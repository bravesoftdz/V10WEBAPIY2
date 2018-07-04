{***********UNITE*************************************************
Auteur  ...... : FV
Créé le ...... : 10/09/2000
Modifié le ... : 15/02/2001
Description .. : Gestion du planning graphique
Mots clefs ... : PLANNING
*****************************************************************}
unit BTPLANNING;

interface
uses
  Windows,
  SysUtils,
  Classes,
  Controls,
  Forms,
  graphics,
  StdCtrls,
  ComCtrls,
  Utob,
{$IFDEF EAGLCLIENT}
  MaineAGL, UtilEagl,                                                      
  eMul,
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fe_main,
  EdtREtat,
  mul,
{$ENDIF}
	AglInit,
  HEnt1,
  HCtrls,
  HMsgBox,
  UIUtil,
  Dialogs,
  Hplanning,
  HPanel,
  Mask,
  ExtCtrls,
  HTB97,
  ParamDat,
  Paramsoc,
  Lookup,
  PlanUtil,
  HeureUtil,
  MsgUtil,
  BTsaisieDate,
  BTPUtil,
  Grids,
  uJoursFeries,
  Utof_VideInside,
  Menus, DBCtrls, HDB, HRichEdt, HRichOLE, HCapCtrl,UDateUtils,
  TntComCtrls, TntStdCtrls, TntGrids, TntExtCtrls, ActnList,UtilsMail;
//
var

  DateAppel     : TDateTime;

  //Permet de déterminer quelles sont les informations affichées dans les onglets
  TypePlanning  : String;
  FamMatGere    : String;
  Fonction      : String;
  ModePlanning  : string;

procedure SaisiePlanning(StArgument: string; mode: integer; action: TActionFiche);
procedure ConsultationPlanning(mode: integer; DateEnCours: TDateTime; Famres, DossierEnCours: string);
procedure OuvrePlanningRessources(Chantier: string; Phase : integer;DateDebut : Tdatetime; Action :TActionFiche);

type

  TTaction = (Ttatransfere,TTaEtirer);

  TMonPlanning = class(TForm)
    MonthCalendar1  : TMonthCalendar;
    PageOnglet      : TPageControl;
    Timer1          : TTimer;
    Dock972         : TDock97;
    Timer2          : TTimer;
    DESCRIPTIF      : THRichEditOLE;
    DesAction       : TMemo;
    GrilleSelection : THGrid;
    MenuCadencement : TPopupMenu;
    //
    CodeDossierRech : TEdit;
    DossierInterdit : TEdit;
    //
    DateEdit        : THCritMaskEdit;
    TheParamPlanning: THCritMaskEdit;
    //
    Toolbar972    : TToolWindow97;
    TTWLegende    : TToolWindow97;
    TTWAppels     : TToolWindow97;
    FenDescriptif : TToolWindow97;
    //
    BPagePrec     : TToolbarButton97;
    BPageSuiv     : TToolbarButton97;
    BCalendrier   : TToolbarButton97;
    BRecharger    : TToolbarButton97;
    BChercher     : TToolbarButton97;
    Bimprimer     : TToolbarButton97;
    BQuitter      : TToolbarButton97;
    ExcelExport   : TToolbarButton97;
    BLegende      : TToolbarButton97;
    BAffAppel     : TToolbarButton97;
    BAffAppel1    : TToolbarButton97;
    BAffAppel2    : TToolbarButton97;
    BValidAff     : TToolbarButton97;
    BAnnulAff     : TToolbarButton97;
    BCREAPPEL     : TToolbarButton97;
    BRefreshAff   : TToolbarButton97;
    BDescriptif   : TToolbarButton97;
    Btaches       : TToolbarButton97;
    Label3        : TLabel;
    Label4        : TLabel;
    Label5        : TLabel;
    Label6        : TLabel;
    Label7        : TLabel;
    Label8        : TLabel;
    Label9        : TLabel;
    Label11       : TLabel;
    CHKTACHES         : TCheckBox;
    //
    HPanel1           : THPanel;
    Panel1            : TPanel;
    BCadencement      : TToolbarButton97;
    BTypePlanning     : TToolbarButton97;
    CBTRAITE          : THCheckbox;
    BtParamListe      : TToolbarButton97;
    BTSupprime        : TToolbarButton97;
    BAffAppel1BvalidImputation: TToolbarButton97;
    BAffAppel1BAnnulImputation: TToolbarButton97;
    EdtDosEnCours     : TEdit;
    NUMPHASE          : TEdit;
    BTMail            : TToolbarButton97;
    BTEditEvt         : TToolbarButton97;
    //
    procedure BAffAppelClick(Sender: TObject);
    procedure BAnnulAffClick(Sender: TObject);
    procedure BCalendrierClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure BCREAPPELClick(Sender: TObject);
    procedure BEnvoiMail(Sender: TObject);
    procedure BImprimeClick(Sender: TObject);
    procedure BLegendeClick(Sender: TObject);
    procedure BPageSuivClick(Sender: TObject);
    procedure BPagePrecClick(Sender: TObject);
    procedure BQuitterClick(Sender: TObject);
    procedure BRechargerClick(Sender: TObject);
    procedure BVALIDAFFClick(Sender: TObject);
    procedure ChangeCadencement(Sender: TObject);
    procedure ExcelexportClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure GestionTimer;
    procedure GrilleSelectionClick(Sender: TObject);
    procedure GrilleSelectionDblClick(Sender: TObject);
    procedure GrilleSelectionKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure PageOngletChange(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure VoirToutBoutons(choix: Boolean);
    procedure BtachesClick(Sender: TObject);
    procedure BTypePlanningClick(Sender: TObject);
    procedure BTAFFMEMOClick(Sender: TObject);
    procedure CBTRAITEClick(Sender: TObject);
    procedure BTSupprimeClick(Sender: TObject);
    procedure BTEditEvtClick(Sender: TObject);

  private
    {Déclarations privées }
    TobModelePlanning : Tob;
    TobOnglet         : Tob;
    TobEtat           : Tob;
    TobItem				  	: Tob;
    TobEvenement		  : Tob;
    TobSuppression	  : Tob;
    TobTaches         : Tob;
    TobAffectChantier : Tob;
    TobJFerie         : Tob;
    TobCalendrier     : Tob;

    Couleur			  : array of TLabel;
    Libelle		  	: array of TLabel;

    OKModif			  : Boolean;
    OkAffect		  : Boolean;
    OkValAction	  : Boolean;
    OKAppelLibre  : Boolean;
    OKVisuAffect  : Boolean;
    OkSpecifCal   : Boolean;
    Ok_TpsReel	  : Boolean;
    Ok_ModifItem  : Boolean;
    TousPossible  : boolean;

    DessineGAppel : AffGrille;

    DatePlanDeb   : TDateTime;
    DatePlanFin   : TDateTime;

    DureeJournee  : Double;
    HeureDeb  	  : Double;
    HeureFin  	  : Double;

    NbOnglet      : Integer;
    NoOnglet      : Integer;

    Nbjour			: String;
    ColNames		: string;
    CodeRes			: string;
    Calendrier  : String;
    Equipe			: string;
    Responsable	: string;
    TypeAction	: string;
    Cadencement	: string;
    CtrlCal     : Boolean;
    Contact     : Integer;

//  ColGAppel 	: String;
//  LibGAppel		: String;
//  TableGapp		: String;
//  TriGapp		 	: String;
//	FormatGapp 	: String;
//  LarGappel	 	: String;

    CreerPar    : String;
    AffaireInit : String;
    CodeTiers   : String;

    OkEnvoiMail : Boolean;

    procedure Affectation(Sender: TObject; Item: Tob; var Cancel: boolean);
    procedure AffichageAffectationRessource;
    procedure AfficheMultiCritere;
    procedure AvertirApplication(Sender: TObject; FromItem: TOB; ToItem: TOB; Actions: THPlanningAction);
    //
    //procedure ChargeAppel;
    function  CalculeDateEvtEtireReduit(Item: TOB; var DateD, DateF: TdateTime; var Delai: integer): boolean;
    procedure ChangePlanning(IndexPage: Integer);
    procedure ChargeEvenementIntervention;
    procedure ChargeInfoAdresse(CodeAffaire: string; var Tel: String);
    procedure ChargeInfoAppel(TobAppel : TOB);
    procedure ChargeInfoContact(Auxiliaire: string);
    procedure ChargeInfotiers(CodeTiers: string; var Auxiliaire, NomTiers: String);
    procedure ChargeInfoTypeAction(Affaire, CodeTiers: string; NoTache: Integer; var TypeAction: string);
    procedure ChargeIntervention;
    procedure ChargeItems;
    function  ChargeLegende: boolean;
    procedure ChargeMenuCadencement;
    procedure ChargementGrilleWithEvtIntervention(TobEvenement : TOB);
    procedure ChargementOngletPlanning(TobOnglet: Tob);
    function  ChargementPlanningStandard(TobModelePlanning: Tob)    : String;
    function  ChargementPlanningSpecifique(TobModelePlanning: Tob)  : String;
    function  ChargeSQLOnglet: String;
    procedure CheckMoveItem (sender : Tobject; Item : TOB; var cancel : boolean);
    procedure ChargementTobappel(StWhere: String);
    function  ConstituStwhereIntervention(Critere: string): string;
    function  Control_BtEtat : Boolean;
    function  ControleAffectation : Boolean;
    function  ControleDansLaGrille : boolean;
    procedure ControleHeureDeb(var HeureDeb: Double);
    procedure ControleHeureFin(var HeureFin: Double);
    function  ControleRessourceArrivee(CodeRes: String; DateDeb, DateFin, HeureDeb, Heurefin: TDateTime): boolean;
    procedure CopyItem(Sender: TObject; Item: TOB; var Cancel: boolean);
    procedure CreateEvenementIntervention;
    procedure CreateEventMateriel;
    procedure CreateEventPlanCharge;
    procedure Creation(Sender: TObject; Item: TOB; var Cancel: boolean);
    procedure CreationActionGRC;
    procedure CreateAppel;
    procedure CreationEvenement(LignePlan : Integer);
    procedure CreationEvtAction(TobLigItem : TOB);
    //
    procedure DefinitZoneDate;
    Function  DeleteEquipePlanCharge(Equipe, Ressource: string; NumEvent: integer) : boolean;
    function  DeleteEvenementMateriel : boolean;
    function  DeleteEvenementPlanCharge : boolean;
    procedure Deplacement(Sender: TObject; Item: TOB; var Cancel: boolean);
    function  DeplacementActionGRC : boolean;
    procedure DoubleClickSpec(ACol, ARow: INTEGER; TypeCellule: TPlanningTypeCellule; T: TOB = nil);
    procedure DrawGrille(NbColFixe: Integer);
    function  DuplicationAction : Boolean;
    //
    procedure EditionIntervention(TobLigItem: TOB);
    procedure EnvoieMail(fSujetMail: String);
    procedure Extension(Sender: TObject; Item: Tob; var Cancel: boolean);
    //
    procedure GererPlanningIntervention;
    procedure GererPlanningMateriel;
    procedure GererPlanningChantier;
    //
    function  InitialisationPlanning: Boolean;
    procedure InitItem(Sender: TObject; var Item: TOB; var Cancel: boolean);
    //
    procedure Link(Sender: TObject; TobSource, TobDestination: TOB; Option: THPlanningOptionLink; var Cancel: Boolean);

    procedure MajActionEquipe;
    procedure MajEvenements(var MajPlanning: THPlanningBTP);
    procedure MajEventBEP;
    procedure MajEventPlanning;
    function  MajEvtAffectation(ModeOption: String): Boolean;
    function  MajEvtParcMateriel(ModeOption: String;Item: TOB): boolean;

    procedure ModifAffectation;
    procedure ModifEventMateriel;
    procedure ModifEventPlanCharge;
    procedure Modification(Sender: TObject; Item: TOB; var Cancel: boolean);
    procedure ModificationAction;
    procedure ModificationEvtAction;
    procedure ModificationsActionGRC;

    //procedure NePasEtirer(Item: TOB);
    //procedure OnPopup(Item: tob; ZCode: integer; var Redraw: boolean);

    procedure P1BeforeChange(const Item: TOB; const LaRessource: string; const LaDateDeDebut, LaDateDeFin: TDateTime; var Cancel: Boolean);

    //procedure ReafficherItem;
    procedure ReChargeActionGRC(Auxiliaire: string; NumAct: integer); overload;
    procedure ReChargeActionGRC(Item : TOB); overload;
    procedure RechargePlanning(TMO : Tob; IndexPage : Integer);
    function  RechEquipe(Equipe : String) : Boolean;
    function  RechEquipeAction: Boolean;
    procedure Reduction(Sender: TObject; Item: Tob;  var Cancel: boolean);
    function  ReductionEtirementActionGRC(TobLigItem: TOB): boolean;

    procedure SetInfosAffectation(TobEvt: tob);
    procedure SupprimeAffectation;
    procedure Suppression(Sender: TObject; Item: TOB; var Cancel: boolean);
    procedure SuppressionEvenementIntervention;
    procedure SuppressionActionGRC;
    function  SuppressionEquipe(Prefixe, Equipe: String) : Boolean;
    procedure SuppressionEvtInterv(Numevent: String);
    procedure SuppressionEvtIntervEquipe(NumEvent: String);
    procedure SuppressionItemPlanning(CritereLecture: String);

    procedure TraiteLigneSelectionne(TobEventMat : TOB);
    procedure TraitementAffectationAppel;
    procedure traitementEventParcMateriel(TOBL: TOB);
    procedure TraitementParcMateriel;
    procedure TraitementPlancharge;
    procedure TraitSuppItem(CodeRes, NumAppel, NumEvenement: String);

    procedure ValidationAffectation;
    procedure VisualisationAffectation;

    procedure CreationItems(LigItem : TOB; TypeCreation: string);
    procedure ModificationItem(LigItem: Tob);
    procedure SuppressionItem(LigItem: Tob);
    procedure ChargementTobItem(TobRes : Tob);
    procedure ModificationEvenementIntervention;
    procedure SuppressionEvenementAppel;
    procedure ChargementDateItem(TobRes, Item: Tob);
    procedure ChargeTobDestinationWithTobSource(TobDest, TobSource: Tob);

  public
    {Déclarations publiques }
    tempscalcule: TTime;
    feuille   : array[0..99] of TTabSheet;
    Planning  : array[0..99] of THPlanningBTP;
{$IFDEF TESTXP}
    PlanningType: array[0..99] of THPlanningBTP;
{$ENDIF}
    TobPlannings: array[0..99] of RecordPlanning;
    PlanningCurrent : THplanningBTP;
    DateSelectionee : TDateTime;
    FindPRien: THpanel;

    procedure GestionSynchroDate(Item: Tob);
    procedure GestionDateLigne(Item: Tob);

  protected // XP 07.06.2006
    CancelAction: boolean ;
  end;

var
  PremierPassage: Boolean = True;
  FicAppel: Boolean;

implementation

{$R *.DFM}

uses uBtpEtatPlanning,
     UTilFonctionCalcul,
     DateUtils,
     UtilsParc,
     AGLInitGC,
     ParamDBG,
     DialogEx,
     UtilsParcPlanning,
     UtilActionPlanning
     ;

function Recupdefaut: string;
var QModelePlanning: TQuery;
    StReq : String;
begin

  result := '';

  stReq := 'Select HPP_PARAMPLANNING from HRPARAMPLANNING where (HPP_MODEPLANNING="' + TypePlanning + '") AND HPP_PLANNINGDEFAUT="X"';

  QModelePlanning := OpenSQL(streq, True,-1,'',true);

  if not QModelePlanning.eof then
    Result := QModelePlanning.FindField('HPP_PARAMPLANNING').AsString;

  Ferme(QModelePlanning);

end;


procedure OuvrePlanningRessources(Chantier: string; Phase : integer;DateDebut : Tdatetime; Action :TActionFiche);
var
  Inside        : THPanel;
  X             : TMonPlanning;
  ParamPlanning : string;
  TousPossible  : Boolean;
begin

  TousPossible := False;
  ModePlanning := 'PCA';
  (*

  //Récupération du modéle de planning (Intervention, Parc/Matériel, Chantier)
  ModePlanning := Rechdom('BTMODEPLANNING', TypePlanning, True);

  //récupération du planning par défaut
  ParamPlanning := Recupdefaut;
  *)
  TypePlanning  := GetParamSocSecur('SO_BTTYPELAN','');
  ParamPlanning := GetParamSocSecur('SO_BTPLANNINGDEF','');
  
  if paramplanning = '' then
  begin
    PGIInfo ('merci de paramétrer vos plannings');
    exit;
  end;

  {le planning détaillé est deja chargé}
  if not PremierPassage then exit;

  NumLigneTO := -1;

  if (Action = Tamodif) then
    ModeSaisie := True
  else if (Action = TaConsult) then
    ModeSaisie := False;

  X := TMonPlanning.Create(Application);

  if (ParamPlanning <> '') then X.TheParamPlanning.text := ParamPlanning;

  X.TousPossible         := false;
  X.DateEdit.Text        := DateTimeToStr(DateDebut);
  X.CodeDossierRech.text := Chantier;
  X.EdtDosEnCours.text   := X.CodeDossierRech.text;
  X.NUMPHASE.Text      := IntToStr(Phase);

{$IFDEF EAGLCLIENT}
  Inside := FindInsidePanel;
{$ELSE}
  Inside := nil;
{$ENDIF}

  if Inside = nil then
  begin
    X.FindPrien := nil;
    try
      X.ShowModal;
    finally
      X.Free;
    end;
  end
  else
  begin
    X.findprien := Inside;
    InitInside(X, Inside);
    X.Show;
  end;

end;

procedure SaisiePlanning(StArgument: string; mode: integer; Action: TActionFiche);
var
  Inside        : THPanel;
  X             : TMonPlanning;
  StDate        : string;
  StContrat     : string;
  ParamPlanning : string;
  CodeAppel     : string;
  CodeChantier  : string;
  TousPossible  : Boolean;
begin

  StContrat := '-';
  TousPossible := False;
  //
  TypePlanning := '';
  ModePlanning := '';
  FamMatGere   := '';
  Fonction     := '';
  //

  if TypePlanning = '' then
  begin
    if (Trim(Copy(StArgument,0,5))) = 'APPEL' then
      TypePlanning := 'PLA'
    else
      TypePlanning := StArgument;
  end
  else
     TypePlanning := StArgument;

  //Récupération du modéle de planning (Intervention, Parc/Matériel, Chantier)
  ModePlanning := Rechdom('BTMODEPLANNING', TypePlanning, True);

  //récupération du planning par défaut
  ParamPlanning := Recupdefaut;

  {le planning détaillé est deja chargé}
  if not PremierPassage then exit;

  NumLigneTO := -1;

  if (Action = Tamodif) then
    ModeSaisie := True
  else if (Action = TaConsult) then
    ModeSaisie := False;

  FicAppel     := False;

  if (Trim(ReadTokenSt(StArgument))) = 'APPEL' then
  Begin
     CodeAppel    := (Trim(ReadTokenSt(StArgument)));
     DateAppel    := StrToDate(Trim(ReadTokenSt(StArgument)));
     StDate       := DateToStr(DateAppel);
     TypePlanning := 'PLA';
     ModePlanning := Rechdom('BTMODEPLANNING', TypePlanning, True);
     FicAppel     := True;
     if ParamPlanning = '' then
       StDate := SaisieDate(CodeAppel, ParamPlanning, TypePlanning,TousPossible, StDate, true, true);
  end
  else if (Trim(ReadTokenSt(StArgument))) = 'CHANTIER' then
  Begin
     CodeChantier := (Trim(ReadTokenSt(StArgument)));
     DateAppel := StrToDate(Trim(ReadTokenSt(StArgument)));
     StDate       := DateToStr(DateAppel);
     TypePlanning := 'PAR';
     ModePlanning := Rechdom('BTMODEPLANNING', TypePlanning, True);
     FicAppel     := True;
     if ParamPlanning = '' then
       StDate := SaisieDate(CodeChantier, ParamPlanning, TypePlanning,TousPossible, StDate, true, true);
  end
  else
     StDate := SaisieDate(StContrat, ParamPlanning, TypePlanning,TousPossible, datetimetostr(v_pgi.dateentree), true, true);

  if (StrToDate(StDate) = IDate1900) then Exit;

  if ParamPlanning = '' then
  begin
    PgiInfo('Aucun modèle de défini', 'Saisie Planning');
    SaisiePlanning(TypePlanning, mode, Action);
  end;

  X := TMonPlanning.Create(Application);

  if (ParamPlanning <> '') then X.TheParamPlanning.text := ParamPlanning;

  X.TousPossible         := TousPossible;
  X.DateEdit.Text        := StDate;

  if CodeAppel <> '' then
    X.CodeDossierRech.text := CodeAppel
  else if CodeChantier <> '' then
    X.CodeDossierRech.text := CodeChantier
  else
    X.CodeDossierRech.text := '';

  X.EdtDosEnCours.text  := X.CodeDossierRech.text;
  X.NUMPHASE.Text       := '0';

{$IFDEF EAGLCLIENT}
  Inside := FindInsidePanel;
{$ELSE}
  Inside := nil;
{$ENDIF}

  if Inside = nil then
  begin
    X.FindPrien := nil;
    TRY
      try
        X.ShowModal;
      EXCEPT
        on E : Exception do PGIError(E.Message);
      END;
    finally
      X.Free;
    end;
  end
  else
  begin
    X.findprien := Inside;
    InitInside(X, Inside);
    X.Show;
  end;

end;

procedure ConsultationPlanning(mode: integer; DateEnCours: TDateTime; Famres, DossierEnCours: string);
var
  Inside: THPanel;
  X: TMonPlanning;
begin

  //le planning détaillé est deja chargé
  if not PremierPassage then exit;

  NumLigneTO := -1;
  //ModeGestion := Mode;
  ModeSaisie := False;
  X := TMonPlanning.Create(Application);
  X.DateEdit.Text := DateTimeToStr(DateEnCours);
  X.DossierInterdit.Text := DossierEnCours;
{$IFDEF EAGLCLIENT}
  Inside := FindInsidePanel;
{$ELSE}
  Inside := nil;
{$ENDIF}
  if Inside = nil then
  begin
    X.FindPrien := nil;
    try
      X.ShowModal;
    finally
      X.Free;
    end;
  end
  else
  begin
    InitInside(X, Inside);
    X.Show;
  end;
end;

procedure TMonPlanning.FormCreate(Sender: TObject);
var
  i: integer;
begin

  PremierPassage := False;

  MenuCadencement.AutoPopup := False;

  Self.WindowState := wsMaximized;

  for i := 0 to 99 do
  begin
    feuille[i] := nil;
    Planning[i] := nil;
    with TobPlannings[I] do
    begin
      TobItems  := nil;
      TobEtat   := nil;
      TobRes    := nil;
      TobCols   := nil;
      TobRows   := nil;
      TobEvents := nil;
      TobPeriodeDivers := nil;
    end;
  end;

  TobItem         := nil;
  TobEtat         := nil;
  TobEvenement    := Nil;
  //
  TobSuppression  := nil;
  TobTaches       := nil;

  Application.HintHidePause := 20000;

  //Creation des tob génériques
  TobItem := Tob.create('BTEVENEMENTPLA', nil, -1);

  //Creation de la tob évènements liées à la gestion des affecations appels
  TobEvenement := Tob.create('Les Evenements', nil, -1);
  TobEvenement.AddchampSupValeur('AFFAIRE', '');
  //
  TobTaches := Tob.create('TACHE', nil, -1);
  TobEtat   := Tob.create('BTETAT', nil, -1);
  //
  TobJFerie := Tob.create('JOURFERIE', nil, -1);
end;

procedure TMonPlanning.FormShow(Sender: TObject);
begin

  //Controle du type d'action
  if ModePlanning = 'PMA'  then
    GererPlanningMateriel
  else if ModePlanning = 'PCA'  then
    GererPlanningChantier
  else
    GererPlanningIntervention;

  // XP 07.06.2006 Par défaut, l'action est toujours réfusée donc il faut mettre CancelAction à True
  CancelAction := True;

  if TypePlanning = '' then
  Begin
    HShowMessage('', TraduireMemoire('Attention !'), TraduireMemoire('Aucun type de planning n''a été défini'));
    BPageSuiv.Enabled := False;
    BPagePRec.Enabled := False;
    BCalendrier.Enabled := False;
    Close;
    exit;
  End;

  {$IFDEF EAGLCLIENT}
  BQUITTER.Enabled := False;
  BCREAPPEL.Visible := False;
  {$ENDIF}

  OkEnvoiMail := GetParamSocSecur ('SO_BTAVERTIRENMODIF',true);

  // Initialisation de la date en cours
  if (DateEdit.Text = DateTimeToStr(iDate1900)) then
  begin
    DateSelectionee := V_PGI.DateEntree;
    DateEdit.Text := DateTimeToStr(V_PGI.DateEntree);
  end
  else
    DateSelectionee := StrToDate(DateEdit.Text);

  If not InitialisationPlanning then
  begin
    Close;
    exit;
  end;

  CHKTACHES.Visible := True;

  BAffAppel1.Down   := False;
  BAffAppel2.Down   := False;
  BAffAppel.Down    := False;

  if (ModeSaisie) And (not FicAppel) then BCREAPPEL.Visible := not FicAppel;

  //Chargement du menu Contextuel pour gestion des cadencements en dynamique
  ChargeMenuCadencement;

  //chargement de la legende
  if not ChargeLegende then Close;

  ChargeJoursFeries(TobJferie);

  GestionTimer;

  if (FicAppel) And (CodeDossierRech.text<>'') then
  begin
    BAffAppel.down    := False;
    BAffAppel.Visible := False;
    BCREAPPEL.Visible := False;
    ChargementTobappel(ConstituStWhereIntervention(CodeDossierRech.text));
  end;

end;

Function TMonPlanning.InitialisationPlanning : Boolean;
Var QModelePlanning : TQuery;
		Sql             :string;
    Ok_Onglet       : string;
Begin

  Result := True;

  // Chargement des modèles de planning dans la TOB les modèles ET LES ONGLETS
  TobModelePlanning := TOB.Create('les modeles', nil, -1);
  TobOnglet         := TOB.Create('les Onglets', nil, -1);

  //on charge TOUS les champs
  //MajParamplanning;

  try
    if TousPossible then
      Sql := 'Select * from HRPARAMPLANNING where HPP_MODEPLANNING="' + TypePlanning + '" ORDER BY HPP_PLANNINGDEFAUT DESC,HPP_PARAMPLANNING'
    else
      Sql := 'Select * from HRPARAMPLANNING where HPP_MODEPLANNING="' + TypePlanning + '" AND HPP_PARAMPLANNING="' + TheParamPlanning.Text + '" ORDER BY HPP_PLANNINGDEFAUT DESC,HPP_PARAMPLANNING';
    //
    QModelePlanning := OpenSQL(Sql , True,-1,'',true);
    //
    if QModelePlanning.eof then
    begin
      Result := False;
      // Test si pas de modèles de planning
      HShowMessage('', TraduireMemoire('Attention !'), TraduireMemoire('Aucun modèle de planning n''a été défini'));
      BPageSuiv.Enabled   := Result;
      BPagePRec.Enabled   := Result;
      BCalendrier.Enabled := Result;
    end
    else
    begin
      TobModelePlanning.LoadDetailDB('HRPARAMPLANNING', '', '', QModelePlanning, True);
    end;
  finally
    Ferme(QModelePlanning);
  end;

  if not result then Exit;

  // Initialisation de la date en cours
  if (DateEdit.Text = DateTimeToStr(iDate1900)) then
  begin
    DateSelectionee := V_PGI.DateEntree;
    DateEdit.Text := DateTimeToStr(V_PGI.DateEntree);
  end
  else
    DateSelectionee := StrToDate(DateEdit.Text);
  //
  //On charge ici les famille matériel que l'on veut voir dans le planning en onglet
  //si le type de planning est PSF (Planning sous-Famille) uniquement
  if TypePlanning = 'PFO' then Fonction := TobModelePlanning.Detail[0].GetValue('HPP_FAMRES');
  //
  FamMatGere := TobModelePlanning.Detail[0].GetValue('HPP_FAMMATGERE');
  FamMatGere := FormatMultiValComboforSQL(FamMatGere);
  //
  if Pos(TypePlanning,'PLA;PMA;PRA;PAR') > 0 then
    Ok_Onglet := ChargementPlanningStandard(TobModelePlanning)
  Else
    Ok_Onglet := ChargementPlanningSpecifique(TobModelePlanning);

  //Si on a pas pu charger un seul élément du planning ==> Erreur

  if Ok_Onglet <> '' then
  begin
    Result := false;
    if      Ok_Onglet = 'PFM' then HShowMessage('', TraduireMemoire('Attention !'), TraduireMemoire('Aucune Famille Matériel à charger'))
    else if Ok_Onglet = 'PTA' then HShowMessage('', TraduireMemoire('Attention !'), TraduireMemoire('Aucun Type d''action intervention à charger'))
    Else if Ok_Onglet = 'PPA' then HShowMessage('', TraduireMemoire('Attention !'), TraduireMemoire('Aucun Type d''action Parc/Matériel à charger'))
    Else if OK_Onglet = 'PAR' then HShowMessage('', TraduireMemoire('Attention !'), TraduireMemoire('Aucun Type d''action Affaire à charger'))
    Else if OK_Onglet = 'PFO' then HShowMessage('', TraduireMemoire('Attention !'), TraduireMemoire('Aucune Fonction de ressource à charger'));
  end;

  BPageSuiv.Enabled   := Result;
  BPagePRec.Enabled   := Result;
  BCalendrier.Enabled := Result;

end;

//
//Procedure de chargement de la requête pour les onglets
//
Function TMonPlanning.ChargeSQLOnglet : String;
begin

  If TypePlanning = 'PSF' then //Planning Sous-Type Ressource
  Begin
     //On ne prends que les type de ressource associés à des ressources qui ne sont pas fermées
     Result := 'SELECT DISTINCT BTR_LIBELLE AS LIBELLE, BTR_TYPRES AS CODEONGLET FROM BTTYPERES ';
     Result := Result + 'LEFT JOIN RESSOURCE ON BTR_TYPRES=ARS_CHAINEORDO ';
     Result := Result + 'WHERE BTR_GEREPLANNING="X" AND ARS_CHAINEORDO <> "" AND ARS_FERME = "-"';
  end
  else if TypePlanning = 'PTA' then //Planning par type d'Action
  Begin
     Result := 'Select BTA_LIBELLE AS LIBELLE, BTA_BTETAT AS CODEONGLET FROM BTETAT ';
     Result := Result + 'WHERE BTA_TYPEACTION="INT" AND BTA_ASSOSRES = "X" ORDER BY BTA_BTETAT';
  end
  else if TypePlanning = 'PTR' then   //Planning par Type Ressource
  begin
     Result := 'SELECT DISTINCT CC_LIBELLE AS LIBELLE, ARS_TYPERESSOURCE AS CODEONGLET FROM CHOIXCOD ';
     Result := Result + 'LEFT JOIN RESSOURCE ON ARS_TYPERESSOURCE=CC_CODE ';
     Result := Result + 'WHERE ARS_TYPERESSOURCE <> "" AND ARS_FERME = "-"';
  end
  else if TypePlanning = 'PFO' then   //Planning par Fonction Ressource
  begin
    Result := 'SELECT DISTINCT AFO_LIBELLE AS LIBELLE, BFO_FONCTION AS CODEONGLET FROM FONCTION ';
    Result := Result + 'LEFT JOIN BFONCTION ON AFO_FONCTION=BFO_FONCTION  ';
    Result := Result + 'LEFT JOIN RESSOURCE ON AFO_FONCTION=ARS_FONCTION1 ';
    Result := Result + 'LEFT JOIN BRESSOURCE ON BFO_FONCTION=ARS_FONCTION1 ';
    Result := Result + 'WHERE BFO_GEREPLANNING = "X"';
    if Fonction <> '' then
      Result := Result + ' AND BFO_FONCTION="' + Fonction + '"';
  end
  else if TypePlanning = 'PPA' then //Planning Parc/Matériel par type d'Action
  Begin
     Result := 'Select BTA_LIBELLE AS LIBELLE, BTA_BTETAT AS CODEONGLET FROM BTETAT ';
     Result := Result + 'WHERE BTA_TYPEACTION="PMA" AND BTA_ASSOSRES = "X" ORDER BY BTA_BTETAT';
  end
  else if TypePlanning = 'PFM' then   //Planning par Famille Matériel
  begin
     Result := 'SELECT DISTINCT BFM_LIBELLE AS LIBELLE, BFM_CODEFAMILLE AS CODEONGLET ';
     Result := Result + '     FROM BTMATERIEL ';
     Result := Result + 'LEFT JOIN BTFAMILLEMAT ON BFM_CODEFAMILLE=BMA_CODEFAMILLE ';
     Result := Result + '    WHERE BFM_NONGEREPLANNING="X" ';
     if FamMatGere <> '' then
        Result := Result + '   AND BFM_CODEFAMILLE IN (' + FamMatGere + ')';
  end;

end;

//
function TMonPlanning.ChargementPlanningSpecifique(TobModelePlanning : Tob) : String;
Var QTableOnglet    : TQuery;
begin

  Result := '';

  try
    QTableOnglet := OpenSQL(ChargeSQLOnglet, True,-1,'',true);
    if Not QTableOnglet.EOF then
    begin
      TobOnglet.LoadDetailDB('ONGLET', '', '', QTableOnglet, True);
      ChargementOngletPlanning(TobOnglet);
    end
    else
      Result := TypePlanning;
  finally
    Ferme(QTableOnglet);
  end;

end;

function TMonPlanning.ChargementPlanningStandard(TobModelePlanning : Tob) : String;
Var NModele     : integer;
    i           : Integer;
    TobLOnglet  : TOB;
begin

  Result  := '';

  NModele := 0;

  for i := 0 to TobModelePlanning.detail.count -1 do
  begin
    //Création de la table Onglet
    TobLOnglet := TOB.Create('ONGLET',TobOnglet,-1);
    TobLOnglet.AddChampSupValeur('LIBELLE',    TobModelePlanning.detail[i].GetValue('HPP_ABREGE'));
    TobLOnglet.AddChampSupValeur('CODEONGLET', TobModelePlanning.detail[i].GetValue('HPP_PARAMPLANNING'));
  end;

  ChargementOngletPlanning(TobOnglet);

  // Charge les paramètres du premier planning à afficher
  if (TheParamPlanning.Text <> '') then
  begin
    for i := 0 to TobModelePlanning.detail.count - 1 do
    begin
      if (TobModelePlanning.detail[i].getvalue('HPP_PARAMPLANNING') = TheParamPlanning.Text) then
      begin
        NModele := i;
        break;
      end;
    end;
  end;

  PlanningCurrent := Planning[NModele];

end;

// Initilisation des items
procedure TMonPlanning.InitItem(Sender: TObject; var Item: TOB; var Cancel: boolean);
begin

  if Item = nil then
  begin
     Item := TOB.Create('BTEVENEMENTPLA', nil, -1);
     AddChampsSupItems (Item);
  end;

  Cancel := False;

end;

procedure TMonPlanning.Affectation(Sender: TObject; Item: Tob; var Cancel: boolean);
Begin

	Cancel := False ;

  if Cancel then exit;

end;


procedure TMonPlanning.FormClose(Sender: TObject; var Action: TCloseAction);
var
  MyMenu: TMenuItem;
begin

  // Free tob des modeles de planning et tob des ressources
  if TobModelePlanning <> nil then TobModelePlanning.Free;
  if TobOnglet         <> nil then TobOnglet.free;
  if TobEtat           <> nil then TobEtat.Free;

  FreeAndNil(Tobtaches);

  FreeAndNil(TobEvenement);
  //
  if TobSuppression <> nil then TobSuppression.free;

  //Lbération du Menu PopUp
  if MenuCadencement <> nil then
  begin
    if MenuCadencement.items.count = 0 then exit;
    repeat
      MyMenu := MenuCadencement.items[0];
      if Assigned(MyMenu) then MyMenu.Free;
    until MenuCadencement.items.Count <= 0;
  end;

end;

procedure TMonPlanning.PageOngletChange(Sender: TObject);
begin

  ChangePlanning(PageOnglet.ActivePageIndex);

end;

procedure TMonPlanning.BCalendrierClick(Sender: TObject);
var DateAvant: TDateTime;
    DateEnCours: TDateTime;
    ParamPlanning: string;
    StContrat: string;
    TousPossible : boolean;
    IndexPage   : Integer;
begin

  IndexPage := PageOnglet.ActivePageIndex;

  StContrat := '-';
  ParamPlanning := '-';
  TousPossible := false;

  DateAvant     := StrToDateTime(DateEdit.text);
  DateEdit.text := SaisieDate(StContrat, ParamPlanning, TypePlanning, TousPossible,DateEdit.text);
  DateEnCours   := StrToDateTime(DateEdit.text);

  if DateEnCours = Idate1900  then  DateEnCours := DateAvant;

  if DateEnCours <> DateAvant then ChangePlanning(IndexPage);

  ChargeItems;

end;


procedure TMonPlanning.BPageSuivClick(Sender: TObject);
var DateEnCours : TDateTime;
    Intervalle  : Integer;
begin

  //Récupération du planning en cours
  Intervalle      := Round((TobModelePlanning.Detail[0].GetValue('HPP_INTERVALLEDEB') +
                            TobModelePlanning.Detail[0].GetValue('HPP_INTERVALLEFIN')) / 2);

  //Mise à jour de l'intervallle de temps du planning (décalage d'un planning vers la droite)
  DateEnCours     := PlanningCurrent.DateOfStart + Intervalle;
  DateEdit.Text   := DateTimeToStr(DateEnCours);

  ChargeItems;

end;


procedure TMonPlanning.BPagePrecClick(Sender: TObject);
var DateEnCours : TDateTime;
    Intervalle  : Integer;
begin

  //Mise à jour de l'intervallle de temps du planning (décalage d'un planning vers la droite)
  Intervalle    := TobModelePlanning.Detail[0].GetValue('HPP_INTERVALLEDEB') +
                   TobModelePlanning.Detail[0].GetValue('HPP_INTERVALLEFIN');
  Intervalle    := Round(intervalle / 2);

  DateEnCours   := PlanningCurrent.DateOfStart - intervalle;
  DateEdit.Text := DateTimeToStr(DateEnCours);

  ChargeItems;

end;

procedure TMonPlanning.BRechargerClick(Sender: TObject);
var  DateEnCours: TDateTime;
begin

  // Rechargement des plannings
  uBtpEtatPlanning.BTP_LesEtats.ReLoad();

  //Mise à jour de l'intervallle de temps du planning (décalage d'un planning vers la droite)
  DateEnCours := PlanningCurrent.DateOfStart;

  ChangePlanning(PageOnglet.ActivePageIndex);

end;

procedure TMonPlanning.BEnvoiMail(Sender: TObject);
Var TobLigItem : Tob;
begin

  Try
    TobLigItem := PlanningCurrent.GetCurItem;

    if (TobLigItem = nil) then Exit;

    if TobLigItem.GetValue('BPL_ORIGINEITEM') = 'INTERV' then EnvoieEmailFromIntervention(Self, TobLigItem);

  except
    ShowMessage('Aucun événement n''est sélectionné');
  end;

end;

procedure TMonPlanning.CBTRAITEClick(Sender: TObject);
begin

  ChangePlanning(PageOnglet.ActivePageIndex);

end;


Procedure TMOnPlanning.ChangePlanning(IndexPage : Integer);
Var NModele : Integer;
begin

  VoirToutBoutons(false);

  //Ici c'est de la grosse merdasse !!!!
  if PlanningCurrent <> nil then
  begin
    PlanningCurrent.Enabled := False;

    if NumLigneTO >= 0 then
    begin
      NumLigneTO := -1;
      PlanningCurrent.SupprimeLigneFixe(0);
    end;
    //
    PlanningCurrent                 := Planning[IndexPage];
    PlanningCurrent.Activate        := False;
    PlanningCurrent.Name            := Format('PLANNING%2.2d', [indexPage]);
    PlanningCurrent.Align           := alClient;
    PlanningCurrent.AppelsTraites   := CBTRAITE.Checked;
    PlanningCurrent.TypePlanning    := Typeplanning;
    PlanningCurrent.CodeOnglet      := TobOnglet.Detail[IndexPage].GetValue('CODEONGLET');
    PlanningCurrent.tag             := indexPage;
    //
    if TousPossible then
      nModele                       := indexPage
    else
      nModele                       := 0;
    //
    Cadencement                     := TobModelePlanning.Detail[NModele].Getvalue('HPP_CADENCEMENT');
    CtrlCal                         := TobModelePlanning.Detail[NModele].GetBoolean('HPP_CONTROLCAL');

    PlanningCurrent.FamRes          := TobModelePlanning.Detail[NModele].GetValue('HPP_FAMRES');

    if ModePlanning = 'PMA' Then //Pos(TypePlanning,'PMA;PFM;PPA;PTA')>0 then     //Planning de type PARC
    Begin
      ChargeCadencement(TobPlannings[IndexPage], PlanningCurrent, TobModelePlanning.Detail[0]);
    end
    else if ModePlanning = 'PCA' Then //pos(TypePlanning, 'PAR;PRA') > 0 then    //Planning de type Affaire/Chantier
    Begin
      ChargeCadencement(TobPlannings[IndexPage], PlanningCurrent, TobModelePlanning.Detail[0]);
    end
    else if ModePlanning = 'PLA' Then //pos(TypePlanning, 'PTA;PSF;PTR;PFO') > 0 then //Planning de Type Intervention/Ressource
      ChargeCadencement(TobPlannings[IndexPage], PlanningCurrent, TobModelePlanning.Detail[0])
    else
      ChargeCadencement(TobPlannings[IndexPage], PlanningCurrent, TobModelePlanning.Detail[IndexPage]);
    //
    PlanningCurrent.Enabled   := True;
    PlanningCurrent.Activate  := true;
    PlanningCurrent.Invalidate;
    //
    ChargeItems;
  end;

  VoirToutBoutons(True);

end;

Procedure TMonPlanning.RechargePlanning(TMO : Tob; IndexPage : Integer);
var CodeFamres : String;
Begin

  // Recharge les Onglets du planning
  if (Typeplanning = 'PLA') then
  Begin
    ChargeParamPlanning(TobPlannings[IndexPage], Planning[IndexPage], TobModelePlanning.Detail[IndexPage], StrToDateTime(DateEdit.Text), True);     // MAJ du titre de la fenetre
    //
    CodeFamres  := TobModelePlanning.Detail[IndexPage].getvalue('HPP_FAMRES');
  end
  Else if Typeplanning = 'PMA' then
  begin
    ChargeParamPlanning(TobPlannings[IndexPage], Planning[IndexPage], TobModelePlanning.Detail[IndexPage], StrToDateTime(DateEdit.Text), True);     // MAJ du titre de la fenetre
    CodeFamres                    := Planning[IndexPage].FamRes;
  end
  Else if Typeplanning = 'PRA' then
  begin
    ChargeParamPlanning(TobPlannings[IndexPage], Planning[IndexPage], TobModelePlanning.Detail[IndexPage], StrToDateTime(DateEdit.Text), True);     // MAJ du titre de la fenetre
    CodeFamres                    := Planning[IndexPage].FamRes;
  end
  else
  Begin
    // MAJ du titre de la fenetre
    ChargeParamPlanning(TobPlannings[IndexPage], Planning[IndexPage], TobModelePlanning.Detail[0], StrToDateTime(DateEdit.Text), True);
  end;

  if CodeFamres = 'GEN' then
    Caption     := TraduireMemoire(TobModelePlanning.Detail[0].GetValue('HPP_LIBELLE'))
  else
    Caption   := TraduireMemoire('Planning Ressources le ' + DateTimeToStr(DateSelectionee)); //' + RechDom('HRFAMRES', CodeFamres, false) + '

end;


// Création d'un item ==> Planning Intervention
procedure TMonPlanning.Creation(Sender: TObject; Item: TOB; var Cancel: boolean);
Var LignePlan   : Integer;
    TobRes      : Tob;
begin

  CreationItems(Item, 'X');
  Cancel := true;

end;

// Modification d'un item de type Evenement
procedure TMonPlanning.Modification(Sender: TObject; Item: TOB; var Cancel: boolean);
begin

  ModificationItem(Item);
  Cancel := true;

end;

// Supression d'un item
procedure TMonPlanning.Suppression(Sender: TObject; Item: TOB; var Cancel: boolean);
Begin

  Cancel := true;

  SuppressionItem(Item);

end;

// Si touche F5 -> calendrier
procedure TMonPlanning.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var cancel : boolean;
begin

  if key = vk_F5 then BCalendrierClick(Self);

{$IFNDEF EAGLCLIENT}
  if Key = VK_ESCAPE then BQuitterClick(Self);
{$ENDIF}
  if Key = VK_DELETE then
  begin
    cancel := false;
    if PlanningCurrent.GetCurItem <> nil then  Suppression(self, PlanningCurrent.GetCurItem, Cancel);
  end;
end;

procedure TMonPlanning.BLegendeClick(Sender: TObject);
begin
  TTWLegende.Visible := Blegende.down;
end;

procedure TMonPlanning.BImprimeClick(Sender: TObject);
var
  Codefamres: string;
  Criteres: string;
begin

  if ModePlanning = 'PMA' then
  begin
    CodeFamRes := TobModelePlanning.Detail[PlanningCurrent.Tag].getvalue('HPP_FAMRES');
    Criteres := 'HCA_CALENDRIER=' + datetimetostr(PlanningCurrent.IntervalDebut) + '`HCA_CALENDRIER_=' + datetimetostr(PlanningCurrent.IntervalFin);
  end
  else
  begin
    CodeFamres := TobModelePlanning.Detail[PlanningCurrent.Tag].getvalue('HPP_FAMRES');
    Criteres := 'HCA_CALENDRIER=' + datetimetostr(PlanningCurrent.IntervalDebut) + '`HCA_CALENDRIER_=' + datetimetostr(PlanningCurrent.IntervalFin) +
    '`FAMRES=' + RechDom('HRFAMRES', CodeFamres, false);
  end;

{$IFNDEF EAGLCLIENT}
  LanceEtat('E', 'HOT', 'PRE', True, False, False, nil, '', '', False, 0, Criteres, PlanningCurrent.DoPrint);
{$ELSE}
  planningCurrent.TypeEtat := 'E';
  planningCurrent.NatureEtat := 'HOT';
  planningCurrent.CodeEtat := 'PRE';
  planningCurrent.Print(Criteres);
{$ENDIF}

end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 07/06/2006
Modifié le ... :   /  /
Description .. : procédure appelée par le planning sur les évènements
Suite ........ : suivants :
Suite ........ : - copie, déplacement, suppression, etirement, réduction...
Mots clefs ... :
*****************************************************************}
procedure TMonPlanning.AvertirApplication(Sender: TObject; FromItem, ToItem: TOB; Actions: THPlanningAction);
begin

  //FromItem 	: TobItem de l'endroit d'où on part
  //ToItem		: TobItem de l'endroit où on arrive
  //Action		: Action effectuée

  (***
  THPlanningAction = (
    paCreate,
    paModif,
    paConsult,
    paDelete,
    paLink,
    paDelLink,
    paMove,
    paCopy,
    paAll,
    paNone,
    paExtend,
    paReduce,
    paClickLeft,
    paClickRight,
    paMouseMove,
    paCellEnter,
    paResizeColumns);
***)

	CancelAction := False;

  case Actions of
    paClickLeft  :
      begin
        BTMail.Visible    := False;
        BTEditEvt.Visible := False;
        //Si nous sommes sur un évènement Intervention affichage du bouton mail et Impression bon d'intevention
        if FromItem <> nil then
        begin
          if FromItem.GetValue('BPL_ORIGINEITEM') = 'INTERV' then
          begin
            if FromItem.getValue('BPL_AFFAIRE') <> '' then
            begin
              BTMail.Visible    := True;
              BTEditEvt.Visible := True;
            end;
          end;
        end;
      end;
    paClickRight :
       Begin
         //THPlanningBTP(Sender).DisplayOptionLiaison := False;
         if FromItem <> nil then
         Begin
            if FromItem.GetValue('BPL_ORIGINEITEM') = 'ACT-GRC' then
            Begin
              ThPlanningBTP(Sender).EnableOptionPopup(1, False);
              ThPlanningBTP(Sender).EnableOptionPopup(2, False);
              ThPlanningBTP(Sender).EnableOptionPopup(3, False);
              ThPlanningBTP(Sender).EnableOptionPopup(4, False);
              ThPlanningBTP(Sender).EnableOptionPopup(5, False);
              ThPlanningBTP(Sender).EnableOptionPopup(6, False);
            end
            Else //if FromItem.GetValue('BPL_ORIGINEITEM') = 'INTERV' then
            begin
              ThPlanningBTP(Sender).EnableOptionPopup(1, True);
              ThPlanningBTP(Sender).EnableOptionPopup(2, True);
              ThPlanningBTP(Sender).EnableOptionPopup(3, True);
              ThPlanningBTP(Sender).EnableOptionPopup(4, True);
              ThPlanningBTP(Sender).EnableOptionPopup(5, True);
              ThPlanningBTP(Sender).EnableOptionPopup(6, True);
              ThPlanningBTP(Sender).EnableOptionPopup(7, True);
              ThPlanningBTP(Sender).EnableOptionPopup(8, True);
            end;
         end;
       end;
    paExtend :
    	begin
      	IF (ToItem <> nil) or Assigned(ToItem) then Extension(Sender, ToItem, CancelAction);
      end;
    paReduce :
    	begin
      	If (ToItem <> nil) or Assigned(ToItem) then Reduction(Sender, ToItem, CancelAction);
      end;
    paMove :
    	begin
        If (ToItem <> nil) or Assigned(ToItem) then Deplacement(Sender, ToItem, CancelAction);
      end;
    paDelete :
    	begin
    		CancelAction := false;
      end;
    paCellEnter :
    	begin
    		Affectation(Sender, ToItem, CancelAction);
    	end;
  else
    //CancelAction := true;
  end ;

end;


procedure TMonPlanning.BQuitterClick(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
  close;
{$ENDIF}
end;

procedure TMonPlanning.GestionSynchroDate(Item: tob);
var
  DateArrivee: TDateTime;
  DateDepart: TDateTime;
  i: integer;
  QDosRes: TQuery;
  QDossier: TQuery;
  TobDosRes: Tob;
  TobDossier: Tob;
begin

  DateArrivee := iDate1900;
  DateDepart := iDate1900;

  //MAJ dates dossier selon dates ligne;
  TobDosRes := Tob.create('les lignes', nil, -1);
  QDosres := OpenSQL('Select BEP_DATEFIN,BEP_DATEDEB from BTEVENPLAN where BEP_CODEEVENT="' + Item.getvalue('BEP_CODEEVENT') + '"', true,-1,'',true);
  try
    TobDosRes.LoadDetailDB('BTEVENPLAN', '', '', QDosRes, True)
  finally
    if QDosres <> nil then
      ferme(QDosres);
  end;

  if TobDosRes.Detail.Count > 0 then
  begin
    DateArrivee := TobDosRes.detail[0].getvalue('BEP_DATEDEB');
    DateDepart := TobDosRes.detail[0].getvalue('BEP_DATEFIN');
    for i := 1 to TobDosRes.Detail.Count - 1 do
    begin
      if TobDosRes.detail[i].getvalue('BEP_DATEDEB') < DateArrivee then
        DateArrivee := TobDosRes.detail[i].getvalue('BEP_DATEDEB');
      if TobDosRes.detail[i].getvalue('BEP_DATEFIN') > Datedepart then
        DateDepart := TobDosRes.detail[i].getvalue('BEP_DATEFIN');
    end;
  end;

  TobDosRes.free;
  TobDossier := Tob.Create('les Appels', nil, -1);
  QDossier := OpenSQL('Select AFF_DATEDEB,AFF_DATEFIN from AFFAIRE where AFF_AFFAIRE="' + Item.getvalue('BEP_CODEEVENT') + '"', true,-1,'',true);
  try
    TobDossier.LoadDetailDB('AFFAIRE', '', '', QDossier, False)
  finally
    if QDossier <> nil then
      ferme(QDossier);
  end;

  if TobDossier.detail.count > 0 then
  begin
    if TobDossier.Detail[0].getvalue('AFF_DATEDEBUT') <> DateArrivee then
      ExecuteSQL('update AFFAIRE set AFF_DATEDEBUT="' + USDATETIME(DateArrivee) + '" where AFF_AFFAIRE="' + item.getvalue('BEP_AFFAIRE') + '"');
    if TobDossier.Detail[0].getvalue('HDC_DATEDEPART') <> DateDepart then
      ExecuteSQL('update AFFAIRE set AFF_DATEFIN="' + USDATETIME(DateDepart) + '" where AFF_AFFAIRE="' + item.getvalue('BEP_AFFAIRE') + '"');
  end;

  TobDossier.free;

end;

procedure TMonPlanning.GestionDateLigne(Item: Tob);
begin
end;

//Recherche ???
procedure TMonPlanning.FormDestroy(Sender: TObject);
var
  i: integer;
begin

  PremierPassage := True;

  for i := 0 to NbOnglet - 1 do
      begin
      if feuille[i] <> nil            then feuille[i].Free;
      with TobPlannings[i] do
           begin
           if TobItems <> nil         then Tobfree(TobItems);
           if TobRes <> nil           then Tobfree(TobRes);
           if TobCols <> nil          then Tobfree(TobCols);
           if TobRows <> nil          then Tobfree(TobRows);
           if TobEvents <> nil        then Tobfree(TobEvents);
           if TobPeriodeDivers <> nil then Tobfree(TobPeriodeDivers);
           end;
      end;

  if TobItem  <> nil then FreeAndNil(TobItem);

  NumLigneTO := -1;

  Application.HintHidePause := 10000;

  TobItem := TOB.Create('BTEVENEMENTPLA', nil, -1);
  AddChampsSupItems (TobItem);

  FreeandNil(TobjFerie);

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Régis HARANG
Créé le ...... : 03/10/2001
Modifié le ... :   /  /
Description .. : Fonction activer ou désativer les boutons pendant le
Suite ........ : chargement du planning
Mots clefs ... : BOUTONS, ACTIVER
*****************************************************************}

procedure TMonPlanning.VoirToutBoutons(choix: Boolean);
begin

  BRecharger.Enabled := Choix;

  PageOnglet.Enabled := Choix;
  BPageSuiv.enabled  := Choix;
  BPagePrec.enabled  := Choix;

  Timer1.Enabled     := Choix;
  Timer2.Enabled     := Choix;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Régis HARANG
Créé le ...... : 16/10/2002
Modifié le ... :   /  /
Description .. : Mode : True creation
Suite ........ : Mode : False modification
Mots clefs ... :
*****************************************************************}

//function TMonPlanning.FonctionModif(Sender: TObject; Item: Tob): Boolean;
//begin
//end;

procedure TMonPlanning.MajEvenements(var MajPlanning: ThPlanningBTP);
begin

  //On bloque l'ensemble des boutons permettant d'avoir accès aux saisies...
  CodeDossierRech.Visible := False;
  Label11.Visible         := False;
  BCREAPPEL.Visible       := ModeSaisie;
  BChercher.Visible       := ModeSaisie;
  BLegende.Visible        := ModeSaisie;
  ExcelExport.Visible     := ModeSaisie;
  //
  MajPlanning.DisplayOptionCreation     := false;
  MajPlanning.DisplayOptionModification := false;
  MajPlanning.DisplayOptionSuppression  := false;

  if not ModeSaisie then
  begin
    MajPlanning.Autorisation :=  []; //gestion des actions étirée, déplacée,...
    Exit;
  end;

  BAffAppel2.Visible      := False;
  BAffAppel1.Visible      := False;
  BAffAppel.Visible       := False;
  //
  (* --
  MajPlanning.OnCreateItem  := Creation;
  MajPlanning.OnModifyItem  := Modification;
  --- *)
  //MajPlanning.OnDeleteItem  := Suppression;
  //

  if (ModePlanning = 'PMA') then
  begin
    CBTRAITE.Visible        := False;
    BAffAppel1.Visible      := ModeSaisie;
    //
    //MajPlanning.AddOptionPopup(4, '-');
    //
    //MajPlanning.AddOptionPopup(3, 'Suppression Evénement Parc/Matériel');
    //MajPlanning.AddOptionPopup(2, 'Modification Evénement Parc/Matériel');
    //
    //If (not FicAppel) and (ModeSaisie) then MajPlanning.AddOptionPopup(1, 'Création Evénement Parc/Matériel');
  end
  else if (ModePlanning = 'PCA') then
  begin
    CBTRAITE.Visible        := False;
    //BAffAppel2.Visible      := ModeSaisie;
    //
    //MajPlanning.AddOptionPopup(4, '-');
    //
    //MajPlanning.AddOptionPopup(3, 'Suppression Evénement Chantier');
    //MajPlanning.AddOptionPopup(2, 'Modification Evénement Chantier');
    //
    //If (not FicAppel) and (ModeSaisie) then MajPlanning.AddOptionPopup(1, 'Création Evénement Chantier');
  end
  else
  begin
    CBTRAITE.Visible          := True;
    BAffAppel.Visible         := ModeSaisie;

    //MajPlanning.AddOptionPopup(6, '-');

    //MajPlanning.AddOptionPopup(5, 'Edition Intervention');

    //MajPlanning.AddOptionPopup(7, '-');
    //MajPlanning.AddOptionPopup(8, 'Ajout d''une action');

    //MajPlanning.AddOptionPopup(4, '-');

    //MajPlanning.AddOptionPopup(2, 'Modification Affectation');
    //MajPlanning.AddOptionPopup(3, 'Visualisation Appel');

    //If (not FicAppel) and (ModeSaisie) then MajPlanning.AddOptionPopup(1, 'Création d''Appel');
  end;

  //MajPlanning.OnOptionPopup := OnPopup;

  MajPlanning.OnCopyItem    := CopyItem;
  MajPlanning.OnInitItem    := InitItem;
  MajPlanning.OnLink        := Link;
  MajPlanning.OnmoveItem    := CheckMoveItem;

  MajPlanning.OnAvertirApplication := AvertirApplication;

  MajPlanning.OnBeforeChange := P1BeforeChange;
  MajPlanning.OnDblClickSpec := DoubleClickSpec;

end;

procedure TMonPlanning.P1BeforeChange(const Item: TOB; const LaRessource: string; const LaDateDeDebut, LaDateDeFin: TDateTime; var Cancel: Boolean);
var codeRes,Coderes2 : string;
		DDeb : TDateTime;
    Item2 : TOB;
    LignePlan : integer;
    TOBresItem : TOB;
begin

	if (not ModeSaisie) then cancel := true;

  Cancel := False;

  if Item = nil then Exit;

  // contrôle de la possibilité de faire l'action
  if not isModifiable(Item) then Cancel := true;

  CodeRes		:= Item.GetValue('BPL_RESSOURCE');

  if Item.getString('BPL_ORIGINEITEM')='ABSPAIE' then
    Cancel := true
  else if Item.getString('BPL_ORIGINEITEM')='ACT-GRC' then
  begin
    if CodeRes <> LaRessource then Cancel := True;
  end;

  // Cas partis on ne peut pas etirer, reduire ou deplacer...
  //PlanningCurrent.DisplayOptionSuppressionLiaison := False;

end;

procedure TMonPlanning.Timer1Timer(Sender: TObject);
begin

	if TTWappels.Visible then exit;

  BRechargerClick(self);

end;

procedure TMonPlanning.GestionTimer;
begin

  Self.Timer1.Enabled := False;
  Self.Timer2.Enabled := False;

  if TobModelePlanning.Detail[PlanningCurrent.Tag].FieldExists('HPP_DUREETIMER') then
  begin
    if (TobModelePlanning.Detail[PlanningCurrent.Tag].GetValue('HPP_DUREETIMER') > 0) then
    begin
      Self.Timer1.Interval := TobModelePlanning.Detail[PlanningCurrent.Tag].GetValue('HPP_DUREETIMER') * 60000;
      Self.Timer1.Enabled  := True;
      Self.Timer2.Interval := TobModelePlanning.Detail[PlanningCurrent.Tag].GetValue('HPP_DUREETIMER') * 60000;
      Self.Timer2.Enabled  := True;
    end;
  end;

end;

procedure TMonPlanning.FormKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin

  if key = vk_F5 then BCalendrierClick(Self);

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Régis HARANG
Créé le ...... : 21/11/2002
Modifié le ... :   /  /
Description .. : Fonction qui géré les cliques dans les zones grisés du
Suite ........ : planning Ressources & evènements
Mots clefs ... : RESSOURCE;EVENEMENT
*****************************************************************}
procedure TMonPlanning.DoubleClickSpec(ACol, ARow: INTEGER; TypeCellule: TPlanningTypeCellule; T: TOB);
begin

	if (not ModeSaisie) then exit;

  OkAffect := False;

  //Si nous sommes sur une item vide à ce moment on crée la fiche qui va bien en fonction du
  //type de planning... Chaque planning permet de créer sa propre fiche...
  if PlanningCurrent.GetCurItem = Nil then
  Begin
    //Affichage de la grille d'affectation ressource...
    if TTWAppels.Visible then
      AffichageAffectationRessource
    else
      CreationItems(PlanningCurrent.GetCurItem, 'X');
  end
  else
    ModificationItem(PlanningCurrent.GetCurItem);

end;

procedure TMonPlanning.ExcelexportClick(Sender: TObject);
var fil : string;
    Rep : String;
begin

  Rep := GetparamSoc('SO_BTREPPLANNING');

  //pourquoi on vérifie pas que ce putain de rep existe ici ???????
  if DirectoryExists(Rep) then
  begin
    //Le nom du Planning ne doit pas contenir de signe spéciaux
    fil := TobModelePlanning.Detail[PlanningCurrent.Tag].GetValue('HPP_LIBELLE');
    fil := RemplaceCaracteresSpeciaux(fil, '_', True, True, True);
    Rep := Rep + '\' + fil;
    planningCurrent.ExportToExcel(true, rep, V_PGI.RunAfterExport);
  end
  else
    PGIError('Répertoire ' + rep + ' inexistant, veuillez modifier les paramètres de stockage planning', 'Erreur Export Excel');
   {+}

end;

{function TMonPlanning.RecupPeriode(NumeroPeriode: integer): string;
begin

  result := '';

  // GF 19/01/06 teste si tob <> Nil
  if Assigned(planningCurrent.TobPeriodeDivers) then
    if (NumeroPeriode >= 0) and (NumeroPeriode < planningCurrent.TobPeriodeDivers.detail.count) then
      Result := planningCurrent.TobPeriodeDivers.detail[NumeroPeriode].getstring('CC_CODE');

end;
}
procedure TMonPlanning.Link(Sender: TObject; TobSource, TobDestination: TOB; Option: THPlanningOptionLink; var Cancel: Boolean);
begin
	if not ModeSaisie then cancel := true;
end;

function TMonPlanning.ChargeLegende : boolean;
var i: integer;
    TobLeg  : tob;
begin

  result := true;

  TobLeg := Tob.create('LesEtats', nil, -1);
  if (ModePlanning = 'PMA') Then
    TobLeg.LoadDetailFromSQL('SELECT * FROM BTETAT WHERE BTA_TYPEACTION="PMA" AND BTA_ASSOSRES = "X"', False)
  else if (ModePlanning = 'PCA') Then
    TobLeg.LoadDetailFromSQL('SELECT * FROM BTETAT WHERE BTA_TYPEACTION="PCA" AND BTA_ASSOSRES = "X"', False)
  else
    TobLeg.LoadDetailFromSQL('SELECT * FROM BTETAT WHERE BTA_TYPEACTION="INT" AND BTA_ASSOSRES = "X"', False);

  // XP 07.06.2006
  //for i := 0 to uBtpEtatPlanning.BTP_LesEtats.Detail.Count - 1 do
  //  if uBtpEtatPlanning.BTP_LesEtats.Detail[i].GetString('BTA_ASSOSRES') = 'X' then
  //    TOB.Create('', TobEtat, -1).Dupliquer(uBtpEtatPlanning.BTP_LesEtats.Detail[i], True, True);

  SetLength(Libelle, TobLeg.Detail.Count);
  SetLength(Couleur, TobLeg.Detail.Count);

  TTWLegende.Caption := 'Legende';
  TTWLegende.Height := (TobLeg.Detail.Count + 1) * 20 + 30;

  {Taille Mini du panneau}
  if TTWLegende.Height < 165 then TTWLegende.Height := 165;

  if TobLeg.detail.count = 0 then
  begin
    result := false;
    TobLeg.free;
    exit;
  end;

  for i := 0 to TobLeg.detail.count - 1 do
  begin
    //Gestion de la Couleur
    Couleur[I]              := TLabel.Create(Self);
    Couleur[I].Parent       := TTWlegende;
    Couleur[I].Name         := 'Couleur' + IntToStr(I);
    Couleur[I].Left         := 10;
    Couleur[I].Height       := 20;
    Couleur[I].caption      := ' ' + TobLeg.Detail[I].GetValue('BTA_BTETAT')+' ';
    Couleur[I].Color        := StringToColor(TobLeg.Detail[I].GetValue('BTA_COULEURFOND'));
    Couleur[I].Font.Color   := StringToColor(TobLeg.Detail[I].GetValue('BTA_COULEUR'));
    Couleur[I].Transparent  := false;

    if I = 0 then
      Couleur[I].Top := 5
    else
      Couleur[I].Top := Couleur[I-1].Top + 20;

    //Gestion du Libelle
    Libelle[I]              := TLabel.Create(Self);
    Libelle[I].Parent       := TTWlegende;
    Libelle[I].Name         := 'Libelle' + IntToStr(I);
    Libelle[I].caption      := TobLeg.Detail[I].GetValue('BTA_ABREGE');
    Libelle[I].Left         := Couleur[I].left + Couleur[I].Width + 10;
    Libelle[I].Transparent  := true;
    if I = 0 then
     Libelle[I].Top := 5
    else
     Libelle[I].Top := Couleur[I-1].Top + 20;
  end;

  TobLeg.free;

end;

//Chargement du menu en fonction de la Tablette Cadencement
procedure TMonPlanning.ChargeMenuCadencement;
var
  QQ: TQuery;
  StSql: string;
  TobCad: Tob;
  I: Integer;
  ITemMenCad: TmenuItem;
begin

  //Cadencement := TobModelePlanning.Detail[PlanningCurrent.Tag].Getvalue('HPP_CADENCEMENT');

  if (ModePlanning = 'PMA') then
    StSql := 'SELECT CO_CODE, CO_LIBELLE FROM COMMUN WHERE CO_TYPE="BCA" AND CO_LIBRE LIKE "PMA%" ORDER BY CO_CODE'
  else if (ModePlanning = 'PCA') then
    StSql := 'SELECT CO_CODE, CO_LIBELLE FROM COMMUN WHERE CO_TYPE="BCA" AND CO_LIBRE LIKE "%PCA%" ORDER BY CO_CODE'
  else
    StSql := 'SELECT CO_CODE, CO_LIBELLE FROM COMMUN WHERE CO_TYPE="BCA" ORDER BY CO_CODE';

  QQ := OpenSQL(StSQl, True,-1,'',true);

  if QQ.Eof then
  begin
    MenuCadencement.Free;
    bcadencement.visible := false;
    Ferme(QQ);
    Exit;
  end;

  TobCad := Tob.create('Les Cadencements', nil, -1);
  TobCad.LoadDetailDB('COMMUN', '', '', QQ, False);
  Ferme(QQ);

  for I := 0 to TobCad.Detail.count - 1 do
  begin
    ITemMenCad := TmenuItem.Create(Self);
    ITemMenCad.Caption := TobCad.Detail[I].getValue('CO_LIBELLE');
    ITemMenCad.OnClick := ChangeCadencement;
    if Cadencement = TobCad.Detail[I].getValue('CO_CODE') then
    begin
      MenuCadencement.Items.Add(ITemMenCad);
      //MenuCadencement.Items.Checked := True;
      ITemMenCad.Checked := true;
    end
    else
    Begin
      MenuCadencement.Items.Add(ITemMenCad);
      ITemMenCad.Checked := False;
    end;
  end;

  TObCad.free;

end;

//Controle de l'existence du type d'action
Function TMonPlanning.Control_BtEtat : boolean;
var QQ    : TQuery;
    StSQL : string;
Begin

  Result := false;
  
  if TypeAction = '' then  Exit;

  StSQl:= 'SELECT * FROM BTETAT WHERE BTA_TYPEACTION="INT" AND BTA_ASSOSRES = "X"';
  StSQL := StSQL + ' AND BTA_BTETAT="' + TypeAction + '" ';

  QQ := OpenSql(StSQL,True,-1,'',true);

  //Création de la Tob Tache
  if not QQ.eof then
  Begin
    TobEtat.SelectDB ('BTETAT',QQ);
    Result := True;
  end;

  ferme(QQ);

end;

procedure TMonPlanning.BAffAppelClick(Sender: TObject);
begin
  Ok_ModifItem := False;

  FicAppel := False;

  AfficheMultiCritere;

end;

//Procedure de recherche de l'équipe associée à une ressource
function TMonPlanning.RechEquipe(Equipe : string) : Boolean;
var StSql     : string;
    ZoneLue   : string;
    Rep       : string;
    Critere   : string;
    NomChamp  : string;
    I         : Integer;
    QQ        : TQuery;
    TobEquipe : Tob;
    TobLEquipe: TOB;
    TobItemEq : Tob;
begin

  Result := False;

  if TobItem = nil then Exit;

  StSQL := 'SELECT * FROM RESSOURCE ';
  StSQL := StSQL + 'WHERE ARS_EQUIPERESS="' + Equipe + '" ';
  StSQL := StSQL + 'ORDER BY ARS_TYPERESSOURCE';

  QQ := OpenSQL(StSql, True,-1,'',true);
  if QQ.eof then
  begin
    Ferme(QQ);
    Exit;
  end;

  if PGIAsk(TraduireMemoire('La Ressource appartient à l''Equipe ' + Equipe + '.' + Chr(10) + 'Voulez-vous gérer l''equipe ?'), rep) = mrno then
  begin
    Ferme(QQ);
    exit;
  end;

  Result := true;

  //Création de la Tob equipe uniquement si Equipe existe et
  //si réponse oui à la question
  TobEquipe := Tob.Create('Equipe', nil, -1);
  TobEquipe.LoadDetailDB('LEQUIPE', '', '', QQ, False);

  Ferme(QQ);

  //Vérification si les info de la tob n'existe pas déjà dans le grille
  for i := 0 to TobEquipe.detail.count - 1 do
  begin
    TobLEquipe := TobEquipe.Detail[I];
    //Creation d'une tob Item Equipe temporaire à Partir du TobItem
    TobItemEq := TOB.Create('Item Equipe', nil, -1);
    TobItemEq.Dupliquer(TobItem, False, True);
    //if not ControleDansLaGrille then Continue;
    //Si la ressource ne se trouve pas déjà dans la grille on charge les info de la ressource
    TobItemEq.PutValue('NUMEVENT', 0);
    TobItemEq.PutValue('BPL_RESSOURCE',    TobLEquipe.GetValue('ARS_RESSOURCE'));
    TobItemEq.PutValue('BPL_NOMRESSOURCE', TobLEquipe.GetValue('ARS_LIBELLE') + ' ' + TobLEquipe.GetValue('ARS_LIBELLE2'));
    TobItemEq.PutValue('BPL_VILLE',        TobLEquipe.GetValue('ARS_VILLE'));
    TobItemEq.PutValue('BPL_TYPERESSOURCE',TobLEquipe.GetValue('ARS_TYPERESSOURCE'));
    TobItemEq.PutValue('BPL_STYPERESSOURCE',TobLEquipe.GetValue('ARS_CHAINEORDO'));
    TobItemEq.PutValue('BPL_FONCTION',     TobLEquipe.GetValue('ARS_FONCTION1'));
    TobItemEq.PutValue('BPL_EQUIPERESS',   TobLEquipe.GetValue('ARS_EQUIPERESS'));
    TobItemEq.PutValue('BPL_EMAIL',        TobLEquipe.GetValue('ARS_EMAIL'));
    TobItemEq.PutValue('BPL_LIBFONCTION',  TobLEquipe.GetValue('AFO_LIBELLE'));
    TobItemEq.PutValue('BPL_STANDCALEN',   TobLEquipe.GetValue('ARS_STANDCALEN'));
    TobItemEq.PutValue('BPL_CALENSPECIF',  TobLEquipe.GetValue('ARS_CALENSPECIF'));
    //chargement des zones de la grille
    Critere := ColNames;
    NomChamp := (Trim(ReadTokenSt(Critere)));
    while (NomChamp <> '') do
    begin
      TobItemEq.PutValue(NomChamp, TobLEquipe.GetString(NomChamp));
      NomChamp := Trim(ReadTokenSt(Critere));
    end;
    //
    TobItem.InitValeurs(False);
    TobItem.Dupliquer(TobItemEq, False, True);
    //
    MajEventPlanning;
    //
    FreeAndNil(TobItemEq);
  end;

  FreeAndNil(TobEquipe);

end;

//Dessin de la grille et création de la TOB de remplissage
procedure TMonPlanning.DrawGrille(NbColFixe: Integer);
var I					 : Integer;
    Ind				 : Integer;
    TGril			 : Integer;
    Champ			 : String;
    NomChamp	 : String;
begin

  //Création de la Tob des Evènements
  TobEvenement.ClearDetail;

  //Création de la Tob Mère de Suppression de Lignes
  TobSuppression := Tob.create('la SUPPRESSION', nil, -1);

  ColNames := '';

  //Initialisation de la Grille
  GrilleSelection.RowCount := 2;
  GrilleSelection.ColCount := NbColFixe;

  For i := 0 To GrilleSelection.RowCount -1 do
  Begin
    For Ind := 0 TO GrilleSelection.ColCount - 1 do
    Begin
      GrilleSelection.Cells[Ind, I] := '';
    end;
  end;

  GrilleSelection.FixedRows := 1;
  GrilleSelection.FixedCols := 0;
  TTWappels.Width := 305;

  TGril := TTWappels.Width;

  //Chargement de l'entete en fonction des paramètres planning
  for I := 4 to NbColFixe + 3 do
  begin
    Ind := I - 4;
    Champ := TobModelePlanning.Detail[PlanningCurrent.tag].GetValue('HPP_LIBCOLENT' + IntToStr(I));
    if Champ <> '' then
    begin
      NomChamp := LectLibCol('CC', 'BLI', Champ, 'CC_LIBRE');
      GrilleSelection.Cells[Ind, 0] := RechDom('BTLIBCOLPLANNING', Champ, True);
      GrilleSelection.ColWidths[Ind] := TGril div NbColFixe;
    end
    else
    begin
      NomChamp := 'ARS_LIBELLE';
      GrilleSelection.Cells[0, 0] := 'Ressource';
      GrilleSelection.ColWidths[0] := TGril div NbColFixe;
      break;
    end;
    if ColNames = '' then
      ColNames := Nomchamp
    else
      ColNames := ColNames + ';' + Nomchamp;
  end;

end;

procedure TMonPlanning.BVALIDAFFClick(Sender: TObject);
begin

   ValidationAffectation;

end;

procedure TMonPlanning.BTSupprimeClick(Sender: TObject);
begin

  SupprimeAffectation;

end;


procedure TMonPlanning.BAnnulAffClick(Sender: TObject);
begin

  BAffAppel1.Down    := False;

  if ModePlanning = 'PLA' then
  begin
    CodeDossierRech.text := '';
    TobEvenement.ClearDetail;
    TobEvenement.PutValue('AFFAIRE', '');
    BAffAppel.Down   := True;
  end;

  BAffAppel2.Down   := False;

  TTWappels.Visible := False;

  if FicAppel then
  begin
    BAffAppel.down    := False;
    BAffAppel.Visible := True;
    Exit;
  end;

  AfficheMultiCritere;

end;

procedure TMonPlanning.GrilleSelectionKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var Cancel : Boolean;
begin

  //Gestion de la suppression de l'affectation
  if Key = VK_DELETE then SuppressionItem(TobItem);

end;

procedure TMonPlanning.ChangeCadencement(Sender: TObject);
var MyMenu    : TMenuItem;
    MenuLib   : string;
    i         : integer;
    IndexPage : Integer;
begin

  IndexPage := PageOnglet.ActivePageIndex;

  for i:=0 to MenuCadencement.Items.Count -1 do
  Begin
    MenuCadencement.items[i].Checked := false;
  end;

  MyMenu          := TmenuItem(Sender);
  MyMenu.Checked  := true;

  MenuLib := StrDelete(MyMenu.Caption, '&');

  Cadencement := RecTypeCommun('BCA', MenuLib);

  BCadencement.Down := False;

  if Cadencement = '' then exit;

  // Recharge les onglets du planning
  if (Typeplanning = 'PLA') Or (Typeplanning = 'PMA') then
    TobModelePlanning.Detail[PlanningCurrent.Tag].putvalue('HPP_CADENCEMENT', Cadencement)
  else
    TobModelePlanning.Detail[0].putvalue('HPP_CADENCEMENT', Cadencement);

  ChangePlanning(PageOnglet.ActivePageIndex);

end;

procedure TMonPlanning.BCREAPPELClick(Sender: TObject);
Var Nomfiche : string;
    LignePlan   : Integer;
begin

  if (ModePlanning = 'PCA') then
  begin
    if (not ModeSaisie) then exit;
    //
    CreationItems(nil, '-');
    //
  end
  else if (ModePlanning = 'PMA') then
    AglLanceFiche ('BTP','BTEVENTMAT','','','ACTION=CREATION;PLANNING')
  else
    AglLanceFiche ('BTP','BTAPPELINT','','','ACTION=CREATION;PLANNING');

  BCREAPPEL.Down := False;

  ChargeItems;

  AfficheMultiCritere;

end;

procedure TMonPlanning.BChercherClick(Sender: TObject);
var TobResultat	    : Tob;
    TobRecherche    : TOB;
    Ind             : Integer;
    ZoneRech        : String;
    ZoneName        : string;
    Z1, Z2, Z3, Z4  : string;
    FamilleRess     : String;
begin

  if PlanningCurrent.TobItems = nil then exit;
     
  //On va d'abord gérer la tob d'envoi
  TobRecherche := TOB.Create('RECHERCHE', nil, -1);

  TobRecherche.AddChampSupValeur('MODEPLANNING', ModePlanning);
  TobRecherche.AddChampSupValeur('TYPERESSOURCE', '');

  FamilleRess := '';

  //Le paramétrage va déterminer le type de ressource que l'on va permettre en reccherche
  if ModePlanning = 'PLA' then
  begin
    if PlanningCurrent.TypePlanning = 'PTR' then
      FamilleRess := PlanningCurrent.CodeOnglet
    else
    begin
      if PlanningCurrent.Famres <> 'GEN' then FamilleRess := PlanningCurrent.FamRes;
    end;
    IF FamilleRess <> '' then  TobRecherche.PutValue('TYPERESSOURCE', 'ARS_TYPERESSOURCE=' + FamilleRess);
  end
  else if ModePlanning = 'PCA' then
  begin
    if PlanningCurrent.TypePlanning = 'PFO' then
    begin
      FamilleRess := PlanningCurrent.CodeOnglet;
      IF FamilleRess <> '' then  TobRecherche.PutValue('TYPERESSOURCE', 'ARS_FONCTION1=' + FamilleRess);
    end
    else
    begin
      if PlanningCurrent.Famres <> 'GEN' then FamilleRess := PlanningCurrent.FamRes;
      IF FamilleRess <> '' then  TobRecherche.PutValue('TYPERESSOURCE', 'ARS_TYPERESSOURCE=' + FamilleRess);
    end;
  end
  else if ModePlanning = 'PMA' then
  begin
    if PlanningCurrent.TypePlanning = 'PFM' then
      FamilleRess := PlanningCurrent.CodeOnglet
    else
    begin
      if PlanningCurrent.Famres <> 'GEN' then FamilleRess := PlanningCurrent.FamRes;
      if Copy(FamilleRess,0,1) = '"' then FamilleRess := Copy(FamilleRess, 2, Length(FamilleRess)-2);
      if Pos(',', FamilleRess) > 0 Then FamilleRess := '';
    end;
    IF FamilleRess <> '' then TobRecherche.PutValue('TYPERESSOURCE', 'BMA_CODEFAMILLE=' + FamilleRess);
  end;

  TobRecherche.AddChampSupValeur('AFFAIRE',   '');
  TobRecherche.AddChampSupValeur('RESSOURCE', '');
  TobRecherche.AddChampSupValeur('MATERIEL',  '');
  TobRecherche.AddChampSupValeur('PHASE',     '0');

  //appel de la fenêtre
  TheTOB := TobRecherche;

  AglLanceFiche ('BTP','BTRECHPLANNING','','','PLANNING');

  TheTob := nil;
  //
  ZoneRech := '';
  ZoneName := '';
  //
  Z1 := TobRecherche.GetValue('AFFAIRE');
  //
  if ModePlanning <> 'PMA' then
    Z2 := TobRecherche.GetValue('RESSOURCE')
  else
    Z2 := TobRecherche.GetValue('MATERIEL');
  //
  Z3 := TobRecherche.GetValue('PHASE');

  //Si le code Affaire est Renseigné
  if Z1 <> '' then
  Begin
    if Z2 <> '' then //Si le code ressource/Matériel est renseigné
    begin
      if Z3 <> '0' then //si le Numéro de phase est renseigné
      Begin
        IF Modeplanning <> 'PMA' then
          TobResultat := PlanningCurrent.TobItems.FindFirst(['BPL_AFFAIRE','BPL_RESSOURCE','BPL_NUMPHASE'], [Z1, Z2, Z3], True)
        else
          TobResultat := PlanningCurrent.TobItems.FindFirst(['BPL_AFFAIRE','BPL_MATERIEL','BPL_NUMPHASE'], [Z1, Z2, Z3], True);
      end
      else
      begin
        IF Modeplanning <> 'PMA' then
          TobResultat := PlanningCurrent.TobItems.FindFirst(['BPL_AFFAIRE','BPL_RESSOURCE'], [Z1, Z2], True)
        else
          TobResultat := PlanningCurrent.TobItems.FindFirst(['BPL_AFFAIRE','BPL_MATERIEL'], [Z1, Z2], True);
      end;
    end
    else
    begin
      If Z3 <> '0' then //Si le Numéro de phase est renseignée
        TobResultat := PlanningCurrent.TobItems.FindFirst(['BPL_AFFAIRE','BPL_NUMPHASE'], [Z1, Z3], True)
      else
      begin
        TobResultat := PlanningCurrent.TobItems.FindFirst(['BPL_AFFAIRE'], [Z1], True)
      end;
    end;
  end
  else
  begin
    if Z2 <> '' then //si le code Ressource/matériel est renseigné
    begin
      if Z3 <> '0' then //si le numéro de phase est renseigné
      Begin
        IF Modeplanning <> 'PMA' then
          TobResultat := PlanningCurrent.TobItems.FindFirst(['BPL_RESSOURCE','BPL_NUMPHASE'], [Z2, Z3], True)
        else
          TobResultat := PlanningCurrent.TobItems.FindFirst(['BPL_MATERIEL','BPL_NUMPHASE'], [Z2, Z3], True);
      end
      else
      begin
        IF Modeplanning <> 'PMA' then
          TobResultat := PlanningCurrent.TobItems.FindFirst(['BPL_RESSOURCE'], [Z2], True)
        else
          TobResultat := PlanningCurrent.TobItems.FindFirst(['BPL_MATERIEL'], [Z2], True);
      end;
    end
    else
    begin
      if Z3 <> '0' then //Si le numéro de phase est renseigné
      Begin
        TobResultat := PlanningCurrent.TobItems.FindFirst(['BPL_NUMPHASE'], [Z3], True);
      end
      else
      begin
        TobResultat := nil;
      end;
    end;
  end;

  //if ZoneRech <> '' then TobResultat := PlanningCurrent.TobItems.FindFirst([ZoneName], [ZoneRech], True);

  if TobResultat <> nil then PlanningCurrent.GoToItem(Tobresultat);

  FreeAndNil(TobRecherche);

end;

function TMonPlanning.MajEvtAffectation(ModeOption: String): Boolean;
Var DFin			 	: TDateTime;
	  DDebut      : TDateTIme;
    DateDeb     : TDateTime;
    DateFin     : TDateTime;
    HeureDEb    : TTime;
    HeureFin    : TTime;
    SDateDeb    : TDateTime;
    SDateFin    : TDateTime;
    SHeureDeb   : TDateTime;
    SHeurefin   : TDateTime;
    OldRessource: String;
	  NumEvenement: Integer;
    NumAppel 	 	: String;
    Rep			 	  : String;
    Equipe      : String;
    StSql,SQLB  : String;
    TobParam    : Tob;
    QQ			 	  : TQuery;
    sender		 	: Tobject;
    NbHeure     : Double;
    Temp        : Integer;
    newDate     : TdateTime;
    An, Mois, Jour : Word;
    H, M, S, Msec   : Word;
    Delta       : Integer;
begin

  Result          := False;

	if not IsModifiable (TobItem) then exit;

  CodeRes		      := TobItem.GetValue('BPL_RESSOURCE');
  OldRessource    := PlanningCurrent.GetCurItem.GetValue('BPL_RESSOURCE');
  //
  NumEvenement    := TobItem.GetValue('BPL_IDEVENT');
  Equipe          := TobItem.GetString('BPL_EQUIPERESS');
  //
  NumAppel        := TobItem.GetString('BPL_AFFAIRE');

  CodeDossierRech.text := NumAppel;

  //control si l'appel n'est pas terminé sinon on fait rien
  if NumAppel <> '' then
  Begin
    IF (TobItem.Getvalue('BPL_CODEETAT') = 'TER') Or (TobItem.Getvalue('BPL_CODEETAT') = 'REA') Then
    Begin
      AfficheErreur('BTPLANNING', '14', 'Modification d''affectation');
      exit;
    end;
  end;

  DateFin         := StrToDateTime(TobItem.GetValue('BPL_DATEAFFFIN'));
  DecodeDateTime(DateFin, an, Mois, Jour, H, M, S, Msec);
  HeureFin        := EncodeTime(H, M, 0, 0);

  //Avant de faire quoi que ce soit il faut ajouter +1 à l'heure de fin
  if Cadencement = '001' then // gestion par 1/4 Heure
  begin
    HeureFin := IncMinute (HeureFin, 15);
  end else if Cadencement = '002' then // gestion par 1/2 Heure
  begin
    HeureFin := IncMinute (HeureFin,30);
  end else if Cadencement = '003' then // gestion par heure
  begin
    HeureFin := IncHour (HeureFin, 1);
  end else if Cadencement = '004' then // gestion par 1/2 Journée
  Begin
    if heureFin > GetFinMatinee then
    begin
      Delta :=  MinutesBetween(heureFin,GetFinApresMidi);
    end else
    begin
      Delta := MinutesBetWeen(HeureFin,GetFinMatinee);
    end;
    HeureFin := incminute (HeureFin, Delta);
  end else if Cadencement = '005' then // gestion par Journée
  begin
    // pas de decalage dans ce cas la
  end else if Cadencement = '008' then // gestion par Période
  begin
    {TobTmp.PutValue('BEP_DATEFIN', TobTmp.GetValue('BEP_DATEFIN'));}
  end;

  DateFin         := Trunc(DateFin) + HeureFin;

  //Chargement des nouvelles valeur debut et fin
  DateDeb         := StrToDateTime(TobItem.GetValue('BPL_DATEAFFDEB'));
  DecodeDateTime(DateDeb, an, Mois, Jour, H, M, S, Msec);
  HeureDeb        := EncodeTime(H, M, 0, 0);

  TobItem.PutValue('BPL_DATEDEB', DateTimeToStr(DateDeb));
  TobItem.PutValue('BPL_DATEFIN', DateTimeToStr(DateFin));

  TobItem.PutValue('BPL_HEUREDEB', TimeToStr(HeureDeb));
  TobItem.PutValue('BPL_HEUREFIN', TimeToStr(HeureFin));
  //
  NbHeure := CalculDureeEvenement(DateDeb, DateFin);
  TobItem.PutValue('BPL_DUREE', NbHeure);
  //
  Result := True;

  //Controle s'il n'y a pas de changement d'affectation
  if Coderes <> OldRessource then
  Begin
    if PgiAsk('Confirmez-vous le changement d''affectation ?', rep)=mryes then
    Begin
      //FV1 : 12/11/2015 - FS#1691 - SERVAIS : Pb en saisie de planning si déplacement d'un appel.
      ModificationAction;
      exit;
    end
    else
    Begin
      TobItem.PutValue('BPL_RESSOURCE', OldRessource);
      result := false;
      exit;
    end;
  end;

  //gestion de la suppression ou de la modification d'un évènement lié à un appel
  TobParam := Tob.Create('AppelActionEvenement', nil, -1);

  //vérification si un ou plusieur évènement pour cet appel
  if NumAppel <> '' then
  Begin
    StSql := 'SELECT BEP_CODEEVENT FROM BTEVENPLAN WHERE BEP_AFFAIRE = "' + NumAppel + '"';
    QQ := OpenSQL(StSQL, True,-1,'',false);
    if not QQ.eof then
    begin
      If QQ.RecordCount > 1 then
      Begin
        TobParam.AddChampSupValeur('RETOUR', -1);
        TobParam.AddChampSupValeur('ACTION', ModeOption);
        TobParam.AddChampSupValeur('EQUIPE', Equipe);
        TheTob := TobParam;
        AglLanceFiche ('BTP','BTSUPPEVENT','','','');
        TheTob := nil;
        if TOBParam.getValue('RETOUR')= 0 then
        begin
          TOBParam.free;
          result := false;
          exit;
        end;
      end
      else
      Begin
        TobParam.AddChampSupValeur('RETOUR', 1);
        TobParam.AddChampSupValeur('ACTION', ModeOption);
        TobParam.AddChampSupValeur('EQUIPE', Equipe);
        MajEventBEP;
        TobEvenement.SetAllModifie(True);
        TobEvenement.InsertOrUpdateDB(true);
        Ferme(QQ);
        Exit;
      end;
    end;
  end
  else
  begin
    TobParam.AddChampSupValeur('RETOUR', 1);
    TobParam.AddChampSupValeur('ACTION', ModeOption);
    TobParam.AddChampSupValeur('EQUIPE', Equipe);
    MajEventBEP;
    TobEvenement.SetAllModifie(True);
    TobEvenement.InsertOrUpdateDB(true);
    Ferme(QQ);
    Exit;
  end;

  Ferme(QQ);

	 //Traitement de Modification ou déplacement (1 - ensemble de l'affectation, 2 - Equipe, 3 - Item Sélectionnées)
   StSQL := 'UPDATE BTEVENPLAN SET BEP_DATEDEB="' + USDATETIME(dateDeb) + '", ';
   StSQL := StSQL + 'BEP_DATEFIN="'   + USDATETIME(datefin) + '", ';
   StSQL := StSQL + 'BEP_HEUREDEB="'  + USTIME(HeureDeb) + '", ';
   StSQl := StSQl + 'BEP_HEUREFIN="'  + USTIME(HeureFin) + '", ';
   StSQl := StSQl + 'BEP_DUREE='      + StringReplace(FloatToStr(NbHeure),',','.',[rfReplaceAll]) + ' ';

  If TobParam.GetValue('RETOUR') = 1 then
  Begin
    SqlB := 'WHERE BEP_AFFAIRE ="' + NumAppel + '"';
    StSql := StSQl +SqlB;
  end
  else if TobParam.GetValue('RETOUR') = 2 Then
  Begin
   		SqlB := 'WHERE BEP_AFFAIRE="' + NumAppel + '" AND BEP_EQUIPERESS="'+ Equipe+ '"';
      StSQL := StSQl + SqlB;
  end
  Else If TobParam.GetValue('RETOUR') = 3 Then
  Begin
   		SqlB := ' WHERE BEP_CODEEVENT="' + IntToStr(NumEvenement) + '"';
      StSQL := StSQl + SqlB;
  end;

   ExecuteSQL(StSql);

  // mise à jour de l'appel
  StSql := 'UPDATE AFFAIRE SET AFF_DATEREPONSE="'+ USDateTIME(DDebut)+'" WHERE AFF_AFFAIRE="'+NumAppel+'"';
  ExecuteSQL(StSql);

  CodeDossierRech.text := '';

  TobParam.free;

  if OkEnvoiMail then	EnvoieMail ('Modification Intervention N°' + IntToStr(NumEvenement));

end;

Procedure TMonPlanning.SuppressionItemPlanning(CritereLecture : String);
Var TobDelete : Tob;
    TobEvt		: Tob;
    StSql			: string;
    QQ				: TQuery;
    NumEvent	: String;
    i					: Integer;
Begin

  //Lecture en fonction de l'option choisie
  StSql := 'SELECT * FROM BTEVENPLAN ' + CritereLecture;
  QQ := OpenSQL(StSQL, True,-1,'',true);

  TobEvt := Tob.Create('Tob_Evenements', nil, -1);
  TobEvt.LoadDetailDB('BTEVENPLAN', '', '', QQ, False);

  Ferme(QQ);

  for i := 0 to TobEvt.Detail.Count -1 do
  Begin
    NumEvent  := TobEvt.detail[i].getValue('BEP_CODEEVENT');
    TobDelete := PlanningCurrent.TobItems.FindFirst(['BPL_IDEVENT'], [NumEvent], True);
    if TobDelete <> nil then
    Begin
     PlanningCurrent.DeleteItem(TobDelete);
    end;
  end;

  TobEvt.free;

  //Suppression de l'évènement du planning
  StSql := 'DELETE FROM BTEVENPLAN ' +  CritereLecture;
  ExecuteSQL(StSql);

  if OkEnvoiMail then	EnvoieMail ('Suppression Action N°' + NumEvent);

end;

Procedure TMonPlanning.TraitSuppItem(CodeRes : String; NumAppel : String; NumEvenement : String);
Var QQ		: TQuery;
		StSQL	: String;
    I			: Integer;
Begin

  //Vérification si la ressource est la ressource principale
  StSql := 'SELECT * FROM BTEVENPLAN WHERE BEP_AFFAIRE="' + NumAppel + '"';
  QQ := OpenSQL(StSQL, True,-1,'',true);
  TRY
    if QQ.eof then
    Begin
      StSql := 'UPDATE AFFAIRE SET AFF_RESPONSABLE="", AFF_ETATAFFAIRE="ECO", AFF_DATEREPONSE=' + DateTimeToStr(iDate2099) + ' WHERE AFF_AFFAIRE ="' + NumAppel + '"';
      ExecuteSQL(StSql);
      exit;
    end;
    for I := 0 to QQ.recordcount -1 do
    Begin
      if QQ.FindField('BEP_RESSOURCE').AsString = CodeRes then
      begin
      if QQ.FindField('BEP_RESPRINCIPAL').AsString = 'X' then
      Begin
        StSql := 'UPDATE AFFAIRE SET AFF_RESPONSABLE="' + CodeRes + '", AFF_ETATAFFAIRE="AFF", AFF_DATEREPONSE=' + USDateTime(DatePlanDeb) + ' WHERE AFF_AFFAIRE ="' + NumAppel + '"';
        ExecuteSQL(StSql);
        exit;
      end;
      end
      else
      Begin
        if QQ.FindField('BEP_RESPRINCIPAL').AsString = 'X' then
        Begin
          StSql := 'UPDATE AFFAIRE SET AFF_RESPONSABLE="' + QQ.FindField('BEP_RESPRINCIPAL').AsString + '", AFF_ETATAFFAIRE="AFF", AFF_DATEREPONSE=' + UsDateTime(DatePlanDeb) + ' WHERE AFF_AFFAIRE ="' + NumAppel + '"';
          ExecuteSQL(StSql);
          exit;
        end;
      end;
    end;
  FINALLY
    Ferme(QQ);
  END;

end;

procedure TMonPlanning.GrilleSelectionClick(Sender: TObject);
Var Nbheure : Double;
begin

  if GrilleSelection.row = 1 then exit;
                 
  if CreerPar <> 'TAC' then
     BTaches.Visible      := False
  else
     BTaches.Visible      := True;

	if Not ficappel then
  Begin
    //CodeDossierRech.text  := TobAppel.Detail[GrilleSelection.Row-1].GetString('AFF_AFFAIRE');
    //Descriptif.Text 	    := TobAppel.Detail[GrilleSelection.Row-1].GetString('AFF_DESCRIPTIF');
    if FenDescriptif.visible then FenDescriptif.SetFocus;
  end
  else
  Begin
    Nbheure  := TobEvenement.Detail[GrilleSelection.Row-1].GetValue('BEP_DUREE');
    HeureDeb := TobEvenement.Detail[GrilleSelection.Row-1].GetValue('BEP_HEUREDEB');
    HeureFin := TobEvenement.Detail[GrilleSelection.Row-1].GetValue('BEP_HEUREFIN');
  end;

end;

procedure TMonPlanning.GrilleSelectionDblClick(Sender: TObject);
Var DateDeb : TDatetime;
    Datefin : TDateTime;
begin

  if ficappel then SetInfosAffectation(TobEvenement.Detail[GrilleSelection.Row - 1]);

	if FenDescriptif.visible then FenDescriptif.SetFocus;

end;

Procedure TMonPlanning.TraitementEventParcMateriel(TOBL : TOB);
var StSQL     : string;
    NumEvent  : Integer;
    QQ        : TQuery;
    TobEvent  : TOB;
    DateFin   : TDateTime;
    DateImputation : String;
begin

  if TOBL = nil then Exit;

  NumEvent        := TOBL.GetValue('BEM_IDEVENTMAT');
  DateFin         := TOBL.GetValue('BEM_DATEFIN');
  DateImputation  := TOBL.GetValue('DATEIMPUTATION');

  if TOBL.GetValue('ACTION') = 'ECO' then
  begin
    //Traitement de passage en En-COURS ET MISE à JOUR DES CONSOS
    StSQL := 'Update BTEVENTMAT SET BEM_CODEETAT="' + TOBL.GetValue('ACTION') + '" WHERE BEM_IDEVENTMAT=' + IntToStr(NumEvent);
    if ExecuteSQL(StSQL)<> 1 then PGIError('La Mise à jour de l''état ne s''est pas effectué', 'Erreur Mise à jour Evènements Parc Matériels');
  end
  else
  begin
    //il n'y a aucun test sur le genere conso !!!
    //if not GereConso(TOBL.GetValue('BEM_BTETAT')) then exit;
    //
    //Traitement de passage en REALISE ET MISE à JOUR DES CONSOS
    StSQL := 'SELECT * FROM BTEVENTMAT WHERE BEM_IDEVENTMAT=' + IntToStr(NumEvent);
    QQ    := OpenSQL(StSQL, false);
    if Not QQ.Eof then
    begin
      TobEvent := Tob.Create('BTEVENTMAT',nil,-1);
      TobEvent.SelectDB(StSQL, QQ);
      TobEvent.AddChampSupValeur('DATEIMPUTATION', DateImputation);
    end;

    Ferme(QQ);

    TraitementRealise(TobEvent);

    FreeandNil(TobEvent);
  end;

end;

procedure TMonPlanning.Timer2Timer(Sender: TObject);

begin
 //if Not ficappel then BRefreshAllClick(Sender);
end;

procedure TMonPlanning.BTachesClick(Sender: TObject);
var StAffaire : String;
    StTiers   : String;
begin

  //StAffaire := TobAppel.Detail[GrilleSelection.Row-1].GetString('AFF_AFFAIREINIT');
  //StTiers   := TobAppel.Detail[GrilleSelection.Row-1].GetString('AFF_TIERS');

  if stAffaire = '' then
     Begin
     pgibox('Il n''y a pas d''affaire de référence sur cet appel', 'Appel Tache');
     exit;
     end;

  AGLLanceFiche('BTP', 'BTTACHES','','', 'BYAFFAIRE;ATA_AFFAIRE:'+StAffaire+';ATA_TIERS:'+StTiers);

end;



procedure TMonPlanning.BTypePlanningClick(Sender: TObject);
Var DateEnCours   : TDateTime;
    StContrat     : string;
    ParamPlanning : string;
    CodeAppel     : string;
    i             : integer;
    LTousPossible : boolean;
    OK_AppelDrawG : Boolean;
begin

  if ModePlanning = 'PMA'  then OK_AppelDrawG   := BAffAppel1.Down;
  if ModePlanning = 'PCA'  then OK_AppelDrawG   := BAffAppel2.Down;

  PremierPassage  := false;
  LTousPossible   := false;
  ParamPlanning   := Recupdefaut;

  DateEdit.Text := SaisieDate(StContrat, ParamPlanning, TypePlanning, LTousPossible,DateEdit.Text, true, true);

  if ParamPlanning = '' then
  begin
    PgiInfo('Aucun modèle de défini', 'Saisie Planning');
    BTypePlanningClick(Self);
    exit;
  end;

  FormDestroy(Sender);

  TheParamPlanning.text := ParamPlanning;

  BTypePlanning.Down := False;

  TousPossible          := LTousPossible;
  CodeDossierRech.text  := CodeAppel;

  DateSelectionee := StrToDate(DateEdit.Text);

  TobModelePlanning.free;
  TobOnglet.Free;

  if ModePlanning = 'PMA'  then BAffAppel1.Down := OK_AppelDrawG;
  if ModePlanning = 'PCA'  then BAffAppel2.Down := OK_AppelDrawG;

  if TypePlanning = '' then
  Begin
    HShowMessage('', TraduireMemoire('Attention !'), TraduireMemoire('Aucun modèle de planning n''a été défini'));
    BPageSuiv.Enabled   := False;
    BPagePRec.Enabled   := False;
    BCalendrier.Enabled := False;
    Close;
    exit;
  End;

  If not InitialisationPlanning then
  begin
     Close;
     exit;
  end;

  if (ModePlanning = 'PLA')  AND (BAffAppel.Down)  then AfficheMultiCritere;
  if (ModePlanning = 'PMA')  AND (BAffAppel1.Down) then AfficheMultiCritere;
  if (ModePlanning = 'PCA')  AND (BAffAppel2.Down) then AfficheMultiCritere;

end;

procedure TMonPlanning.CheckMoveItem(sender: Tobject; Item: TOB; var cancel: boolean);
begin
end;


procedure TMonPlanning.BTAFFMEMOClick(Sender: TObject);
begin
	FenDescriptif.Caption := 'Descriptif de l''évenement';
  FenDescriptif.Visible := true;
  DESCRIPTIF.visible := false;
  DESCRIPTIF.Enabled := false;
  DesAction.Enabled  := true;
  DesAction.visible := true;
end;

procedure TMonPlanning.ReChargeActionGRC (Auxiliaire : string ; NumAct : integer);
var indice : Integer;
    OneTOB,TOBEvt : Tob;
begin
	//DeleteItemGRC (Auxiliaire,NumAct);
  //
  OneTOB := TOB.Create ('LES ACTION GRC',nil,-1);
  //ChargeActionsGRCDetail (OneTOB,Idate1900,Idate2099,PlanningCurrent.AppelsTraites,Auxiliaire,NumAct);

  if OneTOB.detail.count > 0 then
  begin
    for Indice := 0 to OneTOB.detail.count -1 do
    begin
      TOBEvt := TOB.Create ('BTEVENPLAN',PlanningCurrent.TobItems,-1);
      //TransformeActionsEnItem (OneTOB.detail[indice],PlanningCurrent);
      CompleteTobItem (TOBEvt);
      PlanningCurrent.AddItem (TOBEvt) ;
      PlanningCurrent.InvalidateItem (TOBEvt);
    end;
  end;
  OneTOB.free;
end;

procedure TMonPlanning.ReChargeActionGRC (Item : TOB);
var Auxiliaire : string;
    Numaction : integer;
begin
  Auxiliaire:= Item.GetValue('CLIENTGRC');
  NumAction:= Item.GetValue('ACTIONSGRC');
  ReChargeActionGRC (Auxiliaire,NumAction);
end;

function TMonPlanning.CalculeDateEvtEtireReduit (Item : TOB; var DateD,DateF : TdateTime; var Delai : integer): boolean;
var NbHJour       : TDateTime;
    //
    HDeb          : TDateTime;
    HFin          : TDateTime;
begin

	result := true;

  if cadencement = '005' then // journée
  Begin
    Delai	:= 0;
    DateD := Trunc(DateD) + GetDebutMatinee;
    DateF := Trunc(DateF) + GetFinApresMidi;
  End
  Else if Cadencement = '004' then  // 1/2 Journée
  begin
    Delai	:= 0;
    DateF := GestionDateFinPourModif (DateF);
  End
  Else if Cadencement = '003' then // heure
  begin
    Delai	:= 0;
    DateF := GestionDateFinPourModif (DateF);
  End
  Else if Cadencement = '002' then // 1/2 heure
  begin
    Delai	:= 0;
    DateF := GestionDateFinPourModif (DateF);
  End
  Else if Cadencement = '001' then // 1/4 heure
  begin
    Delai	:= 0;
    DateF := GestionDateFinPourModif (DateF);
  end;

  if not ControleHoraires (DateD,DateF,'E',0, CodeRes) then
  begin
    result := false;
    exit;
  end;

  if ModePlanning = 'PMA' then
    delai := StrToInt(CalculNbHeure(DateD, DateF))
  Else
    Delai := round(CalculDureeEvenement(DateD,DateF));

end;

{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 22/09/2015
Modifié le ... :   /  /
Description .. : Procedure et fonction evenement Parc/Materiel
Mots clefs ... : PARC, MATERIEL
*****************************************************************}
function TMonPlanning.MajEvtParcMateriel (ModeOption: String; Item : TOB) : boolean;
var DateD       : TDateTime;
    DateF       : TdateTime;
    DDateFin    : TDateTime;
		Materiel    : string;
    NbJour      : string;
    Rep			 	  : String;
    NumEventMat : integer;
    NbHeure     : Double;
begin

  Result          := False;

  Materiel        := Item.GetValue('BPL_MATERIEL');
  NumEventMat     := Item.GetValue('BPL_CODEEVENT');
  DateD           := Item.GetValue('BPL_DATEDEB');
  DateF           := Item.GetValue('BPL_DATEFIN');
  NbJour			    := '';

  NbHeure         := Item.GetValue('BPL_DUREE');

  DDateFin        := PlanningCurrent.GetCurItem.GetValue('BPL_DATEFIN');

	if not IsModifiable (Item) then exit;

  //Recalcul de la Durée dans le cas de la non suppression
  if ModeOption <> 'S' then
  Begin
	   Heuredeb := 0;
	   HeureFin := 0;
  end;

  if ModeOption = 'D' then    // deplacement
  Begin
    if cadencement = '005' then // Journée
    Begin
      if not ControleHoraires (DateD,DateF,ModeOption,NbHeure,'') then  exit;
    End
    Else if Cadencement = '004' then  // 1/2 Journée
    begin
      Nbheure	    := 0;
      DateD       := CalculDateDebut (DateD);
      DateF       := GestionDateFinPourModif (DateF);
      if not ControleHoraires (DatePlanDeb,DatePlanFin,ModeOption,NbHeure,'') then exit;
    End
    Else if Cadencement = '003' then // Heure
    begin
      Nbheure	    := 0;
      DateF       := GestionDateFinPourModif (DateF);
      DateF       := Item.GetValue('BPL_DATEFIN');
      NbHeure     := Item.GetValue('BPL_DUREE');
      DatePlanFin := GestionDateFinPourModif (DateF);
      if not ControleHoraires (DateD,DateF,ModeOption,NbHeure,'') then exit;
      DateF := GestionDateFinPourEnreg (DateF);
    end
    else if cadencement = '001' then // 1/4 heure
    Begin
      DateF := GestionDateFinPourModif (DateF);
      if not ControleHoraires (DateD,DateF,ModeOption,NbHeure,'') then exit;
      DateF := GestionDateFinPourEnreg (DateF);
    end;
  end
  else if pos(ModeOption,'E;R') > 0  then // etirement ou reduction
  Begin
    if cadencement = '005' then // journée
    Begin
      Nbheure	  := 0;
      DateD := Trunc(DateD) + GetDebutMatinee;
      DateF := Trunc(DateF) + GetFinApresMidi;
    End
    Else if Cadencement = '004' then  // 1/2 Journée
    begin
      Nbheure	  := 0;
      DateF := GestionDateFinPourModif (DateF);
    End
    Else if Cadencement = '003' then // heure
    begin
      Nbheure	  := 0;
      DateF := GestionDateFinPourModif (DateF);
      if not ControleHoraires (DateD,DateF,ModeOption,0,'') then exit;
      DateF := GestionDateFinPourEnreg (DateF);
    End
    Else if Cadencement = '002' then // 1/2 heure
    begin
      Nbheure	  := 0;
      DateF := GestionDateFinPourModif (DateF);
      if not ControleHoraires (DateD,DateF,ModeOption,0,'') then exit;
      DateF := GestionDateFinPourEnreg (DateF);
    End
    Else if Cadencement = '001' then // 1/4 heure
    begin
      Nbheure	  := 0;
      DateF := GestionDateFinPourModif (DateF);
      if not ControleHoraires (DateD,DateF,ModeOption,0,'') then exit;
      DateF := GestionDateFinPourEnreg (DateF);
    end;
	end;

  //Traitement de la modification ou de la suppression d'un évènement libre
  if ModeOption = 'S' then DeleteEvenementMateriel;

end;


Procedure TMonPlanning.GererPlanningIntervention;
begin

  TypeAction          := GetParamSoc('SO_TYPEACTION');
  Ok_TpsReel          := GetParamSoc('SO_GesTempsReel');

  BCREAPPEL.Hint      := 'Création d''un Appel';

  BAffAppel.Visible   := True;
  BAffAppel1.Visible  := False;
  BAffAppel2.Visible  := False;

  BAffAppel1.Down     := False;
  BAffAppel2.Down     := False;
  BAffAppel.Down      := False;

  Btaches.Visible     := True;

  BDescriptif.Visible := True;

  if not Control_BtEtat then
  Begin
    AfficheErreur('BTPLANNING','12','Erreur d''affectation');
    Close;
    exit;
  end;

  //BTMail.Visible := True;

end;

Procedure TMonPlanning.GererPlanningMateriel;
begin

  TypeAction          := '';
  Ok_TpsReel          := False;

  BCREAPPEL.Hint      := 'Création d''un événement Parc/Matériel';
  BAffAppel1.Hint     := 'Liste des événements à réaliser';

  BAffAppel1.Visible  := True;
  BAffAppel.Visible   := False;
  BAffAppel2.Visible  := False;

  BAffAppel1.Down     := False;
  BAffAppel2.Down     := False;
  BAffAppel.Down      := False;

  BChercher.Visible   := False;
  CBTRAITE.Visible    := False;
  Btaches.Visible     := False;
  BDescriptif.Visible := False;

  //BTMail.Visible      := False;

end;

Procedure TMonPlanning.GererPlanningChantier;
begin

  TypeAction          := '';
  Ok_TpsReel          := False;

  BCREAPPEL.Hint      := 'Création d''un événement Chantier';
  BAffAppel2.Hint     := 'Liste des Chantiers à Affecter';

  BAffAppel.Visible   := False;
  BAffAppel1.Visible  := False;
  BAffAppel2.Visible  := False;

  BAffAppel1.Down     := False;
  BAffAppel2.Down     := False;
  BAffAppel.Down      := False;

  BChercher.Visible   := False;

  CBTRAITE.Visible    := False;

  Btaches.Visible     := False;

  BDescriptif.Visible := False;

  //BTMail.Visible      := False;

end;

Procedure TMonPlanning.ChargementOngletPlanning(TobOnglet : Tob);
Var Maxi    : Integer;
    i       : Integer;
    NModele : integer;
begin

  Maxi := 99;

  //On définit le nombre maxi d'onglet dont on aura besoin en fonction du nombre de modèle de planning
  if TobOnglet.detail.count < maxi then maxi := TobOnglet.detail.count;

  NbOnglet := maxi;

  for i := 0 to (NbOnglet - 1) do
  begin
    // Creation des onglets
    Feuille[i] := TTabSheet.Create(Self);
    if i < 10 then
      Feuille[i].name := 'Page0' + IntToStr(i)
    else
      Feuille[i].name := 'Page'  + IntToStr(i);
    //
    if TobOnglet.Detail[i].GetValue('LIBELLE') = '' then
      Feuille[i].Caption      := 'Non Attribué'
    else
      Feuille[i].Caption      := TobOnglet.Detail[i].GetValue('LIBELLE');
    //
    Feuille[I].PageControl     := PageOnglet;
    Feuille[I].Tag := I;

    // Creation des plannings
    Planning[i]                := THPlanningBTP.Create(Self);
    Planning[i].parent         := Feuille[i];
    Planning[i].Name           := Format('PLANNING%2.2d', [i]);
    Planning[i].Align          := alClient;
    Planning[i].AppelsTraites  := CBTRAITE.Checked;
    Planning[i].TypePlanning   := Typeplanning;
    Planning[i].CodeOnglet     := TobOnglet.Detail[i].GetValue('CODEONGLET');
    if TousPossible then
      Planning[i].FamRes       := TobModelePlanning.Detail[i].GetValue('HPP_FAMRES')
    else
    begin
      if TypePlanning = 'PFO' then
      begin
        if Fonction <> '' then //si planning de type PFO (Fonction) il faudra gérer le multi fonction ici
          Planning[i].FamRes     := TobOnglet.Detail[i].GetValue('CODEONGLET')
        else
          Planning[i].FamRes     := 'GEN';
      end
      else
      begin
        if FamMatGere <> '' then //si planning de type PSF (sous-Famille) il faudra gérer le multi famille ici
          Planning[i].FamRes     := TobOnglet.Detail[i].GetValue('CODEONGLET')
        else
          Planning[i].FamRes     := TobModelePlanning.Detail[0].GetValue('HPP_FAMRES');
      end;

    end;
    //
    Planning[i].Tag            := i;

    //Gestion du Menu Popup
    MajEvenements(Planning[i]);

    RechargePlanning(TobOnglet.Detail[i], i);
  end;

  // Charge les paramètres du premier planning
  NModele := 0;


  PlanningCurrent                 := Planning[NModele];

  Cadencement                     := TobModelePlanning.Detail[NModele].Getvalue('HPP_CADENCEMENT');


end;

Procedure TMonPlanning.ChargeItems;
var IndexPage : Integer;
begin

  IndexPage := PageOnglet.ActivePageIndex;

  if       TypePlanning ='PMA' then
    ChargeItemPlanning(TobPlannings[IndexPage], Planning[IndexPage], TobModelePlanning.Detail[IndexPage], StrToDateTime(DateEdit.Text))
  else if Pos(TypePlanning,'PFM;PPA') > 0 then
    ChargeItemPlanning(TobPlannings[IndexPage], Planning[IndexPage], TobModelePlanning.Detail[0], StrToDateTime(DateEdit.Text))
  else if Pos(TypePlanning,'PSF;PTA;PTR;PFO') > 0 then
    ChargeItemPlanning(TobPlannings[IndexPage], Planning[IndexPage], TobModelePlanning.Detail[0], StrToDateTime(DateEdit.Text))
  else
    ChargeItemPlanning(TobPlannings[IndexPage], Planning[IndexPage], TobModelePlanning.Detail[IndexPage], StrToDateTime(DateEdit.Text));

  Planning[IndexPage].Raffraichir;

  if Pos(TypePlanning,'PFM;PPA;PSF;PTA;PTR;PFO') > 0 then
    AfficheLFixe(PlanningCurrent, TobModelePlanning.Detail[0])
  else
    AfficheLFixe(PlanningCurrent, TobModelePlanning.Detail[IndexPage]);
  
end;

Procedure TMonPlanning.AfficheMultiCritere;
begin

  if Ok_ModifItem then
  begin
    Ok_ModifItem := False;
    BAffAppel.Down  := false;
    BAffAppel1.Down := false;
    BAffAppel2.Down := false;
    Exit;
  end;

  if ModePlanning = 'PMA'  then
    TraitementParcMateriel
  else if ModePlanning = 'PLA' then
    TraitementAffectationAppel
  else if ModePlanning = 'PCA' then
    TraitementPlancharge;

  BAffAppel.Down  := false;
  BAffAppel1.Down := false;
  BAffAppel2.Down := false;

end;


{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 04/03/2016
Modifié le ... :   /  /    
Description .. : Traitement d'affectation d'une intervention à partir d'une ou 
Suite ........ : plusieurs ressource
Mots clefs ... : 
*****************************************************************}
Procedure TMonPlanning.TraitementAffectationAppel;
Var ZoneRec : string;
begin

  if (not BAffAppel.Down)  Then Exit;

  if GetParamSoc('SO_BTAFFAIRERECHLIEU') then
    ZoneRec := AGLLanceFiche('BTP','BTMULAPPELSINT','AFF_AFFAIRE0=W','','STATUT=APP;ETAT=ECO;NOCHANGESTATUT;PLANNING')
  else
    ZoneRec := AglLanceFiche('BTP','BTMULAPPELS','','','STATUT=APP;ETAT=ECO;NOCHANGESTATUT;PLANNING');

  if zoneRec = '' then Exit;

  ChargementTobappel(ConstituStWhereIntervention(ZoneRec));

end;

Function TMonPlanning.ConstituStwhereIntervention(Critere : string) : string;
Var ZoneRec : string;
begin

  ZoneRec := Critere;

  Critere := (Trim(ReadTokenSt(ZoneRec)));
  if Critere <> '' then
  begin
    if Result <> '' then
      Result := Result + ' AND AFF_AFFAIRE = "' + Critere + '" '
    else
      Result := ' AFF_AFFAIRE = "' + Critere + '" ';
  end;

  if Result <> '' then
    Result := Result + 'AND AFF_ETATAFFAIRE = "ECO" '
  else
    Result := ' AFF_ETATAFFAIRE = "ECO" ';

  Critere:=(Trim(ReadTokenSt(ZoneRec)));
  if Critere <> '' then
  begin
    if Result <> '' then
      Result := Result + ' AND AFF_TIERS = "' + Critere + '" '
    else
      Result := ' AFF_TIERS = "' + Critere + '" ';
  end;

  if Result <> '' then Result := ' WHERE ' + Result;

end;

Procedure TMonPlanning.ChargementTobappel(StWhere : String);
Var StSQL   : string;
    QQ      : TQuery;
    TobAppel: TOB;
begin

  TobAppel := TOB.Create('LES APPELS',nil,-1);

  //Lecture de la table des appels pour chargement des informations....
  StSQL := 'SELECT * FROM AFFAIRE ' + StWhere;
  QQ    := OpenSQL(StSQL, True,1,'',True);

  if not QQ.Eof then
  begin
    TobAppel.SelectDB('AFFAIRE', QQ, False);
    ChargeInfoAppel(TobAppel);
    if ModePlanning = 'PLA' then ChargeIntervention;
  end;

  Ferme(QQ);

  FreeAndNil(TobAppel);

end;

//chargement de la TOBAppel avec les info de l'appel
procedure TMonPlanning.ChargeInfoAppel(TobAppel : Tob);
var QQ    		: TQuery;
    Auxiliaire: string;
    NomTiers  : string;
    StSQL			: string;
    InfoAppel	: String;
    Tel       : string;
    NoTache   : Integer;
begin

  if TobAppel = nil then exit;
  //
  OkAppelLibre          := True;
  //FicAppel			        := True;
  OKVisuAffect          := True;
  //
  Btaches.Visible       := False;
  BDescriptif.Visible   := False;
  CHKTACHES.Visible     := False;
  //
  TTWAppels.Caption     := 'Liste des Ressources Affectées';
  TTWAppels.Hint        := 'Liste des Ressources Affectées';

  CodeDossierRech.text  := TobAppel.GetString('AFF_AFFAIRE');
  Descriptif.Text 	    := TobAppel.GetString('AFF_DESCRIPTIF');
  Creerpar              := TobAppel.Getvalue('AFF_CREERPAR');
  AffaireInit           := TobAppel.Getvalue('AFF_AFFAIREINIT');
  CodeTiers             := TobAppel.Getvalue('AFF_TIERS');
  Contact               := TobAppel.getvalue('AFF_NUMEROCONTACT');

  NoTache               := TobAppel.Getvalue('AFF_NUMEROTACHE');

  //if Not FicAppel then exit;

  label11.Caption := InfoAppel;
  label11.Visible := true;

  if CreerPar <> 'TAC' then
    BTaches.Visible := False
  else
    BTaches.Visible := True;

  BValidAff.Visible := True;
  BRefreshAff.Visible := False;

  Responsable := TobAppel.GetValue('AFF_RESPONSABLE');

  if ModePlanning = 'PLA' then
  begin
    InfoAppel	:= 'Appel N° ' + CodeDossierRech.text + ' ';
    Label3.Caption := 'APPEL N°'      + TobAppel.GetValue('AFF_AFFAIRE1');
    Label3.Caption := Label3.caption  + TobAppel.GetValue('AFF_AFFAIRE2');
    Label3.Caption := Label3.caption  + TobAppel.GetValue('AFF_AFFAIRE3');
  end
  else
  begin
    InfoAppel	:= 'Chantier N° ' + CodeDossierRech.text + ' ';
    Label3.Caption := 'CHANTIER N°'   + TobAppel.GetValue('AFF_AFFAIRE1');
    Label3.Caption := Label3.caption  + TobAppel.GetValue('AFF_AFFAIRE2');
    Label3.Caption := Label3.caption  + TobAppel.GetValue('AFF_AFFAIRE3');
  end;

  Descriptif.Text:= TobAppel.GetValue('AFF_DESCRIPTIF');

  DateAppel := TobAppel.GetValue('AFF_DATEDEBUT');

  ChargeInfoTiers(CodeTiers, Auxiliaire, NomTiers);
  if NomTiers <> '' then InfoAppel := InfoAppel + 'de ' + NomTiers + ' ';
  //
  ChargeInfoAdresse(CodeDossierRech.text, Tel);
  if Tel <> '' then InfoAppel := InfoAppel + '(' + Tel + ')';
  //
  ChargeInfoContact(Auxiliaire);
  //
  ChargeInfoTypeAction(AffaireInit, CodeTiers, NoTache, TypeAction);

  if Not Control_BtEtat then
  Begin
    AfficheErreur('BTPLANNING','15','Erreur d''affectation');
    exit;
  end;

  HPanel1.Visible     := True;
  BValidAff.Visible   := True;
  BRefreshAff.Visible := False;

  TTWappels.Visible   := True;

end;

Procedure TMonPlanning.ChargeInfotiers(CodeTiers : string; Var Auxiliaire, NomTiers : String);
Var StSQL : String;
    QQ    : TQuery;
begin

  //Lecture Tiers
  StSQL := 'SELECT T_LIBELLE, T_AUXILIAIRE FROM TIERS ';
  StSQL := StSQL + 'WHERE T_NATUREAUXI ="CLI" AND T_TIERS="' + CodeTiers + '"';
  QQ    := OpenSQL(StSQL, True, -1, '', true);

  if not QQ.Eof then
  begin
    Label4.caption  := 'Tiers : ' + QQ.FindField('T_LIBELLE').AsString;
    Label4.caption  := Label4.caption + ' (' + CodeTiers + ')';
    Auxiliaire      := QQ.FindField('T_AUXILIAIRE').AsString;
    NomTiers        := QQ.FindField('T_LIBELLE').AsString;
  end
  else
  begin
    Label4.caption := 'Tiers Inconnu';
    Auxiliaire      := '';
    NomTiers        := '';
  end;

  Ferme(QQ);

end;

procedure TMonPlanning.ChargeInfoAdresse(CodeAffaire : string; Var Tel : String);
Var StSQL : String;
    QQ    : TQuery;
begin

  //Lecture Adresse Intervention
  StSQL := 'SELECT ADR_ADRESSE1, ADR_CODEPOSTAL, ADR_VILLE, ADR_TELEPHONE FROM ADRESSES WHERE ADR_REFCODE="' + CodeAffaire + '" ';
  StSQL := StSQL + 'AND ADR_TYPEADRESSE="INT"';

  QQ    := OpenSQL(StSQL, True,-1,'',true);

  if not QQ.Eof then
  begin
    Label5.caption := QQ.FindField('ADR_ADRESSE1').AsString;
    Label6.caption := QQ.FindField('ADR_CODEPOSTAL').AsString;
    Label7.caption := QQ.FindField('ADR_VILLE').AsString;
    Label9.caption := 'Tel : ' + QQ.FindField('ADR_TELEPHONE').AsString;
    Tel := QQ.FindField('ADR_TELEPHONE').AsString;
  end
  else
  begin
    Label5.caption := '';
    Label6.caption := '';
    Label7.caption := '';
    Label9.caption := '';
    Tel            := 'Tel :';
  end;

  Ferme(QQ);

end;

Procedure TMonPlanning.ChargeInfoContact(Auxiliaire : string);
Var StSQL : string;
    QQ    : TQuery;
begin

  Label8.caption := 'Contact : ';

  StSql := 'SELECT C_NOM FROM CONTACT ';
  StSql := StSQL + 'WHERE C_AUXILIAIRE ="' + Auxiliaire + '"';
  StSql := StSQL + '  AND C_NUMEROCONTACT =' + IntToStr(Contact);

  QQ := OpenSQL(StSQL, True,-1,'',true);

  if not QQ.Eof then Label8.caption := Label8.caption + QQ.FindField('C_NOM').AsString;

  Ferme(QQ);

end;

Procedure TMonPlanning.ChargeInfoTypeAction(Affaire, CodeTiers : string; NoTache : Integer; var  TypeAction : string);
var StSQl : string;
    QQ    : TQuery;
begin

  TobTaches.ClearDetail;

  if CreerPar = 'TAC' then //chargement de la Tâche correspondante à l'appel (Interventions Préventives)
  Begin
    Stsql := 'SELECT * FROM TACHE WHERE ATA_AFFAIRE="' + AffaireInit + '"';
    StSQL := StSQL + '  AND ATA_TIERS="'               + CodeTiers   + '"';
    StSQL := StSQL + '  AND ATA_NUMEROTACHE='          + IntToStr(NoTache) ;
    //
    QQ := OpenSQL(StSQL, True,-1,'',true);
    if not QQ.eof then
    begin
      TobTaches.selectDB('', QQ);
      TypeAction := TobTaches.GetValue('ATA_BTETAT');
    end;
    Ferme(QQ);
  end
  else
    TypeAction := GetParamSoc('SO_TYPEACTION');

end;

procedure TMonPlanning.ChargeIntervention;
Var NbColFixe: Integer;
begin

  //Initialisation de la grille
  NbColFixe := TobModelePlanning.Detail[PlanningCurrent.tag].GetValue('HPP_NBCOLDIVERS');

  //Chargement de la Tablette des zone des colonnes Fixes
  DrawGrille(NbColFixe);

  //Contrôle si les affectations ne sont pas déjà faites et chargement du tableau
  ChargeEvenementIntervention;

  TobEvenement.PutGridDetail(GrilleSelection, False, False, ColNames);

  GrilleSelection.SetFocus;


end;

procedure TMonPlanning.ChargeEvenementIntervention();
var StSql				: string;
    QQ					: TQuery;
begin

  if CodeDossierRech.text = '' then Exit;

  TobEvenement.ClearDetail;

  StSql := 'SELECT "INTERV" AS ORIGINEITEM, BTEVENPLAN.*, ' +  FindEtReplace(ColNames, ';', ',', true) + ' FROM BTEVENPLAN ';
  StSql := StSql + 'LEFT JOIN RESSOURCE ON BEP_RESSOURCE = ARS_RESSOURCE ';
  StSQL := StSQL + 'LEFT JOIN BTTYPERES ON  BTR_TYPRES=ARS_CHAINEORDO ' ;
  StSql := StSql + 'WHERE BEP_AFFAIRE="' + CodeDossierRech.text + '" ORDER BY BEP_CODEEVENT'; //

  QQ := OpenSQL(StSQL, true,-1,'',true);

  if Not QQ.eof then
  begin
    TobEvenement.LoadDetailDB('BTEVENPLAN', '', '', QQ, False);
    ChargementGrilleWithEvtIntervention(TobEvenement);
    TobEvenement.PutValue('AFFAIRE', TobEvenement.detail[0].GetValue('BEP_AFFAIRE'));
  end;

  Ferme(QQ);

end;

procedure TMonPlanning.ChargementGrilleWithEvtIntervention(TobEvenement : TOB);
var Nomchamp		: string;
    ValeurChamp : String;
    NumAppel    : string;
    Ind					: Integer;
    TobLigne		: Tob;
    TobL  			: Tob;
    Critere			: string;
    DetailGrille: string;
    NbHeure     : Double;
    Prefixe     : string;
begin

  //
  DetailGrille := FindEtReplace(ColNames, ';', ',', true);
  if DetailGrille = '' then DetailGrille := 'ARS_LIBELLE';

  for ind := 0 to TobEvenement.detail.Count - 1 do
  begin
    //
    TobL := TobEvenement.Detail[Ind];
    //
    if Tobl.GetValue('ORIGINEITEM') = 'INTERV' then
      Prefixe := 'BEP_';
    //
    Critere := ColNames;
    if Critere = '' then Critere := 'ARS_LIBELLE';
    NomChamp := (Trim(ReadTokenSt(Critere)));
    //
    while (NomChamp <> '') do
    begin
      ValeurChamp := RechDom(TabletteAssociee(NomChamp),TobL.getvalue(nomchamp), False);
      If ValeurChamp = '' then ValeurChamp := TobL.getvalue(nomchamp);
      TOBL.AddChampSupValeur(NomChamp, ValeurChamp);
      NomChamp := Trim(ReadTokenSt(Critere));
    end;
    //
  	TOBL.AddChampSupValeur('NEW_ONE','-');
    if TobL.GetValue('ORIGINEITEM') = 'INTERV' then
      TOBL.AddChampSupValeur('TYPEACTION','INT')
    else
      TOBL.AddChampSupValeur('TYPEACTION','PCA');

    if TobL.GetValue(Prefixe + 'RESPRINCIPALE') <> 'X' then
    Begin
      Nbheure      := TobL.GetValue(Prefixe + 'DUREE');
      HeureDeb     := TobL.GetValue(Prefixe + 'HEUREDEB');
      HeureFin     := TobL.GetValue(Prefixe + 'HEUREFIN');
      DatePlanDeb  := TobL.GetValue(Prefixe + 'DATEDEB');;
      DatePlanFin  := TobL.GetValue(Prefixe + 'DATEFIN');;
    end;
  end;

end;

Procedure TMonPlanning.AffichageAffectationRessource;
Var LignePlan   : Integer;
    StSQL       : string;
    QQ          : TQuery;
    DateDeb     : TDateTime;
    DateFin     : TDateTime;
    HeureDeb    : TDateTime;
    HeureFin    : TDateTime;
    Critere     : String;
    NomChamp    : string;
    TypeEvent   : String;
begin

  LignePlan   := PlanningCurrent.Row - PlanningCurrent.FixedRows;

  TobItem.InitValeurs(False);

  ChargementTobItem(PlanningCurrent.TobRes.Detail[LignePlan]);
  //
  //ChargementDateItem(PlanningCurrent.TobRes.Detail[LignePlan], nil);
  //
  DateDeb := TobItem.GetDateTime('BPL_DATEDEB');
  DateFin := TobItem.GetDateTime('BPL_DATEFIN');
  HeureDeb:= TobItem.GetDateTime('BPL_HEUREDEB');
  HeureFin:= TobItem.GetDateTime('BPL_HEUREFIN');
  //
  if not ControleRessourceDispo(TobItem, CtrlCal) then Exit;
  //
  TobItem.PutValue('BPL_IDEVENT', '');
  //
  TobItem.PutValue('BPL_AFFAIRE',    TobEvenement.GetValue('AFFAIRE'));
  TobItem.PutValue('BPL_TIERS',      CodeTiers);
  //
  Equipe := TobItem.GetValue('BPL_EQUIPERESS');
  //
  //chargement des zones de la grille
  Critere := ColNames;
  NomChamp := (Trim(ReadTokenSt(Critere)));

  while (NomChamp <> '') do
  begin
    TobItem.AddChampSupValeur(NomChamp, PlanningCurrent.TobRes.Detail[LignePlan].GetString(NomChamp));
    NomChamp := Trim(ReadTokenSt(Critere));
  end;

  //Controle si la ressource n'existe pas déjà dans la liste
  if not ControleAffectation then exit;

  //Controle si ressource appartient à une Equipe
  if Equipe = '' then
     MajEventPlanning
  else
     if not RechEquipe(Equipe) then MajEventPlanning;

End;

procedure TMonPlanning.MajEventPlanning;
var NumEvent: Integer;
begin

  //Chargement de la Tob des Evènements avec la tob de l'Appel
  if not GetNumCompteur('BEP',iDate1900, NumEvent) then
    AfficheErreur('BTPLANNING', '1', 'Affectation d''Appel')
  else
  begin
    TobItem.PutValue('BPL_IDEVENT', IntToStr(NumEvent));
    MajEventBEP;
  end;

  TobEvenement.PutGridDetail(GrilleSelection, False, False, ColNames);

  GrilleSelection.RowHeights[GrilleSelection.Row] := 18;

end;

Procedure TMonPlanning.MajEventBEP;
Var Equipe    : String;
    NumEvt    : String;
    Critere   : String;
    NomChamp  : string;
    TobLigEvt : Tob;
begin

  NumEvt := '';
  Equipe := '';

  TobLigEvt := Tob.create('BTEVENPLAN',   TobEvenement, -1);

  TobLigEvt.PutValue('BEP_CODEEVENT',     IntToStr(TobItem.GetValue('BPL_IDEVENT')));
  TobLigEvt.PutValue('BEP_RESSOURCE',     TobItem.GetString('BPL_RESSOURCE'));

  TobLigEvt.PutValue('BEP_BTETAT',        TobItem.getValue('BPL_BTETAT'));
  TobLigEvt.PutValue('BEP_TIERS',         TobItem.GetValue('BPL_TIERS'));
  TobLigEvt.PutValue('BEP_NUMADRESSE',    TobItem.getvalue('CONTACT'));

  TobLigEvt.PutValue('BEP_EQUIPERESS',    TobItem.GetString('BPL_EQUIPERESS'));

  TobLigEvt.PutValue('BEP_MODIFIABLE',    TobItem.GetValue('MODIFIABLE')); // XP 07.06.2006

  if TobEvenement.Detail.Count = 1 then
  begin
    TobLigEvt.PutValue('BEP_RESPRINCIPALE', 'X');
    NumEvt := TobItem.GetValue('BPL_IDEVENT');
    Equipe := TobItem.GetValue('BPL_EQUIPERESS');
    //TobItem.PutValue('REFEQUIPE', NumEvt + ';' + Equipe);
  end
  else
    TobLigEvt.PutValue('BEP_RESPRINCIPALE', '-');

  if TobItem.GetValue('BPL_EQUIPE') <> '' then
    TobLigEvt.PutValue('BEP_REFEQUIPE', NumEvt + ';' + Equipe)
  else
    TobLigEvt.PutValue('BEP_REFEQUIPE',   '');

  TobEvenement.PutValue('AFFAIRE', CodeDossierRech.text);

  TobLigEvt.PutValue('BEP_AFFAIRE',       TobEvenement.getValue('AFFAIRE'));

  // TempoDate := DateToStr(DatePlanning);
  TobLigEvt.PutValue('BEP_GEREPLAN',      TobItem.GetValue('BPL_GEREPLANNING'));
  TobLigEvt.PutValue('BEP_EQUIPESEP',     '-');
  TobLigEvt.PutValue('BEP_OBLIGATOIRE',   'X');
  TobLigEvt.PutValue('BEP_DUREE',         TobItem.getValue('BPL_DUREE'));
  TobLigEvt.PutValue('BEP_DATEDEB',       TobItem.GetValue('BPL_DATEDEB'));
  TobLigEvt.PutValue('BEP_HEUREDEB',      TobItem.GetValue('BPL_HEUREDEB'));
  TobLigEvt.PutValue('BEP_DATEFIN',       TobItem.GetValue('BPL_DATEFIN'));
  TobLigEvt.PutValue('BEP_HEUREFIN',      TobItem.getValue('BPL_HEUREFIN'));
  TobLigEvt.PutValue('BEP_PERIODEDEBUT',  '');
  TobLigEvt.PutValue('BEP_PERIODEFIN',    '');
  TobLigEvt.PutValue('BEP_HEURETRAV',     'X');
  TobLigEvt.PutValue('BEP_CREAPLANNING',  'X');
  TobLigEvt.PutValue('BEP_BLOCNOTE',      '');
  TobLigEvt.AddChampSupValeur('NEW_ONE',  'X');
  TobLigEvt.AddChampSupValeur('TYPEACTION',TobItem.GetDateTime('BPL_TYPEACTION'));

  //Récupération valeur des Noms de zone
  Critere := ColNames;
  NomChamp := (Trim(ReadTokenSt(Critere)));

  while (NomChamp <> '') do
  begin
    TobLigEvt.AddChampSupValeur(NomChamp, TobItem.GetString(NomChamp));
    NomChamp := Trim(ReadTokenSt(Critere));
  end;

end;      

Function TMonPlanning.ControleAffectation : Boolean;
Begin

  Result := false;

  If TobItem = nil Then Exit;

  if not control_BtEtat then
  Begin
    AfficheErreur('BTPLANNING','12','Erreur d''affectation');
    exit;
  end;

  if not ControleDansLaGrille then
  Begin
    if Not okModif then AfficheErreur('BTPLANNING', '2', 'Affectation d''Appel');
    exit;
  end;

	Result := true;

  //contrôle du surbooking à faire ici...

End;

//Fonction de controle des doublons dans la grille des affectations
function TMonPlanning.ControleDansLaGrille : boolean;
var
  I: Integer;
  NomChamp: string;
  ChampTob: string;
  ChampPlanning: string;
  Critere: string;
  TobLigne: Tob;
begin

  Result := True;

  if TobEvenement = nil then exit;

  if not Assigned(TobEvenement) then exit;

  //Récupération valeur des Noms de zone Pour Controle Doublon
  Critere := ColNames;
  NomChamp := (Trim(ReadTokenSt(Critere)));

  while (NomChamp <> '') do
  begin
    if ChampPlanning = '' then
      ChampPlanning := TobItem.GetString(NomChamp)
    else
      ChampPlanning := ChampPlanning + ';' + TobItem.GetString(NomChamp);
    NomChamp := (Trim(ReadTokenSt(Critere)));
  end;

  //Récupération valeur des noms de la TOB pour controle doublons
  Critere := ColNames;
  NomChamp := (Trim(ReadTokenSt(Critere)));

  if TobEvenement.Detail.Count <> 0 then
  begin
    for I := 0 to TobEvenement.Detail.Count - 1 do
    begin
      TOBligne := TobEvenement.detail[I];
      while (NomChamp <> '') do
      begin
        if ChampTob = '' then
          ChampTob := TobLigne.getvalue(NomChamp)
        else
          ChampTob := ChampTob + ';' + TobLigne.GetValue(NomChamp);
        NomChamp := (Trim(ReadTokenSt(Critere)));
      end;
      if ChampPlanning = ChampTob then
      begin
        result := false;
        break;
      end
      else
      begin
        ChampTob := '';
        Critere := ColNames;
        NomChamp := (Trim(ReadTokenSt(Critere)));
      end;
    end;
  end;

end;

procedure TMonPlanning.ControleHeureDeb(Var HeureDeb : Double);
Var AMDebut   : Double;
    AMFin     : Double;
    PMDebut   : Double;
    PMFin     : Double;
    HdebParam : Double;
    HFinParam : Double;
begin
  //
  AMDebut   := GetParamSocSecur('SO_BTAMDEBUT', GetParamSocSecur('SO_HEUREDEB', TimeToFloat(StrToTime('08:45:00'))));
  AMfin     := GetParamSocSecur('SO_BTAMFIN',   TimeToFloat(StrToTime('12:20:00')));

  PMDebut   := GetParamSocSecur('SO_BTPMDEBUT', GetParamSocSecur('SO_HEUREFIN', TimeToFloat(StrToTime('14:00:00'))));
  PMFin     := GetParamSocSecur('SO_BTPMFIN',   TimeToFloat(StrToTime('17:45:00')));

  HdebParam := GetParamSocSecur('SO_HEUREDEB',  TimeToFloat(StrToTime('08:45:00')));
  HFinParam := GetParamSocSecur('SO_HEUREFIN',  TimeToFloat(StrToTime('17:45:00')));

  if HeureDeb < AMDebut then HeureDeb := AMDebut;

  if (HeureDeb > AMFin) and (HeureDeb < PMDebut) then HeureDeb := AMFin;

  if HeureDeb > PMFin   then HeureDeb := PMFin;

end;

procedure TMonPlanning.ControleHeureFin(Var HeureFin : Double);
Var AMDebut   : Double;
    AMFin     : Double;
    PMDebut   : Double;
    PMFin     : Double;
    HdebParam : Double;
    HFinParam : Double;
begin
  //
  AMDebut   := GetParamSocSecur('SO_BTAMDEBUT', GetParamSocSecur('SO_HEUREDEB', TimeToFloat(StrToTime('08:45:00'))));
  AMfin     := GetParamSocSecur('SO_BTAMFIN',   TimeToFloat(StrToTime('12:20:00')));

  PMDebut   := GetParamSocSecur('SO_BTPMDEBUT', GetParamSocSecur('SO_HEUREFIN', TimeToFloat(StrToTime('14:00:00'))));
  PMFin     := GetParamSocSecur('SO_BTPMFIN',   TimeToFloat(StrToTime('17:45:00')));

  HdebParam := GetParamSocSecur('SO_HEUREDEB',  TimeToFloat(StrToTime('08:45:00')));
  HFinParam := GetParamSocSecur('SO_HEUREFIN',  TimeToFloat(StrToTime('17:45:00')));

  if HeureFin < AMDebut then Heurefin := AMDebut;

  if (HeureFin > AMFin) And (Heurefin < PMDebut) then HeureFin := PMDebut;

  if HeureFin > PmFin then Heurefin := PMFin;

end;

Procedure TMonPlanning.ValidationAffectation;
Var StSQL   : string;
    TOBL    : TOB;
    DateDeb : TDatetime;
    DateFin : TDateTime;
    OrigineItem : string;
    Affaire : String;
    IdAffect: string;
begin

  TTWAppels.Visible := false;

  //Si la tobsuppression.cou nt est supérieure à 1 alors on traite la suppression d'enregistrement...
  if TobSuppression.Detail.count > 0 then
  begin
    if PGIAsk('Confirmez-vous la suppression de la ou des affectations', 'Suppression Affectation')= MRyes then
    begin
      TobSuppression.DeleteDB;
    end;
  end;

  //Envoi d'un mail aux intervenants de l'appel (Attention ça sert à rien d'envoyer un mail à la boîte à outils ou à la camionnette

  //Mise à jour de la fiche appels
  if TOBEvenement.detail.count <= 0 then
  begin
    Affaire := TobEvenement.GetValue('AFFAIRE');
    StSQL := 'UPDATE AFFAIRE SET AFF_RESPONSABLE="", AFF_DATEREPONSE="' + USDATETIME(idate2099) + '", AFF_ETATAFFAIRE="ECO" '
  end
  else
  begin
    //on Recherche le responsable pour chargement des info dans l'affaire
    TOBL := TOBEvenement.FindFirst(['BEP_RESPRINCIPALE'], ['X'], False);
    if TOBL = nil then
      StSQL := 'UPDATE AFFAIRE SET AFF_DATEREPONSE="' + USDATETIME(TOBL.GetDateTime('BEP_DATEDEB')) + '", AFF_ETATAFFAIRE="AFF" '
    else
    begin
      DateDeb := TobL.GetDateTime('BEP_DATEDEB');
      DateFin := TobL.GetDateTime('BEP_DATEDEB');
      Affaire := TobL.GetVALUE('BEP_AFFAIRE');
      StSQL := 'UPDATE AFFAIRE SET AFF_RESPONSABLE="' + TOBL.GetString('BEP_RESSOURCE') + '", AFF_DATEREPONSE="' + USDATETIME(TOBL.GetDateTime('BEP_DATEDEB')) + '", AFF_ETATAFFAIRE="AFF" '
    end;
  end;

  StSQL := StSQL + ' WHERE AFF_AFFAIRE ="' + Affaire + '"';

  If ExecuteSQL(StSQL) = 0 then
    AfficheErreur('BTPLANNING', '16', 'Erreur Validation Affectation')
  else
  begin
    TobEvenement.SetAllModifie(True);
    TobEvenement.InsertOrUpdateDB(true);

    BAffAppel.Down  := True
  end;

  CodeDossierRech.text := '';
  TobEvenement.ClearDetail;
  TobEvenement.PutValue('AFFAIRE', '');

  ChargeItems;

  if (FicAppel) then
  Begin
    BAffAppel.down    := False;
    BAffAppel.Visible := True;
    Exit;
  end;

  AfficheMultiCritere;

end;

procedure TMonPlanning.ModifAffectation;
Var NumAppel 	    : string;
    CodeTiers     : string;
    NumRes	 	    : string;
    DateDeb       : TDateTime;
    DateFin       : TDateTime;
    StWhere       : string;
    NbHeure       : Double;
begin

  if (not ModeSaisie) then exit;

  if not Ismodifiable (TobItem) then  exit;

  Ok_ModifItem := True;

  if Assigned(TobItem) then
  Begin
    NumAppel    := TobItem.GetString('BPL_AFFAIRE');
    CodeTiers   := TobItem.GetString('BPL_TIERS');
    NumRes	    := TobItem.GetString('BPL_RESSOURCE');
    DateDeb     := TobItem.GetValue ('BPL_DATEDEB');
    DateFin     := Tobitem.GetValue ('BPL_DATEFIN');
    Nbheure     := TobItem.GetValue ('BPL_DUREE');
    HeureDeb    := TobItem.GetValue ('BPL_HEUREDEB');
    HeureFin    := TobItem.GetValue ('BPL_HEUREFIN');
    //
  end;

  if NumAppel = '' then exit;

  StWhere := ' WHERE AFF_AFFAIRE = "' + Numappel + '" ';
  StWhere := StWhere + ' AND AFF_TIERS = "' + CodeTiers + '" ';

  ChargementTobappel(StWhere);

end;

procedure TMonPlanning.SupprimeAffectation;
var TobSupLig	: TOB;
    Tobl      : TOB;
    TOBLR     : TOB;
  	NumEquipe	: String;
    I					: Integer;
    Prefixe   : String;
begin

  if TobEvenement.detail.Count = 0 then exit;

  Tobl := TOBEvenement.detail[GrilleSelection.Row - 1];

  CodeDossierRech.Text := '';

  //vérification si la ressource appartient à une équipe
  if TOBL.GetString(Prefixe + 'EQUIPERESS') <> '' then
  Begin
    if SuppressionEquipe(Prefixe, TOBL.GetString(Prefixe + 'EQUIPERESS')) then Exit;
  end;

  //Chargement dans TOB Suppression
  TobSupLig := Tob.Create('Supprime Evenement', TobSuppression, -1);

  //si je suis sur le responsable il faut que je trouve la première ligne
  if Tobl.GetValue('BEP_RESPRINCIPALE') = 'X' then
  begin
    //On cherche la première ligne qui n'aurait pas de responsable pour la passer en responsable de l'appel
    TOBLR := TOBEvenement.Findfirst(['BEP_RESPRINCIPALE'], ['-'], False);
    if TOBLR <> Nil then
    begin
      TOBLR.PutValue('BEP_RESPRINCIPALE', 'X');
      //TobSupLig.Dupliquer(TOBL, False, True);
      Tobl.PutValue('BEP_RESPRINCIPALE', '-');
      TobSupLig.Dupliquer(TOBL, False, True);
      Responsable := '';
    end
    else //Si aucune autre ligne on charge comme étant non responsable...
    begin
      Tobl.PutValue('BEP_RESPRINCIPALE', '-');
      TobSupLig.Dupliquer(TOBL, False, True);
      Responsable := '';
    end;
  end
  else //Cette ligne n'est pas un responsable on la supprime simplement...
    TobSupLig.Dupliquer(TOBL, False, True);

  //Suppression de la ligne dans la Tob Evènement
  FreeAndNil(TOBL);

  //On Supprime toutes les lignes de la grille
  For i := 1 to GrilleSelection.rowcount -1 do GrilleSelection.DeleteRow(I);

  //Rechargement de la grille avec la tob Evènement
  TobEvenement.PutGridDetail(GrilleSelection, False, False, ColNames);

end;

Function TmonPlanning.SuppressionEquipe(Prefixe, Equipe : String) : Boolean;
Var I	        : Integer;
    J         : Integer;
    TobSupLig : Tob;
    TOBL      : TOB;
    TOBR      : TOB;
Begin

  Result := False;

  //Si on répond non on ne traitera que la ressource sélectionnée
  if pgiAsk('Désirez-vous supprimer toute l''équipe', 'Suppression d''Equipe')=MrNo then exit;

  Result := True;

  //On cherche tous les membres d'une même équipe
  TOBL := TobEvenement.FindFirst([Prefixe + 'EQUIPERESS'],[equipe], False);

  while TOBL <> nil do
  begin
    TobSupLig := Tob.Create('Supprime Evenement', TobSuppression, -1);
    TobSupLig.Dupliquer(TOBL, False, True);
    FreeAndNil(TOBL);
    TOBL := TobEvenement.FindNext([Prefixe + 'EQUIPERESS'],[equipe], False);
  end;

  //On Supprime toutes les lignes de la grille
  For i := 1 to GrilleSelection.rowcount -1 do GrilleSelection.DeleteRow(I);

  //Rechargement de la grille avec la tob Evènement
  TobEvenement.PutGridDetail(GrilleSelection, False, False, ColNames);

end;

Procedure TMonPlanning.EditionIntervention(TobLigItem : TOB);
Var StWhere   : String;
    TitreEtat : String;
begin

  if TobLigItem = nil then Exit;

  CodeDossierRech.text := TobLigItem.GetString('BPL_AFFAIRE');

  stWhere := 'AFF_AFFAIRE="' + TobLigItem.GetString('BPL_AFFAIRE') + '"';

  TitreEtat := Copy (TobLigItem.GetString('BPL_AFFAIRE'),2,14);

  LanceEtat('E','INT', GetParamsoc('SO_BTMODBI'),True, false, false,Nil,trim(StWhere),'INTERVENTION '+TitreEtat,False);

end;

Procedure TMonPlanning.CreationEvenement(LignePlan : Integer);
Var NbOption : Integer;
    TobParam : TOb;
Begin

  NbOption := 1;

  TobParam := Tob.Create('ChoixEvenement', nil, -1);
  TobParam.AddChampSupValeur('RETOUR','0');
  TobParam.AddChampSupValeur('VAPPEL', 0);
  TobParam.AddChampSupValeur('VINTERVENTION', 0);
  TobParam.AddChampSupValeur('VACTION_GRC', 0);
  TobParam.AddChampSupValeur('VPARCMAT', 0);
  TobParam.AddChampSupValeur('VCHANTIER', 0);

  //il faut savoir sur quel type de planning on se trouve !!!!
  if ModePlanning = 'PLA' then
  begin
    TobParam.SetInteger('VAPPEL', 1);
    TobParam.SetInteger('VINTERVENTION', 1);
    Inc(NbOption);
    //
    if (TobModelePlanning.Detail[0].GetBoolean('HPP_AFFACTGRC')) then
    begin
      TobParam.SetInteger('VACTION_GRC', 1);
      Inc(NbOption);
    end;
    if (TobModelePlanning.Detail[0].GetBoolean('HPP_AFFEVTMAT')) then
    begin
      TobParam.SetInteger('VPARCMAT', 1);
      Inc(NbOption);
    end;
    if (TobModelePlanning.Detail[0].GetBoolean('HPP_AFFEVTCHA')) then
    Begin
      TobParam.SetInteger('VCHANTIER', 1);
      Inc(NbOption);
    end;
    if NbOption = 1 then CreateEvenementIntervention;
  end
  else if ModePlanning = 'PMA' then
  begin
    TobParam.SetInteger('VPARCMAT', 1);
    if (TobModelePlanning.Detail[0].GetBoolean('HPP_AFFEVTINT')) then
    begin
      TobParam.SetInteger('VAPPEL', 1);
      TobParam.SetInteger('VINTERVENTION', 1);
      Inc(NbOption);
    end;
    if (TobModelePlanning.Detail[0].GetBoolean('HPP_AFFACTGRC')) then
    begin
      TobParam.SetInteger('VACTION_GRC', 1);
      Inc(NbOption);
    end;
    if (TobModelePlanning.Detail[0].GetBoolean('HPP_AFFEVTCHA')) then
    begin
      TobParam.SetInteger('VCHANTIER', 1);
      Inc(NbOption);
    end;
    if NbOption = 1 then CreateEventMateriel;
  end
  else if ModePlanning = 'PCA' then
  begin
    TobParam.SetInteger('VCHANTIER', 1);
    if (TobModelePlanning.Detail[0].GetBoolean('HPP_AFFEVTINT')) then
    begin
      TobParam.SetInteger('VAPPEL', 1);
      TobParam.SetInteger('VINTERVENTION', 1);
      Inc(NbOption);
    end;
    //FV1 - 09/12/2016
    if (TobModelePlanning.Detail[0].GetBoolean('HPP_AFFACTGRC')) then
    begin
      TobParam.SetInteger('VACTION_GRC', 1);
      Inc(NbOption);
    end;
    if (TobModelePlanning.Detail[0].GetBoolean('HPP_AFFEVTMAT')) then
    begin
      TobParam.SetInteger('VPARCMAT', 1);
      Inc(NbOption);
    end;
    if NbOption = 1 then CreateEventPlanCharge;
  end;

  if NBOption > 1 then
  begin
    TheTob := TobParam;
    AglLanceFiche ('BTP','BTCHOIXEVENT','','','');
    TheTob := nil;
    if TOBParam.getValue('RETOUR') <> 0 then
    begin
      //
      if TOBParam.GetValue('RETOUR')      = '1' then  TobItem.PutValue('BPL_ORIGINEITEM', 'INTERV')
      Else if TOBParam.GetValue('RETOUR') = '2' then  TobItem.PutValue('BPL_ORIGINEITEM', 'INTERV')
      Else if TOBParam.GetValue('RETOUR') = '3' then  TobItem.PutValue('BPL_ORIGINEITEM', 'ACT-GRC')
      Else if TOBParam.GetValue('RETOUR') = '4' then  TobItem.PutValue('BPL_ORIGINEITEM', 'PARCMAT')
      Else if TOBParam.GetValue('RETOUR') = '5' then  TobItem.PutValue('BPL_ORIGINEITEM', 'CHANTIER');
      //Pourquoi on fait ça deux fois (????)
      ChargementTobItem(PlanningCurrent.TobRes.Detail[LignePlan]);
      //
      if TOBParam.GetValue('RETOUR')      = '1' then CreateAppel
      Else if TOBParam.GetValue('RETOUR') = '2' then CreateEvenementIntervention
      Else if TOBParam.GetValue('RETOUR') = '3' then CreationActionGRC
      Else if TOBParam.GetValue('RETOUR') = '4' then CreateEventMateriel
      Else if TOBParam.GetValue('RETOUR') = '5' then CreateEventPlanCharge;
    end;
  end;

  TOBParam.free;
  
end;

{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 16/03/2016
Modifié le ... :   /  /
Description .. : Gestion des action possible sur évènement action-GRC
Mots clefs ... :
*****************************************************************}
Procedure TMonPlanning.CreationActionGRC;
Var NumRes    : string;
    Retour    : String;
    TheAction : string;
    TobParam 	: Tob;
begin

  //positionnement Ligne du planning
  TobParam := Tob.Create('AppelCreaAction', nil, -1);

  TobParam.AddChampSupValeur('RAC_INTERVENANT', TobItem.GetString('BPL_RESSOURCE'));
  TobParam.AddChampSupValeur('BEP_EQUIPERESS',  TobItem.GetString('BPL_EQUIPERESS'));
  TobParam.AddChampSupValeur('RAC_DATEACTION',  TobItem.GetValue('BPL_DATEDEB'));
  TobParam.AddChampSupValeur('RAC_HEUREACTION', TobItem.GetValue('BPL_HEUREDEB'));
  TobParam.AddChampSupValeur('RAC_DUREEACTION', TobItem.GetValue('BPL_DUREE'));
  //
  //Voir s'il n'y aurait pas la possibilité de charger une action par défaut (???)
  TobParam.AddChampSupValeur('RAC_TYPEACTION',  '');

  TheTob := TobParam;

  if (not ModeSaisie) then TheAction := 'ACTION=CONSULTATION' else TheAction := 'ACTION=CREATION';

  Numres  := TobItem.GetString('ARS_RESSOURCE');

  Retour  := AGLLanceFiche('RT','RTACTIONS','','','RAC_INTERVENANT=' + Numres + ';' + theAction + ';CREATPLANNING') ;

  TobParam.free;

  //chargement des info pour affichage de l'item
  ChargeItems;

end;

procedure TMonPlanning.ModificationsActionGRC;
var resultat    : string;
    Auxiliaire  : string;
    NumAct      : Integer;
    TheAction   : string;
begin

  Auxiliaire  := TobItem.GetValue('AUXILIAIRE');
  NumAct      := TobItem.GetValue('BPL_IDEVENT');

  if (not ModeSaisie) then TheAction := 'ACTION=CONSULTATION' else TheAction := 'ACTION=MODIFICATION';

  resultat := AGLLanceFiche('RT','RTACTIONS','', Auxiliaire + ';' + IntToStr(NumAct), TheAction + ';MODIFPLANNING') ;

end;

Procedure TMonPlanning.SuppressionActionGRC;
var Auxiliaire  : string;
    StSQL       : String;
		NumAction   : integer;
begin

  //If pgiAsk('Confirmez-vous la suppression de l''action GRC ?')<>MrYes then exit;

  Auxiliaire:= TobItem.GetValue('AUXILIAIRE');
  NumAction := TobItem.GetValue('BPL_IDEVENT');

  StSQL :='DELETE FROM ACTIONS ' + ' WHERE RAC_NUMACTION=' + IntToStr(NumAction) + ' AND RAC_AUXILIAIRE="' + Auxiliaire + '"';

  ExecuteSql (StSQL);

end;


function TMonPlanning.ReductionEtirementActionGRC (TobLigItem : TOB) : boolean;
var DateDeb   : TDateTime;
    DateFin   : TdateTime;
    NbHeure   : integer;
		Auxiliaire: string;
    HeureDeb  : Tdatetime;
    Numaction : integer;
begin

  if not UpdateActionGRC (TobLigItem) then
  begin
  	result := true;
    exit;
  end;

end;

function TMonPlanning.DeplacementActionGRC : boolean;
var CodeRes       : string;
    OldRessource  : string;
  	DateDeb       : TDateTime;
    DateFin       : TdateTime;
    HeureDeb      : TdateTime;
    NbHeure       : double;
    Auxiliaire    : string;
    NumAction     : integer;
begin

  Result := CancelAction ;

  if not IsModifiable (TobItem) then
  begin
    Result := true;
    exit;
  end;

  if not UpdateActionGRC (TobItem) then
  begin
  	result := false;
    exit;
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 16/03/2016
Modifié le ... :   /  /
Description .. : Gestion des actions possibles sur évènement hors appels
Mots clefs ... :
*****************************************************************}
Procedure TMonPlanning.CreateEvenementIntervention;
Var TobParam  : TOB;
    theAction : string;
    Btetat    : String;
    StSQL     : string;
    QQ        : TQuery;
begin

  //Chargement du type d'action par défaut ou du type d'action paramétré
  BtEtat := GetParamSocSecur('SO_TYPEACTION','');

  StSQL := 'SELECT BTA_BTETAT, BTA_LIBELLE FROM BTETAT WHERE BTA_BTETAT = "' + BtEtat + '" AND BTA_TYPEACTION="' + TobItem.Getvalue('BPL_TYPEACTION') + '"';
  QQ := OpenSQL(StSQL, false, 1, '', true);
  if not QQ.Eof then
  begin
    TobItem.PutValue('BPL_BTETAT',    QQ.FindField('BTA_BTETAT').AsString);
    TobItem.PutValue('BPL_LIBACTION', QQ.FindField('BTA_LIBELLE').AsString);
  end
  else
  begin
    Ferme(QQ);
    StSQL := 'SELECT BTA_BTETAT, BTA_LIBELLE FROM BTETAT WHERE BTA_DEFAUT="X" AND BTA_TYPEACTION="INT"';
    QQ := OpenSQL(StSQL, false, 1, '', true);
    if not QQ.Eof then
    begin
      TobItem.PutValue('BPL_BTETAT',    QQ.FindField('BTA_BTETAT').AsString);
      TobItem.PutValue('BPL_LIBACTION', QQ.FindField('BTA_LIBELLE').AsString);
    end
    else
    begin
      TobItem.PutValue('BPL_BTETAT', '');
      TobItem.PutValue('BPL_LIBACTION','');
    end;
  end;
  Ferme(QQ);

  TobItem.AddChampSupValeur('RETOUR',        -1);

  TheTob := TobItem;

  if (not ModeSaisie) then TheAction := 'ACTION=CONSULTATION' else TheAction := 'ACTION=CREATION';

  AglLanceFiche ('BTP','BTEVENPLAN','','',theAction+';PLANNING');

  //chargement des info pour affichage de l'item
  if TobItem.GetValue('RETOUR') = 1 then
  begin
    //IndexPage           := PageOnglet.ActivePageIndex;
    EdtDosEnCours.text  := TobItem.GetValue('BPL_AFFAIRE');
  end;

  FreeAndNil(TheTob);

  //chargement des info pour affichage de l'item
  ChargeItems;

end;

Procedure TMonPlanning.ModificationEvenementIntervention;
Var NumEvenPlan : Integer;
    Argument    : string;
begin

  if Assigned(TobItem) then NumEvenPlan := TobItem.GetValue('BPL_IDEVENT');

  if NumEvenPlan <> 0 then
  Begin
    TheTob := TobItem;
    Argument := 'IDEVENEMENT=' + IntToStr(NumEvenPlan) + ';ACTION=MODIFICATION';
    AGLLanceFiche('BTP','BTEVENPLAN','','',Argument+';PLANNING');
    FreeAndNil(TheTOB);                                        
  end;

  ChargeItems;

end;

Procedure TMonPlanning.SuppressionEvenementIntervention;
var Rep       : String;
    Equipe    : string;
    NumEvent  : String;
    Table     : String;
begin

  Equipe    := TobItem.GetValue('BPL_EQUIPERESS');
  NumEvent  := TobItem.GetValue('BPL_IDEVENT');

  //Controle si ressource appartient à une Equipe
  if Equipe <> '' then
  begin
    if PGIAsk(TraduireMemoire('La Ressource appartient à l''Equipe ' + Equipe + '.' + Chr(10) + 'Voulez-vous gérer l''equipe ?'), rep) = mryes then
      SuppressionEvtIntervEquipe(NumEvent)
    else
      SuppressionEvtInterv(NumEvent);
  end
  Else
    SuppressionEvtInterv(NumEvent);

end;

Procedure TMonplanning.SuppressionEvtInterv(Numevent : String);
Var StsQL : string;
begin

  StSQL := 'DELETE FROM BTEVENPLAN WHERE BEP_CODEEVENT = "' + Numevent + '"';
  if ExecuteSql (StSQL) <= 0 then Exit;

end;

Procedure TMonPlanning.SuppressionEvtIntervEquipe(NumEvent : String);
Var StSQL     : String;
    QQ        : Tquery;
    RefEquipe : String;
    TobLigEvt : Tob;
    TobLEquipe: TOB;
    I         : Integer;
begin

  StSQL := 'SELECT BEP_REFEQUIPE FROM BTEVENPLAN ';
  StSQL := StSQL + 'WHERE BEP_CODEEVENT ="' + NumEvent + '"';
  QQ := OpenSQL(StSql, True,-1,'',true);

  if not QQ.eof then
  begin
    RefEquipe := QQ.FindField('BEP_REFEQUIPE').asString;
    Ferme(QQ);
    StSQL := 'DELETE FROM BTEVENPLAN WHERE BEP_REFEQUIPE="' + RefEquipe + '"';
    ExecuteSql (StSQL);
  end
  else Ferme(QQ);

end;

procedure TMonPlanning.SetInfosAffectation(TobEvt : tob);
Var NumEvenPlan : Integer;
    Argument    : string;
Begin

  //Faire en sorte d'afficher la nouvelle fenêtre !!!
  if Assigned(TobEvt) then NumEvenPlan := TobEvt.GetValue('BEP_CODEEVENT');

  if NumEvenPlan <> 0 then
  Begin
    Argument := 'IDEVENEMENT=' + IntToStr(NumEvenPlan) + ';ACTION=MODIFICATION';
    AGLLanceFiche('BTP','BTEVENPLAN','','',Argument);
  end;

end;

function TMonPlanning.RechEquipeAction : Boolean;
var StSql     : string;
    ZoneLue   : string;
    Rep       : string;
    NumEvt    : String;
    I         : Integer;
    QQ        : TQuery;
    TobEquipe : Tob;
    TobLEquipe: TOB;
begin

  Result := False;

  if TobItem = nil then Exit;

  StSQL := 'SELECT * FROM RESSOURCE ';
  StSQL := StSQL + 'WHERE ARS_EQUIPERESS="' + Equipe + '" ';
  StSQL := StSQL + 'ORDER BY ARS_TYPERESSOURCE';

  QQ := OpenSQL(StSql, True,-1,'',true);
  if QQ.eof then
  begin
    Ferme(QQ);
    Exit;
  end;

  if PGIAsk(TraduireMemoire('La Ressource appartient à l''Equipe ' + Equipe + '.' + Chr(10) + 'Voulez-vous gérer l''equipe ?'), rep) = mrno then
  begin
    Ferme(QQ);
    exit;
  end;

  Result := true;

  //Création de la Tob equipe uniquement si Equipe existe et
  //si réponse oui à la question
  TobEquipe := Tob.Create('Equipe', nil, -1);
  TobEquipe.LoadDetailDB('LEQUIPE', '', '', QQ, False);

  Ferme(QQ);

  //Vérification si les info de la tob n'existe pas déjà dans le grille
  for i := 0 to TobEquipe.detail.count - 1 do
  begin
    TobLEquipe := TobEquipe.Detail[I];
    CodeRes := TobLEquipe.getvalue('ARS_RESSOURCE');
    //chargement des nouvelles zones avec les zone items....
    TobLEquipe.PutValue('NUMEVENT', 0);
    TobLEquipe.PutValue('ORIGINEITEM',  'INTERV');
    TobLEquipe.PutValue('TYPEACTION',   'INT');
    TobLEquipe.PutValue('BTETAT',   TobItem.GetValue('BTETAT'));
    TobLEquipe.PutValue('DUREE',    TobItem.GetValue('DUREE'));
    TobLEquipe.PutValue('DATEDEB',  TobItem.GetValue('DATEDEB'));
    TobLEquipe.PutValue('DATEFIN',  TobItem.GetValue('DATEFIN'));
    TobLEquipe.PutValue('HEUREDEB', TobItem.GetValue('HEUREDEB'));
    TobLEquipe.PutValue('HEUREFIN', TobItem.GetValue('HEUREFIN'));
    TobLEquipe.PutValue('REFEQUIPE',TobEquipe.Detail[0].GetValue('REFEQUIPE'));
    CreationEvtAction(TobLEquipe); //==>il faut le TobLigItem...
  end;

  FreeAndNil(TobEquipe);

end;

Procedure TMonPlanning.CreationEvtAction(TobLigItem : TOB);
var NumEvent: Integer;
    Equipe  : String;
begin

  if not GetNumCompteur('BEP',iDate1900, NumEvent) then
    AfficheErreur('BTPLANNING', '1', 'Affectation d''Appel')
  else
  begin
    TobLigItem.PutValue('NUMEVENT',   NumEvent);
    //
    MajEventPlanning;

    TobEvenement.SetAllModifie(True);
    TobEvenement.InsertOrUpdateDB(true);
  end;

end;

Procedure TMonPlanning.ModificationAction;
Var Rep       : string;
begin

  if TobItem = nil then Exit;

  Equipe := TobItem.GetValue('BPL_EQUIPERESS');

  //Controle si ressource appartient à une Equipe
  if Equipe <> '' then
  begin
    if PGIAsk(TraduireMemoire('La Ressource appartient à l''Equipe ' + Equipe + '.' + Chr(10) + 'Voulez-vous gérer l''equipe ?'), rep) = mryes then
      MajActionEquipe
    else
      ModificationEvtaction;
  end
  Else ModificationEvtAction;

  TobEvenement.ClearDetail;

  ChargeItems;

end;

Procedure TMonPlanning.MajActionEquipe;
Var StSQL     : String;
    QQ        : Tquery;
    RefEquipe : String;
    TobLigEvt : Tob;
    TobLEquipe: TOB;
    I         : Integer;
begin

  StSQL := 'SELECT BEP_REFEQUIPE FROM BTEVENPLAN ';
  StSQL := StSQL + 'WHERE BEP_CODEEVENT ="' + IntToStr(TobItem.GetValue('NUMEVENT')) + '"';
  QQ := OpenSQL(StSql, True,-1,'',true);

  if QQ.eof then
  begin
    Ferme(QQ);
     Exit;
  end;

  RefEquipe := QQ.FindField('BEP_REFEQUIPE').asString;

  Ferme(QQ);

  StSQL := 'SELECT * FROM BTEVENPLAN ';
  StSQL := StSQL + 'WHERE BEP_REFEQUIPE ="' + RefEquipe + '"';
  QQ := OpenSQL(StSql, True,-1,'',true);
  if not QQ.eof then
  begin
    //Création de la Tob equipe uniquement si Equipe existe et
    //si réponse oui à la question
    TobLigEvt := Tob.create('LIGEVENT', TobEvenement, -1);
    TobLigEvt.LoadDetailDB('BTEVENPLAN', '', '', QQ, False);
    for i := 0 to TobLigEvt.detail.count - 1 do
    begin
      TobLEquipe := TobLigEvt.Detail[I];
      //chargement des nouvelles zones avec les zone items....
      TobLEquipe.PutValue('BEP_BTETAT',        TobItem.getValue('BTETAT'));
      TobLEquipe.PutValue('BEP_DUREE',         TobItem.getValue('DUREE'));
      TobLEquipe.PutValue('BEP_DATEDEB',       TobItem.GetDateTime('DATEDEB'));
      TobLEquipe.PutValue('BEP_HEUREDEB',      TobItem.GetDateTime('HEUREDEB'));
      TobLEquipe.PutValue('BEP_DATEFIN',       TobItem.GetDateTime('DATEFIN'));
      TobLEquipe.PutValue('BEP_HEUREFIN',      TobItem.GetDateTime('HEUREFIN'));
      TobLEquipe.PutValue('BEP_GEREPLAN',      TobItem.GetValue('GEREPLANNING'));
      //On envoie un mail à chaque personne de l'équipe (???)
      if OkEnvoiMail then	EnvoieMail ('Modification Action N°' + TobLEquipe.GetValue('BPL_CODEEVENT'));
    end;
    TobLigEvt.SetAllModifie(True);
    TobLigEvt.UpdateDB;
    Ferme(QQ);
  end;

end;

Procedure TMonPlanning.ModificationEvtAction;
var StSQL     : string;
    QQ        : TQuery;
    TobLigEvt : Tob;
begin

  //chargement de la tob evenement pour MAJ
  StSQL := 'SELECT * FROM BTEVENPLAN WHERE BEP_CODEEVENT="' + TobItem.GetValue('BPL_IDEVENT') + '"';

  QQ    := OpenSQL(StSQL, False, -1,'',True);

  TobLigEvt := Tob.create('BTEVENPLAN', TobEvenement, -1);
  if Not QQ.EOf then TobLigEvt.SelectDB('', QQ, True);
  Ferme(QQ);
  //
  TobLigEvt.PutValue('BEP_RESSOURCE',     TobItem.GetString('BPL_RESSOURCE'));

  TobLigEvt.PutValue('BEP_BTETAT',        TobItem.getValue('BPL_BTETAT'));
  TobLigEvt.PutValue('BEP_DUREE',         TobItem.getValue('BPL_DUREE'));
  TobLigEvt.PutValue('BEP_DATEDEB',       TobItem.GetDateTime('BPL_DATEDEB'));
  TobLigEvt.PutValue('BEP_HEUREDEB',      TobItem.GetDateTime('BPL_HEUREDEB'));
  TobLigEvt.PutValue('BEP_DATEFIN',       TobItem.GetDateTime('BPL_DATEFIN'));
  TobLigEvt.PutValue('BEP_HEUREFIN',      TobItem.GetDateTime('BPL_HEUREFIN'));
  TobLigEvt.PutValue('BEP_GEREPLAN',      TobItem.GetValue('BPL_GEREPLANNING'));

  TobLigEvt.SetAllModifie(True);
  TobLigEvt.UpdateDB;

  if OkEnvoiMail then	EnvoieMail ('Modification Action N°' + TobItem.GetValue('BPL_IDEVENT'));

end;

{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 22/03/2016
Modifié le ... :   /  /
Description .. : Action possible sur une item (Popupmenu)
Mots clefs ... :
*****************************************************************}
// Extension d'un item
procedure TMonPlanning.Extension(Sender: TObject; Item: Tob; var Cancel: boolean);
Var Origine : String;
    Datefin : TDateTime;
Begin

  if (not ModeSaisie) or (item = nil) then
  BEGIN
    cancel        := true;
    CancelAction  := True;
    exit;
  END;

  Cancel := CancelAction ;
  if CancelAction       then exit;

  TobItem := Tob.create('ITEM',nil,-1);
  TobItem.Dupliquer(Item, False, True);

  Origine := TobItem.getString('BPL_ORIGINEITEM');

  if Origine = 'PARCMAT' then
    cancel := ReductionEtirementParcMateriel (TobItem, Cadencement)
  else if Origine = 'ACT-GRC' then
    Cancel := ReductionEtirementActionGRC (TobItem)
  else if Origine = 'INTERV' then
    //FV1 : 12/11/2015 - FS#1691 - SERVAIS : Pb en saisie de planning si déplacement d'un appel.
    Cancel := not MajEvtAffectation  ('E')
  else if Origine = 'CHANTIER' then
    cancel := ReductionEtirementPlanCharge (TobItem, Cadencement);

  ChargeItems;

end;

//Réduction d'une item
procedure TMonPlanning.Reduction(Sender: TObject; Item: Tob; var Cancel: boolean);
Var Origine : String;
Begin

	if (not ModeSaisie) or (item = nil) then
  BEGIN
    Cancel        := True;
    CancelAction  := True;
    exit;
  END;

  Cancel := CancelAction ;
  if CancelAction       then exit;

  TobItem := Tob.create('ITEM',nil,-1);
  TobItem.Dupliquer(Item, False, True);

  Origine := TobItem.getString('BPL_ORIGINEITEM');

  if Origine = 'PARCMAT' then
    Cancel := ReductionEtirementParcMateriel(TobItem, Cadencement)
  else if Origine = 'ACT-GRC' then
    Cancel := ReductionEtirementActionGRC (TobItem)
  else if Origine = 'INTERV' then
    Cancel := not MajEvtAffectation  ('R')
  else if Origine = 'CHANTIER' then
    Cancel := ReductionEtirementPlanCharge(TobItem, Cadencement);

  ChargeItems;

end;

//Deplacement d'un item
procedure TMonPlanning.Deplacement(Sender: TObject; Item: TOB; var Cancel: boolean);
Var LignePlan : Integer;
    Origine   : String;
Begin

	if (not ModeSaisie) or (item = nil) then
  BEGIN
    Cancel        := True;
    CancelAction  := True;
    exit;
  END;

  TobItem.InitValeurs(False);
  TobItem.Dupliquer(Item, False, True);

  Origine := TobItem.getString('BPL_ORIGINEITEM');

  if Origine = 'PARCMAT' then
    cancel := DeplacementParcMateriel (TobItem,TobJFerie, PlanningCurrent.GetCurItem.GetValue('BPL_MATERIEL'))
  else if Origine = 'ACT-GRC' then
  	cancel := not DeplacementActionGRC
  else if Origine = 'INTERV' then
    //FV1 : 12/11/2015 - FS#1691 - SERVAIS : Pb en saisie de planning si déplacement d'un appel.
    Cancel := not MajEvtAffectation ('D')
 else if  Origine = 'CHANTIER' then
    cancel := DeplacementPlanCharge (TobItem, PlanningCurrent.GetCurItem.GetValue('BPL_RESSOURCE'), Cadencement);

 ChargeItems;

  //if FicAppel then ChargeIntervention;

end;

Function TMonPlanning.DuplicationAction : Boolean;
Var NumEvenement : String;
    CodeRes      : String;
    OldRessource : string;
    NumAppel     : String;
    Equipe       : String;
    StSql        : String;
    Rep          : String;
    NbHeure	     : Double;
    DateDeb      : Tdatetime;
    Datefin      : Tdatetime;
    NumEvent     : Integer;
    I            : Integer;
    QQ           : TQuery;
    TobRessource : Tob;
    TobItemDe    : Tob;
    TobItemA     : Tob;
    Ok_Equipe    : Boolean;
    An, Mois, Jour  : Word;
    H, M, S, Msec   : Word;
    Delta        : Integer;
begin

  //Controle de cohérence de la duplication
  DateFin         := StrToDateTime(TobItem.GetValue('BPL_DATEAFFFIN'));
  DecodeDateTime(DateFin, an, Mois, Jour, H, M, S, Msec);
  HeureFin        := EncodeTime(H, M, 0, 0);

  //Avant de faire quoi que ce soit il faut ajouter +1 à l'heure de fin
  if Cadencement = '001' then // gestion par 1/4 Heure
  begin
    HeureFin := IncMinute (HeureFin, 15);
  end else if Cadencement = '002' then // gestion par 1/2 Heure
  begin
    HeureFin := IncMinute (HeureFin,30);
  end else if Cadencement = '003' then // gestion par heure
  begin
    HeureFin := IncHour (HeureFin, 1);
  end else if Cadencement = '004' then // gestion par 1/2 Journée
  Begin
    if heureFin > GetFinMatinee then
    begin
      Delta :=  MinutesBetween(heureFin,GetFinApresMidi);
    end else
    begin
      Delta := MinutesBetWeen(HeureFin,GetFinMatinee);
    end;
    HeureFin := incminute (HeureFin, Delta);
  end else if Cadencement = '005' then // gestion par Journée
  begin
    // pas de decalage dans ce cas la
  end else if Cadencement = '008' then // gestion par Période
  begin
    {TobTmp.PutValue('BEP_DATEFIN', TobTmp.GetValue('BEP_DATEFIN'));}
  end;

  DateFin         := Trunc(DateFin) + HeureFin;

  //Chargement des nouvelles valeur debut et fin
  DateDeb         := StrToDateTime(TobItem.GetValue('BPL_DATEAFFDEB'));
  DecodeDateTime(DateDeb, an, Mois, Jour, H, M, S, Msec);
  HeureDeb        := EncodeTime(H, M, 0, 0);

  TobItem.PutValue('BPL_DATEDEB', DateTimeToStr(DateDeb));
  TobItem.PutValue('BPL_DATEFIN', DateTimeToStr(DateFin));

  TobItem.PutValue('BPL_HEUREDEB', TimeToStr(HeureDeb));
  TobItem.PutValue('BPL_HEUREFIN', TimeToStr(HeureFin));
  //
  NbHeure := CalculDureeEvenement(DateDeb, DateFin);
  TobItem.PutValue('BPL_DUREE', NbHeure);
  //
  //
  TypeAction  := TobItem.GetString('BPL_BTETAT');
  NumEvenement:= TobItem.GetString('BPL_IDEVENT');
  CodeRes		  := TobItem.GetValue('BPL_RESSOURCE');
  OldRessource:= PlanningCurrent.GetCurItem.GetValue('BPL_RESSOURCE');
  NumAppel    := TobItem.GetString('BPL_AFFAIRE');
  //
  Ok_equipe   := false;

  //if not ControleRessourceArrivee(CodeRes, DateDeb, DateFin, HeureDeb, HeureFin) then
  if not ControleRessourceDispo(TobItem, CtrlCal) then
  begin
    AfficheErreur('BTPLANNING', '4', 'Duplication d''Affectation');
    result := True;
    exit;
  end;

  //if TobItem.GetString('BPL_EQUIPRERESS') <> '' then
  //begin
  //  Equipe := QQ.FindField('ARS_EQUIPERESS').AsString;
  //  If PgiAsk('Voulez-vous appliquer la duplication à l''ensemble de l''équipe '+ Equipe + ' ?', rep)=MrYes then Ok_Equipe := true;
  //end;

  TobItemA := Tob.Create('BTEVENPLAN', nil, -1);

  if not ok_equipe then
  Begin
    //Chargement de la Tob des Evènements avec la tob de l'Appel
    if not GetNumCompteur('BEP',iDate1900, NumEvent) then
    begin
      AfficheErreur('BTPLANNING', '1', 'Affectation d''Appel');
      exit;
    end;
    TobItemA.PutValue('BEP_CODEEVENT',     IntToStr(NumEvent));

    TobItemA.PutValue('BEP_RESSOURCE',     CodeRes);

    TobItemA.PutValue('BEP_BTETAT',        TobItem.getValue('BPL_BTETAT'));
    TobItemA.PutValue('BEP_AFFAIRE',       TobItem.getValue('BPL_AFFAIRE'));
    TobItemA.PutValue('BEP_TIERS',         TobItem.getValue('BPL_TIERS'));
    TobItemA.PutValue('BEP_NUMEROADRESSE', TobItem.getValue('BPL_NUMEROADRESSE'));
    TobItemA.PutValue('BEP_EQUIPERESS',    TobItem.getValue('BPL_EQUIPERESS'));
    TobItemA.PutValue('BEP_MODIFIABLE',    TobItem.getValue('MODIFIABLE'));
    TobItemA.PutValue('BEP_RESPRINCIPALE', TobItem.getValue('-'));
    TobItemA.PutValue('BEP_EQUIPESEP',     TobItem.getValue('-'));
    if NumAppel <> '' then
      TobItemA.PutValue('BEP_OBLIGATOIRE', 'X')
    else
      TobItemA.PutValue('BEP_OBLIGATOIRE', '-');
    TobItemA.PutValue('BEP_DUREE',         TobItem.getValue('BPL_DUREE'));
    TobItemA.PutValue('BEP_DATEDEB',       TobItem.GetDateTime('BPL_DATEDEB'));
    TobItemA.PutValue('BEP_HEUREDEB',      TobItem.GetDateTime('BPL_HEUREDEB'));
    TobItemA.PutValue('BEP_DATEFIN',       TobItem.GetDateTime('BPL_DATEFIN'));
    TobItemA.PutValue('BEP_HEUREFIN',      TobItem.GetDateTime('BPL_HEUREFIN'));
    TobItemA.PutValue('BEP_PERIODEDEBUT',  TobItem.getValue(''));
    TobItemA.PutValue('BEP_PERIODEFIN',    TobItem.getValue(''));
    TobItemA.PutValue('BEP_GEREPLAN',      TobItem.GetValue('BPL_GEREPLANNING'));
    TobItemA.PutValue('BEP_HEURETRAV',     TobItem.getValue('BPL_EQUIPERESS'));
    TobItemA.PutValue('BEP_CREAPLANNING',  'X');
    TobItemA.PutValue('BEP_BLOCNOTE',      TobItem.getValue('BPL_DESCRIPTIF'));
    TobItemA.PutValue('BEP_REFEQUIPE',     '');
    //
    TobItemA.SetAllModifie(True);
    TobItemA.InsertOrUpdateDB(False);
    FreeAndNil(TobItemA);
    //
  end;

end;

procedure TMonPlanning.CopyItem(Sender: TObject; Item: TOB; var Cancel: boolean);
var Origine : string;
begin

  if (not ModeSaisie) or (item = nil) then
  BEGIN
    Cancel        := True;
    CancelAction  := True;
    exit;
  END;

  TobItem.InitValeurs(False);
  TobItem.Dupliquer(Item, False, True);

  Origine := TobItem.getString('BPL_ORIGINEITEM');

  if Origine = 'PARCMAT' then
    DuplicationParcMateriel (TobItem, TobJFerie, PlanningCurrent.GetCurItem.GetValue('BPL_MATERIEL'))
  else if Origine = 'ACT-GRC' then
    Exit
  else if Origine = 'INTERV' then
    DuplicationAction
  else if Origine = 'CHANTIER' then
    DuplicationPlanCharge(TobItem, PlanningCurrent.GetCurItem.GetValue('BPL_RESSOURCE'), Cadencement);

  //ChargeItems;

  if FicAppel then ChargeIntervention;

end;

Function TMonPlanning.ControleRessourceArrivee(CodeRes : String; DateDeb, DateFin, HeureDeb, Heurefin : TDateTime) : boolean;
var StSQL : string;
    QQ    : TQuery;
begin

  result := True;

  //Controle si pas déjà une affectation pour la ressource arrivée sur la date et heure de duplication
  StSql := 'SELECT BEP_AFFAIRE FROM BTEVENPLAN WHERE BEP_RESSOURCE ="' + CodeRes + '" ';
  StSQL := StSql + 'AND BEP_DATEDEB >="' + USDATETIME(DateDeb) + '" ';
  StSQL := StSQL + 'AND BEP_DATEFIN <="' + USDATETIME(DateFin) + '" ';

  QQ := OpenSQL(StSQL, true,-1,'',true);
  if Not QQ.Eof then Result := False;

  Ferme(QQ);

end;

{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 09/05/2016
Modifié le ... :   /  /    
Description .. : Procedure de gestion des évènements sur un appel
Mots clefs ... :
*****************************************************************}
procedure TMonPlanning.CreateAppel;
Var NumRes	 	    : String;
    Equipe        : String;
    TobParam 	    : Tob;
    LignePlan	    : Integer;
    Sender		    : TObject;
    NumEvent	    : string;
    //
    TheAction     : string;
    DebutDemande  : TDateTime;
    FinDemande    : TDateTime;
    retour        : TDateTime;
    TitreEtat     : string;
    BtEtat        : string;
    StSQL         : string;
    QQ            : TQuery;
begin


  //Chargement du type d'action par défaut ou du type d'action paramétré
  //
  BtEtat := GetParamSocSecur('SO_TYPEACTION','');

  StSQL := 'SELECT BTA_BTETAT, BTA_LIBELLE FROM BTETAT WHERE BTA_BTETAT = "' + BtEtat + '" AND BTA_TYPEACTION="' + TobItem.Getvalue('BPL_TYPEACTION') + '"';
  QQ := OpenSQL(StSQL, false, 1, '', true);
  if not QQ.Eof then
  begin
    TobItem.PutValue('BPL_BTETAT',    QQ.FindField('BTA_BTETAT').AsString);
    TobItem.PutValue('BPL_LIBACTION', QQ.FindField('BTA_LIBELLE').AsString);
  end
  else
  begin
    Ferme(QQ);
    StSQL := 'SELECT BTA_BTETAT, BTA_LIBELLE FROM BTETAT WHERE BTA_DEFAUT="X" AND BTA_TYPEACTION="INT"';
    QQ := OpenSQL(StSQL, false, 1, '', true);
    if not QQ.Eof then
    begin
      TobItem.PutValue('BPL_BTETAT',    QQ.FindField('BTA_BTETAT').AsString);
      TobItem.PutValue('BPL_LIBACTION', QQ.FindField('BTA_LIBELLE').AsString);
    end
    else
    begin
      AfficheErreur('BTPLANNING','12','Erreur d''affectation');
      ChargeItems;
      Exit;
    end;
  end;
  Ferme(QQ);

  //positionnement Ligne du planning
  TobParam := Tob.Create('AppelCreaAppel', nil, -1);

  TobParam.AddChampSupValeur('BEP_AFFAIRE',     '');
  TobParam.AddChampSupValeur('RESSOURCE',       TobItem.GetString('BPL_RESSOURCE'));
  TobParam.AddChampSupValeur('BEP_EQUIPERESS',  TobItem.GetString('BPL_EQUIPERESS'));
  TobParam.AddChampSupValeur('DATEDEB',         TobItem.GetValue('BPL_DATEDEB'));
  TobParam.AddChampSupValeur('DATEFIN',         TobItem.GetValue('BPL_DATEFIN'));
  TobParam.AddChampSupValeur('HEUREDEB',        TobItem.GetValue('BPL_HEUREDEB'));
  TobParam.AddChampSupValeur('HEUREFIN',        TobItem.GetValue('BPL_HEUREFIN'));
  TobParam.AddChampSupValeur('DUREE',           TobItem.GetValue('BPL_DUREE'));
  TobParam.AddChampSupValeur('TYPEACTION',      TobItem.GetValue('BPL_BTETAT'));
  TobParam.AddChampSupValeur('RETOUR',          -1);

  TheTob := TobParam;

  if (not ModeSaisie) then TheAction := 'ACTION=CONSULTATION' else TheAction := 'ACTION=CREATION';

  AglLanceFiche ('BTP','BTAPPELINT','','',theAction+';PLANNING');

  TheTob := nil;

  TobParam.free;

  //chargement des info pour affichage de l'item
  ChargeItems;

end;

Procedure TMonPlanning.VisualisationAffectation;
begin

  AGLLanceFiche ('BTP','BTAFFAPPEL','','', TobItem.GetString('BPL_AFFAIRE') + ';' + TobItem.GetString('BPL_RESSOURCE'));

  CodeDossierRech.text := TobItem.GetString('BPL_RESSOURCE');

end;

Procedure TMonPlanning.SuppressionEvenementAppel;
Var DFin			 	: TDateTime;
	  DDebut      : TDateTIme;
    DateDeb     : TDateTime;
    DateFin     : TDateTime;
    SDateDeb    : TDateTime;
    SDateFin    : TDateTime;
    SHeureDeb   : TDateTime;
    SHeurefin   : TDateTime;
    OldRessource: String;
	  NumEvenement: Integer;
    NumAppel 	 	: String;
    Rep			 	  : String;
    Equipe      : String;
    StSql,SQLB  : String;
    TobParam    : Tob;
    QQ			 	  : TQuery;
    sender		 	: Tobject;
    NbHeure     : Integer;
    newDate     : TdateTime;
begin

	if not IsModifiable (TobItem) then exit;

  CodeRes		      := TobItem.GetValue('BPL_RESSOURCE');
  OldRessource    := PlanningCurrent.GetCurItem.GetValue('BPL_RESSOURCE');
  //
  NumAppel        := TobItem.GetString('BPL_AFFAIRE');

  CodeDossierRech.text := NumAppel;

  //control si l'appel n'est pas terminé sinon on fait rien
  if NumAppel <> '' then
  Begin
    if (TobItem.GetValue('BPL_ETATAFFAIRE') = 'TER') OR (TobItem.GetValue('BPL_ETATAFFAIRE') = 'REA') then
    Begin
      AfficheErreur('BTPLANNING', '14', 'Modification d''affectation');
      exit;
    end;
  end;

  NumEvenement    := TobItem.GetValue('BPL_IDEVENT');
  Equipe          := TobItem.GetString('BPL_EQUIPERESS');

  //Traitement de la modification ou de la suppression d'un évènement libre
  if NumAppel = '' then
  Begin
    If pgiAsk('Confirmez-vous la suppression de l''évènement ?', rep)=MrYes then
    begin
      SuppressionItemPlanning('WHERE BEP_CODEEVENT ="' + IntToStr(NumEvenement) + '"');
    end;
  end;

  //gestion de la suppression ou de la modification d'un évènement lié à un appel
  TobParam := Tob.Create('AppelActionEvenement', nil, -1);
  TobParam.AddChampSupValeur('RETOUR', 3);
  TobParam.AddChampSupValeur('ACTION', 'S');
  TobParam.AddChampSupValeur('EQUIPE', '');

  TRY
    //vérification si un ou plusieur évènement pour cet appel
    if numAppel <> '' then
    begin
      StSql := 'SELECT BEP_CODEEVENT FROM BTEVENPLAN WHERE BEP_AFFAIRE = "' + NumAppel + '"';
      QQ := OpenSQL(StSQL, True,-1,'',false);
      if QQ.eof then Exit;
      If QQ.RecordCount > 1 then
      Begin
        TobParam.SetInteger('RETOUR', -1);
        TobParam.SetString('ACTION', 'S');
        TobParam.SetString('EQUIPE', Equipe);
        TheTob := TobParam;
        AglLanceFiche ('BTP','BTSUPPEVENT','','','');
        TheTob := nil;
      end
      Else
      Begin
        TobParam.SetInteger('RETOUR', 1);
        TobParam.SetString('ACTION', 'S');
        TobParam.SetString('EQUIPE', Equipe);
      end;
    end;
    Ferme(QQ);

    //Traitement de suppression (1 - ensemble de l'affectation, 2 - Equipe, 3 - Item Sélectionnées)
    If TobParam.GetValue('RETOUR') = 1 then
    begin
      SqlB := 'WHERE BEP_AFFAIRE ="' + NumAppel + '"';
      SuppressionItemPlanning(SqlB);
    end Else If TobParam.GetValue('RETOUR') = 2 Then
    begin
      SqlB := 'WHERE BEP_AFFAIRE="' + NumAppel + '" AND BEP_EQUIPERESS="'+ Equipe+ '"';
      SuppressionItemPlanning(SqlB);
    end Else If TobParam.GetValue('RETOUR') = 3 Then
    begin
      SqlB := ' WHERE BEP_CODEEVENT="' + IntToStr(NumEvenement) + '"';
      SuppressionItemPlanning(SqlB);
    end;
  FINALLY
    TraitSuppItem(CodeRes, NumAppel, IntToStr(NumEvenement));

    CodeDossierRech.text  := '';
    Label11.caption       := '';
  END;

end;



{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 23/03/2016
Modifié le ... :   /  /
Description .. : Procedure de gestion des evenements Parc/Matériel
Mots clefs ... :
*****************************************************************}
Procedure TMonPlanning.TraitementParcMateriel;
Var TobEventMat : TOB;
    StWhere     : String;
begin

  if (not BAffAppel1.Down) Then Exit;

  StWhere := 'CODEETAT=ECO';

  if (TypePlanning = 'PFM') then
    StWhere :=  StWhere + ';CODEFAMILLE=' + PlanningCurrent.CodeOnglet
  else if (TypePlanning = 'PPA') then
    StWhere := StWhere + ';BTETAT=' + PlanningCurrent.CodeOnglet;

  TobEventMat := TOB.Create('EVENT PARC-MAT', nil, -1);

  TheTob := TobEventMat;

  //FV1 - 15/12/2016 : FS#2263 - Divers pbs en parc matériels
  AGLLanceFiche('BTP','BTEVENTMAT_MUL','','',StWhere + ';CODEETAT=ARE;PLANNING');

  if TheTob <> nil then
  begin
    TraiteLigneSelectionne(TobEventMat);
    FreeAndNil(TobEventMat);
  end;

end;


Procedure TMonPlanning.TraiteLigneSelectionne(TobEventMat : TOB);
Var Indice : Integer;
begin

  for Indice := 0 to TobEventMat.detail.count -1 do
  begin
    TobEventMat.Detail[Indice].AddChampSupValeur('ACTION',          TobEventMat.GetValue('ACTION'));
    TobEventMat.Detail[Indice].AddChampSupValeur('DATEIMPUTATION',  TobEventMat.GetValue('DATEIMPUTATION'));
    TraitementEventParcMateriel(TobEventMat.Detail[Indice]);
  end;

end;

procedure TMonPlanning.CreateEventMateriel;
Var NumMateriel   : String;
    TobParam 	    : Tob;
    TobResItem    : Tob;
    LignePlan	    : Integer;
    IndexPage     : Integer;
    Sender		    : TObject;
    NumEvent	    : string;
    //
    TheAction     : string;
    DebutDemande  : TDateTime;
    FinDemande    : TDateTime;
    retour        : TDateTime;
    TitreEtat     : string;
    NbHeure       : Double;
    StSQL         : string;
    QQ            : TQuery;

begin

  //Chargement du type d'action par défaut ou du type d'action paramétré
  //
  StSQL := 'SELECT BTA_BTETAT, BTA_LIBELLE FROM BTETAT WHERE BTA_DEFAUT="X" AND BTA_TYPEACTION="PMA"';
  QQ := OpenSQL(StSQL, false, 1, '', true);
  if not QQ.Eof then
  begin
    TobItem.PutValue('BPL_BTETAT',    QQ.FindField('BTA_BTETAT').AsString);
    TobItem.PutValue('BPL_LIBACTION', QQ.FindField('BTA_LIBELLE').AsString);
  end
  else
  begin
    TobItem.PutValue('BPL_BTETAT', '');
    TobItem.PutValue('BPL_LIBACTION','');
  end;
  Ferme(QQ);

  //positionnement Ligne du planning
  TobParam      := Tob.Create('BTEVENTMAT', nil, -1);
  //
  TobParam.PutValue('BEM_IDEVENTMAT', 0);
  TobParam.PutValue('BEM_BTETAT',       TOBItem.GetString('BPL_BTETAT'));
  TobParam.PutValue('BEM_CODEMATERIEL', TOBItem.GetString('BPL_MATERIEL'));
  TobParam.PutValue('BEM_CODEETAT',     'ARE');
  TobParam.PutValue('BEM_DATEDEB',      TobItem.GetValue('BPL_DATEDEB'));
  TobParam.PutValue('BEM_DATEFIN',      TobItem.GetValue('BPL_DATEFIN'));
  TobParam.PutValue('BEM_HEUREDEB',     TOBItem.GetValue('BPL_HEUREDEB'));
  TobParam.PutValue('BEM_HEUREFIN',     TOBItem.GetValue('BPL_HEUREFIN'));
  TobParam.PutValue('BEM_NBHEURE',      TobItem.GetValue('BPL_DUREE'));
  TobParam.PutValue('BEM_AFFAIRE',      TOBItem.GetString('BPL_AFFAIRE'));
  TobParam.PutValue('BEM_TIERS',        TOBItem.GetString('BPL_TIERS'));
  TobParam.PutValue('BEM_RESSOURCE',    TOBItem.GetString('BPL_RESSOURCE'));
  TobParam.AddChampSupValeur('RETOUR',      -1);

  TheTob := TobParam;

  if (not ModeSaisie) then TheAction := 'ACTION=CONSULTATION' else TheAction := 'ACTION=CREATION';

  AglLanceFiche ('BTP','BTEVENTMAT','','',theAction+';PLANNING');

  TheTob := nil;

  //chargement des info pour affichage de l'item
  if TOBParam.GetValue('RETOUR') = 1 then IndexPage := PageOnglet.ActivePageIndex;

  TobParam.free;

  //chargement des info pour affichage de l'item
  ChargeItems;


end;

procedure TMonPlanning.ModifEventMateriel;
Var NumEventMat : Integer;
    DateDeb     : TDateTime;
    DateFin     : TDateTime;
    Argument    : string;
    IndexPage   : Integer;
    NbHeure		  : Double;
begin

  if Assigned(TobItem) then NumEventMat := TobItem.GetValue('BPL_IDEVENT');

  if NumEventMat <> 0 then
  Begin
    Argument := 'IDEVENEMENT=' + IntToStr(NumEventMat) + ';ACTION=MODIFICATION';
    AglLanceFiche ('BTP','BTEVENTMAT','','',Argument+';PLANNING');
    //AGLLanceFiche('BTP','BTEVENTMAT','','',Argument);
    //chargement des info pour affichage de l'item
    IndexPage := PageOnglet.ActivePageIndex;
  end;

  ChargeItems;  

end;

//Suppression d'un événement Planning Matériel
function  TMonPlanning.DeleteEvenementMateriel : boolean;
Var NumEvent    : Integer;
    Materiel    : string;
    Ressource   : string;
    Marequete   : string;
    TobDelete   : TOB;
begin

  Result := False;

  if Assigned(TobItem) then
  Begin
    NumEvent    := TobItem.GetValue('BPL_IDEVENT');
    Materiel    := TobItem.GetValue('BPL_MATERIEL');
    Ressource   := TobItem.GetValue('BPL_RESSOURCE');
  end;

  TobDelete := TobItem.FindFirst(['BPL_CODEEVENT'], [NumEvent], True);

  Marequete :='DELETE FROM BTEVENTMAT WHERE BEM_IDEVENTMAT='+IntToStr(NumEvent)+' AND BEM_CODEMATERIEL="'+ Materiel +'"';

  if ExecuteSql (Marequete) = 0 then
  begin
    PlanningCurrent.DeleteItem(TobItem);
    Result := True;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 25/03/2016
Modifié le ... :   /  /
Description .. : Gestion de l'envoi de mail pour l'ensemble des actions
Suite ........ : du planning le nécessitant
Mots clefs ... : 
*****************************************************************}
Procedure TMonPlanning.EnvoieMail(fSujetMail  : String);
Var EnvoiMail   : TGestionMail;
begin

  EnvoiMail := TGestionMail.Create(self);

  EnvoiMail.Sujet := fSujetMail;

  EnvoiMail.Corps := hTStringList.Create;
  EnvoiMail.Corps.Clear ;

  EnvoiMail.Copie         := '';
  EnvoiMail.TypeContact   := '';
  EnvoiMail.Fournisseur   := '';
  EnvoiMail.FichierSource := '';
  EnvoiMail.FichierTempo  := '';
  EnvoiMail.Fichiers      := '';
  EnvoiMail.TypeDoc       := '';
  //Pourrait être déterminé par le type d'enregistrement traité ou par le type de planning (????)
  EnvoiMail.Tiers         := '';
  EnvoiMail.Contact       := '';
  EnvoiMail.Destinataire  := '';
  EnvoiMail.QualifMail    := 'PLA';
  EnvoiMail.TobRapport    := TobItem;

  if EnvoiMail.QualifMail = '' then
    EnvoiMail.GestionParam  := False
  else
    EnvoiMail.GestionParam  := True;

  EnvoiMail.AppelEnvoiMail;

  FreeAndNil(EnvoiMail);

end;

{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 30/03/2016
Modifié le ... :   /  /
Description .. : Procedure de gestion des evenements Chantier
Mots clefs ... :
*****************************************************************}
Procedure TMonPlanning.TraitementPlanCharge;
Var TobEventCha : TOB;
    StWhere     : String;
begin

  if (not BAffAppel2.Down) Then Exit;

  TobEventCha := TOB.Create('EVENT CHANTIER', nil, -1);

  TheTob := TobEventcha;

  AGLLanceFiche('BTP','BTEVENTMAT_MUL','','','PLANNING');

  TraiteLigneSelectionne(TobEventCha);

  FreeAndNil(TobEventcha);

end;


procedure TMonPlanning.CreateEventPlanCharge;
Var NumMateriel   : String;
    TobParam 	    : Tob;
    TobResItem    : Tob;
    LignePlan	    : Integer;
    IndexPage     : Integer;
    Sender		    : TObject;
    NumEvent	    : string;
    //
    TheAction     : string;
    DebutDemande  : TDateTime;
    FinDemande    : TDateTime;
    retour        : TDateTime;
    TitreEtat     : string;
    NbHeure       : Double;
    
    StSQL         : string;
    QQ            : TQuery;

begin

  //Chargement du type d'action par défaut ou du type d'action paramétré
  //
  StSQL := 'SELECT BTA_BTETAT, BTA_LIBELLE FROM BTETAT WHERE BTA_DEFAUT="X" AND BTA_TYPEACTION="PCA"';
  QQ := OpenSQL(StSQL, false, 1, '', true);
  if not QQ.Eof then
  begin
    TobItem.PutValue('BPL_BTETAT',    QQ.FindField('BTA_BTETAT').AsString);
    TobItem.PutValue('BPL_LIBACTION', QQ.FindField('BTA_LIBELLE').AsString);
  end
  else
  begin
    TobItem.PutValue('BPL_BTETAT', '');
    TobItem.PutValue('BPL_LIBACTION','');
  end;
  Ferme(QQ);

  //positionnement Ligne du planning
  TobParam      := Tob.Create('BTEVENTCHA', nil, -1);
  //
  TobParam.PutValue('BEC_IDEVENTCHA','');
  TobParam.PutValue('BEC_BTETAT',    TOBItem.GetString('BPL_BTETAT'));

  if TypePlanning = 'PAR' then
  begin
    TobParam.PutValue('BEC_RESSOURCE', '');
    if EdtDosEnCours.text <> '' then
      TobParam.PutValue('BEC_AFFAIRE', EdtDosEnCours.text)
    else
      TobParam.PutValue('BEC_AFFAIRE', TOBItem.GetString('BPL_AFFAIRE'));
    TobParam.PutValue('BEC_TIERS',     TOBItem.GetString('BPL_TIERS'));
  end
  else
  begin
    TobParam.PutValue('BEC_RESSOURCE', TOBItem.GetString('BPL_RESSOURCE'));
    TobParam.PutValue('BEC_FONCTION',  TOBItem.GetString('BPL_FONCTION'));
    TobParam.PutValue('BEC_EQUIPERESS',TOBItem.GetString('BPL_EQUIPERESS'));
    TobParam.PutValue('BEC_AFFAIRE',   EdtDosEnCours.text);
    TobParam.PutValue('BEC_TIERS',     '');
  end;

  TobParam.PutValue('BEC_NUMPHASE',    StrToInt(NumPhase.text));

  TobParam.PutValue('BEC_DATEDEB',     DatetoStr(TOBItem.GetValue('BPL_DATEDEB')));
  TobParam.PutValue('BEC_DATEFIN',     DatetoStr(TOBItem.GetValue('BPL_DATEFIN')));
  TobParam.PutValue('BEC_HEUREDEB',    TOBItem.GetValue('BPL_HEUREDEB'));
  TobParam.PutValue('BEC_HEUREFIN',    TOBItem.GetValue('BPL_HEUREFIN'));
  TobParam.PutValue('BEC_DUREE',       TOBItem.GetValue('BPL_DUREE'));
  //
  TobParam.AddChampSupValeur('BEC_IDAFFECT',   0);
  TobParam.AddChampSupValeur('RETOUR',        -1);

  TheTob := TobParam;

  if (not ModeSaisie) then TheAction := 'ACTION=CONSULTATION' else TheAction := 'ACTION=CREATION';

  AglLanceFiche ('BTP','BTEVENTCHA','','',theAction+';PLANNING');

  TheTob := nil;

  //chargement des info pour affichage de l'item
  if TobParam.GetValue('RETOUR') = 1 then
  begin
    IndexPage           := PageOnglet.ActivePageIndex;
    EdtDosEnCours.text  := TobParam.GetValue('BEC_AFFAIRE');
    NUMPHASE.Text       := TobParam.GetValue('BEC_NUMPHASE');
  end;

  FreeAndNil(TobParam);

  //chargement des info pour affichage de l'item
  ChargeItems;

end;

procedure TMonPlanning.ModifEventPlanCharge;
Var NumEventCha : Integer;
    DateDeb     : TDateTime;
    DateFin     : TDateTime;
    Argument    : string;
    IndexPage   : Integer;
    NbHeure		  : Double;
begin

  if Assigned(TobItem) then NumEventCha := TobItem.GetValue('BPL_IDEVENT');

  if NumEventCha <> 0 then
  Begin
    Argument := 'IDEVENEMENT=' + IntToStr(NumEventCha) + ';ACTION=MODIFICATION';
    //AGLLanceFiche('BTP','BTEVENTMAT','','',Argument);
    AglLanceFiche ('BTP','BTEVENTCHA','','',Argument+';PLANNING');
    //chargement des info pour affichage de l'item
    IndexPage := PageOnglet.ActivePageIndex;
    ChargeItems;
  end
  else
  begin
    CreateEventPlanCharge
  end;
  
end;

//Suppression d'un événement Planning Matériel
function  TMonPlanning.DeleteEvenementPlanCharge : boolean;
Var NumEvent    : Integer;
    Ressource   : string;
    Marequete   : string;
    TobDelete   : TOB;
    rep         : String;
begin

  Result := False;

 if Assigned(TobItem) then
  Begin
    NumEvent    := TobItem.GetValue('BPL_IDEVENT');
    Ressource   := TobItem.GetValue('BPL_RESSOURCE');
    Equipe      := TobItem.GetValue('BPL_EQUIPERESS');
  end;

  //gestion des équipes...
  if Equipe <> '' then
  begin
    if PGIAsk(TraduireMemoire('La Ressource appartient à l''Equipe ' + Equipe + '.' + Chr(10) + 'Voulez-vous supprimer toute l''equipe ?'), rep) = mrYes then
    begin
      Result := DeleteEquipePlancharge(Equipe, Ressource,NumEvent);
    end;
  end
  else
  begin
     Marequete :='DELETE FROM BTEVENTCHA WHERE BEC_IDEVENTCHA='+IntToStr(NumEvent)+' AND BEC_RESSOURCE="'+ Ressource +'"';

    if ExecuteSql (Marequete) = 0 then
    begin
      PlanningCurrent.DeleteItem(TobItem);
      Result := True;
    end;
  end;

end;

Function TMonPlanning.DeleteEquipePlancharge(Equipe, Ressource : string;  NumEvent : integer ) : Boolean;
Var StSQL     : String;
    QQ        : TQuery;
    IdAffect  : Integer;
begin

  Result := False;

  //Recherche d'un enregistrement correspondant à l'item sélectionné
  StSQL := 'SELECT BEC_IDAFFECT FROM BTEVENTCHA WHERE BEC_IDEVENTCHA='+ IntToStr(NumEvent) + ' AND BEC_RESSOURCE="' + Ressource + '" AND BEC_EQUIPERESS="' + Equipe + '"';
  QQ := OpenSQL(StSQL, True, -1,'', True);

  If not QQ.eof then IdAffect := QQ.FindField('BEC_IDAFFECT').AsInteger;
  Ferme(QQ);

  if IdAffect = 0 then Exit;

  StSQL := 'Delete BTEVENTCHA WHERE BEC_IDAFFECT=' + IntToStr(Idaffect);
  if ExecuteSql (StSQL) = 0 then Result := True;

end;

{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 04/05/2016
Modifié le ... :   /  /    
Description .. : Actions sur une item
Mots clefs ... : 
*****************************************************************}
Procedure TMonPlanning.CreationItems(LigItem : Tob; TypeCreation : string);
Var LignePlan   : Integer;
    StSQL       : string;
    QQ          : TQuery;
    DateDeb     : TDateTime;
    DateFin     : TDateTime;
    HeureDeb    : TDateTime;
    HeureFin    : TDateTime;
    NbHeure     : Double;
begin

	if (not ModeSaisie) then exit;

  OkAffect := False;

  TobItem.InitValeurs(False);

  LignePlan   := PlanningCurrent.Row - PlanningCurrent.FixedRows;

  ChargementTobItem(PlanningCurrent.TobRes.Detail[LignePlan]);

  if not ControleRessourceDispo(TobItem, CtrlCal) then Exit;

  //On devrait faire le calcul de la durée ici une fois tous les contrôles passés....
  //
  DateDeb := TobItem.GetDateTime('BPL_DATEDEB');
  DateFin := TobItem.GetDateTime('BPL_DATEFIN');
  HeureDeb:= TobItem.GetDateTime('BPL_HEUREDEB');
  HeureFin:= TobItem.GetDateTime('BPL_HEUREFIN');
  //
  NbHeure := CalculDureeEvenement(DateDeb, DateFin);
  TobItem.PutValue('BPL_DUREE', NbHeure);

  //Si nous sommes sur une item vide à ce moment on crée la fiche qui va bien en fonction du
  //type de planning... Chaque planning permet de créer sa propre fiche...
  if      TypeCreation = '5' then     CreationActionGRC
  else if TypeCreation = '-' then
  begin
    if      ModePlanning = 'PLA' then CreateEvenementIntervention
    Else if ModePlanning = 'PMA' then CreateEventMateriel
    Else if ModePlanning = 'PCA' then CreateEventPlanCharge;
  end
  else IF TypeCreation = 'X' then CreationEvenement(LignePlan);

end;

Procedure TmonPlanning.ModificationItem(LigItem : Tob);
begin

  if (not ModeSaisie) then exit;
  //
  if LigItem = nil then exit;
  //
  if not IsModifiable(LigItem) then exit;
  //
  TobItem.InitValeurs(False);

  TobItem.Dupliquer(LigItem, False, True);
  //
  //AffichageDateHeure(TobItem, '+');
  //
  if (TobItem.GetString('BPL_ORIGINEITEM') = 'INTERV') then
  begin
    if TobItem.GetString('BPL_AFFAIRE') <> '' then
    begin
      VisualisationAffectation;
      ChargeItems;
    end
    else
      ModificationEvenementIntervention;
  end
  else
  begin
    if (TobItem.GetString('BPL_ORIGINEITEM') = 'PARCMAT') then
      ModifEventMateriel
    else if (TobItem.GetString('BPL_ORIGINEITEM') = 'ACT-GRC') then
      ModificationsActionGRC
    else if (TobItem.GetString('BPL_ORIGINEITEM') = 'CHANTIER') then
    Begin
      ModifEventPlanCharge
    end;
  end;

end;

Procedure TmonPlanning.SuppressionItem(LigItem : Tob);
Var Origine : String;
begin

	if (not ModeSaisie) or (LigItem = nil) then
  BEGIN
    CancelAction  := True;
    exit;
  END;

  if not IsModifiable (LigItem) then  exit;

  If pgiAsk('Confirmez-vous la suppression de l''évènement ?')=MrNo then Exit;

  TobItem.InitValeurs(False);

  TobItem.Dupliquer(LigItem, False, True);

  Origine := TobItem.getString('BPL_ORIGINEITEM');

  if (Origine = 'PARCMAT') then
  begin
    DeleteEvenementMateriel;
  end
  else if (Origine = 'INTERV') then
  begin
    if TobItem.GetString('BPL_AFFAIRE') <> '' then
    begin
      SuppressionEvenementAppel;
    end else
    begin
      SuppressionEvenementIntervention;
    end;
  end
  else if (Origine = 'ACT-GRC')  then SuppressionActionGRC
  else if (Origine = 'CHANTIER') then DeleteEvenementPlanCharge;
  BRechargerClick(self);
end;

Procedure TMonPlanning.ChargementTobItem(TobRes : Tob);
Var StSQL : string;
    QQ    : TQuery;
begin

  TobItem.PutValue('BPL_ORIGINEITEM', TobRes.GetString('ORIGINEITEM'));

  if TobRes.GetString('ORIGINEITEM') = 'INTERV' then
    TobItem.PutValue('BPL_TYPEACTION',  'INT')
  Else if TobRes.GetString('ORIGINEITEM') = 'ACT-GRC' then
    TobItem.PutValue('BPL_TYPEACTION',  'INT')
  Else if TobRes.GetString('ORIGINEITEM') = 'ABSPAIE' then
    TobItem.PutValue('BPL_TYPEACTION',  'INT')
  else if TobRes.GetString('ORIGINEITEM') = 'CHANTIER' then
    TobItem.PutValue('BPL_TYPEACTION',  'PCA')
  else if TobRes.GetString('ORIGINEITEM') = 'PARCMAT' then
    TobItem.PutValue('BPL_TYPEACTION' , 'PMA');
 
  if ModePlanning = 'PCA' then
  begin
    if TypePlanning = 'PAR' then
    begin
      TOBItem.PutValue('BPL_AFFAIRE'   ,    TobRes.GetValue('AFF_AFFAIRE'));
      TOBItem.PutValue('BPL_TIERS'     ,    TobRes.GetValue('AFF_TIERS'));
      TOBItem.PutValue('BPL_RESSOURCE' ,    TobRes.GetValue('AFF_RESPONSABLE'));
    end
    else if (TypePlanning = 'PRA') OR (TypePlanning = 'PFO') then
    begin
      TobItem.PutValue('BPL_RESSOURCE',     TobRes.GetValue('ARS_RESSOURCE'));
      TobItem.PutValue('BPL_NOMRESSOURCE',  TobRes.GetValue('ARS_LIBELLE') + ' ' + TobRes.GetValue('ARS_LIBELLE2'));
      TobItem.PutValue('BPL_VILLE',         TobRes.GetValue('ARS_VILLE'));
      TobItem.PutValue('BPL_TYPERESSOURCE', TobRes.GetValue('ARS_TYPERESSOURCE'));
      TobItem.PutValue('BPL_STYPERESSOURCE',TobRes.GetValue('ARS_CHAINEORDO'));
      TobItem.PutValue('BPL_FONCTION',      TobRes.GetValue('ARS_FONCTION1'));
      TobItem.PutValue('BPL_EQUIPERESS',    TobRes.GetValue('ARS_EQUIPERESS'));
      TobItem.PutValue('BPL_EMAIL',         TobRes.GetValue('ARS_EMAIL'));
      TobItem.PutValue('BPL_LIBFONCTION',   TobRes.GetValue('AFO_LIBELLE'));
      TobItem.PutValue('BPL_STANDCALEN',    TobRes.GetValue('ARS_STANDCALEN'));
      TobItem.PutValue('BPL_CALENSPECIF',   TobRes.GetValue('ARS_CALENSPECIF'));
    end;
  end
  else if ModePlanning = 'PMA' then
  begin
    TobItem.PutValue('BPL_FAMILLEMAT',    TobRes.GetValue('BMA_CODEFAMILLE'));
    TobItem.PutValue('BPL_MATERIEL',      TobRes.GetValue('BMA_CODEMATERIEL'));
    TobItem.PutValue('BPL_LIBMATERIEL',   TobRes.GetValue('BMA_LIBELLE'));
    TobItem.PutValue('BPL_LIBFAMILLEMAT', TobRes.GetValue('BFM_LIBELLE'));
    TobItem.PutValue('BPL_STANDCALEN',    '');
    TobItem.PutValue('BPL_CALENSPECIF',   '');
    TobItem.PutValue('BPL_RESSOURCE',     '');
    TobItem.PutValue('BPL_NOMRESSOURCE',  '');
    TobItem.PutValue('BPL_VILLE',         '');
    TobItem.PutValue('BPL_TYPERESSOURCE', '');
    TobItem.PutValue('BPL_STYPERESSOURCE','');
    TobItem.PutValue('BPL_FONCTION',      '');
    TobItem.PutValue('BPL_EQUIPERESS',    '');
    TobItem.PutValue('BPL_EMAIL',         '');
    TobItem.PutValue('BPL_LIBFONCTION',   '');
  end
  else
  begin     
    if Pos(TypePlanning,'PLA;PSF;PTA;PTR') > 0  then
    begin
      if TobItem.GetValue('BPL_BTETAT') = '' then TobItem.PutValue('BPL_BTETAT',GetParamSocSecur('SO_TYPEACTION',''));
      //
      TobItem.PutValue('BPL_RESSOURCE',     TobRes.GetValue('ARS_RESSOURCE'));
      TobItem.PutValue('BPL_NOMRESSOURCE',  TobRes.GetValue('ARS_LIBELLE') + ' ' + TobRes.GetValue('ARS_LIBELLE2'));
      TobItem.PutValue('BPL_VILLE',         TobRes.GetValue('ARS_VILLE'));
      TobItem.PutValue('BPL_TYPERESSOURCE', TobRes.GetValue('ARS_TYPERESSOURCE'));
      TobItem.PutValue('BPL_STYPERESSOURCE',TobRes.GetValue('ARS_CHAINEORDO'));
      TobItem.PutValue('BPL_FONCTION',      TobRes.GetValue('ARS_FONCTION1'));
      TobItem.PutValue('BPL_EQUIPERESS',    TobRes.GetValue('ARS_EQUIPERESS'));
      TobItem.PutValue('BPL_EMAIL',         TobRes.GetValue('ARS_EMAIL'));
      TobItem.PutValue('BPL_LIBFONCTION',   TobRes.GetValue('AFO_LIBELLE'));
      TobItem.PutValue('BPL_STANDCALEN',    TobRes.GetValue('ARS_STANDCALEN'));
      TobItem.PutValue('BPL_CALENSPECIF',   TobRes.GetValue('ARS_CALENSPECIF'));
      TobItem.PutValue('BPL_FAMILLEMAT',    '');
      TobItem.PutValue('BPL_FAMILLEMAT',    '');
      TobItem.PutValue('BPL_LIBMATERIEL',   '');
      TobItem.PutValue('BPL_LIBFAMILLEMAT', '');
    end;
  end;

  ChargementDateItem(TobRes, nil);

end;

Procedure TMonPlanning.ChargementDateItem(TobRes, Item : Tob);
Var Duree : Double;
    DateDeb   : TDateTime;
    DateFin   : TDateTime;
    HeureDeb  : TTime;
    HeureFin  : TTime;
    An        : Word;
    Mois      : Word;
    Jour      : Word;
    H, M, S, Ms : Word;
begin

  if Item <> nil then
  begin
    DateDeb   := StrToDatetime(Item.GetValue('BPL_DATEDEB'));
    DateFin   := StrToDatetime(Item.GetValue('BPL_DATEFIN'));
    TobItem.PutValue('MODIFIABLE',  Item.GetValue('MODIFIABLE'));
  end
  else
  begin
    //
    DateDeb   := PlanningCurrent.GetDateOfCol(PlanningCurrent.Col);
    DateDeb   := CalculDateDebut(DateDeb);
    DecodeDateTime(DateDeb, An, Mois, Jour, H, M, S, MS);
    HeureDeb  := EncodeTime(H, M, 0, 0);
    //
    DateFin   := PlanningCurrent.GetDateOfCol(PlanningCurrent.Col);
    DateFin   := CalculDateFin(DateFin);
    DecodeDateTime(Datefin, An, Mois, Jour, H, M, S, MS);
    HeureFin  := encodeTime(H, M, 0, 0);
  end;
  //
  //Calcul de la date de début et de la date de fin
  //
  TobItem.PutValue('BPL_DATEDEB', DateTimeToStr(DateDeb));
  TobItem.PutValue('BPL_DATEFIN', DateTimeToStr(DateFin));

  TobItem.PutValue('BPL_HEUREDEB', TimeToStr(HeureDeb));
  TobItem.PutValue('BPL_HEUREFIN', TimeToStr(HeureFin));

end;

Procedure TMonPlanning.DefinitZoneDate;
Var DateDeb : TDateTime;
    DateFin : TDateTime;
    //
    DblHeure: Double;
    //
    NbHeure : Double;
    StHeure : String;
    StSQL   : string;
    QQ      : TQuery;
begin

  // verification disponibilité
  DateDeb := TobItem.GetValue('BPL_DATEDEB');
  DateFin := TobItem.GetValue('BPL_DATEFIN');

  NbHeure := TobItem.GetValue('BPL_DUREE');

  FormatageDate(DateDeb);
  FormatageDate(DateFin);

  TobItem.PutValue('BPL_DATEDEB', DateDeb);
  TobItem.PutValue('BPL_DATEFIN', DateFin);

  StHeure := TimeToStr(DateDeb);
  DblHeure:= StrToTime(StHeure);
  ControleHeureDeb(DblHeure);
  TobItem.PutValue('BPL_HEUREDEB', TimeToStr(DblHeure));
  //
  StHeure := TimeToStr(DateFin);
  DblHeure:= StrToTime(StHeure);
  ControleHeureFin(DblHeure);
  TobItem.PutValue('BPL_HEUREFIN', TimetoStr(DblHeure));

  Nbheure := CalculDureeEvenement(DateDeb,DateFin);

  //Si la durée est égale à zéro on charge celle du type Action...
  if NbHeure = 0 then
  begin
    StSQL := 'Select BTA_DUREEMINI FROM BTETAT WHERE BTA_TYPEACTION="' + TobItem.GetValue('BPL_TYPEACTION') + '" AND BTA_BTETAT="' + TobItem.GetValue('BPL_BTETAT') + '"';
    QQ    := OpenSQL(StSQl, False, -1,'',True);

    if not QQ.eof then
    begin
      if NbHeure < QQ.findfield('BTA_DUREEMINI').Asfloat then NbHeure := QQ.findfield('BTA_DUREEMINI').Asfloat;
      NbHeure := Nbheure * 60;
    end;

    Ferme(QQ);
  end;

  if NbHeure = 0 then NbHeure := 60;

  TobItem.PutValue('BPL_DUREE', NbHeure);

end;

procedure TMonPlanning.BTEditEvtClick(Sender: TObject);
Var TobEvt        : TOB;
    TobEdition    : TOB;
    TobRessource  : TOB;
    StSQL         : String;
		StModele	    : String;
		TitreEtat	    : String;
begin

  //il faut gérer l'édition en fonction de l'événement sélectionné
  TobEvt := PlanningCurrent.GetCurItem;

  TobEdition := Tob.create('EDITION', nil, -1);

  StSQL := 'SELECT * FROM (AFFAIRE LEFT JOIN TIERS ON AFF_TIERS = T_TIERS) ';
  StSQL := StSQL + 'LEFT JOIN RESSOURCE ON AFF_RESPONSABLE = ARS_RESSOURCE ';
  StSQL := StSQL + 'LEFT JOIN ADRESSES ON AFF_AFFAIRE = ADR_REFCODE ';
  StSQL := StSQL + 'WHERE AFF_STATUTAFFAIRE = "APP" ';
  StSQL := StSQL + '  AND AFF_AFFAIRE="'  + TobEvt.GetValue('BPL_AFFAIRE') + '"';

  TobEdition.LoadDetailFromSQL(StSql);

  //Lecture de la tobEdition pour chargement des valeurs de la ressources si différentes
  //de la ressource principale
  if Tobedition.detail[0].getValue('AFF_RESPONSABLE') <> TobEvt.GetValue('BPL_RESSOURCE') then
  Begin
    StSQL := 'SELECT * FROM RESSOURCE WHERE ARS_RESSOURCE="' + TobEvt.GetValue('BPL_RESSOURCE')  + '"';
    TobRessource := Tob.Create('RESSOURCE', nil, -1);
    TobRessource.LoadDetailFromSQL(StSql);
    ChargeTobDestinationWithTobSource(TobEdition.detail[0], TobRessource.Detail[0]);
  end;

  stModele := GetParamSOC('SO_BTMODBI');

  V_PGI.NoPrintDialog := False;

  TitreEtat := 'Edition Intervention : ' + TobEvt.GetValue('BPL_AFFAIRE');
  //LanceEtat('E','INT',stModele,True, false, false,Nil,trim (stWhere),'INTERVENTION '+TitreEtat,False);
  LanceEtatTob('E','INT',stModele,TobEdition, True, False, False, nil, '', TitreEtat, False);

end;

Procedure TMonPlanning.ChargeTobDestinationWithTobSource(TobDest, TobSource : Tob);
Var Ind : Integer;
    Nomchamp : string;
begin

   ind := 1000;
   NomChamp := TobSource.GetNomChamp(ind);
   while NomChamp <> '' do
   begin
      if TobDest.fieldExists(Nomchamp) then TobDest.Putvalue(Nomchamp, TobSource.GetValue(Nomchamp));
      Inc(ind);
      NomChamp := TobSource.GetNomChamp(ind);
   end;

end;

End.


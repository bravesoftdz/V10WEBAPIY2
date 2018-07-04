{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 11/06/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSAISIEHEBDO ()
Mots clefs ... : TOF;BTSAISIEHEBDO
*****************************************************************}
Unit BTSAISIEHEBDO_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     grids,
     graphics,
     windows,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     uTob,
     fe_main,
{$else}
     MainEagl,
     eMul,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HTB97,
     Hpanel,
     utofAfBaseCodeAffaire,
     vierge,
     NewCalendar,
     messages,
     UtilModeleConso,
     SelectPhase,
     UtilPhases,
     UtofRessource_Mul,
     UtilsaisieConso,
     menus,
     lookup,
     UtilSuggestion,
     AffaireUtil,
     CalcOleGenericBtp;

Type

  TmsSaisie = (TmsEntete,TmsDetail);
  Tmodecomp = (TmcEqual,TmcOnlyModel,TmcOnlyConso,TmcNotEqual);
//
//
// la tob de saisie est définie ainsi
//
// 1er niveau : la ligne comportant les temps de chaque jours de la semaine sur une prestation d'un type donné
//   |
//   |--> 2 eme niveau le détail jour par jour des consommations
//
//
//
  TOF_BTSAISIEHEBDO = Class (TOF_AFBASECODEAFFAIRE)
  private
    TOBRessource : TOB;
    TOBTypeHeure : TOB;

    stprev : string;
    LastModele : string;
    fTypeRessource : string;
    fExistsConso : boolean;
    TOBSaisies   : TOB;
    TOBSaisies_O : TOB;
    TOBConso_O   : TOB;
    POPCHOIX     : TpopupMenu;
    POPDECID     : TpopupMenu;
    PENTETE      : THPanel;
    PBAS         : THPanel;
    PCALENDAR    : ThPanel;
    PINDIC       : ThPanel;
    Calendrier   : TmxCalendar;
    PHASE        : THEdit;
    LCHANTIER    : THlabel;
    CH_CHANTIER  : THEdit;
    CH_CHANTIER0 : THEdit;
    CH_CHANTIER1 : THEdit;
    CH_CHANTIER2 : THEdit;
    CH_CHANTIER3 : THEdit;
    CH_AVENANT 	 : THEdit;
    BEFFACEAFF1  : TToolbarButton97;
    BDELETE      : TToolbarButton97;
    LPHASE       : THLabel;
    BRECHPHASE,BRESSCHANTIER : TToolbarButton97;
    RESSOURCE    : THEdit;
    MODELE       : THedit;
    LRESSOURCE   : THLabel;
    LMODELE      : THLabel;
    BCHERCHE     : TToolbarButton97;
    BANNUL       : TToolbarButton97;
    BVALIDER     : TToolbarButton97;
    BFERME       : TToolbarButton97;
    GS           : THgrid;
    FFQTE        : string;
    QMODELE      : THnumedit;
    QCONSO       : THnumedit;
    CUMULSEMMOD  : THnumedit;
    CUMULSEMCON  : THnumedit;
    MnPrest,MnFrais : TmenuItem;
    MnCelluleMod,MnCelluleCon : TmenuItem;
    MnApplicAllMod,MnApplicAllCon : TmenuItem;
    // variable pour la saisie
    {NomList, }FNomTable, FLien, FSortBy, stCols, FTitre,stColListe : string;
    FLargeur, FAlignement, FParams, title, NC, FPerso: string;
    OkTri, OkNumCol : boolean;
    nbcolsInListe : integer;
    ValoPrestaFromRessource : Boolean; 
    //
    G_NUMORDRE,G_CODEARTICLE,G_LIBELLE,G_TYPEH,G_QTE1,G_QTE2,G_QTE3,G_QTE4,G_QTE5,G_QTE6,G_QTE7 : integer;
    // methode sur creation et liberation objet
    procedure GetComponents;
    procedure CreateCalendar;
    procedure SetEvents;
    procedure destroyCalendar;
    // Methode sur evenements de l'entete
    procedure InitEntete;
    procedure PhaseChange (Sender : Tobject);
    procedure RecherchePhase (Sender : TObject);
    procedure RessourceExit (Sender : TObject);
    procedure FindRessource (Sender : Tobject);
    procedure FindRessourceDuChantier (Sender : Tobject);
    procedure ModeleExit (Sender : TObject);
    //
    procedure LetsGoParty (Sender : TObject); // Passage en mode saisie du détail
    procedure AnnulationSaisie (Sender : Tobject); // retour a la saisie de l'entete sans validation
    //
    procedure SetModeSaisie (Mode : TmsSaisie);
    //
    procedure SetGridEvent (Actif : boolean); // positionne ou pas les evenements sur la grille
    // evenement sur grid et form
    procedure GSEnter(Sender: TObject); dynamic;
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer;  var Cancel: Boolean);
    procedure GSElipsisClick(Sender: TOBject);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer;  var Cancel: Boolean; Chg: Boolean);
		procedure GSPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GetCellCanvas(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);

    procedure GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    //
    procedure RetourEntete;
    function VerifEntete : boolean;
    // element de definition de la saisie
    procedure SetparamGrille;
    procedure DefiniGrille;
    procedure AfficheLagrille;
    procedure AfficheLigne(Grid: THgrid; TOBL: TOB; indice: integer);
    procedure AffaireEfface(Sender: TObject);
    procedure OnExitPartieAffaire(Sender: Tobject);
    procedure ChangeChantier(Sender: TObject);
    //
    procedure ReInitModeleSaisie;
    procedure AddLesChampsSups(TOBL : TOB);
    procedure ReinitSaisie;
    procedure CompleteSaisieViaConso;
    procedure DefiniRowCounts;
    procedure RattacheConsoSurModele (TOBConso : TOB);
    procedure PrepareLeModele;
    function  AjouteLigneModeleFromConso (TOBC : TOB) : TOB;
    procedure EnregistreConsoSurModele(TOBAT, TOBC: TOB);
    function GetIndiceTOB(LeJour: integer): integer;
    procedure DefiniCumulSemaine;
    function OKCompareSaisieConso(Acol, Arow: integer): Tmodecomp;
    function AjouteLigneModele: TOB;
    procedure PositionneDansGrid(Arow, Acol: integer);
    function IsLigneVide(Arow: integer): boolean;
    function ControleLigne(ligne: integer; var Acol: integer): boolean;
    procedure ZoneSuivanteOuOk(Grille: THGrid; var ACol, ARow: integer;var Cancel: boolean);
    function ZoneAccessible(Grille: THgrid; var ACol,ARow: integer): boolean;
    function FindCodeArticle(TOBL: TOB; Valeur: string): boolean;
    procedure ShowInfoQte (Acol,Arow : integer);
    procedure RechFRais(Sender: TObject);
    procedure RechPrestations(Sender: TObject);
    function VerifTypeHeure(TOBL: TOB; Valeur: string): boolean;
    function IsArticleExiste (Arow,Acol : integer) : boolean;
    function IsExisteFrais(TOBL: TOB): boolean;
    function IsExistePrestationInterne(TOBL: TOB): boolean;
    function VerifExistsHeure(valeur: string): boolean;
    procedure DeleteLigne;
    procedure ReindiceGrid;
    procedure Remplitgrille;
    procedure MemoriseResource;
    function ExisteLienRessourceAffaire: boolean;
    procedure DefiniMenuDecid(Arow, Acol: integer);
    procedure CopieValeurToModele(TOBL: TOB);
    //
    procedure MnCelluleModClick (Sender : Tobject);
    procedure MnCelluleConClick (Sender : Tobject);
    procedure MnApplicAllModClick (Sender : Tobject);
    procedure MnApplicAllConClick (Sender : Tobject);
    procedure CopieConsoToSaisie(TOBL: TOB);
    procedure CopieModeleToSaisie(TOBL: TOB);
    procedure SupprimelesConsos (Sender : TOBject);
    //
    procedure EnregistreLesConsos;
    procedure EcritLesConsos;
    procedure prepareEcritureConsos(TOBConsos: TOB);
    procedure AppliqueQteSaisieSurConso(TOBCONSO,TOBS: TOB; Jour: TDateTime;IndiceJour: integer; QteJ: double);
    procedure NettoieMoiCa(TOBCC: TOB);
    procedure CreeLigneConso(TOBS, TOBCC: TOB; Jour: TDateTime);
    procedure RecupLesInfosRessource;
    procedure SetvaloFromArticle(TOBL: TOB);
    procedure MetJourLaConso(TOBS,TOBL: TOB; QteJ: double);
    procedure RecupModeleorigine;
    procedure ReinitValSaisie;
    procedure RemplitTypeHeures;
    procedure SetValoFromTypeHeures(TOBS: TOB);
    //
  public
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4,Aff_, Aff0_,
                                Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;
    procedure bSelectAff1Click(Sender: TObject);override;

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

uses factutil,Paramsoc;

procedure TOF_BTSAISIEHEBDO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEHEBDO.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEHEBDO.OnUpdate ;
var cancel : boolean;
begin
  Inherited ;
  cancel := false;
  // Validation des lignes de consos
  GSRowExit(self,gs.row,cancel,false);
  if cancel then BEGIN TFVierge(Ecran).ModalResult := 0; Exit; END;
  //
  MemoriseResource;
  //
  if TRANSACTIONS (EnregistreLesConsos,0) = OeUnKnown then PGIInfo ('Erreur lors de la validation des consommations');
  //
  RetourEntete;
  TFVierge(Ecran).ModalResult := 0; // on ne sort pas en cliquant sur valide...na
end ;

procedure TOF_BTSAISIEHEBDO.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEHEBDO.OnArgument (S : String ) ;
var indice : integer;
begin
  Inherited ;
  ValoPrestaFromRessource := (not GetParamSocSecur ('SO_BTVALOSCONSO',True));
  FFQTE := '###';
  if V_PGI.OkDecQ > 0 then
  begin
    FFQTE := FFQTE+'0.';
    for indice := 1 to V_PGI.OkDecQ do
    begin
      FFQTE := FFQTE + '0';
    end;
  end;

  TOBSaisies := TOB.Create ('LA SAISIE DES TEMPS HEBDO',nil,-1);
  TOBSaisies_O := TOB.Create ('LA SAISIE DES TEMPS SAV',nil,-1);
  TOBConso_O   := TOB.Create ('LES CONSOS',nil,-1);
  TOBRessource := TOB.create ('RESSOURCE',nil,-1);
  TOBTypeHeure := TOB.Create ('LES TYPE HEURES',nil,-1);
  RemplitTypeHeures;

  fTypeRessource := '';
  //
  GetComponents;
  CreateCalendar;
  InitEntete;
  //
  SetEvents;
  SetModeSaisie (TmsEntete);   // définition des boutons de pied et de la grille en mode entete
  SetparamGrille;
  DefiniGrille;
  //
end ;

procedure TOF_BTSAISIEHEBDO.OnClose ;
begin
  Inherited ;
  destroyCalendar;
  TOBSaisies.free;
  TOBSaisies_O.free;
  TOBConso_O.free;
  TOBRessource.free;
  TOBTypeHeure.free;
end ;

procedure TOF_BTSAISIEHEBDO.InitEntete;
begin
  GS.Visible := false;
  LChantier.Caption := '';
  Lressource.Caption := '';
  LModele.Caption := '';
end;

procedure TOF_BTSAISIEHEBDO.destroyCalendar;
begin
  FreeAndNil (Calendrier);
end;

procedure TOF_BTSAISIEHEBDO.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEHEBDO.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEHEBDO.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,
  Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
  Aff  := THEdit(GetControl('CH_CHANTIER'));
  Aff0 := THEdit(GetControl('CH_CHANTIER0'));
  Aff1 := THEdit(GetControl('CH_CHANTIER1'));
  Aff2 := THEdit(GetControl('CH_CHANTIER2'));
  Aff3 := THEdit(GetControl('CH_CHANTIER3'));
  Aff4 := THEdit(GetControl('CH_AVENANT'));
end;

procedure TOF_BTSAISIEHEBDO.GetComponents;
begin
  POPCHOIX := TPopupMenu (GetCOntrol('POPCHOIX'));
  POPDECID := TPopupMenu (GetCOntrol('POPDECID'));
  PCALENDAR := THpanel (GetControl('PCALENDAR'));
  PENTETE := THpanel (GetControl('PENTETE'));
  PBAS := THpanel (GetControl('PBAS'));
  //
  CH_CHANTIER := THEdit (GetControl('CH_CHANTIER0'));
  CH_CHANTIER0 := THEdit (GetControl('CH_CHANTIER0'));
  CH_CHANTIER1 := THEdit (GetControl('CH_CHANTIER1'));
  CH_CHANTIER2 := THEdit (GetControl('CH_CHANTIER2'));
  CH_CHANTIER3 := THEdit (GetControl('CH_CHANTIER3'));
  CH_AVENANT 	 := THEdit (GetControl('CH_AVENANT'));
  LCHANTIER := THLabel (GetControl('LCHANTIER'));
  BEFFACEAFF1 := TToolBarButton97(getCOntrol('BEFFACEAFF1'));
  //
  PHASE := THEdit(GetControl('PHASE'));
  BRECHPHASE := TToolBarButton97(GetControl('BRECHPHASE'));
  LPHASE := THLabel (GetControl('LPHASE'));
  //
  RESSOURCE := THEdit(GetControl('RESSOURCE'));
  BRESSCHANTIER := TToolBarButton97(GetControl('BRESSCHANTIER'));
  MODELE    := THedit(GetControl('MODELE'));
  //
  LRESSOURCE := THLabel(GetControl('LRESSOURCE'));
  LMODELE    := THLabel(GetControl('LMODELE'));
  //
  BCHERCHE := TToolbarButton97(GetControl('BCHERCHE'));
  BANNUL   := TToolbarButton97(GetControl('BANNUL'));
  BVALIDER := TToolbarButton97(GetControl('BVALIDER'));
  BFERME   := TToolbarButton97(GetControl('BFERME'));
  BDELETE  := TToolbarButton97(GetControl('BDELETE'));
  //
  GS := THgrid(GetControl('GS'));
  //
  PINDIC := THpanel (GetControl('PINDIC'));
  QMODELE := THnumedit(GetControl('QMODELE')); QMODELE.Decimals := V_PGI.OkdecQ;
  QCONSO  := THnumedit(GetControl('QCONSO'));  QCONSO.Decimals := V_PGI.OkdecQ;
  CUMULSEMMOD  := THnumedit(GetControl('CUMULSEMMOD')); CUMULSEMMOD.Decimals := V_PGI.OkdecQ;
  CUMULSEMCON  := THnumedit(GetControl('CUMULSEMCON')); CUMULSEMCON.Decimals := V_PGI.OkdecQ;
  MnPrest := TMenuItem(GetControl('MNPREST'));
  MnFrais := TMenuItem(GetControl('MNFRAIS'));
  //
  MnCelluleMod  := TMenuItem(GetControl('MnCelluleMod'));
  MnCelluleCon  := TMenuItem(GetControl('MnCelluleCon'));
  MnApplicAllMod  := TMenuItem(GetControl('MnApplicAllMod'));
  MnApplicAllCon := TMenuItem(GetControl('MnApplicAllCon'));
end;

procedure TOF_BTSAISIEHEBDO.CreateCalendar;
begin
  Calendrier := TmxCalendar.Create (PCALENDAR);
  Calendrier.Align := alClient;
  calendrier.Parent := PCALENDAR;
  Calendrier.DateFormat := 'dd/mm/yyyy';
  calendrier.Options:=calendrier.Options+ [csSelectionEnabled,csWeekSelectionOnly];
end;


procedure TOF_BTSAISIEHEBDO.OnExitPartieAffaire(Sender: Tobject);
var iErr : integer;
begin
  if (CH_CHANTIER1.text = '') and (CH_CHANTIER2.text = '') and (CH_CHANTIER3.text = '') then exit;
  CH_CHANTIER.Text := DechargeCleAffaire(CH_CHANTIER0, CH_CHANTIER1, CH_CHANTIER2, CH_CHANTIER3, CH_AVENANT, '', TaCreat, False, True, false, Ierr);
end;

procedure TOF_BTSAISIEHEBDO.ChangeChantier(Sender: TObject);
var QQ     : TQuery;
    Sql    : String;
begin
  Ch_CHANTIER.text := Trim(UpperCase (Ch_CHANTIER.text));
  if GetChampsAffaire (CH_CHANTIER.Text,'AFF_ETATAFFAIRE')='TER' then
  begin
    PGIBox (TraduireMemoire('ATTENTION : Chantier clôturé'),ecran.caption);
    exit;
  end;

	Sql := 'SELECT AFF_LIBELLE FROM AFFAIRE WHERE AFF_AFFAIRE="'+Ch_CHANTIER.text+'"';

  QQ := OpenSql (SQL,true);
  if not QQ.eof then
	begin
     LCHANTIER.Caption := QQ.findfield('AFF_LIBELLE').asString ;
  end;
  RESSOURCE.Plus := ' AND AFT_AFFAIRE="'+Ch_CHANTIER.text+'" AND AFT_TYPEINTERV="RES"';
  ferme (QQ);
  LastModele := '';
end;

procedure TOF_BTSAISIEHEBDO.SetEvents;
begin

  CH_CHANTIER1.OnExit := OnExitPartieAffaire;
  CH_CHANTIER2.OnExit := OnExitPartieAffaire;
  CH_CHANTIER3.OnExit := OnExitPartieAffaire;
  //
  BEFFACEAFF1.onClick := AffaireEfface;

  //  methodes destinées a réagir aux evenements des zones d'entetes...
  CH_CHANTIER.OnChange := ChangeChantier;
  PHASE.OnChange := PhaseChange;
  BRECHPHASE.OnClick := RecherchePhase;
  RESSOURCE.onExit := RessourceExit;
  RESSOURCE.OnElipsisClick := FindRessource;
  BRESSCHANTIER.OnClick := FindRessourceDuChantier;
  MODELE.OnExit := ModeleExit;
  //.... et au pied (médor ??)
  BCHERCHE.onClick := LetsGoParty;
  BANNUL.onClick := AnnulationSaisie;
  BDELETE.onClick := SupprimelesConsos;
  //
  MnPrest.onClick := RechPrestations;
  MnFrais.onClick := RechFrais;
  //
  MnCelluleMod.OnClick  := MnCelluleModClick;
  MnCelluleCon.OnClick  := MnCelluleConClick;
  MnApplicAllMod.OnClick  := MnApplicAllModClick;
  MnApplicAllCon.OnClick := MnApplicAllConClick;
  //
  TFvierge(ecran).OnKeyDown := FormKeyDown;
end;

(*
procedure TOF_BTSAISIEHEBDO.FindModele(Sender: TObject);
begin

end;
*)

procedure TOF_BTSAISIEHEBDO.FindRessource(Sender: Tobject);
var retour : string;
begin
  retour := AFLanceFiche_Rech_Ressource ('','');
  if retour <> '' then
  begin
    RESSOURCE.text := retour;
    RessourceExit (self);
  end;

end;

procedure TOF_BTSAISIEHEBDO.FindRessourceDuChantier(Sender: Tobject);
var retour : string;
begin
  retour := AglLanceFiche ('BTP','BTRESSAFFAIRE_MUL','AFT_AFFAIRE='+CH_CHANTIER.text,'','');
  if retour <> '' then
  begin
    RESSOURCE.text := retour;
    RessourceExit (self);
  end;
end;

procedure TOF_BTSAISIEHEBDO.ModeleExit(Sender: TObject);
begin
  if Modele.text <> '' then
  begin
    LMODELE.caption := GetLibelleModele (MODELE.text,fTypeRessource);
  end else
  begin
    LMODELE.caption := '';
  end;
end;

procedure TOF_BTSAISIEHEBDO.PhaseChange(Sender: Tobject);
begin
  LPhase.Caption := GetLibellePhase (Ch_CHANTIER.text,Phase.text);
end;

procedure TOF_BTSAISIEHEBDO.RecherchePhase(Sender: TObject);
var Code : string;
begin
  if CH_CHANTIER.Text = '' then exit;
  //
  if SelectionPhase (Ch_CHANTIER.text,Code) then PHASE.Text := Code;
end;

procedure TOF_BTSAISIEHEBDO.RessourceExit(Sender: TObject);
var TheTypeRessource : string;
		Sql : string;
begin
  RESSOURCE.text := UpperCase_(RESSOURCE.text) ;
  //
  Sql :='SELECT 1 FROM RESSOURCE WHERE ARS_RESSOURCE="'+RESSOURCE.text+'"';
  if not ExisteSql (Sql) then
  begin
  	pgiInfo ('Erreur : Ressource inexistante...');
  	ressource.text := '';
    SetFocusControl('RESSOURCE');
    exit;  
  end;
  //
  LRESSOURCE.Caption  := GetMainDoeuvre (RESSOURCE.text);
  TheTypeRessource := GetTypeRessource (RESSOURCE.Text);
  if TheTypeRessource <> fTypeRessource then
  begin
    MODELE.Plus := ' AND BMS_TYPERESSOURCE="'+TheTypeRessource+'"';
    MODELE.Text := ''; LMODELE.caption := ''; LastModele := '';
    fTypeRessource := TheTypeRessource;
  end;
end;

procedure TOF_BTSAISIEHEBDO.AnnulationSaisie(Sender: Tobject);
begin
  if PgiAsk ('Vous êtes sur le point de sortir sans valider. Confimer ?')=Mryes then
  begin
    RetourEntete;
  end;
end;

procedure TOF_BTSAISIEHEBDO.LetsGoParty(Sender: TObject);
begin
  // contrôle de l'entete
  if TFvierge(ecran).ActiveControl.Name <> 'GS' Then
  begin
  	NextControl (TFvierge(ecran),true);
  end;

  if MODELE.text <> LastModele then
  begin
    ReInitModeleSaisie;
    LastModele := MODELE.text;
  end;

  if not VerifEntete then exit;
  //
  RecupLesInfosRessource;
  PrepareLeModele;
  CompleteSaisieViaConso;
  DefiniCumulSemaine;
  DefiniRowCounts;
  // grille remplie..on peut l'afficher
  AfficheLagrille;
  // --
  TFVierge(Ecran).HMTrad.ResizeGridColumns (GS);
  // et hop on active les evenements sur la grille
  GS.Visible := True;
  // entree dans la grille
  GS.SetFocus;
  GSEnter(self);
  SetGridEvent (True);
  SetModeSaisie (TmsDetail);
end;

procedure TOF_BTSAISIEHEBDO.SetModeSaisie(Mode: TmsSaisie);
begin
  BCHERCHE.visible := (Mode=TmsEntete);
  BANNUL.visible := (Mode=TmsDetail);
  BVALIDER.visible := (Mode=TmsDetail);
  BFERME.visible := (Mode=TmsEntete);
  PENTETE.Enabled := (Mode=TmsEntete);
  PBAS.visible := (Mode=TmsDetail);
  BDELETE.visible := (Mode=TmsDetail);
  //
end;

procedure TOF_BTSAISIEHEBDO.SetGridEvent(Actif: boolean);
begin
  if Actif then
  begin
    GS.OnEnter  := GSEnter;
    GS.OnRowEnter := GSRowEnter;
    GS.OnRowExit := GSRowExit;
    GS.OnCellEnter  := GSCellEnter;
    GS.OnCellExit  := GSCellExit;
    GS.OnElipsisClick := GSElipsisClick;
    GS.PostDrawCell := GSPostDrawCell;
    GS.GetCellCanvas := GetCellCanvas;
    GS.OnKeyDown := GSKeyDOwn;
    GS.PopupMenu := POPDECID;
  end
  else
  begin
    GS.OnEnter  := nil;
    GS.OnRowEnter := nil;
    GS.OnRowExit := nil;
    GS.OnCellEnter  := nil;
    GS.OnCellExit  := nil;
    GS.OnElipsisClick := nil;
    GS.PostDrawCell := nil;
    GS.GetCellCanvas := nil;
    GS.OnKeyDown := nil;
    GS.PopupMenu := nil;
  end;
end;

procedure TOF_BTSAISIEHEBDO.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    VK_ESCAPE:
    BEGIN
    	Key := 0;
    END;
  end;
end;


procedure TOF_BTSAISIEHEBDO.GSEnter(Sender: TObject);
var Acol,Arow : integer;
    cancel : boolean;
begin
  cancel := false;
  GS.row := 1;
  GS.col := 1;
  Arow := 1;
  Acol := 1;
  ZoneSuivanteOuOk(GS,ACol, ARow, Cancel);
  cancel := false;
  GSRowEnter (GS,Arow,cancel,false);
  GSCellEnter (GS,Acol,Arow,cancel);
  GS.row := Arow;
  GS.col := Acol;
  ShowInfoQte (Acol,Arow);
  DefiniMenuDecid (Arow,Acol);
  GS.ElipsisButton := ((Acol = G_CODEARTICLE) or (ACOL=G_TYPEH)) ;
  stprev := GS.Cells [Acol,Arow];
  //
end;

procedure TOF_BTSAISIEHEBDO.GSCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var TOBL : TOB;
begin
  ZoneSuivanteOuOk(GS,ACol, ARow, Cancel);
  if not Cancel then
  begin
    TOBL := TOBSaisies.detail[Arow-1];
    if TOBL <> nil then
    begin
      if (TOBL.GetValue('BSM_ARTICLE')='') and (Acol > G_CODEARTICLE) then BEGIN Acol := G_CODEARTICLE; PositionneDansGrid (Arow,Acol); END;
      if (TOBL.GetValue('BSM_ARTICLE')<>'') And (Acol = G_CODEARTICLE) then BEGIN Acol := G_LIBELLE; PositionneDansGrid (Arow,Acol); END;
    end;
    DefiniMenuDecid (Arow,Acol);
    ShowInfoQte (Acol,Arow);
    GS.ElipsisButton := (Acol = G_CODEARTICLE) OR (ACOl = G_TYPEH);
    stprev := GS.Cells [Acol,Arow];
  end;
end;

procedure TOF_BTSAISIEHEBDO.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Var TOBLigne : TOB;
    indice : integer;
begin
  // pas de controle si remonté et ligne vide
  if (GS.row < Arow) and (IsLigneVide (Arow))  then exit;

  if (stPrev <> GS.cells[Acol,Arow]) or (GS.cells[Acol,Arow]= '') then
  begin
    TOBLigne := TOBSaisies.detail[Arow-1];
    if ACol = G_CODEARTICLE then
    Begin
      if not FindCodeArticle (TOBLigne,GS.cells[Acol,Arow]) then BEGIN cancel := true; exit; END;
    end;
    if Acol = G_TYPEH then
    begin
       if not VerifTypeHeure(TOBLigne,GS.cells[Acol,Arow]) then BEGIN cancel := true; exit; END;
    end;
    if (Acol = G_LIBELLE) then
    begin
      TOBLIGNE.putValue('BSM_LIBELLE',GS.cells[Acol,Arow]);
    end;
    if (Acol >= G_QTE1) and (Acol <= G_QTE7) then
    begin
      Indice := (Acol-G_QTE1)+1;
      TOBLIGNE.putValue('BSM_QTEJ'+IntToStr(Indice),valeur(GS.cells[Acol,Arow]));
    end;
    if IsArticleExiste (Arow,Acol) Then
    BEGIN
      PgiBox('Cette prestation/frais est déjà présente dans la saisie');
      cancel := true;
    END;
    AfficheLigne (GS,TOBLIGNE,Arow-1);

    if not cancel then
    begin
      stPrev := GS.cells[Acol,Arow];
    end;
  end;
end;

function TOF_BTSAISIEHEBDO.VerifTypeHeure (TOBL: TOB; Valeur : string) : boolean;
begin
	result := true;
  valeur := trim (Uppercase(Valeur));
  if valeur <> '' then
  begin
    result := VerifExistsHeure(valeur);
    if not result then
    begin
      PgiBox (TraduireMemoire('Type d''heure inconnu'),ecran.caption);
      exit;
    end;
  end;
  TOBL.putValue('TYPEHEURE',valeur);
end;

function TOF_BTSAISIEHEBDO.VerifExistsHeure (valeur : string) : boolean;
var Sql : string;
    QQ : TQuery;
begin
  Sql := 'SELECT CC_LIBELLE,CC_ABREGE FROM CHOIXCOD WHERE CC_TYPE="ATH" AND CC_CODE="'+valeur+'"';
  QQ := OpenSql (SQl,true);
  result := not QQ.eof;
  ferme (QQ);
end;

procedure TOF_BTSAISIEHEBDO.GSElipsisClick(Sender: TOBject);
var Coords : Trect;
    PointD,PointF : Tpoint;
begin
  if GS.col=G_CODEARTICLE then
  begin
    if (Pos(fTypeRessource,'SAL;INT')= 0) then
    begin
      RechPrestations (self);
      exit;
    end;
//
    Coords := GS.CellRect (GS.col,GS.row);
    PointD := GS.ClientToScreen ( Coords.Topleft)  ;
    PointF := GS.ClientToScreen ( Coords.BottomRight )  ;
    POPCHOIX.Popup (pointF.X ,pointD.y+10);
  end else if GS.col = G_TYPEH then
  begin
     LookupList (THEdit(Sender),'Type d''heure','CHOIXCOD','CC_CODE','CC_LIBELLE','CC_TYPE="ATH"','',true,0);
  end;
end;

procedure TOF_BTSAISIEHEBDO.GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    case Key of
    VK_DELETE: if (Shift = [ssCtrl]) then
    begin
      DeleteLigne;
      Key := 0;
    end;
    VK_F5: if (Shift = []) then
    begin
      if TFVierge(Ecran).ActiveControl = GS then
      begin
        GSElipsisClick(Sender);
        Key := 0;
      end;
    end;
    VK_RETURN : if (Shift = []) then
    begin
    	Key := 0;
      SendMessage(THedit(Sender).Handle, WM_KeyDown, VK_TAB, 0);
    end;
  end;


end;

procedure TOF_BTSAISIEHEBDO.GetCellCanvas(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
var TypeComp : Tmodecomp;
begin
  if (Acol = 0) or (Arow = 0) then exit;
  if (Acol>=G_QTE1) and (Acol <= G_QTE7) then
  begin
    TypeComp := OKCompareSaisieConso(Acol,Arow);
    if (TypeComp=TmcNotEqual) then
    begin
      Canvas.Font.Color := clRed;
    end else if (TypeComp=TmcOnlyConso) then
    begin
      Canvas.Font.Color := clBlue;
      Canvas.Font.Style := Canvas.Font.Style + [fsBold];
    end else if (TypeComp=TmcOnlyModel) then
    begin
      Canvas.Font.Color := clGreen;
    end;
  end;
end;

procedure TOF_BTSAISIEHEBDO.GSPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
end;

procedure TOF_BTSAISIEHEBDO.GSRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
Var TobLigne : Tob;
begin
  if (Ou > 1) and (IsLigneVide (Ou-1)) then
  BEGIN
    cancel := true;
    Exit;
  END;
  if (Ou >= GS.rowCount -1) then GS.rowCount := GS.rowCount +1;
  if Ou > TOBSaisies.Detail.count then
  begin
    TOBLigne := AjouteLigneModele;
    AfficheLigne (GS,TOBLigne,Ou-1);
  end;
end;

procedure TOF_BTSAISIEHEBDO.GSRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
var Acol : integer;
Begin
  if IsLigneVide (Ou) then exit;
  if not ControleLigne (Ou,Acol) then
  BEGIN
  	cancel := true;
    PositionneDansGrid (Ou,Acol);
  END;
  DefiniCumulSemaine;
end;

procedure TOF_BTSAISIEHEBDO.ReinitSaisie;
var Indice : integer;
begin
  Indice := 0;
  repeat
    TOBSaisies.detail[Indice].ClearDetail;
    TOBSaisies.detail[Indice].PutValue ('CONSOJ1',0); // Cumul Lundi des consos
    TOBSaisies.detail[Indice].PutValue ('CONSOJ2',0);
    TOBSaisies.detail[Indice].PutValue ('CONSOJ3',0);
    TOBSaisies.detail[Indice].PutValue ('CONSOJ4',0);
    TOBSaisies.detail[Indice].PutValue ('CONSOJ5',0);
    TOBSaisies.detail[Indice].PutValue ('CONSOJ6',0);
    TOBSaisies.detail[Indice].PutValue ('CONSOJ7',0);
    //
    if TOBSaisies.detail[Indice].getValue('FROMCONSO')='X' then TOBSaisies.detail[Indice].Free else Inc(Indice);
  until indice >= TOBSaisies.detail.count;
end;


procedure TOF_BTSAISIEHEBDO.RetourEntete;
begin
  //
  ReinitSaisie;
  //
  GS.VidePile(false);
  SetGridEvent (false);
  GS.Visible := false;
  //
  SetModeSaisie (TmsEntete);
  RESSOURCE.SetFocus; // positionne sur la ressource
end;

procedure TOF_BTSAISIEHEBDO.SetparamGrille;
var LaListe,lelement : string;
begin
  // récupération du paramétrage général des grilles
  FNomTable := 'BTDETMODELSHEB';
  FLien := '';
  FSortBy := '';
  Fparams := '';
  stcols := 'BSM_NUMORDRE;BSM_CODEARTICLE;BSM_LIBELLE;TYPEHEURE;BSM_QTEJ1;BSM_QTEJ2;BSM_QTEJ3;BSM_QTEJ4;BSM_QTEJ5;BSM_QTEJ6;BSM_QTEJ7;';
  FTitre := 'N°;Prestation / Frais;libellé;Type;Lun.;Mar.;Mer.;Jeu.;Ven.;Sam.;Dim.;';
  FLargeur := '2;10;16;8;5;5;5;5;5;5;5;';
  Falignement := 'G.0  ---;G.0  ---;G.0  ---;G.0  ---;D.2  -X-;D.2  -X-;D.2  -X-;D.2  -X-;D.2  -X-;D.2  -X-;D.2  -X-;';
  title := 'Saisie des temps/Modèle';
  NC := '1;1;1;1;1;1;1;1;1;1;1;';
  Fperso := '---';
  OkTri := True;
  OkNumCOl := false;
  //
  stColListe := stCols;
  nbcolsInListe := 0;
  LaListe := stColListe;
  repeat
    lelement := READTOKENST (laListe);
    if lelement <> '' then
    begin
      if lelement = 'BSM_NUMORDRE' then G_NUMORDRE := nbcolsinListe else
      if lelement = 'BSM_CODEARTICLE' then G_CODEARTICLE := nbcolsinListe else
      if lelement = 'BSM_LIBELLE' then G_LIBELLE := nbcolsinListe else
      if lelement = 'TYPEHEURE' then  G_TYPEH := nbcolsinListe else
      if lelement = 'BSM_QTEJ1' then  G_QTE1 := nbcolsinListe else
      if lelement = 'BSM_QTEJ2' then  G_QTE2 := nbcolsinListe else
      if lelement = 'BSM_QTEJ3' then  G_QTE3 := nbcolsinListe else
      if lelement = 'BSM_QTEJ4' then  G_QTE4 := nbcolsinListe else
      if lelement = 'BSM_QTEJ5' then  G_QTE5 := nbcolsinListe else
      if lelement = 'BSM_QTEJ6' then  G_QTE6 := nbcolsinListe else
      if lelement = 'BSM_QTEJ7' then  G_QTE7 := nbcolsinListe ;
      inc(nbcolsInListe);
    end;
  until lelement = '';
end;


procedure TOF_BTSAISIEHEBDO.DefiniGrille;
var st,lestitres,lesalignements,FF,alignement,Nam,leslargeurs,lalargeur,letitre : string;
    Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep : boolean;
    dec : integer;
    indice : integer;
    FFLQTE : string;
begin

  GS.ColCount := NbColsInListe;

  st := stColListe;
  lesalignements := Falignement;
  lestitres := FTitre;
  leslargeurs := flargeur;

  for indice := 0 to nbcolsInListe -1 do
  begin
    Nam := ReadTokenSt (St); // nom
    alignement := ReadTokenSt(lesalignements);
    lalargeur := readtokenst(leslargeurs);
    letitre := readtokenst(lestitres);
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;

    GS.cells[Indice,0] := leTitre;

    if copy(Alignement,1,1)='G' then GS.ColAligns[indice] := taLeftJustify
    else if copy(Alignement,1,1)='D' then GS.ColAligns[indice] := taRightJustify
    else if copy(Alignement,1,1)='C' then GS.ColAligns[indice] := taCenter;

    GS.ColWidths[indice] := strtoint(lalargeur);

    if OkLib then GS.ColFormats[indice] := 'CB=' + Get_Join(Nam)
    else if (Dec<>0) or (Sep) then GS.ColFormats[indice] := FF ;

    if Copy(nam,1,7) = 'BSM_QTE' then
    begin
      if OkNulle then FFLQte := FFQTE+';-'+FFQTE+'; ;';
      GS.ColFormats[indice] := FFLQTE ;
    end;
  end ;
end;

procedure TOF_BTSAISIEHEBDO.AfficheLigne(Grid : THgrid;TOBL : TOB;indice : integer);
begin
  TOBL.PutLigneGrid (Grid,Indice+1,false,false,stColListe);
end;

procedure TOF_BTSAISIEHEBDO.AfficheLagrille;
var Indice : integer;
begin
  for Indice := 0 to TOBSaisies.detail.count -1 do
  begin
    AfficheLigne(GS,TOBSaisies.detail[Indice],Indice);
  end;
end;

function TOF_BTSAISIEHEBDO.VerifEntete: boolean;
begin
  result := true;
  if CH_CHANTIER.Text ='' then BEGIN PGIBox ('Renseignez le chantier...'); Result := false; Exit; END;
  if RESSOURCE.Text ='' then BEGIN PGIBox ('Renseignez la ressource...'); Result := false; Exit; END;
  if MODELE.Text ='' then BEGIN PGIBox ('Renseignez le modèle de saisie...'); Result := false; Exit; END;
  if GetChampsAffaire (CH_CHANTIER.Text,'AFF_ETATAFFAIRE')='TER' then
  begin
    PGIBox (TraduireMemoire('ATTENTION : Chantier clôturé'),ecran.caption);
    Result := false;
    exit;
  end;

end;

procedure TOF_BTSAISIEHEBDO.AffaireEfface(Sender: TObject);
begin
  PHASE.text := '';
  CH_CHANTIER.text := '';
  CH_CHANTIER1.text := '';
  CH_CHANTIER2.text := '';
  CH_CHANTIER3.text := '';
  CH_AVENANT.text := '';
  LCHANTIER.Caption := '';
  LPHASE.Caption := 'Non définie';
end;

procedure TOF_BTSAISIEHEBDO.bSelectAff1Click(Sender: TObject);
var iErr : integer;
begin
  inherited;
  CH_CHANTIER.Text := DechargeCleAffaire(CH_CHANTIER0, CH_CHANTIER1, CH_CHANTIER2, CH_CHANTIER3, CH_AVENANT, '', TaCreat, False, True, false, Ierr);
end;

procedure TOF_BTSAISIEHEBDO.ReInitModeleSaisie;
var Indice : integer;
    TOBL : TOB;
begin
  TOBSaisies.clearDetail;
  LoadModele (TOBSaisies,MODELE.text,fTypeRessource);
  for Indice := 0 to TOBSaisies.detail.count -1 do
  begin
    TOBL := TOBSaisies.detail[Indice];
    AddLesChampsSups(TOBSaisies.detail[Indice]);
    CopieValeurToModele (TOBL);
    TOBL.PutValue ('FROMMODELE','X');
    TOBL.PutValue ('TYPERESSOURCE',TOBL.GetValue ('BNP_TYPERESSOURCE'));
    TOBL.PutValue ('TYPEARTICLE',TOBL.GetValue ('GA_TYPEARTICLE'));
    TOBL.PutValue ('DPA',TOBL.GetValue ('GA_DPA'));
    TOBL.PutValue ('DPR',TOBL.GetValue ('GA_DPR'));
    TOBL.PutValue ('PVHT',TOBL.GetValue ('GA_PVHT'));
    TOBL.PutValue ('QUALIFUNITEMOUV',TOBL.GetValue ('GA_QUALIFUNITEVTE'));

    //
    if TOBL.GetValue('DPA') <> 0 then TOBL.PutValue ('COEFPAPR',TOBL.GetValue ('GA_DPR')/TOBL.GetValue ('GA_DPA'))
                                 else TOBL.PutValue ('COEFPAPR',1);
    if TOBL.GetValue('DPR') <> 0 then TOBL.PutValue ('COEFMARGE',arrondi(TOBL.GetValue ('GA_PVHT')/TOBL.GetValue ('GA_DPR'),4))
                                 else TOBL.PutValue ('COEFMARGE',1);
  end;
  TobSaisies_O.Dupliquer (TOBSaisies,true,true);
end;

procedure TOF_BTSAISIEHEBDO.AddLesChampsSups(TOBL: TOB);
begin
  TOBL.AddChampSupValeur ('TYPEARTICLE','');
  TOBL.AddChampSupValeur ('TYPERESSOURCE','');
  TOBL.AddChampSupValeur ('TYPEHEURE','');
  TOBL.AddChampSupValeur ('FROMMODELE','-');
  TOBL.AddChampSupValeur ('FROMCONSO','-');
  TOBL.AddChampSupValeur ('GETQTEFROM','MOD');
  TOBL.AddChampSupValeur ('MODIFIABLE','X');
  TOBL.AddChampSupValeur ('CONSOJ1',0); // Cumul Lundi des consos
  TOBL.AddChampSupValeur ('CONSOJ2',0);
  TOBL.AddChampSupValeur ('CONSOJ3',0);
  TOBL.AddChampSupValeur ('CONSOJ4',0);
  TOBL.AddChampSupValeur ('CONSOJ5',0);
  TOBL.AddChampSupValeur ('CONSOJ6',0);
  TOBL.AddChampSupValeur ('CONSOJ7',0);
  TOBL.AddChampSupValeur ('MODELJ1',0); // valeur du modele de lundi
  TOBL.AddChampSupValeur ('MODELJ2',0);
  TOBL.AddChampSupValeur ('MODELJ3',0);
  TOBL.AddChampSupValeur ('MODELJ4',0);
  TOBL.AddChampSupValeur ('MODELJ5',0);
  TOBL.AddChampSupValeur ('MODELJ6',0);
  TOBL.AddChampSupValeur ('MODELJ7',0);
  TOBL.AddChampSupValeur ('DPA',0);
  TOBL.AddChampSupValeur ('DPR',0);
  TOBL.AddChampSupValeur ('PVHT',0);
  TOBL.AddChampSupValeur ('COEFPAPR',0);
  TOBL.AddChampSupValeur ('COEFMARGE',0);
  TOBL.AddChampSupValeur ('QUALIFUNITEMOUV',0);
end;

procedure TOF_BTSAISIEHEBDO.CompleteSaisieViaConso;
var REq : string;
    TOBConso : TOB;
begin
  TOBConso := TOB.Create ('LES CONSO',nil,-1);
  Req := 'SELECT *,BNP_TYPERESSOURCE,BNP_LIBELLE,GA_TYPEARTICLE,GA_DPA,GA_DPR,GA_PVHT,GA_QUALIFUNITEVTE FROM CONSOMMATIONS '+
         'LEFT JOIN ARTICLE ON BCO_ARTICLE=GA_ARTICLE '+
         'LEFT JOIN NATUREPREST ON GA_NATUREPRES=BNP_NATUREPRES '+
         'WHERE '+
         '((BNP_TYPERESSOURCE="'+fTypeRessource+'" AND GA_TYPEARTICLE="PRE") OR (GA_TYPEARTICLE="FRA")) AND '+
         'BCO_AFFAIRE="'+CH_CHANTIER.Text+'" AND BCO_RESSOURCE="'+RESSOURCE.Text+'" AND '+
         'BCO_DATEMOUV >= "'+USDATETIME(Calendrier.SelectionStart)+'" '+
         'AND BCO_DATEMOUV <="'+USDATETIME(Calendrier.SelectionEnd)+'" ';
  if PHASE.Text <> '' then
  begin
    req := Req + 'AND BCO_PHASETRA="'+PHASE.text+'" ';
  end;
  Req := REq + 'ORDER BY BCO_DATEMOUV,BCO_ARTICLE,BCO_TYPEHEURE';
  TOBCONSO.LoadDetailDBFromSQL ('CONSOMMATIONS',Req,false);
  // Sauvegarde des consos pour suppression ulterieure
  TOBCONSO_O.dupliquer (TOBCONSO,true,true);
  //
  fExistsConso := (TOBCONSO.detail.count > 0);
  RattacheConsoSurModele (TOBConso);
  TOBCONSO.free;
end;

procedure TOF_BTSAISIEHEBDO.DefiniRowCounts;
begin
  if TOBSaisies.detail.count = 0 then AjouteLigneModele;
  if TOBSaisies.detail.count > 0 then GS.RowCount := TOBSaisies.detail.count +2;
  if GS.RowCount < 2 then GS.rowcount := 3;
end;

procedure TOF_BTSAISIEHEBDO.RattacheConsoSurModele(TOBConso: TOB);
var Indice : integer;
    TOBC,TOBAT : TOB;
begin
  if TOBCOnso.detail.count = 0 then exit;
  ReinitValSaisie;
  indice := 0;
  repeat
    TOBC := TOBConso.detail[Indice];
    TOBAT := TOBSaisies.FindFirst (['BSM_ARTICLE','TYPEHEURE'],[TOBC.GetValue('BCO_ARTICLE'),TOBC.GetValue('BCO_TYPEHEURE')],True);
    if TOBAT = Nil then
    begin
      TOBAT := AjouteLigneModeleFromConso (TOBC);
    end;
    EnregistreConsoSurModele (TOBAT,TOBC);
  until Indice >= TOBConso.detail.count;
end;

procedure AddLesLignesCumuls (TOBS : TOB);
var TOBL : TOB;
begin
  TOBL := TOB.Create ('CUMUL LUNDI',TOBS,-1); TOBL.AddChampSupValeur ('CUMUL',0);
  TOBL := TOB.Create ('CUMUL MARDI',TOBS,-1); TOBL.AddChampSupValeur ('CUMUL',0);
  TOBL := TOB.Create ('CUMUL MERCREDI',TOBS,-1); TOBL.AddChampSupValeur ('CUMUL',0);
  TOBL := TOB.Create ('CUMUL JEUDI',TOBS,-1); TOBL.AddChampSupValeur ('CUMUL',0);
  TOBL := TOB.Create ('CUMUL VENDREDI',TOBS,-1); TOBL.AddChampSupValeur ('CUMUL',0);
  TOBL := TOB.Create ('CUMUL SAMEDI',TOBS,-1); TOBL.AddChampSupValeur ('CUMUL',0);
  TOBL := TOB.Create ('CUMUL DIMANCHE',TOBS,-1); TOBL.AddChampSupValeur ('CUMUL',0);
end;

procedure TOF_BTSAISIEHEBDO.PrepareLeModele;
var Indice : integer;
begin
  for Indice := 0 to TOBSaisies.detail.count -1 do
  begin
    TOBSaisies.detail[Indice].clearDetail;
    AddLesLignesCumuls (TOBSaisies.detail[Indice]);
  end;
end;

function TOF_BTSAISIEHEBDO.AjouteLigneModeleFromConso(TOBC: TOB): TOB;
begin
  result := TOB.Create ('BTDETMODELSHEB',TOBSaisies,-1);
  AddLesChampsSups (result);
  result.PutValue ('TYPERESSOURCE',TOBC.GetValue ('BNP_TYPERESSOURCE'));
  result.PutValue ('TYPEARTICLE',TOBC.GetValue ('GA_TYPEARTICLE'));
  //
  result.PutValue ('DPA',TOBC.GetValue ('GA_DPA'));
  result.PutValue ('DPR',TOBC.GetValue ('GA_DPR'));
  result.PutValue ('PVHT',TOBC.GetValue ('GA_PVHT'));
  result.PutValue ('QUALIFUNITEMOUV',TOBC.GetValue ('GA_QUALIFUNITEVTE'));
  //
  if result.GetValue('DPA') <> 0 then result.PutValue ('COEFPAPR',result.GetValue ('DPR')/result.GetValue ('DPA'))
                                 else result.PutValue ('COEFPAPR',1);
  if result.GetValue('DPR') <> 0 then result.PutValue ('COEFMARGE',result.GetValue ('PVHT')/result.GetValue ('DPR'))
                                 else result.PutValue ('COEFMARGE',1);
  //
  result.putValue('BSM_TYPERESSOURCE',TOBSaisies.getValue('BMS_TYPERESSOURCE'));
  result.putValue('BSM_CODEMODELE',TOBSaisies.getValue('BMS_CODEMODELE'));
  result.putValue('BSM_NUMORDRE',TOBSaisies.detail.count);
  result.putValue('BSM_CODEARTICLE',TOBC.getValue('BCO_CODEARTICLE'));
  result.putValue('BSM_ARTICLE',TOBC.getValue('BCO_ARTICLE'));
  result.putValue('BSM_LIBELLE',TOBC.getValue('BCO_LIBELLE'));
  result.putValue('TYPEHEURE',TOBC.getValue('BCO_TYPEHEURE'));
  result.PutValue('FROMCONSO','X');
  result.PutValue('GETQTEFROM','CON');
  AddLesLignesCumuls (result);
end;

function TOF_BTSAISIEHEBDO.AjouteLigneModele: TOB;
begin
  result := TOB.Create ('BTDETMODELSHEB',TOBSaisies,-1);
  AddLesChampsSups (result);
  result.putValue('BSM_TYPERESSOURCE',TOBSaisies.getValue('BMS_TYPERESSOURCE'));
  result.putValue('BSM_CODEMODELE',TOBSaisies.getValue('BMS_CODEMODELE'));
  result.putValue('BSM_NUMORDRE',TOBSaisies.detail.count);
  result.PutValue('GETQTEFROM','MOD');
  AddLesLignesCumuls (result);
end;

function TOF_BTSAISIEHEBDO.GetIndiceTOB (LeJour : integer) : integer;
begin
  case LeJour of
    1 : result := 6; // Dimanche
    2 : result := 0; // Lundi
    3 : result := 1; // Mardi
    4 : result := 2; // Mercredi
    5 : result := 3; // Jeudi
    6 : result := 4; // Vendredi
    7 : result := 5; // Samedi
  else
    result := 1;
  end;
end;

procedure TOF_BTSAISIEHEBDO.EnregistreConsoSurModele(TOBAT, TOBC : TOB);
var LeJourdelaSemaine,IndiceTOB : integer;
    TOBLJ : TOB;
    LaZoneC,LaZoneS : string;
begin
  // On va enregistrer ici la ligne de conso sur la bonne ligne du modele et au jour indiqué
  LeJourdelaSemaine := DayOfWeek (TOBC.getValue('BCO_DATEMOUV'));
  IndiceTOB := GetIndiceTOB (LeJourdelaSemaine);
  TOBLJ := TOBAT.detail[IndiceTOB];
  TOBC.ChangeParent (TOBLJ,-1);
  TOBLJ.PutValue('CUMUL',TOBLJ.GetValue('CUMUL')+TOBC.GetValue('BCO_QUANTITE'));
  LaZoneC := 'CONSOJ'+IntToStr(IndiceTOB+1);
  TOBAT.PutValue(LaZoneC,TOBAT.GetValue(LaZoneC)+TOBC.GetValue('BCO_QUANTITE'));
  LaZoneS := 'BSM_QTEJ'+IntToStr(IndiceTOB+1);
  TOBAT.PutValue(LaZoneS,TOBC.GetValue('BCO_QUANTITE'));
end;

procedure TOF_BTSAISIEHEBDO.DefiniCumulSemaine;
var Indice,Jour : integer;
    LaZoneC,LaZoneM : string;
    TOBL : TOB;
begin
  CUMULSEMCON.Value := 0;
  CUMULSEMMOD.Value := 0;
  for Indice := 0 to TOBSaisies.detail.count -1 do
  begin
    TOBL := TOBSaisies.detail[Indice];
    for Jour := 1 to 7 do
    begin
      LaZoneM := 'BSM_QTEJ'+IntToStr(Jour);
      LaZoneC := 'CONSOJ'+IntToStr(Jour);
      if TOBL.GetValue('TYPERESSOURCE')='SAL' Then
      begin
        CUMULSEMMOD.Value := CUMULSEMMOD.Value + TOBL.GetValue(LaZoneM);
        CUMULSEMCON.Value := CUMULSEMCON.Value + TOBL.GetValue(LaZoneC);
      end;
    end;
  end;
end;

function TOF_BTSAISIEHEBDO.OKCompareSaisieConso(Acol,Arow : integer) : Tmodecomp ;
var Jour : integer;
    LaZoneC,LaZoneM,LaZoneS : string;
    TOBL : TOB;
begin
  result := TmcEqual;
  if (Arow = 0) Or (Arow > TOBSaisies.detail.count) then exit;
  TOBL := TOBSaisies.detail[Arow-1];
  Jour := (Acol - G_QTE1) + 1;
  LaZoneM := 'MODELJ'+IntToStr(Jour);
  LaZoneC := 'CONSOJ'+IntToStr(Jour);
  LaZoneS := 'BSM_QTEJ'+IntToStr(Jour);
  if (TOBL.getValue(LaZoneM)<>TOBL.getValue(LaZoneS)) then
  begin
    if (TOBL.getValue(LaZoneM)=0) then result := TmcOnlyConso
    else result := TmcNotEqual
  end else
  begin
    if (TOBL.getValue(LaZoneC)=0) then result := TmcOnlyModel
    else result := TmcEqual;
  end;
end;

procedure TOF_BTSAISIEHEBDO.PositionneDansGrid (Arow,Acol : integer);
var cancel : boolean;
    LastMode : boolean;
begin
  LastMode := assigned(GS.OnCellEnter);
  GS.CacheEdit;
  SetGridEvent (false);
  GS.row := Arow;
  GS.col := Acol;
  DefiniMenuDecid (Arow,Acol);
  ShowInfoQte (Acol,Arow);
  GS.ElipsisButton := (Acol = G_CODEARTICLE) OR (ACOl = G_TYPEH);
  stprev := GS.Cells [Acol,Arow];
  SetGridEvent (LastMode);
//  GS.OnRowEnter (self,Arow,Cancel,false);
//  GS.OnCellEnter (self,Acol,Arow,Cancel);
  GS.row := Arow;
  GS.col := Acol;
  GS.ShowEditor;
end;

function TOF_BTSAISIEHEBDO.IsLigneVide (Arow : integer) : boolean;
var TOBL : TOB;
begin
  result := true;
  if Arow > TOBSaisies.detail.count then exit;
  TOBL := TOBSaisies.detail[Arow-1];
  if TOBL.GetValue('BSM_CODEARTICLE') <> '' then result:= false;
end;


function TOF_BTSAISIEHEBDO.ControleLigne (ligne : integer; var Acol : integer) : boolean;
begin
  result := true;
  if IsLIgneVide (Ligne) then exit; // si elle est vide bon ..
  //
  if IsArticleExiste (Ligne,G_TYPEH) Then
  BEGIN
    ACol := G_QTE1;
    PgiBox('Cette prestation/frais est déjà présente dans la saisie');
    result := false;
  END;
  //
end;

procedure TOF_BTSAISIEHEBDO.ZoneSuivanteOuOk(Grille : THGrid;var ACol, ARow : integer; var Cancel : boolean);
var Sens, ii, Lim: integer;
  OldEna, ChgLig, ChgSens: boolean;
begin
  OldEna := Grille.SynEnabled;
  Grille.SynEnabled := False;
  Sens := -1;
  ChgLig := (Grille.Row <> ARow);
  ChgSens := False;
  if Grille.Row > ARow then Sens := 1 else if ((Grille.Row = ARow) and (ACol <= Grille.Col)) then Sens := 1;
  ACol := Grille.Col;
  ARow := Grille.Row;
  ii := 0;
  while not ZoneAccessible(Grille,ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if Sens = 1 then
    begin
      // Modif BTP
      Lim := Grille.RowCount ;
      // ---
      if ((ACol = Grille.ColCount - 1) and (ARow >= Lim)) then
      begin
        if ChgSens then Break else
        begin
          // Ajout d'une ligne
          break;
        end;
      end;
      if ChgLig then
      begin
        ACol := Grille.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < Grille.ColCount - 1 then
      begin
        Inc(ACol);
      end else
      begin
        Inc(ARow);
        ACol := Grille.FixedCols;
      end;
    end else
    begin
      if ((ACol = Grille.FixedCols) and (ARow = 1)) then
      begin
        if ChgSens then Break else
        begin
          Sens := 1;
          Continue;
        end;
      end;
      if ChgLig then
      begin
        ACol := Grille.ColCount;
        ChgLig := False;
      end;
      if ACol > Grille.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := Grille.ColCount - 1;
      end;
    end;
  end;
  Grille.SynEnabled := OldEna;
end;

function TOF_BTSAISIEHEBDO.ZoneAccessible( Grille : THgrid; var ACol, ARow : integer) : boolean;
var TOBL : TOB;
begin
  TOBL := nil;
  if Arow < TOBSaisies.detail.count then TOBL := TOBSaisies.detail[Arow-1];
  result := true;
  if Grille.ColWidths[acol] = 0 then BEGIN result := false; exit; END;
  if (Acol=G_CODEARTICLE) then
  BEGIN
    if (TOBL <> nil) and (TOBL.GetValue('BSM_ARTICLE')<>'') then
    begin
      result := false;
      exit;
    end;
  END;
  if (Acol>G_CODEARTICLE) then
  begin
    if (TOBL<>nil) and (TOBL.GetValue('BSM_ARTICLE')='') then
    begin
      result := false;
      exit;
    end;
  end;
  if (Acol = G_TYPEH) then
  begin
    if (TOBL <> nil) and (TOBL.GetValue('TYPERESSOURCE')<>'SAL') then
    begin
      result := false;
      exit;
    end;
  end;
end;

function TOF_BTSAISIEHEBDO.FindCodeArticle (TOBL : TOB; Valeur : string ) : boolean;
var QQ : TQuery;
    Req : STring;
begin
  if Valeur = '' then BEGIN result := false; Exit; END;
  Req := 'SELECT GA_CODEARTICLE,GA_TYPEARTICLE,GA_ARTICLE,GA_LIBELLE,BNP_TYPERESSOURCE,BNP_LIBELLE,GA_DPA,GA_DPR,GA_PVHT,GA_QUALIFUNITEVTE,GA_FOURNPRINC '+
         'FROM ARTICLE '+
         'LEFT JOIN NATUREPREST ON GA_NATUREPRES=BNP_NATUREPRES WHERE '+
         'GA_CODEARTICLE="'+Valeur+'" AND GA_STATUTART="UNI" '+
         'AND ((GA_TYPEARTICLE="PRE" AND BNP_TYPERESSOURCE="'+fTypeRessource+'") OR (GA_TYPEARTICLE="FRA"))';

  QQ := OpenSql (Req,True);
  result := not QQ.eof;
  if Result then
  begin
    TOBL.PutValue('BSM_LIBELLE',QQ.findField('GA_LIBELLE').asString);
    TOBL.PutValue('BSM_ARTICLE',QQ.findField('GA_ARTICLE').asString);
    TOBL.PutValue('BSM_CODEARTICLE',QQ.findField('GA_CODEARTICLE').asString);
    TOBL.PutValue('TYPERESSOURCE',QQ.findField('BNP_TYPERESSOURCE').asString);
    TOBL.PutValue('TYPEARTICLE',QQ.findField('GA_TYPEARTICLE').asString);
    TOBL.PutValue ('DPA',QQ.findField('GA_DPA').AsFloat);
    TOBL.PutValue ('DPR',QQ.findField('GA_DPR').AsFloat);
    TOBL.PutValue ('PVHT',QQ.findField('GA_PVHT').AsFloat);
    TOBL.PutValue ('QUALIFUNITEMOUV',QQ.findField('GA_QUALIFUNITEVTE').AsString);
    //
    if TOBL.GetValue('DPA') <> 0 then TOBL.PutValue ('COEFPAPR',TOBL.GetValue ('DPR')/TOBL.GetValue ('DPA'))
                                 else TOBL.PutValue ('COEFPAPR',1);
    if TOBL.GetValue('DPR') <> 0 then TOBL.PutValue ('COEFMARGE',arrondi(TOBL.GetValue ('PVHT')/TOBL.GetValue ('DPR'),4))
                                 else TOBL.PutValue ('COEFMARGE',1);
  //


  end else
  begin
    TOBL.InitValeurs;
    TOBL.putValue('BSM_NUMORDRE',GS.row);
  end;
  ferme (QQ);
end;

procedure TOF_BTSAISIEHEBDO.ShowInfoQte(Acol, Arow: integer);
var TOBL : TOB;
    Indice : integer;
begin
  PINDIC.Visible := False;
  TOBL := TOBSaisies.detail[Arow-1]; if TOBL = nil then exit;
  if (Acol >= G_QTE1) and (Acol <= G_QTE7) then
  begin
    Indice := (Acol - G_QTE1)+1;
    PINDIC.visible := true;
    QMODELE.Value := TOBL.GetValue('MODELJ'+IntToStr(Indice));
    QCONSO.Value := TOBL.GetValue('CONSOJ'+IntToStr(Indice));
  end else
  begin
    PINDIC.visible := false;
  end;
end;


Procedure TOF_BTSAISIEHEBDO.RechPrestations (Sender : TObject);
Var  stChamps : String;
     Article	: string;
     TOBL : TOB;
begin
  if GS.Cells [GS.Col,GS.row] <> '' then stchamps := 'GA_CODEARTICLE='+trim(GS.Cells[GS.col,GS.row])+';' else stchamps := '';
  stChamps := stChamps+'TYPERESSOURCE='+ fTypeRessource  ;
  stChamps := stChamps+';GA_TYPEARTICLE=PRE';
  //
  Article := AGLLanceFiche('BTP', 'BTPREST_RECH', '','',stChamps);
  //
  if Article <> '' then
  begin
    GS.cells[GS.col,GS.row] := Trim(Copy(Article,1,18));
    TOBL := TOBSaisies.detail[GS.row-1];
    if FindCodeArticle (TOBL,Trim(Copy(Article,1,18))) then
    begin
      AfficheLigne (GS,TOBL,GS.row-1);
      PositionneDansGrid (GS.row,G_LIBELLE);
      stPrev := GS.cells[GS.col,GS.row];
    end;
  end;
end;

Procedure TOF_BTSAISIEHEBDO.RechFRais (Sender : TObject);
Var  stChamps : String;
     Article	: string;
     TOBL : TOB;
begin
  if GS.Cells [GS.Col,GS.row] <> '' then stchamps := 'GA_CODEARTICLE='+trim(GS.Cells[GS.col,GS.row])+';' else stchamps := '';

  stChamps := stchamps+'XX_WHERE=AND (GA_TYPEARTICLE="FRA")';
  Article := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', StChamps);
  //
  if Article <> '' then
  begin
    GS.cells[GS.col,GS.row] := Trim(Copy(Article,1,18));
    TOBL := TOBSaisies.detail[GS.row-1];
    if FindCodeArticle (TOBL,Trim(Copy(Article,1,18))) then
    begin
      AfficheLigne (GS,TOBL,GS.row-1);
      PositionneDansGrid (GS.row,G_LIBELLE);
      stPrev := GS.cells[GS.col,GS.row];
    end;
  end;
end;

function TOF_BTSAISIEHEBDO.IsArticleExiste (Arow,Acol : integer) : boolean;
var TOBL : TOB;
begin
  result := false;
  TOBL := TOBSaisies.detail[Arow-1];
  if (((TOBL.GetValue('TYPEARTICLE')='PRE') and (TOBL.GetValue('TYPERESSOURCE')<>'SAL')) OR (TOBL.GetValue('TYPEARTICLE')='FRA')) and (Acol >= G_CODEARTICLE) then
  begin
    result := IsExisteFrais (TOBL);
  end else if (TOBL.GetValue('TYPEARTICLE')='PRE') and (TOBL.GetValue('TYPERESSOURCE')='SAL') and (Acol >= G_TYPEH) then
  begin
    result := IsExistePrestationInterne (TOBL);
  end;
end;

function TOF_BTSAISIEHEBDO.IsExisteFrais (TOBL : TOB) : boolean;
var Indice : integer;
    TOBC : TOB;
begin
  result := false;
  for Indice := 0 to TOBSaisies.detail.count -1 do
  begin
    TOBC := TOBSaisies.detail[Indice];
    if (TOBC.getValue('BSM_ARTICLE')=TOBL.GetValue('BSM_ARTICLE')) and (TOBL.getValue('BSM_NUMORDRE') <> TOBC.getValue('BSM_NUMORDRE')) then
    begin
      result := true;
      break;
    end;
  end;
end;

function TOF_BTSAISIEHEBDO.IsExistePrestationInterne (TOBL : TOB) : boolean;
var Indice : integer;
    TOBC : TOB;
begin
  result := false;
  for Indice := 0 to TOBSaisies.detail.count -1 do
  begin
    TOBC := TOBSaisies.detail[Indice];
    if (TOBC.getValue('BSM_ARTICLE')=TOBL.GetValue('BSM_ARTICLE')) and (TOBC.GetValue('TYPEHEURE')= TOBL.GetValue('TYPEHEURE')) and
       (TOBL.getValue('BSM_NUMORDRE') <> TOBC.getValue('BSM_NUMORDRE')) then
    begin
      result := true;
      break;
    end;
  end;
end;

procedure TOF_BTSAISIEHEBDO.DeleteLigne ;
var LastRow : integer;
begin
  if TFVierge(Ecran).ActiveControl = GS then
  begin
    lastrow := GS.row;
    GS.DeleteRow (GS.row);
    if GS.row <= TOBSaisies.detail.count then
    begin
      TOBSaisies.detail[GS.row-1].free;
    end;
    ReindiceGrid;
    Remplitgrille;
    if lastrow > TOBSaisies.detail.count then lastRow := TOBSaisies.detail.count;
    PositionneDansGrid (LastRow,1);
  end;
end;

procedure TOF_BTSAISIEHEBDO.ReindiceGrid;
var Indice : integer;
    TOBL : TOB;
begin
  for Indice := 0 to TOBSaisies.detail.count-1 do
  begin
    TOBL := TOBSaisies.detail[Indice]; if TOBL = nil then break;
    TOBL.putValue('BSM_NUMORDRE',Indice+1);
  end;
end;

procedure TOF_BTSAISIEHEBDO.Remplitgrille;
begin
  GS.VidePile (false);
  DefiniRowCounts;
  AfficheLagrille;
  TFVierge(ecran).HMTrad.ResizeGridColumns (GS);
  DefiniCumulSemaine;
end;

function TOF_BTSAISIEHEBDO.ExisteLienRessourceAffaire : boolean;
var QQ : TQuery;
begin
  QQ := OpenSql ('SELECT AFT_LIBELLE FROM AFFTIERS WHERE AFT_AFFAIRE="'+CH_CHANTIER.text+'" AND AFT_RESSOURCE="'+RESSOURCE.text+'" AND AFT_TYPEINTERV="RES"',True);
  result := not QQ.eof;
  Ferme (QQ);
end;

procedure TOF_BTSAISIEHEBDO.MemoriseResource;
var TOBAR : TOB;
    QQ : Tquery;
    Imax : integer;
begin
  if not ExisteLienRessourceAffaire then
  begin

    QQ := OpenSQL ('SELECT MAX(AFT_RANG) FROM AFFTIERS WHERE AFT_AFFAIRE="' + CH_CHANTIER.text + '"', TRUE) ;
    if not QQ.EOF then
      imax := QQ.Fields [0].AsInteger + 1
    else
      iMax := 1;
    Ferme (QQ) ;

    TOBAR := TOB.Create ('AFFTIERS',nil,-1);
    TOBAR.putValue('AFT_AFFAIRE',CH_CHANTIER.text);
    TOBAR.putValue('AFT_RESSOURCE',RESSOURCE.text);
    TOBAR.putValue('AFT_LIBELLE',LRESSOURCE.caption);
    TOBAR.putValue('AFT_CREATEUR','CEG');
    TOBAR.putValue('AFT_DATECREATION',Now);
    TOBAR.putValue('AFT_TYPEINTERV','RES');
    TOBAR.putValue('AFT_RANG',Imax);
    TOBAR.insertDb (nil);
    TOBAR.free;
  end;
end;

procedure TOF_BTSAISIEHEBDO.DefiniMenuDecid (Arow,Acol : integer);
var TypeComp : Tmodecomp;
begin

  MnCelluleMod  := TMenuItem(GetControl('MnCelluleMod'));
  MnCelluleCon  := TMenuItem(GetControl('MnCelluleCon'));
  MnApplicAllMod  := TMenuItem(GetControl('MnApplicAllMod'));
  MnApplicAllCon := TMenuItem(GetControl('MnApplicAllCon'));


  if Acol >= G_QTE1 then
  begin
    TypeComp := OKCompareSaisieConso(Acol,Arow);
    if (TypeComp = TmcEqual) or (TypeComp = TmcOnlyModel) then
    begin
      MnCelluleMod.enabled := false;
      MnCelluleCon.Enabled := false;
      MnApplicAllMod.Enabled := false;
      MnApplicAllCon.Enabled := false;
    end else if TypeCOmp = TmcOnlyConso then
    begin
      MnCelluleMod.enabled := false;
      MnCelluleCon.Enabled := true;
      MnApplicAllMod.Enabled := false;
      MnApplicAllCon.Enabled := true;
    end else if TypeCOmp = TmcNotEqual then
    begin
      MnCelluleMod.enabled := true;
      MnCelluleCon.Enabled := true;
      MnApplicAllMod.Enabled := true;
      MnApplicAllCon.Enabled := true;
    end;

  end else
  begin
    MnCelluleMod.enabled := false;
    MnCelluleCon.Enabled := false;
    MnApplicAllMod.Enabled := true;
    MnApplicAllCon.Enabled := fExistsConso;
  end;

end;

procedure TOF_BTSAISIEHEBDO.CopieValeurToModele (TOBL : TOB);
var Indice : integer;
    ChampM,ChampS : string;
begin
  for Indice := 1 to 7 do
  begin
    ChampM := 'MODELJ'+IntToStr(Indice);
    ChampS := 'BSM_QTEJ'+IntToStr(Indice);
    TOBL.putValue(ChampM,TOBL.GetValue(ChampS));
  end;
end;

procedure TOF_BTSAISIEHEBDO.CopieConsoToSaisie (TOBL : TOB);
var Jour : integer;
    ChampC,ChampS : string;
begin
  for Jour := 1 to 7 do
  begin
    ChampC := 'CONSOJ'+IntToStr(Jour);
    ChampS := 'BSM_QTEJ'+IntToStr(Jour);
    TOBL.putValue(Champs,TOBL.GetValue(ChampC));
  end;
end;


procedure TOF_BTSAISIEHEBDO.CopieModeleToSaisie (TOBL : TOB);
var Jour : integer;
    ChampC,ChampS : string;
begin
  for Jour := 1 to 7 do
  begin
    ChampC := 'MODELJ'+IntToStr(Jour);
    ChampS := 'BSM_QTEJ'+IntToStr(Jour);
    TOBL.putValue(Champs,TOBL.GetValue(ChampC));
  end;
end;

procedure TOF_BTSAISIEHEBDO.MnApplicAllConClick(Sender: Tobject);
var Indice : integer;
    TOBS : TOB;
    Arow,Acol : integer;
begin
  Arow := GS.row;
  ACol := GS.col;
  SetGridEvent (false);
  GS.VidePile(false);
  DefiniRowCounts ;
//
  for Indice := 0 To TOBSaisies.detail.count -1 do
  begin
    TOBS := TOBSaisies.detail[Indice];
    CopieConsoToSaisie (TOBS);
    AfficheLigne (GS,TOBS,Indice);
  end;
  SetGridEvent (True);
  PositionneDansGrid (Arow,Acol);
end;

procedure TOF_BTSAISIEHEBDO.MnApplicAllModClick(Sender: Tobject);
var Indice : integer;
    TOBS : TOB;
    Arow,Acol : integer;
begin
  Arow := GS.row;
  ACol := GS.col;
  SetGridEvent (false);
  GS.VidePile(false);
  RecupModeleorigine;
  DefiniRowCounts ;
//
  for Indice := 0 To TOBSaisies.detail.count -1 do
  begin
    TOBS := TOBSaisies.detail[Indice];
    AfficheLigne (GS,TOBS,Indice);
  end;
  SetGridEvent (True);
  PositionneDansGrid (Arow,Acol);
end;

procedure TOF_BTSAISIEHEBDO.MnCelluleConClick(Sender: Tobject);
var TOBC : TOB;
    JOUR : INTEGER;
    ChampC,ChampS : string;
begin
  if GS.row > TOBSaisies.detail.count then exit;
  TOBC := TOBSaisies.detail[GS.row-1]; if TOBC = nil then exit;
  jour := (GS.col - G_QTE1) + 1;
  ChampC := 'CONSOJ'+IntToStr(Jour);
  ChampS := 'BSM_QTEJ'+IntToStr(Jour);
  TOBC.putValue(Champs,TOBC.GetValue(ChampC));
  AfficheLigne (GS,TOBC,GS.row-1);
end;

procedure TOF_BTSAISIEHEBDO.MnCelluleModClick(Sender: Tobject);
var TOBC : TOB;
    JOUR : INTEGER;
    ChampC,ChampS : string;
begin
  if GS.row > TOBSaisies.detail.count then exit;
  TOBC := TOBSaisies.detail[GS.row-1]; if TOBC = nil then exit;
  jour := (GS.col - G_QTE1) + 1;
  ChampC := 'MODELJ'+IntToStr(Jour);
  ChampS := 'BSM_QTEJ'+IntToStr(Jour);
  TOBC.putValue(Champs,TOBC.GetValue(ChampC));
  AfficheLigne (GS,TOBC,GS.row-1);
end;

procedure TOF_BTSAISIEHEBDO.SupprimelesConsos (Sender : TObject);
begin
  if PGIAsk ('Etes-vous sûr de vouloir supprimer les saisies ?')=MrYes then
  begin
    TOBCOnso_o.DeleteDB;
    RetourEntete;
  end;
end;

procedure TOF_BTSAISIEHEBDO.EnregistreLesConsos;
begin
  if TOBConso_O.detail.count > 0 then
  begin
    if not TOBCOnso_o.DeleteDB Then V_PGI.ioerror := OeUnknown;
  end;
  if V_PGI.IOError = OeOk then EcritLesConsos;
end;

procedure TOF_BTSAISIEHEBDO.NettoieMoiCa (TOBCC : TOB);
var Indice : integer;
    TOBLD : TOB;
begin
  Indice := 1;
  repeat
    TOBLD := TOBCC.detail[Indice];
    TOBLD.free;
  until Indice >= TOBCC.detail.count;
end;

procedure TOF_BTSAISIEHEBDO.CreeLigneConso (TOBS,TOBCC : TOB; Jour : TDateTime);
Var Part      : String;
    Part0     : String;
    Part1     : String;
    Part2     : String;
    Part3     : String;
    Part4     : String;
    Day,year,Month   :word;
    TypeArticle : String;
    TheRetour : TGncERROR;
    Nature : string;
    NumMouv   : Double;
    TOBC : TOB;
begin
  Part  := CH_CHANTIER.Text;
  Part0 := '';
  Part1 := '';
  Part2 := '';
  Part3 := '';
  Part4 := CH_AVENANT.text;
  BTPCodeAffaireDecoupe(Part,Part0,Part1,Part2,Part3, Part4, TaModif,false);
  TypeArticle := TOBS.GetValue('TYPEARTICLE');
  Nature := TOBS.GetValue('TYPERESSOURCE');
//
  TOBC := Tob.Create('CONSOMMATIONS', TOBCC, -1);
  TheRetour := GetNumUniqueConso (NumMouv);
  if TheRetour = gncAbort then
  BEGIN
    PGIBox ('Erreur lors de l''attribution d''un numéro de consommations');
    V_PGI.IOError := oePointage;
    Exit;
  END;
//
  TOBC.AddChampSupValeur ('COEFMARGE',TOBS.GetValue('COEFMARGE'));
  // Par défaut positionné facturable à non
  TOBC.PutValue ('BCO_FACTURABLE','N');
  TOBC.putValue ('BCO_AFFAIRESAISIE','');
  TOBC.putValue ('BCO_QTEFACTUREE',1);
  TOBC.PutValue ('BCO_TRANSFORME', '-');
  TOBC.PutValue ('BCO_NUMMOUV', NumMouv);
  TOBC.PutValue ('BCO_INDICE',0);

  //
  TOBC.PutValue ('BCO_RESSOURCE',RESSOURCE.Text);

  DecodeDate (Jour,year,month,day);
  TOBC.putValue ('JOUR',IntToStr (day));
  TOBC.putValue ('BCO_SEMAINE',NumSemaine(Jour));
  TOBC.PutValue ('BCO_MOIS',month);
  //
  TOBC.PutValue ('BCO_AFFAIRE', CH_CHANTIER.text);
  TOBC.PutValue ('BCO_AFFAIRE0', Part0);
  TOBC.PutValue ('BCO_AFFAIRE1', Part1);
  TOBC.PutValue ('BCO_AFFAIRE2', Part2);
  TOBC.PutValue ('BCO_AFFAIRE3', Part3);
  TOBC.PutValue ('BCO_PHASETRA', PHASE.text);
  TOBC.PutValue ('BCO_DATEMOUV', Jour);
 //
  TOBC.PutValue ('BCO_CODEARTICLE',TOBS.GetValue('BSM_CODEARTICLE'));
  TOBC.PutValue ('BCO_ARTICLE',TOBS.GetValue('BSM_ARTICLE') );
  TOBC.PutValue ('BCO_LIBELLE',TOBS.GetValue('BSM_LIBELLE') );
  TOBC.putValue ('BCO_QUALIFQTEMOUV',TOBS.GetValue('QUALIFUNITEMOUV'));
  TOBC.putValue ('BCO_TYPEHEURE',TOBS.GetValue('TYPEHEURE'));
  //
  if TypeArticle = 'PRE' then
  begin
    if Nature = 'SAL' then
    begin
      TOBC.PutValue('BCO_NATUREMOUV', 'MO');
      if ValoPrestaFromRessource then
      begin
        SetValoFromRessource (TOBC,TOBRessource);
      end else
      begin
      	SetvaloFromArticle (TOBC);
      end;
      SetValoFromTypeHeures (TOBC);
    end else if  Nature = 'ST' then
    begin
      TOBC.PutValue('BCO_NATUREMOUV', 'EXT');
      SetvaloFromArticle (TOBC);
    end else if  Nature = 'INT' then
    begin
      TOBC.PutValue('BCO_NATUREMOUV', 'EXT');
      if ValoPrestaFromRessource then
      begin
        SetValoFromRessource (TOBC,TOBRessource);
      end else
      begin
      	SetvaloFromArticle (TOBC);
      end;
      SetValoFromTypeHeures (TOBC);
    end else if  Nature = 'AUT' then
    begin
      TOBC.PutValue('BCO_NATUREMOUV', 'EXT');
      SetvaloFromArticle (TOBC);
    end else if  Nature = 'LOC' then
    begin
      TOBC.PutValue('BCO_NATUREMOUV', 'EXT');
      SetvaloFromArticle (TOBC);
    end else if  Nature = 'MAT' then
    begin
      TOBC.PutValue('BCO_NATUREMOUV', 'RES');
      SetvaloFromArticle (TOBC);
    end else if  Nature = 'OUT' then
    begin
      TOBC.PutValue('BCO_NATUREMOUV', 'RES');
      if ValoPrestaFromRessource then
      begin
      	SetValoFromRessource (TOBC,TOBRessource);
      end else
      begin
      	SetvaloFromArticle (TOBC);
      end;
    end;
  end else if (TypeArticle = 'MAR') or (TypeArticle = 'ARP') then
  Begin
    TOBC.PutValue('BCO_NATUREMOUV', 'FOU');
    SetvaloFromArticle (TOBC);
  end else if TypeArticle = 'FRA' then
  Begin
    TOBC.PutValue('BCO_NATUREMOUV', 'FRS');
    SetvaloFromArticle (TOBC);
  end;
end;

procedure TOF_BTSAISIEHEBDO.AppliqueQteSaisieSurConso (TOBCONSO,TOBS : TOB ; Jour : TDateTime; IndiceJour : integer; QteJ : double);
var TOBCC : TOB;
begin
  TOBCC := TOBS.detail[IndiceJour]; // la ligne de centralisation des lignes de conso.
  if TOBCC.detail.count > 1 then NettoieMoiCa (TOBCC);
  if TOBCC.detail.count = 0 then CreeLigneConso (TOBS,TOBCC,Jour);
  MetJourLaConso (TOBS,TOBCC.detail[0],QteJ); // il n'en restera qu'une ^^
  TOBCC.detail[0].ChangeParent (TOBCONSO,-1);
end;

procedure TOF_BTSAISIEHEBDO.prepareEcritureConsos (TOBConsos : TOB);
var Indice,IndiceJour : integer;
    TOBS : TOB;
    QteJ : double;
    Jour : TDateTime;
//    ChampM : string;
begin
  for Indice := 0 to TOBSaisies.detail.count -1 do
  begin
    TOBS := TOBSaisies.detail[Indice];
    for IndiceJour := 1 To 7 do
    begin
      QteJ := TOBS.GetValue('BSM_QTEJ'+IntToStr(IndiceJour));
(*
      ChampM := 'BSM_QTEJ'+IntToStr(IndiceJour);
      TOBS.PutValue (ChampM,QteJ); // stocke les valeurs saisies dans le nouveau modèle actif
*)
      if QTeJ > 0 then
      begin
        if IndiceJour > 1 then JOur := PlusDate  (Calendrier.SelectionStart,indiceJour-1,'J')
                          else JOur := Calendrier.SelectionStart;
        AppliqueQteSaisieSurConso (TOBCONSOS,TOBS,Jour,IndiceJour-1,QteJ);
      end;
    end;
  end;
end;

procedure TOF_BTSAISIEHEBDO.EcritLesConsos;
var TOBConsos : TOB;
begin
  TOBConsos := TOB.Create ('LES CONSOS',nil,-1);
  prepareEcritureConsos (TOBConsos);
  if TOBConsos.detail.count > 0 then
  begin
    if not TOBConsos.InsertDB (nil,true) then V_PGI.Ioerror := OeUnknown; 
  end;
end;

procedure TOF_BTSAISIEHEBDO.RecupLesInfosRessource;
var SQl : String;
    QQ : TQuery;
begin
  TOBRessource.InitValeurs;
  SQL := 'SELECT * FROM RESSOURCE WHERE ARS_RESSOURCE="'+RESSOURCE.Text+'"';
  QQ := OPenSql (SQL,True);
  TOBRessource.SelectDB ('',QQ);
  ferme (QQ);
end;


procedure TOF_BTSAISIEHEBDO.SetvaloFromArticle (TOBL : TOB);
var fournisseur,Article,Ua : string;
    TOBA : TOB;
    QQ : TQuery;
    CoefUaUs : double;
begin
  TOBA := TOB.Create ('ARTICLE',nil,-1);
  QQ := OPENSQL ('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+TOBL.GEtValue('BCO_ARTICLE')+'"',True);
  TOBA.selectdb ('',QQ);
  ferme (QQ);
	Fournisseur := TOBA.GetValue('GA_FOURNPRINC');
  Article := TOBA.GetValue('GA_ARTICLE');
  RecupTarifAch (Fournisseur,Article,Ua,CoefUaUs,TurVente,true,True,TOBA);
  ReactualisePrPv (TOBA);
  TOBL.putValue('BCO_DPA',TOBA.GetValue('GA_PAHT') );
  TOBL.putValue('BCO_DPR',TOBA.GetValue('GA_DPR') );
  TOBL.putValue('BCO_PUHT',TOBA.GetValue('GA_PVHT'));
  TOBA.Free;
end;

procedure TOF_BTSAISIEHEBDO.MetJourLaConso (TOBS,TOBL : TOB; QteJ : double);
begin
  TOBL.PutValue('BCO_QUANTITE',QteJ);
  TOBL.PutValue('BCO_LIBELLE',TOBS.GetValue('BSM_LIBELLE'));
  calculeLaLigne (TOBL);
end;

procedure TOF_BTSAISIEHEBDO.RecupModeleorigine;
var Indice,Jour : integer;
    TOBSO,TOBAT : TOB;
    ChampS,ChampC : string;
begin
  for Indice := 0 to TOBSaisies_O.detail.count -1 do
  begin
    TOBSO := TOBSaisies_O.detail[Indice];
    TOBAT := TOBSaisies.FindFirst (['BSM_ARTICLE','TYPEHEURE'],[TOBSO.GetValue('BSM_ARTICLE'),TOBSO.GetValue('TYPEHEURE')],True);
    if TOBAT = Nil then
    begin
      TOBAT := TOB.Create ('BTDETMODELSHEB',TOBSaisies,-1);
      AddLesChampsSups (TOBAT);
      TOBAT.Dupliquer (TOBSO,true,true);
      TOBAT.putValue('BSM_NUMORDRE',TOBSaisies.detail.count);
    end else
    begin
      for Jour := 1 to 7 do
      begin
        ChampC := 'MODELJ'+IntToStr(Jour);
        ChampS := 'BSM_QTEJ'+IntToStr(Jour);
        TOBAT.putValue(Champs,TOBSO.GetValue(ChampC));
      end;
    end;
  end;
end;

procedure TOF_BTSAISIEHEBDO.ReinitValSaisie;
var Indice,Jour : integer;
    TOBSO : TOB;
    ChampS : string;
begin
  for Indice := 0 to TOBSaisies.detail.count -1 do
  begin
    TOBSO := TOBSaisies.detail[Indice];
    for Jour := 1 to 7 do
    begin
      ChampS := 'BSM_QTEJ'+IntToStr(Jour);
      TOBSO.putValue(Champs,0);
    end;
  end;
end;

procedure TOF_BTSAISIEHEBDO.SetValoFromTypeHeures (TOBS : TOB);
var Sql : string;
    QQ : TQuery;
    LaTOB : TOB;
    TheValeur : string;
    LaValeur : string;
    coef : double;
begin
  TheValeur := TOBS.getValue('BCO_TYPEHEURE');
  if TheValeur = '' then exit;
  laTOB := TOBTypeHeure.findFirst(['CC_CODE'],[TheValeur],True);
  if LaTOB = nil then
  begin
    Sql := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="ATH" AND CC_CODE="'+TheValeur+'"';
    QQ := OpenSql (SQl,true);
    if not QQ.eof then
    begin
      LaTOB := TOB.Create ('CHOIXCOD',TOBTypeheure,-1);
      LaTOB.selectdb ('',QQ);
    end;
    ferme (QQ);
  end;
  if LaTOB <> nil then
  begin
    LaValeur :=  LaTOB.GetValue ('CC_ABREGE');
    if (LaValeur <> '') and (IsNumeric (LaValeur)) then
    begin
      Coef := 1 + VALEUR(laValeur)/100;
	    TOBS.PutValue('BCO_DPA',Arrondi(TOBS.GetValue('BCO_DPA')*Coef,V_PGI.okdecV));
 	    TOBS.PutValue('BCO_DPR',Arrondi(TOBS.GetValue('BCO_DPR')*Coef,V_PGI.okdecV));
 	    TOBS.PutValue('BCO_PUHT',Arrondi(TOBS.GetValue('BCO_PUHT')*Coef,V_PGI.okdecV));
    end;
  end;
end;


procedure TOF_BTSAISIEHEBDO.RemplitTypeHeures;
var Sql : string;
begin
  Sql := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="ATH"';
  TOBTypeHeure.LoadDetailDBFromSQL ('CHOIXCOD',Sql,false);
end;

Initialization
  registerclasses ( [ TOF_BTSAISIEHEBDO ] ) ;
end.


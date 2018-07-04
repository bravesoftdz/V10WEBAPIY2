unit UTomRessource;

interface
uses {$IFDEF VER150} variants,{$ENDIF} Windows,
     StdCtrls,
     Controls,
     Classes,
     forms,
     sysutils,
     ComCtrls,
     ParamSoc,
{$IFDEF EAGLCLIENT}
  	 Maineagl,
     eFiche,
     eFichList,
     utileAGL,
     Spin,
{$ELSE}
  	 FE_Main,
     Fiche,
     FichList,
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     dbctrls,
     HDB,
{$ENDIF}
     UtilConfid,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     Dialogs,
     TiersUtil,
     EntGC,
     LookUp,
  	 Dicobtp,
     HTB97,
     UtilPaieAffaire,
     confidentaffaire,
     menus,
  //activiteutil,    source intuil. si utiliser par GPAO ou BTP à mettre en ifdef
     utilarticle,
     UtilPGI,
     UtilGc,
     UTOB,
     UTOF,
     AGLInit,
     uafo_ressource,
     UtomRessourcePr,
     UtilRessource,
     UtomCompetence, // à laisser pour être sur que ce source utilis2 dans le script soit dans le projet
     {$IFDEF AFFAIRE} // la revalo de l'activité n'est accessible que si Affaire ...
     UTOFREVAL_RESS_MUL, UtofResHistoValo,AffaireUtil,
     {$ENDIF}
     AGLInitGc,
     M3fp,
     graphics,
     //
     CodePost,
     AglMail,
     MailOl;

type
  TOM_Ressource = class(TOM)
    procedure OnNewRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnArgument(sArgument: string); override;
    procedure OnDeleteRecord; override;
    procedure SetArguments(StSQL: string);
    procedure OnClose; override;
    procedure OnCancelRecord             ; override ;
  public
    tsArgsReception	: TStringList;
     //NCX 15/05/01
  private
    TOBRESSOURCEBTP : TOB;
    BREVALORISATION	: TToolbarButton97;
    BHISTOVALO			: TToolbarButton97;
    BTMAIL					: TToolbarButton97;
    BTPERSOCALEND  	: TToolBarButton97;
    BTSALARIE				: TToolBarButton97;
{$IFDEF GPAO}
    BRESSOURCELIE		: tToolBarButton97;
{$ENDIF}
    TempSalarie			: string;
    TypeRessource		: string;
    StSQL						: string;
    //
    PremierPassage	: Boolean;
    QuestHisto			: boolean;
    ModifLot				: boolean;
    bEstFamille			: Boolean;
    CacherPrix	 		: Boolean;
    //
    //Variable ajouté par FV pour gestion des calendrier ressources
    CBOSTANDCALEND 	: THValComboBox;
    CBOTYPERESSOURCE: THDBValComboBox;
    CHKCALENSPECIF 	: TCheckBox;
    CHKMENSUALITE  	: TCheckBox;
    CHKCALCULPR			: TCheckBox;
    //
    AdresseMail			: THDBEdit;
    CodeArticle			: THDBEdit;
    CodePostal			: THDBEdit;
    Ville						: THDBEdit;
    TLabelRessource : THLabel;
    TLabelType			: THLabel;
    //
    //FV1 : 15/12/20016 - FS#2230 - TEAM RESEAUX - fiche ressource aucun contrôle pour la section analytique
    SectionAna      : THDBEdit;

{$IFDEF AFFAIRE}
    procedure BRevalorisationClick(sender: Tobject);
{$ENDIF}
    procedure BHistoValoclick(sender: TObject);
{$IFDEF GPAO}
    function IsEcranGPAO: Boolean;
    procedure BRESSOURCELIE_OnClick(Sender: tObject);
{$ENDIF}
    procedure CBSpecifClick(Sender: TObject);
    procedure StandCalenClick(Sender: TObject);
    //procedure BValiderClick (Sender : TObject);
    procedure Historisation;
    procedure MajHistovalo(parms: string);
    procedure ZoneSalarieOnExit(Sender: TObject);
    procedure ZoneSalarieOnEnter(Sender: TObject);
    function ChargeTheTobHisto(liste: string; Historisation: boolean): string;
    procedure MajEnregCal(quoi: string);
    function VerifChangements: integer;
    procedure Duplication(CleDuplic: string);
    //function VerfiCalSpecifExist : boolean;
    procedure CheckRattachementSalarie;

    //Procedures ajoutées par FV pour gestion de la Sous-Famille
    procedure LectureSFamilleRes(SFamRes: string);
    procedure SsFamilleOnElipsisClick(Sender: TObject);
    procedure SsFamilleOnExit(Sender: TObject);
    procedure BtPersoCalendClick(Sender: TObject);

    //Procedures Ajoutées pour gestion dans TOM du Script de la fiche !!!
    procedure ChkCalculPR_OnClick(Sender: TObject);
    Procedure MailClick(Sender: TObject);
    Procedure BtSalarieClick(Sender: TObject);
    Procedure ControleAdresse(Sender: TObject);

    //Procedures ajoutées par FV pour gestion Line
    procedure GestionFicheRessource(StArgument: String);
    Procedure GestionFicheMateriel(StArgument: String);
    procedure GestionZoneEcran;
    procedure LoadFicheRessource;
    Procedure LoadFicheMateriel;
    procedure MailChange(Sender: TObject);
    procedure MailDblClick(Sender: TObject);
    //FV1 : 15/12/20016 - FS#2230 - TEAM RESEAUX - fiche ressource aucun contrôle pour la section analytique
    procedure FindRessourceBTP(CodeRessource: string);
    procedure SetEvents(Etat: boolean);
    procedure ChangeInfoSUp(Sender: Tobject);
    procedure DeleteRessourceBTP(CodeRessource: string);

  protected
    EnSaisie, ModifSpecif, ModifStand, ModifSalarie: Boolean;
    Oldsalarie: string;
    function Ressource_CalculPR(): double;
    function Ressource_CalculPV(): double; //NCX 18/05/01
    function TestImpactChangeSalarie(CodeSalarie, CodeRessource: string): Boolean;
    procedure MoveFonction(Sender: TObject);
    procedure VerifFonction(Sender: TObject);
{$IFDEF GPAO}
    procedure BImprimer_OnClick(Sender: tObject);
{$ENDIF}
  end;

const
  // libellés des messages de la TOM Ressource
  TexteMsgRessource: array[1..32] of string = (
    {1}'Vous devez renseigner un code ressource',
    {2}'Vous devez renseigner une Famille article',
    {3}'Date incorrecte veuillez la modifier',
    {4}'Code Salarié invalide',
    {5}'Changement impossible : compétences par ressource/salarié existantes',
    {6}'Changement impossible : calendrier par ressource/salarié existant',
    {7}'Changement impossible : absences par ressource/salarié existantes',
    {8}'Date de fonction de l''historique supérieure à la fonction actuelle',
    {9}'Date de fonction actuelle inférieure à celle de l''historique ',
    {10}'Suppression impossible : il existe des lignes de consommations pour la ressource',
    {11}'Suppression impossible : la ressource intervient sur des affaires',
    {12}'Suppression impossible : la ressource est utilisée en historique d''activité',
    {13}'Suppression impossible : la ressource est affectée sur des lignes de pièce',
    {14}'saisi n''existe pas',
    {15}'Le code prestation saisi n''existe pas',
    {16}'Le code article saisi n''existe pas',
    {17}'Vous devez renseigner un département d''affaire',
    {18}'Code salarié déjà existant',
    {19}' ',
    { Merci de ne pas modifier ces numéros de messages, ou mettre à jour wRessource }
    {20}'Suppression impossible : La ressource figure dans les ressources d''une opération',
    {21}'Suppression impossible : La ressource est associée à une autre ressource'
    {22}, 'La saisie du champs suivant est obligatoire : '
    {23}, 'Attention : la ressource va reprendre les valeurs du salarié' // TG 4/04/2003
    {24}, 'Attention : le salarié %SALARIE% est lié à cette ressource.'#13'Confirmez-vous la suppression de cette ressource ?' // TG 4/04/2003
    {25}, 'Changement impossible : le salarié %SALARIE% est déjà lié à une ressource' // TG 4/04/2003
    {26}, 'Le matériel associé n''existe pas.'
    {27}, 'L''atelier saisi n''existe pas.'
    {28}, 'Suppression impossible : la ressource est présente dans des affaires (Intervenants)'
    {29}, 'Suppression impossible : la ressource est présente dans des tâches non générées'
    {30}, 'Suppression impossible : la ressource est présente dans des affaires (Principal)'
    {31}, 'Suppression impossible : la ressource est présente dans le planning'
    {32}, 'Code Section invalide.'
    );

procedure AFLanceFiche_Ressource(lequel, argument: string);

implementation
uses
  {$IFDEF GPAO}
    wRessource,
  {$ENDIF}
  wCommuns,
  BTPUtil,
  HrichOle,
  Ent1
  ;

var
  TauxPR: Double;

{ TOM_Ressource }

procedure TOM_Ressource.OnArgument(sArgument: string);
var sTmp			: String;
		DecPrix		: string;
 		menuPerso	: TMenuItem;
	  i					: integer;
    CodeAxe : string;
begin
  inherited;
  TOBRESSOURCEBTP := TOB.Create ('BRESSOURCE',nil,-1);
  AppliqueFontDefaut (THRichEditOle(GetControl('ARS_BLOCNOTE')));

  ModifSpecif := False;
  ModifStand := False;
  ModifSalarie := False;
  PremierPassage := True;

	TypeRessource := GetArgumentValue(sArgument, 'TYPERESSOURCE');

  GestionZoneEcran;

  if TCheckBox(GetControl('ARS_CALCULPR')) <> nil then
     Begin
		 CHKCALCULPR := TCheckBox(GetControl('ARS_CALCULPR'));
	   ChkCalculPR.OnClick := ChkCalculPR_OnClick;
     end;

  if (GetControl('ARS_TYPERESSOURCE')) <> nil then
		 Begin
     if TypeRessource <> '' then
        Begin
	      CBOTYPERESSOURCE.Value := TypeRessource;
        CBOTYPERESSOURCE.Enabled := False;
        end
     else
        CBOTYPERESSOURCE.Enabled := True;
  	 end;

  //Gestion de la valorisation
  CacherPrix := false;
  if not AffichageValorisation then CacherPrix := True;

  {$IFNDEF AFFAIRE}
  CacherPrix := True;
	{$ENDIF}

  // Si on ne doit pas accéder à la valorisation on vérouille l'accès à l'onglet Valorisation
  if CacherPrix then
     Begin
     if (GetControl('mnPersoPR') <> nil) then
        begin
        menuPerso := TMenuItem(GetControl('mnPersoPR'));
        menuPerso.Enabled := false;
        menuPerso.Visible := false;
        menuPerso := TMenuItem(GetControl('mnPersoPV'));
        menuPerso.Enabled := false;
        menuPerso.Visible := false;
        end;
  	 end
  else
  	 begin
     //Décimales prix unitaires
     if V_PGI.OkDecP=0 then
        DecPrix:='#,##0'
     Else
        DecPrix:='#,##0.' ;
     for i := 1 to V_PGI.OkDecP do DecPrix := DecPrix + '0';
     SetControlProperty('ARS_TAUXUNIT', 'DisplayFormat', DecPrix);
     SetControlProperty('ARS_TAUXREVIENTUN', 'DisplayFormat', DecPrix);
     SetControlProperty('ARS_PVHTCALCUL', 'DisplayFormat', DecPrix);
     SetControlProperty('ARS_PVHT', 'DisplayFormat', DecPrix);
     SetControlProperty('ARS_PVTTC', 'DisplayFormat', DecPrix);
     end;

  //Aiguillage pour gestion écran Ressource ou Matériel
  //Non en line
	GestionFicheRessource(sArgument);
{*
  if (TypeRessource = 'OUT') Or (TypeRessource = 'MAT') Or (TypeRessource = 'LOC') then
		 GestionFicheMateriel(sArgument)
  else
     GestionFicheRessource(sArgument);
*}

  if (sArgument = '') then exit;

  i := pos('MODIFLOT', sArgument);
  ModifLot := (i <> 0);
  if ModifLot then
     begin
     TFfiche(Ecran).MonoFiche := true;
     StSQL := copy(sArgument, i + 9, length(sArgument));
     exit;
     end;

  sTmp := StringReplace(sArgument, ';', chr(VK_RETURN), [rfReplaceAll]);

  // On réceptionne les paramètres passés par l'écran appelant, s'il y en a
  tsArgsReception := TStringList.Create;
  tsArgsReception.Text := sTmp;

  if (tsArgsReception.Count = 0) then
     begin
     tsArgsReception.Free;
     tsArgsReception := nil;
     exit;
     end;

  // On adapte la configuration de l'écran suivant la provenance de l'appel
  if (tsArgsReception.Values['ORIGINE'] = 'PSA') then
     // Si on vient de l'écran de saisie des salariés
     begin
     // On ne doit pas pouvoir rappeler la saisie des salariés
     SetControlEnabled('BSALARIE', False);
     // On initialise le type de ressource à Salarié et on le protège
     SetField('ARS_TYPERESSOURCE', 'SAL');
     SetControlEnabled('ARS_TYPERESSOURCE', False);
     TFFiche(Ecran).Monofiche := True; //Navigation interdite
     end;

  CodeAxe := GetparamSocSecur('SO_BTAXEANALSTOCK','TX1');
  if VH^.LiaisonY2ViaShare then
  begin
    THEdit(GetControl('BRS_SECTION')).DataType := 'TZCSECTION'+Copy(CodeAxe,2,1);
  end else
  begin
    THEdit(GetControl('BRS_SECTION')).DataType := 'TZSECTION'+Copy(CodeAxe,2,1);
  end;

//uniquement en line
//  if (GetControl('ARS_CALCULPR') <> nil) then  SetControlText('ARS_CALCULPR', 'X');
end;

Procedure TOM_RESSOURCE.ChkCalculPR_OnClick(Sender: TObject);
Begin

	if ChkCalculPR.Checked then
		 Setfield('ARS_CALCULPR', 'X')
  else
	   Setfield('ARS_CALCULPR', '-');

end;

procedure TOM_Ressource.GestionFicheMateriel(StArgument: String);
begin

end;

procedure TOM_Ressource.GestionFicheRessource(StArgument: String);
Var PPRIX				: TTabSheet;
    PCOMPLEMENT : TTabSheet;
    PLIBRE			: TTabSheet;
    PINFORMATION: TTabSheet;
begin

  if (GetControl('ARS_CHAINEORDO') <> nil) then
     Begin
     THEdit(GetControl('ARS_CHAINEORDO')).OnExit := SsFamilleOnExit;
     THEdit(GetControl('ARS_CHAINEORDO')).OnElipsisClick := SsFamilleOnElipsisClick;
     end;

  if (GetControl('PPRIX') <> nil) then PPRIX := TTabSheet(GetControl('PPRIX'));
  if (GetControl('TABSHEETCOMPLEMENT') <> nil) then PCOMPLEMENT := TTabSheet(GetControl('TABSHEETCOMPLEMENT'));
  if (GetControl('PLIBRE') <> nil) then PLIBRE := TTabSheet(GetControl('PLIBRE'));
  if (GetControl('GINFORMATIONS') <> nil) then PINFORMATION := TTabSheet(GetControl('GINFORMATIONS'));

	if cacherPrix then
     begin
     if (PPRIX <> nil) then
        begin
        PPRIX.Enabled := false;
        PPRIX.TabVisible := false;
        end;
     end;

  //mcd 10/06/02 ajout pour ne pas afficher si pas lien paie
  if GetParamSoc('SO_AFLIENPAIEVAR') then
     if ctxScot in V_PGI.PgiCOntexte then SetControlCaption('ARS_MENSUALISE', 'Assistant mensualisé')
  else
     SetControlVisible('ARS_MENSUALISE', FALSE);

  SetControlVisible('ARS_ACTIVITEPAIE', False); // En attente ...
  SetField('ARS_ACTIVITEPAIE', '-');
  SetControlVisible('ARS_GROUPERES', False);
  SetControlVisible('TARS_GROUPERES', False); // En attente ...

{$IFNDEF AFFAIRE}
  menuPerso := TMenuItem(GetControl('mnCompetence'));
  if (menuPerso <> nil) then menuPerso.Visible := false;
  SetControlVisible('BPERSOCALENDRIER', False);
  SetControlVisible('ARS_CALENSPECIF', False);
  SetControlvisible('ARS_STANDCALEN', False);
  SetControlvisible('TARS_STANDCALEN', False);
{$ENDIF}

  GCMAJChampLibre(TForm(Ecran), False, 'COMBO', 'ARS_LIBRERES', 10, '');
  GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'ARS_VALLIBRE', 3, '');
  GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'ARS_DATELIBRE', 3, '');
  GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'ARS_CHARLIBRE', 3, '');
  GCMAJChampLibre(TForm(Ecran), False, 'BOOL', 'ARS_BOOLLIBRE', 3, '');

{$IFDEF GPAO}
  if ctxGPAO in V_PGI.PGIContexte then
  	 begin
   	 TypeRessource := GetArgumentValue(sArgument, 'TYPERESSOURCE');
   	 bEstFamille := GetArgumentValue(sArgument, 'ESTFAMILLE') <> '';
   	 if (Ecran <> nil) then
   			begin
      	BRESSOURCELIE := tToolBarButton97(GetControl('BRESSOURCELIE'));
      	if BRESSOURCELIE <> nil then BRESSOURCELIE.OnClick := BRessourceLie_OnClick;
		    end;
  	 end;
{$ENDIF}

  //Initialisation des valeurs propres à l'écran ressource
  CBOSTANDCALEND.Value := '';

	// On ne distingue plus seulement les ressource par S3 ou S5 mais
	//  pour S3, si on a la GA de sérialiser (dans GA et/ou GC)  on a une ressource plus compléte
  if (V_PGI.LaSerie <> S3)  or
{$IFDEF BTP}
	   (V_PGI.LaSerie = S3) then
{$ELSE}
     ((V_PGI.LaSerie = S3) and  (VH_GC.GASERIA) ) then
{$ENDIF}
		 begin  // on met tous les onglets
     if GetControl('ARS_UTILASSOCIE_')  <> nil then   SetControlVisible('ARS_UTILASSOCIE_', False);
   	 if GetControl('TARS_UTILASSOCIE_') <> nil then   SetControlVisible('TARS_UTILASSOCIE_', False);
     if GetControl('ARS_FONCTION1_')    <> nil then   SetControlVisible('ARS_FONCTION1_', False);
     if GetControl('TARS_FONCTION1_')   <> nil then   SetControlVisible('TARS_FONCTION1_', False);
		 end
  else
  	 begin  // on n' a que 2 onglets
     if GetControl('BMENUZOOM')         <> nil then   SetControlVisible('BMENUZOOM', False);
     if GetControl('BPERSOCALENDRIER')  <> nil then   SetControlVisible('BPERSOCALENDRIER', False);
     if GetControl('TARS_STANDCALEN')   <> nil then   SetControlVisible('TARS_STANDCALEN', False);
     if GetControl('ARS_STANDCALEN')    <> nil then   SetControlVisible('ARS_STANDCALEN', False);
     if GetControl('ARS_CALENSPECIF')   <> nil then   SetControlVisible('ARS_CALENSPECIF', False);
     if PPRIX <> Nil then PPRIX.TabVisible := False;
     if PCOMPLEMENT <> Nil then PCOMPLEMENT.TabVisible := False;
     if PLIBRE <> Nil then PLIBRE.TabVisible := False;
     end;
  //


//uniquement en line
{*
	if PPRIX <> Nil then PPRIX.Visible := False;
  if PCOMPLEMENT <> Nil then PCOMPLEMENT.TabVisible := False;
  if PLIBRE <> Nil then PLIBRE.TabVisible := False;
  if PINFORMATION <> Nil then PINFORMATION.TabVisible := False;
//  If CHKCALCULPR <> nil then CHKCALCULPR.Checked := True;
*}

end;

Procedure TOM_RESSOURCE.GestionZoneEcran;
Var Bt	: TToolBarButton97;
    Lbl	: THLabel;
	  {$IFDEF EAGLCLIENT}
	  CC	: THEdit;
		{$ELSE}
	  CC	: THDBEdit;
	  {$ENDIF}
Begin

  if BTPERSOCALEND <> Nil then BTPERSOCALEND.Enabled := True;

  if (GetControl('TARS_RESSOURCE') <> nil) then
		 TLabelRessource := THLabel(ecran.FindComponent('TARS_RESSOURCE'));

  if (GetControl('TARS_TYPE') <> nil) then
	   TLabelType := THLabel(ecran.FindComponent('TARS_TYPE'));

	{$IFDEF EAGLCLIENT}
  if (GetControl('ARS_SALARIE') <> nil) then
     THEdit(GetControl('ARS_SALARIE')).OnExit := ZoneSalarieOnExit;

  if (GetControl('ARS_SALARIE') <> nil) then
     THEdit(GetControl('ARS_SALARIE')).OnEnter := ZoneSalarieOnEnter;

  if GetControl('ARS_DATEFONC1') <> nil then
	   begin
     CC := THEdit(GetControl('ARS_DATEFONC1'));
     CC.OnExit := verifFonction;
  	 end;
  if GetControl('ARS_DATEFONC2') <> nil then
  	 begin
     CC := THEdit(GetControl('ARS_DATEFONC2'));
     CC.OnExit := verifFonction;
  	 end;
  if GetControl('ARS_DATEFONC3') <> nil then
  	 begin
     CC := THEdit(GetControl('ARS_DATEFONC3'));
     CC.OnExit := verifFonction;
  	 end;
  if GetControl('ARS_DATEFONC4') <> nil then
  	 begin
     CC := THEdit(GetControl('ARS_DATEFONC4'));
     CC.OnExit := verifFonction;
  	 end;

	{$ELSE}
  if GetControl('ARS_SALARIE') <> nil then
     THDBEdit(GetControl('ARS_SALARIE')).OnExit := ZoneSalarieOnExit;
  if GetControl('ARS_SALARIE') <> nil then
     THDBEdit(GetControl('ARS_SALARIE')).OnEnter := ZoneSalarieOnEnter;
  if GetControl('ARS_DATEFONC1') <> nil then
  	 begin
     CC := THDBEdit(GetControl('ARS_DATEFONC1'));
     CC.OnExit := verifFonction;
  	 end;
  if GetControl('ARS_DATEFONC2') <> nil then
  	 begin
     CC := THDBEdit(GetControl('ARS_DATEFONC2'));
     CC.OnExit := verifFonction;
  	 end;
  if GetControl('ARS_DATEFONC3') <> nil then
     begin
     CC := THDBEdit(GetControl('ARS_DATEFONC3'));
     CC.OnExit := verifFonction;
  	 end;
  if GetControl('ARS_DATEFONC4') <> nil then
  	 begin
     CC := THDBEdit(GetControl('ARS_DATEFONC4'));
     CC.OnExit := verifFonction;
  	 end;
	{$ENDIF}

  //Gestion des zones ecrans
  if (GetControl('BNEWMAIL') <> nil) then
     Begin
     BTMAIL := TToolbarButton97(ecran.FindComponent('BNEWMAIL'));
     BTMAIL.OnClick := MailClick;
     End;

  if (GetControl('ARS_CODEARTICLE') <> nil) then
	   CodeArticle := THDBEdit(ecran.FindComponent('ARS_CODEARTICLE'));


  if (GetControl('ARS_EMAIL') <> nil) then
     Begin
	   AdresseMail := THDBEdit(ecran.FindComponent('ARS_EMAIL'));
     AdresseMail.OnChange := MailChange;
     AdresseMAil.OnDblClick := MailDblClick;
     end;

  if (GetControl('ARS_CODEPOSTAL') <> nil) then
     Begin
	   CodePostal := THDBEdit(ecran.FindComponent('ARS_CODEPOSTAL'));
     CodePostal.OnDblClick := ControleAdresse;
     End;

  if (GetControl('ARS_VILLE') <> nil) then
     Begin
	   Ville := THDBEdit(ecran.FindComponent('ARS_VILLE'));
     Ville.OnDblClick := ControleAdresse;
     end;

  //Gestion du calendrier spécifique ou standard
  if (GetControl('ARS_STANDCALEN') <> nil) then
     Begin
  	 CBOSTANDCALEND := THValComboBox(ecran.FindComponent('ARS_STANDCALEN'));
	   CBOSTANDCALEND.OnChange:= StandCalenClick;
     end;

  if (GetControl('BPERSOCALENDRIER') <> nil) then
     Begin
  	 BTPERSOCALEND := TToolbarButton97(ecran.FindComponent('BPERSOCALENDRIER'));
  	 BTPERSOCALEND.OnClick := BtPersoCalendClick;
     end;

  if (GetControl('BSALARIE') <> nil) then
     Begin
  	 BTSALARIE := TToolbarButton97(ecran.FindComponent('BSALARIE'));
  	 BTSALARIE.OnClick := BtSalarieClick;
     end;

  if (GetControl('ARS_CALENSPECIF') <> nil) then
     Begin
     CHKCALENSPECIF := TCheckBox(ecran.FindComponent('ARS_CALENSPECIF'));
	   CHKCALENSPECIF.OnClick := CBSpecifClick;
     End;

  //Gestion du type de ressource
  if (GetControl('ARS_TYPERESSOURCE')) <> nil then
		 CBOTYPERESSOURCE := THDBValComboBox(ecran.FindComponent('ARS_TYPERESSOURCE'));

  if (GetControl('BHAUTFONCTION') <> nil) then
  	 begin
     Bt := TToolBarButton97(GetControl('BHAUTFONCTION'));
     Bt.Onclick := MoveFonction;
  	 end;

  if (GetControl('BBASFONCTION') <> nil) then
  	 begin
     Bt := TToolBarButton97(GetControl('BBASFONCTION'));
     Bt.Onclick := MoveFonction;
  	 end;

  if (GetControl('BREVALORISATION') <> nil) then
  	 begin
     BREVALORISATION := TToolbarButton97(GetControl('BREVALORISATION'));
     BRevalorisation.OnClick := BRevalorisationClick;
  	 end;

  if (GetControl('BHISTOVALO') <> nil) then
  	 begin
     BHISTOVALO := TToolbarButton97(GetControl('BHISTOVALO'));
     BHistoValo.OnClick := BHistoValoClick;
  	 end;

{$IFDEF GPAO}
  if (IsEcranGPAO) then
  	 begin
     SetControlProperty('BIMPRIMER', 'VISIBLE', GetArgumentValue(sArgument, 'ESTHUMAIN') = '-');
     SetControlProperty('BTIMPRIMER', 'VISIBLE', GetArgumentValue(sArgument, 'ESTHUMAIN') = 'X');
     if (GetControl('BIMPRIMER') <> nil) then
        TToolBarButton97(GetControl('BIMPRIMER')).OnClick := BImprimer_OnClick;
     if (GetControl('TSPRODUCTION') <> nil) then
        TTabSheet(GetControl('TSPRODUCTION')).Visible := True;
     if (GetControl('BRESSOURCELIE') <> nil) then
        TToolBarButton97(GetControl('BRESSOURCELIE')).Visible := True;
     end;
{$ENDIF}

//Gestion du libellé de l'identification...
	if (GetControl('TARS_IMMAT') <> Nil) then
     Begin
     Lbl := THLabel(ecran.FindComponent('TARS_IMMAT'));
     if (TypeRessource = 'SAL')      Then
        Lbl.Caption := 'N° Sécurité Sociale :'
     Else if (TypeRessource = 'ST')  Then
        Lbl.Caption := 'N° Sécurité Sociale :'
     Else if (TypeRessource = 'INT') Then
        Lbl.Caption := 'N° Sécurité Sociale :'
	   Else if (TypeRessource = 'OUT') Then
        Lbl.Caption := 'N° Série :'
     Else if (TypeRessource = 'MAT') Then
        Lbl.Caption := 'N° Immatriculation :'
     Else if (TypeRessource = 'LOC') then
        Lbl.Caption := 'N° Contrat Location :';
     CC := THDBEdit(Ecran.FindComponent('ARS_IMMAT'));
     CC.Left := lbl.left + lbl.Width + 12;
     end;
end;

procedure TOM_Ressource.OnNewRecord;
var
  sCodeRessource: string;
begin
  //
  EnSaisie := False; // initialisation par défaut
  //
  SetField('ARS_UNITETEMPS', VH_GC.AFMesureActivite);
  SetField('ARS_TAUXFRAISGEN1', Valeur(GetParamSocSecur('SO_AFFRAISGEN1', 0)));
  SetField('ARS_TAUXFRAISGEN2', Valeur(GetParamSocSecur('SO_AFFRAISGEN2', 0)));
  SetField('ARS_TAUXCHARGEPAT', Valeur(GetParamSocSecur('SO_AFCHARGESPAT', 0)));
  SetField('ARS_COEFMETIER', Valeur(GetParamSocSecur('SO_AFCOEFMETIER', 0)));
  SetField('ARS_ARTICLE', VH_GC.AFPrestationRes);
  SetField('ARS_FERME', '-');

  if (GetControl('ARS_TYPERESSOURCE')) <> nil then
		 Begin
     if TypeRessource <> '' then
        Begin
	      CBOTYPERESSOURCE.Value := TypeRessource;
        CBOTYPERESSOURCE.Enabled := False;
        end
     else
        Begin
        CBOTYPERESSOURCE.Value := 'SAL';
        CBOTYPERESSOURCE.Enabled := True;
        end;
  	 end;

  SetField('ARS_MENSUALISE', 'X');
  SetField('ARS_TAUXSIMUL', 0);
  SetField('ARS_GENERIQUE', '-');
  SetField('ARS_SECTIONPDR', '');
  SetField('ARS_RUBRIQUEPDR', '');
  SetField('ARS_ESTHUMAIN', '-');
  SetField('ARS_SITE', '');
  SetField('ARS_DEPOT', '');
  //
  if BTPersoCalend <> nil then BTPersoCalend.Enabled := False;
  If CHKCALENSPECIF <> nil then CHKCALENSPECIF.Checked := False;
  //
  if ctxGPAO in V_PGI.PGIContexte then
     begin
     if (ECRAN <> nil) and (ECRAN.Name = 'WRESSOURCE_FIC') then
        SetField('ARS_TYPERESSOURCE', TypeRessource);
     if (ECRAN <> nil) and (ECRAN.Name = 'RESSOURCE') then
        SetField('ARS_ESTHUMAIN', 'X');
     if bEstFamille then SetField('ARS_GENERIQUE', 'X');
     SetFocusControl('ARS_RESSOURCE');
     end;

  SetField('ARS_PAYS', GetParamSoc('SO_GcTiersPays'));
  SetField('ARS_COEFPRPV', 1.0);

  SetControlVisible('BHISTOVALO', False);
  SetControlEnabled('BPERSOCALENDRIER', False);

//chargement option calcul auto du prix de revient...
  if (VH_GC.AFResCalculPR) then
    CHKCALCULPR.Checked := True
  else
    CHKCALCULPR.Checked := False;

  SetField('ARS_DEBUTDISPO', idate1900);
  SetField('ARS_FINDISPO', idate2099);
  SetField('ARS_DATESORTIE', idate2099);
  SetField('ARS_CREATEUR', V_PGI.User);
  SetField('ARS_UTILISATEUR', V_PGI.User);
  SetField('ARS_CREERPAR', 'SAI'); // saisie

  // Initialisation de l'affichage, due à un appel depuis un autre écran de saisie
  // (ex : saisie des salariés)
  if (tsArgsReception <> nil) then
    // Si une StringList de réception a été passée en paramètre
    // la variable de réception de l'écran Ressource est remplie
    with tsArgsReception do
    begin
      if (tsArgsReception.Values['ORIGINE'] = 'PSA') then
      begin
        if (Values['SALARIE'] <> '') then
            // Si le code salarié n'est pas vide, on va chercher le code ressource
            // associé.
        begin
          sCodeRessource := RessourceDuSalarie(Values['SALARIE']);
          if sCodeRessource = '' then
            SetField('ARS_RESSOURCE', Values['SALARIE'])
          else
            SetField('ARS_RESSOURCE', sCodeRessource);
        end;
        SetField('ARS_SALARIE', Values['SALARIE']);
        SetField('ARS_AUXILIAIRE', Values['AUXILIAIRE']);
        SetField('ARS_TIERS', Values['TIERS']);
        SetField('ARS_LIBELLE', Values['LIBELLE']);
        SetField('ARS_ADRESSE1', Values['ADRESSE1']);
        SetField('ARS_ADRESSE2', Values['ADRESSE2']);
        SetField('ARS_ADRESSE3', Values['ADRESSE3']);
        SetField('ARS_CODEPOSTAL', Values['CODEPOSTAL']);
        SetField('ARS_VILLE', Values['VILLE']);
        SetField('ARS_PAYS', Values['PAYS']);
        SetField('ARS_TAUXUNIT', Values['TAUXUNIT']);
        QuestHisto := False;
      end;
    end;

  //FV1 : 04/05/2017 - FS#2528 - FVE - Gérer le SAV sur les Intérim comme sur les Salariés
  if (GetControlText('ARS_TYPERESSOURCE')='SAL') OR (GetControlText('ARS_TYPERESSOURCE')='INT') then
  begin
    SetControlVisible('BRS_GERESAV',True);
    SetControlVisible('BRS_GERECHANTIER',True);
  end else
  begin
    SetControlVisible('BRS_GERESAV',False);
    SetControlVisible('BRS_GERECHANTIER',False);
  end;

end;

procedure TOM_Ressource.OnDeleteRecord;
{$IFDEF GPAO} //GISE
var
  CleARS: tCleArs;
  {$ENDIF}
begin
  (*
  if ExisteSQL('SELECT ACT_RESSOURCE FROM ACTIVITE WHERE ACT_RESSOURCE="' + GetField('ARS_RESSOURCE') + '"') then
     begin
     LastError := 10;
     LastErrorMsg := TexteMsgRessource[LastError];
     exit;
     end;
  *)

  if ExisteSQL('SELECT BCO_RESSOURCE FROM CONSOMMATIONS WHERE BCO_RESSOURCE="' + GetField('ARS_RESSOURCE') + '"') then
     begin
     LastError := 10;
     LastErrorMsg := TexteMsgRessource[LastError];
     exit;
     end;

  if ExisteSQL('SELECT BEP_RESSOURCE FROM BTEVENPLAN WHERE BEP_RESSOURCE="' + GetField('ARS_RESSOURCE') + '"') then
     begin
     LastError := 31;
     LastErrorMsg := TexteMsgRessource[LastError];
     exit;
     end;

  if ExisteSQL('SELECT AFT_RESSOURCE FROM AFFTIERS WHERE AFT_RESSOURCE="' + GetField('ARS_RESSOURCE') + '"') then
     begin
     LastError := 28;
     LastErrorMsg := TexteMsgRessource[LastError];
     exit;
     end;

  if ExisteSQL('SELECT AFF_RESPONSABLE FROM AFFAIRE WHERE AFF_RESPONSABLE="' + GetField('ARS_RESSOURCE') + '"') then
     begin
     LastError := 28;
     LastErrorMsg := TexteMsgRessource[LastError];
     exit;
     end;

  if ExisteSQL('SELECT ATR_RESSOURCE,ATA_TERMINE FROM TACHERESSOURCE '+
  							'LEFT JOIN TACHE ON ATR_AFFAIRE=ATA_AFFAIRE AND ATR_NUMEROTACHE=ATA_NUMEROTACHE '+
                'WHERE ATR_RESSOURCE="' + GetField('ARS_RESSOURCE') +'" AND '+
                'ATA_TERMINE="-"') then
     begin
     LastError := 30;
     LastErrorMsg := TexteMsgRessource[LastError];
     exit;
     end;

  (* pas de champ ressource pour l'instant dans histo ..if ExisteSQL('SELECT AHI_RESSOURCE FROM HISTOACTIVITE WHERE AHI_RESSOURCE="'+GetField('ARS_RESSOURCE')+'"') then
   Begin LastError:=12 ; LastErrorMsg:=TexteMsgRessource[LastError] ; exit; End; *)

  if ExisteSQL('SELECT Gl_RESSOURCE FROM LIGNE WHERE GL_RESSOURCE="' + GetField('ARS_RESSOURCE') + '"') then
     begin
     LastError := 13;
     LastErrorMsg := TexteMsgRessource[LastError];
     exit;
     end;

  {$IFDEF GPAO}
  if ctxGPAO in V_PGI.PGIContexte then
     begin
     CleARS.Ressource := GetField('ARS_RESSOURCE');
     LastError := WArsAllowDelete(CleARS);
     if LastError <> 0 then
        begin
        LastErrorMsg := TexteMsgRessource[LastError];
        Exit;
        end;
     end;
  {$ENDIF}

  // Ajout TG 04/04/2003
  if GetParamSoc('SO_PGLIENRESSOURCE') then
     if (GetField('ARS_SALARIE') <> '') then
        if (PGIAskAF(StringReplace(TexteMsgRessource[24], '%SALARIE%', RechDom('PGSALARIE', OldSalarie, false), [rfReplaceAll, rfIgnoreCase]), '') <> mrYes) then
           begin
           LastError := -1;
           LastErrorMsg := '';
           exit;
           end;

  // Si la suppression est OK, on supprime la ressource dans les tables connexes
  ExecuteSQL('DELETE FROM CALENDRIER WHERE ACA_RESSOURCE="' + string(GetField('ARS_RESSOURCE')) + '"');
  ExecuteSQL('DELETE FROM CALENDRIERREGLE WHERE ACG_RESSOURCE="' + string(GetField('ARS_RESSOURCE')) + '"');
  ExecuteSQL('DELETE FROM RESSOURCEPR WHERE ARP_RESSOURCE="' + string(GetField('ARS_RESSOURCE')) + '"');
  ExecuteSQL('DELETE FROM COMPETRESSOURCE WHERE ACR_RESSOURCE="' + string(GetField('ARS_RESSOURCE')) + '"');
  //
  DeleteRessourceBTP (GetField('ARS_RESSOURCE'));

end;

procedure TOM_Ressource.OnChangeField(F: TField);
var
{$IFDEF EAGLCLIENT}
  ControlSalarie: THEdit;
{$ELSE}
  ControlSalarie: THDBEDIT;
{$ENDIF}
  sWhere: string;
begin

  if ((F.FieldName = 'ARS_TAUXUNIT') or (F.FieldName = 'ARS_TAUXFRAISGEN1') or (F.FieldName = 'ARS_TAUXFRAISGEN2') or
     (F.FieldName = 'ARS_TAUXCHARGEPAT') or (F.FieldName = 'ARS_COEFMETIER') or (F.FieldName = 'ARS_CALCULPR')) then
     Begin
     SetField('ARS_TAUXREVIENTUN', Ressource_CalculPR());
     Exit;
     end;

  if ((F.FieldName = 'ARS_TAUXREVIENTUN') or (F.FieldName = 'ARS_COEFPRPV') or (F.FieldName = 'ARS_CALCULPV')) then
		 Begin
     SetField('ARS_PVHTCALCUL', Ressource_CalculPV());
     exit;
     end;

  if (F.FieldName = 'ARS_PVHTCALCUL') then
     if (GetField('ARS_PVHT') = 0) and (GetField('ARS_PVHTCALCUL') <> 0) then SetField('ARS_PVHT', getField('ARS_PVHTCALCUL'));

  if (ctxGPAO in V_PGI.PGIContexte) and (Ecran <> nil) and (Ecran.Name = 'WRESSOURCE_FIC') then
      begin
      if (F.FieldName = 'ARS_RESSOURCE') or (F.FieldName = 'ARS_TYPERESSOURCE') then
         begin
            { Filtre sur les ressources liées }
            sWhere := '     ARS_RESSOURCE <> "' + GetField('ARS_RESSOURCE') + '"'
              + ' AND ARS_TYPERESSOURCE = "' + GetField('ARS_TYPERESSOURCE') + '"'
              + ' AND ARS_WGENERIQUE = "X"';
{$IFDEF EAGLCLIENT} //GISE
            if GetControl('ARS_WRESSOURCELIE') <> nil then thEdit(GetControl('ARS_WRESSOURCELIE')).Plus := sWhere;
{$ELSE}
            if GetControl('ARS_WRESSOURCELIE') <> nil then thdbEdit(GetControl('ARS_WRESSOURCELIE')).Plus := sWhere;
{$ENDIF}
         end;
      end;

  if (F.FieldName = 'ARS_TYPERESSOURCE') then
  begin
    if (TypeRessource = 'ST') then
    begin
      // Cacher les zones salarié
      SetControlVisible('ARS_SALARIE', False);
      SetControlVisible('TARS_SALARIE', False);
      SetField('ARS_SALARIE', '');
      SetControlVisible('NOMSALARIE', False);
      SetControlVisible('ARS_MENSUALISE', False);
      SetControlEnabled('BSALARIE', False);
      if (VH_GC.GAAchatSeria) then
      begin
        SetControlVisible('ARS_AUXILIAIRE', True);
        SetControlVisible('TARS_AUXILIAIRE', True);
        SetControlVisible('NOMAUXILIAIRE', True);
      end
      else
      begin
        SetControlVisible('ARS_AUXILIAIRE', False);
        SetControlVisible('NOMAUXILIAIRE', False);
        SetControlVisible('TARS_AUXILIAIRE', False);
      end;

          // mcd 21/03/01 SetControlText('TARS_IMMAT','Ré&férence');
          //SetControlVisible('ARS_ACTIVITEPAIE',False); SetField('ARS_ACTIVITEPAIE','-');
      if Ecran.Name <> 'WRESSOURCEFAM_FIC' then
      begin
        SetControlText('TARS_LIBELLE', '&Nom');
        SetControlText('TARS_LIBELLE2', 'Prén&om');
      end;
    end
    else
      if (typeRessource = 'SAL') then
      begin
        SetControlVisible('ARS_AUXILIAIRE', False);
        SetControlVisible('TARS_AUXILIAIRE', False);
        if not GetParamSoc('SO_PGLIENRESSOURCE') then //Affaire - Lien Paie
          SetField('ARS_AUXILIAIRE', '');
            // mcd 21/03/01 SetControlText('TARS_IMMAT','Matr&icule');
        SetControlVisible('ARS_SALARIE', True);
        SetControlVisible('TARS_SALARIE', True);
        SetControlVisible('NOMAUXILIAIRE', False);
        SetControlVisible('NOMSALARIE', True);
        if GetParamSoc('SO_AFLIENPAIEVAR') then SetControlVisible('ARS_MENSUALISE', True)
        else SetControlVisible('ARS_MENSUALISE', False);
            // SetControlVisible('ARS_ACTIVITEPAIE',True); SetField('ARS_ACTIVITEPAIE','-');
        if Ecran.Name <> 'WRESSOURCEFAM_FIC' then
           begin
           SetControlText('TARS_LIBELLE', '&Nom');
           SetControlText('TARS_LIBELLE2', 'Prén&om');
           end;
      end
      //FV1 : 25/07/2013 ==> FS#529 - AURUS : pb affichage écran si création de ressource depuis la zone "intervenant" 
      else // Type matériel et surtout pas libre sinon zones chargées à blanc...
      if (typeRessource = 'MAT') OR (typeRessource = 'OUT') then
      begin
        SetControlVisible('ARS_SALARIE', False);
        SetControlVisible('ARS_AUXILIAIRE', False);
        SetControlVisible('TARS_SALARIE', False);
        SetControlVisible('TARS_AUXILIAIRE', False);
        SetField('ARS_AUXILIAIRE', '');
        SetField('ARS_SALARIE', '');
        SetControlEnabled('BSALARIE', False);
          // mcd 21/03/01 SetControlText('TARS_IMMAT','Ré&férence');
        SetControlVisible('NOMAUXILIAIRE', False);
        SetControlVisible('NOMSALARIE', False);
        if GetParamSoc('SO_AFLIENPAIEVAR') then SetControlVisible('ARS_MENSUALISE', GetField('ARS_TYPERESSOURCE') = 'INT')
        else SetControlVisible('ARS_MENSUALISE', False);
        //SetControlVisible('ARS_ACTIVITEPAIE',False); SetField('ARS_ACTIVITEPAIE','-');
//Non en line
        SetControlText('TARS_LIBELLE', 'L&ibellé');
        if (GetControl('TARS_LIBELLE2') <> nil) and (GetField('ARS_TYPERESSOURCE') <> 'INT') then
          SetControlText('TARS_LIBELLE2', '');
      end;
  end
  else
    if (F.FieldName = 'ARS_SALARIE') then
    begin
      if (GetField('ARS_SALARIE') <> '') then
      begin
{$IFDEF EAGLCLIENT}
        ControlSalarie := THEdit(GetControl('ARS_SALARIE'));
{$ELSE}
        ControlSalarie := THDBEdit(GetControl('ARS_SALARIE'));
{$ENDIF}
        if (not GetParamSoc('SO_AFLIENPAIEDEC')) and (not LookUpValueExist(ControlSalarie)) then // Ajout GetParamSoc TG 9/04/2003
        begin
          PGIBoxAF(TexteMsgRessource[4], '');
          SetControlEnabled('BSALARIE', False);
          SetFocusControl('ARS_SALARIE');
          Exit;
        end
        else
          SetControlEnabled('BSALARIE', True);
      end;

      if (GetField('ARS_SALARIE') <> OldSalarie) then
      begin
             // Ajout TG 04/04/2003
        if ExisteLienSalarie(GetField('ARS_SALARIE'), GetField('ARS_RESSOURCE')) then
        begin
          PGIBoxAF(StringReplace(TexteMsgRessource[25], '%SALARIE%', RechDom('PGSALARIE', GetField('ARS_SALARIE'), false), [rfReplaceAll, rfIgnoreCase]), '');
          SetField('ARS_SALARIE', OldSalarie);
          SetFocusControl('ARS_SALARIE');
        end else
             //
            //  Test changement de code salarié
          if (OldSalarie <> '') then
          begin
            if not TestImpactChangeSalarie(OldSalarie, GetField('ARS_RESSOURCE')) then
            begin
              SetField('ARS_SALARIE', OldSalarie);
              SetFocusControl('ARS_SALARIE');
            end;
          end;
        OldSalarie := GetField('ARS_SALARIE');
        CheckRattachementSalarie; // Ajout TG 3/04/2003
      end;
    end
    else
    begin
      {$IFDEF GPAO}
      if (ctxGPAO in V_PGI.PGIContexte) then
      begin
        if (F.FieldName = 'ARS_DEPOT') then
        begin
          if GetField('ARS_DEPOT') <> '' then
            SetControlProperty('ARS_SITE','PLUS','QSI_DEPOT="'+GetField('ARS_DEPOT')+'"')
          else
            SetControlProperty('ARS_SITE','PLUS','');
        end
        else if (F.FieldName = 'ARS_SITE') then
        begin
          if GetField('ARS_DEPOT') = '' then
          begin
            SetControlProperty('ARS_SITE','PLUS','');
            SetField('ARS_DEPOT', wGetSqlFieldValue('QSI_DEPOT','QSITE', 'QSI_SITE="' + GetField('ARS_SITE') + '"'));
          end
        end;
      end;
      {$ENDIF}
    end;

    // ajout mcd 11/09/01 pour désactivier bouton historique si modif d'un prix (il faut valider avant
  {If ((F.FieldName='ARS_TAUXUNIT') Or (F.FieldName='ARS_TAUXFRAISGEN1') Or (F.FieldName='ARS_TAUXFRAISGEN2') Or
   (F.FieldName='ARS_TAUXCHARGEPAT') Or(F.FieldName='ARS_PVHT') Or(F.FieldName='ARS_DATEPRIX') Or
   (F.FieldName='ARS_UNITETEMPS') Or(F.FieldName='ARS_PVTTC') Or(F.FieldName='ARS_COEFPRPV') Or (F.FieldName='ARS_PVHTCALCUL') Or
    (F.FieldName='ARS_TAUXCHARGEPAT') Or (F.FieldName='ARS_COEFMETIER') Or (F.FieldName='ARS_TAUXREVIENTUN'))then
     begin
     If (VerifChangements > 0) then SetControlVisible('BHISTOVALO',False);
     end;}// NCX 19/09/01
end;


procedure TOM_Ressource.OnUpdateRecord;
var
  InfTiers: Info_Tiers;
  sCodeAuxi, NomChamp: string;
  stRapport: string;
  modif: integer;

{$IFDEF DEBUG}
  sTmp, sCodeSalarie: string;
{$ENDIF}
begin

  // Gestion des erreurs sur les zones libres
  if EnSaisie then
     begin
    (*if (Getfield('ARS_FONCTION1') <> '') then begin
    // fait pour tester ressourcpr depuis fonction !!!!
    // ex qd valo fct sera faite  (nb penser aussi de passer le typeaction. voir fiche ressource...
     AglLancefiche ('AFF','RESSOURCEPR','V;***;'+GetField('ARS_FONCTION1'),'',';ORIGINE=FCT;Typevalo=V');
     end; *)
     if (GetField('ARS_LIBRERES1') <> '') then
        if not LookupValueExist(GetControl('ARS_LIBRERES1')) then
           begin
           LastError := 14;
           LastErrorMsg := TraduitGa(THLabel(GetControl('TARS_LIBRERES1')).Caption + ' ' + TexteMsgRessource[LastError]);
           SetFocusControl('ARS_LIBRERES1');
           Exit;
           end;
     if (GetField('ARS_LIBRERES2') <> '') then
        if not LookupValueExist(GetControl('ARS_LIBRERES2')) then
           begin
           LastError := 14;
           LastErrorMsg := TraduitGa(THLabel(GetControl('TARS_LIBRERES2')).Caption + ' ' + TexteMsgRessource[LastError]);
           SetFocusControl('ARS_LIBRERES2');
           Exit;
           end;
     if (GetField('ARS_LIBRERES3') <> '') then
        if not LookupValueExist(GetControl('ARS_LIBRERES3')) then
           begin
           LastError := 14;
           LastErrorMsg := TraduitGa(THLabel(GetControl('TARS_LIBRERES3')).Caption + ' ' + TexteMsgRessource[LastError]); SetFocusControl('ARS_LIBRERES3');
           Exit;
           end;
     if GetControl('ARS_CHAINEORDO') <> nil then
        begin
        if (THEdit(GetControl('ARS_CHAINEORDO')).Text = '') then
           SetField('ARS_CHAINEORDO', '')
        else
           SetField('ARS_CHAINEORDO', GetControlText('ARS_CHAINEORDO'));
        end;

     end;

  if (Ensaisie) and (LastError = 0) and (GetControl('ARS_RESSOURCELIE') <> nil) and (GetField('ARS_RESSOURCELIE') <> '') then
     begin
     if not ExisteSql('SELECT 1 FROM RESSOURCE WHERE ARS_RESSOURCE="'+GetField('ARS_RESSOURCELIE')+'" AND ARS_TYPERESSOURCE="MAT"') then
        begin
        LastError := 26;
        LastErrorMsg := TexteMsgRessource[LastError];
        Exit;
        end;
     end;

  // Gestion des erreurs sur le code article
  if Ensaisie then
     begin
     // GPAO : La fiche de GP possédant moins de champ, il faudrait, à chaque "GetControl"... faire tout d'abord un "if getcontrol('...') <> nil"
     if CodeArticle <> nil then
        begin
        if CodeArticle.Text <> '' then
           //if not LookupValueExist(GetControl('ARS_CODEARTICLE')) then
           if not LookupValueExist(CodeArticle) then
              begin
              LastError := 15;
              LastErrorMsg := TraduitGa(TexteMsgRessource[LastError]);
              SetFocusControl('ARS_CODEARTICLE');
              exit;
              end;
        end;
     // Gestion du passage du CODEARTICLE en ARTICLE pour stocker
     //{$IFDEF GPAO}
     if GetControl('ARS_CODEARTICLE') <> nil then
        begin
        if CodeArticle.Text = '' then
           SetField('ARS_ARTICLE', '')
        else
           SetField('ARS_ARTICLE', CodeArticleUnique(CodeArticle.Text, '', '', '', '', ''));
        end;
     //{$ELSE} //mcd 27/01/03 Pas GPAO
     //   // Gestion du passage du CODEARTICLE en ARTICLE pour stocker
     //   if CodeArticle.Text = '' then
     //      SetField('ARS_ARTICLE', '')
     //   else
     //      SetField('ARS_ARTICLE', CodeArticleUnique(CodeArticle.Text, '', '', '', '', ''));
     //{$ENDIF} // GPAO
     end;

  // Création de la fiche sous traitant pour tempo uniquement
  // si on met pour scot, penser de modifier ConfigSaisieRessource  dans utilsocaf
  if ((EnSaisie) and (GetField('ARS_RESSOURCE') <> '') and (GetField('ARS_LIBELLE') <> '')) then
     begin
     SetField('ARS_RESSOURCE', Trim(GetField('ARS_RESSOURCE')));
     // mcd 24/04/02 if (ctxTempo in V_PGI.PGIContexte) then
     if not (CtxScot in V_PGI.PGIContexte) then
        begin
        if ((GetField('ARS_TYPERESSOURCE') = 'ST') and (GetField('ARS_AUXILIAIRE') = '')) then
           // MAJ de la fiche sous Traitant
           if (PGIAskAF('Voulez-vous créer le tiers associé au sous traitant', '') = mrYes) then
              begin
              FillChar(InfTiers, Sizeof(InfTiers), #0);
              InfTiers.Libelle := GetField('ARS_LIBELLE');
              InfTiers.Adresse1 := GetField('ARS_ADRESSE1');
              InfTiers.Adresse2 := GetField('ARS_ADRESSE2');
              InfTiers.Adresse3 := GetField('ARS_ADRESSE3');
              InfTiers.CodePostal := GetField('ARS_CODEPOSTAL');
              InfTiers.Ville := GetField('ARS_VILLE');
              InfTiers.Telephone := GetField('ARS_TELEPHONE');
              InfTiers.Pays := GetField('ARS_PAYS');
              sCodeAuxi := CreationTiers(InfTiers, stRapport);
              SetField('ARS_AUXILIAIRE', sCodeAuxi);
              end;
        end;
      // A voir avec paie ...
      {-$IFDEF DEBUG}
    if ((GetField('ARS_TYPERESSOURCE') = 'SAL') and (GetField('ARS_SALARIE') = '')
      and (not (ctxScot in V_PGI.PGIContexte)))
      and GetParamSoc('SO_PGLIENRESSOURCE') then
      PGIInfoAf('Attention : Vous n''avez pas associé de salarié à cette ressource.', '');
(*        // MAJ de la fiche salarié
        if (PGIAskAF ('Voulez-vous créer le salarié associé', '') = mrYes) then
          begin
            sTmp := 'ORIGINE=ARS;SALARIE=' + GetField('ARS_SALARIE')
              + ';AUXILIAIRE=' + GetField('ARS_AUXILIAIRE')
              + ';TIERS=' + GetField('ARS_TIERS')
              + ';LIBELLE=' + GetField('ARS_LIBELLE')
              + ';ADRESSE1=' + GetField('ARS_ADRESSE1')
              + ';ADRESSE2=' + GetField('ARS_ADRESSE2')
              + ';ADRESSE3=' + GetField('ARS_ADRESSE3')
              + ';CODEPOSTAL=' + GetField('ARS_CODEPOSTAL')
              + ';VILLE=' + GetField('ARS_VILLE')
              + ';PAYS=' + GetField('ARS_PAYS')
              + ';TAUXUNIT=' + FloatToStr(GetField('ARS_TAUXUNIT'));

            // On appelle la saisie du salarié en passant en paramètre le code salarié vide
            // pour qu'on soit positionné automatiquement en ajout.
            sCodeSalarie := AglLanceFiche ('PAY', 'SALARIE', '', '', sTmp);

            // On récupère le code salarié passé en retour de la saisie du salarié
            SetField ('ARS_SALARIE', sCodeSalarie);
            SetFocusControl ('ARS_SALARIE');
          end; *)// Suppr. TG 14/04/2003
      {-$ENDIF}// DEBUG

    if (tsArgsReception <> nil) then
        // Au retour d'une saisie de ressource à partir d'un autre écran
        // on rempli la zone Retour de la fiche avec le code Ressource que
        // l'on vient de créer, on libère la mémoire allouée et on sort de la
        // fiche
    begin
          // On adapte la sortie de l'écran suivant la provenance de l'appel
      if (tsArgsReception.Values['ORIGINE'] = 'PSA') or (tsArgsReception.Values['ORIGINE'] = 'RGP') then
      begin
        TFFiche(ecran).Retour := GetField('ARS_RESSOURCE');
        tsArgsReception.Free;
        tsArgsReception := nil;
        TFFicheListe(ecran).close;
      end;
    end;
  end;

  SetField('ARS_UTILISATEUR', V_PGI.User);

  //NCX 21/05/01
    // mcd 04/07/01 If LastError = 0 then
  if (EnSaisie) and (LastError = 0) then
  begin
      (*// ajout mcd 11/09/01
      if (GetField('ARS_SALARIE') <> '')
        and (ExisteSQL('SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_SALARIE="'+GetField('ARS_SALARIE')
          + '" AND ARS_RESSOURCE<>"'+GetControlText('ARS_RESSOURCE')+'"')) then
        begin
          PGIInfoAF ('Vous ne pouvez pas utiliser plusieurs fois le même code salarié,'+
                    'à chaque ressource doit correspondre un salarié différent', TitreHalley);
          SetFocusControl ('ARS_SALARIE');
          Lasterror := 18;
          exit;
        end;    // fin mcd 11/09 *)// Déplacé dans OnChangeField - TG 04/04/2003
        // mcd 11/06/03 message si même utilisateur
    if (GetField('ARS_UTILASSOCIE') <> '')
      and (ExisteSQL('SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_UTILASSOCIE="' + GetField('ARS_UTILASSOCIE')
      + '" AND ARS_RESSOURCE<>"' + GetControlText('ARS_RESSOURCE') + '"')) then
    begin
      PGIInfoAF('Il est déconseillé d''utiliser plusieurs fois le même code utilisateur associé,' +
        'à chaque ressource doit correspondre un utilisateur différent', TitreHalley);
    end; // fin mcd 11/06/03

    Modif := VerifChangements;
    if (modif = 1) or (modif = 2) or (modif = 3) then
    begin
      if modif = 1 then
      begin
        if not QuestHisto then Historisation;
        SetControlVisible('BREVALORISATION', True);
      end
      else
        if (modif = 2) and (QuestHisto = False) then
        begin
          Historisation;
          SetControlVisible('BREVALORISATION', True); //mc d19/04/02 recal aussi sur pV ...
        end
        else
          if (modif = 3) then
            SetControlVisible('BREVALORISATION', True);
    end;

      // If VarToDateTime(GetField('ARS_DATEPRIX')) = idate1900 then SetField('ARS_DATEPRIX',Date);
    SetControlVisible('BHISTOVALO', True);
  end;

  if (EnSaisie) and (ModifSpecif) and (TCheckBox(GetControl('CBSPECIF')).checked = True) then
    MajEnregCal('specif');

  if ModifStand then MajEnregCal('Stand');
  if ModifSalarie then MajEnregCal('Salarie');
  ModifSpecif := False;
  ModifStand := False;
  ModifSalarie := False;
  if (EnSaisie) then
  begin
    SetControlText('STANDTEMP', GetField('ARS_STANDCALEN'));
    if Assigned(GetControl('ARS_SALARIE')) then
      TempSalarie := GetControlText('ARS_SALARIE');
    SetcontrolEnabled('BPERSOCALENDRIER', true);

    if (GetCheckBoxState('ARS_CALENSPECIF') = cbChecked) then
    begin
      TCheckBox(GetControl('CBSPECIF')).Checked := true;
      SetControlEnabled('ARS_STANDCALEN', false);
    end
    else
    begin
      TCheckBox(GetControl('CBSPECIF')).Checked := false;
      SetControlEnabled('ARS_STANDCALEN', true);
    end;

    SetControlEnabled('B_DUPLICATION', True);
  end;

  QuestHisto := False;

  // mcd 21/05/02 gestion zones obligatoire
  if (Ensaisie) then
  begin
    NomChamp := VerifierChampsObligatoires(Ecran, 'ARS');
    if (NomChamp <> '') then
    begin
      NomChamp := ReadTokenSt(NomChamp);
      SetFocusControl(NomChamp);
      LastError := 22;
      LastErrorMsg := TexteMsgRessource[LastError] + champToLibelle(NomChamp);
    end;
  end;

  if ThEdit(GetControl('BRS_SECTION')).text <> '' then
  begin
    if not ExisteSql('SELECT 1 FROM SECTION WHERE S_SECTION="'+ThEdit(GetControl('BRS_SECTION')).text+'"') then
    begin
      SetFocusControl('BRS_SECTION');
      LastError := 32;
      LastErrorMsg := TexteMsgRessource[LastError];
      Exit;
    end;
  end;
  TOBRESSOURCEBTP.GetEcran(ecran);
  TOBRESSOURCEBTP.setString('BRS_RESSOURCE',GetField('ARS_RESSOURCE'));
  TOBRESSOURCEBTP.SetAllModifie(true);
  TOBRESSOURCEBTP.InsertOrUpdateDB(false);

  // Modif par lot : PL le 28/09/2001
  if ModifLot then TFFiche(ecran).BFermeClick(nil);
  if (EnSaisie) and (LastError = 0) then
  begin
    if (GetControl('ARS_SITE') <> nil) and (GetField('ARS_SITE') <> '') then
    begin
      if not LookupValueExist(GetControl('ARS_SITE')) then
      begin
        LastError := 27;
        LastErrorMsg := TexteMsgRessource[LastError];
        SetFocusControl('ARS_SITE');
      end;
    end;
  end;
end;

// PL le 05/02/03 : on remet un close qui va libérer la liste des arguments si besoin est...
// je ne sais pas pourquoi ça avait été supprimé...

procedure TOM_Ressource.OnClose;
begin
  if (tsArgsReception <> nil) then
  begin
    tsArgsReception.Free;
    tsArgsReception := nil;
  end;
  TOBRESSOURCEBTP.free;
  inherited;
end;

procedure TOM_Ressource.DeleteRessourceBTP (CodeRessource : string);
begin
  ExecuteSql('DELETE FROM BRESSOURCE WHERE BRS_RESSOURCE="'+CodeRessource+'"');
end;

procedure TOM_Ressource.FindRessourceBTP (CodeRessource : string);
var  QQ : TQuery;
begin
  TOBRESSOURCEBTP.InitValeurs(true);
  TOBRESSOURCEBTP.SetString('BRS_RESSOURCE',CodeRessource);
  if CodeRessource = '' then exit;
  QQ := OpenSql('SELECT * FROM BRESSOURCE WHERE BRS_RESSOURCE="'+CodeRessource+'"',true,1,'',true);
  if not QQ.eof then
  begin
    TOBRESSOURCEBTP.SelectDB('',QQ);
  end;
  ferme (QQ);
end;

procedure TOM_Ressource.OnLoadRecord;
var EvtOnChange	: TNotifyEvent;
    LibRessource: String;
    Nom         : String;
    Prenom      : String;
begin
  SetEvents (false);

  if TypeRessource = '' then TypeRessource := Getfield('ARS_TYPERESSOURCE');

  if (GetControl('ARS_TYPERESSOURCE')) <> nil then
		 Begin
     if TypeRessource <> '' then
        Begin
          CBOTYPERESSOURCE.Enabled := False;
        end
     else
        CBOTYPERESSOURCE.Enabled := True;
  	 end;

  //Détermination du Libellé Famille Ressource - LowerCase
	LibRessource := iif(TypeRessource <> '',RECHDOM('AFTTYPERESSOURCE', TypeRessource, true),'');

  //if (ctxGPAO in V_PGI.PGIContexte) and (Assigned(Ecran) and (Pos('WRESSOURCE', Ecran.Name) > 0)) then
  Ecran.Caption := 'Ressource ' + LibRessource + ' : ' + GetField('ARS_RESSOURCE');

  //FV1 : 25/07/2013 ==> FS#529 - AURUS : pb affichage écran si création de ressource depuis la zone "intervenant"
  if LibRessource <> '' then  TLabelRessource.Caption := LibRessource;

  //recherche du libellé salarié
  //FV1 : 12/01/2014 - FS#706 - NCN : Dans la fiche ressource, le nom du salarié associé ne s'affiche pas
  LibelleSalarie(GetField('ARS_SALARIE'),Nom, Prenom);
  SetControlText('NOMSALARIE', Prenom + ' ' + Nom);


//uniquement en line
//  TLabelType.Caption := LibRessource;

  EvtOnChange := nil;
  EnSaisie := True;
  QuestHisto := False;
  OldSalarie := GetField('ARS_SALARIE');
  TauxPR := GetField('ARS_TAUXREVIENTUN');
  SetControlVisible('BREVALORISATION', False);

  if CodeArticle <> nil then
     Begin
     EvtOnChange := CodeArticle.OnChange;
	   CodeArticle.OnChange := nil;
		 CodeArticle.Text := trim(copy(GetField('ARS_ARTICLE'), 1, 18));
		 CodeArticle.OnChange := EvtOnChange;
     end;

  SetControlText('STANDTEMP', GetField('ARS_STANDCALEN'));

  //Aiguillage pour Chargement (Load) écran Ressource ou Matériel
  //Non en line
	LoadFicheRessource;
{*
  if (TypeRessource = 'SAL') Or (TypeRessource = 'ST') Or (TypeRessource = 'INT') Then
     LoadFicheRessource
  else if (TypeRessource = 'OUT') Or (TypeRessource = 'MAT') Or (TypeRessource = 'LOC') then
		 LoadFicheMateriel;
*}

  TempSalarie := GetField('ARS_SALARIE');
  PremierPassage := False;

  AppliquerConfidentialite(Ecran, 'ARS'); // mcd 21/05/02   zones soumises à confid

  // Modif par lot
  if ModifLot then SetArguments(StSQL);

  // Lien ressource-salarié
  CheckRattachementSalarie;

  //lien sousFamille Ressource
   if (GetControl('ARS_CHAINEORDO') <> nil) then
     Begin
	   SetControlText('ARS_CHAINEORDO',  GetField('ARS_CHAINEORDO'));
  	 LectureSFamilleRes(GetControlText('ARS_CHAINEORDO'));
     end;

  //FV1 : 04/05/2017 - FS#2528 - FVE - Gérer le SAV sur les Intérim comme sur les Salariés
  if (GetControlText('ARS_TYPERESSOURCE')='SAL') OR (GetControlText('ARS_TYPERESSOURCE')='INT') then
  begin
    SetControlVisible('BRS_GERESAV',True);
    SetControlVisible('BRS_GERECHANTIER',True);
    SetControlVisible('ARS_RESSOURCELIE',True);
  end else
  begin
    SetControlVisible('BRS_GERESAV',False);
    SetControlVisible('BRS_GERECHANTIER',false);
    SetControlVisible('ARS_RESSOURCELIE',false);
  end;

  //
  FindRessourceBTP (GetField('ARS_RESSOURCE'));
  TOBRESSOURCEBTP.PutEcran(ecran); 

  SetEvents (true);
end;


procedure TOM_Ressource.ChangeInfoSUp (Sender : Tobject);
begin
  if not (DS.State in [dsInsert, dsEdit]) then
  begin
    DS.edit;
    SetControlEnabled('bValider',true);
    SetControlEnabled('bRechercher',false);
    SetControlEnabled('bFirst',false);
    SetControlEnabled('bPrev',false);
    SetControlEnabled('bNext',false);
    SetControlEnabled('bLast',false);
    SetControlEnabled('bDefaire',true);
    SetControlEnabled('bInsert',false);
    SetControlEnabled('bDelete',false);
    SetControlEnabled('bDupliquer',false);
  end;
end;


procedure TOM_Ressource.SetEvents (Etat : boolean);
begin
  if Etat then
  begin
    TCHeckBox(GetControl ('BRS_GERESAV')).onClick := ChangeInfoSUp;
    TCHeckBox(GetControl ('BRS_GERECHANTIER')).onClick := ChangeInfoSUp;
    THEdit(GetControl ('BRS_SECTION')).OnChange := ChangeInfoSUp;
  end else
  begin
    TCHeckBox(GetControl ('BRS_GERESAV')).onClick := nil;
    TCHeckBox(GetControl ('BRS_GERECHANTIER')).onClick := nil;
    THEdit(GetControl ('BRS_SECTION')).OnChange := nil;
  end;
end;

procedure TOM_Ressource.LoadFicheRessource;
begin

	SetActiveTabSheet('PGeneral');

  if AdresseMail.Text <> '' then
  	 BTMAIL.Enabled := True
  Else
		 BTMAil.Enabled := False;

  //Gestion du calendrier spécifique ou standard
  if (CHKCALENSPECIF.Checked=False) And (CBOSTANDCALEND.Value = '') then
     BTPersoCalend.Enabled := False
  else
     BTPersoCalend.Enabled := True;

 {if (GetCheckBoxState('ARS_CALENSPECIF') = cbChecked) then
     begin
     TCheckBox(GetControl('CBSPECIF')).Checked := true;
     SetControlEnabled('ARS_STANDCALEN', false);
     end
  else
     begin
     TCheckBox(GetControl('CBSPECIF')).Checked := false;
     SetControlEnabled('ARS_STANDCALEN', true); //mcd 16/04/02 sinon pas ok sur bouton magnétoscope
     end;}

end;

procedure TOM_Ressource.LoadFicheMateriel;
begin

end;

procedure TOM_Ressource.OnCancelRecord;
begin
  inherited;

  if GetField('ARS_DEPOT') <> '' then
    SetControlProperty('ARS_SITE','PLUS','QSI_DEPOT="'+GetField('ARS_DEPOT')+'"')
  else
    SetControlProperty('ARS_SITE','PLUS','');

  SetEvents (false);
  FindRessourceBTP (GetField('ARS_RESSOURCE'));
  TOBRESSOURCEBTP.PutEcran(ecran);
  SetEvents (true);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 12/01/2000
Modifié le ... :   /  /
Description .. : test
Mots clefs ... :
*****************************************************************}

function TOM_Ressource.TestImpactChangeSalarie(CodeSalarie, CodeRessource: string): Boolean;
var
  CompWhere: string;
begin
  Result := False;

  CompWhere := 'AND ACR_RESSOURCE ="' + CodeRessource + '"';

  // Contrôle de la présence d'enregistrements dans les tables ou code resssource+Salarié
  if SupTablesLiees('COMPETRESSOURCE', 'ACR_SALARIE', CodeSalarie, compWhere, False) then
  begin
    PGIBoxAF(TexteMsgRessource[5], '');
    Exit;
  end;

  CompWhere := 'AND ACA_RESSOURCE ="' + CodeRessource + '"';
  if SupTablesLiees('CALENDRIER', 'ACA_SALARIE', CodeSalarie, compWhere, False) then
  begin
    PGIBoxAF(TexteMsgRessource[6], '');
    Exit;
  end;

  CompWhere := 'AND PCN_RESSOURCE ="' + CodeRessource + '"';
  if SupTablesLiees('ABSENCESALARIE', 'PCN_SALARIE', CodeSalarie, compWhere, False) then
  begin
    PGIBoxAF(TexteMsgRessource[7], '');
    Exit;
  end;

  Result := True;
end;

procedure TOM_Ressource.MoveFonction(Sender: TObject);
begin

  if not (DS.State in [dsInsert, dsEdit]) then DS.edit;

  if THEdit(Sender).Tag = 1 then
  begin // Move en Haut
    SetField('ARS_FONCTION1', GetControlText('ARS_FONCTION2'));
    SetField('ARS_DATEFONC1', GetControlText('ARS_DATEFONC2'));
    SetField('ARS_FONCTION2', GetControlText('ARS_FONCTION3'));
    SetField('ARS_DATEFONC2', GetControlText('ARS_DATEFONC3'));
    SetField('ARS_FONCTION3', GetControlText('ARS_FONCTION4'));
    SetField('ARS_DATEFONC3', GetControlText('ARS_DATEFONC4'));
    SetField('ARS_FONCTION4', '');
    SetField('ARS_DATEFONC4', idate1900);
  end
  else
  begin // Move en bas
    SetField('ARS_FONCTION4', GetControlText('ARS_FONCTION3'));
    SetField('ARS_DATEFONC4', GetControlText('ARS_DATEFONC3'));
    SetField('ARS_FONCTION3', GetControlText('ARS_FONCTION2'));
    SetField('ARS_DATEFONC3', GetControlText('ARS_DATEFONC2'));
    SetField('ARS_FONCTION2', GetControlText('ARS_FONCTION1'));
    SetField('ARS_DATEFONC2', GetControlText('ARS_DATEFONC1'));
    SetField('ARS_FONCTION1', '');
    SetField('ARS_DATEFONC1', idate1900);
  end;

  SetFocusControl('ARS_DATEFONC1'); //pour test que la fonction actuelle est correcte

end;

procedure TOM_Ressource.VerifFonction(Sender: TObject);
var
  i, NumFonc: integer;
begin
  NumFonc := THEdit(Sender).Tag;
  if GetField('ARS_DATEFONC' + IntToStr(NumFonc)) = Null then
    Exit;
  if (VarToDateTime(GetField('ARS_DATEFONC' + IntToStr(NumFonc))) = iDate1900) then
    Exit;

  if NumFonc = 1 then
  begin
    for i := 2 to 4 do
    begin
      if (VarToDateTime(GetField('ARS_DATEFONC' + IntToStr(i))) <> iDate1900)
        and (VarToDateTime(GetField('ARS_DATEFONC' + IntToStr(i))) > VarToDateTime(GetField('ARS_DATEFONC1'))) then
      begin
        PGIBoxAF(TexteMsgRessource[8], '');
        SetFocusControl('ARS_DATEFONC1');
        Exit;
      end;
    end;
  end
  else
  begin
    if (VarToDateTime(GetField('ARS_DATEFONC1')) <> iDate1900)
      and (VarToDateTime(GetField('ARS_DATEFONC1')) < VarToDateTime(GetField('ARS_DATEFONC' + IntToStr(NumFonc)))) then
    begin
      PGIBoxAF(TexteMsgRessource[9], '');
      SetFocusControl('ARS_DATEFONC' + IntToStr(NumFonc));
      Exit;
    end;
  end;

end;

{$IFDEF AFFAIRE}

procedure TOM_Ressource.Brevalorisationclick;
var
  T: TOB;
  Tofreval: TOF_REVAL_RESS_MUL;
  ret: integer;
begin
  ret := 0;
  T := TOB.Create('', nil, -1);
  T.LoadDetailDB('RESSOURCE', '"' + GetField('ARS_RESSOURCE') + '"', '', nil, False);
  Tofreval := TOF_REVAL_RESS_MUL(CreateTof('REVAL_RESS_MUL', nil, False, False));
  if TofReval <> nil then
    ret := Tofreval.RevaloriseRessources(T);
  if (ret <> 0) then PgiInfoAf('Problème lors de la revalorisation', 'Revalorisation');

  Tofreval.free;
  T.free;
end;
{$ENDIF}

//NCX 16/05/01

function TOM_Ressource.VerifChangements: integer;
var
  Q: Tquery;
  listeChamps, Champs: string;
  modif, modmodif: integer;
begin
  Result := 0;
  modif := 0;
  listeChamps := 'ARS_TAUXUNIT,ARS_TAUXFRAISGEN1,ARS_COEFMETIER,ARS_TAUXCHARGEPAT,' +
    'ARS_TAUXREVIENTUN,ARS_UNITETEMPS,ARS_TAUXFRAISGEN2,ARS_PVHT,ARS_PVTTC,' +
    'ARS_DATEPRIX,ARS_COEFPRPV,ARS_PVHTCALCUL';

  Q := OpenSQL('SELECT ' + ListeChamps + ' FROM RESSOURCE WHERE ARS_RESSOURCE="' + GetField('ARS_RESSOURCE') + '"', True);
  if Q.EOF then
  begin
    Ferme(Q);
    exit;
  end;

  while (listeChamps <> '') do
  begin
    Champs := ReadTokenPipe(listeChamps, ',');
    if Q.FindField(Champs).asVariant <> GetField(Champs) then
    begin
      if (Champs = 'ARS_DATEPRIX') or (Champs = 'ARS_COEFPRV') or (Champs = 'ARS_PVHTCALCUL')
        or (Champs = 'ARS_PVHT') or (Champs = 'ARS_PVTTC') then
        modif := modif + 100
      else
        if (Champs = 'ARS_TAUXREVIENTUN') then
          modif := modif + 101
{$IFDEF AFFAIRE}
        else
        if (GetParamSoc ('SO_AFCLIENT') = cInClientAlgoe) and (Champs = 'ARS_TAUXUNIT') then
          modif := modif + 101
{$ENDIF}
        else
          modif := modif + 1;
    end;
  end;
  Ferme(Q);

  //
  if modif = 0 then
  begin
    result := modif;
    exit;
  end; // Rien n'a bougé

  modmodif := modif mod 100;
  if modmodif = 0 then
  begin
    result := 2;
    exit;
  end; // Seules les données relatives au PV ont changées

  if (modmodif < 100) and (modif < 100) then
  begin
    result := 3;
    exit;
  end; // Seules les donnes relatives au PR ont changées

  result := 1; // Les données du PV et du Pr ont changées.

end;

procedure TOM_Ressource.Historisation;
var
  Champ, ListeChamps, ListeChampTot: string;
  Q: Tquery;
begin
  if (TFfiche(Ecran).typeaction <> TaConsult) then
  begin
    if PgiAskAF('Voulez vous conserver un historique des prix tels qu''ils étaient avant modification', TitreHalley) = MrYes then
    begin
      if GetField('ARS_DATEPRIX') > GetField('ARS_DATEPRIX2') then
      begin
        QuestHisto := True;
        TheTOB := Tob.create('RESSOURCE', nil, -1);
        try
              // mcd si modif voir fct ModifAvectrait de UtofRessou_Modiflot
          ListeChamps := 'ARS_TAUXUNIT,ARS_DATEPRIX,ARS_TAUXREVIENTUN,ARS_PVHT,ARS_PVTTC,ARS_PVHTCALCUL';
          ListeChampTot := ChargeTheTobHisto(listeChamps, true);

              // obligation d'aller de faire une requête pour avoir les valeurs avant chgmt écran
          Q := OpenSQL('SELECT ' + ListeChamps + ' FROM RESSOURCE WHERE ARS_RESSOURCE="' + GetField('ARS_RESSOURCE') + '"', True);
          while ListeChamps <> '' do
          begin
            Champ := ReadTokenPipe(ListeChamps, ',');
            if champ = 'ARS_DATEPRIX' then
              TheTob.putValue(champ + IntToStr(2), Q.FindField(champ).asDateTime)
            else
              TheTob.putValue(champ + IntToStr(2), Q.FindField(champ).asFloat);
            ListeChampTot := ListeChampTot + champ + IntToStr(2) + ';';
          end;

          Ferme(Q);
          ListeChampTot := Copy(ListeChampTot, 1, Length(ListeChampTot) - 1);
          MajHistovalo(ListeChampTot);
        finally
          TheTob.free;
          TheTob := nil;
        end;
      end
      else
        Pgiinfo('Vous ne pouvez historiser avec deux fois la même date', Titrehalley);
    end;

  end;

end;

procedure TOM_Ressource.MajHistovalo(parms: string);
var
  champ: string;
begin
  if Ds.State <> DsEdit then Ds.Edit;
  while parms <> '' do
  begin
    Champ := ReadTokenSt(Parms);
    SetField(champ, TheTob.GetValue(Champ));
  end;

  if QuestHisto <> True then TFFiche(ecran).Bouge(nbPost);

end;

procedure TOM_Ressource.BHistoValoclick(sender: TObject);
var
  parms, ListeChamps, ListeChampTot: string;
begin
  if (Ds.State <> DsEdit) or (QuestHisto = True) then
  begin
    try
      TheTOB := Tob.create('RESSOURCE', nil, -1);
      ListeChamps := 'ARS_TAUXUNIT,ARS_DATEPRIX,ARS_TAUXREVIENTUN,ARS_PVHT,ARS_PVTTC,ARS_PVHTCALCUL';
      begin
        ListeChampTot := ChargeTheTobHisto(listeChamps, False);
        ListeChampTot := Copy(ListeChampTot, 1, Length(ListeChampTot) - 1);
{$IFDEF AFFAIRE}
        if (TFFiche(ecran).TypeAction = taConsult) then
            // mcd 12/06/02 then  Parms := AGLLanceFiche('AFF','RESHISTOVALO','','','ACTION=CONSULTATION') else
          Parms := AFLanceFiche_RevalActivPAram('ACTION=CONSULTATION')
        else
          Parms := AFLanceFiche_RevalActivParam('');
{$ENDIF}

          // mcd 12/06/02Parms := AGLLanceFiche('AFF','RESHISTOVALO','','','');
        if Parms = '0' then MajHistovalo(ListeChampTot);
      end;
    finally
      TheTob.free;
      TheTob := nil;
    end;
    QuestHisto := False;
  end
  else
  begin
    PGIInfoAF('Pour avoir accès à l''historique vous devez d''abord valider vos modifications', TitreHalley);
  end;
end;

procedure TOM_Ressource.MajEnregCal(Quoi: string);
var
  tableMAJ, TableQ, Prefixe, rqsalarie, rqstand, rqspe, wr1, wr2, fr, Csalarie, Cressource: string;
begin
  TableMaj := 'ACA=CALENDRIER;ACG=CALENDRIERREGLE';
  while TableMaj <> '' do
  begin
    prefixe := ReadTokenPipe(TableMaj, '=');
    TableQ := ReadTokenSt(TableMaj);
    Csalarie := GetField('ARS_SALARIE');
    if Csalarie = '' then Csalarie := '***';
    Cressource := GetField('ARS_RESSOURCE');
    if Cressource = '' then Cressource := '***';
    rqstand := 'DELETE ';
    rqsalarie := 'UPDATE ' + TableQ + ' SET ' + prefixe + '_SALARIE="' + Csalarie + '" ';
    fr := 'FROM ' + TableQ;
    wr1 := ' WHERE ' + prefixe + '_RESSOURCE="' + Cressource + '"';
    wr2 := ' AND ' + prefixe + '_SALARIE="' + Csalarie + '"';
    rqspe := ' AND ' + prefixe + '_STANDCALEN="' + GetControlText('STANDTEMP') + '"';

    if quoi = 'specif' then ExecuteSql(rqstand + fr + wr1 + wr2 + rqspe) else
      if quoi = 'stand' then ExecuteSql(rqstand + fr + wr1 + wr2) else

        ExecuteSql(rqsalarie + wr1);
  end;
end;

procedure TOM_Ressource.CBSpecifClick(Sender: TObject);
begin

  if CHKCalenSpecif.Checked then
     BTPERSOCALEND.Enabled := true
  else
     BTPERSOCALEND.Enabled := true;

end;

procedure TOM_Ressource.StandCalenClick(Sender: TObject);
begin

  //ModifStand := GetControlText('ARS_STANDCALEN') <> GetControlText('STANDTEMP');

  if CboStandCalend.value = '' then
     BtPersoCalend.Enabled := False
  Else
     BtPersoCalend.Enabled := True;
     
  //if (not ModifStand) and (not Modifspecif) then
  //  SetcontrolEnabled('BPERSOCALENDRIER', True)
  //else
  //  SetcontrolEnabled('BPERSOCALENDRIER', False);

end;

function TOM_Ressource.ChargeTheTobHisto(liste: string; historisation: boolean): string;
var
  ListeChampTot, Champ: string;
  i, min: integer;
begin
  while Liste <> '' do
  begin
    Champ := ReadTokenPipe(Liste, ',');
    if historisation then
      min := 3
    else
      min := 2;

    for i := 4 downto min do
    begin
      if historisation then
        TheTob.putValue(champ + IntToStr(i), GetField(champ + IntToStr(i - 1)))
      else
        TheTob.putValue(champ + IntToStr(i), GetField(champ + IntToStr(i)));

      ListeChampTot := ListeChampTot + champ + IntToStr(i) + ';';
    end;
  end;

  result := ListeChampTot;
end;

procedure TOM_Ressource.ZoneSalarieOnExit(Sender: TObject);
Var Nom     : String;
    Prenom  : String;
begin

 {mcd 11/09/01 fait trop svt. mis dans onupdate If (GetField('ARS_SALARIE')<>'')
 and (ExisteSQL('SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_SALARIE="'+GetField('ARS_SALARIE')+
          '" AND ARS_RESSOURCE<>"'+GetControlText('ARS_RESSOURCE')+'"')) then
   begin
    PGIInfoAF('Vous ne pouvez pas utiliser plusieurs fois le même code salarié,'+
              'à chaque code ressource doit correspondre un code salarié différent',TitreHalley);
    SetFocusControl('ARS_SALARIE');
    exit;
   end;}

  if GetField('ARS_SALARIE') <> TempSalarie then
    ModifSalarie := True
  else
    SetcontrolEnabled('BPERSOCALENDRIER', True);

  //recherche du libellé salarié
  //FV1 : 12/01/2014 - FS#706 - NCN : Dans la fiche ressource, le nom du salarié associé ne s'affiche pas
  LibelleSalarie(GetField('ARS_SALARIE'),Nom, Prenom);
  SetControlText('NOMSALARIE', Prenom + ' ' + Nom);


end;

procedure TOM_Ressource.ZoneSalarieOnEnter(Sender: TObject);
begin
  SetcontrolEnabled('BPERSOCALENDRIER', False);
end;

Procedure TOM_Ressource.SsFamilleOnElipsisClick(Sender : TObject);
Begin

  DS.Edit;

  SetControlProperty('ARS_CHAINEORDO', 'Plus', '');

  LookupCombo(GetControl('ARS_CHAINEORDO'));

  if GetControlText('ARS_CHAINEORDO')='' then exit;

  //Lecture de l'adresse sélectionnée et affichage des informations
  LectureSFamilleRes(GetControlText('ARS_CHAINEORDO'));

end;

procedure TOM_Ressource.SsFamilleOnExit(Sender: TObject);
begin

  LectureSFamilleRes(GetControlText('ARS_CHAINEORDO'));

end;


//Lecture de la Sous-famille ressource (BTTYPRES)
procedure TOM_Ressource.LectureSFamilleRes(SFamRes: string);
var Req       : String;
    TobTypRes : Tob;
begin

  if SFamRes = '' then exit;

  SetControlText('LARS_SSFAM', '');

  //Lecture du Type ACtion Evenement sélectionné et affichage des informations
  req := 'SELECT BTR_LIBELLE FROM BTTYPERES WHERE BTR_TYPRES="' + SFamRes + '"';
  Req := Req + ' AND BTR_FAMRES="' + GetControlText('ARS_TYPERESSOURCE') + '"';

  TobTypRes := Tob.Create('#BTTYPRES',Nil, -1);
  TobTypRes.LoadDetailFromSQL(Req);
  if TobTypRes.Detail.Count > 0 then
     Setcontroltext('LARS_SSFAM', TobTypRes.detail[0].GetString('BTR_LIBELLE'))
  else
     Begin
     Setcontroltext('ARS_CHAINEORDO', '');
     Setcontroltext('LARS_SSFAM', '');
     end;

  TobTypRes.free;

end;

procedure TOM_Ressource.Duplication(CleDuplic: string);
var
  TobForm: TOB;
  i: integer;
  StTable, StCleDuplic: string;
//    TagFicheRech : integer ;
  CC: Tcontrol;
  G_CodeRess: THCritMaskEdit;
  CC_RO: Boolean;
  QQ: TQuery;
begin
  //Duplic:= True ;
  StCleDuplic := CleDuplic;
  // PL le 22/01/02 c'est toujours la table ressource
  StTable := 'RESSOURCE';
  // Fin PL le 22/01/02
  // TagFicheRech := 1; // pour DispatchRecherche
  TobForm := TOB.Create(StTable, nil, -1);
  if TobForm <> nil then
  begin
    if STCleDuplic = '' then
    begin
      if (DS.State = dsInsert) then // si nouvelle fiche  recherche record à dupliquer
      begin
        G_CodeRess := THCritMaskEdit.Create(nil);
        DispatchRecherche(G_CodeRess, 3, '', '', '');
        StCleDuplic := G_CodeRess.Text;
                        // normal de tout prendre, on duplique ...
        QQ := OpenSQL('SELECT * From RESSOURCE WHERE ARS_RESSOURCE="' + StCleDuplic + '"', True);
        if not TobForm.SelectDB('', QQ) then
          StCleDuplic := '';

        Ferme(QQ);
        G_CodeREss.Free;
      end
      else
      begin
        if not TobForm.SelectDB('', TFFiche(Ecran).QFiche) then
          StCleDuplic := '';
        stCleDuplic := '*';
        TFFiche(Ecran).Bouge(NbInsert);
      end;
    end
    else
    begin
      if not TobForm.SelectDB('"' + StCleDuplic + '"', nil) then
        StCleDuplic := '';
    end;

      // on efface calendrier spécif si existe et salarié
    TobForm.PutValue('ARS_DATECREATION', V_PGI.DateEntree);
    TobForm.PutValue('ARS_SALARIE', '');
    TobForm.PutValue('ARS_UTILASSOCIE', '');  //mcd 30/09/03
    TobForm.PutValue('ARS_RESSOURCE', ''); //mcd 01/07/2003 il faut effacer le code, sinon, peut poser des pb sur saisie code salaréi qui existe déja en lien paie
    TobForm.PutValue('ARS_CALENSPECIF', '-');
    if StCleDuplic <> '' then
    begin
      for i := 1 to TobForm.NbChamps do
      begin
        CC := TControl(Ecran.findcomponent(TobForm.GetNomChamp(i)));
{$IFDEF EAGLCLIENT}
        CC_RO := ((CC is TCustomEdit) and (TEdit(CC).ReadOnly));
{$ELSE}
        CC_RO := ((CC is TEdit) and (TEdit(CC).ReadOnly)) or ((CC is TDBEdit) and (TDBEdit(CC).ReadOnly));
{$ENDIF}
        if (CC <> nil) and (CC.enabled and (not CC_RO) and cc.Visible) then SetField(TobForm.GetNomChamp(i), TobForm.GetValeur(i));
      end;

          // les champs histo ne font pas parti de la fiche. il faut les mettre à jour
      for i := 2 to 4 do
      begin
        SetField('ARS_DATEPRIX' + InttoStr(i), TobForm.GetValue('ARS_DATEPRIX' + InttoStr(i)));
        SetField('ARS_TAUXREVIENTUN' + InttoStr(i), TobForm.GetValue('ARS_TAUXREVIENTUN' + InttoStr(i)));
        SetField('ARS_PVHT' + InttoStr(i), TobForm.GetValue('ARS_PVHT' + InttoStr(i)));
        SetField('ARS_PVTTC' + InttoStr(i), TobForm.GetValue('ARS_PVTTC' + InttoStr(i)));
        SetField('ARS_PVHTCALCUL' + InttoStr(i), TobForm.GetValue('ARS_PVHTCALCUL' + InttoStr(i)));
      end;
          // TobForm.PutEcran (F); A mettre quand fonctionnement ok (bloc note affiché)
    end;
  end; // TobForm <> Nil

  TobForm.Free;
  SetActiveTabSheet(TFFiche(Ecran).pages.Pages[0].name);

  SetFocusControl('ARS_RESSOURCE');
  SetControlEnabled('B_DUPLICATION', false);
  SetControlEnabled('ARS_STANDCALEN', True);

end;


procedure TOM_Ressource.CheckRattachementSalarie;
var i, j: integer;
  PasTouche: boolean;
  Q: TQuery;
  fieldz: string;
begin
  if not GetParamSoc('SO_PGLIENRESSOURCE') then exit;

  PasTouche := (GetField('ARS_SALARIE') <> '');
  fieldz := '';

  for i := Low(RSLinkedFields) to High(RSLinkedFields) do
  begin
    SetControlEnabled(RSLinkedFields[i, 0], not PasTouche);
    if fieldz <> '' then fieldz := fieldz + ', ';
    fieldz := fieldz + RSLinkedFields[i, 1];
  end;

  Q := OpenSQL('SELECT ' + fieldz + ' FROM SALARIES WHERE PSA_SALARIE="' + OldSalarie + '"', true);
  if not Q.EOF then
  begin
    for i := Low(RSLinkedFields) to High(RSLinkedFields) do
      if GetField(RSLinkedFields[i, 0]) <> Q.FindField(RSLinkedFields[i, 1]).AsString then
      begin
        PGIInfoAF(TexteMsgRessource[23], '');
        ForceUpdate;
        for j := Low(RSLinkedFields) to High(RSLinkedFields) do
          SetField(RSLinkedFields[j, 0], Q.FindField(RSLinkedFields[j, 1]).AsVariant);
        break;
      end;
  end;
  Ferme(Q);
end;


/////////////////////////////////////////////////////////////////////////////

procedure TOM_Ressource.SetArguments(StSQL: string);
var
  Critere, ChampMul, ValMul: string;
  x, y: integer;
  Ctrl: TControl;
  Fiche: TFFiche;
begin
  SetControlVisible('BSTOP', TRUE);
  DS.Edit;
  Fiche := TFFiche(ecran);
  repeat
    Critere := AnsiUppercase(Trim(ReadTokenPipe(StSQL, '|')));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        y := pos(',', ValMul);
        if y <> 0 then
          ValMul := copy(ValMul, 1, length(ValMul) - 1);
        if copy(ValMul, 1, 1) = '"' then
          ValMul := copy(ValMul, 2, length(ValMul));
        if copy(ValMul, length(ValMul), 1) = '"' then
          ValMul := copy(ValMul, 1, length(ValMul) - 1);
        SetField(ChampMul, ValMul);
        Ctrl := TControl(Fiche.FindComponent(ChampMul));
        if Ctrl = nil then
          exit;
{$IFDEF EAGLCLIENT}
        if (Ctrl is TCustomCheckBox) or (Ctrl is THValComboBox) or (Ctrl is TCustomEdit) then
          TEdit(Ctrl).Font.Color := clRed
        else
          if Ctrl is TSpinEdit then
            TSpinEdit(Ctrl).Font.Color := clRed
          else
            if (Ctrl is TCheckBox) or (Ctrl is THValComboBox) or (Ctrl is THEdit) or (Ctrl is THNumEdit) then
            begin
              TSpinEdit(Ctrl).Font.Color := clRed;
              SetControlText(ChampMul, ValMul);
            end;
{$ELSE}
        if (Ctrl is TDBCheckBox) or (Ctrl is THDBValComboBox) or (Ctrl is THDBEdit) then
          TEdit(Ctrl).Font.Color := clRed
        else
          if Ctrl is THDBSpinEdit then
            THDBSpinEdit(Ctrl).Font.Color := clRed
          else
            if (Ctrl is TCheckBox) or (Ctrl is THValComboBox) or (Ctrl is THEdit) or (Ctrl is THNumEdit) then
            begin
              THDBSpinEdit(Ctrl).Font.Color := clRed;
              SetControlText(ChampMul, ValMul);
            end;
{$ENDIF}
      end;
    end;
  until (Critere = '');

end;

{$IFDEF GPAO}
procedure TOM_RESSOURCE.BRESSOURCELIE_OnClick(Sender: tObject);
begin
  AglLanceFiche('WBT', 'WRESSOURCEVIS_VIE', '', '', 'RESSOURCE=' + GetField('ARS_RESSOURCE'));
end;

procedure TOM_Ressource.BImprimer_OnClick(Sender: tObject);
var
  Params: string;
begin
  Params := 'ARS_RESSOURCE=' + GetField('ARS_RESSOURCE');
  AGLLanceFiche('W', 'WR2_QR1', '', '', Params);
end;

function TOM_Ressource.IsEcranGPAO: Boolean;
begin
  Result := Assigned(Ecran) and ((Ecran.Name = 'WRESSOURCE_FIC') or (Ecran.Name = 'WRESSOURCEPER_FIC'));
end;
{$ENDIF}

//Envoie du mail par appuie sur le bouton
Procedure Tom_Ressource.MailClick(Sender: TObject);
Var ResMail : TResultMailForm;
    AdrMail : String;
    Sujet		: HString;
    Copie		: String;
    Fichier	: String;
    Liste		: HTStringList;
Begin

	AdrMail := AdresseMail.text;
  Sujet		:= '';
  Copie		:= '';
  Fichier	:= '';

  Liste := HTStringList.Create;

  TRY
		 ResMail := AglMailForm(Sujet,AdrMail,Copie,Liste,Fichier,false);
		 if ResMail = rmfOkButNotSend then SendMail(Sujet,AdrMail,'',Liste,'',True,1, '', '', True);
  FINALLY
     Liste.Free;
  END;


end;

Procedure Tom_Ressource.MailChange(Sender: TObject);
Begin

  if AdresseMail.Text <> '' then
  	 BTMAIL.Enabled := True
  Else
		 BTMAil.Enabled := False;

end;

Procedure Tom_Ressource.MailDblClick(Sender: TObject);
Begin

	if AdresseMail.text <> '' then MailClick(Sender);

end;

procedure Tom_Ressource.BtSalarieClick(Sender: TObject);
Var Lequel   : String;
    Argument : String;
begin

  Argument := ActionToString(TFfiche(Ecran).TypeAction) + ';';
  Argument := Argument + '"ORIGINE=ARS;SALARIE=' + GetField('ARS_SALARIE');
  Argument := Argument + ';AUXILIAIRE=' + GetField('ARS_AUXILIAIRE');
  Argument := Argument + ';TIERS=' + GetField('ARS_TIERS');
  Argument := Argument + ';LIBELLE=' + GetField('ARS_LIBELLE');
  Argument := Argument + ';ADRESSE1=' + GetField('ARS_ADRESSE1');
  Argument := Argument + ';ADRESSE2=' + GetField('ARS_ADRESSE2');
  Argument := Argument + ';ADRESSE3=' + GetField('ARS_ADRESSE3');
  Argument := Argument + ';CODEPOSTAL=' + GetField('ARS_CODEPOSTAL');
  Argument := Argument + ';VILLE=' + GetField('ARS_VILLE');
  Argument := Argument + ';PAYS=' + GetField('ARS_PAYS');
	Argument := Argument + ';TAUXUNIT=' + GetField('ARS_TAUXUNIT');

  Lequel := GetField('ARS_SALARIE');

	if (GetField('ARS_SALARIE') <> '') then AGLLAnceFiche ('PAY','SALARIE','',Lequel,Argument);

end;


procedure Tom_Ressource.ControleAdresse(Sender: TObject);
begin          

  VerifCodePostal(nil,CodePostal,Ville,(THEdit(Sender).name = 'ARS_CODEPOSTAL'));

  SetField('T_CODEPOSTAL', GetControlText('T_CODEPOSTAL'));
  SetField('T_VILLE', GetControlText('T_VILLE'));

end;

Procedure Tom_Ressource.BtPersoCalendClick(Sender: TObject);
Var Lequel   : String;
    Argument : String;
Begin

  if CHKCALENSPECIF.Checked then
   BEGIN
   //ACTIONDB('POST');
   Argument := ActionToString(TFfiche(Ecran).TypeAction) + ';';
   Argument := Argument + 'TYPE=RES;CODE='+ GetField('ARS_RESSOURCE');
   Argument := Argument + ';LIBELLE='+ GetField('ARS_LIBELLE');
   Argument := Argument + ';STANDARD='+ CBOSTANDCALEND.Value;
   AGLLanceFiche('BTP', 'BTHORAIRESTD', '', Lequel, Argument);
   END
else
   BEGIN
   Argument := ActionToString(TFfiche(Ecran).TypeAction) + ';';
   Argument := Argument + 'TYPE=STD;STANDARD=' + CBOSTANDCALEND.Value;
   AGLLanceFiche('BTP', 'BTHORAIRESTD', '', Lequel, Argument);
   END;

end;

procedure AGLDuplication_RES(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_RESSOURCE) then TOM_Ressource(OM).Duplication('') else exit;
end;

procedure AFLanceFiche_Ressource(lequel, argument: string);
begin
  AGLLanceFiche('AFF', 'RESSOURCE', '', Lequel, Argument);
end;

function TOM_Ressource.Ressource_CalculPR(): double;
var TauxUnit,
    TauxFr1,
    TauxFr2,
    TauxChargePat,
    CoefMetier: Double;
begin

  if ChkCalculPR <> nil then
     Begin
     if CHKCALCULPR.Checked then
        SetControlEnabled('ARS_TAUXREVIENTUN', False)
     else
        Begin
        SetControlEnabled('ARS_TAUXREVIENTUN', True);
        exit;
        end;
     end;

  TauxUnit := (GetField('ARS_TAUXUNIT'));
  TauxFr1 := (GetField('ARS_TAUXFRAISGEN1'));
  TauxFr2 := (GetField('ARS_TAUXFRAISGEN2'));
  TauxChargePat := (GetField('ARS_TAUXCHARGEPAT'));
  CoefMetier := (GetField('ARS_COEFMETIER'));

  CheckError;

//uniquement en line
{*  if TauxFr1 = 0 then TauxFr1 := 1;
  if TauxFr2 = 0 then TauxFr2 := 1;
  if TauxChargePat = 0 then TauxChargePat := 1;
  if CoefMetier = 0    then CoefMetier := 1;
*}

  // si modif voir fct modifAvec_trait de UtofAfRessou_ModifLot
  result := (TauxUnit * TauxFr1 * TauxFr2 * TauxChargePat * coefMetier);

  // mcd 07/11/02 ajout arrondi
  Result := Arrondi(result, GetParamSoc('SO_DECPRIX'));

end;

//NCX 18/05/01

function TOM_Ressource.Ressource_CalculPV(): double;
var Coef				: Double;
		TauxRevient	: Double;
begin

  if (GetField('ARS_CALCULPV') = 'X') then
     SetControlEnabled('ARS_PVHTCALCUL', False)
  else
     Begin
     SetControlEnabled('ARS_PVHTCALCUL', True);
     exit;
     end;

  Coef := (GetField('ARS_COEFPRPV'));
  TauxRevient := (GetField('ARS_TAUXREVIENTUN'));

  // si modif voir fct modifAvec_trait de UtofAfRessou_ModifLot
  // mcd 07/11/02 ajout arrondi
  result := Arrondi((TauxRevient * Coef), GetParamSoc('SO_DECPRIX'));

end;


initialization
  registerclasses([TOM_Ressource]);
  RegisterAglProc('Duplication_Res', TRUE, 0, AGLDuplication_REs);
end.

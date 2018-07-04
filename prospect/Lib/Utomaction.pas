unit Utomaction;

interface
uses extctrls,Classes,Controls,forms,sysutils,stdctrls,
     HCtrls,HEnt1,HMsgBox,UTOM, UTob,M3FP,TiersUtil,Paramsoc,menus, 
     // CRM_20080901_MNG_FQ;012;10843
{$IFNDEF EAGLSERVER}
     MailOl,
     SaisieList,uTableFiltre,AglIsoflex, AglInit,LookUp,UtilConfid,
{$ENDIF EAGLSERVER}

{$IFDEF EAGLCLIENT}
     MaineAGL,eFiche,UtileAGL,Spin,
{$ELSE}
     db,hdb,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}DBCtrls,
{$IFNDEF EAGLSERVER}
     Fe_Main,HsysMenu,Fiche,
{$ENDIF EAGLSERVER}
{$ENDIF}
{$IFDEF POURVOIR}
OutlookView_TLB,
{$ENDIF}
Graphics,HPanel,grids,windows,EntRT,UtilRT,HRichOle,ComCtrls,UtilPGI
{$IFDEF AFFAIRE}
        ,AffaireUtil,FactUtil
{$ifdef GIGI}
  ,dicobtp
{$endif}
{$ENDIF}
 , HTB97, EntGC, UtilAction
{$IFDEF SAV}
 ,wParcNome,wParc
{$ENDIF}
{$IFNDEF EAGLSERVER} ,UtilChainage{$ENDIF EAGLSERVER},YPlanning,YRessource,UtofRessource_Mul
 ;
// Modif 0602
// La chaîne DESTMAIL contenant les no de contacts commence obligatoirement par ; (ex: ;2;12;5;)
// pr pourvoir faire une recherche d'un no de contact parmi la liste
// la chaîne DESTMAIL est à blanc si aucun contact
// ==> modif des fonctions PutInterlocuteurs,ChargeInterlocuteurs et ActEnvoiMessage
Const
    PREVUE   : string = 'PRE';
    ANNULEE  : String = 'ANU';
    REALISEE : String = 'REA';
    NONREALISABLE : String = 'NRE';
    LGINTERVINT : Integer = 85;
    LGDESTMAIL  : Integer = 80;
    RESPONSABLE : string = 'RSP';
    INTERLOCUTEUR : string = 'INT';
    RCHTableName = 'ACTIONSCHAINEES';
    StVisee : String = 'Chaînage terminé, pièce(s) visée(s) : ';

Type
    TOM_ACTIONS = Class (TOM)
{$IFDEF POURVOIR}
    OVCtl1: TOVCtl;
{$ENDIF}
    Private
     AsInsert : boolean;
{$IFNDEF EAGLSERVER}
     TF : TTableFiltre;
{$ENDIF EAGLSERVER}
     TypeAction : TActionFiche;
{$IFDEF QUALITE}
     iPerspective,iNumeroContact,iIdentParc, iQncNum, iPlanCorrNum, iQLitigieuse : Integer;
{$ELSE QUALITE}
     iPerspective,iNumeroContact,iIdentParc : Integer;
{$ENDIF QUALITE}
     NoChangeProspect,NoChangeProjet,FicheChainage,NoChangeChainage,NoChangeContact,NoChangeAffaire,ActionOblig,SaisieAction,OrigineTiers,VerrouModif,NouvelleAction : Boolean;
     NoCreat : Boolean;
     HeureDeb,TheTime{$IFNDEF EAGLSERVER},Old_DateAction{$ENDIF EAGLSERVER}: TDateTime;
     Timer1: TTimer;
     LesInterlocuteurs,LesIntervenants,TobAct : tob;
{$IFNDEF EAGLSERVER}
     G1,G2 : THGRID;
     G1ColNom,G2ColNom: integer;
     Split : TSplitter;
     G1LesColonnes,G2LesColonnes: string ;
     DerniereCreateTiers : string;
     DerniereCreateNum   : integer;
{$ENDIF EAGLSERVER}
     sttiersduplic       : string;
     iNumActDuplic       : Integer;
     OldEtatAction,Old_Typeaction,Old_Intervenant,TypeActionInitial,EtatActionInitial      : string;
     Duplic,Planning              : Boolean;
     OutLook,EnvoiMail,MailResp,MailAuto,ModifResp,NonSupp : string;
     MailEtatAnu, MailEtatPre, MailEtatRea : string;
     MailRespAnu, MailRespPre, MailRespRea : string;
     MailAnuAuto, MailPreAuto, MailReaAuto : string;

     StDelaiRappel : String;
     TobTypActEncours : Tob; // Tob du type d'action en cours
    { On met les paramsoc en variable }
    soRtactgestech: boolean;
    soRtactgestcout: boolean;
    soRtactgestchrono: boolean;
    soRtactrappel: boolean;
    soRtprojmultitiers: boolean;
    soRtgestinfos001: boolean;
    soRtmesstypeact: boolean;
    soRtmesslibact: boolean;
    soRtmessbl: boolean;
    //
//    procedure MonOnTimer(Sender: TObject);
//    procedure MajIntervint;
     {$IFNDEF EAGLSERVER}stNatureauxi,stParticulier,Old_HeureAction,stLibelle : string;{$ENDIF EAGLSERVER}
		 sTiersCli: boolean;
     {$IFNDEF EAGLSERVER}
     ClotureEncours: Boolean;
     {$ENDIF EAGLSERVER}
     ModifLot : boolean;
{$IFNDEF EAGLSERVER}
     StSQL : string;
{$ENDIF EAGLSERVER}
     ListeInitRessources,ListeFinRessources,stIntervint : string;
     DuplicAct : Boolean;
{$IFDEF QUALITE}
     ClotureFromFiche : boolean;
     sAction, sQNCOrigine, FamAction : string;
{$ENDIF QUALITE}
     sOrigine : string;
    Function ActEnvoiMessage (TypeMessage,MailAuto,EtatAction :String; MajAct : boolean) : String;
{$IFNDEF EAGLSERVER}
    Function ActFormatHeure : String;
    Function ActFormatDuree : Integer;
    function  ChargeInterlocuteurs : Boolean;
    procedure PutInterlocuteurs;
    function  ChargeIntervenants (Auxiliaire : string; NoAction : integer) : Boolean;
    procedure NomContact;
{$ENDIF EAGLSERVER}
    procedure LookEcran (Affect : Boolean);

{$IFNDEF EAGLSERVER}
    procedure ChgtTypeAction;
    // gestion des grids
    procedure AddRemoveItemFromPopup(stPopup,stItem : string; bVisible : boolean);
    Procedure AddSplitter ;
    Procedure ClickInsert(GS : THGrid; Row:integer) ;
    Procedure ClickDel(GS : THGrid; Row:integer) ;
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G1DessineCell (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
    procedure G1CellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
    procedure G1CellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
    procedure G1ElipsisClick(Sender: TObject);
    procedure G1RemplirLaLigne(NewInterlocuteur : string; Row : integer);
    procedure G2DessineCell (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
    procedure G2CellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
    procedure G2CellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
    procedure G2ElipsisClick(Sender: TObject);
    procedure G2RemplirLaLigne(NewIntervenant : string; Row : integer);

    procedure RTAppelParamCL;
    procedure GereIntervenants;
    procedure RattacherChainage;
    procedure ChargeChainage;
    procedure DetacherChainage ;
    procedure PutIntervenants;
    Procedure InitGrids;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure RappelClick(Sender: TObject);
    procedure ActMajIntervint (NewIntervenant:String);
    procedure GereIsoflex;
    procedure OnExit_DateAction(Sender: TObject);
    procedure MnPiece_OnClick(Sender: TObject);
    procedure OnExit_HeureAction(Sender: TObject);
    procedure OnChange_HeureAction(Sender: TObject);
{$ENDIF EAGLSERVER}

{$IFDEF AFFAIRE}
    procedure RechercheAffaire (Sender: TObject);
    procedure EffaceAffaire (Sender: TObject);
    Procedure AffaireEnabled( grise : boolean);
{$ENDIF AFFAIRE}
    {$IFDEF SAV}
    // Gestion Parc
    procedure MAJMenuParc;
    procedure AffecterParc(Sender: TObject);
    procedure SupprimerPArc(Sender: TObject) ;
    procedure VoirFicheParc(Sender: TObject) ;
    {$ENDIF SAV}
    {$IFDEF QUALITE}
    { Module Qualité}
		procedure RAC_INTERVENANT_OnElipsisClick(Sender:TObject);
		procedure RAC_TIERS_OnElipsisClick(Sender:TObject);
		procedure RAC_LigneOrdre_OnElipsisClick(Sender:TObject);
		procedure MnProspect_OnClick(Sender:TObject);
		procedure MnFicheInfos_OnClick(Sender:TObject);
		procedure MnContacts_OnClick(Sender:TObject);
		procedure MnClotureRac_OnClick(Sender:TObject);
		procedure MnLpNonConf_OnClick(Sender:TObject);
		procedure MnDemDerog_OnClick(Sender:TObject);
		procedure MnLpPlanCorr_OnClick(Sender:TObject);
		procedure QDEMDEROGNUM_OnElipsisClick(Sender:TObject);
		procedure QNCNUM_OnElipsisClick(Sender:TObject);
		procedure QPLANCORRNUM_OnElipsisClick(Sender:TObject);
    procedure PmAction_OnPopUp(Sender: TObject);
    procedure MnJournalAction_OnClick(Sender: TObject);
    procedure MnProperties_OnClick(Sender: TObject);
    procedure QUALITETYPE_OnChange(Sender: TObject);
    {$ENDIF QUALITE}
    function ModifChampYplanning : Boolean;
    procedure MAJYPlanning (TobYPlanning : Tob;Ressource,Guid:string);
    function ControleEstLibre : Boolean;
    procedure RACSetFocusControl(Champ : String);
{$IFDEF BTP}
    procedure RAC_INTERVENANTCLICK(Sender: TObject);
      //fv1 - 18/03/2016 - Nouvelle Version Planning
    procedure ControleChamp(Champ, Valeur: String);
    procedure ControleCritere(Critere: String);
{$ENDIF}
    Public
     stTiers,stProjet,stOperation,stIntervenant,stAffaire : String;
     iNumChainage : Integer;
     stProduitpgi : String;
     soRtactresp: string;
    procedure OnClose ; override ;
    procedure OnNewRecord ; override ;
    procedure OnUpdateRecord ; override ;
    procedure OnLoadRecord ; override ;
{$IFNDEF EAGLSERVER}
    procedure SetArguments(StSQL : string);
    procedure OnLoadAlerte  ; override ;
    procedure OnChangeField (F : TField) ; override ;
    procedure RTAjouteOutlook( Ajout : String );
    procedure DuplicationAct (TiersDuplic: string;NumActDuplic:integer);
    function GetInfoGed : string ;
    procedure RTAppelGED_OnClick (Sender: TObject) ;
    procedure GestionBoutonGED;
	function GenerChainageIntervention : Integer;
{$ENDIF EAGLSERVER}
    procedure OnArgument (Arguments : String )  ; override ;
    Procedure OnDeleteRecord ; override ;
    procedure HeureDebut;
    procedure HeureFin;
    procedure OnAfterUpdateRecord        ; override ;
    {$IFDEF QUALITE}
    procedure SetIdentifiant;
    procedure SetNumRQNEnabled;
    {$ENDIF QUALITE}
		procedure SetControlsEnabled(Const FieldsName: array of String; Const Enability: Boolean);
    procedure SetGBClotureEnabled(Enable: boolean);
    {$IFNDEF EAGLSERVER}
    function TestAlerteAction (CodeA : String) : boolean;
    procedure ListAlerte_OnClick_RAC(Sender: TObject);
    procedure Alerte_OnClick_RAC(Sender: TObject);
    {$ENDIF EAGLSERVER}
END ;

{$IFNDEF EAGLSERVER}
procedure AGLHeureDebut(parms:array of variant; nb: integer ) ;
procedure AGLHeureFin(parms:array of variant; nb: integer ) ;
procedure AGLRTAjouteOutlook(parms:array of variant; nb: integer ) ;
Function AGLActFormatHeure(parms:array of variant; nb: integer ) : variant ;
Function AGLActFormatDuree(parms:array of variant; nb: integer ) : variant ;
procedure AGLActEnvoiMessage(parms:array of variant; nb: integer ) ;
procedure DessineCell(GS:THGrid; ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
procedure GereNewLigne(GS:THGrid)  ;
procedure NewLigne(GS:THGrid) ;
Function  EstRempli(GS:THGrid; Lig : integer) : Boolean ;
Procedure InitRow (GS:THGrid; R : integer) ;
function Action_MyAfterImport (Sender : TObject): string ;
procedure Action_GestionBoutonGed (Sender : TObject) ;
{$ENDIF EAGLSERVER}

const
	// libellés des messages
	TexteMessage: array[1..42] of string 	= (
          {1}        'Le code tiers doit être renseigné'
          {2}        ,'Le code tiers n''existe pas.'
          {3}        ,'La date d''échéance doit être postérieure à la date d''action'
          {4}        ,'L''opération n''existe pas'
          {5}        ,'La date de l''action doit être renseignée'
          {6}        ,'Le type d''action doit être renseigné'
          {7}        ,'L''état doit être renseigné'
          {8}        ,'Le projet n''existe pas'
          {9}        ,'La proposition n''existe pas ou ne concerne pas ce tiers'
          {10}       ,'Cet intervenant est déjà sélectionné'
          {11}       ,'La saisie d''une action est obligatoire'
          {12}       ,'Vous ne pouvez plus saisir d''intervenants car la liste est complète'
          {13}       ,'Opération impossible.Ce tiers est fermé'
          {14}       ,'Vous n''êtes pas autorisé à saisir cette action'
          {15}       ,'Le responsable doit être renseigné'
          {16}       ,'Le responsable n''existe pas'
          {17}       ,'L''heure de l''action doit être renseignée dans ce cas'
          {18}       ,'Le tiers du chaînage est différent de celui de l''action'
          {19}       ,'L''état de l''action ne peut pas être réalisé'
          {20}       ,'Cette action ne peut pas être supprimée'
          {21}       ,'Vous n''êtes pas autorisé à modifier cette action'
          {22}       ,'La saisie du champ suivant est obligatoire : '
          {23}       ,'Le libellé de l''action doit être renseigné'
          {24}       ,'Code affaire incorrect'
          {25}       ,'la fiche parc n''existe pas'
          {26}       ,'Vous ne pouvez pas clôturer cette action, elle doit être réalisée'
{ QUALITE}
          {27}       ,'Le n° de demande de dérogation est incorrect.'
          {28}       ,'Le type d''action est obligatoire.'
          {29}       ,'Cette ressource n''existe pas.'
          {30}       ,'Le n° de non conformité est incorrect.'
          {31}       ,'Le n° de plan correcteur est incorrect.'
          {32}       ,''
          {33}       ,'le n° d''ordre saisi est inconnu dans la nature de travail.'
          {34}       ,'Cette action est de type "Production", Vous devez saisir un ordre de production.'
          {35}       ,'Vous devez renseigner la ressource de la clôture.'
          {36}       ,'Vous devez renseigner la date de clôture.'
{ QUALITE}
          {37}       ,'Des intervenants ne sont pas disponibles sur cette période'
          {38}       ,'Vous ne pouvez pas rattacher un plan correcteur à une action immédiate.'
          {39}       ,'Le projet ne concerne pas ce tiers'
          {40}       ,'Le projet n''est pas affecté à ce tiers. Vous devez saisir le tiers dans la fiche projet'
          {41}       ,'Vous devez rattacher cette action à une non conformité ou à un plan correcteur.'
          {42}       ,'Cette action est rattachée à un plan correcteur. Elle ne peut être de type "Demande de dérogation".'
          );
implementation

Uses
	wCommuns
  {$IFDEF VER150}
    ,Variants
  {$ENDIF VER150}
{$IFDEF QUALITE}
	,AglInitGC
  ,rqNonConformite
  ,rqDemDerog_Tof
  ,RqNonConf_Tof
  ,rqPlanCorr_Tof
  ,wAction
  ,wJournalAction
  ,wJetons
  ,wRessource
  ,wMnu
  ,wOrdreLigMul_Tof
  ,EntGp
  ,wOrdreLig
  ,wRapport
{$ENDIF QUALITE}
{$IFNDEF EAGLSERVER}
  ,UtilAlertes,YAlertesConst,EntPgi
{$ENDIF EAGLSERVER}
  ;

{ TOM_ACTIONS }
{$IFNDEF EAGLSERVER}
procedure AGLHeureDebut( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else
   if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
if (OM is TOM_ACTIONS) then TOM_ACTIONS(OM).HeureDebut else exit;
end;

procedure AGLHeureFin( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else
   if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
if (OM is TOM_ACTIONS) then TOM_ACTIONS(OM).HeureFin else exit;
end;

procedure AGLRTAjouteOutlook( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else
   if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
if (OM is TOM_ACTIONS) then TOM_ACTIONS(OM).RTAjouteOutlook(Parms[1]) else exit;
end;

procedure AGLActEnvoiMessage( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else
   if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
if (OM is TOM_ACTIONS) then TOM_ACTIONS(OM).ActEnvoiMessage(Parms[1],Parms[2],TOM_ACTIONS(OM).GetField('RAC_ETATACTION'),true) else exit;
end;

Function AGLActFormatheure( parms: array of variant; nb: integer ) : variant;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else
   if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
if (OM is TOM_ACTIONS) then result:=TOM_ACTIONS(OM).ActFormatHeure else exit;
end;

Function AGLActFormatDuree( parms: array of variant; nb: integer ) : variant;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else
   if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
if (OM is TOM_ACTIONS) then result:=TOM_ACTIONS(OM).ActFormatDuree else exit;
end;

procedure AGLNomContact( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFFiche) then OM:=TFFiche(F).OM else
     if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
  if (OM is TOM_ACTIONS) then TOM_ACTIONS(OM).NomContact else exit;

  if (F is TFSaisieList) then
     if TFSaisieList(F).LeFiltre.State <> dsInsert then
        TFSaisieList(F).LeFiltre.Edit;

  if (F is TFSaisieList) then
    if TFSaisieList(F).LeFiltre.DataSet.State = dsEdit then //pour forcer la fiche à passer en mode Edition
      begin
      {$IFDEF EAGLCLIENT}
      TFSaisieList(F).LeFiltre.DataSet.State := dsBrowse;
      {$ELSE}
      TFSaisieList(F).LeFiltre.DataSet.Post;
      {$ENDIF}
      end;
  { mng 07/09/04 : en cwas, empeche la création de l'action, déjà fait plus haut
  TFSaisieList(F).LeFiltre.Edit;}

end;

procedure AGLDuplicationAct( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else
   if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
if (OM is TOM_ACTIONS) then TOM_ACTIONS(OM).DuplicationAct('',0) else exit;
end;


procedure AGLChgtTypeAction( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else
   if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
if (OM is TOM_ACTIONS) then TOM_ACTIONS(OM).ChgtTypeAction else exit;
end;


procedure AGLRTAppelParamCL( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else
   if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
if (OM is TOM_ACTIONS) then TOM_ACTIONS(OM).RTAppelParamCL else exit;
end;

procedure AGLActMajIntervint( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else
   if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
if (OM is TOM_ACTIONS) then TOM_ACTIONS(OM).ActMajIntervint(Parms[1]) else exit;
end;

procedure AGLRattacherChainage( parms: array of variant; nb: integer ) ;//chainage
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else
   if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
if (OM is TOM_ACTIONS) then TOM_ACTIONS(OM).RattacherChainage else exit;
end;

procedure AGLDetacherChainage( parms: array of variant; nb: integer ) ;//chainage
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else
   if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
if (OM is TOM_ACTIONS) then TOM_ACTIONS(OM).DetacherChainage else exit;
end;

procedure TOM_ACTIONS.OnChangeField(F: TField);
var
  Select: String;
  Q : TQuery;
begin
	inherited;
  if (F.FieldName = 'RAC_TIERS') then
  begin
  	if (GetField('RAC_TIERS')='') then
    else
    begin
      Q := OpenSQL('SELECT T_NATUREAUXI,T_PARTICULIER,T_LIBELLE,T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+GetField('RAC_TIERS')+'"', True);
      try
        if Q.Eof then
        begin
          RACSetFocusControl('RAC_TIERS');
          LastError := 2;
{$ifdef GIGI}
          LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
          LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
          exit;
        end
        else
        begin
          stNatureauxi:=Q.FindField('T_NATUREAUXI').asstring;
          stLibelle:=Q.FindField('T_LIBELLE').asstring;
          stParticulier:=Q.FindField('T_PARTICULIER').asstring;
          if DS.State in [dsInsert,dsEdit] then
            SetField('RAC_AUXILIAIRE',Q.FindField('T_AUXILIAIRE').asstring);
          {$IFDEF AFFAIRE}
            if ( (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) ) and
               ( stProduitpgi<>'GRF') then
            begin
              if ( pos (stNatureauxi, FabriqueWhereNatureAuxiAff('AFF')) = 0 ) and
                 ( ( pos (stNatureauxi, FabriqueWhereNatureAuxiAff('PRO')) = 0 ) or
                 (VH_GC.GASeria=false) ) or (NoChangeAffaire) then
                 AffaireEnabled(false)
              else
                 AffaireEnabled(true);
            end;
          {$ENDIF}
        end;
      finally
        Ferme(Q);
      end;  
      if (DS.State = dsInsert) and (Duplic = False) then
      begin
        //SetField('RAC_AUXILIAIRE',TiersAuxiliaire (GetField('RAC_TIERS'), False));
        if (soRtactresp='COM' ) and ( Getfield('RAC_INTERVENANT') = '' ) then
        begin
          SetField('RAC_INTERVENANT', RTRechResponsable (GetField('RAC_TIERS')));
          {$IFDEF EAGLCLIENT}
            if (GetField ('RAC_INTERVENANT') <> '') and (StTiers = '') then
            begin
              Select := 'SELECT ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE = "'+ GetField ('RAC_INTERVENANT')+'"';
              Q:=OpenSQL(Select, True);
              try
                if not Q.Eof then
                  SetControlText ('TRAC_INTERVENANT_',Q.Fields[0].AsString);
              finally
                Ferme(Q);
              end;
            end;
          {$ENDIF}
          if (GetField ('RAC_INTERVENANT') <> '') then ActMajIntervint (GetField('RAC_INTERVENANT'));
        end;
      end;
      //mcd 30/11/2005 il faut prendre en compte le paramétrage possible intervenent que ressource
      if (  ( soRtactresp = 'RE1' )
         or ( soRtactresp = 'RE2' )
         or ( soRtactresp = 'RE3' ))
        and (CtxAffaire in V_PGI.PGIContexte)
        and (GetField('RAC_TIERS') <>'')
        and (GetControlText('RAC_INTERVENANT')='') then
      begin
        Select :='SELECT YTC_RESSOURCE1,YTC_RESSOURCE2,YTC_RESSOURCE3 FROM  TIERSCOMPL WHERE YTC_TIERS="'+GetField('RAC_TIERS')+'"';
        Q := OpenSQL(Select, True);
        try
          if not Q.Eof then
          begin
            if ( soRtactresp = 'RE1' ) then SetField('RAC_INTERVENANT',Q.FindField('YTC_RESSOURCE1').asstring);
            if ( soRtactresp = 'RE2' ) then SetField('RAC_INTERVENANT',Q.FindField('YTC_RESSOURCE2').asstring);
            if ( soRtactresp = 'RE3' ) then SetField('RAC_INTERVENANT',Q.FindField('YTC_RESSOURCE3').asstring);
          end;
        finally
          Ferme(Q);
        end;
        if (GetField ('RAC_INTERVENANT') <> '') then ActMajIntervint (GetField('RAC_INTERVENANT'));
      end;
                //fin mcd 30/11/2005
    end;

    sTiersCli:= (stProduitPgi='QNC') and (GetField('RAC_TIERS') <> '') and (GetField('RAC_TYPEACTION')<>'') and ExistTiers('CLI', GetField('RAC_TIERS'));
  end
  else if (F.FieldName = 'RAC_INTERVENANT') then
  begin
  	if (DS.State <> dsBrowse) and (Duplic = False) then
  	begin
    	if GetField('RAC_INTERVENANT') <> Old_Intervenant then
    	begin
      	if ((GetField ('RAC_INTERVENANT') <> '') and  (Not ExisteSQL ('SELECT ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE="'+GetField('RAC_INTERVENANT')+'"'))) then
      	begin
        	RACSetFocusControl('RAC_INTERVENANT');
	        LastError := 16;
{$ifdef GIGI}
          LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
          LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
    	    exit;
      	end;

	      if stProduitpgi<>'QNC' then
					GereIntervenants;
    	end;
    end;
  end
  else if F.FieldName = 'RAC_TYPEACTION' then
  begin
  	if Ecran is TFSaisieList  then
	  	ChgtTypeAction();
{$IFDEF QUALITE}
    if (stProduitPGI='QNC') then
    begin
			FamAction:= GetFamAction(GetField('RAC_TYPEACTION'));
      SetControlText('NATUREACTION', FamAction);
      { Impossible de modifier le tiers de l'action : il fait partie de la clé de la table RTACTIONS }
      SetControlEnabled('RAC_TIERS', TypeAction = TaCreat);
      SetControlEnabled('RAC_TYPELOCA', (TypeAction<>TaConsult) and (GetControlEnabled('RAC_TYPELOCA') and (FamAction<>'DER')));
      SetControlProperty('RAC_NATURETRAVAIL', 'ENABLED', (TypeAction<>TaConsult) and (FamAction = 'WPR'));
      SetControlProperty('RAC_LIGNEORDRE'   , 'ENABLED', (TypeAction<>TaConsult) and (FamAction = 'WPR'));
      if (FamAction='DER') then
	    begin
        SetField('RAC_TYPELOCA', '002');
 	  		if (TypeAction=taCreat) and (GetControlText('RAC_TIERS')='') then
        	SetField('RAC_TIERS', stTiers);
      end;
      if (FamAction<>'WPR') and (GetField('RAC_NATURETRAVAIL')<>'') then
        SetControlText('RAC_NATURETRAVAIL', '');
      if (FamAction<>'WPR') and (GetField('RAC_LIGNEORDRE')<>0) then
        SetField('RAC_LIGNEORDRE', 0);
    end;
{$ENDIF QUALITE}
  end
  else if (GetField('RAC_ETATACTION')='ZCL') and (OldEtatAction<>'ZCL') and (OldEtatAction <> REALISEE) then
  begin
    SetField('RAC_ETATACTION', OldEtatAction);
    SetControlText('RAC_ETATACTION', OldEtatAction);
    RACSetFocusControl('RAC_ETATACTION');
    LastError := 26;
{$ifdef GIGI}
    LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
  end;
  ;
{$IFDEF QUALITE}
	if (stProduitPGI='QNC') then
 	begin
 		if F.FieldName = 'RAC_QNCNUM' then
    	SetControlProperty('MnLpNonConf' , 'VISIBLE', GetField('RAC_QNCNUM')<>0)
  	else if F.FieldName = 'RAC_QDEMDEROGNUM' then
    	SetControlProperty('MnDemDerog'  , 'VISIBLE', GetField('RAC_QDEMDEROGNUM')<>0)
  	else if (F.FieldName = 'RAC_QPLANCORRNUM') or (F.FieldName='RAC_QUALITETYPE') then
    begin
      if (GetControlText('RAC_QUALITETYPE')='001') and (GetField('RAC_QPLANCORRNUM')<>0) then
      begin
        SetField('RAC_QPLANCORRNUM', 0);
        LastError:= 38;
        LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
      end;
    	SetControlProperty('MnLpPlanCorr', 'VISIBLE', GetField('RAC_QPLANCORRNUM')<>0);
    end
    ;
    if  (F.FieldName = 'RAC_QNCNUM') or (F.FieldName = 'RAC_QPLANCORRNUM') or (F.FieldName = 'RAC_QUALITETYPE')
     or (F.FieldName = 'RAC_TYPEACTION') then
      SetNumRQNEnabled;
  end;
{$ENDIF QUALITE}
end;

procedure TOM_ACTIONS.RappelClick(Sender: TObject);
begin
  if (DS.State = dsEdit) then
  begin
    if (GetCheckBoxState('RAC_GESTRAPPEL') = cbChecked) then
      begin
      if StDelaiRappel = '' then
         begin
         if TobTypActEncours <> Nil then
             StDelaiRappel:=TobTypActEncours.GetValue('RPA_DELAIRAPPEL');
         end;
      SetField('RAC_DELAIRAPPEL',StDelaiRappel);
      end
    else
       SetField('RAC_DELAIRAPPEL','');
  end;
end;

procedure TOM_ACTIONS.OnLoadAlerte;
begin
  if (not V_Pgi.SilentMode) and (AlerteActive(TableToPrefixe(TableName))) and (not AfterInserting )  and ( not Modiflot )  then
     TestAlerteAction(CodeOuverture+';'+CodeDateAnniv);

end;
{$ENDIF !EAGLSERVER}
procedure TOM_ACTIONS.OnLoadRecord;
var i:integer;
    Heure: TDateTime;
    Hour, Min, Sec, MSec: Word;
{$IFNDEF EAGLSERVER}
    SavEcran : String;
{$ENDIF EAGLSERVER}
    {$IFDEF AFFAIRE}
    Q : TQuery;
    {$ENDIF AFFAIRE}
{$IFNDEF EAGLSERVER}
    ActionNonCloturee: Boolean;
    TI : Tob;
{$ENDIF EAGLSERVER}
begin
inherited;
  { pour debuguer }
  i:=0;
  if i=1 then exit;
  NouvelleAction := ds.State in [dsinsert];
  ListeInitRessources:='';

{$IFNDEF EAGLSERVER}
  if (Ecran is TFSaisieList) then
  begin
    if (copy(ecran.name,1,17) <> 'WINTERVENTION_FSL') then
      begin
      stTiers:=TF.TOBFiltre.GetValue('RCH_TIERS');
      iNumChainage:=TF.TOBFiltre.GetValue('RCH_NUMERO');
      end
    else
      begin
      stTiers:=TF.TOBFiltre.GetValue('WIV_TIERS');
      iNumChainage:=TF.TOBFiltre.GetValue('WIV_NUMCHAINAGE');
      end;
    SetControlEnabled('RAC_TIERS',False) ;
    SetField('RAC_TIERS',stTiers);
    NonSupp:='-';
  end;

  if (TFFiche(Ecran) <> Nil) then    // Si on vient de la fiche RTCOURRIER il n'y a pas de tform
    TPopupMenu(GetControl('POPCOMPLEMENT')).items[3].visible:=(Getfield('RAC_OPERATION')<>'');
  SetControlEnabled('BDUPLICATION',True);
  SetControlEnabled('BDelete',True);
  SetControlEnabled('G1',True);
  SetControlEnabled('G2',True);
  SetControlEnabled('RAC_INTERVENANT',True);
  if (Ecran is TFFiche) then
  begin
  	if stProduitPGI <> 'QNC' then
    begin
      TPopupMenu(GetControl('POPCOMPLEMENT')).items[7].enabled:=True;
      TPopupMenu(GetControl('POPCOMPLEMENT')).items[5].enabled:=True;
    end;
  end;
  for i:=1 to 3 do
    SetControlCaption('TRAC_TABLELIBRE'+IntToStr(i),'&'+RechDom('RTLIBCHAMPSLIBRES','AL'+IntToStr(i),FALSE)) ;

  SetControlEnabled('BDUPLICATION',(TypeAction <> taconsult));


  if stProduitPGI <> 'QNC' then
  begin
		{Etat "Clôturée" ds la liste : invisible en création}
	  SetControlProperty('RAC_ETATACTION', 'PLUS', iif(TypeAction=taCreat,' AND CO_CODE<>"ZCL"', ''));


    if (DS<>nil) and ( not (ds.state in [dsinsert]) and not VerrouModif ) then
    begin
        if OrigineTiers then
          begin
          if GetControlText ('NATUREAUXI') = 'FOU' then
             ModifAutorisee := RTDroitModifActionsF('',Getfield('RAC_TYPEACTION'),Getfield('RAC_INTERVENANT'))
          else
             ModifAutorisee := RTDroitModifActions('',Getfield('RAC_TYPEACTION'),Getfield('RAC_INTERVENANT'));
          end
        else
          { mng 30/11/2007 : cas de l'accès à partir du mul des actions, être en consult seule si
            tiers en consult seule et pas d'extension de création à tous les tiers }
          ModifAutorisee := (RTDroitModiftiers(GetField('RAC_TIERS'))=true)  or ( RTXXWhereConfident('CREATACT') <> '-' )
        ;
        if Ecran is TFSaisieList then
           SetControlEnabled('bPost',ModifAutorisee)
        else SetControlEnabled('BVALIDER',ModifAutorisee);
        SetControlEnabled('BDUPLICATION',ModifAutorisee);
        SetControlEnabled('BDelete',ModifAutorisee);
        SetControlEnabled('G1',ModifAutorisee);
        SetControlEnabled('G2',ModifAutorisee);
        if Ecran is TFFiche then
          begin
          TPopupMenu(GetControl('POPCOMPLEMENT')).items[7].visible:=false;
          TPopupMenu(GetControl('POPCOMPLEMENT')).items[5].visible:=false;
          end;
    end;
  end;

  if (TFFiche(Ecran) <> Nil) then    // Si on vient de la fiche RTCOURRIER il n'y a pas de tform
  begin
    G2.videpile (False);
    G2.RowCount:=2;
  end;
  if ((stTiersDuplic <> '') and (iNumActDuplic <> 0)) or ((stProduitPGI='QNC') and (iNumActDuplic<>0)) then
  begin
   	DuplicationAct (stTiersDuplic,iNumActDuplic);
   	stTiersDuplic := '';
   	iNumActDuplic := 0;
    DuplicAct := True;
  end;

  if NoChangeProjet then
   SetControlEnabled('RAC_PROJET',False) ;
  if NoChangeContact then
   SetControlEnabled('LECONTACT',False) ;
{$ENDIF EAGLSERVER}
  // Fiche ACTION QUALITE
{$IFDEF QUALITE}
	if stProduitPGI = 'QNC' then
  begin
		{ Etat de l'action qualité }
	  SetControlProperty('RAC_ETATACTION', 'PLUS', iif(GetField('RAC_ETATACTION')='ZCL','',' AND CO_CODE<>"ZCL"'));

		{ Type action }
    SetControlEnabled('RAC_TYPEACTION', (GetField('RAC_TYPEACTION')<>'DER') or (TypeAction = TaCreat));

	  { Fiche clôturée non modifiable }
		if (GetField('RAC_ETATACTION')='ZCL') or ClotureEncours then
	  begin
      RACSetFocusControl('RAC_REALISEELE');
			// 		FicheReadOnly(Ecran, true );
      if ClotureEncours then
			begin
      	PgiInfo(TraduireMemoire('Clôture de la fiche action. Veuillez renseigner les informations de clôture.'));
	      if DS.State<>DsEdit then
        	DS.Edit;
	      SetField('RAC_CLOTUREELE'		, V_PGI.DateEntree);
  	    SetField('RAC_CLOTUREEPAR'	, GetRessourceAssocieeUser(V_PGI.User) );
    	  SetField('RAC_REALISEEPAR'	, GetField('RAC_AREALISERPAR') );
      end;
    end;

    { Menus loupes}
    SetControlProperty('MnLpNonConf' , 'VISIBLE', GetField('RAC_QNCNUM')<>0);
    SetControlProperty('MnLpPlanCorr', 'VISIBLE', GetField('RAC_QPLANCORRNUM')<>0);
   	SetControlVisible('mnDemDerog'			, GetField('RAC_QDEMDEROGNUM')<>0);
   	SetControlVisible('mnProspect'			, GetField('RAC_TIERS')<>'');
   	SetControlVisible('mnContacts'			, GetField('RAC_TIERS')<>'');

   	SetControlProperty('RAC_QDEMDEROGNUM' , 'VISIBLE', ((sQNCOrigine='RQN') or (sQNCOrigine='')) );
   	SetControlProperty('TRAC_QDEMDEROGNUM', 'VISIBLE', ((sQNCOrigine='RQN') or (sQNCOrigine='')) );

		RACSetFocusControl('RAC_TYPEACTION');

    if not (CtxGpao in V_Pgi.PgiContexte) then
    	SetControlProperty('RAC_TYPEACTION', 'PLUS', ' AND RPA_QNATUREACTION <> "WPR"');

    if (GetField('RAC_TIERS') <> '') and ExistTiers('CLI', GetField('RAC_TIERS')) then
    	sTiersCli:=True;
//    else
//			{Si pas tier client, pas accès aux actions de type "Demandes de dérogations"}
//    	SetControlProperty('RAC_TYPEACTION', 'PLUS', THDbValComboBox(GetControl('RAC_TYPEACTION')).Plus+' AND RPA_PRODUITPGI="'+stProduitpgi+'" AND RPA_QNATUREACTION <> "DER"');
    if sQNCOrigine='RQP' then // Qualité : Plans correcteurs
    begin
      SetControlProperty('RAC_QUALITETYPE', 'PLUS', ' AND CO_CODE<>"001"');
     	SetControlProperty('RAC_TYPEACTION' , 'PLUS', THDbValComboBox(GetControl('RAC_TYPEACTION')).Plus+' AND RPA_PRODUITPGI="'+stProduitpgi+'" AND RPA_QNATUREACTION <> "DER"');
    end;

		FamAction:= GetFamAction(GetField('RAC_TYPEACTION'));

   	SetControlProperty('RAC_ETATACTION' 	, 'ENABLED', (TypeAction<>TaConsult) and (GetField('RAC_ETATACTION')<>'ZCL'));
   	SetControlProperty('RAC_QDEMDEROGNUM' , 'ENABLED', (TypeAction<>TaConsult) and (GetField('RAC_TYPEACTION')='DER'));
		SetControlproperty('RAC_TIERS'				, 'ENABLED', (TypeAction<>TaConsult) and (GetField('RAC_QDEMDEROGNUM')=0));
	 	SetControlProperty('RAC_NATURETRAVAIL', 'ENABLED', (TypeAction<>TaConsult) and (FamAction = 'WPR'));
	 	SetControlProperty('RAC_LIGNEORDRE'   , 'ENABLED', (TypeAction<>TaConsult) and (FamAction = 'WPR'));

    SetNumRQNEnabled;
  end
  else
  begin
{$ENDIF QUALITE}
	  Heure := GetField ('RAC_CHRONOMETRE');
  	DecodeTime(Heure, Hour, Min, Sec, MSec);
{$IFNDEF EAGLSERVER}
    if (TFFiche(Ecran) <> Nil) then    // Si on vient de la fiche RTCOURRIER il n'y a pas de tform
      SetControlText('CHRONO', IntToStr(Hour) + ':'
        + IntToStr(Min) + ':'+ IntToStr(Sec));

    if (NoChangeProspect) or ((DS<>nil) and (not(DS.State in [dsInsert]))) then
    begin
      SetControlEnabled('RAC_TIERS',False) ;
      //RACSetFocusControl('RAC_TYPEACTION');
    end else
    begin

      {$IFDEF EAGLCLIENT}
      if Ecran is TFFiche then
        TFFiche(Ecran).QFiche.CurrentFille.PutEcran(TFFiche(Ecran)) ;
      {$ENDIF}

      RACSetFocusControl('RAC_TIERS');
    end;
	  if Getfield ('RAC_AUXILIAIRE') <> '' then
  	begin
	//   SetControlProperty ('RAC_NUMEROCONTACT', 'Plus', Getfield ('RAC_AUXILIAIRE'));
	//   SetControlProperty ('RAC_NUMEROCONTACT', 'DataType', 'YYCONTACT');
	//   SetControlProperty ('RAC_DESTMAIL', 'Plus', Getfield ('RAC_AUXILIAIRE'));
	//   SetControlProperty ('RAC_DESTMAIL', 'DataType', 'YYCONTACT');
   		SetControlProperty ('RAC_PERSPECTIVE', 'Plus', Getfield ('RAC_AUXILIAIRE'));
   		SetControlProperty ('RAC_PERSPECTIVE', 'datatype', 'RTPERSPECTIVES');
   		if soRtprojmultitiers = False then SetControlProperty ('RAC_PROJET', 'Plus', Getfield ('RAC_AUXILIAIRE'));
  	end else
  	begin
   		if soRtprojmultitiers = False then SetControlProperty ('RAC_PROJET', 'Plus', '#######');
   		SetControlProperty ('RAC_PERSPECTIVE', 'datatype', '');
   		SetControlProperty ('RAC_PERSPECTIVE', 'Plus', '#######');
  	end;
    if (TFFiche(Ecran) <> Nil) then    // Si on vient de la fiche RTCOURRIER il n'y a pas de tform
    begin
//      G1.videpile (False); G2.videpile (False);
//      G1.RowCount:=2; G2.RowCount:=2;
      G1.videpile (False);
      G1.RowCount:=2;
  //{$IFDEF MNG}
      if LesInterlocuteurs <> Nil then LesInterlocuteurs.Free;
      LesInterlocuteurs:=tob.create('LesInterlocuteurs',Nil , -1);
//      if LesIntervenants <> Nil then LesIntervenants.Free;
//      LesIntervenants:=tob.create('LesIntervenants',Nil , -1);

      if ChargeInterlocuteurs then begin LesInterlocuteurs.PutGridDetail(G1,True,True,G1LesColonnes,True); NewLigne(G1); end;

      if DuplicAct = False then
      begin
        if LesIntervenants <> Nil then LesIntervenants.Free;
        LesIntervenants:=tob.create('LesIntervenants',Nil , -1);
        if (DS.State in [dsInsert]) and (stIntervint <> '') then
        begin
          ActMajIntervint (stIntervint);
        end
        else
        begin
          if ChargeIntervenants (Getfield ('RAC_AUXILIAIRE'),Getfield ('RAC_NUMACTION')) then
          begin
            LesIntervenants.PutGridDetail(G2,True,True,G2LesColonnes,True);
            NewLigne(G2);
            for i := 1 to G2.RowCount-1 do
              begin
              TI:=TOB(G2.Objects[0,i]);
              if TI<> nil then ListeInitRessources:=ListeInitRessources+TOB(G2.Objects[0,i]).GetValue('ARS_RESSOURCE')+';';
              end;
          end;
        end;
      end;  //  DuplicAct = False
  //{$ENDIF}
    end; // fin (TFFiche(Ecran) <> Nil)
		//cd140101   if iPerspective<>0 then Setcontroltext('NUM_PERSPECTIVE',intToStr(iPerspective));
{$IFDEF QUALITE}
  end;
{$ENDIF QUALITE}
  if Getfield ('RAC_TYPEACTION') <> '' then
  begin
  	Old_Typeaction := Getfield ('RAC_TYPEACTION');
   	LookEcran(False);
  	TypeActionInitial := Getfield ('RAC_TYPEACTION');
  end;
  if Getfield ('RAC_INTERVENANT') <> '' then
  	Old_Intervenant := Getfield ('RAC_INTERVENANT')
  else
  	Old_Intervenant := '';
  if Assigned(TFFiche(Ecran)) or Assigned(TFSaisieList(Ecran)) then    // Si on vient de la fiche RTCOURRIER il n'y a pas de tform
    SetControlText ('DUREEACT',PCSDureeCombo(Getfield('RAC_DUREEACTION')));
  NomContact();
  OldEtatAction:='';
  If (DS<>nil) and (Not(DS.State in [dsInsert])) Then
  begin
  	DerniereCreateTiers := '';
   	DerniereCreateNum := 0;
   	OldEtatAction:=GetField('RAC_ETATACTION');
  end;
  Old_DateAction := Getfield ('RAC_DATEACTION');
  EtatActionInitial:=GetField('RAC_ETATACTION');

  if ( getfield ('RAC_NUMCHAINAGE') = 0 ) and (iNumChainage <> 0) then
  begin
    ForceUpdate;
    SetField ('RAC_NUMCHAINAGE',iNumChainage) ;
  end;
	ChargeChainage; //Chainage

  if Ecran is TFSaisieList then
  begin
    SavEcran:=Ecran.Name;
    Ecran.Name:='RTACTIONS';
  end;
  AppliquerConfidentialite(Ecran,stProduitpgi);
  if Ecran is TFSaisieList then
    Ecran.Name:=SavEcran;
{$IFDEF AFFAIRE}
  if ( (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) ) and
    {$ifdef GIGI}
    (GetParamSocSecur ('SO_AFRTPROPOS',False)) and  //mcd 26/10/2005 on ne doit pas voir la mission si champ à faux
    {$endif}
    ( stProduitpgi <> 'GRF' ) then
  begin
    if Assigned(Ecran) then
      ChargeCleAffaire(THEDIT(GetControl('RAC_AFFAIRE0')),THEDIT(GetControl('RAC_AFFAIRE1')), THEDIT(GetControl('RAC_AFFAIRE2')),
                       THEDIT(GetControl('RAC_AFFAIRE3')), THEDIT(GetControl('RAC_AVENANT')),Nil , taModif,GetField ('RAC_AFFAIRE'),True);
    {ChargeAffaire(THEDIT(GetControl('RAC_AFFAIRE')), THEDIT(GetControl('RAC_AFFAIRE1')), THEDIT(GetControl('RAC_AFFAIRE2')),
      THEDIT(GetControl('RAC_AFFAIRE3')), THEDIT(GetControl('RAC_AVENANT')), THEDIT(GetControl('RAC_TIERS')),
      THEDIT(GetControl('RAC_AFFAIRE0')), false, TypeAction, nil, nil, True, False,
    False);}
    if Assigned(Ecran) then
      if (ds.state in [dsinsert]) and (GetField ('RAC_AFFAIRE') <> '') then
      begin
        SetField('RAC_AFFAIRE0',THEDIT(GetControl('RAC_AFFAIRE0')).Text);
        SetField('RAC_AFFAIRE1',THEDIT(GetControl('RAC_AFFAIRE1')).Text);
        SetField('RAC_AFFAIRE2',THEDIT(GetControl('RAC_AFFAIRE2')).Text);
        SetField('RAC_AFFAIRE3',THEDIT(GetControl('RAC_AFFAIRE3')).Text);
        //mcd 22/12/04 si avenant pas géré, ChargeCleaffaire Met blanc,
        //alors que le champs doit être à 00
        if THEDIT(GetControl('RAC_AVENANT')).Text ='' then THEDIT(GetControl('RAC_AVENANT')).Text:='00';
          SetField('RAC_AVENANT',THEDIT(GetControl('RAC_AVENANT')).Text);
      end;
    if (GetField ('RAC_AFFAIRE') <> '') then
    begin
      Q:= OpenSQL ('SELECT AFF_LIBELLE,AFF_NUMEROCONTACT,AFF_RESPONSABLE FROM AFFAIRE WHERE AFF_AFFAIRE="'+Getfield('RAC_AFFAIRE')+'"',True);
      if (Q<>Nil) and (not Q.Eof) then
      begin
        if Assigned(Ecran) then
          SetControlText ('TRAC_AFFAIRE',Q.FindField('AFF_LIBELLE').asstring);
        if (Q.FindField('AFF_NUMEROCONTACT').asinteger <> 0) and ( Getfield ('RAC_NUMEROCONTACT') = 0 ) then
        begin
          SetField ('RAC_NUMEROCONTACT', Q.FindField('AFF_NUMEROCONTACT').asinteger);
          NomContact();
        end;
        if (Q.FindField('AFF_RESPONSABLE').asString <> '') then
          SetField ('RAC_INTERVENANT', Q.FindField('AFF_RESPONSABLE').asString);
      end;
      Ferme (Q);
    end
    else
    begin
      if Assigned(Ecran) then
      begin
        THEDIT(GetControl('RAC_AVENANT')).Text:='';
        SetControlText ('TRAC_AFFAIRE','');
      end;
      SetField('RAC_AVENANT','');
    end;
  end;
{$ENDIF}

	if Planning or NoCreat then
  begin
  	SetControlEnabled ('BInsert',False);
   	SetControlEnabled ('BDUPLICATION',False);
  end;
	GestionBoutonGED;
	{$IFDEF SAV}
	if Assigned(Ecran) and (Ecran.Name <> 'RQACTION_FIC') then
  	MajMenuParc;
{$ENDIF}
	if (Ecran is TFFiche) and (THEdit(GetControl('HEUREACTION')) <> nil) then
  begin
  	Old_HeureAction := FormatDateTime('hh:nn', GetField('RAC_HEUREACTION'));
   	SetControlText('HEUREACTION', Old_HeureAction);
  end;
  if stProduitPGI<>'QNC' then
	begin
		ActionNonCloturee:= (GetField('RAC_ETATACTION')<>'ZCL');
    { Gestion des champs modifiables }
    SetControlsEnabled(['RAC_TIERS','RAC_TYPEACTION','RAC_LIBELLE','LECONTACT','TELEPHONE','RAC_INTERVENANT','RAC_NIVIMP'], ActionNonCloturee);
    SetControlsEnabled(['RAC_OPERATION','RAC_PROJET','NUM_PERSPECTIVE', 'CHAINAGE','RAC_AFFAIRE','RAC_AFFAIRE1', 'RAC_AFFAIRE2', 'RAC_AFFAIRE3','RAC_AVENANT','IDENTPARC'], ActionNonCloturee);
    SetControlsEnabled(['RAC_DATEACTION','HRUREACTION','RAC_NIVIMP', 'RAC_DELAIRAPPEL','RAC_GESTRAPPEL'], ActionNonCloturee);
    SetControlsEnabled(['HEUREACTION','DUREEACT','RAC_DATEECHEANCE','RAC_COUTACTION', 'CHRONO'], ActionNonCloturee);
    SetControlsEnabled(['BSELECTAFF1'], ActionNonCloturee);
  end
  else
  begin
		ActionNonCloturee:= (GetField('RAC_ETATACTION')<>'ZCL') and (not ClotureEncours);
    SetControlsEnabled(['RAC_TYPEACTION','RAC_LIBELLE','RAC_INTERVENANT','RAC_QUALITETYPE','RAC_NIVIMP'], ActionNonCloturee);
    SetControlsEnabled(['RAC_QDEMDEROGNUM', 'RAC_TIERS','RAC_AFFAIRE1', 'RAC_DATEACTION','RAC_DATEECHEANCE','RAC_HEUREACTION'], ActionNonCloturee);
    SetControlsEnabled(['RAC_AREALISERPAR','DUREEACT','RAC_QPCTAVANCT', 'RAC_TYPELOCA'], ActionNonCloturee);
    SetGBClotureEnabled(ActionNonCloturee and (TypeAction<>TaConsult));
  end;
if ModifLot then SetArguments(StSQL);
{$ENDIF EAGLSERVER}
end;

procedure TOM_ACTIONS.OnNewRecord;
//var StIntervint : string;
begin
	inherited;
 	AsInsert := true;
 	if Planning then
 		SetField('RAC_TIERS', TobAct.GetValue('RAC_TIERS'))
 	else if stProduitPgi<>'QNC' then
 		SetField('RAC_TIERS',stTiers);

  if (GetField('RAC_TIERS') <> '') then
  begin
    SetField('RAC_AUXILIAIRE',TiersAuxiliaire (GetField('RAC_TIERS'), False));
  end;
  SetField('RAC_PROJET',stProjet);
  SetField('RAC_OPERATION',stOperation);
  SetField('RAC_CHAINAGE','---');
{$IFNDEF EAGLSERVER}
  if (Ecran is TFSaisieList) then
    begin
    if (copy(ecran.name,1,17) <> 'WINTERVENTION_FSL') then
      begin
      iNumChainage:=TF.TOBFiltre.GetValue('RCH_NUMERO');
      if ( (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) ) and
         ( TF.TOBFiltre.GetValue('RCH_AFFAIRE') <> '' ) then
             SetField('RAC_AFFAIRE',TF.TOBFiltre.GetValue('RCH_AFFAIRE'));
      end
    else
      iNumChainage:=TF.TOBFiltre.GetValue('WIV_NUMCHAINAGE');
    end;
{$ENDIF EAGLSERVER}
  SetField('RAC_NUMCHAINAGE',iNumChainage);
  SetField('RAC_DATEECHEANCE', iDate2099);
  SetField('RAC_ETATACTION', PREVUE);

 	if Planning then
 	begin
  	SetField('RAC_DATEACTION', TobAct.GetValue('RAC_DATEACTION'));
    SetControlEnabled('RAC_DATEACTION',False) ;
  end
  else
  	SetField('RAC_DATEACTION', V_PGI.DateEntree);

  if Planning then
  begin
    if TobAct.GetValue('RAC_HEUREACTION') = IDate1900 then SetField('RAC_HEUREACTION',EncodeTime(0,0,0,0))
    else
    begin
    	SetField('RAC_HEUREACTION', TobAct.GetValue('RAC_HEUREACTION'));
      SetControlEnabled('RAC_HEUREACTION',False) ;
    end;
  end
  else
    SetField('RAC_HEUREACTION', Time);

  if (Planning) then
  begin
    If (TobAct.GetValue('RAC_TYPEACTION') <> '') then
    begin
      SetField('RAC_TYPEACTION', TobAct.GetValue('RAC_TYPEACTION'));
      SetControlEnabled('RAC_TYPEACTION',False) ;
      LookEcran(true);
    end
    else
    begin
      SetControlEnabled('RAC_TYPEACTION',True) ;
      LookEcran(true);
    end;
  end
  //FV1 : 29/11/2016 - FS#2249 - MATFOR - on ne peut renseigner le type d'action en création d'une action
  else if typeaction <> tacreat then SetControlEnabled('RAC_TYPEACTION',False) ;


  if Planning           then stIntervenant := TobAct.GetValue('RAC_INTERVENANT');

  if stProduitpgi='GRF' then
    SetField('RAC_PRODUITPGI','GRF')
  else
  begin
    if stProduitpgi='QNC' then
      SetField('RAC_PRODUITPGI','QNC')
    else
      SetField('RAC_PRODUITPGI','GRC');
  end;

  if stIntervenant <> '' then
    SetField('RAC_INTERVENANT',stIntervenant)
  else
    if soRtactresp = 'UTI' then SetField('RAC_INTERVENANT',VH_RT.RTResponsable);

  if (GetField ('RAC_INTERVENANT') <> '') then
  begin
    Old_Intervenant := GetField ('RAC_INTERVENANT');
//    StIntervint := GetField ('RAC_INTERVENANT') + ';';
//    SetField ('RAC_INTERVINT',StIntervint);
    StIntervint := GetField ('RAC_INTERVENANT');
  end;

  if iPerspective <> 0 then SetField('RAC_PERSPECTIVE',iPerspective);

{$IFDEF QUALITE}
{$ELSE QUALITE}
{$IFNDEF EAGLSERVER}
  if (Ecran is TFFiche) then                    //FQ10654 - TJA
    SetControlEnabled('RAC_TIERS', True);       //champ rac_tiers accessible après "nouveau" depuis une fiche action
  if (Ecran<>Nil) then SetControlText ('LECONTACT','');
{$ENDIF EAGLSERVER}
{$ENDIF QUALITE}
  Old_Typeaction := '';
  if iNumeroContact <> 0 then SetField('RAC_NUMEROCONTACT',iNumeroContact);
  { dans le cas où la liste entête est vide et que l'on en créé une, si on veut créer
    un détail on ne peut pas car modifautorise=false}
{$IFNDEF EAGLSERVER}
  if (TF <> Nil) then
     ModifAutorisee:=true;
{$ENDIF EAGLSERVER}
  if stAffaire <> '' then SetField('RAC_AFFAIRE',stAffaire);

  if Planning then
  begin
    if TobAct.GetValue('RAC_DUREEACTION') > 720 then
      SetField ('RAC_DUREEACTION', 720)
    else if TobAct.GetValue('RAC_DUREEACTION') < 30 then
      SetField ('RAC_DUREEACTION', 30)
    else
      SetField ('RAC_DUREEACTION', TobAct.GetValue('RAC_DUREEACTION'));
  end
  else
    SetField ('RAC_DUREEACTION',30);

  if iIdentParc <> 0 then SetField('RAC_IDENTPARC',iIdentParc);
{$IFDEF QUALITE}
  if stProduitPGI <> 'QNC' then
  begin
  	if (Ecran is TFFiche) then
    	SetControlEnabled('RAC_TIERS',True) ;
	  if (Ecran<>Nil) then
    	SetControlText ('LECONTACT','');
  end
  else   { Module qualité }
  begin
  	SetField('RAC_LIGNEORDRE'					,0);
  	SetField('RAC_QDEMDEROGNUM'  			,0);

		SetField('RAC_QORIGINERQ', iif(sQNCOrigine='RQN', 'RQN', iif(sQNCOrigine='RQP', 'RQP','')));
  	SetField('RAC_QNCNUM'							,iQncNum);
    SetControlText('RAC_QNCNUM'				, intToStr(iQncNum));

  	SetField('RAC_QPLANCORRNUM'				,iPlanCorrNum);
  	SetControlText('RAC_QPLANCORRNUM'	,intToStr(iPlanCorrNum));

    SetField('RAC_TYPELOCA', iif(stTiers<>'', '002','001'));
    SetControlText('RAC_TYPELOCA', iif(stTiers<>'', '002','001'));

		IF sQNCOrigine<>'RQP' then
    begin
	  	SetField('RAC_QUALITETYPE'			 , '001');
    	SetControlText('RAC_QUALITETYPE' , '001');
    end;
  end;
{$ENDIF QUALITE}
end;

procedure TOM_ACTIONS.OnUpdateRecord;
var Q : TQuery;
    nbact:integer;
    Select,NomChamp : string;
    HeureAct : TDateTime;
    EnvoiEmail : boolean;
    i{$IFNDEF EAGLSERVER},j{$ENDIF EAGLSERVER} : integer;
    TobRessourcesOrigine,TobRessources,{$IFNDEF EAGLSERVER}TI,{$ENDIF EAGLSERVER}TR,TobTemp : Tob;
    TobCreatYplanning,TobModifYplanning : Tob;
    Trouve : Boolean;
    Ressource : string;
    Resultat : integer;
    SurBooking : integer;
    TiersDuProjet : string; 
{$IFDEF QUALITE}

    function GetCleWOL: tCleWOL;
    begin
    	result.NatureTravail:= GetField('RAC_NATURETRAVAIL');
      result.LigneOrdre		:= GetField('RAC_LIGNEORDRE');
    end;
{$ENDIF QUALITE}
begin
inherited;
DuplicAct := False;
{$IFDEF QUALITE}
	if (GetField('rac_produitPGI')='QNC') and (not Assigned(ecran)) then // QNC = Module Qualité
  	ModifAutorisee:= true;
{$ENDIF QUALITE}
{ pour parer au cas de la première action d'un chainage, qui ne fonctionne pas}
	if ModifAutorisee = false then
	begin
    LastError := 21;
{$ifdef GIGI}
    LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
   	exit;
  end;
{$IFNDEF EAGLSERVER}
	if ( Ecran is TFSaisieList) and (GetField('RAC_LIBELLE')='') then
  begin
  	RACSetFocusControl('RAC_LIBELLE');
  	LastError := 23;
{$ifdef GIGI}
    LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
  	exit;
  end;
{$ENDIF EAGLSERVER}
	if (DS<>Nil) and (DS.State in [dsInsert]) then
	begin
  	if (GetField('RAC_TIERS')='') then
    begin
      if (stProduitPgi<>'QNC') or ((stProduitPgi='QNC') and sTiersCli) then
      begin
		    RACSetFocusControl('RAC_TIERS');
    	  LastError := 1;
{$ifdef GIGI}
        LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
        LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
       	exit;
      end;
    end else
    { mng 18/10/07 FQ 10741 } 
   	if ( (GetControlText ('NATUREAUXI') <> 'FOU') and
       ( ( (not VH_RT.RTCreatActions) and (not RTDroitModifTiers(GetField('RAC_TIERS')) ) ) or
         ( not RTDroitModifActions(Getfield('RAC_TIERS'),Getfield('RAC_TYPEACTION'),Getfield('RAC_INTERVENANT')))) ) or
       ( (GetControlText ('NATUREAUXI') = 'FOU') and
       ( ( (not VH_RT.RFCreatActions) and (not RTDroitModifFou(GetField('RAC_TIERS')) ) ) or
       ( not RTDroitModifActionsF(Getfield('RAC_TIERS'),Getfield('RAC_TYPEACTION'),Getfield('RAC_INTERVENANT')))) ) then
    begin
      RACSetFocusControl('RAC_TIERS');
      LastError :=14; PGIBox(TraduireMemoire(TexteMessage[LastError]),'Accès en création');
      exit;
    end;

  	if GetField('RAC_AUXILIAIRE') = '' then
    begin
	    if (stProduitPgi<>'QNC') or ((stProduitPgi='QNC') and sTiersCli) then
  	  begin
    	  RACSetFocusControl('RAC_TIERS');
      	LastError := 2;
{$ifdef GIGI}
        LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
        LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
      	exit;
      end;
    end;
{placé dans onchangefield
   if not ExisteSQL ('SELECT T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE="'+GetField('RAC_AUXILIAIRE')+'"') then
      begin
      RACSetFocusControl('RAC_TIERS');
      LastError := 2;
      LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
      exit;
      end;}
	end else
	if ( (GetControlText ('NATUREAUXI') <> 'FOU') and ( not RTDroitModifActions('',Getfield('RAC_TYPEACTION'),Getfield('RAC_INTERVENANT'))) ) or
  	 ( (GetControlText ('NATUREAUXI') = 'FOU') and ( not RTDroitModifActionsF('',Getfield('RAC_TYPEACTION'),Getfield('RAC_INTERVENANT'))) ) then
  begin
  	RACSetFocusControl('RAC_TYPEACTION');
  	LastError :=14;
  	PGIBox(TraduireMemoire(TexteMessage[LastError]),'Accès en modification');
  	exit;
  end;

	if GetField ('RAC_TYPEACTION') = '' then
  begin
  	RACSetFocusControl('RAC_TYPEACTION') ;
   	Lasterror:=6;
{$ifdef GIGI}
    LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
   	exit;
  end;

	if ((GetField ('RAC_OPERATION') <> '') and  (Not ExisteSQL ('SELECT ROP_OPERATION FROM OPERATIONS WHERE ROP_FERME<>"X" AND ROP_OPERATION="'+GetField('RAC_OPERATION')+'" AND ROP_PRODUITPGI="'+stProduitpgi+'"'))) then
  begin
  	RACSetFocusControl('RAC_OPERATION');
   	LastError := 4;
{$ifdef GIGI}
    LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
   	exit;
  end;

	if GetField ('RAC_DATEACTION') = iDate1900 then
  begin
  	RACSetFocusControl('RAC_DATEACTION') ;
   	Lasterror:=5;
{$ifdef GIGI}
    LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
   	exit;
  end;

	if  (GetField ('RAC_GESTRAPPEL') = 'X') and (GetField ('RAC_HEUREACTION') = iDate1900) then
  begin
   	RACSetFocusControl('RAC_HEUREACTION') ;
   	Lasterror:=17;
{$ifdef GIGI}
    LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
   	exit;
  end;

	if  ( (isFieldModified( 'RAC_GESTRAPPEL' )) or (isFieldModified( 'RAC_DELAIRAPPEL' ))
     or (isFieldModified( 'RAC_DATEACTION' )) or (isFieldModified( 'RAC_HEUREACTION' )) ) and
  	  (GetField ('RAC_GESTRAPPEL') = 'X' ) and ( GetField ('RAC_HEUREACTION') <> iDate1900 )
    {and ( GetField ('RAC_DELAIRAPPEL') <> '' )} then
  begin
    HeureAct:=GetField ('RAC_DATEACTION')+GetField ('RAC_HEUREACTION');
    // calcul différent suivant que l'on traite des heures ou des jours.
    if GetField ('RAC_DELAIRAPPEL') = '' then
       SetField('RAC_DATERAPPEL',HeureAct)
    else
    begin
      if GetField ('RAC_DELAIRAPPEL') < '001' then
        SetField('RAC_DATERAPPEL',HeureAct-EncodeTime(0, StrToInt(copy(GetField ('RAC_DELAIRAPPEL'),2,2)), 0, 0))
      else
        if GetField ('RAC_DELAIRAPPEL') < '024' then
          SetField('RAC_DATERAPPEL',HeureAct-EncodeTime(GetField ('RAC_DELAIRAPPEL'), 0, 0, 0))
        else
          SetField('RAC_DATERAPPEL',PlusDate(HeureAct, (GetField ('RAC_DELAIRAPPEL')/24)* (-1),'J'));
    end;
  end;

	if soRtactgestech then
  begin
  	if GetField ('RAC_DATEECHEANCE') < GetField ('RAC_DATEACTION') then
    begin
      RACSetFocusControl('RAC_DATEECHEANCE') ;
      Lasterror:=3;
{$ifdef GIGI}
      LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
      LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
      exit;
    end;
  end;

	if GetField ('RAC_ETATACTION') = '' then
  begin
  	RACSetFocusControl('RAC_ETATACTION') ;
   	Lasterror:=7;
{$ifdef GIGI}
    LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
   	exit;
  end;

	if GetField ('RAC_PROJET') <> '' then
  begin
  	if Not ExisteSQL ('SELECT Rpj_tiers FROM projets WHERE rpj_projet="'+GetField ('RAC_PROJET')+'"') then
    begin
      RACSetFocusControl('RAC_PROJET');
      LastError := 8;
{$ifdef GIGI}
      LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
      LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
      exit;
    end;
  end;

	if soRtprojmultitiers = False then
  begin
  	if ((GetField ('RAC_PROJET') <> '') and
      (Not ExisteSQL ('SELECT RPJ_PROJET FROM PROJETS WHERE (RPJ_PROJET="'+
       GetField('RAC_PROJET')+'" and RPJ_AUXILIAIRE="'+GetField('RAC_AUXILIAIRE')+'")'))) then
    begin
      TiersDuProjet := '';
      Q:=OpenSql('SELECT RPJ_TIERS FROM projets WHERE RPJ_PROJET="'+GetField ('RAC_PROJET')+'"',TRUE);
      try
        if not Q.EOF then TiersDuProjet:= Q.FindField('RPJ_TIERS').asstring;
      finally
        ferme(Q);
      end;  
      RACSetFocusControl('RAC_PROJET');
      if TiersDuProjet = '' then
      begin
        LastErrorMsg := '';
        LastError := 40;
  	    PGIBox(TexteMessage[LastError],'');
      end
      else
      begin
        LastError := 39;
{$ifdef GIGI}
        LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
        LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
      end;
      exit;
 		end;
  end;

	if ((GetField ('RAC_PERSPECTIVE') <> 0) and (Not ExisteSQL ('SELECT RPE_LIBELLE FROM PERSPECTIVES WHERE (RPE_PERSPECTIVE='+IntToStr(GetField('RAC_PERSPECTIVE'))+' and RPE_AUXILIAIRE="'+GetField('RAC_AUXILIAIRE')+'")'))) then
  begin
  	RACSetFocusControl('NUM_PERSPECTIVE');
   	LastError := 9;
{$ifdef GIGI}
    LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
   	exit;
  end;

	if (GetField ('RAC_NUMACTGEN') = 0) and (GetField ('RAC_INTERVENANT') = '') then
  begin
  	RACSetFocusControl('RAC_INTERVENANT') ;
   	Lasterror:=15;
{$ifdef GIGI}
    LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
   	exit;
  end;

	if ((GetField ('RAC_INTERVENANT') <> '') and  (Not ExisteSQL ('SELECT ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE="'+GetField('RAC_INTERVENANT')+'"'))) then
  begin
  	RACSetFocusControl('RAC_INTERVENANT');
   	LastError := 16;
{$ifdef GIGI}
    LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
   	exit;
  end;
//cd 140301 if ((GetField ('RAC_INTERVINT') <> '<<Tous>>') and (GetField ('RAC_INTERVENANT') <> '')) then MajIntervint();
{$IFNDEF EAGLSERVER}
	if (GetField ('RAC_NUMCHAINAGE') <> 0) and ( ecran is TFFiche) then
  begin
  	if Not ExisteSQL ('SELECT Rch_tiers FROM actionschainees WHERE rch_numero='+
      IntToStr(GetField ('RAC_NUMCHAINAGE'))+' and rch_tiers="'+GetField ('RAC_TIERS')+'"') then
    begin
      RACSetFocusControl('RAC_TIERS');
      LastError := 18;
{$ifdef GIGI}
      LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
      LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
      exit;
    end;
  end;
{$ENDIF EAGLSERVER}
{if (ModifResp='X') and (Getfield ('RAC_ETATACTION')=REALISEE)
   and (DS.State in [dsInsert]) then
      begin
      RACSetFocusControl('RAC_ETATACTION');
      LastError := 19;
      LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
      exit;
      end;}
    NomChamp:='';
{$IFNDEF EAGLSERVER}
    NomChamp:=VerifierChampsObligatoires(Ecran,stProduitpgi);
    if NomChamp<>'' then
    begin
      NomChamp:=ReadTokenSt(NomChamp);
      RACSetFocusControl(NomChamp) ;
      LastError:=22;
      LastErrorMsg:=TexteMessage[LastError]+champToLibelle(NomChamp);
      exit;
    end;
{$ENDIF EAGLSERVER}
	if (GetFieldAvantModif('RAC_OPERATION') <> '') and (Getfield('RAC_OPERATION') <> GetFieldAvantModif('RAC_OPERATION')) then
  	SetField ('RAC_NUMACTGEN',0);

  // CONTRÖLE SI RESSOURCES LIBRES PAR RAPPORT AUX AUTRES PLANNINGS
  if (DS <> Nil ) and (RTActEstPlanifiable (GetField('RAC_TYPEACTION')) = True) then
    begin
    if ControleEstLibre() = False then
      begin
      SurBooking := ConfirmeSurBooking;
      if SurBooking < 0 then
        begin
        LastError := 37;
        if SurBooking = -2 then PGIInfo(TexteMessage[LastError],'');
        exit;
        end;
      end;
    end;

{$IFNDEF EAGLSERVER}
  if  (Assigned(Ecran)) and (not V_Pgi.SilentMode) and (AlerteActive(TableToPrefixe(TableName))) and ( not ModifLot ) then
       if (not Inserting) then
         begin
         if not TestAlerteAction (CodeModification+';'+CodeModifChamps) then
           begin
           LastError:=99;
           exit;
           end;
         end
       else
         if not TestAlerteAction (CodeCreation) then
           begin
           LastError:=99;
           exit;
           end;
{$ENDIF !EAGLSERVER}
	{$IFDEF AFFAIRE}
  {if (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) then
    begin
    bProposition := (GetField('RAC_AFFAIRE0') = 'P');
    SetField('RAC_AFFAIRE',DechargeCleAffaire(THEDIT(GetControl('RAC_AFFAIRE0')),
                THEDIT(GetControl('RAC_AFFAIRE1')),THEDIT(GetControl('RAC_AFFAIRE2')),
              THEDIT(GetControl('RAC_AFFAIRE3')),THEDIT(GetControl('RAC_AVENANT')),
              GetField('RAC_TIERS'), taModif,True,True,bProposition,iP));
    end;}
	{$ENDIF}
  if getfield ('RAC_NUMACTION') = 0 then
  begin
    Select := 'SELECT MAX(RAC_NUMACTION) FROM ACTIONS WHERE RAC_AUXILIAIRE = "'+getfield ('RAC_AUXILIAIRE')+'"';
    Q:=OpenSQL(Select, True);
    try
      if not Q.Eof then
         begin
         nbact := Q.Fields[0].AsInteger;
         setfield ('RAC_NUMACTION',Nbact+1);
         end
         else
         setfield ('RAC_NUMACTION',1);
    finally
      Ferme(Q) ;
    end;
  end;

{$IFNDEF EAGLSERVER}
	if (ecran <> nil) and (ecran is TFFiche) and (ModifLot = False) then
  	if Planning then TFFiche(ecran).retour:=Getfield ('RAC_AUXILIAIRE')+';'+intTostr(getField('RAC_NUMACTION'))
   	else TFFiche(ecran).retour:=intTostr(getField('RAC_NUMACTION'));
	SaisieAction:=True;

	If (ds<>nil) and (ecran is TFFiche) then
  begin
  	if (DS.State in [dsInsert]) then
    begin
    	DerniereCreateTiers := GetField('RAC_AUXILIAIRE');
     	DerniereCreateNum   := GetField('RAC_NUMACTION');
    end
    else if (DerniereCreateTiers = GetField('RAC_AUXILIAIRE')) and (DerniereCreateNum = getField('RAC_NUMACTION')) then Ecran.Close; // le bug arrive on se casse !!!
  end;
{$ENDIF EAGLSERVER}

{$IFDEF QUALITE}
	{ Spécifique module qualité }
  if (stProduitPGI='QNC') or ((GetField('RAC_PRODUITPGI')='QNC') and (not Assigned(ecran))) then // QNC = Module Qualité
  begin
    if (GetField('RAC_TIERS') <> '') then
    begin
      if not ExistTiers('', GetField('RAC_TIERS')) then
      begin
        RACSetFocusControl('RAC_TIERS');
        LastError := 2;
      end
//GP_20080625_DS_GP14715
      else if (GetField('RAC_TYPEACTION')='DER') and (not ExistTiers('CLI', GetField('RAC_TIERS'))) then
      begin
        RACSetFocusControl('RAC_TIERS');
        LastError := 38;
      end
    end;

    if (GetField('RAC_QNCNUM')=0) and (GetField('RAC_QPLANCORRNUM')=0) then
    begin
 	 		RACSetFocusControl('RAC_QNCNUM');
 			LastError := 41;
    end
    else if (GetField('RAC_QDEMDEROGNUM')<>0) and (Not ExisteDEMDEROG(GetField('RAC_QDEMDEROGNUM'), GetField('RAC_QNCNUM'))) then
   	begin
   		RACSetFocusControl('RAC_QDEMDEROGNUM');
      SetControlProperty('RAC_QDEMDEROGNUM', 'ENABLED', True);
	  	LastError := 27;
    end
		else if (GetField('RAC_QNCNUM') <> 0) and (Not ExistQNC(GetField('RAC_QNCNUM'),0)) then
   	begin
 	 		RACSetFocusControl('RAC_QNCNUM');
 			LastError := 30;
		end
		else if (GetField('RAC_QPLANCORRNUM') <> 0) and (Not ExistRQP(GetField('RAC_QPLANCORRNUM'),0)) then
   	begin
 	 		RACSetFocusControl('RAC_QPLANCORRNUM');
 			LastError := 31;
		end
		else if (GetFamAction(GetField('RAC_TYPEACTION'))='DER') then
    begin
      {Une action rattachée à un plan correcteur ne peut être de type "Demande de dérogation"}
      if GetField('RAC_QPLANCORRNUM')<>0 then
      begin
        RACSetFocusControl('RAC_TYPEACTION');
        LastError := 42;
      end
    end
		else if (GetFamAction(GetField('RAC_TYPEACTION'))='WPR') and ((GetField('RAC_LIGNEORDRE')=0) or (GetField('RAC_LIGNEORDRE')=null)) then
    begin
   		RACSetFocusControl('RAC_LIGNEORDRE');
   		LastError := 34;
    end
		else if (GetField('RAC_LIGNEORDRE')<>0) and (GetField('RAC_LIGNEORDRE')<>null) and (not wExistWOL(GetCleWOL)) then
    begin
   		RACSetFocusControl('RAC_LIGNEORDRE');
   		LastError := 33;
    end
    else if GetField('RAC_QUALITETYPE') = '' then
    begin
   		RACSetFocusControl('RAC_QUALITETYPE');
   		LastError := 28;
    end
		else if (GetField('RAC_CLOTUREEPAR') <> '') and (not wExistRessource(GetField('RAC_CLOTUREEPAR'),'','',True)) then
    begin
   		RACSetFocusControl('RAC_CLOTUREEPAR');
   		LastError := 29;
    end
		else if (GetField('RAC_REALISEEPAR') <> '') and (not wExistRessource(GetField('RAC_REALISEEPAR'),'','',True)) then
    begin
   		RACSetFocusControl('RAC_REALISEEPAR');
   		LastError := 29;
    end
		else if (GetField('RAC_VERIFIEEPAR') <> '') and (not wExistRessource(GetField('RAC_VERIFIEEPAR'),'','',True)) then
    begin
   		RACSetFocusControl('RAC_VERIFIEEPAR');
   		LastError := 29;
    end
		else if (GetField('RAC_EFFJUGEEPAR') <> '') and (not wExistRessource(GetField('RAC_EFFJUGEEPAR'),'','',True)) then
    begin
   		RACSetFocusControl('RAC_EFFJUGEEPAR');
   		LastError := 29;
    end
		else if (GetField('RAC_AREALISERPAR') <> '') and (not wExistRessource(GetField('RAC_AREALISERPAR'),'','',True)) then
    begin
   		RACSetFocusControl('RAC_AREALISERPAR');
   		LastError := 29;
    end
		else if ((GetField('RAC_ETATACTION')='ZCL') or ClotureEncours) and (assigned(Ecran))then
	  begin
      if GetField('RAC_CLOTUREEPAR')='' then
      begin
      	RACSetFocusControl('RAC_CLOTUREEPAR');
      	LastError:=35
      end
	    else if GetField('RAC_CLOTUREELE')=iDate1900 then
  	 	begin
      	RACSetFocusControl('RAC_CLOTUREELE');
	      LastError:=36
      end;
    end
		else if GetField('RAC_LIBELLE') = '' then
  	begin
      RACSetFocusControl('RAC_LIBELLE');
	    LastError := 23
    end
    ;
	  if LastError > 0 then
  	begin
      {$ifdef GIGI}
      LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
      {$else}
      LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
      {$endif}
    	if not Assigned(ecran) then
	    begin
  	    fTob.AddChampSupValeur('Error', LastErrorMsg, false);
        {$ifdef GIGI}
    	  if V_Pgi.Sav then PgiInfoAf(LastErrorMsg, 'uTomAction.UpdateRecord');
        {$else}
    	  if V_Pgi.Sav then PgiInfo(LastErrorMsg, 'uTomAction.UpdateRecord');
        {$endif}
    	end;

      EXIT;

  	end;

   	if (assigned(ecran)) and (LastError=0) and (TFFiche(Ecran).fTypeAction = taCreat) then
  	begin
     { Jeton }
     SetIdentifiant;
      {Origine de laction (Non conformité ou plan correcteur)}
      SetField('RAC_QORIGINERQ', iif(GetField('RAC_QPLANCORRNUM')<>0, 'RQP', 'RQN') );
   	end;
  end;
  {$ENDIF QUALITE}

{$IFNDEF EAGLSERVER}
	if (ds<>nil) and (DS.State in [dsInsert]) then
  begin
  { on ne positionne le RV Outlook que si le responsable de l'action est bien
    l'utilisateur connecté }
  if (OutLook = 'X') and (GetField('RAC_INTERVENANT')=VH_RT.RTResponsable) then
    RTAjouteOutlook('RDV');
  if EnvoiMail = 'X' then
  begin
  	if MailResp = RESPONSABLE then
    	if GetField ('RAC_INTERVENANT') <> '' then ActEnvoiMessage ('R',MailAuto,GetField('RAC_ETATACTION'),true)
      else PGIBox('Envoi de message impossible : absence de responsable','')
    else
      if GetField ('RAC_NUMEROCONTACT') <> 0 then ActEnvoiMessage ('I',MailAuto,GetField('RAC_ETATACTION'),true)
      else PGIBox('Envoi de message impossible : absence d''interlocuteur','');
    end;
 	end;
{$ENDIF EAGLSERVER}

	if (OldEtatAction <> '') and (Getfield('RAC_ETATACTION') <> OldEtatAction) then
  begin
	  EnvoiEmail:=False; MailResp:=INTERLOCUTEUR; MailAuto:='-';
  	if (Getfield('RAC_ETATACTION') = PREVUE) and (MailEtatPre = 'X') then
    begin
      EnvoiEmail:=True;
      MailAuto:=MailPreAuto;
      MailResp:=MailRespPre;
    end;
  	if (Getfield('RAC_ETATACTION') = ANNULEE) and (MailEtatAnu = 'X') then
    begin
      EnvoiEmail:=True;
      MailAuto:=MailAnuAuto;
      MailResp:=MailRespAnu;
    end;

  	if (Getfield('RAC_ETATACTION') = REALISEE) and (MailEtatRea = 'X') then
    begin
      EnvoiEmail:=True;
      MailAuto:=MailReaAuto;
      MailResp:=MailRespRea;
    end;

  	if EnvoiEmail then
    	if MailResp = RESPONSABLE then
        if GetField ('RAC_INTERVENANT') <> '' then ActEnvoiMessage ('R',MailAuto,GetField('RAC_ETATACTION'),true)
        else PGIBox('Envoi de message impossible : absence de Responsable','')
     	else
        if GetField ('RAC_NUMEROCONTACT') <> 0 then ActEnvoiMessage ('I',MailAuto,GetField('RAC_ETATACTION'),true)
        else PGIBox('Envoi de message impossible : absence d''interlocuteur','');
  	{ on ne positionne le RV Outlook que si le responsable de l'action est bien
    l'utilisateur connecté }
{$IFNDEF EAGLSERVER}
  	if (Outlook='X') and (OldEtatAction = NONREALISABLE) and
     (Getfield('RAC_ETATACTION') = PREVUE) and (GetField('RAC_INTERVENANT')=VH_RT.RTResponsable) then
    	RTAjouteOutlook('RDV');
{$ENDIF EAGLSERVER}
  end;

  { Pour le cas où l'on a créé une action sur une intervention sans chainage, je le crée par défaut }
{$IFNDEF EAGLSERVER}
//GP_20071123_TP_GP14528 <<<
  if (Assigned(Ecran)) and (DS.State in [dsInsert]) and (GetField('RAC_NUMCHAINAGE') = 0) and
     (Ecran<>nil) and ( Ecran is TFSaisieList ) and
     (copy(ecran.name,1,17) = 'WINTERVENTION_FSL') then
    begin
    SetField('RAC_NUMCHAINAGE',GenerChainageIntervention);
    TF.TobFiltre.Putvalue('WIV_NUMCHAINAGE', GetField('RAC_NUMCHAINAGE'));
    TF.TobFiltre.UpdateDB(false,false);
    end;
{$ENDIF EAGLSERVER}
  if (stProduitPGI<>'QNC') or ((GetField('RAC_PRODUITPGI')<>'QNC') and (not Assigned(ecran))) then // QNC = Module Qualité
  begin
    ListeFinRessources:='';
{$IFNDEF EAGLSERVER}
    if (DS <> Nil) then
    begin
      for i := 1 to G2.RowCount-1 do
      begin
        TI:=TOB(G2.Objects[0,i]);
        if TI<> nil then ListeFinRessources:=ListeFinRessources+TOB(G2.Objects[0,i]).GetValue('ARS_RESSOURCE')+';';
      end;
    end
    else ListeFinRessources:=GetField ('RAC_INTERVENANT')+';';   // Cas du courrier
{$ENDIF EAGLSERVER}
    if ((DS <> Nil) and (ds.state in [dsinsert])) or (ListeInitRessources <> ListeFinRessources)
     or ModifChampYplanning then
    begin
    TobCreatYplanning := Tob.create('CREATIONYPLANNING',Nil,-1) ;
    TobModifYplanning := Tob.create('MODIFYPLANNING',Nil,-1) ;
    TobRessourcesOrigine:=Tob.create('les LIENSACTIONS',Nil,-1) ;
    Q:=OpenSql('SELECT * FROM ACTIONINTERVENANT WHERE RAI_AUXILIAIRE = "'+GetField('RAC_AUXILIAIRE')+'" AND RAI_NUMACTION ='+IntToStr(GetField('RAC_NUMACTION')),true);
    try
     TobRessourcesOrigine.loaddetailDB('ACTIONINTERVENANT','','',Q ,false,False);
    finally
     ferme(Q);
    end;
    { mng 11/02/08 : actions crées sans écran, ex CTI, ne pas accéder aux objets écrans }
    if not Assigned (DS) then ListeInitRessources := ListeFinRessources;
    if ListeInitRessources <> ListeFinRessources then
      begin
      for i:=0 to TobRessourcesOrigine.detail.count-1 do
          begin
          Trouve:=false;
{$IFNDEF EAGLSERVER}
          for j:=1 to (G2.RowCount-1) do
             begin
             TI:=TOB(G2.Objects[0,j]);
             if (TI <> nil) then
               begin
               if TobRessourcesOrigine.detail[i].GetValue('RAI_RESSOURCE') = TOB(G2.Objects[0,j]).GetValue('ARS_RESSOURCE') then
                  begin
                  Trouve:=True;
                  break;
                  end;
                end;
             end;
{$ENDIF EAGLSERVER}
             // Ressource existante à l'origine a été enlevée : détruire l'enreg. ds LIENSACTIONS
             if not Trouve then
                begin
                ExecuteSQl('DELETE FROM ACTIONINTERVENANT WHERE RAI_AUXILIAIRE = "'+TobRessourcesOrigine.detail[i].GetValue('RAI_AUXILIAIRE')+'" '
                  +' AND RAI_NUMACTION='+IntToStr(TobRessourcesOrigine.detail[i].GetValue('RAI_NUMACTION'))
                  +' AND RAI_RESSOURCE="'+TobRessourcesOrigine.detail[i].GetValue('RAI_RESSOURCE')+'"') ;
                DeleteYPL ('RAI',TobRessourcesOrigine.Detail[i].GetString('RAI_GUID'));
                end;
          end;
      end;
    TobRessources := TOB.Create ('les ressources', Nil, -1);
    While ListeFinRessources <> '' do
        begin
        Ressource :=ReadTokenSt(ListeFinRessources);
        if (Ressource <> '') then
          begin
          TobTemp := TobRessourcesOrigine.FindFirst (['RAI_AUXILIAIRE','RAI_NUMACTION','RAI_RESSOURCE'],[GetField('RAC_AUXILIAIRE'),GetField('RAC_NUMACTION'),Ressource],TRUE);
          if TobTemp = Nil then
             begin
             TR:=Tob.create('ACTIONINTERVENANT',TobRessources,-1);
             TR.PutValue('RAI_AUXILIAIRE', GetField('RAC_AUXILIAIRE'));
             TR.PutValue('RAI_NUMACTION', GetField('RAC_NUMACTION'));
             TR.PutValue('RAI_RESSOURCE', Ressource);
             TR.PutValue('RAI_GUID', AglGetGuid());
             if RTActEstPlanifiable (GetField('RAC_TYPEACTION')) then
               begin
               if (GetField('RAC_ETATACTION') <> ANNULEE) then
                 begin
                 MAJYPlanning (TobCreatYplanning,Ressource,TR.GetString('RAI_GUID'));
                 end;
               end;
             end
          else
             // Modification de YPLANNING
             begin
             if isFieldModified( 'RAC_TYPEACTION') then
               begin
                 if (RTActEstPlanifiable (TypeActionInitial) = True) and (RTActEstPlanifiable (GetField('RAC_TYPEACTION')) = False) then
                   begin
                   DeleteYPL ('RAI',TobTemp.GetString('RAI_GUID'));
                   end
                 else
                   if (RTActEstPlanifiable (TypeActionInitial) = False) and (RTActEstPlanifiable (GetField('RAC_TYPEACTION')) = True) then
                     begin
                     // Création
                     if (GetField('RAC_ETATACTION') <> ANNULEE) then MAJYPlanning (TobCreatYplanning,Ressource,TobTemp.GetString('RAI_GUID'));
                     end
                   else
                     if (RTActEstPlanifiable (TypeActionInitial) = True) and (RTActEstPlanifiable (GetField('RAC_TYPEACTION')) = True) then
                       begin
                       if (GetField('RAC_ETATACTION') = ANNULEE) then DeleteYPL ('RAI',TobTemp.GetString('RAI_GUID'))
                         else if EtatActionInitial = ANNULEE then MAJYPlanning (TobCreatYplanning,Ressource,TobTemp.GetString('RAI_GUID'))
                         // Modif.
                           else MAJYPlanning (TobModifYplanning,Ressource,TobTemp.GetString('RAI_GUID'));
                       end;
               end
             else
               begin
                 if (RTActEstPlanifiable (GetField('RAC_TYPEACTION')) = True) then
                   begin
                   if isFieldModified( 'RAC_ETATACTION') then
                     begin
                       if (GetField('RAC_ETATACTION') = ANNULEE) then DeleteYPL ('RAI',TobTemp.GetString('RAI_GUID'))
                         else if EtatActionInitial = ANNULEE then MAJYPlanning (TobCreatYplanning,Ressource,TobTemp.GetString('RAI_GUID'))
                         // Modif.
                           else MAJYPlanning (TobModifYplanning,Ressource,TobTemp.GetString('RAI_GUID'));
                     end
                   else MAJYPlanning (TobModifYplanning,Ressource,TobTemp.GetString('RAI_GUID'));
                   end;
               end;
             end; // fin // Modification de YPLANNING
          end; // Fin if ressource # ''
        end; // Fin While
     if TobRessources.Detail.Count <> 0 then
         TobRessources.InsertDB(Nil);
     if TobCreatYplanning.Detail.Count <> 0 then
       begin
         Resultat := CreateYPL(TobCreatYplanning);
         if resultat <> 0 then
            PGIBox('Erreur lors de l''enregistrement dans la table YPLANNING. ' + GetLastErrorMsgYPL(resultat), '');
       end;
     if TobModifYplanning.Detail.Count <> 0 then
       begin
         Resultat := UpdateYPL(TobModifYplanning);
         if resultat <> 0 then
            PGIBox('Erreur lors de l''enregistrement dans la table YPLANNING. ' + GetLastErrorMsgYPL(resultat), '');
       end;
     TobRessourcesOrigine.free;
     TobRessources.free;
     TobCreatYplanning.free;
     TobModifYplanning.free;
    end;
{$IFNDEF EAGLSERVER}
    if ModifLot then TFFiche(ecran).BFermeClick(nil);
{$ENDIF EAGLSERVER}
  end;
end; { fin OnUpdateRecord }

{$IFDEF AFFAIRE}
procedure TOM_ACTIONS.RechercheAffaire (Sender: TObject);
var    Q : TQuery;
     stTiers : string;
    bProposition,NoChangeStatut : boolean;
begin
    { sauvegarde du code tiers qui disparait si aucune affaire n'ai sélectionnée }
    stTiers:=GetField('RAC_TIERS');
    bProposition:=false;
    NoChangeStatut:=VH_GC.GASeria;
    if stNatureauxi <> '' then
       begin
       if pos (stNatureauxi,FabriqueWhereNatureAuxiAff('PRO')) = 0 then
          begin
          bProposition:=false;
          NoChangeStatut:=false;
          end;
       if pos (stNatureauxi,FabriqueWhereNatureAuxiAff('AFF')) = 0 then
          begin
          bProposition:=true;
          NoChangeStatut:=false;
          end;
       end;
    GetAffaireEntete(THEDIT(GetControl('RAC_AFFAIRE')), THEDIT(GetControl('RAC_AFFAIRE1')),
                   THEDIT(GetControl('RAC_AFFAIRE2')), THEDIT(GetControl('RAC_AFFAIRE3')),
                   THEDIT(GetControl('RAC_AVENANT')),
                   THEDIT(GetControl('RAC_TIERS')), bProposition, NoChangeStatut, false, false, true,'');
    if (THEDIT(GetControl('RAC_AFFAIRE')).text <> '') then
    begin
    ForceUpdate;
    SetField('RAC_AUXILIAIRE',TiersAuxiliaire (THEDIT(GetControl('RAC_TIERS')).text, False));
    SetField('RAC_TIERS',THEDIT(GetControl('RAC_TIERS')).text);  //mcd 22/12/2004 pour que cela passe par le Onchange sur le tiers..
    SetField('RAC_AFFAIRE',THEDIT(GetControl('RAC_AFFAIRE')).Text);
    ChargeCleAffaire(THEDIT(GetControl('RAC_AFFAIRE0')),THEDIT(GetControl('RAC_AFFAIRE1')), THEDIT(GetControl('RAC_AFFAIRE2')),
       THEDIT(GetControl('RAC_AFFAIRE3')), THEDIT(GetControl('RAC_AVENANT')),Nil , taModif,THEDIT(GetControl('RAC_AFFAIRE')).text,True);

    SetField('RAC_AFFAIRE0',THEDIT(GetControl('RAC_AFFAIRE0')).Text);
    SetField('RAC_AFFAIRE1',THEDIT(GetControl('RAC_AFFAIRE1')).Text);
    SetField('RAC_AFFAIRE2',THEDIT(GetControl('RAC_AFFAIRE2')).Text);
    SetField('RAC_AFFAIRE3',THEDIT(GetControl('RAC_AFFAIRE3')).Text);
    if THEDIT(GetControl('RAC_AVENANT')).Text ='' then THEDIT(GetControl('RAC_AVENANT')).Text:='00'; //mcd 22/12/2004 si avenant pas gérer, il faut le mettre à 00, ce qui n'est pas fait dans ChargeCLeAffaire
    SetField('RAC_AVENANT',THEDIT(GetControl('RAC_AVENANT')).Text);
     Q:= OpenSQL ('SELECT AFF_LIBELLE,AFF_NUMEROCONTACT FROM AFFAIRE WHERE AFF_AFFAIRE="'+THEDIT(GetControl('RAC_AFFAIRE')).text+'"',True);
     try
       if (not Q.Eof) then
          begin
          SetControlText ('TRAC_AFFAIRE',Q.FindField('AFF_LIBELLE').asstring);
          if (Q.FindField('AFF_NUMEROCONTACT').asinteger <> 0) and ( Getfield ('RAC_NUMEROCONTACT') = 0 ) then
             begin
             SetField ('RAC_NUMEROCONTACT', Q.FindField('AFF_NUMEROCONTACT').asinteger);
             NomContact();
             end;
          end
       else
          begin
          RACSetFocusControl('BSELECTAFF1');
          LastError := 24;
{$ifdef GIGI}
          LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
          LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
          PGIBox(TraduireMemoire(TexteMessage[LastError]),'Saisie code affaire');
          EffaceAffaire(Sender);
          end;
     finally
       Ferme(Q);
     end;  
    end
    else
        begin
        ForceUpdate;
        SetField('RAC_TIERS',stTiers);
        end;
end;

procedure TOM_ACTIONS.EffaceAffaire (Sender: TObject);
begin
  if ( TF <> Nil ) then
      TF.Edit
  else
      ForceUpdate;
  SetControlText ('RAC_AFFAIRE',''); SetControlText ('RAC_AFFAIRE0',''); SetControlText('RAC_AFFAIRE1','');
  SetControlText ('RAC_AFFAIRE2',''); SetControlText ('RAC_AFFAIRE3',''); SetControlText ('RAC_AVENANT','');
  SetField ('RAC_AFFAIRE',''); SetField ('RAC_AFFAIRE0',''); SetField('RAC_AFFAIRE1','');
  SetField ('RAC_AFFAIRE2',''); SetField ('RAC_AFFAIRE3',''); SetField ('RAC_AVENANT','');
  SetControlText ('TRAC_AFFAIRE','');
end;

{$ENDIF}
procedure TOM_ACTIONS.OnAfterUpdateRecord ;
var
  MailEnvoye,stChainage{$IFNDEF EAGLSERVER},chpChainage{$ENDIF EAGLSERVER},mess,StPiecesVisees : String;
  i : integer;
  // CRM_20080901_MNG_FQ;012;10843
  {$IFNDEF EAGLSERVER}TobEcran,{$ENDIF EAGLSERVER}TobTypeEncours,TobPiece,TobActionSuiv,TobRessources : tob;
  {$IFNDEF EAGLSERVER}ExisteOblig,{$ENDIF EAGLSERVER}bTF : boolean;
  Q : TQuery;
{$IFDEF QUALITE}
  Sql        : string;
  DemDerogNum: integer;
{$ENDIF QUALITE}

{$IFDEF QUALITE}
  function GetArgument: string;
  begin
    Result := 'WHERE=(RAC_AUXILIAIRE="' + GetField('RAC_AUXILIAIRE') + '" AND RAC_NUMACTION='+intToStr(GetField('RAC_NUMACTION'))+')'
  end;
{$ENDIF QUALITE}
begin
  Inherited ;
  if Ecran = Nil then exit;
  bTF:=false;
{$IFNDEF EAGLSERVER}
  if ( Ecran is TFSaisieList ) then bTF:=true;
  TobEcran:=nil;

  if bTF then
    begin
    if copy(ecran.name,1,17) <> 'WINTERVENTION_FSL' then
      chpChainage:='RCH_CHAINAGE' else chpChainage:='WIV_CHAINAGE';
    stChainage:= TF.TOBFiltre.GetString(chpChainage);
    end
  else
    stChainage:= GetField('RAC_CHAINAGE');
{$ENDIF EAGLSERVER}

  VH_RT.TobTypesChainage.Load;

  { sur option du modèle chainage : dans le cas d'action chainée,
    si la suivante est non réalisable, on la passe encours !
    IMPORTANT : Ds le cas de modification en série les automatismes liés aux
                actions chainées ne st pas réalisés }
  if stProduitPGI<>'QNC' then
  begin
    if ModiFLot = False then
    begin
  	if ( Getfield('RAC_ETATACTION')=REALISEE) and
    	 ((OldEtatAction <> '') and (Getfield('RAC_ETATACTION') <> OldEtatAction)) and
     	 (stChainage <> '') and (stChainage <> '---') then
    begin
       TobTypeEncours:=VH_RT.TobTypesChainage.FindFirst(['RPG_CHAINAGE','RPG_PRODUITPGI'],[stChainage,stProduitpgi],TRUE) ;
       if (TobTypeEncours = Nil) then
         begin
         mess:='Le modèle de chainage '+stChainage+' est inexistant.';
         PgiBox (mess,'Modèle de chaînage');
         exit;
         end;
       TobActionSuiv:=TOB.Create ('ACTIONS', Nil, -1);
{$IFNDEF EAGLSERVER}
       if bTF then
         begin
         { recherche action suivante, indépendant de l'ordre d'affichage TF.Next; }
         TobEcran := TF.TobFiltre.FindFirst(['RAC_NUMLIGNE' ],[Getfield('RAC_NUMLIGNE')+1],false);
         if Assigned(TobEcran) then
           TobActionSuiv.Dupliquer(TobEcran, False, True);
         end
       else
{$ENDIF EAGLSERVER}
         begin
         Q:=OpenSQL ('SELECT * FROM ACTIONS WHERE RAC_AUXILIAIRE="'+GetField('RAC_AUXILIAIRE')
           +'" AND RAC_NUMACTION='+IntToStr(GetField('RAC_NUMACTION')+1),true);
         TobActionSuiv.SelectDB ('', Q);
         Ferme(Q);
         end;

       if (Assigned(TobActionSuiv)) and (TobActionSuiv.GetString('RAC_ETATACTION') = NONREALISABLE )
          and (TobTypeEncours.GetValue('RPG_NREAPRE') = 'X' ) then
         begin
         { envoi mail et RV Outlook si c'est le cas }
         LookEcran(False);
         MailEnvoye:='-';
         if (MailEtatPre = 'X') then
           begin
            MailAuto:=MailPreAuto;
            MailResp:=MailRespPre;

            if MailResp = RESPONSABLE then
              if TobActionSuiv.GetString ('RAC_INTERVENANT') <> '' then MailEnvoye:=ActEnvoiMessage ('R',MailAuto,PREVUE,false)
              else PGIBox(TraduireMemoire('Envoi de message impossible : absence de responsable'),'')
            else
              if TobActionSuiv.GetInteger ('RAC_NUMEROCONTACT') <> 0 then MailEnvoye:=ActEnvoiMessage ('I',MailAuto,PREVUE,false)
              else PGIBox(TraduireMemoire('Envoi de message impossible : absence d''interlocuteur'),'');
           end;
         { maj Action suivante }
         TobActionSuiv.SetString('RAC_ETATACTION',PREVUE);
         if MailEnvoye='X' then
           begin
           TobActionSuiv.SetString('RAC_MAILENVOYE','X');
           TobActionSuiv.SetDateTime('RAC_DERNIERMAIL',NowH);
           TobActionSuiv.SetString('RAC_MAILAUTO','X');
           end;
         TobActionSuiv.UpdateDB(false,false);
         // Mise à jour de YPLANNING
         if RTActEstPlanifiable (TobActionSuiv.GetString ('RAC_TYPEACTION')) then
           begin
           TobRessources:=Tob.create('les LIENSACTIONS',Nil,-1) ;
           Q:=OpenSql('SELECT RAI_RESSOURCE,RAI_GUID,ARS_LIBELLE FROM ACTIONINTERVENANT LEFT JOIN RESSOURCE ON ARS_RESSOURCE = RAI_RESSOURCE WHERE RAI_AUXILIAIRE = "'+TobActionSuiv.GetString ('RAC_AUXILIAIRE')+'" AND RAI_NUMACTION ='+IntToStr(TobActionSuiv.GetInteger ('RAC_NUMACTION')),True);
           try
            TobRessources.loaddetailDB('','','',Q ,false,False);
           finally
            ferme(Q);
           end;
           RTMAJYPlanning (TobRessources,IDate1900,IDate1900,'',TobActionSuiv.GetString('RAC_ETATACTION'));
           FreeAndNil (TobRessources);
           end;

{$IFNDEF EAGLSERVER}
         if bTF then
           begin
           { rafraichissement de l'écran }
           if Assigned(TobEcran) then
             begin
             TobEcran.LoadDb; // recharge la tob depuis la table
             TF.RefreshControls; // refresh l'affichage de la grille et des controls
             TF.SelectRecord(TobEcran.GetInteger('RAC_NUMLIGNE'));
             end;
           if (Outlook='X') and (GetField('RAC_INTERVENANT')=VH_RT.RTResponsable) then
             RTAjouteOutlook('RDV');
           end;
{$ENDIF EAGLSERVER}
         end;
         TobActionSuiv.free;
    end;
   { rafraichissement de l'écran : mng 15/03/06 : pourquoi une deuxième fois ?
   TF.TobFiltre.LoadDb;
   TF.RefreshControls;}
  OldEtatAction:=GetField('RAC_ETATACTION');
  NouvelleAction := False;
  if NonSupp='X' then SetControlEnabled('BDelete',False);
  { fermeture automatique du chainage si option sur modèle et toutes les actions réalisées }
  if not ( bTF ) and ( GetField('RAC_NUMCHAINAGE') <> 0 ) and
     ( GetField('RAC_CHAINAGE') <> '---' ) and ( GetField('RAC_CHAINAGE') <> '' )
     and (OldEtatAction = REALISEE) then
    begin
    TobTypeEncours:=VH_RT.TobTypesChainage.FindFirst(['RPG_CHAINAGE','RPG_PRODUITPGI'],
           [GetField('RAC_CHAINAGE'),stProduitpgi],TRUE) ;
    if Assigned(TobTypeEncours) then
    begin
      if TobTypeEncours.GetValue('RPG_FERMEAUTO') = 'X' then
      begin
        if not ExisteSQL ('SELECT RAC_NUMACTION FROM ACTIONS WHERE RAC_NUMCHAINAGE='+IntToStr(GetField('RAC_NUMCHAINAGE'))+' AND RAC_ETATACTION <> "'+REALISEE+'"') then
          begin
          if ExecuteSQL('UPDATE ACTIONSCHAINEES SET RCH_TERMINE="X" WHERE RCH_NUMERO='+
            IntToStr(GetField('RAC_NUMCHAINAGE'))+' AND RCH_TERMINE<>"X"') > 0 then
            begin
            { visa des pièces si modèle paramétré pour }
            if TobTypeEncours.GetValue('RPG_VISA') = 'X' then
              begin
              TobPiece := TOB.Create ('les liens', Nil, -1);
              try
                Q:=OpenSQL('SELECT RLC_INDICEG,RLC_NATUREPIECEG,RLC_NUMERO,RLC_PRODUITPGI,RLC_SOUCHE FROM CHAINAGEPIECES WHERE RLC_NUMCHAINAGE='
                     +intToStr(GetField('RAC_NUMCHAINAGE'))+' and RLC_PRODUITPGI="'+stProduitpgi+'"',true);
                try
                  TobPiece.LoadDetailDB('CHAINAGEPIECES', '', '',Q ,false);
                finally
                  Ferme(Q);
                end;

                if TobPiece.detail.count <> 0 then
                  for i:=0 to Pred(TobPiece.detail.count) Do
                    begin
                    SetPieceVisee (TobPiece.detail[i].GetString('RLC_NATUREPIECEG'),TobPiece.detail[i].GetString('RLC_SOUCHE'),
                                  TobPiece.detail[i].GetString('RLC_NUMERO'),TobPiece.detail[i].GetString('RLC_INDICEG'));
                          if StPiecesVisees = '' then
                             StPiecesVisees:=TraduireMemoire(StVisee);
                          StPiecesVisees:=StPiecesVisees+#13#10+'  '+RechDom('GCNATUREPIECEG',TobPiece.detail[i].GetString('RLC_NATUREPIECEG'),False)+TraduireMemoire(' n° : ')+TobPiece.detail[i].GetString('RLC_NUMERO')
                    
                    end;
              finally
                TobPiece.free;
              end;
            end;
            if StPiecesVisees<>'' then StPiecesVisees:=StPiecesVisees+#13#10;
            if StPiecesVisees <> '' then
               PgiInfo (StPiecesVisees,Ecran.Caption);
            end;
            { maj intervention liée si existe : uniquement si GRC ..}
            if stProduitpgi = 'GRC' then
                ExecuteSQL('UPDATE WINTERVENTION SET WIV_ETATINTERV="TER",WIV_DATEFIN="'+UsDateTime(Date)
                   +'",WIV_DATEMODIF="'+UsDateTime(Date)+'",WIV_UTILISATEUR="'+V_PGI.USer
                   +'" WHERE WIV_NUMCHAINAGE='+IntToStr(GetField('RAC_NUMCHAINAGE')));
            end;
            end;
          end;
      end;
    end; // Fin if modiflot = False
  	if GetParamSocSecur('SO_RTGESTIONGED',False) and isFieldModified('RAC_NUMCHAINAGE') then
    begin
    	if ExisteSQL('SELECT RTD_DOCID FROM RTDOCUMENT WHERE RTD_TIERS="'+GetField('RAC_TIERS')+'" AND RTD_NUMACTION= '+IntToStr(GetField ('RAC_NUMACTION'))) then
      	ExecuteSQl('UPDATE RTDOCUMENT SET RTD_NUMCHAINAGE = ' + IntToStr(GetField('RAC_NUMCHAINAGE'))
        	      +' WHERE RTD_TIERS= "'+GetField('RAC_TIERS')
          	    +'" AND RTD_NUMACTION= '+ IntToStr(GetField('RAC_NUMACTION')));
    end;
  end;
{$IFDEF QUALITE}
  if (stProduitPGI='QNC') or ((GetField('RAC_PRODUITPGI')='QNC') and (not Assigned(ecran))) then // QNC = Module Qualité
  begin
		{Famille d'action = Demande de dérogation}
  	if (Assigned(ecran)) and (GetControlText('NATUREACTION') = 'DER')
    										 and ( (TFFiche(Ecran).fTypeAction = taCreat) and (GetField('RAC_QDEMDEROGNUM')=0) ) then
    begin
    	DemDerogNum:= CreateDemDerog(GetField('RAC_TIERS'),iQLitigieuse, GetField('RAC_IDENTIFIANT'), GetField('RAC_QNCNUM'));
      if DemDerogNum <> 0 then
	    begin
        AGLLanceFiche('RT','RQDEMDEROG_FIC', '', intToStr(DemDerogNum), 'ACTION=MODIFICATION', '');
        Sql:='UPDATE ACTIONS SET RAC_QDEMDEROGNUM='+intToStr(DemDerogNum)
      	  	+' WHERE RAC_IDENTIFIANT='+intToStr(GetField('RAC_IDENTIFIANT'));
        ExecuteSQL(Sql);
      end;
    end;
  end;

  if ClotureEncours and ClotureFromFiche then
    wDoAction(wtaClotureRAC, GetArgument, LastErrorMsg);
{$ENDIF QUALITE}
{$IFNDEF EAGLSERVER}
  { Infos complémentaires si champs obligatoire }
  if AsInsert and soRtgestinfos001 then
    begin
    ExisteOblig:=ExisteSQL('SELECT RDE_OBLIGATOIRE FROM RTINFOSDESC WHERE RDE_DESC="1" AND RDE_OBLIGATOIRE="X" ');
    if ExisteOblig then
         AGLLanceFiche('RT','RTPARAMCL','','','FICHEPARAM=RTACTIONS;FICHEINFOS='+GetField('RAC_AUXILIAIRE')+'|'+IntToStr(GetField('RAC_NUMACTION'))+';RAC_TYPEACTION='+GetField('RAC_TYPEACTION')+';EXISTOBLIG');
    end;
{$ENDIF EAGLSERVER}
  AsInsert :=false;
end ;

procedure TOM_ACTIONS.OnArgument (Arguments : String );
var Critere,ChampMul,ValMul : string;
    x : integer;
{$IFNDEF EAGLSERVER}
    Tcb:TCheckBox;
{$ENDIF EAGLSERVER}
    {$IFDEF AFFAIRE}
    bAff : TToolBarButton97;
    {$ENDIF AFFAIRE}
begin
inherited ;
  AsInsert :=false;
// CRM_20081126_MNG_FQ;012;10842
  stProduitpgi:='';

{$IFNDEF EAGLSERVER}
  if (Ecran<>nil) and (Ecran is TFSaisieList ) then
    begin
    TF := TFSaisieList(Ecran).LeFiltre;
    stProduitpgi:=Arguments;
    end
  else

	if (Ecran<>nil) and (ecran is TFFiche) then
    TypeAction:=TFFiche(ecran).TypeAction;
{$ENDIF EAGLSERVER}

{$IFDEF AFFAIRE}
	if Assigned(GetControl('BSELECTAFF1')) then
  begin
	 	bAff:=TToolBarButton97(GetControl('BSELECTAFF1'));
  	bAff.OnClick:=RechercheAffaire;
  end;
	if Assigned(GetControl('BEFFACEAFF1')) then
  begin
		bAff:=TToolBarButton97(GetControl('BEFFACEAFF1'));
		bAff.OnClick:=EffaceAffaire;
  end;
{$ENDIF}

{récup des arguments }
  stTiers         := '';
  stProjet        := '';
  stOperation     := '';
  stIntervenant   := '';
  iNumChainage    := 0;
  stTiersDuplic   :='';
  iNumActDuplic   := 0;
  NoChangeProspect:= False;
  NoChangeProjet  := False;
  NoChangeChainage:= False;
  NoChangeContact := False;
  NoChangeAffaire :=false;
  FicheChainage   := False;
  iPerspective    :=0;
  ActionOblig     :=False;
  SaisieAction    :=False;
  OrigineTiers    := False;
  iNumeroContact  :=0;
  Duplic          := False;
  StDelaiRappel   :='';
  Planning        := False;
  iIdentParc      :=0;
  NoCreat         := False;

  x := pos('MODIFLOT',Arguments);
  ModifLot := x<>0;

{$IFNDEF EAGLSERVER}
  if ModifLot then
    begin
    TFfiche(Ecran).MonoFiche:=true;
    StSQL := copy(Arguments,x+9,length(Arguments));
    x := pos('PRODUITPGI',Arguments);
    stProduitpgi := copy(Arguments,x+11,3);
    end
  else
{$ENDIF EAGLSERVER}
  begin
    //Récupération valeur de argument
    Critere:=uppercase(Trim(ReadTokenSt(Arguments)));
    //
    while (Critere <> '') do
    begin
      if Critere <> '' then
        begin
        X := pos (':', Critere) ;
        if x = 0 then
          X := pos ('=', Critere) ;
        if x <> 0 then
          begin
          ChampMul := copy (Critere, 1, X - 1) ;
          ValMul := Copy (Critere, X + 1, length (Critere) - X) ;
          ControleChamp(ChampMul, ValMul);
	  			end
        end;
      ControleCritere(Critere);
      Critere := (Trim(ReadTokenSt(Arguments)));
    end;
  end;

  {$IFDEF PAULPOURVOIR}
  OVCtl1:= TOVCtl.create(ECRAN);
  OVCtl1.parent:=TWinControl(getcontrol('AGENDA'));
  OVCtl1.folder:='Calendrier';
  OVCtl1.align:=alClient;
  {$ENDIF}

  {$IFNDEF EAGLSERVER}
  InitGrids ;
  {$ENDIF EAGLSERVER}
  if (stProduitpgi<>'GRF') and (stProduitPGI<>'QNC') then stProduitpgi:='GRC';

  if sOrigine='PGISIDE' then Exit;

  if stProduitpgi<>'QNC' then
  begin
    {$IFDEF MNG}
    LesInterlocuteurs:=tob.create('LesInterlocuteurs',Nil , -1);
    LesIntervenants:=tob.create('LesIntervenants',Nil , -1);
    {$ENDIF}
    LesInterlocuteurs:=Nil;
    LesIntervenants:=Nil;
  end;

{$IFDEF AFFAIRE}
  if ( not (ctxAffaire in V_PGI.PGIContexte) ) and
    ( not ( ctxGCAFF  in V_PGI.PGIContexte) ) or
    ( stProduitpgi='GRF') then
{$ENDIF}
  begin
    SetControlVisible ('BEFFACEAFF1',false); SetControlVisible ('BSELECTAFF1',false);
    SetControlVisible ('TRAC_AFFAIRE0',false); SetControlVisible ('RAC_AFFAIRE1',false);
    SetControlVisible ('RAC_AFFAIRE2',false); SetControlVisible ('RAC_AFFAIRE3',false);
    SetControlVisible ('RAC_AVENANT',false); SetControlVisible ('TRAC_AFFAIRE',false);
  end;

  if stProduitpgi = 'GRF' then
  begin
    SetControlText ('NATUREAUXI','FOU');
    SetControlProperty ('RAC_TIERS','Datatype','GCTIERSFOURN');
    SetControlProperty ('RAC_TYPEACTION','Datatype','RTTYPEACTIONFOU');
    SetControlProperty ('RAC_OPERATION','Datatype','RTOPERATIONSFOU');
    SetControlVisible ('RAC_PROJET',false);
    SetControlVisible ('TRAC_PROJET',false);
    SetControlVisible ('TRAC_PROJET_',false);
    SetControlVisible ('RAC_PERSPECTIVE',false);
    SetControlVisible ('TRAC_PERSPECTIVE',false);
    SetControlVisible ('NUM_PERSPECTIVE',false);

    SetControlVisible ('GBCOMPLEMENT',false);
    TPopupMenu(GetControl('POPCOMPLEMENT')).items[1].visible:=false; { proposition }
    TPopupMenu(GetControl('POPCOMPLEMENT')).items[2].visible:=false; { projet }
    TPopupMenu(GetControl('POPCOMPLEMENT')).items[4].visible:=false; { infos compl. }
    { On met les paramsoc en variable }
    soRtactgestech:=GetParamsocSecur('SO_RFACTGESTECH',False);
    soRtactgestcout:=GetParamsocSecur('SO_RFACTGESTCOUT',False);
    soRtactgestchrono:=GetParamsocSecur('SO_RFACTGESTCHRONO',False);
    soRtactrappel:=GetParamsocSecur('SO_RFACTRAPPEL',True);
    //soRtactresp:=GetParamSoc('SO_RFACTRESP');
    soRtactresp:='UTI';
    soRtprojmultitiers:=false;
    soRtgestinfos001:=false;
    soRtmesstypeact:=GetParamSocSecur ('SO_RFMESSTYPEACT',True);
    soRtmesslibact:=GetParamSocSecur ('SO_RFMESSLIBACT',True);
    soRtmessbl:=GetParamSocSecur ('SO_RFMESSBL',True);
    THPanel(GetControl('PP1')).Caption:=TraduireMemoire('Intervenants fournisseur');
    TPopupMenu(GetControl('POPCOMPLEMENT')).items[0].Caption:=TraduireMemoire('Fiche fournisseur');  // fiche client
  end
  else
  begin
    SetControlVisible ('GBCOMPLEMENTF',false);
    { On met les paramsoc en variable }
    soRtactgestech:=GetParamsocSecur('SO_RTACTGESTECH',False);
    soRtactgestcout:=GetParamsocSecur('SO_RTACTGESTCOUT',False);
    soRtactgestchrono:=GetParamsocSecur('SO_RTACTGESTCHRONO',False);
    soRtactrappel:=GetParamsocSecur('SO_RTACTRAPPEL',True);
    soRtactresp:=GetParamSocSecur('SO_RTACTRESP','UTI');
    soRtprojmultitiers:=GetParamsocSecur('SO_RTPROJMULTITIERS',True);
    soRtgestinfos001:=GetParamSocSecur('SO_RTGESTINFOS001',False);
    soRtmesstypeact:=GetParamSocSecur ('SO_RTMESSTYPEACT',True);
    soRtmesslibact:=GetParamSocSecur ('SO_RTMESSLIBACT',True);
    soRtmessbl:=GetParamSocSecur ('SO_RTMESSBL',True);
    SetControlVisible ('BAGENDAOLK',True);
  end;

  if soRtactgestech = FALSE then
  begin
    SetControlEnabled('RAC_DATEECHEANCE',FALSE) ;
    SetControlEnabled('TRAC_DATEECHEANCE',FALSE) ;
  end;

  if soRtactgestcout = FALSE then
  begin
    SetControlEnabled('RAC_COUTACTION',FALSE) ;
    SetControlEnabled('TRAC_COUTACTION',FALSE) ;
  end;

  if soRtactgestchrono = FALSE then
  begin
    SetControlEnabled('CHRONO',FALSE) ;
    SetControlEnabled('TRAC_CHRONO',FALSE) ;
    SetControlEnabled('BCHRONO',FALSE) ;
  end;

  if soRtactrappel = FALSE then
  begin
    SetControlEnabled('RAC_GESTRAPPEL',FALSE) ;
    SetControlEnabled('RAC_DELAIRAPPEL',FALSE) ;
  end
{$IFNDEF EAGLSERVER}

else
   begin
   Tcb:=TCheckBox(GetControl('RAC_GESTRAPPEL'));
   Tcb.OnClick:=RappelClick;
   end;
    GereIsoflex;
THEdit(GetControl('RAC_DATEACTION')).OnExit := OnExit_DateAction;
{$ENDIF EAGLSERVER}
{$Ifdef GIGI}
 if (GetControl('TRAC_NUMEROCONTACT') <> nil) then SetControlText('TRAC_NUMEROCONTACT','Contact');
 TPopupMenu(GetControl('POPNEWMAIL')).items[0].caption := TraduireMemoire('au contact'); // attention si MnCOntact change de position
 if (GetControl('TRAC_AFFAIRE0') <> nil) then  SetControlText('TRAC_AFFAIRE0',TraduireMemoire('Mission'));
 if (GetControl('RAC_OPERATION') <> nil) and Not GetParamsocSecur('SO_AFRTOPERATIONS',False)
    then  begin
    SetControlVisible('RAC_OPERATION',false);
    SetControlVisible('TRAC_OPERATION',false);
    AddRemoveItemFromPopup ('POPCOMPLEMENT','MNOPERATION',False);
    end;
 if (GetControl('RAC_PROJET') <> nil) and Not GetParamsocSecur('SO_RTPROJGESTION',False)
    then  begin
    SetControlVisible('RAC_PROJET',false);
    SetControlVisible('TRAC_PROJET',false);
    AddRemoveItemFromPopup ('POPCOMPLEMENT','MNPROJET',False);
    end;
 If (Not VH_GC.GaSeria) or not (GetParamSocSecur ('SO_AFRTPROPOS',False)) then
   begin
   if (GetControl('TRAC_PERSPECTIVE') <> nil) then  SetControlVisible('TRAC_PERSPECTIVE',false);
   if (GetControl('NUM_PERSPECTIVE') <> nil) then  SetControlVisible('NUM_PERSPECTIVE',false);
   if (GetControl('TRAC_AFFAIRE0') <> nil) then  SetControlVisible('TRAC_AFFAIRE0',false);
   if (GetControl('RAC_AFFAIRE1') <> nil) then  SetControlVisible('RAC_AFFAIRE1',false);
   if (GetControl('RAC_AFFAIRE2') <> nil) then  SetControlVisible('RAC_AFFAIRE2',false);
   if (GetControl('RAC_AFFAIRE3') <> nil) then  SetControlVisible('RAC_AFFAIRE3',false);
   if (GetControl('RAC_AVENANT') <> nil) then  SetControlVisible('RAC_AVENANT',false);
   if (GetControl('BEFFACEAFF1') <> nil) then  SetControlVisible('BEFFACEAFF1',false);
   if (GetControl('BSELECTAFF1') <> nil) then  SetControlVisible('BSELECTAFF1',false);
   AddRemoveItemFromPopup ('POPCOMPLEMENT','MNPERSPECTIVE',False);
   end;
 If (Not GetParamSocSecur ('SO_AFRTCHAINAGE',False)) then
   begin
   AddRemoveItemFromPopup ('POPCOMPLEMENT','MNFICHECHAINAGE',False);
   AddRemoveItemFromPopup ('POPCOMPLEMENT','MNLISTECHAINE',False);
   AddRemoveItemFromPopup ('POPCOMPLEMENT','MNRATTACHER',False);
   SetControlVisible('CHAINAGE',false);
   SetControlVisible('TCHAINAGE',false);
   end;
{$endif}
{$IFNDEF EAGLSERVER}
if (Not GetParamSocSecur('SO_RTGESTIONGED',False)) or (stProduitpgi = 'GRF') then AddRemoveItemFromPopup ('POPCOMPLEMENT','MNSGEDGRC',False);
  if Assigned(GetControl('mnPiece')) then
      TMenuItem(GetControl('mnPiece')).OnClick := MnPiece_OnClick;

if Assigned(GetControl('mnSGEDGRC')) then
   TMenuItem(GetControl('mnSGEDGRC')).OnClick := RTAppelGED_OnClick;
if Assigned(GetControl('BDOCGEDEXIST')) then
   TToolbarButton97(GetControl('BDOCGEDEXIST')).OnClick := RTAppelGED_OnClick;
SetControlVisible('BDOCGEDEXIST',false);
{$IFDEF SAV}
// Gestion Parc
if Assigned(GetControl('mnParcAffectation')) then
   TMenuItem(GetControl('mnParcAffectation')).OnClick := AffecterParc;
if Assigned(GetControl('mnParcSuppression')) then
   TMenuItem(GetControl('mnParcSuppression')).OnClick := SupprimerParc;
if Assigned(GetControl('mnParcConsultation')) then
   TMenuItem(GetControl('mnParcConsultation')).OnClick := VoirFicheParc;
if (Not VH_GC.SAVSeria) or (stProduitpgi = 'GRF') then
  begin
  AddRemoveItemFromPopup ('POPCOMPLEMENT','MNPARC',False);
  SetControlVisible ('IDENTPARC',False);
  SetControlVisible ('TIDENTPARC',False);
  end;
{$else}
       //mcd 04/10/2005 si pas compile SAV, on ne veut pas les info ....
  AddRemoveItemFromPopup ('POPCOMPLEMENT','MNPARC',False);
  SetControlVisible ('IDENTPARC',False);
  SetControlVisible ('TIDENTPARC',False);
{$ENDIF}
if THEdit(GetControl('HEUREACTION')) <> nil then
  begin
  THEdit(GetControl('HEUREACTION')).OnExit := OnExit_HeureAction;
  THEdit(GetControl('HEUREACTION')).OnChange := OnChange_HeureAction;
  end;
{$ENDIF EAGLSERVER}
{$IFDEF QUALITE}
	{ Module qualité }
	if stProduitpgi='QNC' then
  begin
    if sAction = 'CONSULTATION' then
    	TypeAction:=taConsult;
//      TFFiche(ecran).TypeAction

    {$IFDEF EAGLCLIENT}
	  if Assigned(GetControl('RAC_INTERVENANT')) then
    	tHEdit(GetControl('RAC_INTERVENANT')).OnElipsisClick:= RAC_INTERVENANT_OnElipsisClick;
	  if Assigned(GetControl('RAC_AREALISERPAR')) then
    	tHEdit(GetControl('RAC_AREALISERPAR')).OnElipsisClick:= RAC_INTERVENANT_OnElipsisClick;
	  if Assigned(GetControl('RAC_CLOTUREEPAR')) then
    	tHEdit(GetControl('RAC_CLOTUREEPAR')).OnElipsisClick:= RAC_INTERVENANT_OnElipsisClick;
	  if Assigned(GetControl('RAC_REALISEEPAR')) then
    	tHEdit(GetControl('RAC_REALISEEPAR')).OnElipsisClick:= RAC_INTERVENANT_OnElipsisClick;
	  if Assigned(GetControl('RAC_VERIFIEEPAR')) then
    	tHEdit(GetControl('RAC_VERIFIEEPAR')).OnElipsisClick:= RAC_INTERVENANT_OnElipsisClick;
	  if Assigned(GetControl('RAC_EFFJUGEEPAR')) then
    	tHEdit(GetControl('RAC_EFFJUGEEPAR')).OnElipsisClick:= RAC_INTERVENANT_OnElipsisClick;
	  if Assigned(GetControl('RAC_TIERS')) then
    	tHEdit(GetControl('RAC_TIERS')).OnElipsisClick:= RAC_Tiers_OnElipsisClick;
	  if Assigned(GetControl('RAC_LIGNEORDRE')) then
    	tHEdit(GetControl('RAC_LIGNEORDRE')).OnElipsisClick:= RAC_LigneOrdre_OnElipsisClick;
    {$ELSE EAGLCLIENT}
	  if Assigned(GetControl('RAC_INTERVENANT')) then
    	tHDBEdit(GetControl('RAC_INTERVENANT')).OnElipsisClick:= RAC_INTERVENANT_OnElipsisClick;
	  if Assigned(GetControl('RAC_AREALISERPAR')) then
    	tHDBEdit(GetControl('RAC_AREALISERPAR')).OnElipsisClick:= RAC_INTERVENANT_OnElipsisClick;
	  if Assigned(GetControl('RAC_CLOTUREEPAR')) then
    	tHDBEdit(GetControl('RAC_CLOTUREEPAR')).OnElipsisClick:= RAC_INTERVENANT_OnElipsisClick;
	  if Assigned(GetControl('RAC_REALISEEPAR')) then
    	tHDBEdit(GetControl('RAC_REALISEEPAR')).OnElipsisClick:= RAC_INTERVENANT_OnElipsisClick;
	  if Assigned(GetControl('RAC_VERIFIEEPAR')) then
    	tHDBEdit(GetControl('RAC_VERIFIEEPAR')).OnElipsisClick:= RAC_INTERVENANT_OnElipsisClick;
	  if Assigned(GetControl('RAC_EFFJUGEEPAR')) then
    	tHDBEdit(GetControl('RAC_EFFJUGEEPAR')).OnElipsisClick:= RAC_INTERVENANT_OnElipsisClick;
	  if Assigned(GetControl('RAC_TIERS')) then
    	tHDBEdit(GetControl('RAC_TIERS')).OnElipsisClick:= RAC_Tiers_OnElipsisClick;
	  if Assigned(GetControl('RAC_LIGNEORDRE')) then
    	tHDBEdit(GetControl('RAC_LIGNEORDRE')).OnElipsisClick:= RAC_LigneOrdre_OnElipsisClick;
    {$ENDIF EAGLCLIENT}

	  if Assigned(GetControl('MNPROSPECT')) then
    	tMenuItem(GetControl('MNPROSPECT')).OnClick:= MnProspect_OnClick;
	  if Assigned(GetControl('MNCONTACTS')) then
    	tMenuItem(GetControl('MNCONTACTS')).OnClick:= MnContacts_OnClick;
	  if Assigned(GetControl('MNFICHEINFOS')) then
    	tMenuItem(GetControl('MNFICHEINFOS')).OnClick:= MnFicheInfos_OnClick;
	  if Assigned(GetControl('MNCLOTURERAC')) then
    	tMenuItem(GetControl('MNCLOTURERAC')).OnClick:= MnClotureRac_OnClick;
	  if Assigned(GetControl('MNDEMDEROG')) then
    	tMenuItem(GetControl('MNDEMDEROG')).OnClick:= MnDemDerog_OnClick;
	  if Assigned(GetControl('MNLPNONCONF')) then
    	tMenuItem(GetControl('MNLPNONCONF')).OnClick:= MnLpNonConf_OnClick;
	  if Assigned(GetControl('MNLPPLANCORR')) then
    	tMenuItem(GetControl('MNLPPLANCORR')).OnClick:= MnLpPlanCorr_OnClick;

    if Assigned(GetControl('MNJOURNALACTION')) then
      TMenuItem(GetControl('MNJOURNALACTION')).OnClick := MnJournalAction_OnClick;
    if Assigned(GetControl('MNPROPERTIES')) then
      TMenuItem(GetControl('MNPROPERTIES')).OnClick := MnProperties_OnClick;

	  if Assigned(GetControl('RAC_QDEMDEROGNUM')) then
  	{$IFDEF EAGLCLIENT}
	  	if Assigned(GetControl('RAC_QDEMDEROGNUM')) then
    		ThEdit(GetControl('RAC_QDEMDEROGNUM')).OnElipsisClick := QDEMDEROGNUM_OnElipsisClick;
	  	if Assigned(GetControl('RAC_QNCNUM')) then
    		ThEdit(GetControl('RAC_QNCNUM')).OnElipsisClick := QNCNUM_OnElipsisClick;
	  	if Assigned(GetControl('RAC_QPLANCORRNUM')) then
    		ThEdit(GetControl('RAC_QPLANCORRNUM')).OnElipsisClick := QPLANCORRNUM_OnElipsisClick;
  	{$ELSE EAGLCLIENT}
	  	if Assigned(GetControl('RAC_QDEMDEROGNUM')) then
    		ThDbEdit(GetControl('RAC_QDEMDEROGNUM')).OnElipsisClick := QDEMDEROGNUM_OnElipsisClick;
	  	if Assigned(GetControl('RAC_QNCNUM')) then
    		ThDbEdit(GetControl('RAC_QNCNUM')).OnElipsisClick := QNCNUM_OnElipsisClick;
	  	if Assigned(GetControl('RAC_QPLANCORRNUM')) then
    		ThDbEdit(GetControl('RAC_QPLANCORRNUM')).OnElipsisClick := QPLANCORRNUM_OnElipsisClick;
  	{$ENDIF EAGLCLIENT}
    if Assigned(GetControl('RAC_QUALITETYPE')) then
      ThEdit(GetControl('RAC_QUALITETYPE')).OnChange := QualiteType_OnChange;
    if Assigned(GetControl('PMACTION')) then
      TPopUpMenu(GetControl('PMACTION')).OnPopUp := PmAction_OnPopUp;

  	SetControlProperty('RAC_TYPEACTION', 'PLUS', ' AND RPA_PRODUITPGI="'+stProduitpgi+'"');
    if sQNCOrigine='RQN' then
    	SetControlEnabled('RAC_QNCNUM'			, False)
    else if sQNCOrigine='RQP' then
    	SetControlEnabled('RAC_QPLANCORRNUM', False);

    if not GetParamSoc('SO_RTGESTINFOS001') then
   		SetControlVisible('mnFicheInfos', False);

 		SetControlVisible('mnProspect', GetControlText('RAC_TIERS')<>'');

    SetControlProperty('PNPRODUCTION', 'VISIBLE', (ctxGPAO in V_PGI.PGIContexte) );

    { Pour pouvoir appeler n'importe quel type de tiers et pas seulement des clients ou des prospects }
    if Assigned(GetControl('RAC_TIERS')) then
      SetControlProperty('RAC_TIERS', 'DataType', 'GCTIERS');

		{ Non gérés}
    SetControlVisible('mnRattacher'			, False);
    SetControlVisible('mnFicheChainage'	, False);
    SetControlVisible('mnDetacher'			, False);
    SetControlVisible('mnListeChaine'	  , False);
    SetControlVisible('mnPiece'			  	, False);
  end;
{$ENDIF QUALITE}
{$IFDEF GRCLIGHT}
  if ( stProduitpgi='GRC') and (not GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False)) then
    begin
    SetControlVisible('RAC_OPERATION',false);
    SetControlVisible('TRAC_OPERATION',false);
    SetControlVisible('TRAC_OPERATION_',false);
    SetControlVisible('RAC_PROJET',false);
    SetControlVisible('TRAC_PROJET',false);
    SetControlVisible('TRAC_PROJET_',false);
    SetControlVisible('MNPROJET',false);
    SetControlVisible('MNOPERATION',false);
    end;
{$ENDIF GRCLIGHT}

{$IFNDEF EAGLSERVER}
  if Assigned(GetControl('MnAlerte')) then
    if AlerteActive('RAC') then
      TMenuItem(GetControl('MnAlerte')).OnClick := Alerte_OnClick_RAC
    else
      TMenuItem(GetControl('MnAlerte')).visible:=false;

  if Assigned(GetControl('MnListAlerte')) then
    if AlerteActive('RAC') then
      TMenuItem(GetControl('MnListAlerte')).OnClick := ListAlerte_OnClick_RAC
    else
      TMenuItem(GetControl('MnListAlerte')).visible:=false;

  if Assigned(GetControl('MnGestAlerte')) and Assigned(GetControl('MnAlerte'))
     and Assigned(GetControl('MnListAlerte')) then
         TMenuItem(GetControl('MnGestAlerte')).visible := (TMenuItem(GetControl('MnAlerte')).visible)
          and (TMenuItem(GetControl('MnListAlerte')).visible);

  if stProduitPgi = 'GRC' then
    SetControlVisible ('PANCOMPL',GetControlVisible('RAC_TABLELIBRE1') or
        GetControlVisible('RAC_TABLELIBRE2') or GetControlVisible('RAC_TABLELIBRE3'))
  else if stProduitPgi = 'GRF' then
    SetControlVisible ('PANCOMPL',GetControlVisible('RAC_TABLELIBREF1') or
        GetControlVisible('RAC_TABLELIBREF2') or GetControlVisible('RAC_TABLELIBREF3'));


{$ENDIF EAGLSERVER}
end;

procedure TOM_ACTIONS.OnClose;
begin
	if ActionOblig and Not SaisieAction then
  begin
  	LastError := 11;
{$ifdef GIGI}
    LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
   	exit;
  end;
	{$IFDEF POURVOIR}
		OVCtl1.free;
	{$ENDIF}
{$IFNDEF EAGLSERVER}
	if (Ecran is TFFiche) and Assigned(Split) then
  	Split.free;
{$ENDIF EAGLSERVER}
	if LesInterlocuteurs <> Nil  then
  	LesInterlocuteurs.free;
	if LesIntervenants <> Nil then
  	LesIntervenants.free;

	inherited;
end;

procedure TOM_ACTIONS.OnDeleteRecord  ;
var Q    : TQuery;
    TobRess : Tob;
    i : integer;
begin
	Inherited ;
{$IFNDEF EAGLSERVER}
  if (not V_Pgi.SilentMode) and (not AfterInserting)
    and ( not ModifLot ) and (AlerteActive (TableToPrefixe(TableName))) then
      if (not TestAlerteAction(CodeSuppression)) then
        begin
        LastError := 99;
        exit;
        end;
{$ENDIF EAGLSERVER}
  if (NonSupp <> 'X') then
  begin
    if (GetField('RAC_PRODUITPGI') = 'GRC') and GetParamSocSecur ('SO_RTGESTINFOS001',false) then
       ExecuteSQL('DELETE FROM RTINFOS001 where RD1_CLEDATA="'+GetField('RAC_AUXILIAIRE')+';'+IntToStr(GetField('RAC_NUMACTION'))+'"') ;
  end
  else
  begin
    LastError:=20;
{$ifdef GIGI}
    LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
    exit ;
  end;
  if GetParamSocSecur('SO_RTGESTIONGED',False) then
  begin
     if ExisteSQL('SELECT RTD_DOCID FROM RTDOCUMENT WHERE RTD_TIERS="'+GetField('RAC_TIERS')+'" AND RTD_NUMACTION= '+IntToStr(GetField ('RAC_NUMACTION'))) then
        ExecuteSQl('UPDATE RTDOCUMENT SET RTD_NUMACTION = 0 WHERE RTD_TIERS= "'+GetField('RAC_TIERS')
              +'" AND RTD_NUMACTION= '+ IntToStr(GetField('RAC_NUMACTION')));
  end;
  // Destruction des enregs YPLANNING si type d'action planifiable
  if RTActEstPlanifiable (GetField('RAC_TYPEACTION')) then
{  VH_RT.TobTypesAction.Load;
  TobTypAct:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[GetField('RAC_TYPEACTION'),'---',0],TRUE) ;
  if (TobTypAct <> Nil) and (TobTypAct.GetBoolean('RPA_PLANIFIABLE') = True) then   }
  begin
    TobRess := TOB.Create ('Les LIENSACTIONS',NIL,-1);
    Q:=OpenSql('SELECT RAI_GUID FROM ACTIONINTERVENANT WHERE RAI_AUXILIAIRE = "'+GetField('RAC_AUXILIAIRE')+'" AND RAI_NUMACTION ='+IntToStr(GetField('RAC_NUMACTION')),True);
    try
     TobRess.loaddetailDB('','','',Q ,false);
    finally
     ferme(Q);
    end;
    for i:=0 to TobRess.detail.count-1 do
      begin
      DeleteYPL ('RAI',TobRess.Detail[i].GetString('RAI_GUID'));
      end;
    TobRess.Free;
  end;

  ExecuteSQL('DELETE FROM ACTIONINTERVENANT where RAI_AUXILIAIRE="'+GetField('RAC_AUXILIAIRE')+'" AND RAI_NUMACTION = '+IntToStr(GetField('RAC_NUMACTION'))) ;
{$IFDEF QUALITE}
	{ Spécifique QUALITE}
	if (LastError=0) and (GetField('RAC_PRODUITPGI')='QNC') then
  begin
  	if GetField('RAC_QDEMDEROGNUM')<>0 then
	  	LastError:= wAllowDeleteRAC(GetField('RAC_IDENTIFIANT'), GetField('RAC_QDEMDEROGNUM'), LastErrorMsg);

 		if LastError = 0 then
    begin
      { Traitement des suppressions}
      if Ecran = nil then
      begin
        if fTob.FieldExists(WRPTobFieldName) then
          wGetRapportFromTob(fTob).Add(TWRP_Done, intToStr(GetField('RAC_IDENTIFIANT')), TraduireMemoire('Fiche action ')+ string(GetField('RAC_LIBELLE'))+ TraduireMemoire(' supprimée'));

        {Suppression des demandes de dérogation}
        if GetField('RAC_QDEMDEROGNUM')<>0 then
        begin
          if (PgiAsk(TraduireMemoire('Voulez-vous supprimer la demande de dérogation associée à cette action?'), TraduireMemoire('Suppression de l''action') + ' : ' + intToStr(GetField('RAC_IDENTIFIANT')) ) = MrYes) then
            wDeleteTable('RQDEMDEROG', 'RQD_DEMDEROGNUM='+intToStr(GetField('RAC_QDEMDEROGNUM')), True)
          else
          { Suppression du n° d'action  ds table RQDEMDEROG }
            ExecuteSQL('UPDATE RQDEMDEROG'
                      + ' SET RQD_IDACTION=0'
                      + ' WHERE RQD_DEMDEROGNUM='+intToStr(GetField('RAC_QDEMDEROGNUM')));
        end;
        {Suppression des plans correcteurs}
        if GetField('RAC_QPLANCORRNUM')<>0 then
        begin
          if (PgiAsk(TraduireMemoire('Voulez-vous supprimer le plan correcteur associé à cette action?'), TraduireMemoire('Suppression de l''action') + ' : ' + intToStr(GetField('RAC_IDENTIFIANT')) ) = MrYes) then
            wDeleteTable('RQPLANCORR', 'RQP_PLANCORRNUM='+intToStr(GetField('RAC_QPLANCORRNUM')), True);
        end;
      end;
  	end
  	else if ecran = nil then   // Erreur ds le contrôle suppression
  	begin
    	if fTob.FieldExists(WRPTobFieldName) then
      	wGetRapportFromTob(fTob).Add(TWRP_Error, IntToStr(GetField('RAC_IDENTIFIANT')), LastErrorMsg);
  	end;
  end;
{$ENDIF QUALITE}
{$IFNDEF EAGLSERVER}
if (ecran is TFFiche) then TFFiche(ecran).retour:='DELETE' else
   if (ecran is TFSaisieList) then TFSaisieList(ecran).retour:='DELETE';
{$ENDIF EAGLSERVER}
end;
{$IFNDEF EAGLSERVER}
procedure TOM_ACTIONS.DuplicationAct (TiersDuplic: string;NumActDuplic:Integer);
var Q : TQuery ;
    TobDuplication : Tob;
    i,x : Integer;
    CleDuplicTiers : string;
    CleDuplicNum   : Integer;
    ActionChoisie,StArgument1,StArgument2,StFiche  : string;
    Valid : Boolean;
begin
	CleDuplicTiers := TiersDuplic;
	CleDuplicNum   := NumActDuplic;
	Valid := True;
	Q := Nil;
	TobDuplication := TOB.Create ('ACTIONS', Nil, -1);
  try
    if TobDuplication <> Nil then
    begin
      if (CleDuplicTiers = '') and (CleDuplicNum = 0) then
      begin
        if (DS.State=dsInsert) and ( Ecran is TFFiche) then  // si nouvelle fiche  recherche record à dupliquer
        begin
          StArgument1 := '';
          if (StTiers <> '') then StArgument1 := 'RAC_TIERS='+StTiers;
          if NoChangeProspect then
          begin
            StArgument2 := 'NOCHANGEPROSPECT';
            if StProduitpgi = 'GRF' then StFiche := 'RFACTIONS_DUPLI_T'
            else StFiche := 'RTACTIONS_DUPLI_T';
          end
          else
          begin
            StArgument2 := '';
            if StProduitpgi = 'GRF' then
              StFiche := 'RFACTIONS_DUPLI'
            else if StProduitpgi = 'QNC' then
              StFiche := 'RQACTION_DUPLIC'
            else StFiche := 'RTACTIONS_DUPLI';
          end;
          ActionChoisie := AGLLanceFiche('RT',StFiche,StArgument1,'',StArgument2);
          if (ActionChoisie <> '') then
          begin
            x:=pos(';',ActionChoisie);
            if x<>0 then
            begin
              CleDuplicTiers:=copy(ActionChoisie,1,x-1) ;
              CleDuplicNum:=StrToInt(copy(ActionChoisie,x+1,length(ActionChoisie))) ;
            end ;
          end;
          Q:=OpenSQL('SELECT * From ACTIONS WHERE RAC_AUXILIAIRE="'+CleDuplicTiers+'" AND RAC_NUMACTION ='+ IntToStr(CleDuplicNum),True);
        end else
        begin
          CleDuplicTiers := GetField ('RAC_AUXILIAIRE');
          CleDuplicNum   := GetField ('RAC_NUMACTION');
          if (Ecran <> Nil) and ( Ecran is TFFiche) then
            Valid := TFFiche (Ecran).Bouge (TNavigateBtn(NbInsert));
          if (Ecran <> Nil) and ( Ecran is TFSaisieList) then
            TF.Insert;
          if ((Valid) and ( Ecran is TFFiche)) or (Ecran is TFSaisieList) then
            Q:=OpenSQL('SELECT * From ACTIONS WHERE RAC_AUXILIAIRE="'+CleDuplicTiers+'" AND RAC_NUMACTION ='+ IntToStr(CleDuplicNum),True);
        end;
      end else
      begin
        Q:=OpenSQL('SELECT * From ACTIONS WHERE RAC_AUXILIAIRE="'+CleDuplicTiers+'" AND RAC_NUMACTION ='+ IntToStr(CleDuplicNum),True);
      end;
      if (Q<>Nil) and (Valid) and (not Q.Eof) then TobDuplication.SelectDB ('', Q)
      else
      begin
        CleDuplicTiers:= '';
      end;
      if Q <> Nil then Ferme(Q) ;
      if (CleDuplicTiers <> '') or (stProduitPGI='QNC') then
      begin
        //Inum:=getfield ('RAC_NUMCHAINAGE');
        SetControlEnabled('BDUPLICATION',False) ;
        Duplic := True;
        for i := 1 to TobDuplication.NbChamps do
        begin
          SetField (TobDuplication.GetNomChamp(i), TobDuplication.GetValeur (i));
        end;
        SetField ('RAC_NUMACTION',0);
      { si l'onvient d'un chainage, bien reprendre le bon code }
        //Setfield ('RAC_NUMCHAINAGE',INum);
        if stProduitPgi<>'QNC' then
        begin
          if LesInterlocuteurs <> Nil then LesInterlocuteurs.Free;
          LesInterlocuteurs:=tob.create('LesInterlocuteurs',Nil , -1);
          if LesIntervenants <> Nil then LesIntervenants.Free;
          LesIntervenants:=tob.create('LesIntervenants',Nil , -1);
          if ChargeInterlocuteurs then
          begin
            LesInterlocuteurs.PutGridDetail(G1,True,True,G1LesColonnes,True); NewLigne(G1);
          end
          else
            G1.videpile (False);
          if ChargeIntervenants (CleDuplicTiers,CleDuplicNum) then
          begin
            LesIntervenants.PutGridDetail(G2,True,True,G2LesColonnes,True); NewLigne(G2);
          end
          else
            G2.videpile (False);

          SetControlText ('DUREEACT',PCSDureeCombo(Getfield('RAC_DUREEACTION')));
          NomContact();
          SetField('RAC_CHRONOMETRE', 0);
          SetControlText('CHRONO', '0:0:0');
          if (GetField ('RAC_OPERATION') <> '') and (GetField ('RAC_NUMACTGEN') <> 0) then
          begin
            SetField ('RAC_OPERATION','');
            SetField ('RAC_NUMACTGEN',0);
          end;
        end;
        if Getfield ('RAC_TYPEACTION') <> '' then
        begin
          Old_Typeaction := Getfield ('RAC_TYPEACTION');
          if (GetField ('RAC_INTERVENANT') <> '') and (Old_Intervenant <> '') and
             (GetField ('RAC_INTERVENANT') <> Old_Intervenant) then
          begin
            Q:=OpenSQL('SELECT ARS_LIBELLE2,ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE = "'+ GetField ('RAC_INTERVENANT')+'"',True);
            try
              if not Q.Eof then
                SetControlText ('TRAC_INTERVENANT_',Q.FindField('ARS_LIBELLE2').asstring+' '+Q.FindField('ARS_LIBELLE').asstring);
            finally
              Ferme(Q);
            end;  
          end;
          Old_Intervenant:=GetField ('RAC_INTERVENANT');
          LookEcran(false);
        end;
{$IFDEF QUALITE}
        if stProduitPgi='QNC' then
        	SetField('RAC_ETATACTION', PREVUE);
{$ENDIF QUALITE}

        Duplic := False;
        SetField ('RAC_HEUREACTION',Time);
        ChargeChainage; //Chainage
        if (ModifResp='X') and (Getfield ('RAC_ETATACTION')=REALISEE) then
          SetField ('RAC_ETATACTION',Prevue);
        {$IFDEF AFFAIRE}
        if ( (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) ) and
             (GetField ('RAC_AFFAIRE') <> '') then
        begin
          ChargeCleAffaire(THEDIT(GetControl('RAC_AFFAIRE0')),THEDIT(GetControl('RAC_AFFAIRE1')), THEDIT(GetControl('RAC_AFFAIRE2')),
          THEDIT(GetControl('RAC_AFFAIRE3')), THEDIT(GetControl('RAC_AVENANT')),Nil , taModif,GetField ('RAC_AFFAIRE'),True);
          if  THEDIT(GetControl('RAC_AVENANT')).text='' then THEDIT(GetControl('RAC_AVENANT')).text:='00'; //mcd 22/12/2004
        end;
        {$ENDIF}
      end;
    end;
  finally
    TobDuplication.Free;
  end;
end;
{$ENDIF EAGLSERVER}
{$IFNDEF EAGLSERVER}
procedure TOM_ACTIONS.NomContact;
var
  Q : TQuery;
  Nom,Telephone : string;
begin
	if assigned(GetControl('LECONTACT')) then
  begin
		if (GetField('RAC_AUXILIAIRE') <> '') and (GetField('RAC_NUMEROCONTACT')<>0) then
   	begin
			Nom:='';
		  Telephone := '';
   		Q:=OpenSql('Select C_NOM,C_TELEPHONE from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'+GetField('RAC_AUXILIAIRE')+'" AND C_NUMEROCONTACT='+intToStr(GetField('RAC_NUMEROCONTACT')),TRUE);
      try
        if not Q.EOF then
        begin
          Nom:= Q.FindField('C_NOM').asstring;
          Telephone := Q.FindField('C_TELEPHONE').asstring;
        end;
      finally
        ferme(Q);
      end;  
		//   SetControlText ('LECONTACT',RTRechDomContact('T',GetField('RAC_AUXILIAIRE'),GetField('RAC_NUMEROCONTACT')))
   		SetControlText ('LECONTACT',Nom);
   		SetControlText ('TELEPHONE',Telephone);
   	end
		else
    begin
    	SetControlText ('LECONTACT','');
    	SetControlText ('TELEPHONE','');
    end;
		SetControlText ('OldContact',GetControlText('Lecontact'));
  end;
end;
{$ENDIF EAGLSERVER}

procedure TOM_ACTIONS.HeureFin;
var HeureFin: TDateTime;
    Hour, Min, Sec, MSec: Word;
begin
HeureFin:= Time;
HeureFin:=HeureFin-HeureDeb+GetField('RAC_CHRONOMETRE');
{$IFNDEF EAGLSERVER}
if not(DS.State in [dsInsert,dsEdit]) then
    if ( TF <> Nil ) then
        TF.Edit
    else
        DS.edit;
{$ENDIF EAGLSERVER}
setfield ('RAC_CHRONOMETRE',HeureFin);
DecodeTime(HeureFin, Hour, Min, Sec, MSec);
SetControlText('CHRONO', IntToStr(Hour) + ':'
    + IntToStr(Min) + ':'+ IntToStr(Sec));
Timer1.Free ;
end;

{$IFNDEF EAGLSERVER}
procedure TOM_ACTIONS.RTAjouteOutlook( Ajout : String );

  function GetMailResp (TheResp : string) : string;
  var QQ : TQuery;
  begin
    Result := '';
    QQ := OpenSQL('SELECT ARS_EMAIL FROM RESSOURCE WHERE ARS_RESSOURCE="'+TheResp+'"',True,1,'',true);
    if not QQ.eof then
    begin
      Result := QQ.fields[0].asstring;
    end;
    ferme (QQ);
  end;

var  Memo : TRichEdit;
  ThePanel:TPanel;
  Heure : String;
  TheMailresp : string;
begin
  ThePanel := TPanel.Create( nil );
  ThePanel.Visible := False;
  ThePanel.ParentWindow := GetDesktopWindow;
  Memo := TRichEdit.Create( ThePanel );
  Memo.Parent:=ThePanel;
  Memo.width:=525;
  StringTorich(Memo,GetControlText('RAC_BLOCNOTE'));
  if Ajout='RDV' then
  begin
    Heure := ActFormatHeure;
    TheMailresp := GetMailResp (GetControlText('RAC_INTERVENANT'));

    if TheMailresp <> '' then
    begin
      TRY
        AddRDVWithDest(getfield ('RAC_LIBELLE'),GetControlText('TRAC_TIERS'),Memo.Text,getfield ('RAC_DATEACTION'),StrToDateTime (Heure) ,ActFormatDuree,True,TheMailresp) ;
      finally
        PgiInfo ('Une message a été envoyé au responsable de l''action #13#10 et le rendez-vous à été rajouté dans votre calendrier');
      end;
    end else
    begin
      TRY
        AddRDV(getfield ('RAC_LIBELLE'),GetControlText('TRAC_TIERS'),Memo.Text,getfield ('RAC_DATEACTION'),StrToDateTime (Heure) ,ActFormatDuree,True) ;
      finally
        PGIInfo('Le rendez-vous à été rajouté dans votre calendrier');
      end;
    end;
  end;
  if Ajout='TACHE' then
  begin
    AddTache('Tiers : '+GetControlText('TRAC_TIERS')+', Action : '+getfield ('RAC_LIBELLE'),Memo.Text,getfield ('RAC_DATEACTION'),getfield ('RAC_DATEECHEANCE')) ;
  end;
  if Ajout='NOTE' then
  begin
    AddNote(Memo.Text,getfield ('RAC_TYPEACTION'));
  end;
  ThePanel.Free;
end;
{$ENDIF EAGLSERVER}

procedure TOM_ACTIONS.HeureDebut;
begin
HeureDeb:= Time;
Timer1:=TTimer.Create(application) ;
//if Timer1<>nil then Timer1.OnTimer:=MonOnTimer ;
//TheTime:=StrToTime(TDateTime(Getcontrol('CHRONO'))) ;
TheTime:=StrToTime('0:0:0') ;
end;

procedure MonOnTimer(Sender: TObject);
begin
//SetControlText('CHRONO', timetostr(TheTime));
end ;

{procedure TOM_ACTIONS.MajIntervint;
var StIntervint,Critere : string;
    Present : Boolean;
begin
StIntervint:=GetField ('RAC_INTERVINT');
Present := False;
While StIntervint <> '' do
    begin
    Critere :=ReadTokenSt(StIntervint);
    if (Critere = GetField ('RAC_INTERVENANT')) then
       begin
       Present := True;
       break;
       end;
    end;
if (Present = False) then
   begin
   StIntervint:=GetField ('RAC_INTERVINT');
   StIntervint := StIntervint + GetField ('RAC_INTERVENANT') + ';';
   SetField ('RAC_INTERVINT',StIntervint);
   end;
end;    }

Function TOM_ACTIONS.ActEnvoiMessage (TypeMessage,MailAuto,EtatAction:String; MajAct : boolean) : String;
//GP_20071122_TP
var StDestinataire, StSQL, StWhere, StOr, Critere, Objet : string;
    {$IFDEF QUALITE}
    NatureAuxi: string;
    {$ENDIF QUALITE}
    IContact,i : integer;
    Q : TQuery;
    TobContact,TobIntervint:TOB;
//    BODY:Tstrings;
    Memo : TMEMO;
    Liste :HTStringList;
{$IFNDEF EAGLSERVER}
    TI : TOB;
{$ENDIF EAGLSERVER}
begin
Result:='-';
TobContact := Nil;
if (TypeMessage = 'I') then
    begin
    stWhere:='';
    StOr:='';
    StDestinataire:=GetField ('RAC_DESTMAIL');
    if StDestinataire <> TraduireMemoire('<<Tous>>') then
       begin
       if (StDestinataire <> '') or (GetField ('RAC_NUMEROCONTACT') <> 0) then
          begin
          While StDestinataire <> '' do
              begin
              i :=ReadTokenI(StDestinataire);
              if i <> 0 then     // car chaîne commence par ; voir modif 0602
                begin
                StWhere := StWhere + StOr + 'C_NUMEROCONTACT='+ IntToStr(i);
                StOr := ' OR ';
                end
              end;
          if GetField ('RAC_NUMEROCONTACT') <> 0 then
             begin
             StWhere := StWhere + StOr + 'C_NUMEROCONTACT='+ IntToStr(GetField('RAC_NUMEROCONTACT'));
             end;
          end else
          begin
          StWhere := 'C_NUMEROCONTACT =' + IntToStr(0);
          end;
       end;
    StSQL:='SELECT C_NUMEROCONTACT,C_NOM,C_FONCTION,C_TELEPHONE,C_RVA'+
               ',C_CIVILITE FROM CONTACT WHERE C_TYPECONTACT = "T" AND C_AUXILIAIRE = "'+GetField('RAC_AUXILIAIRE')+'"';
    if StWhere<>'' then StSQL:=StSQL+' AND ('+stWhere+')';

    TobContact:=TOB.create ('les contacts',NIL,-1);
    Q := OpenSql (StSQL,TRUE);
    try
      TobContact.LoadDetailDB ('CONTACT', '', '', Q, False);
    finally
      ferme(Q) ;
    end;
    end;

// cd 010301 : Remplacement table Commercial par table Ressource
stWhere:='';
StOr:='';
//StDestinataire:=GetField ('RAC_INTERVINT');
Stwhere := 'WHERE ';
{$IFNDEF EAGLSERVER}
for i := 1 to G2.RowCount-1 do
  begin
  TI:=TOB(G2.Objects[0,i]);
  if TI <> nil then
    begin
     StWhere := StWhere + StOr + 'ARS_RESSOURCE= "'+TOB(G2.Objects[0,i]).GetValue('ARS_RESSOURCE')+'"';
     StOr := ' OR ';
    end;
  end;
{$ENDIF EAGLSERVER}
{While StIntervint <> '' do
    begin
    Critere :=ReadTokenSt(StIntervint);
    StWhere := StWhere + StOr + 'ARS_RESSOURCE= "'+ Critere+'"';
    StOr := ' OR ';
    end;
end;  }
if StWhere = 'WHERE ' then StWhere := 'WHERE ARS_RESSOURCE = "#######"';

if (TypeMessage <> 'I') and (GetField ('RAC_INTERVENANT') <> '') then
   begin
   StWhere := StWhere + ' or ARS_RESSOURCE= "'+ GetField('RAC_INTERVENANT')+'"';
   end;

StSQL:='SELECT ARS_RESSOURCE,ARS_LIBELLE,ARS_TYPERESSOURCE,ARS_EMAIL'+
           ' FROM RESSOURCE ';
if StWhere<>'' then StSQL:=StSQL+stWhere;

TobIntervint:=TOB.create ('les intervenants',NIL,-1);
Q := OpenSql (StSQL,TRUE);
try
  TobIntervint.LoadDetailDB ('RESSOURCE', '', '', Q, False);
finally
  ferme(Q);
end;

StDestinataire := '';
StWhere := '';
if (TypeMessage = 'I') then
begin
for i:=0 to TobContact.detail.count-1 do
    begin
    if TobContact.detail[i].GetValue('C_RVA') <> '' then
       begin
       IContact:=TobContact.detail[i].GetValue('C_NUMEROCONTACT');
       if IContact = GetField('RAC_NUMEROCONTACT') then
          begin
          StWhere := TobContact.detail[i].GetValue('C_RVA');
          end
       else
           begin
           StDestinataire := StDestinataire + TobContact.detail[i].GetValue('C_RVA') +';';
           end;
       end;
    end;
end;

//GP_20071122_TP >>>
  for i:=0 to TobIntervint.detail.count-1 do
  begin
    if TobIntervint.detail[i].GetValue('ARS_EMAIL') <> '' then
    begin
      critere := TobIntervint.detail[i].GetValue('ARS_RESSOURCE');
      if (critere = GetField('RAC_INTERVENANT')) and (TypeMessage <> 'I') then
      begin
        StWhere := TobIntervint.detail[i].GetValue('ARS_EMAIL');
      end
      else
      begin
        StDestinataire := StDestinataire + TobIntervint.detail[i].GetValue('ARS_EMAIL') +';';
      end;
    end;
  end;
//GP_20071122_TP <<<

  if StWhere <> '' then
  begin
//GP_20071122_TP >>>
    Objet := '';
    if GetControlText('TRAC_TIERS') <> '' then
    begin
      if stProduitpgi = 'GRF' then
        Objet := TraduireMemoire('Fournisseur')+' : ' + GetControlText('TRAC_TIERS')
      else if stProduitPGI = 'GRC' then
        Objet := TraduireMemoire('Client')+' : '+GetControlText('TRAC_TIERS');
    end;

    if soRtmesstypeact then
    begin
      if Objet <> '' then Objet := Objet + ' - ' + RechDom('RTTYPEACTIONALL',GetField('RAC_TYPEACTION'),FALSE)
      else Objet := RechDom('RTTYPEACTIONALL',GetField('RAC_TYPEACTION'),FALSE);
    end;

    if soRtmesslibact then
    begin
      if Objet <> '' then Objet := Objet + ' - ' + GetField('RAC_LIBELLE')
      else Objet := GetField('RAC_LIBELLE')
    end;

    if soRtmessbl then
    begin
      Memo := TMemo(GetControl('RAC_BLOCNOTE'));
      Liste := HTStringList.Create;
      try
        Liste.Text := Memo.text;
        if GetField('RAC_PRODUITPGI') <> 'QNC' then
          Liste.Insert(0,Format(TraduireMemoire('Action du %s à %s %s'),[FormatDateTime ('dd/mm/yyyy',GetField('RAC_DATEACTION')),FormatDateTime ('hh:nn:ss',GetField('RAC_HEUREACTION')),RechDom('RTETATACTION',EtatAction,FALSE)]))
        {$IFDEF QUALITE}
        else
        begin
          if GetControlText('TRAC_TIERS') <> '' then
          begin
            NatureAuxi := GetFieldFromTiers('T_NATUREAUXI', '', GetField('RAC_TIERS'));
            if NatureAuxi = 'CLI' then
              Liste.Insert(0, Format( TraduireMemoire('Client : %s'), [GetControlText('TRAC_TIERS')]))
            else if NatureAuxi = 'FOU' then
              Liste.Insert(0, Format( TraduireMemoire('Fournisseur : %s'), [GetControlText('TRAC_TIERS')]))
          end;
          if GetField('RAC_QNCNUM') <> 0 then
            Liste.Insert(0, Format( TraduireMemoire('Non conformité N° %s'), [IntToStr(GetField('RAC_QNCNUM'))]));

          Liste.Insert(0, Format(TraduireMemoire('Action N° %s du %s')
                        , [IntToStr(GetField('RAC_IDENTIFIANT')), FormatDateTime('dd/mm/yyyy', GetField('RAC_DATEACTION'))]));
        end
        {$ENDIF QUALITE}
        ;
        PGIEnvoiMail (Objet,StWhere,StDestinataire,Liste,'',(MailAuto='X'),1,'','');
      finally
        Liste.free;
      end;
    end
    else
      PGIEnvoiMail (Objet,StWhere,StDestinataire,Nil,'',(MailAuto='X'),1,'','');

    Result:='X';
    if MajAct then
    begin
      Setfield ('RAC_MAILENVOYE','X');
      Setfield ('RAC_DERNIERMAIL',Date+Time);
      Setfield ('RAC_MAILAUTO',MailAuto);
    end;
  end
  else
    PGIBox(TraduireMemoire('Envoi de message impossible : pas d''adresse e-mail'),'');

  if TypeMessage = 'I' then TobContact.free;
  TobIntervint.free;
//GP_20071122_TP <<<

end;

{$IFNDEF EAGLSERVER}
Function TOM_ACTIONS.ActFormatHeure : string;
begin
Result:=FormatDateTime ('hh:nn:ss',GetField('RAC_HEUREACTION'));
end;

Function TOM_ACTIONS.ActFormatDuree : Integer;
//var Hour,Min,Sec,Msec : Word;
begin
{DecodeTime(GetField('RAC_DUREEACT'), Hour, Min, Sec, MSec);
Result:=(Hour*60)+Min;   }
Result:=trunc(GetField('RAC_DUREEACTION'));
end;

procedure TOM_ACTIONS.ChgtTypeAction;
begin
if (GetField ('RAC_TYPEACTION') <> '') and (GetField ('RAC_TYPEACTION') <> Old_Typeaction) then
   begin
   Old_Typeaction := GetField ('RAC_TYPEACTION');
   LookEcran(true);
   end;
end;

procedure TOM_ACTIONS.RTAppelParamCL;
var StAction : string;
    TobChampsProFille : tob;
    { mng_fq012;10859 }
    bcreat:boolean;
begin
//if VH_RT.TobChampsPro.detail[1].Detail.Count = 0 then

  VH_RT.TobChampsPro.Load;
  VH_RT.TobTypesAction.Load;

  TobChampsProFille:=VH_RT.TobChampsPro.FindFirst(['CO_CODE'], ['1'], TRUE);
  if (TobChampsProFille = Nil ) or (TobChampsProFille.detail.count = 0 ) then
    begin
    PGIInfo(TraduireMemoire('Le paramétrage de cette saisie n''a pas été effectué'),'');
    exit;
    end;
  StAction:='';
  if (not self.ModifAutorisee) or (TypeAction=taConsult) then
   StAction:='ACTION=CONSULTATION;';
  if  (Ecran <> Nil) and ( Ecran is TFFiche ) then
    begin
    { mng_fq012;10859 }
    bcreat:=false;
    if ds.state = dsInsert then bcreat:=true;
    { mng_fq012;10858 }
    if not TFFiche(Ecran).Bouge(TNavigateBtn(nbPost)) then exit ;
    { mng_fq012;10859 }
    if (not bcreat) or (not Assigned(TobChampsProFille.FindFirst(['RDE_DESC','RDE_OBLIGATOIRE'],['1','X'],TRUE))) then
      AglLancefiche('RT','RTPARAMCL','','',StAction+'FICHEPARAM='+TFFiche(Ecran).Name+';FICHEINFOS='+GetField('RAC_AUXILIAIRE')+'|'+IntToStr(GetField('RAC_NUMACTION'))+';RAC_TYPEACTION='+GetField('RAC_TYPEACTION')) ;
    end;
  if (Ecran <> Nil) and ( Ecran is TFSaisieList) then
    begin
    if DS.State = dsEdit then
        TFSaisieList(Ecran).LeFiltre.Post
    else
      if DS.State = dsInsert then
        TFSaisieList(Ecran).LeFiltre.Insert;
    AglLancefiche('RT','RTPARAMCL','','',StAction+'FICHEPARAM=RTACTIONS;FICHEINFOS='+GetField('RAC_AUXILIAIRE')+'|'+IntToStr(GetField('RAC_NUMACTION'))+';RAC_TYPEACTION='+GetField('RAC_TYPEACTION')) ;
    { il faut valider la fiche pour que les modifs soit prise en compte dans le grid}
    TFSaisieList(Ecran).LeFiltre.Edit;
    TFSaisieList(Ecran).LeFiltre.Post;
    end;

end;
{$ENDIF EAGLSERVER}
procedure TOM_ACTIONS.LookEcran (Affect : Boolean);
var Q : TQUERY;
    Select{,ModifResp} : String;
begin
    VH_RT.TobTypesAction.Load;

    TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[GetField('RAC_TYPEACTION'),GetField('RAC_CHAINAGE'),GetField('RAC_NUMLIGNE')],TRUE) ;
    if TobTypActEncours = Nil then
    begin
      //if GetField('RAC_NUMLIGNE') <> 0 then {pour le cas où la ligne n'existerait plus dans les types d'actions}
      //begin
        TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[GetField('RAC_TYPEACTION'),'---',0],TRUE) ;
        if TobTypActEncours = Nil then exit;
      //end
      //else
      //  exit;
    end;

   if (TobTypActEncours.GetValue('RPA_GESTDATECH') = 'X') then
      begin
      SetControlEnabled('RAC_DATEECHEANCE',TRUE) ;
      SetControlEnabled('TRAC_DATEECHEANCE',TRUE) ;
      if (Affect) and (TobTypActEncours.GetValue('RPA_DELAIDATECH') <> 0) then
          SetField('RAC_DATEECHEANCE', RTCalculEch(GetField('RAC_DATEACTION'),StrToInt(TobTypActEncours.GetValue('RPA_DELAIDATECH')),TobTypActEncours.GetValue('RPA_WEEKEND')));
      end
   else
      begin
      SetControlEnabled('RAC_DATEECHEANCE',FALSE) ;
      SetControlEnabled('TRAC_DATEECHEANCE',FALSE) ;
      if (Affect) and (GetField('RAC_DATEECHEANCE') <> iDate2099) then  SetField('RAC_DATEECHEANCE', iDate2099);
      end;
   if (TobTypActEncours.GetValue('RPA_GESTCOUT') = 'X') then
      begin
      SetControlEnabled('RAC_COUTACTION',TRUE) ;
      SetControlEnabled('TRAC_COUTACTION',TRUE) ;
      end
   else
      begin
      SetControlEnabled('RAC_COUTACTION',FALSE) ;
      SetControlEnabled('TRAC_COUTACTION',FALSE) ;
      if (Affect) and (GetField('RAC_COUTACTION') <> 0) then  SetField('RAC_COUTACTION', 0);
      end;
   if (TobTypActEncours.GetValue('RPA_GESTCHRONO') = 'X') then
      begin
      SetControlEnabled('CHRONO',TRUE) ;
      SetControlEnabled('TRAC_CHRONO',TRUE) ;
      SetControlEnabled('BCHRONO',TRUE) ;
      end
   else
      begin
      SetControlEnabled('CHRONO',FALSE) ;
      SetControlEnabled('TRAC_CHRONO',FALSE) ;
      SetControlEnabled('BCHRONO',FALSE) ;
      if (Affect) and (GetField('RAC_CHRONOMETRE') <> 0 ) then
         begin
         SetField('RAC_CHRONOMETRE', 0);
         SetControlText('CHRONO', '0:0:0');
         end;
      end;
   if Affect then
     begin
     if (VarIsNull (TobTypActEncours.GetValue('RPA_DUREEACTION')) ) or
      (TobTypActEncours.GetValue('RPA_DUREEACTION') = 0) then
         SetField('RAC_DUREEACTION', 30)
     else
       SetField('RAC_DUREEACTION', TobTypActEncours.GetValue('RPA_DUREEACTION'));
     SetControlText ('DUREEACT',PCSDureeCombo(Getfield('RAC_DUREEACTION')));
     end;
   if (Affect) and (stProduitpgi='GRC') then
      begin
      SetField ('RAC_TABLELIBRE1',TobTypActEncours.GetValue('RPA_TABLELIBRE1'));
      SetField ('RAC_TABLELIBRE2',TobTypActEncours.GetValue('RPA_TABLELIBRE2'));
      SetField ('RAC_TABLELIBRE3',TobTypActEncours.GetValue('RPA_TABLELIBRE3'));
      end;
   if (Affect) and (stProduitpgi='GRF') then
      begin
      SetField ('RAC_TABLELIBREF1',TobTypActEncours.GetValue('RPA_TABLELIBREF1'));
      SetField ('RAC_TABLELIBREF2',TobTypActEncours.GetValue('RPA_TABLELIBREF2'));
      SetField ('RAC_TABLELIBREF3',TobTypActEncours.GetValue('RPA_TABLELIBREF3'));
      end;

   if (TobTypActEncours.GetValue('RPA_GESTRAPPEL') = 'X') then
   begin
      SetControlEnabled('RAC_GESTRAPPEL',TRUE) ;
      SetControlEnabled('RAC_DELAIRAPPEL',TRUE) ;
      if (DS<>Nil) and (DS.State in [dsInsert]) then
      begin
        SetField('RAC_GESTRAPPEL', 'X');
        SetField ('RAC_DELAIRAPPEL',TobTypActEncours.GetValue('RPA_DELAIRAPPEL'));
        SetField (StDelaiRappel,TobTypActEncours.GetValue('RPA_DELAIRAPPEL'));
      end;
   end
   else
   begin
      SetControlEnabled('RAC_GESTRAPPEL',false) ;
      SetControlEnabled('RAC_DELAIRAPPEL',false) ;
      if GetField('RAC_GESTRAPPEL') <> '-' then
      begin
         SetField('RAC_GESTRAPPEL', '-');
         SetField('RAC_DATERAPPEL', iDate1900);
      end;
   end;
   OutLook := TobTypActEncours.GetValue('RPA_OUTLOOK');
{$IFNDEF EAGLSERVER}
   if Ecran is TFSaisieList then
       begin
       EnvoiMail := TobTypActEncours.GetValue('RPA_MAILETATVAL');
       MailResp := TobTypActEncours.GetValue('RPA_MAILRESPVAL');
       MailAuto := TobTypActEncours.GetValue('RPA_MAILVALAUTO');
       end
   else
{$ENDIF EAGLSERVER}
       begin
       EnvoiMail := TobTypActEncours.GetValue('RPA_ENVOIMAIL');
       MailResp := TobTypActEncours.GetValue('RPA_MAILRESP');
       MailAuto := TobTypActEncours.GetValue('RPA_MAILRESPAUTO');
       end;

   MailEtatAnu := TobTypActEncours.GetValue('RPA_MAILETATANU');
   MailRespAnu := TobTypActEncours.GetValue('RPA_MAILRESPANU');
   MailAnuAuto := TobTypActEncours.GetValue('RPA_MAILANUAUTO');

   MailEtatPre := TobTypActEncours.GetValue('RPA_MAILETATPRE');
   MailRespPre := TobTypActEncours.GetValue('RPA_MAILRESPPRE');
   MailPreAuto := TobTypActEncours.GetValue('RPA_MAILPREAUTO');

   MailEtatRea := TobTypActEncours.GetValue('RPA_MAILETATREA');
   MailRespRea := TobTypActEncours.GetValue('RPA_MAILRESPREA');
   MailReaAuto := TobTypActEncours.GetValue('RPA_MAILREAAUTO');
   ModifResp:=TobTypActEncours.GetValue('RPA_MODIFRESP');
   NonSupp:=TobTypActEncours.GetValue('RPA_NONSUPP');

	if (ModifResp='X') and (DS.State<>dsInsert) then
  begin
    if not RTDroitModifTypeAction(GetField ('RAC_PRODUITPGI')) then
    begin
      Select := 'SELECT ARS_UTILASSOCIE FROM RESSOURCE WHERE ARS_RESSOURCE = "'+ GetField ('RAC_INTERVENANT')+'"';
      Q:=OpenSQL(Select, True,-1,'RESSOURCE');
      try
        if not Q.Eof then
          if Q.FindField('ARS_UTILASSOCIE').asString <> V_PGI.User then
          begin
            if TypeAction = taModif then
              PGIInfo(TraduireMemoire('Cette fiche est modifiable uniquement par le responsable.'),'');
            ModifAutorisee:=False;
{$IFNDEF EAGLSERVER}
            if Ecran is TFSaisieList then
              SetControlEnabled('bPost',ModifAutorisee)
            else
              SetControlEnabled('BVALIDER',ModifAutorisee);
{$ENDIF EAGLSERVER}
            SetControlEnabled('BDUPLICATION',ModifAutorisee);
            SetControlEnabled('BDelete',ModifAutorisee);
            SetControlEnabled('G1',ModifAutorisee);
            SetControlEnabled('G2',ModifAutorisee);
            SetControlEnabled('RAC_INTERVENANT',False);
{$IFNDEF EAGLSERVER}
            if (ecran is TFFiche) and (ecran.name<>'RQACTION_FIC')then
            begin
              TPopupMenu(GetControl('POPCOMPLEMENT')).items[7].enabled:=false;
              TPopupMenu(GetControl('POPCOMPLEMENT')).items[5].enabled:=false;
            end;
{$ENDIF EAGLSERVER}
          end;
      finally
        Ferme(Q);
      end;
    end;
  end;
  if NonSupp='X' then
  	SetControlEnabled('BDelete',False);
  if (Affect) and (ds.State=dsInsert) and (TobTypActEncours.GetValue('RPA_REPLICLIB') = 'X' ) then
    SetField ('RAC_LIBELLE',THValComboBox(GetControl('RAC_TYPEACTION')).Text);
  if (Affect) and (ds.State=dsInsert) and (TobTypActEncours.GetString('RPA_ETATACTION') <> '' ) then
    SetField ('RAC_ETATACTION',TobTypActEncours.GetString('RPA_ETATACTION'));
end;


// ======== gestion des intervenants et interlocuteurs =========

{$IFNDEF EAGLSERVER}
function TOM_ACTIONS.ChargeInterlocuteurs : boolean;
var Q : TQuery;
    listeContacts : string;
begin
result:=false;
//LesInterlocuteurs:=tob.create('LesInterlocuteurs',Nil , -1);  mis ds ONARGUMENT
LesInterlocuteurs.ClearDetail;
if getfield('RAC_DESTMAIL') <> '' then
   begin
   ListeContacts:=FindEtReplace(getfield('RAC_DESTMAIL'),';',',',True);
   ListeContacts:='('+copy(ListeContacts,2,Length(ListeContacts)-2)+')'; // à partir du 2eme Caract. ald du 1er et sur LG-2 ald LG-1 voir modif 0602
   Q:=OpenSql('SELECT * FROM CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'+GetField('RAC_AUXILIAIRE')+'" AND C_NUMEROCONTACT in '+ListeContacts,true);
   try
     Result:=LesInterlocuteurs.loaddetailDB('CONTACT','','',Q ,false,False,-1,0);
   finally
     ferme(Q);
   end;
   end
else
   begin
   G1.Col:=1; G1.Row:=1;
   G1.ElipsisButton:=true; G1.options:=G1.Options+[goEditing] ;G1.editormode:=True ;
   end ;
end;

function TOM_ACTIONS.ChargeIntervenants (Auxiliaire:string;NoAction:integer) : boolean ;
var Q : TQuery;
//    listeContacts : string;
begin
result:=false;
//LesIntervenants:=tob.create('LesIntervenants',Nil , -1);
LesIntervenants.ClearDetail;
if (Auxiliaire <> '') and (NoAction <> 0) then
   begin
   Q:=OpenSql('SELECT RESSOURCE.* FROM ACTIONINTERVENANT LEFT JOIN RESSOURCE on RAI_RESSOURCE = ARS_RESSOURCE  WHERE RAI_AUXILIAIRE = "'+Auxiliaire+'" AND RAI_NUMACTION ='+IntToStr(NoAction),true);
   try
     Result:=LesIntervenants.loaddetailDB('RESSOURCE','','',Q ,false,False);
   finally
     ferme(Q);
   end;
   end;
if LesIntervenants.Detail.count = 0 then
   begin
   G2.Col:=1; G2.Row:=1;
   G2.ElipsisButton:=true; G2.options:=G2.Options+[goEditing] ;G2.editormode:=True ;
   end ;       

{if getfield('RAC_INTERVINT') <> '' then
   begin
   ListeContacts:=FindEtReplace(getfield('RAC_INTERVINT'),';','","',True);
   ListeContacts:='("'+copy(ListeContacts,1,Length(ListeContacts)-2)+')';
   Q:=OpenSql('SELECT * FROM RESSOURCE where ARS_RESSOURCE  in '+ListeContacts,true,-1,'RESSOURCE');
   try
     Result:=LesIntervenants.loaddetailDB('RESSOURCE','','',Q ,false,False,-1,0);
   finally
     ferme(Q);
   end;
   end
else
   begin
   G2.Col:=1; G2.Row:=1;
   G2.ElipsisButton:=true; G2.options:=G2.Options+[goEditing] ;G2.editormode:=True ;
   end ;       }

end;

procedure TOM_ACTIONS.InitGrids;
var i : integer ;
    st,Nam : string;
begin
G1LesColonnes:='FIXED;C_NOM;C_FONCTION;C_TELEPHONE' ;
if sOrigine='PGISIDE' then
  G1 := THGrid.Create(GetControl('P1'))
else
  G1:=THGRID(GetControl('G1'));
G1.OnCellEnter:=G1CellEnter ;
G1.OnCellExit:=G1CellExit ;
G1.OnRowEnter:=GSRowEnter ;
G1.OnRowExit:=GSRowExit ;
G1.PostDrawCell:= G1DessineCell;
G1.OnElipsisClick:=G1ElipsisClick  ;
G1.ColCount:=4;
G1.ColWidths[0]:=10;
//G1.ColWidths[1]:=0;
St:=G1LesColonnes ;
for i:=0 to G1.ColCount-1 do
   BEGIN
   if i>2 then G1.ColWidths[i]:=100;
   Nam:=ReadTokenSt(St) ;
   if Nam='C_NOM' then G1ColNom:=i
   ;
   END ;
G1.options:=G1.Options-[goEditing] ;

G2LesColonnes:='FIXED;ARS_LIBELLE;ARS_FONCTION1;ARS_TELEPHONE' ;
if sOrigine='PGISIDE' then
  G2 := THGrid.Create(GetControl('P2'))
else
  G2:=THGRID(GetControl('G2'));
G2.OnCellEnter:=G2CellEnter ;
G2.OnCellExit:=G2CellExit ;
G2.OnRowEnter:=GSRowEnter ;
G2.OnRowExit:=GSRowExit ;
G2.PostDrawCell:= G2DessineCell;
G2.OnElipsisClick:=G2ElipsisClick  ;
G2.ColCount:=4;
G2.ColWidths[0]:=10;
//G2.ColWidths[1]:=0;
St:=G2LesColonnes ;
for i:=0 to G2.ColCount-1 do
   BEGIN
   if i>1 then G2.ColWidths[i]:=100;
   Nam:=ReadTokenSt(St) ;
   if Nam='ARS_LIBELLE' then G2ColNom:=i
   else if Nam='ARS_FONCTION1' then G2.ColFormats[i]:='CB='+ChampToTT(Nam)
   ;
   END ;
G2.options:=G2.Options-[goEditing] ;
if Assigned(Ecran) then
  if ecran is TFFiche then
      begin
      TFFiche(Ecran).Hmtrad.ResizeGridColumns(G1) ;
      TFFiche(Ecran).Hmtrad.ResizeGridColumns(G2) ;
      if Split = Nil then AddSplitter;
      TFFiche(Ecran).OnKeyDown:=FormKeyDown ;
      end
  else
      begin
      TFSaisieList (Ecran).Hmtrad.ResizeGridColumns(G1) ;
      TFSaisieList(Ecran).Hmtrad.ResizeGridColumns(G2) ;
      TFSaisieList(Ecran).OnKeyDown:=FormKeyDown ;
      end;
G1.Enabled :=(TypeAction<>taConsult);
G2.Enabled :=(TypeAction<>taConsult);
end;

procedure TOM_ACTIONS.G1ElipsisClick(Sender: TObject);
var NewInterlocuteur : string;
    i : Integer;
    Trouve : Boolean;
    TI : TOB;
begin
if G1.Col = G1ColNom then
    begin
    if G1.Objects[0,G1.Row] = Nil then
       begin
       NewInterlocuteur:= AGLLanceFiche ('YY', 'YYCONTACT','T;'+GetField('RAC_AUXILIAIRE'), '', 'FROMSAISIE='+G1.cells[G1ColNom,G1.Row]);
       if NewInterlocuteur<>'' then
          begin
          Trouve := False;
          for i := 1 to G1.RowCount-1 do
            begin
              TI:=TOB(G1.Objects[0,i]);
              if (TI <> nil) and (intToStr(TOB(G1.Objects[0,i]).GetValue('C_NUMEROCONTACT')) = NewInterlocuteur) then
              begin
              Trouve:=True;
              break;
              end;
            end;
          if not Trouve then
             begin
               if (Length(getfield('RAC_DESTMAIL')) + Length(NewInterlocuteur) + 1) <= LGDESTMAIL then
                  begin
                  G1RemplirLaLigne(NewInterlocuteur,G1.Row);
                  PutInterlocuteurs;
                  end else
                  begin
                  LastError := 12;
                  PGIBox(TraduireMemoire(TexteMessage[LastError]),'Liste des intervenants');
                  end
             end
          else
             begin
             LastError := 10;
             PGIBox(TraduireMemoire(TexteMessage[LastError]),'Liste des intervenants');
             end;
          end;
       end
    else
       AGLLanceFiche ('YY', 'YYCONTACT','T;'+GetField('RAC_AUXILIAIRE'),intToStr(TOB(G1.Objects[0,G1.Row]).getValue('C_NUMEROCONTACT')) , 'FROMSAISIE');
    end;
end;

procedure TOM_ACTIONS.G1RemplirLaLigne(NewInterlocuteur : string; Row : integer);
var TI : Tob;
    Q : TQuery;
begin
TI:=tob.create('CONTACT',LesInterlocuteurs , -1);
Q:=OpenSQL('select * from CONTACT where C_TYPECONTACT="T" and C_AUXILIAIRE="'+GetField('RAC_AUXILIAIRE')+'" AND C_NUMEROCONTACT='+NewInterlocuteur ,true);
try
  TI.SelectDB('',Q ,True);
  TI.PutLigneGrid(G1 , Row , false ,false ,G1LesColonnes);
finally
  Ferme(Q);
end;
GereNewLigne(G1);
end;

procedure TOM_ACTIONS.PutInterlocuteurs;
var TI : TOB;
    st : string;
    i : integer;
begin
st:='';
For i:=G1.fixedrows to G1.RowCount-1 do
   begin
   TI:=TOB(G1.Objects[0,i]);
   if TI<> nil then
       begin
       st:=st+intToStr(TOB(G1.Objects[0,i]).GetValue('C_NUMEROCONTACT'))+';';
       end ;
   end ;
if not(DS.State in [dsInsert,dsEdit])then
    if ( TF <> Nil ) then
        TF.Edit
    else
        DS.edit;
if st <> '' then st := ';' + st;     // voir modif 0602
SetField('RAC_DESTMAIL',st);
end;

procedure TOM_ACTIONS.G2ElipsisClick(Sender: TObject);
var NewIntervenant : String;
    i : Integer;
    Trouve : Boolean;
    TI : TOB;
begin
if G2.Col = G2ColNom then
    begin
    if G2.Objects[0,G2.Row] = Nil then
       begin
//       if G2.RowCount-1 < NbIntervenants then
{$IFDEF BTP}
			 NewIntervenant:=AFLanceFiche_Rech_Ressource ('ARS_LIBELLE='+G2.cells[G2ColNom,G2.Row],'');
{$ELSE}
       NewIntervenant:= AGLLanceFiche ('AFF', 'RESSOURCERECH_MUL','ARS_LIBELLE='+G2.cells[G2ColNom,G2.Row], '', '');
{$ENDIF}
       if NewIntervenant<>'' then
          begin
          Trouve := False;
          for i := 1 to G2.RowCount-1 do
            begin
              TI:=TOB(G2.Objects[0,i]);
              if (TI <> nil) and (TOB(G2.Objects[0,i]).GetValue('ARS_RESSOURCE') = NewIntervenant) then
              begin
              Trouve:=True;
              break;
              end;
            end;
          if not Trouve then
             begin
             G2RemplirLaLigne(NewIntervenant,G2.Row);
             PutIntervenants;
{               if (Length(getfield('RAC_INTERVINT')) + Length(NewIntervenant) + 1) <= LGINTERVINT then
                  begin
                  G2RemplirLaLigne(NewIntervenant,G2.Row);
                  PutIntervenants;
                  end else
                  begin
                  LastError := 12;
                  PGIBox(TraduireMemoire(TexteMessage[LastError]),'Liste des intervenants');
                  end      }
             end
          else
             begin
             LastError := 10;
             PGIBox(TraduireMemoire(TexteMessage[LastError]),'Liste des intervenants');
             end;
          end;
       end
    else
       V_PGI.DispatchTT(6,taConsult,TOB(G2.Objects[0,G2.Row]).getValue('ARS_RESSOURCE'),'','');
    end;
end;

procedure TOM_ACTIONS.G2RemplirLaLigne(NewIntervenant : string; Row : integer);
var TI : Tob;
begin
TI:=tob.create('RESSOURCE',LesIntervenants , -1);
TI.SelectDB('"'+NewIntervenant+'"',Nil ,True);
TI.PutLigneGrid(G2 , Row , false ,false ,G2LesColonnes);
GereNewLigne(G2);
end;

procedure TOM_ACTIONS.PutIntervenants;
var TI : TOB;
    st : string;
    i : integer;
begin
st:='';
For i:=G2.fixedrows to G2.RowCount-1 do
   begin
   TI:=TOB(G2.Objects[0,i]);
   if TI<> nil then
       begin
       st:=st+TOB(G2.Objects[0,i]).GetValue('ARS_RESSOURCE')+';';
       end ;
   end ;
if not(DS.State in [dsInsert,dsEdit])then
    if ( TF <> Nil ) then
        TF.Edit
    else
        DS.edit;
//SetField('RAC_INTERVINT',st);
//StIntervint := st;
end;

procedure TOM_ACTIONS.G1CellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
if TypeAction<>taConsult then GereNewLigne(G1);
if Not Cancel then
   if (G1.Col=G1ColNom) and (GetField('RAC_TIERS')<>'') then begin G1.ElipsisButton:=true; G1.options:=G1.Options+[goEditing] ;G1.editormode:=True ;end ;
end;

procedure TOM_ACTIONS.G1CellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
if (ACol=G1ColNom) then begin G1.ElipsisButton:=false; G1.options:=G1.Options-[goEditing] ;G1.editormode:=False ;end ;
end;

procedure TOM_ACTIONS.G2CellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
if TypeAction<>taConsult then GereNewLigne(G2);
if Not Cancel then
   if (G2.Col=G2ColNom) and (GetField('RAC_TIERS')<>'') then begin G2.ElipsisButton:=true; G2.options:=G2.Options+[goEditing] ;G2.editormode:=True ;end ;
end;

procedure TOM_ACTIONS.G2CellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
if (ACol=G2ColNom) then begin G2.ElipsisButton:=false; G2.options:=G2.Options-[goEditing] ;G2.editormode:=False ;end ;
end;

procedure TOM_ACTIONS.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
if sender is THGrid then THGrid(Sender).InvalidateRow(ou) ;
end;

procedure TOM_ACTIONS.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var GS : THGrid;
begin
if sender is THGrid then GS:=THGrid(Sender) else exit;
if Not EstRempli(GS,ou) then GS.DeleteRow(ou) ;
GS.InvalidateRow(ou) ;
end;

procedure TOM_ACTIONS.G1DessineCell(ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
begin
DessineCell(G1,Acol,Arow,Canvas,AState);
end;

procedure TOM_ACTIONS.G2DessineCell(ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
begin
DessineCell(G2,Acol,Arow,Canvas,AState);
end;

procedure TOM_ACTIONS.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var GS : THGRID;
    OkG : Boolean ;
    //Q : TQuery ;
    stAuxi : string;
begin
OKG:=( (Screen.ActiveControl=G1) or (Screen.ActiveControl=G2));
if OkG then GS:=THGRID(Screen.ActiveControl) else GS := Nil;
Case Key of
    VK_F3 : {GED} if (ssAlt in Shift) then
                     if (GetParamSocSecur('SO_RTGESTIONGED',False)) and (stProduitpgi = 'GRC') then
                       begin
                          if (ds<>nil) and not(DS.State in [dsInsert]) then
                            begin
                            Key:=0 ;
                            RTAppelGED_OnClick (Sender);
                            end;
                       end;
    VK_F5 : begin
            //Key:=0 ;
            if (Screen.ActiveControl=G1) then G1ElipsisClick(Sender)
            else if (Screen.ActiveControl=G2) then G2ElipsisClick(Sender);
            end;
    VK_F6 : {Infos compl.} if (ssAlt in Shift) then
            if (soRtgestinfos001) and (stProduitpgi='GRC')  then RTAppelParamCL ;
    VK_F7 : if (ssAlt in Shift) then RTAffichePieceLiee(GetField('RAC_NUMCHAINAGE'),stProduitpgi);
    VK_INSERT : if ((OkG) and (Shift=[])) then BEGIN Key:=0 ; ClickInsert(GS, GS.Row) ; END ;
    VK_DELETE : if ((OkG) and (Shift=[ssCtrl])) then BEGIN Key:=0 ; ClickDel(GS, GS.Row) ; END ;
    VK_F10 : {Contacts} if OrigineTiers and  (ssAlt in Shift) and (stTiers<>'') then
                        begin
                        Key:=0 ;
                        stAuxi:=TiersAuxiliaire (stTiers);
                        {Q := OpenSQL('SELECT T_NATUREAUXI,T_PARTICULIER,T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE="'+stAuxi+'"', True);
                        if Q.Eof then begin Ferme(Q); exit; end
                          else begin
                          stNatureauxi:=Q.FindField('T_NATUREAUXI').asstring;
                          stLibelle:=Q.FindField('T_NATUREAUXI').asstring;
                          stParticulier:=Q.FindField('T_NATUREAUXI').asstring;
                          end;
                        Ferme(Q);}
                        AGLLanceFiche('YY','YYCONTACT','T;'+TiersAuxiliaire (stTiers),'', ActionToString(TypeAction)+';TYPE=T;'+'TYPE2='+stNatureauxi+';PART='+stParticulier+';TITRE='+stLibelle);
                        end;
    VK_F12 : {encours} if (OrigineTiers) and (stProduitpgi='GRC') then
                       begin
                            if ((ssAlt in Shift) and (ssShift in Shift)) then
                               begin; Key:=0 ; AfficheRisqueClientDetail(stTiers); end
                            else if (ssAlt in Shift) then
                                 begin Key:=0 ; AfficheRisqueClient(stTiers); end;
                       end;
{$IFNDEF EAGLSERVER}
    81 : {Ctrl + Q - Création d'1 alerte} if (ssCtrl in Shift) then
          begin
          Key:=0 ;
          Alerte_OnClick_RAC(Sender);
          end;
    85 : {Ctrl + U - liste des alertes du tiers} if (ssCtrl in Shift) then
          begin
          Key:=0 ;
          ListAlerte_OnClick_RAC(Sender);
          end;
{$ENDIF EAGLSERVER}
    END;
if (ecran <> nil) then
if ecran is TFFiche then
   TFFiche(ecran).FormKeyDown(Sender,Key,Shift)
else
   TFSaisieList(ecran).FormKeyDown(Sender,Key,Shift);
end;


Procedure TOM_ACTIONS.ClickInsert (GS : THGrid; Row:integer);
begin
GS.InsertRow(Row) ;
if Row > GS.FixedRows then GS.Row:=GS.Row-1;
end;

Procedure TOM_ACTIONS.ClickDel (GS : THGrid; Row:integer);
begin
if GS.Objects[0,GS.Row] <> Nil then begin GS.Objects[0,GS.Row].free; GS.Objects[0,GS.Row]:=Nil; end;
GS.DeleteRow (Row);
if GS=G2 then PutIntervenants;
if GS=G1 then PutInterlocuteurs;
end;


Procedure TOM_ACTIONS.AddSplitter ;
var     PP:THPanel;
begin
Split:=TSplitter.Create(Ecran);
Split.Name:='Splitter1';
PP:=THPanel(GetControl('P1')) ;
Split.Parent:=TWincontrol(getControl('INTERVENANTS'));
Split.AutoSnap:=False;
Split.Beveled:=True;
Split.ResizeStyle:=rsUpdate;
Split.Left:=PP.Left ;
Split.Top:=PP.Top+PP.Height+1 ; // s'assurer que le splitter est sous le controle fixe
Split.Align:=alTop;
Split.Width:=PP.Width;
Split.Height:=3;
Split.Cursor:=crVSplit;
Split.Color:=clActiveCaption;
end;

// ====== Gestion des grids =============================================

procedure DessineCell(GS:THGrid; ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
  Arect: Trect ;
begin
If Arow < GS.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=GS.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := GS.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = GS.row) then
         BEGIN
         Canvas.Brush.Color := clBlack ;
         Canvas.Pen.Color := clBlack ;
         Triangle[1].X:=ARect.Right-2 ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
         Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
         Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
         if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
         END ;
    end;
end;

procedure GereNewLigne(GS:THGrid)  ;
BEGIN
   if (GS.RowCount>GS.FixedCols) and EstRempli(GS,GS.RowCount-1) then NewLigne(GS) else
      if (GS.RowCount>GS.FixedCols+1) and (Not EstRempli(GS,GS.RowCount-2)) then GS.RowCount:=GS.RowCount-1 ;
END ;

procedure NewLigne(GS:THGrid) ;
BEGIN
    GS.RowCount:=GS.RowCount+1 ;
    InitRow( GS,GS.RowCount-1 ) ;
END ;

Function EstRempli(GS:THGrid; Lig : integer) : Boolean ;
var i : integer ;
BEGIN
Result:=False ;
for i:=1 to GS.ColCount-1 do
  if (GS.Cells[i,Lig]<>'') then begin result:= true; break; end;
END ;

Procedure InitRow (GS:THGrid; R : integer) ;
var i : integer ;
begin
for i:=0 to GS.ColCount do GS.cells[i,R]:='';
end ;

{$ENDIF EAGLSERVER}

function AGLPCSFormatDuree( parms: array of variant; nb: integer ) : Variant;
var stDuree : string;
//     Heures,Minutes : Integer;
begin
stDuree:=string(parms[0]);
{Heures:=trunc(strToInt(stDuree)/60);
Minutes:=strToInt(stDuree)-(Heures*60);
result:=EncodeTime(Heures,Minutes,0,0); }
result:=Valeur(stDuree);
end;


{$IFNDEF EAGLSERVER}
procedure AGLRTActionImprime( parms: array of variant; nb: integer ) ;
var stWhere,stArgument : string;
    F : TForm ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if F is TFFiche then
  begin
    stWhere:=' RAC_TIERS="'+string(Parms[1])+'" and RAC_NUMACTION='+string(Parms[2]);
    stArgument := string(Parms[3])+';WHERE='+stWhere;
    AGLLanceFiche ('RT','RTPARAMEDITACTION','','',stArgument)
  end
  else
  begin
    stWhere:=' RAC_TIERS="'+string(Parms[1])+'" and RAC_NUMCHAINAGE='+string(Parms[2]);
    stArgument := string(Parms[3])+';WHERE='+stWhere;
    AGLLanceFiche ('RT','RTPARAMEDITACTION','','',stArgument)
  end ;
end;

procedure TOM_ACTIONS.ActMajIntervint (NewIntervenant:String);
var i : Integer;
    Trouve : Boolean;
    TI : TOB;
begin
Trouve := False;
for i := 1 to G2.RowCount-1 do
  begin
    TI:=TOB(G2.Objects[0,i]);
    if (TI <> nil) and (TOB(G2.Objects[0,i]).GetValue('ARS_RESSOURCE') = NewIntervenant) then
    begin
    Trouve:=True;
    break;
    end;
  end;
if not Trouve then
   begin
   G2RemplirLaLigne(NewIntervenant,G2.Row);
   PutIntervenants;
{     if (Length(getfield('RAC_INTERVINT')) + Length(NewIntervenant) + 1) <= LGINTERVINT then
        begin
        G2RemplirLaLigne(NewIntervenant,G2.Row);
        PutIntervenants;
        end else
        begin
        LastError := 12;
        PGIBox(TraduireMemoire(TexteMessage[LastError]),'Liste des intervenants');
        end  }
   end
end;

Procedure TOM_ACTIONS.ChargeChainage ;
var Q : TQuery;
    LibChainage,stAff : string;
    NumChainage : integer;
    MnRattacher,MnSupprimer,MnListe,chainage,MnFiche :boolean;
begin
  if Assigned(Ecran) and (ecran.name = 'RQACTION_FIC') then
  	exit;
  LibChainage :='';
  NumChainage := getfield ('RAC_NUMCHAINAGE');
  Chainage  := (NumChainage <> 0);
  if (Chainage) then
    if (Ecran is TFFiche) then
        begin
        Q:=OpenSQL('SELECT RCH_LIBELLE,RCH_AFFAIRE FROM ACTIONSCHAINEES WHERE RCH_NUMERO = '+
            IntToStr(NumChainage)+' and RCH_TIERS="'+getfield('RAC_TIERS')+'"', True);
        try
          if not Q.Eof then
            begin
            LibChainage := Q.FindField('RCH_LIBELLE').AsString;
            stAff := Q.FindField('RCH_AFFAIRE').AsString;
            end;
        finally
          Ferme(Q) ;
        end;
        end
    else
      if (copy(ecran.name,1,17) <> 'WINTERVENTION_FSL') then
        LibChainage := TF.TOBFiltre.GetValue('RCH_LIBELLE');

  if (TFFiche(Ecran) <> Nil) then    // Si on vient de la fiche RTCOURRIER il n'y a pas de tform
    SetControlText ('CHAINAGE',IntToStr(NumChainage)+' : '+LibChainage);
  if (Ecran is TFSaisieList) then exit;

  MnRattacher:= (Not Chainage); MnSupprimer:= Chainage; MnListe:= Chainage;
  MnFiche:=Chainage;

  if NoChangeChainage or FicheChainage then
  begin
    MnRattacher:= False; MnListe:= False;
    MnFiche:=False;
  end;
  if NoChangeChainage then
     MnSupprimer:= False;

  if (TypeAction = taConsult) then
  begin
    MnRattacher:=False;MnSupprimer:=False;
  end;
  if (Ecran is TFFiche) then
  begin
    TPopupMenu(GetControl('POPCOMPLEMENT')).items[5].visible:=MnRattacher;
    TPopupMenu(GetControl('POPCOMPLEMENT')).items[6].visible:=MnFiche;
    TPopupMenu(GetControl('POPCOMPLEMENT')).items[7].visible:=MnSupprimer;
    TPopupMenu(GetControl('POPCOMPLEMENT')).items[8].visible:=MnListe;
    TMenuItem(GetControl('MNPIECE')).Visible := MnFiche;
  end;
  SetControlEnabled ('CHAINAGE',Chainage);
  SetControlEnabled ('TCHAINAGE',Chainage);
{$IFDEF AFFAIRE}
  if (Ecran is TFFiche) and ( stAff <> '' ) and ( Getfield('RAC_AFFAIRE') = '' ) then
    begin
    ForceUpdate;
    SetField('RAC_AFFAIRE',stAff);
    THEDIT(GetControl('RAC_AFFAIRE')).text:=stAff;
    ChargeCleAffaire(THEDIT(GetControl('RAC_AFFAIRE0')),THEDIT(GetControl('RAC_AFFAIRE1')), THEDIT(GetControl('RAC_AFFAIRE2')),
       THEDIT(GetControl('RAC_AFFAIRE3')), THEDIT(GetControl('RAC_AVENANT')),Nil , taModif,THEDIT(GetControl('RAC_AFFAIRE')).text,True);

    SetField('RAC_AFFAIRE0',THEDIT(GetControl('RAC_AFFAIRE0')).Text);
    SetField('RAC_AFFAIRE1',THEDIT(GetControl('RAC_AFFAIRE1')).Text);
    SetField('RAC_AFFAIRE2',THEDIT(GetControl('RAC_AFFAIRE2')).Text);
    SetField('RAC_AFFAIRE3',THEDIT(GetControl('RAC_AFFAIRE3')).Text);
    if THEDIT(GetControl('RAC_AVENANT')).Text='' then THEDIT(GetControl('RAC_AVENANT')).Text:='00';
    SetField('RAC_AVENANT',THEDIT(GetControl('RAC_AVENANT')).Text);
    Q:= OpenSQL ('SELECT AFF_LIBELLE,AFF_NUMEROCONTACT FROM AFFAIRE WHERE AFF_AFFAIRE="'+THEDIT(GetControl('RAC_AFFAIRE')).text+'"',True);
    try
       if (Q<>Nil) and (not Q.Eof) then
          begin
          SetControlText ('TRAC_AFFAIRE',Q.FindField('AFF_LIBELLE').asstring);
          if (Q.FindField('AFF_NUMEROCONTACT').asinteger <> 0) and ( Getfield ('RAC_NUMEROCONTACT') = 0 ) then
             begin
             SetField ('RAC_NUMEROCONTACT', Q.FindField('AFF_NUMEROCONTACT').asinteger);
             NomContact();
             end;
          end
       else
          begin
          RACSetFocusControl('BSELECTAFF1');
          LastError := 24;
{$ifdef GIGI}
          LastErrorMsg:=TraduitGa(TexteMessage[LastError]);
{$else}
          LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
{$endif}
          EffaceAffaire(Ecran);
          end;
    finally
      Ferme(Q);
    end;
    end;
{$ENDIF}
{$ifdef GIGI}
 If (Not GetParamSocSecur ('SO_AFRTCHAINAGE',False)) then
   begin
   AddRemoveItemFromPopup ('POPCOMPLEMENT','MNFICHECHAINAGE',False);
   AddRemoveItemFromPopup ('POPCOMPLEMENT','MNLISTECHAINE',False);
   AddRemoveItemFromPopup ('POPCOMPLEMENT','MNRATTACHER',False);
   end;
{$ENDIF}
end;

Procedure TOM_ACTIONS.RattacherChainage ;
var ChaineAct : THEdit;
    Typeaction : THDBValcombobox;
    stWhere,stTiers :string;
BEGIN
ChaineAct:= THEdit.Create (Nil);
Typeaction := THDBValcombobox(GetControl('RAC_TYPEACTION'));
ChaineAct.Parent:=Typeaction.Parent;
ChaineAct.visible := False;
ChaineAct.Top := Typeaction.top;
ChaineAct.Left := Typeaction.left+Typeaction.width;
stTiers :=  Getfield ('RAC_TIERS') ;
stWhere := 'RCH_TIERS="'+stTiers+'"';

ChaineAct.plus := 'RCH_TIERS='+stTiers+';PRODUITPGI='+stProduitpgi;
{$IFNDEF EAGLCLIENT}
ChaineAct.text :=  Getfield ('RAC_NUMCHAINAGE') ;
{$ENDIF}
{$IFNDEF EAGLCLIENT}
if (LookupList(ChaineAct,'Chainages d''actions','ACTIONSCHAINEES','RCH_NUMERO','RCH_LIBELLE',stWhere,'RCH_NUMERO DESC',True,24,'',tlLocate )) then
{$ELSE}
if (LookupList(ChaineAct,'Chainages d''actions','ACTIONSCHAINEES','RCH_NUMERO','RCH_LIBELLE',stWhere,'RCH_NUMERO DESC',True,24,'',tlDefault )) then
{$ENDIF}
begin
  if not(DS.State in [dsInsert,dsEdit])then
    if ( TF <> Nil ) then
        TF.Edit
    else
        DS.edit;
  setfield ('RAC_NUMCHAINAGE',strToInt(ChaineAct.Text));
  ChargeChainage;
end;
ChaineAct.Destroy ;
end;

Procedure TOM_ACTIONS.DetacherChainage ;
begin
  if not(DS.State in [dsInsert,dsEdit])then
    if ( TF <> Nil ) then
        TF.Edit
    else
        DS.edit;
  setfield ('RAC_NUMCHAINAGE',0);
  SetControlText ('CHAINAGE','');
  ChargeChainage;  
end;
{$ENDIF EAGLSERVER}

{$IFDEF AFFAIRE}
Procedure TOM_ACTIONS.AffaireEnabled( grise : boolean);
begin
{$ifdef GIGI}
 If (Not VH_GC.GaSeria) or not (GetParamSocSecur ('SO_AFRTPROPOS',False)) then exit;
{$endif}
  SetControlEnabled ('BEFFACEAFF1',grise); SetControlEnabled ('BSELECTAFF1',grise);
  SetControlEnabled ('RAC_AFFAIRE1',grise);
  SetControlEnabled ('RAC_AFFAIRE2',grise); SetControlEnabled ('RAC_AFFAIRE3',grise);
  SetControlEnabled ('RAC_AVENANT',grise);
end;
{$ENDIF AFFAIRE}

{$IFNDEF EAGLSERVER}
procedure TOM_ACTIONS.OnExit_DateAction(Sender: TObject);
var TobTypActEncours : TOB;
begin
if (DS<>Nil) and (DS.State in [dsInsert,dsEdit]) then
  if GetField ('RAC_DATEACTION') <> Old_DateAction then
     begin
     Old_DateAction := GetField ('RAC_DATEACTION');

     VH_RT.TobTypesAction.Load;

     TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[GetField('RAC_TYPEACTION'),GetField('RAC_CHAINAGE'),GetField('RAC_NUMLIGNE')],TRUE) ;
     if TobTypActEncours = Nil then TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[GetField('RAC_TYPEACTION'),'---',0],TRUE) ;
     if TobTypActEncours <> Nil then
        begin
           if TobTypActEncours.GetValue('RPA_DELAIDATECH') <> 0 then SetField('RAC_DATEECHEANCE',RTCalculEch(GetField('RAC_DATEACTION'),StrToInt(TobTypActEncours.GetValue('RPA_DELAIDATECH')),TobTypActEncours.GetValue('RPA_WEEKEND')));
        end;
     end;
end;
procedure TOM_ACTIONS.MnPiece_OnClick(Sender: TObject);
begin
  RTAffichePieceLiee(GetField('RAC_NUMCHAINAGE'),stProduitpgi);
end;

procedure TOM_ACTIONS.OnChange_HeureAction(Sender: TObject);
begin
  if GetControlText('HEUREACTION') <> Old_HeureAction then
    begin
    if not(DS.State in [dsInsert,dsEdit]) then
      if ( TF <> Nil ) then
          TF.Edit
      else
          DS.Edit;
    end;
end;

procedure TOM_ACTIONS.OnExit_HeureAction(Sender: TObject);
var NouvelleHeure: string;
begin
  NouvelleHeure := FormatDateTime('hh:nn', StrToTime(GetControlText('HEUREACTION')));
  if NouvelleHeure <> Old_HeureAction then
    begin
    Old_HeureAction := NouvelleHeure;
    SetField('RAC_HEUREACTION', StrToTime(GetControlText('HEUREACTION')));
    end;
end;
{$ENDIF EAGLSERVER}

{$IFDEF SAV}
Procedure TOM_ACTIONS.MAJMenuParc ;
var Q : TQuery;
    MnRattacher,MnSupprimer,IdentParc,MnFiche :boolean;
begin
  if (Not VH_GC.SAVSeria) or (stProduitpgi = 'GRF') then exit;
  SetControlText ('IDENTPARC','');
  IdentParc  := (getfield ('RAC_IDENTPARC') <> 0);
  if (IdentParc) then
  begin
    Q:=OpenSQL('SELECT WPN_LIBELLE FROM WPARCNOME WHERE WPN_IDENTIFIANT = '+
        IntToStr(getfield ('RAC_IDENTPARC')), True);
    try
      if not Q.Eof then SetControlText ('IDENTPARC',Q.FindField('WPN_LIBELLE').AsString);
    finally
      Ferme(Q) ;
    end;  
  end;

  MnRattacher:= (Not IdentParc); MnSupprimer:= IdentParc; MnFiche:=IdentParc;

  if (TypeAction = taConsult) then
  begin
    MnRattacher:=False;MnSupprimer:=False;
  end;
  TMenuItem(GetControl('MNPARCAFFECTATION')).Visible := MnRattacher;
  TMenuItem(GetControl('MNPARCCONSULTATION')).Visible := MnFiche;
  TMenuItem(GetControl('MNPARCSUPPRESSION')).Visible := MnSupprimer;
  SetControlEnabled ('IDENTPARC',IdentParc);
  SetControlEnabled ('TIDENTPARC',IdentParc);
end;

Procedure TOM_ACTIONS.AffecterParc (Sender: TObject) ;
var st,numIntervention :string;
    Q : TQuery;
begin
  if GetField('RAC_TIERS') = '' then exit;
  numIntervention:='';
  if (GetField('RAC_NUMCHAINAGE') <> 0) then
    begin
    if (Ecran <> nil) and (copy(ecran.name,1,17) = 'WINTERVENTION_FSL') then
      numIntervention:=';IDENTINTERV='+TF.TOBFiltre.GetString('WIV_IDENTIFIANT')
    else
      begin
      Q:=OpenSQL('SELECT WIV_IDENTIFIANT FROM WINTERVENTION WHERE WIV_NUMCHAINAGE='+IntToStr(GetField('RAC_NUMCHAINAGE')),true);
      if not Q.Eof then
        numIntervention:=';IDENTINTERV='+Q.FindField('WIV_IDENTIFIANT').asString;
      Ferme(Q);
      end;
    end;

  st:= AGLLanceFiche ('W', 'WPARCELEM_SELECT','WPC_TIERS='+GetField('RAC_TIERS')+numIntervention, '', 'ACTION=MODIFICATION');
if st <> '' then
  begin
  if not(DS.State in [dsInsert,dsEdit])then
    if ( TF <> Nil ) then
        TF.Edit
    else
        DS.edit;
  setfield ('RAC_IDENTPARC',Valeuri(st));
  MAJMenuParc;
  end;
end;

Procedure TOM_ACTIONS.SupprimerParc (Sender: TObject);
begin
  if not(DS.State in [dsInsert,dsEdit])then
    if ( TF <> Nil ) then
        TF.Edit
    else
        DS.edit;
  setfield ('RAC_IDENTPARC',0);
  SetControlText ('IDENTPARC','');
  MAJMenuParc;
end;

Procedure TOM_ACTIONS.VoirFicheParc (Sender: TObject);
var Q : TQuery;
    CleWPC : tCleWPC;
begin
  Q := OpenSQL('SELECT WPN_IDENTPARC FROM WPARCNOME WHERE WPN_IDENTIFIANT='+IntToStr(GetField('RAC_IDENTPARC')), True);
  try
    if Not Q.Eof then
      begin
       CleWPC.Identifiant := Q.FindField('WPN_IDENTPARC').AsInteger;
     wCallTreeViewWPN (CleWPC, 'CONSULTATION;WPN_IDENT='+IntToStr(GetField('RAC_IDENTPARC')));
      end
    else
       PGIInfo (TexteMessage [25] );
  finally
    ferme(Q);
  end;  
end;
{$ENDIF}

{$IFNDEF EAGLSERVER}
function Action_MyAfterImport (Sender: TObject) : string;
var  OM : TOM ;
begin
result := '';
if sender is TFFICHE then OM := TFFICHE(Sender).OM else
   if (sender is TFSaisieList) then OM:=TFSaisieList(sender).LeFiltre.OM else exit;
if (OM is TOM_ACTIONS) then result := TOM_ACTIONS(OM).GetInfoGed else exit;
end;

procedure Action_GestionBoutonGED (Sender: TObject);
var  OM : TOM ;
begin
if sender is TFFICHE then OM := TFFICHE(Sender).OM else
   if (sender is TFSaisieList) then OM:=TFSaisieList(sender).LeFiltre.OM else exit;
if (OM is TOM_ACTIONS) then TOM_ACTIONS(OM).GestionBoutonGED else exit;
end;

function TOM_ACTIONS.GetInfoGed : string;
begin
if (not(DS.State in [dsInsert])) and (stProduitpgi = 'GRC') then result := 'Tiers='+GetField('RAC_TIERS')+';'+'Action='+intTostr(GetField('RAC_NUMACTION'))+';'+'Chainage='+intTostr(GetField('RAC_NUMCHAINAGE'));
end;

procedure TOM_ACTIONS.RTAppelGED_OnClick (Sender: TObject) ;
begin
if (Ecran is TFSaisieList) then
   begin
   if (TFSaisieList(Ecran).LeFiltre.TOBFiltre.detail.count <> 0) and (ds<>nil) and not(DS.State in [dsInsert]) then
      begin
      AGLLanceFiche('RT','RTRECHDOCGED','RTD_TIERS='+GetField ('RAC_TIERS')+';RTD_NUMACTION='+IntToStr(GetField ('RAC_NUMACTION')),'','Objet=ACT;Tiers='+GetField ('RAC_TIERS')+';Action='+IntToStr(GetField ('RAC_NUMACTION')));
      GestionBoutonGED;
      end;
   end
else
   begin
   if (ds<>nil) and not(DS.State in [dsInsert]) then
      begin
      AGLLanceFiche('RT','RTRECHDOCGED','RTD_TIERS='+GetField ('RAC_TIERS')+';RTD_NUMACTION='+IntToStr(GetField ('RAC_NUMACTION')),'','Objet=ACT;Tiers='+GetField ('RAC_TIERS')+';Action='+IntToStr(GetField ('RAC_NUMACTION')));
      GestionBoutonGED;
      end;
   end;
end;

Procedure TOM_ACTIONS.GestionBoutonGED;
BEGIN
if (GetParamSocSecur('SO_RTGESTIONGED',False)) and (ds<>nil) and not(DS.State in [dsInsert]) and (ExisteSQL('SELECT RTD_DOCID FROM RTDOCUMENT WHERE RTD_TIERS="'+GetField('RAC_TIERS')+'" AND RTD_NUMACTION='+IntToStr(GetField ('RAC_NUMACTION')))) then
     SetControlVisible ('BDOCGEDEXIST',True)
else SetControlVisible ('BDOCGEDEXIST',False);
END;
function TOM_ACTIONS.GenerChainageIntervention : integer;
var
    FTomChainage : TOM;
    FTobChainage : Tob;
begin
    FTomChainage := CreateTom(RCHTableName, nil, False, True);
    FTobChainage := Tob.Create(RCHTableName, nil, -1);
    FTomChainage.Argument('GENETIERS='+GetField('RAC_TIERS')
        + ';GENELIBELLE=' + TF.TOBFiltre.GetValue('WIV_LIBELLE')
        + ';GENEAFFAIRE=' + GetField('RAC_AFFAIRE')
        + ';GENERESSOURCE=' + GetField('RAC_INTERVENANT')
        );
    if FTomChainage.InitTOB(FTobChainage) then
       if FTomChainage.VerifTOB(FTobChainage) then
          FTobChainage.InsertDB(Nil,True);
    Result:=FTobChainage.GetInteger('RCH_NUMERO');
    FTomChainage.free;
    FTobChainage.free;
end;
{$ENDIF EAGLSERVER}

function TOM_ACTIONS.modifChampYPlanning;
begin
Result := False;
if isFieldModified( 'RAC_DATEACTION') or isFieldModified( 'RAC_HEUREACTION') or
   isFieldModified( 'RAC_DUREEACTION') or isFieldModified( 'RAC_LIBELLE') or
   isFieldModified( 'RAC_ETATACTION') or isFieldModified( 'RAC_TYPEACTION') then Result := True;
end;
// *****************************************************************************
// ********************** gestion Isoflex **************************************
// *****************************************************************************
{$IFNDEF EAGLSERVER}
procedure TOM_ACTIONS.GereIsoflex;
var bIso: Boolean;
  MenuIso: TMenuItem;
begin
  MenuIso := TMenuItem(GetControl('mnSGED'));
  bIso := AglIsoflexPresent;
  if MenuIso <> nil then MenuIso.Visible := bIso;
end;

procedure Rac_AppelIsoFlex(parms: array of variant; nb: integer);
var F: TForm;
  Cle1: string;
begin
  F := TForm(Longint(Parms[0]));
  if (F.Name <> 'RTACTIONS') and (F.Name <> 'RTCHAINAGES') then exit;
  Cle1 := string(Parms[1])+';'+string(Parms[2]);
  AglIsoflexViewDoc(NomHalley, F.Name, 'ACTIONS', 'RAC_CLE1', 'RAC_AUXILIAIRE,RAC_NUMACTION', Cle1, '');
end;
procedure TOM_ACTIONS.AddRemoveItemFromPopup(stPopup,stItem : string; bVisible : boolean);
var pop : TPopupMenu ;
    i : integer;
    st : string;
begin
pop :=TPopupMenu(GetControl(stPopup) ) ;
if pop<> nil  then
   for i:=0 to pop.items.count-1 do
       begin
       st:=uppercase(pop.items[i].name);
       if st=stItem then  begin pop.items[i].visible:=bVisible; break; end;
       end;
end ;
{$ENDIF EAGLSERVER}
{$IFDEF QUALITE}
procedure TOM_ACTIONS.RAC_INTERVENANT_OnElipsisClick(Sender: TObject);
Var
	Ressource:string;
begin
  if not(DS.State in [dsInsert,dsEdit]) then
    DS.edit;
{$IFDEF BTP}
	Ressource:=AFLanceFiche_Rech_Ressource ('ARS_RESSOURCE='+THEdit(GetCOntrol('RAC_INTERVENANT')).text,'');
{$ELSE}
  Ressource:= AGLLanceFiche('AFF','RESSOURCERECH_Mul','ARS_ESTHUMAIN=X;ARS_GENERIQUE=-','','ARS_RESSOURCE=' + GetControlText(TControl(Sender).Name)+';PRODUITPGI=QNC');
{$ENDIF}
	if Ressource <> '' then
 	begin
 		SetControlText(TControl(Sender).Name, Ressource);
  	SetField(TControl(Sender).Name,Ressource);
  end;
end;
{$ENDIF QUALITE}

{$IFNDEF EAGLSERVER}
procedure TOM_ACTIONS.GereIntervenants;
Var
	i: integer;
  TI: Tob;
//  ChaineIntervint: string;
  trouve : boolean;
begin
// on remplace l'ancien Responsable par le nouveau dans la liste des intervenants, on supprime éventuellement
// le nouveau s'il est déjà dans la liste ( pour Info cette liste sert au planning des actions d'un Responsable)
  for i := 1 to G2.RowCount-1 do
  begin
    TI:=TOB(G2.Objects[0,i]);
    if (TI <> nil) and (TOB(G2.Objects[0,i]).GetValue('ARS_RESSOURCE') = getfield('RAC_INTERVENANT')) then
    begin
      G2.Objects[0,i].free;
      G2.Objects[0,i]:=Nil;
      G2.DeleteRow (i);
{      ChaineIntervint:=getfield('RAC_INTERVINT');
      if ( copy(ChaineIntervint,1,length(getfield('RAC_INTERVENANT'))) = getfield('RAC_INTERVENANT') ) then
         Delete(ChaineIntervint, 1, length(getfield('RAC_INTERVENANT'))+1)
      else
         Delete(ChaineIntervint, Pos(';'+getfield('RAC_INTERVENANT')+';',ChaineIntervint), length(getfield('RAC_INTERVENANT'))+1);
      SetField('RAC_INTERVINT',ChaineIntervint);  }
      PutIntervenants;
      break;
    end;
  end;
  trouve:=false;
  for i := 1 to G2.RowCount-1 do
  begin
    TI:=TOB(G2.Objects[0,i]);
    if (TI <> nil) and (TOB(G2.Objects[0,i]).GetValue('ARS_RESSOURCE') = Old_Intervenant) then
    begin
{      if Length(getfield('RAC_INTERVINT')) - Length(Old_Intervenant) + Length(getfield('RAC_INTERVENANT'))
          <= LgIntervint then
      begin
        G2RemplirLaLigne(getfield('RAC_INTERVENANT'),i);
        Setfield('RAC_INTERVINT',StringReplace(getfield('RAC_INTERVINT'), Old_Intervenant, GetField ('RAC_INTERVENANT'), []));
        Old_Intervenant:=GetField ('RAC_INTERVENANT');
      end else
      begin
        LastError := 12;
        PGIBox(TraduireMemoire(TexteMessage[LastError]),'Liste des intervenants');
        Setfield('RAC_INTERVENANT',Old_Intervenant);
      end;  }
      G2RemplirLaLigne(getfield('RAC_INTERVENANT'),i);
      PutIntervenants;
      Old_Intervenant:=GetField ('RAC_INTERVENANT');
      trouve:=true;
      break;
    end;
  end;
  if not trouve then
  begin
{    if (Length(getfield('RAC_INTERVINT')) + Length(GetField ('RAC_INTERVENANT')) + 1) <= LgIntervint then
    begin
      G2RemplirLaLigne(GetField ('RAC_INTERVENANT'),G2.RowCount-1);
      PutIntervenants;
      Old_Intervenant:=GetField ('RAC_INTERVENANT');
    end else
    begin
      LastError := 12;
      PGIBox(TraduireMemoire(TexteMessage[LastError]),'Liste des intervenants');
      Setfield('RAC_INTERVENANT',Old_Intervenant);
    end;   }
    G2RemplirLaLigne(GetField ('RAC_INTERVENANT'),G2.RowCount-1);
    PutIntervenants;
    Old_Intervenant:=GetField ('RAC_INTERVENANT');
  end;
end;
{$ENDIF EAGLSERVER}

{$IFDEF QUALITE}
procedure TOM_ACTIONS.MnProspect_OnClick(Sender: TObject);
var
	Q		: tQuery;
  Arg	: string;
begin
	if getField('RAC_AUXILIAIRE') <> '' then
  	if GetControlText('NATUREAUXI') = 'FOU' then
      AglLanceFiche('GC','GCFOURNISSEUR','',getField('RAC_AUXILIAIRE'),'MONOFICHE;T_NATUREAUXI=FOU')
   	else
 		begin
 			Q:=OpenSQL('SELECT T_NATUREAUXI from TIERS where T_AUXILIAIRE ="'+ getField('RAC_AUXILIAIRE')+'"' , true);
      try
        Arg:='MONOFICHE;ACTION=CONSULTATION;T_NATUREAUXI='+Q.Findfield('T_NATUREAUXI').AsString;
      finally
        Ferme(Q);
      end;  
 			AglLanceFiche('GC','GCTIERS','',getField('RAC_AUXILIAIRE'),Arg);
 		end;
end;
{$ENDIF QUALITE}

{$IFDEF QUALITE}
procedure TOM_ACTIONS.MnFicheInfos_OnClick(Sender: TObject);
begin
	RTAppelParamCL;
end;
{$ENDIF QUALITE}

{$IFDEF QUALITE}
procedure TOM_ACTIONS.RAC_TIERS_OnElipsisClick(Sender: TObject);
Var
  sNatureAuxi: string;
begin
  { Etat du DataSet lié à la fiche }
  if (not (Ecran is TFSaisieList)) and Assigned(DS) and (DS.State = dsBrowse) then
    DS.Edit;
    sNatureAuxi:= '("CLI","FOU")';
	DispatchRecherche(THCritMaskEdit(getControl('RAC_TIERS')), 2, 'T_NATUREAUXI IN '+sNatureAuxi, '', '');
end;
{$ENDIF QUALITE}

{$IFDEF QUALITE}
procedure TOM_ACTIONS.MnContacts_OnClick(Sender: TObject);
begin
	if GetField('RAC_AUXILIAIRE') <> '' then
  	AglLanceFiche('YY','YYCONTACT','T;'+GetField('RAC_AUXILIAIRE'),'', 'ACTION=MODIFICATION');
end;
{$ENDIF QUALITE}

{$IFDEF QUALITE}
procedure TOM_ACTIONS.MnClotureRAC_OnClick(Sender: TObject);
Var
	Rapport: twRapport;
  function GetArgument: string;
  begin
    Result := 'WHERE=(RAC_AUXILIAIRE="' + GetField('RAC_AUXILIAIRE') + '" AND RAC_NUMACTION='+intToStr(GetField('RAC_NUMACTION'))+')'
  end;
begin
	if (DS.State = dsBrowse)
		and (PgiAsk(TraduireMemoire('Confirmez-vous la clôture de cette action?'), Ecran.Caption + ' : ' + intToStr(GetField('RAC_NUMACTION')) ) = MrYes) then
  begin
    Rapport  := TWRapport.Create(TraduireMemoire('Contrôle de la clôture des fiches actions.'));
    try
      if (ControleClotureRAC(GetArgument, Rapport)) then
      begin
        try
    			SetControlsEnabled(['RAC_TYPEACTION','RAC_LIBELLE','RAC_INTERVENANT','RAC_QUALITETYPE','RAC_NIVIMP'], False);
    			SetControlsEnabled(['RAC_QDEMDEROGNUM', 'RAC_TIERS','RAC_AFFAIRE1', 'RAC_DATEACTION','RAC_DATEECHEANCE','RAC_HEUREACTION'], False);
    			SetControlsEnabled(['RAC_AREALISERPAR','DUREEACT','RAC_QPCTAVANCT', 'RAC_TYPELOCA'], False);
          SetGBClotureEnabled(TypeAction<>TaConsult);
          RACSetFocusControl('RAC_REALISEELE');
          if DS.State<>DsEdit then
              DS.Edit;
          CloTureEnCours	:=True;
          ClotureFromFiche:=True;
    //    	wDoAction(wtaClotureRAC, GetArgument, LastErrorMsg);
          SetControlEnabled('RAC_ETATACTION', False);
        finally
          RefreshDb;
        end;
      end
    finally
      Rapport.Display;
      Rapport.free;
    end;
  end;
end;
{$ENDIF QUALITE}

{$IFDEF QUALITE}
procedure TOM_ACTIONS.MnDemDerog_OnClick(Sender: TObject);
begin
	CallFic_DemDerog(GetField('RAC_QDEMDEROGNUM'), GetArgumentString(ActionToString(TypeAction), 'ACTION'));
end;
{$ENDIF QUALITE}

{$IFDEF QUALITE}
procedure TOM_ACTIONS.QDEMDEROGNUM_OnElipsisClick(Sender: TObject);
Var
	DemDerogNum: string;
  function GetRange:string;
  begin
  	result:= 'RQD_QNCNUM='+intToStr(GetField('RAC_QNCNUM'));
  end;
begin
	DemDerogNum:=LanceFiche_RtDemDerog_Tof ('RT', 'RQDEMDEROG_MUL' 		, GetRange, '', ActionToString(TypeAction)+';MULRECHERCHE=RQACTION_FIC');
  if (DemDerogNum<>'') then
  begin
    if not(DS.State in [dsInsert,dsEdit]) then
      DS.edit;
   	SetField('RAC_QDEMDEROGNUM', StrToint(DemDerogNum));
   	SetControlText('RAC_QDEMDEROGNUM', DemDerogNum);
  end;
end;
{$ENDIF QUALITE}

{$IFDEF QUALITE}
procedure TOM_ACTIONS.SetIdentifiant;
begin
	{ DS le 15/02/06 : Attention, le jour ou utomction hérite de WTOM, ne pas faire le SetIdentifiant
  	en automatique dans la WTOM car fait ds le OnNew, cela génère des trous ds les n° de jetons }
  SetField('RAC_IDENTIFIANT', wSetJeton('RAC'));
end;
{$ENDIF QUALITE}

procedure TOM_ACTIONS.SetGBClotureEnabled(Enable: boolean);
begin
  SetControlProperty('RAC_REALISEEPAR' , 'ENABLED', Enable);
  SetControlProperty('RAC_REALISEELE'  , 'ENABLED', Enable);
  SetControlProperty('RAC_VERIFIEELE'  , 'ENABLED', Enable);
  SetControlProperty('RAC_VERIFIEEPAR' , 'ENABLED', Enable);
  SetControlProperty('RAC_EFFJUGEELE'  , 'ENABLED', Enable);
  SetControlProperty('RAC_EFFJUGEEPAR' , 'ENABLED', Enable);
  SetControlProperty('RAC_CLOTUREELE'  , 'ENABLED', Enable);
  SetControlProperty('RAC_CLOTUREEPAR' , 'ENABLED', Enable);
  SetControlProperty('RAC_EFFICACITE'  , 'ENABLED', Enable);
  SetControlProperty('RAC_COUTACTION'  , 'ENABLED', Enable);
end;

procedure TOM_ACTIONS.SetControlsEnabled(Const FieldsName: array of String; Const Enability: Boolean);
var
  i: integer;
begin
  for i := 0 to Length(FieldsName) - 1 do
  	if Enability then
    begin
			{ Rend Enabled = True seulement si le controle n'est pas déja Enabled=false }
    	if GetControlEnabled(FieldsName[i]) then
	      SetControlEnabled(FieldsName[i], Enability);
    end
    else
	  	SetControlEnabled(FieldsName[i], Enability);
end;

{$IFDEF QUALITE}
procedure TOM_ACTIONS.PmAction_OnPopUp(Sender: TObject);
begin
	wSetMnuAction(false, 'ETATLIG=' + GetField('RAC_ETATACTION')+ ';'+ActionToString(TypeAction), tPopUpMenu(GetControl('PMACTION')));
end;
{$ENDIF QUALITE}

{$IFDEF QUALITE}
procedure TOM_ACTIONS.MnJournalAction_OnClick(Sender: TObject);
begin
 	wCallMulWJA('RAC', GetField('RAC_IDENTIFIANT'), ActionToString(typeAction), True);
  if DS.State = DsBrowse then
    RefreshDB;
end;
{$ENDIF QUALITE}


{$IFDEF QUALITE}
procedure TOM_ACTIONS.QNCNUM_OnElipsisClick(Sender: TObject);
Var
	QncNum: string;
  Function GetRange: string;
  begin
    result:= 'RQN_TIERS='+GetField('RAC_TIERS');
  end;
begin
	QncNum:=LanceFiche_RtNonConf_Tof ('RT', 'RQNONCONF_MUL' 		, GetRange, '', ActionToString(TypeAction)+';MULRECHERCHE=RQACTION_FIC');
  if (QncNum<>'') then
  begin
    if not(DS.State in [dsInsert,dsEdit]) then
      DS.edit;
   	SetField('RAC_QNCNUM', StrToInt(QncNum));
   	SetControlText('RAC_QNCNUM', QncNum);
  end;
end;

procedure TOM_ACTIONS.QPLANCORRNUM_OnElipsisClick(Sender: TObject);
Var
	QPlanCorrNum: string;
begin
	QPlanCorrNum:=LanceFiche_RtPlanCorr_Tof ('RT', 'RQPLANCORR_MUL' 		, '', '', ActionToString(TypeAction)+';MULRECHERCHE=RQACTION_FIC');
  if (QPlanCorrNum<>'') then
  begin
    if not(DS.State in [dsInsert,dsEdit]) then
      DS.edit;
   	SetField('RAC_QPLANCORRNUM', StrToInt(QPlanCorrNum));
   	SetControlText('RAC_QPLANCORRNUM', QPlanCorrNum);
  end;
end;

procedure TOM_ACTIONS.MnLpNonConf_OnClick(Sender: TObject);
begin
	CallFic_NonConf(GetField('RAC_QNCNUM'), GetArgumentString(ActionToString(TypeAction), 'ACTION'));
end;

procedure TOM_ACTIONS.MnLpPlanCorr_OnClick(Sender: TObject);
begin
	CallFic_PlanCorr(GetField('RAC_QPLANCORRNUM'), GetArgumentString(ActionToString(TypeAction), 'ACTION'));
end;

procedure TOM_ACTIONS.MnProperties_OnClick(Sender: TObject);
var
  Identifiant: String;
begin
  Identifiant := intToStr(GetField('RAC_IDENTIFIANT'));
  if Identifiant <> '' then
    wCallProperties('RAC', Identifiant, 'RAC_IDENTIFIANT', IntToStr(LongInt(Ecran)))
  else
	  PGIInfo(TraduireMemoire('L''enregistrement est vide.'), TraduireMemoire('Proprietés'));
end;
{$ENDIF QUALITE}

{$IFDEF QUALITE}
procedure TOM_ACTIONS.RAC_LigneOrdre_OnElipsisClick(Sender: TObject);
Var
	CleOrdre, Cle1: string;

	function GetRange: string;
  begin
  	result:= 'WOL_NATURETRAVAIL=' + GetControlText('RAC_NATURETRAVAIL')
  end;
begin
  CleOrdre:= LanceFiche_wOrdreLigMul_Tof('W', 'WORDRELIGREC_MUL', GetRange, '', 'LANCEPAR=UTOMACTION;MULRECHERCHE'+iif(GetControlText('RAC_NATURETRAVAIL')<>'', ';WOL_NATURETRAVAIL='+GetControlText('RAC_NATURETRAVAIL'),''));
  if (CleOrdre<>'') then
  begin
  	Cle1:= ReadTokenSt(CleOrdre);
  	if GetControlText('RAC_NATURETRAVAIL') <> Cle1 then
		begin
  		SetControlText('RAC_NATURETRAVAIL', Cle1);
  		SetField('RAC_NATURETRAVAIL', Cle1);
    end;
   	SetControlText('RAC_LIGNEORDRE', CleOrdre);
   	SetField('RAC_LIGNEORDRE', CleOrdre);
  end;
end;
{$ENDIF QUALITE}

{$IFDEF QUALITE}
procedure TOM_ACTIONS.SetNumRQNEnabled;
begin
  if GetField('RAC_QPLANCORRNUM') = null then
    SetField('RAC_QPLANCORRNUM', 0);
  if GetField('RAC_QNCNUM') = null then
    SetField('RAC_QNCNUM', 0);
  if (FamAction<>'DER') and (GetField('RAC_QDEMDEROGNUM')<>0) then
    SetField('RAC_QDEMDEROGNUM', 0);
  if (FamAction='DER') and (GetField('RAC_QPLANCORRNUM')<>0) then
    SetField('RAC_QPLANCORRNUM', 0);

  // Accessible si on n'est pas en création d'une action depuis une non conformité
  SetControlProperty('RAC_QNCNUM'       , 'ENABLED', (sQNCOrigine<>'RQN') and (TFFiche(Ecran).fTypeAction = taCreat)
                                                and  (GetField('RAC_QPLANCORRNUM')=0) );

  SetControlProperty('RAC_QPLANCORRNUM' , 'ENABLED', (sQNCOrigine<>'RQP')
                                                 and (TFFiche(Ecran).fTypeAction = taCreat)
                                                 and (FamAction <> 'DER')
//                                                 and (GetControlText('RAC_QUALITETYPE')<>'001')  // action immédiate
                                                 and (GetField('RAC_QNCNUM')=0) );  // Pas de non conformité et plan correcteur en même temps

  // Accessible si on n'est pas en création d'une action depuis une demande de dérogation
  SetControlProperty('RAC_QDEMDEROGNUM', 'ENABLED', ((sQNCOrigine<>'RQP') and (TFFiche(Ecran).fTypeAction = taCreat) and (FamAction = 'DER'))
                                                or ( (sQNCOrigine<>'RQP') and (TFFiche(Ecran).fTypeAction <> taCreat) and (GetField('RAC_QDEMDEROGNUM')=0)  ) );
end;
{$ENDIF QUALITE}

{$IFDEF QUALITE}
procedure TOM_ACTIONS.QUALITETYPE_OnChange(Sender: TObject);
begin
  SetNumRQNEnabled;
end;
{$ENDIF QUALITE}
{$IFNDEF EAGLSERVER}
procedure TOM_ACTIONS.SetArguments(StSQL : string);
var Critere,ChampMul,ValMul : string ;
    x,y : integer ;
    Ctrl : TControl;
    Fiche : TFFiche;
begin
SetControlVisible('BSTOP',TRUE);
SetControlVisible('BDUPLICATION',False);
SetControlVisible('BCOURRIER',False);
if StSQL <> '' then DS.Edit;
Fiche := TFFiche(ecran);
Repeat
    Critere:=Trim(ReadTokenPipe(StSQL,'|')) ;
    if Critere<>'' then
        begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           y := pos(',',ValMul);
           if y<>0 then ValMul:=copy(ValMul,1,length(ValMul)-1);
           if copy(ValMul,1,1)='"' then ValMul:=copy(ValMul,2,length(ValMul));
           if copy(ValMul,length(ValMul),1)='"' then ValMul:=copy(ValMul,1,length(ValMul)-1);
           SetField(ChampMul,ValMul);
           Ctrl:=TControl(Fiche.FindComponent(ChampMul));
           if Ctrl=nil then exit;
{$IFDEF EAGLCLIENT}
           if (Ctrl is TCustomCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is TCustomEdit) then
           begin
            TEdit(Ctrl).Font.Color:=clRed;
            SetControlText(ChampMul,ValMul);
           end
           else if Ctrl is TSpinEdit then TSpinEdit(Ctrl).Font.Color:=clRed
           else if (Ctrl is TCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is THEdit)Or (Ctrl is THNumEdit)then
              begin
              TSpinEdit(Ctrl).Font.Color:=clRed;
              SetControlText(ChampMul,ValMul);
              end;
{$ELSE}
           if (Ctrl is TDBCheckBox) or (Ctrl is THDBValComboBox) Or (Ctrl is THDBEdit) then TEdit(Ctrl).Font.Color:=clRed
           else if Ctrl is THDBSpinEdit then THDBSpinEdit(Ctrl).Font.Color:=clRed
           else if (Ctrl is TCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is THEdit)Or (Ctrl is THNumEdit)then
              begin
              THDBSpinEdit(Ctrl).Font.Color:=clRed;
              SetControlText(ChampMul,ValMul);
              end;
{$ENDIF}
           end;
        end;
until  Critere='';
end;
{$ENDIF EAGLSERVER}

procedure TOM_ACTIONS.MAJYPlanning (TobYPlanning : Tob;Ressource,Guid:string);
var TYPL : Tob;
    DureeAct : double;
    Heures,Minutes : integer;
    StGuidYRS : string;
begin
StGuidYRS := GetYRS_GUID('', Ressource, '');
if (StGuidYRS <> '') then
  begin
  TYPL:=Tob.create('#YPLANNING',TobYPlanning,-1);
  TYPL.AddChampSupValeur('YPL_PREFIXE', 'RAI');
  TYPL.AddChampSupValeur('YPL_GUIDYRS', StGuidYRS);
  TYPL.AddChampSupValeur('YPL_GUIDORI', Guid);
  TYPL.AddChampSupValeur('YPL_DATEDEBUT', GetField('RAC_DATEACTION') + GetField('RAC_HEUREACTION'));
  Dureeact := GetField('RAC_DUREEACTION');
  Heures:=trunc(Dureeact/60);
  Minutes:=trunc(Dureeact-(Heures*60));
  TYPL.AddChampSupValeur('YPL_DATEFIN', GetField('RAC_DATEACTION') + GetField('RAC_HEUREACTION') + EncodeTime(Heures,Minutes,0,0));
  TYPL.AddChampSupValeur('YPL_LIBELLE', GetField('RAC_LIBELLE'));
  TYPL.AddChampSupValeur('YPL_ABREGE', GetField('RAC_LIBELLE'));
  TYPL.AddChampSupValeur('YPL_STATUTYPL', GetField('RAC_ETATACTION'));
  TYPL.AddChampSupValeur('YPL_PRIVE', '-');
  end;
end;

function TOM_ACTIONS.ControleEstLibre : Boolean;
var TobControleYplanning,{$IFNDEF EAGLSERVER}TF,TI,TobTemp,{$ENDIF EAGLSERVER}TobRessourcesOrigine : Tob;
    DureeAct : double;
    Heures,Minutes : integer;
{$IFNDEF EAGLSERVER}
    i : integer;
{$ENDIF EAGLSERVER}
    Q : TQuery;
    DateDebut,DateFin : TDateTime;
begin
Result := True;
TobControleYplanning := Tob.create('CONTROLEYPLANNING',Nil,-1) ;
if isFieldModified( 'RAC_DATEACTION') or isFieldModified( 'RAC_HEUREACTION') or
   isFieldModified( 'RAC_DUREEACTION') or isFieldModified( 'RAC_TYPEACTION') then
  begin
{$IFNDEF EAGLSERVER}
  for i := 1 to G2.RowCount-1 do
    begin
    TI:=TOB(G2.Objects[0,i]);
    if TI<> nil then
      begin
      TF:=Tob.create('#CONTROLE',TobControleYplanning,-1);
      TF.AddChampSupValeur('RESSOURCE',TOB(G2.Objects[0,i]).GetValue('ARS_RESSOURCE') );
      TF.AddChampSupValeur('NOMRESSOURCE',TOB(G2.Objects[0,i]).GetValue('ARS_LIBELLE') );
      TF.AddChampSupValeur('GUID','');
      Q:=OpenSql('SELECT RAI_GUID FROM ACTIONINTERVENANT WHERE RAI_AUXILIAIRE = "'+GetField('RAC_AUXILIAIRE')+'" AND RAI_NUMACTION ='+IntToStr(GetField('RAC_NUMACTION'))
         + ' AND RAI_RESSOURCE = "'+TOB(G2.Objects[0,i]).GetValue('ARS_RESSOURCE') + '"',True);
      if Not Q.Eof then TF.SetString('GUID',Q.FindField('RAI_GUID').AsString);
      Ferme(Q);
      TF.AddChampSupValeur('LIBRE','X');
      TF.AddChampSupValeur('MOTIF','');
      end;
    end;
{$ENDIF EAGLSERVER}
  end
else
  begin
  TobRessourcesOrigine:=Tob.create('les LIENSACTIONS',Nil,-1) ;
  Q:=OpenSql('SELECT * FROM ACTIONINTERVENANT WHERE RAI_AUXILIAIRE = "'+GetField('RAC_AUXILIAIRE')+'" AND RAI_NUMACTION ='+IntToStr(GetField('RAC_NUMACTION')),true);
  try
   TobRessourcesOrigine.loaddetailDB('ACTIONINTERVENANT','','',Q ,false,False);
  finally
   ferme(Q);
  end;
{$IFNDEF EAGLSERVER}
  for i := 1 to G2.RowCount-1 do
    begin
    TI:=TOB(G2.Objects[0,i]);
    if TI<> nil then
      begin
      TobTemp := TobRessourcesOrigine.FindFirst (['RAI_AUXILIAIRE','RAI_NUMACTION','RAI_RESSOURCE'],[GetField('RAC_AUXILIAIRE'),GetField('RAC_NUMACTION'),TOB(G2.Objects[0,i]).GetValue('ARS_RESSOURCE')],TRUE);
      if TobTemp = Nil then
        begin
        TF:=Tob.create('#CONTROLE',TobControleYplanning,-1);
        TF.AddChampSupValeur('RESSOURCE',TOB(G2.Objects[0,i]).GetValue('ARS_RESSOURCE') );
        TF.AddChampSupValeur('NOMRESSOURCE',TOB(G2.Objects[0,i]).GetValue('ARS_LIBELLE') );
        TF.AddChampSupValeur('GUID','');
        TF.AddChampSupValeur('LIBRE','X');
        TF.AddChampSupValeur('MOTIF','');
        end;
      end;
    end;
{$ENDIF EAGLSERVER}
  FreeAndNil (TobRessourcesOrigine);
  end;
if TobControleYplanning.Detail.Count <> 0 then
  begin
  DateDebut := GetField ('RAC_DATEACTION') + GetField('RAC_HEUREACTION');
  Dureeact := GetField('RAC_DUREEACTION');
  Heures:=trunc(Dureeact/60);
  Minutes:=trunc(Dureeact-(Heures*60));
  DateFin := GetField('RAC_DATEACTION') + GetField('RAC_HEUREACTION') + EncodeTime(Heures,Minutes,0,0);
  Result := ControleIsFreeYPL (TobControleYplanning,DateDebut,DateFin,True);
  end;
FreeAndNil (TobControleYplanning);
end;

{$IFNDEF EAGLSERVER}
function Tom_actions.TestAlerteAction (CodeA : String) : boolean;
var TOBInfosCompl : tob;
    i : integer;
    F: TForm;
    Q : TQuery;
begin
  result:=true;
  { cas de la duplication ou l'on passe dans le loadalerte alors que les champs ne sont pas renseignés }
  if (GetField('RAC_AUXILIAIRE') = '' ) or ( GetField('RAC_NUMACTION') = 0 ) then exit;

  F:=TForm(Ecran);
  if (not (F is TfFiche) ) and (not (F is TFSaisieList) ) then exit;

  if (F is TfFiche) then
    begin
    if  assigned( TfFiche(F).TobFinale) then
      TfFiche(F).TobFinale.free;
    TfFiche(F).TobFinale:=TOB.create ('ACTIONS',NIL,-1);
    TfFiche(F).TobFinale.GetEcran (TFfiche(Ecran),Nil);
    end;

  if (F is TFSaisieList) then
    begin
    if TF.Recno = 0 then exit;
    if assigned( TFSaisieList(F).TobFinale) then
      TFSaisieList(F).TobFinale.free;
    TFSaisieList(F).TobFinale:=TOB.create ('ACTIONS',NIL,-1);
    TFSaisieList(F).TobFinale.Dupliquer(TF.TobFiltre.Detail[TF.Recno-1], False, True);
    end;


  { si passage du load, on sauvegarde les tobs initiales }
  if pos(CodeOuverture,CodeA) > 0 then
    begin
    if (F is TfFiche) then
      begin
      if assigned( TfFiche(F).TobOrigine) then TfFiche(F).TobOrigine.free;
      TfFiche(F).TobOrigine:=TOB.create ('ACTIONS',NIL,-1);
      TfFiche(F).TobOrigine.GetEcran (F,Nil,true);
      end
    else
      begin
      if assigned( TFSaisieList(F).TobOrigine) then TFSaisieList(F).TobOrigine.free;
      TFSaisieList(F).TobOrigine:=TOB.create ('ACTIONS',NIL,-1);
      TFSaisieList(F).TobOrigine.Dupliquer(TF.TobFiltre.Detail[TF.Recno-1], False, True);
      end;
   end;
  { si (F is TFSaisieList) infos compl déjà dans la tob }
  if ( GetParamSocSecur('SO_RTGESTINFOS001',false) ) and (not (F is TFSaisieList)) then
    begin
    TOBInfosCompl:= TOB.Create('RTINFOS001', nil, -1);
    if CodeA<>CodeCreation then
      begin
      Q:=OpenSQL('Select * from RTINFOS001 where RD1_CLEDATA="'+GetField('RAC_AUXILIAIRE')+';'+IntToStr(GetField('RAC_NUMACTION'))+'"',True);
      TOBInfosCompl.selectDB ('',Q,False);
      ferme(Q);
      end;
    for i := 1 to Pred(TOBInfosCompl.NbChamps) do
      if (F is TfFiche) then
        TfFiche(F).TobFinale.AddChampSupValeur(TOBInfosCompl.GetNomChamp(i), TOBInfosCompl.GetValue(TOBInfosCompl.GetNomChamp(i)))
      else
        TFSaisieList(F).TobFinale.AddChampSupValeur(TOBInfosCompl.GetNomChamp(i), TOBInfosCompl.GetValue(TOBInfosCompl.GetNomChamp(i)));

    if (pos(CodeOuverture,CodeA) > 0) then
      begin
      if (pos(CodeOuverture,CodeA) > 0) and assigned( TfFiche(F).TobOrigine) then
        for i := 1 to Pred(TOBInfosCompl.NbChamps) do
          if (F is TfFiche) then
            TfFiche(F).TobOrigine.AddChampSupValeur(TOBInfosCompl.GetNomChamp(i), TOBInfosCompl.GetValue(TOBInfosCompl.GetNomChamp(i)))
          else
            TFSaisieList(F).TobOrigine.AddChampSupValeur(TOBInfosCompl.GetNomChamp(i), TOBInfosCompl.GetValue(TOBInfosCompl.GetNomChamp(i)));
      end;
    TOBInfosCompl.free;
    end;

  if pos(CodeOuverture,CodeA) > 0 then
    result:=ExecuteAlerteLoad(F,false)
  else
    if pos(CodeSuppression,CodeA) > 0 then
      result:=ExecuteAlerteDelete(F,false)
    else
      result:=ExecuteAlerteUpdate(F,false);
end;

{ GC/GRC : MNG / gestion des alertes }
procedure Tom_actions.ListAlerte_OnClick_RAC(Sender: TObject);
begin
if (GetField('RAC_NUMACTION') <> 0) and(AlerteActive(TableToPrefixe(TableName))) then
   AGLLanceFiche('Y','YALERTES_MUL','YAL_PREFIXE=RAC','','ACTION=CREATION;MONOFICHE;CHAMP=RAC_AUXILIAIRE|RAC_NUMACTION;VALEUR='
      +GetField('RAC_AUXILIAIRE')+'|'+IntToStr(GetField('RAC_NUMACTION'))+';LIBELLE='+GetField('RAC_LIBELLE')) ;
end ;

procedure Tom_actions.Alerte_OnClick_RAC(Sender: TObject);
begin
  if (GetField('RAC_NUMACTION') <> 0) and(AlerteActive(TableToPrefixe(TableName))) then
    AGLLanceFiche('Y','YALERTES','','','ACTION=CREATION;MONOFICHE;CHAMP=RAC_AUXILIAIRE|RAC_NUMACTION;VALEUR='
    +GetField('RAC_AUXILIAIRE')+'|'+IntToStr(GetField('RAC_NUMACTION'))+';LIBELLE='+GetField('RAC_LIBELLE')) ;
  VH_EntPgi.TobAlertes.ClearDetail;
end;
{$ENDIF !EAGLSERVER}

procedure Tom_actions.RACSetFocusControl(Champ : String);
begin
  if Assigned (Ecran) then
    SetFocusControl(Champ);
end;


procedure Tom_actions.RAC_INTERVENANTCLICK (Sender : TObject);
var intervenant : string;
begin
Intervenant:=AFLanceFiche_Rech_Ressource ('ARS_RESSOURCE='+THEdit(GetCOntrol('RAC_INTERVENANT')).text,'');
if Intervenant <> '' then
  begin
  THEdit(GetCOntrol('RAC_INTERVENANT')).text := Intervenant;
	if not (DS.State in [dsInsert,dsEdit]) then DS.Edit;
  Setfield ('RAC_INTERVENANT',Intervenant);
  end;

end;

Procedure Tom_actions.ControleChamp(Champ : String;Valeur : String);
Var X : integer;
Begin

  if Champ='RAC_PROJET'     then stProjet     := Valeur;
  if Champ='RAC_OPERATION'  then stOperation  := Valeur;
  if Champ='RAC_TIERS'      then stTiers      := Valeur;

  if Champ='ACTION'    then
  begin
    if Valeur = 'CREATION'          then TypeAction := taCreat
    else if Valeur = 'MODIFICATION' Then TypeAction := taModif
    else if Valeur = 'CONSULTATION' Then TypeAction := taConsult;
  end;

  if Champ='DUPLICATION'    then
  begin
    x := pos('|',Valeur);
    stTiersDuplic := copy(Valeur,1,x-1);
    iNumActDuplic := strToInt(copy(Valeur,x+1,length(Valeur)));
  end;

  if Champ='RAC_PERSPECTIVE' then
  begin
    iPerspective := strToInt(Valeur);
    SetControlEnabled('NUM_PERSPECTIVE',FALSE) ;
  end;

  if Champ='RAC_NUMEROCONTACT'  then iNumeroContact := strToInt(Valeur);

  if Champ='RAC_INTERVENANT'    then stIntervenant := Valeur;
  if Champ='RAC_NUMCHAINAGE'    then iNumChainage := StrToInt(Valeur);
  if Champ='RAC_AFFAIRE'        then stAffaire := Valeur;
  if Champ='PRODUITPGI'         then stProduitpgi := Valeur;
  if Champ='RAC_IDENTPARC'      then iIdentParc := Valeuri(Valeur);
  {$IFDEF QUALITE}
  if Champ='RAC_QNCNUM'         then iQncNum := Valeuri(Valeur);
  if Champ='RAC_QPLANCORRNUM'   then iPlanCorrNum := Valeuri(Valeur);
  if Champ='RQN_QLITIGIEUSE'    then iQLitigieuse := Valeuri(Valeur);
  if Champ='ACTION'             then sAction := Valeur;
  if Champ='QNCORIGINE'         then sQNCOrigine := Valeur;
  if Champ='CLOTUREENCOURS'     then ClotureEncours := StrToBool_(Valeur);
  {$ENDIF QUALITE}

  if Champ='ORIGINE'            then sOrigine := Valeur;

end;

Procedure Tom_actions.ControleCritere(Critere : String);
begin

  if Critere = 'NOCHANGEPROSPECT' then NoChangeProspect := True ;
  if Critere = 'NOCHANGEPROJET'   then NoChangeProjet   := True ;
  if Critere = 'NOCHANGEAFFAIRE'  then NoChangeAffaire  := True;
  if Critere = 'NOCHANGECHAINAGE' then NoChangeChainage := true; //Liste Chainage
  if Critere = 'NOCHANGEAFFAIRE'  then NoChangeAffaire  := True;
  if Critere = 'FICHECHAINAGE'    then FicheChainage    := true; //Liste Chainage
  if Critere = 'MODIFPLANNING'    then Planning         := true;
  if Critere = 'NOCREAT'          then NoCreat          := True ;  // Analyse dynamique des actions

  if Critere = 'FICHETIERS'       then
  begin
{$IFNDEF EAGLSERVER}
    if ecran is TFFiche then
    begin
      RTMajPopMenu(TFFiche(ecran),'POPCOMPLEMENT',0,false) ;
      TFFiche(Ecran).MonoFiche := False;
    end;
    {$ENDIF EAGLSERVER}
    OrigineTiers := true;
    VerrouModif := (TypeAction = taConsult);
  end;

  if Critere = 'ORIGINEPERSP'      then
  begin
    ActionOblig := True ;
    SetControlenabled ('BFERME',False) ;
  end;

  if Critere = 'FICHECONTACT'      then
  begin
  {$IFNDEF EAGLSERVER}
    if ecran is TFFiche then RTMajPopMenu(TFFiche(ecran),'POPCOMPLEMENT',0,false) ;
  {$ENDIF EAGLSERVER}
    OrigineTiers := true;
    VerrouModif := (TypeAction = taConsult);
    SetControlVisible ('BCONTACT',False);
    NoChangeContact := True ;
  end;

  if Critere = 'CREATPLANNING' then  //Planning
  begin
    Planning := true;
    {$IFNDEF EAGLSERVER}
    TobAct := TheTob;
    {$ENDIF EAGLSERVER}
  end;

end;


Initialization
registerclasses([TOM_ACTIONS]) ;
{$IFNDEF EAGLSERVER}
RegisterAglProc( 'HeureDebut', TRUE , 0, AGLHeureDebut);
RegisterAglProc( 'HeureFin', TRUE , 0, AGLHeureFin);
RegisterAglProc( 'ActEnvoiMessage', TRUE , 2, AGLActEnvoiMessage);
RegisterAglFunc( 'ActFormatHeure', TRUE , 0, AGLActFormatHeure);
RegisterAglFunc( 'ActFormatDuree', TRUE , 0, AGLActFormatDuree);
RegisterAglFunc( 'PCSFormatDuree', False , 1, AGLPCSFormatDuree);
//RegisterAglFunc( 'PCSDureeCombo', False , 1, AGLPCSDureeCombo);
RegisterAglProc( 'NomContact', TRUE , 0, AGLNomContact);
RegisterAglProc( 'DuplicationAct', TRUE , 0, AGLDuplicationAct);
RegisterAglProc( 'RTAjouteOutlook', TRUE , 1, AGLRTAjouteOutlook);
RegisterAglProc( 'ChgtTypeAction', TRUE , 0, AGLChgtTypeAction);
RegisterAglProc( 'RTAppelParamCL', True,0,AGLRTAppelParamCL) ;
RegisterAglProc( 'RTActionImprime', TRUE , 3, AGLRTActionImprime);
RegisterAglProc( 'ActMajIntervint', TRUE , 1, AGLActMajIntervint);
RegisterAglProc( 'RattacherChainage', TRUE , 0, AGLRattacherChainage);//chainage
RegisterAglProc( 'DetacherChainage', TRUE , 0, AGLDetacherChainage);//chainage
RegisterAglProc('Rac_AppelIsoFlex', TRUE, 2, Rac_AppelIsoFlex);
{$ENDIF EAGLSERVER}
end.

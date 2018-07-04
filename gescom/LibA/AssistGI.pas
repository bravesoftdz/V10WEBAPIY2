unit AssistGI;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
  ActnList, Mask, HPanel, UTOB,
  ImgList, Menus,
  HEnt1,LicUtil,MajTable,ParamSoc, HStatus, AGLInit,
  // Uses Commun
  FichComm,
  // Uses Compta
  Ent1,

{$IFDEF EAGLCLIENT}
  MaineAGL,etablette,
// Pour l'instant, pas de gestion multi-dossiers en eagl...
{$ELSE}
 AssistPL, Ubob,  PGIExec,  fe_main,tablette,
 dbtables, galoutil,
{$ENDIF}
{$IFNDEF SANSCOMPTA}
  //Devise,ContAbon,
  Exercice, AssistExo, MulGene, MulJal, Guide, ParamLib,
{$ENDIF}
  // Uses GI, GA
  UtilSocAF, AffaireUtil, DicoAF,AssistCodeAffaire,
  //InterSavSisco,
  // PGIEnv
  PGIEnv, Buttons ,UtofAfAssCreatLibRess ,UtofAfAssCreatLibtier
  ;

function LanceAssistantGI_PCL (bInit : boolean) : boolean;
//procedure InitEtablissement;
procedure InitArticle;
procedure InitUnite;
procedure InitProfilGener;
procedure InitCompta;
procedure InitTVA;
procedure InitParPiece;
procedure InitBlocageAffaire;
procedure InitZonesLibres;

function OpenSQLCom(St : string; RO : boolean) : TQuery;
procedure FermeSQLCom (var QCom : TQuery);
Procedure FicheEtablissement_GC ( Mode : TActionFiche ) ;

type
  TFAssistGI = class(TFAssist)
    ActionList1: TActionList;
    PDebut: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    HLabel1: THLabel;
    TSO_LIBELLE: THLabel;
    HLabel5: THLabel;                 
    HLabel6: THLabel;
    bCoordonnees: TToolbarButton97;
    SO_LIBELLE: TEdit;
    SO_CODE: TEdit;
    GroupBox2: TGroupBox;
    ToolbarButton971: TToolbarButton97;
    HLabel8: THLabel;
    GroupBox4: TGroupBox;
    DossierEuro: TCheckBox;
    HLabel2: THLabel;
    ImageList: TImageList;
    HLabel3: THLabel;
    PopParamPlan: TPopupMenu;
    PlanComptesGeneraux: TMenuItem;
    JournauxComptables: TMenuItem;
    LibelleAuto: TMenuItem;
    GuideSaisie: TMenuItem;
    OpenDialog: TOpenDialog;
    HLabel7: THLabel;
    MODE_CREATION: THValComboBox;
    LABEL_CREATION: THLabel;
    FICHIER_IMPORT: THCritMaskEdit;
    SO_NUMPLANREF: THValComboBox;
    bPrefMission: TToolbarButton97;
    bTypMission: TToolbarButton97;
    SO_AFLIENDP: TCheckBox;
    TabSheet5: TTabSheet;
    HLabel4: THLabel;
    GroupBox3: TGroupBox;
    GroupBox5: TGroupBox;
    bActivite: TToolbarButton97;
    GroupBox6: TGroupBox;
    bMonnaie: TToolbarButton97;
    bFourchettes: TToolbarButton97;
    bParamComptable: TToolbarButton97;
    bRessource: TToolbarButton97;
    HLabel9: THLabel;
    GroupBox7: TGroupBox;
    CheckBox2: TCheckBox;
    SBCollab: TSpeedButton;
    HLabel10: THLabel;
    Autre1: THCritMaskEdit;
    SBAutre1: TSpeedButton;
    bComplementsRes: TToolbarButton97;
    Autre2: THCritMaskEdit;
    SBAutre2: TSpeedButton;
    bComplementsAff: TToolbarButton97;
    SO_AFGESTIONCOM: TCheckBox;
    CBFormatExercice: THValComboBox;
    LblFormatExercice: THLabel;
    GBOrganisation: TGroupBox;
    CBAgence: TCheckBox;
    SBAgence: TSpeedButton;
    SBBureau: TSpeedButton;
    CBBureau: TCheckBox;
    CBAssocie: TCheckBox;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    CBChefGroupe: TCheckBox;
    CBResponsable: TCheckBox;
    bComplementsCli: TToolbarButton97;
    CBModeRegle: THValComboBox;
    SBModeRegle: TSpeedButton;
    HLabel11: THLabel;
    Autre0: THCritMaskEdit;
    PopParam: TPopupMenu;
    Activer: TMenuItem;
    Desactiver: TMenuItem;
    bEntiteJur: TToolbarButton97;
    bDates: TToolbarButton97;
    BAffaire: TToolbarButton97;
    LCSO_AFFACTPARRES: THLabel;
    CSO_AFFACTPARRES: THValComboBox;
    BCOMPLEMENTFACT: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure bAnnulerClick(Sender: TObject);
    procedure OnBoutonActivationClick(Sender: TObject);
    procedure ActiverClick(Sender: TObject);
    procedure bCoordonneesClick(Sender: TObject);
    procedure CSO_AFFACTPARRESChange(Sender: TObject);
    procedure PlanComptesGenerauxClick(Sender: TObject);
    procedure JournauxComptablesClick(Sender: TObject);
    procedure LibelleAutoClick(Sender: TObject);
    procedure GuideSaisieClick(Sender: TObject);
    procedure bParamComptableClick(Sender: TObject);
    procedure bFourchettesClick(Sender: TObject);
    procedure bPreferencesClick(Sender: TObject);
    procedure bMonnaieClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure DossierEuroClick(Sender: TObject);
    procedure MODE_CREATIONChange(Sender: TObject);
    procedure FICHIER_IMPORTClick(Sender: TObject);
    procedure bPrefMissionClick(Sender: TObject);
    procedure bRessourceClick(Sender: TObject);
    procedure bActiviteClick(Sender: TObject);
    procedure bComplementFactClick(Sender: TObject);
    procedure bTypMissionClick(Sender: TObject);
    procedure bAffaireClick(Sender: TObject);
    procedure SO_AFLIENDPClick(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure SBCollabClick(Sender: TObject);
    procedure SBAutre1Click(Sender: TObject);
    procedure Autre1Change(Sender: TObject);
    procedure Autre1Exit(Sender: TObject);
    procedure bComplementsResClick(Sender: TObject);
    procedure Autre2Change(Sender: TObject);
    procedure Autre2Exit(Sender: TObject);
    procedure SBAutre2Click(Sender: TObject);
    procedure SO_AFGESTIONCOMClick(Sender: TObject);
    procedure CBFormatExerciceChange(Sender: TObject);
    procedure bComplementsAffClick(Sender: TObject);
    procedure CBBureauClick(Sender: TObject);
    procedure SBAgenceClick(Sender: TObject);
    procedure Autre0Change(Sender: TObject);
    procedure Autre0Exit(Sender: TObject);
    procedure CBAgenceClick(Sender: TObject);
    procedure SBBureauClick(Sender: TObject);
    procedure CBAssocieClick(Sender: TObject);
    procedure CBModeRegleChange(Sender: TObject);
    procedure SBModeRegleClick(Sender: TObject);
    procedure CBChefGroupeClick(Sender: TObject);
    procedure CBResponsableClick(Sender: TObject);
    procedure bComplementsCliClick(Sender: TObject);
    procedure DesactiverClick(Sender: TObject);
    procedure bEntiteJurClick(Sender: TObject);
    procedure bEntiteJurMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure bEntiteJurMouseEnter(Sender: TObject);
    procedure bDatesClick(Sender: TObject);
  private
    { Déclarations privées }
    fNumPlan : integer;
    bChargementEnCours : boolean;
    CurToolbarButton : TToolbarButton97;
    CurToolBarButtonName : string;
    procedure InitApprec(bInit : boolean);
    procedure InitEntiteJuridique(bInit : boolean);
    procedure InitRessource(bInit:boolean);
    procedure InitAffaire(bInit:boolean);
    procedure InitActivite(bInit:boolean);
    procedure InitClient(bInit:boolean);
    procedure InitPreferences(bInit : boolean);
    procedure InitDates(bInit : boolean);
    procedure ChargePlanRef(NumPlan: integer);
    procedure ChargeJournalRef(NumPlan: integer);
    function ChargeStandard(NumPlan: integer): boolean;
    procedure ChargeStandardCompta;
    procedure ChargeParamSocRef(NumPlan: integer);
  public
    { Déclarations publiques }
  end;

procedure ActiveGI;
//function IsComptaActivee : boolean;
//function IsGIActivee : boolean;
procedure InitMODEREGL;

implementation

{$R *.DFM}

{ TFAssistGI }

function OpenSQLCom(St : string; RO : boolean) : TQuery;
var QCom : TQuery;
begin
  V_PGI_Env.Connected := True;
  QCom := TQuery.Create (nil);
  QCom.DatabaseName := 'DBCom';
  QCom.SQL.Text := St;
  ChangeSQL(QCom);
  QCom.RequestLive := RO;
  QCom.Open;
  result := QCom;
end;

procedure FermeSQLCom (var QCom : TQuery);
begin
  QCom.Close;
  QCom.Free;
  V_PGI_Env.Connected := false;
end;
(*
function IsComptaActivee : boolean;
var QCom : TQuery;
begin
  V_PGI_Env.Connected := True;
  QCom := TQuery.Create (nil);
  QCom.DatabaseName := 'DBCOM';
  QCom.SQL.Text := 'SELECT DAP_NODOSSIER,DAP_NOMEXEC FROM DOSSAPPLI WHERE DAP_NODOSSIER="'+V_PGI_Env.NoDossier+'"  AND DAP_NOMEXEC="CCS5.EXE"';
  ChangeSQL(QCom);
  QCom.RequestLive := True;
  QCom.Open;
  result := not QCom.Eof;
  QCom.Close;
  QCom.Free;
  V_PGI_Env.Connected := false;
  // ### Avec double connexion
{  QCom := OpenSQL ('SELECT DAP_NODOSSIER,DAP_NOMEXEC FROM DOSSAPPLI WHERE DAP_NODOSSIER="'+V_PGI_Env.NoDossier+'" AND DAP_NOMEXEC="CCS5.EXE"',True);
  result := not QCom.Eof;
  Ferme(QCom);}
end;

function IsGIActivee : boolean;
var QCom : TQuery;
begin
if V_PGI_Env=Nil then BEGIN Result:=TRUE ; exit ; END ;
  V_PGI_Env.Connected := True;
  QCom := TQuery.Create (nil);
  QCom.DatabaseName := 'DBCOM';
  QCom.SQL.Text := 'SELECT DAP_NODOSSIER,DAP_NOMEXEC FROM DOSSAPPLI WHERE DAP_NODOSSIER="'+V_PGI_Env.NoDossier+'"  AND DAP_NOMEXEC="CGIS5.EXE"';
  ChangeSQL(QCom);
  QCom.RequestLive := True;
  QCom.Open;
  result := not QCom.Eof;
  QCom.Close;
  QCom.Free;
  V_PGI_Env.Connected := false;
  // ### Avec double connexion
{  QCom := OpenSQL ('SELECT DAP_NODOSSIER,DAP_NOMEXEC FROM DOSSAPPLI WHERE DAP_NODOSSIER="'+V_PGI_Env.NoDossier+'" AND DAP_NOMEXEC="CGIS5.EXE"',True);
  result := not QCom.Eof;
  Ferme(QCom);}
end; *)

procedure ActiveGI;
var QCom : TQuery;
begin
  V_PGI_Env.Connected := True;
  QCom := TQuery.Create (nil);
  QCom.DatabaseName := 'DBCom';
  QCom.SQL.Text := 'SELECT DAP_NODOSSIER,DAP_NOMEXEC,DAP_DATEMODIF,DAP_UTILISATEUR FROM DOSSAPPLI WHERE DAP_NODOSSIER="'+V_PGI_Env.NoDossier+'"  AND DAP_NOMEXEC="CGIS5.EXE"';
  ChangeSQL (QCom);
  QCom.RequestLive := True;
  QCom.Open;
  if QCom.Eof then
  begin
    QCom.Insert;
    InitNew(QCom);
    QCom.FindField('DAP_NODOSSIER').AsString := V_PGI_Env.NoDossier;
    QCom.FindField('DAP_NOMEXEC').AsString := 'CGIS5.EXE';
    QCom.FindField('DAP_DATEMODIF').AsDateTime := Date;
    QCom.FindField('DAP_UTILISATEUR').AsString := V_PGI.USER;
    QCom.Post
  end;
  QCom.Close;
  QCom.Free;
  V_PGI_Env.Connected := false;
end;

// ### si bInit = True , alors on est au lancement de l'application
//       ==> le dossier n'est pas encore initialisé pour la GI
// ### si bInit = False, on est dans l'application.
//       ==> le dossier est déjà initialisé pour la GI.
function LanceAssistantGI_PCL (bInit : boolean) : boolean;
var   FAssistGI: TFAssistGI;
      bRetour : boolean;
      bOnRentreDansAssistant:boolean;
begin
  bRetour := false;
  bOnRentreDansAssistant:=true;
{$IFDEF DEBUG} // si on est en DEBUG, on peut voir le menu d'accès à l'assistant dans les paramètres
                // en entrant par le mot de passe du jour
  if (V_PGI<>nil) then
      begin
      if V_PGI.RunFromLanceur then
            begin
            if GetFlagAppli ('CGIS5.EXE') then  // PL le 17/04/03 : chgt IsAppliActive de AssistPL en GetFlagAppli de galoutil
                begin
                bOnRentreDansAssistant:=false;
                bRetour:=true;
                end
            else
                begin
                bOnRentreDansAssistant:=True;
                end;
            end
      else
            begin
            bOnRentreDansAssistant:=True;
            end;
      end;
{$ELSE}
    if GetFlagAppli ('CGIS5.EXE') then   // PL le 17/04/03 : chgt IsAppliActive de AssistPL en GetFlagAppli de galoutil
        begin
        bOnRentreDansAssistant:=false;
        bRetour:=true;
        end ;
{$ENDIF}

  //mcd 27/10/2003 pour interdire de lancer assistant si des affaires existent
  If ExisteSql('SELECT AFF_AFFAIRE FROM AFFAIRE') then
    begin
    bOnRentreDansAssistant:=false;
    bRetour:=true;
    ActiveGI;
    end ;

  if bOnRentreDansAssistant then
  begin
    FAssistGI := TFAssistGI.Create (Application);
    try
      //
      // Avant assistant de création
      //
        // Initialisation des libellés libres à ".-"
      if bInit then InitZonesLibres;

        // création de l'article ACOMPTE s'il n'existe pas
      if Not ExisteArticle('ACOMPTE') then
            InitArticle;
       // mcd 17/09/03 destructiond es valeurs typeressources non gérées en GI et existant dans socref, en attendant solution autre
      ExecuteSql ('DELETE from CHOIXCOD where cc_type="tre" and cc_code not in ("SAL","ST")');
         // création des unités de base si elles n'existent pas
      InitUnite;
         // création des profils de générations s'ils n'existent pas
      if Not ExisteProfilGener('DEF') then
            InitProfilGener;

        // création de l'etablissement par defaut s'il n'existe pas
{$IFNDEF DEBUG}
     If Not (V_PGI_Env.InBaseCommune) then
{$ENDIF}
      if Not ExisteEtabliss(GetParamSoc('SO_SOCIETE')) then
         InitEtablissement;

      // initialisation de la saisie de l'entité juridique
      FAssistGI.InitEntiteJuridique(bInit);

      // initialisation de la saisie des parametres ressource
      FAssistGI.InitRessource(bInit);

      // initialisation de la saisie des parametres affaire
      FAssistGI.InitAffaire(bInit);

      // initialisation de la saisie des parametres activité
      FAssistGI.InitActivite(bInit);
      // initialisation de l'appreciaiton
      FAssistGI.InitApprec(bInit);

      // initialisation de la saisie des parametres client
      FAssistGI.InitClient(bInit);

      // initialisation de la saisie des preferences GA/GI
      FAssistGI.InitPreferences(bInit);

      // initialisation des dates
      FAssistGI.InitDates(bInit);

      // initialisation des parametres du code affaire (AffaireUtil.pas)
      if bInit then
        InitParamCodeAffaire;

      // initialisation des parpieces affaire (AffaireUtil.pas)
      if bInit then
        ModifLesParpieceAffaire;

      // initialisation de la compta
      if bInit then
        InitCompta;

      // initialisation de la table TXCPTTVA
      if bInit then
        InitTVA;

      // initialisation de la table PARPIECE
      if bInit then
        InitParPiece;

      // Récupération du bob des présentations par défauts (graph, tobViewer mcd 16/09/02
{$IFnDEF EAGLCLIENT}
      if bInit then
         AglIntegreBob (V_PGI_Env.RepLocal+'\BOB\GIS5\GIS50582D001.bob');
      if bInit then  // integre bob pour tables planning mcd 10/12/02
          Begin
          AglIntegreBob (V_PGI_Env.RepLocal+'\BOB\GIS5\GIS50596D001.bob');
          AglIntegreBob (V_PGI_Env.RepLocal+'\BOB\GIS5\GIS50621D001.bob'); //mcd 22/09/03
          end;
{$endif}
      // initialisation de la table MODEREGL
      if bInit then
        begin
        InitMODEREGL;
// On rafraichit le lien entre la combo et la tablette pour que celle-ci soit mise à jour en temps réel
// NE PAS SUPPRIMER !
        FAssistGI.CBModeRegle.Refresh;
        FAssistGI.CBModeRegle.datatype := '';
        FAssistGI.CBModeRegle.datatype := 'TTMODEREGLE';
        FAssistGI.CBModeRegle.Value := GetParamsoc('SO_GCMODEREGLEDEFAUT');
///////////////////
        end;

      //
      // Lancement assistant de création
      //
      FAssistGI.ShowModal;

      //
      // Après assistant de création
      //
    finally
      FAssistGI.Free;
    end;
  bRetour := GetFlagAppli ('CGIS5.EXE');   // PL le 17/04/03 : chgt IsAppliActive de AssistPL en GetFlagAppli de galoutil
  end;

  if bRetour then
  begin
{$IFNDEF SANSCOMPTA}
    if not ExisteSQL('SELECT EX_EXERCICE FROM EXERCICE')  then LanceAssistExo;
{$ENDIF}
  end;

  result := ((bRetour) and (ExisteSQL('SELECT EX_EXERCICE FROM EXERCICE')));
end;

procedure TFAssistGI.FormShow(Sender: TObject);
var NumPlan : integer;
begin
  inherited;

  TToolBarButton97(bEntiteJur).Down := false;
  SO_CODE.Text:=GetParamSoc('SO_SOCIETE') ;
  SO_LIBELLE.Text:=GetParamSoc('SO_LIBELLE') ;
  bAffaire.Enabled := False;
  bAnnuler.Enabled := true;
  bPrecedent.Enabled := true;
  bSuivant.Enabled := true;
  bFin.Enabled := true;
  NumPlan := GetParamSoc('SO_NUMPLANREF');
  MODE_CREATION.Enabled := not (NumPlan>0);
  SO_NUMPLANREF.Enabled := not (NumPlan>0);
  if NumPlan > 0 then SO_NUMPLANREF.Value := Format('%.03d',[integer(GetParamSoc('SO_NUMPLANREF'))]);

// Protection du champ Lien DP
// Si on est dans la base 00, on force lien DP à VRAI, à Faux sinon
SetParamsoc('SO_AFANNUAIREIMPETIERS',FALSE);
if (ctxScot in V_PGI.PGIContexte) then
    begin
    if V_PGI.RunFromLanceur then
        if (V_PGI_Env <> nil) then
            begin
            If (V_PGI_Env.InBaseCommune) then
                begin
                SetParamsoc('SO_AFLIENDP',TRUE);
                SetParamsoc('SO_AFANNUAIREIMPETIERS',TRUE);
                end
            else
                SetParamsoc('SO_AFLIENDP',FALSE);
            end
        else
            SetParamsoc('SO_AFLIENDP',FALSE);

    if V_PGI.SAV then
        SO_AFLIENDP.enabled:= true
    else
        SO_AFLIENDP.enabled:= False;

    end
else
    begin
    SetParamsoc('SO_AFLIENDP',FALSE);
    SO_AFLIENDP.visible:=false;
    end;

SO_AFLIENDP.Checked := GetParamsoc ('SO_AFLIENDP');
CBModeRegle.Value := GetParamsoc ('SO_GCMODEREGLEDEFAUT');
CSO_AFFACTPARRES.Value := GetParamsoc ('SO_AFFACTPARRES');
CBFormatExercice.Value := GetParamSoc ('SO_AFFORMATEXER');
end;

procedure TFAssistGI.bAnnulerClick(Sender: TObject);
begin
  inherited;
  Close;
end;

procedure TFAssistGI.OnBoutonActivationClick(Sender: TObject);
begin
if Sender.ClassName <> 'TToolbarButton97' then exit;
CurToolBarButtonName := TToolbarButton97(Sender).Name;
CurToolbarButton := TToolbarButton97(Sender);

if TToolbarButton97(Sender).ImageIndex = 0 then
  begin
    PopParam.Items[0].Visible := True;
    PopParam.Items[1].Visible := False;
    PopParam.Popup(mouse.CursorPos.x,mouse.CursorPos.y);
  end;
  if TToolbarButton97(Sender).ImageIndex = 1 then
  begin
    inherited;
  if CurToolBarButtonName = 'bEntiteJur' then FicheEtablissement_GC(taCreat);
  end;


end;

procedure TFAssistGI.ActiverClick(Sender: TObject);
begin
  inherited;
  CurToolBarButton.ImageIndex := 1;
  if CurToolBarButtonName = 'bEntiteJur' then SetParamSoc ('SO_ETABLISCPTA',True);
end;

procedure TFAssistGI.DesactiverClick(Sender: TObject);
begin
  inherited;
  CurToolBarButton.ImageIndex := 0;
  if CurToolBarButtonName = 'bEntiteJur' then SetParamSoc ('SO_ETABLISCPTA',false);
end;


procedure TFAssistGI.bCoordonneesClick(Sender: TObject);
begin
  inherited;
  ParamSociete(False,'','SCO_COORDONNEES','',nil,nil,nil,nil,110000220, taConsult) ;
end;

procedure TFAssistGI.PlanComptesGenerauxClick(Sender: TObject);
begin
  inherited;
{$IFNDEF SANSCOMPTA}
  MulticritereCpteGene(taModif) ;
{$ENDIF}
end;

procedure TFAssistGI.JournauxComptablesClick(Sender: TObject);
begin
  inherited;
{$IFNDEF SANSCOMPTA}
  MulticritereJournal(taModif);
{$ENDIF}
end;

procedure TFAssistGI.LibelleAutoClick(Sender: TObject);
begin
  inherited;
{$IFNDEF SANSCOMPTA}
  ParamLibelle ;
{$ENDIF}
end;

procedure TFAssistGI.GuideSaisieClick(Sender: TObject);
begin
  inherited;
{$IFNDEF SANSCOMPTA}
  ParamGuide('','NOR',taModif) ;
{$ENDIF}
end;

procedure TFAssistGI.bParamComptableClick(Sender: TObject);
begin
  inherited;
  ParamSociete(False,'','SCO_COMPTABLES','',nil,ChargePageSocGA,SauvePageSocGA,InterfaceSocGA,110000220);
  ChargeSocieteHalley;
end;

procedure TFAssistGI.bFourchettesClick(Sender: TObject);
begin
  inherited;
  ParamSociete(False,'','SCO_GCPONTCOMPTABLE','',nil,ChargePageSocGA,SauvePageSocGA,InterfaceSocGA,110000220);
  ChargeSocieteHalley;
end;

procedure TFAssistGI.bMonnaieClick(Sender: TObject);
begin
  inherited;
  ParamSociete(False,'','SCO_EURO','',nil,ChargePageSocGA,SauvePageSocGA,InterfaceSocGA,110000220);
  ChargeSocieteHalley;
end;

procedure TFAssistGI.bPreferencesClick(Sender: TObject);
begin
  inherited;
  AGLLanceFiche('CP','ASSPREFCOMPTA','','','');
end;

procedure TFAssistGI.bSuivantClick(Sender: TObject);
begin
  if  MODE_CREATION.Value='SI2' then
  begin
    if ((P.ActivePage = PDebut) and (FICHIER_IMPORT.Enabled)) then
    begin
//      RestaureSavSisco(FICHIER_IMPORT.Text, False);
      FICHIER_IMPORT.Enabled := False;
    end;
  end else
  if  MODE_CREATION.Value='PLC' then
  begin
    if ((P.ActivePage = PDebut) and (SO_NUMPLANREF.Enabled)) then
        if (Not V_PGI.RunFromLanceur)  then
            ChargeStandard(StrToInt(SO_NUMPLANREF.Value));
  end else
  MessageAlerte ('Non disponible');

  inherited;
end;

procedure TFAssistGI.bFinClick(Sender: TObject);
var bRetour : boolean;
CptCollClient, CptCollFourn:string;
begin
  bRetour := True;
  if  MODE_CREATION.Value='SI2' then
  begin
    if ((P.ActivePage = PDebut) and (FICHIER_IMPORT.Enabled)) then
    begin
//      RestaureSavSisco(FICHIER_IMPORT.Text, False);
    end;
  end else
  if  MODE_CREATION.Value='PLC' then
  begin
    if ((P.ActivePage = PDebut) and (SO_NUMPLANREF.Enabled)) then
        if Not V_PGI.RunFromLanceur then
            bRetour := ChargeStandard(StrToInt(SO_NUMPLANREF.Value));
  end else
  begin
    MessageAlerte ('Non disponible');
    exit;
  end;
  if bRetour then
  begin
    inherited;
    // Gestion des maj finales
{$IFNDEF DEBUG}
       //mcd 29/07/03 on met correctement les valeurs qui ne sont pas toujours top dans la socref
    SetParamSoc('SO_GcMultiDepots','X'); //25/08/03
    SetParamSoc('SO_GCINVPERM','-');  //mcd 08/10/03
    SetParamSoc('SO_GCVALOARTDEPOT','-');  //mcd 14/10/03
    SetParamSoc('SO_GcToutEuro','X');
    SetParamSoc('SO_GcVenteEuro','X');
    SetParamSoc('SO_AfProfiLGener','DEF');
    If Not (ExisteSql ('SELECT cc_type FROM CHOIXCOD WHERE CC_TYPE="APG" and Cc_CODE="DEF"'))
      then ExecuteSql ('INSERT into  CHOIXCOD  (cc_type,cc_code,cc_libelle,cc_abrege,cc_libre) values ("APG","DEF","Profil par défaut","Défaut","")');
    SetParamSoc('SO_RetGarantie','-');
    SetParamSoc('SO_AcompteOblig','-');
    SetParamSoc('SO_BtRechTarifFou','-');
    SetParamSoc('SO_BtLIvrVisible','-');
    // on ne gere pas la tarification
    SetParamSoc('SO_AFGESTIONTARIF', '-');  //17/10/01
    // on ne gere pas le code avenant
    SetParamSoc('SO_AFFGESTIONAVENANT', '-');
    // Initialisation des natures de pièces  associées aux affaires
    SetParamSoc('SO_AFNATAFFAIRE', 'AFF');
    SetParamSoc('SO_AFNATPROPOSITION','PAF');
    // Init de la TVA encaisst = Oui et des racines des comptes collectifs clients et fournisseur
    SetParamsoc('SO_OUITVAENC', 'X');
    CptCollClient := GetParamsoc('SO_DEFCOLCLI');
    CptCollFourn := GetParamsoc('SO_DEFCOLFOU');
    if CptCollClient<>'' then
        SetParamsoc('SO_COLLCLIENC', copy(CptCollClient, 1, 3));

    if CptCollFourn<>'' then
        SetParamsoc('SO_COLLFOUENC', copy(CptCollFourn, 1, 3));

    // Active la gestion interne
    ActiveGI;
{$ELSE}
    // Active la gestion interne
    if V_PGI.RunFromLanceur then
        ActiveGI;
{$ENDIF}

    ChargeMagHalley;
    Close;
  end;
  // ### Afficher Message d'erreur ###
end;

procedure TFAssistGI.DossierEuroClick(Sender: TObject);
begin
  inherited;
  SetParamSoc('SO_TENUEEURO',DossierEuro.Checked);
end;

procedure TFAssistGI.MODE_CREATIONChange(Sender: TObject);
begin
  if bChargementEnCours then exit;
  inherited;
  if MODE_CREATION.value = 'PLC' then
  begin
    FICHIER_IMPORT.Visible := False;
    SO_NUMPLANREF.Visible := True;
    bSuivant.Enabled := True;
    bFin.Enabled := True;
    LABEL_CREATION.Caption := Msg.Mess[2];
  end else if MODE_CREATION.Value = 'SI2' then
  begin
    FICHIER_IMPORT.Visible := True;
    SO_NUMPLANREF.Visible := False;
    bSuivant.Enabled := True;
    bFin.Enabled := True;
    LABEL_CREATION.Caption := Msg.Mess[3];
  end else
  begin
    FICHIER_IMPORT.Visible := True;
    SO_NUMPLANREF.Visible := False;
    bSuivant.Enabled := True;
    bFin.Enabled := False;
    LABEL_CREATION.Caption := Msg.Mess[3];
  end;
end;

procedure TFAssistGI.FICHIER_IMPORTClick(Sender: TObject);
begin
  inherited;
  OpenDialog.Execute;
  FICHIER_IMPORT.Text := OpenDialog.FileName;
end;

function TFAssistGI.ChargeStandard (NumPlan : integer) : boolean;
var bRetour : boolean;
begin
  fNumPlan := NumPlan;
  if Blocage(['nrBatch','nrCloture'],True,'nrCloture') then
  begin result := false; exit ; end;
  if Transactions (ChargeStandardCompta, 1) <> oeOk then
  begin
    // Erreur durant la mise à jour depuis les standards
     PGIBoxAf('Chargement du standard impossible. Veuillez recommencer.','Chargement d''un standard');
    bRetour := false;
  end else
  begin
    SO_NUMPLANREF.Enabled := false;
    SetParamSoc('SO_NUMPLANREF',StrToInt(SO_NUMPLANREF.Value));
    bRetour := True;
  end;
  Bloqueur('nrCloture',False) ;
  result := bRetour;
end;

procedure TFAssistGI.ChargeStandardCompta;
begin
  ChargeParamSocRef(fNumPlan);
  ChargeSocieteHalley;
//  CreeSoucheMini(GetParamSoc('SO_SOCIETE'));
  ChargePlanRef (fNumPlan);
  ChargeJournalRef (fNumPlan);
end;

procedure TFAssistGI.ChargePlanRef (NumPlan : integer);
var Q, QPlanRef : TQuery ;
    s : string ;
begin
  // SELECT * : indispensable pour chargement du plan de référence : voir fonctions de la compta AssitPL
  QPlanRef := OpenSQLCom ('SELECT * FROM PLANREF WHERE PR_NUMPLAN = ' +
                      IntToStr(NumPlan) + ' ORDER BY PR_COMPTE',True) ;
  Q:=OpenSQL('SELECT * FROM GENERAUX WHERE G_GENERAL="'+W_W+'"',False) ;
  InitMove(RecordsCount(QPlanRef),'Chargement du plan comptable');
  while not QPlanRef.EOF do
  begin
    s := BourreLaDonc(QPlanRef.FindField('PR_COMPTE').AsString,fbGene) ;
    if Not Presence('GENERAUX', 'G_GENERAL', s) then
    begin
      Q.Insert ;
      InitNew(Q) ;
      Q.FindField('G_GENERAL').AsString := s ;
      Q.FindField('G_LIBELLE').AsString := QPlanRef.FindField('PR_LIBELLE').AsString ;
      Q.FindField('G_ABREGE').AsString := QPlanRef.FindField('PR_ABREGE').AsString ;
      Q.FindField('G_CENTRALISABLE').AsString := QPlanRef.FindField('PR_CENTRALISABLE').AsString ;
      Q.FindField('G_SOLDEPROGRESSIF').AsString := QPlanRef.FindField('PR_SOLDEPROGRESSIF').AsString ;
      Q.FindField('G_SAUTPAGE').AsString := QPlanRef.FindField('PR_SAUTPAGE').AsString ;
      Q.FindField('G_TOTAUXMENSUELS').AsString := QPlanRef.FindField('PR_TOTAUXMENSUELS').AsString ;
      Q.FindField('G_COLLECTIF').AsString := QPlanRef.FindField('PR_COLLECTIF').AsString ;
      Q.FindField('G_BLOCNOTE').AsString := QPlanRef.FindField('PR_BLOCNOTE').AsString ;
      Q.FindField('G_SENS').AsString := QPlanRef.FindField('PR_SENS').AsString ;
      Q.FindField('G_LETTRABLE').AsString := QPlanRef.FindField('PR_LETTRABLE').AsString ;
      Q.FindField('G_POINTABLE').AsString := QPlanRef.FindField('PR_POINTABLE').AsString ;
      Q.FindField('G_VENTILABLE1').AsString := QPlanRef.FindField('PR_VENTILABLE1').AsString ;
      Q.FindField('G_VENTILABLE2').AsString := QPlanRef.FindField('PR_VENTILABLE2').AsString ;
      Q.FindField('G_VENTILABLE3').AsString := QPlanRef.FindField('PR_VENTILABLE3').AsString ;
      Q.FindField('G_VENTILABLE4').AsString := QPlanRef.FindField('PR_VENTILABLE4').AsString ;
      Q.FindField('G_VENTILABLE5').AsString := QPlanRef.FindField('PR_VENTILABLE5').AsString ;
      Q.FindField('G_NATUREGENE').AsString := QPlanRef.FindField('PR_NATUREGENE').AsString ;
      Q.Post ;
      MoveCur(False);
    end;
    QPlanRef.Next ;    // ##PLPBNEXT ?
  end;
  Ferme(Q) ;
  FermeSQLCom(QPlanRef);
  FiniMove;
end;

procedure TFAssistGI.ChargeJournalRef (NumPlan : integer);
var Q, QJournalRef : TQuery ;
begin
  // SELECT * : indispensable pour chargement du journal de référence : voir fonctions de la compta AssitPL
  QJournalRef := OpenSQLCom ('SELECT * FROM JALREF WHERE JR_NUMPLAN = ' +
                      IntToStr(NumPlan) + ' ORDER BY JR_JOURNAL',True) ;
  Q:=OpenSQL('SELECT * FROM JOURNAL WHERE J_JOURNAL="'+W_W+'"',False) ;
  InitMove(RecordsCount(QJournalRef),'Chargement des journaux');
  while not QJournalRef.EOF do
  begin
    if Not Presence('JOURNAL', 'J_JOURNAL', QJournalRef.FindField('JR_JOURNAL').AsString) then
    begin
      Q.Insert ;
      InitNew(Q) ;
      Q.FindField('J_JOURNAL').AsString := QJournalRef.FindField('JR_JOURNAL').AsString ;
      Q.FindField('J_LIBELLE').AsString := QJournalRef.FindField('JR_LIBELLE').AsString ;
      Q.FindField('J_ABREGE').AsString := QJournalRef.FindField('JR_ABREGE').AsString ;
      Q.FindField('J_NATUREJAL').AsString := QJournalRef.FindField('JR_NATUREJAL').AsString ;
      Q.FindField('J_COMPTEURNORMAL').AsString := QJournalRef.FindField('JR_COMPTEURNORMAL').AsString ;
      Q.FindField('J_COMPTEURSIMUL').AsString := QJournalRef.FindField('JR_COMPTEURSIMUL').AsString ;
      Q.FindField('J_MODESAISIE').AsString := QJournalRef.FindField('JR_MODESAISIE').AsString ;
      Q.Post ;
      MoveCur (False);
    end;
    QJournalRef.Next ;      // ##PLPBNEXT ?
  end;
  Ferme(Q) ;
  FermeSQLCom(QJournalRef);
  FiniMove;
end;

procedure TFAssistGI.ChargeParamSocRef (NumPlan : integer);
var QParamSocRef : TQuery ;
begin
  // SELECT * : indispensable pour chargement des elements de la socref : voir fonctions de la compta AssitPL
  QParamSocRef := OpenSQLCom ('SELECT * FROM PARSOCREF WHERE PRR_NUMPLAN = ' +
                      IntToStr(NumPlan) + ' ORDER BY PRR_SOCNOM',True) ;
  InitMove(RecordsCount(QParamSocRef),'Chargement des paramètres société');
  while not QParamSocRef.EOF do
  begin
    ExecuteSQL ('UPDATE PARAMSOC SET SOC_DATA="'+QParamSocRef.FindField('PRR_SOCDATA').AsString+
                        '" WHERE SOC_NOM="'+QParamSocRef.FindField('PRR_SOCNOM').AsString+'"');
    QParamSocRef.Next ;     // ##PLPBNEXT ?
    MoveCur (False);
  end;
  FermeSQLCom(QParamSocRef);
  FiniMove;
end;

Procedure FicheEtablissement_GC ( Mode : TActionFiche ) ;
Var Arguments : string;
begin
if Blocage(['nrCloture','nrBatch'],True,'nrAucun') then Exit ;
Arguments:=ActionToString(Mode);
AGLLanceFiche('GC','GCETABLISS_MUL','','',Arguments) ;
end;

procedure TFAssistGI.bPrefMissionClick(Sender: TObject);
begin
  inherited;
ParamSociete(False,'','SCO_AFFPREFERENCES','',nil,ChargePageSocGA,SauvePageSocGA,InterfaceSocGA,110000220);
end;

procedure TFAssistGI.bRessourceClick(Sender: TObject);
begin
  inherited;
ParamSociete(False,'','SCO_RESSOURCE','',nil,ChargePageSocGA,SauvePageSocGA,InterfaceSocGA,110000220);
end;

procedure TFAssistGI.bActiviteClick(Sender: TObject);
begin
  inherited;
ParamSociete(False,'','SCO_AFFACTIVITE','',nil,ChargePageSocGA,SauvePageSocGA,InterfaceSocGA,110000220);
end;

 procedure TFAssistGI.bComplementFactClick(Sender: TObject);
begin
  inherited;
ParamSociete(False,'','SCO_FACTUREECLATEE','',nil,ChargePageSocGA,SauvePageSocGA,InterfaceSocGA,110000220);
ChargeSocieteHalley;
if (getparamsoc('SO_AFFACTPARRES') <> 'SAN') and
 ((getparamsoc('SO_AFFACTMODEPEC')='')
  or (GetparamSoc('SO_AFFACTPRESDEFAUT')='')
  or (GetparamSoc('SO_AFFACTFRAISDEFAUT')='')
  or (GetparamSoc('SO_AFFACTFOURDEFAUT')='')
  or (GetparamSoc('SO_AFFACTRESSDEFAUT')='') )
     then begin
       PGIInfo('Vos paramètres ne sont pas tous renseignés','Assistant GI/Facture éclatée');
       bComplementFactClick(self) ;
     end;
end;

procedure TFAssistGI.bDatesClick(Sender: TObject);
begin
  inherited;
ParamSociete(False,'','SCO_AFDATES','',nil,ChargePageSocGA,SauvePageSocGA,InterfaceSocGA,110000220);

end;

procedure TFAssistGI.bTypMissionClick(Sender: TObject);
begin
  inherited;
ParamTable ('AFFAIREPART1',taCreat,0,nil);
end;
procedure TFAssistGI.bAffaireClick(Sender: TObject);
begin
  inherited;
LanceAssistCodeAffaire;
end;

procedure TFAssistGI.SO_AFLIENDPClick(Sender: TObject);
begin
  inherited;
  SetParamSoc('SO_AFLIENDP', SO_AFLIENDP.Checked);
end;

procedure TFAssistGI.CheckBox2Click(Sender: TObject);
begin
  inherited;
if (TCheckBox(Sender).Checked=true) then
    begin
//    ModifierTablette('CC', 'ZLD', 'AL1', 'Groupe de ressources', 3);
    ModifierTablette('CC', 'ZLI', 'RT1', TraduitGA('Groupe de ressources'), 3);
    Autre0.Enabled:=false;
    Autre0.Text:=TraduitGA('Groupe de ressources');
//    AvertirTable('AFCARESTABLE');
    AvertirTable('GCZONELIBRE');
    end
else
    begin
    Autre0.Enabled:=true;
    Autre0.Text:=' ';
    Autre0.Text:='';
    end;

end;

procedure TFAssistGI.SBCollabClick(Sender: TObject);
begin
  inherited;
ParamTable ('AFTLIBRERES1',taCreat,0,nil,6, Autre0.Text);
end;

procedure TFAssistGI.SBAutre1Click(Sender: TObject);
begin
  inherited;
ParamTable ('AFTLIBRERES2',taCreat,0,nil, 6, Autre1.Text);

end;

procedure TFAssistGI.SBAutre2Click(Sender: TObject);
begin
  inherited;
ParamTable ('AFTLIBRERES3',taCreat,0,nil,6, Autre2.Text);

end;

procedure TFAssistGI.bComplementsResClick(Sender: TObject);
var
Ret, arg:string;
Crit, champ, valeur : string;
x:integer;
begin
  inherited;
arg:='';
nextprevcontrol(self);
if (Autre0.text<>'') and (AnsiUppercase(Autre0.text)<>'TABLE LIBRE 1') then arg:=arg+'CRIT1='+Autre0.text+';';
if (Autre1.text<>'') and (AnsiUppercase(Autre1.text)<>'TABLE LIBRE 2') then arg:=arg+'CRIT2='+Autre1.text+';';
if (Autre2.text<>'') and (AnsiUppercase(Autre2.text)<>'TABLE LIBRE 3') then arg:=arg+'CRIT3='+Autre2.text+';';

Ret:=AFLanceFiche_CreatLibelRess('RES;CRITERES;' + arg) ;

Crit:=(Trim(ReadTokenSt(Ret)));
while (Crit<>'') do
    begin
    Champ:=Crit;
    X:=pos('=',Crit);
    if x<>0 then
           begin
           Champ:=copy(Crit,1,X-1);
           Valeur:=Copy (Crit,X+1,length(Crit)-X);
           end;

    if (Champ='CRIT1') then
        begin
        Autre0.text:=Valeur;
        end
    else
    if (Champ='CRIT2') then
        begin
        Autre1.text:=Valeur;
        end
    else
    if (Champ='CRIT3') then
        begin
        Autre2.text:=Valeur;
        end;
    Crit:=(Trim(ReadTokenSt(Ret)));
    end;
end;

procedure TFAssistGI.Autre1Change(Sender: TObject);
begin
  inherited;
if (THCritMaskEdit(Sender).Text<>'') and (AnsiUppercase(Trim(THCritMaskEdit(Sender).Text))<>'TABLE LIBRE 2') then
    begin
    SBAutre1.Enabled:=true;
//    Autre2.Enabled:=true;
    end
else
    begin
    SBAutre1.Enabled:=false;
//    Autre2.Enabled:=false;
//    Autre2.Text:='';
    end;

end;

procedure TFAssistGI.Autre1Exit(Sender: TObject);
begin
  inherited;
if (THCritMaskEdit(Sender).Text<>'') then
    begin
//    ModifierTablette('CC', 'ZLD', 'AL2', Autre1.Text , 3);
    ModifierTablette('CC', 'ZLI', 'RT2', Autre1.Text, 3);
    end
else
    begin
//    ModifierTablette('CC', 'ZLD', 'AL2', 'Table libre 2', 3);
    ModifierTablette('CC', 'ZLI', 'RT2', 'Table libre 2', 3);
    end;

//AvertirTable('AFCARESTABLE')
AvertirTable('GCZONELIBRE');
end;

procedure TFAssistGI.Autre2Change(Sender: TObject);
begin
  inherited;
if (THCritMaskEdit(Sender).Text<>'') and (AnsiUppercase(Trim(THCritMaskEdit(Sender).Text))<>'TABLE LIBRE 3')  then
    begin
    SBAutre2.Enabled:=true;
    end
else
    begin
    SBAutre2.Enabled:=false;
    end;

end;

procedure TFAssistGI.Autre2Exit(Sender: TObject);
begin
  inherited;
 if (THCritMaskEdit(Sender).Text<>'') then
    begin
//    ModifierTablette('CC', 'ZLD', 'AL3', Autre2.Text , 3);
    ModifierTablette('CC', 'ZLI', 'RT3', Autre2.Text, 3);
    end
else
    begin
//    ModifierTablette('CC', 'ZLD', 'AL3', 'Table libre 3', 3);
    ModifierTablette('CC', 'ZLI', 'RT3', 'Table libre 3', 3);
    end;

//AvertirTable('AFCARESTABLE')
AvertirTable('GCZONELIBRE');
end;


procedure TFAssistGI.SO_AFGESTIONCOMClick(Sender: TObject);
begin
  inherited;
SetParamsoc('SO_AFGESTIONCOM', SO_AFGESTIONCOM.Checked);
end;


procedure TFAssistGI.CBFormatExerciceChange(Sender: TObject);
var
format:string;
begin
  inherited;
format := CBFormatExercice.Value;
SetParamsoc('SO_AFFORMATEXER', format);
If (GetParamSoc('So_AfFormatExer')='AUC') then BAffaire.enabled:=true
   else begin
        Baffaire.enabled:=False;
          // il faut remettre les valeurs pas défaut de l'exercice ...
        InitParamCodeAffaire;
         // valeur ecraser par fct précédente
        SetParamsoc('SO_AFFORMATEXER', format);
        end;
ModifParamCodeAffaire( format) ;
end;

 procedure TFAssistGI.CSO_AFFACTPARRESChange(Sender: TObject);
begin
  inherited;          //mcd 15/01/03 ajout fct
SetParamsoc('SO_AFFACTPARRES', CSO_AFFACTPARRES.value);
If (Cso_AfFactParRes.value)<>'SAN' then BComplementFact.enabled :=true
 else  BComplementFact.enabled :=false;
end;

procedure TFAssistGI.bComplementsAffClick(Sender: TObject);
var
Choix:string;
begin
  inherited;
Choix:=AFLanceFiche_CreatLibelRess('AFF;CRITERES') ;
end;

procedure TFAssistGI.CBBureauClick(Sender: TObject);
begin
  inherited;
if (TCheckBox(Sender).Checked=true) then
    begin
    if (CBAgence.Checked=true) then
        ModifierTablette('CC', 'ZLI', 'CT2', 'Agence', 3);

    ModifierTablette('CC', 'ZLI', 'CT1', 'Bureau', 3);
    SBBureau.Enabled:=true;
    end
else
    begin
    if (CBAgence.Checked=true) then
        begin
        ModifierTablette('CC', 'ZLI', 'CT1', 'Agence', 3);
        ModifierTablette('CC', 'ZLI', 'CT2', 'Table libre 2', 3);
        end
    else
        ModifierTablette('CC', 'ZLI', 'CT1', 'Table libre 1', 3);

    SBBureau.Enabled:=false;
    end;

AvertirTable('GCZONELIBRE');
end;

procedure TFAssistGI.CBAgenceClick(Sender: TObject);
begin
  inherited;
if (TCheckBox(Sender).Checked=true) then
    begin
    if (CBBureau.Checked=true) then
        ModifierTablette('CC', 'ZLI', 'CT2', 'Agence', 3)
    else
        ModifierTablette('CC', 'ZLI', 'CT1', 'Agence', 3);

    SBAgence.Enabled:=true;
    end
else
    begin
    if (CBBureau.Checked=true) then
        ModifierTablette('CC', 'ZLI', 'CT2', 'Table libre 2', 3)
    else
        ModifierTablette('CC', 'ZLI', 'CT1', 'Table libre 1', 3);

    SBAgence.Enabled:=false;
    end;

AvertirTable('GCZONELIBRE');
end;

procedure TFAssistGI.CBAssocieClick(Sender: TObject);
begin
  inherited;
if CBAssocie.Checked then
    begin
    if CBChefGroupe.Checked then
        begin
        ModifierTablette('CC', 'ZLI', 'CR2', 'Chef de groupe', 3);
        if CBResponsable.Checked then
            ModifierTablette('CC', 'ZLI', 'CR3', 'Responsable', 3);
        end
    else
        begin
        if CBResponsable.Checked then
            ModifierTablette('CC', 'ZLI', 'CR2', 'Responsable', 3);
        end;

    ModifierTablette('CC', 'ZLI', 'CR1', 'Associé', 3);
    end
else
    begin
    if CBChefGroupe.Checked then
        begin
        ModifierTablette('CC', 'ZLI', 'CR1', 'Chef de groupe', 3);
        if CBResponsable.Checked then
            begin
            ModifierTablette('CC', 'ZLI', 'CR2', 'Responsable', 3);
            ModifierTablette('CC', 'ZLI', 'CR3', TraduitGA('Ressource 3'), 3);
            end
        else
            ModifierTablette('CC', 'ZLI', 'CR2', TraduitGA('Ressource 2'), 3);
        end
    else
        begin
        if CBResponsable.Checked then
            begin
            ModifierTablette('CC', 'ZLI', 'CR1', 'Responsable', 3);
            ModifierTablette('CC', 'ZLI', 'CR2', TraduitGA('Ressource 2'), 3);
            end
        else
            begin
            ModifierTablette('CC', 'ZLI', 'CR1', TraduitGA('Ressource 1'), 3);
            end;
        end;
    end;

AvertirTable('GCZONELIBRE');
end;

procedure TFAssistGI.CBChefGroupeClick(Sender: TObject);
begin
  inherited;
if CBChefGroupe.Checked then
    begin
    if CBAssocie.Checked then
        begin
        ModifierTablette('CC', 'ZLI', 'CR2', 'Chef de groupe', 3);
        if CBResponsable.Checked then
            ModifierTablette('CC', 'ZLI', 'CR3', 'Responsable', 3);
        end
    else
        begin
        ModifierTablette('CC', 'ZLI', 'CR1', 'Chef de groupe', 3);
        if CBResponsable.Checked then
            ModifierTablette('CC', 'ZLI', 'CR2', 'Responsable', 3);
        end;
    end
else
    begin
    if CBAssocie.Checked then
        begin
        if CBResponsable.Checked then
            begin
            ModifierTablette('CC', 'ZLI', 'CR2', 'Responsable', 3);
            ModifierTablette('CC', 'ZLI', 'CR3', TraduitGA('Ressource 3'), 3);
            end
        else
            ModifierTablette('CC', 'ZLI', 'CR2', TraduitGA('Ressource 2'), 3);
        end
    else
        begin
        if CBResponsable.Checked then
            begin
            ModifierTablette('CC', 'ZLI', 'CR1', 'Responsable', 3);
            ModifierTablette('CC', 'ZLI', 'CR2', TraduitGA('Ressource 2'), 3);
            end
        else
            ModifierTablette('CC', 'ZLI', 'CR1', TraduitGA('Ressource 1'), 3);
        end;
    end;

AvertirTable('GCZONELIBRE');
end;

procedure TFAssistGI.CBResponsableClick(Sender: TObject);
begin
  inherited;
if CBResponsable.Checked then
    begin
    if CBChefGroupe.Checked then
        begin
        if CBAssocie.Checked then
            ModifierTablette('CC', 'ZLI', 'CR3', 'Responsable', 3)
        else
            ModifierTablette('CC', 'ZLI', 'CR2', 'Responsable', 3);
        end
    else
        begin
        if CBAssocie.Checked then
            ModifierTablette('CC', 'ZLI', 'CR2', 'Responsable', 3)
        else
            ModifierTablette('CC', 'ZLI', 'CR1', 'Responsable', 3);
        end;
    end
else
    begin
    if CBChefGroupe.Checked then
        begin
        if CBAssocie.Checked then
            ModifierTablette('CC', 'ZLI', 'CR3', TraduitGA('Ressource 3'), 3)
        else
            ModifierTablette('CC', 'ZLI', 'CR2', TraduitGA('Ressource 2'), 3);
        end
    else
        begin
        if CBAssocie.Checked then
            ModifierTablette('CC', 'ZLI', 'CR2', TraduitGA('Ressource 2'), 3)
        else
            ModifierTablette('CC', 'ZLI', 'CR1', TraduitGA('Ressource 1'), 3);
        end;
    end;

AvertirTable('GCZONELIBRE');
end;

procedure TFAssistGI.SBAgenceClick(Sender: TObject);
begin
  inherited;
ParamTable ('GCLIBRETIERS2',taCreat,0,nil,6, 'Agence');
end;

procedure TFAssistGI.Autre0Change(Sender: TObject);
begin
  inherited;
if (THCritMaskEdit(Sender).Text<>'') and (AnsiUppercase(Trim(THCritMaskEdit(Sender).Text))<>'TABLE LIBRE 1') then
    begin
    SBCollab.Enabled:=true;
//    Autre1.Enabled:=true;
    end
else
    begin
    SBCollab.Enabled:=false;
//    Autre1.Enabled:=false;
//    Autre1.Text:='';
    end;

end;

procedure TFAssistGI.Autre0Exit(Sender: TObject);
begin
  inherited;
if (THCritMaskEdit(Sender).Text<>'') then
    begin
//    ModifierTablette('CC', 'ZLD', 'AL1', Autre0.Text , 3);
    ModifierTablette('CC', 'ZLI', 'RT1', Autre0.Text, 3);
    end
else
    begin
//    ModifierTablette('CC', 'ZLD', 'AL1', 'Table libre 1', 3);
    ModifierTablette('CC', 'ZLI', 'RT1', 'Table libre 1', 3);
    end;

//AvertirTable('AFCARESTABLE')
AvertirTable('GCZONELIBRE');
end;

procedure TFAssistGI.SBBureauClick(Sender: TObject);
begin
  inherited;
ParamTable ('GCLIBRETIERS1',taCreat,0,nil,6, 'Bureau');
end;

procedure TFAssistGI.CBModeRegleChange(Sender: TObject);
begin
  inherited;
SetParamsoc('SO_GCMODEREGLEDEFAUT', CBModeRegle.Value);
end;

procedure TFAssistGI.SBModeRegleClick(Sender: TObject);
begin
  inherited;
FicheRegle_AGL ('',False,taModif) ;
AvertirTable('TTMODEREGLE');
// On rafraichit le lien entre la combo et la tablette pour que celle-ci soit mise à jour en temps réel
// NE PAS SUPPRIMER !
self.CBModeRegle.Refresh;
self.CBModeRegle.datatype := '';
self.CBModeRegle.datatype := 'TTMODEREGLE';
self.CBModeRegle.value := GetParamsoc('SO_GCMODEREGLEDEFAUT');
////////////////////////////
end;


procedure TFAssistGI.bComplementsCliClick(Sender: TObject);
var
Choix, arg:string;
begin
  inherited;
arg:='';
if CBBureau.Checked then arg:=arg+'BUREAU;';
if CBAgence.Checked then arg:=arg+'AGENCE;';
if CBAssocie.Checked then arg:=arg+'ASSOCIE;';
if CBChefGroupe.Checked then arg:=arg+'CHEFGROUPE;';
if CBResponsable.Checked then arg:=arg+'RESPONSABLE;';

Choix:=AFLanceFiche_CreatLibelClt(arg) ;
end;

procedure TFAssistGI.bEntiteJurClick(Sender: TObject);
begin
  inherited;
FicheEtablissement_GC(taCreat);
end;

procedure TFAssistGI.bEntiteJurMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  inherited;
  if CurToolBarButton = nil then exit;
  if ((Button = mbRight) and (CurToolBarButton.ImageIndex=1)) then
  begin
    PopParam.Items[0].Visible := False;
    PopParam.Items[1].Visible := True;
    PopParam.Popup(mouse.CursorPos.x,mouse.CursorPos.y);
  end;
  inherited;
end;

procedure TFAssistGI.bEntiteJurMouseEnter(Sender: TObject);
begin
  inherited;
  CurToolbarButton := TToolbarButton97(Sender);
  CurToolBarButtonName := TToolbarButton97(Sender).Name;
end;

(*******************************************************************************
* fonctions d'initialisation des tables initialisée pour la compta, la gescom ou ...
* avec modifs souvent propres à la GA/GI
*******************************************************************************)
// Création de l'article 'ACOMPTE                          X'
procedure InitArticle;
var  OB , OB_DETAIL : TOB;
begin
  // Création de l'enregistrement
  OB := TOB.Create('un article',nil,-1);
  OB_DETAIL := TOB.Create('ARTICLE',OB,-1);
  OB_DETAIL.PutValue('GA_ARTICLE','ACOMPTE                          X');
  OB_DETAIL.PutValue('GA_CODEARTICLE','ACOMPTE');
  OB_DETAIL.PutValue('GA_LIBELLE','Acompte forfaitaire');
  OB_DETAIL.PutValue('GA_TYPEARTICLE','PRE');
  OB_DETAIL.PutValue('GA_FAMILLETAXE1','NOR');
  OB_DETAIL.PutValue('GA_DATESUPPRESSION',idate2099);
  OB_DETAIL.PutValue('GA_STATUTART','UNI');
  OB_DETAIL.PutValue('GA_CALCPRIXHT','AUC');
  OB_DETAIL.PutValue('GA_CALCPRIXTTC','AUC');
  OB_DETAIL.PutValue('GA_PRIXPOURQTE',1);
  OB_DETAIL.PutValue('GA_ACTIVITEREPRISE','F');
  OB_DETAIL.PutValue('GA_QUALIFUNITEVTE',GetParamSoc('SO_AFMESUREACTIVITE'));
  OB_DETAIL.PutValue('GA_QUALIFUNITEACT',GetParamSoc('SO_AFMESUREACTIVITE'));
  OB.InsertOrUpdateDB(True);
  OB.Free;
end;

// Création des unités minimales Heure et Jour
procedure InitUnite;
var  OB , OB_DETAIL : TOB;
begin
  // Création de l'enregistrement
OB := TOB.Create('une unite',nil,-1);

if Not ExisteUnite('H') then
   begin
   OB_DETAIL := TOB.Create('MEA',OB,-1);
   OB_DETAIL.PutValue('GME_QUALIFMESURE','TEM');
   OB_DETAIL.PutValue('GME_MESURE','H');
   OB_DETAIL.PutValue('GME_LIBELLE','Heure');
   OB_DETAIL.PutValue('GME_QUOTITE',1);
   end;

if Not ExisteUnite('J') then
   begin
   OB_DETAIL := TOB.Create('MEA',OB,-1);
   OB_DETAIL.PutValue('GME_QUALIFMESURE','TEM');
   OB_DETAIL.PutValue('GME_MESURE','J');
   OB_DETAIL.PutValue('GME_LIBELLE','Jour');
   OB_DETAIL.PutValue('GME_QUOTITE',7);
   end;

if (OB.Detail.Count<>0) then
   OB.InsertOrUpdateDB(True);
OB.Free;
end;

// Création d'un enregistrement de profil gener
procedure InitProfilGener;
var  OB , OB_DETAIL : TOB;
begin
  // Création de l'enregistrement
  OB := TOB.Create('un profil',nil,-1);
  OB_DETAIL := TOB.Create('PROFILGENER',OB,-1);
  OB_DETAIL.PutValue('APG_PROFILGENER','DEF');
  OB_DETAIL.PutValue('APG_NUMORDRE',0);
  OB_DETAIL.PutValue('APG_ELEMENT','A00');
  OB_DETAIL.PutValue('APG_CUMUL','X');

  OB_DETAIL := TOB.Create('PROFILGENER',OB,-1);
  OB_DETAIL.PutValue('APG_PROFILGENER','DEF');
  OB_DETAIL.PutValue('APG_NUMORDRE',1);
  OB_DETAIL.PutValue('APG_ELEMENT','G00');
  OB_DETAIL.PutValue('APG_CUMUL','-');
  OB_DETAIL.PutValue('APG_TYPEARTICLE','PRE');
  OB_DETAIL.PutValue('APG_ARTICLE','ACOMPTE                          X');
  OB_DETAIL.PutValue('APG_CODEARTICLE','ACOMPTE');
  OB_DETAIL.PutValue('APG_LIBART','Montant forfaitaire');
  OB.InsertOrUpdateDB(True);
  OB.Free;
end;

// initialisation d'un établissement par défaut
(*procedure InitEtablissement;
var  OB , OB_DETAIL : TOB;
begin
if (GetParamSoc('SO_SOCIETE')='') then exit;

  // Création de l'établissement par défaut
  OB := TOB.Create('Les établissements',nil,-1);
  OB_DETAIL := TOB.Create('ETABLISS',OB,-1);
  OB_DETAIL.PutValue('ET_ETABLISSEMENT',GetParamSoc('SO_SOCIETE'));
  OB_DETAIL.PutValue('ET_SOCIETE',GetParamSoc('SO_SOCIETE'));
  OB_DETAIL.PutValue('ET_LIBELLE',GetParamSoc('SO_LIBELLE'));
  OB_DETAIL.PutValue('ET_ABREGE',GetParamSoc('SO_LIBELLE'));
  OB_DETAIL.PutValue('ET_ADRESSE1',GetParamSoc('SO_ADRESSE1'));
  OB_DETAIL.PutValue('ET_ADRESSE2',GetParamSoc('SO_ADRESSE2'));
  OB_DETAIL.PutValue('ET_ADRESSE3',GetParamSoc('SO_ADRESSE3'));
  OB_DETAIL.PutValue('ET_CODEPOSTAL',GetParamSoc('SO_CODEPOSTAL'));
  OB_DETAIL.PutValue('ET_VILLE',GetParamSoc('SO_VILLE'));
  OB_DETAIL.PutValue('ET_PAYS',GetParamSoc('SO_PAYS'));
  OB_DETAIL.PutValue('ET_TELEPHONE',GetParamSoc('SO_TELEPHONE'));
  OB_DETAIL.PutValue('ET_FAX',GetParamSoc('SO_FAX'));
  OB_DETAIL.PutValue('ET_TELEX',GetParamSoc('SO_TELEX'));
  OB_DETAIL.PutValue('ET_SIRET',GetParamSoc('SO_SIRET'));
  OB_DETAIL.PutValue('ET_APE',GetParamSoc('SO_APE'));
  OB_DETAIL.PutValue('ET_JURIDIQUE',GetParamSoc('SO_NATUREJURIDIQUE'));
  OB.InsertOrUpdateDB(True);
  OB.Free;

  // Etablissement par défaut = société = établissement unique.
  SetParamSoc('SO_ETABLISDEFAUT',GetParamSoc('SO_SOCIETE'));
end;
*)
// initialisation de la compta
procedure InitCompta;
var
fDecPrix, fDecQte : currency;
lgCptes : integer;
sBourrage:string;
StrBase:string;
i:integer;
begin
StrBase:='';
lgCptes := GetParamsoc('SO_LGCPTEGEN');
sBourrage := GetParamsoc('SO_BOURREGEN');
if sBourrage='' then sBourrage:='0';

if (lgCptes >= 3) then
    begin
    for i:=1 to lgCptes-3 do StrBase := StrBase + sBourrage;
// Valeur par défaut comptes collectifs
    SetParamsoc('SO_DEFCOLCLI', '411'+StrBase);
    SetParamsoc('SO_DEFCOLFOU', '401'+StrBase);
    SetParamsoc('SO_DEFCOLSAL', '421'+StrBase);
    SetParamsoc('SO_DEFCOLDDIV', '411'+StrBase);
    SetParamsoc('SO_DEFCOLCDIV', '401'+StrBase);
    SetParamsoc('SO_DEFCOLDIV', '467'+StrBase);
// Valeur par défaut compte de vente
    SetParamsoc('SO_GCCPTEHTVTE', '706'+StrBase);
    end;

// Valeur par défaut TVA Encaissement
SetParamsoc('SO_TVAENCAISSEMENT', 'TE');

// Modification de J_JOURNAL='VEN' dans la table JOURNAL : J_MULTIDEVISE='-'
//  Q := OpenSQLCom ('SELECT * FROM JOURNAL WHERE J_JOURNAL ="VEN"',True) ;
//  try
//  if not Q.EOF then
    ExecuteSQL ('UPDATE JOURNAL SET J_MULTIDEVISE="X" WHERE J_JOURNAL="VEN"');
//  finally
//  FermeSQLCom(Q);
//  end;

// init du nombre de décimales pour les prix et les quantités
fDecPrix:=GetParamsoc('SO_DECPRIX');
fDecQte:=GetParamsoc('SO_DECQTE');
if (fDecPrix=0) then SetParamsoc('SO_DECPRIX', 2);
if (fDecQte=0) then SetParamsoc('SO_DECQTE', 2);

end;

// initialisation de la table PARPIECE
procedure InitParPiece;
begin
// Modification de GPP_IMPIMMEDIATE='-' dans la table PARPIECE   pour GPP_NATUREPIECEG="AFF" et "PAF"
    ExecuteSQL ('UPDATE PARPIECE SET GPP_IMPIMMEDIATE="-", GPP_ENCOURS="-" WHERE GPP_NATUREPIECEG="AFF"');
    ExecuteSQL ('UPDATE PARPIECE SET GPP_IMPIMMEDIATE="-" WHERE GPP_NATUREPIECEG="PAF"');
      // mcd 07/02/03  pour les natures APR, il faut comme pour les AVC, ne pas permettre les acompte. (NB OK dans socref > 595)
    ExecuteSQL ('UPDATE PARPIECE SET GPP_ACOMPTE="-" WHERE GPP_NATUREPIECEG="APR"'); 

    ExecuteSQL ('UPDATE PARPIECE SET GPP_GEREECHEANCE="SAN",GPP_VISA="-" WHERE GPP_NATUREPIECEG="FPR"');
    ExecuteSQL ('UPDATE PARPIECE SET GPP_GEREECHEANCE="SAN" WHERE GPP_NATUREPIECEG="APR"');
      //mcd 16/06/03 on met le champs "masquernature" à faux sur toutes les natures non gérée en GI
    ExecuteSql ('UPDATE PARPIECE SET gpp_masquernature="X" where gpp_naturepieceg not in ("AFF","PAF","AVC","APR","FAC","FPR","FRE","FSI","FF","AF")');
      //mcd 15/07/03 on gère le contexte pour ne voir que les natures voulues
    ExecuteSql('UPDATE parpiece set gpp_contextes="AUT;" where gpp_naturepieceg  not in ("AFF","PAF","FF","AF","FAC","AVC","FPR","FRE","APR")');
    ExecuteSql('UPDATE parpiece set gpp_contextes="AFF;" where gpp_naturepieceg  in ("AFF","PAF","FF","AF","FAC","AVC","FPR","FRE","APR")');
end;

// initialisation de la table BLOCAGEAFFAIRE
procedure InitBlocageAffaire;
var  OB , OB_DETAIL : TOB;
begin
  // Création de l'enregistrement
OB := TOB.Create('un blocage',nil,-1);

   OB_DETAIL := TOB.Create('BLOCAGEAFFAIRE',OB,-1);
   OB_DETAIL.PutValue('ABA_NUMBLOCAGE',1);
   OB_DETAIL.PutValue('ABA_EVENEMENTAFF','MAF');
   OB_DETAIL.PutValue('ABA_TYPEBLOCAGE','DDB');
   OB_DETAIL.PutValue('ABA_ETATAFFAIRE','');
   OB_DETAIL.PutValue('ABA_ALERTE','X');
   OB_DETAIL.PutValue('ABA_OPERATEUR','INF');
   OB_DETAIL.PutValue('ABA_GROUPE','');
   OB_DETAIL.PutValue('ABA_LIBELLE','Inférieure à la Date début, Modification affaire, Tous groupes');


   OB_DETAIL := TOB.Create('BLOCAGEAFFAIRE',OB,-1);
   OB_DETAIL.PutValue('ABA_NUMBLOCAGE',2);
   OB_DETAIL.PutValue('ABA_EVENEMENTAFF','SAT');
   OB_DETAIL.PutValue('ABA_TYPEBLOCAGE','DDB');
   OB_DETAIL.PutValue('ABA_ETATAFFAIRE','');
   OB_DETAIL.PutValue('ABA_ALERTE','X');
   OB_DETAIL.PutValue('ABA_OPERATEUR','INF');
   OB_DETAIL.PutValue('ABA_GROUPE','');
   OB_DETAIL.PutValue('ABA_LIBELLE','Inférieure à la Date début, Saisie activité, Tous groupes');

   OB_DETAIL := TOB.Create('BLOCAGEAFFAIRE',OB,-1);
   OB_DETAIL.PutValue('ABA_NUMBLOCAGE',3);
   OB_DETAIL.PutValue('ABA_EVENEMENTAFF','SAH');
   OB_DETAIL.PutValue('ABA_TYPEBLOCAGE','DDB');
   OB_DETAIL.PutValue('ABA_ETATAFFAIRE','');
   OB_DETAIL.PutValue('ABA_ALERTE','X');
   OB_DETAIL.PutValue('ABA_OPERATEUR','INF');
   OB_DETAIL.PutValue('ABA_GROUPE','');
   OB_DETAIL.PutValue('ABA_LIBELLE','Inférieure à la Date début, Saisie achats, Tous groupes');

   OB_DETAIL := TOB.Create('BLOCAGEAFFAIRE',OB,-1);
   OB_DETAIL.PutValue('ABA_NUMBLOCAGE',4);
   OB_DETAIL.PutValue('ABA_EVENEMENTAFF','FAC');
   OB_DETAIL.PutValue('ABA_TYPEBLOCAGE','DDB');
   OB_DETAIL.PutValue('ABA_ETATAFFAIRE','');
   OB_DETAIL.PutValue('ABA_ALERTE','X');
   OB_DETAIL.PutValue('ABA_OPERATEUR','INF');
   OB_DETAIL.PutValue('ABA_GROUPE','');
   OB_DETAIL.PutValue('ABA_LIBELLE','Inférieure à la Date début, facturation, Tous groupes');


   OB_DETAIL := TOB.Create('BLOCAGEAFFAIRE',OB,-1);
   OB_DETAIL.PutValue('ABA_NUMBLOCAGE',5);
   OB_DETAIL.PutValue('ABA_EVENEMENTAFF','MAF');
   OB_DETAIL.PutValue('ABA_TYPEBLOCAGE','DAF');
   OB_DETAIL.PutValue('ABA_ETATAFFAIRE','');
   OB_DETAIL.PutValue('ABA_ALERTE','X');
   OB_DETAIL.PutValue('ABA_OPERATEUR','SUP');
   OB_DETAIL.PutValue('ABA_GROUPE','');
   OB_DETAIL.PutValue('ABA_LIBELLE','Supérieure à la Date fin, Modification affaire, Tous groupes');

   OB_DETAIL := TOB.Create('BLOCAGEAFFAIRE',OB,-1);
   OB_DETAIL.PutValue('ABA_NUMBLOCAGE',6);
   OB_DETAIL.PutValue('ABA_EVENEMENTAFF','SAT');
   OB_DETAIL.PutValue('ABA_TYPEBLOCAGE','DAF');
   OB_DETAIL.PutValue('ABA_ETATAFFAIRE','');
   OB_DETAIL.PutValue('ABA_ALERTE','X');
   OB_DETAIL.PutValue('ABA_OPERATEUR','SUP');
   OB_DETAIL.PutValue('ABA_GROUPE','');
   OB_DETAIL.PutValue('ABA_LIBELLE','Supérieure à la Date fin, Saisie activité, Tous groupes');

   OB_DETAIL := TOB.Create('BLOCAGEAFFAIRE',OB,-1);
   OB_DETAIL.PutValue('ABA_NUMBLOCAGE',7);
   OB_DETAIL.PutValue('ABA_EVENEMENTAFF','SAH');
   OB_DETAIL.PutValue('ABA_TYPEBLOCAGE','DAF');
   OB_DETAIL.PutValue('ABA_ETATAFFAIRE','');
   OB_DETAIL.PutValue('ABA_ALERTE','X');
   OB_DETAIL.PutValue('ABA_OPERATEUR','SUP');
   OB_DETAIL.PutValue('ABA_GROUPE','');
   OB_DETAIL.PutValue('ABA_LIBELLE','Supérieure à la Date fin, Saisie achats, Tous groupes');

   OB_DETAIL := TOB.Create('BLOCAGEAFFAIRE',OB,-1);
   OB_DETAIL.PutValue('ABA_NUMBLOCAGE',8);
   OB_DETAIL.PutValue('ABA_EVENEMENTAFF','FAC');
   OB_DETAIL.PutValue('ABA_TYPEBLOCAGE','DAF');
   OB_DETAIL.PutValue('ABA_ETATAFFAIRE','');
   OB_DETAIL.PutValue('ABA_ALERTE','X');
   OB_DETAIL.PutValue('ABA_OPERATEUR','SUP');
   OB_DETAIL.PutValue('ABA_GROUPE','');
   OB_DETAIL.PutValue('ABA_LIBELLE','Supérieure à la Date fin, facturation, Tous groupes');

if (OB.Detail.Count<>0) then
   OB.InsertOrUpdateDB(True);
OB.Free;
end;

procedure InitZonesLibres;               // PL le 11/12/01 V571
begin
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "CT%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "CM%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "CD%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "CR%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "CB%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "CC%"');

ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "RT%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "RB%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "RC%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "RD%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "RM%"');

ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "MT%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "MB%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "MC%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "MD%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "MM%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "MR%"');
// mcd 19/07/02 ajout pour zone tâches
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "TT%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "TC%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "TD%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "TM%"');
//mcd 24/09/03 ajout table article, piece..
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "ET%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "EM%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "ED%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "EC%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "EB%"');

ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "AT%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "AM%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "AD%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "AC%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "AB%"');

ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "PT%"');

ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "BT%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "BM%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "BD%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "BC%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "BB%"');

ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "FT%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "FM%"');
ExecuteSQL ('UPDATE CHOIXCOD SET CC_LIBELLE=".-", CC_ABREGE=".-" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "FD%"');

AvertirTable('GCZONELIBRE');
end;

// initialisation de la table TXCPTTVA
procedure InitTVA;
var
lgCptes : integer;
sBourrage:string;
StrBase, StrCptVente:string;
i:integer;
OB , OB_DETAIL : TOB;
begin
StrBase:=''; StrCptVente:='';
lgCptes := GetParamsoc('SO_LGCPTEGEN');
sBourrage := GetParamsoc('SO_BOURREGEN');
if sBourrage='' then sBourrage:='0';

if (lgCptes >= 5) then
    begin
    for i:=1 to lgCptes-5 do StrBase := StrBase + sBourrage;
    StrCptVente := '44571' + StrBase;
    end;

  // Création de l'enregistrement
  OB := TOB.Create('un TxCptTva',nil,-1);

  OB_DETAIL := TOB.Create('TXCPTTVA',OB,-1);
  OB_DETAIL.PutValue('TV_TVAOUTPF','TX1');
  OB_DETAIL.PutValue('TV_CODETAUX','NOR');
  OB_DETAIL.PutValue('TV_REGIME','FRA');
  OB_DETAIL.PutValue('TV_TAUXACH',0);
  OB_DETAIL.PutValue('TV_CPTEACH','');
  OB_DETAIL.PutValue('TV_TAUXVTE',19.6);
  OB_DETAIL.PutValue('TV_CPTEVTE',StrCptVente);
  OB_DETAIL.PutValue('TV_ENCAISVTE',StrCptVente);

  OB_DETAIL := TOB.Create('TXCPTTVA',OB,-1);
  OB_DETAIL.PutValue('TV_TVAOUTPF','TX1');
  OB_DETAIL.PutValue('TV_CODETAUX','RED');
  OB_DETAIL.PutValue('TV_REGIME','FRA');
  OB_DETAIL.PutValue('TV_TAUXACH',0);
  OB_DETAIL.PutValue('TV_CPTEACH','');
  OB_DETAIL.PutValue('TV_CPTEVTE',StrCptVente);
  OB_DETAIL.PutValue('TV_ENCAISVTE',StrCptVente);

  OB_DETAIL := TOB.Create('TXCPTTVA',OB,-1);
  OB_DETAIL.PutValue('TV_TVAOUTPF','TX1');
  OB_DETAIL.PutValue('TV_CODETAUX','EXO');
  OB_DETAIL.PutValue('TV_REGIME','FRA');

  OB.InsertOrUpdateDB(True);
  OB.Free;

// Modification de (TV_TVAOUTPF="TX1" AND TV_CODETAUX="NOR" AND TV_REGIME="FRA") dans la table TXCPTTVA
(*  Q := OpenSQLCom ('SELECT * FROM TXCPTTVA WHERE TV_TVAOUTPF="TX1" AND TV_CODETAUX="NOR" AND TV_REGIME="FRA"',True) ;
  try
  if not Q.EOF then
    ExecuteSQL ('UPDATE TXCPTTVA SET TV_TAUXACH=0, TV_CPTEACH="", TV_TAUXVTE=19.6, TV_CPTEVTE="44571000", TV_ENCAISVTE="44571000"  WHERE TV_TVAOUTPF="TX1" AND TV_CODETAUX="NOR" AND TV_REGIME="FRA"');
  finally
  FermeSQLCom(Q);
  end;

// Modification de (TV_TVAOUTPF="TX1" AND TV_CODETAUX="RED" AND TV_REGIME="FRA") dans la table TXCPTTVA
  Q := OpenSQLCom ('SELECT * FROM TXCPTTVA WHERE TV_TVAOUTPF="TX1" AND TV_CODETAUX="RED" AND TV_REGIME="FRA"',True) ;
  try
  if not Q.EOF then
    ExecuteSQL ('UPDATE TXCPTTVA SET TV_TAUXACH=0, TV_CPTEACH="", TV_CPTEVTE="44571000", TV_ENCAISVTE="44571000"  WHERE TV_TVAOUTPF="TX1" AND TV_CODETAUX="RED" AND TV_REGIME="FRA"');
  finally
  FermeSQLCom(Q);
  end;
*)
end;

// initialisation de la table MODEREGL
procedure InitMODEREGL;
var  OB , OB_DETAIL : TOB;
begin
//  MessageAlerte ('Début Création table MODEREGL');
  // Création de l'enregistrement
  OB := TOB.Create('un ModeRegl',nil,-1);
  OB_DETAIL := TOB.Create('MODEREGL',OB,-1);
  OB_DETAIL.PutValue('MR_MODEREGLE','002');
  OB_DETAIL.PutValue('MR_LIBELLE','CHEQUE');
  OB_DETAIL.PutValue('MR_ABREGE','CHEQUE');
  OB_DETAIL.PutValue('MR_APARTIRDE','ECR');
  OB_DETAIL.PutValue('MR_PLUSJOUR',5);
  OB_DETAIL.PutValue('MR_ARRONDIJOUR','PAS');
  OB_DETAIL.PutValue('MR_NOMBREECHEANCE',1);
  OB_DETAIL.PutValue('MR_SEPAREPAR','QUI');
  //mcd 15/09/03 OB_DETAIL.PutValue('MR_MONTANTMIN',999999);
  OB_DETAIL.PutValue('MR_MONTANTMIN',0);
  OB_DETAIL.PutValue('MR_REMPLACEMIN','002');
  OB_DETAIL.PutValue('MR_MP1','CHQ');
  OB_DETAIL.PutValue('MR_TAUX1',100);
  OB.InsertOrUpdateDB(True);
  OB.Free;

AvertirTable('TTMODEREGLE');
//MessageAlerte ('Fin Création table MODEREGL');

SetParamsoc('SO_GCMODEREGLEDEFAUT', '002');
end;


(*******************************************************************************
* méthodes d'initialisation des saisies (grisage, visibilité, init des champs)
* sur les différents écrans de l'assistant
*******************************************************************************)
// initialisation des paramètres appréciation
procedure TFAssistGI.InitApprec(bInit : boolean);
begin
if bInit then
    begin
    SetParamSoc('SO_AFGESTIONAPPRECIATION','-');
    SetParamSoc('SO_AFAPPAVECBM','-');
    SetParamSoc('SO_AFAPPPOINT','-');
    SetParamSoc('SO_AFAPPCUTOFF','-');
    SetParamSoc('SO_AFAPPPRES','');
    SetParamSoc('SO_AFAPPFOUR','');
    SetParamSoc('SO_AFAPPFRAIS','');
    end;
end;
// Initialisation de la saisie de l'entite juridique
procedure TFAssistGI.InitEntiteJuridique(bInit : boolean);
begin
  if not bInit then
    begin
    if GetParamSoc ('SO_ETABLISCPTA')=True then
        bEntiteJur.ImageIndex := 1
    else
        bEntiteJur.ImageIndex := 0;
    end
  else
    begin
    SetParamSoc ('SO_ETABLISCPTA',false);
    // PL le 09/12/02 : init de l'option multidepots par établissement à faux
    SetParamsoc('SO_GCMULTIDEPOTS', false);
    end;
end;

// initialisation de la saisie des parametres ressource
procedure TFAssistGI.InitRessource(bInit : boolean);
var
LibTab:string;
begin
if bInit then
    begin
    SetParamSoc ('SO_AFPRESTATIONRES', '');
    SetParamSoc ('SO_AFRESCALCULPR', False);
    end;

//LibTab:=RechDom('AFCARESTABLE','AL1',False);
LibTab:=RechDom('GCZONELIBRE','RT1',False);
if (AnsiUppercase(Trim(LibTab))=AnsiUppercase(TraduitGA('groupe de ressources'))) then
    begin
    CheckBox2.Checked:=true;
    end
else
if (AnsiUppercase(Trim(LibTab))<>'TABLE LIBRE 1') and (LibTab<>'') then
    begin
    SBCollab.Enabled:=true;
    end
else
    begin
    SBCollab.Enabled:=false;
    end;
Autre0.text := LibTab;

//LibTab:=RechDom('AFCARESTABLE','AL2',False);
LibTab:=RechDom('GCZONELIBRE','RT2',False);
if (AnsiUppercase(Trim(LibTab))<>'TABLE LIBRE 2') and (LibTab<>'') then
    begin
    SBAutre1.Enabled:=true;
    end
else
    begin
    SBAutre1.Enabled:=false;
    end;
Autre1.text := LibTab;

LibTab:=RechDom('GCZONELIBRE','RT3',False);
if (AnsiUppercase(Trim(LibTab))<>'TABLE LIBRE 3') and (LibTab<>'') then
    begin
    SBAutre2.Enabled:=true;
    end
else
    begin
    SBAutre2.Enabled:=false;
    end;
Autre2.text := LibTab;

end;

// initialisation de la saisie des parametres activite
procedure TFAssistGI.InitActivite(bInit : boolean);
begin
// Init de la date de fin de l'activité
if bInit then
    begin
    SetParamsoc('SO_AFDATEFINACT', iDate2099);
    SetParamSoc('SO_AFVALOACTPR','RES');
    SetParamSoc('SO_AFVALOACTPV','RES');
      // mcd 04/04/2002
    SetParamSoc('SO_AFPROPOSACT','-');
    SetParamSoc('SO_AFTYPEHEUREACT','-');
    SetParamSoc('SO_AFRECHCLIMISSAUTO','X');
    SetParamSoc('SO_TYPEDATEFINACTIVITE','DAF');
    SetParamSoc('SO_AFTYPEVALOACT','001');
    end;
end;

// initialisation de la saisie des parametres affaire
procedure TFAssistGI.InitAffaire(bInit : boolean);
begin
// Gestion du format du code affaire
if bInit then
    begin
    CBFormatExercice.Value:=GetParamSoc ('SO_AFFORMATEXER');
    CBFormatExercice.OnChange:=CBFormatExerciceChange;
    end
else
    begin
    CBFormatExercice.visible:=false;
    LblFormatExercice.visible:=false;
    end;
CBFormatExercice.Enabled := Not ExisteAffaires;

if bInit then
    begin
    // Init de la gestion des apporteurs
    SetParamsoc('SO_AFGESTIONCOM', False);
    // Init eclatement des factures par ressource
    SetParamsoc('SO_AFFACTPARRES', 'SAN');
    // Gestion des saisies affaire
    SetParamsoc('SO_AFGERESSAFFAIRE', False);
    // Elements supplémentaires dans profil facture pour gestion de la facturation aux temps passés
    SetParamsoc('SO_AFTRTFTP', False);
    // mise otpion libelle mission mcd 23/12/02
    SetParamSoc('SO_AFALIMDESCAUTO','Mission $libPartie1$');
    // Insertion des blocages affaire par defaut
    InitBlocageAffaire;
    end;

SO_AFGESTIONCOM.Checked:=GetParamsoc('SO_AFGESTIONCOM');
CSO_AFFACTPARRES.Value:=GetParamsoc('SO_AFFACTPARRES');

// On rafraichit le lien entre la combo et la tablette pour que celle-ci soit mise à jour en temps réel
// NE PAS SUPPRIMER !
//self.CSO_AFFACTPARRES.Refresh;
//self.CSO_AFFACTPARRES.datatype := '';
//self.CSO_AFFACTPARRES.datatype := 'AFFACTECLATASSIST';
//self.CSO_AFFACTPARRES.Value := GetParamsoc('SO_AFFACTPARRES');
///////////////////
end;

// initialisation de la saisie des parametres client
procedure TFAssistGI.InitClient(bInit : boolean);
var
LibTab:string;
begin
// Valeurs par defaut
if bInit then
    begin        
    SetParamsoc('SO_GCNUMTIERSAUTO', false);
    SetParamsoc('SO_GCDEFFACTUREHT', True);
    SetParamsoc('SO_GCTIERSPAYS', 'FRA');
    SetParamsoc('SO_REGIMEDEFAUT', 'FRA');
    SetParamsoc('SO_GCMODEREGLEDEFAUT', '002');
    SetParamsoc('SO_GCPIECEADRESSE', 'X');    // mcd 06/06/2003 mise à vrai gestion adresse
    end;

//LibTab:=Uppercase(Trim(RechDom('GCCATTIERSTABLE','AL1',False)));
LibTab:=AnsiUppercase(Trim(RechDom('GCZONELIBRE','CT1',False)));
if (LibTab='BUREAU') then CBBureau.Checked:=true;

//LibTab:=Uppercase(Trim(RechDom('GCCATTIERSTABLE','AL2',False)));
LibTab:=AnsiUppercase(Trim(RechDom('GCZONELIBRE','CT2',False)));
if (LibTab='AGENCE') then CBAgence.Checked:=true;

// Ressources
//LibTab:=Uppercase(Trim(RechDom('AFCATIERSTABLE','AL1',False)));
LibTab:=AnsiUppercase(Trim(RechDom('GCZONELIBRE','CR1',False)));
if (LibTab='ASSOCIé') then CBAssocie.Checked:=true;

//LibTab:=Uppercase(Trim(RechDom('AFCATIERSTABLE','AL2',False)));
LibTab:=AnsiUppercase(Trim(RechDom('GCZONELIBRE','CR2',False)));
if (LibTab='CHEF DE GROUPE') then CBChefGroupe.Checked:=true;

//LibTab:=Uppercase(Trim(RechDom('AFCATIERSTABLE','AL3',False)));
LibTab:=AnsiUppercase(Trim(RechDom('GCZONELIBRE','CR3',False)));
if (LibTab='RESPONSABLE') then CBResponsable.Checked:=true;

//LibTab:=Uppercase(Trim(RechDom('GCCATTIERSTABLE','CL1',False)));
LibTab:=AnsiUppercase(Trim(RechDom('GCZONELIBRE','CD1',False)));
if bInit and ((LibTab='') or (LibTab='DATE LIBRE 1')) then
    begin
    ModifierTablette('CC', 'ZLI', 'CD1', 'Date entrée cabinet', 3);
//    AvertirTable('GCCATTIERSTABLE');
    AvertirTable('GCZONELIBRE');
    end;

CBModeRegle.Value:= GetParamSoc('SO_GCMODEREGLEDEFAUT');
end;

// initialisation de la saisie des préférences
procedure TFAssistGI.InitPreferences(bInit : boolean);
begin
if bInit then
    begin
    SetParamSoc('SO_AFACOMPTE', 'ACOMPTE');
    SetParamSoc('SO_AFFGENERAUTO','FOR');
    SetParamSoc('SO_AFFRECONDUCTION','TAC');
    SetParamSoc('SO_AFREPRISEACTIV','NON');
    SetParamSoc('SO_AFPROFILGENER', 'DEF');
    end;
end;

// initialisation des dates
procedure TFAssistGI.InitDates(bInit : boolean);
var
Jour, Mois, Annee : word;
dDebExCiv, dFinExCiv : TDateTime;
begin
// mc d04/04/2002 decodedate(Date, Annee, Mois, Jour);
decodedate(V_PGI.DateEntree, Annee, Mois, Jour);
dDebExCiv := EncodeDate(Annee,1,1);
dFinExCiv := EncodeDate(Annee, 12, 31);

if bInit then
    begin
    SetParamSoc('SO_AFDATEDEB35', dDebExCiv);
    SetParamSoc('SO_AFDATEFIN35', dFinExCiv);
    SetParamSoc('SO_AFDATEDEBCAB', dDebExCiv);
    SetParamSoc('SO_AFDATEFINCAB', dFinExCiv);
    SetParamSoc('SO_DATECUTOFF', idate1900);
    end;
end;

end.

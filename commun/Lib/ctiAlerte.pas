unit ctiAlerte;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, StdCtrls, ctiInterface, ctiConst, TrayIcon, Buttons, Menus, HPop97,
  ComCtrls, utob,
{$IFNDEF BUREAU}
  utom,hctrls,
{$ENDIF}
  ImgList,extctrls
  ;
const
{$IFDEF BUREAU}
     cct_Annuaire = 0;
     cct_RepPerso = 1;
{$ELSE}
     cct_Tiers = 0;
     fiche_tiers : string = 'TIE';
     fiche_Appel : string = 'APP';
     fiche_contact : string = 'CON';
     fiche_visu360 : string = 'VIS';
{$ENDIF BUREAU}
     cImgTel                 = 0;
     cImgTelIncoming         = 1;
     cImgTelConnected        = 2;
     cImgTelHold             = 3;
     cImgTelIdle             = 4;
{$IFDEF BUREAU}
     cImgDossier             = 5;
     cImgAnnuaire            = 6;
{$ELSE}
     cImgTiers               = 5;
     cImgContact             = 6;
{$ENDIF BUREAU}
     cImgRepertPerso         = 7;
     cImgTelOutgoing         = 8;
     cImgTelInactive         = 9;

     TRAYICON_NOTIFICATION_MSG = WM_USER + 1000;

     {WS_EX_LAYERED = $80000;
     LWA_COLORKEY  = 1;
     LWA_ALPHA     = 2;}

type
	TCtiCallstate = (TCtiDisable,TCtiNonDecroche,TctiDecroche,TCtiNothingToDo);

  TFCtiAlerte = class(TForm)
    TVCalls: TTreeView;
    ICti: TImageList;
    PBoutons: TPanel;
    BRepondre: TSpeedButton;
    BTerminer: TSpeedButton;
    BTransferer: TSpeedButton;
    BConference: TSpeedButton;
    ETelephone: TEdit;
    BAppeller: TSpeedButton;
    BDelTerminee: TSpeedButton;
    PMTray: TPopupMenu;
    MDesactiver: TMenuItem;
    BRepertoire: TSpeedButton;
    PMFiches: TPopupMenu;
    MAnnuaire: TMenuItem;
    MDossiers: TMenuItem;
    MRepertPerso: TMenuItem;
    MCollaborateurs: TMenuItem;
    BLog: TSpeedButton;
    PMLog: TPopupMenu;
    MLogVoir: TMenuItem;
    MLogClear: TMenuItem;
    MLogStop: TMenuItem;
    N1: TMenuItem;
    MActiver: TMenuItem;
    BSerieAppels: TSpeedButton;
    BRechTiers: TSpeedButton;
    BActionEnCours: TSpeedButton;

    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);

    procedure BRepondreClick (Sender: TObject);
    procedure BTerminerClick (Sender: TObject);
    procedure BTransfererClick(Sender: TObject);
    procedure BDelTermineeClick(Sender: TObject);
    procedure BAppellerClick(Sender: TObject);
    procedure ETelephoneChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure TVCallsDblClick(Sender: TObject);
    procedure TrayEvent (var mMsg:TMessage); message TRAYICON_NOTIFICATION_MSG;
    procedure OnTrayTimer (Sender:TObject);
    procedure TVCallsChange(Sender: TObject; Node: TTreeNode);
    procedure FormActivate(Sender: TObject);
    procedure MDesactiverClick(Sender: TObject);
    procedure BConferenceClick(Sender: TObject);

    procedure MDossiersClick(Sender: TObject);
    procedure MAnnuaireClick(Sender: TObject);
    procedure MCollaborateursClick(Sender: TObject);
    procedure MRepertPersoClick(Sender: TObject);
    procedure BRepertoireClick(Sender: TObject);
    procedure PMFichesPopup(Sender: TObject);

    procedure MTiersClick(Sender: TObject);
    procedure BSerieTiersClick(Sender: TObject);
    procedure BActionEnCoursClick (Sender: TObject);

    procedure BLogClick(Sender: TObject);
    procedure MLogVoirClick(Sender: TObject);
    procedure MLogClearClick(Sender: TObject);
    procedure MLogStopClick(Sender: TObject);
    procedure MActiverClick(Sender: TObject);
    procedure ETelephoneKeyPress(Sender: TObject; var Key: Char);
{$IFNDEF BUREAU}
    procedure OuvreLaBonneFiche (ctiContact : Tob);
{$ENDIF BUREAU}
  protected
           m_ctiInterface   :TCtiInterface;
           //m_bActif         :boolean;
           m_bConnActif     :boolean; // $$$ JP 31/10/06: le connecteur peut ne pas être actif de temps en temps (il peut rencontrer des difficultés)

           m_ConnectedCall  :TCtiCall;
           m_DialToneCall   :TCtiCall; // $$$ JP 12/12/05

           m_HTel           :TIcon;
           m_HTelIncoming   :TIcon;
           m_HTelOutgoing   :TIcon;
           m_HTelConnected  :TIcon;
           m_HTelHold       :TIcon;
           m_HTelInactive   :TIcon;

           m_Timer          :TTimer;
           m_hSwapIcon1     :HIcon;
           m_hSwapIcon2     :HIcon;

           // $$$ JP 21/04/06 - pour faire une fenêtre réellement topmost, transparente et surtout détachée de la mainform delphi
           procedure  CreateParams (var Params: TCreateParams); override;

           procedure  CreateTrayIcon;
           procedure  DestroyTrayIcon;
           procedure  UpdateTrayIcon;

           procedure  UpdateButtons (ctiCall:TCtiCall);

           procedure  SetContactNodeStateImage (CallNode:TTreeNode; iImage:integer);
           procedure  SetCallNodeStateImage    (CallNode:TTreeNode; iImage:integer);

           function   ModalExist:boolean;

           function   Start:boolean;
           procedure  Stop;

           // Noeud lié à l'appel et noeud lié au contact de d'un noeud d'appel
           function   GetCallNode     (ctiCall:TCtiCall):TTreeNode;
           function   GetContactNode  (CallNode:TTreeNode):TTreeNode;

           // Appel lié au noeud en sélection
           function   GetSelectedCall :TCtiCall;

           // Gestionnaires d'événement CTI: événement cti, chargement des contacts et lecture nom d'un contact
           function   OnCallEvent      (ctiCall:TCtiCall; lEvent:LongInt):LongInt;
           function   OnLoadContacts   (ctiCall:TCtiCall):integer;
           function   OnUnloadContacts (ctiCall:TCtiCall):boolean;
           function   OnGetContactName (ctiCall:TCtiCall):string;

           // Icone représentant le type de contact
           function   GetContactIcon (ctiCall:TCtiCall):integer;

           function   GetStarted:boolean;

  public
  				 function   GetState        : TCtiCallstate;
           procedure  AddCallNode     (ctiCall:TCtiCall);
           procedure  UpdateCallNode  (ctiCall:TCtiCall);
           procedure  RemoveCallNode  (ctiCall:TCtiCall);

{$IFNDEF BUREAU}
           procedure  AddActionCTI    (ctiCall:TCtiCall; ctiContact : Tob ; connectEvent : boolean);
           procedure  UpdateActionCTI (ctiCall:TCtiCall);
{$ENDIF BUREAU}

           procedure  MakeCall        (strTelephone:string; strOutgoingRef:string='');
           procedure  ClearCalls;
           // $$$ JP 16/08/07: obsolète il me semble, à réactiver si nécessaire
           //function   GetCallDealing  (strCode:string; iType:integer):TCtiCall;

           procedure  ShowAlert;

           property   ctiInterface:TCtiInterface read m_ctiInterface;
           property   SelectedCall:TCtiCall      read GetSelectedCall;

           property   Started:boolean            read GetStarted;
  end;


const
     ModeleAction : String = 'MODELES D''ACTIONS';
     // $$$ JP ne pas effacer svp TSetLayeredWindowAttributes = function (hWnd: HWND; crKey: TColorRef; bAlpha: Byte; dwFlags: LongWord):LongBool; stdcall;


implementation

uses
{$IFDEF BUREAU}
    galOutil, entDp, uDossierSelect,galmenudisp,
{$ELSE}
    UtilCtiAlerte,UTofTiersCti_Mul,UtilRT,EntGC,TiersUtil,
{$ENDIF BUREAU}
{$IFDEF EAGLCLIENT}
    MaineAgl,
    menuolx,
{$ELSE}
    fe_main,
    menuolg,
{$ENDIF}
    hent1, hmsgbox, paramsoc, utilpgi
    , shellapi
    ;

{$R *.DFM}

procedure TFCtiAlerte.CreateTrayIcon;
var
   SHNotifyIcon   :TNotifyIconData;
begin
     ZeroMemory (@SHNotifyIcon, sizeof (SHNotifyIcon));
     SHNotifyIcon.cbSize           := sizeof (SHNotifyIcon);
     SHNotifyIcon.Wnd              := Handle;
     SHNotifyIcon.uID              := 0;
     if m_ctiInterface <> nil then
     begin
          if m_ctiInterface.Started = TRUE then
              SHNotifyIcon.hIcon  := m_HTel.Handle
          else
              SHNotifyIcon.hIcon  := m_HTelInactive.Handle;
          strLCopy (SHNotifyIcon.szTip, pchar (Caption), 63);
          MActiver.Visible    := FALSE;
          MDesactiver.Visible := TRUE;
     end
     else
     begin
          SHNotifyIcon.hIcon  := m_HTelInactive.Handle;
          strLCopy (SHNotifyIcon.szTip, pchar ('Interface CTI LSE indisponible'), 63);
          MActiver.Visible    := FALSE;
          MDesactiver.Visible := FALSE;
     end;

     SHNotifyIcon.uCallbackMessage := TRAYICON_NOTIFICATION_MSG;
     SHNotifyIcon.uFlags           := NIF_ICON or NIF_TIP or NIF_MESSAGE;
     Shell_NotifyIcon (NIM_ADD, @SHNotifyIcon);
end;

procedure TFCtiAlerte.DestroyTrayIcon;
var
   SHNotifyIcon   :TNotifyIconData;
begin
     ZeroMemory (@SHNotifyIcon, sizeof (SHNotifyIcon));
     SHNotifyIcon.cbSize           := sizeof (SHNotifyIcon);
     SHNotifyIcon.Wnd              := Handle;
     SHNotifyIcon.uID              := 0;
     Shell_NotifyIcon (NIM_DELETE, @SHNotifyIcon);
end;

procedure TFCtiAlerte.UpdateTrayIcon;
var
   SHNotifyIcon   :TNotifyIconData;
begin
     // Pour màj du tray icon
     ZeroMemory (@SHNotifyIcon, sizeof (SHNotifyIcon));
     SHNotifyIcon.cbSize := sizeof (SHNotifyIcon);
     SHNotifyIcon.Wnd    := Handle;
     SHNotifyIcon.uID    := 0;

     // Si inactif, on affiche l'icone d'appel terminé
     if (Started = FALSE) or (m_bConnActif = FALSE) then //(m_bActif = FALSE) or (m_bConnActif = FALSE) then
     begin
          SHNotifyIcon.hIcon := m_HTelInactive.Handle;
          if Started = FALSE then
              strLCopy (SHNotifyIcon.szTip, pchar ('CTI LSE - Interface non démarrée'), 63)
          else
              strLCopy (SHNotifyIcon.szTip, pchar ('CTI LSE - Connecteur non disponible'), 63);
          FreeAndNil (m_Timer);

          if m_ctiInterface <> nil then
          begin
               MActiver.Visible    := TRUE;
               MDesactiver.Visible := FALSE;
          end;
     end
     else
     begin
          MActiver.Visible    := FALSE;
          MDesactiver.Visible := TRUE;

          if m_ctiInterface.OfferingCount  > 0 then
          begin
               SHNotifyIcon.hIcon := m_HTelIncoming.Handle;
               strLCopy (SHNotifyIcon.szTip, pchar ('CTI LSE - Appel entrant'), 63);
               m_hSwapIcon1 := m_HTelIncoming.Handle;
               m_hSwapIcon2 := m_HTel.Handle;
               if m_Timer = nil then
               begin
                    m_Timer          := TTimer.Create (Application);
                    m_Timer.Interval := 300;
                    m_Timer.OnTimer  := OnTrayTimer;
                    m_Timer.Enabled  := TRUE;
               end;
          end
          else if (m_ctiInterface.RingingCount > 0) or (m_ctiInterface.DialToneCount > 0) or (m_ctiInterface.DialingCount > 0) then
          begin
               SHNotifyIcon.hIcon := m_HTelOutgoing.Handle;
               strLCopy (SHNotifyIcon.szTip, pchar ('CTI LSE - Appel sortant'), 63);
               m_hSwapIcon1 := m_HTelOutgoing.Handle;
               m_hSwapIcon2 := m_HTel.Handle;
               if m_Timer = nil then
               begin
                    m_Timer          := TTimer.Create (Application);
                    m_Timer.Interval := 300;
                    m_Timer.OnTimer  := OnTrayTimer;
                    m_Timer.Enabled  := TRUE;
               end;
          end
          else if m_ctiInterface.ConnectedCount > 0 then
          begin
               SHNotifyIcon.hIcon := m_HTelConnected.Handle;
               strLCopy (SHNotifyIcon.szTip, pchar ('CTI LSE - Communication en cours'), 63);
               FreeAndNil (m_Timer);
          end
          else if m_ctiInterface.HoldCount > 0 then
          begin
               SHNotifyIcon.hIcon := m_HTelHold.Handle;
               strLCopy (SHNotifyIcon.szTip, pchar ('CTI LSE - Communication en attente'), 63);
               FreeAndNil (m_Timer);
          end
          else
          begin
               SHNotifyIcon.hIcon := m_HTel.Handle;
               strLCopy (SHNotifyIcon.szTip, pchar (Caption), 63);
               FreeAndNil (m_Timer);
          end;
     end;

     // Màj tray icon du shell
     SHNotifyIcon.uFlags := NIF_ICON or NIF_TIP;
     Shell_NotifyIcon (NIM_MODIFY, @SHNotifyIcon);
end;

procedure TFCtiAlerte.TrayEvent (var mMsg:TMessage);
var
   MousePos   :TPoint;
begin
     if mMsg.LParam = WM_LBUTTONDBLCLK then
     begin
          // On montre la fenêtre d'alerte CTI
          ShowAlert;
          //PostMessage (Handle, WM_NULL, 0, 0);
     end
     else if mMsg.LParam = WM_RBUTTONDOWN then
     begin
          GetCursorPos (MousePos);
          PMTray.Popup (MousePos.X, MousePos.Y);
          //Windows.Setfocus (PMTray.WindowHandle);
          PostMessage  (Handle, WM_NULL, 0, 0);
     end;
end;

procedure TFCtiAlerte.OnTrayTimer (Sender:TObject);
var
   SHNotifyIcon   :TNotifyIconData;
begin
     m_Timer.Enabled := FALSE;

     // Màj icone (on swappe d'un icone à l'autre)
     ZeroMemory (@SHNotifyIcon, sizeof (SHNotifyIcon));
     SHNotifyIcon.cbSize := sizeof (SHNotifyIcon);
     SHNotifyIcon.Wnd    := Handle;
     SHNotifyIcon.uID    := 0;
     SHNotifyIcon.hIcon  := m_hSwapIcon2;
     SHNotifyIcon.uFlags := NIF_ICON;
     Shell_NotifyIcon (NIM_MODIFY, @SHNotifyIcon);

     // Icon swappé
     m_hSwapIcon2 := m_hSwapIcon1;
     m_hSwapIcon1 := SHNotifyIcon.hIcon;

     m_Timer.Enabled := TRUE;
end;

function TFCtiAlerte.Start:boolean;
begin
     //FMenuG.Status.Caption := 'Démarrage de l''interface CTI ...';
     if m_ctiInterface.Start = TRUE then
     begin
          Result       := TRUE;
          TVCalls.Items.Clear;
          Caption      := 'CTI LSE (' + m_ctiInterface.MonitoredLines + ')';
          m_bConnActif := TRUE;
     end
     else
     begin
          Result       := FALSE;
          Caption      := 'CTI LSE non démarré';
          m_bConnActif := FALSE;
          with TVCalls.Items.Add (nil, 'Démarrage de l''interface CTI impossible') do
          begin
               ImageIndex    := cImgTelIdle;
               StateIndex    := -1;
               SelectedIndex := cImgTelIdle;
          end;
     end;
end;

procedure TFCtiAlerte.Stop;
begin
     m_ctiInterface.Stop;

     m_ConnectedCall  := nil;
     m_DialToneCall   := nil;
     m_bConnActif     := FALSE;

     Caption := 'CTI LSE arrêté';
end;

procedure TFCtiAlerte.MActiverClick(Sender: TObject);
begin
     if Started = FALSE then //) or (m_bConnActif = FALSE) then //(m_bActif = FALSE) or (m_bConnActif = FALSE) then
        if PgiAsk ('Confirmez-vous le démarrage du CTI LSE ?') = mrYes then
           Start;

     UpdateButtons (nil);
     UpdateTrayIcon;
end;

procedure TFCtiAlerte.MDesactiverClick (Sender:TObject);
begin
     if Started = TRUE then //and (m_bConnActif = FALSE) then //(m_bActif = FALSE) or (m_bConnActif = FALSE) then
        if PgiAsk ('Confirmez-vous l''arrêt du CTI LSE ?') = mrYes then
           Stop;

     UpdateButtons (nil);
     UpdateTrayIcon;
end;

procedure TFCtiAlerte.ShowAlert;
begin
     // On passe la fenêtre en premier plan
     Show;
end;

procedure TFCtiAlerte.CreateParams (var Params: TCreateParams);
begin
     inherited;

     // On veut une fenêtre détachée de la mainform delphi, et transparente si windows 2000/XP...
     with Params do
     begin
          WndParent := GetDesktopWindow;
          // Ne pas supprimer les commentaires de cette fonction, servira plus tard...
          {if (Win32Platform = VER_PLATFORM_WIN32_NT) and (Win32MajorVersion = 5) then
             ExStyle := ExStyle or WS_EX_LAYERED;}
     end;
end;

procedure TFCtiAlerte.FormCreate (Sender:TObject);
var
   strError                   :string;
   //strStatus                  :string;
// Ne pas supprimer les commentaires de cette fonction, servira plus tard...
{   SetLayeredWindowAttributes :TSetLayeredWindowAttributes;
   User32                     :HMODULE;}
begin
     // Appels connecté et décroché pour l'instant inconnu
     m_ConnectedCall := nil;
     m_DialToneCall  := nil;

{$IFDEF BUREAU}
     BActionEnCours.visible:=false;
     BSerieAppels.visible:=false;
     BRechTiers.visible:=false;
{$ELSE}
     BRepertoire.visible:=false;
{$ENDIF}

     // Si l'interface CTI est déjà instanciée, on ne la recréer pas
     strError  := '';
     //strStatus := FMenuG.Status.Caption;
     SourisSablier;
     try
        // Initialisation interface CTI
        if m_ctiInterface = nil then
           try
              //FMenuG.Status.Caption := 'Initialisation de l''interface CTI ...';
              //Application.ProcessMessages;
{$IFDEF BUREAU}
              m_ctiInterface := CreateCtiInterface (GetParamSocSecur ('SO_CTIINTERFACE', ''), GetParamSocSecur ('SO_CTILIGNEEXT', '0'));
{$ELSE}
              m_ctiInterface := CreateCtiInterface (GetParamSocSecur ('SO_RTINTERFACECTI', ''), GetParamSocSecur ('SO_RTCTILIGNEEXT', '0'));
{$ENDIF BUREAU}
           except
                 on E:Exception do strError := E.Message;
           end;

        // Si interface CTI instanciée, on la démarre (sauf si déjà fait)
        if m_ctiInterface <> nil then
        begin
             // Si log, on affiche le bouton
             BLog.Visible := m_ctiInterface.HasLog;

             // Gestionnaires d'événements propres au bureau PGI
             m_ctiInterface.OnAppCallEvent    := OnCallEvent;
             m_ctiInterface.AppLoadContacts   := OnLoadContacts;
             m_ctiInterface.AppUnloadContacts := OnUnloadContacts; // $$$ JP 16/12/05
             m_ctiInterface.AppGetContactName := OnGetContactName;

             // Démarrage (sauf si déjà fait): on instancie la fenêtre des communications téléphoniques
             Start;
        end;
     finally
            SourisNormale;
            //FMenuG.Status.Caption := strStatus;
     end;

     // Si interface CTI non présente, on affiche l'erreur
     {if Started = FALSE then
     begin
          with TVCalls.Items.Add (nil, strError) do
          begin
               ImageIndex    := cImgTelIdle;
               StateIndex    := -1;
               SelectedIndex := cImgTelIdle;
          end;
          //m_bActif := FALSE;
     end
     else
     begin
          //m_bActif     := TRUE;
          m_bConnActif := TRUE;

          // $$$ JP 21/04/06 - rendre transparente la fenêtre à 70%
          // Ne pas supprimer les commentaires de cette fonction, servira plus tard...
          {User32 := SafeLoadLibrary ('user32.dll');
          if User32 <> 0 then
          begin
               SetLayeredWindowAttributes := GetProcAddress (User32, 'SetLayeredWindowAttributes');
               if @SetLayeredWindowAttributes <> nil then
                  SetLayeredWindowAttributes (Handle, TColorRef (clBlue), 0, LWA_COLORKEY);
               FreeLibrary(User32);
          end;}
     //end;}

     // On énumère les 4 icônes pour le shell tray
     m_HTel          := TIcon.Create;
     m_HTelIncoming  := TIcon.Create;
     m_HTelOutgoing  := TIcon.Create;
     m_HTelConnected := TIcon.Create;
     m_HTelHold      := TIcon.Create;
     m_HTelInactive  := TIcon.Create;
     ICti.GetIcon (cImgTel,          m_HTel);
     ICti.GetIcon (cImgTelIncoming,  m_HTelIncoming);
     ICti.GetIcon (cImgTelOutgoing,  m_HTelOutgoing);
     ICti.GetIcon (cImgTelConnected, m_HTelConnected);
     ICti.GetIcon (cImgTelHold,      m_HTelHold);
     ICti.GetIcon (cImgTelInactive,  m_HTelInactive);

     // Création du trayicon dans le shell
     CreateTrayIcon;
end;

procedure TFCtiAlerte.FormDestroy (Sender:TObject);
begin
     ClearCalls;

     FreeAndNil (m_ctiInterface);

     DestroyTrayIcon;
     FreeAndNil (m_HTel);
     FreeAndNil (m_HTelIncoming);
     FreeAndNil (m_HTelOutgoing);
     FreeAndNil (m_HTelConnected);
     FreeAndNil (m_HTelHold);
     FreeAndNil (m_HTelInactive);
end;

procedure TFCtiAlerte.FormActivate(Sender: TObject);
begin
     // On déplace la fenêtre en bas à droite de l'écran (vers les tray icones)
     Left  := Screen.Width  - Width  +  1;
     Top   := Screen.Height - Height - 30;
end;

function TFCtiAlerte.OnCallEvent (ctiCall:TCtiCall; lEvent:LongInt):LongInt;
begin
     Result := 0;
     if (Started = FALSE) or (ctiCall = nil) then
        exit;

     // $$$ JP 12/12/05 - si combiné décroché, il faut le savoir pour éviter de faire des traitements incompatibles
     if lEvent = ctiEvent_CallDialTone then
          m_DialToneCall := ctiCall
     else if m_DialToneCall = ctiCall then
          m_DialToneCall := nil;

     // On fait la mise à jour de la fenêtre CTI
     case lEvent of
          ctiEvent_CallNew:
          begin
               AddCallNode (ctiCall);
          end;

          ctiEvent_CallConnect:
          begin
               m_ConnectedCall := ctiCall;
{$IFNDEF BUREAU}
               if (ctiCall.Contact <> nil) and (TOB (ctiCall.Contact).GetValue ('RAC_NUMACTION') = 0) then // $$$ JP 13/08/07 ctiCall.NumAction = 0 then
                  AddActionCTI (ctiCall, TOB (ctiCall.Contact),true);
{$ENDIF BUREAU}

               UpdateCallNode(ctiCall);
          end;

          ctiEvent_CallOffering,
          ctiEvent_CallDisconnect,
          ctiEvent_CallHold,
          ctiEvent_CallConference,
          ctiEvent_CallDialTone,
          ctiEvent_CallDialing,
          ctiEvent_CallRingBack:
          begin
               if m_ConnectedCall = ctiCall then
                  m_ConnectedCall := nil;

{$IFNDEF BUREAU}
               if lEvent = ctiEvent_CallDisconnect then
               begin
                    if (ctiCall.CalledTel <> '') or (ctiCall.CallerTel <> '') then
                       UpdateActionCTI (ctiCall);
               end;
{$ENDIF BUREAU}

               UpdateCallNode (ctiCall);
          end;

          ctiEvent_CallInfoCallerId,
          ctiEvent_CallInfoCalledId,
          ctiEvent_CallInfoConnectedId,
          ctiEvent_CallInfoOrigin,
          ctiEvent_CallInfoReason:
          begin
               UpdateCallNode (ctiCall);
          end;

          ctiEvent_CallDestroy:
          begin
               if (ctiCall.Origin = ctiOrigin_Outgoing) and (ctiCall.CalledTel = '') then
                   RemoveCallNode (ctiCall)
               else
                   UpdateCallNode (ctiCall);

               // $$$ JP 12/01/06 - à chaque passage à l'état inactif, on ferme l'alerte s'il n'y a plus d'appel non idle
               if m_ctiInterface.NonIdleCount < 1 then
                  Hide;
          end;

          // $$$ JP 31/10/06
          ctiEvent_ConnectorActive:
          begin
               m_bConnActif := TRUE;

               UpdateButtons (nil);
               UpdateTrayIcon;
          end;


          ctiEvent_ConnectorInactive:
          begin
               m_bConnActif := FALSE;

               UpdateButtons (nil);
               UpdateTrayIcon;
          end;
     end;
end;

// -----------------------------------------------------
//       STRATEGIE DE RECHERCHE DU CORRESPONDANT TELEPHONIQUE
//
// On cherche TOUTES les personnes de l'annuaire ayant un lien avec le n° de téléphone du correspondant
// Ces personnes sont liés ou non à un dossier, et le n° de téléphone est celui de la personne de l'annuaire
// ou d'un contact qui lui est lié. On charge en + les contacts du répertoire personnel
// On fournit une TOB au ctiCall, identifiant le ou les correspondant(s).
// --------------------------------------------------------
function TFCtiAlerte.OnLoadContacts (ctiCall:TCtiCall):integer;
var
   TOBAnn      :TOB;
{$IFDEF BUREAU}
   TOBTel      :TOB;
{$ENDIF}
   strCleTel   :string;
   strSelect   :string;
   i           :integer;
   CallNode    :TTreeNode;
   NewNode     :TTreeNode;

   // $$$ JP 13/08/07: pour recherche bon contact lors d'un appel sortant depuis l'application
   strOutgoingField  :string;
   strOutgoingValue  :string;
   iContact          :integer;
   iNumChamp         :integer;
begin
     Result := 0;

     if (Started = FALSE) or (ctiCall = nil) then
        exit;

     // On recherche le téléphone du correspondant: celui-ci est soit l'appellant, soit l'appellé, soit le connecté (peut être différent si routage dans l'autocom distant)
     strCleTel := ctiCall.CallTel;
     if strCleTel <> '' then
     begin
          strCleTel := CleTelephone (strCleTel);

          // 1 - Sélection des personnes (dossier ou non) lié au numéro de téléphone donné
{$IFDEF BUREAU}
          {strSelect := 'SELECT ' + IntToStr (cct_Annuaire) + ' AS CTITYPE,';
          strSelect := strSelect + 'ANN_GUIDPER,ANN_NOM1,ANN_NOM2,ANN_NOMPER,ANN_TYPEPER,ANN_TIERS,'; // $$$ JP 16/08/07: code tiers important
          strSelect := strSelect + 'DOS_NODOSSIER,';
          strSelect := strSelect + 'C_GUIDPER,C_PRENOM,C_NOM,C_FONCTION FROM ANNUAIRE ';
          strSelect := strSelect + 'LEFT JOIN DOSSIER ON ANN_GUIDPER=DOS_GUIDPER ';
          strSelect := strSelect + 'LEFT JOIN CONTACT ON ANN_GUIDPER=C_GUIDPER AND C_CLETELEPHONE="' + strCleTel + '" ';
          strSelect := strSelect + 'WHERE (ANN_CLETELEPHONE="' + strCleTel + '" OR C_CLETELEPHONE="' + strCleTel + '") ';
          strSelect := strSelect + 'AND ((DOS_NODOSSIER IS NULL) OR EXISTS (SELECT 1 FROM DOSSIERGRP, USERCONF WHERE DOS_NODOSSIER=DOG_NODOSSIER ' + ' AND ( (DOG_GROUPECONF=UCO_GROUPECONF AND UCO_USER="' + V_PGI.User + '") OR DOG_GROUPECONF="")))';
          strSelect := strSelect + 'ORDER BY ANN_NOM1,ANN_NOM2';}

          // $$$ JP 16/08/07: pb sur groupe de confidentialité
          strSelect := 'SELECT DISTINCT ' + IntToStr (cct_Annuaire) + ' AS CTITYPE,';
          strSelect := strSelect + 'ANN_GUIDPER,ANN_NOM1,ANN_NOM2,ANN_NOMPER,ANN_TYPEPER,ANN_TIERS,'; // $$$ JP 16/08/07: code tiers
          strSelect := strSelect + 'DOS_NODOSSIER,';
          strSelect := strSelect + 'C_GUIDPER,C_PRENOM,C_NOM,C_FONCTION ';
          strSelect := strSelect + 'FROM ANNUAIRE ';
          strSelect := strSelect + 'LEFT JOIN DOSSIER ON ANN_GUIDPER=DOS_GUIDPER ';
          strSelect := strSelect + 'LEFT JOIN CONTACT ON ANN_GUIDPER=C_GUIDPER ';
          strSelect := strSelect + 'LEFT JOIN DOSSIERGRP ON DOS_NODOSSIER=DOG_NODOSSIER ';
          strSelect := strSelect + 'LEFT JOIN USERCONF ON (DOG_GROUPECONF=UCO_GROUPECONF AND UCO_USER="CEG") ';
          strSelect := strSelect + 'WHERE (ANN_CLETELEPHONE="' + strCleTel + '" OR ';
          strSelect := strSelect + 'ANN_CLETELEPHONE2="' + strCleTel + '" OR ';
          strSelect := strSelect + 'C_CLETELEPHONE="' + strCleTel + '") AND ';

          // soit pas de confidentialité dossier, soit le user doit être dans le groupe de confidentialité du dossier
          strSelect := strSelect + '(DOG_NODOSSIER IS NULL OR NOT (UCO_GROUPECONF IS NULL))';
{$ELSE}
          strSelect := 'SELECT ' + IntToStr (cct_Tiers) + ' AS CTITYPE, 0 AS RAC_NUMACTION,';
          strSelect := strSelect + 'T_AUXILIAIRE,T_NATUREAUXI,T_LIBELLE,T_TIERS,C_PRENOM,C_NOM,C_FONCTIONCODEE,C_NUMEROCONTACT FROM TIERS ';
          strSelect := strSelect + 'LEFT JOIN CONTACT ON C_AUXILIAIRE=T_AUXILIAIRE AND C_TYPECONTACT = "T" AND C_CLETELEPHONE="' + strCleTel + '" ';
          strSelect := strSelect + 'WHERE (T_CLETELEPHONE="' + strCleTel + '" OR C_CLETELEPHONE="' + strCleTel;
          strSelect := strSelect + '") AND ( T_NATUREAUXI= "CLI" OR T_NATUREAUXI= "PRO" )';
{$ENDIF BUREAU}

          TOBAnn    := TOB.Create ('les personnes', nil, -1);
          try
             TOBAnn.LoadDetailFromSQL (strSelect);
             for i := 0 to TOBAnn.Detail.Count-1 do
                 ctiCall.Contacts.Add (TOBAnn.Detail [i]);
          finally
                 // $$$ JP 16/12/05 - on vire la tob si aucun correspondants chargés (sinon, sera fait lors de la dest. du cticall)
                 if TOBAnn.Detail.Count = 0 then
                    FreeAndNil (TOBAnn);
          end;
          //Application.ProcessMessages; // $$$ JP 10/07/07: pour "tenter" d'éviter des problèmes de blocage de messagesqueue constaté avec pgictitapi32.dll

{$IFDEF BUREAU}
          // 2 - Sélection des contacts personnels
          strSelect := 'SELECT ' + IntToStr (cct_RepPerso) + ' AS CTITYPE,';
          strSelect := strSelect + 'YRP_CODEREP,YRP_PRENOM,YRP_NOM FROM YREPERTPERSO ';
          strSelect := strSelect + 'WHERE YRP_USER="' + V_PGI.User + '" AND YRP_CLETELEPHONE="' + strCleTel + '" ';
          strSelect := strSelect + 'ORDER BY YRP_NOM';
          TOBTel    := TOB.Create ('rep perso', nil, -1);
          try
             TOBTel.LoadDetailFromSQL (strSelect);
             for i := 0 to TOBTel.Detail.Count-1 do
                 ctiCall.Contacts.Add (TOBTel.Detail [i]);
          finally
                 // $$$ JP 16/12/05 - on vire la tob si aucun correspondants chargés (sinon, sera fait lors de la dest. du cticall)
                 if TOBTel.Detail.Count = 0 then
                    FreeAndNil (TOBTel);
          end;
{$ENDIF BUREAU}

          // Par défaut, si une seule identité, on la sélectionne
          CallNode := GetCallNode (ctiCall);
          if CallNode <> nil then
          begin
               // Il faut supprimer les anciennes identifications (peut être fait plus d'une fois)
               CallNode.DeleteChildren;

               // Si une seule identité, on sélectionne forcément celle-ci
               iContact := -1;
               if ctiCall.Contacts.Count = 1 then
                    iContact := 0
               else
               begin
                    // $$$ JP 13/08/07: préparation recherche du bon contact sur appel sortant depuis l'application (reference de la forme CHAMP=VALEUR)
                    // $$$ todo: gérer plusieurs champs
                    strOutgoingField := ctiCall.OutgoingRef;
                    strOutgoingValue := '';
                    if strOutgoingField <> '' then
                    begin
                         i := Pos ('=', strOutgoingField);
                         if i > 0 then
                         begin
                              strOutgoingValue := Copy (strOutgoingField, i+1, Length (strOutgoingField)-i);
                              strOutgoingField := Copy (strOutgoingField, 1, i-1);
                         end
                         else
                              strOutgoingField := '';
                    end;

                    // On alimente le TreeView avec tous les contacts trouvés
                    for i := 0 to ctiCall.Contacts.Count-1 do
                    begin
                         ctiCall.ContactIndex := i;
                         if ctiCall.Contact <> nil then
                         begin
                              // Construction d'un noeud supplémentaire pour représenter le contact
                              NewNode := TVCalls.Items.AddChildObject (CallNode, ctiCall.ContactName, ctiCall.Contact);
                              if NewNode <> nil then
                              begin
                                   NewNode.ImageIndex    := -1;
                                   NewNode.SelectedIndex := -1;
                                   NewNode.StateIndex    := GetContactIcon (ctiCall);
                              end;
                              //begin
                                   //NewNode.ImageIndex    := GetContactIcon (ctiCall);
                                   //NewNode.SelectedIndex := NewNode.ImageIndex;
                              //end;

                              // $$$ JP 13/08/07: si référence sur appel sortant, on l'utilise pour sélectionner le bon contact
                              if (iContact = -1) and (strOutgoingField <> '') then
                              begin
                                   iNumChamp := TOB (ctiCall.Contact).GetNumChamp (strOutgoingField);
                                   if (iNumChamp <> -1) and (TOB (ctiCall.Contact).GetValeur (iNumChamp) = strOutgoingValue) then
                                      iContact := i;
                              end;
                         end;
                    end;
               end;

               // $$$ JP 13/08/07: sélection du contact par défaut
               ctiCall.ContactIndex := iContact;

               // Si pas d'enfant (pas plusieurs contacts)
               if CallNode.HasChildren = FALSE then
               //begin
                    //CallNode.ImageIndex    := -1;
                    //CallNode.SelectedIndex := -1;
                    CallNode.StateIndex    := GetContactIcon (ctiCall)
               //end
               else
               begin
                    CallNode.StateIndex := -1;
                    CallNode.Expand (FALSE);
               end;
               {begin
                    CallNode.ImageIndex    := GetContactIcon (ctiCall);
                    CallNode.SelectedIndex := CallNode.ImageIndex;
               end
               else
                   CallNode.Expand (FALSE);}
          end;
     end;
end;

function TFCtiAlerte.OnUnloadContacts (ctiCall:TCtiCall):boolean;
var
   TOBParent    :TOB;
   TOBNewParent :TOB;
   i            :integer;
begin
     // On considère que les correspondants vont être supprimés correctement et en totalité
     Result := TRUE;

     // On supprime les TOB parentes des TOB stockées dans la liste des correspondants du cticall
     // Suppose que les TOB stockées sont triées par TOB parentes, ce qui est logiquement le cas
     with ctiCall do
     begin
          TOBParent := nil;
          for i := 0 to Contacts.Count - 1 do
          begin
               TOBNewParent := TOB (Contacts [i]).Parent;

               // Sur rupture, on peut supprimer la TOB parente
               if (TOBParent <> nil) and (TOBNewParent <> TOBParent) then
                  TOBParent.Free;
               TOBParent := TOBNewParent;
          end;
          TOBParent.Free;
     end;
end;

function TFCtiAlerte.OnGetContactName (ctiCall:TCtiCall):string;
begin
     Result := '';
     if (Started = FALSE) or (ctiCall = nil) then
        exit;

     // On créer le libellé de l'identité spécifiée
     if ctiCall.Contact <> nil then
        with TOB (ctiCall.Contact) do
        begin
             case GetInteger ('CTITYPE') of
{$IFDEF BUREAU}
                  cct_Annuaire:
                  begin
                       // Si un contact est détecté, il doit apparaitre en priorité
                       if GetString ('C_GUIDPER') <> '' then
                       begin
                            Result := Trim (GetString ('C_PRENOM') + ' ' + GetString ('C_NOM'));

                            // $$$ JP 04/04/06
                            if GetString ('C_FONCTION') <> '' then
                               Result := Result + ' (' + GetString ('C_FONCTION') + ')';
                            Result := Result + ' - ';
                       end;

                       // Si dossier ou pas
                       if GetString ('DOS_NODOSSIER') <> '' then
                       begin
                            Result := Result + GetString ('DOS_NODOSSIER') + ' ';
                            Result := Result + Trim (GetString ('ANN_NOM2') + ' ' + GetString ('ANN_NOM1'));
                            // MD 23/01/06 - suite à gestion multi-groupe avec DOSSIERGRP, que fait-on ?
                            // if Trim (GetString ('DOS_GROUPECONF')) <> '' then
                            //   Result := Result + ' [' + Trim (GetString ('DOS_GROUPECONF')) + '] ';
                       end
                       else
                       begin
                            // Ajout des infos annuaires: libellé long ou court si déjà contact identifié
                            Result := Result + Trim (GetString ('ANN_NOM2') + ' ' + GetString ('ANN_NOM1'));
                            if Trim (GetString ('ANN_TYPEPER')) <> '' then
                               Result := Result + ' [' + Trim (GetString ('ANN_TYPEPER')) + ']'
                       end;
                  end;

                  cct_RepPerso:
                     Result := Trim (GetString ('YRP_PRENOM') + ' ' + GetString ('YRP_NOM'));
{$ELSE}
                  cct_Tiers:
                  begin
                       // Si un contact est détecté, il doit apparaitre en priorité
                       if GetInteger ('C_NUMEROCONTACT') <> 0 then
                         begin
                         Result := Trim (GetString ('C_PRENOM') + ' ' + GetString ('C_NOM'));
                         if GetString ('C_FONCTIONCODEE') <> '' then
                           Result := Result + ' (' + Rechdom('TTFONCTION',GetString('C_FONCTIONCODEE'),false) + ')';
                         Result := Result + ' - '+Rechdom('TTNATTIERS',GetString('T_NATUREAUXI'),FALSE)+' '+GetString('T_LIBELLE');
                         end
                       else
                         if GetString ('T_TIERS') <> '' then
                         begin
                         Result := Trim (Rechdom('TTNATTIERS',GetString('T_NATUREAUXI'),FALSE)+' '+
                           GetString ('T_TIERS') + ' ' + GetString ('T_LIBELLE'));
                         end;
                  end;
{$ENDIF BUREAU}
             end;
        end;
end;

function TFCtiAlerte.GetContactIcon (ctiCall:TCtiCall):integer;
begin
     Result := -1;
     if (ctiCall <> nil) and (ctiCall.Contact <> nil) then
        with TOB (ctiCall.Contact) do
        begin
             case GetInteger ('CTITYPE') of
{$IFDEF BUREAU}
                  cct_Annuaire:
                     if GetString ('DOS_NODOSSIER') <> '' then
                         Result := cImgDossier
                     else
                         Result := cImgAnnuaire;

                  cct_RepPerso:
                     Result := cImgRepertPerso;
{$ELSE}
                  cct_Tiers:
                     if GetInteger ('C_NUMEROCONTACT') <> 0 then
                         Result := cImgContact
                     else
                         Result := cImgTiers;
{$ENDIF BUREAU}
             else
                 Result := cImgTel;
             end;
        end;
end;

function TFCtiAlerte.GetStarted:boolean;
begin
     Result := (m_ctiInterface <> nil) and (m_ctiInterface.Started);
end;

// $$$ JP 16/08/07: on renvoie bien le noeud lié au TCtiCall (on remonte sur la racine du noeud en sélection)
function TFCtiAlerte.GetSelectedCall:TCtiCall;
var
   CallNode  :TTreeNode;
begin
     Result := nil;

     // $$$ JP 16/08/07: l'appel est référencé au noeud racine
     CallNode := TVCalls.Selected;
     if CallNode <> nil then
     begin
          while CallNode.Parent <> nil do
                CallNode := CallNode.Parent;
          Result := TCtiCall (CallNode.Data);
     end;

     //if TVCalls.Selected <> nil then
     //   Result := TCtiCall (TVCalls.Selected.Data)
end;

procedure TFCtiAlerte.UpdateButtons (ctiCall:TCtiCall);
begin
     // $$$ JP 15/01/07: si interface ou connecteur non actif, aucun bouton n'est disponible à part l'accès aux répertoires
     if (Started = FALSE) or (m_bConnActif = FALSE) then
     begin
          BRepondre.Enabled    := FALSE;
          BTerminer.Enabled    := FALSE;
          BTransferer.Enabled  := FALSE;
          BConference.Enabled  := FALSE;
          BAppeller.Enabled    := FALSE;
          BDelTerminee.Enabled := FALSE;
          ETelephone.Enabled   := FALSE;
     end
     else
     begin
          // $$$ JP 15/01/07: saisie d'un numéro de téléphone autorisé si interface et connecteurs actifs
          ETelephone.Enabled := TRUE;

          // SI un appel est spécifié, il conditionne l'activation des boutons de commande (sur cet appel)
          if ctiCall <> nil then // $$$ JP 15/01/07 and (m_bActif = TRUE) and (m_bConnActif = TRUE)then
          begin
               with ctiCall do
               begin
                    BRepondre.Enabled   := (m_ctiInterface.bCanAnswer = TRUE)     and ((State = ctiState_Offering) or (State = ctiState_Accepted));
                    BTerminer.Enabled   := (m_ctiInterface.bCanDisconnect = TRUE) and ((State = ctiState_Offering) or (State = ctiState_Proceeding) or (State = ctiState_Ringback) or (State = ctiState_Connected));
                    BTransferer.Enabled := (m_ctiInterface.bCanTransfer = TRUE)   and ((State = ctiState_OnHold)   and (m_ConnectedCall <> nil));
                    BConference.Enabled := (m_ctiInterface.bCanConference = TRUE) and ((State = ctiState_OnHold)   and (m_ConnectedCall <> nil));
{$IFNDEF BUREAU}
                    BActionEnCours.Enabled := ((State = ctiState_Offering) or (State = ctiState_Proceeding) or (State = ctiState_Ringback) or (State = ctiState_Connected));
{$ENDIF BUREAU}
               end;
          end
          else
          begin
               BRepondre.Enabled    := FALSE;
               BTerminer.Enabled    := FALSE;
               BTransferer.Enabled  := FALSE;
               BConference.Enabled  := FALSE;
{$IFNDEF BUREAU}
               BActionEnCours.Enabled := FALSE;
{$ENDIF BUREAU}
          end;

          // On peut appeler si le connecteur l'autorise, qu'il y a un n° de téléphone saisi et que le combiné ne soit pas décroché
          BAppeller.Enabled  := (m_ctiInterface.bCanMakeCall = TRUE) and (ETelephone.Text <> '') and (m_DialToneCall = nil);

          // Bouton de vidage de la liste: uniquement s'il y a quelque chose à vider
          BDelTerminee.Enabled := TVCalls.Items.Count > 0; //m_bActif = TRUE) and (m_bConnActif = TRUE) and (TVCalls.Items.Count > 0);
     end;
end;

function TFCtiAlerte.ModalExist:boolean;
var
   i   :integer;
begin
     Result := FALSE;
     for i := 0 to Screen.Formcount-1 do
         if (fsModal in Screen.Forms [i].FormState) and (Screen.Forms[i].Name <> Name) then
         begin
              Result := TRUE;
              exit;
         end;
end;

procedure TFCtiAlerte.AddCallNode (ctiCall:TCtiCall);
var
   NewNode  :TTreeNode;
begin
     NewNode := TVCalls.Items.AddChildObjectFirst (nil, ctiCall.FaceName, ctiCall);
     if NewNode <> nil then
     begin
          NewNode.ImageIndex    := -1;
          NewNode.SelectedIndex := -1;
          NewNode.StateIndex    := cImgTel; //cImgTelIncoming;
          {NewNode.ImageIndex    := cImgTelIncoming;
          NewNode.SelectedIndex := cImgTelIncoming;
          NewNode.StateIndex    := -1;}

          NewNode.Selected      := TRUE;

          UpdateButtons (ctiCall);
          UpdateTrayIcon;

          // $$$ JP 11/01/06 - désormais, il faut voir la fenêtre dès lors qu'un appel est créer (sortant, entrant, ...)
          ShowAlert;
     end;
end;

procedure TFCtiAlerte.SetContactNodeStateImage (CallNode:TTreeNode; iImage:integer);
var
   ContactNode :TTreeNode;
begin
     ContactNode := GetContactNode (CallNode);
     if ContactNode <> nil then
     begin
          ContactNode.ImageIndex    := iImage;
          ContactNode.SelectedIndex := iImage;
     end;
end;

procedure TFCtiAlerte.SetCallNodeStateImage (CallNode:TTreeNode; iImage:integer);
begin
     CallNode.ImageIndex    := iImage;
     CallNode.SelectedIndex := iImage;

     // Màj du noeud contact en cours
     SetContactNodeStateImage (CallNode, iImage);
end;

procedure TFCtiAlerte.UpdateCallNode (ctiCall:TCtiCall);
var
   CallNode    :TTreeNode;
begin
     CallNode := GetCallNode (ctiCall);
     if CallNode <> nil then
     begin
          // Toujours le texte du noeud lié à l'appel
          CallNode.Text := ctiCall.FaceName;

          case ctiCall.State of
               ctiState_Offering:
               begin
                    // $$$ JP 16/08/07: si un parent, on est sur un contact parmi plusieurs: il faut mettre à jour noeud contact ET noeud parent (appel)
                    SetCallNodeStateImage (CallNode, cImgTelIncoming);
                    {if ctiCall.Contact <> nil then
                        CallNode.StateIndex := cImgTelIncoming
                    else
                    begin
                         CallNode.ImageIndex    := cImgTelIncoming;
                         CallNode.SelectedIndex := cImgTelIncoming;
                    end;}

                    CallNode.Selected := TRUE;
               end;

               ctiState_Dialtone,
               ctiState_Dialing,
               ctiState_Proceeding,
               ctiState_Ringback:
               begin
                    // $$$ JP 16/08/07: si un parent, on est sur un contact parmi plusieurs: il faut mettre à jour noeud contact ET noeud parent (appel)
                    SetCallNodeStateImage (CallNode, cImgTelOutgoing);
                    {if ctiCall.Contact <> nil then
                        CallNode.StateIndex := cImgTelOutgoing
                    else
                    begin
                         CallNode.ImageIndex    := cImgTelOutgoing;
                         CallNode.SelectedIndex := cImgTelOutgoing;
                    end;}

                    CallNode.Selected := TRUE;
               end;

               ctiState_Connected:
               begin
                    // $$$ JP 16/08/07: si un parent, on est sur un contact parmi plusieurs: il faut mettre à jour noeud contact ET noeud parent (appel)
                    SetCallNodeStateImage (CallNode, cImgTelConnected);
                    {if ctiCall.Contact <> nil then
                        CallNode.StateIndex := cImgTelConnected
                    else
                    begin
                        CallNode.ImageIndex    := cImgTelConnected;
                        CallNode.SelectedIndex := cImgTelConnected;
                    end;}

                    CallNode.Selected := TRUE;

// $$$ JP 13/08/07: pas dans cette fonction, ici que pour le treeview
//{$IFNDEF BUREAU}
                    //if ( cticall.NumAction = 0 ) then
                    //   AddActionCTI (ctiCall,tob(ctiCall.Contact))
//{$ENDIF BUREAU}
               end;

               ctiState_Disconnected,
               ctiState_Idle:
               begin
                    // $$$ JP 16/08/07: si un parent, on est sur un contact parmi plusieurs: il faut mettre à jour noeud contact ET noeud parent (appel)
                    SetCallNodeStateImage (CallNode, cImgTelIdle);
                    CallNode.Collapse (FALSE);
                    {if ctiCall.Contact <> nil then
                       CallNode.StateIndex := cImgTelIdle
                    else
                    begin
                       CallNode.ImageIndex    := cImgTelIdle;
                       CallNode.SelectedIndex := cImgTelIdle;

                       CallNode.Collapse (FALSE);
                    end;}

// $$$ JP 13/08/07: pas dans cette fonction, ici que pour le treeview
//{$IFNDEF BUREAU}
//                  if ctiCall.state=ctiState_Disconnected then
//                    begin
//                      if (ctiCall.CalledTel <> '') or (ctiCall.CallerTel <> '') then
//                        UpdateActionCTI (ctiCall);
//                      cticall.NumAction := 0;
//                    end;
//{$ENDIF BUREAU}
               end;

               // $$$ JP 06/11/06: sur occupé, ce n'est pas "déconnecté": peut repasser en état "ringing" ou "connected"
               ctiState_Busy:
               begin
                    // $$$ JP 16/08/07: si un parent, on est sur un contact parmi plusieurs: il faut mettre à jour noeud contact ET noeud parent (appel)
                    SetCallNodeStateImage (CallNode, cImgTelInactive);
                    {CallNode.ImageIndex    := cImgTelInactive;
                    CallNode.SelectedIndex := cImgTelInactive;}
               end;

               ctiState_OnHold:
               begin
                    // $$$ JP 16/08/07: si un parent, on est sur un contact parmi plusieurs: il faut mettre à jour noeud contact ET noeud parent (appel)
                    SetCallNodeStateImage (CallNode, cImgTelHold);
                    {if ctiCall.Contact <> nil then
                       CallNode.StateIndex := cImgTelHold
                    else
                    begin
                       CallNode.ImageIndex    := cImgTelHold;
                       CallNode.SelectedIndex := cImgTelHold;
                    end;}
               end;
          else
              SetCallNodeStateImage (CallNode, cImgTel);
              {if ctiCall.Contact <> nil then
                 CallNode.StateIndex := cImgTel
              else
              begin
                   CallNode.ImageIndex    := cImgTel;
                   CallNode.SelectedIndex := cImgTel;
              end;}
          end;

          UpdateButtons (ctiCall);
          UpdateTrayIcon;
     end;
end;

procedure TFCtiAlerte.RemoveCallNode (ctiCall:TCtiCall);
var
   CallNode   :TTreeNode;
begin
     CallNode := GetCallNode (ctiCall);
     if CallNode <> nil then
     begin
          CallNode.Delete;

          UpdateButtons (ctiCall);
          UpdateTrayIcon;
     end;
end;

procedure TFCtiAlerte.MakeCall (strTelephone:string; strOutgoingRef:string='');
begin
     if (m_ctiInterface <> nil) and (m_DialToneCall = nil) then
       m_ctiInterface.CallMake (strTelephone, strOutgoingRef);
end;

function TFCtiAlerte.GetContactNode (CallNode:TTreeNode):TTreeNode;
var
   ctiCall      :TCtiCall;
   ContactNode  :TTreeNode;
begin
     // Par défaut, pas de noeud trouvé
     Result := nil;

     // Recherche du noeud contact lié à l'appel (du noeud d'appel)
     ctiCall := TCtiCall (CallNode.Data);
     if (ctiCall <> nil) and (ctiCall.Contact <> nil) then
     begin
          // $$$ JP 16/08/07: le callnode est désormais le noeud associé au contact sélectionné (si contact non trouvé, noeud de l'appel)
          ContactNode := CallNode.GetFirstChild;
          while (ContactNode <> nil) and (Result = nil) do
          begin
               // Si contact trouvé, le noeud renvoyé est le noeud lié au contact
               if ContactNode.Data = ctiCall.Contact then
                   Result := ContactNode
               else
                   ContactNode := ContactNode.GetNextSibling;
          end;
     end;
end;

function TFCtiAlerte.GetCallNode (ctiCall:TCtiCall):TTreeNode;
var
   CallNode     :TTreeNode;
begin
     // Par défaut, pas de noeud trouvé
     Result := nil;

     // Recherche du noeud d'appel
     CallNode := TVCalls.Items [0];
     while (CallNode <> nil) and (Result = nil) do
     begin
          if TCtiCall (CallNode.Data) = ctiCall then
              Result := CallNode
          else
              CallNode := CallNode.GetNextSibling;
     end;
end;

// $$$ JP 16/08/07: obsolète il me semble, à réactiver si nécessaire
//function TFCtiAlerte.GetCallDealing (strCode:string; iType:integer):TCtiCall;
//var
//   i, j   :integer;
//begin
//     for i := 0 to m_ctiInterface.Calls.Count-1 do
//     begin
//          Result := TCtiCall (m_ctiInterface.Calls [i]);
//          for j := Result.Contacts.Count-1 downto 0 do
//          begin
//               if TOB (Result.Contacts [j]).GetInteger ('CTITYPE') = iType then
//               begin
//                    case iType of
//{$IFDEF BUREAU}
//                         cct_Annuaire:
//                            if TOB (Result.Contacts [j]).GetString ('ANN_GUIDPER') = strCode then
//                               exit;
//
//                         //cct_RepPerso:
//                         //   ; // $$$ todo
//{$ELSE}
//                         cct_Tiers:
//                            if TOB (Result.Contacts [j]).GetString ('T_AUXILIAIRE') = strCode then
//                               exit;
//{$ENDIF BUREAU}
//                    end;
//               end;
//          end;
//     end;
//
//     Result := nil;
//end;

procedure TFCtiAlerte.ClearCalls;
begin
     UpdateTrayIcon;
end;

procedure TFCtiAlerte.BRepondreClick (Sender:TObject);
begin
     if m_ctiInterface <> nil then
        m_ctiInterface.CallConnect (SelectedCall);
end;

procedure TFCtiAlerte.BTerminerClick (Sender:TObject);
begin
     if m_ctiInterface <> nil then
        m_ctiInterface.CallDisconnect (SelectedCall);
end;

procedure TFCtiAlerte.BTransfererClick (Sender: TObject);
begin
     if m_ctiInterface <> nil then
        m_ctiInterface.CallTransfer (m_ConnectedCall, SelectedCall);
end;

procedure TFCtiAlerte.BConferenceClick(Sender: TObject);
begin
     if m_ctiInterface <> nil then
        m_ctiInterface.CallConference (m_ConnectedCall, SelectedCall);
end;

procedure TFCtiAlerte.BAppellerClick(Sender: TObject);
begin
     if m_ctiInterface <> nil then
        m_ctiInterface.CallMake (ETelephone.Text);
end;

procedure TFCtiAlerte.BDelTermineeClick (Sender:TObject);
var
   CallNode :TTreeNode;
   NextNode :TTreeNode;
   ctiCall  :TCtiCall;
begin
     if TVCalls.Items.Count > 0 then
     begin
          CallNode := TVCalls.Items [0];
          while CallNode <> nil do
          begin
               // Tout de suite connaitre le noeud suivant
               NextNode := CallNode.GetNextSibling;

               // Si l'appel associé au noeud est "idle", on supprime le noeud
               ctiCall := TCtiCall (CallNode.Data);
               if (ctiCall <> nil) and (ctiCall.State = ctiState_Idle) then
                  CallNode.Delete;

               // Noeud suivant
               CallNode := NextNode;
          end;
     end;

     // On remet à jour les boutons avec le nouvel appel sélectionné
     UpdateButtons (SelectedCall);
     UpdateTrayIcon;
end;

procedure TFCtiAlerte.ETelephoneChange (Sender: TObject);
begin
     UpdateButtons (SelectedCall);
end;

// $$$ JP 15/01/07: limiter la saisie à des chiffres (numéro de téléphone)
procedure TFCtiAlerte.ETelephoneKeyPress(Sender: TObject; var Key: Char);
begin
     if (Key > #31) and ((Key < '0') or (Key > '9')) then
        Key := #0;
end;

procedure TFCtiAlerte.FormKeyDown (Sender:TObject; var Key:Word; Shift:TShiftState);
begin
     if Key = VK_ESCAPE then
        Hide;
end;

procedure TFCtiAlerte.TVCallsDblClick (Sender:TObject);
var
   ctiCall       :TCtiCall;
   ctiContact    :TOB;
   {$IFDEF BUREAU}strGuidPer :string; {$ENDIF BUREAU}
begin
     // Si une modale déjà en cours, on ne peut rien faire
     if ModalExist = TRUE then
        exit;

     // Appel lié au noeud en sélection
     ctiCall := GetSelectedCall;
     if ctiCall = nil then
        exit;

     // Si on se on se trouve sur un noeud "contact", on le sélectionne dans l'appel
     if TVCalls.Selected.Parent <> nil then
          ctiContact := TOB (TVCalls.Selected.Data)
     else if TVCalls.Selected.HasChildren = FALSE then
          ctiContact := TOB (ctiCall.Contact)
     else
          ctiContact := nil;

     // Selon l'objet associé au noeud (appel, contact, ...)
     {ctiContact := nil;
     ctiCall    := nil;
     case TVCalls.Selected.Level of
          // Référence sur un cticall
          0:
          begin
               ctiCall := TCtiCall (TVCalls.Selected.Data);
               if ctiCall <> nil then
                  ctiContact := TOB (ctiCall.Contact);
          end;

          // Référence sur un cticontact (une tob)
          1:
          begin
               ctiContact := TOB (TVCalls.Selected.Data);
          end;
     end;}

     if ctiContact <> nil then
     begin
          // $$$ JP 16/08/07 il devient le contact en cours de l'appel, mais seulement si appel toujours vivant
          if (ctiCall.Contact <> ctiContact) and (ctiCall.State <> ctiState_Idle) then
          begin
               // Noeud de l'ancien contact ne doit plus avoir l'état en cours de l'appel
               SetContactNodeStateImage (GetCallNode (ctiCall), -1);

               // Nouveau contact
               ctiCall.Contact := ctiContact;
               UpdateCallNode (ctiCall);
          end;

{$IFNDEF BUREAU}
          { mng si pas d'action crée (plusieurs contacts pour le meme numéro), on la crée si oui
            uniquement si state = connecté }
          if ( ctxGrc in V_PGI.PGIContexte ) then
          begin
               // $$$ JP 16/08/07 if TVCalls.Selected.Level = 1 then
               //    ctiCall := TCtiCall (TVCalls.Selected.Parent.Data);
               if {(ctiCall <> nil) and }(ctiCall.state = ctiState_Connected) and (ctiCall.Origin = ctiOrigin_Incoming)
               and (GetParamsocSecur ('SO_RTCTIGENACTENTOK',false)) and (ctiContact.GetValue ('RAC_NUMACTION') = 0) then // $$$ JP 13/08/07 and ( cticall.NumAction = 0 ) then
               begin
                    if PgiAsk('Voulez-vous créer l''action CTI pour ce tiers ?','Génération action CTI') = mrYes then
                       AddActionCTI (ctiCall, ctiContact,true);
               end;
          end;
{$ENDIF BUREAU}

          Hide;
          Application.Restore;
          Application.BringToFront;

          case ctiContact.GetInteger ('CTITYPE') of
{$IFDEF BUREAU}
               cct_Annuaire:
               begin
                    strGuidPer := ctiContact.GetString ('ANN_GUIDPER');
                    if ctiContact.GetString ('DOS_NODOSSIER') <> '' then
                        RetourDossierClient (strGuidPer, TRUE)
                    else
                        AGLLanceFiche ('YY','ANNUAIRE', strGuidPer, strGuidPer, 'ACTION=CONSULTATION');
               end;

               cct_RepPerso:
               begin
                    AGLLanceFiche ('YY', 'YYREPERTPERSO', '', IntToStr (ctiContact.GetInteger ('YRP_CODEREP')) + ';' + V_PGI.User, 'ACTION=CONSULTATION');
               end;
{$ELSE}
             cct_Tiers:
             begin
                  if ( ctxGrc in V_PGI.PGIContexte ) then
                     V_PGI.ZoomOLE := true;
                  OuvreLaBonneFiche (ctiContact);
                  V_PGI.ZoomOLE := false;
             end;
{$ENDIF BUREAU}
          end;
          ShowAlert;
     end;
end;

{$IFNDEF BUREAU}
procedure TFCtiAlerte.OuvreLaBonneFiche (ctiContact : Tob);
var stArg,stAppelant: string;
		ctiCall    : TCtiCall;
    stContact : string;
begin
	 stContact := '';
   if Assigned(TVCalls.Selected) then
     begin
     if TVCalls.Selected.Level = 1 then
       begin
       ctiCall := TCtiCall (TVCalls.Selected.Parent.Data);
       end
     else
       begin
       ctiCall := TCtiCall (TVCalls.Selected.Data);
       end;
     if Assigned(ctiCall) and (ctiCall.CallerTel <> '') then stAppelant:= ';APPELANT='+ctiCall.CallerTel;
     end;
  // SO_RTCTICHOIXFICHE : TIE, CON, VIS
  stArg:='ACTION=MODIFICATION;';
  if (ctxGRC in V_PGI.PGIContexte) then
    if (RTDroitModiftiers(ctiContact.GetString('T_TIERS'))=False) then stArg:= 'ACTION=CONSULTATION;';

  if GetParamSocSecur ('SO_RTCTICHOIXFICHE', '') = fiche_contact then
    begin
    if ctiContact.GetString ('C_NUMEROCONTACT') <> '0' then
      AGLLanceFiche('YY','YYCONTACT','T;'+ctiContact.GetString ('T_AUXILIAIRE'),ctiContact.GetString ('C_NUMEROCONTACT'),'ACTION=MODIFICATION')
    else
      if (ctiContact.GetString ('T_NATUREAUXI') = 'CLI' ) or ( ctiContact.GetString ('T_NATUREAUXI') = 'PRO' ) then
        AGLLanceFiche('GC','GCTIERS','',ctiContact.GetString ('T_AUXILIAIRE'),stArg+'T_NATUREAUXI='+ctiContact.GetString ('T_NATUREAUXI')) ;
    end
  else
    if GetParamSocSecur ('SO_RTCTICHOIXFICHE', '') = fiche_tiers then
      begin
    	if ctiContact.GetString ('C_NUMEROCONTACT') <> '0' then stContact := ';CONTACT='+ctiContact.GetString ('C_NUMEROCONTACT');
      if (ctiContact.GetString ('T_NATUREAUXI') = 'CLI' ) or ( ctiContact.GetString ('T_NATUREAUXI') = 'PRO' ) then
        AGLLanceFiche('GC','GCTIERS','',ctiContact.GetString ('T_AUXILIAIRE'),stArg+'T_NATUREAUXI='+ctiContact.GetString ('T_NATUREAUXI')+stContact+stAppelant) ;
      end
    else if GetParamSocSecur ('SO_RTCTICHOIXFICHE', '') = Fiche_Appel then
      begin
  		stArg:='ACTION=CREATION;';
      if (ctiContact.GetString ('T_NATUREAUXI') = 'CLI' ) or ( ctiContact.GetString ('T_NATUREAUXI') = 'PRO' ) then
        AGLLanceFiche('BTP','BTAPPELINT','','',stArg+'T_NATUREAUXI='+ctiContact.GetString ('T_NATUREAUXI')) ;
      end
    else
      if GetParamSocSecur ('SO_RTCTICHOIXFICHE', '') = fiche_visu360 then
        AGLLanceFiche('RT','RTTIERS360','','','AUXILIAIRE='+ctiContact.GetString ('T_AUXILIAIRE')+';TIERS='+TiersAuxiliaire(ctiContact.GetString ('T_AUXILIAIRE'),true));
end;
{$ENDIF BUREAU}

procedure TFCtiAlerte.TVCallsChange (Sender:TObject; Node:TTreeNode);
begin
     UpdateButtons (SelectedCall);
end;

procedure TFCtiAlerte.MDossiersClick (Sender:TObject);
begin
{$IFDEF BUREAU}
     Hide;
     RetourSelectionDossier;
     ShowAlert;
{$ENDIF}
end;

procedure TFCtiAlerte.MAnnuaireClick (Sender:TObject);
begin
{$IFDEF BUREAU}
     Hide;
     RetourFicheAnnuaire;
     ShowAlert;
{$ENDIF}
end;

procedure TFCtiAlerte.MCollaborateursClick (Sender:TObject);
begin
{$IFDEF BUREAU}
     Hide;
     RetourUtilisatRessMul;
     ShowAlert;
{$ENDIF}
end;

procedure TFCtiAlerte.MRepertPersoClick (Sender:TObject);
begin
{$IFDEF BUREAU}
     Hide;
     RetourRepertPersoMul;
     ShowAlert;
{$ENDIF}
end;

procedure TFCtiAlerte.BRepertoireClick (Sender:TObject);
var
   MousePos   :TPoint;
begin
     GetCursorPos (MousePos);
     PMFiches.Popup (MousePos.X, MousePos.Y);
end;


procedure TFCtiAlerte.PMFichesPopup (Sender:TObject);
begin
     // On autorise l'accès rapide à un répertoire seulement si pas de fiche modale en cours
     if ModalExist = TRUE then
     begin
          MAnnuaire.Enabled       := FALSE;
          MDossiers.Enabled       := FALSE;
          MRepertPerso.Enabled    := FALSE;
          MCollaborateurs.Enabled := FALSE;
     end
     else
     begin
          MAnnuaire.Enabled       := TRUE;
          MDossiers.Enabled       := TRUE;
          MRepertPerso.Enabled    := TRUE;
          MCollaborateurs.Enabled := TRUE;
     end;
end;

procedure TFCtiAlerte.BLogClick(Sender: TObject);
var
   MousePos   :TPoint;
begin
     if (m_ctiInterface <> nil) and (V_PGI.SAV = TRUE) then
     begin
          GetCursorPos (MousePos);
          PMLog.Popup (MousePos.X, MousePos.Y);
     end;
end;

procedure TFCtiAlerte.MLogVoirClick(Sender: TObject);
begin
     if m_ctiInterface <> nil then
        m_ctiInterface.LogWrite ('', TRUE);
end;

procedure TFCtiAlerte.MLogClearClick(Sender: TObject);
begin
     if m_ctiInterface <> nil then
        m_ctiInterface.LogClear;
end;

procedure TFCtiAlerte.MLogStopClick(Sender: TObject);
begin
     if m_ctiInterface <> nil then
     begin
          if MLogStop.Tag = 0 then
          begin
               if m_ctiInterface.LogStop ('') = TRUE then
               begin
                    MLogStop.Tag := 1;
                    MLogStop.Caption := 'Démarrer le log';
               end;
          end
          else
          begin
               if m_ctiInterface.LogStart = TRUE then
               begin
                    MLogStop.Tag := 0;
                    MLogStop.Caption := 'Arrêter le log';
               end;
          end;
     end;
end;

{$IFNDEF BUREAU}
procedure TFCtiAlerte.AddActionCTI (ctiCall:TCtiCall; ctiContact : Tob; connectEvent : boolean);
begin
     if ctxGrc in V_PGI.PGIContexte then
        RTCreatActionCTI (ctiCall, ctiContact,connectEvent);
end;

procedure TFCtiAlerte.UpdateActionCTI (ctiCall:TCtiCall);
var
   ctiContact :TOB;
begin
     if not (ctxGrc in V_PGI.PGIContexte ) or (ctiCall = nil) then
        exit;

     // $$$ JP 13/08/07: à confirmer par MNG...: sélection du contact désigné par l'élément courant dans le treeview
     ctiContact := nil;
     case TVCalls.Selected.Level of
          // Référence sur un cticall
          0: ctiContact := TOB (ctiCall.Contact);

          // Référence sur un cticontact (une tob)
          1: ctiContact := TOB (TVCalls.Selected.Data);
     end;

     // Màj ou création de l'action sur ce contact
     if (ctiContact <> nil) and
     ( ({(ctiCall.ConnectedTel <> '') and }(ctiCall.Origin = ctiOrigin_Outgoing)) or (ctiCall.Origin = ctiOrigin_Incoming) )then
        if ctiContact.GetValue ('RAC_NUMACTION') > 0 then
            RTUpdateActionCTI (ctiCall, ctiContact)
        else
            RTCreatActionCTI (ctiCall, ctiContact,false);
end;
{$ENDIF BUREAU}

procedure TFCtiAlerte.BActionEnCoursClick (Sender: TObject);
{$IFNDEF BUREAU}
var ctiCall  :TCtiCall;
    ctiContact : tob;
{$ENDIF BUREAU}
begin
{$IFNDEF BUREAU}
  if TVCalls.Selected.Level = 1 then
    begin
    ctiCall := TCtiCall (TVCalls.Selected.Parent.Data);
    ctiContact := TOB (TVCalls.Selected.Data);
    end
  else
    begin
    ctiCall := TCtiCall (TVCalls.Selected.Data);
    ctiContact := TOB (ctiCall.Contact);
    end;
  if (ctiCall <> nil) and (ctiContact <> nil) and (ctiContact.GetValue ('RAC_NUMACTION') <> 0) then // $$$ JP 13/08/07 ( cticall.NumAction <> 0 ) then
    AglLanceFiche('RT','RTACTIONS','',ctiContact.GetString('T_AUXILIAIRE')+';'+IntToStr (ctiContact.GetValue ('RAC_NUMACTION')), '') // $$$ JP 13/08/07 cticall.NumAction),'')
  else PgiInfo ('L''action est inexistante ou pas encore créée','Accès action CTI');
{$ENDIF BUREAU}
end;

procedure TFCtiAlerte.MTiersClick (Sender:TObject);
{$IFNDEF BUREAU}
var iMenu : integer;
    stAppelant,strCleTel,strSelect : String;
    ctiCall    : TCtiCall;
    ctiContact,TOBAnn : Tob;
    CallNode   : TTreeNode;
{$ENDIF}
begin
{$IFNDEF BUREAU}
   Hide;
   V_PGI.ZoomOLE := true;  { Affichage en mode modal }

   { si GRC, comportement menu 92 quelque soit l'endroit }
   iMenu:=V_PGI.MenuCourant;
      if (ctxGRC in V_PGI.PGIContexte) then
     V_PGI.MenuCourant:=92;

   { si appel en cours, on propose le numéro de l'appelant dans la fiche }
   stAppelant:='';
   ctiContact := nil;
   ctiCall := nil;
   if Assigned(TVCalls.Selected) then
     begin
     if TVCalls.Selected.Level = 1 then
       begin
       ctiCall := TCtiCall (TVCalls.Selected.Parent.Data);
       ctiContact := TOB (TVCalls.Selected.Data);
       end
     else
       begin
       ctiCall := TCtiCall (TVCalls.Selected.Data);
       ctiContact := TOB (ctiCall.Contact);
       end;
     if Assigned(ctiCall) and (ctiCall.CallerTel <> '') then stAppelant:= 'APPELANT='+ctiCall.CallerTel;
     end;
   AglLanceFiche('RT','RTPROSPECT_MUL',stAppelant,'','');
   { si pas de ctiContact (pas de tiers ou contact trouvé au départ)
     on recherche s'il existe, sans action, suite à la recherche en liste.
     Si trouvé, c'est qu'il y a eu création de tiers,
     on alimente le ctiContact et on créé éventuellement l'action }
   if Assigned(ctiCall) and not Assigned (ctiContact) then
     begin
      strCleTel := CleTelephone (stAppelant);
      strSelect := 'SELECT ' + IntToStr (cct_Tiers) + ' AS TYPECONTACT, 0 AS RAC_NUMACTION'; // $$$ JP 13/08/07: RAC_NUMACTION nécessaire
      strSelect := strSelect + 'T_AUXILIAIRE,T_NATUREAUXI,T_LIBELLE,T_TIERS,C_PRENOM,C_NOM,C_FONCTIONCODEE,C_NUMEROCONTACT FROM TIERS ';
      strSelect := strSelect + 'LEFT JOIN CONTACT ON C_AUXILIAIRE=T_AUXILIAIRE and C_TYPECONTACT = "T" AND C_CLETELEPHONE="' + strCleTel + '" ';
      strSelect := strSelect + 'WHERE (T_CLETELEPHONE="' + strCleTel + '" OR C_CLETELEPHONE="' + strCleTel + '") ';
      strSelect := strSelect + 'AND NOT EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE RAC_AUXILIAIRE=T_AUXILIAIRE)';

      TOBAnn    := TOB.Create ('les personnes', nil, -1);
      try
         TOBAnn.LoadDetailFromSQL (strSelect);
         if TOBAnn.Detail.Count <> 0 then
             ctiCall.Contacts.Add (TOBAnn.Detail[0]);
      finally
         if TOBAnn.Detail.Count = 0 then
            FreeAndNil (TOBAnn);
      end;

      CallNode := GetCallNode (ctiCall);
      if CallNode <> nil then
        begin
        ctiCall.ContactIndex   := 0;
        CallNode.ImageIndex    := GetContactIcon (ctiCall);
        CallNode.SelectedIndex := CallNode.ImageIndex;
        end;
     AddActionCTI (ctiCall,tob(ctiCall.Contact),true);
     end;
   V_PGI.ZoomOLE := false;
   V_PGI.MenuCourant:=iMenu;
   ShowAlert;
{$ENDIF BUREAU}
end;

procedure TFCtiAlerte.BSerieTiersClick (Sender:TObject);
begin
{$IFNDEF BUREAU}
   Hide;
   V_PGI.ZoomOLE := true;  //Affichage en mode modal
   RTLanceFiche_TiersCti_Mul ('RT', 'RTCTI_SERIETIERS','' , '', '');
   V_PGI.ZoomOLE := false;
   ShowAlert;
{$ENDIF BUREAU}
end;


function TFCtiAlerte.GetState : TCtiCallstate;
var CtiCall :TCtiCall;
begin
	CtiCall := GetSelectedCall;
  if (Started = FALSE) or (ctiCall = nil) then
  BEGIN
  	result := TCtiDisable;
    exit;
  end;
  if (m_ctiInterface.bCanDisconnect = TRUE) and (CtiCall.State = ctiState_Connected) then
  begin
  	result := TctiDecroche;
  end else if (m_ctiInterface.bCanAnswer = TRUE) AND ((CtiCall.State = ctiState_Offering) or (CtiCall.State = ctiState_Accepted)) then
  begin
  	result := TctiNonDecroche;
  end else
  begin
  	result := TctiNothingToDo;
  end;

end;

end.

unit ctiAlerte;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, StdCtrls, ctiInterface, ctiConst, TrayIcon, Buttons, Menus, HPop97,
  ComCtrls, utob, ImgList, extctrls;

const
     cct_Annuaire = 0;
     cct_RepPerso = 1;

     cImgTel                 = 0;
     cImgTelIncoming         = 1;
     cImgTelConnected        = 2;
     cImgTelHold             = 3;
     cImgTelIdle             = 4;
     cImgDossier             = 5;
     cImgAnnuaire            = 6;
     cImgRepertPerso         = 7;
     cImgTelOutgoing         = 8;
     cImgTelInactive         = 9;

     TRAYICON_NOTIFICATION_MSG = WM_USER + 1000;

     {WS_EX_LAYERED = $80000;
     LWA_COLORKEY  = 1;
     LWA_ALPHA     = 2;}

type
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
    procedure BLogClick(Sender: TObject);
    procedure MLogVoirClick(Sender: TObject);
    procedure MLogClearClick(Sender: TObject);
    procedure MLogStopClick(Sender: TObject);
    procedure MActiverClick(Sender: TObject);

  protected
           m_ctiInterface   :TCtiInterface;
           m_bActif         :boolean;
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

           function   ModalExist:boolean;

           // Recherche appel et noeud lié à l'appel
           function   GetCallNode     (ctiCall:TCtiCall):TTreeNode;
           function   GetSelectedCall :TCtiCall;

           // Gestionnaires d'événement CTI: événement cti, chargement des contacts et lecture nom d'un contact
           function   OnCallEvent      (ctiCall:TCtiCall; lEvent:LongInt):LongInt;
           function   OnLoadContacts   (ctiCall:TCtiCall):integer;
           function   OnUnloadContacts (ctiCall:TCtiCall):boolean;
           function   OnGetContactName (ctiCall:TCtiCall):string;

           // Icone représentant le type de contact
           function   GetContactIcon (ctiCall:TCtiCall):integer;


  public
           procedure  AddCall         (ctiCall:TCtiCall);
           procedure  UpdateCall      (ctiCall:TCtiCall);
           procedure  RemoveCall      (ctiCall:TCtiCall);
           procedure  MakeCall        (strTelephone:string);

           function   GetCallDealing  (strCode:string; iType:integer):TCtiCall;
           procedure  ClearCalls;

           procedure  ShowAlert;

           property   ctiInterface:TCtiInterface read m_ctiInterface;
           property   SelectedCall:TCtiCall      read GetSelectedCall;
  end;


  //TSetLayeredWindowAttributes = function (hWnd: HWND; crKey: TColorRef; bAlpha: Byte; dwFlags: LongWord):LongBool; stdcall;


implementation

uses
    galOutil, entDp, uDossierSelect,
{$IFDEF EAGLCLIENT}
    MaineAgl,
    menuolx,
{$ELSE}
    fe_main,
    menuolg,
{$ENDIF}
    hent1, hmsgbox, paramsoc, utilpgi
    ,galmenudisp
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
          SHNotifyIcon.hIcon  := m_HTel.Handle;
          strLCopy (SHNotifyIcon.szTip, pchar (Caption), 63);
          MActiver.Visible    := FALSE;
          MDesactiver.Visible := TRUE;
     end
     else
     begin
          SHNotifyIcon.hIcon  := m_HTelInactive.Handle;
          strLCopy (SHNotifyIcon.szTip, pchar ('CTI Cegid (indisponible)'), 63);
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
     if (m_bActif = FALSE) or (m_bConnActif = FALSE) then
     begin
          SHNotifyIcon.hIcon := m_HTelInactive.Handle;
          if m_bActif = FALSE then
              strLCopy (SHNotifyIcon.szTip, pchar ('CTI Cegid - Désactivé'), 63)
          else
              strLCopy (SHNotifyIcon.szTip, pchar ('CTI Cegid - Connecteur non disponible'), 63);
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
               strLCopy (SHNotifyIcon.szTip, pchar ('CTI Cegid - Appel entrant'), 63);
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
               strLCopy (SHNotifyIcon.szTip, pchar ('CTI Cegid - Appel sortant'), 63);
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
               strLCopy (SHNotifyIcon.szTip, pchar ('CTI Cegid - Communication en cours'), 63);
               FreeAndNil (m_Timer);
          end
          else if m_ctiInterface.HoldCount > 0 then
          begin
               SHNotifyIcon.hIcon := m_HTelHold.Handle;
               strLCopy (SHNotifyIcon.szTip, pchar ('CTI Cegid - Communication en attente'), 63);
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

procedure TFCtiAlerte.MDesactiverClick (Sender:TObject);
begin
     if m_ctiInterface = nil then
        exit;

     if (m_bActif = TRUE) and (m_bConnActif = TRUE) then
     begin
          if PgiAsk ('Confirmez-vous la désactivation du CTI ?') = mrYes then
             m_bActif := FALSE;
     end;

     UpdateButtons (nil);
     UpdateTrayIcon;
end;

procedure TFCtiAlerte.MActiverClick(Sender: TObject);
begin
     if m_ctiInterface = nil then
        exit;

     if (m_bActif = FALSE) and (m_bConnActif = TRUE) then
     begin
          if PgiAsk ('Confirmez-vous l''activation du CTI ?') = mrYes then
             m_bActif := TRUE;
     end;

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
   strStatus                  :string;
// Ne pas supprimer les commentaires de cette fonction, servira plus tard...
{   SetLayeredWindowAttributes :TSetLayeredWindowAttributes;
   User32                     :HMODULE;}
begin
     // Appels connecté et décroché pour l'instant inconnu
     m_ConnectedCall := nil;
     m_DialToneCall  := nil;

     // Si l'interface CTI est déjà instanciée, on ne la recréer pas
     strError  := '';
     strStatus := FMenuG.Status.Caption;
     SourisSablier;
     try
        // Initialisation interface CTI
        if m_ctiInterface = nil then
           try
              FMenuG.Status.Caption := 'Initialisation de l''interface CTI ...';
              Application.ProcessMessages;
              m_ctiInterface := CreateCtiInterface (GetParamSocSecur ('SO_CTIINTERFACE', ''), GetParamSocSecur ('SO_CTILIGNEEXT', '0'));
           except
                 on E:Exception do strError := E.Message;
           end;

        // Si interface CTI instanciée, on la démarre (sauf si déjà fait)
        if m_ctiInterface <> nil then
        begin
             // Gestionnaires d'événements propres au bureau PGI
             m_ctiInterface.OnAppCallEvent    := OnCallEvent;
             m_ctiInterface.AppLoadContacts   := OnLoadContacts;
             m_ctiInterface.AppUnloadContacts := OnUnloadContacts; // $$$ JP 16/12/05
             m_ctiInterface.AppGetContactName := OnGetContactName;

             // Démarrage (sauf si déjà fait): on instancie la fenêtre des communications téléphoniques
             FMenuG.Status.Caption := 'Démarrage de l''interface CTI ...';
             Application.ProcessMessages;
             if m_ctiInterface.Start = FALSE then
             begin
                  Caption  := 'CTI Cegid (indisponible)';
                  strError := 'Démarrage de l''interface CTI impossible';
                  FreeAndNil (m_ctiInterface);
             end
             else
                 Caption := 'CTI Cegid (' + m_ctiInterface.MonitoredLines + ')';
        end;
     finally
            SourisNormale;
            FMenuG.Status.Caption := strStatus;
     end;

     // Si interface CTI non présente, on affiche l'erreur
     if m_ctiInterface = nil then
     begin
          with TVCalls.Items.Add (nil, strError) do
          begin
               ImageIndex    := cImgTelIdle;
               StateIndex    := -1;
               SelectedIndex := cImgTelIdle;
          end;
          m_bActif := FALSE;
     end
     else
     begin
          m_bActif     := TRUE;
          m_bConnActif := TRUE;

          // $$$ JP 09/01/06 - Log visible qu'en mode SAV
          if V_PGI.Sav = TRUE then
          begin
               if m_ctiInterface.LogStart = TRUE then
                  BLog.Visible := TRUE;
          end;

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
     end;

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
     if (m_bActif = FALSE) or (ctiCall = nil) then
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
               AddCall (ctiCall);
          end;

          ctiEvent_CallConnect:
          begin
               m_ConnectedCall := ctiCall;
               UpdateCall (ctiCall);
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
               UpdateCall (ctiCall);
          end;

          ctiEvent_CallInfoCallerId,
          ctiEvent_CallInfoCalledId,
          ctiEvent_CallInfoConnectedId,
          ctiEvent_CallInfoOrigin,
          ctiEvent_CallInfoReason:
          begin
               UpdateCall (ctiCall);
          end;

          ctiEvent_CallDestroy:
          begin
               if (ctiCall.Origin = ctiOrigin_Outgoing) and (ctiCall.CalledTel = '') then
                   RemoveCall (ctiCall)
               else
                   UpdateCall (ctiCall);

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
   TOBTel      :TOB;
   strCleTel   :string;
   strSelect   :string;
   i           :integer;
   CallNode    :TTreeNode;
   NewNode     :TTreeNode;
begin
     Result := 0;
     if (m_bActif = FALSE) or (ctiCall = nil) then
        exit;

     // On recherche le téléphone du correspondant: celui-ci est soit l'appellant, soit l'appellé, soit le connecté (peut être différent si routage dans l'autocom distant)
     strCleTel := ctiCall.CorrespondentTel;
     if strCleTel <> '' then
     begin
          strCleTel := CleTelephone (strCleTel);

          // 1 - Sélection des personnes (dossier ou non) lié au numéro de téléphone donné
          strSelect := 'SELECT ' + IntToStr (cct_Annuaire) + ' AS TYPECONTACT,';
          strSelect := strSelect + 'ANN_GUIDPER,ANN_NOM1,ANN_NOM2,ANN_NOMPER,ANN_TYPEPER,';
          strSelect := strSelect + 'DOS_NODOSSIER,';
          strSelect := strSelect + 'C_GUIDPER,C_PRENOM,C_NOM,C_FONCTION FROM ANNUAIRE ';
          strSelect := strSelect + 'LEFT JOIN DOSSIER ON ANN_GUIDPER=DOS_GUIDPER ';
          //mcd 12/2005 strSelect := strSelect + 'LEFT JOIN ANNUINTERLOC ON ANN_CODEPER=ANI_CODEPER AND ANI_CLETELEPHONE="' + strCleTel + '" ';
          strSelect := strSelect + 'LEFT JOIN CONTACT ON ANN_GUIDPER=C_GUIDPER AND C_CLETELEPHONE="' + strCleTel + '" ';

          // $$$ JP 04/10/05 - gestion groupe de confidentialité
          // MD 23/01/06 - req modifiée suite gestion multi-groupe DOSSIERGRP
          strSelect := strSelect + 'WHERE (ANN_CLETELEPHONE="' + strCleTel + '" OR C_CLETELEPHONE="' + strCleTel + '") ';
          strSelect := strSelect + 'AND ((DOS_NODOSSIER IS NULL) OR EXISTS (SELECT 1 FROM DOSSIERGRP, USERCONF WHERE DOS_NODOSSIER=DOG_NODOSSIER ' + ' AND ( (DOG_GROUPECONF=UCO_GROUPECONF AND UCO_USER="' + V_PGI.User + '") OR DOG_GROUPECONF="")))';
          strSelect := strSelect + 'ORDER BY ANN_NOM1,ANN_NOM2';
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

          // 2 - Sélection des contacts personnels
          strSelect := 'SELECT ' + IntToStr (cct_RepPerso) + ' AS TYPECONTACT,';
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

          // Par défaut, si une seule identité, on la sélectionne
          CallNode := GetCallNode (ctiCall);
          if CallNode <> nil then
             if ctiCall.Contacts.Count = 1 then
             begin
                  ctiCall.ContactIndex   := 0;
                  CallNode.ImageIndex    := GetContactIcon (ctiCall);
                  CallNode.SelectedIndex := CallNode.ImageIndex;
             end
             else
             begin
                  // $$$ JP 31/10/06: il faut effacer les éléments déjà existant, car l'identification peut être faite plusieurs fois
                  CallNode.DeleteChildren;

                  // On alimente le TreeView avec tous les contacts trouvés
                  for i := 0 to ctiCall.Contacts.Count-1 do
                  begin
                       ctiCall.ContactIndex := i;
                       if ctiCall.Contact <> nil then
                       begin
                            NewNode := TVCalls.Items.AddChildObject (CallNode, ctiCall.ContactName, ctiCall.Contact);
                            if NewNode <> nil then
                            begin
                                 NewNode.ImageIndex    := GetContactIcon (ctiCall);
                                 NewNode.SelectedIndex := NewNode.ImageIndex;
                            end;
                       end;
                  end;

                  ctiCall.ContactIndex := -1; //SetContact (-1);
                  CallNode.Expand (FALSE);
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
     if (m_bActif = FALSE) or (ctiCall = nil) then
        exit;

     // On créer le libellé de l'identité spécifiée
     if ctiCall.Contact <> nil then
        with TOB (ctiCall.Contact) do
        begin
             case GetInteger ('TYPECONTACT') of
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
             end;
        end;
end;

function TFCtiAlerte.GetContactIcon (ctiCall:TCtiCall):integer;
begin
     Result := -1;
     if (ctiCall <> nil) and (ctiCall.Contact <> nil) then
        with TOB (ctiCall.Contact) do
        begin
             case GetInteger ('TYPECONTACT') of
                  cct_Annuaire:
                     if GetString ('DOS_NODOSSIER') <> '' then
                         Result := cImgDossier
                     else
                         Result := cImgAnnuaire;

                  cct_RepPerso:
                     Result := cImgRepertPerso;
             else
                 Result := cImgTel;
             end;
        end;
end;

function TFCtiAlerte.GetSelectedCall:TCtiCall;
begin
     Result := nil;
     if TVCalls.Selected <> nil then
        Result := TCtiCall (TVCalls.Selected.Data);
end;

procedure TFCtiAlerte.UpdateButtons (ctiCall:TCtiCall);
begin
     if (ctiCall <> nil) and (m_bActif = TRUE) and (m_bConnActif = TRUE)then
     begin
          with ctiCall do
          begin
               BRepondre.Enabled   := (m_ctiInterface.bCanAnswer = TRUE)     and ((State = ctiState_Offering) or (State = ctiState_Accepted));
               BTerminer.Enabled   := (m_ctiInterface.bCanDisconnect = TRUE) and ((State = ctiState_Offering) or (State = ctiState_Proceeding) or (State = ctiState_Ringback) or (State = ctiState_Connected));
               BTransferer.Enabled := (m_ctiInterface.bCanTransfer = TRUE)   and ((State = ctiState_OnHold)   and (m_ConnectedCall <> nil));
               BConference.Enabled := (m_ctiInterface.bCanConference = TRUE) and ((State = ctiState_OnHold)   and (m_ConnectedCall <> nil));
               BAppeller.Enabled   := (m_ctiInterface.bCanMakeCall = TRUE)   and ((ETelephone.Text <> '')     and (m_DialToneCall = nil));
          end;
     end
     else
     begin
          BRepondre.Enabled    := FALSE;
          BTerminer.Enabled    := FALSE;
          BTransferer.Enabled  := FALSE;
          BConference.Enabled  := FALSE;
          BAppeller.Enabled    := (m_ctiInterface.bCanMakeCall = TRUE) and ((m_bActif = TRUE) and (m_bConnActif = TRUE) and (ETelephone.Text <> '') and (m_DialToneCall = nil));
     end;

     // Bouton de vidage de la liste
     BDelTerminee.Enabled := (m_bActif = TRUE) and (m_bConnActif = TRUE) and (TVCalls.Items.Count > 0);
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

procedure TFCtiAlerte.AddCall (ctiCall:TCtiCall);
var
   NewNode  :TTreeNode;
begin
     NewNode := TVCalls.Items.AddChildObjectFirst (nil, ctiCall.FaceName, ctiCall);
     if NewNode <> nil then
     begin
          NewNode.ImageIndex    := cImgTelIncoming;
          NewNode.StateIndex    := -1;
          NewNode.SelectedIndex := cImgTelIncoming;
          NewNode.Selected      := TRUE;

          UpdateButtons (ctiCall);
          UpdateTrayIcon;

          // $$$ JP 11/01/06 - désormais, il faut voir la fenêtre dès lors qu'un appel est créer (sortant, entrant, ...)
          ShowAlert;
     end;
end;

procedure TFCtiAlerte.UpdateCall (ctiCall:TCtiCall);
var
   CallNode   :TTreeNode;
begin
     CallNode := GetCallNode (ctiCall);
     if CallNode <> nil then
     begin
          CallNode.Text := ctiCall.FaceName;
          case ctiCall.State of
               ctiState_Offering:
               begin
                    if ctiCall.Contact <> nil then
                        CallNode.StateIndex := cImgTelIncoming
                    else
                    begin
                         CallNode.ImageIndex    := cImgTelIncoming;
                         CallNode.SelectedIndex := cImgTelIncoming;
                    end;

                    CallNode.Selected := TRUE;
               end;

               ctiState_Dialtone,
               ctiState_Dialing,
               ctiState_Proceeding,
               ctiState_Ringback:
               begin
                    if ctiCall.Contact <> nil then
                        CallNode.StateIndex := cImgTelOutgoing
                    else
                    begin
                         CallNode.ImageIndex    := cImgTelOutgoing;
                         CallNode.SelectedIndex := cImgTelOutgoing;
                    end;

                    CallNode.Selected := TRUE;
               end;

               ctiState_Connected:
               begin
                    if ctiCall.Contact <> nil then
                        CallNode.StateIndex := cImgTelConnected
                    else
                    begin
                        CallNode.ImageIndex    := cImgTelConnected;
                        CallNode.SelectedIndex := cImgTelConnected;
                    end;

                    CallNode.Selected := TRUE;
               end;

               ctiState_Disconnected,
               ctiState_Idle:
               begin
                    if ctiCall.Contact <> nil then
                       CallNode.StateIndex := cImgTelIdle
                    else
                    begin
                       CallNode.ImageIndex    := cImgTelIdle;
                       CallNode.SelectedIndex := cImgTelIdle;

                       CallNode.Collapse (FALSE);
                    end;
               end;

               // $$$ JP 06/11/06: sur occupé, ce n'est pas "déconnecté": peut repasser en état "ringing" ou "connected"
               ctiState_Busy:
               begin
                    CallNode.ImageIndex    := cImgTelInactive;
                    CallNode.SelectedIndex := cImgTelInactive;
               end;

               ctiState_OnHold:
               begin
                    if ctiCall.Contact <> nil then
                       CallNode.StateIndex := cImgTelHold
                    else
                    begin
                       CallNode.ImageIndex    := cImgTelHold;
                       CallNode.SelectedIndex := cImgTelHold;
                    end;
               end;
          else
              if ctiCall.Contact <> nil then
                 CallNode.StateIndex := cImgTel
              else
              begin
                   CallNode.ImageIndex    := cImgTel;
                   CallNode.SelectedIndex := cImgTel;
              end;
          end;

          UpdateButtons (ctiCall);
          UpdateTrayIcon;
     end;
end;

procedure TFCtiAlerte.RemoveCall (ctiCall:TCtiCall);
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

procedure TFCtiAlerte.MakeCall (strTelephone:string);
begin
     if (m_ctiInterface <> nil) and (m_DialToneCall = nil) then
        m_ctiInterface.CallMake (strTelephone);
end;

function TFCtiAlerte.GetCallNode (ctiCall:TCtiCall):TTreeNode;
begin
     Result := TVCalls.Items [0];
     while Result <> nil do
     begin
          if TCtiCall (Result.Data) = ctiCall then
             exit;

          Result := Result.getNextSibling;
     end;

     Result := nil;
end;

function TFCtiAlerte.GetCallDealing (strCode:string; iType:integer):TCtiCall;
var
   i, j   :integer;
begin
     for i := 0 to m_ctiInterface.Calls.Count-1 do
     begin
          Result := TCtiCall (m_ctiInterface.Calls [i]);
          for j := Result.Contacts.Count-1 downto 0 do
          begin
               if TOB (Result.Contacts [j]).GetInteger ('TYPECONTACT') = iType then
               begin
                    case iType of
                         cct_Annuaire:
                            // if TOB (Result.Contacts [j]).GetInteger ('ANN_CODEPER') = StrToInt (strCode) then
                            if TOB (Result.Contacts [j]).GetString ('ANN_GUIDPER') = strCode then
                               exit;

                         cct_RepPerso:
                            ; // $$$ todo
                    end;
               end;
          end;
     end;

     Result := nil;
end;

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

procedure TFCtiAlerte.FormKeyDown (Sender:TObject; var Key:Word; Shift:TShiftState);
begin
     if Key = VK_ESCAPE then
        Hide;
end;

procedure TFCtiAlerte.TVCallsDblClick (Sender:TObject);
var
   ctiCall       :TCtiCall;
   ctiContact    :TOB;
   strGuidPer    :string;
begin
     // Si une modale déjà en cours, on ne peut rien faire
     if ModalExist = TRUE then
        exit;

     // Selon l'objet associé au noeud (appel, contact, ...)
     ctiContact := nil;
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
     end;

     if ctiContact <> nil then
     begin
          Hide;
          Application.Restore;
          Application.BringToFront;
          case ctiContact.GetInteger ('TYPECONTACT') of
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
          end;
          ShowAlert;
     end;
end;

procedure TFCtiAlerte.TVCallsChange (Sender:TObject; Node:TTreeNode);
begin
     UpdateButtons (SelectedCall);
end;

procedure TFCtiAlerte.MDossiersClick (Sender:TObject);
begin
     Hide;
     RetourSelectionDossier;
     ShowAlert;
end;

procedure TFCtiAlerte.MAnnuaireClick (Sender:TObject);
begin
     Hide;
     RetourFicheAnnuaire;
     ShowAlert;
end;

procedure TFCtiAlerte.MCollaborateursClick (Sender:TObject);
begin
     Hide;
     RetourUtilisatRessMul;
     ShowAlert;
end;

procedure TFCtiAlerte.MRepertPersoClick (Sender:TObject);
begin
     Hide;
     RetourRepertPersoMul;
     ShowAlert;
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
        m_ctiInterface.LogWrite;
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
               if m_ctiInterface.LogStop = TRUE then
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


end.

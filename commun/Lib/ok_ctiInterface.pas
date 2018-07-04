unit ctiInterface;

interface

uses
    sysutils, ctiConst, uTOB, classes, stdctrls, contnrs;

type
    TCtiInterface       = class;

    TCtiStart             = function  (sProfil:pchar; sLine:pchar):LongInt;
    TCtiStop              = procedure;
    TCtiGetCapacities     = function  :LongInt;
    TCtiGetMonitoredLines = function  :pchar;

    // $$$ JP 09/01/06 - Gestion du log CTI
    TCtiLogStart        = function  :LongInt;
    TCtiLogStop         = function  :LongInt;
    TCtiLogClear        = function  :LongInt;
    TCtiLogWrite        = function  (sFileName:pchar; iShowMode:integer):LongInt;
    // $$$

    TCtiCallMake        = function  (sTelephone:pchar):LongInt;
    TCtiCallConnect     = function  (lCall:LongInt):LongInt;
    TCtiCallDisconnect  = function  (lCall:LongInt):LongInt;
    TCtiCallRedirect    = function  (lCall:LongInt; sTelephone:pchar):LongInt;
    TCtiCallTransfer    = function  (lHeldCall:LongInt; lConnectedCall:LongInt):LongInt;
    TCtiCallConference  = function  (lHeldCall:LongInt; lConnectedCall:LongInt):LongInt;

    TCtiSetEventHandler = procedure (pCtiCallBack:TCtiEvent);

    // Classe représentant une communication. Peut être dériver pour enrichir les informations de la communication
    TCtiCall = class (TObject)
    protected
             m_ctiInterface    :TCtiInterface;  // Interface CTI parente
             m_lID             :LongInt;        // Identifiant de l'appel, spécifié par le connecteur
             m_lState          :LongInt;        // Etat de l'appel
             m_lStateDetails   :LongInt;        // Complément d'information sur l'état d'un appel (pour disconnected surtout)
             m_lOrigin         :LongInt;        // Origine de l'appel (entrant, sortant, interne ou externe, ...)
             m_Contacts        :TObjectList;    // Liste d'objet, chacun de ces objets étant une identité de l'appellant
             m_Contact         :TObject;        // L'objet représentant l'identité sélectionnée (s'il y en a une)
             m_strContactName  :string;         // Nom affichable du contact en cours (à alimenter par l'appli)
             m_CallerTel       :string;         // Numéro de téléphone de l'appellant
             m_CalledTel       :string;         // Numéro de téléphone de l'appellé
             m_ConnectedTel    :string;         // Numéro du téléphone réellement connecté (si transfert, redirection...)

             m_TimeStart       :TDateTime;      // Horodatage: début
             m_TimeEnd         :TDateTime;      // Horodatage: fin

             procedure         SetContact (iIndex:integer);

             function          GetHorodate         :string;
             function          GetCorrespondentTel (iIndex:integer):string;
             function          GetFaceName         (iIndex:integer):string;

    public
          constructor Create (ctiInterface:TCtiInterface);
          destructor  Destroy;                  override;

          procedure   UnloadContacts;

          property    ID             :LongInt     read m_lID             write m_lID;
          property    State          :LongInt     read m_lState          write m_lState;
          property    StateDetails   :LongInt     read m_lStateDetails   write m_lStateDetails;
          property    Origin         :LongInt     read m_lOrigin         write m_lOrigin;

          property    CallerTel        :string      read m_CallerTel       write m_CallerTel;
          property    CalledTel        :string      read m_CalledTel       write m_CalledTel;
          property    ConnectedTel     :string      read m_ConnectedTel    write m_ConnectedTel;

          property    CorrespondentTel         :string    index 0          read GetCorrespondentTel;
          property    CorrespondentTelFaceName :string    index 1          read GetCorrespondentTel;

          property    FaceName       :string      index 1                  read GetFaceName;
          property    ShortFaceName  :string      index 2                  read GetFaceName;

          property    Contacts       :TObjectList read  m_Contacts;
          property    Contact        :TObject     read  m_Contact;
          property    ContactIndex   :integer     write SetContact;
          property    ContactName    :string      read  m_strContactName;

          property    TimeStart      :TDateTime   read m_TimeStart         write m_TimeStart;
          property    TimeEnd        :TDateTime   read m_TimeEnd           write m_TimeEnd;
          property    Horodate       :string      read GetHorodate;
    end;

    // Gestionnaire d'événement applicatifs
    TCtiAppEvent          = function (ctiCall:TCtiCall; lEvent:LongInt):LongInt of object;

    // Gestion des identités du correspondant (il peut y avoir plusieurs identités connues pour un n° de tel)
    TCtiAppLoadContacts   = function  (ctiCall:TCtiCall):integer of object;
    TCtiAppUnloadContacts = function  (ctiCall:TCtiCall):boolean of object;
    TCtiAppGetContactName = function  (ctiCall:TCtiCall):string  of object;

    // Classe d'interface avec le connecteur CTI (dll)
    TCtiInterface = class (TObject)
    protected
          m_hModule           :LongWord;     // handle de la dll
          m_bStarted          :boolean;      // connecteur démarré ou non
          m_strInterface      :string;       // code de l'interface CTI utilisé (tablette RTINTERFACECTI)
          m_strConnector      :string;       // nom de la dll associé à cette interface
          m_strProfil         :string;       // profil de démarrage du connecteur
          m_strLine           :string;       // ligne à monitorée
          m_strPrefixe        :string;       // N° préfixe pour atteindre l'extérieur
          m_lCapacities       :LongInt;      // Possibilité de CTI: lancement d'appel, répondre, transférer...
          m_bHorodatage       :boolean;      // On gère l'horodatage ou pas

          // Liste des communications en cours
          m_Calls             :TObjectList;

          // Fonctions de bases de la dll
          m_ctiStart          :TCtiStart;                // Démarrage du connecteur CTI PGI
          m_ctiStop           :TCtiStop;                 // Arrêt du connecteur CTI PGI
          m_ctiCapacities     :TCtiGetCapacities;        // Possibilités cti du connecteur CTI PGI
          m_ctiMonitoredLines :TCtiGetMonitoredLines;    // Nom des lignes monitorées (pour l'instant une seule)

          // $$$ JP 09/01/06 - gestion du log
          m_ctiLogStart       :TCtiLogStart;
          m_ctiLogStop        :TCtiLogStop;
          m_ctiLogClear       :TCtiLogClear;
          m_ctiLogWrite       :TCtiLogWrite;
          // $$$

          // Fonctions de traitements téléphoniques
          m_ctiCallMake       :TCtiCallMake;             // fonction dll de lancement d'appel
          m_ctiCallConnect    :TCtiCallConnect;          // fonction dll d'acceptation d'un appel
          m_ctiCallDisconnect :TCtiCallDisconnect;       // fonction dll de rejet/terminaison d'un appel
          m_ctiCallRedirect   :TCtiCallRedirect;         // fonction dll de redirection d'un appel (non connecté)
          m_ctiCallTransfer   :TCtiCallTransfer;         // fonction dll de transfert d'un appel en attente sur un appel connecté
          m_ctiCallConference :TCtiCallConference;       // fonction dll de conférence à 3 avec un appel en attente et un connecté

          // Fonctions callback pour l'application sur événements CTI
          m_OnAppCallEvent    :TCtiAppEvent;

          // Fonctions de déterminations des identités du correspondant
          m_AppLoadContacts   :TCtiAppLoadContacts;
          m_AppUnloadContacts :TCtiAppUnloadContacts;
          m_AppGetContactName :TCtiAppGetContactName;

          // Nom des lignes monitorées (pour l'instant une seule)
          function    GetMonitoredLines:string;

          // Capacités du connecteur CTI
          function    GetCapacity (iIndex:integer):boolean;

          // Gestion de la liste des communications
          function    GetCtiCall     (lID:LongInt):TCtiCall;
          procedure   SetCtiCall     (lID:LongInt; Call:TCtiCall);

          // Fonctions callback de l'interface CTI, déclarées auprès du connecteur cti, appel retransmis à l'application
          function    OnCallEvent  (lCall:LongInt; lEvent:LongInt; lInfo:LongInt; strInfo:string):LongInt;

          function    GetCallCount (iIndex:integer):integer;

    public
          constructor Create (strInterface:string);
          destructor  Destroy; override;

          function    Start (bForceRestart:boolean=FALSE):boolean;
          procedure   Stop;

          // $$$ JP 09/01/06 - gestion du log
          function    LogStart:boolean;
          function    LogStop:boolean;
          function    LogClear:boolean;
          function    LogWrite:boolean;
          // $$$

          function    CallMake       (strTelephone:string):LongInt;                        // Emission d'un appel téléphonique
          function    CallConnect    (ctiCall:TCtiCall):LongInt;                           // Acceptation d'un appel téléphonique entrant
          function    CallDisconnect (ctiCall:TCtiCall):LongInt;                           // Fin d'appel (connecté ou non)
          function    CallRedirect   (ctiCall:TCtiCall; strTelephone:string):LongInt;      // Redirection d'un appel (non connecté)
          function    CallTransfer   (ConnectedCall:TCtiCall; HeldCall:TCtiCall):LongInt;  // Transfert d'un appel sur un autre
          function    CallConference (ConnectedCall:TCtiCall; HeldCall:TCtiCall):LongInt;  // Conférence à 3

          function    GetContactName (ctiCall:TCtiCall):string;                            // Récupère le nom du contact en cours de l'appel
          function    LoadContacts   (ctiCall:TCtiCall):integer;                           // Charge en mémoie les contacts liés au n° de téléphone
          function    UnloadContacts (ctiCall:TCtiCall):boolean;                           // Décharge de la mémoire les contacts chargés

          // Gestionnaire d'événements CTI applicatif,
          property    OnAppCallEvent     :TCtiAppEvent           read m_OnAppCallEvent    write m_OnAppCallEvent;
          property    AppLoadContacts    :TCtiAppLoadContacts    read m_AppLoadContacts   write m_AppLoadContacts;
          property    AppUnloadContacts  :TCtiAppUnloadContacts  read m_AppUnloadContacts write m_AppUnloadContacts;
          property    AppGetContactName  :TCtiAppGetContactName  read m_AppGetContactName write m_AppGetContactName;

          // Nom des lignes monitorées (une seule pour l'instant)
          property    MonitoredLines     :string                 read GetMonitoredLines;

          // Capacités du connecteur CTI
          property    bCanMakeCall:boolean    index ctiCaps_CallMake       read GetCapacity;
          property    bCanAnswer:boolean      index ctiCaps_CallAnswer     read GetCapacity;
          property    bCanDisconnect:boolean  index ctiCaps_CallDisconnect read GetCapacity;
          property    bCanRedirect:boolean    index ctiCaps_CallRedirect   read GetCapacity;
          property    bCanTransfer:boolean    index ctiCaps_CallTransfer   read GetCapacity;
          property    bCanConference:boolean  index ctiCaps_CallConference read GetCapacity;
          property    bCanSwapHold:boolean    index ctiCaps_CallSwap       read GetCapacity;

          // Liste de communication
          property    Call [lID:LongInt] :TCtiCall      read GetCtiCall        write SetCtiCall;
          property    Calls              :TObjectList   read m_Calls;
          property    Prefixe            :string        read m_strPrefixe      write m_strPrefixe;

          // Compteurs divers
          property    CallCount         :integer  index 0                      read GetCallCount;
          property    OfferingCount     :integer  index ctiState_Offering      read GetCallCount;
          property    ConnectedCount    :integer  index ctiState_Connected     read GetCallCount;
          property    HoldCount         :integer  index ctiState_OnHold        read GetCallCount;
          property    RingingCount      :integer  index ctiState_Ringback      read GetCallCount;
          property    DialToneCount     :integer  index ctiState_DialTone      read GetCallCount;
          property    DialingCount      :integer  index ctiState_Dialing       read GetCallCount;
          property    IdleCount         :integer  index ctiState_Idle          read GetCallCount;
          property    NonIdleCount      :integer  index -ctiState_Idle         read GetCallCount;

          // Comportement divers de l'interface
          property    Horodatage        :boolean        read m_bHorodatage     write m_bHorodatage;
    end;

    ECtiError = class (Exception);

// Créateur de l'interface CTI, en fonction du code interface (tablette RTINTERFACECTI)
function CreateCtiInterface (strInterfaceName:string; strPrefixe:string):TCtiInterface;


implementation

uses
    windows, hctrls, hent1, utilPgi, entDp,
    Registry;


function CreateCtiInterface (strInterfaceName:string; strPrefixe:string):TCtiInterface;
begin
     if strInterfaceName <> '' then
     begin
          Result := TCtiInterface.Create (strInterfaceName);
          if Result <> nil then
             Result.Prefixe := strPrefixe;
     end
     else
         Result := nil;
end;


              // -------------------------------
              //    Fonctions callback de la dll
              // -------------------------------
function DllOnCallEvent (lCall:LongInt; lEvent:LongInt; lInfo:LongInt; sInfo:pchar):LongInt;
begin
     Result := ctiStatus_Error;

{$IFDEF BUREAU}
     if VH_DP.ctiAlerte <> nil then
        Result := VH_DP.ctiAlerte.ctiInterface.OnCallEvent (lCall, lEvent, lInfo, string (sInfo));
{$ENDIF}
end;




              // --------------------------
              //       TCtiInterface
              // --------------------------
constructor TCtiInterface.Create (strInterface:string);
var
   TOBCti              :TOB;
   RegCti              :TRegistry;
   ctiSetEventHandler  :TCtiSetEventHandler;
begin
     // Code de l'interface CTI dans la table PGI
     m_strInterface := strInterface;
     m_bHorodatage  := TRUE;

     // Nom du connecteur (dll) associé, et lecture du profil de démarrage (paramètre de type chaine de car.)
     TOBCti := TOB.Create ('interface CTI', nil, -1);
     if TOBCti <> nil then
     begin
          // On détermine le profil d'utilisation du CTI en fonction de l'interface: en général le code provider
          try
             TOBCti.LoadDetailFromSQL ('SELECT CO_ABREGE,CO_LIBRE FROM COMMUN WHERE CO_TYPE="RCT" AND CO_CODE="' + m_strInterface + '"');
             if TOBCti.Detail.Count = 0 then
                  raise ECtiError.Create ('L''interface CTI n''est pas (ou mal) configurée')
             else
             begin
                  m_strConnector := TOBCti.Detail [0].GetString ('CO_ABREGE');
                  m_strProfil    := TOBCti.Detail [0].GetString ('CO_LIBRE');
             end;
          finally
                 TOBCti.Free;
          end;

          // On détermine la ligne téléphonique à écouter
          // $$$todo v7 - dans la fiche utilisateur, pas dans la registry
          m_strLine := '';
          RegCti := TRegistry.Create;
          try
             if V_PGI.WinNT = TRUE then
                 RegCti.RootKey := HKEY_CURRENT_USER
             else
                 RegCti.RootKey := HKEY_LOCAL_MACHINE;
             if RegCti.OpenKey ('\Software\Cegid\Bureau PGI S5\', FALSE) = TRUE then
             begin
                  m_strLine := RegCti.ReadString ('CTI');
                  RegCti.CloseKey;
             end;
          finally
                 FreeAndNil (RegCti);
          end;

          // Chargement de la dll (le connecteur) dans le process
          m_hModule := LoadLibrary (pchar (m_strConnector));
          if m_hModule = 0 then
             raise ECtiError.Create ('Le connecteur CTI "' + m_strConnector + '" n''est pas utilisable (absent ou défectueux)');

          // Identification des points d'entrée principaux: démarrage et fonctions téléphoniques
          @m_ctiStart          := GetProcAddress (m_hModule, pchar ('ctiStart'));
          @m_ctiStop           := GetProcAddress (m_hModule, pchar ('ctiStop'));
          @m_ctiCapacities     := GetProcAddress (m_hModule, pchar ('ctiGetCapacities'));
          @m_ctiMonitoredLines := GetProcAddress (m_hModule, pchar ('ctiGetMonitoredLines'));

          // $$$ JP 09/01/06 - Gestion du log CTI
          @m_ctiLogStart       := GetProcAddress (m_hModule, pchar ('ctiLogStart'));
          @m_ctiLogStop        := GetProcAddress (m_hModule, pchar ('ctiLogStop'));
          @m_ctiLogClear       := GetProcAddress (m_hModule, pchar ('ctiLogClear'));
          @m_ctiLogWrite       := GetProcAddress (m_hModule, pchar ('ctiLogWrite'));
          // $$$
          
          @m_ctiCallMake       := GetProcAddress (m_hModule, pchar ('ctiCallMake'));
          @m_ctiCallConnect    := GetProcAddress (m_hModule, pchar ('ctiCallConnect'));
          @m_ctiCallDisconnect := GetProcAddress (m_hModule, pchar ('ctiCallDisconnect'));
          @m_ctiCallRedirect   := GetProcAddress (m_hModule, pchar ('ctiCallRedirect'));
          @m_ctiCallTransfer   := GetProcAddress (m_hModule, pchar ('ctiCallTransfer'));
          @m_ctiCallConference := GetProcAddress (m_hModule, pchar ('ctiCallConference'));

          // Il faut au moins ctiStart/ctiStop
          if (@m_ctiStart = nil) or (@m_ctiStop = nil) or (@m_ctiCallMake = nil) then
             raise ECtiError.Create ('Le connecteur CTI n''est pas conforme');

          // Spécification au conecteur du gestionnaire d'événement CTI de l'application
          @ctiSetEventHandler := GetProcAddress (m_hModule, pchar ('ctiSetEventHandler'));
          if (@ctiSetEventHandler <> nil) then
             ctiSetEventHandler (DllOnCallEvent);

          // Création de la liste des communications (vide pour l'instant)
          m_Calls := TObjecTList.Create (TRUE);
     end
     else
         raise ECtiError.Create ('Mémoire insuffisante pour instancier l''interface CTI');
end;

destructor TCtiInterface.Destroy;
begin
     // Arrêt de l'interface
     Stop;

     // Plus besoin de la liste des communications
     FreeAndNil (m_Calls);

     // Plus besoin de la dll
     if m_hModule <> 0 then
        FreeLibrary (m_hModule);

     inherited;
end;

function TCtiInterface.Start (bForceRestart:boolean):boolean;
begin
     // On lit la registry du user, pour voir si pas des param' "cachés" (doit disparaitre à terme, à stocker dans base)

     // Si déjà démarré, on ne fait rien (sauf si force le redémarrage)
     try
        // Si démarrage forcé, on arrête s'il était déjà démarré
        if bForceRestart = TRUE then
           Stop;

        // Démarrage du connecteur
        if (m_bStarted = FALSE) then
        begin
             m_bStarted := m_ctiStart (pchar (m_strProfil), pchar (m_strLine)) = ctiStatus_OK;

             // Si connecteur démarré, on lui demande ses possibilités
             if (m_bStarted = TRUE) and (@m_ctiCapacities <> nil) then
                m_lCapacities := m_ctiCapacities;
        end;
     finally
            // On indique si interface démarrée
            Result := m_bStarted;
     end;
end;

procedure TCtiInterface.Stop;
begin
     if m_bStarted = TRUE then
     begin
          // Arrêt du connecteur
          m_ctiStop;

          // Plus de référence sur les appels CTI (suppression automatique, car liste propriétaire)
          m_Calls.Clear;

          // Cti arrêté
          m_bStarted    := FALSE;
          m_lCapacities := 0;
     end;
end;

function TCtiInterface.GetMonitoredLines:string;
begin
     Result := '';
     if @m_ctiMonitoredLines <> nil then
        Result := string (m_ctiMonitoredLines);
end;

function TCtiInterface.GetCapacity (iIndex:integer):boolean;
begin
     if m_lCapacities <> ctiStatus_Error then
         Result := (m_lCapacities and iIndex) = iIndex
     else
         Result := FALSE;
end;

function TCtiInterface.LogStart:boolean;
begin
     Result := FALSE;
     if @m_ctiLogStart <> nil then
        Result := m_ctiLogStart = ctiStatus_OK;
end;

function TCtiInterface.LogStop:boolean;
begin
     Result := FALSE;
     if @m_ctiLogStop <> nil then
        Result := m_ctiLogStop = ctiStatus_OK;
end;

function TCtiInterface.LogClear:boolean;
begin
     Result := FALSE;
     if @m_ctiLogClear <> nil then
        Result := m_ctiLogClear = ctiStatus_OK;
end;

function TCtiInterface.LogWrite:boolean;
var
   sLogFile  :array [0..1023] of char;
begin
     Result := FALSE;
     if @m_ctiLogWrite <> nil then
     begin
          // Nom de fichier temporaire
          GetTempPath     (SizeOf (sLogFile) - 1, sLogFile);
          GetTempFileName (sLogfile, 'CTI', 0, sLogFile);

          // On enregistre le log dans le fichier, et on le visualise
          Result := m_ctiLogWrite (sLogFile, SW_NORMAL) = ctiStatus_OK;
     end;
end;

function TCtiInterface.OnCallEvent (lCall:LongInt; lEvent:LongInt; lInfo:LongInt; strInfo:string):LongInt;
var
   ctiCall        :TCtiCall;
   bNeedIdentify  :boolean;
   bCanIdentify   :boolean;
begin
     Result := ctiStatus_OK;

     // Recherche de l'appel cti dans la liste déjà établie
     ctiCall := Call [lCall];

     // Traitement de l'événement CTI en fonction de sa nature
     bNeedIdentify := FALSE;
     case lEvent of
          // Création d'un appel dans le connecteur
          ctiEvent_CallNew:
          begin
               if ctiCall = nil then
               begin
                    ctiCall := TCtiCall.Create (Self);
                    if ctiCall <> nil then
                       Call [lCall] := ctiCall;
               end;
          end;

          // Suppression d'un appel dans le connecteur: on le marque comme n'existant plus (non supprimé en mémoire)
          ctiEvent_CallDestroy:
          begin
               if ctiCall <> nil then
               begin
                    // $$$ JP 12/01/06 - Si pas encore fait, on màj l'horodate de fin
                    if (Horodatage = TRUE) and (ctiCall.TimeEnd = 0) then
                       ctiCall.TimeEnd := Time;
                    ctiCall.State := ctiState_Idle;
               end;
          end;

          ctiEvent_CallOffering:
          begin
               if ctiCall <> nil then
                  ctiCall.State := ctiState_Offering;
          end;

          ctiEvent_CallConnect:
          begin
               if ctiCall <> nil then
               begin
                    ctiCall.State        := ctiState_Connected;
                    ctiCall.StateDetails := lInfo;

                    // Horodatage
                    if Horodatage = TRUE then
                       ctiCall.TimeStart := Time;
               end;
          end;

          ctiEvent_CallDisconnect:
          begin
               if ctiCall <> nil then
               begin
                    // $$$ JP 12/01/06 - Si pas encore fait, on màj l'horodate de fin
                    if (Horodatage = TRUE) and (ctiCall.TimeEnd = 0) then
                       ctiCall.TimeEnd   := Time;

                    // $$$ JP 06/11/06: si un appel sortant n'a jamais été connecté, on considère que le correspondant était injoignable (raison inconnue)
                    if (ctiCall.State <> ctiState_Connected) and (ctiCall.Origin = ctiOrigin_Outgoing) and ((lInfo = ctiStateC_None) or (lInfo = ctiStateC_DisconnectedNormal)) then
                        ctiCall.StateDetails := ctiStateC_DisconnectedNoAnswer
                    else
                        ctiCall.StateDetails := lInfo;
                    ctiCall.State := ctiState_Disconnected;
               end;
          end;

          ctiEvent_CallHold:
          begin
               if ctiCall <> nil then
                  ctiCall.State := ctiState_OnHold;
          end;

          ctiEvent_CallConference:
          begin
               if ctiCall <> nil then
                  ctiCall.State := ctiState_Conferenced;
          end;

          ctiEvent_CallDialTone:
          begin
               if ctiCall <> nil then
                  ctiCall.State := ctiState_Dialtone;
          end;

          ctiEvent_CallDialing:
          begin
               if ctiCall <> nil then
                  ctiCall.State := ctiState_Dialing;
          end;

          ctiEvent_CallRingBack:
          begin
               if ctiCall <> nil then
                  ctiCall.State := ctiState_Ringback;
          end;

          ctiEvent_CallProceeding:
          begin
               if ctiCall <> nil then
                  ctiCall.State := ctiState_Proceeding;
          end;

          ctiEvent_CallBusy:
          begin
               if ctiCall <> nil then
                  ctiCall.State := ctiState_Busy;
          end;

          ctiEvent_CallInfoCallerId:
          begin
               // $$$ JP 31/10/06: modification que si changement
               if ctiCall <> nil then
               begin
                    strInfo := Trim (strInfo);
                    if strInfo <> ctiCall.CallerTel then
                    begin
                         bNeedIdentify     := TRUE;
                         ctiCall.CallerTel := strInfo;
                    end;
               end;
          end;

          ctiEvent_CallInfoCalledId:
          begin
               // $$$ JP 31/10/06: modification que si changement
               if ctiCall <> nil then
               begin
                    strInfo := Trim (strInfo);
                    if strInfo <> ctiCall.CalledTel then
                    begin
                         bNeedIdentify     := TRUE;
                         ctiCall.CalledTel := strInfo;
                    end;
               end;
          end;

          ctiEvent_CallInfoConnectedId:
          begin
               // $$$ JP 31/10/06: modification que si changement, et nécessite identification (peut avoir changé via le PABX distant)
               if ctiCall <> nil then
               begin
                    strInfo := Trim (strInfo);
                    if strInfo <> ctiCall.ConnectedTel then
                    begin
                         bNeedIdentify        := TRUE;
                         ctiCall.ConnectedTel := strInfo;
                    end;
               end;
          end;

          ctiEvent_CallInfoOrigin:
          begin
               bNeedIdentify := TRUE;
               if ctiCall <> nil then
                  ctiCall.Origin := lInfo;
          end;

          ctiEvent_CallInfoReason:
          begin
          end;
     end;

     // Appel connu
     if ctiCall <> nil then
     begin
          // D'abord les identités du correspondant, si c'est nécessaire et possible
          bCanIdentify := (ctiCall.m_CallerTel <> '') and (ctiCall.m_CalledTel <> '') and (ctiCall.m_lOrigin <> ctiOrigin_Unknown);
          if (bNeedIdentify = TRUE) and (bCanIdentify = TRUE) then
             LoadContacts (ctiCall);

          // Transmission de l'événement à l'application
          if @m_OnAppCallEvent <> nil then
             m_OnAppCallEvent (ctiCall, lEvent);
     end
     else
         // Des événements peuvent ne pas concerner un appel (mais le connecteur, par exemple ctiEvent_ConnectorInactive)
         if @m_OnAppCallEvent <> nil then
            m_OnAppCallEvent (nil, lEvent);
end;

function TCtiInterface.CallMake (strTelephone:string):LongInt;
begin
     // Par défaut, appel non lancé
     Result := 0;

     // Si procédure d'appel définie dans la dll, on l'invoque
     if @m_ctiCallMake <> nil then
     begin
          if Length (strTelephone) > 4 then
             strTelephone := m_strPrefixe + strTelephone;
          Result := m_ctiCallMake (pchar (CleTelephone (strTelephone, FALSE)));
     end;
end;

function TCtiInterface.CallConnect (ctiCall:TCtiCall):LongInt;
begin
     // Par défaut, appel non connecté
     Result := 0;
     if ctiCall.State = ctiState_Idle then
        exit;

     // Si procédure d'appel définie dans la dll, on l'invoque
     if @m_ctiCallConnect <> nil then
        Result := m_ctiCallConnect (ctiCall.ID);
end;

function TCtiInterface.CallDisconnect (ctiCall:TCtiCall):LongInt;
begin
     // Par défaut, appel non terminée
     Result := 0;
     if ctiCall.State = ctiState_Idle then
        exit;

     // Si procédure définie dans la dll, on l'invoque
     if @m_ctiCallDisconnect <> nil then
        Result := m_ctiCallDisconnect (ctiCall.ID);
end;

function TCtiInterface.CallRedirect (ctiCall:TCtiCall; strTelephone:string):LongInt;
begin
     // Par défaut, appel non redirigée
     Result := 0;

     // Si procédure définie dans la dll, on l'invoque
     if @m_ctiCallRedirect <> nil then
        Result := m_ctiCallRedirect (ctiCall.ID, pchar (CleTelephone (strTelephone, FALSE)));
end;

function TCtiInterface.CallTransfer (ConnectedCall:TCtiCall; HeldCall:TCtiCall):LongInt;
begin
     // Par défaut, communication non transférée
     Result := 0;

     // Si procédure d'appel définie dans la dll, on l'invoque
     if (@m_ctiCallTransfer <> nil) and (ConnectedCall <> nil) and (HeldCall <> nil) then
        Result := m_ctiCallTransfer (ConnectedCall.ID, HeldCall.ID);
end;

function TCtiInterface.CallConference (ConnectedCall:TCtiCall; HeldCall:TCtiCall):LongInt;
begin
     // Par défaut, appel non mis en conférence
     Result := 0;

     // Si procédure d'appel définie dans la dll, on l'invoque
     if (@m_ctiCallConference <> nil) and (ConnectedCall <> nil) and (HeldCall <> nil) then
        Result := m_ctiCallConference (ConnectedCall.ID, HeldCall.ID);
end;

function TCtiInterface.GetContactName (ctiCall:TCtiCall):string;
begin
     Result := '';
     if (ctiCall <> nil) and (@m_AppGetContactName <> nil) then
        Result := m_AppGetContactName (ctiCall);
end;

// $$$ JP 31/10/06: on enlève les anciens contacts déjà existant (plusieurs identification peuvent être demandées)
function TCtiInterface.LoadContacts (ctiCall:TCtiCall):integer;
begin
     // PAr défaut, aucun contact
     Result := 0;

     // On enlève les contacts déjà présents
     ctiCall.UnloadContacts;

     // On charge les contacts (c'est l'appli qui sait faire)
     if @m_AppLoadContacts <> nil then
        Result := m_AppLoadContacts (ctiCall);
end;

function TCtiInterface.UnloadContacts (ctiCall:TCtiCall):boolean;
begin
     if @m_AppUnloadContacts <> nil then
         Result := m_AppUnloadContacts (ctiCall)
     else
         Result := FALSE;
end;

function TCtiInterface.GetCtiCall (lID:LongInt):TCtiCall;
var
   i           :integer;
   SearchCall  :TCtiCall;
begin
     // Par défaut, communication non trouvée
     Result := nil;

     // Recherche parmi les communications existantes
     for i := 0 to m_Calls.Count-1 do
     begin
          SearchCall := TCtiCall (m_Calls [i]);
          if SearchCall.ID = lID then
          begin
               Result := SearchCall;
               exit;
          end;
     end;
end;

procedure TCtiInterface.SetCtiCall (lID:LongInt; Call:TCtiCall);
begin
     // Si on référence un nouvel appel, on l'ajoute à la liste (pas de controle de doublon sur l'id)
     if Call <> nil then
     begin
          // On spécifié l'identifiant de l'appel
          Call.ID := lId;

          // Ajout à la liste
          m_Calls.Add (Call);
     end
     else
     begin
          // On supprime l'appel (de la liste, en mémoire)
          Call := GetCtiCall (lId);
          if Call <> nil then
             m_Calls.Remove (Call);
     end;
end;

function TCtiInterface.GetCallCount (iIndex:integer):integer;
var
   i       :integer;
   iState  :integer;
begin

     // $$$ JP 06/01/06 - pour tous les appels, on renvoie le nb de la liste
     if iIndex = 0 then
         Result := m_Calls.Count
     else
     begin
          Result := 0;
          for i := 0 to m_Calls.Count-1 do
          begin
               iState := integer (TCtiCall (m_Calls [i]).State);

               // $$$ JP 12/01/06 - pouvoir compter les appels qui NE sont PAS dans un état donné
               if iIndex > 0 then
               begin
                    if iState = iIndex then
                       Inc (Result);
               end
               else
               begin
                    if iState <> -iIndex then
                       Inc (Result);
               end;
          end;
     end;
end;




                 // -----------------------------------------
                 //      CLASSE DES APPELS CTI
                 // -----------------------------------------
constructor TCtiCall.Create (ctiInterface:TCtiInterface);
begin
     m_ctiInterface   := ctiInterface;
     m_lId            := LongInt ($FFFFFFFF);
     m_lState         := ctiState_Unknown;
     m_lStateDetails  := ctiStateC_Unknown;
     m_lOrigin        := ctiOrigin_Unknown;
     m_Contacts       := TObjectList.Create (FALSE);
     m_Contact        := nil;
     m_strContactName := '';
     m_CallerTel      := '';
     m_CalledTel      := '';
     m_ConnectedTel   := '';

     // $$$ JP 12/01/06 - horodatage
     m_TimeStart      := 0;
     m_TimeEnd        := 0;
end;

destructor TCtiCall.Destroy;
begin
     // On vide la liste des contacts
     UnloadContacts;

     // On supprime la liste des contacts
     FreeAndNil (m_Contacts);

     inherited;
end;

procedure TCtiCall.UnloadContacts;
var
   i   :integer;
begin
     // Demande à l'interface CTI de supprimer les contacts qu'elle a mis dans ce cticall. Sinon, on fait ce qu'on peut: on les vire tel quel, sans fioritures
     if m_ctiInterface.UnloadContacts (Self) = FALSE then
        for i := 0 to m_Contacts.Count - 1 do
            m_Contacts [i].Free;

     // vide la liste
     m_Contacts.Clear;
end;

procedure TCtiCall.SetContact (iIndex:integer);
begin
     // Par défaut, identité non trouvée
     m_Contact        := nil;
     m_strContactName := '';

     // Si identité trouvée, on la sélectionne
     if (iIndex >= 0) and (iIndex < m_Contacts.Count) then
     begin
          // On demande également le nom du contact
          m_Contact := m_Contacts [iIndex];
          if m_Contact <> nil then
              m_strContactName := m_ctiInterface.GetContactName (Self);
     end;
end;

function TCtiCall.GetCorrespondentTel (iIndex:integer):string;
var
   strConnectedTel  :string; // $$$ JP 31/10/06
begin
     // Par défaut, c'est le numéro connecté
     Result := Trim (m_ConnectedTel);
     if Result = '' then
     begin
          // Sinon, selon origine de l'appel, l'appellant ou l'appellé
          case m_lOrigin of
               ctiOrigin_Outgoing:
                  Result := Trim (m_CalledTel);

               ctiOrigin_Incoming:
                  Result := Trim (m_CallerTel);
          end;

          // $$$ JP 31/10/06: on prend le n° connecté si existe et différent de celui sélectionné
          strConnectedTel := Trim (m_ConnectedTel);
          if (strConnectedTel <> '') and (Result <> strConnectedTel) then
             Result := strConnectedTel;

          // Au pire, numéro invisible
          if (iIndex = 1) and (Result = '') then
             Result := '<n° invisible>'
     end;
end;

function TCtiCall.GetHorodate:string;
begin
     Result := '';
     if m_TimeStart > 0 then
     begin
          // Heure de début de l'horodatage
          Result := FormatDateTime ('hh"h"nn', m_TimeStart);

          // On veut aussi la durée, si on a l'heure de fin
          if m_TimeEnd > 0 then
             Result := Result + ', ' + FormatDateTime ('nn"''"ss"''''"', m_TimeEnd-m_TimeStart);
     end;
end;

function TCtiCall.GetFaceName (iIndex:integer):string;
begin
     Result := '';
     if iIndex = 1 then
     begin
          // Texte en fonction de l'état de la communication
          case m_lState of
               ctiState_Offering:
                  Result := 'De ' + CorrespondentTelFaceName; //GetNumTel (m_CallerTel);

               ctiState_Connected:
               begin
                    // Horodatage
                    Result := Horodate;
                    if Result <> '' then
                       Result := '(' + Result + ') ';

                    case m_lOrigin of
                         ctiOrigin_Outgoing:
                            Result := Result + 'Vers ';

                         ctiOrigin_Incoming:
                            Result := Result + 'De ';
                    end;

                    // $$$ JP 11/01/06 - n° de téléphone du correspondant
                    Result := Result + CorrespondentTelFaceName; // GetNumTel (m_ConnectedTel)
               end;

               ctiState_Disconnected,
               ctiState_Idle:
               begin
                    // $$$ JP 12/01/06 - heure de fin de l'appel (s'il était connecté)
                    Result := Horodate;
                    if Result <> '' then
                       Result := '(' + Result + ') ';

                    // $$$ JP 11/01/06 - normalement le connecté
                    case m_lStateDetails of
                         ctiStateC_DisconnectedNormal,
                         ctiStateC_Unknown,
                         ctiStateC_None:
                            Result := Result + 'Terminé ';

                         ctiStateC_DisconnectedBusy:
                            Result := Result + 'Occupé ';

                         ctiStateC_DisconnectedNoAnswer:
                            Result := Result + 'Injoignable ';
                    end;
                    Result := Result + CorrespondentTelFaceName;
               end;

               ctiState_OnHold:
               begin
                    Result := 'Attente ' + CorrespondentTelFaceName;
               end;

               ctiState_Dialtone:
                  Result := 'Attente de numérotation';

               ctiState_Dialing:
                  Result := 'Numérotation en cours';

               ctiState_Proceeding:
                  Result := 'Acheminement vers ' + CorrespondentTelFaceName;

               ctiState_Ringback:
                  Result := 'Sonnerie sur ' + CorrespondentTelFaceName;

               ctiState_Conferenced:
               begin
                    // $$$ JP 11/01/06 - normalement le connected id
                    Result := 'Conférence avec ' + CorrespondentTelFaceName;
               end;
          end;
     end;

     // Si une identité en sélection, on l'indique
     if (m_Contact <> nil) and (m_strContactName <> '') then
     begin
          if Result <> '' then
             Result := Result + ': ';
          Result := Result + m_strContactName;
     end;
end;


end.


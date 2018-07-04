unit CCACONTROLSLib_TLB;

// ************************************************************************ //
// AVERTISSEMENT                                                                 
// -------                                                                    
// Les types déclarés dans ce fichier ont été générés à partir de données lues 
// depuis la bibliothèque de types. Si cette dernière (via une autre bibliothèque de types 
// s'y référant) est explicitement ou indirectement ré-importée, ou la commande "Rafraîchir"  
// de l'éditeur de bibliothèque de types est activée lors de la modification de la bibliothèque 
// de types, le contenu de ce fichier sera régénéré et toutes les modifications      
// manuellement apportées seront perdues.                                     
// ************************************************************************ //

// PASTLWTR : $Revision:   1.88  $
// Fichier généré le 26/01/2004 10:27:38 depuis la bibliothèque de types ci-dessous.

// *************************************************************************//
// REMARQUE :                                                                      
// Les éléments gardés par $IFDEF_LIVE_SERVER_AT_DESIGN_TIME sont utilisés par les propriétés 
// qui renvoient des objets qui peuvent nécessiter d'être explicitement créés via un appel de 
// fonction avant tout accès par une propriété. Ces éléments ont été désactivés pour 
// éviter une utilisation accidentelle depuis l'inspecteur d'objets. Vous   
// pouvez les activer en définissant LIVE_SERVER_AT_DESIGN_TIME ou en les  
// retirant sélectivement des blocs $IFDEF. Cependant, de tels éléments doivent toujours  
// être créés par programmation via une méthode de la classe CoClass appropriée   
// avant de pouvoir être utilisées.                                                          
// ************************************************************************ //
// Bibl.Types     : C:\Program Files\Alcatel\CCagent\CCAControls.dll (1)
// IID\LCID       : {C45074B5-DDA3-11D5-B773-00508B79E1ED}\0
// Fichier d'aide : 
// DepndLst       : 
//   (1) v2.0 stdole, (C:\WINNT\System32\STDOLE2.TLB)
//   (2) v4.0 StdVCL, (C:\WINNT\System32\STDVCL40.DLL)
// Erreurs :
//   Erreur à la création du bitmap de palette de (TDataOfCall) : Format GUID incorrect
//   Erreur à la création du bitmap de palette de (TItemOfHistoric) : Le serveur C:\Program Files\Alcatel\CCagent\CCAControls.dll ne contient pas d'icones
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
interface

uses Windows, ActiveX, Classes, Graphics, OleServer, OleCtrls;

// *********************************************************************//
// GUIDS déclarés dans la bibliothèque de types. Préfixes utilisés :    
//   Bibliothèques de types : LIBID_xxxx                                      
//   CoClasses              : CLASS_xxxx                                      
//   DISPInterfaces         : DIID_xxxx                                       
//   Non-DISP interfaces    : IID_xxxx                                        
// *********************************************************************//
const
  // Version majeure et mineure de la bibliothèque de types
  CCACONTROLSLibMajorVersion = 1;
  CCACONTROLSLibMinorVersion = 0;

  LIBID_CCACONTROLSLib: TGUID = '{C45074B5-DDA3-11D5-B773-00508B79E1ED}';

  IID_IAxApplication: TGUID = '{C45074C1-DDA3-11D5-B773-00508B79E1ED}';
  DIID__IAxApplicationEvents: TGUID = '{C45074C2-DDA3-11D5-B773-00508B79E1ED}';
  CLASS_AxApplication: TGUID = '{9E78C026-DD92-11D5-B772-00508B79E1ED}';
  IID_IDataOfCall: TGUID = '{D71863E5-5B66-11D6-8C70-0008C7BF3D97}';
  CLASS_DataOfCall: TGUID = '{D71863E6-5B66-11D6-8C70-0008C7BF3D97}';
  IID_IAxLine: TGUID = '{C45074C3-DDA3-11D5-B773-00508B79E1ED}';
  DIID__IAxLineEvents: TGUID = '{C45074C4-DDA3-11D5-B773-00508B79E1ED}';
  CLASS_AxLine: TGUID = '{9E78C029-DD92-11D5-B772-00508B79E1ED}';
  IID_IAxMakeCall: TGUID = '{AFA40FF3-E319-11D5-B776-00508B79E1ED}';
  DIID__IAxMakeCallEvents: TGUID = '{AFA40FF5-E319-11D5-B776-00508B79E1ED}';
  CLASS_AxMakeCall: TGUID = '{AFA40FF4-E319-11D5-B776-00508B79E1ED}';
  IID_IAxHangUp: TGUID = '{AFA40FFA-E319-11D5-B776-00508B79E1ED}';
  DIID__IAxHangUpEvents: TGUID = '{AFA40FFC-E319-11D5-B776-00508B79E1ED}';
  CLASS_AxHangUp: TGUID = '{AFA40FFB-E319-11D5-B776-00508B79E1ED}';
  IID_IAxTakeCall: TGUID = '{AA674F6F-E801-11D5-B778-00508B79E1ED}';
  DIID__IAxTakeCallEvents: TGUID = '{AA674F70-E801-11D5-B778-00508B79E1ED}';
  CLASS_AxTakeCall: TGUID = '{BF1B0DB8-E7FC-11D5-B778-00508B79E1ED}';
  IID_IAxConference: TGUID = '{37D5EB36-E88F-11D5-B779-00508B79E1ED}';
  DIID__IAxConferenceEvents: TGUID = '{37D5EB38-E88F-11D5-B779-00508B79E1ED}';
  CLASS_AxConference: TGUID = '{37D5EB37-E88F-11D5-B779-00508B79E1ED}';
  IID_IAxLog: TGUID = '{7F272F84-E98E-11D5-B779-00508B79E1ED}';
  DIID__IAxLogEvents: TGUID = '{7F272F86-E98E-11D5-B779-00508B79E1ED}';
  CLASS_AxLog: TGUID = '{7F272F85-E98E-11D5-B779-00508B79E1ED}';
  IID_IItemOfHistoric: TGUID = '{7CB3D129-5C21-11D6-8C70-0008C7BF3D97}';
  CLASS_ItemOfHistoric: TGUID = '{7CB3D12A-5C21-11D6-8C70-0008C7BF3D97}';
  IID_IAxHistoric: TGUID = '{1ECE7398-EAF5-11D5-B77A-00508B79E1ED}';
  DIID__IAxHistoricEvents: TGUID = '{1ECE739A-EAF5-11D5-B77A-00508B79E1ED}';
  CLASS_AxHistoric: TGUID = '{1ECE7399-EAF5-11D5-B77A-00508B79E1ED}';
  IID_IAxTransfertCall: TGUID = '{C989EA69-F2C8-11D5-B77B-00508B79E1ED}';
  DIID__IAxTransfertCallEvents: TGUID = '{C989EA6B-F2C8-11D5-B77B-00508B79E1ED}';
  CLASS_AxTransfertCall: TGUID = '{C989EA6A-F2C8-11D5-B77B-00508B79E1ED}';
  IID_IAxStartStopRecording: TGUID = '{E810B9BA-F45D-11D5-B77E-00508B79E1ED}';
  DIID__IAxStartStopRecordingEvents: TGUID = '{E810B9BC-F45D-11D5-B77E-00508B79E1ED}';
  CLASS_AxStartStopRecording: TGUID = '{E810B9BB-F45D-11D5-B77E-00508B79E1ED}';
  IID_IAxAgentStatistics: TGUID = '{47B38D57-F9ED-11D5-B77F-00508B79E1ED}';
  DIID__IAxAgentStatisticsEvents: TGUID = '{47B38D59-F9ED-11D5-B77F-00508B79E1ED}';
  CLASS_AxAgentStatistics: TGUID = '{47B38D58-F9ED-11D5-B77F-00508B79E1ED}';
  IID_IAxTeamStatistics: TGUID = '{94EDB881-FA19-11D5-B77F-00508B79E1ED}';
  DIID__IAxTeamStatisticsEvents: TGUID = '{94EDB883-FA19-11D5-B77F-00508B79E1ED}';
  CLASS_AxTeamStatistics: TGUID = '{94EDB882-FA19-11D5-B77F-00508B79E1ED}';
  IID_IAxHold: TGUID = '{4BB01D3F-035A-11D6-B782-00508B79E1ED}';
  DIID__IAxHoldEvents: TGUID = '{4BB01D41-035A-11D6-B782-00508B79E1ED}';
  CLASS_AxHold: TGUID = '{4BB01D40-035A-11D6-B782-00508B79E1ED}';
  IID_IAxDoNotCallRecord: TGUID = '{114C6FA3-04FC-11D6-B785-00508B79E1ED}';
  DIID__IAxDoNotCallRecordEvents: TGUID = '{114C6FA5-04FC-11D6-B785-00508B79E1ED}';
  CLASS_AxDoNotCallRecord: TGUID = '{114C6FA4-04FC-11D6-B785-00508B79E1ED}';
  IID_IAxRequestRejectPreviewRecord: TGUID = '{50A0753E-0513-11D6-B785-00508B79E1ED}';
  DIID__IAxRequestRejectPreviewRecordEvents: TGUID = '{50A07540-0513-11D6-B785-00508B79E1ED}';
  CLASS_AxRequestRejectPreviewRecord: TGUID = '{50A0753F-0513-11D6-B785-00508B79E1ED}';
  IID_IAxScheduleCallback: TGUID = '{50A0755E-0513-11D6-B785-00508B79E1ED}';
  DIID__IAxScheduleCallbackEvents: TGUID = '{50A07560-0513-11D6-B785-00508B79E1ED}';
  CLASS_AxScheduleCallback: TGUID = '{50A0755F-0513-11D6-B785-00508B79E1ED}';
  IID_IAxClassify: TGUID = '{50A0757C-0513-11D6-B785-00508B79E1ED}';
  DIID__IAxClassifyEvents: TGUID = '{50A0757E-0513-11D6-B785-00508B79E1ED}';
  CLASS_AxClassify: TGUID = '{50A0757D-0513-11D6-B785-00508B79E1ED}';

// *********************************************************************//
// Déclaration d'énumérations définies dans la bibliothèque de types    
// *********************************************************************//
// Constantes pour enum AxTServerErrorTypes
type
  AxTServerErrorTypes = TOleEnum;
const
  AxAxALC_TSERVER_SUCCESS = $00000000;
  AxALC_TSERVER_OBJECT_UNKNOWN = $FFFFFC18;
  AxALC_TSERVER_OBJECT_NORESPONSE = $FFFFFC17;
  AxALC_TSERVER_CLIENT_NOT_CONNECTED = $FFFFFC16;
  AxALC_TSERVER_CLIENT_ALREADY_CONNECTED = $FFFFFC15;
  AxALC_TSERVER_NO_HOST = $FFFFFC14;
  AxALC_TSERVER_OBJECT_NOT_CREATED = $FFFFFC13;
  AxALC_TSERVER_SERV_NOT_REGISTERED = $FFFFFC12;
  AxALC_TSERVER_SERV_NOT_EXIST = $FFFFFC11;
  AxALC_TSERVER_CLIENT_NOT_EXIST = $FFFFFC10;
  AxALC_TSERVER_CANT_ACTIVATE_SERVICE = $FFFFFC0F;
  AxALC_TSERVER_OUT_BOUND = $FFFFFC04;
  AxALC_TSERVER_ID_CALL_INVALID = $FFFFFC03;
  AxALC_TSERVER_INVALID_PARAMETER = $FFFFFC02;
  AxALC_TSERVER_OPERATION_NOT_EXIST = $FFFFFC01;
  AxALC_TSERVER_BAD_PASSWORD = $FFFFFC00;
  AxALC_TSERVER_SEND_NOT_OK = $FFFFFBFF;
  AxALC_TSERVER_READ_FILE = $FFFFFBFE;
  AxALC_TSERVER_NO_RESPONSE_FROM_RCSTA = $FFFFFBFD;
  AxALC_TSERVER_REJ_ACAPI = $FFFFFBFA;
  AxALC_TSERVER_BAD_SECRET_CODE = $FFFFFBF9;
  AxALC_TSERVER_INVALID_VERSION = $FFFFFBF8;
  AxALC_TSERVER_UNKNOWN = $FFFFFBF7;
  AxALC_TSERVER_PROTOCOL = $FFFFFBF6;
  AxALC_TSERVER_OP_ALREADY_PROGRESS = $FFFFFBF5;
  AxALC_TSERVER_USER_ALREADY_CHECKIN = $FFFFFBF4;
  AxALC_TSERVER_MAX_CHECKINS_REACHED = $FFFFFBF3;
  AxALC_TSERVER_MAX_NOTES_REACHED = $FFFFFBF2;
  AxALC_TSERVER_REPERTORY_NUMBER = $FFFFFBF0;
  AxALC_TSERVER_PHONE_BOOK_UNKNOWN = $FFFFFBEF;
  AxALC_TSERVER_PHONE_BOOK_NOT_RECORDED = $FFFFFBEE;
  AxALC_TSERVER_REPERTORY_FULL = $FFFFFBED;
  AxALC_TSERVER_ADMISTRATOR_RECORD = $FFFFFBE6;
  AxALC_TSERVER_USER_LOGINNAME_ALREADY_USED = $FFFFFBE5;
  AxALC_TSERVER_USER_USER_NOMADIC_NOT_IDLE = $FFFFFBE4;
  AxALC_TSERVER_LDAP_REJECT = $FFFFFBDC;
  AxALC_TSERVER_LDAP_SERVEUR = $FFFFFBDB;
  AxALC_TSERVER_LDAP_SERVEUR_RECORD = $FFFFFBDA;
  AxALC_TSERVER_LDAP_NB_RSP = $FFFFFBD9;
  AxALC_TSERVER_MAX_PLANNINGS_REACHED = $FFFFFBD8;
  AxALC_TSERVER_LINE_NOT_PCCM2 = $FFFFFBD2;
  AxALC_TSERVER_PCCM2_AUDIO_BUSY = $FFFFFBD1;
  AxALC_TSERVER_NO_RESPONSE_FROM_TLIB = $FFFFFBC8;
  AxALC_TSERVER_NO_ACTIVE_CAMPAIGNS = $FFFFFBC7;
  AxALC_TSERVER_NO_ACTIVE_PREVIEW_CAMPAIGNS = $FFFFFBC6;
  AxALC_TSERVER_NO_RECORDS_AVAILABLE = $FFFFFBC5;
  AxALC_TSERVER_RECORD_NOT_FOUND = $FFFFFBC4;
  AxALC_TSERVER_INVALID_REQUEST_DATA = $FFFFFBC3;
  AxALC_TSERVER_INVALID_REQUEST = $FFFFFBC2;
  AxALC_TSERVER_INVALID_TIME = $FFFFFBC1;
  AxALC_TSERVER_INVALID_TIME_FORMAT = $FFFFFBC0;
  AxALC_TSERVER_RECORD_ALREADY_PROCESSED = $FFFFFBBF;
  AxALC_TSERVER_DB_ERROR = $FFFFFBBE;
  AxALC_TSERVER_CHAINED_RECORD_NOT_FOUND = $FFFFFBBD;
  AxALC_TSERVER_RECORD_ALREADY_EXIST = $FFFFFBBC;
  AxALC_TSERVER_ADD_RECORD_ERROR = $FFFFFBBB;
  AxALC_TSERVER_SCHEDULED_RECORD_NOT_FOUND = $FFFFFBBA;
  AxALC_TSERVER_CSTA_SERVICE_NOT_PROVIDED = $FFFFEC77;
  AxALC_TSERVER_SERVER_CSTA_NOT_CONNECTED = $FFFFEC76;
  AxALC_TSERVER_CSTA_REQUEST_FAILURE = $FFFFEC75;
  AxALC_TSERVER_UNKNOWN_USER = $FFFFEC74;
  AxALC_TSERVER_UNKNOWN_LINE = $FFFFEC73;
  AxALC_TSERVER_NOT_IN_CLIENTS_LIST = $FFFFEC72;
  AxALC_TSERVER_USER_NOT_CHECKED_IN = $FFFFEC71;
  AxALC_TSERVER_INCORRECT_PASSWORD = $FFFFEC70;
  AxALC_TSERVER_USER_NOT_AUTHORIZED = $FFFFEC6F;
  AxALC_TSERVER_TSA_UNREACHABLE = $FFFFEC6E;
  AxALC_TSERVER_FROZEN_APPLI = $FFFFEC6D;
  AxALC_TSERVER_INCORRECT_MSG = $FFFFEC6C;
  AxALC_TSERVER_REQ_INCOMPAT_WITH_STATE = $FFFFEC6B;
  AxALC_TSERVER_TEL_SERVICE_NOT_PROVIDED = $FFFFEC6A;
  AxALC_TSERVER_CFG_SERVICE_NOT_PROVIDED = $FFFFEC69;
  AxALC_TSERVER_USER_NOT_ON_HIS_PC = $FFFFEC68;
  AxALC_TSERVER_VMAIL_SERVICE_NOT_PROVIDED = $FFFFEC67;
  AxALC_TSERVER_TICKETS_SERVICE_NOT_READY = $FFFFEC66;
  AxALC_TSERVER_UNKNOWN_VMAIL = $FFFFEC65;
  AxALC_TSERVER_ACAPI = $FFFFEC64;
  AxALC_TSERVER_ERROR_FILE = $FFFFEC63;
  AxALC_TSERVER_INCORRECT_PARAMETERS = $FFFFEC62;
  AxALC_TSERVER_OTHER_ADMIN_WORKING = $FFFFEC61;
  AxALC_TSERVER_USER_IS_ON_OTHER_TSA = $FFFFEC60;
  AxALC_TSERVER_NO_SPECIFIC_FUNCTION = $FFFFEC5F;
  AxALC_TSERVER_LOGIN_NOMADIC = $FFFFEC5E;
  AxALC_TSERVER_USER_LINE_BUSY = $FFFFEC5D;
  AxALC_TSERVER_LOCK_NOT_LOADED = $FFFFEC5C;
  AxALC_TSERVER_PASSWORD_TO_CHANGE = $FFFFEC5B;
  AxALC_TSERVER_Z_NOMADIC_CLEARING = $FFFFEC5A;
  AxALC_TSERVER_TEMP_ALERT_NOMADIC = $FFFFEC59;
  AxALC_TSERVER_TEMP_CONNECT_NOMADIC = $FFFFEC58;
  AxALC_TSERVER_TEMP_ANSWER_GSM = $FFFFEC57;
  AxALC_TSERVER_NO_GHOST_AVAILABLE = $FFFFEC56;
  AxALC_TSERVER_UNKNOWN_NOMADIC_MODE = $FFFFEC55;
  AxALC_TSERVER_INSUFFICIENT_NOMADIC_RIGHTS = $FFFFEC54;
  AxALC_TSERVER_REJ_MINIMESS_SRVCE_NOT_PROVIDED = $FFFFEC53;
  AxALC_TSERVER_BLOCKED_LOGIN = $FFFFEC52;
  AxALC_TSERVER_Z_NOMADIC_MANAGEMENT = $FFFFEC51;
  AxALC_TSERVER_NOMADIC_PHONIE = $FFFFEC50;
  AxALC_TSERVER_FREE_NOMADIC_RESOURCE = $FFFFEC4F;
  AxALC_TSERVER_UNKNOWN_NOTE_ID = $FFFFEC4E;
  AxALC_TSERVER_GRP_ERROR_UNKNOWN_GROUP = $FFFFEC4D;
  AxALC_TSERVER_GRP_ERROR_UNKNOWN_ELT_GROUP = $FFFFEC4C;
  AxALC_TSERVER_GRP_ERROR_LOG_STATE = $FFFFEC4B;
  AxALC_TSERVER_GRP_INVALID_LOG_MOD = $FFFFEC4A;
  AxALC_TSERVER_NOMAD_UNPARC = $FFFFEC49;
  AxALC_TSERVER_NOT_PROVIDED_ON_THIS_TSA = $FFFFEC48;
  AxALC_TSERVER_CCA_USER_CHECKIN_TWICE = $FFFFEC47;
  AxALC_TSERVER_CCA_CONFIG_SDK_NOT_PROVIDED = $FFFFE890;

// Constantes pour enum AxPCAgentErrorTypes
type
  AxPCAgentErrorTypes = TOleEnum;
const
  AxALC_PCAGENT_SUCCESS = $00000000;
  AxALC_ERR_PCAGENT_OBJECT_UNKNOWN = $000003E8;
  AxALC_ERR_PCAGENT_CLIENT_CREATION = $000003E9;
  AxALC_ERR_PCAGENT_CONNECT_IN_PROGRESS = $000003EA;
  AxALC_ERR_PCAGENT_CLIENT_ALLREADY_CONNECTED = $000003EB;
  AxALC_ERR_PCAGENT_HOST_NAME_REQUIRED = $000003EC;
  AxALC_ERR_PCAGENT_INVALID_HOST_NAME = $000003ED;
  AxALC_ERR_PCAGENT_INVALID_PORT_NUMBER = $000003EE;
  AxALC_ERR_PCAGENT_TERMINAL_NAME_REQUIRED = $000003EF;
  AxALC_ERR_PCAGENT_INVALID_TERMINAL_NAME = $000003F0;
  AxALC_ERR_PCAGENT_TERMINAL_NAME_TWICE = $000003F1;
  AxALC_ERR_PCAGENT_INVALID_SITE_NAME = $000003F2;
  AxALC_ERR_PCAGENT_OPEN_SOCKET = $000003F3;
  AxALC_ERR_PCAGENT_CONNECT_TIMEOUT = $000003F4;
  AxALC_ERR_PCAGENT_SOCKET_WRITE = $000003F5;
  AxALC_ERR_PCAGENT_PING_PONG_TIMEOUT = $000003F6;
  AxALC_ERR_PCAGENT_DISCONNECTED_BY_SERVER = $000003F7;
  AxALC_ERR_PCAGENT_IP_ADDRESS_UNAVIALABLE = $000003F8;
  AxALC_ERR_PCAGENT_DIR_NUMBER_REQUIRED = $000003F9;
  AxALC_ERR_PCAGENT_CLIENT_ALLREADY_ATTACHED = $000003FA;
  AxALC_ERR_PCAGENT_CLIENT_NOT_FOUND = $000003FB;
  AxALC_ERR_PCAGENT_INVALID_DIR_NUMBER = $000003FC;
  AxALC_ERR_PCAGENT_SAME_OBJECT_ATTACHED = $000003FD;
  AxALC_ERR_PCAGENT_DETTACHED_FROM_CLIENT = $000003FE;
  AxALC_ERR_PCAGENT_CLIENT_DISCONNECTED = $000003FF;
  AxALC_ERR_PCAGENT_OBJECT_NOT_REGISTERED = $00000400;
  AxALC_ERR_PCAGENT_OPERATION_UNKNOWN = $00000401;
  AxALC_ERR_PCAGENT_SUBOPERATION_UNKNOWN = $00000402;
  AxALC_ERR_PCAGENT_RESPONSE_TIMEOUT = $00000403;
  AxALC_ERR_PCAGENT_RESPONSE_DIR_NUMBER = $00000404;
  AxALC_ERR_PCAGENT_RESPONSE_NO_INFORMATION = $00000405;
  AxALC_ERR_PCAGENT_FORBIDDEN_OPERATION = $00000406;
  AxALC_ERR_PCAGENT_NULL_OBJECT_KEY = $00000407;

// Constantes pour enum AxVisibleStatus
type
  AxVisibleStatus = TOleEnum;
const
  AxALC_DEFAULT_VISIBLE = $00000001;
  AxALC_MAINWINDOW_VISIBLE = $00000002;
  AxALC_TELEPHONEBAR_VISIBLE = $00000004;
  AxALC_AGENTBAR_VISIBLE = $00000008;
  AxALC_CALLSBAR_VISIBLE = $00000010;
  AxALC_STATISTICSBAR_VISIBLE = $00000020;
  AxALC_NICE_VISIBLE = $00000100;
  AxALC_SERVICEBAR_VISIBLE = $00000200;
  AxALC_OUTBOUND_VISIBLE = $00000400;
  AxALC_INTERNETTOOLS_VISIBLE = $00000800;
  AxALC_SCRIPTINGWND_VISIBLE = $00001000;
  AxALC_UADWND_VISIBLE = $00002000;

// Constantes pour enum AxServerStatus
type
  AxServerStatus = TOleEnum;
const
  AxALC_PHONE_SERVER_OK = $00000001;
  AxALC_STATISTICS_SERVER_OK = $00000002;
  AxALC_OLE_SERVER_OK = $00000004;

// Constantes pour enum AxServiceOptions
type
  AxServiceOptions = TOleEnum;
const
  AxALC_SO_NICE_RECORDER = $00000001;
  AxALC_SO_OUTBOUND_CAMPAIGN = $00000002;
  AxALC_SO_CCI = $00000004;

// Constantes pour enum AxAgentServiceStates
type
  AxAgentServiceStates = TOleEnum;
const
  AxALC_SERVICE_STATE_UNKNOWN = $00000000;
  AxALC_SERVICE_STATE_LOGGED_OUT = $00000001;
  AxALC_SERVICE_STATE_LOGGED_IN = $00000002;
  AxALC_SERVICE_STATE_AFFECTED = $00000003;
  AxALC_SERVICE_STATE_WITHDRAWED = $00000004;

// Constantes pour enum AxAgentLineStates
type
  AxAgentLineStates = TOleEnum;
const
  AxALC_LINE_STATE_UNKNOWN = $00000000;
  AxALC_LINE_STATE_FREE = $00000001;
  AxALC_LINE_STATE_LOCK_OUT = $00000002;
  AxALC_LINE_STATE_OUT_OF_SERVICE = $00000003;
  AxALC_LINE_STATE_RINGING_ACD = $00000004;
  AxALC_LINE_STATE_TALKING_ACD = $00000005;
  AxALC_LINE_STATE_CONSULTATING_ACD = $00000006;
  AxALC_LINE_STATE_HELP = $00000007;
  AxALC_LINE_STATE_CONFERENCING_ACD = $00000008;
  AxALC_LINE_STATE_TRANSACTION = $00000009;
  AxALC_LINE_STATE_PAUSE = $0000000A;
  AxALC_LINE_STATE_WRAPUP = $0000000B;
  AxALC_LINE_STATE_SILENT_LISTENING = $0000000C;
  AxALC_LINE_STATE_MONITORED = $0000000D;
  AxALC_LINE_STATE_RECORDING = $0000000E;
  AxALC_LINE_STATE_LOGOFF = $0000000F;
  AxALC_LINE_STATE_HELD = $00000010;
  AxALC_LINE_STATE_DIALING = $00000011;
  AxALC_LINE_STATE_RINGING = $00000012;
  AxALC_LINE_STATE_INTERNAL_TALKING = $00000013;
  AxALC_LINE_STATE_EXTERNAL_TALKING = $00000014;
  AxALC_LINE_STATE_CONSULTING = $00000015;
  AxALC_LINE_STATE_CONFERENCING = $00000016;
  AxALC_LINE_STATE_BUSY_TONE = $00000017;
  AxALC_LINE_STATE_BLOCKED = $00000018;
  AxALC_LINE_STATE_TALKING_ACD_OUT = $00000019;

// Constantes pour enum AxACDTeamStateChanges
type
  AxACDTeamStateChanges = TOleEnum;
const
  AxALC_ST_TEAM_NUMBER_OF_LOGGED_AGENTS = $00000001;
  AxALC_ST_TEAM_NUMBER_OF_AFFECTED_AGENTS = $00000002;
  AxALC_ST_TEAM_NUMBER_OF_WITHDRAWED_AGENTS = $00000004;
  AxALC_ST_TEAM_NUMBER_OF_BUSY_AGENTS = $00000008;
  AxALC_ST_TEAM_NUMBER_OF_FREE_AGENTS = $00000010;
  AxALC_ST_TEAM_NUMBER_OF_ACD_CALLS = $00000020;
  AxALC_ST_TEAM_NUMBER_OF_PRIVATE_CALLS = $00000040;
  AxALC_ST_TEAM_NUMBER_OF_WAITING_CALLS = $00000080;
  AxALC_ST_TEAM_MAXIMUM_WAITING_TIME = $00000100;
  AxALC_ST_TEAM_AVERAGE_WAITING_TIME = $00000200;
  AxALC_ST_TEAM_ALL_PARAMETERS = $000003FF;

// Constantes pour enum AxLineStates
type
  AxLineStates = TOleEnum;
const
  AxALC_AGEMT_STATE_UNKNOWN = $00000000;
  AxALC_AGENT_STATE_LOGGED_OUT = $00000001;
  AxALC_AGENT_STATE_LOGGED_IN = $00000002;
  AxALC_AGENT_STATE_AFFECTED = $00000003;
  AxALC_AGENT_STATE_WITHDRAWED = $00000004;
  AxALC_AGENT_STATE_WRAPUP = $00000005;
  AxALC_AGENT_STATE_PAUSE = $00000006;
  AxALC_AGENT_STATE_TRANSACTION = $00000007;
  AxALC_AGENT_STATE_BUSY = $00000008;

// Constantes pour enum AxAgentCallStates
type
  AxAgentCallStates = TOleEnum;
const
  AxALC_CALL_STATE_UNKNOWN = $00000000;
  AxALC_CALL_STATE_RECORDING = $00000004;
  AxALC_CALL_STATE_DIALING = $00000005;
  AxALC_CALL_STATE_CALLING = $00000006;
  AxALC_CALL_STATE_RINGING = $00000007;
  AxALC_CALL_STATE_TALKING = $00000008;
  AxALC_CALL_STATE_HELD = $00000009;
  AxALC_CALL_STATE_CONFERENCED = $0000000B;
  AxALC_CALL_STATE_BUSYTONE = $0000000C;
  AxALC_CALL_STATE_BLOCKED = $0000000D;
  AxALC_CALL_STATE_WAITING = $0000000E;
  AxALC_CALL_STATE_RELEASING = $0000000F;
  AxALC_CALL_STATE_INIT = $00000011;

// Constantes pour enum AxAgentCallCauses
type
  AxAgentCallCauses = TOleEnum;
const
  AxALC_CALL_CAUSE_NONE = $00000001;
  AxALC_CALL_CAUSE_BUSY = $00000002;
  AxALC_CALL_CAUSE_NOTOBTAINABLE = $00000003;
  AxALC_CALL_CAUSE_DONOTDISTURB = $00000004;
  AxALC_CALL_CAUSE_OVERFLOW = $00000007;
  AxALC_CALL_CAUSE_TRUNKSBUSY = $00000009;
  AxALC_CALL_CAUSE_REDIRECTEDINCOMING = $0000000E;
  AxALC_CALL_CAUSE_REDIRECTEDOUTGOING = $0000000F;

// Constantes pour enum AxPilotStatus
type
  AxPilotStatus = TOleEnum;
const
  AxALC_PILOT_OPENED = $00000000;
  AxALC_PILOT_BLOCKED = $00000001;
  AxALC_PILOT_BLOCKEDONRULE = $00000002;
  AxALC_PILOT_BLOCKEDONBLOCKEDRULE = $00000003;
  AxALC_PILOT_FORWARDING = $00000004;
  AxALC_PILOT_FORWARDINGONRULE = $00000005;
  AxALC_PILOT_BLOCKEDONGENERALFORWARDINGRULE = $00000006;
  AxALC_PILOT_OTHERSTATUS = $000000FF;

// Constantes pour enum AxAgentLineStateChanges
type
  AxAgentLineStateChanges = TOleEnum;
const
  AxALC_ST_AGENT_SERVICE_STATE = $00000001;
  AxALC_ST_AGENT_SERVICE_STATE_DATE = $00000002;
  AxALC_ST_AGENT_LAST_LOGIN_DATE = $00000004;
  AxALC_ST_AGENT_STATE = $00000008;
  AxALC_ST_AGENT_STATE_DATE = $00000010;
  AxALC_ST_AGENT_NUMBER_OF_WITHDRAWS = $00000020;
  AxALC_ST_AGENT_WITHDRAW_DURATION = $00000040;
  AxALC_ST_AGENT_NUMBER_OF_ACD_CALLS = $00000080;
  AxALC_ST_AGENT_ACD_CALLS_DURATION = $00000100;
  AxALC_ST_AGENT_NUMBER_OF_PRIVATE_CALLS = $00000200;
  AxALC_ST_AGENT_PRIVATE_CALLS_DURATION = $00000400;
  AxALC_ST_AGENT_NUMBER_OF_INTERCEPTED_CALLS = $00000800;
  AxALC_ST_AGENT_NUMBER_OF_TRANSFERRED_CALLS = $00001000;
  AxALC_ST_AGENT_NUMBER_OF_REFUSED_CALLS = $00002000;
  AxALC_ST_AGENT_SERVED_PILOT = $00004000;
  AxALC_ST_AGENT_PAUSE_AFTER_WRAPUP = $00008000;
  AxALC_ST_AGENT_ALL_PARAMETERS = $0000FFFF;

// Constantes pour enum AxCampaignStatus
type
  AxCampaignStatus = TOleEnum;
const
  AxALC_NO_CAMPAIGN = $00000000;
  AxALC_CAMPAIGN_PREDICTIVE = $00000001;
  AxALC_CAMPAIGN_PROGRESSIVE = $00000002;
  AxALC_CAMPAIGN_PREVIEW = $00000003;
  AxALC_CAMPAIGN_ENGAGED_PREDICTIVE = $00000004;
  AxALC_CAMPAIGN_ENGAGED_PROGRESSIVE = $00000005;

// Constantes pour enum AxCallType
type
  AxCallType = TOleEnum;
const
  AxCCE_TREATMENT = $00000000;
  AxCCO_TREATMENT = $00000001;
  AxCCW_TREATMENT = $00000002;
  AxCCBK_TREATMENT = $00000003;
  AxAUDIO_TREATMENT = $000000FF;

type

// *********************************************************************//
// Déclaration Forward des types définis dans la bibliothèque de types    
// *********************************************************************//
  IAxApplication = interface;
  IAxApplicationDisp = dispinterface;
  _IAxApplicationEvents = dispinterface;
  IDataOfCall = interface;
  IDataOfCallDisp = dispinterface;
  IAxLine = interface;
  IAxLineDisp = dispinterface;
  _IAxLineEvents = dispinterface;
  IAxMakeCall = interface;
  IAxMakeCallDisp = dispinterface;
  _IAxMakeCallEvents = dispinterface;
  IAxHangUp = interface;
  IAxHangUpDisp = dispinterface;
  _IAxHangUpEvents = dispinterface;
  IAxTakeCall = interface;
  IAxTakeCallDisp = dispinterface;
  _IAxTakeCallEvents = dispinterface;
  IAxConference = interface;
  IAxConferenceDisp = dispinterface;
  _IAxConferenceEvents = dispinterface;
  IAxLog = interface;
  IAxLogDisp = dispinterface;
  _IAxLogEvents = dispinterface;
  IItemOfHistoric = interface;
  IItemOfHistoricDisp = dispinterface;
  IAxHistoric = interface;
  IAxHistoricDisp = dispinterface;
  _IAxHistoricEvents = dispinterface;
  IAxTransfertCall = interface;
  IAxTransfertCallDisp = dispinterface;
  _IAxTransfertCallEvents = dispinterface;
  IAxStartStopRecording = interface;
  IAxStartStopRecordingDisp = dispinterface;
  _IAxStartStopRecordingEvents = dispinterface;
  IAxAgentStatistics = interface;
  IAxAgentStatisticsDisp = dispinterface;
  _IAxAgentStatisticsEvents = dispinterface;
  IAxTeamStatistics = interface;
  IAxTeamStatisticsDisp = dispinterface;
  _IAxTeamStatisticsEvents = dispinterface;
  IAxHold = interface;
  IAxHoldDisp = dispinterface;
  _IAxHoldEvents = dispinterface;
  IAxDoNotCallRecord = interface;
  IAxDoNotCallRecordDisp = dispinterface;
  _IAxDoNotCallRecordEvents = dispinterface;
  IAxRequestRejectPreviewRecord = interface;
  IAxRequestRejectPreviewRecordDisp = dispinterface;
  _IAxRequestRejectPreviewRecordEvents = dispinterface;
  IAxScheduleCallback = interface;
  IAxScheduleCallbackDisp = dispinterface;
  _IAxScheduleCallbackEvents = dispinterface;
  IAxClassify = interface;
  IAxClassifyDisp = dispinterface;
  _IAxClassifyEvents = dispinterface;

// *********************************************************************//
// Déclaration de CoClasses définies dans la bibliothèque de types 
// (REMARQUE: On affecte chaque CoClass à son Interface par défaut)              
// *********************************************************************//
  AxApplication = IAxApplication;
  DataOfCall = IDataOfCall;
  AxLine = IAxLine;
  AxMakeCall = IAxMakeCall;
  AxHangUp = IAxHangUp;
  AxTakeCall = IAxTakeCall;
  AxConference = IAxConference;
  AxLog = IAxLog;
  ItemOfHistoric = IItemOfHistoric;
  AxHistoric = IAxHistoric;
  AxTransfertCall = IAxTransfertCall;
  AxStartStopRecording = IAxStartStopRecording;
  AxAgentStatistics = IAxAgentStatistics;
  AxTeamStatistics = IAxTeamStatistics;
  AxHold = IAxHold;
  AxDoNotCallRecord = IAxDoNotCallRecord;
  AxRequestRejectPreviewRecord = IAxRequestRejectPreviewRecord;
  AxScheduleCallback = IAxScheduleCallback;
  AxClassify = IAxClassify;


// *********************************************************************//
// Interface   : IAxApplication
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {C45074C1-DDA3-11D5-B773-00508B79E1ED}
// *********************************************************************//
  IAxApplication = interface(IDispatch)
    ['{C45074C1-DDA3-11D5-B773-00508B79E1ED}']
    function  Get_StatServerName: WideString; safecall;
    function  Get_Version: WideString; safecall;
    function  Get_Time: TDateTime; safecall;
    function  Get_Visible: Integer; safecall;
    procedure Set_Visible(pVal: Integer); safecall;
    function  Get_ServerStatus: AxServerStatus; safecall;
    function  Get_PhoneServerName: WideString; safecall;
    function  Get_PhoneServerPort: Integer; safecall;
    function  Get_StatServerSite: WideString; safecall;
    function  Get_NbEvtQueued: Integer; safecall;
    function  Get_ServiceOption: Integer; safecall;
    function  Get_FirstName: WideString; safecall;
    function  Get_LastName: WideString; safecall;
    function  Get_LoginName: WideString; safecall;
    function  Connect: Integer; safecall;
    function  Get_StatServerPort: Integer; safecall;
    function  Get_CCagentVisibility: AxVisibleStatus; safecall;
    procedure Set_CCagentVisibility(pVal: AxVisibleStatus); safecall;
    property StatServerName: WideString read Get_StatServerName;
    property Version: WideString read Get_Version;
    property Time: TDateTime read Get_Time;
    property Visible: Integer read Get_Visible write Set_Visible;
    property ServerStatus: AxServerStatus read Get_ServerStatus;
    property PhoneServerName: WideString read Get_PhoneServerName;
    property PhoneServerPort: Integer read Get_PhoneServerPort;
    property StatServerSite: WideString read Get_StatServerSite;
    property NbEvtQueued: Integer read Get_NbEvtQueued;
    property ServiceOption: Integer read Get_ServiceOption;
    property FirstName: WideString read Get_FirstName;
    property LastName: WideString read Get_LastName;
    property LoginName: WideString read Get_LoginName;
    property StatServerPort: Integer read Get_StatServerPort;
    property CCagentVisibility: AxVisibleStatus read Get_CCagentVisibility write Set_CCagentVisibility;
  end;

// *********************************************************************//
// DispIntf:  IAxApplicationDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C45074C1-DDA3-11D5-B773-00508B79E1ED}
// *********************************************************************//
  IAxApplicationDisp = dispinterface
    ['{C45074C1-DDA3-11D5-B773-00508B79E1ED}']
    property StatServerName: WideString readonly dispid 1;
    property Version: WideString readonly dispid 2;
    property Time: TDateTime readonly dispid 3;
    property Visible: Integer dispid 4;
    property ServerStatus: AxServerStatus readonly dispid 5;
    property PhoneServerName: WideString readonly dispid 6;
    property PhoneServerPort: Integer readonly dispid 7;
    property StatServerSite: WideString readonly dispid 8;
    property NbEvtQueued: Integer readonly dispid 9;
    property ServiceOption: Integer readonly dispid 10;
    property FirstName: WideString readonly dispid 11;
    property LastName: WideString readonly dispid 12;
    property LoginName: WideString readonly dispid 13;
    function  Connect: Integer; dispid 14;
    property StatServerPort: Integer readonly dispid 15;
    property CCagentVisibility: AxVisibleStatus dispid 16;
  end;

// *********************************************************************//
// DispIntf:  _IAxApplicationEvents
// Flags:     (4096) Dispatchable
// GUID:      {C45074C2-DDA3-11D5-B773-00508B79E1ED}
// *********************************************************************//
  _IAxApplicationEvents = dispinterface
    ['{C45074C2-DDA3-11D5-B773-00508B79E1ED}']
    procedure OnQuit; dispid 1;
    procedure OnServerStatus(Status: AxServerStatus); dispid 2;
    procedure OnVisible(Status: AxVisibleStatus); dispid 3;
  end;

// *********************************************************************//
// Interface   : IDataOfCall
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {D71863E5-5B66-11D6-8C70-0008C7BF3D97}
// *********************************************************************//
  IDataOfCall = interface(IDispatch)
    ['{D71863E5-5B66-11D6-8C70-0008C7BF3D97}']
    function  Get_Id: Integer; safecall;
    function  Get_LastName: WideString; safecall;
    function  Get_FirstName: WideString; safecall;
    function  Get_Number: WideString; safecall;
    function  Get_Incoming: Integer; safecall;
    function  Get_Extern: Integer; safecall;
    function  Get_DestinationNumber: WideString; safecall;
    function  Get_ConferenceMaster: WideString; safecall;
    function  Get_State: AxAgentCallStates; safecall;
    function  Get_Cause: AxAgentCallCauses; safecall;
    function  Get_ACDCall: Integer; safecall;
    function  Get_TrunkUsed: WideString; safecall;
    function  Get_EstimatedWaitingTime: Integer; safecall;
    function  Get_PilotStatus: AxPilotStatus; safecall;
    function  Get_PossibleTransfer: Integer; safecall;
    function  Get_WaitingQueueSaturation: Integer; safecall;
    function  Get_LastRedirectionDevice: WideString; safecall;
    function  Get_CorrelatorData: WideString; safecall;
    procedure Set_CorrelatorData(const pVal: WideString); safecall;
    property Id: Integer read Get_Id;
    property LastName: WideString read Get_LastName;
    property FirstName: WideString read Get_FirstName;
    property Number: WideString read Get_Number;
    property Incoming: Integer read Get_Incoming;
    property Extern: Integer read Get_Extern;
    property DestinationNumber: WideString read Get_DestinationNumber;
    property ConferenceMaster: WideString read Get_ConferenceMaster;
    property State: AxAgentCallStates read Get_State;
    property Cause: AxAgentCallCauses read Get_Cause;
    property ACDCall: Integer read Get_ACDCall;
    property TrunkUsed: WideString read Get_TrunkUsed;
    property EstimatedWaitingTime: Integer read Get_EstimatedWaitingTime;
    property PilotStatus: AxPilotStatus read Get_PilotStatus;
    property PossibleTransfer: Integer read Get_PossibleTransfer;
    property WaitingQueueSaturation: Integer read Get_WaitingQueueSaturation;
    property LastRedirectionDevice: WideString read Get_LastRedirectionDevice;
    property CorrelatorData: WideString read Get_CorrelatorData write Set_CorrelatorData;
  end;

// *********************************************************************//
// DispIntf:  IDataOfCallDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {D71863E5-5B66-11D6-8C70-0008C7BF3D97}
// *********************************************************************//
  IDataOfCallDisp = dispinterface
    ['{D71863E5-5B66-11D6-8C70-0008C7BF3D97}']
    property Id: Integer readonly dispid 1;
    property LastName: WideString readonly dispid 2;
    property FirstName: WideString readonly dispid 3;
    property Number: WideString readonly dispid 4;
    property Incoming: Integer readonly dispid 5;
    property Extern: Integer readonly dispid 6;
    property DestinationNumber: WideString readonly dispid 7;
    property ConferenceMaster: WideString readonly dispid 8;
    property State: AxAgentCallStates readonly dispid 9;
    property Cause: AxAgentCallCauses readonly dispid 10;
    property ACDCall: Integer readonly dispid 11;
    property TrunkUsed: WideString readonly dispid 12;
    property EstimatedWaitingTime: Integer readonly dispid 13;
    property PilotStatus: AxPilotStatus readonly dispid 14;
    property PossibleTransfer: Integer readonly dispid 15;
    property WaitingQueueSaturation: Integer readonly dispid 16;
    property LastRedirectionDevice: WideString readonly dispid 17;
    property CorrelatorData: WideString dispid 18;
  end;

// *********************************************************************//
// Interface   : IAxLine
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {C45074C3-DDA3-11D5-B773-00508B79E1ED}
// *********************************************************************//
  IAxLine = interface(IDispatch)
    ['{C45074C3-DDA3-11D5-B773-00508B79E1ED}']
    function  Get_ActiveCall: Integer; safecall;
    function  Get_AffectedTeamDirNumber(nTeam: SYSINT): WideString; safecall;
    function  Get_ContactDataKey(handle: Integer; field: SYSINT): WideString; safecall;
    function  Get_ContactDataValue(handle: Integer; const key: WideString): OleVariant; safecall;
    procedure Set_ContactDataValue(handle: Integer; const key: WideString; pVal: OleVariant); safecall;
    function  Get_CampaignMode: Integer; safecall;
    function  Get_ContactDataVisible(handle: Integer): Integer; safecall;
    procedure Set_ContactDataVisible(handle: Integer; pVal: Integer); safecall;
    function  Get_IsChainedRecordPossible(handle: Integer): Integer; safecall;
    function  Get_IsConferencePossible: Integer; safecall;
    function  Get_IsDoNotCallPossible(handle: Integer): Integer; safecall;
    function  Get_IsHoldCallPossible: Integer; safecall;
    function  Get_IsLogoffPossible: Integer; safecall;
    function  Get_IsLogonPossible: Integer; safecall;
    function  Get_IsMakeCallPossible: Integer; safecall;
    function  Get_IsPauseEndPossible: Integer; safecall;
    function  Get_IsPausePossible: Integer; safecall;
    function  Get_IsPreviewRecordPossible: Integer; safecall;
    function  Get_IsRecordCancelPossible(handle: Integer): Integer; safecall;
    function  Get_IsRecordProcessedPossible(handle: Integer): Integer; safecall;
    function  Get_IsRecordRejectPossible(handle: Integer): Integer; safecall;
    function  Get_IsRecordReschedulePossible(handle: Integer): Integer; safecall;
    function  Get_IsReleaseCallPossible(callId: Integer): Integer; safecall;
    function  Get_IsTakeCallPossible(callId: Integer): Integer; safecall;
    function  Get_IsTransferCallPossible: Integer; safecall;
    function  Get_IsUpdateCallCompStatsPossible(handle: Integer): Integer; safecall;
    function  Get_IsWithdrawEndPossible: Integer; safecall;
    function  Get_IsWithdrawPossible: Integer; safecall;
    function  Get_IsWrapUpEndPossible: Integer; safecall;
    function  Get_IsWrapUpPossible: Integer; safecall;
    function  Get_LastError: AxTServerErrorTypes; safecall;
    function  Get_LineServiceState: AxAgentServiceStates; safecall;
    function  Get_LoggedTeam: WideString; safecall;
    function  Get_NbCalls: Integer; safecall;
    function  Get_NbEvtQueued: Integer; safecall;
    function  Get_Number: WideString; safecall;
    function  Get_NumberOfAffectedTeam: Integer; safecall;
    function  Get_NumberOfRecordFields(handle: Integer): Integer; safecall;
    function  Get_NumberOfWithdrawCauses: Integer; safecall;
    function  Get_ProACDExtension: WideString; safecall;
    procedure Set_ProACDExtension(const pVal: WideString); safecall;
    function  Get_WithdrawCause(nCause: Integer): WideString; safecall;
    function  ChainedRecord(handle: Integer): Integer; safecall;
    function  Conference: Integer; safecall;
    function  DoNotCall(handle: Integer): Integer; safecall;
    function  HoldCall: Integer; safecall;
    function  Logoff(const Agent: WideString; const Team: WideString; const Password: WideString): Integer; safecall;
    function  Logon(const Agent: WideString; const Team: WideString; const Password: WideString; 
                    outbound: Integer; multimediaTools: Integer): Integer; safecall;
    function  MakeCall(const destNumber: WideString; const correlator: WideString): Integer; safecall;
    function  Pause: Integer; safecall;
    function  PauseEnd: Integer; safecall;
    function  PreviewRecord: Integer; safecall;
    function  RecordCancel(handle: Integer): Integer; safecall;
    function  RecordProcessed(handle: Integer): Integer; safecall;
    function  RecordReject(handle: Integer): Integer; safecall;
    function  RecordReschedule(handle: Integer): Integer; safecall;
    function  ReleaseCall(callId: Integer): Integer; safecall;
    function  SetDateAndTime(const szDate: WideString): Integer; safecall;
    function  TakeCall(callId: Integer): Integer; safecall;
    function  TransferCall(heldCallId: Integer): Integer; safecall;
    function  UpdateCallCompletionStats(handle: Integer): Integer; safecall;
    function  Withdraw(Cause: Integer): Integer; safecall;
    function  WithdrawEnd: Integer; safecall;
    function  WrapUp: Integer; safecall;
    function  WrapUpEnd: Integer; safecall;
    function  Get_HandleOfCurrentRecord: Integer; safecall;
    procedure ShowEvents(Enabled: Integer); safecall;
    function  Connect: Integer; safecall;
    function  Get_CurrentReservationType: AxCallType; safecall;
    function  Get_CallById(callId: Integer): IDataOfCall; safecall;
    function  Get_CallByIndex(CallIndex: Integer): IDataOfCall; safecall;
    function  EndConference: Integer; safecall;
    function  IsEndConferencePossible: Integer; safecall;
    function  ContactAdd(const pContactNumber: WideString): Integer; safecall;
    function  ContactSave(nHandle: Integer): Integer; safecall;
    function  ContactCancel(nHandle: Integer): Integer; safecall;
    function  IsAddContactPossible: Integer; safecall;
    property ActiveCall: Integer read Get_ActiveCall;
    property AffectedTeamDirNumber[nTeam: SYSINT]: WideString read Get_AffectedTeamDirNumber;
    property ContactDataKey[handle: Integer; field: SYSINT]: WideString read Get_ContactDataKey;
    property ContactDataValue[handle: Integer; const key: WideString]: OleVariant read Get_ContactDataValue write Set_ContactDataValue;
    property CampaignMode: Integer read Get_CampaignMode;
    property ContactDataVisible[handle: Integer]: Integer read Get_ContactDataVisible write Set_ContactDataVisible;
    property IsChainedRecordPossible[handle: Integer]: Integer read Get_IsChainedRecordPossible;
    property IsConferencePossible: Integer read Get_IsConferencePossible;
    property IsDoNotCallPossible[handle: Integer]: Integer read Get_IsDoNotCallPossible;
    property IsHoldCallPossible: Integer read Get_IsHoldCallPossible;
    property IsLogoffPossible: Integer read Get_IsLogoffPossible;
    property IsLogonPossible: Integer read Get_IsLogonPossible;
    property IsMakeCallPossible: Integer read Get_IsMakeCallPossible;
    property IsPauseEndPossible: Integer read Get_IsPauseEndPossible;
    property IsPausePossible: Integer read Get_IsPausePossible;
    property IsPreviewRecordPossible: Integer read Get_IsPreviewRecordPossible;
    property IsRecordCancelPossible[handle: Integer]: Integer read Get_IsRecordCancelPossible;
    property IsRecordProcessedPossible[handle: Integer]: Integer read Get_IsRecordProcessedPossible;
    property IsRecordRejectPossible[handle: Integer]: Integer read Get_IsRecordRejectPossible;
    property IsRecordReschedulePossible[handle: Integer]: Integer read Get_IsRecordReschedulePossible;
    property IsReleaseCallPossible[callId: Integer]: Integer read Get_IsReleaseCallPossible;
    property IsTakeCallPossible[callId: Integer]: Integer read Get_IsTakeCallPossible;
    property IsTransferCallPossible: Integer read Get_IsTransferCallPossible;
    property IsUpdateCallCompStatsPossible[handle: Integer]: Integer read Get_IsUpdateCallCompStatsPossible;
    property IsWithdrawEndPossible: Integer read Get_IsWithdrawEndPossible;
    property IsWithdrawPossible: Integer read Get_IsWithdrawPossible;
    property IsWrapUpEndPossible: Integer read Get_IsWrapUpEndPossible;
    property IsWrapUpPossible: Integer read Get_IsWrapUpPossible;
    property LastError: AxTServerErrorTypes read Get_LastError;
    property LineServiceState: AxAgentServiceStates read Get_LineServiceState;
    property LoggedTeam: WideString read Get_LoggedTeam;
    property NbCalls: Integer read Get_NbCalls;
    property NbEvtQueued: Integer read Get_NbEvtQueued;
    property Number: WideString read Get_Number;
    property NumberOfAffectedTeam: Integer read Get_NumberOfAffectedTeam;
    property NumberOfRecordFields[handle: Integer]: Integer read Get_NumberOfRecordFields;
    property NumberOfWithdrawCauses: Integer read Get_NumberOfWithdrawCauses;
    property ProACDExtension: WideString read Get_ProACDExtension write Set_ProACDExtension;
    property WithdrawCause[nCause: Integer]: WideString read Get_WithdrawCause;
    property HandleOfCurrentRecord: Integer read Get_HandleOfCurrentRecord;
    property CurrentReservationType: AxCallType read Get_CurrentReservationType;
    property CallById[callId: Integer]: IDataOfCall read Get_CallById;
    property CallByIndex[CallIndex: Integer]: IDataOfCall read Get_CallByIndex;
  end;

// *********************************************************************//
// DispIntf:  IAxLineDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C45074C3-DDA3-11D5-B773-00508B79E1ED}
// *********************************************************************//
  IAxLineDisp = dispinterface
    ['{C45074C3-DDA3-11D5-B773-00508B79E1ED}']
    property ActiveCall: Integer readonly dispid 1;
    property AffectedTeamDirNumber[nTeam: SYSINT]: WideString readonly dispid 2;
    property ContactDataKey[handle: Integer; field: SYSINT]: WideString readonly dispid 3;
    property ContactDataValue[handle: Integer; const key: WideString]: OleVariant dispid 4;
    property CampaignMode: Integer readonly dispid 5;
    property ContactDataVisible[handle: Integer]: Integer dispid 6;
    property IsChainedRecordPossible[handle: Integer]: Integer readonly dispid 7;
    property IsConferencePossible: Integer readonly dispid 8;
    property IsDoNotCallPossible[handle: Integer]: Integer readonly dispid 9;
    property IsHoldCallPossible: Integer readonly dispid 10;
    property IsLogoffPossible: Integer readonly dispid 11;
    property IsLogonPossible: Integer readonly dispid 12;
    property IsMakeCallPossible: Integer readonly dispid 13;
    property IsPauseEndPossible: Integer readonly dispid 14;
    property IsPausePossible: Integer readonly dispid 15;
    property IsPreviewRecordPossible: Integer readonly dispid 16;
    property IsRecordCancelPossible[handle: Integer]: Integer readonly dispid 17;
    property IsRecordProcessedPossible[handle: Integer]: Integer readonly dispid 18;
    property IsRecordRejectPossible[handle: Integer]: Integer readonly dispid 19;
    property IsRecordReschedulePossible[handle: Integer]: Integer readonly dispid 20;
    property IsReleaseCallPossible[callId: Integer]: Integer readonly dispid 21;
    property IsTakeCallPossible[callId: Integer]: Integer readonly dispid 22;
    property IsTransferCallPossible: Integer readonly dispid 23;
    property IsUpdateCallCompStatsPossible[handle: Integer]: Integer readonly dispid 24;
    property IsWithdrawEndPossible: Integer readonly dispid 25;
    property IsWithdrawPossible: Integer readonly dispid 26;
    property IsWrapUpEndPossible: Integer readonly dispid 27;
    property IsWrapUpPossible: Integer readonly dispid 28;
    property LastError: AxTServerErrorTypes readonly dispid 29;
    property LineServiceState: AxAgentServiceStates readonly dispid 30;
    property LoggedTeam: WideString readonly dispid 31;
    property NbCalls: Integer readonly dispid 32;
    property NbEvtQueued: Integer readonly dispid 33;
    property Number: WideString readonly dispid 34;
    property NumberOfAffectedTeam: Integer readonly dispid 35;
    property NumberOfRecordFields[handle: Integer]: Integer readonly dispid 36;
    property NumberOfWithdrawCauses: Integer readonly dispid 37;
    property ProACDExtension: WideString dispid 38;
    property WithdrawCause[nCause: Integer]: WideString readonly dispid 39;
    function  ChainedRecord(handle: Integer): Integer; dispid 40;
    function  Conference: Integer; dispid 41;
    function  DoNotCall(handle: Integer): Integer; dispid 42;
    function  HoldCall: Integer; dispid 43;
    function  Logoff(const Agent: WideString; const Team: WideString; const Password: WideString): Integer; dispid 44;
    function  Logon(const Agent: WideString; const Team: WideString; const Password: WideString; 
                    outbound: Integer; multimediaTools: Integer): Integer; dispid 45;
    function  MakeCall(const destNumber: WideString; const correlator: WideString): Integer; dispid 46;
    function  Pause: Integer; dispid 47;
    function  PauseEnd: Integer; dispid 48;
    function  PreviewRecord: Integer; dispid 49;
    function  RecordCancel(handle: Integer): Integer; dispid 50;
    function  RecordProcessed(handle: Integer): Integer; dispid 51;
    function  RecordReject(handle: Integer): Integer; dispid 52;
    function  RecordReschedule(handle: Integer): Integer; dispid 53;
    function  ReleaseCall(callId: Integer): Integer; dispid 54;
    function  SetDateAndTime(const szDate: WideString): Integer; dispid 55;
    function  TakeCall(callId: Integer): Integer; dispid 56;
    function  TransferCall(heldCallId: Integer): Integer; dispid 57;
    function  UpdateCallCompletionStats(handle: Integer): Integer; dispid 58;
    function  Withdraw(Cause: Integer): Integer; dispid 59;
    function  WithdrawEnd: Integer; dispid 60;
    function  WrapUp: Integer; dispid 61;
    function  WrapUpEnd: Integer; dispid 62;
    property HandleOfCurrentRecord: Integer readonly dispid 63;
    procedure ShowEvents(Enabled: Integer); dispid 64;
    function  Connect: Integer; dispid 65;
    property CurrentReservationType: AxCallType readonly dispid 66;
    property CallById[callId: Integer]: IDataOfCall readonly dispid 67;
    property CallByIndex[CallIndex: Integer]: IDataOfCall readonly dispid 68;
    function  EndConference: Integer; dispid 69;
    function  IsEndConferencePossible: Integer; dispid 70;
    function  ContactAdd(const pContactNumber: WideString): Integer; dispid 71;
    function  ContactSave(nHandle: Integer): Integer; dispid 72;
    function  ContactCancel(nHandle: Integer): Integer; dispid 73;
    function  IsAddContactPossible: Integer; dispid 74;
  end;

// *********************************************************************//
// DispIntf:  _IAxLineEvents
// Flags:     (4096) Dispatchable
// GUID:      {C45074C4-DDA3-11D5-B773-00508B79E1ED}
// *********************************************************************//
  _IAxLineEvents = dispinterface
    ['{C45074C4-DDA3-11D5-B773-00508B79E1ED}']
    procedure OnLineNewCall(const pCall: IDataOfCall; callType: AxCallType); dispid 1;
    procedure OnLineChangeCall(const pCall: IDataOfCall); dispid 2;
    procedure OnLineDeleteCall(callId: Integer); dispid 3;
    procedure OnMakeCallPossible(Status: Integer); dispid 4;
    procedure OnTakeCallPossible(callId: Integer; Status: Integer); dispid 5;
    procedure OnReleaseCallPossible(callId: Integer; Status: Integer); dispid 6;
    procedure OnTransferCallPossible(Status: Integer); dispid 7;
    procedure OnConferenceCallPossible(Status: Integer); dispid 8;
    procedure OnHoldCallPossible(Status: Integer); dispid 9;
    procedure OnRequestStatus(requestId: Integer; Status: AxTServerErrorTypes); dispid 10;
    procedure OnLineStateChange(State: Integer); dispid 11;
    procedure OnLogonPossible(Status: Integer); dispid 12;
    procedure OnLogoffPossible(Status: Integer); dispid 13;
    procedure OnWrapUpPossible(Status: Integer); dispid 14;
    procedure OnWrapUpEndPossible(Status: Integer); dispid 15;
    procedure OnPausePossible(Status: Integer); dispid 16;
    procedure OnPauseEndPossible(Status: Integer); dispid 17;
    procedure OnWithdrawPossible(Status: Integer); dispid 18;
    procedure OnWithdrawEndPossible(Status: Integer); dispid 19;
    procedure OnCampaignModeChange(Status: Integer); dispid 20;
    procedure OnPreviewRecordPossible(Status: Integer); dispid 21;
    procedure OnRecordCancelPossible(Status: Integer); dispid 22;
    procedure OnRecordRejectPossible(Status: Integer); dispid 23;
    procedure OnDoNotCallPossible(Status: Integer); dispid 24;
    procedure OnRecordProcessedPossible(Status: Integer); dispid 25;
    procedure OnUpdateCallCompStatsPossible(Status: Integer); dispid 26;
    procedure OnRecordReschedulePossible(Status: Integer); dispid 27;
    procedure OnChainedRecordPossible(Status: Integer); dispid 28;
    procedure OnNewRecord(handle: Integer); dispid 29;
    procedure OnDeleteRecord(handle: Integer); dispid 30;
    procedure OnACRDataAvailable(callId: Integer; Status: Integer); dispid 31;
    procedure OnAvailableURLs(const CampaignName: WideString); dispid 32;
    procedure OnDebug(Status: Integer); dispid 33;
    procedure OnEndConferencePossible(Status: Integer); dispid 34;
    procedure OnAlternateCallPossible(Status: Integer); dispid 35;
    procedure OnDoNotCallByPhoneNumberPossible(Status: Integer); dispid 36;
    procedure OnMultiPreviewRecordsSelection(nbRecords: Integer); dispid 37;
    procedure OnAddRecordPossible(Status: Integer); dispid 38;
    procedure OnNewAddRecordPossible(nHandle: Integer); dispid 39;
  end;

// *********************************************************************//
// Interface   : IAxMakeCall
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {AFA40FF3-E319-11D5-B776-00508B79E1ED}
// *********************************************************************//
  IAxMakeCall = interface(IDispatch)
    ['{AFA40FF3-E319-11D5-B776-00508B79E1ED}']
    function  MakeCall(const destNumber: WideString; const correlator: WideString): Integer; safecall;
    function  Connect: Integer; safecall;
    function  Get_destNumber: WideString; safecall;
    procedure Set_destNumber(const pVal: WideString); safecall;
    function  Get_MakeCallPossible: Integer; safecall;
    procedure Set_DialButtonText(const Param1: WideString); safecall;
    procedure Set_ClrButtonText(const Param1: WideString); safecall;
    property destNumber: WideString read Get_destNumber write Set_destNumber;
    property MakeCallPossible: Integer read Get_MakeCallPossible;
    property DialButtonText: WideString write Set_DialButtonText;
    property ClrButtonText: WideString write Set_ClrButtonText;
  end;

// *********************************************************************//
// DispIntf:  IAxMakeCallDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {AFA40FF3-E319-11D5-B776-00508B79E1ED}
// *********************************************************************//
  IAxMakeCallDisp = dispinterface
    ['{AFA40FF3-E319-11D5-B776-00508B79E1ED}']
    function  MakeCall(const destNumber: WideString; const correlator: WideString): Integer; dispid 1;
    function  Connect: Integer; dispid 2;
    property destNumber: WideString dispid 3;
    property MakeCallPossible: Integer readonly dispid 4;
    property DialButtonText: WideString writeonly dispid 5;
    property ClrButtonText: WideString writeonly dispid 6;
  end;

// *********************************************************************//
// DispIntf:  _IAxMakeCallEvents
// Flags:     (4096) Dispatchable
// GUID:      {AFA40FF5-E319-11D5-B776-00508B79E1ED}
// *********************************************************************//
  _IAxMakeCallEvents = dispinterface
    ['{AFA40FF5-E319-11D5-B776-00508B79E1ED}']
    procedure OnMakeCallPossible(Status: Integer); dispid 1;
    procedure OnRequestStatus(requestId: Integer; Status: AxTServerErrorTypes); dispid 2;
  end;

// *********************************************************************//
// Interface   : IAxHangUp
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {AFA40FFA-E319-11D5-B776-00508B79E1ED}
// *********************************************************************//
  IAxHangUp = interface(IDispatch)
    ['{AFA40FFA-E319-11D5-B776-00508B79E1ED}']
    procedure Set_Picture(const ppPicture: IPictureDisp); safecall;
    procedure _Set_Picture(const ppPicture: IPictureDisp); safecall;
    function  Get_Picture: IPictureDisp; safecall;
    function  ReleaseCall(callId: Integer): Integer; safecall;
    function  Connect: Integer; safecall;
    function  Get_ReleaseCallPossible(callId: Integer): Integer; safecall;
    property Picture: IPictureDisp read Get_Picture write _Set_Picture;
    property ReleaseCallPossible[callId: Integer]: Integer read Get_ReleaseCallPossible;
  end;

// *********************************************************************//
// DispIntf:  IAxHangUpDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {AFA40FFA-E319-11D5-B776-00508B79E1ED}
// *********************************************************************//
  IAxHangUpDisp = dispinterface
    ['{AFA40FFA-E319-11D5-B776-00508B79E1ED}']
    property Picture: IPictureDisp dispid -523;
    function  ReleaseCall(callId: Integer): Integer; dispid 1;
    function  Connect: Integer; dispid 2;
    property ReleaseCallPossible[callId: Integer]: Integer readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf:  _IAxHangUpEvents
// Flags:     (4096) Dispatchable
// GUID:      {AFA40FFC-E319-11D5-B776-00508B79E1ED}
// *********************************************************************//
  _IAxHangUpEvents = dispinterface
    ['{AFA40FFC-E319-11D5-B776-00508B79E1ED}']
    procedure OnReleaseCallPossible(callId: Integer; Status: Integer); dispid 1;
    procedure OnRequestStatus(requestId: Integer; Status: AxTServerErrorTypes); dispid 2;
  end;

// *********************************************************************//
// Interface   : IAxTakeCall
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {AA674F6F-E801-11D5-B778-00508B79E1ED}
// *********************************************************************//
  IAxTakeCall = interface(IDispatch)
    ['{AA674F6F-E801-11D5-B778-00508B79E1ED}']
    procedure Set_Picture(const ppPicture: IPictureDisp); safecall;
    procedure _Set_Picture(const ppPicture: IPictureDisp); safecall;
    function  Get_Picture: IPictureDisp; safecall;
    function  TakeCall(callId: Integer): Integer; safecall;
    function  Connect: Integer; safecall;
    function  Get_TakeCallPossible(callId: Integer): Integer; safecall;
    function  Get_AlternateCallPossible: Integer; safecall;
    function  AlternateCall: Integer; safecall;
    property Picture: IPictureDisp read Get_Picture write _Set_Picture;
    property TakeCallPossible[callId: Integer]: Integer read Get_TakeCallPossible;
    property AlternateCallPossible: Integer read Get_AlternateCallPossible;
  end;

// *********************************************************************//
// DispIntf:  IAxTakeCallDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {AA674F6F-E801-11D5-B778-00508B79E1ED}
// *********************************************************************//
  IAxTakeCallDisp = dispinterface
    ['{AA674F6F-E801-11D5-B778-00508B79E1ED}']
    property Picture: IPictureDisp dispid -523;
    function  TakeCall(callId: Integer): Integer; dispid 1;
    function  Connect: Integer; dispid 2;
    property TakeCallPossible[callId: Integer]: Integer readonly dispid 3;
    property AlternateCallPossible: Integer readonly dispid 4;
    function  AlternateCall: Integer; dispid 5;
  end;

// *********************************************************************//
// DispIntf:  _IAxTakeCallEvents
// Flags:     (4096) Dispatchable
// GUID:      {AA674F70-E801-11D5-B778-00508B79E1ED}
// *********************************************************************//
  _IAxTakeCallEvents = dispinterface
    ['{AA674F70-E801-11D5-B778-00508B79E1ED}']
    procedure OnTakeCallPossible(callId: Integer; Status: Integer); dispid 1;
    procedure OnRequestStatus(requestId: Integer; Status: AxTServerErrorTypes); dispid 2;
    procedure OnAlternateCallPossible(Status: Integer); dispid 3;
  end;

// *********************************************************************//
// Interface   : IAxConference
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {37D5EB36-E88F-11D5-B779-00508B79E1ED}
// *********************************************************************//
  IAxConference = interface(IDispatch)
    ['{37D5EB36-E88F-11D5-B779-00508B79E1ED}']
    procedure Set_Picture(const ppPicture: IPictureDisp); safecall;
    procedure _Set_Picture(const ppPicture: IPictureDisp); safecall;
    function  Get_Picture: IPictureDisp; safecall;
    function  Conference: Integer; safecall;
    function  Connect: Integer; safecall;
    function  Get_ConferencePossible: Integer; safecall;
    function  Get_ConferenceEndPossible: Integer; safecall;
    function  ConferenceEnd: Integer; safecall;
    property Picture: IPictureDisp read Get_Picture write _Set_Picture;
    property ConferencePossible: Integer read Get_ConferencePossible;
    property ConferenceEndPossible: Integer read Get_ConferenceEndPossible;
  end;

// *********************************************************************//
// DispIntf:  IAxConferenceDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {37D5EB36-E88F-11D5-B779-00508B79E1ED}
// *********************************************************************//
  IAxConferenceDisp = dispinterface
    ['{37D5EB36-E88F-11D5-B779-00508B79E1ED}']
    property Picture: IPictureDisp dispid -523;
    function  Conference: Integer; dispid 1;
    function  Connect: Integer; dispid 2;
    property ConferencePossible: Integer readonly dispid 3;
    property ConferenceEndPossible: Integer readonly dispid 4;
    function  ConferenceEnd: Integer; dispid 5;
  end;

// *********************************************************************//
// DispIntf:  _IAxConferenceEvents
// Flags:     (4096) Dispatchable
// GUID:      {37D5EB38-E88F-11D5-B779-00508B79E1ED}
// *********************************************************************//
  _IAxConferenceEvents = dispinterface
    ['{37D5EB38-E88F-11D5-B779-00508B79E1ED}']
    procedure OnConferenceCallPossible(Status: Integer); dispid 1;
    procedure OnRequestStatus(requestId: Integer; Status: AxTServerErrorTypes); dispid 2;
    procedure OnEndConferenceCallPossible(Status: Integer); dispid 3;
  end;

// *********************************************************************//
// Interface   : IAxLog
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {7F272F84-E98E-11D5-B779-00508B79E1ED}
// *********************************************************************//
  IAxLog = interface(IDispatch)
    ['{7F272F84-E98E-11D5-B779-00508B79E1ED}']
    procedure Set_Picture(const ppPicture: IPictureDisp); safecall;
    procedure _Set_Picture(const ppPicture: IPictureDisp); safecall;
    function  Get_Picture: IPictureDisp; safecall;
    function  Logon(const Agent: WideString; const Team: WideString; const Password: WideString; 
                    outbound: Integer; multimediaTools: Integer): Integer; safecall;
    function  Logoff(const Agent: WideString; const Team: WideString; const Password: WideString): Integer; safecall;
    function  Connect: Integer; safecall;
    function  Get_LogonPossible: Integer; safecall;
    function  Get_LogoffPossible: Integer; safecall;
    procedure Set_LogonPopupWindowTitle(const Param1: WideString); safecall;
    procedure Set_LogoffPopupWindowTitle(const Param1: WideString); safecall;
    procedure Set_AgentNumberText(const Param1: WideString); safecall;
    procedure Set_PasswordText(const Param1: WideString); safecall;
    procedure Set_AssignedGroupText(const Param1: WideString); safecall;
    procedure Set_CheckboxOutboundCampaignText(const Param1: WideString); safecall;
    procedure Set_CheckBoxInternetToolsText(const Param1: WideString); safecall;
    procedure Set_CancelButtonText(const Param1: WideString); safecall;
    procedure Set_OkButtonText(const Param1: WideString); safecall;
    property Picture: IPictureDisp read Get_Picture write _Set_Picture;
    property LogonPossible: Integer read Get_LogonPossible;
    property LogoffPossible: Integer read Get_LogoffPossible;
    property LogonPopupWindowTitle: WideString write Set_LogonPopupWindowTitle;
    property LogoffPopupWindowTitle: WideString write Set_LogoffPopupWindowTitle;
    property AgentNumberText: WideString write Set_AgentNumberText;
    property PasswordText: WideString write Set_PasswordText;
    property AssignedGroupText: WideString write Set_AssignedGroupText;
    property CheckboxOutboundCampaignText: WideString write Set_CheckboxOutboundCampaignText;
    property CheckBoxInternetToolsText: WideString write Set_CheckBoxInternetToolsText;
    property CancelButtonText: WideString write Set_CancelButtonText;
    property OkButtonText: WideString write Set_OkButtonText;
  end;

// *********************************************************************//
// DispIntf:  IAxLogDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {7F272F84-E98E-11D5-B779-00508B79E1ED}
// *********************************************************************//
  IAxLogDisp = dispinterface
    ['{7F272F84-E98E-11D5-B779-00508B79E1ED}']
    property Picture: IPictureDisp dispid -523;
    function  Logon(const Agent: WideString; const Team: WideString; const Password: WideString; 
                    outbound: Integer; multimediaTools: Integer): Integer; dispid 1;
    function  Logoff(const Agent: WideString; const Team: WideString; const Password: WideString): Integer; dispid 2;
    function  Connect: Integer; dispid 3;
    property LogonPossible: Integer readonly dispid 4;
    property LogoffPossible: Integer readonly dispid 5;
    property LogonPopupWindowTitle: WideString writeonly dispid 6;
    property LogoffPopupWindowTitle: WideString writeonly dispid 7;
    property AgentNumberText: WideString writeonly dispid 8;
    property PasswordText: WideString writeonly dispid 9;
    property AssignedGroupText: WideString writeonly dispid 10;
    property CheckboxOutboundCampaignText: WideString writeonly dispid 11;
    property CheckBoxInternetToolsText: WideString writeonly dispid 12;
    property CancelButtonText: WideString writeonly dispid 13;
    property OkButtonText: WideString writeonly dispid 14;
  end;

// *********************************************************************//
// DispIntf:  _IAxLogEvents
// Flags:     (4096) Dispatchable
// GUID:      {7F272F86-E98E-11D5-B779-00508B79E1ED}
// *********************************************************************//
  _IAxLogEvents = dispinterface
    ['{7F272F86-E98E-11D5-B779-00508B79E1ED}']
    procedure OnLogonPossible(Status: Integer); dispid 1;
    procedure OnLogoffPossible(Status: Integer); dispid 2;
    procedure OnRequestStatus(requestId: Integer; Status: AxTServerErrorTypes); dispid 3;
  end;

// *********************************************************************//
// Interface   : IItemOfHistoric
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {7CB3D129-5C21-11D6-8C70-0008C7BF3D97}
// *********************************************************************//
  IItemOfHistoric = interface(IDispatch)
    ['{7CB3D129-5C21-11D6-8C70-0008C7BF3D97}']
    function  Get_IdHigh: Integer; safecall;
    function  Get_IdLow: Integer; safecall;
    function  Get_Number: WideString; safecall;
    function  Get_LastName: WideString; safecall;
    function  Get_FirstName: WideString; safecall;
    function  Get_State: AxAgentCallStates; safecall;
    function  Get_Incoming: Integer; safecall;
    function  Get_Extern: Integer; safecall;
    function  Get_BeginDate: TDateTime; safecall;
    function  Get_EndDate: TDateTime; safecall;
    function  Get_TalkingDate: TDateTime; safecall;
    property IdHigh: Integer read Get_IdHigh;
    property IdLow: Integer read Get_IdLow;
    property Number: WideString read Get_Number;
    property LastName: WideString read Get_LastName;
    property FirstName: WideString read Get_FirstName;
    property State: AxAgentCallStates read Get_State;
    property Incoming: Integer read Get_Incoming;
    property Extern: Integer read Get_Extern;
    property BeginDate: TDateTime read Get_BeginDate;
    property EndDate: TDateTime read Get_EndDate;
    property TalkingDate: TDateTime read Get_TalkingDate;
  end;

// *********************************************************************//
// DispIntf:  IItemOfHistoricDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {7CB3D129-5C21-11D6-8C70-0008C7BF3D97}
// *********************************************************************//
  IItemOfHistoricDisp = dispinterface
    ['{7CB3D129-5C21-11D6-8C70-0008C7BF3D97}']
    property IdHigh: Integer readonly dispid 1;
    property IdLow: Integer readonly dispid 2;
    property Number: WideString readonly dispid 3;
    property LastName: WideString readonly dispid 4;
    property FirstName: WideString readonly dispid 5;
    property State: AxAgentCallStates readonly dispid 6;
    property Incoming: Integer readonly dispid 7;
    property Extern: Integer readonly dispid 8;
    property BeginDate: TDateTime readonly dispid 9;
    property EndDate: TDateTime readonly dispid 10;
    property TalkingDate: TDateTime readonly dispid 11;
  end;

// *********************************************************************//
// Interface   : IAxHistoric
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {1ECE7398-EAF5-11D5-B77A-00508B79E1ED}
// *********************************************************************//
  IAxHistoric = interface(IDispatch)
    ['{1ECE7398-EAF5-11D5-B77A-00508B79E1ED}']
    function  Get_LastError: Integer; safecall;
    function  Get_NbEvtQueued: Integer; safecall;
    function  Get_NbItems: Integer; safecall;
    function  RemoveItem(IdHigh: Integer; IdLow: Integer): Integer; safecall;
    function  Connect: Integer; safecall;
    property LastError: Integer read Get_LastError;
    property NbEvtQueued: Integer read Get_NbEvtQueued;
    property NbItems: Integer read Get_NbItems;
  end;

// *********************************************************************//
// DispIntf:  IAxHistoricDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {1ECE7398-EAF5-11D5-B77A-00508B79E1ED}
// *********************************************************************//
  IAxHistoricDisp = dispinterface
    ['{1ECE7398-EAF5-11D5-B77A-00508B79E1ED}']
    property LastError: Integer readonly dispid 1;
    property NbEvtQueued: Integer readonly dispid 2;
    property NbItems: Integer readonly dispid 3;
    function  RemoveItem(IdHigh: Integer; IdLow: Integer): Integer; dispid 4;
    function  Connect: Integer; dispid 5;
  end;

// *********************************************************************//
// DispIntf:  _IAxHistoricEvents
// Flags:     (4096) Dispatchable
// GUID:      {1ECE739A-EAF5-11D5-B77A-00508B79E1ED}
// *********************************************************************//
  _IAxHistoricEvents = dispinterface
    ['{1ECE739A-EAF5-11D5-B77A-00508B79E1ED}']
    procedure OnHistoDeleteItem(IdHigh: Integer; IdLow: Integer); dispid 1;
    procedure OnRequestStatus(requestId: Integer; Status: AxTServerErrorTypes); dispid 2;
  end;

// *********************************************************************//
// Interface   : IAxTransfertCall
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {C989EA69-F2C8-11D5-B77B-00508B79E1ED}
// *********************************************************************//
  IAxTransfertCall = interface(IDispatch)
    ['{C989EA69-F2C8-11D5-B77B-00508B79E1ED}']
    procedure Set_Picture(const ppPicture: IPictureDisp); safecall;
    procedure _Set_Picture(const ppPicture: IPictureDisp); safecall;
    function  Get_Picture: IPictureDisp; safecall;
    function  TransferCall(heldCallId: Integer): Integer; safecall;
    function  Connect: Integer; safecall;
    function  Get_TransferCallPossible: Integer; safecall;
    property Picture: IPictureDisp read Get_Picture write _Set_Picture;
    property TransferCallPossible: Integer read Get_TransferCallPossible;
  end;

// *********************************************************************//
// DispIntf:  IAxTransfertCallDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C989EA69-F2C8-11D5-B77B-00508B79E1ED}
// *********************************************************************//
  IAxTransfertCallDisp = dispinterface
    ['{C989EA69-F2C8-11D5-B77B-00508B79E1ED}']
    property Picture: IPictureDisp dispid -523;
    function  TransferCall(heldCallId: Integer): Integer; dispid 1;
    function  Connect: Integer; dispid 2;
    property TransferCallPossible: Integer readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf:  _IAxTransfertCallEvents
// Flags:     (4096) Dispatchable
// GUID:      {C989EA6B-F2C8-11D5-B77B-00508B79E1ED}
// *********************************************************************//
  _IAxTransfertCallEvents = dispinterface
    ['{C989EA6B-F2C8-11D5-B77B-00508B79E1ED}']
    procedure OnTransferCallPossible(Status: Integer); dispid 1;
    procedure OnRequestStatus(requestId: Integer; Status: AxTServerErrorTypes); dispid 2;
  end;

// *********************************************************************//
// Interface   : IAxStartStopRecording
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {E810B9BA-F45D-11D5-B77E-00508B79E1ED}
// *********************************************************************//
  IAxStartStopRecording = interface(IDispatch)
    ['{E810B9BA-F45D-11D5-B77E-00508B79E1ED}']
    procedure Set_Picture(const ppPicture: IPictureDisp); safecall;
    procedure _Set_Picture(const ppPicture: IPictureDisp); safecall;
    function  Get_Picture: IPictureDisp; safecall;
    function  Get_IsNiceStartRecordPossible: Integer; safecall;
    function  Get_IsNiceStopRecordPossible: Integer; safecall;
    function  Get_NbEvtQueued: Integer; safecall;
    function  NiceStartRecord(const Comment: WideString): Integer; safecall;
    function  NiceStopRecord(const Comment: WideString): Integer; safecall;
    function  UpdateCall(Comment1: OleVariant; Comment2: OleVariant; Comment3: OleVariant): Integer; safecall;
    function  Connect: Integer; safecall;
    property Picture: IPictureDisp read Get_Picture write _Set_Picture;
    property IsNiceStartRecordPossible: Integer read Get_IsNiceStartRecordPossible;
    property IsNiceStopRecordPossible: Integer read Get_IsNiceStopRecordPossible;
    property NbEvtQueued: Integer read Get_NbEvtQueued;
  end;

// *********************************************************************//
// DispIntf:  IAxStartStopRecordingDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {E810B9BA-F45D-11D5-B77E-00508B79E1ED}
// *********************************************************************//
  IAxStartStopRecordingDisp = dispinterface
    ['{E810B9BA-F45D-11D5-B77E-00508B79E1ED}']
    property Picture: IPictureDisp dispid -523;
    property IsNiceStartRecordPossible: Integer readonly dispid 1;
    property IsNiceStopRecordPossible: Integer readonly dispid 2;
    property NbEvtQueued: Integer readonly dispid 3;
    function  NiceStartRecord(const Comment: WideString): Integer; dispid 4;
    function  NiceStopRecord(const Comment: WideString): Integer; dispid 5;
    function  UpdateCall(Comment1: OleVariant; Comment2: OleVariant; Comment3: OleVariant): Integer; dispid 6;
    function  Connect: Integer; dispid 7;
  end;

// *********************************************************************//
// DispIntf:  _IAxStartStopRecordingEvents
// Flags:     (4096) Dispatchable
// GUID:      {E810B9BC-F45D-11D5-B77E-00508B79E1ED}
// *********************************************************************//
  _IAxStartStopRecordingEvents = dispinterface
    ['{E810B9BC-F45D-11D5-B77E-00508B79E1ED}']
    procedure OnNiceStartRecordPossible(Status: Integer); dispid 1;
    procedure OnNiceStopRecordPossible(Status: Integer); dispid 2;
  end;

// *********************************************************************//
// Interface   : IAxAgentStatistics
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {47B38D57-F9ED-11D5-B77F-00508B79E1ED}
// *********************************************************************//
  IAxAgentStatistics = interface(IDispatch)
    ['{47B38D57-F9ED-11D5-B77F-00508B79E1ED}']
    function  Get_ServiceState: Integer; safecall;
    function  Get_ServiceStateTime: Integer; safecall;
    function  Get_ServiceStateDate: TDateTime; safecall;
    function  Get_LogOnTime: Integer; safecall;
    function  Get_LogOnDate: TDateTime; safecall;
    function  Get_AgentState: Integer; safecall;
    function  Get_AgentStateTime: Integer; safecall;
    function  Get_AgentStateDate: TDateTime; safecall;
    function  Get_NumberOfWithdraws: Integer; safecall;
    function  Get_WithdrawsDuration: Integer; safecall;
    function  Get_NumberOfACDCalls: Integer; safecall;
    function  Get_NumberOfPrivateCalls: Integer; safecall;
    function  Get_PrivateCallsDuration: Integer; safecall;
    function  Get_NumberOfInterceptedACDCalls: Integer; safecall;
    function  Get_NumberOfTransferredACDCalls: Integer; safecall;
    function  Get_NumberOfRefusedACDCalls: Integer; safecall;
    function  Get_ServedPilotDirNumber: WideString; safecall;
    function  Get_NbEvtQueued: Integer; safecall;
    function  Get_LastError: Integer; safecall;
    function  Connect: Integer; safecall;
    procedure Set_AgentNameText(const Param1: WideString); safecall;
    procedure Set_GeneralStatisticsText(const Param1: WideString); safecall;
    procedure Set_LogonText(const Param1: WideString); safecall;
    procedure Set_LogonTimeText(const Param1: WideString); safecall;
    procedure Set_LogonDurationText(const Param1: WideString); safecall;
    procedure Set_WithdrawalsText(const Param1: WideString); safecall;
    procedure Set_WithdrawalsCountText(const Param1: WideString); safecall;
    procedure Set_WithdrawalsDurationText(const Param1: WideString); safecall;
    procedure Set_CallsStatisticsText(const Param1: WideString); safecall;
    procedure Set_PrivatesText(const Param1: WideString); safecall;
    procedure Set_PrivatesCountText(const Param1: WideString); safecall;
    procedure Set_PrivatesDurationText(const Param1: WideString); safecall;
    procedure Set_AcdText(const Param1: WideString); safecall;
    procedure Set_AcdProcessedText(const Param1: WideString); safecall;
    procedure Set_AcdRefusedText(const Param1: WideString); safecall;
    procedure Set_AcdPickedUpText(const Param1: WideString); safecall;
    procedure Set_AcdTransferredText(const Param1: WideString); safecall;
    property ServiceState: Integer read Get_ServiceState;
    property ServiceStateTime: Integer read Get_ServiceStateTime;
    property ServiceStateDate: TDateTime read Get_ServiceStateDate;
    property LogOnTime: Integer read Get_LogOnTime;
    property LogOnDate: TDateTime read Get_LogOnDate;
    property AgentState: Integer read Get_AgentState;
    property AgentStateTime: Integer read Get_AgentStateTime;
    property AgentStateDate: TDateTime read Get_AgentStateDate;
    property NumberOfWithdraws: Integer read Get_NumberOfWithdraws;
    property WithdrawsDuration: Integer read Get_WithdrawsDuration;
    property NumberOfACDCalls: Integer read Get_NumberOfACDCalls;
    property NumberOfPrivateCalls: Integer read Get_NumberOfPrivateCalls;
    property PrivateCallsDuration: Integer read Get_PrivateCallsDuration;
    property NumberOfInterceptedACDCalls: Integer read Get_NumberOfInterceptedACDCalls;
    property NumberOfTransferredACDCalls: Integer read Get_NumberOfTransferredACDCalls;
    property NumberOfRefusedACDCalls: Integer read Get_NumberOfRefusedACDCalls;
    property ServedPilotDirNumber: WideString read Get_ServedPilotDirNumber;
    property NbEvtQueued: Integer read Get_NbEvtQueued;
    property LastError: Integer read Get_LastError;
    property AgentNameText: WideString write Set_AgentNameText;
    property GeneralStatisticsText: WideString write Set_GeneralStatisticsText;
    property LogonText: WideString write Set_LogonText;
    property LogonTimeText: WideString write Set_LogonTimeText;
    property LogonDurationText: WideString write Set_LogonDurationText;
    property WithdrawalsText: WideString write Set_WithdrawalsText;
    property WithdrawalsCountText: WideString write Set_WithdrawalsCountText;
    property WithdrawalsDurationText: WideString write Set_WithdrawalsDurationText;
    property CallsStatisticsText: WideString write Set_CallsStatisticsText;
    property PrivatesText: WideString write Set_PrivatesText;
    property PrivatesCountText: WideString write Set_PrivatesCountText;
    property PrivatesDurationText: WideString write Set_PrivatesDurationText;
    property AcdText: WideString write Set_AcdText;
    property AcdProcessedText: WideString write Set_AcdProcessedText;
    property AcdRefusedText: WideString write Set_AcdRefusedText;
    property AcdPickedUpText: WideString write Set_AcdPickedUpText;
    property AcdTransferredText: WideString write Set_AcdTransferredText;
  end;

// *********************************************************************//
// DispIntf:  IAxAgentStatisticsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {47B38D57-F9ED-11D5-B77F-00508B79E1ED}
// *********************************************************************//
  IAxAgentStatisticsDisp = dispinterface
    ['{47B38D57-F9ED-11D5-B77F-00508B79E1ED}']
    property ServiceState: Integer readonly dispid 1;
    property ServiceStateTime: Integer readonly dispid 2;
    property ServiceStateDate: TDateTime readonly dispid 3;
    property LogOnTime: Integer readonly dispid 4;
    property LogOnDate: TDateTime readonly dispid 5;
    property AgentState: Integer readonly dispid 6;
    property AgentStateTime: Integer readonly dispid 7;
    property AgentStateDate: TDateTime readonly dispid 8;
    property NumberOfWithdraws: Integer readonly dispid 9;
    property WithdrawsDuration: Integer readonly dispid 10;
    property NumberOfACDCalls: Integer readonly dispid 11;
    property NumberOfPrivateCalls: Integer readonly dispid 12;
    property PrivateCallsDuration: Integer readonly dispid 13;
    property NumberOfInterceptedACDCalls: Integer readonly dispid 14;
    property NumberOfTransferredACDCalls: Integer readonly dispid 15;
    property NumberOfRefusedACDCalls: Integer readonly dispid 16;
    property ServedPilotDirNumber: WideString readonly dispid 17;
    property NbEvtQueued: Integer readonly dispid 18;
    property LastError: Integer readonly dispid 19;
    function  Connect: Integer; dispid 20;
    property AgentNameText: WideString writeonly dispid 21;
    property GeneralStatisticsText: WideString writeonly dispid 22;
    property LogonText: WideString writeonly dispid 23;
    property LogonTimeText: WideString writeonly dispid 24;
    property LogonDurationText: WideString writeonly dispid 25;
    property WithdrawalsText: WideString writeonly dispid 26;
    property WithdrawalsCountText: WideString writeonly dispid 27;
    property WithdrawalsDurationText: WideString writeonly dispid 28;
    property CallsStatisticsText: WideString writeonly dispid 29;
    property PrivatesText: WideString writeonly dispid 30;
    property PrivatesCountText: WideString writeonly dispid 31;
    property PrivatesDurationText: WideString writeonly dispid 32;
    property AcdText: WideString writeonly dispid 33;
    property AcdProcessedText: WideString writeonly dispid 34;
    property AcdRefusedText: WideString writeonly dispid 35;
    property AcdPickedUpText: WideString writeonly dispid 36;
    property AcdTransferredText: WideString writeonly dispid 37;
  end;

// *********************************************************************//
// DispIntf:  _IAxAgentStatisticsEvents
// Flags:     (4096) Dispatchable
// GUID:      {47B38D59-F9ED-11D5-B77F-00508B79E1ED}
// *********************************************************************//
  _IAxAgentStatisticsEvents = dispinterface
    ['{47B38D59-F9ED-11D5-B77F-00508B79E1ED}']
    procedure OnPropertyChanges(Status: AxPCAgentErrorTypes; changes: AxACDTeamStateChanges); dispid 1;
  end;

// *********************************************************************//
// Interface   : IAxTeamStatistics
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {94EDB881-FA19-11D5-B77F-00508B79E1ED}
// *********************************************************************//
  IAxTeamStatistics = interface(IDispatch)
    ['{94EDB881-FA19-11D5-B77F-00508B79E1ED}']
    function  Get_NumberOfLoggedAgents: Integer; safecall;
    function  Get_NumberOfAffectedAgents: Integer; safecall;
    function  Get_NumberOfWithdrawedAgents: Integer; safecall;
    function  Get_NumberOfBusyAgents: Integer; safecall;
    function  Get_NumberOfFreeAgents: Integer; safecall;
    function  Get_NumberOfACDCalls: Integer; safecall;
    function  Get_NumberOfPrivateCalls: Integer; safecall;
    function  Get_NumberOfWaitingCalls: Integer; safecall;
    function  Get_MaximumWaitingTime: Integer; safecall;
    function  Get_NbEvtQueued: Integer; safecall;
    function  Get_LastError: Integer; safecall;
    function  Get_AverageWaitingTime: Single; safecall;
    function  Connect: Integer; safecall;
    procedure Set_GeneralStatisticsText(const Param1: WideString); safecall;
    procedure Set_ProcessingGroupText(const Param1: WideString); safecall;
    procedure Set_WaitText(const Param1: WideString); safecall;
    procedure Set_WaitMaxText(const Param1: WideString); safecall;
    procedure Set_WaitAverageText(const Param1: WideString); safecall;
    procedure Set_CallsText(const Param1: WideString); safecall;
    procedure Set_CallsAcdText(const Param1: WideString); safecall;
    procedure Set_CallsPrivateText(const Param1: WideString); safecall;
    procedure Set_CallsWaitText(const Param1: WideString); safecall;
    procedure Set_AgentStatisticsText(const Param1: WideString); safecall;
    procedure Set_LoggedText(const Param1: WideString); safecall;
    procedure Set_SetText(const Param1: WideString); safecall;
    procedure Set_WithdrawsText(const Param1: WideString); safecall;
    procedure Set_BusyText(const Param1: WideString); safecall;
    procedure Set_IdleText(const Param1: WideString); safecall;
    function  Get_MonitoredGroupNumber: WideString; safecall;
    procedure Set_MonitoredGroupNumber(const pVal: WideString); safecall;
    property NumberOfLoggedAgents: Integer read Get_NumberOfLoggedAgents;
    property NumberOfAffectedAgents: Integer read Get_NumberOfAffectedAgents;
    property NumberOfWithdrawedAgents: Integer read Get_NumberOfWithdrawedAgents;
    property NumberOfBusyAgents: Integer read Get_NumberOfBusyAgents;
    property NumberOfFreeAgents: Integer read Get_NumberOfFreeAgents;
    property NumberOfACDCalls: Integer read Get_NumberOfACDCalls;
    property NumberOfPrivateCalls: Integer read Get_NumberOfPrivateCalls;
    property NumberOfWaitingCalls: Integer read Get_NumberOfWaitingCalls;
    property MaximumWaitingTime: Integer read Get_MaximumWaitingTime;
    property NbEvtQueued: Integer read Get_NbEvtQueued;
    property LastError: Integer read Get_LastError;
    property AverageWaitingTime: Single read Get_AverageWaitingTime;
    property GeneralStatisticsText: WideString write Set_GeneralStatisticsText;
    property ProcessingGroupText: WideString write Set_ProcessingGroupText;
    property WaitText: WideString write Set_WaitText;
    property WaitMaxText: WideString write Set_WaitMaxText;
    property WaitAverageText: WideString write Set_WaitAverageText;
    property CallsText: WideString write Set_CallsText;
    property CallsAcdText: WideString write Set_CallsAcdText;
    property CallsPrivateText: WideString write Set_CallsPrivateText;
    property CallsWaitText: WideString write Set_CallsWaitText;
    property AgentStatisticsText: WideString write Set_AgentStatisticsText;
    property LoggedText: WideString write Set_LoggedText;
    property SetText: WideString write Set_SetText;
    property WithdrawsText: WideString write Set_WithdrawsText;
    property BusyText: WideString write Set_BusyText;
    property IdleText: WideString write Set_IdleText;
    property MonitoredGroupNumber: WideString read Get_MonitoredGroupNumber write Set_MonitoredGroupNumber;
  end;

// *********************************************************************//
// DispIntf:  IAxTeamStatisticsDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {94EDB881-FA19-11D5-B77F-00508B79E1ED}
// *********************************************************************//
  IAxTeamStatisticsDisp = dispinterface
    ['{94EDB881-FA19-11D5-B77F-00508B79E1ED}']
    property NumberOfLoggedAgents: Integer readonly dispid 1;
    property NumberOfAffectedAgents: Integer readonly dispid 2;
    property NumberOfWithdrawedAgents: Integer readonly dispid 3;
    property NumberOfBusyAgents: Integer readonly dispid 4;
    property NumberOfFreeAgents: Integer readonly dispid 5;
    property NumberOfACDCalls: Integer readonly dispid 6;
    property NumberOfPrivateCalls: Integer readonly dispid 7;
    property NumberOfWaitingCalls: Integer readonly dispid 8;
    property MaximumWaitingTime: Integer readonly dispid 9;
    property NbEvtQueued: Integer readonly dispid 10;
    property LastError: Integer readonly dispid 11;
    property AverageWaitingTime: Single readonly dispid 12;
    function  Connect: Integer; dispid 13;
    property GeneralStatisticsText: WideString writeonly dispid 14;
    property ProcessingGroupText: WideString writeonly dispid 15;
    property WaitText: WideString writeonly dispid 16;
    property WaitMaxText: WideString writeonly dispid 17;
    property WaitAverageText: WideString writeonly dispid 18;
    property CallsText: WideString writeonly dispid 19;
    property CallsAcdText: WideString writeonly dispid 20;
    property CallsPrivateText: WideString writeonly dispid 21;
    property CallsWaitText: WideString writeonly dispid 22;
    property AgentStatisticsText: WideString writeonly dispid 23;
    property LoggedText: WideString writeonly dispid 24;
    property SetText: WideString writeonly dispid 25;
    property WithdrawsText: WideString writeonly dispid 26;
    property BusyText: WideString writeonly dispid 27;
    property IdleText: WideString writeonly dispid 28;
    property MonitoredGroupNumber: WideString dispid 29;
  end;

// *********************************************************************//
// DispIntf:  _IAxTeamStatisticsEvents
// Flags:     (4096) Dispatchable
// GUID:      {94EDB883-FA19-11D5-B77F-00508B79E1ED}
// *********************************************************************//
  _IAxTeamStatisticsEvents = dispinterface
    ['{94EDB883-FA19-11D5-B77F-00508B79E1ED}']
    procedure OnPropertyChanges(Status: AxPCAgentErrorTypes; changes: AxACDTeamStateChanges); dispid 1;
  end;

// *********************************************************************//
// Interface   : IAxHold
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {4BB01D3F-035A-11D6-B782-00508B79E1ED}
// *********************************************************************//
  IAxHold = interface(IDispatch)
    ['{4BB01D3F-035A-11D6-B782-00508B79E1ED}']
    procedure Set_Picture(const ppPicture: IPictureDisp); safecall;
    procedure _Set_Picture(const ppPicture: IPictureDisp); safecall;
    function  Get_Picture: IPictureDisp; safecall;
    function  HoldCall: Integer; safecall;
    function  Connect: Integer; safecall;
    function  Get_HoldCallPossible: Integer; safecall;
    property Picture: IPictureDisp read Get_Picture write _Set_Picture;
    property HoldCallPossible: Integer read Get_HoldCallPossible;
  end;

// *********************************************************************//
// DispIntf:  IAxHoldDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {4BB01D3F-035A-11D6-B782-00508B79E1ED}
// *********************************************************************//
  IAxHoldDisp = dispinterface
    ['{4BB01D3F-035A-11D6-B782-00508B79E1ED}']
    property Picture: IPictureDisp dispid -523;
    function  HoldCall: Integer; dispid 1;
    function  Connect: Integer; dispid 2;
    property HoldCallPossible: Integer readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf:  _IAxHoldEvents
// Flags:     (4096) Dispatchable
// GUID:      {4BB01D41-035A-11D6-B782-00508B79E1ED}
// *********************************************************************//
  _IAxHoldEvents = dispinterface
    ['{4BB01D41-035A-11D6-B782-00508B79E1ED}']
    procedure OnHoldCallPossible(Status: Integer); dispid 1;
    procedure OnRequestStatus(requestId: Integer; Status: AxTServerErrorTypes); dispid 2;
  end;

// *********************************************************************//
// Interface   : IAxDoNotCallRecord
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {114C6FA3-04FC-11D6-B785-00508B79E1ED}
// *********************************************************************//
  IAxDoNotCallRecord = interface(IDispatch)
    ['{114C6FA3-04FC-11D6-B785-00508B79E1ED}']
    procedure Set_Picture(const ppPicture: IPictureDisp); safecall;
    procedure _Set_Picture(const ppPicture: IPictureDisp); safecall;
    function  Get_Picture: IPictureDisp; safecall;
    function  DoNotCall(handle: Integer): Integer; safecall;
    function  Connect: Integer; safecall;
    function  Get_DoNotCallRecord: Integer; safecall;
    property Picture: IPictureDisp read Get_Picture write _Set_Picture;
    property DoNotCallRecord: Integer read Get_DoNotCallRecord;
  end;

// *********************************************************************//
// DispIntf:  IAxDoNotCallRecordDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {114C6FA3-04FC-11D6-B785-00508B79E1ED}
// *********************************************************************//
  IAxDoNotCallRecordDisp = dispinterface
    ['{114C6FA3-04FC-11D6-B785-00508B79E1ED}']
    property Picture: IPictureDisp dispid -523;
    function  DoNotCall(handle: Integer): Integer; dispid 1;
    function  Connect: Integer; dispid 2;
    property DoNotCallRecord: Integer readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf:  _IAxDoNotCallRecordEvents
// Flags:     (4096) Dispatchable
// GUID:      {114C6FA5-04FC-11D6-B785-00508B79E1ED}
// *********************************************************************//
  _IAxDoNotCallRecordEvents = dispinterface
    ['{114C6FA5-04FC-11D6-B785-00508B79E1ED}']
    procedure OnDoNotCallPossible(Status: Integer); dispid 1;
    procedure OnRequestStatus(requestId: Integer; Status: AxTServerErrorTypes); dispid 2;
  end;

// *********************************************************************//
// Interface   : IAxRequestRejectPreviewRecord
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {50A0753E-0513-11D6-B785-00508B79E1ED}
// *********************************************************************//
  IAxRequestRejectPreviewRecord = interface(IDispatch)
    ['{50A0753E-0513-11D6-B785-00508B79E1ED}']
    procedure Set_Picture(const ppPicture: IPictureDisp); safecall;
    procedure _Set_Picture(const ppPicture: IPictureDisp); safecall;
    function  Get_Picture: IPictureDisp; safecall;
    function  PreviewRecord: Integer; safecall;
    function  RecordReject(handle: Integer): Integer; safecall;
    function  Connect: Integer; safecall;
    function  Get_RequestPreviewRecordPossible: Integer; safecall;
    function  Get_RejectPreviewRecordPossible: Integer; safecall;
    property Picture: IPictureDisp read Get_Picture write _Set_Picture;
    property RequestPreviewRecordPossible: Integer read Get_RequestPreviewRecordPossible;
    property RejectPreviewRecordPossible: Integer read Get_RejectPreviewRecordPossible;
  end;

// *********************************************************************//
// DispIntf:  IAxRequestRejectPreviewRecordDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {50A0753E-0513-11D6-B785-00508B79E1ED}
// *********************************************************************//
  IAxRequestRejectPreviewRecordDisp = dispinterface
    ['{50A0753E-0513-11D6-B785-00508B79E1ED}']
    property Picture: IPictureDisp dispid -523;
    function  PreviewRecord: Integer; dispid 1;
    function  RecordReject(handle: Integer): Integer; dispid 2;
    function  Connect: Integer; dispid 3;
    property RequestPreviewRecordPossible: Integer readonly dispid 4;
    property RejectPreviewRecordPossible: Integer readonly dispid 5;
  end;

// *********************************************************************//
// DispIntf:  _IAxRequestRejectPreviewRecordEvents
// Flags:     (4096) Dispatchable
// GUID:      {50A07540-0513-11D6-B785-00508B79E1ED}
// *********************************************************************//
  _IAxRequestRejectPreviewRecordEvents = dispinterface
    ['{50A07540-0513-11D6-B785-00508B79E1ED}']
    procedure OnPreviewRecordPossible(Status: Integer); dispid 1;
    procedure OnRecordRejectPossible(Status: Integer); dispid 2;
    procedure OnRequestStatus(requestId: Integer; Status: AxTServerErrorTypes); dispid 3;
  end;

// *********************************************************************//
// Interface   : IAxScheduleCallback
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {50A0755E-0513-11D6-B785-00508B79E1ED}
// *********************************************************************//
  IAxScheduleCallback = interface(IDispatch)
    ['{50A0755E-0513-11D6-B785-00508B79E1ED}']
    procedure Set_Picture(const ppPicture: IPictureDisp); safecall;
    procedure _Set_Picture(const ppPicture: IPictureDisp); safecall;
    function  Get_Picture: IPictureDisp; safecall;
    function  Connect: Integer; safecall;
    procedure Set_CallbackTypesList(const Param1: WideString); safecall;
    procedure Set_CallbackDateFormat(const Param1: WideString); safecall;
    function  RescheduleRecord(handle: Integer; const callbackType: WideString; 
                               callbackDate: OleVariant): Integer; safecall;
    function  Get_ScheduleCallbackPossible: Integer; safecall;
    procedure Set_PopupWindowTitle(const Param1: WideString); safecall;
    procedure Set_CallbackTypeText(const Param1: WideString); safecall;
    procedure Set_CallbackDateText(const Param1: WideString); safecall;
    procedure Set_OkButtonText(const Param1: WideString); safecall;
    procedure Set_CancelButtonText(const Param1: WideString); safecall;
    property Picture: IPictureDisp read Get_Picture write _Set_Picture;
    property CallbackTypesList: WideString write Set_CallbackTypesList;
    property CallbackDateFormat: WideString write Set_CallbackDateFormat;
    property ScheduleCallbackPossible: Integer read Get_ScheduleCallbackPossible;
    property PopupWindowTitle: WideString write Set_PopupWindowTitle;
    property CallbackTypeText: WideString write Set_CallbackTypeText;
    property CallbackDateText: WideString write Set_CallbackDateText;
    property OkButtonText: WideString write Set_OkButtonText;
    property CancelButtonText: WideString write Set_CancelButtonText;
  end;

// *********************************************************************//
// DispIntf:  IAxScheduleCallbackDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {50A0755E-0513-11D6-B785-00508B79E1ED}
// *********************************************************************//
  IAxScheduleCallbackDisp = dispinterface
    ['{50A0755E-0513-11D6-B785-00508B79E1ED}']
    property Picture: IPictureDisp dispid -523;
    function  Connect: Integer; dispid 1;
    property CallbackTypesList: WideString writeonly dispid 2;
    property CallbackDateFormat: WideString writeonly dispid 3;
    function  RescheduleRecord(handle: Integer; const callbackType: WideString; 
                               callbackDate: OleVariant): Integer; dispid 5;
    property ScheduleCallbackPossible: Integer readonly dispid 6;
    property PopupWindowTitle: WideString writeonly dispid 7;
    property CallbackTypeText: WideString writeonly dispid 8;
    property CallbackDateText: WideString writeonly dispid 9;
    property OkButtonText: WideString writeonly dispid 10;
    property CancelButtonText: WideString writeonly dispid 11;
  end;

// *********************************************************************//
// DispIntf:  _IAxScheduleCallbackEvents
// Flags:     (4096) Dispatchable
// GUID:      {50A07560-0513-11D6-B785-00508B79E1ED}
// *********************************************************************//
  _IAxScheduleCallbackEvents = dispinterface
    ['{50A07560-0513-11D6-B785-00508B79E1ED}']
    procedure OnRecordReschedulePossible(Status: Integer); dispid 1;
    procedure OnRequestStatus(requestId: Integer; Status: AxTServerErrorTypes); dispid 2;
  end;

// *********************************************************************//
// Interface   : IAxClassify
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {50A0757C-0513-11D6-B785-00508B79E1ED}
// *********************************************************************//
  IAxClassify = interface(IDispatch)
    ['{50A0757C-0513-11D6-B785-00508B79E1ED}']
    procedure Set_Picture(const ppPicture: IPictureDisp); safecall;
    procedure _Set_Picture(const ppPicture: IPictureDisp); safecall;
    function  Get_Picture: IPictureDisp; safecall;
    function  RecordProcessed(handle: Integer): Integer; safecall;
    function  Connect: Integer; safecall;
    function  Get_RecordToClassify: Integer; safecall;
    property Picture: IPictureDisp read Get_Picture write _Set_Picture;
    property RecordToClassify: Integer read Get_RecordToClassify;
  end;

// *********************************************************************//
// DispIntf:  IAxClassifyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {50A0757C-0513-11D6-B785-00508B79E1ED}
// *********************************************************************//
  IAxClassifyDisp = dispinterface
    ['{50A0757C-0513-11D6-B785-00508B79E1ED}']
    property Picture: IPictureDisp dispid -523;
    function  RecordProcessed(handle: Integer): Integer; dispid 1;
    function  Connect: Integer; dispid 2;
    property RecordToClassify: Integer readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf:  _IAxClassifyEvents
// Flags:     (4096) Dispatchable
// GUID:      {50A0757E-0513-11D6-B785-00508B79E1ED}
// *********************************************************************//
  _IAxClassifyEvents = dispinterface
    ['{50A0757E-0513-11D6-B785-00508B79E1ED}']
    procedure OnRecordProcessedPossible(Status: Integer); dispid 1;
    procedure OnRequestStatus(requestId: Integer; Status: AxTServerErrorTypes); dispid 2;
  end;


// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TAxApplication
// Chaîne d'aide        : AxApplication Class
// Interface par défaut : IAxApplication
// DISP Int. Déf. ?     : No
// Interface événements : _IAxApplicationEvents
// TypeFlags            : (2) CanCreate
// *********************************************************************//
  TAxApplicationOnServerStatus = procedure(Sender: TObject; Status: AxServerStatus) of object;
  TAxApplicationOnVisible = procedure(Sender: TObject; Status: AxVisibleStatus) of object;

  TAxApplication = class(TOleControl)
  private
    FOnQuit: TNotifyEvent;
    FOnServerStatus: TAxApplicationOnServerStatus;
    FOnVisible: TAxApplicationOnVisible;
    FIntf: IAxApplication;
    function  GetControlInterface: IAxApplication;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    function  Connect: Integer;
    property  ControlInterface: IAxApplication read GetControlInterface;
    property  DefaultInterface: IAxApplication read GetControlInterface;
    property StatServerName: WideString index 1 read GetWideStringProp;
    property Version: WideString index 2 read GetWideStringProp;
    property Time: TDateTime index 3 read GetTDateTimeProp;
    property ServerStatus: TOleEnum index 5 read GetTOleEnumProp;
    property PhoneServerName: WideString index 6 read GetWideStringProp;
    property PhoneServerPort: Integer index 7 read GetIntegerProp;
    property StatServerSite: WideString index 8 read GetWideStringProp;
    property NbEvtQueued: Integer index 9 read GetIntegerProp;
    property ServiceOption: Integer index 10 read GetIntegerProp;
    property FirstName: WideString index 11 read GetWideStringProp;
    property LastName: WideString index 12 read GetWideStringProp;
    property LoginName: WideString index 13 read GetWideStringProp;
    property StatServerPort: Integer index 15 read GetIntegerProp;
  published
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property Visible: Integer index 4 read GetIntegerProp write SetIntegerProp stored False;
    property CCagentVisibility: TOleEnum index 16 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property OnQuit: TNotifyEvent read FOnQuit write FOnQuit;
    property OnServerStatus: TAxApplicationOnServerStatus read FOnServerStatus write FOnServerStatus;
    property OnVisible: TAxApplicationOnVisible read FOnVisible write FOnVisible;
  end;

// *********************************************************************//
// La classe CoDataOfCall fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut IDataOfCall exposée             
// pas la CoClass DataOfCall. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClass exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoDataOfCall = class
    class function Create: IDataOfCall;
    class function CreateRemote(const MachineName: string): IDataOfCall;
  end;


// *********************************************************************//
// Déclaration de classe proxy de serveur OLE
// Objet serveur        : TDataOfCall
// Chaîne d'aide        : DataOfCall Class
// Interface par défaut : IDataOfCall
// DISP Int. Déf. ?     : No
// Interface événements : 
// TypeFlags            : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TDataOfCallProperties= class;
{$ENDIF}
  TDataOfCall = class(TOleServer)
  private
    FIntf:        IDataOfCall;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TDataOfCallProperties;
    function      GetServerProperties: TDataOfCallProperties;
{$ENDIF}
    function      GetDefaultInterface: IDataOfCall;
  protected
    procedure InitServerData; override;
    function  Get_Id: Integer;
    function  Get_LastName: WideString;
    function  Get_FirstName: WideString;
    function  Get_Number: WideString;
    function  Get_Incoming: Integer;
    function  Get_Extern: Integer;
    function  Get_DestinationNumber: WideString;
    function  Get_ConferenceMaster: WideString;
    function  Get_State: AxAgentCallStates;
    function  Get_Cause: AxAgentCallCauses;
    function  Get_ACDCall: Integer;
    function  Get_TrunkUsed: WideString;
    function  Get_EstimatedWaitingTime: Integer;
    function  Get_PilotStatus: AxPilotStatus;
    function  Get_PossibleTransfer: Integer;
    function  Get_WaitingQueueSaturation: Integer;
    function  Get_LastRedirectionDevice: WideString;
    function  Get_CorrelatorData: WideString;
    procedure Set_CorrelatorData(const pVal: WideString);
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: IDataOfCall);
    procedure Disconnect; override;
    property  DefaultInterface: IDataOfCall read GetDefaultInterface;
    property Id: Integer read Get_Id;
    property LastName: WideString read Get_LastName;
    property FirstName: WideString read Get_FirstName;
    property Number: WideString read Get_Number;
    property Incoming: Integer read Get_Incoming;
    property Extern: Integer read Get_Extern;
    property DestinationNumber: WideString read Get_DestinationNumber;
    property ConferenceMaster: WideString read Get_ConferenceMaster;
    property State: AxAgentCallStates read Get_State;
    property Cause: AxAgentCallCauses read Get_Cause;
    property ACDCall: Integer read Get_ACDCall;
    property TrunkUsed: WideString read Get_TrunkUsed;
    property EstimatedWaitingTime: Integer read Get_EstimatedWaitingTime;
    property PilotStatus: AxPilotStatus read Get_PilotStatus;
    property PossibleTransfer: Integer read Get_PossibleTransfer;
    property WaitingQueueSaturation: Integer read Get_WaitingQueueSaturation;
    property LastRedirectionDevice: WideString read Get_LastRedirectionDevice;
    property CorrelatorData: WideString read Get_CorrelatorData write Set_CorrelatorData;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TDataOfCallProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// Classe proxy Propriétés Serveur OLE
// Objet serveur    : TDataOfCall
// (Cet objet est utilisé pas l'inspecteur de propriété de l'EDI pour la 
//  modification des propriétés de ce serveur)
// *********************************************************************//
 TDataOfCallProperties = class(TPersistent)
  private
    FServer:    TDataOfCall;
    function    GetDefaultInterface: IDataOfCall;
    constructor Create(AServer: TDataOfCall);
  protected
    function  Get_Id: Integer;
    function  Get_LastName: WideString;
    function  Get_FirstName: WideString;
    function  Get_Number: WideString;
    function  Get_Incoming: Integer;
    function  Get_Extern: Integer;
    function  Get_DestinationNumber: WideString;
    function  Get_ConferenceMaster: WideString;
    function  Get_State: AxAgentCallStates;
    function  Get_Cause: AxAgentCallCauses;
    function  Get_ACDCall: Integer;
    function  Get_TrunkUsed: WideString;
    function  Get_EstimatedWaitingTime: Integer;
    function  Get_PilotStatus: AxPilotStatus;
    function  Get_PossibleTransfer: Integer;
    function  Get_WaitingQueueSaturation: Integer;
    function  Get_LastRedirectionDevice: WideString;
    function  Get_CorrelatorData: WideString;
    procedure Set_CorrelatorData(const pVal: WideString);
  public
    property DefaultInterface: IDataOfCall read GetDefaultInterface;
  published
    property CorrelatorData: WideString read Get_CorrelatorData write Set_CorrelatorData;
  end;
{$ENDIF}



// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TAxLine
// Chaîne d'aide        : AxLine Class
// Interface par défaut : IAxLine
// DISP Int. Déf. ?     : No
// Interface événements : _IAxLineEvents
// TypeFlags            : (2) CanCreate
// *********************************************************************//
  TAxLineOnLineNewCall = procedure(Sender: TObject; const pCall: IDataOfCall; callType: AxCallType) of object;
  TAxLineOnLineChangeCall = procedure(Sender: TObject; const pCall: IDataOfCall) of object;
  TAxLineOnLineDeleteCall = procedure(Sender: TObject; callId: Integer) of object;
  TAxLineOnMakeCallPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnTakeCallPossible = procedure(Sender: TObject; callId: Integer; Status: Integer) of object;
  TAxLineOnReleaseCallPossible = procedure(Sender: TObject; callId: Integer; Status: Integer) of object;
  TAxLineOnTransferCallPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnConferenceCallPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnHoldCallPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnRequestStatus = procedure(Sender: TObject; requestId: Integer; 
                                                      Status: AxTServerErrorTypes) of object;
  TAxLineOnLineStateChange = procedure(Sender: TObject; State: Integer) of object;
  TAxLineOnLogonPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnLogoffPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnWrapUpPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnWrapUpEndPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnPausePossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnPauseEndPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnWithdrawPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnWithdrawEndPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnCampaignModeChange = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnPreviewRecordPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnRecordCancelPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnRecordRejectPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnDoNotCallPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnRecordProcessedPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnUpdateCallCompStatsPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnRecordReschedulePossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnChainedRecordPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnNewRecord = procedure(Sender: TObject; handle: Integer) of object;
  TAxLineOnDeleteRecord = procedure(Sender: TObject; handle: Integer) of object;
  TAxLineOnACRDataAvailable = procedure(Sender: TObject; callId: Integer; Status: Integer) of object;
  TAxLineOnAvailableURLs = procedure(Sender: TObject; const CampaignName: WideString) of object;
  TAxLineOnDebug = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnEndConferencePossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnAlternateCallPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnDoNotCallByPhoneNumberPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnMultiPreviewRecordsSelection = procedure(Sender: TObject; nbRecords: Integer) of object;
  TAxLineOnAddRecordPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLineOnNewAddRecordPossible = procedure(Sender: TObject; nHandle: Integer) of object;

  TAxLine = class(TOleControl)
  private
    FOnLineNewCall: TAxLineOnLineNewCall;
    FOnLineChangeCall: TAxLineOnLineChangeCall;
    FOnLineDeleteCall: TAxLineOnLineDeleteCall;
    FOnMakeCallPossible: TAxLineOnMakeCallPossible;
    FOnTakeCallPossible: TAxLineOnTakeCallPossible;
    FOnReleaseCallPossible: TAxLineOnReleaseCallPossible;
    FOnTransferCallPossible: TAxLineOnTransferCallPossible;
    FOnConferenceCallPossible: TAxLineOnConferenceCallPossible;
    FOnHoldCallPossible: TAxLineOnHoldCallPossible;
    FOnRequestStatus: TAxLineOnRequestStatus;
    FOnLineStateChange: TAxLineOnLineStateChange;
    FOnLogonPossible: TAxLineOnLogonPossible;
    FOnLogoffPossible: TAxLineOnLogoffPossible;
    FOnWrapUpPossible: TAxLineOnWrapUpPossible;
    FOnWrapUpEndPossible: TAxLineOnWrapUpEndPossible;
    FOnPausePossible: TAxLineOnPausePossible;
    FOnPauseEndPossible: TAxLineOnPauseEndPossible;
    FOnWithdrawPossible: TAxLineOnWithdrawPossible;
    FOnWithdrawEndPossible: TAxLineOnWithdrawEndPossible;
    FOnCampaignModeChange: TAxLineOnCampaignModeChange;
    FOnPreviewRecordPossible: TAxLineOnPreviewRecordPossible;
    FOnRecordCancelPossible: TAxLineOnRecordCancelPossible;
    FOnRecordRejectPossible: TAxLineOnRecordRejectPossible;
    FOnDoNotCallPossible: TAxLineOnDoNotCallPossible;
    FOnRecordProcessedPossible: TAxLineOnRecordProcessedPossible;
    FOnUpdateCallCompStatsPossible: TAxLineOnUpdateCallCompStatsPossible;
    FOnRecordReschedulePossible: TAxLineOnRecordReschedulePossible;
    FOnChainedRecordPossible: TAxLineOnChainedRecordPossible;
    FOnNewRecord: TAxLineOnNewRecord;
    FOnDeleteRecord: TAxLineOnDeleteRecord;
    FOnACRDataAvailable: TAxLineOnACRDataAvailable;
    FOnAvailableURLs: TAxLineOnAvailableURLs;
    FOnDebug: TAxLineOnDebug;
    FOnEndConferencePossible: TAxLineOnEndConferencePossible;
    FOnAlternateCallPossible: TAxLineOnAlternateCallPossible;
    FOnDoNotCallByPhoneNumberPossible: TAxLineOnDoNotCallByPhoneNumberPossible;
    FOnMultiPreviewRecordsSelection: TAxLineOnMultiPreviewRecordsSelection;
    FOnAddRecordPossible: TAxLineOnAddRecordPossible;
    FOnNewAddRecordPossible: TAxLineOnNewAddRecordPossible;
    FIntf: IAxLine;
    function  GetControlInterface: IAxLine;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
    function  Get_AffectedTeamDirNumber(nTeam: SYSINT): WideString;
    function  Get_ContactDataKey(handle: Integer; field: SYSINT): WideString;
    function  Get_ContactDataValue(handle: Integer; const key: WideString): OleVariant;
    procedure Set_ContactDataValue(handle: Integer; const key: WideString; pVal: OleVariant);
    function  Get_ContactDataVisible(handle: Integer): Integer;
    procedure Set_ContactDataVisible(handle: Integer; pVal: Integer);
    function  Get_IsChainedRecordPossible(handle: Integer): Integer;
    function  Get_IsDoNotCallPossible(handle: Integer): Integer;
    function  Get_IsRecordCancelPossible(handle: Integer): Integer;
    function  Get_IsRecordProcessedPossible(handle: Integer): Integer;
    function  Get_IsRecordRejectPossible(handle: Integer): Integer;
    function  Get_IsRecordReschedulePossible(handle: Integer): Integer;
    function  Get_IsReleaseCallPossible(callId: Integer): Integer;
    function  Get_IsTakeCallPossible(callId: Integer): Integer;
    function  Get_IsUpdateCallCompStatsPossible(handle: Integer): Integer;
    function  Get_NumberOfRecordFields(handle: Integer): Integer;
    function  Get_WithdrawCause(nCause: Integer): WideString;
    function  Get_CallById(callId: Integer): IDataOfCall;
    function  Get_CallByIndex(CallIndex: Integer): IDataOfCall;
  public
    function  ChainedRecord(handle: Integer): Integer;
    function  Conference: Integer;
    function  DoNotCall(handle: Integer): Integer;
    function  HoldCall: Integer;
    function  Logoff(const Agent: WideString; const Team: WideString; const Password: WideString): Integer;
    function  Logon(const Agent: WideString; const Team: WideString; const Password: WideString; 
                    outbound: Integer; multimediaTools: Integer): Integer;
    function  MakeCall(const destNumber: WideString; const correlator: WideString): Integer;
    function  Pause: Integer;
    function  PauseEnd: Integer;
    function  PreviewRecord: Integer;
    function  RecordCancel(handle: Integer): Integer;
    function  RecordProcessed(handle: Integer): Integer;
    function  RecordReject(handle: Integer): Integer;
    function  RecordReschedule(handle: Integer): Integer;
    function  ReleaseCall(callId: Integer): Integer;
    function  SetDateAndTime(const szDate: WideString): Integer;
    function  TakeCall(callId: Integer): Integer;
    function  TransferCall(heldCallId: Integer): Integer;
    function  UpdateCallCompletionStats(handle: Integer): Integer;
    function  Withdraw(Cause: Integer): Integer;
    function  WithdrawEnd: Integer;
    function  WrapUp: Integer;
    function  WrapUpEnd: Integer;
    procedure ShowEvents(Enabled: Integer);
    function  Connect: Integer;
    function  EndConference: Integer;
    function  IsEndConferencePossible: Integer;
    function  ContactAdd(const pContactNumber: WideString): Integer;
    function  ContactSave(nHandle: Integer): Integer;
    function  ContactCancel(nHandle: Integer): Integer;
    function  IsAddContactPossible: Integer;
    property  ControlInterface: IAxLine read GetControlInterface;
    property  DefaultInterface: IAxLine read GetControlInterface;
    property ActiveCall: Integer index 1 read GetIntegerProp;
    property AffectedTeamDirNumber[nTeam: SYSINT]: WideString read Get_AffectedTeamDirNumber;
    property ContactDataKey[handle: Integer; field: SYSINT]: WideString read Get_ContactDataKey;
    property ContactDataValue[handle: Integer; const key: WideString]: OleVariant read Get_ContactDataValue write Set_ContactDataValue;
    property CampaignMode: Integer index 5 read GetIntegerProp;
    property ContactDataVisible[handle: Integer]: Integer read Get_ContactDataVisible write Set_ContactDataVisible;
    property IsChainedRecordPossible[handle: Integer]: Integer read Get_IsChainedRecordPossible;
    property IsConferencePossible: Integer index 8 read GetIntegerProp;
    property IsDoNotCallPossible[handle: Integer]: Integer read Get_IsDoNotCallPossible;
    property IsHoldCallPossible: Integer index 10 read GetIntegerProp;
    property IsLogoffPossible: Integer index 11 read GetIntegerProp;
    property IsLogonPossible: Integer index 12 read GetIntegerProp;
    property IsMakeCallPossible: Integer index 13 read GetIntegerProp;
    property IsPauseEndPossible: Integer index 14 read GetIntegerProp;
    property IsPausePossible: Integer index 15 read GetIntegerProp;
    property IsPreviewRecordPossible: Integer index 16 read GetIntegerProp;
    property IsRecordCancelPossible[handle: Integer]: Integer read Get_IsRecordCancelPossible;
    property IsRecordProcessedPossible[handle: Integer]: Integer read Get_IsRecordProcessedPossible;
    property IsRecordRejectPossible[handle: Integer]: Integer read Get_IsRecordRejectPossible;
    property IsRecordReschedulePossible[handle: Integer]: Integer read Get_IsRecordReschedulePossible;
    property IsReleaseCallPossible[callId: Integer]: Integer read Get_IsReleaseCallPossible;
    property IsTakeCallPossible[callId: Integer]: Integer read Get_IsTakeCallPossible;
    property IsTransferCallPossible: Integer index 23 read GetIntegerProp;
    property IsUpdateCallCompStatsPossible[handle: Integer]: Integer read Get_IsUpdateCallCompStatsPossible;
    property IsWithdrawEndPossible: Integer index 25 read GetIntegerProp;
    property IsWithdrawPossible: Integer index 26 read GetIntegerProp;
    property IsWrapUpEndPossible: Integer index 27 read GetIntegerProp;
    property IsWrapUpPossible: Integer index 28 read GetIntegerProp;
    property LastError: TOleEnum index 29 read GetTOleEnumProp;
    property LineServiceState: TOleEnum index 30 read GetTOleEnumProp;
    property LoggedTeam: WideString index 31 read GetWideStringProp;
    property NbCalls: Integer index 32 read GetIntegerProp;
    property NbEvtQueued: Integer index 33 read GetIntegerProp;
    property Number: WideString index 34 read GetWideStringProp;
    property NumberOfAffectedTeam: Integer index 35 read GetIntegerProp;
    property NumberOfRecordFields[handle: Integer]: Integer read Get_NumberOfRecordFields;
    property NumberOfWithdrawCauses: Integer index 37 read GetIntegerProp;
    property WithdrawCause[nCause: Integer]: WideString read Get_WithdrawCause;
    property HandleOfCurrentRecord: Integer index 63 read GetIntegerProp;
    property CurrentReservationType: TOleEnum index 66 read GetTOleEnumProp;
    property CallById[callId: Integer]: IDataOfCall read Get_CallById;
    property CallByIndex[CallIndex: Integer]: IDataOfCall read Get_CallByIndex;
  published
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  Visible;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property ProACDExtension: WideString index 38 read GetWideStringProp write SetWideStringProp stored False;
    property OnLineNewCall: TAxLineOnLineNewCall read FOnLineNewCall write FOnLineNewCall;
    property OnLineChangeCall: TAxLineOnLineChangeCall read FOnLineChangeCall write FOnLineChangeCall;
    property OnLineDeleteCall: TAxLineOnLineDeleteCall read FOnLineDeleteCall write FOnLineDeleteCall;
    property OnMakeCallPossible: TAxLineOnMakeCallPossible read FOnMakeCallPossible write FOnMakeCallPossible;
    property OnTakeCallPossible: TAxLineOnTakeCallPossible read FOnTakeCallPossible write FOnTakeCallPossible;
    property OnReleaseCallPossible: TAxLineOnReleaseCallPossible read FOnReleaseCallPossible write FOnReleaseCallPossible;
    property OnTransferCallPossible: TAxLineOnTransferCallPossible read FOnTransferCallPossible write FOnTransferCallPossible;
    property OnConferenceCallPossible: TAxLineOnConferenceCallPossible read FOnConferenceCallPossible write FOnConferenceCallPossible;
    property OnHoldCallPossible: TAxLineOnHoldCallPossible read FOnHoldCallPossible write FOnHoldCallPossible;
    property OnRequestStatus: TAxLineOnRequestStatus read FOnRequestStatus write FOnRequestStatus;
    property OnLineStateChange: TAxLineOnLineStateChange read FOnLineStateChange write FOnLineStateChange;
    property OnLogonPossible: TAxLineOnLogonPossible read FOnLogonPossible write FOnLogonPossible;
    property OnLogoffPossible: TAxLineOnLogoffPossible read FOnLogoffPossible write FOnLogoffPossible;
    property OnWrapUpPossible: TAxLineOnWrapUpPossible read FOnWrapUpPossible write FOnWrapUpPossible;
    property OnWrapUpEndPossible: TAxLineOnWrapUpEndPossible read FOnWrapUpEndPossible write FOnWrapUpEndPossible;
    property OnPausePossible: TAxLineOnPausePossible read FOnPausePossible write FOnPausePossible;
    property OnPauseEndPossible: TAxLineOnPauseEndPossible read FOnPauseEndPossible write FOnPauseEndPossible;
    property OnWithdrawPossible: TAxLineOnWithdrawPossible read FOnWithdrawPossible write FOnWithdrawPossible;
    property OnWithdrawEndPossible: TAxLineOnWithdrawEndPossible read FOnWithdrawEndPossible write FOnWithdrawEndPossible;
    property OnCampaignModeChange: TAxLineOnCampaignModeChange read FOnCampaignModeChange write FOnCampaignModeChange;
    property OnPreviewRecordPossible: TAxLineOnPreviewRecordPossible read FOnPreviewRecordPossible write FOnPreviewRecordPossible;
    property OnRecordCancelPossible: TAxLineOnRecordCancelPossible read FOnRecordCancelPossible write FOnRecordCancelPossible;
    property OnRecordRejectPossible: TAxLineOnRecordRejectPossible read FOnRecordRejectPossible write FOnRecordRejectPossible;
    property OnDoNotCallPossible: TAxLineOnDoNotCallPossible read FOnDoNotCallPossible write FOnDoNotCallPossible;
    property OnRecordProcessedPossible: TAxLineOnRecordProcessedPossible read FOnRecordProcessedPossible write FOnRecordProcessedPossible;
    property OnUpdateCallCompStatsPossible: TAxLineOnUpdateCallCompStatsPossible read FOnUpdateCallCompStatsPossible write FOnUpdateCallCompStatsPossible;
    property OnRecordReschedulePossible: TAxLineOnRecordReschedulePossible read FOnRecordReschedulePossible write FOnRecordReschedulePossible;
    property OnChainedRecordPossible: TAxLineOnChainedRecordPossible read FOnChainedRecordPossible write FOnChainedRecordPossible;
    property OnNewRecord: TAxLineOnNewRecord read FOnNewRecord write FOnNewRecord;
    property OnDeleteRecord: TAxLineOnDeleteRecord read FOnDeleteRecord write FOnDeleteRecord;
    property OnACRDataAvailable: TAxLineOnACRDataAvailable read FOnACRDataAvailable write FOnACRDataAvailable;
    property OnAvailableURLs: TAxLineOnAvailableURLs read FOnAvailableURLs write FOnAvailableURLs;
    property OnDebug: TAxLineOnDebug read FOnDebug write FOnDebug;
    property OnEndConferencePossible: TAxLineOnEndConferencePossible read FOnEndConferencePossible write FOnEndConferencePossible;
    property OnAlternateCallPossible: TAxLineOnAlternateCallPossible read FOnAlternateCallPossible write FOnAlternateCallPossible;
    property OnDoNotCallByPhoneNumberPossible: TAxLineOnDoNotCallByPhoneNumberPossible read FOnDoNotCallByPhoneNumberPossible write FOnDoNotCallByPhoneNumberPossible;
    property OnMultiPreviewRecordsSelection: TAxLineOnMultiPreviewRecordsSelection read FOnMultiPreviewRecordsSelection write FOnMultiPreviewRecordsSelection;
    property OnAddRecordPossible: TAxLineOnAddRecordPossible read FOnAddRecordPossible write FOnAddRecordPossible;
    property OnNewAddRecordPossible: TAxLineOnNewAddRecordPossible read FOnNewAddRecordPossible write FOnNewAddRecordPossible;
  end;


// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TAxMakeCall
// Chaîne d'aide        : AxMakeCall Class
// Interface par défaut : IAxMakeCall
// DISP Int. Déf. ?     : No
// Interface événements : _IAxMakeCallEvents
// TypeFlags            : (2) CanCreate
// *********************************************************************//
  TAxMakeCallOnMakeCallPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxMakeCallOnRequestStatus = procedure(Sender: TObject; requestId: Integer; 
                                                          Status: AxTServerErrorTypes) of object;

  TAxMakeCall = class(TOleControl)
  private
    FOnMakeCallPossible: TAxMakeCallOnMakeCallPossible;
    FOnRequestStatus: TAxMakeCallOnRequestStatus;
    FIntf: IAxMakeCall;
    function  GetControlInterface: IAxMakeCall;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    function  MakeCall(const destNumber: WideString; const correlator: WideString): Integer;
    function  Connect: Integer;
    property  ControlInterface: IAxMakeCall read GetControlInterface;
    property  DefaultInterface: IAxMakeCall read GetControlInterface;
    property MakeCallPossible: Integer index 4 read GetIntegerProp;
    property DialButtonText: WideString index 5 write SetWideStringProp;
    property ClrButtonText: WideString index 6 write SetWideStringProp;
  published
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  Visible;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property destNumber: WideString index 3 read GetWideStringProp write SetWideStringProp stored False;
    property OnMakeCallPossible: TAxMakeCallOnMakeCallPossible read FOnMakeCallPossible write FOnMakeCallPossible;
    property OnRequestStatus: TAxMakeCallOnRequestStatus read FOnRequestStatus write FOnRequestStatus;
  end;


// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TAxHangUp
// Chaîne d'aide        : AxHangUp Class
// Interface par défaut : IAxHangUp
// DISP Int. Déf. ?     : No
// Interface événements : _IAxHangUpEvents
// TypeFlags            : (2) CanCreate
// *********************************************************************//
  TAxHangUpOnReleaseCallPossible = procedure(Sender: TObject; callId: Integer; Status: Integer) of object;
  TAxHangUpOnRequestStatus = procedure(Sender: TObject; requestId: Integer; 
                                                        Status: AxTServerErrorTypes) of object;

  TAxHangUp = class(TOleControl)
  private
    FOnReleaseCallPossible: TAxHangUpOnReleaseCallPossible;
    FOnRequestStatus: TAxHangUpOnRequestStatus;
    FIntf: IAxHangUp;
    function  GetControlInterface: IAxHangUp;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
    function  Get_ReleaseCallPossible(callId: Integer): Integer;
  public
    function  ReleaseCall(callId: Integer): Integer;
    function  Connect: Integer;
    property  ControlInterface: IAxHangUp read GetControlInterface;
    property  DefaultInterface: IAxHangUp read GetControlInterface;
    property ReleaseCallPossible[callId: Integer]: Integer read Get_ReleaseCallPossible;
  published
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  Visible;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property Picture: TPicture index -523 read GetTPictureProp write SetTPictureProp stored False;
    property OnReleaseCallPossible: TAxHangUpOnReleaseCallPossible read FOnReleaseCallPossible write FOnReleaseCallPossible;
    property OnRequestStatus: TAxHangUpOnRequestStatus read FOnRequestStatus write FOnRequestStatus;
  end;


// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TAxTakeCall
// Chaîne d'aide        : AxTakeCall Class
// Interface par défaut : IAxTakeCall
// DISP Int. Déf. ?     : No
// Interface événements : _IAxTakeCallEvents
// TypeFlags            : (2) CanCreate
// *********************************************************************//
  TAxTakeCallOnTakeCallPossible = procedure(Sender: TObject; callId: Integer; Status: Integer) of object;
  TAxTakeCallOnRequestStatus = procedure(Sender: TObject; requestId: Integer; 
                                                          Status: AxTServerErrorTypes) of object;
  TAxTakeCallOnAlternateCallPossible = procedure(Sender: TObject; Status: Integer) of object;

  TAxTakeCall = class(TOleControl)
  private
    FOnTakeCallPossible: TAxTakeCallOnTakeCallPossible;
    FOnRequestStatus: TAxTakeCallOnRequestStatus;
    FOnAlternateCallPossible: TAxTakeCallOnAlternateCallPossible;
    FIntf: IAxTakeCall;
    function  GetControlInterface: IAxTakeCall;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
    function  Get_TakeCallPossible(callId: Integer): Integer;
  public
    function  TakeCall(callId: Integer): Integer;
    function  Connect: Integer;
    function  AlternateCall: Integer;
    property  ControlInterface: IAxTakeCall read GetControlInterface;
    property  DefaultInterface: IAxTakeCall read GetControlInterface;
    property TakeCallPossible[callId: Integer]: Integer read Get_TakeCallPossible;
    property AlternateCallPossible: Integer index 4 read GetIntegerProp;
  published
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  Visible;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property Picture: TPicture index -523 read GetTPictureProp write SetTPictureProp stored False;
    property OnTakeCallPossible: TAxTakeCallOnTakeCallPossible read FOnTakeCallPossible write FOnTakeCallPossible;
    property OnRequestStatus: TAxTakeCallOnRequestStatus read FOnRequestStatus write FOnRequestStatus;
    property OnAlternateCallPossible: TAxTakeCallOnAlternateCallPossible read FOnAlternateCallPossible write FOnAlternateCallPossible;
  end;


// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TAxConference
// Chaîne d'aide        : AxConference Class
// Interface par défaut : IAxConference
// DISP Int. Déf. ?     : No
// Interface événements : _IAxConferenceEvents
// TypeFlags            : (2) CanCreate
// *********************************************************************//
  TAxConferenceOnConferenceCallPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxConferenceOnRequestStatus = procedure(Sender: TObject; requestId: Integer; 
                                                            Status: AxTServerErrorTypes) of object;
  TAxConferenceOnEndConferenceCallPossible = procedure(Sender: TObject; Status: Integer) of object;

  TAxConference = class(TOleControl)
  private
    FOnConferenceCallPossible: TAxConferenceOnConferenceCallPossible;
    FOnRequestStatus: TAxConferenceOnRequestStatus;
    FOnEndConferenceCallPossible: TAxConferenceOnEndConferenceCallPossible;
    FIntf: IAxConference;
    function  GetControlInterface: IAxConference;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    function  Conference: Integer;
    function  Connect: Integer;
    function  ConferenceEnd: Integer;
    property  ControlInterface: IAxConference read GetControlInterface;
    property  DefaultInterface: IAxConference read GetControlInterface;
    property ConferencePossible: Integer index 3 read GetIntegerProp;
    property ConferenceEndPossible: Integer index 4 read GetIntegerProp;
  published
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  Visible;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property Picture: TPicture index -523 read GetTPictureProp write SetTPictureProp stored False;
    property OnConferenceCallPossible: TAxConferenceOnConferenceCallPossible read FOnConferenceCallPossible write FOnConferenceCallPossible;
    property OnRequestStatus: TAxConferenceOnRequestStatus read FOnRequestStatus write FOnRequestStatus;
    property OnEndConferenceCallPossible: TAxConferenceOnEndConferenceCallPossible read FOnEndConferenceCallPossible write FOnEndConferenceCallPossible;
  end;


// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TAxLog
// Chaîne d'aide        : AxLog Class
// Interface par défaut : IAxLog
// DISP Int. Déf. ?     : No
// Interface événements : _IAxLogEvents
// TypeFlags            : (2) CanCreate
// *********************************************************************//
  TAxLogOnLogonPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLogOnLogoffPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxLogOnRequestStatus = procedure(Sender: TObject; requestId: Integer; Status: AxTServerErrorTypes) of object;

  TAxLog = class(TOleControl)
  private
    FOnLogonPossible: TAxLogOnLogonPossible;
    FOnLogoffPossible: TAxLogOnLogoffPossible;
    FOnRequestStatus: TAxLogOnRequestStatus;
    FIntf: IAxLog;
    function  GetControlInterface: IAxLog;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    function  Logon(const Agent: WideString; const Team: WideString; const Password: WideString; 
                    outbound: Integer; multimediaTools: Integer): Integer;
    function  Logoff(const Agent: WideString; const Team: WideString; const Password: WideString): Integer;
    function  Connect: Integer;
    property  ControlInterface: IAxLog read GetControlInterface;
    property  DefaultInterface: IAxLog read GetControlInterface;
    property LogonPossible: Integer index 4 read GetIntegerProp;
    property LogoffPossible: Integer index 5 read GetIntegerProp;
    property LogonPopupWindowTitle: WideString index 6 write SetWideStringProp;
    property LogoffPopupWindowTitle: WideString index 7 write SetWideStringProp;
    property AgentNumberText: WideString index 8 write SetWideStringProp;
    property PasswordText: WideString index 9 write SetWideStringProp;
    property AssignedGroupText: WideString index 10 write SetWideStringProp;
    property CheckboxOutboundCampaignText: WideString index 11 write SetWideStringProp;
    property CheckBoxInternetToolsText: WideString index 12 write SetWideStringProp;
    property CancelButtonText: WideString index 13 write SetWideStringProp;
    property OkButtonText: WideString index 14 write SetWideStringProp;
  published
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  Visible;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property Picture: TPicture index -523 read GetTPictureProp write SetTPictureProp stored False;
    property OnLogonPossible: TAxLogOnLogonPossible read FOnLogonPossible write FOnLogonPossible;
    property OnLogoffPossible: TAxLogOnLogoffPossible read FOnLogoffPossible write FOnLogoffPossible;
    property OnRequestStatus: TAxLogOnRequestStatus read FOnRequestStatus write FOnRequestStatus;
  end;

// *********************************************************************//
// La classe CoItemOfHistoric fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut IItemOfHistoric exposée             
// pas la CoClass ItemOfHistoric. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClass exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoItemOfHistoric = class
    class function Create: IItemOfHistoric;
    class function CreateRemote(const MachineName: string): IItemOfHistoric;
  end;


// *********************************************************************//
// Déclaration de classe proxy de serveur OLE
// Objet serveur        : TItemOfHistoric
// Chaîne d'aide        : HistoricItem Class
// Interface par défaut : IItemOfHistoric
// DISP Int. Déf. ?     : No
// Interface événements : 
// TypeFlags            : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TItemOfHistoricProperties= class;
{$ENDIF}
  TItemOfHistoric = class(TOleServer)
  private
    FIntf:        IItemOfHistoric;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TItemOfHistoricProperties;
    function      GetServerProperties: TItemOfHistoricProperties;
{$ENDIF}
    function      GetDefaultInterface: IItemOfHistoric;
  protected
    procedure InitServerData; override;
    function  Get_IdHigh: Integer;
    function  Get_IdLow: Integer;
    function  Get_Number: WideString;
    function  Get_LastName: WideString;
    function  Get_FirstName: WideString;
    function  Get_State: AxAgentCallStates;
    function  Get_Incoming: Integer;
    function  Get_Extern: Integer;
    function  Get_BeginDate: TDateTime;
    function  Get_EndDate: TDateTime;
    function  Get_TalkingDate: TDateTime;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: IItemOfHistoric);
    procedure Disconnect; override;
    property  DefaultInterface: IItemOfHistoric read GetDefaultInterface;
    property IdHigh: Integer read Get_IdHigh;
    property IdLow: Integer read Get_IdLow;
    property Number: WideString read Get_Number;
    property LastName: WideString read Get_LastName;
    property FirstName: WideString read Get_FirstName;
    property State: AxAgentCallStates read Get_State;
    property Incoming: Integer read Get_Incoming;
    property Extern: Integer read Get_Extern;
    property BeginDate: TDateTime read Get_BeginDate;
    property EndDate: TDateTime read Get_EndDate;
    property TalkingDate: TDateTime read Get_TalkingDate;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TItemOfHistoricProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// Classe proxy Propriétés Serveur OLE
// Objet serveur    : TItemOfHistoric
// (Cet objet est utilisé pas l'inspecteur de propriété de l'EDI pour la 
//  modification des propriétés de ce serveur)
// *********************************************************************//
 TItemOfHistoricProperties = class(TPersistent)
  private
    FServer:    TItemOfHistoric;
    function    GetDefaultInterface: IItemOfHistoric;
    constructor Create(AServer: TItemOfHistoric);
  protected
    function  Get_IdHigh: Integer;
    function  Get_IdLow: Integer;
    function  Get_Number: WideString;
    function  Get_LastName: WideString;
    function  Get_FirstName: WideString;
    function  Get_State: AxAgentCallStates;
    function  Get_Incoming: Integer;
    function  Get_Extern: Integer;
    function  Get_BeginDate: TDateTime;
    function  Get_EndDate: TDateTime;
    function  Get_TalkingDate: TDateTime;
  public
    property DefaultInterface: IItemOfHistoric read GetDefaultInterface;
  published
  end;
{$ENDIF}



// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TAxHistoric
// Chaîne d'aide        : AxHistoric Class
// Interface par défaut : IAxHistoric
// DISP Int. Déf. ?     : No
// Interface événements : _IAxHistoricEvents
// TypeFlags            : (2) CanCreate
// *********************************************************************//
  TAxHistoricOnHistoDeleteItem = procedure(Sender: TObject; IdHigh: Integer; IdLow: Integer) of object;
  TAxHistoricOnRequestStatus = procedure(Sender: TObject; requestId: Integer; 
                                                          Status: AxTServerErrorTypes) of object;

  TAxHistoric = class(TOleControl)
  private
    FOnHistoDeleteItem: TAxHistoricOnHistoDeleteItem;
    FOnRequestStatus: TAxHistoricOnRequestStatus;
    FIntf: IAxHistoric;
    function  GetControlInterface: IAxHistoric;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    function  RemoveItem(IdHigh: Integer; IdLow: Integer): Integer;
    function  Connect: Integer;
    property  ControlInterface: IAxHistoric read GetControlInterface;
    property  DefaultInterface: IAxHistoric read GetControlInterface;
    property LastError: Integer index 1 read GetIntegerProp;
    property NbEvtQueued: Integer index 2 read GetIntegerProp;
    property NbItems: Integer index 3 read GetIntegerProp;
  published
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  Visible;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property OnHistoDeleteItem: TAxHistoricOnHistoDeleteItem read FOnHistoDeleteItem write FOnHistoDeleteItem;
    property OnRequestStatus: TAxHistoricOnRequestStatus read FOnRequestStatus write FOnRequestStatus;
  end;


// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TAxTransfertCall
// Chaîne d'aide        : AxTransfertCall Class
// Interface par défaut : IAxTransfertCall
// DISP Int. Déf. ?     : No
// Interface événements : _IAxTransfertCallEvents
// TypeFlags            : (2) CanCreate
// *********************************************************************//
  TAxTransfertCallOnTransferCallPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxTransfertCallOnRequestStatus = procedure(Sender: TObject; requestId: Integer; 
                                                               Status: AxTServerErrorTypes) of object;

  TAxTransfertCall = class(TOleControl)
  private
    FOnTransferCallPossible: TAxTransfertCallOnTransferCallPossible;
    FOnRequestStatus: TAxTransfertCallOnRequestStatus;
    FIntf: IAxTransfertCall;
    function  GetControlInterface: IAxTransfertCall;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    function  TransferCall(heldCallId: Integer): Integer;
    function  Connect: Integer;
    property  ControlInterface: IAxTransfertCall read GetControlInterface;
    property  DefaultInterface: IAxTransfertCall read GetControlInterface;
    property TransferCallPossible: Integer index 3 read GetIntegerProp;
  published
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  Visible;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property Picture: TPicture index -523 read GetTPictureProp write SetTPictureProp stored False;
    property OnTransferCallPossible: TAxTransfertCallOnTransferCallPossible read FOnTransferCallPossible write FOnTransferCallPossible;
    property OnRequestStatus: TAxTransfertCallOnRequestStatus read FOnRequestStatus write FOnRequestStatus;
  end;


// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TAxStartStopRecording
// Chaîne d'aide        : AxStartStopRecording Class
// Interface par défaut : IAxStartStopRecording
// DISP Int. Déf. ?     : No
// Interface événements : _IAxStartStopRecordingEvents
// TypeFlags            : (2) CanCreate
// *********************************************************************//
  TAxStartStopRecordingOnNiceStartRecordPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxStartStopRecordingOnNiceStopRecordPossible = procedure(Sender: TObject; Status: Integer) of object;

  TAxStartStopRecording = class(TOleControl)
  private
    FOnNiceStartRecordPossible: TAxStartStopRecordingOnNiceStartRecordPossible;
    FOnNiceStopRecordPossible: TAxStartStopRecordingOnNiceStopRecordPossible;
    FIntf: IAxStartStopRecording;
    function  GetControlInterface: IAxStartStopRecording;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    function  NiceStartRecord(const Comment: WideString): Integer;
    function  NiceStopRecord(const Comment: WideString): Integer;
    function  UpdateCall(Comment1: OleVariant): Integer; overload;
    function  UpdateCall(Comment1: OleVariant; Comment2: OleVariant): Integer; overload;
    function  UpdateCall(Comment1: OleVariant; Comment2: OleVariant; Comment3: OleVariant): Integer; overload;
    function  Connect: Integer;
    property  ControlInterface: IAxStartStopRecording read GetControlInterface;
    property  DefaultInterface: IAxStartStopRecording read GetControlInterface;
    property IsNiceStartRecordPossible: Integer index 1 read GetIntegerProp;
    property IsNiceStopRecordPossible: Integer index 2 read GetIntegerProp;
    property NbEvtQueued: Integer index 3 read GetIntegerProp;
  published
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  Visible;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property Picture: TPicture index -523 read GetTPictureProp write SetTPictureProp stored False;
    property OnNiceStartRecordPossible: TAxStartStopRecordingOnNiceStartRecordPossible read FOnNiceStartRecordPossible write FOnNiceStartRecordPossible;
    property OnNiceStopRecordPossible: TAxStartStopRecordingOnNiceStopRecordPossible read FOnNiceStopRecordPossible write FOnNiceStopRecordPossible;
  end;


// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TAxAgentStatistics
// Chaîne d'aide        : AxAgentStatistics Class
// Interface par défaut : IAxAgentStatistics
// DISP Int. Déf. ?     : No
// Interface événements : _IAxAgentStatisticsEvents
// TypeFlags            : (2) CanCreate
// *********************************************************************//
  TAxAgentStatisticsOnPropertyChanges = procedure(Sender: TObject; Status: AxPCAgentErrorTypes; 
                                                                   changes: AxACDTeamStateChanges) of object;

  TAxAgentStatistics = class(TOleControl)
  private
    FOnPropertyChanges: TAxAgentStatisticsOnPropertyChanges;
    FIntf: IAxAgentStatistics;
    function  GetControlInterface: IAxAgentStatistics;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    function  Connect: Integer;
    property  ControlInterface: IAxAgentStatistics read GetControlInterface;
    property  DefaultInterface: IAxAgentStatistics read GetControlInterface;
    property ServiceState: Integer index 1 read GetIntegerProp;
    property ServiceStateTime: Integer index 2 read GetIntegerProp;
    property ServiceStateDate: TDateTime index 3 read GetTDateTimeProp;
    property LogOnTime: Integer index 4 read GetIntegerProp;
    property LogOnDate: TDateTime index 5 read GetTDateTimeProp;
    property AgentState: Integer index 6 read GetIntegerProp;
    property AgentStateTime: Integer index 7 read GetIntegerProp;
    property AgentStateDate: TDateTime index 8 read GetTDateTimeProp;
    property NumberOfWithdraws: Integer index 9 read GetIntegerProp;
    property WithdrawsDuration: Integer index 10 read GetIntegerProp;
    property NumberOfACDCalls: Integer index 11 read GetIntegerProp;
    property NumberOfPrivateCalls: Integer index 12 read GetIntegerProp;
    property PrivateCallsDuration: Integer index 13 read GetIntegerProp;
    property NumberOfInterceptedACDCalls: Integer index 14 read GetIntegerProp;
    property NumberOfTransferredACDCalls: Integer index 15 read GetIntegerProp;
    property NumberOfRefusedACDCalls: Integer index 16 read GetIntegerProp;
    property ServedPilotDirNumber: WideString index 17 read GetWideStringProp;
    property NbEvtQueued: Integer index 18 read GetIntegerProp;
    property LastError: Integer index 19 read GetIntegerProp;
    property AgentNameText: WideString index 21 write SetWideStringProp;
    property GeneralStatisticsText: WideString index 22 write SetWideStringProp;
    property LogonText: WideString index 23 write SetWideStringProp;
    property LogonTimeText: WideString index 24 write SetWideStringProp;
    property LogonDurationText: WideString index 25 write SetWideStringProp;
    property WithdrawalsText: WideString index 26 write SetWideStringProp;
    property WithdrawalsCountText: WideString index 27 write SetWideStringProp;
    property WithdrawalsDurationText: WideString index 28 write SetWideStringProp;
    property CallsStatisticsText: WideString index 29 write SetWideStringProp;
    property PrivatesText: WideString index 30 write SetWideStringProp;
    property PrivatesCountText: WideString index 31 write SetWideStringProp;
    property PrivatesDurationText: WideString index 32 write SetWideStringProp;
    property AcdText: WideString index 33 write SetWideStringProp;
    property AcdProcessedText: WideString index 34 write SetWideStringProp;
    property AcdRefusedText: WideString index 35 write SetWideStringProp;
    property AcdPickedUpText: WideString index 36 write SetWideStringProp;
    property AcdTransferredText: WideString index 37 write SetWideStringProp;
  published
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  Visible;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property OnPropertyChanges: TAxAgentStatisticsOnPropertyChanges read FOnPropertyChanges write FOnPropertyChanges;
  end;


// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TAxTeamStatistics
// Chaîne d'aide        : AxTeamStatistics Class
// Interface par défaut : IAxTeamStatistics
// DISP Int. Déf. ?     : No
// Interface événements : _IAxTeamStatisticsEvents
// TypeFlags            : (2) CanCreate
// *********************************************************************//
  TAxTeamStatisticsOnPropertyChanges = procedure(Sender: TObject; Status: AxPCAgentErrorTypes; 
                                                                  changes: AxACDTeamStateChanges) of object;

  TAxTeamStatistics = class(TOleControl)
  private
    FOnPropertyChanges: TAxTeamStatisticsOnPropertyChanges;
    FIntf: IAxTeamStatistics;
    function  GetControlInterface: IAxTeamStatistics;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    function  Connect: Integer;
    property  ControlInterface: IAxTeamStatistics read GetControlInterface;
    property  DefaultInterface: IAxTeamStatistics read GetControlInterface;
    property NumberOfLoggedAgents: Integer index 1 read GetIntegerProp;
    property NumberOfAffectedAgents: Integer index 2 read GetIntegerProp;
    property NumberOfWithdrawedAgents: Integer index 3 read GetIntegerProp;
    property NumberOfBusyAgents: Integer index 4 read GetIntegerProp;
    property NumberOfFreeAgents: Integer index 5 read GetIntegerProp;
    property NumberOfACDCalls: Integer index 6 read GetIntegerProp;
    property NumberOfPrivateCalls: Integer index 7 read GetIntegerProp;
    property NumberOfWaitingCalls: Integer index 8 read GetIntegerProp;
    property MaximumWaitingTime: Integer index 9 read GetIntegerProp;
    property NbEvtQueued: Integer index 10 read GetIntegerProp;
    property LastError: Integer index 11 read GetIntegerProp;
    property AverageWaitingTime: Single index 12 read GetSingleProp;
    property GeneralStatisticsText: WideString index 14 write SetWideStringProp;
    property ProcessingGroupText: WideString index 15 write SetWideStringProp;
    property WaitText: WideString index 16 write SetWideStringProp;
    property WaitMaxText: WideString index 17 write SetWideStringProp;
    property WaitAverageText: WideString index 18 write SetWideStringProp;
    property CallsText: WideString index 19 write SetWideStringProp;
    property CallsAcdText: WideString index 20 write SetWideStringProp;
    property CallsPrivateText: WideString index 21 write SetWideStringProp;
    property CallsWaitText: WideString index 22 write SetWideStringProp;
    property AgentStatisticsText: WideString index 23 write SetWideStringProp;
    property LoggedText: WideString index 24 write SetWideStringProp;
    property SetText: WideString index 25 write SetWideStringProp;
    property WithdrawsText: WideString index 26 write SetWideStringProp;
    property BusyText: WideString index 27 write SetWideStringProp;
    property IdleText: WideString index 28 write SetWideStringProp;
  published
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  Visible;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property MonitoredGroupNumber: WideString index 29 read GetWideStringProp write SetWideStringProp stored False;
    property OnPropertyChanges: TAxTeamStatisticsOnPropertyChanges read FOnPropertyChanges write FOnPropertyChanges;
  end;


// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TAxHold
// Chaîne d'aide        : AxHold Class
// Interface par défaut : IAxHold
// DISP Int. Déf. ?     : No
// Interface événements : _IAxHoldEvents
// TypeFlags            : (2) CanCreate
// *********************************************************************//
  TAxHoldOnHoldCallPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxHoldOnRequestStatus = procedure(Sender: TObject; requestId: Integer; 
                                                      Status: AxTServerErrorTypes) of object;

  TAxHold = class(TOleControl)
  private
    FOnHoldCallPossible: TAxHoldOnHoldCallPossible;
    FOnRequestStatus: TAxHoldOnRequestStatus;
    FIntf: IAxHold;
    function  GetControlInterface: IAxHold;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    function  HoldCall: Integer;
    function  Connect: Integer;
    property  ControlInterface: IAxHold read GetControlInterface;
    property  DefaultInterface: IAxHold read GetControlInterface;
    property HoldCallPossible: Integer index 3 read GetIntegerProp;
  published
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  Visible;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property Picture: TPicture index -523 read GetTPictureProp write SetTPictureProp stored False;
    property OnHoldCallPossible: TAxHoldOnHoldCallPossible read FOnHoldCallPossible write FOnHoldCallPossible;
    property OnRequestStatus: TAxHoldOnRequestStatus read FOnRequestStatus write FOnRequestStatus;
  end;


// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TAxDoNotCallRecord
// Chaîne d'aide        : AxDoNotCallRecord Class
// Interface par défaut : IAxDoNotCallRecord
// DISP Int. Déf. ?     : No
// Interface événements : _IAxDoNotCallRecordEvents
// TypeFlags            : (2) CanCreate
// *********************************************************************//
  TAxDoNotCallRecordOnDoNotCallPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxDoNotCallRecordOnRequestStatus = procedure(Sender: TObject; requestId: Integer; 
                                                                 Status: AxTServerErrorTypes) of object;

  TAxDoNotCallRecord = class(TOleControl)
  private
    FOnDoNotCallPossible: TAxDoNotCallRecordOnDoNotCallPossible;
    FOnRequestStatus: TAxDoNotCallRecordOnRequestStatus;
    FIntf: IAxDoNotCallRecord;
    function  GetControlInterface: IAxDoNotCallRecord;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    function  DoNotCall(handle: Integer): Integer;
    function  Connect: Integer;
    property  ControlInterface: IAxDoNotCallRecord read GetControlInterface;
    property  DefaultInterface: IAxDoNotCallRecord read GetControlInterface;
    property DoNotCallRecord: Integer index 3 read GetIntegerProp;
  published
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  Visible;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property Picture: TPicture index -523 read GetTPictureProp write SetTPictureProp stored False;
    property OnDoNotCallPossible: TAxDoNotCallRecordOnDoNotCallPossible read FOnDoNotCallPossible write FOnDoNotCallPossible;
    property OnRequestStatus: TAxDoNotCallRecordOnRequestStatus read FOnRequestStatus write FOnRequestStatus;
  end;


// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TAxRequestRejectPreviewRecord
// Chaîne d'aide        : AxRequestRejectPreviewRecord Class
// Interface par défaut : IAxRequestRejectPreviewRecord
// DISP Int. Déf. ?     : No
// Interface événements : _IAxRequestRejectPreviewRecordEvents
// TypeFlags            : (2) CanCreate
// *********************************************************************//
  TAxRequestRejectPreviewRecordOnPreviewRecordPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxRequestRejectPreviewRecordOnRecordRejectPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxRequestRejectPreviewRecordOnRequestStatus = procedure(Sender: TObject; requestId: Integer; 
                                                                            Status: AxTServerErrorTypes) of object;

  TAxRequestRejectPreviewRecord = class(TOleControl)
  private
    FOnPreviewRecordPossible: TAxRequestRejectPreviewRecordOnPreviewRecordPossible;
    FOnRecordRejectPossible: TAxRequestRejectPreviewRecordOnRecordRejectPossible;
    FOnRequestStatus: TAxRequestRejectPreviewRecordOnRequestStatus;
    FIntf: IAxRequestRejectPreviewRecord;
    function  GetControlInterface: IAxRequestRejectPreviewRecord;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    function  PreviewRecord: Integer;
    function  RecordReject(handle: Integer): Integer;
    function  Connect: Integer;
    property  ControlInterface: IAxRequestRejectPreviewRecord read GetControlInterface;
    property  DefaultInterface: IAxRequestRejectPreviewRecord read GetControlInterface;
    property RequestPreviewRecordPossible: Integer index 4 read GetIntegerProp;
    property RejectPreviewRecordPossible: Integer index 5 read GetIntegerProp;
  published
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  Visible;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property Picture: TPicture index -523 read GetTPictureProp write SetTPictureProp stored False;
    property OnPreviewRecordPossible: TAxRequestRejectPreviewRecordOnPreviewRecordPossible read FOnPreviewRecordPossible write FOnPreviewRecordPossible;
    property OnRecordRejectPossible: TAxRequestRejectPreviewRecordOnRecordRejectPossible read FOnRecordRejectPossible write FOnRecordRejectPossible;
    property OnRequestStatus: TAxRequestRejectPreviewRecordOnRequestStatus read FOnRequestStatus write FOnRequestStatus;
  end;


// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TAxScheduleCallback
// Chaîne d'aide        : AxScheduleCallback Class
// Interface par défaut : IAxScheduleCallback
// DISP Int. Déf. ?     : No
// Interface événements : _IAxScheduleCallbackEvents
// TypeFlags            : (2) CanCreate
// *********************************************************************//
  TAxScheduleCallbackOnRecordReschedulePossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxScheduleCallbackOnRequestStatus = procedure(Sender: TObject; requestId: Integer; 
                                                                  Status: AxTServerErrorTypes) of object;

  TAxScheduleCallback = class(TOleControl)
  private
    FOnRecordReschedulePossible: TAxScheduleCallbackOnRecordReschedulePossible;
    FOnRequestStatus: TAxScheduleCallbackOnRequestStatus;
    FIntf: IAxScheduleCallback;
    function  GetControlInterface: IAxScheduleCallback;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    function  Connect: Integer;
    function  RescheduleRecord(handle: Integer; const callbackType: WideString; 
                               callbackDate: OleVariant): Integer;
    property  ControlInterface: IAxScheduleCallback read GetControlInterface;
    property  DefaultInterface: IAxScheduleCallback read GetControlInterface;
    property CallbackTypesList: WideString index 2 write SetWideStringProp;
    property CallbackDateFormat: WideString index 3 write SetWideStringProp;
    property ScheduleCallbackPossible: Integer index 6 read GetIntegerProp;
    property PopupWindowTitle: WideString index 7 write SetWideStringProp;
    property CallbackTypeText: WideString index 8 write SetWideStringProp;
    property CallbackDateText: WideString index 9 write SetWideStringProp;
    property OkButtonText: WideString index 10 write SetWideStringProp;
    property CancelButtonText: WideString index 11 write SetWideStringProp;
  published
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  Visible;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property Picture: TPicture index -523 read GetTPictureProp write SetTPictureProp stored False;
    property OnRecordReschedulePossible: TAxScheduleCallbackOnRecordReschedulePossible read FOnRecordReschedulePossible write FOnRecordReschedulePossible;
    property OnRequestStatus: TAxScheduleCallbackOnRequestStatus read FOnRequestStatus write FOnRequestStatus;
  end;


// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TAxClassify
// Chaîne d'aide        : AxClassify Class
// Interface par défaut : IAxClassify
// DISP Int. Déf. ?     : No
// Interface événements : _IAxClassifyEvents
// TypeFlags            : (2) CanCreate
// *********************************************************************//
  TAxClassifyOnRecordProcessedPossible = procedure(Sender: TObject; Status: Integer) of object;
  TAxClassifyOnRequestStatus = procedure(Sender: TObject; requestId: Integer; 
                                                          Status: AxTServerErrorTypes) of object;

  TAxClassify = class(TOleControl)
  private
    FOnRecordProcessedPossible: TAxClassifyOnRecordProcessedPossible;
    FOnRequestStatus: TAxClassifyOnRequestStatus;
    FIntf: IAxClassify;
    function  GetControlInterface: IAxClassify;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    function  RecordProcessed(handle: Integer): Integer;
    function  Connect: Integer;
    property  ControlInterface: IAxClassify read GetControlInterface;
    property  DefaultInterface: IAxClassify read GetControlInterface;
    property RecordToClassify: Integer index 3 read GetIntegerProp;
  published
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  Visible;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property Picture: TPicture index -523 read GetTPictureProp write SetTPictureProp stored False;
    property OnRecordProcessedPossible: TAxClassifyOnRecordProcessedPossible read FOnRecordProcessedPossible write FOnRecordProcessedPossible;
    property OnRequestStatus: TAxClassifyOnRequestStatus read FOnRequestStatus write FOnRequestStatus;
  end;

procedure Register;

implementation

uses ComObj;

procedure TAxApplication.InitControlData;
const
  CEventDispIDs: array [0..2] of DWORD = (
    $00000001, $00000002, $00000003);
  CControlData: TControlData2 = (
    ClassID: '{9E78C026-DD92-11D5-B772-00508B79E1ED}';
    EventIID: '{C45074C2-DDA3-11D5-B773-00508B79E1ED}';
    EventCount: 3;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnQuit) - Cardinal(Self);
end;

procedure TAxApplication.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IAxApplication;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TAxApplication.GetControlInterface: IAxApplication;
begin
  CreateControl;
  Result := FIntf;
end;

function  TAxApplication.Connect: Integer;
begin
  Result := DefaultInterface.Connect;
end;

class function CoDataOfCall.Create: IDataOfCall;
begin
  Result := CreateComObject(CLASS_DataOfCall) as IDataOfCall;
end;

class function CoDataOfCall.CreateRemote(const MachineName: string): IDataOfCall;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_DataOfCall) as IDataOfCall;
end;

procedure TDataOfCall.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{D71863E6-5B66-11D6-8C70-0008C7BF3D97}';
    IntfIID:   '{D71863E5-5B66-11D6-8C70-0008C7BF3D97}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TDataOfCall.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as IDataOfCall;
  end;
end;

procedure TDataOfCall.ConnectTo(svrIntf: IDataOfCall);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TDataOfCall.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TDataOfCall.GetDefaultInterface: IDataOfCall;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface est NULL. Le composant n''est pas connecté au serveur. Vous devez appeler ''Connect'' ou ''ConnectTo'' avant cette opération');
  Result := FIntf;
end;

constructor TDataOfCall.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TDataOfCallProperties.Create(Self);
{$ENDIF}
end;

destructor TDataOfCall.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TDataOfCall.GetServerProperties: TDataOfCallProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function  TDataOfCall.Get_Id: Integer;
begin
  Result := DefaultInterface.Get_Id;
end;

function  TDataOfCall.Get_LastName: WideString;
begin
  Result := DefaultInterface.Get_LastName;
end;

function  TDataOfCall.Get_FirstName: WideString;
begin
  Result := DefaultInterface.Get_FirstName;
end;

function  TDataOfCall.Get_Number: WideString;
begin
  Result := DefaultInterface.Get_Number;
end;

function  TDataOfCall.Get_Incoming: Integer;
begin
  Result := DefaultInterface.Get_Incoming;
end;

function  TDataOfCall.Get_Extern: Integer;
begin
  Result := DefaultInterface.Get_Extern;
end;

function  TDataOfCall.Get_DestinationNumber: WideString;
begin
  Result := DefaultInterface.Get_DestinationNumber;
end;

function  TDataOfCall.Get_ConferenceMaster: WideString;
begin
  Result := DefaultInterface.Get_ConferenceMaster;
end;

function  TDataOfCall.Get_State: AxAgentCallStates;
begin
  Result := DefaultInterface.Get_State;
end;

function  TDataOfCall.Get_Cause: AxAgentCallCauses;
begin
  Result := DefaultInterface.Get_Cause;
end;

function  TDataOfCall.Get_ACDCall: Integer;
begin
  Result := DefaultInterface.Get_ACDCall;
end;

function  TDataOfCall.Get_TrunkUsed: WideString;
begin
  Result := DefaultInterface.Get_TrunkUsed;
end;

function  TDataOfCall.Get_EstimatedWaitingTime: Integer;
begin
  Result := DefaultInterface.Get_EstimatedWaitingTime;
end;

function  TDataOfCall.Get_PilotStatus: AxPilotStatus;
begin
  Result := DefaultInterface.Get_PilotStatus;
end;

function  TDataOfCall.Get_PossibleTransfer: Integer;
begin
  Result := DefaultInterface.Get_PossibleTransfer;
end;

function  TDataOfCall.Get_WaitingQueueSaturation: Integer;
begin
  Result := DefaultInterface.Get_WaitingQueueSaturation;
end;

function  TDataOfCall.Get_LastRedirectionDevice: WideString;
begin
  Result := DefaultInterface.Get_LastRedirectionDevice;
end;

function  TDataOfCall.Get_CorrelatorData: WideString;
begin
  Result := DefaultInterface.Get_CorrelatorData;
end;

procedure TDataOfCall.Set_CorrelatorData(const pVal: WideString);
begin
  DefaultInterface.Set_CorrelatorData(pVal);
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TDataOfCallProperties.Create(AServer: TDataOfCall);
begin
  inherited Create;
  FServer := AServer;
end;

function TDataOfCallProperties.GetDefaultInterface: IDataOfCall;
begin
  Result := FServer.DefaultInterface;
end;

function  TDataOfCallProperties.Get_Id: Integer;
begin
  Result := DefaultInterface.Get_Id;
end;

function  TDataOfCallProperties.Get_LastName: WideString;
begin
  Result := DefaultInterface.Get_LastName;
end;

function  TDataOfCallProperties.Get_FirstName: WideString;
begin
  Result := DefaultInterface.Get_FirstName;
end;

function  TDataOfCallProperties.Get_Number: WideString;
begin
  Result := DefaultInterface.Get_Number;
end;

function  TDataOfCallProperties.Get_Incoming: Integer;
begin
  Result := DefaultInterface.Get_Incoming;
end;

function  TDataOfCallProperties.Get_Extern: Integer;
begin
  Result := DefaultInterface.Get_Extern;
end;

function  TDataOfCallProperties.Get_DestinationNumber: WideString;
begin
  Result := DefaultInterface.Get_DestinationNumber;
end;

function  TDataOfCallProperties.Get_ConferenceMaster: WideString;
begin
  Result := DefaultInterface.Get_ConferenceMaster;
end;

function  TDataOfCallProperties.Get_State: AxAgentCallStates;
begin
  Result := DefaultInterface.Get_State;
end;

function  TDataOfCallProperties.Get_Cause: AxAgentCallCauses;
begin
  Result := DefaultInterface.Get_Cause;
end;

function  TDataOfCallProperties.Get_ACDCall: Integer;
begin
  Result := DefaultInterface.Get_ACDCall;
end;

function  TDataOfCallProperties.Get_TrunkUsed: WideString;
begin
  Result := DefaultInterface.Get_TrunkUsed;
end;

function  TDataOfCallProperties.Get_EstimatedWaitingTime: Integer;
begin
  Result := DefaultInterface.Get_EstimatedWaitingTime;
end;

function  TDataOfCallProperties.Get_PilotStatus: AxPilotStatus;
begin
  Result := DefaultInterface.Get_PilotStatus;
end;

function  TDataOfCallProperties.Get_PossibleTransfer: Integer;
begin
  Result := DefaultInterface.Get_PossibleTransfer;
end;

function  TDataOfCallProperties.Get_WaitingQueueSaturation: Integer;
begin
  Result := DefaultInterface.Get_WaitingQueueSaturation;
end;

function  TDataOfCallProperties.Get_LastRedirectionDevice: WideString;
begin
  Result := DefaultInterface.Get_LastRedirectionDevice;
end;

function  TDataOfCallProperties.Get_CorrelatorData: WideString;
begin
  Result := DefaultInterface.Get_CorrelatorData;
end;

procedure TDataOfCallProperties.Set_CorrelatorData(const pVal: WideString);
begin
  DefaultInterface.Set_CorrelatorData(pVal);
end;

{$ENDIF}

procedure TAxLine.InitControlData;
const
  CEventDispIDs: array [0..38] of DWORD = (
    $00000001, $00000002, $00000003, $00000004, $00000005, $00000006,
    $00000007, $00000008, $00000009, $0000000A, $0000000B, $0000000C,
    $0000000D, $0000000E, $0000000F, $00000010, $00000011, $00000012,
    $00000013, $00000014, $00000015, $00000016, $00000017, $00000018,
    $00000019, $0000001A, $0000001B, $0000001C, $0000001D, $0000001E,
    $0000001F, $00000020, $00000021, $00000022, $00000023, $00000024,
    $00000025, $00000026, $00000027);
  CControlData: TControlData2 = (
    ClassID: '{9E78C029-DD92-11D5-B772-00508B79E1ED}';
    EventIID: '{C45074C4-DDA3-11D5-B773-00508B79E1ED}';
    EventCount: 39;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnLineNewCall) - Cardinal(Self);
end;

procedure TAxLine.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IAxLine;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TAxLine.GetControlInterface: IAxLine;
begin
  CreateControl;
  Result := FIntf;
end;

function  TAxLine.Get_AffectedTeamDirNumber(nTeam: SYSINT): WideString;
begin
  Result := DefaultInterface.Get_AffectedTeamDirNumber(nTeam);
end;

function  TAxLine.Get_ContactDataKey(handle: Integer; field: SYSINT): WideString;
begin
  Result := DefaultInterface.Get_ContactDataKey(handle, field);
end;

function  TAxLine.Get_ContactDataValue(handle: Integer; const key: WideString): OleVariant;
begin
  Result := DefaultInterface.Get_ContactDataValue(handle, key);
end;

procedure TAxLine.Set_ContactDataValue(handle: Integer; const key: WideString; pVal: OleVariant);
begin
  DefaultInterface.Set_ContactDataValue(handle, key, pVal);
end;

function  TAxLine.Get_ContactDataVisible(handle: Integer): Integer;
begin
  Result := DefaultInterface.Get_ContactDataVisible(handle);
end;

procedure TAxLine.Set_ContactDataVisible(handle: Integer; pVal: Integer);
begin
  DefaultInterface.Set_ContactDataVisible(handle, pVal);
end;

function  TAxLine.Get_IsChainedRecordPossible(handle: Integer): Integer;
begin
  Result := DefaultInterface.Get_IsChainedRecordPossible(handle);
end;

function  TAxLine.Get_IsDoNotCallPossible(handle: Integer): Integer;
begin
  Result := DefaultInterface.Get_IsDoNotCallPossible(handle);
end;

function  TAxLine.Get_IsRecordCancelPossible(handle: Integer): Integer;
begin
  Result := DefaultInterface.Get_IsRecordCancelPossible(handle);
end;

function  TAxLine.Get_IsRecordProcessedPossible(handle: Integer): Integer;
begin
  Result := DefaultInterface.Get_IsRecordProcessedPossible(handle);
end;

function  TAxLine.Get_IsRecordRejectPossible(handle: Integer): Integer;
begin
  Result := DefaultInterface.Get_IsRecordRejectPossible(handle);
end;

function  TAxLine.Get_IsRecordReschedulePossible(handle: Integer): Integer;
begin
  Result := DefaultInterface.Get_IsRecordReschedulePossible(handle);
end;

function  TAxLine.Get_IsReleaseCallPossible(callId: Integer): Integer;
begin
  Result := DefaultInterface.Get_IsReleaseCallPossible(callId);
end;

function  TAxLine.Get_IsTakeCallPossible(callId: Integer): Integer;
begin
  Result := DefaultInterface.Get_IsTakeCallPossible(callId);
end;

function  TAxLine.Get_IsUpdateCallCompStatsPossible(handle: Integer): Integer;
begin
  Result := DefaultInterface.Get_IsUpdateCallCompStatsPossible(handle);
end;

function  TAxLine.Get_NumberOfRecordFields(handle: Integer): Integer;
begin
  Result := DefaultInterface.Get_NumberOfRecordFields(handle);
end;

function  TAxLine.Get_WithdrawCause(nCause: Integer): WideString;
begin
  Result := DefaultInterface.Get_WithdrawCause(nCause);
end;

function  TAxLine.Get_CallById(callId: Integer): IDataOfCall;
begin
  Result := DefaultInterface.Get_CallById(callId);
end;

function  TAxLine.Get_CallByIndex(CallIndex: Integer): IDataOfCall;
begin
  Result := DefaultInterface.Get_CallByIndex(CallIndex);
end;

function  TAxLine.ChainedRecord(handle: Integer): Integer;
begin
  Result := DefaultInterface.ChainedRecord(handle);
end;

function  TAxLine.Conference: Integer;
begin
  Result := DefaultInterface.Conference;
end;

function  TAxLine.DoNotCall(handle: Integer): Integer;
begin
  Result := DefaultInterface.DoNotCall(handle);
end;

function  TAxLine.HoldCall: Integer;
begin
  Result := DefaultInterface.HoldCall;
end;

function  TAxLine.Logoff(const Agent: WideString; const Team: WideString; const Password: WideString): Integer;
begin
  Result := DefaultInterface.Logoff(Agent, Team, Password);
end;

function  TAxLine.Logon(const Agent: WideString; const Team: WideString; 
                        const Password: WideString; outbound: Integer; multimediaTools: Integer): Integer;
begin
  Result := DefaultInterface.Logon(Agent, Team, Password, outbound, multimediaTools);
end;

function  TAxLine.MakeCall(const destNumber: WideString; const correlator: WideString): Integer;
begin
  Result := DefaultInterface.MakeCall(destNumber, correlator);
end;

function  TAxLine.Pause: Integer;
begin
  Result := DefaultInterface.Pause;
end;

function  TAxLine.PauseEnd: Integer;
begin
  Result := DefaultInterface.PauseEnd;
end;

function  TAxLine.PreviewRecord: Integer;
begin
  Result := DefaultInterface.PreviewRecord;
end;

function  TAxLine.RecordCancel(handle: Integer): Integer;
begin
  Result := DefaultInterface.RecordCancel(handle);
end;

function  TAxLine.RecordProcessed(handle: Integer): Integer;
begin
  Result := DefaultInterface.RecordProcessed(handle);
end;

function  TAxLine.RecordReject(handle: Integer): Integer;
begin
  Result := DefaultInterface.RecordReject(handle);
end;

function  TAxLine.RecordReschedule(handle: Integer): Integer;
begin
  Result := DefaultInterface.RecordReschedule(handle);
end;

function  TAxLine.ReleaseCall(callId: Integer): Integer;
begin
  Result := DefaultInterface.ReleaseCall(callId);
end;

function  TAxLine.SetDateAndTime(const szDate: WideString): Integer;
begin
  Result := DefaultInterface.SetDateAndTime(szDate);
end;

function  TAxLine.TakeCall(callId: Integer): Integer;
begin
  Result := DefaultInterface.TakeCall(callId);
end;

function  TAxLine.TransferCall(heldCallId: Integer): Integer;
begin
  Result := DefaultInterface.TransferCall(heldCallId);
end;

function  TAxLine.UpdateCallCompletionStats(handle: Integer): Integer;
begin
  Result := DefaultInterface.UpdateCallCompletionStats(handle);
end;

function  TAxLine.Withdraw(Cause: Integer): Integer;
begin
  Result := DefaultInterface.Withdraw(Cause);
end;

function  TAxLine.WithdrawEnd: Integer;
begin
  Result := DefaultInterface.WithdrawEnd;
end;

function  TAxLine.WrapUp: Integer;
begin
  Result := DefaultInterface.WrapUp;
end;

function  TAxLine.WrapUpEnd: Integer;
begin
  Result := DefaultInterface.WrapUpEnd;
end;

procedure TAxLine.ShowEvents(Enabled: Integer);
begin
  DefaultInterface.ShowEvents(Enabled);
end;

function  TAxLine.Connect: Integer;
begin
  Result := DefaultInterface.Connect;
end;

function  TAxLine.EndConference: Integer;
begin
  Result := DefaultInterface.EndConference;
end;

function  TAxLine.IsEndConferencePossible: Integer;
begin
  Result := DefaultInterface.IsEndConferencePossible;
end;

function  TAxLine.ContactAdd(const pContactNumber: WideString): Integer;
begin
  Result := DefaultInterface.ContactAdd(pContactNumber);
end;

function  TAxLine.ContactSave(nHandle: Integer): Integer;
begin
  Result := DefaultInterface.ContactSave(nHandle);
end;

function  TAxLine.ContactCancel(nHandle: Integer): Integer;
begin
  Result := DefaultInterface.ContactCancel(nHandle);
end;

function  TAxLine.IsAddContactPossible: Integer;
begin
  Result := DefaultInterface.IsAddContactPossible;
end;

procedure TAxMakeCall.InitControlData;
const
  CEventDispIDs: array [0..1] of DWORD = (
    $00000001, $00000002);
  CControlData: TControlData2 = (
    ClassID: '{AFA40FF4-E319-11D5-B776-00508B79E1ED}';
    EventIID: '{AFA40FF5-E319-11D5-B776-00508B79E1ED}';
    EventCount: 2;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnMakeCallPossible) - Cardinal(Self);
end;

procedure TAxMakeCall.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IAxMakeCall;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TAxMakeCall.GetControlInterface: IAxMakeCall;
begin
  CreateControl;
  Result := FIntf;
end;

function  TAxMakeCall.MakeCall(const destNumber: WideString; const correlator: WideString): Integer;
begin
  Result := DefaultInterface.MakeCall(destNumber, correlator);
end;

function  TAxMakeCall.Connect: Integer;
begin
  Result := DefaultInterface.Connect;
end;

procedure TAxHangUp.InitControlData;
const
  CEventDispIDs: array [0..1] of DWORD = (
    $00000001, $00000002);
  CTPictureIDs: array [0..0] of DWORD = (
    $FFFFFDF5);
  CControlData: TControlData2 = (
    ClassID: '{AFA40FFB-E319-11D5-B776-00508B79E1ED}';
    EventIID: '{AFA40FFC-E319-11D5-B776-00508B79E1ED}';
    EventCount: 2;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401;
    FontCount: 0;
    FontIDs: nil;
    PictureCount: 1;
    PictureIDs: @CTPictureIDs);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnReleaseCallPossible) - Cardinal(Self);
end;

procedure TAxHangUp.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IAxHangUp;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TAxHangUp.GetControlInterface: IAxHangUp;
begin
  CreateControl;
  Result := FIntf;
end;

function  TAxHangUp.Get_ReleaseCallPossible(callId: Integer): Integer;
begin
  Result := DefaultInterface.Get_ReleaseCallPossible(callId);
end;

function  TAxHangUp.ReleaseCall(callId: Integer): Integer;
begin
  Result := DefaultInterface.ReleaseCall(callId);
end;

function  TAxHangUp.Connect: Integer;
begin
  Result := DefaultInterface.Connect;
end;

procedure TAxTakeCall.InitControlData;
const
  CEventDispIDs: array [0..2] of DWORD = (
    $00000001, $00000002, $00000003);
  CTPictureIDs: array [0..0] of DWORD = (
    $FFFFFDF5);
  CControlData: TControlData2 = (
    ClassID: '{BF1B0DB8-E7FC-11D5-B778-00508B79E1ED}';
    EventIID: '{AA674F70-E801-11D5-B778-00508B79E1ED}';
    EventCount: 3;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401;
    FontCount: 0;
    FontIDs: nil;
    PictureCount: 1;
    PictureIDs: @CTPictureIDs);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnTakeCallPossible) - Cardinal(Self);
end;

procedure TAxTakeCall.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IAxTakeCall;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TAxTakeCall.GetControlInterface: IAxTakeCall;
begin
  CreateControl;
  Result := FIntf;
end;

function  TAxTakeCall.Get_TakeCallPossible(callId: Integer): Integer;
begin
  Result := DefaultInterface.Get_TakeCallPossible(callId);
end;

function  TAxTakeCall.TakeCall(callId: Integer): Integer;
begin
  Result := DefaultInterface.TakeCall(callId);
end;

function  TAxTakeCall.Connect: Integer;
begin
  Result := DefaultInterface.Connect;
end;

function  TAxTakeCall.AlternateCall: Integer;
begin
  Result := DefaultInterface.AlternateCall;
end;

procedure TAxConference.InitControlData;
const
  CEventDispIDs: array [0..2] of DWORD = (
    $00000001, $00000002, $00000003);
  CTPictureIDs: array [0..0] of DWORD = (
    $FFFFFDF5);
  CControlData: TControlData2 = (
    ClassID: '{37D5EB37-E88F-11D5-B779-00508B79E1ED}';
    EventIID: '{37D5EB38-E88F-11D5-B779-00508B79E1ED}';
    EventCount: 3;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401;
    FontCount: 0;
    FontIDs: nil;
    PictureCount: 1;
    PictureIDs: @CTPictureIDs);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnConferenceCallPossible) - Cardinal(Self);
end;

procedure TAxConference.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IAxConference;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TAxConference.GetControlInterface: IAxConference;
begin
  CreateControl;
  Result := FIntf;
end;

function  TAxConference.Conference: Integer;
begin
  Result := DefaultInterface.Conference;
end;

function  TAxConference.Connect: Integer;
begin
  Result := DefaultInterface.Connect;
end;

function  TAxConference.ConferenceEnd: Integer;
begin
  Result := DefaultInterface.ConferenceEnd;
end;

procedure TAxLog.InitControlData;
const
  CEventDispIDs: array [0..2] of DWORD = (
    $00000001, $00000002, $00000003);
  CTPictureIDs: array [0..0] of DWORD = (
    $FFFFFDF5);
  CControlData: TControlData2 = (
    ClassID: '{7F272F85-E98E-11D5-B779-00508B79E1ED}';
    EventIID: '{7F272F86-E98E-11D5-B779-00508B79E1ED}';
    EventCount: 3;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401;
    FontCount: 0;
    FontIDs: nil;
    PictureCount: 1;
    PictureIDs: @CTPictureIDs);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnLogonPossible) - Cardinal(Self);
end;

procedure TAxLog.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IAxLog;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TAxLog.GetControlInterface: IAxLog;
begin
  CreateControl;
  Result := FIntf;
end;

function  TAxLog.Logon(const Agent: WideString; const Team: WideString; const Password: WideString; 
                       outbound: Integer; multimediaTools: Integer): Integer;
begin
  Result := DefaultInterface.Logon(Agent, Team, Password, outbound, multimediaTools);
end;

function  TAxLog.Logoff(const Agent: WideString; const Team: WideString; const Password: WideString): Integer;
begin
  Result := DefaultInterface.Logoff(Agent, Team, Password);
end;

function  TAxLog.Connect: Integer;
begin
  Result := DefaultInterface.Connect;
end;

class function CoItemOfHistoric.Create: IItemOfHistoric;
begin
  Result := CreateComObject(CLASS_ItemOfHistoric) as IItemOfHistoric;
end;

class function CoItemOfHistoric.CreateRemote(const MachineName: string): IItemOfHistoric;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ItemOfHistoric) as IItemOfHistoric;
end;

procedure TItemOfHistoric.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{7CB3D12A-5C21-11D6-8C70-0008C7BF3D97}';
    IntfIID:   '{7CB3D129-5C21-11D6-8C70-0008C7BF3D97}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TItemOfHistoric.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as IItemOfHistoric;
  end;
end;

procedure TItemOfHistoric.ConnectTo(svrIntf: IItemOfHistoric);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TItemOfHistoric.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TItemOfHistoric.GetDefaultInterface: IItemOfHistoric;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface est NULL. Le composant n''est pas connecté au serveur. Vous devez appeler ''Connect'' ou ''ConnectTo'' avant cette opération');
  Result := FIntf;
end;

constructor TItemOfHistoric.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TItemOfHistoricProperties.Create(Self);
{$ENDIF}
end;

destructor TItemOfHistoric.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TItemOfHistoric.GetServerProperties: TItemOfHistoricProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function  TItemOfHistoric.Get_IdHigh: Integer;
begin
  Result := DefaultInterface.Get_IdHigh;
end;

function  TItemOfHistoric.Get_IdLow: Integer;
begin
  Result := DefaultInterface.Get_IdLow;
end;

function  TItemOfHistoric.Get_Number: WideString;
begin
  Result := DefaultInterface.Get_Number;
end;

function  TItemOfHistoric.Get_LastName: WideString;
begin
  Result := DefaultInterface.Get_LastName;
end;

function  TItemOfHistoric.Get_FirstName: WideString;
begin
  Result := DefaultInterface.Get_FirstName;
end;

function  TItemOfHistoric.Get_State: AxAgentCallStates;
begin
  Result := DefaultInterface.Get_State;
end;

function  TItemOfHistoric.Get_Incoming: Integer;
begin
  Result := DefaultInterface.Get_Incoming;
end;

function  TItemOfHistoric.Get_Extern: Integer;
begin
  Result := DefaultInterface.Get_Extern;
end;

function  TItemOfHistoric.Get_BeginDate: TDateTime;
begin
  Result := DefaultInterface.Get_BeginDate;
end;

function  TItemOfHistoric.Get_EndDate: TDateTime;
begin
  Result := DefaultInterface.Get_EndDate;
end;

function  TItemOfHistoric.Get_TalkingDate: TDateTime;
begin
  Result := DefaultInterface.Get_TalkingDate;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TItemOfHistoricProperties.Create(AServer: TItemOfHistoric);
begin
  inherited Create;
  FServer := AServer;
end;

function TItemOfHistoricProperties.GetDefaultInterface: IItemOfHistoric;
begin
  Result := FServer.DefaultInterface;
end;

function  TItemOfHistoricProperties.Get_IdHigh: Integer;
begin
  Result := DefaultInterface.Get_IdHigh;
end;

function  TItemOfHistoricProperties.Get_IdLow: Integer;
begin
  Result := DefaultInterface.Get_IdLow;
end;

function  TItemOfHistoricProperties.Get_Number: WideString;
begin
  Result := DefaultInterface.Get_Number;
end;

function  TItemOfHistoricProperties.Get_LastName: WideString;
begin
  Result := DefaultInterface.Get_LastName;
end;

function  TItemOfHistoricProperties.Get_FirstName: WideString;
begin
  Result := DefaultInterface.Get_FirstName;
end;

function  TItemOfHistoricProperties.Get_State: AxAgentCallStates;
begin
  Result := DefaultInterface.Get_State;
end;

function  TItemOfHistoricProperties.Get_Incoming: Integer;
begin
  Result := DefaultInterface.Get_Incoming;
end;

function  TItemOfHistoricProperties.Get_Extern: Integer;
begin
  Result := DefaultInterface.Get_Extern;
end;

function  TItemOfHistoricProperties.Get_BeginDate: TDateTime;
begin
  Result := DefaultInterface.Get_BeginDate;
end;

function  TItemOfHistoricProperties.Get_EndDate: TDateTime;
begin
  Result := DefaultInterface.Get_EndDate;
end;

function  TItemOfHistoricProperties.Get_TalkingDate: TDateTime;
begin
  Result := DefaultInterface.Get_TalkingDate;
end;

{$ENDIF}

procedure TAxHistoric.InitControlData;
const
  CEventDispIDs: array [0..1] of DWORD = (
    $00000001, $00000002);
  CControlData: TControlData2 = (
    ClassID: '{1ECE7399-EAF5-11D5-B77A-00508B79E1ED}';
    EventIID: '{1ECE739A-EAF5-11D5-B77A-00508B79E1ED}';
    EventCount: 2;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnHistoDeleteItem) - Cardinal(Self);
end;

procedure TAxHistoric.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IAxHistoric;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TAxHistoric.GetControlInterface: IAxHistoric;
begin
  CreateControl;
  Result := FIntf;
end;

function  TAxHistoric.RemoveItem(IdHigh: Integer; IdLow: Integer): Integer;
begin
  Result := DefaultInterface.RemoveItem(IdHigh, IdLow);
end;

function  TAxHistoric.Connect: Integer;
begin
  Result := DefaultInterface.Connect;
end;

procedure TAxTransfertCall.InitControlData;
const
  CEventDispIDs: array [0..1] of DWORD = (
    $00000001, $00000002);
  CTPictureIDs: array [0..0] of DWORD = (
    $FFFFFDF5);
  CControlData: TControlData2 = (
    ClassID: '{C989EA6A-F2C8-11D5-B77B-00508B79E1ED}';
    EventIID: '{C989EA6B-F2C8-11D5-B77B-00508B79E1ED}';
    EventCount: 2;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401;
    FontCount: 0;
    FontIDs: nil;
    PictureCount: 1;
    PictureIDs: @CTPictureIDs);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnTransferCallPossible) - Cardinal(Self);
end;

procedure TAxTransfertCall.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IAxTransfertCall;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TAxTransfertCall.GetControlInterface: IAxTransfertCall;
begin
  CreateControl;
  Result := FIntf;
end;

function  TAxTransfertCall.TransferCall(heldCallId: Integer): Integer;
begin
  Result := DefaultInterface.TransferCall(heldCallId);
end;

function  TAxTransfertCall.Connect: Integer;
begin
  Result := DefaultInterface.Connect;
end;

procedure TAxStartStopRecording.InitControlData;
const
  CEventDispIDs: array [0..1] of DWORD = (
    $00000001, $00000002);
  CTPictureIDs: array [0..0] of DWORD = (
    $FFFFFDF5);
  CControlData: TControlData2 = (
    ClassID: '{E810B9BB-F45D-11D5-B77E-00508B79E1ED}';
    EventIID: '{E810B9BC-F45D-11D5-B77E-00508B79E1ED}';
    EventCount: 2;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401;
    FontCount: 0;
    FontIDs: nil;
    PictureCount: 1;
    PictureIDs: @CTPictureIDs);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnNiceStartRecordPossible) - Cardinal(Self);
end;

procedure TAxStartStopRecording.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IAxStartStopRecording;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TAxStartStopRecording.GetControlInterface: IAxStartStopRecording;
begin
  CreateControl;
  Result := FIntf;
end;

function  TAxStartStopRecording.NiceStartRecord(const Comment: WideString): Integer;
begin
  Result := DefaultInterface.NiceStartRecord(Comment);
end;

function  TAxStartStopRecording.NiceStopRecord(const Comment: WideString): Integer;
begin
  Result := DefaultInterface.NiceStopRecord(Comment);
end;

function  TAxStartStopRecording.UpdateCall(Comment1: OleVariant): Integer;
begin
  Result := DefaultInterface.UpdateCall(Comment1, EmptyParam, EmptyParam);
end;

function  TAxStartStopRecording.UpdateCall(Comment1: OleVariant; Comment2: OleVariant): Integer;
begin
  Result := DefaultInterface.UpdateCall(Comment1, Comment2, EmptyParam);
end;

function  TAxStartStopRecording.UpdateCall(Comment1: OleVariant; Comment2: OleVariant; 
                                           Comment3: OleVariant): Integer;
begin
  Result := DefaultInterface.UpdateCall(Comment1, Comment2, Comment3);
end;

function  TAxStartStopRecording.Connect: Integer;
begin
  Result := DefaultInterface.Connect;
end;

procedure TAxAgentStatistics.InitControlData;
const
  CEventDispIDs: array [0..0] of DWORD = (
    $00000001);
  CControlData: TControlData2 = (
    ClassID: '{47B38D58-F9ED-11D5-B77F-00508B79E1ED}';
    EventIID: '{47B38D59-F9ED-11D5-B77F-00508B79E1ED}';
    EventCount: 1;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnPropertyChanges) - Cardinal(Self);
end;

procedure TAxAgentStatistics.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IAxAgentStatistics;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TAxAgentStatistics.GetControlInterface: IAxAgentStatistics;
begin
  CreateControl;
  Result := FIntf;
end;

function  TAxAgentStatistics.Connect: Integer;
begin
  Result := DefaultInterface.Connect;
end;

procedure TAxTeamStatistics.InitControlData;
const
  CEventDispIDs: array [0..0] of DWORD = (
    $00000001);
  CControlData: TControlData2 = (
    ClassID: '{94EDB882-FA19-11D5-B77F-00508B79E1ED}';
    EventIID: '{94EDB883-FA19-11D5-B77F-00508B79E1ED}';
    EventCount: 1;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnPropertyChanges) - Cardinal(Self);
end;

procedure TAxTeamStatistics.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IAxTeamStatistics;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TAxTeamStatistics.GetControlInterface: IAxTeamStatistics;
begin
  CreateControl;
  Result := FIntf;
end;

function  TAxTeamStatistics.Connect: Integer;
begin
  Result := DefaultInterface.Connect;
end;

procedure TAxHold.InitControlData;
const
  CEventDispIDs: array [0..1] of DWORD = (
    $00000001, $00000002);
  CTPictureIDs: array [0..0] of DWORD = (
    $FFFFFDF5);
  CControlData: TControlData2 = (
    ClassID: '{4BB01D40-035A-11D6-B782-00508B79E1ED}';
    EventIID: '{4BB01D41-035A-11D6-B782-00508B79E1ED}';
    EventCount: 2;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401;
    FontCount: 0;
    FontIDs: nil;
    PictureCount: 1;
    PictureIDs: @CTPictureIDs);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnHoldCallPossible) - Cardinal(Self);
end;

procedure TAxHold.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IAxHold;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TAxHold.GetControlInterface: IAxHold;
begin
  CreateControl;
  Result := FIntf;
end;

function  TAxHold.HoldCall: Integer;
begin
  Result := DefaultInterface.HoldCall;
end;

function  TAxHold.Connect: Integer;
begin
  Result := DefaultInterface.Connect;
end;

procedure TAxDoNotCallRecord.InitControlData;
const
  CEventDispIDs: array [0..1] of DWORD = (
    $00000001, $00000002);
  CTPictureIDs: array [0..0] of DWORD = (
    $FFFFFDF5);
  CControlData: TControlData2 = (
    ClassID: '{114C6FA4-04FC-11D6-B785-00508B79E1ED}';
    EventIID: '{114C6FA5-04FC-11D6-B785-00508B79E1ED}';
    EventCount: 2;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401;
    FontCount: 0;
    FontIDs: nil;
    PictureCount: 1;
    PictureIDs: @CTPictureIDs);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnDoNotCallPossible) - Cardinal(Self);
end;

procedure TAxDoNotCallRecord.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IAxDoNotCallRecord;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TAxDoNotCallRecord.GetControlInterface: IAxDoNotCallRecord;
begin
  CreateControl;
  Result := FIntf;
end;

function  TAxDoNotCallRecord.DoNotCall(handle: Integer): Integer;
begin
  Result := DefaultInterface.DoNotCall(handle);
end;

function  TAxDoNotCallRecord.Connect: Integer;
begin
  Result := DefaultInterface.Connect;
end;

procedure TAxRequestRejectPreviewRecord.InitControlData;
const
  CEventDispIDs: array [0..2] of DWORD = (
    $00000001, $00000002, $00000003);
  CTPictureIDs: array [0..0] of DWORD = (
    $FFFFFDF5);
  CControlData: TControlData2 = (
    ClassID: '{50A0753F-0513-11D6-B785-00508B79E1ED}';
    EventIID: '{50A07540-0513-11D6-B785-00508B79E1ED}';
    EventCount: 3;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401;
    FontCount: 0;
    FontIDs: nil;
    PictureCount: 1;
    PictureIDs: @CTPictureIDs);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnPreviewRecordPossible) - Cardinal(Self);
end;

procedure TAxRequestRejectPreviewRecord.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IAxRequestRejectPreviewRecord;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TAxRequestRejectPreviewRecord.GetControlInterface: IAxRequestRejectPreviewRecord;
begin
  CreateControl;
  Result := FIntf;
end;

function  TAxRequestRejectPreviewRecord.PreviewRecord: Integer;
begin
  Result := DefaultInterface.PreviewRecord;
end;

function  TAxRequestRejectPreviewRecord.RecordReject(handle: Integer): Integer;
begin
  Result := DefaultInterface.RecordReject(handle);
end;

function  TAxRequestRejectPreviewRecord.Connect: Integer;
begin
  Result := DefaultInterface.Connect;
end;

procedure TAxScheduleCallback.InitControlData;
const
  CEventDispIDs: array [0..1] of DWORD = (
    $00000001, $00000002);
  CTPictureIDs: array [0..0] of DWORD = (
    $FFFFFDF5);
  CControlData: TControlData2 = (
    ClassID: '{50A0755F-0513-11D6-B785-00508B79E1ED}';
    EventIID: '{50A07560-0513-11D6-B785-00508B79E1ED}';
    EventCount: 2;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401;
    FontCount: 0;
    FontIDs: nil;
    PictureCount: 1;
    PictureIDs: @CTPictureIDs);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnRecordReschedulePossible) - Cardinal(Self);
end;

procedure TAxScheduleCallback.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IAxScheduleCallback;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TAxScheduleCallback.GetControlInterface: IAxScheduleCallback;
begin
  CreateControl;
  Result := FIntf;
end;

function  TAxScheduleCallback.Connect: Integer;
begin
  Result := DefaultInterface.Connect;
end;

function  TAxScheduleCallback.RescheduleRecord(handle: Integer; const callbackType: WideString; 
                                               callbackDate: OleVariant): Integer;
begin
  Result := DefaultInterface.RescheduleRecord(handle, callbackType, callbackDate);
end;

procedure TAxClassify.InitControlData;
const
  CEventDispIDs: array [0..1] of DWORD = (
    $00000001, $00000002);
  CTPictureIDs: array [0..0] of DWORD = (
    $FFFFFDF5);
  CControlData: TControlData2 = (
    ClassID: '{50A0757D-0513-11D6-B785-00508B79E1ED}';
    EventIID: '{50A0757E-0513-11D6-B785-00508B79E1ED}';
    EventCount: 2;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401;
    FontCount: 0;
    FontIDs: nil;
    PictureCount: 1;
    PictureIDs: @CTPictureIDs);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnRecordProcessedPossible) - Cardinal(Self);
end;

procedure TAxClassify.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IAxClassify;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TAxClassify.GetControlInterface: IAxClassify;
begin
  CreateControl;
  Result := FIntf;
end;

function  TAxClassify.RecordProcessed(handle: Integer): Integer;
begin
  Result := DefaultInterface.RecordProcessed(handle);
end;

function  TAxClassify.Connect: Integer;
begin
  Result := DefaultInterface.Connect;
end;

procedure Register;
begin
  RegisterComponents('CTI',[TAxApplication, TAxLine, TAxMakeCall, TAxHangUp, 
    TAxTakeCall, TAxConference, TAxLog, TAxHistoric, TAxTransfertCall, 
    TAxStartStopRecording, TAxAgentStatistics, TAxTeamStatistics, TAxHold, TAxDoNotCallRecord, 
    TAxRequestRejectPreviewRecord, TAxScheduleCallback, TAxClassify]);
  RegisterComponents('CTI',[TDataOfCall, TItemOfHistoric]);
end;

end.

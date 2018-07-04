unit PCb_TLB;

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
// Fichier généré le 10/08/2001 22:07:27 depuis la bibliothèque de types ci-dessous.

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
// Bibl.Types     : C:\PGIS5\V1\Exemples\PCB\PCB.TLB (1)
// IID\LCID       : {B693A910-5E47-11D5-B887-00400556B940}\0
// Fichier d'aide : C:\PGIS5\V1\Exemples\PCB\pcb.Chm
// DepndLst       : 
//   (1) v2.0 stdole, (C:\WINDOWS\System32\stdole2.tlb)
//   (2) v4.0 StdVCL, (C:\WINDOWS\System32\STDVCL40.DLL)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
interface

uses Windows, Classes, OleCtrls;

// *********************************************************************//
// GUIDS déclarés dans la bibliothèque de types. Préfixes utilisés :    
//   Bibliothèques de types : LIBID_xxxx                                      
//   CoClasses              : CLASS_xxxx                                      
//   DISPInterfaces         : DIID_xxxx                                       
//   Non-DISP interfaces    : IID_xxxx                                        
// *********************************************************************//
const
  // Version majeure et mineure de la bibliothèque de types
  PCbMajorVersion = 1;
  PCbMinorVersion = 0;

  LIBID_PCb: TGUID = '{B693A910-5E47-11D5-B887-00400556B940}';

  IID__ControlPCB: TGUID = '{B693A911-5E47-11D5-B887-00400556B940}';
  DIID___ControlPCB: TGUID = '{B693A915-5E47-11D5-B887-00400556B940}';
  IID__ClassPCB: TGUID = '{B693A918-5E47-11D5-B887-00400556B940}';
  CLASS_ClassPCB: TGUID = '{B693A919-5E47-11D5-B887-00400556B940}';
  CLASS_ControlPCB: TGUID = '{B693A912-5E47-11D5-B887-00400556B940}';

// *********************************************************************//
// Déclaration d'énumérations définies dans la bibliothèque de types    
// *********************************************************************//
// Constantes pour enum EtatAgentACD
type
  EtatAgentACD = TOleEnum;
const
  Agent_Pause = $00000000;
  Agent_Free = $00000001;
  Agent_Busy = $00000002;
  Agent_Disconnect = $00000003;
  Agent_Informatique = $00000004;
  Agent_Call = $00000005;
  Agent_Connection = $00000006;
  Agent_NotReady = $00000007;
  Agent_NotWorkReady = $00000008;

// Constantes pour enum CdeEtatAgent
type
  CdeEtatAgent = TOleEnum;
const
  Evt_Agent_Connect = $00000001;
  Evt_Agent_Disconnect = $00000002;
  Evt_Agent_Pause = $00000003;
  Evt_Agent_EndPause = $00000004;
  Evt_Agent_PauseAdministrative = $00000005;
  Evt_Agent_EndPauseAdministrative = $00000006;

type

// *********************************************************************//
// Déclaration Forward des types définis dans la bibliothèque de types    
// *********************************************************************//
  _ControlPCB = interface;
  _ControlPCBDisp = dispinterface;
  __ControlPCB = dispinterface;
  _ClassPCB = interface;
  _ClassPCBDisp = dispinterface;

// *********************************************************************//
// Déclaration de CoClasses définies dans la bibliothèque de types 
// (REMARQUE: On affecte chaque CoClass à son Interface par défaut)              
// *********************************************************************//
  ClassPCB = _ClassPCB;
  ControlPCB = _ControlPCB;


// *********************************************************************//
// Interface   : _ControlPCB
// Indicateurs : (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID        : {B693A911-5E47-11D5-B887-00400556B940}
// *********************************************************************//
  _ControlPCB = interface(IDispatch)
    ['{B693A911-5E47-11D5-B887-00400556B940}']
    function  AppelUnClient(const NumeroTelSortant: WideString): Smallint; safecall;
    procedure SendIdentificationPCB(RefApp: Integer; const Nom: WideString; 
                                    const Prenom: WideString; const Societe: WideString); safecall;
    procedure ReceptionMsg(const msg: WideString); safecall;
    procedure StartMode(Mode: Smallint); safecall;
    function  FctCstaMakeCall(const Number: WideString): WordBool; safecall;
    function  FctCstaConsultCall(const Number: WideString): WordBool; safecall;
    function  FctCstaReconnectCall(HoldRefCom: Integer): WordBool; safecall;
    function  FctCstaAnswerCall(RefCom: Integer): WordBool; safecall;
    function  FctCstaClearCall(RefCom: Integer): WordBool; safecall;
    function  FctCstaTransfertCall(HoldRefCom: Integer): WordBool; safecall;
    function  FctPrivateTransfertData(const Data: WideString): WordBool; safecall;
    function  FctCstaRedirectCall(HoldRefCom: Integer; const NewDestination: WideString): WordBool; safecall;
    function  FctAgentEtatACD(NewEtatACD: CdeEtatAgent): WordBool; safecall;
    procedure Quit; safecall;
    function  FctCstaHoldCall(RefCom: Integer): WordBool; safecall;
    function  FctCstaRetrieveCall(HoldRefCom: Integer): WordBool; safecall;
    function  FctAckAppelOut: WordBool; safecall;
    function  FctFinWrapUp: WordBool; safecall;
    function  TraceFileDebug(const FileDebug: WideString): OleVariant; safecall;
  end;

// *********************************************************************//
// DispIntf:  _ControlPCBDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {B693A911-5E47-11D5-B887-00400556B940}
// *********************************************************************//
  _ControlPCBDisp = dispinterface
    ['{B693A911-5E47-11D5-B887-00400556B940}']
    function  AppelUnClient(const NumeroTelSortant: WideString): Smallint; dispid 1610809350;
    procedure SendIdentificationPCB(RefApp: Integer; const Nom: WideString; 
                                    const Prenom: WideString; const Societe: WideString); dispid 1610809351;
    procedure ReceptionMsg(const msg: WideString); dispid 1610809354;
    procedure StartMode(Mode: Smallint); dispid 1610809355;
    function  FctCstaMakeCall(const Number: WideString): WordBool; dispid 1610809356;
    function  FctCstaConsultCall(const Number: WideString): WordBool; dispid 1610809357;
    function  FctCstaReconnectCall(HoldRefCom: Integer): WordBool; dispid 1610809358;
    function  FctCstaAnswerCall(RefCom: Integer): WordBool; dispid 1610809359;
    function  FctCstaClearCall(RefCom: Integer): WordBool; dispid 1610809360;
    function  FctCstaTransfertCall(HoldRefCom: Integer): WordBool; dispid 1610809361;
    function  FctPrivateTransfertData(const Data: WideString): WordBool; dispid 1610809362;
    function  FctCstaRedirectCall(HoldRefCom: Integer; const NewDestination: WideString): WordBool; dispid 1610809363;
    function  FctAgentEtatACD(NewEtatACD: CdeEtatAgent): WordBool; dispid 1610809364;
    procedure Quit; dispid 1610809365;
    function  FctCstaHoldCall(RefCom: Integer): WordBool; dispid 1610809366;
    function  FctCstaRetrieveCall(HoldRefCom: Integer): WordBool; dispid 1610809367;
    function  FctAckAppelOut: WordBool; dispid 1610809368;
    function  FctFinWrapUp: WordBool; dispid 1610809369;
    function  TraceFileDebug(const FileDebug: WideString): OleVariant; dispid 1610809370;
  end;

// *********************************************************************//
// DispIntf:  __ControlPCB
// Flags:     (4240) Hidden NonExtensible Dispatchable
// GUID:      {B693A915-5E47-11D5-B887-00400556B940}
// *********************************************************************//
  __ControlPCB = dispinterface
    ['{B693A915-5E47-11D5-B887-00400556B940}']
    procedure AppelTelephonique(const appelant: WideString; const appele: WideString); dispid 1;
    procedure CommunicationTelephonique(const appelant: WideString; const appele: WideString); dispid 2;
    procedure FinCommunicationTelephonique(const appelant: WideString; const appele: WideString); dispid 3;
    procedure AppelIdClientEntrant(const idclient: WideString; const appele: WideString); dispid 4;
    procedure CommunicationIdClientEntrant(const idclient: WideString; const appele: WideString); dispid 5;
    procedure ResultatAppel(Resultat: Smallint; const NumeroTelephone: WideString); dispid 6;
    procedure RequestIdentification(RefAppel: Integer; const appelant: WideString; 
                                    const appele: WideString); dispid 7;
    procedure EtatConnectionPCB(EtatPCB: WordBool); dispid 8;
    procedure JournalAppelant(const appelant: WideString; const appele: WideString); dispid 9;
    procedure CstaServiceInitiated(RefCom: Integer); dispid 10;
    procedure CstaFailed(RefCom: Integer); dispid 11;
    procedure CstaConnectionClear(RefCom: Integer); dispid 12;
    procedure CstaOriginated(RefCom: Integer; const Correspondant: WideString); dispid 13;
    procedure CstaDelivered(RefCom: Integer; const Correspondant: WideString); dispid 14;
    procedure CstaEstablished(RefCom: Integer; const Correspondant: WideString; 
                              const appele: WideString); dispid 15;
    procedure CstaQueue(RefCom: Integer; const Correspondant: WideString); dispid 16;
    procedure CstaDiverted(RefCom: Integer; const Correspondant: WideString); dispid 17;
    procedure CstaAlerting(RefCom: Integer; const Correspondant: WideString; 
                           const appele: WideString); dispid 18;
    procedure CstaClearCall(RefCom: Integer; const Correspondant: WideString); dispid 19;
    procedure CstaConference(RefCom: Integer); dispid 20;
    procedure CstaHoldMe(RefCom: Integer; const Correspondant: WideString); dispid 21;
    procedure CstaNetworkReached(RefCom: Integer; const Correspondant: WideString); dispid 22;
    procedure CstaRetrieved(RefCom: Integer; const Correspondant: WideString); dispid 23;
    procedure CstaHoldOther(RefCom: Integer; const Correspondant: WideString); dispid 24;
    procedure CstaTransferred(oldRefcom: Integer; const Correspondant: WideString; 
                              NewRefCom: Integer); dispid 25;
    procedure TransferedData(RefCom: Integer; const Data: WideString); dispid 26;
    procedure EvtAgentACD(EtatACD: EtatAgentACD); dispid 28;
    procedure AutomateAppel(const TypeCampagne: WideString; const NomCampagne: WideString; 
                            const TelephoneOut: WideString; const InfoOut: WideString); dispid 27;
  end;

// *********************************************************************//
// Interface   : _ClassPCB
// Indicateurs : (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID        : {B693A918-5E47-11D5-B887-00400556B940}
// *********************************************************************//
  _ClassPCB = interface(IDispatch)
    ['{B693A918-5E47-11D5-B887-00400556B940}']
    function  SendOCX(var msg: WideString): OleVariant; safecall;
  end;

// *********************************************************************//
// DispIntf:  _ClassPCBDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {B693A918-5E47-11D5-B887-00400556B940}
// *********************************************************************//
  _ClassPCBDisp = dispinterface
    ['{B693A918-5E47-11D5-B887-00400556B940}']
    function  SendOCX(var msg: WideString): OleVariant; dispid 1610809345;
  end;

// *********************************************************************//
// La classe CoClassPCB fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut _ClassPCB exposée             
// pas la CoClass ClassPCB. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClass exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoClassPCB = class
    class function Create: _ClassPCB;
    class function CreateRemote(const MachineName: string): _ClassPCB;
  end;


// *********************************************************************//
// Déclaration de classe proxy de contrôle OLE
// Nom du contrôle      : TControlPCB
// Chaîne d'aide        : 
// Interface par défaut : _ControlPCB
// DISP Int. Déf. ?     : No
// Interface événements : __ControlPCB
// TypeFlags            : (32) Control
// *********************************************************************//
  TControlPCBAppelTelephonique = procedure(Sender: TObject; const appelant: WideString; 
                                                            const appele: WideString) of object;
  TControlPCBCommunicationTelephonique = procedure(Sender: TObject; const appelant: WideString; 
                                                                    const appele: WideString) of object;
  TControlPCBFinCommunicationTelephonique = procedure(Sender: TObject; const appelant: WideString; 
                                                                       const appele: WideString) of object;
  TControlPCBAppelIdClientEntrant = procedure(Sender: TObject; const idclient: WideString; 
                                                               const appele: WideString) of object;
  TControlPCBCommunicationIdClientEntrant = procedure(Sender: TObject; const idclient: WideString; 
                                                                       const appele: WideString) of object;
  TControlPCBResultatAppel = procedure(Sender: TObject; Resultat: Smallint; 
                                                        const NumeroTelephone: WideString) of object;
  TControlPCBRequestIdentification = procedure(Sender: TObject; RefAppel: Integer; 
                                                                const appelant: WideString; 
                                                                const appele: WideString) of object;
  TControlPCBEtatConnectionPCB = procedure(Sender: TObject; EtatPCB: WordBool) of object;
  TControlPCBJournalAppelant = procedure(Sender: TObject; const appelant: WideString; 
                                                          const appele: WideString) of object;
  TControlPCBCstaServiceInitiated = procedure(Sender: TObject; RefCom: Integer) of object;
  TControlPCBCstaFailed = procedure(Sender: TObject; RefCom: Integer) of object;
  TControlPCBCstaConnectionClear = procedure(Sender: TObject; RefCom: Integer) of object;
  TControlPCBCstaOriginated = procedure(Sender: TObject; RefCom: Integer; 
                                                         const Correspondant: WideString) of object;
  TControlPCBCstaDelivered = procedure(Sender: TObject; RefCom: Integer; 
                                                        const Correspondant: WideString) of object;
  TControlPCBCstaEstablished = procedure(Sender: TObject; RefCom: Integer; 
                                                          const Correspondant: WideString; 
                                                          const appele: WideString) of object;
  TControlPCBCstaQueue = procedure(Sender: TObject; RefCom: Integer; const Correspondant: WideString) of object;
  TControlPCBCstaDiverted = procedure(Sender: TObject; RefCom: Integer; 
                                                       const Correspondant: WideString) of object;
  TControlPCBCstaAlerting = procedure(Sender: TObject; RefCom: Integer; 
                                                       const Correspondant: WideString; 
                                                       const appele: WideString) of object;
  TControlPCBCstaClearCall = procedure(Sender: TObject; RefCom: Integer; 
                                                        const Correspondant: WideString) of object;
  TControlPCBCstaConference = procedure(Sender: TObject; RefCom: Integer) of object;
  TControlPCBCstaHoldMe = procedure(Sender: TObject; RefCom: Integer; 
                                                     const Correspondant: WideString) of object;
  TControlPCBCstaNetworkReached = procedure(Sender: TObject; RefCom: Integer; 
                                                             const Correspondant: WideString) of object;
  TControlPCBCstaRetrieved = procedure(Sender: TObject; RefCom: Integer; 
                                                        const Correspondant: WideString) of object;
  TControlPCBCstaHoldOther = procedure(Sender: TObject; RefCom: Integer; 
                                                        const Correspondant: WideString) of object;
  TControlPCBCstaTransferred = procedure(Sender: TObject; oldRefcom: Integer; 
                                                          const Correspondant: WideString; 
                                                          NewRefCom: Integer) of object;
  TControlPCBTransferedData = procedure(Sender: TObject; RefCom: Integer; const Data: WideString) of object;
  TControlPCBEvtAgentACD = procedure(Sender: TObject; EtatACD: EtatAgentACD) of object;
  TControlPCBAutomateAppel = procedure(Sender: TObject; const TypeCampagne: WideString; 
                                                        const NomCampagne: WideString; 
                                                        const TelephoneOut: WideString; 
                                                        const InfoOut: WideString) of object;

  TControlPCB = class(TOleControl)
  private
    FOnAppelTelephonique: TControlPCBAppelTelephonique;
    FOnCommunicationTelephonique: TControlPCBCommunicationTelephonique;
    FOnFinCommunicationTelephonique: TControlPCBFinCommunicationTelephonique;
    FOnAppelIdClientEntrant: TControlPCBAppelIdClientEntrant;
    FOnCommunicationIdClientEntrant: TControlPCBCommunicationIdClientEntrant;
    FOnResultatAppel: TControlPCBResultatAppel;
    FOnRequestIdentification: TControlPCBRequestIdentification;
    FOnEtatConnectionPCB: TControlPCBEtatConnectionPCB;
    FOnJournalAppelant: TControlPCBJournalAppelant;
    FOnCstaServiceInitiated: TControlPCBCstaServiceInitiated;
    FOnCstaFailed: TControlPCBCstaFailed;
    FOnCstaConnectionClear: TControlPCBCstaConnectionClear;
    FOnCstaOriginated: TControlPCBCstaOriginated;
    FOnCstaDelivered: TControlPCBCstaDelivered;
    FOnCstaEstablished: TControlPCBCstaEstablished;
    FOnCstaQueue: TControlPCBCstaQueue;
    FOnCstaDiverted: TControlPCBCstaDiverted;
    FOnCstaAlerting: TControlPCBCstaAlerting;
    FOnCstaClearCall: TControlPCBCstaClearCall;
    FOnCstaConference: TControlPCBCstaConference;
    FOnCstaHoldMe: TControlPCBCstaHoldMe;
    FOnCstaNetworkReached: TControlPCBCstaNetworkReached;
    FOnCstaRetrieved: TControlPCBCstaRetrieved;
    FOnCstaHoldOther: TControlPCBCstaHoldOther;
    FOnCstaTransferred: TControlPCBCstaTransferred;
    FOnTransferedData: TControlPCBTransferedData;
    FOnEvtAgentACD: TControlPCBEvtAgentACD;
    FOnAutomateAppel: TControlPCBAutomateAppel;
    FIntf: _ControlPCBDisp;
    function  GetControlInterface: _ControlPCBDisp;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    function  AppelUnClient(const NumeroTelSortant: WideString): Smallint;
    procedure SendIdentificationPCB(RefApp: Integer; const Nom: WideString; 
                                    const Prenom: WideString; const Societe: WideString);
    procedure ReceptionMsg(const msg: WideString);
    procedure StartMode(Mode: Smallint);
    function  FctCstaMakeCall(const Number: WideString): WordBool;
    function  FctCstaConsultCall(const Number: WideString): WordBool;
    function  FctCstaReconnectCall(HoldRefCom: Integer): WordBool;
    function  FctCstaAnswerCall(RefCom: Integer): WordBool;
    function  FctCstaClearCall(RefCom: Integer): WordBool;
    function  FctCstaTransfertCall(HoldRefCom: Integer): WordBool;
    function  FctPrivateTransfertData(const Data: WideString): WordBool;
    function  FctCstaRedirectCall(HoldRefCom: Integer; const NewDestination: WideString): WordBool;
    function  FctAgentEtatACD(NewEtatACD: CdeEtatAgent): WordBool;
    procedure Quit;
    function  FctCstaHoldCall(RefCom: Integer): WordBool;
    function  FctCstaRetrieveCall(HoldRefCom: Integer): WordBool;
    function  FctAckAppelOut: WordBool;
    function  FctFinWrapUp: WordBool;
    function  TraceFileDebug(const FileDebug: WideString): OleVariant;
    property  ControlInterface: _ControlPCBDisp read GetControlInterface;
    property  DefaultInterface: _ControlPCBDisp read GetControlInterface;
  published
    property OnAppelTelephonique: TControlPCBAppelTelephonique read FOnAppelTelephonique write FOnAppelTelephonique;
    property OnCommunicationTelephonique: TControlPCBCommunicationTelephonique read FOnCommunicationTelephonique write FOnCommunicationTelephonique;
    property OnFinCommunicationTelephonique: TControlPCBFinCommunicationTelephonique read FOnFinCommunicationTelephonique write FOnFinCommunicationTelephonique;
    property OnAppelIdClientEntrant: TControlPCBAppelIdClientEntrant read FOnAppelIdClientEntrant write FOnAppelIdClientEntrant;
    property OnCommunicationIdClientEntrant: TControlPCBCommunicationIdClientEntrant read FOnCommunicationIdClientEntrant write FOnCommunicationIdClientEntrant;
    property OnResultatAppel: TControlPCBResultatAppel read FOnResultatAppel write FOnResultatAppel;
    property OnRequestIdentification: TControlPCBRequestIdentification read FOnRequestIdentification write FOnRequestIdentification;
    property OnEtatConnectionPCB: TControlPCBEtatConnectionPCB read FOnEtatConnectionPCB write FOnEtatConnectionPCB;
    property OnJournalAppelant: TControlPCBJournalAppelant read FOnJournalAppelant write FOnJournalAppelant;
    property OnCstaServiceInitiated: TControlPCBCstaServiceInitiated read FOnCstaServiceInitiated write FOnCstaServiceInitiated;
    property OnCstaFailed: TControlPCBCstaFailed read FOnCstaFailed write FOnCstaFailed;
    property OnCstaConnectionClear: TControlPCBCstaConnectionClear read FOnCstaConnectionClear write FOnCstaConnectionClear;
    property OnCstaOriginated: TControlPCBCstaOriginated read FOnCstaOriginated write FOnCstaOriginated;
    property OnCstaDelivered: TControlPCBCstaDelivered read FOnCstaDelivered write FOnCstaDelivered;
    property OnCstaEstablished: TControlPCBCstaEstablished read FOnCstaEstablished write FOnCstaEstablished;
    property OnCstaQueue: TControlPCBCstaQueue read FOnCstaQueue write FOnCstaQueue;
    property OnCstaDiverted: TControlPCBCstaDiverted read FOnCstaDiverted write FOnCstaDiverted;
    property OnCstaAlerting: TControlPCBCstaAlerting read FOnCstaAlerting write FOnCstaAlerting;
    property OnCstaClearCall: TControlPCBCstaClearCall read FOnCstaClearCall write FOnCstaClearCall;
    property OnCstaConference: TControlPCBCstaConference read FOnCstaConference write FOnCstaConference;
    property OnCstaHoldMe: TControlPCBCstaHoldMe read FOnCstaHoldMe write FOnCstaHoldMe;
    property OnCstaNetworkReached: TControlPCBCstaNetworkReached read FOnCstaNetworkReached write FOnCstaNetworkReached;
    property OnCstaRetrieved: TControlPCBCstaRetrieved read FOnCstaRetrieved write FOnCstaRetrieved;
    property OnCstaHoldOther: TControlPCBCstaHoldOther read FOnCstaHoldOther write FOnCstaHoldOther;
    property OnCstaTransferred: TControlPCBCstaTransferred read FOnCstaTransferred write FOnCstaTransferred;
    property OnTransferedData: TControlPCBTransferedData read FOnTransferedData write FOnTransferedData;
    property OnEvtAgentACD: TControlPCBEvtAgentACD read FOnEvtAgentACD write FOnEvtAgentACD;
    property OnAutomateAppel: TControlPCBAutomateAppel read FOnAutomateAppel write FOnAutomateAppel;
  end;

procedure Register;

implementation

uses ComObj;

class function CoClassPCB.Create: _ClassPCB;
begin
  Result := CreateComObject(CLASS_ClassPCB) as _ClassPCB;
end;

class function CoClassPCB.CreateRemote(const MachineName: string): _ClassPCB;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ClassPCB) as _ClassPCB;
end;

procedure TControlPCB.InitControlData;
const
  CEventDispIDs: array [0..27] of DWORD = (
    $00000001, $00000002, $00000003, $00000004, $00000005, $00000006,
    $00000007, $00000008, $00000009, $0000000A, $0000000B, $0000000C,
    $0000000D, $0000000E, $0000000F, $00000010, $00000011, $00000012,
    $00000013, $00000014, $00000015, $00000016, $00000017, $00000018,
    $00000019, $0000001A, $0000001C, $0000001B);
  CControlData: TControlData2 = (
    ClassID: '{B693A912-5E47-11D5-B887-00400556B940}';
    EventIID: '{B693A915-5E47-11D5-B887-00400556B940}';
    EventCount: 28;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80040154*);
    Flags: $00000000;
    Version: 401);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnAppelTelephonique) - Cardinal(Self);
end;

procedure TControlPCB.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as _ControlPCBDisp;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TControlPCB.GetControlInterface: _ControlPCBDisp;
begin
  CreateControl;
  Result := FIntf;
end;

function  TControlPCB.AppelUnClient(const NumeroTelSortant: WideString): Smallint;
begin
  Result := DefaultInterface.AppelUnClient(NumeroTelSortant);
end;

procedure TControlPCB.SendIdentificationPCB(RefApp: Integer; const Nom: WideString; 
                                            const Prenom: WideString; const Societe: WideString);
begin
  DefaultInterface.SendIdentificationPCB(RefApp, Nom, Prenom, Societe);
end;

procedure TControlPCB.ReceptionMsg(const msg: WideString);
begin
  DefaultInterface.ReceptionMsg(msg);
end;

procedure TControlPCB.StartMode(Mode: Smallint);
begin
  DefaultInterface.StartMode(Mode);
end;

function  TControlPCB.FctCstaMakeCall(const Number: WideString): WordBool;
begin
  Result := DefaultInterface.FctCstaMakeCall(Number);
end;

function  TControlPCB.FctCstaConsultCall(const Number: WideString): WordBool;
begin
  Result := DefaultInterface.FctCstaConsultCall(Number);
end;

function  TControlPCB.FctCstaReconnectCall(HoldRefCom: Integer): WordBool;
begin
  Result := DefaultInterface.FctCstaReconnectCall(HoldRefCom);
end;

function  TControlPCB.FctCstaAnswerCall(RefCom: Integer): WordBool;
begin
  Result := DefaultInterface.FctCstaAnswerCall(RefCom);
end;

function  TControlPCB.FctCstaClearCall(RefCom: Integer): WordBool;
begin
  Result := DefaultInterface.FctCstaClearCall(RefCom);
end;

function  TControlPCB.FctCstaTransfertCall(HoldRefCom: Integer): WordBool;
begin
  Result := DefaultInterface.FctCstaTransfertCall(HoldRefCom);
end;

function  TControlPCB.FctPrivateTransfertData(const Data: WideString): WordBool;
begin
  Result := DefaultInterface.FctPrivateTransfertData(Data);
end;

function  TControlPCB.FctCstaRedirectCall(HoldRefCom: Integer; const NewDestination: WideString): WordBool;
begin
  Result := DefaultInterface.FctCstaRedirectCall(HoldRefCom, NewDestination);
end;

function  TControlPCB.FctAgentEtatACD(NewEtatACD: CdeEtatAgent): WordBool;
begin
  Result := DefaultInterface.FctAgentEtatACD(NewEtatACD);
end;

procedure TControlPCB.Quit;
begin
  DefaultInterface.Quit;
end;

function  TControlPCB.FctCstaHoldCall(RefCom: Integer): WordBool;
begin
  Result := DefaultInterface.FctCstaHoldCall(RefCom);
end;

function  TControlPCB.FctCstaRetrieveCall(HoldRefCom: Integer): WordBool;
begin
  Result := DefaultInterface.FctCstaRetrieveCall(HoldRefCom);
end;

function  TControlPCB.FctAckAppelOut: WordBool;
begin
  Result := DefaultInterface.FctAckAppelOut;
end;

function  TControlPCB.FctFinWrapUp: WordBool;
begin
  Result := DefaultInterface.FctFinWrapUp;
end;

function  TControlPCB.TraceFileDebug(const FileDebug: WideString): OleVariant;
begin
  Result := DefaultInterface.TraceFileDebug(FileDebug);
end;

procedure Register;
begin
  RegisterComponents('ActiveX',[TControlPCB]);
end;

end.

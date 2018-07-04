unit ctiConst;

interface

const
     // Codes d'erreur lors des appels cti (ou sur le gestionnaire d'événements CTI)
     ctiStatus_OK               = LongInt ($00000000);
     ctiStatus_Error            = LongInt ($80000001);

     // Codes d'origine: appel émis ou appel reçu
     ctiOrigin_Unknown               = $00000000;
     ctiOrigin_Incoming              = $00000001;
     ctiOrigin_Outgoing              = $00000002;

     // Etat d'un appel (en fait identique aux différents états d'un appel Tapi32)
     ctiState_Idle                   = $00000001;
     ctiState_Offering               = $00000002;
     ctiState_Accepted               = $00000004;
     ctiState_Dialtone               = $00000008;
     ctiState_Dialing                = $00000010;
     ctiState_Ringback               = $00000020;
     ctiState_Busy                   = $00000040;
     ctiState_SpecialInfo            = $00000080;
     ctiState_Connected              = $00000100;
     ctiState_Proceeding             = $00000200;
     ctiState_OnHold                 = $00000400;
     ctiState_Conferenced            = $00000800;
     ctiState_OnHoldPendConf         = $00001000;
     ctiState_OnHoldPendTransfer     = $00002000;
     ctiState_Disconnected           = $00004000;
     ctiState_Unknown                = $00008000;

     // Complément d'information sur l'état d'un appel (un peu comme Tapi et ses infos supplémentaires lors d'un événement de type CallInfoState)
     ctiStateC_None                  = $00000000;
     ctiStateC_DisconnectedNormal    = $00000001;
     ctiStateC_Unknown               = $00000002;
     ctiStateC_DisconnectedBusy      = $00000020;
     ctiStateC_DisconnectedNoAnswer  = $00000040;

     // Codes événement CTI applicatifs
     ctiEvent_CallNew                = $00000001;
     ctiEvent_CallDestroy            = $00000002;
     ctiEvent_CallOffering           = $00000003;
     ctiEvent_CallConnect            = $00000004;
     ctiEvent_CallDisconnect         = $00000005;
     ctiEvent_CallHold               = $00000006;
     ctiEvent_CallConference         = $00000007;
     ctiEvent_CallDialTone           = $00000008;
     ctiEvent_CallDialing            = $00000009;
     ctiEvent_CallRingBack           = $0000000A;
     ctiEvent_CallProceeding         = $0000000B;
     ctiEvent_CallBusy               = $0000000C;

     ctiEvent_CallInfoCallerId       = $00000010;
     ctiEvent_CallInfoCalledId       = $00000011;
     ctiEvent_CallInfoConnectedId    = $00000012;
     ctiEvent_CallInfoOrigin         = $00000013;
     ctiEvent_CallInfoReason         = $00000014;

     ctiEvent_ConnectorInactive      = $00000100;
     ctiEvent_ConnectorActive        = $00000101;

     // Codes capacités de CTI
     ctiCaps_CallMake                = $00000001;
     ctiCaps_CallAnswer              = $00000002;
     ctiCaps_CallDisconnect          = $00000004;
     ctiCaps_CallRedirect            = $00000008;
     ctiCaps_CallTransfer            = $00000010;
     ctiCaps_CallConference          = $00000020;
     ctiCaps_CallSwap                = $00000040;

type
     TCtiEvent = function (lCall:LongInt; lEvent:LongInt; lInfo:LongInt; sInfo:pchar):LongInt;


implementation

end.


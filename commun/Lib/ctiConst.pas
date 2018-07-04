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
     // Pour l'instant, seuls les compléments sur l'état ctiState_Disconnected sont référencés
     ctiStateC_None                     = $00000000;
     ctiStateC_DisconnectNormal         = $00000001;
     ctiStateC_DisconnectUnknown        = $00000002;
     ctiStateC_DisconnectBusy           = $00000020;
     ctiStateC_DisconnectNoAnswer       = $00000040;

     // $$$ JP 23/08/07: new, on en a besoin pour affiner la raison de la déconnexion
     ctiStateC_DisconnectReject         = $00000004;
     ctiStateC_DisconnectForwarded      = $00000010;
     ctiStateC_DisconnectBadAddress     = $00000080;
     ctiStateC_DisconnectUnreachable    = $00000100;
     ctiStateC_DisconnectCongestion     = $00000200;
     ctiStateC_DisconnectNoDialTone     = $00001000;
     ctiStateC_DisconnectNumberChanged  = $00002000;
     ctiStateC_DisconnectOutOfOrder     = $00004000;
     ctiStateC_DisconnectTempFailure    = $00008000;
     ctiStateC_DisconnectQosUnvail      = $00010000;
     ctiStateC_DisconnectBlocked        = $00020000;
     ctiStateC_DisconnectDoNotDisturb   = $00040000;
     ctiStateC_DisconnectCancelled      = $00080000;

{
LINEDISCONNECTMODE_BADADDRESS
The destination address is invalid.

LINEDISCONNECTMODE_BLOCKED
The call could not be connected because calls from the origination address are not being accepted at the destination address.
This differs from LINEDISCONNECTMODE_REJECT in that blocking is implemented in the network (a passive reject)
while a rejection is implemented in the destination equipment (an active reject). The blocking can be due to a specific exclusion
of the origination address, or because the destination accepts calls from only a selected set of origination address (closed user group).
It's appropriate as a blacklisted response. For example, a modem has received an answer, gone more than six seconds without detecting Ringback, failed to connect a defined number of times, determines that the phone number is not valid to call, and issues a 'blacklisted' response.

LINEDISCONNECTMODE_BUSY
The remote user's station is busy.

LINEDISCONNECTMODE_CANCELLED
The call was cancelled. (TAPI versions 2.0 and later)

LINEDISCONNECTMODE_CONGESTION
The network is congested.

LINEDISCONNECTMODE_DONOTDISTURB
The call could not be connected because the destination has invoked the Do Not Disturb feature. (TAPI versions 2.0 and later)

LINEDISCONNECTMODE_FORWARDED
The call was forwarded by the switch.

LINEDISCONNECTMODE_INCOMPATIBLE
The remote user's station equipment is incompatible with the type of call requested.

LINEDISCONNECTMODE_NOANSWER
The remote user's station does not answer.

LINEDISCONNECTMODE_NODIALTONE
A dial tone was not detected within a service-provider defined timeout, at a point during dialing when one was expected
(such as at a "W" in the dialable string). This can also occur without a service-provider-defined timeout period or without
a value specified in the dwWaitForDialTone member of the LINEDIALPARAMS structure. (TAPI versions 1.4 and later)

LINEDISCONNECTMODE_NORMAL
This is a normal disconnect request by the remote party. The call was terminated normally.

LINEDISCONNECTMODE_NUMBERCHANGED
The call could not be connected because the destination number has been changed, but automatic redirection to the new number is not provided. (TAPI versions 2.0 and later)

LINEDISCONNECTMODE_OUTOFORDER
The call could not be connected or was disconnected because the destination device is out of order (hardware failure). (TAPI versions 2.0 and later)

LINEDISCONNECTMODE_PICKUP
The call was picked up from elsewhere.

LINEDISCONNECTMODE_QOSUNAVAIL
The call could not be connected or was disconnected because the minimum quality of service could not be obtained or sustained.
This differs from LINEDISCONNECTMODE_INCOMPATIBLE in that the lack of resources may be a temporary condition at the destination. (TAPI versions 2.0 and later)

LINEDISCONNECTMODE_REJECT
The remote user has rejected the call.

LINEDISCONNECTMODE_TEMPFAILURE
The call could not be connected or was disconnected because of a temporary failure in the network; the call can be reattempted later and is expected to eventually complete. (TAPI versions 2.0 and later)
It's appropriate as a delayed response. For example, a modem getting a busy signal or equivalent too many times in a particular time period
concludes that the number should not be called again until a defined time has elapsed and issues a 'delayed' response.

LINEDISCONNECTMODE_UNAVAIL
The reason for the disconnect is unavailable and will not become known later.

LINEDISCONNECTMODE_UNKNOWN
The reason for the disconnect request is unknown but may become known later.

LINEDISCONNECTMODE_UNREACHABLE
The remote user could not be reached.
}

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
     ctiEvent_CallAccepted           = $0000000D;
     ctiEvent_CallInfoChanged        = $0000000F;

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

     // Constantes pour log
     ctiLog_AppendLine               = Word ($0000);
     ctiLog_UpdateLastLine           = Word ($0001);
     ctiLog_ConcatLastLine           = Word ($0002);
     ctiLog_Clear                    = Word ($0003);
     ctiLog_Write                    = Word ($0004);

type
     TCtiEvent    = function (lCall:LongInt; lEvent:LongInt; lInfo:LongInt; sInfo:pchar):LongInt;

     // $$$ JP 02/07/07
     TCtiLogEvent = procedure (ctiLogParam:pchar; ctiLogCommand:Word); // iLine:integer);


implementation

end.







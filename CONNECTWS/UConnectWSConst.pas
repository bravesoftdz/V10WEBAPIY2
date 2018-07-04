unit UConnectWSConst;

interface

type
  T_WSEntryType      = (wsetNone, wsetDocument, wsetPayment, wsetPayer, wsetExtourne, wsetSubContractPayment, wsetStock);
  T_WSPSocType       = (wspsNone, wspsServer, wspsPort, wspsFolder, wspsLastSynchro);
  T_WSDataService    = (  wsdsNone               // Aucun                Doc génération fichier .TRA
                        , wsdsThird              // Tiers                page 34
                        , wsdsAnalyticalSection  // Section analytique   page 25
                        , wsdsAccount            // Comptes comptable
                        , wsdsJournal            // Journaux comptable
                        , wsdsBankIdentification // RIB                  page 51
                        , wsdsChoixCod           // Table CHOIXCOD       page 64 (GDM;NVR;GOR;JUR;LGU;GCT;RTV;GZC;SCC;TRC)
                        , wsdsCommon             // Table COMMUN
                        , wsdsRecovery           // Relance
                        , wsdsCountry            // Pays                 page 
                        , wsdsCurrency           // Devise               page 22
                        , wsdsChangeRate         // Taux de change       page 77
                        , wsdsCorrespondence     // Correspondance
                        , wsdsPaymenChoice       // Mode de règlement    page 20
                        , wsdsFiscalYear         // Exercice comptable
                        , wsdsSocietyParameters  // Paramètre société
                        , wsdsEstablishment      // Etablissement
                        , wsdsPaymentMode        // Mode de paiment
                        , wsdsZipCode            // Code postaux
                        , wsdsContact            // Contact
                        , wsdsFieldsList         // Liste des champs
                        );
  T_WSInfoFromDSType = (wsidNone, wsidTableName, wsidFieldsKey, wsidExcludeFields, wsidFieldsList, wsidRequest);
  T_WSAction         = (wsacNone, wsacUpdate, wsacInsert);
  T_WSBTPValues      = Record
                         ConnectionName : string;
                         UserAdmin      : string;
                         Server         : string;
                         DataBase       : string;
                         LastSynchro    : string;
                       end;
  T_WSY2Values       = Record
                         ConnectionName : string;
                         Server         : string;
                         DataBase       : string;
                       end;
  T_WSConnectionValues = Record
                           UserAdmin   : string;
                           BTPServer   : string;
                           BTPDataBase : string;
                           BTPLastSync : string;
                           Y2Server    : string;
                           Y2DataBase  : string;
                           Y2LastSync  : string;
                         end;
  T_WSLogValues       = Record
                          LogLevel          : integer;
                          LogMoMaxSize      : double;
                          LogMaxQty         : integer;
                          LogDebug          : integer;
                          LogDebugMoMaxSize : double;
                       end;

const
  WSCDS_UpdDateFieldName       = 'DateModif';
  WSCDS_CreDateFieldName       = 'DateCreation';
  WSCDS_SocLastSync            = 'SO_BTWSLASTSYNC';
  WSCDS_SocCegidDos            = 'SO_BTWSCEGIDDOS';
  WSCDS_SocNumPort             = 'SO_BTWSCEGIDPORT';
  WSCDS_SocServer              = 'SO_BTWSSERVEUR';
  WSCDS_EmptyValue             = '#None#';
  WSCDS_GetDataOk              = 'OK';
  WSCDS_IndiceField            = '#INDICE';
  WSCDS_NextUrlValue           = '"@odata.nextLink":"';
  WSCDS_SectionGlobalSettings  = 'GLOBALSETTINGS';
  WSCDS_SectionConnection      = 'CONNECTION';
  WSCDS_SectionUpdateFrequency = 'UPDATEFREQUENCY';
  WSCDS_EndUrlEntries          = 'entries';
  WSCDS_EndUrlUploadBytes      = 'importEntries/upload/bytes';

implementation


end.

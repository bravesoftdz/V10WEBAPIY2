unit Edisys_IULM_Core_TLB;

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

// PASTLWTR : 1.2
// Fichier généré le 26/02/2014 11:17:05 depuis la bibliothèque de types ci-dessous.

// ************************************************************************  //
// Bibl. types : C:\Program Files\Edisys\IULM\x86\Edisys.IULM.Core.tlb (1)
// LIBID: {D3660F56-C7EA-499C-AF18-A4F2FD5E3435}
// LCID: 0
// Fichier d'aide : 
// Chaîne d'aide : 
// DepndLst: 
//   (1) v2.0 stdole, (C:\Windows\system32\stdole2.tlb)
//   (2) v2.0 mscorlib, (C:\Windows\Microsoft.NET\Framework\v2.0.50727\mscorlib.tlb)
// Bibliothèque de types parent :
//   (0) v3.2 Edisys_IULM_Alert, (C:\Program Files\Edisys\IULM\x86\Edisys.IULM.Alert.tlb)
// Erreurs :
//   Conseil : Membre 'Unit' de 'IPriceItem' modifié en 'Unit_'
//   Remarque : le symbole 'Type' a été renommé en 'type_'
//   Remarque : le symbole 'Type' a été renommé en 'type_'
//   Remarque : le symbole 'Type' a été renommé en 'type_'
//   Remarque : le symbole 'Type' a été renommé en 'type_'
//   Remarque : le symbole 'Type' a été renommé en 'type_'
// ************************************************************************ //
// *************************************************************************//
// REMARQUE :                                                                      
// Les éléments gardés par $IFDEF_LIVE_SERVER_AT_DESIGN_TIME sont  utilisés 
// par des propriétés qui renvoient des objets pouvant nécessiter d'être  
// explicitement créés par un appel de fonction avant tout accès via la   
// propriété. Ces éléments ont été désactivés afin de prévenir une utilisation  
// accidentelle depuis l'inspecteur d'objets. Vous pouvez les activer en définissant  
// LIVE_SERVER_AT_DESIGN_TIME ou en les retirant sélectivement des blocs 
//  $IFDEF. Cependant, ces éléments doivent toujours être créés par programmation
//  via une méthode de la CoClasse appropriée avant d'être utilisés.                          
{$TYPEDADDRESS OFF} // L'unité doit être compilée sans pointeur à type contrôlé. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, mscorlib_TLB, OleServer, StdVCL, Variants;
  


// *********************************************************************//
// GUIDS déclarés dans la bibliothèque de types. Préfixes utilisés :    
//   Bibliothèques de types : LIBID_xxxx                                      
//   CoClasses              : CLASS_xxxx                                      
//   DISPInterfaces         : DIID_xxxx                                       
//   Non-DISP interfaces    : IID_xxxx                                        
// *********************************************************************//
const
  // Versions majeure et mineure de la bibliothèque de types
  Edisys_IULM_CoreMajorVersion = 3;
  Edisys_IULM_CoreMinorVersion = 2;

  LIBID_Edisys_IULM_Core: TGUID = '{D3660F56-C7EA-499C-AF18-A4F2FD5E3435}';

  DIID_IDocumentViewer: TGUID = '{15A10AB3-BA7D-46C7-BD9E-37C1CDB1D643}';
  DIID_IReportMessageCollection: TGUID = '{B34023FC-D78F-4C2E-8D0C-B6B8E8D6181B}';
  CLASS_ReportMessageCollection: TGUID = '{727E7EED-5827-4CAD-9DC9-99DEE0C91C4C}';
  DIID_IProcessingUserConfig: TGUID = '{355FDA9D-B6EF-4FF5-B9BD-22E5753FC18E}';
  DIID_IWorkItemType: TGUID = '{C9A834F0-A70C-4A4C-A21E-9EACA9C3A9B6}';
  DIID_IVersion: TGUID = '{77C83E13-6254-472C-A328-B7C1DE04D17F}';
  DIID_IProxySettings: TGUID = '{F0961E06-04BB-4796-9EEF-A71ABA2B6696}';
  DIID_IDetectionEngine: TGUID = '{1E9D57E3-67C8-4A57-8F93-BC05A761CBA9}';
  DIID_IDetectionEvents: TGUID = '{23A7D2C4-050F-4B90-80B0-F5DD7119AD0C}';
  DIID_IDetectionEventProvider: TGUID = '{8BA7B2B0-B8E0-473F-B1C6-2409A83A2917}';
  DIID_IDetectionAlerts: TGUID = '{7BA54285-A051-4ACF-B748-61172525CDAD}';
  DIID_IDIEDownloadEventProvider: TGUID = '{2C67F73A-388F-43CE-9F36-281394BB26C8}';
  DIID_IDIEDownloaderEvents: TGUID = '{DB5D7BF0-C449-11E1-8021-CBE76088709B}';
  DIID_IProjectImporter: TGUID = '{6FAF9E2B-F38E-41D4-A25F-CE638DF04158}';
  DIID_IProcessingUtils: TGUID = '{0B0CBCBD-2362-4588-9788-625E45857CDE}';
  DIID_IWorkItemUnitPrice: TGUID = '{5F0992B6-DFB3-4F29-B950-4F63D65FB690}';
  DIID_IWorkItemPriceNumber: TGUID = '{F956992F-2583-45B5-9FC2-CA2569A2037A}';
  DIID_IWorkItem: TGUID = '{0D11D53C-27C1-4567-A373-02191DFCF743}';
  DIID_INoticeCollection: TGUID = '{10BCCE4C-30D8-42D5-AE4D-969B0EBDD906}';
  DIID_IDownloadCompletedEventArgs: TGUID = '{F5D3291A-6C6E-47E9-A260-8DE622B82EB6}';
  DIID_IDetectionEvents2: TGUID = '{3EC5F34C-F53A-4400-A841-801A08B5E71B}';
  DIID_IDealCollection: TGUID = '{2D034F70-09EC-4170-A442-1517F567FE66}';
  DIID_IDealUIOptions: TGUID = '{A1241010-3E0A-4098-AA81-E1C7C21708FC}';
  CLASS_DealUIOptions: TGUID = '{E06965AE-CB35-46C2-B323-97BBDA34550A}';
  DIID_IGroupItem: TGUID = '{AC628ACC-0789-49B3-BB25-2FA4B4365265}';
  DIID_IChapterItem: TGUID = '{47438293-365A-481A-A854-D8553E66CE25}';
  DIID_ILot: TGUID = '{A555300F-EEAB-4DD9-A8E3-69A24E9BD783}';
  DIID_IIntegerValue: TGUID = '{E476DB4D-4071-4413-AAF2-8E2E7576D677}';
  DIID_ICommentItem: TGUID = '{26BB641B-6203-4FE8-8C2C-A6E3FEE6ACCC}';
  DIID_IDIEResult: TGUID = '{250F3B59-E483-4BE4-A079-13A1B34B20F4}';
  DIID_IReport: TGUID = '{339A4FAE-B4BA-47BB-9086-C41F8C1753FE}';
  CLASS_Report: TGUID = '{A54216A2-729B-4F99-BFDD-96AA30B01427}';
  DIID_IRootItem: TGUID = '{1CAA25B7-DF28-49AD-90E5-D61EBF2F950F}';
  DIID_IWorkItemDesignation: TGUID = '{63E40920-070E-47C4-B280-164389E33345}';
  DIID_IMarking: TGUID = '{FE491689-DC81-4444-88FD-78C1A83E9251}';
  DIID_IBillOfQuantity: TGUID = '{B1475EC2-EF9D-46DB-8B99-C7E7CA25CCB6}';
  DIID_IAddress: TGUID = '{2317FFBB-45CD-4390-9B98-81B177F61623}';
  DIID_IDealResult: TGUID = '{73F675E7-DC1F-4F11-B909-C307F44AC420}';
  DIID_IDIEDownloader: TGUID = '{459C04A2-CC9B-4F20-81B0-1D549F33D0D9}';
  DIID_IProcedureType: TGUID = '{B37B9697-99CC-4806-BB2C-412FFDFEF735}';
  DIID_IDetectionEventsTester: TGUID = '{CFD41425-28E3-4960-8912-0C29BE7DF9BD}';
  CLASS_DetectionEventsTester: TGUID = '{8F4BFB9A-5275-488F-B855-792DC061E474}';
  DIID_ISpigaoCredentials: TGUID = '{59B6A2A2-0AB8-4601-9399-05B44AB299CD}';
  DIID_IDCEDocumentViewer: TGUID = '{58971E8C-A1CE-459C-B5C0-C0D20A189A6C}';
  DIID_IWorkItemSelection: TGUID = '{B2E87ED2-83E6-40A1-9146-5B3B2E7E69BE}';
  DIID_IWorkItemParent: TGUID = '{AA6B168A-BD59-4B00-B011-4B5DE635DEC9}';
  DIID_IPriceItem: TGUID = '{8771822E-7F90-4879-A1C9-B33E1EF935E8}';
  DIID_INumericValue: TGUID = '{79870506-7AA1-4D32-9E47-6FBAC903E791}';
  DIID_IMessage: TGUID = '{84CEF100-7B36-4A22-A833-0E568AEE7FA0}';
  DIID_ICompany: TGUID = '{6CEE5935-AAA3-442B-962B-59098CFECB8A}';
  CLASS_DIEDownloadEventProvider: TGUID = '{D29AE4A7-22B6-4DC4-B621-FD9F141FABE1}';
  DIID_ITender: TGUID = '{740564AB-49F0-433F-A412-A4E5993F907E}';
  DIID_IDIE: TGUID = '{2B888D4C-1248-4E8B-8AFC-A47D966DCC65}';
  DIID_IAccountData: TGUID = '{65E7B8CF-C5F2-4772-BE0F-65898FC4B5D6}';
  CLASS_AccountData: TGUID = '{56B7CE44-4EE5-435C-BF1A-68F47BB93326}';
  CLASS_ComponentVersion: TGUID = '{92D5357D-F831-417D-B035-DFE57F4C9C67}';
  DIID_IProcessingUI: TGUID = '{D4ED8C67-1284-40A6-B7D9-2405AA1CC744}';
  DIID_IBillOfQuantityEditor: TGUID = '{961A7B33-016E-452B-B3A4-014D03D2237D}';
  DIID_IProject: TGUID = '{9BF53378-3410-4317-9B95-3CDC86B64269}';
  DIID_ISpigaoMessage: TGUID = '{2BF4A787-2A10-427B-8DC3-8E861D2381FF}';
  DIID_IDIECollection: TGUID = '{B47EA317-25FC-4A31-A018-29BD4F823EEE}';
  DIID_IProcessingUIOptions: TGUID = '{95560364-14E3-4352-B1F7-9FFB1D9B38C0}';
  DIID_IImportResult: TGUID = '{A6B2DBA7-9E8F-4B28-AF64-6BCE9096B977}';
  DIID_IWorkItemUnit: TGUID = '{52BB6C76-9CD4-40EA-866D-A00A6A16F044}';
  DIID_IWorkItemLevel: TGUID = '{DAB6448C-FBCA-4EAC-A9B5-6D52A786C4D3}';
  DIID_IMessageCollection: TGUID = '{ABD88D6E-3E34-4C09-A222-E0BA1073530A}';
  DIID_IError: TGUID = '{CD2DA701-D5DC-490C-B84F-A4A5EB845371}';
  DIID_INotificationService: TGUID = '{69629DC5-4017-49FA-8688-30CCC92D9088}';
  DIID_IGroupItemValue: TGUID = '{56F79E31-D5CF-4EF3-A47B-2962228DB16D}';
  DIID_IBillOfQuantityProperties: TGUID = '{03ABDABB-CABB-42A9-B32F-81AB06232127}';
  DIID_IDetectionUI: TGUID = '{428CA8F9-BB3D-40C5-8C3F-F3C8BF25B5BE}';
  DIID_IDealSummary: TGUID = '{398A1177-E48E-4E8A-A9A9-940D17E450B6}';
  DIID_IPropertyCollection: TGUID = '{10A300BF-44D5-48B1-B62B-15534C04B4A2}';
  DIID_INotice: TGUID = '{AAB12904-3549-494B-A30A-C6F5CA47F9AC}';
  DIID_IStringCollection: TGUID = '{50F9DD60-62B6-40E4-96EB-FEFAB7A9BC9F}';
  DIID_ISpigao: TGUID = '{A6935D58-5651-439D-8DE7-5DC5441136CE}';
  DIID_ISpigaoMessageCollection: TGUID = '{36BC6A5F-740D-4DFE-B2C0-A155648145FA}';
  DIID_ISpigaoWebSite: TGUID = '{4C5961E0-25CF-48E9-A645-74F7D67762AD}';
  DIID_IDCEFile: TGUID = '{81CBD17F-F5BF-415D-A0EF-E3CCA3D72BD4}';
  DIID_IActivityCollection: TGUID = '{E69C079D-C5CE-485C-B57C-63751B8C5FB6}';
  DIID_ITrialAccountState: TGUID = '{8B1A2847-80CF-4EA3-9B1A-6FBA01309944}';
  DIID_IWorkItemTypeValue: TGUID = '{45821B51-333D-4E04-A1CD-131B3739A2CA}';
  DIID_IStringValue: TGUID = '{75C64598-7052-4731-BD20-40A8297D40E4}';
  CLASS_DetectionEventProvider: TGUID = '{2FC05B1B-D0DF-4FEC-9CD7-10A89127FA72}';
  DIID_IProjectExporter: TGUID = '{FEEE5573-72CF-445F-8F5F-A324511CD481}';
  DIID_IProcessingEngine: TGUID = '{D73ECB5A-E5CA-46E3-B164-84C50D7EFF52}';
  DIID_IContract: TGUID = '{04ED3DFF-6A97-41C9-A94A-81336769527F}';
  DIID_ITaskResult: TGUID = '{5D675097-FF60-43A7-BF62-B0D047A4D713}';
  DIID_IDealUI: TGUID = '{EF506804-2AEE-42B0-A6BE-03D192E57D34}';
  DIID_ISpigaoAlert: TGUID = '{FFC4247D-2706-4850-9D13-D641019899A8}';
  DIID_IDealSummaryCollection: TGUID = '{754E11E6-A29D-44B8-9A8A-A8C669F30289}';
  DIID_IDealUpdateState: TGUID = '{A2A121B0-2527-46C1-870D-3B2298D58753}';
  CLASS_Error: TGUID = '{F3C52C7E-8172-49EB-BAEE-BB235E5BDF16}';
  DIID_IDownloadProgressChangedEventArgs: TGUID = '{015A840E-CFBA-42D8-8F2F-72F13DC7D3AE}';
  DIID_IDetectionAppConfig: TGUID = '{CABEEE57-810B-4256-8C99-A7674BEE996D}';
  DIID_IDeal: TGUID = '{AF76D763-D1CC-4CBC-BA4F-B888BA09F031}';
  DIID_IAccount: TGUID = '{5CC64EE1-9C83-48CF-A772-E07C0DFB948A}';
  CLASS_SpigaoLoader: TGUID = '{4E1073C9-FDC1-4A32-A765-7E878FB8EBF3}';
  DIID_IWorkItemQuantity: TGUID = '{BE8D140D-F90A-41C5-B82F-DE7C6C1563DD}';
  DIID_IWorkItemAmount: TGUID = '{8184A93D-7E57-490E-B08B-2ABE17C6551D}';
  DIID_IProperty: TGUID = '{5640CC99-97AB-4551-87A5-AC215CC01F9B}';
  DIID_IAmountItem: TGUID = '{0017D291-9DBF-4467-A72F-5EF69F9EC9C7}';
  DIID_IReportMessage: TGUID = '{45289823-532F-4367-BBC3-6F035FFA7329}';
  CLASS_TaskResult: TGUID = '{AB47F3DF-EF1C-48AB-B839-F477A6F25EFD}';
  DIID_ILotCollection: TGUID = '{B6D39D68-0C8B-4856-8067-7AA77DCEC95C}';
  DIID_IResultEventArgs: TGUID = '{600FF0C6-00CF-4EB2-8F5C-84705A58AEA9}';
  DIID_IMarketType: TGUID = '{D7098A5B-5369-4960-BB68-9F48130319EF}';
  CLASS_ReportMessage: TGUID = '{B85340BB-0639-4C30-9D0D-0C41401167B5}';
  CLASS_StringCollection: TGUID = '{DBDB82D6-1C79-4FD4-99A1-DD8FE011D1E3}';
  DIID_IFormula: TGUID = '{646286ED-83E9-423E-9E7C-0FA3B0429B56}';
  DIID_IDetectionUserConfig: TGUID = '{8915E454-FBA0-45B9-8B29-4C8C398288F2}';
  DIID_ITerritoryCollection: TGUID = '{7191CF2B-C3EE-4222-B481-B2F196996F5B}';
  DIID_IActivity: TGUID = '{D1DDE808-7F5E-4EDE-A26F-195DD058524F}';
  DIID_ISearchResult: TGUID = '{DC15D625-3BE9-41A7-B793-1F9CA2BBD64E}';
  DIID_IDCEFileCollection: TGUID = '{3B3F3331-0727-4484-A446-CF79D901C836}';
  DIID_IPDFGeneration: TGUID = '{B367C992-C24E-45D3-A525-1CC00036BE82}';
  DIID_IProcessingAppConfig: TGUID = '{2BDB3EE3-0401-4300-B7D2-169FD9E819C9}';
  DIID_IWorkItemCollection: TGUID = '{1B10FD5F-83C8-44F3-AE09-B25E870F6E2D}';
  DIID_ITerritory: TGUID = '{AE7A746E-33F3-4A35-AD3F-5D01AE04A11C}';
  DIID_INotificationUI: TGUID = '{BD40D123-EB68-45D3-91EB-25CF6E02F4B0}';
  DIID_IAccountUI: TGUID = '{8A91ED49-620C-4314-B5C4-E9DB87EBA1FF}';
  DIID_IDealService: TGUID = '{1A83ECE4-146F-4927-BEAD-5273C9971CAA}';

// *********************************************************************//
// Déclaration d'énumérations définies dans la bibliothèque de types    
// *********************************************************************//
// Constantes pour enum AccountType
type
  AccountType = TOleEnum;
const
  AccountType_Undefined = $00000000;
  AccountType_Starter = $00000001;
  AccountType_Trial = $00000002;
  AccountType_Full = $00000003;

// Constantes pour enum AccountStatus
type
  AccountStatus = TOleEnum;
const
  AccountStatus_NotCreated = $00000000;
  AccountStatus_Active = $00000001;
  AccountStatus_Disabled = $00000002;

// Constantes pour enum DealUpdatedData
type
  DealUpdatedData = TOleEnum;
const
  DealUpdatedData_NewCorrectionNotice = $00000001;
  DealUpdatedData_DIEPublished = $00000002;
  DealUpdatedData_DIEUpdated = $00000004;

// Constantes pour enum ComDocumentType
type
  ComDocumentType = TOleEnum;
const
  ComDocumentType_Unknown = $00000000;
  ComDocumentType_DET = $00000001;
  ComDocumentType_BPX = $00000002;
  ComDocumentType_AE = $00000003;

// Constantes pour enum AlertFrequency
type
  AlertFrequency = TOleEnum;
const
  AlertFrequency_Min30 = $0000001E;
  AlertFrequency_Min60 = $0000003C;
  AlertFrequency_Min120 = $00000078;
  AlertFrequency_Min240 = $000000F0;
  AlertFrequency_Min1440 = $000005A0;

// Constantes pour enum PriceNumberRenumberingScope
type
  PriceNumberRenumberingScope = TOleEnum;
const
  PriceNumberRenumberingScope_Global = $00000000;
  PriceNumberRenumberingScope_WithinChapter = $00000001;

// Constantes pour enum ImportMode
type
  ImportMode = TOleEnum;
const
  ImportMode_Allow = $00000000;
  ImportMode_Exclude = $00000001;
  ImportMode_Change = $00000002;

// Constantes pour enum BillOfQuantityEditCommand
type
  BillOfQuantityEditCommand = TOleEnum;
const
  BillOfQuantityEditCommand_LevelUp = $00000000;
  BillOfQuantityEditCommand_LevelDown = $00000001;
  BillOfQuantityEditCommand_ChangeToComment = $00000002;
  BillOfQuantityEditCommand_ChangeToChapter = $00000003;
  BillOfQuantityEditCommand_Undo = $00000004;
  BillOfQuantityEditCommand_Redo = $00000005;
  BillOfQuantityEditCommand_DiscardAll = $00000006;
  BillOfQuantityEditCommand_ActivateLinks = $00000007;
  BillOfQuantityEditCommand_DeactivateLinks = $00000008;

// Constantes pour enum ComChapterType
type
  ComChapterType = TOleEnum;
const
  ComChapterType_Standard = $00000000;
  ComChapterType_Alternative = $00000001;
  ComChapterType_Option = $00000002;
  ComChapterType_Summary = $00000003;

// Constantes pour enum ComAmountType
type
  ComAmountType = TOleEnum;
const
  ComAmountType_Unknown = $00000000;
  ComAmountType_ExcludingTaxes = $00000001;
  ComAmountType_VAT = $00000002;
  ComAmountType_IncludingTaxes = $00000003;
  ComAmountType_SubTotal = $00000004;
  ComAmountType_Summary = $00000005;

// Constantes pour enum ExecResult
type
  ExecResult = TOleEnum;
const
  ExecResult_Completed = $00000000;
  ExecResult_Cancelled = $FFFFFFFF;
  ExecResult_Failed = $FFFFFFFE;

// Constantes pour enum DIEStatus
type
  DIEStatus = TOleEnum;
const
  DIEStatus_NotProduced = $00000000;
  DIEStatus_InProduction = $00000001;
  DIEStatus_Available = $00000002;
  DIEStatus_Unavailable = $00000003;

// Constantes pour enum MaxLevelAction
type
  MaxLevelAction = TOleEnum;
const
  MaxLevelAction_Top = $00000000;
  MaxLevelAction_Bottom = $00000001;

// Constantes pour enum FileFormat
type
  FileFormat = TOleEnum;
const
  FileFormat_Txt = $00000000;
  FileFormat_Xml = $00000001;

// Constantes pour enum ComPriceType
type
  ComPriceType = TOleEnum;
const
  ComPriceType_Standard = $00000000;
  ComPriceType_ForMemory = $00000001;
  ComPriceType_Fixed = $00000002;

// Constantes pour enum DCEFileType
type
  DCEFileType = TOleEnum;
const
  DCEFileType_Avis = $00000001;
  DCEFileType_DET = $00000002;
  DCEFileType_BPU = $00000004;
  DCEFileType_RC = $00000008;
  DCEFileType_CC = $00000010;
  DCEFileType_AvisRectificatif = $00000020;

// Constantes pour enum ProxyMethod
type
  ProxyMethod = TOleEnum;
const
  ProxyMethod_AutoDetect = $00000000;
  ProxyMethod_PACFile = $00000001;
  ProxyMethod_Manual = $00000002;

// Constantes pour enum ReportSeverity
type
  ReportSeverity = TOleEnum;
const
  ReportSeverity_Info = $00000000;
  ReportSeverity_Warning = $00000001;
  ReportSeverity_Error = $00000002;

// Constantes pour enum AllowLinksMode
type
  AllowLinksMode = TOleEnum;
const
  AllowLinksMode_ByRef = $00000000;
  AllowLinksMode_ByPriceNumber = $00000001;
  AllowLinksMode_None = $00000002;

// Constantes pour enum ComMessageCategory
type
  ComMessageCategory = TOleEnum;
const
  ComMessageCategory_Information = $00000000;
  ComMessageCategory_Warning = $00000001;
  ComMessageCategory_Error = $00000002;

// Constantes pour enum ExecutionMode
type
  ExecutionMode = TOleEnum;
const
  ExecutionMode_Trial = $00000000;
  ExecutionMode_Full = $00000001;

// Constantes pour enum DealStatus
type
  DealStatus = TOleEnum;
const
  DealStatus_New = $00000000;
  DealStatus_Selected = $00000001;
  DealStatus_Integrated = $00000002;
  DealStatus_Deleted = $00000003;

// Constantes pour enum AlertType
type
  AlertType = TOleEnum;
const
  AlertType_NewDeals = $00000000;
  AlertType_UpdatedDeals = $00000001;
  AlertType_DealsToLoad = $00000002;
  AlertType_SpigaoMessage = $00000003;
  AlertType_UpdateAvailable = $00000004;
  AlertType_Error = $00000005;

// Constantes pour enum ComWorkItemType
type
  ComWorkItemType = TOleEnum;
const
  ComWorkItemType_Unknown = $00000000;
  ComWorkItemType_Title = $00000001;
  ComWorkItemType_Chapter = $00000002;
  ComWorkItemType_Price = $00000003;
  ComWorkItemType_Amount = $00000004;
  ComWorkItemType_Comment = $00000005;

// Constantes pour enum ComWorkItemFilter
type
  ComWorkItemFilter = TOleEnum;
const
  ComWorkItemFilter_All = $00000000;
  ComWorkItemFilter_Title = $00000001;
  ComWorkItemFilter_ChapterStandard = $00000002;
  ComWorkItemFilter_ChapterAlternative = $00000004;
  ComWorkItemFilter_ChapterOption = $00000008;
  ComWorkItemFilter_ChapterSummary = $00000010;
  ComWorkItemFilter_ChapterAll = $0000001E;
  ComWorkItemFilter_PriceStandard = $00000020;
  ComWorkItemFilter_PriceForMemory = $00000040;
  ComWorkItemFilter_PriceFixed = $00000080;
  ComWorkItemFilter_PriceAll = $000000E0;
  ComWorkItemFilter_AmountExcludingTaxes = $00000100;
  ComWorkItemFilter_AmountVAT = $00000200;
  ComWorkItemFilter_AmountIncludingTaxes = $00000400;
  ComWorkItemFilter_AmountSubTotal = $00000800;
  ComWorkItemFilter_AmountSummary = $00001000;
  ComWorkItemFilter_AmountAll = $00001F00;
  ComWorkItemFilter_Comment = $00002000;

// Constantes pour enum TerritoryType
type
  TerritoryType = TOleEnum;
const
  TerritoryType_Departement = $00000000;
  TerritoryType_Township = $00000001;

// Constantes pour enum RestrictionType
type
  RestrictionType = TOleEnum;
const
  RestrictionType_No = $00000000;
  RestrictionType_Yes = $00000001;
  RestrictionType_Unknown = $00000002;

type

// *********************************************************************//
// Déclaration Forward des types définis dans la bibliothèque de types    
// *********************************************************************//
  IDocumentViewer = dispinterface;
  IReportMessageCollection = dispinterface;
  IProcessingUserConfig = dispinterface;
  IWorkItemType = dispinterface;
  IVersion = dispinterface;
  IProxySettings = dispinterface;
  IDetectionEngine = dispinterface;
  IDetectionEvents = dispinterface;
  IDetectionEventProvider = dispinterface;
  IDetectionAlerts = dispinterface;
  IDIEDownloadEventProvider = dispinterface;
  IDIEDownloaderEvents = dispinterface;
  IProjectImporter = dispinterface;
  IProcessingUtils = dispinterface;
  IWorkItemUnitPrice = dispinterface;
  IWorkItemPriceNumber = dispinterface;
  IWorkItem = dispinterface;
  INoticeCollection = dispinterface;
  IDownloadCompletedEventArgs = dispinterface;
  IDetectionEvents2 = dispinterface;
  IDealCollection = dispinterface;
  IDealUIOptions = dispinterface;
  IGroupItem = dispinterface;
  IChapterItem = dispinterface;
  ILot = dispinterface;
  IIntegerValue = dispinterface;
  ICommentItem = dispinterface;
  IDIEResult = dispinterface;
  IReport = dispinterface;
  IRootItem = dispinterface;
  IWorkItemDesignation = dispinterface;
  IMarking = dispinterface;
  IBillOfQuantity = dispinterface;
  IAddress = dispinterface;
  IDealResult = dispinterface;
  IDIEDownloader = dispinterface;
  IProcedureType = dispinterface;
  IDetectionEventsTester = dispinterface;
  ISpigaoCredentials = dispinterface;
  IDCEDocumentViewer = dispinterface;
  IWorkItemSelection = dispinterface;
  IWorkItemParent = dispinterface;
  IPriceItem = dispinterface;
  INumericValue = dispinterface;
  IMessage = dispinterface;
  ICompany = dispinterface;
  ITender = dispinterface;
  IDIE = dispinterface;
  IAccountData = dispinterface;
  IProcessingUI = dispinterface;
  IBillOfQuantityEditor = dispinterface;
  IProject = dispinterface;
  ISpigaoMessage = dispinterface;
  IDIECollection = dispinterface;
  IProcessingUIOptions = dispinterface;
  IImportResult = dispinterface;
  IWorkItemUnit = dispinterface;
  IWorkItemLevel = dispinterface;
  IMessageCollection = dispinterface;
  IError = dispinterface;
  INotificationService = dispinterface;
  IGroupItemValue = dispinterface;
  IBillOfQuantityProperties = dispinterface;
  IDetectionUI = dispinterface;
  IDealSummary = dispinterface;
  IPropertyCollection = dispinterface;
  INotice = dispinterface;
  IStringCollection = dispinterface;
  ISpigao = dispinterface;
  ISpigaoMessageCollection = dispinterface;
  ISpigaoWebSite = dispinterface;
  IDCEFile = dispinterface;
  IActivityCollection = dispinterface;
  ITrialAccountState = dispinterface;
  IWorkItemTypeValue = dispinterface;
  IStringValue = dispinterface;
  IProjectExporter = dispinterface;
  IProcessingEngine = dispinterface;
  IContract = dispinterface;
  ITaskResult = dispinterface;
  IDealUI = dispinterface;
  ISpigaoAlert = dispinterface;
  IDealSummaryCollection = dispinterface;
  IDealUpdateState = dispinterface;
  IDownloadProgressChangedEventArgs = dispinterface;
  IDetectionAppConfig = dispinterface;
  IDeal = dispinterface;
  IAccount = dispinterface;
  IWorkItemQuantity = dispinterface;
  IWorkItemAmount = dispinterface;
  IProperty = dispinterface;
  IAmountItem = dispinterface;
  IReportMessage = dispinterface;
  ILotCollection = dispinterface;
  IResultEventArgs = dispinterface;
  IMarketType = dispinterface;
  IFormula = dispinterface;
  IDetectionUserConfig = dispinterface;
  ITerritoryCollection = dispinterface;
  IActivity = dispinterface;
  ISearchResult = dispinterface;
  IDCEFileCollection = dispinterface;
  IPDFGeneration = dispinterface;
  IProcessingAppConfig = dispinterface;
  IWorkItemCollection = dispinterface;
  ITerritory = dispinterface;
  INotificationUI = dispinterface;
  IAccountUI = dispinterface;
  IDealService = dispinterface;

// *********************************************************************//
// Déclaration de CoClasses définies dans la bibliothèque de types 
// (REMARQUE: On affecte chaque CoClasse à son Interface par défaut)              
// *********************************************************************//
  ReportMessageCollection = IReportMessageCollection;
  DealUIOptions = IDealUIOptions;
  Report = IReport;
  DetectionEventsTester = IDetectionEventsTester;
  DIEDownloadEventProvider = IDIEDownloadEventProvider;
  AccountData = IAccountData;
  ComponentVersion = IVersion;
  DetectionEventProvider = IDetectionEventProvider;
  Error = IError;
  SpigaoLoader = ISpigao;
  TaskResult = ITaskResult;
  ReportMessage = IReportMessage;
  StringCollection = IStringCollection;


// *********************************************************************//
// DispIntf :  IDocumentViewer
// Flags :     (4096) Dispatchable
// GUID :      {15A10AB3-BA7D-46C7-BD9E-37C1CDB1D643}
// *********************************************************************//
  IDocumentViewer = dispinterface
    ['{15A10AB3-BA7D-46C7-BD9E-37C1CDB1D643}']
    procedure Show(const project: IProject); dispid 128;
    function NavigateTo(const project: IProject; const workItemID: WideString; 
                        documentType: ComDocumentType): WordBool; dispid 129;
  end;

// *********************************************************************//
// DispIntf :  IReportMessageCollection
// Flags :     (4096) Dispatchable
// GUID :      {B34023FC-D78F-4C2E-8D0C-B6B8E8D6181B}
// *********************************************************************//
  IReportMessageCollection = dispinterface
    ['{B34023FC-D78F-4C2E-8D0C-B6B8E8D6181B}']
    property Count: Integer readonly dispid 1;
    property Item[index: Integer]: IReportMessage readonly dispid 2;
  end;

// *********************************************************************//
// DispIntf :  IProcessingUserConfig
// Flags :     (4096) Dispatchable
// GUID :      {355FDA9D-B6EF-4FF5-B9BD-22E5753FC18E}
// *********************************************************************//
  IProcessingUserConfig = dispinterface
    ['{355FDA9D-B6EF-4FF5-B9BD-22E5753FC18E}']
    property DefaultStorageFolder: WideString dispid 1;
    property DoMergePriceNumberOnChange: WordBool dispid 2;
    property DoActivateLinkedWorkitemsByDefault: WordBool dispid 3;
    procedure Save; dispid 128;
  end;

// *********************************************************************//
// DispIntf :  IWorkItemType
// Flags :     (4096) Dispatchable
// GUID :      {C9A834F0-A70C-4A4C-A21E-9EACA9C3A9B6}
// *********************************************************************//
  IWorkItemType = dispinterface
    ['{C9A834F0-A70C-4A4C-A21E-9EACA9C3A9B6}']
    property Current: IWorkItemTypeValue readonly dispid 1;
    property Original: IWorkItemTypeValue readonly dispid 2;
    property IsModified: WordBool readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf :  IVersion
// Flags :     (4096) Dispatchable
// GUID :      {77C83E13-6254-472C-A328-B7C1DE04D17F}
// *********************************************************************//
  IVersion = dispinterface
    ['{77C83E13-6254-472C-A328-B7C1DE04D17F}']
    property Major: Integer readonly dispid 1;
    property Minor: Integer readonly dispid 2;
    property Build: Integer readonly dispid 3;
    property Text: WideString readonly dispid 4;
  end;

// *********************************************************************//
// DispIntf :  IProxySettings
// Flags :     (4096) Dispatchable
// GUID :      {F0961E06-04BB-4796-9EEF-A71ABA2B6696}
// *********************************************************************//
  IProxySettings = dispinterface
    ['{F0961E06-04BB-4796-9EEF-A71ABA2B6696}']
    property ConfigType: ProxyMethod dispid 1;
    property PACFileAddress: WideString dispid 2;
    property ProxyAddress: WideString dispid 3;
    property ProxyPort: Integer dispid 4;
    property ProxyLogin: WideString dispid 5;
    property ProxyPassword: WideString dispid 6;
    procedure Save; dispid 128;
  end;

// *********************************************************************//
// DispIntf :  IDetectionEngine
// Flags :     (4096) Dispatchable
// GUID :      {1E9D57E3-67C8-4A57-8F93-BC05A761CBA9}
// *********************************************************************//
  IDetectionEngine = dispinterface
    ['{1E9D57E3-67C8-4A57-8F93-BC05A761CBA9}']
    property GlobalSettings: IDetectionAppConfig readonly dispid 1;
    property UserSettings: IDetectionUserConfig readonly dispid 2;
    property Alerts: IDetectionAlerts readonly dispid 3;
    property UI: IDetectionUI readonly dispid 4;
    property WebSite: ISpigaoWebSite readonly dispid 5;
    property DisplayAlerts: WordBool dispid 6;
    property Notifications: INotificationService readonly dispid 7;
    property Deals: IDealService readonly dispid 8;
    property Events: IDetectionEventProvider readonly dispid 9;
    property Account: IAccount readonly dispid 10;
    function TestConnection: WordBool; dispid 128;
    function SearchNewDeals: ISearchResult; dispid 129;
    function SearchUpdatedDeals: ISearchResult; dispid 130;
    function ListDealsToLoad: IDealSummaryCollection; dispid 131;
    function ListSpigaoMessages: ISpigaoMessageCollection; dispid 132;
    function LoadDeal(dealId: Integer): IDeal; dispid 133;
    function SetIntegrationStatus(dealId: Integer; value: WordBool): WordBool; dispid 134;
    function DeleteAlertOnDeal(dealId: Integer): WordBool; dispid 135;
    function GetIntegratedDealModifiedId: {??PSafeArray}OleVariant; dispid 136;
    function GetLastError: IError; dispid 137;
  end;

// *********************************************************************//
// DispIntf :  IDetectionEvents
// Flags :     (4096) Dispatchable
// GUID :      {23A7D2C4-050F-4B90-80B0-F5DD7119AD0C}
// *********************************************************************//
  IDetectionEvents = dispinterface
    ['{23A7D2C4-050F-4B90-80B0-F5DD7119AD0C}']
    procedure OnNewDeals(const rea: IResultEventArgs); dispid 1;
    procedure OnUpdatedDeals(const rea: IResultEventArgs); dispid 2;
    procedure OnDealsToLoad(const rea: IResultEventArgs); dispid 3;
    procedure OnSpigaoMessages(const rea: IResultEventArgs); dispid 4;
    procedure OnUpdateAvailable(const rea: IResultEventArgs); dispid 5;
    procedure OnError(const rea: IResultEventArgs); dispid 6;
  end;

// *********************************************************************//
// DispIntf :  IDetectionEventProvider
// Flags :     (4096) Dispatchable
// GUID :      {8BA7B2B0-B8E0-473F-B1C6-2409A83A2917}
// *********************************************************************//
  IDetectionEventProvider = dispinterface
    ['{8BA7B2B0-B8E0-473F-B1C6-2409A83A2917}']
    property Test: IDetectionEventsTester readonly dispid 1;
  end;

// *********************************************************************//
// DispIntf :  IDetectionAlerts
// Flags :     (4096) Dispatchable
// GUID :      {7BA54285-A051-4ACF-B748-61172525CDAD}
// *********************************************************************//
  IDetectionAlerts = dispinterface
    ['{7BA54285-A051-4ACF-B748-61172525CDAD}']
    property DisplayAlerts: WordBool dispid 1;
    procedure Configure(AlertFrequency: {??TimeSpan}OleVariant); dispid 128;
    procedure Configure_2(AlertFrequency: Integer); dispid 129;
  end;

// *********************************************************************//
// DispIntf :  IDIEDownloadEventProvider
// Flags :     (4096) Dispatchable
// GUID :      {2C67F73A-388F-43CE-9F36-281394BB26C8}
// *********************************************************************//
  IDIEDownloadEventProvider = dispinterface
    ['{2C67F73A-388F-43CE-9F36-281394BB26C8}']
  end;

// *********************************************************************//
// DispIntf :  IDIEDownloaderEvents
// Flags :     (4096) Dispatchable
// GUID :      {DB5D7BF0-C449-11E1-8021-CBE76088709B}
// *********************************************************************//
  IDIEDownloaderEvents = dispinterface
    ['{DB5D7BF0-C449-11E1-8021-CBE76088709B}']
    procedure DownloadProgressChanged(const sender: IDIEDownloader; 
                                      const e: IDownloadProgressChangedEventArgs); dispid 128;
    procedure DownloadCompleted(const sender: IDIEDownloader; const e: IDownloadCompletedEventArgs); dispid 129;
  end;

// *********************************************************************//
// DispIntf :  IProjectImporter
// Flags :     (4096) Dispatchable
// GUID :      {6FAF9E2B-F38E-41D4-A25F-CE638DF04158}
// *********************************************************************//
  IProjectImporter = dispinterface
    ['{6FAF9E2B-F38E-41D4-A25F-CE638DF04158}']
    property Report: IReport readonly dispid 2;
    procedure Format; dispid 128;
    function Edit: IBillOfQuantityEditor; dispid 129;
    procedure Done; dispid 130;
    function SaveToTxt: WideString; dispid 131;
    function SaveToXml: WideString; dispid 132;
  end;

// *********************************************************************//
// DispIntf :  IProcessingUtils
// Flags :     (4096) Dispatchable
// GUID :      {0B0CBCBD-2362-4588-9788-625E45857CDE}
// *********************************************************************//
  IProcessingUtils = dispinterface
    ['{0B0CBCBD-2362-4588-9788-625E45857CDE}']
    property NullProject: IProject readonly dispid 1;
  end;

// *********************************************************************//
// DispIntf :  IWorkItemUnitPrice
// Flags :     (4096) Dispatchable
// GUID :      {5F0992B6-DFB3-4F29-B950-4F63D65FB690}
// *********************************************************************//
  IWorkItemUnitPrice = dispinterface
    ['{5F0992B6-DFB3-4F29-B950-4F63D65FB690}']
    property Current: INumericValue readonly dispid 1;
    property Original: INumericValue readonly dispid 2;
    property IsModified: WordBool readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf :  IWorkItemPriceNumber
// Flags :     (4096) Dispatchable
// GUID :      {F956992F-2583-45B5-9FC2-CA2569A2037A}
// *********************************************************************//
  IWorkItemPriceNumber = dispinterface
    ['{F956992F-2583-45B5-9FC2-CA2569A2037A}']
    property Current: IStringValue readonly dispid 1;
    property Original: IStringValue readonly dispid 2;
    property IsModified: WordBool readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf :  IWorkItem
// Flags :     (4096) Dispatchable
// GUID :      {0D11D53C-27C1-4567-A373-02191DFCF743}
// *********************************************************************//
  IWorkItem = dispinterface
    ['{0D11D53C-27C1-4567-A373-02191DFCF743}']
    property Id: WideString readonly dispid 1;
    property PriceNumber: IWorkItemPriceNumber readonly dispid 2;
    property Designation: IWorkItemDesignation readonly dispid 3;
    property Parent: IWorkItemParent readonly dispid 4;
    property Level: IWorkItemLevel readonly dispid 5;
    property ItemType: IWorkItemType readonly dispid 6;
    property AsChapter: IChapterItem readonly dispid 7;
    property AsPrice: IPriceItem readonly dispid 8;
    property AsAmount: IAmountItem readonly dispid 9;
  end;

// *********************************************************************//
// DispIntf :  INoticeCollection
// Flags :     (4096) Dispatchable
// GUID :      {10BCCE4C-30D8-42D5-AE4D-969B0EBDD906}
// *********************************************************************//
  INoticeCollection = dispinterface
    ['{10BCCE4C-30D8-42D5-AE4D-969B0EBDD906}']
    property Item[index: Integer]: INotice readonly dispid 1;
  end;

// *********************************************************************//
// DispIntf :  IDownloadCompletedEventArgs
// Flags :     (4096) Dispatchable
// GUID :      {F5D3291A-6C6E-47E9-A260-8DE622B82EB6}
// *********************************************************************//
  IDownloadCompletedEventArgs = dispinterface
    ['{F5D3291A-6C6E-47E9-A260-8DE622B82EB6}']
    property ProjectPath: WideString readonly dispid 1;
    property Status: ExecResult readonly dispid 2;
  end;

// *********************************************************************//
// DispIntf :  IDetectionEvents2
// Flags :     (4096) Dispatchable
// GUID :      {3EC5F34C-F53A-4400-A841-801A08B5E71B}
// *********************************************************************//
  IDetectionEvents2 = dispinterface
    ['{3EC5F34C-F53A-4400-A841-801A08B5E71B}']
    procedure EngineReady; dispid 128;
    procedure AccountDataRequired(AccountType: AccountType; var accountKey: WideString; 
                                  const AccountData: IAccountData); dispid 129;
    procedure AccountStateChanged; dispid 130;
    procedure UserRequestNewDeals; dispid 131;
    procedure UserRequestSelectedDeals; dispid 132;
    procedure UserRequestUpdatedDeals; dispid 133;
    procedure UserRequestDealsWithAvailableDIE; dispid 134;
    procedure UserValidationRequired(var password: WideString; var isValidated: WordBool); dispid 135;
  end;

// *********************************************************************//
// DispIntf :  IDealCollection
// Flags :     (4096) Dispatchable
// GUID :      {2D034F70-09EC-4170-A442-1517F567FE66}
// *********************************************************************//
  IDealCollection = dispinterface
    ['{2D034F70-09EC-4170-A442-1517F567FE66}']
    property FullCount: Integer readonly dispid 1;
    property Item[index: Integer]: IDeal readonly dispid 2;
  end;

// *********************************************************************//
// DispIntf :  IDealUIOptions
// Flags :     (4096) Dispatchable
// GUID :      {A1241010-3E0A-4098-AA81-E1C7C21708FC}
// *********************************************************************//
  IDealUIOptions = dispinterface
    ['{A1241010-3E0A-4098-AA81-E1C7C21708FC}']
    property CanSelectDeal: WordBool dispid 1;
    property CanImportDeal: WordBool dispid 2;
    property CanDeleteDeal: WordBool dispid 3;
    property CanMarkUpdatedDealAsRead: WordBool dispid 4;
  end;

// *********************************************************************//
// DispIntf :  IGroupItem
// Flags :     (4096) Dispatchable
// GUID :      {AC628ACC-0789-49B3-BB25-2FA4B4365265}
// *********************************************************************//
  IGroupItem = dispinterface
    ['{AC628ACC-0789-49B3-BB25-2FA4B4365265}']
    property Id: WideString readonly dispid 1;
    property PriceNumber: IWorkItemPriceNumber readonly dispid 2;
    property Designation: IWorkItemDesignation readonly dispid 3;
    property Parent: IWorkItemParent readonly dispid 4;
    property Level: IWorkItemLevel readonly dispid 5;
    property ItemType: IWorkItemType readonly dispid 6;
    property WorkItems: IWorkItemCollection readonly dispid 7;
    function GetFlatWorkItemList(filter: ComWorkItemFilter): IWorkItemCollection; dispid 128;
  end;

// *********************************************************************//
// DispIntf :  IChapterItem
// Flags :     (4096) Dispatchable
// GUID :      {47438293-365A-481A-A854-D8553E66CE25}
// *********************************************************************//
  IChapterItem = dispinterface
    ['{47438293-365A-481A-A854-D8553E66CE25}']
    property Id: WideString readonly dispid 1;
    property PriceNumber: IWorkItemPriceNumber readonly dispid 2;
    property Designation: IWorkItemDesignation readonly dispid 3;
    property Parent: IWorkItemParent readonly dispid 4;
    property Level: IWorkItemLevel readonly dispid 5;
    property ItemType: IWorkItemType readonly dispid 6;
    property WorkItems: IWorkItemCollection readonly dispid 7;
    function GetFlatWorkItemList(filter: ComWorkItemFilter): IWorkItemCollection; dispid 128;
    property ItemSubType: ComChapterType readonly dispid 9;
  end;

// *********************************************************************//
// DispIntf :  ILot
// Flags :     (4096) Dispatchable
// GUID :      {A555300F-EEAB-4DD9-A8E3-69A24E9BD783}
// *********************************************************************//
  ILot = dispinterface
    ['{A555300F-EEAB-4DD9-A8E3-69A24E9BD783}']
    property Reference: WideString readonly dispid 1;
    property Subject: WideString readonly dispid 2;
    property Caption: WideString readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf :  IIntegerValue
// Flags :     (4096) Dispatchable
// GUID :      {E476DB4D-4071-4413-AAF2-8E2E7576D677}
// *********************************************************************//
  IIntegerValue = dispinterface
    ['{E476DB4D-4071-4413-AAF2-8E2E7576D677}']
    property value: Integer readonly dispid 1;
  end;

// *********************************************************************//
// DispIntf :  ICommentItem
// Flags :     (4096) Dispatchable
// GUID :      {26BB641B-6203-4FE8-8C2C-A6E3FEE6ACCC}
// *********************************************************************//
  ICommentItem = dispinterface
    ['{26BB641B-6203-4FE8-8C2C-A6E3FEE6ACCC}']
    property Id: WideString readonly dispid 1;
    property PriceNumber: IWorkItemPriceNumber readonly dispid 2;
    property Designation: IWorkItemDesignation readonly dispid 3;
    property Parent: IWorkItemParent readonly dispid 4;
    property Level: IWorkItemLevel readonly dispid 5;
    property ItemType: IWorkItemType readonly dispid 6;
  end;

// *********************************************************************//
// DispIntf :  IDIEResult
// Flags :     (4096) Dispatchable
// GUID :      {250F3B59-E483-4BE4-A079-13A1B34B20F4}
// *********************************************************************//
  IDIEResult = dispinterface
    ['{250F3B59-E483-4BE4-A079-13A1B34B20F4}']
    property Result: ExecResult readonly dispid 1;
    property SelectedDIE: IDIE readonly dispid 2;
  end;

// *********************************************************************//
// DispIntf :  IReport
// Flags :     (4096) Dispatchable
// GUID :      {339A4FAE-B4BA-47BB-9086-C41F8C1753FE}
// *********************************************************************//
  IReport = dispinterface
    ['{339A4FAE-B4BA-47BB-9086-C41F8C1753FE}']
    property Title: WideString readonly dispid 1;
    property Infos: IReportMessageCollection readonly dispid 2;
    property Warnings: IReportMessageCollection readonly dispid 3;
    property Errors: IReportMessageCollection readonly dispid 4;
    function ToHtmlString: WideString; dispid 128;
    procedure ToHtmlStream(const output: _Stream); dispid 129;
  end;

// *********************************************************************//
// DispIntf :  IRootItem
// Flags :     (4096) Dispatchable
// GUID :      {1CAA25B7-DF28-49AD-90E5-D61EBF2F950F}
// *********************************************************************//
  IRootItem = dispinterface
    ['{1CAA25B7-DF28-49AD-90E5-D61EBF2F950F}']
    property Id: WideString readonly dispid 1;
    property PriceNumber: IWorkItemPriceNumber readonly dispid 2;
    property Designation: IWorkItemDesignation readonly dispid 3;
    property Parent: IWorkItemParent readonly dispid 4;
    property Level: IWorkItemLevel readonly dispid 5;
    property ItemType: IWorkItemType readonly dispid 6;
    property WorkItems: IWorkItemCollection readonly dispid 7;
    function GetById(const Id: WideString): IWorkItem; dispid 128;
  end;

// *********************************************************************//
// DispIntf :  IWorkItemDesignation
// Flags :     (4096) Dispatchable
// GUID :      {63E40920-070E-47C4-B280-164389E33345}
// *********************************************************************//
  IWorkItemDesignation = dispinterface
    ['{63E40920-070E-47C4-B280-164389E33345}']
    property Current: IStringValue readonly dispid 1;
    property Original: IStringValue readonly dispid 2;
    property IsModified: WordBool readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf :  IMarking
// Flags :     (4096) Dispatchable
// GUID :      {FE491689-DC81-4444-88FD-78C1A83E9251}
// *********************************************************************//
  IMarking = dispinterface
    ['{FE491689-DC81-4444-88FD-78C1A83E9251}']
    property PageNumber: Integer readonly dispid 1;
    property Top: Integer readonly dispid 2;
  end;

// *********************************************************************//
// DispIntf :  IBillOfQuantity
// Flags :     (4096) Dispatchable
// GUID :      {B1475EC2-EF9D-46DB-8B99-C7E7CA25CCB6}
// *********************************************************************//
  IBillOfQuantity = dispinterface
    ['{B1475EC2-EF9D-46DB-8B99-C7E7CA25CCB6}']
    property project: IProject readonly dispid 1;
    property Properties: IBillOfQuantityProperties readonly dispid 2;
    property WorkItems: IWorkItemCollection readonly dispid 3;
    property Title: WideString readonly dispid 4;
    function FindWorkItemById(const workItemID: WideString): IWorkItem; dispid 128;
    function GetFlatWorkItemList(filter: ComWorkItemFilter): IWorkItemCollection; dispid 129;
  end;

// *********************************************************************//
// DispIntf :  IAddress
// Flags :     (4096) Dispatchable
// GUID :      {2317FFBB-45CD-4390-9B98-81B177F61623}
// *********************************************************************//
  IAddress = dispinterface
    ['{2317FFBB-45CD-4390-9B98-81B177F61623}']
    property Street1: WideString readonly dispid 1;
    property Street2: WideString readonly dispid 2;
    property PostalCode: WideString readonly dispid 3;
    property Town: WideString readonly dispid 4;
    property Cedex: WideString readonly dispid 5;
    property Country: WideString readonly dispid 6;
    property MailAddress: WideString readonly dispid 7;
  end;

// *********************************************************************//
// DispIntf :  IDealResult
// Flags :     (4096) Dispatchable
// GUID :      {73F675E7-DC1F-4F11-B909-C307F44AC420}
// *********************************************************************//
  IDealResult = dispinterface
    ['{73F675E7-DC1F-4F11-B909-C307F44AC420}']
    property Result: ExecResult readonly dispid 1;
    property SelectedDeal: IDeal readonly dispid 2;
  end;

// *********************************************************************//
// DispIntf :  IDIEDownloader
// Flags :     (4096) Dispatchable
// GUID :      {459C04A2-CC9B-4F20-81B0-1D549F33D0D9}
// *********************************************************************//
  IDIEDownloader = dispinterface
    ['{459C04A2-CC9B-4F20-81B0-1D549F33D0D9}']
    property dealId: Integer readonly dispid 1;
    property Events: IDIEDownloadEventProvider readonly dispid 2;
    function Start: WordBool; dispid 128;
    procedure Cancel; dispid 129;
  end;

// *********************************************************************//
// DispIntf :  IProcedureType
// Flags :     (4096) Dispatchable
// GUID :      {B37B9697-99CC-4806-BB2C-412FFDFEF735}
// *********************************************************************//
  IProcedureType = dispinterface
    ['{B37B9697-99CC-4806-BB2C-412FFDFEF735}']
    property Code: WideString readonly dispid 1;
    property Designation: WideString readonly dispid 2;
  end;

// *********************************************************************//
// DispIntf :  IDetectionEventsTester
// Flags :     (4096) Dispatchable
// GUID :      {CFD41425-28E3-4960-8912-0C29BE7DF9BD}
// *********************************************************************//
  IDetectionEventsTester = dispinterface
    ['{CFD41425-28E3-4960-8912-0C29BE7DF9BD}']
    procedure RaiseEngineReady(delay: Integer); dispid 128;
    procedure RaiseAccountDataRequired(AccountType: AccountType; delay: Integer); dispid 129;
    procedure RaiseAccountStateChanged(delay: Integer); dispid 130;
    procedure RaiseUserRequestNewDeals(delay: Integer); dispid 131;
    procedure RaiseUserRequestSelectedDeals(delay: Integer); dispid 132;
    procedure RaiseUserRequestUpdatedDeals(delay: Integer); dispid 133;
    procedure RaiseUserRequestDealsWithAvailableDIE(delay: Integer); dispid 134;
  end;

// *********************************************************************//
// DispIntf :  ISpigaoCredentials
// Flags :     (4096) Dispatchable
// GUID :      {59B6A2A2-0AB8-4601-9399-05B44AB299CD}
// *********************************************************************//
  ISpigaoCredentials = dispinterface
    ['{59B6A2A2-0AB8-4601-9399-05B44AB299CD}']
    property Domain: WideString readonly dispid 1;
    property UserId: WideString readonly dispid 2;
    property password: WideString readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf :  IDCEDocumentViewer
// Flags :     (4096) Dispatchable
// GUID :      {58971E8C-A1CE-459C-B5C0-C0D20A189A6C}
// *********************************************************************//
  IDCEDocumentViewer = dispinterface
    ['{58971E8C-A1CE-459C-B5C0-C0D20A189A6C}']
    procedure ShowDocument(DCEFileType: DCEFileType); dispid 128;
  end;

// *********************************************************************//
// DispIntf :  IWorkItemSelection
// Flags :     (4096) Dispatchable
// GUID :      {B2E87ED2-83E6-40A1-9146-5B3B2E7E69BE}
// *********************************************************************//
  IWorkItemSelection = dispinterface
    ['{B2E87ED2-83E6-40A1-9146-5B3B2E7E69BE}']
    property SelectedCount: Integer readonly dispid 1;
    property WorkItems: IWorkItemCollection readonly dispid 2;
    procedure Add(const Item: IWorkItem); dispid 128;
    procedure Remove(const Item: IWorkItem); dispid 129;
    procedure Clear; dispid 130;
  end;

// *********************************************************************//
// DispIntf :  IWorkItemParent
// Flags :     (4096) Dispatchable
// GUID :      {AA6B168A-BD59-4B00-B011-4B5DE635DEC9}
// *********************************************************************//
  IWorkItemParent = dispinterface
    ['{AA6B168A-BD59-4B00-B011-4B5DE635DEC9}']
    property Current: IGroupItemValue readonly dispid 1;
    property Original: IGroupItemValue readonly dispid 2;
    property IsModified: WordBool readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf :  IPriceItem
// Flags :     (4096) Dispatchable
// GUID :      {8771822E-7F90-4879-A1C9-B33E1EF935E8}
// *********************************************************************//
  IPriceItem = dispinterface
    ['{8771822E-7F90-4879-A1C9-B33E1EF935E8}']
    property Id: WideString readonly dispid 1;
    property PriceNumber: IWorkItemPriceNumber readonly dispid 2;
    property Designation: IWorkItemDesignation readonly dispid 3;
    property Parent: IWorkItemParent readonly dispid 4;
    property Level: IWorkItemLevel readonly dispid 5;
    property ItemType: IWorkItemType readonly dispid 6;
    property Quantity: IWorkItemQuantity readonly dispid 7;
    property Unit_: IWorkItemUnit readonly dispid 8;
    property UnitPrice: IWorkItemUnitPrice readonly dispid 9;
    property IsDuplicated: WordBool readonly dispid 10;
    property IsLinked: WordBool readonly dispid 11;
    property IsReference: WordBool readonly dispid 12;
    property ReferenceItem: IPriceItem readonly dispid 13;
    property IsPM: WordBool readonly dispid 14;
    property IsFixed: WordBool readonly dispid 15;
    function GetDuplicatedItems: IWorkItemCollection; dispid 128;
    function GetLinkedItems: IWorkItemCollection; dispid 129;
    procedure SetUnitPrice(value: Double); dispid 130;
    procedure SetQuantity(value: Double); dispid 131;
  end;

// *********************************************************************//
// DispIntf :  INumericValue
// Flags :     (4096) Dispatchable
// GUID :      {79870506-7AA1-4D32-9E47-6FBAC903E791}
// *********************************************************************//
  INumericValue = dispinterface
    ['{79870506-7AA1-4D32-9E47-6FBAC903E791}']
    property IsFormula: WordBool readonly dispid 1;
    property Formula: IFormula readonly dispid 2;
    property value: Double readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf :  IMessage
// Flags :     (4096) Dispatchable
// GUID :      {84CEF100-7B36-4A22-A833-0E568AEE7FA0}
// *********************************************************************//
  IMessage = dispinterface
    ['{84CEF100-7B36-4A22-A833-0E568AEE7FA0}']
    property Category: ComMessageCategory readonly dispid 1;
    property Text: WideString readonly dispid 2;
  end;

// *********************************************************************//
// DispIntf :  ICompany
// Flags :     (4096) Dispatchable
// GUID :      {6CEE5935-AAA3-442B-962B-59098CFECB8A}
// *********************************************************************//
  ICompany = dispinterface
    ['{6CEE5935-AAA3-442B-962B-59098CFECB8A}']
    property Name: WideString readonly dispid 1;
    property ShortName: WideString readonly dispid 2;
    property RegistrationNumber: WideString readonly dispid 3;
    property Address: IAddress readonly dispid 4;
    property WebSite: WideString readonly dispid 5;
    property Id: Integer readonly dispid 6;
  end;

// *********************************************************************//
// DispIntf :  ITender
// Flags :     (4096) Dispatchable
// GUID :      {740564AB-49F0-433F-A412-A4E5993F907E}
// *********************************************************************//
  ITender = dispinterface
    ['{740564AB-49F0-433F-A412-A4E5993F907E}']
    property Reference: WideString readonly dispid 1;
    property Subject: WideString readonly dispid 2;
    property Caption: WideString readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf :  IDIE
// Flags :     (4096) Dispatchable
// GUID :      {2B888D4C-1248-4E8B-8AFC-A47D966DCC65}
// *********************************************************************//
  IDIE = dispinterface
    ['{2B888D4C-1248-4E8B-8AFC-A47D966DCC65}']
    property PublishedOn: TDateTime readonly dispid 2;
    property LastModifiedOn: TDateTime readonly dispid 3;
    property Status: DIEStatus readonly dispid 4;
  end;

// *********************************************************************//
// DispIntf :  IAccountData
// Flags :     (4096) Dispatchable
// GUID :      {65E7B8CF-C5F2-4772-BE0F-65898FC4B5D6}
// *********************************************************************//
  IAccountData = dispinterface
    ['{65E7B8CF-C5F2-4772-BE0F-65898FC4B5D6}']
    property CompanyName: WideString dispid 1;
    property RegistrationNumber: WideString dispid 2;
    property PostCode: WideString dispid 3;
    property ActivityCode: WideString dispid 4;
    property FirstName: WideString dispid 5;
    property LastName: WideString dispid 6;
    property Email: WideString dispid 7;
    property PhoneNumber: WideString dispid 8;
  end;

// *********************************************************************//
// DispIntf :  IProcessingUI
// Flags :     (4096) Dispatchable
// GUID :      {D4ED8C67-1284-40A6-B7D9-2405AA1CC744}
// *********************************************************************//
  IProcessingUI = dispinterface
    ['{D4ED8C67-1284-40A6-B7D9-2405AA1CC744}']
    property Options: IProcessingUIOptions readonly dispid 1;
    function ShowImportWizard(const project: IProject; generateImportFile: WordBool): IImportResult; dispid 128;
    function ShowExportWizard(const project: IProject; const exportFilePath: WideString): ExecResult; dispid 129;
    function ShowBPX(const project: IProject; const workItemID: WideString; docType: ComDocumentType): ExecResult; dispid 130;
    function ShowGlobalSettings: ExecResult; dispid 131;
    function ShowUserSettings: ExecResult; dispid 132;
    function ShowAbout: ExecResult; dispid 133;
    function ImportProject(dealId: Integer): IImportResult; dispid 134;
    function ImportProjectFromDisk(const ProjectPath: WideString): IImportResult; dispid 135;
    function ExportProject(const project: IProject; const exportFilePath: WideString): ExecResult; dispid 136;
    function ViewDocuments(const project: IProject; const workItemID: WideString; 
                           docType: ComDocumentType): ExecResult; dispid 137;
  end;

// *********************************************************************//
// DispIntf :  IBillOfQuantityEditor
// Flags :     (4096) Dispatchable
// GUID :      {961A7B33-016E-452B-B3A4-014D03D2237D}
// *********************************************************************//
  IBillOfQuantityEditor = dispinterface
    ['{961A7B33-016E-452B-B3A4-014D03D2237D}']
    property CanUndo: WordBool readonly dispid 1;
    property CanRedo: WordBool readonly dispid 2;
    property CurrentSelection: IWorkItemSelection readonly dispid 3;
    function IsCommandAvailable(command: BillOfQuantityEditCommand): WordBool; dispid 128;
    function LevelUp: WordBool; dispid 129;
    function LevelDown(const newChapterPriceNumber: WideString; 
                       const newChapterDesignation: WideString): WordBool; dispid 130;
    function TransformToComment: WordBool; dispid 131;
    function TransformToChapter: WordBool; dispid 132;
    procedure Undo; dispid 133;
    procedure Redo; dispid 134;
    procedure Validate; dispid 135;
    function EnablePriceItemLinks: WordBool; dispid 136;
    function DisablePriceItemLinks: WordBool; dispid 137;
  end;

// *********************************************************************//
// DispIntf :  IProject
// Flags :     (4096) Dispatchable
// GUID :      {9BF53378-3410-4317-9B95-3CDC86B64269}
// *********************************************************************//
  IProject = dispinterface
    ['{9BF53378-3410-4317-9B95-3CDC86B64269}']
    property Id: WideString readonly dispid 1;
    property Name: WideString readonly dispid 2;
    property Contract: IContract readonly dispid 3;
    property BillOfQuantity: IBillOfQuantity readonly dispid 4;
    property Messages: IMessageCollection readonly dispid 5;
    property Path: WideString readonly dispid 6;
    property Reference: WideString readonly dispid 7;
  end;

// *********************************************************************//
// DispIntf :  ISpigaoMessage
// Flags :     (4096) Dispatchable
// GUID :      {2BF4A787-2A10-427B-8DC3-8E861D2381FF}
// *********************************************************************//
  ISpigaoMessage = dispinterface
    ['{2BF4A787-2A10-427B-8DC3-8E861D2381FF}']
    property Content: WideString dispid 1;
    property Title: WideString dispid 2;
    property Url: WideString dispid 3;
  end;

// *********************************************************************//
// DispIntf :  IDIECollection
// Flags :     (4096) Dispatchable
// GUID :      {B47EA317-25FC-4A31-A018-29BD4F823EEE}
// *********************************************************************//
  IDIECollection = dispinterface
    ['{B47EA317-25FC-4A31-A018-29BD4F823EEE}']
    property Count: Integer readonly dispid 1;
    property Item[index: Integer]: IDIE readonly dispid 2;
  end;

// *********************************************************************//
// DispIntf :  IProcessingUIOptions
// Flags :     (4096) Dispatchable
// GUID :      {95560364-14E3-4352-B1F7-9FFB1D9B38C0}
// *********************************************************************//
  IProcessingUIOptions = dispinterface
    ['{95560364-14E3-4352-B1F7-9FFB1D9B38C0}']
    property generateImportFile: WordBool dispid 1;
    property UseScribe: WordBool dispid 2;
  end;

// *********************************************************************//
// DispIntf :  IImportResult
// Flags :     (4096) Dispatchable
// GUID :      {A6B2DBA7-9E8F-4B28-AF64-6BCE9096B977}
// *********************************************************************//
  IImportResult = dispinterface
    ['{A6B2DBA7-9E8F-4B28-AF64-6BCE9096B977}']
    property ExecResult: ExecResult readonly dispid 1;
    property project: IProject readonly dispid 2;
    property ImportFilename: WideString readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf :  IWorkItemUnit
// Flags :     (4096) Dispatchable
// GUID :      {52BB6C76-9CD4-40EA-866D-A00A6A16F044}
// *********************************************************************//
  IWorkItemUnit = dispinterface
    ['{52BB6C76-9CD4-40EA-866D-A00A6A16F044}']
    property Current: IStringValue readonly dispid 1;
    property Original: IStringValue readonly dispid 2;
    property IsModified: WordBool readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf :  IWorkItemLevel
// Flags :     (4096) Dispatchable
// GUID :      {DAB6448C-FBCA-4EAC-A9B5-6D52A786C4D3}
// *********************************************************************//
  IWorkItemLevel = dispinterface
    ['{DAB6448C-FBCA-4EAC-A9B5-6D52A786C4D3}']
    property Current: IIntegerValue readonly dispid 1;
    property Original: IIntegerValue readonly dispid 2;
    property IsModified: WordBool readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf :  IMessageCollection
// Flags :     (4096) Dispatchable
// GUID :      {ABD88D6E-3E34-4C09-A222-E0BA1073530A}
// *********************************************************************//
  IMessageCollection = dispinterface
    ['{ABD88D6E-3E34-4C09-A222-E0BA1073530A}']
    property Informations: {??PSafeArray}OleVariant readonly dispid 1;
    property Warnings: {??PSafeArray}OleVariant readonly dispid 2;
    property Errors: {??PSafeArray}OleVariant readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf :  IError
// Flags :     (4096) Dispatchable
// GUID :      {CD2DA701-D5DC-490C-B84F-A4A5EB845371}
// *********************************************************************//
  IError = dispinterface
    ['{CD2DA701-D5DC-490C-B84F-A4A5EB845371}']
    property Code: {??Int64}OleVariant readonly dispid 3;
    property Message: WideString readonly dispid 1;
    property StackTrace: WideString readonly dispid 2;
  end;

// *********************************************************************//
// DispIntf :  INotificationService
// Flags :     (4096) Dispatchable
// GUID :      {69629DC5-4017-49FA-8688-30CCC92D9088}
// *********************************************************************//
  INotificationService = dispinterface
    ['{69629DC5-4017-49FA-8688-30CCC92D9088}']
    function RetrieveAlert: ISpigaoAlert; dispid 128;
    function RetrieveMessages: ISpigaoMessageCollection; dispid 129;
  end;

// *********************************************************************//
// DispIntf :  IGroupItemValue
// Flags :     (4096) Dispatchable
// GUID :      {56F79E31-D5CF-4EF3-A47B-2962228DB16D}
// *********************************************************************//
  IGroupItemValue = dispinterface
    ['{56F79E31-D5CF-4EF3-A47B-2962228DB16D}']
    property value: IGroupItem readonly dispid 1;
  end;

// *********************************************************************//
// DispIntf :  IBillOfQuantityProperties
// Flags :     (4096) Dispatchable
// GUID :      {03ABDABB-CABB-42A9-B32F-81AB06232127}
// *********************************************************************//
  IBillOfQuantityProperties = dispinterface
    ['{03ABDABB-CABB-42A9-B32F-81AB06232127}']
    property HasDuplicates: WordBool readonly dispid 2;
    property HasLinks: WordBool readonly dispid 3;
    property DecimalDigits: Integer readonly dispid 4;
    property VatRate: Double readonly dispid 5;
    property HasMultipleColumns: WordBool readonly dispid 6;
  end;

// *********************************************************************//
// DispIntf :  IDetectionUI
// Flags :     (4096) Dispatchable
// GUID :      {428CA8F9-BB3D-40C5-8C3F-F3C8BF25B5BE}
// *********************************************************************//
  IDetectionUI = dispinterface
    ['{428CA8F9-BB3D-40C5-8C3F-F3C8BF25B5BE}']
    property Deals: IDealUI readonly dispid 1;
    property Notifications: INotificationUI readonly dispid 2;
    property Account: IAccountUI readonly dispid 3;
    function ShowDealImporter(allowMultiple: WordBool): IDealCollection; dispid 128;
    procedure ShowUserSettings; dispid 129;
    procedure ShowProxySettings; dispid 130;
    procedure ShowNewAndUpdatedDeals; dispid 131;
  end;

// *********************************************************************//
// DispIntf :  IDealSummary
// Flags :     (4096) Dispatchable
// GUID :      {398A1177-E48E-4E8A-A9A9-940D17E450B6}
// *********************************************************************//
  IDealSummary = dispinterface
    ['{398A1177-E48E-4E8A-A9A9-940D17E450B6}']
    property Id: Integer readonly dispid 1;
    property Subject: WideString readonly dispid 2;
    property Reference: WideString readonly dispid 3;
    property Title: WideString readonly dispid 4;
    property ContractorName: WideString readonly dispid 5;
    property DeadLine: TDateTime readonly dispid 6;
    property Departements: WideString readonly dispid 7;
    property IsDCELinkAvailable: WordBool readonly dispid 8;
    property HasDCEDocuments: WordBool readonly dispid 9;
    property ProcedureType: WideString readonly dispid 10;
    property Caption: WideString readonly dispid 11;
    property DIEStatus: DIEStatus readonly dispid 12;
    property DealStatus: DealStatus readonly dispid 13;
    property UpdateState: IDealUpdateState readonly dispid 14;
    function IsDCEDocumentAvailable(docType: DCEFileType): WordBool; dispid 128;
  end;

// *********************************************************************//
// DispIntf :  IPropertyCollection
// Flags :     (4096) Dispatchable
// GUID :      {10A300BF-44D5-48B1-B62B-15534C04B4A2}
// *********************************************************************//
  IPropertyCollection = dispinterface
    ['{10A300BF-44D5-48B1-B62B-15534C04B4A2}']
    property Item[index: Integer]: IProperty readonly dispid 1;
  end;

// *********************************************************************//
// DispIntf :  INotice
// Flags :     (4096) Dispatchable
// GUID :      {AAB12904-3549-494B-A30A-C6F5CA47F9AC}
// *********************************************************************//
  INotice = dispinterface
    ['{AAB12904-3549-494B-A30A-C6F5CA47F9AC}']
    property Origin: WideString readonly dispid 1;
    property OriginReference: WideString readonly dispid 2;
    property Code: WideString readonly dispid 3;
    property Designation: WideString readonly dispid 4;
  end;

// *********************************************************************//
// DispIntf :  IStringCollection
// Flags :     (4096) Dispatchable
// GUID :      {50F9DD60-62B6-40E4-96EB-FEFAB7A9BC9F}
// *********************************************************************//
  IStringCollection = dispinterface
    ['{50F9DD60-62B6-40E4-96EB-FEFAB7A9BC9F}']
    property Count: Integer readonly dispid 1;
    property Item[index: Integer]: WideString readonly dispid 2;
  end;

// *********************************************************************//
// DispIntf :  ISpigao
// Flags :     (4096) Dispatchable
// GUID :      {A6935D58-5651-439D-8DE7-5DC5441136CE}
// *********************************************************************//
  ISpigao = dispinterface
    ['{A6935D58-5651-439D-8DE7-5DC5441136CE}']
    property Processing: IProcessingEngine readonly dispid 1;
    property Detection: IDetectionEngine readonly dispid 2;
    property Version: IVersion readonly dispid 3;
    property Mode: ExecutionMode readonly dispid 4;
    function Unlock(const editorName: WideString; const softwareName: WideString; 
                    const licenseKey: WideString): WordBool; dispid 128;
    procedure ShowAbout; dispid 129;
  end;

// *********************************************************************//
// DispIntf :  ISpigaoMessageCollection
// Flags :     (4096) Dispatchable
// GUID :      {36BC6A5F-740D-4DFE-B2C0-A155648145FA}
// *********************************************************************//
  ISpigaoMessageCollection = dispinterface
    ['{36BC6A5F-740D-4DFE-B2C0-A155648145FA}']
    property FullCount: Integer readonly dispid 1;
    property Item[index: Integer]: ISpigaoMessage readonly dispid 2;
    property Count: Integer readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf :  ISpigaoWebSite
// Flags :     (4096) Dispatchable
// GUID :      {4C5961E0-25CF-48E9-A645-74F7D67762AD}
// *********************************************************************//
  ISpigaoWebSite = dispinterface
    ['{4C5961E0-25CF-48E9-A645-74F7D67762AD}']
    property HomePageUrl: WideString readonly dispid 1;
    property SupportPageUrl: WideString readonly dispid 2;
    property AlertSubscriptionPageUrl: WideString readonly dispid 3;
    function GetDIEPageUrl(dealId: Integer): WideString; dispid 128;
    function GetDealPageUrl(dealId: Integer): WideString; dispid 129;
    function GetDCEFilePageUrl(dealId: Integer): WideString; dispid 130;
    function GetDCEDisplayPageUrl(dealId: Integer; const fileType: WideString): WideString; dispid 131;
    function GetFullAccountActivationPageUrl: WideString; dispid 132;
    function GetLinkEditorPartnership: WideString; dispid 133;
  end;

// *********************************************************************//
// DispIntf :  IDCEFile
// Flags :     (4096) Dispatchable
// GUID :      {81CBD17F-F5BF-415D-A0EF-E3CCA3D72BD4}
// *********************************************************************//
  IDCEFile = dispinterface
    ['{81CBD17F-F5BF-415D-A0EF-E3CCA3D72BD4}']
    property Name: WideString readonly dispid 1;
    property fileType: WideString readonly dispid 2;
    property PublishedOn: TDateTime readonly dispid 3;
    property type_: DCEFileType readonly dispid 4;
  end;

// *********************************************************************//
// DispIntf :  IActivityCollection
// Flags :     (4096) Dispatchable
// GUID :      {E69C079D-C5CE-485C-B57C-63751B8C5FB6}
// *********************************************************************//
  IActivityCollection = dispinterface
    ['{E69C079D-C5CE-485C-B57C-63751B8C5FB6}']
    property FullCount: Integer readonly dispid 1;
    property Item[index: Integer]: IActivity readonly dispid 2;
    property Count: Integer readonly dispid 3;
    property ToString[const separator: WideString]: WideString readonly dispid 128;
  end;

// *********************************************************************//
// DispIntf :  ITrialAccountState
// Flags :     (4096) Dispatchable
// GUID :      {8B1A2847-80CF-4EA3-9B1A-6FBA01309944}
// *********************************************************************//
  ITrialAccountState = dispinterface
    ['{8B1A2847-80CF-4EA3-9B1A-6FBA01309944}']
    property EndDate: TDateTime readonly dispid 1;
    property ImportedDeals: Integer readonly dispid 2;
    property ImportedDealsLimit: Integer readonly dispid 3;
    property RemainingDays: Integer readonly dispid 4;
    property StartDate: TDateTime readonly dispid 5;
  end;

// *********************************************************************//
// DispIntf :  IWorkItemTypeValue
// Flags :     (4096) Dispatchable
// GUID :      {45821B51-333D-4E04-A1CD-131B3739A2CA}
// *********************************************************************//
  IWorkItemTypeValue = dispinterface
    ['{45821B51-333D-4E04-A1CD-131B3739A2CA}']
    property value: ComWorkItemType readonly dispid 1;
  end;

// *********************************************************************//
// DispIntf :  IStringValue
// Flags :     (4096) Dispatchable
// GUID :      {75C64598-7052-4731-BD20-40A8297D40E4}
// *********************************************************************//
  IStringValue = dispinterface
    ['{75C64598-7052-4731-BD20-40A8297D40E4}']
    property value: WideString readonly dispid 1;
  end;

// *********************************************************************//
// DispIntf :  IProjectExporter
// Flags :     (4096) Dispatchable
// GUID :      {FEEE5573-72CF-445F-8F5F-A324511CD481}
// *********************************************************************//
  IProjectExporter = dispinterface
    ['{FEEE5573-72CF-445F-8F5F-A324511CD481}']
    procedure LoadData(const documentId: WideString); dispid 128;
    function Run: ExecResult; dispid 129;
    function GeneratePDFFiles(const folder: WideString): IPDFGeneration; dispid 130;
    property Report: IReport readonly dispid 1;
  end;

// *********************************************************************//
// DispIntf :  IProcessingEngine
// Flags :     (4096) Dispatchable
// GUID :      {D73ECB5A-E5CA-46E3-B164-84C50D7EFF52}
// *********************************************************************//
  IProcessingEngine = dispinterface
    ['{D73ECB5A-E5CA-46E3-B164-84C50D7EFF52}']
    property GlobalSettings: IProcessingAppConfig readonly dispid 1;
    property UserSettings: IProcessingUserConfig readonly dispid 2;
    property UI: IProcessingUI readonly dispid 3;
    property DataFolder: WideString readonly dispid 4;
    property Utils: IProcessingUtils readonly dispid 5;
    function LoadProject(const documentId: WideString): IProject; dispid 128;
    function ImportProject(const project: IProject): IProjectImporter; dispid 129;
    function ExportProject(const project: IProject): IProjectExporter; dispid 130;
    function GetLastError: IError; dispid 131;
  end;

// *********************************************************************//
// DispIntf :  IContract
// Flags :     (4096) Dispatchable
// GUID :      {04ED3DFF-6A97-41C9-A94A-81336769527F}
// *********************************************************************//
  IContract = dispinterface
    ['{04ED3DFF-6A97-41C9-A94A-81336769527F}']
    property DeadLine: TDateTime readonly dispid 1;
    property MarketType: WideString readonly dispid 2;
    property ProcedureType: WideString readonly dispid 3;
    property Contractor: ICompany readonly dispid 4;
    property Supervisor: ICompany readonly dispid 5;
    property Lots: ILotCollection readonly dispid 6;
    property Notices: INoticeCollection readonly dispid 7;
  end;

// *********************************************************************//
// DispIntf :  ITaskResult
// Flags :     (4096) Dispatchable
// GUID :      {5D675097-FF60-43A7-BF62-B0D047A4D713}
// *********************************************************************//
  ITaskResult = dispinterface
    ['{5D675097-FF60-43A7-BF62-B0D047A4D713}']
    property Result: ExecResult readonly dispid 1;
    property Message: WideString readonly dispid 2;
  end;

// *********************************************************************//
// DispIntf :  IDealUI
// Flags :     (4096) Dispatchable
// GUID :      {EF506804-2AEE-42B0-A6BE-03D192E57D34}
// *********************************************************************//
  IDealUI = dispinterface
    ['{EF506804-2AEE-42B0-A6BE-03D192E57D34}']
    property Options: IDealUIOptions readonly dispid 1;
    function ListNewDeals: IDealResult; dispid 128;
    function ListSelectedDeals: IDealResult; dispid 129;
    function ListUpdatedDeals: ExecResult; dispid 130;
    function ListDealsWithAvailableDIE: IDealResult; dispid 131;
    function ViewDCEDocuments(dealId: Integer; documentType: DCEFileType): ExecResult; dispid 132;
    function DownloadDIE(dealId: Integer): IDIEResult; dispid 133;
    function ViewDeal(dealId: Integer): ExecResult; dispid 134;
  end;

// *********************************************************************//
// DispIntf :  ISpigaoAlert
// Flags :     (4096) Dispatchable
// GUID :      {FFC4247D-2706-4850-9D13-D641019899A8}
// *********************************************************************//
  ISpigaoAlert = dispinterface
    ['{FFC4247D-2706-4850-9D13-D641019899A8}']
    property Title: WideString readonly dispid 1;
    property Content: WideString readonly dispid 2;
    property NbNewDeals: Integer readonly dispid 3;
    property NbUpdatedDeals: Integer readonly dispid 4;
    property NbAvailableDIE: Integer readonly dispid 5;
    property HelpUrl: WideString readonly dispid 6;
  end;

// *********************************************************************//
// DispIntf :  IDealSummaryCollection
// Flags :     (4096) Dispatchable
// GUID :      {754E11E6-A29D-44B8-9A8A-A8C669F30289}
// *********************************************************************//
  IDealSummaryCollection = dispinterface
    ['{754E11E6-A29D-44B8-9A8A-A8C669F30289}']
    property FullCount: Integer readonly dispid 1;
    property Item[index: Integer]: IDealSummary readonly dispid 2;
    property Count: Integer readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf :  IDealUpdateState
// Flags :     (4096) Dispatchable
// GUID :      {A2A121B0-2527-46C1-870D-3B2298D58753}
// *********************************************************************//
  IDealUpdateState = dispinterface
    ['{A2A121B0-2527-46C1-870D-3B2298D58753}']
    property HasUpdatedData: WordBool readonly dispid 1;
    property UpdatedData: DealUpdatedData readonly dispid 2;
    function IsDataUpdated(data: DealUpdatedData): WordBool; dispid 128;
  end;

// *********************************************************************//
// DispIntf :  IDownloadProgressChangedEventArgs
// Flags :     (4096) Dispatchable
// GUID :      {015A840E-CFBA-42D8-8F2F-72F13DC7D3AE}
// *********************************************************************//
  IDownloadProgressChangedEventArgs = dispinterface
    ['{015A840E-CFBA-42D8-8F2F-72F13DC7D3AE}']
    property TotalBytes: {??Int64}OleVariant readonly dispid 1;
    property DownloadedBytes: {??Int64}OleVariant readonly dispid 2;
  end;

// *********************************************************************//
// DispIntf :  IDetectionAppConfig
// Flags :     (4096) Dispatchable
// GUID :      {CABEEE57-810B-4256-8C99-A7674BEE996D}
// *********************************************************************//
  IDetectionAppConfig = dispinterface
    ['{CABEEE57-810B-4256-8C99-A7674BEE996D}']
    property UseStub: WordBool readonly dispid 1;
    property ServiceUrl: WideString readonly dispid 2;
    property AllowAccountActivationUI: WordBool readonly dispid 3;
    property DisplayTestData: WordBool readonly dispid 4;
  end;

// *********************************************************************//
// DispIntf :  IDeal
// Flags :     (4096) Dispatchable
// GUID :      {AF76D763-D1CC-4CBC-BA4F-B888BA09F031}
// *********************************************************************//
  IDeal = dispinterface
    ['{AF76D763-D1CC-4CBC-BA4F-B888BA09F031}']
    property Id: Integer readonly dispid 1;
    property Reference: WideString readonly dispid 2;
    property Subject: WideString readonly dispid 3;
    property Activities: IActivityCollection readonly dispid 4;
    property Tender: ITender readonly dispid 5;
    property Lot: ILot readonly dispid 6;
    property Contractor: ICompany readonly dispid 7;
    property Supervisor: ICompany readonly dispid 8;
    property DeadLine: TDateTime readonly dispid 9;
    property Territories: ITerritoryCollection readonly dispid 10;
    property DCEFiles: IDCEFileCollection readonly dispid 11;
    property PublishedOn: TDateTime readonly dispid 12;
    property MarketType: IMarketType readonly dispid 14;
    property ProcedureType: IProcedureType readonly dispid 15;
    property IsRestricted: RestrictionType readonly dispid 16;
    property LastModifiedOn: TDateTime readonly dispid 17;
    property IsPrivate: WordBool readonly dispid 18;
    property Notice: INotice readonly dispid 19;
    property DIE: IDIE readonly dispid 20;
    property OriginUrl: WideString readonly dispid 21;
    property Caption: WideString readonly dispid 22;
    property Status: DealStatus readonly dispid 23;
    function IsDocumentAvailable(docType: DCEFileType): WordBool; dispid 128;
  end;

// *********************************************************************//
// DispIntf :  IAccount
// Flags :     (4096) Dispatchable
// GUID :      {5CC64EE1-9C83-48CF-A772-E07C0DFB948A}
// *********************************************************************//
  IAccount = dispinterface
    ['{5CC64EE1-9C83-48CF-A772-E07C0DFB948A}']
    property Key: WideString readonly dispid 1;
    property data: IAccountData readonly dispid 2;
    property type_: AccountType readonly dispid 3;
    property Status: AccountStatus readonly dispid 4;
    property Credentials: ISpigaoCredentials readonly dispid 5;
    property Trial: ITrialAccountState readonly dispid 6;
    function CanActivateTrialAccount(var isActivated: WordBool): WordBool; dispid 7;
    function ActivateStarterAccount(const accountKey: WideString): WordBool; dispid 128;
    function ActivateTrialAccount: WordBool; dispid 129;
    function SetFullAccountCredentials(const Domain: WideString; const UserId: WideString; 
                                       const password: WideString): WordBool; dispid 130;
    function DisableAccount: WordBool; dispid 131;
  end;

// *********************************************************************//
// DispIntf :  IWorkItemQuantity
// Flags :     (4096) Dispatchable
// GUID :      {BE8D140D-F90A-41C5-B82F-DE7C6C1563DD}
// *********************************************************************//
  IWorkItemQuantity = dispinterface
    ['{BE8D140D-F90A-41C5-B82F-DE7C6C1563DD}']
    property Current: INumericValue readonly dispid 1;
    property Original: INumericValue readonly dispid 2;
    property IsModified: WordBool readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf :  IWorkItemAmount
// Flags :     (4096) Dispatchable
// GUID :      {8184A93D-7E57-490E-B08B-2ABE17C6551D}
// *********************************************************************//
  IWorkItemAmount = dispinterface
    ['{8184A93D-7E57-490E-B08B-2ABE17C6551D}']
  end;

// *********************************************************************//
// DispIntf :  IProperty
// Flags :     (4096) Dispatchable
// GUID :      {5640CC99-97AB-4551-87A5-AC215CC01F9B}
// *********************************************************************//
  IProperty = dispinterface
    ['{5640CC99-97AB-4551-87A5-AC215CC01F9B}']
    property Name: WideString readonly dispid 1;
    property value: WideString readonly dispid 2;
  end;

// *********************************************************************//
// DispIntf :  IAmountItem
// Flags :     (4096) Dispatchable
// GUID :      {0017D291-9DBF-4467-A72F-5EF69F9EC9C7}
// *********************************************************************//
  IAmountItem = dispinterface
    ['{0017D291-9DBF-4467-A72F-5EF69F9EC9C7}']
    property Amount: IWorkItemAmount readonly dispid 7;
    property ItemSubType: ComAmountType readonly dispid 8;
  end;

// *********************************************************************//
// DispIntf :  IReportMessage
// Flags :     (4096) Dispatchable
// GUID :      {45289823-532F-4367-BBC3-6F035FFA7329}
// *********************************************************************//
  IReportMessage = dispinterface
    ['{45289823-532F-4367-BBC3-6F035FFA7329}']
    property Text: WideString readonly dispid 1;
    property Severity: ReportSeverity readonly dispid 2;
    property HasDetails: WordBool readonly dispid 3;
    property Details: IReportMessageCollection readonly dispid 4;
  end;

// *********************************************************************//
// DispIntf :  ILotCollection
// Flags :     (4096) Dispatchable
// GUID :      {B6D39D68-0C8B-4856-8067-7AA77DCEC95C}
// *********************************************************************//
  ILotCollection = dispinterface
    ['{B6D39D68-0C8B-4856-8067-7AA77DCEC95C}']
    property Item[index: Integer]: ILot readonly dispid 1;
  end;

// *********************************************************************//
// DispIntf :  IResultEventArgs
// Flags :     (4096) Dispatchable
// GUID :      {600FF0C6-00CF-4EB2-8F5C-84705A58AEA9}
// *********************************************************************//
  IResultEventArgs = dispinterface
    ['{600FF0C6-00CF-4EB2-8F5C-84705A58AEA9}']
    property Result: OleVariant dispid 1;
    property type_: AlertType dispid 2;
  end;

// *********************************************************************//
// DispIntf :  IMarketType
// Flags :     (4096) Dispatchable
// GUID :      {D7098A5B-5369-4960-BB68-9F48130319EF}
// *********************************************************************//
  IMarketType = dispinterface
    ['{D7098A5B-5369-4960-BB68-9F48130319EF}']
    property Code: WideString readonly dispid 1;
    property Designation: WideString readonly dispid 2;
  end;

// *********************************************************************//
// DispIntf :  IFormula
// Flags :     (4096) Dispatchable
// GUID :      {646286ED-83E9-423E-9E7C-0FA3B0429B56}
// *********************************************************************//
  IFormula = dispinterface
    ['{646286ED-83E9-423E-9E7C-0FA3B0429B56}']
    property Definition: WideString readonly dispid 1;
  end;

// *********************************************************************//
// DispIntf :  IDetectionUserConfig
// Flags :     (4096) Dispatchable
// GUID :      {8915E454-FBA0-45B9-8B29-4C8C398288F2}
// *********************************************************************//
  IDetectionUserConfig = dispinterface
    ['{8915E454-FBA0-45B9-8B29-4C8C398288F2}']
    property SpigaoAccount: WideString dispid 1;
    property SpigaoUserId: WideString dispid 2;
    property SpigaoPassword: WideString dispid 3;
    property ProxySettings: IProxySettings readonly dispid 4;
    property AlertFrequency: AlertFrequency dispid 5;
    property DisplayAlerts: WordBool dispid 6;
    property DefaultStorageFolder: WideString dispid 7;
    property accountKey: WideString dispid 8;
    procedure Save; dispid 128;
  end;

// *********************************************************************//
// DispIntf :  ITerritoryCollection
// Flags :     (4096) Dispatchable
// GUID :      {7191CF2B-C3EE-4222-B481-B2F196996F5B}
// *********************************************************************//
  ITerritoryCollection = dispinterface
    ['{7191CF2B-C3EE-4222-B481-B2F196996F5B}']
    property FullCount: Integer readonly dispid 1;
    property Item[index: Integer]: ITerritory readonly dispid 2;
    property Count: Integer readonly dispid 3;
    property ToString[const separator: WideString]: WideString readonly dispid 128;
  end;

// *********************************************************************//
// DispIntf :  IActivity
// Flags :     (4096) Dispatchable
// GUID :      {D1DDE808-7F5E-4EDE-A26F-195DD058524F}
// *********************************************************************//
  IActivity = dispinterface
    ['{D1DDE808-7F5E-4EDE-A26F-195DD058524F}']
    property Code: WideString readonly dispid 1;
    property Designation: WideString readonly dispid 2;
  end;

// *********************************************************************//
// DispIntf :  ISearchResult
// Flags :     (4096) Dispatchable
// GUID :      {DC15D625-3BE9-41A7-B793-1F9CA2BBD64E}
// *********************************************************************//
  ISearchResult = dispinterface
    ['{DC15D625-3BE9-41A7-B793-1F9CA2BBD64E}']
    property Count: Integer dispid 1;
    property Url: WideString dispid 2;
  end;

// *********************************************************************//
// DispIntf :  IDCEFileCollection
// Flags :     (4096) Dispatchable
// GUID :      {3B3F3331-0727-4484-A446-CF79D901C836}
// *********************************************************************//
  IDCEFileCollection = dispinterface
    ['{3B3F3331-0727-4484-A446-CF79D901C836}']
    property FullCount: Integer readonly dispid 1;
    property Item[index: Integer]: IDCEFile readonly dispid 2;
    property Count: Integer readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf :  IPDFGeneration
// Flags :     (4096) Dispatchable
// GUID :      {B367C992-C24E-45D3-A525-1CC00036BE82}
// *********************************************************************//
  IPDFGeneration = dispinterface
    ['{B367C992-C24E-45D3-A525-1CC00036BE82}']
    property Status: ExecResult readonly dispid 1;
    property PDFFiles: IStringCollection readonly dispid 2;
  end;

// *********************************************************************//
// DispIntf :  IProcessingAppConfig
// Flags :     (4096) Dispatchable
// GUID :      {2BDB3EE3-0401-4300-B7D2-169FD9E819C9}
// *********************************************************************//
  IProcessingAppConfig = dispinterface
    ['{2BDB3EE3-0401-4300-B7D2-169FD9E819C9}']
    property MaxLineCount: Integer dispid 1;
    property ChapterImportMode: ImportMode dispid 2;
    property CommentImportMode: ImportMode dispid 3;
    property ChapterMandatory: WordBool dispid 4;
    property IdentAllowAlpha: WordBool dispid 5;
    property IdentMaxLength: Integer dispid 6;
    property DesignationMaxLength: Integer dispid 7;
    property MaxLevel: Integer dispid 8;
    property MaxLevelAction: MaxLevelAction dispid 9;
    property PriceNumberAllowAlpha: WordBool dispid 10;
    property PriceNumberMandatory: WordBool dispid 11;
    property PriceNumberMaxLength: Integer dispid 12;
    property PriceNumberForComment: WideString dispid 13;
    property PriceNumberRenumberingScope: PriceNumberRenumberingScope dispid 14;
    property UnitAllowedChars: WideString dispid 15;
    property UnitDefaultValue: WideString dispid 16;
    property UnitMaxLength: Integer dispid 17;
    property UnitReplaceString: WideString dispid 18;
    property QuantityDecimals: Integer dispid 19;
    property QuantityDefaultValue: Double dispid 20;
    property QuantityMandatory: WordBool dispid 21;
    property UnitPriceAllowNegative: WordBool dispid 22;
    property UnitPriceDecimals: Integer dispid 23;
    property ImportFileFormat: FileFormat dispid 31;
    property ImportTxtMapping: WideString dispid 32;
    property ImportTxtFieldSeparator: WideString dispid 33;
    property ExportTxtRowMapping: WideString dispid 34;
    property ExportTxtFieldSeparator: WideString dispid 35;
    property UIShowReport: WordBool readonly dispid 41;
    property ExportFileExtension: WideString dispid 42;
    property ExportFileFormat: FileFormat dispid 43;
    property ExportTxtSkippedRows: Integer dispid 44;
    property ImportFileExtension: WideString dispid 45;
    property ImportFileFormatter: WideString dispid 46;
    property UnitPriceAllowLinks: AllowLinksMode dispid 47;
  end;

// *********************************************************************//
// DispIntf :  IWorkItemCollection
// Flags :     (4096) Dispatchable
// GUID :      {1B10FD5F-83C8-44F3-AE09-B25E870F6E2D}
// *********************************************************************//
  IWorkItemCollection = dispinterface
    ['{1B10FD5F-83C8-44F3-AE09-B25E870F6E2D}']
    property Item[index: Integer]: IWorkItem readonly dispid 1;
    property Count: Integer readonly dispid 2;
  end;

// *********************************************************************//
// DispIntf :  ITerritory
// Flags :     (4096) Dispatchable
// GUID :      {AE7A746E-33F3-4A35-AD3F-5D01AE04A11C}
// *********************************************************************//
  ITerritory = dispinterface
    ['{AE7A746E-33F3-4A35-AD3F-5D01AE04A11C}']
    property Designation: WideString readonly dispid 1;
    property Code: WideString readonly dispid 2;
    property type_: TerritoryType readonly dispid 3;
  end;

// *********************************************************************//
// DispIntf :  INotificationUI
// Flags :     (4096) Dispatchable
// GUID :      {BD40D123-EB68-45D3-91EB-25CF6E02F4B0}
// *********************************************************************//
  INotificationUI = dispinterface
    ['{BD40D123-EB68-45D3-91EB-25CF6E02F4B0}']
    function ShowMessages: WordBool; dispid 128;
    function ShowAlert: WordBool; dispid 129;
  end;

// *********************************************************************//
// DispIntf :  IAccountUI
// Flags :     (4096) Dispatchable
// GUID :      {8A91ED49-620C-4314-B5C4-E9DB87EBA1FF}
// *********************************************************************//
  IAccountUI = dispinterface
    ['{8A91ED49-620C-4314-B5C4-E9DB87EBA1FF}']
    function ActivateTrialAccount: ExecResult; dispid 128;
  end;

// *********************************************************************//
// DispIntf :  IDealService
// Flags :     (4096) Dispatchable
// GUID :      {1A83ECE4-146F-4927-BEAD-5273C9971CAA}
// *********************************************************************//
  IDealService = dispinterface
    ['{1A83ECE4-146F-4927-BEAD-5273C9971CAA}']
    function CountNewDeals: Integer; dispid 128;
    function CountSelectedDeals: Integer; dispid 129;
    function CountUpdatedDeals: Integer; dispid 130;
    function CountDealsWithAvailableDIE: Integer; dispid 131;
    function ListNewDeals: IDealSummaryCollection; dispid 132;
    function ListSelectedDeals: IDealSummaryCollection; dispid 133;
    function ListUpdatedDeals: IDealSummaryCollection; dispid 134;
    function ListDealsWithAvailableDIE: IDealSummaryCollection; dispid 135;
    function GetDeal(dealId: Integer): IDeal; dispid 136;
    function MarkDealAsSelected(dealId: Integer; selected: WordBool): WordBool; dispid 137;
    function IsDIEAvailable(dealId: Integer): WordBool; dispid 138;
    function DownloadDIE(dealId: Integer): IDIEDownloader; dispid 139;
    function MarkDealAsIntegrated(dealId: Integer; integrated: WordBool): WordBool; dispid 140;
    function GetDealSummary(dealId: Integer): IDealSummary; dispid 141;
  end;

// *********************************************************************//
// La classe CoReportMessageCollection fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut IReportMessageCollection exposée             
// par la CoClasse ReportMessageCollection. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoReportMessageCollection = class
    class function Create: IReportMessageCollection;
    class function CreateRemote(const MachineName: string): IReportMessageCollection;
  end;

// *********************************************************************//
// La classe CoDealUIOptions fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut IDealUIOptions exposée             
// par la CoClasse DealUIOptions. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoDealUIOptions = class
    class function Create: IDealUIOptions;
    class function CreateRemote(const MachineName: string): IDealUIOptions;
  end;

// *********************************************************************//
// La classe CoReport fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut IReport exposée             
// par la CoClasse Report. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoReport = class
    class function Create: IReport;
    class function CreateRemote(const MachineName: string): IReport;
  end;

// *********************************************************************//
// La classe CoDetectionEventsTester fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut IDetectionEventsTester exposée             
// par la CoClasse DetectionEventsTester. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoDetectionEventsTester = class
    class function Create: IDetectionEventsTester;
    class function CreateRemote(const MachineName: string): IDetectionEventsTester;
  end;

// *********************************************************************//
// La classe CoDIEDownloadEventProvider fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut IDIEDownloadEventProvider exposée             
// par la CoClasse DIEDownloadEventProvider. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoDIEDownloadEventProvider = class
    class function Create: IDIEDownloadEventProvider;
    class function CreateRemote(const MachineName: string): IDIEDownloadEventProvider;
  end;

// *********************************************************************//
// La classe CoAccountData fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut IAccountData exposée             
// par la CoClasse AccountData. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoAccountData = class
    class function Create: IAccountData;
    class function CreateRemote(const MachineName: string): IAccountData;
  end;

// *********************************************************************//
// La classe CoComponentVersion fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut IVersion exposée             
// par la CoClasse ComponentVersion. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoComponentVersion = class
    class function Create: IVersion;
    class function CreateRemote(const MachineName: string): IVersion;
  end;

// *********************************************************************//
// La classe CoDetectionEventProvider fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut IDetectionEventProvider exposée             
// par la CoClasse DetectionEventProvider. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoDetectionEventProvider = class
    class function Create: IDetectionEventProvider;
    class function CreateRemote(const MachineName: string): IDetectionEventProvider;
  end;

// *********************************************************************//
// La classe CoError fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut IError exposée             
// par la CoClasse Error. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoError = class
    class function Create: IError;
    class function CreateRemote(const MachineName: string): IError;
  end;

// *********************************************************************//
// La classe CoSpigaoLoader fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut ISpigao exposée             
// par la CoClasse SpigaoLoader. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoSpigaoLoader = class
    class function Create: ISpigao;
    class function CreateRemote(const MachineName: string): ISpigao;
  end;

// *********************************************************************//
// La classe CoTaskResult fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut ITaskResult exposée             
// par la CoClasse TaskResult. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoTaskResult = class
    class function Create: ITaskResult;
    class function CreateRemote(const MachineName: string): ITaskResult;
  end;

// *********************************************************************//
// La classe CoReportMessage fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut IReportMessage exposée             
// par la CoClasse ReportMessage. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoReportMessage = class
    class function Create: IReportMessage;
    class function CreateRemote(const MachineName: string): IReportMessage;
  end;

// *********************************************************************//
// La classe CoStringCollection fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut IStringCollection exposée             
// par la CoClasse StringCollection. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoStringCollection = class
    class function Create: IStringCollection;
    class function CreateRemote(const MachineName: string): IStringCollection;
  end;

implementation

uses ComObj;

class function CoReportMessageCollection.Create: IReportMessageCollection;
begin
  Result := CreateComObject(CLASS_ReportMessageCollection) as IReportMessageCollection;
end;

class function CoReportMessageCollection.CreateRemote(const MachineName: string): IReportMessageCollection;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ReportMessageCollection) as IReportMessageCollection;
end;

class function CoDealUIOptions.Create: IDealUIOptions;
begin
  Result := CreateComObject(CLASS_DealUIOptions) as IDealUIOptions;
end;

class function CoDealUIOptions.CreateRemote(const MachineName: string): IDealUIOptions;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_DealUIOptions) as IDealUIOptions;
end;

class function CoReport.Create: IReport;
begin
  Result := CreateComObject(CLASS_Report) as IReport;
end;

class function CoReport.CreateRemote(const MachineName: string): IReport;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Report) as IReport;
end;

class function CoDetectionEventsTester.Create: IDetectionEventsTester;
begin
  Result := CreateComObject(CLASS_DetectionEventsTester) as IDetectionEventsTester;
end;

class function CoDetectionEventsTester.CreateRemote(const MachineName: string): IDetectionEventsTester;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_DetectionEventsTester) as IDetectionEventsTester;
end;

class function CoDIEDownloadEventProvider.Create: IDIEDownloadEventProvider;
begin
  Result := CreateComObject(CLASS_DIEDownloadEventProvider) as IDIEDownloadEventProvider;
end;

class function CoDIEDownloadEventProvider.CreateRemote(const MachineName: string): IDIEDownloadEventProvider;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_DIEDownloadEventProvider) as IDIEDownloadEventProvider;
end;

class function CoAccountData.Create: IAccountData;
begin
  Result := CreateComObject(CLASS_AccountData) as IAccountData;
end;

class function CoAccountData.CreateRemote(const MachineName: string): IAccountData;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_AccountData) as IAccountData;
end;

class function CoComponentVersion.Create: IVersion;
begin
  Result := CreateComObject(CLASS_ComponentVersion) as IVersion;
end;

class function CoComponentVersion.CreateRemote(const MachineName: string): IVersion;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ComponentVersion) as IVersion;
end;

class function CoDetectionEventProvider.Create: IDetectionEventProvider;
begin
  Result := CreateComObject(CLASS_DetectionEventProvider) as IDetectionEventProvider;
end;

class function CoDetectionEventProvider.CreateRemote(const MachineName: string): IDetectionEventProvider;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_DetectionEventProvider) as IDetectionEventProvider;
end;

class function CoError.Create: IError;
begin
  Result := CreateComObject(CLASS_Error) as IError;
end;

class function CoError.CreateRemote(const MachineName: string): IError;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Error) as IError;
end;

class function CoSpigaoLoader.Create: ISpigao;
begin
  Result := CreateComObject(CLASS_SpigaoLoader) as ISpigao;
end;

class function CoSpigaoLoader.CreateRemote(const MachineName: string): ISpigao;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_SpigaoLoader) as ISpigao;
end;

class function CoTaskResult.Create: ITaskResult;
begin
  Result := CreateComObject(CLASS_TaskResult) as ITaskResult;
end;

class function CoTaskResult.CreateRemote(const MachineName: string): ITaskResult;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_TaskResult) as ITaskResult;
end;

class function CoReportMessage.Create: IReportMessage;
begin
  Result := CreateComObject(CLASS_ReportMessage) as IReportMessage;
end;

class function CoReportMessage.CreateRemote(const MachineName: string): IReportMessage;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ReportMessage) as IReportMessage;
end;

class function CoStringCollection.Create: IStringCollection;
begin
  Result := CreateComObject(CLASS_StringCollection) as IStringCollection;
end;

class function CoStringCollection.CreateRemote(const MachineName: string): IStringCollection;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_StringCollection) as IStringCollection;
end;

end.

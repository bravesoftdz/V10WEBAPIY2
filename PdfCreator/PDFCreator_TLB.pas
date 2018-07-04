unit PDFCreator_TLB;

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
// Fichier généré le 18/07/2011 10:56:09 depuis la bibliothèque de types ci-dessous.

// ************************************************************************  //
// Bibl. types : C:\Program Files\PDFCreator\PDFCreator.exe (1)
// LIBID: {1CE9DC08-9FBC-45C6-8A7C-4FE1E208A613}
// LCID: 0
// Fichier d'aide : 
// Chaîne d'aide : 
// DepndLst: 
//   (1) v2.0 stdole, (C:\Windows\system32\stdole2.tlb)
//   (2) v6.0 VBA, (C:\Windows\system32\MSVBVM60.DLL)
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

uses Windows, ActiveX, Classes, Graphics, OleServer, StdVCL, Variants, VBA_TLB;
  


// *********************************************************************//
// GUIDS déclarés dans la bibliothèque de types. Préfixes utilisés :    
//   Bibliothèques de types : LIBID_xxxx                                      
//   CoClasses              : CLASS_xxxx                                      
//   DISPInterfaces         : DIID_xxxx                                       
//   Non-DISP interfaces    : IID_xxxx                                        
// *********************************************************************//
const
  // Versions majeure et mineure de la bibliothèque de types
  PDFCreatorMajorVersion = 7;
  PDFCreatorMinorVersion = 1;

  LIBID_PDFCreator: TGUID = '{1CE9DC08-9FBC-45C6-8A7C-4FE1E208A613}';

  IID__clsPDFCreatorOptions: TGUID = '{F9A69DB6-B04F-4F31-9B15-3DEC9F1F2472}';
  CLASS_clsPDFCreatorOptions: TGUID = '{FCC886F6-E0DF-4302-8BE4-F8A8D9CB881C}';
  IID__clsPDFCreatorError: TGUID = '{A030F401-6045-4942-A5F5-9CCBF2C1872D}';
  CLASS_clsPDFCreatorError: TGUID = '{84D26557-2990-4B3E-A99F-C4DC1CB6C225}';
  IID__clsPDFCreatorInfoSpoolFile: TGUID = '{253F17D9-1678-4B0E-843E-A2D37C2C6B4E}';
  CLASS_clsPDFCreatorInfoSpoolFile: TGUID = '{411DBF4B-0B78-4C28-9997-ECD80CC371C4}';
  IID__clsPDFCreator: TGUID = '{280807DB-86BB-4FD1-B477-A7A6E285A3FE}';
  DIID___clsPDFCreator: TGUID = '{4B1FCFC1-EB7C-4180-B0B6-68EBA12FCF88}';
  IID__clsTools: TGUID = '{58CCA1EC-FC65-4750-A7E7-FFA191BD1042}';
  CLASS_clsTools: TGUID = '{718E1546-3738-4DC8-AA9E-A1BDE33237FE}';
  IID__clsUpdate: TGUID = '{6381BF10-F5C2-4577-B6CA-EDAF0287185D}';
  CLASS_clsUpdate: TGUID = '{D4E7A610-1FDE-4F38-B498-60926031591C}';
  CLASS_clsPDFCreator: TGUID = '{40108C54-9352-46C9-822C-027727352F00}';
type

// *********************************************************************//
// Déclaration Forward des types définis dans la bibliothèque de types    
// *********************************************************************//
  _clsPDFCreatorOptions = interface;
  _clsPDFCreatorOptionsDisp = dispinterface;
  _clsPDFCreatorError = interface;
  _clsPDFCreatorErrorDisp = dispinterface;
  _clsPDFCreatorInfoSpoolFile = interface;
  _clsPDFCreatorInfoSpoolFileDisp = dispinterface;
  _clsPDFCreator = interface;
  _clsPDFCreatorDisp = dispinterface;
  __clsPDFCreator = dispinterface;
  _clsTools = interface;
  _clsToolsDisp = dispinterface;
  _clsUpdate = interface;
  _clsUpdateDisp = dispinterface;

// *********************************************************************//
// Déclaration de CoClasses définies dans la bibliothèque de types 
// (REMARQUE: On affecte chaque CoClasse à son Interface par défaut)              
// *********************************************************************//
  clsPDFCreatorOptions = _clsPDFCreatorOptions;
  clsPDFCreatorError = _clsPDFCreatorError;
  clsPDFCreatorInfoSpoolFile = _clsPDFCreatorInfoSpoolFile;
  clsTools = _clsTools;
  clsUpdate = _clsUpdate;
  clsPDFCreator = _clsPDFCreator;


// *********************************************************************//
// Déclaration de structures, d'unions et d'alias.                        
// *********************************************************************//

  clsPDFCreatorOptions___v0 = _clsPDFCreatorOptions; 
  clsPDFCreatorOptions___v1 = _clsPDFCreatorOptions; 
  clsPDFCreator___v0 = _clsPDFCreator; 

// *********************************************************************//
// Interface   : _clsPDFCreatorOptions
// Indicateurs : (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID        : {F9A69DB6-B04F-4F31-9B15-3DEC9F1F2472}
// *********************************************************************//
  _clsPDFCreatorOptions = interface(IDispatch)
    ['{F9A69DB6-B04F-4F31-9B15-3DEC9F1F2472}']
    function Get_AdditionalGhostscriptParameters: WideString; safecall;
    procedure Set_AdditionalGhostscriptParameters(const AdditionalGhostscriptParameters: WideString); safecall;
    function Get_AdditionalGhostscriptSearchpath: WideString; safecall;
    procedure Set_AdditionalGhostscriptSearchpath(const AdditionalGhostscriptSearchpath: WideString); safecall;
    function Get_AddWindowsFontpath: Integer; safecall;
    procedure Set_AddWindowsFontpath(AddWindowsFontpath: Integer); safecall;
    function Get_AllowSpecialGSCharsInFilenames: Integer; safecall;
    procedure Set_AllowSpecialGSCharsInFilenames(AllowSpecialGSCharsInFilenames: Integer); safecall;
    function Get_AutosaveDirectory: WideString; safecall;
    procedure Set_AutosaveDirectory(const AutosaveDirectory: WideString); safecall;
    function Get_AutosaveFilename: WideString; safecall;
    procedure Set_AutosaveFilename(const AutosaveFilename: WideString); safecall;
    function Get_AutosaveFormat: Integer; safecall;
    procedure Set_AutosaveFormat(AutosaveFormat: Integer); safecall;
    function Get_AutosaveStartStandardProgram: Integer; safecall;
    procedure Set_AutosaveStartStandardProgram(AutosaveStartStandardProgram: Integer); safecall;
    function Get_BMPColorscount: Integer; safecall;
    procedure Set_BMPColorscount(BMPColorscount: Integer); safecall;
    function Get_BMPResolution: Integer; safecall;
    procedure Set_BMPResolution(BMPResolution: Integer); safecall;
    function Get_ClientComputerResolveIPAddress: Integer; safecall;
    procedure Set_ClientComputerResolveIPAddress(ClientComputerResolveIPAddress: Integer); safecall;
    function Get_Counter: Currency; safecall;
    procedure Set_Counter(Counter: Currency); safecall;
    function Get_DeviceHeightPoints: Double; safecall;
    procedure Set_DeviceHeightPoints(DeviceHeightPoints: Double); safecall;
    function Get_DeviceWidthPoints: Double; safecall;
    procedure Set_DeviceWidthPoints(DeviceWidthPoints: Double); safecall;
    function Get_DirectoryGhostscriptBinaries: WideString; safecall;
    procedure Set_DirectoryGhostscriptBinaries(const DirectoryGhostscriptBinaries: WideString); safecall;
    function Get_DirectoryGhostscriptFonts: WideString; safecall;
    procedure Set_DirectoryGhostscriptFonts(const DirectoryGhostscriptFonts: WideString); safecall;
    function Get_DirectoryGhostscriptLibraries: WideString; safecall;
    procedure Set_DirectoryGhostscriptLibraries(const DirectoryGhostscriptLibraries: WideString); safecall;
    function Get_DirectoryGhostscriptResource: WideString; safecall;
    procedure Set_DirectoryGhostscriptResource(const DirectoryGhostscriptResource: WideString); safecall;
    function Get_DisableEmail: Integer; safecall;
    procedure Set_DisableEmail(DisableEmail: Integer); safecall;
    function Get_DontUseDocumentSettings: Integer; safecall;
    procedure Set_DontUseDocumentSettings(DontUseDocumentSettings: Integer); safecall;
    function Get_EPSLanguageLevel: Integer; safecall;
    procedure Set_EPSLanguageLevel(EPSLanguageLevel: Integer); safecall;
    function Get_FilenameSubstitutions: WideString; safecall;
    procedure Set_FilenameSubstitutions(const FilenameSubstitutions: WideString); safecall;
    function Get_FilenameSubstitutionsOnlyInTitle: Integer; safecall;
    procedure Set_FilenameSubstitutionsOnlyInTitle(FilenameSubstitutionsOnlyInTitle: Integer); safecall;
    function Get_JPEGColorscount: Integer; safecall;
    procedure Set_JPEGColorscount(JPEGColorscount: Integer); safecall;
    function Get_JPEGQuality: Integer; safecall;
    procedure Set_JPEGQuality(JPEGQuality: Integer); safecall;
    function Get_JPEGResolution: Integer; safecall;
    procedure Set_JPEGResolution(JPEGResolution: Integer); safecall;
    function Get_Language: WideString; safecall;
    procedure Set_Language(const Language: WideString); safecall;
    function Get_LastSaveDirectory: WideString; safecall;
    procedure Set_LastSaveDirectory(const LastSaveDirectory: WideString); safecall;
    function Get_LastUpdateCheck: WideString; safecall;
    procedure Set_LastUpdateCheck(const LastUpdateCheck: WideString); safecall;
    function Get_Logging: Integer; safecall;
    procedure Set_Logging(Logging: Integer); safecall;
    function Get_LogLines: Integer; safecall;
    procedure Set_LogLines(LogLines: Integer); safecall;
    function Get_NoConfirmMessageSwitchingDefaultprinter: Integer; safecall;
    procedure Set_NoConfirmMessageSwitchingDefaultprinter(NoConfirmMessageSwitchingDefaultprinter: Integer); safecall;
    function Get_NoProcessingAtStartup: Integer; safecall;
    procedure Set_NoProcessingAtStartup(NoProcessingAtStartup: Integer); safecall;
    function Get_NoPSCheck: Integer; safecall;
    procedure Set_NoPSCheck(NoPSCheck: Integer); safecall;
    function Get_OnePagePerFile: Integer; safecall;
    procedure Set_OnePagePerFile(OnePagePerFile: Integer); safecall;
    function Get_OptionsDesign: Integer; safecall;
    procedure Set_OptionsDesign(OptionsDesign: Integer); safecall;
    function Get_OptionsEnabled: Integer; safecall;
    procedure Set_OptionsEnabled(OptionsEnabled: Integer); safecall;
    function Get_OptionsVisible: Integer; safecall;
    procedure Set_OptionsVisible(OptionsVisible: Integer); safecall;
    function Get_Papersize: WideString; safecall;
    procedure Set_Papersize(const Papersize: WideString); safecall;
    function Get_PCLColorsCount: Integer; safecall;
    procedure Set_PCLColorsCount(PCLColorsCount: Integer); safecall;
    function Get_PCLResolution: Integer; safecall;
    procedure Set_PCLResolution(PCLResolution: Integer); safecall;
    function Get_PCXColorscount: Integer; safecall;
    procedure Set_PCXColorscount(PCXColorscount: Integer); safecall;
    function Get_PCXResolution: Integer; safecall;
    procedure Set_PCXResolution(PCXResolution: Integer); safecall;
    function Get_PDFAes128Encryption: Integer; safecall;
    procedure Set_PDFAes128Encryption(PDFAes128Encryption: Integer); safecall;
    function Get_PDFAllowAssembly: Integer; safecall;
    procedure Set_PDFAllowAssembly(PDFAllowAssembly: Integer); safecall;
    function Get_PDFAllowDegradedPrinting: Integer; safecall;
    procedure Set_PDFAllowDegradedPrinting(PDFAllowDegradedPrinting: Integer); safecall;
    function Get_PDFAllowFillIn: Integer; safecall;
    procedure Set_PDFAllowFillIn(PDFAllowFillIn: Integer); safecall;
    function Get_PDFAllowScreenReaders: Integer; safecall;
    procedure Set_PDFAllowScreenReaders(PDFAllowScreenReaders: Integer); safecall;
    function Get_PDFColorsCMYKToRGB: Integer; safecall;
    procedure Set_PDFColorsCMYKToRGB(PDFColorsCMYKToRGB: Integer); safecall;
    function Get_PDFColorsColorModel: Integer; safecall;
    procedure Set_PDFColorsColorModel(PDFColorsColorModel: Integer); safecall;
    function Get_PDFColorsPreserveHalftone: Integer; safecall;
    procedure Set_PDFColorsPreserveHalftone(PDFColorsPreserveHalftone: Integer); safecall;
    function Get_PDFColorsPreserveOverprint: Integer; safecall;
    procedure Set_PDFColorsPreserveOverprint(PDFColorsPreserveOverprint: Integer); safecall;
    function Get_PDFColorsPreserveTransfer: Integer; safecall;
    procedure Set_PDFColorsPreserveTransfer(PDFColorsPreserveTransfer: Integer); safecall;
    function Get_PDFCompressionColorCompression: Integer; safecall;
    procedure Set_PDFCompressionColorCompression(PDFCompressionColorCompression: Integer); safecall;
    function Get_PDFCompressionColorCompressionChoice: Integer; safecall;
    procedure Set_PDFCompressionColorCompressionChoice(PDFCompressionColorCompressionChoice: Integer); safecall;
    function Get_PDFCompressionColorCompressionJPEGHighFactor: Double; safecall;
    procedure Set_PDFCompressionColorCompressionJPEGHighFactor(PDFCompressionColorCompressionJPEGHighFactor: Double); safecall;
    function Get_PDFCompressionColorCompressionJPEGLowFactor: Double; safecall;
    procedure Set_PDFCompressionColorCompressionJPEGLowFactor(PDFCompressionColorCompressionJPEGLowFactor: Double); safecall;
    function Get_PDFCompressionColorCompressionJPEGManualFactor: Double; safecall;
    procedure Set_PDFCompressionColorCompressionJPEGManualFactor(PDFCompressionColorCompressionJPEGManualFactor: Double); safecall;
    function Get_PDFCompressionColorCompressionJPEGMaximumFactor: Double; safecall;
    procedure Set_PDFCompressionColorCompressionJPEGMaximumFactor(PDFCompressionColorCompressionJPEGMaximumFactor: Double); safecall;
    function Get_PDFCompressionColorCompressionJPEGMediumFactor: Double; safecall;
    procedure Set_PDFCompressionColorCompressionJPEGMediumFactor(PDFCompressionColorCompressionJPEGMediumFactor: Double); safecall;
    function Get_PDFCompressionColorCompressionJPEGMinimumFactor: Double; safecall;
    procedure Set_PDFCompressionColorCompressionJPEGMinimumFactor(PDFCompressionColorCompressionJPEGMinimumFactor: Double); safecall;
    function Get_PDFCompressionColorResample: Integer; safecall;
    procedure Set_PDFCompressionColorResample(PDFCompressionColorResample: Integer); safecall;
    function Get_PDFCompressionColorResampleChoice: Integer; safecall;
    procedure Set_PDFCompressionColorResampleChoice(PDFCompressionColorResampleChoice: Integer); safecall;
    function Get_PDFCompressionColorResolution: Integer; safecall;
    procedure Set_PDFCompressionColorResolution(PDFCompressionColorResolution: Integer); safecall;
    function Get_PDFCompressionGreyCompression: Integer; safecall;
    procedure Set_PDFCompressionGreyCompression(PDFCompressionGreyCompression: Integer); safecall;
    function Get_PDFCompressionGreyCompressionChoice: Integer; safecall;
    procedure Set_PDFCompressionGreyCompressionChoice(PDFCompressionGreyCompressionChoice: Integer); safecall;
    function Get_PDFCompressionGreyCompressionJPEGHighFactor: Double; safecall;
    procedure Set_PDFCompressionGreyCompressionJPEGHighFactor(PDFCompressionGreyCompressionJPEGHighFactor: Double); safecall;
    function Get_PDFCompressionGreyCompressionJPEGLowFactor: Double; safecall;
    procedure Set_PDFCompressionGreyCompressionJPEGLowFactor(PDFCompressionGreyCompressionJPEGLowFactor: Double); safecall;
    function Get_PDFCompressionGreyCompressionJPEGManualFactor: Double; safecall;
    procedure Set_PDFCompressionGreyCompressionJPEGManualFactor(PDFCompressionGreyCompressionJPEGManualFactor: Double); safecall;
    function Get_PDFCompressionGreyCompressionJPEGMaximumFactor: Double; safecall;
    procedure Set_PDFCompressionGreyCompressionJPEGMaximumFactor(PDFCompressionGreyCompressionJPEGMaximumFactor: Double); safecall;
    function Get_PDFCompressionGreyCompressionJPEGMediumFactor: Double; safecall;
    procedure Set_PDFCompressionGreyCompressionJPEGMediumFactor(PDFCompressionGreyCompressionJPEGMediumFactor: Double); safecall;
    function Get_PDFCompressionGreyCompressionJPEGMinimumFactor: Double; safecall;
    procedure Set_PDFCompressionGreyCompressionJPEGMinimumFactor(PDFCompressionGreyCompressionJPEGMinimumFactor: Double); safecall;
    function Get_PDFCompressionGreyResample: Integer; safecall;
    procedure Set_PDFCompressionGreyResample(PDFCompressionGreyResample: Integer); safecall;
    function Get_PDFCompressionGreyResampleChoice: Integer; safecall;
    procedure Set_PDFCompressionGreyResampleChoice(PDFCompressionGreyResampleChoice: Integer); safecall;
    function Get_PDFCompressionGreyResolution: Integer; safecall;
    procedure Set_PDFCompressionGreyResolution(PDFCompressionGreyResolution: Integer); safecall;
    function Get_PDFCompressionMonoCompression: Integer; safecall;
    procedure Set_PDFCompressionMonoCompression(PDFCompressionMonoCompression: Integer); safecall;
    function Get_PDFCompressionMonoCompressionChoice: Integer; safecall;
    procedure Set_PDFCompressionMonoCompressionChoice(PDFCompressionMonoCompressionChoice: Integer); safecall;
    function Get_PDFCompressionMonoResample: Integer; safecall;
    procedure Set_PDFCompressionMonoResample(PDFCompressionMonoResample: Integer); safecall;
    function Get_PDFCompressionMonoResampleChoice: Integer; safecall;
    procedure Set_PDFCompressionMonoResampleChoice(PDFCompressionMonoResampleChoice: Integer); safecall;
    function Get_PDFCompressionMonoResolution: Integer; safecall;
    procedure Set_PDFCompressionMonoResolution(PDFCompressionMonoResolution: Integer); safecall;
    function Get_PDFCompressionTextCompression: Integer; safecall;
    procedure Set_PDFCompressionTextCompression(PDFCompressionTextCompression: Integer); safecall;
    function Get_PDFDisallowCopy: Integer; safecall;
    procedure Set_PDFDisallowCopy(PDFDisallowCopy: Integer); safecall;
    function Get_PDFDisallowModifyAnnotations: Integer; safecall;
    procedure Set_PDFDisallowModifyAnnotations(PDFDisallowModifyAnnotations: Integer); safecall;
    function Get_PDFDisallowModifyContents: Integer; safecall;
    procedure Set_PDFDisallowModifyContents(PDFDisallowModifyContents: Integer); safecall;
    function Get_PDFDisallowPrinting: Integer; safecall;
    procedure Set_PDFDisallowPrinting(PDFDisallowPrinting: Integer); safecall;
    function Get_PDFEncryptor: Integer; safecall;
    procedure Set_PDFEncryptor(PDFEncryptor: Integer); safecall;
    function Get_PDFFontsEmbedAll: Integer; safecall;
    procedure Set_PDFFontsEmbedAll(PDFFontsEmbedAll: Integer); safecall;
    function Get_PDFFontsSubSetFonts: Integer; safecall;
    procedure Set_PDFFontsSubSetFonts(PDFFontsSubSetFonts: Integer); safecall;
    function Get_PDFFontsSubSetFontsPercent: Integer; safecall;
    procedure Set_PDFFontsSubSetFontsPercent(PDFFontsSubSetFontsPercent: Integer); safecall;
    function Get_PDFGeneralASCII85: Integer; safecall;
    procedure Set_PDFGeneralASCII85(PDFGeneralASCII85: Integer); safecall;
    function Get_PDFGeneralAutorotate: Integer; safecall;
    procedure Set_PDFGeneralAutorotate(PDFGeneralAutorotate: Integer); safecall;
    function Get_PDFGeneralCompatibility: Integer; safecall;
    procedure Set_PDFGeneralCompatibility(PDFGeneralCompatibility: Integer); safecall;
    function Get_PDFGeneralDefault: Integer; safecall;
    procedure Set_PDFGeneralDefault(PDFGeneralDefault: Integer); safecall;
    function Get_PDFGeneralOverprint: Integer; safecall;
    procedure Set_PDFGeneralOverprint(PDFGeneralOverprint: Integer); safecall;
    function Get_PDFGeneralResolution: Integer; safecall;
    procedure Set_PDFGeneralResolution(PDFGeneralResolution: Integer); safecall;
    function Get_PDFHighEncryption: Integer; safecall;
    procedure Set_PDFHighEncryption(PDFHighEncryption: Integer); safecall;
    function Get_PDFLowEncryption: Integer; safecall;
    procedure Set_PDFLowEncryption(PDFLowEncryption: Integer); safecall;
    function Get_PDFOptimize: Integer; safecall;
    procedure Set_PDFOptimize(PDFOptimize: Integer); safecall;
    function Get_PDFOwnerPass: Integer; safecall;
    procedure Set_PDFOwnerPass(PDFOwnerPass: Integer); safecall;
    function Get_PDFOwnerPasswordString: WideString; safecall;
    procedure Set_PDFOwnerPasswordString(const PDFOwnerPasswordString: WideString); safecall;
    function Get_PDFSigningMultiSignature: Integer; safecall;
    procedure Set_PDFSigningMultiSignature(PDFSigningMultiSignature: Integer); safecall;
    function Get_PDFSigningPFXFile: WideString; safecall;
    procedure Set_PDFSigningPFXFile(const PDFSigningPFXFile: WideString); safecall;
    function Get_PDFSigningPFXFilePassword: WideString; safecall;
    procedure Set_PDFSigningPFXFilePassword(const PDFSigningPFXFilePassword: WideString); safecall;
    function Get_PDFSigningSignatureContact: WideString; safecall;
    procedure Set_PDFSigningSignatureContact(const PDFSigningSignatureContact: WideString); safecall;
    function Get_PDFSigningSignatureLeftX: Double; safecall;
    procedure Set_PDFSigningSignatureLeftX(PDFSigningSignatureLeftX: Double); safecall;
    function Get_PDFSigningSignatureLeftY: Double; safecall;
    procedure Set_PDFSigningSignatureLeftY(PDFSigningSignatureLeftY: Double); safecall;
    function Get_PDFSigningSignatureLocation: WideString; safecall;
    procedure Set_PDFSigningSignatureLocation(const PDFSigningSignatureLocation: WideString); safecall;
    function Get_PDFSigningSignatureOnPage: Integer; safecall;
    procedure Set_PDFSigningSignatureOnPage(PDFSigningSignatureOnPage: Integer); safecall;
    function Get_PDFSigningSignatureReason: WideString; safecall;
    procedure Set_PDFSigningSignatureReason(const PDFSigningSignatureReason: WideString); safecall;
    function Get_PDFSigningSignatureRightX: Double; safecall;
    procedure Set_PDFSigningSignatureRightX(PDFSigningSignatureRightX: Double); safecall;
    function Get_PDFSigningSignatureRightY: Double; safecall;
    procedure Set_PDFSigningSignatureRightY(PDFSigningSignatureRightY: Double); safecall;
    function Get_PDFSigningSignatureVisible: Integer; safecall;
    procedure Set_PDFSigningSignatureVisible(PDFSigningSignatureVisible: Integer); safecall;
    function Get_PDFSigningSignPDF: Integer; safecall;
    procedure Set_PDFSigningSignPDF(PDFSigningSignPDF: Integer); safecall;
    function Get_PDFUpdateMetadata: Integer; safecall;
    procedure Set_PDFUpdateMetadata(PDFUpdateMetadata: Integer); safecall;
    function Get_PDFUserPass: Integer; safecall;
    procedure Set_PDFUserPass(PDFUserPass: Integer); safecall;
    function Get_PDFUserPasswordString: WideString; safecall;
    procedure Set_PDFUserPasswordString(const PDFUserPasswordString: WideString); safecall;
    function Get_PDFUseSecurity: Integer; safecall;
    procedure Set_PDFUseSecurity(PDFUseSecurity: Integer); safecall;
    function Get_PNGColorscount: Integer; safecall;
    procedure Set_PNGColorscount(PNGColorscount: Integer); safecall;
    function Get_PNGResolution: Integer; safecall;
    procedure Set_PNGResolution(PNGResolution: Integer); safecall;
    function Get_PrintAfterSaving: Integer; safecall;
    procedure Set_PrintAfterSaving(PrintAfterSaving: Integer); safecall;
    function Get_PrintAfterSavingBitsPerPixel: Integer; safecall;
    procedure Set_PrintAfterSavingBitsPerPixel(PrintAfterSavingBitsPerPixel: Integer); safecall;
    function Get_PrintAfterSavingDuplex: Integer; safecall;
    procedure Set_PrintAfterSavingDuplex(PrintAfterSavingDuplex: Integer); safecall;
    function Get_PrintAfterSavingMaxResolution: Integer; safecall;
    procedure Set_PrintAfterSavingMaxResolution(PrintAfterSavingMaxResolution: Integer); safecall;
    function Get_PrintAfterSavingMaxResolutionEnabled: Integer; safecall;
    procedure Set_PrintAfterSavingMaxResolutionEnabled(PrintAfterSavingMaxResolutionEnabled: Integer); safecall;
    function Get_PrintAfterSavingNoCancel: Integer; safecall;
    procedure Set_PrintAfterSavingNoCancel(PrintAfterSavingNoCancel: Integer); safecall;
    function Get_PrintAfterSavingPrinter: WideString; safecall;
    procedure Set_PrintAfterSavingPrinter(const PrintAfterSavingPrinter: WideString); safecall;
    function Get_PrintAfterSavingQueryUser: Integer; safecall;
    procedure Set_PrintAfterSavingQueryUser(PrintAfterSavingQueryUser: Integer); safecall;
    function Get_PrintAfterSavingTumble: Integer; safecall;
    procedure Set_PrintAfterSavingTumble(PrintAfterSavingTumble: Integer); safecall;
    function Get_PrinterStop: Integer; safecall;
    procedure Set_PrinterStop(PrinterStop: Integer); safecall;
    function Get_PrinterTemppath: WideString; safecall;
    procedure Set_PrinterTemppath(const PrinterTemppath: WideString); safecall;
    function Get_ProcessPriority: Integer; safecall;
    procedure Set_ProcessPriority(ProcessPriority: Integer); safecall;
    function Get_ProgramFont: WideString; safecall;
    procedure Set_ProgramFont(const ProgramFont: WideString); safecall;
    function Get_ProgramFontCharset: Integer; safecall;
    procedure Set_ProgramFontCharset(ProgramFontCharset: Integer); safecall;
    function Get_ProgramFontSize: Integer; safecall;
    procedure Set_ProgramFontSize(ProgramFontSize: Integer); safecall;
    function Get_PSDColorsCount: Integer; safecall;
    procedure Set_PSDColorsCount(PSDColorsCount: Integer); safecall;
    function Get_PSDResolution: Integer; safecall;
    procedure Set_PSDResolution(PSDResolution: Integer); safecall;
    function Get_PSLanguageLevel: Integer; safecall;
    procedure Set_PSLanguageLevel(PSLanguageLevel: Integer); safecall;
    function Get_RAWColorsCount: Integer; safecall;
    procedure Set_RAWColorsCount(RAWColorsCount: Integer); safecall;
    function Get_RAWResolution: Integer; safecall;
    procedure Set_RAWResolution(RAWResolution: Integer); safecall;
    function Get_RemoveAllKnownFileExtensions: Integer; safecall;
    procedure Set_RemoveAllKnownFileExtensions(RemoveAllKnownFileExtensions: Integer); safecall;
    function Get_RemoveSpaces: Integer; safecall;
    procedure Set_RemoveSpaces(RemoveSpaces: Integer); safecall;
    function Get_RunProgramAfterSaving: Integer; safecall;
    procedure Set_RunProgramAfterSaving(RunProgramAfterSaving: Integer); safecall;
    function Get_RunProgramAfterSavingProgramname: WideString; safecall;
    procedure Set_RunProgramAfterSavingProgramname(const RunProgramAfterSavingProgramname: WideString); safecall;
    function Get_RunProgramAfterSavingProgramParameters: WideString; safecall;
    procedure Set_RunProgramAfterSavingProgramParameters(const RunProgramAfterSavingProgramParameters: WideString); safecall;
    function Get_RunProgramAfterSavingWaitUntilReady: Integer; safecall;
    procedure Set_RunProgramAfterSavingWaitUntilReady(RunProgramAfterSavingWaitUntilReady: Integer); safecall;
    function Get_RunProgramAfterSavingWindowstyle: Integer; safecall;
    procedure Set_RunProgramAfterSavingWindowstyle(RunProgramAfterSavingWindowstyle: Integer); safecall;
    function Get_RunProgramBeforeSaving: Integer; safecall;
    procedure Set_RunProgramBeforeSaving(RunProgramBeforeSaving: Integer); safecall;
    function Get_RunProgramBeforeSavingProgramname: WideString; safecall;
    procedure Set_RunProgramBeforeSavingProgramname(const RunProgramBeforeSavingProgramname: WideString); safecall;
    function Get_RunProgramBeforeSavingProgramParameters: WideString; safecall;
    procedure Set_RunProgramBeforeSavingProgramParameters(const RunProgramBeforeSavingProgramParameters: WideString); safecall;
    function Get_RunProgramBeforeSavingWindowstyle: Integer; safecall;
    procedure Set_RunProgramBeforeSavingWindowstyle(RunProgramBeforeSavingWindowstyle: Integer); safecall;
    function Get_SaveFilename: WideString; safecall;
    procedure Set_SaveFilename(const SaveFilename: WideString); safecall;
    function Get_SendEmailAfterAutoSaving: Integer; safecall;
    procedure Set_SendEmailAfterAutoSaving(SendEmailAfterAutoSaving: Integer); safecall;
    function Get_SendMailMethod: Integer; safecall;
    procedure Set_SendMailMethod(SendMailMethod: Integer); safecall;
    function Get_ShowAnimation: Integer; safecall;
    procedure Set_ShowAnimation(ShowAnimation: Integer); safecall;
    function Get_StampFontColor: WideString; safecall;
    procedure Set_StampFontColor(const StampFontColor: WideString); safecall;
    function Get_StampFontname: WideString; safecall;
    procedure Set_StampFontname(const StampFontname: WideString); safecall;
    function Get_StampFontsize: Integer; safecall;
    procedure Set_StampFontsize(StampFontsize: Integer); safecall;
    function Get_StampOutlineFontthickness: Integer; safecall;
    procedure Set_StampOutlineFontthickness(StampOutlineFontthickness: Integer); safecall;
    function Get_StampString: WideString; safecall;
    procedure Set_StampString(const StampString: WideString); safecall;
    function Get_StampUseOutlineFont: Integer; safecall;
    procedure Set_StampUseOutlineFont(StampUseOutlineFont: Integer); safecall;
    function Get_StandardAuthor: WideString; safecall;
    procedure Set_StandardAuthor(const StandardAuthor: WideString); safecall;
    function Get_StandardCreationdate: WideString; safecall;
    procedure Set_StandardCreationdate(const StandardCreationdate: WideString); safecall;
    function Get_StandardDateformat: WideString; safecall;
    procedure Set_StandardDateformat(const StandardDateformat: WideString); safecall;
    function Get_StandardKeywords: WideString; safecall;
    procedure Set_StandardKeywords(const StandardKeywords: WideString); safecall;
    function Get_StandardMailDomain: WideString; safecall;
    procedure Set_StandardMailDomain(const StandardMailDomain: WideString); safecall;
    function Get_StandardModifydate: WideString; safecall;
    procedure Set_StandardModifydate(const StandardModifydate: WideString); safecall;
    function Get_StandardSaveformat: Integer; safecall;
    procedure Set_StandardSaveformat(StandardSaveformat: Integer); safecall;
    function Get_StandardSubject: WideString; safecall;
    procedure Set_StandardSubject(const StandardSubject: WideString); safecall;
    function Get_StandardTitle: WideString; safecall;
    procedure Set_StandardTitle(const StandardTitle: WideString); safecall;
    function Get_StartStandardProgram: Integer; safecall;
    procedure Set_StartStandardProgram(StartStandardProgram: Integer); safecall;
    function Get_SVGResolution: Integer; safecall;
    procedure Set_SVGResolution(SVGResolution: Integer); safecall;
    function Get_TIFFColorscount: Integer; safecall;
    procedure Set_TIFFColorscount(TIFFColorscount: Integer); safecall;
    function Get_TIFFResolution: Integer; safecall;
    procedure Set_TIFFResolution(TIFFResolution: Integer); safecall;
    function Get_Toolbars: Integer; safecall;
    procedure Set_Toolbars(Toolbars: Integer); safecall;
    function Get_UpdateInterval: Integer; safecall;
    procedure Set_UpdateInterval(UpdateInterval: Integer); safecall;
    function Get_UseAutosave: Integer; safecall;
    procedure Set_UseAutosave(UseAutosave: Integer); safecall;
    function Get_UseAutosaveDirectory: Integer; safecall;
    procedure Set_UseAutosaveDirectory(UseAutosaveDirectory: Integer); safecall;
    function Get_UseCreationDateNow: Integer; safecall;
    procedure Set_UseCreationDateNow(UseCreationDateNow: Integer); safecall;
    function Get_UseCustomPaperSize: WideString; safecall;
    procedure Set_UseCustomPaperSize(const UseCustomPaperSize: WideString); safecall;
    function Get_UseFixPapersize: Integer; safecall;
    procedure Set_UseFixPapersize(UseFixPapersize: Integer); safecall;
    function Get_UseStandardAuthor: Integer; safecall;
    procedure Set_UseStandardAuthor(UseStandardAuthor: Integer); safecall;
    property AdditionalGhostscriptParameters: WideString read Get_AdditionalGhostscriptParameters write Set_AdditionalGhostscriptParameters;
    property AdditionalGhostscriptSearchpath: WideString read Get_AdditionalGhostscriptSearchpath write Set_AdditionalGhostscriptSearchpath;
    property AddWindowsFontpath: Integer read Get_AddWindowsFontpath write Set_AddWindowsFontpath;
    property AllowSpecialGSCharsInFilenames: Integer read Get_AllowSpecialGSCharsInFilenames write Set_AllowSpecialGSCharsInFilenames;
    property AutosaveDirectory: WideString read Get_AutosaveDirectory write Set_AutosaveDirectory;
    property AutosaveFilename: WideString read Get_AutosaveFilename write Set_AutosaveFilename;
    property AutosaveFormat: Integer read Get_AutosaveFormat write Set_AutosaveFormat;
    property AutosaveStartStandardProgram: Integer read Get_AutosaveStartStandardProgram write Set_AutosaveStartStandardProgram;
    property BMPColorscount: Integer read Get_BMPColorscount write Set_BMPColorscount;
    property BMPResolution: Integer read Get_BMPResolution write Set_BMPResolution;
    property ClientComputerResolveIPAddress: Integer read Get_ClientComputerResolveIPAddress write Set_ClientComputerResolveIPAddress;
    property Counter: Currency read Get_Counter write Set_Counter;
    property DeviceHeightPoints: Double read Get_DeviceHeightPoints write Set_DeviceHeightPoints;
    property DeviceWidthPoints: Double read Get_DeviceWidthPoints write Set_DeviceWidthPoints;
    property DirectoryGhostscriptBinaries: WideString read Get_DirectoryGhostscriptBinaries write Set_DirectoryGhostscriptBinaries;
    property DirectoryGhostscriptFonts: WideString read Get_DirectoryGhostscriptFonts write Set_DirectoryGhostscriptFonts;
    property DirectoryGhostscriptLibraries: WideString read Get_DirectoryGhostscriptLibraries write Set_DirectoryGhostscriptLibraries;
    property DirectoryGhostscriptResource: WideString read Get_DirectoryGhostscriptResource write Set_DirectoryGhostscriptResource;
    property DisableEmail: Integer read Get_DisableEmail write Set_DisableEmail;
    property DontUseDocumentSettings: Integer read Get_DontUseDocumentSettings write Set_DontUseDocumentSettings;
    property EPSLanguageLevel: Integer read Get_EPSLanguageLevel write Set_EPSLanguageLevel;
    property FilenameSubstitutions: WideString read Get_FilenameSubstitutions write Set_FilenameSubstitutions;
    property FilenameSubstitutionsOnlyInTitle: Integer read Get_FilenameSubstitutionsOnlyInTitle write Set_FilenameSubstitutionsOnlyInTitle;
    property JPEGColorscount: Integer read Get_JPEGColorscount write Set_JPEGColorscount;
    property JPEGQuality: Integer read Get_JPEGQuality write Set_JPEGQuality;
    property JPEGResolution: Integer read Get_JPEGResolution write Set_JPEGResolution;
    property Language: WideString read Get_Language write Set_Language;
    property LastSaveDirectory: WideString read Get_LastSaveDirectory write Set_LastSaveDirectory;
    property LastUpdateCheck: WideString read Get_LastUpdateCheck write Set_LastUpdateCheck;
    property Logging: Integer read Get_Logging write Set_Logging;
    property LogLines: Integer read Get_LogLines write Set_LogLines;
    property NoConfirmMessageSwitchingDefaultprinter: Integer read Get_NoConfirmMessageSwitchingDefaultprinter write Set_NoConfirmMessageSwitchingDefaultprinter;
    property NoProcessingAtStartup: Integer read Get_NoProcessingAtStartup write Set_NoProcessingAtStartup;
    property NoPSCheck: Integer read Get_NoPSCheck write Set_NoPSCheck;
    property OnePagePerFile: Integer read Get_OnePagePerFile write Set_OnePagePerFile;
    property OptionsDesign: Integer read Get_OptionsDesign write Set_OptionsDesign;
    property OptionsEnabled: Integer read Get_OptionsEnabled write Set_OptionsEnabled;
    property OptionsVisible: Integer read Get_OptionsVisible write Set_OptionsVisible;
    property Papersize: WideString read Get_Papersize write Set_Papersize;
    property PCLColorsCount: Integer read Get_PCLColorsCount write Set_PCLColorsCount;
    property PCLResolution: Integer read Get_PCLResolution write Set_PCLResolution;
    property PCXColorscount: Integer read Get_PCXColorscount write Set_PCXColorscount;
    property PCXResolution: Integer read Get_PCXResolution write Set_PCXResolution;
    property PDFAes128Encryption: Integer read Get_PDFAes128Encryption write Set_PDFAes128Encryption;
    property PDFAllowAssembly: Integer read Get_PDFAllowAssembly write Set_PDFAllowAssembly;
    property PDFAllowDegradedPrinting: Integer read Get_PDFAllowDegradedPrinting write Set_PDFAllowDegradedPrinting;
    property PDFAllowFillIn: Integer read Get_PDFAllowFillIn write Set_PDFAllowFillIn;
    property PDFAllowScreenReaders: Integer read Get_PDFAllowScreenReaders write Set_PDFAllowScreenReaders;
    property PDFColorsCMYKToRGB: Integer read Get_PDFColorsCMYKToRGB write Set_PDFColorsCMYKToRGB;
    property PDFColorsColorModel: Integer read Get_PDFColorsColorModel write Set_PDFColorsColorModel;
    property PDFColorsPreserveHalftone: Integer read Get_PDFColorsPreserveHalftone write Set_PDFColorsPreserveHalftone;
    property PDFColorsPreserveOverprint: Integer read Get_PDFColorsPreserveOverprint write Set_PDFColorsPreserveOverprint;
    property PDFColorsPreserveTransfer: Integer read Get_PDFColorsPreserveTransfer write Set_PDFColorsPreserveTransfer;
    property PDFCompressionColorCompression: Integer read Get_PDFCompressionColorCompression write Set_PDFCompressionColorCompression;
    property PDFCompressionColorCompressionChoice: Integer read Get_PDFCompressionColorCompressionChoice write Set_PDFCompressionColorCompressionChoice;
    property PDFCompressionColorCompressionJPEGHighFactor: Double read Get_PDFCompressionColorCompressionJPEGHighFactor write Set_PDFCompressionColorCompressionJPEGHighFactor;
    property PDFCompressionColorCompressionJPEGLowFactor: Double read Get_PDFCompressionColorCompressionJPEGLowFactor write Set_PDFCompressionColorCompressionJPEGLowFactor;
    property PDFCompressionColorCompressionJPEGManualFactor: Double read Get_PDFCompressionColorCompressionJPEGManualFactor write Set_PDFCompressionColorCompressionJPEGManualFactor;
    property PDFCompressionColorCompressionJPEGMaximumFactor: Double read Get_PDFCompressionColorCompressionJPEGMaximumFactor write Set_PDFCompressionColorCompressionJPEGMaximumFactor;
    property PDFCompressionColorCompressionJPEGMediumFactor: Double read Get_PDFCompressionColorCompressionJPEGMediumFactor write Set_PDFCompressionColorCompressionJPEGMediumFactor;
    property PDFCompressionColorCompressionJPEGMinimumFactor: Double read Get_PDFCompressionColorCompressionJPEGMinimumFactor write Set_PDFCompressionColorCompressionJPEGMinimumFactor;
    property PDFCompressionColorResample: Integer read Get_PDFCompressionColorResample write Set_PDFCompressionColorResample;
    property PDFCompressionColorResampleChoice: Integer read Get_PDFCompressionColorResampleChoice write Set_PDFCompressionColorResampleChoice;
    property PDFCompressionColorResolution: Integer read Get_PDFCompressionColorResolution write Set_PDFCompressionColorResolution;
    property PDFCompressionGreyCompression: Integer read Get_PDFCompressionGreyCompression write Set_PDFCompressionGreyCompression;
    property PDFCompressionGreyCompressionChoice: Integer read Get_PDFCompressionGreyCompressionChoice write Set_PDFCompressionGreyCompressionChoice;
    property PDFCompressionGreyCompressionJPEGHighFactor: Double read Get_PDFCompressionGreyCompressionJPEGHighFactor write Set_PDFCompressionGreyCompressionJPEGHighFactor;
    property PDFCompressionGreyCompressionJPEGLowFactor: Double read Get_PDFCompressionGreyCompressionJPEGLowFactor write Set_PDFCompressionGreyCompressionJPEGLowFactor;
    property PDFCompressionGreyCompressionJPEGManualFactor: Double read Get_PDFCompressionGreyCompressionJPEGManualFactor write Set_PDFCompressionGreyCompressionJPEGManualFactor;
    property PDFCompressionGreyCompressionJPEGMaximumFactor: Double read Get_PDFCompressionGreyCompressionJPEGMaximumFactor write Set_PDFCompressionGreyCompressionJPEGMaximumFactor;
    property PDFCompressionGreyCompressionJPEGMediumFactor: Double read Get_PDFCompressionGreyCompressionJPEGMediumFactor write Set_PDFCompressionGreyCompressionJPEGMediumFactor;
    property PDFCompressionGreyCompressionJPEGMinimumFactor: Double read Get_PDFCompressionGreyCompressionJPEGMinimumFactor write Set_PDFCompressionGreyCompressionJPEGMinimumFactor;
    property PDFCompressionGreyResample: Integer read Get_PDFCompressionGreyResample write Set_PDFCompressionGreyResample;
    property PDFCompressionGreyResampleChoice: Integer read Get_PDFCompressionGreyResampleChoice write Set_PDFCompressionGreyResampleChoice;
    property PDFCompressionGreyResolution: Integer read Get_PDFCompressionGreyResolution write Set_PDFCompressionGreyResolution;
    property PDFCompressionMonoCompression: Integer read Get_PDFCompressionMonoCompression write Set_PDFCompressionMonoCompression;
    property PDFCompressionMonoCompressionChoice: Integer read Get_PDFCompressionMonoCompressionChoice write Set_PDFCompressionMonoCompressionChoice;
    property PDFCompressionMonoResample: Integer read Get_PDFCompressionMonoResample write Set_PDFCompressionMonoResample;
    property PDFCompressionMonoResampleChoice: Integer read Get_PDFCompressionMonoResampleChoice write Set_PDFCompressionMonoResampleChoice;
    property PDFCompressionMonoResolution: Integer read Get_PDFCompressionMonoResolution write Set_PDFCompressionMonoResolution;
    property PDFCompressionTextCompression: Integer read Get_PDFCompressionTextCompression write Set_PDFCompressionTextCompression;
    property PDFDisallowCopy: Integer read Get_PDFDisallowCopy write Set_PDFDisallowCopy;
    property PDFDisallowModifyAnnotations: Integer read Get_PDFDisallowModifyAnnotations write Set_PDFDisallowModifyAnnotations;
    property PDFDisallowModifyContents: Integer read Get_PDFDisallowModifyContents write Set_PDFDisallowModifyContents;
    property PDFDisallowPrinting: Integer read Get_PDFDisallowPrinting write Set_PDFDisallowPrinting;
    property PDFEncryptor: Integer read Get_PDFEncryptor write Set_PDFEncryptor;
    property PDFFontsEmbedAll: Integer read Get_PDFFontsEmbedAll write Set_PDFFontsEmbedAll;
    property PDFFontsSubSetFonts: Integer read Get_PDFFontsSubSetFonts write Set_PDFFontsSubSetFonts;
    property PDFFontsSubSetFontsPercent: Integer read Get_PDFFontsSubSetFontsPercent write Set_PDFFontsSubSetFontsPercent;
    property PDFGeneralASCII85: Integer read Get_PDFGeneralASCII85 write Set_PDFGeneralASCII85;
    property PDFGeneralAutorotate: Integer read Get_PDFGeneralAutorotate write Set_PDFGeneralAutorotate;
    property PDFGeneralCompatibility: Integer read Get_PDFGeneralCompatibility write Set_PDFGeneralCompatibility;
    property PDFGeneralDefault: Integer read Get_PDFGeneralDefault write Set_PDFGeneralDefault;
    property PDFGeneralOverprint: Integer read Get_PDFGeneralOverprint write Set_PDFGeneralOverprint;
    property PDFGeneralResolution: Integer read Get_PDFGeneralResolution write Set_PDFGeneralResolution;
    property PDFHighEncryption: Integer read Get_PDFHighEncryption write Set_PDFHighEncryption;
    property PDFLowEncryption: Integer read Get_PDFLowEncryption write Set_PDFLowEncryption;
    property PDFOptimize: Integer read Get_PDFOptimize write Set_PDFOptimize;
    property PDFOwnerPass: Integer read Get_PDFOwnerPass write Set_PDFOwnerPass;
    property PDFOwnerPasswordString: WideString read Get_PDFOwnerPasswordString write Set_PDFOwnerPasswordString;
    property PDFSigningMultiSignature: Integer read Get_PDFSigningMultiSignature write Set_PDFSigningMultiSignature;
    property PDFSigningPFXFile: WideString read Get_PDFSigningPFXFile write Set_PDFSigningPFXFile;
    property PDFSigningPFXFilePassword: WideString read Get_PDFSigningPFXFilePassword write Set_PDFSigningPFXFilePassword;
    property PDFSigningSignatureContact: WideString read Get_PDFSigningSignatureContact write Set_PDFSigningSignatureContact;
    property PDFSigningSignatureLeftX: Double read Get_PDFSigningSignatureLeftX write Set_PDFSigningSignatureLeftX;
    property PDFSigningSignatureLeftY: Double read Get_PDFSigningSignatureLeftY write Set_PDFSigningSignatureLeftY;
    property PDFSigningSignatureLocation: WideString read Get_PDFSigningSignatureLocation write Set_PDFSigningSignatureLocation;
    property PDFSigningSignatureOnPage: Integer read Get_PDFSigningSignatureOnPage write Set_PDFSigningSignatureOnPage;
    property PDFSigningSignatureReason: WideString read Get_PDFSigningSignatureReason write Set_PDFSigningSignatureReason;
    property PDFSigningSignatureRightX: Double read Get_PDFSigningSignatureRightX write Set_PDFSigningSignatureRightX;
    property PDFSigningSignatureRightY: Double read Get_PDFSigningSignatureRightY write Set_PDFSigningSignatureRightY;
    property PDFSigningSignatureVisible: Integer read Get_PDFSigningSignatureVisible write Set_PDFSigningSignatureVisible;
    property PDFSigningSignPDF: Integer read Get_PDFSigningSignPDF write Set_PDFSigningSignPDF;
    property PDFUpdateMetadata: Integer read Get_PDFUpdateMetadata write Set_PDFUpdateMetadata;
    property PDFUserPass: Integer read Get_PDFUserPass write Set_PDFUserPass;
    property PDFUserPasswordString: WideString read Get_PDFUserPasswordString write Set_PDFUserPasswordString;
    property PDFUseSecurity: Integer read Get_PDFUseSecurity write Set_PDFUseSecurity;
    property PNGColorscount: Integer read Get_PNGColorscount write Set_PNGColorscount;
    property PNGResolution: Integer read Get_PNGResolution write Set_PNGResolution;
    property PrintAfterSaving: Integer read Get_PrintAfterSaving write Set_PrintAfterSaving;
    property PrintAfterSavingBitsPerPixel: Integer read Get_PrintAfterSavingBitsPerPixel write Set_PrintAfterSavingBitsPerPixel;
    property PrintAfterSavingDuplex: Integer read Get_PrintAfterSavingDuplex write Set_PrintAfterSavingDuplex;
    property PrintAfterSavingMaxResolution: Integer read Get_PrintAfterSavingMaxResolution write Set_PrintAfterSavingMaxResolution;
    property PrintAfterSavingMaxResolutionEnabled: Integer read Get_PrintAfterSavingMaxResolutionEnabled write Set_PrintAfterSavingMaxResolutionEnabled;
    property PrintAfterSavingNoCancel: Integer read Get_PrintAfterSavingNoCancel write Set_PrintAfterSavingNoCancel;
    property PrintAfterSavingPrinter: WideString read Get_PrintAfterSavingPrinter write Set_PrintAfterSavingPrinter;
    property PrintAfterSavingQueryUser: Integer read Get_PrintAfterSavingQueryUser write Set_PrintAfterSavingQueryUser;
    property PrintAfterSavingTumble: Integer read Get_PrintAfterSavingTumble write Set_PrintAfterSavingTumble;
    property PrinterStop: Integer read Get_PrinterStop write Set_PrinterStop;
    property PrinterTemppath: WideString read Get_PrinterTemppath write Set_PrinterTemppath;
    property ProcessPriority: Integer read Get_ProcessPriority write Set_ProcessPriority;
    property ProgramFont: WideString read Get_ProgramFont write Set_ProgramFont;
    property ProgramFontCharset: Integer read Get_ProgramFontCharset write Set_ProgramFontCharset;
    property ProgramFontSize: Integer read Get_ProgramFontSize write Set_ProgramFontSize;
    property PSDColorsCount: Integer read Get_PSDColorsCount write Set_PSDColorsCount;
    property PSDResolution: Integer read Get_PSDResolution write Set_PSDResolution;
    property PSLanguageLevel: Integer read Get_PSLanguageLevel write Set_PSLanguageLevel;
    property RAWColorsCount: Integer read Get_RAWColorsCount write Set_RAWColorsCount;
    property RAWResolution: Integer read Get_RAWResolution write Set_RAWResolution;
    property RemoveAllKnownFileExtensions: Integer read Get_RemoveAllKnownFileExtensions write Set_RemoveAllKnownFileExtensions;
    property RemoveSpaces: Integer read Get_RemoveSpaces write Set_RemoveSpaces;
    property RunProgramAfterSaving: Integer read Get_RunProgramAfterSaving write Set_RunProgramAfterSaving;
    property RunProgramAfterSavingProgramname: WideString read Get_RunProgramAfterSavingProgramname write Set_RunProgramAfterSavingProgramname;
    property RunProgramAfterSavingProgramParameters: WideString read Get_RunProgramAfterSavingProgramParameters write Set_RunProgramAfterSavingProgramParameters;
    property RunProgramAfterSavingWaitUntilReady: Integer read Get_RunProgramAfterSavingWaitUntilReady write Set_RunProgramAfterSavingWaitUntilReady;
    property RunProgramAfterSavingWindowstyle: Integer read Get_RunProgramAfterSavingWindowstyle write Set_RunProgramAfterSavingWindowstyle;
    property RunProgramBeforeSaving: Integer read Get_RunProgramBeforeSaving write Set_RunProgramBeforeSaving;
    property RunProgramBeforeSavingProgramname: WideString read Get_RunProgramBeforeSavingProgramname write Set_RunProgramBeforeSavingProgramname;
    property RunProgramBeforeSavingProgramParameters: WideString read Get_RunProgramBeforeSavingProgramParameters write Set_RunProgramBeforeSavingProgramParameters;
    property RunProgramBeforeSavingWindowstyle: Integer read Get_RunProgramBeforeSavingWindowstyle write Set_RunProgramBeforeSavingWindowstyle;
    property SaveFilename: WideString read Get_SaveFilename write Set_SaveFilename;
    property SendEmailAfterAutoSaving: Integer read Get_SendEmailAfterAutoSaving write Set_SendEmailAfterAutoSaving;
    property SendMailMethod: Integer read Get_SendMailMethod write Set_SendMailMethod;
    property ShowAnimation: Integer read Get_ShowAnimation write Set_ShowAnimation;
    property StampFontColor: WideString read Get_StampFontColor write Set_StampFontColor;
    property StampFontname: WideString read Get_StampFontname write Set_StampFontname;
    property StampFontsize: Integer read Get_StampFontsize write Set_StampFontsize;
    property StampOutlineFontthickness: Integer read Get_StampOutlineFontthickness write Set_StampOutlineFontthickness;
    property StampString: WideString read Get_StampString write Set_StampString;
    property StampUseOutlineFont: Integer read Get_StampUseOutlineFont write Set_StampUseOutlineFont;
    property StandardAuthor: WideString read Get_StandardAuthor write Set_StandardAuthor;
    property StandardCreationdate: WideString read Get_StandardCreationdate write Set_StandardCreationdate;
    property StandardDateformat: WideString read Get_StandardDateformat write Set_StandardDateformat;
    property StandardKeywords: WideString read Get_StandardKeywords write Set_StandardKeywords;
    property StandardMailDomain: WideString read Get_StandardMailDomain write Set_StandardMailDomain;
    property StandardModifydate: WideString read Get_StandardModifydate write Set_StandardModifydate;
    property StandardSaveformat: Integer read Get_StandardSaveformat write Set_StandardSaveformat;
    property StandardSubject: WideString read Get_StandardSubject write Set_StandardSubject;
    property StandardTitle: WideString read Get_StandardTitle write Set_StandardTitle;
    property StartStandardProgram: Integer read Get_StartStandardProgram write Set_StartStandardProgram;
    property SVGResolution: Integer read Get_SVGResolution write Set_SVGResolution;
    property TIFFColorscount: Integer read Get_TIFFColorscount write Set_TIFFColorscount;
    property TIFFResolution: Integer read Get_TIFFResolution write Set_TIFFResolution;
    property Toolbars: Integer read Get_Toolbars write Set_Toolbars;
    property UpdateInterval: Integer read Get_UpdateInterval write Set_UpdateInterval;
    property UseAutosave: Integer read Get_UseAutosave write Set_UseAutosave;
    property UseAutosaveDirectory: Integer read Get_UseAutosaveDirectory write Set_UseAutosaveDirectory;
    property UseCreationDateNow: Integer read Get_UseCreationDateNow write Set_UseCreationDateNow;
    property UseCustomPaperSize: WideString read Get_UseCustomPaperSize write Set_UseCustomPaperSize;
    property UseFixPapersize: Integer read Get_UseFixPapersize write Set_UseFixPapersize;
    property UseStandardAuthor: Integer read Get_UseStandardAuthor write Set_UseStandardAuthor;
  end;

// *********************************************************************//
// DispIntf :  _clsPDFCreatorOptionsDisp
// Flags :     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID :      {F9A69DB6-B04F-4F31-9B15-3DEC9F1F2472}
// *********************************************************************//
  _clsPDFCreatorOptionsDisp = dispinterface
    ['{F9A69DB6-B04F-4F31-9B15-3DEC9F1F2472}']
    property AdditionalGhostscriptParameters: WideString dispid 1073938432;
    property AdditionalGhostscriptSearchpath: WideString dispid 1073938433;
    property AddWindowsFontpath: Integer dispid 1073938434;
    property AllowSpecialGSCharsInFilenames: Integer dispid 1073938602;
    property AutosaveDirectory: WideString dispid 1073938435;
    property AutosaveFilename: WideString dispid 1073938436;
    property AutosaveFormat: Integer dispid 1073938437;
    property AutosaveStartStandardProgram: Integer dispid 1073938438;
    property BMPColorscount: Integer dispid 1073938439;
    property BMPResolution: Integer dispid 1073938440;
    property ClientComputerResolveIPAddress: Integer dispid 1073938441;
    property Counter: Currency dispid 1073938442;
    property DeviceHeightPoints: Double dispid 1073938443;
    property DeviceWidthPoints: Double dispid 1073938444;
    property DirectoryGhostscriptBinaries: WideString dispid 1073938445;
    property DirectoryGhostscriptFonts: WideString dispid 1073938446;
    property DirectoryGhostscriptLibraries: WideString dispid 1073938447;
    property DirectoryGhostscriptResource: WideString dispid 1073938448;
    property DisableEmail: Integer dispid 1073938449;
    property DontUseDocumentSettings: Integer dispid 1073938450;
    property EPSLanguageLevel: Integer dispid 1073938451;
    property FilenameSubstitutions: WideString dispid 1073938452;
    property FilenameSubstitutionsOnlyInTitle: Integer dispid 1073938453;
    property JPEGColorscount: Integer dispid 1073938454;
    property JPEGQuality: Integer dispid 1073938455;
    property JPEGResolution: Integer dispid 1073938456;
    property Language: WideString dispid 1073938457;
    property LastSaveDirectory: WideString dispid 1073938458;
    property LastUpdateCheck: WideString dispid 1073938603;
    property Logging: Integer dispid 1073938459;
    property LogLines: Integer dispid 1073938460;
    property NoConfirmMessageSwitchingDefaultprinter: Integer dispid 1073938461;
    property NoProcessingAtStartup: Integer dispid 1073938462;
    property NoPSCheck: Integer dispid 1073938463;
    property OnePagePerFile: Integer dispid 1073938464;
    property OptionsDesign: Integer dispid 1073938465;
    property OptionsEnabled: Integer dispid 1073938466;
    property OptionsVisible: Integer dispid 1073938467;
    property Papersize: WideString dispid 1073938468;
    property PCLColorsCount: Integer dispid 1073938469;
    property PCLResolution: Integer dispid 1073938470;
    property PCXColorscount: Integer dispid 1073938471;
    property PCXResolution: Integer dispid 1073938472;
    property PDFAes128Encryption: Integer dispid 1073938606;
    property PDFAllowAssembly: Integer dispid 1073938473;
    property PDFAllowDegradedPrinting: Integer dispid 1073938474;
    property PDFAllowFillIn: Integer dispid 1073938475;
    property PDFAllowScreenReaders: Integer dispid 1073938476;
    property PDFColorsCMYKToRGB: Integer dispid 1073938477;
    property PDFColorsColorModel: Integer dispid 1073938478;
    property PDFColorsPreserveHalftone: Integer dispid 1073938479;
    property PDFColorsPreserveOverprint: Integer dispid 1073938480;
    property PDFColorsPreserveTransfer: Integer dispid 1073938481;
    property PDFCompressionColorCompression: Integer dispid 1073938482;
    property PDFCompressionColorCompressionChoice: Integer dispid 1073938483;
    property PDFCompressionColorCompressionJPEGHighFactor: Double dispid 1073938484;
    property PDFCompressionColorCompressionJPEGLowFactor: Double dispid 1073938485;
    property PDFCompressionColorCompressionJPEGManualFactor: Double dispid 1073938607;
    property PDFCompressionColorCompressionJPEGMaximumFactor: Double dispid 1073938486;
    property PDFCompressionColorCompressionJPEGMediumFactor: Double dispid 1073938487;
    property PDFCompressionColorCompressionJPEGMinimumFactor: Double dispid 1073938488;
    property PDFCompressionColorResample: Integer dispid 1073938489;
    property PDFCompressionColorResampleChoice: Integer dispid 1073938490;
    property PDFCompressionColorResolution: Integer dispid 1073938491;
    property PDFCompressionGreyCompression: Integer dispid 1073938492;
    property PDFCompressionGreyCompressionChoice: Integer dispid 1073938493;
    property PDFCompressionGreyCompressionJPEGHighFactor: Double dispid 1073938494;
    property PDFCompressionGreyCompressionJPEGLowFactor: Double dispid 1073938495;
    property PDFCompressionGreyCompressionJPEGManualFactor: Double dispid 1073938608;
    property PDFCompressionGreyCompressionJPEGMaximumFactor: Double dispid 1073938496;
    property PDFCompressionGreyCompressionJPEGMediumFactor: Double dispid 1073938497;
    property PDFCompressionGreyCompressionJPEGMinimumFactor: Double dispid 1073938498;
    property PDFCompressionGreyResample: Integer dispid 1073938499;
    property PDFCompressionGreyResampleChoice: Integer dispid 1073938500;
    property PDFCompressionGreyResolution: Integer dispid 1073938501;
    property PDFCompressionMonoCompression: Integer dispid 1073938502;
    property PDFCompressionMonoCompressionChoice: Integer dispid 1073938503;
    property PDFCompressionMonoResample: Integer dispid 1073938504;
    property PDFCompressionMonoResampleChoice: Integer dispid 1073938505;
    property PDFCompressionMonoResolution: Integer dispid 1073938506;
    property PDFCompressionTextCompression: Integer dispid 1073938507;
    property PDFDisallowCopy: Integer dispid 1073938508;
    property PDFDisallowModifyAnnotations: Integer dispid 1073938509;
    property PDFDisallowModifyContents: Integer dispid 1073938510;
    property PDFDisallowPrinting: Integer dispid 1073938511;
    property PDFEncryptor: Integer dispid 1073938512;
    property PDFFontsEmbedAll: Integer dispid 1073938513;
    property PDFFontsSubSetFonts: Integer dispid 1073938514;
    property PDFFontsSubSetFontsPercent: Integer dispid 1073938515;
    property PDFGeneralASCII85: Integer dispid 1073938516;
    property PDFGeneralAutorotate: Integer dispid 1073938517;
    property PDFGeneralCompatibility: Integer dispid 1073938518;
    property PDFGeneralDefault: Integer dispid 1073938519;
    property PDFGeneralOverprint: Integer dispid 1073938520;
    property PDFGeneralResolution: Integer dispid 1073938521;
    property PDFHighEncryption: Integer dispid 1073938522;
    property PDFLowEncryption: Integer dispid 1073938523;
    property PDFOptimize: Integer dispid 1073938524;
    property PDFOwnerPass: Integer dispid 1073938525;
    property PDFOwnerPasswordString: WideString dispid 1073938526;
    property PDFSigningMultiSignature: Integer dispid 1073938527;
    property PDFSigningPFXFile: WideString dispid 1073938528;
    property PDFSigningPFXFilePassword: WideString dispid 1073938529;
    property PDFSigningSignatureContact: WideString dispid 1073938530;
    property PDFSigningSignatureLeftX: Double dispid 1073938531;
    property PDFSigningSignatureLeftY: Double dispid 1073938532;
    property PDFSigningSignatureLocation: WideString dispid 1073938533;
    property PDFSigningSignatureOnPage: Integer dispid 1073938609;
    property PDFSigningSignatureReason: WideString dispid 1073938534;
    property PDFSigningSignatureRightX: Double dispid 1073938535;
    property PDFSigningSignatureRightY: Double dispid 1073938536;
    property PDFSigningSignatureVisible: Integer dispid 1073938537;
    property PDFSigningSignPDF: Integer dispid 1073938538;
    property PDFUpdateMetadata: Integer dispid 1073938539;
    property PDFUserPass: Integer dispid 1073938540;
    property PDFUserPasswordString: WideString dispid 1073938541;
    property PDFUseSecurity: Integer dispid 1073938542;
    property PNGColorscount: Integer dispid 1073938543;
    property PNGResolution: Integer dispid 1073938544;
    property PrintAfterSaving: Integer dispid 1073938545;
    property PrintAfterSavingBitsPerPixel: Integer dispid 1073938610;
    property PrintAfterSavingDuplex: Integer dispid 1073938546;
    property PrintAfterSavingMaxResolution: Integer dispid 1073938611;
    property PrintAfterSavingMaxResolutionEnabled: Integer dispid 1073938612;
    property PrintAfterSavingNoCancel: Integer dispid 1073938547;
    property PrintAfterSavingPrinter: WideString dispid 1073938548;
    property PrintAfterSavingQueryUser: Integer dispid 1073938549;
    property PrintAfterSavingTumble: Integer dispid 1073938550;
    property PrinterStop: Integer dispid 1073938551;
    property PrinterTemppath: WideString dispid 1073938552;
    property ProcessPriority: Integer dispid 1073938553;
    property ProgramFont: WideString dispid 1073938554;
    property ProgramFontCharset: Integer dispid 1073938555;
    property ProgramFontSize: Integer dispid 1073938556;
    property PSDColorsCount: Integer dispid 1073938557;
    property PSDResolution: Integer dispid 1073938558;
    property PSLanguageLevel: Integer dispid 1073938559;
    property RAWColorsCount: Integer dispid 1073938560;
    property RAWResolution: Integer dispid 1073938561;
    property RemoveAllKnownFileExtensions: Integer dispid 1073938562;
    property RemoveSpaces: Integer dispid 1073938563;
    property RunProgramAfterSaving: Integer dispid 1073938564;
    property RunProgramAfterSavingProgramname: WideString dispid 1073938565;
    property RunProgramAfterSavingProgramParameters: WideString dispid 1073938566;
    property RunProgramAfterSavingWaitUntilReady: Integer dispid 1073938567;
    property RunProgramAfterSavingWindowstyle: Integer dispid 1073938568;
    property RunProgramBeforeSaving: Integer dispid 1073938569;
    property RunProgramBeforeSavingProgramname: WideString dispid 1073938570;
    property RunProgramBeforeSavingProgramParameters: WideString dispid 1073938571;
    property RunProgramBeforeSavingWindowstyle: Integer dispid 1073938572;
    property SaveFilename: WideString dispid 1073938573;
    property SendEmailAfterAutoSaving: Integer dispid 1073938574;
    property SendMailMethod: Integer dispid 1073938575;
    property ShowAnimation: Integer dispid 1073938576;
    property StampFontColor: WideString dispid 1073938577;
    property StampFontname: WideString dispid 1073938578;
    property StampFontsize: Integer dispid 1073938579;
    property StampOutlineFontthickness: Integer dispid 1073938580;
    property StampString: WideString dispid 1073938581;
    property StampUseOutlineFont: Integer dispid 1073938582;
    property StandardAuthor: WideString dispid 1073938583;
    property StandardCreationdate: WideString dispid 1073938584;
    property StandardDateformat: WideString dispid 1073938585;
    property StandardKeywords: WideString dispid 1073938586;
    property StandardMailDomain: WideString dispid 1073938587;
    property StandardModifydate: WideString dispid 1073938588;
    property StandardSaveformat: Integer dispid 1073938589;
    property StandardSubject: WideString dispid 1073938590;
    property StandardTitle: WideString dispid 1073938591;
    property StartStandardProgram: Integer dispid 1073938592;
    property SVGResolution: Integer dispid 1073938604;
    property TIFFColorscount: Integer dispid 1073938593;
    property TIFFResolution: Integer dispid 1073938594;
    property Toolbars: Integer dispid 1073938595;
    property UpdateInterval: Integer dispid 1073938605;
    property UseAutosave: Integer dispid 1073938596;
    property UseAutosaveDirectory: Integer dispid 1073938597;
    property UseCreationDateNow: Integer dispid 1073938598;
    property UseCustomPaperSize: WideString dispid 1073938599;
    property UseFixPapersize: Integer dispid 1073938600;
    property UseStandardAuthor: Integer dispid 1073938601;
  end;

// *********************************************************************//
// Interface   : _clsPDFCreatorError
// Indicateurs : (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID        : {A030F401-6045-4942-A5F5-9CCBF2C1872D}
// *********************************************************************//
  _clsPDFCreatorError = interface(IDispatch)
    ['{A030F401-6045-4942-A5F5-9CCBF2C1872D}']
    function Get_Number: Integer; safecall;
    procedure Set_Number(Number: Integer); safecall;
    function Get_Description: WideString; safecall;
    procedure Set_Description(const Description: WideString); safecall;
    property Number: Integer read Get_Number write Set_Number;
    property Description: WideString read Get_Description write Set_Description;
  end;

// *********************************************************************//
// DispIntf :  _clsPDFCreatorErrorDisp
// Flags :     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID :      {A030F401-6045-4942-A5F5-9CCBF2C1872D}
// *********************************************************************//
  _clsPDFCreatorErrorDisp = dispinterface
    ['{A030F401-6045-4942-A5F5-9CCBF2C1872D}']
    property Number: Integer dispid 1073938432;
    property Description: WideString dispid 1073938433;
  end;

// *********************************************************************//
// Interface   : _clsPDFCreatorInfoSpoolFile
// Indicateurs : (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID        : {253F17D9-1678-4B0E-843E-A2D37C2C6B4E}
// *********************************************************************//
  _clsPDFCreatorInfoSpoolFile = interface(IDispatch)
    ['{253F17D9-1678-4B0E-843E-A2D37C2C6B4E}']
    function Get_REDMON_PORT: WideString; safecall;
    procedure Set_REDMON_PORT(const REDMON_PORT: WideString); safecall;
    function Get_REDMON_JOB: WideString; safecall;
    procedure Set_REDMON_JOB(const REDMON_JOB: WideString); safecall;
    function Get_REDMON_PRINTER: WideString; safecall;
    procedure Set_REDMON_PRINTER(const REDMON_PRINTER: WideString); safecall;
    function Get_REDMON_MACHINE: WideString; safecall;
    procedure Set_REDMON_MACHINE(const REDMON_MACHINE: WideString); safecall;
    function Get_REDMON_USER: WideString; safecall;
    procedure Set_REDMON_USER(const REDMON_USER: WideString); safecall;
    function Get_REDMON_DOCNAME: WideString; safecall;
    procedure Set_REDMON_DOCNAME(const REDMON_DOCNAME: WideString); safecall;
    function Get_REDMON_FILENAME: WideString; safecall;
    procedure Set_REDMON_FILENAME(const REDMON_FILENAME: WideString); safecall;
    function Get_REDMON_SESSIONID: WideString; safecall;
    procedure Set_REDMON_SESSIONID(const REDMON_SESSIONID: WideString); safecall;
    function Get_SpoolFilename: WideString; safecall;
    procedure Set_SpoolFilename(const SpoolFilename: WideString); safecall;
    function Get_SpoolerAccount: WideString; safecall;
    procedure Set_SpoolerAccount(const SpoolerAccount: WideString); safecall;
    function Get_Computer: WideString; safecall;
    procedure Set_Computer(const Computer: WideString); safecall;
    function Get_Created: WideString; safecall;
    procedure Set_Created(const Created: WideString); safecall;
    property REDMON_PORT: WideString read Get_REDMON_PORT write Set_REDMON_PORT;
    property REDMON_JOB: WideString read Get_REDMON_JOB write Set_REDMON_JOB;
    property REDMON_PRINTER: WideString read Get_REDMON_PRINTER write Set_REDMON_PRINTER;
    property REDMON_MACHINE: WideString read Get_REDMON_MACHINE write Set_REDMON_MACHINE;
    property REDMON_USER: WideString read Get_REDMON_USER write Set_REDMON_USER;
    property REDMON_DOCNAME: WideString read Get_REDMON_DOCNAME write Set_REDMON_DOCNAME;
    property REDMON_FILENAME: WideString read Get_REDMON_FILENAME write Set_REDMON_FILENAME;
    property REDMON_SESSIONID: WideString read Get_REDMON_SESSIONID write Set_REDMON_SESSIONID;
    property SpoolFilename: WideString read Get_SpoolFilename write Set_SpoolFilename;
    property SpoolerAccount: WideString read Get_SpoolerAccount write Set_SpoolerAccount;
    property Computer: WideString read Get_Computer write Set_Computer;
    property Created: WideString read Get_Created write Set_Created;
  end;

// *********************************************************************//
// DispIntf :  _clsPDFCreatorInfoSpoolFileDisp
// Flags :     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID :      {253F17D9-1678-4B0E-843E-A2D37C2C6B4E}
// *********************************************************************//
  _clsPDFCreatorInfoSpoolFileDisp = dispinterface
    ['{253F17D9-1678-4B0E-843E-A2D37C2C6B4E}']
    property REDMON_PORT: WideString dispid 1073938432;
    property REDMON_JOB: WideString dispid 1073938433;
    property REDMON_PRINTER: WideString dispid 1073938434;
    property REDMON_MACHINE: WideString dispid 1073938435;
    property REDMON_USER: WideString dispid 1073938436;
    property REDMON_DOCNAME: WideString dispid 1073938437;
    property REDMON_FILENAME: WideString dispid 1073938438;
    property REDMON_SESSIONID: WideString dispid 1073938439;
    property SpoolFilename: WideString dispid 1073938440;
    property SpoolerAccount: WideString dispid 1073938441;
    property Computer: WideString dispid 1073938442;
    property Created: WideString dispid 1073938443;
  end;

// *********************************************************************//
// Interface   : _clsPDFCreator
// Indicateurs : (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID        : {280807DB-86BB-4FD1-B477-A7A6E285A3FE}
// *********************************************************************//
  _clsPDFCreator = interface(IDispatch)
    ['{280807DB-86BB-4FD1-B477-A7A6E285A3FE}']
    function Get_cPrinterProfile(const Printername: WideString): WideString; safecall;
    procedure Set_cPrinterProfile(const Printername: WideString; const Param2: WideString); safecall;
    function Get_cIsClosed: WordBool; safecall;
    function Get_cError: _clsPDFCreatorError; safecall;
    function Get_cErrorDetail(const PropertyName: WideString): OleVariant; safecall;
    procedure cErrorClear; safecall;
    function Get_cGhostscriptVersion: WideString; safecall;
    function Get_cOutputFilename: WideString; safecall;
    function Get_cPDFCreatorApplicationPath: WideString; safecall;
    function Get_cIsLogfileDialogDisplayed: WordBool; safecall;
    function Get_cIsOptionsDialogDisplayed: WordBool; safecall;
    function Get_cProgramRelease(WithBeta: WordBool): WideString; safecall;
    function Get_cProgramIsRunning: WordBool; safecall;
    function Get_cWindowsVersion: WideString; safecall;
    function Get_cVisible: WordBool; safecall;
    procedure Set_cVisible(Param1: WordBool); safecall;
    function Get_cInstalledAsServer: WordBool; safecall;
    function Get_cPrinterStop: WordBool; safecall;
    procedure Set_cPrinterStop(Param1: WordBool); safecall;
    function Get_cOptionsNames: _Collection; safecall;
    function Get_cOption(const PropertyName: WideString): OleVariant; safecall;
    procedure Set_cOption(const PropertyName: WideString; Param2: OleVariant); safecall;
    function Get_cOptionProfile(const ProfileName: WideString; const PropertyName: WideString): OleVariant; safecall;
    procedure Set_cOptionProfile(const ProfileName: WideString; const PropertyName: WideString; 
                                 Param3: OleVariant); safecall;
    function Get_cOptions: _clsPDFCreatorOptions; safecall;
    procedure _Set_cOptions(const Param1: _clsPDFCreatorOptions); safecall;
    function Get_cOptionsProfile(const ProfileName: WideString): _clsPDFCreatorOptions; safecall;
    procedure _Set_cOptionsProfile(const ProfileName: WideString; 
                                   const Param2: _clsPDFCreatorOptions); safecall;
    function Get_cStandardOption(const PropertyName: WideString): OleVariant; safecall;
    function Get_cStandardOptions: _clsPDFCreatorOptions; safecall;
    function Get_cPostscriptInfo(const PostscriptFilename: WideString; 
                                 const PropertyName: WideString): WideString; safecall;
    function Get_cPrintjobInfos(const PrintjobFilename: WideString): _clsPDFCreatorInfoSpoolFile; safecall;
    function Get_cPrintjobInfo(const PrintjobFilename: WideString; const PropertyName: WideString): WideString; safecall;
    function Get_cCountOfPrintjobs: Integer; safecall;
    function Get_cPrintjobFilename(JobNumber: Integer): WideString; safecall;
    function Get_cDefaultPrinter: WideString; safecall;
    procedure Set_cDefaultPrinter(const Param1: WideString); safecall;
    function Get_cStopURLPrinting: WordBool; safecall;
    procedure Set_cStopURLPrinting(Param1: WordBool); safecall;
    function Get_cWindowState: Integer; safecall;
    procedure Set_cWindowState(Param1: Integer); safecall;
    function Get_cIsConverted: WordBool; safecall;
    procedure Set_cIsConverted(Param1: WordBool); safecall;
    function Get_cInstanceCounter: Integer; safecall;
    function cIsAdministrator: WordBool; safecall;
    function cPrinterIsInstalled(const Printername: WideString): WordBool; safecall;
    function cAddPDFCreatorPrinter(const Printername: WideString; const ProfileName: WideString): WordBool; safecall;
    function cProfileExists(const ProfileName: WideString): WordBool; safecall;
    function cDeletePDFCreatorPrinter(const Printername: WideString): WordBool; safecall;
    function cGetProfileNames: _Collection; safecall;
    function cAddProfile(const ProfileName: WideString; const Options1: _clsPDFCreatorOptions): WordBool; safecall;
    function cRenameProfile(const OldProfileName: WideString; const NewProfileName: WideString): WordBool; safecall;
    function cDeleteProfile(const ProfileName: WideString): WordBool; safecall;
    procedure cAddPrintjob(const filename: WideString); safecall;
    procedure cDeletePrintjob(JobNumber: Integer); safecall;
    procedure cMovePrintjobBottom(JobNumber: Integer); safecall;
    procedure cMovePrintjobTop(JobNumber: Integer); safecall;
    procedure cMovePrintjobUp(JobNumber: Integer); safecall;
    procedure cMovePrintjobDown(JobNumber: Integer); safecall;
    function cClose: WordBool; safecall;
    function cStart(const Params: WideString; ForceInitialize: WordBool): WordBool; safecall;
    procedure cClearCache; safecall;
    procedure cClearLogfile; safecall;
    procedure cConvertPostscriptfile(const InputFilename: WideString; 
                                     const OutputFilename: WideString); safecall;
    procedure cConvertFile(const InputFilename: WideString; const OutputFilename: WideString; 
                           const SubFormat: WideString); safecall;
    procedure cTestEvent(const EventName: WideString); safecall;
    procedure cShowLogfileDialog(value: WordBool); safecall;
    procedure cShowOptionsDialog(value: WordBool); safecall;
    procedure cSendMail(const OutputFilename: WideString; const Recipients: WideString); safecall;
    function cIsPrintable(const filename: WideString): WordBool; safecall;
    procedure cCombineAll; safecall;
    function cGetPDFCreatorPrinters: _Collection; safecall;
    function cGetPrinterProfiles: _Collection; safecall;
    function cGetLogfile: WideString; safecall;
    procedure cWriteToLogfile(const LogStr: WideString); safecall;
    procedure cPrintFile(const filename: WideString); safecall;
    procedure cPrintURL(const URL: WideString; TimeBetweenLoadAndPrint: Integer); safecall;
    procedure cPrintPDFCreatorTestpage; safecall;
    procedure cPrintPrinterTestpage(const Printername: WideString); safecall;
    function cReadOptions(const ProfileName: WideString): _clsPDFCreatorOptions; safecall;
    procedure cSaveOptions(Options1: OleVariant; const ProfileName: WideString); safecall;
    function cReadOptionsFromFile(const INIFilename: WideString): _clsPDFCreatorOptions; safecall;
    procedure cSaveOptionsToFile(const INIFilename: WideString; Options1: OleVariant); safecall;
    function cGhostscriptRun(var Arguments: PSafeArray): WordBool; safecall;
    property cPrinterProfile[const Printername: WideString]: WideString read Get_cPrinterProfile write Set_cPrinterProfile;
    property cIsClosed: WordBool read Get_cIsClosed;
    property cError: _clsPDFCreatorError read Get_cError;
    property cErrorDetail[const PropertyName: WideString]: OleVariant read Get_cErrorDetail;
    property cGhostscriptVersion: WideString read Get_cGhostscriptVersion;
    property cOutputFilename: WideString read Get_cOutputFilename;
    property cPDFCreatorApplicationPath: WideString read Get_cPDFCreatorApplicationPath;
    property cIsLogfileDialogDisplayed: WordBool read Get_cIsLogfileDialogDisplayed;
    property cIsOptionsDialogDisplayed: WordBool read Get_cIsOptionsDialogDisplayed;
    property cProgramRelease[WithBeta: WordBool]: WideString read Get_cProgramRelease;
    property cProgramIsRunning: WordBool read Get_cProgramIsRunning;
    property cWindowsVersion: WideString read Get_cWindowsVersion;
    property cVisible: WordBool read Get_cVisible write Set_cVisible;
    property cInstalledAsServer: WordBool read Get_cInstalledAsServer;
    property cPrinterStop: WordBool read Get_cPrinterStop write Set_cPrinterStop;
    property cOptionsNames: _Collection read Get_cOptionsNames;
    property cOption[const PropertyName: WideString]: OleVariant read Get_cOption write Set_cOption;
    property cOptionProfile[const ProfileName: WideString; const PropertyName: WideString]: OleVariant read Get_cOptionProfile write Set_cOptionProfile;
    property cOptions: _clsPDFCreatorOptions read Get_cOptions write _Set_cOptions;
    property cOptionsProfile[const ProfileName: WideString]: _clsPDFCreatorOptions read Get_cOptionsProfile write _Set_cOptionsProfile;
    property cStandardOption[const PropertyName: WideString]: OleVariant read Get_cStandardOption;
    property cStandardOptions: _clsPDFCreatorOptions read Get_cStandardOptions;
    property cPostscriptInfo[const PostscriptFilename: WideString; const PropertyName: WideString]: WideString read Get_cPostscriptInfo;
    property cPrintjobInfos[const PrintjobFilename: WideString]: _clsPDFCreatorInfoSpoolFile read Get_cPrintjobInfos;
    property cPrintjobInfo[const PrintjobFilename: WideString; const PropertyName: WideString]: WideString read Get_cPrintjobInfo;
    property cCountOfPrintjobs: Integer read Get_cCountOfPrintjobs;
    property cPrintjobFilename[JobNumber: Integer]: WideString read Get_cPrintjobFilename;
    property cDefaultPrinter: WideString read Get_cDefaultPrinter write Set_cDefaultPrinter;
    property cStopURLPrinting: WordBool read Get_cStopURLPrinting write Set_cStopURLPrinting;
    property cWindowState: Integer read Get_cWindowState write Set_cWindowState;
    property cIsConverted: WordBool read Get_cIsConverted write Set_cIsConverted;
    property cInstanceCounter: Integer read Get_cInstanceCounter;
  end;

// *********************************************************************//
// DispIntf :  _clsPDFCreatorDisp
// Flags :     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID :      {280807DB-86BB-4FD1-B477-A7A6E285A3FE}
// *********************************************************************//
  _clsPDFCreatorDisp = dispinterface
    ['{280807DB-86BB-4FD1-B477-A7A6E285A3FE}']
    property cPrinterProfile[const Printername: WideString]: WideString dispid 1745027172;
    property cIsClosed: WordBool readonly dispid 1745027171;
    property cError: _clsPDFCreatorError readonly dispid 1745027138;
    property cErrorDetail[const PropertyName: WideString]: OleVariant readonly dispid 1745027137;
    procedure cErrorClear; dispid 1610809412;
    property cGhostscriptVersion: WideString readonly dispid 1745027136;
    property cOutputFilename: WideString readonly dispid 1745027135;
    property cPDFCreatorApplicationPath: WideString readonly dispid 1745027134;
    property cIsLogfileDialogDisplayed: WordBool readonly dispid 1745027133;
    property cIsOptionsDialogDisplayed: WordBool readonly dispid 1745027132;
    property cProgramRelease[WithBeta: WordBool]: WideString readonly dispid 1745027131;
    property cProgramIsRunning: WordBool readonly dispid 1745027130;
    property cWindowsVersion: WideString readonly dispid 1745027129;
    property cVisible: WordBool dispid 1745027128;
    property cInstalledAsServer: WordBool readonly dispid 1745027127;
    property cPrinterStop: WordBool dispid 1745027126;
    property cOptionsNames: _Collection readonly dispid 1745027125;
    property cOption[const PropertyName: WideString]: OleVariant dispid 1745027124;
    property cOptionProfile[const ProfileName: WideString; const PropertyName: WideString]: OleVariant dispid 1745027123;
    property cOptions: _clsPDFCreatorOptions dispid 1745027122;
    property cOptionsProfile[const ProfileName: WideString]: _clsPDFCreatorOptions dispid 1745027121;
    property cStandardOption[const PropertyName: WideString]: OleVariant readonly dispid 1745027120;
    property cStandardOptions: _clsPDFCreatorOptions readonly dispid 1745027119;
    property cPostscriptInfo[const PostscriptFilename: WideString; const PropertyName: WideString]: WideString readonly dispid 1745027118;
    property cPrintjobInfos[const PrintjobFilename: WideString]: _clsPDFCreatorInfoSpoolFile readonly dispid 1745027117;
    property cPrintjobInfo[const PrintjobFilename: WideString; const PropertyName: WideString]: WideString readonly dispid 1745027116;
    property cCountOfPrintjobs: Integer readonly dispid 1745027115;
    property cPrintjobFilename[JobNumber: Integer]: WideString readonly dispid 1745027114;
    property cDefaultPrinter: WideString dispid 1745027113;
    property cStopURLPrinting: WordBool dispid 1745027112;
    property cWindowState: Integer dispid 1745027111;
    property cIsConverted: WordBool dispid 1745027110;
    property cInstanceCounter: Integer readonly dispid 1745027109;
    function cIsAdministrator: WordBool; dispid 1610809447;
    function cPrinterIsInstalled(const Printername: WideString): WordBool; dispid 1610809448;
    function cAddPDFCreatorPrinter(const Printername: WideString; const ProfileName: WideString): WordBool; dispid 1610809449;
    function cProfileExists(const ProfileName: WideString): WordBool; dispid 1610809450;
    function cDeletePDFCreatorPrinter(const Printername: WideString): WordBool; dispid 1610809451;
    function cGetProfileNames: _Collection; dispid 1610809452;
    function cAddProfile(const ProfileName: WideString; const Options1: _clsPDFCreatorOptions): WordBool; dispid 1610809453;
    function cRenameProfile(const OldProfileName: WideString; const NewProfileName: WideString): WordBool; dispid 1610809454;
    function cDeleteProfile(const ProfileName: WideString): WordBool; dispid 1610809455;
    procedure cAddPrintjob(const filename: WideString); dispid 1610809413;
    procedure cDeletePrintjob(JobNumber: Integer); dispid 1610809414;
    procedure cMovePrintjobBottom(JobNumber: Integer); dispid 1610809415;
    procedure cMovePrintjobTop(JobNumber: Integer); dispid 1610809416;
    procedure cMovePrintjobUp(JobNumber: Integer); dispid 1610809417;
    procedure cMovePrintjobDown(JobNumber: Integer); dispid 1610809418;
    function cClose: WordBool; dispid 1610809419;
    function cStart(const Params: WideString; ForceInitialize: WordBool): WordBool; dispid 1610809420;
    procedure cClearCache; dispid 1610809421;
    procedure cClearLogfile; dispid 1610809422;
    procedure cConvertPostscriptfile(const InputFilename: WideString; 
                                     const OutputFilename: WideString); dispid 1610809423;
    procedure cConvertFile(const InputFilename: WideString; const OutputFilename: WideString; 
                           const SubFormat: WideString); dispid 1610809424;
    procedure cTestEvent(const EventName: WideString); dispid 1610809425;
    procedure cShowLogfileDialog(value: WordBool); dispid 1610809426;
    procedure cShowOptionsDialog(value: WordBool); dispid 1610809427;
    procedure cSendMail(const OutputFilename: WideString; const Recipients: WideString); dispid 1610809428;
    function cIsPrintable(const filename: WideString): WordBool; dispid 1610809429;
    procedure cCombineAll; dispid 1610809430;
    function cGetPDFCreatorPrinters: _Collection; dispid 1610809431;
    function cGetPrinterProfiles: _Collection; dispid 1610809456;
    function cGetLogfile: WideString; dispid 1610809432;
    procedure cWriteToLogfile(const LogStr: WideString); dispid 1610809433;
    procedure cPrintFile(const filename: WideString); dispid 1610809434;
    procedure cPrintURL(const URL: WideString; TimeBetweenLoadAndPrint: Integer); dispid 1610809435;
    procedure cPrintPDFCreatorTestpage; dispid 1610809436;
    procedure cPrintPrinterTestpage(const Printername: WideString); dispid 1610809437;
    function cReadOptions(const ProfileName: WideString): _clsPDFCreatorOptions; dispid 1610809438;
    procedure cSaveOptions(Options1: OleVariant; const ProfileName: WideString); dispid 1610809439;
    function cReadOptionsFromFile(const INIFilename: WideString): _clsPDFCreatorOptions; dispid 1610809440;
    procedure cSaveOptionsToFile(const INIFilename: WideString; Options1: OleVariant); dispid 1610809441;
    function cGhostscriptRun(var Arguments: {??PSafeArray}OleVariant): WordBool; dispid 1610809442;
  end;

// *********************************************************************//
// DispIntf :  __clsPDFCreator
// Flags :     (4240) Hidden NonExtensible Dispatchable
// GUID :      {4B1FCFC1-EB7C-4180-B0B6-68EBA12FCF88}
// *********************************************************************//
  __clsPDFCreator = dispinterface
    ['{4B1FCFC1-EB7C-4180-B0B6-68EBA12FCF88}']
    procedure eReady; dispid 1;
    procedure eError; dispid 2;
  end;

// *********************************************************************//
// Interface   : _clsTools
// Indicateurs : (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID        : {58CCA1EC-FC65-4750-A7E7-FFA191BD1042}
// *********************************************************************//
  _clsTools = interface(IDispatch)
    ['{58CCA1EC-FC65-4750-A7E7-FFA191BD1042}']
    function cOpenFileDialog(var files: OleVariant; const InitFilename: WideString; 
                             const Filter: WideString; const DefaultFileExtension: WideString; 
                             const InitDir: WideString; const DialogTitle: WideString; 
                             Flags: Integer; hwnd: Integer; FilterIndex: Integer): Integer; safecall;
    function cSaveFileDialog(var filename: OleVariant; var InitFilename: WideString; 
                             var Filter: WideString; var DefaultFileExtension: WideString; 
                             var InitDir: WideString; var DialogTitle: WideString; 
                             var Flags: Integer; var hwnd: Integer; var FilterIndex: Integer): Integer; safecall;
  end;

// *********************************************************************//
// DispIntf :  _clsToolsDisp
// Flags :     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID :      {58CCA1EC-FC65-4750-A7E7-FFA191BD1042}
// *********************************************************************//
  _clsToolsDisp = dispinterface
    ['{58CCA1EC-FC65-4750-A7E7-FFA191BD1042}']
    function cOpenFileDialog(var files: OleVariant; const InitFilename: WideString; 
                             const Filter: WideString; const DefaultFileExtension: WideString; 
                             const InitDir: WideString; const DialogTitle: WideString; 
                             Flags: Integer; hwnd: Integer; FilterIndex: Integer): Integer; dispid 1610809344;
    function cSaveFileDialog(var filename: OleVariant; var InitFilename: WideString; 
                             var Filter: WideString; var DefaultFileExtension: WideString; 
                             var InitDir: WideString; var DialogTitle: WideString; 
                             var Flags: Integer; var hwnd: Integer; var FilterIndex: Integer): Integer; dispid 1610809345;
  end;

// *********************************************************************//
// Interface   : _clsUpdate
// Indicateurs : (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID        : {6381BF10-F5C2-4577-B6CA-EDAF0287185D}
// *********************************************************************//
  _clsUpdate = interface(IDispatch)
    ['{6381BF10-F5C2-4577-B6CA-EDAF0287185D}']
    procedure CheckForUpdates(var ShowMessageNoNewUpdates: WordBool; 
                              var ShowErrorMessage: WordBool; var TimeOutInMs: Integer); safecall;
  end;

// *********************************************************************//
// DispIntf :  _clsUpdateDisp
// Flags :     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID :      {6381BF10-F5C2-4577-B6CA-EDAF0287185D}
// *********************************************************************//
  _clsUpdateDisp = dispinterface
    ['{6381BF10-F5C2-4577-B6CA-EDAF0287185D}']
    procedure CheckForUpdates(var ShowMessageNoNewUpdates: WordBool; 
                              var ShowErrorMessage: WordBool; var TimeOutInMs: Integer); dispid 1610809344;
  end;

// *********************************************************************//
// La classe CoclsPDFCreatorOptions fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut _clsPDFCreatorOptions exposée             
// par la CoClasse clsPDFCreatorOptions. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoclsPDFCreatorOptions = class
    class function Create: _clsPDFCreatorOptions;
    class function CreateRemote(const MachineName: string): _clsPDFCreatorOptions;
  end;


// *********************************************************************//
// Déclaration de classe proxy de serveur OLE
// Objet serveur        : TclsPDFCreatorOptions
// Chaîne d'aide        : 
// Interface par défaut : _clsPDFCreatorOptions
// DISP Int. Déf. ?     : No
// Interface événements : 
// TypeFlags            : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TclsPDFCreatorOptionsProperties= class;
{$ENDIF}
  TclsPDFCreatorOptions = class(TOleServer)
  private
    FIntf:        _clsPDFCreatorOptions;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TclsPDFCreatorOptionsProperties;
    function      GetServerProperties: TclsPDFCreatorOptionsProperties;
{$ENDIF}
    function      GetDefaultInterface: _clsPDFCreatorOptions;
  protected
    procedure InitServerData; override;
    function Get_AdditionalGhostscriptParameters: WideString;
    procedure Set_AdditionalGhostscriptParameters(const AdditionalGhostscriptParameters: WideString);
    function Get_AdditionalGhostscriptSearchpath: WideString;
    procedure Set_AdditionalGhostscriptSearchpath(const AdditionalGhostscriptSearchpath: WideString);
    function Get_AddWindowsFontpath: Integer;
    procedure Set_AddWindowsFontpath(AddWindowsFontpath: Integer);
    function Get_AllowSpecialGSCharsInFilenames: Integer;
    procedure Set_AllowSpecialGSCharsInFilenames(AllowSpecialGSCharsInFilenames: Integer);
    function Get_AutosaveDirectory: WideString;
    procedure Set_AutosaveDirectory(const AutosaveDirectory: WideString);
    function Get_AutosaveFilename: WideString;
    procedure Set_AutosaveFilename(const AutosaveFilename: WideString);
    function Get_AutosaveFormat: Integer;
    procedure Set_AutosaveFormat(AutosaveFormat: Integer);
    function Get_AutosaveStartStandardProgram: Integer;
    procedure Set_AutosaveStartStandardProgram(AutosaveStartStandardProgram: Integer);
    function Get_BMPColorscount: Integer;
    procedure Set_BMPColorscount(BMPColorscount: Integer);
    function Get_BMPResolution: Integer;
    procedure Set_BMPResolution(BMPResolution: Integer);
    function Get_ClientComputerResolveIPAddress: Integer;
    procedure Set_ClientComputerResolveIPAddress(ClientComputerResolveIPAddress: Integer);
    function Get_Counter: Currency;
    procedure Set_Counter(Counter: Currency);
    function Get_DeviceHeightPoints: Double;
    procedure Set_DeviceHeightPoints(DeviceHeightPoints: Double);
    function Get_DeviceWidthPoints: Double;
    procedure Set_DeviceWidthPoints(DeviceWidthPoints: Double);
    function Get_DirectoryGhostscriptBinaries: WideString;
    procedure Set_DirectoryGhostscriptBinaries(const DirectoryGhostscriptBinaries: WideString);
    function Get_DirectoryGhostscriptFonts: WideString;
    procedure Set_DirectoryGhostscriptFonts(const DirectoryGhostscriptFonts: WideString);
    function Get_DirectoryGhostscriptLibraries: WideString;
    procedure Set_DirectoryGhostscriptLibraries(const DirectoryGhostscriptLibraries: WideString);
    function Get_DirectoryGhostscriptResource: WideString;
    procedure Set_DirectoryGhostscriptResource(const DirectoryGhostscriptResource: WideString);
    function Get_DisableEmail: Integer;
    procedure Set_DisableEmail(DisableEmail: Integer);
    function Get_DontUseDocumentSettings: Integer;
    procedure Set_DontUseDocumentSettings(DontUseDocumentSettings: Integer);
    function Get_EPSLanguageLevel: Integer;
    procedure Set_EPSLanguageLevel(EPSLanguageLevel: Integer);
    function Get_FilenameSubstitutions: WideString;
    procedure Set_FilenameSubstitutions(const FilenameSubstitutions: WideString);
    function Get_FilenameSubstitutionsOnlyInTitle: Integer;
    procedure Set_FilenameSubstitutionsOnlyInTitle(FilenameSubstitutionsOnlyInTitle: Integer);
    function Get_JPEGColorscount: Integer;
    procedure Set_JPEGColorscount(JPEGColorscount: Integer);
    function Get_JPEGQuality: Integer;
    procedure Set_JPEGQuality(JPEGQuality: Integer);
    function Get_JPEGResolution: Integer;
    procedure Set_JPEGResolution(JPEGResolution: Integer);
    function Get_Language: WideString;
    procedure Set_Language(const Language: WideString);
    function Get_LastSaveDirectory: WideString;
    procedure Set_LastSaveDirectory(const LastSaveDirectory: WideString);
    function Get_LastUpdateCheck: WideString;
    procedure Set_LastUpdateCheck(const LastUpdateCheck: WideString);
    function Get_Logging: Integer;
    procedure Set_Logging(Logging: Integer);
    function Get_LogLines: Integer;
    procedure Set_LogLines(LogLines: Integer);
    function Get_NoConfirmMessageSwitchingDefaultprinter: Integer;
    procedure Set_NoConfirmMessageSwitchingDefaultprinter(NoConfirmMessageSwitchingDefaultprinter: Integer);
    function Get_NoProcessingAtStartup: Integer;
    procedure Set_NoProcessingAtStartup(NoProcessingAtStartup: Integer);
    function Get_NoPSCheck: Integer;
    procedure Set_NoPSCheck(NoPSCheck: Integer);
    function Get_OnePagePerFile: Integer;
    procedure Set_OnePagePerFile(OnePagePerFile: Integer);
    function Get_OptionsDesign: Integer;
    procedure Set_OptionsDesign(OptionsDesign: Integer);
    function Get_OptionsEnabled: Integer;
    procedure Set_OptionsEnabled(OptionsEnabled: Integer);
    function Get_OptionsVisible: Integer;
    procedure Set_OptionsVisible(OptionsVisible: Integer);
    function Get_Papersize: WideString;
    procedure Set_Papersize(const Papersize: WideString);
    function Get_PCLColorsCount: Integer;
    procedure Set_PCLColorsCount(PCLColorsCount: Integer);
    function Get_PCLResolution: Integer;
    procedure Set_PCLResolution(PCLResolution: Integer);
    function Get_PCXColorscount: Integer;
    procedure Set_PCXColorscount(PCXColorscount: Integer);
    function Get_PCXResolution: Integer;
    procedure Set_PCXResolution(PCXResolution: Integer);
    function Get_PDFAes128Encryption: Integer;
    procedure Set_PDFAes128Encryption(PDFAes128Encryption: Integer);
    function Get_PDFAllowAssembly: Integer;
    procedure Set_PDFAllowAssembly(PDFAllowAssembly: Integer);
    function Get_PDFAllowDegradedPrinting: Integer;
    procedure Set_PDFAllowDegradedPrinting(PDFAllowDegradedPrinting: Integer);
    function Get_PDFAllowFillIn: Integer;
    procedure Set_PDFAllowFillIn(PDFAllowFillIn: Integer);
    function Get_PDFAllowScreenReaders: Integer;
    procedure Set_PDFAllowScreenReaders(PDFAllowScreenReaders: Integer);
    function Get_PDFColorsCMYKToRGB: Integer;
    procedure Set_PDFColorsCMYKToRGB(PDFColorsCMYKToRGB: Integer);
    function Get_PDFColorsColorModel: Integer;
    procedure Set_PDFColorsColorModel(PDFColorsColorModel: Integer);
    function Get_PDFColorsPreserveHalftone: Integer;
    procedure Set_PDFColorsPreserveHalftone(PDFColorsPreserveHalftone: Integer);
    function Get_PDFColorsPreserveOverprint: Integer;
    procedure Set_PDFColorsPreserveOverprint(PDFColorsPreserveOverprint: Integer);
    function Get_PDFColorsPreserveTransfer: Integer;
    procedure Set_PDFColorsPreserveTransfer(PDFColorsPreserveTransfer: Integer);
    function Get_PDFCompressionColorCompression: Integer;
    procedure Set_PDFCompressionColorCompression(PDFCompressionColorCompression: Integer);
    function Get_PDFCompressionColorCompressionChoice: Integer;
    procedure Set_PDFCompressionColorCompressionChoice(PDFCompressionColorCompressionChoice: Integer);
    function Get_PDFCompressionColorCompressionJPEGHighFactor: Double;
    procedure Set_PDFCompressionColorCompressionJPEGHighFactor(PDFCompressionColorCompressionJPEGHighFactor: Double);
    function Get_PDFCompressionColorCompressionJPEGLowFactor: Double;
    procedure Set_PDFCompressionColorCompressionJPEGLowFactor(PDFCompressionColorCompressionJPEGLowFactor: Double);
    function Get_PDFCompressionColorCompressionJPEGManualFactor: Double;
    procedure Set_PDFCompressionColorCompressionJPEGManualFactor(PDFCompressionColorCompressionJPEGManualFactor: Double);
    function Get_PDFCompressionColorCompressionJPEGMaximumFactor: Double;
    procedure Set_PDFCompressionColorCompressionJPEGMaximumFactor(PDFCompressionColorCompressionJPEGMaximumFactor: Double);
    function Get_PDFCompressionColorCompressionJPEGMediumFactor: Double;
    procedure Set_PDFCompressionColorCompressionJPEGMediumFactor(PDFCompressionColorCompressionJPEGMediumFactor: Double);
    function Get_PDFCompressionColorCompressionJPEGMinimumFactor: Double;
    procedure Set_PDFCompressionColorCompressionJPEGMinimumFactor(PDFCompressionColorCompressionJPEGMinimumFactor: Double);
    function Get_PDFCompressionColorResample: Integer;
    procedure Set_PDFCompressionColorResample(PDFCompressionColorResample: Integer);
    function Get_PDFCompressionColorResampleChoice: Integer;
    procedure Set_PDFCompressionColorResampleChoice(PDFCompressionColorResampleChoice: Integer);
    function Get_PDFCompressionColorResolution: Integer;
    procedure Set_PDFCompressionColorResolution(PDFCompressionColorResolution: Integer);
    function Get_PDFCompressionGreyCompression: Integer;
    procedure Set_PDFCompressionGreyCompression(PDFCompressionGreyCompression: Integer);
    function Get_PDFCompressionGreyCompressionChoice: Integer;
    procedure Set_PDFCompressionGreyCompressionChoice(PDFCompressionGreyCompressionChoice: Integer);
    function Get_PDFCompressionGreyCompressionJPEGHighFactor: Double;
    procedure Set_PDFCompressionGreyCompressionJPEGHighFactor(PDFCompressionGreyCompressionJPEGHighFactor: Double);
    function Get_PDFCompressionGreyCompressionJPEGLowFactor: Double;
    procedure Set_PDFCompressionGreyCompressionJPEGLowFactor(PDFCompressionGreyCompressionJPEGLowFactor: Double);
    function Get_PDFCompressionGreyCompressionJPEGManualFactor: Double;
    procedure Set_PDFCompressionGreyCompressionJPEGManualFactor(PDFCompressionGreyCompressionJPEGManualFactor: Double);
    function Get_PDFCompressionGreyCompressionJPEGMaximumFactor: Double;
    procedure Set_PDFCompressionGreyCompressionJPEGMaximumFactor(PDFCompressionGreyCompressionJPEGMaximumFactor: Double);
    function Get_PDFCompressionGreyCompressionJPEGMediumFactor: Double;
    procedure Set_PDFCompressionGreyCompressionJPEGMediumFactor(PDFCompressionGreyCompressionJPEGMediumFactor: Double);
    function Get_PDFCompressionGreyCompressionJPEGMinimumFactor: Double;
    procedure Set_PDFCompressionGreyCompressionJPEGMinimumFactor(PDFCompressionGreyCompressionJPEGMinimumFactor: Double);
    function Get_PDFCompressionGreyResample: Integer;
    procedure Set_PDFCompressionGreyResample(PDFCompressionGreyResample: Integer);
    function Get_PDFCompressionGreyResampleChoice: Integer;
    procedure Set_PDFCompressionGreyResampleChoice(PDFCompressionGreyResampleChoice: Integer);
    function Get_PDFCompressionGreyResolution: Integer;
    procedure Set_PDFCompressionGreyResolution(PDFCompressionGreyResolution: Integer);
    function Get_PDFCompressionMonoCompression: Integer;
    procedure Set_PDFCompressionMonoCompression(PDFCompressionMonoCompression: Integer);
    function Get_PDFCompressionMonoCompressionChoice: Integer;
    procedure Set_PDFCompressionMonoCompressionChoice(PDFCompressionMonoCompressionChoice: Integer);
    function Get_PDFCompressionMonoResample: Integer;
    procedure Set_PDFCompressionMonoResample(PDFCompressionMonoResample: Integer);
    function Get_PDFCompressionMonoResampleChoice: Integer;
    procedure Set_PDFCompressionMonoResampleChoice(PDFCompressionMonoResampleChoice: Integer);
    function Get_PDFCompressionMonoResolution: Integer;
    procedure Set_PDFCompressionMonoResolution(PDFCompressionMonoResolution: Integer);
    function Get_PDFCompressionTextCompression: Integer;
    procedure Set_PDFCompressionTextCompression(PDFCompressionTextCompression: Integer);
    function Get_PDFDisallowCopy: Integer;
    procedure Set_PDFDisallowCopy(PDFDisallowCopy: Integer);
    function Get_PDFDisallowModifyAnnotations: Integer;
    procedure Set_PDFDisallowModifyAnnotations(PDFDisallowModifyAnnotations: Integer);
    function Get_PDFDisallowModifyContents: Integer;
    procedure Set_PDFDisallowModifyContents(PDFDisallowModifyContents: Integer);
    function Get_PDFDisallowPrinting: Integer;
    procedure Set_PDFDisallowPrinting(PDFDisallowPrinting: Integer);
    function Get_PDFEncryptor: Integer;
    procedure Set_PDFEncryptor(PDFEncryptor: Integer);
    function Get_PDFFontsEmbedAll: Integer;
    procedure Set_PDFFontsEmbedAll(PDFFontsEmbedAll: Integer);
    function Get_PDFFontsSubSetFonts: Integer;
    procedure Set_PDFFontsSubSetFonts(PDFFontsSubSetFonts: Integer);
    function Get_PDFFontsSubSetFontsPercent: Integer;
    procedure Set_PDFFontsSubSetFontsPercent(PDFFontsSubSetFontsPercent: Integer);
    function Get_PDFGeneralASCII85: Integer;
    procedure Set_PDFGeneralASCII85(PDFGeneralASCII85: Integer);
    function Get_PDFGeneralAutorotate: Integer;
    procedure Set_PDFGeneralAutorotate(PDFGeneralAutorotate: Integer);
    function Get_PDFGeneralCompatibility: Integer;
    procedure Set_PDFGeneralCompatibility(PDFGeneralCompatibility: Integer);
    function Get_PDFGeneralDefault: Integer;
    procedure Set_PDFGeneralDefault(PDFGeneralDefault: Integer);
    function Get_PDFGeneralOverprint: Integer;
    procedure Set_PDFGeneralOverprint(PDFGeneralOverprint: Integer);
    function Get_PDFGeneralResolution: Integer;
    procedure Set_PDFGeneralResolution(PDFGeneralResolution: Integer);
    function Get_PDFHighEncryption: Integer;
    procedure Set_PDFHighEncryption(PDFHighEncryption: Integer);
    function Get_PDFLowEncryption: Integer;
    procedure Set_PDFLowEncryption(PDFLowEncryption: Integer);
    function Get_PDFOptimize: Integer;
    procedure Set_PDFOptimize(PDFOptimize: Integer);
    function Get_PDFOwnerPass: Integer;
    procedure Set_PDFOwnerPass(PDFOwnerPass: Integer);
    function Get_PDFOwnerPasswordString: WideString;
    procedure Set_PDFOwnerPasswordString(const PDFOwnerPasswordString: WideString);
    function Get_PDFSigningMultiSignature: Integer;
    procedure Set_PDFSigningMultiSignature(PDFSigningMultiSignature: Integer);
    function Get_PDFSigningPFXFile: WideString;
    procedure Set_PDFSigningPFXFile(const PDFSigningPFXFile: WideString);
    function Get_PDFSigningPFXFilePassword: WideString;
    procedure Set_PDFSigningPFXFilePassword(const PDFSigningPFXFilePassword: WideString);
    function Get_PDFSigningSignatureContact: WideString;
    procedure Set_PDFSigningSignatureContact(const PDFSigningSignatureContact: WideString);
    function Get_PDFSigningSignatureLeftX: Double;
    procedure Set_PDFSigningSignatureLeftX(PDFSigningSignatureLeftX: Double);
    function Get_PDFSigningSignatureLeftY: Double;
    procedure Set_PDFSigningSignatureLeftY(PDFSigningSignatureLeftY: Double);
    function Get_PDFSigningSignatureLocation: WideString;
    procedure Set_PDFSigningSignatureLocation(const PDFSigningSignatureLocation: WideString);
    function Get_PDFSigningSignatureOnPage: Integer;
    procedure Set_PDFSigningSignatureOnPage(PDFSigningSignatureOnPage: Integer);
    function Get_PDFSigningSignatureReason: WideString;
    procedure Set_PDFSigningSignatureReason(const PDFSigningSignatureReason: WideString);
    function Get_PDFSigningSignatureRightX: Double;
    procedure Set_PDFSigningSignatureRightX(PDFSigningSignatureRightX: Double);
    function Get_PDFSigningSignatureRightY: Double;
    procedure Set_PDFSigningSignatureRightY(PDFSigningSignatureRightY: Double);
    function Get_PDFSigningSignatureVisible: Integer;
    procedure Set_PDFSigningSignatureVisible(PDFSigningSignatureVisible: Integer);
    function Get_PDFSigningSignPDF: Integer;
    procedure Set_PDFSigningSignPDF(PDFSigningSignPDF: Integer);
    function Get_PDFUpdateMetadata: Integer;
    procedure Set_PDFUpdateMetadata(PDFUpdateMetadata: Integer);
    function Get_PDFUserPass: Integer;
    procedure Set_PDFUserPass(PDFUserPass: Integer);
    function Get_PDFUserPasswordString: WideString;
    procedure Set_PDFUserPasswordString(const PDFUserPasswordString: WideString);
    function Get_PDFUseSecurity: Integer;
    procedure Set_PDFUseSecurity(PDFUseSecurity: Integer);
    function Get_PNGColorscount: Integer;
    procedure Set_PNGColorscount(PNGColorscount: Integer);
    function Get_PNGResolution: Integer;
    procedure Set_PNGResolution(PNGResolution: Integer);
    function Get_PrintAfterSaving: Integer;
    procedure Set_PrintAfterSaving(PrintAfterSaving: Integer);
    function Get_PrintAfterSavingBitsPerPixel: Integer;
    procedure Set_PrintAfterSavingBitsPerPixel(PrintAfterSavingBitsPerPixel: Integer);
    function Get_PrintAfterSavingDuplex: Integer;
    procedure Set_PrintAfterSavingDuplex(PrintAfterSavingDuplex: Integer);
    function Get_PrintAfterSavingMaxResolution: Integer;
    procedure Set_PrintAfterSavingMaxResolution(PrintAfterSavingMaxResolution: Integer);
    function Get_PrintAfterSavingMaxResolutionEnabled: Integer;
    procedure Set_PrintAfterSavingMaxResolutionEnabled(PrintAfterSavingMaxResolutionEnabled: Integer);
    function Get_PrintAfterSavingNoCancel: Integer;
    procedure Set_PrintAfterSavingNoCancel(PrintAfterSavingNoCancel: Integer);
    function Get_PrintAfterSavingPrinter: WideString;
    procedure Set_PrintAfterSavingPrinter(const PrintAfterSavingPrinter: WideString);
    function Get_PrintAfterSavingQueryUser: Integer;
    procedure Set_PrintAfterSavingQueryUser(PrintAfterSavingQueryUser: Integer);
    function Get_PrintAfterSavingTumble: Integer;
    procedure Set_PrintAfterSavingTumble(PrintAfterSavingTumble: Integer);
    function Get_PrinterStop: Integer;
    procedure Set_PrinterStop(PrinterStop: Integer);
    function Get_PrinterTemppath: WideString;
    procedure Set_PrinterTemppath(const PrinterTemppath: WideString);
    function Get_ProcessPriority: Integer;
    procedure Set_ProcessPriority(ProcessPriority: Integer);
    function Get_ProgramFont: WideString;
    procedure Set_ProgramFont(const ProgramFont: WideString);
    function Get_ProgramFontCharset: Integer;
    procedure Set_ProgramFontCharset(ProgramFontCharset: Integer);
    function Get_ProgramFontSize: Integer;
    procedure Set_ProgramFontSize(ProgramFontSize: Integer);
    function Get_PSDColorsCount: Integer;
    procedure Set_PSDColorsCount(PSDColorsCount: Integer);
    function Get_PSDResolution: Integer;
    procedure Set_PSDResolution(PSDResolution: Integer);
    function Get_PSLanguageLevel: Integer;
    procedure Set_PSLanguageLevel(PSLanguageLevel: Integer);
    function Get_RAWColorsCount: Integer;
    procedure Set_RAWColorsCount(RAWColorsCount: Integer);
    function Get_RAWResolution: Integer;
    procedure Set_RAWResolution(RAWResolution: Integer);
    function Get_RemoveAllKnownFileExtensions: Integer;
    procedure Set_RemoveAllKnownFileExtensions(RemoveAllKnownFileExtensions: Integer);
    function Get_RemoveSpaces: Integer;
    procedure Set_RemoveSpaces(RemoveSpaces: Integer);
    function Get_RunProgramAfterSaving: Integer;
    procedure Set_RunProgramAfterSaving(RunProgramAfterSaving: Integer);
    function Get_RunProgramAfterSavingProgramname: WideString;
    procedure Set_RunProgramAfterSavingProgramname(const RunProgramAfterSavingProgramname: WideString);
    function Get_RunProgramAfterSavingProgramParameters: WideString;
    procedure Set_RunProgramAfterSavingProgramParameters(const RunProgramAfterSavingProgramParameters: WideString);
    function Get_RunProgramAfterSavingWaitUntilReady: Integer;
    procedure Set_RunProgramAfterSavingWaitUntilReady(RunProgramAfterSavingWaitUntilReady: Integer);
    function Get_RunProgramAfterSavingWindowstyle: Integer;
    procedure Set_RunProgramAfterSavingWindowstyle(RunProgramAfterSavingWindowstyle: Integer);
    function Get_RunProgramBeforeSaving: Integer;
    procedure Set_RunProgramBeforeSaving(RunProgramBeforeSaving: Integer);
    function Get_RunProgramBeforeSavingProgramname: WideString;
    procedure Set_RunProgramBeforeSavingProgramname(const RunProgramBeforeSavingProgramname: WideString);
    function Get_RunProgramBeforeSavingProgramParameters: WideString;
    procedure Set_RunProgramBeforeSavingProgramParameters(const RunProgramBeforeSavingProgramParameters: WideString);
    function Get_RunProgramBeforeSavingWindowstyle: Integer;
    procedure Set_RunProgramBeforeSavingWindowstyle(RunProgramBeforeSavingWindowstyle: Integer);
    function Get_SaveFilename: WideString;
    procedure Set_SaveFilename(const SaveFilename: WideString);
    function Get_SendEmailAfterAutoSaving: Integer;
    procedure Set_SendEmailAfterAutoSaving(SendEmailAfterAutoSaving: Integer);
    function Get_SendMailMethod: Integer;
    procedure Set_SendMailMethod(SendMailMethod: Integer);
    function Get_ShowAnimation: Integer;
    procedure Set_ShowAnimation(ShowAnimation: Integer);
    function Get_StampFontColor: WideString;
    procedure Set_StampFontColor(const StampFontColor: WideString);
    function Get_StampFontname: WideString;
    procedure Set_StampFontname(const StampFontname: WideString);
    function Get_StampFontsize: Integer;
    procedure Set_StampFontsize(StampFontsize: Integer);
    function Get_StampOutlineFontthickness: Integer;
    procedure Set_StampOutlineFontthickness(StampOutlineFontthickness: Integer);
    function Get_StampString: WideString;
    procedure Set_StampString(const StampString: WideString);
    function Get_StampUseOutlineFont: Integer;
    procedure Set_StampUseOutlineFont(StampUseOutlineFont: Integer);
    function Get_StandardAuthor: WideString;
    procedure Set_StandardAuthor(const StandardAuthor: WideString);
    function Get_StandardCreationdate: WideString;
    procedure Set_StandardCreationdate(const StandardCreationdate: WideString);
    function Get_StandardDateformat: WideString;
    procedure Set_StandardDateformat(const StandardDateformat: WideString);
    function Get_StandardKeywords: WideString;
    procedure Set_StandardKeywords(const StandardKeywords: WideString);
    function Get_StandardMailDomain: WideString;
    procedure Set_StandardMailDomain(const StandardMailDomain: WideString);
    function Get_StandardModifydate: WideString;
    procedure Set_StandardModifydate(const StandardModifydate: WideString);
    function Get_StandardSaveformat: Integer;
    procedure Set_StandardSaveformat(StandardSaveformat: Integer);
    function Get_StandardSubject: WideString;
    procedure Set_StandardSubject(const StandardSubject: WideString);
    function Get_StandardTitle: WideString;
    procedure Set_StandardTitle(const StandardTitle: WideString);
    function Get_StartStandardProgram: Integer;
    procedure Set_StartStandardProgram(StartStandardProgram: Integer);
    function Get_SVGResolution: Integer;
    procedure Set_SVGResolution(SVGResolution: Integer);
    function Get_TIFFColorscount: Integer;
    procedure Set_TIFFColorscount(TIFFColorscount: Integer);
    function Get_TIFFResolution: Integer;
    procedure Set_TIFFResolution(TIFFResolution: Integer);
    function Get_Toolbars: Integer;
    procedure Set_Toolbars(Toolbars: Integer);
    function Get_UpdateInterval: Integer;
    procedure Set_UpdateInterval(UpdateInterval: Integer);
    function Get_UseAutosave: Integer;
    procedure Set_UseAutosave(UseAutosave: Integer);
    function Get_UseAutosaveDirectory: Integer;
    procedure Set_UseAutosaveDirectory(UseAutosaveDirectory: Integer);
    function Get_UseCreationDateNow: Integer;
    procedure Set_UseCreationDateNow(UseCreationDateNow: Integer);
    function Get_UseCustomPaperSize: WideString;
    procedure Set_UseCustomPaperSize(const UseCustomPaperSize: WideString);
    function Get_UseFixPapersize: Integer;
    procedure Set_UseFixPapersize(UseFixPapersize: Integer);
    function Get_UseStandardAuthor: Integer;
    procedure Set_UseStandardAuthor(UseStandardAuthor: Integer);
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: _clsPDFCreatorOptions);
    procedure Disconnect; override;
    property DefaultInterface: _clsPDFCreatorOptions read GetDefaultInterface;
    property AdditionalGhostscriptParameters: WideString read Get_AdditionalGhostscriptParameters write Set_AdditionalGhostscriptParameters;
    property AdditionalGhostscriptSearchpath: WideString read Get_AdditionalGhostscriptSearchpath write Set_AdditionalGhostscriptSearchpath;
    property AddWindowsFontpath: Integer read Get_AddWindowsFontpath write Set_AddWindowsFontpath;
    property AllowSpecialGSCharsInFilenames: Integer read Get_AllowSpecialGSCharsInFilenames write Set_AllowSpecialGSCharsInFilenames;
    property AutosaveDirectory: WideString read Get_AutosaveDirectory write Set_AutosaveDirectory;
    property AutosaveFilename: WideString read Get_AutosaveFilename write Set_AutosaveFilename;
    property AutosaveFormat: Integer read Get_AutosaveFormat write Set_AutosaveFormat;
    property AutosaveStartStandardProgram: Integer read Get_AutosaveStartStandardProgram write Set_AutosaveStartStandardProgram;
    property BMPColorscount: Integer read Get_BMPColorscount write Set_BMPColorscount;
    property BMPResolution: Integer read Get_BMPResolution write Set_BMPResolution;
    property ClientComputerResolveIPAddress: Integer read Get_ClientComputerResolveIPAddress write Set_ClientComputerResolveIPAddress;
    property Counter: Currency read Get_Counter write Set_Counter;
    property DeviceHeightPoints: Double read Get_DeviceHeightPoints write Set_DeviceHeightPoints;
    property DeviceWidthPoints: Double read Get_DeviceWidthPoints write Set_DeviceWidthPoints;
    property DirectoryGhostscriptBinaries: WideString read Get_DirectoryGhostscriptBinaries write Set_DirectoryGhostscriptBinaries;
    property DirectoryGhostscriptFonts: WideString read Get_DirectoryGhostscriptFonts write Set_DirectoryGhostscriptFonts;
    property DirectoryGhostscriptLibraries: WideString read Get_DirectoryGhostscriptLibraries write Set_DirectoryGhostscriptLibraries;
    property DirectoryGhostscriptResource: WideString read Get_DirectoryGhostscriptResource write Set_DirectoryGhostscriptResource;
    property DisableEmail: Integer read Get_DisableEmail write Set_DisableEmail;
    property DontUseDocumentSettings: Integer read Get_DontUseDocumentSettings write Set_DontUseDocumentSettings;
    property EPSLanguageLevel: Integer read Get_EPSLanguageLevel write Set_EPSLanguageLevel;
    property FilenameSubstitutions: WideString read Get_FilenameSubstitutions write Set_FilenameSubstitutions;
    property FilenameSubstitutionsOnlyInTitle: Integer read Get_FilenameSubstitutionsOnlyInTitle write Set_FilenameSubstitutionsOnlyInTitle;
    property JPEGColorscount: Integer read Get_JPEGColorscount write Set_JPEGColorscount;
    property JPEGQuality: Integer read Get_JPEGQuality write Set_JPEGQuality;
    property JPEGResolution: Integer read Get_JPEGResolution write Set_JPEGResolution;
    property Language: WideString read Get_Language write Set_Language;
    property LastSaveDirectory: WideString read Get_LastSaveDirectory write Set_LastSaveDirectory;
    property LastUpdateCheck: WideString read Get_LastUpdateCheck write Set_LastUpdateCheck;
    property Logging: Integer read Get_Logging write Set_Logging;
    property LogLines: Integer read Get_LogLines write Set_LogLines;
    property NoConfirmMessageSwitchingDefaultprinter: Integer read Get_NoConfirmMessageSwitchingDefaultprinter write Set_NoConfirmMessageSwitchingDefaultprinter;
    property NoProcessingAtStartup: Integer read Get_NoProcessingAtStartup write Set_NoProcessingAtStartup;
    property NoPSCheck: Integer read Get_NoPSCheck write Set_NoPSCheck;
    property OnePagePerFile: Integer read Get_OnePagePerFile write Set_OnePagePerFile;
    property OptionsDesign: Integer read Get_OptionsDesign write Set_OptionsDesign;
    property OptionsEnabled: Integer read Get_OptionsEnabled write Set_OptionsEnabled;
    property OptionsVisible: Integer read Get_OptionsVisible write Set_OptionsVisible;
    property Papersize: WideString read Get_Papersize write Set_Papersize;
    property PCLColorsCount: Integer read Get_PCLColorsCount write Set_PCLColorsCount;
    property PCLResolution: Integer read Get_PCLResolution write Set_PCLResolution;
    property PCXColorscount: Integer read Get_PCXColorscount write Set_PCXColorscount;
    property PCXResolution: Integer read Get_PCXResolution write Set_PCXResolution;
    property PDFAes128Encryption: Integer read Get_PDFAes128Encryption write Set_PDFAes128Encryption;
    property PDFAllowAssembly: Integer read Get_PDFAllowAssembly write Set_PDFAllowAssembly;
    property PDFAllowDegradedPrinting: Integer read Get_PDFAllowDegradedPrinting write Set_PDFAllowDegradedPrinting;
    property PDFAllowFillIn: Integer read Get_PDFAllowFillIn write Set_PDFAllowFillIn;
    property PDFAllowScreenReaders: Integer read Get_PDFAllowScreenReaders write Set_PDFAllowScreenReaders;
    property PDFColorsCMYKToRGB: Integer read Get_PDFColorsCMYKToRGB write Set_PDFColorsCMYKToRGB;
    property PDFColorsColorModel: Integer read Get_PDFColorsColorModel write Set_PDFColorsColorModel;
    property PDFColorsPreserveHalftone: Integer read Get_PDFColorsPreserveHalftone write Set_PDFColorsPreserveHalftone;
    property PDFColorsPreserveOverprint: Integer read Get_PDFColorsPreserveOverprint write Set_PDFColorsPreserveOverprint;
    property PDFColorsPreserveTransfer: Integer read Get_PDFColorsPreserveTransfer write Set_PDFColorsPreserveTransfer;
    property PDFCompressionColorCompression: Integer read Get_PDFCompressionColorCompression write Set_PDFCompressionColorCompression;
    property PDFCompressionColorCompressionChoice: Integer read Get_PDFCompressionColorCompressionChoice write Set_PDFCompressionColorCompressionChoice;
    property PDFCompressionColorCompressionJPEGHighFactor: Double read Get_PDFCompressionColorCompressionJPEGHighFactor write Set_PDFCompressionColorCompressionJPEGHighFactor;
    property PDFCompressionColorCompressionJPEGLowFactor: Double read Get_PDFCompressionColorCompressionJPEGLowFactor write Set_PDFCompressionColorCompressionJPEGLowFactor;
    property PDFCompressionColorCompressionJPEGManualFactor: Double read Get_PDFCompressionColorCompressionJPEGManualFactor write Set_PDFCompressionColorCompressionJPEGManualFactor;
    property PDFCompressionColorCompressionJPEGMaximumFactor: Double read Get_PDFCompressionColorCompressionJPEGMaximumFactor write Set_PDFCompressionColorCompressionJPEGMaximumFactor;
    property PDFCompressionColorCompressionJPEGMediumFactor: Double read Get_PDFCompressionColorCompressionJPEGMediumFactor write Set_PDFCompressionColorCompressionJPEGMediumFactor;
    property PDFCompressionColorCompressionJPEGMinimumFactor: Double read Get_PDFCompressionColorCompressionJPEGMinimumFactor write Set_PDFCompressionColorCompressionJPEGMinimumFactor;
    property PDFCompressionColorResample: Integer read Get_PDFCompressionColorResample write Set_PDFCompressionColorResample;
    property PDFCompressionColorResampleChoice: Integer read Get_PDFCompressionColorResampleChoice write Set_PDFCompressionColorResampleChoice;
    property PDFCompressionColorResolution: Integer read Get_PDFCompressionColorResolution write Set_PDFCompressionColorResolution;
    property PDFCompressionGreyCompression: Integer read Get_PDFCompressionGreyCompression write Set_PDFCompressionGreyCompression;
    property PDFCompressionGreyCompressionChoice: Integer read Get_PDFCompressionGreyCompressionChoice write Set_PDFCompressionGreyCompressionChoice;
    property PDFCompressionGreyCompressionJPEGHighFactor: Double read Get_PDFCompressionGreyCompressionJPEGHighFactor write Set_PDFCompressionGreyCompressionJPEGHighFactor;
    property PDFCompressionGreyCompressionJPEGLowFactor: Double read Get_PDFCompressionGreyCompressionJPEGLowFactor write Set_PDFCompressionGreyCompressionJPEGLowFactor;
    property PDFCompressionGreyCompressionJPEGManualFactor: Double read Get_PDFCompressionGreyCompressionJPEGManualFactor write Set_PDFCompressionGreyCompressionJPEGManualFactor;
    property PDFCompressionGreyCompressionJPEGMaximumFactor: Double read Get_PDFCompressionGreyCompressionJPEGMaximumFactor write Set_PDFCompressionGreyCompressionJPEGMaximumFactor;
    property PDFCompressionGreyCompressionJPEGMediumFactor: Double read Get_PDFCompressionGreyCompressionJPEGMediumFactor write Set_PDFCompressionGreyCompressionJPEGMediumFactor;
    property PDFCompressionGreyCompressionJPEGMinimumFactor: Double read Get_PDFCompressionGreyCompressionJPEGMinimumFactor write Set_PDFCompressionGreyCompressionJPEGMinimumFactor;
    property PDFCompressionGreyResample: Integer read Get_PDFCompressionGreyResample write Set_PDFCompressionGreyResample;
    property PDFCompressionGreyResampleChoice: Integer read Get_PDFCompressionGreyResampleChoice write Set_PDFCompressionGreyResampleChoice;
    property PDFCompressionGreyResolution: Integer read Get_PDFCompressionGreyResolution write Set_PDFCompressionGreyResolution;
    property PDFCompressionMonoCompression: Integer read Get_PDFCompressionMonoCompression write Set_PDFCompressionMonoCompression;
    property PDFCompressionMonoCompressionChoice: Integer read Get_PDFCompressionMonoCompressionChoice write Set_PDFCompressionMonoCompressionChoice;
    property PDFCompressionMonoResample: Integer read Get_PDFCompressionMonoResample write Set_PDFCompressionMonoResample;
    property PDFCompressionMonoResampleChoice: Integer read Get_PDFCompressionMonoResampleChoice write Set_PDFCompressionMonoResampleChoice;
    property PDFCompressionMonoResolution: Integer read Get_PDFCompressionMonoResolution write Set_PDFCompressionMonoResolution;
    property PDFCompressionTextCompression: Integer read Get_PDFCompressionTextCompression write Set_PDFCompressionTextCompression;
    property PDFDisallowCopy: Integer read Get_PDFDisallowCopy write Set_PDFDisallowCopy;
    property PDFDisallowModifyAnnotations: Integer read Get_PDFDisallowModifyAnnotations write Set_PDFDisallowModifyAnnotations;
    property PDFDisallowModifyContents: Integer read Get_PDFDisallowModifyContents write Set_PDFDisallowModifyContents;
    property PDFDisallowPrinting: Integer read Get_PDFDisallowPrinting write Set_PDFDisallowPrinting;
    property PDFEncryptor: Integer read Get_PDFEncryptor write Set_PDFEncryptor;
    property PDFFontsEmbedAll: Integer read Get_PDFFontsEmbedAll write Set_PDFFontsEmbedAll;
    property PDFFontsSubSetFonts: Integer read Get_PDFFontsSubSetFonts write Set_PDFFontsSubSetFonts;
    property PDFFontsSubSetFontsPercent: Integer read Get_PDFFontsSubSetFontsPercent write Set_PDFFontsSubSetFontsPercent;
    property PDFGeneralASCII85: Integer read Get_PDFGeneralASCII85 write Set_PDFGeneralASCII85;
    property PDFGeneralAutorotate: Integer read Get_PDFGeneralAutorotate write Set_PDFGeneralAutorotate;
    property PDFGeneralCompatibility: Integer read Get_PDFGeneralCompatibility write Set_PDFGeneralCompatibility;
    property PDFGeneralDefault: Integer read Get_PDFGeneralDefault write Set_PDFGeneralDefault;
    property PDFGeneralOverprint: Integer read Get_PDFGeneralOverprint write Set_PDFGeneralOverprint;
    property PDFGeneralResolution: Integer read Get_PDFGeneralResolution write Set_PDFGeneralResolution;
    property PDFHighEncryption: Integer read Get_PDFHighEncryption write Set_PDFHighEncryption;
    property PDFLowEncryption: Integer read Get_PDFLowEncryption write Set_PDFLowEncryption;
    property PDFOptimize: Integer read Get_PDFOptimize write Set_PDFOptimize;
    property PDFOwnerPass: Integer read Get_PDFOwnerPass write Set_PDFOwnerPass;
    property PDFOwnerPasswordString: WideString read Get_PDFOwnerPasswordString write Set_PDFOwnerPasswordString;
    property PDFSigningMultiSignature: Integer read Get_PDFSigningMultiSignature write Set_PDFSigningMultiSignature;
    property PDFSigningPFXFile: WideString read Get_PDFSigningPFXFile write Set_PDFSigningPFXFile;
    property PDFSigningPFXFilePassword: WideString read Get_PDFSigningPFXFilePassword write Set_PDFSigningPFXFilePassword;
    property PDFSigningSignatureContact: WideString read Get_PDFSigningSignatureContact write Set_PDFSigningSignatureContact;
    property PDFSigningSignatureLeftX: Double read Get_PDFSigningSignatureLeftX write Set_PDFSigningSignatureLeftX;
    property PDFSigningSignatureLeftY: Double read Get_PDFSigningSignatureLeftY write Set_PDFSigningSignatureLeftY;
    property PDFSigningSignatureLocation: WideString read Get_PDFSigningSignatureLocation write Set_PDFSigningSignatureLocation;
    property PDFSigningSignatureOnPage: Integer read Get_PDFSigningSignatureOnPage write Set_PDFSigningSignatureOnPage;
    property PDFSigningSignatureReason: WideString read Get_PDFSigningSignatureReason write Set_PDFSigningSignatureReason;
    property PDFSigningSignatureRightX: Double read Get_PDFSigningSignatureRightX write Set_PDFSigningSignatureRightX;
    property PDFSigningSignatureRightY: Double read Get_PDFSigningSignatureRightY write Set_PDFSigningSignatureRightY;
    property PDFSigningSignatureVisible: Integer read Get_PDFSigningSignatureVisible write Set_PDFSigningSignatureVisible;
    property PDFSigningSignPDF: Integer read Get_PDFSigningSignPDF write Set_PDFSigningSignPDF;
    property PDFUpdateMetadata: Integer read Get_PDFUpdateMetadata write Set_PDFUpdateMetadata;
    property PDFUserPass: Integer read Get_PDFUserPass write Set_PDFUserPass;
    property PDFUserPasswordString: WideString read Get_PDFUserPasswordString write Set_PDFUserPasswordString;
    property PDFUseSecurity: Integer read Get_PDFUseSecurity write Set_PDFUseSecurity;
    property PNGColorscount: Integer read Get_PNGColorscount write Set_PNGColorscount;
    property PNGResolution: Integer read Get_PNGResolution write Set_PNGResolution;
    property PrintAfterSaving: Integer read Get_PrintAfterSaving write Set_PrintAfterSaving;
    property PrintAfterSavingBitsPerPixel: Integer read Get_PrintAfterSavingBitsPerPixel write Set_PrintAfterSavingBitsPerPixel;
    property PrintAfterSavingDuplex: Integer read Get_PrintAfterSavingDuplex write Set_PrintAfterSavingDuplex;
    property PrintAfterSavingMaxResolution: Integer read Get_PrintAfterSavingMaxResolution write Set_PrintAfterSavingMaxResolution;
    property PrintAfterSavingMaxResolutionEnabled: Integer read Get_PrintAfterSavingMaxResolutionEnabled write Set_PrintAfterSavingMaxResolutionEnabled;
    property PrintAfterSavingNoCancel: Integer read Get_PrintAfterSavingNoCancel write Set_PrintAfterSavingNoCancel;
    property PrintAfterSavingPrinter: WideString read Get_PrintAfterSavingPrinter write Set_PrintAfterSavingPrinter;
    property PrintAfterSavingQueryUser: Integer read Get_PrintAfterSavingQueryUser write Set_PrintAfterSavingQueryUser;
    property PrintAfterSavingTumble: Integer read Get_PrintAfterSavingTumble write Set_PrintAfterSavingTumble;
    property PrinterStop: Integer read Get_PrinterStop write Set_PrinterStop;
    property PrinterTemppath: WideString read Get_PrinterTemppath write Set_PrinterTemppath;
    property ProcessPriority: Integer read Get_ProcessPriority write Set_ProcessPriority;
    property ProgramFont: WideString read Get_ProgramFont write Set_ProgramFont;
    property ProgramFontCharset: Integer read Get_ProgramFontCharset write Set_ProgramFontCharset;
    property ProgramFontSize: Integer read Get_ProgramFontSize write Set_ProgramFontSize;
    property PSDColorsCount: Integer read Get_PSDColorsCount write Set_PSDColorsCount;
    property PSDResolution: Integer read Get_PSDResolution write Set_PSDResolution;
    property PSLanguageLevel: Integer read Get_PSLanguageLevel write Set_PSLanguageLevel;
    property RAWColorsCount: Integer read Get_RAWColorsCount write Set_RAWColorsCount;
    property RAWResolution: Integer read Get_RAWResolution write Set_RAWResolution;
    property RemoveAllKnownFileExtensions: Integer read Get_RemoveAllKnownFileExtensions write Set_RemoveAllKnownFileExtensions;
    property RemoveSpaces: Integer read Get_RemoveSpaces write Set_RemoveSpaces;
    property RunProgramAfterSaving: Integer read Get_RunProgramAfterSaving write Set_RunProgramAfterSaving;
    property RunProgramAfterSavingProgramname: WideString read Get_RunProgramAfterSavingProgramname write Set_RunProgramAfterSavingProgramname;
    property RunProgramAfterSavingProgramParameters: WideString read Get_RunProgramAfterSavingProgramParameters write Set_RunProgramAfterSavingProgramParameters;
    property RunProgramAfterSavingWaitUntilReady: Integer read Get_RunProgramAfterSavingWaitUntilReady write Set_RunProgramAfterSavingWaitUntilReady;
    property RunProgramAfterSavingWindowstyle: Integer read Get_RunProgramAfterSavingWindowstyle write Set_RunProgramAfterSavingWindowstyle;
    property RunProgramBeforeSaving: Integer read Get_RunProgramBeforeSaving write Set_RunProgramBeforeSaving;
    property RunProgramBeforeSavingProgramname: WideString read Get_RunProgramBeforeSavingProgramname write Set_RunProgramBeforeSavingProgramname;
    property RunProgramBeforeSavingProgramParameters: WideString read Get_RunProgramBeforeSavingProgramParameters write Set_RunProgramBeforeSavingProgramParameters;
    property RunProgramBeforeSavingWindowstyle: Integer read Get_RunProgramBeforeSavingWindowstyle write Set_RunProgramBeforeSavingWindowstyle;
    property SaveFilename: WideString read Get_SaveFilename write Set_SaveFilename;
    property SendEmailAfterAutoSaving: Integer read Get_SendEmailAfterAutoSaving write Set_SendEmailAfterAutoSaving;
    property SendMailMethod: Integer read Get_SendMailMethod write Set_SendMailMethod;
    property ShowAnimation: Integer read Get_ShowAnimation write Set_ShowAnimation;
    property StampFontColor: WideString read Get_StampFontColor write Set_StampFontColor;
    property StampFontname: WideString read Get_StampFontname write Set_StampFontname;
    property StampFontsize: Integer read Get_StampFontsize write Set_StampFontsize;
    property StampOutlineFontthickness: Integer read Get_StampOutlineFontthickness write Set_StampOutlineFontthickness;
    property StampString: WideString read Get_StampString write Set_StampString;
    property StampUseOutlineFont: Integer read Get_StampUseOutlineFont write Set_StampUseOutlineFont;
    property StandardAuthor: WideString read Get_StandardAuthor write Set_StandardAuthor;
    property StandardCreationdate: WideString read Get_StandardCreationdate write Set_StandardCreationdate;
    property StandardDateformat: WideString read Get_StandardDateformat write Set_StandardDateformat;
    property StandardKeywords: WideString read Get_StandardKeywords write Set_StandardKeywords;
    property StandardMailDomain: WideString read Get_StandardMailDomain write Set_StandardMailDomain;
    property StandardModifydate: WideString read Get_StandardModifydate write Set_StandardModifydate;
    property StandardSaveformat: Integer read Get_StandardSaveformat write Set_StandardSaveformat;
    property StandardSubject: WideString read Get_StandardSubject write Set_StandardSubject;
    property StandardTitle: WideString read Get_StandardTitle write Set_StandardTitle;
    property StartStandardProgram: Integer read Get_StartStandardProgram write Set_StartStandardProgram;
    property SVGResolution: Integer read Get_SVGResolution write Set_SVGResolution;
    property TIFFColorscount: Integer read Get_TIFFColorscount write Set_TIFFColorscount;
    property TIFFResolution: Integer read Get_TIFFResolution write Set_TIFFResolution;
    property Toolbars: Integer read Get_Toolbars write Set_Toolbars;
    property UpdateInterval: Integer read Get_UpdateInterval write Set_UpdateInterval;
    property UseAutosave: Integer read Get_UseAutosave write Set_UseAutosave;
    property UseAutosaveDirectory: Integer read Get_UseAutosaveDirectory write Set_UseAutosaveDirectory;
    property UseCreationDateNow: Integer read Get_UseCreationDateNow write Set_UseCreationDateNow;
    property UseCustomPaperSize: WideString read Get_UseCustomPaperSize write Set_UseCustomPaperSize;
    property UseFixPapersize: Integer read Get_UseFixPapersize write Set_UseFixPapersize;
    property UseStandardAuthor: Integer read Get_UseStandardAuthor write Set_UseStandardAuthor;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TclsPDFCreatorOptionsProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// Classe proxy Propriétés Serveur OLE
// Objet serveur    : TclsPDFCreatorOptions
// (Cet objet est utilisé pas l'inspecteur de propriété de l'EDI pour la 
//  modification des propriétés de ce serveur)
// *********************************************************************//
 TclsPDFCreatorOptionsProperties = class(TPersistent)
  private
    FServer:    TclsPDFCreatorOptions;
    function    GetDefaultInterface: _clsPDFCreatorOptions;
    constructor Create(AServer: TclsPDFCreatorOptions);
  protected
    function Get_AdditionalGhostscriptParameters: WideString;
    procedure Set_AdditionalGhostscriptParameters(const AdditionalGhostscriptParameters: WideString);
    function Get_AdditionalGhostscriptSearchpath: WideString;
    procedure Set_AdditionalGhostscriptSearchpath(const AdditionalGhostscriptSearchpath: WideString);
    function Get_AddWindowsFontpath: Integer;
    procedure Set_AddWindowsFontpath(AddWindowsFontpath: Integer);
    function Get_AllowSpecialGSCharsInFilenames: Integer;
    procedure Set_AllowSpecialGSCharsInFilenames(AllowSpecialGSCharsInFilenames: Integer);
    function Get_AutosaveDirectory: WideString;
    procedure Set_AutosaveDirectory(const AutosaveDirectory: WideString);
    function Get_AutosaveFilename: WideString;
    procedure Set_AutosaveFilename(const AutosaveFilename: WideString);
    function Get_AutosaveFormat: Integer;
    procedure Set_AutosaveFormat(AutosaveFormat: Integer);
    function Get_AutosaveStartStandardProgram: Integer;
    procedure Set_AutosaveStartStandardProgram(AutosaveStartStandardProgram: Integer);
    function Get_BMPColorscount: Integer;
    procedure Set_BMPColorscount(BMPColorscount: Integer);
    function Get_BMPResolution: Integer;
    procedure Set_BMPResolution(BMPResolution: Integer);
    function Get_ClientComputerResolveIPAddress: Integer;
    procedure Set_ClientComputerResolveIPAddress(ClientComputerResolveIPAddress: Integer);
    function Get_Counter: Currency;
    procedure Set_Counter(Counter: Currency);
    function Get_DeviceHeightPoints: Double;
    procedure Set_DeviceHeightPoints(DeviceHeightPoints: Double);
    function Get_DeviceWidthPoints: Double;
    procedure Set_DeviceWidthPoints(DeviceWidthPoints: Double);
    function Get_DirectoryGhostscriptBinaries: WideString;
    procedure Set_DirectoryGhostscriptBinaries(const DirectoryGhostscriptBinaries: WideString);
    function Get_DirectoryGhostscriptFonts: WideString;
    procedure Set_DirectoryGhostscriptFonts(const DirectoryGhostscriptFonts: WideString);
    function Get_DirectoryGhostscriptLibraries: WideString;
    procedure Set_DirectoryGhostscriptLibraries(const DirectoryGhostscriptLibraries: WideString);
    function Get_DirectoryGhostscriptResource: WideString;
    procedure Set_DirectoryGhostscriptResource(const DirectoryGhostscriptResource: WideString);
    function Get_DisableEmail: Integer;
    procedure Set_DisableEmail(DisableEmail: Integer);
    function Get_DontUseDocumentSettings: Integer;
    procedure Set_DontUseDocumentSettings(DontUseDocumentSettings: Integer);
    function Get_EPSLanguageLevel: Integer;
    procedure Set_EPSLanguageLevel(EPSLanguageLevel: Integer);
    function Get_FilenameSubstitutions: WideString;
    procedure Set_FilenameSubstitutions(const FilenameSubstitutions: WideString);
    function Get_FilenameSubstitutionsOnlyInTitle: Integer;
    procedure Set_FilenameSubstitutionsOnlyInTitle(FilenameSubstitutionsOnlyInTitle: Integer);
    function Get_JPEGColorscount: Integer;
    procedure Set_JPEGColorscount(JPEGColorscount: Integer);
    function Get_JPEGQuality: Integer;
    procedure Set_JPEGQuality(JPEGQuality: Integer);
    function Get_JPEGResolution: Integer;
    procedure Set_JPEGResolution(JPEGResolution: Integer);
    function Get_Language: WideString;
    procedure Set_Language(const Language: WideString);
    function Get_LastSaveDirectory: WideString;
    procedure Set_LastSaveDirectory(const LastSaveDirectory: WideString);
    function Get_LastUpdateCheck: WideString;
    procedure Set_LastUpdateCheck(const LastUpdateCheck: WideString);
    function Get_Logging: Integer;
    procedure Set_Logging(Logging: Integer);
    function Get_LogLines: Integer;
    procedure Set_LogLines(LogLines: Integer);
    function Get_NoConfirmMessageSwitchingDefaultprinter: Integer;
    procedure Set_NoConfirmMessageSwitchingDefaultprinter(NoConfirmMessageSwitchingDefaultprinter: Integer);
    function Get_NoProcessingAtStartup: Integer;
    procedure Set_NoProcessingAtStartup(NoProcessingAtStartup: Integer);
    function Get_NoPSCheck: Integer;
    procedure Set_NoPSCheck(NoPSCheck: Integer);
    function Get_OnePagePerFile: Integer;
    procedure Set_OnePagePerFile(OnePagePerFile: Integer);
    function Get_OptionsDesign: Integer;
    procedure Set_OptionsDesign(OptionsDesign: Integer);
    function Get_OptionsEnabled: Integer;
    procedure Set_OptionsEnabled(OptionsEnabled: Integer);
    function Get_OptionsVisible: Integer;
    procedure Set_OptionsVisible(OptionsVisible: Integer);
    function Get_Papersize: WideString;
    procedure Set_Papersize(const Papersize: WideString);
    function Get_PCLColorsCount: Integer;
    procedure Set_PCLColorsCount(PCLColorsCount: Integer);
    function Get_PCLResolution: Integer;
    procedure Set_PCLResolution(PCLResolution: Integer);
    function Get_PCXColorscount: Integer;
    procedure Set_PCXColorscount(PCXColorscount: Integer);
    function Get_PCXResolution: Integer;
    procedure Set_PCXResolution(PCXResolution: Integer);
    function Get_PDFAes128Encryption: Integer;
    procedure Set_PDFAes128Encryption(PDFAes128Encryption: Integer);
    function Get_PDFAllowAssembly: Integer;
    procedure Set_PDFAllowAssembly(PDFAllowAssembly: Integer);
    function Get_PDFAllowDegradedPrinting: Integer;
    procedure Set_PDFAllowDegradedPrinting(PDFAllowDegradedPrinting: Integer);
    function Get_PDFAllowFillIn: Integer;
    procedure Set_PDFAllowFillIn(PDFAllowFillIn: Integer);
    function Get_PDFAllowScreenReaders: Integer;
    procedure Set_PDFAllowScreenReaders(PDFAllowScreenReaders: Integer);
    function Get_PDFColorsCMYKToRGB: Integer;
    procedure Set_PDFColorsCMYKToRGB(PDFColorsCMYKToRGB: Integer);
    function Get_PDFColorsColorModel: Integer;
    procedure Set_PDFColorsColorModel(PDFColorsColorModel: Integer);
    function Get_PDFColorsPreserveHalftone: Integer;
    procedure Set_PDFColorsPreserveHalftone(PDFColorsPreserveHalftone: Integer);
    function Get_PDFColorsPreserveOverprint: Integer;
    procedure Set_PDFColorsPreserveOverprint(PDFColorsPreserveOverprint: Integer);
    function Get_PDFColorsPreserveTransfer: Integer;
    procedure Set_PDFColorsPreserveTransfer(PDFColorsPreserveTransfer: Integer);
    function Get_PDFCompressionColorCompression: Integer;
    procedure Set_PDFCompressionColorCompression(PDFCompressionColorCompression: Integer);
    function Get_PDFCompressionColorCompressionChoice: Integer;
    procedure Set_PDFCompressionColorCompressionChoice(PDFCompressionColorCompressionChoice: Integer);
    function Get_PDFCompressionColorCompressionJPEGHighFactor: Double;
    procedure Set_PDFCompressionColorCompressionJPEGHighFactor(PDFCompressionColorCompressionJPEGHighFactor: Double);
    function Get_PDFCompressionColorCompressionJPEGLowFactor: Double;
    procedure Set_PDFCompressionColorCompressionJPEGLowFactor(PDFCompressionColorCompressionJPEGLowFactor: Double);
    function Get_PDFCompressionColorCompressionJPEGManualFactor: Double;
    procedure Set_PDFCompressionColorCompressionJPEGManualFactor(PDFCompressionColorCompressionJPEGManualFactor: Double);
    function Get_PDFCompressionColorCompressionJPEGMaximumFactor: Double;
    procedure Set_PDFCompressionColorCompressionJPEGMaximumFactor(PDFCompressionColorCompressionJPEGMaximumFactor: Double);
    function Get_PDFCompressionColorCompressionJPEGMediumFactor: Double;
    procedure Set_PDFCompressionColorCompressionJPEGMediumFactor(PDFCompressionColorCompressionJPEGMediumFactor: Double);
    function Get_PDFCompressionColorCompressionJPEGMinimumFactor: Double;
    procedure Set_PDFCompressionColorCompressionJPEGMinimumFactor(PDFCompressionColorCompressionJPEGMinimumFactor: Double);
    function Get_PDFCompressionColorResample: Integer;
    procedure Set_PDFCompressionColorResample(PDFCompressionColorResample: Integer);
    function Get_PDFCompressionColorResampleChoice: Integer;
    procedure Set_PDFCompressionColorResampleChoice(PDFCompressionColorResampleChoice: Integer);
    function Get_PDFCompressionColorResolution: Integer;
    procedure Set_PDFCompressionColorResolution(PDFCompressionColorResolution: Integer);
    function Get_PDFCompressionGreyCompression: Integer;
    procedure Set_PDFCompressionGreyCompression(PDFCompressionGreyCompression: Integer);
    function Get_PDFCompressionGreyCompressionChoice: Integer;
    procedure Set_PDFCompressionGreyCompressionChoice(PDFCompressionGreyCompressionChoice: Integer);
    function Get_PDFCompressionGreyCompressionJPEGHighFactor: Double;
    procedure Set_PDFCompressionGreyCompressionJPEGHighFactor(PDFCompressionGreyCompressionJPEGHighFactor: Double);
    function Get_PDFCompressionGreyCompressionJPEGLowFactor: Double;
    procedure Set_PDFCompressionGreyCompressionJPEGLowFactor(PDFCompressionGreyCompressionJPEGLowFactor: Double);
    function Get_PDFCompressionGreyCompressionJPEGManualFactor: Double;
    procedure Set_PDFCompressionGreyCompressionJPEGManualFactor(PDFCompressionGreyCompressionJPEGManualFactor: Double);
    function Get_PDFCompressionGreyCompressionJPEGMaximumFactor: Double;
    procedure Set_PDFCompressionGreyCompressionJPEGMaximumFactor(PDFCompressionGreyCompressionJPEGMaximumFactor: Double);
    function Get_PDFCompressionGreyCompressionJPEGMediumFactor: Double;
    procedure Set_PDFCompressionGreyCompressionJPEGMediumFactor(PDFCompressionGreyCompressionJPEGMediumFactor: Double);
    function Get_PDFCompressionGreyCompressionJPEGMinimumFactor: Double;
    procedure Set_PDFCompressionGreyCompressionJPEGMinimumFactor(PDFCompressionGreyCompressionJPEGMinimumFactor: Double);
    function Get_PDFCompressionGreyResample: Integer;
    procedure Set_PDFCompressionGreyResample(PDFCompressionGreyResample: Integer);
    function Get_PDFCompressionGreyResampleChoice: Integer;
    procedure Set_PDFCompressionGreyResampleChoice(PDFCompressionGreyResampleChoice: Integer);
    function Get_PDFCompressionGreyResolution: Integer;
    procedure Set_PDFCompressionGreyResolution(PDFCompressionGreyResolution: Integer);
    function Get_PDFCompressionMonoCompression: Integer;
    procedure Set_PDFCompressionMonoCompression(PDFCompressionMonoCompression: Integer);
    function Get_PDFCompressionMonoCompressionChoice: Integer;
    procedure Set_PDFCompressionMonoCompressionChoice(PDFCompressionMonoCompressionChoice: Integer);
    function Get_PDFCompressionMonoResample: Integer;
    procedure Set_PDFCompressionMonoResample(PDFCompressionMonoResample: Integer);
    function Get_PDFCompressionMonoResampleChoice: Integer;
    procedure Set_PDFCompressionMonoResampleChoice(PDFCompressionMonoResampleChoice: Integer);
    function Get_PDFCompressionMonoResolution: Integer;
    procedure Set_PDFCompressionMonoResolution(PDFCompressionMonoResolution: Integer);
    function Get_PDFCompressionTextCompression: Integer;
    procedure Set_PDFCompressionTextCompression(PDFCompressionTextCompression: Integer);
    function Get_PDFDisallowCopy: Integer;
    procedure Set_PDFDisallowCopy(PDFDisallowCopy: Integer);
    function Get_PDFDisallowModifyAnnotations: Integer;
    procedure Set_PDFDisallowModifyAnnotations(PDFDisallowModifyAnnotations: Integer);
    function Get_PDFDisallowModifyContents: Integer;
    procedure Set_PDFDisallowModifyContents(PDFDisallowModifyContents: Integer);
    function Get_PDFDisallowPrinting: Integer;
    procedure Set_PDFDisallowPrinting(PDFDisallowPrinting: Integer);
    function Get_PDFEncryptor: Integer;
    procedure Set_PDFEncryptor(PDFEncryptor: Integer);
    function Get_PDFFontsEmbedAll: Integer;
    procedure Set_PDFFontsEmbedAll(PDFFontsEmbedAll: Integer);
    function Get_PDFFontsSubSetFonts: Integer;
    procedure Set_PDFFontsSubSetFonts(PDFFontsSubSetFonts: Integer);
    function Get_PDFFontsSubSetFontsPercent: Integer;
    procedure Set_PDFFontsSubSetFontsPercent(PDFFontsSubSetFontsPercent: Integer);
    function Get_PDFGeneralASCII85: Integer;
    procedure Set_PDFGeneralASCII85(PDFGeneralASCII85: Integer);
    function Get_PDFGeneralAutorotate: Integer;
    procedure Set_PDFGeneralAutorotate(PDFGeneralAutorotate: Integer);
    function Get_PDFGeneralCompatibility: Integer;
    procedure Set_PDFGeneralCompatibility(PDFGeneralCompatibility: Integer);
    function Get_PDFGeneralDefault: Integer;
    procedure Set_PDFGeneralDefault(PDFGeneralDefault: Integer);
    function Get_PDFGeneralOverprint: Integer;
    procedure Set_PDFGeneralOverprint(PDFGeneralOverprint: Integer);
    function Get_PDFGeneralResolution: Integer;
    procedure Set_PDFGeneralResolution(PDFGeneralResolution: Integer);
    function Get_PDFHighEncryption: Integer;
    procedure Set_PDFHighEncryption(PDFHighEncryption: Integer);
    function Get_PDFLowEncryption: Integer;
    procedure Set_PDFLowEncryption(PDFLowEncryption: Integer);
    function Get_PDFOptimize: Integer;
    procedure Set_PDFOptimize(PDFOptimize: Integer);
    function Get_PDFOwnerPass: Integer;
    procedure Set_PDFOwnerPass(PDFOwnerPass: Integer);
    function Get_PDFOwnerPasswordString: WideString;
    procedure Set_PDFOwnerPasswordString(const PDFOwnerPasswordString: WideString);
    function Get_PDFSigningMultiSignature: Integer;
    procedure Set_PDFSigningMultiSignature(PDFSigningMultiSignature: Integer);
    function Get_PDFSigningPFXFile: WideString;
    procedure Set_PDFSigningPFXFile(const PDFSigningPFXFile: WideString);
    function Get_PDFSigningPFXFilePassword: WideString;
    procedure Set_PDFSigningPFXFilePassword(const PDFSigningPFXFilePassword: WideString);
    function Get_PDFSigningSignatureContact: WideString;
    procedure Set_PDFSigningSignatureContact(const PDFSigningSignatureContact: WideString);
    function Get_PDFSigningSignatureLeftX: Double;
    procedure Set_PDFSigningSignatureLeftX(PDFSigningSignatureLeftX: Double);
    function Get_PDFSigningSignatureLeftY: Double;
    procedure Set_PDFSigningSignatureLeftY(PDFSigningSignatureLeftY: Double);
    function Get_PDFSigningSignatureLocation: WideString;
    procedure Set_PDFSigningSignatureLocation(const PDFSigningSignatureLocation: WideString);
    function Get_PDFSigningSignatureOnPage: Integer;
    procedure Set_PDFSigningSignatureOnPage(PDFSigningSignatureOnPage: Integer);
    function Get_PDFSigningSignatureReason: WideString;
    procedure Set_PDFSigningSignatureReason(const PDFSigningSignatureReason: WideString);
    function Get_PDFSigningSignatureRightX: Double;
    procedure Set_PDFSigningSignatureRightX(PDFSigningSignatureRightX: Double);
    function Get_PDFSigningSignatureRightY: Double;
    procedure Set_PDFSigningSignatureRightY(PDFSigningSignatureRightY: Double);
    function Get_PDFSigningSignatureVisible: Integer;
    procedure Set_PDFSigningSignatureVisible(PDFSigningSignatureVisible: Integer);
    function Get_PDFSigningSignPDF: Integer;
    procedure Set_PDFSigningSignPDF(PDFSigningSignPDF: Integer);
    function Get_PDFUpdateMetadata: Integer;
    procedure Set_PDFUpdateMetadata(PDFUpdateMetadata: Integer);
    function Get_PDFUserPass: Integer;
    procedure Set_PDFUserPass(PDFUserPass: Integer);
    function Get_PDFUserPasswordString: WideString;
    procedure Set_PDFUserPasswordString(const PDFUserPasswordString: WideString);
    function Get_PDFUseSecurity: Integer;
    procedure Set_PDFUseSecurity(PDFUseSecurity: Integer);
    function Get_PNGColorscount: Integer;
    procedure Set_PNGColorscount(PNGColorscount: Integer);
    function Get_PNGResolution: Integer;
    procedure Set_PNGResolution(PNGResolution: Integer);
    function Get_PrintAfterSaving: Integer;
    procedure Set_PrintAfterSaving(PrintAfterSaving: Integer);
    function Get_PrintAfterSavingBitsPerPixel: Integer;
    procedure Set_PrintAfterSavingBitsPerPixel(PrintAfterSavingBitsPerPixel: Integer);
    function Get_PrintAfterSavingDuplex: Integer;
    procedure Set_PrintAfterSavingDuplex(PrintAfterSavingDuplex: Integer);
    function Get_PrintAfterSavingMaxResolution: Integer;
    procedure Set_PrintAfterSavingMaxResolution(PrintAfterSavingMaxResolution: Integer);
    function Get_PrintAfterSavingMaxResolutionEnabled: Integer;
    procedure Set_PrintAfterSavingMaxResolutionEnabled(PrintAfterSavingMaxResolutionEnabled: Integer);
    function Get_PrintAfterSavingNoCancel: Integer;
    procedure Set_PrintAfterSavingNoCancel(PrintAfterSavingNoCancel: Integer);
    function Get_PrintAfterSavingPrinter: WideString;
    procedure Set_PrintAfterSavingPrinter(const PrintAfterSavingPrinter: WideString);
    function Get_PrintAfterSavingQueryUser: Integer;
    procedure Set_PrintAfterSavingQueryUser(PrintAfterSavingQueryUser: Integer);
    function Get_PrintAfterSavingTumble: Integer;
    procedure Set_PrintAfterSavingTumble(PrintAfterSavingTumble: Integer);
    function Get_PrinterStop: Integer;
    procedure Set_PrinterStop(PrinterStop: Integer);
    function Get_PrinterTemppath: WideString;
    procedure Set_PrinterTemppath(const PrinterTemppath: WideString);
    function Get_ProcessPriority: Integer;
    procedure Set_ProcessPriority(ProcessPriority: Integer);
    function Get_ProgramFont: WideString;
    procedure Set_ProgramFont(const ProgramFont: WideString);
    function Get_ProgramFontCharset: Integer;
    procedure Set_ProgramFontCharset(ProgramFontCharset: Integer);
    function Get_ProgramFontSize: Integer;
    procedure Set_ProgramFontSize(ProgramFontSize: Integer);
    function Get_PSDColorsCount: Integer;
    procedure Set_PSDColorsCount(PSDColorsCount: Integer);
    function Get_PSDResolution: Integer;
    procedure Set_PSDResolution(PSDResolution: Integer);
    function Get_PSLanguageLevel: Integer;
    procedure Set_PSLanguageLevel(PSLanguageLevel: Integer);
    function Get_RAWColorsCount: Integer;
    procedure Set_RAWColorsCount(RAWColorsCount: Integer);
    function Get_RAWResolution: Integer;
    procedure Set_RAWResolution(RAWResolution: Integer);
    function Get_RemoveAllKnownFileExtensions: Integer;
    procedure Set_RemoveAllKnownFileExtensions(RemoveAllKnownFileExtensions: Integer);
    function Get_RemoveSpaces: Integer;
    procedure Set_RemoveSpaces(RemoveSpaces: Integer);
    function Get_RunProgramAfterSaving: Integer;
    procedure Set_RunProgramAfterSaving(RunProgramAfterSaving: Integer);
    function Get_RunProgramAfterSavingProgramname: WideString;
    procedure Set_RunProgramAfterSavingProgramname(const RunProgramAfterSavingProgramname: WideString);
    function Get_RunProgramAfterSavingProgramParameters: WideString;
    procedure Set_RunProgramAfterSavingProgramParameters(const RunProgramAfterSavingProgramParameters: WideString);
    function Get_RunProgramAfterSavingWaitUntilReady: Integer;
    procedure Set_RunProgramAfterSavingWaitUntilReady(RunProgramAfterSavingWaitUntilReady: Integer);
    function Get_RunProgramAfterSavingWindowstyle: Integer;
    procedure Set_RunProgramAfterSavingWindowstyle(RunProgramAfterSavingWindowstyle: Integer);
    function Get_RunProgramBeforeSaving: Integer;
    procedure Set_RunProgramBeforeSaving(RunProgramBeforeSaving: Integer);
    function Get_RunProgramBeforeSavingProgramname: WideString;
    procedure Set_RunProgramBeforeSavingProgramname(const RunProgramBeforeSavingProgramname: WideString);
    function Get_RunProgramBeforeSavingProgramParameters: WideString;
    procedure Set_RunProgramBeforeSavingProgramParameters(const RunProgramBeforeSavingProgramParameters: WideString);
    function Get_RunProgramBeforeSavingWindowstyle: Integer;
    procedure Set_RunProgramBeforeSavingWindowstyle(RunProgramBeforeSavingWindowstyle: Integer);
    function Get_SaveFilename: WideString;
    procedure Set_SaveFilename(const SaveFilename: WideString);
    function Get_SendEmailAfterAutoSaving: Integer;
    procedure Set_SendEmailAfterAutoSaving(SendEmailAfterAutoSaving: Integer);
    function Get_SendMailMethod: Integer;
    procedure Set_SendMailMethod(SendMailMethod: Integer);
    function Get_ShowAnimation: Integer;
    procedure Set_ShowAnimation(ShowAnimation: Integer);
    function Get_StampFontColor: WideString;
    procedure Set_StampFontColor(const StampFontColor: WideString);
    function Get_StampFontname: WideString;
    procedure Set_StampFontname(const StampFontname: WideString);
    function Get_StampFontsize: Integer;
    procedure Set_StampFontsize(StampFontsize: Integer);
    function Get_StampOutlineFontthickness: Integer;
    procedure Set_StampOutlineFontthickness(StampOutlineFontthickness: Integer);
    function Get_StampString: WideString;
    procedure Set_StampString(const StampString: WideString);
    function Get_StampUseOutlineFont: Integer;
    procedure Set_StampUseOutlineFont(StampUseOutlineFont: Integer);
    function Get_StandardAuthor: WideString;
    procedure Set_StandardAuthor(const StandardAuthor: WideString);
    function Get_StandardCreationdate: WideString;
    procedure Set_StandardCreationdate(const StandardCreationdate: WideString);
    function Get_StandardDateformat: WideString;
    procedure Set_StandardDateformat(const StandardDateformat: WideString);
    function Get_StandardKeywords: WideString;
    procedure Set_StandardKeywords(const StandardKeywords: WideString);
    function Get_StandardMailDomain: WideString;
    procedure Set_StandardMailDomain(const StandardMailDomain: WideString);
    function Get_StandardModifydate: WideString;
    procedure Set_StandardModifydate(const StandardModifydate: WideString);
    function Get_StandardSaveformat: Integer;
    procedure Set_StandardSaveformat(StandardSaveformat: Integer);
    function Get_StandardSubject: WideString;
    procedure Set_StandardSubject(const StandardSubject: WideString);
    function Get_StandardTitle: WideString;
    procedure Set_StandardTitle(const StandardTitle: WideString);
    function Get_StartStandardProgram: Integer;
    procedure Set_StartStandardProgram(StartStandardProgram: Integer);
    function Get_SVGResolution: Integer;
    procedure Set_SVGResolution(SVGResolution: Integer);
    function Get_TIFFColorscount: Integer;
    procedure Set_TIFFColorscount(TIFFColorscount: Integer);
    function Get_TIFFResolution: Integer;
    procedure Set_TIFFResolution(TIFFResolution: Integer);
    function Get_Toolbars: Integer;
    procedure Set_Toolbars(Toolbars: Integer);
    function Get_UpdateInterval: Integer;
    procedure Set_UpdateInterval(UpdateInterval: Integer);
    function Get_UseAutosave: Integer;
    procedure Set_UseAutosave(UseAutosave: Integer);
    function Get_UseAutosaveDirectory: Integer;
    procedure Set_UseAutosaveDirectory(UseAutosaveDirectory: Integer);
    function Get_UseCreationDateNow: Integer;
    procedure Set_UseCreationDateNow(UseCreationDateNow: Integer);
    function Get_UseCustomPaperSize: WideString;
    procedure Set_UseCustomPaperSize(const UseCustomPaperSize: WideString);
    function Get_UseFixPapersize: Integer;
    procedure Set_UseFixPapersize(UseFixPapersize: Integer);
    function Get_UseStandardAuthor: Integer;
    procedure Set_UseStandardAuthor(UseStandardAuthor: Integer);
  public
    property DefaultInterface: _clsPDFCreatorOptions read GetDefaultInterface;
  published
    property AdditionalGhostscriptParameters: WideString read Get_AdditionalGhostscriptParameters write Set_AdditionalGhostscriptParameters;
    property AdditionalGhostscriptSearchpath: WideString read Get_AdditionalGhostscriptSearchpath write Set_AdditionalGhostscriptSearchpath;
    property AddWindowsFontpath: Integer read Get_AddWindowsFontpath write Set_AddWindowsFontpath;
    property AllowSpecialGSCharsInFilenames: Integer read Get_AllowSpecialGSCharsInFilenames write Set_AllowSpecialGSCharsInFilenames;
    property AutosaveDirectory: WideString read Get_AutosaveDirectory write Set_AutosaveDirectory;
    property AutosaveFilename: WideString read Get_AutosaveFilename write Set_AutosaveFilename;
    property AutosaveFormat: Integer read Get_AutosaveFormat write Set_AutosaveFormat;
    property AutosaveStartStandardProgram: Integer read Get_AutosaveStartStandardProgram write Set_AutosaveStartStandardProgram;
    property BMPColorscount: Integer read Get_BMPColorscount write Set_BMPColorscount;
    property BMPResolution: Integer read Get_BMPResolution write Set_BMPResolution;
    property ClientComputerResolveIPAddress: Integer read Get_ClientComputerResolveIPAddress write Set_ClientComputerResolveIPAddress;
    property Counter: Currency read Get_Counter write Set_Counter;
    property DeviceHeightPoints: Double read Get_DeviceHeightPoints write Set_DeviceHeightPoints;
    property DeviceWidthPoints: Double read Get_DeviceWidthPoints write Set_DeviceWidthPoints;
    property DirectoryGhostscriptBinaries: WideString read Get_DirectoryGhostscriptBinaries write Set_DirectoryGhostscriptBinaries;
    property DirectoryGhostscriptFonts: WideString read Get_DirectoryGhostscriptFonts write Set_DirectoryGhostscriptFonts;
    property DirectoryGhostscriptLibraries: WideString read Get_DirectoryGhostscriptLibraries write Set_DirectoryGhostscriptLibraries;
    property DirectoryGhostscriptResource: WideString read Get_DirectoryGhostscriptResource write Set_DirectoryGhostscriptResource;
    property DisableEmail: Integer read Get_DisableEmail write Set_DisableEmail;
    property DontUseDocumentSettings: Integer read Get_DontUseDocumentSettings write Set_DontUseDocumentSettings;
    property EPSLanguageLevel: Integer read Get_EPSLanguageLevel write Set_EPSLanguageLevel;
    property FilenameSubstitutions: WideString read Get_FilenameSubstitutions write Set_FilenameSubstitutions;
    property FilenameSubstitutionsOnlyInTitle: Integer read Get_FilenameSubstitutionsOnlyInTitle write Set_FilenameSubstitutionsOnlyInTitle;
    property JPEGColorscount: Integer read Get_JPEGColorscount write Set_JPEGColorscount;
    property JPEGQuality: Integer read Get_JPEGQuality write Set_JPEGQuality;
    property JPEGResolution: Integer read Get_JPEGResolution write Set_JPEGResolution;
    property Language: WideString read Get_Language write Set_Language;
    property LastSaveDirectory: WideString read Get_LastSaveDirectory write Set_LastSaveDirectory;
    property LastUpdateCheck: WideString read Get_LastUpdateCheck write Set_LastUpdateCheck;
    property Logging: Integer read Get_Logging write Set_Logging;
    property LogLines: Integer read Get_LogLines write Set_LogLines;
    property NoConfirmMessageSwitchingDefaultprinter: Integer read Get_NoConfirmMessageSwitchingDefaultprinter write Set_NoConfirmMessageSwitchingDefaultprinter;
    property NoProcessingAtStartup: Integer read Get_NoProcessingAtStartup write Set_NoProcessingAtStartup;
    property NoPSCheck: Integer read Get_NoPSCheck write Set_NoPSCheck;
    property OnePagePerFile: Integer read Get_OnePagePerFile write Set_OnePagePerFile;
    property OptionsDesign: Integer read Get_OptionsDesign write Set_OptionsDesign;
    property OptionsEnabled: Integer read Get_OptionsEnabled write Set_OptionsEnabled;
    property OptionsVisible: Integer read Get_OptionsVisible write Set_OptionsVisible;
    property Papersize: WideString read Get_Papersize write Set_Papersize;
    property PCLColorsCount: Integer read Get_PCLColorsCount write Set_PCLColorsCount;
    property PCLResolution: Integer read Get_PCLResolution write Set_PCLResolution;
    property PCXColorscount: Integer read Get_PCXColorscount write Set_PCXColorscount;
    property PCXResolution: Integer read Get_PCXResolution write Set_PCXResolution;
    property PDFAes128Encryption: Integer read Get_PDFAes128Encryption write Set_PDFAes128Encryption;
    property PDFAllowAssembly: Integer read Get_PDFAllowAssembly write Set_PDFAllowAssembly;
    property PDFAllowDegradedPrinting: Integer read Get_PDFAllowDegradedPrinting write Set_PDFAllowDegradedPrinting;
    property PDFAllowFillIn: Integer read Get_PDFAllowFillIn write Set_PDFAllowFillIn;
    property PDFAllowScreenReaders: Integer read Get_PDFAllowScreenReaders write Set_PDFAllowScreenReaders;
    property PDFColorsCMYKToRGB: Integer read Get_PDFColorsCMYKToRGB write Set_PDFColorsCMYKToRGB;
    property PDFColorsColorModel: Integer read Get_PDFColorsColorModel write Set_PDFColorsColorModel;
    property PDFColorsPreserveHalftone: Integer read Get_PDFColorsPreserveHalftone write Set_PDFColorsPreserveHalftone;
    property PDFColorsPreserveOverprint: Integer read Get_PDFColorsPreserveOverprint write Set_PDFColorsPreserveOverprint;
    property PDFColorsPreserveTransfer: Integer read Get_PDFColorsPreserveTransfer write Set_PDFColorsPreserveTransfer;
    property PDFCompressionColorCompression: Integer read Get_PDFCompressionColorCompression write Set_PDFCompressionColorCompression;
    property PDFCompressionColorCompressionChoice: Integer read Get_PDFCompressionColorCompressionChoice write Set_PDFCompressionColorCompressionChoice;
    property PDFCompressionColorCompressionJPEGHighFactor: Double read Get_PDFCompressionColorCompressionJPEGHighFactor write Set_PDFCompressionColorCompressionJPEGHighFactor;
    property PDFCompressionColorCompressionJPEGLowFactor: Double read Get_PDFCompressionColorCompressionJPEGLowFactor write Set_PDFCompressionColorCompressionJPEGLowFactor;
    property PDFCompressionColorCompressionJPEGManualFactor: Double read Get_PDFCompressionColorCompressionJPEGManualFactor write Set_PDFCompressionColorCompressionJPEGManualFactor;
    property PDFCompressionColorCompressionJPEGMaximumFactor: Double read Get_PDFCompressionColorCompressionJPEGMaximumFactor write Set_PDFCompressionColorCompressionJPEGMaximumFactor;
    property PDFCompressionColorCompressionJPEGMediumFactor: Double read Get_PDFCompressionColorCompressionJPEGMediumFactor write Set_PDFCompressionColorCompressionJPEGMediumFactor;
    property PDFCompressionColorCompressionJPEGMinimumFactor: Double read Get_PDFCompressionColorCompressionJPEGMinimumFactor write Set_PDFCompressionColorCompressionJPEGMinimumFactor;
    property PDFCompressionColorResample: Integer read Get_PDFCompressionColorResample write Set_PDFCompressionColorResample;
    property PDFCompressionColorResampleChoice: Integer read Get_PDFCompressionColorResampleChoice write Set_PDFCompressionColorResampleChoice;
    property PDFCompressionColorResolution: Integer read Get_PDFCompressionColorResolution write Set_PDFCompressionColorResolution;
    property PDFCompressionGreyCompression: Integer read Get_PDFCompressionGreyCompression write Set_PDFCompressionGreyCompression;
    property PDFCompressionGreyCompressionChoice: Integer read Get_PDFCompressionGreyCompressionChoice write Set_PDFCompressionGreyCompressionChoice;
    property PDFCompressionGreyCompressionJPEGHighFactor: Double read Get_PDFCompressionGreyCompressionJPEGHighFactor write Set_PDFCompressionGreyCompressionJPEGHighFactor;
    property PDFCompressionGreyCompressionJPEGLowFactor: Double read Get_PDFCompressionGreyCompressionJPEGLowFactor write Set_PDFCompressionGreyCompressionJPEGLowFactor;
    property PDFCompressionGreyCompressionJPEGManualFactor: Double read Get_PDFCompressionGreyCompressionJPEGManualFactor write Set_PDFCompressionGreyCompressionJPEGManualFactor;
    property PDFCompressionGreyCompressionJPEGMaximumFactor: Double read Get_PDFCompressionGreyCompressionJPEGMaximumFactor write Set_PDFCompressionGreyCompressionJPEGMaximumFactor;
    property PDFCompressionGreyCompressionJPEGMediumFactor: Double read Get_PDFCompressionGreyCompressionJPEGMediumFactor write Set_PDFCompressionGreyCompressionJPEGMediumFactor;
    property PDFCompressionGreyCompressionJPEGMinimumFactor: Double read Get_PDFCompressionGreyCompressionJPEGMinimumFactor write Set_PDFCompressionGreyCompressionJPEGMinimumFactor;
    property PDFCompressionGreyResample: Integer read Get_PDFCompressionGreyResample write Set_PDFCompressionGreyResample;
    property PDFCompressionGreyResampleChoice: Integer read Get_PDFCompressionGreyResampleChoice write Set_PDFCompressionGreyResampleChoice;
    property PDFCompressionGreyResolution: Integer read Get_PDFCompressionGreyResolution write Set_PDFCompressionGreyResolution;
    property PDFCompressionMonoCompression: Integer read Get_PDFCompressionMonoCompression write Set_PDFCompressionMonoCompression;
    property PDFCompressionMonoCompressionChoice: Integer read Get_PDFCompressionMonoCompressionChoice write Set_PDFCompressionMonoCompressionChoice;
    property PDFCompressionMonoResample: Integer read Get_PDFCompressionMonoResample write Set_PDFCompressionMonoResample;
    property PDFCompressionMonoResampleChoice: Integer read Get_PDFCompressionMonoResampleChoice write Set_PDFCompressionMonoResampleChoice;
    property PDFCompressionMonoResolution: Integer read Get_PDFCompressionMonoResolution write Set_PDFCompressionMonoResolution;
    property PDFCompressionTextCompression: Integer read Get_PDFCompressionTextCompression write Set_PDFCompressionTextCompression;
    property PDFDisallowCopy: Integer read Get_PDFDisallowCopy write Set_PDFDisallowCopy;
    property PDFDisallowModifyAnnotations: Integer read Get_PDFDisallowModifyAnnotations write Set_PDFDisallowModifyAnnotations;
    property PDFDisallowModifyContents: Integer read Get_PDFDisallowModifyContents write Set_PDFDisallowModifyContents;
    property PDFDisallowPrinting: Integer read Get_PDFDisallowPrinting write Set_PDFDisallowPrinting;
    property PDFEncryptor: Integer read Get_PDFEncryptor write Set_PDFEncryptor;
    property PDFFontsEmbedAll: Integer read Get_PDFFontsEmbedAll write Set_PDFFontsEmbedAll;
    property PDFFontsSubSetFonts: Integer read Get_PDFFontsSubSetFonts write Set_PDFFontsSubSetFonts;
    property PDFFontsSubSetFontsPercent: Integer read Get_PDFFontsSubSetFontsPercent write Set_PDFFontsSubSetFontsPercent;
    property PDFGeneralASCII85: Integer read Get_PDFGeneralASCII85 write Set_PDFGeneralASCII85;
    property PDFGeneralAutorotate: Integer read Get_PDFGeneralAutorotate write Set_PDFGeneralAutorotate;
    property PDFGeneralCompatibility: Integer read Get_PDFGeneralCompatibility write Set_PDFGeneralCompatibility;
    property PDFGeneralDefault: Integer read Get_PDFGeneralDefault write Set_PDFGeneralDefault;
    property PDFGeneralOverprint: Integer read Get_PDFGeneralOverprint write Set_PDFGeneralOverprint;
    property PDFGeneralResolution: Integer read Get_PDFGeneralResolution write Set_PDFGeneralResolution;
    property PDFHighEncryption: Integer read Get_PDFHighEncryption write Set_PDFHighEncryption;
    property PDFLowEncryption: Integer read Get_PDFLowEncryption write Set_PDFLowEncryption;
    property PDFOptimize: Integer read Get_PDFOptimize write Set_PDFOptimize;
    property PDFOwnerPass: Integer read Get_PDFOwnerPass write Set_PDFOwnerPass;
    property PDFOwnerPasswordString: WideString read Get_PDFOwnerPasswordString write Set_PDFOwnerPasswordString;
    property PDFSigningMultiSignature: Integer read Get_PDFSigningMultiSignature write Set_PDFSigningMultiSignature;
    property PDFSigningPFXFile: WideString read Get_PDFSigningPFXFile write Set_PDFSigningPFXFile;
    property PDFSigningPFXFilePassword: WideString read Get_PDFSigningPFXFilePassword write Set_PDFSigningPFXFilePassword;
    property PDFSigningSignatureContact: WideString read Get_PDFSigningSignatureContact write Set_PDFSigningSignatureContact;
    property PDFSigningSignatureLeftX: Double read Get_PDFSigningSignatureLeftX write Set_PDFSigningSignatureLeftX;
    property PDFSigningSignatureLeftY: Double read Get_PDFSigningSignatureLeftY write Set_PDFSigningSignatureLeftY;
    property PDFSigningSignatureLocation: WideString read Get_PDFSigningSignatureLocation write Set_PDFSigningSignatureLocation;
    property PDFSigningSignatureOnPage: Integer read Get_PDFSigningSignatureOnPage write Set_PDFSigningSignatureOnPage;
    property PDFSigningSignatureReason: WideString read Get_PDFSigningSignatureReason write Set_PDFSigningSignatureReason;
    property PDFSigningSignatureRightX: Double read Get_PDFSigningSignatureRightX write Set_PDFSigningSignatureRightX;
    property PDFSigningSignatureRightY: Double read Get_PDFSigningSignatureRightY write Set_PDFSigningSignatureRightY;
    property PDFSigningSignatureVisible: Integer read Get_PDFSigningSignatureVisible write Set_PDFSigningSignatureVisible;
    property PDFSigningSignPDF: Integer read Get_PDFSigningSignPDF write Set_PDFSigningSignPDF;
    property PDFUpdateMetadata: Integer read Get_PDFUpdateMetadata write Set_PDFUpdateMetadata;
    property PDFUserPass: Integer read Get_PDFUserPass write Set_PDFUserPass;
    property PDFUserPasswordString: WideString read Get_PDFUserPasswordString write Set_PDFUserPasswordString;
    property PDFUseSecurity: Integer read Get_PDFUseSecurity write Set_PDFUseSecurity;
    property PNGColorscount: Integer read Get_PNGColorscount write Set_PNGColorscount;
    property PNGResolution: Integer read Get_PNGResolution write Set_PNGResolution;
    property PrintAfterSaving: Integer read Get_PrintAfterSaving write Set_PrintAfterSaving;
    property PrintAfterSavingBitsPerPixel: Integer read Get_PrintAfterSavingBitsPerPixel write Set_PrintAfterSavingBitsPerPixel;
    property PrintAfterSavingDuplex: Integer read Get_PrintAfterSavingDuplex write Set_PrintAfterSavingDuplex;
    property PrintAfterSavingMaxResolution: Integer read Get_PrintAfterSavingMaxResolution write Set_PrintAfterSavingMaxResolution;
    property PrintAfterSavingMaxResolutionEnabled: Integer read Get_PrintAfterSavingMaxResolutionEnabled write Set_PrintAfterSavingMaxResolutionEnabled;
    property PrintAfterSavingNoCancel: Integer read Get_PrintAfterSavingNoCancel write Set_PrintAfterSavingNoCancel;
    property PrintAfterSavingPrinter: WideString read Get_PrintAfterSavingPrinter write Set_PrintAfterSavingPrinter;
    property PrintAfterSavingQueryUser: Integer read Get_PrintAfterSavingQueryUser write Set_PrintAfterSavingQueryUser;
    property PrintAfterSavingTumble: Integer read Get_PrintAfterSavingTumble write Set_PrintAfterSavingTumble;
    property PrinterStop: Integer read Get_PrinterStop write Set_PrinterStop;
    property PrinterTemppath: WideString read Get_PrinterTemppath write Set_PrinterTemppath;
    property ProcessPriority: Integer read Get_ProcessPriority write Set_ProcessPriority;
    property ProgramFont: WideString read Get_ProgramFont write Set_ProgramFont;
    property ProgramFontCharset: Integer read Get_ProgramFontCharset write Set_ProgramFontCharset;
    property ProgramFontSize: Integer read Get_ProgramFontSize write Set_ProgramFontSize;
    property PSDColorsCount: Integer read Get_PSDColorsCount write Set_PSDColorsCount;
    property PSDResolution: Integer read Get_PSDResolution write Set_PSDResolution;
    property PSLanguageLevel: Integer read Get_PSLanguageLevel write Set_PSLanguageLevel;
    property RAWColorsCount: Integer read Get_RAWColorsCount write Set_RAWColorsCount;
    property RAWResolution: Integer read Get_RAWResolution write Set_RAWResolution;
    property RemoveAllKnownFileExtensions: Integer read Get_RemoveAllKnownFileExtensions write Set_RemoveAllKnownFileExtensions;
    property RemoveSpaces: Integer read Get_RemoveSpaces write Set_RemoveSpaces;
    property RunProgramAfterSaving: Integer read Get_RunProgramAfterSaving write Set_RunProgramAfterSaving;
    property RunProgramAfterSavingProgramname: WideString read Get_RunProgramAfterSavingProgramname write Set_RunProgramAfterSavingProgramname;
    property RunProgramAfterSavingProgramParameters: WideString read Get_RunProgramAfterSavingProgramParameters write Set_RunProgramAfterSavingProgramParameters;
    property RunProgramAfterSavingWaitUntilReady: Integer read Get_RunProgramAfterSavingWaitUntilReady write Set_RunProgramAfterSavingWaitUntilReady;
    property RunProgramAfterSavingWindowstyle: Integer read Get_RunProgramAfterSavingWindowstyle write Set_RunProgramAfterSavingWindowstyle;
    property RunProgramBeforeSaving: Integer read Get_RunProgramBeforeSaving write Set_RunProgramBeforeSaving;
    property RunProgramBeforeSavingProgramname: WideString read Get_RunProgramBeforeSavingProgramname write Set_RunProgramBeforeSavingProgramname;
    property RunProgramBeforeSavingProgramParameters: WideString read Get_RunProgramBeforeSavingProgramParameters write Set_RunProgramBeforeSavingProgramParameters;
    property RunProgramBeforeSavingWindowstyle: Integer read Get_RunProgramBeforeSavingWindowstyle write Set_RunProgramBeforeSavingWindowstyle;
    property SaveFilename: WideString read Get_SaveFilename write Set_SaveFilename;
    property SendEmailAfterAutoSaving: Integer read Get_SendEmailAfterAutoSaving write Set_SendEmailAfterAutoSaving;
    property SendMailMethod: Integer read Get_SendMailMethod write Set_SendMailMethod;
    property ShowAnimation: Integer read Get_ShowAnimation write Set_ShowAnimation;
    property StampFontColor: WideString read Get_StampFontColor write Set_StampFontColor;
    property StampFontname: WideString read Get_StampFontname write Set_StampFontname;
    property StampFontsize: Integer read Get_StampFontsize write Set_StampFontsize;
    property StampOutlineFontthickness: Integer read Get_StampOutlineFontthickness write Set_StampOutlineFontthickness;
    property StampString: WideString read Get_StampString write Set_StampString;
    property StampUseOutlineFont: Integer read Get_StampUseOutlineFont write Set_StampUseOutlineFont;
    property StandardAuthor: WideString read Get_StandardAuthor write Set_StandardAuthor;
    property StandardCreationdate: WideString read Get_StandardCreationdate write Set_StandardCreationdate;
    property StandardDateformat: WideString read Get_StandardDateformat write Set_StandardDateformat;
    property StandardKeywords: WideString read Get_StandardKeywords write Set_StandardKeywords;
    property StandardMailDomain: WideString read Get_StandardMailDomain write Set_StandardMailDomain;
    property StandardModifydate: WideString read Get_StandardModifydate write Set_StandardModifydate;
    property StandardSaveformat: Integer read Get_StandardSaveformat write Set_StandardSaveformat;
    property StandardSubject: WideString read Get_StandardSubject write Set_StandardSubject;
    property StandardTitle: WideString read Get_StandardTitle write Set_StandardTitle;
    property StartStandardProgram: Integer read Get_StartStandardProgram write Set_StartStandardProgram;
    property SVGResolution: Integer read Get_SVGResolution write Set_SVGResolution;
    property TIFFColorscount: Integer read Get_TIFFColorscount write Set_TIFFColorscount;
    property TIFFResolution: Integer read Get_TIFFResolution write Set_TIFFResolution;
    property Toolbars: Integer read Get_Toolbars write Set_Toolbars;
    property UpdateInterval: Integer read Get_UpdateInterval write Set_UpdateInterval;
    property UseAutosave: Integer read Get_UseAutosave write Set_UseAutosave;
    property UseAutosaveDirectory: Integer read Get_UseAutosaveDirectory write Set_UseAutosaveDirectory;
    property UseCreationDateNow: Integer read Get_UseCreationDateNow write Set_UseCreationDateNow;
    property UseCustomPaperSize: WideString read Get_UseCustomPaperSize write Set_UseCustomPaperSize;
    property UseFixPapersize: Integer read Get_UseFixPapersize write Set_UseFixPapersize;
    property UseStandardAuthor: Integer read Get_UseStandardAuthor write Set_UseStandardAuthor;
  end;
{$ENDIF}


// *********************************************************************//
// La classe CoclsPDFCreatorError fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut _clsPDFCreatorError exposée             
// par la CoClasse clsPDFCreatorError. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoclsPDFCreatorError = class
    class function Create: _clsPDFCreatorError;
    class function CreateRemote(const MachineName: string): _clsPDFCreatorError;
  end;


// *********************************************************************//
// Déclaration de classe proxy de serveur OLE
// Objet serveur        : TclsPDFCreatorError
// Chaîne d'aide        : 
// Interface par défaut : _clsPDFCreatorError
// DISP Int. Déf. ?     : No
// Interface événements : 
// TypeFlags            : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TclsPDFCreatorErrorProperties= class;
{$ENDIF}
  TclsPDFCreatorError = class(TOleServer)
  private
    FIntf:        _clsPDFCreatorError;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TclsPDFCreatorErrorProperties;
    function      GetServerProperties: TclsPDFCreatorErrorProperties;
{$ENDIF}
    function      GetDefaultInterface: _clsPDFCreatorError;
  protected
    procedure InitServerData; override;
    function Get_Number: Integer;
    procedure Set_Number(Number: Integer);
    function Get_Description: WideString;
    procedure Set_Description(const Description: WideString);
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: _clsPDFCreatorError);
    procedure Disconnect; override;
    property DefaultInterface: _clsPDFCreatorError read GetDefaultInterface;
    property Number: Integer read Get_Number write Set_Number;
    property Description: WideString read Get_Description write Set_Description;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TclsPDFCreatorErrorProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// Classe proxy Propriétés Serveur OLE
// Objet serveur    : TclsPDFCreatorError
// (Cet objet est utilisé pas l'inspecteur de propriété de l'EDI pour la 
//  modification des propriétés de ce serveur)
// *********************************************************************//
 TclsPDFCreatorErrorProperties = class(TPersistent)
  private
    FServer:    TclsPDFCreatorError;
    function    GetDefaultInterface: _clsPDFCreatorError;
    constructor Create(AServer: TclsPDFCreatorError);
  protected
    function Get_Number: Integer;
    procedure Set_Number(Number: Integer);
    function Get_Description: WideString;
    procedure Set_Description(const Description: WideString);
  public
    property DefaultInterface: _clsPDFCreatorError read GetDefaultInterface;
  published
    property Number: Integer read Get_Number write Set_Number;
    property Description: WideString read Get_Description write Set_Description;
  end;
{$ENDIF}


// *********************************************************************//
// La classe CoclsPDFCreatorInfoSpoolFile fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut _clsPDFCreatorInfoSpoolFile exposée             
// par la CoClasse clsPDFCreatorInfoSpoolFile. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoclsPDFCreatorInfoSpoolFile = class
    class function Create: _clsPDFCreatorInfoSpoolFile;
    class function CreateRemote(const MachineName: string): _clsPDFCreatorInfoSpoolFile;
  end;


// *********************************************************************//
// Déclaration de classe proxy de serveur OLE
// Objet serveur        : TclsPDFCreatorInfoSpoolFile
// Chaîne d'aide        : 
// Interface par défaut : _clsPDFCreatorInfoSpoolFile
// DISP Int. Déf. ?     : No
// Interface événements : 
// TypeFlags            : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TclsPDFCreatorInfoSpoolFileProperties= class;
{$ENDIF}
  TclsPDFCreatorInfoSpoolFile = class(TOleServer)
  private
    FIntf:        _clsPDFCreatorInfoSpoolFile;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TclsPDFCreatorInfoSpoolFileProperties;
    function      GetServerProperties: TclsPDFCreatorInfoSpoolFileProperties;
{$ENDIF}
    function      GetDefaultInterface: _clsPDFCreatorInfoSpoolFile;
  protected
    procedure InitServerData; override;
    function Get_REDMON_PORT: WideString;
    procedure Set_REDMON_PORT(const REDMON_PORT: WideString);
    function Get_REDMON_JOB: WideString;
    procedure Set_REDMON_JOB(const REDMON_JOB: WideString);
    function Get_REDMON_PRINTER: WideString;
    procedure Set_REDMON_PRINTER(const REDMON_PRINTER: WideString);
    function Get_REDMON_MACHINE: WideString;
    procedure Set_REDMON_MACHINE(const REDMON_MACHINE: WideString);
    function Get_REDMON_USER: WideString;
    procedure Set_REDMON_USER(const REDMON_USER: WideString);
    function Get_REDMON_DOCNAME: WideString;
    procedure Set_REDMON_DOCNAME(const REDMON_DOCNAME: WideString);
    function Get_REDMON_FILENAME: WideString;
    procedure Set_REDMON_FILENAME(const REDMON_FILENAME: WideString);
    function Get_REDMON_SESSIONID: WideString;
    procedure Set_REDMON_SESSIONID(const REDMON_SESSIONID: WideString);
    function Get_SpoolFilename: WideString;
    procedure Set_SpoolFilename(const SpoolFilename: WideString);
    function Get_SpoolerAccount: WideString;
    procedure Set_SpoolerAccount(const SpoolerAccount: WideString);
    function Get_Computer: WideString;
    procedure Set_Computer(const Computer: WideString);
    function Get_Created: WideString;
    procedure Set_Created(const Created: WideString);
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: _clsPDFCreatorInfoSpoolFile);
    procedure Disconnect; override;
    property DefaultInterface: _clsPDFCreatorInfoSpoolFile read GetDefaultInterface;
    property REDMON_PORT: WideString read Get_REDMON_PORT write Set_REDMON_PORT;
    property REDMON_JOB: WideString read Get_REDMON_JOB write Set_REDMON_JOB;
    property REDMON_PRINTER: WideString read Get_REDMON_PRINTER write Set_REDMON_PRINTER;
    property REDMON_MACHINE: WideString read Get_REDMON_MACHINE write Set_REDMON_MACHINE;
    property REDMON_USER: WideString read Get_REDMON_USER write Set_REDMON_USER;
    property REDMON_DOCNAME: WideString read Get_REDMON_DOCNAME write Set_REDMON_DOCNAME;
    property REDMON_FILENAME: WideString read Get_REDMON_FILENAME write Set_REDMON_FILENAME;
    property REDMON_SESSIONID: WideString read Get_REDMON_SESSIONID write Set_REDMON_SESSIONID;
    property SpoolFilename: WideString read Get_SpoolFilename write Set_SpoolFilename;
    property SpoolerAccount: WideString read Get_SpoolerAccount write Set_SpoolerAccount;
    property Computer: WideString read Get_Computer write Set_Computer;
    property Created: WideString read Get_Created write Set_Created;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TclsPDFCreatorInfoSpoolFileProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// Classe proxy Propriétés Serveur OLE
// Objet serveur    : TclsPDFCreatorInfoSpoolFile
// (Cet objet est utilisé pas l'inspecteur de propriété de l'EDI pour la 
//  modification des propriétés de ce serveur)
// *********************************************************************//
 TclsPDFCreatorInfoSpoolFileProperties = class(TPersistent)
  private
    FServer:    TclsPDFCreatorInfoSpoolFile;
    function    GetDefaultInterface: _clsPDFCreatorInfoSpoolFile;
    constructor Create(AServer: TclsPDFCreatorInfoSpoolFile);
  protected
    function Get_REDMON_PORT: WideString;
    procedure Set_REDMON_PORT(const REDMON_PORT: WideString);
    function Get_REDMON_JOB: WideString;
    procedure Set_REDMON_JOB(const REDMON_JOB: WideString);
    function Get_REDMON_PRINTER: WideString;
    procedure Set_REDMON_PRINTER(const REDMON_PRINTER: WideString);
    function Get_REDMON_MACHINE: WideString;
    procedure Set_REDMON_MACHINE(const REDMON_MACHINE: WideString);
    function Get_REDMON_USER: WideString;
    procedure Set_REDMON_USER(const REDMON_USER: WideString);
    function Get_REDMON_DOCNAME: WideString;
    procedure Set_REDMON_DOCNAME(const REDMON_DOCNAME: WideString);
    function Get_REDMON_FILENAME: WideString;
    procedure Set_REDMON_FILENAME(const REDMON_FILENAME: WideString);
    function Get_REDMON_SESSIONID: WideString;
    procedure Set_REDMON_SESSIONID(const REDMON_SESSIONID: WideString);
    function Get_SpoolFilename: WideString;
    procedure Set_SpoolFilename(const SpoolFilename: WideString);
    function Get_SpoolerAccount: WideString;
    procedure Set_SpoolerAccount(const SpoolerAccount: WideString);
    function Get_Computer: WideString;
    procedure Set_Computer(const Computer: WideString);
    function Get_Created: WideString;
    procedure Set_Created(const Created: WideString);
  public
    property DefaultInterface: _clsPDFCreatorInfoSpoolFile read GetDefaultInterface;
  published
    property REDMON_PORT: WideString read Get_REDMON_PORT write Set_REDMON_PORT;
    property REDMON_JOB: WideString read Get_REDMON_JOB write Set_REDMON_JOB;
    property REDMON_PRINTER: WideString read Get_REDMON_PRINTER write Set_REDMON_PRINTER;
    property REDMON_MACHINE: WideString read Get_REDMON_MACHINE write Set_REDMON_MACHINE;
    property REDMON_USER: WideString read Get_REDMON_USER write Set_REDMON_USER;
    property REDMON_DOCNAME: WideString read Get_REDMON_DOCNAME write Set_REDMON_DOCNAME;
    property REDMON_FILENAME: WideString read Get_REDMON_FILENAME write Set_REDMON_FILENAME;
    property REDMON_SESSIONID: WideString read Get_REDMON_SESSIONID write Set_REDMON_SESSIONID;
    property SpoolFilename: WideString read Get_SpoolFilename write Set_SpoolFilename;
    property SpoolerAccount: WideString read Get_SpoolerAccount write Set_SpoolerAccount;
    property Computer: WideString read Get_Computer write Set_Computer;
    property Created: WideString read Get_Created write Set_Created;
  end;
{$ENDIF}


// *********************************************************************//
// La classe CoclsTools fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut _clsTools exposée             
// par la CoClasse clsTools. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoclsTools = class
    class function Create: _clsTools;
    class function CreateRemote(const MachineName: string): _clsTools;
  end;


// *********************************************************************//
// Déclaration de classe proxy de serveur OLE
// Objet serveur        : TclsTools
// Chaîne d'aide        : 
// Interface par défaut : _clsTools
// DISP Int. Déf. ?     : No
// Interface événements : 
// TypeFlags            : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TclsToolsProperties= class;
{$ENDIF}
  TclsTools = class(TOleServer)
  private
    FIntf:        _clsTools;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TclsToolsProperties;
    function      GetServerProperties: TclsToolsProperties;
{$ENDIF}
    function      GetDefaultInterface: _clsTools;
  protected
    procedure InitServerData; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: _clsTools);
    procedure Disconnect; override;
    function cOpenFileDialog(var files: OleVariant; const InitFilename: WideString; 
                             const Filter: WideString; const DefaultFileExtension: WideString; 
                             const InitDir: WideString; const DialogTitle: WideString; 
                             Flags: Integer; hwnd: Integer; FilterIndex: Integer): Integer;
    function cSaveFileDialog(var filename: OleVariant; var InitFilename: WideString; 
                             var Filter: WideString; var DefaultFileExtension: WideString; 
                             var InitDir: WideString; var DialogTitle: WideString; 
                             var Flags: Integer; var hwnd: Integer; var FilterIndex: Integer): Integer;
    property DefaultInterface: _clsTools read GetDefaultInterface;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TclsToolsProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// Classe proxy Propriétés Serveur OLE
// Objet serveur    : TclsTools
// (Cet objet est utilisé pas l'inspecteur de propriété de l'EDI pour la 
//  modification des propriétés de ce serveur)
// *********************************************************************//
 TclsToolsProperties = class(TPersistent)
  private
    FServer:    TclsTools;
    function    GetDefaultInterface: _clsTools;
    constructor Create(AServer: TclsTools);
  protected
  public
    property DefaultInterface: _clsTools read GetDefaultInterface;
  published
  end;
{$ENDIF}


// *********************************************************************//
// La classe CoclsUpdate fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut _clsUpdate exposée             
// par la CoClasse clsUpdate. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoclsUpdate = class
    class function Create: _clsUpdate;
    class function CreateRemote(const MachineName: string): _clsUpdate;
  end;


// *********************************************************************//
// Déclaration de classe proxy de serveur OLE
// Objet serveur        : TclsUpdate
// Chaîne d'aide        : 
// Interface par défaut : _clsUpdate
// DISP Int. Déf. ?     : No
// Interface événements : 
// TypeFlags            : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TclsUpdateProperties= class;
{$ENDIF}
  TclsUpdate = class(TOleServer)
  private
    FIntf:        _clsUpdate;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TclsUpdateProperties;
    function      GetServerProperties: TclsUpdateProperties;
{$ENDIF}
    function      GetDefaultInterface: _clsUpdate;
  protected
    procedure InitServerData; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: _clsUpdate);
    procedure Disconnect; override;
    procedure CheckForUpdates(var ShowMessageNoNewUpdates: WordBool; 
                              var ShowErrorMessage: WordBool; var TimeOutInMs: Integer);
    property DefaultInterface: _clsUpdate read GetDefaultInterface;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TclsUpdateProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// Classe proxy Propriétés Serveur OLE
// Objet serveur    : TclsUpdate
// (Cet objet est utilisé pas l'inspecteur de propriété de l'EDI pour la 
//  modification des propriétés de ce serveur)
// *********************************************************************//
 TclsUpdateProperties = class(TPersistent)
  private
    FServer:    TclsUpdate;
    function    GetDefaultInterface: _clsUpdate;
    constructor Create(AServer: TclsUpdate);
  protected
  public
    property DefaultInterface: _clsUpdate read GetDefaultInterface;
  published
  end;
{$ENDIF}


// *********************************************************************//
// La classe CoclsPDFCreator fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut _clsPDFCreator exposée             
// par la CoClasse clsPDFCreator. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoclsPDFCreator = class
    class function Create: _clsPDFCreator;
    class function CreateRemote(const MachineName: string): _clsPDFCreator;
  end;


// *********************************************************************//
// Déclaration de classe proxy de serveur OLE
// Objet serveur        : TclsPDFCreator
// Chaîne d'aide        : 
// Interface par défaut : _clsPDFCreator
// DISP Int. Déf. ?     : No
// Interface événements : __clsPDFCreator
// TypeFlags            : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TclsPDFCreatorProperties= class;
{$ENDIF}
  TclsPDFCreator = class(TOleServer)
  private
    FOneReady: TNotifyEvent;
    FOneError: TNotifyEvent;
    FIntf:        _clsPDFCreator;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TclsPDFCreatorProperties;
    function      GetServerProperties: TclsPDFCreatorProperties;
{$ENDIF}
    function      GetDefaultInterface: _clsPDFCreator;
  protected
    procedure InitServerData; override;
    procedure InvokeEvent(DispID: TDispID; var Params: TVariantArray); override;
    function Get_cPrinterProfile(const Printername: WideString): WideString;
    procedure Set_cPrinterProfile(const Printername: WideString; const Param2: WideString);
    function Get_cIsClosed: WordBool;
    function Get_cError: _clsPDFCreatorError;
    function Get_cErrorDetail(const PropertyName: WideString): OleVariant;
    function Get_cGhostscriptVersion: WideString;
    function Get_cOutputFilename: WideString;
    function Get_cPDFCreatorApplicationPath: WideString;
    function Get_cIsLogfileDialogDisplayed: WordBool;
    function Get_cIsOptionsDialogDisplayed: WordBool;
    function Get_cProgramRelease(WithBeta: WordBool): WideString;
    function Get_cProgramIsRunning: WordBool;
    function Get_cWindowsVersion: WideString;
    function Get_cVisible: WordBool;
    procedure Set_cVisible(Param1: WordBool);
    function Get_cInstalledAsServer: WordBool;
    function Get_cPrinterStop: WordBool;
    procedure Set_cPrinterStop(Param1: WordBool);
    function Get_cOptionsNames: _Collection;
    function Get_cOption(const PropertyName: WideString): OleVariant;
    procedure Set_cOption(const PropertyName: WideString; Param2: OleVariant);
    function Get_cOptionProfile(const ProfileName: WideString; const PropertyName: WideString): OleVariant;
    procedure Set_cOptionProfile(const ProfileName: WideString; const PropertyName: WideString; 
                                 Param3: OleVariant);
    function Get_cOptions: _clsPDFCreatorOptions;
    procedure _Set_cOptions(const Param1: _clsPDFCreatorOptions);
    function Get_cOptionsProfile(const ProfileName: WideString): _clsPDFCreatorOptions;
    procedure _Set_cOptionsProfile(const ProfileName: WideString; 
                                   const Param2: _clsPDFCreatorOptions);
    function Get_cStandardOption(const PropertyName: WideString): OleVariant;
    function Get_cStandardOptions: _clsPDFCreatorOptions;
    function Get_cPostscriptInfo(const PostscriptFilename: WideString; 
                                 const PropertyName: WideString): WideString;
    function Get_cPrintjobInfos(const PrintjobFilename: WideString): _clsPDFCreatorInfoSpoolFile;
    function Get_cPrintjobInfo(const PrintjobFilename: WideString; const PropertyName: WideString): WideString;
    function Get_cCountOfPrintjobs: Integer;
    function Get_cPrintjobFilename(JobNumber: Integer): WideString;
    function Get_cDefaultPrinter: WideString;
    procedure Set_cDefaultPrinter(const Param1: WideString);
    function Get_cStopURLPrinting: WordBool;
    procedure Set_cStopURLPrinting(Param1: WordBool);
    function Get_cWindowState: Integer;
    procedure Set_cWindowState(Param1: Integer);
    function Get_cIsConverted: WordBool;
    procedure Set_cIsConverted(Param1: WordBool);
    function Get_cInstanceCounter: Integer;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: _clsPDFCreator);
    procedure Disconnect; override;
    procedure cErrorClear;
    function cIsAdministrator: WordBool;
    function cPrinterIsInstalled(const Printername: WideString): WordBool;
    function cAddPDFCreatorPrinter(const Printername: WideString; const ProfileName: WideString): WordBool;
    function cProfileExists(const ProfileName: WideString): WordBool;
    function cDeletePDFCreatorPrinter(const Printername: WideString): WordBool;
    function cGetProfileNames: _Collection;
    function cAddProfile(const ProfileName: WideString; const Options1: _clsPDFCreatorOptions): WordBool;
    function cRenameProfile(const OldProfileName: WideString; const NewProfileName: WideString): WordBool;
    function cDeleteProfile(const ProfileName: WideString): WordBool;
    procedure cAddPrintjob(const filename: WideString);
    procedure cDeletePrintjob(JobNumber: Integer);
    procedure cMovePrintjobBottom(JobNumber: Integer);
    procedure cMovePrintjobTop(JobNumber: Integer);
    procedure cMovePrintjobUp(JobNumber: Integer);
    procedure cMovePrintjobDown(JobNumber: Integer);
    function cClose: WordBool;
    function cStart(const Params: WideString; ForceInitialize: WordBool): WordBool;
    procedure cClearCache;
    procedure cClearLogfile;
    procedure cConvertPostscriptfile(const InputFilename: WideString; 
                                     const OutputFilename: WideString);
    procedure cConvertFile(const InputFilename: WideString; const OutputFilename: WideString; 
                           const SubFormat: WideString);
    procedure cTestEvent(const EventName: WideString);
    procedure cShowLogfileDialog(value: WordBool);
    procedure cShowOptionsDialog(value: WordBool);
    procedure cSendMail(const OutputFilename: WideString; const Recipients: WideString);
    function cIsPrintable(const filename: WideString): WordBool;
    procedure cCombineAll;
    function cGetPDFCreatorPrinters: _Collection;
    function cGetPrinterProfiles: _Collection;
    function cGetLogfile: WideString;
    procedure cWriteToLogfile(const LogStr: WideString);
    procedure cPrintFile(const filename: WideString);
    procedure cPrintURL(const URL: WideString; TimeBetweenLoadAndPrint: Integer);
    procedure cPrintPDFCreatorTestpage;
    procedure cPrintPrinterTestpage(const Printername: WideString);
    function cReadOptions(const ProfileName: WideString): _clsPDFCreatorOptions;
    procedure cSaveOptions(Options1: OleVariant; const ProfileName: WideString);
    function cReadOptionsFromFile(const INIFilename: WideString): _clsPDFCreatorOptions;
    procedure cSaveOptionsToFile(const INIFilename: WideString); overload;
    procedure cSaveOptionsToFile(const INIFilename: WideString; Options1: OleVariant); overload;
    function cGhostscriptRun(var Arguments: PSafeArray): WordBool;
    property DefaultInterface: _clsPDFCreator read GetDefaultInterface;
    property cPrinterProfile[const Printername: WideString]: WideString read Get_cPrinterProfile write Set_cPrinterProfile;
    property cIsClosed: WordBool read Get_cIsClosed;
    property cError: _clsPDFCreatorError read Get_cError;
    property cErrorDetail[const PropertyName: WideString]: OleVariant read Get_cErrorDetail;
    property cGhostscriptVersion: WideString read Get_cGhostscriptVersion;
    property cOutputFilename: WideString read Get_cOutputFilename;
    property cPDFCreatorApplicationPath: WideString read Get_cPDFCreatorApplicationPath;
    property cIsLogfileDialogDisplayed: WordBool read Get_cIsLogfileDialogDisplayed;
    property cIsOptionsDialogDisplayed: WordBool read Get_cIsOptionsDialogDisplayed;
    property cProgramRelease[WithBeta: WordBool]: WideString read Get_cProgramRelease;
    property cProgramIsRunning: WordBool read Get_cProgramIsRunning;
    property cWindowsVersion: WideString read Get_cWindowsVersion;
    property cInstalledAsServer: WordBool read Get_cInstalledAsServer;
    property cOptionsNames: _Collection read Get_cOptionsNames;
    property cOption[const PropertyName: WideString]: OleVariant read Get_cOption write Set_cOption;
    property cOptionProfile[const ProfileName: WideString; const PropertyName: WideString]: OleVariant read Get_cOptionProfile write Set_cOptionProfile;
    property cOptions: _clsPDFCreatorOptions read Get_cOptions write _Set_cOptions;
    property cOptionsProfile[const ProfileName: WideString]: _clsPDFCreatorOptions read Get_cOptionsProfile write _Set_cOptionsProfile;
    property cStandardOption[const PropertyName: WideString]: OleVariant read Get_cStandardOption;
    property cStandardOptions: _clsPDFCreatorOptions read Get_cStandardOptions;
    property cPostscriptInfo[const PostscriptFilename: WideString; const PropertyName: WideString]: WideString read Get_cPostscriptInfo;
    property cPrintjobInfos[const PrintjobFilename: WideString]: _clsPDFCreatorInfoSpoolFile read Get_cPrintjobInfos;
    property cPrintjobInfo[const PrintjobFilename: WideString; const PropertyName: WideString]: WideString read Get_cPrintjobInfo;
    property cCountOfPrintjobs: Integer read Get_cCountOfPrintjobs;
    property cPrintjobFilename[JobNumber: Integer]: WideString read Get_cPrintjobFilename;
    property cInstanceCounter: Integer read Get_cInstanceCounter;
    property cVisible: WordBool read Get_cVisible write Set_cVisible;
    property cPrinterStop: WordBool read Get_cPrinterStop write Set_cPrinterStop;
    property cDefaultPrinter: WideString read Get_cDefaultPrinter write Set_cDefaultPrinter;
    property cStopURLPrinting: WordBool read Get_cStopURLPrinting write Set_cStopURLPrinting;
    property cWindowState: Integer read Get_cWindowState write Set_cWindowState;
    property cIsConverted: WordBool read Get_cIsConverted write Set_cIsConverted;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TclsPDFCreatorProperties read GetServerProperties;
{$ENDIF}
    property OneReady: TNotifyEvent read FOneReady write FOneReady;
    property OneError: TNotifyEvent read FOneError write FOneError;
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// Classe proxy Propriétés Serveur OLE
// Objet serveur    : TclsPDFCreator
// (Cet objet est utilisé pas l'inspecteur de propriété de l'EDI pour la 
//  modification des propriétés de ce serveur)
// *********************************************************************//
 TclsPDFCreatorProperties = class(TPersistent)
  private
    FServer:    TclsPDFCreator;
    function    GetDefaultInterface: _clsPDFCreator;
    constructor Create(AServer: TclsPDFCreator);
  protected
    function Get_cPrinterProfile(const Printername: WideString): WideString;
    procedure Set_cPrinterProfile(const Printername: WideString; const Param2: WideString);
    function Get_cIsClosed: WordBool;
    function Get_cError: _clsPDFCreatorError;
    function Get_cErrorDetail(const PropertyName: WideString): OleVariant;
    function Get_cGhostscriptVersion: WideString;
    function Get_cOutputFilename: WideString;
    function Get_cPDFCreatorApplicationPath: WideString;
    function Get_cIsLogfileDialogDisplayed: WordBool;
    function Get_cIsOptionsDialogDisplayed: WordBool;
    function Get_cProgramRelease(WithBeta: WordBool): WideString;
    function Get_cProgramIsRunning: WordBool;
    function Get_cWindowsVersion: WideString;
    function Get_cVisible: WordBool;
    procedure Set_cVisible(Param1: WordBool);
    function Get_cInstalledAsServer: WordBool;
    function Get_cPrinterStop: WordBool;
    procedure Set_cPrinterStop(Param1: WordBool);
    function Get_cOptionsNames: _Collection;
    function Get_cOption(const PropertyName: WideString): OleVariant;
    procedure Set_cOption(const PropertyName: WideString; Param2: OleVariant);
    function Get_cOptionProfile(const ProfileName: WideString; const PropertyName: WideString): OleVariant;
    procedure Set_cOptionProfile(const ProfileName: WideString; const PropertyName: WideString; 
                                 Param3: OleVariant);
    function Get_cOptions: _clsPDFCreatorOptions;
    procedure _Set_cOptions(const Param1: _clsPDFCreatorOptions);
    function Get_cOptionsProfile(const ProfileName: WideString): _clsPDFCreatorOptions;
    procedure _Set_cOptionsProfile(const ProfileName: WideString; 
                                   const Param2: _clsPDFCreatorOptions);
    function Get_cStandardOption(const PropertyName: WideString): OleVariant;
    function Get_cStandardOptions: _clsPDFCreatorOptions;
    function Get_cPostscriptInfo(const PostscriptFilename: WideString; 
                                 const PropertyName: WideString): WideString;
    function Get_cPrintjobInfos(const PrintjobFilename: WideString): _clsPDFCreatorInfoSpoolFile;
    function Get_cPrintjobInfo(const PrintjobFilename: WideString; const PropertyName: WideString): WideString;
    function Get_cCountOfPrintjobs: Integer;
    function Get_cPrintjobFilename(JobNumber: Integer): WideString;
    function Get_cDefaultPrinter: WideString;
    procedure Set_cDefaultPrinter(const Param1: WideString);
    function Get_cStopURLPrinting: WordBool;
    procedure Set_cStopURLPrinting(Param1: WordBool);
    function Get_cWindowState: Integer;
    procedure Set_cWindowState(Param1: Integer);
    function Get_cIsConverted: WordBool;
    procedure Set_cIsConverted(Param1: WordBool);
    function Get_cInstanceCounter: Integer;
  public
    property DefaultInterface: _clsPDFCreator read GetDefaultInterface;
  published
    property cVisible: WordBool read Get_cVisible write Set_cVisible;
    property cPrinterStop: WordBool read Get_cPrinterStop write Set_cPrinterStop;
    property cDefaultPrinter: WideString read Get_cDefaultPrinter write Set_cDefaultPrinter;
    property cStopURLPrinting: WordBool read Get_cStopURLPrinting write Set_cStopURLPrinting;
    property cWindowState: Integer read Get_cWindowState write Set_cWindowState;
    property cIsConverted: WordBool read Get_cIsConverted write Set_cIsConverted;
  end;
{$ENDIF}


procedure Register;

resourcestring
  dtlServerPage = 'ActiveX';

  dtlOcxPage = 'ActiveX';

implementation

uses ComObj;

class function CoclsPDFCreatorOptions.Create: _clsPDFCreatorOptions;
begin
  Result := CreateComObject(CLASS_clsPDFCreatorOptions) as _clsPDFCreatorOptions;
end;

class function CoclsPDFCreatorOptions.CreateRemote(const MachineName: string): _clsPDFCreatorOptions;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_clsPDFCreatorOptions) as _clsPDFCreatorOptions;
end;

procedure TclsPDFCreatorOptions.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{FCC886F6-E0DF-4302-8BE4-F8A8D9CB881C}';
    IntfIID:   '{F9A69DB6-B04F-4F31-9B15-3DEC9F1F2472}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TclsPDFCreatorOptions.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as _clsPDFCreatorOptions;
  end;
end;

procedure TclsPDFCreatorOptions.ConnectTo(svrIntf: _clsPDFCreatorOptions);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TclsPDFCreatorOptions.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TclsPDFCreatorOptions.GetDefaultInterface: _clsPDFCreatorOptions;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface est NULL. Le composant n''est pas connecté au serveur. Vous devez appeler ''Connect'' ou ''ConnectTo'' avant cette opération');
  Result := FIntf;
end;

constructor TclsPDFCreatorOptions.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TclsPDFCreatorOptionsProperties.Create(Self);
{$ENDIF}
end;

destructor TclsPDFCreatorOptions.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TclsPDFCreatorOptions.GetServerProperties: TclsPDFCreatorOptionsProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TclsPDFCreatorOptions.Get_AdditionalGhostscriptParameters: WideString;
begin
    Result := DefaultInterface.AdditionalGhostscriptParameters;
end;

procedure TclsPDFCreatorOptions.Set_AdditionalGhostscriptParameters(const AdditionalGhostscriptParameters: WideString);
  { Avertissement : cette propriété AdditionalGhostscriptParameters a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.AdditionalGhostscriptParameters := AdditionalGhostscriptParameters;
end;

function TclsPDFCreatorOptions.Get_AdditionalGhostscriptSearchpath: WideString;
begin
    Result := DefaultInterface.AdditionalGhostscriptSearchpath;
end;

procedure TclsPDFCreatorOptions.Set_AdditionalGhostscriptSearchpath(const AdditionalGhostscriptSearchpath: WideString);
  { Avertissement : cette propriété AdditionalGhostscriptSearchpath a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.AdditionalGhostscriptSearchpath := AdditionalGhostscriptSearchpath;
end;

function TclsPDFCreatorOptions.Get_AddWindowsFontpath: Integer;
begin
    Result := DefaultInterface.AddWindowsFontpath;
end;

procedure TclsPDFCreatorOptions.Set_AddWindowsFontpath(AddWindowsFontpath: Integer);
begin
  DefaultInterface.Set_AddWindowsFontpath(AddWindowsFontpath);
end;

function TclsPDFCreatorOptions.Get_AllowSpecialGSCharsInFilenames: Integer;
begin
    Result := DefaultInterface.AllowSpecialGSCharsInFilenames;
end;

procedure TclsPDFCreatorOptions.Set_AllowSpecialGSCharsInFilenames(AllowSpecialGSCharsInFilenames: Integer);
begin
  DefaultInterface.Set_AllowSpecialGSCharsInFilenames(AllowSpecialGSCharsInFilenames);
end;

function TclsPDFCreatorOptions.Get_AutosaveDirectory: WideString;
begin
    Result := DefaultInterface.AutosaveDirectory;
end;

procedure TclsPDFCreatorOptions.Set_AutosaveDirectory(const AutosaveDirectory: WideString);
  { Avertissement : cette propriété AutosaveDirectory a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.AutosaveDirectory := AutosaveDirectory;
end;

function TclsPDFCreatorOptions.Get_AutosaveFilename: WideString;
begin
    Result := DefaultInterface.AutosaveFilename;
end;

procedure TclsPDFCreatorOptions.Set_AutosaveFilename(const AutosaveFilename: WideString);
  { Avertissement : cette propriété AutosaveFilename a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.AutosaveFilename := AutosaveFilename;
end;

function TclsPDFCreatorOptions.Get_AutosaveFormat: Integer;
begin
    Result := DefaultInterface.AutosaveFormat;
end;

procedure TclsPDFCreatorOptions.Set_AutosaveFormat(AutosaveFormat: Integer);
begin
  DefaultInterface.Set_AutosaveFormat(AutosaveFormat);
end;

function TclsPDFCreatorOptions.Get_AutosaveStartStandardProgram: Integer;
begin
    Result := DefaultInterface.AutosaveStartStandardProgram;
end;

procedure TclsPDFCreatorOptions.Set_AutosaveStartStandardProgram(AutosaveStartStandardProgram: Integer);
begin
  DefaultInterface.Set_AutosaveStartStandardProgram(AutosaveStartStandardProgram);
end;

function TclsPDFCreatorOptions.Get_BMPColorscount: Integer;
begin
    Result := DefaultInterface.BMPColorscount;
end;

procedure TclsPDFCreatorOptions.Set_BMPColorscount(BMPColorscount: Integer);
begin
  DefaultInterface.Set_BMPColorscount(BMPColorscount);
end;

function TclsPDFCreatorOptions.Get_BMPResolution: Integer;
begin
    Result := DefaultInterface.BMPResolution;
end;

procedure TclsPDFCreatorOptions.Set_BMPResolution(BMPResolution: Integer);
begin
  DefaultInterface.Set_BMPResolution(BMPResolution);
end;

function TclsPDFCreatorOptions.Get_ClientComputerResolveIPAddress: Integer;
begin
    Result := DefaultInterface.ClientComputerResolveIPAddress;
end;

procedure TclsPDFCreatorOptions.Set_ClientComputerResolveIPAddress(ClientComputerResolveIPAddress: Integer);
begin
  DefaultInterface.Set_ClientComputerResolveIPAddress(ClientComputerResolveIPAddress);
end;

function TclsPDFCreatorOptions.Get_Counter: Currency;
begin
    Result := DefaultInterface.Counter;
end;

procedure TclsPDFCreatorOptions.Set_Counter(Counter: Currency);
begin
  DefaultInterface.Set_Counter(Counter);
end;

function TclsPDFCreatorOptions.Get_DeviceHeightPoints: Double;
begin
    Result := DefaultInterface.DeviceHeightPoints;
end;

procedure TclsPDFCreatorOptions.Set_DeviceHeightPoints(DeviceHeightPoints: Double);
begin
  DefaultInterface.Set_DeviceHeightPoints(DeviceHeightPoints);
end;

function TclsPDFCreatorOptions.Get_DeviceWidthPoints: Double;
begin
    Result := DefaultInterface.DeviceWidthPoints;
end;

procedure TclsPDFCreatorOptions.Set_DeviceWidthPoints(DeviceWidthPoints: Double);
begin
  DefaultInterface.Set_DeviceWidthPoints(DeviceWidthPoints);
end;

function TclsPDFCreatorOptions.Get_DirectoryGhostscriptBinaries: WideString;
begin
    Result := DefaultInterface.DirectoryGhostscriptBinaries;
end;

procedure TclsPDFCreatorOptions.Set_DirectoryGhostscriptBinaries(const DirectoryGhostscriptBinaries: WideString);
  { Avertissement : cette propriété DirectoryGhostscriptBinaries a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.DirectoryGhostscriptBinaries := DirectoryGhostscriptBinaries;
end;

function TclsPDFCreatorOptions.Get_DirectoryGhostscriptFonts: WideString;
begin
    Result := DefaultInterface.DirectoryGhostscriptFonts;
end;

procedure TclsPDFCreatorOptions.Set_DirectoryGhostscriptFonts(const DirectoryGhostscriptFonts: WideString);
  { Avertissement : cette propriété DirectoryGhostscriptFonts a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.DirectoryGhostscriptFonts := DirectoryGhostscriptFonts;
end;

function TclsPDFCreatorOptions.Get_DirectoryGhostscriptLibraries: WideString;
begin
    Result := DefaultInterface.DirectoryGhostscriptLibraries;
end;

procedure TclsPDFCreatorOptions.Set_DirectoryGhostscriptLibraries(const DirectoryGhostscriptLibraries: WideString);
  { Avertissement : cette propriété DirectoryGhostscriptLibraries a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.DirectoryGhostscriptLibraries := DirectoryGhostscriptLibraries;
end;

function TclsPDFCreatorOptions.Get_DirectoryGhostscriptResource: WideString;
begin
    Result := DefaultInterface.DirectoryGhostscriptResource;
end;

procedure TclsPDFCreatorOptions.Set_DirectoryGhostscriptResource(const DirectoryGhostscriptResource: WideString);
  { Avertissement : cette propriété DirectoryGhostscriptResource a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.DirectoryGhostscriptResource := DirectoryGhostscriptResource;
end;

function TclsPDFCreatorOptions.Get_DisableEmail: Integer;
begin
    Result := DefaultInterface.DisableEmail;
end;

procedure TclsPDFCreatorOptions.Set_DisableEmail(DisableEmail: Integer);
begin
  DefaultInterface.Set_DisableEmail(DisableEmail);
end;

function TclsPDFCreatorOptions.Get_DontUseDocumentSettings: Integer;
begin
    Result := DefaultInterface.DontUseDocumentSettings;
end;

procedure TclsPDFCreatorOptions.Set_DontUseDocumentSettings(DontUseDocumentSettings: Integer);
begin
  DefaultInterface.Set_DontUseDocumentSettings(DontUseDocumentSettings);
end;

function TclsPDFCreatorOptions.Get_EPSLanguageLevel: Integer;
begin
    Result := DefaultInterface.EPSLanguageLevel;
end;

procedure TclsPDFCreatorOptions.Set_EPSLanguageLevel(EPSLanguageLevel: Integer);
begin
  DefaultInterface.Set_EPSLanguageLevel(EPSLanguageLevel);
end;

function TclsPDFCreatorOptions.Get_FilenameSubstitutions: WideString;
begin
    Result := DefaultInterface.FilenameSubstitutions;
end;

procedure TclsPDFCreatorOptions.Set_FilenameSubstitutions(const FilenameSubstitutions: WideString);
  { Avertissement : cette propriété FilenameSubstitutions a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.FilenameSubstitutions := FilenameSubstitutions;
end;

function TclsPDFCreatorOptions.Get_FilenameSubstitutionsOnlyInTitle: Integer;
begin
    Result := DefaultInterface.FilenameSubstitutionsOnlyInTitle;
end;

procedure TclsPDFCreatorOptions.Set_FilenameSubstitutionsOnlyInTitle(FilenameSubstitutionsOnlyInTitle: Integer);
begin
  DefaultInterface.Set_FilenameSubstitutionsOnlyInTitle(FilenameSubstitutionsOnlyInTitle);
end;

function TclsPDFCreatorOptions.Get_JPEGColorscount: Integer;
begin
    Result := DefaultInterface.JPEGColorscount;
end;

procedure TclsPDFCreatorOptions.Set_JPEGColorscount(JPEGColorscount: Integer);
begin
  DefaultInterface.Set_JPEGColorscount(JPEGColorscount);
end;

function TclsPDFCreatorOptions.Get_JPEGQuality: Integer;
begin
    Result := DefaultInterface.JPEGQuality;
end;

procedure TclsPDFCreatorOptions.Set_JPEGQuality(JPEGQuality: Integer);
begin
  DefaultInterface.Set_JPEGQuality(JPEGQuality);
end;

function TclsPDFCreatorOptions.Get_JPEGResolution: Integer;
begin
    Result := DefaultInterface.JPEGResolution;
end;

procedure TclsPDFCreatorOptions.Set_JPEGResolution(JPEGResolution: Integer);
begin
  DefaultInterface.Set_JPEGResolution(JPEGResolution);
end;

function TclsPDFCreatorOptions.Get_Language: WideString;
begin
    Result := DefaultInterface.Language;
end;

procedure TclsPDFCreatorOptions.Set_Language(const Language: WideString);
  { Avertissement : cette propriété Language a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.Language := Language;
end;

function TclsPDFCreatorOptions.Get_LastSaveDirectory: WideString;
begin
    Result := DefaultInterface.LastSaveDirectory;
end;

procedure TclsPDFCreatorOptions.Set_LastSaveDirectory(const LastSaveDirectory: WideString);
  { Avertissement : cette propriété LastSaveDirectory a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.LastSaveDirectory := LastSaveDirectory;
end;

function TclsPDFCreatorOptions.Get_LastUpdateCheck: WideString;
begin
    Result := DefaultInterface.LastUpdateCheck;
end;

procedure TclsPDFCreatorOptions.Set_LastUpdateCheck(const LastUpdateCheck: WideString);
  { Avertissement : cette propriété LastUpdateCheck a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.LastUpdateCheck := LastUpdateCheck;
end;

function TclsPDFCreatorOptions.Get_Logging: Integer;
begin
    Result := DefaultInterface.Logging;
end;

procedure TclsPDFCreatorOptions.Set_Logging(Logging: Integer);
begin
  DefaultInterface.Set_Logging(Logging);
end;

function TclsPDFCreatorOptions.Get_LogLines: Integer;
begin
    Result := DefaultInterface.LogLines;
end;

procedure TclsPDFCreatorOptions.Set_LogLines(LogLines: Integer);
begin
  DefaultInterface.Set_LogLines(LogLines);
end;

function TclsPDFCreatorOptions.Get_NoConfirmMessageSwitchingDefaultprinter: Integer;
begin
    Result := DefaultInterface.NoConfirmMessageSwitchingDefaultprinter;
end;

procedure TclsPDFCreatorOptions.Set_NoConfirmMessageSwitchingDefaultprinter(NoConfirmMessageSwitchingDefaultprinter: Integer);
begin
  DefaultInterface.Set_NoConfirmMessageSwitchingDefaultprinter(NoConfirmMessageSwitchingDefaultprinter);
end;

function TclsPDFCreatorOptions.Get_NoProcessingAtStartup: Integer;
begin
    Result := DefaultInterface.NoProcessingAtStartup;
end;

procedure TclsPDFCreatorOptions.Set_NoProcessingAtStartup(NoProcessingAtStartup: Integer);
begin
  DefaultInterface.Set_NoProcessingAtStartup(NoProcessingAtStartup);
end;

function TclsPDFCreatorOptions.Get_NoPSCheck: Integer;
begin
    Result := DefaultInterface.NoPSCheck;
end;

procedure TclsPDFCreatorOptions.Set_NoPSCheck(NoPSCheck: Integer);
begin
  DefaultInterface.Set_NoPSCheck(NoPSCheck);
end;

function TclsPDFCreatorOptions.Get_OnePagePerFile: Integer;
begin
    Result := DefaultInterface.OnePagePerFile;
end;

procedure TclsPDFCreatorOptions.Set_OnePagePerFile(OnePagePerFile: Integer);
begin
  DefaultInterface.Set_OnePagePerFile(OnePagePerFile);
end;

function TclsPDFCreatorOptions.Get_OptionsDesign: Integer;
begin
    Result := DefaultInterface.OptionsDesign;
end;

procedure TclsPDFCreatorOptions.Set_OptionsDesign(OptionsDesign: Integer);
begin
  DefaultInterface.Set_OptionsDesign(OptionsDesign);
end;

function TclsPDFCreatorOptions.Get_OptionsEnabled: Integer;
begin
    Result := DefaultInterface.OptionsEnabled;
end;

procedure TclsPDFCreatorOptions.Set_OptionsEnabled(OptionsEnabled: Integer);
begin
  DefaultInterface.Set_OptionsEnabled(OptionsEnabled);
end;

function TclsPDFCreatorOptions.Get_OptionsVisible: Integer;
begin
    Result := DefaultInterface.OptionsVisible;
end;

procedure TclsPDFCreatorOptions.Set_OptionsVisible(OptionsVisible: Integer);
begin
  DefaultInterface.Set_OptionsVisible(OptionsVisible);
end;

function TclsPDFCreatorOptions.Get_Papersize: WideString;
begin
    Result := DefaultInterface.Papersize;
end;

procedure TclsPDFCreatorOptions.Set_Papersize(const Papersize: WideString);
  { Avertissement : cette propriété Papersize a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.Papersize := Papersize;
end;

function TclsPDFCreatorOptions.Get_PCLColorsCount: Integer;
begin
    Result := DefaultInterface.PCLColorsCount;
end;

procedure TclsPDFCreatorOptions.Set_PCLColorsCount(PCLColorsCount: Integer);
begin
  DefaultInterface.Set_PCLColorsCount(PCLColorsCount);
end;

function TclsPDFCreatorOptions.Get_PCLResolution: Integer;
begin
    Result := DefaultInterface.PCLResolution;
end;

procedure TclsPDFCreatorOptions.Set_PCLResolution(PCLResolution: Integer);
begin
  DefaultInterface.Set_PCLResolution(PCLResolution);
end;

function TclsPDFCreatorOptions.Get_PCXColorscount: Integer;
begin
    Result := DefaultInterface.PCXColorscount;
end;

procedure TclsPDFCreatorOptions.Set_PCXColorscount(PCXColorscount: Integer);
begin
  DefaultInterface.Set_PCXColorscount(PCXColorscount);
end;

function TclsPDFCreatorOptions.Get_PCXResolution: Integer;
begin
    Result := DefaultInterface.PCXResolution;
end;

procedure TclsPDFCreatorOptions.Set_PCXResolution(PCXResolution: Integer);
begin
  DefaultInterface.Set_PCXResolution(PCXResolution);
end;

function TclsPDFCreatorOptions.Get_PDFAes128Encryption: Integer;
begin
    Result := DefaultInterface.PDFAes128Encryption;
end;

procedure TclsPDFCreatorOptions.Set_PDFAes128Encryption(PDFAes128Encryption: Integer);
begin
  DefaultInterface.Set_PDFAes128Encryption(PDFAes128Encryption);
end;

function TclsPDFCreatorOptions.Get_PDFAllowAssembly: Integer;
begin
    Result := DefaultInterface.PDFAllowAssembly;
end;

procedure TclsPDFCreatorOptions.Set_PDFAllowAssembly(PDFAllowAssembly: Integer);
begin
  DefaultInterface.Set_PDFAllowAssembly(PDFAllowAssembly);
end;

function TclsPDFCreatorOptions.Get_PDFAllowDegradedPrinting: Integer;
begin
    Result := DefaultInterface.PDFAllowDegradedPrinting;
end;

procedure TclsPDFCreatorOptions.Set_PDFAllowDegradedPrinting(PDFAllowDegradedPrinting: Integer);
begin
  DefaultInterface.Set_PDFAllowDegradedPrinting(PDFAllowDegradedPrinting);
end;

function TclsPDFCreatorOptions.Get_PDFAllowFillIn: Integer;
begin
    Result := DefaultInterface.PDFAllowFillIn;
end;

procedure TclsPDFCreatorOptions.Set_PDFAllowFillIn(PDFAllowFillIn: Integer);
begin
  DefaultInterface.Set_PDFAllowFillIn(PDFAllowFillIn);
end;

function TclsPDFCreatorOptions.Get_PDFAllowScreenReaders: Integer;
begin
    Result := DefaultInterface.PDFAllowScreenReaders;
end;

procedure TclsPDFCreatorOptions.Set_PDFAllowScreenReaders(PDFAllowScreenReaders: Integer);
begin
  DefaultInterface.Set_PDFAllowScreenReaders(PDFAllowScreenReaders);
end;

function TclsPDFCreatorOptions.Get_PDFColorsCMYKToRGB: Integer;
begin
    Result := DefaultInterface.PDFColorsCMYKToRGB;
end;

procedure TclsPDFCreatorOptions.Set_PDFColorsCMYKToRGB(PDFColorsCMYKToRGB: Integer);
begin
  DefaultInterface.Set_PDFColorsCMYKToRGB(PDFColorsCMYKToRGB);
end;

function TclsPDFCreatorOptions.Get_PDFColorsColorModel: Integer;
begin
    Result := DefaultInterface.PDFColorsColorModel;
end;

procedure TclsPDFCreatorOptions.Set_PDFColorsColorModel(PDFColorsColorModel: Integer);
begin
  DefaultInterface.Set_PDFColorsColorModel(PDFColorsColorModel);
end;

function TclsPDFCreatorOptions.Get_PDFColorsPreserveHalftone: Integer;
begin
    Result := DefaultInterface.PDFColorsPreserveHalftone;
end;

procedure TclsPDFCreatorOptions.Set_PDFColorsPreserveHalftone(PDFColorsPreserveHalftone: Integer);
begin
  DefaultInterface.Set_PDFColorsPreserveHalftone(PDFColorsPreserveHalftone);
end;

function TclsPDFCreatorOptions.Get_PDFColorsPreserveOverprint: Integer;
begin
    Result := DefaultInterface.PDFColorsPreserveOverprint;
end;

procedure TclsPDFCreatorOptions.Set_PDFColorsPreserveOverprint(PDFColorsPreserveOverprint: Integer);
begin
  DefaultInterface.Set_PDFColorsPreserveOverprint(PDFColorsPreserveOverprint);
end;

function TclsPDFCreatorOptions.Get_PDFColorsPreserveTransfer: Integer;
begin
    Result := DefaultInterface.PDFColorsPreserveTransfer;
end;

procedure TclsPDFCreatorOptions.Set_PDFColorsPreserveTransfer(PDFColorsPreserveTransfer: Integer);
begin
  DefaultInterface.Set_PDFColorsPreserveTransfer(PDFColorsPreserveTransfer);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionColorCompression: Integer;
begin
    Result := DefaultInterface.PDFCompressionColorCompression;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionColorCompression(PDFCompressionColorCompression: Integer);
begin
  DefaultInterface.Set_PDFCompressionColorCompression(PDFCompressionColorCompression);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionColorCompressionChoice: Integer;
begin
    Result := DefaultInterface.PDFCompressionColorCompressionChoice;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionColorCompressionChoice(PDFCompressionColorCompressionChoice: Integer);
begin
  DefaultInterface.Set_PDFCompressionColorCompressionChoice(PDFCompressionColorCompressionChoice);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionColorCompressionJPEGHighFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionColorCompressionJPEGHighFactor;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionColorCompressionJPEGHighFactor(PDFCompressionColorCompressionJPEGHighFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionColorCompressionJPEGHighFactor(PDFCompressionColorCompressionJPEGHighFactor);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionColorCompressionJPEGLowFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionColorCompressionJPEGLowFactor;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionColorCompressionJPEGLowFactor(PDFCompressionColorCompressionJPEGLowFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionColorCompressionJPEGLowFactor(PDFCompressionColorCompressionJPEGLowFactor);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionColorCompressionJPEGManualFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionColorCompressionJPEGManualFactor;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionColorCompressionJPEGManualFactor(PDFCompressionColorCompressionJPEGManualFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionColorCompressionJPEGManualFactor(PDFCompressionColorCompressionJPEGManualFactor);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionColorCompressionJPEGMaximumFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionColorCompressionJPEGMaximumFactor;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionColorCompressionJPEGMaximumFactor(PDFCompressionColorCompressionJPEGMaximumFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionColorCompressionJPEGMaximumFactor(PDFCompressionColorCompressionJPEGMaximumFactor);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionColorCompressionJPEGMediumFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionColorCompressionJPEGMediumFactor;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionColorCompressionJPEGMediumFactor(PDFCompressionColorCompressionJPEGMediumFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionColorCompressionJPEGMediumFactor(PDFCompressionColorCompressionJPEGMediumFactor);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionColorCompressionJPEGMinimumFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionColorCompressionJPEGMinimumFactor;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionColorCompressionJPEGMinimumFactor(PDFCompressionColorCompressionJPEGMinimumFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionColorCompressionJPEGMinimumFactor(PDFCompressionColorCompressionJPEGMinimumFactor);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionColorResample: Integer;
begin
    Result := DefaultInterface.PDFCompressionColorResample;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionColorResample(PDFCompressionColorResample: Integer);
begin
  DefaultInterface.Set_PDFCompressionColorResample(PDFCompressionColorResample);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionColorResampleChoice: Integer;
begin
    Result := DefaultInterface.PDFCompressionColorResampleChoice;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionColorResampleChoice(PDFCompressionColorResampleChoice: Integer);
begin
  DefaultInterface.Set_PDFCompressionColorResampleChoice(PDFCompressionColorResampleChoice);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionColorResolution: Integer;
begin
    Result := DefaultInterface.PDFCompressionColorResolution;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionColorResolution(PDFCompressionColorResolution: Integer);
begin
  DefaultInterface.Set_PDFCompressionColorResolution(PDFCompressionColorResolution);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionGreyCompression: Integer;
begin
    Result := DefaultInterface.PDFCompressionGreyCompression;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionGreyCompression(PDFCompressionGreyCompression: Integer);
begin
  DefaultInterface.Set_PDFCompressionGreyCompression(PDFCompressionGreyCompression);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionGreyCompressionChoice: Integer;
begin
    Result := DefaultInterface.PDFCompressionGreyCompressionChoice;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionGreyCompressionChoice(PDFCompressionGreyCompressionChoice: Integer);
begin
  DefaultInterface.Set_PDFCompressionGreyCompressionChoice(PDFCompressionGreyCompressionChoice);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionGreyCompressionJPEGHighFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionGreyCompressionJPEGHighFactor;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionGreyCompressionJPEGHighFactor(PDFCompressionGreyCompressionJPEGHighFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionGreyCompressionJPEGHighFactor(PDFCompressionGreyCompressionJPEGHighFactor);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionGreyCompressionJPEGLowFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionGreyCompressionJPEGLowFactor;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionGreyCompressionJPEGLowFactor(PDFCompressionGreyCompressionJPEGLowFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionGreyCompressionJPEGLowFactor(PDFCompressionGreyCompressionJPEGLowFactor);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionGreyCompressionJPEGManualFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionGreyCompressionJPEGManualFactor;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionGreyCompressionJPEGManualFactor(PDFCompressionGreyCompressionJPEGManualFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionGreyCompressionJPEGManualFactor(PDFCompressionGreyCompressionJPEGManualFactor);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionGreyCompressionJPEGMaximumFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionGreyCompressionJPEGMaximumFactor;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionGreyCompressionJPEGMaximumFactor(PDFCompressionGreyCompressionJPEGMaximumFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionGreyCompressionJPEGMaximumFactor(PDFCompressionGreyCompressionJPEGMaximumFactor);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionGreyCompressionJPEGMediumFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionGreyCompressionJPEGMediumFactor;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionGreyCompressionJPEGMediumFactor(PDFCompressionGreyCompressionJPEGMediumFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionGreyCompressionJPEGMediumFactor(PDFCompressionGreyCompressionJPEGMediumFactor);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionGreyCompressionJPEGMinimumFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionGreyCompressionJPEGMinimumFactor;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionGreyCompressionJPEGMinimumFactor(PDFCompressionGreyCompressionJPEGMinimumFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionGreyCompressionJPEGMinimumFactor(PDFCompressionGreyCompressionJPEGMinimumFactor);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionGreyResample: Integer;
begin
    Result := DefaultInterface.PDFCompressionGreyResample;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionGreyResample(PDFCompressionGreyResample: Integer);
begin
  DefaultInterface.Set_PDFCompressionGreyResample(PDFCompressionGreyResample);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionGreyResampleChoice: Integer;
begin
    Result := DefaultInterface.PDFCompressionGreyResampleChoice;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionGreyResampleChoice(PDFCompressionGreyResampleChoice: Integer);
begin
  DefaultInterface.Set_PDFCompressionGreyResampleChoice(PDFCompressionGreyResampleChoice);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionGreyResolution: Integer;
begin
    Result := DefaultInterface.PDFCompressionGreyResolution;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionGreyResolution(PDFCompressionGreyResolution: Integer);
begin
  DefaultInterface.Set_PDFCompressionGreyResolution(PDFCompressionGreyResolution);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionMonoCompression: Integer;
begin
    Result := DefaultInterface.PDFCompressionMonoCompression;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionMonoCompression(PDFCompressionMonoCompression: Integer);
begin
  DefaultInterface.Set_PDFCompressionMonoCompression(PDFCompressionMonoCompression);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionMonoCompressionChoice: Integer;
begin
    Result := DefaultInterface.PDFCompressionMonoCompressionChoice;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionMonoCompressionChoice(PDFCompressionMonoCompressionChoice: Integer);
begin
  DefaultInterface.Set_PDFCompressionMonoCompressionChoice(PDFCompressionMonoCompressionChoice);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionMonoResample: Integer;
begin
    Result := DefaultInterface.PDFCompressionMonoResample;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionMonoResample(PDFCompressionMonoResample: Integer);
begin
  DefaultInterface.Set_PDFCompressionMonoResample(PDFCompressionMonoResample);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionMonoResampleChoice: Integer;
begin
    Result := DefaultInterface.PDFCompressionMonoResampleChoice;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionMonoResampleChoice(PDFCompressionMonoResampleChoice: Integer);
begin
  DefaultInterface.Set_PDFCompressionMonoResampleChoice(PDFCompressionMonoResampleChoice);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionMonoResolution: Integer;
begin
    Result := DefaultInterface.PDFCompressionMonoResolution;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionMonoResolution(PDFCompressionMonoResolution: Integer);
begin
  DefaultInterface.Set_PDFCompressionMonoResolution(PDFCompressionMonoResolution);
end;

function TclsPDFCreatorOptions.Get_PDFCompressionTextCompression: Integer;
begin
    Result := DefaultInterface.PDFCompressionTextCompression;
end;

procedure TclsPDFCreatorOptions.Set_PDFCompressionTextCompression(PDFCompressionTextCompression: Integer);
begin
  DefaultInterface.Set_PDFCompressionTextCompression(PDFCompressionTextCompression);
end;

function TclsPDFCreatorOptions.Get_PDFDisallowCopy: Integer;
begin
    Result := DefaultInterface.PDFDisallowCopy;
end;

procedure TclsPDFCreatorOptions.Set_PDFDisallowCopy(PDFDisallowCopy: Integer);
begin
  DefaultInterface.Set_PDFDisallowCopy(PDFDisallowCopy);
end;

function TclsPDFCreatorOptions.Get_PDFDisallowModifyAnnotations: Integer;
begin
    Result := DefaultInterface.PDFDisallowModifyAnnotations;
end;

procedure TclsPDFCreatorOptions.Set_PDFDisallowModifyAnnotations(PDFDisallowModifyAnnotations: Integer);
begin
  DefaultInterface.Set_PDFDisallowModifyAnnotations(PDFDisallowModifyAnnotations);
end;

function TclsPDFCreatorOptions.Get_PDFDisallowModifyContents: Integer;
begin
    Result := DefaultInterface.PDFDisallowModifyContents;
end;

procedure TclsPDFCreatorOptions.Set_PDFDisallowModifyContents(PDFDisallowModifyContents: Integer);
begin
  DefaultInterface.Set_PDFDisallowModifyContents(PDFDisallowModifyContents);
end;

function TclsPDFCreatorOptions.Get_PDFDisallowPrinting: Integer;
begin
    Result := DefaultInterface.PDFDisallowPrinting;
end;

procedure TclsPDFCreatorOptions.Set_PDFDisallowPrinting(PDFDisallowPrinting: Integer);
begin
  DefaultInterface.Set_PDFDisallowPrinting(PDFDisallowPrinting);
end;

function TclsPDFCreatorOptions.Get_PDFEncryptor: Integer;
begin
    Result := DefaultInterface.PDFEncryptor;
end;

procedure TclsPDFCreatorOptions.Set_PDFEncryptor(PDFEncryptor: Integer);
begin
  DefaultInterface.Set_PDFEncryptor(PDFEncryptor);
end;

function TclsPDFCreatorOptions.Get_PDFFontsEmbedAll: Integer;
begin
    Result := DefaultInterface.PDFFontsEmbedAll;
end;

procedure TclsPDFCreatorOptions.Set_PDFFontsEmbedAll(PDFFontsEmbedAll: Integer);
begin
  DefaultInterface.Set_PDFFontsEmbedAll(PDFFontsEmbedAll);
end;

function TclsPDFCreatorOptions.Get_PDFFontsSubSetFonts: Integer;
begin
    Result := DefaultInterface.PDFFontsSubSetFonts;
end;

procedure TclsPDFCreatorOptions.Set_PDFFontsSubSetFonts(PDFFontsSubSetFonts: Integer);
begin
  DefaultInterface.Set_PDFFontsSubSetFonts(PDFFontsSubSetFonts);
end;

function TclsPDFCreatorOptions.Get_PDFFontsSubSetFontsPercent: Integer;
begin
    Result := DefaultInterface.PDFFontsSubSetFontsPercent;
end;

procedure TclsPDFCreatorOptions.Set_PDFFontsSubSetFontsPercent(PDFFontsSubSetFontsPercent: Integer);
begin
  DefaultInterface.Set_PDFFontsSubSetFontsPercent(PDFFontsSubSetFontsPercent);
end;

function TclsPDFCreatorOptions.Get_PDFGeneralASCII85: Integer;
begin
    Result := DefaultInterface.PDFGeneralASCII85;
end;

procedure TclsPDFCreatorOptions.Set_PDFGeneralASCII85(PDFGeneralASCII85: Integer);
begin
  DefaultInterface.Set_PDFGeneralASCII85(PDFGeneralASCII85);
end;

function TclsPDFCreatorOptions.Get_PDFGeneralAutorotate: Integer;
begin
    Result := DefaultInterface.PDFGeneralAutorotate;
end;

procedure TclsPDFCreatorOptions.Set_PDFGeneralAutorotate(PDFGeneralAutorotate: Integer);
begin
  DefaultInterface.Set_PDFGeneralAutorotate(PDFGeneralAutorotate);
end;

function TclsPDFCreatorOptions.Get_PDFGeneralCompatibility: Integer;
begin
    Result := DefaultInterface.PDFGeneralCompatibility;
end;

procedure TclsPDFCreatorOptions.Set_PDFGeneralCompatibility(PDFGeneralCompatibility: Integer);
begin
  DefaultInterface.Set_PDFGeneralCompatibility(PDFGeneralCompatibility);
end;

function TclsPDFCreatorOptions.Get_PDFGeneralDefault: Integer;
begin
    Result := DefaultInterface.PDFGeneralDefault;
end;

procedure TclsPDFCreatorOptions.Set_PDFGeneralDefault(PDFGeneralDefault: Integer);
begin
  DefaultInterface.Set_PDFGeneralDefault(PDFGeneralDefault);
end;

function TclsPDFCreatorOptions.Get_PDFGeneralOverprint: Integer;
begin
    Result := DefaultInterface.PDFGeneralOverprint;
end;

procedure TclsPDFCreatorOptions.Set_PDFGeneralOverprint(PDFGeneralOverprint: Integer);
begin
  DefaultInterface.Set_PDFGeneralOverprint(PDFGeneralOverprint);
end;

function TclsPDFCreatorOptions.Get_PDFGeneralResolution: Integer;
begin
    Result := DefaultInterface.PDFGeneralResolution;
end;

procedure TclsPDFCreatorOptions.Set_PDFGeneralResolution(PDFGeneralResolution: Integer);
begin
  DefaultInterface.Set_PDFGeneralResolution(PDFGeneralResolution);
end;

function TclsPDFCreatorOptions.Get_PDFHighEncryption: Integer;
begin
    Result := DefaultInterface.PDFHighEncryption;
end;

procedure TclsPDFCreatorOptions.Set_PDFHighEncryption(PDFHighEncryption: Integer);
begin
  DefaultInterface.Set_PDFHighEncryption(PDFHighEncryption);
end;

function TclsPDFCreatorOptions.Get_PDFLowEncryption: Integer;
begin
    Result := DefaultInterface.PDFLowEncryption;
end;

procedure TclsPDFCreatorOptions.Set_PDFLowEncryption(PDFLowEncryption: Integer);
begin
  DefaultInterface.Set_PDFLowEncryption(PDFLowEncryption);
end;

function TclsPDFCreatorOptions.Get_PDFOptimize: Integer;
begin
    Result := DefaultInterface.PDFOptimize;
end;

procedure TclsPDFCreatorOptions.Set_PDFOptimize(PDFOptimize: Integer);
begin
  DefaultInterface.Set_PDFOptimize(PDFOptimize);
end;

function TclsPDFCreatorOptions.Get_PDFOwnerPass: Integer;
begin
    Result := DefaultInterface.PDFOwnerPass;
end;

procedure TclsPDFCreatorOptions.Set_PDFOwnerPass(PDFOwnerPass: Integer);
begin
  DefaultInterface.Set_PDFOwnerPass(PDFOwnerPass);
end;

function TclsPDFCreatorOptions.Get_PDFOwnerPasswordString: WideString;
begin
    Result := DefaultInterface.PDFOwnerPasswordString;
end;

procedure TclsPDFCreatorOptions.Set_PDFOwnerPasswordString(const PDFOwnerPasswordString: WideString);
  { Avertissement : cette propriété PDFOwnerPasswordString a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.PDFOwnerPasswordString := PDFOwnerPasswordString;
end;

function TclsPDFCreatorOptions.Get_PDFSigningMultiSignature: Integer;
begin
    Result := DefaultInterface.PDFSigningMultiSignature;
end;

procedure TclsPDFCreatorOptions.Set_PDFSigningMultiSignature(PDFSigningMultiSignature: Integer);
begin
  DefaultInterface.Set_PDFSigningMultiSignature(PDFSigningMultiSignature);
end;

function TclsPDFCreatorOptions.Get_PDFSigningPFXFile: WideString;
begin
    Result := DefaultInterface.PDFSigningPFXFile;
end;

procedure TclsPDFCreatorOptions.Set_PDFSigningPFXFile(const PDFSigningPFXFile: WideString);
  { Avertissement : cette propriété PDFSigningPFXFile a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.PDFSigningPFXFile := PDFSigningPFXFile;
end;

function TclsPDFCreatorOptions.Get_PDFSigningPFXFilePassword: WideString;
begin
    Result := DefaultInterface.PDFSigningPFXFilePassword;
end;

procedure TclsPDFCreatorOptions.Set_PDFSigningPFXFilePassword(const PDFSigningPFXFilePassword: WideString);
  { Avertissement : cette propriété PDFSigningPFXFilePassword a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.PDFSigningPFXFilePassword := PDFSigningPFXFilePassword;
end;

function TclsPDFCreatorOptions.Get_PDFSigningSignatureContact: WideString;
begin
    Result := DefaultInterface.PDFSigningSignatureContact;
end;

procedure TclsPDFCreatorOptions.Set_PDFSigningSignatureContact(const PDFSigningSignatureContact: WideString);
  { Avertissement : cette propriété PDFSigningSignatureContact a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.PDFSigningSignatureContact := PDFSigningSignatureContact;
end;

function TclsPDFCreatorOptions.Get_PDFSigningSignatureLeftX: Double;
begin
    Result := DefaultInterface.PDFSigningSignatureLeftX;
end;

procedure TclsPDFCreatorOptions.Set_PDFSigningSignatureLeftX(PDFSigningSignatureLeftX: Double);
begin
  DefaultInterface.Set_PDFSigningSignatureLeftX(PDFSigningSignatureLeftX);
end;

function TclsPDFCreatorOptions.Get_PDFSigningSignatureLeftY: Double;
begin
    Result := DefaultInterface.PDFSigningSignatureLeftY;
end;

procedure TclsPDFCreatorOptions.Set_PDFSigningSignatureLeftY(PDFSigningSignatureLeftY: Double);
begin
  DefaultInterface.Set_PDFSigningSignatureLeftY(PDFSigningSignatureLeftY);
end;

function TclsPDFCreatorOptions.Get_PDFSigningSignatureLocation: WideString;
begin
    Result := DefaultInterface.PDFSigningSignatureLocation;
end;

procedure TclsPDFCreatorOptions.Set_PDFSigningSignatureLocation(const PDFSigningSignatureLocation: WideString);
  { Avertissement : cette propriété PDFSigningSignatureLocation a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.PDFSigningSignatureLocation := PDFSigningSignatureLocation;
end;

function TclsPDFCreatorOptions.Get_PDFSigningSignatureOnPage: Integer;
begin
    Result := DefaultInterface.PDFSigningSignatureOnPage;
end;

procedure TclsPDFCreatorOptions.Set_PDFSigningSignatureOnPage(PDFSigningSignatureOnPage: Integer);
begin
  DefaultInterface.Set_PDFSigningSignatureOnPage(PDFSigningSignatureOnPage);
end;

function TclsPDFCreatorOptions.Get_PDFSigningSignatureReason: WideString;
begin
    Result := DefaultInterface.PDFSigningSignatureReason;
end;

procedure TclsPDFCreatorOptions.Set_PDFSigningSignatureReason(const PDFSigningSignatureReason: WideString);
  { Avertissement : cette propriété PDFSigningSignatureReason a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.PDFSigningSignatureReason := PDFSigningSignatureReason;
end;

function TclsPDFCreatorOptions.Get_PDFSigningSignatureRightX: Double;
begin
    Result := DefaultInterface.PDFSigningSignatureRightX;
end;

procedure TclsPDFCreatorOptions.Set_PDFSigningSignatureRightX(PDFSigningSignatureRightX: Double);
begin
  DefaultInterface.Set_PDFSigningSignatureRightX(PDFSigningSignatureRightX);
end;

function TclsPDFCreatorOptions.Get_PDFSigningSignatureRightY: Double;
begin
    Result := DefaultInterface.PDFSigningSignatureRightY;
end;

procedure TclsPDFCreatorOptions.Set_PDFSigningSignatureRightY(PDFSigningSignatureRightY: Double);
begin
  DefaultInterface.Set_PDFSigningSignatureRightY(PDFSigningSignatureRightY);
end;

function TclsPDFCreatorOptions.Get_PDFSigningSignatureVisible: Integer;
begin
    Result := DefaultInterface.PDFSigningSignatureVisible;
end;

procedure TclsPDFCreatorOptions.Set_PDFSigningSignatureVisible(PDFSigningSignatureVisible: Integer);
begin
  DefaultInterface.Set_PDFSigningSignatureVisible(PDFSigningSignatureVisible);
end;

function TclsPDFCreatorOptions.Get_PDFSigningSignPDF: Integer;
begin
    Result := DefaultInterface.PDFSigningSignPDF;
end;

procedure TclsPDFCreatorOptions.Set_PDFSigningSignPDF(PDFSigningSignPDF: Integer);
begin
  DefaultInterface.Set_PDFSigningSignPDF(PDFSigningSignPDF);
end;

function TclsPDFCreatorOptions.Get_PDFUpdateMetadata: Integer;
begin
    Result := DefaultInterface.PDFUpdateMetadata;
end;

procedure TclsPDFCreatorOptions.Set_PDFUpdateMetadata(PDFUpdateMetadata: Integer);
begin
  DefaultInterface.Set_PDFUpdateMetadata(PDFUpdateMetadata);
end;

function TclsPDFCreatorOptions.Get_PDFUserPass: Integer;
begin
    Result := DefaultInterface.PDFUserPass;
end;

procedure TclsPDFCreatorOptions.Set_PDFUserPass(PDFUserPass: Integer);
begin
  DefaultInterface.Set_PDFUserPass(PDFUserPass);
end;

function TclsPDFCreatorOptions.Get_PDFUserPasswordString: WideString;
begin
    Result := DefaultInterface.PDFUserPasswordString;
end;

procedure TclsPDFCreatorOptions.Set_PDFUserPasswordString(const PDFUserPasswordString: WideString);
  { Avertissement : cette propriété PDFUserPasswordString a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.PDFUserPasswordString := PDFUserPasswordString;
end;

function TclsPDFCreatorOptions.Get_PDFUseSecurity: Integer;
begin
    Result := DefaultInterface.PDFUseSecurity;
end;

procedure TclsPDFCreatorOptions.Set_PDFUseSecurity(PDFUseSecurity: Integer);
begin
  DefaultInterface.Set_PDFUseSecurity(PDFUseSecurity);
end;

function TclsPDFCreatorOptions.Get_PNGColorscount: Integer;
begin
    Result := DefaultInterface.PNGColorscount;
end;

procedure TclsPDFCreatorOptions.Set_PNGColorscount(PNGColorscount: Integer);
begin
  DefaultInterface.Set_PNGColorscount(PNGColorscount);
end;

function TclsPDFCreatorOptions.Get_PNGResolution: Integer;
begin
    Result := DefaultInterface.PNGResolution;
end;

procedure TclsPDFCreatorOptions.Set_PNGResolution(PNGResolution: Integer);
begin
  DefaultInterface.Set_PNGResolution(PNGResolution);
end;

function TclsPDFCreatorOptions.Get_PrintAfterSaving: Integer;
begin
    Result := DefaultInterface.PrintAfterSaving;
end;

procedure TclsPDFCreatorOptions.Set_PrintAfterSaving(PrintAfterSaving: Integer);
begin
  DefaultInterface.Set_PrintAfterSaving(PrintAfterSaving);
end;

function TclsPDFCreatorOptions.Get_PrintAfterSavingBitsPerPixel: Integer;
begin
    Result := DefaultInterface.PrintAfterSavingBitsPerPixel;
end;

procedure TclsPDFCreatorOptions.Set_PrintAfterSavingBitsPerPixel(PrintAfterSavingBitsPerPixel: Integer);
begin
  DefaultInterface.Set_PrintAfterSavingBitsPerPixel(PrintAfterSavingBitsPerPixel);
end;

function TclsPDFCreatorOptions.Get_PrintAfterSavingDuplex: Integer;
begin
    Result := DefaultInterface.PrintAfterSavingDuplex;
end;

procedure TclsPDFCreatorOptions.Set_PrintAfterSavingDuplex(PrintAfterSavingDuplex: Integer);
begin
  DefaultInterface.Set_PrintAfterSavingDuplex(PrintAfterSavingDuplex);
end;

function TclsPDFCreatorOptions.Get_PrintAfterSavingMaxResolution: Integer;
begin
    Result := DefaultInterface.PrintAfterSavingMaxResolution;
end;

procedure TclsPDFCreatorOptions.Set_PrintAfterSavingMaxResolution(PrintAfterSavingMaxResolution: Integer);
begin
  DefaultInterface.Set_PrintAfterSavingMaxResolution(PrintAfterSavingMaxResolution);
end;

function TclsPDFCreatorOptions.Get_PrintAfterSavingMaxResolutionEnabled: Integer;
begin
    Result := DefaultInterface.PrintAfterSavingMaxResolutionEnabled;
end;

procedure TclsPDFCreatorOptions.Set_PrintAfterSavingMaxResolutionEnabled(PrintAfterSavingMaxResolutionEnabled: Integer);
begin
  DefaultInterface.Set_PrintAfterSavingMaxResolutionEnabled(PrintAfterSavingMaxResolutionEnabled);
end;

function TclsPDFCreatorOptions.Get_PrintAfterSavingNoCancel: Integer;
begin
    Result := DefaultInterface.PrintAfterSavingNoCancel;
end;

procedure TclsPDFCreatorOptions.Set_PrintAfterSavingNoCancel(PrintAfterSavingNoCancel: Integer);
begin
  DefaultInterface.Set_PrintAfterSavingNoCancel(PrintAfterSavingNoCancel);
end;

function TclsPDFCreatorOptions.Get_PrintAfterSavingPrinter: WideString;
begin
    Result := DefaultInterface.PrintAfterSavingPrinter;
end;

procedure TclsPDFCreatorOptions.Set_PrintAfterSavingPrinter(const PrintAfterSavingPrinter: WideString);
  { Avertissement : cette propriété PrintAfterSavingPrinter a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.PrintAfterSavingPrinter := PrintAfterSavingPrinter;
end;

function TclsPDFCreatorOptions.Get_PrintAfterSavingQueryUser: Integer;
begin
    Result := DefaultInterface.PrintAfterSavingQueryUser;
end;

procedure TclsPDFCreatorOptions.Set_PrintAfterSavingQueryUser(PrintAfterSavingQueryUser: Integer);
begin
  DefaultInterface.Set_PrintAfterSavingQueryUser(PrintAfterSavingQueryUser);
end;

function TclsPDFCreatorOptions.Get_PrintAfterSavingTumble: Integer;
begin
    Result := DefaultInterface.PrintAfterSavingTumble;
end;

procedure TclsPDFCreatorOptions.Set_PrintAfterSavingTumble(PrintAfterSavingTumble: Integer);
begin
  DefaultInterface.Set_PrintAfterSavingTumble(PrintAfterSavingTumble);
end;

function TclsPDFCreatorOptions.Get_PrinterStop: Integer;
begin
    Result := DefaultInterface.PrinterStop;
end;

procedure TclsPDFCreatorOptions.Set_PrinterStop(PrinterStop: Integer);
begin
  DefaultInterface.Set_PrinterStop(PrinterStop);
end;

function TclsPDFCreatorOptions.Get_PrinterTemppath: WideString;
begin
    Result := DefaultInterface.PrinterTemppath;
end;

procedure TclsPDFCreatorOptions.Set_PrinterTemppath(const PrinterTemppath: WideString);
  { Avertissement : cette propriété PrinterTemppath a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.PrinterTemppath := PrinterTemppath;
end;

function TclsPDFCreatorOptions.Get_ProcessPriority: Integer;
begin
    Result := DefaultInterface.ProcessPriority;
end;

procedure TclsPDFCreatorOptions.Set_ProcessPriority(ProcessPriority: Integer);
begin
  DefaultInterface.Set_ProcessPriority(ProcessPriority);
end;

function TclsPDFCreatorOptions.Get_ProgramFont: WideString;
begin
    Result := DefaultInterface.ProgramFont;
end;

procedure TclsPDFCreatorOptions.Set_ProgramFont(const ProgramFont: WideString);
  { Avertissement : cette propriété ProgramFont a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.ProgramFont := ProgramFont;
end;

function TclsPDFCreatorOptions.Get_ProgramFontCharset: Integer;
begin
    Result := DefaultInterface.ProgramFontCharset;
end;

procedure TclsPDFCreatorOptions.Set_ProgramFontCharset(ProgramFontCharset: Integer);
begin
  DefaultInterface.Set_ProgramFontCharset(ProgramFontCharset);
end;

function TclsPDFCreatorOptions.Get_ProgramFontSize: Integer;
begin
    Result := DefaultInterface.ProgramFontSize;
end;

procedure TclsPDFCreatorOptions.Set_ProgramFontSize(ProgramFontSize: Integer);
begin
  DefaultInterface.Set_ProgramFontSize(ProgramFontSize);
end;

function TclsPDFCreatorOptions.Get_PSDColorsCount: Integer;
begin
    Result := DefaultInterface.PSDColorsCount;
end;

procedure TclsPDFCreatorOptions.Set_PSDColorsCount(PSDColorsCount: Integer);
begin
  DefaultInterface.Set_PSDColorsCount(PSDColorsCount);
end;

function TclsPDFCreatorOptions.Get_PSDResolution: Integer;
begin
    Result := DefaultInterface.PSDResolution;
end;

procedure TclsPDFCreatorOptions.Set_PSDResolution(PSDResolution: Integer);
begin
  DefaultInterface.Set_PSDResolution(PSDResolution);
end;

function TclsPDFCreatorOptions.Get_PSLanguageLevel: Integer;
begin
    Result := DefaultInterface.PSLanguageLevel;
end;

procedure TclsPDFCreatorOptions.Set_PSLanguageLevel(PSLanguageLevel: Integer);
begin
  DefaultInterface.Set_PSLanguageLevel(PSLanguageLevel);
end;

function TclsPDFCreatorOptions.Get_RAWColorsCount: Integer;
begin
    Result := DefaultInterface.RAWColorsCount;
end;

procedure TclsPDFCreatorOptions.Set_RAWColorsCount(RAWColorsCount: Integer);
begin
  DefaultInterface.Set_RAWColorsCount(RAWColorsCount);
end;

function TclsPDFCreatorOptions.Get_RAWResolution: Integer;
begin
    Result := DefaultInterface.RAWResolution;
end;

procedure TclsPDFCreatorOptions.Set_RAWResolution(RAWResolution: Integer);
begin
  DefaultInterface.Set_RAWResolution(RAWResolution);
end;

function TclsPDFCreatorOptions.Get_RemoveAllKnownFileExtensions: Integer;
begin
    Result := DefaultInterface.RemoveAllKnownFileExtensions;
end;

procedure TclsPDFCreatorOptions.Set_RemoveAllKnownFileExtensions(RemoveAllKnownFileExtensions: Integer);
begin
  DefaultInterface.Set_RemoveAllKnownFileExtensions(RemoveAllKnownFileExtensions);
end;

function TclsPDFCreatorOptions.Get_RemoveSpaces: Integer;
begin
    Result := DefaultInterface.RemoveSpaces;
end;

procedure TclsPDFCreatorOptions.Set_RemoveSpaces(RemoveSpaces: Integer);
begin
  DefaultInterface.Set_RemoveSpaces(RemoveSpaces);
end;

function TclsPDFCreatorOptions.Get_RunProgramAfterSaving: Integer;
begin
    Result := DefaultInterface.RunProgramAfterSaving;
end;

procedure TclsPDFCreatorOptions.Set_RunProgramAfterSaving(RunProgramAfterSaving: Integer);
begin
  DefaultInterface.Set_RunProgramAfterSaving(RunProgramAfterSaving);
end;

function TclsPDFCreatorOptions.Get_RunProgramAfterSavingProgramname: WideString;
begin
    Result := DefaultInterface.RunProgramAfterSavingProgramname;
end;

procedure TclsPDFCreatorOptions.Set_RunProgramAfterSavingProgramname(const RunProgramAfterSavingProgramname: WideString);
  { Avertissement : cette propriété RunProgramAfterSavingProgramname a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.RunProgramAfterSavingProgramname := RunProgramAfterSavingProgramname;
end;

function TclsPDFCreatorOptions.Get_RunProgramAfterSavingProgramParameters: WideString;
begin
    Result := DefaultInterface.RunProgramAfterSavingProgramParameters;
end;

procedure TclsPDFCreatorOptions.Set_RunProgramAfterSavingProgramParameters(const RunProgramAfterSavingProgramParameters: WideString);
  { Avertissement : cette propriété RunProgramAfterSavingProgramParameters a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.RunProgramAfterSavingProgramParameters := RunProgramAfterSavingProgramParameters;
end;

function TclsPDFCreatorOptions.Get_RunProgramAfterSavingWaitUntilReady: Integer;
begin
    Result := DefaultInterface.RunProgramAfterSavingWaitUntilReady;
end;

procedure TclsPDFCreatorOptions.Set_RunProgramAfterSavingWaitUntilReady(RunProgramAfterSavingWaitUntilReady: Integer);
begin
  DefaultInterface.Set_RunProgramAfterSavingWaitUntilReady(RunProgramAfterSavingWaitUntilReady);
end;

function TclsPDFCreatorOptions.Get_RunProgramAfterSavingWindowstyle: Integer;
begin
    Result := DefaultInterface.RunProgramAfterSavingWindowstyle;
end;

procedure TclsPDFCreatorOptions.Set_RunProgramAfterSavingWindowstyle(RunProgramAfterSavingWindowstyle: Integer);
begin
  DefaultInterface.Set_RunProgramAfterSavingWindowstyle(RunProgramAfterSavingWindowstyle);
end;

function TclsPDFCreatorOptions.Get_RunProgramBeforeSaving: Integer;
begin
    Result := DefaultInterface.RunProgramBeforeSaving;
end;

procedure TclsPDFCreatorOptions.Set_RunProgramBeforeSaving(RunProgramBeforeSaving: Integer);
begin
  DefaultInterface.Set_RunProgramBeforeSaving(RunProgramBeforeSaving);
end;

function TclsPDFCreatorOptions.Get_RunProgramBeforeSavingProgramname: WideString;
begin
    Result := DefaultInterface.RunProgramBeforeSavingProgramname;
end;

procedure TclsPDFCreatorOptions.Set_RunProgramBeforeSavingProgramname(const RunProgramBeforeSavingProgramname: WideString);
  { Avertissement : cette propriété RunProgramBeforeSavingProgramname a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.RunProgramBeforeSavingProgramname := RunProgramBeforeSavingProgramname;
end;

function TclsPDFCreatorOptions.Get_RunProgramBeforeSavingProgramParameters: WideString;
begin
    Result := DefaultInterface.RunProgramBeforeSavingProgramParameters;
end;

procedure TclsPDFCreatorOptions.Set_RunProgramBeforeSavingProgramParameters(const RunProgramBeforeSavingProgramParameters: WideString);
  { Avertissement : cette propriété RunProgramBeforeSavingProgramParameters a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.RunProgramBeforeSavingProgramParameters := RunProgramBeforeSavingProgramParameters;
end;

function TclsPDFCreatorOptions.Get_RunProgramBeforeSavingWindowstyle: Integer;
begin
    Result := DefaultInterface.RunProgramBeforeSavingWindowstyle;
end;

procedure TclsPDFCreatorOptions.Set_RunProgramBeforeSavingWindowstyle(RunProgramBeforeSavingWindowstyle: Integer);
begin
  DefaultInterface.Set_RunProgramBeforeSavingWindowstyle(RunProgramBeforeSavingWindowstyle);
end;

function TclsPDFCreatorOptions.Get_SaveFilename: WideString;
begin
    Result := DefaultInterface.SaveFilename;
end;

procedure TclsPDFCreatorOptions.Set_SaveFilename(const SaveFilename: WideString);
  { Avertissement : cette propriété SaveFilename a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.SaveFilename := SaveFilename;
end;

function TclsPDFCreatorOptions.Get_SendEmailAfterAutoSaving: Integer;
begin
    Result := DefaultInterface.SendEmailAfterAutoSaving;
end;

procedure TclsPDFCreatorOptions.Set_SendEmailAfterAutoSaving(SendEmailAfterAutoSaving: Integer);
begin
  DefaultInterface.Set_SendEmailAfterAutoSaving(SendEmailAfterAutoSaving);
end;

function TclsPDFCreatorOptions.Get_SendMailMethod: Integer;
begin
    Result := DefaultInterface.SendMailMethod;
end;

procedure TclsPDFCreatorOptions.Set_SendMailMethod(SendMailMethod: Integer);
begin
  DefaultInterface.Set_SendMailMethod(SendMailMethod);
end;

function TclsPDFCreatorOptions.Get_ShowAnimation: Integer;
begin
    Result := DefaultInterface.ShowAnimation;
end;

procedure TclsPDFCreatorOptions.Set_ShowAnimation(ShowAnimation: Integer);
begin
  DefaultInterface.Set_ShowAnimation(ShowAnimation);
end;

function TclsPDFCreatorOptions.Get_StampFontColor: WideString;
begin
    Result := DefaultInterface.StampFontColor;
end;

procedure TclsPDFCreatorOptions.Set_StampFontColor(const StampFontColor: WideString);
  { Avertissement : cette propriété StampFontColor a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StampFontColor := StampFontColor;
end;

function TclsPDFCreatorOptions.Get_StampFontname: WideString;
begin
    Result := DefaultInterface.StampFontname;
end;

procedure TclsPDFCreatorOptions.Set_StampFontname(const StampFontname: WideString);
  { Avertissement : cette propriété StampFontname a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StampFontname := StampFontname;
end;

function TclsPDFCreatorOptions.Get_StampFontsize: Integer;
begin
    Result := DefaultInterface.StampFontsize;
end;

procedure TclsPDFCreatorOptions.Set_StampFontsize(StampFontsize: Integer);
begin
  DefaultInterface.Set_StampFontsize(StampFontsize);
end;

function TclsPDFCreatorOptions.Get_StampOutlineFontthickness: Integer;
begin
    Result := DefaultInterface.StampOutlineFontthickness;
end;

procedure TclsPDFCreatorOptions.Set_StampOutlineFontthickness(StampOutlineFontthickness: Integer);
begin
  DefaultInterface.Set_StampOutlineFontthickness(StampOutlineFontthickness);
end;

function TclsPDFCreatorOptions.Get_StampString: WideString;
begin
    Result := DefaultInterface.StampString;
end;

procedure TclsPDFCreatorOptions.Set_StampString(const StampString: WideString);
  { Avertissement : cette propriété StampString a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StampString := StampString;
end;

function TclsPDFCreatorOptions.Get_StampUseOutlineFont: Integer;
begin
    Result := DefaultInterface.StampUseOutlineFont;
end;

procedure TclsPDFCreatorOptions.Set_StampUseOutlineFont(StampUseOutlineFont: Integer);
begin
  DefaultInterface.Set_StampUseOutlineFont(StampUseOutlineFont);
end;

function TclsPDFCreatorOptions.Get_StandardAuthor: WideString;
begin
    Result := DefaultInterface.StandardAuthor;
end;

procedure TclsPDFCreatorOptions.Set_StandardAuthor(const StandardAuthor: WideString);
  { Avertissement : cette propriété StandardAuthor a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StandardAuthor := StandardAuthor;
end;

function TclsPDFCreatorOptions.Get_StandardCreationdate: WideString;
begin
    Result := DefaultInterface.StandardCreationdate;
end;

procedure TclsPDFCreatorOptions.Set_StandardCreationdate(const StandardCreationdate: WideString);
  { Avertissement : cette propriété StandardCreationdate a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StandardCreationdate := StandardCreationdate;
end;

function TclsPDFCreatorOptions.Get_StandardDateformat: WideString;
begin
    Result := DefaultInterface.StandardDateformat;
end;

procedure TclsPDFCreatorOptions.Set_StandardDateformat(const StandardDateformat: WideString);
  { Avertissement : cette propriété StandardDateformat a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StandardDateformat := StandardDateformat;
end;

function TclsPDFCreatorOptions.Get_StandardKeywords: WideString;
begin
    Result := DefaultInterface.StandardKeywords;
end;

procedure TclsPDFCreatorOptions.Set_StandardKeywords(const StandardKeywords: WideString);
  { Avertissement : cette propriété StandardKeywords a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StandardKeywords := StandardKeywords;
end;

function TclsPDFCreatorOptions.Get_StandardMailDomain: WideString;
begin
    Result := DefaultInterface.StandardMailDomain;
end;

procedure TclsPDFCreatorOptions.Set_StandardMailDomain(const StandardMailDomain: WideString);
  { Avertissement : cette propriété StandardMailDomain a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StandardMailDomain := StandardMailDomain;
end;

function TclsPDFCreatorOptions.Get_StandardModifydate: WideString;
begin
    Result := DefaultInterface.StandardModifydate;
end;

procedure TclsPDFCreatorOptions.Set_StandardModifydate(const StandardModifydate: WideString);
  { Avertissement : cette propriété StandardModifydate a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StandardModifydate := StandardModifydate;
end;

function TclsPDFCreatorOptions.Get_StandardSaveformat: Integer;
begin
    Result := DefaultInterface.StandardSaveformat;
end;

procedure TclsPDFCreatorOptions.Set_StandardSaveformat(StandardSaveformat: Integer);
begin
  DefaultInterface.Set_StandardSaveformat(StandardSaveformat);
end;

function TclsPDFCreatorOptions.Get_StandardSubject: WideString;
begin
    Result := DefaultInterface.StandardSubject;
end;

procedure TclsPDFCreatorOptions.Set_StandardSubject(const StandardSubject: WideString);
  { Avertissement : cette propriété StandardSubject a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StandardSubject := StandardSubject;
end;

function TclsPDFCreatorOptions.Get_StandardTitle: WideString;
begin
    Result := DefaultInterface.StandardTitle;
end;

procedure TclsPDFCreatorOptions.Set_StandardTitle(const StandardTitle: WideString);
  { Avertissement : cette propriété StandardTitle a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StandardTitle := StandardTitle;
end;

function TclsPDFCreatorOptions.Get_StartStandardProgram: Integer;
begin
    Result := DefaultInterface.StartStandardProgram;
end;

procedure TclsPDFCreatorOptions.Set_StartStandardProgram(StartStandardProgram: Integer);
begin
  DefaultInterface.Set_StartStandardProgram(StartStandardProgram);
end;

function TclsPDFCreatorOptions.Get_SVGResolution: Integer;
begin
    Result := DefaultInterface.SVGResolution;
end;

procedure TclsPDFCreatorOptions.Set_SVGResolution(SVGResolution: Integer);
begin
  DefaultInterface.Set_SVGResolution(SVGResolution);
end;

function TclsPDFCreatorOptions.Get_TIFFColorscount: Integer;
begin
    Result := DefaultInterface.TIFFColorscount;
end;

procedure TclsPDFCreatorOptions.Set_TIFFColorscount(TIFFColorscount: Integer);
begin
  DefaultInterface.Set_TIFFColorscount(TIFFColorscount);
end;

function TclsPDFCreatorOptions.Get_TIFFResolution: Integer;
begin
    Result := DefaultInterface.TIFFResolution;
end;

procedure TclsPDFCreatorOptions.Set_TIFFResolution(TIFFResolution: Integer);
begin
  DefaultInterface.Set_TIFFResolution(TIFFResolution);
end;

function TclsPDFCreatorOptions.Get_Toolbars: Integer;
begin
    Result := DefaultInterface.Toolbars;
end;

procedure TclsPDFCreatorOptions.Set_Toolbars(Toolbars: Integer);
begin
  DefaultInterface.Set_Toolbars(Toolbars);
end;

function TclsPDFCreatorOptions.Get_UpdateInterval: Integer;
begin
    Result := DefaultInterface.UpdateInterval;
end;

procedure TclsPDFCreatorOptions.Set_UpdateInterval(UpdateInterval: Integer);
begin
  DefaultInterface.Set_UpdateInterval(UpdateInterval);
end;

function TclsPDFCreatorOptions.Get_UseAutosave: Integer;
begin
    Result := DefaultInterface.UseAutosave;
end;

procedure TclsPDFCreatorOptions.Set_UseAutosave(UseAutosave: Integer);
begin
  DefaultInterface.Set_UseAutosave(UseAutosave);
end;

function TclsPDFCreatorOptions.Get_UseAutosaveDirectory: Integer;
begin
    Result := DefaultInterface.UseAutosaveDirectory;
end;

procedure TclsPDFCreatorOptions.Set_UseAutosaveDirectory(UseAutosaveDirectory: Integer);
begin
  DefaultInterface.Set_UseAutosaveDirectory(UseAutosaveDirectory);
end;

function TclsPDFCreatorOptions.Get_UseCreationDateNow: Integer;
begin
    Result := DefaultInterface.UseCreationDateNow;
end;

procedure TclsPDFCreatorOptions.Set_UseCreationDateNow(UseCreationDateNow: Integer);
begin
  DefaultInterface.Set_UseCreationDateNow(UseCreationDateNow);
end;

function TclsPDFCreatorOptions.Get_UseCustomPaperSize: WideString;
begin
    Result := DefaultInterface.UseCustomPaperSize;
end;

procedure TclsPDFCreatorOptions.Set_UseCustomPaperSize(const UseCustomPaperSize: WideString);
  { Avertissement : cette propriété UseCustomPaperSize a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.UseCustomPaperSize := UseCustomPaperSize;
end;

function TclsPDFCreatorOptions.Get_UseFixPapersize: Integer;
begin
    Result := DefaultInterface.UseFixPapersize;
end;

procedure TclsPDFCreatorOptions.Set_UseFixPapersize(UseFixPapersize: Integer);
begin
  DefaultInterface.Set_UseFixPapersize(UseFixPapersize);
end;

function TclsPDFCreatorOptions.Get_UseStandardAuthor: Integer;
begin
    Result := DefaultInterface.UseStandardAuthor;
end;

procedure TclsPDFCreatorOptions.Set_UseStandardAuthor(UseStandardAuthor: Integer);
begin
  DefaultInterface.Set_UseStandardAuthor(UseStandardAuthor);
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TclsPDFCreatorOptionsProperties.Create(AServer: TclsPDFCreatorOptions);
begin
  inherited Create;
  FServer := AServer;
end;

function TclsPDFCreatorOptionsProperties.GetDefaultInterface: _clsPDFCreatorOptions;
begin
  Result := FServer.DefaultInterface;
end;

function TclsPDFCreatorOptionsProperties.Get_AdditionalGhostscriptParameters: WideString;
begin
    Result := DefaultInterface.AdditionalGhostscriptParameters;
end;

procedure TclsPDFCreatorOptionsProperties.Set_AdditionalGhostscriptParameters(const AdditionalGhostscriptParameters: WideString);
  { Avertissement : cette propriété AdditionalGhostscriptParameters a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.AdditionalGhostscriptParameters := AdditionalGhostscriptParameters;
end;

function TclsPDFCreatorOptionsProperties.Get_AdditionalGhostscriptSearchpath: WideString;
begin
    Result := DefaultInterface.AdditionalGhostscriptSearchpath;
end;

procedure TclsPDFCreatorOptionsProperties.Set_AdditionalGhostscriptSearchpath(const AdditionalGhostscriptSearchpath: WideString);
  { Avertissement : cette propriété AdditionalGhostscriptSearchpath a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.AdditionalGhostscriptSearchpath := AdditionalGhostscriptSearchpath;
end;

function TclsPDFCreatorOptionsProperties.Get_AddWindowsFontpath: Integer;
begin
    Result := DefaultInterface.AddWindowsFontpath;
end;

procedure TclsPDFCreatorOptionsProperties.Set_AddWindowsFontpath(AddWindowsFontpath: Integer);
begin
  DefaultInterface.Set_AddWindowsFontpath(AddWindowsFontpath);
end;

function TclsPDFCreatorOptionsProperties.Get_AllowSpecialGSCharsInFilenames: Integer;
begin
    Result := DefaultInterface.AllowSpecialGSCharsInFilenames;
end;

procedure TclsPDFCreatorOptionsProperties.Set_AllowSpecialGSCharsInFilenames(AllowSpecialGSCharsInFilenames: Integer);
begin
  DefaultInterface.Set_AllowSpecialGSCharsInFilenames(AllowSpecialGSCharsInFilenames);
end;

function TclsPDFCreatorOptionsProperties.Get_AutosaveDirectory: WideString;
begin
    Result := DefaultInterface.AutosaveDirectory;
end;

procedure TclsPDFCreatorOptionsProperties.Set_AutosaveDirectory(const AutosaveDirectory: WideString);
  { Avertissement : cette propriété AutosaveDirectory a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.AutosaveDirectory := AutosaveDirectory;
end;

function TclsPDFCreatorOptionsProperties.Get_AutosaveFilename: WideString;
begin
    Result := DefaultInterface.AutosaveFilename;
end;

procedure TclsPDFCreatorOptionsProperties.Set_AutosaveFilename(const AutosaveFilename: WideString);
  { Avertissement : cette propriété AutosaveFilename a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.AutosaveFilename := AutosaveFilename;
end;

function TclsPDFCreatorOptionsProperties.Get_AutosaveFormat: Integer;
begin
    Result := DefaultInterface.AutosaveFormat;
end;

procedure TclsPDFCreatorOptionsProperties.Set_AutosaveFormat(AutosaveFormat: Integer);
begin
  DefaultInterface.Set_AutosaveFormat(AutosaveFormat);
end;

function TclsPDFCreatorOptionsProperties.Get_AutosaveStartStandardProgram: Integer;
begin
    Result := DefaultInterface.AutosaveStartStandardProgram;
end;

procedure TclsPDFCreatorOptionsProperties.Set_AutosaveStartStandardProgram(AutosaveStartStandardProgram: Integer);
begin
  DefaultInterface.Set_AutosaveStartStandardProgram(AutosaveStartStandardProgram);
end;

function TclsPDFCreatorOptionsProperties.Get_BMPColorscount: Integer;
begin
    Result := DefaultInterface.BMPColorscount;
end;

procedure TclsPDFCreatorOptionsProperties.Set_BMPColorscount(BMPColorscount: Integer);
begin
  DefaultInterface.Set_BMPColorscount(BMPColorscount);
end;

function TclsPDFCreatorOptionsProperties.Get_BMPResolution: Integer;
begin
    Result := DefaultInterface.BMPResolution;
end;

procedure TclsPDFCreatorOptionsProperties.Set_BMPResolution(BMPResolution: Integer);
begin
  DefaultInterface.Set_BMPResolution(BMPResolution);
end;

function TclsPDFCreatorOptionsProperties.Get_ClientComputerResolveIPAddress: Integer;
begin
    Result := DefaultInterface.ClientComputerResolveIPAddress;
end;

procedure TclsPDFCreatorOptionsProperties.Set_ClientComputerResolveIPAddress(ClientComputerResolveIPAddress: Integer);
begin
  DefaultInterface.Set_ClientComputerResolveIPAddress(ClientComputerResolveIPAddress);
end;

function TclsPDFCreatorOptionsProperties.Get_Counter: Currency;
begin
    Result := DefaultInterface.Counter;
end;

procedure TclsPDFCreatorOptionsProperties.Set_Counter(Counter: Currency);
begin
  DefaultInterface.Set_Counter(Counter);
end;

function TclsPDFCreatorOptionsProperties.Get_DeviceHeightPoints: Double;
begin
    Result := DefaultInterface.DeviceHeightPoints;
end;

procedure TclsPDFCreatorOptionsProperties.Set_DeviceHeightPoints(DeviceHeightPoints: Double);
begin
  DefaultInterface.Set_DeviceHeightPoints(DeviceHeightPoints);
end;

function TclsPDFCreatorOptionsProperties.Get_DeviceWidthPoints: Double;
begin
    Result := DefaultInterface.DeviceWidthPoints;
end;

procedure TclsPDFCreatorOptionsProperties.Set_DeviceWidthPoints(DeviceWidthPoints: Double);
begin
  DefaultInterface.Set_DeviceWidthPoints(DeviceWidthPoints);
end;

function TclsPDFCreatorOptionsProperties.Get_DirectoryGhostscriptBinaries: WideString;
begin
    Result := DefaultInterface.DirectoryGhostscriptBinaries;
end;

procedure TclsPDFCreatorOptionsProperties.Set_DirectoryGhostscriptBinaries(const DirectoryGhostscriptBinaries: WideString);
  { Avertissement : cette propriété DirectoryGhostscriptBinaries a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.DirectoryGhostscriptBinaries := DirectoryGhostscriptBinaries;
end;

function TclsPDFCreatorOptionsProperties.Get_DirectoryGhostscriptFonts: WideString;
begin
    Result := DefaultInterface.DirectoryGhostscriptFonts;
end;

procedure TclsPDFCreatorOptionsProperties.Set_DirectoryGhostscriptFonts(const DirectoryGhostscriptFonts: WideString);
  { Avertissement : cette propriété DirectoryGhostscriptFonts a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.DirectoryGhostscriptFonts := DirectoryGhostscriptFonts;
end;

function TclsPDFCreatorOptionsProperties.Get_DirectoryGhostscriptLibraries: WideString;
begin
    Result := DefaultInterface.DirectoryGhostscriptLibraries;
end;

procedure TclsPDFCreatorOptionsProperties.Set_DirectoryGhostscriptLibraries(const DirectoryGhostscriptLibraries: WideString);
  { Avertissement : cette propriété DirectoryGhostscriptLibraries a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.DirectoryGhostscriptLibraries := DirectoryGhostscriptLibraries;
end;

function TclsPDFCreatorOptionsProperties.Get_DirectoryGhostscriptResource: WideString;
begin
    Result := DefaultInterface.DirectoryGhostscriptResource;
end;

procedure TclsPDFCreatorOptionsProperties.Set_DirectoryGhostscriptResource(const DirectoryGhostscriptResource: WideString);
  { Avertissement : cette propriété DirectoryGhostscriptResource a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.DirectoryGhostscriptResource := DirectoryGhostscriptResource;
end;

function TclsPDFCreatorOptionsProperties.Get_DisableEmail: Integer;
begin
    Result := DefaultInterface.DisableEmail;
end;

procedure TclsPDFCreatorOptionsProperties.Set_DisableEmail(DisableEmail: Integer);
begin
  DefaultInterface.Set_DisableEmail(DisableEmail);
end;

function TclsPDFCreatorOptionsProperties.Get_DontUseDocumentSettings: Integer;
begin
    Result := DefaultInterface.DontUseDocumentSettings;
end;

procedure TclsPDFCreatorOptionsProperties.Set_DontUseDocumentSettings(DontUseDocumentSettings: Integer);
begin
  DefaultInterface.Set_DontUseDocumentSettings(DontUseDocumentSettings);
end;

function TclsPDFCreatorOptionsProperties.Get_EPSLanguageLevel: Integer;
begin
    Result := DefaultInterface.EPSLanguageLevel;
end;

procedure TclsPDFCreatorOptionsProperties.Set_EPSLanguageLevel(EPSLanguageLevel: Integer);
begin
  DefaultInterface.Set_EPSLanguageLevel(EPSLanguageLevel);
end;

function TclsPDFCreatorOptionsProperties.Get_FilenameSubstitutions: WideString;
begin
    Result := DefaultInterface.FilenameSubstitutions;
end;

procedure TclsPDFCreatorOptionsProperties.Set_FilenameSubstitutions(const FilenameSubstitutions: WideString);
  { Avertissement : cette propriété FilenameSubstitutions a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.FilenameSubstitutions := FilenameSubstitutions;
end;

function TclsPDFCreatorOptionsProperties.Get_FilenameSubstitutionsOnlyInTitle: Integer;
begin
    Result := DefaultInterface.FilenameSubstitutionsOnlyInTitle;
end;

procedure TclsPDFCreatorOptionsProperties.Set_FilenameSubstitutionsOnlyInTitle(FilenameSubstitutionsOnlyInTitle: Integer);
begin
  DefaultInterface.Set_FilenameSubstitutionsOnlyInTitle(FilenameSubstitutionsOnlyInTitle);
end;

function TclsPDFCreatorOptionsProperties.Get_JPEGColorscount: Integer;
begin
    Result := DefaultInterface.JPEGColorscount;
end;

procedure TclsPDFCreatorOptionsProperties.Set_JPEGColorscount(JPEGColorscount: Integer);
begin
  DefaultInterface.Set_JPEGColorscount(JPEGColorscount);
end;

function TclsPDFCreatorOptionsProperties.Get_JPEGQuality: Integer;
begin
    Result := DefaultInterface.JPEGQuality;
end;

procedure TclsPDFCreatorOptionsProperties.Set_JPEGQuality(JPEGQuality: Integer);
begin
  DefaultInterface.Set_JPEGQuality(JPEGQuality);
end;

function TclsPDFCreatorOptionsProperties.Get_JPEGResolution: Integer;
begin
    Result := DefaultInterface.JPEGResolution;
end;

procedure TclsPDFCreatorOptionsProperties.Set_JPEGResolution(JPEGResolution: Integer);
begin
  DefaultInterface.Set_JPEGResolution(JPEGResolution);
end;

function TclsPDFCreatorOptionsProperties.Get_Language: WideString;
begin
    Result := DefaultInterface.Language;
end;

procedure TclsPDFCreatorOptionsProperties.Set_Language(const Language: WideString);
  { Avertissement : cette propriété Language a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.Language := Language;
end;

function TclsPDFCreatorOptionsProperties.Get_LastSaveDirectory: WideString;
begin
    Result := DefaultInterface.LastSaveDirectory;
end;

procedure TclsPDFCreatorOptionsProperties.Set_LastSaveDirectory(const LastSaveDirectory: WideString);
  { Avertissement : cette propriété LastSaveDirectory a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.LastSaveDirectory := LastSaveDirectory;
end;

function TclsPDFCreatorOptionsProperties.Get_LastUpdateCheck: WideString;
begin
    Result := DefaultInterface.LastUpdateCheck;
end;

procedure TclsPDFCreatorOptionsProperties.Set_LastUpdateCheck(const LastUpdateCheck: WideString);
  { Avertissement : cette propriété LastUpdateCheck a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.LastUpdateCheck := LastUpdateCheck;
end;

function TclsPDFCreatorOptionsProperties.Get_Logging: Integer;
begin
    Result := DefaultInterface.Logging;
end;

procedure TclsPDFCreatorOptionsProperties.Set_Logging(Logging: Integer);
begin
  DefaultInterface.Set_Logging(Logging);
end;

function TclsPDFCreatorOptionsProperties.Get_LogLines: Integer;
begin
    Result := DefaultInterface.LogLines;
end;

procedure TclsPDFCreatorOptionsProperties.Set_LogLines(LogLines: Integer);
begin
  DefaultInterface.Set_LogLines(LogLines);
end;

function TclsPDFCreatorOptionsProperties.Get_NoConfirmMessageSwitchingDefaultprinter: Integer;
begin
    Result := DefaultInterface.NoConfirmMessageSwitchingDefaultprinter;
end;

procedure TclsPDFCreatorOptionsProperties.Set_NoConfirmMessageSwitchingDefaultprinter(NoConfirmMessageSwitchingDefaultprinter: Integer);
begin
  DefaultInterface.Set_NoConfirmMessageSwitchingDefaultprinter(NoConfirmMessageSwitchingDefaultprinter);
end;

function TclsPDFCreatorOptionsProperties.Get_NoProcessingAtStartup: Integer;
begin
    Result := DefaultInterface.NoProcessingAtStartup;
end;

procedure TclsPDFCreatorOptionsProperties.Set_NoProcessingAtStartup(NoProcessingAtStartup: Integer);
begin
  DefaultInterface.Set_NoProcessingAtStartup(NoProcessingAtStartup);
end;

function TclsPDFCreatorOptionsProperties.Get_NoPSCheck: Integer;
begin
    Result := DefaultInterface.NoPSCheck;
end;

procedure TclsPDFCreatorOptionsProperties.Set_NoPSCheck(NoPSCheck: Integer);
begin
  DefaultInterface.Set_NoPSCheck(NoPSCheck);
end;

function TclsPDFCreatorOptionsProperties.Get_OnePagePerFile: Integer;
begin
    Result := DefaultInterface.OnePagePerFile;
end;

procedure TclsPDFCreatorOptionsProperties.Set_OnePagePerFile(OnePagePerFile: Integer);
begin
  DefaultInterface.Set_OnePagePerFile(OnePagePerFile);
end;

function TclsPDFCreatorOptionsProperties.Get_OptionsDesign: Integer;
begin
    Result := DefaultInterface.OptionsDesign;
end;

procedure TclsPDFCreatorOptionsProperties.Set_OptionsDesign(OptionsDesign: Integer);
begin
  DefaultInterface.Set_OptionsDesign(OptionsDesign);
end;

function TclsPDFCreatorOptionsProperties.Get_OptionsEnabled: Integer;
begin
    Result := DefaultInterface.OptionsEnabled;
end;

procedure TclsPDFCreatorOptionsProperties.Set_OptionsEnabled(OptionsEnabled: Integer);
begin
  DefaultInterface.Set_OptionsEnabled(OptionsEnabled);
end;

function TclsPDFCreatorOptionsProperties.Get_OptionsVisible: Integer;
begin
    Result := DefaultInterface.OptionsVisible;
end;

procedure TclsPDFCreatorOptionsProperties.Set_OptionsVisible(OptionsVisible: Integer);
begin
  DefaultInterface.Set_OptionsVisible(OptionsVisible);
end;

function TclsPDFCreatorOptionsProperties.Get_Papersize: WideString;
begin
    Result := DefaultInterface.Papersize;
end;

procedure TclsPDFCreatorOptionsProperties.Set_Papersize(const Papersize: WideString);
  { Avertissement : cette propriété Papersize a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.Papersize := Papersize;
end;

function TclsPDFCreatorOptionsProperties.Get_PCLColorsCount: Integer;
begin
    Result := DefaultInterface.PCLColorsCount;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PCLColorsCount(PCLColorsCount: Integer);
begin
  DefaultInterface.Set_PCLColorsCount(PCLColorsCount);
end;

function TclsPDFCreatorOptionsProperties.Get_PCLResolution: Integer;
begin
    Result := DefaultInterface.PCLResolution;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PCLResolution(PCLResolution: Integer);
begin
  DefaultInterface.Set_PCLResolution(PCLResolution);
end;

function TclsPDFCreatorOptionsProperties.Get_PCXColorscount: Integer;
begin
    Result := DefaultInterface.PCXColorscount;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PCXColorscount(PCXColorscount: Integer);
begin
  DefaultInterface.Set_PCXColorscount(PCXColorscount);
end;

function TclsPDFCreatorOptionsProperties.Get_PCXResolution: Integer;
begin
    Result := DefaultInterface.PCXResolution;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PCXResolution(PCXResolution: Integer);
begin
  DefaultInterface.Set_PCXResolution(PCXResolution);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFAes128Encryption: Integer;
begin
    Result := DefaultInterface.PDFAes128Encryption;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFAes128Encryption(PDFAes128Encryption: Integer);
begin
  DefaultInterface.Set_PDFAes128Encryption(PDFAes128Encryption);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFAllowAssembly: Integer;
begin
    Result := DefaultInterface.PDFAllowAssembly;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFAllowAssembly(PDFAllowAssembly: Integer);
begin
  DefaultInterface.Set_PDFAllowAssembly(PDFAllowAssembly);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFAllowDegradedPrinting: Integer;
begin
    Result := DefaultInterface.PDFAllowDegradedPrinting;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFAllowDegradedPrinting(PDFAllowDegradedPrinting: Integer);
begin
  DefaultInterface.Set_PDFAllowDegradedPrinting(PDFAllowDegradedPrinting);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFAllowFillIn: Integer;
begin
    Result := DefaultInterface.PDFAllowFillIn;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFAllowFillIn(PDFAllowFillIn: Integer);
begin
  DefaultInterface.Set_PDFAllowFillIn(PDFAllowFillIn);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFAllowScreenReaders: Integer;
begin
    Result := DefaultInterface.PDFAllowScreenReaders;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFAllowScreenReaders(PDFAllowScreenReaders: Integer);
begin
  DefaultInterface.Set_PDFAllowScreenReaders(PDFAllowScreenReaders);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFColorsCMYKToRGB: Integer;
begin
    Result := DefaultInterface.PDFColorsCMYKToRGB;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFColorsCMYKToRGB(PDFColorsCMYKToRGB: Integer);
begin
  DefaultInterface.Set_PDFColorsCMYKToRGB(PDFColorsCMYKToRGB);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFColorsColorModel: Integer;
begin
    Result := DefaultInterface.PDFColorsColorModel;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFColorsColorModel(PDFColorsColorModel: Integer);
begin
  DefaultInterface.Set_PDFColorsColorModel(PDFColorsColorModel);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFColorsPreserveHalftone: Integer;
begin
    Result := DefaultInterface.PDFColorsPreserveHalftone;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFColorsPreserveHalftone(PDFColorsPreserveHalftone: Integer);
begin
  DefaultInterface.Set_PDFColorsPreserveHalftone(PDFColorsPreserveHalftone);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFColorsPreserveOverprint: Integer;
begin
    Result := DefaultInterface.PDFColorsPreserveOverprint;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFColorsPreserveOverprint(PDFColorsPreserveOverprint: Integer);
begin
  DefaultInterface.Set_PDFColorsPreserveOverprint(PDFColorsPreserveOverprint);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFColorsPreserveTransfer: Integer;
begin
    Result := DefaultInterface.PDFColorsPreserveTransfer;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFColorsPreserveTransfer(PDFColorsPreserveTransfer: Integer);
begin
  DefaultInterface.Set_PDFColorsPreserveTransfer(PDFColorsPreserveTransfer);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionColorCompression: Integer;
begin
    Result := DefaultInterface.PDFCompressionColorCompression;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionColorCompression(PDFCompressionColorCompression: Integer);
begin
  DefaultInterface.Set_PDFCompressionColorCompression(PDFCompressionColorCompression);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionColorCompressionChoice: Integer;
begin
    Result := DefaultInterface.PDFCompressionColorCompressionChoice;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionColorCompressionChoice(PDFCompressionColorCompressionChoice: Integer);
begin
  DefaultInterface.Set_PDFCompressionColorCompressionChoice(PDFCompressionColorCompressionChoice);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionColorCompressionJPEGHighFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionColorCompressionJPEGHighFactor;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionColorCompressionJPEGHighFactor(PDFCompressionColorCompressionJPEGHighFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionColorCompressionJPEGHighFactor(PDFCompressionColorCompressionJPEGHighFactor);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionColorCompressionJPEGLowFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionColorCompressionJPEGLowFactor;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionColorCompressionJPEGLowFactor(PDFCompressionColorCompressionJPEGLowFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionColorCompressionJPEGLowFactor(PDFCompressionColorCompressionJPEGLowFactor);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionColorCompressionJPEGManualFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionColorCompressionJPEGManualFactor;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionColorCompressionJPEGManualFactor(PDFCompressionColorCompressionJPEGManualFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionColorCompressionJPEGManualFactor(PDFCompressionColorCompressionJPEGManualFactor);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionColorCompressionJPEGMaximumFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionColorCompressionJPEGMaximumFactor;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionColorCompressionJPEGMaximumFactor(PDFCompressionColorCompressionJPEGMaximumFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionColorCompressionJPEGMaximumFactor(PDFCompressionColorCompressionJPEGMaximumFactor);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionColorCompressionJPEGMediumFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionColorCompressionJPEGMediumFactor;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionColorCompressionJPEGMediumFactor(PDFCompressionColorCompressionJPEGMediumFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionColorCompressionJPEGMediumFactor(PDFCompressionColorCompressionJPEGMediumFactor);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionColorCompressionJPEGMinimumFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionColorCompressionJPEGMinimumFactor;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionColorCompressionJPEGMinimumFactor(PDFCompressionColorCompressionJPEGMinimumFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionColorCompressionJPEGMinimumFactor(PDFCompressionColorCompressionJPEGMinimumFactor);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionColorResample: Integer;
begin
    Result := DefaultInterface.PDFCompressionColorResample;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionColorResample(PDFCompressionColorResample: Integer);
begin
  DefaultInterface.Set_PDFCompressionColorResample(PDFCompressionColorResample);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionColorResampleChoice: Integer;
begin
    Result := DefaultInterface.PDFCompressionColorResampleChoice;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionColorResampleChoice(PDFCompressionColorResampleChoice: Integer);
begin
  DefaultInterface.Set_PDFCompressionColorResampleChoice(PDFCompressionColorResampleChoice);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionColorResolution: Integer;
begin
    Result := DefaultInterface.PDFCompressionColorResolution;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionColorResolution(PDFCompressionColorResolution: Integer);
begin
  DefaultInterface.Set_PDFCompressionColorResolution(PDFCompressionColorResolution);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionGreyCompression: Integer;
begin
    Result := DefaultInterface.PDFCompressionGreyCompression;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionGreyCompression(PDFCompressionGreyCompression: Integer);
begin
  DefaultInterface.Set_PDFCompressionGreyCompression(PDFCompressionGreyCompression);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionGreyCompressionChoice: Integer;
begin
    Result := DefaultInterface.PDFCompressionGreyCompressionChoice;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionGreyCompressionChoice(PDFCompressionGreyCompressionChoice: Integer);
begin
  DefaultInterface.Set_PDFCompressionGreyCompressionChoice(PDFCompressionGreyCompressionChoice);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionGreyCompressionJPEGHighFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionGreyCompressionJPEGHighFactor;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionGreyCompressionJPEGHighFactor(PDFCompressionGreyCompressionJPEGHighFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionGreyCompressionJPEGHighFactor(PDFCompressionGreyCompressionJPEGHighFactor);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionGreyCompressionJPEGLowFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionGreyCompressionJPEGLowFactor;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionGreyCompressionJPEGLowFactor(PDFCompressionGreyCompressionJPEGLowFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionGreyCompressionJPEGLowFactor(PDFCompressionGreyCompressionJPEGLowFactor);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionGreyCompressionJPEGManualFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionGreyCompressionJPEGManualFactor;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionGreyCompressionJPEGManualFactor(PDFCompressionGreyCompressionJPEGManualFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionGreyCompressionJPEGManualFactor(PDFCompressionGreyCompressionJPEGManualFactor);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionGreyCompressionJPEGMaximumFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionGreyCompressionJPEGMaximumFactor;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionGreyCompressionJPEGMaximumFactor(PDFCompressionGreyCompressionJPEGMaximumFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionGreyCompressionJPEGMaximumFactor(PDFCompressionGreyCompressionJPEGMaximumFactor);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionGreyCompressionJPEGMediumFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionGreyCompressionJPEGMediumFactor;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionGreyCompressionJPEGMediumFactor(PDFCompressionGreyCompressionJPEGMediumFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionGreyCompressionJPEGMediumFactor(PDFCompressionGreyCompressionJPEGMediumFactor);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionGreyCompressionJPEGMinimumFactor: Double;
begin
    Result := DefaultInterface.PDFCompressionGreyCompressionJPEGMinimumFactor;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionGreyCompressionJPEGMinimumFactor(PDFCompressionGreyCompressionJPEGMinimumFactor: Double);
begin
  DefaultInterface.Set_PDFCompressionGreyCompressionJPEGMinimumFactor(PDFCompressionGreyCompressionJPEGMinimumFactor);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionGreyResample: Integer;
begin
    Result := DefaultInterface.PDFCompressionGreyResample;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionGreyResample(PDFCompressionGreyResample: Integer);
begin
  DefaultInterface.Set_PDFCompressionGreyResample(PDFCompressionGreyResample);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionGreyResampleChoice: Integer;
begin
    Result := DefaultInterface.PDFCompressionGreyResampleChoice;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionGreyResampleChoice(PDFCompressionGreyResampleChoice: Integer);
begin
  DefaultInterface.Set_PDFCompressionGreyResampleChoice(PDFCompressionGreyResampleChoice);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionGreyResolution: Integer;
begin
    Result := DefaultInterface.PDFCompressionGreyResolution;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionGreyResolution(PDFCompressionGreyResolution: Integer);
begin
  DefaultInterface.Set_PDFCompressionGreyResolution(PDFCompressionGreyResolution);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionMonoCompression: Integer;
begin
    Result := DefaultInterface.PDFCompressionMonoCompression;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionMonoCompression(PDFCompressionMonoCompression: Integer);
begin
  DefaultInterface.Set_PDFCompressionMonoCompression(PDFCompressionMonoCompression);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionMonoCompressionChoice: Integer;
begin
    Result := DefaultInterface.PDFCompressionMonoCompressionChoice;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionMonoCompressionChoice(PDFCompressionMonoCompressionChoice: Integer);
begin
  DefaultInterface.Set_PDFCompressionMonoCompressionChoice(PDFCompressionMonoCompressionChoice);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionMonoResample: Integer;
begin
    Result := DefaultInterface.PDFCompressionMonoResample;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionMonoResample(PDFCompressionMonoResample: Integer);
begin
  DefaultInterface.Set_PDFCompressionMonoResample(PDFCompressionMonoResample);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionMonoResampleChoice: Integer;
begin
    Result := DefaultInterface.PDFCompressionMonoResampleChoice;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionMonoResampleChoice(PDFCompressionMonoResampleChoice: Integer);
begin
  DefaultInterface.Set_PDFCompressionMonoResampleChoice(PDFCompressionMonoResampleChoice);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionMonoResolution: Integer;
begin
    Result := DefaultInterface.PDFCompressionMonoResolution;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionMonoResolution(PDFCompressionMonoResolution: Integer);
begin
  DefaultInterface.Set_PDFCompressionMonoResolution(PDFCompressionMonoResolution);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFCompressionTextCompression: Integer;
begin
    Result := DefaultInterface.PDFCompressionTextCompression;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFCompressionTextCompression(PDFCompressionTextCompression: Integer);
begin
  DefaultInterface.Set_PDFCompressionTextCompression(PDFCompressionTextCompression);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFDisallowCopy: Integer;
begin
    Result := DefaultInterface.PDFDisallowCopy;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFDisallowCopy(PDFDisallowCopy: Integer);
begin
  DefaultInterface.Set_PDFDisallowCopy(PDFDisallowCopy);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFDisallowModifyAnnotations: Integer;
begin
    Result := DefaultInterface.PDFDisallowModifyAnnotations;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFDisallowModifyAnnotations(PDFDisallowModifyAnnotations: Integer);
begin
  DefaultInterface.Set_PDFDisallowModifyAnnotations(PDFDisallowModifyAnnotations);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFDisallowModifyContents: Integer;
begin
    Result := DefaultInterface.PDFDisallowModifyContents;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFDisallowModifyContents(PDFDisallowModifyContents: Integer);
begin
  DefaultInterface.Set_PDFDisallowModifyContents(PDFDisallowModifyContents);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFDisallowPrinting: Integer;
begin
    Result := DefaultInterface.PDFDisallowPrinting;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFDisallowPrinting(PDFDisallowPrinting: Integer);
begin
  DefaultInterface.Set_PDFDisallowPrinting(PDFDisallowPrinting);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFEncryptor: Integer;
begin
    Result := DefaultInterface.PDFEncryptor;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFEncryptor(PDFEncryptor: Integer);
begin
  DefaultInterface.Set_PDFEncryptor(PDFEncryptor);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFFontsEmbedAll: Integer;
begin
    Result := DefaultInterface.PDFFontsEmbedAll;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFFontsEmbedAll(PDFFontsEmbedAll: Integer);
begin
  DefaultInterface.Set_PDFFontsEmbedAll(PDFFontsEmbedAll);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFFontsSubSetFonts: Integer;
begin
    Result := DefaultInterface.PDFFontsSubSetFonts;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFFontsSubSetFonts(PDFFontsSubSetFonts: Integer);
begin
  DefaultInterface.Set_PDFFontsSubSetFonts(PDFFontsSubSetFonts);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFFontsSubSetFontsPercent: Integer;
begin
    Result := DefaultInterface.PDFFontsSubSetFontsPercent;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFFontsSubSetFontsPercent(PDFFontsSubSetFontsPercent: Integer);
begin
  DefaultInterface.Set_PDFFontsSubSetFontsPercent(PDFFontsSubSetFontsPercent);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFGeneralASCII85: Integer;
begin
    Result := DefaultInterface.PDFGeneralASCII85;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFGeneralASCII85(PDFGeneralASCII85: Integer);
begin
  DefaultInterface.Set_PDFGeneralASCII85(PDFGeneralASCII85);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFGeneralAutorotate: Integer;
begin
    Result := DefaultInterface.PDFGeneralAutorotate;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFGeneralAutorotate(PDFGeneralAutorotate: Integer);
begin
  DefaultInterface.Set_PDFGeneralAutorotate(PDFGeneralAutorotate);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFGeneralCompatibility: Integer;
begin
    Result := DefaultInterface.PDFGeneralCompatibility;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFGeneralCompatibility(PDFGeneralCompatibility: Integer);
begin
  DefaultInterface.Set_PDFGeneralCompatibility(PDFGeneralCompatibility);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFGeneralDefault: Integer;
begin
    Result := DefaultInterface.PDFGeneralDefault;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFGeneralDefault(PDFGeneralDefault: Integer);
begin
  DefaultInterface.Set_PDFGeneralDefault(PDFGeneralDefault);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFGeneralOverprint: Integer;
begin
    Result := DefaultInterface.PDFGeneralOverprint;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFGeneralOverprint(PDFGeneralOverprint: Integer);
begin
  DefaultInterface.Set_PDFGeneralOverprint(PDFGeneralOverprint);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFGeneralResolution: Integer;
begin
    Result := DefaultInterface.PDFGeneralResolution;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFGeneralResolution(PDFGeneralResolution: Integer);
begin
  DefaultInterface.Set_PDFGeneralResolution(PDFGeneralResolution);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFHighEncryption: Integer;
begin
    Result := DefaultInterface.PDFHighEncryption;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFHighEncryption(PDFHighEncryption: Integer);
begin
  DefaultInterface.Set_PDFHighEncryption(PDFHighEncryption);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFLowEncryption: Integer;
begin
    Result := DefaultInterface.PDFLowEncryption;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFLowEncryption(PDFLowEncryption: Integer);
begin
  DefaultInterface.Set_PDFLowEncryption(PDFLowEncryption);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFOptimize: Integer;
begin
    Result := DefaultInterface.PDFOptimize;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFOptimize(PDFOptimize: Integer);
begin
  DefaultInterface.Set_PDFOptimize(PDFOptimize);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFOwnerPass: Integer;
begin
    Result := DefaultInterface.PDFOwnerPass;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFOwnerPass(PDFOwnerPass: Integer);
begin
  DefaultInterface.Set_PDFOwnerPass(PDFOwnerPass);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFOwnerPasswordString: WideString;
begin
    Result := DefaultInterface.PDFOwnerPasswordString;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFOwnerPasswordString(const PDFOwnerPasswordString: WideString);
  { Avertissement : cette propriété PDFOwnerPasswordString a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.PDFOwnerPasswordString := PDFOwnerPasswordString;
end;

function TclsPDFCreatorOptionsProperties.Get_PDFSigningMultiSignature: Integer;
begin
    Result := DefaultInterface.PDFSigningMultiSignature;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFSigningMultiSignature(PDFSigningMultiSignature: Integer);
begin
  DefaultInterface.Set_PDFSigningMultiSignature(PDFSigningMultiSignature);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFSigningPFXFile: WideString;
begin
    Result := DefaultInterface.PDFSigningPFXFile;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFSigningPFXFile(const PDFSigningPFXFile: WideString);
  { Avertissement : cette propriété PDFSigningPFXFile a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.PDFSigningPFXFile := PDFSigningPFXFile;
end;

function TclsPDFCreatorOptionsProperties.Get_PDFSigningPFXFilePassword: WideString;
begin
    Result := DefaultInterface.PDFSigningPFXFilePassword;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFSigningPFXFilePassword(const PDFSigningPFXFilePassword: WideString);
  { Avertissement : cette propriété PDFSigningPFXFilePassword a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.PDFSigningPFXFilePassword := PDFSigningPFXFilePassword;
end;

function TclsPDFCreatorOptionsProperties.Get_PDFSigningSignatureContact: WideString;
begin
    Result := DefaultInterface.PDFSigningSignatureContact;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFSigningSignatureContact(const PDFSigningSignatureContact: WideString);
  { Avertissement : cette propriété PDFSigningSignatureContact a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.PDFSigningSignatureContact := PDFSigningSignatureContact;
end;

function TclsPDFCreatorOptionsProperties.Get_PDFSigningSignatureLeftX: Double;
begin
    Result := DefaultInterface.PDFSigningSignatureLeftX;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFSigningSignatureLeftX(PDFSigningSignatureLeftX: Double);
begin
  DefaultInterface.Set_PDFSigningSignatureLeftX(PDFSigningSignatureLeftX);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFSigningSignatureLeftY: Double;
begin
    Result := DefaultInterface.PDFSigningSignatureLeftY;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFSigningSignatureLeftY(PDFSigningSignatureLeftY: Double);
begin
  DefaultInterface.Set_PDFSigningSignatureLeftY(PDFSigningSignatureLeftY);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFSigningSignatureLocation: WideString;
begin
    Result := DefaultInterface.PDFSigningSignatureLocation;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFSigningSignatureLocation(const PDFSigningSignatureLocation: WideString);
  { Avertissement : cette propriété PDFSigningSignatureLocation a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.PDFSigningSignatureLocation := PDFSigningSignatureLocation;
end;

function TclsPDFCreatorOptionsProperties.Get_PDFSigningSignatureOnPage: Integer;
begin
    Result := DefaultInterface.PDFSigningSignatureOnPage;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFSigningSignatureOnPage(PDFSigningSignatureOnPage: Integer);
begin
  DefaultInterface.Set_PDFSigningSignatureOnPage(PDFSigningSignatureOnPage);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFSigningSignatureReason: WideString;
begin
    Result := DefaultInterface.PDFSigningSignatureReason;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFSigningSignatureReason(const PDFSigningSignatureReason: WideString);
  { Avertissement : cette propriété PDFSigningSignatureReason a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.PDFSigningSignatureReason := PDFSigningSignatureReason;
end;

function TclsPDFCreatorOptionsProperties.Get_PDFSigningSignatureRightX: Double;
begin
    Result := DefaultInterface.PDFSigningSignatureRightX;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFSigningSignatureRightX(PDFSigningSignatureRightX: Double);
begin
  DefaultInterface.Set_PDFSigningSignatureRightX(PDFSigningSignatureRightX);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFSigningSignatureRightY: Double;
begin
    Result := DefaultInterface.PDFSigningSignatureRightY;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFSigningSignatureRightY(PDFSigningSignatureRightY: Double);
begin
  DefaultInterface.Set_PDFSigningSignatureRightY(PDFSigningSignatureRightY);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFSigningSignatureVisible: Integer;
begin
    Result := DefaultInterface.PDFSigningSignatureVisible;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFSigningSignatureVisible(PDFSigningSignatureVisible: Integer);
begin
  DefaultInterface.Set_PDFSigningSignatureVisible(PDFSigningSignatureVisible);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFSigningSignPDF: Integer;
begin
    Result := DefaultInterface.PDFSigningSignPDF;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFSigningSignPDF(PDFSigningSignPDF: Integer);
begin
  DefaultInterface.Set_PDFSigningSignPDF(PDFSigningSignPDF);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFUpdateMetadata: Integer;
begin
    Result := DefaultInterface.PDFUpdateMetadata;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFUpdateMetadata(PDFUpdateMetadata: Integer);
begin
  DefaultInterface.Set_PDFUpdateMetadata(PDFUpdateMetadata);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFUserPass: Integer;
begin
    Result := DefaultInterface.PDFUserPass;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFUserPass(PDFUserPass: Integer);
begin
  DefaultInterface.Set_PDFUserPass(PDFUserPass);
end;

function TclsPDFCreatorOptionsProperties.Get_PDFUserPasswordString: WideString;
begin
    Result := DefaultInterface.PDFUserPasswordString;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFUserPasswordString(const PDFUserPasswordString: WideString);
  { Avertissement : cette propriété PDFUserPasswordString a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.PDFUserPasswordString := PDFUserPasswordString;
end;

function TclsPDFCreatorOptionsProperties.Get_PDFUseSecurity: Integer;
begin
    Result := DefaultInterface.PDFUseSecurity;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PDFUseSecurity(PDFUseSecurity: Integer);
begin
  DefaultInterface.Set_PDFUseSecurity(PDFUseSecurity);
end;

function TclsPDFCreatorOptionsProperties.Get_PNGColorscount: Integer;
begin
    Result := DefaultInterface.PNGColorscount;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PNGColorscount(PNGColorscount: Integer);
begin
  DefaultInterface.Set_PNGColorscount(PNGColorscount);
end;

function TclsPDFCreatorOptionsProperties.Get_PNGResolution: Integer;
begin
    Result := DefaultInterface.PNGResolution;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PNGResolution(PNGResolution: Integer);
begin
  DefaultInterface.Set_PNGResolution(PNGResolution);
end;

function TclsPDFCreatorOptionsProperties.Get_PrintAfterSaving: Integer;
begin
    Result := DefaultInterface.PrintAfterSaving;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PrintAfterSaving(PrintAfterSaving: Integer);
begin
  DefaultInterface.Set_PrintAfterSaving(PrintAfterSaving);
end;

function TclsPDFCreatorOptionsProperties.Get_PrintAfterSavingBitsPerPixel: Integer;
begin
    Result := DefaultInterface.PrintAfterSavingBitsPerPixel;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PrintAfterSavingBitsPerPixel(PrintAfterSavingBitsPerPixel: Integer);
begin
  DefaultInterface.Set_PrintAfterSavingBitsPerPixel(PrintAfterSavingBitsPerPixel);
end;

function TclsPDFCreatorOptionsProperties.Get_PrintAfterSavingDuplex: Integer;
begin
    Result := DefaultInterface.PrintAfterSavingDuplex;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PrintAfterSavingDuplex(PrintAfterSavingDuplex: Integer);
begin
  DefaultInterface.Set_PrintAfterSavingDuplex(PrintAfterSavingDuplex);
end;

function TclsPDFCreatorOptionsProperties.Get_PrintAfterSavingMaxResolution: Integer;
begin
    Result := DefaultInterface.PrintAfterSavingMaxResolution;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PrintAfterSavingMaxResolution(PrintAfterSavingMaxResolution: Integer);
begin
  DefaultInterface.Set_PrintAfterSavingMaxResolution(PrintAfterSavingMaxResolution);
end;

function TclsPDFCreatorOptionsProperties.Get_PrintAfterSavingMaxResolutionEnabled: Integer;
begin
    Result := DefaultInterface.PrintAfterSavingMaxResolutionEnabled;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PrintAfterSavingMaxResolutionEnabled(PrintAfterSavingMaxResolutionEnabled: Integer);
begin
  DefaultInterface.Set_PrintAfterSavingMaxResolutionEnabled(PrintAfterSavingMaxResolutionEnabled);
end;

function TclsPDFCreatorOptionsProperties.Get_PrintAfterSavingNoCancel: Integer;
begin
    Result := DefaultInterface.PrintAfterSavingNoCancel;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PrintAfterSavingNoCancel(PrintAfterSavingNoCancel: Integer);
begin
  DefaultInterface.Set_PrintAfterSavingNoCancel(PrintAfterSavingNoCancel);
end;

function TclsPDFCreatorOptionsProperties.Get_PrintAfterSavingPrinter: WideString;
begin
    Result := DefaultInterface.PrintAfterSavingPrinter;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PrintAfterSavingPrinter(const PrintAfterSavingPrinter: WideString);
  { Avertissement : cette propriété PrintAfterSavingPrinter a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.PrintAfterSavingPrinter := PrintAfterSavingPrinter;
end;

function TclsPDFCreatorOptionsProperties.Get_PrintAfterSavingQueryUser: Integer;
begin
    Result := DefaultInterface.PrintAfterSavingQueryUser;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PrintAfterSavingQueryUser(PrintAfterSavingQueryUser: Integer);
begin
  DefaultInterface.Set_PrintAfterSavingQueryUser(PrintAfterSavingQueryUser);
end;

function TclsPDFCreatorOptionsProperties.Get_PrintAfterSavingTumble: Integer;
begin
    Result := DefaultInterface.PrintAfterSavingTumble;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PrintAfterSavingTumble(PrintAfterSavingTumble: Integer);
begin
  DefaultInterface.Set_PrintAfterSavingTumble(PrintAfterSavingTumble);
end;

function TclsPDFCreatorOptionsProperties.Get_PrinterStop: Integer;
begin
    Result := DefaultInterface.PrinterStop;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PrinterStop(PrinterStop: Integer);
begin
  DefaultInterface.Set_PrinterStop(PrinterStop);
end;

function TclsPDFCreatorOptionsProperties.Get_PrinterTemppath: WideString;
begin
    Result := DefaultInterface.PrinterTemppath;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PrinterTemppath(const PrinterTemppath: WideString);
  { Avertissement : cette propriété PrinterTemppath a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.PrinterTemppath := PrinterTemppath;
end;

function TclsPDFCreatorOptionsProperties.Get_ProcessPriority: Integer;
begin
    Result := DefaultInterface.ProcessPriority;
end;

procedure TclsPDFCreatorOptionsProperties.Set_ProcessPriority(ProcessPriority: Integer);
begin
  DefaultInterface.Set_ProcessPriority(ProcessPriority);
end;

function TclsPDFCreatorOptionsProperties.Get_ProgramFont: WideString;
begin
    Result := DefaultInterface.ProgramFont;
end;

procedure TclsPDFCreatorOptionsProperties.Set_ProgramFont(const ProgramFont: WideString);
  { Avertissement : cette propriété ProgramFont a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.ProgramFont := ProgramFont;
end;

function TclsPDFCreatorOptionsProperties.Get_ProgramFontCharset: Integer;
begin
    Result := DefaultInterface.ProgramFontCharset;
end;

procedure TclsPDFCreatorOptionsProperties.Set_ProgramFontCharset(ProgramFontCharset: Integer);
begin
  DefaultInterface.Set_ProgramFontCharset(ProgramFontCharset);
end;

function TclsPDFCreatorOptionsProperties.Get_ProgramFontSize: Integer;
begin
    Result := DefaultInterface.ProgramFontSize;
end;

procedure TclsPDFCreatorOptionsProperties.Set_ProgramFontSize(ProgramFontSize: Integer);
begin
  DefaultInterface.Set_ProgramFontSize(ProgramFontSize);
end;

function TclsPDFCreatorOptionsProperties.Get_PSDColorsCount: Integer;
begin
    Result := DefaultInterface.PSDColorsCount;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PSDColorsCount(PSDColorsCount: Integer);
begin
  DefaultInterface.Set_PSDColorsCount(PSDColorsCount);
end;

function TclsPDFCreatorOptionsProperties.Get_PSDResolution: Integer;
begin
    Result := DefaultInterface.PSDResolution;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PSDResolution(PSDResolution: Integer);
begin
  DefaultInterface.Set_PSDResolution(PSDResolution);
end;

function TclsPDFCreatorOptionsProperties.Get_PSLanguageLevel: Integer;
begin
    Result := DefaultInterface.PSLanguageLevel;
end;

procedure TclsPDFCreatorOptionsProperties.Set_PSLanguageLevel(PSLanguageLevel: Integer);
begin
  DefaultInterface.Set_PSLanguageLevel(PSLanguageLevel);
end;

function TclsPDFCreatorOptionsProperties.Get_RAWColorsCount: Integer;
begin
    Result := DefaultInterface.RAWColorsCount;
end;

procedure TclsPDFCreatorOptionsProperties.Set_RAWColorsCount(RAWColorsCount: Integer);
begin
  DefaultInterface.Set_RAWColorsCount(RAWColorsCount);
end;

function TclsPDFCreatorOptionsProperties.Get_RAWResolution: Integer;
begin
    Result := DefaultInterface.RAWResolution;
end;

procedure TclsPDFCreatorOptionsProperties.Set_RAWResolution(RAWResolution: Integer);
begin
  DefaultInterface.Set_RAWResolution(RAWResolution);
end;

function TclsPDFCreatorOptionsProperties.Get_RemoveAllKnownFileExtensions: Integer;
begin
    Result := DefaultInterface.RemoveAllKnownFileExtensions;
end;

procedure TclsPDFCreatorOptionsProperties.Set_RemoveAllKnownFileExtensions(RemoveAllKnownFileExtensions: Integer);
begin
  DefaultInterface.Set_RemoveAllKnownFileExtensions(RemoveAllKnownFileExtensions);
end;

function TclsPDFCreatorOptionsProperties.Get_RemoveSpaces: Integer;
begin
    Result := DefaultInterface.RemoveSpaces;
end;

procedure TclsPDFCreatorOptionsProperties.Set_RemoveSpaces(RemoveSpaces: Integer);
begin
  DefaultInterface.Set_RemoveSpaces(RemoveSpaces);
end;

function TclsPDFCreatorOptionsProperties.Get_RunProgramAfterSaving: Integer;
begin
    Result := DefaultInterface.RunProgramAfterSaving;
end;

procedure TclsPDFCreatorOptionsProperties.Set_RunProgramAfterSaving(RunProgramAfterSaving: Integer);
begin
  DefaultInterface.Set_RunProgramAfterSaving(RunProgramAfterSaving);
end;

function TclsPDFCreatorOptionsProperties.Get_RunProgramAfterSavingProgramname: WideString;
begin
    Result := DefaultInterface.RunProgramAfterSavingProgramname;
end;

procedure TclsPDFCreatorOptionsProperties.Set_RunProgramAfterSavingProgramname(const RunProgramAfterSavingProgramname: WideString);
  { Avertissement : cette propriété RunProgramAfterSavingProgramname a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.RunProgramAfterSavingProgramname := RunProgramAfterSavingProgramname;
end;

function TclsPDFCreatorOptionsProperties.Get_RunProgramAfterSavingProgramParameters: WideString;
begin
    Result := DefaultInterface.RunProgramAfterSavingProgramParameters;
end;

procedure TclsPDFCreatorOptionsProperties.Set_RunProgramAfterSavingProgramParameters(const RunProgramAfterSavingProgramParameters: WideString);
  { Avertissement : cette propriété RunProgramAfterSavingProgramParameters a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.RunProgramAfterSavingProgramParameters := RunProgramAfterSavingProgramParameters;
end;

function TclsPDFCreatorOptionsProperties.Get_RunProgramAfterSavingWaitUntilReady: Integer;
begin
    Result := DefaultInterface.RunProgramAfterSavingWaitUntilReady;
end;

procedure TclsPDFCreatorOptionsProperties.Set_RunProgramAfterSavingWaitUntilReady(RunProgramAfterSavingWaitUntilReady: Integer);
begin
  DefaultInterface.Set_RunProgramAfterSavingWaitUntilReady(RunProgramAfterSavingWaitUntilReady);
end;

function TclsPDFCreatorOptionsProperties.Get_RunProgramAfterSavingWindowstyle: Integer;
begin
    Result := DefaultInterface.RunProgramAfterSavingWindowstyle;
end;

procedure TclsPDFCreatorOptionsProperties.Set_RunProgramAfterSavingWindowstyle(RunProgramAfterSavingWindowstyle: Integer);
begin
  DefaultInterface.Set_RunProgramAfterSavingWindowstyle(RunProgramAfterSavingWindowstyle);
end;

function TclsPDFCreatorOptionsProperties.Get_RunProgramBeforeSaving: Integer;
begin
    Result := DefaultInterface.RunProgramBeforeSaving;
end;

procedure TclsPDFCreatorOptionsProperties.Set_RunProgramBeforeSaving(RunProgramBeforeSaving: Integer);
begin
  DefaultInterface.Set_RunProgramBeforeSaving(RunProgramBeforeSaving);
end;

function TclsPDFCreatorOptionsProperties.Get_RunProgramBeforeSavingProgramname: WideString;
begin
    Result := DefaultInterface.RunProgramBeforeSavingProgramname;
end;

procedure TclsPDFCreatorOptionsProperties.Set_RunProgramBeforeSavingProgramname(const RunProgramBeforeSavingProgramname: WideString);
  { Avertissement : cette propriété RunProgramBeforeSavingProgramname a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.RunProgramBeforeSavingProgramname := RunProgramBeforeSavingProgramname;
end;

function TclsPDFCreatorOptionsProperties.Get_RunProgramBeforeSavingProgramParameters: WideString;
begin
    Result := DefaultInterface.RunProgramBeforeSavingProgramParameters;
end;

procedure TclsPDFCreatorOptionsProperties.Set_RunProgramBeforeSavingProgramParameters(const RunProgramBeforeSavingProgramParameters: WideString);
  { Avertissement : cette propriété RunProgramBeforeSavingProgramParameters a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.RunProgramBeforeSavingProgramParameters := RunProgramBeforeSavingProgramParameters;
end;

function TclsPDFCreatorOptionsProperties.Get_RunProgramBeforeSavingWindowstyle: Integer;
begin
    Result := DefaultInterface.RunProgramBeforeSavingWindowstyle;
end;

procedure TclsPDFCreatorOptionsProperties.Set_RunProgramBeforeSavingWindowstyle(RunProgramBeforeSavingWindowstyle: Integer);
begin
  DefaultInterface.Set_RunProgramBeforeSavingWindowstyle(RunProgramBeforeSavingWindowstyle);
end;

function TclsPDFCreatorOptionsProperties.Get_SaveFilename: WideString;
begin
    Result := DefaultInterface.SaveFilename;
end;

procedure TclsPDFCreatorOptionsProperties.Set_SaveFilename(const SaveFilename: WideString);
  { Avertissement : cette propriété SaveFilename a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.SaveFilename := SaveFilename;
end;

function TclsPDFCreatorOptionsProperties.Get_SendEmailAfterAutoSaving: Integer;
begin
    Result := DefaultInterface.SendEmailAfterAutoSaving;
end;

procedure TclsPDFCreatorOptionsProperties.Set_SendEmailAfterAutoSaving(SendEmailAfterAutoSaving: Integer);
begin
  DefaultInterface.Set_SendEmailAfterAutoSaving(SendEmailAfterAutoSaving);
end;

function TclsPDFCreatorOptionsProperties.Get_SendMailMethod: Integer;
begin
    Result := DefaultInterface.SendMailMethod;
end;

procedure TclsPDFCreatorOptionsProperties.Set_SendMailMethod(SendMailMethod: Integer);
begin
  DefaultInterface.Set_SendMailMethod(SendMailMethod);
end;

function TclsPDFCreatorOptionsProperties.Get_ShowAnimation: Integer;
begin
    Result := DefaultInterface.ShowAnimation;
end;

procedure TclsPDFCreatorOptionsProperties.Set_ShowAnimation(ShowAnimation: Integer);
begin
  DefaultInterface.Set_ShowAnimation(ShowAnimation);
end;

function TclsPDFCreatorOptionsProperties.Get_StampFontColor: WideString;
begin
    Result := DefaultInterface.StampFontColor;
end;

procedure TclsPDFCreatorOptionsProperties.Set_StampFontColor(const StampFontColor: WideString);
  { Avertissement : cette propriété StampFontColor a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StampFontColor := StampFontColor;
end;

function TclsPDFCreatorOptionsProperties.Get_StampFontname: WideString;
begin
    Result := DefaultInterface.StampFontname;
end;

procedure TclsPDFCreatorOptionsProperties.Set_StampFontname(const StampFontname: WideString);
  { Avertissement : cette propriété StampFontname a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StampFontname := StampFontname;
end;

function TclsPDFCreatorOptionsProperties.Get_StampFontsize: Integer;
begin
    Result := DefaultInterface.StampFontsize;
end;

procedure TclsPDFCreatorOptionsProperties.Set_StampFontsize(StampFontsize: Integer);
begin
  DefaultInterface.Set_StampFontsize(StampFontsize);
end;

function TclsPDFCreatorOptionsProperties.Get_StampOutlineFontthickness: Integer;
begin
    Result := DefaultInterface.StampOutlineFontthickness;
end;

procedure TclsPDFCreatorOptionsProperties.Set_StampOutlineFontthickness(StampOutlineFontthickness: Integer);
begin
  DefaultInterface.Set_StampOutlineFontthickness(StampOutlineFontthickness);
end;

function TclsPDFCreatorOptionsProperties.Get_StampString: WideString;
begin
    Result := DefaultInterface.StampString;
end;

procedure TclsPDFCreatorOptionsProperties.Set_StampString(const StampString: WideString);
  { Avertissement : cette propriété StampString a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StampString := StampString;
end;

function TclsPDFCreatorOptionsProperties.Get_StampUseOutlineFont: Integer;
begin
    Result := DefaultInterface.StampUseOutlineFont;
end;

procedure TclsPDFCreatorOptionsProperties.Set_StampUseOutlineFont(StampUseOutlineFont: Integer);
begin
  DefaultInterface.Set_StampUseOutlineFont(StampUseOutlineFont);
end;

function TclsPDFCreatorOptionsProperties.Get_StandardAuthor: WideString;
begin
    Result := DefaultInterface.StandardAuthor;
end;

procedure TclsPDFCreatorOptionsProperties.Set_StandardAuthor(const StandardAuthor: WideString);
  { Avertissement : cette propriété StandardAuthor a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StandardAuthor := StandardAuthor;
end;

function TclsPDFCreatorOptionsProperties.Get_StandardCreationdate: WideString;
begin
    Result := DefaultInterface.StandardCreationdate;
end;

procedure TclsPDFCreatorOptionsProperties.Set_StandardCreationdate(const StandardCreationdate: WideString);
  { Avertissement : cette propriété StandardCreationdate a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StandardCreationdate := StandardCreationdate;
end;

function TclsPDFCreatorOptionsProperties.Get_StandardDateformat: WideString;
begin
    Result := DefaultInterface.StandardDateformat;
end;

procedure TclsPDFCreatorOptionsProperties.Set_StandardDateformat(const StandardDateformat: WideString);
  { Avertissement : cette propriété StandardDateformat a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StandardDateformat := StandardDateformat;
end;

function TclsPDFCreatorOptionsProperties.Get_StandardKeywords: WideString;
begin
    Result := DefaultInterface.StandardKeywords;
end;

procedure TclsPDFCreatorOptionsProperties.Set_StandardKeywords(const StandardKeywords: WideString);
  { Avertissement : cette propriété StandardKeywords a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StandardKeywords := StandardKeywords;
end;

function TclsPDFCreatorOptionsProperties.Get_StandardMailDomain: WideString;
begin
    Result := DefaultInterface.StandardMailDomain;
end;

procedure TclsPDFCreatorOptionsProperties.Set_StandardMailDomain(const StandardMailDomain: WideString);
  { Avertissement : cette propriété StandardMailDomain a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StandardMailDomain := StandardMailDomain;
end;

function TclsPDFCreatorOptionsProperties.Get_StandardModifydate: WideString;
begin
    Result := DefaultInterface.StandardModifydate;
end;

procedure TclsPDFCreatorOptionsProperties.Set_StandardModifydate(const StandardModifydate: WideString);
  { Avertissement : cette propriété StandardModifydate a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StandardModifydate := StandardModifydate;
end;

function TclsPDFCreatorOptionsProperties.Get_StandardSaveformat: Integer;
begin
    Result := DefaultInterface.StandardSaveformat;
end;

procedure TclsPDFCreatorOptionsProperties.Set_StandardSaveformat(StandardSaveformat: Integer);
begin
  DefaultInterface.Set_StandardSaveformat(StandardSaveformat);
end;

function TclsPDFCreatorOptionsProperties.Get_StandardSubject: WideString;
begin
    Result := DefaultInterface.StandardSubject;
end;

procedure TclsPDFCreatorOptionsProperties.Set_StandardSubject(const StandardSubject: WideString);
  { Avertissement : cette propriété StandardSubject a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StandardSubject := StandardSubject;
end;

function TclsPDFCreatorOptionsProperties.Get_StandardTitle: WideString;
begin
    Result := DefaultInterface.StandardTitle;
end;

procedure TclsPDFCreatorOptionsProperties.Set_StandardTitle(const StandardTitle: WideString);
  { Avertissement : cette propriété StandardTitle a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.StandardTitle := StandardTitle;
end;

function TclsPDFCreatorOptionsProperties.Get_StartStandardProgram: Integer;
begin
    Result := DefaultInterface.StartStandardProgram;
end;

procedure TclsPDFCreatorOptionsProperties.Set_StartStandardProgram(StartStandardProgram: Integer);
begin
  DefaultInterface.Set_StartStandardProgram(StartStandardProgram);
end;

function TclsPDFCreatorOptionsProperties.Get_SVGResolution: Integer;
begin
    Result := DefaultInterface.SVGResolution;
end;

procedure TclsPDFCreatorOptionsProperties.Set_SVGResolution(SVGResolution: Integer);
begin
  DefaultInterface.Set_SVGResolution(SVGResolution);
end;

function TclsPDFCreatorOptionsProperties.Get_TIFFColorscount: Integer;
begin
    Result := DefaultInterface.TIFFColorscount;
end;

procedure TclsPDFCreatorOptionsProperties.Set_TIFFColorscount(TIFFColorscount: Integer);
begin
  DefaultInterface.Set_TIFFColorscount(TIFFColorscount);
end;

function TclsPDFCreatorOptionsProperties.Get_TIFFResolution: Integer;
begin
    Result := DefaultInterface.TIFFResolution;
end;

procedure TclsPDFCreatorOptionsProperties.Set_TIFFResolution(TIFFResolution: Integer);
begin
  DefaultInterface.Set_TIFFResolution(TIFFResolution);
end;

function TclsPDFCreatorOptionsProperties.Get_Toolbars: Integer;
begin
    Result := DefaultInterface.Toolbars;
end;

procedure TclsPDFCreatorOptionsProperties.Set_Toolbars(Toolbars: Integer);
begin
  DefaultInterface.Set_Toolbars(Toolbars);
end;

function TclsPDFCreatorOptionsProperties.Get_UpdateInterval: Integer;
begin
    Result := DefaultInterface.UpdateInterval;
end;

procedure TclsPDFCreatorOptionsProperties.Set_UpdateInterval(UpdateInterval: Integer);
begin
  DefaultInterface.Set_UpdateInterval(UpdateInterval);
end;

function TclsPDFCreatorOptionsProperties.Get_UseAutosave: Integer;
begin
    Result := DefaultInterface.UseAutosave;
end;

procedure TclsPDFCreatorOptionsProperties.Set_UseAutosave(UseAutosave: Integer);
begin
  DefaultInterface.Set_UseAutosave(UseAutosave);
end;

function TclsPDFCreatorOptionsProperties.Get_UseAutosaveDirectory: Integer;
begin
    Result := DefaultInterface.UseAutosaveDirectory;
end;

procedure TclsPDFCreatorOptionsProperties.Set_UseAutosaveDirectory(UseAutosaveDirectory: Integer);
begin
  DefaultInterface.Set_UseAutosaveDirectory(UseAutosaveDirectory);
end;

function TclsPDFCreatorOptionsProperties.Get_UseCreationDateNow: Integer;
begin
    Result := DefaultInterface.UseCreationDateNow;
end;

procedure TclsPDFCreatorOptionsProperties.Set_UseCreationDateNow(UseCreationDateNow: Integer);
begin
  DefaultInterface.Set_UseCreationDateNow(UseCreationDateNow);
end;

function TclsPDFCreatorOptionsProperties.Get_UseCustomPaperSize: WideString;
begin
    Result := DefaultInterface.UseCustomPaperSize;
end;

procedure TclsPDFCreatorOptionsProperties.Set_UseCustomPaperSize(const UseCustomPaperSize: WideString);
  { Avertissement : cette propriété UseCustomPaperSize a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.UseCustomPaperSize := UseCustomPaperSize;
end;

function TclsPDFCreatorOptionsProperties.Get_UseFixPapersize: Integer;
begin
    Result := DefaultInterface.UseFixPapersize;
end;

procedure TclsPDFCreatorOptionsProperties.Set_UseFixPapersize(UseFixPapersize: Integer);
begin
  DefaultInterface.Set_UseFixPapersize(UseFixPapersize);
end;

function TclsPDFCreatorOptionsProperties.Get_UseStandardAuthor: Integer;
begin
    Result := DefaultInterface.UseStandardAuthor;
end;

procedure TclsPDFCreatorOptionsProperties.Set_UseStandardAuthor(UseStandardAuthor: Integer);
begin
  DefaultInterface.Set_UseStandardAuthor(UseStandardAuthor);
end;

{$ENDIF}

class function CoclsPDFCreatorError.Create: _clsPDFCreatorError;
begin
  Result := CreateComObject(CLASS_clsPDFCreatorError) as _clsPDFCreatorError;
end;

class function CoclsPDFCreatorError.CreateRemote(const MachineName: string): _clsPDFCreatorError;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_clsPDFCreatorError) as _clsPDFCreatorError;
end;

procedure TclsPDFCreatorError.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{84D26557-2990-4B3E-A99F-C4DC1CB6C225}';
    IntfIID:   '{A030F401-6045-4942-A5F5-9CCBF2C1872D}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TclsPDFCreatorError.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as _clsPDFCreatorError;
  end;
end;

procedure TclsPDFCreatorError.ConnectTo(svrIntf: _clsPDFCreatorError);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TclsPDFCreatorError.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TclsPDFCreatorError.GetDefaultInterface: _clsPDFCreatorError;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface est NULL. Le composant n''est pas connecté au serveur. Vous devez appeler ''Connect'' ou ''ConnectTo'' avant cette opération');
  Result := FIntf;
end;

constructor TclsPDFCreatorError.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TclsPDFCreatorErrorProperties.Create(Self);
{$ENDIF}
end;

destructor TclsPDFCreatorError.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TclsPDFCreatorError.GetServerProperties: TclsPDFCreatorErrorProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TclsPDFCreatorError.Get_Number: Integer;
begin
    Result := DefaultInterface.Number;
end;

procedure TclsPDFCreatorError.Set_Number(Number: Integer);
begin
  DefaultInterface.Set_Number(Number);
end;

function TclsPDFCreatorError.Get_Description: WideString;
begin
    Result := DefaultInterface.Description;
end;

procedure TclsPDFCreatorError.Set_Description(const Description: WideString);
  { Avertissement : cette propriété Description a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.Description := Description;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TclsPDFCreatorErrorProperties.Create(AServer: TclsPDFCreatorError);
begin
  inherited Create;
  FServer := AServer;
end;

function TclsPDFCreatorErrorProperties.GetDefaultInterface: _clsPDFCreatorError;
begin
  Result := FServer.DefaultInterface;
end;

function TclsPDFCreatorErrorProperties.Get_Number: Integer;
begin
    Result := DefaultInterface.Number;
end;

procedure TclsPDFCreatorErrorProperties.Set_Number(Number: Integer);
begin
  DefaultInterface.Set_Number(Number);
end;

function TclsPDFCreatorErrorProperties.Get_Description: WideString;
begin
    Result := DefaultInterface.Description;
end;

procedure TclsPDFCreatorErrorProperties.Set_Description(const Description: WideString);
  { Avertissement : cette propriété Description a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.Description := Description;
end;

{$ENDIF}

class function CoclsPDFCreatorInfoSpoolFile.Create: _clsPDFCreatorInfoSpoolFile;
begin
  Result := CreateComObject(CLASS_clsPDFCreatorInfoSpoolFile) as _clsPDFCreatorInfoSpoolFile;
end;

class function CoclsPDFCreatorInfoSpoolFile.CreateRemote(const MachineName: string): _clsPDFCreatorInfoSpoolFile;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_clsPDFCreatorInfoSpoolFile) as _clsPDFCreatorInfoSpoolFile;
end;

procedure TclsPDFCreatorInfoSpoolFile.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{411DBF4B-0B78-4C28-9997-ECD80CC371C4}';
    IntfIID:   '{253F17D9-1678-4B0E-843E-A2D37C2C6B4E}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TclsPDFCreatorInfoSpoolFile.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as _clsPDFCreatorInfoSpoolFile;
  end;
end;

procedure TclsPDFCreatorInfoSpoolFile.ConnectTo(svrIntf: _clsPDFCreatorInfoSpoolFile);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TclsPDFCreatorInfoSpoolFile.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TclsPDFCreatorInfoSpoolFile.GetDefaultInterface: _clsPDFCreatorInfoSpoolFile;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface est NULL. Le composant n''est pas connecté au serveur. Vous devez appeler ''Connect'' ou ''ConnectTo'' avant cette opération');
  Result := FIntf;
end;

constructor TclsPDFCreatorInfoSpoolFile.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TclsPDFCreatorInfoSpoolFileProperties.Create(Self);
{$ENDIF}
end;

destructor TclsPDFCreatorInfoSpoolFile.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TclsPDFCreatorInfoSpoolFile.GetServerProperties: TclsPDFCreatorInfoSpoolFileProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TclsPDFCreatorInfoSpoolFile.Get_REDMON_PORT: WideString;
begin
    Result := DefaultInterface.REDMON_PORT;
end;

procedure TclsPDFCreatorInfoSpoolFile.Set_REDMON_PORT(const REDMON_PORT: WideString);
  { Avertissement : cette propriété REDMON_PORT a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.REDMON_PORT := REDMON_PORT;
end;

function TclsPDFCreatorInfoSpoolFile.Get_REDMON_JOB: WideString;
begin
    Result := DefaultInterface.REDMON_JOB;
end;

procedure TclsPDFCreatorInfoSpoolFile.Set_REDMON_JOB(const REDMON_JOB: WideString);
  { Avertissement : cette propriété REDMON_JOB a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.REDMON_JOB := REDMON_JOB;
end;

function TclsPDFCreatorInfoSpoolFile.Get_REDMON_PRINTER: WideString;
begin
    Result := DefaultInterface.REDMON_PRINTER;
end;

procedure TclsPDFCreatorInfoSpoolFile.Set_REDMON_PRINTER(const REDMON_PRINTER: WideString);
  { Avertissement : cette propriété REDMON_PRINTER a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.REDMON_PRINTER := REDMON_PRINTER;
end;

function TclsPDFCreatorInfoSpoolFile.Get_REDMON_MACHINE: WideString;
begin
    Result := DefaultInterface.REDMON_MACHINE;
end;

procedure TclsPDFCreatorInfoSpoolFile.Set_REDMON_MACHINE(const REDMON_MACHINE: WideString);
  { Avertissement : cette propriété REDMON_MACHINE a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.REDMON_MACHINE := REDMON_MACHINE;
end;

function TclsPDFCreatorInfoSpoolFile.Get_REDMON_USER: WideString;
begin
    Result := DefaultInterface.REDMON_USER;
end;

procedure TclsPDFCreatorInfoSpoolFile.Set_REDMON_USER(const REDMON_USER: WideString);
  { Avertissement : cette propriété REDMON_USER a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.REDMON_USER := REDMON_USER;
end;

function TclsPDFCreatorInfoSpoolFile.Get_REDMON_DOCNAME: WideString;
begin
    Result := DefaultInterface.REDMON_DOCNAME;
end;

procedure TclsPDFCreatorInfoSpoolFile.Set_REDMON_DOCNAME(const REDMON_DOCNAME: WideString);
  { Avertissement : cette propriété REDMON_DOCNAME a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.REDMON_DOCNAME := REDMON_DOCNAME;
end;

function TclsPDFCreatorInfoSpoolFile.Get_REDMON_FILENAME: WideString;
begin
    Result := DefaultInterface.REDMON_FILENAME;
end;

procedure TclsPDFCreatorInfoSpoolFile.Set_REDMON_FILENAME(const REDMON_FILENAME: WideString);
  { Avertissement : cette propriété REDMON_FILENAME a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.REDMON_FILENAME := REDMON_FILENAME;
end;

function TclsPDFCreatorInfoSpoolFile.Get_REDMON_SESSIONID: WideString;
begin
    Result := DefaultInterface.REDMON_SESSIONID;
end;

procedure TclsPDFCreatorInfoSpoolFile.Set_REDMON_SESSIONID(const REDMON_SESSIONID: WideString);
  { Avertissement : cette propriété REDMON_SESSIONID a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.REDMON_SESSIONID := REDMON_SESSIONID;
end;

function TclsPDFCreatorInfoSpoolFile.Get_SpoolFilename: WideString;
begin
    Result := DefaultInterface.SpoolFilename;
end;

procedure TclsPDFCreatorInfoSpoolFile.Set_SpoolFilename(const SpoolFilename: WideString);
  { Avertissement : cette propriété SpoolFilename a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.SpoolFilename := SpoolFilename;
end;

function TclsPDFCreatorInfoSpoolFile.Get_SpoolerAccount: WideString;
begin
    Result := DefaultInterface.SpoolerAccount;
end;

procedure TclsPDFCreatorInfoSpoolFile.Set_SpoolerAccount(const SpoolerAccount: WideString);
  { Avertissement : cette propriété SpoolerAccount a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.SpoolerAccount := SpoolerAccount;
end;

function TclsPDFCreatorInfoSpoolFile.Get_Computer: WideString;
begin
    Result := DefaultInterface.Computer;
end;

procedure TclsPDFCreatorInfoSpoolFile.Set_Computer(const Computer: WideString);
  { Avertissement : cette propriété Computer a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.Computer := Computer;
end;

function TclsPDFCreatorInfoSpoolFile.Get_Created: WideString;
begin
    Result := DefaultInterface.Created;
end;

procedure TclsPDFCreatorInfoSpoolFile.Set_Created(const Created: WideString);
  { Avertissement : cette propriété Created a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.Created := Created;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TclsPDFCreatorInfoSpoolFileProperties.Create(AServer: TclsPDFCreatorInfoSpoolFile);
begin
  inherited Create;
  FServer := AServer;
end;

function TclsPDFCreatorInfoSpoolFileProperties.GetDefaultInterface: _clsPDFCreatorInfoSpoolFile;
begin
  Result := FServer.DefaultInterface;
end;

function TclsPDFCreatorInfoSpoolFileProperties.Get_REDMON_PORT: WideString;
begin
    Result := DefaultInterface.REDMON_PORT;
end;

procedure TclsPDFCreatorInfoSpoolFileProperties.Set_REDMON_PORT(const REDMON_PORT: WideString);
  { Avertissement : cette propriété REDMON_PORT a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.REDMON_PORT := REDMON_PORT;
end;

function TclsPDFCreatorInfoSpoolFileProperties.Get_REDMON_JOB: WideString;
begin
    Result := DefaultInterface.REDMON_JOB;
end;

procedure TclsPDFCreatorInfoSpoolFileProperties.Set_REDMON_JOB(const REDMON_JOB: WideString);
  { Avertissement : cette propriété REDMON_JOB a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.REDMON_JOB := REDMON_JOB;
end;

function TclsPDFCreatorInfoSpoolFileProperties.Get_REDMON_PRINTER: WideString;
begin
    Result := DefaultInterface.REDMON_PRINTER;
end;

procedure TclsPDFCreatorInfoSpoolFileProperties.Set_REDMON_PRINTER(const REDMON_PRINTER: WideString);
  { Avertissement : cette propriété REDMON_PRINTER a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.REDMON_PRINTER := REDMON_PRINTER;
end;

function TclsPDFCreatorInfoSpoolFileProperties.Get_REDMON_MACHINE: WideString;
begin
    Result := DefaultInterface.REDMON_MACHINE;
end;

procedure TclsPDFCreatorInfoSpoolFileProperties.Set_REDMON_MACHINE(const REDMON_MACHINE: WideString);
  { Avertissement : cette propriété REDMON_MACHINE a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.REDMON_MACHINE := REDMON_MACHINE;
end;

function TclsPDFCreatorInfoSpoolFileProperties.Get_REDMON_USER: WideString;
begin
    Result := DefaultInterface.REDMON_USER;
end;

procedure TclsPDFCreatorInfoSpoolFileProperties.Set_REDMON_USER(const REDMON_USER: WideString);
  { Avertissement : cette propriété REDMON_USER a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.REDMON_USER := REDMON_USER;
end;

function TclsPDFCreatorInfoSpoolFileProperties.Get_REDMON_DOCNAME: WideString;
begin
    Result := DefaultInterface.REDMON_DOCNAME;
end;

procedure TclsPDFCreatorInfoSpoolFileProperties.Set_REDMON_DOCNAME(const REDMON_DOCNAME: WideString);
  { Avertissement : cette propriété REDMON_DOCNAME a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.REDMON_DOCNAME := REDMON_DOCNAME;
end;

function TclsPDFCreatorInfoSpoolFileProperties.Get_REDMON_FILENAME: WideString;
begin
    Result := DefaultInterface.REDMON_FILENAME;
end;

procedure TclsPDFCreatorInfoSpoolFileProperties.Set_REDMON_FILENAME(const REDMON_FILENAME: WideString);
  { Avertissement : cette propriété REDMON_FILENAME a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.REDMON_FILENAME := REDMON_FILENAME;
end;

function TclsPDFCreatorInfoSpoolFileProperties.Get_REDMON_SESSIONID: WideString;
begin
    Result := DefaultInterface.REDMON_SESSIONID;
end;

procedure TclsPDFCreatorInfoSpoolFileProperties.Set_REDMON_SESSIONID(const REDMON_SESSIONID: WideString);
  { Avertissement : cette propriété REDMON_SESSIONID a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.REDMON_SESSIONID := REDMON_SESSIONID;
end;

function TclsPDFCreatorInfoSpoolFileProperties.Get_SpoolFilename: WideString;
begin
    Result := DefaultInterface.SpoolFilename;
end;

procedure TclsPDFCreatorInfoSpoolFileProperties.Set_SpoolFilename(const SpoolFilename: WideString);
  { Avertissement : cette propriété SpoolFilename a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.SpoolFilename := SpoolFilename;
end;

function TclsPDFCreatorInfoSpoolFileProperties.Get_SpoolerAccount: WideString;
begin
    Result := DefaultInterface.SpoolerAccount;
end;

procedure TclsPDFCreatorInfoSpoolFileProperties.Set_SpoolerAccount(const SpoolerAccount: WideString);
  { Avertissement : cette propriété SpoolerAccount a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.SpoolerAccount := SpoolerAccount;
end;

function TclsPDFCreatorInfoSpoolFileProperties.Get_Computer: WideString;
begin
    Result := DefaultInterface.Computer;
end;

procedure TclsPDFCreatorInfoSpoolFileProperties.Set_Computer(const Computer: WideString);
  { Avertissement : cette propriété Computer a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.Computer := Computer;
end;

function TclsPDFCreatorInfoSpoolFileProperties.Get_Created: WideString;
begin
    Result := DefaultInterface.Created;
end;

procedure TclsPDFCreatorInfoSpoolFileProperties.Set_Created(const Created: WideString);
  { Avertissement : cette propriété Created a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.Created := Created;
end;

{$ENDIF}

class function CoclsTools.Create: _clsTools;
begin
  Result := CreateComObject(CLASS_clsTools) as _clsTools;
end;

class function CoclsTools.CreateRemote(const MachineName: string): _clsTools;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_clsTools) as _clsTools;
end;

procedure TclsTools.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{718E1546-3738-4DC8-AA9E-A1BDE33237FE}';
    IntfIID:   '{58CCA1EC-FC65-4750-A7E7-FFA191BD1042}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TclsTools.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as _clsTools;
  end;
end;

procedure TclsTools.ConnectTo(svrIntf: _clsTools);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TclsTools.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TclsTools.GetDefaultInterface: _clsTools;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface est NULL. Le composant n''est pas connecté au serveur. Vous devez appeler ''Connect'' ou ''ConnectTo'' avant cette opération');
  Result := FIntf;
end;

constructor TclsTools.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TclsToolsProperties.Create(Self);
{$ENDIF}
end;

destructor TclsTools.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TclsTools.GetServerProperties: TclsToolsProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TclsTools.cOpenFileDialog(var files: OleVariant; const InitFilename: WideString; 
                                   const Filter: WideString; 
                                   const DefaultFileExtension: WideString; 
                                   const InitDir: WideString; const DialogTitle: WideString; 
                                   Flags: Integer; hwnd: Integer; FilterIndex: Integer): Integer;
begin
  Result := DefaultInterface.cOpenFileDialog(files, InitFilename, Filter, DefaultFileExtension, 
                                             InitDir, DialogTitle, Flags, hwnd, FilterIndex);
end;

function TclsTools.cSaveFileDialog(var filename: OleVariant; var InitFilename: WideString; 
                                   var Filter: WideString; var DefaultFileExtension: WideString; 
                                   var InitDir: WideString; var DialogTitle: WideString; 
                                   var Flags: Integer; var hwnd: Integer; var FilterIndex: Integer): Integer;
begin
  Result := DefaultInterface.cSaveFileDialog(filename, InitFilename, Filter, DefaultFileExtension, 
                                             InitDir, DialogTitle, Flags, hwnd, FilterIndex);
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TclsToolsProperties.Create(AServer: TclsTools);
begin
  inherited Create;
  FServer := AServer;
end;

function TclsToolsProperties.GetDefaultInterface: _clsTools;
begin
  Result := FServer.DefaultInterface;
end;

{$ENDIF}

class function CoclsUpdate.Create: _clsUpdate;
begin
  Result := CreateComObject(CLASS_clsUpdate) as _clsUpdate;
end;

class function CoclsUpdate.CreateRemote(const MachineName: string): _clsUpdate;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_clsUpdate) as _clsUpdate;
end;

procedure TclsUpdate.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{D4E7A610-1FDE-4F38-B498-60926031591C}';
    IntfIID:   '{6381BF10-F5C2-4577-B6CA-EDAF0287185D}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TclsUpdate.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as _clsUpdate;
  end;
end;

procedure TclsUpdate.ConnectTo(svrIntf: _clsUpdate);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TclsUpdate.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TclsUpdate.GetDefaultInterface: _clsUpdate;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface est NULL. Le composant n''est pas connecté au serveur. Vous devez appeler ''Connect'' ou ''ConnectTo'' avant cette opération');
  Result := FIntf;
end;

constructor TclsUpdate.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TclsUpdateProperties.Create(Self);
{$ENDIF}
end;

destructor TclsUpdate.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TclsUpdate.GetServerProperties: TclsUpdateProperties;
begin
  Result := FProps;
end;
{$ENDIF}

procedure TclsUpdate.CheckForUpdates(var ShowMessageNoNewUpdates: WordBool; 
                                     var ShowErrorMessage: WordBool; var TimeOutInMs: Integer);
begin
  DefaultInterface.CheckForUpdates(ShowMessageNoNewUpdates, ShowErrorMessage, TimeOutInMs);
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TclsUpdateProperties.Create(AServer: TclsUpdate);
begin
  inherited Create;
  FServer := AServer;
end;

function TclsUpdateProperties.GetDefaultInterface: _clsUpdate;
begin
  Result := FServer.DefaultInterface;
end;

{$ENDIF}

class function CoclsPDFCreator.Create: _clsPDFCreator;
begin
  Result := CreateComObject(CLASS_clsPDFCreator) as _clsPDFCreator;
end;

class function CoclsPDFCreator.CreateRemote(const MachineName: string): _clsPDFCreator;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_clsPDFCreator) as _clsPDFCreator;
end;

procedure TclsPDFCreator.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{40108C54-9352-46C9-822C-027727352F00}';
    IntfIID:   '{280807DB-86BB-4FD1-B477-A7A6E285A3FE}';
    EventIID:  '{4B1FCFC1-EB7C-4180-B0B6-68EBA12FCF88}';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TclsPDFCreator.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    ConnectEvents(punk);
    Fintf:= punk as _clsPDFCreator;
  end;
end;

procedure TclsPDFCreator.ConnectTo(svrIntf: _clsPDFCreator);
begin
  Disconnect;
  FIntf := svrIntf;
  ConnectEvents(FIntf);
end;

procedure TclsPDFCreator.DisConnect;
begin
  if Fintf <> nil then
  begin
    DisconnectEvents(FIntf);
    FIntf := nil;
  end;
end;

function TclsPDFCreator.GetDefaultInterface: _clsPDFCreator;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface est NULL. Le composant n''est pas connecté au serveur. Vous devez appeler ''Connect'' ou ''ConnectTo'' avant cette opération');
  Result := FIntf;
end;

constructor TclsPDFCreator.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TclsPDFCreatorProperties.Create(Self);
{$ENDIF}
end;

destructor TclsPDFCreator.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TclsPDFCreator.GetServerProperties: TclsPDFCreatorProperties;
begin
  Result := FProps;
end;
{$ENDIF}

procedure TclsPDFCreator.InvokeEvent(DispID: TDispID; var Params: TVariantArray);
begin
  case DispID of
    -1: Exit;  // DISPID_UNKNOWN
    1: if Assigned(FOneReady) then
         FOneReady(Self);
    2: if Assigned(FOneError) then
         FOneError(Self);
  end; {case DispID}
end;

function TclsPDFCreator.Get_cPrinterProfile(const Printername: WideString): WideString;
begin
    Result := DefaultInterface.cPrinterProfile[Printername];
end;

procedure TclsPDFCreator.Set_cPrinterProfile(const Printername: WideString; const Param2: WideString);
  { Avertissement : cette propriété cPrinterProfile a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.cPrinterProfile := Param2;
end;

function TclsPDFCreator.Get_cIsClosed: WordBool;
begin
    Result := DefaultInterface.cIsClosed;
end;

function TclsPDFCreator.Get_cError: _clsPDFCreatorError;
begin
    Result := DefaultInterface.cError;
end;

function TclsPDFCreator.Get_cErrorDetail(const PropertyName: WideString): OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.cErrorDetail[PropertyName];
end;

function TclsPDFCreator.Get_cGhostscriptVersion: WideString;
begin
    Result := DefaultInterface.cGhostscriptVersion;
end;

function TclsPDFCreator.Get_cOutputFilename: WideString;
begin
    Result := DefaultInterface.cOutputFilename;
end;

function TclsPDFCreator.Get_cPDFCreatorApplicationPath: WideString;
begin
    Result := DefaultInterface.cPDFCreatorApplicationPath;
end;

function TclsPDFCreator.Get_cIsLogfileDialogDisplayed: WordBool;
begin
    Result := DefaultInterface.cIsLogfileDialogDisplayed;
end;

function TclsPDFCreator.Get_cIsOptionsDialogDisplayed: WordBool;
begin
    Result := DefaultInterface.cIsOptionsDialogDisplayed;
end;

function TclsPDFCreator.Get_cProgramRelease(WithBeta: WordBool): WideString;
begin
    Result := DefaultInterface.cProgramRelease[WithBeta];
end;

function TclsPDFCreator.Get_cProgramIsRunning: WordBool;
begin
    Result := DefaultInterface.cProgramIsRunning;
end;

function TclsPDFCreator.Get_cWindowsVersion: WideString;
begin
    Result := DefaultInterface.cWindowsVersion;
end;

function TclsPDFCreator.Get_cVisible: WordBool;
begin
    Result := DefaultInterface.cVisible;
end;

procedure TclsPDFCreator.Set_cVisible(Param1: WordBool);
begin
  DefaultInterface.Set_cVisible(Param1);
end;

function TclsPDFCreator.Get_cInstalledAsServer: WordBool;
begin
    Result := DefaultInterface.cInstalledAsServer;
end;

function TclsPDFCreator.Get_cPrinterStop: WordBool;
begin
    Result := DefaultInterface.cPrinterStop;
end;

procedure TclsPDFCreator.Set_cPrinterStop(Param1: WordBool);
begin
  DefaultInterface.Set_cPrinterStop(Param1);
end;

function TclsPDFCreator.Get_cOptionsNames: _Collection;
begin
    Result := DefaultInterface.cOptionsNames;
end;

function TclsPDFCreator.Get_cOption(const PropertyName: WideString): OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.cOption[PropertyName];
end;

procedure TclsPDFCreator.Set_cOption(const PropertyName: WideString; Param2: OleVariant);
begin
  DefaultInterface.cOption[PropertyName] := Param2;
end;

function TclsPDFCreator.Get_cOptionProfile(const ProfileName: WideString; 
                                           const PropertyName: WideString): OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.cOptionProfile[ProfileName, PropertyName];
end;

procedure TclsPDFCreator.Set_cOptionProfile(const ProfileName: WideString; 
                                            const PropertyName: WideString; Param3: OleVariant);
begin
  DefaultInterface.cOptionProfile[ProfileName, PropertyName] := Param3;
end;

function TclsPDFCreator.Get_cOptions: _clsPDFCreatorOptions;
begin
    Result := DefaultInterface.cOptions;
end;

procedure TclsPDFCreator._Set_cOptions(const Param1: _clsPDFCreatorOptions);
  { Avertissement : cette propriété cOptions a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.cOptions := Param1;
end;

function TclsPDFCreator.Get_cOptionsProfile(const ProfileName: WideString): _clsPDFCreatorOptions;
begin
    Result := DefaultInterface.cOptionsProfile[ProfileName];
end;

procedure TclsPDFCreator._Set_cOptionsProfile(const ProfileName: WideString; 
                                              const Param2: _clsPDFCreatorOptions);
  { Avertissement : cette propriété cOptionsProfile a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.cOptionsProfile := Param2;
end;

function TclsPDFCreator.Get_cStandardOption(const PropertyName: WideString): OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.cStandardOption[PropertyName];
end;

function TclsPDFCreator.Get_cStandardOptions: _clsPDFCreatorOptions;
begin
    Result := DefaultInterface.cStandardOptions;
end;

function TclsPDFCreator.Get_cPostscriptInfo(const PostscriptFilename: WideString; 
                                            const PropertyName: WideString): WideString;
begin
    Result := DefaultInterface.cPostscriptInfo[PostscriptFilename, PropertyName];
end;

function TclsPDFCreator.Get_cPrintjobInfos(const PrintjobFilename: WideString): _clsPDFCreatorInfoSpoolFile;
begin
    Result := DefaultInterface.cPrintjobInfos[PrintjobFilename];
end;

function TclsPDFCreator.Get_cPrintjobInfo(const PrintjobFilename: WideString; 
                                          const PropertyName: WideString): WideString;
begin
    Result := DefaultInterface.cPrintjobInfo[PrintjobFilename, PropertyName];
end;

function TclsPDFCreator.Get_cCountOfPrintjobs: Integer;
begin
    Result := DefaultInterface.cCountOfPrintjobs;
end;

function TclsPDFCreator.Get_cPrintjobFilename(JobNumber: Integer): WideString;
begin
    Result := DefaultInterface.cPrintjobFilename[JobNumber];
end;

function TclsPDFCreator.Get_cDefaultPrinter: WideString;
begin
    Result := DefaultInterface.cDefaultPrinter;
end;

procedure TclsPDFCreator.Set_cDefaultPrinter(const Param1: WideString);
  { Avertissement : cette propriété cDefaultPrinter a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.cDefaultPrinter := Param1;
end;

function TclsPDFCreator.Get_cStopURLPrinting: WordBool;
begin
    Result := DefaultInterface.cStopURLPrinting;
end;

procedure TclsPDFCreator.Set_cStopURLPrinting(Param1: WordBool);
begin
  DefaultInterface.Set_cStopURLPrinting(Param1);
end;

function TclsPDFCreator.Get_cWindowState: Integer;
begin
    Result := DefaultInterface.cWindowState;
end;

procedure TclsPDFCreator.Set_cWindowState(Param1: Integer);
begin
  DefaultInterface.Set_cWindowState(Param1);
end;

function TclsPDFCreator.Get_cIsConverted: WordBool;
begin
    Result := DefaultInterface.cIsConverted;
end;

procedure TclsPDFCreator.Set_cIsConverted(Param1: WordBool);
begin
  DefaultInterface.Set_cIsConverted(Param1);
end;

function TclsPDFCreator.Get_cInstanceCounter: Integer;
begin
    Result := DefaultInterface.cInstanceCounter;
end;

procedure TclsPDFCreator.cErrorClear;
begin
  DefaultInterface.cErrorClear;
end;

function TclsPDFCreator.cIsAdministrator: WordBool;
begin
  Result := DefaultInterface.cIsAdministrator;
end;

function TclsPDFCreator.cPrinterIsInstalled(const Printername: WideString): WordBool;
begin
  Result := DefaultInterface.cPrinterIsInstalled(Printername);
end;

function TclsPDFCreator.cAddPDFCreatorPrinter(const Printername: WideString; 
                                              const ProfileName: WideString): WordBool;
begin
  Result := DefaultInterface.cAddPDFCreatorPrinter(Printername, ProfileName);
end;

function TclsPDFCreator.cProfileExists(const ProfileName: WideString): WordBool;
begin
  Result := DefaultInterface.cProfileExists(ProfileName);
end;

function TclsPDFCreator.cDeletePDFCreatorPrinter(const Printername: WideString): WordBool;
begin
  Result := DefaultInterface.cDeletePDFCreatorPrinter(Printername);
end;

function TclsPDFCreator.cGetProfileNames: _Collection;
begin
  Result := DefaultInterface.cGetProfileNames;
end;

function TclsPDFCreator.cAddProfile(const ProfileName: WideString; 
                                    const Options1: _clsPDFCreatorOptions): WordBool;
begin
  Result := DefaultInterface.cAddProfile(ProfileName, Options1);
end;

function TclsPDFCreator.cRenameProfile(const OldProfileName: WideString; 
                                       const NewProfileName: WideString): WordBool;
begin
  Result := DefaultInterface.cRenameProfile(OldProfileName, NewProfileName);
end;

function TclsPDFCreator.cDeleteProfile(const ProfileName: WideString): WordBool;
begin
  Result := DefaultInterface.cDeleteProfile(ProfileName);
end;

procedure TclsPDFCreator.cAddPrintjob(const filename: WideString);
begin
  DefaultInterface.cAddPrintjob(filename);
end;

procedure TclsPDFCreator.cDeletePrintjob(JobNumber: Integer);
begin
  DefaultInterface.cDeletePrintjob(JobNumber);
end;

procedure TclsPDFCreator.cMovePrintjobBottom(JobNumber: Integer);
begin
  DefaultInterface.cMovePrintjobBottom(JobNumber);
end;

procedure TclsPDFCreator.cMovePrintjobTop(JobNumber: Integer);
begin
  DefaultInterface.cMovePrintjobTop(JobNumber);
end;

procedure TclsPDFCreator.cMovePrintjobUp(JobNumber: Integer);
begin
  DefaultInterface.cMovePrintjobUp(JobNumber);
end;

procedure TclsPDFCreator.cMovePrintjobDown(JobNumber: Integer);
begin
  DefaultInterface.cMovePrintjobDown(JobNumber);
end;

function TclsPDFCreator.cClose: WordBool;
begin
  Result := DefaultInterface.cClose;
end;

function TclsPDFCreator.cStart(const Params: WideString; ForceInitialize: WordBool): WordBool;
begin
  Result := DefaultInterface.cStart(Params, ForceInitialize);
end;

procedure TclsPDFCreator.cClearCache;
begin
  DefaultInterface.cClearCache;
end;

procedure TclsPDFCreator.cClearLogfile;
begin
  DefaultInterface.cClearLogfile;
end;

procedure TclsPDFCreator.cConvertPostscriptfile(const InputFilename: WideString; 
                                                const OutputFilename: WideString);
begin
  DefaultInterface.cConvertPostscriptfile(InputFilename, OutputFilename);
end;

procedure TclsPDFCreator.cConvertFile(const InputFilename: WideString; 
                                      const OutputFilename: WideString; const SubFormat: WideString);
begin
  DefaultInterface.cConvertFile(InputFilename, OutputFilename, SubFormat);
end;

procedure TclsPDFCreator.cTestEvent(const EventName: WideString);
begin
  DefaultInterface.cTestEvent(EventName);
end;

procedure TclsPDFCreator.cShowLogfileDialog(value: WordBool);
begin
  DefaultInterface.cShowLogfileDialog(value);
end;

procedure TclsPDFCreator.cShowOptionsDialog(value: WordBool);
begin
  DefaultInterface.cShowOptionsDialog(value);
end;

procedure TclsPDFCreator.cSendMail(const OutputFilename: WideString; const Recipients: WideString);
begin
  DefaultInterface.cSendMail(OutputFilename, Recipients);
end;

function TclsPDFCreator.cIsPrintable(const filename: WideString): WordBool;
begin
  Result := DefaultInterface.cIsPrintable(filename);
end;

procedure TclsPDFCreator.cCombineAll;
begin
  DefaultInterface.cCombineAll;
end;

function TclsPDFCreator.cGetPDFCreatorPrinters: _Collection;
begin
  Result := DefaultInterface.cGetPDFCreatorPrinters;
end;

function TclsPDFCreator.cGetPrinterProfiles: _Collection;
begin
  Result := DefaultInterface.cGetPrinterProfiles;
end;

function TclsPDFCreator.cGetLogfile: WideString;
begin
  Result := DefaultInterface.cGetLogfile;
end;

procedure TclsPDFCreator.cWriteToLogfile(const LogStr: WideString);
begin
  DefaultInterface.cWriteToLogfile(LogStr);
end;

procedure TclsPDFCreator.cPrintFile(const filename: WideString);
begin
  DefaultInterface.cPrintFile(filename);
end;

procedure TclsPDFCreator.cPrintURL(const URL: WideString; TimeBetweenLoadAndPrint: Integer);
begin
  DefaultInterface.cPrintURL(URL, TimeBetweenLoadAndPrint);
end;

procedure TclsPDFCreator.cPrintPDFCreatorTestpage;
begin
  DefaultInterface.cPrintPDFCreatorTestpage;
end;

procedure TclsPDFCreator.cPrintPrinterTestpage(const Printername: WideString);
begin
  DefaultInterface.cPrintPrinterTestpage(Printername);
end;

function TclsPDFCreator.cReadOptions(const ProfileName: WideString): _clsPDFCreatorOptions;
begin
  Result := DefaultInterface.cReadOptions(ProfileName);
end;

procedure TclsPDFCreator.cSaveOptions(Options1: OleVariant; const ProfileName: WideString);
begin
  DefaultInterface.cSaveOptions(Options1, ProfileName);
end;

function TclsPDFCreator.cReadOptionsFromFile(const INIFilename: WideString): _clsPDFCreatorOptions;
begin
  Result := DefaultInterface.cReadOptionsFromFile(INIFilename);
end;

procedure TclsPDFCreator.cSaveOptionsToFile(const INIFilename: WideString);
begin
  DefaultInterface.cSaveOptionsToFile(INIFilename, EmptyParam);
end;

procedure TclsPDFCreator.cSaveOptionsToFile(const INIFilename: WideString; Options1: OleVariant);
begin
  DefaultInterface.cSaveOptionsToFile(INIFilename, Options1);
end;

function TclsPDFCreator.cGhostscriptRun(var Arguments: PSafeArray): WordBool;
begin
  Result := DefaultInterface.cGhostscriptRun(Arguments);
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TclsPDFCreatorProperties.Create(AServer: TclsPDFCreator);
begin
  inherited Create;
  FServer := AServer;
end;

function TclsPDFCreatorProperties.GetDefaultInterface: _clsPDFCreator;
begin
  Result := FServer.DefaultInterface;
end;

function TclsPDFCreatorProperties.Get_cPrinterProfile(const Printername: WideString): WideString;
begin
    Result := DefaultInterface.cPrinterProfile[Printername];
end;

procedure TclsPDFCreatorProperties.Set_cPrinterProfile(const Printername: WideString; 
                                                       const Param2: WideString);
  { Avertissement : cette propriété cPrinterProfile a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.cPrinterProfile := Param2;
end;

function TclsPDFCreatorProperties.Get_cIsClosed: WordBool;
begin
    Result := DefaultInterface.cIsClosed;
end;

function TclsPDFCreatorProperties.Get_cError: _clsPDFCreatorError;
begin
    Result := DefaultInterface.cError;
end;

function TclsPDFCreatorProperties.Get_cErrorDetail(const PropertyName: WideString): OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.cErrorDetail[PropertyName];
end;

function TclsPDFCreatorProperties.Get_cGhostscriptVersion: WideString;
begin
    Result := DefaultInterface.cGhostscriptVersion;
end;

function TclsPDFCreatorProperties.Get_cOutputFilename: WideString;
begin
    Result := DefaultInterface.cOutputFilename;
end;

function TclsPDFCreatorProperties.Get_cPDFCreatorApplicationPath: WideString;
begin
    Result := DefaultInterface.cPDFCreatorApplicationPath;
end;

function TclsPDFCreatorProperties.Get_cIsLogfileDialogDisplayed: WordBool;
begin
    Result := DefaultInterface.cIsLogfileDialogDisplayed;
end;

function TclsPDFCreatorProperties.Get_cIsOptionsDialogDisplayed: WordBool;
begin
    Result := DefaultInterface.cIsOptionsDialogDisplayed;
end;

function TclsPDFCreatorProperties.Get_cProgramRelease(WithBeta: WordBool): WideString;
begin
    Result := DefaultInterface.cProgramRelease[WithBeta];
end;

function TclsPDFCreatorProperties.Get_cProgramIsRunning: WordBool;
begin
    Result := DefaultInterface.cProgramIsRunning;
end;

function TclsPDFCreatorProperties.Get_cWindowsVersion: WideString;
begin
    Result := DefaultInterface.cWindowsVersion;
end;

function TclsPDFCreatorProperties.Get_cVisible: WordBool;
begin
    Result := DefaultInterface.cVisible;
end;

procedure TclsPDFCreatorProperties.Set_cVisible(Param1: WordBool);
begin
  DefaultInterface.Set_cVisible(Param1);
end;

function TclsPDFCreatorProperties.Get_cInstalledAsServer: WordBool;
begin
    Result := DefaultInterface.cInstalledAsServer;
end;

function TclsPDFCreatorProperties.Get_cPrinterStop: WordBool;
begin
    Result := DefaultInterface.cPrinterStop;
end;

procedure TclsPDFCreatorProperties.Set_cPrinterStop(Param1: WordBool);
begin
  DefaultInterface.Set_cPrinterStop(Param1);
end;

function TclsPDFCreatorProperties.Get_cOptionsNames: _Collection;
begin
    Result := DefaultInterface.cOptionsNames;
end;

function TclsPDFCreatorProperties.Get_cOption(const PropertyName: WideString): OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.cOption[PropertyName];
end;

procedure TclsPDFCreatorProperties.Set_cOption(const PropertyName: WideString; Param2: OleVariant);
begin
  DefaultInterface.cOption[PropertyName] := Param2;
end;

function TclsPDFCreatorProperties.Get_cOptionProfile(const ProfileName: WideString; 
                                                     const PropertyName: WideString): OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.cOptionProfile[ProfileName, PropertyName];
end;

procedure TclsPDFCreatorProperties.Set_cOptionProfile(const ProfileName: WideString; 
                                                      const PropertyName: WideString; 
                                                      Param3: OleVariant);
begin
  DefaultInterface.cOptionProfile[ProfileName, PropertyName] := Param3;
end;

function TclsPDFCreatorProperties.Get_cOptions: _clsPDFCreatorOptions;
begin
    Result := DefaultInterface.cOptions;
end;

procedure TclsPDFCreatorProperties._Set_cOptions(const Param1: _clsPDFCreatorOptions);
  { Avertissement : cette propriété cOptions a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.cOptions := Param1;
end;

function TclsPDFCreatorProperties.Get_cOptionsProfile(const ProfileName: WideString): _clsPDFCreatorOptions;
begin
    Result := DefaultInterface.cOptionsProfile[ProfileName];
end;

procedure TclsPDFCreatorProperties._Set_cOptionsProfile(const ProfileName: WideString; 
                                                        const Param2: _clsPDFCreatorOptions);
  { Avertissement : cette propriété cOptionsProfile a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.cOptionsProfile := Param2;
end;

function TclsPDFCreatorProperties.Get_cStandardOption(const PropertyName: WideString): OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.cStandardOption[PropertyName];
end;

function TclsPDFCreatorProperties.Get_cStandardOptions: _clsPDFCreatorOptions;
begin
    Result := DefaultInterface.cStandardOptions;
end;

function TclsPDFCreatorProperties.Get_cPostscriptInfo(const PostscriptFilename: WideString; 
                                                      const PropertyName: WideString): WideString;
begin
    Result := DefaultInterface.cPostscriptInfo[PostscriptFilename, PropertyName];
end;

function TclsPDFCreatorProperties.Get_cPrintjobInfos(const PrintjobFilename: WideString): _clsPDFCreatorInfoSpoolFile;
begin
    Result := DefaultInterface.cPrintjobInfos[PrintjobFilename];
end;

function TclsPDFCreatorProperties.Get_cPrintjobInfo(const PrintjobFilename: WideString; 
                                                    const PropertyName: WideString): WideString;
begin
    Result := DefaultInterface.cPrintjobInfo[PrintjobFilename, PropertyName];
end;

function TclsPDFCreatorProperties.Get_cCountOfPrintjobs: Integer;
begin
    Result := DefaultInterface.cCountOfPrintjobs;
end;

function TclsPDFCreatorProperties.Get_cPrintjobFilename(JobNumber: Integer): WideString;
begin
    Result := DefaultInterface.cPrintjobFilename[JobNumber];
end;

function TclsPDFCreatorProperties.Get_cDefaultPrinter: WideString;
begin
    Result := DefaultInterface.cDefaultPrinter;
end;

procedure TclsPDFCreatorProperties.Set_cDefaultPrinter(const Param1: WideString);
  { Avertissement : cette propriété cDefaultPrinter a un setter et un getter dont
    les types ne sont pas en concordance. Delphi n'a pas pu générer une propriété
    de cette sorte et utilise donc un Variant à la place. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.cDefaultPrinter := Param1;
end;

function TclsPDFCreatorProperties.Get_cStopURLPrinting: WordBool;
begin
    Result := DefaultInterface.cStopURLPrinting;
end;

procedure TclsPDFCreatorProperties.Set_cStopURLPrinting(Param1: WordBool);
begin
  DefaultInterface.Set_cStopURLPrinting(Param1);
end;

function TclsPDFCreatorProperties.Get_cWindowState: Integer;
begin
    Result := DefaultInterface.cWindowState;
end;

procedure TclsPDFCreatorProperties.Set_cWindowState(Param1: Integer);
begin
  DefaultInterface.Set_cWindowState(Param1);
end;

function TclsPDFCreatorProperties.Get_cIsConverted: WordBool;
begin
    Result := DefaultInterface.cIsConverted;
end;

procedure TclsPDFCreatorProperties.Set_cIsConverted(Param1: WordBool);
begin
  DefaultInterface.Set_cIsConverted(Param1);
end;

function TclsPDFCreatorProperties.Get_cInstanceCounter: Integer;
begin
    Result := DefaultInterface.cInstanceCounter;
end;

{$ENDIF}

procedure Register;
begin
  RegisterComponents(dtlServerPage, [TclsPDFCreatorOptions, TclsPDFCreatorError, TclsPDFCreatorInfoSpoolFile, TclsTools, 
    TclsUpdate, TclsPDFCreator]);
end;

end.

unit MSDATASRC_TLB;

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

// PASTLWTR : $Revision:   1.88.1.0.1.0  $
// Fichier généré le 24/03/2006 17:21:46 depuis la bibliothèque de types ci-dessous.

// ************************************************************************ //
// Bibl.Types     : C:\WINDOWS\System32\msdatsrc.tlb (1)
// IID\LCID       : {7C0FFAB0-CD84-11D0-949A-00A0C91110ED}\0
// Fichier d'aide : 
// DepndLst       : 
//   (1) v2.0 stdole, (C:\WINDOWS\System32\stdole2.tlb)
// Bibliothèque de types parent :
//   (0) v1.0 OWC11, (C:\Program Files\Fichiers communs\Microsoft Shared\Web Components\11\OWC11.DLL)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // L'unité doit être compilée sans vérification de type des pointeurs. 
interface

uses Windows, ActiveX, Classes, Graphics, OleServer, OleCtrls, StdVCL;

// *********************************************************************//
// GUIDS déclarés dans la bibliothèque de types. Préfixes utilisés :    
//   Bibliothèques de types : LIBID_xxxx                                      
//   CoClasses              : CLASS_xxxx                                      
//   DISPInterfaces         : DIID_xxxx                                       
//   Non-DISP interfaces    : IID_xxxx                                        
// *********************************************************************//
const
  // Version majeure et mineure de la bibliothèque de types
  MSDATASRCMajorVersion = 1;
  MSDATASRCMinorVersion = 0;

  LIBID_MSDATASRC: TGUID = '{7C0FFAB0-CD84-11D0-949A-00A0C91110ED}';

  IID_DataSourceListener: TGUID = '{7C0FFAB2-CD84-11D0-949A-00A0C91110ED}';
  IID_DataSource: TGUID = '{7C0FFAB3-CD84-11D0-949A-00A0C91110ED}';
type

// *********************************************************************//
// Déclaration Forward des types définis dans la bibliothèque de types    
// *********************************************************************//
  DataSourceListener = interface;
  DataSource = interface;

// *********************************************************************//
// Déclaration de structures, d'unions et alias.                         
// *********************************************************************//
  PUserType1 = ^TGUID; {*}

  DataMember = WideString; 

// *********************************************************************//
// Interface   : DataSourceListener
// Indicateurs : (272) Hidden OleAutomation
// GUID        : {7C0FFAB2-CD84-11D0-949A-00A0C91110ED}
// *********************************************************************//
  DataSourceListener = interface(IUnknown)
    ['{7C0FFAB2-CD84-11D0-949A-00A0C91110ED}']
    function  dataMemberChanged(const bstrDM: DataMember): HResult; stdcall;
    function  dataMemberAdded(const bstrDM: DataMember): HResult; stdcall;
    function  dataMemberRemoved(const bstrDM: DataMember): HResult; stdcall;
  end;

// *********************************************************************//
// Interface   : DataSource
// Indicateurs : (256) OleAutomation
// GUID        : {7C0FFAB3-CD84-11D0-949A-00A0C91110ED}
// *********************************************************************//
  DataSource = interface(IUnknown)
    ['{7C0FFAB3-CD84-11D0-949A-00A0C91110ED}']
    function  getDataMember(const bstrDM: DataMember; var riid: TGUID; out ppunk: IUnknown): HResult; stdcall;
    function  getDataMemberName(lIndex: Integer; out pbstrDM: DataMember): HResult; stdcall;
    function  getDataMemberCount(out plCount: Integer): HResult; stdcall;
    function  addDataSourceListener(const pDSL: DataSourceListener): HResult; stdcall;
    function  removeDataSourceListener(const pDSL: DataSourceListener): HResult; stdcall;
  end;

implementation

uses ComObj;

end.

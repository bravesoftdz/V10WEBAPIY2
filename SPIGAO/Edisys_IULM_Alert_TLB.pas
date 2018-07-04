unit Edisys_IULM_Alert_TLB;

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
// Bibl. types : C:\Program Files\Edisys\IULM\x86\Edisys.IULM.Alert.tlb (1)
// LIBID: {F80E576F-2CF6-41AE-AEB5-6DD998D1DBF6}
// LCID: 0
// Fichier d'aide : 
// Chaîne d'aide : 
// DepndLst: 
//   (1) v2.0 stdole, (C:\Windows\system32\stdole2.tlb)
//   (2) v2.0 mscorlib, (C:\Windows\Microsoft.NET\Framework\v2.0.50727\mscorlib.tlb)
//   (3) v3.2 Edisys_IULM_Core, (C:\Program Files\Edisys\IULM\x86\Edisys.IULM.Core.tlb)
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

uses Windows, ActiveX, Classes, Edisys_IULM_Core_TLB, Graphics, mscorlib_TLB, OleServer, StdVCL, 
Variants;
  

// *********************************************************************//
// GUIDS déclarés dans la bibliothèque de types. Préfixes utilisés :    
//   Bibliothèques de types : LIBID_xxxx                                      
//   CoClasses              : CLASS_xxxx                                      
//   DISPInterfaces         : DIID_xxxx                                       
//   Non-DISP interfaces    : IID_xxxx                                        
// *********************************************************************//
const
  // Versions majeure et mineure de la bibliothèque de types
  Edisys_IULM_AlertMajorVersion = 3;
  Edisys_IULM_AlertMinorVersion = 2;

  LIBID_Edisys_IULM_Alert: TGUID = '{F80E576F-2CF6-41AE-AEB5-6DD998D1DBF6}';

  CLASS_ResultEventArgs: TGUID = '{CAF58E65-2F4F-4EA2-B19D-9861606D002B}';
  IID__ResultHandler: TGUID = '{7BE49544-9221-36AB-B7C7-E11F2AABC728}';
  CLASS_DetectionAlerts: TGUID = '{755155DC-AB81-4DAC-AD64-EAA36EE27501}';
  CLASS_ResultHandler: TGUID = '{8506CD28-4910-4E51-B6CF-7B30C3B4D1F3}';
type

// *********************************************************************//
// Déclaration Forward des types définis dans la bibliothèque de types    
// *********************************************************************//
  _ResultHandler = interface;
  _ResultHandlerDisp = dispinterface;

// *********************************************************************//
// Déclaration de CoClasses définies dans la bibliothèque de types 
// (REMARQUE: On affecte chaque CoClasse à son Interface par défaut)              
// *********************************************************************//
  ResultEventArgs = IResultEventArgs;
  DetectionAlerts = IDetectionAlerts;
  ResultHandler = _ResultHandler;


// *********************************************************************//
// Interface   : _ResultHandler
// Indicateurs : (4432) Hidden Dual OleAutomation Dispatchable
// GUID        : {7BE49544-9221-36AB-B7C7-E11F2AABC728}
// *********************************************************************//
  _ResultHandler = interface(IDispatch)            
    ['{7BE49544-9221-36AB-B7C7-E11F2AABC728}']
  end;

// *********************************************************************//
// DispIntf :  _ResultHandlerDisp
// Flags :     (4432) Hidden Dual OleAutomation Dispatchable
// GUID :      {7BE49544-9221-36AB-B7C7-E11F2AABC728}
// *********************************************************************//
  _ResultHandlerDisp = dispinterface
    ['{7BE49544-9221-36AB-B7C7-E11F2AABC728}']
  end;

// *********************************************************************//
// La classe CoResultEventArgs fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut IResultEventArgs exposée             
// par la CoClasse ResultEventArgs. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoResultEventArgs = class
    class function Create: IResultEventArgs;
    class function CreateRemote(const MachineName: string): IResultEventArgs;
  end;

// *********************************************************************//
// La classe CoDetectionAlerts fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut IDetectionAlerts exposée             
// par la CoClasse DetectionAlerts. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoDetectionAlerts = class
    class function Create: IDetectionAlerts;
    class function CreateRemote(const MachineName: string): IDetectionAlerts;
  end;

// *********************************************************************//
// La classe CoResultHandler fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut _ResultHandler exposée             
// par la CoClasse ResultHandler. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClasse exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoResultHandler = class
    class function Create: _ResultHandler;
    class function CreateRemote(const MachineName: string): _ResultHandler;
  end;

implementation

uses ComObj;

class function CoResultEventArgs.Create: IResultEventArgs;
begin
  Result := CreateComObject(CLASS_ResultEventArgs) as IResultEventArgs;
end;

class function CoResultEventArgs.CreateRemote(const MachineName: string): IResultEventArgs;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ResultEventArgs) as IResultEventArgs;
end;

class function CoDetectionAlerts.Create: IDetectionAlerts;
begin
  Result := CreateComObject(CLASS_DetectionAlerts) as IDetectionAlerts;
end;

class function CoDetectionAlerts.CreateRemote(const MachineName: string): IDetectionAlerts;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_DetectionAlerts) as IDetectionAlerts;
end;

class function CoResultHandler.Create: _ResultHandler;
begin
  Result := CreateComObject(CLASS_ResultHandler) as _ResultHandler;
end;

class function CoResultHandler.CreateRemote(const MachineName: string): _ResultHandler;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ResultHandler) as _ResultHandler;
end;

end.

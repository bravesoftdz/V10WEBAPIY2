unit CRCS5_TLB;

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
// Fichier généré le 06/08/2001 11:40:29 depuis la bibliothèque de types ci-dessous.

// ************************************************************************ //
// Bibl.Types     : Y:\PgiS5\VDEV\Prospect\CRCS5.tlb (1)
// IID\LCID       : {53F352D6-702F-4F81-8DDD-782F69C61176}\0
// Fichier d'aide :
// DepndLst       : 
//   (1) v2.0 stdole, (C:\WINDOWS\System32\stdole2.tlb)
//   (2) v4.0 StdVCL, (C:\WINDOWS\System32\STDVCL40.DLL)
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
  CRCS5MajorVersion = 1;
  CRCS5MinorVersion = 0;

  LIBID_CRCS5: TGUID = '{53F352D6-702F-4F81-8DDD-782F69C61176}';

  IID_IGRCWINDOW: TGUID = '{6F339FD2-EF7F-47CD-8801-281A43B91FB4}';
  CLASS_GRCWINDOW: TGUID = '{B4519199-7216-4B34-9DF9-B8E04D68059D}';
type

// *********************************************************************//
// Déclaration Forward des types définis dans la bibliothèque de types    
// *********************************************************************//
  IGRCWINDOW = interface;
  IGRCWINDOWDisp = dispinterface;

// *********************************************************************//
// Déclaration de CoClasses définies dans la bibliothèque de types 
// (REMARQUE: On affecte chaque CoClass à son Interface par défaut)              
// *********************************************************************//
  GRCWINDOW = IGRCWINDOW;


// *********************************************************************//
// Interface   : IGRCWINDOW
// Indicateurs : (4416) Dual OleAutomation Dispatchable
// GUID        : {6F339FD2-EF7F-47CD-8801-281A43B91FB4}
// *********************************************************************//
  IGRCWINDOW = interface(IDispatch)
    ['{6F339FD2-EF7F-47CD-8801-281A43B91FB4}']
    function  GRCCumulPersp(const etat: WideString; datcredeb: TDateTime; datcrefin: TDateTime; 
                            datfinviedeb: TDateTime; datfinviefin: TDateTime): OleVariant; safecall;
    function  GRCRechDOMOLE(const TT: WideString; const Code: WideString): OleVariant; safecall;
    procedure GRCInitDico; safecall;
    procedure GRCNewDico; safecall;
    procedure GRCShowDico; safecall;
    procedure GRCFreeDico; safecall;
  end;

// *********************************************************************//
// DispIntf:  IGRCWINDOWDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {6F339FD2-EF7F-47CD-8801-281A43B91FB4}
// *********************************************************************//
  IGRCWINDOWDisp = dispinterface
    ['{6F339FD2-EF7F-47CD-8801-281A43B91FB4}']
    function  GRCCumulPersp(const etat: WideString; datcredeb: TDateTime; datcrefin: TDateTime; 
                            datfinviedeb: TDateTime; datfinviefin: TDateTime): OleVariant; dispid 1;
    function  GRCRechDOMOLE(const TT: WideString; const Code: WideString): OleVariant; dispid 2;
    procedure GRCInitDico; dispid 3;
    procedure GRCNewDico; dispid 4;
    procedure GRCShowDico; dispid 5;
    procedure GRCFreeDico; dispid 6;
  end;

// *********************************************************************//
// La classe CoGRCWINDOW fournit une méthode Create et CreateRemote pour          
// créer des instances de l'interface par défaut IGRCWINDOW exposée             
// pas la CoClass GRCWINDOW. Les fonctions sont destinées à être utilisées par            
// les clients désirant automatiser les objets CoClass exposés par       
// le serveur de cette bibliothèque de types.                                            
// *********************************************************************//
  CoGRCWINDOW = class
    class function Create: IGRCWINDOW;
    class function CreateRemote(const MachineName: string): IGRCWINDOW;
  end;

implementation

uses ComObj;

class function CoGRCWINDOW.Create: IGRCWINDOW;
begin
  Result := CreateComObject(CLASS_GRCWINDOW) as IGRCWINDOW;
end;

class function CoGRCWINDOW.CreateRemote(const MachineName: string): IGRCWINDOW;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_GRCWINDOW) as IGRCWINDOW;
end;

end.

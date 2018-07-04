{***********UNITE*************************************************
Auteur  ...... : Paul Chapuis
Créé le ...... : 15/09/2003
Modifié le ... :   /  /    
Description .. : Classe OLE générique pour PGI
Suite ........ : 
Suite ........ : enregistrement de la classe par appel à 
Suite ........ : RegisterPGIOLEAUTO
Suite ........ : 
Suite ........ : Peut fonctionner conjointement à d'autres classes métier
Mots clefs ... : 
*****************************************************************}
unit PGIOLEAUTO;

interface

uses ComObj, ActiveX, HEnt1, HCtrls,
{$IFDEF EAGLCLIENT}
     UTob, eFichList,
{$ELSE}
     DB, {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}FichList,
{$ENDIF}
     StdVcl ,OLEAuto;
type
  TPGIOLEAUTO = class(TAutoObject)
  automated
    function  PO_GetUser:variant ;
    function  PO_GetUserName:variant ;
    function  PO_GetUserLogin:variant ;
 protected
     procedure OnTheLast (var shutdown : boolean) ;
 public
     constructor Create; override;
  end;

 procedure RegisterPGIOLEAUTO ;

implementation

uses {ComServ,} Dialogs, Forms, UTom,  UtomLiensOle,sysutils;

procedure RegisterPGIOLEAUTO ; // appelé dans la phase initialization
const
  AutoClassInfo: TAutoClassInfo = (
    AutoClass: TPGIOLEAUTO;
    ProgID: 'PGIOLEAUTO';
// ATTENTION : CE NUMERO DE CLASSE ID EST A CHANGER
// IL DOIT LIE A UNE APPLICATION ET A UNE SEULE
// POUR CALCULER UN GUID UNIQUE, FAIRE DANS DELPHI :
// CTRL+SHIFT+G

    ClassID: '{8D9A00A6-EA29-4CA8-B693-25282EEFAB1A}';
    Description: 'Cegid PGI OLE';
    Instancing: acMultiInstance);
begin
  Automation.RegisterClass(AutoClassInfo);
end;

procedure TPGIOLEAUTO.OnTheLast(var shutdown: boolean);
begin
  //ShutDown:=FALSE ;  // pour ne pas fermer l'application automatiquement lorsqu'il n'y a plus de réference à l'objet.
  ShutDown:=(FindCmdLineSwitch('REGSERVER') OR FindCmdLineSwitch('UNREGSERVER'));
end;

constructor TPGIOLEAUTO.create;
begin
inherited;
Automation.OnLastRelease:=OnTheLast;
end;

function  TPGIOLEAUTO.PO_GetUser:variant ;
begin
Result:=V_PGI.User;
end;

function  TPGIOLEAUTO.PO_GetUserName:variant ;
begin
Result:=V_PGI.UserName;
end;

function  TPGIOLEAUTO.PO_GetUserLogin:variant ;
begin
Result:=V_PGI.UserLogin;
end;


end.

{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/06/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : DEPRECIATION ()
Mots clefs ... : TOF;DEPRECIATION
*****************************************************************}
Unit CPAVERTNEWRAPPRO_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Sysutils,
     Graphics,
     Classes,
     Windows,
     Grids,
{$IFDEF EAGLCLIENT}
     eMul,
     MaineAGL,
{$else}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,
{$ENDIF}
     forms,
     HSysMenu,
     LookUp,
     ComCtrls,
     HCtrls,
     HEnt1,
     Vierge,
     HMsgBox,
     HTB97,
     utob,
     UTOF ;


Function INITNEWPOINTAGE : Boolean ;
Function OkNewPointage : Boolean ;

Type
  TOF_CPAVERTNEWPOINTAGE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  Private
    procedure MajHelp ;
    procedure BValiderClick (Sender : TObject) ;
  end ;

Implementation

Uses HStatus, PARAMSOC ;

Function OkNewPointage : Boolean ;
BEGIN
Result:=GetparamSocSecur('SO_CP_NEWRAPPRO',FALSE) ;
END ;

Function INITNEWPOINTAGE : Boolean ;
Var St : String ;
    OldZoomOle : Boolean ;
BEGIN
Result:=FALSE ;
If OkNewPointage Then
  BEGIN
  PgiError('Opération impossible : ce traitement a déjà été effectué.','') ;
  Exit ;
  END ;
OldZoomOle:=V_PGI.ZoomOle ; V_PGI.ZoomOle:=TRUE ;
St:=AGLLanceFiche('CP','CPAVERTNEWPOINTAG','','','') ;
Result:=St='X' ;
V_PGI.ZoomOle:=OldZoomOle ;
END ;

procedure TOF_CPAVERTNEWPOINTAGE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPAVERTNEWPOINTAGE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPAVERTNEWPOINTAGE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPAVERTNEWPOINTAGE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPAVERTNEWPOINTAGE.MajHelp ;
BEGIN
TFVierge(Ecran).HelpContext:=0 ;
end ;

procedure TOF_CPAVERTNEWPOINTAGE.OnArgument (S : String ) ;
Var BT : TToolBarButton97 ;
    F : TfVierge ;
begin
  Inherited ;
F := TfVierge(Ecran) ;
F.Retour:='' ;
F.ModalResult:=mrNone ;
MajHelp ;
BT:=TToolbarButton97(GetControl('BVALIDER')) ; If BT<>NIL Then BT.OnClick:=BValiderClick ;
END ;

procedure TOF_CPAVERTNEWPOINTAGE.BValiderClick (Sender : TObject) ;
var F : TfVierge ;
begin
F := TfVierge(Ecran) ;
F.ModalResult:=mrNone ;
F.Retour:='' ;
if GetControlText('FOUI')<>'OUI' then BEGIN PgiError('Cette opération est impossible : le mot de passe est incorrect') ; Exit ; END ;
If PGIAsk('Confirmez-vous le traitement ?','')<>mrYes Then Exit ;

ExecuteSQL( 'UPDATE EEXBQLIG SET CEL_DEVISE = (SELECT MAX(EE_DEVISE) FROM EEXBQ ' +
                          'WHERE CEL_REFPOINTAGE = EE_REFPOINTAGE ' +
                          'AND CEL_NUMRELEVE = EE_NUMERO) ' +
                          'WHERE (CEL_DEVISE = "" OR CEL_DEVISE IS NULL)' );
ExecuteSQL( 'UPDATE EEXBQLIG SET CEL_REFPOINTAGE = "" WHERE (CEL_DATEPOINTAGE = "' + UsDateTime(iDate1900) + '" OR CEL_DATEPOINTAGE IS NULL)' );

PgiInfo('Traitement terminé')  ;
SetParamSoc('SO_CP_NEWRAPPRO','X') ;
F.ModalResult:=mrOk ;
F.Retour:='X';
end ;

procedure TOF_CPAVERTNEWPOINTAGE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPAVERTNEWPOINTAGE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPAVERTNEWPOINTAGE.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CPAVERTNEWPOINTAGE ] ) ;
end.


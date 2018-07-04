{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 07/09/2016
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSAISIEAPP ()
Mots clefs ... : TOF;BTSAISIEAPP
*****************************************************************}
Unit BTSAISIEAPP_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
{$ENDIF}
     uTob, 
     forms,
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1,
     HTB97,
     HMsgBox,
     HRichOLE,
     HRichEdt,
     UTOF ;

Type
  TOF_BTSAISIEAPP = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TOBAPP : TOB;
  end ;

Implementation
uses AffaireUtil,AppelsUtil;

procedure TOF_BTSAISIEAPP.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEAPP.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEAPP.OnUpdate ;
begin
  Inherited ;
  if not ControleSaisieOK (THEdit(getControl('AFF_APPEL1')).text, THEdit(getControl('AFF_APPEL2')).text, THEdit(getControl('AFF_APPEL3')).text) then
  begin
    PgiInfo('Merci de renseigner le code intervention');
    ecran.ModalResult := mrNone;
    exit;
  end;
  TOBAPP.setString('SAISIEOK','X')
end ;

procedure TOF_BTSAISIEAPP.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEAPP.OnArgument (S : String ) ;
var TheAffaire : string;
begin
  Inherited ;
  TOBAPP := LaTOB;
  SetControlText('AFF_APPEL0','A');
  ChargeCleAffaire(THEdit(getControl('AFF_APPEL0')), THEdit(getControl('AFF_APPEL1')), THEdit(getControl('AFF_APPEL2')), THEdit(getControl('AFF_APPEL3')), THEdit(getControl('APP_AVENANT')), TToolBarButton97(getControl('BEFFACEAPP')), TaCreat, TheAffaire, True);
  SetControlText('AFF_APPEL0','W');
  TOBAPP.PutEcran(ecran);
  StringToRich(THricheditOle(GetCOntrol('MEMO')), TOBAPP.GetValue('BAA_DESCRIPTIF'));
end ;

procedure TOF_BTSAISIEAPP.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEAPP.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEAPP.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTSAISIEAPP ] ) ; 
end.

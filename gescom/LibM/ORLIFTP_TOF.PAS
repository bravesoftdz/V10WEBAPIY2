{***********UNITE*************************************************
Auteur  ...... : Julien D
Créé le ...... : 17/06/2002
Modifié le ... : 17/06/2002
Description .. : Source TOF de la FICHE : ORLIFTP ()
Mots clefs ... : TOF;ORLIFTP
*****************************************************************}
Unit ORLIFTP_TOF ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,dbtables,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,
     ParamSoc,Vierge;

Type
  TOF_ORLIFTP = Class (TOF)
    Private
      SauveParamSoc : Boolean;
    Public
      procedure OnUpdate                 ; override ;
      procedure OnLoad                   ; override ;
      procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

procedure TOF_ORLIFTP.OnUpdate ;
begin
  Inherited ;
  if SauveParamSoc then
    begin
    SetParamSoc('SO_GCORLIHOST',GetControlText('GCORLIHOST'));
    SetParamSoc('SO_GCORLIUSERNAME',GetControlText('GCORLIUSERNAME'));
    SetParamSoc('SO_GCORLIPASSWORD',GetControlText('GCORLIPASSWORD'));
    SetParamSoc('SO_GCORLIREPFTP',GetControlText('GCORLIREPFTP'));
    end;
  TFVierge(ecran).Retour:='SO_GCORLIHOST='+GetControlText('GCORLIHOST')+
                          ';SO_GCORLIUSERNAME='+GetControlText('GCORLIUSERNAME')+
                          ';SO_GCORLIPASSWORD='+GetControlText('GCORLIPASSWORD')+
                          ';SO_GCORLIREPFTP='+GetControlText('GCORLIREPFTP');
end ;

procedure TOF_ORLIFTP.OnLoad ;
begin
  Inherited ;
  SetControlText('GCORLIHOST',GetParamSoc('SO_GCORLIHOST'));
  SetControlText('GCORLIUSERNAME',GetParamSoc('SO_GCORLIUSERNAME'));
  SetControlText('GCORLIPASSWORD',GetParamSoc('SO_GCORLIPASSWORD'));
  SetControlText('GCORLIREPFTP',GetParamSoc('SO_GCORLIREPFTP'));
end ;

procedure TOF_ORLIFTP.OnArgument (S : String ) ;
begin
  Inherited ;
  SauveParamSoc := S='SAUVEPARAMSOC';
end ;

Initialization
  registerclasses ( [ TOF_ORLIFTP ] ) ;
end.

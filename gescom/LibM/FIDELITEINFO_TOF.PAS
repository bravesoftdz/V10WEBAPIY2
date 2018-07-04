{***********UNITE*************************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 05/06/2003
Modifié le ... : 10/07/2003
Description .. : Source TOF de la FICHE : FIDELITEINFO ()
Mots clefs ... : TOF;FIDELITEINFO
*****************************************************************}
Unit FIDELITEINFO_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     dbtables, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,
     Vierge;

Type
  TOF_FIDELITEINFO = Class (TOF)
      procedure OnUpdate                 ; override ;
      procedure OnLoad                   ; override ;
      procedure OnArgument (S : String ) ; override ;
      procedure OnClose                  ; override ;
  end ;

Implementation


procedure TOF_FIDELITEINFO.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_FIDELITEINFO.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_FIDELITEINFO.OnArgument (S : String ) ;
var Critere : String;
    x : Integer;
begin
  Inherited ;
  Repeat
    Critere := ReadTokenSt(S) ;
    if Critere<>'' then
    begin
      x := pos('=',Critere);
      if x<>0 then
        SetControlText( 'T'+copy(Critere,1,x-1), copy(Critere,x+1,length(Critere)) );
    end;
  until Critere='';
  SetControlText( 'VALRETOUR', '0.0' );
end ;

procedure TOF_FIDELITEINFO.OnClose ;
begin
  Inherited ;
  TFVierge(Ecran).Retour := GetControlText('ISMODIFCUMUL');
end ;

Initialization
  registerclasses ( [ TOF_FIDELITEINFO ] ) ; 
end.

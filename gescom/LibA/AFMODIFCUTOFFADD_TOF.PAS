{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 07/08/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFMODIFCUTOFFADD ()
Mots clefs ... : TOF;AFMODIFCUTOFFADD
*****************************************************************}
Unit UtofAFMODIFCUTOFFADD ;

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
     UTOF ; 

Type
  TOF_AFMODIFCUTOFFADD = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

Implementation

procedure TOF_AFMODIFCUTOFFADD.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFMODIFCUTOFFADD.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFMODIFCUTOFFADD.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFMODIFCUTOFFADD.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFMODIFCUTOFFADD.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_AFMODIFCUTOFFADD.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_AFMODIFCUTOFFADD ] ) ; 
end.

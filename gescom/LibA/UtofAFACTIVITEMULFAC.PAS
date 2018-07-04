{***********UNITE*************************************************
Auteur  ...... : PL
Créé le ...... : 14/11/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFACTIVITEMULFAC ()
Mots clefs ... : TOF;AFACTIVITEMULFAC
*****************************************************************}
Unit UtofAFACTIVITEMULFAC ;

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
  TOF_AFACTIVITEMULFAC = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

Implementation

procedure TOF_AFACTIVITEMULFAC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFACTIVITEMULFAC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFACTIVITEMULFAC.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFACTIVITEMULFAC.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFACTIVITEMULFAC.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_AFACTIVITEMULFAC.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_AFACTIVITEMULFAC ] ) ; 
end.

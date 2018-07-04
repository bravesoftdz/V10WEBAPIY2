{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMUL_FORMCURSUS ()
Mots clefs ... : TOF;PGMUL_FORMCURSUS
*****************************************************************}
Unit UTofPGMul_FormCursus ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF ;

Type
  TOF_PGMUL_FORMCURSUS = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

Implementation

procedure TOF_PGMUL_FORMCURSUS.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGMUL_FORMCURSUS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGMUL_FORMCURSUS.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_PGMUL_FORMCURSUS.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_PGMUL_FORMCURSUS ] ) ; 
end.


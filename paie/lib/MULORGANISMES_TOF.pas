{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 07/07/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MULORGANISMES ()
Mots clefs ... : TOF;MULORGANISMES
*****************************************************************}
Unit MULORGANISMES_TOF ;

Interface

Uses
//unused     StdCtrls, 
//unused     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
//unused     db,
//unused     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
//unused     FE_Main,
{$ELSE}
//unused     MaineAGL,
//unused     UTOB,
//unused     eMul,
{$ENDIF}
//unused     AGLInit,
//unused     forms,
//unused     sysutils,
//unused     ComCtrls,
//unused     HCtrls,
//unused     HEnt1,
//unused     HMsgBox,
     UTOF ;

Type
  TOF_MULORGANISMES = Class (TOF)
    private

    public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

procedure TOF_MULORGANISMES.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_MULORGANISMES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_MULORGANISMES.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_MULORGANISMES.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_MULORGANISMES.OnArgument (S : String ) ;

begin
  Inherited ;
end ;

procedure TOF_MULORGANISMES.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_MULORGANISMES.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_MULORGANISMES.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_MULORGANISMES ] ) ; 
end.

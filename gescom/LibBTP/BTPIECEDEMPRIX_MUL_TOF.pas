{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 23/01/2012
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BPIECEDEMPRIX_MUL ()
Mots clefs ... : TOF;BPIECEDEMPRIX_MUL
*****************************************************************}
Unit BPIECEDEMPRIX_MUL_TOF;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,

{$IFDEF EAGLCLIENT}
     MaineAGL,
     eMul,
     uTob,
{$ELSE}
     DBCtrls, Db,
     {$IFNDEF DBXPRESS}
     dbTables,
     {$ELSE}
     uDbxDataSet,
     {$ENDIF DBXPRESS}
     fe_main,
     HDB,
     Mul,
{$ENDIF EAGLCLIENT}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     UTOF ;

Type
  TOF_BPIECEDEMPRIX_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  Private

    BtCreate : TToolbarButton97;
  
    procedure OnCreate(Sender: Tobject);

  end ;

Implementation
uses uentcommun;

procedure TOF_BPIECEDEMPRIX_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BPIECEDEMPRIX_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BPIECEDEMPRIX_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BPIECEDEMPRIX_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BPIECEDEMPRIX_MUL.OnArgument (S : String ) ;
begin
  Inherited ;

  if assigned(Getcontrol('BTInsert')) then
  begin
     BtCreate := TToolbarButton97(ecran.FindComponent('BTInsert'));
     BtCreate.OnClick := OnCreate;
  end;
end ;

procedure TOF_BPIECEDEMPRIX_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BPIECEDEMPRIX_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BPIECEDEMPRIX_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BPIECEDEMPRIX_MUL.OnCreate(Sender: Tobject);
var cledoc : R_Cledoc;
begin
  AGLLanceFiche('BTP','BTCREATEDEMPRIX','','','') ;   // Demande de prix
end;

Initialization
  registerclasses ( [ TOF_BPIECEDEMPRIX_MUL ] ) ;
end.


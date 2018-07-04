{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 16/03/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTVALIDCTRETUDE ()
Mots clefs ... : TOF;BTVALIDCTRETUDE
*****************************************************************}
Unit BTVALIDCTRETUDE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$ELSE}
     MainEagl,
{$ENDIF}
     forms,
     menus,
     AglInit,
     saisUtil,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     UTOB,
     UTOF,
     UPlannifchUtil,
     FactComm,
     FactUtil,
     FactCalc,
     FactGrp,
     FactGrpBTP,
     paramSoc,
     EntGC,
     FactOuvrage,
     FactVariante,
     ExtCtrls,
     Graphics,
     FactTob,
     FactAdresse,
		 UtilPlannifchantier
     ;

Type
  TOF_BTVALIDCTRETUDE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

procedure PlannificationCtrEtude (TOBDEVIS,TobOuvrage,TOBFrais,TOBOuvrageFrais : TOB; Mode : integer);

Implementation

procedure PlannificationCtrEtude (TOBDEVIS,TobOuvrage,TOBFrais,TOBOuvrageFrais : TOB; Mode : integer);
var Param : string;
begin
  Param := 'MODE='+ inttostr(Mode);
  TheTob := TOBDevis;
  TheTob.Data := TOBOuvrage;
  TOBOuvrage.data := TOBFrais;
  TOBFrais.data := TOBOuvrageFrais;
  AGLLanceFiche('BTP','BTVALIDCTRETUDE','','',Param) ;
  TheTob := nil;
end;

procedure TOF_BTVALIDCTRETUDE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTVALIDCTRETUDE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTVALIDCTRETUDE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTVALIDCTRETUDE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTVALIDCTRETUDE.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_BTVALIDCTRETUDE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTVALIDCTRETUDE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTVALIDCTRETUDE.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTVALIDCTRETUDE ] ) ; 
end.


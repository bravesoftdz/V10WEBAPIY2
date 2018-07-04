{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFREVISION_MUL ()
Mots clefs ... : TOF;AFREVISION_MUL
*****************************************************************}
Unit uTofAFRevisionMul;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,mul,
{$Else}
     MainEagl,emul,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     UTOF,HTB97,utofafrevision ;

Type
  TOF_AFREVISION_MUL = Class (TOF)
      LaListe : THGrid ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure LaListeDblClick(sender : Tobject) ;
  end ;

procedure AFLanceFiche_MulRevision ;

Implementation

procedure TOF_AFREVISION_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFREVISION_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFREVISION_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFREVISION_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFREVISION_MUL.LaListeDblClick(sender : Tobject) ;
Var Clause_where,affaire : String ;
begin
  try
  affaire:=GetField('AFR_AFFAIRE') ;
  //affaire:='pascal' ;
  Clause_where:=' WHERE AFR_AFFAIRE ="'+affaire+'" '  ;
  AglLanceFicheAFREVISION('','Clause_where='+Clause_where+';ACTION=MODIFICATION') ;
  except
  end ;
TFMul(Ecran).ChercheClick;
end ;

procedure TOF_AFREVISION_MUL.OnArgument (S : String ) ;
begin
  Inherited ;
   LaListe:=THGrid(GetControl('Fliste')) ;
   LaListe.OnDblClick:= LaListeDblClick;
end ;

procedure TOF_AFREVISION_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AFREVISION_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFREVISION_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure AFLanceFiche_MulRevision ;
begin
AglLanceFiche('AFF','AFREVISION_MUL','','','');
end ;

Initialization
  registerclasses ( [ TOF_AFREVISION_MUL ] ) ;
end.

{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 22/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFORMULEVAR_MUL ()
Mots clefs ... : TOF;AFORMULEVAR_MUL
*****************************************************************}
Unit uTofAFormuleVar_Mul ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
      Fe_Main,mul,
{$Else}
     MainEagl,emul,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,uTofAFormuleVar,HTB97 ;

Type
  TOF_AFORMULEVAR_MUL = Class (TOF)
    LaListe : THGrid ;
    binsert : TToolbarButton97 ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure LaListeDblClick(sender : Tobject) ;
    procedure BinsertClick(sender : Tobject) ;

  end ;

procedure AFLanceFicheFormuleVar_Mul ;
Implementation


procedure TOF_AFORMULEVAR_MUL.LaListeDblClick(sender : Tobject) ;
Var Formule : String ;
begin
  try
    Formule:=GetField('AVF_FORMULEVAR') ;

    AFLanceFiche_FormuleVariable('','AVF_FORMULEVAR='+Formule+';ACTION=MODIFICATION') ;
  except
    AFLanceFiche_FormuleVariable('','AVF_FORMULEVAR=;ACTION=MODIFICATION') ;
  end ;
  TFMul(Ecran).Retour:=Formule;
  TFMul(Ecran).ChercheClick;
end ;

procedure TOF_AFORMULEVAR_MUL.BinsertClick(sender : Tobject) ;
begin
AFLanceFiche_FormuleVariable('','ACTION=CREATION') ;
TFMul(Ecran).ChercheClick;
end ;


procedure TOF_AFORMULEVAR_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFORMULEVAR_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFORMULEVAR_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFORMULEVAR_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFORMULEVAR_MUL.OnArgument (S : String ) ;
begin
  Inherited ;
  LaListe:=THGrid(GetControl('Fliste')) ;
  binsert:=TToolbarButton97(GetControl('Binsert')) ;
  LaListe.OnDblClick:= LaListeDblClick;
  binsert.OnClick:=binsertClick ;
end ;

procedure TOF_AFORMULEVAR_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AFORMULEVAR_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFORMULEVAR_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure AFLanceFicheFormuleVar_Mul ;
begin
AglLanceFiche ('AFF','AFORMULEVAR_MUL','','','');
end ;

Initialization
  registerclasses ( [ TOF_AFORMULEVAR_MUL ] ) ;
end.

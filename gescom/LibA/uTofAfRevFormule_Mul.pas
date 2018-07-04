{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFREVFORMULE ()
Mots clefs ... : TOF;AFREVFORMULE
*****************************************************************}
Unit uTofAfRevFormule_Mul;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main, mul,
{$Else}
     MainEagl, emul,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,HTB97,uTofAfRevFormule ;

Type
  TOF_AFREVFORMULE_MUL = Class (TOF)
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

procedure AFLanceFiche_MulFormule ;
Implementation

procedure TOF_AFREVFORMULE_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFREVFORMULE_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFREVFORMULE_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFREVFORMULE_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFREVFORMULE_MUL.LaListeDblClick(sender : Tobject) ;
Var indice : String ;
begin
try
  indice:=GetField('AFE_FORCODE') ;
  AFLanceFiche_Formule('','AFE_FORCODE='+indice+';ACTION=MODIFICATION') ;
  except
  AFLanceFiche_Formule('','AFE_FORCODE=;ACTION=CREATION') ;
  end ;
TFMul(Ecran).ChercheClick;
end ;

procedure TOF_AFREVFORMULE_MUL.BinsertClick(sender : Tobject) ;
begin
AFLanceFiche_Formule('','AFE_FORCODE=;ACTION=CREATION') ;
TFMul(Ecran).ChercheClick;

end ;
procedure TOF_AFREVFORMULE_MUL.OnArgument (S : String ) ;
begin
  Inherited ;
  LaListe:=THGrid(GetControl('Fliste')) ;
  binsert:=TToolbarButton97(GetControl('Binsert')) ;
  LaListe.OnDblClick:= LaListeDblClick;
  binsert.OnClick:=binsertClick ;
end ;

procedure TOF_AFREVFORMULE_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AFREVFORMULE_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFREVFORMULE_MUL.OnCancel () ;
begin
  Inherited ;
end;

procedure AFLanceFiche_MulFormule ;
begin
AglLanceFiche ('AFF','AFREVFORMULE_MUL','','','');
end ;

Initialization
  registerclasses ( [ TOF_AFREVFORMULE_MUL] ) ;
end.

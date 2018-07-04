{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFINDICE_MUL ()
Mots clefs ... : TOF;AFINDICE_MUL
*****************************************************************}
Unit uTofAfIndiceMul ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main, mul,
{$Else}
     MainEagl, emul,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,HTB97,UtobView,uTomAfIndice,AglInit;
                                       
Type
  TOF_AFINDICE_MUL = Class (TOF)
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


procedure AFLanceFiche_MulIndice ;

Implementation

procedure TOF_AFINDICE_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFINDICE_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFINDICE_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFINDICE_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFINDICE_MUL.LaListeDblClick(sender : Tobject) ;
Var indice : String ;
begin

  {$IFDEF EAGLCLIENT}
  TheMulQ:=TFMul(Ecran).Q.TQ;
  {$ELSE}                    
  TheMulQ:=TFMul(Ecran).Q;
  {$ENDIF}

  try
    indice:=GetField('AIN_INDCODE') ;
    AglLanceFicheAFINDICE(indice,'ACTION=MODIFICATION') ;
  except
    AglLanceFicheAFINDICE('','ACTION=CREATION') ;
  end ;
  TFMul(Ecran).ChercheClick;
end ;

procedure TOF_AFINDICE_MUL.BinsertClick(sender : Tobject) ;
begin
AglLanceFicheAFINDICE('','ACTION=CREATION') ;
TFMul(Ecran).ChercheClick;
end ;


procedure TOF_AFINDICE_MUL.OnArgument (S : String ) ;

begin
  Inherited ;

  LaListe:=THGrid(GetControl('Fliste')) ;
  binsert:=TToolbarButton97(GetControl('Binsert')) ;
  LaListe.OnDblClick:= LaListeDblClick;
  binsert.OnClick:=binsertClick ;
end ;

procedure TOF_AFINDICE_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AFINDICE_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFINDICE_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure AFLanceFiche_MulIndice ;
begin
AglLanceFiche ('AFF','AFINDICE_MUL','','','');
end ;


Initialization
  registerclasses ( [ TOF_AFINDICE_MUL ] ) ;
end.


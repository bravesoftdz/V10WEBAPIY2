{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 01/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFPUBLICATION_MUL ()
Mots clefs ... : TOF;AFPUBLICATION_MUL
*****************************************************************}
Unit uTofAfPublicationMul ;

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
     UTOF, HTB97, UtobView, uTomAfPublication, AGLInit ;

Type
  TOF_AFPUBLICATION_MUL = Class (TOF)
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

procedure AFLanceFiche_MulPublication ;
Implementation

procedure TOF_AFPUBLICATION_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFPUBLICATION_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFPUBLICATION_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFPUBLICATION_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFPUBLICATION_MUL.LaListeDblClick(sender : Tobject) ;
Var Pub : String ;
begin

  {$IFDEF EAGLCLIENT}
  TheMulQ:=TFMul(Ecran).Q.TQ;
  {$ELSE}                    
  TheMulQ:=TFMul(Ecran).Q;
  {$ENDIF}

  try
    Pub:=GetField('AFP_PUBCODE') ;
    AglLanceFicheAFPUBLICATION(Pub,'ACTION=MODIFICATION');
  except
    AglLanceFicheAFPUBLICATION('','ACTION=CREATION');
  end ;
  TFMul(Ecran).ChercheClick;

end ;

procedure TOF_AFPUBLICATION_MUL.BinsertClick(sender : Tobject) ;
begin
  {$IFDEF EAGLCLIENT}
  TheMulQ:=TFMul(Ecran).Q.TQ;
  {$ELSE}                    
  TheMulQ:=TFMul(Ecran).Q;
  {$ENDIF}

  AglLanceFicheAFPUBLICATION('','ACTION=CREATION');
  TFMul(Ecran).ChercheClick;
end ;
   
procedure TOF_AFPUBLICATION_MUL.OnArgument (S : String ) ;


begin
  Inherited ;
  LaListe:=THGrid(GetControl('Fliste')) ;
  binsert:=TToolbarButton97(GetControl('Binsert')) ;
  LaListe.OnDblClick:= LaListeDblClick;
  binsert.OnClick:=binsertClick ;

end ;

procedure TOF_AFPUBLICATION_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AFPUBLICATION_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFPUBLICATION_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure AFLanceFiche_MulPublication ;
begin
  AglLanceFiche ('AFF','AFPUBLICATION_MUL','','','');
end ;

Initialization
  registerclasses ( [ TOF_AFPUBLICATION_MUL ] ) ;
end.

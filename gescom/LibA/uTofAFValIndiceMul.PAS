{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFVALINDICE ()
Mots clefs ... : TOF;AFVALINDICE
*****************************************************************}
Unit uTofAFValIndiceMul ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,FichList,mul,
{$Else}
     MainEagl,eFichList,emul,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1, HQry,
     HMsgBox,
     UTOF,HTB97,UtobView,uTomAfValIndice, AglInit ;

Type
  TOF_AFVALINDICE_MUL = Class (TOF)
    LaListe   : THGrid ;
    binsert   : TToolbarButton97 ;
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

  end;

procedure AFLanceFiche_MulValIndice ;

Implementation

procedure TOF_AFVALINDICE_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFVALINDICE_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFVALINDICE_MUL.OnUpdate ;
begin
  Inherited;
end;

procedure TOF_AFVALINDICE_MUL.OnLoad ;
begin
  Inherited;
end;

procedure TOF_AFVALINDICE_MUL.LaListeDblClick(sender : Tobject) ;
Var AFV_PUBCODE,AFV_INDCODE,AFV_INDDATEVAL,st : String ;
begin
  Inherited ;
  try
    AFV_PUBCODE:=GetField('AFV_PUBCODE') ;
    AFV_INDCODE:=GetField('AFV_INDCODE') ;
    AFV_INDDATEVAL:=GetField('AFV_INDDATEVAL') ;
    st:=AFV_PUBCODE+';'+AFV_INDCODE+';'+AFV_INDDATEVAL ;
    AglLanceFicheAFVALINDICE(st,'ACTION=MODIFICATION' ) ;
  except
    AglLanceFicheAFVALINDICE('','ACTION=CREATION' ) ;
  end ;
  TFMul(Ecran).ChercheClick;
end ;

procedure TOF_AFVALINDICE_MUL.BinsertClick(sender : Tobject) ;
begin
  Inherited ;
  AglLanceFicheAFVALINDICE('','ACTION=CREATION' ) ;
  TFMul(Ecran).ChercheClick;
end ;


procedure TOF_AFVALINDICE_MUL.OnArgument (S : String ) ;
begin
  Inherited ;

  LaListe:=THGrid(GetControl('Fliste')) ;
  binsert:=TToolbarButton97(GetControl('Binsert')) ;
  LaListe.OnDblClick:= LaListeDblClick;
  binsert.OnClick:=binsertClick ;
end;

procedure TOF_AFVALINDICE_MUL.OnClose ;
begin
  Inherited ;
end ;
                                              
procedure TOF_AFVALINDICE_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFVALINDICE_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure AFLanceFiche_MulValIndice ;
begin
AglLanceFiche('AFF','AFVALINDICE_MUL','','','');
end ;

Initialization
  registerclasses ( [ TOF_AFVALINDICE_MUL ] ) ;
end.

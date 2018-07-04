{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 16/10/2017
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTMARCHEST_MUL ()
Mots clefs ... : TOF;BTMARCHEST_MUL
*****************************************************************}
Unit BTMARCHEST_MUL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
{$ENDIF}
     uTob, 
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF ; 

Type
  TOF_BTMARCHEST_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TOBParam : TOB;
    TOBSSTRAIT : TOB;
    TOBPieceTrait : TOB;
    //
    procedure GSDblclick (Sender : TObject);
  end ;

Implementation

procedure TOF_BTMARCHEST_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTMARCHEST_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTMARCHEST_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTMARCHEST_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTMARCHEST_MUL.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBParam := LaTOB;
  TOBSSTRAIT := TOB(LaTOB.Data);
  TOBPieceTrait := TOB(TOBSSTRAIT.data);
  TOBSSTRAIT.data := nil;
  THGrid(GetControl('fListe')).OnDblClick := GSDblclick;
end ;

procedure TOF_BTMARCHEST_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTMARCHEST_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTMARCHEST_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTMARCHEST_MUL.GSDblclick(Sender: TObject);
var Tiers,CodeMarche : string;
begin
  Tiers:= TFMul(Ecran).Fliste.datasource.dataset.FindField('BM1_FOURNISSEUR').AsString;
  CodeMarche:=TFMul(Ecran).Fliste.datasource.dataset.FindField('BM1_CODEMARCHE').AsString;
  TOBParam.putValue('SOUSTRAIT',Tiers);
  TOBParam.putValue('CODEMARCHE',CodeMarche);
  Ecran.Close;
end;

Initialization
  registerclasses ( [ TOF_BTMARCHEST_MUL ] ) ; 
end.

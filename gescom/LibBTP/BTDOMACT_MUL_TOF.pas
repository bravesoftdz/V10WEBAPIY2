{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 30/05/2008
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTDOMACT_MUL ()
Mots clefs ... : TOF;BTDOMACT_MUL
*****************************************************************}
Unit BTDOMACT_MUL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     DBGrids,
     FE_Main,
     mul,
{$else}
     eMul,
     uTob,
     Maineagl,
{$ENDIF}
     forms,
     HDB,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     BTMUL_TOF,
     HTB97,
     UTOF ;

Type
  TOF_BTDOMACT_MUL = Class (TOF_BTMUL)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
	Private
    //
    BOuvrir1   	: TToolBarButton97;
    BInsert			: TToolBarButton97;
    //
		FListe 			: THDbGrid;
    //
		Procedure FlisteDblClick (Sender : TObject);
    Procedure BInsert_OnClick(Sender: TObject);
    //
  end ;

Implementation

procedure TOF_BTDOMACT_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTDOMACT_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTDOMACT_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTDOMACT_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTDOMACT_MUL.OnArgument (S : String ) ;
begin
  Inherited ;

  BInsert := TToolbarButton97(ecran.FindComponent('BInsert'));
  BInsert.OnClick := BInsert_OnClick;
  //
  Fliste := THDbGrid(GetControl('FLISTE'));
  Fliste.OnDblClick := FlisteDblClick;

end ;

procedure TOF_BTDOMACT_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTDOMACT_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTDOMACT_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTDOMACT_MUL.FListeDblClick(Sender: TObject);
var critere : String;
		Action 	: String;
begin

   critere := FListe.datasource.dataset.FindField('BTD_CODE').AsString;

   Action := GetControlText('XXAction');

   AGLLanceFiche ('BTP','BDOMAINEACTIVITE','',critere,'ACTION=MODIFICATION');

	 TToolBarButton97(GetCOntrol('Bcherche')).Click;

end;

procedure TOF_BTDOMACT_MUL.BInsert_OnClick(Sender: TObject);
begin

  AGLLanceFiche ('BTP','BDOMAINEACTIVITE','','','ACTION=CREATION');

  TToolBarButton97(GetCOntrol('Bcherche')).Click;

end;


Initialization
  registerclasses ( [ TOF_BTDOMACT_MUL ] ) ; 
end.

 
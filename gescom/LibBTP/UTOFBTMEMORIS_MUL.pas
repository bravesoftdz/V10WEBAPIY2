{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 17/06/2008
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTMEMORIS_MUL ()
Mots clefs ... : TOF;BTMEMORIS_MUL
*****************************************************************}
Unit UTOFBTMEMORIS_MUL ;

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
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1, 
     HMsgBox,
     HTB97,
     HDB,
     UTOF;
//

Type
  TOF_BTMEMORIS_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

	Private
    BInsert : TToolBarButton97;
    BDelete	: TToolBarButton97;
    //
    FListe 	: THDbGrid;
    //
  	Procedure BOuvrir_OnClick (Sender : TObject);
	  Procedure FListe_OnDblClick(Sender: TObject);
 	  Procedure BDelete_OnClick(Sender: TObject);

  end ;

Implementation

procedure TOF_BTMEMORIS_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTMEMORIS_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTMEMORIS_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTMEMORIS_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTMEMORIS_MUL.OnArgument (S : String ) ;
begin
  Inherited ;

  //RECUPERATION DES ZONES ECRAN
	Fliste := THDbGrid (GetControl('FLISTE'));
  Fliste.OnDblClick := FListe_OnDblClick;
  //
  BInsert := TToolbarButton97(ecran.FindComponent('Binsert'));
  BInsert.OnClick := BOuvrir_OnClick;
  //
  BDelete := TToolbarButton97(ecran.FindComponent('BDelete'));
  BDelete.OnClick := BDelete_OnClick;


end ;

procedure TOF_BTMEMORIS_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTMEMORIS_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTMEMORIS_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTMEMORIS_MUL.FListe_OnDblClick(Sender : TObject);
begin

TFMul(Ecran).Retour := Fliste.datasource.dataset.FindField('BMO_CODEMEMO').AsString;

TFMUL(Ecran).Close;

end;

procedure TOF_BTMEMORIS_MUL.BOuvrir_OnClick(Sender : TObject);
begin

TFMul(Ecran).Retour := Fliste.datasource.dataset.FindField('BMO_CODEMEMO').AsString;

TFMUL(Ecran).Close;

end;

procedure TOF_BTMEMORIS_MUL.BDelete_OnClick(Sender : TObject);
Var CodeMemo : String;
		TypeMemo : String;
begin

TypeMemo := Fliste.datasource.dataset.FindField('BMO_TYPEMEMO').AsString;
CodeMemo := Fliste.datasource.dataset.FindField('BMO_CODEMEMO').AsString;

//Suppression de l'enregistrement
if (PGIAsk(TraduireMemoire('Confirmez-vous la suppression de cet enregistrement ?'), 'Suppression Mémorisation') = mrYes) then
   Begin
	 ExecuteSQL('Delete FROM BMEMORISATION Where BMO_CODEMEMO = "' + CodeMemo + '" AND BMO_TYPEMEMO="'+ TypeMemo + '"');
   end;

TToolBarButton97(GetCOntrol('Bcherche')).Click;

end;

Initialization
  registerclasses ( [ TOF_BTMEMORIS_MUL ] ) ;
end.


{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 21/01/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : EXPPLANREF ()
Mots clefs ... : TOF;EXPPLANREF
*****************************************************************}
Unit EXPPLANREF_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     dbtables, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     HSysMenu,
     Vierge,
     HTB97,
     UTOF
      ;

Type
  TOF_EXPPLANREF = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    procedure BValiderOnClick(Sender: TObject);
  end ;

Implementation

procedure TOF_EXPPLANREF.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_EXPPLANREF.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_EXPPLANREF.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_EXPPLANREF.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_EXPPLANREF.BValiderOnClick(Sender : TObject) ;
begin
  TFVierge(Ecran).retour:='1';
end ;


procedure TOF_EXPPLANREF.OnArgument (S : String ) ;
var argNumPlanRef, argCompte : string ;
    Q : TQuery ;
begin
  Inherited ;
  argNumPlanRef := ReadTokenPipe(S,';') ;
  argCompte := S ;
  Q := OpenSQL('SELECT * From PLANREF Where PR_NUMPLAN='+ArgNumPlanRef+' AND PR_COMPTE="'+argCompte+'"',true ) ;
  SetControlText('PR_COMPTE', Q.FindField('PR_COMPTE').AsString ) ;
  SetControlText('PR_LIBELLE', Q.FindField('PR_LIBELLE').AsString ) ;
  SetControlText('PR_ABREGE', Q.FindField('PR_ABREGE').AsString ) ;
  SetControlText('PR_CENTRALISABLE', Q.FindField('PR_CENTRALISABLE').AsString ) ;
  SetControlText('PR_SOLDEPROGRESSIF', Q.FindField('PR_SOLDEPROGRESSIF').AsString ) ;
  SetControlText('PR_SAUTPAGE', Q.FindField('PR_SAUTPAGE').AsString ) ;
  SetControlText('PR_TOTAUXMENSUELS', Q.FindField('PR_TOTAUXMENSUELS').AsString ) ;
  SetControlText('PR_COLLECTIF', Q.FindField('PR_COLLECTIF').AsString ) ;
  SetControlText('PR_BLOCNOTE', Q.FindField('PR_BLOCNOTE').AsString) ;
  SetControlText('PR_SENS', Q.FindField('PR_SENS').AsString ) ;
  SetControlText('PR_LETTRABLE', Q.FindField('PR_LETTRABLE').AsString ) ;
  SetControlText('PR_POINTABLE', Q.FindField('PR_POINTABLE').AsString ) ;
  SetControlText('PR_VENTILABLE1', Q.FindField('PR_VENTILABLE1').AsString ) ;
  SetControlText('PR_VENTILABLE2', Q.FindField('PR_VENTILABLE2').AsString ) ;
  SetControlText('PR_VENTILABLE3', Q.FindField('PR_VENTILABLE3').AsString ) ;
  SetControlText('PR_VENTILABLE4', Q.FindField('PR_VENTILABLE4').AsString ) ;
  SetControlText('PR_VENTILABLE5', Q.FindField('PR_VENTILABLE5').AsString ) ;
  SetControlText('PR_NATUREGENE', Q.FindField('PR_NATUREGENE').AsString ) ;
  Ferme(Q) ;
  TToolBarButton97(GetControl('BVALIDER')).OnClick := BValiderOnClick ;
end ;

procedure TOF_EXPPLANREF.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_EXPPLANREF ] ) ; 
end.

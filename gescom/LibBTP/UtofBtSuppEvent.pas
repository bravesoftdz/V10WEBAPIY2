{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 26/06/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSUPPEVENT ()
Mots clefs ... : TOF;BTSUPPEVENT
*****************************************************************}
Unit UtofBtSuppEvent ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
{$else}
     eMul,
     uTob,
{$ENDIF}
     AglInit,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF ;

Type
  TOF_BTSUPPEVENT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

procedure TOF_BTSUPPEVENT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSUPPEVENT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSUPPEVENT.OnUpdate ;
begin
  Inherited ;

	If TRadioButton(GetControl('ENSEMBLE')).Checked then
     LaTob.PutValue('RETOUR', 1)
  Else If TRadioButton(GetControl('EQUIPE')).Checked then
     LaTob.PutValue('RETOUR', 2)
  Else If TRadioButton(GetControl('EVENEMENT')).Checked then
     LaTob.PutValue('RETOUR', 3);

	TheTob := LaTob;

end ;

procedure TOF_BTSUPPEVENT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTSUPPEVENT.OnArgument (S : String ) ;
begin
  Inherited ;

  LaTob.PutValue('RETOUR', 0);
  ecran.caption := 'Action à effectuer';
  UpdateCaption(ecran);
  if LaTob.GetValue('ACTION') = 'S' then
     Begin
     TRadioButton(GetControl('ENSEMBLE')).Caption := 'Supprimer l''ensemble de l''affectation';
     TRadioButton(GetControl('EQUIPE')).Caption := 'Supprimer toute l''équipe';
     TRadioButton(GetControl('EVENEMENT')).Caption := 'Supprimer l''Evénement Seul';
     end
  else If LaTob.GetValue('ACTION') = 'E' then
  	 Begin
     TRadioButton(GetControl('ENSEMBLE')).Caption := 'Etirer l''ensemble de l''affectation';
     TRadioButton(GetControl('EQUIPE')).Caption := 'Etirer toute l''équipe';
     TRadioButton(GetControl('EVENEMENT')).Caption := 'Etirer l''Evénement Seul';
     end
  else If LaTob.GetValue('ACTION') = 'R' then
  	 Begin
     TRadioButton(GetControl('ENSEMBLE')).Caption := 'Réduire l''ensemble de l''affectation';
     TRadioButton(GetControl('EQUIPE')).Caption := 'Réduire toute l''équipe';
     TRadioButton(GetControl('EVENEMENT')).Caption := 'Réduire l''Evénement Seul';
     end
  else If LaTob.GetValue('ACTION') = 'D' then
  	 Begin
     TRadioButton(GetControl('ENSEMBLE')).Caption := 'Déplacer l''ensemble de l''affectation';
     TRadioButton(GetControl('EQUIPE')).Caption := 'Déplacer toute l''équipe';
     TRadioButton(GetControl('EVENEMENT')).Caption := 'Déplacer l''Evénement Seul';
     end
  else If LaTob.GetValue('ACTION') = 'M' then
  	 Begin
     TRadioButton(GetControl('ENSEMBLE')).Caption := 'Modifier l''ensemble de l''affectation';
     TRadioButton(GetControl('EQUIPE')).Caption := 'Modifier toute l''équipe';
     TRadioButton(GetControl('EVENEMENT')).Caption := 'Modifier l''Evénement Seul';
     end;

  if LaTob.GetValue('EQUIPE') = '' then TRadioButton(GetControl('EQUIPE')).Enabled := false;

end ;

procedure TOF_BTSUPPEVENT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTSUPPEVENT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSUPPEVENT.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTSUPPEVENT ] ) ; 
end.

 

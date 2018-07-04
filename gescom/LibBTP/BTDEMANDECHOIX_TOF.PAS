{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 14/04/2010
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTDEMANDECHOIX ()
Mots clefs ... : TOF;BTDEMANDECHOIX
*****************************************************************}
Unit BTDEMANDECHOIX_TOF ;

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
  TOF_BTDEMANDECHOIX = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	TOBSELECT : TOB;
  end ;

Implementation

procedure TOF_BTDEMANDECHOIX.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTDEMANDECHOIX.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTDEMANDECHOIX.OnUpdate ;
begin
  Inherited ;
  if TRadioButton (GetControl('CHOIX2')).Checked then TOBSelect.putValue('RETOUR',2)
  else if TRadioButton (GetControl('CHOIX3')).Checked then TOBSelect.putValue('RETOUR',3)
  else if TRadioButton (GetControl('CHOIX4')).Checked then TOBSelect.putValue('RETOUR',4);
end ;

procedure TOF_BTDEMANDECHOIX.OnLoad ;
begin
  Inherited ;
  TRadioButton (GetControl('CHOIX1')).Checked := true;
  if TRadioButton (GetControl('CHOIX2')).caption = '' then TRadioButton (GetControl('CHOIX2')).visible := false;
  if TRadioButton (GetControl('CHOIX3')).caption = '' then TRadioButton (GetControl('CHOIX3')).visible := false;
  if TRadioButton (GetControl('CHOIX4')).caption = '' then TRadioButton (GetControl('CHOIX4')).visible := false;
end ;

procedure TOF_BTDEMANDECHOIX.OnArgument (S : String ) ;
var NbChoix : integer;
begin
  Inherited ;
  TOBSELECT := LaTOB;
  NbChoix := TOBSELECT.GetValue('NBCHOIX');
  ThLabel(getControl('TEXTE1')).Caption := TOBSelect.getValue('TEXTE1');
  ThLabel(getControl('TEXTE2')).Caption := TOBSelect.getValue('TEXTE2');
  TRadioButton (GetControl('CHOIX4')).caption := TOBSelect.getValue('CHOIX4');
  TRadioButton (GetControl('CHOIX3')).caption := TOBSelect.getValue('CHOIX3');
  TRadioButton (GetControl('CHOIX2')).Caption := TOBSelect.getValue('CHOIX2');
  TRadioButton (GetControl('CHOIX1')).Caption := TOBSelect.getValue('CHOIX1');
  ecran.caption := TOBSelect.getvalue('PERIODE');
end ;

procedure TOF_BTDEMANDECHOIX.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTDEMANDECHOIX.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTDEMANDECHOIX.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTDEMANDECHOIX ] ) ; 
end.

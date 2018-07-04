{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 16/03/2010
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCHOIXAFFAIRE ()
Mots clefs ... : TOF;BTAFFCHANGETIER
*****************************************************************}
Unit BTCHOIXAFFAIRE_TOF ;

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
		 M3FP,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     uTOB,
     UTOF,
     AGLInit,
     UtilSaisieconso, AffaireUtil ;

Type
  TOF_BTCHOIXAFFAIRE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    THAffaire     : THEdit;

    procedure ListeAffaire(Sender: TObject);
  end ;

Implementation

procedure TOF_BTCHOIXAFFAIRE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXAFFAIRE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXAFFAIRE.OnUpdate ;
begin
  Inherited ;
  if ThEdit(getControl('AFFAIRE')).Text <> '' then
  begin
    LaTOB.PutValue ('RETOUR',ThEdit(getControl('AFFAIRE')).Text);
    TheTOB := LaTOB;
  end;
end ;

procedure TOF_BTCHOIXAFFAIRE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXAFFAIRE.OnArgument (S : String ) ;
begin
  Inherited ;
  THAffaire	:= THEdit(GetControl('AFFAIRE'));
  THAffaire.OnElipsisClick := ListeAffaire;
end ;

procedure TOF_BTCHOIXAFFAIRE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXAFFAIRE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXAFFAIRE.OnCancel () ;
begin
  Inherited ;
end ;

Procedure TOF_BTCHOIXAFFAIRE.ListeAffaire(Sender : TObject);
begin
 THAFFAIRE.text := GetChantier (THAFFAIRE.text);
 THLabel(GetControl('LIBAFFAIRE')).Caption := GetChampsAffaire (THAFFAIRE.text,'AFF_LIBELLE');
end;

Initialization
  registerclasses ( [ TOF_BTCHOIXAFFAIRE ] ) ;

end.

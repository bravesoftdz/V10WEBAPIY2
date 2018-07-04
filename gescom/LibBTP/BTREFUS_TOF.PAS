{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 25/09/2013
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTREFUS ()
Mots clefs ... : TOF;BTREFUS
*****************************************************************}
Unit BTREFUS_TOF ;

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
     AglInit,
     HMsgBox,
     UTOF ;

Type
  TOF_BTREFUS = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    OneTOB : TOB;
  end ;

Implementation

procedure TOF_BTREFUS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTREFUS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTREFUS.OnUpdate ;
begin
  Inherited ;
  if GetControlText('MOTIFREFUS')='' then
  begin
    PGIError('Merci de renseigner le motif du refus');
    Ecran.ModalResult := 0;
    Exit;
  end;
  LaTOB.SetString('RESULT','X');
  TheTOB := LATob;
end ;

procedure TOF_BTREFUS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTREFUS.OnArgument (S : String ) ;
begin
  Inherited ;
  OneTOB := LaTOB;
  SetControlText('DATEREFUS',DateToStr(OneTOB.GetDateTime('DATEREFUS')));
  SetControlText('MOTIFREFUS',OneTOB.GetString('MOTIFREFUS'));
end ;

procedure TOF_BTREFUS.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTREFUS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTREFUS.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTREFUS ] ) ; 
end.

{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 25/06/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTDEFLICENCES ()
Mots clefs ... : TOF;BTDEFLICENCES
*****************************************************************}
Unit BTDEFLICENCES_TOF ;

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
     HTB97,
     Paramsoc, 
     UTOF ; 

Type
  TOF_BTDEFLICENCES = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    procedure BALEATOIREClick (Sender : TObject);
  end ;

Implementation

procedure TOF_BTDEFLICENCES.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTDEFLICENCES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTDEFLICENCES.OnUpdate ;
var Mess : string;
begin
  Inherited ;
  //
  if GetControlText('ID')='' then
  begin
    PGIError('Vous devez renseigner une identification pour le client / Database');
    Exit;
  end;
  Mess := 'ATTENTION : Vous allez mettre en place la gestion des licences pour BUSINESS BTP'+#13#10+' Confirmez-vous ?';
  if PGIAsk(Mess)=Mryes then
  begin
     SetParamSoc('SO_IDSECURITE',GetControlText('ID'));
     SetParamSoc('SO_BTACTIVATE',True);
  end;
end ;

procedure TOF_BTDEFLICENCES.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTDEFLICENCES.OnArgument (S : String ) ;
begin
  Inherited ;
  SetControlText('ID',GetParamSoc('SO_IDSECURITE'));
  if GetControl('BALEATOIRE') <> nil then
  begin
    TToolbarButton97 (GetControl('BALEATOIRE')).OnClick := BALEATOIREClick;
  end;
end ;

procedure TOF_BTDEFLICENCES.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTDEFLICENCES.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTDEFLICENCES.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTDEFLICENCES.BALEATOIREClick(Sender: TObject);
begin
  SetControlText('ID',AglGetGuid);
end;

Initialization
  registerclasses ( [ TOF_BTDEFLICENCES ] ) ; 
end.

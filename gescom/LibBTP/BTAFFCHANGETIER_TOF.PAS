{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 16/03/2010
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTAFFCHANGETIER ()
Mots clefs ... : TOF;BTAFFCHANGETIER
*****************************************************************}
Unit BTAFFCHANGETIER_TOF ;

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
     TiersUtil ;

Type
  TOF_BTAFFCHANGETIER = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	UneTob : TOB;
  end ;

Implementation

procedure TOF_BTAFFCHANGETIER.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFCHANGETIER.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFCHANGETIER.OnUpdate ;
begin
  Inherited ;
  if ThEdit(getControl('TIERS')).Text <> '' then
  begin
    if not ExisteSql('SELECT T_FERME FROM TIERS WHERE T_TIERS="'+ThEdit(getControl('TIERS')).Text+'" AND T_NATUREAUXI="CLI"') then
    begin
      PgiError ('Ce client n''existe pas ou est fermé');
      ecran.ModalResult := mrNone;
      exit;
    end;
    UneTOB.PutValue ('RETOUR',ThEdit(getControl('TIERS')).Text);
  end;
end ;

procedure TOF_BTAFFCHANGETIER.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFCHANGETIER.OnArgument (S : String ) ;
begin
  Inherited ;
  UneTOB := LaTOB;
end ;

procedure TOF_BTAFFCHANGETIER.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFCHANGETIER.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFCHANGETIER.OnCancel () ;
begin
  Inherited ;
end ;

Function AGLControleTiers( parms: array of variant; nb: integer ): variant;
var CodeTiers : string;
		QQ : TQuery;
    EtatRisque  : string;
begin
	result := 0;
  CodeTiers := Parms[0];
  QQ := OpenSql ('SELECT T_FERME FROM TIERS WHERE T_TIERS="'+CodeTiers+'" AND T_NATUREAUXI="CLI"',true,1,'',true);
  if not QQ.eof then
  begin
  	if (QQ.Findfield ('T_FERME') .AsString = 'X') then result := -1;
    if result = 0 then
    begin
      EtatRisque := GetEtatRisqueClient (CodeTiers) ;
      if EtatRisque = 'R' then BEGIN result := -1 END;
    end;
  end else
  begin
		result := -1
  end;
  ferme (QQ);
end;


Initialization
  registerclasses ( [ TOF_BTAFFCHANGETIER ] ) ;
	RegisterAglFunc('ControleTiers',False,1,AGLControleTiers) ;

end.

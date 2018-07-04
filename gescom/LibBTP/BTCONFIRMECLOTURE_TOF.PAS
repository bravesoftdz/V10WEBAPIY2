{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/03/2009
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCONFIRMECLOTURE ()
Mots clefs ... : TOF;BTCONFIRMECLOTURE
*****************************************************************}
Unit BTCONFIRMECLOTURE_TOF ;

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
		 HrichOle,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HTB97,
     Hpanel ;

Type
  TOF_BTCONFIRMECLOTURE = Class (TOF)
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

procedure TOF_BTCONFIRMECLOTURE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMECLOTURE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMECLOTURE.OnUpdate ;
begin
  Inherited ;
  laTOB.PutValue('CONFIRMATION','X');

  //FV1 : 11/09/2013 - FS#654 - BAGE : en cloture de chantier, ajouter la saisie de la date de clôture
  if Assigned(GetControl('AFF_DATEFIN')) then LaTob.PutValue('AFF_DATEFIN', GetControlText('AFF_DATEFIN'));
  //
end ;

procedure TOF_BTCONFIRMECLOTURE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMECLOTURE.OnArgument (S : String ) ;
begin
  Inherited ;

  //FV1 : 11/09/2013 - FS#654 - BAGE : en cloture de chantier, ajouter la saisie de la date de clôture
  if Assigned(GetControl('AFF_DATEFIN')) then THEdit(GetControl('AFF_DATEFIN')).Visible := False;
  if Assigned(GetControl('AV_TXT3'))     then THLabel(GetControl('AV_TXT3')).Visible := False;

  if laTOB.GetValue('MODECLOTURE')='AVANCEMENT' then
  begin
  	THpanel(GetControl('PAVANC')).visible := true;
  end else if laTOB.GetValue('MODECLOTURE')='MEMOIREFINAL' then
  begin
  end else if laTOB.GetValue('MODECLOTURE')='MEMOIREFINAL' then
  begin
  end else
  begin
  	THpanel(GetControl('PAVANC')).visible := true;
    //FV1 : 11/09/2013 - FS#654 - BAGE : en cloture de chantier, ajouter la saisie de la date de clôture
    if Assigned(GetControl('AFF_DATEFIN')) then THEdit(GetControl('AFF_DATEFIN')).Visible := True;
    if Assigned(GetControl('AV_TXT3'))     then THLabel(GetControl('AV_TXT3')).Visible    := True;
    if Assigned(GetControl('AFF_DATEFIN')) then THEdit(GetControl('AFF_DATEFIN')).Text    := DateToStr(Now);
  end;

end ;

procedure TOF_BTCONFIRMECLOTURE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMECLOTURE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMECLOTURE.OnCancel () ;
begin
  Inherited ;
end ;


Initialization
  registerclasses ( [ TOF_BTCONFIRMECLOTURE ] ) ;
end.

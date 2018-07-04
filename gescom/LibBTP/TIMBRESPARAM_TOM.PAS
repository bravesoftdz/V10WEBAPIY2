{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 25/08/2011
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : TIMBRESPARAM (TIMBRESPARAM)
Mots clefs ... : TOM;TIMBRESPARAM
*****************************************************************}
Unit TIMBRESPARAM_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     Fiche, 
     FichList, 
{$else}
     eFiche, 
     eFichList, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     UTOM, 
     UTob ;

Type
  TOM_TIMBRESPARAM = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    procedure OnLoadAlerte               ; override ;
    procedure OnAfterCommit              ; override ;
    procedure OnAfterRollBack            ; override ;
    procedure OnAfterBeginTrans          ; override ;
  	private
    procedure SetLastError(Num: integer; ou: string);
    procedure ModePaieChange (Sender : TObject);
    end ;

const
  // libellés des messages
  TexteMessage: array[1..5] of string = (
    {1}  'Vous devez renseigner une désignation'
    {2}, 'Vous devez le pourcentage'
    {3}, 'Vous devez renseigner le montant'
    {4}, 'Vous devez renseigner la valeur de la tranche'
    {5}, 'Vous devez renseigner le compte comptable'
    );

Implementation

procedure TOM_TIMBRESPARAM.OnNewRecord ;
begin
  Inherited ;
  SetField('BTP_TYPETIMBRE', '001');
  SetField('BTP_MAXI', '999999');
  SetField('BTP_LISTEMODPAIE', '');
  SetControlText('LISTEMODPAIE','');
end ;

procedure TOM_TIMBRESPARAM.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_TIMBRESPARAM.OnUpdateRecord ;
begin
  Inherited ;

  if GetField('BTP_LIBELLE')='001' then
  begin
  	SetLastError(1,'BTP_LIBELLE');
    Exit;
  end;

  if GetField('BTP_TYPETIMBRE')='001' then
  begin
    if GetField('BTP_POURCENT')=0 then
    begin
      SetLastError(2,'BTP_POURCENT');
      Exit;
    end;
  end else if GetField('BTP_TYPETIMBRE')='002' then
  begin
    if GetField('BTP_MONTANT')=0 then
    begin
      SetLastError(3,'BTP_MONTANT');
      Exit;
    end;
  end else if GetField('BTP_TYPETIMBRE')='003' then
  begin
    if GetField('BTP_MONTANT')=0 then
    begin
      SetLastError(3,'BTP_MONTANT');
      Exit;
    end;
    if GetField('BTP_TRANCHE')=0 then
    begin
      SetLastError(4,'BTP_TRANCHE');
      Exit;
    end;

  end;
  if GetField('BTP_GENERAL')='' then
  begin
    SetLastError(5,'BTP_GENERAL');
    Exit;
  end;
end ;

procedure TOM_TIMBRESPARAM.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_TIMBRESPARAM.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_TIMBRESPARAM.OnLoadRecord ;
begin
  Inherited ;
  SetControlText('LISTEMODPAIE',getField('BTP_LISTEMODPAIE'));
  THMultiValComboBox (GetControl('LISTEMODPAIE')).OnChange := ModePaieChange;
end ;

procedure TOM_TIMBRESPARAM.OnChangeField ( F: TField ) ;
begin
  Inherited ;
//  if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
  if F.FieldName = 'BTP_TYPETIMBRE' then
  begin
  	if GetField('BTP_TYPETIMBRE') = '001' then
    begin
      SetControlVisible('BTP_POURCENT',true);
      SetControlVisible('TBTP_POURCENT',true);
      SetControlVisible('TTBTP_POURCENT',True);
      SetControlVisible('BTP_MONTANT',false);
      SetControlVisible('TBTP_MONTANT',false);
      SetControlVisible('BTP_SEUILAPPLIC',false);
      SetControlVisible('TBTP_SEUILAPPLIC',false);
      SetControlVisible('BTP_TRANCHE',false);
      SetControlVisible('TBTP_TRANCHE',false);
      SetControlVisible('BTP_MINI',false);
      SetControlVisible('TBTP_MINI',false);
      SetControlVisible('BTP_MAXI',false);
      SetControlVisible('TBTP_MAXI',false);
      SetControlVisible('BTP_FRACTION',false);
    end else if GetField('BTP_TYPETIMBRE') = '002' then
    begin
      SetControlVisible('BTP_MONTANT',true);
      SetControlVisible('TBTP_MONTANT',true);
      SetControlVisible('BTP_POURCENT',false);
      SetControlVisible('TBTP_POURCENT',false);
      SetControlVisible('TTBTP_POURCENT',false);
      SetControlVisible('BTP_SEUILAPPLIC',false);
      SetControlVisible('TBTP_SEUILAPPLIC',false);
      SetControlVisible('BTP_TRANCHE',false);
      SetControlVisible('TBTP_TRANCHE',false);
      SetControlVisible('BTP_MINI',false);
      SetControlVisible('TBTP_MINI',false);
      SetControlVisible('BTP_MAXI',false);
      SetControlVisible('TBTP_MAXI',false);
      SetControlVisible('BTP_FRACTION',false);
    end else if GetField('BTP_TYPETIMBRE') = '003' then
    begin
      SetControlCaption('TBTP_MONTANT','Montant de la tranche');
      SetControlVisible('BTP_MONTANT',true);
      SetControlVisible('TBTP_MONTANT',true);
      SetControlVisible('BTP_POURCENT',false);
      SetControlVisible('TBTP_POURCENT',false);
      SetControlVisible('TTBTP_POURCENT',false);
      SetControlVisible('BTP_SEUILAPPLIC',true);
      SetControlVisible('TBTP_SEUILAPPLIC',true);
      SetControlVisible('BTP_TRANCHE',true);
      SetControlVisible('TBTP_TRANCHE',true);
      SetControlVisible('BTP_MINI',true);
      SetControlVisible('TBTP_MINI',true);
      SetControlVisible('BTP_MAXI',true);
      SetControlVisible('TBTP_MAXI',true);
      SetControlVisible('BTP_FRACTION',true);
    end;
  end else if F.FieldName = 'BTP_LISTEMODPAIE' then
  begin
  end;
end ;

procedure TOM_TIMBRESPARAM.OnArgument ( S: String ) ;
begin
  Inherited ;
end ;

procedure TOM_TIMBRESPARAM.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_TIMBRESPARAM.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_TIMBRESPARAM.OnLoadAlerte;
begin
  Inherited ;
end ;

procedure TOM_TIMBRESPARAM.OnAfterBeginTrans;
begin
  Inherited ;
end ;

procedure TOM_TIMBRESPARAM.OnAfterCommit;
begin
  Inherited ;
end ;

procedure TOM_TIMBRESPARAM.OnAfterRollBack;
begin
  Inherited ;
end ;

procedure TOM_TIMBRESPARAM.SetLastError(Num: integer; ou: string);
begin
  if ou <> '' then SetFocusControl(ou);
  LastError := Num;
  LastErrorMsg := TexteMessage[LastError];
end;

procedure TOM_TIMBRESPARAM.ModePaieChange(Sender: TObject);
begin
  if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
  SetField('BTP_LISTEMODPAIE',THMultiValComboBox (GetControl('LISTEMODPAIE')).Value);
end;

Initialization
  registerclasses ( [ TOM_TIMBRESPARAM ] ) ;
end.

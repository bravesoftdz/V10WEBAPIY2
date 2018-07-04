{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 26/10/2017
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : BDEPENSESPOC (BDEPENSESPOC)
Mots clefs ... : TOM;BDEPENSESPOC
*****************************************************************}
Unit BDEPENSESPOC_TOM ;

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
  TOM_BDEPENSESPOC = Class (TOM)
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
    procedure SetEvents(State : boolean);
    procedure ChangeValue (Sender : TObject);
  end ;

Implementation

procedure TOM_BDEPENSESPOC.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_BDEPENSESPOC.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_BDEPENSESPOC.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_BDEPENSESPOC.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_BDEPENSESPOC.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_BDEPENSESPOC.OnLoadRecord ;
begin
  Inherited ;
  if getField('BM7_CODE')='' then SetField('BM7_CODE',001);
  SetControlText('MBM7_PRODPOC',GetField('BM7_PRODPOC'));
  SetControlText('MBM7_PRODMOI',GetField('BM7_PRODMOI'));
  SetControlText('MBM7_PRODFI',GetField('BM7_PRODFI'));
  SetControlText('MBM7_PRODST',GetField('BM7_PRODST'));
  SetEvents(True);
end ;

procedure TOM_BDEPENSESPOC.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_BDEPENSESPOC.OnArgument ( S: String ) ;
begin
  Inherited ;
end ;

procedure TOM_BDEPENSESPOC.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_BDEPENSESPOC.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_BDEPENSESPOC.OnLoadAlerte;
begin
  Inherited ;
end ;

procedure TOM_BDEPENSESPOC.OnAfterBeginTrans;
begin
  Inherited ;
end ;

procedure TOM_BDEPENSESPOC.OnAfterCommit;
begin
  Inherited ;
end ;

procedure TOM_BDEPENSESPOC.OnAfterRollBack;
begin
  Inherited ;
end ;

procedure TOM_BDEPENSESPOC.SetEvents(State: boolean);
begin
  if State then
  begin
    THMultiValComboBox (GetControl('MBM7_PRODPOC')).OnChange := ChangeValue;
    THMultiValComboBox (GetControl('MBM7_PRODMOI')).OnChange := ChangeValue;
    THMultiValComboBox (GetControl('MBM7_PRODFI')).OnChange := ChangeValue;
    THMultiValComboBox (GetControl('MBM7_PRODST')).OnChange := ChangeValue;
  end else
  begin
    THMultiValComboBox (GetControl('MBM7_PRODPOC')).OnChange := nil;
    THMultiValComboBox (GetControl('MBM7_PRODMOI')).OnChange := nil;
    THMultiValComboBox (GetControl('MBM7_PRODFI')).OnChange := nil;
    THMultiValComboBox (GetControl('MBM7_PRODST')).OnChange := nil;
  end;
end;

procedure TOM_BDEPENSESPOC.ChangeValue(Sender: TObject);
begin
  if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
  if THMultiValComboBox (Sender).name = 'MBM7_PRODPOC' then
  begin
    SetField('BM7_PRODPOC',GetControlText('MBM7_PRODPOC'));
  end else if THMultiValComboBox (Sender).name = 'MBM7_PRODMOI' then
  begin
    SetField('BM7_PRODMOI',GetControlText('MBM7_PRODMOI'));
  end else if THMultiValComboBox (Sender).name = 'MBM7_PRODFI' then
  begin
    SetField('BM7_PRODFI',GetControlText('MBM7_PRODFI'));
  end else if THMultiValComboBox (Sender).name = 'MBM7_PRODST' then
  begin
    SetField('BM7_PRODST',GetControlText('MBM7_PRODST'));
  end;
end;

Initialization
  registerclasses ( [ TOM_BDEPENSESPOC ] ) ; 
end.

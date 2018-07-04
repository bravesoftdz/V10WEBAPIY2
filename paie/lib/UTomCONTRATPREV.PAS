{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 13/11/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CONTRATPREV (CONTRATPREV)
Mots clefs ... : TOM;CONTRATPREV
*****************************************************************}
Unit UTomCONTRATPREV ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     Fiche, 
     FichList,
     HDB,
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
  TOM_CONTRATPREV = Class (TOM)
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
  Public
    procedure OnCBDateFinClick(Sender : TObject);
  Private
    Loading : Boolean;
  end ;

Implementation

procedure TOM_CONTRATPREV.OnNewRecord ;
begin
  Inherited ;
  SetField('POP_CODEPOPPREV', '90');
  SetField('POP_NBENFANT', 90);
  SetField('POP_DATEFIN', iDate2099);
end ;

procedure TOM_CONTRATPREV.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_CONTRATPREV.OnUpdateRecord ;
  Procedure ControleChampsNonVide(FieldName : String; MsgLibelleChamps : String = '');
    Procedure FieldFocus(FieldName : String);
    begin
      if Assigned(GetControl(FieldName)) then SetFocusControl(FieldName);
    end;
  begin
    if (GetField(FieldName) = '') then
    begin
      LastError := 1;
      if MsgLibelleChamps = '' then MsgLibelleChamps := FieldName;
      LastErrorMsg := 'Vous devez renseigner le champ '+MsgLibelleChamps;
      FieldFocus(FieldName);
    end;
  end;
begin
  if (GetField('POP_INSTITUTION') <> '') and (not IsNumeric(GetField('POP_INSTITUTION'))) then
  begin
    SetField('POP_INSTITUTION', '');
  end;
  ControleChampsNonVide('POP_INSTITUTION', 'Institution');
  ControleChampsNonVide('POP_REFERCONTRAT', 'Référence du contrat');
  Inherited ;
end ;

procedure TOM_CONTRATPREV.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_CONTRATPREV.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_CONTRATPREV.OnLoadRecord ;
begin
  Inherited ;
  if Assigned(GetControl('CB_DATEFIN')) then
  begin
    (GetControl('CB_DATEFIN') as THCheckbox).Checked := (GetField('POP_DATEFIN') <> iDate2099);
    SetControlVisible('POP_DATEFIN', (GetField('POP_DATEFIN') <> iDate2099));
  end;
  Loading := False;
end ;

procedure TOM_CONTRATPREV.OnChangeField ( F: TField ) ;
var
  TabletteName : String;
begin
  Inherited ;
  if F.FieldName = 'POP_INSTITUTION' then
  begin
    if (GetField('POP_INSTITUTION') <> '') and (not IsNumeric(GetField('POP_INSTITUTION'))) then
    begin
      SetField('POP_INSTITUTION', '');
    end;
  end;
  if F.FieldName = 'POP_DATEDEBUT' then
  begin
    if (GetField('POP_DATEDEBUT') > GetField('POP_DATEFIN')) then
    begin
      SetField('POP_DATEFIN', GetField('POP_DATEDEBUT'));
    end;
  end;
  if F.FieldName = 'POP_TYPECONTRAT' then
  begin
    TabletteName := '';
    if (GetField('POP_TYPECONTRAT') = 'DOS') then
    begin
      TabletteName := '';
      SetField('POP_NOMPOP', '');
    end else if (GetField('POP_TYPECONTRAT') = 'ETB') then
    begin
      TabletteName := 'TTETABLISSEMENT';
    end else if (GetField('POP_TYPECONTRAT') = 'POP') then
    begin
      TabletteName := 'PGPOPULATION';
    end else if (GetField('POP_TYPECONTRAT') = 'SAL') then
    begin
      TabletteName := 'PGSALARIE';
    end;
    if Assigned(GetControl('POP_NOMPOP')) then
    begin
{$IFNDEF EAGLCLIENT}
      (GetControl('POP_NOMPOP') as THDBEdit).Enabled := (TabletteName <> '');
      (GetControl('POP_NOMPOP') as THDBEdit).DataType := TabletteName;
{$else}
      (GetControl('POP_NOMPOP') as THEdit).Enabled := (TabletteName <> '');
      (GetControl('POP_NOMPOP') as THEdit).DataType := TabletteName;
{$ENDIF}
    end;
  end;
end ;

procedure TOM_CONTRATPREV.OnArgument ( S: String ) ;
begin
  Inherited ;
  { Recherche des composants de la fiche }
  Loading := True;
  if assigned(GetControl('CB_DATEFIN')) then
  begin
    (GetControl('CB_DATEFIN') as THCheckbox).OnClick := OnCBDateFinClick;
  end;
end ;

procedure TOM_CONTRATPREV.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_CONTRATPREV.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_CONTRATPREV.OnLoadAlerte;
begin
  Inherited ;
end ;

procedure TOM_CONTRATPREV.OnCBDateFinClick(Sender: TObject);
begin
  if Loading then exit;
  if (GetControl('CB_DATEFIN') as THCheckbox).Checked then
  begin
    SetField('POP_DATEFIN', GetField('POP_DATEDEBUT'));
    SetControlVisible('POP_DATEFIN', True);
  end else begin
    SetField('POP_DATEFIN', iDate2099);
    SetControlVisible('POP_DATEFIN', False);
  end;
end;

Initialization
  registerclasses ( [ TOM_CONTRATPREV ] ) ; 
end.

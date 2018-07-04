{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 14/04/2005
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PARAMFORMABS (PARAMFORMABS)
Mots clefs ... : TOM;PARAMFORMABS
*****************************************************************}
{
PT1     02/08/2005 PH V_60 FQ 12467 Ergonomie

}
Unit UTomParamFormAbs ;

Interface

Uses StdCtrls,
     Controls,
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     Fiche,
     HDB,
     FichList,
     DBCtrls,
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
     PGOutils2,
     PGOutils,
     UTob ;

Type
  TOM_PARAMFORMABS = Class (TOM)
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
    private
    LectureSeule, CEG, STD, DOS, OnFerme: boolean;
    procedure GestionRubrique;
    procedure AccesCheckBox(Sender : TObject);
    end ;

Implementation

procedure TOM_PARAMFORMABS.OnNewRecord ;
begin
  Inherited ;
  SetField('PRM_PREDEFINI', 'DOS');
  SetField('PPF_TYPEPLANPREV','DIF');
end ;

procedure TOM_PARAMFORMABS.OnDeleteRecord ;
begin
  Inherited ;
  TFFiche(Ecran).Retour := 'MODIF';
end ;

procedure TOM_PARAMFORMABS.OnUpdateRecord ;
var Predef : String;
begin
  Inherited ;
  if (DS.State = dsinsert) then
  begin
    if (GetField('PPF_PREDEFINI') <> 'DOS') then
      SetField('PPF_NODOSSIER', '000000')
    else
      SetField('PPF_NODOSSIER', PgRendNoDossier());
  end;
  Predef := GetField('PPF_PREDEFINI');
  if (Predef <> 'CEG') and (Predef <> 'DOS') and (Predef <> 'STD') then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez renseigner le champ prédéfini';
    SetFocusControl('PPF_PREDEFINI');
  end;
  TFFiche(Ecran).Retour := 'MODIF';
end ;

procedure TOM_PARAMFORMABS.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_PARAMFORMABS.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_PARAMFORMABS.OnLoadRecord ;
begin
  Inherited ;
  SetControlEnabled('PPF_RUBRIQUE',GetCheckBoxState('PPF_ALIMRUB') = CbChecked);
  SetControlEnabled('PPF_ALIMENT',GetCheckBoxState('PPF_ALIMRUB') = CbChecked);
  SetControlEnabled('PPF_PROFIL',GetCheckBoxState('PPF_ALIMRUB') = CbChecked);
  SetControlEnabled('PPF_TYPECONGE',GetCheckBoxState('PPF_ALIMABS') = CbChecked);
  SetControlEnabled('PPF_ALIMABS',GetCheckBoxState('PPF_ALIMRUB') <> CbChecked);
  SetControlEnabled('PPF_ALIMRUB',GetCheckBoxState('PPF_ALIMABS') <> CbChecked);
  GestionRubrique;
  TFFiche(Ecran).Retour := '';
  AccesPredefini('TOUS', CEG, STD, DOS);
  LectureSeule := FALSE;
  if (Getfield('PPF_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
    //*  SetControlEnabled('BDefaire',CEG);
    SetControlEnabled('BDelete', CEG);
    if CEG = True then SetControlProperty('ACCES', 'Text', 'TRUE')
    else SetControlProperty('ACCES', 'Text', 'FALSE');
    end
    else
    if (Getfield('PPF_PREDEFINI') = 'STD') then
    begin
         LectureSeule := (STD = False);
         PaieLectureSeule(TFFiche(Ecran), (STD = False));
         SetControlEnabled('BDelete', STD);
         if STD = True then SetControlProperty('ACCES', 'Text', 'TRUE')
         else SetControlProperty('ACCES', 'Text', 'FALSE');
         end
    else
    if (Getfield('PPF_PREDEFINI') = 'DOS') then
    begin
         LectureSeule := False;
         PaieLectureSeule(TFFiche(Ecran), False);
         SetControlEnabled('BDelete', DOS);
         if DOS = True then SetControlProperty('ACCES', 'Text', 'TRUE')
         else SetControlProperty('ACCES', 'Text', 'FALSE');
    end;
end ;

procedure TOM_PARAMFORMABS.OnChangeField ( F: TField ) ;
var Rubrique,TempRub,Mes,Pred,Vide : String;
    iCode: integer;
    OKRub: boolean;
begin
  Inherited ;
  if F.FieldName = 'PPF_PROFIL' then
  begin
    if getfield('PMA_PROFILABS') = '' then GestionRubrique;
  end;
  if (F.FieldName = 'PPF_PARAMFORMPAIE') then
  begin
    Rubrique := Getfield('PPF_PARAMFORMPAIE');
    if Rubrique = '' then exit;
    if ((isnumeric(Rubrique)) and (Rubrique <> '    ')) then
    begin
      iCode := strtoint(trim(Rubrique));
      TempRub := ColleZeroDevant(iCode, 4);
      if (DS.State = dsinsert) and (TempRub <> '') and (GetField('PPF_PREDEFINI') <> '') then
      begin
        OKRub := TestRubrique(GetField('PPF_PREDEFINI'), TempRub, 0);
        if (OkRub = False) or (rubrique = '0000') then
        begin
          Mes := MesTestRubrique('AAA', GetField('PPF_PREDEFINI'), 0);
          HShowMessage('2;Code erroné: ' + TempRub + ' ;' + Mes + ';W;O;O;;;', '', '');
          TempRub := '';
        end;
      end;
      if TempRub <> Rubrique then
      begin
        SetField('PPF_PARAMFORMPAIE', TempRub);
        SetFocusControl('PPF_PARAMFORMPAIE');
      end;
    end;
  end;
  if (F.FieldName = 'PPF_PREDEFINI') and (DS.State = dsinsert) then
  begin
    Pred := GetField('PPF_PREDEFINI');
    Rubrique := (GetField('PPF_PARAMFORMPAIE'));
    if Pred = '' then exit;
    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de rubrique prédéfinie CEGID', 'Accès refusé'); //PT1
      Pred := 'DOS';
      SetControlProperty('PPF_PREDEFINI', 'Value', Pred);
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de rubrique prédéfinie Standard', 'Accès refusé'); // PT1
      Pred := 'DOS';
      SetControlProperty('PPF_PREDEFINI', 'Value', Pred);
    end;
    if (rubrique <> '') and (Pred <> '') then
    begin
      OKRub := TestRubrique(pred, rubrique, 0);
      if (OkRub = False) or (rubrique = '0000') then
      begin
        Mes := MesTestRubrique('AAA', pred, 0);
        HShowMessage('2;Code Erroné: ' + Rubrique + ' ;' + Mes + ';W;O;O;;;', '', '');
        SetField('PPF_PARAMFORMPAIE', vide);
        if Pred <> GetField('PPF_PREDEFINI') then SetField('PPF_PREDEFINI', pred);
        SetFocusControl('PPF_PARAMFORMPAIE');
        exit;
      end;
    end;
    if Pred <> GetField('PPF_PREDEFINI') then SetField('PPF_PREDEFINI', pred);
  end;


end ;

procedure TOM_PARAMFORMABS.OnArgument ( S: String ) ;
var
 {$IFNDEF EAGLCLIENT}
  CheckAbs,CheckRub : TDBCheckBox;
  {$ELSE}
  CheckAbs,CheckRub : TCheckBox;
  {$ENDIF}
begin
  Inherited ;
  {$IFNDEF EAGLCLIENT}
  CheckAbs := TDBCheckBox(GetControl('PPF_ALIMABS'));
  CheckRub := TDBCheckBox(GetControl('PPF_ALIMRUB'));
  {$ELSE}
  CheckAbs := TCheckBox(GetControl('PPF_ALIMABS'));
  CheckRub := TCheckBox(GetControl('PPF_ALIMRUB'));
  {$ENDIF}
  If CheckAbs <> Nil then CheckAbs.OnClick := AccesCheckBox;
  If CheckRub <> Nil then CheckRub.OnClick := AccesCheckBox;
end ;

procedure TOM_PARAMFORMABS.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_PARAMFORMABS.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_PARAMFORMABS.GestionRubrique;
var
  profil,rubrique : boolean;
  {$IFNDEF EAGLCLIENT}
  rub: thdbedit;
  {$ELSE}
  rub: thedit;
  {$ENDIF}
begin
  profil := (getfield('PPF_PROFIL') <> '');
  rubrique := (getfield('PPF_RUBRIQUE') <> '');
  {$IFNDEF EAGLCLIENT}
  rub := thdbedit(getcontrol('PPF_RUBRIQUE'));
  {$ELSE}
  rub := thedit(getcontrol('PPF_RUBRIQUE'));
  {$ENDIF}
//  setcontrolenabled('PPF_ALIMENT', rubrique { and profil});
    if rub <> nil then
    if profil then
    begin
      Rub.datatype := 'PGRUBABSENCE';
      Rub.plus := getfield('PPF_PROFIL');
    end
    else
    begin
      Rub.datatype := 'PGREMUNERATION';
      Rub.plus := ' PPF_THEMEREM = "ABS"';
    end;
end;

procedure TOM_PARAMFORMABS.AccesCheckBox(Sender : TObject);
var
 {$IFNDEF EAGLCLIENT}
  Check : TDbCheckBox;
  {$ELSE}
  Check : TCheckBox;
  {$ENDIF}
begin
     If Sender = Nil then Exit;
      {$IFNDEF EAGLCLIENT}
      Check := TDbCheckBox(Sender);
      {$ELSE}
      Check := TCheckBox(Sender);
      {$ENDIF}
     If Check.Name = 'PPF_ALIMABS' then
     begin
          If GetCheckBoxState('PPF_ALIMABS') = CbChecked then
          begin
               SetControlEnabled('PPF_ALIMRUB',False);
               SetControlEnabled('PPF_RUBRIQUE',False);
               SetControlEnabled('PPF_ALIMENT',False);
               SetControlEnabled('PPF_PROFIL',False);
               SetField('PPF_PROFIL','');
               SetField('PPF_RUBRIQUE','');
               SetField('PPF_ALIMENT','');
               SetControlEnabled('PPF_TYPECONGE',True);
          end
          else
          begin
               SetControlEnabled('PPF_ALIMRUB',True);
               SetControlEnabled('PPF_TYPECONGE',False);
          end;
     end;
     If Check.Name = 'PPF_ALIMRUB' then
     begin
          If GetCheckBoxState('PPF_ALIMRUB') = CbChecked then
          begin
               SetControlEnabled('PPF_ALIMABS',False);
               SetControlEnabled('PPF_TYPECONGE',False);
               SetField('PPF_TYPECONGE','');
               SetControlEnabled('PPF_RUBRIQUE',True);
               SetControlEnabled('PPF_ALIMENT',True);
               SetControlEnabled('PPF_PROFIL',True);
          end
          else
          begin
               SetControlEnabled('PPF_ALIMABS',True);
               SetControlEnabled('PPF_PROFIL',False);
               SetControlEnabled('PPF_ALIMENT',False);
               SetControlEnabled('PPF_RUBRIQUE',False);
          end;
     end;
end;

Initialization
  registerclasses ( [ TOM_PARAMFORMABS ] ) ;
end.


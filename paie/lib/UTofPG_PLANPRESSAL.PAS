{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 04/07/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PLANPRESSAL ()
Mots clefs ... : TOF;PLANPRESSAL
*****************************************************************}
Unit UTofPG_PLANPRESSAL ;

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
     uTob, 
{$ENDIF}
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,
{$IFNDEF NA3}
     UTOFPGPlaning ;
{$ELSE} //Pour Nicole
     UTOFPGPlanningAbsPre ;
{$ENDIF}

Type
  TOF_PLANPRESSAL = Class (TOF_PGPLANNINGMERE)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnChangeRupture (Sender : TObject);

  private
//    procedure OnClickBValider(Sender : TObject);
  end ;

Implementation

Uses
  HTB97, PGPlanningPresenceSal, PGPlanningOutils, EntPaie, P5Def, PGOutils2;

procedure TOF_PLANPRESSAL.OnUpdate ;
var
  NiveauRupt: TNiveauRupture;
  controleOK, rythmetrav, Evenements: boolean;
begin
  NiveauRupt.ChampsRupt[1] := ChampsRupt[1];
  NiveauRupt.ChampsRupt[2] := ChampsRupt[2];
  NiveauRupt.ChampsRupt[3] := ChampsRupt[3];
  NiveauRupt.ChampsRupt[4] := ChampsRupt[4];
  NiveauRupt.CondRupt := '';
  NiveauRupt.NiveauRupt := 0;
  if      GetControlText('RUPTURE4') <> '' then
     NiveauRupt.NiveauRupt := 4
  else if GetControlText('RUPTURE3') <> '' then
    NiveauRupt.NiveauRupt := 3
  else if GetControlText('RUPTURE2') <> '' then
    NiveauRupt.NiveauRupt := 2
  else if GetControlText('RUPTURE1') <> '' then
    NiveauRupt.NiveauRupt := 1;

  //  if UtiliSat = 'ADM' then Salarie := GetControlText('RESPONS');

//  PGPlanningAbsence(StrToDate(GetControlText('DATEABS')), StrToDate(GetControlText('DATEABS_')),
//    Utilisat, '', RecupCritere('ABSENCE'), RecupCritere('SALARIE'), Salarie, NiveauRupt, (GetControlText('MULTINIVEAU') = 'X'),
//    RecupCritere('FILTRE'), RecupCritere('HIERARCHIE'), RecupCritere('CBFILTREABS'), { PT8 } { PT11 }
//    StrtoFloat(GetControltext('JOURABS')),StrtoFloat(GetControltext('SOLDEABS')), { PT12 }
//    GetControltext('TYPESOLDEABS'),mode);   { PT12 }

  controleOK := true;
  if (StrToDate(GetControlText('DATEABS')) = iDate1900)
  or (StrToDate(GetControlText('DATEABS_')) - StrToDate(GetControlText('DATEABS')) > 366) then
  begin
      PGIError('Vous devez choisir une période d''affichage d''un an maximum.', Ecran.Caption);
      controleOK := false;
  end;
  Rythmetrav  := TCheckBox(GetControl('RYTHMETRAV')).checked;
  Evenements  := TCheckBox(GetControl('EVENEMENTS')).checked;

  if not rythmetrav and not Evenements  then
  begin
      PGIError('Vous devez choisir Rythme de travail et/ou Evènements. ', Ecran.Caption);
      setfocuscontrol('RYTHMETRAV');
      controleOK := false;
  end;

   if controleOK then
    //PGPlanningPresenceSalOpen(StrToDate(GetControlText('DATEABS')),StrToDate(GetControlText('DATEABS_')),'','', NiveauRupt, [piaDemiJournee, piaSemaine, piaMois, piaHeure, piaOutlook], True, True, RecupCritere('SALARIE')+' '+RecupCritere('HIERARCHIE'));
    PGPlanningPresenceSalOpen(StrToDate(GetControlText('DATEABS')),StrToDate(GetControlText('DATEABS_')),'','', NiveauRupt, [piaDemiJournee, piaSemaine, piaMois, piaHeure, piaOutlook],Evenements,Rythmetrav, RecupCritere('SALARIE')+' '+RecupCritere('HIERARCHIE'));
end ;

procedure TOF_PLANPRESSAL.OnArgument (S : String ) ;
var
  Num: integer;
//  Edit: THEdit;
//  NumEdit: THNumEdit;
  Combo: TControl;
  GrpBox: TGroupBox;
//  ChbxH: TCheckBox;
begin
  inherited;
//  (GetControl('BValider') as TToolBarButton97).OnClick := OnClickBValider;
  for num := 1 to 4 do ChampsRupt[num] := '';
  SetControlVisible('GBRUPTURE', True);
  if VH_Paie.PGNbreStatOrg > 0 then SetControlVisible('GBORGANISATION', True);
  for Num := 1 to 4 do
    VisibiliteChampSalarie(IntToStr(Num), GetControl('PSA_TRAVAILN' + IntToStr(Num)), GetControl('TPSA_TRAVAILN' + IntToStr(Num)));
  VisibiliteStat(GetControl('PSA_CODESTAT'), GetControl('TPSA_CODESTAT'));
  Num := 1;
  while num <= VH_Paie.PgNbCombo do
  begin
    GrpBox := TGroupBox(GetControl('GBCHAMPLIBRE'));
    if GrpBox <> nil then GrpBox.visible := True;
    SetControlVisible('PSA_LIBREPCMB' + IntToStr(Num), True);
    SetControlVisible('TPSA_LIBREPCMB' + IntToStr(Num), True);
    if      (Num = 1) then SetControlText('TPSA_LIBREPCMB1', VH_Paie.PgLibCombo1)
    else if (Num = 2) then SetControlText('TPSA_LIBREPCMB2', VH_Paie.PgLibCombo2)
    else if (Num = 3) then SetControlText('TPSA_LIBREPCMB3', VH_Paie.PgLibCombo3)
    else if (Num = 4) then SetControlText('TPSA_LIBREPCMB4', VH_Paie.PgLibCombo4);
    Num := Num + 1;
  end;

//  SetControlText('NOMRESP', RechDom('PGSALARIE', GetControlText('RESPONS'), False));
//  if GetControlText('NOMRESP') = 'Error' then SetControlText('NOMRESP', '');
//  Edit := THEdit(GetControl('RESPONS'));
//  if Edit <> nil then begin Edit.OnElipsisClick := AfficheTabletteResp; Edit.OnExit := ExitRespons; end;
//  Edit := THEdit(GetControl('PSA_SALARIE'));
//  if Edit <> nil then begin Edit.OnElipsisClick := AfficheTabletteSal; Edit.OnExit := OnExitSalarie; end;
//  Edit := THEdit(GetControl('PSA_SALARIE_'));
//  if Edit <> nil then begin Edit.OnElipsisClick := AfficheTabletteSal; Edit.OnExit := OnExitSalarie; end;

  SetControlEnabled('PSA_ETABLISSEMENT', True);
  SetControlEnabled('PSA_TRAVAILN1', True);
  SetControlEnabled('PSA_TRAVAILN2', True);
  SetControlEnabled('PSA_TRAVAILN3', True);
  SetControlEnabled('PSA_TRAVAILN4', True);
  SetControlEnabled('PSA_CODESTAT', True);

  for Num := 1 to 4 do
  begin
    Combo := ThValComboBox(GetControl('RUPTURE' + IntToStr(num)));
    if Combo <> nil then THValComboBox(Combo).OnChange := OnChangeRupture;
    InitialiseCombo(Combo);
  end;

//  ChbxH := TCheckBox(GetControl('MULTINIVEAU'));
//  if Assigned(ChbxH) then ChbxH.OnClick := ClickHierarchie; { PT8 }

//  NumEdit := THNumEdit(GetControl('SOLDEABS'));                 { PT12 }
//  if Assigned(NumEdit) then NumEdit.OnChange := ChangeSoldeAbs; { PT12 }

end ;

//procedure TOF_PLANPRESSAL.OnClickBValider(Sender: TObject);
//begin
//  if (StrToDate(GetControlText('DATEABS')) = iDate1900)
//  or (StrToDate(GetControlText('DATEABS_'))-StrToDate(GetControlText('DATEABS')) > 366) then
//    PGIError('Vous devez choisir une période d''affichage d''un an maximum', Ecran.Caption)
//  else
//    PGPlanningPresenceSalOpen(StrToDate(GetControlText('DATEABS')),StrToDate(GetControlText('DATEABS_')),'','', [piaSemaine, piaMois, piaHeure, piaOutlook], True, True);  //
//end;

procedure TOF_PLANPRESSAL.OnChangeRupture(Sender: TObject);
var
  Combo: THValComboBox;
  Name, Ordre: string;
  IntOrdre, i: integer;
begin
  if Sender is THValComboBox then Combo := THValComboBox(Sender) else exit;
  Name := Combo.name;
  Ordre := Copy(name, Length(name), 1);
  if isnumeric(Ordre) then IntOrdre := StrToInt(Ordre) else IntOrdre := 1;
  if (Combo.value = '') or (Combo.value = 'SO_PGSERVICE') then
  begin
    if combo.value = '' then ChampsRupt[IntOrdre] := '';
    for i := IntOrdre + 1 to 4 do
    begin
      SetControlEnabled('THRUPTURE' + IntToStr(i), False);
      SetControlEnabled('RUPTURE' + IntToStr(i), False);
      ChampsRupt[IntOrdre] := '';
      SetControlText('RUPTURE' + IntToStr(i), '');
    end;
  end
  else
  begin
    SetControlEnabled('THRUPTURE' + IntToStr(IntOrdre + 1), True);
    SetControlEnabled('RUPTURE' + IntToStr(IntOrdre + 1), True);
    ChampsRupt[IntOrdre] := RendNomChampSal(Combo.value);
  end;
end;

Initialization
  registerclasses ( [ TOF_PLANPRESSAL ] ) ; 
end.

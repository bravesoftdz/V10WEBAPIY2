{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 27/03/2002
Modifié le ... :   /  /
Description .. : Edition du solde des jours Rtt
Mots clefs ... : PAIE;RTT;CP
*****************************************************************
PT1 18/04/2002 SB V571 Fiche de bug n°10087 : V_PGI_env.LibDossier non renseigné en Mono
PT2 01/08/2002 SB V585 FQ n° 10192 Intégration des cumuls RTT : contrôle de vraisemblance
PT3 07/10/2002 SB V585 Requête Sql inopérante sur ORACLE 7
PT4 22/10/2002 SB V585 Manque critère de sélection dans requête
PT5 12/12/2002 SB V585 La requête compatible Oracle 7 est erronée..
                       Dans l'attente de la version AGL permettant de conditionner les sommes
                       le requête diffusée reste celle non compatible Oracle7
PT6 15/01/2004 SB V591 FQ10447 Gestion multicombo des paramètres soc RTT Acquis Pris
PT7 02/10/2003 SB V_42 Affichage des ongles si gestion paramsoc des combos libres
PT8 05/05/2004 SB V_50 FQ 11281 Exclusion des salariés sortis
PT9 11/05/2006 SB V_65 FQ 13015 Ajout du date d'arrêté de debut
PT10 02/11/2006 SB V_65 FQ 13361 Contrôle cohérence date d'arrêté
}
unit UTOFPGEditSolde;

interface
uses StdCtrls, Controls, Classes,   sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
  eQRS1,
  {$ELSE}
  //    {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  QRS1,
  {$ENDIF}
  ParamDat, UTOF, HCtrls,
  HEnt1,ParamSoc,HQry,HMsgBox;

Type
       TOF_PGSOLDERTT_ETAT = Class (TOF)
         public
         procedure OnArgument(Arguments : String ) ; override ;
         procedure OnUpdate; override ;
         private
         procedure DateElipsisclick(Sender: TObject);
         procedure ExitEdit(Sender: TObject);
         procedure OnChangeRupture(Sender : TObject);
         function  RendNomChampPaie(st: string): string;
         procedure OnChangeAlpha(Sender : TObject);
       END ;

  implementation

  uses EntPaie,Pgoutils2,PgEditOutils,PGEditOutils2;

  { TOF_PGSOLDERTT_ETAT }
procedure TOF_PGSOLDERTT_ETAT.OnArgument(Arguments: string);
var
  Edit: THEdit;
  Min, Max: string;
  i: integer;
  Combo: THValComboBox;
  Check: TCheckBox;
begin
  inherited;
  {Valeur par défaut et gestionnaire d'évenements}
  VisibiliteChamp(Ecran);
  VisibiliteChampLibre(Ecran);
  { DEB PT7 }
  SetControlProperty('TBCOMPLEMENT', 'Tabvisible', (VH_Paie.PGNbreStatOrg > 0) or (VH_Paie.PGLibCodeStat <> ''));
  SetControlProperty('TBCHAMPLIBRE', 'Tabvisible', (VH_Paie.PgNbCombo > 0));
  { FIN PT7 }
  RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
  Edit := ThEdit(getcontrol('PSA_SALARIE'));
  if Edit <> nil then
  begin
    Edit.text := Min;
    Edit.OnExit := ExitEdit;
  end;
  Edit := ThEdit(getcontrol('PSA_SALARIE_'));
  if Edit <> nil then
  begin
    Edit.text := Max;
    Edit.OnExit := ExitEdit;
  end;
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  Edit := ThEdit(getcontrol('PSA_ETABLISSEMENT'));
  if Edit <> nil then Edit.text := Min;
  Edit := ThEdit(getcontrol('PSA_ETABLISSEMENT_'));
  if Edit <> nil then Edit.text := Max;
  { DEB PT9 }
  Edit := THEdit(GetControl('DATEARRET'));
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
  Edit := THEdit(GetControl('DATEARRET_'));
  if Edit <> nil then
  begin
    Edit.OnDblClick := DateElipsisclick;
    Edit.Text := DateToStr(Date);
  end;
  { FIN PT9 }

  Edit := ThEdit(getcontrol('DOSSIER'));
  if Edit <> nil then
    //  Edit.text:=V_PGI_env.LibDossier;  //PT1 Mise en commentaire
    Edit.text := GetParamSoc('SO_LIBELLE');

  for i := 1 to 4 do
  begin
    Combo := ThValComboBox(GetControl('THVALRUPTURE' + IntToStr(i)));
    if Combo <> nil then combo.onchange := OnChangeRupture;
  end;

  //Evenement ONCHANGE
  Check := TCheckBox(GetControl('CALPHA'));
  if Check <> nil then Check.OnClick := OnChangeAlpha;
  //DEB PT2
  if (VH_Paie.PgRttAcquis = '') or (VH_Paie.PgRttPris = '') then
  begin
    PgiBox('Veuiller renseigner les cumuls de RTT acquis et pris dans les paramètres société.', Ecran.caption);
    SetControlEnabled('BValider', False);
  end;
  //FIN PT2
end;
procedure TOF_PGSOLDERTT_ETAT.OnUpdate;
var
  SQL, Order, St, StAcquis, StPris: string;
  Critere, Champ, DateArret, DateArret2 : string;
  Pages: TPageControl;
begin
  inherited;
  Champ := '';
  Pages := TPageControl(GetControl('Pages'));
  critere := RecupWhereCritere(Pages);

  if GetControlText('XX_RUPTURE1') <> '' then Champ := GetControlText('XX_RUPTURE1') + ',';
  if GetControlText('XX_RUPTURE2') <> '' then Champ := Champ + GetControlText('XX_RUPTURE2') + ',';
  if GetControlText('XX_RUPTURE3') <> '' then Champ := Champ + GetControlText('XX_RUPTURE3') + ',';
  if GetControlText('XX_RUPTURE4') <> '' then Champ := Champ + GetControlText('XX_RUPTURE4') + ',';
  if GetControlText('CALPHA') = 'X' then Order := 'PSA_LIBELLE,PSA_PRENOM,' else Order := '';
  DateArret := UsDateTime(StrToDate(GetControlText('DATEARRET')));
  DateArret2 := UsDateTime(StrToDate(GetControlText('DATEARRET_')));   { PT9 }
  { DEB PT10 }
  if DateArret > DateArret2 then
    Begin
    PgiInfo('Les dates d''arrêté sont incohérentes.', TFQRS1(Ecran).caption);
    SetControlText('DATEARRET',GetControlText('DATEARRET_'));
    End;
  { FIN PT10 }

  {DEB PT6 Ajout d'un traitement multi cumul}
  StAcquis := '';
  StPris := '';
  St := VH_Paie.PgRttAcquis;
  while St <> '' do
    StAcquis := StAcquis + ' PHC_CUMULPAIE="' + ReadTokenSt(St) + '" OR';
  if StAcquis <> '' then StAcquis := ' AND (' + Copy(StAcquis, 1, Length(StAcquis) - 2) + ')';
  St := VH_Paie.PgRttPris;
  while St <> '' do
    StPris := StPris + ' PHC_CUMULPAIE="' + ReadTokenSt(St) + '" OR';
  if StPris <> '' then StPris := ' AND (' + Copy(StPris, 1, Length(StPris) - 2) + ')';
  {FIN PT11}

  { DEB PT8 }
  if pos('WHERE',Critere)>0 then
     Critere := Critere +' AND (PSA_DATESORTIE>="' +DateArret2+ '" OR PSA_DATESORTIE="' + UsdateTime(Idate1900) + '" OR PSA_DATESORTIE is null) ' +
    'AND PSA_DATEENTREE<="' +DateArret2+ '" '    { PT9 }
  else
     Critere := Critere +' WHERE (PSA_DATESORTIE>="' + DateArret2+ '" OR PSA_DATESORTIE="' + UsdateTime(Idate1900) + '" OR PSA_DATESORTIE is null) ' +
    'AND PSA_DATEENTREE<="' +DateArret2+ '" ';    { PT9 }
  { FIN PT8 }



  {PT3 Modification de la requête..plante sous Oracle 7
  DEB PT5 Reprise de l'ancienne requête dont le résultat est exact mais non compatible Oracle 7}
  Sql := 'SELECT DISTINCT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_DATEENTREE,' + Champ +
    '(SELECT SUM(PHC_MONTANT) FROM HISTOCUMSAL WHERE PHC_DATEFIN>="' + DateArret + '" AND PHC_DATEFIN<="' + DateArret2 + '" ' + { PT9 }
    'AND PHC_SALARIE=PSA_SALARIE ' + StAcquis + ') ACQUIS,' + //PT2 //PT6 Modif requête
  '(SELECT SUM(PHC_MONTANT) FROM HISTOCUMSAL WHERE PHC_DATEFIN>="' + DateArret + '" AND PHC_DATEFIN<="' + DateArret2 + '" ' +   { PT9 }
    'AND PHC_SALARIE=PSA_SALARIE ' + StPris + ') PRIS ' + //PT2 //PT6 Modif requête
  'FROM SALARIES ' +
    critere +
    'ORDER BY ' + Champ + Order + 'PSA_SALARIE ';
  {Sql:='SELECT DISTINCT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_DATEENTREE,'+Champ+
       'SUM (H1.PHC_MONTANT) ACQUIS,SUM (H2.PHC_MONTANT) PRIS '+
       'FROM SALARIES '+
       'LEFT JOIN HISTOCUMSAL H1 ON H1.PHC_SALARIE=PSA_SALARIE '+
       'AND H1.PHC_CUMULPAIE="'+VH_Paie.PgRttAcquis+'" AND H1.PHC_DATEFIN<="'+DateArret+'" '+  //PT4
       'LEFT JOIN HISTOCUMSAL H2 ON H2.PHC_SALARIE=PSA_SALARIE '+
       'AND H2.PHC_CUMULPAIE="'+VH_Paie.PgRttPris+'" AND H2.PHC_DATEFIN<="'+DateArret+'" '+    //PT4
        critere+  //PT4
       'GROUP BY PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,'+Champ+'PSA_DATEENTREE '+
       'ORDER BY '+Champ+Order+'PSA_SALARIE ';
  FIN PT5}
  TFQRS1(Ecran).WhereSQL := SQL;

end;

procedure TOF_PGSOLDERTT_ETAT.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGSOLDERTT_ETAT.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;



procedure TOF_PGSOLDERTT_ETAT.OnChangeRupture(Sender: TObject);
var
  Combo: THValComboBox;
  ChampRupture, Name, Ordre: string;
  IntOrdre, i: integer;
begin
  if Sender is THValComboBox then Combo := THValComboBox(Sender) else combo := nil;
  if Combo = nil then exit;
  Name := Combo.name;
  Ordre := Copy(name, Length(name), 1);
  IntOrdre := StrToInt(Ordre) + 1;
  ChampRupture := 'XX_RUPTURE' + Ordre;
  if Combo.value = '' then
  begin
    SetControlText(ChampRupture, '');
    for i := IntOrdre to 4 do
    begin
      SetControlText('XX_RUPTURE' + IntToStr(i), '');
      SetControlText('THVALRUPTURE' + IntToStr(i), '');
      SetControlEnabled('THVALRUPTURE' + IntToStr(i), False);
    end;
  end
  else
  begin
    SetControlEnabled('THVALRUPTURE' + IntToStr(IntOrdre), True);
    SetControlText(ChampRupture, RendNomChampPaie(Combo.value));
  end;
end;

function TOF_PGSOLDERTT_ETAT.RendNomChampPaie(st: string): string;
begin
  Result := '';
  if st = 'SO_PGETABLISSEMENT' then result := 'PSA_ETABLISSEMENT'
  else
    if st = 'SO_PGLIBCODESTAT' then result := 'PSA_CODESTAT'
  else
    if Pos('SO_PGLIBORGSTAT', St) > 0 then result := 'PSA_TRAVAILN' + Copy(St, Length(St), 1)
  else
    if Pos('SO_PGLIBCOMBO', St) > 0 then result := 'PSA_LIBREPCMB' + Copy(St, Length(St), 1);

end;


procedure TOF_PGSOLDERTT_ETAT.OnChangeAlpha(Sender: TObject);
begin
  if TCheckBox(Sender).Name = 'CALPHA' then
    AffectCritereAlpha(Ecran, TCheckBox(Sender).Checked, 'PSA_SALARIE', 'PSA_LIBELLE');
end;

initialization
  registerclasses([TOF_PGSOLDERTT_ETAT]);
end.


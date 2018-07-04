{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 31/01/2002
Modifié le ... :   /  /
Description .. : Planning des absences
Mots clefs ... : PAIE;ABSENCE;CP
*****************************************************************}
{
PT1 05/12/2002 V591 SB Gestion paramétre soc de la hiérarchie dans la planning
PT2 07/11/2003 V_42 SB Intégration ou non des salariés sortis
PT3 17/05/2004 V_50 SB FQ 11287 Collezerodevant pour matricule responsable
PT4 17/05/2004 V_50 SB FQ 11304 Modif. zone Type absence et état en multicombo
PT5 19/09/2005 V_65 SB FQ 11313 Ajout planning des absences de la paie, Refonte du source
PT6 15/09/2005 V_65 SB Econges : recherche responsable table intérimaire
PT7 07/10/2005 V_65 SB Planning absence : suppression des <<tous>>
PT8 05/12/2005 V_65 SB FQ 12737 Ajout sélection du niveau hierarchique
PT9 28/02/2006 SB V_65 FQ 12926 Optimisation du planning
PT10 09/05/2006 SB V_65 FQ 13038 gestion des salariés sortis sur option
PT11 27/06/2006 SB V_65 Nouveaux filtres d'absence
PT12 06/07/2006 SB V_65 Nouveaux paramètres d'execution
PT13 24/01/2007 FC V_80 Gestion du filtrage habilitations
}

unit UTOFPGPlaning;

interface
uses SysUtils, Classes, Controls, Dialogs,
  stdctrls,
{$IFDEF EAGLCLIENT}

{$ELSE}
{$ENDIF}
  HPanel, HCtrls, Utof, P5Def,
  LookUp, entpaie,
  PgOutilsEAgl, HEnt1, PGPlanning, PGPlanningAbsPaie, PgOutils2, Hmsgbox;

type
  TOF_PGPLANNINGMERE = class(TOF)
  public
    ChampsRupt: array[1..4] of string;
    procedure OnChangeRupture(Sender: TObject);
    procedure OnExitSalarie(Sender: TObject);
    function RendNomChampSal(st: string): string;
    function RecupCritere(ARecuperer: string): string;
    function RecupMultiCritereCombo(Champ: string): string;
    function RendSyntaxeCondSqlIn(Cond, Champ, Val: string): string; { PT11 }
  end;

  TOF_PGPLANNING = class(TOF_PGPLANNINGMERE)
    PPanel: THPanel;
    procedure OnArgument(stArgument: string); override;
    procedure OnLoad; override;
    procedure OnClose; override;
    procedure OnUpdate; override;
  private
    procedure AfficheChampHierarchie(Afficher: boolean); //PT1
  protected
    procedure AfficheTabletteSal(Sender: TObject);
    procedure AfficheTabletteResp(Sender: TObject);
//       procedure GestionColors(Sender: TObject);
  //     procedure GestionFont(Sender: TObject);
    procedure ExitRespons(Sender: TObject);
    procedure ClickHierarchie(Sender: TObject); { PT8 }
    Procedure ChangeSoldeAbs(Sender: TObject);  { PT12 }
  public
        { Déclarations publiques }
    Colors: TColorDialog;
    LaFont: TFontDialog;
    Utilisat, Salarie, Mode, GBDatedeb : string;
  end;

  TOF_PGPLANNINGABSPAIE = class(TOF_PGPLANNINGMERE)
    procedure OnArgument(stArgument: string); override;
    procedure OnUpdate; override;
  end;

implementation
uses PGPlanningOutils, P5Util, PGHierarchie;


{ TOF_PGPLANNINGMERE }

procedure TOF_PGPLANNINGMERE.OnChangeRupture(Sender: TObject);
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

procedure TOF_PGPLANNINGMERE.OnExitSalarie(Sender: TObject);
var edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

function TOF_PGPLANNINGMERE.RecupCritere(ARecuperer: string): string;
var
  StWhere, StAbs, St, StVal : string;
  Hierar: THMultiValComboBox;
begin
  result := '';
  if ARecuperer = 'SALARIE' then
  begin
    StWhere := '';
{ if Utilisat='ADM'  GetControlText('RUPTURE1')<>'SO_PGSERVICE' then
  if (GetControlText('RESPONS')<>'') AND ((GetControlText('PSA_SALARIE')<>'') or (GetControlText('PSA_SALARIE_')<>''))then
    StWhere:=StWhere+'(PSE_RESPONSABS="'+GetControlText('RESPONS')+'" AND ('
  else
    if (GetControlText('RESPONS')<>'') AND (GetControlText('PSA_SALARIE')='') AND (GetControlText('PSA_SALARIE_')='')then
      StWhere:=StWhere+'PSE_RESPONSABS="'+GetControlText('RESPONS')+'" ';
                    }
    if (GetControlText('PSA_SALARIE') <> '') and (GetControlText('PSA_SALARIE_') <> '') then
      StWhere := StWhere + 'PSA_SALARIE>="' + GetControlText('PSA_SALARIE') + '" AND PSA_SALARIE<="' + GetControlText('PSA_SALARIE_') + '" '
    else
      if (GetControlText('PSA_SALARIE') <> '') and (GetControlText('PSA_SALARIE_') = '') then
        StWhere := StWhere + 'PSA_SALARIE>="' + GetControlText('PSA_SALARIE') + '" '
      else
        if (GetControlText('PSA_SALARIE') = '') and (GetControlText('PSA_SALARIE_') <> '') then
          StWhere := StWhere + 'PSA_SALARIE<="' + GetControlText('PSA_SALARIE_') + '" ';
    //DEB PT13
    if Assigned(MonHabilitation) then
      if MonHabilitation.Active then
        if StWhere = '' then
          StWhere := MonHabilitation.LeSQL
        else
          StWhere := StWhere + ' AND ' + MonHabilitation.LeSQL;
    //FIN PT13
    if (Pos('PSE_RESPONSABS', StWhere) > 0) and (Pos('PSA_SALARIE', StWhere) > 0) then
      StWhere := StWhere + ')) AND'
    else
      if Stwhere <> '' then StWhere := StWhere + 'AND ';

    if (GetControlText('PSA_ETABLISSEMENT') <> '') and (Pos('<Tous', GetControlText('PSA_ETABLISSEMENT')) < 1) then
      StWhere := StWhere + RecupMultiCritereCombo('PSA_ETABLISSEMENT') + ' AND '; //PSA_ETABLISSEMENT="'+GetControlText('PSA_ETABLISSEMENT')+'"
    if (GetControlText('PSA_TRAVAILN1') <> '') and (Pos('<Tous', GetControlText('PSA_TRAVAILN1')) < 1) then
      StWhere := StWhere + RecupMultiCritereCombo('PSA_TRAVAILN1') + ' AND ';
    if (GetControlText('PSA_TRAVAILN2') <> '') and (Pos('<Tous', GetControlText('PSA_TRAVAILN2')) < 1) then
      StWhere := StWhere + RecupMultiCritereCombo('PSA_TRAVAILN2') + ' AND ';
    if (GetControlText('PSA_TRAVAILN3') <> '') and (Pos('<Tous', GetControlText('PSA_TRAVAILN3')) < 1) then
      StWhere := StWhere + RecupMultiCritereCombo('PSA_TRAVAILN3') + ' AND ';
    if (GetControlText('PSA_TRAVAILN4') <> '') and (Pos('<Tous', GetControlText('PSA_TRAVAILN4')) < 1) then
      StWhere := StWhere + RecupMultiCritereCombo('PSA_TRAVAILN4') + ' AND ';
    if (GetControlText('PSA_CODESTAT') <> '') and (Pos('<Tous', GetControlText('PSA_CODESTAT')) < 1) then
      StWhere := StWhere + RecupMultiCritereCombo('PSA_CODESTAT') + ' AND ';
    if (GetControlText('PSA_LIBREPCMB1') <> '') and (Pos('<Tous', GetControlText('PSA_LIBREPCMB1')) < 1) then
      StWhere := StWhere + RecupMultiCritereCombo('PSA_LIBREPCMB1') + ' AND ';
    if (GetControlText('PSA_LIBREPCMB2') <> '') and (Pos('<Tous', GetControlText('PSA_LIBREPCMB2')) < 1) then
      StWhere := StWhere + RecupMultiCritereCombo('PSA_LIBREPCMB2') + ' AND ';
    if (GetControlText('PSA_LIBREPCMB3') <> '') and (Pos('<Tous', GetControlText('PSA_LIBREPCMB3')) < 1) then
      StWhere := StWhere + RecupMultiCritereCombo('PSA_LIBREPCMB3') + ' AND ';
    if (GetControlText('PSA_LIBREPCMB4') <> '') and (Pos('<Tous', GetControlText('PSA_LIBREPCMB4')) < 1) then
      StWhere := StWhere + RecupMultiCritereCombo('PSA_LIBREPCMB4') + ' AND ';
    if VH_Paie.PGEcabHierarchie then //PT1
      if (GetControlText('PSE_CODESERVICE') <> '') and (Pos('<Tous', GetControlText('PSE_CODESERVICE')) < 1) then
        StWhere := StWhere + RecupMultiCritereCombo('PSE_CODESERVICE') + ' AND ';
  //Ajout clause date sortie
  { DEB PT10 refonte de la clause }
    if GetControlText('CKSORTIE') = '-' then ////PT2
      StWhere := StWhere + ' (PSA_DATESORTIE is null ' +
        'OR PSA_DATESORTIE<="' + UsDateTime(idate1900) + '") '
    else
      StWhere := StWhere + ' (PSA_DATESORTIE is null ' +
        'OR PSA_DATESORTIE<="' + UsDateTime(idate1900) + '" ' +
        'OR (PSA_DATESORTIE>="' + UsDatetime(StrToDate(GetControlText('DATEABS'))) + '" ' +
        'AND PSA_DATESORTIE<="' + UsDatetime(StrToDate(GetControlText('DATEABS_'))) + '"  ))';
  { FIN PT10 }

//  if Length(StWhere)<10 then StWhere:=''
//  else StWhere:=Copy(StWhere,1,Length(StWhere)-5)+' ';
  {if (Utilisat='SAL') AND (StWhere<>'') then
    StWhere:='WHERE ('+StWhere+' OR PSE_RESPONSABS =(SELECT PSE_RESPONSABS FROM DEPORTSAL '+
    'WHERE PSE_SALARIE="'+GetControlText('PSA_SALARIE')+'"))';     }
    if Pos('WHERE', StWhere) = 0 then StWhere := 'WHERE ' + StWhere;
  end
  else
    if ARecuperer = 'ABSENCE' then
    begin { DEB PT4 }
      StWhere := ''; StAbs := '';
      if (GetControlText('PMA_MOTIFABSENCE') <> '') and (GetControlText('CBFILTREABS') <> 'SA1') and (GetControlText('CBFILTREABS') <> 'SA2') then { PT11 }
      begin
        St := GetControlText('PMA_MOTIFABSENCE');
        while (St <> '') and (st <> '<<Tous>>') do { PT7 10/10/2005}
        begin
          StVal := ReadTokenSt(St);
          StAbs := StAbs + ' PCN_TYPECONGE="' + StVal + '" OR';
        end;
        if StAbs <> '' then StAbs := ' AND (' + Copy(StAbs, 1, Length(StAbs) - 3) + ') ';
        StWhere := StAbs;
      end;
      if GetControlText('PCN_VALIDRESP') <> '' then
      begin
        StAbs := '';
        St := GetControlText('PCN_VALIDRESP');
        while (St <> '') and (st <> '<<Tous>>') do { PT7 10/10/2005}
        begin
          StVal := ReadTokenSt(St);
          StAbs := StAbs + ' PCN_VALIDRESP="' + StVal + '" OR';
        end;
        if StAbs <> '' then StAbs := 'AND (' + Copy(StAbs, 1, Length(StAbs) - 3) + ') ';
        StWhere := StWhere + StAbs;
      end; { FIN PT4 }
      if Assigned(THMultiValComboBox(GetControl('PCN_CODETAPE'))) then
        if GetControlText('PCN_CODETAPE') <> '' then
        begin
          StAbs := '';
          St := GetControlText('PCN_CODETAPE');
          while (St <> '') and (st <> '<<Tous>>') do { PT7 }
          begin
            StVal := ReadTokenSt(St);
            StAbs := StAbs + ' PCN_CODETAPE="' + StVal + '" OR';
          end;
          if StAbs <> '' then StAbs := 'AND (' + Copy(StAbs, 1, Length(StAbs) - 3) + ') ';
          StWhere := StWhere + StAbs;
        end; 
    end
    else
      if ARecuperer = 'CBFILTREABS' then
      begin
        StWhere := RendSyntaxeCondSqlIn('IN', 'PMA_MOTIFABSENCE', GetControlText('PMA_MOTIFABSENCE'))
      end
      else
        if ARecuperer = 'FILTRE' then { DEB PT7 }
        begin
          st := GetControlText('CBFILTREABS');
          if St <> '<<Tous>>' then
            StWhere := StWhere + GetControlText('CBFILTREABS');
        end { FIN PT7 }
        else
          if ARecuperer = 'HIERARCHIE' then { DEB PT8 }
          begin
            Hierar := THMultiValComboBox(GetControl('HIERARCHIE'));
            if Assigned(Hierar) then StWhere := RendClauseHierarchie(Hierar.Text);
          end; { FIN PT8 }

  result := StWhere;
end;

function TOF_PGPLANNINGMERE.RecupMultiCritereCombo(Champ: string): string;
var
  StValue: string;
begin
  result := '';
  StValue := GetControlText(Champ);
  result := '(';
  while StValue <> '' do
    Result := Result + Champ + '="' + ReadTokenSt(StValue) + '" OR ';
  if Length(result) > 10 then Result := Copy(result, 1, Length(result) - 4) + ')'
  else result := '';
end;

function TOF_PGPLANNINGMERE.RendNomChampSal(st: string): string;
begin
  Result := '';
  if st = 'SO_PGLIBCODESTAT' then result := 'PSA_CODESTAT'
  else
    if Pos('SO_PGLIBORGSTAT', St) > 0 then result := 'PSA_TRAVAILN' + Copy(St, Length(St), 1)
    else
      if Pos('SO_PGLIBCOMBO', St) > 0 then result := 'PSA_LIBREPCMB' + Copy(St, Length(St), 1)
      else
        if Pos('SO_PGETABLISSEMENT', St) > 0 then result := 'PSA_ETABLISSEMENT'
        else
          if Pos('SO_PGRESPONSABS', St) > 0 then result := 'PSE_RESPONSABS';
end;

{ DEB PT11 }
function TOF_PGPLANNINGMERE.RendSyntaxeCondSqlIn(Cond, Champ, Val: string): string;
var St: string;
begin
  result := '';
  if (Trim(Champ) = '') then exit;
  St := ReadTokenSt(Val);
  while St <> '' do
  begin
    Result := Result + '"' + St + '",';
    St := ReadTokenSt(Val);
  end;
  if Result <> '' then
    Result := ' AND ' + Champ + ' ' + Cond + ' (' + Copy(Result, 1, Length(Result) - 1) + ')'
  else
    Result := ' AND ' + Champ + ' <> "" ';
end;
{ FIN PT11 }

{ TOF_PGPLANNING }

procedure TOF_PGPLANNING.OnArgument(stArgument: string);
var
  Num: integer;
  Edit: THEdit;
  NumEdit: THNumEdit;
  Combo: TControl;
  GrpBox: TGroupBox;
  ChbxH: TCheckBox;
begin
  inherited;
  for num := 1 to 4 do ChampsRupt[num] := '';
  Utilisat := ReadtokenSt(StArgument);
  Salarie := ReadtokenSt(StArgument);
{DEB PT12}
  Mode := ReadtokenSt(StArgument);
  if Mode = 'SUIVI' then
    Begin
    Ecran.Caption := 'Suivi des absences';
    UpdateCaption(Ecran);
    End;
{FIN PT12}
  SetControlEnabled('RESPONS', ((Utilisat = '') or (Utilisat = 'ADM')));
  SetControlEnabled('PSA_SALARIE', (Utilisat <> 'SAL'));
  SetControlEnabled('PSA_SALARIE_', (Utilisat <> 'SAL'));
  SetControlvisible('CKSORTIE', (Utilisat = 'ADM')); //PT2
  if Utilisat = 'RESP' then
    SetControlText('RESPONS', Salarie)
  else
    if Utilisat = 'SAL' then
    begin
//    SetControlText('PSA_SALARIE',Salarie);
   // SetControlText('PSA_SALARIE_',Salarie);
      SetControlEnabled('PMA_MOTIFABSENCE', False);
      SetControlText('RUPTURE1', 'SO_PGSERVICE');
      SetControlText('RUPTURE2', '');
      SetControlText('RUPTURE3', '');
      SetControlText('RUPTURE4', '');
//    SetControlChecked('CKORGANIGRAMME',True);
//    SetControlVisible('CKORGANIGRAMME',False);
    end;
  if (Utilisat <> 'SAL') then
  begin
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
      if (Num = 1) then SetControlText('TPSA_LIBREPCMB1', VH_Paie.PgLibCombo1)
      else
        if (Num = 2) then SetControlText('TPSA_LIBREPCMB2', VH_Paie.PgLibCombo2)
        else
          if (Num = 3) then SetControlText('TPSA_LIBREPCMB3', VH_Paie.PgLibCombo3)
          else
            if (Num = 4) then SetControlText('TPSA_LIBREPCMB4', VH_Paie.PgLibCombo4);
      Num := Num + 1;
    end;
 { Combo:=ThValComboBox(getcontrol('PSA_TRAVAILN1'));
  InitialiseCombo(Combo);
  Combo:=ThValComboBox(getcontrol('PSA_TRAVAILN2'));
  InitialiseCombo(Combo);
  Combo:=ThValComboBox(getcontrol('PSA_TRAVAILN3'));
  InitialiseCombo(Combo);
  Combo:=ThValComboBox(getcontrol('PSA_TRAVAILN4'));
  InitialiseCombo(Combo);
  Combo:=ThValComboBox(getcontrol('PSA_CODESTAT'));
  InitialiseCombo(Combo);
  Combo:=ThValComboBox(getcontrol('PSA_LIBREPCMB1'));
  InitialiseCombo(Combo);
  Combo:=ThValComboBox(getcontrol('PSA_LIBREPCMB2'));
  InitialiseCombo(Combo);
  Combo:=ThValComboBox(getcontrol('PSA_LIBREPCMB3'));
  InitialiseCombo(Combo);
  Combo:=ThValComboBox(getcontrol('PSA_LIBREPCMB4'));
  InitialiseCombo(Combo);}
  end;

  SetControlText('NOMRESP', RechDom('PGSALARIE', GetControlText('RESPONS'), False));
  if GetControlText('NOMRESP') = 'Error' then SetControlText('NOMRESP', '');
  Edit := THEdit(GetControl('RESPONS'));
  if Edit <> nil then begin Edit.OnElipsisClick := AfficheTabletteResp; Edit.OnExit := ExitRespons; end;
  Edit := THEdit(GetControl('PSA_SALARIE'));
  if Edit <> nil then begin Edit.OnElipsisClick := AfficheTabletteSal; Edit.OnExit := OnExitSalarie; end;
  Edit := THEdit(GetControl('PSA_SALARIE_'));
  if Edit <> nil then begin Edit.OnElipsisClick := AfficheTabletteSal; Edit.OnExit := OnExitSalarie; end;
//Combo:=ThValComboBox(getcontrol('PMA_MOTIFABSENCE')); PT4 Mise en commentaire
//InitialiseCombo(Combo);
{Combo:=ThValComboBox(getcontrol('PSA_ETABLISSEMENT'));
InitialiseCombo(Combo);   }

  SetControlEnabled('PSA_ETABLISSEMENT', (Utilisat <> 'SAL'));
  SetControlEnabled('PSA_TRAVAILN1', (Utilisat <> 'SAL'));
  SetControlEnabled('PSA_TRAVAILN2', (Utilisat <> 'SAL'));
  SetControlEnabled('PSA_TRAVAILN3', (Utilisat <> 'SAL'));
  SetControlEnabled('PSA_TRAVAILN4', (Utilisat <> 'SAL'));
  SetControlEnabled('PSA_CODESTAT', (Utilisat <> 'SAL'));

  for Num := 1 to 4 do
  begin
    Combo := ThValComboBox(GetControl('RUPTURE' + IntToStr(num)));
    if Combo <> nil then THValComboBox(Combo).OnChange := OnChangeRupture;
    InitialiseCombo(Combo);
  end;

{Check := TRadioButton(GetControl('CKORGANIGRAMME'));
if check<>nil then check.onclick:=OnClickOrganigramme;
Check := TRadioButton(GetControl('CKORGANISATION'));
if check<>nil then check.onclick:=OnClickOrganisation;}

{PT9 Mise en commentaire
SetControlProperty('RUPTURE1','Plus','(SOC_NOM LIKE "SO_PGLIBORGSTAT%" OR SOC_NOM LIKE "SO_PGLIBCOMBO%" OR SOC_NOM="SO_PGLIBCODESTAT" OR SOC_NOM="SO_PGETABLISSEMENT" OR SOC_NOM="SO_PGRESPONSABS") AND SOC_DATA<>""');
SetControlProperty('RUPTURE2','Plus','(SOC_NOM LIKE "SO_PGLIBORGSTAT%" OR SOC_NOM LIKE "SO_PGLIBCOMBO%" OR SOC_NOM="SO_PGLIBCODESTAT" OR SOC_NOM="SO_PGETABLISSEMENT" OR SOC_NOM="SO_PGRESPONSABS") AND SOC_DATA<>""');
SetControlProperty('RUPTURE3','Plus','(SOC_NOM LIKE "SO_PGLIBORGSTAT%" OR SOC_NOM LIKE "SO_PGLIBCOMBO%" OR SOC_NOM="SO_PGLIBCODESTAT" OR SOC_NOM="SO_PGETABLISSEMENT" OR SOC_NOM="SO_PGRESPONSABS") AND SOC_DATA<>""');
SetControlProperty('RUPTURE4','Plus','(SOC_NOM LIKE "SO_PGLIBORGSTAT%" OR SOC_NOM LIKE "SO_PGLIBCOMBO%" OR SOC_NOM="SO_PGLIBCODESTAT" OR SOC_NOM="SO_PGETABLISSEMENT" OR SOC_NOM="SO_PGRESPONSABS") AND SOC_DATA<>""');}

  ChbxH := TCheckBox(GetControl('MULTINIVEAU'));
  if Assigned(ChbxH) then ChbxH.OnClick := ClickHierarchie; { PT8 }

  NumEdit := THNumEdit(GetControl('SOLDEABS'));                 { PT12 }
  if Assigned(NumEdit) then NumEdit.OnChange := ChangeSoldeAbs; { PT12 }

end;

procedure TOF_PGPLANNING.OnLoad;
begin
  inherited;
  AfficheChampHierarchie(VH_PAIE.PGEcabHierarchie); //PT1
end;

procedure TOF_PGPLANNING.OnClose;
begin
  inherited;
  if Colors <> nil then Colors.free;
  if LaFont <> nil then LaFont.free;
end;





procedure TOF_PGPLANNING.OnUpdate;
var NiveauRupt: TNiveauRupture;
begin
  inherited;

  NiveauRupt.ChampsRupt[1] := ChampsRupt[1]; NiveauRupt.ChampsRupt[2] := ChampsRupt[2];
  NiveauRupt.ChampsRupt[3] := ChampsRupt[3]; NiveauRupt.ChampsRupt[4] := ChampsRupt[4];
  NiveauRupt.CondRupt := ''; NiveauRupt.NiveauRupt := 0;
  if GetControlText('RUPTURE1') <> '' then NiveauRupt.NiveauRupt := 1;
  if GetControlText('RUPTURE2') <> '' then NiveauRupt.NiveauRupt := 2;
  if GetControlText('RUPTURE3') <> '' then NiveauRupt.NiveauRupt := 3;
  if GetControlText('RUPTURE4') <> '' then NiveauRupt.NiveauRupt := 4;
  if UtiliSat = 'ADM' then Salarie := GetControlText('RESPONS');
  PGPlanningAbsence(StrToDate(GetControlText('DATEABS')), StrToDate(GetControlText('DATEABS_')),
    Utilisat, '', RecupCritere('ABSENCE'), RecupCritere('SALARIE'), Salarie, NiveauRupt, (GetControlText('MULTINIVEAU') = 'X'),
    RecupCritere('FILTRE'), RecupCritere('HIERARCHIE'), RecupCritere('CBFILTREABS'), { PT8 } { PT11 }
    StrtoFloat(GetControltext('JOURABS')),StrtoFloat(GetControltext('SOLDEABS')), { PT12 }
    GetControltext('TYPESOLDEABS'),mode);   { PT12 }
  //(GetControlText('RUPTURE1')='SO_PGSERVICE')

end;


//Function TOF_PGPLANNING.RecupCritere(ARecuperer: string) : String;
//end;

procedure TOF_PGPLANNING.AfficheTabletteResp(Sender: TObject);
begin
  if sender = nil then exit;
{ DEB PT6 }
  if GblTypeSal = 'SAL' then
    LookupList(TControl(Sender), 'Responsable', 'DEPORTSAL,SALARIES', 'DISTINCT PSE_RESPONSABS', 'PSA_LIBELLE+"  "+PSA_PRENOM', 'PSE_RESPONSABS=PSA_SALARIE', 'PSE_RESPONSABS', True, -1)
  else
    if GblTypeSal = 'INT' then
      LookupList(TControl(Sender), 'Responsable', 'DEPORTSAL,INTERIMAIRES', 'DISTINCT PSE_RESPONSABS', 'PSI_LIBELLE+"  "+PSI_PRENOM', 'PSE_RESPONSABS=PSI_INTERIMAIRE', 'PSE_RESPONSABS', True, -1)
{ FIN PT6 }
end;

procedure TOF_PGPLANNING.AfficheTabletteSal(Sender: TObject);
begin
  AfficheTabSalResp(Sender, GetControltext('RESPONS'));
end;


procedure TOF_PGPLANNING.ExitRespons(Sender: TObject);
var edit: thedit;
begin
  edit := THEdit(Sender); { PT3 }
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
  SetControlText('NOMRESP', RechDom('PGSALARIE', GetControlText('RESPONS'), False));
end;



//function TOF_PGPLANNING.RecupMultiCritereCombo(Champ: string): string;
//end;

//DEB PT1

procedure TOF_PGPLANNING.AfficheChampHierarchie(Afficher: boolean);
begin
  SetControlVisible('TPSE_CODESERVICE', Afficher);
  SetControlVisible('PSE_CODESERVICE', Afficher);
  SetControlVisible('MULTINIVEAU', Afficher);
  SetControlVisible('GBRUPTURE', Afficher);
  SetControlVisible('HIERARCHIE', Afficher); { PT8 }
  SetControlVisible('HHIERARCHIE', Afficher); { PT8 }
end;
//FIN PT1
{ DEB PT8 }

procedure TOF_PGPLANNING.ClickHierarchie(Sender: TObject);
begin
  SetcontrolEnabled('HIERARCHIE', (TCheckBox(Sender).Checked = True));
  if TCheckBox(Sender).Checked = False then SetcontrolText('HIERARCHIE', '');
end;
{ FIN PT8 }
{ DEB PT12 }
procedure TOF_PGPLANNING.ChangeSoldeAbs(Sender: TObject);
begin
if isnumeric(GetControlText('SOLDEABS')) then
   Begin
   if StrToFloat(GetControlText('SOLDEABS')) <> 0 then
     Begin
     GBDatedeb := GetControlText('DATEABS');
     SetControlEnabled('TYPESOLDEABS',True);
     SetControlText('TYPESOLDEABS','TOUS');
     if (VH_PAIE.PGEcabDateIntegration > idate1900) And  (StrToDate(GBDatedeb) > VH_PAIE.PGEcabDateIntegration) then
       Begin
       PGiInfo('Le début de périodicité d''absence ne doit pas être postérieur à la date de calcul des soldes.',Ecran.caption);
       SetControlText('DATEABS',DateToStr(VH_PAIE.PGEcabDateIntegration));
       End
     end
   else
     Begin
     SetControlEnabled('TYPESOLDEABS',False);
     SetControlEnabled('DATEABS',True);
     SetControlText('DATEABS',GBDatedeb); 
     End;
   End;
end;
{ FIN PT12 }

{ TOF_PGPLANNINGABSPAIE }

procedure TOF_PGPLANNINGABSPAIE.OnArgument(stArgument: string);
var
  Num: integer;
  Edit: THEdit;
  Combo: TControl;
  GrpBox: TGroupBox;
begin
  inherited;
  for num := 1 to 4 do ChampsRupt[num] := '';

  SetControlvisible('CKSORTIE', True);
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
    if (Num = 1) then SetControlText('TPSA_LIBREPCMB1', VH_Paie.PgLibCombo1)
    else
      if (Num = 2) then SetControlText('TPSA_LIBREPCMB2', VH_Paie.PgLibCombo2)
      else
        if (Num = 3) then SetControlText('TPSA_LIBREPCMB3', VH_Paie.PgLibCombo3)
        else
          if (Num = 4) then SetControlText('TPSA_LIBREPCMB4', VH_Paie.PgLibCombo4);
    Num := Num + 1;
  end;

  Edit := THEdit(GetControl('PSA_SALARIE'));
  if Edit <> nil then Edit.OnExit := OnExitSalarie;
  Edit := THEdit(GetControl('PSA_SALARIE_'));
  if Edit <> nil then Edit.OnExit := OnExitSalarie;

  for Num := 1 to 4 do
  begin
    Combo := ThValComboBox(GetControl('RUPTURE' + IntToStr(num)));
    if Combo <> nil then THValComboBox(Combo).OnChange := OnChangeRupture;
    InitialiseCombo(Combo);
  end;

{ PT9 Mise en commentaire
SetControlProperty('RUPTURE1','Plus','(SOC_NOM LIKE "SO_PGLIBORGSTAT%" OR SOC_NOM LIKE "SO_PGLIBCOMBO%" OR SOC_NOM="SO_PGLIBCODESTAT" OR SOC_NOM="SO_PGETABLISSEMENT") AND SOC_DATA<>""');
SetControlProperty('RUPTURE2','Plus','(SOC_NOM LIKE "SO_PGLIBORGSTAT%" OR SOC_NOM LIKE "SO_PGLIBCOMBO%" OR SOC_NOM="SO_PGLIBCODESTAT" OR SOC_NOM="SO_PGETABLISSEMENT") AND SOC_DATA<>""');
SetControlProperty('RUPTURE3','Plus','(SOC_NOM LIKE "SO_PGLIBORGSTAT%" OR SOC_NOM LIKE "SO_PGLIBCOMBO%" OR SOC_NOM="SO_PGLIBCODESTAT" OR SOC_NOM="SO_PGETABLISSEMENT") AND SOC_DATA<>""');
SetControlProperty('RUPTURE4','Plus','(SOC_NOM LIKE "SO_PGLIBORGSTAT%" OR SOC_NOM LIKE "SO_PGLIBCOMBO%" OR SOC_NOM="SO_PGLIBCODESTAT" OR SOC_NOM="SO_PGETABLISSEMENT") AND SOC_DATA<>""');}

end;



procedure TOF_PGPLANNINGABSPAIE.OnUpdate;
var NiveauRupt: TNiveauRupture;
begin
  inherited;
  NiveauRupt.ChampsRupt[1] := ChampsRupt[1]; NiveauRupt.ChampsRupt[2] := ChampsRupt[2];
  NiveauRupt.ChampsRupt[3] := ChampsRupt[3]; NiveauRupt.ChampsRupt[4] := ChampsRupt[4];
  NiveauRupt.CondRupt := ''; NiveauRupt.NiveauRupt := 0;
  if GetControlText('RUPTURE1') <> '' then NiveauRupt.NiveauRupt := 1;
  if GetControlText('RUPTURE2') <> '' then NiveauRupt.NiveauRupt := 2;
  if GetControlText('RUPTURE3') <> '' then NiveauRupt.NiveauRupt := 3;
  if GetControlText('RUPTURE4') <> '' then NiveauRupt.NiveauRupt := 4;
  PGPlanningAbsencePaie(StrToDate(GetControlText('DATEABS')), StrToDate(GetControlText('DATEABS_')),
    'ADM', '', RecupCritere('ABSENCE'), RecupCritere('SALARIE'), NiveauRupt, RecupCritere('FILTRE'));
end;






{procedure TOF_PGPLANNING.GestionColors(Sender: TObject);
begin
Colors:=TColorDialog.Create(Application);
if Colors=nil then exit;
if (Colors.Execute) AND (H<>nil) then
  Begin
  H.ColorBackground:=Colors.Color;
  ColorBackground:=ColorToString(Colors.Color);
  H.Raffraichir;
  End;
end;

procedure TOF_PGPLANNING.GestionFont(Sender: TObject);
begin
LaFont:=TFontDialog.Create(Application);
if LaFont=nil then exit;
LaFont.Font.Color:=StringToColor(FontColor);
If Pos('G',Fontstyle)>0 then LaFont.Font.Style:=[fsBold];
If Pos('I',Fontstyle)>0 then LaFont.Font.Style:=[fsItalic];
If Pos('U',Fontstyle)>0 then LaFont.Font.Style:=[fsUnderline];
If Pos('S',Fontstyle)>0 then LaFont.Font.Style:=[fsStrikeout];
LaFont.Font.Name:=FontName;
LaFont.Font.Size:=FontSize;
if (LaFont.Execute) then
  Begin
  FontColor:=ColorToString(LaFont.Font.Color);
  If LaFont.Font.Style = [fsBold]      Then Fontstyle:='G';
  If LaFont.Font.Style = [fsItalic]    Then Fontstyle:='I';
  IF LaFont.Font.Style = [fsUnderline] Then Fontstyle:='U';
  IF LaFont.Font.Style = [fsStrikeout] Then Fontstyle:='S';
  FontName:=String(LaFont.Font.Name);
  FontSize:=Integer(LaFont.Font.Size);
  End;
end;
       }


initialization
  registerclasses([TOF_PGPLANNINGMERE, TOF_PGPLANNING, TOF_PGPLANNINGABSPAIE]);
end.

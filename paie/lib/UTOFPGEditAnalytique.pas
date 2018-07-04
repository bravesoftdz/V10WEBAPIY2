{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 12/09/2001
Modifié le ... :   /  /
Description .. : Etat analytique de la paie
Mots clefs ... : PAIE;ANALYTIQUE
*****************************************************************
PT1-1 17/01/2002 SB V571 Fiche de bug n°434 Nom Prénom erronée
PT1-2 24/01/2002 SB V571 Fiche de bug n°250 rename champ salarié
PT2-1 18/04/2002 SB V571 Fiche de bug n°10087 : V_PGI_env.LibDossier non renseigné en Mono
PT3   03/09/2002 SB V585 FQ n° 10011 Ajout de la colonne Divers Modification requête
PT4   19/03/2003 SB V591 FQ n° 10557 Incompatibilité d'édition
PT5   18/04/2005 SB V_65 FQ 11993 Ajout champ pha_ordreetat
PT6   04/07/2005 SB V_65 FQ 12020 Erreur Sql Oracle , 2 fois même champs dans select
PT7   21/09/2005 SB V_65 FQ 11786 Ajout champ rupture organisation
PT8   03/07/2007 FC V_72 FQ 13368 Pouvoir exclure les salariés sortis
}
unit UTOFPGEditAnalytique;

interface
uses Controls, Classes, sysutils, ComCtrls, HQry,
  {$IFNDEF EAGLCLIENT}
  QRS1,
  {$ELSE}
  eQRS1,
  {$ENDIF}
  HCtrls, HEnt1,  UTOF,
  ParamDat, LookUp,
  ParamSoc,StdCtrls;

type
  TOF_PGANALYTIQUE_ETAT = class(TOF)
    procedure OnArgument(Arguments: string); override;
  private
    StChamp, StJoin: string;
    procedure ExitEdit(Sender: TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure ControlRupture;
    function GetOrderBy(StOrder: string): string;
    procedure OnChangeRupture(Sender: TObject);
    procedure OnSectionClick(Sender: TObject);
  end;

  TOF_PGANALYSE_ETAT = class(TOF_PGANALYTIQUE_ETAT)
    procedure OnUpdate; override;
  end;

  TOF_PGANALRUBRIQUE_ETAT = class(TOF_PGANALYTIQUE_ETAT)
    procedure OnUpdate; override;
  end;

  TOF_PGVENTILSAL_ETAT = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;   //PT8
  private
    procedure ExitEdit(Sender: TObject);
    procedure OnSectionClick(Sender: TObject);
    procedure OnClickSalarieSortie(Sender: TObject); //PT8

  end;


implementation

uses PgEditOutils2, EntPaie, PGoutils2 ; 

{ TOF_PGANALYTIQUE_ETAT }

procedure TOF_PGANALYTIQUE_ETAT.OnArgument(Arguments: string);
var
  CDd, CDf, Defaut: THEdit;
  Control: TControl;
  Min, Max, DebPer, FinPer, ExerPerEncours: string;
  i, j : integer;
  Combo: THValComboBox;
begin
  inherited;
  //Valeur par défaut
//SetControlText('DOSSIER',V_PGI_env.LibDossier); //PT2-1 mise en commentaire
  SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
  SetControlChecked('CKEURO', VH_Paie.PGTenueEuro);

  RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
  Defaut := ThEdit(getcontrol('PHA_SALARIE'));
  if Defaut <> nil then
  begin
    Defaut.text := Min;
    Defaut.OnExit := ExitEdit;
  end;
  Defaut := ThEdit(getcontrol('PHA_SALARIE_'));
  if Defaut <> nil then
  begin
    Defaut.text := Max;
    Defaut.OnExit := ExitEdit;
  end;
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  Defaut := ThEdit(getcontrol('PHA_ETABLISSEMENT'));
  if Defaut <> nil then Defaut.text := Min;
  Defaut := ThEdit(getcontrol('PHA_ETABLISSEMENT_'));
  if Defaut <> nil then Defaut.text := Max;
  RecupMinMaxTablette('PG', 'SECTION', 'S_SECTION', Min, Max);
  Defaut := ThEdit(getcontrol('PHA_SECTION'));
  if Defaut <> nil then
  begin
    Defaut.text := Min;
    Defaut.OnElipsisClick := OnSectionClick;
  end;
  Defaut := ThEdit(getcontrol('PHA_SECTION_'));
  if Defaut <> nil then
  begin
    Defaut.text := Max;
    Defaut.OnElipsisClick := OnSectionClick;
  end;
  RecupMinMaxTablette('PG', 'GENERAUX', 'G_GENERAL', Min, Max);
  Defaut := ThEdit(getcontrol('PHA_GENERAL'));
  if Defaut <> nil then Defaut.text := Min;
  Defaut := ThEdit(getcontrol('PHA_GENERAL_'));
  if Defaut <> nil then Defaut.text := Max;
  Control := ThValComboBox(getcontrol('PHA_AXE'));
  InitialiseCombo(Control);


  for j := 1 to 4 do
  begin
    Combo := ThValComboBox(GetControl('THVALRUPTURE' + IntToStr(j)));
    { DEB PT7 }
    if Assigned(Combo) then
     Begin
     Combo.onchange := OnChangeRupture;
     Combo.Itemindex := 0;
     Combo.Items.Add('<<Aucune>>');          Combo.Values.Add('');
     Combo.Items.Add('Etablissement');       Combo.Values.Add('PHA_ETABLISSEMENT');
     Combo.Items.Add('Salarié');             Combo.Values.Add('PHA_SALARIE');
     Combo.Items.Add('Compte');              Combo.Values.Add('PHA_GENERAL');
     Combo.Items.Add('Axe');                 Combo.Values.Add('PHA_AXE');
     Combo.Items.Add('Section analytique');  Combo.Values.Add('PHA_SECTION');
     For i :=1 To VH_Paie.PGNbreStatOrg Do
       Begin
       IF (i= 1) AND (VH_Paie.PGLibelleOrgStat1 <> '') then Combo.Items.Add(VH_Paie.PGLibelleOrgStat1)
       else IF (i= 2) AND (VH_Paie.PGLibelleOrgStat2 <> '') then Combo.Items.Add(VH_Paie.PGLibelleOrgStat2)
       else IF (i= 3) AND (VH_Paie.PGLibelleOrgStat3 <> '') then Combo.Items.Add(VH_Paie.PGLibelleOrgStat3)
       else IF (i= 4) AND (VH_Paie.PGLibelleOrgStat4 <> '') then Combo.Items.Add(VH_Paie.PGLibelleOrgStat4);
       Combo.Values.Add('PHA_TRAVAILN'+IntToStr(i));
       End;
     if VH_Paie.PGLibCodeStat <> '' then
       Begin
       Combo.Items.Add(VH_Paie.PGLibCodeStat);
       Combo.Values.Add('PHA_CODESTAT');
       End;
     For i :=1 To VH_Paie.PgNbCombo Do
       Begin
       IF (i= 1) AND (VH_Paie.PgLibCombo1 <> '') then Combo.Items.Add(VH_Paie.PgLibCombo1)
       else IF (i= 2) AND (VH_Paie.PgLibCombo2 <> '') then Combo.Items.Add(VH_Paie.PgLibCombo2)
       else IF (i= 3) AND (VH_Paie.PgLibCombo3 <> '') then Combo.Items.Add(VH_Paie.PgLibCombo3)
       else IF (i= 4) AND (VH_Paie.PgLibCombo4 <> '') then Combo.Items.Add(VH_Paie.PgLibCombo4);
       Combo.Values.Add('PHA_LIBREPCMB'+IntToStr(i));
       End;
     End;
    { FIN PT7 }
  end;

  CDd := THEdit(GetControl('XX_VARIABLEDEB'));
  CDf := THEdit(GetControl('XX_VARIABLEFIN'));
  if CDd <> nil then CDd.OnElipsisClick := DateElipsisclick;
  if CDf <> nil then CDf.OnElipsisClick := DateElipsisclick;
  if RendPeriodeEnCours(ExerPerEncours, DebPer, FinPer) = True then
  begin
    if CDd <> nil then CDd.text := DebPer;
    if CDf <> nil then CDf.text := FinPer;
  end;


end;

procedure TOF_PGANALYTIQUE_ETAT.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGANALYTIQUE_ETAT.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;


procedure TOF_PGANALYTIQUE_ETAT.OnChangeRupture(Sender: TObject);
var
  Combo: THValComboBox;
  ChampRupture, Name, Ordre: string;
  IntOrdre, i: integer;
begin
  if Sender is THValComboBox then Combo := THValComboBox(Sender) else exit;
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
    SetControlText(ChampRupture, Combo.value);
  end;
end;

procedure TOF_PGANALYTIQUE_ETAT.ControlRupture;
var
  Rupt1, Rupt2, Rupt3, Rupt4: string;
begin
  Rupt1 := GetControlText('XX_RUPTURE1');
  Rupt2 := GetControlText('XX_RUPTURE2');
  Rupt3 := GetControlText('XX_RUPTURE3');
  Rupt4 := GetControlText('XX_RUPTURE4');

  if Rupt4 = Rupt3 then Rupt4 := ''
  else if Rupt4 = Rupt2 then Rupt4 := ''
  else if Rupt4 = Rupt1 then Rupt4 := '';

  if Rupt3 = Rupt2 then Rupt3 := ''
  else if Rupt3 = Rupt1 then Rupt3 := '';

  if Rupt2 = Rupt1 then Rupt2 := '';
  SetControlText('XX_RUPTURE2', Rupt2);
  SetControlText('XX_RUPTURE3', Rupt3);
  SetControlText('XX_RUPTURE4', Rupt4);
  if Rupt2 = '' then SetControlText('THVALRUPTURE2', '');
  if Rupt3 = '' then SetControlText('THVALRUPTURE3', '');
  if Rupt4 = '' then SetControlText('THVALRUPTURE4', '');
  StChamp := '';
  if Rupt1 <> '' then StChamp := Rupt1 + ',';
  if Rupt2 <> '' then StChamp := StChamp + Rupt2 + ',';
  if Rupt3 <> '' then StChamp := StChamp + Rupt3 + ',';
  if Rupt4 <> '' then StChamp := StChamp + Rupt4 + ',';
  if Pos('PHA_SALARIE', StChamp) > 0 then // PT1-1 Chgmt cond. =1
  begin
    StChamp := StChamp + 'PSA_LIBELLE,PSA_PRENOM,';
    StJoin := ' LEFT JOIN SALARIES ON PSA_SALARIE=PHA_SALARIE ';
  end
  else StJoin := '';
end;

function TOF_PGANALYTIQUE_ETAT.GetOrderBy(StOrder: string): string;
var
  Order1, Order2, Order3, Order4, NewOrderBY: string;
begin
  Order1 := '';
  Order2 := '';
  Order3 := '';
  Order4 := '';
  result := StOrder;
  if StOrder <> '' then Order1 := Trim(ReadTokenPipe(StOrder, ','));
  if StOrder <> '' then Order2 := Trim(ReadTokenPipe(StOrder, ','));
  if StOrder <> '' then Order3 := Trim(ReadTokenPipe(StOrder, ','));
  if StOrder <> '' then Order4 := Trim(ReadTokenPipe(StOrder, ','));

  if Order4 = Order3 then order4 := ''
  else
    if Order4 = Order2 then order4 := ''
  else
    if Order4 = Order1 then order4 := '';
  if Order3 = Order2 then Order3 := ''
  else
    if Order3 = Order1 then Order3 := '';
  if Order2 = Order1 then Order2 := '';

  NewOrderBY := Order1;
  if (Order2 <> '') then NewOrderBY := NewOrderBY + ',' + Order2;
  if (Order3 <> '') then NewOrderBY := NewOrderBY + ',' + Order3;
  if (Order4 <> '') then NewOrderBY := NewOrderBY + ',' + Order4;
  if (Order1 <> '') then result := 'ORDER BY ' + NewOrderBY;
end;

procedure TOF_PGANALYTIQUE_ETAT.OnSectionClick(Sender: TObject);
var
  Edit: THEdit;
begin
  if Sender is ThEdit then Edit := THEdit(Sender) else Exit;
  LookUpList(Edit, 'Section Analytique', 'SECTION', 'S_SECTION', 'S_LIBELLE', '', 'S_SECTION', TRUE, -1);
end;


{ TOF_PGANALYSE_ETAT }

procedure TOF_PGANALYSE_ETAT.OnUpdate;
var
  DateDeb, DateFin: TDateTime;
  SQL, Temp, Tempo, Critere, OrderBy : string;
  Pages: TPageControl;
  x: integer;
begin
  inherited;
  ControlRupture;
  DateDeb := StrToDate(GetControlText('XX_VARIABLEDEB'));
  DateFin := StrToDate(GetControlText('XX_VARIABLEFIN'));
  X := Pos('|', TFQRS1(Ecran).WhereSQL);
  if X > 0 then
  begin
    OrderBy := GetOrderBy(Copy(TFQRS1(Ecran).WhereSQL, x + 1, Length(TFQRS1(Ecran).WhereSQL))); //PT4+',PHA_GENERAL' //PT3
    if (OrderBy <> '') and (Pos('PHA_GENERAL', OrderBy) < 1) then OrderBy := OrderBy + ',PHA_GENERAL'; //PT4 ajout
  end
  else
    OrderBy := 'ORDER BY PHA_GENERAL';

  Pages := TPageControl(GetControl('Pages'));
  Temp := RecupWhereCritere(Pages);
  tempo := '';
  critere := '';
  x := Pos('(', Temp);

  if x > 0 then Tempo := copy(Temp, x, (Length(temp) - 5));
  if tempo <> '' then critere := 'AND ' + Tempo;
  IF Pos('PHA_GENERAL',StChamp)>0 then Sql:= '' else Sql := 'PHA_GENERAL,';  { PT6 }

  SQL := 'SELECT ' + StChamp + Sql+'PRM_THEMEREM,' + { PT6 }
    'SUM(PHA_MTREM) AS MTREM,SUM(PHA_MTSALARIAL) AS MTSALARIAL,' +
    'SUM(PHA_MTPATRONAL) AS MTPATRONAL ' +
    'FROM HISTOANALPAIE ' + StJoin + //PT3 Ajout join et champ PRM_THEMEREM
  'LEFT JOIN REMUNERATION ON PHA_NATURERUB=PRM_NATURERUB AND ##PRM_PREDEFINI## PHA_RUBRIQUE=PRM_RUBRIQUE ' +
    'WHERE PHA_DATEDEBUT>="' + UsDateTime(DateDeb) + '" AND PHA_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
    ' ' + critere +
    'GROUP BY ' + StChamp + Sql +'PRM_THEMEREM ' + OrderBy; { PT6 }

  TFQRS1(Ecran).WhereSQL := SQL;
end;


{ TOF_PGANALRUBRIQUE_ETAT }

procedure TOF_PGANALRUBRIQUE_ETAT.OnUpdate;
var
  DateDeb, DateFin: TDateTime;
  SQL, Temp, Tempo, Critere, OrderBy: string;
  Pages: TPageControl;
  x: integer;
begin
  inherited;
  ControlRupture;
  DateDeb := StrToDate(GetControlText('XX_VARIABLEDEB'));
  DateFin := StrToDate(GetControlText('XX_VARIABLEFIN'));
  X := Pos('|', TFQRS1(Ecran).WhereSQL);
  if X > 0 then
    OrderBy := GetOrderBy(Copy(TFQRS1(Ecran).WhereSQL, x + 1, Length(TFQRS1(Ecran).WhereSQL))) + ',PHA_ORDREETAT,PHA_RUBRIQUE' //PT3 PT5
  else OrderBy := 'ORDER BY PHA_ORDREETAT,PHA_RUBRIQUE';     //PT5

  Pages := TPageControl(GetControl('Pages'));
  Temp := RecupWhereCritere(Pages);
  tempo := '';
  critere := '';
  x := Pos('(', Temp);

  if x > 0 then Tempo := copy(Temp, x, (Length(temp) - 5));
  if tempo <> '' then critere := 'AND ' + Tempo;

  SQL := 'SELECT ' + StChamp + 'PHA_RUBRIQUE,PHA_LIBELLE,PRM_THEMEREM,PRM_BASEMTQTE,SUM(PHA_BASEREM) AS BASEREM,' +
    'SUM(PHA_MTREM) AS MTREM,SUM(PHA_MTSALARIAL) AS MTSALARIAL,' +
    'SUM(PHA_MTPATRONAL) AS MTPATRONAL ' +
    'FROM HISTOANALPAIE ' + StJoin + '  ' + //PT3 Ajout champ PRM_THEMEREM
  'LEFT JOIN REMUNERATION ON PHA_RUBRIQUE=PRM_RUBRIQUE AND ##PRM_PREDEFINI## PHA_NATURERUB=PRM_NATURERUB ' +
    'WHERE PHA_DATEDEBUT>="' + UsDateTime(DateDeb) + '" AND PHA_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
    ' ' + critere + ' ' +
    'GROUP BY ' + StChamp + 'PHA_ORDREETAT,PHA_RUBRIQUE,PHA_LIBELLE,PRM_THEMEREM,PRM_BASEMTQTE ' + OrderBy; //PT5

  TFQRS1(Ecran).WhereSQL := SQL;
end;



{ TOF_PGVENTILSAL_ETAT }

procedure TOF_PGVENTILSAL_ETAT.OnArgument(Arguments: string);
var
  Defaut: THEdit;
  Min, Max: string;
  Check:TCheckBox;//PT8
begin
  inherited;
  //Valeur par défaut
  Defaut := ThEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
    //  Defaut.text:=V_PGI_env.LibDossier;  //PT2-1 mise en commentaire
    Defaut.text := GetParamSoc('SO_LIBELLE');
  RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
  Defaut := ThEdit(getcontrol('PSA_SALARIE')); //PT1-2  V_COMPTE
  if Defaut <> nil then
  begin
    Defaut.text := Min;
    Defaut.OnExit := ExitEdit;
  end;
  Defaut := ThEdit(getcontrol('PSA_SALARIE_')); //PT1-2
  if Defaut <> nil then
  begin
    Defaut.text := Max;
    Defaut.OnExit := ExitEdit;
  end;
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  Defaut := ThEdit(getcontrol('PSA_ETABLISSEMENT'));
  if Defaut <> nil then Defaut.text := Min;
  Defaut := ThEdit(getcontrol('PSA_ETABLISSEMENT_'));
  if Defaut <> nil then Defaut.text := Max;
  RecupMinMaxTablette('PG', 'SECTION', 'S_SECTION', Min, Max);
  Defaut := ThEdit(getcontrol('V_SECTION'));
  if Defaut <> nil then
  begin
    Defaut.text := Min;
    Defaut.OnElipsisClick := OnSectionClick;
  end;
  Defaut := ThEdit(getcontrol('V_SECTION_'));
  if Defaut <> nil then
  begin
    Defaut.text := Max;
    Defaut.OnElipsisClick := OnSectionClick;
  end;

  //DEB PT8
  Check := TCheckBox(GetControl('CKSORTIE'));
  if Check <> nil then
    Check.OnClick := OnClickSalarieSortie;
  //FIN PT8
end;

procedure TOF_PGVENTILSAL_ETAT.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;


procedure TOF_PGVENTILSAL_ETAT.OnSectionClick(Sender: TObject);
var
  Edit: THEdit;
begin
  Edit := nil;
  if Sender is ThEdit then Edit := THEdit(Sender);
  if Edit <> nil then
    LookUpList(Edit, 'Section Analytique', 'SECTION', 'S_SECTION', 'S_LIBELLE', '', 'S_SECTION', TRUE, -1);
end;

//DEB PT8
procedure TOF_PGVENTILSAL_ETAT.OnUpdate;
var
  StWhere : String;
  Pages : TPageControl;
  DateArret : TDateTime;
begin
  inherited;
  Pages := TPageControl(GetControl('Pages'));
  StWhere := TFQRS1(Ecran).WhereSQL;

  if TCheckBox(GetControl('CKSORTIE')) <> nil then
  begin
    if (GetControlText('CKSORTIE') = 'X') and (IsValidDate(GetControlText('DATEARRET')))then
    begin
      DateArret := StrtoDate(GetControlText('DATEARRET'));
      if (StWhere <> '') then
        StWhere := StWhere + ' AND (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) '
                           + ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"'
      else
        StWhere := ' (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) '
                           + ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
    end;
  end;

  TFQRS1(Ecran).WhereSQL := StWhere;
end;

procedure TOF_PGVENTILSAL_ETAT.OnClickSalarieSortie(Sender: TObject);
begin
  SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
  SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;
//FIN PT8

initialization
  registerclasses([TOF_PGANALYTIQUE_ETAT, TOF_PGANALYSE_ETAT, TOF_PGANALRUBRIQUE_ETAT, TOF_PGVENTILSAL_ETAT]);
end.


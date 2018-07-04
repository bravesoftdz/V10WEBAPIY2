{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 24/01/2002
Modifié le ... :   /  /
Description .. : Edition des bases de cotisation
Mots clefs ... : PAIE;BASE
*****************************************************************
PT1 24/01/2002 SB V571 Changement Cond. requêtes
PT2 28/01/2002 SB V571 Sélection alphanumérique du code salarié
PT3 18/04/2002 SB V571 Fiche de bug n°10087 : V_PGI_env.LibDossier non renseigné en Mono
PT4 02/10/2003 SB V_42 Affichage des ongles si gestion paramsoc des combos libres
PT5 28/06/2007 FC V_72 FQ 13508 Ajout case salariés sortis
PT6 : 01/10/2008 SJ FQ n°15267 Pouvoir sélectionner une rubrique de base
}
unit UTOFPGEditBaseCotisation;

interface
uses StdCtrls, Classes, sysutils, ComCtrls, //forms,Graphics,Controls,   
  {$IFDEF EAGLCLIENT}
  eQRS1, UTOB,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} QRS1,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, ParamDat,
  ParamSoc, HQry;

type
  TOF_PGBASECOT = class(TOF)
  private
    AddChamp, ExDateDeb, ExDateFin: string;
    procedure DateElipsisclick(Sender: TObject);
    procedure Change(Sender: TObject);
    procedure ChangeLieuTravail(Sender: TObject);
    procedure ChangeExercice(Sender: TObject);
    function CreateOrderBy: string;
    procedure ControlPeriodeDeb(Sender: TObject);
    procedure ControlPeriodeFin(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
  end;

implementation

uses PgEditOutils,PgEditOutils2, EntPaie, PGoutils2;

{ TOF_PGBASECOT }

procedure TOF_PGBASECOT.Change(Sender: TObject);
begin
  if GetControlText('CPERIODE') = 'X' then SetControlText('XX_RUPTURE3', 'PHB_DATEFIN')
  else SetControlText('XX_RUPTURE3', '');
  if GetControltext('CSAL') = 'X' then
  begin
    SetControlEnabled('CALPHA', True);
    SetControlText('XX_RUPTURE2', 'PHB_SALARIE');
    SetControlEnabled('CKSORTIE',True);   //PT5
  end
  else
  begin
    SetControlChecked('CALPHA', False);
    SetControlEnabled('CALPHA', False);
    SetControlText('XX_RUPTURE2', '');
    SetControlEnabled('CKSORTIE',False);   //PT5
    SetControlChecked('CKSORTIE', False);  //PT5
  end;
  if TCheckBox(Sender).Name = 'CALPHA' then //PT2
    AffectCritereAlpha(Ecran, TCheckBox(Sender).Checked, 'PHB_SALARIE', 'PSA_LIBELLE');
end;

procedure TOF_PGBASECOT.ChangeExercice(Sender: TObject);
var
  QPeriode: TQuery;
begin
  QPeriode := OpenSql('SELECT MIN(PPU_DATEDEBUT),MAX(PPU_DATEFIN) FROM PAIEENCOURS ' +
    'LEFT JOIN EXERSOCIAL ON PPU_DATEDEBUT>=PEX_DATEDEBUT AND PPU_DATEFIN<=PEX_DATEFIN ' +
    'WHERE PEX_EXERCICE="' + GetControlText('EDEXERSOC') + '"', TRUE);
  if not QPeriode.eof then
  begin
    SetControlText('XX_VARIABLEDEB', DateToStr(QPeriode.Fields[0].AsDateTime));
    SetControlText('XX_VARIABLEFIN', DateToStr(QPeriode.Fields[1].AsDateTime));
  end
  else
  begin
    SetControlText('XX_VARIABLEDEB', DateToStr(idate1900));
    SetControlText('XX_VARIABLEFIN', DateToStr(idate1900));
  end;
  Ferme(QPeriode);
end;

procedure TOF_PGBASECOT.ChangeLieuTravail(Sender: TObject);
begin
  RecupChampRupture(Ecran);
  BloqueChampLibre(Ecran);
end;

procedure TOF_PGBASECOT.ControlPeriodeDeb(Sender: TObject);
var
  Edit: ThEdit;
  DebExer: TDateTime;
  Q: TQuery;
begin
  Q := OpenSql('SELECT PEX_DATEDEBUT FROM EXERSOCIAL WHERE PEX_EXERCICE="' + GetControlText('EDEXERSOC') + '" ', True);
  if not Q.eof then DebExer := Q.findfield('PEX_DATEDEBUT').AsDateTime
  else DebExer := idate1900;
  Ferme(Q);
  Edit := ThEdit(getcontrol('XX_VARIABLEDEB'));
  if StrToDate(Edit.Text) < (DebExer) then
  begin
    PgiBox('La date de début ne peut être inferieur au ' + DateToStr(DebExer) + '', 'Date Erronée!');
    Edit.text := DateToStr(DebExer);
  end;
end;

procedure TOF_PGBASECOT.ControlPeriodeFin(Sender: TObject);
var
  FinExer: TDateTime;
  Q: TQuery;
begin
  Q := OpenSql('SELECT PEX_DATEFIN FROM EXERSOCIAL WHERE PEX_EXERCICE="' + GetControlText('EDEXERSOC') + '" ', True);
  if not Q.Eof then //PORTAGECWAS
    FinExer := Q.findfield('PEX_DATEFIN').AsDateTime
  else
    FinExer := idate1900;
  Ferme(Q);
  if (StrToDate(GetControlText('XX_VARIABLEFIN')) > FinExer) and (FinExer <> idate1900) then
  begin
    PgiBox('La date de fin ne peut être supérieur au ' + DateToStr(FinExer) + '', 'Date Erronée!');
    SetControlText('XX_VARIABLEFIN', DateToStr(FinExer));
  end;
end;

function TOF_PGBASECOT.CreateOrderBy: string;
var
  StRupt, StAlpha, StPeriode, StSal, St: string;
begin
  StRupt := '';
  StAlpha := '';
  StSal := '';
  St := '';
  AddChamp := '';
  if GetControlText('CN1') = 'X' then StRupt := 'PHB_TRAVAILN1';
  if GetControlText('CN2') = 'X' then StRupt := 'PHB_TRAVAILN2';
  if GetControlText('CN3') = 'X' then StRupt := 'PHB_TRAVAILN3';
  if GetControlText('CN4') = 'X' then StRupt := 'PHB_TRAVAILN4';
  if GetControlText('CN5') = 'X' then StRupt := 'PHB_CODESTAT';
  if GetControlText('CETAB') = 'X' then StRupt := 'PHB_ETABLISSEMENT';

  if StRupt <> '' then AddChamp := StRupt + ',';
  if GetControlText('CSAL') = 'X' then StSal := 'PHB_SALARIE,PSA_LIBELLE,PSA_PRENOM';
  if GetControlText('CALPHA') = 'X' then
  begin
    StSal := '';
    StAlpha := 'PSA_LIBELLE,PSA_PRENOM,PHB_SALARIE';
  end;
  if GetControlText('CPERIODE') = 'X' then StPeriode := 'PHB_DATEDEBUT,PHB_DATEFIN';

  if (StRupt <> '') and (StSal <> '') and (StAlpha <> '') and (StPeriode <> '') then St := StRupt + ',' + StAlpha + ',' + StSal + ',' + StPeriode;


  if (StRupt <> '') and (StSal = '') and (StAlpha = '') and (StPeriode = '') then St := StRupt;
  if (StRupt <> '') and (StSal <> '') and (StAlpha = '') and (StPeriode = '') then St := StRupt + ',' + StSal;
  if (StRupt <> '') and (StSal = '') and (StAlpha <> '') and (StPeriode = '') then St := StRupt + ',' + StAlpha;
  if (StRupt <> '') and (StSal = '') and (StAlpha = '') and (StPeriode <> '') then St := StRupt + ',' + StPeriode;
  if (StRupt <> '') and (StSal <> '') and (StAlpha = '') and (StPeriode <> '') then St := StRupt + ',' + StSal + ',' + StPeriode;

  if (StRupt = '') and (StSal <> '') and (StAlpha = '') and (StPeriode = '') then St := StSal;
  if (StRupt = '') and (StSal = '') and (StAlpha <> '') and (StPeriode = '') then St := StAlpha;
  if (StRupt = '') and (StSal <> '') and (StAlpha = '') and (StPeriode <> '') then St := StSal + ',' + StPeriode;
  if (StRupt = '') and (StSal = '') and (StAlpha <> '') and (StPeriode <> '') then St := StAlpha + ',' + StPeriode;
  if (StRupt = '') and (StSal = '') and (StAlpha = '') and (StPeriode <> '') then St := StPeriode;

  if St <> '' then result := St + ',' else result := '';
end;

procedure TOF_PGBASECOT.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGBASECOT.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if edit.text <> '' then
      if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
        edit.text := AffectDefautCode(edit, 10);
end;


procedure TOF_PGBASECOT.OnArgument(Arguments: string);
var
  THDateDeb, THDateFin, Defaut: THEdit;
  Exercice: THValComboBox;
  QPeriode, QExer: TQuery;
  Check: TCheckBox;
  Min, Max: string;
begin
  inherited;
  //SetControlText('DOSSIER',V_PGI_env.LibDossier); PT3 Mise en commentaire
  SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));

  SetControlChecked('CKEURO', VH_Paie.PGTenueEuro);
  SetControlChecked('CKPAIE', VH_Paie.PGDecalage);
  SetControlText('XX_RUPTURE2', 'PHB_SALARIE');

  VisibiliteChamp(Ecran);
  VisibiliteChampLibre(Ecran);
  { DEB PT4 }
  SetControlProperty('TBCOMPLEMENT', 'Tabvisible', (VH_Paie.PGNbreStatOrg > 0) or (VH_Paie.PGLibCodeStat <> ''));
  SetControlProperty('TBCHAMPLIBRE', 'Tabvisible', (VH_Paie.PgNbCombo > 0));
  { FIN PT4 }

  //Valeur par défaut
  RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
  Defaut := ThEdit(getcontrol('PHB_SALARIE'));
  if Defaut <> nil then
  begin
    Defaut.text := Min;
    Defaut.OnExit := ExitEdit;
  end;
  Defaut := ThEdit(getcontrol('PHB_SALARIE_'));
  if Defaut <> nil then
  begin
    Defaut.text := Max;
    Defaut.OnExit := ExitEdit;
  end;
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  Defaut := ThEdit(getcontrol('PHB_ETABLISSEMENT'));
  if Defaut <> nil then
  begin
    Defaut.text := Min;
  end;
  Defaut := ThEdit(getcontrol('PHB_ETABLISSEMENT_'));
  if Defaut <> nil then
  begin
    Defaut.text := Max;
  end;
  //debut PT6
  RecupMinMaxTablette('PG', 'COTISATION', 'PCT_RUBRIQUE', Min, Max);
  Defaut := ThEdit(getcontrol('PHB_RUBRIQUE'));
  if Defaut <> nil then
  begin
    Defaut.text := Min;
  end;
  Defaut := ThEdit(getcontrol('PHB_RUBRIQUE_'));
  if Defaut <> nil then
  begin
    Defaut.text := Max;;
  end;
  //fin PT6
  QExer := OpenSql('SELECT MAX(PEX_EXERCICE),MAX(PEX_DATEDEBUT) FROM EXERSOCIAL ' +
    'WHERE PEX_ACTIF="X"', true);
  if not QExer.eof then
    SetControlText('EDEXERSOC', QExer.Fields[0].asstring)
  else
    SetControlText('EDEXERSOC', '');
  Ferme(QExer);

  QPeriode := OpenSql('SELECT MIN(PPU_DATEDEBUT),MAX(PPU_DATEFIN) FROM PAIEENCOURS ' +
    'LEFT JOIN EXERSOCIAL ON PPU_DATEDEBUT>=PEX_DATEDEBUT AND PPU_DATEFIN<=PEX_DATEFIN ' +
    'WHERE PEX_EXERCICE="' + GetControltext('EDEXERSOC') + '"', TRUE);
  if not QPeriode.eof then
  begin
    ExDateDeb := DateToStr(QPeriode.Fields[0].AsDateTime);
    ExDateFin := DateToStr(QPeriode.Fields[1].AsDateTime);
  end
  else
  begin
    ExDateDeb := DateToStr(idate1900);
    ExDateFin := DateToStr(idate1900);
  end;
  Ferme(QPeriode);
  Exercice := THValComboBox(getcontrol('EDEXERSOC'));
  THDateDeb := ThEdit(getcontrol('XX_VARIABLEDEB'));
  THDateFin := ThEdit(getcontrol('XX_VARIABLEFIN'));
  if (THDateDeb <> nil) and (THDateFin <> nil) and (Exercice <> nil) then
  begin
    THDateDeb.text := ExDateDeb;
    THDateDeb.OnElipsisClick := DateElipsisclick;
    THDateDeb.OnExit := ControlPeriodeDeb;
    THDateFin.text := ExDateFin;
    THDateFin.OnElipsisClick := DateElipsisclick;
    THDateFin.OnExit := ControlPeriodeFin;
    Exercice.OnChange := ChangeExercice;
  end;

  Check := TCheckBox(GetControl('CN1'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CN2'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CN3'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CN4'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CN5'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CETAB'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL1'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL2'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL3'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL4'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;

  Check := TCheckBox(GetControl('CSAL'));
  if Check <> nil then
  begin
    Check.Checked := True;
    Check.OnClick := Change;
  end;
  Check := TCheckBox(GetControl('CALPHA'));
  if Check <> nil then Check.OnClick := Change;
  Check := TCheckBox(GetControl('CPERIODE'));
  if Check <> nil then Check.OnClick := Change;
end;

procedure TOF_PGBASECOT.OnUpdate;
var
  Temp, Tempo, Critere, DateDeb, DateFin, Join, Order: string;
  Pages: TPageControl;
  x: integer;
begin
  inherited;
  Pages := TPageControl(GetControl('Pages'));
  Temp := RecupWhereCritere(Pages);
  tempo := '';
  critere := '';
  x := Pos('(', Temp);
  if x > 0 then Tempo := copy(Temp, x, (Length(temp) - 5));
  if tempo <> '' then critere := 'AND ' + Tempo;
  DateDeb := GetControlText('XX_VARIABLEDEB');
  DateFin := GetControlText('XX_VARIABLEFIN');
  { DEB PT5 }
  if (GetControltext('CKSORTIE')='X') and (UsdateTime(StrToDate(DateFin))<>UsdateTime(idate1900)) then
  Critere:=Critere+' AND PHB_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_DATESORTIE IS NULL '+
                      'OR PSA_DATESORTIE="'+Usdatetime(idate1900)+'" '+
                      'OR PSA_DATESORTIE>"'+UsdateTime(StrToDate(DateFin))+'")';
  { FIN PT5 }

  Order := CreateOrderBy;
  AddChamp := Order;
  if Pos('PHB_SALARIE', AddChamp) > 0 then Join := 'LEFT JOIN SALARIES ON PSA_SALARIE=PHB_SALARIE ' else Join := '';
  TFQRS1(Ecran).WhereSQL := 'SELECT ' + AddChamp + ' PHB_RUBRIQUE, ' +
    'PHB_LIBELLE,SUM(PHB_BASECOT) BASECOT,SUM(PHB_PLAFOND) PLAFOND,' +
    'SUM(PHB_PLAFOND1) PLAFOND1,SUM(PHB_PLAFOND2) PLAFOND2,' +
    'SUM(PHB_PLAFOND3) PLAFOND3,SUM(PHB_TRANCHE1) TRANCHE1,' +
    'SUM(PHB_TRANCHE2) TRANCHE2,SUM(PHB_TRANCHE3) TRANCHE3 ' +
    'FROM HISTOBULLETIN ' + Join +
    'WHERE PHB_NATURERUB="BAS" ' +
    'AND PHB_DATEFIN>="' + UsDateTime(StrToDate(DateDeb)) + '"  ' + //PT1
  'AND PHB_DATEFIN<="' + UsDateTime(StrToDate(DateFin)) + '" ' +
    'AND (PHB_BASECOT<>0 OR PHB_PLAFOND<>0 OR PHB_PLAFOND<>0 OR PHB_PLAFOND1<>0 ' +
    'OR PHB_PLAFOND2 <>0 OR PHB_PLAFOND3<>0 OR PHB_TRANCHE1<>0 ' +
    'OR PHB_TRANCHE2<>0 OR PHB_TRANCHE3<>0 ) ' + critere + ' ' +
    'GROUP BY ' + AddChamp + 'PHB_RUBRIQUE,PHB_LIBELLE ' +
    'ORDER BY ' + Order + 'PHB_RUBRIQUE';

end;

initialization
  registerclasses([TOF_PGBASECOT]);
end.

 

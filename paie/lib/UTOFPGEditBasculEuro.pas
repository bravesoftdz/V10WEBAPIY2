{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 12/09/2001
Modifié le ... :   /  /
Description .. : Etat de contrôle de la bascule euro
Mots clefs ... : PAIE;EURO
*****************************************************************
PT-1 26/09/2001 V562 SB Contrôle date de bascule en euro
PT-2 11/03/2002 V571 SB Correction Conversion de variant incorrecte
PT-3 18/04/2002 V571 SB Fiche de bug n°10087 : V_PGI_env.LibDossier non renseigné en Mono
PT-4 02/10/2003 V_42 SB Affichage des ongles si gestion paramsoc des combos libres
}
unit UTOFPGEditBasculEuro;

interface
uses StdCtrls, Classes,  sysutils, ComCtrls,
  HCtrls, UTOF, HQry, HEnt1, HMsgBox, ParamSoc,
  {$IFDEF EAGLCLIENT}
  eQRS1, UTOB,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} QRS1,
  {$ENDIF}
  ParamDat;

type

  TOF_PGBASCULEURO_ETAT = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
  private
    AddChamp, ExDateDeb, ExDateFin: string;
    Recap: Boolean;
    procedure DateElipsisclick(Sender: TObject);
    procedure ChangeLieuTravail(Sender: TObject);
    procedure ChangeExercice(Sender: TObject);
    function CreateOrderBy: string;
    procedure ControlPeriodeDeb(Sender: TObject);
    procedure ControlPeriodeFin(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
  end;
implementation

uses EntPaie, PgEditOutils,PgEditOutils2, PGoutils2;

procedure TOF_PGBASCULEURO_ETAT.OnArgument(Arguments: string);
var
  THDateDeb, THDateFin, Defaut: THEdit;
  Exercice: THValComboBox;
  QPeriode, QExer: TQuery;
  CTrav1, CTrav2, CTrav3, CTrav4, CTrav5, CEtab: TCheckBox;
  Min, Max: string;
  check: TCheckBox;
begin

  Recap := False;
  //SetControlText('DOSSIER',V_PGI_env.LibDossier);    //PT-3 Mise en commentaire
  SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
  SetControlChecked('CKEURO', VH_Paie.PGTenueEuro);
  SetControlChecked('CKPAIE', VH_Paie.PGDecalage);
  SetControlChecked('CSAL', True);
  SetControltext('XX_RUPTURE2', 'PHB_SALARIE');
  {DEB1 PT-1}
  if VH_Paie.PGDateBasculEuro <= idate1900 then
  begin
    PgiBox('Vous n''avez effectué aucune bascule en euro.Edition impossible.', 'Basculement en euro');
    TFQRS1(Ecran).BValider.Enabled := False;
  end;
  {FIN1 PT-1}
  VisibiliteChamp(Ecran);
  VisibiliteChampLibre(Ecran);
  { DEB PT-4 }
  SetControlProperty('TBCOMPLEMENT', 'Tabvisible', (VH_Paie.PGNbreStatOrg > 0) or (VH_Paie.PGLibCodeStat <> ''));
  SetControlProperty('TBCHAMPLIBRE', 'Tabvisible', (VH_Paie.PgNbCombo > 0));
  { FIN PT-4 }
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

  QExer := OpenSql('SELECT MAX(PEX_EXERCICE),MAX(PEX_DATEDEBUT) FROM EXERSOCIAL ' +
    'WHERE PEX_ACTIF="X"', true);
  if not QExer.eof then SetControlText('EDEXERSOC', QExer.Fields[0].asstring); // PORTAGECWAS
  Ferme(QExer);

  QPeriode := OpenSql('SELECT MIN(PPU_DATEDEBUT),MAX(PPU_DATEFIN) FROM PAIEENCOURS ' +
    'LEFT JOIN EXERSOCIAL ON PPU_DATEDEBUT>=PEX_DATEDEBUT AND PPU_DATEFIN<=PEX_DATEFIN ' +
    'WHERE PEX_EXERCICE="' + GetControlText('EDEXERSOC') + '"', TRUE);
  if not QPeriode.Eof then
  begin
    ExDateDeb := QPeriode.Fields[0].asstring;
    ExDateFin := QPeriode.Fields[1].asstring;
  end
  else
  begin
    ExDateDeb := DateToStr(idate1900);
    ExDateFin := DateToStr(idate1900);
  end;
  Ferme(QPeriode);

  THDateDeb := ThEdit(getcontrol('XX_VARIABLEDEB'));
  THDateFin := ThEdit(getcontrol('XX_VARIABLEFIN'));
  Exercice := ThValComboBox(GetControl('EDEXERSOC'));
  if (THDateDeb <> nil) and (THDateFin <> nil) and (Exercice <> nil) then
  begin
    THDateDeb.text := ExDateDeb;
    THDateDeb.OnElipsisClick := DateElipsisclick;
    THDateDeb.OnExit := ControlPeriodeDeb;
    THDateFin.text := ExDateFin;
    if IsValidDate(ExDateFin) then //PT-2 Ajout Cond.
      if VH_Paie.PGDateBasculEuro < StrToDate(ExDateFin) then
        THDateFin.text := DateToStr(VH_Paie.PGDateBasculEuro);
    THDateFin.OnElipsisClick := DateElipsisclick;
    THDateFin.OnExit := ControlPeriodeFin;
    Exercice.OnChange := ChangeExercice;
  end;

  CTrav1 := TCheckBox(GetControl('CN1'));
  if CTrav1 <> nil then CTrav1.OnClick := ChangeLieuTravail;
  CTrav2 := TCheckBox(GetControl('CN2'));
  if CTrav2 <> nil then CTrav2.OnClick := ChangeLieuTravail;
  CTrav3 := TCheckBox(GetControl('CN3'));
  if CTrav3 <> nil then CTrav3.OnClick := ChangeLieuTravail;
  CTrav4 := TCheckBox(GetControl('CN4'));
  if CTrav4 <> nil then CTrav4.OnClick := ChangeLieuTravail;
  CTrav5 := TCheckBox(GetControl('CN5'));
  if CTrav5 <> nil then CTrav5.OnClick := ChangeLieuTravail;
  CEtab := TCheckBox(GetControl('CETAB'));
  if CEtab <> nil then CEtab.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL1'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL2'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL3'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;
  Check := TCheckBox(GetControl('CL4'));
  if Check <> nil then Check.OnClick := ChangeLieuTravail;


end;

procedure TOF_PGBASCULEURO_ETAT.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGBASCULEURO_ETAT.ChangeExercice(Sender: TObject);
var
  Exercice: string;
  QPeriode: TQuery;
begin
  Exercice := GetControlText('EDEXERSOC');
  QPeriode := OpenSql('SELECT MIN(PPU_DATEDEBUT),MAX(PPU_DATEFIN) FROM PAIEENCOURS ' +
    'LEFT JOIN EXERSOCIAL ON PPU_DATEDEBUT>=PEX_DATEDEBUT AND PPU_DATEFIN<=PEX_DATEFIN WHERE PEX_EXERCICE="' + Exercice + '"', TRUE);
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
  SetControlText('XX_VARIABLEDEB', ExDateDeb);
  SetControlText('XX_VARIABLEFIN', ExDateFin);
  {DEB2 PT-1}
  if StrToDate(GetControlText('XX_VARIABLEFIN')) > VH_Paie.PGDateBasculEuro then
  begin
    PgiBox('La date de fin ne peut être supérieur à la date de basculement en euro.', 'Date Erronée!');
    SetControlText('XX_VARIABLEFIN', DateToStr(VH_Paie.PGDateBasculEuro));
  end;
  {FIN2 PT-1}
  Ferme(QPeriode);
end;

function TOF_PGBASCULEURO_ETAT.CreateOrderBy: string;
var
  StRupt, StAlpha, StSal, St: string;
begin
  StRupt := '';
  StAlpha := '';
  StSal := '';
  St := '';
  AddChamp := '';

  if (GetControlText('CN1') = 'X') then StRupt := 'PHB_TRAVAILN1';
  if (GetControlText('CN2') = 'X') then StRupt := 'PHB_TRAVAILN2';
  if (GetControlText('CN3') = 'X') then StRupt := 'PHB_TRAVAILN3';
  if (GetControlText('CN4') = 'X') then StRupt := 'PHB_TRAVAILN4';
  if (GetControlText('CN5') = 'X') then StRupt := 'PHB_CODESTAT';
  if (GetControlText('CETAB') = 'X') then StRupt := 'PHB_ETABLISSEMENT';

  if StRupt <> '' then AddChamp := StRupt + ',';
  if (GetControlText('CSAL') = 'X') then
  begin
    StSal := 'PHB_SALARIE';
    AddChamp := AddChamp + 'PHB_SALARIE,PSA_LIBELLE,PSA_PRENOM,';
  end;
  if (GetControlText('CALPHA') = 'X') then
  begin
    StAlpha := 'PSA_LIBELLE';
    AddChamp := AddChamp + 'PHB_SALARIE,PSA_LIBELLE,PSA_PRENOM,';
  end;

  if (StRupt <> '') and (StSal = '') and (StAlpha = '') then St := StRupt;
  if (StRupt = '') and (StSal <> '') and (StAlpha = '') then St := StSal;
  if (StRupt = '') and (StSal = '') and (StAlpha <> '') then St := StAlpha;

  if (StRupt <> '') and (StSal <> '') and (StAlpha <> '') then St := StRupt + ',' + StAlpha + ',' + StSal;

  if (StRupt <> '') and (StSal <> '') and (StAlpha = '') then St := StRupt + ',' + StSal;
  if (StRupt = '') and (StSal <> '') and (StAlpha <> '') then St := StAlpha + ',' + StSal;
  if (StRupt <> '') and (StSal = '') and (StAlpha <> '') then St := StRupt + ',' + StAlpha;

  if St <> '' then result := St + ',PHB_ORDREETAT,PHB_NATURERUB,PHB_RUBRIQUE';
  if st = '' then result := 'PHB_ORDREETAT,PHB_NATURERUB,PHB_RUBRIQUE';
end;

procedure TOF_PGBASCULEURO_ETAT.OnUpdate;
var
  Temp, Tempo, Critere: string;
  Pages: TPageControl;
  x: integer;
  DateDeb, DateFin: TDateTime;
  Order, join: string;
begin
  Pages := TPageControl(GetControl('Pages'));
  Temp := RecupWhereCritere(Pages);
  tempo := '';
  critere := '';
  x := Pos('(', Temp);
  if x > 0 then Tempo := copy(Temp, x, (Length(temp) - 5));
  if tempo <> '' then critere := 'AND ' + Tempo;

  if IsValidDate(GetControlText('XX_VARIABLEDEB')) then
    Datedeb := StrToDate(GetControlText('XX_VARIABLEDEB'))
  else
    Datedeb := idate1900;

  if IsValidDate(GetControlText('XX_VARIABLEFIN')) then
    DateFin := StrToDate(GetControlText('XX_VARIABLEFIN'))
  else
    DateFin := idate1900;

  if (GetControlText('CSAL') = 'X') then Join := 'LEFT JOIN SALARIES ON PSA_SALARIE=PHB_SALARIE '
  else join := '';

  Order := CreateOrderBy;

  TFQRS1(Ecran).WhereSQL := 'SELECT DISTINCT ' + AddChamp + 'PHB_RUBRIQUE,' +
    'PHB_LIBELLE,PHB_ORDREETAT,PHB_NATURERUB,' +
    'SUM(PHB_MTREM) AS MTREM,SUM(PHB_OMTREM) AS OMTREM,' +
    'SUM(PHB_MTSALARIAL) AS MTSALARIAL,SUM(PHB_OMTSALARIAL) AS OMTSALARIAL ' +
    'FROM HISTOBULLETIN ' +
    '' + Join + ' ' +
    'WHERE PHB_DATEDEBUT>="' + UsDateTime(DateDeb) + '"  ' +
    'AND PHB_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
    'AND ((PHB_MTREM+PHB_MTSALARIAL>0 AND PHB_IMPRIMABLE="X") ' +
    'AND PHB_NATURERUB<>"BAS") ' +
    '' + critere + ' ' +
    'GROUP BY ' + AddChamp + 'PHB_RUBRIQUE,PHB_LIBELLE,PHB_ORDREETAT,PHB_NATURERUB ' +
    'ORDER BY ' + order;
end;


procedure TOF_PGBASCULEURO_ETAT.ChangeLieuTravail(Sender: TObject);
begin
  RecupChampRupture(Ecran);
  BloqueChampLibre(Ecran);
end;

procedure TOF_PGBASCULEURO_ETAT.ControlPeriodeDeb(Sender: TObject);
var
  DebExer: TDateTime;
  Q: TQuery;
begin
  Q := OpenSql('SELECT PEX_DATEDEBUT FROM EXERSOCIAL WHERE PEX_EXERCICE="' + GetControlText('EDEXERSOC') + '" ', True);
  if not Q.eof then
    DebExer := Q.findfield('PEX_DATEDEBUT').AsDateTime
  else
    DebExer := idate1900;
  Ferme(Q);
  if StrToDate(GetControlText('XX_VARIABLEDEB')) < (DebExer) then
  begin
    PgiBox('La date de début ne peut être inferieur au ' + DateToStr(DebExer) + '', 'Date Erronée!');
    SetControlText('XX_VARIABLEDEB', DateToStr(DebExer));
  end;
end;

procedure TOF_PGBASCULEURO_ETAT.ControlPeriodeFin(Sender: TObject);
var
  FinPer, FinExer: TDateTime;
  Q: TQuery;
begin
  Q := OpenSql('SELECT PEX_DATEFIN FROM EXERSOCIAL WHERE PEX_EXERCICE="' + GetControlText('EDEXERSOC') + '" ', True);
  if not Q.eof then
    FinExer := Q.findfield('PEX_DATEFIN').AsDateTime
  else
    FinExer := idate1900;
  Ferme(Q);
  if IsValidDate(GetControlText('XX_VARIABLEFIN')) then
    FinPer := StrToDate(GetControlText('XX_VARIABLEFIN'))
  else FinPer := idate1900;
  if FinPer > VH_Paie.PGDateBasculEuro then {PT-1 FinPer au lieu de FinExer}
  begin
    FinPer := VH_Paie.PGDateBasculEuro;
    PgiBox('La date de fin ne peut être supérieur à la date de basculement en euro.', 'Date Erronée!');
    SetControlText('XX_VARIABLEFIN', DateToStr(FinPer));
    Exit;
  end;
  if FinPer > FinExer then
  begin
    PgiBox('La date de fin ne peut être supérieur au ' + DateToStr(FinExer) + '', 'Date Erronée!');
    SetControlText('XX_VARIABLEFIN', DateToStr(FinExer));
  end;
end;


procedure TOF_PGBASCULEURO_ETAT.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if edit.text <> '' then
      if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
        edit.text := AffectDefautCode(edit, 10);
end;

initialization
  registerclasses([TOF_PGBASCULEURO_ETAT]);
end.

 

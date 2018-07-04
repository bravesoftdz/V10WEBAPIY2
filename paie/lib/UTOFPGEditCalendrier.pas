{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 12/09/2001
Modifié le ... :   /  /
Description .. : Edition du calendrier absence et temps de travail
Mots clefs ... : PAIE;CALENDRIER
*****************************************************************
PT1 29/01/2002 V571 SB Idate1900 au lieu de 01/01/1900
PT2 18/04/2002 V571 SB Fiche de bug n°10087 : V_PGI_env.LibDossier non renseigné en Mono}
unit UTOFPGEditCalendrier;

interface
uses  Controls, Classes, sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
  eQRS1, UTOB,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} QRS1,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF,
  ParamSoc, HQry, ParamDat;

type
  TOF_PGCALENDRIER_ETAT = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
    procedure Change(Sender: TObject);
    procedure InitJourMois(Sender: TObject);
    procedure AffectDateCumul(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
    procedure DateElipsisclick(Sender: TObject);
  end;

implementation

uses PgEditOutils2, EntPaie, PgOutils,PGoutils2, PGEdtEtat; 

{ TOF_PGCALENDRIER_ETAT }

procedure TOF_PGCALENDRIER_ETAT.AffectDateCumul(Sender: TObject);
var
  DateFin, DateCumul: TdateTime;
begin
  if GetCOntrolText('DATEFIN') <> '' then DateFin := StrToDate(GetCOntrolText('DATEFIN'))
  else DateFin := Idate1900;
  if GetControltext('DEBREPRISE') <> '' then DateCumul := StrToDate(GetControltext('DEBREPRISE'))
  else DateCumul := IDate1900;
  if (DateCumul > DateFin) and (DateCumul <> IDate1900) and (DateFin <> IDate1900) then
  begin
    PGIBox('Vous devez saisir une date inférieur au ' + DateToStr(DateFin) + '!', 'Date debut cumul');
    SetControltext('DEBREPRISE', DateToStr(DateFin));
  end;
  SetControlText('DEBREPRISEMOIS2', GetControltext('DEBREPRISE'));
end;

procedure TOF_PGCALENDRIER_ETAT.Change(Sender: TObject);
begin

end;

procedure TOF_PGCALENDRIER_ETAT.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;


procedure TOF_PGCALENDRIER_ETAT.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

procedure TOF_PGCALENDRIER_ETAT.InitJourMois(Sender: TObject);
var
  Mois, Mois2, Annee, AnneeOk: string;
  JJ, MM, YY: WORD;
  Q: TQuery;
  DateDebutMois1, DateFinMois1, DateDebutMois2, DateFinMois2: TDateTime;
begin
  Mois := GetControlText('CBMOIS');
  Annee := GetControlText('CBANNEE');
  if (mois = '') or (annee = '') then exit;
  ControlMoisAnneeExer(Mois, RechDom('PGANNEESOCIALE', Annee, FALSE), AnneeOk);
  Annee := AnneeOk;
  if Annee <> '' then
  begin
    DateDebutMois1 := EncodeDate(StrToInt(Annee), StrToInt(Mois), 1);
    DateFinMois1 := FindeMois(DateDebutMois1);
    DateDebutMois2 := PlusMois(DateDebutMois1, 1);
    DateFinMois2 := FinDeMois(DateDebutMois2);
  end
  else
  begin
    DateDebutMois1 := iDate1900;
    DateFinMois1 := iDate1900;
    //DateDebutMois2:=iDate1900;
    DateFinMois2 := iDate1900;
  end;

  SetControlText('DATEDEBUT', DateToStr(DateDebutMois1));
  SetControlText('DATEFIN', DateToStr(DateFinMois2));
  Decodedate(DateDebutMois1, YY, MM, JJ);
  SetControlText('ANNEEMOIS1', IntToStr(YY));
  Decodedate(DateFinMois2, YY, MM, JJ);
  SetControlText('ANNEEMOIS2', IntToStr(YY));
  if Mois <> '12' then Mois2 := IntToStr(StrToInt(Mois) + 1) else Mois2 := '01';
  if Length(Mois2) = 1 then Mois2 := '0' + Mois2;
  SetControlText('CBMOIS_', Mois2);

  Q := OpenSql('SELECT PEX_DATEDEBUT,PEX_DATEFIN FROM EXERSOCIAL ' +
    'WHERE PEX_DATEDEBUT<="' + USDateTime(DateDebutMois1) + '" ' +
    'AND PEX_DATEFIN>="' + USDateTime(DateFinMois1) + '" ', True);
  if not Q.eof then
    SetControlText('DEBREPRISE', DateToStr(Q.FindField('PEX_DATEDEBUT').AsDateTime))
  else
    SetControlText('DEBREPRISE', DateToStr(Date));
  Ferme(Q);

  SetControlText('DEBREPRISEMOIS2', GetControltext('DEBREPRISE'));
  {Q:=OpenSql('SELECT PEX_DATEDEBUT,PEX_DATEFIN FROM EXERSOCIAL '+
     'WHERE PEX_DATEDEBUT<="'+USDateTime(DateDebutMois2)+'" '+
     'AND PEX_DATEFIN>="'+USDateTime(DateFinMois2)+'" ',True);
  If Not Q.eof then
    SetControlText('DEBREPRISEMOIS2',DateToStr(Q.FindField('PEX_DATEDEBUT').AsDateTime))
  else
    SetControlText('DEBREPRISEMOIS2',DateToStr(Date));
  Ferme(Q);      }

end;

procedure TOF_PGCALENDRIER_ETAT.OnArgument(Arguments: string);
var
  Min, Max, Mois, ExerPerEncours, DebPer, FinPer: string;
  YY, MM, JJ: Word;
  Combo: THValComboBox;
  Defaut: ThEdit;
  Zone: Tcontrol;
begin
  inherited;
  Combo := THValComboBox(GetControl('CBMOIS'));
  if Combo <> nil then Combo.OnChange := InitJourMois;
  Combo := THValComboBox(GetControl('CBANNEE'));
  if Combo <> nil then Combo.OnChange := InitJourMois;
  Zone := ThValComboBox(getcontrol('PSA_CALENDRIER'));
  InitialiseCombo(Zone);
  //Valeur par défaut
  RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
  Defaut := ThEdit(getcontrol('PSA_SALARIE'));
  if Defaut <> nil then
  begin
    Defaut.text := Min;
    Defaut.OnExit := ExitEdit;
  end;
  Defaut := ThEdit(getcontrol('PSA_SALARIE_'));
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
  Defaut := ThEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
    //  Defaut.text:=V_PGI_env.LibDossier;    //PT2 Mise en commentaire
    Defaut.text := GetParamSoc('SO_LIBELLE');
  RendPeriodeEnCours(ExerPerEncours, DebPer, FinPer);
  DecodeDate(StrToDate(DebPer), YY, MM, JJ);
  Mois := IntToStr(MM);
  if Length(Mois) = 1 then Mois := '0' + Mois;
  SetControlproperty('CBMOIS', 'Value', Mois);
  SetControlproperty('CBANNEE', 'Value', ExerPerEncours);
  if Ecran <> nil then PGCalendrierEtat_Ecran := Ecran;
  defaut := THedit(GetControl('DEBREPRISE'));
  if defaut <> nil then
  begin
    Defaut.OnExit := AffectDateCumul;
    Defaut.OnElipsisClick := DateElipsisclick;
  end;

end;

procedure TOF_PGCALENDRIER_ETAT.OnUpdate;
var
  SQL, Temp, Tempo, Critere: string;
  DD, DF: TDateTime;
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
  DD := StrToDate(GetControlText('DATEDEBUT'));
  DF := StrToDate(GetControlText('DATEFIN'));
  if (DD > 2) and (DF > 2) then
    SQL := 'SELECT DISTINCT PSA_SALARIE,PSA_LIBELLE, PSA_ETABLISSEMENT,PSA_STANDCALEND,' +
      'PSA_CALENDRIER,PSA_PRENOM,PSA_DATEENTREE,PSA_DATESORTIE,PSA_JOURHEURE ' +
      'FROM SALARIES ' +
      'WHERE PSA_DATEENTREE<="' + USDateTime(DF) + '" ' + //PT1
    'AND (PSA_DATESORTIE>="' + USDateTime(DD) + '" OR PSA_DATESORTIE="' + UsdateTime(Idate1900) + '" OR PSA_DATESORTIE is null) ' +
      '' + critere + ' ' +
      'ORDER BY PSA_SALARIE ';
  TFQRS1(Ecran).WhereSQL := SQL;
end;

initialization
  registerclasses([TOF_PGCALENDRIER_ETAT]);
end.


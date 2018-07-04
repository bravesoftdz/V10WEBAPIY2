{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 29/07/2004
Modifié le ... :   /  /
Description .. : Edition du bilan social
Mots clefs ... : PAIE;BILAN SOCIAL
*****************************************************************
}
unit UTOFPGEditBilanSocial;

interface
uses Classes,sysutils,
  {$IFDEF EAGLCLIENT}
  UTOB,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, ParamDat,
  ParamSoc;



type

  TOF_PGEDITBILANSOCIAL = class(TOF)
    procedure OnArgument(Arguments: string); override;
  private
    ExDateDeb, ExDateFin: string;
    procedure DateElipsisclick(Sender: TObject);
    procedure ChangeExercice(Sender: TObject);
    procedure PgIFValidPeriode(Sender: TObject);
  end;
implementation

{-------------------------------------------------------------------------------
                     TOF  FICHE INDIVIDUEL
--------------------------------------------------------------------------------}
procedure TOF_PGEDITBILANSOCIAL.OnArgument(Arguments: string);
var
  THDateDeb, THDateFin : THEdit;
  Exercice: THValComboBox;
  QPeriode, QExer: TQuery;        
begin
  SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
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
    THDateDeb.OnExit := PgIFValidPeriode;
    THDateFin.text := ExDateFin;
    THDateFin.OnElipsisClick := DateElipsisclick;
    THDateFin.OnExit := PgIFValidPeriode;
    Exercice.OnChange := ChangeExercice;
  end;
end;

procedure TOF_PGEDITBILANSOCIAL.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGEDITBILANSOCIAL.ChangeExercice(Sender: TObject);
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

procedure TOF_PGEDITBILANSOCIAL.PgIFValidPeriode(Sender : TObject);
var
  YYD, MMD, JJ, YYF, MMF: WORD;
begin
  if IsValidDate(GetControlText('XX_VARIABLEDEB')) and IsValidDate(GetControlText('XX_VARIABLEFIN')) then
  begin
    DecodeDate(StrToDate(GetControlText('XX_VARIABLEDEB')), YYD, MMD, JJ);
    DecodeDate(StrToDate(GetControlText('XX_VARIABLEFIN')), YYF, MMF, JJ);
    if (YYF > YYD) and (MMF >= MMD) then
    begin
      PgiBox('La période d''édition ne peut excéder douze mois civils.', 'Date Erronée!'); { PT10-1 04/05/2004 }
      SetControlText('XX_VARIABLEFIN', DateToStr(FinDeMois(PlusDate(StrToDate(GetControlText('XX_VARIABLEDEB')), 11, 'M'))));
    end;
  end;
end;

initialization
  registerclasses([TOF_PGEDITBILANSOCIAL]);
end.


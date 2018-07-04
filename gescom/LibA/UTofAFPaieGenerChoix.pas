{***********UNITE*************************************************
Auteur  ...... : TG
Créé le ...... : 30/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFPaieGenerChoix ()
Mots clefs ... : TOF;AFPaieGenerChoix
*****************************************************************}
unit UTofAFPaieGenerChoix;

interface

uses StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
  db,
  dbtables,
  {$ELSE}
  utob,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  DicoAF,
  UTOF,FileCtrl;

type
  TOF_AFPaieGenerChoix = class(TOF)
    procedure OnArgument(S: string); override;
    procedure OnUpdate; override;
    procedure OnClose; override;
  private
    DontClose: Boolean;
    DebExer, FinExer: TDateTime;
    CBExercice: THValCombobox;
    procedure CBExerciceExit(Sender: TObject);
    function ControlDate: boolean;
  end;

implementation
uses Vierge, pgienv;

const TexteMessage: array[1..1] of string = (
    {1}'Les dates de périodes ne correspondent pas à l''exercice ');

procedure TOF_AFPaieGenerChoix.OnArgument(S: string);
var StChemin : string;
begin
  inherited;
  DontClose := false;
  CBExercice := THValCombobox(Getcontrol('EXERCICESOC'));
  CBExercice.OnExit := CBExerciceExit;
  CBExercice.ItemIndex := CBExercice.Items.count - 1;
  CBExerciceExit(nil);
  StChemin := ExtractFilePath(ParamStr(0));
  if (StChemin = '') or not DirectoryExists(StChemin) then
    StChemin := 'C:\PGI00\';
  SetControlText('CHEMINDEST', StChemin);
end;

procedure TOF_AFPaieGenerChoix.OnUpdate;
var S: string;
begin
  inherited;
  if ControlDate then
  begin
    S := 'EXERCICE='+ GetControlText('EXERCICESOC')+ ','+THValCombobox(GetControl('EXERCICESOC')).text+ ';' +
      'PERIODE='+ GetControlText('DATEPERIODE')+ ',' + GetControlText('DATEPERIODE_') + ';' +
      'ACTIVITE='+ GetControlText('DATEPERIODEACT') + ',' + GetControlText('DATEPERIODEACT_') + ';' +
      'DEST=' + GetControlText('CHEMINDEST');
  end
  else S := '';

  TFVierge(Ecran).Retour := S;
end;

procedure TOF_AFPaieGenerChoix.CBExerciceExit(Sender: TObject);
var StSql: string;
  Q: TQuery;
  DatF,DatFAct: TDateTime;
begin
  StSql := 'SELECT * FROM EXERSOCIAL WHERE PEX_ACTIF="X" ';
  if trim(CBExercice.value) <> '' then
    StSql := StSql + ' AND PEX_EXERCICE = "' + CBExercice.value + '" ';
  StSql := StSql + ' ORDER BY PEX_ANNEEREFER DESC';
  Q := OpenSQL(StSql, TRUE);
  if not Q.EOF then
  begin
    DatF := Q.FindField('PEX_FINPERIODE').AsFloat;
    DebExer := Q.FindField('PEX_DATEDEBUT').AsDateTime;
    FinExer := Q.FindField('PEX_DATEFIN').AsDateTime;
    SetControlText('EXERCICESOC', Q.FindField('PEX_EXERCICE').AsString);
    SetControlText('DATEPERIODE', DateToStr(DebutDeMois(DatF)));
    SetControlText('DATEPERIODE_', DateToStr(DatF));
    DatFAct := DatF;
    //DatFAct := PlusMois (DatF,-1); paye décalé
    SetControlText('DATEPERIODEACT', DateToStr(DebutDeMois(DatFAct)));
    SetControlText('DATEPERIODEACT_', DateToStr(FinDeMois(DatFAct)));
  end
  else
  begin
    SetControlText('DATEPERIODE', DateToStr(DebutDeMois(now)));
    SetControlText('DATEPERIODE_', DateToStr(FinDeMois(now)));
    SetControlText('DATEPERIODEACT', DateToStr(DebutDeMois(now)));
    SetControlText('DATEPERIODEACT_', DateToStr(FinDeMois(now)));
    DebExer := iDate1900;
    FinExer := iDate2099;
  end;
  Ferme(Q);
end;

procedure TOF_AFPaieGenerChoix.OnClose;
begin
  if not ControlDate then
  begin
    LastError := 1;
    PGIInfoAf(textemessage[LastError], Ecran.Caption);
    exit;
  end
end;

function TOF_AFPaieGenerChoix.ControlDate: boolean;
var vDtDeb, vDtFin: TDateTime;
begin
  vDtDeb := strtodate(getcontroltext('DATEPERIODE'));
  vDtFin := strtodate(getcontroltext('DATEPERIODE_'));
  result := not ((vDtDeb < DebExer) or (vDtFin > FinExer));
end;

initialization
  registerclasses([TOF_AFPaieGenerChoix]);
end.

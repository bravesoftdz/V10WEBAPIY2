{***********UNITE*************************************************
Auteur  ...... : TG
Créé le ...... : 30/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPaieGenerChoix ()
Mots clefs ... : TOF;BTPaieGenerChoix
*****************************************************************}
unit UTofBTPaieGenerChoix;

interface

uses StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ELSE}
  utob,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  Dicobtp,
  UTOF,
  CBPPath,
  FileCtrl;

type
  TOF_BTPaieGenerChoix = class(TOF)
    procedure OnArgument(S: string); override;
    procedure OnUpdate; override;
    procedure OnClose; override;
  private
    DontClose: Boolean;
    DebExer, FinExer: TDateTime;
    CBExercice: THValCombobox;
  	DatF,DatFAct: TDateTime;
    procedure CBExerciceExit(Sender: TObject);
    function ControlDate: boolean;
    procedure ChangeTypeLiaison (Sender : Tobject);
  end;

implementation
uses Vierge;

const TexteMessage: array[1..1] of string = (
    {1}'Les dates de périodes ne correspondent pas à l''exercice ');

procedure TOF_BTPaieGenerChoix.OnArgument(S: string);
var StChemin : string;
begin
  inherited;
  TgroupBox(getControl('GPPAIEDECAL')).Enabled := false;

  DontClose := false;
  CBExercice := THValCombobox(Getcontrol('EXERCICESOC'));
  CBExercice.OnExit := CBExerciceExit;
  CBExercice.ItemIndex := CBExercice.Items.count - 1;
  CBExerciceExit(nil);

  TCheckBox(GetControl('CPAIEDECAL')).OnClick := ChangeTypeLiaison;

  StChemin := ExtractFilePath(ParamStr(0));
  if (StChemin = '') or not DirectoryExists(StChemin) then
     //Modif FV passage Version 5 - Version 8
     //StChemin := 'C:\PGI00\';
     StChemin := IncludeTrailingBackSlash(TCBPPath.GetCegidDataDistri);

  SetControlText('CHEMINDEST', StChemin);

end;

procedure TOF_BTPaieGenerChoix.OnUpdate;
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

procedure TOF_BTPaieGenerChoix.CBExerciceExit(Sender: TObject);
var StSql: string;
  Q: TQuery;
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
    //
    if TcheckBox(GetCOntrol('CPAIEDECAL')).Checked then
    begin
      DatFAct := DatF;
      DatFAct := PlusMois (DatF,-1); // paye décalé
      SetControlText('DATEPERIODEACT', DateToStr(DebutDeMois(DatFAct)));
      SetControlText('DATEPERIODEACT_', DateToStr(FinDeMois(DatFAct)));
    end else
    begin
      DatFAct := DatF;
      SetControlText('DATEPERIODEACT', DateToStr(DebutDeMois(DatFAct)));
      SetControlText('DATEPERIODEACT_', DateToStr(FinDeMois(DatFAct)));
    end;
  end
  else
  begin
    Datf := now;
    DatFact := now;
    TcheckBox(GetCOntrol('CPAIEDECAL')).Checked := false;
    //
    SetControlText('DATEPERIODE', DateToStr(DebutDeMois(now)));
    SetControlText('DATEPERIODE_', DateToStr(FinDeMois(now)));
    SetControlText('DATEPERIODEACT', DateToStr(DebutDeMois(now)));
    SetControlText('DATEPERIODEACT_', DateToStr(FinDeMois(now)));
    DebExer := iDate1900;
    FinExer := iDate2099;
  end;
  Ferme(Q);
end;

procedure TOF_BTPaieGenerChoix.OnClose;
begin
  if not ControlDate then
  begin
    LastError := 1;
    PGIInfoAf(textemessage[LastError], Ecran.Caption);
    exit;
  end
end;

function TOF_BTPaieGenerChoix.ControlDate: boolean;
var vDtDeb, vDtFin: TDateTime;
begin
  vDtDeb := strtodate(getcontroltext('DATEPERIODE'));
  vDtFin := strtodate(getcontroltext('DATEPERIODE_'));
  result := not ((vDtDeb < DebExer) or (vDtFin > FinExer));
end;

procedure TOF_BTPaieGenerChoix.ChangeTypeLiaison(Sender: Tobject);
begin
  if TcheckBox(GetCOntrol('CPAIEDECAL')).Checked then
  begin
    TgroupBox(getControl('GPPAIEDECAL')).Enabled := true;
    DatFAct := DatF;
    DatFAct := PlusMois (DatF,-1); // paye décalé
    SetControlText('DATEPERIODEACT', DateToStr(DebutDeMois(DatFAct)));
    SetControlText('DATEPERIODEACT_', DateToStr(FinDeMois(DatFAct)));
  end else
  begin
    DatFAct := DatF;
    SetControlText('DATEPERIODEACT', DateToStr(DebutDeMois(DatFAct)));
    SetControlText('DATEPERIODEACT_', DateToStr(FinDeMois(DatFAct)));
    TgroupBox(getControl('GPPAIEDECAL')).Enabled := false;
  end;
end;

initialization
  registerclasses([TOF_BTPaieGenerChoix]);
end.

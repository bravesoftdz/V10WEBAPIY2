{***********UNITE*************************************************
Auteur  ...... : GF
Créé le ...... : 06/11/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : HRSERVICES ()
Mots clefs ... : TOF;HRSERVICES
*****************************************************************}
unit BTSERVICES_TOF;

interface

uses
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}

{$ENDIF}
  Vierge,
  forms,
  sysutils,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  UTob,
  M3FP;

type
  TOF_BTSERVICES = class(TOF)
    procedure OnNew; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnDelete; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  private
    indice_tob: integer;
    Action: TActionFiche;
    TOBServices: TOB;
    OKFermer: boolean;
    procedure ChargeLaTob();
    procedure ChargeLafiche();
    procedure MAJTobServices();
    function ZonesOk(): boolean;
    procedure ChangerIndice(stind: string);
  end;

implementation

procedure TOF_BTSERVICES.OnNew;
begin
  inherited;
  SetControlEnabled('CC_CODE',True);
  SetFocusControl('CC_CODE');
  SetControlText('CC_CODE', '');
  SetControlText('CC_LIBELLE', '');
  SetControlText('CC_ABREGE', '');
  SetControlText('CC_LIBRE', '');
  SetControlText('HEUREDEB', '');
  SetControlText('HEUREFIN', '');
  Action := TaCreat;
end;

procedure TOF_BTSERVICES.OnUpdate;
begin
  inherited;
  MAJTobServices;
  if ZonesOk then
  begin
    TOBServices.InsertOrUpdateDB(True); {MAJ table}
    TOBServices.free;
    TOBServices := TOB.Create('CHOIXCOD',nil, -1);
    ChargeLatob;
    OKFermer := True;
  end
  else
    OKFermer := False;
end;

procedure TOF_BTSERVICES.OnLoad;
begin
  inherited;
  ChargeLaTob;
  if (TobServices.detail.count <> 0) then
    Action := TaModif
  else
    Action := TaCreat;
end;

procedure TOF_BTSERVICES.OnArgument(S: string);
begin
  inherited;
  OkFermer := True;
  {Init TOB}
  TOBServices := TOB.Create('CHOIXCOD',nil, -1);
  {Init CC_TYPE}
  SetControltext('CC_TYPE', 'BSE');
  indice_tob := 0;
end;

procedure TOF_BTSERVICES.OnClose;
begin
  inherited;
  if not OKFermer then
  begin
    LastError := -1;
    OKFermer := True;
    Exit;
  end
  else
    {Libere la tob}
    TOBServices.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Guillaume FONTANA
Créé le ...... : 26/04/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOF_BTSERVICES.OnDelete;
var
  StSQL: string;
  TOBTemp: TOB;
  LeCode: string;
  Reponse: word;
begin
  inherited;

  LeCode := GetcontrolText('CC_CODE');
  if LeCode = '' then
    Exit;
  {Vérifie que le code n'est pas utilisé dans HRDOSRES}
  StSQL := 'SELECT HDR_DOSRES FROM HRDOSRES WHERE (HDR_HRPERIODEDEBUT="' + GetControlText('CC_CODE') + '")'
    + ' OR (HDR_HRPERIODEFIN="' + GetControlText('CC_CODE') + '")';
  if not ExisteSQL(StSQL) then
  begin
    {Message de confirmation}
    Reponse := PGIAsk('Confirmez-vous la suppression de l''enregistrement ?',Ecran.Caption + ' : ' + LeCode);
    if Reponse = mrYes then
    begin
      TOBTemp := TOBServices.FindFirst(['CC_CODE', 'CC_TYPE'], [GetControlText('CC_CODE'), 'BSE'],True);
      if Assigned(TOBtemp) then
      begin
        try
          {Efface l'enreg. dans la table}
          BEGINTRANS;
          ExecuteSQL('DELETE FROM CHOIXCOD WHERE CC_TYPE="BSE" AND CC_CODE="' + LeCode + '"');
          COMMITTRANS;
        except
          ROLLBACK;
        end;
        TOBTemp.Free;
        {Placement sur l'enregistrement le plus proche}
        if (indice_tob - 1 < TOBServices.Detail.Count) and (TOBServices.Detail.Count <> 0) then
        begin
          if indice_tob - 1 >= 0 then
            ChangerIndice('DEC')
          else
            ChangerIndice('PRE');
        end
        else
        begin
          {Mode Creation si plus de services}
          indice_tob := 0;
          OnNew;
        end;
      end;
    end;
  end
  else
    PGIError('Ce service est utilisé.', 'Suppression impossible.');
end;

procedure TOF_BTSERVICES.ChangerIndice(stind: string);
begin
  if (Stind = 'DEC') then
    if (indice_tob >= 1) then
      indice_tob := indice_tob - 1;
  if (Stind = 'INC') then
    if (indice_tob < TobServices.detail.count - 1) then
      indice_tob := indice_tob + 1;
  if (Stind = 'PRE') then
    indice_tob := 0;
  if (Stind = 'DER') then
    indice_tob := TOBServices.detail.count - 1;
  chargelafiche;
end;

procedure TOF_BTSERVICES.ChargeLaTob;
var
  Q: TQuery;
  StSQL: string;
begin
  StSQl := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="BSE"';
  Q := OpenSQL(StSQL,True);
  if not Q.Eof then
    TOBservices.LoadDetailDB('LES SERVICES', '', '',Q,True,True);
  Ferme(Q);
  if (TOBservices.detail.count <> 0) then
    ChargeLafiche
  else
  begin
    SetControlEnabled('CC_CODE',True);
    SetFocusControl('CC_CODE');
  end;
end;

procedure TOF_BTSERVICES.ChargeLafiche;
begin
  SetControlText('CC_CODE',TOBservices.detail[indice_tob].Getvalue('CC_CODE'));
  SetControlEnabled('CC_CODE',False);
  SetControlText('CC_LIBELLE',TOBservices.detail[indice_tob].Getvalue('CC_LIBELLE'));
  SetControlText('CC_ABREGE',TOBservices.detail[indice_tob].Getvalue('CC_ABREGE'));
  if (TOBservices.detail[indice_tob].Getvalue('CC_LIBRE') <> '') then
  begin
    SetControlText('HEUREDEB',Copy(TOBservices.detail[indice_tob].Getvalue('CC_LIBRE'),1,5));
    SetControlText('HEUREFIN',Copy(TOBservices.detail[indice_tob].Getvalue('CC_LIBRE'),7,5));
  end
  else
  begin
    SetControlText('HEUREDEB', '00:00');
    SetControlText('HEUREFIN', '00:00');
  end;
  SetControlText('CC_LIBRE',TOBservices.detail[indice_tob].Getvalue('CC_LIBRE'));
end;

procedure TOF_BTSERVICES.MAJTobServices;
var
  stheure: string;
begin
  stheure := GetControlText('HEUREDEB') + ' ' + GetControlText('HEUREFIN');
  SetControlText('CC_LIBRE',StHeure);
  TOBservices.GetEcran(Ecran);
end;

function TOF_BTSERVICES.ZonesOk: boolean;
var
  StSQL: string;
  TempHeure: TTime;
begin
  TempHeure := iDate1900;
  result := True;
  if (GetControltext('CC_CODE') = '') then
  begin
    PGIError('Veuillez saisir un code', 'Validation impossible');
    SetFocusControl('CC_CODE');
    Result := False;
    Exit;
  end;
  StSQL := 'SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="BSE"'
    + ' AND CC_CODE="' + GetControltext('CC_CODE') + '"';
  if (ExisteSQL(StSQL)) and (Action = TaCreat) then
  begin
    PGIError('Ce code existe déja', 'Validation impossible');
    Result := False;
    SetControlText('CC_CODE', '');
    SetFocusControl('CC_CODE');
    Exit;
  end;
  if (GetControltext('CC_LIBELLE') = '') then
  begin
    PGIError('Veuillez renseigner le libellé', 'Validation impossible');
    SetFocusControl('CC_LIBELLE');
    Result := False;
    Exit;
  end;
  if (GetControltext('HEUREDEB') <> '') then
  begin
    try
      TempHeure := StrToTime(GetControltext('HEUREDEB'));
    except
      on E: EConvertError do
      begin
        HShowMessage('0;Zone mal renseignée;L''heure saisie n''est pas valide.;E;O;O;O;', '', '');
      end;
    end;
    SetControlText('HEUREDEB',FormatDateTime('hh:nn',TempHeure));
  end;
  if (GetControltext('HEUREFIN') <> '') then
  begin
    try
      TempHeure := StrToTime(GetControltext('HEUREFIN'));
    except
      on E: EConvertError do
      begin
        HShowMessage('0;Zone mal renseignée;L''heure saisie n''est pas valide.;E;O;O;O;', '', '');
      end;
    end;
    SetControlText('HEUREFIN',FormatDateTime('hh:nn',TempHeure));
  end;
end;

procedure ScriptDecrementer_indicetob(Parms: array of variant; nb: integer);
var
  Totof: TOF;
  F: TForm;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then
    ToTof := TFVierge(F).LaTof
  else
    exit;
  TOF_BTSERVICES(totof).changerindice('DEC');
end;

procedure ScriptIncrementer_indicetob(Parms: array of variant; nb: integer);
var
  Totof: TOF;
  F: TForm;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then
    ToTof := TFVierge(F).LaTof
  else
    exit;
  TOF_BTSERVICES(totof).changerindice('INC');
end;

procedure ScriptDernier_indicetob(Parms: array of variant; nb: integer);
var
  Totof: TOF;
  F: TForm;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then
    ToTof := TFVierge(F).LaTof
  else
    exit;
  TOF_BTSERVICES(totof).changerindice('DER');
end;

procedure ScriptPremier_indicetob(Parms: array of variant; nb: integer);
var
  Totof: TOF;
  F: TForm;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then
    ToTof := TFVierge(F).LaTof
  else
    exit;
  TOF_BTSERVICES(totof).changerindice('PRE');
end;

initialization
  RegisterClasses([TOF_BTSERVICES]);
  RegisterAglProc('Decrementer_indicetob',True,0,ScriptDecrementer_indicetob);
  RegisterAglProc('Incrementer_indicetob',True,0,ScriptIncrementer_indicetob);
  RegisterAglProc('Dernier_indicetob',True,0,ScriptDernier_indicetob);
  RegisterAglProc('Premier_indicetob',True,0,ScriptPremier_indicetob);
end.

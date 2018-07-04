{***********UNITE*************************************************
Auteur  ...... : MF
Créé le ...... : 03/10/2001
Modifié le ... : 03/10/2001
Description .. : Multicritère de sélection des Codes DUCS. Choix
Suite ........ : possible du type de DUCS
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
{
PT1 : 02/09/03  : V_421 surchage de la clause where du mul dans
                   le cas du XX_WHERE ==> Double condition ==> mauvais SQL

}
unit UTofPG_MulDucsParam;

interface
uses
  Classes, HCtrls,
{$IFDEF EAGLCLIENT}
  UtileAGL, eMul,
{$ELSE}
  mul,
{$ENDIF}
  UTOF, Hqry;
type
  TOF_PGMULDUCSPARAM = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  private
    WW: THEdit;
    Q_Mul: THQuery;
    NatureOrg: THValComboBox;
    procedure ActiveWhere(Okok: Boolean);
  end;
implementation

procedure TOF_PGMULDUCSPARAM.OnArgument(Arguments: string);
begin
  WW := THEdit(GetControl('XX_WHERE'));
  NatureOrg := THValComboBox(getcontrol('NATUREORG'));
end;

procedure TOF_PGMULDUCSPARAM.OnLoad;
var
  Okok: Boolean;
begin
  inherited;
  Okok := TRUE;
  ActiveWhere(Okok);
end;

procedure TOF_PGMULDUCSPARAM.ActiveWhere(Okok: Boolean);
var
  St, DebCod, FinCod: string;

begin
  WW.Text := '';
  st := '';

  DebCod := Copy(NatureOrg.Value, 1, 1);
  FinCod := DebCod + 'ZZZZZZ';
  if (DebCod = '') then FinCod := 'ZZZZZZZ';
  if (DebCod = '3') then FinCod := '5ZZZZZZ';
// PT1 : 02/09/03  : V_421 FQ 10784 surchage de la clause where du mul dans
  if DebCod <> '' then St := '((PDP_CODIFICATION >="' + DebCod + '")AND (PDP_CODIFICATION <="' + FinCod + '"))';
  if (DebCod = '3') then St := St + ' OR ((PDP_CODIFICATION >= "8") AND (PDP_CODIFICATION <= "9ZZZZZZ"))';
  if St <> '' then WW.Text := st;
  if Q_Mul <> nil then
  begin
    TFMul(Ecran).SetDBListe('PGMULDUCSPARAM');
  end;
end;

initialization
  registerclasses([TOF_PGMULDUCSPARAM]);
end.



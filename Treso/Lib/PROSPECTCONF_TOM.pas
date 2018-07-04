{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
08.10.001.001  02/08/07   JP   Création de l'unité : fiche des confidentialités
08.10.001.013  09/10/07   JP   On ne gère que les confidentialités en Tréso
--------------------------------------------------------------------------------------}
unit PROSPECTCONF_TOM;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL, UTob , eFiche,
  {$ELSE}
  HDB, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, Fiche,
  {$ENDIF}
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  Controls, Classes, SysUtils, HCtrls, HEnt1, UTOM;

type
  TOM_PROSPECTCONF = class(TOM)
    procedure OnNewRecord              ; override;
    procedure OnUpdateRecord           ; override;
    procedure OnAfterUpdateRecord      ; override;
    procedure OnArgument   (S : string); override;
    procedure OnDeleteRecord           ; override;
    procedure OnLoadRecord             ; override;
    procedure OnChangeField(F : TField); override;
  private
    FTypeConf : string;

    procedure InitAffichage;
    procedure ChangeTypeConf;
    procedure ChangeProduit;
    procedure VideControls(Ind : Char);
    procedure RazZones;
    function  GenereSQL : string;
  public
    procedure GenereSQLClick(Sender : TObject);
    procedure GetValeurClick(Sender : TObject);
    procedure OnEditKeyDown (Sender : TObject; var Key : Word; Shift : TShiftState);
    procedure EcranKeyDown  (Sender : TObject; var Key : Word; Shift : TShiftState);
  end;

procedure TRLanceFiche_Confidentialite(Range, Lequel, Arguments : string);

implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  {$ENDIF MODENT1}
  ComCtrls, HTB97, Ent1, HMsgBox, ULibConfidentialite, HQry, ParamSoc, TRTABCONF_TOF,
  Windows, ULibWindows;


{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_Confidentialite(Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('TR', 'TRFICCONFIDENTIEL', Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_PROSPECTCONF.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  InitAffichage;
  (GetControl('BGENERESQL') as TToolbarButton97).OnClick := GenereSQLClick;
  (GetControl('BTVAL1'    ) as TToolbarButton97).OnClick := GetValeurClick;
  (GetControl('BTVAL2'    ) as TToolbarButton97).OnClick := GetValeurClick;
  (GetControl('BTVAL3'    ) as TToolbarButton97).OnClick := GetValeurClick;
  {$IFDEF EAGLCLIENT}
  (GetControl('RTC_VAL1') as THEdit).OnKeyDown := OnEditKeyDown;
  (GetControl('RTC_VAL2') as THEdit).OnKeyDown := OnEditKeyDown;
  (GetControl('RTC_VAL3') as THEdit).OnKeyDown := OnEditKeyDown;
  {$ELSE}
  (GetControl('RTC_VAL1') as THDBEdit).OnKeyDown := OnEditKeyDown;
  (GetControl('RTC_VAL2') as THDBEdit).OnKeyDown := OnEditKeyDown;
  (GetControl('RTC_VAL3') as THDBEdit).OnKeyDown := OnEditKeyDown;
  {$ENDIF EAGLCLIENT}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_PROSPECTCONF.OnChangeField(F : TField);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if F.FieldName = 'RTC_TYPECONF' then ChangeTypeConf
  else if F.FieldName = 'RTC_PRODUITPGI' then ChangeProduit
  else if (F.FieldName = 'RTC_CHAMP1') then VideControls('1')
  else if (F.FieldName = 'RTC_CHAMP2') then VideControls('2')
  else if (F.FieldName = 'RTC_CHAMP3') then VideControls('3');
end;

{---------------------------------------------------------------------------------------}
procedure TOM_PROSPECTCONF.OnDeleteRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOM_PROSPECTCONF.OnNewRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;                       
  if (CtxTreso in V_PGI.PGIContexte) or EstComptaTreso then
    {09/10/07 : FQ 10529 : Pour le moment on met que Tréso
    SetField('RTC_PRODUITPGI', typ_ComptaTreso)}
    SetField('RTC_PRODUITPGI', typ_Treso)
  else
    SetField('RTC_PRODUITPGI', typ_Compta);
  SetField('RTC_TYPECONF'  , tyc_Banque);

  if not EstComptaTreso then SetControlEnabled('RTC_PRODUITPGI', False);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_PROSPECTCONF.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if VarToStr(GetField('RTC_TYPECONF')) = '' then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Veuillez renseigner le type de confidentialité');
    SetFocusControl('RTC_TYPECONF');
  end

  else if VarToStr(GetField('RTC_PRODUITPGI')) = '' then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Veuillez renseigner le produit');
    SetFocusControl('RTC_PRODUITPGI');
  end

  else if VarToStr(GetField('RTC_INTERVENANT')) = '' then begin
    LastError := 1;
    if GetParamSocSecur(NOMCONFIDENTIEL, False) then
      LastErrorMsg := TraduireMemoire('Veuillez renseigner le groupe utilisateur')
    else
      LastErrorMsg := TraduireMemoire('Veuillez renseigner l'' utilisateur');
    SetFocusControl('RTC_INTERVENANT');
  end

  {Vérification du SQL}
  else begin
    LastErrorMsg := GenereSQL;
    if LastErrorMsg <> '' then LastError := 1;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_PROSPECTCONF.OnLoadRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {Pour le moment on désactive le choix du produit et du type de confidentialité}
  SetControlEnabled('RTC_PRODUITPGI', False);
  SetControlEnabled('RTC_TYPECONF', False);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_PROSPECTCONF.OnAfterUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOM_PROSPECTCONF.InitAffichage;
{---------------------------------------------------------------------------------------}
begin
  FTypeConf := '';

  GereTabletteUser(GetControl('RTC_INTERVENANT'));
  GereCaptionUser(THLabel(GetControl('TRTC_INTERVENANT')));

  {$IFDEF EAGLCLIENT}
  THValComboBox(GetControl('RTC_PRODUITPGI')).Plus := FiltreTabletteProduit;
  THValComboBox(GetControl('RTC_TYPECONF'  )).Plus := FiltreTabletteTypeCnf('');
  {$ELSE}
  THDBValComboBox(GetControl('RTC_PRODUITPGI')).Plus := FiltreTabletteProduit;
  THDBValComboBox(GetControl('RTC_TYPECONF'  )).Plus := FiltreTabletteTypeCnf('');
  {$ENDIF EAGLCLIENT}

  Ecran.OnKeyDown := EcranKeyDown; 
end;

{---------------------------------------------------------------------------------------}
procedure TOM_PROSPECTCONF.ChangeProduit;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF EAGLCLIENT}
  THValComboBox(GetControl('RTC_TYPECONF'  )).Plus := FiltreTabletteTypeCnf(GetField('RTC_PRODUITPGI'));
  {$ELSE}
  THDBValComboBox(GetControl('RTC_TYPECONF'  )).Plus := FiltreTabletteTypeCnf(GetField('RTC_PRODUITPGI'));
  {$ENDIF EAGLCLIENT}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_PROSPECTCONF.ChangeTypeConf;
{---------------------------------------------------------------------------------------}
begin
  FTypeConf := VarToStr(GetField('RTC_TYPECONF'));
  {$IFDEF EAGLCLIENT}
  THValComboBox(GetControl('RTC_CHAMP1'  )).Plus := FiltreTableChamps(FTypeConf);
  THValComboBox(GetControl('RTC_CHAMP2'  )).Plus := FiltreTableChamps(FTypeConf);
  THValComboBox(GetControl('RTC_CHAMP3'  )).Plus := FiltreTableChamps(FTypeConf);
  {$ELSE}
  THDBValComboBox(GetControl('RTC_CHAMP1')).Plus := FiltreTableChamps(FTypeConf);
  THDBValComboBox(GetControl('RTC_CHAMP2')).Plus := FiltreTableChamps(FTypeConf);
  THDBValComboBox(GetControl('RTC_CHAMP3')).Plus := FiltreTableChamps(FTypeConf);
  {$ENDIF EAGLCLIENT}
end;

{---------------------------------------------------------------------------------------}
function TOM_PROSPECTCONF.GenereSQL : string;
{---------------------------------------------------------------------------------------}
var
  wSQL : string;
  tSQL : string;
  w    : string;
  n    : Byte;
begin
  Result := '';
  if (GetControlText('RTC_CHAMP1') = '') or (GetControlText('RTC_OPER1') = '') then Exit;

  {Constitution de la requête}
  wSQL := '';
  for n := 1 to 3 do begin
    if (GetControlText('RTC_CHAMP' + IntToStr(n)) = '') or
       (GetControlText('RTC_OPER' + IntToStr(n)) = '') then Break;
    {Récupération du SQL}
    w := CodeSQL(GetControlText('RTC_CHAMP' + IntToStr(n)),
                 GetControlText('RTC_OPER' + IntToStr(n)),
                 GetControlText('RTC_VAL' + IntToStr(n)));

    {Gestion d'un éventuel lien}
    w := '(' + w + ')';

    if (n > 1) and (GetControlText('RTC_LIEN' + IntToStr(n - 1)) <> '') and
       (w <> '') and (wSQL <> '') then begin
        w := ' ' + UpperCase(GetControlText('RTC_LIEN' + IntToStr(n - 1))) + ' ' + w;
    end;
    wSQL := wSQL + w;
    if (n < 3) and (GetControlText('RTC_LIEN' + IntToStr(n)) = '') then Break;
  end;

  {Mise à jour du Memo, même si la requête est fausse}
  if not (DS.State in [dsEdit, dsInsert]) then DS.Edit;
  SetField('RTC_SQLCONF', wSQL);
  SetControlText('RTC_SQLCONF', wSQL);

  if wSQL = '' then Exit;

  tSQL := 'SELECT ##TOP 1 ## "A" FROM ' + GetTableFromConf(GetField('RTC_TYPECONF'));
  tSQL := tSQL + ' WHERE ' + wSQL;

  {On teste la requête}
  try
    ExecuteSQL(tSQL);
  except
    on E : Exception do
      Result := TraduireMemoire('Confidentialité incorrecte :') + #13 + E.Message;
  end;

end;

{---------------------------------------------------------------------------------------}
procedure TOM_PROSPECTCONF.GenereSQLClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  NextPrevControl(Ecran);
  s := GenereSQL;
  if s <> '' then PGIError(s, Ecran.Caption);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_PROSPECTCONF.GetValeurClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Ind : Char;
  Val : string;
  Rge : string;
  Lql : string;
begin
  Ind := TToolbarButton97(Sender).Name[Length(TToolbarButton97(Sender).Name)];

  Rge := '';
  if TestJoker(GetControlText('RTC_VAL' + Ind)) then Rge := GetControlText('RTC_VAL' + Ind);
  Lql := GetControlText('RTC_VAL' + Ind);

  Val := TrLanceFiche_TabletteConf(Rge, Lql, GetControlText('RTC_CHAMP' + Ind));
  {Si on a valider la sélection, on met à jour le champs}
  if Val <> '@@@@' then SetField('RTC_VAL' + Ind, Val);//SetControlText('RTC_VAL' + Ind, Val);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_PROSPECTCONF.VideControls(Ind : Char);
{---------------------------------------------------------------------------------------}
begin
  if DS.State in [dsInsert, dsEdit] then begin
    SetField('RTC_OPER' + Ind, '');
    SetField('RTC_VAL'  + Ind, '');
    SetControlText('RTC_OPER' + Ind, '');
    SetControlText('RTC_VAL'  + Ind, '');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_PROSPECTCONF.OnEditKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  case Key of
    VK_F5 : GetValeurClick(Sender);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_PROSPECTCONF.EcranKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  case Key of
    Ord('Z') : if Shift = [ssCtrl	] then RazZones;
  else
    TFFiche(Ecran).FormKeyDown(Sender, Key, Shift);
  end;
end;

{09/10/07 : Réinitialisation sur le Ctrl+Z
{---------------------------------------------------------------------------------------}
procedure TOM_PROSPECTCONF.RazZones;
{---------------------------------------------------------------------------------------}
begin
  VideControls('1');
  VideControls('2');
  VideControls('3');
  SetField('RTC_CHAMP1', '');
  SetField('RTC_CHAMP2', '');
  SetField('RTC_CHAMP3', '');
  SetControlText('RTC_CHAMP1', '');
  SetControlText('RTC_CHAMP2', '');
  SetControlText('RTC_CHAMP3', '');
  SetField('RTC_LIEN1', '');
  SetField('RTC_LIEN2', '');
  SetControlText('RTC_LIEN1', '');
  SetControlText('RTC_LIEN2', '');
  SetField('RTC_SQLCONF', '');
  SetControlText('RTC_SQLCONF', '');
end;

initialization
  RegisterClasses([TOM_PROSPECTCONF]);

end.

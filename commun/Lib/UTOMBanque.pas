unit UTOMBanque;

interface

uses
{$IFDEF VER150}
  Variants,
{$ENDIF}

  Controls, Classes, sysutils,
{$IFDEF EAGLCLIENT}
  Maineagl, UTob, eFiche, eFichList,
{$ELSE}
  Fe_Main, Fiche, db, FichList, HDB,
  {$IFNDEF DBXPRESS}dbtables {BDE}, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
  UTOM;

type
  TOM_BANQUES = class(TOM)
    procedure OnArgument    (S: string); override;
    procedure OnUpdateRecord           ; override;
    procedure OnLoadRecord             ; override;
    procedure OnChangeField(F : TField); override;
    procedure OnNewRecord              ; override;
    procedure OnDeleteRecord           ; override;

  private
    Modifier: Boolean;
    {JP 11/03/04 : Fusion des tom de la Tréso et la Compa}
    Validation: TNotifyEvent;
    procedure PresenteFiche;
    function  DenominateurOk(Denomin : string): boolean;
    procedure RempliTaux    (Mode : string);
    procedure TbbDateModifier;
  public
    procedure CacheLesChamps   (Nom : string);
    procedure CodeCalOnKeyPress(Sender : TObject; var Key : Char);
    procedure ValiderClick     (Sender : TObject);
  end;



const
 // libellés des messages
  TexteMessage: array[1..4] of string = (
          {1}'Vous devez renseigner un code.'
          {2}, 'Vous devez renseigner un libellé.'
          {3}, 'Cette saisie n''est pas valide pour ce type d''opération.'
          {4}, 'Cet établissement bancaire est référencé par un compte bancaire, vous ne pouvez pas le supprimer.'
    );

procedure TRLanceFiche_Banques(Dom, Fiche, Range, Lequel, Arguments: string);

implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  {$ENDIF MODENT1}
  HTB97,HMsgBox, HEnt1, Ent1, Windows, ParamSoc, Constantes, Commun, HCtrls,
  ComCtrls, StdCtrls, UtilPgi, Forms, Messages, M3FP;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_Banques(Dom, Fiche, Range, Lequel, Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_BANQUES.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 500004;
  if EstComptaTreso then begin
    {$IFDEF EAGLCLIENT}
    THEdit(GetControl('PQ_BANQUE')).OnKeyPress := CodeCalOnKeyPress;
    THEdit(GetControl('PQ_ETABBQ')).OnKeyPress := CodeCalOnKeyPress;
    {$ELSE}
    THDBEdit(GetControl('PQ_BANQUE')).OnKeyPress := CodeCalOnKeyPress;
    THDBEdit(GetControl('PQ_ETABBQ')).OnKeyPress := CodeCalOnKeyPress;
    {$ENDIF EAGLCLIENT}
    {JP 23/09/03 : On récupère le traitement de BValiderClick de l'ancêtre Pour le lancer au début
                   ValiderClick}
    Validation := TToolbarButton97(GetControl('BVALIDER')).OnClick;
    TToolbarButton97(GetControl('BVALIDER')).OnClick := ValiderClick;
  end;
  {JP 16/08/07 : fusion des deux fiches}
  PresenteFiche;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_BANQUES.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if EstComptaTreso and (GetControlText('PQ_ETABBQ') = '') then begin
    LastErrorMsg := TraduireMemoire('Veuillez saisir le code établissement');
  // SBO 28/08/2007 plantage de la fiche si tréso activée
  if GetControl('TSGENERAL')<>nil then
    SetActiveTabSheet('TSGENERAL');
    SetFocusControl('PQ_ETABBQ');
    LastError := 1;
  end;
  if (GetField('PQ_BANQUE') = '') then begin
    // SBO 28/08/2007 plantage de la fiche si tréso activée
    if GetControl('TSGENERAL')<>nil then
      SetActiveTabSheet('TSGENERAL');
    SetFocusControl('PQ_BANQUE');
    LastError := 1;
    LastErrorMsg := TexteMessage[LastError];
    Exit;
  end;
  if (GetField('PQ_LIBELLE') = '') then begin
    // SBO 28/08/2007 plantage de la fiche si tréso activée
    if GetControl('TSGENERAL')<>nil then
      SetActiveTabSheet('TSGENERAL');
    SetFocusControl('PQ_LIBELLE');
    LastError := 2;
    LastErrorMsg := TexteMessage[LastError];
    Exit;
  end;

  if EstComptaTreso then Exit;

  if not DenominateurOk('PQ_DE_DENOMINATEUR') then Exit;
  if not DenominateurOk('PQ_PD_DENOMINATEUR') then Exit;
  if not DenominateurOk('PQ_CO_DENOMINATEUR') then Exit;
  if not DenominateurOk('PQ_CR_DENOMINATEUR') then Exit;

  RempliTaux('PQ_DE_MODE');
  RempliTaux('PQ_PD_MODE');
  RempliTaux('PQ_CR_MODE');
  RempliTaux('PQ_CO_MODE');

  TbbDateModifier;
end;


{---------------------------------------------------------------------------------------}
procedure TOM_BANQUES.OnLoadRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if (not EstComptaTreso) then begin
    Modifier := False;
    CacheLesChamps('PQ_DE_MODE'); CacheLesChamps('PQ_PD_MODE');
    CacheLesChamps('PQ_CR_MODE'); CacheLesChamps('PQ_CO_MODE');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_BANQUES.OnChangeField(F: TField);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if not EstComptaTreso then
    Modifier := True;
  if (F.FieldName = 'PQ_CO_MODE') or
     (F.FieldName = 'PQ_CR_MODE') or
     (F.FieldName = 'PQ_DE_MODE') or
     (F.FieldName = 'PQ_PD_MODE') then CacheLesChamps(F.FieldName);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_BANQUES.OnNewRecord;
{---------------------------------------------------------------------------------------}
var
  i  : Integer;
  {$IFDEF EAGLCLIENT}
  CH : THValComboBox;
  {$ELSE}
  CH : THDBValComboBox;
  {$ENDIF EAGLCLIENT}
  BoxRem, BoxRes: TGroupBox;
begin
  inherited;
  if not EstComptaTreso then begin
    SetField('PQ_DE_MODE', 'TX'); SetField('PQ_PD_MODE', 'TX');
    SetField('PQ_CR_MODE', 'TX'); SetField('PQ_CO_MODE', 'TX');
    SetField('PQ_DE_TXREF', 'TBB'); SetField('PQ_PD_TXREF', 'TBB');
    SetField('PQ_CR_TXREF', 'TBB'); SetField('PQ_CO_TXREF', 'TBB');
    SetField('PQ_BB_DATETAUX', Date); SetField('PQ_DE_DATETAUX', Date);
    SetField('PQ_PD_DATETAUX', Date); SetField('PQ_CR_DATETAUX', Date);
    SetField('PQ_CO_DATETAUX', Date);
    {JP 28/03/07 : FQ 19170 : PQ_CR_TYPEPLAFOND est champs combo, mais pas PQ_CR_PLAFOND
    SetField('PQ_CR_TYPEPLAFOND',0) ;}
    SetField('PQ_CR_PLAFOND', 0);

    CacheLesChamps('PQ_DE_MODE');
    CacheLesChamps('PQ_PD_MODE');
    CacheLesChamps('PQ_CR_MODE');
    CacheLesChamps('PQ_CO_MODE');

    BoxRem := TGroupBox(GetControl('BoxRemise'));
    BoxRes := TGroupBox(GetControl('BoxRestitution'));

    {$IFDEF EAGLCLIENT}
    for i := 0 to BoxRem.ControlCount - 1 do
      if BoxRem.Controls[i] is THValComboBox then
      begin
        CH := THValComboBox(BoxRem.Controls[i]);
        if CH.Values.Count > 0 then CH.Value := CH.Values[0];
      end;
    for i := 0 to BoxRes.ControlCount - 1 do
      if BoxRes.Controls[i] is THValComboBox then
      begin
        CH := THValComboBox(BoxRes.Controls[i]);
        if CH.values.Count > 0 then CH.Value := CH.Values[0];
      end;
    {$ELSE}
    for i := 0 to BoxRem.ControlCount - 1 do
      if BoxRem.Controls[i] is THDBValComboBox then
      begin
        CH := THDBValComboBox(BoxRem.Controls[i]);
        if CH.Values.Count > 0 then CH.Value := CH.Values[0];
      end;
    for i := 0 to BoxRes.ControlCount - 1 do
      if BoxRes.Controls[i] is THDBValComboBox then
      begin
        CH := THDBValComboBox(BoxRes.Controls[i]);
        if CH.values.Count > 0 then CH.Value := CH.Values[0];
      end;
    {$ENDIF EAGLCLIENT}
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_BANQUES.OnDeleteRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if not EstComptaTreso then begin
    {JP 16/05/06 : FQ 10349 : gestion du multisoc car pour le moment BANQUECP n'est pas partagée}
    if PresenceMS('BANQUECP', 'BQ_BANQUE', VarToStr(GetField('PQ_BANQUE'))) then begin
      LastError := 4;
      LastErrorMsg := TexteMessage[LastError];
    end;
  end
  {JP 26/04/04 : si la trésorerie accompagne la comptabilité, on fait le test sur les agences}
  else begin
    {JP 03/10/06 : Interdiction de la suppression de la banque des comptes courants si Tréso multi sociétés}
    if EstMultiSoc and (GetParamSocSecur('SO_TRBASETRESO', '') <> '') and
      (VarToStr(GetField('PQ_BANQUE')) = CODECOURANTS) then begin
      LastError := 6;
      LastErrorMsg := TraduireMemoire('Cet établissement bancaire est obligatoire pour la gestion des comptes courants :'#13 +
        'vous ne pouvez pas le supprimer.');
    end
    {JP 16/05/06 : FQ 10349 : gestion du multisoc car pour le moment BANQUECP n'est pas partagée}
    else if PresenceMS('AGENCE', 'TRA_BANQUE', VarToStr(GetField('PQ_BANQUE'))) then begin
      LastError := 5;
      LastErrorMsg := TraduireMemoire('Cet établissement bancaire est référencé par une agence bancaire, vous ne pouvez pas le supprimer.');
    end;
  end;
end;


{---------------------------------------------------------------------------------------}
function TOM_BANQUES.DenominateurOk(Denomin : string): Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := True;

  // SBO 29/08/2007 : Pb composant inexistant suite fusion TR / CP...
  if GetControl(Denomin) = nil then Exit ;

  if not GetControlVisible(Denomin) then Exit;
  if GetControlText(Denomin) = '0' then begin
    Result := False;
    SetActiveTabSheet('TSCONDITION');
    SetFocusControl(Denomin);
    LastError := 3;
    LastErrorMsg := TexteMessage[LastError];
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_BANQUES.RempliTaux(Mode : string);
{---------------------------------------------------------------------------------------}
var Pref, Tau: string;
  Sai, Num, Den, Cor, Tbb: Extended;
  Q: TQuery;
begin

  // SBO 29/08/2007 : Pb composant inexistant suite fusion TR / CP...
  if GetControl(Mode) = nil then Exit ;

  Pref := Copy(Mode, 4, 2);
  Sai := GetField('PQ_' + Pref + '_SAISIE');
  Num := GetField('PQ_' + Pref + '_NUMERATEUR');
  Den := GetField('PQ_' + Pref + '_DENOMINATEUR');
  Cor := GetField('PQ_' + Pref + '_CORRECTION');
  
  if VarToStr(GetField(Mode)) = 'TX' then
    SetField('PQ_' + Pref + '_TAUX', Sai)
  else if VarToStr(GetField(Mode)) = 'TF' then
    SetField('PQ_' + Pref + '_TAUX', Sai + (Num / Den))
  else if VarToStr(GetField(Mode)) = 'TT' then begin
    Tau := GetField('PQ_' + Pref + '_TXREF');
    Q := OpenSql('SELECT BT_TAUX FROM BQTAUX WHERE BT_TYPETAUX="' + Tau + '" And ' +
      'BT_DATE<="' + UsDateTime(GetField('PQ_BB_DATETAUX')) + '" order by BT_DATE DESC', True);
    Tbb := Q.Fields[0].AsFloat;
    Ferme(Q);
    SetField('PQ_' + Pref + '_TAUX', Tbb + Cor);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure AGLCacheLesChamps(Parms : array of Variant; Nb : Integer);
{---------------------------------------------------------------------------------------}
var
  F  : TForm;
  OM : TOM;
begin
  if not EstComptaTreso then begin
    F := TForm(Longint(Parms[0]));
    {16/08/07 : Bien qu'il s'agisse maintenant d'une ficheListe dans tous les cas, je maintiens les
                les deux options tant qu'il n'y a pas de SocRef}
    if (F is TFfiche) then OM := TOM_BANQUES(TFfiche(F).OM)
    else if (F is TFFicheListe) then OM := TOM_BANQUES(TFFicheListe(F).OM)
    else exit;
    if (OM is TOM_BANQUES) then TOM_BANQUES(OM).CacheLesChamps(string(Parms[1]))
    else exit;
  end;
end;

{Cache les champs Numérateur, Dénomination et Correction en fonction du mode de saisie des taux
{---------------------------------------------------------------------------------------}
procedure TOM_BANQUES.CacheLesChamps(Nom : string);
{---------------------------------------------------------------------------------------}
var
  Val, Pref : string;
begin

  // SBO 29/08/2007 : Pb composant inexistant suite fusion TR / CP...
  if GetControl(Nom) = nil then Exit ;

  // SBO 28/08/2007 plantage à l'ouverture de la fiche si tréso activée
//  Val := THDBValComboBox(GetControl(Nom)).Value;
  val := GetControlText( Nom ) ;
  
  Pref := Copy(Nom, 4, 2);
  if Val = 'TX' then
  begin
    SetControlVisible('PQ_' + Pref + '_SAISIE', True);
    SetControlVisible('PQ_' + Pref + '_TXREF', False);
    SetControlVisible('PQ_' + Pref + '_PLUS', False);
    SetControlVisible('PQ_' + Pref + '_DIV', False);
    SetControlVisible('PQ_' + Pref + '_NUMERATEUR', False);
    SetControlVisible('PQ_' + Pref + '_DENOMINATEUR', False);
    SetControlVisible('PQ_' + Pref + '_CORRECTION', False);
    SetField('PQ_' + Pref + '_NUMERATEUR', 0);
    SetField('PQ_' + Pref + '_DENOMINATEUR', 0);
    SetField('PQ_' + Pref + '_CORRECTION', 0);
  end
  else if Val = 'TF' then
  begin
    SetControlVisible('PQ_' + Pref + '_SAISIE', True);
    SetControlVisible('PQ_' + Pref + '_NUMERATEUR', True);
    SetControlVisible('PQ_' + Pref + '_PLUS', True);
    SetControlVisible('PQ_' + Pref + '_DIV', True);
    SetControlVisible('PQ_' + Pref + '_DENOMINATEUR', True);
    SetControlVisible('PQ_' + Pref + '_TXREF', False);
    SetControlVisible('PQ_' + Pref + '_CORRECTION', False);
    SetField('PQ_' + Pref + '_CORRECTION', 0);
  end
  else if Val = 'TT' then
  begin
    SetControlVisible('PQ_' + Pref + '_TXREF', True);
    SetControlVisible('PQ_' + Pref + '_CORRECTION', True);
    SetControlVisible('PQ_' + Pref + '_SAISIE', False);
    SetControlVisible('PQ_' + Pref + '_PLUS', False);
    SetControlVisible('PQ_' + Pref + '_DIV', False);
    SetControlVisible('PQ_' + Pref + '_NUMERATEUR', False);
    SetControlVisible('PQ_' + Pref + '_DENOMINATEUR', False);
    SetField('PQ_' + Pref + '_NUMERATEUR', 0);
    SetField('PQ_' + Pref + '_DENOMINATEUR', 0);
    SetField('PQ_' + Pref + '_SAISIE', 0);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_BANQUES.TbbDateModifier;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  C : string;
begin
  if not Modifier then Exit;
  try
    C := StrfPoint(GetField('PQ_BB_TAUX'));
  except
    C := '0';
  end;

  Q := OpenSql('SELECT BT_DATE FROM BQTAUX WHERE BT_DATE="' + UsDateTime(StrToDate(GetField('PQ_BB_DATETAUX'))) + '" ' +
    'And BT_TYPETAUX="TBB" And BT_BANQUE="' + GetField('PQ_BANQUE') + '"', True);
  if Q.Eof then
  begin
    ExecuteSql('INSERT INTO BQTAUX (BT_BANQUE, BT_TYPETAUX, BT_DATE, BT_TAUX) ' +
      'VALUES ("' + GetField('PQ_BANQUE') + '", "TBB", "' + UsDateTime(StrToDate(GetField('PQ_BB_DATETAUX'))) + '", ' +
      '' + C + ')');
  end
  else
    ExecuteSql('UPDATE BQTAUX SET BT_DATE="' + UsDateTime(StrToDate(GetField('PQ_BB_DATETAUX'))) + '", ' +
      'BT_TAUX=' + C + ' Where ' +
      'BT_BANQUE="' + GetField('PQ_BANQUE') + '" ' +
      'And BT_DATE="' + UsDateTime(StrToDate(GetField('PQ_BB_DATETAUX'))) + '"');
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_BANQUES.CodeCalOnKeyPress(Sender: TObject; var Key: Char);
{---------------------------------------------------------------------------------------}
begin
  {JP 05/08/03 : Ce code est nécessaire pour la version CWAS. En effet mettre '00000' comme EditMask
                 dans decla est interprété en CWAS comme le formatage d'une date !!!!}
  if (UpperCase(THEdit(Sender).Name) = 'PQ_ETABBQ') and not
    (Key in ['0'..'9', chr(VK_BACK), chr(VK_RIGHT), chr(VK_LEFT), chr(VK_DELETE)]) then Key := #0;
  inherited;
  ValidCodeOnKeyPress(Key);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_BANQUES.ValiderClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {On exécute d'abord le traitement de l'ancêtre}
  Validation(Sender);
  {Puis si on n'est pas dans le panel, on ferme la fiche}
  if not IsInside(Ecran) then
    SendMessage(Ecran.Handle, WM_CLOSE, 0, 0);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_BANQUES.PresenteFiche;
{---------------------------------------------------------------------------------------}
var
  pc : TPageControl;
  n  : Integer;
begin
  // SBO 28/08/2007 plantage de la fiche si tréso activée
  if GetControl('PAGES')=nil then Exit ;
  if EstComptaTreso then begin
    pc := (GetControl('PAGES') as TPageControl);
    for n := 0 to pc.PageCount - 1 do
      pc.Pages[n].TabVisible := False;
      Ecran.Height := 240;
      Ecran.Width  := 604;
  end;
  // SBO 28/08/2007 plantage de la fiche si tréso activée
  if GetControl('TSGENERAL')<>nil then
    SetActiveTabSheet('TSGENERAL');
end;

initialization
  registerclasses([TOM_BANQUES]);
  RegisterAGLProc('CacheLesChamps', true, 1, AGLCacheLesChamps);

end.




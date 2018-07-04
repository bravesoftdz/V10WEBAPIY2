{***********UNITE*************************************************
Auteur  ...... : TG
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : ACTIVITEPAIE (ACTIVITEPAIE)
Mots clefs ... : TOM;ACTIVITEPAIE;LIENACTIVPAIE
*****************************************************************}
unit UTomActivitePaie;

interface

uses StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fiche,
  FichList,
  {$ELSE}
  eFiche,
  eFichList,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOM,
  ParamSoc,
  UtilArticle,
  UTob;

type
  TOM_ACTIVITEPAIE = class(TOM)
    procedure OnNewRecord; override;
    procedure OnLoadRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    private
    procedure NONMENSUALISEClick(sender: Tobject);
  end;

const
  TexteMsgActivitePaie: array[1..1] of string = (
    {1}'Vous devez saisir au moins un type, un code ou une famille d''article'
    );

implementation
//uses AfUtilArticle;


procedure TOM_ACTIVITEPAIE.OnArgument(S: string);
begin
  inherited;
  if GetParamSoc('SO_AFLIENPAIEDEC') then
  begin
    SetControlProperty('ACP_PROFIL', 'Style', csDropDown);
    SetControlProperty('ACP_RUBRIQUE', 'ElipsisButton', false);
  end;
  SetControlVisible('ACP_VENTANARUB', GetParamSoc('SO_AFLIENPAIEANA'));
  TCheckBox(GetControl('NONMENSUALISE')).OnClick := NONMENSUALISEClick;
{$IFDEF BTP}
  THValComboBox(GetControl('ACP_TYPEARTICLE')).Plus := 'AND (CO_CODE="PRE" OR CO_CODE="FRA")';
{$ENDIF}
end;

procedure TOM_ACTIVITEPAIE.OnNewRecord;
var Q: TQuery;
  i: integer;
begin
  inherited;
  Q := OpenSQL('SELECT MAX(ACP_RANG) AS MUX FROM ACTIVITEPAIE', true);
  if Q.EOF then i := 0
  else i := Q.FindField('MUX').AsInteger + 1;
  Ferme(Q);
  SetField('ACP_RANG', i);
  SetControlChecked('NONMENSUALISE', false);
  SetField ('ACP_MENSUALISE','X');
end;

procedure TOM_ACTIVITEPAIE.OnLoadRecord;
begin
  inherited;
  if not (ds.state in [dsinsert]) then
  SetControlChecked('NONMENSUALISE', (GetField('ACP_MENSUALISE') = '-'));
end;

procedure TOM_ACTIVITEPAIE.OnUpdateRecord;
begin
  inherited;
  if (GetField('ACP_ARTICLE') = '') and (GetField('ACP_CODEARTICLE') = '') and (GetField('ACP_TYPEARTICLE') = '') and
  (GetField('ACP_TYPEHEURE') = '')  and
  (GetField('ACP_FAMILLENIV1') = '') and (GetField('ACP_FAMILLENIV2') = '') and (GetField('ACP_FAMILLENIV3') = '') then
  begin
    LastError := 1;
    LastErrorMsg := TexteMsgActivitePaie[LastError];
  end ;
end;

procedure TOM_ACTIVITEPAIE.NONMENSUALISEClick(sender: Tobject);
begin
  if not(DS.State in [dsInsert,dsEdit]) and (GetControltext('NONMENSUALISE') = GetField ('ACP_MENSUALISE')) then
   DS.edit;
  if GetControltext('NONMENSUALISE') = 'X' then SetField ('ACP_MENSUALISE','-')
  else SetField ('ACP_MENSUALISE','X');
end;

procedure TOM_ACTIVITEPAIE.OnChangeField(F: TField);
begin
  inherited;

  if (F.FieldName = 'ACP_CODEARTICLE') or (F.FieldName = 'ACP_TYPEARTICLE') then
  begin
    SetControlEnabled('ACP_FAMILLENIV1', F.Value = '');
    SetControlEnabled('ACP_FAMILLENIV2', F.Value = '');
    SetControlEnabled('ACP_FAMILLENIV3', F.Value = '');

    if F.FieldName = 'ACP_CODEARTICLE' then
      if F.Value = '' then SetField('ACP_ARTICLE', '')
      else SetField('ACP_ARTICLE', CodeArticleUnique(F.Value, '', '', '', '', ''));
  end else
    if (F.FieldName = 'ACP_FAMILLENIV1') or (F.FieldName = 'ACP_FAMILLENIV2') or (F.FieldName = 'ACP_FAMILLENIV3') then
  begin
    SetControlEnabled('ACP_CODEARTICLE', F.Value = '');
    SetControlEnabled('ACP_TYPEARTICLE', F.Value = '');
  end;
  if (F.FieldName = 'ACP_RESSOURCE') then
  begin
    if trim(F.Value) <> '' then
    begin
      SetField('ACP_PROFIL', '');
      SetControlEnabled('ACP_PROFIL', false);
    end
    else SetControlEnabled('ACP_PROFIL', true);
  end;
end;

initialization
  registerclasses([TOM_ACTIVITEPAIE]);
end.

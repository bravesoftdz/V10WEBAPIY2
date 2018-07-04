{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/03/2002
Modifié le ... :   /  /
Description .. : unit de gestion du multi critère des objectifs des stages
Suite ........ :
Mots clefs ... : PAIE;FORMATION
*****************************************************************}
unit UTofPG_MulObjectifStage;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Mul, Fe_Main, DBGrids,
{$ELSE}
  MaineAgl, eMul,
{$ENDIF}
  HCtrls, Hqry, UTOF;

type
  TOF_PGMulObjectifStage = class(TOF)
  private
    Q_Mul: THQuery;
    BtnCherche: TToolbarButton97;
    Nat: THValComboBox;
    TypeSaisie, LeStage: string;
    procedure ActiveWhere;
    procedure NatChange(Sender: TObject);
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  end;

implementation

procedure TOF_PGMulObjectifStage.ActiveWhere;
begin
  if GetControlText('POS_NATOBJSTAGE') <> 'COM' then
  begin
    if Q_Mul <> nil then begin  TFMul(Ecran).SetDBListe('PGOBJECTIFSTAGE'); end;
  end
  else
  begin
    if Q_Mul <> nil then begin  TFMul(Ecran).SetDBListe('PGCOMPETENCESTAGE'); end;
  end;
end;

procedure TOF_PGMulObjectifStage.NatChange(Sender: TObject);
begin
  if GetControlText('POS_NATOBJSTAGE') <> 'COM' then
  begin
    SetControlVisible('POS_COMPETENCE', FALSE);
    SetControlVisible('TPOS_COMPETENCE', FALSE);
  end
  else
  begin
    SetControlVisible('POS_COMPETENCE', TRUE);
    SetControlVisible('TPOS_COMPETENCE', TRUE);
  end;
  ActiveWhere;
end;

procedure TOF_PGMulObjectifStage.OnArgument(Arguments: string);
var
{$IFNDEF EAGLCLIENT}
  Grille: THDBGrid;
{$ELSE}
  Grille: THGrid;
{$ENDIF}
begin
  inherited;
  TypeSaisie := Trim(ReadTokenPipe(Arguments, ';'));
  LeStage := Trim(ReadTokenPipe(Arguments, ';'));
  if LeStage <> '' then
  begin
    SetControltext('POS_CODESTAGE', LeStage);
    SetControlEnabled('POS_CODESTAGE', False);
    SetControltext('POS_CODESTAGE_', LeStage);
    SetControlEnabled('POS_CODESTAGE_', False);
  end;
  BtnCherche := TToolbarButton97(GetControl('BCherche'));
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
{$IFNDEF EAGLCLIENT}
  Grille := THDBGrid(GetControl('Fliste'));
{$ELSE}
  Grille := THGrid(GetControl('Fliste'));
{$ENDIF}
  SetControlVisible('BInsert', TRUE);
  Nat := THvalComboBox(GetControl('POS_NATOBJSTAGE'));
  if Nat <> nil then Nat.Onchange := NatChange;
end;


procedure TOF_PGMulObjectifStage.OnLoad;
begin
  inherited;
  if not (Ecran is TFMul) then exit;
  if GetControlText('POS_NATOBJSTAGE') = 'COM' then
  begin
    SetControlVisible('POS_COMPETENCE', TRUE);
    SetControlVisible('TPOS_COMPETENCE', TRUE);
    SetControlVisible('POS_COUTFORFAIT', TRUE);
    SetControlVisible('TPOS_COUTFORFAIT', TRUE);
  end
  else
  begin
    SetControlVisible('POS_COMPETENCE', FALSE);
    SetControlVisible('TPOS_COMPETENCE', FALSE);
    SetControlVisible('POS_COUTFORFAIT', FALSE);
    SetControlVisible('TPOS_COUTFORFAIT', FALSE);
  end;
// ActiveWhere;
end;


initialization
  registerclasses([TOF_PGMulObjectifStage]);
end.


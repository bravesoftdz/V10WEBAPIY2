{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : unit de gestion du multi critère de la ventilation comptable
Suite ........ : par défaut des cotisations
Mots clefs ... : PAIE;GENERCOMPTA
*****************************************************************}
unit UTofPG_MulVentilCot;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Mul, Fe_Main, DBGrids,
{$ELSE}
  MaineAgl, eMul,
{$ENDIF}
  Grids, HCtrls, HEnt1, vierge, EntPaie, HMsgBox, Hqry, UTOF, UTOB, UTOM,
  AGLInit;
type
  TOF_PGMULVENTILCOT = class(TOF)
  private
    Affect: TCheckBox;
    Q_Mul: THQuery;
    BtnCherche: TToolbarButton97;
    procedure ActiveWhere;
    procedure GrilleDblClick(Sender: TObject);
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  end;

implementation

procedure TOF_PGMULVENTILCOT.ActiveWhere;
begin
  if Affect = nil then exit;
  if Affect.Checked = FALSE then
  begin
    if Q_Mul <> nil then begin TFMul(Ecran).SetDBListe('PGVENTILACOTISAT'); end;
  end
  else
  begin
    if Q_Mul <> nil then begin TFMul(Ecran).SetDBListe('PGVENTILSCOTISAT'); end;
  end;
  if BtnCherche <> nil then BtnCherche.click;
end;


procedure TOF_PGMULVENTILCOT.GrilleDblClick(Sender: TObject);
begin
  if not Affect.checked then
    AglLanceFiche('PAY', 'VENTILCOT', '', Q_Mul.FindField('PVT_PREDEFINI').AsString + ';' + Q_Mul.FindField('PVT_NODOSSIER').AsString + ';' + Q_Mul.FindField('PVT_RUBRIQUE').AsString, '')
  else
    AglLanceFiche('PAY', 'VENTILCOT', '', '', Q_Mul.FindField('PCT_PREDEFINI').AsString + ';' + Q_Mul.FindField('PCT_NODOSSIER').AsString + ';' + Q_Mul.FindField('PCT_RUBRIQUE').AsString + ';' + 'ACTION=CREATION');
  if BtnCherche <> nil then BtnCherche.click;
end;

procedure TOF_PGMULVENTILCOT.OnArgument(Arguments: string);
var
{$IFNDEF EAGLCLIENT}
  Grille: THDBGrid;
{$ELSE}
  Grille: THGrid;
{$ENDIF}
begin
  inherited;
  BtnCherche := TToolbarButton97(GetControl('BCherche'));
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  Affect := TCheckBox(GetControl('CHBXANNOMALIE'));
{$IFNDEF EAGLCLIENT}
  Grille := THDBGrid(GetControl('Fliste'));
{$ELSE}
  Grille := THGrid(GetControl('Fliste'));
{$ENDIF}
  SetControlVisible('BInsert', FALSE);
  if Grille <> nil then Grille.OnDblClick := GrilleDblClick;
end;


procedure TOF_PGMULVENTILCOT.OnLoad;
var TT: TFMul;
begin
  inherited;
  if not (Ecran is TFMul) then exit;
  if Affect <> nil then
  begin
    if Affect.Checked = FALSE then TFMul(Ecran).Caption := 'Liste des cotisations affectées à un compte'
    else TFMul(Ecran).Caption := 'Liste des cotisations non affectées à un compte';
    TT := TFMul(Ecran);
    if TT <> nil then UpdateCaption(TT);
  end;
  ActiveWhere;
end;


initialization
  registerclasses([TOF_PGMULVENTILCOT]);
end.


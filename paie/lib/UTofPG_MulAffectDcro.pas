{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 16/03/2001
Modifié le ... : 02/10/2001
Description .. :  TOF affectation par defaut des codes ducs,
                  cotisations, remunérations et organismes de paie
                  TOF unique pour gérer le cas des non affectation
                  et la création automatique des ventilations
Objectif ......:  visualiser les rubriques et les organismes non
                  affectés.On joue sur les listes basées sur des
                  vues
Mots clefs ... : PAIE, PGDUCS, PGCOTISATION
*****************************************************************}
{
 PT1 06/07/2001   V540 MF Correction multicritère MUL_DUCSAFFECT - PGDUCS
 PT2 02/10/2001   V562 MF Modif RendLeWhere pour recup Critères CODIFICATION - PGDUCS
 PT3 03/10/2001   V562 MF Modif RendLeWhere pour nouveau critère NATUREORG - PGDUCS
 PT4 01/07/2002   V582 MF Modif RendLeWhere pour ne Sélectionner que les Cotisation
                          du dossier (AND ##PCT_PREDEFINI##)
 PT5 18/10/2002   V585 PH Affectation des bornes par defaut des rubriques pour faire
                          une requete correcte en mono bug Agl540n
// **** Refonte accès V_PGI_env ***** V_PGI_env.ModeFonc remplacé par PgRendModeFonc () *****
 PT6 17/09/2003   V_421 PH FQ 10429 Sans affectation par defaut
 PT7 03/10/2003  MF V_421  FQ 10873 mise au point CWAS : concerne codif ducs,
                           cotisations, rémunérations
 PT8 04/05/2006   V_65 PH FQ 12256 FocusControl sur la saisie du champ predefini
}

unit UTofPG_MulAffectDcro;

interface
uses
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBGrids, mul, FE_Main,
{$ELSE}
  MaineAgl, eMul,
{$ENDIF}
  StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
  Grids, HCtrls, HEnt1, vierge, EntPaie, HMsgBox, Hqry, UTOF, UTOB, UTOM,
  AGLInit;
type
  TOF_PGMULAFFECTDCRO = class(TOF)
  private
    Affect: TCheckBox;
    Q_Mul: THQuery;
    BtnCherche: TToolbarButton97;
    QFonct: string; // Quelle fonctionnalité = multi critère
    procedure ActiveWhere;
    procedure GrilleDblClick(Sender: TObject);
    function RendLeWhere(R1, R2, R3, R4, Org: THEdit; R5, Pred: THValComboBox; Racine: string): string; // PT2, PT3
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  end;

implementation

uses P5Def;

procedure TOF_PGMULAFFECTDCRO.ActiveWhere;
var R1, R2, R3, R4, Org, WW: THEdit; // PT2
  R5, Pred: THValComboBox; // PT3

begin
  if Affect = nil then exit;
  R1 := THEdit(GetControl('RUBRIQUE'));
  R2 := THEdit(GetControl('RUBRIQUE_'));
// PT2
  R3 := nil;
  R4 := nil;
  R5 := nil; // PT3
  if (TFmul(Ecran).Name = 'MUL_DUCSAFFECT') then
  begin
    R3 := THEdit(GetControl('CODIFICATION'));
    R4 := THEdit(GetControl('CODIFICATION_'));
    R5 := THValComboBox(GetControl('NATUREORG')); // PT3
  end;
// FIN PT2
  Org := THEdit(GetControl('ORGANISME'));
  Pred := THValComboBox(GetControl('PREDEFINI'));
  WW := THEdit(GetControl('XX_WHERE'));
  if QFonct = 'D' then
  begin // cas des ducs
// début PT1
    if WW <> nil then
    begin
      WW.text := '';
      if Affect.Checked = FALSE then
      begin
        WW.Text := ' ##PDF_PREDEFINI## ';
        WW.Text := WW.Text + RendLeWhere(R1, R2, R3, R4, Org, R5, Pred, 'PDF'); // PT2 , PT3
      end
      else
      begin
        R3 := nil; // PT2
        R4 := nil; // PT2
        R5 := nil; // PT3
        WW.Text := RendLeWhere(R1, R2, R3, R4, Org, R5, Pred, 'PCT'); // PT2, PT3
        if WW.Text <> '' then WW.Text := WW.TEXT + 'AND ';
        WW.Text := WW.TEXT + ' NOT EXISTS (SELECT PDF_RUBRIQUE FROM DUCSAFFECT WHERE ##PDF_PREDEFINI## PDF_RUBRIQUE=COTISATION.PCT_RUBRIQUE) AND PCT_NATURERUB="COT"';
      end;
    end;
// fin PT1
    if Affect.Checked = FALSE then
    begin
      TFMul(Ecran).SetDBListe('PGDUCSAFFECT');
    end
    else
    begin
      TFMul(Ecran).SetDBListe('PGCOTSANSDUCS');
    end;
  end;
  if QFonct = 'C' then
  begin // cas des cotisations ventilations comptables
    if WW <> nil then
    begin
      WW.text := '';
      if Affect.Checked = FALSE then
      begin
        WW.Text := ' ##PCT_PREDEFINI## ';
        WW.Text := WW.Text + RendLeWhere(R1, R2, R3, R4, Org, R5, Pred, 'PVT'); //PT2, PT3
      end
      else
      begin
        WW.Text := RendLeWhere(R1, R2, R3, R4, Org, R5, Pred, 'PCT'); // PT2, PT3
        if WW.Text <> '' then WW.Text := WW.TEXT + 'AND ';
        WW.Text := WW.TEXT + ' NOT EXISTS (SELECT PVT_RUBRIQUE FROM VENTICOTPAIE WHERE ##PVT_PREDEFINI## PVT_RUBRIQUE=COTISATION.PCT_RUBRIQUE) AND PCT_NATURERUB="COT"';
      end;
//      if Affect.Checked=FALSE then WW.text := WW.text + ' PVT_PREDEFINI <> "ZZZ" LEFT JOIN COTISATION ON PCT_RUBRIQUE=PVT_RUBRIQUE';
    end;
    if Affect.Checked = FALSE then
    begin
      TFMul(Ecran).SetDBListe('PGVENTILACOTISAT');
    end
    else
    begin
      TFMul(Ecran).SetDBListe('PGVENTILSCOTISAT');
    end;
  end;

  if QFonct = 'R' then
  begin // cas des rémunérations ventilations comptables
    if WW <> nil then
    begin
      WW.text := '';
      if Affect.Checked = FALSE then
      begin
        WW.Text := ' ##PRM_PREDEFINI## ';
        WW.Text := WW.Text + RendLeWhere(R1, R2, R3, R4, Org, R5, Pred, 'PVS'); // PT2, PT3
      end
      else
      begin
        WW.Text := RendLeWhere(R1, R2, R3, R4, Org, R5, Pred, 'PRM'); // PT2 , PT3
        if WW.Text <> '' then WW.Text := WW.TEXT + ' AND ';
        WW.Text := WW.TEXT + ' NOT EXISTS (SELECT PVS_RUBRIQUE FROM VENTIREMPAIE WHERE ##PVS_PREDEFINI## PVS_RUBRIQUE=REMUNERATION.PRM_RUBRIQUE)';
      end;
    end;
    if Affect.Checked = FALSE then
    begin
      TFMul(Ecran).SetDBListe('PGVENTILAREM');
    end
    else
    begin
      TFMul(Ecran).SetDBListe('PGVENTILSREM');
    end;
  end;
  if QFonct = 'O' then
  begin // cas des organismes ventilations comptables
    if WW <> nil then
    begin
      WW.text := '';
      if Affect.Checked = FALSE then
      begin
        WW.text := '##PVO_PREDEFINI##';
        WW.Text := WW.text + RendLeWhere(R1, R2, R3, R4, Org, R5, Pred, 'PVO_TYPORGANISME'); // PT2, PT3
      end
      else
      begin
        WW.Text := RendLeWhere(R1, R2, R3, R4, Org, R5, nil, 'CC_CODE'); // PT2, PT3
        if WW.Text <> '' then WW.Text := WW.TEXT + ' AND ';
        WW.Text := WW.TEXT + ' CC_TYPE="PTG"  AND (NOT EXISTS (SELECT PVO_TYPORGANISME  FROM VENTIORGPAIE WHERE ##PVO_PREDEFINI## (PVO_TYPORGANISME=CHOIXCOD.CC_CODE)))';
      end;
    end;
    if Affect.Checked = FALSE then
    begin
      TFMul(Ecran).SetDBListe('PGVENTILAORGAN');
    end
    else
    begin
      TFMul(Ecran).SetDBListe('PGVENTILSORGAN');
    end;
  end;
end;


procedure TOF_PGMULAFFECTDCRO.GrilleDblClick(Sender: TObject);
begin
  if Affect = nil then exit;

{$IFDEF EAGLCLIENT}
  Q_Mul.TQ.Seek(TFmul(Ecran).Fliste.Row - 1); // PT7
{$ENDIF}

  if QFonct = 'D' then
  begin // cas des ducs
    if not Affect.checked then
      AglLanceFiche('PAY', 'DUCSAFFECT', '', Q_Mul.FindField('PDF_PREDEFINI').AsString + ';' + Q_Mul.FindField('PDF_NODOSSIER').AsString + ';' + Q_Mul.FindField('PDF_RUBRIQUE').AsString, '')
    else
      AglLanceFiche('PAY', 'DUCSAFFECT', '', '', Q_Mul.FindField('PCT_PREDEFINI').AsString + ';' + Q_Mul.FindField('PCT_NODOSSIER').AsString + ';' + Q_Mul.FindField('PCT_RUBRIQUE').AsString + ';' + 'ACTION=CREATION');
  end;
  if QFonct = 'C' then
  begin // cas des cotisations
    if not Affect.checked then
      AglLanceFiche('PAY', 'VENTILCOT', '', Q_Mul.FindField('PVT_PREDEFINI').AsString + ';' + Q_Mul.FindField('PVT_NODOSSIER').AsString + ';' + Q_Mul.FindField('PVT_RUBRIQUE').AsString, '')
    else
      AglLanceFiche('PAY', 'VENTILCOT', '', '', Q_Mul.FindField('PCT_PREDEFINI').AsString + ';' + Q_Mul.FindField('PCT_NODOSSIER').AsString + ';' + Q_Mul.FindField('PCT_RUBRIQUE').AsString + ';' + 'ACTION=CREATION');
  end;
  if QFonct = 'R' then
  begin // cas des rémunérations
    if not Affect.checked then
      AglLanceFiche('PAY', 'VENTILSTD', '', Q_Mul.FindField('PVS_PREDEFINI').AsString + ';' + Q_Mul.FindField('PVS_NODOSSIER').AsString + ';' + Q_Mul.FindField('PVS_RUBRIQUE').AsString, '')
    else
      AglLanceFiche('PAY', 'VENTILSTD', '', '', Q_Mul.FindField('PRM_PREDEFINI').AsString + ';' + Q_Mul.FindField('PRM_NODOSSIER').AsString + ';' + Q_Mul.FindField('PRM_RUBRIQUE').AsString + ';' + 'ACTION=CREATION');
  end;
  if QFonct = 'O' then
  begin // cas des rémunérations
    if not Affect.checked then
    begin
      SetControlEnabled('PREDEFINI', TRUE);
      AglLanceFiche('PAY', 'VENTILORGANISME', '', Q_Mul.FindField('PVO_PREDEFINI').AsString + ';' + Q_Mul.FindField('PVO_NODOSSIER').AsString + ';' + Q_Mul.FindField('PVO_TYPORGANISME').AsString, '')
    end
    else
    begin
      setControlProperty('PREDEFINI', 'Value', '');
      SetControlEnabled('PREDEFINI', FALSE);
      AglLanceFiche('PAY', 'VENTILORGANISME', '', '', Q_Mul.FindField('CC_CODE').AsString + ';' + 'ACTION=CREATION');
    end;
  end;
  if BtnCherche <> nil then BtnCherche.click;
end;

procedure TOF_PGMULAFFECTDCRO.OnArgument(Arguments: string);
var
{$IFDEF EAGLCLIENT}
  Grille: THGrid;
{$ELSE}
  Grille: THDBGrid;
{$ENDIF}

begin
  inherited;
  QFonct := Arguments;
  BtnCherche := TToolbarButton97(GetControl('BCherche'));
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  Affect := TCheckBox(GetControl('CHBXANNOMALIE'));
{$IFDEF EAGLCLIENT}
  Grille := THGrid(GetControl('Fliste'));
{$ELSE}
  Grille := THDBGrid(GetControl('Fliste'));
{$ENDIF}
  SetControlVisible('BInsert', FALSE);
  if Grille <> nil then Grille.OnDblClick := GrilleDblClick;
//  PT5 18/10/2002   V585 MF Affectation des bornes par defaut des rubriques
  if (QFonct = 'R') or (QFonct = 'C') then
  begin // initialisation des bornes par defaut
    SetControlText('RUBRIQUE', '0000');
    SetControlText('RUBRIQUE_', '9999');
//  PT6 17/09/2003   V_421 PH FQ 10429 Sans affectation par defaut
    SetControlChecked('CHBXANNOMALIE', TRUE);
  end
  else SetFocusControl ('PREDEFINI'); // PT8

end;


procedure TOF_PGMULAFFECTDCRO.OnLoad;
var TT: TFMul;
begin
  inherited;
  if not (Ecran is TFMul) then exit;
  if Affect <> nil then
  begin
    if QFonct = 'D' then
    begin // affectation des codes ducs aux cotisations
      if Affect.Checked = FALSE then TFMul(Ecran).Caption := 'Liste des cotisations affectées à un code ducs'
      else TFMul(Ecran).Caption := 'Liste des cotisations non affectées à un code ducs';
    end;
    if QFonct = 'C' then
    begin // affectation des codes ducs aux cotisations
      if Affect.Checked = FALSE then TFMul(Ecran).Caption := 'Liste des cotisations affectées à un compte'
      else TFMul(Ecran).Caption := 'Liste des cotisations non affectées à un compte';
    end;
    if QFonct = 'R' then
    begin // affectation des codes ducs aux cotisations
      if Affect.Checked = FALSE then TFMul(Ecran).Caption := 'Liste des rémunérations affectées à un compte'
      else TFMul(Ecran).Caption := 'Liste des rémunérations non affectées à un compte';
    end;
    if QFonct = 'O' then
    begin // affectation des codes ducs aux cotisations
      if Affect.Checked = FALSE then TFMul(Ecran).Caption := 'Liste des organismes affectés à un compte'
      else TFMul(Ecran).Caption := 'Liste des organismes non affectés à un compte';
    end;
    TT := TFMul(Ecran);

    if TT <> nil then UpdateCaption(TT);
  end;
  ActiveWhere;
end;

// fonction qui rend la clause where du multi critère en fonction des saisies et de la liste
// PT2 function TOF_PGMULAFFECTDCRO.RendLeWhere (R1,R2,Org : THEdit; Pred : THValComboBox; Racine : String) : String;

function TOF_PGMULAFFECTDCRO.RendLeWhere(R1, R2, R3, R4, Org: THEdit; R5, Pred: THValComboBox; Racine: string): string; // PT3
var
  DebCod, FinCod, LeText: string; // PT3
begin
  LeText := '';
  if (R1 <> nil) and (R1.Text <> '') then LeText := ' (' + Racine + '_RUBRIQUE>="' + R1.Text + '")';
  if (R2 <> nil) and (R2.Text <> '') then
  begin
    if LeText <> '' then LeText := LeText + ' AND (' + Racine + '_RUBRIQUE<="' + R2.Text + '")'
    else LeText := ' (' + Racine + '_RUBRIQUE<="' + R2.Text + '")';
  end;
// PT2 deb
  if (R3 <> nil) and (R3.Text <> '') then
  begin
    if LeText <> '' then
      LeText := LeText + ' AND (' + Racine + '_CODIFICATION>="' + R3.Text + '")'
    else
      LeText := ' (' + Racine + '_CODIFICATION>="' + R3.Text + '")';
  end;
  if (R4 <> nil) and (R4.Text <> '') then
  begin
    if LeText <> '' then
      LeText := LeText + ' AND (' + Racine + '_CODIFICATION<="' + R4.Text + '")'
    else
      LeText := ' (' + Racine + '_CODIFICATION<="' + R4.Text + '")';
  end;
// PT2 fin
// PT3 deb
  if (R5 <> nil) then
  begin
    DebCod := Copy(R5.Value, 1, 1);
    FinCod := DebCod + 'ZZZZZZ';
    if (DebCod = '') then FinCod := 'ZZZZZZZ';
    if (DebCod = '3') then FinCod := '5ZZZZZZ';

    if LeText <> '' then
    begin
      if (DebCod = '3') then
        LeText := LeText + 'AND (((' + Racine + '_CODIFICATION >="' + DebCod + '")AND (' + Racine + '_CODIFICATION <="' + FinCod + '"))' +
          ' OR ((' + Racine + '_CODIFICATION >= "8") AND (' + Racine + '_CODIFICATION <= "9ZZZZZZ")))'
      else
        LeText := LeText + 'AND ((' + Racine + '_CODIFICATION >="' + DebCod + '")AND (' + Racine + '_CODIFICATION <="' + FinCod + '"))';
    end
    else
    begin
      if (DebCod = '3') then
        LeText := 'AND (((' + Racine + '_CODIFICATION >="' + DebCod + '")AND (' + Racine + '_CODIFICATION <="' + FinCod + '"))' +
          ' OR ((' + Racine + '_CODIFICATION >= "8") AND (' + Racine + '_CODIFICATION <= "9ZZZZZZ")))'
      else
        LeText := 'AND ((' + Racine + '_CODIFICATION >="' + DebCod + '")AND (' + Racine + '_CODIFICATION <="' + FinCod + '"))';
    end;
  end;
// PT3 fin
  if (Pred <> nil) and (Pred.Value <> '') then
  begin
// PT1   if (Racine = 'PVS') OR (Racine = 'PVT') OR (Racine = 'PRM') OR (Racine = 'PCT')then
    if (Racine = 'PVS') or (Racine = 'PVT') or (Racine = 'PRM') or (Racine = 'PCT') or (Racine = 'PDF') then
    begin
      if LeText <> '' then LeText := LeText + ' AND (' + Racine + '_PREDEFINI="' + Pred.Value + '")'
      else LeText := ' (' + Racine + '_PREDEFINI="' + Pred.Value + '")';
    end
    else
    begin
      if (Copy(Racine, 1, 3) = 'PVO') then
      begin
        if LeText <> '' then LeText := LeText + ' AND (PVO_PREDEFINI="' + Pred.Value + '")'
        else LeText := ' (PVO_PREDEFINI="' + Pred.Value + '")';
      end;
    end;
  end;
  if (Org <> nil) and (Org.Text <> '') then
  begin
    if LeText <> '' then LeText := LeText + ' AND (' + Racine + '="' + Org.Text + '")'
    else LeText := ' (' + Racine + '="' + Org.Text + '")';
  end;
// **** Refonte accès V_PGI_env ***** V_PGI_env.ModeFonc remplacé par PgRendModeFonc () *****
  if (Pred <> nil) and (PgRendModeFonc() = 'MULTI') then LeText := LeText + ' AND ##' + Racine + '_PREDEFINI##';
  if (TFmul(Ecran).Name = 'MUL_DUCSAFFECT') and (Affect.Checked = FALSE) then
    LeText := LeText + ' AND ##PCT_PREDEFINI##'; //PT4
  result := LeText;
end;

initialization
  registerclasses([TOF_PGMULAFFECTDCRO]);
end.


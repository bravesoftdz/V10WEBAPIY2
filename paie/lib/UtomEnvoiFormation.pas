{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : ENVOIFORMATION (ENVOIFORMATION)
Mots clefs ... : TOM;ENVOIFORMATION
*****************************************************************
PT1 19/11/2003  V_50 JL  Mise à jour validation OPCA dans table SESSIONSTAGE
---- JL 20/03/2006 modification clé annuaire ----
}
unit UtomEnvoiFormation;

interface

uses Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fe_Main,
  {$ELSE}
  eFiche, eFichList, MainEAGL,entPaie,
  {$ENDIF}
  sysutils, HCtrls, HEnt1, HMsgBox, UTOM, UTob, MailOL,HTB97;

type
  TOM_ENVOIFORMATION = class(TOM)
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;
  private
    Arg: string;
    InitAvoir, InitTotFacture, InitComplement: Double;
    procedure CalculerNet;
    procedure KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GrilleCopierColler;
    procedure SuppressionSessions(Sender: TObject);
    procedure AjoutSessions(Sender: TObject);
    procedure RemplirGrilleSessions;
    procedure CreerFichierFacture(Sender: TObject);
    procedure EnvoiFichierFacture(Sender: TObject);
    procedure CopieNumOPCA(Sender: TObject);
    procedure ReGenereFchier(Sender: TObject);

  end;

implementation

procedure TOM_ENVOIFORMATION.OnNewRecord;
var
  Q: TQuery;
  NumEnvoi: Integer;
begin
  inherited;
  Q := OpenSQL('SELECT MAX(PVF_NUMENVOI) NUMERO FROM ENVOIFORMATION', True);
  if not Q.eof then NumEnvoi := Q.FindField('NUMERO').AsInteger + 1
  else NumEnvoi := 1;
  Ferme(Q);
  SetField('PVF_NUMENVOI', NumEnvoi);
  SetField('PVF_DATEENVOI', IDate1900);
  SetField('PVF_TOTALFACTURE', 0);
  SetField('PVF_COMPLEMENT', 0);
  SetField('PVF_AVOIR', 0);
  SetField('PVF_TOTALNET', 0);
  SetField('PVF_NUMFACTURE', '');
  SetField('PVF_NOMFICHIER', 'TableFormationEnvoiN°' + IntToStr(GetField('PVF_NUMENVOI')));
  SetField('PVF_FICHIERGEN', '-');
end;

procedure TOM_ENVOIFORMATION.OnDeleteRecord;
begin
  inherited;
  ExecuteSQL('UPDATE SSESSIONSTAGE SET PSS_NUMENVOI=-1 WHERE PSS_NUMENVOI=' + GetField('PVF_NUMENVOI') + ''); //DB2
end;

procedure TOM_ENVOIFORMATION.OnUpdateRecord;
begin
  inherited;
end;

procedure TOM_ENVOIFORMATION.OnAfterUpdateRecord;
begin
  inherited;
end;

procedure TOM_ENVOIFORMATION.OnLoadRecord;
begin
  inherited;
  InitAvoir := GetField('PVF_AVOIR');
  InitTotFacture := GetField('PVF_TOTALFACTURE');
  InitComplement := GetField('PVF_COMPLEMENT');
  RemplirGrilleSessions;
end;

procedure TOM_ENVOIFORMATION.OnChangeField(F: TField);
begin
  inherited;
  if Arg = 'FACTURE' then
  begin
    if F.FieldName = 'PVF_AVOIR' then
    begin
      if InitAvoir <> GetField('PVF_AVOIR') then CalculerNet;
    end;
    if F.FieldName = 'PVF_TOTALFACTURE' then
    begin
      if InitTotFacture <> GetField('PVF_TOTALFACTURE') then CalculerNet;
    end;
    if F.FieldName = 'PVF_COMPLEMENT' then
    begin
      if InitComplement <> GetField('PVF_COMPLEMENT') then CalculerNet;
    end;
  end;
end;

procedure TOM_ENVOIFORMATION.OnArgument(S: string);
var
  BAjout, BSupp, BFacture, BMail, BRecopier: TToolBarButton97;
  LaGrilleRecup: THGrid;
begin
  inherited;
  Arg := ReadTokenPipe(S, ';');
  if Arg = 'MODIF' then
  begin
    SetControlProperty('PRECUP', 'TabVisible', False);
    SetControlVisible('BMAIL', False);
    SetControlVisible('BFACTURE', False);
    SetControlVisible('BRECOPIER', True);
    SetControlVisible('BAJOUT', True);
    SetControlVisible('BSUPP', True);
    BAjout := TToolBarButton97(GetControl('BAJOUT'));
    BSupp := TToolBarButton97(GetControl('BSUPP'));
    SetControlVIsible('GBFACTURE', False);
    SetControlVisible('GBETAT', False);
    if BAjout <> nil then BAjout.OnClick := AjoutSessions;
    if BSupp <> nil then BSupp.OnClick := SuppressionSessions;
    BRecopier := TToolBarButton97(GetControl('BRECOPIER'));
    if BRecopier <> nil then
    begin
      BRecopier.Hint := 'Regénérer le fichier';
      BRecopier.OnClick := ReGenereFchier;
    end;
  end
  else
  begin
    SetControlProperty('PRECUP', 'TabVisible', True);
    SetControlVisible('BMAIL', True);
    SetControlVisible('BFACTURE', True);
    SetControlVisible('BRECOPIER', True);
    SetControlVisible('BAJOUT', False);
    SetControlVisible('BSUPP', False);
    BMail := TToolBarButton97(GetControl('BMAIL'));
    BFacture := TToolBarButton97(GetControl('BFACTURE'));
    if BMail <> nil then BMail.OnClick := EnvoiFichierFacture;
    if BFacture <> nil then BFacture.OnClick := CreerFichierFacture;
    BRecopier := TToolBarButton97(GetControl('BRECOPIER'));
    if BRecopier <> nil then
    begin
      BRecopier.OnClick := CopieNumOPCA;
      BRecopier.Hint := 'Récupération des n° OPCA';
    end;
  end;
  LaGrilleRecup := THGrid(GetControl('GRECUPNUM'));
  LaGrilleRecup.OnKeyDown := KeyDown;
end;

procedure TOM_ENVOIFORMATION.OnClose;
begin
  inherited;
end;

procedure TOM_ENVOIFORMATION.OnCancelRecord;
begin
  inherited;
end;

procedure TOM_ENVOIFORMATION.CalculerNet;
var
  Net: Double;
begin
  InitComplement := GetField('PVF_COMPLEMENT');
  InitTotFacture := GetField('PVF_TOTALFACTURE');
  InitAvoir := GetField('PVF_AVOIR');
  Net := InitTotFacture + InitComplement - InitAvoir;
  SetField('PVF_TOTALNET', Net);
end;

procedure TOM_ENVOIFORMATION.KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = 86) and (ssCtrl in Shift) then
    GrilleCopierColler;
end;

procedure TOM_ENVOIFORMATION.GrilleCopierColler;
var
  T: Tob;
begin
  T := TOBLoadFromClipBoard;
  T.Detail[0].Free;
  T.PutGridDetail(THGrid(GetControl('GRECUPNUM')), False, True, '', False);
  T.Free;
end;

procedure TOM_ENVOIFORMATION.SuppressionSessions(Sender: TObject);
var
  St: string;
begin
  St := IntToStr(GetField('PVF_NUMENVOI'));
  AGLLanceFiche('PAY', 'ENVOISESSION_MUL', '', '', 'SUPP;' + St);
  RemplirGrilleSessions;
end;

procedure TOM_ENVOIFORMATION.AjoutSessions(Sender: TObject);
var
  St: string;
begin
  St := IntToStr(GetField('PVF_NUMENVOI'));
  AGLLanceFiche('PAY', 'ENVOISESSION_MUL', '', '', 'AJOUT;' + St);
  RemplirGrilleSessions;
end;

procedure TOM_ENVOIFORMATION.RemplirGrilleSessions;
var
  GSessions: THGrid;
  Q: TQuery;
  TobSessions, TGrille, T: Tob;
  i : Integer;
  OrgCollectP, OrgCollectS : String;
  NAction, Libelle: string;

begin
  GSessions := THGrid(GetControl('GSESSIONS'));
  GSessions.RowCount := 2;
  GSessions.Rows[1].Clear;
  if GSessions = nil then exit;
  Q := OpenSQL('SELECT COUNT (PFO_SALARIE) NBSAL,PSS_ORDRE,PSS_CODESTAGE,PSS_DATEDEBUT,PSS_DATEFIN,PSS_ORGCOLLECTPGU,PSS_ORGCOLLECTSGU' +
    ' FROM SESSIONSTAGE LEFT JOIN FORMATIONS ' +
    'ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_ORDRE=PFO_ORDRE AND PSS_MILLESIME=PFO_MILLESIME ' +
    'WHERE PSS_NUMENVOI=' + IntToStr(GetField('PVF_NUMENVOI')) + '' + //DB2
    ' GROUP BY PSS_ORDRE,PSS_CODESTAGE,PSS_DATEDEBUT,PSS_DATEFIN,PSS_ORGCOLLECTPGU,PSS_ORGCOLLECTSGU', True);
  TobSessions := Tob.Create('Les sessions', nil, -1);
  TobSessions.LoadDetailDB('SESSIONSTAGE', '', '', Q, False);
  Ferme(Q);
  TGrille := Tob.Create('Remplir Grille', nil, -1);
  for i := 0 to TobSessions.Detail.Count - 1 do
  begin
    T := Tob.Create('Remplir Grille', TGrille, -1);
    NAction := IntToStr(TobSessions.detail[i].GetValue('PSS_ORDRE')) + TobSessions.detail[i].GetValue('PSS_CODESTAGE');
    Libelle := RechDom('PGSTAGEFORM', TobSessions.detail[i].GetValue('PSS_CODESTAGE'), False);
    OrgCollectP := TobSessions.Detail[i].GetValue('PSS_ORGCOLLECTPGU');
    OrgCollectS := TobSessions.Detail[i].GetValue('PSS_ORGCOLLECTSGU');
    T.AddChampSupValeur('ACTION', NAction, False);
    T.AddChampSupValeur('LIBELLE', Libelle, False);
    T.AddChampSupValeur('DEBUT', TobSessions.detail[i].GetValue('PSS_DATEDEBUT'), False);
    T.AddChampSupValeur('FIN', TobSessions.detail[i].GetValue('PSS_DATEFIN'), False);
    T.AddChampSupValeur('OPCAP', RechDom('PGORGCOLLECTEUR', OrgCollectP, False), False);
    T.AddChampSupValeur('OPCAS', RechDom('PGORGCOLLECTEUR', OrgCollectS, False), False);
    T.AddChampSupValeur('NBSAL', TobSessions.detail[i].GetValue('NBSAL'), False);
  end;
  TGrille.PutGridDetail(GSEssions, False, True, '', False);
  TobSessions.free;
  TGrille.Free;
end;

procedure TOM_ENVOIFORMATION.CreerFichierFacture(Sender: TObject);
var
  Q: TQuery;
  TobExport, TobSession, T: Tob;
  i: Integer;
  Fichier: string;
begin
  Q := OpenSQL('SELECT PSS_NUMENVOI,PSS_ORDRE,PSS_CODESTAGE,PSS_NUMOCS,PSS_TOTALHT FROM SESSIONSTAGE WHERE PSS_NUMENVOI=' + IntToStr(GetField('PVF_NUMENVOI')) + '', True); //DB2
  TobSession := Tob.Create('Les sessions', nil, -1);
  TobSession.LoadDetailDB('Les sessions', '', '', Q, False);
  Ferme(Q);
  TobExport := Tob.Create('Export données', nil, -1);
  for i := 0 to TobSession.Detail.Count - 1 do
  begin
    T := Tob.Create('Une fille', TobExport, -1);
    T.AddChampSupValeur('NUMENVOI', TobSession.Detail[i].GetValue('PSS_NUMENVOI'), false);
    T.AddChampSupValeur('NUMACTION', IntToStr(TobSession.Detail[i].GetValue('PSS_ORDRE')) + TobSession.Detail[i].GetValue('PSS_CODESTAGE'), false);
    T.AddChampSupValeur('NUMAGEFOS', TobSession.Detail[i].GetValue('PSS_NUMOCS'), false);
    T.AddChampSupValeur('TOTALHT', TobSession.Detail[i].GetValue('PSS_TOTALHT'), false);
  end;
  TobSession.Free;
  {$IFDEF EAGLCLIENT}
  Fichier := VH_PAIE.PGCheminEagl + '\' + GetField('PVF_FICHIEFACTURE') + '.xls';
  {$ELSE}
  Fichier := V_PGI.DatPath + '\' + GetField('PVF_FICHIEFACTURE') + '.xls';
  {$ENDIF}
  if FileExists(Fichier) then DeleteFile(PChar(Fichier));
  TobExport.SaveToExcelFile(Fichier);
  TobExport.Free;
end;

procedure TOM_ENVOIFORMATION.EnvoiFichierFacture(Sender: TObject);
var
  Liste: HTStringList;
begin
  {$IFDEF EAGLCLIENT}
  if not FileExists(VH_PAIE.PGCheminEagl + '\' + GetField('PVF_FICHIEFACTURE') + '.xls') then
    {$ELSE}
  if not FileExists(V_PGI.DatPath + '\' + GetField('PVF_FICHIEFACTURE') + '.xls') then
    {$ENDIF}
  begin
    PGIBox('Le fichier n''existe pas', Ecran.Caption);
    Exit;
  end;
  Liste := HTStringList.Create;
  Liste.Add('Veuillez trouver ci-joint le fichier');
  {$IFDEF EAGLCLIENT}
  SendMail('Envoi facturation formation', '', '', Liste, VH_PAIE.PGCheminEagl + '\' + GetField('PVF_FICHIEFACTURE') + '.xls', FALSE);
  {$ELSE}
  SendMail('Envoi facturation formation', '', '', Liste, V_PGI.DatPath + '\' + GetField('PVF_FICHIEFACTURE') + '.xls', FALSE);
  {$ENDIF}
  Liste.Clear;
  Liste.Free;
  ForceUpdate;
  SetField('PVF_TRANSFACT', 'X');
end;

procedure TOM_ENVOIFORMATION.CopieNumOPCA(Sender: TObject);
var
  LaGrilleRecup: THGrid;
  TobSessions, TS: Tob;
  Q: TQuery;
  Longueur,  i: Integer;
  NumOrdre, CodeStage, NumAction: string;
begin
  LaGrilleRecup := THGrid(GetControl('GRECUPNUM'));
  LaGrilleRecup.OnKeyDown := KeyDown;
  Q := OpenSQL('SELECT * FROM SESSIONSTAGE WHERE PSS_NUMENVOI=' + IntToStr(GetField('PVF_NUMENVOI')) + '', True); //DB2
  TobSessions := Tob.Create('SESSIONSTAGE', nil, -1);
  TobSessions.LoadDetailDB('SESSIONSTAGE', '', '', Q, False);
  Ferme(Q);
  for i := 1 to LaGrilleRecup.RowCount - 1 do
  begin
    NumAction := LaGrilleRecup.CellValues[0, i];
    if NumAction = '' then Continue;
    Longueur := Length(NumAction);
    if Longueur < 7 then Continue;
    NumOrdre := Copy(NumAction, 1, Longueur - 6);
    if not IsNumeric(NumOrdre) then Continue;
    CodeStage := Copy(NumAction, Longueur - 5, Longueur);
    TS := TobSessions.FindFirst(['PSS_ORDRE', 'PSS_CODESTAGE'], [NumOrdre, CodeStage], False);
    if TS <> nil then
    begin
      if GetField('PVF_AVECCTPEDAG') = 'X' then TS.Putvalue('PSS_NUMOCP', LaGrilleRecup.CellValues[1, i]);
      if GetField('PVF_AVECCTSAL') = 'X' then TS.Putvalue('PSS_NUMOCS', LaGrilleRecup.CellValues[1, i]);
      TS.UpdateDB(False);
      TS.Putvalue('PSS_VALIDORG', 'X'); //PT1
    end;
  end;
  TobSessions.Free;
  ForceUpdate;
  SetField('PVF_RETOUROPCA', 'X');
end;

procedure TOM_ENVOIFORMATION.ReGenereFchier(Sender: TObject);
begin
  AGLLanceFiche('PAY', 'PREP_ENVOIFORM', '', '', 'REGENERE;' + IntToStr(GetField('PVF_NUMENVOI')) + ';;');
end;

initialization
  registerclasses([TOM_ENVOIFORMATION]);
end.


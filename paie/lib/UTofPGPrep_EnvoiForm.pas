{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 18/02/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGPREP_ENVOIFORM ()
Mots clefs ... : TOF;PGPREP_ENVOIFORM
*****************************************************************
---- JL 20/03/2006 modification clé annuaire ----
---- JL 17/10/2006 Modification contrôle des exercices de formations -----
}
unit UTofPGPrep_EnvoiForm;

interface

uses StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DBGrids, HDB, DBCtrls, Mul,
  {$ELSE}
  emul,
  {$ENDIF}
  EntPAie,forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, HQry, UTOB, ParamDat, PGOutilsFormation, Vierge, HTB97;

type
  TOF_PGPREP_ENVOIFORM = class(TOF)
    procedure OnArgument(S: string); override;
  private
    DateDebutMul, DateFinMul: TDateTime;
    NumeroEnvoi: Integer;
    Arg: string;
    procedure CreerEnvoi(Sender: TObject);
    procedure AjoutSessions();
    procedure CreerFichier();
    procedure MAjEnvoiForm(Cout: Double);
    procedure DateElipsisclick(Sender: TObject);
    procedure SuppEnvoiForm();
  end;

implementation
uses UTofPGMulEnvoiSession;

procedure TOF_PGPREP_ENVOIFORM.OnArgument(S: string);
var
  Q: TQuery;
  BGenere: TToolBarButton97;
  THDate: THEdit;
begin
  inherited;
  Arg := ReadTokenPipe(S, ';');
  NumeroEnvoi := StrToInt(ReadTokenPipe(S, ';'));
  if Arg <> 'REGENERE' then
  begin
    DateDebutMul := StrToDate(ReadTokenPipe(S, ';'));
    DateFinMul := StrToDate(ReadTokenPipe(S, ';'));
  end;
  if NumeroEnvoi <= 0 then
  begin
    Q := OpenSQL('SELECT MAX(PVF_NUMENVOI) NUMERO FROM ENVOIFORMATION', True);
    if not Q.eof then NumeroEnvoi := Q.FindField('NUMERO').AsInteger + 1
    else NumeroEnvoi := 1;
    Ferme(Q);
    SetControlText('FICHIER', 'EnvoiFormationN°' + IntToStr(NumeroEnvoi));
    SetControlText('CKCOUTPEDAG', '-');
    SetControlText('CKCOUTSAL', '-');
  end
  else
  begin
    Q := OpenSQL('SELECT PVF_LIBELLE,PVF_NOMFICHIER,PVF_AVECCTPEDAG,PVF_AVECCTSAL,PVF_PREPARELE,PVF_ORGCOLLECTGU ' +
      'FROM ENVOIFORMATION WHERE PVF_NUMENVOI=' + IntToStr(NumeroEnvoi) + '', True);
    if not Q.Eof then
    begin
      SetControlText('PREPARELE', DateToStr(Q.FindField('PVF_PREPARELE').AsdateTime));
      if Q.FindField('PVF_AVECCTPEDAG').AsString = 'X' then SetControlChecked('CKCOUTPEDAG', True)
      else SetControlChecked('CKCOUTPEDAG', False);
      if Q.FindField('PVF_AVECCTSAL').AsString = 'X' then SetControlChecked('CKCOUTSAL', True)
      else SetControlChecked('CKCOUTSAL', False);
      SetControlText('ORGCOLLECT', Q.FindField('PVF_ORGCOLLECTGU').AsString);
      SetControlText('FICHIER', Q.FindField('PVF_NOMFICHIER').AsString);
      SetControlText('LIBELLE', Q.FindField('PVF_LIBELLE').AsString);
    end;
    Ferme(Q);
  end;
  BGenere := TToolBarButton97(GetControl('BGENERE'));
  if BGenere <> nil then BGenere.OnClick := CreerEnvoi;
  SetControlText('DATEPREP', DateToStr(Date));
  THDate := THEdit(GetControl('PVF_DATEENVOI'));
  if THDate <> nil then THDate.OnElipsisClick := DateElipsisClick;

end;

procedure TOF_PGPREP_ENVOIFORM.CreerEnvoi(Sender: TObject);
begin
  if GetControlText('ORGCOLLECT') = '' then
  begin
    PGIBox('Vous devez choisir l''OPCA', Ecran.Caption);
    Exit;
  end;
  if (GetControlText('CKCOUTPEDAG') = '-') and (GetControlText('CKCOUTSAL') = '-') then
  begin
    PGIBox('Vous devez choisir le type de coût', Ecran.Caption);
    Exit;
  end;
  SetActiveTabSheet('PTRAITEMENT');
  if Arg = 'REGENERE' then CreerFichier
  else AjoutSessions;
end;

procedure TOF_PGPREP_ENVOIFORM.AjoutSessions();
var
  Stage, Millesime: string;
  Session: Integer;
  TobSession, TS: Tob;
  Q: TQuery;
  i: Integer;
begin
  Q := OpenSQL('SELECT * FROM SESSIONSTAGE WHERE PSS_DATEDEBUT>="' + UsDateTime(DateDebutMul) + '"' +
    ' AND PSS_DATEFIN<="' + UsDateTime(DateFinMul) + '"', True);
  TobSession := Tob.Create('SESSIONSTAGE', nil, -1);
  TobSession.LoadDetailDB('SESSIONSTAGE', '', '', Q, False);
  Ferme(Q);
  if (Grille = nil) then Exit;
  if (Grille.NbSelected = 0) and (not Grille.AllSelected) then
  begin
    MessageAlerte('Aucun élément sélectionné');
    exit;
  end;
  {$IFDEF EAGLCLIENT}
  if (TFM.bSelectAll.Down) then
    TFM.Fetchlestous;
  {$ENDIF}
  if ((Grille.nbSelected) > 0) and (not Grille.AllSelected) then
  begin
    for i := 0 to Grille.NbSelected - 1 do
    begin
      Grille.GotoLeBookmark(i);
      {$IFDEF EAGLCLIENT}
      TFM.Q.TQ.Seek(Grille.Row - 1);
      {$ENDIF}
      Stage := q_mul.FindField('PSS_CODESTAGE').AsString;
      Session := q_mul.FindField('PSS_ORDRE').AsInteger;
      Millesime := q_mul.FindField('PSS_MILLESIME').AsString;
      TS := TobSession.FindFirst(['PSS_CODESTAGE', 'PSS_ORDRE', 'PSS_MILLESIME'], [Stage, Session, Millesime], False);
      if TS <> nil then
      begin
        if Arg = 'SUPP' then TS.PutValue('PSS_NUMENVOI', 0)
        else TS.PutValue('PSS_NUMENVOI', NumeroEnvoi);
        TS.UpdateDB(False);
      end;
    end;
    Grille.ClearSelected;
  end;
  if (Grille.AllSelected = TRUE) then
  begin
    q_mul.First;
    while not q_mul.EOF do
    begin
      Stage := q_mul.FindField('PSS_CODESTAGE').AsString;
      Session := q_mul.FindField('PSS_ORDRE').AsInteger;
      Millesime := q_mul.FindField('PSS_MILLESIME').AsString;
      TS := TobSession.FindFirst(['PSS_CODESTAGE', 'PSS_ORDRE', 'PSS_MILLESIME'], [Stage, Session, Millesime], False);
      if TS <> nil then
      begin
        if Arg = 'SUPP' then TS.PutValue('PSS_NUMENVOI', 0)
        else TS.PutValue('PSS_NUMENVOI', NumeroEnvoi);
        TS.UpdateDB(False);
      end;
      q_mul.Next;
    end;
    Grille.AllSelected := False;
  end;
  TobSession.Free;
  CreerFichier;
end;

procedure TOF_PGPREP_ENVOIFORM.CreerFichier();
var
  i: Integer;
  CoutSal, CoutTotal : Double;
  Fichier: string;
  Trace, TraceErr: TListBox;
  OPCACP, OPCACS: String;                         
begin
  Trace := TListBox(GetControl('LSTBXTRACE'));
  TraceErr := TListBox(GetControl('LSTBXERREUR'));
  {$IFDEF EAGLCLIENT}
  Fichier := VH_PAIE.PGCheminEagl + '\' + GetControlText('FICHIER') + '.xls';
  {$ELSE}
  Fichier := V_PGI.DatPath + '\' + GetControlText('FICHIER') + '.xls';
  {$ENDIF}
  ChargeTobCreerFichierForm(NumeroEnvoi);
  OPCACP := '';
  OPCACS := '';
  if GetControlText('CKCOUTPEDAG') = 'X' then OPCACP := GetControlText('ORGCOLLECT');
  if GetControlText('CKCOUTSAL') = 'X' then OPCACS := GetControlText('ORGCOLLECT');
  ExecuteSQL('UPDATE SESSIONSTAGE SET PSS_TOTALHT=0 WHERE PSS_NUMENVOI=' + IntToStr(NumeroEnvoi) + ''); //DB2
  RecupDonneesEnvoiAnim(NumeroEnvoi, OPCACP, OPCACS, Trace, TraceErr);
  RecupDonneesEnvoiSta(NumeroEnvoi, OPCACP, OPCACS, Trace, TraceErr);
  if TobPrepAGEFOS <> nil then TobPrepAGEFOS.Free;
  CoutTotal := 0;
  for i := 0 to TobPasserelle.Detail.Count - 1 do
  begin
    CoutSal := TobPasserelle.Detail[i].GetValue('TOTALHT');
    CoutTotal := CoutTotal + CoutSal;
  end;
  TobPasserelle.Detail.Sort('ACTION;TYPE;MATRICULE');
  if FileExists(Fichier) then DeleteFile(PChar(Fichier));
  TobPasserelle.SaveToExcelFile(Fichier);
  TobPasserelle.Free;
  if arg <> '' then SuppEnvoiForm;
  MajEnvoiForm(CoutTotal);
end;

procedure TOF_PGPREP_ENVOIFORM.MajEnvoiForm(Cout: Double);
var
  TobMere, TobEnvoi: Tob;
  Q: TQuery;
  Millesime: string;
  DD,DF : TDateTime;
begin
  Millesime := RendMillesimeRealise(DD,DF);
  TobMere := Tob.Create('Les envoi', nil, -1);
  TobEnvoi := Tob.Create('ENVOIFORMATION', TobMere, -1);
  TobEnvoi.PutValue('PVF_NUMENVOI', NumeroEnvoi);
  TobEnvoi.PutValue('PVF_LIBELLE', GetControlText('LIBELLE'));
  TobEnvoi.PutValue('PVF_NUMFACTURE', '');
  TobEnvoi.PutValue('PVF_DATEENVOI', IDate1900);
  TobEnvoi.PutValue('PVF_ORGCOLLECTGU', GetControlText('ORGCOLLECT'));
  TobEnvoi.PutValue('PVF_AVECCTPEDAG', GetControlText('CKCOUTPEDAG'));
  TobEnvoi.PutValue('PVF_AVECCTSAL', GetControlText('CKCOUTSAL'));
  TobEnvoi.PutValue('PVF_TOTALFACTURE', Cout);
  TobEnvoi.PutValue('PVF_COMPLEMENT', 0);
  TobEnvoi.PutValue('PVF_AVOIR', 0);
  TobEnvoi.PutValue('PVF_NBSESSIONS', 0);
  TobEnvoi.PutValue('PVF_TOTALNET', Cout);
  TobEnvoi.PutValue('PVF_NOMFICHIER', GetControlText('FICHIER'));
  TobEnvoi.PutValue('PVF_ENVOIREEL', '-');
  TobEnvoi.PutValue('PVF_SIRETDO', '');
  TobEnvoi.PutValue('PVF_FICHIEFACTURE', 'FactureFormationEnvoiN°' + IntToStr(NumeroEnvoi));
  TobEnvoi.PutValue('PVF_TAILLEFIC', 0);
  TobEnvoi.PutValue('PVF_PREPARELE', StrToDate(GetControlText('DATEPREP')));
  TobEnvoi.PutValue('PVF_EMETSOC', '');
  TobEnvoi.PutValue('PVF_SUPPORTEMIS', '');
  TobEnvoi.PutValue('PVF_STATUTENVOI', '005');
  TobEnvoi.PutValue('PVF_NUMDOSSIER', '');
  TobEnvoi.PutValue('PVF_ENVOYEPAR', '');
  TobEnvoi.PutValue('PVF_CODAPPLI', '');
  TobEnvoi.PutValue('PVF_RETOUROPCA', '-');
  TobEnvoi.PutValue('PVF_TRANSFACT', '-');
  TobEnvoi.PutValue('PVF_RETOURFAC', '-');
  TobEnvoi.PutValue('PVF_MILLESSOC', Millesime);
  TobEnvoi.InsertOrUpdateDB(False);
  TobMere.Free;
end;

procedure TOF_PGPREP_ENVOIFORM.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGPREP_ENVOIFORM.SuppEnvoiForm();
begin
  ExecuteSQL('DELETE FROM ENVOIFORMATION WHERE PVF_NUMENVOI="' + IntToStr(NumeroEnvoi) + '"');
end;

initialization
  registerclasses([TOF_PGPREP_ENVOIFORM]);
end.


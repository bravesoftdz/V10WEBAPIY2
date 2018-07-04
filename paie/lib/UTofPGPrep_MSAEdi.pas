{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 08/04/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGPREP_CONGESCPEC ()
Mots clefs ... : TOF;PGPREP_CONGESCPEC
*****************************************************************
PT1 05/04/2005 JL V_60 Corrections envoi réél dans fichier
PT2 13/04/2005 JL V_60 Corrections envoi multi dossier
}
unit UTofPGPrep_MSAEdi;

interface

uses Windows, StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
  Dialogs, ed_tools, Mailol,
  {$IFDEF EAGLCLIENT}
  UtileAGL, eFiche, eMul,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fiche, Qre, DbGrids, HDB,
  {$ENDIF}
  Grids, HCtrls, HEnt1, EntPaie, HMsgBox, UTOF, UTOB, UTOM, Vierge, HQry, HStatus,
  AGLInit, ParamSoc;

type
  TOF_PGPREP_MSAEDI = class(TOF)
    procedure OnArgument(Arguments: string); override;
  private
    Trace, TraceErr: TListBox;
    QMul: TQUERY;
    Pan: TPageControl;
    Tbsht: TTabSheet;
    FileE, FileL: string;
    Reel, Emet, Monnaie, Support, CPEmmetteur: string;
    FLect, FEcrt: TextFile;
    BtnLance, BImprime: TToolBarButton97;
    NumEnvoi: Integer;
    NbCertificats: Integer;
    BaseConges: Double;
    DateDebutConges, DateFinConges: TDateTime;
    procedure LancePrepMSA(Sender: TObject);
    procedure ImprimeClick(Sender: TObject);
    procedure AbandonTraitement;
    function TraiteFichier: Boolean;
    function ReecritChrono(Chrono: Integer): Boolean;
    procedure VerifPeriodes;
  end;

implementation
uses UTofPG_MulEnvoiSocial;

procedure TOF_PGPREP_MSAEDI.OnArgument(Arguments: string);
var
  S, St: string;
  F: TFVierge;
  Q: TQuery;
begin
  inherited;
  st := Trim(Arguments);
  S := ReadTokenSt(st);
  SetControlText('LBLTYPE', GetControlText('LBLTYPE') + ' ' + RechDom('PGENVOISOCIAL', S, FALSE));
  S := ReadTokenSt(st);
  SetControlText('LBLMILLESIME', GetControlText('LBLMILLESIME') + ' ' + RechDom('PGANNEE', S, FALSE));
  S := ReadTokenSt(st);
  SetControlText('LBLPERIODICITE', GetControlText('LBLPERIODICITE') + ' ' + RechDom('PGPERIODICITEDUCS', S, FALSE));
  S := ReadTokenSt(st);
  SetControlText('LBLDESTINATAIRE', GetControlText('LBLDESTINATAIRE') + ' ' + RechDom('PGINSTITUTION', S, FALSE));
  Emet := ReadTokenSt(st);
  SetControlText('LBLEMETTEUR', GetControlText('LBLEMETTEUR') + ' ' + Copy(RechDom('PGEMETTEURSOC', Emet, FALSE), 1, 25));
  S := ReadTokenSt(st);
  Support := S;
  Q := OpenSQL('SELECT PET_CODEPOSTAL FROM EMETTEURSOCIAL WHERE PET_EMETTSOC="' + Emet + '"', True);
  if not Q.Eof then CPEmmetteur := Q.FindField('PET_CODEPOSTAL').AsString
  else CPEmmetteur := '';
  Ferme(Q);
  SetControlText('LBLSUPPORT', GetControlText('LBLSUPPORT') + ' ' + RechDom('PGSUPPORTEDI', S, FALSE));
  Reel := ReadTokenSt(st);
  if Reel = 'X' then SetControlProperty('LBLREEL', 'Checked', TRUE)
  else SetControlProperty('LBLREEL', 'Checked', FALSE);
  S := ReadTokenSt(st); // Monnaie
  Monnaie := S;
  SetControlText('LBLMONNAIE', GetControlText('LBLMONNAIE') + ' ' + RechDom('PGMONNAIE', S, FALSE));
  BtnLance := TToolbarButton97(GetControl('BLANCE'));
  if BtnLance <> nil then
    BtnLance.OnClick := LancePrepMSA;
  BImprime := ttoolbarbutton97(getcontrol('BIMPRIMER'));
  if Bimprime <> nil then
    Bimprime.Onclick := ImprimeClick;
  if not (Ecran is TFVierge) then exit;
  F := TFVierge(Ecran);
  {$IFDEF EAGLCLIENT}
  QMUL := THQuery(F.FMULQ).TQ;
  {$ELSE}
  QMUL := F.FMULQ;
  {$ENDIF}
  VerifPeriodes;
end;

procedure TOF_PGPREP_MSAEDI.LancePrepMSA(Sender: TObject);
var
  ListeJ: HTStrings;
  St: string;
  Chrono, i, reponse, Rep: integer;
  okok: Boolean;
  QQ, Q: TQuery;
begin
  Pan := TPageControl(GetControl('PAGES'));
  Tbsht := TTabSheet(GetControl('TBSHTTRACE'));
  Trace := TListBox(GetControl('LSTBXTRACE'));
  TraceErr := TListBox(GetControl('LSTBXERROR'));
  DateDebutConges := IDate1900;
  DateFinConges := IDate1900;
  if (Trace = nil) or (TraceErr = nil) then
  begin
    PGIBox('La préparation MSA EDI ne peut pas être lancée', 'Car les composants trace ne sont pas disponibles');
    exit;
  end;
  if GetControlText('NOMFIC') = '' then
  begin
    PgiBox('Vous devez indiquer un nom de fichier', Ecran.Caption);
    exit;
  end;
  if CPEmmetteur = '' then
  begin
    PgiBox('Vous devez renseigner le code postal de l''émeteur', Ecran.Caption);
    exit;
  end;
  BtnLance.Enabled := FALSE;

  if (Pan <> nil) and (Tbsht <> nil) then Pan.ActivePage := Tbsht;
  if QMul = nil then
  begin
    PGIBox('Erreur sur la liste des fichiers à traiter', 'Abandon du traitement');
    exit;
  end;
  if (Grille = nil) then
  begin
    MessageAlerte('Grille de données inexistantes');
    exit;
  end;
  if (Grille.NbSelected = 0) and (not Grille.AllSelected) then
  begin
    MessageAlerte('Aucun élément sélectionné');
    exit;
  end;
  Trace.Items.Add('Debut du traitement');
  {$IFNDEF EAGLCLIENT}
  if support <> 'DTK' then
  begin
    FileE := '$DAT\' + GetControlText('NOMFIC');
    FileE := ChangeStdDatPath(FILEE);
  end
    {$ELSE}
  if support <> 'DTK' then FileE := VH_Paie.PGCheminEagl + '\' + GetControlText('NOMFIC')
    {$ENDIF}
  else FileE := 'A:\' + GetControlText('NOMFIC');
  if support = 'TRA' then
    FileE := VH_Paie.PGCheminEagl + '\' + GetControlText('NOMFIC');
  if FileExists(FileE) then
  begin
    reponse := HShowMessage('5;;Voulez-vous supprimer votre fichier MSA ' + FileE + ';Q;YN;Y;N', '', '');
    if reponse = 6 then DeleteFile(PChar(FileE))
    else exit;
  end;
  AssignFile(FEcrt, FileE);
  ReWrite(FEcrt);
  QQ := OpenSql('SELECT MAX(PES_NUMEROENVOI) FROM ENVOISOCIAL', TRUE);
  if not QQ.EOF then NumEnvoi := QQ.Fields[0].AsInteger + 1 // beurk à supprimer
  else NumEnvoi := 1;
  Ferme(QQ);
  if ((Grille.nbSelected) > 0) and (not Grille.AllSelected) then
  begin
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', i, FALSE, TRUE);
    InitMove(Grille.NbSelected, '');
    for i := 0 to Grille.NbSelected - 1 do
    begin
      Grille.GotoLeBOOKMARK(i);
      {$IFDEF EAGLCLIENT}
      QMul.Seek(Grille.Row - 1);
      {$ENDIF}
      MoveCur(False);
      Chrono := QMul.findfield('PES_CHRONOMESS').asInteger;
      //DEBUT PT1
      {$IFNDEF EAGLCLIENT}
      FileL := V_PGI.DatPath + '\' + QMul.findfield('PES_FICHIERRECU').AsString;
      {$ELSE}
      FileL := VH_Paie.PGCheminEagl + '\' + QMul.findfield('PES_FICHIERRECU').AsString;
      {$ENDIF}
      //FIN PT1
      okok := TraiteFichier;
      if not okok then
      begin
        AbandonTraitement;
        break;
      end;
      st := QMul.findfield('PES_LIBELLE').asString;
      MoveCurProgressForm(St);
    end;
    FiniMove;
    FiniMoveProgressForm;
  end;
  if (Grille.AllSelected = TRUE) then
  begin
    {$IFDEF EAGLCLIENT}
    if (TFMul(Ecran).bSelectAll.Down) then
      TFMul(Ecran).Fetchlestous;
    {$ENDIF}
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', i, FALSE, TRUE);
    InitMove(QMul.RecordCount, '');
    QMul.First;
    while not QMul.EOF do
    begin
      MoveCur(False);
      Chrono := QMul.findfield('PES_CHRONOMESS').asInteger;
      //DEBUT PT1
      {$IFNDEF EAGLCLIENT}
      FileL := V_PGI.DatPath + '\' + QMul.findfield('PES_FICHIERRECU').AsString;
      {$ELSE}
      FileL := VH_Paie.PGCheminEagl + '\' + QMul.findfield('PES_FICHIERRECU').AsString;
      {$ENDIF}
      //FIN PT1
      okok := TraiteFichier;
      if not okok then
      begin
        AbandonTraitement;
        break;
      end;
      st := QMul.findfield('PES_LIBELLE').asString;
      MoveCurProgressForm(St);
      QMul.Next;
    end;
    FiniMove;
    FiniMoveProgressForm;
  end;
  if not OkOk then Exit;
  // Traitement de mise à jour des enregistrements de la table envoisocial
  Trace.Items.Add('Mise à jour de la liste des fichiers envoyés');
  if ((Grille.nbSelected) > 0) and (not Grille.AllSelected) then
  begin
    for i := 0 to Grille.NbSelected - 1 do
    begin
      Grille.GotoLeBOOKMARK(i);
      Chrono := QMul.findfield('PES_CHRONOMESS').asInteger;
      ReecritChrono(Chrono);
    end;
    Grille.ClearSelected;
  end;
  if (Grille.AllSelected = TRUE) then
  begin
    QMul.First;
    while not QMul.EOF do
    begin
      Chrono := QMul.findfield('PES_CHRONOMESS').asInteger;
      ReecritChrono(Chrono);
      QMul.Next;
    end;
    Grille.AllSelected := False;
  end;
  CloseFile(FEcrt);
  Trace.Items.Add('Fin de mise à jour de la liste des fichiers envoyés');
  if TraceErr.items.Count >= 1 then Trace.Items.Add('Fin du traitement, consultez vos anomalies')
  else Trace.Items.Add('Fin de traitement');
  Pan.ActivePage := Tbsht;
  // Traitement specifique envoi fichier par messagerie Outlook ou autre en fonction des préférences
  if support = 'TEL' then
  begin
    rep := mrYes;
    if TraceErr.Items.Count - 1 > 0 then
    begin
      Rep := PGIAsk('Vous avez des erreurs#13#10Voulez vous quand même envoyer votre fichier', 'Emission par Internet');
    end;
    if rep = mrYes then
    begin
      ListeJ := HTStringList.Create;
      ListeJ.Add('Veuillez trouver ci-joint notre déclaration des salaires MSA');
      ListeJ.Add('Cordialement');
      SendMail('MSA', '', '', ListeJ, FileE, FALSE);
      ListeJ.Clear;
      ListeJ.Free;
    end;
  end;
  if support <> 'TEL' then
    PGIBox('Vous devez émettre le fichier ' + FileE, Ecran.Caption)
  else
    PGIBox('Fin de la procédure', Ecran.Caption);
end;

procedure TOF_PGPREP_MSAEDI.ImprimeClick(Sender: TObject);
var
  MPages: tpagecontrol;
begin
  {$IFNDEF EAGLCLIENT}
  MPages := TPageControl(getcontrol('PAGES'));
  if MPages <> nil then
    PrintPageDeGarde(MPages, TRUE, TRUE, FALSE, Ecran.Caption, 0);
  {$ENDIF}
end;

procedure TOF_PGPREP_MSAEDI.AbandonTraitement;
begin
  TraceErr.items.add('Une erreur est survenue lors de la conception du fichier');
  TraceErr.items.add('Le traitement est interrompu');
  Trace.items.add('Le traitement est abandonné, vérifiez vos erreurs !');
end;

function TOF_PGPREP_MSAEDI.TraiteFichier: Boolean;
var
  S, St: string;
begin
  result := FALSE;
  Trace.Items.Add('Traitement déclaration des salaires de l''entreprise ' + QMul.findfield('PES_LIBELLE').AsString);
  if not FileExists(FileL) then
  begin
    PgiBox('Fichier MSA inexistant pour l''entreprise ' + FileL, Ecran.caption);
    TraceErr.Items.Add('Fichier MSA inexistant pour l''entreprise ' + QMul.findfield('PES_FICHIERRECU').AsString);
    exit;
  end;
  AssignFile(FLect, FileL);
  {$I-}Reset(FLect);
  {$I+}if IoResult <> 0 then
  begin
    PGIBox('Fichier inaccessible : ' + FileL, 'Abandon du traitement');
    TraceErr.Items.Add('Fichier inaccessible : ' + FileL);
    Exit;
  end;

  while not eof(FLect) do
  begin
    {$I-}Readln(FLect, S);
    {$I+}if IoResult <> 0 then
    begin
      PGIBox('Erreur de lecture du fichier : ' + FileL, 'Abandon du traitement');
      TraceErr.Items.Add('Erreur de lecture du fichier : ' + FileL);
      closeFile(FLect);
      Exit;
    end;
    if Copy(S, 34, 4) = 'PE11' then //PT1
    begin
      if Reel = 'X' then S := Copy(S, 1, 58) + 'R' + Copy(S, 60, 141)
      else S := Copy(S, 1, 58) + 'T' + Copy(S, 60, 141)
    end;
    Write(FEcrt, S);
    WriteLN(FEcrt, '');
  end;
  closeFile(FLect);
  Trace.Items.Add('Fin de traitement de l''entreprise ' + QMul.findfield('PES_LIBELLE').AsString);
  result := TRUE;
end;

function TOF_PGPREP_MSAEDI.ReecritChrono(Chrono: Integer): Boolean;
var
  st: string;
begin
  st := 'UPDATE ENVOISOCIAL SET PES_ENVOIREEL="' + Reel + '", PES_NUMEROENVOI=' +
    IntToStr(NumEnvoi) + ', PES_FICHIEREMIS="' + GetControlText('NOMFIC') + '", PES_EMETSOC="' +
    Emet + '", PES_SUPPORTEMIS="' + Support + '", PES_STATUTENVOI="TRA", PES_PREPARELE="' +
    UsDateTime(Date) + '", PES_ENVOYEPAR="' + V_PGI.User + '", PES_MTPAYER="' + FloatToStr(BaseConges) + '" WHERE PES_CHRONOMESS="' +
    IntToStr(Chrono) + '"';
  ExecuteSql(st);
  Result := True;
end;

procedure TOF_PGPREP_MSAEDI.VerifPeriodes;
var
  DateDebutEnvoiPrec, DateDebutEnvoi: TDateTime;
  i: Integer;
  Annee, TrimestreDeb, Trimestre: string;
  Q : TQuery;
begin
  DateDebutEnvoi := IDate1900;
  DateDebutEnvoiPrec := IDate1900;
  if ((Grille.nbSelected) > 0) and (not Grille.AllSelected) then
  begin
    for i := 0 to Grille.NbSelected - 1 do
    begin
      Grille.GotoLeBOOKMARK(i);
      {$IFDEF EAGLCLIENT}
      QMul.Seek(Grille.Row - 1);
      {$ENDIF}
      Q := OpenSQL('SELECT PES_DATEDEBUT FROM ENVOISOCIAL WHERE PES_CHRONOMESS='+IntToStr(QMul.findfield('PES_CHRONOMESS').AsInteger),True);
      If Not Q.Eof then DateDebutEnvoi := Q.findfield('PES_DATEDEBUT').AsDateTime
      else DateDebutEnvoi := IDate1900;
      Ferme(Q);
      if DateDebutEnvoiPrec <> IDate1900 then
      begin
        if DateDebutEnvoi <> DateDebutEnvoiPrec then
        begin
          PGIBox('Traitement impossible car les envois sélectionnés ne sont pas sur la même période', Ecran.Caption);
          TFFiche(Ecran).Close;
          exit;
        end;
      end;
      DateDebutEnvoiPrec := DateDebutEnvoi;
    end;
  end;
  if (Grille.AllSelected = TRUE) then
  begin
    QMul.First;
    while not QMul.EOF do
    begin
      Q := OpenSQL('SELECT PES_DATEDEBUT FROM ENVOISOCIAL WHERE PES_CHRONOMESS='+IntToStr(QMul.findfield('PES_CHRONOMESS').AsInteger),True);
      If Not Q.Eof then DateDebutEnvoi := Q.findfield('PES_DATEDEBUT').AsDateTime
      else DateDebutEnvoi := IDate1900;
      if DateDebutEnvoiPrec <> IDate1900 then
      begin
        if DateDebutEnvoi <> DateDebutEnvoiPrec then
        begin
          PGIBox('Traitement impossible car les envois sélectionnés ne sont pas sur la même période', Ecran.Caption);
          TFFiche(Ecran).Close;
          exit;
        end;
      end;
      DateDebutEnvoiPrec := DateDebutEnvoi;
      QMul.Next;
    end;
  end;
  Annee := Copy(FormatDateTime('yyyy', DateDebutEnvoi), 3, 2);
  TrimestreDeb := FormatDateTime('ddmm', DateDebutEnvoi);
  if TrimestreDeb = '0101' then Trimestre := '1'
  else if TrimestreDeb = '0104' then Trimestre := '2'
  else if TrimestreDeb = '0107' then Trimestre := '3'
  else if TrimestreDeb = '0110' then Trimestre := '4';
  SetControlText('NOMFIC', 'DS' + Copy(CPEmmetteur, 1, 2) + '2' + Annee + Trimestre + '.DAT');
end;

initialization
  registerclasses([TOF_PGPREP_MSAEDI]);
end.


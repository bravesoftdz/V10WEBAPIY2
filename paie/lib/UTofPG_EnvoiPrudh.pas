{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 12/11/2001
Modifié le ... : 15/11/2001
Description .. : Unit de lancement de la confection du fichiers
Suite ........ : prud'hom
Suite ........ : Ecriture et controle entete emetteur
Suite ........ : Mise à jour des champs de la table des envois du social
Mots clefs ... : PAIE;PRUDH
*****************************************************************}
{
PT1   : 11/02/2003 VG V_42 On permet de choisir dorénavant un répertoire de
                           travail pour déposer le fichier (celui défini dans
                           les paramètres société) - FQ N°10471
}
unit UTofPG_EnvoiPrudh;

interface
uses Windows, StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
  Dialogs, ed_tools, Mailol,
  {$IFDEF EAGLCLIENT}
  UtileAGL, eFiche, emul,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fiche, Qre, mul,
  EdtREtat,
  {$ENDIF}
  Grids, HCtrls, HEnt1, EntPaie, HMsgBox, UTOF, UTOB, Vierge, HQry, HStatus,
  AGLInit, PgOutils,PgOutils2;
type
  TOF_PG_Envoi_Prudh = class(TOF)
  private
    Pan: TPageControl;
    Tbsht: TTabSheet;
    Trace, TraceErr: TListBox;
    QMul: TQUERY; // Query recuperee du mul
    FEcrt, FLect, FZ: TextFile; // Canaux des fichiers
    FileN, FileM, FileZ: string; // Nom des fichiers
    NumEnvoi: Integer; // Numero de l'envoi
    Reel, Emet, Support: string; // Reel = X sinon - ...
    BtnLance: TToolbarButton97;
    procedure LanceEnvoi_Prudh(Sender: TObject);
    procedure PreparBorderAcc;
    procedure ImprimeClick(Sender: TObject);
    function TraiteFicPrudh: Boolean;
    function PrudhEmetteur: Boolean;
    function EcritFinFich: Boolean;
    procedure ReecritChronoPrudh(Chrono: Integer);
    procedure AbandonTraitementPrudh;
  public
    procedure OnArgument(Arguments: string); override;
  end;

procedure EditBorderAcc(NomFic: string);

implementation
uses
  UTofPG_MulEnvoiSocial;

type
  E000 = record // ENREGISTREMENT DE TYPE DEBUT DE FICHIER
    Indicatif: string[19];
    TypeEnr: string[3];
    L1: string[1];
    siret: string[14];
    L2: string[200];
    L3: string[200];
    L4: string[127];
  end;

type
  E990 = record // ENREGISTREMENT DE TYPE FIN DE FICHIER
    Indicatif: string[19];
    TypeEnr: string[3];
    L1: string[200];
    L2: string[200];
    L3: string[142];
  end;

procedure TOF_PG_Envoi_Prudh.LanceEnvoi_Prudh(Sender: TObject);
var
  ListeJ: HTStrings;
  St: string;
  Chrono, i, reponse, rep: integer;
  okok: Boolean;
  QQ: TQuery;
begin
  Pan := TPageControl(GetControl('PANELPREP'));
  Tbsht := TTabSheet(GetControl('TBSHTTRACE'));
  Trace := TListBox(GetControl('LSTBXTRACE'));
  TraceErr := TListBox(GetControl('LSTBXERROR'));
  if (Trace = nil) or (TraceErr = nil) then
  begin
    PGIBox('La génération du fichier prud''hom ne peut pas être lancée', 'Car les composants trace ne sont pas disponibles');
    exit;
  end;
  if GetControlText('NOMFIC') = '' then
  begin
    PgiBox('Vous devez indiquer un nom de fichier', Ecran.Caption);
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
  //            V_PGI_Env.PathDat       V_PGI.DatPath
  if support <> 'DTK' then
  begin
    FileN := '$DAT' + '\' + GetControlText('NOMFIC');
    FileN := ChangeStdDatPath(FileN);
  end
  else FileN := 'A:\' + GetControlText('NOMFIC');

  //PT1
  if support = 'TRA' then
    FileN := VH_Paie.PGCheminEagl + '\' + GetControlText('NOMFIC');
  //FIN PT1
  if FileExists(FileN) then
  begin
    reponse := HShowMessage('5;;Voulez-vous supprimer votre fichier prud''hom ' + FileN + ';Q;YN;Y;N', '', '');
    if reponse = 6 then DeleteFile(PChar(FileN))
    else exit;
  end;
  AssignFile(FEcrt, FileN);
  {$I-}ReWrite(FEcrt);
  {$I+}if IoResult <> 0 then
  begin
    PGIBox('Fichier inaccessible : ' + FileN, 'Abandon du traitement');
    Exit;
  end;
  QQ := OpenSql('SELECT MAX(PES_NUMEROENVOI) FROM ENVOISOCIAL', TRUE);
  if QQ <> nil then NumEnvoi := QQ.Fields[0].AsInteger + 1
  else NumEnvoi := 1;
  Ferme(QQ);
  Trace.Items.Add('Traitement de l''émetteur');
  okok := PrudhEmetteur;
  Trace.Items.Add('Fin traitement de l''émetteur');
  if not okok then
  begin
    AbandonTraitementPrudh;
    exit;
  end;
  if ((Grille.nbSelected) > 0) and (not Grille.AllSelected) and okok then
  begin
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', Grille.NbSelected, FALSE, TRUE);

    InitMove(Grille.NbSelected, '');
    for i := 0 to Grille.NbSelected - 1 do
    begin
      Grille.GotoLeBOOKMARK(i);
      MoveCur(False);
      {$IFDEF EAGLCLIENT}
      TFmul(Ecran).Q.TQ.Seek(TFmul(Ecran).FListe.Row - 1);
      {$ENDIF}
      okok := TraiteFicPrudh;
      if not okok then
      begin
        AbandonTraitementPrudh;
        break;
      end;
      st := QMul.findfield('PES_LIBELLE').asString;
      MoveCurProgressForm(St); // PORTAGECWAS
    end;
    FiniMove;
    FiniMoveProgressForm(); // PORTAGECWAS
  end;
  if (Grille.AllSelected = TRUE) and okok then
  begin
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', QMul.RecordCount, FALSE, TRUE);

    InitMove(QMul.RecordCount, '');
    QMul.First;
    while not QMul.EOF do
    begin
      MoveCur(False);
      okok := TraiteFicPrudh;
      if not okok then
      begin
        AbandonTraitementPrudh;
        break;
      end;
      st := QMul.findfield('PES_LIBELLE').asString;
      MoveCurProgressForm(St); // PORTAGECWAS
      QMul.Next;
    end;
    FiniMove;
    FiniMoveProgressForm(); // PORTAGECWAS
  end;

  if okok then
  begin
    okok := EcritFinFich;
    if okok then Trace.Items.Add('Ecriture Fin de fichier terminée');
  end;
  closeFile(FEcrt);
  // Traitement de mise à jour des enregistrements de la table envoisocial
  if okok then
  begin // B1
    Trace.Items.Add('Mise à jour de la liste des fichiers envoyés');
    if ((Grille.nbSelected) > 0) and (not Grille.AllSelected) then
    begin // B2
      for i := 0 to Grille.NbSelected - 1 do
      begin // B3
        Grille.GotoLeBOOKMARK(i);
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(TFmul(Ecran).FListe.Row - 1);
        {$ENDIF}
        Chrono := QMul.findfield('PES_CHRONOMESS').asInteger;
        ReecritChronoPrudh(Chrono);
      end; // B3
      Grille.ClearSelected;
    end; // B2
    if (Grille.AllSelected = TRUE) then
    begin // B4
      QMul.First;
      while not QMul.EOF do
      begin // B5
        Chrono := QMul.findfield('PES_CHRONOMESS').asInteger;
        ReecritChronoPrudh(Chrono);
        QMul.Next;
      end; // B5
      Grille.AllSelected := False;
    end; // B4
    Trace.Items.Add('Fin de mise à jour de la liste des fichiers envoyés');
  end // B1
  else DeleteFile(PChar(FileN));

  if TraceErr.items.Count >= 1 then Trace.Items.Add('Fin du traitement, consultez vos anomalies')
  else Trace.Items.Add('Fin de traitement');
  Pan.ActivePage := Tbsht;
  if TraceErr.Items.Count > 0 then
  begin
    //Génération d'un fichier de log
    {$IFNDEF EAGLCLIENT}
    if MessageDlg('Voulez-vous générez le fichier ErreurPRUDH.log sous le répertoire ' + V_PGI.DatPath, mtConfirmation, [mbYes, mbNo], 0) = mrYes then //PT2
    {$ELSE}
    if MessageDlg('Voulez-vous générez le fichier ErreurPRUDH.log sous le répertoire ' + VH_PAIE.PGCheminEagl, mtConfirmation, [mbYes, mbNo], 0) = mrYes then //PT2
    {$ENDIF}
    begin
    {$IFNDEF EAGLCLIENT}
      if V_PGI.DatPath <> '' then FileZ := V_PGI.DatPath + '\ErreurPRUDH.log' //PT2
      else FileZ := 'C:\ErreurPRUDH.log';
     {$ELSE}
      if VH_PAIE.PGCheminEagl <> '' then FileZ := VH_PAIE.PGCheminEagl + '\ErreurPRUDH.log' //PT2
      else FileZ := 'C:\ErreurPRUDH.log';
     {$ENDIF}
      if SupprimeFichier(FileZ) = False then exit;
      AssignFile(FZ, FileZ);
      {$I-}ReWrite(FZ);
      {$I+}
      if IoResult <> 0 then
      begin
        PGIBox('Fichier inaccessible : ' + FileZ, 'Abandon du traitement');
        Exit;
      end;
      writeln(FZ, 'Création fichier Prud''hom : Gestion des messages d''erreur.');
      for i := 0 to TraceErr.Items.Count - 1 do
      begin
        St := TraceErr.Items.Strings[i];
        writeln(FZ, St);
      end;
      CloseFile(FZ);
      PGIInfo('La génération du fichier d''erreurs est terminée', Ecran.Caption);
    end;
  end;
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
      ListeJ.Add('Veuillez trouver ci-joint notre déclaration prud''homale');
      ListeJ.Add('Cordialement');
      SendMail('Déclaration prud''homale', '', '', ListeJ, FileN, FALSE);
      ListeJ.Clear;
      ListeJ.Free;
    end;
  end;
  {PT1
  PGIBox ('Fin de la procédure', Ecran.Caption);
  }
  if support <> 'TEL' then
    PGIBox('Vous devez émettre le fichier ' + FileN, Ecran.Caption)
  else
    PGIBox('Fin de la procédure', Ecran.Caption);
  //FIN PT1
  Rep := PGIAsk('Voulez vous éditer le bordereau d''accompagnement', Ecran.caption);
  if rep = mrYes then
  begin
    PreparBorderAcc;
    // LanceEtat;
  end;
end;

procedure TOF_PG_Envoi_Prudh.ImprimeClick(Sender: TObject);
var
  MPages: tpagecontrol;
begin
  {$IFNDEF EAGLCLIENT}
  MPages := TPageControl(getcontrol('PANELPREP'));
  if MPages <> nil then
    PrintPageDeGarde(MPages, TRUE, TRUE, FALSE, Ecran.Caption, 0);
  {$ENDIF}
end;

procedure TOF_PG_Envoi_Prudh.OnArgument(Arguments: string);
var
  F: TFVierge;
  st, S: string;
  BImprime: ttoolbarbutton97;
begin
  inherited;
  st := Trim(Arguments);
  S := ReadTokenSt(st); // Type de traitement prud'hom
  SetControlText('LBLTYPE', GetControlText('LBLTYPE') + ' ' + RechDom('PGENVOISOCIAL', S, FALSE));
  S := ReadTokenSt(st); // Millesime
  SetControlText('LBLMILLESIME', GetControlText('LBLMILLESIME') + ' ' + RechDom('PGANNEE', S, FALSE));
  S := ReadTokenSt(st); // Périodicité non traité specif DADS-U
  S := ReadTokenSt(st); // Destinataire non traité
  Emet := ReadTokenSt(st); // Emetteur
  SetControlText('LBLEMETTEUR', GetControlText('LBLEMETTEUR') + ' ' + Copy(RechDom('PGEMETTEURSOC', Emet, FALSE), 1, 25));
  S := ReadTokenSt(st); // Support
  Support := S;
  SetControlText('LBLSUPPORT', GetControlText('LBLSUPPORT') + ' ' + RechDom('PGSUPPORTEDI', S, FALSE));
  Reel := ReadTokenSt(st); // Envoi Reel
  if Reel = 'X' then SetControlProperty('LBLREEL', 'Checked', TRUE)
  else SetControlProperty('LBLREEL', 'Checked', FALSE);
  BtnLance := TToolbarButton97(GetControl('BLANCE'));
  if BtnLance <> nil then BtnLance.OnClick := LanceEnvoi_Prudh;
  BImprime := ttoolbarbutton97(getcontrol('BIMPRIMER'));
  if Bimprime <> nil then
    Bimprime.Onclick := ImprimeClick;
  if not (Ecran is TFVierge) then exit;
  F := TFVierge(Ecran);
  if F <> nil then QMUL := F.FMULQ;
end;
// Fonction de prise en compte du fichier code retour pour savoir
function TOF_PG_Envoi_Prudh.TraiteFicPrudh: Boolean;
var
  S: string;
begin
  result := FALSE;
  {$IFNDEF EAGLCLIENT}
  FileM := V_PGI.DatPath + '\' + QMul.findfield('PES_FICHIERRECU').AsString;
  {$ELSE}
  FileM := VH_PAIE.PGCheminEagl + '\' + QMul.findfield('PES_FICHIERRECU').AsString;
  {$ENDIF}
  Trace.Items.Add('Traitement des prud''hommes de l''entreprise ' + QMul.findfield('PES_LIBELLE').AsString);
  if not FileExists(FileM) then
  begin
    PgiBox('Fichier prud''hom inexistant pour l''entreprise ' + FileM,
      Ecran.caption);
    TraceErr.Items.Add('Fichier prud''hom inexistant pour l''entreprise ' + QMul.findfield('PES_FICHIERRECU').AsString);
    exit;
  end;
  AssignFile(FLect, FileM);
  {$I-}Reset(FLect);
  {$I+}if IoResult <> 0 then
  begin
    PGIBox('Fichier inaccessible : ' + FileM, 'Abandon du traitement');
    TraceErr.Items.Add('Fichier inaccessible : ' + FileM);
    Exit;
  end;

  while not eof(FLect) do
  begin
    {$I-}Readln(FLect, S);
    {$I+}if IoResult <> 0 then
    begin
      PGIBox('Erreur de lecture du fichier : ' + FileM, 'Abandon du traitement');
      TraceErr.Items.Add('Erreur de lecture du fichier : ' + FileM);
      closeFile(FLect);
      Exit;
    end;
    {$I-}Writeln(FEcrt, S);
    {$I+}if IoResult <> 0 then
    begin
      PGIBox('Erreur d''écriture du fichier : ' + FileN, 'Abandon du traitement');
      TraceErr.Items.Add('Erreur d''écriture du fichier : ' + FileN);
      closeFile(FLect);
      Exit;
    end;
  end;
  closeFile(FLect);
  Trace.Items.Add('Fin de traitement l''entreprise ' + QMul.findfield('PES_LIBELLE').AsString);
  result := TRUE;
end;
// Traitement des enregistrements emetteur
function TOF_PG_Envoi_Prudh.PrudhEmetteur: Boolean;
var
  st, S: string;
  Q1: Tquery;
  Ok: Boolean;
  E: E000;
begin
  result := FALSE;
  OK := False;
  St := 'SELECT * FROM EMETTEURSOCIAL WHERE PET_EMETTSOC="' + Emet + '"';
  Q1 := OpenSql(St, TRUE);
  if (not Q1.EOF) and (Q1 <> nil) then
  begin
    with E do
    begin
      indicatif := StringOfChar('0', sizeof(Indicatif));
      TypeEnr := '000';
      L1 := StringOfChar(' ', sizeof(L1));
      siret := Format_String(Q1.findField('PET_SIRET').AsString, 14);
      ok := TRUE;
      if not ControlSiret(siret) then
      begin
        PgiBox('Attention, le siret de l''émetteur n''est pas conforme', Ecran.caption);
        TraceErr.Items.Add('le siret de l''émetteur n''est pas conforme');
        Ok := FALSE;
      end;
      L2 := StringOfChar(' ', sizeof(L2));
      L3 := StringOfChar(' ', sizeof(L3));
      L4 := StringOfChar(' ', sizeof(L4));
    end;
  end;
  Ferme(Q1);

  if ok then
  begin
    S := E.indicatif + E.TypeEnr + E.L1 + E.Siret + E.L2 + E.L3 + E.L4;
    {$I-}Writeln(FEcrt, S);
    {$I+}if IoResult <> 0 then
    begin
      PGIBox('Erreur d''écriture début de fichier : ' + FileN, 'Abandon du traitement');
      TraceErr.Items.Add('Erreur d''écriture début de fichier : ' + FileN);
      Exit;
    end;
    result := TRUE;
  end;

end;

procedure TOF_PG_Envoi_Prudh.AbandonTraitementPrudh;
begin
  TraceErr.items.add('Une erreur est survenue lors de la conceptiondu fichier');
  TraceErr.items.add('Le traitement est interrompu');
  Trace.items.add('Le traitement est abandonné, vérifiez vos erreurs !');
end;
procedure TOF_PG_Envoi_Prudh.ReecritChronoPrudh(Chrono: Integer);
var
  st: string;
begin
  st := 'UPDATE ENVOISOCIAL SET PES_ENVOIREEL="' + Reel + '", PES_NUMEROENVOI=' +
    IntToStr(NumEnvoi) + ', PES_FICHIEREMIS="' + GetControlText('NOMFIC') + '", PES_EMETSOC="' +
    Emet + '", PES_SUPPORTEMIS="' + Support + '", PES_STATUTENVOI="TRA", PES_PREPARELE="' +
    UsDateTime(Date) + '", PES_ENVOYEPAR="' + V_PGI.User + '" WHERE PES_CHRONOMESS=' +
    IntToStr(Chrono) + ''; // DB2
  ExecuteSql(st);
end;

function TOF_PG_Envoi_Prudh.EcritFinFich: Boolean;
var
  S: string;
  E: E990;
begin
  result := FALSE;
  with E do
  begin
    indicatif := StringOfChar('9', sizeof(Indicatif));
    TypeEnr := '990';
    L1 := StringOfChar(' ', sizeof(L1));
    L2 := StringOfChar(' ', sizeof(L2));
    L3 := StringOfChar(' ', sizeof(L3));
  end;

  S := E.indicatif + E.TypeEnr + E.L1 + E.L2 + E.L3;
  {$I-}Writeln(FEcrt, S);
  {$I+}if IoResult <> 0 then
  begin
    PGIBox('Erreur d''écriture fin de fichier : ' + FileN, 'Abandon du traitement');
    TraceErr.Items.Add('Erreur d''écriture fin de fichier : ' + FileN);
    Exit;
  end;
  result := TRUE;
end;
// Fonction d'alimentation de la table du bordereau des prud'hom
procedure TOF_PG_Envoi_Prudh.PreparBorderAcc;
begin
  if support <> 'DTK' then
  begin
    FileN := '$DAT' + '\' + GetControlText('NOMFIC');
    FileN := ChangeStdDatPath(FileN);
  end
  else FileN := 'A:\' + GetControlText('NOMFIC');
  if not FileExists(FileN) then
  begin
    PgiBox('Le fichier ' + GetControlText('NOMFIC') + ' est inconnu ?', Ecran.caption);
    exit;
  end;
  EditBorderAcc(FileN);
end;

procedure EditBorderAcc(NomFic: string);
var
  TPrud, T1: TOB;
  NumOrdre, i: Integer;
  S: string;
  FText: TextFile; // Canaux des fichiers
begin
  NumOrdre := 1;
  Tprud := TOB.create('mon bordereau', nil, -1);
  AssignFile(FText, NomFic);
  {$I-}Reset(FText);
  {$I+}if IoResult <> 0 then
  begin
    PGIBox('Fichier inaccessible : ' + NomFic, 'Abandon du traitement');
    Exit;
  end;
  i := FileSize(FText) + 1; // calcul du nbre d'enregistrements,de rang
  InitMoveProgressForm(nil, 'Préparation du bordereau', 'Veuillez patienter SVP ...', i, FALSE, TRUE);

  while not eof(FText) do
  begin
    {$I-}Readln(FText, S);
    {$I+}if IoResult <> 0 then
    begin
      PGIBox('Erreur de lecture du fichier : ' + NomFic, 'Abandon du traitement');
      Exit;
    end;
    // On remplit la tob pour mettre à jour la table afin de préparer l'édition du bordereau
    T1 := TOB.create('PRUDBORD', TPrud, -1);
    if T1 <> nil then
    begin
      T1.putvalue('PPD_NUMORDRE', NumOrdre);
      T1.putvalue('PPD_SIRET', Copy(S, 1, 14));
      T1.putvalue('PPD_TYPEENR', Copy(S, 20, 3));
      T1.putvalue('PPD_PRUDHLIGNE', S);
    end;
    MoveCurProgressForm(IntToStr(NumOrdre)); // PORTAGECWAS
    NumOrdre := NumOrdre + 1;
  end;
  closeFile(FText);
  try
    BeginTrans;
    MoveCurProgressForm(IntToStr(NumOrdre + 1)); // PORTAGECWAS
    MoveCurProgressForm('Sauvegarde en cours du bordereau'); // PORTAGECWAS
    ExecuteSql('DELETE FROM PRUDBORD'); // On detruit le contenu de la table
    TPrud.SetAllModifie(TRUE);
    TPrud.InsertOrUpdateDB(TRUE);
    CommitTrans;
    LanceEtat('E', 'PDA', 'PRU', True, False, False, nil, '', '', False);
    i := PGIAsk('Voulez vous imprimer les lettres d''information aux maires', 'Déclarations prud''homales');
    if i = mrYes then
      LanceEtat('E', 'PDA', 'PLC', True, False, False, nil, '', '', False);
  except
    Rollback;
  end;
  FiniMoveProgressForm(); // PORTAGECWAS
  TPrud.Free;
end;

initialization
  registerclasses([TOF_PG_Envoi_Prudh]);
end.


{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 18/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGPREP_CONGESCPEC ()
Mots clefs ... : TOF;PGPREP_CONGESCPEC
*****************************************************************
PT1 16/01/2006 JL V_650 FQ 12736 Edition du n° compte entreprise
}
unit UTofPGPrep_CongesSpec;

interface

uses Windows, StdCtrls, Controls, Classes, sysutils, ComCtrls, HTB97,
  ed_tools, Mailol,
  {$IFDEF EAGLCLIENT}
  UtileAGL, eFiche, eMul,UTob,HQry,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Qre, HDB, EdtREtat,
  {$ENDIF}
  HCtrls, HEnt1, EntPaie, HMsgBox, UTOF, Vierge, HStatus,ParamSoc;


type
  TOF_PGPREP_CONGESCPEC = class(TOF)
    procedure OnArgument(Arguments: string); override;
  private
    Trace, TraceErr: TListBox;
    QMul: TQUERY;
    Pan: TPageControl;
    Tbsht: TTabSheet;
    FileN, FileM: string;
    Reel, Emet, Monnaie, Support: string;
    FLect: TextFile;
    BtnLance, BImprime: TToolBarButton97;
    NumEnvoi: Integer;
    NbCertificats: Integer;
    BaseConges: Double;
    DateDebutConges, DateFinConges: TDateTime;
    procedure LancePrepCongSpec(Sender: TObject);
    procedure ImprimeClick(Sender: TObject);
    procedure AbandonTraitement;
    function TraiteFichier: Boolean;
    function ReecritChrono(Chrono: Integer): Boolean;
  end;

implementation
uses
  UTofPG_MulEnvoiSocial;

procedure TOF_PGPREP_CONGESCPEC.OnArgument(Arguments: string);
var
  S, St,EtabPrinc,NumEntre: string;
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
  SetControlText('LBLSUPPORT', GetControlText('LBLSUPPORT') + ' ' + RechDom('PGSUPPORTEDI', S, FALSE));
  Reel := ReadTokenSt(st);
  if Reel = 'X' then SetControlProperty('LBLREEL', 'Checked', TRUE)
  else SetControlProperty('LBLREEL', 'Checked', FALSE);
  S := ReadTokenSt(st); // Monnaie
  Monnaie := S;
  SetControlText('LBLMONNAIE', GetControlText('LBLMONNAIE') + ' ' + RechDom('PGMONNAIE', S, FALSE));
  BtnLance := TToolbarButton97(GetControl('BLANCE'));
  if BtnLance <> nil then
    BtnLance.OnClick := LancePrepCongSpec;
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
  SetControlText('RAISONSOC', '');
  SetControlText('ADRESSE1', '');
  SetControlText('ADRESSE2', '');
  Q := OpenSQL('SELECT PET_RAISONSOC,PET_NUMEROVOIE,PET_BISTER,PET_NOMVOIE,PET_BURDISTRIB,PET_CODEPOSTAL FROM EMETTEURSOCIAL WHERE PET_EMETTSOC="' + Emet + '"', True);
  if not Q.Eof then
  begin
    SetControlText('RAISONSOC', Q.FindField('PET_RAISONSOC').AsString);
    SetControlText('ADRESSE1', Q.FindField('PET_NUMEROVOIE').AsString + ' ' + Q.FindField('PET_BISTER').AsString + ' ' + Q.FindField('PET_NOMVOIE').AsString);
    SetControlText('ADRESSE2', Q.FindField('PET_CODEPOSTAL').AsString + ' ' + Q.FindField('PET_BURDISTRIB').AsString);
  end;
  Ferme(Q);
  //DEBUT PT1
  EtabPrinc := GetParamSoc('SO_ETABLISDEFAUT');
  Q := OpenSQL('SELECT ETB_ISNUMCPAY FROM ETABCOMPL WHERE ETB_ETABLISSEMENT="'+EtabPrinc+'"',True);
  If Not Q.Eof Then NumEntre := Q.FindField('ETB_ISNUMCPAY').AsString;
  Ferme(Q);
  If NumEntre <> '' then
  begin
       SetControlText('NUMCOMPTEENTRE',Copy(NumEntre,1,8));
       SetControlText('CLECOMPTEENTRE',Copy(NumEntre,9,1));
  end;
  //FIN PT1
end;

procedure TOF_PGPREP_CONGESCPEC.LancePrepCongSpec(Sender: TObject);
var
  ListeJ: HTStrings;
  St: string;
  Chrono, i, reponse, rep: integer;
  okok: Boolean;
  QQ, Q: TQuery;
  StPages: string;
begin
  Pan := TPageControl(GetControl('PAGES'));
  Tbsht := TTabSheet(GetControl('TBSHTTRACE'));
  Trace := TListBox(GetControl('LSTBXTRACE'));
  TraceErr := TListBox(GetControl('LSTBXERROR'));
  DateDebutConges := IDate1900;
  DateFinConges := IDate1900;
  {$IFNDEF EAGLCLIENT}
  FileM := V_PGI.DatPath + '\' + QMul.findfield('PES_FICHIERRECU').AsString;
  {$ELSE}
  FileM := VH_Paie.PGCheminEagl + '\' + QMul.findfield('PES_FICHIERRECU').AsString;
  {$ENDIF}
  if (Trace = nil) or (TraceErr = nil) then
  begin
    PGIBox('La préparation des congés spectacles ne peut pas être lancée', 'Car les composants trace ne sont pas disponibles');
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
   if support <> 'DTK' then
   begin
   FileN := '$DAT' + '\' + GetControlText('NOMFIC') ;
   FileN := ChangeStdDatPath(FileN);
   end
  else FileN := 'A:\' + GetControlText('NOMFIC');
  if support = 'TRA' then
    FileN := VH_Paie.PGCheminEagl + '\' + GetControlText('NOMFIC');
  if FileExists(FileN) then
  begin
    reponse := HShowMessage('5;;Voulez-vous supprimer votre fichier congés spectacles ' + FileN + ';Q;YN;Y;N', '', '');
    if reponse = 6 then DeleteFile(PChar(FileN))
    else exit;
  end;
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
      Q := OpenSQL('SELECT PES_DATEDEBUT,PES_DATEFIN FROM ENVOISOCIAL WHERE PES_CHRONOMESS=' + IntToStr(Chrono) + '', True); // DB2
      if not Q.Eof then
      begin
        if (DateDebutConges = IDate1900) or (DateDebutConges > (Q.findfield('PES_DATEDEBUT').ASDateTime)) then
          DateDebutConges := Q.findfield('PES_DATEDEBUT').ASDateTime;
        if (DateFinConges = IDate1900) or (DateFinConges < (Q.findfield('PES_DATEFIN').ASDateTime)) then
          DateFinConges := Q.findfield('PES_DATEFIN').ASDateTime;
      end;
      Ferme(Q);
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
      Q := OpenSQL('SELECT PES_DATEDEBUT,PES_DATEFIN FROM ENVOISOCIAL WHERE PES_CHRONOMESS=' + IntToStr(Chrono) + '', True); // DB2
      if not Q.Eof then
      begin
        if (DateDebutConges = IDate1900) or (DateDebutConges > (Q.findfield('PES_DATEDEBUT').ASDateTime)) then
          DateDebutConges := Q.findfield('PES_DATEDEBUT').ASDateTime;
        if (DateFinConges = IDate1900) or (DateFinConges < (Q.findfield('PES_DATEFIN').ASDateTime)) then
          DateFinConges := Q.findfield('PES_DATEFIN').ASDateTime;
      end;
      Ferme(Q);
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
  Trace.Items.Add('Fin de mise à jour de la liste des fichiers envoyés');
  CopyFile(PChar(FileM), PChar(FileN), False);
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
      ListeJ.Add('Veuillez trouver ci-joint notre déclaration congés spectacles');
      ListeJ.Add('Cordialement');
      SendMail('Déclaration Congés spectacles', '', '', ListeJ, FileN, FALSE);
      ListeJ.Clear;
      ListeJ.Free;
    end;
  end;
  if support <> 'TEL' then
    PGIBox('Vous devez émettre le fichier ' + FileN, Ecran.Caption)
  else
    PGIBox('Fin de la procédure', Ecran.Caption);
  Rep := PGIAsk('Voulez-vous éditer la déclaration à joindre avec le fichier', Ecran.Caption);
  if rep = mrYes then
  begin
    SetControlText('BASECONGES', FloatToStr(BaseConges));
    SetControlText('NBCERTIFICATS', IntToStr(NbCertificats));
    SetControlText('DEBUTCONGES', DateToStr(DateDebutConges));
    SetControlText('FINCONGES', DateToStr(DateFinConges));
    SetControlText('DATEENVOI', DateToStr(date));
    St := 'SELECT * FROM ENVOISOCIAL WHERE PES_CHRONOMESS=' + IntToStr(Chrono) + ''; // DB2
    StPages := '';
    {$IFDEF EAGLCLIENT}
    StPages := AglGetCriteres(Pan, FALSE);
    {$ENDIF}
    LanceEtat('E', 'PCT', 'PDC', True, False, False, Pan, St, '', False, 0, StPages);
  end;
end;

procedure TOF_PGPREP_CONGESCPEC.ImprimeClick(Sender: TObject);
var
  MPages: tpagecontrol;
begin
  {$IFNDEF EAGLCLIENT}
  MPages := TPageControl(getcontrol('PAGES'));
  if MPages <> nil then
    PrintPageDeGarde(MPages, TRUE, TRUE, FALSE, Ecran.Caption, 0);
  {$ENDIF}
end;

procedure TOF_PGPREP_CONGESCPEC.AbandonTraitement;
begin
  TraceErr.items.add('Une erreur est survenue lors de la conception du fichier');
  TraceErr.items.add('Le traitement est interrompu');
  Trace.items.add('Le traitement est abandonné, vérifiez vos erreurs !');
end;

function TOF_PGPREP_CONGESCPEC.TraiteFichier: Boolean;
var
  S, St: string;
begin
  BaseConges := 0;
  NbCertificats := 0;
  result := FALSE;
  Trace.Items.Add('Traitement des Congés Spectacles de l''entreprise ' + QMul.findfield('PES_LIBELLE').AsString);
  if not FileExists(FileM) then
  begin
    PgiBox('Fichier Congés Spectacles inexistant pour l''entreprise ' + FileM, Ecran.caption);
    TraceErr.Items.Add('Fichier Congés Spectacles inexistant pour l''entreprise ' + QMul.findfield('PES_FICHIERRECU').AsString);
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
    st := Copy(S, 131, 9);
    if not IsNumeric(St) then
    begin
      PGIBox('Erreur de lecture du fichier : ' + FileM, 'Abandon du traitement');
      TraceErr.items.add('Le calcul de la base congés est impossible');
      closeFile(FLect);
      Exit;
    end;
    BaseConges := BaseConges + StrToFloat(St);
    NbCertificats := NbCertificats + 1;
  end;
  closeFile(FLect);
  Trace.Items.Add('Fin de traitement de l''entreprise ' + QMul.findfield('PES_LIBELLE').AsString);
  result := TRUE;
end;

function TOF_PGPREP_CONGESCPEC.ReecritChrono(Chrono: Integer): Boolean;
var
  st: string;
begin
  st := 'UPDATE ENVOISOCIAL SET PES_ENVOIREEL="' + Reel + '", PES_NUMEROENVOI=' +
    IntToStr(NumEnvoi) + ', PES_FICHIEREMIS="' + GetControlText('NOMFIC') + '", PES_EMETSOC="' +
    Emet + '", PES_SUPPORTEMIS="' + Support + '", PES_STATUTENVOI="TRA", PES_PREPARELE="' +
    UsDateTime(Date) + '", PES_ENVOYEPAR="' + V_PGI.User + '", PES_MTPAYER=' + FloatToStr(BaseConges) + ' WHERE PES_CHRONOMESS=' + // DB2
  IntToStr(Chrono) + ''; // DB2
  ExecuteSql(st);
  Result := True;
end;

initialization
  registerclasses([TOF_PGPREP_CONGESCPEC]);
end.


{***********UNITE*************************************************
Auteur  ...... : Julien D
Créé le ...... : 19/03/2002
Modifié le ... : 20/03/2002
Description .. : Source TOF de la FICHE : MBOEXPORTPIECE ()
Mots clefs ... : TOF;MBOEXPORTPIECE
*****************************************************************}
unit MBOEXPORTPIECE_TOF;

interface

uses UTOF, UTOB,
  {$IFDEF EAGLCLIENT}
  eMul, MaineAGL,
  {$ELSE}
  mul, Fe_Main, dbTables,
  {$ENDIF}
  HMsgBox, FileCtrl, sysutils, HEnt1, Controls, HStatus,
  StdCtrls, HCtrls, ComCtrls, forms, Classes, HQry, EntGC, Paramsoc, M3FP,
  EXPORTASCII_ENTETE_TOM, uToxFTP;

type FichierFournisseur = record
    Fournisseur: string;
    Emplacement: string;
    FichierASCII: TextFile;
  end;
type
  TOF_MBOEXPORTPIECE = class(TOF)
  private
    stWhere, Separateur, ListeDesChamps, PathFichier: string;
    DateFormat, BufferAlpha, BufferNumerique, LimiteArticle: string;
    BFormatDynamique: Boolean;
    StGCORLIHOST, StGCORLIUSERNAME, StGCORLIPASSWORD, StGCORLIREPFTP: string;
    TOBChampASCII: TOB;
    TabFichierASCII: array[0..2056] of FichierFournisseur;
    IdTab: integer;
    procedure BatchExport;
    procedure SetEXPORT;
    procedure SetAllEXPORT;
    function ChargeTOBChampASCII: boolean;
    function QuelFichierASCII(Q: Tquery): integer;
    procedure AjouteLignesDuQueryDansFichier(Q: TQuery);
    procedure FermeLesFichiers;
    procedure ChargeCheminExportTRF;
    procedure ParamFTP;
    procedure ExportDuFichierParFTP;
  public
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
  end;

procedure LanceExportPiece(NomFiche, NaturePiece, NatureExport: string);

implementation

procedure LanceExportPiece(NomFiche, NaturePiece, NatureExport: string);
begin
  AGLLanceFiche('MBO', NomFiche, 'GP_NATUREPIECEG=' + NaturePiece + ';NATUREEXPORT=' + NatureExport, '', '');
end;

procedure TOF_MBOEXPORTPIECE.BatchExport;
var i: integer;
begin
  IdTab := -1;
  with TFMul(Ecran) do
  begin
    if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
    begin
      PGIInfo('Veuillez sélectionner les pièces à exporter', Caption);
      exit;
    end;

    PathFichier := GetControlText('EMPLACEMENT');
    if (PathFichier = '') or not DirectoryExists(ExtractFileDir(PathFichier)) then
    begin
      PGIInfo('Vous devez sélectionner un emplacement de fichier.', Caption);
      exit;
    end;
    //SetParamSoc('SO_GENERATIONFICHIERASCII',PathFichier);

    if not ChargeTOBChampASCII then
    begin
      PGIError('Impossible de charger le paramétrage d''export.', Caption);
      exit;
    end;

    LimiteArticle := '';
    for i := 0 to TOBChampASCII.Detail.Count - 1 do
    begin
      if TOBChampASCII.Detail[i].GetValue('GEX_NOMCHAMP') = '$$_SUPPQTEFACTZERO' then
      begin
        LimiteArticle := ' AND GL_QTEFACT<>0';
        TOBChampASCII.Detail[i].Free;
        Break;
      end;
    end;

    SourisSablier;
    try
      if FListe.AllSelected then
      begin
        if PGIAsk('Voulez-vous exporter toutes les pièces ?', Caption) <> mrYes then exit;
        if Transactions(SetAllEXPORT, 2) <> oeOK then PGIError('Impossible d''exporter toutes les pièces', Caption);
        FListe.AllSelected := false;
      end
      else
      begin
        if PGIAsk('Voulez-vous exporter les pièces sélectionnées ?', Caption) <> mrYes then exit;
        InitMove(FListe.NbSelected, '');
        for i := 0 to FListe.NbSelected - 1 do
        begin
          FListe.GotoLeBookMark(i);
          {$IFDEF EAGLCLIENT}
          Q.TQ.Seek(FListe.Row - 1);
          {$ENDIF}
          if Transactions(SetEXPORT, 2) <> oeOK then PGIBox('Impossible d''exporter la pièce n° ' + Q.FindField('GP_NUMERO').AsString, Caption);
          MoveCur(False);
        end;
        FListe.ClearSelected;
        FiniMove;
      end;
      ChercheClick;
    finally
      FermeLesFichiers;

      //Si EXPORTFTP est coché : envoie le fichier par FTP
      if GetCheckBoxState('EXPORTFTP') = cbChecked then ExportDuFichierParFTP;

      if TOBChampASCII <> nil then
      begin
        TOBChampASCII.Free;
        TOBChampASCII := nil;
      end;
      bSelectAll.Down := False;
      SourisNormale;
    end;
  end;
end;

procedure TOF_MBOEXPORTPIECE.ExportDuFichierParFTP;
var xpFtp: TOXFTP;
  PortFTP, i: integer;
  LocalFileName, HostFileName: string;
begin
  if (StGCORLIHOST = '') or (StGCORLIUSERNAME = '') or (StGCORLIPASSWORD = '') then exit;
  PortFTP := 21;
  xpFtp := TOXFTP.Create(PortFtp);
  if xpFTP.Connect(StGCORLIHOST, StGCORLIUSERNAME, StGCORLIPASSWORD) then
  begin
    for i := 0 to IdTab do
    begin
      LocalFileName := TabFichierASCII[i].Emplacement;
      HostFileName := StGCORLIREPFTP + ExtractFileName(LocalFileName);
      xpFTP.PutFileAscii(LocalFileName, HostFileName);
    end;
    xpFTP.DisConnect;
  end;
  xpFTP.Free;
end;

function TOF_MBOEXPORTPIECE.ChargeTOBChampASCII: boolean;
var Q, QSep: TQuery;
  CodeExport, NatureExport: string;
begin
  Result := False;

  CodeExport := GetControlText('LISTEEXPORT');
  NatureExport := GetControlText('NATUREEXPORT');
  //Charge les champs du fichier ASCII
  TOBChampASCII := Tob.Create('Table EXPORTASCII', nil, -1);
  Q := OpenSQL('SELECT * FROM EXPORTASCII WHERE GEX_NATURE="' + NatureExport + '" AND GEX_CODEEX="' + CodeExport + '" ORDER BY GEX_POSITION', True);
  if not Q.Eof then TOBChampASCII.LoadDetailDB('EXPORTASCII', '', '', Q, False)
  else
  begin
    Ferme(Q);
    TOBChampASCII.Free;
    exit;
  end;
  Ferme(Q);

  ASCIIAjouteChampGaucheEtDate(TOBChampASCII);

  ListeDesChamps := ASCIIConstruireListeDeChamps(TOBChampASCII);

  // Ajoute le champ GA_FOURNFRINC pour pouvoir séparer les lignes par fournisseur de l'article
  if GetCheckBoxState('FICHIERPARFOURNPRINC') = cbChecked then
    if Pos('GA_FOURNPRINC', ListeDesChamps) = 0 then ListeDesChamps := ListeDesChamps + ',GA_FOURNPRINC';

  QSep := OpenSQL('SELECT GAS_SEPARATEUR,GAS_DATEFORMAT,GAS_ALPHA,GAS_NUMERIQUE,GAS_FORMATFIXE ' +
    'FROM EXPORTASCII_ENTETE WHERE GAS_NATURE="' + NatureExport + '" AND GAS_CODEEX="' + CodeExport + '"', True);
  Separateur := '';
  DateFormat := '';
  BufferAlpha := ' ';
  BufferNumerique := ' ';
  if not QSep.Eof then
  begin
    Separateur := QSep.FindField('GAS_SEPARATEUR').AsString;
    DateFormat := QSep.FindField('GAS_DATEFORMAT').AsString;
    BufferAlpha := QSep.FindField('GAS_ALPHA').AsString;
    BufferNumerique := QSep.FindField('GAS_NUMERIQUE').AsString;
    BFormatDynamique := QSep.FindField('GAS_FORMATFIXE').AsString = '-';
  end;
  Ferme(QSep);

  Result := True;
end;

// Procédure appelée par une transaction pour exporter les lignes de la pièce

procedure TOF_MBOEXPORTPIECE.SetEXPORT;
var Nature, DateP, Souche, Nb, Indice, DateExport: string;
  StWherePiece, StSQL, StJointure: string;
  Q: TQuery;
begin
  DateExport := USDateTime(NowH);
  StJointure := CreationJointure('GL', ListeDesChamps);

  with TFMul(Ecran) do
  begin
    Nature := Q.FindField('GP_NATUREPIECEG').AsString;
    DateP := USDateTime(Q.FindField('GP_DATEPIECE').AsDateTime);
    Souche := Q.FindField('GP_SOUCHE').AsString;
    Nb := Q.FindField('GP_NUMERO').AsString;
    Indice := Q.FindField('GP_INDICEG').AsString;
  end;
  StWherePiece := 'WHERE GP_NATUREPIECEG="' + Nature + '" AND GP_DATEPIECE="' + DateP + '" AND ' +
    'GP_SOUCHE="' + Souche + '" AND GP_NUMERO=' + Nb + ' AND GP_INDICEG=' + Indice;
  StSQL := 'SELECT ' + ListeDesChamps + ' FROM PIECE ' +
    'LEFT JOIN LIGNE ON GL_NATUREPIECEG=GP_NATUREPIECEG ' +
    'AND GL_SOUCHE=GP_SOUCHE AND GL_NUMERO=GP_NUMERO ' +
    'AND GL_INDICEG=GP_INDICEG ' +
    StJointure + StWherePiece + ' AND GL_TYPELIGNE="ART"' + LimiteArticle;
  Q := OpenSQL(StSQL, True);
  if not Q.Eof then
  begin
    AjouteLignesDuQueryDansFichier(Q);
    ExecuteSQL('UPDATE PIECE SET GP_ETATEXPORT="EXP",GP_DATEEXPORT="' + DateExport + '" ' + StWherePiece);
  end;
  Ferme(Q);
end;

// Procédure appelée par une transaction pour exporter les lignes de toutes
// les pièces sélectionnées

procedure TOF_MBOEXPORTPIECE.SetAllEXPORT;
var DateExport, StJointure, StSQL: string;
  Q: TQuery;
begin
  DateExport := USDateTime(NowH);
  StJointure := CreationJointure('GL', ListeDesChamps);

  StSQL := 'SELECT ' + ListeDesChamps + ' FROM PIECE ' +
    'LEFT JOIN LIGNE ON GL_NATUREPIECEG=GP_NATUREPIECEG ' +
    'AND GL_SOUCHE=GP_SOUCHE AND GL_NUMERO=GP_NUMERO ' +
    'AND GL_INDICEG=GP_INDICEG ' +
    StJointure +
    'WHERE ' + Copy(StWhere, 7, length(StWhere) - 7) +
    ' AND GL_TYPELIGNE="ART"' + LimiteArticle;

  Q := OpenSQL(StSQL, True);
  if not Q.Eof then
  begin
    AjouteLignesDuQueryDansFichier(Q);
    ExecuteSQL('UPDATE PIECE SET GP_ETATEXPORT="EXP",GP_DATEEXPORT="' + DateExport + '" ' + StWhere);
  end;
  Ferme(Q);
end;

function TOF_MBOEXPORTPIECE.QuelFichierASCII(Q: Tquery): integer;
var FournPrinc, StChemin: string;
  i: integer;
  Trouver: boolean;
begin
  Result := -1;
  FournPrinc := '';
  Trouver := False;
  if GetCheckBoxState('FICHIERPARFOURNPRINC') = cbChecked then
    FournPrinc := Q.Findfield('GA_FOURNPRINC').AsString;
  for i := 0 to IdTab do
  begin
    if TabFichierASCII[i].Fournisseur = FournPrinc then
    begin
      Result := i;
      Trouver := True;
      Break;
    end;
  end;
  if not Trouver then
  begin
    IdTab := IdTab + 1;
    if IdTab > 2056 then
    begin
      PGIError('Le nombre de fournisseur exède 2056.', Ecran.Caption);
      exit;
    end;
    TabFichierASCII[IdTab].Fournisseur := FournPrinc;
    if FournPrinc <> '' then FournPrinc := '_' + FournPrinc;
    StChemin := PathFichier;
    Insert(FournPrinc, StChemin, Length(StChemin) - 3);
    AssignFile(TabFichierASCII[IdTab].FichierASCII, StChemin);
    TabFichierASCII[IdTab].Emplacement := StChemin;
    if (GetCheckBoxState('COMPLETER') = cbChecked) and FileExists(StChemin) then
      Append(TabFichierASCII[IdTab].FichierASCII)
    else
      Rewrite(TabFichierASCII[IdTab].FichierASCII);
    Result := IdTab;
  end;
end;

procedure TOF_MBOEXPORTPIECE.FermeLesFichiers;
var i: integer;
begin
  for i := 0 to IdTab do
    CloseFile(TabFichierASCII[i].FichierASCII);
end;

procedure TOF_MBOEXPORTPIECE.AjouteLignesDuQueryDansFichier(Q: TQuery);
var UneLigne: string;
  id: integer;
begin
  Q.First;
  while not Q.Eof do
  begin
    UneLigne := ASCIIConstruireUneLigne(Q, TOBChampASCII, GetControlText('LISTEEXPORT'),
      Separateur, DateFormat, BufferAlpha, BufferNumerique, BFormatDynamique);
    id := QuelFichierASCII(Q);
    if id = -1 then break;
    WriteLn(TabFichierASCII[id].FichierASCII, UneLigne);
    Q.Next;
  end;
end;

procedure TOF_MBOEXPORTPIECE.OnLoad;
begin
  inherited;
  stWhere := RecupWhereCritere(TPageControl(TFMul(Ecran).Pages));
end;

procedure TOF_MBOEXPORTPIECE.OnArgument(S: string);
begin
  inherited;
  // paramétrage du libellé etablissement quand on est en multi-dépôts
  if VH_GC.GCMultiDepots then
  begin
    SetControlText('TGP_DEPOT', TraduireMemoire('Dépôt émetteur'));
    if TFMul(Ecran).Name = 'EXPORTPIECE_MUL' then SetControlText('TGP_DEPOTDEST', TraduireMemoire('Dépôt récepteur'));
  end
  else
  begin
    SetControlText('TGP_DEPOT', TraduireMemoire('Etabliss. émetteur'));
    if TFMul(Ecran).Name = 'EXPORTPIECE_MUL' then SetControlText('TGP_DEPOT', TraduireMemoire('Etabliss. récepteur'));
  end;
  THValComboBox(GetControl('GP_ETATEXPORT')).Value := 'ATT';
  StGCORLIHOST := GetParamSoc('SO_GCORLIHOST');
  StGCORLIUSERNAME := GetParamSoc('SO_GCORLIUSERNAME');
  StGCORLIPASSWORD := GetParamSoc('SO_GCORLIPASSWORD');
  StGCORLIREPFTP := GetParamSoc('SO_GCORLIREPFTP');
end;

procedure TOF_MBOEXPORTPIECE.ChargeCheminExportTRF;
var Q: TQuery;
begin
  Q := OpenSQL('SELECT GAS_FICHIEREXPORT FROM EXPORTASCII_ENTETE WHERE GAS_NATURE="' + GetControlText('NATUREEXPORT') + '" AND GAS_CODEEX="' +
    GetControlText('LISTEEXPORT') + '"', True);
  if not Q.EOF then SetControlText('EMPLACEMENT', Q.FindField('GAS_FICHIEREXPORT').AsString);
  Ferme(Q);
end;

procedure TOF_MBOEXPORTPIECE.ParamFTP;
var LesParametres, UnParam, StParam, ValParam: string;
  x: integer;
begin
  LesParametres := AGLLanceFiche('MBO', 'ORLIFTP', '', '', '');
  while LesParametres <> '' do
  begin
    UnParam := READTOKENST(LesParametres);
    if UnParam <> '' then
    begin
      x := pos('=', UnParam);
      if x <> 0 then
      begin
        StParam := copy(UnParam, 1, x - 1);
        ValParam := copy(UnParam, x + 1, length(UnParam));
        if StParam = 'SO_GCORLIHOST' then StGCORLIHOST := ValParam
        else
          if StParam = 'SO_GCORLIUSERNAME' then StGCORLIUSERNAME := ValParam
        else
          if StParam = 'SO_GCORLIPASSWORD' then StGCORLIPASSWORD := ValParam
        else
          if StParam = 'SO_GCORLIREPFTP' then StGCORLIREPFTP := ValParam
          ;
      end;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////

// procedure appellée par le bouton BOuvrir

procedure AGLBatchEXPORT(Parms: array of Variant; Nb: Integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFmul) then TOTOF := TFMul(F).LaTOF else exit;
  if (TOTOF is TOF_MBOEXPORTPIECE) then TOF_MBOEXPORTPIECE(TOTOF).BatchExport;
end;

procedure AGLChargeCheminExportTRF(Parms: array of Variant; Nb: Integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFmul) then TOTOF := TFMul(F).LaTOF else exit;
  if (TOTOF is TOF_MBOEXPORTPIECE) then TOF_MBOEXPORTPIECE(TOTOF).ChargeCheminExportTRF;
end;

procedure AGLParamFTP(Parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFmul) then TOTOF := TFMul(F).LaTOF else exit;
  if (TOTOF is TOF_MBOEXPORTPIECE) then TOF_MBOEXPORTPIECE(TOTOF).ParamFTP;
end;

initialization
  registerclasses([TOF_MBOEXPORTPIECE]);
  RegisterAGLProc('BatchEXPORT', True, 0, AGLBatchEXPORT);
  RegisterAGLProc('ChargeCheminExportTRF', True, 0, AGLChargeCheminExportTRF);
  RegisterAglProc('ParamFTP', True, 0, AGLParamFTP);
end.

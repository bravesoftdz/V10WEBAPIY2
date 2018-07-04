{***********UNITE*************************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 22/03/2001
Modifié le ... : 23/11/2001
Description .. : Source TOF de la TABLE : GENETRANSF ()
Mots clefs ... : TOF;GENETRANSF
*****************************************************************}
unit UTOFMBOGENETRANSF;

interface

uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
  HCtrls, HEnt1, HMsgBox, UTOF, UTOB, Hqry, FactUtil, Paramsoc,
  {$IFDEF EAGLCLIENT}
  UtileAGL, Maineagl,
  {$ELSE}
  FE_Main, dbTables, db, EdtDoc, EdtREtat,
  {$ENDIF}
  Vierge, ent1, SaisUtil, FactComm, EntGC, FactCalc, HStatus, M3FP, Graphics,
  Windows, FileCtrl, EXPORTASCII_ENTETE_TOM, uRecupSQLModele, uToxFTP, StockUtil,
  FactTOB, FactPiece, FactAdresse, FactArticle;

type
  TOF_MBOGENETRANSF = class(TOF)
  private
    CodePtrf, CodeTiers, CodeDepotEmmetteur, NaturePiece: string;
    CleDoc: R_CleDoc;
    DEV: RDEVISE;
    QualifMvt: string; // qualifiant mouvement de la nature de pièce
    Numligne: integer;
    CanClose, EnHT, BImpression, BGenerationCF: boolean;
    TOBToutesLesPieces: TOB;
    Tob_Tiers: TOB; // TOB du tiers
    Tob_Article: TOB; // TOB des articles de la pièce
    Tob_PiedBase: TOB; // TOB PiedBase
    Tob_PiedEche: TOB; // TOB PiedEche
    Tob_Catalogue: TOB; // TOB Catalogues
    Tob_Noment: TOB; // TOB Nomenclatures
    Tob_Adresses: TOB; // TOB adresses
    StGCORLIHOST, StGCORLIUSERNAME, StGCORLIPASSWORD, StGCORLIREPFTP: string;
    TransfertPGI : boolean;
    procedure GenerationFichierASCII;
    procedure GenerationPGI;
    procedure ConstruireLesFondations(TOBFondation: TOB);
    function ConstruireTOBTrfDepot(TOBTRF: TOB; QDepot: TQuery; DepotEmetteur: string): TOB;
    function ConstruireTOBCFFournPrinc(TOBCF: TOB; QFourn: TQuery; DepotEmetteur: string): TOB;
    procedure ConstruireTOBTrfArt(TOBTrfDepot: TOB; QArt: TQuery);
    procedure ConstruireTousLesDocuments(TOBTRF: TOB);
    function ConstruirePiece(TOBTrfDepot: TOB): TOB;
    function ChargeTOBClient(TOBCF: TOB): Boolean;
    function ConstruireLignes(TOBPiece, TOBTrfDepot: TOB): boolean;
    function CreatLignePiece(TOBPiece, TobTrfArt: TOB): boolean;
    procedure BuildTob_ArticleAndDispo(TOBDepot: TOB);
    function FindStInTAB(St: string; Tab: array of string): boolean;
    procedure MAJPrixLignesGEN(TOBP: TOB);
    function CreationTRV(Depot: string): boolean;
    procedure ConstruireTRVouTRE(TOBP: TOB);
    function ValideLesPieces(TOBTrfDepot: TOB): boolean;
    procedure MarquePROPTRANSFLIG(TOBTrfDepot: TOB);
    procedure ImpressionDesEtats(BApercu, BDuplicata: boolean; StModele: string);
    procedure Affichage(Etape: integer; StDetail: string = '');
    procedure MAJBlocNote(Etape: integer; NbTransfert: integer = 0);
    procedure AfterGeneration;
    procedure SetLastError(Num: integer; ou: string);
    procedure ChargeCheminExport;
    procedure ParametrageFTP;
    procedure ExportDuFichierParFTP;
    procedure EditionClick;
    procedure ChoixTransfert(TypeGeneration: string);
  public
    procedure OnUpdate; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnLoad; override;
    function RetourneLibelleAvecDimensions(RefArtDim: string; TOBArt: TOB): string;
  end;

const
  // libellés des messages
  TexteMessage: array[1..13] of string = (
    {1}'Impossible de déterminer le client par défaut.'
    {2}, 'Erreur à la création de la pièce.'
    {3}, 'Erreur à la création d''une ligne.'
    {4}, 'Impossible de trouver cet article :'
    {5}, 'Aucun état disponible.'
    {6}, 'Erreur à la création des documents.'
    {7}, 'Erreur à la validation de la pièce.'
    {8}, 'Aucune pièce à générer.'
    {9}, 'Un problème sur un article, pas de ligne générique !'
    {10}, 'Aucun état sélectionné pour l''impression.'
    {11}, 'Vous devez sélectionner un dossier existant.'
    {12}, 'Vous devez renseigner un type de fichier.'
    {13}, 'Impossible de sélectionner les champs du fichier ASCII.'
    );

  NbArtParRequete: integer = 1000;

implementation

procedure TOF_MBOGENETRANSF.SetLastError(Num: integer; ou: string);
begin
  if ou <> '' then SetFocusControl(ou);
  LastError := -1;
  PGIError(TraduireMemoire(TexteMessage[Num]), Ecran.Caption);
end;

function TOF_MBOGENETRANSF.RetourneLibelleAvecDimensions(RefArtDim: string; TOBArt: TOB): string;
var k: integer;
  Grille, CodeDim, LibDim: string;
begin
  LibDim := '';
  if TOBArt <> nil then
  begin
    for k := 1 to 5 do
    begin
      Grille := TOBArt.GetValue('GA_GRILLEDIM' + IntToStr(k));
      CodeDim := TOBArt.GetValue('GA_CODEDIM' + IntToStr(k));
      if ((Grille <> '') and (CodeDim <> '')) then
        LibDim := LibDim + '  ' + RechDom('GCGRILLEDIM' + IntToStr(k), Grille, True) + ' ' + GCGetCodeDim(Grille, CodeDim, k);
    end;
  end;
  Result := LibDim;
end;

procedure TOF_MBOGENETRANSF.OnUpdate;
begin
  inherited;
  CanClose := False;
  try
    //if TRadioButton(GetControl('TRANSFERTPGI')).checked then GenerationPGI
    if TransfertPGI then GenerationPGI
    else GenerationFichierASCII;
  finally
    CanClose := LastError <> -1;
  end;
end;

procedure TOF_MBOGENETRANSF.OnArgument(S: string);
begin
  inherited;
  CodePtrf := S;
  CanClose := True;
  BGenerationCF := Ecran.Name = 'PR_GENECF';
  StGCORLIHOST := GetParamSoc('SO_GCORLIHOST');
  StGCORLIUSERNAME := GetParamSoc('SO_GCORLIUSERNAME');
  StGCORLIPASSWORD := GetParamSoc('SO_GCORLIPASSWORD');
  StGCORLIREPFTP := GetParamSoc('SO_GCORLIREPFTP');
end;

procedure TOF_MBOGENETRANSF.OnClose;
begin
  if CanClose then LastError := 0 else LastError := -1;
  CanClose := True;
  inherited;
end;

procedure TOF_MBOGENETRANSF.OnLoad;
var Q: TQuery;
  StSQL: string;
begin
  CodeDepotEmmetteur := '';
  if BGenerationCF then NaturePiece := 'CF' else NaturePiece := 'TEM';
  TCheckbox(GetControl('EDITAUTOTRANSF')).Checked := GetInfoParPiece(NaturePiece, 'GPP_IMPIMMEDIATE') = 'X';
  TCheckbox(GetControl('APERCU')).Checked := GetInfoParPiece(NaturePiece, 'GPP_APERCUAVIMP') = 'X';
  EditionClick;
  StSQL := 'Select GTE_DEPOTEMET,GTE_COMMANDEFOU,GTE_GENEREPTRF FROM PROPTRANSFENT WHERE GTE_CODEPTRF="' + CodePtrf + '"';
  Q := OpenSQL(StSQL, True);
  if not Q.EOF then
  begin
    CodeDepotEmmetteur := Q.FindField('GTE_DEPOTEMET').AsString;
    if CodeDepotEmmetteur = '' then PGIError('Pas de dépôt émetteur défini, impossible de faire la génération.', Ecran.Caption)
    else
    begin
      if ((BGenerationCF) and (Q.FindField('GTE_COMMANDEFOU').AsString = 'X'))
        or ((not BGenerationCF) and (Q.FindField('GTE_GENEREPTRF').AsString = 'X')) then
      begin
        TGroupBox(GetControl('GROUPOPT')).Caption := 'Traitement déjà généré';
        TGroupBox(GetControl('GROUPOPT')).Font.Color := clRed;
        SetControlEnabled('TRANSFERTPGI', False);
        ChoixTransfert('ASCII');
      end
      else
      begin
        if DepotGererSurSite(CodeDepotEmmetteur) then ChoixTransfert('PGI')
        else ChoixTransfert('ASCII');
      end;
    end;
  end;
  Ferme(Q);
end;

procedure TOF_MBOGENETRANSF.Affichage(Etape: integer; StDetail: string = '');
var i: integer;
begin
  case Etape of
    1:
      begin
        BImpression := TCheckbox(GetControl('EDITAUTOTRANSF')).checked;
        if BImpression then SetControlVisible('TRF4', True);
        SetControlVisible('GROUPOPT', False);
        SetControlVisible('GROUPINFO', True);
      end;
    2:
      begin
        SetControlVisible('MOUETTEGRISE1', False);
        SetControlVisible('MOUETTEBLEUE1', True);
        THLabel(GetControl('PARTIE1')).Font.Style := [];
        SetControlEnabled('PARTIE2', True);
        THLabel(GetControl('PARTIE2')).Font.Style := [fsBold];
      end;
    3: SetControlCaption('PARTIE2', 'Validation des transferts : ' + StDetail);
    4:
      begin
        for i := 1 to 3 do SetControlEnabled('TRF' + IntToStr(i), True);
        if BImpression then SetControlEnabled('TRF4', True);
      end;
    5:
      begin
        if BImpression then THLabel(GetControl('TRF4')).Font.Style := []
        else THLabel(GetControl('TRF3')).Font.Style := [];
        THLabel(GetControl('TRF1')).Font.Style := [fsBold];
      end;
    6:
      begin
        THLabel(GetControl('TRF1')).Font.Style := [];
        THLabel(GetControl('TRF2')).Font.Style := [fsBold];
      end;
    7:
      begin
        THLabel(GetControl('TRF2')).Font.Style := [];
        THLabel(GetControl('TRF3')).Font.Style := [fsBold];
      end;
    8:
      begin
        THLabel(GetControl('TRF3')).Font.Style := [];
        THLabel(GetControl('TRF4')).Font.Style := [fsBold];
      end;
    9:
      begin
        SetControlVisible('MOUETTEGRISE2', False);
        SetControlVisible('MOUETTEBLEUE2', True);
      end;
  end;
  Application.ProcessMessages;
end;

procedure TOF_MBOGENETRANSF.MAJBlocNote(Etape: integer; NbTransfert: integer = 0);
var BlocNote: TStringList;
  GenerationEnCours: Boolean;
  TOBProp: TOB;
begin
  { exemple de bloc notes :
  ------------------------------------
  Début génération : 28/12/01 15:05
  Nombre total de transferts : 3
  Restant à générer : 3
  Reprise génération : 28/12/01 15:05
  Restant à générer : 2
  Restant à générer : 1
  Fin génération : 28/12/01 à 15:20
  ------------------------------------ }
  BlocNote := TStringList.Create;
  TOBProp := TOB.Create('PROPTRANSFENT', nil, -1);
  TOBProp.SelectDB('"' + CodePtrf + '"', nil);
  BlocNote.Text := TOBProp.GetValue('GTE_BLOCNOTE');
  if BGenerationCF then GenerationEnCours := TOBProp.GetValue('GTE_ENCOURSCF') = 'X'
  else GenerationEnCours := TOBProp.GetValue('GTE_ENCOURSGENE') = 'X';
  with BlocNote do
  begin
    case Etape of
      1:
        begin
          if not GenerationEnCours then
          begin
            Add('++++++++++++++++++++++++++++++++');
            Add('Début génération : ' + FormatDateTime('dd/mm/yy tt', Now));
            if BGenerationCF then Add('Nb total commandes fournisseurs : ' + IntToStr(NbTransfert))
            else Add('Nombre total de transferts : ' + IntToStr(NbTransfert));
          end
          else Add('Reprise : ' + FormatDateTime('dd/mm/yy tt', Now));
        end;
      2: Add('Restant à générer : ' + IntToStr(NbTransfert - 1));
      3:
        begin
          Add('Fin génération : ' + FormatDateTime('dd/mm/yy tt', Now));
          Add('++++++++++++++++++++++++++++++++');
        end;
      4:
        begin
          Add('++++++++++++++++++++++++++++++++');
          Add('Début génération : ' + FormatDateTime('dd/mm/yy tt', Now));
        end;
      5: Add('Fichier ASCII de : ' + IntToStr(NbTransfert) + ' enregistrements.');
      6: Add('Export par FTP du fichier ASCII.');
    end;
  end;
  TOBProp.PutValue('GTE_BLOCNOTE', BlocNote.Text);
  TOBProp.UpdateDB;
  BlocNote.Free;
  TOBProp.Free;
end;

procedure TOF_MBOGENETRANSF.GenerationFichierASCII;
var PathFichier, StSQL, UneLigne, CodeExport, ListeDesChamps, Sep: string;
  DateFormat, BufferAlpha, BufferNumerique, StPositive: string;
  BFormatDynamique: Boolean;
  StGEX_NATURE: string;
  Q, QSEp: TQuery;
  NbRecords: integer;
  FichierASCII: TextFile;
  TOB_CHAMPSELECT: TOB;
begin
  NbRecords := 0;
  // Test si la création du fichier ASCII peut se faire
  CodeExport := GetControlText('GAS_CODEEX');
  if CodeExport = '' then
  begin
    SetLastError(12, 'GAS_CODEEX');
    exit;
  end;
  PathFichier := GetControlText('EMPLACEMENT');
  if PathFichier = '' then
  begin
    SetLastError(11, 'EMPLACEMENT');
    exit;
  end;
  MAJBlocNote(4);

  //Charge les champs du fichier ASCII
  TOB_CHAMPSELECT := Tob.Create('Table EXPORTASCII', nil, -1);
  if BGenerationCF then StGEX_NATURE := 'GCF' else StGEX_NATURE := 'MPT';
  Q := OpenSQL('SELECT * FROM EXPORTASCII WHERE GEX_NATURE="' + StGEX_NATURE +
    '" AND GEX_CODEEX="' + CodeExport + '" ORDER BY GEX_POSITION', True);
  if not Q.Eof then TOB_CHAMPSELECT.LoadDetailDB('EXPORTASCII', '', '', Q, False)
  else
  begin
    Ferme(Q);
    SetLastError(12, '');
    exit;
  end;
  Ferme(Q);

  ASCIIAjouteChampGaucheEtDate(TOB_CHAMPSELECT);

  ListeDesChamps := ASCIIConstruireListeDeChamps(TOB_CHAMPSELECT);

  if BGenerationCF then
    StSQL := 'Select ' + ListeDesChamps +
      ' From PROPTRANSFLIG Left Join ARTICLE On GTL_ARTICLE=GA_ARTICLE' +
      ' Where GTL_CODEPTRF="' + CodePtrf + '" and GTL_DEPOTDEST="..."' +
      ' and GTL_MANQUANT>0 and GTL_STATUTART<>"GEN"'
  else
  begin
    if GetCheckBoxState('PROPOSITION_POSITIVE') = cbChecked then
      StPositive := 'and GTL_PROPOSITION>0 ' else StPositive := '';
    StSQL := 'Select ' + ListeDesChamps +
      ' From PROPTRANSFLIG Left Join ARTICLE On GTL_ARTICLE=GA_ARTICLE' +
      ' Where GTL_CODEPTRF="' + CodePtrf + '" and GTL_DEPOTDEST<>"..." ' +
      StPositive + 'and GTL_STATUTART<>"GEN"';
  end;

  Q := OpenSQL(StSQL, True);

  try
    AssignFile(FichierASCII, PathFichier);
    Rewrite(FichierASCII);

    QSep := OpenSQL('SELECT GAS_SEPARATEUR,GAS_DATEFORMAT,GAS_ALPHA,GAS_NUMERIQUE,GAS_FORMATFIXE ' +
      'FROM EXPORTASCII_ENTETE WHERE GAS_NATURE="' + StGEX_NATURE +
      '" AND GAS_CODEEX="' + CodeExport + '"', True);
    Sep := '';
    DateFormat := '';
    BufferAlpha := ' ';
    BufferNumerique := ' ';
    BFormatDynamique := False;
    if not QSep.Eof then
    begin
      Sep := QSep.FindField('GAS_SEPARATEUR').AsString;
      DateFormat := QSep.FindField('GAS_DATEFORMAT').AsString;
      BufferAlpha := QSep.FindField('GAS_ALPHA').AsString;
      BufferNumerique := QSep.FindField('GAS_NUMERIQUE').AsString;
      BFormatDynamique := QSep.FindField('GAS_FORMATFIXE').AsString = '-';
    end;
    Ferme(QSep);

    Q.First;
    while not Q.Eof do
    begin
      Inc(NbRecords);
      UneLigne := ASCIIConstruireUneLigne(Q, TOB_CHAMPSELECT, CodeExport, Sep,
        DateFormat, BufferAlpha, BufferNumerique, BFormatDynamique);
      WriteLn(FichierASCII, UneLigne);
      Q.Next;
    end;

    if TOB_CHAMPSELECT <> nil then TOB_CHAMPSELECT.Free;
    CloseFile(FichierASCII);
    MAJBlocNote(5, NbRecords);

    //Si EXPORTFTP est coché : envoie le fichier par FTP
    if GetCheckBoxState('EXPORTFTP') = cbChecked then
    begin
      ExportDuFichierParFTP;
      MAJBlocNote(6, NbRecords);
    end;

    AfterGeneration;
    MAJBlocNote(3);
  finally
    Ferme(Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 19/12/2002
Modifié le ... :   /  /
Description .. : Création des documents :
Suite ........ : - transferts
Suite ........ : ou
Suite ........ : - commandes fournisseur
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOGENETRANSF.GenerationPGI;
var TOBFondation: TOB;
begin
  if CodePtrf = '' then exit;
  Affichage(1);
  TOBFondation := TOB.Create('', nil, -1);
  try
    ConstruireLesFondations(TOBFondation);
    if TOBFondation <> nil then
    begin
      MAJBlocNote(1, TOBFondation.Detail.Count);
      if TOBFondation.Detail.Count = 0 then SetLastError(8, '')
      else ConstruireTousLesDocuments(TOBFondation);
      MAJBlocNote(3);
    end;
  finally
    TOBFondation.Free;
  end;
end;

procedure TOF_MBOGENETRANSF.ConstruireLesFondations(TOBFondation: TOB);
var Q, QDepot: TQuery;
  TOBTrfDepot, TOBCFFournPrinc: TOB;
  StSQL, Boutique, StPositive, FournPrinc, DepotEmett: string;
begin
  QDepot := OpenSQL('Select GTE_DEPOTEMET From PROPTRANSFENT ' +
    'Where GTE_CODEPTRF="' + CodePtrf + '"', True);
  DepotEmett := QDepot.FindField('GTE_DEPOTEMET').AsString;
  Ferme(QDepot);

  if BGenerationCF then
    StSQL := 'Select GA_FOURNPRINC,GTL_ARTICLE,GTL_CODEARTICLE,' +
      ' GTL_MANQUANT,GTL_STATUTART,GTL_UTILISATEUR' +
      ' From PROPTRANSFLIG Left Join ARTICLE On GTL_ARTICLE=GA_ARTICLE' +
      ' Where GTL_CODEPTRF="' + CodePtrf + '" and GTL_DEPOTDEST="..." ' +
      ' and GTL_MANQUANT>0 and GTL_GENECF<>"X" ' +
      ' Order By GA_FOURNPRINC,GTL_ARTICLE'
  else
  begin
    if GetCheckBoxState('PROPOSITION_POSITIVE') = cbChecked then
      StPositive := 'and GTL_PROPOSITION>0 ' else StPositive := '';
    StSQL := 'Select GTL_DEPOTDEST,GTL_ARTICLE,GTL_CODEARTICLE,' +
      ' GTL_PROPOSITION,GTL_STATUTART,GTL_UTILISATEUR' +
      ' From PROPTRANSFLIG' +
      ' Where GTL_CODEPTRF="' + CodePtrf + '" and GTL_DEPOTDEST<>"..." ' +
      StPositive + ' and GTL_GENERER<>"X" ' +
      ' Order By GTL_DEPOTDEST,GTL_ARTICLE';
  end;

  Q := OpenSQL(StSQL, True);
  if not Q.Eof then
  begin
    Q.First;
    Boutique := '';
    FournPrinc := '';
    INITMOVE(100, '');
    TOBTrfDepot := nil;
    TOBCFFournPrinc := nil;
    while not Q.Eof do
    begin
      MoveCur(False);
      if BGenerationCF then
      begin // TOBFondation contient autant de TOBCFFournPrinc qu'il y a de fournisseur principal différent
        if FournPrinc <> Q.FindField('GA_FOURNPRINC').AsString then
        begin
          FournPrinc := Q.FindField('GA_FOURNPRINC').AsString;
          TOBCFFournPrinc := ConstruireTOBCFFournPrinc(TOBFondation, Q, DepotEmett);
        end;
        if TOBCFFournPrinc <> nil then ConstruireTOBTrfArt(TOBCFFournPrinc, Q);
      end
      else
      begin // TOBFondation contient autant de TOBTrfDepot qu'il y a de dépôt récepteur différent
        if Boutique <> Q.FindField('GTL_DEPOTDEST').AsString then
        begin
          Boutique := Q.FindField('GTL_DEPOTDEST').AsString;
          TOBTrfDepot := ConstruireTOBTrfDepot(TOBFondation, Q, DepotEmett);
        end;
        if TOBTrfDepot <> nil then ConstruireTOBTrfArt(TOBTrfDepot, Q);
      end;
      Q.Next;
    end;
    FINIMOVE;
  end;
  Ferme(Q);
end;

function TOF_MBOGENETRANSF.ConstruireTOBTrfDepot(TOBTRF: TOB; QDepot: TQuery; DepotEmetteur: string): TOB;
var TOBTrfD: TOB;
begin
  TOBTrfD := TOB.Create('', TOBTRF, -1);
  with TOBTrfD do
  begin
    AddChampSupValeur('GTE_DEPOTEMET', DepotEmetteur);
    AddChampSupValeur('GTL_DEPOTDEST', QDepot.FindField('GTL_DEPOTDEST').AsString);
    AddChampSupValeur('GTE_UTILISATEUR', QDepot.FindField('GTL_UTILISATEUR').AsString);
  end;
  result := TOBTrfD;
end;

function TOF_MBOGENETRANSF.ConstruireTOBCFFournPrinc(TOBCF: TOB; QFourn: TQuery; DepotEmetteur: string): TOB;
var TOBCFUnique: TOB;
begin
  TOBCFUnique := TOB.Create('', TOBCF, -1);
  with TOBCFUnique do
  begin
    AddChampSupValeur('GTE_DEPOTEMET', DepotEmetteur);
    AddChampSupValeur('GA_FOURNPRINC', QFourn.FindField('GA_FOURNPRINC').AsString);
    AddChampSupValeur('GTE_UTILISATEUR', QFourn.FindField('GTL_UTILISATEUR').AsString);
  end;
  result := TOBCFUnique;
end;

procedure TOF_MBOGENETRANSF.ConstruireTOBTrfArt(TOBTrfDepot: TOB; QArt: TQuery);
var TOBTrfArt: TOB;
begin
  TOBTrfArt := TOB.Create('', TOBTrfDepot, -1);
  with TOBTrfArt do
  begin
    AddChampSupValeur('GTL_ARTICLE', QArt.FindField('GTL_ARTICLE').AsString);
    AddChampSupValeur('GTL_CODEARTICLE', QArt.FindField('GTL_CODEARTICLE').AsString);
    AddChampSupValeur('GTL_STATUTART', QArt.FindField('GTL_STATUTART').AsString);
    if BGenerationCF then AddChampSupValeur('GTL_MANQUANT', QArt.FindField('GTL_MANQUANT').AsInteger)
    else AddChampSupValeur('GTL_PROPOSITION', QArt.FindField('GTL_PROPOSITION').AsInteger);
    AddChampSupValeur('GTL_UTILISATEUR', QArt.FindField('GTL_UTILISATEUR').AsString);
  end;
end;

procedure TOF_MBOGENETRANSF.ConstruireTousLesDocuments(TOBTRF: TOB);
var i, j, NbAdresses: integer;
  TOBP: TOB;
  BApercu: boolean;
  StSQL, StModele: string;
begin
  if BGenerationCF then StSQL := 'Update PROPTRANSFENT set GTE_ENCOURSCF="X" where GTE_CODEPTRF="' + CodePtrf + '"'
  else StSQL := 'Update PROPTRANSFENT set GTE_ENCOURSGENE="X" where GTE_CODEPTRF="' + CodePtrf + '"';
  ExecuteSQL(StSQL);
  TOBToutesLesPieces := TOB.CREATE('Toutes les pièces', nil, -1);
  if BGenerationCF then INITMOVE(TOBTRF.Detail.Count, '') else INITMOVE(TOBTRF.Detail.Count * 2, '');
  try
    Affichage(2);
    Affichage(4);
    QualifMvt := GetInfoParPiece(NaturePiece, 'GPP_QUALIFMVT');
    BApercu := TCheckbox(GetControl('APERCU')).checked;
    StModele := GetInfoParPiece(NaturePiece, 'GPP_IMPETAT');
    for i := 0 to TOBTRF.Detail.Count - 1 do
    begin
      /////////////////////////////////////////////////////
      // Création des TOB Virtuelles rattachées à la pièce
      /////////////////////////////////////////////////////
      Tob_PiedBase := TOB.CREATE('Les Taxes', nil, -1);
      Tob_PiedEche := TOB.CREATE('Les Echeances', nil, -1);
      Tob_Catalogue := TOB.CREATE('Les Catalogues', nil, -1);
      Tob_noment := TOB.CREATE('Les Noments', nil, -1);
      Tob_Article := TOB.CREATE('Les Articles', nil, -1);
      try
        Affichage(3, IntToStr(i + 1) + '/' + IntToStr(TOBTRF.Detail.Count));
        if (i = TOBTRF.Detail.Count - 1) then Affichage(9);
        Affichage(5);
        MoveCur(False);
        TOBP := ConstruirePiece(TOBTRF.Detail[i]);
        if TOBP = nil then
        begin
          SetLastError(2, '');
          Break;
        end;
        EnHT := TOBP.GetValue('GP_FACTUREHT') = 'X';
        if not ConstruireLignes(TOBP, TOBTRF.Detail[i]) then
        begin
          SetLastError(3, '');
          Break;
        end;
        MoveCur(False);
        PutValueDetail(TOBP, 'GP_RECALCULER', 'X');
        MAJPrixLignesGEN(TOBP);
        if TOBP.Detail.Count = 0 then TOBP.Free
        else
        begin
          IncNumSoucheG(CleDoc.Souche);
          //////////////////////////////////////////////////////////////////
          // Gestion des adresses
          //////////////////////////////////////////////////////////////////
          Tob_Adresses := TOB.Create('LES ADRESSES', nil, -1);
          TOB.Create('PIECEADRESSE', Tob_Adresses, -1); {Livraison}
          TOB.Create('PIECEADRESSE', Tob_Adresses, -1); {Facturation}
          for NbAdresses := 0 to Tob_Adresses.Detail.Count - 1 do GetAdrFromCode(Tob_Adresses.Detail[NbAdresses], TOB_Tiers.GetValue ('T_TIERS'));

          CalculFacture(TOBP, Tob_PiedBase, Tob_Tiers, Tob_Article, nil, nil, nil, DEV);
          if not BGenerationCF then ConstruireTRVouTRE(TOBP);
          // Valide les pieces TEM et TRE  ou  CF
          if ValideLesPieces(TOBTRF.Detail[i]) then
          begin
            MAJBlocNote(2, TOBTRF.Detail.Count - i);
            //Impression du Transfert Emis
            if BImpression then
            begin
              Affichage(8);
              if StModele = '' then SetLastError(10, '')
              else ImpressionDesEtats(BApercu, False, StModele);
            end;
          end
          else
          begin
            SetLastError(7, '');
            Break;
          end;
          //Libere les tobs correspondant aux pieces
          for j := TOBToutesLesPieces.Detail.Count - 1 downto 0 do TOBToutesLesPieces.Detail[j].Free;
        end;
      finally
        Tob_PiedBase.Free;
        Tob_PiedEche.Free;
        Tob_Catalogue.Free;
        Tob_noment.Free;
        Tob_Article.Free;
        Tob_Adresses.Free;
      end;
    end;
    AfterGeneration;
  finally
    TOBToutesLesPieces.Free;
    Tob_Tiers.Free;
    FINIMOVE;
  end;
end;

function TOF_MBOGENETRANSF.ConstruirePiece(TOBTrfDepot: TOB): TOB;
var Etab: string;
  DD: TDateTime;
  Per, Sem: integer;
  Tob_Piece: TOB; // TOB entête de pièce
begin
  result := nil;
  Numligne := 0;
  Etab := TOBTrfDepot.GetValue('GTE_DEPOTEMET');
  Tob_Piece := TOB.CREATE('PIECE', TOBToutesLesPieces, -1);
  Tob_Piece.InitValeurs;
  AddLesSupEntete(Tob_Piece);
  InitTOBPiece(TOB_Piece); // Init date de création, modif, utilisateur ......
  Tob_Piece.PutValue('GP_NATUREPIECEG', NaturePiece);
  /////////////////////////
  // Souche, Numéro, .....
  ////////////////////////
  CleDoc.Souche := GetSoucheG(NaturePiece, Etab, '');
  CleDoc.NumeroPiece := GetNumSoucheG(CleDoc.Souche);
  CleDoc.DatePiece := V_PGI.DateEntree;

  Tob_Piece.PutValue('GP_DATEPIECE', CleDoc.DatePiece);
  PutValueDetail(Tob_Piece, 'GP_SOUCHE', CleDoc.Souche);
  PutValueDetail(Tob_Piece, 'GP_NUMERO', CleDoc.NumeroPiece);
  /////////////////////////////////////////////////////////////////////////
  // Etablissement / Dépôt Emetteur / Dépôt Destinataire
  /////////////////////////////////////////////////////////////////////////
  Tob_Piece.PutValue('GP_ETABLISSEMENT', Etab);
  Tob_Piece.PutValue('GP_DEPOT', Etab);
  if not BGenerationCF then Tob_Piece.PutValue('GP_DEPOTDEST', TOBTrfDepot.GetValue('GTL_DEPOTDEST'));
  ///////////////////////////////////////////////////////////////
  // Chargement en TOB du client
  ///////////////////////////////////////////////////////////////
  if not ChargeTOBClient(TOBTrfDepot) then exit;
  Tob_Piece.PutValue('GP_TIERS', CodeTiers);
  Tob_Piece.PutValue('GP_TIERSLIVRE', CodeTiers);
  Tob_Piece.PutValue('GP_TIERSFACTURE', Tob_Tiers.GetValue('T_FACTURE'));
  Tob_Piece.PutValue('GP_TIERSPAYEUR', Tob_Tiers.GetValue('T_PAYEUR'));
  ////////////////////////////////////////////////////////////////////////////
  // Initialise les zones de l'entête en fonction du client
  ////////////////////////////////////////////////////////////////////////////
  Tob_Piece.PutValue('GP_ESCOMPTE', Tob_Tiers.GetValue('T_ESCOMPTE'));
  Tob_Piece.PutValue('GP_MODEREGLE', Tob_Tiers.GetValue('T_MODEREGLE'));
  Tob_Piece.PutValue('GP_REGIMETAXE', Tob_Tiers.GetValue('T_REGIMETVA'));
  Tob_Piece.PutValue('GP_TVAENCAISSEMENT', Tob_Tiers.GetValue('T_TVAENCAISSEMENT'));
  Tob_Piece.PutValue('GP_QUALIFESCOMPTE', Tob_Tiers.GetValue('T_QUALIFESCOMPTE'));
  Tob_Piece.PutValue('GP_TARIFTIERS', Tob_Tiers.GetValue('T_TARIFTIERS'));
  /////////////////////////////////////////////////////////////////////////////
  // Facturation HT ou TTC
  /////////////////////////////////////////////////////////////////////////////
  Tob_Piece.PutValue('GP_FACTUREHT', Tob_Tiers.GetValue('T_FACTUREHT'));
  Tob_Piece.PutValue('GP_REPRESENTANT', ChoixCommercial('REP', Tob_Tiers.GetValue('T_ZONECOM')));
  ///////////////////////////////////////////////
  // Champs par  défaut
  //////////////////////////////////////////////
  Tob_Piece.PutValue('GP_UTILISATEUR', TOBTrfDepot.GetValue('GTE_UTILISATEUR'));
  if BGenerationCF then Tob_Piece.PutValue('GP_VENTEACHAT', 'ACH')
  else Tob_Piece.PutValue('GP_VENTEACHAT', 'TRF'); // Document  : Transfert
  Tob_Piece.PutValue('GP_NUMADRESSELIVR', -1);
  Tob_Piece.PutValue('GP_NUMADRESSEFACT', -1);
  Tob_Piece.PutValue('GP_CREATEUR', V_PGI.User);
  Tob_Piece.PutValue('GP_SOCIETE', V_PGI.CodeSociete);
  /////////////////////////
  // Devise du document
  ////////////////////////
  Tob_Piece.PutValue('GP_DEVISE', V_PGI.DevisePivot);

  DEV.Code := Tob_Piece.GetValue('GP_DEVISE');
  GETINFOSDEVISE(DEV);
  DEV.Taux := GETTAUX(DEV.Code, DEV.DateTaux, CleDoc.DatePiece);
  PutValueDetail(Tob_Piece, 'GP_DEVISE', DEV.Code);
  PutValueDetail(Tob_Piece, 'GP_TAUXDEV', DEV.Taux);
  PutValueDetail(Tob_Piece, 'GP_COTATION', DEV.Taux);
  PutValueDetail(Tob_Piece, 'GP_DATETAUXDEV', DEV.DateTaux);

  DD := Tob_Piece.GetValue('GP_DATEPIECE');
  Per := GetPeriode(DD);
  Sem := NumSemaine(DD);
  Tob_Piece.PutValue('GP_PERIODE', Per);
  Tob_Piece.PutValue('GP_SEMAINE', Sem);
  result := Tob_Piece;
end;

function TOF_MBOGENETRANSF.ChargeTOBClient(TOBCF: TOB): Boolean;
var Q: TQuery;
  BChargeTiers: boolean;
begin
  result := true;
  BChargeTiers := Tob_Tiers = nil;

  if (BGenerationCF) and not (BChargeTiers)
    and (TOBCF.GetValue('GA_FOURNPRINC') <> CodeTiers) then
  begin
    Tob_Tiers.Free;
    Tob_Tiers := nil;
    BChargeTiers := True;
  end;

  if BChargeTiers then
  begin
    if BGenerationCF then CodeTiers := TOBCF.GetValue('GA_FOURNPRINC')
    else CodeTiers := GetParamsoc('SO_GCTIERSDEFAUT');

    Q := OpenSQL('Select * from Tiers Where T_TIERS="' + CodeTiers + '"', False);
    if not Q.EOF then
    begin
      Tob_Tiers := TOB.CREATE('TIERS', nil, -1);
      Tob_Tiers.initValeurs;
      Tob_Tiers.SelectDB('', Q);
    end
    else
    begin
      SetLastError(1, '');
      result := False;
    end;
    Ferme(Q);
  end;
end;

function TOF_MBOGENETRANSF.ConstruireLignes(TOBPiece, TOBTrfDepot: TOB): boolean;
var i: integer;
  Continue: boolean;
begin
  i := 0;
  Continue := True;
  BuildTob_ArticleAndDispo(TOBTrfDepot);
  Affichage(6);
  while (i < TOBTrfDepot.Detail.Count) and (Continue) do
  begin
    Continue := CreatLignePiece(TOBPiece, TOBTrfDepot.Detail[i]);
    i := i + 1;
  end;
  result := Continue;
end;

///////////////////////////////////////////////////////////////////////////////////////
// Création d'une ligne de document
///////////////////////////////////////////////////////////////////////////////////////

function TOF_MBOGENETRANSF.CreatLignePiece(TOBPiece, TobTrfArt: TOB): boolean;
var TobArt: TOB;
  Tob_Ligne: TOB; // TOB lignes
  ReferenceArtDim, ReferenceArtGen, SQL, TypeDim, Utilisateur, Designation: string;
  Q: TQuery;
  i: integer;
  Proposition: double;
  Prix: double;
begin
  result := True;

  ReferenceArtDim := TobTrfArt.GetValue('GTL_ARTICLE');
  ReferenceArtGen := TobTrfArt.GetValue('GTL_CODEARTICLE');
  TypeDim := TobTrfArt.GetValue('GTL_STATUTART');
  Utilisateur := TobTrfArt.GetValue('GTL_UTILISATEUR');
  if BGenerationCF then Proposition := TobTrfArt.GetValue('GTL_MANQUANT')
  else Proposition := TobTrfArt.GetValue('GTL_PROPOSITION');

  //Empêche de créer des lignes avec une quantité à zéro
  if Proposition = 0 then exit;

  Tob_Ligne := TOB.CREATE('LIGNE', TobPiece, -1);
  Tob_Ligne.initValeurs;
  ////////////////////////////////////////////////////////////////
  // Initialise les champs de la ligne à partir de l'entête
  ////////////////////////////////////////////////////////////////
  PieceVersLigne(TobPiece, Tob_Ligne);
  Tob_Ligne.PutValue('GL_PERIODE', TobPiece.GetValue('GP_PERIODE'));
  Tob_Ligne.PutValue('GL_SEMAINE', TobPiece.GetValue('GP_SEMAINE'));
  /////////////////////////////////////////////
  // Nouveau numéro de ligne
  //////////////////////////////////////////////
  NumLigne := NumLigne + 1;
  Tob_Ligne.PutValue('GL_NUMLIGNE', NumLigne);
  Tob_Ligne.PutValue('GL_NUMORDRE', NumLigne);
  ////////////////////////////////////////////////////////////////
  // Ajoute des champs supplémentaires pour le calcul des cumuls
  ////////////////////////////////////////////////////////////////
  //AddLesSupLigne (Tob_Ligne, True) ; // Ajout Champs spécifiques pour calcul de la pièce
  AddLesSupLigne(Tob_Ligne, False);
  ///////////////////////////////////////////////////////////////
  // Détermination de l'identifiant article
  //////////////////////////////////////////////////////////////
  TobArt := Tob_Article.FindFirst(['GA_ARTICLE'], [ReferenceArtDim], False);
  if TobArt = nil then //Ne doit pas passer ici
  begin
    SQL := 'Select * From ARTICLE WHERE GA_ARTICLE="' + ReferenceArtDim + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      TobArt := CreerTOBArt(Tob_Article);
      TobArt.SelectDB('', Q);
      if TypeDim <> 'GEN' then LoadTOBDispo(TobArt, True);
    end
    else
    begin
      PGIError(TraduireMemoire(TexteMessage[4]) + ' ' + ReferenceArtDim, TFVierge(Ecran).Caption);
      result := False;
    end;
    Ferme(Q);
  end;
  if TobArt <> nil then
  begin
    /////////////////////////////////////////////////////////////////////////////
    // Initialisation des champs de la ligne à partir de l'article dimensionné
    /////////////////////////////////////////////////////////////////////////////
    Designation := TobArt.GetValue('GA_LIBELLE');
    if TypeDim <> 'GEN' then
    begin
      Tob_Ligne.PutValue('GL_PRIXPOURQTE', TobArt.GetValue('GA_PRIXPOURQTE'));
      Tob_Ligne.PutValue('GL_ESCOMPTABLE', TobArt.GetValue('GA_ESCOMPTABLE'));
      Tob_Ligne.PutValue('GL_REMISABLEPIED', TobArt.GetValue('GA_REMISEPIED'));
      Tob_Ligne.PutValue('GL_REMISABLELIGNE', TobArt.GetValue('GA_REMISELIGNE'));
      Tob_Ligne.PutValue('GL_TENUESTOCK', TobArt.GetValue('GA_TENUESTOCK'));
      Tob_Ligne.PutValue('GL_TARIFARTICLE', TobArt.GetValue('GA_TARIFARTICLE'));
      Tob_Ligne.PutValue('GL_QUALIFSURFACE', TobArt.GetValue('GA_QUALIFSURFACE'));
      Tob_Ligne.PutValue('GL_QUALIFVOLUME', TobArt.GetValue('GA_QUALIFVOLUME'));
      Tob_Ligne.PutValue('GL_QUALIFPOIDS', TobArt.GetValue('GA_QUALIFPOIDS'));
      Tob_Ligne.PutValue('GL_QUALIFLINEAIRE', TobArt.GetValue('GA_QUALIFLINEAIRE'));
      Tob_Ligne.PutValue('GL_QUALIFHEURE', TobArt.GetValue('GA_QUALIFHEURE'));
      Tob_Ligne.PutValue('GL_SURFACE', TobArt.GetValue('GA_SURFACE'));
      Tob_Ligne.PutValue('GL_VOLUME', TobArt.GetValue('GA_VOLUME'));
      Tob_Ligne.PutValue('GL_POIDSBRUT', TobArt.GetValue('GA_POIDSBRUT'));
      Tob_Ligne.PutValue('GL_POIDSNET', TobArt.GetValue('GA_POIDSNET'));
      Tob_Ligne.PutValue('GL_POIDSDOUA', TobArt.GetValue('GA_POIDSDOUA'));
      Tob_Ligne.PutValue('GL_LINEAIRE', TobArt.GetValue('GA_LINEAIRE'));
      Tob_Ligne.PutValue('GL_HEURE', TobArt.GetValue('GA_HEURE'));
      Tob_Ligne.PutValue('GL_QUALIFQTESTO', TobArt.GetValue('GA_QUALIFUNITESTO'));
      Tob_Ligne.PutValue('GL_FAMILLENIV1', TobArt.GetValue('GA_FAMILLENIV1'));
      Tob_Ligne.PutValue('GL_FAMILLENIV2', TobArt.GetValue('GA_FAMILLENIV2'));
      Tob_Ligne.PutValue('GL_FAMILLENIV3', TobArt.GetValue('GA_FAMILLENIV3'));
      for i := 1 to 9 do Tob_Ligne.PutValue('GL_LIBREART' + IntToStr(i), TobArt.GetValue('GA_LIBREART' + IntToStr(i)));
      Tob_Ligne.PutValue('GL_LIBREARTA', TobArt.GetValue('GA_LIBREARTA'));
    end;
  end;
  for i := 1 to 5 do Tob_Ligne.PutValue('GL_FAMILLETAXE' + IntToStr(i), TobArt.GetValue('GA_FAMILLETAXE' + IntToStr(i)));
  if result then
  begin
    if TypeDim = 'GEN' then
    begin
      Tob_Ligne.PutValue('GL_LIBELLE', Designation);
      Tob_Ligne.PutValue('GL_TYPELIGNE', 'COM'); // Ligne commentaire
      Tob_Ligne.PutValue('GL_CODESDIM', ReferenceArtGen);
    end
    else
    begin
      Tob_Ligne.PutValue('GL_ARTICLE', ReferenceArtDim);
      Tob_Ligne.PutValue('GL_CODEARTICLE', ReferenceArtGen);
      if TypeDim = 'DIM' then Tob_Ligne.PutValue('GL_LIBELLE', Copy(Designation + RetourneLibelleAvecDimensions(ReferenceArtDim, TOBArt), 1, 70))
      else Tob_Ligne.PutValue('GL_LIBELLE', Designation);
      Tob_Ligne.PutValue('GL_TARIF', 0);
      Tob_Ligne.PutValue('GL_TYPELIGNE', 'ART');
    end;
    Tob_Ligne.PutValue('GL_REFARTSAISIE', ReferenceArtGen);
    Tob_Ligne.PutValue('GL_TYPEARTICLE', 'MAR');
    Tob_Ligne.PutValue('GL_VALIDECOM', 'AFF');
    if TypeDim = 'UNI' then Tob_Ligne.PutValue('GL_TYPEDIM', 'NOR')
    else Tob_Ligne.PutValue('GL_TYPEDIM', TypeDim);
    Tob_Ligne.PutValue('GL_PRIXPOURQTE', 1);
    Tob_Ligne.PutValue('GL_QUALIFMVT', QualifMvt);
    ////////////////////////////////////////////////////////////////////////////
    // Traite les quantités
    ////////////////////////////////////////////////////////////////////////////
    Tob_Ligne.PutValue('GL_QTESTOCK', Proposition);
    Tob_Ligne.PutValue('GL_QTEFACT', Proposition);
    Tob_Ligne.PutValue('GL_QTERESTE', Proposition); { NEWPIECE }
    /////////////////////////////////////////////////////////
    // Récupération des Prix
    //////////////////////////////////////////////////////////
    if TypeDim <> 'GEN' then
    begin
      Prix := QuelPrixBase(NaturePiece, TOBPiece.GetValue('GP_DEPOT'), TobArt, Tob_Ligne);
      if EnHT then
      begin
        Tob_Ligne.PutValue('GL_PUHT', Prix);
        Tob_Ligne.PutValue('GL_PUHTDEV', Prix);
      end
      else
      begin
        Tob_Ligne.PutValue('GL_PUTTC', Prix);
        Tob_Ligne.PutValue('GL_PUTTCDEV', Prix);
      end;
      AffectePrixValo(Tob_Ligne, TobArt);
    end;
  end
  else
    Tob_Ligne.free;
end;

procedure TOF_MBOGENETRANSF.BuildTob_ArticleAndDispo(TOBDepot: TOB);
var i, NbArt, x, CountStArt, NbRequete, y: integer;
  StSelect, StSelectDepot, StArticle, ListeDepot: string;
  TabWhere, TabWhereDepot: array of string;
  TOBDispo, TOBArt, TobDispoArt: TOB;
  QArticle, QDepot: TQuery;
begin
  StSelect := 'SELECT GA_ARTICLE,GA_LIBELLE,GA_GRILLEDIM1,GA_GRILLEDIM2,' +
    'GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5,GA_CODEDIM1,GA_CODEDIM2,' +
    'GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5,GA_PRIXPOURQTE,GA_ESCOMPTABLE,' +
    'GA_REMISEPIED,GA_REMISELIGNE,GA_TENUESTOCK,GA_TARIFARTICLE,GA_QUALIFSURFACE,' +
    'GA_QUALIFVOLUME,GA_QUALIFPOIDS,GA_QUALIFLINEAIRE,GA_QUALIFHEURE,GA_SURFACE,' +
    'GA_VOLUME,GA_POIDSBRUT,GA_POIDSNET,GA_POIDSDOUA,GA_LINEAIRE,GA_HEURE,' +
    'GA_QUALIFUNITESTO,GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3,GA_LIBREART1,' +
    'GA_LIBREART2,GA_LIBREART3,GA_LIBREART4,GA_LIBREART5,GA_LIBREART6,GA_LIBREART7,' +
    'GA_LIBREART8,GA_LIBREART9,GA_LIBREARTA,GA_FAMILLETAXE1,GA_FAMILLETAXE2,' +
    'GA_FAMILLETAXE3,GA_FAMILLETAXE4,GA_FAMILLETAXE5,' +
    'GA_DPA,GA_DPR,GA_PAHT,GA_PRHT,GA_PVHT,GA_PVTTC FROM ARTICLE';
  StSelectDepot := 'SELECT GQ_ARTICLE,GQ_DEPOT,GQ_CLOTURE,GQ_DATECLOTURE,GQ_DPA,GQ_DPR,GQ_PMAP,' +
    'GQ_PMRP,GQ_PHYSIQUE,GQ_TRANSFERT,GQ_RESERVEFOU,' +
    'GQ_RESERVECLI,GQ_PREPACLI FROM DISPO';
  CountStArt := 0;
  NbRequete := 0;
  ListeDepot := '';
  SetLength(TabWhere, 1);
  SetLength(TabWhereDepot, 1);

  if BGenerationCF then ListeDepot := '"' + TOBDepot.GetValue('GTE_DEPOTEMET') + '"'
  else ListeDepot := '"' + TOBDepot.GetValue('GTE_DEPOTEMET') + '","' + TOBDepot.GetValue('GTL_DEPOTDEST') + '"';

  for i := 0 to TOBDepot.Detail.Count - 1 do
  begin
    StArticle := TOBDepot.Detail[i].GetValue('GTL_ARTICLE');
    if CountStArt >= NbArtParRequete then
    begin
      NbRequete := NbRequete + 1;
      SetLength(TabWhere, NbRequete + 1);
      SetLength(TabWhereDepot, NbRequete + 1);
      CountStArt := 0;
    end;
    if not FindStInTAB(StArticle, TabWhere) then
    begin
      CountStArt := CountStArt + 1;
      if TabWhere[NbRequete] = '' then TabWhere[NbRequete] := '"' + StArticle + '"'
      else TabWhere[NbRequete] := TabWhere[NbRequete] + ',"' + StArticle + '"';
      if TOBDepot.Detail[i].GetValue('GTL_STATUTART') <> 'GEN' then
      begin
        if TabWhereDepot[NbRequete] = '' then TabWhereDepot[NbRequete] := '"' + StArticle + '"'
        else TabWhereDepot[NbRequete] := TabWhereDepot[NbRequete] + ',"' + StArticle + '"';
      end;
    end;
  end;

  if TabWhere[0] <> '' then
  begin
    TOBDispo := TOB.CREATE('Les Dispos', nil, -1);
    for y := Low(TabWhere) to High(TabWhere) do
    begin
      if TabWhere[y] <> '' then
      begin
        QArticle := OpenSQL(StSelect + ' WHERE GA_ARTICLE IN (' + TabWhere[y] + ')', True);
        if not QArticle.EOF then Tob_Article.LoadDetailDB('ARTICLE', '', '', QArticle, True);
        Ferme(QArticle);
      end;
      if TabWhereDepot[y] <> '' then
      begin
        QDepot := OpenSQL(StSelectDepot + ' WHERE GQ_ARTICLE IN (' + TabWhereDepot[y] + ') AND GQ_DEPOT IN (' + ListeDepot + ') AND GQ_CLOTURE="-"', True);
        if not QDepot.EOF then TOBDispo.LoadDetailDB('DISPO', '', '', QDepot, True);
        Ferme(QDepot);
      end;
    end;
    for x := 0 to Tob_Article.detail.count - 1 do
    begin
      Tob_Article.detail[x].AddChampSup('UTILISE', False);
      Tob_Article.detail[x].PutValue('UTILISE', '-');
    end;
    //Affecte les stocks aux articles sélectionnés
    for NbArt := 0 to Tob_Article.detail.Count - 1 do
    begin
      TOBArt := Tob_Article.detail[NbArt];
      //Dépôt émetteur
      TobDispoArt := TOBDispo.FindFirst(['GQ_ARTICLE'], [TOBArt.GetValue('GA_ARTICLE')], False);
      if TobDispoArt <> nil then
      begin
        TobDispoArt.Changeparent(TOBArt, -1);
        //Dépôt récepteur (seulement pour les transferts)
        if not BGenerationCF then
        begin
          TobDispoArt := TOBDispo.FindNext(['GQ_ARTICLE'], [TOBArt.GetValue('GA_ARTICLE')], False);
          if TobDispoArt <> nil then TobDispoArt.Changeparent(TOBArt, -1);
        end;
      end;
    end;
    TOBDispo.Free;
    DispoChampSupp(Tob_Article);
  end;
end;

function TOF_MBOGENETRANSF.FindStInTAB(St: string; Tab: array of string): boolean;
var i: integer;
begin
  result := false;
  for i := Low(Tab) to High(Tab) do
    if Pos(St, Tab[i]) <> 0 then
    begin
      result := true;
      break;
    end;
end;

procedure TOF_MBOGENETRANSF.MAJPrixLignesGEN(TOBP: TOB);
var NbLignes, i, NbDec: integer;
  CodeArtGEN, TypeDim, CodeArt: string;
  TOBLGEN: TOB;
  MoyDPA, MoyPMAP, MoyDPR, MoyPMRP: double;
begin
  i := 0;
  TOBLGEN := nil;
  NbDec := DEV.Decimale;
  while i < TOBP.Detail.Count do
  begin
    TypeDim := TOBP.Detail[i].GetValue('GL_TYPEDIM');
    if TypeDim = 'GEN' then
    begin
      TOBLGEN := TOBP.Detail[i];
      CodeArtGEN := TOBP.Detail[i].GetValue('GL_REFARTSAISIE');
      i := i + 1;
    end
    else
      if TypeDim = 'DIM' then
    begin
      CodeArt := TOBP.Detail[i].GetValue('GL_CODEARTICLE');
      //les prix de la ligne générique correspondent à la moyenne
      MoyDPA := 0;
      MoyPMAP := 0;
      MoyDPR := 0;
      MoyPMRP := 0;
      NbLignes := 0;
      while (TypeDim = 'DIM') and (CodeArtGEN = CodeArt) and (i < TOBP.Detail.Count) do
      begin
        MoyDPA := MoyDPA + TOBP.Detail[i].GetValue('GL_DPA');
        MoyPMAP := MoyPMAP + TOBP.Detail[i].GetValue('GL_PMAP');
        MoyDPR := MoyDPR + TOBP.Detail[i].GetValue('GL_DPR');
        MoyPMRP := MoyPMRP + TOBP.Detail[i].GetValue('GL_PMRP');
        NbLignes := NbLignes + 1;
        i := i + 1;
        if i < TOBP.Detail.Count then
        begin
          CodeArt := TOBP.Detail[i].GetValue('GL_CODEARTICLE');
          TypeDim := TOBP.Detail[i].GetValue('GL_TYPEDIM');
        end;
      end;
      if NbLignes <> 0 then
      begin
        MoyDPA := Arrondi(MoyDPA / NbLignes, NbDec);
        MoyPMAP := Arrondi(MoyPMAP / NbLignes, NbDec);
        MoyDPR := Arrondi(MoyDPR / NbLignes, NbDec);
        MoyPMRP := Arrondi(MoyPMRP / NbLignes, NbDec);
      end;
      if TOBLGEN = nil then
      begin
        SetLastError(8, '');
        Break
      end;
      TOBLGEN.PutValue('GL_DPA', MoyDPA);
      TOBLGEN.PutValue('GL_PMAP', MoyPMAP);
      TOBLGEN.PutValue('GL_DPR', MoyDPR);
      TOBLGEN.PutValue('GL_PMRP', MoyPMRP);
      if EnHT then
      begin
        TOBLGEN.PutValue('GL_PUHTDEV', MoyDPA);
        CalculeLigneHT(TOBLGEN, Tob_PiedBase, TOBP, NbDec, True);
      end
      else
      begin
        TOBLGEN.PutValue('GL_PUTTCDEV', MoyDPA);
        CalculeLigneTTC(TOBLGEN, Tob_PiedBase, TOBP, NbDec, True);
      end;
    end
    else i := i + 1;
  end;
end;

function TOF_MBOGENETRANSF.CreationTRV(Depot: string): boolean;
var Q: TQuery;
  GereSurSiteDistant: Boolean;
begin
  result := False;
  GereSurSiteDistant := False;
  Q := OpenSQL('Select GDE_SURSITEDISTANT from DEPOTS WHERE GDE_DEPOT="' + Depot + '"', True);
  if not Q.EOF then GereSurSiteDistant := Q.FindField('GDE_SURSITEDISTANT').AsString = 'X';
  Ferme(Q);
  if GereSurSiteDistant and GetParamSoc('SO_GCTRV') then Result := True;
end;

procedure TOF_MBOGENETRANSF.ConstruireTRVouTRE(TOBP: TOB);
var i: integer;
  RefPieceTRE, DepotDest, NatureP: string;
  Tob_Piece: TOB; // TOB entête de pièce
begin
  DepotDest := TOBP.GetValue('GP_DEPOTDEST');
  if CreationTRV(DepotDest) then NatureP := 'TRV' else NatureP := 'TRE';
  Tob_Piece := TOB.CREATE('PIECE', TOBToutesLesPieces, -1);
  Tob_Piece.Dupliquer(TOBP, True, True);
  Tob_Piece.PutValue('GP_NATUREPIECEG', NatureP);
  RefPieceTRE := EncodeRefPiece(Tob_Piece);
  TOBP.PutValue('GP_DEVENIRPIECE', RefPieceTRE);
  for i := 0 to Tob_Piece.Detail.Count - 1 do
  begin
    Tob_Piece.Detail[i].PutValue('GL_NATUREPIECEG', NatureP);
    Tob_Piece.Detail[i].PutValue('GL_DEPOT', DepotDest);
    Tob_Piece.Detail[i].PutValue('GL_ETABLISSEMENT', DepotDest);
  end;
end;

function TOF_MBOGENETRANSF.ValideLesPieces(TOBTrfDepot: TOB): boolean;
var i, NbLigne : integer;
  TOBP: TOB;
begin
  BEGINTRANS;
  result := False;
  try
    // TOBToutesLesPieces.Detail.Count doit être égal à 2 pour les Transferts (Emis + Reçu)
    // et égal à 1 pour les commandes fournisseurs
    TOBP := nil;
    for i := 0 to TOBToutesLesPieces.Detail.Count - 1 do
    begin
      TOBP := TOBToutesLesPieces.Detail[i];
      if i = 1 then
        begin
        // MAJ de PIEDBASE
        for NbLigne := 0 to Tob_PiedBase.Detail.Count - 1 do
          Tob_PiedBase.Detail[NbLigne].PutValue('GPB_NATUREPIECEG', TOBP.getvalue('gp_naturepieceg'));
        // MAJ de PIEDECHE
        for NbLigne := 0 to Tob_PiedEche.Detail.Count - 1 do
          Tob_PiedEche.Detail[NbLigne].PutValue('GPE_NATUREPIECEG', TOBP.getvalue('gp_naturepieceg'));
        end;
      Tob_PiedBase.SetAllModifie(True);
      Tob_PiedEche.SetAllModifie(True);
      ValideLesLignes(TOBP, Tob_Article, Tob_Catalogue, Tob_Noment, nil, nil, nil, False, True);
      ValideLeTiers(TOBP, Tob_Tiers);
      Tob_PiedBase.InsertDB(nil);
      Tob_PiedEche.InsertDB(nil);
      ValideLesAdresses(TOBP, TOBP, Tob_Adresses);
      TOBP.InsertDB(nil);
    end;
    ////////////////////////////////////////////////////////////////////
    // MAJ des articles et du stock
    ////////////////////////////////////////////////////////////////////
    Affichage(7);
    ValideLesArticles(TOBP, Tob_Article);

    MarquePROPTRANSFLIG(TOBTrfDepot);
    COMMITTRANS;
    result := true;
  except
    ROLLBACK;
  end;
end;

procedure TOF_MBOGENETRANSF.MarquePROPTRANSFLIG(TOBTrfDepot: TOB);
var StUpdate, StArticle: string;
  TabIN: array of string;
  i, CountStArt, NbRequete: integer;
begin
  if BGenerationCF then
    StUpdate := 'UPDATE PROPTRANSFLIG SET GTL_GENECF="X" WHERE GTL_CODEPTRF="' + CodePtrf +
      '" AND GTL_DEPOTDEST="..." AND GTL_ARTICLE IN ('
  else
    StUpdate := 'UPDATE PROPTRANSFLIG SET GTL_GENERER="X" WHERE GTL_CODEPTRF="' + CodePtrf + '" AND GTL_DEPOTDEST="' +
      TOBTrfDepot.GetValue('GTL_DEPOTDEST') + '" AND GTL_ARTICLE IN (';
  CountStArt := 0;
  NbRequete := 0;
  SetLength(TabIN, 1);
  for i := 0 to TOBTrfDepot.Detail.Count - 1 do
  begin
    StArticle := TOBTrfDepot.Detail[i].GetValue('GTL_ARTICLE');
    if CountStArt >= NbArtParRequete then
    begin
      NbRequete := NbRequete + 1;
      CountStArt := 0;
      SetLength(TabIN, NbRequete + 1);
    end;
    if not FindStInTAB(StArticle, TabIN) then
    begin
      CountStArt := CountStArt + 1;
      if TabIN[NbRequete] = '' then TabIN[NbRequete] := '"' + StArticle + '"'
      else TabIN[NbRequete] := TabIN[NbRequete] + ',"' + StArticle + '"';
    end;
  end;
  if TabIN[0] <> '' then
  begin
    for i := Low(TabIN) to High(TabIN) do
    begin
      if TabIN[i] <> '' then ExecuteSQL(StUpdate + TabIN[i] + ')');
    end;
  end;
end;

procedure TOF_MBOGENETRANSF.ImpressionDesEtats(BApercu, BDuplicata: boolean; StModele: string);
var stCle, stSql, StSQLUpdate, stWhere: string;
begin
  stWhere := '';
  with TOBToutesLesPieces.Detail[0] do
  begin
    stCle := NaturePiece + ';' + DateToStr(GetValue('GP_DATEPIECE')) + ';' +
      GetValue('GP_SOUCHE') + ';' + IntToStr(GetValue('GP_NUMERO')) + ';' +
      IntToStr(GetValue('GP_INDICEG'));
  end;
  StringToCleDoc(stCle, CleDoc);
  stWhere := WherePiece(CleDoc, ttdPiece, False);
  stSql := RecupSQLModele('E', 'GPJ', stModele, '', '', '', ' WHERE ' + stWhere);
  stSql := stSql + ' and GL_NONIMPRIMABLE<>"X"';
  if (pos('ORDER BY', uppercase(stSql)) <= 0) then stSql := stSql + ' order by GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE,GL_ARTICLE';
  if (not BApercu) then V_PGI.NoPrintDialog := True else V_PGI.NoPrintDialog := False;
  LanceEtat('E', 'GPJ', stModele, BApercu, false, false, nil, trim(stSql), '', BDuplicata);
  //Requête de MAJ de GP_EDITEE pour la pièce éditée
  StSQLUpdate := 'Update PIECE set GP_EDITEE="X" where ' + stWhere;
  ExecuteSQL(StSQLUpdate);
end;

procedure TOF_MBOGENETRANSF.AfterGeneration;
var StSQL: string;
begin
  //Suppression de la proposition après génération
  if TCheckbox(GetControl('SUPPROP')).checked then
  begin
    StSQL := 'Delete from PROPTRANSFENT where GTE_CODEPTRF="' + CodePtrf + '"';
    ExecuteSQL(StSQL);
    StSQL := 'Delete from PROPTRANSFLIG where GTL_CODEPTRF="' + CodePtrf + '"';
    ExecuteSQL(StSQL);
  end
  else
  begin
    if BGenerationCF then StSQL := 'Update PROPTRANSFENT set GTE_ENCOURSPTRF="-",GTE_COMMANDEFOU="X",GTE_ENCOURSCF="-" where GTE_CODEPTRF="' + CodePtrf + '"'
    else StSQL := 'Update PROPTRANSFENT set GTE_ENCOURSPTRF="-",GTE_GENEREPTRF="X",GTE_ENCOURSGENE="-" where GTE_CODEPTRF="' + CodePtrf + '"';
    ExecuteSQL(StSQL);
  end;
end;

procedure TOF_MBOGENETRANSF.EditionClick;
begin
  TCheckbox(GetControl('APERCU')).Enabled := TCheckbox(GetControl('EDITAUTOTRANSF')).checked;
end;

procedure TOF_MBOGENETRANSF.ChoixTransfert(TypeGeneration: string);
var Etat: boolean;
begin
  Etat := TypeGeneration = 'PGI';
  TRadioButton(GetControl('TRANSFERTPGI')).checked := Etat;
  TRadioButton(GetControl('TRANSFERTASCII')).checked := not Etat;
  SetControlVisible('EDITAUTOTRANSF', Etat);
  SetControlVisible('APERCU', Etat);
  SetControlVisible('GAS_CODEEX', not Etat);
  SetControlVisible('TGAS_CODEEX', not Etat);
  SetControlVisible('EMPLACEMENT', not Etat);
  SetControlVisible('TEMPLACEMENT', not Etat);
  SetControlVisible('EXPORTFTP', not Etat);
  SetControlVisible('PROPOSITION_POSITIVE', not Etat);
  if Etat then
  begin
    EditionClick;
    SetControlVisible('BPARAMFTP', False);
  end
  else
    if TCheckBox(GetControl('EXPORTFTP')).checked then SetControlVisible('BPARAMFTP', True);
  TransfertPGI := Etat;
end;

procedure TOF_MBOGENETRANSF.ChargeCheminExport;
var Q: TQuery;
  St_Nature: string;
begin
  if BGenerationCF then St_Nature := 'GCF' else St_Nature := 'MPT';
  Q := OpenSQL('SELECT GAS_FICHIEREXPORT FROM EXPORTASCII_ENTETE WHERE GAS_NATURE="' +
    St_Nature + '" AND GAS_CODEEX="' + GetControlText('GAS_CODEEX') + '"', True);
  if not Q.EOF then SetControlText('EMPLACEMENT', Q.FindField('GAS_FICHIEREXPORT').AsString);
  Ferme(Q);
end;

procedure TOF_MBOGENETRANSF.ParametrageFTP;
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

procedure TOF_MBOGENETRANSF.ExportDuFichierParFTP;
var xpFtp: TOXFTP;
  PortFTP: integer;
  LocalFileName, HostFileName: string;
begin
  if (StGCORLIHOST = '') or (StGCORLIUSERNAME = '') or (StGCORLIPASSWORD = '') then exit;
  PortFTP := 21;
  xpFtp := TOXFTP.Create(PortFtp);
  if xpFTP.Connect(StGCORLIHOST, StGCORLIUSERNAME, StGCORLIPASSWORD) then
  begin
    LocalFileName := GetControlText('EMPLACEMENT');
    HostFileName := StGCORLIREPFTP + ExtractFileName(LocalFileName);
    xpFTP.PutFileAscii(LocalFileName, HostFileName);
    xpFTP.DisConnect;
  end;
  xpFTP.Free;
end;

///////////////////////////////////////////////////////////////////////////////

procedure AGLEditionClick(Parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then TOTOF := TFVierge(F).LaTOF else exit;
  if (TOTOF is TOF_MBOGENETRANSF) then TOF_MBOGENETRANSF(TOTOF).EditionClick;
end;

procedure AGLChoixTransfert(Parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then TOTOF := TFVierge(F).LaTOF else exit;
  if (TOTOF is TOF_MBOGENETRANSF) then TOF_MBOGENETRANSF(TOTOF).ChoixTransfert(string(Parms[1]));
end;

procedure AGLChargeCheminExport(Parms: array of Variant; Nb: Integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then TOTOF := TFVierge(F).LaTOF else exit;
  if (TOTOF is TOF_MBOGENETRANSF) then TOF_MBOGENETRANSF(TOTOF).ChargeCheminExport;
end;

procedure AGLParametrageFTP(Parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then TOTOF := TFVierge(F).LaTOF else exit;
  if (TOTOF is TOF_MBOGENETRANSF) then TOF_MBOGENETRANSF(TOTOF).ParametrageFTP;
end;

initialization
  registerclasses([TOF_MBOGENETRANSF]);
  RegisterAglProc('EditionClick', True, 0, AGLEditionClick);
  RegisterAglProc('ChoixTransfert', True, 1, AGLChoixTransfert);
  RegisterAGLProc('ChargeCheminExport', True, 0, AGLChargeCheminExport);
  RegisterAglProc('ParametrageFTP', True, 0, AGLParametrageFTP);
end.

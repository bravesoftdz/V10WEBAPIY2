{***********UNITE*************************************************
Auteur  ...... : Julien D
Créé le ...... : 19/03/2002
Modifié le ... : 20/03/2002
Description .. : Source TOF de la FICHE : MBOEXPORTPIECE ()
Mots clefs ... : TOF;MBOEXPORTPIECE
*****************************************************************}
unit MBORECALCULPIECE_TOF;

interface

uses StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  eMul, MaineAGL,
  {$ELSE}
  Fiche, HDB, mul, DBGrids, db, dbTables, Fe_Main,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF,
  HStatus, HQry, UTob, EXPORTASCII_ENTETE_TOM, FileCtrl, Paramsoc, M3FP, FactUtil,
  SaisUtil, FactCalc, FactComm, EntGC,
  FactTOB;

type FichierFournisseur = record
    Fournisseur: string;
    FichierASCII: TextFile;
  end;
type
  TOF_MBORECALCULPIECE = class(TOF)
  private
    stWhere: string;
    IdTab: integer;
    CodeTiers: string;
    CodeTiersOld: string;
    Tob_Tiers: TOB;
    procedure BatchRecalcul;
    procedure SetRecalcul;
    procedure SetAllRecalcul;
    procedure RecalculPiece(var Tob_Piece: TOB);
    procedure ControleLigneGenerique(var Tob_Piece: TOB);
  public
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  end;

implementation

procedure TOF_MBORECALCULPIECE.BatchRecalcul;
var i: integer;
begin
  //
  // Libération de la TOB Tiers
  //
  Tob_Tiers := nil;
  CodeTiersOld := '';
  CodeTiers := '';

  IdTab := -1;
  with TFMul(Ecran) do
  begin
    if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
    begin
      PGIInfo('Veuillez sélectionner les pièces à recalculer', Caption);
      exit;
    end;

    SourisSablier;
    try
      if FListe.AllSelected then
      begin
        if PGIAsk('Voulez-vous recalculer toutes les pièces ?', Caption) <> mrYes then exit;

        if Transactions(SetallRecalcul, 2) <> oeOK then PGIError('Impossible de recalculer toutes les pièces', Caption);

        FListe.AllSelected := false;
      end else
      begin
        if PGIAsk('Voulez-vous recalculer les pièces sélectionnées ?', Caption) <> mrYes then exit;

        InitMove(FListe.NbSelected, '');

        for i := 0 to FListe.NbSelected - 1 do
        begin
          FListe.GotoLeBookMark(i);
          {$IFDEF EAGLCLIENT}
          Q.TQ.Seek(FListe.Row - 1);
          {$ENDIF}
          if Transactions(SetRecalcul, 2) <> oeOK then PGIBox('Impossible de recalculer la pièce n° ' + Q.FindField('GP_NUMERO').AsString, Caption);
          MoveCur(False);
        end;
        FListe.ClearSelected;
        FiniMove;
      end;
      ChercheClick;
    finally
      bSelectAll.Down := False;
      SourisNormale;
    end;
  end;
  //
  // Libération de la TOB Tiers
  //
  if Tob_Tiers <> nil then Tob_Tiers.free;
  Tob_Tiers := nil;
end;

// Procédure appelée par une transaction pour exporter les lignes de la pièce

procedure TOF_MBORECALCULPIECE.SetRecalcul;
var Nature, Souche, Nb, Indice: string;
  StWherePiece, StSQL: string;
  Q1: TQuery;
  Tob_Piece: TOB;
begin

  with TFMul(Ecran) do
  begin
    Nature := Q.FindField('GP_NATUREPIECEG').AsString;
    Souche := Q.FindField('GP_SOUCHE').AsString;
    Nb := Q.FindField('GP_NUMERO').AsString;
    Indice := Q.FindField('GP_INDICEG').AsString;
  end;

  StWherePiece := 'WHERE GP_NATUREPIECEG="' + Nature + '" AND ' +
    'GP_SOUCHE="' + Souche + '" AND GP_NUMERO=' + Nb + ' AND GP_INDICEG=' + Indice;

  StSQL := 'SELECT * FROM PIECE ' + StWherePiece;

  Q1 := OpenSQL(StSQL, True);
  if not Q1.Eof then
  begin
    //
    // Chargement de la pièce
    //
    Tob_Piece := TOB.Create('PIECE', nil, -1);
    Tob_Piece.SelectDB('', Q1);
    AddLesSupEntete(Tob_Piece);
    //
    // Recalcul
    //
    RecalculPiece(Tob_Piece);
    //
    // Libération de la pièce
    //
    Tob_Piece.free;
    Tob_Piece := nil;
  end;
  Ferme(Q1);
end;

// Procédure appelée par une transactions pour recalculer toutes les pièces

procedure TOF_MBORECALCULPIECE.SetAllRecalcul;
var StSQL, StSQL2: string;
  Q1: TQuery;
  Q2: TQuery;
  Tob_Piece: TOB;
  Nature, Souche, Nb, Indice, StWherePiece: string;
begin
  //
  // Chargement de toutes les entêtes de pièce sélectionnées
  //
  StSQL := 'SELECT GP_NATUREPIECEG, GP_SOUCHE, GP_NUMERO, GP_INDICEG FROM PIECE ' + StWhere;

  Q1 := OpenSQL(StSQL, True);
  if not Q1.Eof then
  begin
    Q1.First;

    while not Q1.Eof do
    begin
      //
      // Chargement de la pièce
      //
      Nature := Q1.FindField('GP_NATUREPIECEG').AsString;
      Souche := Q1.FindField('GP_SOUCHE').AsString;
      Nb := Q1.FindField('GP_NUMERO').AsString;
      Indice := Q1.FindField('GP_INDICEG').AsString;

      StWherePiece := 'WHERE GP_NATUREPIECEG="' + Nature + '" AND ' +
        'GP_SOUCHE="' + Souche + '" AND GP_NUMERO=' + Nb + ' AND GP_INDICEG=' + Indice;

      StSQL2 := 'SELECT * FROM PIECE ' + StWherePiece;
      Q2 := OpenSQL(StSQL2, True);

      if not Q2.Eof then
      begin
        //
        // Chargement de la pièce
        //
        Tob_Piece := TOB.Create('PIECE', nil, -1);
        Tob_Piece.SelectDB('', Q2);
        AddLesSupEntete(Tob_Piece);

        RecalculPiece(Tob_Piece);

        //
        // Libération de la pièce
        //
        Tob_Piece.free;
        Tob_Piece := nil;
      end;
      Ferme(Q2);

      //
      // Pièce suivante
      //
      Q1.Next;
    end;
  end;
  Ferme(Q1);
end;

procedure TOF_MBORECALCULPIECE.ControleLigneGenerique(var Tob_Piece: TOB);
var cpt: integer;
  CodeArticle: string;
  ArticleTraite: string;
  TypeLigne: string;
  Qte: double;
  TotQte: double;
  PrixHT: double;
  PrixHTNet: double;
  TotPrixHT: double;
  TotPrixHTNet: double;
  PrixTTC: double;
  PrixTTCNet: double;
  TotPrixTTC: double;
  TotPrixTTCNet: double;
  Tob_Ligne: TOB;
  First: boolean;
begin
  First := True;

  for Cpt := Tob_Piece.Detail.Count - 1 downto 0 do
  begin
    Tob_Ligne := Tob_Piece.Detail[Cpt];

    if first then
    begin
      first := False;
      ArticleTraite := Tob_Ligne.GetValue('GL_CODEARTICLE');
      Qte := 0.0;
      TotQte := 0.0;
      TotPrixHT := 0.0;
      TotPrixHTNet := 0.0;
      PrixHT := 0.0;
      PrixHTNet := 0.0;
      TotPrixTTC := 0.0;
      TotPrixTTCNet := 0.0;
      PrixTTC := 0.0;
      PrixTTCNet := 0.0;
    end;

    CodeArticle := Tob_Ligne.GetValue('GL_CODEARTICLE');
    TypeLigne := Tob_Ligne.GetValue('GL_TYPEDIM');

    if TypeLigne = 'DIM' then
    begin
      if ArticleTraite = '' then
      begin
        ArticleTraite := Tob_Ligne.GetValue('GL_CODEARTICLE');
        Qte := 0.0;
        TotQte := 0.0;
        TotPrixHT := 0.0;
        TotPrixHTNet := 0.0;
        PrixHT := 0.0;
        PrixHTNet := 0.0;
        TotPrixTTC := 0.0;
        TotPrixTTCNet := 0.0;
        PrixTTC := 0.0;
        PrixTTCNet := 0.0;
      end;

      if CodeArticle = ArticleTraite then
      begin
        Qte := Tob_Ligne.GetValue('GL_QTEFACT');

        ////////////////////////////////////////////////////////////////
        // Recherche du prix de vente
        ////////////////////////////////////////////////////////////////
        PrixHT := Tob_Ligne.GetValue('GL_PUHTDEV');
        PrixHTNet := Tob_Ligne.GetValue('GL_PUHTNETDEV');
        PrixTTC := Tob_Ligne.GetValue('GL_PUTTCDEV');
        PrixTTCNet := Tob_Ligne.GetValue('GL_PUTTCNETDEV');
        TotQte := TotQte + Qte;
        TotPrixHT := TotPrixHT + (Qte * PrixHT);
        TotPrixHTNet := TotPrixHTNet + (Qte * PrixHTNet);
        TotPrixTTC := TotPrixTTC + (Qte * PrixTTC);
        TotPrixTTCNet := TotPrixTTCNet + (Qte * PrixTTCNet);
      end else
      begin
        ArticleTraite := Tob_Ligne.GetValue('GL_CODEARTICLE');
        Qte := 0.0;
        TotQte := 0.0;
        TotPrixHT := 0.0;
        TotPrixHTNet := 0.0;
        PrixHT := 0.0;
        PrixHTNet := 0.0;
        TotPrixTTC := 0.0;
        TotPrixTTCNet := 0.0;
        PrixTTC := 0.0;
        PrixTTCNet := 0.0;
      end;
    end else
    begin
      if TypeLigne = 'GEN' then
      begin
        CodeArticle := Tob_Ligne.GetValue('GL_CODESDIM');
        if CodeArticle = ArticleTraite then
        begin
          ////////////////////////////////////////////
          // Quantité
          ///////////////////////////////////////////
          Tob_Ligne.PutValue('GL_QTESTOCK', TotQte); // Total des quantités
          Tob_Ligne.PutValue('GL_QTEFACT', TotQte);
          ////////////////////////////////////////////
          // Prix unitaire
          ///////////////////////////////////////////
          if (TotQte <> 0) then
          begin
            PrixHT := Arrondi(TotPrixHT / TotQte, 2);
            PrixHTNet := Arrondi(TotPrixHTNet / TotQte, 2);
            PrixTTC := Arrondi(TotPrixTTC / TotQte, 2);
            PrixTTCNet := Arrondi(TotPrixTTCNet / TotQte, 2);
          end else
          begin
            PrixHT := 0.0;
            PrixHTNet := 0.0;
            PrixTTC := 0.0;
            PrixTTCNet := 0.0;
          end;

          Tob_Ligne.PutValue('GL_PUHTDEV', PrixHT);
          Tob_Ligne.PutValue('GL_PUHTNETDEV', PrixHTNet);
          Tob_Ligne.PutValue('GL_PUHT', PrixHT);
          Tob_Ligne.PutValue('GL_PUHTNET', PrixHTNet);
          Tob_Ligne.PutValue('GL_MONTANTHTDEV', TotPrixHT);
          Tob_Ligne.PutValue('GL_MONTANTHT', TotPrixHTNet);
          Tob_Ligne.PutValue('GL_PUTTCDEV', PrixTTC);
          Tob_Ligne.PutValue('GL_PUTTCNETDEV', PrixTTCNet);
          Tob_Ligne.PutValue('GL_PUTTC', PrixTTC);
          Tob_Ligne.PutValue('GL_PUTTCNET', PrixTTCNet);
          Tob_Ligne.PutValue('GL_MONTANTTTCDEV', TotPrixTTC);
          Tob_Ligne.PutValue('GL_MONTANTTTC', TotPrixTTCNet);
          //
          // Ré-initialisation des variables pour la prochaine ligne
          //
          ArticleTraite := '';
          TotPrixHT := 0.0;
          TotPrixHTNet := 0.0;
          PrixHT := 0.0;
          PrixHTNet := 0.0;
          TotPrixTTC := 0.0;
          TotPrixTTCNet := 0.0;
          PrixTTC := 0.0;
          PrixTTCNet := 0.0;
        end;
      end else
      begin
        ArticleTraite := '';
        TotPrixHT := 0.0;
        TotPrixHTNet := 0.0;
        PrixHT := 0.0;
        PrixHTNet := 0.0;
        TotPrixTTC := 0.0;
        TotPrixTTCNet := 0.0;
        PrixTTC := 0.0;
        PrixTTCNet := 0.0;
      end;
    end;
  end;
end;

procedure TOF_MBORECALCULPIECE.RecalculPiece(var Tob_Piece: TOB);
var Tob_Article: TOB;
  Tob_Ligne: TOB;
  TobFilleArticle: TOB;
  Tob_Piedbase: TOB;
  TOB_Ports:TOB;
  Nature, Souche: string;
  Numero, indice: integer;
  Condition: string;
  SQL: string;
  CodeArt: string;
  Q2: Tquery;
  cpt: integer;
  DEV: RDEVISE;
  CleDoc: R_CleDoc;
begin
  CodeArt := '';

  //
  // Création et chargement TOB TIERS
  //
  CodeTiers := Tob_Piece.GetValue('GP_TIERS');

  if CodeTiers <> CodeTiersOld then
  begin
    if Tob_Tiers <> nil then
    begin
      Tob_Tiers.free;
      Tob_Tiers := nil;
    end;

    SQL := 'Select * From TIERS WHERE T_TIERS="' + CodeTiers + '"';
    Q2 := OpenSQL(SQL, True);
    if not Q2.EOF then
    begin
      Tob_Tiers := TOB.Create('TIERS', nil, -1);
      Tob_Tiers.SelectDB('', Q2);
    end;
    Ferme(Q2);
    //
    // Sauvegarde du code tiers
    //
    CodeTiersOld := CodeTiers;
  end;

  /////////////////////////////////////////////////////////////////////////
  // Chargement de la TOB des lignes du document
  // (qui doivent être fille de l'entête)
  /////////////////////////////////////////////////////////////////////////
  Nature := TOB_Piece.GetValue('GP_NATUREPIECEG');
  Souche := TOB_Piece.GetValue('GP_SOUCHE');
  Numero := TOB_Piece.GetValue('GP_NUMERO');
  Indice := TOB_Piece.GetValue('GP_INDICEG');

  Condition := 'GL_NATUREPIECEG="' + Nature + '" AND GL_SOUCHE="' + Souche + '" AND GL_NUMERO="' + IntToStr(Numero) + '" AND GL_INDICEG="' + IntToStr(Indice) +
    '" ORDER BY GL_NUMLIGNE';
  SQL := 'SELECT * From LIGNE WHERE' + ' ' + Condition;
  Q2 := OpenSQL(SQL, True);
  if not Q2.EOF then TOB_Piece.LoadDetailDB('LIGNE', '', '', Q2, True, True);
  Ferme(Q2);
  //
  // Ajout des champs supplémentaires
  //
  for Cpt := 0 to Tob_Piece.Detail.Count - 1 do
  begin
    Tob_Ligne := Tob_Piece.Detail[Cpt];
    AddLesSupLigne(Tob_Ligne, False);
  end;

  //
  //  CodeArt   := Tob_Ligne.GetValue ('GL_ARTICLE');
  //  if CodeArt <> '' then
  //  begin
  //    SQL:='Select * From ARTICLE WHERE GA_ARTICLE="'+CodeArt+'"';
  //    Q2:=OpenSQL(SQL,True) ;
  //    if Not Q2.EOF then
  //    begin
  //      TobFilleArticle := CreerTOBArt(Tob_Article);
  //      TobFilleArticle.SelectDB('',Q2);
  //      //LoadTOBDispo (TobFilleArticle, True, CreerQuelDepot(TobEntete)) ;
  //    end;
  //    Ferme (Q2);
  //  end;
  //end;

  /////////////////////////////////////////////////////////////////////////
  // Chargement de la TOB des articles
  /////////////////////////////////////////////////////////////////////////
  CleDoc := TOB2CleDoc(TOB_Piece);
  Tob_Article := TOB.CREATE('Les Articles', nil, -1);

  Q2 := OpenSQL('SELECT ARTICLE.* FROM ARTICLE, LIGNE WHERE ' + WherePiece(CleDoc, ttdLigne, False) + ' AND GL_ARTICLE=GA_ARTICLE', True);
  Tob_Article.LoadDetailDB('ARTICLE', '', '', Q2, False, True);
  Ferme(Q2);

  /////////////////////////////////////////////////////////////////////////
  // Chargement de Piedport
  /////////////////////////////////////////////////////////////////////////
  TOB_Ports := TOB.CREATE('Les Ports', nil, -1);
  Q2 := OpenSQL('SELECT * FROM PIEDPORT WHERE ' + WherePiece(CleDoc, ttdPorc, False), True);
  if not Q2.EOF then TOB_Ports.LoadDetailDB('PIEDPORT', '', '', Q2, True, False);
  Ferme(Q2);

  /////////////////////////////////////////////////////////////////////////
  // Chargement de Piedbase
  /////////////////////////////////////////////////////////////////////////
  Tob_PiedBase := TOB.CREATE('Les PiesBases', nil, -1);

  Condition := 'GPB_NATUREPIECEG="' + Nature + '" AND GPB_SOUCHE="' + Souche + '" AND GPB_NUMERO="' + IntToStr(Numero) + '" AND GPB_INDICEG="' +
    IntToStr(Indice) + '"';
  SQL := 'SELECT * From PIEDBASE WHERE' + ' ' + Condition;
  Q2 := OpenSQL(SQL, True);
  if not Q2.EOF then
  begin
    Tob_PiedBase.LoadDetailDB('PIEDBASE', '', '', Q2, True, True);
  end;
  Ferme(Q2);

  //Recalcul de la pièce
  CleDoc := TOB2CleDoc(Tob_Piece);
  DEV.Code := Tob_Piece.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := GetTaux(DEV.Code, DEV.DateTaux, CleDoc.DatePiece);

  PutValueDetail(Tob_Piece, 'GP_RECALCULER', 'X');
  CalculFacture(Tob_Piece, Tob_PiedBase, Tob_Tiers, Tob_Article, TOB_Ports, nil, nil, DEV);
  //
  // Contrôle des lignes génériques
  //
  ControleLigneGenerique(Tob_Piece);
  //
  // MAJ
  //
  Tob_Piece.SetAllModifie(True);
  Tob_PiedBase.SetAllModifie(True);
  Tob_Piece.UpdateDB(False);
  Tob_PiedBase.UpdateDB(False);

  if Tob_Article  <> nil then Tob_Article.free;
  if Tob_PiedBase <> nil then Tob_PiedBase.free;
  if TOB_Ports    <> nil then TOB_Ports.free;
end;

procedure TOF_MBORECALCULPIECE.OnLoad;
begin
  inherited;
  stWhere := RecupWhereCritere(TPageControl(TFMul(Ecran).Pages));
end;

procedure TOF_MBORECALCULPIECE.OnArgument(S: string);
begin
  inherited;
  //THValComboBox(GetControl('GP_ETATEXPORT')).Value := 'ATT';
  //SetControlText('EMPLACEMENT',GetParamSoc('SO_GENERATIONFICHIERASCII'));
end;

procedure TOF_MBORECALCULPIECE.OnClose;
begin
  inherited;
end;

// procedure appellée par le bouton BOuvrir

procedure AGLBatchRecalcul(Parms: array of Variant; Nb: Integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFmul) then TOTOF := TFMul(F).LaTOF else exit;
  if (TOTOF is TOF_MBORECALCULPIECE) then TOF_MBORECALCULPIECE(TOTOF).BatchRecalcul;
end;

initialization
  registerclasses([TOF_MBORECALCULPIECE]);
  RegisterAGLProc('BatchRecalcul', True, 0, AGLBatchRecalcul);
end.

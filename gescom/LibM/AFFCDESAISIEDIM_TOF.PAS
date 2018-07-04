{***********UNITE*************************************************
Auteur  ...... : A. CATHELINEAU
Créé le ...... : 04/12/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFFCDESAISIEDIM ()
Mots clefs ... : TOF;AFFCDESAISIEDIM
*****************************************************************}
unit AFFCDESAISIEDIM_TOF;

interface

uses StdCtrls, Controls, Classes, Windows,
  {$IFNDEF EAGLCLIENT}
  db, dbtables,
  {$ENDIF}
  vierge, forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, UDimArticle, UTob,
  UtilArticle, HDimension, UtilDimArticle, EntGc, M3FP, AGLInit;

type
  TOF_AFFCDESAISIEDIM = class(TOF)
    DimensionsArticle: TODimArticle;

  private
    dimAction: string; // SELECT, SAISIE, CONSULT, MULTI
    dimChamp: string;
    dimMasque: string;
    CodeArticle: string;
    NatureDoc, NaturePiece: string;
    Depot: string;
    NotClose, PrixUnique: Boolean;
    MAJ: Boolean;

    procedure InitDimensionsArticle;

    // Gestion document
    procedure PreparationSaisie;
    procedure InitTobSelect(var TobSelect: TOB);
    procedure SetReadOnly;
    procedure OnSaisiePiece(Validation: boolean);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    function VerifStock(ItemDim: THDimensionItem): Boolean;

  public
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure AdapterFiche;
    procedure OnChangeItem(Sender: TObject);
  end;

procedure AGLOnClickParamGrilleDoc2(parms: array of variant; nb: integer);

implementation

procedure TOF_AFFCDESAISIEDIM.OnUpdate;
var ItemDim: THDimensionItem;
begin
  ItemDim := DimensionsArticle.Dim.CurrentItem;
  if (NatureDoc = NAT_AFFCDEAFF) and (VerifStock(ItemDim)) then
  begin
    NotClose := True;
    exit;
  end;
  //if DimensionsArticle.NewDimChamp[1]='' then OnDoubleClick(ItemDim)
  OnSaisiePiece(True);
  inherited;
end;

procedure TOF_AFFCDESAISIEDIM.OnLoad;
begin
  inherited;
  if DimensionsArticle.TOBArticleDim = nil then
  begin
    SetControlVisible('LPasDimensions', True);
    SetControlEnabled('BPARAMDIM', False);
    Ecran.Caption := 'Reservation des commandes clients de l''article: ' + LibelleArticleGenerique(CodeArticle);
    Exit;
  end;
  MAJ := False;
  if NatureDoc = NAT_AFFCDERES then
    Ecran.Caption := 'Reservation de l''article: ' + LibelleArticleGenerique(CodeArticle)
  else Ecran.Caption := 'Affectation de l''article: ' + LibelleArticleGenerique(CodeArticle);
  AdapterFiche;
end;

procedure TOF_AFFCDESAISIEDIM.AdapterFiche;
var CoordDernCol, CoordDernLign: TRect;
  GridDim: THGrid;
  ValOnglet: Integer;
begin
  Ecran.Width := 1000; // pour pouvoir voir toutes les colonnes
  Ecran.Height := 1000; // pour pouvoir voir toutes les lignes
  THDimension(GetControl('FDIM')).Align := alNone;
  GridDim := DimensionsArticle.Dim.GridDim;
  if DimensionsArticle.Dim.DimOngl = nil then ValOnglet := 0 else ValOnglet := GridDim.DefaultRowHeight;
  CoordDernCol := GridDim.CellRect(GridDim.VisibleColCount + GridDim.FixedCols, 1);
  CoordDernLign := GridDim.CellRect(2, GridDim.VisibleRowCount + GridDim.FixedRows); // Dernière ligne
  GridDim.ScrollBars := ssNone;
  Ecran.Height := CoordDernLign.Bottom + 76 + ValOnglet; // 76= titre et panel bouton de ecran + encadrment objet Dim
  if CoordDernCol.Right < 230 then Ecran.Width := 230
  else Ecran.Width := CoordDernCol.Right + 20;
  // Ajustement taille si ajout d'un ascenseur
  if (GridDim.ScrollBars = ssVertical) or (GridDim.ScrollBars = ssBoth) then
    Ecran.Width := CoordDernCol.Right + 16 + 20; //ScrollBars+encadrement Dim
  if (GridDim.ScrollBars = ssHorizontal) or (GridDim.ScrollBars = ssBoth) then
  else Ecran.Height := CoordDernLign.Bottom + ValOnglet + 16 + 76; //ScrollBars+titre,panel bouton de ecran et encadrment objet Dim
  THDimension(GetControl('FDIM')).Align := alClient;
  if Ecran.Width < 250 then Ecran.Width := 250;
  V_PGI.FormCenter := True;
end;

procedure TOF_AFFCDESAISIEDIM.OnArgument(S: string);
var Lequel, CodeUnique: string;
  Critere: string;
  ChampMul, ValMul: string;
  x: integer;
  SQL: string;
  Q: TQuery;
begin
  inherited;
  dimAction := 'SAISIE';
  dimChamp := '';
  repeat
    Critere := uppercase(Trim(ReadTokenSt(S)));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));

        if ChampMul = 'GA_CODEARTICLE' then Lequel := ValMul;
        if ChampMul = 'ACTION' then dimAction := ValMul;
        if ChampMul = 'CHAMP' then dimChamp := ValMul;
        if ChampMul = 'DEPOT' then Depot := ValMul;
        if ChampMul = 'TYPEPARAM' then NatureDoc := ValMul;
      end;
    end;
  until Critere = '';
  NaturePiece := NatureDoc;
  CodeUnique := CodeArticleUnique(Lequel, '', '', '', '', '');

  SQL := 'SELECT GA_ARTICLE, GA_CODEARTICLE, GA_DIMMASQUE, GA_STATUTART, GA_PRIXUNIQUE,GA_REMISELIGNE,GA_TARIFARTICLE ';
  SQL := SQL + ' from Article where GA_CODEARTICLE="' + Lequel + '" AND GA_STATUTART="GEN"';
  Q := OpenSQl(SQL, True);
  //SetControlText('GA_ARTICLE', CodeUnique);
  CodeArticle := Q.FindField('GA_CODEARTICLE').AsString;
  PrixUnique := Boolean(Q.FindField('GA_PRIXUNIQUE').AsString = 'X');
  SetControlText('GA_CODEARTICLE', Lequel);
  dimMasque := Q.FindField('GA_DIMMASQUE').AsString;
  SetControlText('GA_STATUTART', Q.FindField('GA_STATUTART').AsString);
  TheTob := LaTob;
  try
    InitDimensionsArticle;
  finally
    Ferme(Q);
  end;
  Ecran.OnKeyDown := FormKeyDown;
end;

procedure TOF_AFFCDESAISIEDIM.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  //if Key=vk_valide then BEGIN Key:=0 ; if (Ecran<>Nil) and (Ecran is TFVierge) then TFVierge(Ecran).BValider.Click ;  END ;
end;

procedure TOF_AFFCDESAISIEDIM.InitDimensionsArticle;
var detail: string;
  item: THDimensionItem;
begin
  if DimensionsArticle <> nil then
  begin
    DimensionsArticle.free;
  end;
  DimensionsArticle := TODimArticle.Create(THDimension(GetControl('FDIM'))
    , CodeArticle
    , dimMasque
    , dimChamp, 'GCDIMCHAMP', NatureDoc, NaturePiece, Depot, '-', PrixUnique);
  if DimensionsArticle.TOBArticleDim = nil then exit;
  DimensionsArticle.Dim.PopUp.Items[2].Visible := False; // Menu Existant invisible
  DimensionsArticle.Dim.PopUp.Items[3].Visible := False; // Menu Inexistant invisible
  DimensionsArticle.Dim.OnChange := OnChangeItem;
  //DimensionsArticle.Dim.OnDblClick:=OnDoubleClick ;
  if DimAction = 'CONSULT' then DimensionsArticle.Action := taConsult;

  PreparationSaisie();
  detail := '-';
  AfficheUserPref(DimensionsArticle, NatureDoc, NaturePiece);
  SetControlVisible('BPARAMDIM', NaturePieceGeree(naturePiece));
  //Consultation(False) ;
  // Positionnement sur 1ère cellule en saisie
  //item := FindItemToFocused(DimensionsArticle.Dim.GridDim, DimensionsArticle.MasqueDim, VH_GC.BOTypeMasque_Defaut);
  //if item <> nil then itemDimSetFocus(DimensionsArticle, item, '');
  FocusedFirstItem(THDimension(GetControl('FDIM')));
  SetReadOnly;
end;

procedure TOF_AFFCDESAISIEDIM.PreparationSaisie;
var ItemDim: THDimensionItem;
  TobSelect, TobDoc: TOB;
  i: integer;
  TableDim: THDimensionItemList;
begin
  TableDim := DimensionsArticle.TableDim;
  if LaTob <> nil then
  begin
    for i := 0 to TableDim.count - 1 do
    begin
      if TableDim.Items[i] <> nil then
      begin
        ItemDim := THDimensionItem(TableDim.Items[i]);
        TobSelect := TOB(ItemDim.data);
        TobDoc := LaTob.findfirst(['GA_ARTICLE'], [TobSelect.GetValue('GA_ARTICLE')], false);
        if TobDoc <> nil then
        begin
          InitTobSelect(TobSelect);
          TobSelect.PutValue('GEL_QTERESERVEE', TobDoc.GetValue('GEL_QTERESERVEE'));
          TobSelect.PutValue('GEL_MODIFIABLE', TobDoc.GetValue('GEL_MODIFIABLE'));
          TobSelect.PutValue('GED_QTEPREVI', TobDoc.GetValue('GED_QTEPREVI'));
          if NatureDoc = NAT_AFFCDERES then
          begin
            TobSelect.PutValue('GEL_QTEALIVRER', TobDoc.GetValue('GEL_QTEALIVRER'));
            TobSelect.PutValue('GED_DISPORESERVEE', TobDoc.GetValue('GED_DISPORESERVEE'));
          end;
          if NatureDoc = NAT_AFFCDEAFF then
          begin
            TobSelect.PutValue('GEL_QTEAFFECTEE', TobDoc.GetValue('GEL_QTEAFFECTEE'));
            TobSelect.PutValue('GED_DISPOAFFECTEE', TobDoc.GetValue('GED_DISPOAFFECTEE'));
          end;
        end;
      end;
    end;
  end;
end;

// Pour initialiser les valeur null à zero

procedure TOF_AFFCDESAISIEDIM.InitTobSelect(var TobSelect: TOB);
var iChamp: Integer;
  ValeurChamp: Variant;
begin
  for iChamp := 1000 to (1000 + (TobSelect.ChampsSup.count - 1)) do
  begin
    ValeurChamp := TobSelect.GetValeur(iChamp);
    if VarType(ValeurChamp) = varNull then
    begin
      TobSelect.PutValue(TobSelect.GetNomChamp(iChamp), 0);
    end;
  end;
  for iChamp := 1 to MaxDimChamp do
  begin
    ValeurChamp := TobSelect.GetValeur(iChamp);
    if VarType(ValeurChamp) = varNull then
    begin
      TobSelect.PutValue(TobSelect.GetNomChamp(iChamp), 0);
    end;
  end;
end;

procedure TOF_AFFCDESAISIEDIM.SetReadOnly;
var iChamp: integer;
begin
  begin
    for iChamp := 1 to MaxDimChamp do
    begin
      if DimensionsArticle.NewDimChamp[iChamp] = 'GED_QTEPREVI' then DimensionsArticle.Dim.ReadOnly[iChamp] := True;
      if DimensionsArticle.NewDimChamp[iChamp] = 'GED_DISPORESERVEE' then DimensionsArticle.Dim.ReadOnly[iChamp] := True;
      if DimensionsArticle.NewDimChamp[iChamp] = 'GED_DISPOAFFECTEE' then DimensionsArticle.Dim.ReadOnly[iChamp] := True;
      if DimensionsArticle.NewDimChamp[iChamp] = 'GEL_QTEALIVRER' then DimensionsArticle.Dim.ReadOnly[iChamp] := True;
      if NatureDoc = NAT_AFFCDEAFF then
        if DimensionsArticle.NewDimChamp[iChamp] = 'GEL_QTERESERVEE' then DimensionsArticle.Dim.ReadOnly[iChamp] := True;
    end;
  end;
end;

procedure TOF_AFFCDESAISIEDIM.OnSaisiePiece(Validation: boolean);
var ItemDim: THDimensionItem;
  TobSelect, TobDoc: TOB;
  ichamp, i: integer;
  TableDim: THDimensionItemList;
begin
  MAJ := True;
  if not Validation then
  begin
    LaTob.ClearDetail;
    LaTob.AddChampSup('ANNULE', False);
    TheTob := DimensionsArticle.TOBArticleDim;
    exit;
  end;
  DimensionsArticle.ChangeChampDimMul(false);
  TableDim := DimensionsArticle.TableDim;
  if LaTob <> nil then
  begin
    for i := 0 to TableDim.count - 1 do
    begin
      if TableDim.Items[i] <> nil then
      begin
        ItemDim := THDimensionItem(TableDim.Items[i]);
        TobSelect := TOB(ItemDim.data);
        TobDoc := LaTob.findfirst(['GA_ARTICLE'], [TobSelect.GetValue('GA_ARTICLE')], false);
        if not TobDoc.FieldExists('GA_ARTICLE') then TobDoc.AddChampSup('GA_ARTICLE', false);
        TobDoc.PutValue('GA_ARTICLE', TobSelect.GetValue('GA_ARTICLE'));

        if not TobDoc.FieldExists('GEL_MODIFIABLE') then TobDoc.AddChampSup('GEL_MODIFIABLE', false);
        if not TobDoc.FieldExists('GEL_QTERESERVEE') then TobDoc.AddChampSup('GEL_QTERESERVEE', false);
        if not TobDoc.FieldExists('GEL_QTEAFFECTEE') then TobDoc.AddChampSup('GEL_QTEAFFECTEE', false);
        TobDoc.AddChampSup('MODIFIE', False);
        TobDoc.PutValue('MODIFIE', TobSelect.modifie);
        for iChamp := 0 to TobSelect.NbChamps do
        begin
          if TobDoc.FieldExists(TobSelect.GetNomChamp(iChamp)) then
            TobDoc.PutValue(TobSelect.GetNomChamp(iChamp), TobSelect.GetValeur(iChamp));
        end;
        for iChamp := 1000 to (1000 + (TobSelect.ChampsSup.count - 1)) do
        begin
          if TobDoc.FieldExists(TobSelect.GetNomChamp(iChamp)) then
            TobDoc.PutValue(TobSelect.GetNomChamp(iChamp), TobSelect.GetValeur(iChamp));
        end;
      end;
    end;
    TheTob := DimensionsArticle.TOBArticleDim;
  end;
  exit;
end;

procedure TOF_AFFCDESAISIEDIM.OnClose;
begin
  if NotClose then LastError := -1 else LastError := 0;
  if not MAJ then OnSaisiePiece(False);
  NotClose := False;
  // Sortie avec valider -> TobArticleDim conservé et retourné par TheTob
  DimensionsArticle.ConserveTobArtDim := True;
  if LastError = 0 then
    if (DimensionsArticle <> nil) then DimensionsArticle.Destroy;
end;

{==============================================================================================}
{=============================== Verification Stock  ==========================================}
{==============================================================================================}

procedure TOF_AFFCDESAISIEDIM.OnChangeItem(Sender: TObject);
var ItemDim: THDimensionItem;
begin
  ItemDim := THDimensionItem(Sender);
  if ItemDim = nil then exit;
  if (NatureDoc = NAT_AFFCDEAFF) then
  begin
    VerifStock(ItemDim);
  end;
end;

function TOF_AFFCDESAISIEDIM.VerifStock(ItemDim: THDimensionItem): Boolean;
var NomChamp: string;
  QteSaisie, QteDispo, QteReservee, QteDispoaff: Double;
  TobSelect: TOB;
  iChamp: Integer;
begin
  Result := False;
  QteSaisie := 0;
  QteDispoaff := 0;
  if ItemDim = nil then exit;
  for iChamp := 1 to MaxDimChamp do
  begin
    NomChamp := DimensionsArticle.OldDimChamp[iChamp];
    if NomChamp = 'GEL_QTEAFFECTEE' then
    begin
      QteSaisie := valeur(ItemDim.Valeur[iChamp]);
      Break;
    end;
  end;
  if QteSaisie = 0 then exit;
  TobSelect := TOB(ItemDim.data);
  QteDispo := TobSelect.GetValue('GED_QTEPREVI');
  QteReservee := TobSelect.GetValue('GEL_QTERESERVEE');
  if (NatureDoc = NAT_AFFCDEAFF) then QteDispoaff := TobSelect.GetValue('GED_DISPOAFFECTEE');
  if QteSaisie > QteReservee then
  begin
    PGIInfo(TraduireMemoire('La quantité demandée est supérieur à la quantité reservée'), Ecran.Caption);
    DimensionsArticle.Dim.FocusDim(ItemDim, iChamp);
    Result := True;
  end
  else
    if QteSaisie > QteDispoAff then //if QteSaisie > QteDispo then
  begin
    PGIInfo(TraduireMemoire('Quantité insuffisante pour cet article'), Ecran.Caption);
    DimensionsArticle.Dim.FocusDim(ItemDim, iChamp);
    Result := True;
  end;
end;

procedure AGLOnClickParamGrilleDoc2(parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then TOTOF := TFVierge(F).LaTOF else exit;
  {parm[1] contient la nature des champs affichés : préférences article/stock/achat/vente/autre}
  if (TOTOF is TOF_AFFCDESAISIEDIM) then
  begin
    ParamGrille(TOF_AFFCDESAISIEDIM(TOTOF).DimensionsArticle, TOF_AFFCDESAISIEDIM(TOTOF).NatureDoc, TOF_AFFCDESAISIEDIM(TOTOF).NaturePiece);
    TOF_AFFCDESAISIEDIM(TOTOF).AdapterFiche;
    TOF_AFFCDESAISIEDIM(TOTOF).SetReadOnly;
  end else exit;
end;

initialization
  registerclasses([TOF_AFFCDESAISIEDIM]);
  RegisterAglProc('OnClickParamGrilleDoc2', TRUE, 0, AGLOnClickParamGrilleDoc2);
end.

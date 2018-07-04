unit UtofSelectDimDoc;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
  HCtrls, HEnt1, HMsgBox, UTOF, Windows, ent1, SaisUtil,
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  DBGrids, dbTables, db,
  {$ENDIF}
  vierge, HDimension, UTOB, UtilArticle, UdimArticle, UtilPGI,
  AglInit, entgc, UtilDimArticle, M3FP, UtofMBOParamDim, ParamSoc, FactUtil, UtilTarif,
  TarifUtil;

type

  TOF_SELECTDIMDOC = class(TOF)
    DimensionsArticle: TODimArticle;
    Q: TQuery;

  private
    dimAction: string; // SELECT, SAISIE, CONSULT, MULTI
    dimChamp: string;
    dimMasque: string;
    CodeArticle, TarfArt, CodeTiers, TarifTiers, ArrondiT: string;
    PrixUnique: boolean;
    PrixHT, PrixTTC, Rem: Double;
    Top, Left, Height, Width: Integer;
    AuDessus, TarifGenCharge: Boolean; //AC TARIF
    NatureDoc, NaturePiece: string;
    MAJ: Boolean;
    Depot: string; // JD TRF
    QteStock: Double; // JD TRF
    NotClose: Boolean;
    RemA: Boolean; // remise autorisée?
    TOBTarif, TOBArticle, TOBCata, TOBMode: TOB; // Pour recherche tarif
    TypeTarfEtab, NatureType: string; // Pour Recherche tarif Detail ou ACH
    ModeTypeEtab: Integer; // Pour Recherche tarif Detail
    QTarifNegoce: TQuery; // Pour Recherche tarif Negoce
    RemiseTiers: Integer; // Pour Recherche tarif Negoce

    procedure InitDimensionsArticle;

    // Gestion document
    procedure PreparationSaisie;
    procedure InitTobSelect(var TobSelect: TOB);
    procedure InitPrix(TobSelect, TobDoc: TOB);
    function RecupTarifDetail(TobSelect, TobDoc: TOB): string;
    function RecupTarifNegoce(TobSelect, TobDoc: TOB; PrixBase: Double): string;
    function RecupPrixBase(TobSelect: TOB): Double;
    {$IFDEF NON_UTILISE}
    function GestionStock(TobSelect, TobDoc: TOB): Double;
    {$ENDIF}
    procedure CreerTobArticles;
    procedure CreerTobTarif;
    procedure OnSaisiePiece(Validation: boolean);
    procedure SetReadOnly;
    //procedure Consultation(Change: Boolean);
    procedure OnDoubleClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);

    // Gestion tarif ac
    procedure AffichageTarif;
    procedure Miseajourtarif(Validation: boolean); // pour article dim

  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
    procedure OnUpdate; override;
    procedure OnClose; override;
    procedure AdapterFiche;
    procedure OnChangeItem(Sender: TObject); //TRF
    function VerifStock(ItemDim: THDimensionItem): Boolean;
    function TraiteQteReste(ItemDim: THDimensionItem): Boolean;
    procedure TraiteParCombien(ItemDim: THDimensionItem);
    procedure TraiteReliquat(ItemDim: THDimensionItem);
    procedure LibereTout;
  end;

procedure AGLOnClickParamGrilleDoc(parms: array of variant; nb: integer);

implementation

Uses                { NEWPIECE }
    FactComm
    ;

procedure TOF_SELECTDIMDOC.OnArgument(Arguments: string);
var Lequel, CodeUnique: string;
  Critere: string;
  ChampMul, ValMul: string;
  x: integer;
  SQL: string;
begin
  inherited;
  dimAction := 'SAISIE';
  dimChamp := '';
  repeat
    Critere := uppercase(Trim(ReadTokenSt(Arguments)));
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
        if ChampMul = 'TOP' then Top := StrToInt(ValMul);
        if ChampMul = 'LEFT' then Left := StrToInt(ValMul);
        if ChampMul = 'HEIGTH' then Height := StrToInt(ValMul);
        if ChampMul = 'WIDTH' then Width := StrToInt(ValMul);
        if ChampMul = 'OU' then Audessus := Boolean(ValMul = 'X');
        //           if ChampMul='TARIF' then Tarif:=Boolean(ValMul='X') ; //AC TARIF
        if ChampMul = 'TYPEPARAM' then NatureDoc := ValMul;
        if ChampMul = 'CODEPARAM' then NaturePiece := ValMul;
        if ChampMul = 'DEPOT' then Depot := ValMul;

      end;
    end;
  until Critere = '';
  if (NaturePiece = NAT_ENTREXC) or (NaturePiece = NAT_SORTEXC) then NatureDoc := NaturePiece;
  CodeUnique := CodeArticleUnique(Lequel, '', '', '', '', '');

  SQL := 'SELECT GA_ARTICLE, GA_CODEARTICLE, GA_DIMMASQUE, GA_STATUTART, GA_PRIXUNIQUE,GA_REMISELIGNE,GA_TARIFARTICLE ';
  if (dimChamp <> '') and (NatureDoc <> 'TAF') then SQl := SQL + ',' + dimChamp;
  SQL := SQL + ' from Article where GA_CODEARTICLE="' + Lequel + '" AND GA_STATUTART="GEN"';
  Q := OpenSQl(SQL, True);
  //SetControlText('GA_ARTICLE', CodeUnique);
  CodeArticle := Q.FindField('GA_CODEARTICLE').AsString;
  TarfArt := Q.FindField('GA_TARIFARTICLE').AsString;
  PrixUnique := Boolean(Q.FindField('GA_PRIXUNIQUE').AsString = 'X');
  SetControlText('GA_CODEARTICLE', Lequel);
  dimMasque := Q.FindField('GA_DIMMASQUE').AsString;
  SetControlText('GA_STATUTART', Q.FindField('GA_STATUTART').AsString);
  RemA := Boolean(Q.FindField('GA_REMISELIGNE').AsString = 'X');
  try
    InitDimensionsArticle;
  finally
    Ferme(Q);
  end;
  Ecran.OnKeyDown := FormKeyDown;
end;

procedure TOF_SELECTDIMDOC.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = vk_valide then
  begin
    Key := 0;
    if (Ecran <> nil) and (Ecran is TFVierge) then TFVierge(Ecran).BValider.Click;
  end;
end;

procedure TOF_SELECTDIMDOC.InitDimensionsArticle;
var detail: string;
  item: THDimensionItem;
begin
  if DimensionsArticle <> nil then
    DimensionsArticle.free;
  DimensionsArticle := TODimArticle.Create(THDimension(GetControl('FDIM'))
    , CodeArticle
    , dimMasque
    , dimChamp, 'GCDIMCHAMP', NatureDoc, NaturePiece, Depot, '-', PrixUnique);
  //                     ,dimChamp,'GCDIMCHAMP',NatureDoc,NaturePiece, Depot, '-', PrixUnique, False);
  if DimensionsArticle.TOBArticleDim = nil then exit;
  DimensionsArticle.Dim.PopUp.Items[2].Visible := False; // Menu Existant invisible
  DimensionsArticle.Dim.PopUp.Items[3].Visible := False; // Menu Inexistant invisible
  DimensionsArticle.Dim.OnChange := OnChangeItem;
  DimensionsArticle.Dim.OnDblClick := OnDoubleClick;
  // DCA - FQ MODE 10445
  // Rend invisible l'option "Rendre existant" du "collage spéciale" de l'objet dimension (Option disponible en saisie de pièce / saisie de transferts)
  DimensionsArticle.Dim.Options.RendreExistanteVisible := False;
  // Rend invisible l'option "Rendre existant" du "Aller à" de l'objet dimension (Option disponible en saisie de transferts)
  DimensionsArticle.Dim.Options.AllerARendreExistanteVisible := False;
  if DimAction = 'CONSULT' then DimensionsArticle.Action := taConsult;
  if NatureDoc = 'TAF' then //AC TARIF
  begin
    AffichageTarif;
    if PrixUnique then DimensionsArticle.Action := taConsult; //Si prix unique interdire la modif des prix pour dim
  end else
  begin
    PreparationSaisie();
    detail := '-';
  end;
  {Affichage des champs par défaut de l'utilisateur}
  AfficheUserPref(DimensionsArticle, NatureDoc, NaturePiece);
  SetControlVisible('BPARAMDIM', NaturePieceGeree(naturePiece));
  SetReadOnly;
  // Gestion du solde reliquat
  if (LaTOB.FieldExists('_NotModifDim')) then DimensionsArticle.Action := taConsult ;
  //Consultation(False);
  // Positionnement sur 1ère cellule en saisie
  FocusedFirstItem(THDimension(GetControl('FDIM')));
end;

procedure TOF_SELECTDIMDOC.OnLoad; // AC TARIF
begin
  inherited;
  if DimensionsArticle.TOBArticleDim = nil then
  begin
    SetControlVisible('LPasDimensions', True);
    SetControlEnabled('BPARAMDIM', False);
    Ecran.Caption := 'Saisie articles: ' + LibelleArticleGenerique(CodeArticle);
    if Audessus then Ecran.Top := Top - Ecran.Height else Ecran.Top := Top;
    Ecran.Left := Left;
    if ExisteSQL('Select * from article where GA_CodeArticle="' + CodeArticle + '" and GA_Statutart="DIM"') then
      SetControlProperty('LPasDimensions', 'caption', ' Cet article ne se trouve pas en stock dans l''établissment ' + Depot + ' ' + RechDom('TTETABLISSEMENT',
        Depot, False) + ',' + #13 + #10 +
        ' vous ne pouvez pas le sélectionner');
    Exit;
  end;
  Ecran.Width := 1000; // pour pouvoir diminuer la fiche dans Adapter fiche
  Ecran.Height := 1000;
  MAJ := False;
  if NatureDoc = 'TAF' then Ecran.Caption := 'Tarif articles: ' + LibelleArticleGenerique(CodeArticle) else
    Ecran.Caption := 'Saisie articles: ' + LibelleArticleGenerique(CodeArticle);
  AdapterFiche;
  if Audessus then Ecran.Top := Top - Ecran.Height else Ecran.Top := Top;
  Ecran.Left := Left;
  //If (NatureDoc='TAF') and prixunique then SetControlProperty('LPXUNIQUE','visible',True) else
  SetControlProperty('LPXUNIQUE', 'visible', False);
end;

procedure TOF_SELECTDIMDOC.AdapterFiche;
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

  if (CoordDernLign.Bottom > Height) or (CoordDernLign.Bottom + 76 > Height) then
  begin
    Ecran.Height := Height;
    GridDim.ScrollBars := ssVertical;
  end else
    Ecran.Height := CoordDernLign.Bottom + 76 + ValOnglet; // 76= titre et panel bouton de ecran + encadrment objet Dim
  if (CoordDernCol.Right > Width) or (CoordDernCol.Right + 20 > Width) then
  begin
    Ecran.Width := Width;
    if GridDim.ScrollBars = ssVertical then GridDim.ScrollBars := ssBoth else GridDim.ScrollBars := ssHorizontal;
  end
  else
    if CoordDernCol.Right < 230 then Ecran.Width := 230
  else Ecran.Width := CoordDernCol.Right + 20;

  // Ajustement taille si ajout d'un ascenseur
  if (GridDim.ScrollBars = ssVertical) or (GridDim.ScrollBars = ssBoth) then
  begin
    if (CoordDernCol.Right + 16 > Width) or (CoordDernCol.Right + 16 + 20 > Width) then Ecran.Width := Width
    else Ecran.Width := CoordDernCol.Right + 16 + 20; //ScrollBars+encadrement Dim
  end;
  if (GridDim.ScrollBars = ssHorizontal) or (GridDim.ScrollBars = ssBoth) then
  begin
    if (CoordDernLign.Bottom + 16 > Height) or (CoordDernLign.Bottom + 16 + 76 > Height) then Ecran.Height := Height
    else Ecran.Height := CoordDernLign.Bottom + ValOnglet + 16 + 76; //ScrollBars+titre,panel bouton de ecran et encadrment objet Dim
  end;
  if Ecran.Width < 250 then Ecran.Width := 250;
  THDimension(GetControl('FDIM')).Align := alClient;
end;

procedure TOF_SELECTDIMDOC.OnClose;
begin
  if NotClose then LastError := -1 else LastError := 0;
  if not MAJ then
    if NatureDoc = 'TAF' then Miseajourtarif(False)
    else OnSaisiePiece(False);
  NotClose := False;
  if LastError = 0 then LibereTout;
end;

procedure TOF_SELECTDIMDOC.OnUpdate;
var ItemDim: THDimensionItem;
begin
  //if not (ctxMode in V_PGI.PGIContexte) then exit ;
  ItemDim := DimensionsArticle.Dim.CurrentItem;
  if NatureDoc = 'TAF' then Miseajourtarif(True)
  else
  begin
    if VerifStock(ItemDim) or TraiteQteReste(ItemDim) then
    begin
      NotClose := True;
      exit;
    end;
    //if NatureDoc = 'TAF' then Miseajourtarif(True)
    //else
      if DimensionsArticle.NewDimChamp[1] = '' then OnDoubleClick(ItemDim)
    else OnSaisiePiece(True);
  end ;
  inherited;
end;

procedure TOF_SELECTDIMDOC.LibereTout;
begin
  TOBTarif.Free;
  TOBTarif := nil;
  TOBArticle.Free;
  TOBArticle := nil;
  TOBCata.Free;
  TOBCata := nil;
  TOBMode.Free;
  TOBMode := nil;
  if QTarifNegoce <> nil then Ferme(QTarifNegoce);
  DimensionsArticle.Destroy;
end;

{==============================================================================================}
{=============================== Document + TRF =====================================================}
{==============================================================================================}

procedure TOF_SelectDimDoc.SetReadOnly;
var iChamp: integer;
begin
  if ((NatureDoc = NAT_VENTE) or (NatureDoc = NAT_ACHAT)) and not RemA then
  begin
    for iChamp := 1 to MaxDimChamp do
    begin
      if DimensionsArticle.NewDimChamp[iChamp] = 'GL_REMISELIGNE' then DimensionsArticle.Dim.ReadOnly[iChamp] := True;
    end;
  end;
end;

{procedure TOF_SELECTDIMDOC.Consultation(Change: Boolean);
begin
  if DimensionsArticle.NewDimChamp[1] = '' then DimensionsArticle.Action := taConsult
  else
    if Change then DimensionsArticle.Action := taCreat;
end;}

procedure TOF_SELECTDIMDOC.PreparationSaisie;
var ItemDim: THDimensionItem;
  TobSelect, TobDoc: TOB;
  i: integer;
  TableDim: THDimensionItemList;
begin
  TableDim := DimensionsArticle.TableDim;
  if LaTob <> nil then
  begin
    //TOBTarif:=TOB.Create('',nil,-1) ;
    TarifGenCharge := False;
    CreerTobArticles;
    CreerTobTarif;
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
          InitPrix(TobSelect, TobDoc);
          if TobSelect.FieldExists('GL_PIECEPRECEDENTE') and TobDoc.FieldExists('GL_PIECEPRECEDENTE') then
             TobSelect.PutValue('GL_PIECEPRECEDENTE', TobDoc.GetValue('GL_PIECEPRECEDENTE'));
          if TobSelect.FieldExists('GL_QTERESTE') and TobDoc.FieldExists('GL_QTERESTE') then
             TobSelect.PutValue('GL_QTERESTE',TobDoc.GetValue('GL_QTERESTE')) ;
          if (TobDoc.GetValue('GL_QTEFACT') <> 0) or (TobDoc.GetValue('GL_QTERELIQUAT') <> 0) then
          begin
            TobSelect.PutValue('GL_QTEFACT', TobDoc.GetValue('GL_QTEFACT'));
            TobSelect.PutValue('GL_PUHTDEV', TobDoc.GetValue('GL_PUHTDEV'));
            TobSelect.PutValue('GL_QTERELIQUAT', TobDoc.GetValue('GL_QTERELIQUAT'));
            if (NatureDoc <> NAT_INVENTAIRE) then
            begin
              TobSelect.PutValue('GL_REMISELIGNE', TobDoc.GetValue('GL_REMISELIGNE'));
              if (NatureDoc <> NAT_TRANSFERT) and (NatureDoc <> NAT_ENTREXC) and (NatureDoc <> NAT_SORTEXC) then
              begin
                TobSelect.PutValue('GL_PUTTCDEV', TobDoc.GetValue('GL_PUTTCDEV'));
                if (NatureDoc <> NAT_TOBVIEW) then
                  TobSelect.PutValue('GL_CODEARRONDI', TobDoc.GetValue('GL_CODEARRONDI'));
              end;
            end;
          end else
          begin
            TobSelect.PutValue('GL_QTEFACT', 0);
            TobSelect.PutValue('GL_PUHTDEV', PrixHT);
            TobSelect.PutValue('GL_QTERELIQUAT', TobDoc.GetValue('GL_QTERELIQUAT'));
            if (NatureDoc <> NAT_INVENTAIRE) then
            begin
              TobSelect.PutValue('GL_REMISELIGNE', Rem);
              if (NatureDoc <> NAT_TRANSFERT) and (NatureDoc <> NAT_ENTREXC) and (NatureDoc <> NAT_SORTEXC) then
              begin
                TobSelect.PutValue('GL_PUTTCDEV', PrixTTC);
                if (NatureDoc <> NAT_TOBVIEW) then
                  TobSelect.PutValue('GL_CODEARRONDI', ArrondiT);
              end;
            end;
          end;
          TobSelect.PutValue('GQ_PHYSIQUE', QteStock);
        end;
      end;
    end;
  end;
end;

/// Initialisation des champs de l'objet dimension

procedure TOF_SELECTDIMDOC.InitPrix(TobSelect, TobDoc: TOB);
var EnHt, RechTarif: boolean;
  VenteAchat, PxRem: string;
  //QteDejaSaisi: Double ;
begin
  PrixHT := 0;
  PrixTTC := 0;
  Rem := 0;
  QteStock := 0;
  //QteDejaSaisi:=TobDoc.GetValue('_QteDejaSaisi') ;
  VenteAchat := LaTob.GetValue('GP_VENTEACHAT');
  EnHt := (LaTob.GetValue('GP_FACTUREHT') = 'X');
  RechTarif := (GetInfoParPiece(naturePiece, 'GPP_CONDITIONTARIF') = 'X');
  QteStock := TobSelect.getValue('GQ_PHYSIQUE');
  if EnHt then
  begin
    if RechTarif then
    begin
      if VenteAchat = 'VEN' then PxRem := RecupTarifNegoce(TobSelect, TobDoc, PrixHt)
      else PxRem := RecupTarifDetail(TobSelect, TobDoc);
      PrixHt := Valeur(ReadTokenSt(PxRem));
      Rem := Valeur(PxRem);
    end;
    if PrixHT = 0 then PrixHT := RecupPrixBase(TOBSelect);
  end else
  begin
    if RechTarif then
    begin
      PxRem := RecupTarifDetail(TobSelect, TobDoc);
      PrixTTC := Valeur(ReadTokenSt(PxRem));
      Rem := Valeur(PxRem);
    end;
    if PrixTTC = 0 then PrixTTC := RecupPrixBase(TOBSelect);
  end;
  //if VenteAchat='ACH' then PrixTTC:=0 ;  A revoir
end;

{$IFDEF NON_UTILISE}

function TOF_SELECTDIMDOC.GestionStock(TobSelect, TobDoc: TOB): Double;
var QtePlus, QteMoins, Champ: string;
  QteDejaSaisi, StockPhy: Double;
begin
  QteStock := 0;
  QteDejaSaisi := TobDoc.GetValue('_QteDejaSaisi');
  StockPhy := TobSelect.getValue('GQ_PHYSIQUE');
  QteMoins := GetInfoParPiece(naturePiece, 'GPP_QTEMOINS');
  QtePlus := GetInfoParPiece(naturePiece, 'GPP_QTEPLUS');
  repeat
    Champ := uppercase(Trim(ReadTokenSt(QteMoins)));
    if Champ = 'PHY' then QteStock := StockPhy - QteDejaSaisi;
  until Champ = '';
  if QteStock = 0 then
    repeat
      Champ := uppercase(Trim(ReadTokenSt(QtePlus)));
      if Champ = 'PHY' then QteStock := StockPhy + QteDejaSaisi;
    until Champ = '';
  if QteStock = 0 then QteStock := StockPhy;
  Result := QteStock;
end;
{$ENDIF}

{==============================================================================================}
{============================================== Tarifs ========================================}
{==============================================================================================}

procedure TOF_SELECTDIMDOC.CreerTobTarif;
var codeArt, TarifArticle, CodeDevise, VenteAchat, NomChamp: string;
  LaDate: TDateTime;
  Q, QTiers: TQuery;
  TobM: Tob;
  EnHt: boolean;
  SQL: string;
begin
  ModeTypeEtab := -1;
  TOBMode := TOB.Create('', nil, -1);
  EnHt := (LaTob.GetValue('GP_FACTUREHT') = 'X');
  if EnHT then
  begin
    NatureType := 'ACH';
    NomChamp := 'ET_TYPETARIFACH';
  end else
  begin
    NatureType := 'VTE';
    NomChamp := 'ET_TYPETARIF';
  end;
  Q := OpenSQL('Select ' + NomChamp + ' from ETABLISS Where ET_ETABLISSEMENT="' + Depot + '"', True);
  if not Q.EOF then TypeTarfEtab := Q.Fields[0].AsString;
  if TypeTarfEtab = '' then TypeTarfEtab := '...';
  Ferme(Q);
  SQL := 'Select GFM_TARFMODE,GFM_TYPETARIF,GFM_NATURETYPE,GFM_DATEDEBUT,GFM_PROMO,GFM_COEF,GFM_DEVISE from TARIFMODE where gfm_typetarif in ("' + TypeTarfEtab
    + '","...") and gfm_naturetype="' + NatureType + '" order by GFM_TYPETARIF DESC,GFM_DATEDEBUT DESC';
  Q := OpenSQL(SQL, True);
  TOBMode.LoadDetailDB('TARIFMODE', '', '', Q, False);
  Ferme(Q);
  Codetiers := LaTob.GetValue('GP_TIERS');
  TarifTiers := LaTob.GetValue('GP_TARIFTIERS');
  CodeArt := CodeArticleUnique(CodeArticle, '', '', '', '', '');
  TarifArticle := TarfArt;
  CodeDevise := LaTob.GetValue('GP_DEVISE');
  VenteAchat := LaTob.GetValue('GP_VENTEACHAT');
  LaDate := LaTob.GetValue('GP_DATEPIECE');
 //TOBTarif := CreerTobTarifArt(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, Depot, VenteAchat, LaDate, EnHT);
  if (VenteAchat = 'VEN') and EnHt then
  begin
    // Charge QTarifNegoce
    QTiers := OpenSQL('Select T_REMISE from tiers where T_TIERS="' + CodeTiers + '"', True);
    RemiseTiers := QTiers.Fields[0].AsInteger;
    ferme(QTiers);
    SQL := SQLTarifGen(CodeArt, TarfArt, CodeTiers, TarifTiers, CodeDevise, Depot, LaDate, 0);
    SQL := SQl + CompleteSQLTarif(VenteAchat, False, EnHt);
    QTarifNegoce := OpenSQL(SQL, True);
    //
  end else TOBTarif := CreerTobTarifArt(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, Depot, VenteAchat, LaDate, EnHT);
end;

function TOF_SELECTDIMDOC.RecupTarifDetail(TobSelect, TobDoc: TOB): string;
var codeArt, CodeDevise, InfoRem: string;
  Remise, PrixRemise, PrixBase: double;
begin
  Remise := 0;
  CodeDevise := LaTob.GetValue('GP_DEVISE');
  CodeArt := TobSelect.GetValue('GA_ARTICLE');
  // Prix remisé (Période Promo)
  InfoRem := RechTarifSpec(TOBTarif, TOBMode, CodeArt, Depot, CodeDevise, TypeTarfEtab, NatureType);
  PrixRemise := Valeur(ReadTokenSt(InfoRem));
  // Prix de base (Hors promotion)
  if PrixRemise = 0 then
  begin
    PrixBase := RechPrixTarifBase(TOBTarif, TOBMode, CodeArt, Depot, CodeDevise, TypeTarfEtab, NatureType);
    // remise
    InfoRem := '';
    InfoRem := ChercheMieRem(TOBTarif, TobMode, CodeArt, TarfArt, CodeTiers, TarifTiers, Depot, CodeDevise, TypeTarfEtab, NatureType);
    Remise := Valeur(ReadTokenSt(InfoRem));
    ArrondiT := ReadTokenSt(InfoRem);
  end
  else PrixBase := 0;
  if (PrixRemise = 0) and (PrixBase <> 0) then PrixRemise := PrixBase;
  result := FloatToStr(PrixRemise) + ';' + FloatToStr(Remise);
end;

function TOF_SELECTDIMDOC.RecupTarifNegoce(TobSelect, TobDoc: TOB; PrixBase: Double): string;
var Q, QTiers: TQuery;
  QueQte, EnHT, Find: boolean;
  SQL, VenteAchat: AnsiString;
  QuelleCascade, CodeArt, CodeTiers, CodeDepot, CodeDevise, TarifArticle, TarifTiers, StCascade: string;
  LaDate: TDateTime;
  Qte, RemiseTiers, Remise, PrixUnitaire: Double;
  DEV: Rdevise;
begin
  StCascade := '';
  Remise := 0.0;
  QuelleCascade := 'MIE';
  PrixUnitaire := 0.0;
  QueQte := False;
  EnHt := True;
  Find := False;
  CodeArt := TobSelect.GetValue('GA_ARTICLE');
  CodeTiers := LaTob.GetValue('GP_TIERS');
  CodeDepot := Depot;
  CodeDevise := LaTob.GetValue('GP_DEVISE');
  Qte := 0;
  TarifArticle := TobSelect.GetValue('GA_TARIFARTICLE');
  LaDate := LaTob.GetValue('GP_DATEPIECE');
  TarifTiers := LaTob.GetValue('GP_TARIFTIERS');
  VenteAchat := LaTob.GetValue('GP_VENTEACHAT');

  //CodeArt:=CodeArticleUnique(Copy(CodeArt,1,18),'','','','','') ;
  //SQL:=SQLTarif(CodeArt,TarifArticle,CodeTiers,TarifTiers,CodeDevise,CodeDepot,LaDate,Qte) ;
  //SQL:=SQl+CompleteSQLTarif(VenteAchat,QueQte,EnHT) ;
  //Q:=OpenSQL(SQL,True) ;

  if not QTarifNegoce.EOF then
  begin
    if QTarifNegoce.EOF then
    begin
      Ferme(QTarifNegoce);
      Exit;
    end;
    StCascade := StCascade + QTarifNegoce.FindField('GF_CALCULREMISE').AsString;
    QuelleCascade := QTarifNegoce.FindField('GF_CASCADEREMISE').AsString;
    Remise := QTarifNegoce.FindField('GF_REMISE').AsFloat;
    // MODIF MODE LM
    {$IFDEF MODE}
    QTarifNegoce.First;
    while not QTarifNegoce.EOF do
    begin
      if QTarifNegoce.FindField('GF_ARTICLE').AsString = CodeArt then
      begin
        PrixUnitaire := QTarifNegoce.FindField('GF_PRIXUNITAIRE').AsFloat;
        break;
      end;
      QTarifNegoce.Next;
    end;
    QTarifNegoce.First;
    {$ENDIF}
    Find := True;
    if ((not QueQte) and (QTarifNegoce.FindField('GF_CASCADEREMISE').AsString = 'CAS')) then
      // Si premier = cascade --> dérouler cascade, y compris l'éventuel dernier non cascade
    begin
      repeat
        QTarifNegoce.Next;
        if QTarifNegoce.EOF then Break;
        QuelleCascade := QTarifNegoce.FindField('GF_CASCADEREMISE').AsString;
        if QTarifNegoce.FindField('GF_PRIXUNITAIRE').AsFloat <> 0 then PrixUnitaire := QTarifNegoce.FindField('GF_PRIXUNITAIRE').AsFloat;
        Remise := 100.0 * (1.0 - (1.0 - Remise / 100.0) * (1.0 - QTarifNegoce.FindField('GF_REMISE').AsFloat / 100));
        StCascade := StCascade + QTarifNegoce.FindField('GF_CALCULREMISE').AsString;
      until ((QTarifNegoce.EOF) or (QTarifNegoce.FindField('GF_CASCADEREMISE').AsString <> 'CAS'));
    end;
  end;
  //Ferme(QTarifNegoce) ;

  //if QuelleCascade='FOR' then else {rien à faire, le tarif a été trouvé}
  //if QuelleCascade='CAS' then else {rien à faire, la remise cascade a été caclulée}
  //if ((Not QueQte) and (QuelleCascade='MIE')) then
  //   BEGIN
     {prendre la meilleure entre remise issue du tarif et celle du tiers}
     //if RemiseTiers>Remise then Remise:=RemiseTiers ;
  //   END else if ((Not QueQte) and (QuelleCascade='CUM')) then
  //   BEGIN
     {cumuler la remise du tarif et celle du tiers}
  //   Remise:=100.0*(1.0-(1.0-Remise/100.0)*(1.0-RemiseTiers/100)) ;
  //   END ;
  // Si pas de tarif
  if not Find then PrixUnitaire := PrixBase;
  result := FloatToStr(PrixUnitaire) + ';' + FloatToStr(Remise);
end;
{==============================================================================================}

procedure TOF_SELECTDIMDOC.CreerTobArticles;
var SQLArt, SQLCata: string;
  QArticle, QCata: TQuery;
begin
  Codetiers := LaTob.GetValue('GP_TIERS');
  TOBArticle := TOB.Create('LesArticles', nil, -1);
  SQLArt := 'SELECT * FROM ARTICLE WHERE GA_ARTICLE like "' + CodeArticle + '%" ORDER BY GA_ARTICLE';
  QArticle := OpenSQL(SQLArt, True);
  if not QArticle.EOF then TOBArticle.LoadDetailDB('ARTICLE', '', '', QArticle, False);
  Ferme(QArticle);
  TOBCata := TOB.Create('LeCata', nil, -1);
  SQLCata := 'SELECT * FROM CATALOGU WHERE GCA_TIERS="' + CodeTiers + '" ORDER BY GCA_ARTICLE';
  QCata := OpenSQL(SQLCata, True);
  if not QCata.EOF then TOBCata.LoadDetailDB('CATALOGU', '', '', QCata, False);
  Ferme(QCata);
end;

function TOF_SELECTDIMDOC.RecupPrixBase(TobSelect: TOB): Double;
var CodeP, NaturePiece, CodeArt, CodeDevise: string;
  Prix: Double;
  DPACat, BaseCat: Double;
  TOBA, TOBC, TOBM: TOB;
  DEV: RDEVISE;
  LaDate: TDateTime;
begin
  CodeArt := TobSelect.GetValue('GA_ARTICLE');
  TOBA := TOBArticle.FindFirst(['GA_ARTICLE'], [CodeArt], False);
  TOBC := TOBCata.FindFirst(['GCA_ARTICLE'], [CodeArticleUnique2(CodeArticle, '')], False);
  NaturePiece := LaTob.GetValue('GP_NATUREPIECEG');
  CodeDevise := LaTob.GetValue('GP_DEVISE');
  LaDate := LaTob.GetValue('GP_DATEPIECE');
  Result := 0;
  if TOBA = nil then Exit;
  CodeP := GetInfoParPiece(NaturePiece, 'GPP_APPELPRIX');
  Prix := 0;
  if CodeP = 'DPA' then
  begin
    if TOBC <> nil then DPACat := TOBC.GetValue('GCA_DPA') else DPACat := 0;
    if DPACat <> 0 then Prix := DPACat else
    begin
      if TOBSelect.FieldExists('GQ_DPA') then
      begin
        Prix := Arrondi(TOBSelect.GetValue('GQ_DPA'), V_PGI.OkDecV);
        if (Prix <= 0) then Prix := Arrondi(TOBA.GetValue('GA_DPA'), V_PGI.OkDecV);
      end
      else Prix := Arrondi(TOBA.GetValue('GA_DPA'), V_PGI.OkDecV);
    end;
  end else
    if CodeP = 'DPR' then
  begin
    if TOBSelect.FieldExists('GQ_DPR') then
    begin
      Prix := Arrondi(TOBSelect.GetValue('GQ_DPR'), V_PGI.OkDecV);
      if (Prix <= 0) then Prix := Arrondi(TOBA.GetValue('GA_DPR'), V_PGI.OkDecV);
    end
    else Prix := Arrondi(TOBA.GetValue('GA_DPR'), V_PGI.OkDecV);
  end else
    if CodeP = 'PPA' then
  begin
    if TOBSelect.FieldExists('GQ_PMAP') then
    begin
      Prix := Arrondi(TOBSelect.GetValue('GQ_PMAP'), V_PGI.OkDecV);
      if (Prix <= 0) then Prix := Arrondi(TOBA.GetValue('GA_PMAP'), V_PGI.OkDecV);
    end
    else Prix := Arrondi(TOBA.GetValue('GA_PMAP'), V_PGI.OkDecV);
  end else
    if CodeP = 'PPR' then
  begin
    if TOBSelect.FieldExists('GQ_PMRP') then
    begin
      Prix := Arrondi(TOBSelect.GetValue('GQ_PMRP'), V_PGI.OkDecV);
      if (Prix <= 0) then Prix := Arrondi(TOBA.GetValue('GA_PMRP'), V_PGI.OkDecV);
    end
    else Prix := Arrondi(TOBA.GetValue('GA_PMRP'), V_PGI.OkDecV);
  end else
    if CodeP = 'PAS' then Prix := Arrondi(TOBA.GetValue('GA_PAHT'), V_PGI.OkDecV) else
    if CodeP = 'PRS' then Prix := Arrondi(TOBA.GetValue('GA_PRHT'), V_PGI.OkDecV) else
    if CodeP = 'CAT' then
  begin
    if TOBC <> nil then BaseCat := TOBC.GetValue('GCA_PRIXBASE') else BaseCat := 0;
    if BaseCat <> 0 then Prix := BaseCat else Prix := Arrondi(TOBA.GetValue('GA_DPA'), V_PGI.OkDecV);
  end else
    if CodeP = 'PUH' then
  begin
    if LaTob.GetValue('GP_FACTUREHT') = 'X' then Prix := Arrondi(TOBA.GetValue('GA_PVHT'), V_PGI.OkDecV)
    else Prix := Arrondi(TOBA.GetValue('GA_PVTTC'), V_PGI.OkDecV);
  end else
    if CodeP = 'MT1' then Prix := Arrondi(TOBA.GetValue('GA_VALLIBRE1'), V_PGI.OkDecV) else
    if CodeP = 'MT2' then Prix := Arrondi(TOBA.GetValue('GA_VALLIBRE2'), V_PGI.OkDecV) else
    if CodeP = 'MT3' then Prix := Arrondi(TOBA.GetValue('GA_VALLIBRE3'), V_PGI.OkDecV);

  if (Prix <= 0) and (CodeP <> 'PUH') then // DC - Valeur par défaut - Fonctionnement Mode généralisé
  begin
    if (CodeP = 'DPR') or (CodeP = 'PPR')
      then Prix := Arrondi(TOBA.GetValue('GA_PRHT'), V_PGI.OkDecV)
    else Prix := Arrondi(TOBA.GetValue('GA_PAHT'), V_PGI.OkDecV);
  end; // DC Fin

  ////// Gestion des devises
  if LaTob.GetValue('GP_SAISIECONTRE') = 'X' then
  begin
    if VH^.TenueEuro then Prix := EuroToFranc(Prix) else Prix := FrancToEuro(Prix);
  end;
  /// Conversion  MODIF AC
  if CodeDevise <> V_PGI.DevisePivot then // Saisie en devise
  begin
    //TOBM:=TOBMode.FindFirst(['GFM_DEVISE'],[CodeDevise],False) ;
    //if TOBM<>nil then Prix:=Prix*TOBM.GetValue('GFM_COEF') else
    begin
      DEV.Code := CodeDevise;
      GetInfosDevise(DEV);
      DEV.Taux := GetTaux(DEV.Code, DEV.DateTaux, LaDate);
      if (VH_GC.GCComportePrixDev = 'ALE') or (VH_GC.GCComportePrixDev = 'BLO') then Prix := 0
      else
      begin
        Prix := PivotToDevise(Prix, DEV.Taux, DEV.Quotite, DEV.Decimale);
        // MODIF MODE LM
        Prix := Prix * V_PGI.TauxEuro;
      end;
    end;
  end;
  // MODIF AC
  Result := Prix;
end;

// Pour initialiser les valeur null à zero

procedure TOF_SELECTDIMDOC.InitTobSelect(var TobSelect: TOB);
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
end;

procedure TOF_SELECTDIMDOC.OnSaisiePiece(Validation: boolean);
var ItemDim: THDimensionItem;
  TobSelect, TobDoc: TOB;
  ichamp, i, j: integer;
  TableDim: THDimensionItemList;
begin
  MAJ := True;
  if not Validation then
  begin
    LaTob.ClearDetail;
    LaTob.AddChampSup('ANNULE', False);
    TheTob := LaTob; // retourne la Tob originale
    exit;
  end;
  j := 0;
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
        { NEWPIECE }
        if ((TobSelect.GetValue('GL_QTEFACT') = 0) and (not WithPiecePrecedente(TobDoc))) or (TobDoc = nil) then
        begin
          if TobDoc <> nil then TobDoc.Free;
          continue;
        end;
        if not TobDoc.FieldExists('GA_ARTICLE') then TobDoc.AddChampSup('GA_ARTICLE', false);
        TobDoc.PutValue('GA_ARTICLE', TobSelect.GetValue('GA_ARTICLE'));

        if not TobDoc.FieldExists('GL_QTEFACT') then TobDoc.AddChampSup('GL_QTEFACT', false);
        if not TobDoc.FieldExists('GL_PUHTDEV') then TobDoc.AddChampSup('GL_PUHTDEV', false);
        if not TobDoc.FieldExists('GL_PUTTCDEV') then TobDoc.AddChampSup('GL_PUTTCDEV', false);
        if not TobDoc.FieldExists('GL_REMISELIGNE') then TobDoc.AddChampSup('GL_REMISELIGNE', false);
        if not TobDoc.FieldExists('GL_QTERELIQUAT') then TobDoc.AddChampSup('GL_QTERELIQUAT', false);
        if not TobDoc.FieldExists('GL_CODEARRONDI') then TobDoc.AddChampSup('GL_CODEARRONDI', false);
        TobDoc.AddChampSup('_TRAITEMENT', false);

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
    while j <= LaTob.Detail.Count - 1 do
    begin
      if not LaTob.Detail[j].FieldExists('_TRAITEMENT') then LaTob.Detail[j].Free
      else
      begin
        LaTob.Detail[j].DelChampSup('_TRAITEMENT', false);
        j := j + 1;
      end;
    end;
    TheTob := LaTob;
  end;
  exit;
end;

procedure TOF_SELECTDIMDOC.OnDoubleClick(Sender: TObject);
var ItemDim, ItemDimFact: THDimensionItem;
  TobSelect, TobDoc: TOB;
  i, j: Integer;
  TableDim: THDimensionItemList;
begin
  if DimensionsArticle.NewDimChamp[1] <> '' then exit;
  MAJ := True;
  j := 0;
  ItemDim := THDimensionItem(Sender);
  if ItemDim = nil then exit;
  TableDim := DimensionsArticle.TableDim;
  if LaTob <> nil then
  begin
    for i := 0 to TableDim.count - 1 do
    begin
      if TableDim.Items[i] <> nil then
      begin
        ItemDimFact := THDimensionItem(TableDim.Items[i]);
        if ItemDim = ItemDimFact then
        begin
          TobSelect := TOB(ItemDim.data);
          TobDoc := tob.create('_Article Dim', LaTob, -1);
          TobDoc.AddChampSup('GA_ARTICLE', false);
          TobDoc.PutValue('GA_ARTICLE', TobSelect.GetValue('GA_ARTICLE'));
          TobDoc.AddChampSup('GL_QTEFACT', false);
          TobDoc.AddChampSup('GL_PUHTDEV', false);
          TobDoc.PutValue('GL_QTEFACT', 1);
          TobDoc.PutValue('GL_PUHTDEV', TobSelect.GetValue('GL_PUHTDEV'));
          TobDoc.AddChampSup('GL_REMISELIGNE', false);
          TobDoc.PutValue('GL_REMISELIGNE', TobSelect.GetValue('GL_REMISELIGNE'));
          if (NatureDoc <> NAT_TRANSFERT) and (NatureDoc <> NAT_INVENTAIRE) and (NatureDoc <> NAT_ENTREXC) and (NatureDoc <> NAT_SORTEXC) then
          begin
            TobDoc.AddChampSup('GL_PUTTCDEV', false);
            TobDoc.PutValue('GL_PUTTCDEV', TobSelect.GetValue('GL_PUTTCDEV'));
          end;
          if (NatureDoc <> NAT_TOBVIEW) then
          begin
            TobDoc.AddChampSup('GL_CODEARRONDI', false);
            TobDoc.PutValue('GL_CODEARRONDI', TobSelect.GetValue('GL_CODEARRONDI'));
          end;
          TobDoc.AddChampSup('_TRAITEMENT', false);
          while j <= LaTob.Detail.Count - 1 do
          begin
            if not LaTob.Detail[j].FieldExists('_TRAITEMENT') then LaTob.Detail[j].Free
            else
            begin
              LaTob.Detail[j].DelChampSup('_TRAITEMENT', false);
              j := j + 1;
            end;
          end;
          TheTob := LaTob;
          TFVierge(Ecran).close;
        end;
      end;
    end;
  end;
end;

{==============================================================================================}
{=============================== Tarif ========================================================}
{==============================================================================================}
//Affiche tarif avec Tob provenant de tarif

procedure TOF_SELECTDIMDOC.AffichageTarif; //AC TARIF
var ItemDim: THDimensionItem;
  TobDimAff, TobTarfArtDim: TOB;
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
        TobDimAff := TOB(ItemDim.data);
        TobTarfArtDim := LaTob.findfirst(['GF_ARTICLE'], [TobDimAff.GetValue('GA_ARTICLE')], false);
        if TobTarfArtDim <> nil then
        begin
          TobDimAff.PutValue('GF_PRIXUNITAIRE', TobTarfArtDim.GetValue('GF_PRIXUNITAIRE'));
          TobDimAff.PutValue('GF_CALCULREMISE', TobTarfArtDim.GetValue('GF_CALCULREMISE'));
        end;
      end;
    end;
  end;
end;

// Mise à jour avec Tob de tarif (avec retour Tob à Tarif

procedure TOF_SELECTDIMDOC.Miseajourtarif(Validation: boolean); // Modif Tarif Dim
var ItemDim: THDimensionItem;
  TobDimAff, TobTarfArtDim: TOB;
  i, iChamp: integer;
  TableDim: THDimensionItemList;
begin
  MAJ := True;
  if not Validation then
  begin
    TheTob := LaTob; // retourne la Tob originale
    TheTob.SetAllModifie(False);
    exit;
  end;
  TableDim := DimensionsArticle.TableDim;
  if LaTob <> nil then
  begin
    for i := 0 to TableDim.count - 1 do
    begin
      if TableDim.Items[i] <> nil then
      begin
        ItemDim := THDimensionItem(TableDim.Items[i]);
        TobDimAff := TOB(ItemDim.data);
        TobTarfArtDim := LaTob.findfirst(['GF_ARTICLE'], [TobDimAff.GetValue('GA_ARTICLE')], false);
        if TobTarfArtDim <> nil then
        begin
          for iChamp := 1 to MaxDimChamp do
          begin
            if (DimensionsArticle.OldDimChamp[iChamp] <> '') then //and (ItemDim.valeur[iChamp]<>TobDimAff.GetValue(DimensionsArticle.OldDimChamp[iChamp])) then
            begin
              TobTarfArtDim.PutValue(DimensionsArticle.OldDimChamp[iChamp], ItemDim.Valeur[iChamp]);
            end;
          end;
        end;
      end;
    end;
    TheTob := LaTob;
    { for i_ind := TobTarfArtDim.Detail.Count - 1 downto 0 do
     BEGIN
     TobTarfArtDim.Detail[i_ind].Free;
     END;}
  end;
  exit;
end;

{==============================================================================================}
{=============================== Verification Stock + PCB =====================================}
{==============================================================================================}

procedure TOF_SELECTDIMDOC.OnChangeItem(Sender: TObject);
var ItemDim: THDimensionItem;
begin
  ItemDim := THDimensionItem(Sender);
  if ItemDim = nil then exit;
  if (NatureDoc <> NAT_ACHAT) and (NatureDoc <> NAT_TARIF) and (NatureDoc <> NAT_INVENTAIRE) then
  begin
    TraiteParCombien(ItemDim);
    VerifStock(ItemDim);
  end;
  if (NatureDoc = NAT_ACHAT) or (NatureDoc = NAT_VENTE) then
  begin
    TraiteQteReste(ItemDim) ;
    TraiteReliquat(ItemDim);
  end;
end;

function TOF_SELECTDIMDOC.TraiteQteReste(ItemDim: THDimensionItem): Boolean;
var Ichamp,IQteSaisie :Integer ;
    NomChamp: String ;
    QteSaisie,QteReste,QteFact : Double ;
    TobSelect,TobDoc: TOB;
begin
  Result :=False ;
  QteSaisie := 0;
  QteReste := 0 ;
  if ItemDim = nil then exit;
  for iChamp := 1 to MaxDimChamp do
  begin
    NomChamp := DimensionsArticle.OldDimChamp[iChamp];
    if NomChamp = 'GL_QTEFACT' then
    begin
      QteSaisie := valeur(ItemDim.Valeur[iChamp]);
      IQteSaisie := iChamp ;
      Break;
    end ;
  end;
   TobSelect := TOB(ItemDim.data);
   TobDoc := LaTob.findfirst(['GA_ARTICLE'], [TobSelect.GetValue('GA_ARTICLE')], false);
   QteFact :=  TobDoc.GetValue('GL_QTEFACT');
   QteReste := TobSelect.GetValue('GL_QTERESTE');
   if QteFact <> QteReste then
   begin
     if QteSaisie < QteFact - QteReste then
     begin
          Result:=True ;
          PGIInfo('Attention ! La quantité saisie est inférieure au total des quantités des pièces suivantes',Ecran.Caption) ;
          ItemDim.Valeur[IQteSaisie] := QteFact ;                // Modif LM 121103
          DimensionsArticle.Dim.FocusDim(ItemDim, IQteSaisie) ;  // Modif LM 121103
          DimensionsArticle.Dim.InvalidateDim(ItemDim, False);   // Modif LM 121103
          DimensionsArticle.Dim.UpdateEditing;   // Modif LM 121103
     end ;
   end ;
end ;

function TOF_SELECTDIMDOC.VerifStock(ItemDim: THDimensionItem): Boolean;
var ForceRupt, EnRupture: Boolean;
  CalcRupture, NomChamp: string;
  QteSaisie, QteDispo, QteReserveCli, QtePrepaCli: Double;
  TobSelect: TOB;
  iChamp: Integer;
begin
  Result := False;
  EnRupture := False;
  QteSaisie := 0;
  if ItemDim = nil then exit;
  for iChamp := 1 to MaxDimChamp do
  begin
    NomChamp := DimensionsArticle.OldDimChamp[iChamp];
    if NomChamp = 'GL_QTEFACT' then
    begin
      QteSaisie := valeur(ItemDim.Valeur[iChamp]);
      Break;
    end ;
  end;
  if QteSaisie = 0 then exit;
  TobSelect := TOB(ItemDim.data);
  QteDispo := TobSelect.GetValue('GQ_PHYSIQUE');
  QteReserveCli := TobSelect.GetValue('GQ_RESERVECLI');
  QtePrepaCli := TobSelect.GetValue('GQ_PREPACLI');
  ForceRupt := (GetInfoParPiece(NaturePiece, 'GPP_FORCERUPTURE') = 'X');
  CalcRupture := GetInfoParPiece(NaturePiece, 'GPP_CALCRUPTURE');
  if CalcRupture = 'DIS' then
    if QteSaisie > (QteDispo - QteReserveCli - QtePrepaCli) then EnRupture := True;
  if CalcRupture = 'PHY' then
    if QteSaisie > QteDispo then EnRupture := True;
  if EnRupture then
  begin
    if ForceRupt then
      if HShowMessage('1;' + Ecran.Caption + ';' + 'Cet article est en rupture, confirmez-vous malgré tout la quantité ?' + ';Q;YN;Y;N', '', '') = MrYes then
        exit
      else DimensionsArticle.Dim.FocusDim(ItemDim, iChamp)
    else
    begin
      PGIInfo('Cet article est en rupture', Ecran.Caption);
      DimensionsArticle.Dim.FocusDim(ItemDim, iChamp);
    end;
  end;
  Result := EnRupture;
end;

procedure TOF_SELECTDIMDOC.TraiteParCombien(ItemDim: THDimensionItem);
var NomChamp: string;
  QteSaisie, PCB, X: Double;
  TobSelect: TOB;
  iChamp, Champ: Integer;
  Neg: Boolean;
begin
  QteSaisie := 0;
  Champ := 1;
  if ItemDim = nil then exit;
  for iChamp := 1 to MaxDimChamp do
  begin
    NomChamp := DimensionsArticle.OldDimChamp[iChamp];
    if (NomChamp = 'GL_QTEFACT') or (NomChamp = 'GL_QTERELIQUAT') then
    begin
      QteSaisie := valeur(ItemDim.Valeur[iChamp]);
      Champ := iChamp;
    end;
  end;
  if QteSaisie = 0 then exit;
  TobSelect := TOB(ItemDim.data);
  PCB := TobSelect.GetValue('GA_PCB');
  // Prix par combien
  if ((PCB <> 0) and (QteSaisie <> 0) and (NatureDoc = 'VEN')) then
  begin
    Neg := (QteSaisie < 0);
    QteSaisie := Abs(QteSaisie);
    X := Arrondi(QteSaisie / PCB, 9);
    if Arrondi(X - Trunc(X), 6) <> 0 then
    begin
      X := Arrondi(X + 0.5, 0);
      X := X * PCB;
      if Neg then X := -X;
      QteSaisie := X;
      ItemDim.Valeur[Champ] := QteSaisie;
      THDimension(GetControl('FDIM')).DisplayFormat := FormateChamp(TobSelect.GetValue('GA_ARTICLE'), PCB);
    end;
  end;
end;

procedure TOF_SELECTDIMDOC.TraiteReliquat(ItemDim: THDimensionItem);
var IReliquat, iChamp: Integer;
  ValQteFact, AncQteReliq: Double;
  NomChamp: string;
  TobSelect, TobDoc: Tob;
begin
  IReliquat := 0;
  ValQteFact := 0.0;
  if ItemDim = nil then exit;
  TobSelect := TOB(ItemDim.data);
  if TobSelect = nil then exit;
  if LaTob <> nil then
  begin
    TobDoc := LaTob.findfirst(['GA_ARTICLE'], [TobSelect.GetValue('GA_ARTICLE')], false);
    AncQteReliq := TobDoc.GetValue('_ReliqDejaSaisi');
  end else AncQteReliq := 0;
  for iChamp := 1 to MaxDimChamp do
  begin
    NomChamp := DimensionsArticle.OldDimChamp[iChamp];
    if NomChamp = 'GL_QTEFACT' then ValQteFact := valeur(ItemDim.Valeur[iChamp])
    else
      if NomChamp = 'GL_QTERELIQUAT' then IReliquat := iChamp;
  end;
  if IReliquat <> 0 then
  begin
    if (ValQteFact > AncQteReliq) then ItemDim.Valeur[IReliquat] := 0
    else ItemDim.Valeur[IReliquat] := AncQteReliq - ValQteFact;
  end;
end;

{==============================================================================================}
{===================================== Procedure AGL ==========================================}
{==============================================================================================}

procedure AGLOnClickParamGrilleDoc(parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then TOTOF := TFVierge(F).LaTOF else exit;
  {parm[1] contient la nature des champs affichés : préférences article/stock/achat/vente/autre}
  if (TOTOF is TOF_SELECTDIMDOC) then
  begin
    ParamGrille(TOF_SELECTDIMDOC(TOTOF).DimensionsArticle, TOF_SELECTDIMDOC(TOTOF).NatureDoc, TOF_SELECTDIMDOC(TOTOF).NaturePiece);
    TOF_SELECTDIMDOC(TOTOF).AdapterFiche;
    TOF_SELECTDIMDOC(TOTOF).SetReadOnly;
    //TOF_SELECTDIMDOC(TOTOF).Consultation(True);
  end else exit;
end;

initialization
  registerclasses([TOF_SELECTDIMDOC]);
  RegisterAglProc('OnClickParamGrilleDoc', TRUE, 0, AGLOnClickParamGrilleDoc);
end.

{***********UNITE*************************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 13/08/2003
Modifié le ... : 13/08/2003
Description .. :
Mots clefs ... : ARTICLE;CATALOGUE;FACTURE;
*****************************************************************}

unit FactArticle ;

interface

uses {$IFDEF VER150} variants,{$ENDIF} HEnt1, UTOB, Ent1, EntGC, FactTOB, UtilArticle,
{$IFDEF EAGLCLIENT}
     Utileagl, MaineAgl,
{$ELSE}
     Fe_Main, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     SysUtils, HCtrls, LookUp, Classes ;

procedure CodesCataToCodesLigne(TOBCata, TOBArt: TOB; ArticleRef: string; ARow: integer;
                                TobPiece : TOB);
function  GetCodeArtUnique(TOBPiece: TOB; ARow: integer): string;
procedure GetCodeCataUnique(TOBPiece: TOB; ARow: integer; var RefCata, RefFour: string);
procedure CodesLigneToCodesArt(TOBL, TOBArt: TOB);
function  FindTOBArtRow(TOBPiece, TOBArticles: TOB; ARow: integer): TOB;
function  FindTOBCataRow(TOBPiece, TOBCatalogu: TOB; ARow: integer): TOB;
function  FindTOBArtSais(TOBArticles: TOB; RefSais: string): TOB;
function  FindTOBCataSais(TOBCatalogu: TOB; RefCata, RefFour: string): TOB;
function  ArticleAutorise(TOBPiece, TOBArticles: TOB; NaturePiece: string; ARow: integer): boolean;
Function  ArticleEnStock (TOBPiece, TOBArticles: TOB; NaturePiece: string; ARow: integer) : boolean;
function  ISLigneGerableCC(TOBPiece, TOBArticles: TOB; GereLot, GereSerie: boolean; Arow: integer): boolean;
function  CreerTOBCata(TOBParent: TOB): TOB;
function  InitTOBCata(TOBCatalogu: TOB; RefCata, RefFour: string): TOB;
function  CreerTOBArt(TOBParent: TOB): TOB;
function  GetArticleRecherche(GS: THGrid; TitreSel, NaturePieceG, DomainePiece, SelectFourniss: string; stWhere: string = ''): boolean;
function  GetCatalogueMul(GS: THGrid; TitreSel, NaturePieceG, DomainePiece, SelectFourniss: string): boolean;
function  GetCatalogueMulArticle(GS: THGrid; TitreSel, NaturePieceG: string): string;
function  CatalogueChoixArticleRef: string;
procedure AjouteCatalogueArt(TOBL, TOBA: TOB);
function AjusteQte_MiniMultiple(nQte, nQteMini, nQteMultiple : double) : double;
procedure ChargeNewDims(TOBConds, TOBArticles, TOBDim : TOB;
  StQuelDepot, Tiers : string; BTransfert : Boolean);

implementation

uses
  wCommuns,
  Math,
  BTPUtil,
  paramsoc;

procedure CodesCataToCodesLigne(TOBCata, TOBArt: TOB; ArticleRef: string; ARow: integer;
                                TobPiece : TOB);
var TOBL : TOB;
    No : integer ;
begin
  TOBL := GetTOBLigne(TOBPiece, ARow); if TOBL=Nil then Exit ;
  No:=TOBL.GetValue('GL_NUMORDRE') ; TOBL.InitValeurs; TOBL.PutValue('GL_NUMORDRE',No) ;
  InitLesSupLigne(TOBL);
  if TOBCata = nil then EXIT;
  TOBL.PutValue('GL_ENCONTREMARQUE', 'X');
  TOBL.PutValue('GL_REFARTSAISIE', TOBCata.GetValue('GCA_REFERENCE'));
  TOBL.PutValue('GL_ARTICLE', TOBCata.GetValue('GCA_ARTICLE'));
  TOBL.PutValue('GL_REFCATALOGUE', TOBCata.GetValue('GCA_REFERENCE'));
  TOBL.PutValue('GL_FOURNISSEUR', TOBCata.GetValue('GCA_TIERS'));
  TOBL.PutValue('_CONTREMARTREF', ArticleRef);
  TOBL.PutValue('GL_TYPEREF', 'CAT');
  if TobArt <> nil then
  begin
    TOBL.PutValue('GL_REFARTSAISIE', TOBArt.GetValue('REFARTSAISIE'));
    TOBL.PutValue('GL_CODEARTICLE', TOBArt.GetValue('GA_CODEARTICLE'));
    TOBL.PutValue('GL_REFARTBARRE', TOBArt.GetValue('REFARTBARRE'));
    TOBL.PutValue('GL_REFARTTIERS', TOBArt.GetValue('REFARTTIERS'));
    TOBL.PutValue('GL_TYPEREF', 'ART');
  end;
end;

function GetCodeArtUnique(TOBPiece: TOB; ARow: integer): string;
var TypeDim, CodeArticle: string;
begin
  TypeDim := VarToStr(GetChampLigne(TobPiece, 'GL_TYPEDIM', ARow));
  if TypeDim = 'GEN' then
  begin
    CodeArticle := VarToStr(GetChampLigne(TobPiece, 'GL_CODESDIM', ARow));
    Result := CodeArticleUnique2(CodeArticle, '');
  end else
    Result := VarToStr(GetChampLigne(TobPiece, 'GL_ARTICLE', ARow));
end;

procedure CodesLigneToCodesArt(TOBL, TOBArt: TOB);
begin
  TOBArt.PutValue('GA_ARTICLE', TOBL.GetValue('GL_ARTICLE'));
  TOBArt.PutValue('REFARTSAISIE', TOBL.GetValue('GL_REFARTSAISIE'));
  TOBArt.PutValue('GA_CODEARTICLE', TOBL.GetValue('GL_CODEARTICLE'));
  TOBArt.PutValue('REFARTBARRE', TOBL.GetValue('GL_REFARTBARRE'));
  TOBArt.PutValue('REFARTTIERS', TOBL.GetValue('GL_REFARTTIERS'));
end;

procedure GetCodeCataUnique(TOBPiece: TOB; ARow: integer; var RefCata, RefFour: string);
var VenteAchat, EnContremarque: string;
  MouvExContrem: boolean;
begin
  VenteAchat := GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_VENTEACHAT');
  RefCata := GetChampLigne(TobPiece, 'GL_REFCATALOGUE', ARow);
  EnContremarque := GetChampLigne(TobPiece, 'GL_ENCONTREMARQUE', ARow);
  MouvExContrem := ((TobPiece.GetValue('GP_NATUREPIECEG') = 'EEX') or (TobPiece.GetValue('GP_NATUREPIECEG') = 'SEX'))
    and (EnContremarque = 'X');

  if (VenteAchat = 'ACH') or (MouvExContrem) then
    RefFour := GetChampLigne(TobPiece, 'GL_TIERS', ARow)
  else
    RefFour := GetChampLigne(TobPiece, 'GL_FOURNISSEUR', ARow);
end;

function FindTOBArtRow(TOBPiece, TOBArticles: TOB; ARow: integer): TOB;
var RefUnique: string;
begin
  Result := nil;
  RefUnique := GetCodeArtUnique(TobPiece, ARow);
  if RefUnique = '' then Exit;
  Result := TOBArticles.FindFirst(['GA_ARTICLE'], [RefUnique], False);
end;

function FindTOBCataRow(TOBPiece, TOBCatalogu: TOB; ARow: integer): TOB;
var RefCata, RefFour: string;
begin
  Result := nil;
  GetCodeCataUnique(TobPiece, ARow, RefCata, RefFour);
  if (Trim(RefCata) = '') or (Trim(RefFour) = '') then Exit;
  Result := TOBCatalogu.FindFirst(['GCA_REFERENCE', 'GCA_TIERS'], [RefCata, RefFour], False);
end;

function FindTOBArtSais(TOBArticles: TOB; RefSais: string): TOB;
var TOBArt: TOB;
begin
  Result := nil;
  if RefSais = '' then Exit;
  TOBArt := TOBArticles.FindFirst(['REFARTSAISIE'], [RefSais], False);
  if TOBArt <> nil then
  begin
    if TOBArt.GetValue('GA_STATUTART') <> 'UNI' then TOBArt := nil;
  end;
  if TOBArt = nil then
  begin
    TOBArt := TOBArticles.FindFirst(['GA_ARTICLE'], [RefSais], False);
  end;
  if TOBArt = nil then TOBArt := TOBArticles.FindFirst(['REFARTBARRE'], [RefSais], False);
  if TOBArt = nil then TOBArt := TOBArticles.FindFirst(['REFARTTIERS'], [RefSais], False);
  Result := TOBArt;
end;

function FindTOBCataSais(TOBCatalogu: TOB; RefCata, RefFour: string): TOB;
begin
  Result := nil;
  if (Trim(RefCata) = '') or (Trim(RefFour) = '') then Exit;
  Result := TOBCatalogu.FindFirst(['GCA_REFERENCE', 'GCA_TIERS'], [RefCata, RefFour], False);
end;

//FV1 - 10/01/2018 : FS#2806 - DELABOUDINIERE - Avertissement si utilisation d'un article non tenu en stock en saisie livraison
Function ArticleEnStock(TOBPiece, TOBArticles: TOB; NaturePiece: string; ARow: integer) : boolean;
Var TOBA      : TOB;
    RefUnique : string;
    IsStock   : Boolean;
Begin

  Result := True;

  //FV1 : 07/02/2018 - FS#2854 - Report V9 : Avertissement si utilisation d'un article non tenu en stock en saisie livraison
  If Not GetParamSocSecur('SO_ARTNONTENUESTOCK', False) then Exit;

  if (NaturePiece <> 'LBT') then Exit;

  TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  if TOBA = nil then Exit;

  RefUnique := TOBA.GetValue('GA_ARTICLE');
  if RefUnique = '' then Exit;

  IsStock := (TOBA.GetValue('GA_TENUESTOCK') = 'X');
  if (NaturePiece = 'LBT') And (Not IsStock) then
  begin
    Result := False;
    Exit;
  end;

end;

function ArticleAutorise(TOBPiece, TOBArticles: TOB; NaturePiece: string; ARow: integer): boolean;
var TOBA, TOBNat, TOBG: TOB;
  RefUnique: string;
  isFerme: boolean;
begin
  Result := True;
  TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  if TOBA = nil then Exit;
  RefUnique := TOBA.GetValue('GA_ARTICLE');
  if RefUnique = '' then Exit;

  isFerme := (TOBA.GetValue('GA_FERME') = 'X');
  if not isFerme then Exit;
  //

  // FQ Mode 10371 AC
  if TOBA.GetValue('GA_STATUTART')='DIM' then
  begin
    RefUnique := CodeArticleUnique2(TOBA.GetValue('GA_CODEARTICLE'),'') ;
  end ;

  TOBNat := VH_GC.TOBParPiece.FindFirst(['GPP_NATUREPIECEG'], [NaturePiece], False);
  if TOBNat = nil then Exit;

  TOBG := TOBNat.FindFirst(['GAP_NATUREPIECEG', 'GAP_ARTICLE'], [NaturePiece, RefUnique], False);
  if TOBG <> nil then Result := (TOBG.GetValue('GAP_SUSPENSION') = 'X') else Result := False;

end;

function ISLigneGerableCC(TOBPiece, TOBArticles: TOB; GereLot, GereSerie: boolean; Arow: integer): boolean;
var TOBA, TOBL: TOB;
		StockPlusMoins : string;
begin
  result := false;
  TOBL := GetTOBLigne(TOBPiece, Arow);
  if TOBL = nil then Exit;
  StockPlusMoins := getInfoParPiece (TOBPiece.getValue('GP_NATUREPIECEG'),'GPP_QTEPLUS')+
  									getInfoParPiece (TOBPiece.getValue('GP_NATUREPIECEG'),'GPP_QTEMOINS');
  if (TOBL.FieldExists('SUPPRIME')) and (TOBL.GetValue('SUPPRIME') = 'X') then Exit;
  if (TOBL.FieldExists('GL_ENCONTREMARQUE')) and (TOBL.GetValue('GL_ENCONTREMARQUE') = 'X') then exit;
  TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  if TOBA <> nil then
  begin
    if (TOBA.GetValue('GA_TENUESTOCK') = 'X') and (Pos('PHY',StockPlusMoins)> 0) then exit;
    if GereLot then
    begin
      if TOBA.GetValue('GA_LOT') <> 'X' then Exit;
    end;
    if (GereSerie) then
    begin
      if (TOBA.GetValue('GA_NUMEROSERIE') = 'X') then Exit;
    end;
  end;
  result := true;
end;

function CreerTOBCata(TOBParent: TOB): TOB;
var TOBCata: TOB;
begin
  TOBCata := TOB.Create('CATALOGU', TOBParent, -1);
  // ??
  TOBCata.AddChampSup('UTILISE', False);
  TOBCata.PutValue('UTILISE', '-');
  Result := TOBCata;
end;

function InitTOBCata(TOBCatalogu: TOB; RefCata, RefFour: string): TOB;
var
  QQ: TQuery;
begin
  Result := FindTOBCataSais(TOBCatalogu, RefCata, RefFour);
  if Result = nil then
  begin
    QQ := OpenSQL('Select * from CATALOGU Where GCA_REFERENCE="' + RefCATA + '" AND GCA_TIERS="' + RefFour + '"', True,-1, '', True);
    if not QQ.EOF then
    begin
      Result := CreerTOBCata(TOBCatalogu);
      Result.SelectDB('', QQ);
    end;
    Ferme(QQ);
  end;
end;

function CreerTOBArt(TOBParent: TOB): TOB;
var TOBArt: TOB;
begin
  TOBArt := TOB.Create('ARTICLE', TOBParent, -1);
  AddChampsSupTable (TOBARt,'GA2');
	if not TOBArt.FieldExists ('BNP_TYPERESSOURCE') then
  begin
  	TOBArt.AddChampSupValeur ('BNP_TYPERESSOURCE','');
  end;
	if not TOBArt.FieldExists ('BNP_LIBELLE') then
  begin
  	TOBArt.AddChampSupValeur ('BNP_LIBELLE','');
  end;
  TOBArt.AddChampSup('REFARTBARRE', False);
  TOBArt.PutValue('REFARTBARRE', '');
  TOBArt.AddChampSup('REFARTTIERS', False);
  TOBArt.PutValue('REFARTTIERS', '');
  TOBArt.AddChampSup('REFARTSAISIE', False);
  TOBArt.PutValue('REFARTSAISIE', '');
  TOBArt.AddChampSup('UTILISE', False);
  TOBArt.PutValue('UTILISE', '-');
  TOBArt.AddChampSupValeur('SUPPRIME', '-');
  TOBArt.AddChampSupValeur('LIBELLEFOU','');
  Result := TOBArt;
end;

function GetArticleLookUp(GS: THGrid; TitreSel, NaturePieceG, DomainePiece, SelectFourniss: string;
  stWhere: string = ''): boolean;
var sWhere, sArt: string;
begin
  sArt := 'GA_STATUTART<>"DIM" AND GA_TYPEARTICLE<>"FI"';
  sWhere := FabricWhereNatArt(NaturePieceG, DomainePiece, SelectFourniss);
  if sWhere <> '' then sArt := sArt + ' AND ' + sWhere;
  Result := LookupList(GS, TitreSel, 'ARTICLE', 'GA_ARTICLE', 'GA_LIBELLE', sArt + stWhere, 'GA_ARTICLE', True, 7);
end;

function GetArticleMul(GS: THGrid; TitreSel, NaturePieceG, DomainePiece, SelectFourniss: string; stWhere: string = ''): boolean;
var sw, sWhere, CodeArt, StChamps : string;
{$IFNDEF BTP}
	 SNatArt: string;
{$ENDIF}
begin
  Result := False;
  CodeArt := GS.Cells[GS.Col, GS.Row];
  //sWhere := '';
  {$IFDEF BTP}
  sW := FabricWhereNatArt(NaturePieceG, DomainePiece, SelectFourniss);
  if sw <> '' then sWhere := stWhere + ' AND (' + sw + ')';
  if CodeArt <> '' then StChamps := 'GA_CODEARTICLE=' + Trim(Copy(CodeArt, 1, 18)) + ';XX_WHERE=' + sWhere
                   else StChamps := 'XX_WHERE=' + sWhere;

  CodeArt := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', StChamps+';RECHERCHEARTICLE;MULTISELECTION');

  {$ELSE}
  sW := FabricWhereNatArt(NaturePieceG, DomainePiece, SelectFourniss);
  if (GetInfoParpiece(NaturePieceG, 'GPP_PILOTEORDRE') = 'X') then
  begin

    sW := sW + ' AND EXISTS' +
      ' (SELECT WAN_CODEARTICLE FROM WARTNAT LEFT JOIN WNATURETRAVAIL ON WAN_NATURETRAVAIL=WNA_NATURETRAVAIL ' +
      ' WHERE WAN_CODEARTICLE = GA_CODEARTICLE' +
      ' AND WNA_NATUREPIECEA IN (SELECT GPP_NATUREPIECEG FROM PARPIECE WHERE GPP_CONTEXTES like "GP%" )) ';
  end;
  if (ctxMode in V_PGI.PGIContexte) then
  begin
    // Restriction possible de l'affichage des articles d'un seul fournisseur
    SNatArt := GetInfoParPiece(NaturePieceG, 'GPP_TYPEARTICLE');
    while SNatArt <> '' do
    begin
      if sWhere <> '' then sWhere := sWhere + ' OR ';
      sWhere := sWhere + '(CO_CODE="' + ReadTokenSt(SNatArt) + '")';
    end;
    sWhere := 'TYPEARTICLE=' + sWhere;
    sWhere := sWhere + ';GA_CODEARTICLE=' + Trim(Copy(CodeArt, 1, 18));
    sWhere := sWhere + ';XX_WHERE_=' + sw;
    if SelectFourniss <> '' then sWhere := sWhere + ';EXCLUSIF;GA_FOURNPRINC=' + SelectFourniss;
    {$IFDEF CHR}
    StChamps := '';
    StChamps := 'GA_CODEARTICLE=' + Trim(Copy(CodeArt, 1, 18));
    CodeArt := AGLLanceFiche('H', 'HRPRESTATIONS_MUL', StChamps, '', 'SELECTION;NOFILTRE');
    {$ELSE}
    CodeArt := DispatchArtMode(1, '', '', 'SELECTION;' + sWhere);
    {$ENDIF}
  end
  else
  begin
    if sw <> '' then sWhere := sWhere + ' AND ' + sw;
    StChamps := 'GA_CODEARTICLE=' + Trim(Copy(CodeArt, 1, 18)) + ';XX_WHERE=' + sWhere + stWhere;
    if ctxAffaire in V_PGI.PGIContexte then CodeArt := AGLLanceFiche('AFF', 'AFARTICLE_RECH', '', '', StChamps)
    else CodeArt := AGLLanceFiche('GC', 'GCARTICLE_RECH', '', '', StChamps);
  end;
  {$ENDIF}
  if CodeArt <> '' then
  begin
    Result := True;
    GS.Cells[GS.Col, GS.Row] := CodeArt;
  end;
end;

function GetArticleRecherche(GS: THGrid; TitreSel, NaturePieceG, DomainePiece, SelectFourniss: string; stWhere: string = ''): boolean;
begin
  if GetParamSocSecur('SO_GCRECHARTAV', False) then
    Result := GetArticleMul(GS, TitreSel, NaturePieceG, DomainePiece, SelectFourniss, stWhere)
  else
    Result := GetArticleLookUp(GS, TitreSel, NaturePieceG, DomainePiece, SelectFourniss, stWhere);
end;

function GetCatalogueMul(GS: THGrid; TitreSel, NaturePieceG, DomainePiece, SelectFourniss: string): boolean;
var sWhere, CodeArt, StChamps: string;
begin
  Result := False;
  CodeArt := GS.Cells[GS.Col, GS.Row];
  sWhere := '';
  //sW:=FabricWhereNatArt(NaturePieceG,DomainePiece,SelectFourniss) ; if sw<>'' then sWhere:=sWhere+' AND '+sw ;
  StChamps := 'GCA_REFERENCE=' + Trim(Copy(CodeArt, 1, 18)) + ';XX_WHERE=' + sWhere;
  CodeArt := AGLLanceFiche('GC', 'GCMULCATALOG_RECH', '', '', StChamps);

  if CodeArt <> '' then
  begin
    Result := True;
    GS.Cells[GS.Col, GS.Row] := CodeArt;
  end;
end;

function GetCatalogueMulArticle(GS: THGrid; TitreSel, NaturePieceG: string): string;
var StChamps: string;
begin
  // Il faut adapter le mul pour ne pas pouvoir changer le code article
  StChamps := 'GCA_ARTICLE=' + GS.Cells[SG_RefArt, GS.Row];
  Result := AGLLanceFiche('GC', 'GCMULCATALOG_RECH', '', '', StChamps + ';MULARTICLE');
end;

function CatalogueChoixArticleRef: string;
begin
  Result := AGLLanceFiche('GC', 'GCCONTREM_ARTREF', '', '', '');
end;

procedure AjouteCatalogueArt(TOBL, TOBA: TOB);
var Q: TQuery;
  i: integer;
  Nam, req : string;
begin
  if TOBA = nil then Exit;
  if TOBL = nil then Exit;
  req := 'SELECT * FROM CATALOGU WHERE GCA_TIERS="' + TOBL.GetValue('GL_TIERS') + '" AND GCA_ARTICLE="' + TOBL.GetValue('GL_ARTICLE') + '" ' +
          'AND GCA_DATESUP>="' + USDateTime(TOBL.parent.GetValue('GP_DATEPIECE')) + '" AND GCA_DATEREFERENCE<="' + USDateTime(TOBL.parent.GetValue('GP_DATEPIECE')) + '" ORDER BY GCA_DATEREFERENCE DESC';
  Q := OpenSQL(req, True,-1, '', True);
  if not Q.EOF then
  begin
    for i := 0 to Q.FieldCount - 1 do
    begin
      Nam := Q.Fields[i].FieldName;
      TOBA.AddChampSup(Nam, False);
      TOBA.PutValue(Nam, Q.Fields[i].AsVariant);
    end;
    TOBA.AddChampSup('DPANEW', False);
    TOBA.PutValue('DPANEW', TOBA.GetValue('GCA_DPA'));
  end;
  Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 24/09/2003
Modifié le ... :   /  /
Description .. :
Suite ........ : Ajuste la quantité d'une pièce en fonction
Suite ........ : * minimum de pièce (quantité économique)
Suite ........ : * multiple de (quantité de conditionnement)
Mots clefs ... : _PCB;_QECO;_QPCB
*****************************************************************}
function AjusteQte_MiniMultiple(nQte, nQteMini, nQteMultiple : double) : double;
var
  iSens : integer;
begin
  //On mémorise le signe de la qté
  if (nQte>=0) then iSens:=+1 else iSens:=-1;
  nQte:=abs(nQte);

  // Prise en compte du Minimum de commande (Qté économique de vente)
  if (nQte<=nQteMini) then
    nQte := nQteMini;

  // Prise en compte du Multiple de vente (Qté conditionnement)
  if (nQteMultiple>0) then
    nQte := Max(nQteMultiple,nQteMultiple*wArrondir(nQte/nQteMultiple, tApUnite, tAMInferieure));

  //On redonne le bon signe à la qté
  nQte:=nQte*iSens;

  Result := nQte;
end;

function CreerInArticle(Const TOBDim : TOB) : String;
var NbArtDim, i : integer;
    ListeArticles : TStringList;
    StArticle, StIn : string;
begin
  ListeArticles := TStringList.Create;
  ListeArticles.Sorted := True;
  for NbArtDim := 0 to TOBDim.Detail.Count - 1 do
  begin
    StArticle := TOBDim.Detail[NbArtDim].GetValue('GA_ARTICLE');
    if Not ListeArticles.Find(StArticle, i) then ListeArticles.Add(StArticle);
  end;
  StIn := '';
  For i := 0 to ListeArticles.Count - 1 do
  begin
    if StIn <> '' then StIn := StIn + ',';
    StIn := StIn + '"' + ListeArticles.Strings[i] + '"';
  end;
  Result := StIn;
end;

procedure ChargeNewDims(TOBConds, TOBArticles, TOBDim : TOB;
  StQuelDepot, Tiers : string; BTransfert : Boolean);
var TobTempArt, TobDispo, TobTempCata : TOB;
  TobTempCataArt, TobDispoArt, TOBItem, TobTempArtDim : TOB;
  SQLArticle, SQLDispo, SQLCondi, SQLCata : String;
  ListeArticles, RefUnique, NomChamp : String;
  iDim, iChamp : Integer;
begin
  ListeArticles := CreerInArticle(TOBDim);
  if ListeArticles = '' then exit;
  SQLArticle := 'Select * from Article LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE=GA_ARTICLE where GA_ARTICLE IN (' + ListeArticles + ')';
  //Chargement de tous les dépôts en multi-dépôts sinon chargement du dépôt courant
  //et si transfert, chargement du dépôt récepteur en plus du dépôt émetteur
  if StQuelDepot <> '' then StQuelDepot := ' AND GQ_DEPOT IN (' + StQuelDepot + ')';
  SQLDispo := 'Select * from Dispo where GQ_ARTICLE IN (' + ListeArticles
  + ') AND GQ_CLOTURE="-"' + StQuelDepot;
  SQLCondi := 'SELECT * FROM CONDITIONNEMENT WHERE GCO_ARTICLE IN ('
  + ListeArticles + ') ORDER BY GCO_NBARTICLE';
  SQLCata := 'SELECT * FROM CATALOGU WHERE GCA_TIERS="' + Tiers
  + '" AND GCA_ARTICLE IN (' + ListeArticles + ') ORDER BY GCA_ARTICLE';
  TobTempArt := TOB.Create('', nil, -1);
  TobDispo := TOB.Create('', nil, -1);
  TobTempCata := TOB.Create('', nil, -1);
  TobTempArt.LoadDetailDBFromSQL('ARTICLE', SQLArticle);
  TobDispo.LoadDetailDBFromSQL('DISPO', SQLDispo);
  TobTempCata.LoadDetailDBFromSQL('CATALOGU', SQLCata);
  TOBConds.LoadDetailDBFromSQL('CONDITIONNEMENT', SQLCondi, True);
  for iDim := 0 to TOBDim.Detail.Count - 1 do
  begin
    TOBItem := TOBDim.Detail[iDim];
    if TOBItem = nil then exit;
    RefUnique := TOBItem.GetValue('GA_ARTICLE');
    if TobTempArt = nil then TobTempArtDim := nil
    else TobTempArtDim := TobTempArt.FindFirst(['GA_ARTICLE'], [RefUnique], False);
    if (TobTempArtDim <> nil) and (TOBArticles.FindFirst(['GA_ARTICLE'], [RefUnique], False) = nil) then
    begin
      With TobTempArtDim do
      begin
        AddChampSupValeur('REFARTBARRE', '', False);
        AddChampSupValeur('REFARTTIERS', '', False);
        AddChampSupValeur('REFARTSAISIE', '', False);
        AddChampSupValeur('UTILISE', '-', False);
        Changeparent(TOBArticles, -1);
        if FindFirst(['GQ_ARTICLE'], [RefUnique], False) = nil then
        begin
          if TobDispo <> nil then
          begin
            TobDispoArt := TobDispo.FindFirst(['GQ_ARTICLE'], [RefUnique], False);
            if TobDispoArt <> nil then TobDispoArt.Changeparent(TobTempArtDim, -1);
            if BTransfert then
            begin
              TobDispoArt := TobDispo.FindNext(['GQ_ARTICLE'], [RefUnique], False);
              if TobDispoArt <> nil then TobDispoArt.Changeparent(TobTempArtDim, -1);
            end;
          end;
        end;
      end;
    end;
    //Affecte les catalogues aux articles sélectionnés
    if TobTempCata <> nil then
    begin
      TobTempCataArt := TobTempCata.FindFirst(['GCA_ARTICLE'], [RefUnique], False);
      while TobTempCataArt <> nil do
      begin
        for iChamp := 0 to TobTempCataArt.NbChamps do
        begin
          NomChamp := TobTempCataArt.GetNomChamp(iChamp);
          TobTempArtDim.AddChampSupValeur(NomChamp, TobTempCataArt.GetValue(NomChamp));
        end;
        TobTempArtDim.AddChampSupValeur('DPANEW', TobTempArtDim.GetValue('GCA_DPA'));
        TobTempCataArt := TobTempCata.FindNext(['GCA_ARTICLE'], [RefUnique], False);
      end;
    end;
  end;
  TobTempArt.Free;
  TobDispo.Free;
  TobTempCata.Free ;
end;


end.

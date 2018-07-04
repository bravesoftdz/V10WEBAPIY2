{***********UNITE*************************************************
Auteur  ...... : Dominique BROSSET
Créé le ...... : 14/08/2003
Modifié le ... :   /  /    
Description .. : Fonctions utilitaires pour la gestion des lots et des séries 
Suite ........ : dans les pièces
Mots clefs ... : PIECE;LOT;SERIE;UTIL
*****************************************************************}
unit FactLotSerie ;

interface

uses HEnt1, UTOB, HCtrls, LookUp, SysUtils,
     FactTob, FactNomen,
     StockUtil,
{$IFDEF EAGLCLIENT}
     Utileagl, MaineAgl,
{$ELSE}
     Fe_Main, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     AGLInitGC, EntGC, UtilPGI,uEntCommun,UtilTOBPiece ;

{============================= LOTS ==================================}

procedure MAJDispoLot(TOBArticles, TOBL, TOBDL: TOB; VenteAchat: string; DepotEmetteur: string = ''; Inverse: boolean = false);
function DateLotDepotEmetteur(RefUnique, DepotEmetteur, NumLot: string): TDateTime;
function CreerNoeudLot(TOBPiece, TOBDesLots: TOB): TOB;
procedure LoadLesLots(TOBPiece, TOBDesLots: TOB; CleDoc: R_CleDoc);
procedure UpdateNoeudLot(TOBL, TOBDL: TOB);

procedure InverseStockLot(TOBPiece_O,TOBLOT_O,TOBArticles : TOB) ; { NEWPIECE }
procedure InverseStockLotTransfo(TOBPiece_O, TobPiece, TOBLOT_O, TOBArticles: TOB); { NEWPIECE }

{============================ SERIES ==================================}

procedure DetruitSerie(ARow: integer; Action : TActionFiche; GereSerie : boolean; TobPiece, TobSerie : TOB);
procedure MAJDispoSerie(TOBL, TOBSerLig: TOB; Inverse: Boolean);
procedure LoadLesSeries(TOBPiece, TOBSerie : TOB; CleDoc: R_CleDoc ; TransfoPiece : boolean = False);

procedure UpdateLesSeries(TOBL, TOBSerLig: TOB);
procedure InitialiseReliquatSerie (TobSerRel, TobSerie_O : TOB);

procedure InverseStockSerie(TOBPiece_O, TOBSerie_O: TOB);
procedure InverseStockSerieTransfo(TOBPiece_O, TobPiece, TOBSerie_O: TOB);

procedure RazLesSeries(TobSerie, TobSerie_O : TOB);
procedure RazAndSaveLesSeries(TobSerie, TobSerie_O : TOB);
procedure ReaffecteLesReliquatSerie(TobSerRel, TobPiece_O : TOB);

function QuelQteEnSerie(TobSerie, TobL : TOB ; NbCompoSerie : double) : double;

implementation

{============================= LOTS ==================================}

procedure DetruitSerie(ARow: integer; Action : TActionFiche; GereSerie : boolean; TobPiece, TobSerie : TOB);
var TOBL, TOBLS: TOB;
  IndiceSerie, i_ind: integer;
begin
  {$IFDEF CCS3}
  Exit;
  {$ELSE}
  if Action = taConsult then Exit;
  if not GereSerie then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if TOBL.GetValue('GL_NUMEROSERIE') <> 'X' then Exit;
  IndiceSerie := TOBL.GetValue('GL_INDICESERIE');
  if IndiceSerie <= 0 then Exit;
  TOBL.PutValue('GL_INDICESERIE', 0);
  for i_ind := TOBSerie.Detail[IndiceSerie - 1].Detail.Count - 1 downto 0 do
  begin
    TOBLS := TOBSerie.Detail[IndiceSerie - 1].Detail[i_ind];
    if i_ind = 0 then TOBLS.PutValue('GLS_IDSERIE', '')
    else TOBLS.Free;
  end;
  {$ENDIF}
end;

{ NEWPIECE }

procedure MAJDispoLot(TOBArticles, TOBL, TOBDL: TOB; VenteAchat: string; DepotEmetteur: string = ''; Inverse: boolean = false);
var TOBA, TOBDepot, TOBDepotLot: TOB;
  Depot, RefUnique, NaturePiece, NumLot: string;
  Qte, RatioStock: Double;
  i, iAvoir: Integer;
begin
  Depot := TOBL.GetValue('GL_DEPOT');
  RefUnique := TOBL.GetValue('GL_ARTICLE');
  if RefUnique = '' then Exit;
  NaturePiece := TOBL.GetValue('GL_NATUREPIECEG');
  RatioStock := GetRatio(TOBL, nil, trsStock);
  TOBA := TOBArticles.FindFirst(['GA_ARTICLE'], [RefUnique], False);
  TOBDepot := TOBA.FindFirst(['GQ_DEPOT'], [Depot], False);
  if not Inverse then
  begin
    for i := 0 to TOBDL.Detail.Count - 1 do
    begin
      NumLot := TOBDL.Detail[i].GetValue('GLL_NUMEROLOT');
      TOBDepotLot := TOBDepot.FindFirst(['GQL_NUMEROLOT'], [NumLot], False);
      if TOBDepotLot = nil then
      begin
        TOBDepotLot := TOB.Create('DISPOLOT', TOBDepot, -1);
        if TOBDL.Detail[i].FieldExists('DATELOT') then TOBDepotLot.PutValue('GQL_DATELOT', TOBDL.Detail[i].GetValue('DATELOT'))
        else TOBDepotLot.PutValue('GQL_DATELOT', iDate1900);
        TOBDepotLot.PutValue('GQL_ARTICLE', RefUnique);
        TOBDepotLot.PutValue('GQL_DEPOT', Depot);
        TOBDepotLot.PutValue('GQL_NUMEROLOT', NumLot);
        if NaturePiece = 'TRE' then TOBDepotLot.PutValue('GQL_DATELOT', DateLotDepotEmetteur(RefUnique, DepotEmetteur, NumLot));
        TOBDepotLot.PutValue('GQL_PHYSIQUE', 0);
        TOBDepotLot.AddChampSupValeur('NEW_ENREG', 'X');
      end;
      Qte := TOBDepotLot.GetValue('GQL_PHYSIQUE');
      if (NaturePiece = 'TEM') or (VenteAchat = 'VEN') or (NaturePiece = 'SEX') then
        Qte := Qte - TOBDL.Detail[i].GetValue('GLL_QUANTITE') / RatioStock
      else if (NaturePiece = 'TRE') or (VenteAchat = 'ACH') or (NaturePiece = 'EEX') then
        Qte := Qte + TOBDL.Detail[i].GetValue('GLL_QUANTITE') / RatioStock;
      Qte := Arrondi(Qte, V_PGI.OkDecQ);
      TOBDepotLot.PutValue('GQL_PHYSIQUE', Qte);
    end;
    //if Not TOBDepot.InsertOrUpdateDB(False) then V_PGI.IoError:=oeUnknown ; //desactivé 20/03/2003
  end else
  begin
    if GetInfoParPiece(NaturePiece, 'GPP_ESTAVOIR') = 'X' then iAvoir := -1 else iAvoir := 1;
    for i := 0 to TOBDL.Detail.Count - 1 do
    begin
      TOBDepotLot := TOBDepot.FindFirst(['GQL_NUMEROLOT'], [TOBDL.Detail[i].GetValue('GLL_NUMEROLOT')], true);
      if TOBDepotLot = nil then continue;
      Qte := (TOBDL.Detail[i].GetValue('GLL_QUANTITE') * iAvoir) / RatioStock;
      if (NaturePiece = 'TEM') or (VenteAchat = 'VEN') or (NaturePiece = 'SEX') then
        Qte := TOBDepotLot.GetValue('GQL_PHYSIQUE') + Qte
      else if (NaturePiece = 'TRE') or (VenteAchat = 'ACH') or (NaturePiece = 'EEX') then
        Qte := TOBDepotLot.GetValue('GQL_PHYSIQUE') - Qte;
      Qte := Arrondi(Qte, V_PGI.OkDecQ);
      TOBDepotLot.PutValue('GQL_PHYSIQUE', Qte);
    end;
  end;
end;

function DateLotDepotEmetteur(RefUnique, DepotEmetteur, NumLot: string): TDateTime;
var Q: TQuery;
begin
  Result := iDate1900;
  Q := OpenSQL('SELECT GQL_DATELOT FROM DISPOLOT WHERE GQL_ARTICLE="' + RefUnique +
    '" AND GQL_DEPOT="' + DepotEmetteur + '" AND GQL_NUMEROLOT="' + NumLot + '"', True,-1, '', True);
  if not Q.EOF then Result := Q.FindField('GQL_DATELOT').AsDateTime;
  Ferme(Q);
end;

procedure _InverseStockLot(TOBPiece_O, TobPiece, TOBLOT_O, TOBArticles: TOB; Transfo: Boolean);
var
  i, IndiceLot: integer;
  TOBDL, TOBL: TOB;
  NaturePiece, VenteAchat: string;
begin
  {$IFDEF CCS3}
  Exit;
  {$ELSE}
  NaturePiece := TOBPiece_O.GetValue('GP_NATUREPIECEG');
  if GetInfoParPiece(NaturePiece, 'GPP_LOT') <> 'X' then Exit;
  if TOBLOT_O.Detail.Count <= 0 then Exit;
  TOBLOT_O.Detail[0].AddChampSup('UTILISE', True);
  VenteAchat := GetInfoParPiece(NaturePiece, 'GPP_VENTEACHAT');
  for i := 0 to TOBPiece_O.Detail.Count - 1 do
  begin
    IndiceLot := GetChampLigne(TOBPiece_O, 'GL_INDICELOT', i + 1);
    if IndiceLot > 0 then
    begin
      TOBL := TOBPiece_O.Detail[i];
      TOBDL := TOBLOT_O.Detail[IndiceLot - 1];
      TOBDL.PutValue('UTILISE', 'X');
      MAJDispoLot(TOBArticles, TOBL, TOBDL, VenteAchat, '', True);
    end;
  end;
  {$ENDIF}
end;

{ NEWPIECE }

procedure InverseStockLot(TOBPiece_O, TOBLOT_O, TOBArticles: TOB);
begin
  _InverseStockLot(TOBPiece_O, nil, TOBLOT_O, TobArticles, False);
end;

{ NEWPIECE }

procedure InverseStockLotTransfo(TOBPiece_O, TobPiece, TOBLOT_O, TOBArticles: TOB);
begin
  _InverseStockLot(TOBPiece_O, TobPiece, TOBLOT_O, TobArticles, True);
end;

function CreerNoeudLot(TOBPiece, TOBDesLots: TOB): TOB;
var TOBN: TOB;
begin
  TOBN := TOB.Create('', TOBDesLots, -1);
  TOBN.AddChampSup('NATUREPIECEG', False);
  TOBN.PutValue('NATUREPIECEG', TOBPiece.GetValue('GP_NATUREPIECEG'));
  TOBN.AddChampSup('SOUCHE', False);
  TOBN.PutValue('SOUCHE', TOBPiece.GetValue('GP_SOUCHE'));
  TOBN.AddChampSup('NUMERO', False);
  TOBN.PutValue('NUMERO', TOBPiece.GetValue('GP_NUMERO'));
  TOBN.AddChampSup('INDICEG', False);
  TOBN.PutValue('INDICEG', TOBPiece.GetValue('GP_INDICEG'));
  TOBN.AddChampSup('NUMLIGNE', False);
  TOBN.PutValue('NUMLIGNE', 0);
  TOBN.AddChampSup('RANG', False);
  TOBN.PutValue('RANG', 0);
  TOBN.AddChampSup('QUANTITE', False);
  TOBN.PutValue('QUANTITE', 0);
  Result := TOBN;
end;

procedure LoadLesLots(TOBPiece, TOBDesLots: TOB; CleDoc: R_CleDoc);
var QQ: TQuery;
  i, k, NumL, Indicelot: integer;
  CodeLot: string;
  TOBG, TOBLigne, TOBLoc, TOBNoeud: TOB;
begin
  if TOBPiece.GetValue('GP_ARTICLESLOT') <> 'X' then Exit;
  TOBG := TOB.Create('', nil, -1);
  QQ := OpenSQl('SELECT * FROM LIGNELOT WHERE ' + WherePiece(CleDoc, ttdLot, False), True,-1, '', True);
  if not QQ.EOF then TOBG.LoadDetailDB('LIGNELOT', '', '', QQ, False);
  Ferme(QQ);
  for i := TOBG.Detail.Count - 1 downto 0 do
  begin
    {Lecture des LigneLot "en vrac"}
    TOBLoc := TOBG.Detail[i];
    NumL := TOBLoc.GetValue('GLL_NUMLIGNE');
    CodeLot := TOBLoc.GetValue('GLL_NUMEROLOT');
    TOBLigne := TOBPiece.FindFirst(['GL_NUMLIGNE'], [NumL], False);
    {recherche de la ligne de pièce concernée}
    if TOBLigne <> nil then
    begin
      {Recherche du noeud des lots de la ligne, si pas ok le créer}
      TOBNoeud := nil;
      IndiceLot := 0;
      for k := 0 to TOBDesLots.Detail.Count - 1 do
      begin
        if TOBDesLots.Detail[k].GetValue('NUMLIGNE') = NumL then
        begin
          TOBNoeud := TOBDeslots.Detail[k];
          IndiceLot := k + 1;
          Break;
        end;
      end;
      if TOBNoeud = nil then
      begin
        TOBNoeud := CreerNoeudLot(TOBPiece, TOBDesLots);
        TOBNoeud.PutValue('NUMLIGNE', NumL);
        IndiceLot := TOBDesLots.Detail.Count;
        TOBNoeud.PutValue('QUANTITE', TobLigne.Getvalue('GL_QTEFACT')); //JS 25/03/03
      end;
      TOBLigne.PutValue('GL_INDICELOT', IndiceLot);
      TOBLoc.ChangeParent(TOBNoeud, -1);
    end;
  end;
  TOBG.Free;
end;

procedure UpdateNoeudLot(TOBL, TOBDL: TOB);
var NatureG, Souche, RefArticle: string;
  Numero, IndiceG, i, NumL: integer;
  Qte: Double;
  TOBDet: TOB;
begin
  if ((TOBL = nil) or (TOBDL = nil)) then Exit;
  NatureG := TOBL.GetValue('GL_NATUREPIECEG');
  Souche := TOBL.GetValue('GL_SOUCHE');
  Numero := TOBL.GetValue('GL_NUMERO');
  IndiceG := TOBL.GetValue('GL_INDICEG');
  RefArticle := TOBL.GetValue('GL_ARTICLE');
  Qte := TOBL.GetValue('GL_QTEFACT');
  NumL := TOBL.GetValue('GL_NUMLIGNE');
  TOBDL.PutValue('NATUREPIECEG', NatureG);
  TOBDL.PutValue('SOUCHE', Souche);
  TOBDL.PutValue('NUMERO', Numero);
  TOBDL.PutValue('INDICEG', Souche);
  TOBDL.PutValue('QUANTITE', Qte);
  for i := 0 to TOBDL.Detail.Count - 1 do
  begin
    TOBDet := TOBDL.Detail[i];
    TOBDet.PutValue('GLL_NATUREPIECEG', NatureG);
    TOBDet.PutValue('GLL_SOUCHE', Souche);
    TOBDet.PutValue('GLL_NUMERO', Numero);
    TOBDet.PutValue('GLL_INDICEG', IndiceG);
    TOBDet.PutValue('GLL_ARTICLE', RefArticle);
    TOBDet.PutValue('GLL_NUMLIGNE', NumL);
  end;
end;

{============================ SERIES ==================================}

procedure MAJDispoSerie(TOBL, TOBSerLig: TOB; Inverse: Boolean);
var TOBLS, TOBDispoSerie, TOBDS: TOB;
  Depot, RefUnique, NaturePiece, NumSerie: string;
  ColPlus, ColMoins, Sens, WhereSerie, St: string;
  EstAvoir, EntreeSerie, ResaCli, PrepaCli: Boolean;
  i: Integer;
  QQ: TQuery;
begin
  {$IFDEF CCS3}
  Exit;
  {$ELSE}
  Depot := TOBL.GetValue('GL_DEPOT');
  RefUnique := TOBL.GetValue('GL_ARTICLE');
  if RefUnique = '' then Exit;
  NaturePiece := TOBL.GetValue('GL_NATUREPIECEG');
  ColPlus := GetInfoParPiece(NaturePiece, 'GPP_QTEPLUS');
  ColMoins := GetInfoParPiece(NaturePiece, 'GPP_QTEMOINS');
  Sens := GetInfoParPiece(NaturePiece, 'GPP_SENSPIECE');
  EstAvoir := (GetInfoParPiece(NaturePiece, 'GPP_ESTAVOIR') = 'X');
  if Sens = 'ENT' then EntreeSerie := not (EstAvoir) else EntreeSerie := EstAvoir;
  TOBDispoSerie := TOB.Create('', nil, -1);
  if TOBDispoSerie = nil then
  begin
    V_PGI.IoError := oeUnknown;
    exit;
  end;
  if not Inverse then
  begin
    if EntreeSerie then
    begin // Entrée de numéro série
      for i := 0 to TOBSerLig.Detail.Count - 1 do
      begin
        TOBLS := TOBSerLig.Detail[i];
        if TOBLS.GetValue('GLS_TENUESTOCK') <> 'X' then Continue;
        NumSerie := TOBLS.GetValue('GLS_IDSERIE');
        TOBDS := TOB.Create('DISPOSERIE', TOBDispoSerie, -1);
        if TOBDS = nil then
        begin
          V_PGI.IoError := oeUnknown;
          TOBDispoSerie.free;
          exit;
        end;
        TOBDS.PutValue('GQS_IDSERIE', NumSerie);
        TOBDS.PutValue('GQS_ARTICLE', RefUnique);
        TOBDS.PutValue('GQS_DEPOT', Depot);
        TOBDS.PutValue('GQS_NUMEROLOT', TOBLS.GetValue('GLS_NUMEROLOT'));
        TOBDS.PutValue('GQS_ENRESERVECLI', '-');
        TOBDS.PutValue('GQS_ENPREPACLI', '-');
      end;
      if TOBDispoSerie.Detail.Count > 0 then
        if not TOBDispoSerie.InsertDBByNivel(False) then V_PGI.IoError := oeUnknown;
    end else
    begin // Sortie de numéro série
      WhereSerie := '';
      St := '';
      for i := 0 to TOBSerLig.Detail.Count - 1 do
      begin
        TOBLS := TOBSerLig.Detail[i];
        if TOBLS.GetValue('GLS_TENUESTOCK') <> 'X' then Continue;
        NumSerie := TOBLS.GetValue('GLS_IDSERIE');
        St := 'GQS_IDSERIE="' + NumSerie + '" ';
        if WhereSerie = '' then WhereSerie := '(GQS_ARTICLE="' + TOBLS.GetValue('GLS_ARTICLE') + '" AND ' + St + ')'
        else WhereSerie := WhereSerie + ' OR ' + '(GQS_ARTICLE="' + TOBLS.GetValue('GLS_ARTICLE') + '" AND ' + St + ')';
      end;
      if WhereSerie <> '' then
      begin
        QQ := OpenSQl('Select * From DISPOSERIE Where ' + WhereSerie, True,-1, '', True);
        if not QQ.EOF then TOBDispoSerie.LoadDetailDB('DISPOSERIE', '', '', QQ, False);
        Ferme(QQ);
        if TOBDispoSerie.Detail.Count > 0 then
        begin
          ResaCli := (Pos('RC', ColPlus) > 0);
          PrepaCli := (Pos('PRE', ColPlus) > 0);
          if (ResaCli) or (PrepaCli) then
          begin
            for i := 0 to TOBDispoSerie.Detail.Count - 1 do
            begin
              if (ResaCli) then TOBDispoSerie.Detail[i].PutValue('GQS_ENRESERVECLI', 'X')
              else TOBDispoSerie.Detail[i].PutValue('GQS_ENPREPACLI', 'X');
            end;
            if not TOBDispoSerie.UpdateDB(False) then V_PGI.IoError := oeUnknown;
          end else if not TOBDispoSerie.DeleteDB(False) then V_PGI.IoError := oeUnknown;
        end;
      end;
    end;
  end else
  begin // Inversion gestion stock série
    if EntreeSerie then
    begin // Inversion entrée de numéro série
      WhereSerie := '';
      St := '';
      for i := 0 to TOBSerLig.Detail.Count - 1 do
      begin
        TOBLS := TOBSerLig.Detail[i];
        if TOBLS.GetValue('GLS_TENUESTOCK') <> 'X' then Continue;
        if TOBLS.GetValue('GLS_RANG') = 1 then Continue;
        NumSerie := TOBLS.GetValue('GLS_IDSERIE');
        St := 'GQS_IDSERIE="' + NumSerie + '" ';
        if WhereSerie = '' then WhereSerie := 'GQS_ARTICLE="' + RefUnique + '" AND (' + St
        else WhereSerie := WhereSerie + ' OR ' + St;
      end;
      if St <> '' then WhereSerie := WhereSerie + ')';
      QQ := OpenSQl('Select * From DISPOSERIE Where ' + WhereSerie, True,-1, '', True);
      if not QQ.EOF then TOBDispoSerie.LoadDetailDB('DISPOSERIE', '', '', QQ, False);
      Ferme(QQ);
      if TOBDispoSerie.Detail.Count > 0 then
      begin
        if not TOBDispoSerie.DeleteDB(False) then V_PGI.IoError := oeUnknown;
      end;
    end else
    begin // Inversion sortie de numéro série
      ResaCli := (Pos('RC', ColPlus) > 0);
      PrepaCli := (Pos('PRE', ColPlus) > 0);
      if (ResaCli) or (PrepaCli) then
      begin
        WhereSerie := '';
        St := '';
        for i := 0 to TOBSerLig.Detail.Count - 1 do
        begin
          TOBLS := TOBSerLig.Detail[i];
          if TOBLS.GetValue('GLS_TENUESTOCK') <> 'X' then Continue;
          if TOBLS.GetValue('GLS_RANG') = 1 then Continue;
          NumSerie := TOBLS.GetValue('GLS_IDSERIE');
          St := 'GQS_IDSERIE="' + NumSerie + '" ';
          if WhereSerie = '' then WhereSerie := '(GQS_ARTICLE="' + TOBLS.GetValue('GLS_ARTICLE') + '" AND ' + St + ')'
          else WhereSerie := WhereSerie + ' OR ' + '(GQS_ARTICLE="' + TOBLS.GetValue('GLS_ARTICLE') + '" AND ' + St + ')';
        end;
        QQ := OpenSQl('Select * From DISPOSERIE Where ' + WhereSerie, True,-1, '', True);
        if not QQ.EOF then TOBDispoSerie.LoadDetailDB('DISPOSERIE', '', '', QQ, False);
        Ferme(QQ);
        for i := 0 to TOBDispoSerie.Detail.Count - 1 do
        begin
          if (ResaCli) then TOBDispoSerie.Detail[i].PutValue('GQS_ENRESERVECLI', '-')
          else TOBDispoSerie.Detail[i].PutValue('GQS_ENPREPACLI', '-');
        end;
        if TOBDispoSerie.Detail.Count > 0 then
          if not TOBDispoSerie.UpdateDB(False) then V_PGI.IoError := oeUnknown;
      end else
      begin
        for i := 0 to TOBSerLig.Detail.Count - 1 do
        begin
          TOBLS := TOBSerLig.Detail[i];
          if TOBLS.GetValue('GLS_TENUESTOCK') <> 'X' then Continue;
          if TOBLS.GetValue('GLS_RANG') = 1 then Continue;
          NumSerie := TOBLS.GetValue('GLS_IDSERIE');
          TOBDS := TOB.Create('DISPOSERIE', TOBDispoSerie, -1);
          if TOBDS = nil then
          begin
            V_PGI.IoError := oeUnknown;
            TOBDispoSerie.free;
            exit;
          end;
          TOBDS.PutValue('GQS_IDSERIE', NumSerie);
          TOBDS.PutValue('GQS_ARTICLE', TOBLS.GetValue('GLS_ARTICLE'));
          TOBDS.PutValue('GQS_DEPOT', Depot);
          TOBDS.PutValue('GQS_NUMEROLOT', TOBLS.GetValue('GLS_NUMEROLOT'));
          TOBDS.PutValue('GQS_ENRESERVECLI', '-');
          TOBDS.PutValue('GQS_ENPREPACLI', '-');
        end;
        if TOBDispoSerie.Detail.Count > 0 then
          if not TOBDispoSerie.InsertOrUpdateDb(False) then V_PGI.IoError := oeUnknown;
      end;
    end;
  end;
  TOBDispoSerie.free;
  {$ENDIF}
end;

procedure LoadLesSeries(TOBPiece, TOBSerie : TOB; CleDoc: R_CleDoc ; TransfoPiece : boolean = False);
var QQ: TQuery;
  i_ind, NumL, NumLPrec, IndiceSerie: integer;
  TOBVrac, TOBLigne, TOBLS, TOBSerLig: TOB;
  stRang : string;
begin
  {$IFNDEF CCS3}
  //if TOBPiece.GetValue('GP_ARTICLESLOT')<>'X' then Exit ;
  TOBVrac := TOB.Create('', nil, -1);
  if TransfoPiece then stRang := ' AND GLS_RANG=0' else stRang := '';
  QQ := OpenSQl('SELECT * FROM LIGNESERIE WHERE ' + WherePiece(CleDoc, ttdSerie, False) + stRang, True,-1, '', True);
if not QQ.EOF then TOBVrac.LoadDetailDB('LIGNESERIE', '', '', QQ, False);
  Ferme(QQ);
  NumLPrec := 0;
  TOBSerLig := nil;
  for i_ind := TOBVrac.Detail.Count - 1 downto 0 do
  begin
    {Lecture des LigneSerie "en vrac"}
    TOBVrac.Detail.Sort('GLS_NUMLIGNE');
    TOBLS := TOBVrac.Detail[i_ind];
    NumL := TOBLS.GetValue('GLS_NUMLIGNE');
    if Numl <> NumLPrec then
    begin
      NumLPrec := NumL;
      {recherche de la ligne de pièce concernée}
      TOBLigne := TOBPiece.FindFirst(['GL_NUMLIGNE'], [NumL], False);
      if TOBLigne <> nil then
      begin
        {Recherche du noeud des lots de la ligne, si pas ok le créer}
        TOBSerLig := TOB.Create('', TOBSerie, -1);
        if TOBSerLig <> nil then
        begin
          IndiceSerie := TOBSerie.Detail.Count;
          TOBLigne.PutValue('GL_INDICESERIE', IndiceSerie);
          TOBLS.ChangeParent(TOBSerLig, -1);
        end;
      end;
    end else
    begin
      if TOBSerLig <> nil then TOBLS.ChangeParent(TOBSerLig, -1);
    end;
  end;
  TOBVrac.Free;
  {$ENDIF}
end;

procedure UpdateLesSeries(TOBL, TOBSerLig: TOB);
var NatureG, Souche: string;
  Numero, IndiceG, i, NumL: integer;
  TOBDet: TOB;
begin
  {$IFNDEF CCS3}
  if ((TOBL = nil) or (TOBSerLig = nil)) then Exit;
  NatureG := TOBL.GetValue('GL_NATUREPIECEG');
  Souche := TOBL.GetValue('GL_SOUCHE');
  Numero := TOBL.GetValue('GL_NUMERO');
  IndiceG := TOBL.GetValue('GL_INDICEG');
  NumL := TOBL.GetValue('GL_NUMLIGNE');
  for i := 0 to TOBSerLig.Detail.Count - 1 do
  begin
    TOBDet := TOBSerLig.Detail[i];
    TOBDet.PutValue('GLS_NATUREPIECEG', NatureG);
    TOBDet.PutValue('GLS_SOUCHE', Souche);
    TOBDet.PutValue('GLS_NUMERO', Numero);
    TOBDet.PutValue('GLS_INDICEG', IndiceG);
    TOBDet.PutValue('GLS_NUMLIGNE', NumL);
  end;
  {$ENDIF}
end;

procedure InitialiseReliquatSerie (TobSerRel, TobSerie_O : TOB);
var i_ind1: integer;
  //TOBSRel : TOB;
begin
  for i_ind1 := 0 to TOBSerie_O.Detail.Count - 1 do
  begin
    //TOBSRel:=TOB.Create('',TOBSerRel,-1);
    TOB.Create('', TOBSerRel, -1);
    //   for i_ind2:=0 to TOBSerie_O.Detail[i_ind1].Detail.Count-1 do TOB.Create('LIGNESERIE',TOBSRel,-1);
  end;
end;

{ NEWPIECE }

procedure _InverseStockSerie(TOBPiece_O, TobPiece, TOBSerie_O: TOB; Transfo: Boolean);
var
  TOBL, TOBSerLig: TOB;
  NaturePiece, VenteAchat: string;
  GereSerie: Boolean;
  i_ind, IndiceSerie: Integer;
begin
  {$IFDEF CCS3}
  Exit;
  {$ELSE}
  NaturePiece := TOBPiece_O.GetValue('GP_NATUREPIECEG');

  GereSerie := (GetInfoParPiece(NaturePiece, 'GPP_NUMEROSERIE') = 'X');
  if not GereSerie then Exit;
  VenteAchat := GetInfoParPiece(NaturePiece, 'GPP_VENTEACHAT');
  if TOBSerie_O.Detail.Count <= 0 then exit;
  TOBSerie_O.Detail[0].AddChampSup('UTILISE', True);
  for i_ind := 0 to TOBPiece_O.Detail.Count - 1 do
  begin
    IndiceSerie := GetChampLigne(TOBPiece_O, 'GL_INDICESERIE', i_ind + 1);
    if IndiceSerie > 0 then
    begin
      TOBL := TOBPiece_O.Detail[i_ind];
      TOBSerLig := TOBSerie_O.Detail[IndiceSerie - 1];
      TOBSerLig.PutValue('UTILISE', 'X');
      MAJDispoSerie(TOBL, TOBSerLig, True);
    end;
  end;
  {$ENDIF}
end;

{ NEWPIECE }
procedure InverseStockSerie(TOBPiece_O, TOBSerie_O: TOB);
begin
  _InverseStockSerie(TOBPiece_O, nil, TOBSerie_O, False);
end;

{ NEWPIECE }
procedure InverseStockSerieTransfo(TOBPiece_O, TobPiece, TOBSerie_O: TOB);
begin
  _InverseStockSerie(TOBPiece_O, TobPiece, TOBSerie_O, True);
end;

{ NEWPIECE }

{***********A.G.L.***********************************************
Auteur  ...... : JS
Créé le ...... : 08/10/2003
Modifié le ... : 08/10/2003
Description .. : Flag le champ GLS_RANG des lignes de document de la 
Suite ........ : table LIGNESERIE à 0 ou à 1.
Suite ........ : 0 : Numéro de série non transformé donc encore "vivant" 
Suite ........ : ou mis en reliquat sur pièce suivante
Suite ........ : 1 : Numéro affecté sur pièce suivante donc non modifiable
Mots clefs ... : NUMERO;SERIE;RELIQUAT
*****************************************************************}
Procedure RazLesSeries(TobSerie, TobSerie_O : TOB);
var iInd,jInd : integer;
    TOBSerLig_O, TOBS_O : TOB;
begin
{$IFNDEF CCS3}
  if not Assigned(TobSerie) then
  begin
    for iInd := 0 to TOBSerie_O.Detail.Count - 1 do
      TobSerie_O.Detail[iInd].PutValueAllFille('GLS_RANG',1);
  end else
  begin
    for iInd := 0 to TOBSerie_O.Detail.Count - 1 do
    begin
      TOBSerLig_O := TOBSerie_O.Detail[iInd];       //Noeud
      for jInd := 0 to TOBSerLig_O.Detail.Count -1 do
      begin
        TOBS_O := TOBSerLig_O.Detail[jInd];
        if TOBSerie.FindFirst(['GLS_ARTICLE','GLS_IDSERIE'],[TOBS_O.GetValue('GLS_ARTICLE'),TOBS_O.GetValue('GLS_IDSERIE')],True) = nil then
          TOBS_O.PutValue('GLS_RANG',0)
        else TOBS_O.PutValue('GLS_RANG',1);
      end;
    end;
  end;
{$ENDIF}
end;

{ NEWPIECE }
Procedure RazAndSaveLesSeries(TobSerie, TobSerie_O : TOB);
begin
{$IFNDEF CCS3}
  RazLesSeries(TobSerie, TobSerie_O);
  TobSerie_O.UpdateDB(False);
{$ENDIF}
end;

{ NEWPIECE }

{***********A.G.L.***********************************************
Auteur  ...... : JS
Créé le ...... : 08/10/2003
Modifié le ... :   /  /    
Description .. : Réaffecte les n° de série en reliquat dans la table 
Suite ........ : LIGNESERIE après avoir inversé le stock
Mots clefs ... : NUMERO;SERIE;RELIQUAT
*****************************************************************}
procedure ReaffecteLesReliquatSerie(TobSerRel, TobPiece_O : TOB);
var iInd, IndiceSerie: integer;
begin
{$IFNDEF CCS3}
if TobSerRel.Detail.Count > 0 then
  begin
    if V_PGI.IoError = oeOk then
    begin
      for iInd := 0 to TobPiece_O.Detail.Count-1 do
      begin
        IndiceSerie := GetChampLigne(TobPiece_O,'GL_INDICESERIE',iInd+1) ;
        if IndiceSerie > 0 then MAJDispoSerie(TobPiece_O.Detail[iInd],TobSerRel.Detail[IndiceSerie-1],False);
      end ;
    end;
  end;
{$ENDIF}
end;


{***********A.G.L.***********************************************
Auteur  ...... : JS
Créé le ...... : 08/10/2003
Modifié le ... :   /  /    
Description .. : Calcul le nombre de numéro de série affecté sur la ligne
Mots clefs ... : NUMERO;SERIE;QUANTITE
*****************************************************************}
function QuelQteEnSerie(TobSerie, TobL : TOB ; NbCompoSerie : double) : double;
var IndiceNomen, IndiceSerie : integer;
    VenteAchat, QualQte : string;
    TOBM : TOB;
    UnitePiece : double;
begin
  Result := 0;
  IndiceSerie := TobL.GetValue('GL_INDICESERIE');
  if IndiceSerie <= 0 then exit;
  if IndiceSerie > TOBSerie.Detail.Count then exit;
  VenteAchat := GetInfoParPiece(TobL.GetValue('GL_NATUREPIECEG'), 'GPP_VENTEACHAT');
  IndiceNomen := TobL.GetValue('GL_INDICENOMEN');
  if IndiceNomen <= 0 then
    Result := TOBSerie.Detail[IndiceSerie - 1].Detail.Count
  else begin
         if NbCompoSerie = 0 then exit;
         Result := TOBSerie.Detail[IndiceSerie - 1].Detail.Count / NbCompoSerie;
       end;
  if VenteAchat = 'ACH' then QualQte := TOBL.GetValue('GLN_QUALIFQTEACH')
  else if (VenteAchat = 'AUT') or (VenteAchat = 'TRF') then QualQte := TOBL.GetValue('GL_QUALIFQTESTO')
  else QualQte := TOBL.GetValue('GL_QUALIFQTEVTE');
  TOBM := VH_GC.MTOBMEA.FindFirst(['GME_QUALIFMESURE', 'GME_MESURE'], ['PIE', QualQte], False);
  UnitePiece := 0;
  if TOBM <> nil then UnitePiece := TOBM.GetValue('GME_QUOTITE');
  if (UnitePiece = 0) or (IndiceNomen > 0) then UnitePiece := 1.0;
  Result := Arrondi(Result / UnitePiece, 6);
end;

end.

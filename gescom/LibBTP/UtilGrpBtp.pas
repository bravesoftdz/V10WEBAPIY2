unit UtilGrpBtp;

interface
Uses HEnt1, HCtrls, UTOB, Ent1, SysUtils, SaisUtil, UtilPGI, AGLInit,
     FactUtil, UtilSais, EntGC, Classes, SaisComm, FactComm, FactNomen,
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     FactAffaire, FactPiece, FactCpta, FactCalc, ParamSoc, UtilArticle,
{$IFDEF BTP}
     BTPUtil,UplannifChUtil,
{$ENDIF}
     FactArticle, DepotUtil,
     FactContreM , StockUtil,uEntCommun ;

procedure UG_reajusteQtePiecePrecedent(TobGenere,TobPiece : TOB);
procedure UG_recupSSTrait (TOBL,TOBSSTRAIT : TOB);
procedure UG_SetSSTrait (TOBL,TOBDESSOUSTRAIT,TOBSSTRAITSDOC : TOB);
procedure CompletelesSSTrait (cledoc : R_CLEDOC;TOBSSTrait : TOB);

implementation
uses factgrpBtp,FactTOB,UtilTOBPiece,UCotraitance;

procedure UG_reajusteQtePiecePrecedent(TobGenere,TobPiece : TOB);
var TOBL,TOBOLDL : TOB;
    indice : integer;
    cledoc : R_CLEDOC;
    precedent,Courant : string;
    ColMoins,ColPlus : string;
    RatioVA : double;
    GereReliquat : boolean;
begin
  for Indice := 0 To TOBGenere.detail.count -1 do
  begin
    TOBL := TOBGenere.detail[Indice];
    if TOBL.GetValue('GL_TYPELIGNE')<>'ART' Then continue;
    // Réajustement des stocks
    Courant := EncodeRefPiece (TOBL,0,true);
    Precedent := TOBL.GetValue('GL_PIECEPRECEDENTE');
    if IsSameDocument (Courant,Precedent) then continue;
    //
    if Precedent <> '' then
    begin
      DecoderefPiece (precedent,cledoc);
      GereReliquat := (GetInfoParPiece(Cledoc.Naturepiece, 'GPP_RELIQUAT')='X');
      TOBOLDL := TOBpiece.findFirst(['GL_NATUREPIECEG','GL_NUMERO','GL_NUMORDRE'],
                                    [cledoc.NaturePiece,cledoc.numeroPiece,Cledoc.NumOrdre],true);
      if TOBOLDL = nil then continue;
      if TOBL.getValue('GL_QTEFACT') <> 0 then
      begin
      	if GereReliquat then // le document ne gere pas des reliquats
        begin
          TOBOLDL.PutValue('GL_QTERESTE',TOBOLDL.GetValue('GL_QTERESTE')-TOBL.GetValue('GL_QTEFACT'));
          if (TOBOLDL.getValue('GL_QTERESTE') < 0) and ( TOBOLDL.GetValue('GL_QTEFACT')> 0) then TOBOLDL.PutValue('GL_QTERESTE',0);
          if CtrlOkReliquat(TOBOLDL, 'GL') then TOBOLDL.PutValue('GL_MTRESTE',TOBOLDL.GetValue('GL_MTRESTE')-TOBL.GetValue('GL_MONTANTHT'));
        end;
      end;
    end;
  end;
end;

procedure UG_recupSSTrait (TOBL,TOBSSTRAIT : TOB);
var Sql : string;
		Where : string;
    Q : TQuery;
    TOBLOC : TOB;
    cledoc : R_CLEDOC;
begin
  TOBLOC := TOBSSTrait.FindFirst (['BPI_TIERSFOU'],[TOBL.GetString('GL_FOURNISSEUR')],True);
  if TOBLOC = nil then
  begin
    cledoc := TOB2CleDoc(TOBL);
    Sql := 'SELECT * FROM PIECEINTERV ';
    Where := 'WHERE ' + WherePiece(CleDoc, ttdPieceInterv, False)+' AND BPI_TIERSFOU="'+TOBL.GetString('GL_FOURNISSEUR')+'"';
    Q := OpenSQL(Sql + where, True,-1, '', True);
    if not Q.Eof then
    begin
    	TOBLOc := TOB.Create ('PIECEINTERV',TOBSSTRAIT,-1);
    	TOBLOc.SelectDB('',Q);
      TOBLOC.SetInteger('BPI_ORDRE',TOBSSTRAIT.detail.count -1);
    end;
    ferme (Q);
  end;
end;

procedure CompletelesSSTrait (cledoc : R_CLEDOC;TOBSSTrait : TOB);
var TOBSousTrait : TOB;
		indice : Integer;
begin
	TOBSousTrait := TOB.Create ('LES SOUS TRAITS',nil,-1);
  LoadLesSousTraitants(cledoc,TOBSousTrait);
  Indice := 0;
  repeat
    if TOBSSTrait.FindFirst(['BPI_TIERSFOU'],[TOBSousTrait.Detail[indice].GetString('BPI_TIERSFOU')],true) = nil then
    begin
			TOBSousTrait.Detail[indice].ChangeParent(TOBSSTrait,-1);
      TOBSousTrait.Detail[TOBSousTrait.Detail.Count-1].SetInteger('BPI_ORDRE',TOBSousTrait.detail.count -1);
      TOBSousTrait.Detail[TOBSousTrait.Detail.Count-1].SetString('UTILISE','X');
    end else inc(indice);
  until Indice >= TOBSousTrait.detail.Count;
  TOBSousTrait.free;
end;

procedure UG_SetSSTrait (TOBL,TOBDESSOUSTRAIT,TOBSSTRAITSDOC : TOB);
var TOBLOC : TOB;
begin
  TOBLOC := TOBSSTRAITSDOC.FindFirst (['BPI_TIERSFOU'],[TOBL.GetString('GL_FOURNISSEUR')],True);
  if TOBLOC = nil then
  begin
  	TOBLOC := TOBDESSOUSTRAIT.FindFirst (['BPI_TIERSFOU'],[TOBL.GetString('GL_FOURNISSEUR')],True);
    if TOBLOC <> nil then
    begin
      TOBLOC.ChangeParent(TOBSSTRAITSDOC,-1);
      TOBLOC.SetInteger('BPI_ORDRE',TOBSSTRAITSDOC.detail.count -1);
    end;
  end;
end;
end.


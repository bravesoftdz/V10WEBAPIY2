unit UtilEpurePiece;

interface
Uses Classes,
     Windows,
     sysutils,
     ComCtrls,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$ELSE}
     MainEagl,
{$ENDIF}
     HCtrls,
     Hpanel,
     HEnt1,
     HMsgBox,
     Vierge,
     EntGc,
     BTPUtil,
     ExtCtrls,
     SaisUtil,
     UtilPgi,
     uEntCommun,
     UtilTOBPiece,
     UTOB;

procedure	EpureLaPiece (fTOBPiece : TOB);

implementation
uses UtilBTPgestChantier,FactCpta;

procedure	EpureLaPiece (fTOBPiece : TOB);
var cledoc : r_cledoc;
    SQl : string;
    OldEcr,OldStk : RMVT;
    TP : TOB;
    QQ : TQuery;
begin
  TP := TOB.Create('PIECE',nil,-1);
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  CleDoc.NaturePiece := fTOBPiece.getValue('GP_NATUREPIECEG');
  CleDoc.DatePiece := fTOBPiece.getValue('GP_DATEPIECE');
  CleDoc.Souche := fTOBPiece.getValue('GP_SOUCHE');
  CleDoc.NumeroPiece := fTOBPiece.getValue('GP_NUMERO');
  CleDoc.Indice := fTOBPiece.getValue('GP_INDICEG');
  QQ := OpenSQL('SELECT * FROM PIECE WHERE '+WherePiece(cledoc,ttdPiece,false),True,1,'',True);
  if not QQ.Eof then
  begin
    TP.SelectDB('',QQ);
    DetruitCompta(TP,NowH, OldEcr, OldStk);
  end;
  ferme (QQ);
  TP.Free;

  if (fTOBPiece.GetValue('GP_AFFAIREDEVIS') <> '') and (copy(fTOBPiece.GetValue('GP_AFFAIREDEVIS'), 1, 1) = 'Z') and
      ((fTOBPiece.GetValue('GP_NATUREPIECEG')='DBT') or (fTOBPiece.GetValue('GP_NATUREPIECEG')='BCE')) then
  begin
    // suppression de la sous-affaire associée
    ExecuteSQL('DELETE FROM AFFAIRE WHERE AFF_AFFAIRE="' + fTOBPiece.GetValue('GP_AFFAIREDEVIS') + '"');
    ExecuteSQL('DELETE FROM BSITUATIONS WHERE BST_SSAFFAIRE="' + fTOBPiece.GetValue('GP_AFFAIREDEVIS') + '"');
    SQL := 'DELETE FROM BTMEMOFACTURE WHERE '+
           'BMF_NATUREPIECEG="' + cledoc.NaturePiece + '" AND '+
           'BMF_SOUCHE="' + cledoc.Souche +'" AND '+
           'BMF_NUMERO='+InttOStr(cledoc.NumeroPiece)+' AND '+
           'BMF_INDICEG='+InttOStr(cledoc.Indice);
    ExecuteSQL(SQL);
  end;
  ExecuteSQL('DELETE FROM BTPARDOC WHERE '+WherePiece(CleDoc, ttdParDoc, False));
	ExecuteSQL('DELETE FROM LIGNE WHERE ' + WherePiece(CleDoc, ttdLigne, False));
  ExecuteSQL('DELETE FROM LIGNECOMPL WHERE ' + WherePiece(CleDoc, ttdLigneCompl, False));
  ExecuteSQL('DELETE FROM PIEDBASE WHERE ' + WherePiece(CleDoc, ttdPiedBase, False));
  ExecuteSQL('DELETE FROM LIGNEBASE WHERE ' + WherePiece(CleDoc, ttdLigneBase, False));
  ExecuteSQL('DELETE FROM PIEDECHE WHERE ' + WherePiece(CleDoc, ttdEche, False));
  ExecuteSQL('DELETE FROM ACOMPTES WHERE GAC_NATUREPIECEG="' + fTOBPiece.GetValue('GP_NATUREPIECEG')+ '" '+
  					 'AND GAC_SOUCHE="' + fTOBPiece.GetValue('GP_SOUCHE')+'" '+
    				 'AND GAC_NUMERO=' + IntToStr(fTOBPiece.GetValue('GP_NUMERO'))+
    				 'AND GAC_INDICEG=' + IntToStr(fTOBPiece.GetValue('GP_INDICEG')));
  ExecuteSQL('DELETE FROM PIEDPORT WHERE GPT_NATUREPIECEG="' + fTOBPiece.GetValue('GP_NATUREPIECEG')+ '" '+
  					 'AND GPT_SOUCHE="' + fTOBPiece.GetValue('GP_SOUCHE')+'" '+
    				 'AND GPT_NUMERO=' + IntToStr(fTOBPiece.GetValue('GP_NUMERO'))+
    				 'AND GPT_INDICEG=' + IntToStr(fTOBPiece.GetValue('GP_INDICEG')));
	ExecuteSQL('DELETE FROM LIGNEOUV WHERE ' + WherePiece(CleDoc, ttdOuvrage, False));
  ExecuteSQL('DELETE FROM LIGNEOUVPLAT WHERE ' + WherePiece(CleDoc, ttdOuvrageP , False));
  ExecuteSQL('DELETE FROM LIGNEFAC WHERE ' + WherePiece(CleDoc, ttdLigneFac , False));
  ExecuteSQL('DELETE FROM LIGNEPHASES WHERE ' + WherePiece(CleDoc, TtdLignePhase , False));
  ExecuteSQL('DELETE FROM PIECERG WHERE ' + WherePiece(CleDoc, ttdRetenuG, False));
	ExecuteSQL('DELETE FROM PIEDBASERG WHERE ' + WherePiece(CleDoc, ttdBaseRG, False));
  ExecuteSQL('DELETE FROM BTPIECEMILIEME WHERE ' + WherePiece(CleDoc, TTdRepartmill , False));
  ExecuteSql('DELETE FROM LIENSOLE WHERE ' + WherePiece(CleDoc, ttdLienOle, False));
  ExecuteSQL('DELETE FROM PIECEADRESSE WHERE ' + WherePiece(CleDoc, ttdPieceAdr, False));
  ExecuteSQL('DELETE FROM VENTANA WHERE YVA_TABLEANA="GL" AND YVA_IDENTIFIANT LIKE "' + EncodeRefPresqueCPGescom(fTOBPiece) + '"');
  ExecuteSQL('DELETE FROM VENTANA WHERE YVA_TABLEANA="GS" AND YVA_IDENTIFIANT LIKE "' + EncodeRefPresqueCPGescom(fTOBPiece) + '"');
  ExecuteSql('DELETE FROM BVARDOC WHERE ' + WherePiece(CleDoc, ttdVardoc , False));
  ExecuteSql('DELETE FROM LIENDEVCHA WHERE BDA_REFD="'+EncodeLienDevCHA(fTOBPiece)+'"');
  ExecuteSQL('DELETE FROM PIECETRAIT WHERE ' + WherePiece(CleDoc, ttdPieceTrait  , False));
  ExecuteSQL('DELETE FROM PIECEINTERV WHERE ' + WherePiece(CleDoc, ttdPieceInterv  , False));
  ExecuteSQL('DELETE FROM BLIGNEMETRE WHERE ' + WherePiece(CleDoc, ttdLigneMetre  , False));
  ExecuteSQL('DELETE FROM DETAILDEMPRIX WHERE ' + WherePiece(CleDoc, TtdDetailDemPrix , False));
  ExecuteSQL('DELETE FROM FOURLIGDEMPRIX WHERE ' + WherePiece(CleDoc, TtdFournDemprix , False));
  ExecuteSQL('DELETE FROM PIECEDEMPRIX WHERE ' + WherePiece(CleDoc, ttdPieceDemPrix, False));
  ExecuteSQL('DELETE FROM TIMBRESPIECE WHERE ' + WherePiece(CleDoc, ttdTimbres, False));
	ExecuteSQL('DELETE FROM PIECE WHERE ' + WherePiece(CleDoc, ttdPiece, False));

  if Pos(cledoc.NaturePiece , 'FBT;FBP;DAC') > 0 then
  begin
    SQL := 'DELETE FROM BSITUATIONS WHERE '+
           'BST_NATUREPIECE="' + cledoc.NaturePiece + '" AND '+
           'BST_SOUCHE="' + cledoc.Souche +'" AND '+
           'BST_NUMEROFAC='+InttOStr(cledoc.NumeroPiece);
    ExecuteSQL(SQL);
  end;
end;

end.

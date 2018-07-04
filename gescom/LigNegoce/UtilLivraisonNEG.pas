unit UtilLivraisonNEG;

interface
Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     AglInit,
     HMsgBox,
     uEntCommun,
     FactComm,
     FactTOB,
     UTOF ;


procedure GetLivrableNEG(TOBL : TOB; var QteLivrable : double; var PMAP : double );

implementation

uses UtilTOBPiece;

function EncodeFindRefOrigine (Cledoc : r_cledoc) : string;
begin
	Result := '%;'+Cledoc.NaturePiece+';'+Cledoc.Souche+';'+Inttostr(Cledoc.NumeroPiece)+';'+
  					Inttostr(Cledoc.Indice)+';'+Inttostr(Cledoc.NumOrdre)+';';
end;

procedure GetLivrableNEG(TOBL : TOB; var QteLivrable : double; var PMAP : double );
{$IFDEF PLUSTARD}
var QQ : TQuery;
		SQL : string;
    Cledoc,CledocPC : R_CLEDOC;
    TOBR,TOBLL : TOB;
{$ENDIF}
begin
  if TOBL.GetString('GL_PIECEPRECEDENTE')='' then Exit;
  //
  {$IFDEF PLUSTARD}
  TOB2Cledoc(TOBL,CledocPC);
  TOBR := TOB.Create ('LES LIGNES RECEP',nil,-1);
  TOBLL := TOB.Create ('LES LIGNES DEJA LIVRES',nil,-1);
  //
  TRY
    DecodeRefPiece(TOBL.GetString('GL_PIECEPRECEDENTE'),Cledoc);
    // Détail des réceptionnés
    SQL := 'SELECT GL_QTEFACT,GL_QUALIFQTEACH,GL_PUHTDEV AS DPA, '+
    			 '(SELECT GA_QUALIFUNITEVTE FROM ARTICLE WHERE GA_ARTICLE=GL_ARTICLE) AS GL_QUALIFQTEVTE,'+
           '0 AS QTEVTE FROM LIGNE WHERE GL_PIECEORIGINE LIKE "'+EncodeFindRefOrigine(Cledoc)+'" AND '+
           'GL_NATUREPIECEG IN ("BLF","FF") AND (GL_QTERESTE > 0 OR GL_MTRESTE > 0) AND GL_VIVANTE="X"';
    QQ := OpenSQl(Sql,True,-1,'',True);
    if not QQ.eof then
    begin
      TOBR.LoadDetailDB('LIGNE','','',QQ,false);
    end;
    Ferme(QQ);
    // Détails des déjà livrés et/ou facturés (sans compter la pièce courante)
    SQL := 'SELECT SUM(GL_QTEFACT) FROM LIGNE WHERE GL_PIECEORIGINE LIKE "'+EncodeFindRefOrigine(Cledoc)+'"';
           'GL_NATUREPIECEG IN ("BLC","FBC") AND (GL_QTERESTE > 0 OR GL_MTRESTE > 0) AND GL_VIVANTE="X" AND NOT ('+WherePiece(CledocPC,ttdLigne,False)+')';
    QQ := OpenSQL(SQL,True,-1,'',true);
    if not QQ.eof then
    begin
      TOBLL.LoadDetailDB('LIGNE','','',QQ,false);
    end;
    ferme (QQ);
  FINALLY
    TOBR.Free;
    TOBLL.Free;
  end;
  {$ENDIF}
end;

end.

unit ULienVP;

interface
uses  Classes,
    {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
    HDB,
    HEnt1,
    HMsgBox,
    AffaireUtil,
    UtilPGI,
    UTOB,
    AglInit,
    HCtrls,
    paramsoc,
    UtilTOBPiece,
    UentCommun,
    UdateUtils;


procedure PreparePlannif (cledoc : r_cledoc);
procedure LanceVp;

implementation

procedure PreparePlannif (cledoc : r_cledoc);
var QQ : TQuery;
  	TOBP,TOBL,TOBA : TOB;
    Affaire,CodeAffaire : string;
    Indice : Integer;
begin
  TOBA := TOB.Create ('AFFAIRE',nil,-1);
  TOBP := TOB.Create ('PIECE',nil,-1);
  TOBL := TOB.Create ('LIGNE',nil,-1);
  TRY
    QQ := OpenSql ('SELECT GP_AFFAIRE FROM PIECE WHERE '+WherePiece(cledoc,ttdPiece,False),True,1,'',true);
    if not QQ.eof then
    begin
    	Affaire := QQ.fields[0].AsString;
      TOBP.SelectDB('',QQ);
    end;
  	Ferme(QQ);
    if Affaire = '' then exit;
    QQ := OpenSql ('SELECT AFF_AFFAIRE,AFF_LIBELLE,AFF_RESPONSABLE,'+
                   'AFF_DATEDEBUT FROM AFFAIRE WHERE AFF_AFFAIRE="'+Affaire+'"',True,1,'',true);
    if not QQ.eof then
    begin
      TOBA.SelectDB('',QQ);
    end;
    ferme (QQ);
    QQ := OpenSQL('SELECT GP_TYPELIGNE,GL_TYPEARTICLE,GL_ARTICLE,GL_LIBELLE,'+
    							'GL_QTEFACT,GL_QUALIFQTEVTE,GL_TPSUNITAIRE FROM LIGNE WHERE '+
                  WherePiece(cledoc,ttdLigne,false)+ ' ORDER GL_NUMLIGNE',True,-1,'',true);
    if not QQ.eof then
    begin
      TOBP.LoadDetailDB('LIGNE','','',QQ,false);
    end;
    Ferme(QQ);
    //
    For Indice := 0 to TOBP.detail.count -1 do
    begin

    end;
	finally
    TOBA.free;
    TOBP.free;
    TOBL.free;
  End;
end;

procedure LanceVp;
begin

end;

end.

unit UtilFermeturePiece;
 
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
     EntGc,
     ExtCtrls,
     UtilPgi,
     stockUtil,
     UTOB;

procedure EnleveReserveAffaire (Affaire : string);
procedure ReajusteStock (fTOBPiece : TOB);

implementation


procedure ReajusteStock (fTOBPiece : TOB);
var Req : String;
		TOBLTR,TOBLL : TOB;
    TOBDEpot,TOBD : TOB;
    Indice : integer;
    RCCLient,RCFourn : boolean;
    Ratio,Qte : double;
    QQ : TQuery;
    NaturePiece : string;
begin
	TOBDepot := TOB.Create ('LES DISPO',nil,-1);
  //
  TRY
    NaturePiece := fTOBPiece.getValue('GP_NATUREPIECEG');
    RCClient := (Pos('RC',GetInfoParPiece (NaturePiece,'GPP_QTEPLUS'))>0);
    RCFOurn := (Pos('RF',GetInfoParPiece (NaturePiece,'GPP_QTEPLUS'))>0);
    if (RCClient) or (RCFourn) then
    begin
      TOBLTR := TOB.Create ('LIGNE',nil,-1);
      Req := 'SELECT GL_ARTICLE,GL_DEPOT,GL_QTERESTE, GL_MTRESTE, GL_QUALIFQTEVTE,GL_QUALIFQTESTO,GL_QUALIFQTEACH FROM LIGNE WHERE '+
             'GL_NATUREPIECEG="'+ fTOBPiece.getValue('GP_NATUREPIECEG')+'" AND '+
             'GL_SOUCHE="'+fTOBPiece.getValue('GP_SOUCHE')+'" AND '+
             'GL_NUMERO='+IntToStr(fTOBPiece.getValue('GP_NUMERO'))+ ' AND '+
             'GL_TYPELIGNE<>"CEN" AND '+
             'GL_INDICEG='+IntToStr(fTOBPiece.getValue('GP_INDICEG'))+' AND '+
             'GL_ARTICLE<>"" AND (GL_QTERESTE <> 0)';
      TOBLTR.LoadDetailDBFromSQL('LIGNE',Req,false);
      if TOBLTR.detail.count > 0 then
      begin
        for Indice := 0 to TOBLTR.detail.count -1 do
        begin
          TOBLL := TOBLTR.detail[Indice];
          Req := 'SELECT * FROM DISPO WHERE GQ_ARTICLE="'+TOBLL.GetValue('GL_ARTICLE')+'" AND '+
                 'GQ_DEPOT="'+TOBLL.GetVAlue('GL_DEPOT')+'"';
          QQ := OpenSQL(REq,True,-1,'',true);
          if not QQ.eof then
          begin
            TOBD := TOB.Create ('DISPO',TOBdepot,-1);
            TOBD.SelectDB('',QQ);
            Ratio:=GetRatio (TOBLL,nil, trsStock);
            Qte := TOBLL.GetValue('GL_QTERESTE');
            Qte:=Arrondi(Qte/Ratio,V_PGI.OkDecQ) ;
            if RCFOurn then
            begin
              TOBD.PutValue('GQ_RESERVEFOU',TOBD.GetValue('GQ_RESERVEFOU')-Qte);
            end else
            begin
              TOBD.PutValue('GQ_RESERVECLI',TOBD.GetValue('GQ_RESERVECLI')-Qte);
            end;
          end;
          ferme (QQ);
        end;
      	if not TOBDepot.UpdateDB(false) then V_PGI.Ioerror := OeUnknown;
      end;
    end;
  FINALLY
  	TOBDepot.free;
  END;
end;

procedure EnleveReserveAffaire (Affaire : string);
var TOBPieces : TOB;
		REqNaturePiece,Req : string;
    Indice : integer;
begin
	TOBPieces := TOB.Create ('LES PIECES', nil,-1);
  TRY
  	REqNaturePiece := 'SELECT GP_NATUREPIECEG FROM PARPIECE WHERE GPP_MASQUERNATURE="-" AND ((GPP_QTEPLUS LIKE "%RC%") '+
                      'OR (GPP_QTEPLUS LIKE "%RF%"))';
    Req := 'SELECT * FROM PIECE WHERE GP_AFFAIRE="'+Affaire+'" AND '+
    							 'GP_NATUREPIECEG IN ('+ReqNaturePiece+') AND GP_VIVANTE="X"';
    TOBPIEces.LoadDetailDBFromSQL  ('PIECE',req,false,false);
    if TOBPieces.detail.count > 0 then
    begin
    	for Indice := 0 to TOBPieces.detail.count -1 do
      begin
        ReajusteStock (TOBPieces.detail[Indice]);
        if V_PGI.ioerror <> OeOk then break;
      end;
    end;
  FINALLY
  	TOBPieces.free;
  END;
end;

end.

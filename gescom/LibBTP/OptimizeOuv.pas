unit OptimizeOuv;

interface

uses Classes,HEnt1,Ent1,EntGC,UTob, AGLInit, SaisUtil, UtilPGI,
{$IFDEF EAGLCLIENT}
     UtileAGL,MaineAGL,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}EdtREtat,FE_Main,uPDFBatch,
{$ENDIF}
     HCtrls,ParamSoc,SysUtils,FactTOB,StockUtil,NomenUtil,UtilArticle,uEntCommun,UtilTOBPiece;

type
	TOptimizeOuv = class
  	private
    	fTobArticles : TOB;
      fTOBDispo : TOB;
      fTOBFournisseur : TOB;
      fARecalc : Boolean;
    public
      property Arecalc: boolean read fARecalc write fARecalc;
      property TOBArticles : TOB read fTobArticles write fTobArticles;
      property TOBDispo : TOB read fTOBDispo write fTOBDispo;
      property TOBFournisseur : TOB read fTOBFournisseur write fTOBFournisseur;
      procedure ChargeLesTOBS (Cledoc : R_CLEDOC);
    	Constructor create;
      destructor destroy; override;
      procedure fusionneArticles (TOBArticlesIni : TOB);
      procedure initialisation;
  end;


implementation

{ TOptimizeOuv }

procedure TOptimizeOuv.ChargeLesTOBS(Cledoc: R_CLEDOC);
var Q : Tquery;
		Select : string;
    Indice : integer;
    TOBA,TOBD : TOB;
begin
	// On charge les articles utilisés dans les ouvrages
  Q := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE IN (SELECT DISTINCT BLO_ARTICLE FROM LIGNEOUV WHERE '+
                WherePiece(CleDoc,ttdOuvrage,False)+') ORDER BY GA_ARTICLE',true,-1, '', True);
  if not Q.eof then fTobArticles.LoadDetailDB ('ARTICLE','','',Q,True,True);
  ferme(Q);
	// On charge les dispos articles par dépot utilisés dans les ouvrages
  Select := 'SELECT * FROM DISPO WHERE GQ_ARTICLE '+
  					'IN (SELECT DISTINCT BLO_ARTICLE FROM LIGNEOUV WHERE '+
            WherePiece(CleDoc,ttdOuvrage,False)+') AND '+
            'GQ_DEPOT IN (SELECT DISTINCT BLO_DEPOT FROM LIGNEOUV WHERE '+
            WherePiece(CleDoc,ttdOuvrage,False)+
            ') AND GQ_CLOTURE="-" ORDER BY GQ_ARTICLE';
  Q:=OpenSQL(Select,True,-1, '', True) ;
  if not Q.eof then fTOBDispo.LoadDetailDB ('DISPO','','',Q,True,True);
  ferme(Q);

	// On charge les fournisseurs des articles utilisés dans les ouvrages
  Select := 'SELECT * FROM TIERS WHERE T_TIERS '+
  					'IN (SELECT DISTINCT BLO_FOURNISSEUR FROM LIGNEOUV WHERE '+
            WherePiece(CleDoc,ttdOuvrage,False)+' AND BLO_FOURNISSEUR<>"")';
  Q:=OpenSQL(Select,True,-1, '', True) ;
  if not Q.eof then fTOBFournisseur.LoadDetailDB ('TIERS','','',Q,True,True);
  ferme(Q);

  for Indice := 0 to fTobArticles.detail.count -1 do
  begin
  	TOBA := fTOBArticles.detail[Indice];
    InitChampsSupArticle (TOBA);
    TOBA.AddChampSupValeur('REFARTSAISIE', '', False);
    TOBA.AddChampSupvaleur('REFARTBARRE', '', False);
    TOBA.AddChampSupValeur('REFARTTIERS', '', False);
    TOBA.AddChampSupValeur('SUPPRIME', '-', False);
    TOBA.AddChampSupValeur('UTILISE', '-', False);
    TOBD := fTOBDispo.findFirst(['GQ_ARTICLE'],[TOBA.GetValue('GA_ARTICLE')],True);
    while TOBD <> nil do
    begin
      TobD.Changeparent(TOBA, -1);
      TobD := fTOBDispo.FindNext(['GQ_ARTICLE'], [TOBA.GetValue('GA_ARTICLE')], False);
    end;
		DispoChampSupp (TOBA);
  end;
  fTobdispo.clearDetail; // le reste .. on le degage..liberation memoire
end;

constructor TOptimizeOuv.create;
begin
	fTobArticles := TOB.Create ('LES ARTICLES OUV',nil,-1);
	fTOBDispo := TOB.Create ('LES STOCKS',nil,-1);
  fTOBFournisseur := TOB.Create ('LES FOURNISSEURS',nil,-1);
end;

destructor TOptimizeOuv.destroy;
begin
  inherited;
  fTobArticles.free;
	fTOBDispo.free;
  fTOBFournisseur.free;
end;

procedure TOptimizeOuv.fusionneArticles(TOBArticlesIni: TOB);
var TOBA : TOB;
		indice : integer;
begin
	if fTOBArticles.detail.count = 0 then exit;
	indice := 0;
	repeat
    TOBA := fTOBArticles.detail[Indice];
    if TOBArticlesIni.FindFirst (['GA_ARTICLE'],[TOBA.GetValue('GA_ARTICLE')],true) = nil then
    begin
    	TOBA.ChangeParent (TOBArticlesIni,-1);
    end else inc(Indice);
  until indice > fTobArticles.detail.count -1;
  fTOBArticles.clearDetail; // plus la peine de garder le reste..on libere la memeoire.
  fTOBFournisseur.clearDetail;
end;

procedure TOptimizeOuv.initialisation;
begin
	fTobArticles.ClearDetail;
	fTOBDispo.ClearDetail;
  fTOBFournisseur.clearDetail;
end;

end.

unit UtilOutilsLivraisons;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, StdCtrls, HSysMenu, HTB97,HCtrls,EntGc
{$IFDEF EAGLCLIENT}
      ,MaineAGL,
{$ELSE}
      ,Fe_Main,
  {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  UTOB,FactComm,FactTOb,UtilConso,HEnt1,uEntCommun,UtilTOBPiece;

type
  TFOutilsLivr = class(TFVierge)
    MRetour: TMemo;
    Blance: TToolbarButton97;
    procedure BValiderClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure TraiteDetailLivraisons;
    procedure TraiteLigne (TOBL : TOB);
  	procedure LoadlesTobs (TOBL : TOB);
    procedure LibereToutCa;

  private
    { Déclarations privées }
    TOBLiv : TOB;
    TOBCOnso : TOB;
    TOBCde,TOBRF : TOB;
    TOBA : TOB;
  public
    { Déclarations publiques }
  end;

procedure RecalculLivrNonValorise;

implementation

{$R *.DFM}

procedure RecalculLivrNonValorise;
var XX : TFOutilsLivr;
begin
	XX := TFOutilsLivr.Create (Application);
  XX.ShowModal;
  XX.free;
end;

procedure TFOutilsLivr.BValiderClick(Sender: TObject);
var QQ : TQuery;
begin
  inherited;
  Mretour.Clear;
  TOBLiv.ClearDetail;
  QQ := OpenSql ('SELECT LIGNE.*,BCO_NUMMOUV,BCO_INDICE FROM LIGNE '+
  							 'LEFT JOIN CONSOMMATIONS ON BCO_NATUREPIECEG=GL_NATUREPIECEG AND BCO_SOUCHE=GL_SOUCHE AND BCO_NUMERO=GL_NUMERO AND BCO_INDICEG=GL_INDICEG AND BCO_NUMORDRE=GL_NUMORDRE '+
                 'WHERE GL_NATUREPIECEG="LBT" AND GL_PIECEPRECEDENTE<>"" AND GL_CODEARTICLE<>"" AND GL_DPA=0 AND GL_AFFAIRE <>""',true);
  if not QQ.eof then
  begin
  	TOBLIv.LoadDetailDB ('LIGNE','','',QQ,false,true);
    ferme (QQ);
    Mretour.lines.add ('Début de traitement à : '+TimeToStr (Now));
    TraiteDetailLivraisons;
    Mretour.lines.add ('Fin de traitement à : '+TimeToStr (Now));
  end else
  begin
  	ferme (QQ);
  	Mretour.lines.Add ('Rien à traiter');
  end;

end;

procedure TFOutilsLivr.FormCreate(Sender: TObject);
begin
  inherited;
	TOBLIv := TOB.create ('LES LIVRAISONS',nil,-1);
  TOBCOnso := TOB.create ('CONSOMMATIONS',nil,-1);
  TOBCde := TOB.Create ('LES LIGNES DE BESOINS',nil,-1);
  TOBRF := TOB.Create ('LES LIGNES DE RECEPTIONS',nil,-1);
  TOBA := TOB.Create ('ARTICLE',nil,-1);
end;

procedure TFOutilsLivr.FormDestroy(Sender: TObject);
begin
  inherited;
	TOBLiv.free;
  TOBCOnso.free;
  TOBCde.free;
  TOBRF.free;
  TOBA.free;
end;

procedure TFOutilsLivr.LibereToutCa;
begin
	TOBConso.Initvaleurs;
	TOBCde.ClearDetail;
	TOBRF.ClearDetail;
  TOBA.InitValeurs;
end;

procedure TFOutilsLivr.LoadlesTobs(TOBL: TOB);
var QQ : TQuery;
		cledoc : R_CLEDOC;
begin
	DecodeRefPiece (TOBL.GetValue('GL_PIECEPRECEDENTE'),Cledoc);
	TOBConso.Initvaleurs;
	TOBCde.ClearDetail;
	TOBRF.ClearDetail;
  TOBA.InitValeurs;

  QQ := OpenSql ('SELECT GA_ARTICLE,GA_PAHT,GA_DPR FROM ARTICLE WHERE GA_ARTICLE="'+TOBL.GetValue('GL_ARTICLE')+'"',true);
  TOBA.SelectDB ('',QQ);
  ferme (QQ);

	if not VarIsNull (TOBL.GetValue('BCO_NUMMOUV')) then
  begin
  	QQ := OpenSql ('SELECT * FROM CONSOMMATIONS WHERE BCO_NUMMOUV='+InttoStr(TOBL.GetValue('BCO_NUMMOUV'))+
    							 ' AND BCO_INDICE='+InttoStr(TOBL.GetVAlue('BCO_INDICE')),true);
    if not QQ.eof then TOBConso.SelectDB ('',QQ);
    ferme (QQ);
  end else
  begin
  end;
	// on recup comme ca la derniere piece et donc les derniers prix
  QQ := OpenSql ('SELECT * FROM LIGNE WHERE GL_PIECEORIGINE="'+TOBL.GetValue('GL_PIECEPRECEDENTE')+
  							 '" AND GL_NATUREPIECEG IN ("BLF","FF") AND (GL_QTERESTE <> 0  OR GL_MTRESTE <> 0)',True);
  if QQ.eof then
  begin
  	ferme (QQ);
    QQ := OpenSql ('SELECT * FROM LIGNE WHERE GL_ARTICLE="'+TOBL.GetValue('GL_ARTICLE')+
                   '" AND GL_AFFAIRE="'+TOBL.GetValue('GL_AFFAIRE')+'" AND GL_NATUREPIECEG IN ("BLF","FF") AND (GL_QTERESTE <> 0 OR GL_MTRESTE <> 0)',True);
  end;
  if not QQ.eof then TOBRF.LoadDetailDB ('LIGNE','','',QQ,true);
  Ferme (QQ);
  QQ := OpenSql ('SELECT * FROM LIGNE WHERE '+WherePiece(cledoc,ttdLigne,true,true),True);
  if not QQ.eof then
  begin
  	// cas super top
  	TOBCde.LoadDetailDB ('LIGNE','','',QQ,true);
  end else
  begin
    Ferme (QQ);
    // on essaye de trouver l'article dans le document mais pas forcement a la ligne attendu..
    QQ := OpenSql ('SELECT GL_QUALIFQTEVTE,GL_QTERESTE, GL_MTRESTE, GL_PRIXPOURQTE,GL_ARTICLE,GL_CODEARTICLE,GL_AVENANT,GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_QTEFACT,GL_DPA,GL_DPR,GL_PUHTDEV FROM LIGNE WHERE GL_NATUREPIECEG="'+Cledoc.NaturePiece+'" AND '+
    							 ' GL_SOUCHE="'+Cledoc.Souche+'" AND GL_NUMERO='+InttOStr(Cledoc.NumeroPiece)+' AND GL_ARTICLE="'+TOBL.GeTValue('GL_ARTICLE')+'"',True);
    if not QQ.eof then TOBCde.LoadDetailDB ('LIGNE','','',QQ,true);
  end;
  Ferme (QQ);
end;

procedure TFOutilsLivr.TraiteDetailLivraisons;
var Indice : integer;
begin
  for Indice := 0 to TOBLiv.detail.count -1 do
  begin
  	TraiteLigne(TOBLiv.detail[Indice]);
  end;
end;

procedure TFOutilsLivr.TraiteLigne(TOBL: TOB);
var Indice : integer;
		PAA,PAV,PR,PV,MontantPa,MontantPR,Qte,Coeffg : Double;
    TOBR : TOB;
begin
  PAA := 0;
  PAV := 0;
  PR := 0;
  PV := 0;
  Coeffg := 0;
  // --
  LoadlesTobs (TOBL);
  if TOBRf.detail.count > 0 then
  begin
  	// ici ce sont les achats
  	Montantpa := 0;
    Qte := 0;
    for Indice := 0 to TOBRF.detail.count -1 do
    begin
    	TOBR := TOBRF.detail[Indice];
      MontantPA := MontantPA + (TOBR.GetValue('GL_QTEFACT')*TOBR.GetValue('GL_PUHTDEV'));
      Qte := QTE + TOBR.GetValue('GL_QTEFACT');
    end;
    if Qte <> 0 then PAA := Arrondi(MontantPA / Qte,V_PGI.OkdecP);
  end;
  if TOBCde.detail.count > 0 then
  begin
  	// la les besoins.....de chantiers
  	Montantpa := 0;
  	MontantpR := 0;
    Qte := 0;
    for Indice := 0 to TOBCde.detail.count -1 do
    begin
    	TOBR := TOBCde.detail[Indice];
      MontantPA := MontantPA + (TOBR.GetValue('GL_QTEFACT')*TOBR.GetValue('GL_DPA'));
      MontantPR := MontantPR + (TOBR.GetValue('GL_QTEFACT')*TOBR.GetValue('GL_PUHTDEV'));
      Qte := QTE + TOBR.GetValue('GL_QTEFACT');
    end;
    if QTE <> 0 then
    begin
      PAV := Arrondi(MontantPA / Qte,V_PGI.OkdecP);
    	PR := Arrondi(MontantPR / Qte,V_PGI.OkdecP);
    end;
    if PAV <> 0 then Coeffg := PR / PAV;
  end;
  if (PAA<>0) or (PAV<>0) then
  begin
  	// on a trouve un pa a appliquer
    if PAA <> 0 then TOBL.putValue('GL_DPA',PAA)
    						else TOBL.putValue('GL_DPA',PAV);
    if Coeffg <> 0 then
    begin
    	TOBL.putValue('GL_DPR',Arrondi(TOBL.GetValue('GL_DPA')*Coeffg,V_PGI.OkdecP));
    end else
    begin
      if PR <> 0 then TOBL.putValue('GL_DPR',PR)
                 else TOBL.putValue('GL_DPR',TOBL.GetValue('GL_DPA'));
    end;

  end else
  begin
  	// on a rien trouve
    // on prend ce qu'on trouve
    TOBL.putValue('GL_DPA',TOBA.GetValue('GA_PAHT'));
    TOBL.putValue('GL_DPR',TOBA.GetValue('GA_DPR'));
  end;

	if not VarIsNull (TOBL.GetValue('BCO_NUMMOUV')) then
  begin
    TOBConso.putValue('BCO_DPA',TOBL.GEtValue('GL_DPA'));
    TOBConso.putValue('BCO_DPR',TOBL.GEtValue('GL_DPR'));
  end else
  begin
  	if not NewConsoFromLigne (TOBL,TOBConso) then exit;
  end;
  TOBL.SetAllModifie (true);
  if TOBL.UpdateDB then
  begin
  	Mretour.Lines.add ('Livraison '+IntToStr(TOBL.GEtValue('GL_NUMERO'))+' Article : '+TOBL.GetValue('GL_CODEARTICLE')+' Mis à jour');
    if TOBL.GetValue('GL_DPA')=0 then
    begin
    	MRetour.lines.add ('	Attention ---> Prix à zéro');
    end;
  end;
  if TOBConso.GetValue('BCO_NUMMOUV')<>0 then
  begin
  	TOBConso.SetAllModifie (true);
  	TOBConso.InsertOrUpdateDB ;
  end;
  LibereToutCa;
end;

end.

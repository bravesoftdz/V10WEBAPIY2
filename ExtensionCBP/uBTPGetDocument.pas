unit uBTPGetDocument;

interface
uses hctrls,
  sysutils,HEnt1,
{$IFNDEF EAGLCLIENT}
  uDbxDataSet, DB,
{$ELSE}
  uWaini,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DB,
{$ENDIF}
  uhttp,
  uEntCommun,
  UtilTOBPiece,
  ParamSoc,
  utob;

procedure BTPGetDocument (TOBParams,TOBResult : TOB);

implementation

uses Classes,UtilBTPgestChantier,UmetresUtil,UtilBordereau, FactOuvrage;


procedure chargeLaPiece (Cledoc : r_cledoc; TOBREsult : TOB);
var TOBM,TOBD : TOB;
begin
  TOBM := TOB.Create ('TOBPIECE',TOBresult,-1);
  TOBD := TOB.Create ('PIECE',TOBM,-1);
  LoadPieceLignes(CleDoc, TobD,true,false);  // charge la PIECE et les LIGNES
end;

procedure  ChargeLesAffaires (TOBPiece,TOBREsult : TOB);
var TOBM,TOBD : TOB;
		QQ : TQuery;
begin
	TOBM := nil;
  if TOBPiece.getString('GP_AFFAIREDEVIS') <> '' then
  begin
  	TOBM := TOB.Create ('TOBAFFAIRES',TOBresult,-1);
    QQ := OpenSQL ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+TOBPiece.getString('GP_AFFAIREDEVIS')+'"',true,1,'',true);
    if not QQ.eof then
    begin
      TOBD := TOB.Create ('AFFAIRE',TOBM,-1);
      TOBD.SelectDB('',QQ);
    end;
    ferme (QQ);
  end;
  QQ := OpenSQL ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+TOBPiece.getString('GP_AFFAIRE')+'"',true,1,'',true);
  if not QQ.eof then
  begin
  	if TOBM = nil then   TOBM := TOB.Create ('TOBAFFAIRES',TOBresult,-1);
    TOBD := TOB.Create ('AFFAIRE',TOBM,-1);
    TOBD.SelectDB('',QQ);
  end;
  ferme (QQ);
end;

procedure ChargeLesBasesLignes (Cledoc: r_cledoc; TOBresult : TOB);
var TOBM : TOB;
		QQ : TQuery;
begin
  QQ := OpenSQL('SELECT * FROM LIGNEBASE WHERE ' + WherePiece(CleDoc, ttdLigneBase, False), True,-1, '', True);
  if not QQ.eof then
  begin
    TOBM := TOB.Create ('TOBBASESL',TOBresult,-1);
    TOBM.LoadDetailDB('LIGNEBASE', '', '', QQ, False);
  end;
  Ferme(QQ);
end;

procedure ChargeLesBases (cledoc : r_cledoc; TOBresult : TOB);
var TOBM : TOB;
		QQ : TQuery;
begin
  QQ := OpenSQL('SELECT * FROM PIEDBASE WHERE ' + WherePiece(CleDoc, ttdPiedBase, False), True,-1, '', True);
  if not QQ.eof then
  begin
	  TOBM := TOB.Create ('TOBBASES',TOBresult,-1);
  	TOBM.LoadDetailDB('PIEDBASE', '', '', QQ, False);
  end;
  Ferme(QQ);
end;

procedure ChargeLesEcheances (cledoc : r_cledoc; TOBresult : TOB);
var TOBM : TOB;
		QQ : TQuery;
begin
  // Lecture Echéances
  QQ := OpenSQL('SELECT * FROM PIEDECHE WHERE ' + WherePiece(CleDoc, ttdEche, False), True,-1, '', True);
  if not QQ.eof then
  begin
  	TOBM := TOB.Create ('TOBECHES',TOBresult,-1);
  	TOBM.LoadDetailDB('PIEDECHE', '', '', QQ, False);
  end;
  Ferme(QQ);
end;

procedure ChargeLesPorts (cledoc : r_cledoc; TOBresult : TOB);
var TOBM : TOB;
		QQ : TQuery;
begin
  // Lecture Ports
  QQ := OpenSQL('SELECT * FROM PIEDPORT WHERE ' + WherePiece(CleDoc, ttdPorc, False), True,-1, '', True);
  if not QQ.eof then
  begin
  	TOBM := TOB.Create ('TOBPORCS',TOBresult,-1);
  	TOBM.LoadDetailDB('PIEDPORT', '', '', QQ, False);
  end;
  Ferme(QQ);
end;

procedure ChargeLesAdresses(cledoc : r_cledoc; TOBPiece,TOBresult : TOB);
var TOBAdresses,TOBADR : TOB;
		QQ,Q : TQuery;
    NumL,NumF : integer;
begin
  if GetParamSoc('SO_GCPIECEADRESSE') then
  begin
    QQ := OpenSQL('SELECT * FROM PIECEADRESSE WHERE ' + WherePiece(CleDoc, ttdPieceAdr, False) + ' ORDER BY GPA_TYPEPIECEADR', True,-1, '', True);
    if not QQ.eof then
    begin
      TOBAdresses := TOB.Create ('TOBADRESSES',TOBresult,-1);
      TOBAdresses.LoadDetailDB('PIECEADRESSE', '', '', QQ, False);
      if TOBAdresses.Detail.Count = 1 then
      begin
        TOBAdr := TOB.Create('PIECEADRESSE', TOBAdresses, -1);
        TOBAdr.Dupliquer(TOBAdresses.Detail[0], False, True);
      end;
    end;
    Ferme(QQ);
  end else
  begin
    NumL := TOBPiece.GetValue('GP_NUMADRESSELIVR');
    QQ := OpenSQL('SELECT * FROM ADRESSES WHERE ADR_NUMEROADRESSE=' + IntToStr(NumL), True,-1, '', True);
    if not QQ.eof then
    begin
      TOBAdresses := TOB.Create ('TOBADRESSES',TOBresult,-1);
      TOB.Create('PIECEADRESSE', TOBAdresses, -1);
      TOB.Create('PIECEADRESSE', TOBAdresses, -1);
      TOBAdresses.Detail[0].SelectDB('', QQ);
      NumF := TOBPiece.GetValue('GP_NUMADRESSEFACT');
      if NumF = NumL then
      begin
        TOBAdresses.Detail[1].Dupliquer(TOBAdresses.Detail[0], True, True);
      end else
      begin
        Q := OpenSQL('SELECT * FROM ADRESSES WHERE ADR_NUMEROADRESSE=' + IntToStr(NumF), True,-1, '', True);
        if not QQ.eof then
        begin
          TOBAdresses.Detail[1].SelectDB('', QQ);
        end;
        ferme (Q);
      end;
    end;
    Ferme(QQ);
  end;
end;

procedure ChargeLesAcomptes (Cledoc : r_cledoc; TOBPiece,TOBresult : TOB);
var TOBM : TOB;
		QQ : TQuery;
    GereAcompte : boolean;
begin
	GereAcompte := false;
	QQ := OpenSql ('SELECT GPP_ACOMPTE FROM PARPIECE WHERE GPP_NATUREPIECEG="'+Cledoc.NaturePiece+'"',true,1,'',true);
  if not QQ.eof then GereAcompte := (QQ.FindField('GPP_ACOMPTE').AsString = 'X');
  ferme (QQ);

  if ((not GereAcompte) and (TOBPiece.GetValue('GP_ACOMPTEDEV') = 0)) then Exit;
  QQ := OpenSQl('SELECT * FROM ACOMPTES WHERE ' + WherePiece(CleDoc, ttdAcompte, False), True,-1, '', True);
  if not QQ.EOF then
  begin
  	TOBM := TOB.Create ('TOBACOMPTES',TOBresult,-1);
  	TOBM.LoadDetailDB('ACOMPTES', '', '', QQ, False);
  end;
  Ferme(QQ);
end;

procedure ChargeLesAna (cledoc : r_cledoc; TOBPiece,TOBResult : TOB);

	function encodeRef (cledoc : r_cledoc) : string;
  begin
    Result := Cledoc.NaturePiece
            + ';' + Cledoc.Souche
            + ';' + FormatDateTime('ddmmyyyy',Cledoc.DatePiece)
            + ';' + IntToStr(cledoc.NumeroPiece)
            + ';' + IntToStr(Cledoc.NumOrdre)+';';
  end;

Var Q : TQuery ;
    TOBM,TOBANAP,TOBANAS,TOBA : TOB ;
    i: integer ;
    RefA   : String ;
BEGIN

  RefA:=EncodeRef(cledoc) ; if RefA='' then Exit ;
  Q:=OpenSQL('SELECT * FROM VENTANA WHERE (YVA_TABLEANA="GL" OR YVA_TABLEANA="GS") AND YVA_IDENTIFIANT="'+RefA+'"',True,-1, '', True) ;
  if Not Q.EOF then
  BEGIN
  	TOBM := TOB.Create ('TOBANALS',TOBresult,-1);
    TOBAnaP := TOB.Create ('TOBANAP',TOBM,-1);
    TOBAnaS := TOB.Create ('TOBANAS',TOBM,-1);
    //
    TOBAnaP.LoadDetailDB('VENTANA','','',Q,False) ;
    for i:=TOBAnaP.Detail.Count-1 downto 0 do
    BEGIN
      TOBA:=TOBAnaP.Detail[i] ;
      if TOBA.GetValue('YVA_TABLEANA')='GS' then TOBA.ChangeParent(TOBAnaS,-1) ;
    END ;
    TOBAnaP.Detail.Sort('-YVA_AXE;-YVA_IDENTLIGNE') ;
    TOBAnaS.Detail.Sort('-YVA_AXE;-YVA_IDENTLIGNE') ;
  END ;
  Ferme(Q) ;
end;

procedure ChargeLesBlocNotes (Cledoc : r_cledoc; TOBresult : TOB);
var Q : TQuery;
		TOBM : TOB;
begin
  Q := OpenSQL('SELECT * FROM LIENSOLE WHERE ' + WherePiece(CleDoc, ttdLienOle, False), True,-1, '', True);
  if not Q.eof then
  begin
    TOBM := TOB.Create ('TOBBLOCNOTES',TOBresult,-1);
    TOBM.LoadDetailDB('LIENSOLE', '', '', Q, False);
  end;
  Ferme(Q);
end;

procedure ChargelesRG (cledoc : r_cledoc; TOBresult : TOB);
var Q : TQuery;
		TOBM,TOBM1 : TOB;
    i : integer;
begin
  Q := OpenSQL('SELECT * FROM PIECERG WHERE ' + WherePiece(CleDoc, ttdretenuG, False), True,-1, '', True);
  if not Q.eof then
  begin
  	TOBM := TOB.Create ('RGS',TOBresult,-1);
    TOBM1 := TOB.Create ('TOBPIECERG',TOBM,-1);
    TOBM1.loaddetailDB('PIECERG', '', '', Q, False, True);
    Ferme(Q);
    //
    for i := 0 to TOBM1.detail.count - 1 do TOBM1.detail[i].addchampsupValeur('INDICERG', 0);
    // Chargement de la tob des Bases de tva retenues de garantie
    Q := OpenSQL('SELECT * FROM PIEDBASERG WHERE ' + WherePiece(CleDoc, ttdBaseRG, False), True,-1, '', True);
    if not Q.eof then
    begin
    	TOBM1 := TOB.Create ('TOBBASERG',TOBM,-1);
    	TOBM1.loaddetaildb('PIEDBASERG', '', '', Q, False);
    end;
  end;
  Ferme(Q);
end;

procedure ChargeLesOuv (cledoc : r_cledoc;TOBresult: TOB);
var Q : TQuery;
		TOBM : TOB;
    requete : string;
begin
	TOBM := TOB.Create ('TOBOUVRAGES',TOBresult,-1);
  ChargeLesOuvrages (TobResult.Detail[0].Detail[0], TobM,cledoc);
(*
	if not IsOuvrageOkInPiece (cledoc.NaturePiece) then exit;
  requete :='SELECT O.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM LIGNEOUV O '+
          'LEFT JOIN NATUREPREST N ON BNP_NATUREPRES=(SELECT GA_NATUREPRES FROM ARTICLE WHERE GA_ARTICLE=O.BLO_ARTICLE) '+
          'WHERE '+WherePiece(CleDoc,ttdOuvrage,False)+
          ' ORDER BY BLO_NUMLIGNE,BLO_N1, BLO_N2, BLO_N3, BLO_N4,BLO_N5';
  Q:=OpenSQL(requete,True,-1, '', True) ;
  if not Q.eof then
  begin
  	TOBM := TOB.Create ('TOBOUVRAGES',TOBresult,-1);
  	TOBM.LoadDetailDB ('LIGNEOUV','','',Q,True,true);
  end;
  Ferme(Q) ;
*)
end;

procedure ChargeLesOuvragePlat (cledoc : r_cledoc; TOBresult : TOB);
var Q : TQuery;
		TOBM : TOB;
begin
  Q := OpenSQl('SELECT * FROM LIGNEOUVPLAT WHERE ' + WherePiece(CleDoc, ttdOuvrageP, False), True,-1, '', True);
  if not Q.EOF then
  begin
  	TOBM := TOB.Create ('TOBOUVRAGESP',TOBresult,-1);
  	TOBM.LoadDetailDB('LIGNEOUVPLAT', '', '', Q, False);
  end;
  Ferme(Q);
end;

procedure ChargeLesArticles (cledoc : r_cledoc; TOBresult : TOB);
var Q : TQuery;
		TOBM : TOB;
    TOBI,TOBII : TOB;
    StSelect,StWhereLigne,StWhereOuvrage : string;
    Indice : integer;
begin
  //
	TOBM := nil;
	TOBI := TOB.Create ('TOBARTICLES',nil,-1);
  TRY
    StSelect := 'SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
              'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
              'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE';
    // 1ere passe --> provenance des lignes de documents
    stWhereLigne := WherePiece(CleDoc, ttdLigne, False);
    Q := OpenSQL(StSelect + ' WHERE GA_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne + ') ORDER BY GA_ARTICLE', True,-1, '', True);
    if not Q.eof then
    begin
      TOBM := TOB.Create ('TOBARTICLES',TOBresult,-1);
      TOBM.LoadDetailDB('ARTICLE', '', '', Q, True, True);
    end;
    Ferme(Q);

    // 2eme passe --> provenance des lignes de détail ouvrage
    stWhereOuvrage := WherePiece(CleDoc, ttdOuvrage, False);
    Q := OpenSQL(StSelect + ' WHERE GA_ARTICLE IN (SELECT BLO_ARTICLE FROM LIGNEOUV WHERE ' + stWhereOuvrage + ') ORDER BY GA_ARTICLE', True,-1, '', True);
    if not Q.eof then
    begin
    	if not Assigned(TOBM) then TOBM := TOB.Create ('TOBARTICLES',TOBresult,-1);
      //
      TOBI.LoadDetailDB('ARTICLE', '', '', Q, True, True);
      Indice := 0;
      repeat
        TOBII := TOBI.detail[Indice];
        if TOBM.FindFirst(['GA_ARTICLE'],[TOBII.GetString('GA_ARTICLe')],true) = nil then
        begin
        	TOBII.ChangeParent(TOBM,-1);
        end else inc(Indice);
      Until Indice > TOBI.detail.Count -1;
    end;
    Ferme(Q);
    //

  FINALLY
  	TOBI.free;
  END;
end;

procedure  ChargeLesDispo (TOBPiece,TOBresult: TOB; cledoc :r_cledoc);
var stselectdepot,ListeDepot ,StWhereLigne: string;
		Q : TQuery;
    DoubleDepot : boolean;
    TOBM : TOB;
begin
	StSelectDepot := 'SELECT * FROM DISPO';
  stWhereLigne := WherePiece(CleDoc, ttdLigne, False);
	DoubleDepot := IsTransfert(TOBPiece.getString('GP_NATUREPIECEG') );
  if DoubleDepot then
    ListeDepot := '"' + TOBPiece.GetValue('GP_DEPOT') + '","' + TOBPiece.GetValue('GP_DEPOTDEST') + '"'
  else ListeDepot := '"' + TOBPiece.GetValue('GP_DEPOT') + '"';
  if DoubleDepot then
  begin
    Q := OpenSQL(StSelectDepot + ' WHERE GQ_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne +
      ') AND GQ_DEPOT IN (' + ListeDepot + ') AND GQ_CLOTURE="-"', True,-1, '', True);
  end else
  begin
  	Q := OpenSQL(StSelectDepot + ' WHERE GQ_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne +
      ') AND GQ_DEPOT IN (SELECT DISTINCT GL_DEPOT FROM LIGNE WHERE ' + stWhereLigne + ') AND GQ_CLOTURE="-"', True,-1, '', True);
  end;
  if not Q.eof then
  begin
    TOBM := TOB.Create ('TOBDISPO',TOBresult,-1);
    TOBM.LoadDetailDB('DISPO', '', '', Q, True, True);
  end;
  ferme (Q);
end;

procedure ChargeLesFournisseurs (cledoc: r_cledoc; TOBresult : TOB);
var select : string;
		TOBM : TOB;
		Q : TQuery;
begin
  Select := 'SELECT * FROM TIERS WHERE T_TIERS '+
  					'IN (SELECT DISTINCT BLO_FOURNISSEUR FROM LIGNEOUV WHERE '+
            WherePiece(CleDoc,ttdOuvrage,False)+' AND BLO_FOURNISSEUR<>"")';
  Q:=OpenSQL(Select,True,-1, '', True) ;
  if not Q.eof then
  begin
    TOBM := TOB.Create ('TOBFOURNISSEURS',TOBresult,-1);
  	TOBM.LoadDetailDB ('TIERS','','',Q,True,True);
  end;
  ferme(Q);
end;

function ChargeLeTiers (TOBPiece,TOBresult: TOB; cledoc : r_cledoc) : TOB;
  procedure InitTiersCompl(Auxiliaire, CodeTiers: String);
  var
    TobTiersCompl: Tob;
  begin
    if (Auxiliaire <> '') and (CodeTiers <> '') then
    begin
      TobTiersCompl := Tob.Create('TIERSCOMPL', nil, -1);
      try
         TobTiersCompl.P('YTC_AUXILIAIRE', Auxiliaire);
         TobTiersCompl.P('YTC_TIERS', CodeTiers);
         TobTiersCompl.InsertDB(nil);
      finally
        TobTiersCompl.Free;
      end;
    end;
  end;

var Q : TQuery;
		CodeTiers,Naturepiece : string;
    TOBM,TOBD : TOB;
begin
	result := nil;
	CodeTiers := TOBPiece.getString('GP_TIERS');
	Naturepiece := TOBPiece.getString('GP_NATUREPIECEG');
  if CodeTiers = '' then exit;
  Q := OpenSQL('SELECT * FROM TIERS LEFT JOIN TIERSCOMPL ON T_TIERS=YTC_TIERS '+
  						 'WHERE T_TIERS="' + CodeTiers + '"', True,-1, '', True);
  if not Q.EOF then
  begin
  	TOBM := TOB.Create ('TOBTIERS',TOBresult,-1);
    TOBD := TOB.Create ('TIERS',TOBM,-1);
    TOBD.selectDb ('',Q);
    result := TOBD;
    if (TobD.G('YTC_AUXILIAIRE')='') or (TobD.G('YTC_TIERS')='') then
    begin
      InitTiersCompl(TobD.G('T_AUXILIAIRE'), TobD.G('T_TIERS'))
    end;
    ferme (Q);
    //
    Q := OpenSQL('SELECT * FROM TIERSPIECE WHERE GTP_TIERS="' + CodeTiers + '" AND GTP_NATUREPIECEG="' + NaturePiece + '"', True,-1, '', True);
    if not Q.eof then
    begin
    	TOBD.LoadDetailDB('TIERSPIECE', '', '', Q, False);
    end;
  end;
  ferme (Q);
end;

procedure ChargeLeRIB (TOBTiers,TOBresult : TOB);
var Auxi,sRIB : string;
		Q : TQuery;
    TOBM,TOBD : TOB;
begin
	if TOBTiers = nil then exit;
  Auxi := TOBTiers.GetValue('T_AUXILIAIRE');
  sRib := '';
  Q := OpenSQL('SELECT * FROM RIB WHERE R_AUXILIAIRE="' + Auxi + '" AND R_PRINCIPAL="X"', True,-1, '', True);
  if not Q.eof then
  begin
  	TOBM := TOB.Create ('TOBTIERSRIB',TOBresult,-1);
    TOBD := TOB.Create ('RIB',TOBM,-1);
    TOBD.SelectDB('',Q);
  end;
  Ferme(Q);
end;

procedure ChargeLesCommerciaux (cledoc : r_cledoc;TOBresult : TOB);
var Req,CodeComm : string;
		QQ : TQuery;
    TOBM,TOBDD : TOB;
    I : integer;
begin
	TOBM := nil;
	Req := 'SELECT * FROM COMMERCIAL WHERE GCL_COMMERCIAL IN (SELECT GL_REPRESENTANT FROM LIGNE WHERE '+
  				WherePiece(CleDoc,ttdligne,False)+' AND GL_REPRESENTANT <>"")';
  QQ := OPenSql (Req,true,-1,'',true);
  if not QQ.eof then
  begin
  	TOBM := TOB.Create ('TOBCOMMS',TOBresult,-1);
    TOBM.LoadDetailDb('COMMERCIAL','','',QQ,false);
  end;
  ferme (QQ);
  if Assigned(TOBM) then
  begin
    for I:=0 to TOBM.detail.count -1 do
    begin
      TOBDD := TOBM.detail[I]; CodeComm := TOBDD.GetString('GCL_COMMERCIAL');
      QQ := OpenSQL('SELECT * FROM COMMISSION WHERE GCM_COMMERCIAL="' + CodeComm + '"', True,-1, '', True);
      if not QQ.eof then
      begin
        TOBM.LoadDetailDB('COMMISSION', '', '', QQ, False, True);
      end;
      Ferme(QQ);
    end;
  end;
end;

procedure ChargeLesLienDevCha (cledoc : r_cledoc; TOBPiece,TOBresult : TOB);
var Req : string;
		TOBM : TOB;
		QQ : TQuery;
begin
	if cledoc.Naturepiece = 'DBT' then
  begin
    Req := 'SELECT * FROM LIENDEVCHA WHERE BDA_REFD="'+EncodeLienDevCHA(TOBPiece)+'"';
  end else
  begin
    Req := 'SELECT * FROM LIENDEVCHA WHERE BDA_REFC="'+EncodeLienDevCHA(TOBPiece)+'" ORDER BY BDA_NUMLC' ;
  end;
  QQ := OpenSql (Req,true,-1,'',true);
  if not QQ.eof then
  begin
  	TOBM := TOB.Create ('TOBLIENDEVCHA',TOBresult,-1);
  	TOBM.LoadDetailDB ('LIENDEVCHA','','',QQ,false,false);
  end;
  ferme (QQ);
end;

procedure ChargeLesmetres (cledoc: r_cledoc; TOBPiece,TOBresult : TOB);
var Laref,Select : string;
		QQ : TQuery;
    TOBM : TOB;
begin
  LaRef := GenereLaReferenceMetre (TOBPiece,taModif,false);
  Select := 'SELECT * FROM LIENSOLE WHERE (LO_TABLEBLOB="MTR") AND (LO_IDENTIFIANT LIKE "'+LaRef + '%")';
  QQ := OpenSql (Select,true);
  if not QQ.eof then
  begin
  	TOBM := TOB.Create ('TOBMETRES',TOBresult,-1);
  	TOBM.LoadDetailDB ('LIENSOLE','','',QQ,false);
  end;
  ferme (QQ);
end;

procedure ChargeLesBordereaux (cledoc : r_cledoc; TOBPiece,TOBTiers,TOBresult : TOB);
var requete,requeteref : string;
		Q : Tquery;
    TOBM : TOB;
begin
	requete := ConstitueRequeteBordereau (cledoc.NaturePiece ,TOBPiece.getString('GP_AFFAIRE'),TOBTIers.getString('T_TIERS'),
  																			TOBTiers.getString('T_NATUREAUXI'),TOBPiece.GetDateTime('GP_DATEPIECE'),requeteRef);
  if Requete = '' then exit;
  Q := OpenSql (Requete,true,-1,'',true);
  if not Q.eof then
  begin
  	TOBM := TOB.Create ('TOBBORDEREAUX',TOBresult,-1); TOBResult.AddChampSupValeur('REQUETEREF',RequeteRef);
    TOBM.LoadDetailDB('BDETETUDE','','',Q,false);
  end;
  ferme(Q);
end;

procedure ChargeLesTOBS (Cledoc : r_cledoc;TOBresult : TOB);
var TOBPiece,TOBTiers : TOB;
begin
(* -------------------------------------------------------
on va constiture la TOB de sortie de la maniere suivante
chaque TOB (ex TOBPIECE,TOBBASES,TOBBASESL,TOBECHE,...)
sera créée dans une tob mère virtuelle portant son nom
--------------------------------------------------------- *)
  ChargeLaPiece (Cledoc,TOBREsult);
  TOBPiece := TOBresult.detail[0].detail[0];
  TOBTiers := ChargeLeTiers (TOBPiece,TOBresult,cledoc);
  ChargeLeRIB (TOBTiers,TOBresult);
// en premier cela permet lors de la récupération d'avoir de suite les elements de valoraisation en ligne
  ChargeLesArticles (cledoc, TOBresult);
  ChargeLesFournisseurs (cledoc, TOBresult);
  ChargeLesDispo (TOBPiece,TOBresult,cledoc);
// ----------------
  ChargeLesAffaires (TOBPiece,TOBREsult);
  ChargeLesBasesLignes (cledoc,TOBresult);
  ChargeLesBases (cledoc,TOBresult);
  ChargeLesEcheances (cledoc,TOBresult);
  ChargeLesPorts (cledoc,TOBresult);
  ChargeLesAdresses (cledoc,TOBPiece,TOBresult);
  ChargeLesAcomptes (Cledoc ,TOBPiece,TOBresult);
  ChargeLesAna (cledoc,TOBPiece,TOBResult);
  ChargeLesBlocNotes (Cledoc,TOBresult);
  ChargelesRG (cledoc,TOBresult);
  ChargeLesOuv (cledoc,TOBresult);
  ChargeLesOuvragePlat (cledoc,TOBresult);
  ChargeLesCommerciaux (cledoc,TOBresult);
  ChargeLesLienDevCha (cledoc,TOBPiece,TOBresult);
  ChargeLesmetres (cledoc,TOBPiece,TOBresult);
  ChargeLesBordereaux (cledoc,TOBPiece,TOBTiers,TOBresult);
end;



procedure BTPGetDocument (TOBParams,TOBResult : TOB);
var cledoc : r_cledoc;
		ViewParam : string;
begin
  if (TOBParams.FieldExists('CLEDOC')) and (TOBParams.GetString('CLEDOC') <> '') then
  begin
  	TRY
    	ViewParam := TOBParams.GetString('CLEDOC');
  		DecodeRefPiece  (ViewParam,Cledoc);
    EXCEPT
  		TOBREsult.SetString('ERROR','Chargement de pièce : Document incorrectement défini');
      exit;
    END;
    ChargeLesTOBS (Cledoc,TOBresult);
  end else
  begin
  	TOBREsult.SetString('ERROR','Chargement de pièce : Document non défini');
  end;
end;

end.

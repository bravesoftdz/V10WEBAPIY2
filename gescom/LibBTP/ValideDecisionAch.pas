unit ValideDecisionAch;

interface                                                         

uses Classes,UTOB,FactComm,Hctrls,HEnt1,SaisUtil,UtilPGi,SysUtils,Forms,
{$IFDEF EAGLCLIENT}
  maineagl
{$ELSE}
  Doc_Parser,DBCtrls, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} fe_main
{$ENDIF}
  ,ParamSoc,AglInit,FactCommBTP,Controls
	,HmsgBox,Entgc,Splash,FactTOB,UtilValidReapDecis,FactUtil,FactPiece,FactCalc,UtofListeInv,
  UtilArticle,uEntCommun
;

Type
MerrCreatFour = (MerdArt,MerdFourn,MerdCreat,MerdAucune);
RRefArtFour = record
	codeArtFour : string;
  Libelle : string;
  DelaiFou : integer;
end;

TvalidDecision = class
	private
  	fChantierTermine : boolean;
  	fusable : boolean;
  	fNumDecisionel : integer;
    fTOBEntree : TOB;
    fiche : TFsplashScreen;
    TOBPieces,TOBBases,TOBBasesL,TOBEches,TOBTiers,TOBDISPO : TOB;
    TOBCdeFou : TOB;
    TOBArticles : TOB;
    TOBResult : TOB;
    Mode : string;
    DEVbesoin : RDevise;
    procedure ArticleVersLigne (TOBA,TOBL : TOB);
    procedure ChargeLesTOBs;
    function CreerLigneCdeFou (TOBD : TOB) : MerrCreatFour;
    procedure DecisionnelVersLigne (TOBD,TOBL : TOB; TheRefArtFourn : RRefArtFour);
    function FindTiers (CodeTiers,TypeTier : String) : TOB;
    function FindArticle (CodeArticle : String) : TOB;
    function isAnnulee (TOBL : TOB) : boolean;
    function isRemplacee (TOBL : TOB) : boolean;
    procedure InitLigne (TOBL,TOBART,TOBFourn : TOB);
    procedure GenereCdesFourn;
    procedure PrepareTObs;
    procedure FindLienFournArt(TOBD: TOB; var TheRefArtFour: RRefArtFour);
    procedure ClotureDecisionnel ;
    procedure InitLigneBesoin(TOBL: TOB);
    procedure DecisionnelVersLigneBesoin(TOBD, TOBL, TOBA: TOB);
    function EnregistreLesBesoins : boolean;
    procedure findTOBs(reference: string; var TOBPiece, TOBbase,TOBeche: TOB);
    procedure ChargeBesoin(reference: string; var TOBPiece, TOBbase, TOBeche,TOBTier: TOB);
    procedure CreeEntree(reference: string; var TOBpiece, TObbase,TOBeche: TOB);
    procedure chargeAjouteArticles(Cledoc: r_cledoc);
    procedure LoadLesTOB(cledoc: r_cledoc; TOBPiece, TOBBase,TOBEche: TOB);
    procedure RemplacementArticle(TOBD, TOBDecisionnel, TOBPiece, TOBBase,TOBeche, TOBTier: TOB);
    procedure SoldeLignePartiel(reference: string; TOBpiece, TOBBase,TOBEche, TOBTier: TOB; Quantite: double);
    function IsScindee (TOBR : TOB) : boolean;
    procedure MajLigneBesoin(TOBR, TOBPiece, TOBBase, TOBeche,TOBTier: TOB);
    procedure GereScindeBesoin (TOBD,TOBPiece, TOBBase,TOBeche, TOBTier : TOB);
    procedure VireLesEltNonNecessaires(TOBPiece: TOB);
    function supprimePiece(TOBPiece: TOB): boolean;
    procedure AjouteArticleDansbesoin(TOBR, TOBPiece, TOBBase, TOBTier,TOBREMPL: TOB; IndiceInsert: integer);
    function AnnuleArticleDansBesoin(TOBD, TOBPiece: TOB;var TOBREMPL: TOB): integer;
    procedure PositionneAncienneValeurs(TOBPiece: TOB);
    procedure ChargeAjouteDispo(Cledoc: r_cledoc);
    function FindDispo(Article, Depot,GereStock: string): TOB;
    procedure ReajusteReserve;
    function AddDispo(TOBD: TOB): TOB; overload;
    function AddDispo(Article,Depot : string): TOB; overload;
    function isRemplaceant(TOBL: TOB): boolean;

  public
  	property NumDecisionnel : integer read fNumDecisionel write fNumDecisionel;
    procedure lanceValidation;
    constructor create;
    destructor destroy; override;
end;

procedure ValideDecisionnel (Decisionnel : integer);

implementation
uses facttiers,FactGRp,FactLignebase,affaireUtil,UtilTOBPiece;

procedure ValideDecisionnel (Decisionnel : integer);
var ValidDecis : TvalidDecision;
begin
  ValidDecis :=TvalidDecision.create;
  ValidDecis.NumDecisionnel := Decisionnel;
  TRY
  	ValidDecis.lanceValidation;
  FINALLY
  	ValidDecis.free;
  END;
end;

{ TvalidDecision }

procedure TvalidDecision.ChargeLesTOBs;
var QQ : TQuery;
begin
	fusable := false;
  QQ := OPenSQL ('SELECT * FROM DECISIONACH WHERE BAE_NUMERO='+inttoStr(fNumDecisionel),true);
  if not QQ.eof then
  begin
     fusable := true;
     fTOBEntree.selectdb ('',QQ);
  end;
  ferme (QQ);
  if not fusable then
  begin
  	PGIBOX (TraduireMemoire('Erreur durant le chargement'),'Validation');
    exit;
  end;
  QQ := OPenSQL ('SELECT * FROM DECISIONACHLIG WHERE BAD_NUMERO='+inttoStr(fNumDecisionel),true);
  if not QQ.eof then
  begin
     fTOBEntree.LoadDetailDB ('DECISIONACHLIG','','',QQ,false,true);
  end else
  begin
  	fusable := false;
  	PGIBOX (TraduireMemoire('Erreur durant le chargement des lignes'),'Validation');
  end;
  ferme (QQ);
end;

constructor TvalidDecision.create;
begin
	fiche := TFsplashScreen.create (application);
  fTOBEntree := TOB.Create ('DECISIONACH',nil,-1);
  TOBCdeFou := TOB.Create ('LES CDE FOURN',nil,-1);
  TOBresult := TOB.Create ('LES PARAMS',nil,-1);
  TOBPieces := TOB.Create ('LES PIECES',nil,-1);
  TOBBases := TOB.Create('LES BASES', nil, -1);
  TOBBasesL := TOB.Create('LES BASES LIGNES', nil, -1);
  TOBTiers := TOB.Create('LES TIERS', nil, -1);
  TOBArticles := TOB.Create('ARTICLES', nil, -1);
  TOBEches := TOB.Create('Les ECHEANCES', nil, -1);
  TOBDispo := TOB.Create ('LES STOCK',nil,-1);
  fChantierTermine := false;
end;

destructor TvalidDecision.destroy;
begin
  inherited;
  fiche.free;
  fTOBEntree.free;
  TOBCdeFou.free;
  TOBArticles.free;
  TOBresult.free;
  TOBPieces.free;
  TOBBases.free;
  TOBBAsesL.free;
  TOBTiers.free;
  TOBEches.free;
  TOBDISPO.free;
	ReinitTOBAffaires;
end;

procedure TvalidDecision.lanceValidation;
begin
  if GetparamSoc('SO_BTECLATCDEFOU') then TOBResult.AddChampSupValeur ('ECLAT','X')
  																	 else TOBResult.AddChampSupValeur ('ECLAT','-');
  TOBResult.AddChampSupValeur ('DELAISECURITE',0);
  TOBResult.AddChampSupValeur ('OK','-');
  TheTOB := TOBresult;
  AglLanceFiche ('BTP','BTDECISIONACH_VAL','','','ACTION=MODIFICATION');
  TheTOB := nil;
  if TOBresult.getValue('OK') = 'X' then
  begin
    if TOBresult.GetValue('ECLAT')<>'X' then Mode:='VALIDDECIS' else Mode := 'VALIDDECISECLAT';
    fiche.label1.caption := 'Chargement du décisionnel...';
    fiche.Show;
    fiche.Refresh;
    ChargeLesTOBs;
    if not fusable then BEGIN exit; END;
    fiche.label1.caption := 'Chargement des besoins...';
    fiche.Refresh;
    PrepareTObs;
    fiche.label1.caption := 'Création des commandes fournisseurs...';
    fiche.Refresh;
    if TRANSACTIONS (GenereCdesFourn,1) <> oEOk then
    begin
    	PgiBox (Traduirememoire('Une erreur s''est produite lors de la fin de la procédure'),'Validation');
    end else
    begin
    	if fChantierTermine then
      begin
    		PgiInfo (Traduirememoire('Des lignes en provenance de chantier terminé ont été ignoré'),'Validation');
      end;
    end;
  end;
end;

procedure TvalidDecision.SoldeLignePartiel(reference : string; TOBpiece,TOBBase,TOBEche,TOBTier : TOB; Quantite : double);
var cledoc : r_cledoc;
		TOBL,TOBD : TOB;
    FUV,FUS,RatioStk : double;
begin
	DecodeRefPiece (reference,cledoc);
  if TOBPiece = nil then exit;
	TOBL := TOBPiece.findFirst (['GL_NUMORDRE'],[cledoc.NumOrdre],true);
  if TOBL <> nil then
  begin
    FUV := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTEVTE')); if FUV = 0 then FUV := 1;
    FUS := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTESTO')); if FUS = 0 then FUS := 1;
    RatioStk := FUV/FUS;

    TOBD := FindDispo (TOBL.GetValue('GL_ARTICLE'),TOBL.GEtValue('GL_DEPOT'),TOBL.GetValue('__TENUESTOCK'));
  	if Quantite > TOBL.GetValue('GL_QTERESTE') then
    begin
    	if TOBD <> nil then
      begin
    		TOBD.PutValue('__QTEMOINS',TOBD.GEtVAlue('__QTEMOINS')+(TOBL.GetValue('GL_QTERESTE')*RatioStk));
      end;
    	TOBL.putValue('GL_QTERESTE',0);
    end else
    begin
    	if TOBD <> nil then
      begin
    		TOBD.PutValue('__QTEMOINS',TOBD.GEtVAlue('__QTEMOINS')+(Quantite*RatioStk));
      end;
    	TOBL.PutValue('GL_QTERESTE',TOBL.GetValue('GL_QTERESTE')-Quantite); // solde la ligne
    end;
  end;
end;

procedure TvalidDecision.PrepareTObs;
var Indice : integer;
    TOBD : TOB;
    Tri,RuptureDoc,RuptureArt : string;
    TOBPiece,TOBBase,TOBeche,TOBTier : TOB;
    EtatAffaire : string;
begin
  for Indice := 0 to fTOBEntree.detail.count -1 do
  begin
  	TOBD := fTOBEntree.detail[Indice];
    if TOBD.GetValue('BAD_NUMN4')<> 0 then
    begin
    	if TOBD.GetValue('BAD_AFFAIRE') <> '' then
      begin
        EtatAffaire := GetChampsAffaire (TOBD.GetString('BAD_AFFAIRE'),'AFF_ETATAFFAIRE');
        if EtatAffaire='TER' then
        begin
        	fChantierTermine := true;
          Continue;
        end;
      end;
    	// On ne traite pas les lignes de cumuls
    	if TOBD.getValue('BAD_REFGESCOM') <> '' then
      begin
      	ChargeBesoin (TOBD.GetValue('BAD_REFGESCOM'),TOBPiece,TOBbase,TOBEche,TOBTier);

        // modif BRL 1/04/2010 : si jamais le besoin n'existe plus, on supprime l'élément dans la tob
        // pour éviter une erreur lors du traitement de génération ((pb POUCHAIN)
        if TOBPiece.detail.count = 0 then
        begin
          TOBPiece.free;
          Continue;
        end;

        if isAnnulee(TOBD) then
        begin
          SoldeLignePartiel (TOBD.GetValue('BAD_REFGESCOM'),TOBPiece,TOBbase,TOBEche,TOBTier,TOBD.getValue('BAD_QUANTITEINIT'));
        end else if isRemplacee (TOBD) then
        begin
          RemplacementArticle (TOBD,fTOBEntree,TOBPiece, TOBBase,TOBeche, TOBTier);
        end else if IsScindee (TOBD) then
        begin
        	GereScindeBesoin (TOBD,TOBPiece, TOBBase,TOBeche, TOBTier);
        end else if not isRemplaceant (TOBD) then // on saute les article remplaceant (deja géré)
        begin
        	MajLigneBesoin (TOBD,TOBPiece, TOBBase,TOBeche, TOBTier);
        end;
      end;
      if CreerLigneCdeFou (TOBD) <> MerdAucune then PgiBox('Une erreur s''est produite durant la création','VALIDATION');
    end;
  end;
  // génération
  ConstitueChampsTriRupt_UtilVRD (Tri,RuptureDoc,RuptureArt,(ToBresult.getValue('ECLAT')='X'));
	TrielaTOB_UtilVRD(TOBCdeFou,Tri);
  ConstitueLaTOBFinale_UtilVRD (TOBCdeFou,Tri,RuptureDoc,RuptureArt,Mode,0);
end;

function TvalidDecision.isAnnulee(TOBL: TOB): boolean;
begin
	result := (TOBL.GetValue('BAD_SUPPRIME')='X');
end;

procedure TvalidDecision.GenereCdesFourn;
var Mrres : boolean;
		WithDemande : boolean;
    ResultTraite : boolean;
    MessageInfo : string;
begin
	WithDemande := false;
  ResultTraite := true;
	V_PGI.IOError := OeOk;
  if TOBCdeFou.detail.count > 0 then mrRes:= CreerPiecesFromLignes(TOBCdeFou, Mode,iDate1900) else WithDemande := true;
  if WithDemande then
  begin
  	MessageInfo := 'Aucune commande fournisseur ne sera générée #13#10 '+
    							 'Désirez vous appliquer quand même les modifications et valider le décisionnel?';
  	if PGIAsk (TraduireMemoire(MessageInfo),'Validation') <> Mryes then ResultTraite := false;
  end;
  if ResultTraite = true then
  begin
    if (V_PGI.Ioerror = OeOk) and (not EnregistreLesBesoins ) Then V_PGI.Ioerror := oeUnknown;
    if (V_PGI.Ioerror = OeOk) THEN ClotureDecisionnel;
    if (V_PGI.ioError = OeOk) then ReajusteReserve;
  end;
end;

function TvalidDecision.CreerLigneCdeFou(TOBD: TOB) : MerrCreatFour;
var TheRefArtFour : RRefArtFour;
		TOBFourn,TOBArt,TOBL : TOB;
    Fournisseur,NAturePiece  : string;
begin
	result := MerdAucune;
  Fournisseur := TOBD.GetValue('BAD_FOURNISSEUR');

  Naturepiece := 'CF';
  if (TOBD.GetValue('BAD_PRISENCOMPTE') = '-') then exit;
  if Fournisseur='' then exit;
  // Dans ce cas la tout vaaaa
  TOBART := FindArticle (TOBD.GetValue('BAD_ARTICLE'));
  if (TOBART = nil)  then BEGIN result := MerdArt; Exit; END;
  TOBFourn := FindTiers (Fournisseur,'FOU');
  if (TOBFourn = nil)  then
  BEGIN
  	TOBFourn := TOB.Create ('TIERS',TOBTiers,-1);
    TOBFourn.AddChampSupValeur('RIB','', False);
		if (RemplirTOBTiers (TOBFourn,Fournisseur,Naturepiece,false)<> trtOk) then BEGIN result := MerdFourn; Exit; END;
  end;
  FindLienFournArt(TOBD,TheRefArtFour);
  TOBL := TOB.Create ('LIGNE',TOBCdeFou,-1);
  AddLesSupLigne (tobl,false);
  InitLigne (TOBL,TOBART,TOBFourn);
  ArticleVersLigne (TOBArt,TOBL);
  DecisionnelVersLigne (TOBD,TOBL,TheRefArtFour);
end;

procedure TvalidDecision.FindLienFournArt(TOBD: TOB; var TheRefArtFour: RRefArtFour);
var QQ : Tquery;
begin
	TheRefArtFour.codeArtFour := '';
	TheRefArtFour.libelle := '';
  QQ := OpenSql ('SELECT GCA_REFERENCE,GCA_LIBELLE,GCA_DELAILIVRAISON FROM CATALOGU WHERE GCA_ARTICLE="'+TOBD.GEtValue('BAD_ARTICLE')+'" AND '+
  							 'GCA_TIERS="'+TOBD.GetValue('BAD_FOURNISSEUR')+'"',true);
  if not QQ.eof then
  begin
  	TheRefArtFour.codeArtFour := QQ.findField('GCA_REFERENCE').AsString;
  	TheRefArtFour.Libelle     := QQ.findField('GCA_LIBELLE').AsString;
  	TheRefArtFour.DelaiFou    := QQ.findField('GCA_DELAILIVRAISON').Asinteger;
  end else
  begin
  	TheRefArtFour.codeArtFour := TOBD.GetValue('BAD_CODEARTICLE');
  	TheRefArtFour.Libelle     := TOBD.GetValue('BAD_LIBELLE');
  	TheRefArtFour.DelaiFou    := 0;
  end;
  ferme (QQ);
end;

procedure TvalidDecision.InitLigne (TOBL,TOBART,TOBFourn : TOB);
var TOBPiece : TOB;
		DEV : Rdevise;
begin
	TOBPiece := TOB.Create ('PIECE',nil,-1);
  TRY
  	TOBPiece.putValue('GP_NATUREPIECEG','CF'); // pour le cas ou...

    TOBL.PutValue('GL_PRIXPOURQTE', 1);
    TOBL.PutValue('GL_CODECOND', '');
    TOBL.PutValue('GL_INDICENOMEN', 0);
    TOBL.PutValue('GL_TYPELIGNE', 'ART');
    TOBL.PutValue('GL_TYPEREF', 'ART');
    TOBL.PutValue('GL_TYPEDIM', 'NOR');
    TOBL.PutValue('GL_NIVEAUIMBRIC', 0);
    //
    TiersVersPiece(TOBFourn, TOBPiece);
    TOBPiece.PutValue('GP_DEVISE',TOBFourn.getValue('T_DEVISE') );
    // recuperation
    DEV.Code := TOBPIECE.GetValue('GP_DEVISE');
    GetInfosDevise(DEV);
    TOBPiece.putValue('GP_TAUXDEV',DEV.Taux);
    AttribCotation (TOBPiece);
    //
    TOBL.PutValue('GL_DEVISE',TOBFourn.getValue('T_DEVISE') );
    if (PositionneExige(TOBFourn)='TE') then TOBL.PutValue('GL_TVAENCAISSEMENT','X')
    																		else TOBL.PutValue('GL_TVAENCAISSEMENT','-');
    TOBL.PutValue('GL_TIERS',TOBFourn.GetValue('T_TIERS'));
    TOBL.PutValue('GL_TIERSFACTURE', TOBPiece.GetValue('GP_TIERSFACTURE'));
    TOBL.PutValue('GL_TIERSLIVRE', TOBPiece.GetValue('GP_TIERSLIVRE'));
    TOBL.PutValue('GL_TIERSPAYEUR', TOBPiece.GetValue('GP_TIERSPAYEUR'));
    TOBL.PutValue('GL_TARIFSPECIAL', TOBPiece.GetValue('GP_TARIFSPECIAL'));
    TOBL.PutValue('GL_TARIFSPECIAL', TOBPiece.GetValue('GP_TARIFSPECIAL'));
    TOBL.PutValue('GL_TARIFTIERS', TOBPiece.GetValue('GP_TARIFTIERS')); //mcd 20/01/03 oubli ...
    TOBL.PutValue('GL_DATEPIECE', TOBPiece.GetValue('GP_DATEPIECE'));
    TOBL.PutValue('GL_NATUREPIECEG', TOBPiece.GetValue('GP_NATUREPIECEG'));
    TOBL.PutValue('GL_ETABLISSEMENT', TOBPiece.GetValue('GP_ETABLISSEMENT'));
    TOBL.PutValue('GL_FACTUREHT', TOBPiece.GetValue('GP_FACTUREHT'));
    TOBL.PutValue('GL_TAUXDEV', TOBPiece.GetValue('GP_TAUXDEV'));
    TOBL.PutValue('GL_COTATION', TOBPiece.GetValue('GP_COTATION'));
    TOBL.PutValue('GL_REGIMETAXE', TOBPiece.GetValue('GP_REGIMETAXE'));
    TOBL.PutValue('GL_REPRESENTANT', TOBPiece.GetValue('GP_REPRESENTANT'));
    TOBL.PutValue('GL_APPORTEUR', TOBPiece.GetValue('GP_APPORTEUR'));
    TOBL.PutValue('GL_ESCOMPTE', TOBPiece.GetValue('GP_ESCOMPTE'));
    TOBL.PutValue('GL_REMISEPIED', TOBPiece.GetValue('GP_REMISEPIED'));
    TOBL.PutValue('GL_SAISIECONTRE', TOBPiece.GetValue('GP_SAISIECONTRE'));
    TOBL.PutValue('GL_DOMAINE', TOBPiece.GetValue('GP_DOMAINE'));
    TOBL.PutValue('GL_VIVANTE','X');
  FINALLY
  	TOBPiece.free;
  END;
end;

procedure TvalidDecision.InitLigneBesoin (TOBL : TOB);
begin
  TOBL.PutValue('GL_PRIXPOURQTE', 1);
  TOBL.PutValue('GL_CODECOND', '');
  TOBL.PutValue('GL_INDICENOMEN', 0);
  TOBL.PutValue('GL_TYPELIGNE', 'ART');
  TOBL.PutValue('GL_TYPEREF', 'ART');
  TOBL.PutValue('GL_TYPEDIM', 'NOR');
  TOBL.PutValue('GL_NIVEAUIMBRIC', 0);
  TOBL.PutValue('GL_VIVANTE','X');
end;

function TvalidDecision.FindArticle(CodeArticle: String): TOB;
var QQ : TQuery;
begin
  result := TOBArticles.findFirst(['GA_ARTICLE'],[CodeArticle],true);
  if result = nil then
  begin
    QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
    							 'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
    							 'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+CodeArticle+'"',true);
    if not QQ.eof then
    begin
    	result := TOB.Create ('ARTICLE',TOBArticles,-1);
      result.SelectDB ('',QQ);
      InitChampsSupArticle (result);
    end;
    ferme(QQ);
  end;
end;

procedure TvalidDecision.DecisionnelVersLigne(TOBD, TOBL: TOB;TheRefArtFourn : RRefArtFour);
var DateL : TdateTime;
		Unite : string;
    MTPAF,MTPAN : double;
    FUS,FUA,FUV,CoefUaUS,CoefUSUV : double;
begin
  unite := TOBD.GetValue('BAD_QUALIFQTEACH');
  // Passage de l'unite de stock en unite d'achat (normal on génère une commande d'achat)
  FUV := RatioMesure('PIE', TOBD.GetValue('BAD_QUALIFQTEVTE'));
  FUS := RatioMesure('PIE', TOBD.GetValue('BAD_QUALIFQTESTO'));
  FUA := RatioMesure('PIE', TOBD.GetValue('BAD_QUALIFQTEACH'));
  CoefUaUs := TOBD.GetValue('BAD_COEFUAUS');
  CoefUSUV := TOBD.GetValue('BAD_COEFUSUV');
  if CoefUSUV = 0 then CoefUSUV := FUS/FUV;

  // si PA à zéro, reprise du PA Budget
 	MTPAF := TOBD.GetValue('BAD_PRIXACH');
  if MTPAF = 0 then MTPAF := TOBD.GetValue('BAD_PABASE');
 	MTPAN := TOBD.GetValue('BAD_PRIXACHNET');
  if MTPAN = 0 then MTPAN := TOBD.GetValue('BAD_PABASE');
  (*
  // voila voila sauf que les prix sont exprimé en UV --> donc passage en UA
  if CoefuaUs <> 0 then
  begin
    if CoefUSUV <> 0 then
    BEGIN
      MTPAF :=  Arrondi(MTPAF * CoefUSUV * CoefuaUs,V_PGI.okdecP);
      MTPAN :=  Arrondi(MTPAN * CoefUSUV * CoefuaUs,V_PGI.okdecP);
    END else
    begin
      MTPAF :=  Arrondi(MTPAF * CoefuaUs / FUV * FUS ,V_PGI.okdecP);
      MTPAN :=  Arrondi(MTPAN * CoefuaUs / FUV * FUS ,V_PGI.okdecP);
    end;
  end else
  begin
    MTPAF :=  Arrondi((MTPAF / FUv ) * FUS * FUA,V_PGI.okdecP);
    MTPAN :=  Arrondi((MTPAN / FUV ) * FUS * FUA,V_PGI.okdecP);
  end;
  *)
  MTPAF := MTPAF * CoefUAUS;
 	MTPAN := MTPAN * CoefUAUS;
  // --
  if Unite = '' then unite := TOBD.GetValue('BAD_QUALIFQTEVTE');
  TOBL.PutValue('GL_QUALIFQTEACH',Unite);
  TOBL.PutValue('GL_CODEARTICLE',TOBD.GetValue('BAD_CODEARTICLE'));
  TOBL.PutValue('GL_ARTICLE',TOBD.GetValue('BAD_ARTICLE'));
  TOBL.PutValue('GL_TYPEARTICLE',TOBD.GetValue('BAD_TYPEARTICLE'));
  TOBL.PutValue('GL_REFARTSAISIE',TheRefArtFourn.codeArtFour );
  TOBL.PutValue('GL_REFARTTIERS',TheRefArtFourn.codeArtFour );
  TOBL.PutValue('GL_BLOCNOTE', TOBD.GetValue('BAD_BLOCNOTE'));
//  if TheRefArtFourn.Libelle <> '' then
//  begin
//  	TOBL.PutValue('GL_LIBELLE',TheRefArtFourn.Libelle );
//  end else
//  begin
  	TOBL.PutValue('GL_LIBELLE',TOBD.GetValue('BAD_LIBELLE'));
//  end;
  TOBL.PutValue('GL_ETABLISSEMENT',TOBD.GetValue('BAD_ETABLISSEMENT'));
  TOBL.PutValue('GL_DEPOT',TOBD.GetValue('BAD_DEPOT'));
  TOBL.PutValue('GL_AFFAIRE',TOBD.GetValue('BAD_AFFAIRE'));
  TOBL.PutValue('GL_AFFAIRE1',TOBD.GetValue('BAD_AFFAIRE1'));
  TOBL.PutValue('GL_AFFAIRE2',TOBD.GetValue('BAD_AFFAIRE2'));
  TOBL.PutValue('GL_AFFAIRE3',TOBD.GetValue('BAD_AFFAIRE3'));
  TOBL.PutValue('GL_AVENANT',TOBD.GetValue('BAD_AVENANT'));
  TOBL.PutValue('GL_PUHTDEV',MTPAF);
  TOBL.PutValue('GL_PUHTNETDEV',MTPAN);
  TOBL.PutValue('GL_REMISELIGNE',TOBD.GetValue('BAD_REMISE'));
  TOBL.PutValue('GL_REMISECASCADE',TOBD.GetValue('BAD_CASCADEREMISE'));
  TOBL.PutValue('GL_TARIF',TOBD.GetValue('BAD_TARIF'));
  TOBL.PutValue('GL_DPA',MTPAN);
  TOBL.PutValue('GL_DPR',MTPAN);
  TOBL.PutValue('GL_QTEFACT',TOBD.GEtValue('BAD_QUANTITEACH'));
  TOBL.PutValue('GL_QTESTOCK',TOBD.GEtValue('BAD_QUANTITEACH'));
  TOBL.PutValue('GL_QTERESTE',TOBD.GEtValue('BAD_QUANTITEACH'));
  if TOBD.GetValue('BAD_LIVCHANTIER')='X' Then TOBL.PutValue('GL_IDENTIFIANTWOL',-1)
  																				else TOBL.PutValue('GL_IDENTIFIANTWOL',0);
  TOBL.PutValue('GL_TENUESTOCK', TOBD.GetValue('BAD_TENUESTOCK'));
  DateL:=TOBD.GetValue('BAD_DATELIVRAISON')-TOBResult.GetValue ('DELAISECURITE')-TheRefArtFourn.DelaiFou;
  if DateL < Date then TOBL.putValue('GL_DATELIVRAISON',Date);
  TOBL.PutValue('GL_DATELIVRAISON', DateL);
  TOBL.PutValue('GL_PIECEPRECEDENTE','') ;
  TOBL.PutValue('GL_PIECEORIGINE',TOBD.getValue('BAD_REFGESCOM')) ;
end;

procedure TvalidDecision.DecisionnelVersLigneBesoin(TOBD, TOBL, TOBA : TOB);
var Unite : string;
		PaPR,PrPv : double;
    Pa,Pr,Pv : double;
    TOBDis : TOB;
    FUS,FUV,RatioStk : double;
begin
	PA := TOBA.getValue('GA_PAHT');
	PR := TOBA.getValue('GA_DPR');
	PV := TOBA.getValue('GA_PVHT');
  if PA <> 0 then PAPR := PR/PA else PAPR := 1;
  if PR <> 0 then PRPV := PV/PR else PRPV := 1;
  unite := TOBD.GetValue('BAD_QUALIFQTEVTE');
  TOBL.PutValue('GL_QUALIFQTEVTE',Unite);
  TOBL.PutValue('GL_CODEARTICLE',TOBD.GetValue('BAD_CODEARTICLE'));
  TOBL.PutValue('GL_ARTICLE',TOBD.GetValue('BAD_ARTICLE'));
  TOBL.PutValue('GL_TYPEARTICLE',TOBD.GetValue('BAD_TYPEARTICLE'));
  TOBL.PutValue('GL_REFARTSAISIE',TOBD.GetValue('BAD_CODEARTICLE'));
  TOBL.PutValue('GL_REFARTTIERS',TOBD.GetValue('BAD_CODEARTICLE') );
  TOBL.PutValue('GL_ETABLISSEMENT',TOBD.GetValue('BAD_ETABLISSEMENT'));
  TOBL.PutValue('GL_DEPOT',TOBD.GetValue('BAD_DEPOT'));
  TOBL.PutValue('GL_AFFAIRE',TOBD.GetValue('BAD_AFFAIRE'));
  TOBL.PutValue('GL_AFFAIRE1',TOBD.GetValue('BAD_AFFAIRE1'));
  TOBL.PutValue('GL_AFFAIRE2',TOBD.GetValue('BAD_AFFAIRE2'));
  TOBL.PutValue('GL_AFFAIRE3',TOBD.GetValue('BAD_AFFAIRE3'));
  TOBL.PutValue('GL_AVENANT',TOBD.GetValue('BAD_AVENANT'));
  if TOBD.GetValue('BAD_FOURNISSEUR') <> '' then TOBL.PutValue('GL_DPA',TOBD.GetValue('BAD_PRIXACHNET'))
  																				  else TOBL.PutValue('GL_DPA',TOBA.GetValue('GA_PAHT')) ;
  TOBL.PutValue('GL_DPR',TOBL.GetValue('GL_DPA')*PAPR);
  TOBL.PutValue('GL_PUHT',Arrondi(TOBL.GetValue('GL_DPR')*PRPV,V_PGI.OkDecP ));
  TOBL.PutValue('GL_PUHTDEV',Arrondi(TOBL.GetValue('GL_DPR')*PRPV,V_PGI.OkDecP ));
  TOBL.PutValue('GL_REMISELIGNE',0);
  TOBL.PutValue('GL_REMISECASCADE','');
  TOBL.PutValue('GL_QTEFACT',TOBD.GEtValue('BAD_QUANTITEVTE'));
  TOBL.PutValue('GL_QTESTOCK',TOBD.GEtValue('BAD_QUANTITEVTE'));
  TOBL.PutValue('GL_QTERESTE',TOBD.GEtValue('BAD_QUANTITEVTE'));
  if TOBD.GetValue('BAD_LIVCHANTIER')='X' Then TOBL.PutValue('GL_IDENTIFIANTWOL',-1)
  																				else TOBL.PutValue('GL_IDENTIFIANTWOL',0);
  TOBL.PutValue('GL_TENUESTOCK', TOBD.GetValue('BAD_TENUESTOCK'));
  TOBL.putValue('GL_DATELIVRAISON',TOBD.GetValue('BAD_DATELIVRAISON'));
  TOBDis := FindDispo (TOBL.GEtVAlue('GL_ARTICLE'),TOBL.GEtVAlue('GL_DEPOT'),TOBL.GEtValue('GL_TENUESTOCK'));
  if TOBDis <> nil Then
  begin
    FUV := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTEVTE')); if FUV = 0 then FUV := 1;
    FUS := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTESTO')); if FUS = 0 then FUS := 1;
    RatioStk := FUV/FUS;
  	TOBDis.PutValue('__QTEPLUS',TOBDis.getValue('__QTEPLUS')+(TOBL.GEtValue('GL_QTEFACT')*RatioStk));
  end;
end;

procedure TvalidDecision.ArticleVersLigne(TOBA, TOBL: TOB);
var i : integer;
begin
	TOBL.PutValue('GL_PRIXPOURQTE', TOBA.GetValue('GA_PRIXPOURQTE'));
  TOBL.PutValue('GL_PCB',TOBA.GetValue('GA_PCB'));
  //
  TOBL.PutValue('GL_LIBELLE', TOBA.GetValue('GA_LIBELLE'));
  TOBL.PutValue('GL_LIBCOMPL', TOBA.GetValue('GA_LIBCOMPL'));
  TOBL.PutValue('GL_REFARTBARRE', TOBA.GetValue('GA_CODEBARRE'));
  TOBL.PutValue('GL_ESCOMPTABLE', TOBA.GetValue('GA_ESCOMPTABLE'));
  TOBL.PutValue('GL_REMISABLEPIED', TOBA.GetValue('GA_REMISEPIED'));
  TOBL.PutValue('GL_REMISABLELIGNE', TOBA.GetValue('GA_REMISELIGNE'));
  TOBL.PutValue('GL_TARIFARTICLE', TOBA.GetValue('GA_TARIFARTICLE'));
  TOBL.PutValue('GL_NUMEROSERIE', TOBA.GetValue('GA_NUMEROSERIE'));
  TOBL.PutValue('GL_PAYSORIGINE', TOBA.GetValue('GA_PAYSORIGINE'));  //JS FI10226
  TOBL.PutValue('GL_CODEDOUANE', TOBA.GetValue('GA_CODEDOUANIER'));
  {Familles, collection, domaine}
  TOBL.PutValue('GL_FAMILLENIV1', TOBA.GetValue('GA_FAMILLENIV1'));
  TOBL.PutValue('GL_FAMILLENIV2', TOBA.GetValue('GA_FAMILLENIV2'));
  TOBL.PutValue('GL_FAMILLENIV3', TOBA.GetValue('GA_FAMILLENIV3'));
  TOBL.PutValue('GL_COLLECTION', TOBA.GetValue('GA_COLLECTION'));
  TOBL.PutValue('GL_DOMAINE', TOBA.GetValue('GA_DOMAINE'));
  for i := 1 to 5 do TOBL.PutValue('GL_FAMILLETAXE' + IntToStr(i), TOBA.GetValue('GA_FAMILLETAXE' + IntToStr(i)));
  TOBL.PutValue('GL_QUALIFSURFACE', TOBA.GetValue('GA_QUALIFSURFACE'));
  TOBL.PutValue('GL_QUALIFVOLUME', TOBA.GetValue('GA_QUALIFVOLUME'));
  TOBL.PutValue('GL_QUALIFPOIDS', TOBA.GetValue('GA_QUALIFPOIDS'));
  TOBL.PutValue('GL_QUALIFLINEAIRE', TOBA.GetValue('GA_QUALIFLINEAIRE'));
  TOBL.PutValue('GL_QUALIFHEURE', TOBA.GetValue('GA_QUALIFHEURE'));
  TOBL.PutValue('GL_SURFACE', TOBA.GetValue('GA_SURFACE'));
  TOBL.PutValue('GL_VOLUME', TOBA.GetValue('GA_VOLUME'));
  TOBL.PutValue('GL_POIDSBRUT', TOBA.GetValue('GA_POIDSBRUT'));
  TOBL.PutValue('GL_POIDSNET', TOBA.GetValue('GA_POIDSNET'));
  TOBL.PutValue('GL_POIDSDOUA', TOBA.GetValue('GA_POIDSDOUA'));
  TOBL.PutValue('GL_LINEAIRE', TOBA.GetValue('GA_LINEAIRE'));
  TOBL.PutValue('GL_HEURE', TOBA.GetValue('GA_HEURE'));
  for i := 1 to 9 do TOBL.PutValue('GL_LIBREART' + IntToStr(i), TOBA.GetValue('GA_LIBREART' + IntToStr(i)));
  TOBL.PutValue('GL_LIBREARTA', TOBA.GetValue('GA_LIBREARTA'));
end;

procedure TvalidDecision.ClotureDecisionnel;
begin
  if ExecuteSQL ('UPDATE DECISIONACH SET BAE_VALIDE="X" WHERE BAE_NUMERO='+inttoStr(fNumDecisionel)) < 0 then V_PGI.IOError := oeUnknown;
end;

procedure TvalidDecision.RemplacementArticle(TOBD, TOBDecisionnel,TOBPiece,TOBBase,TOBeche,TOBTier: TOB);
var TOBR,TOBREMPL : TOB;
		IndiceInsert : integer;
begin
	TOBR := TOBDecisionnel.findFirst(['BAD_REMPLACEANT','BAD_NUMREMPLACE'],['X',TOBD.getValue('BAD_NUMREMPLACE')],true);
  repeat
    if (TOBR <> nil) and (TOBR <> TOBD) then break;
    if TOBR = nil then TOBR := TOBDecisionnel.findNext(['BAD_REMPLACEANT','BAD_NUMREPLACE'],['X',TOBD.getValue('BAD_NUMREMPLACE')],true);
  until TOBR = nil;
  IndiceInsert:= AnnuleArticleDansBesoin (TOBD,TOBPiece,TOBREMPL);
  AjouteArticleDansbesoin (TOBR,TOBPiece,TOBBase,TOBTier,TOBREMPL,IndiceInsert);
end;

function TvalidDecision.isRemplacee(TOBL: TOB): boolean;
begin
  result := (TOBL.Getvalue('BAD_NUMREMPLACE')<>0) and (TOBL.Getvalue('BAD_REMPLACE')='X');
end;


function TvalidDecision.isRemplaceant(TOBL: TOB): boolean;
begin
  result := (TOBL.Getvalue('BAD_NUMREMPLACE')<>0) and (TOBL.Getvalue('BAD_REMPLACEANT')='X');
end;

procedure TvalidDecision.findTOBs (reference: string;var  TOBPiece,TOBbase,TOBeche: TOB);
var cledoc : R_Cledoc;
		TheRef : string;
begin
 	DecodeRefPiece (reference,cledoc);
  TheRef := cledoc.NaturePiece+';'+cledoc.Souche+';'+inttostr(cledoc.NumeroPiece)+';'+inttostr(cledoc.Indice);
	TOBPiece := TOBpieces.findFirst(['REFERENCE'],[TheRef],true);
  TOBBase := TOBBases.findFirst(['REFERENCE'],[TheRef],true);
  TOBeche := TOBeches.findFirst(['REFERENCE'],[TheRef],true);
end;

procedure TvalidDecision.CreeEntree(reference : string;var TOBpiece,TObbase,TOBeche : TOB);
var cledoc : R_Cledoc;
begin
	DecodeRefPiece (reference,cledoc);
  Reference := cledoc.NaturePiece+';'+cledoc.Souche+';'+inttostr(cledoc.NumeroPiece)+';'+inttostr(cledoc.Indice);
	TOBPiece := TOB.Create ('PIECE',TOBPieces,-1);
  AddLesSupEntete(tobpiece);
  TOBPiece.AddChampSupValeur ('REFERENCE',reference);
  TOBBase := TOB.create ('BASES',Tobbases,-1);
  TOBBase.AddChampSupValeur ('REFERENCE',reference);
  TOBeche := TOB.create ('BASES',TOBEches,-1);
  TOBeche.AddChampSupValeur ('REFERENCE',reference);
end;

procedure TvalidDecision.ChargeBesoin (reference : string;var TOBPiece,TOBbase,TOBeche,TOBTier : TOB);
var cledoc : R_CLEDOC;
begin
	DecodeRefPiece (reference,cledoc);
  findTOBs (reference,TOBPiece,TOBbase,TOBeche);
  if TOBPiece = nil then
  begin
  	CreeEntree(reference,TOBpiece,TObbase,TOBeche);
  	LoadLesTOB (cledoc,TOBPiece,TOBBase,TOBEche);
  	chargeAjouteArticles (Cledoc);
    ChargeAjouteDispo (Cledoc);
  end;
  TOBTier := FindTiers (TOBPiece.getValue('GP_TIERS'),'CLI');
  if (TOBTier = nil)  then
  BEGIN
    TOBTier := TOB.Create ('TIERS',TOBTiers,-1);
    TOBTier.AddChampSupValeur('RIB','', False);
    RemplirTOBTiers(TOBTier, TOBPiece.getValue('GP_TIERS'), TOBPIece.getValue('GP_NATUREPIECEG'), True);
  END;
  DEVBesoin.Code := TOBPIECE.GetValue('GP_DEVISE');
  DEVBesoin.Cotation := TOBPiece.GetValue('GP_COTATION');
  DEVBesoin.Taux := TOBPiece.GetValue('GP_TAUXDEV');
end;

function TvalidDecision.AddDispo (TOBD:TOB)  : TOB;
begin
	result := TOB.Create (' THE DISPO',TOBDISPO,-1);
  result.AddChampSupValeur ('__ARTICLE',TOBD.GEtValue('GQ_ARTICLE'));
  result.AddChampSupValeur ('__DEPOT',TOBD.GEtValue('GQ_DEPOT'));
  result.AddChampSupValeur ('__QTEPLUS',0);
  result.AddChampSupValeur ('__QTEMOINS',0);
end;

Function TvalidDecision.FindDispo (Article,Depot,GereStock : string) : TOB;
begin
	result := nil;
  if gereStock<>'X' then exit;
	result := TOBDispo.findFirst(['__ARTICLE','__DEPOT'],[Article,Depot],true);
  if result = nil then
  begin
  	result := AddDispo (Article,Depot);
  end;
end;


procedure TvalidDecision.ChargeAjouteDispo (Cledoc : r_cledoc);
var Q :Tquery;
		TOBDisp,TOBD : TOB;
    Indice : integer;
begin
	TOBDisp := TOB.Create('LES DISPO',nil,-1);
  TRY
    Q := OpenSQL('SELECT * FROM DISPO WHERE GQ_ARTICLE IN (SELECT DISTINCT GL_ARTICLE FROM LIGNE WHERE '+WherePiece(cledoc,ttdligne,false)+') AND '+
    						 'GQ_DEPOT IN (SELECT DISTINCT GL_DEPOT FROM LIGNE WHERE '+WherePiece(cledoc,ttdligne,false)+')'	, True);
    TOBDisp.LoadDetailDB ('DISPO','','',Q,false);
    ferme(Q);
    for indice := 0 to TOBDisp.detail.count -1 do
    begin
      TOBD := TOBDispo.findFirst(['__ARTICLE','__DEPOT'],
      													 [TOBDisp.detail[Indice].getValue('GQ_ARTICLE'),TOBDisp.detail[Indice].getValue('GQ_DEPOT')],true);
      AddDispo (TOBDisp.detail[Indice])
    end;
  FINALLY
  	TOBDisp.free;
  END;
end;

procedure TvalidDecision.chargeAjouteArticles (Cledoc : r_cledoc);
var Q :Tquery;
		TOBArticle,TOBA : TOB;
    Indice : integer;
begin
	TOBArticle := TOB.Create ('LES ARTICLES',nil,-1);
  TRY
    Q := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
    							 'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
    							 'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE IN (SELECT DISTINCT GL_ARTICLE FROM LIGNE WHERE '+
                   WherePiece(cledoc,ttdligne,false)+')',true);
    TOBArticle.LoadDetailDB ('ARTICLE','','',Q,false);
    ferme(Q);
    if TOBArticle.detail.count > 0 then
    begin
      Indice := 0;
      repeat
        TOBA := TOBArticles.findFirst(['GA_ARTICLE'],[TOBArticle.detail[Indice].getValue('GA_ARTICLE')],true);
        if TOBA = nil then
        begin
      	  InitChampsSupArticle (TOBArticle.detail[Indice]);
      	  TOBArticle.detail[Indice].ChangeParent (TOBArticles,-1) // si pas existant dans la liste des article on ajoute
        end else Inc(Indice);
      until indice >= TOBArticle.detail.count;
    end;
  FINALLY
  	TOBArticle.free;
  END;
end;

function TvalidDecision.AnnuleArticleDansBesoin (TOBD,TOBPiece: TOB ;var TOBREMPL : TOB) : integer;
var TOBL,TOBDis : TOB;
		cledoc : R_CLEDOC;
    FUV,FUS,RatioStk : double;
begin
	result := -1;
	DecodeRefPiece (TOBD.GetVAlue('BAD_REFGESCOM'),cledoc);
  TOBL := TOBPiece.findFirst(['GL_NUMORDRE'],[cledoc.NumOrdre],true);
  if TOBL <> nil then
  begin
    TOBDis := FindDispo (TOBL.GetValue('GL_ARTICLE'),TOBL.GetValue('GL_DEPOT'),TOBL.GEtValue('__TENUESTOCK'));
    if TOBDis <> nil then
    begin
      FUV := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTEVTE')); if FUV = 0 then FUV := 1;
      FUS := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTESTO')); if FUS = 0 then FUS := 1;
      RatioStk := FUV/FUS;
    	TOBDis.PutValue('__QTEMOINS',TOBDis.GEtVAlue('__QTEMOINS')+(TOBL.GetValue('GL_QTERESTE')*RatioStk));
    end;
  	TOBL.putValue('GL_QTERESTE',0);
    TOBREMPL := TOBL;
    result := TOBL.GetValue('GL_NUMLIGNE');
  end;
end;

procedure TvalidDecision.AjouteArticleDansbesoin (TOBR,TOBPiece,TOBBase,TOBTier,TOBREMPL : TOB; IndiceInsert : integer);
var TOBA,TOBL : TOB;
		Article : string;
    reference : string;
begin
	if TOBR = nil then exit;
	Article := TOBR.GEtVAlue('BAD_ARTICLE');
	TOBA := FindArticle (Article);
  TOBL := TOB.Create ('LIGNE',TOBPiece,IndiceInsert);
  AddLesSupLigne (TOBL,false);
  InitLigneBesoin (TOBL);
	PieceVersLigne (TOBPiece,TOBL);
  ArticleVersLigne (TOBA,TOBL);
  DecisionnelVersLigneBesoin (TOBR,TOBL,TOBA);
  TOBR.putValue('BLP_PHASETRA',TOBREMPL.getValue('BLP_PHASETRA')); // on le remet dans la meme phase de travaux
  NumeroteLignesGC(Nil, TOBPiece);
  PutValueDetail (TOBpiece,'GP_RECALCULER','X');
  CalculFacture (nil,TOBPiece,nil,nil,nil,nil,TOBBase,TOBBasesL,TOBTier,TOBArticles,nil,nil,nil,nil,DEVbesoin,false);
  reference := EncodeRefPiece (TOBL);
  TOBR.putValue('BAD_REFGESCOM',reference);
end;

procedure TvalidDecision.PositionneAncienneValeurs (TOBPiece : TOB);
var Indice : integer;
		TOBL : TOB;
begin
	for Indice := 0 to TOBPiece.detail.count -1 do
  begin
    TOBL := TOBPIece.detail[Indice];
    TOBL.AddChampSupValeur ('__TENUESTOCK',TOBL.GEtVAlue('GL_TENUESTOCK'));
  end;
end;

procedure TvalidDecision.LoadLesTOB (cledoc : r_cledoc; TOBPiece,TOBBase,TOBEche : TOB);
var Q: TQuery;
begin
  LoadPieceLignes(CleDoc, TobPiece);
  PieceAjouteSousDetail(TOBPiece);
  PositionneAncienneValeurs (TOBPiece);
  // Lecture bases
  Q := OpenSQL('SELECT * FROM PIEDBASE WHERE ' + WherePiece(CleDoc, ttdPiedBase, False), True);
  TOBBase.LoadDetailDB('PIEDBASE', '', '', Q, False);
  Ferme(Q);
  TOBBasesL.clearDetail;
  // Lecture bases Lignes
  Q := OpenSQL('SELECT * FROM LIGNEBASE WHERE ' + WherePiece(CleDoc, ttdLigneBase, False), True,-1, '', True);
  TOBBasesL.LoadDetailDB('LIGNEBASE', '', '', Q, False);
  Ferme(Q);
  OrdonnelignesBases (TOBBasesL);
  // Lecture Echéances
  Q := OpenSQL('SELECT * FROM PIEDECHE WHERE ' + WherePiece(CleDoc, ttdEche, False), True);
  TOBEche.LoadDetailDB('PIEDECHE', '', '', Q, False);
  Ferme(Q);
end;

procedure TvalidDecision.ReajusteReserve;
var Indice : integer;
		TOBD : TOB;
    Maj : double;
begin
	for Indice := 0 to TOBDispo.detail.count -1 do
  begin
  	TOBD := TOBDispo.detail[Indice];
    if (TOBD.GetValue('__QTEPLUS')<> 0) or (TOBD.GetValue('__QTEMOINS')<> 0) then
    begin
      Maj := Arrondi(TOBD.GetValue('__QTEPLUS') - TOBD.GetValue('__QTEMOINS'),V_PGI.OkDecQ);
      if Maj <> 0 then
      begin
        if Maj < 0 then
        begin
        	Maj := Maj *-1;
          if ExecuteSql ('UPDATE DISPO SET GQ_RESERVECLI=GQ_RESERVECLI - '+StrfPoint(Maj)+' '+
          							'WHERE GQ_ARTICLE="'+TOBD.GEtValue('__ARTICLE') +'" '+
                        'AND GQ_DEPOT="'+TOBD.GetValue('__DEPOT')+'" ') < 0 then
          begin
          	V_PGI.IOError := OeUnknown;
          end;
        end else
        begin
          if ExecuteSql ('UPDATE DISPO SET GQ_RESERVECLI=GQ_RESERVECLI + '+StrfPoint(Maj)+' '+
          							'WHERE GQ_ARTICLE="'+TOBD.GEtValue('__ARTICLE') +'" '+
                        'AND GQ_DEPOT="'+TOBD.GetValue('__DEPOT')+'" ') < 0 then
          begin
          	V_PGI.IOError := OeUnknown;
          end;
        end;
      end;
    end;
  end;
end;

function TvalidDecision.EnregistreLesBesoins : boolean;
var Indice : integer;
begin
	result := true;
	for Indice := 0 to TOBPieces.detail.count -1 do
  begin
  	if not supprimePiece (TOBPieces.detail[Indice]) then BEGIN result := false; exit; END;
    TOBPieces.detail[Indice].SetAllModifie (true);
    VireLesEltNonNecessaires(TOBPieces.detail[Indice]);
  end;
	for Indice := 0 to TOBbases.detail.count -1 do
  begin
    TOBbases.detail[Indice].SetAllModifie (true);
  end;
	for Indice := 0 to TOBEches.detail.count -1 do
  begin
    TOBEches.detail[Indice].SetAllModifie (true);
  end;
//
	for Indice := 0 to TOBPieces.detail.count -1 do
  begin
		if (V_PGI.Ioerror = OeOk) and (not TOBPieces.detail[Indice].InsertDBByNivel (true)) then
    begin
    	V_PGI.ioerror := OeUnknown;
      result := false;
      break;
    end;
  end;
	if (V_PGI.Ioerror = OeOk) and (not TOBbases.InsertDB (nil,true)) then begin V_PGI.ioerror := OeUnknown; result := false; end;
	if (V_PGI.Ioerror = OeOk) and (not TOBEches.InsertDB (nil,true)) then begin V_PGI.ioerror := OeUnknown; result := false; end;
end;

function TvalidDecision.supprimePiece (TOBPiece: TOB) : boolean;
var Cledoc : r_cledoc;
		reference,sql : string ;
begin
	result := true;
	reference := EncodeRefPiece (TOBPiece);
  DecodeRefPiece (reference,cledoc);
  Sql := 'DELETE FROM PIECE WHERE '+WherePiece(CleDoc, ttdPiece, true,true);
  if ExecuteSQL (SQl) < 0  then BEGIN result := false; exit; end;
  Sql := 'DELETE FROM LIGNE WHERE '+WherePiece(CleDoc, ttdLigne,false);
  if ExecuteSQL (SQl) < 0  then BEGIN result := false; exit; end;
  Sql := 'DELETE FROM PIEDBASE WHERE '+WherePiece(CleDoc, ttdPiedBase, true,true);
  if ExecuteSQL (SQl) < 0  then BEGIN result := false; exit; end;
  Sql := 'DELETE FROM PIEDECHE WHERE '+WherePiece(CleDoc, ttdEche, true,true);
  if ExecuteSQL (SQl) < 0  then BEGIN result := false; exit; end;
end;

procedure TvalidDecision.MajLigneBesoin (TOBR,TOBPiece, TOBBase,TOBeche, TOBTier: TOB ) ;
var cledoc : r_cledoc;
		TOBL,TOBD  : TOB;
    FUV,FUS,RatioStk : double;
begin
	DecodeRefPiece (TOBR.GetValue('BAD_REFGESCOM'),cledoc);
  if TOBPiece = nil then exit;
	TOBL := TOBPiece.findFirst (['GL_NUMORDRE'],[cledoc.NumOrdre],true);
  if TOBL <> nil then
  begin
  	if TOBR.getValue('BAD_LIVCHANTIER')='X' then TOBL.PutValue('GL_IDENTIFIANTWOL',-1)
    																				else TOBL.PutValue('GL_IDENTIFIANTWOL',0);
    TOBL.putValue('GL_TENUESTOCK',TOBR.getValue('BAD_TENUESTOCK'));
    TOBL.putValue('GL_DATELIVRAISON',TOBR.getValue('BAD_DATELIVRAISON'));
    TOBL.putValue('GL_DEPOT',TOBR.getValue('BAD_DEPOT'));
    //
    FUV := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTEVTE')); if FUV = 0 then FUV := 1;
    FUS := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTESTO')); if FUS = 0 then FUS := 1;
    RatioStk := FUV/FUS;
    if TOBL.GetDouble('GL_COEFCONVQTEVTE')<>0 then RatioStk := TOBL.GetDouble('GL_COEFCONVQTEVTE'); 

    if (TOBL.GEtValue('__TENUESTOCK') = 'X')  then
    begin
      TOBD := FindDispo (TOBL.GEtVAlue('GL_ARTICLE'),TOBL.GEtVAlue('GL_DEPOT'),TOBL.GEtValue('__TENUESTOCK'));
      if TOBD <> nil Then
      begin
        TOBD.PutValue('__QTEMOINS',TOBD.getValue('__QTEMOINS')+(TOBL.GEtValue('GL_QTERESTE')*RatioStk));
      end;
		end;
    // Maj du QTEFACT si le reste devient supérieur a la quantite initiale
    if TOBR.GEtValue('BAD_QUANTITEVTE') > TOBL.GetValue('GL_QTEFACT') then
    begin
      TOBL.PutValue('GL_QTEFACT',TOBR.GEtValue('BAD_QUANTITEVTE'));
      TOBL.PutValue('GL_QTESTOCK',TOBR.GEtValue('BAD_QUANTITEVTE'));
    	TOBL.PutValue('GL_QTERESTE',TOBR.GEtValue('BAD_QUANTITEVTE'));
    end;

    if (TOBL.GEtValue('GL_TENUESTOCK') = 'X') then
    begin
      TOBD := FindDispo (TOBL.GEtVAlue('GL_ARTICLE'),TOBL.GEtVAlue('GL_DEPOT'),TOBL.GEtValue('GL_TENUESTOCK'));
      if TOBD <> nil Then
      begin
        TOBD.PutValue('__QTEPLUS',TOBD.getValue('__QTEPLUS')+(TOBL.GEtValue('GL_QTERESTE')*RatioStk));
      end;
    end;

  end;
end;

function TvalidDecision.IsScindee (TOBR : TOB) : boolean;
begin
	Result := (TOBR.getValue('BAD_SCINDEE')='X');
end;

procedure TvalidDecision.GereScindeBesoin(TOBD, TOBPiece, TOBBase, TOBeche,TOBTier: TOB);

	procedure InitFindLigne (TOBpiece : TOB);
  var Indice : integer;
  begin
  	for Indice := 0 TO TOBPiece.detail.count -1 do
    begin
    	TOBPiece.detail[Indice].putValue('GL_NUMLIGNE',0);
    end;
  end;

  function FindLignebesoin (TOBPiece : TOB;Numordre : integer) : integer;
  var Indice : integer;
  begin
  	result := -1;
  	for Indice := 0 TO TOBPiece.detail.count -1 do
    begin
    	if TOBPiece.detail[Indice].GetValue('GL_NUMORDRE') = Numordre then
      begin
      	result := indice;
        break;
      end;
    end;
  end;

var TOBI,NewTOBL,TOBDis : TOB;
		cledoc :r_cledoc;
    reference : string;
    Indice : integer;
    FUV,FUS,RatioStk,Delta : double;
begin
	newTOBl := nil;
  // c'est la ligne rajouté il faut donc recuperer la ligne de base
  DecodeRefPiece (TOBD.GetValue('BAD_REFGESCOM'),cledoc);
  if TOBPiece = nil then exit;
  InitFindLigne (TOBPiece);
  Indice := FindLignebesoin (TOBPiece,Cledoc.numordre);
  TOBI := TOBPiece.detail[Indice];
  FUV := RatioMesure('PIE', TobI.GetValue('GL_QUALIFQTEVTE')); if FUV = 0 then FUV := 1;
  FUS := RatioMesure('PIE', TobI.GetValue('GL_QUALIFQTESTO')); if FUS = 0 then FUS := 1;
  RatioStk := FUV/FUS;
	if TOBD.GetValue('BAD_ADDLIGNE')='X' then
  begin
    if Indice <> -1 then
    begin
      NewTOBL := TOB.create ('LIGNE',TOBPiece,Indice);
      AddLesSupLigne (NewTobl,false);
      NewTOBL.Dupliquer (TOBI,false,true);
      NewTOBL.putvalue('GL_NUMORDRE',0);
      NewTOBL.putvalue('GL_NUMLIGNE',0);
      NewTOBL.PutValue('GL_QTERESTE',TOBD.GEtValue('BAD_QUANTITEVTE'));
      NewTOBL.PutValue('GL_QTEFACT',TOBD.GEtValue('BAD_QUANTITEVTE'));
      NewTOBL.PutValue('GL_QTESTOCK',TOBD.GEtValue('BAD_QUANTITEVTE'));
      NewTOBL.PutValue('GL_DEPOT',TOBD.GEtValue('BAD_DEPOT'));
  		newTOBL.PutValue('GL_RECALCULER','X');
    	NumeroteLignesGC(Nil, TOBPiece);
      reference := EncodeRefPiece (newTOBL);
      TOBD.putValue('BAD_REFGESCOM',reference);
      TOBDis := FindDispo (TOBI.GEtVAlue('GL_ARTICLE'),TOBI.GEtVAlue('GL_DEPOT'),TOBD.GetValue('BAD_TENUESTOCK'));
      if TOBDis <> nil Then
      begin
        TOBDis.PutValue('__QTEPLUS',TOBDis.getValue('__QTEPLUS')+TOBD.GEtValue('BAD_QUANTITESTK'));
      end;
  		NewTOBL.putValue('GL_TENUESTOCK','-');
    end;
  end else
  begin
    TOBDis := FindDispo (TOBI.GEtVAlue('GL_ARTICLE'),TOBI.GEtVAlue('GL_DEPOT'),TOBI.GEtValue('__TENUESTOCK'));
    if TOBDis <> nil Then
    begin
    	// on deduit l'ancien
      TOBDis.PutValue('__QTEMOINS',TOBDis.getValue('__QTEMOINS')+(TOBI.GEtValue('GL_QTERESTE')*RatioStk));
    end;
    DelTa :=TOBI.GetValue('GL_QTERESTE') - TOBD.GEtValue('BAD_QUANTITEVTE');
  	TOBI.PutValue('GL_QTERESTE',TOBD.GEtValue('BAD_QUANTITEVTE'));
    if Delta <> 0 then
    begin
    	// il faut Réajuster le besoin initial
      TOBI.putValue('GL_QTEFACT',TOBI.GetValue('GL_QTEFACT')-Delta);
      TOBI.putValue('GL_QTESTOCK',TOBI.GetValue('GL_QTESTOCK')-Delta);
    end;
    TOBI.PutValue('GL_RECALCULER','X');
    TOBDis := FindDispo (TOBI.GEtVAlue('GL_ARTICLE'),TOBI.GEtVAlue('GL_DEPOT'),TOBD.GEtValue('BAD_TENUESTOCK'));
    if TOBDis <> nil Then
    begin
    	// on Ajoute le nouveau
      TOBDis.PutValue('__QTEPLUS',TOBDis.getValue('__QTEPLUS')+(TOBI.GEtValue('GL_QTERESTE')*RatioStk));
    end;
  	TOBI.putValue('GL_TENUESTOCK','X');
  end;
  if TOBD.GetValue('BAD_LIVCHANTIER')='X' Then TOBI.PutValue('GL_IDENTIFIANTWOL',-1)
                                          else TOBI.PutValue('GL_IDENTIFIANTWOL',0);
  TOBI.putValue('GL_DATELIVRAISON',TOBD.GetValue('BAD_DATELIVRAISON'));
  CalculFacture (nil,TOBPiece,nil,nil,nil,nil,TOBBase,TOBBasesL,TOBTier,TOBArticles,nil,nil,nil,nil,DEVbesoin,false);
end;

function TvalidDecision.FindTiers(CodeTiers,TypeTier: String): TOB;
begin
  result := TOBTiers.findFirst(['T_TIERS','T_NATUREAUXI'],[CodeTiers,TypeTier],true);
end;

procedure TvalidDecision.VireLesEltNonNecessaires(TOBPiece : TOB);
var Indice : integer;
		TOBL : TOB;
begin
	for Indice := 0 to TOBPiece.detail.count -1 do
  begin
    TOBL := TOBPiece.detail[Indice];
    TOBL.ClearDetail;
  end;
end;

function TvalidDecision.AddDispo(Article, Depot: string): TOB;
begin
	result := TOB.Create (' THE DISPO',TOBDISPO,-1);
  result.AddChampSupValeur ('__ARTICLE',Article);
  result.AddChampSupValeur ('__DEPOT',Depot);
  result.AddChampSupValeur ('__QTEPLUS',0);
  result.AddChampSupValeur ('__QTEMOINS',0);
end;

end.

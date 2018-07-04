unit BTUtilDecisionAch;

interface
uses
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ENDIF}
		 UTOB,HCtrls,EntGc,UtilSuggestion,HmsgBox,HRichOle,BtpUtil,BtPrepaLivr,
     Forms,SysUtils,UTofListeInv,Factcomm,HEnt1,DECISIONACHAT_TOF,Vierge
;
type
	TConstructDecisionAch = class
  private
  	 fecran : TForm;
     fDateDeb,fDateFin : TDateTime;
     fArticleDeb,fArticleFin : string;
     fDepotDeb,fDepotFin : string;
     fEtablissement : String;
     fWithSelectDoc : boolean;
     TheNumero : integer;
     frecap : THRichEditOle;
     fTOBEnt,fTOBDetail : TOB; // TOB entete et detail de decisionnel achat
     fTOBSelect : TOB;  // TOB des entetes de documents selectionnes sur les criteres
     fTOBLigneBrut : TOB;
     fRegroupLibellediff : boolean;
     function AjouteFiltreDoc (select : string) : string;
     procedure AjouteTOBLigne (TOBL : TOB);
     function SelectDocATraiter : boolean;
     procedure Initligne (LaTOBCourante,TOBL : TOB; Niveau : integer);
     function ConstitueRequeteEnteteDoc : string;
     function ConstitueRequeteLigneDoc : string;
     function DejaCommande (TOBL : TOB) : double;
     function MtDejaCommande(TOBL: TOB): double;
     function EcritDocument : boolean;
     procedure EcritLesElements;
     procedure GenereCumul;
     function GetNumeroPiece (var TheNumero : integer) : boolean;
     function RecupereLigneDetail : boolean;
     function SetCumul (TOBPere,TOBDetail : TOB) : boolean;
     function SetInfosEntete : boolean;

     function TraiteLigneDetail : boolean;
  public
     constructor create;
     destructor destroy; override;
     property DateDebut : TDateTime read fDateDeb write fDateDeb;
     property DateFin : TDateTime read fDateFin write fDateFin;
     property ArticleDebut : String read fArticleDeb write fArticleDeb;
     property ArticleFin : String read fDepotFin write fDepotFin;
     property DepotDebut : String read fDepotDeb write fDepotDeb;
     property DepotFin : String read fArticleFin write fArticleFin;
     property MessageRecap : THRichEditOle read fRecap write frecap;
     property Etablissement : String read fEtablissement write fEtablissement;
     property WithSelection : boolean read fWithSelectDoc write fWithSelectDoc;
     property ecran : TForm read fecran write fecran;
     Function Generate : boolean;
     procedure LanceFicheSaisie;
  end;

procedure NumeroteLigne (FF : Tform; TOBdecisionnel : TOB);
procedure AppliqueNumerotationLigne(TOBDetail: TOB; Numero: integer);

implementation
uses UtilArticle,Paramsoc;
{ TConstructDecisionAch }

function TConstructDecisionAch.AjouteFiltreDoc(select: string): string;
var locRequete : string;
		i_ind1,NumeroPiece : integer;
    Naturepiece : string;
begin
	LocRequete := '';
  frecap.Lines.add ('');
  if fTOBSelect.Detail.Count > 1 then
  begin
  	frecap.Lines.add ('Document');
    LocRequete := ' AND GL_NUMERO IN (';
    for i_ind1 := 0 to fTOBSelect.Detail.Count - 1 do
    begin
    	Naturepiece := fTOBSelect.Detail[i_ind1].GetValue('GP_NATUREPIECEG');
      NumeroPiece := fTOBSelect.Detail[i_ind1].GetValue('GP_NUMERO');
      if i_ind1 = 0 then
        LocRequete := LocRequete + IntToStr(NumeroPiece)
      else
        LocRequete := LocRequete + ', ' + IntToStr(NumeroPiece);
  		frecap.Lines.add (rechDom('TTNATUREPIECE',Naturepiece,false)+' '+InttoStr(NumeroPiece));
    end;
    locRequete := locRequete + ') ';
  end else if fTOBSelect.Detail.Count = 1 then
  begin
    Naturepiece := fTOBSelect.Detail[0].GetValue('GP_NATUREPIECEG');
    NumeroPiece := fTOBSelect.Detail[0].GetValue('GP_NUMERO');
    frecap.Lines.add ('Document : '+rechDom('TTNATUREPIECE',Naturepiece,false)+' '+InttoStr(NumeroPiece));
    LocRequete := ' And GL_NUMERO=' + IntToStr(NumeroPiece) + ' '
  end;
  result := Select + Locrequete;
end;

procedure TConstructDecisionAch.AjouteTOBLigne(TOBL: TOB);
var LaTOBN1,LaTOBN2,LaTOBN3,LaTobCourrante : TOB;
		LivChantier : String;
begin
	if (TOBL.GetValue('GL_IDENTIFIANTWOL')=-1) then LivChantier := 'X' else LivChantier := '-';
  if fRegroupLibellediff then
  begin
    LaTOBN1 := fTOBDetail.findFirst(['BAD_TYPEL','BAD_ARTICLE'],['NI1',TOBL.GetValue('GL_ARTICLE')],true);
  end else
  begin
    LaTOBN1 := fTOBDetail.findFirst(['BAD_TYPEL','BAD_ARTICLE','BAD_LIBELLE'],['NI1',TOBL.GetValue('GL_ARTICLE'),TOBL.GetValue('GL_LIBELLE')],true);
  end;
  if LaTOBN1 = nil then
  begin
  	// y'en a po ..alors on la cree
    LaTOBN1 := TOB.Create('DECISIONACHLIG',fTOBdetail,-1);
    LaTOBN1.AddChampSupValeur ('MULTI','-');
    LaTOBN1.AddChampSupValeur ('MULTIPABASE','-');
    LaTOBN1.PutValue('BAD_DATELIVRAISON',iDate1900);
    Initligne (LaTOBN1,TOBL,1);
  end;
  if fRegroupLibellediff then
  begin
    LaTOBN2 := LaTOBN1.findFirst(['BAD_TYPEL','BAD_ARTICLE','BAD_LIVCHANTIER'],['NI2',TOBL.GetValue('GL_ARTICLE'),LivChantier],True);
  end else
  begin
    LaTOBN2 := LaTOBN1.findFirst(['BAD_TYPEL','BAD_ARTICLE','BAD_LIBELLE','BAD_LIVCHANTIER'],['NI2',TOBL.GetValue('GL_ARTICLE'),TOBL.GetValue('GL_LIBELLE'),LivChantier],True);
  end;
  if LaTOBN2 = nil then
  begin
  	// y'en a po ..alors on la cree
    LaTOBN2 := TOB.Create('DECISIONACHLIG',LaTOBN1,-1);
    LaTOBN2.AddChampSupValeur ('MULTI','-');
    LaTOBN2.AddChampSupValeur ('MULTIPABASE','-');
    Initligne (LaTOBN2,TOBL,2);
  end;
  if fRegroupLibellediff then
  begin
    LaTOBN3 := LaTOBN2.findFirst(['BAD_TYPEL','BAD_ARTICLE','BAD_LIVCHANTIER','BAD_DEPOT'],['NI3',TOBL.GetValue('GL_ARTICLE'),LivChantier,TOBL.GetValue('GL_DEPOT')],True);
  end else
  begin
    LaTOBN3 := LaTOBN2.findFirst(['BAD_TYPEL','BAD_ARTICLE','BAD_LIBELLE','BAD_LIVCHANTIER','BAD_DEPOT'],['NI3',TOBL.GetValue('GL_ARTICLE'),TOBL.GetValue('GL_LIBELLE'),LivChantier,TOBL.GetValue('GL_DEPOT')],True);
  end;
  if LaTOBN3 = nil then
  begin
  	// y'en a po ..alors on la cree
    LaTOBN3 := TOB.Create('DECISIONACHLIG',LaTOBN2,-1);
    LaTOBN3.AddChampSupValeur ('MULTI','-');
    LaTOBN3.AddChampSupValeur ('MULTIPABASE','-');
    Initligne (LaTOBN3,TOBL,3);
  end;

  LaTobCourrante := TOB.Create('DECISIONACHLIG',LaTOBN3,-1);
  LaTobCourrante.AddChampSupValeur ('MULTI','-');
  LaTobCourrante.AddChampSupValeur ('MULTIPABASE','-');

  Initligne (LaTobCourrante,TOBL,4);
  
end;

procedure AppliqueNumerotationLigne(TOBDetail : TOB; Numero: integer);
var Indice : integer;
		TOBD : TOB;
begin
  for Indice := 0 To TOBDetail.detail.count -1 do
  begin
    TOBD := TOBdetail.detail[Indice];
    TOBD.PutValue ('BAD_NUMERO',Numero);
    if TOBD.Detail.count > 0 then
    begin
      AppliqueNumerotationLigne (TOBD , Numero);
    end;
  end;
end;

function TConstructDecisionAch.ConstitueRequeteEnteteDoc: string;
begin
  result := 'SELECT DISTINCT GP_DATEPIECE,GP_AFFAIRE,GP_TOTALHT,GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO,GP_INDICEG,GP_VIVANTE,GP_INDICEG,GP_TIERS,GP_DEVISE,T_LIBELLE FROM PIECE '+
            'left join LIGNE on (GL_NATUREPIECEG=GP_NATUREPIECEG) and (GL_SOUCHE=GP_SOUCHE) AND (GL_NUMERO=GP_NUMERO) and (GL_INDICEG=GP_INDICEG) ' +
            'left join TIERS on (T_TIERS=GP_TIERS)  ' +
  					'WHERE GP_NATUREPIECEG IN ("'+GetNaturePieceCde +'") AND '+
  					'GL_DATELIVRAISON >="'+USDATETIME(fDateDeb)+'" AND GL_DATELIVRAISON <="'+USDATETIME(fDateFin)+'"';
  if fArticleDeb <> '' then result := result + ' AND GL_CODEARTICLE >="'+fArticleDeb+'"';
  if fArticleFin <> '' then result := result + ' AND GL_CODEARTICLE <="'+fArticleFin+'"';
  if fDepotDeb <> '' then result := result + ' AND GL_DEPOT >="'+fDepotDeb+'"';
  if fDepotFin <> '' then result := result + ' AND GL_DEPOT <="'+fDepotFin+'"';
  if fEtablissement <> '' then result := result + ' AND GL_ETABLISSEMENT ="'+fEtablissement+'"';
  result := result + ' AND (GL_QTERESTE <> 0 OR GL_MTRESTE <> 0) '; // --- GUINIER ---
  result := result + ' AND GL_VIVANTE="X" AND GP_VIVANTE="X"';
  result := result + ' AND GL_TYPEDIM<>"GEN" ' ;
  result := result + ' AND GL_TYPELIGNE="ART" ' ;
  result := result + ' ORDER BY GP_NUMERO';
end;

function TConstructDecisionAch.ConstitueRequeteLigneDoc: string;
begin
  result := 'SELECT *,GA_LIBELLE FROM LIGNE '+
  					'LEFT JOIN ARTICLE ON GA_ARTICLE=GL_ARTICLE '+
  					'WHERE GL_NATUREPIECEG IN ("'+GetNaturePieceCde +'") AND '+
  					'GL_DATELIVRAISON >="'+USDATETIME(fDateDeb)+'" AND GL_DATELIVRAISON <="'+USDATETIME(fDateFin)+'"';
  if fArticleDeb <> '' then result := result + ' AND GL_CODEARTICLE >="'+fArticleDeb+'"';
  if fArticleFin <> '' then result := result + ' AND GL_CODEARTICLE <="'+fArticleFin+'"';
  if fDepotDeb <> '' then result := result + ' AND GL_DEPOT >="'+fDepotDeb+'"';
  if fDepotFin <> '' then result := result + ' AND GL_DEPOT <="'+fDepotFin+'"';
  if fEtablissement <> '' then result := result + ' AND GL_ETABLISSEMENT ="'+fEtablissement+'"';
  result := result + ' AND (GL_QTERESTE <> 0 OR GL_MTRESTE <> 0) AND GL_VIVANTE="X"';
  result := result + ' AND GL_SOLDERELIQUAT <> "X"';
  result := result + ' AND GL_TYPEDIM<>"GEN" ' ;
  result := result + ' AND GL_TYPELIGNE="ART" ' ;
  if fTOBSelect.detail.count > 0 then result := AjouteFiltreDoc (result);
  result := result + ' ORDER BY GL_ARTICLE';
end;

constructor TConstructDecisionAch.create;
begin
  fRegroupLibellediff := GetParamSocSecur('SO_REGROUPLIBELLEDIFF',false);
  fTOBSelect := TOB.Create ('LES ENTETES SELECTIONNES',nil,-1);
  fTOBEnt := TOB.Create ('DECISIONACH',nil,-1);
  fTOBDetail := TOB.Create ('LES LIGNES DE DECISIONACH',nil,-1);
  fTOBLigneBrut := TOB.Create ('LISTE DES LIGNES A TRAITER',nil,-1);
end;

function TConstructDecisionAch.DejaCommande(TOBL: TOB): double;
var QQ : TQuery;
		req : string;
    TOBAch,TOBA : TOB;
    indice : integer;
    FUA,FUV : double;
begin
	result := 0;
  TOBAch := TOB.Create ('LES LIGNES COMMANDES FOUR',nil,-1);
  Req := 'SELECT GL_NATUREPIECEG,GL_ARTICLE,GL_QTERESTE, GL_QUALIFQTEACH,GL_COEFCONVQTE FROM LIGNE WHERE GL_NATUREPIECEG IN ('+
  				GetPieceAchat (true,true,true,true)+
  			  ') AND GL_PIECEORIGINE="'+EncodeRefPiece (TOBL)+'" AND '+
          '(GL_QTERESTE <> 0) AND GL_VIVANTE="X"';
  QQ := OpenSql (Req,true,-1,'',true);
  if not QQ.eof then
  begin
  	TOBAch.LoadDetailDB('LIGNE','','',QQ,false);
    for Indice := 0 to TOBAch.detail.count -1 do
    begin
      TOBA := TOBACH.detail[Indice];
      FUV := RatioMesure('PIE', TobA.GetValue('GL_QUALIFQTEVTE')); if FUV = 0 then FUV := 1;
      FUA := RatioMesure('PIE', TobA.GetValue('GL_QUALIFQTEACH')); if FUA = 0 then FUA := 1;
      result := result + (TOBA.GetValue('GL_QTERESTE') * FUA / FUV);
    end;
  end;
  ferme (QQ);

  TOBAch.clearDetail;

  Req := 'SELECT BAD_QUANTITEVTE FROM DECISIONACHLIG LEFT JOIN DECISIONACH ON BAD_NUMERO=BAE_NUMERO WHERE BAD_ARTICLE="'+
  			 TOBL.GetValue('GL_ARTICLE')+'" AND BAD_REFGESCOM="'+EncodeRefPiece (TOBL)+
         '" AND BAE_VALIDE="-"';
  QQ := OpenSql (Req,true,-1,'',true);
  if not QQ.eof then
  begin
  	TOBAch.LoadDetailDB('DECISIONACHLIG','','',QQ,false);
    for Indice := 0 to TOBAch.detail.count -1 do
    begin
      TOBA := TOBACH.detail[Indice];
      result := result + (TOBA.GetValue('BAD_QUANTITEVTE'));
    end;
  end;
  ferme (QQ);

end;

function TConstructDecisionAch.MtDejaCommande(TOBL: TOB): double;
var QQ : TQuery;
		req : string;
    TOBAch,TOBA : TOB;
    indice : integer;
    FUA,FUV : double;
begin

	result := 0;

  TOBAch := TOB.Create ('LES LIGNES COMMANDES FOUR',nil,-1);

  Req := 'SELECT GL_NATUREPIECEG,GL_ARTICLE,GL_MTRESTE,GL_QUALIFQTEACH,GL_COEFCONVQTE ';
  req := Req + ' FROM LIGNE WHERE GL_NATUREPIECEG IN ('+ GetPieceAchat (true,true,true,true) + ')';
  req := req + '  AND GL_PIECEORIGINE="' + EncodeRefPiece(TOBL) + '"';
  req := req + '  AND (GL_MTRESTE >0) AND GL_VIVANTE="X"';

  QQ := OpenSql (Req,true,-1,'',true);

  if not QQ.eof then
  begin
  	TOBAch.LoadDetailDB('LIGNE','','',QQ,false);
    for Indice := 0 to TOBAch.detail.count -1 do
    begin
      TOBA := TOBACH.detail[Indice];
      result := result + TOBA.GetValue('GL_MTRESTE');
    end;
  end;

  ferme (QQ);

  FreeAndNil(TOBAch);

end;

destructor TConstructDecisionAch.destroy;
begin
  inherited;
  fTOBSelect.free;
  fTOBEnt.free;
  fTOBDetail.free;
  fTOBLigneBrut.free;
end;

function TConstructDecisionAch.EcritDocument: boolean;
begin
	result := false;
  if fTOBDetail.detail.count = 0 then exit;
	if TRANSACTIONS (EcritLesElements,1) = oeOk then result := true;
end;

procedure TConstructDecisionAch.EcritLesElements;
begin
	fTOBDetail.SetAllModifie(true);
  if not fTOBDetail.InsertDB (nil,true) then V_PGI.IOError := Oeunknown;
  if V_PGI.Ioerror = OeOk then if not fTOBEnt.InsertDB (nil,true) then V_PGI.IOError := Oeunknown;
end;

function TConstructDecisionAch.Generate: boolean;
var Parametres : R_ParamTrait;
begin
	result := true;
  if ISPrepaLivFromAppro then
  begin
    parametres.ArtDeb := ArticleDebut;
    parametres.ArtFin := ArticleFin;
    parametres.DepDeb := DepotDebut;
    parametres.DepFin := DepotFin;
    parametres.Fam1 := '';
    parametres.Fam2 := '';
    parametres.Fam3 := '';
    parametres.Fam1_ := 'zzz';
    parametres.Fam2_ := 'zzz';
    parametres.Fam3_ := 'zzz';
    parametres.DatDeb := DateToStr (DateDebut);
    parametres.DatFin := DateToStr (DateFin);
    GenereCommandeChantierFromRea (parametres);
  end;
  if WithSelection then
  begin
  	if not SelectDocATraiter then BEGIN result := false; exit; END;
  end;
  if not RecupereLigneDetail then BEGIN result := false; exit; END;
  if not TraiteLigneDetail then BEGIN RESULT := false; exit; END;
  if not SetInfosEntete then BEGIN RESULT := false ; exit; END;
  if not EcritDocument then BEGIN Result := false ; exit; END;
end;

procedure TConstructDecisionAch.GenereCumul;
var Indice : integer;
		TOBD : TOB;
begin
  for Indice := 0 to fTOBDetail.detail.count -1 do
  begin
  	TOBD := fTOBDetail.Detail[Indice];
    SetCumul (nil,TOBD);
  end;
end;

function TConstructDecisionAch.GetNumeroPiece(var TheNumero: integer): boolean;
var TheNumPrec : integer;
		QQ : TQuery;
    requete : string;
    CptArret : integer;
    MajOk : boolean;
begin
	result := true;
  CptArret := 1;
  MajOk := false;
	QQ := OpenSQL('SELECT SH_NUMDEPART FROM SOUCHE WHERE SH_TYPE="CPT" AND SH_SOUCHE="BDA"', True,-1,'',true);
  if not QQ.eof then
  begin
  	ferme (QQ);
    repeat
      QQ := OpenSQL('SELECT SH_NUMDEPART FROM SOUCHE WHERE SH_TYPE="CPT" AND SH_SOUCHE="BDA"', True,-1,'',true);
      TheNumprec := QQ.findfield('SH_NUMDEPART').AsInteger;
      ferme (QQ);
      TheNumero := TheNumPrec + 1;
      requete := 'UPDATE SOUCHE SET SH_NUMDEPART ='+IntToStr(TheNumero)+' WHERE SH_TYPE="CPT" AND SH_SOUCHE="BDA" AND'+
                 ' SH_NUMDEPART='+IntToStr(TheNumprec);
      if ExecuteSQL (requete) <= 0 then inc(CptArret) else MajOk := true;
    until (CptArret > 10) or (MajOk);
    result := MajOk;
  end else
  begin
    ferme (QQ);
  	PgiBox ('La souche BDA n''existe pas..il faut la creer',fecran.caption);
  end
end;

procedure TConstructDecisionAch.Initligne(LaTOBCourante,TOBL: TOB; Niveau : integer);
var LivChantier : string;
		FUS,FUV,FUA,PQQ,COEFUAUS,CoefUSUV : double;
    TarifFourn : TGinfostarifFour;
    LeFournisseur,Larticle,UniteAch : string;
    QQ : TQuery;
begin

	LaTOBCourante.putValue('MULTIPABASE','-');

	if (TOBL.GetValue('GL_IDENTIFIANTWOL')=-1) then LivChantier := 'X' else LivChantier := '-';
  UniteAch :=  TobL.GetValue('GL_QUALIFQTEACH');
  LeFournisseur := TOBL.GetValue('GL_FOURNISSEUR');
  Larticle := TOBL.GetValue('GL_ARTICLE');
  CoefUaUs := 0;
  CoefUSUV := TOBL.GetDouble('GL_COEFCONVQTEVTE');
  if leFournisseur <> '' then
  begin
  	QQ := OpenSql ('SELECT GCA_COEFCONVQTEACH FROM CATALOGU WHERE GCA_TIERS="'+leFournisseur+'" AND GCA_ARTICLE="'+Larticle+'"',true,1,'',true);
    if not QQ.eof then
    begin
    	CoefUaUs := QQ.Fields[0].AsFloat;
    end;
    ferme (QQ);
  end;
  if COEFUAUS = 0 then
  begin
  	QQ := OpenSql ('SELECT GA_COEFCONVQTEACH,GA_QUALIFUNITEACH FROM ARTICLE WHERE GA_ARTICLE="'+Larticle+'"',true,1,'',true);
    if not QQ.eof then
    begin
    	CoefUaUs := QQ.Fields[0].AsFloat;
      UniteAch := QQ.Fields[1].AsString;
    end;
    ferme (QQ);
  end;
  FUV := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTEVTE')); if FUV = 0 then FUV := 1;
  FUS := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTESTO')); if FUS = 0 then FUS := 1;
  FUA := RatioMesure('PIE', UniteAch); if FUA = 0 then FUA := 1;
  if CoefUsUv = 0 then CoefUsUv := FUS/FUV;

  if Niveau < 4 then
  begin
    if Niveau = 1 then
    begin
      LaTOBCourante.PutValue('BAD_LIVCHANTIER','');
      LaTOBCourante.putValue('BAD_TYPEL','NI1');
      if fRegroupLibellediff then
      begin
    		LaTOBCOurante.putValue('BAD_LIBELLE',TOBL.GetValue('GA_LIBELLE'));
      end else
      begin
  		  LaTOBCOurante.putValue('BAD_LIBELLE',TOBL.GetValue('GL_LIBELLE'));
      end;
    end else if Niveau = 2 then
    begin
      LaTOBCourante.putValue('BAD_TYPEL','NI2');
      if fRegroupLibellediff then
      begin
    		LaTOBCOurante.putValue('BAD_LIBELLE',TOBL.GetValue('GA_LIBELLE'));
      end else
      begin
  		  LaTOBCOurante.putValue('BAD_LIBELLE',TOBL.GetValue('GL_LIBELLE'));
      end;
    end else if Niveau = 3 then
    begin
      LaTOBCourante.putValue('BAD_TYPEL','NI3');
      if fRegroupLibellediff then
      begin
    		LaTOBCOurante.putValue('BAD_LIBELLE',TOBL.GetValue('GA_LIBELLE'));
      end else
      begin
    		LaTOBCOurante.putValue('BAD_LIBELLE',TOBL.GetValue('GL_LIBELLE'));
      end;
    end;
  end else
  begin
  	LaTOBCourante.putValue('BAD_TYPEL','ART');
  	LaTOBCOurante.putValue('BAD_LIBELLE',TOBL.GetValue('GL_LIBELLE'));
  end;
  LaTOBCOurante.putValue('BAD_ETABLISSEMENT',TOBL.GetValue('GL_ETABLISSEMENT'));
  LaTOBCOurante.putValue('BAD_TYPEARTICLE',TOBL.GetValue('GL_TYPEARTICLE'));
  LaTOBCOurante.putValue('BAD_CODEARTICLE',TOBL.GetValue('GL_CODEARTICLE'));
  LaTOBCOurante.putValue('BAD_ARTICLE',TOBL.GetValue('GL_ARTICLE'));
  LaTOBCOurante.putValue('BAD_PRISENCOMPTE','X');
  LaTOBCOurante.putValue('BAD_QUALIFQTEACH',UniteAch);
  LaTOBCOurante.putValue('BAD_QUALIFQTEVTE',TOBL.GetValue('GL_QUALIFQTEVTE'));
  LaTOBCOurante.putValue('BAD_QUALIFQTESTO',TOBL.GetValue('GL_QUALIFQTESTO'));
  if Niveau > 1 then
  begin
  	LaTOBCourante.PutValue('BAD_LIVCHANTIER',LivChantier);
  end;

  if Niveau > 2 then
  begin
  	LaTOBCourante.PutValue('BAD_DEPOT',TOBL.GEtValue('GL_DEPOT'));
  end;

  LaTOBCOurante.putValue('BAD_TENUESTOCK',TOBL.GetValue('GL_TENUESTOCK'));
  LaTOBCOurante.putValue('BAD_COEFUSUV',CoefUSUV);

  if Niveau > 3 then
  begin
    if CtrlOkReliquat(TOBL, 'GL') then
     	LaTOBCOurante.putValue('BAD_MTINIT', TOBL.GetValue('GL_MTRESTE') - TOBL.getValue('DEJA COMMANDE'))
    else
    	LaTOBCOurante.putValue('BAD_QUANTITEINIT',TOBL.GetValue('GL_QTERESTE') - TOBL.getValue('DEJA COMMANDE'));
  	LaTOBCOurante.putValue('BAD_QUANTITEVTE',LaTOBCOurante.GetValue('BAD_QUANTITEINIT'));
    LaTOBCOurante.putValue('BAD_QUANTITESTK',LaTOBCOurante.GetValue('BAD_QUANTITEINIT')/ CoefUSUV);
    if CoefUaus = 0 then CoefuaUs := FUA/FUS;
    LaTOBCOurante.putValue('BAD_QUANTITEACH',LaTOBCOurante.GetValue('BAD_QUANTITESTK')/CoefUaUs);
  	LaTOBCOurante.putValue('BAD_COEFUAUS',CoefuaUs);
  	LaTOBCOurante.putValue('BAD_AFFAIRE',TOBL.GetValue('GL_AFFAIRE'));
//  	LaTOBCOurante.putValue('BAD_AFFAIRE0',TOBL.GetValue('GL_AFFAIRE0'));
  	LaTOBCOurante.putValue('BAD_AFFAIRE1',TOBL.GetValue('GL_AFFAIRE1'));
  	LaTOBCOurante.putValue('BAD_AFFAIRE2',TOBL.GetValue('GL_AFFAIRE2'));
  	LaTOBCOurante.putValue('BAD_AFFAIRE3',TOBL.GetValue('GL_AFFAIRE3'));
  	LaTOBCOurante.putValue('BAD_AVENANT',TOBL.GetValue('GL_AVENANT'));
    if (LeFournisseur <> '') and (TOBL.GetString('GL_NATUREPIECEG')<>'CC') then
    begin
      GetInfoFromCatalog(LeFournisseur,TOBL.GetValue('GL_ARTICLE'),UniteAch,PQQ,CoefUaUs);
      if UniteAch <> '' then
      begin
      	LaTOBCOurante.putValue('BAD_QUALIFQTEACH',UniteAch);
      end;
      if CoefUaUs = 0 then CoefUaus := 1;
      if CoefUsUV = 0 then CoefUsUV := 1;
  		LaTOBCOurante.putValue('BAD_COEFUAUS',CoefuaUs);
      LaTOBCOurante.putValue('BAD_PRIXACH',TOBL.GetValue('GL_DPA')*CoefUSUV);
      LaTOBCOurante.putValue('BAD_PRIXACHNET',TOBL.GetValue('GL_DPA')*CoefUSUV);
      LaTOBCOurante.putValue('BAD_FOURNISSEUR',LeFournisseur);
    end else
    begin
      TarifFourn := GetInfosTarifAch (LeFournisseur,Larticle,TurStock,false,true,nil,LaTOBCOurante.GetValue('BAD_QUANTITEACH'),TOBL.GetValue('GL_DATELIVRAISON'));
      LaTOBCOurante.putValue('BAD_PRIXACH',TarifFourn.TarifAchBrut );
      LaTOBCOurante.putValue('BAD_CALCULREMISE',TarifFourn.Remise  );
      LaTOBCOurante.putValue('BAD_PRIXACHNET',TarifFourn.TarifAch );
  		LaTOBCOurante.putValue('BAD_COEFUAUS',TarifFourn.CoefUAUs );
      if (LeFournisseur <> '') then LaTOBCOurante.putValue('BAD_FOURNISSEUR',LeFournisseur);
    end;
    LaTOBCOurante.putValue('BAD_PABASE',TOBL.GetValue('GL_DPA'));
    LaTOBCOurante.putValue('BAD_REFGESCOM',EncodeRefPiece(TOBL));
  	LaTOBCOurante.putValue('BAD_DATELIVRAISON',TOBL.GetValue('GL_DATELIVRAISON'));
  	LaTOBCOurante.putValue('BAD_BLOCNOTE',TOBL.GetValue('GL_BLOCNOTE'));
  end else
  begin
    if (LaTOBCOurante.GetValue('BAD_DATELIVRAISON')=iDate1900) and (LaTOBCOurante.GetValue('MULTI')='-') then
    begin
  	  LaTOBCOurante.putValue('BAD_DATELIVRAISON',TOBL.GetValue('GL_DATELIVRAISON'));
    end else if (LaTOBCOurante.GetValue('BAD_DATELIVRAISON')<>TOBL.GetValue('GL_DATELIVRAISON')) and (LaTOBCOurante.GetValue('MULTI')<>'X') then
    begin
      LaTOBCOurante.putValue('BAD_FOURNISSEUR','');
      LaTOBCOurante.putValue('BAD_PRIXACH',0);
      LaTOBCOurante.putValue('BAD_CALCULREMISE','');
      LaTOBCOurante.putValue('BAD_PRIXACHNET',0);
  	  LaTOBCOurante.putValue('BAD_DATELIVRAISON',iDate2099);
  	  LaTOBCOurante.putValue('MULTI','X');
    end;
    if (LaTOBCOurante.GetValue('BAD_FOURNISSEUR') = '') and (LaTOBCOurante.GetValue('MULTI')='-') then
    begin
      LaTOBCOurante.putValue('BAD_FOURNISSEUR',LeFournisseur);
      if (LeFournisseur <> '') and (TOBL.GetString('GL_NATUREPIECEG')<>'CC') then
      begin
        GetInfoFromCatalog(LeFournisseur,TOBL.GetValue('GL_ARTICLE'),UniteAch,PQQ,CoefuaUs);
        if UniteAch <> '' then
        begin
          LaTOBCOurante.putValue('BAD_QUALIFQTEACH',UniteAch);
        end;
        if CoefUaUs = 0 then CoefUaus := 1;
        if CoefUsUV = 0 then CoefUsUV := 1;
  			LaTOBCOurante.putValue('BAD_COEFUAUS',CoefuaUs);
        LaTOBCOurante.putValue('BAD_PRIXACH',TOBL.GetValue('GL_DPA')*COEFUSUV);
        LaTOBCOurante.putValue('BAD_PRIXACHNET',TOBL.GetValue('GL_DPA')*COEFUSUV);
        LaTOBCOurante.putValue('BAD_FOURNISSEUR',LeFournisseur);
      end else
      begin
        TarifFourn := GetInfosTarifAch (LeFournisseur,Larticle,Turstock,false,true,nil,LaTOBCOurante.GetValue('BAD_QUANTITEACH'),TOBL.GetValue('GL_DATELIVRAISON'));
        if CoefUaUs = 0 then CoefUaus := 1;
        if CoefUsUV = 0 then CoefUsUV := 1;
        LaTOBCOurante.putValue('BAD_PRIXACH',TarifFourn.TarifAchBrut );
        LaTOBCOurante.putValue('BAD_CALCULREMISE',TarifFourn.Remise  );
        LaTOBCOurante.putValue('BAD_PRIXACHNET',TarifFourn.TarifAch );
        LaTOBCOurante.putValue('BAD_QUALIFQTEACH',UniteAch);
        LaTOBCOurante.putValue('BAD_COEFUAUS',COEFUAUS );
      end;
//      LaTOBCOurante.putValue('BAD_PABASE',TOBL.GetValue('GL_DPA'));
    end else if (LaTOBCOurante.GetValue('BAD_FOURNISSEUR') <> leFournisseur) and (LaTOBCOurante.GetValue('MULTI')<>'X') then
    begin
  	  LaTOBCOurante.putValue('MULTI','X');
      LaTOBCOurante.putValue('BAD_FOURNISSEUR','');
      LaTOBCOurante.putValue('BAD_PRIXACH',0);
      LaTOBCOurante.putValue('BAD_CALCULREMISE','');
      LaTOBCOurante.putValue('BAD_PRIXACHNET',0);
  	  LaTOBCOurante.putValue('BAD_DATELIVRAISON',iDate2099);
//      LaTOBCOurante.putValue('BAD_PABASE',0);
    end;
  end;
end;

procedure TConstructDecisionAch.LanceFicheSaisie;
begin
	Ouvredecisionnel('NUMDECISIONNEL='+IntToStr(TheNumero));
end;

procedure NumeroteLigne (FF : Tform; TOBdecisionnel : TOB);

  procedure SetNumligne(FF: Tform; TOBPere: TOB; IndiceN1,IndiceN2,IndiceN3,IndiceN4,Niveau: integer);
  var Indice : integer;
      TOBD : TOB;
      Cpt1,Cpt2,Cpt3,Cpt4 : integer;
  begin
    if TOBPere.detail.count = 0 then exit;

    Cpt1 := IndiceN1;
    Cpt2 := IndiceN2;
    Cpt3 := IndiceN3;
    Cpt4 := IndiceN4;

    for Indice := 0 to TOBpere.detail.count -1 do
    begin
      TOBD := TOBPere.detail[Indice];
      if Niveau = 2 then inc(Cpt2);
      if Niveau = 3 then inc(Cpt3);
      if Niveau = 4 then inc(Cpt4);
      //
      TOBD.putValue('BAD_NUMN1',Cpt1);
      TOBD.putValue('BAD_NUMN2',Cpt2);
      TOBD.putValue('BAD_NUMN3',Cpt3);
      TOBD.putValue('BAD_NUMN4',Cpt4);
      if (FF IS TFvierge) and (TOF_DECISIONACHAT(TFVierge(FF).LaTOF).TheConsultFou <> nil) then
      begin
        TOF_DECISIONACHAT(TFVierge(FF).LaTOF).TheConsultFou.Renumerote (TOBD);
      end;
      if TOBD.detail.count > 0 then SetNumligne (FF,TOBD,cpt1,Cpt2,CPt3,CPt4,Niveau+1);
    end;

  end;

var Indice : integer;
		TOBD : TOB;
begin
  for Indice := 0 to TOBdecisionnel.detail.count -1 do
  begin
    TOBD := TOBdecisionnel.detail[Indice];
    TOBD.putvalue('BAD_NUMN1',Indice+1);
    if (FF IS TFvierge) and (TOF_DECISIONACHAT(TFVierge(FF).LaTOF).TheConsultFou <> nil) then
    begin
      TOF_DECISIONACHAT(TFVierge(FF).LaTOF).TheConsultFou.Renumerote (TOBD);
    end;
    SetNumligne (FF,TOBD,Indice+1,0,0,0,2);
  end;
end;

function TConstructDecisionAch.RecupereLigneDetail: boolean;
var Select : string;
		QQ : Tquery;
begin
	result := true;
	select := ConstitueRequeteLigneDoc;
  QQ := OpenSql (Select,true,-1,'',true);
  if not QQ.eof then fTOBLigneBrut.loadDetailDb ('LIGNE','','',QQ,false);
  ferme (QQ);
  if fTOBLigneBrut.detail.count = 0 then
  BEGIN
  	PGIBox ('Aucun document à traiter',Fecran.caption);
    result := false;
    Exit;
  END;
end;

function TConstructDecisionAch.SelectDocATraiter: boolean;
var TheRequete : string;
		QQ : TQuery;
begin
  TheRequete := ConstitueRequeteEnteteDoc;
  QQ := OpenSql (TheRequete,true,-1,'',true);
  if not QQ.eof then fTOBSelect.LoadDetailDB ('PIECE','','',QQ,false,true);
  ferme(QQ);
  if fTOBSelect.detail.count = 0 then
  BEGIN
  	PGIBox ('Aucun document à traiter',Fecran.caption);
    result := false;
    exit ;
  END;
  result := SelectionDocuments ('','','',fTOBSelect);
  if not result then
  begin
  	PGIBox ('Sélection annulée',Fecran.caption);
  end;
end;

function TConstructDecisionAch.SetCumul(TOBPere, TOBDetail: TOB): boolean;
var Indice : integer;
		TOBD : TOB;
begin
	result := true;
  if TOBDetail.detail.count > 0 then
  begin
  	for Indice := 0 to TOBDetail.detail.count -1 do
    begin
      TOBD := TOBDetail.detail[Indice];
      SetCumul (TOBDetail,TOBD);
    end;
  end;
  if TOBPere <> nil then
  begin
    if TOBpere.getValue ('MULTIPABASE') = '-' then
    begin
      if (Valeur(TOBpere.GetValue('BAD_PABASE'))=0) then
      begin
         TOBPere.putValue('BAD_PABASE',TOBdetail.GetValue('BAD_PABASE'));
      end else if (TOBpere.GetValue('BAD_PABASE')<>TOBdetail.GetValue('BAD_PABASE')) then
      begin
         TOBPere.putValue('BAD_PABASE',0);
         TOBpere.putValue ('MULTIPABASE','X');
      end;
    end;
  	TOBPere.putvalue ('BAD_QUANTITEACH',TOBPere.Getvalue ('BAD_QUANTITEACH')+TOBDetail.Getvalue ('BAD_QUANTITEACH'));
  	TOBPere.putvalue ('BAD_QUANTITEVTE',TOBPere.Getvalue ('BAD_QUANTITEVTE')+TOBDetail.Getvalue ('BAD_QUANTITEVTE'));
  	TOBPere.putvalue ('BAD_QUANTITEINIT',TOBPere.Getvalue ('BAD_QUANTITEINIT')+TOBDetail.Getvalue ('BAD_QUANTITEINIT'));
  	TOBPere.putvalue ('BAD_QUANTITESTK',TOBPere.Getvalue ('BAD_QUANTITESTK')+TOBDetail.Getvalue ('BAD_QUANTITESTK'));
    if TOBPere.Getvalue ('BAD_DATELIVRAISON') > TOBDetail.Getvalue ('BAD_DATELIVRAISON') then
  			TOBPere.putvalue ('BAD_DATELIVRAISON',TOBDetail.Getvalue ('BAD_DATELIVRAISON'));
  end;
end;

function TConstructDecisionAch.SetInfosEntete: boolean;
begin
	result := true;
	if not GetNumeroPiece (TheNumero) then BEGIN result := false; exit; END;
  fTOBent.PutValue ('BAE_ETABLISSEMENT',fEtablissement);
  fTOBent.PutValue ('BAE_DATE',V_PGI.DateEntree);
  fTOBEnt.PutValue ('BAE_BLOCNOTE',frecap.LinesRTF.Text );
  fTOBEnt.putValue ('BAE_NUMERO',TheNumero);
  AppliqueNumerotationLigne (fTOBDetail,TheNumero);
end;

function TConstructDecisionAch.TraiteLigneDetail: boolean;
var Indice : integer;
		TOBL : TOB;
begin
	result := true;
  for Indice := 0 to fTOBLigneBrut.detail.count -1 do
  begin
  	TOBL := fTOBLigneBrut.detail[Indice];
    // --- GUINIER ---
    if CtrlOkReliquat(TOBL, 'GL') then
    begin
      TOBL.AddChampSupValeur ('DEJA COMMANDE', MtDejaCommande(TOBL));
      if abs(TOBL.getValue('GL_MTRESTE')) <= abs(TOBL.getValue('DEJA COMMANDE')) then continue;
    end
    else
    begin
      TOBL.AddChampSupValeur ('DEJA COMMANDE', DejaCommande(TOBL));
      if abs(TOBL.getValue('GL_QTERESTE')) <= abs(TOBL.getValue('DEJA COMMANDE')) then continue;
    end;
    AjouteTOBLigne (TOBL);
  end;
  GenereCumul;
  NumeroteLigne (fecran,fTOBDetail);
end;

end.

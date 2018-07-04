unit UImportLigneExcelFac;

interface
uses sysutils,classes,windows,messages,controls,forms,hmsgbox,stdCtrls,clipbrd,nomenUtil,
     HCtrls,SaisUtil,HEnt1,Ent1,EntGC,UtilPGI,UTOB,HTB97,FactUtil,FactComm,Menus,
     FactTOB, FactPiece, FactArticle,
		 UtilXlsBTP,EtudesExt,ed_tools,variants,
{$IFDEF EAGLCLIENT}
  maineagl,
{$ELSE}
  Doc_Parser,DBCtrls, Db,  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} fe_main,UserChg,AglIsoflex,
{$ENDIF}
{$IFDEF BTP}
 FactOuvrage,BTPUtil, CalcOLEGenericBTP,
{$ENDIF}
 facture,
 FactCpta,
 ParamSoc,FactCalc,
 UtilSuggestion,
 UtilArticle,UEntCommun;

type

	TimportLigneFacture = class (Tobject)
  	private
      uForm : TForm;
      fdocExcel : string;
      fWinExcel: OleVariant;
      fLignePGI : integer;
      LigneIni : integer;
      WorkBook,WorkSheet : Variant;
      TOBL : TOB;
      TOBPARIMP : TOB;
      procedure TraitePrixLigneCourante(TOBL: TOB; Pu: double);
      procedure TraitePALigneCourante (TOBL: TOB; PA: double);
      procedure TraitePrixRevientLigneCourante(TOBL: TOB; PR: double);
      procedure ConstitueThisOuvrage (TOBpiece,TOBL,TOBARTicles,TOBOuvrage: TOb ; fLignePgi : integer; MtMar,QTMO,MtINT,MtLOC,MtMAT,MTOUT,MTSt : string);
      function FindFather (TOBParag : TOB; Parag : string) : TOB;
      function AddNewParag (TOBPARAG : TOB; Parag,Designation : string) : TOB;
      procedure AjouteFinParag(TOBPARAG, TOBPiece, TOBTiers, TOBAffaire: TOB; var FlignePgi: integer);
      function FindLastparag(TOBParag: TOB): TOB;
      procedure RecalcPaPRPV(TOBL: TOB);

    public
      TOBPiece,TOBTiers,TOBAffaire,TOBArticles,TOBART,TOBOuvrage,TOBOuvrageP : TOB;
      TOBBases,TOBBasesL,TOBPorcs,TOBpieceRG,TOBbasesRG,TOBConds,TOBTarif : TOB;
      DEV : RDevise;
      affichage : integer;
    	constructor create;
      destructor destroy; override;
    	procedure ajusteDocument;
      procedure execute;
      property documentPGI :TForm write uForm;
      property documentExcel : string read fdocExcel write fdocExcel ;
      property LignePGI : integer read fLignePGI write fLignePGI ;
    	function findTobArt(ArticleDivers: string): TOB;

  end;
implementation
uses factspec,factureBTP,FactDomaines,UspecifPOC;

{ TimportLigneFacture }

constructor TimportLigneFacture.create;
var QQ : Tquery;
begin
  fWinExcel := unassigned;
  TOBTarif := TOB.Create ('TARIF',nil,-1);
  TOBPARIMP := TOB.Create ('BPARAMIMPXLS',nil,-1);
  QQ := OpenSql ('SELECT * FROM BPARAMIMPXLS',true,1,'',true);
  if not QQ.eof then TOBPARIMP.SelectDB('',QQ)
                else TOBPARIMP.SetInteger('B01_NUMPARAM',1);
end;

destructor TimportLigneFacture.destroy;
begin
  if not VarIsEmpty(fWinExcel) then fWinExcel.Quit;
  fWinExcel := unassigned;
  TOBPARIMP.free;
  FreeAndNil(TOBTarif);
  inherited;
end;

procedure TimportLigneFacture.execute;
var WinNew,PError : boolean;
		LigneXls,Vide,Error,Integre : integer;
    CodeArticle,Qte,Designation,PA,ArticleDivers,PR,Ua,UV : string;
    MtMar,QTMO,MtINT,MtLOC,MtMAT,MTOUT,MTSt,CodeParag,PUV : string;
    PxAch,Px,CoefuaUs : double;
    Rem : double;
    AvecSDO,Comment,Paragraphe : boolean;
    TypeImport : integer;
    TOBTMP,TOBPARAG,TOBP,TOBPC : TOB;
    DQte,dpu : double;
begin
	if fdocExcel = '' then exit;
  TOBPARAG := TOB.Create ('LES PARAGRAPHES',nil,-1);
  Affichage:=GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'),'GPP_TYPEPRESENT');
  TypeImport := TOBPARIMP.GetInteger ('B01_NUMPARAM');
  //
  LigneIni := fLignePGI;
  ArticleDivers := GetParamSoc ('SO_BTARTICLEDIV');
  ArticleDivers := CodeArticleUnique(ArticleDivers,'','','','','');

  //
  if not OpenExcel (true,fWinExcel,WinNew) then
  begin
    PGIBox (TraduireMemoire('Liaison Excel impossible'),'Liaison EXCEL');
    FDocExcel := '';
    exit;
  end;
  WorkBook := OpenWorkBook (FDocExcel ,fWinExcel);
  WorkSheet := SelectSheet (WorkBook,''); //modif BRL 9/08/16 : Pas de nom de feuille, la feuille active sera prise par défaut.
  LigneXls := 0;
  vide := 1;
  Integre := 0;
  Error := 0;
  repeat
  	if Vide > 20 then break;
  	inc(LigneXls);
    Perror := false;
    AvecSDO := false;
    Comment := false;
    TOBPC := nil;
    Paragraphe := false;
    //
    CodeParag := '';
    designation := '';
    Uv := '';
    Qte := '';
    QTMO := '';
    PA := '';
    Puv := '';
    CodeArticle := '';
    MtMar := '';
    MtINT := '';
    MtLOC := '';
    MtMAT := '';
    MTOUT := '';
    MTSt := '';
    //
    if TypeImport<>2 then
    begin
      CodeArticle := GetExcelFormated (WorkSheet,LigneXls,2);
      designation := trimright(trimleft(GetExcelText (WorkSheet,LigneXls,3)));
      Qte := GetExcelText (WorkSheet,LigneXls,1);
      if pos(',',Qte) > 0 then StringReplace (Qte,',','.',[rfReplaceAll]);
      PA := GetExcelText (WorkSheet,LigneXls,4);
      if pos(',',pA) > 0 then StringReplace (pA,',','.',[rfReplaceAll]);
      // Pas de récupération du PR en document d'achat
      if (TOBPiece.GetValue('GP_VENTEACHAT') = 'ACH') then
      begin
         PR := ''
      end else
      begin
        PR := GetExcelText (WorkSheet,LigneXls,5);
        if pos(',',pR) > 0 then StringReplace (pR,',','.',[rfReplaceAll]);
      end;
    end else
    begin
      // Colonne A
      CodeParag := GetExcelFormated (WorkSheet,LigneXls,1);    // paragraphe
      // Colonne B
      designation := trimright(trimleft(GetExcelText (WorkSheet,LigneXls,2)));
      // Colonne C
      UV := GetExcelText (WorkSheet,LigneXls,3);
      // COLONNE D
      Qte := GetExcelText (WorkSheet,LigneXls,4);
      if pos(',',Qte) > 0 then StringReplace (Qte,',','.',[rfReplaceAll]);
      // Colonne E
      QTMO := GetExcelText (WorkSheet,LigneXls,5);
      if pos(',',QTMO) > 0 then StringReplace (QTMO,',','.',[rfReplaceAll]);
      if (QTMO <> '') and ((TOBPARIMP.GetString('B01_CODEOUV')='') OR (TOBPARIMP.GetString('B01_CODEMO')='')) then PError := true;
      // Colonne F
      PA := GetExcelText (WorkSheet,LigneXls,6);
      if pos(',',pA) > 0 then StringReplace (pA,',','.',[rfReplaceAll]);
      // Colonne G
      PuV := GetExcelText (WorkSheet,LigneXls,7);
      if pos(',',puV) > 0 then StringReplace (puV,',','.',[rfReplaceAll]);
      // Colonne H
      CodeArticle := GetExcelFormated (WorkSheet,LigneXls,8);
      // Colonne I (Montant)
      // Colonne J
      MtMar := GetExcelText (WorkSheet,LigneXls,10);
      if pos(',',MtMar) > 0 then StringReplace (MtMar,',','.',[rfReplaceAll]);
      if (MtMar <> '') and ((TOBPARIMP.GetString('B01_CODEOUV')='') OR (TOBPARIMP.GetString('B01_CODEMAR')='')) then PError := true;
      // Colonne K
      MtINT := GetExcelText (WorkSheet,LigneXls,11);
      if pos(',',MtINT) > 0 then StringReplace (MtINT,',','.',[rfReplaceAll]);
      if (MtINT <> '') and ((TOBPARIMP.GetString('B01_CODEOUV')='') OR (TOBPARIMP.GetString('B01_CODEINTERIM')='')) then PError := true;
      // Colonne L
      MtLOC := GetExcelText (WorkSheet,LigneXls,12);
      if pos(',',MtLOC) > 0 then StringReplace (MtLOC,',','.',[rfReplaceAll]);
      if (MtLOC <> '') and ((TOBPARIMP.GetString('B01_CODEOUV')='') OR (TOBPARIMP.GetString('B01_CODELOCATION')='')) then PError := true;
      // Colonne M
      MtMAT := GetExcelText (WorkSheet,LigneXls,13);
      if pos(',',MtMAT) > 0 then StringReplace (MtMAT,',','.',[rfReplaceAll]);
      if (MtMAT <> '') and ((TOBPARIMP.GetString('B01_CODEOUV')='') OR (TOBPARIMP.GetString('B01_CODEMAT')='')) then PError := true;
      // Colonne N
      MTOUT := GetExcelText (WorkSheet,LigneXls,14);
      if pos(',',MTOUT) > 0 then StringReplace (MTOUT,',','.',[rfReplaceAll]);
      if (MTOUT <> '') and ((TOBPARIMP.GetString('B01_CODEOUV')='') OR (TOBPARIMP.GetString('B01_CODEOUTIL')='')) then PError := true;
      // Colonne O
      MTSt := GetExcelText (WorkSheet,LigneXls,15);
      if pos(',',MTSt) > 0 then StringReplace (MTSt,',','.',[rfReplaceAll]);
      if (MTSt <> '') and ((TOBPARIMP.GetString('B01_CODEOUV')='') OR (TOBPARIMP.GetString('B01_CODESOUSTRAIT')='')) then PError := true;

      if (Perror) then BEGIN inc(error); continue; END;
    end;
    // pour éliminer ligne entête :
    if  (CodeArticle<>'') and (qte<>'') and (designation<>'') and (pA<>'') and (Puv<> '') and (valeur(Qte)=0) then continue;

    if TypeImport<>2 then
    begin
      if (CodeArticle='') or (qte='') or (valeur(Qte)=0) then
      begin
        inc(vide);
        if (CodeArticle<>'') or (qte<>'')  or (designation<>'') or (pA<>'') or (PUV<> '') then inc(error);
        continue;
      end;
    end else
    begin
      if (CodeParag='') and (CodeArticle='') and (qte='') and (designation = '') then
      begin
        inc(vide);
        continue;
      end;
    end;

    vide := 0;
    if (MtMar<>'') or (QTMO<>'') or (MtINT<>'') or (MtLOC<>'') or (MtMAT<>'') or (MTOUT<>'') or (MTSt<>'') then
    begin
      AvecSDO := true;
      TOBART := findTobArt(CodeArticleUnique2(CodeArticle, ''));
      if (TOBART <> nil) and (TOBART.GetString('GA_TYPEARTICLE')<>'OUV') then
      begin
        TOBART := findTobArt(TOBPARIMP.GetString('B01_CODEOUV'));
        if TOBART = nil then BEGIN inc(Error); continue; END;
      end;
      if (TOBArt = nil) then
      begin
        TOBART := findTobArt(TOBPARIMP.GetString('B01_CODEOUV'));
        if TOBART = nil then BEGIN inc(Error); continue; END;
      end;
      if TOBPARIMP.GetString('B01_CODEMAR')<> '' then
      begin
        TOBTMP := findTobArt(TOBPARIMP.GetString('B01_CODEMAR'));
        if TOBTMP = nil then BEGIN inc(Error); continue; END;
      end;
      if TOBPARIMP.GetString('B01_CODEMO')<> '' then
      begin
        TOBTMP := findTobArt(TOBPARIMP.GetString('B01_CODEMO'));
        if TOBTMP = nil then BEGIN inc(Error); continue; END;
      end;
      if TOBPARIMP.GetString('B01_CODEINTERIM')<> '' then
      begin
        TOBTMP := findTobArt(TOBPARIMP.GetString('B01_CODEINTERIM'));
        if TOBTMP = nil then BEGIN inc(Error); continue; END;
      end;
      if TOBPARIMP.GetString('B01_CODELOCATION')<> '' then
      begin
        TOBTMP := findTobArt(TOBPARIMP.GetString('B01_CODELOCATION'));
        if TOBTMP = nil then BEGIN inc(Error); continue; END;
      end;
      if TOBPARIMP.GetString('B01_CODEMAT')<> '' then
      begin
        TOBTMP := findTobArt(TOBPARIMP.GetString('B01_CODEMAT'));
        if TOBTMP = nil then BEGIN inc(Error); continue; END;
      end;
      if TOBPARIMP.GetString('B01_CODEOUTIL')<> '' then
      begin
        TOBTMP := findTobArt(TOBPARIMP.GetString('B01_CODEOUTIL'));
        if TOBTMP = nil then BEGIN inc(Error); continue; END;
      end;
      if TOBPARIMP.GetString('B01_CODESOUSTRAIT')<> '' then
      begin
        TOBTMP := findTobArt(TOBPARIMP.GetString('B01_CODESOUSTRAIT'));
        if TOBTMP = nil then BEGIN inc(Error); continue; END;
      end;
    end else
    begin
      TOBART := findTobArt(CodeArticleUnique2(CodeArticle, ''));
      if (TOBART = nil) and (ArticleDivers <> '') then
      begin
        TOBART := findTobArt(ArticleDivers);
      end;
      
      if (CodeParag <> '') and (Qte = '') and (PA = '') and (Puv = '') then
      begin
        Paragraphe := true;
      end else if (qte='') and (pA='') and (Designation <> '')then
      begin
        Comment := true;
      end;
    end;
    //
    if (TOBART= nil) and (not Paragraphe) and (not Comment) then BEGIN inc(Error); continue; END;
    //
    TOBPC := nil;
    if Paragraphe then
    begin
      // pour la gestion de la fin de paragraphe
      TOBP :=  FindFather(TOBPARAG,CodeParag);
      if TOBP = nil then
      begin
        AjouteFinParag (TOBPARAG,TOBPiece,TOBTiers,TOBAffaire,FlignePgi);
        TOBPC := AddNewParag (TOBPARAG,CodeParag,designation);
      end else
      begin
        AjouteFinParag (TOBP,TOBPiece,TOBTiers,TOBAffaire,FlignePgi);
        TOBPC := AddNewParag (TOBP,CodeParag,designation);
      end;
    end;
    //
    TOBL := Newtobligne (TOBPiece,fLignePgi);
    InitialiseTOBLigne(TOBPiece,TOBTiers,TOBAffaire,flignePgi,false) ;
    if Paragraphe then
    begin
      // debut de paragraphe
      TOBL.PutValue('GL_TYPELIGNE', 'DP'+InttoStr(TOBPC.getInteger('NIVEAU')));
      TOBL.putValue('GL_LIBELLE',copy(designation,1,70));
      TOBL.putValue('GL_NIVEAUIMBRIC',InttoStr(TOBPC.getInteger('NIVEAU')));
      TOBL.putValue('GLC_NUMEROTATION',CodeParag);
      TOBL.SetDouble('POURCENTMARG',0);
      TOBL.SetDouble('POURCENTMARQ',0);
      TOBL.AddChampSupValeur ('ANCIENNUMLIGNE', 0);
      TOBL.AddChampSupValeur ('ANCIENNUMORDRE', 0);
    end else if Comment then
    begin
      TOBL.PutValue('GL_TYPELIGNE', 'COM');
      TOBL.putValue('GL_LIBELLE',copy(designation,1,70));
      TOBL.SetString('GL_BLOCNOTE',designation);
      if TOBPC <> nil then TOBL.putValue('GL_NIVEAUIMBRIC',InttoStr(TOBPC.getInteger('NIVEAU')));
    end else
    begin
      TOBPC := FindLastparag (TOBPARAG);
      //
      TOBL.PutValue('GL_ARTICLE',TOBART.GetValue('GA_ARTICLE'));
      TOBL.PutValue('GL_TYPELIGNE','ART');
      TOBL.PutValue('GL_TYPEARTICLE',TOBART.GetValue('GA_TYPEARTICLE'));
      TOBL.PutValue('GL_CODEARTICLE',TOBART.GetValue('GA_CODEARTICLE'));
      TOBL.putValue('GLC_NUMEROTATION',CodeParag);

      if CodeArticle = '' then TOBL.putValue('GL_REFARTSAISIE',TOBART.GetValue('GA_CODEARTICLE'))
                          else TOBL.putValue('GL_REFARTSAISIE',CodeArticle);

      TOBL.PutValue('GL_REFARTBARRE', TOBART.GetValue('GA_CODEBARRE'));
      TOBL.PutValue('GL_PRIXPOURQTE', TOBART.GetValue('GA_PRIXPOURQTE'));
      TOBL.PutValue('GL_LIBELLE', TOBART.GetValue('GA_LIBELLE'));
      TOBL.PutValue('GL_TYPEREF', 'ART');
      TOBL.PutValue('GL_RECALCULER','X');
      TOBL.PutValue('GL_IDENTIFIANTWOL',0);
      if (TOBPiece.GetString('GP_NATUREPIECEG')<>'CBT') or
         ((TOBPiece.GetString('GP_NATUREPIECEG')='CBT') and (not GetParamSocSecur('SO_BTLIVBESOINDEF',false))) then
      begin
        if GetParamSocSecur ('SO_BTLIVCHANTIER',True) then TOBL.PutValue('GL_IDENTIFIANTWOL',-1);
      end;
      ArticleVersLigne (TOBPiece,TOBART,TobConds,TOBL,TOBTiers);
      TOBL.PutValue('GL_QUALIFQTEVTE',UV);
  //  Récupération remises fournisseur si document d'achat et pas de pu renseigné sur la ligne
      if (TOBPiece.GetValue('GP_VENTEACHAT') = 'ACH') and (valeur(PA) = 0) then
      begin
        rem := 0;
        CoefUaUs := 1;
        TOBTARIF.InitValeurs;
        //fv1 : 22/05/2014 - FS#1008 - DELABOUDINIERE : Ajout contrôle référencement fournisseur en saisie de pièces d'achats
        Px := RechTarifArticle(TOBArt, TOBTarif, TOBTiers,TOBPIECE, TOBL);
        //
        Rem := TOBTarif.GetValue('GF_REMISE');
        TOBL.PutValue('GL_PUHTDEV',Px);
        TOBL.PutValue('GL_COEFCONVQTE',CoefUaUs);
        TOBL.PutValue('GL_REMISELIGNE', Rem);
        TOBL.PutValue('GL_REMISECASCADE', TOBTarif.GetValue('GF_CASCADEREMISE'));
      end;
  //
      dpu := valeur(PuV);
      if (pos(TOBART.GetString('GA_TYPEARTICLE'),'OUV;ARP')>0) and (not AvecSDO) then
      begin
        TraiteLesOuvrages(nil, TOBPiece, TOBArticles, TOBOuvrage ,nil,nil, fLignePGI,  False, DEV, true);
      end else if (TypeImport = 2) and (AvecSDO) then
      begin
        ConstitueThisOuvrage (TOBpiece,TOBL,TOBARTicles,TOBOuvrage,fLignePgi,MtMar,QtMO,MtINT,MtLOC,MtMAT,MTOUT,MTSt);
      end;
      
      if (designation<>'') then
      begin
        TOBL.putValue('GL_LIBELLE',copy(designation,1,70));
      end;

      if CodeArticle <> '' then
      begin
        if GetInfoParPiece(TOBPiece.getValue('GP_NATUREPIECEG'), 'GPP_BLOBART') = 'X' then
        begin
          TOBL.PutValue('GL_BLOCNOTE', TOBART.GetValue('GA_BLOCNOTE'));
        end;
      end;

      if valeur(Qte)=0 then DQte := 1 else DQte := Valeur(Qte);
      TOBL.SetDouble('GL_QTEFACT',  DQte);
      TOBL.SetDouble('GL_QTERESTE', DQte);
      TOBL.SetDouble('GL_QTESTOCK', DQte);

      if TOBPiece.getValue('GP_DOMAINE')<>'' then
      begin
        TOBL.putvalue('GL_DOMAINE',TOBPiece.getValue('GP_DOMAINE'));
      end;
      if TOBART.getValue('GA_PRIXPASMODIF')<>'X' Then
      begin
        AppliqueCoefDomaineLig (TOBL);
        if VH_GC.BTCODESPECIF = '001' then ApliqueCoeflignePOC (TOBL,DEV);
      end;

      if (not AvecSDO) then
      begin
        if (Valeur(PA) <> 0) then
        begin
          TraitePALigneCourante (TOBL,valeur(PA));
        end;
        if (dpu <> 0)then
        begin
          TraitePrixLigneCourante (TOBL,Dpu);
        end;
        if valeur(PR) <> 0 then
        begin
          TraitePrixRevientLigneCourante (TOBL,valeur(Pr));
        end;
      end;
      if TOBPC <> nil then TOBL.putValue('GL_NIVEAUIMBRIC',InttoStr(TOBPC.getInteger('NIVEAU')));
    end;
    inc(fLignePGI);
    inc(Integre);
  until LigneXls > 5000;
  //
  AjouteFinParag (TOBPARAG,TOBPiece,TOBTiers,TOBAffaire,FlignePgi);
  //
  if fLignePGi <> LigneINi then
  begin
   	NumeroteLignesGC (nil,TOBPiece);
    PgiInfo (format('%d enregistrements intégrés - %d erreurs rencontrées',[integre,error]));
  end;
  TOBPARAG.free;
end;

function TimportLigneFacture.findTobArt (ArticleDivers : string) : TOB;
var TOBA : TOB;
    QQ : Tquery;
begin
  result := nil;
  TOBA := TOBArticles.FindFirst (['GA_ARTICLE'],[ArticleDivers],true);
  if TOBA = nil then
  begin
    QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE FROM ARTICLE A '+
    							 'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
    							 'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+ArticleDivers+'"',true,-1,'',true);
    if not QQ.eof then
    begin
      TOBA := TOB.create ('ARTICLE',TOBArticles,-1);
      TOBA.SelectDB ('',QQ);
      InitChampsSupArticle (TOBA);
      CalculePrixArticle (TOBA);
      result := TOBA;
    end;
    ferme (QQ);
  end else result := TOBA;
end;

procedure TimportLigneFacture.ajusteDocument;
var indice : integer;
    TOBL : TOB;
    IndiceOuv : integer;
begin
  {$IFDEF BTP}
  for Indice := LigneIni to TOBPiece.detail.count -1 do
  begin
    TOBL := TOBpiece.detail[indice];
    IndiceOuv := TOBL.GetValue('GL_INDICENOMEN');
    if (IndiceOuv > 0) and (IsOuvrage(TOBL)) then ReAffecteLigOuv(IndiceOuv, TobL, TobOuvrage);
  end;
  {$ENDIF}
end;

procedure TimportLigneFacture.TraitePrixLigneCourante (TOBL : TOB;Pu : double);
var TypeArticle : string;
		EnHt,EnVte : boolean;

begin
//
  if (TOBPiece.GetValue('GP_VENTEACHAT') <> 'ACH') then EnVte := true else EnVte := false;

	TypeArticle := TOBL.getValue('GL_TYPEARTICLE');
  EnHt := (TOBPiece.getValue('GP_FACTUREHT')='X');
  if (TypeArticle = 'OUV') or (TypeArticle = 'ARP') then
  begin
  	TraitePrixOuvrage(TOBPiece,TOBL,TOBBases,TOBBasesL,TOBOuvrage, EnHt, Pu,0,DEV,false);
  end;
//
  if EnHt then
  begin
    TOBL.PutValue('GL_PUHTDEV',Pu);
    TOBL.putValue('GL_PUHT',DEVISETOPIVOTEx (Pu,DEV.Taux,DEV.Quotite,V_PGI.okdecP));
  end else
  begin
    TOBL.PutValue('GL_PUTTCDEV',Pu);
    TOBL.putValue('GL_PUTTC',DEVISETOPIVOTEx (Pu,DEV.Taux,DEV.Quotite,V_PGI.okdecP));
  end;
  if EnVte then
  begin
    if TOBL.GetDouble ('GL_DPA')= 0 then TOBL.SetDouble('GL_DPA',1);
    if TOBL.GetDouble ('GL_DPR')= 0 then TOBL.SetDouble('GL_DPR',1);
    TOBL.SetDouble('GL_COEFMARG',0);
  end;
//
	TOBL.putValue('GL_BLOQUETARIF','X');
end;

procedure TimportLigneFacture.TraitePrixRevientLigneCourante(TOBL: TOB; PR: double);
var TypeArticle : string;
		TOBOUV,TOBOA : TOB;
begin
	TypeArticle := TOBL.getValue('GL_TYPEARTICLE');
  if (TypeArticle = 'OUV') then exit;
  //
  TOBL.putValue('GL_DPR',PR);
  if TOBL.getValue('GL_DPA') <> 0 then
  begin
  	TOBL.putValue('GL_COEFFG',Arrondi((TOBL.getValue('GL_DPR')/TOBL.getValue('GL_DPA'))-1,4));
  end else
  begin
  	TOBL.putValue('GL_COEFFG',0);
  end;
  TOBL.putValue('GL_COEFMARG',0);
	if (TypeArticle = 'ARP') then
  begin
    if TOBL.getValue('GL_INDICENOMEN') = 0 then exit;
    TOBOUV := TOBOuvrage.detail[TOBL.getValue('GL_INDICENOMEN')-1];
    if TOBOUV = nil then exit;
    TOBOA := TOBOUV.detail[0].detail[0];
    if TOBOA = nil then exit;
    TOBOA.putValue('BLO_DPR',Pr);
    TOBOA.putValue('BLO_COEFFG',0);
    TOBOA.putValue('BLO_COEFMARG',0);
  end;
end;


procedure TimportLigneFacture.ConstitueThisOuvrage(TOBpiece, TOBL,TOBARTicles,TOBOuvrage: TOb; fLignePgi: integer; MtMar,QtMO,MtINT,MtLOC,MtMAT,MTOUT,MTSt : string);

  function  AjouteSousDetail (TOBLN,TOBPiece,TOBL: TOB) : TOB;
  begin
    result:=TOB.Create('LIGNEOUV',TOBLN,-1) ;
    InsertionChampSupOuv (result,false);
    TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
    result.putValue('BLO_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));
    //
    result.PutValue('BLO_NATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG')) ;
    result.PutValue('BLO_SOUCHE',TOBL.GetValue('GL_SOUCHE')) ;
    result.PutValue('BLO_DATEPIECE',TOBL.GetValue('GL_DATEPIECE')) ;
    result.PutValue('BLO_AFFAIRE',TOBL.GetValue('GL_AFFAIRE')) ;
    result.PutValue('BLO_AFFAIRE1',TOBL.GetValue('GL_AFFAIRE1')) ;
    result.PutValue('BLO_AFFAIRE2',TOBL.GetValue('GL_AFFAIRE2')) ;
    result.PutValue('BLO_AFFAIRE3',TOBL.GetValue('GL_AFFAIRE3')) ;
    result.PutValue('BLO_AVENANT',TOBL.GetValue('GL_AVENANT')) ;
    result.PutValue('BLO_NUMERO',TOBL.GetValue('GL_NUMERO')) ;
    result.PutValue('BLO_INDICEG',TOBL.GetValue('GL_INDICEG')) ;
    result.PutValue('BLO_NUMLIGNE',TOBL.GetValue('GL_NUMLIGNE')) ;
    result.PutValue('BLO_QTEDUDETAIL',1) ;
    result.PutValue('BLO_COMPOSE',TOBL.GetValue('GL_ARTICLE')) ;
    CopieOuvFromLigne(result,TOBL); // recup des infos en provenance de la ligne de doc.
  end;

  procedure GestionCommuneLigneOuv(TOBPiece,TOBLND,TOBTMP : TOB);
  begin

    if (IsPrestationSt(TOBLND)) then
    begin
      if TOBPiece.getValue('GP_APPLICFGST')='X' then
        TOBLND.putValue('BLO_NONAPPLICFRAIS','-')
      else
        TOBLND.putValue('BLO_NONAPPLICFRAIS','X');
      //
      if TOBPiece.getValue('GP_APPLICFCST')='X' then
        TOBLND.putValue('BLO_NONAPPLICFC','-')
      else
        TOBLND.putValue('BLO_NONAPPLICFC','X');
    end else
    begin
      TOBLND.putValue('BLO_NONAPPLICFRAIS','-');
      TOBLND.putValue('BLO_NONAPPLICFC','-');
    end;
    if TOBTMP.getValue('GA_PRIXPASMODIF')<>'X' Then
    begin
      AppliqueCoefDomaineActOuv (TOBLND);
      if VH_GC.BTCODESPECIF = '001' then ApliqueCoeflignePOC (TOBLND,DEV);
    end;
    if (Pos (TOBTMP.getValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne) > 0) then
    begin
      TOBLND.PutValue('BLO_TPSUNITAIRE',1);
    end else
    begin
      TOBLND.PutValue('BLO_TPSUNITAIRE',0);
    end;
    TOBLND.PutValue('GA_HEURE',TOBLND.GetValue('BLO_TPSUNITAIRE'));
    if TOBLND.GetDouble('BLO_COEFMARG')=0 then TOBLND.SetDouble('BLO_COEFMARG',1); 
    CalculeLigneAcOuv (TOBLND,DEV,true,TOBTMP);
    CalculMontantHtDevLigOuv (TOBLND,DEV);
    TOBLND.putvalue ('ANCPA',TOBLND.Getvalue ('BLO_DPA'));
    TOBLND.putvalue ('ANCPR',TOBLND.Getvalue ('BLO_DPR'));
    TOBLND.PutValue ('GA_PVHT',TOBLND.GetValue('BLO_PUHTDEV'));
    TOBLND.PutValue ('GA_PAHT',TOBLND.GetValue('BLO_DPA'));
  end;

  procedure getValoArticle (TOBLND,TOBTMp : TOB);
  begin
    TOBLND.PutValue('BLO_TENUESTOCK',  TOBTMp.GetValue('GA_TENUESTOCK'));
    TOBLND.PutValue('BLO_QUALIFQTEVTE',TOBTMp.GetValue('GA_QUALIFUNITEVTE'));
    TOBLND.PutValue('BLO_QUALIFQTESTO',TOBTMp.GetValue('GA_QUALIFUNITESTO'));
    TOBLND.PutValue('BLO_PRIXPOURQTE', TOBTMp.GetValue('GA_PRIXPOURQTE'));
    TOBLND.PutValue('BLO_DPA',         TOBTMp.GetValue('GA_PAHT'));
    TOBLND.PutValue('BLO_PMAP',        TOBTMp.GetValue('GA_PMAP'));
    TOBLND.PutValue('ANCPA',           TOBLND.GetValue('BLO_DPA'));
    TOBLND.PutValue('ANCPR',           TOBLND.GetValue('BLO_DPR'));
    TOBLND.PutValue('BLO_PUHT',        TOBTMp.GetValue('GA_PVHT'));
    TOBLND.PutValue('BLO_PUHTBASE',    TOBTMp.GetValue('GA_PVHT'));
    TOBLND.PutValue('BLO_PUHTDEV',     pivottodevise(TOBLND.GetValue('BLO_PUHTBASE'),DEV.Taux,DEV.quotite,V_PGI.OkdecP ));
    TOBLND.PutValue('ANCPV',           TOBLND.GetValue('BLO_PUHTDEV'));
    TOBLND.PutValue('BLO_PUTTC',       TOBTMp.GetValue('GA_PVTTC'));
    TOBLND.PutValue('BLO_PUTTCBASE',   TOBTMp.GetValue('GA_PVTTC'));
    TOBLND.PutValue('BLO_PUTTCDEV',    pivottodevise(TOBLND.GetValue('BLO_PUTTCBASE'),DEV.Taux,DEV.quotite,V_PGI.OkdecP ));
  end;

var TOBLN,TOBLND,TOBTMP : TOB;
    Valeurs : T_Valeurs;
    EnHt : boolean;
begin
  EnHt := (TOBPiece.getValue('GP_FACTUREHT')='X');
  //
  TOBLN:=TOB.Create('',TOBOuvrage,-1) ;
  InsertionChampSupOuv (TOBLN,false);
  TOBL.putValue('GL_INDICENOMEN',TOBOuvrage.detail.count);

  if (MtMar<> '') and (valeur(MtMar)<>0) then
  begin
    TOBLND := AjouteSousDetail (TOBLN,TOBPiece,TOBL);
    TOBTMP := findTobArt(TOBPARIMP.GetString('B01_CODEMAR'));
    CopieOuvFromArt (TOBPiece,TOBLND,TOBTMP,DEV);
    GetValoArticle (TOBLND,TOBTMP);
    TOBLND.SetString('BLO_CODEARTICLE',TOBTMP.GetString('GA_CODEARTICLE'));
    TOBLND.SetString('BLO_ARTICLE',TOBTMP.GetString('GA_ARTICLE'));
    TOBLND.SetString('BLO_LIBELLE',TOBTMP.GetString('GA_LIBELLE'));
    TOBLND.SetDouble('BLO_QTEFACT',1);
    TOBLND.SetDouble('BLO_DPA',valeur(MtMar));
    TOBLND.PutValue('ANCPA',           TOBLND.GetValue('BLO_DPA'));
    TOBLND.SetDouble('BLO_QTEDUDETAIL',1);
    TOBLND.SetString('BNP_TYPERESSOURCE',TOBTMP.GEtString('BNP_TYPERESSOURCE'));
    GestionCommuneLigneOuv(TOBpiece,TOBLND,TOBTMP);
  end;
  if (QtMo<> '') and (valeur(QtMo)<>0) then
  begin
    TOBLND := AjouteSousDetail (TOBLN,TOBPiece,TOBL);
    TOBTMP := findTobArt(TOBPARIMP.GetString('B01_CODEMO'));
    CopieOuvFromArt (TOBPiece,TOBLND,TOBTMP,DEV);
    GetValoArticle (TOBLND,TOBTMP);
    TOBLND.SetString('BLO_CODEARTICLE',TOBTMP.GetString('GA_CODEARTICLE'));
    TOBLND.SetString('BLO_ARTICLE',TOBTMP.GetString('GA_ARTICLE'));
    TOBLND.SetString('BLO_LIBELLE',TOBTMP.GetString('GA_LIBELLE'));
    TOBLND.SetDouble('BLO_QTEFACT',Valeur(Qtmo));
    TOBLND.SetDouble('BLO_QTEDUDETAIL',1);
    TOBLND.SetString('BNP_TYPERESSOURCE',TOBTMP.GEtString('BNP_TYPERESSOURCE'));
    GestionCommuneLigneOuv(TOBpiece,TOBLND,TOBTMP);
  end;
  if (MtINT<> '') and (valeur(MtINT)<>0) then
  begin
    TOBLND := AjouteSousDetail (TOBLN,TOBPiece,TOBL);
    TOBTMP := findTobArt(TOBPARIMP.GetString('B01_CODEINTERIM'));
    CopieOuvFromArt (TOBPiece,TOBLND,TOBTMP,DEV);
    GetValoArticle (TOBLND,TOBTMP);
    TOBLND.SetString('BLO_CODEARTICLE',TOBTMP.GetString('GA_CODEARTICLE'));
    TOBLND.SetString('BLO_ARTICLE',TOBTMP.GetString('GA_ARTICLE'));
    TOBLND.SetString('BLO_LIBELLE',TOBTMP.GetString('GA_LIBELLE'));
    TOBLND.SetDouble('BLO_QTEFACT',1);
    TOBLND.SetDouble('BLO_DPA',valeur(MtINT));
    TOBLND.PutValue('ANCPA', TOBLND.GetValue('BLO_DPA'));
    TOBLND.SetDouble('BLO_QTEDUDETAIL',1);
    TOBLND.SetString('BNP_TYPERESSOURCE',TOBTMP.GEtString('BNP_TYPERESSOURCE'));
    GestionCommuneLigneOuv(TOBpiece,TOBLND,TOBTMP);
  end;
  if (MtLOC<> '') and (valeur(MtLOC)<>0) then
  begin
    TOBLND := AjouteSousDetail (TOBLN,TOBPiece,TOBL);
    TOBTMP := findTobArt(TOBPARIMP.GetString('B01_CODELOCATION'));
    CopieOuvFromArt (TOBPiece,TOBLND,TOBTMP,DEV);
    GetValoArticle (TOBLND,TOBTMP);
    TOBLND.SetString('BLO_CODEARTICLE',TOBTMP.GetString('GA_CODEARTICLE'));
    TOBLND.SetString('BLO_ARTICLE',TOBTMP.GetString('GA_ARTICLE'));
    TOBLND.SetString('BLO_LIBELLE',TOBTMP.GetString('GA_LIBELLE'));
    TOBLND.SetDouble('BLO_QTEFACT',1);
    TOBLND.SetDouble('BLO_DPA',valeur(MtLOC));
    TOBLND.PutValue('ANCPA', TOBLND.GetValue('BLO_DPA'));
    TOBLND.SetDouble('BLO_QTEDUDETAIL',1);
    TOBLND.SetString('BNP_TYPERESSOURCE',TOBTMP.GEtString('BNP_TYPERESSOURCE'));
    GestionCommuneLigneOuv(TOBpiece,TOBLND,TOBTMP);
  end;
  if (MtMAT<> '') and (valeur(MtMAT)<>0) then
  begin
    TOBLND := AjouteSousDetail (TOBLN,TOBPiece,TOBL);
    TOBTMP := findTobArt(TOBPARIMP.GetString('B01_CODEMAT'));
    CopieOuvFromArt (TOBPiece,TOBLND,TOBTMP,DEV);
    GetValoArticle (TOBLND,TOBTMP);
    TOBLND.SetString('BLO_CODEARTICLE',TOBTMP.GetString('GA_CODEARTICLE'));
    TOBLND.SetString('BLO_ARTICLE',TOBTMP.GetString('GA_ARTICLE'));
    TOBLND.SetString('BLO_LIBELLE',TOBTMP.GetString('GA_LIBELLE'));
    TOBLND.SetDouble('BLO_QTEFACT',1);
    TOBLND.SetDouble('BLO_DPA',valeur(MtMAT));
    TOBLND.PutValue('ANCPA', TOBLND.GetValue('BLO_DPA'));
    TOBLND.SetDouble('BLO_QTEDUDETAIL',1);
    TOBLND.SetString('BNP_TYPERESSOURCE',TOBTMP.GEtString('BNP_TYPERESSOURCE'));
    GestionCommuneLigneOuv(TOBpiece,TOBLND,TOBTMP);
  end;
  if (MTOUT<> '') and (valeur(MTOUT)<>0) then
  begin
    TOBLND := AjouteSousDetail (TOBLN,TOBPiece,TOBL);
    TOBTMP := findTobArt(TOBPARIMP.GetString('B01_CODEOUTIL'));
    CopieOuvFromArt (TOBPiece,TOBLND,TOBTMP,DEV);
    GetValoArticle (TOBLND,TOBTMP);
    TOBLND.SetString('BLO_CODEARTICLE',TOBTMP.GetString('GA_CODEARTICLE'));
    TOBLND.SetString('BLO_ARTICLE',TOBTMP.GetString('GA_ARTICLE'));
    TOBLND.SetString('BLO_LIBELLE',TOBTMP.GetString('GA_LIBELLE'));
    TOBLND.SetDouble('BLO_QTEFACT',1);
    TOBLND.SetDouble('BLO_DPA',valeur(MTOUT));
    TOBLND.PutValue('ANCPA', TOBLND.GetValue('BLO_DPA'));
    TOBLND.SetDouble('BLO_QTEDUDETAIL',1);
    TOBLND.SetString('BNP_TYPERESSOURCE',TOBTMP.GEtString('BNP_TYPERESSOURCE'));
    GestionCommuneLigneOuv(TOBpiece,TOBLND,TOBTMP);
  end;
  if (MTSt<> '') and (valeur(MTSt)<>0) then
  begin
    TOBLND := AjouteSousDetail (TOBLN,TOBPiece,TOBL);
    TOBTMP := findTobArt(TOBPARIMP.GetString('B01_CODESOUSTRAIT'));
    CopieOuvFromArt (TOBPiece,TOBLND,TOBTMP,DEV);
    GetValoArticle (TOBLND,TOBTMP);
    TOBLND.SetString('BLO_CODEARTICLE',TOBTMP.GetString('GA_CODEARTICLE'));
    TOBLND.SetString('BLO_ARTICLE',TOBTMP.GetString('GA_ARTICLE'));
    TOBLND.SetString('BLO_LIBELLE',TOBTMP.GetString('GA_LIBELLE'));
    TOBLND.SetDouble('BLO_QTEFACT',1);
    TOBLND.SetDouble('BLO_DPA',valeur(MTSt));
    TOBLND.PutValue('ANCPA', TOBLND.GetValue('BLO_DPA'));
    TOBLND.SetDouble('BLO_QTEDUDETAIL',1);
    TOBLND.SetString('BNP_TYPERESSOURCE',TOBTMP.GEtString('BNP_TYPERESSOURCE'));
    GestionCommuneLigneOuv(TOBpiece,TOBLND,TOBTMP);
  end;

  InitTableau (Valeurs);
  CalculeOuvrageDoc (TOBLN,1,1,true,DEV,valeurs,EnHt);
  TOBL.Putvalue('GL_PUHTDEV',valeurs[2]);
  TOBL.Putvalue('GL_PUTTCDEV',valeurs[3]);
  TOBL.Putvalue('GL_PUHTNETDEV',valeurs[2]);
  TOBL.Putvalue('GL_PUTTCNETDEV',valeurs[3]);
  TOBL.Putvalue('GL_PUHT',DeviseToPivotEx(TOBL.GetValue('GL_PUHTDEV'),DEV.Taux,DEV.quotite,V_PGI.OkdecP));
  TOBL.Putvalue('GL_PUTTC',DevisetoPivotEx(TOBL.GetValue('GL_PUTTCDEV'),DEV.taux,DEV.quotite,V_PGI.okdecP));
  TOBL.Putvalue('GL_PUHTBASE',TOBL.GetValue('GL_PUHT'));
  TOBL.Putvalue('GL_PUTTCBASE',TOBL.GetValue('GL_PUTTC'));
  TOBL.Putvalue('GL_PUHTNET',TOBL.getValue('GL_PUHT'));
  TOBL.Putvalue('GL_PUTTCNET',TOBL.GetValue('GL_PUTTC'));
  TOBL.Putvalue('GL_DPA',valeurs[0]);
  TOBL.Putvalue('GL_DPR',valeurs[1]);
  TOBL.Putvalue('GL_PMAP',valeurs[6]);
  TOBL.Putvalue('GL_PMRP',valeurs[7]);
  TOBL.putvalue('GL_TPSUNITAIRE',valeurs[9]);
  //
  TOBL.PutValue('GL_MONTANTPAFG',valeurs[10]);
  TOBL.PutValue('GL_MONTANTPAFR',valeurs[11]);
  TOBL.PutValue('GL_MONTANTPAFC',valeurs[12]);
  TOBL.PutValue('GL_MONTANTFG',valeurs[13]);
  TOBL.PutValue('GL_MONTANTFR',valeurs[14]);
  TOBL.PutValue('GL_MONTANTFC',valeurs[15]);
  TOBL.PutValue('GL_MONTANTPA',valeurs[16]);
  TOBL.PutValue('GL_MONTANTPR',valeurs[17]);
  //
  if TOBL.getValue('GL_TYPEARTICLE') <> 'ARP' then TOBL.Putvalue('GL_TYPEPRESENT', Affichage)
                                              else TOBL.Putvalue('GL_TYPEPRESENT', 0);
  StockeInfoTypeLigne (TOBL,valeurs);
  NumeroteLigneOuv (TOBLN,TOBL,1,1,0,0,0);
end;

function TimportLigneFacture.FindFather(TOBParag: TOB; Parag: string): TOB;
var S1,S2,S3,S4,S5 : string;
    CodeParag,S : string;
    II : integer;
begin
  result := nil;
  S := Parag;
  II :=  Pos ('.',S);
  if II <= 0 then exit;
  //
  S1 := ''; S2 := ''; S3 := ''; S4 := ''; S5 := '';
  S1 := READTOKENPipe(S,'.');
  if S <> '' then S2 := READTOKENPipe(S,'.');
  if S <> '' then S3 := READTOKENPipe(S,'.');
  if S <> '' then S4 := READTOKENPipe(S,'.');
  if S <> '' then S5 := READTOKENPipe(S,'.');
  //
  if S5 <> '' then CodeParag := S1+'.'+S2+'.'+S3+'.'+S4+'.'
  else if S4 <> '' then CodeParag := S1+'.'+S2+'.'+S3+'..'
  else if S3 <> '' then CodeParag := S1+'.'+S2+'...'
  else CodeParag := S1+'....';

  result := TOBPARAG.findFirst(['CODEPARAG'],[CodeParag],true);
end;

function TimportLigneFacture.FindLastparag(TOBParag: TOB): TOB;
var TOBP,TOBL : TOB;
    II : integer;
begin
  if TOBPARAG.detail.count = 0 then
  begin
    result := TOBPARAG;
    exit;
  end;
  TOBP := TOBPARAG.detail[TOBPARAG.detail.count-1];
  result := FindLastparag(TOBP);
end;

procedure TimportLigneFacture.AjouteFinParag(TOBPARAG,TOBPiece,TOBTiers,TOBAffaire: TOB;var FlignePgi: integer);
var TOBP,TOBL : TOB;
    II : integer;
begin
  if TOBPARAG.detail.count = 0 then exit;
  for II := 0 to TOBPARAG.detail.count -1 do
  begin
    TOBP := TOBPARAG.detail[II];
    if TOBP.detail.count > 0 then
    begin
      AjouteFinParag(TOBP,TOBPiece,TOBTiers,TOBAffaire,FlignePgi);
    end;
    TOBL := Newtobligne (TOBPiece,fLignePgi);
    InitialiseTOBLigne(TOBPiece,TOBTiers,TOBAffaire,flignePgi,false) ;
    TOBL.PutValue('GL_TYPELIGNE', 'TP'+InttoStr(TOBP.getInteger('NIVEAU')));
    TOBL.putValue('GL_NIVEAUIMBRIC',InttoStr(TOBP.getInteger('NIVEAU')));
    TOBL.putValue('GL_LIBELLE',copy('TOTAL '+TOBP.getString('LIBELLE'),1,70));
    inc(fLignePgi);
  end;
  TOBPARAG.ClearDetail;
end;

function TimportLigneFacture.AddNewParag(TOBPARAG : TOb;  Parag,Designation : string): TOB;
var S,CodeParag,S1,S2,S3,S4,S5 : string;
    Niv : integer;
begin
  S := Parag;
  if Pos ('.',S) <= 0 then
  begin
    CodeParag := S+'....';
    Niv := 1;
  end else
  begin
    //
    S1 := ''; S2 := ''; S3 := ''; S4 := ''; S5 := '';
    S1 := READTOKENPipe(S,'.');
    if S <> '' then S2 := READTOKENPipe(S,'.');
    if S <> '' then S3 := READTOKENPipe(S,'.');
    if S <> '' then S4 := READTOKENPipe(S,'.');
    if S <> '' then S5 := READTOKENPipe(S,'.');
    //
    if S5 <> '' then
    begin
      CodeParag := S1+'.'+S2+'.'+S3+'.'+S4+'.'+S5;
      Niv := 5;
    end else if S4 <> '' then
    begin
      CodeParag := S1+'.'+S2+'.'+S3+'.'+S4+'.';
      Niv := 4;
    end else if S3 <> '' then
    begin
      CodeParag := S1+'.'+S2+'.'+S3+'..';
      Niv := 3;
    end else
    begin
      CodeParag := S1+'.'+S2+'...';
      Niv := 2;
    end;
  end;
  result := TOB.Create ('UN PARAG',TOBPARAG,-1);
  result.AddChampSupValeur('CODEPARAG',CodeParag);
  result.AddChampSupValeur('LIBELLE',Designation);
  result.AddChampSupValeur('NIVEAU',Niv);
end;

procedure TimportLigneFacture.RecalcPaPRPV(TOBL: TOB);
var coeffg : double;
		prefixe : string;
begin
  prefixe := GetPrefixeTable (TOBL);

  if TOBL.Getvalue(prefixe+'_COEFMARG')=0 then TOBL.putvalue(prefixe+'_COEFMARG',1);
  TOBL.PutValue(prefixe+'_DPR',Arrondi(TOBL.getValue(prefixe+'_DPA')*(1+TOBL.Getvalue(prefixe+'_COEFFG')),V_PGI.okdecP));
  TOBL.PutValue(prefixe+'_PUHT',Arrondi(TOBL.getValue(prefixe+'_DPR')*TOBL.Getvalue(prefixe+'_COEFMARG'),V_PGI.okdecP));
  TOBL.putValue(prefixe+'_PUHTDEV',Pivottodevise(TOBL.getValue(prefixe+'_PUHT'),DEV.Taux,DEV.Quotite,DEV.Decimale));
end;

procedure TimportLigneFacture.TraitePALigneCourante(TOBL: TOB; PA: double);
var TypeArticle : string;
    EnHt : boolean;
begin
	TypeArticle := TOBL.getValue('GL_TYPEARTICLE');
  EnHt := (TOBPiece.getValue('GP_FACTUREHT')='X');
  if (TypeArticle = 'OUV') or (TypeArticle = 'ARP') then
  begin
  	TraitePrixOuvrage(TOBPiece,TOBL,TOBBases,TOBBasesL,TOBOuvrage, EnHT, PA,0,DEV,true);
	  CalculMontantsDocOuv (TOBpiece,TOBL,TOBOuvrage,(TOBL.GetValue('GL_BLOQUETARIF')='-'));
    TOBL.PutValue('GL_DPA',PA);
  end else
  begin
    TOBL.PutValue('GL_DPA',PA);
    RecalcPaPRPV (TOBL);
  end;
end;

end.

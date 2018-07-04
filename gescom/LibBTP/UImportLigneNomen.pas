unit UImportLigneNomen;

interface
uses sysutils,classes,windows,messages,controls,forms,hmsgbox,stdCtrls,clipbrd,nomenUtil,
     HCtrls,SaisUtil,HEnt1,Ent1,EntGC,UtilPGI,UTOB,HTB97,FactUtil,FactComm,Menus,
     FactTOB, FactPiece, FactArticle,
		 UtilXlsBTP,EtudesExt,ed_tools,variants,
{$IFDEF EAGLCLIENT}
  maineagl,
{$ELSE}
  Doc_Parser,DBCtrls, Db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} fe_main,UserChg,AglIsoflex,
{$ENDIF}

{$IFDEF BTP}
 BTPUtil, CalcOLEGenericBTP,
{$ENDIF}
	Choix,
 ParamSoc,
 UtilSuggestion,
 UtilArticle,UentCommun;

type

	TimportLigneNomen = class (Tobject)
  	private
      uForm : TForm;
      fdocExcel : string;
      fWinExcel: OleVariant;
      fLignePGI : integer;
      LigneIni : integer;
      WorkBook,WorkSheet : Variant;
    	procedure reajustePv(TOBL: TOB; Newvaleur: double);
    procedure SetForm(const Value: TForm);
    function VerifNomen(TOBL, TobArt: TOB; Arow: integer; var retour,
      Libelle: string): boolean;
    function ChargeSousDetail(CodeNomen: string; TobOuvrage: TOB; Niv,
      OrdreCompo: integer; var Tresult: T_valeurs): boolean;
    procedure AffecteLibelle(TobLN: TOB; stCol: string);
    procedure RenseigneTOBOuv(TOBOuv, TobArticle: TOB; libelle: string;
      AvecPv: boolean);
    public
      TOBPiece,TOBArticles,TOBART,TOBL,TOBLIG,TOBOUV,TOBInfo : TOB;
      DEV : RDevise;
      GestionHt : boolean;
      InfoNiv : integer;
      QteDuDetail : double;
    	constructor create;
      destructor destroy; override;
      procedure execute;
      property documentPGI :TForm write SetForm;
      property documentExcel : string read fdocExcel write fdocExcel ;
      property LignePGI : integer read fLignePGI write fLignePGI ;
    	function findTobArt(ArticleDivers: string): TOB;
  end;
implementation
uses factspec,factureBTP,LigNomen,factOuvrage,FactDomaines,UCotraitance,UspecifPOC;

{ TimportLigneNomen }

constructor TimportLigneNomen.create;
begin
  fWinExcel := unassigned;
end;

destructor TimportLigneNomen.destroy;
begin
  if not VarIsEmpty(fWinExcel) then fWinExcel.Quit;
  fWinExcel := unassigned;
  inherited;
end;

procedure TimportLigneNomen.execute;
var WinNew,TrouvOk : boolean;
		LigneXls,Vide,Error,Integre,Indice : integer;
    CodeArticle,Qte,Designation,Pu,ArticleDivers,PR : string;
    CodeOuv,LibelleOuv : string;
    Px : double;
    Rem : double;
    Tresult : T_valeurs;
begin
	if fdocExcel = '' then exit;
  //
  if TOBOUV.detail[TOBOUV.detail.count-1].getValue('BLO_ARTICLE')='' then TOBOUV.detail[TOBOUV.detail.count-1].free;
  fLignePGi := TOBOUv.detail.count +1;
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
  WorkSheet := SelectSheet (WorkBook,'Feuil1');
  LigneXls := 0;
  vide := 1;
  Integre := 0;
  Error := 0;
  repeat
  	if Vide > 20 then break;
  	inc(LigneXls);
		CodeArticle := trim(GetExcelFormated (WorkSheet,LigneXls,2));
    designation := trimright(trimleft(GetExcelText (WorkSheet,LigneXls,3)));
		Qte := GetExcelText (WorkSheet,LigneXls,1);
    if pos(',',Qte) > 0 then StringReplace (Qte,',','.',[rfReplaceAll]);
		Pu := GetExcelText (WorkSheet,LigneXls,4);
    if pos(',',pu) > 0 then StringReplace (pu,',','.',[rfReplaceAll]);
		PR := GetExcelText (WorkSheet,LigneXls,5);
    if pos(',',pR) > 0 then StringReplace (pR,',','.',[rfReplaceAll]);

    // pour éliminer ligne entête :
    if (CodeArticle<>'') and (qte<>'') and (designation<>'') and (pu<>'') and (valeur(Qte)=0) then continue;

    if (CodeArticle='') or (qte='') or (valeur(Qte)=0) then
    begin
      inc(vide);
      if (CodeArticle<>'') or (qte<>'')  or (designation<>'') or (pu<>'') then inc(error);
      continue;
    end;
    vide := 0;
    TOBART := findTobArt(CodeArticleUnique2(CodeArticle, ''));
    if (TOBART = nil) and (ArticleDivers <> '') then
    begin
    	TOBART := findTobArt(ArticleDivers);
    end;
    if TOBART= nil then BEGIN inc(Error); continue; END;
    //
    TOBL := TOB.Create('LIGNEOUV',TOBOUV,flignePgi-1);
    InsertionChampSupOuv (TOBL,false);
    TOBL.PutValue('BLO_NUMORDRE',fLignePgi);
    //
    TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
    TOBL.putValue('BLO_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));
    //
    TFligNomen(uForm).SetNumerotation(TOBL,TOBOUV,InfoNiv,fLignePgi);
    //
    if Pos(TOBArt.getValue('GA_TYPEARTICLE'),'OUV;ARP')>0 then
    begin
    	Trouvok := VerifNomen(TOBL,TobArt,fLignePgi,CodeOuv,LibelleOuv);
      if not TrouvOk then
      begin
        inc(error);
        Continue;
      end;
      TOBL.PutValue('BLO_COMPOSE',CodeOuv);
      TOBL.PutValue('BLO_LIBELLE',LibelleOuv);
    end else
    begin
      TOBL.PutValue('BLO_LIBELLE',TOBART.getValue('GA_LIBELLE'));
    end;
    TOBL.PutValue('BLO_CODEARTICLE',trim(TOBART.getValue('GA_CODEARTICLE')));
    TOBL.PutValue('BLO_NOMENCLATURE',TOBART.getValue('GA_CODEARTICLE'));
    TOBL.PutValue('BLO_REFARTSAISIE',TOBART.getValue('GA_CODEARTICLE'));
    TOBL.PutValue('BLO_ARTICLE',TOBART.getValue('GA_ARTICLE'));
    TOBL.PutValue('BLO_NIVEAU',InfoNiv);
    TFligNomen(uForm).SetTOBLigne (TOBL,TOBART,TOBInfo);
    TobL.putvalue('BLO_QTEDUDETAIL',QteduDetail);
    RenseigneTOBOuv (TOBL,TOBART,TOBL.GetValue('BLO_LIBELLE'),true);
    if TOBL.getValue('BLO_COMPOSE')<>'' then
    begin
    	// on charge le détail de l'ouvrage
      if not ChargeSousDetail (TobL.GetValue('BLO_COMPOSE'), TOBL ,
      			 									 TobL.GetValue('BLO_NIVEAU')+1,
                               TOBL.GetValue('BLO_NUMORDRE'),Tresult) then
      begin
      	TOBL.InitValeurs;
        inc(error);
        Continue;
      end;
      //
      TFligNomen(uForm).RenumeroteLesFilles (TOBL,InfoNiv);
      TOBL.Putvalue ('BLO_DPA',Tresult[0]);
      TOBL.Putvalue ('BLO_DPR',Tresult[1]);
      TOBL.Putvalue ('BLO_PUHT',Tresult[2]);
      TOBL.Putvalue ('BLO_PUHTBASE',Tresult[2]);
      TOBL.Putvalue ('BLO_PUHTDEV',pivottodevise(Tresult[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
      TOBL.Putvalue ('BLO_PUTTC',Tresult[3]);
      TOBL.Putvalue ('BLO_PUTTCDEV',pivottodevise(Tresult[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
      TOBL.Putvalue ('BLO_PMAP',Tresult[6]);
      TOBL.PutValue('ANCPV', TOBL.getValue('BLO_PUHTDEV'));
      TOBL.PutValue('ANCPA', TOBL.getValue('BLO_DPA'));
      TOBL.PutValue('ANCPR', TOBL.getValue('BLO_DPR'));
      TOBL.Putvalue ('BLO_PMRP',Tresult[7]);
      TOBL.Putvalue ('BLO_QTEDUDETAIL',QteduDetail);
      TOBL.putValue('BLO_TPSUNITAIRE',TResult[9]);
      TOBL.putValue('GA_HEURE',Tresult[9]);
      TOBL.putValue('GA_PAHT',Tresult[0]);
      TOBL.PutValue('BLO_TPSUNITAIRE',Tresult[9]);
      TOBL.PutValue('GA_HEURE',Tresult[9]);
      TOBL.PutValue('BLO_MONTANTFG',Arrondi(Tresult[13]*TOBL.getValue('BLO_QTEFACT'),V_PGI.OkdecP));
      TOBL.PutValue('BLO_MONTANTFC',Arrondi(Tresult[15]*TOBL.getValue('BLO_QTEFACT'),V_PGI.OkdecP));
      TOBL.PutValue('BLO_MONTANTFR',Arrondi(Tresult[14]*TOBL.getValue('BLO_QTEFACT'),V_PGI.OkdecP));
      //
    end;

    if (designation<>'') then
    begin
      TOBL.putValue('BLO_LIBELLE',copy(designation,1,70));
    end;
    TOBL.PutValue('BLO_QTEFACT', Qte);

    if valeur(Pu) <> 0 then
    begin
      reajustePv(TOBL,Valeur(pu));
    end;
    //
    CalculeLigneAcOuvCumul (TOBL);
    CalculMontantHtDevLigOuv (TOBL,DEV);
    GetValoDetail (TOBL);
    //
    inc(fLignePGI);
    inc(Integre);
  until LigneXls > 5000;
  //
	if fLignePGi <> LigneINi then
  begin
  	//
    AffecteLibelle(TOBOuv,'BLO_CODEARTICLE');

    //
    PgiInfo (format('%d enregistrements intégrés - %d erreurs rencontrées',[integre,error]));
  end;
end;

function TimportLigneNomen.findTobArt (ArticleDivers : string) : TOB;
var TOBA : TOB;
    QQ : Tquery;
begin
  result := nil;
  TOBA := TOBArticles.FindFirst (['GA_ARTICLE'],[ArticleDivers],true);
  if TOBA = nil then
  begin
    QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE FROM ARTICLE A '+
    							 'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
    							 'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+ArticleDivers+'"',true);
    if not QQ.eof then
    begin
      TOBA := TOB.create ('ARTICLE',TOBArticles,-1);
      TOBA.SelectDB ('',QQ);
      InitChampsSupArticle (TOBA);
      result := TOBA;
    end;
    ferme (QQ);
  end else result := TOBA;
end;


procedure TimportLigneNomen.reajustePv(TOBL:TOB; NewValeur : double);
var valeurAnc,ValeurPR : double;
begin
	Valeuranc := TOBL.GetValue('ANCPV');
  if (NewValeur <> valeuranc) then
  begin
    if (newValeur <> valeuranc) and (TOBL.Detail.Count > 0) then
        ReajusteMontantOuvrage (TOBART,TOBpiece,TOBLig,TOBL,valeuranc,ValeurPr,newValeur,DEV,GestionHt);
    if GestionHt then
    BEGIN
      TOBL.PutValue('BLO_PUHTDEV', NewValeur);
      TOBL.PutValue('BLO_PUHT', devisetopivotEx(NewValeur,DEV.taux,DEV.quotite,V_PGI.OkdecP));
      CalculeLigneHTOuv  (TOBL,TOBPiece,DEV);
    END ELSE
    BEGIN
      TOBL.PutValue('BLO_PUTTCDEV', NewValeur);
      CalculeLigneTTCOuv  (TOBL,TOBPiece,DEV);
    END;
  end;
  TOBL.PutValue('ANCPV', newValeur);
end;

procedure TimportLigneNomen.SetForm(const Value: TForm);
begin
	uForm := TFligNomen(value);
end;

function  TimportLigneNomen.VerifNomen(TOBL,TobArt : TOB ; Arow : integer; Var retour : string; var Libelle : string ) : boolean;
var Q : TQuery;
    st1 : string;
    i_ind1 : integer;
    TypeArticle,Code : string;
begin
result := true;
st1 := '';
i_ind1 := 0;
TYpeArticle := TOBArt.GetValue('GA_TYPEARTICLE');
Code := TOBArt.GetValue('GA_CODEARTICLE');
if (TypeArticle <> 'OUV') and (TypeArticle <> 'ARP') then
   BEGIN
   exit;
   END else TOBL.PutValue('SOUSNOMEN','X');

Q := OpenSQL('Select Count(GNE_NOMENCLATURE) from NOMENENT Where GNE_ARTICLE="' +
             TOBArt.GetValue('GA_ARTICLE') + '"', True);
if not(Q.Eof) then i_ind1 := Q.Fields[0].AsInteger;
Ferme(Q);

if i_ind1 > 1 then
   BEGIN
   st1:=Choisir('Choix d''une décomposition pour l''ouvrage '+Code,'NOMENENT','GNE_LIBELLE',
          'GNE_NOMENCLATURE','GNE_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'"','');
   if st1 = '' then result := false;
   END else if i_ind1 = 1 then
   BEGIN
   Q := OpenSQL('Select GNE_NOMENCLATURE from NOMENENT Where GNE_ARTICLE="' +
                TOBArt.GetValue('GA_ARTICLE') + '"', True);
   st1 := Q.FindField('GNE_NOMENCLATURE').AsString;
   Ferme(Q);
   end;
if st1 <> '' then // recupération du libellé
   begin
   Q := OpenSQL('Select GNE_LIBELLE from NOMENENT Where GNE_NOMENCLATURE="' + st1 + '"', True);
   Libelle := Q.FindField('GNE_Libelle').AsString;
   Ferme(Q);
   end;
Retour := st1;
end;

function TimportLigneNomen.ChargeSousDetail(CodeNomen : string ; TobOuvrage : TOB ; Niv,OrdreCompo : integer ;var  Tresult: T_valeurs):boolean;
var
    i_ind1, i_ind2,IndPou : integer;
    Select : string;
    Q,QQ : TQuery;
    TOBDetOuv, TobLigne, TobFille : TOB;
    TobArt : TOB;
    // Modif BTP
    TOBLigNom : TOB;
    QteDudetail,QTe,QTedupv,MontantInterm,MontantPr : Double;
    Valloc : T_Valeurs;
    TOBPiece : TOB;
begin
if niv > 5 then
  BEGIN
  PGIBox ('Nombre de niveaux trop important : 5 Maximum','Ouvrages');
  Result := false;
  Exit;
  END;
result := true;
QteDuDetail := 1;
indpou := -1;
//  CodeNomen obligatoire
if CodeNomen <> '' then
   begin
   InitTableau (Tresult);
   TobOuvrage.PutValue('SOUSNOMEN','X');   //En cas de récursivité
   TobLigne := TOB.Create('NOMENLIG', nil, -1);
   //  Chargement du niveau courant de nomenclature
   Select := 'Select NOMENLIG.*,GNE_QTEDUDETAIL,GNE_LIBELLE from NOMENLIG LEFT OUTER JOIN NOMENENT ON GNE_NOMENCLATURE = GNL_NOMENCLATURE where GNL_NOMENCLATURE = "' + CodeNomen + '"';
   Q := OpenSQL(Select, true);
   if not Q.EOF then
      begin
      TobLigne.LoadDetailDB('DETAILNOMEN', '', '', Q, False);
      ferme (Q);
      for i_ind1:=0 to TobLigne.Detail.count -1 do
          begin
          TOBLIGNom := TOBLIgne.detail[i_ind1];
          QteDuDetail := TOBLigNom.getValue('GNE_QTEDUDETAIL');
          if QteDuDetail = 0 then QteDuDetail := 1;
          TOBDetOuv:=TOB.Create('LIGNEOUV',TobOuvrage,-1);
          InsertionChampSupOuv (TOBDetOuv,false);
          TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
    			TOBDetOuv.putValue('BLO_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));

          // prise en compte document d'origine
          TOBDetOuv.PutValue('BLO_NATUREPIECEG',TOBLIG.getValue('GL_NATUREPIECEG'));
          TOBDetOuv.PutValue('BLO_SOUCHE',TOBLig.getValue('GL_SOUCHE'));
          TOBDetOuv.PutValue('BLO_NUMERO',TOBLig.getValue('GL_NUMERO'));
          TOBDetOuv.PutValue('BLO_INDICEG',TOBLig.getValue('GL_INDICEG'));
          // ------
          TOBDetOuv.PutValue('BLO_COMPOSE',TobOuvrage.GetValue('BLO_ARTICLE'));
          TOBDetOuv.PutValue('BLO_NIVEAU',Niv);
          TOBDetOuv.PutValue('BLO_ORDRECOMPO',OrdreCompo);

          TOBDetOuv.PutValue('BLO_NUMORDRE',i_ind1+1);
          TOBDetOuv.PutValue('CODENOMEN',TobLigNom.GetValue('GNL_SOUSNOMEN'));
          TOBDETOuv.putvalue('BLO_QTEDUDETAIL',TOBLigNom.getvalue('GNE_QTEDUDETAIL'));
          TOBDetOuv.PutValue('BLO_ARTICLE', TOBLigNom.GetValue('GNL_ARTICLE'));

//
          QQ := OPenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                          'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                          'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE GA_ARTICLE ="'+
                          TobLigNom.GetValue('GNL_ARTICLE')+'"',True);
          TOBArt:= TOB.create ('ARTICLE',nil,-1);
          TOBArt.selectDB ('',QQ);
          ferme (QQ);
          InitChampsSupArticle (TOBART);
//
          TOBDetOuv.PutValue('BLO_QTEFACT', TOBLigNom.GetValue('GNL_QTE'));
          TFligNomen(uForm).SetTOBLigne(TOBDetOuv,TobArt,TOBInfo);
          TOBDetOuv.PutValue('BLO_CODEARTICLE',TobLigNom.GetValue('GNL_CODEARTICLE'));
          // Modif BTP
          TOBDetOuv.PutValue('BLO_REFARTSAISIE',TobLigNom.GetValue('GNL_CODEARTICLE'));
          // ---
          TOBDetOuv.PutValue('BLO_LIBELLE',TOBLigNom.GetValue('GNL_LIBELLE'));
    			RenseigneTOBOuv (TOBDetOuv,TOBART,TOBDetOuv.GetValue('BLO_LIBELLE'),true);
          if TobDetOuv.GetValue('CODENOMEN') <> '' then
             begin
             if not ChargeSousDetail(TobDetOuv.GetValue('CODENOMEN'), TobDetOuv , Niv+1,TobDetOuv.GetValue('BLO_NUMORDRE'),valloc) then
                begin
                TOBArt.free;
                TOBDetOuv.free;
                TOBligne.free;
                Result := false;
                Exit;
                end;
             TOBDetOuv.Putvalue ('BLO_DPA',valloc[0]);
             TOBDetOuv.Putvalue ('BLO_DPR',valloc[1]);
             TOBDetOuv.Putvalue ('BLO_PUHT',Valloc[2]);
             TOBDetOuv.Putvalue ('BLO_PUHTBASE',Valloc[2]);
             TOBDetOuv.Putvalue ('BLO_PUHTDEV',pivottodevise(Valloc[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
             TOBDetOuv.Putvalue ('BLO_PUTTC',Valloc[3]);
             TOBDetOuv.Putvalue ('BLO_PUTTCDEV',pivottodevise(Valloc[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
             TOBDetOuv.Putvalue ('BLO_PMAP',Valloc[6]);
             TOBDetOuv.Putvalue ('BLO_PMRP',VAlloc[7]);
             TOBDetOuv.Putvalue ('BLO_TPSUNITAIRE',VAlloc[9]);
            TOBDetOuv.PutValue('BLO_MONTANTFG',Arrondi(ValLoc[13]*TOBDetOuv.getValue('BLO_QTEFACT'),V_PGI.OkdecP));
            TOBDetOuv.PutValue('BLO_MONTANTFC',Arrondi(ValLoc[15]*TOBDetOuv.getValue('BLO_QTEFACT'),V_PGI.OkdecP));
            TOBDetOuv.PutValue('BLO_MONTANTFR',Arrondi(ValLoc[14]*TOBDetOuv.getValue('BLO_QTEFACT'),V_PGI.OkdecP));
            CalculeLigneAcOuvCumul (TOBDETOUV);
             CalculMontantHtDevLigOuv (TOBDetOuv,DEV);
             end;
          TOBDetOuv.PutValue('ANCPA',TOBDetOuv.Getvalue('BLO_DPA') );
          TOBDetOuv.PutValue('ANCPR',TOBDetOuv.Getvalue('BLO_DPR') );
          TOBDetOuv.PutValue('ANCPV',TOBDetOuv.Getvalue('BLO_PUHTDEV') );
          if TOBLigNom.GetValue('GNL_PRIXFIXE') <> 0 then
             begin
             TOBDetOuv.PutValue('BLO_PUHT', TOBLigNom.GetValue('GNL_PRIXFIXE'));
             TOBDetOuv.PutValue('BLO_PUHTDEV', pivottodevise(TOBLigNom.GetValue('GNL_PRIXFIXE'),DEV.Taux,DEV.quotite,V_PGI.OkdecP));
             end;
          Qte := TOBDetOuv.getValue('BLO_QTEFACT');
          QteDuPv := TOBDetOuv.getValue('BLO_PRIXPOURQTE'); if QteDuPv = 0 then QteDuPv := 1;
          CalculeLigneAcOuv (TOBDETOUV,DEV,true,TOBART);
          CalculMontantHtDevLigOuv (TOBDetOuv,DEV);
          TobArt.free;
          if TobDetOuv.Getvalue('BLO_TYPEARTICLE') <> 'POU' then
             begin
             Tresult[0] := Tresult [0] + arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_DPA')),V_PGI.okdecP);
             Tresult[1] := Tresult [1] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_DPR')),V_PGI.okdecP);
             Tresult[2] := Tresult [2] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PUHT')),V_PGI.okdecP);
             Tresult[3] := Tresult [3] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PUTTC')),V_PGI.okdecP);
             Tresult[6] := Tresult [6] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PMAP')),V_PGI.okdecP);
             Tresult[7] := Tresult [7] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PMRP')),V_PGI.okdecP);
             Tresult[9] := Tresult [9] +(Qte * TOBDetOuv.GetValue ('BLO_TPSUNITAIRE'));
             Tresult[13] := Tresult [13] +TOBDetOuv.GetValue('BLO_MONTANTFG');
             Tresult[14] := Tresult [14] +TOBDetOuv.GetValue('BLO_MONTANTFR');
             Tresult[15] := Tresult [15] +TOBDetOuv.GetValue('BLO_MONTANTFC');
             end else
             begin
             indpou := 1;
             end;
          end;
      end else Ferme(Q);
   TOBLigne.free;
   FormatageTableau (Tresult,V_PGI.OkdecP);
   // calcule des eventuels pourcentages
   if indpou >= 0 then
      begin
      for I_ind1 := 0 to TOBOuvrage.detail.count -1 do
          begin
          TOBLIGNom := TOBOuvrage.detail[i_ind1];
          if TobLigNom.Getvalue('BLO_TYPEARTICLE') <> 'POU' then continue;
          TOBLIGNom.Putvalue ('BLO_DPA',Tresult[0]);
          TOBLIGNom.Putvalue ('BLO_DPR',Tresult[1]);
          TOBLIGNom.Putvalue ('BLO_PUHT',Tresult[2]);
          TOBLIGNom.Putvalue ('BLO_PUHTBASE',Tresult[2]);
          TOBLIGNom.Putvalue ('BLO_PUHTDEV',pivottodevise(Tresult[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
          TOBLIGNom.Putvalue ('BLO_PUTTC',Tresult[3]);
          TOBLIGNom.Putvalue ('BLO_PUTTCDEV',pivottodevise(Tresult[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
          TOBLIGNom.Putvalue ('BLO_PMAP',Tresult[6]);
          TOBLIGNom.Putvalue ('BLO_PMRP',Tresult[7]);
          Qte := TOBDetOuv.getValue('BLO_QTEFACT');
          CalculMontantHtDevLigOuv (TOBLigNom,DEV);
          QteDuPv := TOBDetOuv.getValue('BLO_PRIXPOURQTE'); if QteDuPv = 0 then QteDuPv := 1;
          Tresult[0] := Tresult [0] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_DPA')),V_PGI.OkdecP);
          Tresult[1] := Tresult [1] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_DPR')),V_PGI.OkdecP);
          Tresult[2] := Tresult [2] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PUHT')),V_PGI.OkdecP);
          Tresult[3] := Tresult [3] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PUTTC')),V_PGI.OkdecP);
          Tresult[6] := Tresult [6] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PMAP')),V_PGI.OkdecP);
          Tresult[7] := Tresult [7] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PMRP')),V_PGI.OkdecP);
          FormatageTableau (Tresult,V_PGI.OkdecP );
          end;
      end;
   for i_ind2 := 0 to TobOuvrage.Detail.Count - 1 do
       begin
       TobFille := TobOuvrage.Detail[i_ind2];
       if (TobFille.getvalue('ANCPV') <> TOBFille.getValue('BLO_PUHTDEV')) and
          (TobFille.detail.count > 0) then
          begin
          MontantInterm := TOBfille.getValue('BLO_PUHTDEV');
          ReajusteMontantOuvrage (nil,TOBPiece,TOBLig,TOBFille,TOBFille.getValue('ANCPV'),MontantPr,MontantInterm,DEV,GestionHt);
          TOBFille.putvalue('BLO_PUHTDEV',MontantInterm);
          TOBDetOuv.PutValue('ANCPV',TOBDetOuv.Getvalue('BLO_PUHTDEV') );
          CalculMontantHtDevLigOuv (TOBDetOuv,DEV);
          end;
       end;
   GestionDetailPrixPose (TobOuvrage);
   end;
   Tresult := CalculSurTableau ('/',Tresult,QteDudetail);
   FormatageTableau (Tresult,V_PGI.OkdecP);
end;

procedure TimportLigneNomen.AffecteLibelle (TobLN : TOB; stCol : string);
var TOBL : TOB ;
    i_ind1 : integer ;
    st1, st2, st3, st4 : string;
begin
for i_ind1 := 0 to TobLN.Detail.Count - 1 do
    begin
    if i_ind1 = 0 then
       begin
       if not TOBLN.FieldExists ('LIBELLE') then TobLN.Detail[i_ind1].AddChampSup('LIBELLE', True);
       end;
    st2 := VarAsType(TobLN.Detail[i_ind1].GetValue(stCol), varString);
    // Modif BTP
    { avant
    st4 := VarAsType(TobLN.Detail[i_ind1].GetValue('GLN_CODEARTICLE'), varString);
    st3 := VarAsType(TobLN.Detail[i_ind1].GetValue('GLN_QTE'), varString);
    }
    st4 := VarAsType(TobLN.Detail[i_ind1].GetValue('BLO_CODEARTICLE'), varString);
    st3 := VarAsType(TobLN.Detail[i_ind1].GetValue('BLO_QTEFACT'), varString);
    if st4 <> '' then
        st1 := st4 + ' (' + StrS0(Valeur(st3)) + ')'
    else
        st1 := st2 + ' (' + StrS0(Valeur(st3)) + ')';
    TobLN.Detail[i_ind1].PutValue('LIBELLE', st1);
    if TobLN.Detail[i_ind1].Detail.Count > 0 then
        begin
        TOBL := TobLN.Detail[i_ind1];
        AffecteLibelle(TOBL, stCol);
        end;
    end;
end;

procedure TimportLigneNomen.RenseigneTOBOuv(TOBOuv, TobArticle : TOB; libelle : string;AvecPv:boolean);
var TOBPiece : TOB;
begin
  If TOBInfo <> nil then
  begin
    // Modif BTP
    TOBOuv.PutValue('BLO_NATUREPIECEG',TOBInfo.GetValue('BLO_NATUREPIECEG'));
    TOBOuv.PutValue('BLO_SOUCHE',TOBInfo.GetValue('BLO_SOUCHE'));
    TOBOuv.PutValue('BLO_NUMERO',TOBInfo.GetValue('BLO_NUMERO'));
    TOBOuv.PutValue('BLO_INDICEG',TOBInfo.GetValue('BLO_INDICEG'));
    TOBOuv.PutValue('BLO_NUMLIGNE',TOBInfo.GetValue('BLO_NUMLIGNE'));
  end;
  CopieOuvFromLigne (TOBOuv,TOBLIG);
	if TobArticle <> nil then
  begin
    // Modif BTP
    TOBOuv.PutValue('BLO_TENUESTOCK',TobArticle.GetValue('GA_TENUESTOCK'));
    //TobLN.PutValue('GLN_QUALIFQTEACH',TobArticle.GetValue('GA_QUALIFQTEACH'));
    TOBOuv.PutValue('BLO_QUALIFQTEVTE',TobArticle.GetValue('GA_QUALIFUNITEVTE'));
    TOBOuv.PutValue('BLO_QUALIFQTESTO',TobArticle.GetValue('GA_QUALIFUNITESTO'));
    TOBOuv.PutValue('BLO_PRIXPOURQTE',TobArticle.GetValue('GA_PRIXPOURQTE'));
    TOBOuv.PutValue('BLO_DPA',TobArticle.GetValue('GA_PAHT'));
    TOBOuv.PutValue('BLO_PMAP',TobArticle.GetValue('GA_PMAP'));
    TOBOuv.PutValue('ANCPA',TobOuv.GetValue('BLO_DPA'));
    TOBOuv.PutValue('ANCPR',TobOuv.GetValue('BLO_DPR'));
		CopieOuvFromArt (TOBPiece,TOBOuv,TobArticle,DEV);
    if AvecPv then
    begin
      TOBOuv.PutValue('BLO_PUHT',TobArticle.GetValue('GA_PVHT'));
      TOBOuv.PutValue('BLO_PUHTBASE',TobArticle.GetValue('GA_PVHT'));
      TOBOuv.PutValue('BLO_PUHTDEV',pivottodevise(TobOuv.GetValue('BLO_PUHTBASE'),DEV.Taux,DEV.quotite,V_PGI.OkdecP ));
      TOBOuv.PutValue('ANCPV',TobOuv.GetValue('BLO_PUHTDEV'));
      TOBOuv.PutValue('BLO_PUTTC',TobArticle.GetValue('GA_PVTTC'));
      TOBOuv.PutValue('BLO_PUTTCBASE',TobArticle.GetValue('GA_PVTTC'));
      TOBOuv.PutValue('BLO_PUTTCDEV',pivottodevise(TOBOuv.GetValue('BLO_PUTTCBASE'),DEV.Taux,DEV.quotite,V_PGI.OkdecP ));
    end;
   // Modif BTP
    if Libelle <> '' then TOBOuv.PutValue('BLO_LIBELLE',Libelle);
   	TOBOUv.putValue('BNP_TYPERESSOURCE',TOBArticle.getValue('BNP_TYPERESSOURCE'));
    //
    if (Pos (TOBOUV.getValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne) > 0) and (not IsLigneExternalise(TOBOUV))   then
    begin
      TOBOUV.PutValue('BLO_TPSUNITAIRE',1);
    end else
    begin
      TOBOUV.PutValue('BLO_TPSUNITAIRE',0);
    end;
    //
    if (IsPrestationSt(TOBOUV)) then
    begin
      if TOBPiece.getValue('GP_APPLICFGST')='X' then
      begin
        TOBOuv.putValue('BLO_NONAPPLICFRAIS','-');
      end else
      begin
        TOBOuv.putValue('BLO_NONAPPLICFRAIS','X');
      end;
      if TOBPiece.getValue('GP_APPLICFCST')='X' then
      begin
        TOBOuv.putValue('BLO_NONAPPLICFC','-');
      end else
      begin
        TOBOuv.putValue('BLO_NONAPPLICFC','X');
      end;
    end else
    begin
      TOBOuv.putValue('BLO_NONAPPLICFRAIS','-');
      TOBOuv.putValue('BLO_NONAPPLICFC','-');
    end;
    if TobArticle.getValue('GA_PRIXPASMODIF')<>'X' Then
    begin
      AppliqueCoefDomaineActOuv (TOBOuv);
      if VH_GC.BTCODESPECIF = '001' then ApliqueCoeflignePOC (TOBOuv,DEV);
    end;
    CalculeLigneAcOuv (TOBOUV,DEV,true,TOBARTICLE);
   end else
   begin
   if AvecPv then
      begin
      TOBOuv.PutValue('BLO_PUHT',0);
      TOBOuv.PutValue('BLO_PUHTBASE',0);
      TOBOuv.PutValue('BLO_PUHTDEV',0);
      TOBOuv.PutValue('ANCPV',0);
      TOBOuv.PutValue('BLO_PUTTC',0);
      TOBOuv.PutValue('BLO_PUTTCBASE',0);
      TOBOuv.PutValue('BLO_PUTTCDEV',0);
      TOBOuv.PutValue('BLO_DPA',0);
      TOBOuv.PutValue('BLO_DPR',0);
      TOBOuv.PutValue('BLO_PMAP',0);
      TOBOuv.PutValue('BLO_PMRP',0);
      TOBOuv.PutValue('BLO_QUALIFQTEVTE','');
      TOBOuv.PutValue('BLO_QUALIFQTESTO','');
      end;
   TOBOuv.PutValue('ANCPA',0);
   TOBOuv.PutValue('ANCPR',0);
   CalculeLigneAcOuv (TOBOUV,DEV,True,TOBARTICLE);
   end;
end;

end.

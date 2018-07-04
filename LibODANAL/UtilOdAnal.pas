unit UtilOdAnal;

interface
uses  DB,
      Fe_Main,       // AGLLanceFiche
      {$IFDEF DBXPRESS}uDbxDataSet,{$ELSE}dbtables,{$ENDIF}
      Ed_Tools,
      HMsgBox,      // PgiAsk
      Forms,        // Application
      SysUtils,     // Except
      HEnt1,        // BeginTrans
      Ent1,         // VH^.
      uTob,         // TOB
      HCtrls,
      UentCommun,
      ParamSoc,
      SAISUTIL;


procedure MakeExtourneOdAnal (TOBODANAL : TOB);
procedure SupprimeODAnal  (TOBODANAL : TOB);
function  OkSupprimeAnal (TOBODANAL : TOB) : Boolean;
procedure ValideLesODStocK(TOBPiece,TOBODANALPIECE,TOBCPTASTock : TOB; DEV : RDevise);
procedure ConstitueODAnalConso (TOBODANAL,TOBCONSO : TOB; TypeMvt : string);
function EncodeRefAna(TOBOD : TOB; typemvt : string='C' ) : string;
procedure DecodeRefAna (refAnal : string; var CleRefAna : r_cleAna);

implementation

uses BTPUtil,CalcOLEGenericBTP,UtilPGI,FactTOB,UtilTOBPiece,
		 FactCpta,EntGC,UtilSaisieConso,Controls,Chancel;

function GetLeTauxDevise(CodeD: string; var DateTaux: TDateTime; DateP: TDateTime): Double;
var Taux: Double;
  ChoixT, Taux1: Boolean;
  ii: integer;
begin
  ChoixT := False;
  Taux1 := False;
  Taux := GetTaux(CodeD, DateTaux, DateP);
  //
  // MODIF LM 02/04/2003
  // Saisie d'un document EUR sur un dossier en devise OUT
  // La table de chancellerie stocke le taux par rapport à l'EUR, par par rapport à la devise pivot du dossier.
  //
  //if ((CodeD<>V_PGI.DevisePivot) and (Not EstMonnaieIN(CodeD))) then
  if ((CodeD <> 'EUR') and (not EstMonnaieIN(CodeD))) then
  begin
    if VH_GC.GCAlerteDevise = 'JOU' then ChoixT := (DateTaux <> DateP) else
      if VH_GC.GCAlerteDevise = 'SEM' then ChoixT := (NumSemaine(DateTaux) <> NumSemaine(DateP)) else
      if VH_GC.GCAlerteDevise = 'MOI' then ChoixT := (GetPeriode(DateTaux) <> GetPeriode(DateP));
    if Taux = 1 then
    begin
      Taux1 := True;
      ChoixT := True;
    end;
    if ChoixT then
    begin
      if Taux1 then ii := PgiAsk('ATTENTION : Le taux en cours est de 1. Voulez-vous saisir ce taux dans la table de chancellerie')
               else ii := PgiAsk('Voulez-vous saisir ce taux dans la table de chancellerie ?');
      if ii = mrYes then
      begin
        {$IFNDEF SANSCOMPTA}
        FicheChancel(CodeD, True, DateP, taCreat, (DateP >= V_PGI.DateDebutEuro));
        {$ENDIF}
        Taux := GetTaux(CodeD, DateTaux, DateP);
      end;
    end;
  end;
  Result := Taux;
end;

procedure InitODAnal(TOBA: TOB;Etablissement,Libelle,Axe,Jnl : string; NumPiece : Integer; DateModif : TdateTime; DEV : RDevise; Indice : integer);
begin
  TOBA.putValue('Y_AXE', Axe);
  TOBA.putValue('Y_DATECOMPTABLE', StrToDate(DateToStr(DateModif)));
  TOBA.putValue('Y_NUMEROPIECE', NumPiece);
  TOBA.putValue('Y_NUMVENTIL', Indice);
  TOBA.putValue('Y_NUMLIGNE', 0);
  TOBA.putValue('Y_EXERCICE', QuelExoDT(DateModif));
  TOBA.putValue('Y_TYPEANALYTIQUE', 'X');
  TOBA.putValue('Y_JOURNAL', Jnl);
  TOBA.putValue('Y_LIBELLE', Libelle);
  TOBA.putValue('Y_REFINTERNE', Libelle);
  TOBA.putValue('Y_NATUREPIECE', 'OD');
  TOBA.putValue('Y_QUALIFPIECE', 'N');
  TOBA.putValue('Y_ETAT', '0000000000');
  TOBA.putValue('Y_UTILISATEUR', V_PGI.User);
  TOBA.putValue('Y_DATECREATION', StrToDate(DateToStr(DateModif)));
  TOBA.putValue('Y_DATEMODIF', DateModif);
  TOBA.putValue('Y_SOCIETE', V_PGI.CodeSociete);
  TOBA.putValue('Y_ETABLISSEMENT',Etablissement);
  TOBA.putValue('Y_DEVISE', DEV.code);
  TOBA.putValue('Y_TAUXDEV', DEV.Taux);
  TOBA.putValue('Y_DATETAUXDEV', DEV.DateTaux);
  TOBA.putValue('Y_QUALIFQTE1', '');
  TOBA.putValue('Y_QUALIFQTE2', '');
  TOBA.putValue('Y_QUALIFECRQTE1', '');
  TOBA.putValue('Y_QUALIFECRQTE2', '');
  TOBA.putValue('Y_ECRANOUVEAU', 'N');
  TOBA.putValue('Y_CREERPAR', 'SAI');
  TOBA.putValue('Y_EXPORTE', '---');
  TOBA.putValue('Y_CONFIDENTIEL', '0');
  TOBA.PutValue('Y_PERIODE',GetPeriode(DateModif));
  TOBA.PutValue('Y_SEMAINE',NumSemaine(DateModif));
  TOBA.PutValue('Y_VALIDE','-');
  // V9 CEGID
  TOBA.PutValue('Y_DATPER',iDate1900) ;
  TOBA.PutValue('Y_ENTITY',0) ;
  TOBA.PutValue('Y_REFGUID','') ;
end;

procedure ConstitueOdAnalPiece (TOBpiece,TOBODA, TOBPieceOD : TOB);
begin
  TOBPieceOD.SetString('BOA_NATUREPIECEG',TOBpiece.GetString('GP_NATUREPIECEG'));
  TOBPieceOD.SetString('BOA_SOUCHE',TOBpiece.GetString('GP_SOUCHE'));
  TOBPieceOD.SetString('BOA_NUMERO',TOBpiece.GetString('GP_NUMERO'));
  TOBPieceOD.SetString('BOA_INDICEG',TOBpiece.GetString('GP_INDICEG'));
  TOBPieceOD.SetString('BOA_JOURNAL',TOBODA.GetString('Y_JOURNAL'));
  TOBPieceOD.SetString('BOA_EXERCICE',TOBODA.GetString('Y_EXERCICE'));
  TOBPieceOD.SetString('BOA_ENTITY',TOBODA.GetString('Y_ENTITY'));
  TOBPieceOD.SetString('BOA_DATECOMPTABLE',TOBODA.GetString('Y_DATECOMPTABLE'));
  TOBPieceOD.SetString('BOA_NUMEROPIECE',TOBODA.GetString('Y_NUMEROPIECE'));
  TOBPieceOD.SetString('BOA_DATECOMPTABLE',TOBODA.GetString('Y_DATECOMPTABLE'));
end;


procedure ValideLesODStocK(TOBPiece,TOBODANALPIECE,TOBCPTASTock : TOB; DEV : RDevise);
var TOBEE,TOBV,TOBODAnal,TOBODA,TOBOD,TOBVS,TOBOG,TOBODE : TOB;
		NumCpta,II,JJ : Integer;
    DateModif : TDateTime;
    Jnl,Axe,Libelle,cptStock,NatureG : string;
    MtGlob : Double;
    NumPiece,NumPieceSto : Integer;
begin
  if not GetParamSocSecur('SO_OPTANALSTOCK',false) then Exit; // Si pas de gestion de l'analytique du stock on sort
  TOBODANALPIECE.ClearDetail;
  if TOBCPTASTock.detail.Count = 0 then Exit;
  Jnl :=  GetParamSocSecur('SO_BTJNLANALSTOCK','');
  Axe := GetParamSocSecur('SO_BTAXEANALSTOCK','');
  if (Axe = '') or (jnl = '') then Exit;
  cptStock:=copy(BTGetSectionStock(TOBPiece.GetString('GP_DEPOT')),1,VH^.Cpta[AxeToFb(Axe)].lg);
  if cptStock = '' then Exit;
  NatureG := TOBPiece.GetString('GP_NATUREPIECEG');
  DateModif := NowH;
  TOBODAnal := TOB.Create ('LES OD ANALYTIQUES',nil,-1);
  TOBODA := TOB.Create('UNE VENTIL ANAL',nil,-1);
  TRY
    if NatureG='LBT' then // livraison chantier
    begin
    	Libelle := Copy('Livraison Chantier N° '+ TOBPiece.GetString('GP_NUMERO'),1,35);
    end else if NatureG='BFC' then // retour chantier
    begin
    	Libelle := Copy('Retour de Chantier '+ TOBpiece.GetString('GP_NUMERO'),1,35);
    end;
    for II :=0 to TOBCPTASTock.detail.count -1 do
    begin
      TOBEE := TOBCPTASTock.detail[II]; // cpta général
      if TOBEE.Detail.Count = 0 then Continue;
	    NumCpta:=GetNewNumJal(Jnl,true,DateModif) ;
  		TOBODA.ClearDetail;
      TOBVS := TOBEE.detail[0]; // les ventilations analytique
      if TOBVS.detail.count = 0 then continue;
      TOBOG := TOB.Create('ANALYTIQ',TOBODA,-1);
			InitODAnal (TOBOG,TOBPiece.GetString('GP_ETABLISSEMENT'),Libelle,Axe,Jnl,NumCpta,DateModif,DEV,1);
      TOBOG.SetString('Y_AFFAIRE',BTPCodeAffaireAffiche(TOBpiece.GetString('GP_AFFAIRE')));
      TOBOG.SetString('Y_GENERAL',TOBEE.GetString('GCP_CPTEGENEACH'));
      TOBOG.SetString('Y_SECTION',cptStock);
  		TOBOG.SetString('Y_TYPEMVT', 'AE');
      //
      MtGlob := 0;
      for JJ := 0 to TOBVS.detail.count -1 do
      begin
      	TOBV := TOBVS.detail[JJ];
        //
        //
      	TOBOD := TOB.Create('ANALYTIQ',TOBODA,-1);
				InitODAnal (TOBOD,TOBPiece.GetString('GP_ETABLISSEMENT'),Libelle,Axe,Jnl,NumCpta,DateModif,DEV,JJ+2);
      	TOBOD.SetString('Y_AFFAIRE',BTPCodeAffaireAffiche(TOBpiece.GetString('GP_AFFAIRE')));
      	TOBOD.SetString('Y_GENERAL',TOBEE.GetString('GCP_CPTEGENEACH'));
      	TOBOD.SetString('Y_SECTION',TOBV.GetString('V_SECTION'));
  			TOBOD.SetString('Y_TYPEMVT', 'AL');
        if NatureG = 'LBT' then
        begin
        	MtGlob := MtGlob + TOBV.GetDouble('V_MONTANT');
      		TOBOD.SetDouble('Y_DEBITDEV',TOBV.GetDouble('V_MONTANT'));
      		TOBOD.SetDouble('Y_DEBIT',DeviseToPivotEx(TOBV.GetDouble('V_MONTANT'),DEV.Taux,DEV.Quotite,DEV.Decimale));
        end else if NatureG = 'BFC' then
        begin
        	MtGlob := MtGlob + (TOBV.GetDouble('V_MONTANT')*(-1));
      		TOBOD.SetDouble('Y_CREDITDEV',TOBV.GetDouble('V_MONTANT')*(-1));
      		TOBOD.SetDouble('Y_CREDIT',DeviseToPivotEx(TOBV.GetDouble('V_MONTANT')*(-1),DEV.Taux,DEV.Quotite,DEV.Decimale));
        end;
      end;
      if NatureG = 'LBT' then
      begin
        TOBOG.SetDouble('Y_CREDITDEV',mtGlob);
        TOBOG.SetDouble('Y_CREDIT',DeviseToPivotEx(mtGlob,DEV.Taux,DEV.Quotite,DEV.Decimale));
      end else if NatureG = 'BFC' then
      begin
        TOBOG.SetDouble('Y_DEBITDEV',mtGlob);
        TOBOG.SetDouble('Y_DEBIT',DeviseToPivotEx(mtGlob,DEV.Taux,DEV.Quotite,DEV.Decimale));
      end;
      // ----------
    	NumPieceSto := 0;
      repeat
        // Alimentation de la TOB PIECE OD ANA
        NumPiece := TOBODA.detail[0].GetInteger('Y_NUMEROPIECE');
        if NumPieceSto <> NumPiece then
        begin
          NumPieceSto := NumPiece;
          TOBODE := TOB.Create ('PIECEODANA',TOBODANALPIECE,-1);
          ConstitueOdAnalPiece (TOBPiece,TOBODA.detail[0], TOBODE);
        end;
        //
      	TOBODA.detail[0].ChangeParent(TOBODAnal,-1)
      until TOBODA.detail.count = 0;
      // -------------
    end;
    if TOBODAnal.Detail.count > 0 then
    begin
      if not TOBODAnal.InsertDB(nil) then
      begin
        V_PGI.IOError := oePiece;
        exit;
      end;
      if not TOBODANALPIECE.InsertDB(nil) then
      begin
        V_PGI.IOError := oePiece;
        exit;
      end;
    end;
  FINALLY
    TOBODA.free;
    TOBODAnal.free;
  end;
end;

function OkSupprimeAnal (TOBODANAL : TOB) : Boolean;
var Exercice : string;
		QQ : TQuery;
begin
	result := false;
  if TOBODANAL = nil then Exit;
	exercice := TOBODANAL.Detail[0].GetString('BOA_EXERCICE');
  QQ := OpenSql ('SELECT EX_ETATCPTA FROM EXERCICE WHERE EX_EXERCICE="'+exercice+'"',True,1,'',true);
  if not QQ.Eof then
  begin
    if QQ.FindField('EX_ETATCPTA').AsString = 'OUV' then
    begin
    	Result := true;
    end;
  end;
  ferme (QQ);
end;

procedure SupprimeODAnal  (TOBODANAL : TOB);
var ii : Integer;
	  Sql : string;
    TOBOD : TOB;
begin
  for ii := 0 to TOBODANAL.Detail.Count -1 do
  begin
    TOBOD := TOBODANAL.Detail[II];
    SQl :='DELETE FROM ANALYTIQ WHERE '+
          'Y_EXERCICE="'+TOBOD.GetString('BOA_EXERCICE')+'" AND '+
          'Y_ENTITY='+TOBOD.GetString('BOA_ENTITY')+' AND '+
          'Y_JOURNAL="'+TOBOD.GetString('BOA_JOURNAL')+'" AND '+
          'Y_NUMEROPIECE='+TOBOD.GetString('BOA_NUMEROPIECE')+' AND '+
          'Y_DATECOMPTABLE="'+USDateTime(StrToDate(DateToStr(TOBOD.GetValue('BOA_DATECOMPTABLE'))))+'"';
    ExecuteSQL(Sql);
  end;
end;

procedure GenereExtourneOdAnal (TOBCPANAL : TOB; DateModif : TDateTime);
var TOBEXTOURNE,TOBOD,TOBEX : TOB;
		ii : Integer;
    Libelle,Jnl : string;
    numCpta : Integer;
    DEV : Rdevise;
    MtAtt : double;
begin
  TOBEXTOURNE := TOB.Create('LES EXTOURNES',nil,-1);
  TRY
    for ii := 0 to TOBCPANAL.detail.count -1 do
    begin
      TOBOD := TOBCPANAL.detail[II];
      TOBEX := TOB.Create ('ANALYTIQ',TOBEXTOURNE,-1);
      //
	    NumCpta:=GetNewNumJal(Jnl,true,DateModif) ;
      Libelle := Copy('Annulation - '+TOBOD.GetValue('Y_LIBELLE'),1,35);
      Jnl := TOBOD.GetString('Y_JOURNAL');
      DEV.Code := TOBOD.GetString('Y_DEVISE');
      DEV.TAUX := TOBOD.GetDouble('Y_TAUXDEV');
      DEV.DateTaux := TOBOD.GetDateTime('Y_DATETAUXDEV');
      //
      InitODAnal (TOBEX,TOBOD.GetString('Y_ETABLISSEMENT'),Libelle,TOBOD.GetString('Y_AXE'),
      						Jnl,NumCpta,DateModif,DEV,ii+1);
      //
      MtAtt := TOBOD.GetDouble('Y_CREDIT');
      TOBEX.SetDouble('Y_CREDIT',TOBOD.GetDouble('Y_DEBIT'));
      TOBEX.SetDouble('Y_DEBIT',MtAtt);
      TOBEX.SetAllModifie(true);
    end;
    TOBEXTOURNE.InsertDB(nil);
  FINALLY
    TOBEXTOURNE.free;
  end;
end;

procedure MakeExtourneOdAnal (TOBODANAL : TOB);
var QQ1 : TQuery;
		TOBCPANAL,TOBOD : TOB;
    ii : Integer;
    Sql : string;
    DateModif : TDateTime;
begin
  if TOBODANAL.Detail.Count = 0 then Exit;
  TOBCPANAL := TOB.Create('LES CONTREPARTIES',nil,-1);
  TRY
    for ii := 0 to TOBODANAL.Detail.Count -1 do
    begin
      TOBOD := TOBODANAL.Detail[II];
      SQl :='SELECT * FROM ANALYTIQ WHERE '+
            'Y_EXERCICE="'+TOBOD.GetString('BOA_EXERCICE')+'" AND '+
            'Y_ENTITY='+TOBOD.GetString('BOA_ENTITY')+' AND '+
            'Y_JOURNAL="'+TOBOD.GetString('BOA_JOURNAL')+'" AND '+
            'Y_NUMEROPIECE='+TOBOD.GetString('BOA_NUMEROPIECE')+' AND '+
            'Y_DATECOMPTABLE="'+TOBOD.GetString('BOA_DATECOMPTABLE')+'"';
      QQ1 := OpenSQL(Sql,True,-1,'',true);
      TOBCPANAL.LoadDetailDB('ANALTYTIQ','','',QQ1,true);
      ferme (QQ1);
    end;
    if TOBCPANAL.Detail.count = 0 then exit;
    //-------------------------------
    DateModif := NowH;
    //-------------------------------
    GenereExtourneOdAnal (TOBCPANAL,DateModif);
  finally
    TOBCPANAL.free;
  end;
end;


function GetCptegeneralFromConso(TOBCONSO,TOBTIERS,TOBARTICLES,TOBAFFAIRES,TOBRESSOURCES : TOB; DEV : Rdevise) : string;
var TOBART,TOBTIER,TOBAFFAIRE,TOBRESS,TOBC,TOBXX,TOBXXC : TOB;
    QQ : TQuery;
    Sql,FamArt,Regime,Etab,FamTiers,NatV,FamTaxe,FamAFF,Tiers : string;
begin
  TOBC := TOB.Create ('CODECPTA',nil,-1);
  //
  FillChar(DEV, Sizeof(DEV), #0);
  FamArt := '';
  FamTaxe := '';
  Tiers := '';
  FamAFF := '';
  FamTiers := '';
  Regime := '';
  Etab := '';
  //
  try
    // --- Infos PRESTATION ou marchandise
    TOBART := TOBARTICLES.FindFirst(['GA_ARTICLE'],[TOBCONSO.GetString('BCO_ARTICLE')],true);
    if TOBART = nil then
    begin
      QQ := OpenSql ('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+TOBCONSO.GetString('BCO_ARTICLE')+'"',
                     true,1,'',True);
      if not QQ.Eof then
      begin
      	TOBART := TOB.Create ('ARTICLE',TOBARTICLES,-1);
        TOBART.SelectDB('',QQ);
      end;
      Ferme(QQ);
    end;
    if TOBART <> nil then
    begin
      FamArt := TOBART.GetString('GA_COMPTAARTICLE');
      FamTaxe := TOBART.GetString('GA_FAMILLENIV1');
    end;
    // --- Infos Chantier
    TOBAFFAIRE := TOBAFFAIRES.FindFirst(['AFF_AFFAIRE'],[TOBCONSO.GetString('BCO_AFFAIRE')],true);
    if TOBAFFAIRE = nil then
    begin
      QQ := OpenSql ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+TOBCONSO.GetString('BCO_AFFAIRE')+'"',
                     true,1,'',True);
      if not QQ.Eof then
      begin
  			TOBAFFAire := TOB.Create ('AFFAIRE',TOBAFFAIRES,-1);
        TOBAFFAIRE.SelectDB('',QQ);
      end;
      Ferme(QQ);
    end;
    if TOBAFFAIRE <> nil then
    begin
      Tiers := TOBAFFAIRE.GetString('AFF_TIERS');
      FamAFF := TOBAFFAIRE.GetString('AFF_COMPTAAFFAIRE');
      Etab := TOBAFFAIRE.GetString('AFF_ETABLISSEMENT');
    end;
    // --- Infos TIERS
    TOBTIER := TOBTIERS.FindFirst(['T_TIERS'],[Tiers],true);
    if TOBTIER = nil then
    begin
      QQ := OpenSQL('SELECT * FROM TIERS WHERE T_TIERS="'+Tiers+'" '+
                    'AND T_NATUREAUXI="CLI"',True,1,'',True);
      if not QQ.Eof then
      begin
  			TOBTier := TOB.Create ('TIERS',TOBTIERS,-1);
        TOBTIER.SelectDB('',QQ);
      end;
      ferme (QQ);
    end;
    if TOBTIER <> nil then
    begin
      FamTiers:=TOBTIER.GetString('T_COMPTATIERS');
      Regime := TOBTIER.GetString('T_REGIMETVA');
    end;
    // -------- Info Intervenant (ressource)
    TOBRESS := TOBRESSOURCES.FindFirst(['ARS_RESSOURCE'],[TOBCONSO.GetString('BCO_RESSOURCE')],true);
    if TOBRESS = nil then
    begin
      QQ := OpenSQL('SELECT * FROM RESSOURCE WHERE ARS_RESSOURCE="'+TOBCONSO.GetString('BCO_RESSOURCE')+'" ',True,1,'',True);
      if not QQ.Eof then
      begin
        TOBRESS := TOB.Create ('RESSOURCE',nil,-1);
        TOBRESS.SelectDB('',QQ);
      end;
      ferme (QQ);
    end;
    if TOBRESS <> nil then
    begin
    	if Etab = '' then Etab:=TOBRESS.GetString('ARS_ETABLISSEMENT');
    end;
    // -------------------------------------------------------------------------------------------------
    NatV:='HA' ; // on se positionne sur des comptes de charges (transfert de charges)
    SQL:=FabricSQLCompta(FamArt,FamTiers,FamAff,Etab,Regime,FamTaxe,'CON') ;
    QQ:=OpenSQL(SQL,True,-1, '', True) ;
    if Not QQ.EOF then
    BEGIN
      TOBXX := TOB.Create ('LES VENTIL',nil,-1);
      TOBXX.LoadDetailDB('CODECPTA','','',QQ,false);
      TOBXXC:=FindTOBCode(TOBXX,FamArt,FamTiers,FamAff,Etab,Regime,FamTaxe) ;
      if TOBXXC <> nil then
      begin
        TOBC.Dupliquer(TOBXXC,False,true);
      end;
      TOBXX.Free;
    END else
    begin
      TOBC.SetString('GCP_COMPTAARTICLE',FamArt);
      TOBC.SetString('GCP_COMPTATIERS',FamTiers);
      TOBC.SetString('GCP_COMPTAAFFAIRE',FamAff);
      TOBC.SetString('GCP_REGIMETAXE',Regime);
      TOBC.SetString('GCP_FAMILLETAXE',FamTaxe);
      TOBC.SetString('GCP_ETABLISSEMENT',Etab);
      TOBC.SetString('GCP_VENTEACHAT','ACH');
      TOBC.SetString('GCP_CPTEGENEACH',VH_GC.GCCpteHTACH);
      TOBC.SetString('GCP_CPTEGENEVTE',VH_GC.GCCpteHTVTE);
    end;
    Ferme(QQ) ;
    result := TOBC.GetString('GCP_CPTEGENEACH');
  finally
    TOBC.free;
  end;
end;

procedure DecodeRefAna (refAnal : string; var CleRefAna : r_cleAna);
begin
  FillChar(CleRefAna,SizeOf(CleRefAna),#0);
  if refAnal = '' then Exit;
  CleRefAna.TYPEMVT := READTOKENST(RefAnal);
  CleRefAna.ENTITY := StrToInt(readTokenST (refAnal));
  CleRefAna.JOURNAL := readTokenST (refAnal);
  CleRefAna.EXERCICE  := readTokenST (refAnal);
  CleRefAna.DATECOMPTABLE := UsDateTimeToDateTime(readTokenST (refAnal));
  CleRefAna.NUMEROPIECE := StrToInt(readTokenST (refAnal));
  CleRefAna.NUMLIGNE  := StrToInt(readTokenST (refAnal));
  CleRefAna.AXE  := readTokenST (refAnal);
  CleRefAna.NUMVENTIL  := StrToInt(readTokenST (refAnal));
  CleRefAna.QUALIFPIECE  := readTokenST (refAnal);
end;

function EncodeRefAna(TOBOD : TOB; typemvt : string='C' ) : string;
begin
  Result := TypeMvt+';'+
  					IntToStr(TOBOD.GetInteger('Y_ENTITY'))+';'+
  					TOBOD.GetString('Y_JOURNAL')+';'+
  					TOBOD.GetString('Y_EXERCICE')+';'+
  					USDATETIME(TOBOD.GetDateTime('Y_DATECOMPTABLE'))+';'+
  					IntToStr(TOBOD.GetInteger('Y_NUMEROPIECE'))+';'+
  					IntToStr(TOBOD.GetInteger('Y_NUMLIGNE'))+';'+
  					TOBOD.GetString('Y_AXE')+';'+
  					IntToStr(TOBOD.GetInteger('Y_NUMVENTIL'))+';'+
  					TOBOD.GetString('Y_QUALIFPIECE')+';';
end;

function GetEtablissement (TOBAFFAIRES,TOBCO : TOB) : string;
var Affaire : string;
		QQ : Tquery;
    TOBAFF : TOB;
begin
  Result := '';
  Affaire := TOBCO.GEtSTring('BCO_AFFAIRE');
  TOBAFF := TOBAFFAIRES.FindFirst(['AFF_AFFAIRE'],[Affaire],true);
  if TOBAFF<> nil then
  begin
    Result := TOBAFF.GetString('AFF_ETABLISSEMENT');
  end else
  begin
    QQ := OpenSQL('SELECT ARS_ETABLISSEMENT FROM RESSOURCE WHERE ARS_RESSOURCE="'+TOBCO.GetString('BCO_RESSOURCE')+'" ',True,1,'',True);
    if not QQ.Eof then
    begin
      Result:=QQ.findField('ARS_ETABLISSEMENT').AsString;
    end;
    ferme (QQ);
  end;
  if result = '' then result := GetParamSocSecur('SO_ETABLISDEFAUT','');
end;

procedure ConstitueODAnalConso (TOBODANAL,TOBCONSO : TOB; TypeMvt : string);
var General,Etab,JNL,Libelle : string;
		TOBOD,TOBCO : TOB;
    DEV : Rdevise;
    Axe : string;
    NumCpta,II :integer;
    DateModif : TDateTime;
    Section : string;
    TOBTIERS,TOBARTICLES,TOBAFFAIRES,TOBRESSOURCES : TOB;
    Mtt : double;
begin
  TOBTIERS := TOB.Create('LES TIERS',nil,-1);
  TOBARTICLES := TOB.Create ('LES ARTICLES',nil,-1);
  TOBAFFAIRES := TOB.Create('LES AFFAIRES',nil,-1);
  TOBRESSOURCES := TOB.Create('LES RESSOURCES',nil,-1);
  TRY
    if TOBCONSO.detail.Count = 0 then Exit;
    if not GetParamSocSecur('SO_OPTANALCONSO',false) then Exit; // Si pas de gestion de l'analytique des consommations
    Jnl :=  GetParamSocSecur('SO_BTJNLANALCONSO','');
    Axe := GetParamSocSecur('SO_BTAXEANALCONSO','');
    if (Jnl = '') or (Axe = '') then Exit;
    //
    //
    DateModif := TOBCONSO.detail[0].GetDateTime('BCO_DATEMOUV');
    //
    DEV.Code := GetParamSocSecur('SO_DEVISEPRINC','EUR');
    GetInfosDevise(DEV);
  	DEV.Taux := GetLeTauxDevise(DEV.Code, DEV.DateTaux,DateModif);
    //
    NumCpta:=GetNewNumJal(Jnl,true,DateModif) ;
    //
    for II := 0 to TOBCONSO.detail.count -1 do
    begin
      TOBCO := TOBCONSO.detail[II];
      Mtt := TOBCO.GetDouble('BCO_MONTANTACH');
      //
      Etab := GetEtablissement (TOBAFFAIREs,TOBCO);
      General := GetCptegeneralFromConso(TOBCO,TOBTIERS,TOBARTICLES,TOBAFFAIRES,TOBRESSOURCES,DEV);
      GetSectionConso (Section,Libelle,Axe,TOBCO,TOBTIERS,TOBARTICLES,TOBAFFAIRES,TOBRESSOURCES);
      Libelle := TOBCO.GetString('BCO_LIBELLE');
      //
      TOBOD := TOB.Create('ANALYTIQ',TOBODANAL,-1);
      InitODAnal (TOBOD,Etab,Libelle,Axe,jnl,NumCpta,TOBCO.GetDateTime('BCO_DATEMOUV'),DEV,II+1);
      TOBOD.SetString('Y_AFFAIRE',BTPCodeAffaireAffiche(TOBCO.GetString('BCO_AFFAIRE')));
      TOBOD.SetString('Y_GENERAL',General);
      TOBOD.SetString('Y_SECTION',Section);
      if II = 0 then TOBOD.SetString('Y_TYPEMVT', 'AE')
      					else TOBOD.SetString('Y_TYPEMVT', 'AL');
      if mtt > 0 then
      begin
        TOBOD.SetDouble('Y_DEBITDEV',mtt);
        TOBOD.SetDouble('Y_DEBIT',DeviseToPivotEx(mtt,DEV.Taux,DEV.Quotite,DEV.Decimale));
      end else
      begin
        TOBOD.SetDouble('Y_CREDITDEV',mtt*(-1));
        TOBOD.SetDouble('Y_CREDIT',DeviseToPivotEx(TOBOD.GetDouble('Y_CREDITDEV'),DEV.Taux,DEV.Quotite,DEV.Decimale));
      end;
      TOBCO.SetString('BCO_LIENANAL',EncodeRefAna(TOBOD,typeMvt));
      TOBOD.SetAllModifie(true);
      TOBCO.SetAllModifie(true);
    end;
  FINALLY
    TOBTIERS.Free;
    TOBARTICLES.free;
    TOBAFFAIRES.Free;
    TOBRESSOURCES.Free;
  END;
end;

end.

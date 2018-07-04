unit UDossierVP;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
    {$IFNDEF ERADIO}
    Fe_Main,
    {$ENDIF !ERADIO}
    {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
    HDB,
    Mul,
    UtilGc,
		AGLInit,
		uTOB,
    HEnt1,
    UentCommun,
    HCtrls,
    HMsgBox,
    DialogEx;

procedure TraitementProjetDocument (Cledoc : r_cledoc);

implementation


uses  UtilTOBPiece,
			FactUtil,
      factOuvrage,
      CalcOLEGenericBTP,
      ParamSoc,
      FactVariante,
      UEchangeVP,
      FactComm,
      DateUtils,
      UdateUtils,
      galPatience,
      CPProcGen,
      UTilFonctionCalcul, DB;


function DebutJournee (DateATrait : TDateTime) : TDateTime;
var DebJour : TTime;
begin
  DebJour := GetparamsocSecur('SO_BTAMDEBUT',StrToTime('08:00'));
  result := StrToDate(DateToStr(DateATrait))+DebJour;
end;


procedure SuprimeThisParagraphe(TOBL,TobDetails : TOB; var IndDep: integer);
var Niveau    : integer;
    TOBI      : TOB;
    StopItNow : boolean;
begin

  Niveau := TOBL.GetValue('GL_NIVEAUIMBRIC');

  StopItNow := false;

  repeat
    TOBI := TobDetails.detail[IndDep];
    if IsDebutParagraphe (TOBI,Niveau) then StopItNow := true;
    TOBI.free;
    Dec(IndDep);
  until (IndDep < 0 ) or (StopItNow);

end;


function IsDetailInside(IndDep: integer; TobDetails,TOBL: TOB): boolean;
var indice : integer;
    Niveau : integer;
    TOBI : TOB;
begin

  result := false;

  Niveau := TOBL.GetValue('GL_NIVEAUIMBRIC');

  for Indice := Inddep-1 downto 0 do
  begin
    TOBI := TobDetails.detail[Indice];
    if IsArticle (TOBI) then BEGIN result := true; break; END;
    if IsSousDetail (TOBI) then BEGIN result := true; break; END;
    if Isouvrage (TOBI) then BEGIN result := true; break; END;
    if IsDebutParagraphe (TOBI,Niveau) then break;
  end;

end;

procedure Nettoieparagraphes(TobTempo: TOB);
var Indice  : integer;
    TOBL    : TOB;
begin

  Indice := TobTempo.detail.count -1;

  repeat
    if Indice >= 0 then
    begin
      TOBL := TobTempo.detail[Indice];
      if IsFinParagraphe(TOBL) then
      begin
        if not IsDetailInside (Indice,TobTempo,TOBL) then
        begin
          SuprimeThisParagraphe (TOBL,TobTempo,Indice);
        end;
      end;
      dec(Indice);
    end;
  until indice <= 0;

end;

procedure TotalProjet (TOBTaches : TOB);

var i : integer;
		TOBL : TOB;
    ILien,Niv : Integer;
    Ptr : array [1..9] of Integer;

	procedure reinitPtr (Niveau : Integer);
  begin
    if Niveau < 2 then Ptr[1] := -1;
    if Niveau < 3 then Ptr[2] := -1;
    if Niveau < 4 then Ptr[3] := -1;
    if Niveau < 5 then Ptr[4] := -1;
    if Niveau < 6 then Ptr[5] := -1;
    if Niveau < 7 then Ptr[6] := -1;
    if Niveau < 8 then Ptr[7] := -1;
    if Niveau < 9 then Ptr[8] := -1;
    if Niveau < 10 then Ptr[9] := -1;
  end;

begin
  //
  iLien := -1;
	reinitPtr(1);
  //
	for i := 0 to TOBtaches.detail.count - 1 do
  begin
  	TOBL := TOBtaches.detail[i];
    if IsVariante ( TOBL) then continue;

    if IsDebutParagraphe(TOBL) then
    begin
      if TOBL.GeTString('LINKOK')= '-' then continue; 
      // --------------------------------------------------
      // Affectation des liens entre taches de meme niveau
      // --------------------------------------------------
      niv := TOBL.GetInteger('GL_NIVEAUIMBRIC');
      if Ptr[niv]<> -1 then
      begin
        TOBL.SetInteger('LIENTACHE',Ptr[niv]);
      end;
      Ptr[Niv] := i;
      reinitPtr(Niv+1); // reinit de tous les niveaux inférieurs
      // --------------------------------------------------
      if Niv= 1 then
      begin
        if iLien <> -1 then TOBL.SetInteger('LIENTACHE',iLien);
        iLien := i;
        if TobTaches.GetDateTime('DATEDEBUTPROJET') > TOBL.GetDateTime('DATEDEBUT') then
        begin
          TobTaches.SetDateTime('DATEDEBUTPROJET',TOBL.GetDateTime('DATEDEBUT'));
        end;
        if TobTaches.GetDateTime('DATEFINPROJET') < TOBL.GetDateTime('DATEFIN') then
        begin
          TobTaches.SetDateTime('DATEFINPROJET',TOBL.GetDateTime('DATEFIN'));
        end;
        TobTaches.SetDouble('DUREE', TobTaches.GetDouble('DUREE') + TOBL.GetDouble('MAXHRS'));
        TobTaches.SetDouble('TOTAL', TobTaches.GetDouble('TOTAL') + TOBL.GetDouble('TOTALHRS'));
        TobTaches.SetDouble('COUT', TobTaches.GetDouble('COUT') + TOBL.GetDouble('COUT'));
      end;
    end;
  end;
end;

procedure SommerLignes(TOBtaches: TOB; ii, Niv: integer; Var DecalagePhase : integer);
var TOBL, TOBTT, TOBDT: TOB;
  i,  iTypeL: integer;
  TypL : string;
  Max,Cumul,COUT : Double;
  DD : TDateTime;
begin
	iTypeL := -1;
  TOBDT := nil;
  TOBTT := TOBTAches.detail[ii]; // La tob de la fin de tache du niveau niv
  if TOBTT = nil then Exit;
  //
  DD := iDate2099;
  Max :=0 ;
  Cumul := 0;
  COUT := 0;
  // on essaye de cumuler les totaux de sous paragraphes   (sous phases)
  for i := ii -1 downto 0 do
  begin
		TOBL := TOBtaches.detail[i];
    if IsDebutParagraphe(TOBL,niv) then
    begin
      TOBDT := TOBL;
      Break;
    end;
    if isFinParagraphe(TOBL,niv+1) then
    begin
      if DD > TOBL.GetDateTime('DATEDEBUT') then
      begin
        DD := TOBL.GetDateTime('DATEDEBUT');
        TOBTT.SetDateTime('DATEDEBUT',DD);
      end;
      Max := Max + (TOBL.GetDouble('MAXHRS'));
      Cumul := Cumul +(TOBl.GetDouble('TOTALHRS'));
      Cout := Cout + TOBL.GetDouble('COUT')
    end;
  end;

  if Max <> 0 then
  begin
    // si on trouve des sous paragraphes on stocke les valeurs cumulés et l'on sort
    TOBTT.SetDateTime('DATEDEBUT',DD);
  	TOBTT.SetDouble('TOTALHRS',Cumul);
  	TOBTT.SetDouble('MAXHRS',Max);
  	TOBTT.SetDouble('COUT',Cout);
    TOBTT.SetDateTime('DATEFIN', AjouteDuree(TOBTT.GetDateTime('DATEDEBUT'),HeureBase100ToMinutes(TOBTT.GetDouble('MAXHRS')))); // remplacement ou nouveau
  end else
  begin
    DD := iDate2099;
    Max :=0 ;
    // ---------
    for i := ii -1 downto 0 do
    begin
      TOBL := TOBTAches.detail[i];
      if ii = 0 then
      begin
        iTypeL := TOBL.GetNumChamp('GL_TYPELIGNE');
      end;
      if TOBL = nil then Break;
      TypL := TOBL.Getstring(iTypeL);
      if IsParagraphe (TOBL,niv) then
      begin
        TOBDT := TOBL;
        Break;
      end else if IsFinParagraphe(TOBL) then
      begin
        //
      end else if IsParagraphe(TOBL) then Continue;
      //
      if DD > TOBL.GetDateTime('DATEDEBUT') then
      begin
        DD := TOBL.GetDateTime('DATEDEBUT');
        TOBTT.SetDateTime('DATEDEBUT',DD);
      end;
      //
      TOBTT.SetDouble('TOTALHRS',TOBTT.GetDouble('TOTALHRS')+TOBL.getDouble('TOTALHRS'));
      TOBTT.SetDouble('COUT',TOBTT.GetDouble('COUT')+TOBL.getDouble('COUT'));
      if TOBL.getDouble('TOTALHRS') > Max then
      begin
        TOBTT.SetDouble('MAXHRS',TOBL.getDouble('TOTALHRS'));
        max := TOBL.getDouble('TOTALHRS');
      end;
    end;
    //
    DecalagePhase := DecalagePhase + HeureBase100ToMinutes(TOBTT.GetDouble('MAXHRS'));
    TOBTT.SetDateTime('DATEFIN', AjouteDuree(TOBTT.GetDateTime('DATEDEBUT'),HeureBase100ToMinutes(TOBTT.GetDouble('MAXHRS')))); // remplacement ou nouveau
  end;

  if TOBDT <> nil then
  begin
  	TOBDT.SetDouble('TOTALHRS',TOBTT.GetDouble('TOTALHRS'));
  	TOBDT.SetDouble('MAXHRS',TOBTT.GetDouble('MAXHRS'));
  	TOBDT.SetDouble('COUT',TOBTT.GetDouble('COUT'));
    TOBDT.SetDateTime('DATEDEBUT',TOBTT.GetDateTime('DATEDEBUT'));
    TOBDT.SetDateTime('DATEFIN', AjouteDuree(TOBDT.GetDateTime('DATEDEBUT'),HeureBase100ToMinutes(TOBDT.GetDouble('MAXHRS')))); // remplacement ou nouveau
    TobTaches.SetDateTime('DATEFINPROJET',TOBTT.GetDateTime('DATEFIN'));
  end;
end;

procedure prepareLiaisonVP(TobTaches: TOB ; DateDebut : TDateTime; ModeTraitement : integer);
var ii,Niv,iNum : Integer;
		TOBTT : TOB;
    DateDebutprojet : TdateTime;
    DecalagePhase : integer;
    DateFinProjet : TDateTime;
    iTime0 : TTime;
    first : Boolean;
    Ratio : double;
begin
  first := true;
  iTime0 := StrToTime('00:00:00');
  if (TimeOf(DateDebut) = ITime0) then
  begin
    DateDebutProjet := DebutJournee(DateDebut);
  end;
  // --
  DateFinProjet := iDate1900;
  DecalagePhase := 0;
  iNum := 1;
	for ii := 0 to TobTaches.detail.count -1 do
  begin
		TOBTT := TobTaches.detail[ii];
    if isvariante(TOBTT) then Continue;

    if first then
    begin
      if (TOBTT.GetDateTime('DATEDEBUT')=IDate2099) or (TOBTT.GetDateTime('DATEDEBUT')=IDate1900) then
      begin
      	TobTaches.SetDateTime('DATEDEBUTPROJET',DateDebutProjet);
      end else
      begin
        DateDebutProjet := DebutJournee(TOBTT.GetDateTime('DATEDEBUT'));
      	TobTaches.SetDateTime('DATEDEBUTPROJET',DateDebutProjet);
      end;
      first := false;
    end;

    Niv := TOBTT.GetInteger('GL_NIVEAUIMBRIC');

    // ----------
    if not IsFinParagraphe(TOBTT) then
    begin
    	TOBTT.SetInteger('NUMTACHE',iNum);
      Inc(iNum);
    end;

    Niv := TOBTT.GetInteger('GL_NIVEAUIMBRIC');
    //
    if IsParagraphe(TOBTT) then
    begin
    	TOBTT.SetInteger('NIVTACHE',niv-1);
    end else
    begin
    	TOBTT.SetInteger('NIVTACHE',niv);
    end;
    //
    if (TOBTT.GetValue('DATEDEBUT')=iDate2099) or
       (TOBTT.GetDateTime('DATEDEBUT')=IDate1900) then
    begin
      TOBTT.SetDateTime('DATEDEBUT',AjouteDuree(DateDebutProjet,DecalagePhase)); // remplacement ou nouveau
    end else
    begin
			TOBTT.SetString('LINKOK','-');
      TOBTT.SetDateTime('DATEDEBUT', DebutJournee(TOBTT.GetDateTime('DATEDEBUT')));
    end;
    //
    if not IsParagraphe(TOBTT) then
    begin
      TOBTT.SetDouble('COUT',TOBTT.GetDouble('GL_MONTANTPR'));
      // ----------
      // -- reajustement a l'unite de mesure Heure --
      Ratio := Unite2Heures (TOBTT.GetString('GL_QUALIFQTEVTE'));
      TOBTT.SetDouble('MAXHRS',TOBTT.GetDouble('MAXHRS')*Ratio);
      TOBTT.SetDouble('TOTALHRS',TOBTT.GetDouble('TOTALHRS')*Ratio);
      //
      TOBTT.SetDateTime('DATEFIN', AjouteDuree(TOBTT.GetDateTime('DATEDEBUT'),HeureBase100ToMinutes(TOBTT.GetDouble('MAXHRS')))); // remplacement ou nouveau
      if TOBTT.GetDateTime('DATEFIN') > DateFinProjet then
      begin
        DateFinProjet := TOBTT.GetDateTime('DATEFIN');
        TobTaches.SetDateTime('DATEFINPROJET',DateFinProjet);
      end;
    end else if IsFinParagraphe(TOBTT) then
    begin
      SommerLignes(TobTaches, ii, Niv,DecalagePhase); // calcule le MAXHRS pour la fin de phase
    end;
  end;

  TotalProjet (TobTaches);
end;

procedure AddChampsSupTache (TobTaches : TOB);
begin
  TobTaches.AddChampSupValeur('IDPROJET',0);
  TobTaches.AddChampSupValeur('NUMDOSS','');
  TobTaches.AddChampSupValeur('REFERENCE','');
  TobTaches.AddChampSupValeur('RESPONSABLE','');
  TobTaches.AddChampSupValeur('COUT',0);
  TobTaches.AddChampSupValeur('DATEDEBUTPROJET',iDate2099);
  TobTaches.AddChampSupValeur('DATEFINPROJET',iDate1900);
  TobTaches.AddChampSupValeur('TOTAL',0);
  TobTaches.AddChampSupValeur('DUREE',0);
  TobTaches.AddChampSupValeur('DATABASE','');
  TobTaches.AddChampSupValeur('CLEDOC','');
end;

procedure TraitementProjetDocument (Cledoc : r_cledoc);
var SQL,NumDoss,reference,responsable,CodeResp : string;
		TOBpiece,TobTaches,TobAffaire : TOB;
    QQ : TQuery;
    CodeAffaire,CodeAffEdit : string;
		CBtnTxt : PCustBtnText;
    RespDLG : Integer;
    MajVp : TEchangeVP;
    Idprojet : Integer;
    DateDebut : TDateTime;
    ModeTraitement : Integer;
    respWarn : Integer;
    XX : TFPatience;
begin

  XX := FenetrePatience('Liaison VisualProjet',aoMilieu, False,true);
  XX.lAide.Caption := 'Vérification des données du projet en cours...';
  XX.lcreation.visible := false ;
  XX.StartK2000 ;

  //
	if Cledoc.NumeroPiece = 0 then Exit;
  ModeTraitement := XmodeCreat;
  RespDLG := mrCust1;
  //
  TobPiece := Tob.Create ('PIECE', nil,-1);
  TobTaches := TOB.Create('LES TACHES', nil, -1);
  AddChampsSupTache (TobTaches);
  TobAffaire := Tob.Create('LES AFFAIRES', nil, -1);
  MajVp := TEchangeVP.create;
  MajVp.ecran := XX;
  if not MajVp.OkLiaison then Exit;
  TRY
    Sql := 'SELECT * FROM PIECE WHERE '+WherePiece(Cledoc,ttdPiece,False);
    QQ := OpenSql (Sql,True,1,'',true);
    if not QQ.eof then
    begin
    	TobPiece.SelectDB('', QQ);
    end;
    ferme (QQ);
    //
    Idprojet := TOBpiece.GetInteger('GP_IDENTIFIANTWOT');
    if Idprojet  > 0 then // Déjà transféré
    begin
      if MajVp.BaseExiste(Idprojet) then
      begin
        New(CBtnTxt);
        CBtnTxt[mbCust1] := 'Remplacer';
        CBtnTxt[mbCust4] := 'Annuler';
        RespDLG :=  ExMessageDlg('Le projet existe déjà' + #13#10+'Que désirez-vous faire ?', mtxConfirmation, [mbCust1, mbCust4], 0, mbCust2, Application.Icon, CBtnTxt);
        dispose(CBtnTxt);
      end else
      begin
        Idprojet := 0;
      end;
    end;

    if RespDLG = mrCust4 then Exit;
    if (idProjet > 0) and (RespDLG = mrCust1) then
    begin
      respWarn := PGIAsk('ATTENTION : Vous allez réinitialiser le projet sous Visual Projet.#13#10 Confirmez-Vous ?');
      if respWarn = mrno then Exit;
    	ModeTraitement := XmodeRemplace
    end;
    //
  	XX.lAide.Caption := 'Préparation des données a envoyer';
    XX.Refresh;
    //
    //Chargement de la Tob Affaire pour export Excel
    CodeAffaire := TobPiece.GetString('GP_AFFAIRE');
    if codeAffaire <> '' then
    begin
    	CodeAffEdit := BTPCodeAffaireAffiche(CodeAffaire);
      //
      SQL := 'SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="' + CodeAffaire + '"';
      QQ := OpenSQL(Sql, false);
      TobAffaire.SelectDB(CodeAffaire,QQ,false);
      ferme(QQ);
    end;
    //chargement des lignes de prestation : TYPEARTICLE = 'PRE'
    SQL := 'SELECT GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMORDRE,'+ // element de la ressource
             'A.GA_NATUREPRES,A.GA_LIBELLE,A.GA_DPR,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,GL_NUMORDRE,GL_TYPELIGNE,GL_ARTICLE,GL_CODEARTICLE,GL_LIBELLE,GL_QTEFACT,GL_QUALIFQTEVTE,GL_DATELIVRAISON,'+
             'GP_TOTALHTDEV, '+
             'GL_TYPEARTICLE,GL_NIVEAUIMBRIC, GL_MONTANTPR, GLC_IDENTIFIANTWNT, '+ // element de la Tache
						 'GL_TYPELIGNE AS TYPELIGNE, (0) AS NUMTACHE, (0) AS NIVTACHE, GL_QTEFACT AS TOTALHRS, GL_QTEFACT AS MAXHRS,GL_DATELIVRAISON AS DATEDEBUT,'+
             'GL_DATELIVRAISON AS DATEFIN, (0) AS COUT, (-1) AS LIENTACHE, "X" AS LINKOK '+
             'FROM LIGNE '+
             'LEFT JOIN PIECE ON GL_NATUREPIECEG=GP_NATUREPIECEG AND GL_SOUCHE=GP_SOUCHE AND GL_NUMERO=GP_NUMERO AND GL_INDICEG=GP_INDICEG '+
             'LEFT JOIN ARTICLE A ON GA_ARTICLE=GL_ARTICLE '+
             'LEFT JOIN LIGNECOMPL ON GL_NATUREPIECEG=GLC_NATUREPIECEG AND GL_SOUCHE=GLC_SOUCHE AND GL_NUMERO=GLC_NUMERO AND GL_INDICEG=GLC_INDICEG AND Gl_NUMORDRE=GLC_NUMORDRE '+
             'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=GA_NATUREPRES '+
             'WHERE '+
             '(((GL_TYPEARTICLE ="PRE") AND (N.BNP_TYPERESSOURCE IN ("SAL","INT","MAT","OUT"))) '+
             'OR (GL_TYPEARTICLE ="FRA") '+
             'OR (GL_TYPELIGNE LIKE "TP%") OR (GL_TYPELIGNE LIKE "DV%") OR (GL_TYPELIGNE LIKE "DP%") OR (GL_TYPELIGNE LIKE "TV%") ) AND '+
             WherePiece(Cledoc, ttdLigne, false,false) + ' ORDER BY GL_NUMLIGNE';

    QQ  := OpenSQL(SQL, false);
    if Not QQ.eof then
    begin
      TobTaches.LoadDetailDB('LIGNE', '', '', QQ,false);
    end;
    ferme(QQ);
    //
    TobTaches.SetString('CLEDOC',EncodeRefPiece(cledoc));
    Nettoieparagraphes(TobTaches);
    // Création du projet
    if CodeAffaire <> '' then
    begin
      NumDoss := CodeAffEdit;
      Responsable := TOBAffaire.GetString('AFF_RESPONSABLE');
      DateDebut := TobAffaire.GetDateTime('AFF_DATEDEBUT');
    end else
    begin
      NumDoss := Cledoc.NaturePiece+':'+Cledoc.Souche+':'+InttOStr(Cledoc.NumeroPiece);
      Reference := NumDoss + ' - '+TOBPiece.GetString('GP_REFINTERNE');
      if Reference = '' then reference := NumDoss;
      Responsable := '';
      DateDebut := V_PGI.DateEntree;
    end;
    Reference := CodeAffEdit + ' - '+TobAffaire.GetString('AFF_LIBELLE');
    ModifCarAccentueKeep(reference); // pour enlever les catacrères accentués
    SupprimeCaracteresSpeciaux(reference,false,True,True); // suppression des caractères spéciaux
    Reference := Trim(Reference);

  	TobTaches.SetInteger('IDPROJET',IdProjet);
  	TobTaches.SetString('NUMDOSS',NumDoss);
  	TobTaches.SetString('REFERENCE',Reference);
    if responsable <> '' then
    begin
      QQ := OpenSQL('SELECT ARS_LIBELLE2,ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE="'+Responsable+'"',True,1,'',true);
      if not QQ.Eof then
      begin
      	CodeResp := QQ.findField('ARS_LIBELLE2').AsString+' '+QQ.findField('ARS_LIBELLE').AsString;
  			TobTaches.SetString('RESPONSABLE',CodeResp);
      end;
      ferme (QQ);
    end;

    //
    (*
    if (TimeOf(DateDebut) = ITime0) then
    begin
      DateDebut := DebutJournee(DateDebut);
    end;
    *)
    //
    prepareLiaisonVP(TobTaches,DateDebut,ModeTraitement);
    //
  	XX.lAide.Caption := 'Envoi des données..';
    XX.Refresh;
    if (idProjet <= 0) then
    begin
      MajVp.AjouteProjet (TOBTaches);
    end else
    begin
			MajVp.UpdateProjet (TOBTaches,true);
    end;
  FINALLY
    TobPiece.Free;
    TobTaches.Free;
    TobAffaire.Free;
    if MajVp <> nil then
    begin
    	MajVp.Free;
    end;
    XX.StopK2000 ;
  	XX.free;
  end;
end;

end.

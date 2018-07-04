{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/07/2013
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTGENODANAL ()
Mots clefs ... : TOF;BTGENODANAL
*****************************************************************}
Unit BTGENODANAL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul, MainEagl,
{$ENDIF}
     uTob,
		 HTB97,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     UtilsTOB,
     Vierge,
     Messages,
     Windows,
     AglInit,
     UdateUtils,
     UtilsRapport,
     VentilCpta
      ;

Type
  TResultControlMois = (TcmOk,TcmDejaTransfere,TcmARetransfere);
  TTyperesult = (TTRCode,TTRLibelle);

  TRefAna = record
  	Entity : Integer;
    Journal : string;
    Exercice : String;
    Axe : string;
    Piece : integer;
    Ligne : Integer;
  end;
//
// TOBresult
//     |
//     |  --> Livraison chantier
//     |       	 |
//     |				 |-->  TOB Compte comptable Général -- clef ( Compte General)
//     |         |   	 | --> TOB Section Générale (Section contrepartie) -- Clef ( 0,Section)
//     |         | 	   | --> TOB Section Détail (Chantier) -- Clef ( 1,Section)
//     |         | 	   | 		 | --> TOB Consommation
//     |         | 	   | --> ...
//     |				 |-->  ...
//     |  --> Retour chantier
//     |       	 |
//     |				 |-->  TOB Compte comptable Général -- clef ( Compte General)
//     |         |   	 | --> TOB Section Générale (Section contrepartie) -- Clef ( 0,Section)
//     |         | 	   | --> TOB Section Détail (Chantier) -- Clef ( 1,Section)
//     |         | 	   | 		 | --> TOB Consommation
//     |         | 	   | --> ...
//     |				 |-->  ...
//     |  --> Consommations chantier / nature prestation (MO,...)
//     |       	 |
//     |				 |-->  TOB Compte comptable Général -- clef ( Compte General)
//     |         |   	 | --> TOB Section Générale (Section contrepartie) -- Clef ( 0,Section )
//     |         | 	   | --> TOB Section Détail (Chantier) -- Clef ( 1,Section)
//     |         | 	   | 		 | --> TOB Consommation
//     |         | 	   | --> ...
//     |				 |-->  ...

  TOF_BTGENODANAL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    OneRapport : TGestionRapport;
    TOBResult,TOBVentil,TOBStructAna,TOBECRANAMOIS_O,TOBECRANAMOIS,TOBSTATUSANALMOIS,TOBREJETES,TOBStructAnalib : TOB;
    TOBREFAFFAIRE : TOB;
    Etat,EtatPrec : TResultControlMois;
		OptAnalStock,OptAnalConso : boolean;
    AxeVentil,JNLOD : string;
    Firstmois : string;
    GlastDay : TDateTime;
    LastErrorMessage : string;
    //
  	procedure SetDate;
    procedure ControleMouvementMois (TheDate : TdateTime);
    procedure ChangeDate (Sender : TObject);
    procedure GenOdAnalClick ( Sender : TObject);
    procedure ConstitueTexte(TheDate : TdateTime);
    procedure ConstitueTOBLivraisons(FirstDay, LastDay: TdateTime; ModeGenere : boolean);
    procedure ConstitueTOBRetourChantiers(FirstDay, LastDay: TdateTime; ModeGenere : boolean);
    procedure ConstitueTobsStocks(TheDate : TDateTime; ModeGenere : Boolean=true);
    procedure ConstitueTobsAutres(TheDate : TDateTime; ModeGenere : Boolean=true);
    function FindCpteGene(CodeCptaArt : string) : string;
		procedure GetVentilations;
		function RecupSection (TOBATrait: TOB; TTR : TTyperesult; IsAppel : boolean; AFFAIREREF : TOB) : string;

    procedure AfficheErreurs (Sender : TObject);
    function AffecteLigneStock(TOBPere: TOB; TOBAInsere: TOB; LastDay : TDateTime; LibMvt: string; ModeGenere : Boolean; Sens : string='+'): integer;
    function AffecteLigneConso(TOBPere: TOB; TOBAInsere: TOB; LastDay: TDateTime; LibMvt: string ; ModeGenere : Boolean ): integer;
    function GetSectionGeneConso(TOBTT : TOB) : string;
    procedure TraitementMouvements;
    procedure DeleteOldAnalytiques;
    procedure DeleteOldMvtsMois;

    procedure SetNumerotation;
    procedure GetParamsOdPrec;
    function AjouteRejets(TOBAREJ: TOB): TOB;
    procedure VERIFOdAnalClick(Sender: TOBject);
    function GenerationOK : boolean;
    procedure ConstitueRapport (TOBREJETES: TOB; OneRapport : TGestionRapport);
    procedure DeduitCumul(TOBAA: TOB);
		procedure AjouteCumul(TOBAA: TOB; Sens : string='+');
    procedure EcritOdAnalytiques;
    procedure GereStatusMois;
    procedure EcritMvtMois;
    function CreatSectionGenere(Axe,Code, libelle: string): Boolean;
    procedure ClickDelete (Sender : TObject);
    procedure Deletemouvements;
    procedure DeleteStatusMois;
    procedure ReinitAll;
    procedure GetOldAna;
    function AnalSurFacture(Naturepiece: string): boolean;
  end ;


function EncodeRefAna(TOBTT : TOB) : string;
procedure DecodeRefAna(Chaine : string ; var RefAna : TRefAna);
procedure UpdateStatusMoisOD (TOBPiece : TOB);

Implementation

uses DateUtils,ParamSoc,uEntCommun,Ent1,
     SaisUtil,CalcOLEGenericBTP,utilPGI,galPatience,
     BTPUtil,StockUtil,EntGC,Wcommuns;


function EncodeRefAna(TOBTT : TOB) : string;
begin
  Result := InttoStr(TOBTT.GetInteger('Y_ENTITY'))+';'+
            TOBTT.GetString('Y_JOURNAL') + ';' +
            TOBTT.GetString('Y_EXERCICE') + ';' +
            TOBTT.GetString('Y_AXE') + ';' +
            IntToStr(TOBTT.GetInteger('Y_NUMEROPIECE')) + ';' +
            IntToStr(TOBTT.GetInteger('Y_NUMLIGNE'));
end;

procedure DecodeRefAna(Chaine : string ; var RefAna : TRefAna);
var LocStr : string;
begin
  LocStr := Chaine;
	FillChar(RefAna,SizeOf(RefAna),#0);
  //
  RefAna.Entity := StrToInt(ReadTokenSt(LocStr));
  RefAna.Journal := READTOKENST(LocStr);
  RefAna.Exercice := READTOKENST(LocStr);
  RefAna.Axe := READTOKENST(LocStr);
  RefAna.Piece  := StrToInt(READTOKENST(LocStr));
  RefAna.Ligne  := StrToInt(READTOKENST(LocStr));
end;

procedure UpdateStatusMoisOD (TOBPiece : TOB);
var Sql : string;
		TheDate : TDateTime;
    QQ : TQuery;
    okok : Boolean;
begin
  okok := false;
  if TOBPiece.NomTable = 'PIECE' then
  begin
    TheDate := TOBPiece.GetDateTime('GP_DATEPIECE');
  end else if TOBPiece.NomTable = 'CONSOMMATIONS' then
  begin
    TheDate := TOBPiece.GetDateTime('BCO_DATEMOUV');
  end else Exit;

  Sql:= 'SELECT BAS_STATUSGEN FROM BSTATUSANALMOIS WHERE BAS_MOIS="'+ConstitueMois(TheDate)+'"';
  QQ := OpenSQL(SQl,True,1,'',true);
  if not QQ.eof then
  begin
    if QQ.fields[0].AsString <> 'ATT' then okok := True;
  end;
  ferme (QQ);
  if okok then
  begin
    SQl := 'UPDATE BSTATUSANALMOIS SET BAS_STATUSGEN="ATT" WHERE BAS_MOIS="'+ConstitueMois(TheDate)+'"';
    if ExecuteSQL(SQL) < 0 then
    begin
      V_PGI.IOError := oeUnknown;
    end;
  end;

end;

procedure TOF_BTGENODANAL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTGENODANAL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTGENODANAL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTGENODANAL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTGENODANAL.OnArgument (S : String ) ;
var TheDate : TDateTime;
begin
  Inherited ;
  OptAnalStock := GetparamSocSecur ('SO_OPTANALSTOCK',false);
  OptAnalConso := GetparamSocSecur ('SO_OPTANALCONSO',false);
	AxeVentil :=  GetParamSocSecur('SO_BTAXEANALSTOCK','') ;
  JNLOD :=  GetParamSocSecur('SO_BTJNLANALSTOCK','') ;
  Firstmois := ConstitueMois(GetParamSocSecur('SO_BTDEPARTODANAL',StrToDate('01/01/1900')));
  //
	TheDate := StrToDate(THEdit(GetControl('EDATE')).Text);
  //
  GLastDay := EndOfAMonth(YearOf(TheDate),MonthOf(TheDate));
  //
  TToolbarButton97 (GetControl('BLANCE')).OnClick := GenOdAnalClick;
  TToolbarButton97 (GetControl('BVERIF')).OnClick := VERIFOdAnalClick;
  THEdit(GetControl('EDATE')).OnChange := ChangeDate;
  TOBSTATUSANALMOIS := TOB.Create ('BSTATUSANALMOIS',nil,-1);
  TOBresult := TOB.Create('LES RESULTATS',nil,-1);
  TOBVentil := TOB.Create ('LES VENTILATIONS',nil,-1);
  TOBStructAna := TOB.Create('STRUCTURES ANAL',nil,-1);
  TOBStructAnalib := TOB.Create ('STRUCT ANAL LIB',nil,-1);
	TOBECRANAMOIS := TOB.Create ('LES ECRITURES ANAL / MOIS',nil,-1);
	TOBECRANAMOIS_O := TOB.Create ('LES ANC ECRITURES ANAL / MOIS',nil,-1);
  TOBREJETES := TOB.Create ('LES REJETS',nil,-1);
  TOBREFAFFAIRE := TOB.Create ('AFFAIRE',nil,-1);
  //
  OneRapport := TGestionRapport.Create(ecran);
  OneRapport.Titre := 'Anomalies détectées';
  OneRapport.Affiche := True;
  OneRapport.Close := True;
  OneRapport.InitRapport;
  //
  SetDate;
  GetVentilations;
  //
  if (not OptAnalConso) and (not OptAnalStock) then
  begin
    PgiInfo('Vous ne gérez l''analytique, ni sur stock, ni sur les consommations');
    PostMessage(TFVierge(Ecran).Handle, WM_CLOSE, 0, 0);
  end;
  if TOBVentil.Detail.count = 0 then
  begin
    PGIInfo('ATTENTION : Aucun paramétrage pour la ventilation comptable des consommations n''a été défini');
    PostMessage(TFVierge(Ecran).Handle, WM_CLOSE, 0, 0);
  end;
  if TOBStructAna.Detail.count = 0 then
  begin
    PGIError('ATTENTION : Aucun paramétrage pour la ventilation analytique des consommations n''a été défini');
    PostMessage(TFVierge(Ecran).Handle, WM_CLOSE, 0, 0);
  end;

  TToolbarButton97 (GetControl('BDELETE')).OnClick := ClickDelete;
end ;

procedure TOF_BTGENODANAL.OnClose ;
begin
  TOBresult.Free;
  TOBVentil.free;
  TOBStructAna.Free;
  TOBStructAnalib.Free;
  TOBECRANAMOIS.free;
  TOBECRANAMOIS_O.free;
  TOBREFAFFAIRE.free;
  TOBREJETES.free;
  OneRapport.free;

  Inherited ;
end ;

procedure TOF_BTGENODANAL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTGENODANAL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTGENODANAL.SetDate;
var TheDate : TdateTime;
		MMLoc : string;
begin
	TToolbarButton97 (GetControl('BDELETE')).Visible := false;
  THImage(GetControl('TSTATUS')).OnClick := nil;
	TheDate := StrToDate(THEdit(GetControl('EDATE')).Text);
  MMLoc := ConstitueMois (TheDate);
  if MMLoc <  Firstmois then
  begin
    PgiInfo ('Mouvements analytiques non gérés sur ce mois');
    SetControlText('EDATE',DateToStr(GLastDay));
    Exit;
  end;
  GLastDay := EndOfAMonth(YearOf(TheDate),MonthOf(TheDate));
  ControleMouvementMois (TheDate);
  ConstitueTexte(TheDate);
  if (Etat = TcmOk) and (EtatPrec = TcmOk) then THImage(GetControl('TSTATUS')).GlobalIndexImage := 'M0047_S24G1'
  																				 else THImage(GetControl('TSTATUS')).GlobalIndexImage := 'M0046_S24G1';
  if EtatPrec <> TcmOk then
  begin
  	THImage(GetControl('TSTATUS')).OnClick := AfficheErreurs;
  end;
  if Etat = TcmDejaTransfere then
  begin
    TToolbarButton97 (GetControl('BDELETE')).Visible := True;
  end;
  GetParamsOdPrec;
end;

procedure TOF_BTGENODANAL.VERIFOdAnalClick (Sender : TOBject);
var TheDate : TdateTime;
		Regen : Boolean;
    XX : TFPatience;
begin
  Regen := true;
  if TOBREJETES.Detail.Count > 0 then
  begin
    if PGIAsk('ATTENTION : Des anomalies sont déjà présentes.#13#10Désirez-vous refaire le contrôle ?',Ecran.caption) <> mryes then
    	Regen := false;
  end;
  if Regen then
  begin
  	if PGIAsk('Confirmez-vous le lancement de contrôle ?',Ecran.caption) <> mryes then Exit;
    OneRapport.InitRapport;
    TOBResult.ClearDetail;
    TOBREJETES.ClearDetail;
    //
    XX := FenetrePatience('Vérifications des données',aoMilieu, False,true);
    XX.lcreation.visible := false ;
    XX.StartK2000 ;
    //
    TheDate := StrToDate(THEdit(GetControl('EDATE')).Text);
  	XX.lAide.Caption := 'Mouvements de stock...';
    XX.refresh;
    if OptAnalStock then ConstitueTobsStocks (TheDate,False);
  	XX.lAide.Caption := 'Mouvements de consommations...';
    XX.refresh;
    if OptAnalConso then ConstitueTobsAutres (TheDate,False);
    XX.StopK2000 ;
    XX.free;
    if not GenerationOK then
    begin
      ConstitueRapport (TOBREJETES,OneRapport);
      OneRapport.PosLeft := 0;
      OneRapport.Postop := Ecran.Height div 2 ;
      OneRapport.Width := Ecran.Width;
      OneRapport.Height := Ecran.Height div 2;
      OneRapport.AfficheRapport;
    end else
    begin
      PGIInfo('Aucune anomalie détectée...');
    end;
  end else
  begin
//    ConstitueRapport (TOBREJETES,OneRapport);
    OneRapport.AfficheRapport;
  end;
end;

procedure TOF_BTGENODANAL.GenOdAnalClick(Sender: TObject);
var TheDate : TdateTime;
		result : TIOErr;
    XX : TFPatience;
begin
  if PGIAsk('Confirmez-vous le lancement de génération des OD Analytiques ?',Ecran.caption) <> mryes then Exit;
  LastErrorMessage := '';
  ReinitAll;
  XX := FenetrePatience('Vérifications des données',aoMilieu, False,true);
  XX.lcreation.visible := false ;
  //
	TheDate := StrToDate(THEdit(GetControl('EDATE')).Text);
  XX.lAide.Caption := 'Mouvements de stock...';
  XX.StartK2000 ;
  XX.refresh;
  if OptAnalStock then ConstitueTobsStocks (TheDate);
  xx.stopK2000;
  XX.lAide.Caption := 'Consommations...';
  XX.StartK2000 ;
  XX.refresh;
	if OptAnalConso then ConstitueTobsAutres (TheDate);
  XX.StopK2000 ;
  XX.free;

  if not GenerationOK then
  begin
    if PGIAsk('ATTENTION : Des anomalies ont été constatées.#13#10Désirez-vous quand même générer les écritures ?',Ecran.caption) <> mryes then Exit;
  end;
  XX := FenetrePatience('Ecritures des OD Analytiques',aoMilieu, False,true);
  XX.lcreation.visible := false ;
  XX.StartK2000 ;
  XX.Refresh;
  result := TRANSACTIONS (TraitementMouvements,0);
  XX.StopK2000 ;
  XX.free;
  if result <> OeOk then
  begin
    PGIError('ABANDON : '+LastErrorMessage);
  end else
  begin
    PGIInfo('Traitement terminé.');
  end;
  SetDate;
end;

procedure TOF_BTGENODANAL.ChangeDate(Sender: TObject);
begin
	SetDate;
end;

procedure TOF_BTGENODANAL.ConstitueTexte(TheDate: TdateTime);
begin
	TMemo(GetControl('MINFOS')).Clear;
  if Etat = TcmOk then
  begin
  	TMemo(GetControl('MINFOS')).Lines.Add('');
  	TMemo(GetControl('MINFOS')).Lines.Add('Les mouvements analytiques n''ont jamais été générés sur ce mois');
  end else if Etat = TcmDejaTransfere then
  begin
  	TMemo(GetControl('MINFOS')).Lines.Add('');
  	TMemo(GetControl('MINFOS')).Lines.Add('ATTENTION : Les mouvements analytiques ont déjà été générés sur ce mois');
  	TMemo(GetControl('MINFOS')).Lines.Add('Si vous relancez ce traitement l''ensemble des mouvements seront réinitialisés');
  	TMemo(GetControl('MINFOS')).Lines.Add('');
  end else if Etat = TcmARetransfere then
  begin
  	TMemo(GetControl('MINFOS')).Lines.Add('');
  	TMemo(GetControl('MINFOS')).Lines.Add('Des éléments ont été modifiés ou ajoutés après génération des mouvements');
  	TMemo(GetControl('MINFOS')).Lines.Add('Vous devez relancer la génération...');
  end;
  if EtatPrec <> TcmOk then
  begin
    TMemo(GetControl('MINFOS')).Lines.Add('');
    TMemo(GetControl('MINFOS')).Lines.Add('ATTENTION : Certains mois précédents devraient être regénérés...');
    TMemo(GetControl('MINFOS')).Lines.Add('Cliquer sur le bouton rouge pour voir la liste');
  end;
end;

procedure TOF_BTGENODANAL.ControleMouvementMois(TheDate: TdateTime);

		procedure  DefiniLesEtatMensuels (TOBDATAS: TOB; FirstDate,LastDate : TdateTime);
    var DateCourante : TDateTime;
    		MoisCourant,SQL : string;
        TOBD : TOB;
        QQ : TQuery;
    begin
      TOBDATAS.ClearDetail;
      DateCourante := EncodeDate(YearOf(FirstDate),MonthOf(FirstDate),1);
      //
      DateCourante := IncMonth(DateCourante,1);
      repeat
        if MonthOf(DateCourante) >= MonthOf(LastDate) then Exit;
        MoisCourant := ConstitueMois(DateCourante);
        TOBD := TOB.Create('BSTATUSANALMOIS',TOBDATAS,-1);
        TOBD.SetString('BAS_MOIS',MoisCourant);
        TOBD.SetString('BAS_STATUSGEN','ATT');
        DateCourante := IncMonth(DateCourante,1);
      until MonthOf(DateCourante) > MonthOf(LastDate);
      //
      SQL := 'SELECT * FROM BSTATUSANALMOIS WHERE BAS_MOIS > "'+ConstitueMois(FirstDate)+'" AND BAS_MOIS < "'+ConstitueMois(LastDate)+'"';
      QQ := OpenSQL(SQL,True,-1,'',true);
      if not QQ.Eof then
      begin
        QQ.First;
        Repeat
          TOBD := TOBDATAS.FindFirst(['BAS_MOIS'],[QQ.FindField('BAS_MOIS').AsString],True);
          if TOBD <> nil then
          begin
            TOBD.SetString('BAS_STATUSGEN',QQ.FindField('BAS_STATUSGEN').AsString);
          end;
          QQ.Next;
        until QQ.Eof;
      end;
      ferme (QQ);
    end;

    procedure ControleEtatPrec(TOBDATAS: TOB; var EtatPrec : TResultControlMois);
    var Indice : Integer;
    begin
			for Indice := 0 to TOBDATAS.detail.count -1 do
      begin
        if TOBDATAS.Detail[Indice].GetString('BAS_STATUSGEN') <> 'DEJ' then
        begin
        	EtatPrec:= TcmARetransfere;
          break;
        end;
      end;
    end;


var QQ : TQuery;
		Req : String;
    FirstDate,lastDate : TDateTime;
    TOBDATAS : TOB;
begin
  TOBDATAS := TOB.Create ('LES ETAT MENSUELS',nil,-1);
  TRY
    Etat := TcmOk;
    EtatPrec := TcmOk;
    FirstDate := GetParamSocSecur('SO_BTDEPARTODANAL',StrToDate('01/01/1900'));
    LastDate := EndOfAMonth(YearOf(TheDate),MonthOf(TheDate));
    DefiniLesEtatMensuels (TOBDATAS,FirstDate,lastDate);
    //
    Req := 'SELECT BAS_STATUSGEN FROM BSTATUSANALMOIS WHERE BAS_MOIS="'+ConstitueMois(TheDate)+'"';
    QQ := OpenSQL(Req,True,1,'',true);
    if not QQ.Eof then
    begin
      if QQ.Fields [0].AsString = 'DEJ' then Etat := TcmDejaTransfere else
      if QQ.Fields [0].AsString = 'ATT' then Etat := TcmARetransfere ;
    end;
    ferme (QQ);
    //
    if FirstMois <> ConstitueMois(TheDate) then
    begin
    	ControleEtatPrec(TOBDATAS,EtatPrec);
    end;
  FINALLY
    TOBDATAS.Free;
  END;
end;

procedure TOF_BTGENODANAL.AfficheErreurs(Sender: TObject);
var OneTOB : TOB;
begin
	OneTOB := TOB.Create('UNE TOB',nil,-1);
  OneTOB.AddChampSupValeur('DEPART',FirstMois);
  OneTOB.AddChampSupValeur('FIN',ConstitueMois(StrToDate(THEdit(GetControl('EDATE')).Text)));
  TheTOB := OneTOB;
	AGLLanceFiche('BTP','BTCONSULTODANAL','','','');
  TheTOB := nil;
  OneTOB.Free;
end;

procedure TOF_BTGENODANAL.ConstitueTobsAutres(TheDate : TDateTime; ModeGenere : boolean);

	function LienConsoPaie : Boolean;
  begin
    Result := false;
    if GetParamSocSecur('SO_AFLIENPAIEVAR',False) and GetParamSocSecur('SO_AFLIENPAIEANA',False) then
    begin
      Result := True;
    end;
  end;

var Req,MvtPaie,NaturePrest,LibNature,MM : string;
    QQ : TQuery;
    TOBMvts,TOBFather : TOB;
    II : Integer;
    FirstDay,LastDay : TDateTime;
begin
  FirstDay := RecodeDay (TheDate,1);
  LastDay := EndOfAMonth(YearOf(TheDate),MonthOf(TheDate));
  MM := ConstitueMois (LastDay);

  MvtPaie := '';
  NaturePrest := '';
  // Attention travaille par exclusion
  If LienConsoPaie then
  begin
    MvtPaie := ',"MO","FRS"'; // Si liaison avec la paie les OD arrivent via la paie depuis le transfert
    													// des consommations dans la paie
  end;
  // ------------------------
  TOBMvts := TOB.Create ('LES MVTS',nil,-1);
  TRY
    // Constitution des retours de chantier sur stock (donc)
    Req := 'SELECT "NATUREPRES" = CASE GA_TYPEARTICLE WHEN "FRA" THEN "---" WHEN "MAR" THEN "---" ELSE GA_NATUREPRES END,' +
    			 '0 AS ERREUR,'+	
    			 'CONSOMMATIONS.*,ARTICLE.*,AFFAIRE.*'+
           'FROM CONSOMMATIONS '+
           'LEFT JOIN ARTICLE ON GA_ARTICLE=BCO_ARTICLE '+
           'LEFT JOIN AFFAIRE ON AFF_AFFAIRE=BCO_AFFAIRE '+
           'WHERE '+
           'BCO_DATEMOUV >= "' + UsDateTime(FirstDay) + '" AND BCO_DATEMOUV <= "'+UsDateTime(LastDay)+'" AND '+
           '(BCO_NATUREPIECEG="" OR BCO_NATUREPIECEG="LBT" OR BCO_NATUREPIECEG = "FPR" OR (BCO_NATUREPIECEG="FAC")) AND ' +
           'BCO_NATUREMOUV NOT IN ("FOU","FAN","RAN"'+MvtPaie+ ') ' +
           'ORDER BY NATUREPRES,BCO_NATUREMOUV,BCO_AFFAIRE';
    //
    QQ := OpenSql (Req,True,-1,'',true);
    if not QQ.eof then
    begin
      TOBFather := nil;
      //
      TOBMvts.LoadDetailDB('CONSOMMATIONS','','',QQ,false);
      II := 0;
      repeat
        if TOBMvts.detail[II].GetString('NATUREPRES')<> NaturePrest then
        begin
        	NaturePrest := TOBMvts.detail[II].GetString('NATUREPRES');
          if NaturePrest <> '---' then
          begin
            LibNature := RechDom('BTNATPRESTATION',NaturePrest,false);
          end else
          begin
            LibNature := 'Frais';
          end;
          TOBFather := TOB.Create('CONSOMMATIONS '+LibNature,TOBResult ,-1);
        end;
      	AffecteLigneConso (TOBFather,TOBMvts.detail[II],LastDay,'OD '+LibNature+' '+MM, ModeGenere);
      until II >= TOBMvts.detail.count ;
    end;
  FINALLY
  	ferme (QQ);
    TOBMvts.Free;
  end;
end;

procedure TOF_BTGENODANAL.ConstitueTOBRetourChantiers (FirstDay,LastDay  : TdateTime; ModeGenere : boolean);
var Req : string;
    QQ : TQuery;
    TOBMvts,TOBFather : TOB;
    II : Integer;
    MM : string;
begin
	MM := ConstitueMois (LastDay);

  TOBMvts := TOB.Create ('LES MVTS',nil,-1);
  TRY
    // Constitution des retours de chantier sur stock (donc)
    Req := 'SELECT "---" AS NATUREPRES, 0 AS ERREUR,'+
    			 'CONSOMMATIONS.*,ARTICLE.*,AFFAIRE.*, LIGNE.GL_DEPOT,LIGNE.GL_PIECEPRECEDENTE,LIGNE.GL_PIECEORIGINE '+
           'FROM CONSOMMATIONS '+
           'LEFT JOIN LIGNE ON GL_NATUREPIECEG=BCO_NATUREPIECEG AND GL_SOUCHE=BCO_SOUCHE AND GL_NUMERO=BCO_NUMERO AND '+
           'GL_INDICEG=BCO_INDICEG AND GL_NUMORDRE=BCO_NUMORDRE '+
           'LEFT JOIN ARTICLE ON GA_ARTICLE=BCO_ARTICLE '+
           'LEFT JOIN AFFAIRE ON AFF_AFFAIRE=BCO_AFFAIRE '+
           'WHERE '+
//           '(GL_PIECEORIGINE="" OR GL_PIECEORIGINE IS NULL OR SUBSTRING(GL_PIECEORIGINE,10,3)="PBT") AND '+
           'BCO_DATEMOUV >= "' + UsDateTime(FirstDay) + '" AND BCO_DATEMOUV <= "'+UsDateTime(LastDay)+'" AND '+
           'BCO_NATUREPIECEG IN ("BFC") AND BCO_NATUREMOUV="FOU" ' +
           'ORDER BY BCO_AFFAIRE,GA_COMPTAARTICLE';
    //
    QQ := OpenSql (Req,True,-1,'',true);
    if not QQ.eof then
    begin

      TOBFather := TOB.Create('LES RETOURS',TOBresult,-1);
      //
      TOBMvts.LoadDetailDB('CONSOMMATIONS','','',QQ,false);
      II := 0;
      repeat
        if not IsLivChantier (TOBMvts.detail[II].GetString('GL_PIECEPRECEDENTE'),TOBMvts.detail[II].GetString('GL_PIECEORIGINE')) then
        begin
      	  AffecteLigneStock (TOBFather,TOBMvts.detail[II],LastDay,'OD RETOUR CHANTIER - '+MM,ModeGenere,'-');
        end
        //FV1 : 20/03/2018 - FS#3020 - TEAM - Génération OD analytique boucle
        else inc(II);
      until II >= TOBMvts.detail.count ;
    end;
  FINALLY
  	ferme (QQ);
    TOBMvts.Free;
  end;
end;

function TOF_BTGENODANAL.AffecteLigneConso (TOBPere: TOB; TOBAInsere : TOB; LastDay : TDateTime;  LibMvt : string; ModeGenere : Boolean) : integer;
var TOBCGENE,TTF1,TTG,TOBREJ  : TOB;
		CompteGene,Section1,SectionGene,CodChantier,LibSection : string;
    DateMvt : TDateTime;
    MtMvt : Double;
    FromAppel : Boolean;
    Appel : string;
    QQ : Tquery;
begin
  Result := 0;
  FromAppel := false;
  DateMvt := StrToDate(DateToStr(LastDay));
  MtMvt := TOBAInsere.GetDouble('BCO_MONTANTACH');
  Appel := TOBAInsere.Getvalue('BCO_AFFAIRE');

  if MtMvt = 0 then
  begin
  	TOBAInsere.setInteger('ERREUR',5);
    TOBREJ := AjouteRejets(TOBAInsere);
    TOBAInsere.ChangeParent(TOBREJ,-1);
    Exit;
  end;

	CompteGene := FindCpteGene(TOBAInsere.GetString('GA_COMPTAARTICLE'));
  if Comptegene = '' then CompteGene := GetParamSocSecur('SO_GCCPTEHTACH',''); // si le compte n'est pas définie
  																																						 // on recup le compta d'achat par défaut
  if CompteGene = '' then
  begin
  	TOBAInsere.setInteger('ERREUR',2);
    TOBREJ := AjouteRejets(TOBAInsere);
    TOBAInsere.ChangeParent(TOBREJ,-1) ;
    Exit;
  end;
  SectionGene := GetSectionGeneConso(TOBAInsere);
  if SectionGene = '<' then
  begin
  	TOBAInsere.setInteger('ERREUR',1);
    TOBREJ := AjouteRejets(TOBAInsere);
    TOBAInsere.ChangeParent(TOBREJ,-1);
    Exit;
  end else if SectionGene = '' then
  begin
  	TOBAInsere.setInteger('ERREUR',3);
    TOBREJ := AjouteRejets(TOBAInsere);
    TOBAInsere.ChangeParent(TOBREJ,-1);
    Exit;
  end;

  If (Copy(TOBAInsere.Getvalue('BCO_AFFAIRE'),1,1) = 'W') then
  begin
    CodChantier:=TOBAInsere.GetString('AFF_CHANTIER');
    if (CodChantier <> '') and (Trim(Copy(CodChantier,2,14)) <> '')  then
    begin
    	FromAppel := True;
      QQ := OpenSql ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+CodChantier+'"',true,1,'',true);
      if not QQ.eof then TOBREFAFFAIRE.SelectDB ('',QQ);
      ferme (QQ);
    end;
  end;
 	Section1 := RecupSection (TOBAInsere,TTRCode,FromAppel,TOBREFAFFAIRE);
  LibSection := RecupSection(TOBAInsere,TTRLibelle,FromAppel,TOBREFAFFAIRE);
  CreatSectionGenere(AxeVentil,Section1,LibSection);
  if FromAppel then
  begin
    TOBREFAFFAIRE.InitValeurs(false);
  end;
  if Section1 = '' then
  BEGIN
  	TOBAInsere.setInteger('ERREUR',4);
    TOBREJ := AjouteRejets(TOBAInsere);
    TOBAInsere.ChangeParent(TOBREJ,-1);
  	Exit; // pas de ventilation analytique donc on degage
  end;

  if not ModeGenere then BEGIN TOBAInsere.Free; Exit; end;

  TOBCGENE := TOBGetInto(TOBPere,['COMPTEGENE'],[Comptegene]);
  if TOBCGENE = nil then
  begin
    TOBCGENE := TOB.Create('UN COMPTE GENE',nil,-1);
    TOBCGENE.AddChampSupValeur('COMPTEGENE',CompteGene);
  	TOBInsertInto (TOBPere,TOBCGENE,['COMPTEGENE']);
  end;
  // Ajoute Section détail
  TTF1 := TOBGetInto (TOBCGENE,['INDICE','Y_SECTION'],[1,Section1]);
  if TTF1 = nil then
  begin
    TTF1 := TOB.Create('ANALYTIQ',nil,-1);
    TTF1.AddChampSupValeur('INDICE',1);
    TTF1.setString  ('Y_GENERAL',CompteGene);
    TTF1.setString  ('Y_AXE',AxeVentil);
    TTF1.setString  ('Y_SECTION',Section1);
    TTF1.setString  ('Y_EXERCICE',QUELEXODT(DateMvt));
    //
    TTF1.SetString  ('Y_LIBELLE',LibMvt);
    TTF1.SetString  ('Y_NATUREPIECE','OD');
    TTF1.SetString  ('Y_QUALIFPIECE','N');
    TTF1.SetString  ('Y_TYPEANALYTIQUE','X');
    TTF1.SetString  ('Y_ETAT','0000000000');
    TTF1.SetString  ('Y_UTILISATEUR',V_PGI.User);
    TTF1.SetDateTime('Y_DATECREATION',DateMvt);
    TTF1.SetString  ('Y_SOCIETE',V_PGI.CodeSociete);
    TTF1.SetString  ('Y_ETABLISSEMENT',GetParamSocSecur('SO_ETABLISDEFAUT', ''));
    TTF1.SetString  ('Y_DEVISE'            , V_PGI.DevisePivot ) ;
    TTF1.SetDouble  ('Y_TAUXDEV'           , 1 ) ;
    TTF1.SetDateTime('Y_DATETAUXDEV'       , DateMvt ) ;
    TTF1.SetString  ('Y_CONTROLE'          , '-' ) ;
    TTF1.SetString  ('Y_QUALIFQTE1'        , '...' ) ;
    TTF1.SetString  ('Y_QUALIFQTE2'        , '...' ) ;
    TTF1.SetString  ('Y_QUALIFECRQTE1'     , '...' ) ;
    TTF1.SetString  ('Y_QUALIFECRQTE2'     , '...' ) ;
    TTF1.SetString  ('Y_ECRANOUVEAU'       , 'N' ) ;
    TTF1.SetString  ('Y_CREERPAR'          , 'SAI' ) ;
    TTF1.SetString  ('Y_EXPORTE'           , '---' ) ;
    TTF1.SetString  ('Y_JOURNAL'           , JNLOD ) ;
    TTF1.SetString  ('Y_VALIDE'            , '-' ) ;
    TTF1.SetString  ('Y_CONFIDENTIEL'      , '0' ) ;
    TTF1.SetDateTime('Y_DATPER',iDate1900) ;
    TTF1.SetInteger ('Y_ENTITY',0) ;
    TTF1.SetString  ('Y_REFGUID','') ;
//	  TTF1.SetDateTime('Y_DATEMODIF',LastDay);
    TOBInsertInto(TOBCGENE,TTF1,['INDICE','Y_SECTION']);
  end;
  if MtMvt >= 0 then
  begin
    TTF1.SetDouble('Y_DEBITDEV',TTF1.GetDouble('Y_DEBITDEV')+Abs(MtMvt));
    TTF1.SetDouble('Y_DEBIT',TTF1.GetDouble('Y_DEBIT')+Abs(MtMvt));
  end else
  begin
    TTF1.SetDouble('Y_CREDITDEV',TTF1.GetDouble('Y_CREDITDEV')+Abs(MtMvt));
    TTF1.SetDouble('Y_CREDIT',TTF1.GetDouble('Y_CREDIT')+Abs(MtMvt));
  end;
  TOBAInsere.ChangeParent(TTF1,-1);
  // --- La même chose sur la section générale
  // Ajoute Section détail
  TTG := TOBGetInto (TOBCGENE,['INDICE','Y_SECTION'],[0,SectionGene]);
  if TTG = nil then
  begin
    TTG := TOB.Create('ANALYTIQ',nil,-1);
    TTG.AddChampSupValeur('INDICE',0);
    TTG.setString  ('Y_GENERAL'           ,CompteGene);
    TTG.setString  ('Y_AXE'               ,AxeVentil);
    TTG.SetDateTime('Y_DATECOMPTABLE'     ,DateMvt);
    TTG.setString  ('Y_SECTION'           ,SectionGene);
    TTG.setString  ('Y_EXERCICE'          ,QUELEXODT(DateMvt));
    //
    TTG.SetString  ('Y_LIBELLE'           ,LibMvt);
    TTG.SetString  ('Y_NATUREPIECE'       ,'OD');
    TTG.SetString  ('Y_QUALIFPIECE'       ,'N');
    TTG.SetString  ('Y_TYPEANALYTIQUE'    ,'X');
    TTG.SetString  ('Y_ETAT'              ,'0000000000');
    TTG.SetString  ('Y_UTILISATEUR'       ,V_PGI.User);
    TTG.SetDateTime('Y_DATECREATION'      ,DateMvt);
    TTG.SetString  ('Y_SOCIETE'           ,V_PGI.CodeSociete);
    TTG.SetString  ('Y_ETABLISSEMENT'     ,GetParamSocSecur('SO_ETABLISDEFAUT', ''));
    TTG.SetString  ('Y_DEVISE'            , V_PGI.DevisePivot ) ;
    TTG.SetDouble  ('Y_TAUXDEV'           , 1 ) ;
    TTG.SetDateTime('Y_DATETAUXDEV'       , DateMvt ) ;
    TTG.SetString  ('Y_CONTROLE'          , '-' ) ;
    TTG.SetString  ('Y_QUALIFQTE1'        , '...' ) ;
    TTG.SetString  ('Y_QUALIFQTE2'        , '...' ) ;
    TTG.SetString  ('Y_QUALIFECRQTE1'     , '...' ) ;
    TTG.SetString  ('Y_QUALIFECRQTE2'     , '...' ) ;
    TTG.SetString  ('Y_ECRANOUVEAU'       , 'N' ) ;
    TTG.SetString  ('Y_CREERPAR'          , 'SAI' ) ;
    TTG.SetString  ('Y_EXPORTE'           , '---' ) ;
    TTG.SetString  ('Y_JOURNAL'           ,  JNLOD) ;
    TTG.SetString  ('Y_VALIDE'            , '-' ) ;
    TTG.SetString  ('Y_CONFIDENTIEL'      , '0' ) ;
    TTG.SetDateTime('Y_DATPER'            ,	iDate1900) ;
    TTG.SetInteger ('Y_ENTITY'            ,0) ;
    TTG.SetString  ('Y_REFGUID'           ,'') ;
//	  TTG.SetDateTime('Y_DATEMODIF'         ,LastDay);
    TOBInsertInto(TOBCGENE,TTG,['INDICE','Y_SECTION']);
  end;
  if MtMvt >= 0 then
  begin
    TTG.SetDouble('Y_CREDITDEV',TTG.GetDouble('Y_CREDITDEV')+Abs(MtMvt));
    TTG.SetDouble('Y_CREDIT',TTG.GetDouble('Y_CREDIT')+Abs(MtMvt));
  end else
  begin
    TTG.SetDouble('Y_DEBITDEV',TTG.GetDouble('Y_DEBITDEV')+Abs(MtMvt));
    TTG.SetDouble('Y_DEBIT',TTG.GetDouble('Y_DEBIT')+Abs(MtMvt));
  end;
end;

function TOF_BTGENODANAL.AffecteLigneStock (TOBPere: TOB; TOBAInsere : TOB; LastDay : TDateTime;  LibMvt : string; ModeGenere : Boolean; Sens : string='+') : integer;
var TOBCGENE,TTF1,TTG,TOBREJ : TOB;
		CompteGene,Section1,SectionGene,CodChantier,LibSection : string;
    DateMvt : TDateTime;
    MtMvt : Double;
    FromAppel : Boolean;
    Depot : string;
    QQ : TQuery;
begin
  Result := 0;
  TOBREFAFFAIRE.InitValeurs(false); 
  FromAppel := false;
  DateMvt := StrToDate(DateToStr(LastDay));
  MtMvt := TOBAInsere.GetDouble('BCO_MONTANTACH');
  //
  if MtMvt = 0 then
  begin
  	TOBAInsere.setInteger('ERREUR',5);
    TOBREJ := AjouteRejets(TOBAInsere);
    TOBAInsere.ChangeParent(TOBREJ,-1);
    Exit;
  end;

	CompteGene := FindCpteGene(TOBAInsere.GetString('GA_COMPTAARTICLE'));
  if Comptegene = '' then CompteGene := GetParamSocSecur('SO_GCCPTEHTACH',''); // si le compte n'est pas définie
  																																						 // on recup le compta d'achat par défaut
  if CompteGene = '' then
  begin
  	TOBAInsere.setInteger('ERREUR',2);
    TOBREJ := AjouteRejets(TOBAInsere);
    TOBAInsere.ChangeParent(TOBREJ,-1);
    Exit;
  end;

//  SectionGene := GetParamSocSecur('SO_BTCPTANALSTOCK','');
	Depot := TOBAInsere.GetString('GL_DEPOT');
	SectionGene := BTGetSectionStock (TOBAInsere.GetString('GL_DEPOT'));
  if SectionGene = '<' then
  begin
  	TOBAInsere.setInteger('ERREUR',1);
    TOBREJ := AjouteRejets(TOBAInsere);
    TOBAInsere.ChangeParent(TOBREJ,-1);
    Exit;
  end else if SectionGene = '' then
  begin

  	TOBAInsere.setInteger('ERREUR',3);
    TOBREJ := AjouteRejets(TOBAInsere);
    TOBAInsere.ChangeParent(TOBREJ,-1);
    Exit;
  end;
  If (Copy(TOBAInsere.Getvalue('BCO_AFFAIRE'),1,1) = 'W') then
  begin
    CodChantier:=TOBAInsere.GetString('AFF_CHANTIER');
    if (CodChantier <> '') and (Trim(Copy(CodChantier,2,14)) <> '')  then
    begin
    	FromAppel := True;
      QQ := OpenSql ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+CodChantier+'"',true,1,'',true);
      if not QQ.eof then TOBREFAFFAIRE.SelectDB ('',QQ);
      ferme (QQ); 
    end;
  end;
 	Section1 := RecupSection (TOBAInsere,TTRCode,FromAppel,TOBREFAFFAIRE);
  LibSection := RecupSection(TOBAInsere,TTRLibelle,FromAppel,TOBREFAFFAIRE);
  CreatSectionGenere(AxeVentil,Section1,LibSection);
  if FromAppel then
  begin
    TOBREFAFFAIRE.InitValeurs(false);
  end;
  if Section1 = '' then
  BEGIN
  	TOBAInsere.setInteger('ERREUR',4);
    TOBREJ := AjouteRejets(TOBAInsere);
    TOBAInsere.ChangeParent(TOBREJ,-1);
  	Exit;
  END; // pas de ventilation analytique donc on degage

  if not ModeGenere then BEGIN TOBAInsere.Free; Exit; end;


  TOBCGENE := TOBGetInto(TOBPere,['COMPTEGENE'],[Comptegene]);
  if TOBCGENE = nil then
  begin
    TOBCGENE := TOB.Create('UN COMPTE GENE',nil,-1);
    TOBCGENE.AddChampSupValeur('COMPTEGENE',CompteGene);
  	TOBInsertInto (TOBPere,TOBCGENE,['COMPTEGENE']);
  end;

  // Ajoute Section détail
  TTF1 := TOBGetInto (TOBCGENE,['INDICE','Y_SECTION'],[1,Section1]);
  if TTF1 = nil then
  begin
    TTF1 := TOB.Create('ANALYTIQ',nil,-1);
    TTF1.AddChampSupValeur('INDICE',1);
    TTF1.setString  ('Y_GENERAL',CompteGene);
    TTF1.setString  ('Y_AXE',AxeVentil);
    TTF1.SetDateTime('Y_DATECOMPTABLE',DateMvt);
    TTF1.setString  ('Y_SECTION',Section1);
    TTF1.setString  ('Y_EXERCICE',QUELEXODT(DateMvt));
    //
    TTF1.SetString  ('Y_LIBELLE',LibMvt);
    TTF1.SetString  ('Y_NATUREPIECE','OD');
    TTF1.SetString  ('Y_QUALIFPIECE','N');
    TTF1.SetString  ('Y_TYPEANALYTIQUE','X');
    TTF1.SetString  ('Y_ETAT','0000000000');
    TTF1.SetString  ('Y_UTILISATEUR',V_PGI.User);
    TTF1.SetDateTime('Y_DATECREATION',DateMvt);
    TTF1.SetString  ('Y_SOCIETE',V_PGI.CodeSociete);
    TTF1.SetString  ('Y_ETABLISSEMENT',GetParamSocSecur('SO_ETABLISDEFAUT', ''));
    TTF1.SetString  ('Y_DEVISE'            , V_PGI.DevisePivot ) ;
    TTF1.SetDouble  ('Y_TAUXDEV'           , 1 ) ;
    TTF1.SetDateTime('Y_DATETAUXDEV'       , DateMvt ) ;
    TTF1.SetString  ('Y_CONTROLE'          , '-' ) ;
    TTF1.SetString  ('Y_QUALIFQTE1'        , '...' ) ;
    TTF1.SetString  ('Y_QUALIFQTE2'        , '...' ) ;
    TTF1.SetString  ('Y_QUALIFECRQTE1'     , '...' ) ;
    TTF1.SetString  ('Y_QUALIFECRQTE2'     , '...' ) ;
    TTF1.SetString  ('Y_ECRANOUVEAU'       , 'N' ) ;
    TTF1.SetString  ('Y_CREERPAR'          , 'SAI' ) ;
    TTF1.SetString  ('Y_EXPORTE'           , '---' ) ;
    TTF1.SetString  ('Y_JOURNAL'           , JNLOD ) ;
    TTF1.SetString  ('Y_VALIDE'            , '-' ) ;
    TTF1.SetString  ('Y_CONFIDENTIEL'      , '0' ) ;
    TTF1.SetDateTime('Y_DATPER',iDate1900) ;
    TTF1.SetInteger ('Y_ENTITY',0) ;
    TTF1.SetString  ('Y_REFGUID','') ;
//	  TTF1.SetDateTime('Y_DATEMODIF',LastDay);
    TOBInsertInto(TOBCGENE,TTF1,['INDICE','Y_SECTION']);
  end;
  if Sens = '+' then
  begin
    if MtMvt > 0 then
    begin
      TTF1.SetDouble('Y_DEBITDEV',TTF1.GetDouble('Y_DEBITDEV')+Abs(MtMvt));
      TTF1.SetDouble('Y_DEBIT',TTF1.GetDouble('Y_DEBIT')+Abs(MtMvt));
    end else
    begin
      TTF1.SetDouble('Y_CREDITDEV',TTF1.GetDouble('Y_CREDITDEV')+Abs(MtMvt));
      TTF1.SetDouble('Y_CREDIT',TTF1.GetDouble('Y_CREDIT')+Abs(MtMvt));
    end;
  end else
  begin
    if MtMvt > 0 then
    begin
      TTF1.SetDouble('Y_CREDITDEV',TTF1.GetDouble('Y_CREDITDEV')+Abs(MtMvt));
      TTF1.SetDouble('Y_CREDIT',TTF1.GetDouble('Y_CREDIT')+Abs(MtMvt));
    end else
    begin
      TTF1.SetDouble('Y_DEBITDEV',TTF1.GetDouble('Y_DEBITDEV')+Abs(MtMvt));
      TTF1.SetDouble('Y_DEBIT',TTF1.GetDouble('Y_DEBIT')+Abs(MtMvt));
    end;
  end;
  TOBAInsere.ChangeParent(TTF1,-1);
  // --- La même chose sur la section générale
  // Ajoute Section détail
  TTG := TOBGetInto (TOBCGENE,['INDICE','Y_SECTION'],[0,SectionGene]);
  if TTG = nil then
  begin
    TTG := TOB.Create('ANALYTIQ',nil,-1);
    TTG.AddChampSupValeur('INDICE',0);

    TTG.setString  ('Y_GENERAL'           ,CompteGene);
    TTG.setString  ('Y_AXE'               ,AxeVentil);
    TTG.SetDateTime('Y_DATECOMPTABLE'     ,DateMvt);
    TTG.setString  ('Y_SECTION'           ,SectionGene);
    TTG.setString  ('Y_EXERCICE'          ,QUELEXODT(DateMvt));
    //
    TTG.SetString  ('Y_LIBELLE'           ,LibMvt);
    TTG.SetString  ('Y_NATUREPIECE'       ,'OD');
    TTG.SetString  ('Y_QUALIFPIECE'       ,'N');
    TTG.SetString  ('Y_TYPEANALYTIQUE'    ,'X');
    TTG.SetString  ('Y_ETAT'              ,'0000000000');
    TTG.SetString  ('Y_UTILISATEUR'       ,V_PGI.User);
    TTG.SetDateTime('Y_DATECREATION'      ,DateMvt);
    TTG.SetString  ('Y_SOCIETE'           ,V_PGI.CodeSociete);
    TTG.SetString  ('Y_ETABLISSEMENT'     ,GetParamSocSecur('SO_ETABLISDEFAUT', ''));
    TTG.SetString  ('Y_DEVISE'            , V_PGI.DevisePivot ) ;
    TTG.SetDouble  ('Y_TAUXDEV'           , 1 ) ;
    TTG.SetDateTime('Y_DATETAUXDEV'       , DateMvt ) ;
    TTG.SetString  ('Y_CONTROLE'          , '-' ) ;
    TTG.SetString  ('Y_QUALIFQTE1'        , '...' ) ;
    TTG.SetString  ('Y_QUALIFQTE2'        , '...' ) ;
    TTG.SetString  ('Y_QUALIFECRQTE1'     , '...' ) ;
    TTG.SetString  ('Y_QUALIFECRQTE2'     , '...' ) ;
    TTG.SetString  ('Y_ECRANOUVEAU'       , 'N' ) ;
    TTG.SetString  ('Y_CREERPAR'          , 'SAI' ) ;
    TTG.SetString  ('Y_EXPORTE'           , '---' ) ;
    TTG.SetString  ('Y_JOURNAL'           ,  JNLOD) ;
    TTG.SetString  ('Y_VALIDE'            , '-' ) ;
    TTG.SetString  ('Y_CONFIDENTIEL'      , '0' ) ;
    TTG.SetDateTime('Y_DATPER'            ,	iDate1900) ;
    TTG.SetInteger ('Y_ENTITY'            ,0) ;
    TTG.SetString  ('Y_REFGUID'           ,'') ;
//	  TTG.SetDateTime('Y_DATEMODIF'         ,LastDay);
    TOBInsertInto(TOBCGENE,TTG,['INDICE','Y_SECTION']);
  end;
  if Sens = '+' then
  begin
    if MtMvt > 0 then
    begin
      TTG.SetDouble('Y_CREDITDEV',TTG.GetDouble('Y_CREDITDEV')+Abs(MtMvt));
      TTG.SetDouble('Y_CREDIT',TTG.GetDouble('Y_CREDIT')+Abs(MtMvt));
    end else
    begin
      TTG.SetDouble('Y_DEBITDEV',TTG.GetDouble('Y_DEBITDEV')+Abs(MtMvt));
      TTG.SetDouble('Y_DEBIT',TTG.GetDouble('Y_DEBIT')+Abs(MtMvt));
    end;
  end else
  begin
    if MtMvt > 0 then
    begin
      TTG.SetDouble('Y_DEBITDEV',TTG.GetDouble('Y_DEBITDEV')+Abs(MtMvt));
      TTG.SetDouble('Y_DEBIT',TTG.GetDouble('Y_DEBIT')+Abs(MtMvt));
    end else
    begin
      TTG.SetDouble('Y_CREDITDEV',TTG.GetDouble('Y_CREDITDEV')+Abs(MtMvt));
      TTG.SetDouble('Y_CREDIT',TTG.GetDouble('Y_CREDIT')+Abs(MtMvt));
    end;
  end;
end;

function TOF_BTGENODANAL.AnalSurFacture (Naturepiece : string) : boolean;
var Passp : string;
begin
  PassP := GetInfoParPiece(Naturepiece, 'GPP_TYPEECRCPTA');
  result := ((PassP <> '') and (PassP <> 'RIE'));
end;

procedure TOF_BTGENODANAL.ConstitueTOBLivraisons (FirstDay,LastDay : TdateTime; ModeGenere : boolean);
var Req,MM : string;
    QQ : TQuery;
    TOBMvts,TOBFather : TOB;
    II : Integer;
    StFacture : string;
begin
  Stfacture := '(BCO_NATUREPIECEG="FAC") AND (BCO_FACTURABLE="N")'; // par defaut
  if not AnalSurFacture('FAC') then stfacture := 'BCO_NATUREPIECEG="FAC"';
	MM := ConstitueMois (LastDay);
  TOBMvts := TOB.Create ('LES MVTS',nil,-1);
  TRY
    // Constitution des livraisons sur chantier ou consommations de fourniture vers chantier (provenance stock)
    Req := 'SELECT "---" AS NATUREPRES, 0 AS ERREUR,'+
    			 'CONSOMMATIONS.*,ARTICLE.*,AFFAIRE.*,LIGNE.GL_DEPOT,LIGNE.GL_PIECEPRECEDENTE,LIGNE.GL_PIECEORIGINE '+
           'FROM CONSOMMATIONS '+
           'LEFT JOIN LIGNE ON GL_NATUREPIECEG=BCO_NATUREPIECEG AND GL_SOUCHE=BCO_SOUCHE AND GL_NUMERO=BCO_NUMERO AND '+
           'GL_INDICEG=BCO_INDICEG AND GL_NUMORDRE=BCO_NUMORDRE '+
           'LEFT JOIN ARTICLE ON GA_ARTICLE=BCO_ARTICLE '+
           'LEFT JOIN AFFAIRE ON AFF_AFFAIRE=BCO_AFFAIRE '+
           'WHERE '+
//           '(GL_PIECEORIGINE="" OR GL_PIECEORIGINE IS NULL OR SUBSTRING(GL_PIECEORIGINE,10,3)="PBT") AND '+
           'BCO_DATEMOUV >= "' + UsDateTime(FirstDay) + '" AND BCO_DATEMOUV <= "'+UsDateTime(LastDay)+'" AND '+
           '(BCO_NATUREPIECEG="" OR BCO_NATUREPIECEG="LBT" OR ('+Stfacture+')) AND ' +
           'BCO_NATUREMOUV="FOU" ' +
           'ORDER BY BCO_AFFAIRE,GA_COMPTAARTICLE';
    //
    QQ := OpenSql (Req,True,-1,'',true);
    if not QQ.eof then
    begin
      TOBFather := TOB.Create('LES LIVRAISONS',TOBresult,-1);
      //
      TOBMvts.LoadDetailDB('CONSOMMATIONS','','',QQ,false);
      II := 0;
      repeat
        if not IsLivChantier (TOBMvts.detail[II].GetString('GL_PIECEPRECEDENTE'),TOBMvts.detail[II].GetString('GL_PIECEORIGINE')) then
        begin
      	  AffecteLigneStock (TOBFather,TOBMvts.detail[II],LastDay,'OD STOCK - '+MM, ModeGenere);
        end else inc(II);
      until II >= TOBMvts.detail.count;
    end;
  FINALLY
  	ferme (QQ);
    TOBMvts.Free;
  end;
end;

procedure TOF_BTGENODANAL.ConstitueTobsStocks (TheDate : TDateTime; ModeGenere : Boolean);
	procedure ErrParamGene ;
  begin
		PGIError('Le paramétrage de cette fonctionnalité n''a pas été effectué');
  end;

var		FirstDay,LastDay : TDateTime;
			XX : string;
begin
  FirstDay := RecodeDay (TheDate,1);
  LastDay := EndOfAMonth(YearOf(TheDate),MonthOf(TheDate));
  //
  if (not IsParamGeneDepot) then begin ErrParamGene; Exit; end;
  XX :=  GetParamSocSecur('SO_BTAXEANALSTOCK','') ;
  if XX = '' then begin ErrParamGene; Exit; end;
  XX :=  GetParamSocSecur('SO_BTJNLANALSTOCK','') ;
  if XX = '' then begin ErrParamGene; Exit; end;
  //
  ConstitueTOBLivraisons (FirstDay,LastDay, ModeGenere);
  ConstitueTOBRetourChantiers (FirstDay,LastDay,ModeGenere);
end;

procedure TOF_BTGENODANAL.GetVentilations;
var REq : String;
		QQ : TQuery;
begin
	Req := 'SELECT GCP_COMPTAARTICLE,GCP_CPTEGENEACH FROM CODECPTA WHERE GCP_VENTEACHAT="CON" ORDER BY GCP_COMPTAARTICLE';
  QQ := OpenSql (Req,True,-1,'',True);
  if not QQ.Eof then
  begin
  	TOBVentil.LoadDetailDB('CODECPTA','','',QQ,false);
  end;
  Ferme(QQ);
  //
  if AxeVentil = '' then Exit;
  REq := 'SELECT * '+
  			 'FROM DECOUPEANA '+
         'WHERE GDA_AXE="'+AxeVentil+'" AND GDA_TYPECOMPTE="CON" AND GDA_TYPESTRUCTANA="SEC" '+
  			 'ORDER BY GDA_RANG';
  QQ := OpenSql (Req,True,-1,'',True);
  if not QQ.Eof then
  begin
  	TOBStructAna.LoadDetailDB('DECOUPEANA','','',QQ,false);
  end;
  ferme (QQ);
  //
  REq := 'SELECT * '+
  			 'FROM DECOUPEANA '+
         'WHERE GDA_AXE="'+AxeVentil+'" AND GDA_TYPECOMPTE="CON" AND GDA_TYPESTRUCTANA="LSE" '+
  			 'ORDER BY GDA_RANG';
  QQ := OpenSql (Req,True,-1,'',True);
  if not QQ.Eof then
  begin
  	TOBStructAnalib.LoadDetailDB('DECOUPEANA','','',QQ,false);
  end;
  ferme (QQ);
end;

function TOF_BTGENODANAL.FindCpteGene(CodeCptaArt: string): string;
var TT : TOB;
begin
  Result := '';
	TT := TOBVentil.findFirst(['GCP_COMPTAARTICLE'],[CodeCptaArt],False);
  if TT = nil then
  begin
		TT := TOBVentil.findFirst(['GCP_COMPTAARTICLE'],[''],False);
  end;
  if TT <> nil then Result := TT.GetString('GCP_CPTEGENEACH');
end;

function TOF_BTGENODANAL.RecupSection(TOBATrait: TOB; TTR : TTyperesult; IsAppel : boolean; AFFAIREREF : TOB ): string;
var II : Integer;
		TOBS : TOB;
    Lng : Integer;
    Champs,CodeVide,value,CodeAna : string;
begin
	Result := '';
  CodeVide := '';
  if TTR = TTRCode then
  begin
    for II:=0 to TOBStructAna.Detail.Count-1 do
    begin
      TOBS := TOBStructAna.detail[II];
      Lng := TOBS.GetInteger('GDA_LONGUEUR');
      Champs := TOBS.GetString('GDA_LIBCHAMP');
      if (wGetPrefixe(Champs)='AFF') and (IsAppel) and (AFFAIREREF.getString('AFF_AFFAIRE')<>'') then
      begin
        if AFFAIREREF.FieldExists(Champs) then value := AFFAIREREF.GetString(Champs);
      end else
      begin
        if TOBATrait.FieldExists(Champs) then value := TOBATrait.GetString(Champs);
      end;
      CodeVide := CodeVide + BourreChamps ('',Lng,TFichierBase(Ord(fbAxe1)),True);
      Value:=BourreChamps (Value,Lng,TFichierBase(Ord(fbAxe1)),True);
      CodeAna := CodeAna + value;
    end;
    if length(CodeAna)>0 then Result := BourreLaDonc(CodeAna,TFichierBase(Ord(fbAxe1)))
    else Result := '';
  end else
  begin
    for II:=0 to TOBStructAnalib .Detail.Count-1 do
    begin
      TOBS := TOBStructAnalib.detail[II];
      Lng := TOBS.GetInteger('GDA_LONGUEUR');
      Champs := TOBS.GetString('GDA_LIBCHAMP');
      if (wGetPrefixe(Champs)='AFF') and (IsAppel) and (AFFAIREREF.getString('AFF_AFFAIRE')<>'') then
      begin
        if AFFAIREREF.FieldExists(Champs) then value := AFFAIREREF.GetString(Champs);
      end else
      begin
        if TOBATrait.FieldExists(Champs) then value := Trim(Copy(TOBATrait.GetString(Champs),1,Lng));
      end;
      CodeAna := CodeAna + value;
    end;
    Result := Trim(CodeAna);
  end;
end;

function TOF_BTGENODANAL.GetSectionGeneConso(TOBTT: TOB): string;
var QQ : TQuery;
		SectionArt : string;
    SortieTTSuite : boolean;
begin
  Result := '';
  SortieTTSuite := false;
  QQ := OpenSQL('SELECT GA_SECTION,GA_GEREANAL FROM ARTICLE WHERE GA_ARTICLE="'+TOBTT.GetString('BCO_ARTICLE')+'"',True,1,'',true);
  if not QQ.eof then
  begin
    SectionArt := QQ.Fields[0].AsString;
    SortieTTSuite := (QQ.fields[1].AsString <> 'X');
  end;
  ferme (QQ);

  if SortieTTSuite then BEGIN Result := '<' ; Exit; end;

	if TOBTT.GetString('BCO_RESSOURCE')<> '' then
  begin
		QQ := OpenSQL('SELECT BRS_SECTION FROM BRESSOURCE WHERE BRS_RESSOURCE="'+TOBTT.GetString('BCO_RESSOURCE')+'"',True,1,'',true);
    if not QQ.eof then
    begin
      Result := QQ.Fields[0].AsString;
    end;
    ferme (QQ);
  end;
  if Result <> '' then Exit;
  //
  if SectionArt <> '' then BEGIN result := SectionArt; end;
  if Result <> '' then Exit;
  //
  if TOBTT.GetString('NATUREPRES') = '---' then Exit; // pas de nature de prestation pour les frais ou les matériaux
  QQ := OpenSQL('SELECT BNP_SECTION FROM NATUREPREST WHERE BNP_NATUREPRES="'+TOBTT.GetString('GA_NATUREPRES')+'"',True,1,'',true);
  if not QQ.eof then
  begin
    Result := QQ.Fields[0].AsString;
  end;
  ferme (QQ);

end;

procedure TOF_BTGENODANAL.DeleteOldMvtsMois;
var TheDate : TDateTime;
		Mois,Sql : string;
begin
  //
	TheDate := StrToDate(THEdit(GetControl('EDATE')).Text);
  Mois := ConstitueMois(TheDate);
  if not ExisteSql ('SELECT * FROM BDETANALMOIS WHERE BA1_MOIS="'+ Mois + '"') then Exit;
  Sql := 'DELETE FROM BDETANALMOIS WHERE BA1_MOIS="'+ Mois + '"';
  if ExecuteSQL(SQL)< 1 then
  begin
    LastErrorMessage := 'Erreur 1200104';
    V_PGI.IOError := oeSaisie;
  end;
end;


procedure TOF_BTGENODANAL.SetNumerotation;

	procedure  SetNumCpta(TOBTT : TOB;JNL: string; DateMvt : TdateTime);
  var TOBII,TOBLL : TOB;
  		II : Integer;
      NumCpta : integer;
  begin
    //
    NumCpta:=GetNewNumJal(JNL,true,DateMvt); // pour chaque compte général
		for II := 0 to TOBTT.detail.count -1 do
    begin
      TOBLL := TOBTT.detail[II];
      TOBLL.SetInteger ('Y_NUMVENTIL',II+1);
      TOBLL.SetDateTime('Y_DATECOMPTABLE',DateMvt);
      TOBLL.SetInteger ('Y_PERIODE',GetPeriode(DateMvt)) ;
      TOBLL.SetInteger ('Y_SEMAINE',NumSemaine(DateMvt)) ;
      TOBLL.SetInteger ('Y_NUMEROPIECE',NumCpta);
      TOBLL.SetInteger ('Y_NUMLIGNE',0);
      if II = 0 then
      begin
        TOBLL.SetString('Y_TYPEMVT','AE');
        //
        TOBII := TOB.Create ('BDETANALMOIS',TOBECRANAMOIS,-1);
        TOBII.SetString('BA1_MOIS',ConstitueMois(DateMvt));
        TOBII.SetString('BA1_REFANALYTIQUE',EncodeRefAna(TOBLL));
      end else
      begin
      	TOBLL.SetString('Y_TYPEMVT','AL');
      end;
    end;
  end;

var II,JJ : Integer;
		TOBUT : TOB;
    LastDay,TheDate : TDateTime;
begin
	TheDate := StrToDate(THEdit(GetControl('EDATE')).Text);
	LastDay := StrToDate(DateToStr(EndOfAMonth(YearOf(TheDate),MonthOf(TheDate))));

//	JNLOD
	for II := 0 to TOBResult.Detail.Count -1 do
  begin
		TOBUT := TOBResult.Detail[II];
    if TOBUT.Detail.Count = 0 then continue;
    // on va traiter un lot d'écriture sur un compte général
    for JJ := 0 to TOBUT.Detail.count -1 do
    begin
    	SetNumCpta(TOBUT.detail[JJ],JNLOD,LastDay);
    end;
  end;

end;

procedure TOF_BTGENODANAL.EcritMvtMois;
var Indice : Integer;
begin
	if TOBECRANAMOIS.Detail.count = 0 then Exit;
  for Indice := 0 to TOBECRANAMOIS.Detail.Count -1 do
  begin
		if not TOBECRANAMOIS.detail[Indice].InsertDB(nil) then
    begin
    	LastErrorMessage := 'Erreur 1200106';
    	V_PGI.IOError := oeUnknown;
      break;
    end;
  end;
end;

procedure TOF_BTGENODANAL.GetOldAna;
var QQ : TQuery;
		TheDate : TDateTime;
    MMLOc : string;
begin
	TheDate := StrToDate(THEdit(GetControl('EDATE')).Text);
  MMLoc := ConstitueMois (TheDate);
  TOBECRANAMOIS_O.ClearDetail;
  QQ := OpenSQL('SELECT * FROM BDETANALMOIS WHERE BA1_MOIS="'+MMLoc +'"',True,-1,'',true);
  if not QQ.Eof then
  begin
		TOBECRANAMOIS_O.LoadDetailDB('BDETANALMOIS','','',QQ,false);
  end;
  Ferme (QQ);
end;

procedure TOF_BTGENODANAL.TraitementMouvements;
begin
  // Recupération des anicennes écritures analytiques
  GetOldAna;
	// Suppression des anciennes OD
  DeleteOldAnalytiques;
  // Supprime les anciens mouvements du mois
  if V_PGI.IOError = OeOk then DeleteOldMvtsMois;
  // numérotation des ecritures
  if V_PGI.IOError = OeOk then SetNumerotation;
  // Ecritures des nouvelles OD Analytiques
  if V_PGI.IOError = OeOk then EcritOdAnalytiques;
  // Ecriture dans BDETANALMOIS
  if V_PGI.IOError = OeOk then EcritMvtMois;
  // ECRITURE DANS BSTATUSANALMOIS
  if V_PGI.IOError = OeOk then GereStatusMois;
end;

procedure TOF_BTGENODANAL.AjouteCumul(TOBAA: TOB; Sens : string='+');
var XD,XC,CE,DS,CS,DP,CP,DE : double;
		DateC : TDateTime;
    Num : Integer;
    Jal,AX,SECTION : string;
    Ste : string3;
    SQl : string;
begin
{capture des infos en vue des MAJ}
  DateC:=TOBAA.GetValue('Y_DATECOMPTABLE') ; StE:=QuelExo(DateToStr(DATEC)) ;
  Num:=TOBAA.GetValue('Y_NUMEROPIECE') ;
  XD:=TOBAA.GetValue('Y_DEBIT') ; XC:=TOBAA.GetValue('Y_CREDIT') ;
  Jal:=TOBAA.GetValue('Y_JOURNAL') ;
  AX := TOBAA.GetValue('Y_AXE') ;
  Section := TOBAA.GetValue('Y_SECTION') ;
  DE:=0 ; CE:=0 ; DP:=0 ; CP:=0 ; DS:=0 ; CS:=0 ;
  if StE=GetEncours.Code then
  BEGIN
  	DE:=XD ;
    CE:=XC;
  END else
  begin
    if StE=GetSuivant.Code then
    BEGIN
      DS:=XD ;
      CS:=XC ;
    END else
    BEGIN
      DP:=XC ;
      CP:=XC ;
    END ;
  end;

  if not EstTablePartagee( 'JOURNAL' ) then
  begin
    if Sens = '+' then
    begin
      SQL:='UPDATE JOURNAL SET J_DEBITDERNMVT='+StrfPoint(XD)+', J_CREDITDERNMVT='+StrfPoint(XC)+', '
          +'J_DATEDERNMVT="'+USDateTime(DateC)+'", J_NUMDERNMVT='+IntToStr(Num)+', J_TOTALDEBIT=J_TOTALDEBIT+'+Strfpoint(XD)+', '
          +'J_TOTALCREDIT=J_TOTALCREDIT+'+StrfPoint(XC)+', J_TOTDEBE=J_TOTDEBE+'+StrfPoint(DE)+', J_TOTCREE=J_TOTCREE+'+StrfPoint(CE)+','
          +'J_TOTDEBS=J_TOTDEBS+'+StrfPoint(DS)+', J_TOTCRES=J_TOTCRES+'+StrfPoint(CS)+', '
          +'J_TOTDEBP=J_TOTDEBP+'+StrfPoint(DP)+', J_TOTCREP=J_TOTCREP+'+StrfPoint(CP)+' WHERE J_JOURNAL="'+Jal+'"' ;
    end else
    begin
      SQL:='UPDATE JOURNAL SET '
          +'J_TOTALDEBIT=J_TOTALDEBIT-'+Strfpoint(XD)+','
          +'J_TOTALCREDIT=J_TOTALCREDIT-'+StrfPoint(XC)+','
          +'J_TOTDEBE=J_TOTDEBE-'+StrfPoint(DE)+','
          +'J_TOTCREE=J_TOTCREE-'+StrfPoint(CE)+','
          +'J_TOTDEBS=J_TOTDEBS-'+StrfPoint(DS)+','
          +'J_TOTCRES=J_TOTCRES-'+StrfPoint(CS)+', '
          +'J_TOTDEBP=J_TOTDEBP-'+StrfPoint(DP)+','
          +'J_TOTCREP=J_TOTCREP-'+StrfPoint(CP)+' WHERE J_JOURNAL="'+Jal+'"' ;
    end;
    if ExecuteSQL(SQL)<>1 then
    begin
      LastErrorMessage := 'Erreur 1200101';
    	V_PGI.IoError:=oeStock;  // FQ 20001 BVE 25.04.07 On utilise oeStock car pas de oeJournal
    end;
  end;
  if VH^.LiaisonY2ViaShare then
  begin
      if Sens = '+' then
      begin
        SQL:='UPDATE CSECTION SET CSP_DEBITDERNMVT='+StrfPoint(XD)+',  CSP_CREDITDERNMVT='+StrfPoint(XC)+', '
        +'CSP_DATEDERNMVT="'+UsDateTime(DateC)+'", CSP_NUMDERNMVT='+IntToStr(Num)+', '
        +'CSP_TOTALDEBIT=CSP_TOTALDEBIT+'+StrfPoint(XD)+', CSP_TOTALCREDIT=CSP_TOTALCREDIT+'+StrfPoint(XC)+', '
        +'CSP_TOTDEBE=CSP_TOTDEBE+'+StrfPoint(DE)+', CSP_TOTCREE=CSP_TOTCREE+'+StrfPoint(CE)+', '
        +'CSP_TOTDEBS=CSP_TOTDEBS+'+Strfpoint(DS)+', CSP_TOTCRES=CSP_TOTCRES+'+StrfPoint(CS)+', '
        +'CSP_TOTDEBP=CSP_TOTDEBP+'+StrfPoint(DP)+', CSP_TOTCREP=CSP_TOTCREP+'+StrfPoint(CP)+
        ' WHERE CSP_AXE="'+AX+'" AND CSP_SECTION="'+Section+'"' ;
      end else
      begin
        SQL:='UPDATE CSECTION SET  '
        +'CSP_TOTALDEBIT=CSP_TOTALDEBIT-'+StrfPoint(XD)+', CSP_TOTALCREDIT=CSP_TOTALCREDIT-'+StrfPoint(XC)+', '
        +'CSP_TOTDEBE=CSP_TOTDEBE-'+StrfPoint(DE)+', CSP_TOTCREE=CSP_TOTCREE-'+StrfPoint(CE)+', '
        +'CSP_TOTDEBS=CSP_TOTDEBS-'+Strfpoint(DS)+', CSP_TOTCRES=CSP_TOTCRES-'+StrfPoint(CS)+', '
        +'CSP_TOTDEBP=CSP_TOTDEBP-'+StrfPoint(DP)+', CSP_TOTCREP=CSP_TOTCREP-'+StrfPoint(CP)+
        ' WHERE CSP_AXE="'+AX+'" AND CSP_SECTION="'+Section+'"' ;
      end;
      if ExecuteSQL(SQL)<>1 then
      begin
        LastErrorMessage := 'Erreur 1200102';
        V_PGI.IoError:=oeSaisie ;
      end;
  end else
  begin
    if not EstTablePartagee( 'SECTION' ) then
    begin
      if Sens = '+' then
      begin
        SQL:='UPDATE SECTION SET S_DEBITDERNMVT='+StrfPoint(XD)+',  S_CREDITDERNMVT='+StrfPoint(XC)+', '
        +'S_DATEDERNMVT="'+UsDateTime(DateC)+'", S_NUMDERNMVT='+IntToStr(Num)+', '
        +'S_TOTALDEBIT=S_TOTALDEBIT+'+StrfPoint(XD)+', S_TOTALCREDIT=S_TOTALCREDIT+'+StrfPoint(XC)+', '
        +'S_TOTDEBE=S_TOTDEBE+'+StrfPoint(DE)+', S_TOTCREE=S_TOTCREE+'+StrfPoint(CE)+', '
        +'S_TOTDEBS=S_TOTDEBS+'+Strfpoint(DS)+', S_TOTCRES=S_TOTCRES+'+StrfPoint(CS)+', '
        +'S_TOTDEBP=S_TOTDEBP+'+StrfPoint(DP)+', S_TOTCREP=S_TOTCREP+'+StrfPoint(CP)+
        ' WHERE S_AXE="'+AX+'" AND S_SECTION="'+Section+'"' ;
      end else
      begin
        SQL:='UPDATE SECTION SET  '
        +'S_TOTALDEBIT=S_TOTALDEBIT-'+StrfPoint(XD)+', S_TOTALCREDIT=S_TOTALCREDIT-'+StrfPoint(XC)+', '
        +'S_TOTDEBE=S_TOTDEBE-'+StrfPoint(DE)+', S_TOTCREE=S_TOTCREE-'+StrfPoint(CE)+', '
        +'S_TOTDEBS=S_TOTDEBS-'+Strfpoint(DS)+', S_TOTCRES=S_TOTCRES-'+StrfPoint(CS)+', '
        +'S_TOTDEBP=S_TOTDEBP-'+StrfPoint(DP)+', S_TOTCREP=S_TOTCREP-'+StrfPoint(CP)+
        ' WHERE S_AXE="'+AX+'" AND S_SECTION="'+Section+'"' ;
      end;
      if ExecuteSQL(SQL)<>1 then
      begin
        LastErrorMessage := 'Erreur 1200102';
        V_PGI.IoError:=oeSaisie ;
      end;
    end;
  end;
end;

procedure TOF_BTGENODANAL.DeduitCumul (TOBAA : TOB);
begin
	AjouteCumul(TOBAA,'-');
end;

procedure TOF_BTGENODANAL.DeleteOldAnalytiques;

	procedure  ReinitCumuls(TOBECRO : TOB);
  var indice : Integer;
  		TOBAA : TOB;
  begin
		for Indice:= 0 to TOBECRO.detail.count -1 do
    begin
			TOBAA := TOBECRO.detail[Indice];
      DeduitCumul (TOBAA);
      if V_PGI.IOError <> OeOk then break;
    end;
  end;


var TOBAA,TOBECRO : TOB;
		Indice : Integer;
    RefAna : TRefAna;
    Req : string;
    QQ : TQuery;
begin
  if TOBECRANAMOIS_O.Detail.count = 0 then exit;
	TOBECRO := TOB.Create('LES ANALYTIQUES',nil,-1);
  TRY
    for Indice := 0 to TOBECRANAMOIS_O.Detail.count -1 do
    begin
      TOBECRO.ClearDetail;
      TOBAA := TOBECRANAMOIS_O.detail[Indice];
      // Récupération des lots d'écritures
      DecodeRefAna(TOBAA.GetString('BA1_REFANALYTIQUE'),RefAna);
			Req := 'SELECT * FROM ANALYTIQ WHERE '+
      			 'Y_ENTITY="'+IntToStr(RefAna.Entity)+'" AND '+
      			 'Y_JOURNAL="'+RefAna.Journal+'" AND '+
      			 'Y_EXERCICE="'+RefAna.Exercice+'" AND '+
      			 'Y_NUMEROPIECE='+IntToStr(RefAna.Piece);
      QQ := OpenSQL(Req,True,-1,'',true);
      //
      TOBECRO.LoadDetailDB('ANALYTIQ','','',QQ,false);
      ReinitCumuls(TOBECRO);
      //
      if V_PGI.IOError <> OeOk then break;
			Req := 'DELETE FROM ANALYTIQ WHERE '+
      			 'Y_ENTITY="'+IntToStr(RefAna.Entity)+'" AND '+
      			 'Y_JOURNAL="'+RefAna.Journal+'" AND '+
      			 'Y_EXERCICE="'+RefAna.Exercice+'" AND '+
      			 'Y_NUMEROPIECE='+IntToStr(RefAna.Piece);
      if ExecuteSQL(Req) < 1 then
      begin
      	LastErrorMessage := 'Erreur 1200103';
      	V_PGI.IOError := oeUnknown;
        break;
      end;
    end;
  FINALLY
  	TOBECRO.Free;
  END;
end;

procedure TOF_BTGENODANAL.GetParamsOdPrec;
var TheDate : TDateTime;
		MMLoc : string;
    QQ : TQuery;
begin
  TOBSTATUSANALMOIS.InitValeurs(false);
  //
	TheDate := StrToDate(THEdit(GetControl('EDATE')).Text);
  MMLoc := ConstitueMois (TheDate);
  QQ := OpenSQL('SELECT * FROM BSTATUSANALMOIS WHERE BAS_MOIS="'+MMLoc+'"',True,1,'',True);
  if not QQ.Eof then
  begin
  	TOBSTATUSANALMOIS.SelectDB('',QQ);
  end;
  ferme (QQ);
end;

function TOF_BTGENODANAL.AjouteRejets(TOBAREJ: TOB) : TOB;
var TOBREJ: TOB;
begin
	TOBREJ := TOBREJETES.FindFirst(['ERREUR'],[TOBAREJ.GetInteger('ERREUR')],false);
  if TOBREJ = nil then
  begin
    TOBREJ := TOB.Create('UNE ERREUR',TOBREJETES,-1);
    TOBREJ.AddChampSupValeur('ERREUR',TOBAREJ.GetInteger('ERREUR'));
  end;
  Result := TOBREJ;
end;

function TOF_BTGENODANAL.GenerationOK: boolean;
var Indice : Integer;
begin
	Result := True;
  if TOBREJETES.Detail.count = 0 then Exit;
  for Indice := 0 to TOBREJETES.Detail.Count -1 do
  begin
    if TOBREJETES.detail[Indice].GetInteger('ERREUR')<> 1 then
    begin
      Result := false;
      Exit;
    end;
  end;
end;

procedure TOF_BTGENODANAL.ConstitueRapport(TOBREJETES: TOB; OneRapport: TGestionRapport);
var Indice,II : integer;
		texte : string;
begin
	for Indice := 0 to TOBREJETES.Detail.count -1 do
  begin
    if TOBREJETES.detail[Indice].GetInteger('ERREUR') <> 1 then
    begin
			for II := 0 to TOBREJETES.detail[Indice].Detail.count -1 do
      begin
				if TOBREJETES.detail[Indice].Detail[II].GetInteger('ERREUR') = 5 then
        begin
					texte := 'Montant nul pour ';
				end else if TOBREJETES.detail[Indice].Detail[II].GetInteger('ERREUR') = 3 then
        begin
					texte := 'Section générale non définie pour ';
        end else if TOBREJETES.detail[Indice].Detail[II].GetInteger('ERREUR') = 2 then
        begin
					texte := 'Compte comptable général non défini pour ';
        end else if TOBREJETES.detail[Indice].Detail[II].GetInteger('ERREUR') = 4 then
        begin
					texte := 'Section analytique non applicable pour ';
        end;
        if TOBREJETES.detail[Indice].Detail[II].GetString('GA_TYPEARTICLE')='FRA' then
        begin
          texte := texte + 'l''article de frais ';
        end else if TOBREJETES.detail[Indice].Detail[II].GetString('GA_TYPEARTICLE')='PRE' then
        begin
          texte := texte + 'la prestation ';
        end else if TOBREJETES.detail[Indice].Detail[II].GetString('GA_TYPEARTICLE')='MAR' then
        begin
          texte := texte + 'le matériel ';
        end;
        texte := texte + TOBREJETES.detail[Indice].Detail[II].GetString('GA_CODEARTICLE');
        texte := texte + ' --> ';
        //
        if (TOBREJETES.detail[Indice].Detail[II].GetInteger('ERREUR') <> 3) then
        begin
          if TOBREJETES.detail[Indice].Detail[II].GetString('NATUREPRES')<>'---' then
          begin
            texte := texte + 'Nature de prestation '+TOBREJETES.detail[Indice].Detail[II].GetString('NATUREPRES')+' ';
          end;
          texte := texte + 'Code resssource '+TOBREJETES.detail[Indice].Detail[II].GetString('BCO_RESSOURCE');
        end else
        begin
          texte := Texte + ' Section analytique Dépot';
        end;
        //
      	OneRapport.SauveLigMemo(texte);
      end;
    end;
  end;
end;

procedure TOF_BTGENODANAL.EcritOdAnalytiques;

	procedure  EcritDansJournal(TOBTT : TOB);
  var TOBLL : TOB;
  		II : Integer;
      SoldeDev,Solde : Double;
  begin
		for II := 0 to TOBTT.detail.count -1 do
    begin
      TOBLL := TOBTT.detail[II];
      TOBLL.ClearDetail;
      SoldeDev := Arrondi(TOBLL.GetDouble('Y_DEBITDEV')- TOBLL.GetDouble('Y_CREDITDEV'),V_PGI.okdecV);
      Solde := Arrondi(TOBLL.GetDouble('Y_DEBIT')- TOBLL.GetDouble('Y_CREDIT'),V_PGI.okdecV);
      if Solde = 0 then continue;
      if SoldeDev > 0 then
      begin
        TOBLL.SetDouble('Y_DEBITDEV',SoldeDev);
        TOBLL.SetDouble('Y_DEBIT',Solde);
        TOBLL.SetDouble('Y_CREDITDEV',0);
        TOBLL.SetDouble('Y_CREDIT',0);
      end else
      begin
        TOBLL.SetDouble('Y_CREDITDEV',Abs(SoldeDev));
        TOBLL.SetDouble('Y_CREDIT',Abs(Solde));
        TOBLL.SetDouble('Y_DEBITDEV',0);
        TOBLL.SetDouble('Y_DEBIT',0);
      end;
      if not TOBLL.InsertDB(nil,false) then
      begin
    		LastErrorMessage := 'Erreur 1200105';
        V_PGI.IOError := oeUnknown;
        break;
      end;
    end;
  end;


var II,JJ : Integer;
		TOBUT : TOB;
begin
	for II := 0 to TOBResult.Detail.Count -1 do
  begin
		TOBUT := TOBResult.Detail[II];
    if TOBUT.Detail.Count = 0 then continue;
    // on va traiter un lot d'écriture sur un compte général
    for JJ := 0 to TOBUT.Detail.count -1 do
    begin
    	EcritDansJournal(TOBUT.detail[JJ]);
      if V_PGI.IOError <> oeOk then Exit;
    end;
  end;

end;

procedure TOF_BTGENODANAL.GereStatusMois;
var TheMois : string;
		TOBSTATUS : TOB;
begin
  TheMois := ConstitueMois(StrToDate(THEdit(GetControl('EDATE')).Text));
  TOBSTATUS := TOB.Create ('BSTATUSANALMOIS',nil,-1);
  TRY
    TOBSTATUS.SetString('BAS_MOIS',TheMois);
    TOBSTATUS.SetString('BAS_STATUSGEN','DEJ');
    if not TOBSTATus.InsertOrUpdateDB(false) then
    begin
    	LastErrorMessage := 'Erreur 1200107';
    	V_PGI.IOError := oeSaisie;
    end;
  FINALLY
  	TOBSTATUS.Free;
  end;
end ;

Function TOF_BTGENODANAL.CreatSectionGenere(Axe,Code,libelle : string) : Boolean;
BEGIN
  Result := CreerSectionVolee ( Code,Libelle,Axe );
END;

procedure TOF_BTGENODANAL.DeleteStatusMois;
var SQl : string;
begin
  SQl := 'DELETE FROM BSTATUSANALMOIS WHERE BAS_MOIS="'+ConstitueMois(StrToDate(THEdit(GetControl('EDATE')).Text))+'"';
  if ExecuteSQL(Sql)<> 1 then V_PGI.IOError := oeUnknown;
end;

procedure TOF_BTGENODANAL.Deletemouvements ;
begin
  ReinitAll;
  GetOldAna;
  DeleteOldAnalytiques;
  // Supprime les anciens mouvements du mois
  if V_PGI.IOError = OeOk then DeleteOldMvtsMois;
  if V_PGI.IOError = OeOk then DeleteStatusMois;
end;

procedure TOF_BTGENODANAL.ClickDelete(Sender: TObject);
var result : TIOErr;
begin
  if PgiAsk('Etes-vous sur de supprimer les écritures de ce mois ?') <> mryes then exit;
  result := TRANSACTIONS (Deletemouvements,0);
  if result <> OeOk then
  begin
    PGIError('ABANDON : '+LastErrorMessage);
  end else
  begin
    PGIInfo('Traitement terminé.');
  end;
  SetDate;
end;

procedure TOF_BTGENODANAL.ReinitAll;
begin
	TOBECRANAMOIS_O.ClearDetail;
  TOBECRANAMOIS.ClearDetail;
  TOBResult.ClearDetail;
  TOBREJETES.ClearDetail;
end;

Initialization
  registerclasses ( [ TOF_BTGENODANAL ] ) ;
end.

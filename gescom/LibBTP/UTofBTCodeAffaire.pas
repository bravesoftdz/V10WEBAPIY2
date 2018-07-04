unit UTofBTCodeAffaire;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
{$IFDEF EAGLCLIENT}
    MaineAGL,emul,
{$ELSE}
    {$IFNDEF ERADIO}
    Fe_Main,
    {$ENDIF !ERADIO}
    {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
    HDB,
    Mul,
    UtilGc,
{$ENDIF}
		M3Fp,
    HCtrls,
    HEnt1,
    HMsgBox,
    UTOF,
    AffaireUtil,
    Dicobtp,
    SaisUtil,
    EntGC,
    utofAfBaseCodeAffaire,
    HTB97,
    ParamSoc,
    Hpanel,
    UtilPGI,
    Facture,
    AGLInitBTP,
    UTOB,
    Menus,
    AglInit,
    UentCommun,
    UdefExport,
//		UDossierVP,
    HRichOle,
    Ent1,
    UtilsRapport,
    BTREAJUSTEANA_TOF
    ;
// Ajout FV

Type
    RParamDoc = record
      Cledoc : R_CLEDOC;
      Affaire : string;
    end;

     TOF_BTCODEAFFAIRE = Class (TOF_AFBASECODEAFFAIRE)
     private

        isMemoire : boolean;
        BInsert   : TToolBarButton97;
        BInsert1  : TToolBarButton97;
        BTAnalyse : TToolBarButton97;
        BTDuplic  : TToolBarButton97;
        BTDuplic1 : TToolBaRButton97;

        MnZTiers  : TMenuItem;
        MnZSuiviCde  : TMenuItem;
        MnZAppels : TMenuItem;

        MnZInsertFBT : TMenuItem;
        MnZInsertABT : TMenuItem;
        //
        MnZInsertFBP : TMenuItem;
        MnZInsertABP : TMenuItem;
        MnZDuplicFBP : TMenuItem;
        MnZDuplicABP : TMenuItem;
        //
        XX_WHERE     : THEdit;
        TypePiece    : string;


{$IFDEF EAGLCLIENT}
        Fliste	: THGrid;
{$ELSE}
        Fliste	: THDbGrid;
{$ENDIF}
        ReajusteSit  : boolean;
        Consultation : Boolean;
        ModifDateSit : boolean;
        Stock        : Boolean;
        GensurDevisNonACPT  : Boolean;
        RapportGen      : TGestionRapport;
        StatutAff       : string;
        ReajusteAnal    : Boolean;
        SaisieAvancPoc  : boolean;
        SaisieInterdite : Boolean;

        Recup           : TCheckBox;
        EtatAffaire     : THValCombobox;

        procedure RenommageEnteteColonnes(TypePiece : string);
        procedure ControleChamp(Champ, Valeur: String);
        //
//uniquement en line
{*
		    procedure GestionMulDevis_S1(StArg: String);
        procedure GestionLine;
		    procedure FlisteDblClickLine(Sender: TObject);
*}
		    procedure FlisteDblClick(Sender: TObject);
//
		    procedure FlisteRowEnter(Sender: TObject);
        procedure ExportExcelClick (Sender : TObject);
				procedure ExportXMLClick (Sender : Tobject);
        procedure ExportDocMandateClick (Sender : Tobject);
        procedure ExportDocssTraiteClick (Sender : Tobject);
        Procedure ExportDocIntervenantClick(Sender : Tobject);
				procedure MnExpCotraitXml (Sender : Tobject);
        Procedure BInsertClick(Sender : TObject);
        procedure GestionMulDevisSuitePlace(StArg: String);
        procedure GestionSuitePlace;
				procedure ModificationDateSituation (TheChaine : string);
        procedure MnZTiersClick(Sender: TObject);
        procedure MnZSuiviCdeClick(Sender: TObject);
        procedure MnZAppelsClick(Sender: TObject);
        Procedure BTAnalyseClick(Sender: Tobject);
        procedure MnOuvPlatClick (Sender : Tobject);

        Procedure BTDuplicationClick(Sender: TObject);
        procedure MnInsertFBT(Sender: TObject);
        procedure MnInsertABT(Sender: TObject);
        //
        procedure MnInsertFBP(Sender: TObject);
        procedure MnInsertABP(Sender: TObject);
        procedure MnDuplicABP(Sender: TObject);
        procedure MnDuplicFBP(Sender: TObject);
        //
        procedure AppelDuplication(TypeFacture : String);
        //
        function  ChargeCleDoc: RParamDoc;
        function  ControleFournisseur(CodeFrs: String; Cledoc: R_cledoc): TTypeExport;
        procedure BTVisualPClick(Sender: Tobject);
        procedure TRVISUALPClick(Sender: TObject);
        procedure RefuseDevis ;
    		procedure RefuseMultiDevis(TOBPieces: TOB);
        procedure RechRepresentant;
        procedure ValidelesfacturesProv;
     public
        procedure OnArgument(stArgument : String ) ; override ;
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
        procedure Onload ;override;
        procedure OnUpdate ;override;
        procedure OnClose                  ; override ;
     END;
implementation

uses Grids,
     USpecifLSE,
     UfactExportXLS,
     UcoTraitance,
     UtofBtanalDev,
     TiersUtil,
     BTPUtil,
     UtilTobPiece, DB,Vierge,FactTOB,FactUtil,FactSituations,UFonctionsCBP,
     UtilsOuvragesPlat
     ;

procedure TOF_BTCODEAFFAIRE.OnArgument(stArgument : String );
var Bacomptes : TToolbarButton97;
    critere,valmul,champmul : string;
    x : integer;
    CC : THValCOmboBox;
Begin
fMulDeTraitement  := true;
Inherited;
  fTableName  := 'PIECE';
  XX_WHERE    := THEdit(Getcontrol('XX_WHERE'));
  if XX_WHERE <> nil then XX_WHERE.text:= '';

  SaisieAvancPoc := false;
  ReajusteSit := false;

  IsMemoire    := false;
  ModifDateSit := false;
  Stock        := false;
  GensurDevisNonACPT  := GetParamSocSecur('SO_GENCESURDEVNACPT', False);
  SaisieInterdite     := GetParamSocSecur('SO_SAISIEDIRECTEINTERDITE', False);
  //
  if Assigned(GetControl('RECUP')) then Recup := TCheckBox(GetControl('RECUP'));
  if Assigned(GetControl('AFF_ETATAFFAIRE')) then EtatAffaire := THValComboBox(GetControl('AFF_ETATAFFAIRE'));
  //
  Repeat
    Critere:=uppercase(ReadTokenSt(stArgument)) ;
    valMul := '';
    if Critere<>'' then
        begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           end
        else
           ChampMul := Critere;
        ControleChamp(ChampMul, ValMul);
        end;
  until  Critere='';

  UpdateCaption(Ecran);

  if GetParamSocSecur('SO_CHANTIEROBLIGATOIRE', False) then
  begin
    if TypePiece = 'CF' then
    begin
      IF not Stock then
         XX_WHERE.text := XX_WHERE.Text + ' AND GP_AFFAIRE = ""'
      else
         XX_WHERE.text := XX_WHERE.text + ' AND GP_AFFAIRE <> ""';
    end;
  end;

	 if TToolbarButton97 (GetControl('BACOMPTES')) <> nil then
   	  Begin
	    BAcomptes :=TToolbarButton97 (GetControl('BACOMPTES'));
   	  if GetParamSoc ('SO_ACOMPTEOBLIG') then
         Bacomptes.down := true
		    else
         Bacomptes.down := false;
	    end;

   if TToolbarButton97 (GetControl('BInsert')) <> nil then
      Begin
      BInsert := TToolbarButton97 (GetControl('BINSERT'));
      BInsert.OnClick := BInsertClick;
      end;

  if TToolbarButton97 (GetControl('BInsert1')) <> nil then
  Begin
    BInsert1 := TToolbarButton97 (GetControl('BINSERT1'));
  end;

	// gestion des etablissements
	CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ;
  if CC<>Nil then PositionneEtabUser(CC) ;
  //
  CC:=THValComboBox(GetControl('GP_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

  SetControlProperty('BSELECTCON', 'Visible', VH_GC.BTSeriaContrat);
  SetControlProperty('TGPAFFAIRE1', 'Visible', VH_GC.BTSeriaContrat);
  SetControlProperty('AFF_CONTRAT1', 'Visible', VH_GC.BTSeriaContrat);
  SetControlProperty('AFF_CONTRAT2', 'Visible', VH_GC.BTSeriaContrat);
  SetControlProperty('AFF_CONTRAT3', 'Visible', VH_GC.BTSeriaContrat);
  SetControlProperty('CON_AVENANT', 'Visible', VH_GC.BTSeriaContrat);

  if Assigned(GetControl('MnZTiers')) then
  begin
    MnZTiers := TMenuItem(GetControl('MnZTiers'));
    MnZTiers.OnClick := MnZTiersClick;
  end;

  if Assigned(GetControl('MnZSuiviCde')) then
  begin
    MnZSuiviCde := TMenuItem(GetControl('MnZSuiviCde'));
    //
    if TypePiece <> 'CF' then
      MnZSuiviCde.Visible := False
    else
    begin
      MnZSuiviCde.Visible := ExJaiLeDroitConcept(TConcept(bt530),False);
      MnZSuiviCde.OnClick := MnZSuiviCdeClick;
    end;
  end;

  if assigned(GetControl('MnOuvPlat')) then
  begin
    ThmenuItem(GetCOntrol('MnOuvPlat')).onclick := MnOuvPlatClick;
  end;

  if Assigned(GetControl('MnZAppels')) then
  begin
    MnZAppels := TmenuItem(GetControl('MnZAppels'));
    MnZAppels.OnClick := MnZAppelsClick;
  end;

  if Assigned(GetControl('B_ANALYSE')) then
  begin
    BTAnalyse := TToolbarButton97(GetControl('B_ANALYSe'));
    BTAnalyse.OnClick := BTAnalyseClick;
  end;

  if assigned(GetControl('B_DUPLICATION')) then
  begin
    BTDuplic  := TToolbarButton97(GetControl('B_DUPLICATION'));
    if BTDuplic <> nil then if BTDuplic.DropdownMenu = nil then BTDuplic.OnClick := BTDuplicationClick;
  end;

  if assigned(GetControl('B_DUPLICATION1')) then
  begin
    BTDuplic1 := TToolBarButton97(Getcontrol('B_DUPLICATION1'));
    if BTDuplic1 <> nil Then BTDuplic1.OnClick := BTDuplicationClick;
  end;

  if Assigned(GetControl('PInsertFBT')) then
  begin
    MnZInsertFBT := TMenuItem(GetControl('PInsertFBT'));
    if (Not jaiLeDroitTag(145421)) then MnZInsertFBT.Enabled := false;
    MnZInsertFBT.OnClick := MnInsertFBT;
  end;

  if Assigned(GetControl('PInsertABT')) then
  begin
    MnZInsertABT := TmenuItem(GetControl('PInsertABT'));
    if (Not jaiLeDroitTag(145431)) then MnZInsertABT.Enabled := false;
    MnZInsertABT.OnClick := MnInsertABT;
  end;

  if assigned(GetControl('TRVISUALP')) then
  begin
    TCheckBox(GetControl('TRVISUALP')).OnClick := TRVISUALPClick;
    THedit(GetControl('XX_WHERE2')).Text := ' AND GP_IDENTIFIANTWOT<=0';
  end;

// NEW
  if Assigned(GetControl('PInsertFBP')) then
  begin
    MnZInsertFBP := TMenuItem(GetControl('PInsertFBP'));
    if (Not jaiLeDroitTag(145421)) then MnZInsertFBP.Enabled := false;
    MnZInsertFBP.OnClick := MnInsertFBP;
  end;

  if Assigned(GetControl('PInsertABP')) then
  begin
    MnZInsertABP := TmenuItem(GetControl('PInsertABP'));
    if (Not jaiLeDroitTag(145431)) then MnZInsertABP.Enabled := false;
    MnZInsertABP.OnClick := MnInsertABP;
  end;

  if Assigned(GetControl('PDuplicFBP')) then
  begin
    MnZDuplicFBP := TMenuItem(GetControl('PDuplicFBP'));
    if (Not jaiLeDroitTag(145421)) then MnZDuplicFBP.Enabled := false;
    MnZDuplicFBP.OnClick := MnDuplicFBP;
  end;

  if Assigned(GetControl('PDuplicABP')) then
  begin
    MnZDuplicABP := TmenuItem(GetControl('PDuplicABP'));
    if (Not jaiLeDroitTag(145431)) then MnZDuplicABP.Enabled := false;
    MnZDuplicABP.OnClick := MnDuplicABP;
  end;
// -----

  SetControlProperty('BEFFACECON', 'Visible', VH_GC.BTSeriaContrat);

//uniquement en line
//	GestionLine;
	GestionSuitePlace;
//
  RapportGen   := TGestionRapport.Create(TFMul(Ecran));
  with RapportGen do
  begin
    Titre   := 'Validation des factures provisoires';
    Affiche := True;
    Close   := True;
    Sauve   := False;
    Print   := False;
    InitRapport;
    Visible := false;
    PosLeft := Round(TFMul(Ecran).Width / 1.5);
    PosTop  := Round(TFMul(Ecran).Top);
  end;

  if pos('BTCONSULTPIEC_MUL',ecran.name) > 0 then
  begin
    SetControlVisible('GP_NATUREPIECEG', True);
    SetControlEnabled('GP_NATUREPIECEG', True);
    SetControlProperty('GP_NATUREPIECEG', 'Plus', ' AND GPP_NATUREPIECEG IN ("CBT", "BFC", "LBT")');
  end;

End;

{$IFNDEF LINE}
function IsDevisIntModifiable (cledoc : R_CLEDOC) : boolean;
var QQ : TQuery;
		Sql : string;
begin
	result := false;
	SQl := 'SELECT AFF_RESPONSABLE, AFF_ETATAFFAIRE FROM AFFAIRE WHERE '+
  			 'AFF_AFFAIRE=(SELECT GP_AFFAIRE FROM PIECE WHERE '+
         'GP_NATUREPIECEG="'+Cledoc.NaturePiece +'" AND GP_SOUCHE="'+Cledoc.Souche +'" AND '+
         'GP_NUMERO='+IntToStr(Cledoc.NumeroPiece )+' AND GP_INDICEG='+IntToStr(Cledoc.Indice)+')';
  QQ := OpenSql (Sql,true,1,'',true);
  TRY
    if not QQ.eof then
    begin
      if Pos(QQ.FindField('AFF_ETATAFFAIRE').AsString,'ACD;ECO;REA')> 0 then
        result := true
      else if QQ.FindField('AFF_RESPONSABLE').AsString <> '' then
        Result := True;        
    end;
  FINALLY
  	Ferme (QQ);
  END;
end;
{$ENDIF}

procedure TOF_BTCODEAFFAIRE.FlisteRowEnter(Sender: TObject);
var mandataire  : string;
    SousTraite  : Integer;
begin

  if pos('BTDEVIS_MUL',ecran.name) > 0 then
  begin
		if TToolbarButton97(getControl('BEXPORTS')) <> nil then
    	TToolbarButton97 (getControl('BEXPORTS')).visible := true;

    if TmenuItem(getControl('MnSepCot')) <> nil then
    begin
      TmenuItem(getControl('MnSepCot')).visible := false;
    end;
    //export document mandataire
    if TmenuItem(getControl('MnExportMandate')) <> nil then
    begin
      TmenuItem(getControl('MnExportMandate')).visible := false;
    end;
    //Export document sous-traitant
    if TmenuItem(getControl('MnExportSsTraite')) <> nil then
    begin
      TmenuItem(getControl('MnExportSsTraite')).visible := false;
    end;
    //export document intervenant (cotraitant & sous-traitant)
    if TmenuItem(getControl('MnExportIntervenant')) <> nil then
    begin
      TmenuItem(getControl('MnExportIntervenant')).visible := false;
    end;
    //
    if TmenuItem(getControl('MnExpCotraitXml')) <> nil then
    begin
      TmenuItem(getControl('MnExpCotraitXml')).visible := false;
    end;

 		mandataire := '';
		sousTraite := 0;

{$IFDEF EAGLCLIENT}
    TFMul(ecran).Q.TQ.Seek(TFMul(ecran).FListe.Row-1) ;
    mandataire := TFMul(ecran).Q.FindField('AFF_MANDATAIRE').AsString;
    SousTraite := TFMul(ecran).Q.FindField('SSTRAITE').AsInteger;
{$ELSE}
		if Fliste.DataSource.DataSet.FindField('AFF_MANDATAIRE') <> nil then
    begin
    	mandataire := Fliste.datasource.dataset.FindField('AFF_MANDATAIRE').AsString;
    end;
  	if Fliste.DataSource.DataSet.FindField('SSTRAITE') <> nil then
    begin
      SousTraite := Fliste.datasource.dataset.FindField('SSTRAITE').AsInteger;
    end;
{$ENDIF}

	  if (Mandataire = 'X') or (SousTraite > 0) then
    begin
    	if TmenuItem(getControl('MnSepCot')) <> nil then
    		TmenuItem(getControl('MnSepCot')).visible := true;
    	if TmenuItem(getControl('MnExportIntervenant')) <> nil then
    		TmenuItem(getControl('MnExportIntervenant')).visible := true;
      if TmenuItem(getControl('MnExpCotraitXml')) <> nil then
        TmenuItem(getControl('MnExpCotraitXml')).visible := true;
    end;
  end;

end;

procedure TOF_BTCODEAFFAIRE.ExportExcelClick (Sender : TObject);
var ParamDoc : RParamDoc ;
		ExportXls : TexportDocXls;
begin

  if pos(ecran.name,'BTDEVIS_MUL;BTDEVISSTRAITE;BTDEVISCOTRAI_MUL') > 0 then
  begin
    ParamDoc := chargeCleDoc;
    //
    ExportXls := TexportDocXls.create;
    ExportXls.ModeExport  := TmeExcel;
    ExportXls.TypeExport := TteDevis;
    ExportXls.CleDoc := ParamDoc.cledoc;
    ExportXls.Affaire := ParamDoc.Affaire;
    ExportXls.ModeAction := XmaDocument;
    ExportXls.Genere;
    ExportXls.free;
  end;

end;

procedure TOF_BTCODEAFFAIRE.ExportXmlClick (Sender : Tobject);
var ParamDoc : RParamDoc ;
		ExportXls : TexportDocXls;
begin

  if pos(ecran.name,'BTDEVIS_MUL;BTDEVISSTRAITE;BTDEVISCOTRAI_MUL') > 0 then
  begin
    ParamDoc := chargeCleDoc;
    //
    ExportXls := TexportDocXls.create;
    ExportXls.ModeExport  := TmeXml;
    ExportXls.TypeExport := TteDevis;
    ExportXls.CleDoc := ParamDoc.cledoc;
    ExportXls.Affaire := ParamDoc.Affaire;
    ExportXls.ModeAction := XmaDocument;
    ExportXls.Genere;
    ExportXls.free;
  end;

end;


procedure TOF_BTCODEAFFAIRE.ExportDocMandateClick (Sender : Tobject);
var Paramdoc : RParamDoc;
		ExportXls   : TexportDocXls;
    fournisseur,LibFournisseur : string;
begin

  if pos(ecran.name,'BTDEVIS_MUL;BTDEVISSTRAITE;BTDEVISCOTRAI_MUL') > 0 then
  begin
    ParamDoc := chargeCleDoc;
  	fournisseur := AGLLanceFiche('BTP','BTMULAFFECTEXTE','','','COTRAITANCE;DOCUMENT='+EncodeCleDoc(Paramdoc.cledoc)+';SELECTION');
    if fournisseur <> '' then
    begin
      GetLibTiers('FOU', Fournisseur, LibFournisseur);
      ExportXls := TexportDocXls.create;
    	ExportXls.ModeExport  := TmeExcel;
      ExportXls.TypeExport := TteCotrait;
      ExportXls.CleDoc := Paramdoc.cledoc;
      ExportXls.ModeAction := XmaDocument;
      ExportXls.Fournisseur := Fournisseur;
      ExportXls.NomFrs :=  LibFournisseur;
      ExportXls.Affaire := ParamDoc.Affaire;
      ExportXls.Genere;
      ExportXls.free;
    end;
  end;
end;

procedure TOF_BTCODEAFFAIRE.ExportDocSstraiteClick (Sender : Tobject);
var ParamDoc : RParamDoc;
		ExportXls   : TexportDocXls;
    fournisseur,LibFournisseur : string;
begin

  if pos(ecran.name,'BTDEVIS_MUL;BTDEVISSTRAITE;BTDEVISCOTRAI_MUL') > 0 then
  begin
    ParamDoc := chargeCleDoc;
  	fournisseur := AGLLanceFiche('BTP','BTMULAFFECTEXTE','','','SOUSTRAITANCE;DOCUMENT='+EncodeCleDoc(ParamDoc.cledoc)+';SELECTION');
    if fournisseur <> '' then
    begin
      LibFournisseur := '';
      GetLibTiers('FOU', Fournisseur, LibFournisseur);

      ExportXls := TexportDocXls.create;
    	ExportXls.ModeExport  := TmeExcel;
      ExportXls.TypeExport := TteSousTrait;
      ExportXls.CleDoc := ParamDoc.cledoc;
      ExportXls.ModeAction := XmaDocument;
      ExportXls.Fournisseur := Fournisseur;
      ExportXls.NomFrs :=  LibFournisseur;
      ExportXls.Affaire := ParamDoc.Affaire;
      ExportXls.Genere;
      ExportXls.free;
    end;
  end;

end;

procedure TOF_BTCODEAFFAIRE.ExportDocIntervenantClick (Sender : Tobject);
var ParamDoc        : RParamDoc;
		ExportXls       : TexportDocXls;
    fournisseur     : String;
    LibFournisseur  : string;
begin

  if pos(ecran.name,'BTDEVIS_MUL') > 0 then
  begin
    ParamDoc := ChargeCleDoc;
  	fournisseur := AGLLanceFiche('BTP','BTMULAFFECTEXTE','','','DOCUMENT='+ EncodeCleDoc(ParamDoc.cledoc)+';SELECTION');
    if fournisseur <> '' then
    begin
    	GetLibTiers('FOU', Fournisseur, LibFournisseur);
      ExportXls := TexportDocXls.create;
    	ExportXls.ModeExport  := TmeExcel;
      //vérification si cotraitant ou Sous-traitant !!!
      ExportXls.TypeExport := ControleFournisseur(Fournisseur, ParamDoc.cledoc);
      ExportXls.CleDoc := ParamDoc.cledoc;
      ExportXls.ModeAction := XmaDocument;
      ExportXls.Fournisseur := Fournisseur;
      ExportXls.NomFrs :=  LibFournisseur;
      ExportXls.Affaire := ParamDoc.Affaire;
      ExportXls.Genere;
      ExportXls.free;
    end;
  end;

end;

function TOF_BTCODEAFFAIRE.ControleFournisseur(CodeFrs : String; Cledoc : R_cledoc) : TTypeExport;
var StSQL : String;
    TypeI : String;
    QQ    : TQuery;

begin

  StSql := 'SELECT BPE_TYPEINTERV FROM PIECETRAIT WHERE BPE_FOURNISSEUR="' + CodeFrs +'" AND ' + WherePiece(CleDoc, ttdpieceTrait, False);
  QQ := opensql(StSql, false);

  if QQ.eof then
    result := TteUndefined
  else
  begin
    TypeI := QQ.FindField('BPE_TYPEINTERV').AsString;
    if (TypeI = 'X00') OR (TypeI = 'X01') then
      result := TteCotrait
    else
      result := TteSousTrait;
  end;

  ferme(QQ);

end;

function TOF_BTCODEAFFAIRE.ChargeCleDoc: RParamDoc;
begin

{$IFDEF EAGLCLIENT}
    TFMul(ecran).Q.TQ.Seek(TFMul(ecran).FListe.Row-1) ;
    Result.cledoc.naturePiece := TFMul(ecran).Q.FindField('GP_NATUREPIECEG').AsString;
    Result.cledoc.DatePiece := TFMul(ecran).Q.FindField('GP_DATEPIECE').AsDateTime;
    Result.cledoc.souche := TFMul(ecran).Q.FindField('GP_SOUCHE').AsString;
    Result.cledoc.numeroPiece := StrToInt(TFMul(ecran).Q.FindField('GP_NUMERO').AsString);
    Result.cledoc.indice := StrToInt(TFMul(ecran).Q.FindField('GP_INDICEG').AsString);
    result.affaire := TFMul(ecran).Q.FindField('GP_AFFAIRE').AsString;
{$ELSE}
    Result.Cledoc.naturePiece := Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
    Result.Cledoc.DatePiece  := Fliste.datasource.dataset.FindField('GP_DATEPIECE').AsDateTime;
    Result.Cledoc.souche := Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
    Result.Cledoc.numeroPiece := StrToInt(Fliste.datasource.dataset.FindField('GP_NUMERO').AsString);
    Result.Cledoc.indice := StrToInt(Fliste.datasource.dataset.FindField('GP_INDICEG').AsString);
    if Assigned(Fliste.DataSource.FindComponent('GP_AFFAIRE')) then
      Result.Affaire          := Fliste.datasource.dataset.FindField('GP_AFFAIRE').AsString
    else
      Result.affaire          := '';

{$ENDIF}

end;

{$IFNDEF LINE}
procedure TOF_BTCODEAFFAIRE.FlisteDblClick(Sender: TObject);
Var ClePiece			: array [0..7] of Variant;
    TheAction			: ThEdit;
    TheModeSais   : TcheckBox;
    TheChaine     : string;
    ParamDoc      : RParamDoc;
begin

    if copy(ecran.name,1,14)='BTPIECENEG_MUL' then
    begin
      if recup.Checked then
      begin
        TFMul(Ecran).Retour := TFMul(ecran).Q.FindField('GP_NUMERO').AsString;
        if Assigned(TFMul(ecran).Q.FindField('GP_AFFAIRE')) then
          TFMul(Ecran).Retour := TFMul(Ecran).Retour + ';' + TFMul(ecran).Q.FindField('GP_AFFAIRE').AsString
        else
          TFMul(Ecran).Retour := TFMul(Ecran).Retour + ';';
        if Assigned(TFMul(ecran).Q.FindField('GP_TIERS'))   then
          TFMul(Ecran).Retour := TFMul(Ecran).Retour + ';' + TFMul(ecran).Q.FindField('GP_TIERS').AsString
        else
          TFMul(Ecran).Retour := TFMul(Ecran).Retour + ';';
        TFMUL(Ecran).Close;
        exit;
      end;
    end;

    if copy(ecran.name,1,14)='BTDEVISINT_MUL' then
    begin
      ParamDoc := ChargeCleDoc;
      //
      if (Consultation) or (not IsDevisIntModifiable (ParamDoc.cledoc)) then
      begin
      	SaisiePiece (ParamDoc.CleDoc, taconsult);
    		Refreshdb;
      end else
      begin
       	SaisiePiece(ParamDoc.Cledoc,taModif);
    		Refreshdb;
      end;
			exit;
    end;
    // -------------------------------------------------------------------
		TheAction := THEdit(GetControl('ZEACTION'));

		if pos(ecran.name,'BTDEVIS_MUL;BTDEVISSTRAITE;BTDEVISCOTRAI_MUL;BTPIECENEG_MUL;BTDEVISNEG_MUL') > 0 then
    begin
      ClePiece[0] := TFMul(Ecran).Name ;
{$IFDEF EAGLCLIENT}
			TFMul(ecran).Q.TQ.Seek(TFMul(ecran).FListe.Row-1) ;
      ClePiece[1] := TFMul(ecran).Q.FindField('GP_NATUREPIECEG').AsString;
      ClePiece[2] := DateTimeToStr(TFMul(ecran).Q.FindField('GP_DATEPIECE').AsDateTime);
      ClePiece[3] := TFMul(ecran).Q.FindField('GP_SOUCHE').AsString;
      ClePiece[4] := TFMul(ecran).Q.FindField('GP_NUMERO').AsString;
      ClePiece[5] := TFMul(ecran).Q.FindField('GP_INDICEG').AsString;
      ClePiece[6] := 'ACTION=CONSULTATION';
{$ELSE}
      ClePiece[1] := Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
      ClePiece[2] := DateTimeToStr(Fliste.datasource.dataset.FindField('GP_DATEPIECE').AsDateTime);
      ClePiece[3] := Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
      ClePiece[4] := Fliste.datasource.dataset.FindField('GP_NUMERO').AsString;
      ClePiece[5] := Fliste.datasource.dataset.FindField('GP_INDICEG').AsString;
      ClePiece[6] := 'ACTION=CONSULTATION';
{$ENDIF EAGLCLIENT}

    end
    else
    begin
      ClePiece[0] := TFMul(Ecran).Name ;

{$IFDEF EAGLCLIENT}
			TFMul(ecran).Q.TQ.Seek(TFMul(ecran).FListe.Row-1) ;
      ClePiece[1] := TFMul(ecran).Q.FindField('GP_NATUREPIECEG').AsString;
      ClePiece[2] := DateTimeToStr(TFMul(ecran).Q.FindField('GP_DATEPIECE').AsDateTime);
      ClePiece[3] := TFMul(ecran).Q.FindField('GP_SOUCHE').AsString;
      ClePiece[4] := TFMul(ecran).Q.FindField('GP_NUMERO').AsString;
      ClePiece[5] := TFMul(ecran).Q.FindField('GP_INDICEG').AsString;
      ClePiece[6] := TFMul(ecran).Q.FindField('GP_TIERS').AsString;
      if TFMul(ecran).Q.findField('GP_AFFAIRE') <> nil then
      begin
      	ClePiece[7] := TFMul(ecran).Q.FindField('GP_AFFAIRE').AsString;
      end;
{$ELSE}
      ClePiece[1] := Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
      ClePiece[2] := DateTimeToStr(Fliste.datasource.dataset.FindField('GP_DATEPIECE').AsDateTime);
      ClePiece[3] := Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
      ClePiece[4] := Fliste.datasource.dataset.FindField('GP_NUMERO').AsString;
      ClePiece[5] := Fliste.datasource.dataset.FindField('GP_INDICEG').AsString;
      ClePiece[6] := Fliste.datasource.dataset.FindField('GP_TIERS').AsString;
      if Fliste.datasource.DataSet.findField('GP_AFFAIRE') <> nil then
      begin
      	ClePiece[7] := Fliste.datasource.dataset.FindField('GP_AFFAIRE').AsString;
      end;
{$ENDIF EAGLCLIENT}
    end;

		TheChaine := ClePiece[1]+';'+ClePiece[2]+';'+ClePiece[3]+';'+ClePiece[4]+';'+ClePiece[5]+';';


    if ReajusteAnal then
    begin
      ReajusteAnalytique (TheChaine);
      Exit;
    end;

    if ReajusteSit then
    begin
      ReajusteSituationSais([TheChaine,'ACTION=CONSULTATION',BoolToStr_(Stock)],3);
    	Refreshdb;
      Exit;
    end;

    if Consultation Then
    Begin
      AppelPiece([TheChaine,'ACTION=CONSULTATION',BoolToStr_(Stock)],3);
    	Refreshdb;
      Exit;
    End;

    if ModifDateSit then
    begin
    	ModificationDateSituation (TheChaine);
    	Refreshdb;
//    TToolBarButton97(GetControl('Bcherche')).Click;
      exit;
    end;

    // generation de la prevision de chantier
 	  if (TheAction <> nil) and  (TheAction.Text = 'PLANNIFICATION') then
    Begin
      AglPlannificationChantier(ClePiece,5);
    	Refreshdb;
//		TToolBarButton97(GetControl('Bcherche')).Click;
      Exit;
    End;

    // generation de la contre etude
    if (TheAction <> nil) and  (TheAction.text = 'GENCONTRETU') then
    Begin
      AGLGenereContreEtude(ClePiece,5);
      Refreshdb;
      //		 TToolBarButton97(GetControl('Bcherche')).Click;
      Exit;
    End;

    if assigned(GetControl('AVANC')) then
    begin
      TheModeSais := TcheckBox(GetControl('AVANC'));
      if (TheModeSais <> nil) and  (TheModeSais.checked ) then
      Begin
        AglAvancementChantier (ClePiece,7);
        Refreshdb;
        //			 TToolBarButton97(GetControl('Bcherche')).Click;
        Exit;
      End;
    end;

    // Modification de devis
		AppelPiece([TheChaine,'ACTION=MODIFICATION',BoolToStr_(Stock)],3);
    //
    //if GetParamSocSecur('SO_REFRESHMUL', true) then TToolBarButton97(GetControl('Bcherche')).Click;
    Refreshdb;

end;
{$ENDIF}

Procedure TOF_BTCODEAFFAIRE.GestionSuitePlace;
begin

	// (copy(ecran.name,1,11)='BTDEVIS_MUL') then
  if pos(ecran.name,'BTDEVIS_MUL;BTDEVISSTRAITE;BTDEVISCOTRAI_MUL;BTPIECENEG_MUL;BTDEVISNEG_MUL') > 0 then
     Begin
     GestionMulDevisSuitePlace(StArgument);
     SetControlText('TGPAFFAIRE', 'Code Chantier');
     end
	Else if (copy(ecran.name,1,14)='BTDEVISINT_MUL') then
     Begin
     GestionMulDevisSuitePlace(StArgument);
     SetControlText('TGPAFFAIRE', 'Code Appel');
     end
	Else if (copy(ecran.name,1,11)='BTPIECE_MUL') then
     Begin
     GestionMulDevisSuitePlace(StArgument);
     SetControlText('TGPAFFAIRE', 'Code Chantier');
     end
  else if (copy(ecran.name,1,12)='BTACCDEV_MUL') then
     Begin
     SetControlText('TGPAFFAIRE', 'Code Chantier');
     end
  else if (copy(ecran.name,1,12)='BTREJDEVNEG_MUL') then
     Begin
     SetControlText('TAFF_AFFAIRE', 'Code Chantier');
     end
	else if (copy(ecran.name,1,12)='BTPREFAC_MUL') then
     Begin
     SetControlText('TAFF_AFFAIRE', 'Code Chantier');
     End;



//  else
//		TFMUL(Ecran).SetDbliste('BTMULPIECE_S1');
end;


Procedure TOF_BTCODEAFFAIRE.GestionMulDevisSuitePlace(StArg : String);
Begin

   SetControlText('TGPAFFAIRE', 'Code Chantier');

{$IFDEF EAGLCLIENT}
   if THGrid (GetControl('FLISTE')) <> nil then
     	Begin
	  	Fliste := THGrid (GetControl('FLISTE'));
//uniquement en line
//    	Fliste.OnDblClick := FlisteDblClickLine;
    	Fliste.OnDblClick := FlisteDblClick;
    	Fliste.OnRowEnter := FlisteRowEnter;
     	end;
{$ELSE}
   if THDbGrid (GetControl('FLISTE')) <> nil then
     	Begin
	  	Fliste := THDbGrid (GetControl('FLISTE'));
//uniquement en line
//    Fliste.OnDblClick := FlisteDblClickLine;
    	Fliste.OnDblClick := FlisteDblClick;
    	Fliste.OnRowEnter := FlisteRowEnter;
     	end;
{$ENDIF}
   if TmenuItem(getControl('MnExportExcel')) <> nil then
   begin
     TmenuItem(getControl('MnExportExcel')).OnClick := ExportExcelClick;
   end;
   if TmenuItem(getControl('MnExportXML')) <> nil then
   begin
     TmenuItem(getControl('MnExportXML')).OnClick := ExportXmlClick;
   end;
   if TmenuItem(getControl('MnExportXML')) <> nil then
   begin
     TmenuItem(getControl('MnExportXML')).OnClick := ExportXMLClick;
   end;
   if TmenuItem(getControl('MnExportMandate')) <> nil then
   begin
     TmenuItem(getControl('MnExportMandate')).OnClick := ExportDocMandateClick;
   end;
   if TmenuItem(getControl('MnExportSsTraite')) <> nil then
   begin
     TmenuItem(getControl('MnExportSsTraite')).OnClick := ExportDocSsTraiteClick;
   end;
   if TmenuItem(getControl('MnExportIntervenant')) <> nil then
   begin
     TmenuItem(getControl('MnExportIntervenant')).OnClick := ExportDocIntervenantClick;
   end;
   if TmenuItem(getControl('MnExpCotraitXml')) <> nil then
   begin
     TmenuItem(getControl('MnExpCotraitXml')).OnClick := MnExpCotraitXml;
   end;

   if GetArgumentValue(StArg, 'ACTION') = 'MODIFICATION'  then
	    begin
     if BTDuplic <> nil then BTDuplic.Visible := True;
	   BInsert1.visible := True;
		 BInsert.visible  := True;
  		Consultation := False;
	    end ;

	 if GetArgumentValue(StArg, 'ACTION') = 'CONSULTATION'  then
  		begin
  		Consultation := True;
	   BInsert1.visible := False;
		 BInsert.visible  := False;
     if BTDuplic <> nil then BTDuplic.Visible := False;
	    end ;

end;



//uniquement en line
{*
procedure TOF_BTCODEAFFAIRE.FlisteDblClickLine(Sender: TObject);
Var StArg 				: String;
		TypeRessource : String;
    Ressource 		: String;
    ClePiece			: array [0..6] of Variant;
    TheAction			: ThEdit;
begin

		TheAction := THEdit(GetControl('ZEACTION'));

    ClePiece[0] := TFMul(Ecran).Name ;
	  ClePiece[1] := Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
    ClePiece[2] := DateTimeToStr(Fliste.datasource.dataset.FindField('GP_DATEPIECE').AsDateTime);
    ClePiece[3] := Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
    ClePiece[4] := Fliste.datasource.dataset.FindField('GP_NUMERO').AsString;
    ClePiece[5] := Fliste.datasource.dataset.FindField('GP_INDICEG').AsString;
    ClePiece[6] := 'ACTION=CONSULTATION';

    if Consultation Then
       Begin
       AppelPiece(ClePiece,6);
       Exit;
       End;

    // generation de la prevision de chantier
 	  if TheAction.Text = 'PLANNIFICATION' then
       Begin
       AglPlannificationChantier(ClePiece,5);
    		Refreshdb;
			 //TToolBarButton97(GetControl('Bcherche')).Click;
       Exit;
       End;

    // generation de la contre etude
    if TheAction.text = 'GENCONTRETU' then
       Begin
       AGLGenereContreEtude(ClePiece,5);
    	Refreshdb;
			 //TToolBarButton97(GetControl('Bcherche')).Click;
       Exit;
       End;

    // Modification de devis
    if GetControlText('GP_VIVANTE')='X' then
	     AppelPiece(ClePiece, 5)
    else
	     AppelPiece(ClePiece, 6);
    //
    if GetParamSocSecur('SO_REFRESHMUL', true) then Refreshdb;
//TToolBarButton97(GetControl('Bcherche')).Click;

end;

Procedure TOF_BTCODEAFFAIRE.GestionLine;
var Bacomptes : TToolbarButton97;
    critere,valmul,champmul : string;
    x : integer;
    F : TFMul ;
    i : integer;
    CC : THValCOmboBox;
    TT : TToolbarButton97;
    TL : THlabel;
    TP : TTabSheet;
    TE : THEdit;
begin
  if cc <> nil then CC.visible := false;

  TT:=TToolBarButton97(GetControl('BINSERT1')) ; if TT<>Nil then TT.visible := false;
  TL:=THlabel(GetControl('TGP_ETABLISSEMENT')) ; if TL<>Nil then TL.visible := false;
  CC := THValCOmboBox(GetControl('AFF_GENERAUTO')); if CC <> nil then CC.plus := 'BTP" AND CO_CODE <> "DAC';
  CC := THValCOmboBox(GetControl('GP_REPRESENTANT')); if CC <> nil then CC.visible := false;
  TL:=THlabel(GetControl('TGP_REPRESENTANT')) ; if TL<>Nil then TL.visible := false;
  CC := THValCOmboBox(GetControl('GP_DEVISE')); if CC <> nil then CC.visible := false;
  TL:=THlabel(GetControl('TGP_DEVISE')) ; if TL<>Nil then TL.visible := false;
  CC := THValCOmboBox(GetControl('GP_LIBREPIECE1')); if CC <> nil then CC.visible := false;
  TL:=THlabel(GetControl('TGP_LIBREPIECE1')) ; if TL<>Nil then TL.visible := false;
  CC := THValCOmboBox(GetControl('GP_LIBREPIECE2')); if CC <> nil then CC.visible := false;
  TL:=THlabel(GetControl('TGP_LIBREPIECE2')) ; if TL<>Nil then TL.visible := false;
  CC := THValCOmboBox(GetControl('GP_LIBREPIECE3')); if CC <> nil then CC.visible := false;
  TL:=THlabel(GetControl('TGP_LIBREPIECE3')) ; if TL<>Nil then TL.visible := false;
  TT:=TToolBarButton97(GetControl('BACOMPTES')) ; if TT<>Nil then TT.visible := false;
  if Copy(Ecran.Name,1,12)='BTACCDEV_MUL' then
  begin
    TP := TTabSheet(GetControl('PCOMPLEMENT')) ; if TP<>Nil then TP.tabvisible := false;
  end;

  if BAcomptes <> nil then
     BEGIN
     BAcomptes.down := false;
     BAcomptes.visible := false;
     END;

  TE := THEdit(GetCOntrol('AFF_RESPONSABLE'));
  if TE <> nil then TE.visible := false;

  TL := THLabel(GetCOntrol('TAFF_RESPONSABLE'));
  if TL <> nil then TL.visible := false;

  if (copy(ecran.name,1,14)='BTDEVIS_MUL_S1') then
     GestionMulDevis_S1(StArgument)
	Else if (copy(ecran.name,1,11)='BTDEVIS_MUL') then
     Begin
     TFMUL(Ecran).SetDbliste('BTMULDEVIS_S1');
     SetControlText('TGPAFFAIRE', 'Code Chantier');
     end
	Else if (copy(ecran.name,1,11)='BTPIECE_MUL') then
     Begin
     TFMUL(Ecran).SetDbliste('BTMULDEVIS_S1');
     SetControlText('TGPAFFAIRE', 'Code Chantier');
     SetControlVisible('GP_VIVANTE', false);
     end
  else if (copy(ecran.name,1,12)='BTACCDEV_MUL') then
     Begin
     TFMUL(Ecran).SetDbliste('BTMULDEVIS_S1');
     SetControlText('TAFF_AFFAIRE', 'Code Chantier');
     end
	else if (copy(ecran.name,1,12)='BTPREFAC_MUL') then
     Begin
     TFMUL(Ecran).SetDbliste('BTPREFACIMM_S1');
     SetControlText('TAFF_AFFAIRE', 'Code Chantier');
     end;
//  else
//		TFMUL(Ecran).SetDbliste('BTMULPIECE_S1');
end;

Procedure TOF_BTCODEAFFAIRE.GestionMulDevis_S1(StArg : String);
Begin

   TFMUL(Ecran).SetDbliste('BTMULDEVIS_S1');

   SetControlText('TGPAFFAIRE', 'Code Chantier');

   TTABSheet(GetControl('PCOMPLEMENT')).TabVisible := False;

   if THDbGrid (GetControl('FLISTE')) <> nil then
     	Begin
	  	Fliste := THDbGrid (GetControl('FLISTE'));
    	Fliste.OnDblClick := FlisteDblClickLine;
     	end;

   if GetArgumentValue(StArg, 'ACTION') = 'MODIFICATION'  then
	    begin
		  //GP_NATUREPIECEG.enabled:=false;
  		Consultation := False;
	    end ;

	 if GetArgumentValue(StArg, 'ACTION') = 'CONSULTATION'  then
  		begin
	    //GP_NATUREPIECEG.enabled:=false;
  		Consultation := True;
	    SetControlProperty('BInsert1','visible',False);
		  SetControlProperty('BInsert','visible',False);
	    end ;

end;
*}

procedure TOF_BTCODEAFFAIRE.MnZTiersClick(Sender: TObject);
var CodeTiers : String;
begin
  CodeTiers := Fliste.datasource.dataset.FindField('GP_TIERS').AsString;
  AGLLanceFiche('GC','GCTIERS','',TiersAuxiliaire(CodeTiers,False),'ACTION=CONSULTATION;MONOFICHE');
end;

procedure TOF_BTCODEAFFAIRE.MnZSuiviCdeClick(Sender: TObject);
var ParamDoc : RParamDoc;
    Toto : String;
begin

  ParamDoc := ChargeCleDoc;
  Toto := EncodeCleDoc(ParamDoc.cledoc);

  AGLLanceFiche('BTP','BTSUIVICDE','','', 'DOCUMENT='+ Toto);
  
end;


procedure TOF_BTCODEAFFAIRE.MnZAppelsClick(Sender: TObject);
var Affaire : String;
begin

  Affaire := Fliste.datasource.dataset.FindField('GP_AFFAIRE').AsString;

  AglLanceFiche('BTP','BTMULAPPELS','AFF_AFFAIRE0=W;AFF_CHANTIER=' + Affaire,'','STATUT=APP'); // Appels

end;

Procedure TOF_BTCODEAFFAIRE.ControleChamp(Champ : String;Valeur : String);
Begin

  if champ = 'REAJUSTEANA' then ReajusteAnal := True;
  
  if Champ = 'RECUP' then Recup.Checked := True;

	if Champ='MEMOIRE' then IsMemoire:=true;

  if Champ = 'CF'    then TypePiece := Champ;

  if Champ='MODIFDATSIT' then
  begin
  	ModifDateSit := true;
    SetControlEnabled('AFF_GENERAUTO', False);
  end;
  
  if Champ = 'REAJUSTESIT' then
  begin
    ReajusteSit := true;
  end;

  IF      Champ = 'CHANTIER'  Then Stock := True
  else if Champ = 'STOCK'     Then Stock := False;

  if champ='COTRAITANCE' then
    Setcontroltext('XX_WHERE', GetControltext('XX_WHERE')+' AND AFF_MANDATAIRE IN ("X", "-")')
  else if champ='SOUSTRAITANCE' then
    Setcontroltext('XX_WHERE', GetControltext('XX_WHERE')+' AND SSTRAITE > 0');

  //if champ='STATUT' then SetControlText('GP_STATUTAFFAIRE0', Valeur);
  if Champ='DATEDEB'  then SetControlText('GP_DATEPIECE', Valeur);
  if Champ='DATEFIN'  then SetControlText('GP_DATEPIECE_', Valeur);

  if champ='ETAT' then TheEtatAffaire := Valeur;

  if Champ='STATUT' then
     Begin
      StatutAff := valeur;
     if Valeur = 'APP' then
        Begin
	    	SetControlText('XX_WHERE', GetControltext('XX_WHERE')+' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("","W")');
        if GetControl('AFF_AFFAIRE0') <> nil then SetControlText('AFF_AFFAIRE0', 'W');
        SetControlText('AFFAIRE0', 'W');
        SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel');
        SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Appel');
        SetControlText('TGPAFFAIRE', 'Code Appel');
        end
     else if Valeur = 'INT' then
        Begin
    		SetControlText('XX_WHERE', GetControltext('XX_WHERE')+' AND ((SUBSTRING(GP_AFFAIRE,1,1)="I") OR (GP_GENERAUTO="CON"))');
        SetControlText('AFFAIRE0', 'I');
        if GetControl('AFF_AFFAIRE0') <> nil then SetControlText('AFF_AFFAIRE0', 'I');
        SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Contrat');
        SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Contrat');
        SetControlText('TGPAFFAIRE', 'Code Contrat');
        End
     else if Valeur = 'GRP' then
        Begin
	    	SetControlText('XX_WHERE', GetControltext('XX_WHERE')+' AND SUBSTRING(GP_AFFAIRE,1,1)IN ("W","A")');
        if GetControl('AFF_AFFAIRE0') <> nil then SetControlText('AFF_AFFAIRE0', 'A');
        SetControlText('AFFAIRE0', 'A');
        SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
        SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Affaire');
        SetControlText('TGPAFFAIRE', 'Code Affaire');
        end
     else if Valeur = 'AFF' then
        Begin
      	SetControlText('XX_WHERE', GetControltext('XX_WHERE')+' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("A","")');
        if pos(ecran.name,'BTDEVISSTRAITE;BTDEVISCOTRAI_MUL') = 0 then
          if GetControl('AFF_AFFAIRE0') <> nil then SetControlText('AFF_AFFAIRE0', 'A');
        SetControlText('AFFAIRE0', 'A');
        SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Chantier');
        SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Chantier');
        SetControlText('TGPAFFAIRE', 'Code Chantier');
        end
     else if Valeur = 'PRO' then
        Begin
      	SetControlText('XX_WHERE', GetControltext('XX_WHERE')+' AND SUBSTRING(GP_AFFAIRE,1,1)="P"');
        if GetControl('AFF_AFFAIRE0') <> nil then SetControlText('AFF_AFFAIRE0', 'P');
        SetControlText('AFFAIRE0', 'P');
        SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel d''offre');
        SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Appel d''offre');
        SetControlText('TGPAFFAIRE', 'Code Appel d''offre');
        end
     else
        Begin
        if GetControl('AFF_AFFAIRE0') <> nil then SetControlText('AFF_AFFAIRE0', '');
        SetControlText('AFFAIRE0', '');
        SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
        SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Affaire');
        SetControlText('TGPAFFAIRE', 'Code Affaire');
        end;
     end;
  if Champ = 'SAVPOC' then
  begin
    SaisieAvancPoc := True;
  end;

  if Champ='ACCEPTINVERSE' then
  begin
    SetControltext('ACCEPTINVERSE','X');
    SetControlText('AFF_ETATAFFAIRE', 'ACP');
  end;

  if Champ='AFF_GENERAUTO' then
  begin
    SetControltext('AFF_GENERAUTO',Valeur);
  end;

end;

procedure TOF_BTCODEAFFAIRE.Onload;
var TypePiece : String;
    Caption   : string;
    QQ        : TQuery;
begin
Inherited;

  TypePiece := GetControlText ('GP_NATUREPIECEG');

  //En cas d'appel multi natures de pices, remplacement de : par ;
  TypePiece := StringReplace(TypePiece,':',';',[rfReplaceAll]);
  SetControlText ('GP_NATUREPIECEG',TypePiece);

  if (TypePiece <> ''){ and (Length(TypePiece) = 3) } then
  Begin
    if copy(ecran.name,1,15)='BTPREPACHANTIER' then
    begin
      Caption := 'Préparation de chantier depuis les '+rechdom('GCNATUREPIECEG',GetControlText ('GP_NATUREPIECEG'),false);
    end
    else if copy(ecran.name,1,15)='BTACCDEVNEG_MUL' then
    begin
      Caption := 'Acceptation des '+rechdom('GCNATUREPIECEG',GetControlText ('GP_NATUREPIECEG'),false);
    end
    else if copy(ecran.name,1,15)='BTREJDEVNEG_MUL' then
    begin
      Caption := 'Rejet des '+rechdom('GCNATUREPIECEG',GetControlText ('GP_NATUREPIECEG'),false);
    end
    else if copy(ecran.name,1,12)<>'BTACCDEV_MUL' then
    begin
      if copy(ecran.name,1,13)='BTRECTSIT_MUL' then
      begin
        if THCheckBox(getControl('AVOIRUNIQUEMENT')).checked then
          Caption := 'Génération d''avoir global depuis situation'
        else
          Caption := 'Modification Situation (via Avoir)';
      end
      else
      begin
        Caption := 'Liste '+rechdom('GCNATUREPIECEG',GetControlText ('GP_NATUREPIECEG'),false);
        if IsMemoire then
        begin
          Caption := TraduireMemoire ( 'Traitement de cloture');
        end;
      end;
    end
    else
    begin
      if GetControlText('ACCEPTINVERSE')<>'X' then
        Caption := 'Acceptation '+rechdom('GCNATUREPIECEG',GetControlText ('GP_NATUREPIECEG'),false)
      else
        Caption := 'Annulation de l''acceptation '+rechdom('GCNATUREPIECEG',GetControlText ('GP_NATUREPIECEG'),false);
    end;
    ecran.caption := caption;
    UpdateCaption(Ecran);
    RenommageEnteteColonnes(TypePiece);
  end;

	if (GetInfoParPiece(TypePiece,'GPP_VENTEACHAT'))='ACH'  then
  begin
    Setcontrolvisible('AFF_GENERAUTO',False);
    Setcontrolvisible('TAFF_GENERAUTO',False);
    if BTDuplic <> nil then BTDuplic.Visible := False;
    Setcontrolproperty('GP_TIERS','Plus',' AND T_NATUREAUXI="FOU"');
    //FV1 : 19/12/2017 - FS#2818 - GUINIER - Interdire la saisie directe de Factures/Avoirs/Factures fournisseur
    if (TypePiece = 'FF') OR  (TypePiece = 'AFS') then
    BEGIN
      if SaisieInterdite then
      begin
        BInsert.Visible := false;
        if BTDuplic  <> nil Then BTDuplic.visible  := false;
        if BTDuplic1 <> nil Then BTDuplic1.visible := false;
      end;
    end;
  end
  else if TypePiece = 'DBT' then
  begin
    if THEdit(GetControl ('ZEACTION')) <> nil then
    begin
    if GetControlText ('ZEACTION')='GENCONTRETU' then
    begin
      ecran.caption := 'Génération des contre études';
      UpdateCaption(Ecran);
      Setcontrolvisible('BINSERT',False);
      Setcontrolvisible('BINSERT1',False);
      if BTDuplic <> nil then BTDuplic.Visible := False;
    end;
    end;
  end
  else if (TypePiece = 'ABT') or (TypePiece = 'ABP')  then
  begin
    Setcontrolvisible('AFF_GENERAUTO',False);
    Setcontrolvisible('TAFF_GENERAUTO',False);
    //FV1 : 19/12/2017 - FS#2818 - GUINIER - Interdire la saisie directe de Factures/Avoirs/Factures fournisseur
    if SaisieInterdite then
    begin
      BInsert.Visible := false;
      if BTDuplic  <> nil Then BTDuplic.visible  := false;
      if BTDuplic1 <> nil Then BTDuplic1.visible := false;
    end;
  end
  else if (TypePiece = 'PBT') or (TypePiece = 'CBT') or (TypePiece = 'FPR') or (TypePiece = 'FAC') or (TypePiece = 'AVC') then
  begin
    Setcontrolvisible('AFF_GENERAUTO',False);
    Setcontrolvisible('TAFF_GENERAUTO',False);
    if BTDuplic <> nil then BTDuplic.Visible := False;
    if (typePiece = 'FPR') or (typePiece = 'FAC') or (typePiece = 'AVC') then SetControlText ('AFF_GENERAUTO','');
  end
  else if (TypePiece = 'DAP') or (TypePiece = 'DAC') then
  begin
    if copy(Ecran.name,1,14)<>'BTDEVISINT_MUL' then
    begin
      Setcontrolvisible('BINSERT',False);
    end;
    Setcontrolvisible('BACOMPTES',False);
    if BTDuplic <> nil then BTDuplic.Visible := False;
    Setcontrolvisible('AFF_GENERAUTO',False);
    Setcontrolvisible('TAFF_GENERAUTO',False);
  end
  else if TypePiece = 'ETU' then
  begin
    Setcontrolvisible('BINSERT1',False);
    Setcontrolvisible('BACOMPTES',False);
    Setcontrolproperty('BINSERT','Hint',TraduireMemoire('Nouvelle étude'));
    Setcontrolproperty('BOUVRIR','Hint',TraduireMemoire('Valider les études sélectionnées'));
    if BTDuplic1 <> nil Then BTDuplic1.Hint := TraduireMemoire('Dupliquer l''étude');
    if BTDuplic  <> nil then BTDuplic.Visible := True;
  end
  else if TypePiece = 'LBT' then
  begin
    if BTDuplic <> nil then BTDuplic.Visible := False;
  end
  else if TypePiece = 'BCE' then
  begin
    if copy(ecran.name,1,14)<>'BTSELPIECE_POC' then
    begin
     Setcontrolproperty('BINSERT','Hint',TraduireMemoire('Nouvelle contre-étude'));
     if BTDuplic1 <> nil Then
     Begin
        BTDuplic1.Hint    := TraduireMemoire('Dupliquer la contre étude');
        BTDuplic1.Visible := True;
     end;
     if BTDuplic <> nil then BTDuplic.Visible := False;
    end;
  end
  else if TypePiece = 'BFC' then
  begin
   if BTDuplic1 <> nil Then BTDuplic1.visible := false;
   if BTDuplic  <> nil then BTDuplic.Visible := False;
  end
  else if TypePiece = 'B00' then
  begin
    if BTDuplic1 <> nil Then BTDuplic1.visible := false;
    if BTDuplic  <> nil then BTDuplic.Visible := False;
  end
  else if TypePiece = 'PBT' then
  begin
    if GetControl('BVISUALP')<> nil then
    begin
    	TToolbarButton97(GetControl('BVISUALP')).visible := true;
      TToolbarButton97(GetControl('BVISUALP')).onclick := BTVisualPClick;
    end;
  end
  else  if TypePiece = 'FBT' then
  begin
    Setcontrolproperty('BOUVRIR','Hint',TraduireMemoire('Valider les Factures provisoires'));
    //FV1 : 19/12/2017 - FS#2818 - GUINIER - Interdire la saisie directe de Factures/Avoirs/Factures fournisseur
    if SaisieInterdite then
    begin
      BInsert.Visible := false;
      if BTDuplic  <> nil Then BTDuplic.visible  := false;
      if BTDuplic1 <> nil Then BTDuplic1.visible := false;
    end;
  end;

if UpperCase(copy(Ecran.Name,1,15)) = 'BTPREPACHANTIER' then
begin
  if THEdit(GetControl ('ZEACTION')) <> nil then
  begin
    if THEdit(GetControl ('ZEACTION')).Text = 'PLANNIFICATION' then
    begin
      THEdit(GetControl('AFF_ETATAFFAIRE')).Enabled := false;
{    Mis en commentaire par BRL le 6/08/2013
     Pour rétablir le code tel qu'il était avant la modif de LS le 19/06.
     Le code AFF_ETATAFFAIRE est passé en paramètre lors de l'appel de la fiche :  

      //FV1 - 30/07/2013 - impossible de faire préparation sur contre étude parce que cette dernière n'est pas acceptée !!!
      //FS#617 - TEAM RESEAUX - Préparation de chantier sur contre étude
    	SetControlText('AFF_ETATAFFAIRE', 'ACP');
      if TypePiece = 'BCE' then
      begin
        SetControlText ('AFF_ETATAFFAIRE', 'ENC');
      End;
      //
}
      TToolbarButton97(GetControl('BInsert')).visible := false;
      TToolbarButton97(GetControl('BINSERT1')).visible := false;
      if BTDuplic <> nil then BTDuplic.Visible := False;
    end;
  end;
end else if UpperCase(copy(Ecran.Name,1,11)) = 'BTDEVIS_MUL' then
begin
  //
  if THEdit(GetControl ('ETATAFFAIRE')) <> nil then THMultiValComboBox(GetControl('AFF_ETATAFFAIRE')).Text := TheEtatAffaire;
  //
  if THEdit(GetControl ('ZEACTION')) <> nil then
  begin
    if THEdit(GetControl ('ZEACTION')).Text = 'GENCONTRETU' then
    begin
      //FV1 : 14/12/2017 - FS#2808 - TEAM RESEAUX : autorisation de passer les devis non acceptés en contre-étude
      if Assigned(EtatAffaire) then
      begin
        EtatAffaire.Enabled := GensurDevisNonACPT;
        if GensurDevisNonACPT then
        begin
          EtatAffaire.Plus := 'AND CC_CODE = "ACP" OR CC_CODE = "ENC"';
          EtatAffaire.Text := '';
        end;
      end;
      //THEdit(GetControl('AFF_ETATAFFAIRE')).Enabled := false;
      TToolbarButton97(GetControl('BInsert')).visible := false;
      TToolbarButton97(GetControl('BINSERT1')).visible := false;
      if BTDuplic <> nil then BTDuplic.Visible := False;
    end else if THEdit(GetControl ('ZEACTION')).Text = '' then
    begin
      if THEdit(GetControl('GP_AFFAIRE')).text <> '' then
      begin
        QQ := OpenSql ('SELECT AFF_ETATAFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE="' + THEdit(GetControl('GP_AFFAIRE')).text + '"',true,1,'',True);
        if not QQ.eof then
        begin
          if Pos(QQ.FindField('AFF_ETATAFFAIRE').AsString,'TER;CLO')>0 then
          begin
            TToolbarButton97(GetControl('BInsert')).visible := false;
            TToolbarButton97(GetControl('BINSERT1')).visible := false;
            if BTDuplic <> nil then BTDuplic.Visible := False;
          end;
        end;
        ferme (QQ);
      end;
    end;
  end;
  //
end;

  if UpperCase(copy(Ecran.Name,1,13)) = 'BTREGLAFFAIRE' then
  begin
    if GetControlText ('GP_AFFAIRE') = '' then
    begin
  Setcontrolvisible('BSELECTAFF1',True);
  Setcontrolvisible('BEFFACEAFF1',True);
  Setcontrolenabled('GP_AFFAIRE1',True);
  Setcontrolenabled('GP_AFFAIRE2',True);
  Setcontrolenabled('GP_AFFAIRE3',True);
  Setcontrolproperty('TGPAFFAIRE','Caption',TraduireMemoire('Affaire'));
  end;
  end;

  if UpperCase(copy(Ecran.Name,1,15)) = 'BTPIECESAFF_MUL' then
  begin
    if TCheckBox(GetControl('CONSULTATION')).checked=True then
    begin
  Setcontrolvisible('BSELECTAFF1',False);
  Setcontrolvisible('BEFFACEAFF1',False);
  Setcontrolenabled('GP_AFFAIRE1',False);
  Setcontrolenabled('GP_AFFAIRE2',False);
  Setcontrolenabled('GP_AFFAIRE3',False);
  Setcontrolenabled('GP_TIERS',False);
  Setcontrolenabled('T_LIBELLE',False);
  Setcontrolvisible('Binsert',False);
      if BTDuplic <> nil then BTDuplic.Visible := False;
  end;
  end;

  if assigned(GetControl('AVANC')) then
  if GetControlText('AVANC')='X' then SetControlVisible('BInsert',false);

end;

procedure TOF_BTCODEAFFAIRE.OnUpdate;
var  F : TFMul ;
    i:integer;
begin
inherited ;

  if not (Ecran is TFMul) then exit;

  //If ((Ecran.name='BTDEVIS_MUL') or (Ecran.name='BTACCDEV_MUL'))
  If (pos(ecran.name,'BTDEVIS_MUL;BTACCDEV_MUL') > 0) {$IFNDEF LINE} and (GetControlText ('GP_NATUREPIECEG') = 'ETU') {$ENDIF} then
   Begin
   F:=TFMul(Ecran) ;
{$IFDEF EAGLCLIENT}
   for i:=0 to F.FListe.ColCount-1 do
       begin
// A FAIRE
       end;
{$ELSE}
    for i:=0 to F.FListe.Columns.Count-1 do
       begin
{$IFNDEF LINE}
       if F.FListe.Columns[i].FieldName='GP_NUMERO' then
          F.FListe.Columns[i].Title.caption:=TraduireMemoire('N Etude');
{$ENDIF}
       if F.FListe.Columns[i].FieldName='AFF_AVENANT' then
       begin
          F.Fliste.columns[i].visible:=false ;
          F.SetVisibleColumn(i,false)
        end;
       end;
{$ENDIF}
    end;
	FListeRowEnter(self);
end;

procedure TOF_BTCODEAFFAIRE.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('GP_AFFAIRE'));
Aff0:=THEdit(GetControl('AFFAIRE0'));
Aff1:=THEdit(GetControl('GP_AFFAIRE1'));
Aff2:=THEdit(GetControl('GP_AFFAIRE2'));
Aff3:=THEdit(GetControl('GP_AFFAIRE3'));
Aff4:=THEdit(GetControl('GP_AVENANT'));
Tiers:=THEdit(GetControl('GP_TIERS'));
end;

procedure TOF_BTCODEAFFAIRE.RenommageEnteteColonnes(TypePiece: string);
var i : integer;
{$IFDEF EAGLCLIENT}
		Gr : THgrid;
{$ELSE}
		Gr : THDbgrid;
{$ENDIF}

    stChamp,Libelle : string;
begin
	Gr := TFMul(ecran).fliste;
{$IFDEF EAGLCLIENT}
	For i:=0 to Gr.ColCount -1 do
  Begin
    StChamp := TFMul(Ecran).Q.FormuleQ.GetFormule(Gr.ColNames [i]); //21/10/2005 on rcupre le nom ttoaldu champ
    if UpperCase (stChamp)='GP_NUMERO' then
    begin
    	if TypePiece = 'ETU' then Libelle := 'N Etude'
    	else if TypePiece = 'BCE' then Libelle := 'N C.Etude'
      else if (TypePiece = 'DBT') or (TypePiece = 'DAP') then Libelle := 'N Devis'
      else Libelle := 'Numéro';
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
    end;
  end;
{$ELSE}
	For i:=0 to Gr.Columns.Count-1 do
  Begin
    StChamp := TFMul(Ecran).Q.FormuleQ.GetFormule(Gr.Columns[i].FieldName); //21/10/2005 on rcupre le nom ttoaldu champ
    if UpperCase (stChamp)='GP_NUMERO' then
    begin
    	if TypePiece = 'ETU' then Libelle := 'N Etude'
    	else if TypePiece = 'BCE' then Libelle := 'N C.Etude'
      else if (TypePiece = 'DBT') or (TypePiece = 'DAP') then Libelle := 'N Devis'
      else Libelle := 'Numéro';
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
    end;
  end;
{$ENDIF}

end;

procedure TOF_BTCODEAFFAIRE.BInsertClick(Sender: TObject);
var Document : string;
begin
  if (copy(ecran.name,1,13)='BTSITPROV_MUL') then Exit;
  Document := GetControltext('GP_NATUREPIECEG');
	if GetControlText('GP_AFFAIRE') = '' then
		 AGLCreerPieceBTP(['','',Document,'','',StatutAff, Stock], 0)
	else
		 AGLCreerPieceBTP([GetControlText('GP_TIERS'),GetControlText('GP_AFFAIRE'),Document,'','CRE', GetControlText('AFFAIRE0'),Stock],0);
  if GetParamSocSecur('SO_REFRESHMUL', true) then
  begin
  	//TToolBarButton97(GetControl('Bcherche')).Click;
    Refreshdb;
  end;

end;

procedure TOF_BTCODEAFFAIRE.ModificationDateSituation (TheChaine : string);
var naturepiece: string;
    Souche    : string;
    LaChaine  : string;
    bof       : string;
    Numero    : Integer;
    //Indice : integer;
    TOBDates  : TOB;
    TOBSit    : TOB;
    //DateSit   : TDateTime;
    QQ        : TQuery;
    Req       : String;
begin

	laChaine := TheChaine;
	naturePiece := READTOKENST(LaChaine);
  Bof := READTOKENST(LaChaine); // Date
  Souche :=READTOKENST(LaChaine);
  Numero := StrToInt(READTOKENST(LaChaine));
  //Indice := StrToInt(READTOKENST(LaChaine));
  TOBSIT := TOB.Create ('BSITUATIONS',nil,-1);
  TOBDates := TOB.Create ('LES DATES', nil,-1);
  //
  Req := 'SELECT * FROM BSITUATIONS WHERE BST_NATUREPIECE="'+NaturePiece+'" AND '+
  				'BST_SOUCHE="'+Souche+'" AND BST_NUMEROFAC="'+IntToStr(Numero)+'"';
  //
  QQ := openSql(Req,true,-1,'',true);
  if not QQ.eof then
  begin
  	TOBSit.SelectDB('',QQ);
  end;
  Ferme (QQ);
  if TOBSit.GetInteger('BST_NUMEROFAC')=Numero then
  begin
    TOBDates.AddChampSupValeur('RETOUROK','-');
    TOBDates.AddChampSupValeur('CTRLEX','X');
    TOBDates.AddChampSupValeur('DATFAC',TOBSit.getValue('BST_DATESIT'));
    TOBDates.AddChampSupValeur('DATESITUATION','X');
    TRY
      TheTOB := TOBDates;
      AGLLanceFiche('BTP','BTDEMANDEDATES','','','');
      TheTOB := nil;
    FINALLY
      if TOBDates.getValue('RETOUROK')='X' then
      begin
        TOBSit.SetDateTime ('BST_DATESIT',TOBDates.GetDateTime('DATFAC'));
        TOBSit.UpdateDB(false);
      end;
    END;
  end;
  freeAndNil(TOBDates);
  FreeAndNil(TOBSIT);

end;

procedure BTPgenerePropal(parms: array of variant; nb: integer);
var F: TForm;
  Titre, NaturePiece,Souche: string;
  Numero,Indice : integer;
begin
  F := TForm(Longint(Parms[0]));
  Titre := F.Caption;
  NaturePiece := Trim(string(Parms[1]));
  Souche := string(Parms[2]);
  Numero := integer(Parms[3]);
  Indice := integer(Parms[4]);
  LSEGenerepropal (naturePiece,Souche,Numero,Indice);
end;

procedure BTPRefuseDevis (parms: array of variant; nb: integer);
var
  F: TForm;
  OM: TOF;
begin
  F := TForm (Longint (Parms [0] ) ) ;
  if (F is TFMul) then
    OM := TFMul (F).LaTOF
  else
    OM := nil;
  if (OM is TOF_BTCODEAFFAIRE) then TOF_BTCODEAFFAIRE (OM).RefuseDevis;
end;

procedure BTPGenereConvention (parms: array of variant; nb: integer);
var F: TForm;
  Titre, NaturePiece,Souche: string;
  Numero,Indice : integer;
begin
  F := TForm(Longint(Parms[0]));
  Titre := F.Caption;
  NaturePiece := Trim(string(Parms[1]));
  Souche := string(Parms[2]);
  Numero := integer(Parms[3]);
  Indice := integer(Parms[4]);
  LSEGenereConvention (naturePiece,Souche,Numero,Indice);
end;

procedure ValidefactProvisoire (parms: array of variant; nb: integer);
var F: TForm;
    OM: TOF;
begin
  F := TForm (Longint (Parms [0] ) ) ;
  if (F is TFMul) then
    OM := TFMul (F).LaTOF
  else
    OM := nil;
  if (OM is TOF_BTCODEAFFAIRE) then TOF_BTCODEAFFAIRE (OM).ValidelesfacturesProv;
end;


//TestLionel
procedure TOF_BTCODEAFFAIRE.MnExpCotraitXml (Sender : Tobject);
var ParamDoc : RParamDoc;
		ExportXls : TexportDocXls;
    fournisseur : string;
begin

  if pos(ecran.name,'BTDEVIS_MUL;BTDEVISSTRAITE;BTDEVISCOTRAI_MUL') > 0 then
  begin
    ParamDoc := ChargeCleDoc;
  	fournisseur := AGLLanceFiche('BTP','BTMULAFFECTEXTE','','','COTRAITANCE;DOCUMENT='+EncodeCleDoc(ParamDoc.cledoc)+';SELECTION');
    if fournisseur <> '' then
    begin
      ExportXls := TexportDocXls.create;
    	ExportXls.ModeExport  := TmeXml;
      ExportXls.TypeExport := TteCotrait;
      ExportXls.CleDoc := ParamDoc.cledoc;
      ExportXls.ModeAction := XmaDocument;
      ExportXls.Fournisseur := Fournisseur;
      ExportXls.Affaire := ParamDoc.Affaire;
      ExportXls.Genere;
      ExportXls.free;
    end;
  end;

end;

procedure TOF_BTCODEAFFAIRE.BTAnalyseClick(Sender: Tobject);
var i :integer;
    TOBPieces :TOB;
    Numero, Indice : integer;
    NaturePiece, Souche :string;
    Mess : string;
begin

  If (TFMul(ecran).Fliste.nbSelected = 0)  and not (TFMul(ecran).Fliste.AllSelected) then
  begin
    PGIInfo(TexteMessage [13]);
    Exit;
  end;

  Mess:=TexteMessage[12];
  if (PGIAskAF (Mess, Ecran.Caption)<>mrYes) then exit;

  SourisSablier;

  // on crée une TOB de toutes les études sélectionnées
  TobPieces := TOB.Create ('Liste des études',NIL, -1);

  try
  if TFMul(Ecran).Fliste.AllSelected then
  BEGIN
    TFMul(Ecran).Q.First;
    while Not TFMul(Ecran).Q.EOF do
    BEGIN
      NaturePiece:=TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
      Souche:=TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
      Numero:=TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger;
      Indice:=TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger;

      PrepaAnalyseMultiDoc (NaturePiece, Souche, Numero, Indice,TobPieces);

      TFMul(Ecran).Q.NEXT;
    END;
    TFMul(Ecran).Fliste.AllSelected:=False;
  END
  else
  begin
    for i:=0 to TFMul(Ecran).Fliste.nbSelected-1 do
    begin
      TFMul(Ecran).Fliste.GotoLeBookmark(i);
      NaturePiece:=TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
      Souche:=TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
      Numero:=TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger;
      Indice:=TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger;

      PrepaAnalyseMultiDoc (NaturePiece, Souche, Numero, Indice,TobPieces);
    end;
  end;

  If TobPieces.detail.count > 0 then AnalyseMultiDoc (TobPieces);

  finally
  TobPieces.Free;
  SourisNormale ;
  end;

  RefreshDB('FListe');

end;

procedure TOF_BTCODEAFFAIRE.BTDuplicationClick(Sender: TObject);
begin
  AppelDuplication('');
  refreshDB('FLISTE');
end;

procedure TOF_BTCODEAFFAIRE.MnInsertABT(Sender: TObject);
begin
  if (Not jaiLeDroitTag(145431)) then Exit;
  AppelDuplication('ABT');
  refreshDB('FLISTE');
end;

procedure TOF_BTCODEAFFAIRE.MnInsertFBT(Sender: TObject);
begin
  if (Not jaiLeDroitTag(145421)) then Exit;
  AppelDuplication('FBT');
  refreshDB('FLISTE');
end;


procedure TOF_BTCODEAFFAIRE.MnInsertABP(Sender: TObject);
var Document : string;
begin
  Document := 'ABP';
	if GetControlText('GP_AFFAIRE') = '' then
		 AGLCreerPieceBTP(['','',Document,'','',GetControlText('AFFAIRE0')], 0)
	else
		 AGLCreerPieceBTP([GetControlText('GP_TIERS'),GetControlText('GP_AFFAIRE'),Document,'','CRE', GetControlText('AFFAIRE0')],0);
  if GetParamSocSecur('SO_REFRESHMUL', true) then
  begin
  	//TToolBarButton97(GetControl('Bcherche')).Click;
    Refreshdb;
  end;
end;

procedure TOF_BTCODEAFFAIRE.MnInsertFBP(Sender: TObject);
var Document : string;
    TheTypeAffaire : string;
begin
  Document := 'FBP';
  if GetCOntrol('AFFAIRE0')<> nil then TheTypeAffaire := GetControlText('AFFAIRE0')
  else if GetCOntrol('AFF_AFFAIRE0')<> nil then TheTypeAffaire := GetControlText('AFF_AFFAIRE0');
  if theTypeAffaire = '' then TheTypeAffaire := 'A';
	if GetControlText('GP_AFFAIRE') = '' then
  begin
		 AGLCreerPieceBTP(['','',Document,'','',TheTypeAffaire,true], 0)
	end else
		 AGLCreerPieceBTP([GetControlText('GP_TIERS'),GetControlText('GP_AFFAIRE'),Document,'','CRE', TheTypeAffaire,true],0);
  if GetParamSocSecur('SO_REFRESHMUL', true) then
  begin
  	//TToolBarButton97(GetControl('Bcherche')).Click;
    Refreshdb;
  end;
end;


procedure TOF_BTCODEAFFAIRE.MnDuplicABP(Sender: TObject);
begin
  AppelDuplication('ABP');
  refreshDB('FLISTE');
end;

procedure TOF_BTCODEAFFAIRE.MnDuplicFBP(Sender: TObject);
begin
  AppelDuplication('FBP');
  refreshDB('FLISTE');
end;



procedure TOF_BTCODEAFFAIRE.AppelDuplication(TypeFacture: String);
var Avenant     : String;
    EtatAffaire : String;
    TheChaine   : String;
    //
    ClePiece		: array [0..7] of Variant;
    //
    Fliste	    : THDbGrid;
begin

  //ClePiece[7] : Saisie Cde Stock ou chantier (True or False)

  if (pos(ecran.name,'BTDEVIS_MUL;BTDEVISSTRAITE;BTDEVISCOTRAI_MUL;BTPIECE_MUL;BTPIECENEG_MUL;BTDEVISNEG_MUL;BTSITPROV_MUL') = 0) then exit;

  Fliste := THdbgrid(GetControl('FLISTE'));

  ClePiece[0] := Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
  ClePiece[1] := DateToSTr(Fliste.datasource.dataset.FindField('GP_DATEPIECE').AsDateTime);
  ClePiece[2] := Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
  ClePiece[3] := IntToStr(Fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger);
  ClePiece[4] := IntToStr(Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger);

  if TypeFacture = '' then
    ClePiece[5] := Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString
  else
    ClePiece[5] := TypeFacture;

  ClePiece[6] := Fliste.datasource.dataset.FindField('AFF_GENERAUTO').AsString;

	TRY
  	EtatAffaire := Fliste.datasource.dataset.findfield ('ETATAFFAIRE').AsString;
  except
    PGIInfo('Merci d''ajouter le champs ETATAFFAIRE dans la présentation');
    exit;
  end;

  TheChaine := ClePiece[0]+';'+ClePiece[1]+';'+ClePiece[2]+';'+ClePiece[3]+';'+ClePiece[4]+';';

  if pos(ecran.name,'BTDEVIS_MUL;BTDEVISSTRAITE;BTDEVISCOTRAI_MUL') > 0 then
  begin
    Avenant:=Fliste.datasource.dataset.FindField('AFF_AVENANT').AsString;
    if (Avenant = '00') or (ClePiece[0]= 'ETU') then
     	AppelDupliquePiece ([thechaine,ClePiece[5],'',Stock],3)
    else
      PGIError('Duplication d''un avenant impossible. Veuillez utiliser la création d''un nouvel avenant.',Ecran.Caption);
  end
  else if pos(ecran.name,'BTPIECE_MUL;BTPIECENEG_MUL;BTDEVISNEG_MUL;BTSITPROV_MUL') > 0 then
  begin
    if EtatAffaire= 'TER' then
    begin
      PGIError('Duplication d''un chantier terminé impossible. Veuillez utiliser un autre chantier.',Ecran.Caption);
      exit;
    end;
   	AppelDupliquePiece ([thechaine,ClePiece[5],ClePiece[6],Stock],4);
  end;

end;

procedure TOF_BTCODEAFFAIRE.BTVisualPClick(Sender: Tobject);
Var Cledoc        : R_Cledoc;
begin
  fillchar(Cledoc,SizeOf(Cledoc),#0);
  // transfert de prévision vers excel pour récupération dans visual projet
  //chargement de la TOBPIECE avec la prévision de chantier
  Cledoc.NaturePiece:=TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
  Cledoc.Souche:=TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
  cledoc.NumeroPiece :=TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger;
  Cledoc.Indice:=TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger;
//  TraitementProjetDocument (Cledoc);

end;

procedure TOF_BTCODEAFFAIRE.TRVISUALPClick (Sender : TObject);
begin
	if TCheckBox (GetControl('TRVISUALP')).State = cbUnchecked then
  begin
    THedit(GetControl('XX_WHERE2')).Text := ' AND GP_IDENTIFIANTWOT<=0';
  end else if TCheckBox (GetControl('TRVISUALP')).State = cbchecked then
  begin
    THedit(GetControl('XX_WHERE2')).Text := ' AND GP_IDENTIFIANTWOT>0';
  end else
  begin
    THedit(GetControl('XX_WHERE2')).Text := '';
  end;
end;

procedure TOF_BTCODEAFFAIRE.RefuseMultiDevis(TOBPieces : TOB);

	procedure  RefuseOneDevis (TOBP : TOB);
  var Sql,Affairedevis : string;
  		Cledoc : R_CLEDOC;
      QQ : TQuery;
  begin
		Cledoc := TOB2CleDoc(TOBP);
    Affairedevis := '';
		Sql := 'SELECT GP_AFFAIREDEVIS FROM PIECE WHERE '+WherePiece(Cledoc,ttdPiece ,false);
    QQ := OpenSQL(Sql,True,1,'',true);
    if not QQ.Eof then
    begin
      Affairedevis := QQ.Fields [0].AsString;
    end;
    ferme (QQ);
    
    if Affairedevis <> '' then
    begin
      Sql := 'UPDATE AFFAIRE SET AFF_ETATAFFAIRE="REF",AFF_DATEFIN="'+USDateTime(TOBPieces.GetDateTime('DATEREFUS'))+'",'+
             'AFF_DATERESIL="'+USDateTime(TOBPieces.GetDateTime('DATEREFUS'))+'",AFF_RESILAFF="'+TOBPieces.GetString('MOTIFREFUS')+'" '+
             'WHERE AFF_AFFAIRE="'+AffaireDevis+'"';
      ExecuteSQL(Sql);
    end;
    Sql := 'UPDATE PIECE SET GP_VIVANTE="-" WHERE '+WherePiece(Cledoc,ttdPiece ,false);
    ExecuteSQL(Sql);
  end;

var TOBP : TOB;
		Indice : integer;
begin
	For Indice := 0 to TOBPieces.Detail.Count - 1 do
  begin
    TOBP := TOBPieces.detail[Indice];
    RefuseOneDevis (TOBP);
  end;
end;


procedure TOF_BTCODEAFFAIRE.RefuseDevis;
var OneTOB,TOBPIECE : TOB;
		i : Integer;
begin
	OneTOB := TOB.Create ('UNE TOB',nil,-1);
  TRY
    OneTOB.AddChampSupValeur('RESULT','-');
    OneTOB.AddChampSupValeur('DATEREFUS',Date);
    OneTOB.AddChampSupValeur('MOTIFREFUS','');
    TheTOB := OneTOB;
    AGLLanceFiche('BTP','BTREFUS','','','');
    TheTOB := nil;
    If OneTOB.GetString('RESULT')='X' then
    begin
      if (PGIAskAF (TraduireMemoire('Etes-vous sur d''indiquer ce(s) devis comme refusé(s) ?'), Ecran.Caption)<>mrYes) then exit;
      SourisSablier;
      if TFMul(Ecran).Fliste.AllSelected then
      BEGIN
        TFMul(Ecran).Q.First;
        while Not TFMul(Ecran).Q.EOF do
        BEGIN
          TOBPIECE := TOB.create ('PIECE',OneTOB,-1);
          TOBPIECE.SetString('GP_NATUREPIECEG',TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString);
          TOBPIECE.SetString('GP_SOUCHE',TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString);
          TOBPIECE.SetInteger('GP_NUMERO',TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger);
          TOBPIECE.SetInteger('GP_INDICEG',TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger);
          TFMul(Ecran).Q.NEXT;
        END;
        TFMul(Ecran).Fliste.AllSelected:=False;
      END
      else
      begin
        for i:=0 to TFMul(Ecran).Fliste.nbSelected-1 do
        begin
          TFMul(Ecran).Fliste.GotoLeBookmark(i);
          TOBPIECE := TOB.create ('PIECE',OneTOB,-1);
          TOBPIECE.SetString('GP_NATUREPIECEG',TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString);
          TOBPIECE.SetString('GP_SOUCHE',TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString);
          TOBPIECE.SetInteger('GP_NUMERO',TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger);
          TOBPIECE.SetInteger('GP_INDICEG',TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger);
        end;
      end;
      If OneTOB.detail.count > 0 then RefuseMultiDevis(OneTOB);
    end;
  FINALLY
  	OneTOB.Free;
  	SourisNormale ;
  	RefreshDB('FListe');
  END;
end;

procedure GetLookuprepresentant (parms: array of variant; nb: integer);
var
  F: TForm;
  OM: TOF;
begin
  F := TForm (Longint (Parms [0] ) ) ;
  if (F is TFMul) then
    OM := TFMul (F).LaTOF
  else
    OM := nil;
  if (OM is TOF_BTCODEAFFAIRE) then TOF_BTCODEAFFAIRE (OM).RechRepresentant;
end;

procedure TOF_BTCODEAFFAIRE.RechRepresentant;
var Titre : string;
    TOBPiece : TOB;
begin
  TOBPiece := TOB.Create ('PIECE',nil,-1);
  TOBPiece.SetString('GP_NATUREPIECEG',GetControlText('GP_NATUREPIECEG'));
  Titre := 'Sélection du commercial';
	GetRepresentantEntete(THEdit(GetControl('GP_REPRESENTANT')),Titre,'',TOBPiece);
  TOBPIece.free;
end;


procedure TOF_BTCODEAFFAIRE.ValidelesfacturesProv;

  function OkValideFac (Cledoc : r_cledoc) : boolean;
  var ssAffaire,MessageErr : string;
      NumSit : integer;
      Sql : string;
      QQ : TQuery;
  begin
    MessageErr := 'ERREUR : Une des précédentes situations n''a pas été validée';
    Result := true;
    ssAffaire := '';
    NumSit     := 0;
    Sql := 'SELECT GP_VIVANTE FROM PIECE WHERE '+WherePiece(Cledoc,ttdPiece,True);
    QQ := OpenSQL(Sql,True,1,'',true);
    if not QQ.eof then
    begin
      Result := (QQ.fields[0].AsString='X');
      if not result then MessageErr := 'ERREUR : On ne peut pas valider un document déjà validé';
    end;
    ferme (QQ);
    if Result then
    begin
      // pas de controle sur avoir
      if Cledoc.NaturePiece = 'ABP' then Exit;
      // ---
    Sql := 'SELECT BST_NUMEROSIT,BST_SSAFFAIRE FROM BSITUATIONS WHERE '+WherePiece(Cledoc,ttdSit,True);
    QQ := OpenSQL(Sql,True,1,'',true);
    if not QQ.eof then
    begin
      NumSit := QQ.fields[0].AsInteger;
      ssAffaire := QQ.fields[1].AsString;
    end;
    ferme (QQ);
    if NumSit = 1 then Exit; // si c'est la première situation ...il n'y a en pas avant....etonnant non ?
    if not (ssAffaire='') then
    begin
      Result := false;
      Sql := 'SELECT BST_NATUREPIECE FROM BSITUATIONS WHERE '+
             'BST_SSAFFAIRE="'+ssAffaire+'" AND '+
             'BST_VIVANTE="X" AND '+
             'BST_NUMEROSIT='+IntToStr(NumSit-1);
      QQ := OpenSQL(Sql,True,1,'',true);
      if not QQ.Eof then
      begin
        if QQ.fields[0].AsString = 'FBT' then Result := True;
      end;
      Ferme(QQ);
    end;
    end;
    if not result then PGIError(MessageErr);
  end;

var cledoc : r_cledoc;
    Q : Tquery;
    i : Integer;
    DateSit : TDateTime;
    XX : T_genereFacture;
begin
  //
  if (PGIAskAF ('Vous allez valider définitivement les factures/avoirs sélectionnées.#13#10Confirmez-vous ?', Ecran.Caption)<>mrYes) then exit;
  if not DemandeDatesFacturation (DateSit) then Exit;
  XX := T_genereFacture.create;
  TRY
    if TFMul(Ecran).Fliste.AllSelected then
    BEGIN
      Q:=TFMul(Ecran).Q;
      Q.First;
      while Not Q.EOF do
      BEGIN
        fillchar(Cledoc,SizeOf(Cledoc),#0);
        Cledoc.NaturePiece:=Q.FindField('GP_NATUREPIECEG').AsString;
        Cledoc.Souche:=Q.FindField('GP_SOUCHE').AsString;
        cledoc.NumeroPiece :=Q.FindField('GP_NUMERO').AsInteger;
        Cledoc.Indice:=Q.FindField('GP_INDICEG').AsInteger;
        if not OkValideFac (Cledoc) then continue;
        XX.GenereFacDefinitif(Cledoc,DateSit);
        RapportGen.SauveLigMemo(XX.LibRetour);
        Q.NEXT;
      END;
      TFMul(Ecran).Fliste.AllSelected:=False;
    END ELSE
    BEGIN
      for i:=0 to TFMul(Ecran).Fliste.nbSelected-1 do
      begin
        TFMul(Ecran).Fliste.GotoLeBookmark(i);
        Cledoc.NaturePiece:=TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
        Cledoc.Souche:=TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
        cledoc.NumeroPiece :=TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger;
        Cledoc.Indice:=TFMul(Ecran).Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger;
        if not OkValideFac (Cledoc) then continue;
        XX.GenereFacDefinitif(Cledoc,DateSit);
        RapportGen.SauveLigMemo(XX.LibRetour); 
      END;
    END;
  FINALLY
    if RapportGen.Memo.Lines.count > 0 then
    begin
      RapportGen.AfficheRapport;    //affichage du rapport d'intégration
      RapportGen.Affiche := True;
    end;
    XX.Free;
  END;
end;


procedure TOF_BTCODEAFFAIRE.OnClose;
begin
  RapportGen.free;
  inherited;
end;

procedure TOF_BTCODEAFFAIRE.MnOuvPlatClick(Sender: Tobject);
var cledoc : r_cledoc;
    XX : TGenOuvPlat;
begin
  //
  Fliste := THdbgrid(GetControl('FLISTE'));
  XX := TGenOuvPlat.create;

  cledoc.NaturePiece := Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
  cledoc.DatePiece := Fliste.datasource.dataset.FindField('GP_DATEPIECE').AsDateTime;
  Cledoc.Souche :=  Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
  cledoc.NumeroPiece  := Fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger;
  cledoc.Indice  := Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger;
  XX.TraitePiece(cledoc);
  PGIInfo(XX.LibRetour);
  XX.Free; 

end;

Initialization
registerclasses([TOF_BTCODEAFFAIRE]);
RegisterAglProc( 'BTPGenerepropal', TRUE , 5, BTPgenerePropal);
RegisterAglProc( 'BTPGenereConvention', TRUE , 5, BTPGenereConvention);
registerAglProc( 'RefuseDevis',true,0,BTPRefuseDevis);
registerAglProc('AglGetLookUprepresentant',true,0,GetLookuprepresentant);
registerAglProc('ValidefactProvisoire',true,0,ValidefactProvisoire);
end.

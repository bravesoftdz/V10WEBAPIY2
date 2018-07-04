{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 27/10/2017
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BSAISIEAVANCPOC ()
Mots clefs ... : TOF;BSAISIEAVANCPOC
*****************************************************************}
Unit BSAISIEAVANCPOC_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     AglInit,
     M3FP,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,Fe_Main,
{$else}
     eMul,
{$ENDIF}
     uTob,
     Grids,
     graphics,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     utofAfBaseCodeAffaire,
     UTOF,
     Vierge,
     types,
     Messages,
     UFactgestionAff,
     galPatience,
     Windows
     ;
const
   stLigPP  = 'IND;TYPELIGNE;BM8_PHASETRA;BM8_FAMILLENIV2;BM8_LIBELLE;BM8_QTEMARCHE;BM8_UNITE;BM8_POURCENTAVANC;BM8_QTECUMUL;SUMQTECONSO;BM8_RADQTE;QTEPROJETE'; // production propre
   stLigMOI = 'IND;TYPELIGNE;BM8_PHASETRA;BM8_FAMILLENIV2;BM8_LIBELLE;BM8_QTEMARCHE;BM8_UNITE;BM8_RADQTE;SUMQTECONSO;BM8_POURCENTAVANC;'; // Main d'oeuvre indirecte
   stLigFI  = 'IND;TYPELIGNE;BM8_PHASETRA;BM8_FAMILLENIV2;BM8_LIBELLE;BM8_QTEMARCHE;BM8_UNITE;BM8_RADQTE;BM8_RADPU;BM8_RADMT;MTAVANCCALC;SUMQTECONSO;DEPENSECUMUL;ECARTDEPENSE;'; // Frais indirect
   stLigST  = 'IND;TYPELIGNE;BM8_PHASETRA;BM8_CODEMARCHE;BM8_LIBELLE;BM8_QTEMARCHE;BM8_UNITE;PU;DEPENSECUMUL;DEPENSEESTIME;POURCENTCALC;BM8_POURCENTAVANC;MTAVANCCORRIGE;ECARTDEPENSE;'; // Sous-traitance
   stLigHP  = 'IND;DATEMVT;CODEARTICLE;LIBELLE;QUANTITE;PUACH;MONTANTACH;'; // hors phase
   STAVANCSQL = 'SELECT BAVANCTECHDET.*, '+
                '0 AS PU,'+
                '0 AS QTEAVANC,'+
                '0 AS QTEPROJETE,'+
                '0 AS MTAVANCCALC,'+
                '0 AS MTCUMULCALC,'+
                '0 AS SUMQTECONSO,'+
                '0 AS ECARTDEPENSE,'+
                '0 AS DEPENSEESTIME,'+
                '0 AS POURCENTCALC,'+
                '0 AS DEPENSECUMUL,'+
                '0 AS MTAVANCCORRIGE,'+
                '"LIG" AS TYPELIGNE,'+
                '"" AS DATEMVT,'+
                '"" AS CODEARTICLE,'+
                '"" AS LIBELLE,'+
                '0 AS QUANTITE,'+
                '0 AS PUACH,'+
                '0 AS MONTANTACH,'+
                '1 AS NIVEAU,'+
                '"-" AS NEWONE '+
                'FROM BAVANCTECHDET';
Type

  TTSaisie = (TmsPP,TmsMOI,TmsFI,TmsST);

  TGSaisie = class (TObject)
    fName : string;
    fGS : THGrid;
    fTOB : TOB;
    fStCellCur : string;
    fListSaisie : string;
    TheMode : TTSaisie;
    fLastRow,fLastCol : integer;
    SG_TYPELIG : Integer;
    SG_LIBELLE : Integer;
    SG_PHASETRA : Integer;
    SG_FAMILLENIV2 : Integer;
    SG_QTEMARCHE : Integer;
    SG_UNITE : Integer;
    SG_POURCENTAVANC : Integer;
    SG_QTEMOIS : Integer;
    SG_QTECUMUL : Integer;
    SG_RADQTE : Integer;
    SG_RADPU : Integer;
    SG_RADMT : Integer;
    SG_QTEPROJETE : Integer;
    SG_MTAVANCCALC : Integer;
    SG_SUMQTECONSO : Integer;
    SG_ECARTDEPENSE : integer;
    SG_DEPENSECUMUL : Integer;
    SG_DEPENSEESTIME : Integer;
    SG_POURCENTCALC : Integer;
    SG_MTAVANCCORRIGE : Integer;
    SG_CODEMARCHE : Integer;
    SG_PU : Integer;
    SG_QTECONSOMME : Integer;
  public
    constructor create;
    property GS : Thgrid read fGS write fGS;
    property StCellCur : String read fStCellCur write fStCellCur;
  end;

  TListGs = class(TList)
  private
    function Add(AObject: TGSaisie): Integer;
    function GetItems(Indice: integer): TGSaisie;
    procedure SetItems(Indice: integer; const Value: TGSaisie);
    procedure AddGrille(TG : THgrid; NomGrille, ListeSaisie: String; ZeTOB : TOB);
  public
    constructor create;
    destructor destroy; override;
    property Items [Indice : integer] : TGSaisie read GetItems write SetItems;
    function findGrille (NomGrille : string ): TGSaisie;
    procedure clear; override;
  end;

  TOF_BSAISIEAVANCPOC = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4,Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;
  private
    XX : TFPatience;
    TTS,TTD : TOB;
    ActionResult : TActionFiche;
    TOBSTRUCTLIG : TOB;
    TheListGrid : TListGs;
    ZeCurSaisie : TGSaisie;
    TOBPP,TOBMOI,TOBFI,TOBST,TOBHP : TOB;
    GSPP,GSMOI,GSFI,GSST,GSHP : THGrid;
    FFP,FFQ : string;
    ImList : TImageList;
    FamFraisIndirect : string;
    fGestionAff : TAffichageDoc;
    DebutMois,FinMois : TDateTime;
    CodeAffaire,Aff0,Aff1,Aff2,Aff3,Avenant : string;
    DataModif : boolean;
    // ---
    function ChargeTOBS (CodeAffaire : string) : boolean;
    procedure ConstitueLesTOBS(TTS, TTD, TTL, TTOP: TOB);
    //
    procedure ChargeGrids;
    procedure PrepareGrids;
    procedure PrepareGrid(Nomgrille : String);
    procedure TraiteChamps(TGS: TGSaisie;NomChamps: string;Colonne : integer);
    procedure MemoriseLigneSaisie;
    procedure SMemoriseLigneSaisie(QQ: TQuery);
    procedure AddlesChampsSupSaisie(TOBPT: TOB);
    function CreTOBLigne: TOB;
    procedure afficheLaGrid(Nomgrille: string); overload;
    procedure afficheLaGrid(TGS : TGSaisie); overload;
    procedure afficheLigne(zeGS: ThGrid; ZeTOB: TOB; ZeListe: string;II: integer);
    procedure SetEventGrid(statut : Boolean);
    procedure EnterGrid (NomGrid : string);
    //
    function VisiblePP(TheCur: TGSaisie; TOBL: TOB;Acol: integer): Boolean;
    function ZoneAccessiblePP(ACol, ARow: Integer): boolean;
    procedure ZoneSuivanteOuOkPP(var ACol, ARow: Integer;var Cancel: boolean);
    procedure PostDrawcellPP(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GetCellCanvasPP(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GSCellEnterPP(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExitPP(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowEnterPP(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExitPP(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GridKeyDownPP(Sender: TObject; var Key: Word;Shift: TShiftState);
    //
    function VisibleMOI(TheCur: TGSaisie; TOBL: TOB;Acol: integer): Boolean;
    function ZoneAccessibleMOI(ACol, ARow: Integer): boolean;
    procedure ZoneSuivanteOuOkMOI(var ACol, ARow: Integer;var Cancel: boolean);
    procedure PostDrawcellMOI(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GetCellCanvasMOI(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GSCellEnterMOI(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExitMOI(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowEnterMOI(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExitMOI(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GridKeyDownMOI(Sender: TObject; var Key: Word;Shift: TShiftState);
    //
    function VisibleFI(TheCur: TGSaisie; TOBL: TOB;Acol: integer): Boolean;
    function ZoneAccessibleFI(ACol, ARow: Integer): boolean;
    procedure ZoneSuivanteOuOkFI(var ACol, ARow: Integer;var Cancel: boolean);
    procedure PostDrawcellFI(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GetCellCanvasFI(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GSCellEnterFI(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExitFI(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowEnterFI(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExitFI(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GridKeyDownFI(Sender: TObject; var Key: Word;Shift: TShiftState);
    //
    function VisibleST(TheCur: TGSaisie; TOBL: TOB;Acol: integer): Boolean;
    function ZoneAccessibleST(ACol, ARow: Integer): boolean;
    procedure ZoneSuivanteOuOkST(var ACol, ARow: Integer;var Cancel: boolean);
    procedure GSCellEnterST(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExitST(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowEnterST(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExitST(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GetCellCanvasST(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure PostDrawcellST(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GSSTElipsisClick (Sender : TObject);
    procedure GridKeyDownST(Sender: TObject; var Key: Word;Shift: TShiftState);
    //
    procedure CalculFinal;
    function RecupTypeGraph(TOBExamin: Tob): Integer;
    procedure ChangePage (Sender : Tobject);
    procedure BeforeChangePage (Sender : TObject; Var AllowChanging: Boolean);
    procedure RecupereConsos;
    procedure CalculeMtAvt(TOBL : TOB);
    procedure CalculeQteAvt(TOBL: TOB);
    procedure CalculeProjete(TOBL: TOB);
    procedure CalculePourcentAvt (TOBL:TOB);
    procedure CalculeMtRAD(TOBL: TOB);
    //
    procedure CalcPourcentAvcPP(TOBL: TOB);
    procedure CalcQteCumul(TOBL: TOB);
    procedure CalculePURAD (TOBL : TOB);
    procedure ReajusteRegion(Arect: Trect);
    function FindLigne(TheType, CurrentPhase, CodeAnal, CodeMarche: string): TOB;
    function FindEmplacementLigne(TheType, CurrentPhase: string): TOB;
    procedure CalculSousTraitance;
    procedure CalculeAvancCorrigeST(TOBL: TOB);
    procedure GetDepensesEstimes(TL: TOB);
    function FindLigneConso(TheType, CurrentPhase, CodeAnal,CodeMarche: string): TOB;
    procedure BeforeUpdate;
    procedure UpdateDatas;
    procedure SetInConsultation;
    function GetEmplacementLigne(TheType, CurrentPhase, FamilleNiv2: string): TOB;
  end ;

Implementation
uses CalcOLEGenericBTP,UtilchampsSup,FactUtil,FactOuvrage,ImgList,Paramsoc,AglInitBtp
;


procedure AGLAvancementPOC (Parms : array of variant ; nb : integer) ;
var TheParam  :string;
begin
  TheParam := 'BNATURE='+Parms[1]+';BSOUCHE='+Parms[2]+';BNUMERO='+Parms[3]+';BINDICEG='+Parms[4]+';BMOIS='+Parms[5]+';BAFFAIRE='+Parms[6];
  AGLLanceFiche('BTP','BSAISIEAVANCPOC','','','ACTION=MODIFICATION;'+TheParam);
  Sleep(1);
end;

procedure TOF_BSAISIEAVANCPOC.MemoriseLigneSaisie;
var QQ : TQuery;
begin
  QQ := OpenSQL(STAVANCSQL+' WHERE 1=2', True,-1, '', True);
  SMemoriseLigneSaisie(QQ);
  ferme (QQ);
end;

procedure TOF_BSAISIEAVANCPOC.SMemoriseLigneSaisie (QQ : TQuery);
var TOBEtalon : TOB;
		Indice : integer;
    NomChamps : string;
begin
  //
  TOBSTRUCTLIG.clearDetail;
	TOBEtalon := TOB.Create ('LA LIGNE ETALON',TOBSTRUCTLIG,-1);
  //
  for Indice := 0 to QQ.FieldCount -1 do
  begin
  	NomChamps := QQ.Fields[Indice].FieldName;
		if ExtractPrefixe (NomChamps) = 'BM9' then continue; // pour ne prendre que les champs sup
  	TOBEtalon.AddChampSupValeur (NomChamps,getValInitChampsSup(NomChamps));
  end;
end;

procedure TOF_BSAISIEAVANCPOC.AddlesChampsSupSaisie (TOBPT : TOB);
var Indice : integer;
		TOBEtl : TOB;
    NomChamps : string;
begin
	if TOBSTRUCTLIG.detail.count = 0 then exit;
  TOBEtl := TOBSTRUCTLIG.detail[0];
  //
  for Indice := 0 to TOBETL.NombreChampSup -1 do
  begin
  	NomChamps := TOBEtl.GetNomChamp(1000+Indice);
    if not TOBPT.FieldExists(NomChamps) then
    begin
  		TOBPT.AddChampSupValeur (NomChamps,TOBEtl.Getvaleur(1000+Indice));
    end;
  end;
end;


procedure TOF_BSAISIEAVANCPOC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BSAISIEAVANCPOC.OnDelete ;

  procedure AjouteEventSuppr;
  var QQ: TQuery;
    MotifPiece: TStrings;
    NumEvent: integer;
    LeLibelle: string;
  begin
    LeLibelle := Format('Suppression de l''avancement du mois %s pour le chantier %s',[GetControlText('BMOIS'),BTPCodeAffaireAffiche(CodeAffaire)]);
    MotifPiece := TStringList.Create;
    MotifPiece.Add(LeLibelle);
    NumEvent := 0;
    QQ := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', True,-1, '', True);
    if not QQ.EOF then NumEvent := QQ.Fields[0].AsInteger;
    Ferme(QQ);
    Inc(NumEvent);
    QQ := OpenSQL('SELECT * FROM JNALEVENT WHERE GEV_TYPEEVENT="PO1" AND GEV_NUMEVENT=-1', False);
    QQ.Insert;
    InitNew(QQ);
    QQ.FindField('GEV_NUMEVENT').AsInteger := NumEvent;
    QQ.FindField('GEV_TYPEEVENT').AsString := 'PO1';
    QQ.FindField('GEV_LIBELLE').AsString := LeLibelle;
    QQ.FindField('GEV_DATEEVENT').AsDateTime := Date;
    QQ.FindField('GEV_UTILISATEUR').AsString := V_PGI.User;
    QQ.FindField('GEV_ETATEVENT').AsString := 'OK';
    TMemoField(QQ.FindField('GEV_BLOCNOTE')).Assign(MotifPiece);
    QQ.Post;
    Ferme(QQ);
    MotifPiece.Free;
  end;

begin
  Inherited ;
  if PgiAsk('Etes-vous sur(e) de vouloir supprimer cet avancement ?') <> mryes then exit;
  BEGINTRANS;
  TRY
    BeforeUpdate; // permet de supprimer d'un coup
    AjouteEventSuppr;
    COMMITTRANS;
  EXCEPT
    PGIError('Une erreur s''est produite durant la suppression de cet avancement');
    ROLLBACK;
  END;
  SendMessage (Ecran.Handle,WM_CLOSE,0,0); 
end ;

procedure TOF_BSAISIEAVANCPOC.OnUpdate ;
begin
  Inherited ;
  BEGINTRANS;
  TRY
    BeforeUpdate;
    UpdateDatas;
    COMMITTRANS;
    DataModif := false;
  except
    PGIError('Une erreur s''est produite durant l''enregistrement de cet avancement');
    DataModif := false;
    ROLLBACK;
  end;
end ;

procedure TOF_BSAISIEAVANCPOC.OnLoad ;
begin
  Inherited ;
  XX := FenetrePatience('Récupération de l''avancement',aoMilieu, False,true);
  XX.lAide.Caption := 'Récupération du prévisionnel  ...';
  XX.StartK2000 ;
  XX.Refresh;
  TRY
    //
    SetControlText('BAFFAIRE0',aff0);
    SetControlText('BAFFAIRE1',aff1);
    SetControlText('BAFFAIRE2',aff2);
    SetControlText('BAFFAIRE3',aff3);
    SetControlText('BAVENANT',Avenant);
    //
    TOBPP := TOB.Create('SAISIE PP',nil,-1);
    TOBMOI := TOB.Create('SAISIE MOI',nil,-1);
    TOBFI := TOB.Create('SAISIE FI',nil,-1);
    TOBST := TOB.Create('SAISIE ST',nil,-1);
    TOBHP := TOB.Create ('HORS PHASES',nil,-1);
    //
    TheListGrid := TListGs.create;
    TheListGrid.AddGrille(THGrid(GetControl('GSPP')),'GSPP',stLigPP,TOBPP);
    TheListGrid.AddGrille(THGrid(GetControl('GSMOI')),'GSMOI',stLigMOI,TOBMOI);
    TheListGrid.AddGrille(THGrid(GetControl('GSFI')),'GSFI',stLigFI,TOBFI);
    TheListGrid.AddGrille(THGrid(GetControl('GSST')),'GSST',stLigST,TOBST);
    TheListGrid.AddGrille(THGrid(GetControl('GSHP')),'GSHP',stLigHP,TOBHP);
    PrepareGrids;
    //
    ChargeTOBS (CodeAffaire);
    XX.StopK2000 ;
    //
    XX.lAide.Caption := 'Récupérations des consommations ...';
    XX.StartK2000 ;
    XX.refresh;
    RecupereConsos;
    XX.StopK2000 ;
    //
    XX.lAide.Caption := 'Récupération avancements sous traitance ...';
    XX.StartK2000 ;
    XX.refresh;
    CalculSousTraitance;
    XX.StopK2000 ;
    //
    XX.lAide.Caption := 'Calculs finaux en cours...';
    XX.StartK2000 ;
    XX.refresh;
    CalculFinal;
    //
    ChargeGrids;
    TPageControl(GetControl('PG1')).ActivePageIndex := 0; // on positionne le premier onglet comme courant
    ZeCurSaisie := TheListGrid.findGrille('GSPP');
    TPageControl(GetControl('PG1')).OnChanging := BeforeChangePage;
    TPageControl(GetControl('PG1')).OnChange := ChangePage;
    SetEventGrid (True);
    TFvierge(Ecran).HMTrad.ResizeGridColumns(GSPP);
    TFvierge(Ecran).HMTrad.ResizeGridColumns(GSMOI);
    TFvierge(Ecran).HMTrad.ResizeGridColumns(GSFI);
    TFvierge(Ecran).HMTrad.ResizeGridColumns(GSST);
    TFvierge(Ecran).HMTrad.ResizeGridColumns(GSHP);
    //
    EnterGrid ('GSPP');
  FINALLY
    XX.StopK2000 ;
    XX.free;
  END;
end ;

procedure TOF_BSAISIEAVANCPOC.OnArgument (S : String ) ;

  function GetLibelleAffaire (TheAff : string) : string;
  var QQ: TQuery;
  begin
    Result := '';
    QQ := OpenSql ('SELECT AFF_LIBELLE FROM AFFAIRE WHERE AFF_AFFAIRE="'+TheAff+'"',True,1,'',True);
    if not QQ.eof then
    begin
      Result := QQ.fields[0].asstring;
    end;
    ferme (QQ);
  end;

  procedure ControleChamp (ChampMul,ValMul : string);
  begin
    if GetControl(ChampMul) <> nil then
    begin
      SetControlText(ChampMul,ValMul);
      if ChampMul = 'BMOIS' then
      begin
        SetControlText('MAAAA',Copy(ValMul,1,4));
        SetControlText('MMM',Copy(ValMul,5,2));
        DebutMois := EncodeDate(StrToInt(Copy(ValMul,1,4)),StrToInt(Copy(ValMul,5,2)),1);
        FinMois := FINDEMOIS(DebutMois);
      end else if ChampMul = 'BAFFAIRE' then
      begin
        SetControlText('TAFFAIRELIBELLE',GetLibelleAffaire(ValMul));
        CodeAffaire:=ValMul;
        BTPCodeAffaireDecoupe (CodeAffaire,Aff0,Aff1,Aff2,Aff3,Avenant,tacreat,false);
      end;
    end;
  end;

  function GetFamillesIndirect: string;
  var QQ : TQuery;
  begin
    result := '';
    QQ := OpenSql ('SELECT BM7_PRODFI FROM BDEPENSESPOC WHERE BM7_CODE="1"',True,1,'',True);
    if Not QQ.eof then
    begin
      Result := QQ.fields[0].AsString;
    end;
    Ferme(QQ);
  end;

var i : Integer;
    STSt,Critere,ChampMul,ValMul : string;
    X : Integer;
    UneImage : THImage;
    TheBm : Graphics.TBitmap;
begin
  Inherited ;
  DataModif := false;
  ActionResult := taModif;
  FamFraisIndirect := GetFamillesIndirect;
  //
  fGestionAff := TAffichageDoc.create(ecran);
  fGestionAff.gestion := TtaNormal;
  fGestionAff.Action := taModif;
  //
  TOBSTRUCTLIG := TOB.Create (' LA STRUCTURE LIGNE',nil,-1);
  TTS := TOB.Create ('LES PHASES',nil,-1);
  TTD := TOB.Create ('BAVANCTECH',nil,-1);

  MemoriseLigneSaisie;
  //
  GSPP := THGrid(getControl('GSPP'));
  GSMOI := THGrid(getControl('GSMOI'));
  GSFI := THGrid(getControl('GSFI'));
  GSST:= THGrid(getControl('GSST'));
  GSHP:= THGrid(getControl('GSHP'));
  //
  STSt := S;
  //
  ImList := TImageList.Create(Ecran);
  ImList.Height := 16;
  ImList.Width := 16;
  //
  UneImage := THImage(GetControl('IMG1'));
  TheBm := graphics.TBitMap(UneImage.Picture.Bitmap);
  ImList.add(TheBM,nil);
  UneImage := THImage(GetControl('IMG2'));
  ImList.add(graphics.TBitMap(UneImage.Picture.Bitmap ),nil);
  UneImage := THImage(GetControl('IMG3'));
  ImList.add(graphics.TBitMap(UneImage.Picture.Bitmap ),nil);
  UneImage := THImage(GetControl('IMG4'));
  ImList.add(graphics.TBitMap(UneImage.Picture.Bitmap ),nil);
  UneImage := THImage(GetControl('IMG5'));
  ImList.add(graphics.TBitMap(UneImage.Picture.Bitmap ),nil);
  UneImage := THImage(GetControl('IMG6'));
  ImList.add(graphics.TBitMap(UneImage.Picture.Bitmap ),nil);
  //
  Repeat
    Critere:=uppercase(ReadTokenSt(STSt)) ;
    valMul := '';
    if Critere<>'' then
    begin
      x:=pos('=',Critere);
      if x<>0 then
      begin
        ChampMul:=copy(Critere,1,x-1);
        ValMul:=copy(Critere,x+1,length(Critere));
      end else
      begin
        ChampMul := Critere;
      end;
      ControleChamp(ChampMul, ValMul);
    end;
  until  Critere='';

  FFP := '#,##';
  if V_PGI.OkDecP > 0 then
  begin
    FFP := '0.';
    for i := 1 to V_PGI.OkDecP - 1 do
    begin
      FFP := FFP + '0';
    end;
    FFP := FFP + '0';
  end;
  //
  FFQ := '#,##';
  if V_PGI.OkDecQ > 0 then
  begin
    FFQ := '0.';
    for i := 1 to V_PGI.OkDecQ - 1 do
    begin
      FFQ := FFQ + '0';
    end;
    FFQ := FFQ + '0';
  end;
  //
end ;

procedure TOF_BSAISIEAVANCPOC.OnClose ;
begin
  if DataModif then
  begin
    if PgiAsk ('ATTENTION : Vous avez modifiés des données.#13#10 Validez-vous la sortie sans sauvegarder ?')<> MrYes then
    begin
      ecran.modalresult := 0;
      exit;
    end;
  end;
  TOBSTRUCTLIG.free;
  TTS.free;
  TTD.Free;
  ImList.Free;
  Inherited ;
end ;

procedure TOF_BSAISIEAVANCPOC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BSAISIEAVANCPOC.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BSAISIEAVANCPOC.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
  Aff:=THEdit(GetControl('BAFFAIRE'));
  Aff0:=THEdit(GetControl('BAFFAIRE0'));
  Aff1:=THEdit(GetControl('BAFFAIRE1'));
  Aff2:=THEdit(GetControl('BAFFAIRE2')); Aff3:=THEdit(GetControl('BAFFAIRE3'));
  Aff4:=THEdit(GetControl('BAVENANT'));
end;

procedure TOF_BSAISIEAVANCPOC.afficheLigne(zeGS : ThGrid;ZeTOB : TOB;ZeListe : string;II : integer);
begin
  ZeTOB.PutLigneGrid(zeGS,II+1,False,False,ZeListe); 
end;

procedure TOF_BSAISIEAVANCPOC.afficheLaGrid(TGS : TGSaisie);
var II : Integer;
begin
  TGS.fGS.RowCount := TGS.fTOB.detail.count +1; if TGS.fGS.RowCount = 1 then TGS.fGS.RowCount := 2;
  if TGS.fTOB.detail.count = 0 then Exit;
  for II:= 0 to TGS.fTOB.detail.count -1 do
  begin
    afficheLigne(TGS.fGS,TGS.fTOB.detail[II],TGS.fListSaisie,II);
  end;
end;

procedure TOF_BSAISIEAVANCPOC.afficheLaGrid(Nomgrille : string);
var TGS : TGSaisie;
begin
  TGS := TheListGrid.findGrille(Nomgrille);
  afficheLaGrid(TGS);
end;

procedure TOF_BSAISIEAVANCPOC.ChargeGrids;
begin
  afficheLaGrid('GSPP');
  afficheLaGrid('GSMOI');
  afficheLaGrid('GSFI');
  afficheLaGrid('GSST');
  afficheLaGrid('GSHP');
end;



function TOF_BSAISIEAVANCPOC.CreTOBLigne : TOB;
begin
  Result := TOB.Create ('BAVANCTECHDET',nil,-1);
  //
  Result.SetString ('BM8_AFFAIRE',CodeAffaire);
  Result.SetString ('BM8_AFFAIRE0',Aff0);
  Result.SetString ('BM8_AFFAIRE1',Aff1);
  Result.SetString ('BM8_AFFAIRE2',Aff2);
  Result.SetString ('BM8_AFFAIRE3',Aff3);
  Result.SetString ('BM8_AVENANT',Avenant);
  //
  Result.SetString ('BM8_NATUREPIECEG',GetControlText('BNATURE'));
  Result.SetString ('BM8_SOUCHE',GetControlText('BSOUCHE'));
  Result.SetInteger ('BM8_NUMERO',StrToInt(GetControlText('BNUMERO')));
  Result.SetInteger ('BM8_INDICEG',StrToInt(GetControlText('BINDICEG')));
  Result.SetString ('BM8_MOIS',GetControlText('BMOIS'));
  AddlesChampsSupSaisie (Result);
  Result.SetString ('NEWONE','X');
  Result.SetString ('LIG','X');
end;



function TOF_BSAISIEAVANCPOC.FindLigne (TheType,CurrentPhase,CodeAnal,CodeMarche : string) : TOB;
begin
  Result := nil;
  if TheType = '000' then Result := TOBPP.FindFirst(['BM8_PHASETRA','BM8_FAMILLENIV2'],[CurrentPhase,CodeAnal],true)
  else if TheType = '001' then Result := TOBMOI.FindFirst(['BM8_PHASETRA','BM8_FAMILLENIV2'],[CurrentPhase,CodeAnal],true)
  else if TheType = '002' then Result := TOBFI.FindFirst(['BM8_PHASETRA','BM8_FAMILLENIV2'],[CurrentPhase,CodeAnal],true)
  else if TheType = '003' then Result := TOBST.FindFirst(['BM8_PHASETRA','BM8_CODEMARCHE'],[CurrentPhase,CodeMarche],true);
end;

function TOF_BSAISIEAVANCPOC.FindLigneConso (TheType,CurrentPhase,CodeAnal,CodeMarche : string) : TOB;
begin
  Result := nil;
  if TheType = '000' then Result := TOBPP.FindFirst(['BM8_PHASETRA','BM8_FAMILLENIV2'],[CurrentPhase,CodeAnal],true)
  else if TheType = '001' then Result := TOBMOI.FindFirst(['BM8_PHASETRA','BM8_FAMILLENIV2'],[CurrentPhase,CodeAnal],true)
  else if TheType = '002' then Result := TOBFI.FindFirst(['BM8_FAMILLENIV2'],[CodeAnal],true)
  else if TheType = '003' then Result := TOBST.FindFirst(['BM8_PHASETRA','BM8_CODEMARCHE'],[CurrentPhase,CodeMarche],true);
end;


function TOF_BSAISIEAVANCPOC.FindEmplacementLigne (TheType,CurrentPhase : string) : TOB;
var TheLast : TOB;
begin
  Result := nil;
  TheLast := nil;
  if TheTYpe = '000' then TheLast := TOBPP.FindFirst(['BM8_PHASETRA'],[CurrentPhase],true)
  else if TheTYpe = '001' then TheLast := TOBMOI.FindFirst(['BM8_PHASETRA'],[CurrentPhase],true)
  else if TheTYpe = '002' then TheLast := TOBFI.FindFirst(['BM8_PHASETRA'],[CurrentPhase],true)
  else if TheTYpe = '003' then TheLast := TOBST.FindFirst(['BM8_PHASETRA'],[CurrentPhase],true);
  repeat
    if TheLast <> nil then
    begin
      Result := TheLast;
      if TheTYpe = '000' then TheLast := TOBPP.FindNext(['BM8_PHASETRA'],[CurrentPhase],true)
      else if TheTYpe = '001' then TheLast := TOBMOI.FindNext(['BM8_PHASETRA'],[CurrentPhase],true)
      else if TheTYpe = '002' then TheLast := TOBFI.FindNext(['BM8_PHASETRA'],[CurrentPhase],true)
      else if TheTYpe = '003' then TheLast := TOBST.FindNext(['BM8_PHASETRA'],[CurrentPhase],true);
    end;
  until TheLast = nil;
end;

function TOF_BSAISIEAVANCPOC.GetEmplacementLigne (TheType,CurrentPhase,FamilleNiv2 : string) : TOB;
var ThePhase,TheTOBP,TheCUrr : TOB;
    II : Integer;
begin
  Result := nil;
  if TheTYpe = '000' then
  begin
    TheTOBP := TOBPP;
    ThePhase := TOBPP.FindFirst(['BM8_PHASETRA'],[CurrentPhase],true)
  end else if TheTYpe = '001' then
  begin
    TheTOBP := TOBMOI;
    ThePhase := TOBMOI.FindFirst(['BM8_PHASETRA'],[CurrentPhase],true)
  end else if TheTYpe = '002' then
  begin
    TheTOBP := TOBFI;
    ThePhase := TOBFI.FindFirst(['BM8_PHASETRA'],[CurrentPhase],true)
  end else if TheTYpe = '003' then
  begin
    TheTOBP := TOBST;
    ThePhase := TOBST.FindFirst(['BM8_PHASETRA'],[CurrentPhase],true);
  end;
  if ThePhase = nil Then exit;
  for II := ThePhase.GetIndex+1 to TheTOBP.detail.count -1 do
  begin
    TheCUrr := TheTOBP.detail[II];
    if (TheCUrr.GetString('BM8_PHASETRA')<>CurrentPhase) and (TheCUrr.GetString('BM8_PHASETRA')<>'') then
    begin
      result := TheCurr;
      break;
    end;
    if (TheCUrr.GetString('BM8_FAMILLENIV2') > FamilleNiv2) then
    begin
      result := TheCurr;
      break;
    end;
  end;
end;


procedure TOF_BSAISIEAVANCPOC.ConstitueLesTOBS (TTS,TTD,TTL,TTOP : TOB);

  procedure AddChampsTOBNiv (TN : TOB);
  begin
    TN.AddChampSupValeur('PN1',0);
    TN.AddChampSupValeur('PN2',0);
    TN.AddChampSupValeur('PN3',0);
    TN.AddChampSupValeur('PN4',0);
    TN.AddChampSupValeur('PN5',0);
    TN.AddChampSupValeur('PN6',0);
    TN.AddChampSupValeur('PN7',0);
    TN.AddChampSupValeur('PN8',0);
    TN.AddChampSupValeur('PN9',0);
  end;

  function AddNiv(TOBPere,TL : TOB) : TOB;
  var TT,TOBZ : TOB;
      Niv : integer;
  begin
    TOBZ := TOBPere.detail[TOBPere.Detail.count-1];
    result := TOB.Create('UN NIVEAU',TOBPere,-1);
    AddChampsTOBNiv (result);
    result.Dupliquer(TOBZ,false,true);
    Niv := TL.GetInteger('GL_NIVEAUIMBRIC'); if IsOuvrage(TL) then Inc(Niv);
    result.SetInteger('PN'+InttoStr(Niv),TL.GetInteger('GL_NUMORDRE'))
  end;

  procedure RemonteNiv(TOBpere : TOB);
  begin
    TOBPere.detail[TOBPere.Detail.count-1].Free;
  end;

  procedure AjouteLigneP(TL : TOB;TheType,CurrentPhase : string;NiveauImbric : integer);
  var TF,TP : TOB;
      Index : Integer;
  begin
    TF := CreTOBLigne;
    if IsOuvrage(TL) then TF.SetString('TYPELIGNE','OUV')
                     else TF.SetString('TYPELIGNE',TL.GetString('GL_TYPELIGNE'));
    if (not IsParagraphe(TL)) and (not IsOuvrage(TL)) then TF.SetInteger('NIVEAU',NiveauImbric+1)
                                                      else TF.SetInteger('NIVEAU',NiveauImbric);
    TF.SetString('BM8_PHASETRA',CurrentPhase);
    TF.SetString('BM8_CODEMARCHE','');

    TF.SetString('BM8_LIBELLE',TL.GetString('GL_LIBELLE'));
    TF.SetString('BM8_UNITE','');
    TF.SetString('NEWONE','-');
    if not IsParagraphe(TL) then
    begin
      TP := FindEmplacementLigne (TheType,CurrentPhase);
      if TP = nil then Index := -1
                  else Index := TP.GetIndex+1;
    end else
    begin
      Index := -1;
    end;
    if TheType = '000' then TF.Changeparent(TOBPP,Index)
    else if TheType = '001' then TF.Changeparent(TOBMOI,Index)
    else if TheType = '002' then TF.Changeparent(TOBFI,Index)
    else if TheType = '003' then TF.Changeparent(TOBST,Index);
  end;

  procedure AjouteLigne(TL : TOB;TheType,CurrentPhase : string;NiveauImbric : integer);
  var TF,TP : TOB;
      Index : Integer;
  begin

    if TheType <> '003' then TF := FindLigne(TheType,CurrentPhase,TL.GetString('GL_FAMILLENIV2'),'')
                        else TF := FindLigne(TheType,CurrentPhase,'',TL.GetString('GL_CODEMARCHE'));
    //
    if (TF = nil) then
    begin
      TF := CreTOBLigne;
      if IsOuvrage(TL) then TF.SetString('TYPELIGNE','OUV')
                       else TF.SetString('TYPELIGNE',TL.GetString('GL_TYPELIGNE'));
      if (not IsParagraphe(TL)) and (not IsOuvrage(TL)) then TF.SetInteger('NIVEAU',NiveauImbric+1)
                                                        else TF.SetInteger('NIVEAU',NiveauImbric);
      TF.SetString('BM8_PHASETRA',CurrentPhase);
      TF.SetString('BM8_CODEMARCHE',TL.GetString('GL_CODEMARCHE'));

      if (not IsParagraphe(TL)) and (not IsOuvrage(TL)) and (TheType <>'003') then TF.SetString('BM8_FAMILLENIV2',TL.GetString('GL_FAMILLENIV2'));

      if Pos(Copy(TF.GetString('TYPELIGNE'),1,2),'DP;TP')>0 then
      begin
        TF.SetString('BM8_LIBELLE',TL.GetString('GL_LIBELLE'));
      end else
      begin
        if TheType <>'003' then TF.SetString('BM8_LIBELLE',TL.GetString('LIBELLE'))
                           else TF.SetString('BM8_LIBELLE',TL.GetString('GL_LIBELLE'));
      end;
      TF.SetString('BM8_UNITE','');
      TF.SetString('NEWONE','-');
      if not IsParagraphe(TL) then
      begin
        if TheType <>'003' then
        begin
          TP := GetEmplacementLigne (TheType,CurrentPhase,TF.GetString('BM8_FAMILLENIV2'));
        end else
        begin
          TP := FindEmplacementLigne (TheType,CurrentPhase) ;
        end;
        if TP = nil then Index := -1
                    else Index := TP.GetIndex+1;
      end else
      begin
        Index := -1;
      end;
      if TheType = '000' then TF.Changeparent(TOBPP,Index)
      else if TheType = '001' then TF.Changeparent(TOBMOI,Index)
      else if TheType = '002' then TF.Changeparent(TOBFI,Index)
      else if TheType = '003' then TF.Changeparent(TOBST,Index);
    end;
    TF.SetDouble('BM8_QTEMARCHE',TF.GetDouble('BM8_QTEMARCHE')+TL.GetDouble('GL_QTEFACT')-TL.GetDouble('QTETRANSFERT'));
    TF.SetDouble('BM8_MTMARCHE',TF.GetDouble('BM8_MTMARCHE')+TL.GetDouble('GL_MONTANTPA')-TL.GetDouble('MTTRANSFERT'));
  end;

  procedure AjouteLigneOUV(TL,TD : TOB;TheType,CurrentPhase : string;NiveauImbric : integer);
  var TF,TP : TOB;
      Index : Integer;
      QteReajuste,MtReajuste : double;
  begin
    if TheType <> '003' then TF := FindLigne(TheType,CurrentPhase,TD.GetString('BOP_FAMILLENIV2'),TL.GetString('GL_CODEMARCHE'))
                        else TF := FindLigne(TheType,CurrentPhase,'',TL.GetString('GL_CODEMARCHE'));
    if (TF = nil) then
    begin
      TF := CreTOBLigne;
      TF.SetString('TYPELIGNE','ART');
      TF.SetString('BM8_PHASETRA',CurrentPhase);
      TF.SetString('BM8_CODEMARCHE',TL.GetString('GL_CODEMARCHE'));
      if TheType <> '003' then TF.SetString('BM8_FAMILLENIV2',TD.GetString('BOP_FAMILLENIV2'));
      TF.SetString('BM8_LIBELLE',TD.GetString('LIBELLE'));
      TF.SetString('BM8_UNITE','');
      TF.SetInteger('NIVEAU',NiveauImbric+1);
      TF.SetString('NEWONE','-');
      if TheType <>'003' then
      begin
        TP := GetEmplacementLigne (TheType,CurrentPhase,TF.GetString('BM8_FAMILLENIV2'));
      end else
      begin
        TP := FindEmplacementLigne (TheType,CurrentPhase) ;
      end;
      if TP = nil then Index := -1
                  else Index := TP.GetIndex;
      if TheType = '000' then TF.Changeparent(TOBPP,Index)
      else if TheType = '001' then TF.Changeparent(TOBMOI,Index)
      else if TheType = '002' then TF.Changeparent(TOBFI,Index)
      else if TheType = '003' then TF.Changeparent(TOBST,Index);
    end;
    //
    QteReajuste := ARRONDI(TD.GetDouble('BOP_QTEFACT') / (TL.GetDouble('GL_QTEFACT')) * ( TL.GetDouble('GL_QTEFACT') - TL.GetDouble('QTETRANSFERT')),V_PGI.OkDecQ);
    MtReajuste := ARRONDI(TD.GetDouble('BOP_MONTANTPA') / (TL.GetDouble('GL_QTEFACT')) * ( TL.GetDouble('GL_QTEFACT') - TL.GetDouble('QTETRANSFERT')),V_PGI.OkDecP);
    TF.SetDouble('BM8_QTEMARCHE',TF.GetDouble('BM8_QTEMARCHE')+QteReajuste);
    TF.SetDouble('BM8_MTMARCHE',TF.GetDouble('BM8_MTMARCHE')+MtReajuste);
  end;

  function FindPhase (TS : TOB) : string;
  var TheTS : TOB;
  begin
    Result := '';
    TheTS := TTS.findfirst(['BPC_PN1','BPC_PN2','BPC_PN3','BPC_PN4','BPC_PN5','BPC_PN6','BPC_PN7','BPC_PN8','BPC_PN9'],
                            [TS.GetInteger('PN1'),
                            TS.GetInteger('PN2'),
                            TS.GetInteger('PN3'),
                            TS.GetInteger('PN4'),
                            TS.GetInteger('PN5'),
                            TS.GetInteger('PN6'),
                            TS.GetInteger('PN7'),
                            TS.GetInteger('PN8'),
                            TS.GetInteger('PN9')],true);
    if TheTS <> nil then Result := TheTS.GetString('BPC_PHASETRA');
  end;

  function FindTypeAV(TLC : TOB ) : string;
  var TT : TOB;
  begin
    Result := '';
    TT := TTS.FindFirst(['BPC_PHASETRA'],[TLC.GetString('BM8_PHASETRA')],true);
    if TT <> nil then
    begin
      Result := TT.GetString('TYPE');
    end;
  end;

var TOBPere,TT,TD : TOB;
    II : Integer;
    TL,TX : TOB;
    TheType : string;
    CurrentPhase : string;
begin
  TOBPere := TOB.Create ('UNE STRUCTURE',nil,-1);
  TT := TOB.Create('UN NIVEAU',TOBPere,-1);
  AddChampsTOBNiv (TT);
  //
  // on va constituer la structure via le document Contre étude
  TheType := '';
  for II := 0 to TTL.Detail.Count -1 do
  begin
    XX.MoveK2000;
    TL := TTL.detail[II];
    if TL.GetString('GL_TYPELIGNE')='DP1' then
    begin
      TheType := TL.GetString('GL_TYPECADENCE'); // on récupère la nature d'éclatement via le paragraphe de niveau 1
      TT.SetInteger('PN1',TL.getInteger('GL_NUMORDRE'));
      CurrentPhase := FindPhase(TT);
      AjouteLigneP(TL,TheType,CurrentPhase,TL.getInteger('GL_NIVEAUIMBRIC'));
    end else if TL.getString('GL_TYPELIGNE')='TP1' then
    begin
      AjouteLigneP(TL,TheType,CurrentPhase,TL.getInteger('GL_NIVEAUIMBRIC'));
      TheType := '';
      TT.SetInteger('PN1',0);
      CurrentPhase := '';
    end else
    begin
      if TheType <> '' then
      begin
        //
        if TL.GetString('TYPETRANSFERT')='X' then continue;
        //
        if (IsDebutParagraphe(TL)) then
        begin
          TX := AddNiv(TOBPere,TL);
          CurrentPhase := FindPhase(TX);
          AjouteLigneP(TL,TheType,CurrentPhase,TL.getInteger('GL_NIVEAUIMBRIC'));
        end else if IsFinParagraphe(TL) then
        begin
          AjouteLigneP(TL,TheType,CurrentPhase,TL.getInteger('GL_NIVEAUIMBRIC'));
          RemonteNiv(TOBPere);
          CurrentPhase := FindPhase(TOBPere.detail[TOBPere.Detail.count-1]);
        end else
        begin
          if ISOuvrage(TL) then
          begin
            TX := AddNiv(TOBPere,TL);
            CurrentPhase := FindPhase(TX);
            if CurrentPhase = '' then
            begin
              XX.lAide.Caption := 'Phase a blanc dans Sous détail ligne '+TL.GetString('GL_NUMLIGNE') ;
            end;
            AjouteLigneP(TL,TheType,CurrentPhase,TL.getInteger('GL_NIVEAUIMBRIC')+1);
            //
            TD := TTOP.FindFirst(['BOP_NUMORDRE'],[TL.GetInteger('GL_NUMORDRE')],true);
            repeat
              if TD = nil then break;
              AjouteLigneOUV(TL,TD,TheType,CurrentPhase,TL.getInteger('GL_NIVEAUIMBRIC')+1);
              TD := TTOP.FindNext(['BOP_NUMORDRE'],[TL.GetInteger('GL_NUMORDRE')],true);
            until TD = nil;
            //
            RemonteNiv(TOBPere);
            CurrentPhase := FindPhase(TOBPere.detail[TOBPere.Detail.count-1]);
          end else
          begin
            if CurrentPhase = '' then XX.lAide.Caption := 'Phase a blanc dans ligne'+TL.GetString('GL_NUMLIGNE');
            AjouteLigne(TL,TheType,CurrentPhase,TL.getInteger('GL_NIVEAUIMBRIC'));
          end;
        end;
      end;
    end;
  end;
  // récup de l'avencement saisi sur le mois
  for II := 0 to TTD.detail.count -1 do
  begin
    TL := TTD.detail[II];
    TheType := FindTypeAV(TL);
    if TheType <> '003' then TD := FindLigne (TheType,TL.GetString('BM8_PHASETRA'),TL.GetString('BM8_FAMILLENIV2'),TL.GetString('BM8_CODEMARCHE'))
                        else TD := FindLigne (TheType,TL.GetString('BM8_PHASETRA'),'',TL.GetString('BM8_CODEMARCHE'));
    if TD <> nil then
    begin
      TD.SetDouble('BM8_MTPREC',TL.GetDouble('BM8_MTPREC'));
      TD.SetDouble('BM8_MTMOIS',TL.GetDouble('BM8_MTMOIS'));
      TD.SetDouble('BM8_MTCUMUL',TL.GetDouble('BM8_MTCUMUL'));
      //
      TD.SetDouble('BM8_POURCENTAVANC',TL.GetDouble('BM8_POURCENTAVANC'));
      //
      TD.SetDouble('BM8_QTEMOIS',TL.GetDouble('BM8_QTEMOIS'));
      TD.SetDouble('BM8_QTEPREC',TL.GetDouble('BM8_QTEPREC'));
      TD.SetDouble('BM8_QTECUMUL',TL.GetDouble('BM8_QTECUMUL'));
      //
      TD.SetDouble('BM8_RADQTE',TL.GetDouble('BM8_RADQTE'));
      TD.SetDouble('BM8_RADMT',TL.GetDouble('BM8_RADMT'));
      TD.SetDouble('BM8_RADPU',TL.GetDouble('BM8_RADPU'));
      //
      TD.SetDouble('BM8_QTECONSOPRE',TL.GetDouble('BM8_QTECONSOPRE'));
      TD.SetDouble('BM8_MTCONSOPRE',TL.GetDouble('BM8_QTECONSOPRE'));
    end;
  end;
  //
  // TOBMOI.detail.Sort('TYPELIGNE;BM8_PHASETRA;BM8_FAMILLENIV2');
  //
end;

function TOF_BSAISIEAVANCPOC.ChargeTOBS  (CodeAffaire : string) : boolean;


  function WhereDoc (racine : string) : string;
  begin
    Result := Racine + '_NATUREPIECEG="'+GetControlText('BNATURE')+'" AND '+
              Racine + '_SOUCHE="'+GetControlText('BSOUCHE')+'" AND '+
              Racine + '_NUMERO='+GetControlText('BNUMERO')+' AND '+
              Racine + '_INDICEG='+GetControlText('BINDICEG');
  end;

  procedure AssocieTypePhase (TTS : TOB) ;
  var SQL : String;
      QQ : TQuery;
  begin
    SQL := 'SELECT GL_TYPECADENCE '+
           'FROM LIGNE '+
           'WHERE '+ WhereDoc('GL')+' AND '+
           'GL_NUMORDRE ='+TTS.GetString('BPC_PN1');
    QQ := OpenSQL(SQL,True,1,'',true);
    if not QQ.eof then
    begin
      TTS.SetString('TYPE',QQ.fields[0].AsString);
    end;
    Ferme(QQ);
  end;

  function SaisieMoisOk  : Boolean;
  var QQ : TQuery;
      Mois : string;
  begin
    Mois := GetControlText('BMOIS');
    Result := true;
    QQ := OpenSQL('SELECT BM9_CLOTURE FROM BAVANCTECH WHERE '+ WhereDoc ('BM9')+' AND BM9_MOIS = '+Mois,True,1,'',true);
    if Not QQ.eof then
    begin
      if QQ.Fields[0].AsString = 'X' then ActionResult := taConsult;
    end;
    Ferme(QQ);
    //
    if ExisteSQL ('SELECT 1 FROM BAVANCTECH WHERE '+ WhereDoc ('BM9')+' AND BM9_MOIS > '+Mois) then Result := false;
  end;

var QQ : Tquery;
    SQL : String;
    TTL,TTOP : TOB;
    Mois : string;
    II : Integer;
begin
  Result := True;

  if not SaisieMoisOk then
  begin
    SetControlVisible('Bdelete',false);
  end;
  if ActionResult = TaConsult then
  begin
    SetInConsultation;
  end;
  TTL := TOB.Create ('LE DOCUMENT',nil,-1);
  TTOP := TOB.Create ('OUVRAGES PLAT',nil,-1);
  //
  SQL := 'SELECT PHASESCHANTIER.*,"" AS TYPE '+
         'FROM PHASESCHANTIER '+
         'WHERE BPC_AFFAIRE="'+CodeAffaire+'"';
  QQ := OpenSQL(SQL,True,-1,'',True);
  if not QQ.eof then TTS.loadDetailDb('PHASESCHANTIER','','',QQ,false);
  Ferme(QQ);
  //
  for II := 0 to TTS.Detail.count -1 do
  begin
    AssocieTypePhase (TTS.detail[II]);
  end;
  //
  Mois := GetControlText('BMOIS');
  TRY
    SQL := 'SELECT 1 FROM BAVANCTECH WHERE '+ WhereDoc ('BM9')+' AND BM9_MOIS = '+Mois;
    if not ExisteSQL(SQL) then
    begin
      // recherche du dernier mois saisi
      QQ := OpenSQL('SELECT MAX(BM9_MOIS) FROM BAVANCTECH WHERE '+WhereDoc('BM9'),True,1,'',True);
      if not QQ.eof then
      begin
        Mois := QQ.Fields[0].AsString;
      end else
      begin
        Mois := '';
      end;
      ferme (QQ);
    end;
    //
    if Mois <> '' then
    begin
      SQL := 'SELECT * FROM BAVANCTECH WHERE '+WhereDoc ('BM9')+' AND BM9_MOIS = '+Mois;
      if ExisteSQL(SQL) then
      begin
        QQ := OpenSQL(SQL,True,-1,'',true);
        //
        if not QQ.eof then
        begin
          TTD.SelectDB('',QQ);
          ferme (QQ);
          TTD.SetString ('BM9_MOIS',GetControlText('BMOIS'));
          //
          SQL := STAVANCSQL+' WHERE '+ WhereDoc('BM8')+' AND BM8_MOIS="'+Mois+'"';
          QQ := OpenSQL(SQL,True,-1,'',true);
          if not QQ.Eof then TTD.LoadDetailDB('BAVANCTECHDET','','',QQ,false);
        end;
        Ferme(QQ);
      end;
    end else
    begin
      TTD.SetString ('BM9_AFFAIRE',CodeAffaire);
      TTD.SetString ('BM9_AFFAIRE0',Aff0);
      TTD.SetString ('BM9_AFFAIRE1',Aff1);
      TTD.SetString ('BM9_AFFAIRE2',Aff2);
      TTD.SetString ('BM9_AFFAIRE3',Aff3);
      TTD.SetString ('BM9_AVENANT',Avenant);
      //
      TTD.SetString ('BM9_NATUREPIECEG',GetControlText('BNATURE'));
      TTD.SetString ('BM9_SOUCHE',GetControlText('BSOUCHE'));
      TTD.SetInteger ('BM9_NUMERO',StrToInt(GetControlText('BNUMERO')));
      TTD.SetInteger ('BM9_INDICEG',StrToInt(GetControlText('BINDICEG')));
      TTD.SetString ('BM9_MOIS',GetControlText('BMOIS'));
    end;

    For II := 0 to TTD.detail.Count -1 do
    begin
      if Mois <> GetControlText('BMOIS') then  // reprise d'un mois antérieur
      begin
        TTD.detail[II].SetString('BM8_MOIS', GetControlText('BMOIS'));
        TTD.detail[II].SetDouble('BM8_QTEPREC',TTD.detail[II].GetDouble('BM8_QTECUMUL'));
        TTD.detail[II].SetDouble('BM8_QTEMOIS',0);
        TTD.detail[II].SetDouble('BM8_MTPREC',TTD.detail[II].GetDouble('BM8_MTCUMUL'));
        TTD.detail[II].SetDouble('BM8_MTMOIS',0);
        TTD.detail[II].SetDouble('BM8_RADQTE',0);
        TTD.detail[II].SetDouble('BM8_RADMT',0);
        TTD.detail[II].SetDouble('BM8_RADPU',0);
        TTD.detail[II].SetDouble('BM8_QTECONSOPRE',TTD.detail[II].GetDouble('BM8_QTECONSOPRE')+TTD.detail[II].GetDouble('BM8_QTECONSOMOIS') );
        TTD.detail[II].SetDouble('BM8_MTCONSOPRE',TTD.detail[II].GetDouble('BM8_MTCONSOPRE')+TTD.detail[II].GetDouble('BM8_MTCONSOMOIS'));
      end;
      TTD.detail[II].SetDouble('BM8_QTEMARCHE',0);
      TTD.detail[II].SetDouble('BM8_MTMARCHE',0);
    end;
    //
    SQL := 'SELECT LIGNE.*,'+
          '('+
          'SELECT SUM(BT3_MTTRANSFERT) '+
          'FROM BTRFDETAIL '+
          'WHERE '+
          'BT3_NATUREPIECEG=GL_NATUREPIECEG AND '+
          'BT3_SOUCHE=GL_SOUCHE AND '+
          'BT3_NUMERO=GL_NUMERO AND '+
          'BT3_INDICEG=GL_INDICEG AND '+
          'BT3_NUMORDRE=GL_NUMORDRE '+
          ') AS MTTRANSFERT,' +
          '('+
          'SELECT DISTINCT IIF(BT3_TYPELIGNETRF="000" AND BT3_CONTREP="-",NULL,BT3_CONTREP) '+
          'FROM BTRFDETAIL '+
          'WHERE '+
          'BT3_NATUREPIECEG=GL_NATUREPIECEG AND '+
          'BT3_SOUCHE=GL_SOUCHE AND '+
          'BT3_NUMERO=GL_NUMERO AND '+
          'BT3_INDICEG=GL_INDICEG AND '+
          'BT3_NUMORDRE=GL_NUMORDRE AND '+
          'BT3_TYPELIGNETRF="001" '+
          ') AS TYPETRANSFERT, '+
          '('+
          'SELECT SUM(BT3_QTETRANSFERT) '+
          'FROM BTRFDETAIL '+
          'WHERE '+
          'BT3_NATUREPIECEG=GL_NATUREPIECEG AND '+
          'BT3_SOUCHE=GL_SOUCHE AND '+
          'BT3_NUMERO=GL_NUMERO AND '+
          'BT3_INDICEG=GL_INDICEG AND '+
          'BT3_NUMORDRE=GL_NUMORDRE '+
          ') AS QTETRANSFERT,' +
          '(SELECT CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="FN2" AND CC_CODE=GL_FAMILLENIV2) AS LIBELLE FROM LIGNE '+
          'WHERE '+ WhereDoc('GL')+
          ' ORDER BY GL_NUMLIGNE';
    QQ := OpenSQL(SQL,true,-1,'',true);
    if not QQ.Eof then
    begin
      TTL.LoadDetailDB('LIGNE','','',QQ,false);
    end;
    Ferme(QQ);
    //
    SQL := 'SELECT LIGNEOUVPLAT.*,(SELECT CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="FN2" AND CC_CODE=BOP_FAMILLENIV2) AS LIBELLE FROM LIGNEOUVPLAT WHERE '+ WhereDoc('BOP')+' ORDER BY BOP_NUMORDRE';
    QQ := OpenSQL(SQL,true,-1,'',true);
    if not QQ.Eof then
    begin
      TTOP.LoadDetailDB('LIGNEOUVPLAT','','',QQ,false);
    end;
    Ferme(QQ);
    //
    ConstitueLesTOBS (TTS,TTD,TTL,TTOP);
    TTD.ClearDetail;
  FINALLY
    TTL.free;
    TTOP.Free;
  END;

end;

procedure TOF_BSAISIEAVANCPOC.PrepareGrids;
begin
  PrepareGrid('GSPP');
  PrepareGrid('GSMOI');
  PrepareGrid('GSFI');
  PrepareGrid('GSST');
  PrepareGrid('GSHP');
end;


procedure TOF_BSAISIEAVANCPOC.PrepareGrid ( Nomgrille : String);
var II : Integer;
    STST,STdet : string;
    TGS : TGSaisie;
begin
  TGS := TheListGrid.findGrille(Nomgrille);
  STST := TGS.fListSaisie;
  II := 0;
  STdet := READTOKENST(STST);
  repeat
    if STDET = '' then exit;
    TraiteChamps(TGS,STDET,II);
    STdet := READTOKENST(STST);
    inc(II);
  until STdet = '';
  TGS.GS.row := 1;
  TGS.fLastRow := TGS.GS.row;
  TGS.fLastcol := TGS.GS.Col;
end;

procedure TOF_BSAISIEAVANCPOC.TraiteChamps(TGS: TGSaisie; NomChamps: string; Colonne: integer);
begin
  if TGS.GS.colCount < colonne +1 then TGS.GS.colCount := TGS.GS.ColCOunt + 1;  
  if NomChamps='IND' then
  begin
    TGS.GS.cells[Colonne,0] := ' ';
    TGS.GS.ColWidths [Colonne] := 20;
    TGS.GS.ColEditables[Colonne] := false;
  end else if NomChamps='TYPELIGNE' then
  begin
    TGS.GS.cells[Colonne,0] := ' ';
    TGS.GS.ColWidths [Colonne] := 38;
    TGS.GS.ColEditables[Colonne] := false;
    TGS.SG_TYPELIG := Colonne;
  end else if NomChamps='BM8_PHASETRA' then
  begin
    TGS.GS.cells[Colonne,0] := 'Phase';
    TGS.GS.ColWidths[colonne] := 48;
    TGS.GS.ColAligns[Colonne] := taLeftJustify;
    TGS.SG_PHASETRA := colonne;
  end else if NomChamps='BM8_CODEMARCHE' then
  begin
    TGS.GS.cells[Colonne,0] := 'Marché S/T';
    TGS.GS.ColWidths[colonne] := 70;
    TGS.GS.ColLengths[colonne] := 17;
    TGS.GS.ColAligns[Colonne] := taLeftJustify;
    TGS.SG_CODEMARCHE := colonne;
  end else if NomChamps='BM8_FAMILLENIV2' then
  begin
    TGS.GS.cells[Colonne,0] := 'Code';
    TGS.GS.ColWidths[colonne] := 75;
    TGS.GS.ColLengths[colonne] := 3;
    TGS.GS.ColAligns[Colonne] := taLeftJustify;
    TGS.SG_FAMILLENIV2 := colonne;
  end else if NomChamps='QTEPROJETE' then
  begin
    TGS.GS.cells[Colonne,0] := 'Dep. fin Chantier';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := FFQ+';'+FFQ+'; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
    TGS.SG_QTEPROJETE := colonne;
  end else if NomChamps='BM8_QTEMARCHE' then
  begin
    TGS.GS.cells[Colonne,0] := 'Quantité';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := FFQ+';'+FFQ+'; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
    TGS.SG_QTEMARCHE := colonne;
  end else if NomChamps='BM8_UNITE' then
  begin
    TGS.GS.cells[Colonne,0] := 'Unité';
    TGS.GS.ColWidths[colonne] := 5*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 5;
    TGS.GS.ColAligns[Colonne] := taCenter;
    TGS.SG_UNITE := Colonne;
  end else if NomChamps='BM8_POURCENTAVANC' then
  begin
    if TGS.fName = 'GSST' then TGS.GS.cells[Colonne,0] := '% Avct corrigé'
                          else TGS.GS.cells[Colonne,0] := '% Avct';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := '#0.00;#0.00; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
    TGS.SG_POURCENTAVANC := colonne;
  end else if NomChamps='BM8_QTEMOIS' then
  begin
    TGS.GS.cells[Colonne,0] := 'Qté Avct';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := FFQ+';'+FFQ+'; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
    TGS.SG_QTEMOIS := colonne;
  end else if NomChamps='BM8_QTECUMUL' then
  begin
    TGS.GS.cells[Colonne,0] := 'Qté Avct';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := FFQ+';'+FFQ+'; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
    TGS.SG_QTECUMUL := colonne;
  end else if NomChamps='BM8_RADQTE' then
  begin
    TGS.GS.cells[Colonne,0] := 'Qté RAD';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := FFQ+';'+FFQ+'; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
    TGS.SG_RADQTE := colonne;
  end else if NomChamps='BM8_QTECUMUL' then
  begin
    TGS.GS.cells[Colonne,0] := 'Qté Fin chantier';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := FFQ+';'+FFQ+'; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
    TGS.SG_QTECUMUL := colonne;
  end else if NomChamps='BM8_LIBELLE' then
  begin
    TGS.GS.cells[Colonne,0] := 'Désignation';
    TGS.GS.ColWidths[colonne] := 35*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 35;
    TGS.GS.ColAligns[Colonne] := taLeftJustify;
    TGS.SG_LIBELLE := colonne;
    TGS.GS.Col := Colonne;
  end else if NomChamps='BM8_RADPU' then
  begin
    TGS.GS.cells[Colonne,0] := 'RAD Pu';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := FFP+';'+FFP+'; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
    TGS.SG_RADPU := colonne;
  end else if NomChamps='BM8_RADMT' then
  begin
    TGS.GS.cells[Colonne,0] := 'RAD Mt';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := '#,##0.00;#,##0.00; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
    TGS.SG_RADMT := colonne;
  end else if NomChamps='MTAVANCCALC' then
  begin
    TGS.GS.cells[Colonne,0] := 'Mt Avanc. Calc.';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := '#,##0.00;#,##0.00; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
    TGS.SG_MTAVANCCALC := colonne;
  end else if NomChamps='SUMQTECONSO' then
  begin
    TGS.GS.cells[Colonne,0] := 'Qté Dépensées cumulées';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := FFQ+';'+FFQ+'; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
    TGS.SG_SUMQTECONSO  := colonne;
  end else if NomChamps='ECARTDEPENSE' then
  begin
    TGS.GS.cells[Colonne,0] := 'Variation M/M-1';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := '#,##0.00;#,##0.00; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
    TGS.SG_ECARTDEPENSE  := colonne;
  end else if NomChamps='PU' then
  begin
    TGS.GS.cells[Colonne,0] := 'Pu Achat';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := FFP+';'+FFP+'; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
    TGS.SG_PU  := colonne;
  end else if NomChamps='DEPENSECUMUL' then
  begin
    TGS.GS.cells[Colonne,0] := 'Dépenses Cumul.';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := '#,##0.00;#,##0.00; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
    TGS.SG_DEPENSECUMUL  := colonne;
  end else if NomChamps='SG_QTECONSOMME' then
  begin
    TGS.GS.cells[Colonne,0] := 'Qte Dépensés Cum.';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := '#,##0.00;#,##0.00; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
    TGS.SG_DEPENSEESTIME  := colonne;
  end else if NomChamps='DEPENSEESTIME' then
  begin
    TGS.GS.cells[Colonne,0] := 'Dépenses Estimées';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := '#,##0.00;#,##0.00; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
    TGS.SG_DEPENSEESTIME  := colonne;
  end else if NomChamps='POURCENTCALC' then
  begin
    TGS.GS.cells[Colonne,0] := '% Avct';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := '#0.00;#0.00; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
    TGS.SG_POURCENTCALC := colonne;
  end else if NomChamps='MTAVANCCORRIGE' then
  begin
    TGS.GS.cells[Colonne,0] := 'MT Avct corrigé';
    TGS.GS.ColWidths[colonne] := 18*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := '#,##0.00;#,##0.00; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
    TGS.SG_MTAVANCCORRIGE := colonne;
  end else if NomChamps='DATEMVT' then
  begin
    TGS.GS.cells[Colonne,0] := 'Date Mvt';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColFormats[colonne] := ShortDateFormat;
    TGS.GS.ColAligns[Colonne] := taLeftJustify;
  end else if NomChamps='CODEARTICLE' then
  begin
    TGS.GS.cells[Colonne,0] := 'Code';
    TGS.GS.ColWidths[colonne] := 18*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColAligns[Colonne] := taLeftJustify;
  end else if NomChamps='LIBELLE' then
  begin
    TGS.GS.cells[Colonne,0] := 'Libellé';
    TGS.GS.ColWidths[colonne] := 70*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColAligns[Colonne] := taLeftJustify;
  end else if NomChamps='QUANTITE' then
  begin
    TGS.GS.cells[Colonne,0] := 'Qte';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := '#,##0.00;#,##0.00; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
  end else if NomChamps='PUACH' then
  begin
    TGS.GS.cells[Colonne,0] := 'PUA';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := '#,##0.00;#,##0.00; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
  end else if NomChamps='MONTANTACH' then
  begin
    TGS.GS.cells[Colonne,0] := 'Montant achat';
    TGS.GS.ColWidths[colonne] := 12*TGS.GS.canvas.TextWidth('w');
    TGS.GS.ColLengths[colonne] := 12;
    TGS.GS.ColFormats[colonne] := '#,##0.00;#,##0.00; ;';
    TGS.GS.ColAligns[Colonne] := taRightJustify;
  end;
end;

{ TListGs }

function TListGs.Add(AObject: TGSaisie): Integer;
begin
	Result := inherited ADD(Aobject);
end;

procedure TListGs.AddGrille(TG : THgrid; NomGrille, ListeSaisie: String; ZeTOB : TOB);
var TT : TGSaisie;
begin
  TT := TGSaisie.Create;
  TT.fName := NomGrille;
  TT.fGS := TG;
  TT.fListSaisie := ListeSaisie;
  TT.fTOB := ZeTOB;
  add(TT);
end;

procedure TListGs.clear;
var indice : integer;
begin
  if count > 0 then
  begin
    for Indice := count -1 downto 0 do
    begin
      if TGSaisie(Items [Indice])<> nil then
      begin
         TGSaisie (Items [Indice]).free;
         Items[Indice] := nil;
      end;
    end;
  end;
  Pack;
  inherited;
end;

constructor TListGs.create;
begin

end;

destructor TListGs.destroy;
begin
	clear;
  inherited;
end;

function TListGs.findGrille(NomGrille: string): TGSaisie;
var Indice : integer;
begin
  result := nil;
  for Indice := 0 to Count -1 do
  begin
    if Items[Indice].fName = NomGrille then
    begin
      result:=Items[Indice];
      break;
    end;
  end;
end;

function TListGs.GetItems(Indice: integer): TGSaisie;
begin
  result := TGSaisie (Inherited Items[Indice]);
end;

procedure TListGs.SetItems(Indice: integer; const Value: TGSaisie);
begin
  Inherited Items[Indice]:= Value;
end;

procedure TOF_BSAISIEAVANCPOC.SetEventGrid(statut: Boolean);
begin
  if statut then
  begin
    GSPP.PostDrawCell := PostDrawcellPP;
    GSPP.GetCellcanvas := GetCellcanvasPP;
    GSPP.OnRowEnter := GSRowEnterPP;
    GSPP.OnRowExit := GSRowExitPP;
    GSPP.OnCellEnter := GSCellEnterPP;
    GSPP.OnCellExit := GSCellExitPP;
    GSPP.Onkeydown := GridKeyDownPP;
    //
    GSMOI.PostDrawCell := PostDrawcellMOI;
    GSMOI.GetCellcanvas := GetCellcanvasMOI;
    GSMOI.OnRowEnter := GSRowEnterMOI;
    GSMOI.OnRowExit := GSRowExitMOI;
    GSMOI.OnCellEnter := GSCellEnterMOI;
    GSMOI.OnCellExit := GSCellExitMOI;
    GSMOI.Onkeydown := GridKeyDownMOI;
    //
    GSFI.PostDrawCell := PostDrawcellFI;
    GSFI.GetCellcanvas := GetCellcanvasFI;
    GSFI.OnRowEnter := GSRowEnterFI;
    GSFI.OnRowExit := GSRowExitFI;
    GSFI.OnCellEnter := GSCellEnterFI;
    GSFI.OnCellExit := GSCellExitFI;
    GSFI.Onkeydown := GridKeyDownFI;
    //
    GSST.PostDrawCell := PostDrawcellST;
    GSST.GetCellcanvas := GetCellcanvasST;
    GSST.OnRowEnter := GSRowEnterST;
    GSST.OnCellEnter := GSCellEnterST;
    GSST.OnCellExit := GSCellExitST;
    GSST.OnRowExit := GSRowExitST;
    GSST.OnElipsisClick := GSSTElipsisClick;
    GSST.Onkeydown := GridKeyDownST;
  end else
  begin
    GSPP.PostDrawCell := nil;
  end;
end;
function TOF_BSAISIEAVANCPOC.RecupTypeGraph (TOBExamin : Tob) : Integer;
var Niveau : integer;
    TypeLigne : String;
begin
  TypeLigne := TOBExamin.getValue('TYPELIGNE');
  Niveau := TOBExamin.getValue('NIVEAU');

  if TypeLigne = 'ART' Then Result := 3
  else if (TypeLigne = 'OUV') then Result := 0
  else if (TypeLigne = 'COM') then Result := 5
  else if (copy(TypeLigne,1,2)='DP') and (Niveau=1) then Result := 1
  else if (copy(TypeLigne,1,2)='DP') and (Niveau>1) then Result := 2
  else Result := -1;
end;

procedure TOF_BSAISIEAVANCPOC.ChangePage(Sender: Tobject);
var TheOnglet : string;
begin
  TheOnglet := GetActiveTabSheet ('PG1').Name;
  if TheOnglet = 'TSPRODPOC' then
  begin
    ZeCurSaisie := TheListGrid.findGrille('GSPP');
    EnterGrid ('GSPP');
  end else if TheOnglet = 'TSMOIND' then
  begin
    ZeCurSaisie := TheListGrid.findGrille('GSMOI');
    EnterGrid ('GSMOI');
  end else if TheOnglet = 'TSFRAIS' then
  begin
    ZeCurSaisie := TheListGrid.findGrille('GSFI');
    EnterGrid ('GSFI');
  end else if TheOnglet = 'TSST' then
  begin
    ZeCurSaisie := TheListGrid.findGrille('GSST');
    EnterGrid ('GSST');
  end;
  ZeCurSaisie.GS.Invalidate;
end;

procedure TOF_BSAISIEAVANCPOC.BeforeChangePage(Sender: TObject;var AllowChanging: Boolean);
var TheOnglet : string;
begin
  TheOnglet := GetActiveTabSheet ('PG1').name;
end;

procedure TOF_BSAISIEAVANCPOC.ReajusteRegion (Arect : Trect);
begin
  Arect.Top := Arect.Top +1;
  Arect.Left := Arect.left +1;
  Arect.Right := Arect.Right -1;
  Arect.Bottom := Arect.Bottom -1;
end;


procedure TOF_BSAISIEAVANCPOC.PostDrawcellPP(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var Arect : Trect;
    TOBL : TOB;
    NumGraph : Integer;
    TheCur : TGSaisie;
    decal : integer;
    NbCars : integer;
    TheText : string;
begin
  TheCur := TheListGrid.findGrille('GSPP');
  if ARow <= 0 then exit;
  if Acol <= 0 then exit;
  if csDestroying in GSPP.ComponentState then Exit;
  if GSPP.RowHeights[ARow] <= 0 then Exit;
  if TOBPP.detail.count = 0 then exit;
  TOBL := TOBPP.detail[ARow-1];
  if TOBL = nil then exit;

  ARect := GSPP.CellRect(ACol, ARow);
  ReajusteRegion (Arect);
  GSPP.Canvas.Pen.Style := psSolid;
  GSPP.Canvas.Pen.Color := clgray;
  GSPP.Canvas.Brush.Style := BsSolid;
  if ((TheCur.SG_TYPELIG > -1) and (Acol = TheCur.SG_TYPELIG) and (Arow >= TheCur.GS.fixedRows)) then
  begin
    GSPP.Canvas.FillRect(ARect);
    NumGraph := RecupTypeGraph(TOBL);
    if NumGraph >= 0 then
    begin
      ImList.DrawingStyle := dsTransparent;
      if (TOBL.GetInteger('NIVEAU')>1) then
      begin
      	ImList.Draw(CanVas, ARect.left+16, ARect.top, NumGraph);
      end else
      begin
      	ImList.Draw(CanVas, ARect.left+1, ARect.top, NumGraph);
      end;
    end;
  end;
  if ((TheCur.SG_LIBELLE > -1) and (Acol = TheCur.SG_LIBELLE) and (Arow >= TheCur.GS.fixedRows)) then
  begin
    Nbcars := TOBL.GetInteger('NIVEAU');
    decal := canvas.TextWidth('W') * (Nbcars-1);
    TheText := copy(TOBL.getvalue('BM8_LIBELLE'),1,35-nbCars);
    Canvas.FillRect(ARect);
    GSPP.Canvas.Brush.Style := bsSolid;
    GSPP.Canvas.TextOut (Arect.left+Decal+1,Arect.Top +2 ,Thetext);
  end;

  if (not VisiblePP(TheCur,TOBL,Acol) and (Arow >= TheCur.GS.fixedRows)) then
  begin
    Canvas.FillRect(ARect);
  end;
end;

procedure TOF_BSAISIEAVANCPOC.PostDrawcellFI(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var Arect : Trect;
    TOBL : TOB;
    NumGraph : Integer;
    TheCur : TGSaisie;
    decal : integer;
    TheText : string;
begin
  if TOBPP.detail.count = 0 then exit;
  TheCur := TheListGrid.findGrille('GSFI');
  if ARow <= 0 then exit;
  if Acol <= 0 then exit;
  if csDestroying in GSFI.ComponentState then Exit;
  if GSFI.RowHeights[ARow] <= 0 then Exit;
  TOBL := TOBFI.detail[ARow-1];
  if TOBL = nil then exit;

  ARect := GSFI.CellRect(ACol, ARow);
  ReajusteRegion (Arect);
  GSFI.Canvas.Pen.Style := psSolid;
  GSFI.Canvas.Pen.Color := clgray;
  GSFI.Canvas.Brush.Style := BsSolid;
  if ((TheCur.SG_TYPELIG > -1) and (Acol = TheCur.SG_TYPELIG) and (Arow >= TheCur.GS.fixedRows)) then
  begin
    GSFI.Canvas.FillRect(ARect);
    NumGraph := RecupTypeGraph(TOBL);
    if NumGraph >= 0 then
    begin
      ImList.DrawingStyle := dsTransparent;
      if (TOBL.GetInteger('NIVEAU')>1) then
      begin
      	ImList.Draw(CanVas, ARect.left+16, ARect.top, NumGraph);
      end else
      begin
      	ImList.Draw(CanVas, ARect.left+1, ARect.top, NumGraph);
      end;
    end;
  end;
  if ((TheCur.SG_LIBELLE > -1) and (Acol = TheCur.SG_LIBELLE) and (Arow >= TheCur.GS.fixedRows)) then
  begin
    decal := canvas.TextWidth('W') * (TOBL.GetInteger('NIVEAU')-1);
    TheText := TOBL.getvalue('BM8_LIBELLE');
    Canvas.FillRect(ARect);
    GSFI.Canvas.Brush.Style := bsSolid;
    GSFI.Canvas.TextOut (Arect.left+Decal+1,Arect.Top +2 ,Thetext);
  end;

  if (not VisibleFI(TheCur,TOBL,Acol) and (Arow >= TheCur.GS.fixedRows)) then
  begin
    Canvas.FillRect(ARect);
  end;
end;

procedure TOF_BSAISIEAVANCPOC.PostDrawcellMOI(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var Arect : Trect;
    TOBL : TOB;
    NumGraph : Integer;
    TheCur : TGSaisie;
    decal : integer;
    TheText : string;
begin
  TheCur := TheListGrid.findGrille('GSMOI');
  if TOBPP.detail.count = 0 then exit;
  if ARow <= 0 then exit;
  if Acol <= 0 then exit;
  if csDestroying in GSMOI.ComponentState then Exit;
  if GSMOI.RowHeights[ARow] <= 0 then Exit;
  TOBL := TOBMOI.detail[ARow-1];
  if TOBL = nil then exit;

  ARect := GSMOI.CellRect(ACol, ARow);
  ReajusteRegion (Arect);
  GSMOI.Canvas.Pen.Style := psSolid;
  GSMOI.Canvas.Pen.Color := clgray;
  GSMOI.Canvas.Brush.Style := BsSolid;
  if ((TheCur.SG_TYPELIG > -1) and (Acol = TheCur.SG_TYPELIG) and (Arow >= TheCur.GS.fixedRows)) then
  begin
    GSMOI.Canvas.FillRect(ARect);
    NumGraph := RecupTypeGraph(TOBL);
    if NumGraph >= 0 then
    begin
      ImList.DrawingStyle := dsTransparent;
      if (TOBL.GetInteger('NIVEAU')>1) then
      begin
      	ImList.Draw(CanVas, ARect.left+16, ARect.top, NumGraph);
      end else
      begin
      	ImList.Draw(CanVas, ARect.left+1, ARect.top, NumGraph);
      end;
    end;
  end;
  if ((TheCur.SG_LIBELLE > -1) and (Acol = TheCur.SG_LIBELLE) and (Arow >= TheCur.GS.fixedRows)) then
  begin
    decal := canvas.TextWidth('W') * (TOBL.GetInteger('NIVEAU')-1);
    TheText := TOBL.getvalue('BM8_LIBELLE');
    Canvas.FillRect(ARect);
    GSMOI.Canvas.Brush.Style := bsSolid;
    GSMOI.Canvas.TextOut (Arect.left+Decal+1,Arect.Top +2 ,Thetext);
  end;

  if (not VisibleMOI(TheCur,TOBL,Acol) and (Arow >= TheCur.GS.fixedRows)) then
  begin
    Canvas.FillRect(ARect);
  end;
end;

procedure TOF_BSAISIEAVANCPOC.PostDrawcellST(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var Arect : Trect;
    TOBL : TOB;
    NumGraph : Integer;
    TheCur : TGSaisie;
    decal : integer;
    TheText : string;
begin
  TheCur := TheListGrid.findGrille('GSST');
  if TOBPP.detail.count = 0 then exit;
  if ARow <= 0 then exit;
  if Acol <= 0 then exit;
  if csDestroying in GSST.ComponentState then Exit;
  if GSST.RowHeights[ARow] <= 0 then Exit;
  TOBL := TOBST.detail[ARow-1];
  if TOBL = nil then exit;

  ARect := GSST.CellRect(ACol, ARow);
  ReajusteRegion (Arect);
  GSST.Canvas.Pen.Style := psSolid;
  GSST.Canvas.Pen.Color := clgray;
  GSST.Canvas.Brush.Style := BsSolid;
  if ((TheCur.SG_TYPELIG > -1) and (Acol = TheCur.SG_TYPELIG) and (Arow >= TheCur.GS.fixedRows)) then
  begin
    GSST.Canvas.FillRect(ARect);
    NumGraph := RecupTypeGraph(TOBL);
    if NumGraph >= 0 then
    begin
      ImList.DrawingStyle := dsTransparent;
      if (TOBL.GetInteger('NIVEAU')>1) then
      begin
      	ImList.Draw(CanVas, ARect.left+16, ARect.top, NumGraph);
      end else
      begin
      	ImList.Draw(CanVas, ARect.left+1, ARect.top, NumGraph);
      end;
    end;
  end;
  if ((TheCur.SG_LIBELLE > -1) and (Acol = TheCur.SG_LIBELLE) and (Arow >= TheCur.GS.fixedRows)) then
  begin
    decal := canvas.TextWidth('W') * (TOBL.GetInteger('NIVEAU')-1);
    TheText := TOBL.getvalue('BM8_LIBELLE');
    Canvas.FillRect(ARect);
    GSST.Canvas.Brush.Style := bsSolid;
    GSST.Canvas.TextOut (Arect.left+Decal+1,Arect.Top +2 ,Thetext);
  end;

  if (not VisibleST(TheCur,TOBL,Acol) and (Arow >= TheCur.GS.fixedRows)) then
  begin
    Canvas.FillRect(ARect);
  end;
end;

procedure TOF_BSAISIEAVANCPOC.GetCellCanvasFI(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var TOBL : TOB;
    TypeLigne : string;
    Niveau : integer;
    Height : Integer;
begin
  if ACol < GSFI.FixedCols then Exit;
  if Arow < GSFI.Fixedrows then Exit;
  if TOBFI.detail.count = 0 then exit;
  TOBL := TOBFI.detail[Arow-1];
  if TOBL = nil then Exit;
  TypeLigne := TOBL.getString('TYPELIGNE');
  Niveau := TOBL.GetInteger('NIVEAU');

  if (TypeLigne = 'COM') then
  begin
    AppliqueStyleCommentaire(Canvas,fGestionAff);
  end else if (TypeLigne = 'ART') then
  begin
    AppliqueStyleArticle(Canvas,fGestionAff);
  end else if (TypeLigne = 'OUV') then
  begin
    AppliqueStyleOuvrage(Canvas,fGestionAff);
  end else if (TypeLigne = 'COM') then
  begin
  end else if ((copy(TypeLigne,1,2)='DP') or (copy(TypeLigne,1,2)='TP')) and (Niveau=1) then
  begin
    AppliqueStyleParag(Canvas,fGestionAff,1);
  end else if ((copy(TypeLigne,1,2)='DP') or (copy(TypeLigne,1,2)='TP')) and (Niveau>1) then
  begin
    AppliqueStyleParag(Canvas,fGestionAff,2);
  end else exit;
  
  Height := GSFI.Canvas.TextHeight ('W') + 2;
  if Height < 18 then height := 18;
  GSFI.RowHeights [Arow] := height;
end;

procedure TOF_BSAISIEAVANCPOC.GetCellCanvasMOI(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var TOBL : TOB;
    TypeLigne : string;
    Niveau : integer;
    Height : Integer;
begin
  if ACol < GSMOI.FixedCols then Exit;
  if Arow < GSMOI.Fixedrows then Exit;
  if TOBMOI.detail.count = 0 then exit;
  TOBL := TOBMOI.detail[Arow-1];
  if TOBL = nil then Exit;
  TypeLigne := TOBL.getString('TYPELIGNE');
  Niveau := TOBL.GetInteger('NIVEAU');

  if (TypeLigne = 'COM') then
  begin
    AppliqueStyleCommentaire(Canvas,fGestionAff);
  end else if (TypeLigne = 'ART') then
  begin
    AppliqueStyleArticle(Canvas,fGestionAff);
  end else if (TypeLigne = 'OUV') then
  begin
    AppliqueStyleOuvrage(Canvas,fGestionAff);
  end else if (TypeLigne = 'COM') then
  begin
  end else if ((copy(TypeLigne,1,2)='DP') or (copy(TypeLigne,1,2)='TP')) and (Niveau=1) then
  begin
    AppliqueStyleParag(Canvas,fGestionAff,1);
  end else if ((copy(TypeLigne,1,2)='DP') or (copy(TypeLigne,1,2)='TP')) and (Niveau>1) then
  begin
    AppliqueStyleParag(Canvas,fGestionAff,2);
  end else exit;
  
  Height := GSMOI.Canvas.TextHeight ('W') + 2;
  if Height < 18 then height := 18;
  GSMOI.RowHeights [Arow] := height;

end;

procedure TOF_BSAISIEAVANCPOC.GetCellCanvasPP(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var TOBL : TOB;
    TypeLigne : string;
    Niveau : integer;
    Height : Integer;
begin
  if ACol < GSPP.FixedCols then Exit;
  if Arow < GSPP.Fixedrows then Exit;
  if TOBPP.detail.count = 0 then exit;
  TOBL := TOBPP.detail[Arow-1];
  if TOBL = nil then Exit;
  TypeLigne := TOBL.getString('TYPELIGNE');
  Niveau := TOBL.GetInteger('NIVEAU');

  if (TypeLigne = 'COM') then
  begin
    AppliqueStyleCommentaire(Canvas,fGestionAff);
  end else if (TypeLigne = 'ART') then
  begin
    AppliqueStyleArticle(Canvas,fGestionAff);
  end else if (TypeLigne = 'OUV') then
  begin
    AppliqueStyleOuvrage(Canvas,fGestionAff);
  end else if (TypeLigne = 'COM') then
  begin
  end else if ((copy(TypeLigne,1,2)='DP') or (copy(TypeLigne,1,2)='TP')) and (Niveau=1) then
  begin
    AppliqueStyleParag(Canvas,fGestionAff,1);
  end else if ((copy(TypeLigne,1,2)='DP') or (copy(TypeLigne,1,2)='TP')) and (Niveau>1) then
  begin
    AppliqueStyleParag(Canvas,fGestionAff,2);
  end else exit;
  
  Height := GSPP.Canvas.TextHeight ('W') + 2;
  if Height < 18 then height := 18;
  GSPP.RowHeights [Arow] := height;

end;

procedure TOF_BSAISIEAVANCPOC.GetCellCanvasST(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var TOBL : TOB;
    TypeLigne : string;
    Niveau : integer;
    Height : Integer;
begin
  if ACol < GSST.FixedCols then Exit;
  if Arow < GSST.Fixedrows then Exit;
  if TOBST.detail.count = 0 then exit;
  TOBL := TOBST.detail[Arow-1];
  if TOBL = nil then Exit;
  TypeLigne := TOBL.getString('TYPELIGNE');
  Niveau := TOBL.GetInteger('NIVEAU');

  if (TypeLigne = 'COM') then
  begin
    AppliqueStyleCommentaire(Canvas,fGestionAff);
  end else if (TypeLigne = 'ART') then
  begin
    AppliqueStyleArticle(Canvas,fGestionAff);
  end else if (TypeLigne = 'OUV') then
  begin
    AppliqueStyleOuvrage(Canvas,fGestionAff);
  end else if (TypeLigne = 'COM') then
  begin
  end else if ((copy(TypeLigne,1,2)='DP') or (copy(TypeLigne,1,2)='TP')) and (Niveau=1) then
  begin
    AppliqueStyleParag(Canvas,fGestionAff,1);
  end else if ((copy(TypeLigne,1,2)='DP') or (copy(TypeLigne,1,2)='TP')) and (Niveau>1) then
  begin
    AppliqueStyleParag(Canvas,fGestionAff,2);
  end else exit;
  
  Height := GSST.Canvas.TextHeight ('W') + 2;
  if Height < 18 then height := 18;
  GSST.RowHeights [Arow] := height;

end;

{ TGSaisie }

constructor TGSaisie.create;
begin
  fLastCol := 1;
  fLastRow := 1;
  SG_TYPELIG := -1;
  SG_LIBELLE := -1;
  SG_PHASETRA := -1;
  SG_FAMILLENIV2 := -1;
  SG_QTEMARCHE := -1;
  SG_UNITE := -1;
  SG_POURCENTAVANC := -1;
  SG_QTEMOIS := -1;
  SG_QTECUMUL := -1;
  SG_RADQTE := -1;
  SG_RADPU := -1;
  SG_RADMT := -1;
  SG_QTEPROJETE := -1;
  SG_MTAVANCCALC := -1;
  SG_SUMQTECONSO := -1;
  SG_ECARTDEPENSE := -1;
  SG_DEPENSECUMUL := -1;
  SG_DEPENSEESTIME := -1;
  SG_POURCENTCALC := -1;
  SG_CODEMARCHE := -1;
  SG_MTAVANCCORRIGE := -1;
  SG_PU := -1;
end;


procedure TOF_BSAISIEAVANCPOC.GridKeyDownPP (Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    13 : if (Shift = []) then
    begin
    	Key := 0;
      SendMessage(THedit(Sender).Handle, WM_KeyDown, VK_TAB, 0);
    end;
  end;
end;

function TOF_BSAISIEAVANCPOC.VisiblePP (TheCur : TGSaisie;TOBL : TOB;  Acol : integer) : Boolean;
var TypeLigne : string;
begin
  Result := true;
  TypeLigne := TOBL.getValue('TYPELIGNE');
  if (TypeLigne = 'ART')  then
  begin
    if (Acol = TheCur.SG_PHASETRA) then
    begin
      Result := false;
    end;
    exit;
  end;
  if (TypeLigne = 'COM')  then
  begin
    if (Acol <> TheCur.SG_LIBELLE) and
       (Acol <> TheCur.SG_TYPELIG) then
    begin
      Result := false;
    end;
    Exit;
  end;
  //
  if ((Acol = TheCur.SG_QTEMARCHE) or
     (Acol = TheCur.SG_POURCENTAVANC) or
     (Acol = TheCur.SG_QTEMOIS) or
     (Acol = TheCur.SG_QTECUMUL) or
     (Acol = TheCur.SG_RADQTE) or
     (Acol = TheCur.SG_SUMQTECONSO) or
     (Acol = TheCur.SG_QTEPROJETE)) and
     (TypeLigne <> 'ART') then
  begin
    Result := false;
    exit;
  end;
end;

function TOF_BSAISIEAVANCPOC.ZoneAccessiblePP(ACol, ARow: Longint): boolean; { NEWPIECE }
var TOBL: TOB;
    TheGS : TGSaisie;
begin
  Result := True;
  TheGS := TheListGrid.findGrille('GSPP');
  if TOBPP.detail.count = 0 then exit;
  if Arow > TOBPP.detail.count then begin result := false; exit; end;
  TOBL := TOBPP.detail[Arow-1];
  if TOBL = nil then Exit;
  if (Acol = 0) then BEGIN result := false; Exit; END;
  if (Acol = TheGS.SG_TYPELIG) or (ACol = TheGS.SG_PHASETRA) or
     (ACol = TheGS.SG_FAMILLENIV2) or (Acol = TheGS.SG_UNITE) or
     (Acol = TheGS.SG_SUMQTECONSO ) or
     (Acol = TheGS.SG_QTECUMUL) or (Acol = TheGS.SG_QTEPROJETE) or (Acol = TheGS.SG_QTEMARCHE) then BEGIN  Result := false; Exit; END;
  if not VisiblePP(TheGS,TOBL,Acol) then BEGIN result := false; Exit; end;
end;


procedure TOF_BSAISIEAVANCPOC.ZoneSuivanteOuOkPP(var ACol, ARow: Longint; var Cancel: boolean);
var Sens, ii, Lim: integer;
  OldEna, ChgLig, ChgSens: boolean;
  RowFirst : integer;
begin
	RowFirst := ARow;
  OldEna := GSPP.SynEnabled;
  GSPP.SynEnabled := False;
  Sens := -1;
  ChgLig := (GSPP.Row <> ARow);
  ChgSens := False;
  if GSPP.Row > ARow then Sens := 1 else if ((GSPP.Row = ARow) and (ACol <= GSPP.Col)) then Sens := 1;
  ACol := GSPP.Col;
  ARow := GSPP.Row;
  ii := 0;
  while not ZoneAccessiblePP(ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if Sens = 1 then
    begin
      Lim := GSPP.RowCount;
      // ---
      if ((ACol = GSPP.ColCount - 1) and (ARow >= Lim)) then
      begin
        if (ChgSens) then Break;
      end;
      if ChgLig then
      begin
        ACol := GSPP.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < GSPP.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        if ARow >= Lim then
        begin
          ARow := 1;
          Acol := 1;
        end else
        begin
          ACol := GSPP.FixedCols;
        end;
      end;
    end else
    begin
      if ((ACol = GSPP.FixedCols) and (ARow = 1)) then
      begin
        if ChgSens then Break else
        begin
          Arow := GSPP.rowCount;
          ChgLig := true;
        end;
      end;
      if ChgLig then
      begin
        ACol := GSPP.ColCount;
        ChgLig := False;
      end;
      if ACol > GSPP.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := GSPP.ColCount - 1;
      end;
    end;
  end;
  GSPP.SynEnabled := OldEna;
  if (Arow <> RowFirst) then
  BEGIN
  	GSRowExitPP (self,RowFirst,cancel,false);
  	if Cancel then
    begin
    	GSRowEnterPP (self,RowFirst,cancel,false);
    end else
    begin
    	GSRowEnterPP (self,Arow,cancel,false);
    end;
  END;
end;

procedure TOF_BSAISIEAVANCPOC.GSCellEnterFI(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var CurSaisie : TGSaisie;
begin
  if ActionResult = taConsult then Exit;
  CurSaisie := TheListGrid.findGrille('GSFI');
  ZoneSuivanteOuOkFI(ACol, ARow, Cancel);
  if not Cancel then
  begin
    CurSaisie.StCellCur := GSFI.Cells[Acol,Arow];
  end;
end;

procedure TOF_BSAISIEAVANCPOC.GSCellEnterMOI(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var CurSaisie : TGSaisie;
begin
  if ActionResult = taConsult then Exit;
  CurSaisie := TheListGrid.findGrille('GSMOI');
  ZoneSuivanteOuOkMOI(ACol, ARow, Cancel);
  if not Cancel then
  begin
    CurSaisie.StCellCur := GSPP.Cells[Acol,Arow];
  end;
end;

procedure TOF_BSAISIEAVANCPOC.GSCellEnterPP(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var CurSaisie : TGSaisie;
begin
  if ActionResult = taConsult then Exit;
  CurSaisie := TheListGrid.findGrille('GSPP');
  ZoneSuivanteOuOkPP(ACol, ARow, Cancel);
  if not Cancel then
  begin
    CurSaisie.StCellCur := GSPP.Cells[Acol,Arow];
  end;
end;

procedure TOF_BSAISIEAVANCPOC.GSCellEnterST(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var CurSaisie : TGSaisie;
begin
  if ActionResult = taConsult then Exit;
  CurSaisie := TheListGrid.findGrille('GSST');
  ZoneSuivanteOuOkST(ACol, ARow, Cancel);
  if not Cancel then
  begin
    GSST.ElipsisButton := (ACol=CurSaisie.SG_DEPENSEESTIME);
    CurSaisie.StCellCur := GSPP.Cells[Acol,Arow];
  end;
end;

procedure TOF_BSAISIEAVANCPOC.GSCellExitFI(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var TOBL : TOB;
    TGS : TGSaisie;
begin
  if ActionResult = taConsult then Exit;
  if TOBFI.detail.count = 0 then exit;
  TOBL := TOBFI.detail[Arow-1];
  TGS := TheListGrid.findGrille('GSFI');
  if TGS.fStCellCur <> GSFI.Cells[ACol,Arow] then
  begin
    if Acol = TGS.SG_LIBELLE then
    begin
      GSFI.Cells[ACol,Arow] := TGS.fStCellCur;
    end else if Acol = TGS.SG_RADQTE then
    begin
      TOBL.SetDouble('BM8_RADQTE',Valeur(GSFI.Cells [Acol,Arow]));
      CalculeMtRAD (TOBL);
      DataModif := true;
    end else if Acol = TGS.SG_RADPU then
    begin
      TOBL.SetDouble('BM8_RADPU',Valeur(GSFI.Cells [Acol,Arow]));
      CalculeMtRAD (TOBL);
      DataModif := true;
    end else if Acol = TGS.SG_RADMT then
    begin
      TOBL.SetDouble('BM8_RADMT',Valeur(GSFI.Cells [Acol,Arow]));
      CalculePURAD (TOBL);
      DataModif := true;
    end;
    TOBL.SetDouble('MTAVANCCALC',TOBL.GetDouble('BM8_MTMARCHE')-TOBL.GetDouble('BM8_RADMT'));
    //
    TOBL.SetDouble('ECARTDEPENSE',TOBL.GetDouble('BM8_MTCONSOMOIS'));
    //
  end;
  TOBL.PutLigneGrid(GSFI,ARow,false,false,stLigFI);
  GSFI.invalidateRow(Arow);
end;

procedure TOF_BSAISIEAVANCPOC.GSCellExitMOI(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var TOBL : TOB;
    TGS : TGSaisie;
begin
  if ActionResult = taConsult then Exit;
  if TOBMOI.detail.count = 0 then exit;
  TOBL := TOBMOI.detail[Arow-1];
  TGS := TheListGrid.findGrille('GSMOI');
  if TGS.fStCellCur <> GSMOI.Cells[ACol,Arow] then
  begin
    if Acol = TGS.SG_LIBELLE then
    begin
      GSMOI.Cells[ACol,Arow] := TGS.fStCellCur;
    end else if Acol = TGS.SG_RADQTE then
    begin
      TOBL.SetDouble('BM8_RADQTE',Valeur(GSMOI.Cells [Acol,Arow]));
      CalculeMtAvt (TOBL);
      CalculePourcentAvt (TOBL);
      CalculeMtRAD (TOBL);
      DataModif := true;
    end;
  end;
  TOBL.PutLigneGrid(GSMOI,ARow,false,false,stLigMOI);
  GSMOI.invalidateRow(Arow);
end;

procedure TOF_BSAISIEAVANCPOC.GSCellExitPP(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var TOBL : TOB;
    TGS : TGSaisie;
begin
  if ActionResult = taConsult then Exit;
  if TOBPP.detail.count = 0 then exit;
  TOBL := TOBPP.detail[Arow-1];
  TGS := TheListGrid.findGrille('GSPP');
  if TGS.fStCellCur <> GSPP.Cells[ACol,Arow] then
  begin
    if Acol = TGS.SG_LIBELLE then
    begin
      GSPP.Cells[ACol,Arow] := TGS.fStCellCur;
    end else if Acol = TGS.SG_POURCENTAVANC then
    begin
      TOBL.SetDouble('BM8_POURCENTAVANC', Valeur(GSPP.Cells [Acol,Arow]));
      CalculeQteAvt(TOBL);
      CalculeMtAvt(TOBL);
      DataModif := true;
    end else if Acol = TGS.SG_RADQTE then
    begin
      TOBL.SetDouble('BM8_RADQTE',Valeur(GSPP.Cells [Acol,Arow]));
      CalculeMtRAD (TOBL);
      CalculeProjete (TOBL);
      DataModif := true;
    end;
  end;
  TOBL.PutLigneGrid(GSPP,ARow,false,false,stLigPP);
  GSPP.invalidateRow(Arow);
end;

procedure TOF_BSAISIEAVANCPOC.GSCellExitST(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var TOBL : TOB;
    TGS : TGSaisie;
begin
  if ActionResult = taConsult then Exit;
  if TOBST.detail.count = 0 then exit;
  TOBL := TOBST.detail[Arow-1];
  TGS := TheListGrid.findGrille('GSST');
  if TGS.fStCellCur <> GSST.Cells[ACol,Arow] then
  begin
    if Acol = TGS.SG_LIBELLE then
    begin
      GSST.Cells[ACol,Arow] := TGS.fStCellCur;
    end else if Acol = TGS.SG_POURCENTAVANC then
    begin
      if (Valeur(GSST.Cells [Acol,Arow]) > TOBL.GetDouble('POURCENTCALC')) then
      begin
        PGIInfo('Vous ne pouvez pas dépasser le % calculé');
        cancel := True;
        Exit;
      end;
      TOBL.SetDouble('BM8_POURCENTAVANC', Valeur(GSST.Cells [Acol,Arow]));
      CalculeAvancCorrigeST (TOBL);
      DataModif := true;
    end else if Acol = TGS.SG_DEPENSEESTIME then
    begin
      GSST.Cells[ACol,Arow] := TGS.fStCellCur;
    end;
  end;
  TOBL.PutLigneGrid(GSST,ARow,false,false,stLigST);
  GSST.invalidateRow(Arow);
end;

procedure TOF_BSAISIEAVANCPOC.GSRowEnterFI(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
end;

procedure TOF_BSAISIEAVANCPOC.GSRowEnterMOI(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
end;

procedure TOF_BSAISIEAVANCPOC.GSRowEnterPP(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
end;

procedure TOF_BSAISIEAVANCPOC.GSRowEnterST(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin

end;

procedure TOF_BSAISIEAVANCPOC.GSRowExitFI(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin

end;

procedure TOF_BSAISIEAVANCPOC.GSRowExitMOI(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin

end;

procedure TOF_BSAISIEAVANCPOC.GSRowExitPP(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
  GSPP.InvalidateRow(Ou);
end;

procedure TOF_BSAISIEAVANCPOC.GSRowExitST(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin

end;

procedure TOF_BSAISIEAVANCPOC.EnterGrid(NomGrid: string);
var CurGS : TGSaisie;
    cancel :Boolean;
begin
  if NomGrid = 'GSPP' then
  begin
    CurGS := TheListGrid.findGrille(NomGrid);
    if TOBPP.detail.count = 0 then BEGIN GSPP.enabled := false; exit; END;
    GSRowEnterPP(Self,CurGS.fLastRow,cancel,false);
    GSCellEnterPP(Self,CurGS.fLastcol,CurGS.fLastRow,cancel);
    GSPP.SetFocus;
    GSPP.ShowEditor;
  end else if NomGrid = 'GSMOI' then
  begin
    CurGS := TheListGrid.findGrille(NomGrid);
    if TOBMOI.detail.count = 0 then BEGIN GSMOI.enabled := false; exit; END;
    GSRowEnterMOI(Self,CurGS.fLastRow,cancel,false);
    GSCellEnterMOI(Self,CurGS.fLastcol,CurGS.fLastRow,cancel);
    GSMOI.SetFocus;
    GSMOI.ShowEditor;
  end else if NomGrid = 'GSFI' then
  begin
    CurGS := TheListGrid.findGrille(NomGrid);
    if TOBFI.detail.count = 0 then BEGIN GSFI.enabled := false; Exit; END;
    GSRowEnterFI(Self,CurGS.fLastRow,cancel,false);
    GSCellEnterFI(Self,CurGS.fLastcol,CurGS.fLastRow,cancel);
    GSFI.SetFocus;
    GSFI.ShowEditor;
  end else if NomGrid = 'GSST' then
  begin
    CurGS := TheListGrid.findGrille(NomGrid);
    if TOBST.detail.count = 0 then BEGIN GSST.enabled := false; EXIT; END;
    GSRowEnterST(Self,CurGS.fLastRow,cancel,false);
    GSCellEnterST(Self,CurGS.fLastcol,CurGS.fLastRow,cancel);
    GSST.SetFocus;
    GSST.ShowEditor;
  end
end;

function TOF_BSAISIEAVANCPOC.VisibleMOI(TheCur: TGSaisie; TOBL: TOB;Acol: integer): Boolean;
var TypeLigne : string;
begin
  Result := true;
  TypeLigne := TOBL.getValue('TYPELIGNE');
  if (TypeLigne = 'ART')  then
  begin
    if (Acol = TheCur.SG_PHASETRA) then
    begin
      Result := false;
    end;
    exit;
  end;
  if (TypeLigne = 'COM')  then
  begin
    if (Acol <> TheCur.SG_LIBELLE) and
       (Acol <> TheCur.SG_TYPELIG) then
    begin
      Result := false;
    end;
    Exit;
  end;
  //
  if ((Acol = TheCur.SG_QTEMARCHE) or
     (Acol = TheCur.SG_UNITE) or
     (Acol = TheCur.SG_POURCENTAVANC) or
     (Acol = TheCur.SG_QTEMOIS) or
     (Acol = TheCur.SG_FAMILLENIV2) or
     (Acol = TheCur.SG_QTECUMUL) or
     (Acol = TheCur.SG_RADQTE)) then
  begin
    Result := false;
    exit;
  end;
end;

function TOF_BSAISIEAVANCPOC.ZoneAccessibleMOI(ACol,ARow: Integer): boolean;
var TOBL: TOB;
    TheGS : TGSaisie;
begin
  Result := True;
  if TOBMOI.detail.count = 0 then exit;
  TheGS := TheListGrid.findGrille('GSMOI');
  if Arow > TOBMOI.detail.count then begin result := false; exit; end;
  TOBL := TOBMOI.detail[Arow-1];
  if TOBL = nil then Exit;
  if (Acol = 0) then BEGIN result := false; Exit; END;
  if (Acol = TheGS.SG_TYPELIG) or (ACol = TheGS.SG_PHASETRA) or
     (ACol = TheGS.SG_FAMILLENIV2) or (Acol = TheGS.SG_UNITE) or
     (Acol = TheGS.SG_SUMQTECONSO) or (Acol = TheGS.SG_QTEMARCHE) or (Acol = TheGS.SG_POURCENTAVANC) then BEGIN  Result := false; Exit; END;
  if not VisibleMOI(TheGS,TOBL,Acol) then BEGIN result := false; Exit; end;
end;

procedure TOF_BSAISIEAVANCPOC.ZoneSuivanteOuOkMOI(var ACol, ARow: Integer;var Cancel: boolean);
var Sens, ii, Lim: integer;
  OldEna, ChgLig, ChgSens: boolean;
  RowFirst : integer;
begin
	RowFirst := ARow;
  OldEna := GSMOI.SynEnabled;
  GSMOI.SynEnabled := False;
  Sens := -1;
  ChgLig := (GSMOI.Row <> ARow);
  ChgSens := False;
  if GSMOI.Row > ARow then Sens := 1 else if ((GSMOI.Row = ARow) and (ACol <= GSMOI.Col)) then Sens := 1;
  ACol := GSMOI.Col;
  ARow := GSMOI.Row;
  ii := 0;
  while not ZoneAccessibleMOI(ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if Sens = 1 then
    begin
      Lim := GSMOI.RowCount;
      // ---
      if ((ACol = GSMOI.ColCount - 1) and (ARow >= Lim)) then
      begin
        if (ChgSens) then Break;
      end;
      if ChgLig then
      begin
        ACol := GSMOI.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < GSMOI.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        if ARow >= Lim then
        begin
          ARow := 1;
          Acol := 1;
        end else
        begin
          ACol := GSMOI.FixedCols;
        end;
      end;
    end else
    begin
      if ((ACol = GSMOI.FixedCols) and (ARow = 1)) then
      begin
        if ChgSens then Break else
        begin
          Arow := GSMOI.rowCount;
          ChgLig := true;
        end;
      end;
      if ChgLig then
      begin
        ACol := GSMOI.ColCount;
        ChgLig := False;
      end;
      if ACol > GSMOI.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := GSMOI.ColCount - 1;
      end;
    end;
  end;
  GSMOI.SynEnabled := OldEna;
  if (Arow <> RowFirst) then
  BEGIN
  	GSRowExitMOI (self,RowFirst,cancel,false);
  	if Cancel then
    begin
    	GSRowEnterMOI (self,RowFirst,cancel,false);
    end else
    begin
    	GSRowEnterMOI (self,Arow,cancel,false);
    end;
  END;
end;

function TOF_BSAISIEAVANCPOC.VisibleFI(TheCur: TGSaisie; TOBL: TOB;Acol: integer): Boolean;
var TypeLigne : string;
begin
  Result := true;
  TypeLigne := TOBL.getValue('TYPELIGNE');
  if (TypeLigne = 'ART')  then
  begin
    if (Acol = TheCur.SG_PHASETRA) then
    begin
      Result := false;
    end;
    exit;
  end;
  if (TypeLigne = 'COM')  then
  begin
    if (Acol <> TheCur.SG_LIBELLE) and
       (Acol <> TheCur.SG_TYPELIG) then
    begin
      Result := false;
    end;
    Exit;
  end;
  //
  if ((Acol = TheCur.SG_QTEMARCHE) or
     (Acol = TheCur.SG_UNITE) or
     (Acol = TheCur.SG_POURCENTAVANC) or
     (Acol = TheCur.SG_QTEMOIS) or
     (Acol = TheCur.SG_QTECUMUL) or
     (Acol = TheCur.SG_RADPU) or
     (Acol = TheCur.SG_FAMILLENIV2) or
     (Acol = TheCur.SG_RADMT) or
     (Acol = TheCur.SG_MTAVANCCALC) or
     (Acol = TheCur.SG_SUMQTECONSO) or
     (Acol = TheCur.SG_DEPENSECUMUL) or
     (Acol = TheCur.SG_ECARTDEPENSE) or
     (Acol = TheCur.SG_RADQTE)) and
     (TypeLigne <> 'ART') then
  begin
    Result := false;
    exit;
  end;
end;

function TOF_BSAISIEAVANCPOC.ZoneAccessibleFI(ACol,ARow: Integer): boolean;
var TOBL: TOB;
    TheGS : TGSaisie;
begin
  Result := True;
  if TOBFI.detail.count = 0 then exit;
  TheGS := TheListGrid.findGrille('GSFI');
  if Arow > TOBFI.detail.count then begin result := false; exit; end;
  TOBL := TOBFI.detail[Arow-1];
  if TOBL = nil then Exit;
  if (Acol = 0) then BEGIN result := false; Exit; END;
  if (Acol = TheGS.SG_TYPELIG) or (ACol = TheGS.SG_PHASETRA) or
     (ACol = TheGS.SG_FAMILLENIV2) or (Acol = TheGS.SG_UNITE) or
     (Acol = TheGS.SG_QTEMARCHE) or (Acol = TheGS.SG_UNITE) or
     (Acol = TheGS.SG_MTAVANCCALC) or (Acol = TheGS.SG_DEPENSECUMUL) or
     (Acol = TheGS.SG_SUMQTECONSO) or (Acol = TheGS.SG_ECARTDEPENSE) then BEGIN  Result := false; Exit; END;
  if not VisibleFI(TheGS,TOBL,Acol) then BEGIN result := false; Exit; end;
end;

procedure TOF_BSAISIEAVANCPOC.ZoneSuivanteOuOkFI(var ACol, ARow: Integer;var Cancel: boolean);
var Sens, ii, Lim: integer;
  OldEna, ChgLig, ChgSens: boolean;
  RowFirst : integer;
begin
	RowFirst := ARow;
  OldEna := GSFI.SynEnabled;
  GSFI.SynEnabled := False;
  Sens := -1;
  ChgLig := (GSFI.Row <> ARow);
  ChgSens := False;
  if GSFI.Row > ARow then Sens := 1 else if ((GSFI.Row = ARow) and (ACol <= GSFI.Col)) then Sens := 1;
  ACol := GSFI.Col;
  ARow := GSFI.Row;
  ii := 0;
  while not ZoneAccessibleFI(ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if Sens = 1 then
    begin
      Lim := GSFI.RowCount ;
      // ---
      if ((ACol = GSFI.ColCount - 1) and (ARow >= Lim)) then
      begin
        if (ChgSens) then Break;
      end;
      if ChgLig then
      begin
        ACol := GSFI.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < GSFI.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        if ARow >= Lim then
        begin
          ARow := 1;
          Acol := 1;
        end else
        begin
          ACol := GSFI.FixedCols;
        end;
      end;
    end else
    begin
      if ((ACol = GSFI.FixedCols) and (ARow = 1)) then
      begin
        if ChgSens then Break else
        begin
          Arow := GSFI.rowCount;
          ChgLig := true;
        end;
      end;
      if ChgLig then
      begin
        ACol := GSFI.ColCount;
        ChgLig := False;
      end;
      if ACol > GSFI.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := GSFI.ColCount - 1;
      end;
    end;
  end;
  GSFI.SynEnabled := OldEna;
  if (Arow <> RowFirst) then
  BEGIN
  	GSRowExitFI (self,RowFirst,cancel,false);
  	if Cancel then
    begin
    	GSRowEnterFI (self,RowFirst,cancel,false);
    end else
    begin
    	GSRowEnterFI (self,Arow,cancel,false);
    end;
  END;
end;

function TOF_BSAISIEAVANCPOC.VisibleST(TheCur: TGSaisie; TOBL: TOB;Acol: integer): Boolean;
//   stLigST  = 'IND;TYPELIGNE;BM8_PHASETRA;BM8_CODEMARCHE;BM8_FAMILLENIV2;BM8_LIBELLE;BM8_QTEMARCHE;BM8_UNITE;PU;DEPENSECUMUL;DEPENSEESTIME;POURCENTCALC;BM8_POURCENTAVANC;MTAVANCCORRIGE;'; // Sous-traitance
var TypeLigne : string;
begin
  Result := true;
  TypeLigne := TOBL.getValue('TYPELIGNE');
  if (TypeLigne = 'ART')  then
  begin
    if (Acol = TheCur.SG_PHASETRA) then
    begin
      Result := false;
    end;
    exit;
  end;
  if (TypeLigne = 'COM')  then
  begin
    if (Acol <> TheCur.SG_LIBELLE) and
       (Acol <> TheCur.SG_TYPELIG) then
    begin
      Result := false;
    end;
    Exit;
  end;
  //
  if ((Acol = TheCur.SG_QTEMARCHE) or
     (Acol = TheCur.SG_UNITE) or
     (Acol = TheCur.SG_PU) or
     (Acol = TheCur.SG_DEPENSECUMUL) or
     (Acol = TheCur.SG_DEPENSEESTIME) or
     (Acol = TheCur.SG_POURCENTCALC) or
     (Acol = TheCur.SG_POURCENTAVANC) or
     (Acol = TheCur.SG_MTAVANCCORRIGE) ) and
     (TypeLigne <> 'ART') then
  begin
    Result := false;
    exit;
  end;
end;

function TOF_BSAISIEAVANCPOC.ZoneAccessibleST(ACol,ARow: Integer): boolean;
var TOBL: TOB;
    TheGS : TGSaisie;
begin
   Result := True;
  if TOBST.detail.count = 0 then exit;
  TheGS := TheListGrid.findGrille('GSST');
  if Arow > TOBST.detail.count then begin result := false; exit; end;
  TOBL := TOBST.detail[Arow-1];
  if TOBL = nil then Exit;
  if (Acol = 0) then BEGIN result := false; Exit; END;
  if (Acol = TheGS.SG_TYPELIG) or (ACol = TheGS.SG_PHASETRA) or
     (ACol = TheGS.SG_CODEMARCHE) or
     (ACol = TheGS.SG_QTEMARCHE) or (Acol = TheGS.SG_UNITE) or
     (ACol = TheGS.SG_DEPENSECUMUL) or (Acol =TheGS.SG_POURCENTCALC) or (Acol = TheGS.SG_UNITE) or
     (Acol = TheGS.SG_PU) or (Acol =TheGS.SG_MTAVANCCORRIGE) or (Acol=TheGS.SG_ECARTDEPENSE) then
     BEGIN
       Result := false;
       Exit;
     END;
  if not VisibleST(TheGS,TOBL,Acol) then BEGIN result := false; Exit; end;
end;

procedure TOF_BSAISIEAVANCPOC.ZoneSuivanteOuOkST(var ACol, ARow: Integer;var Cancel: boolean);
var Sens, ii, Lim: integer;
  OldEna, ChgLig, ChgSens: boolean;
  RowFirst : integer;
begin
	RowFirst := ARow;
  OldEna := GSST.SynEnabled;
  GSST.SynEnabled := False;
  Sens := -1;
  ChgLig := (GSST.Row <> ARow);
  ChgSens := False;
  if GSST.Row > ARow then Sens := 1 else if ((GSST.Row = ARow) and (ACol <= GSST.Col)) then Sens := 1;
  ACol := GSST.Col;
  ARow := GSST.Row;
  ii := 0;
  while not ZoneAccessibleST(ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if Sens = 1 then
    begin
      Lim := GSST.RowCount;
      // ---
      if ((ACol = GSST.ColCount - 1) and (ARow >= Lim)) then
      begin
        if (ChgSens) then Break;
      end;
      if ChgLig then
      begin
        ACol := GSST.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < GSST.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        if ARow >= Lim then
        begin
          ARow := 1;
          Acol := 1;
        end else
        begin
          ACol := GSST.FixedCols;
        end;
      end;
    end else
    begin
      if ((ACol = GSST.FixedCols) and (ARow = 1)) then
      begin
        if ChgSens then Break else
        begin
          Arow := GSST.rowCount;
          ChgLig := true;
        end;
      end;
      if ChgLig then
      begin
        ACol := GSST.ColCount;
        ChgLig := False;
      end;
      if ACol > GSST.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := GSST.ColCount - 1;
      end;
    end;
  end;
  GSST.SynEnabled := OldEna;
  if (Arow <> RowFirst) then
  BEGIN
  	GSRowExitST (self,RowFirst,cancel,false);
  	if Cancel then
    begin
    	GSRowEnterST (self,RowFirst,cancel,false);
    end else
    begin
    	GSRowEnterST (self,Arow,cancel,false);
    end;
  END;
end;

procedure TOF_BSAISIEAVANCPOC.RecupereConsos;

  function FindType(TLC : TOB; var Niveau : integer) : string;
  var TT : TOB;
      II : Integer;
      NN : integer;
  begin
    Niveau := 0;
    Result := '';
    // cas particuler des Frais indirects
    if Pos(TLC.GetString('BCO_FAMILLENIV2'),FamFraisIndirect)>0 then
    begin
      Result := '002';
      TT := TTS.FindFirst(['TYPE'],['002'],true);
      if TT <> nil then Niveau := TT.GetInteger('BPC_PN1')+1;
      exit;
    end;
    // --------------------
    TT := TTS.FindFirst(['BPC_PHASETRA'],[TLC.GetString('BCO_PHASETRA')],true);
    if TT <> nil then
    begin
      Result := TT.GetString('TYPE');
      for II := 1 to 9 do
      begin
        NN := TT.GetInteger('BPC_PN'+InttoStr(II));
        if NN <> 0 then Niveau := II+1 else break;
      end;
    end;
  end;

  procedure SetCumulConsosPRE (TL,TLC : TOB);
  begin
    TL.SetDouble('BM8_QTECONSOPRE',TLC.GetDouble('SUMQTE'));
    TL.SetDouble('BM8_MTCONSOPRE',TLC.GetDouble('SUMMT'));
  end;

  procedure SetCumulConsos (TL,TLC : TOB);
  begin
    TL.SetDouble('BM8_QTECONSOMOIS',TL.GetDouble('BM8_QTECONSOMOIS')+TLC.GetDouble('SUMQTE'));
    TL.SetDouble('BM8_MTCONSOMOIS',TL.GetDouble('BM8_MTCONSOMOIS')+TLC.GetDouble('SUMMT'));
  end;

  procedure AjouteLigneHP(TL : TOB);
  var TF,TP : TOB;
      Index : Integer;
  begin
    TF := CreTOBLigne;
    TF.SetDateTime('DATEMVT',TL.GetDateTime('BCO_DATEMOUV'));
    TF.SetString('CODEARTICLE',TL.GetString('BCO_CODEARTICLE'));
    TF.SetString('LIBELLE',TL.GetString('BCO_LIBELLE'));
    TF.SetDouble('QUANTITE',TL.Getdouble('BCO_QUANTITE'));
    TF.SetDouble('PUACH',TL.GetDouble('BCO_DPA'));
    TF.SetDouble('MONTANTACH',TL.GetDouble('BCO_MONTANTACH'));
    TF.ChangeParent(TOBHP,-1);
  end;


  procedure AjouteLigne(TL : TOB;TheType,CurrentPhase : string;NiveauImbric : integer; MoisCourant : boolean=true);
  var TF,TP : TOB;
      Index : Integer;
  begin
    TF := FindLigneConso(TheType,CurrentPhase,TL.GetString('BCO_FAMILLENIV2'),'');
    //
    if (TF = nil) then
    begin
      TF := CreTOBLigne;
      TF.SetString('TYPELIGNE','ART');
      if TheType = '002' then
      begin
        NiveauImbric := 1;
        CurrentPhase := '';
      end;
      TF.SetInteger('NIVEAU',NiveauImbric);
      TF.SetString('BM8_PHASETRA',CurrentPhase);
      TF.SetString('BM8_CODEMARCHE','');
      //
      TF.SetString('BM8_FAMILLENIV2',TL.GetString('BCO_FAMILLENIV2'));
      //
      TF.SetString('BM8_LIBELLE',TL.GetString('LIBELLE'));
      TF.SetString('BM8_UNITE','');
      TF.SetString('NEWONE','-');
      if not IsParagraphe(TL) then
      begin
        if TheType <>'003' then
        begin
          TP := GetEmplacementLigne (TheType,CurrentPhase,TF.GetString('BM8_FAMILLENIV2'));
        end else
        begin
          TP := FindEmplacementLigne (TheType,CurrentPhase) ;
        end;
//        TP := FindEmplacementLigne (TheType,CurrentPhase);
        if TP = nil then Index := -1
                    else Index := TP.GetIndex;
      end else
      begin
        Index := -1;
      end;
      if TheType = '000' then TF.Changeparent(TOBPP,Index)
      else if TheType = '001' then TF.Changeparent(TOBMOI,Index)
      else if TheType = '002' then TF.Changeparent(TOBFI,Index)
      else if TheType = '003' then TF.Changeparent(TOBST,Index);
    end;
    if MoisCourant then SetCumulConsos (TF,TL)
                   else SetCumulConsosPRE (TF,TL);
  end;


var II : Integer;
    NiveauImbric : Integer; 
    TTC,TLC : TOB;
    QQ : TQuery;
    SQL : String;
    TheType,ThePhase : string;
begin
  TTC := TOB.Create ('LES CONSO MOIS',nil,-1);
  TRY
    // conso précédentes
    SQL := 'SELECT BCO_PHASETRA,BCO_FAMILLENIV2,SUM(BCO_QUANTITE) AS SUMQTE, SUM(BCO_MONTANTACH) AS SUMMT,'+
           '(SELECT CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="FN2" AND CC_CODE=BCO_FAMILLENIV2) AS LIBELLE '+
           'FROM CONSOMMATIONS '+
           'WHERE '+
           'BCO_AFFAIRE="'+CodeAffaire+'" AND '+
           'BCO_DATEMOUV < "'+USDATETIME(DebutMois)+'" '+
           'GROUP BY BCO_PHASETRA,BCO_FAMILLENIV2 '+
           'ORDER BY BCO_PHASETRA';
    QQ := OpenSQL(SQL,True,-1,'',true);
    if Not QQ.eof then
    begin
      TTC.LoadDetailDB('CONSOMMATIONS','','',QQ,false);
    end;
    ferme(QQ);
    for II := 0 to TTC.detail.count -1 do
    begin
      TLC := TTC.detail[II];
      XX.MoveK2000;
      TheType := FindType(TLC,NiveauImbric);
      ThePhase := TLC.GetString('BCO_PHASETRA');
      if (TheTYpe <> '003') and (TheType <> '') then
      begin
        AjouteLigne (TLC,TheType,ThePhase,NiveauImbric,false);
      end;
    end;
    //
    TTC.ClearDetail;
    //  --> consommations du mois
    SQL := 'SELECT BCO_PHASETRA,BCO_FAMILLENIV2,SUM(BCO_QUANTITE) AS SUMQTE, SUM(BCO_MONTANTACH) AS SUMMT,'+
           '(SELECT CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="FN2" AND CC_CODE=BCO_FAMILLENIV2) AS LIBELLE '+
           'FROM CONSOMMATIONS '+
           'WHERE '+
           'BCO_AFFAIRE="'+CodeAffaire+'" AND '+
           'BCO_DATEMOUV >= "'+USDATETIME(DebutMois)+'" AND '+
           'BCO_DATEMOUV <= "'+USDATETIME(FinMois)+'" '+
           'GROUP BY BCO_PHASETRA,BCO_FAMILLENIV2 '+
           'ORDER BY BCO_PHASETRA';
    QQ := OpenSQL(SQL,True,-1,'',true);
    if Not QQ.eof then
    begin
      TTC.LoadDetailDB('CONSOMMATIONS','','',QQ,false);
    end;
    ferme(QQ);
    for II := 0 to TTC.detail.count -1 do
    begin
      XX.MoveK2000;
      TLC := TTC.detail[II];
      TheType := FindType(TLC,NiveauImbric);
      ThePhase := TLC.GetString('BCO_PHASETRA');
      if (TheTYpe <> '003') and (TheType <> '') then
      begin
        AjouteLigne (TLC,TheType,ThePhase,NiveauImbric);
      end;
    end;
    // recupération des hors phases --> erreur de saisie
    TTC.ClearDetail;
    SQL := 'SELECT BCO_DATEMOUV,BCO_CODEARTICLE,BCO_LIBELLE,BCO_QUANTITE,BCO_DPA,BCO_MONTANTACH '+
           'FROM CONSOMMATIONS '+
           'WHERE '+
           'BCO_NATUREPIECEG IN ("LBT","") AND '+
           'BCO_AFFAIRE="'+CodeAffaire+'" AND '+
           'BCO_PHASETRA="" AND '+
           'BCO_DATEMOUV <= "'+USDATETIME(FinMois)+'" ';
    QQ := OpenSQL(SQL,True,-1,'',true);
    if Not QQ.eof then
    begin
      TTC.LoadDetailDB('CONSOMMATIONS','','',QQ,false);
    end;
    ferme(QQ);
    for II := 0 to TTC.detail.count -1 do
    begin
      XX.MoveK2000;
      TLC := TTC.detail[II];
      AjouteLigneHP (TLC);
    end;

  FINALLY
    TTC.Free;
  END;
end;

procedure TOF_BSAISIEAVANCPOC.CalculeMtAvt(TOBL: TOB);
begin
  if TOBL.GetDouble('BM8_QTEMARCHE') <> 0 then TOBL.SetDouble('BM8_MTCUMUL',ARRONDI(TOBL.GetDouble('BM8_QTECUMUL')*(TOBL.GetDouble('BM8_MTMARCHE')/TOBL.GetDouble('BM8_QTEMARCHE')),V_PGI.okdecQ));
  TOBL.SetDouble('BM8_MTMOIS',arrondi(TOBL.GetDouble('BM8_MTCUMUL')-TOBL.GetDouble('BM8_MTPREC'),V_PGI.OkdecV));
end;

procedure TOF_BSAISIEAVANCPOC.CalculeQteAvt(TOBL : TOB);
begin
  TOBL.SetDouble('BM8_QTECUMUL',ARRONDI(TOBL.GetDouble('BM8_QTEMARCHE')*(TOBL.GetDouble('BM8_POURCENTAVANC')/100),V_PGI.OkDecQ));
  TOBL.SetDouble('BM8_QTEMOIS',ARRONDI(TOBL.GetDouble('BM8_QTECUMUL')-TOBL.GetDouble('BM8_QTEPREC'),V_PGI.OkDecQ));
end;

procedure TOF_BSAISIEAVANCPOC.CalculeProjete (TOBL : TOB);
begin
  TOBL.SetDouble('QTEPROJETE',TOBL.GetDouble('SUMQTECONSO')+TOBL.GetDouble('BM8_RADQTE'));
end;


procedure TOF_BSAISIEAVANCPOC.CalculePourcentAvt(TOBL:TOB);
begin
  if TOBL.GetDouble('SUMQTECONSO') + TOBL.GetDouble('BM8_RADQTE') <> 0 then
  begin
    TOBL.SetDouble('BM8_POURCENTAVANC',ARRONDI((TOBL.GetDouble('SUMQTECONSO') / (TOBL.GetDouble('SUMQTECONSO') + TOBL.GetDouble('BM8_RADQTE')))*100,2))
  end else
  begin
    TOBL.SetDouble('BM8_POURCENTAVANC',0);
  end;
end;

procedure TOF_BSAISIEAVANCPOC.CalcPourcentAvcPP(TOBL : TOB);
begin
  if TOBL.GetDouble('BM8_QTEMARCHE') <> 0 then
  begin
    TOBL.SetDouble('BM8_POURCENTAVANC',ARRONDI(TOBL.GetDouble('BM8_QTECUMUL') / TOBL.GetDouble('BM8_QTEMARCHE')*100,2));
  end else
  begin
    TOBL.SetDouble('BM8_POURCENTAVANC',0);
  end;
end;


procedure TOF_BSAISIEAVANCPOC.CalculeMtRAD (TOBL : TOB);
begin
  TOBL.SetDouble('BM8_RADMT',ARRONDI(TOBL.GetDouble('BM8_RADQTE') *TOBL.GetDouble('BM8_RADPU'),V_PGI.OkDecV));
end;

procedure TOF_BSAISIEAVANCPOC.CalcQteCumul (TOBL : TOB);
begin
  TOBL.SetDouble('BM8_QTECUMUL',ARRONDI(TOBL.GetDouble('BM8_QTEPREC')+TOBL.GetDouble('BM8_QTEMOIS'),V_PGI.OkDecQ));
end;

procedure TOF_BSAISIEAVANCPOC.CalculeAvancCorrigeST (TOBL : TOB);
var PP : double;
begin
  PP := 0;
  if TOBL.GetDouble('BM8_POURCENTAVANC')<> 0 then
  begin
    PP := TOBL.GetDouble('BM8_POURCENTAVANC');
    TOBL.SetDouble('MTAVANCCORRIGE',ARRONDI(TOBL.GetDouble('DEPENSEESTIME')* PP /100,2));
  end else
  begin
    TOBL.SetDouble('MTAVANCCORRIGE',TOBL.GetDouble('DEPENSECUMUL'));
  end;

end;

procedure TOF_BSAISIEAVANCPOC.CalculFinal;
var II : Integer;
    TOBL  : TOB;
begin
  for II := 0 to TOBPP.detail.count -1 do
  begin
    TOBL := TOBPP.detail[II];
    XX.MoveK2000;
    if TOBL.GetString('TYPELIGNE') <> 'ART' then continue;
    if TOBL.GetDouble('BM8_QTEMARCHE')<> 0 then TOBL.SetDouble('BM8_RADPU',ARRONDI(TOBL.GetDouble('BM8_MTMARCHE')/TOBL.GetDouble('BM8_QTEMARCHE'),V_PGI.OkDecP))
                                           else TOBL.SetDouble('BM8_RADPU',ARRONDI(TOBL.GetDouble('BM8_MTMARCHE'),V_PGI.OkDecP));
    TOBL.SetDouble('SUMQTECONSO',TOBL.GetDouble('BM8_QTECONSOPRE')+TOBL.GetDouble('BM8_QTECONSOMOIS'));
    TOBL.SetDouble('DEPENSECUMUL',TOBL.GetDouble('BM8_MTCONSOPRE')+TOBL.GetDouble('BM8_MTCONSOMOIS'));
    CalcQteCumul(TOBL);
    CalcPourcentAvcPP(TOBL);
    CalculeProjete(TOBL);
  end;
  //
  for II := 0 to TOBMOI.detail.count -1 do
  begin
    TOBL := TOBMOI.detail[II];
    XX.MoveK2000;
    if TOBL.GetString('TYPELIGNE') <> 'ART' then continue;
    if TOBL.GetDouble('BM8_QTEMARCHE')<> 0 then TOBL.SetDouble('BM8_RADPU',ARRONDI(TOBL.GetDouble('BM8_MTMARCHE')/TOBL.GetDouble('BM8_QTEMARCHE'),V_PGI.OkDecP))
                                           else TOBL.SetDouble('BM8_RADPU',ARRONDI(TOBL.GetDouble('BM8_MTMARCHE'),V_PGI.OkDecP));
    TOBL.SetDouble('SUMQTECONSO',TOBL.GetDouble('BM8_QTECONSOPRE')+TOBL.GetDouble('BM8_QTECONSOMOIS'));
    TOBL.SetDouble('DEPENSECUMUL',TOBL.GetDouble('BM8_MTCONSOPRE')+TOBL.GetDouble('BM8_MTCONSOMOIS'));
    CalcQteCumul(TOBL);
    CalculeMtAvt (TOBL);
    CalculePourcentAvt (TOBL);
    CalculeMtRAD (TOBL);
  end;
  //
  for II := 0 to TOBFI.detail.count -1 do
  begin
    TOBL := TOBFI.detail[II];
    XX.MoveK2000;
    if TOBL.GetString('TYPELIGNE') <> 'ART' then continue;
    if (TOBL.GetDouble('BM8_RADPU') =0) or (TOBL.GetDouble('BM8_RADQTE') =0) then
    begin
      if TOBL.GetDouble('BM8_QTEMARCHE')<> 0 then TOBL.SetDouble('BM8_RADPU',ARRONDI(TOBL.GetDouble('BM8_MTMARCHE')/TOBL.GetDouble('BM8_QTEMARCHE'),V_PGI.OkDecP))
                                             else TOBL.SetDouble('BM8_RADPU',ARRONDI(TOBL.GetDouble('BM8_MTMARCHE'),V_PGI.OkDecP));
    end;
    TOBL.SetDouble('SUMQTECONSO',TOBL.GetDouble('BM8_QTECONSOPRE')+TOBL.GetDouble('BM8_QTECONSOMOIS'));
    TOBL.SetDouble('DEPENSECUMUL',TOBL.GetDouble('BM8_MTCONSOPRE')+TOBL.GetDouble('BM8_MTCONSOMOIS'));
    CalculeMtRAD (TOBL);
    TOBL.SetDouble('MTAVANCCALC',TOBL.GetDouble('BM8_MTMARCHE')-TOBL.GetDouble('BM8_RADMT'));
    //
    TOBL.SetDouble('ECARTDEPENSE',TOBL.GetDouble('BM8_MTCONSOMOIS'));
  end;
  //
  for II := 0 to TOBST.detail.count -1 do
  begin
    TOBL := TOBST.detail[II];
    XX.MoveK2000;
    if TOBL.GetString('TYPELIGNE') <> 'ART' then continue;
    //
    if TOBL.GetDouble('DEPENSEESTIME') <> 0 then TOBL.SetDouble('POURCENTCALC', ARRONDI(TOBL.GetDouble('DEPENSECUMUL') /TOBL.GetDouble('DEPENSEESTIME')*100,2))
                                            else TOBL.SetDouble('POURCENTCALC', 100);
    if TOBL.GetDouble('BM8_QTEMARCHE') <> 0 then TOBL.SetDouble('PU',ARRONDI(TOBL.GetDouble('BM8_MTMARCHE')/TOBL.GetDouble('BM8_QTEMARCHE'),V_PGI.OkDecP));
    TOBL.SetDouble('ECARTDEPENSE',TOBL.GetDouble('BM8_MTCONSOMOIS'));
    CalculeAvancCorrigeST (TOBL);
  end;
end;

procedure TOF_BSAISIEAVANCPOC.CalculePURAD(TOBL : TOB);
begin
  if TOBL.GetDouble('BM8_RADQTE') <> 0 then TOBL.SetDouble('BM8_RADPU',ARRONDI(TOBL.GetDouble('BM8_RADMT')/TOBL.GetDouble('BM8_RADQTE'),V_PGI.OkdecP));
end;

procedure TOF_BSAISIEAVANCPOC.GSSTElipsisClick(Sender: TObject);
  function GetTiersST(Affaire,CodeMarche : string) : string;
  var QQ : TQuery;
  begin
    QQ := OpenSQL('SELECT BM0_FOURNISSEUR FROM BTMARCHEST WHERE BM0_AFFAIRE="'+Affaire+'" AND BM0_MARCHE="'+Copy(CodeMarche,1,8)+'"',True,1,'',true);
    if not QQ.eof then
    begin
      Result := QQ.fields[0].AsString;
    end;
    ferme (QQ);
  end;
var FGS : TGSaisie;
    TT : TOB;
    TL : TOB;
begin
  FGS := TheListGrid.findGrille('GSST');
  if GSST.col = FGS.SG_DEPENSEESTIME then
  begin
    TL := TOBST.detail[GSST.row-1]; if TL = nil then Exit;
    if TL.GetDouble('DEPENSEESTIME')= 0 then exit;
    TT := TOB.Create ('LES PARAMS',nil,-1);
    try
      TT.AddChampSupValeur('AFFAIRE',TL.GetString('BM8_AFFAIRE'));
      TT.AddChampSupValeur('CODEMARCHE',TL.GetString('BM8_CODEMARCHE'));
      TT.AddChampSupValeur('SOUSTRAIT',GetTiersST(TL.GetString('BM8_AFFAIRE'),TL.GetString('BM8_CODEMARCHE')));
      TT.AddChampSupValeur('MODIF','-');
      TheTOB := TT;
      AGLLanceFiche('BTP','BTMODIFMARCHEST','','','ACTION=MODIFICATION');
      TheTOB := nil;
      if TT.GetString('MODIF')='X' then
      begin
        GetDepensesEstimes (TL);
        if TL.GetDouble('DEPENSEESTIME') <> 0 then TL.SetDouble('POURCENTCALC', ARRONDI(TL.GetDouble('DEPENSECUMUL') /TL.GetDouble('DEPENSEESTIME')*100,2))
                                              else TL.SetDouble('POURCENTCALC', 100);
        CalculeAvancCorrigeST (TL);
        afficheLigne(GSST,TL,FGS.fListSaisie,GSST.row-1);
        FGS.StCellCur := GSST.cells[GSST.Col,GSST.Row];
        DataModif := true;
      end;
    finally
      TT.free;
    end;
  end;
end;

procedure TOF_BSAISIEAVANCPOC.GetDepensesEstimes (TL : TOB);
var SQL : string;
    QQ : TQuery;
begin
  TL.SetDouble('DEPENSEESTIME',0);
  SQL :='SELECT SUM(BM1_MONTANT) '+
        'FROM BTMARCHESTDET '+
        'WHERE '+
        'BM1_AFFAIRE="'+TL.GetString('BM8_AFFAIRE')+'" AND '+
        'BM1_MARCHE="'+copy(TL.GetString('BM8_CODEMARCHE'),1,8)+'" AND '+
        'BM1_TYPE IN ("001","002")';
  QQ := OpenSQL(SQL,True,-1,'',true);
  if not QQ.eof then
  begin
    QQ.first;
    TL.SetDouble('DEPENSEESTIME',TL.GetDouble('DEPENSEESTIME')+QQ.Fields[0].AsFloat);
  end;
  TL.SetDouble('DEPENSEESTIME',Arrondi(TL.GetDouble('DEPENSEESTIME'),2));
  Ferme(QQ);
end;


procedure TOF_BSAISIEAVANCPOC.CalculSousTraitance;

  procedure GetBASTFAC(TL : TOB);
  var SQL : string;
      QQ : TQuery;
  begin
    TL.SetDouble('DEPENSECUMUL',0);
    SQL :='SELECT BM4_CODEMARCHE,BM4_NUMSITUATION,BM4_REFGESCOM,'+
          '(SELECT SUM(BM5_MTSITUATION) FROM BASTLIG WHERE BM5_AFFAIRE=BM4_AFFAIRE AND BM5_CODEMARCHE=BM4_CODEMARCHE AND BM5_TYPELBAST="001" AND BM5_TYPELIG="TOT" AND BM5_NUMSITUATION=BM4_NUMSITUATION AND BM5_PAIEMENTPOC=BM4_PAIEMENTPOC) AS MTSITUATION '+
          'FROM BASTENT '+
          'WHERE '+
          'BM4_AFFAIRE="'+TL.GetString('BM8_AFFAIRE')+'" AND '+
          'BM4_CODEMARCHE="'+TL.GetString('BM8_CODEMARCHE')+'" AND '+
          'BM4_REFGESCOM <> ""';
    QQ := OpenSQL(SQL,True,-1,'',true);
    if not QQ.eof then
    begin
      QQ.First;
      repeat
        TL.SetDouble('DEPENSECUMUL',TL.GetDouble('DEPENSECUMUL')+QQ.Fields[3].AsFloat);
        QQ.Next;
      until QQ.Eof;
    end;
    TL.SetDouble('BM8_MTCONSOMOIS',TL.getDouble('DEPENSECUMUL')-TL.GetDouble('BM8_MTCONSOPRE'));
    Ferme(QQ);
  end;

var  II : integer;
     TL : TOB;
begin
  for II := 0 to TOBST.detail.count -1 do
  begin
    TL := TOBST.detail[II];
    if TL.GetString('TYPELIGNE')= 'ART' then
    begin
      XX.MoveK2000;
      GetBASTFAC(TL);
      GetDepensesEstimes (TL);
    end;
  end;
end;

procedure TOF_BSAISIEAVANCPOC.GridKeyDownFI(Sender: TObject; var Key: Word;Shift: TShiftState);
begin
  if ActionResult = taConsult then Exit;
  case Key of
    13 : if (Shift = []) then
    begin
    	Key := 0;
      SendMessage(THedit(Sender).Handle, WM_KeyDown, VK_TAB, 0);
    end;
  end;

end;

procedure TOF_BSAISIEAVANCPOC.GridKeyDownMOI(Sender: TObject;var Key: Word; Shift: TShiftState);
begin
  case Key of
    13 : if (Shift = []) then
    begin
    	Key := 0;
      SendMessage(THedit(Sender).Handle, WM_KeyDown, VK_TAB, 0);
    end;
  end;

end;

procedure TOF_BSAISIEAVANCPOC.GridKeyDownST(Sender: TObject; var Key: Word;Shift: TShiftState);
begin
  case Key of
    13 : if (Shift = []) then
    begin
    	Key := 0;
      SendMessage(THedit(Sender).Handle, WM_KeyDown, VK_TAB, 0);
    end;
  end;

end;

procedure TOF_BSAISIEAVANCPOC.BeforeUpdate;
var Mois : string;
begin
  Mois := GetControlText('BMOIS');
  ExecuteSQL('DELETE FROM BAVANCTECH WHERE BM9_AFFAIRE="'+CodeAffaire+'" AND BM9_MOIS='+Mois);
  ExecuteSQL('DELETE FROM BAVANCTECHDET WHERE BM8_AFFAIRE="'+CodeAffaire+'" AND BM8_MOIS='+Mois);
end;

procedure TOF_BSAISIEAVANCPOC.UpdateDatas;
var II : Integer;
begin
  TTD.ClearDetail;
  II := 0;
  repeat
    if TOBPP.detail[II].GetString('TYPELIGNE')='ART' then TOBPP.detail[II].ChangeParent(TTD,-1) else inc(II);
  until II >= TOBPP.detail.Count;
  II := 0;
  repeat
    if TOBMOI.detail[II].GetString('TYPELIGNE')='ART' then TOBMOI.detail[II].ChangeParent(TTD,-1) else inc(II);
  until II >= TOBMOI.detail.Count;
  II := 0;
  repeat
    if TOBFI.detail[II].GetString('TYPELIGNE')='ART' then TOBFI.detail[II].ChangeParent(TTD,-1) else inc(II);
  until II >= TOBFI.detail.Count;
  II := 0;
  repeat
    if TOBST.detail[II].GetString('TYPELIGNE')='ART' then TOBST.detail[II].ChangeParent(TTD,-1) else inc(II);
  until II >= TOBST.detail.Count;
  //
  TTD.SetAllmodifie(true);
  TTD.detail.Sort('BM8_NATUREPIECEG;BM8_SOUCHE;BM8_NUMERO;BM8_INDICEG;BM8_MOIS;BM8_PHASETRA;BM8_FAMILLENIV2;BM8_CODEMARCHE');
  TTD.InsertDBByNivel(false); 
end;

procedure TOF_BSAISIEAVANCPOC.SetInConsultation;
begin
  TfVierge(Ecran).TypeAction := taConsult;
  GSPP.Options := GSPP.Options+ [goRowSelect]; 
  GSMOI.Options := GSPP.Options;
  GSFI.Options := GSPP.Options;
  GSST.Options := GSPP.Options;
  SetControlVisible('Bvalider',false);
  SetControlVisible('Bdelete',false);
end;

Initialization
  RegisterAglProc( 'SaisAvancementPOC',True,6,AGLAvancementPOC);
  registerclasses ( [ TOF_BSAISIEAVANCPOC ] ) ;

end.

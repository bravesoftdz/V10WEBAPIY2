unit UCotraitance;

interface
Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     MaineAGL,
{$ENDIF}
     Hent1,
		 Menus,
		 HCtrls,
     HMsgBox,
     SysUtils,
     AglInit,
     forms,
     UentCommun,
     SaisUtil,
     CopierCollerUtil,
     UtilPGI,
     uTob,
     Ent1,
     NomenUtil;
type
  TModegestion = (TmgIntervenant,TmgCotraitance,TmgSousTraitance);

	TInfoSSTrait = class (TObject)
  	private
      fCode : string;
      fCoefFG : double;
      fCoefMarg : double;
    	procedure SetCode(const Value: string);
    public
      constructor create;
      destructor destroy; override;
      property Code : string write SetCode;
      property CoefFg : double read fCoefFG;
      property Coefmarg : double read fCoefMarg;
  end;

	TPieceCotrait = class (Tobject)
  	private
      fDEV : Rdevise;
      fInfoSSTrait : TInfoSSTrait;
      fModifSsDetail : boolean;
    	fusable : boolean;
      fCreatedPop : boolean;
      fmaxItems : integer;
      fMaxpopyItems : integer;
      fcopiercoller : TCopieColleDoc;
    	FF : TForm;
      fTOBpiece : TOB;
      fTOBaffaire : TOB;
      fTOBOuvrage : TOB;
      fTOBArticles : TOB;
      ftobPieceTrait : TOB;
      fTOBPieceInterv : TOB;
      fTOBligneArecalc : TOB;
      fTOBBases : TOB;
      GS : THGrid;
      fPOPY : TPopupMenu;
      fPOPGS : TPopupMenu;
      fMenuItem: array[0..5] of TMenuItem;
      fPOPYItem : array[0..3] of TmenuItem;
    	procedure DefiniMenuPop(Parent: Tform);
      //
    	procedure Setaffaire(const Value: TOB);
    	procedure MenuEnabled(State: boolean);
      procedure GSGestionSoustrait (Sender : Tobject);

      procedure GSAffectInterne (Sender : tobject);
    	procedure AffecteInterne(var Ligne: integer);
    	procedure AffecteInterneSousDetail(TOBPere : TOB; IndiceNomen : integer; TOBOUV: TOB; var indice : integer; PrixBloque,FromExcel : boolean);
    	procedure AffecteInterneSousDetail1(TOBOUV: TOB; IndiceNomen : integer; PrixBloque,FromExcel : boolean);
      procedure AffecteInterneOuvrage(TOBL: TOB; var indice : integer);
    	procedure AffecteInterneSDetail (TOBL : TOB; PrixBloque,FromExcel : boolean);
    	procedure InitLigneInterne(TOBL: TOB; var indice : integer);

      procedure GSAffectCotrait (Sender : tobject);
    	procedure AffecteCoTraitanceLigSDetail(TOBL : TOB; Fournisseur, LibelleFou,NatureTravail: string);
    	procedure AffecteCoTraitanceOuvrage(Fournisseur, LibelleFou,NatureTravail: string;TOBL: TOB; Var NextLig : integer; OkInc : boolean);
    	procedure AffecteCoTraitanceSousDetail(TOBPere,TOBOUV: TOB; IndiceNomen : Integer;Fournisseur,LibelleFou,NatureTravail: string;Var NextLig : integer; OkINc : boolean);
    	procedure AffecteCoTraitanceSousDetail1(TOBOUV: TOB;indiceNomen: integer; Fournisseur,LibelleFou,NatureTravail: string);
    	procedure AffecteCoTraitanceLigne(Fournisseur: string; var Indice: integer);
    	procedure SetInfoCotraitancelig(TOBL: TOB; Fournisseur,LibelleFou: string; var Indice : integer);
      //
      procedure GSAffectSoustrait (Sender : tobject);
    	procedure AffecteSousTraitanceLigSDetail(TOBL : TOB; Fournisseur, ModePaie,CodeMarche,LibelleFou,NatureTravail: string);
    	procedure AffecteSousTraitanceOuvrage(TOBL: TOB; Fournisseur, LibelleFou,ModePaie,CodeMarche,NatureTravail: string;Var NextLig : integer; OkInc : boolean; PrixBloque,FromExcel : boolean);
    	procedure AffecteSousTraitanceSousDetail(TOBPere,TOBOUV: TOB; IndiceNomen : Integer;Fournisseur,ModePaie,CodeMarche,LibelleFou,NatureTravail: string;Var NextLig : integer; OkINc : boolean; PrixBloque,FromExcel : boolean);
    	procedure AffecteSousTraitanceSousDetail1(TOBOUV: TOB;indiceNomen: integer; Fournisseur, ModePaie,CodeMarche,LibelleFou,NatureTravail: string; PrixBloque,FromExcel : boolean);
    	procedure SetInfoSoustraitancelig(TOBL: TOB; Fournisseur,ModePaie,CodeMarche,LibelleFou: string; var Ligne : integer; PrixBloque,FromExcel : boolean);
    	procedure AffecteSousTraitanceLigne(Fournisseur,ModePaie,CodeMarche: string; var Ligne: integer);
      //
      procedure GSCotraittableau (sender : TObject);
    	procedure NettoiePieceTrait (TOBpieceTrait : TOB);
    	function  FindLignePere(TOBL: TOB): TOB;
      procedure ReinitDetailOuvrage(TOBO: TOB);
      procedure ReinitOuvrage(TOBL: TOB);
    	function AddLigneCalcul(TOBLP: TOB): Boolean;
      procedure Setpopup (WithCotrait,WithSousTrait : boolean);
      function FindSousTrait : Boolean;
      function FindCoTrait : Boolean;
    	function PieceTraitUsable: boolean;
      function ExisteCoTraitants(CodeAffaire : string) : boolean;
      procedure SetInfoArt (TOBL : TOB; PrixBloque,FromExcel : boolean);
    	procedure RecalcPaPRPV(TOBL: TOB; PrixBloque,FromExcel : boolean);
    	procedure SetValeurs(TOBL: TOB; valeurs: T_Valeurs);
    	procedure SetValeursOuv(TOBL: TOB; valeurs: T_Valeurs);
    	function AfficheSousDetail(TOBL: TOB): boolean;
      procedure GSAffectSoustraitPOC(Sender: tobject);
      procedure GSPaieSTPOC (Sender : TObject); 
    public
    	constructor create (laForm : Tform);
      destructor destroy; override;
      property Affaire : TOB read fTOBAffaire write Setaffaire;
      property TOBpiece : TOB read fTOBpiece write fTOBpiece;
      property TOBOuvrage:TOB read fTOBouvrage write fTOBouvrage;
      property TOBPieceTrait : TOB read ftobPieceTrait write ftobPieceTrait;
      property TOBPieceInterv : TOB read fTOBPieceInterv write fTOBPieceInterv;
      property TOBARticles : TOB read fTOBArticles write fTOBArticles;
      property TOBBases : TOB read fTOBBases  write fTOBbases;
      property InUse : boolean read fusable;
    	procedure SetGrilleSaisie(State: boolean);
    	procedure SetSaisie;
    	procedure ReinitSaisie;
      function FindSousTraitantPOC(Fournisseur, CodeMarche: string): TOB;
      procedure AddSousTraitant(Fournisseur, CodeMarche: string);
  end;

function  CotraitanceAffectationOk (CodeAffaire : string) : boolean;
procedure DecodeCleDoc (laChaine : string; var Cledoc : R_cledoc);
function  EncodeCleDoc (Cledoc : R_cledoc) : string; overload;
function  EncodeCleDoc (TOBL : TOB) : string; overload;
function  GestionMenuCotraitance (Numtag : integer) : integer;
function  GestionMenuSousTraitance (Numtag : integer) : integer;
function  GetLibIntervenant (TOBpiece : TOB; Arow : integer) : string; overload;
function  GetLibIntervenant (Fournisseur : string) : string; overload;
function  IsIntervenant (CodeAffaire,intervenant : string) : boolean;
function  IsLigneCotraite (TOBL : TOB) : boolean;
function  IsLigneSoustraite (TOBL : TOB) : boolean;
function  IsPieceGerableCoTraitance (TOBpiece,TOBaffaire : TOB) : boolean;
function  IsPieceGerableSousTraitance (TOBpiece : TOB) : boolean;
procedure LoadLaTOBPieceAffaireTrait (TOBpieceTrait: TOB;Faffaire : string; TypeInterv : String);
procedure LoadLaTOBPieceTrait(TOBpieceTrait : TOB; cledoc : r_Cledoc; TypeInterv : String);
procedure LoadLaTOBAffaireInterv(TOBAffaireInterv : TOB; Affaire : string);
procedure LoadLesSousTraitants(cledoc : R_CLEDOC;TOBSSTRAIT: TOB);
Function  IsPieceIntervenant(TOBT: TOB; Cledoc: R_Cledoc): Boolean;
procedure ReinitMontantPieceTrait(TOBPiece,TOBaffaire,TOBpieceTrait : TOB);
procedure ClearPieceTrait(TOBpieceTrait : TOB);
procedure SetTOBpieceTrait(TOBpiece,TOBAffaire,TOBpieceTrait: TOB);
procedure SetPieceCotraitant (TOBAffaire,TOBPiece,TOBTiers,TOBAdresses : TOB);
procedure SommePieceTrait (TOBpiece,TOBaffaire,TOBpieceTrait,TOBL : TOB;Sens:string='+');
procedure ValideLesSousTrait(TOBPiece,TOBSSTrait : TOB; DEV : rdevise);
procedure ValideLesPieceTrait(TOBPiece, TOBaffaire,TOBPieceTrait,TOBSousTrait : TOB; DEV : Rdevise);
procedure CalculeReglementsIntervenants(TOBSousTrait,TOBPiece,TOBPieceRG,TOBAcomptes,TOBPorcs : TOB; CodeAffaire : string;TOBPieceTrait : TOB; DEv : Rdevise; lastfacture : boolean=false);
function  ExistsCotrait(Fournisseur,Affaire : string ) : Boolean;
function  ExistsSoustrait (Fournisseur,NaturePiece,Souche : string ; Numero,Indiceg : integer) : Boolean;
function  PieceTraitUsable (fTOBpiece ,fTOBAffaire : TOB) : Boolean;
procedure InitTrait(TOBOUvrage,TOBL: TOB; ModeGestion : Tmodegestion);
procedure  InitTraitAvoir(TOBSSTrait,TOBOUvrage,TOBL: TOB; ModeGestion : Tmodegestion);

function  IsComptabilisationInterv (TOBPieceTrait : TOB) : boolean;
procedure AddChampsSupTrait (TOBSSTRAIT : TOB);
procedure AddChampsSupTraitCreat (TOBSSTRAIT : TOB);
function  GetModePaie (TOBPiece: TOB; Affaire,Fournisseur : string) : string;
function IsLigneExternalise (TOBL : TOB) : boolean;
procedure DefiniPieceTrait (TOBBases,TOBPiece,TOBAffaire,TOBPieceTrait,TOBSSTRAIT :TOB; DEV : Rdevise);
procedure GestionRGIntervenants (FF: Tform; TobPiece, TOBpieceTrait,TOBBases,TOBPorcs,TOBPieceRG,TOBBasesRG: TOB; DEV : RDevise);
function GetPaiementSSTrait (TOBSStrait : TOB; Intervenant : string) : string;
function IsMultiIntervRg (TOBpieceTrait : TOB) : boolean; overload;
FUNCTION FindRecapCotrait(Affaire : string) : boolean;

procedure GetReglementsPrec (Intervenant : string; TOBPiecesTraitPrec: TOB;EnPA : boolean;var MtDejaFacture,MtDejaRegle : double);
procedure  ConstitueCumuleAnterieur (TOBpiece,TOBPiecesTraitPrec : TOB);
procedure  ConstitueCumuleAnterieurAffaire (Affaire : string;TOBPiecesTraitPrec : TOB);

function GetMontantEntreprise (cledoc : R_CLEDOC) : double;
procedure Setallused (TOBSSTRAIT : TOB);
function ExisteReglDirect (TOBPieceTrait : TOB) : boolean;
function GetCumulPaiementDirect(TOBPieceTrait : TOB) : double;
function GetSqlPieceTrait : string;
function GetNbEcheances (TOBpieceTrait : TOB) : integer;
//procedure ConstitueEchesFromPieceTrait(Mode: T_ModeRegl; TOBPiece, TOBEches,TOBpieceTrait,TOBAcomptes : TOB; NbLigAcpte : integer);
function GetMontantEntreprisePDir (TOBpieceTrait : TOB) : double;
procedure AjoutePaiementDirectMode (var Mode: T_ModeRegl;TOBPiece,TOBTiers,TOBPieceTrait,TOBPieceRG,TOBAcomptes: TOB;DEV: RDevise);
procedure ReinitReglPieceTrait(TOBpiecetrait : TOB);
function GetLibelleFou( Code : string) : string;

implementation

uses facture,
     Entgc,
     AffaireUtil,
     FactTOB,
     FACTUREBTP,
     factUtil,
     factComm,
     FactCalc,
     UtilTOBpiece,
     LigNomen,
		 factVariante,
     Paramsoc,
     FactAdresse,
     UTOF_VideInside,
     factdomaines,
     UtilArticle,
     FactOuvrage,
     FactRG,
     BTPUtil,
     BTStructChampSup,
     TypInfo
     ,FactRetenues
     ,ULiquidTva2014
     ,UspecifPOC
     ;


function GetMontantEntreprise (cledoc : R_CLEDOC) : double;
var QQ : Tquery ;
begin
  Result := 0;
  QQ := OpenSQL('SELECT SUM(BPE_TOTALHTDEV) AS TOTALHT FROM PIECETRAIT WHERE '+WherePiece(cledoc,ttdPieceTrait ,True)+' AND ((BPE_FOURNISSEUR="") OR (BPE_TYPEINTERV="Y00"))',true,1,'',true);
  if not QQ.Eof then Result := QQ.findField('TOTALHT').AsFloat;
  ferme (QQ);
end;

function GetModeRegl (TOBSoustrait : TOB; Fournisseur : string) : string;
var TOBST : TOB;
begin
	result := '';
  TOBST := TOBSoustrait.FindFirst(['BPI_TIERSFOU'],[Fournisseur],true);
  if TOBST <> nil then result := TOBST.GetValue('BPI_TYPEPAIE');
end;


function GetLibelleFou( Code : string) : string;
var Sql : string;
    Q : TQuery;
begin
  Sql := 'SELECT T_LIBELLE FROM TIERS WHERE '+
         'T_NATUREAUXI="FOU" AND T_TIERS="'+Code+'"';
  Q := OpenSQL(Sql, True,-1, '', True);
  if not Q.eof then
  begin
    Result := Q.findField('T_LIBELLE').AsString;
  end;
  ferme (Q);
end;

procedure GestionRGIntervenants (FF: Tform; TobPiece, TOBpieceTrait,TOBBases,TOBPorcs,TOBPieceRG,TOBBasesRG: TOB; DEV : RDevise);
var indice : integer;
		TOBRG,TOBBRG,TOBDRG,TOBPT,UnePiece,LesBases,OneTOB : TOB;
    Interv : string;
    XX : TFFacture;
begin
  XX:= TFFacture(FF);
	TOBRG := TOB.Create ('LES RG',nil,-1);
	TOBBRG := TOB.Create ('LES BASES RG',nil,-1);
  UnePiece := TOB.Create ('PIECE',nil,-1);
  LesBases := TOB.Create ('LES BASES',nil,-1);
  //
  TOBPieceTrait.data := TOBRG;
  TOBRG.data := UnePiece;
  UnePiece.data := TOBBRG;
  TOBBRG.data := LesBases;
  LesBases.data := TOBPorcs;
  TRY
    UnePiece.Dupliquer(TOBPiece,true,true);
    LesBases.Dupliquer(TOBBases,true,true);
    TOBPieceTrait.detail.Sort('BPE_FOURNISSEUR'); // pour etre sur de traiter l'entreprise en premier
    if TOBPieceTrait.detail[0].getString('BPE_FOURNISSEUR') <> '' then
    begin
      TOBPT := TOB.Create ('PIECETRAIT',TOBPieceTrait,0);
      AddlesChampsSupPieceTrait (TOBPT);
      TOBPT.SetString ('LIBELLE',GetParamSocSecur('SO_LIBELLE','Notre Société'));
    end;
    //
    for Indice := 0 to TOBpieceTrait.detail.count -1 do
    begin
      TOBPT := TOBpieceTrait.detail[indice];
      Interv := '';
//      Interv := TOBPT.GetString('BPE_FOURNISSEUR');
//      if TOBPT.getString('TYPEPAIE')='002' then Interv := '';
      TOBDRG :=TOBPieceRG.findFirst(['PRG_FOURN'],[Interv],true);
      OneTOB := TOBRG.findFirst(['PRG_FOURN'],[Interv],true);
      if OneTOB = nil then
      begin
        OneTOB := TOB.create('PIECERG', TOBRG, -1);
        OneTOB.addchampsupValeur ('INDICERG',0);
        OneTOB.AddChampSupValeur ('CAUTIONUTIL',0);
        OneTOB.AddChampSupValeur ('CAUTIONUTILDEV',0);
        OneTOB.AddChampSupValeur ('CAUTIONAPRES',0);
        OneTOB.AddChampSupValeur ('CAUTIONAPRESDEV',0);
        if TOBDRG = nil then
        begin
          initLigneRg(OneTOB, TOBPIECE,Interv);
        end else
        begin
          OneTOB.dupliquer(TOBDRG,true,true);
        end;
      end;
    end;
    //
    TheTOB := TOBPieceTrait;
    AGLLanceFiche ('BTP','BTPIECERGCOT','','',ActionToString(XX.TheAction));
    if TheTOB <> nil then
    begin
      TOBpieceRG.clearDetail;
      TOBPieceRG.Dupliquer(TOBRG,true,true);
      TOBBasesRG.clearDetail;
      TOBBASesRG.dupliquer (TOBBRG,True,True);
    end;
  FINALLY
    TheTOB := nil;
    TOBRG.free;
    TOBBRG.free;
    UnePiece.free;
    LesBases.free;
  END;
end;


procedure DefiniPieceTrait (TOBBases,TOBPiece,TOBAffaire,TOBPieceTrait,TOBSSTRAIT :TOB; DEV : Rdevise);
var Indice : integer;
		Fournisseur,TypeInterv : string;
    TOBPT : TOB;
    X : double;
    TheModeregl : string;
begin
	if not PieceTraitUsable (TOBPiece,TOBaffaire) then exit;
  //
  for Indice := 0 to TOBPieceTrait.detail.count -1 do
  begin
  	TOBPieceTrait.detail[Indice].putValue('BPE_TOTALHTDEV',0);
  	TOBPieceTrait.detail[Indice].putValue('BPE_TOTALTTCDEV',0);
  	TOBPieceTrait.detail[Indice].putValue('BPE_MONTANTPA',0);
  	TOBPieceTrait.detail[Indice].putValue('BPE_MONTANTPR',0);
  	if TOBPieceTrait.detail[Indice].GetValue('BPE_MONTANTREGL')<0 then
    begin
			TOBPieceTrait.detail[Indice].PutValue('BPE_MONTANTREGL',0)
    end;
  end;

  For Indice := 0 to TOBBases.detail.count -1 do
  begin
		Fournisseur := TOBBases.Detail[Indice].GetValue('GPB_FOURN');
		TypeInterv := TOBBases.Detail[Indice].GetValue('GPB_TYPEINTERV');
    TheModeregl := GetModeRegl(TOBSSTRAIT, Fournisseur);
		TOBPT := TOBPieceTrait.findFirst(['BPE_FOURNISSEUR'],[Fournisseur],true);
    if TOBPT = nil then
    begin
      TOBPT := TOB.Create ('PIECETRAIT',TOBPieceTrait,-1);
      AddlesChampsSupPieceTrait (TOBPT);
      TOBPT.SetString('TYPEPAIE','');
      TOBPT.putValue('BPE_FOURNISSEUR',Fournisseur);
      TOBPT.putValue('BPE_TYPEINTERV',TypeInterv);
      if Fournisseur='' then
      begin
        TOBPT.SetString ('LIBELLE',GetParamSocSecur('SO_LIBELLE','Notre Société'));
      end else
      begin
        TOBPT.SetString ('LIBELLE',GetLibelleFou(Fournisseur));
      	TOBPT.SetString ('TYPEPAIE',GetModeRegl(TOBSSTRAIT, Fournisseur));
//      	TOBPT.SetString ('BPI_TYPEPAIE',TOBPT.GetString('TYPEPAIE'));
      end;
    end;
    X:=arrondi(TOBPT.GetValue('BPE_TOTALHTDEV')+TOBBases.Detail[Indice].GetValue('GPB_BASEDEV'),DEV.Decimale);
    TOBPT.PutValue('BPE_TOTALHTDEV',X) ;
    X:=arrondi(TOBPT.GetValue('BPE_TOTALTTCDEV')+TOBBases.Detail[Indice].GetValue('GPB_BASEDEV')+TOBBases.Detail[Indice].GetValue('GPB_VALEURDEV'),DEV.Decimale);
    TOBPT.PutValue('BPE_TOTALTTCDEV',X) ;
    X:=arrondi(TOBPT.GetValue('BPE_MONTANTPA')+TOBBases.Detail[Indice].GetValue('GPB_BASEACHAT')+TOBBases.Detail[Indice].GetValue('GPB_VALEURACHAT'),DEV.Decimale);
    TOBPT.PutValue('BPE_MONTANTPA',X) ;
  end;
end;

function CalculeTheFrais (MtFact : Double; TOBT : TOB; DEV : Rdevise) : double;
var MtFrais,CoefFrais : double;
begin
  MtFrais := TOBT.GetDouble('BAF_PVTTC');
  if MtFrais = 0 then
  begin
    CoefFrais := TOBT.GetDouble('BAF_COEFF');
    if CoefFrais <> 0 then MtFrais := Arrondi(MtFact * (CoefFrais/100),DEV.Decimale);
  end;
  Result := MtFrais;
end;

function ChargeEtCalculeLesFrais (CodeAffaire : string; TOBPieceT : TOB; DEV : Rdevise) : double;
var QQ : TQuery;
		Sql : string;
    Indice : Integer;
    TOBF : TOB;
    MtFrs,MtFrsTot : double;
begin
 	result := 0;
  if CodeAffaire = '' then Exit;
  if TOBPieceT.GetValue('BPE_FOURNISSEUR')='' then Exit;
  SQL := 'SELECT *, GPO_LIBELLE FROM AFFAIREFRSGEST '+
  			 'LEFT JOIN PORT ON BAF_CODEPORT=GPO_CODEPORT '+
         'WHERE BAF_AFFAIRE="' + CodeAffaire + '" ' +
  			 'AND BAF_TIERS="' + TOBPieceT.GetValue('BPE_FOURNISSEUR') + '"';
  QQ := OpenSQL(Sql,True,-1,'',True);
  TOBPieceT.LoadDetailDB('AFFAIREFRSGEST','','',QQ,false);
  ferme (QQ);
  //
  MtFrsTot := 0;
  for Indice := 0 to TOBPieceT.detail.count -1 do
  begin
    TOBF := TOBPieceT.detail[Indice];
    MtFrs := CalculeTheFrais (TOBPieceT.GetDouble('BPE_TOTALTTCDEV'),TOBF,DEV);
    MtFrsTot := MtFrsTot + MtFrs;
  end;
  Result := MtFrsTot;
end;


function ISSousTraitant (TOBSOUSTRAIT,TOBPT : TOB) : boolean;
begin
  result := false;
	if TOBPT.GetValue('BPE_TYPEINTERV')='Y00' then Result := true;
//	if GetModeRegl (TOBSOUSTRAIT,TOBPT.GetValue('BPE_FOURNISSEUR'))<>'002' then exit;
end;

procedure  ConstitueCumuleAnterieurAffaire (Affaire : string;TOBPiecesTraitPrec : TOB);
var QQ : TQuery;
		TOBSituations,TOBINTERM,TOBPP : TOB;
    Sql : string;
    Indice : integer;
begin
  TOBSituations := TOB.Create ('LES SITUATIONS',nil,-1);
  TOBINTERM := TOB.Create ('LES PIECETRAIT',nil,-1);
  TRY
    // recuperation des numeros de facture depuis les situations
    Sql := 'SELECT GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO FROM PIECE WHERE GP_AFFAIRE = "'+Affaire+'" AND GP_NATUREPIECEG="FBT"';
    QQ := OpenSql (Sql,true,-1,'',true);
    TOBSituations.LoadDetailDB('BSITUATIONS','','',QQ,false);
    ferme (QQ);
    if TOBSituations.detail.count = 0 then exit;
    for Indice := 0 to TOBSituations.detail.count -1 do
    begin
      Sql :='SELECT BPE_FOURNISSEUR,BPE_TYPEINTERV,BPE_MONTANTPA,BPE_TOTALTTCDEV,BPE_MONTANTREGL FROM PIECETRAIT WHERE '+
      			'BPE_NATUREPIECEG="'+TOBsituations.detail[Indice].getString('GP_NATUREPIECEG')+'" AND '+
      			'BPE_SOUCHE="'+TOBsituations.detail[Indice].getString('GP_SOUCHE')+'" AND '+
      			'BPE_NUMERO="'+TOBsituations.detail[Indice].getString('GP_NUMERO')+'" AND '+
      			'ORDER BY BPE_FOURNISSEUR';
      QQ := OpenSql (Sql,true,-1,'',true);
      if not QQ.eof then
      begin
				TOBInterm.LoadDetailDB('PIECETRAIT','','',QQ,true);
      end;
      ferme (QQ);
    end;
    Indice := 0;
    repeat
      TOBPP := TOBPiecesTraitPrec.findFirst(['INTERVENANT'],[TOBInterm.detail[Indice].GetSTring('BPE_FOURNISSEUR')],true);
      if TOBPP = nil then
      begin
        TOBPP := TOB.Create ('UN INTERVENANT',TOBPiecesTraitPrec,-1);
        TOBPP.AddChampSupValeur ('INTERVENANT',TOBInterm.detail[Indice].GetSTring('BPE_FOURNISSEUR'));
        TOBPP.AddChampSupValeur ('REGLE',0.0);
        TOBPP.AddChampSupValeur ('FACTUREPA',0.0);
        TOBPP.AddChampSupValeur ('FACTUREPV',0.0);
      end;
      TOBPP.SetDouble ('REGLE',TOBPP.GetDouble('REGLE')+TOBInterm.detail[Indice].GetDouble('BPE_MONTANTREGL'));
      TOBPP.SetDouble ('FACTUREPA',TOBPP.GetDouble('FACTUREPA')+TOBInterm.detail[Indice].GetDouble('BPE_MONTANTPA'));
      TOBPP.SetDouble ('FACTUREPV',TOBPP.GetDouble('FACTUREPV')+TOBInterm.detail[Indice].GetDouble('BPE_TOTALTTCDEV'));
      inc(indice);
    until Indice > TOBInterm.detail.count -1;
  FINALLY
  	TOBSituations.free;
  	TOBInterm.free;
  END;
end;

procedure  ConstitueCumuleAnterieur (TOBpiece,TOBPiecesTraitPrec : TOB);
var QQ : TQuery;
		TOBSituations,TOBINTERM,TOBPP : TOB;
    Sql : string;
    Indice : integer;
begin
  TOBSituations := TOB.Create ('LES SITUATIONS',nil,-1);
  TOBINTERM := TOB.Create ('LES PIECETRAIT',nil,-1);
  TRY
    // recuperation des numeros de facture depuis les situations
    Sql := 'SELECT BST_NATUREPIECE,BST_SOUCHE,BST_NUMEROFAC,BST_MONTANTTTC FROM BSITUATIONS WHERE '+
           'BST_SSAFFAIRE="'+TOBpiece.getString('GP_AFFAIREDEVIS')+'" AND '+
           'BST_VIVANTE="X" AND '+
           'BST_NUMEROFAC <> '+IntToStr(TOBPiece.GetInteger('GP_NUMERO'));
    QQ := OpenSql (Sql,true,-1,'',true);
    TOBSituations.LoadDetailDB('BSITUATIONS','','',QQ,false);
    ferme (QQ);
    if TOBSituations.detail.count = 0 then exit;
    for Indice := 0 to TOBSituations.detail.count -1 do
    begin
      Sql :='SELECT BPE_FOURNISSEUR,BPE_TYPEINTERV,BPE_MONTANTPA,BPE_TOTALTTCDEV,BPE_MONTANTREGL FROM PIECETRAIT WHERE '+
      			'BPE_NATUREPIECEG="'+TOBsituations.detail[Indice].getString('BST_NATUREPIECE')+'" AND '+
      			'BPE_SOUCHE="'+TOBsituations.detail[Indice].getString('BST_SOUCHE')+'" AND '+
      			'BPE_NUMERO="'+TOBsituations.detail[Indice].getString('BST_NUMEROFAC')+'" AND '+
      			'BPE_FOURNISSEUR<>"" ORDER BY BPE_FOURNISSEUR';
      QQ := OpenSql (Sql,true,-1,'',true);
      if not QQ.eof then
      begin
				TOBInterm.LoadDetailDB('PIECETRAIT','','',QQ,true);
      end else
      begin
        TOBPP := TOB.Create ('PIECETRAIT',TOBInterm,-1);
        TOBPP.SetString('BPE_FOURNISSEUR','');
        TOBPP.SetString('BPE_TYPEINTERV','');
        TOBPP.SetString('BPE_NATUREPIECEG',TOBsituations.detail[Indice].getString('BST_NATUREPIECE'));
        TOBPP.SetString('BPE_SOUCHE',TOBsituations.detail[Indice].getString('BST_SOUCHE'));
        TOBPP.SetString('BPE_NUMERO',TOBsituations.detail[Indice].getString('BST_NUMEROFAC'));
        TOBPP.SetDouble('BPE_TOTALTTCDEV',TOBsituations.detail[Indice].GeTDouble('BST_MONTANTTTC'));
        TOBPP.SetDouble('BPE_MONTANTREGL',TOBsituations.detail[Indice].GeTDouble('BST_MONTANTTTC'));
      end;
      ferme (QQ);
    end;
    Indice := 0;
    repeat
      TOBPP := TOBPiecesTraitPrec.findFirst(['INTERVENANT'],[TOBInterm.detail[Indice].GetSTring('BPE_FOURNISSEUR')],true);
      if TOBPP = nil then
      begin
        TOBPP := TOB.Create ('UN INTERVENANT',TOBPiecesTraitPrec,-1);
        TOBPP.AddChampSupValeur ('INTERVENANT',TOBInterm.detail[Indice].GetSTring('BPE_FOURNISSEUR'));
        TOBPP.AddChampSupValeur ('REGLE',0.0);
        TOBPP.AddChampSupValeur ('FACTUREPA',0.0);
        TOBPP.AddChampSupValeur ('FACTUREPV',0.0);
      end;
      TOBPP.SetDouble ('REGLE',TOBPP.GetDouble('REGLE')+TOBInterm.detail[Indice].GetDouble('BPE_MONTANTREGL'));
      TOBPP.SetDouble ('FACTUREPA',TOBPP.GetDouble('FACTUREPA')+TOBInterm.detail[Indice].GetDouble('BPE_MONTANTPA'));
      TOBPP.SetDouble ('FACTUREPV',TOBPP.GetDouble('FACTUREPV')+TOBInterm.detail[Indice].GetDouble('BPE_TOTALTTCDEV'));
      inc(indice);
    until Indice > TOBInterm.detail.count -1;
  FINALLY
  	TOBSituations.free;
  	TOBInterm.free;
  END;
end;

procedure GetReglementsPrec (Intervenant : string; TOBPiecesTraitPrec: TOB;EnPA : boolean;var MtDejaFacture,MtDejaRegle : double);
var TOBPT : TOB;
begin
  MtDejaFacture:= 0; MtDejaRegle := 0;
	TOBPT := TOBPiecesTraitPrec.findFirst(['INTERVENANT'],[Intervenant],true);
  if TOBPT = nil then exit;
  MtDejaRegle := TOBPT.GetDouble('REGLE');
  if EnPA then
  begin
    MtDejaFacture := TOBPT.GetDouble('FACTUREPA');
  end else
  begin
    MtDejaFacture := TOBPT.GetDouble('FACTUREPV');
  end;
end;

procedure CalculeReglementsIntervenants(TOBSousTrait,TOBPiece,TOBPieceRG,TOBAcomptes,TOBPorcs : TOB; CodeAffaire : string;TOBPieceTrait : TOB; DEv : Rdevise; lastfacture : boolean=false);
var TOBPT         : TOB;
    TOBMANDATAIRE : TOB;
    TOBPiecesTraitPrec : TOB;
		Indice        : Integer;
    Xp,AP,AD,RXP,RXD  : Double;
    Xd            : Double;
    MtRegle,Mtreglable : Double;
    MtFrs         : Double;
    MtFrsCoTrait  : double;
    DiffReglement : double;
    MtDejaRegle   : double;
    MtDejaFacture : double;
    NumCaution    : String;
    EcartRegl,RGSt,SommePaie,SommeAPaie,NetApayer,MtAcomptesSoc : double;
begin

  if TOBPieceTrait.detail.Count = 0 then Exit;
//  if TOBPieceTrait.detail[0].GetBoolean('BPE_REGLSAISIE')= true then exit; // evite de recalculer les reglemets si deja saisi
  TOBPiecesTraitPrec := TOB.Create ('LES PIECES TRAITS',nil,-1);
  MtFrsCoTrait := 0;
  MtDejaRegle   := 0;
  MtDejaFacture := 0;
  EcartRegl := 0;
  RGSt := 0;
  SommePaie := 0;
  SommeAPaie := 0;
  ConstitueCumuleAnterieur (TOBpiece,TOBPiecesTraitPrec);
  for Indice := 0 to TOBPieceTrait.detail.count -1 do
  begin
  	EcartRegl := 0;
    MtRegle := 0;
    MtFrs := 0;
    EcartRegl := 0;
  	TOBPT := TOBPieceTrait.detail[Indice];
    if TOBPT.getValue('BPE_TYPEINTERV')='X01' then
    begin
  		MtFrs := ChargeEtCalculeLesFrais(CodeAffaire,TOBPT,DEV);
    end else
    begin
      MtFrs := 0;
    end;
    TOBPT.ClearDetail;
    //Calcul montant retenue garantie
    GetRg(TOBPieceRG, false, true, Xp, Xd, NumCaution,TOBPT.getString('BPE_FOURNISSEUR'),False);
    //chargement du montant de réglement si paiement directe
    if not ISSousTraitant (TOBSousTrait,TOBPT) then
    begin
      if TOBPT.getString('BPE_FOURNISSEUR')<>'' then
      begin
        if lastfacture then
        begin
          GetReglementsPrec (TOBPT.getString('BPE_FOURNISSEUR'),TOBPiecesTraitPrec,false,MtDejaFacture,MtDejaRegle);
          EcartRegl := Arrondi(MtDejafacture - MtDejaRegle,Dev.decimale);
        end;
        MtRegle := TOBPT.GetDouble('BPE_TOTALTTCDEV');
        MtRegle := MtRegle - MtFrs;
        MtRegle := MtRegle - Xd;   // Deduction de la RG
        MtRegle := Mtregle + EcartRegl;
        MtRegle := ARRONDI(MtRegle,DEv.Decimale);
        GetSommeReglement(TOBAcomptes, AP, AD,'',false);
        if (Ap = 0) and ( not TOBPT.GetBoolean('BPE_REGLSAISIE')) then
        begin
        	TOBPT.putvalue('BPE_MONTANTREGL', MtRegle);
        end;
        TOBPT.PutValue('MONTANTREGLABLE',  MtRegle);
        TOBPT.PutValue('MONTANTFACT',      TOBPT.GetDouble('BPE_TOTALTTCDEV'));
        if lastfacture then
        begin
          TOBPT.putvalue('BPE_REGLSAISIE','X');
        end;
      end;
    end else // reglement d'un sous traitant
    begin
      if lastfacture then
      begin
        GetReglementsPrec (TOBPT.getString('BPE_FOURNISSEUR'),TOBPiecesTraitPrec,true,MtDejaFacture,MtDejaRegle);
        EcartRegl := Arrondi(MtDejafacture - MtDejaRegle,Dev.decimale);
      end;
      //
      RGSt := RGSt + Xd; // Pour ajouter plus tard la Rg de sous traitance calculé sur l'entreprise
      									 // Normal puisque c'est l'entreprise qui a la RG vis a vis du client
      //
      if TOBPT.GetString('TYPEPAIE')='002' then
      begin
        Mtreglable := 0;
      end else
      begin
        Mtreglable := TOBPT.GetValue('BPE_MONTANTPA');
        if not TOBPT.GetBoolean('BPE_REGLSAISIE') then
        begin
          Mtregle := TOBPT.GetValue('BPE_MONTANTPA');
        end else
        begin
          Mtregle := TOBPT.GetValue('BPE_MONTANTREGL');
        end;
        GetSommeReglement(TOBAcomptes, AP, AD,'',false);
        //
      end;
      Mtreglable := Mtreglable + EcartRegl;
      Mtreglable := ARRONDI(Mtreglable,DEv.Decimale);
      TOBPT.PutValue('MONTANTREGLABLE',  Mtreglable);
      //
      if (Ap = 0) and ( not TOBPT.GetBoolean('BPE_REGLSAISIE')) then
      begin
    		TOBPT.putvalue('BPE_MONTANTREGL', Mtreglable);
        SommeAPaie := SommeAPaie + MtReglable;
      end;
      TOBPT.PutValue('MONTANTFACT', TOBPT.GetValue('BPE_MONTANTPA'));
      //
      if lastFacture then
      begin
    		TOBPT.putvalue('BPE_REGLSAISIE','X');
      end;
			//end;
    end;
    //
    Sommepaie := SommePaie + MtRegle;
    MtFrsCoTrait := MtFrsCoTrait + MtFrs;
    DiffReglement := DiffReglement + EcartRegl;
  end;
  // --
  SommePaie := Arrondi(Sommepaie,Dev.decimale);
  MtFrsCoTrait := Arrondi(MtFrsCoTrait,Dev.decimale);
  DiffReglement := Arrondi(DiffReglement,Dev.decimale);
  //
  TOBMANDATAIRE := TOBPieceTrait.FindFirst(['BPE_FOURNISSEUR'],[''],true);
  if TOBMANDATAIRE = nil then
  begin
    TOBMANDATAIRE := TOB.Create('PIECETRAIT',TOBPieceTrait,-1);
    AddlesChampsSupPieceTrait (TOBMANDATAIRE);
    TOBMANDATAIRE.PutValue('BPE_FOURNISSEUR','');
  	TOBMANDATAIRE.PutValue('LIBELLE', GetParamSocSecur('SO_LIBELLE','Notre Société'));
    TOBMANDATAIRE.PutValue('BPE_TOTALHTDEV',0);
    TOBMANDATAIRE.PutValue('BPE_TOTALTTCDEV',0);
    TOBMANDATAIRE.PutValue('BPE_MONTANTPA',0);
    TOBMANDATAIRE.PutValue('BPE_MONTANTPR',0);
    TOBMANDATAIRE.PutValue('BPE_MONTANTFAC',0);
  end;
  //
  GetMontantRGReliquat(TOBPIeceRG, XD, XP,true);
  GetRetenuesCollectif (TOBPorcs,RXP,RXD);
  MtAcomptesSoc := GetMontantEntreprisePDir (TOBPieceTrait);
  //
  NetApayer := TOBPiece.GetValue('GP_TOTALTTCDEV') - XD - RXD;
  TOBMANDATAIRE.putvalue('MONTANTREGLABLE',Arrondi(NetAPayer - SommePaie ,DEV.Decimale));
  if ( not TOBMANDATAIRE.GetBoolean('BPE_REGLSAISIE')) then
  begin
  	TOBMANDATAIRE.putvalue('BPE_MONTANTREGL',TOBMANDATAIRE.GetDouble('MONTANTREGLABLE'));
  end;
  TOBMANDATAIRE.PutValue('MONTANTFACT',TOBMANDATAIRE.GetDouble('BPE_TOTALTTCDEV'));
  TOBMANDATAIRE.PutValue('LIBELLE', GetParamSocSecur('SO_LIBELLE','Notre Société'));
  TOBMANDATAIRE.PutValue('TOTALFRAIS', MtFrsCoTrait);

  if lastfacture then
  begin
    TOBMANDATAIRE.putvalue('BPE_REGLSAISIE','X');
  end;
  TOBPieceTrait.Detail.Sort('BPE_FOURNISSEUR'); 
  TOBPiecesTraitPrec.free;
end;

procedure LoadLaTOBAffaireInterv(TOBAffaireInterv : TOB; Affaire : string);
var Q : Tquery;
		req : string;
begin
	TOBAffaireInterv.clearDetail;
  if Affaire = '' then exit;
  Req := 'SELECT * FROM AFFAIREINTERV '+
         'WHERE BAI_AFFAIRE="'+Affaire+'"';
  Q := OpenSql (Req, True,-1, '', True);
  if not Q.eof then TOBAffaireInterv.LoadDetailDB('AFFAIREINTERV', '', '', Q, False);
  Ferme(Q);
end;

procedure LoadLaTOBPieceAffaireTrait (TOBpieceTrait: TOB;Faffaire : string; TypeInterv : String);
var Q       : Tquery;
		req     : string;
		Indice  : integer;
    Where   : string;
    Group   : string;
    StSQL   : string;
begin
	TOBPieceTrait.clearDetail;
  //
  StSQL := 'SELECT BPE_FOURNISSEUR,BPE_TYPEINTERV,SUM(BPE_TOTALHTDEV) AS BPE_TOTALHTDEV ,'+
           'SUM(BPE_TOTALTTCDEV) AS BPE_TOTALTTCDEV, SUM(BPE_MONTANTPR) AS BPE_MONTANTPR,'+
           'SUM(BPE_MONTANTFAC) AS BPE_MONTANTFAC, SUM(BPE_MONTANTPA) AS BPE_MONTANTPA,'+
           '(SELECT T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="FOU" AND T_TIERS=BPE_FOURNISSEUR) AS LIBELLE,'+
           'AFF_TYPEPAIE AS TYPEPAIE,'+
           '0 AS FACTURE, 0 AS REGLE '+
           'FROM PIECETRAIT '+
           'LEFT JOIN AFFAIRE ON AFF_AFFAIRE=BPE_AFFAIRE ';
  //
  Where := 'WHERE BPE_AFFAIRE="'+Faffaire+'" AND BPE_NATUREPIECEG IN ("DBT", "ETU")';

  if TypeInterv = 'COTRAITE' then
    Where := Where + ' AND BPE_TYPEINTERV IN ("X00", "X01","")'
  else if TypeInterv = 'SSTRAITE' then
    Where := Where + ' AND BPE_TYPEINTERV IN ("Y00","")';

  Group :=  'GROUP BY BPE_FOURNISSEUR,AFF_TYPEPAIE,BPE_TYPEINTERV';

  req := StSQL + Where + Group;

  Q := OpenSql (Req, True,-1, '', True);

  if not Q.eof then TOBPieceTrait.LoadDetailDB('AFFAIRE AFFECT', '', '', Q, False);

  for Indice := 0 to TOBpieceTrait.detail.count -1 do
  begin
  	if TOBpieceTrait.detail[Indice].getValue('BPE_FOURNISSEUR')='' then
    begin
      TOBpieceTrait.detail[Indice].puTvalue('LIBELLE', GetParamSocSecur('SO_LIBELLE','Notre Société'));
      break;
    end;
  end;

  Ferme(Q);

end;

procedure LoadLesSousTraitants(cledoc : R_CLEDOC;TOBSSTRAIT: TOB);
var Q       : Tquery;
		Indice  : integer;
    Sql     : string;
    where   : string;
begin

  Sql := 'SELECT *,'+
  			 '(SELECT T_LIBELLE FROM TIERS WHERE T_TIERS=BPI_TIERSFOU AND T_NATUREAUXI="FOU") AS TLIBTIERSFOU,'+
         '(SELECT R_DOMICILIATION FROM RIB WHERE R_AUXILIAIRE=BPI_TIERSFOU AND R_NUMERORIB=BPI_NUMERORIB) AS TLIB_NUMERORIB,' +
  			 '(SELECT C_NOM FROM CONTACT WHERE C_TYPECONTACT="T" AND C_NATUREAUXI="FOU" AND '+
         'C_AUXILIAIRE=(SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS=BPI_TIERSFOU AND T_NATUREAUXI="FOU") '+
         'AND C_NUMEROCONTACT=BPI_NUMEROCONTACT)  AS TLIBCONTACT '+
         'FROM PIECEINTERV ';
  //
  Where := 'WHERE ' + WherePiece(CleDoc, ttdPieceInterv, False);
  Q := OpenSQL(Sql + where, True,-1, '', True);

  TOBSSTRAIT.LoadDetailDB('PIECEINTERV', '', '', Q, False);

  for Indice := 0 to TOBSSTRAIT.detail.count -1 do
  begin
  	AddChampsSupTrait (TOBSSTRAIT.detail[Indice]);
  end;
  Ferme(Q);
end;

procedure AddChampsSupTrait (TOBSSTRAIT : TOB);
begin
	TOBSSTRAIT.AddChampSupValeur('UTILISE','-');
  TOBssTrait.AddChampSupValeur('ETABBQ', '');
  TOBssTrait.AddChampSupValeur('GUICHET', '');
  TOBssTrait.AddChampSupValeur('NUMEROCOMPTE', '');
  TOBssTrait.AddChampSupValeur('CLERIB', '');
  TOBssTrait.AddChampSupValeur('IBAN', '');
  TOBssTrait.AddChampSupValeur('AFFAIRE', '');
  TOBssTrait.AddChampSupValeur('NBUSED', 0);
end;

procedure AddChampsSupTraitCreat (TOBSSTRAIT : TOB);
begin
	TOBSSTRAIT.AddChampSupValeur('TLIBTIERSFOU','');
  TOBSSTRAIT.AddChampSupValeur('TLIB_NUMERORIB','');
  TOBSSTRAIT.AddChampSupValeur('TLIBCONTACT','');
end;

function GetSqlPieceTrait : string;
begin
  result := 'SELECT PIECETRAIT.*, T_LIBELLE AS LIBELLE, BPI_TYPEPAIE AS TYPEPAIE, 0 AS FACTURE, '+
  					'0 AS FACTURAPA, 0 AS FACTURPEV , 0 AS REGLE , '+
  					'0 AS TOTALFRAIS , 0 AS MONTANTREGLABLE, 0 AS MONTANTFACT, '+
            '" " AS FIXED '+
            'FROM PIECETRAIT '+
            'LEFT JOIN TIERS ON T_NATUREAUXI="FOU" AND T_TIERS=BPE_FOURNISSEUR '+
            'LEFT JOIN PIECEINTERV ON BPI_NATUREPIECEG=BPE_NATUREPIECEG AND '+
            'BPI_SOUCHE=BPE_SOUCHE AND BPI_NUMERO=BPE_NUMERO AND BPI_INDICEG=BPE_INDICEG AND '+
            'BPI_TIERSFOU=BPE_FOURNISSEUR ';
end;

procedure LoadLaTOBPieceTrait(TOBpieceTrait : TOB; cledoc : r_Cledoc; TypeInterv : String);
var Q       : Tquery;
		Indice  : integer;
    Sql     : string;
    where   : string;
begin
  Sql := GetSqlPieceTrait;
  //
  Where := ' WHERE ' + WherePiece(CleDoc, ttdpieceTrait, False);

  if TypeInterv = 'COTRAITE' then
    Where := Where + ' AND BPE_TYPEINTERV IN ("X00", "X01")'
  else if TypeInterv = 'SSTRAITE' then
    Where := Where + ' AND BPE_TYPEINTERV="Y00"';

  Q := OpenSQL(Sql + where, True,-1, '', True);

  TOBPieceTrait.LoadDetailDB('PIECETRAIT', '', '', Q, False);

  for Indice := 0 to TOBpieceTrait.detail.count -1 do
  begin
  	if TOBpieceTrait.detail[Indice].getValue('BPE_FOURNISSEUR')='' then
    begin
//      TOBpieceTrait.detail[Indice].puTvalue('LIBELLE','Notre Société');
      TOBpieceTrait.detail[Indice].puTvalue('LIBELLE',GetParamSocSecur('SO_LIBELLE','Notre Société'));
      break;
    end;
  end;

  Ferme(Q);

end;

procedure DecodeCleDoc (laChaine : string; var Cledoc : R_cledoc);
var TheChaine : string;
begin
	TheChaine := laChaine;
  if TheChaine ='' then exit;
  FillChar (Cledoc,Sizeof(Cledoc),0);
  Cledoc.NaturePiece := READTOKENPipe(TheChaine,':');
  if TheChaine = '' then exit;
  Cledoc.Souche := READTOKENPipe(TheChaine,':');
  if TheChaine = '' then exit;
  Cledoc.NumeroPiece  := StrToInt(READTOKENPipe(TheChaine,':'));
  if TheChaine = '' then exit;
  Cledoc.Indice   := StrToInt(READTOKENPipe(TheChaine,':'));
end;

function EncodeCleDoc (Cledoc : R_cledoc) : string;
begin
	result := Cledoc.NaturePiece+':'+Cledoc.Souche+':'+inttoStr(Cledoc.NumeroPiece)+':'+IntTostr(Cledoc.Indice);
end;

function EncodeCleDoc (TOBL : TOB) : string;
var prefixe : string;
begin
	if Pos(TOBL.NomTable,'LIGNE;LIGNEOUV;PIECE;LIGNEOUVPLAT') = 0 then exit;
  prefixe := GetPrefixeTable (TOBL);
  result := TOBL.getValue(prefixe+'_NATUREPIECEG')+':'+
  					TOBL.getValue(prefixe+'_SOUCHE')+':'+
  					IntToStr(TOBL.getValue(prefixe+'_NUMERO'))+':'+
  					IntToStr(TOBL.getValue(prefixe+'_INDICEG'));
end;

function IsPieceGerableCoTraitance (TOBpiece,TOBaffaire : TOB) : boolean;
begin
	result := false;
  if pos (TOBPiece.GetString('GP_NATUREPIECEG'),'ETU;DBT;FBT;FBP;ABP;')= 0 then exit;
  if TOBPiece.getValue('GP_AFFAIRE')='' then exit;
  //
  if (VH_GC.SeriaCoTraitance) and
  	 (TOBaffaire.getValue ('AFF_MANDATAIRE') = 'X') and
     (TOBpiece.getValue('GP_VENTEACHAT')='VEN') then result := true;
end;

function IsPieceGerableSousTraitance (TOBpiece: TOB) : boolean;
begin
	result := false;
  if pos (TOBPiece.GetString('GP_NATUREPIECEG'),'ETU;BCE;DBT;FBT;PBT;ABT;FBP;ABP;')= 0 then exit;
  if (VH_GC.SeriaSousTraitance) and
     (TOBpiece.getValue('GP_VENTEACHAT')='VEN') then result := true;
end;


procedure SetPieceCotraitant (TOBAffaire,TOBPiece,TOBTiers,TOBAdresses : TOB);

	procedure  SetAdresseFac (TOBPiece,TOBAdresses : TOB);
  begin

  end;

var QQ : TQuery;
		TOBAFFINT,TOBCLI : TOB;
		Sql,CodeFact : string;
begin

  if (TOBAffaire.GetString('AFF_AFFAIRE0') = 'I')
  Or (TOBAffaire.GetString('AFF_AFFAIRE0') = 'W') then exit;

  if TOBAffaire.getValue('AFF_AFFAIRE')='' then Exit;

	// positionnement du tiers facture, tiers payeur (en fonction du type de paiment du groupement)
  //      ''        de l'adresse de facturation
  //      ''        du rib
  TOBAFFINT := TOB.Create ('AFFAIREINTERV',nil,-1);
  TOBCLI := TOB.Create('TIERS',nil,-1);
  TRY
    Sql := 'SELECT BAI_TIERSCLI,BAI_NUMERORIB FROM AFFAIREINTERV WHERE '+
           'BAI_AFFAIRE="'+TOBAFFAire.GetValue('AFF_AFFAIRE')+'" AND '+
           'BAI_TYPEINTERV="X00"';

    QQ := OpenSQL(Sql,True,1,'',true);
    if not QQ.eof then
    begin
      TOBAFFINT.SelectDB('',QQ);
    end else
    begin
      PGIInfo('Vous n''avez pas renseigné les informations du groupement pour la cotraitance');
    end;
    Ferme(QQ);

    if TOBAFFINT.GetValue('BAI_TIERSCLI')<> '' then
    begin
      CodeFact := TOBAFFINT.GetValue('BAI_TIERSCLI');
      TOBPiece.PutValue('GP_TIERSFACTURE', CodeFact);
      GetAdrFromCode(TOBAdresses.Detail[1], CodeFact);
      if TOBAffaire.GetValue('AFF_TYPEPAIE') = '000' then
      begin
        // Compte de groupement
        Sql := 'SELECT T_TIERS,T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+CodeFact+'" '+
        			 'AND T_NATUREAUXI="CLI"';
        QQ := OpenSQL(Sql,True,1,'',true);
        TRY
          if not QQ.Eof then
          begin
            TOBCLI.SelectDB('',QQ);
            TOBPiece.PutValue('GP_RIB', GetRIBParticulier ( TOBCLI.GetValue('T_AUXILIAIRE'), TOBAFFINT.GetValue('BAI_NUMERORIB')) );
            TOBPiece.PutValue('GP_TIERSPAYEUR', TOBCLI.GetValue('T_TIERS'));
          end;
        FINALLY
        	ferme (QQ);
        end;
      end;
    end;
  FINALLY
    TOBAFFINT.free;
    TOBCLI.Free;
  end;
end;

procedure SetTOBpieceTrait(TOBpiece,TOBaffaire,TOBpieceTrait: TOB);
begin
	if not PieceTraitUsable (TOBpiece,TOBaffaire) then exit;
	SommePieceTrait (TOBpiece,TOBaffaire,TOBpieceTrait,TOBpiece);
end;

procedure ReinitMontantPieceTrait(TOBPiece,TOBaffaire,TOBpieceTrait : TOB);
var TOBP : TOB;
	Indice : integer;
begin
	if not PieceTraitUsable (TOBPiece,TOBaffaire) then exit;
	For Indice := 0 to TOBpieceTrait.detail.count -1 do
  begin
  	TOBP := TOBpieceTrait.detail[Indice];
    TOBP.PutValue('BPE_TOTALHTDEV',0) ;
    TOBP.PutValue('BPE_TOTALTTCDEV',0) ;
    TOBP.PutValue('BPE_MONTANTPA',0) ;
    TOBP.PutValue('BPE_MONTANTPR',0) ;
  end;
end;

procedure SommePieceTrait (TOBPiece,TOBaffaire,TOBpieceTrait,TOBL : TOB;Sens:string='+');
var Fournisseur : string;
		TOBPT : TOB;
    prefixe,TypeInterv : string;
    NatureTravail : double;
    X,MTTaxe : double;
    DEV : RDevise;
begin
	if not PieceTraitUsable (TOBPiece,TOBaffaire) then exit;
  if TOBpieceTrait = nil then Exit;
  prefixe := GetPrefixeTable(TOBL);
	Fournisseur := '';
  NatureTravail := 0;
  TypeInterv := '';
  DEV.Code := TOBPiece.getvalue('GP_DEVISE');
	GetInfosDevise(DEV) ;
  if TOBL.NomTable = 'LIGNE' then
  begin
  	if TOBL.getvalue('GLC_NATURETRAVAIL')<>'' then
    begin
      NatureTravail := valeur(TOBL.getvalue('GLC_NATURETRAVAIL'));
    	Fournisseur := TOBL.getvalue('GL_FOURNISSEUR');
    end;
  end else if TOBL.Nomtable = 'LIGNEOUV' then
  begin
  	if TOBL.getvalue('BLO_NATURETRAVAIL')<>'' then
    begin
      NatureTravail := valeur(TOBL.getvalue('BLO_NATURETRAVAIL'));
    	Fournisseur := TOBL.getvalue('BLO_FOURNISSEUR');
    end;
  end else if TOBL.Nomtable = 'LIGNEOUVPLAT' then
  begin
  	if TOBL.getvalue('BOP_NATURETRAVAIL')<>'' then
    begin
      NatureTravail := valeur(TOBL.getvalue('BOP_NATURETRAVAIL'));
    	Fournisseur := TOBL.getvalue('BOP_FOURNISSEUR');
    end;
  end;
  TOBPT := TOBpieceTrait.findFirst(['BPE_FOURNISSEUR'],[Fournisseur],true);
  if TOBPT = nil then
  begin
  	TOBPT := TOB.Create ('PIECETRAIT',TOBpieceTrait,-1);
    //
    AddlesChampsSupPieceTrait (TOBPT);
    //
    if NatureTravail = 1 then TypeInterv := 'X01'
    else if NatureTravail = 2  then TypeInterv := 'Y00';
    TOBPT.putvalue('BPE_TYPEINTERV',TypeInterv);
    TOBPT.putvalue('BPE_FOURNISSEUR',Fournisseur);
  	if Fournisseur='' then
    begin
      TOBPT.SetString ('LIBELLE',GetParamSocSecur('SO_LIBELLE','Notre Société'));
    end else
    begin
      TOBPT.SetString ('LIBELLE',GetLibelleFou(Fournisseur));
    end;
  end;
  if Sens = '+' then X:=arrondi(TOBPT.GetValue('BPE_TOTALHTDEV')+TOBL.GetValue(prefixe+'_TOTALHTDEV'),DEV.Decimale)
                else X:=arrondi(TOBPT.GetValue('BPE_TOTALHTDEV')-TOBL.GetValue(prefixe+'_TOTALHTDEV'),DEV.decimale);
  TOBPT.PutValue('BPE_TOTALHTDEV',X) ;
  if Sens = '+' then X:=arrondi(TOBPT.GetValue('BPE_TOTALTTCDEV')+TOBL.GetValue(prefixe+'_TOTALTTCDEV'),DEV.Decimale )
                else X:=arrondi(TOBPT.GetValue('BPE_TOTALTTCDEV')-TOBL.GetValue(prefixe+'_TOTALTTCDEV'),DEV.decimale);
  TOBPT.PutValue('BPE_TOTALTTCDEV',X) ;
  if TOBPT.GetValue('BPE_TOTALHTDEV')= 0 then TOBPT.PutValue('BPE_TOTALTTCDEV',0);
  // -- Gestion des achats TTC (Sous traitance)
  if Prefixe <> 'GP' then
  begin
    MTTaxe := GetMontantTaxeLigne (TOBL,TOBL.GetValue(prefixe+'_MONTANTPA'),DEV);
    if Sens = '+' then X:=arrondi(TOBPT.GetValue('BPE_MONTANTPA')+TOBL.GetValue(prefixe+'_MONTANTPA')+MTTaxe,DEV.Decimale)
                  else X:=arrondi(TOBPT.GetValue('BPE_MONTANTPA')-TOBL.GetValue(prefixe+'_MONTANTPA')-MTTaxe,DEV.decimale);
    TOBPT.PutValue('BPE_MONTANTPA',X) ;
  // ---
    if Sens = '+' then X:=arrondi(TOBPT.GetValue('BPE_MONTANTPR')+TOBL.GetValue(prefixe+'_MONTANTPR'),DEV.DEcimale)
                  else X:=arrondi(TOBPT.GetValue('BPE_MONTANTPR')-TOBL.GetValue(prefixe+'_MONTANTPR'),DEV.DEcimale);
    TOBPT.PutValue('BPE_MONTANTPR',X) ;
    if TOBPT.GetValue('BPE_TOTALHTDEV')= 0 then TOBPT.PutValue('BPE_MONTANTPA',0);
  end;
end;

function GetLibIntervenant (TOBpiece : TOB; Arow : integer) : string;
var TOBL : TOB;
		QQ : TQuery;
begin
	TOBL := TOBPiece.detail[Arow-1];
  if TOBL.GetValue('GL_FOURNISSEUR')<> '' then
  begin
    QQ := OpenSql('SELECT T_LIBELLE FROM TIERS WHERE T_TIERS="'+TOBL.getvalue('GL_FOURNISSEUR')+'" AND T_NATUREAUXI="FOU"',
    							true,1,'',true);
    if not QQ.eof then
    begin
      result := QQ.findField('T_LIBELLE').AsString;
    end;
    ferme (QQ);
  end else
  begin
  	result := GetParamSocSecur('SO_LIBELLE','Notre Société');
  end;
end;

function GetLibIntervenant (Fournisseur : string) : string;
var QQ : TQuery;
begin
	result := '';
	if Fournisseur <> '' then
  begin
    QQ := OpenSql('SELECT T_LIBELLE FROM TIERS WHERE T_TIERS="'+Fournisseur+'" AND T_NATUREAUXI="FOU"',
    							true,1,'',true);
    if not QQ.eof then
    begin
      result := QQ.findField('T_LIBELLE').AsString;
    end;
    ferme (QQ);
  end else
  begin
  	result := GetParamSocSecur('SO_LIBELLE','Notre Société');
  end;
end;

function IsLigneExternalise (TOBL : TOB) : boolean;
var NatureTravail : double;
begin
  NatureTravail := 0;
  if TOBL.NomTable = 'LIGNE' then
  begin
 		NatureTravail := valeur(TOBL.getValue('GLC_NATURETRAVAIL'));
  end else if TOBL.NomTable = 'LIGNEOUV' then
  begin
 		NatureTravail := valeur(TOBL.getValue('BLO_NATURETRAVAIL'));
  end else if TOBL.NomTable = 'LIGNEOUVPLAT' then
  begin
 		NatureTravail := valeur(TOBL.getValue('BOP_NATURETRAVAIL'));
  end;
  result := (NatureTravail >= 1) and (NatureTravail < 10);
end;

function  IsLigneSoustraite (TOBL : TOB) : boolean;
begin
	result := false;
  if TOBL.NomTable = 'LIGNE' then
  begin
		if Pos(TOBL.getValue('GLC_NATURETRAVAIL'),'002')> 0 then result := true;
  end else if TOBL.NomTable = 'LIGNEOUV' then
  begin
		if Pos(TOBL.getValue('BLO_NATURETRAVAIL'),'002')> 0 then result := true;
  end else if TOBL.NomTable = 'LIGNEOUVPLAT' then
  begin
		if Pos(TOBL.getValue('BOP_NATURETRAVAIL'),'002')> 0 then result := true;
  end else if TOBL.NomTable = 'LIGNEFAC' then
  begin
		if Pos(TOBL.getValue('BLF_NATURETRAVAIL'),'002')> 0 then result := true;
  end;
end;

function IsLigneCotraite (TOBL : TOB) : boolean;
begin
	result := false;
  if TOBL.NomTable = 'LIGNE' then
  begin
		if Pos(TOBL.getValue('GLC_NATURETRAVAIL'),'001;011')> 0 then result := true;
  end else if TOBL.NomTable = 'LIGNEOUV' then
  begin
		if Pos(TOBL.getValue('BLO_NATURETRAVAIL'),'001;011')> 0 then result := true;
  end else if TOBL.NomTable = 'LIGNEOUVPLAT' then
  begin
		if Pos(TOBL.getValue('BOP_NATURETRAVAIL'),'001;011')> 0 then result := true;
  end;
end;

function CotraitanceAffectationOk (CodeAffaire : string) : boolean;
var QQ : TQuery;
begin
	QQ := OpenSql ('SELECT BAI_TYPEINTERV FROM AFFAIREINTERV WHERE BAI_AFFAIRE="'+CodeAffaire+'"',true,
  							 -1,'',true);
	result := not QQ.eof;
  ferme (QQ);
end;

function GestionMenuCotraitance (Numtag : integer) : integer;
begin
	if not VH_GC.SeriaCoTraitance then
  begin
  	PgiBox('Ce module n''est pas sérialisé. Il ne peut être utilisé.','CoTraitance');
    result := -1;
    exit;
  end;
  result := 0;
  case Numtag of
    325111 : AGLLanceFiche('BTP','BTAFFAIRE_MUL','AFF_AFFAIRE=P','','COTRAITANCE; ETAT=ENC; STATUT=PRO');  //Selection Appel d'offre
    325112 : AGLLanceFiche('BTP','BTAPPOFF_MUL' ,'AFF_AFFAIRE=P','','COTRAITANCE; ETAT=ENC; STATUT=PRO; APPOFFACCEPT');  //Acceptation Appel d'offre
    325120 : AGLLanceFiche('BTP','BTAFFAIRE_MUL','AFF_AFFAIRE=A','','COTRAITANCE; ETAT=ENC; STATUT=AFF');  //Selection affaire
    325130 : AGLLanceFiche('BTP','BTINTEGREDOC','','','TYPEIMPORT=SST');                                                  //intégration document XLS cotraitant
    325140 : AGLLanceFiche('BTP','BTDEVIS_MUL', 'GP_NATUREPIECEG=DBT;GP_VENTEACHAT=VEN','','MODIFICATION');
    //
    325210 : AGLLanceFiche('BTP','BTEDTBONPAIEMENT','GP_NATUREPIECEG=FBT;GP_VENTEACHAT=VEN','','COTRAITANCE;MODIFICATION;STATUT=AFF');  //Edition lettre d'éclatement et bon de paiement
    325310 : BEGIN
    					AGLLanceFiche('GC','GCPORT_MUL','','','COTRAITANCE') ;
    			   END;
    325320 : BEGIN SaisirHorsInside('BTPARAMMAIL','TYPEMAIL=C'); Result := -1 ; end;
    325330 : BEGIN SaisirHorsInside('BTPARAMMAIL','TYPEMAIL=T'); Result := -1 ; end;
    //Devis
    325411 : AGLLanceFiche('BTP','BTDEVISCOTRAI_MUL','GP_NATUREPIECEG=DBT;GP_VENTEACHAT=VEN','','MODIFICATION;COTRAITANCE;STATUT=AFF') ;   // Devis
    325412 : AGLLanceFiche('BTP','BTACCDEVCOT_MUL','GP_NATUREPIECEG=DBT','','MODIFICATION;COTRAITANCE;STATUT=AFF') ;   // Devis
    325413 : AGLLanceFiche('BTP','BTACCDEVCOT_MUL','GP_NATUREPIECEG=DBT','','MODIFICATION;COTRAITANCE;ACCEPTINVERSE=X;STATUT=AFF') ;   // Devis
    325414 : AGLLanceFiche('BTP','BTEDITDOCCOT_MUL','GP_NATUREPIECEG=DBT','','COTRAITANCE;VEN;STATUT=AFF') ;
    325415 : AGLLanceFiche('BTP','BTPTFPIECECOT','GP_NATUREPIECEG=DBT','','DEVIS;COTRAITANCE;STATUT=AFF') ;  // Portefeuille pièces
    //Factures
    325510 : AGLLanceFiche('BTP','BTSELFAC_MUL','AFF_ETATAFFAIRE=ACP','','MODIFICATION;STATUT=AFF;COTRAITANCE;STATUT=AFF') ;   // Prépa Factures
    325521 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=FBT;GP_VENTEACHAT=VEN','','COTRAITANCE;MODIFICATION;STATUT=AFF');  // Modif Factures
    325523 : AGLLanceFiche('BTP','BTREGLCOTRAIT_MUL','GP_NATUREPIECEG=FBT','','COTRAITANCE;STATUT=AFF') ; // Saisie des reglements Cotraitants
    325522 : AGLLanceFiche('BTP','BTEDITDOCCOT_MUL','GP_NATUREPIECEG=FBT','','COTRAITANCE;VEN;STATUT=AFF') ;                   // Edition de factures
    //
    else HShowMessage('2;?caption?;'+'Fonction non disponible : '+';W;O;O;O;','CoTraitance',IntToStr(Numtag)) ;
  end;
end;

function GestionMenuSoustraitance (Numtag : integer) : integer;
begin

	if not VH_GC.SeriaSousTraitance then
  begin
  	PgiBox('Ce module n''est pas sérialisé. Il ne peut être utilisé.','Sous-Traitance');
    result := -1;
    exit;
  end;

  result := 0;

  Case Numtag of
    //
    327110 : BEGIN SaisirHorsInside('BTINTEGREDOC','TYPEIMPORT=SST'); Result := -1 ; end;               //intégration document XLS cotraitant
    327220 : AGLLanceFiche('BTP','BTEDITDOCCOT_MUL','GP_NATUREPIECEG=DBT','','SOUSTRAITANCE;VEN;STATUT=AFF') ;
    327230 : AGLLanceFiche('BTP','BTPTFPIECECOT','GP_NATUREPIECEG=DBT','','DEVIS;SOUSTRAITANCE;STATUT=AFF') ;  // Portefeuille pièces
    // Devis
    327211 : AGLLanceFiche('BTP','BTDEVISSTRAITE','GP_NATUREPIECEG=DBT;GP_VENTEACHAT=VEN','','MODIFICATION;SOUSTRAITANCE;STATUT=AFF') ;  // Devis
    327212 : AGLLanceFiche('BTP','BTACCDEVCOT_MUL','GP_NATUREPIECEG=DBT','','MODIFICATION;SOUSTRAITANCE;STATUT=AFF') ;                   // Devis
    327213 : AGLLanceFiche('BTP','BTACCDEVCOT_MUL','GP_NATUREPIECEG=DBT','','MODIFICATION;SOUSTRAITANCE;ACCEPTINVERSE=X;STATUT=AFF') ;   // Devis
    327214 : AGLLanceFiche('BTP','BTEDITDOCCOT_MUL','GP_NATUREPIECEG=DBT','','SOUSTRAITANCE;VEN;STATUT=AFF') ;
    327215 : AGLLanceFiche('BTP','BTPTFPIECECOT','GP_NATUREPIECEG=DBT','','DEVIS;SOUSTRAITANCE;STATUT=AFF') ;// Portefeuille pièces
    // Facturation
    327330 : AGLLanceFiche('BTP','BTSELFAC_MUL','AFF_ETATAFFAIRE=ACP','','MODIFICATION;STATUT=AFF;SOUSTRAITANCE;STATUT=AFF') ;      // Prépa Factures
    327320 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=FBT;GP_VENTEACHAT=VEN','','SOUSTRAITANCE;MODIFICATION;STATUT=AFF'); // Modif Factures
    327340 : AGLLanceFiche('BTP','BTEDITDOCCOT_MUL','GP_NATUREPIECEG=FBT','','SOUSTRAITANCE;VEN;STATUT=AFF') ;                        // Edition de factures
    327310 : AGLLanceFiche('BTP','BTREGLCOTRAIT_MUL','GP_NATUREPIECEG=FBT','','SOUSTRAITANCE;STATUT=AFF') ;                         // Saisie des reglements Cotraitants

    //Edition des lettres de réglements
    327410 : AGLLanceFiche('BTP','BTEDTBONPAIEMENT','GP_NATUREPIECEG=FBT;GP_VENTEACHAT=VEN','','SOUSTRAITANCE;MODIFICATION;STATUT=AFF');  //Edition lettre d'éclatement et bon de paiement
    //
    327510 : BEGIN SaisirHorsInside('BTPARAMMAIL', 'TYPEMAIL=S'); Result := -1 ; end;

    else HShowMessage('2;?caption?;'+'Fonction non disponible : '+';W;O;O;O;','Sous-Traitance',IntToStr(Numtag)) ;
  end;

end;

function IsIntervenant (CodeAffaire,Intervenant : string) : boolean;
var QQ : TQuery;
begin
	QQ := OpenSql ('SELECT BAI_TYPEINTERV FROM AFFAIREINTERV WHERE BAI_AFFAIRE="'+CodeAffaire+'" AND BAI_TIERSFOU="'+Intervenant+'"',true,
  							 -1,'',true);
	result := not QQ.eof;
  ferme (QQ);
end;

PROCEDURE SetSousTraitantUsed (TOBSousTrait: TOB;Fournisseur : string);
var TOBST : TOB;
begin
  if TOBSousTrait = nil then Exit;
	TOBST := TOBSousTrait.FindFirst(['BPI_TIERSFOU'],[Fournisseur],true);
  if TOBST <> nil then TOBST.PutValue('UTILISE','X');
end;
procedure Setallused (TOBSSTRAIT : TOB);
var i : Integer;
begin
  for i := 0 to TOBSSTrait.Detail.Count - 1 do
  begin
    TOBSSTrait.Detail[i].SetString('UTILISE','X');
  end;
end;


procedure ValideLesSousTrait(TOBPiece,TOBSSTrait : TOB; DEV : rdevise);
var i, Numero, Indice: integer;
  Nature, Souche: string;
  TOBP: TOB;
  okok : boolean;
begin
  if TOBSSTrait = nil then Exit;
  if TOBSSTrait.Detail.Count = 0 then Exit;
    i:=0;
  repeat
    TOBP := TOBSSTrait.Detail[i];
    if TOBP.GetDouble('NBUSED') > 0 then TOBP.SetString('UTILISE','X');
    if (TOBP.getvalue('UTILISE')='-') then
    begin
    	TOBP.free;
    end else inc(I);
  until i >= TOBSSTrait.Detail.Count;
  //
  Nature := TOBPiece.GetValue('GP_NATUREPIECEG');
  Souche := TOBPiece.GetValue('GP_SOUCHE');
  Numero := TOBPiece.GetValue('GP_NUMERO');
  Indice := TOBPiece.GetValue('GP_INDICEG');
  for i := 0 to TOBSSTrait.Detail.Count - 1 do
  begin
    TOBP := TOBSSTrait.Detail[i];
    TOBP.PutValue('BPI_NATUREPIECEG', Nature);
    TOBP.PutValue('BPI_SOUCHE', Souche);
    TOBP.PutValue('BPI_NUMERO', Numero);
    TOBP.PutValue('BPI_INDICEG', Indice);
    TOBP.PutValue('BPI_ORDRE', I+1);
  end;
  TOBSSTrait.SetAllModifie(true);
  okok := false;
  TRY
    okok := TOBSSTrait.InsertDB(nil);
  EXCEPT
    on E: Exception do
    begin
      PgiError('Erreur SQL : ' + E.Message, 'Erreur mise à jour des PIECEINTERV');
    end;
  END;
  if not okok then
  begin
    V_PGI.IoError := oeUnknown;
  end;
end;

procedure ValideLesPieceTrait(TOBPiece, TOBaffaire,TOBPieceTrait,TOBSousTrait : TOB; DEV : Rdevise);
var i, Numero, Indice: integer;
  Nature, Souche: string;
  TOBP: TOB;
  TOTHt,TOTTTC,MaxHt,Diff : double;
  TmaxHt : TOB;
  okok : boolean;
begin
	if not PieceTraitUsable (TOBPiece,TOBaffaire) then exit;
  if TOBPieceTrait = nil then Exit;
  if TOBPieceTrait.Detail.Count = 0 then Exit;
  //
  i:=0;
  repeat
    TOBP := TOBPieceTrait.Detail[i];
    if (TOBP.getvalue('BPE_TOTALHTDEV')=0) and (TOBP.getvalue('BPE_MONTANTREGL')=0) then
    begin
    	TOBP.free;
    end else inc(I);
  until i >= TOBPieceTrait.Detail.Count;
  //
  Nature := TOBPiece.GetValue('GP_NATUREPIECEG');
  Souche := TOBPiece.GetValue('GP_SOUCHE');
  Numero := TOBPiece.GetValue('GP_NUMERO');
  Indice := TOBPiece.GetValue('GP_INDICEG');
  TmaxHt := nil;
  TOTHT := 0; TOTTTC := 0;
  MaxHt := 0;
  for i := 0 to TOBPieceTrait.Detail.Count - 1 do
  begin
    TOBP := TOBPieceTrait.Detail[i];
    TOBP.PutValue('BPE_NATUREPIECEG', Nature);
    TOBP.PutValue('BPE_SOUCHE', Souche);
    TOBP.PutValue('BPE_NUMERO', Numero);
    TOBP.PutValue('BPE_INDICEG', Indice);
    TOBP.PutValue('BPE_AFFAIRE', TOBpiece.getValue('GP_AFFAIRE'));
    if TOBP.GetValue('BPE_TYPEINTERV')='Y00' then
    begin
    	SetSousTraitantUsed (TOBSousTrait,TOBP.GetValue('BPE_FOURNISSEUR'));
    end;
    TOTHT := TOTHT + TOBP.getvalue('BPE_TOTALHTDEV');
    TOTTTC := TOTTTC + TOBP.getvalue('BPE_TOTALTTCDEV');
    if TOBP.getvalue('BPE_TOTALHTDEV') > maxHt then TmaxHt := TOBP;
  end;
  Diff := Arrondi(TOTHT - TOBpiece.getValue('GP_TOTALHTDEV'),Dev.Decimale);
  if TMaxHt <> nil then
  begin
    if diff <> 0 then
    begin
      TMaxHT.Putvalue('BPE_TOTALHTDEV',TMaxHT.getvalue('BPE_TOTALHTDEV')+Diff);
    end;
    Diff := Arrondi(TOTTTC - TOBpiece.getValue('GP_TOTALTTCDEV'),Dev.Decimale);
    if diff <> 0 then
    begin
      TMaxHT.Putvalue('BPE_TOTALTTCDEV',TMaxHT.getvalue('BPE_TOTALTTCDEV')+Diff);
    end;
  end;
  okok := false;
  TRY
    okok := TOBPieceTrait.InsertDB(nil);
  EXCEPT
    on E: Exception do
    begin
      PgiError('Erreur SQL : ' + E.Message, 'Erreur mise à jour des PIECETRAIT');
    end;
  END;
  if not okok then
  begin
    V_PGI.IoError := oeUnknown;
  end;
end;


{ TPieceCotrait }

constructor TPieceCotrait.create ( LaForm : Tform);
var ThePop : Tcomponent;
		NatPrest : string;
begin
  NatPrest := GetParamSocSecur('SO_BTNATPRESTATION','');
  if (NatPrest = '') and (VH_GC.SeriaSousTraitance)  then PgiInfo ('Merci de paramétrer la nature de prestation dans les paramètres sociétés BTP/Sous-Traitance');
	fInfoSSTrait := TInfoSSTrait.create;
  fInfoSSTrait.Code := GetParamSocSecur('SO_BTNATPRESTATION','');
	fusable := true; // par défaut
  fTOBligneArecalc := TOB.Create ('LES LIGNES A RECALC',nil,-1);
	if (not VH_GC.SeriaCoTraitance) and (not VH_GC.SeriaSousTraitance) then fusable := false; // meme pas la peine de continuer
  FF := laForm;
	if (FF is TFLigNomen) then
  begin
    fusable := false;
    GS := TFligNomen(FF).G_NLIG;
  end else if (FF is TFFacture) then
  begin
    GS := TFfacture(FF).GS;
		fcopiercoller := TFFacture(FF).CopierColleObj;
  end else fusable := false;
  //
  if FF is TFfacture then
  begin
		fModifSsDetail := TFFacture(FF).ModifSSDetail;
    ThePop := FF.Findcomponent  ('POPBTP');
    if ThePop = nil then
    BEGIN
      // pas de menu BTP trouve ..on le cree
      fPOPGS := TPopupMenu.Create(FF);
      fPOPGS.Name := 'POPBTP';
      fCreatedPop := true;
    END else
    BEGIN
      fCreatedPop := false;
      fPOPGS := TPopupMenu(thePop);
    END;
    ThePop := FF.Findcomponent  ('POPY');
    if ThePop <> nil then
    begin
      fPOPY := TPopupMenu(thePop);
    end;
  end else
  begin
    ThePop := FF.Findcomponent  ('PopupG_NLIG');
    if ThePop <> nil then
    begin
      fCreatedPop := false;
      fPOPGS := TPopupMenu(thePop);
    end;
  end;
  DefiniMenuPop(FF);
//  PieceTraitUsable := fusable;
end;

destructor TPieceCotrait.destroy;
var indice : integer;
begin
  fTOBligneArecalc.free;
  for Indice := 0 to fMaxItems -1 do
  begin
  	if fMenuItem[Indice]<> nil then fMenuItem[Indice].Free;
  end;
	if fcreatedPop then fPOPGS.free;
	fInfoSSTrait.free;
//  PieceTraitUsable := false;
  inherited;
end;


procedure TPieceCotrait.DefiniMenuPop (Parent : Tform);
var Indice : integer;
		ThePop : TComponent;
begin
  fMaxItems := 0;
  if not fcreatedPop then
  begin
    fMenuItem[fMaxItems] := TmenuItem.Create (parent);
    with fMenuItem[fMaxItems] do
      begin
    	Name := 'BCOTRAITSEP';
      Caption := '-';
      Visible := false;
      end;
    inc (fMaxItems);
  end;
	if (VH_GC.SeriaCoTraitance) then
  begin
    // Affectation à un cotraitant
    fMenuItem[fMaxItems] := TmenuItem.Create (parent);
    with fMenuItem[fMaxItems] do
      begin
      Name := 'BCOTRAITAFFECT';
      Caption := 'Affectation à un cotraitant';
      Visible := false;
      OnClick := GSAffectCotrait;
      end;
    inc (fMaxItems);
  end;
	if (VH_GC.SeriaSousTraitance) then
  begin
    fMenuItem[fMaxItems] := TmenuItem.Create (parent);
    with fMenuItem[fMaxItems] do
    begin
    	Name := 'BSOUSTRAITSEP';
      Caption := '-';
      Visible := false;
    end;
    inc (fMaxItems);

    fMenuItem[fMaxItems] := TmenuItem.Create (parent);
    with fMenuItem[fMaxItems] do
    begin
      Name := 'BGESTSOUSTRAIT';
      Caption := 'Gestion des sous-traitants';
      Visible := false;
      OnClick := GSGestionSoustrait;
    end;
    inc (fMaxItems);
    // Affectation à un sous traitant
    fMenuItem[fMaxItems] := TmenuItem.Create (parent);
    with fMenuItem[fMaxItems] do
    begin
      Name := 'BSOUSTRAITAFFECT';
      Caption := 'Affectation à un Sous-traitant';
      Visible := false;
      if VH_GC.BTCODESPECIF = '001' then
      begin
        OnClick := GSAffectSoustraitPOC;
      end else
      begin
        OnClick := GSAffectSoustrait;
      end;
    end;
    inc (fMaxItems);
  end;
  // Récupération des travaux
  fMenuItem[fMaxItems] := TmenuItem.Create (parent);
  with fMenuItem[fMaxItems] do
    begin
    Name := 'BCOINTERNE';
    Caption := 'Affectation en interne';
    Visible := false;
    OnClick := GSAffectInterne;
    end;
  inc (fMaxItems);

  for Indice := 0 to fMaxItems -1 do
  begin
  	if fMenuItem [Indice] <> nil then fPOPGS.Items.Add (fMenuItem[Indice]);
  end;

  // definition sur POPY du facture.pas
  if FF is TFFacture then
  begin
    fMaxpopyItems := 0;
    if fPOPY <> nil then
    begin
      ThePOp := TFFacture(FF).FindComponent('BPOPYCOTRAITSEP');
      if ThePop <> nil then fPOPYItem[fMaxpopyItems]  := TmenuItem(ThePop);
      with fPOPYItem[fMaxpopyItems] do
      begin
        Visible := false;
      end;
      inc (fMaxpopyItems);
      ThePOp := TFFacture(FF).FindComponent('BCOTRAITTABLEAU');
      if ThePop <> nil then fPOPYItem[fMaxpopyItems]  := TmenuItem(ThePop);
      with fPOPYItem[fMaxpopyItems] do
      begin
        Visible := false;
        OnClick := GSCotraittableau;
      end;
      inc (fMaxpopyItems);
      ThePOp := TFFacture(FF).FindComponent('PAIESTPOC');
      if ThePop <> nil then fPOPYItem[fMaxpopyItems]  := TmenuItem(ThePop);
      with fPOPYItem[fMaxpopyItems] do
      begin
        Visible := false;
        OnClick := GSPaieSTPOC;
      end;
      inc (fMaxpopyItems);

    end;
  end;
end;

function TPieceCotrait.FindLignePere (TOBL : TOB) : TOB;
var indice : integer;
		TOBI : TOB;
    IndiceNomen : integer;
begin
	IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
  //
	result := nil;
  for Indice := TOBL.GetIndex -1 downto 0 do
  begin
    TOBI := fTOBpiece.detail[Indice];
    if Isouvrage(TOBI) and (not IsSousDetail(TOBI)) and TOBI.GetValue('GL_INDICENOMEN')=IndiceNomen then
    begin
    	result := TOBI;
      break;
    end;
  end;
end;

function TPieceCotrait.AddLigneCalcul (TOBLP : TOB) : Boolean;
var TOBI : TOB;
begin
  TOBI := fTOBligneArecalc.FindFirst(['NUMORDRE'],[TOBLP.GetValue('GL_NUMORDRE')],true);
  result := (TOBI = nil);
  if result then
  begin
  	TOBI := TOB.Create ('UNE LIGNE',fTOBligneArecalc,-1);
    TOBI.AddChampSupValeur ('NUMORDRE',TOBLP.GetValue('GL_NUMORDRE'));
    TOBI.Data := TOBLP;
  end;
end;

procedure TPieceCotrait.AffecteCoTraitanceLigSDetail( TOBL : TOB; Fournisseur, LibelleFou, NatureTravail: string);
var TOBOUV,TOBOL,TOBLP : TOB;
		Indice : integer;
    indiceNomen,UniqueBlo : integer;
    found : boolean;
    Affect,WhoAffect : string;
    Multiple : boolean;
    Qte,QteDuDetail,OldTpsUnitaire : double;
begin
	found := false;
  TOBLP := FindLignePere(TOBL);
  if TOBLP = nil then exit;
  if (not IsVariante(TOBLP)) and (TOBLP.GetValue('GL_ARTICLE') <> '') then
  begin
    if AddLigneCalcul (TOBLP) then TFfacture(FF).deduitLaLigne (TOBLP);
  end;
  IndiceNomen := TOBL.getValue('GL_INDICENOMEN');
  UniqueBlo := TOBL.getValue('UNIQUEBLO');
	if IndiceNomen = 0 then exit;
	TOBOUV := TOBOuvrage.detail[IndiceNomen-1];
	for Indice := 0 to TOBOUV.detail.count -1 do
  begin
		TOBOL := TOBOUV.detail[Indice];
    if UniqueBlo <> TOBOL.getValue('BLO_UNIQUEBLO') then continue;
    Qte :=  TOBOL.getValue('BLO_QTEFACT');
    QteDuDetail :=  TOBOL.getValue('BLO_QTEDUDETAIL'); if QteDuDetail = 0 then QteDuDetail := 1;
    OldTpsUnitaire := Arrondi(((Qte/QteDuDetail) * TOBOL.GetValue ('BLO_TPSUNITAIRE')),V_PGI.OkDecQ);

    TOBOL.putValue('BLO_FOURNISSEUR',Fournisseur);
    TOBOL.putValue('GA_FOURNPRINC',Fournisseur);
    TOBOL.putValue('LIBELLEFOU',LibelleFou);
    TOBOL.putValue('BLO_NATURETRAVAIL',NatureTravail);
  	TOBOL.PutValue('BLO_TPSUNITAIRE',0);
    if TOBOL.FieldExists('BLF_NATURETRAVAIL') then
    begin
    	TOBOL.PutValue('BLF_NATURETRAVAIL',TOBOL.GetValue('BLO_NATURETRAVAIL'));
    	TOBOL.PutValue('BLF_FOURNISSEUR',TOBOL.GetValue('BLO_FOURNISSEUR'));
    end;
    TOBOL.putValue('BLO_DPA',TOBOL.GetValue('BLO_PUHT'));
    TOBOL.putValue('BLO_DPR',TOBOL.GetValue('BLO_PUHT'));
    TOBOL.putValue('BLO_COEFFR',0);
    TOBOL.putValue('BLO_COEFFG',0);
    TOBOL.putValue('BLO_COEFFC',0);
    TOBOL.putValue('BLO_COEFMARG',1);
    TOBOL.putValue('BLO_NONAPPLICFRAIS','X');
    TOBOL.putValue('BLO_NONAPPLICFC','X');
    TOBOL.putValue('BLO_NONAPPLICFG','X');
    if TOBOL.detail.count > 0 then
    begin
    	AffecteCoTraitanceSousDetail1 (TOBOL,indiceNomen,Fournisseur,LibelleFou,NatureTravail);
    end;
    found := true;
    break;
  end;
  if found then
  begin
    TOBL.putvalue('GLC_NATURETRAVAIL',NatureTravail);
    TOBL.putvalue('GL_FOURNISSEUR',Fournisseur);
    if TOBL.FieldExists('BLF_NATURETRAVAIL') then
    begin
      TOBL.putvalue('BLF_NATURETRAVAIL',NatureTravail);
      TOBL.putvalue('BLF_FOURNISSEUR',Fournisseur);
    end;
    TOBL.putvalue('LIBELLEFOU',LibelleFou);
    TOBL.PutValue('GL_TPSUNITAIRE',TOBL.GetValue('GL_TPSUNITAIRE')- OldTpsUnitaire);
    TOBL.putvalue('GL_HEURE',Arrondi(TOBL.getValue('GL_TPSUNITAIRE')*TOBL.GetValue('GL_QTEFACT'),V_PGI.okdecQ));
    TOBL.putvalue('GL_TOTALHEURE',TOBL.Getvalue('GL_HEURE'));
    TFFActure(FF).AfficheLaLigne(TOBL.getIndex);
  	Multiple := false;
    for Indice := 0 to TOBOUV.detail.count -1 do
    begin
      TOBOL := TOBOUV.detail[Indice];
      if Indice = 0 then
      begin
      	Affect := TOBOL.getValue('BLO_NATURETRAVAIL');
      	WhoAffect := TOBOL.getValue('BLO_FOURNISSEUR');
      end;
      if (TOBOL.getValue('BLO_NATURETRAVAIL')<>Affect) or (TOBOL.getValue('BLO_FOURNISSEUR')<>WhoAffect) then Multiple := true;
    end;
    //
    if (Multiple) then
    begin
      TOBLP.putvalue('GLC_NATURETRAVAIL','010');
      TOBLP.putvalue('GL_FOURNISSEUR','');
      TOBLP.putvalue('LIBELLEFOU','');
    end else if (not Multiple) and (Affect='') then
    begin
      TOBLP.putvalue('GLC_NATURETRAVAIL','');
      TOBLP.putvalue('GL_FOURNISSEUR','');
      TOBLP.putvalue('LIBELLEFOU','');
    end else if (not Multiple) and (Affect<>'') then
    begin
      TOBLP.putvalue('GLC_NATURETRAVAIL',TOBOUV.detail[0].getValue('BLO_NATURETRAVAIL'));
      TOBLP.putvalue('GL_FOURNISSEUR',TOBOUV.detail[0].getValue('BLO_FOURNISSEUR'));
      TOBLP.putvalue('LIBELLEFOU',TOBOUV.detail[0].getValue('LIBELLEFOU'));
    end;
    if TOBLP.FieldExists('BLF_NATURETRAVAIL') then
    begin
      TOBLP.putvalue('BLF_NATURETRAVAIL',TOBLP.getValue('GLC_NATURETRAVAIL'));
      TOBLP.putvalue('BLF_FOURNISSEUR',TOBLP.getValue('GL_FOURNISSEUR'));
    end;
    TFFActure(FF).AfficheLaLigne(TOBLP.getIndex);
    TOBLP.putvalue('GL_RECALCULER','X');
    fTOBPiece.putvalue('GP_RECALCULER','X');
  end;
  // deporte
  // if (not IsVariante(TOBLP)) and (TOBLP.GetValue('GL_ARTICLE') <> '')  then TFfacture(FF).AjouteLaLigneCotrait (TOBLP);
end;

procedure TPieceCotrait.AffecteSousTraitanceLigSDetail( TOBL : TOB; Fournisseur, ModePaie,CodeMarche,LibelleFou, NatureTravail: string);
var TOBOUV,TOBOL,TOBLP : TOB;
		Indice : integer;
    indiceNomen,UniqueBlo : integer;
    found : boolean;
    Affect,WhoAffect : string;
    Multiple,PrixBloque,FromExcel : boolean;
    OldTpsUnitaire,Qte,QteDuDetail : double;
    valeurs : t_valeurs;
begin
	found := false;
  TOBLP := FindLignePere(TOBL);
  if TOBLP = nil then exit;
  PrixBloque := (TOBLP.getValue('GL_BLOQUETARIF')='X');
  FromExcel := IsFromExcel(TOBLP);
  if (not IsVariante(TOBLP)) and (TOBLP.GetValue('GL_ARTICLE') <> '') then
  begin
    if AddLigneCalcul (TOBLP) then TFfacture(FF).deduitLaLigne (TOBLP);
  end;
  IndiceNomen := TOBL.getValue('GL_INDICENOMEN');
  UniqueBlo := TOBL.getValue('UNIQUEBLO');
	if IndiceNomen = 0 then exit;
  OldTpsUnitaire := 0;
  
	TOBOUV := TOBOuvrage.detail[IndiceNomen-1];
	for Indice := 0 to TOBOUV.detail.count -1 do
  begin
		TOBOL := TOBOUV.detail[Indice];
    if UniqueBlo <> TOBOL.getValue('BLO_UNIQUEBLO') then continue;
    Qte :=  TOBOL.getValue('BLO_QTEFACT');
    QteDuDetail :=  TOBOL.getValue('BLO_QTEDUDETAIL'); if QteDuDetail = 0 then QteDuDetail := 1;
    OldTpsUnitaire := Arrondi(((Qte/QteDuDetail) * TOBOL.GetValue ('BLO_TPSUNITAIRE')),V_PGI.OkDecQ);
    TOBOL.putValue('BLO_FOURNISSEUR',Fournisseur);
    TOBOL.putValue('GA_FOURNPRINC',Fournisseur);
    TOBOL.putValue('LIBELLEFOU',LibelleFou);
    TOBOL.putValue('BLO_NATURETRAVAIL',NatureTravail);
    TOBOL.putValue('BLO_CODEMARCHE',CodeMarche);
    if TOBOL.FieldExists('BLF_NATURETRAVAIL') then
    begin
      TOBOL.putValue('BLF_FOURNISSEUR',TOBOL.getValue('BLO_FOURNISSEUR'));
      TOBOL.putValue('BLF_NATURETRAVAIL',TOBOL.getValue('BLO_NATURETRAVAIL'));
      TOBOL.putValue('BLF_CODEMARCHE',TOBOL.getValue('BLO_CODEMARCHE'));
    end;
    TOBOL.PutValue('BLO_TPSUNITAIRE',0);
    if TOBOL.detail.count > 0 then
    begin
      AffecteSousTraitanceSousDetail1 (TOBOL,indiceNomen,Fournisseur,ModePaie,CodeMarche,LibelleFou,NatureTravail,PrixBloque,FromExcel);
      InitTableau (Valeurs);
      CalculeOuvrageDoc (TOBOL,1,1,true,fDEV,valeurs,(TOBLP.getValue('GL_FACTUREHT')='X'),true,true);
      SetValeursOuv(TOBOL,valeurs);
    end else
    begin
      if TOBOL.getDouble('BLO_DPA')=0 then
      begin
        TOBOL.setDouble('BLO_DPA',TOBOL.getDOuble('BLO_PUHT'));
        TOBOL.setDouble('BLO_DPR',TOBOL.getDOuble('BLO_PUHT'));
        TOBOL.setDouble('BLO_COEFFG',0);
        TOBOL.setDouble('BLO_COEFFC',0);
        TOBOL.setDouble('BLO_COEFFR',0);
        TOBOL.setDouble('BLO_COEFMARG',0);
        TOBOL.PutValue('POURCENTMARQ',0);
      end else
      begin
        if fInfoSSTrait.CoefFG <> 0 then TOBOL.putvalue('BLO_COEFFG',fInfoSSTrait.CoefFG-1);
        if fInfoSSTrait.CoefMarg <> 0 then
        begin
          TOBOL.putvalue('BLO_COEFMARG',fInfoSSTrait.CoefMarg);
          TOBOL.PutValue('POURCENTMARG',Arrondi((TOBOL.GetValue('BLO_COEFMARG')-1)*100,2));
        end;
        RecalcPaPRPV(TOBOL,prixBloque,FromExcel);
        if TOBOL.GetValue('BLO_PUHT') <> 0 then
        begin
          TOBOL.PutValue('POURCENTMARQ',Arrondi(((TOBOL.GetValue('BLO_PUHT')- TOBOL.GetValue('BLO_DPR'))/TOBOL.GetValue('BLO_PUHT'))*100,2));
        end else
        begin
          TOBOL.PutValue('POURCENTMARQ',0);
        end;
      end;
    end;

    found := true;
    break;
  end;

  if TOBLP <> nil then
  begin
  	InitTableau (Valeurs);
    CalculeOuvrageDoc (TOBOUV,1,1,true,fDEV,valeurs,(TOBLP.getValue('GL_FACTUREHT')='X'),true,true);
    SetValeurs(TOBLP,valeurs);
  end;

  if found then
  begin
    TOBL.putvalue('GLC_NATURETRAVAIL',NatureTravail);
    TOBL.putvalue('GL_FOURNISSEUR',Fournisseur);
    TOBL.putvalue('GL_CODEMARCHE',CodeMarche);
    TOBL.putvalue('LIBELLEFOU',LibelleFou);
    TOBL.PutValue('GL_TPSUNITAIRE',TOBL.GetValue('GL_TPSUNITAIRE')- OldTpsUnitaire);
    TOBL.putvalue('GL_HEURE',Arrondi(TOBL.getValue('GL_TPSUNITAIRE')*TOBL.GetValue('GL_QTEFACT'),V_PGI.okdecQ));
    TOBL.putvalue('GL_TOTALHEURE',TOBL.Getvalue('GL_HEURE'));

    if TOBL.FieldExists('BLF_NATURETRAVAIL') then
    begin
      TOBL.putValue('BLO_FOURNISSEUR',TOBL.getValue('GL_FOURNISSEUR'));
      TOBL.putValue('BLO_NATURETRAVAIL',TOBL.getValue('GLC_NATURETRAVAIL'));
      TOBL.putValue('BLO_CODEMARCHE',TOBL.getValue('GL_CODEMARCHE'));
    end;

    TFFActure(FF).AfficheLaLigne(TOBL.getIndex);
  	Multiple := false;
    for Indice := 0 to TOBOUV.detail.count -1 do
    begin
      TOBOL := TOBOUV.detail[Indice];
      if Indice = 0 then
      begin
      	Affect := TOBOL.getValue('BLO_NATURETRAVAIL');
      	WhoAffect := TOBOL.getValue('BLO_FOURNISSEUR');
      end;
      if (TOBOL.getValue('BLO_NATURETRAVAIL')<>Affect) or (TOBOL.getValue('BLO_FOURNISSEUR')<>WhoAffect) then Multiple := true;
    end;
    //
    if (Multiple) then
    begin
      TOBLP.putvalue('GLC_NATURETRAVAIL','010');
      TOBLP.putvalue('GL_FOURNISSEUR','');
      TOBLP.putvalue('LIBELLEFOU','');
    end else if (not Multiple) and (Affect='') then
    begin
      TOBLP.putvalue('GLC_NATURETRAVAIL','');
      TOBLP.putvalue('GL_FOURNISSEUR','');
      TOBLP.putvalue('LIBELLEFOU','');
    end else if (not Multiple) and (Affect<>'') then
    begin
      TOBLP.putvalue('GLC_NATURETRAVAIL',TOBOUV.detail[0].getValue('BLO_NATURETRAVAIL'));
      TOBLP.putvalue('GL_FOURNISSEUR',TOBOUV.detail[0].getValue('BLO_FOURNISSEUR'));
      TOBLP.putvalue('LIBELLEFOU',TOBOUV.detail[0].getValue('LIBELLEFOU'));
    end;
    if TOBLP.FieldExists('BLF_NATURETRAVAIL') then
    begin
      TOBLP.putvalue('BLF_NATURETRAVAIL',TOBLP.Getvalue('GLC_NATURETRAVAIL'));
      TOBLP.putvalue('BLF_FOURNISSEUR',TOBLP.Getvalue('GL_FOURNISSEUR'));
      TOBLP.putvalue('BLF_CODEMARCHE',TOBLP.Getvalue('GL_CODEMARCHE'));
    end;
    TFFActure(FF).AfficheLaLigne(TOBLP.getIndex);
    TOBLP.putvalue('GL_RECALCULER','X');
    fTOBPiece.putvalue('GP_RECALCULER','X');
  end;
  // deporte
  // if (not IsVariante(TOBLP)) and (TOBLP.GetValue('GL_ARTICLE') <> '')  then TFfacture(FF).AjouteLaLigneCotrait (TOBLP);
end;

procedure TPieceCotrait.AffecteInterneSDetail(TOBL: TOB; PrixBloque,FromExcel : boolean);
var TOBOUV,TOBOL,TOBLP : TOB;
		Indice : integer;
    indiceNomen,UniqueBlo : integer;
    found : boolean;
    Affect,WhoAffect : string;
    Multiple : boolean;
    valeurs : t_valeurs;
begin
	found := false;
  multiple := false;
  TOBLP := FindLignePere(TOBL);
  if TOBLP = nil then exit;

  if (not IsVariante(TOBLP)) and (TOBLP.GetValue('GL_ARTICLE') <> '') then
  begin
  	if AddLigneCalcul (TOBLP) then TFfacture(FF).deduitLaLigne (TOBLP);
  end;

  IndiceNomen := TOBL.getValue('GL_INDICENOMEN');
  UniqueBlo := TOBL.getValue('UNIQUEBLO');
	if IndiceNomen = 0 then exit;
	TOBOUV := TOBOuvrage.detail[IndiceNomen-1];

	for Indice := 0 to TOBOUV.detail.count -1 do
  begin
		TOBOL := TOBOUV.detail[Indice];
    if UniqueBlo <> TOBOL.getValue('BLO_UNIQUEBLO') then continue;
    TOBOL.putValue('BLO_FOURNISSEUR','');
    TOBOL.putValue('BLO_CODEMARCHE','');
    TOBOL.putValue('GA_FOURNPRINC','');
    TOBOL.putValue('LIBELLEFOU','');
    TOBOL.putValue('BLO_NATURETRAVAIL','');
    TOBOL.putValue('BLO_DPA',TOBOL.GetValue('BLO_PUHT'));
    TOBOL.putValue('BLO_DPR',TOBOL.GetValue('BLO_PUHT'));
    TOBOL.putValue('BLO_COEFFR',0);
    TOBOL.putValue('BLO_COEFFG',0);
    TOBOL.putValue('BLO_COEFFC',0);
    TOBOL.putValue('BLO_COEFMARG',1);
    TOBOL.putValue('BLO_NONAPPLICFRAIS','-');
    TOBOL.putValue('BLO_NONAPPLICFC','-');
    TOBOL.putValue('BLO_NONAPPLICFG','-');
    if TOBOL.FieldExists('BLF_NATURETRAVAIL') then
    begin
      TOBOL.putValue('BLF_FOURNISSEUR','');
      TOBOL.putValue('BLF_NATURETRAVAIL','');
    end;
    if TOBOL.detail.count > 0 then
    begin
    	AffecteInterneSousDetail1 (TOBOL,indiceNomen,PrixBloque,FromExcel);
      InitTableau (Valeurs);
      CalculeOuvrageDoc (TOBOL,1,1,true,fDEV,valeurs,(TOBLP.getValue('GL_FACTUREHT')='X'),true,true);
      SetValeursOuv(TOBOL,valeurs);
    end else
    begin
    	SetInfoArt (TOBOL,PrixBloque,FromExcel); // on remet à jour via l'article
    end;

    found := true;
    break;
  end;

  if found then
  begin
  	InitTableau (Valeurs);
    CalculeOuvrageDoc (TOBOUV,1,1,true,fDEV,valeurs,(TOBLP.getValue('GL_FACTUREHT')='X'),true,true);
    SetValeurs(TOBLP,valeurs);
    //
    TOBL.putvalue('GLC_NATURETRAVAIL','');
    TOBL.putvalue('GL_FOURNISSEUR','');
    TOBL.putvalue('GL_CODEMARCHE','');
    TOBL.putvalue('LIBELLEFOU','');
    if TOBL.FieldExists('BLF_NATURETRAVAIL') then
    begin
      TOBL.putValue('BLF_FOURNISSEUR',TOBL.getValue('GL_FOURNISSEUR'));
      TOBL.putValue('BLF_NATURETRAVAIL',TOBL.getValue('GLC_NATURETRAVAIL'));
      TOBL.putvalue('BLF_CODEMARCHE','');
    end;

    TFFActure(FF).AfficheLaLigne(TOBL.getIndex);
  //
  	Multiple := false;
    for Indice := 0 to TOBOUV.detail.count -1 do
    begin
      TOBOL := TOBOUV.detail[Indice];
      if Indice = 0 then
      begin
      	Affect := TOBOL.getValue('BLO_NATURETRAVAIL');
        WhoAffect := TOBOL.getValue('BLO_FOURNISSEUR');
      end;
      if (TOBOL.getValue('BLO_NATURETRAVAIL')<>Affect) or (TOBOL.getValue('BLO_FOURNISSEUR')<>WhoAffect) then Multiple := true;
    end;
    if (Multiple) then
    begin
      TOBLP.putvalue('GLC_NATURETRAVAIL','010');
      TOBLP.putvalue('GL_FOURNISSEUR','');
      TOBLP.putvalue('GL_CODEMARCHE','');
      TOBLP.putvalue('LIBELLEFOU','');
    end else if (not Multiple) and (Affect='') then
    begin
      TOBLP.putvalue('GLC_NATURETRAVAIL','');
      TOBLP.putvalue('GL_FOURNISSEUR','');
      TOBLP.putvalue('GL_CODEMARCHE','');
      TOBLP.putvalue('LIBELLEFOU','');
    end else if (not Multiple) and (Affect<>'') then
    begin
      TOBLP.putvalue('GLC_NATURETRAVAIL',TOBOUV.detail[0].getValue('BLO_NATURETRAVAIL'));
      TOBLP.putvalue('GL_FOURNISSEUR',TOBOUV.detail[0].getValue('BLO_FOURNISSEUR'));
      TOBLP.putvalue('GL_CODEMARCHE',TOBOUV.detail[0].getValue('BLO_CODEMARCHE'));
      TOBLP.putvalue('LIBELLEFOU',TOBOUV.detail[0].getValue('LIBELLEFOU'));
    end;
    if TOBLP.FieldExists('BLF_NATURETRAVAIL') then
    begin
      TOBLP.putvalue('BLF_NATURETRAVAIL',TOBLP.Getvalue('GLC_NATURETRAVAIL'));
      TOBLP.putvalue('BLF_FOURNISSEUR',TOBLP.Getvalue('GL_FOURNISSEUR'));
      TOBLP.putvalue('BLF_CODEMARCHE',TOBLP.Getvalue('GL_CODEMARCHE'));
    end;
    TFFActure(FF).AfficheLaLigne(TOBLP.getIndex);
    TOBLP.putvalue('GL_RECALCULER','X');
    fTOBPiece.putvalue('GP_RECALCULER','X');
  end;
  // if (not IsVariante(TOBLP)) and (TOBLP.GetValue('GL_ARTICLE') <> '')  then TFfacture(FF).AjouteLaLigneCotrait (TOBLP);
end;

procedure TPieceCotrait.AffecteCoTraitanceOuvrage (Fournisseur,LibelleFou,NatureTravail: string; TOBL : TOB; Var NextLig : integer; OkInc : boolean);
var TOBOUV : TOB;
		IndiceNomen : integer;
    PrixBloque : boolean;
    valeurs : t_valeurs;
begin
	IndiceNomen := TOBL.getvalue('GL_INDICENOMEN');
  PrixBloque := (TOBL.getvalue('GL_BLOQUETARIF')='X');
	if IndiceNomen > 0 then
  begin
    TOBOUV := TOBOuvrage.detail[IndiceNomen-1];
    AffecteCoTraitanceSousDetail(TOBL,TOBOUV,IndiceNomen,Fournisseur,LibelleFou,NatureTravail,nextLig,OkInc);
    InitTableau (Valeurs);
    CalculeOuvrageDoc (TOBOUV,1,1,true,fDEV,valeurs,(TOBL.getValue('GL_FACTUREHT')='X'),true,true);
    SetValeurs(TOBL,valeurs);
  end else Inc(NextLig);
end;

procedure TPieceCotrait.AffecteSousTraitanceOuvrage(TOBL: TOB; Fournisseur,LibelleFou, ModePaie,CodeMarche,NatureTravail: string; var NextLig: integer;OkInc: boolean; PrixBloque,FromExcel : boolean);
var TOBOUV : TOB;
		IndiceNomen : integer;
    valeurs:T_Valeurs ;
begin
	IndiceNomen := TOBL.getvalue('GL_INDICENOMEN');
	if IndiceNomen > 0 then
  begin
    TOBOUV := TOBOuvrage.detail[IndiceNomen-1];
    AffecteSousTraitanceSousDetail(TOBL,TOBOUV,IndiceNomen,Fournisseur,ModePaie,CodeMarche,LibelleFou,NatureTravail,nextLig,OkInc,PrixBloque,FromExcel);
    InitTableau (Valeurs);
    CalculeOuvrageDoc (TOBOUV,1,1,true,fDEV,valeurs,(TOBL.getValue('GL_FACTUREHT')='X'),true,true);
    SetValeurs(TOBL,valeurs);
  end;
end;

procedure TPieceCotrait.SetValeurs(TOBL: TOB; valeurs:T_Valeurs );
var Qte : double;
begin
  Qte := TOBL.Getvalue('GL_QTEFACT');
  TOBL.Putvalue('GL_MONTANTPAFG',valeurs[10]*Qte);
  TOBL.Putvalue('GL_MONTANTPAFR',valeurs[11]*Qte);
  TOBL.Putvalue('GL_MONTANTPAFC',valeurs[12]*Qte);
  TOBL.Putvalue('GL_MONTANTFG',valeurs[13]*Qte);
  TOBL.Putvalue('GL_MONTANTFR',valeurs[14]*Qte);
  TOBL.Putvalue('GL_MONTANTFC',valeurs[15]*Qte);

  TOBL.Putvalue('GL_DPA',valeurs[16]);
  TOBL.Putvalue('GL_DPR',valeurs[17]);
  //
  TOBL.Putvalue('GL_PUHTDEV',valeurs[2]);
  TOBL.Putvalue('GL_PUTTCDEV',valeurs[3]);
  TOBL.Putvalue('GL_PUHT',DeviseToPivotEx(TOBL.GetValue('GL_PUHTDEV'),fDEV.Taux,fDEV.quotite,V_PGI.OkdecP));
  TOBL.Putvalue('GL_PUTTC',DevisetoPivotEx(TOBL.GetValue('GL_PUTTCDEV'),fDEV.taux,fDEV.quotite,V_PGI.okdecP));
  TOBL.Putvalue('GL_PUHTBASE',TOBL.GetValue('GL_PUHT'));
  TOBL.Putvalue('GL_PUTTCBASE',TOBL.GetValue('GL_PUTTC'));
  TOBL.Putvalue('GL_DPA',valeurs[0]);
  TOBL.Putvalue('GL_DPR',valeurs[1]);
  TOBL.Putvalue('GL_PMAP',valeurs[6]);
  TOBL.Putvalue('GL_PMRP',valeurs[7]);
  TOBL.putvalue('GL_TPSUNITAIRE',valeurs[9]);
  TOBL.putvalue('GL_HEURE',Arrondi(TOBL.getValue('GL_TPSUNITAIRE')*Qte,V_PGI.okdecQ));
  TOBL.putvalue('GL_TOTALHEURE',TOBL.Getvalue('GL_HEURE'));
  TOBL.Putvalue('GL_MONTANTPA',Arrondi((Qte * TOBL.GetValue('GL_DPA')),V_PGI.okdecV));
  TOBL.Putvalue('GL_MONTANTPR',Arrondi((Qte * TOBL.GetValue('GL_DPR')),V_PGI.okdecV));
end;

procedure TPieceCotrait.SetValeursOuv (TOBL: TOB; valeurs:T_Valeurs );
var EnHt : boolean;
begin
  EnHt := (TOBL.getValue('BLO_FACTUREHT')='X');
  TOBL.PutValue ('GA_PAHT',valeurs[0]);
  TOBL.putvalue ('BLO_DPA',valeurs[0]);
  TOBL.putvalue ('BLO_DPR',valeurs[1]);
  TOBL.PutValue('BLO_PMAP',valeurs[6]);
  TOBL.PutValue('BLO_PMRP',valeurs[7]);
  // --
  TOBL.PutValue('BLO_TPSUNITAIRE',valeurs[9]);
  TOBL.PutValue('GA_HEURE',valeurs[9]);
  TOBL.PutValue('BLO_MONTANTPAFG',valeurs[10]);
  TOBL.PutValue('BLO_MONTANTPAFR',valeurs[11]);
  TOBL.PutValue('BLO_MONTANTPAFC',valeurs[12]);
  TOBL.PutValue('BLO_MONTANTFG',valeurs[13]);
  TOBL.PutValue('BLO_MONTANTFR',valeurs[14]);
  TOBL.PutValue('BLO_MONTANTFC',valeurs[15]);
  TOBL.PutValue('BLO_MONTANTPA',valeurs[16]);
  TOBL.PutValue('BLO_MONTANTPR',valeurs[17]);
  TOBL.PutValue('GA_PVHT',valeurs[2]);
  TOBL.putvalue ('BLO_PUHTDEV',valeurs[2]);
  TOBL.putvalue ('BLO_PUTTCDEV',valeurs[3]);
  TOBL.putvalue ('BLO_PUHT',devisetopivotEx(valeurs[2],fDEV.Taux,fDEV.quotite,V_PGI.okdecP));
  TOBL.putvalue ('BLO_PUTTC',devisetopivotEx(valeurs[3],fDEV.Taux,fDEV.quotite,V_PGI.okdecP));
  TOBL.putvalue ('BLO_PUHTBASE',TOBL.GetValue('BLO_PUHT'));
  TOBL.putvalue ('BLO_PUTTCBASE',TOBL.getValue('BLO_PUTTC'));
  TOBL.Putvalue('ANCPA',TOBL.Getvalue('BLO_DPA'));
  TOBL.Putvalue('ANCPR',TOBL.Getvalue('BLO_DPR'));
  if EnHt then TOBL.Putvalue('ANCPV',TOBL.Getvalue('BLO_PUHTDEV'))
          else TOBL.Putvalue('ANCPV',TOBL.Getvalue('BLO_PUTTCDEV'));
  CalculMontantHtDevLigOuv (TOBL,fDEV);
  StockeLesTypes (TOBL,valeurs);
end;


procedure  TPieceCotrait.AffecteCoTraitanceSousDetail1 (TOBOUV : TOB; indiceNomen : integer;Fournisseur,LibelleFou,NatureTravail: string);
var TOBOL : TOB;
		II : Integer;
    valeurs : t_valeurs;
begin
	for II := 0 to TOBOUV.detail.count -1 do
  begin
		TOBOL := TOBOUV.detail[II];
    TOBOL.putValue('BLO_FOURNISSEUR',Fournisseur);
    TOBOL.putValue('GA_FOURNPRINC',Fournisseur);
    TOBOL.putValue('LIBELLEFOU',LibelleFou);
    TOBOL.putValue('BLO_NATURETRAVAIL',NatureTravail);
    TOBOL.putValue('BLO_DPA',TOBOL.GetValue('BLO_PUHT'));
    TOBOL.putValue('BLO_DPR',TOBOL.GetValue('BLO_PUHT'));
    TOBOL.putValue('BLO_COEFFR',0);
    TOBOL.putValue('BLO_COEFFG',0);
    TOBOL.putValue('BLO_COEFFC',0);
    TOBOL.putValue('BLO_COEFMARG',1);
    TOBOL.putValue('BLO_NONAPPLICFRAIS','X');
    TOBOL.putValue('BLO_NONAPPLICFC','X');
    TOBOL.putValue('BLO_NONAPPLICFG','X');
  	TOBOL.PutValue('BLO_TPSUNITAIRE',0);
    if TOBOL.FieldExists('BLF_NATURETRAVAIL') then
    begin
      TOBOL.putValue('BLF_FOURNISSEUR',TOBOL.GetValue('BLO_FOURNISSEUR'));
      TOBOL.putValue('BLF_NATURETRAVAIL',TOBOL.GetValue('BLO_NATURETRAVAIL'));
    end;
    if TOBOL.detail.count > 0 then
    begin
    	AffecteCoTraitanceSousDetail1 (TOBOL,IndiceNomen,Fournisseur,LibelleFou,NatureTravail);
      InitTableau (Valeurs);
      CalculeOuvrageDoc (TOBOL,1,1,true,fDEV,valeurs,(TOBOL.getValue('BLO_FACTUREHT')='X'),true,true);
      SetValeursOuv(TOBOL,valeurs);
    end;
  end;
end;

procedure TPieceCotrait.AffecteSousTraitanceSousDetail1(TOBOUV: TOB;indiceNomen: integer; Fournisseur, ModePaie,CodeMarche,LibelleFou, NatureTravail: string; PrixBloque,FromExcel : boolean);
var TOBOL : TOB;
		II : Integer;
    valeurs : t_valeurs;
begin
	for II := 0 to TOBOUV.detail.count -1 do
  begin
		TOBOL := TOBOUV.detail[II];
    TOBOL.putValue('BLO_FOURNISSEUR',Fournisseur);
    TOBOL.putValue('GA_FOURNPRINC',Fournisseur);
    TOBOL.putValue('LIBELLEFOU',LibelleFou);
    TOBOL.putValue('BLO_NATURETRAVAIL',NatureTravail);
    TOBOL.putValue('BLO_CODEMARCHE',CodeMarche);
    if TOBOL.FieldExists('BLF_NATURETRAVAIL') then
    begin
    	TOBOL.PutValue('BLF_NATURETRAVAIL',TOBOL.GetValue('BLO_NATURETRAVAIL'));
    	TOBOL.putValue('BLF_FOURNISSEUR',TOBOL.GetValue('BLO_FOURNISSEUR'));
    	TOBOL.putValue('BLF_CODEMARCHE',TOBOL.GetValue('BLO_CODEMARCHE'));
    end;
    if TOBOL.detail.count > 0 then
    begin
      AffecteSousTraitanceSousDetail1 (TOBOL,IndiceNomen,Fournisseur,ModePaie,CodeMarche,LibelleFou,NatureTravail,prixBloque,FromExcel);
      InitTableau (Valeurs);
      CalculeOuvrageDoc (TOBOL,1,1,true,fDEV,valeurs,(TOBOL.getValue('BLO_FACTUREHT')='X'),true,true);
      SetValeursOuv(TOBOL,valeurs);
    end else
    begin
      if TOBOL.getDouble('BLO_DPA')=0 then
      begin
        TOBOL.setDouble('BLO_DPA',TOBOL.getDOuble('BLO_PUHT'));
        TOBOL.setDouble('BLO_DPR',TOBOL.getDOuble('BLO_PUHT'));
        TOBOL.setDouble('BLO_COEFFG',0);
        TOBOL.setDouble('BLO_COEFFC',0);
        TOBOL.setDouble('BLO_COEFFR',0);
        TOBOL.setDouble('BLO_COEFMARG',0);
        TOBOL.PutValue('POURCENTMARQ',0);
      end else
      begin
        if fInfoSSTrait.CoefFG <> 0 then TOBOL.putvalue('BLO_COEFFG',fInfoSSTrait.CoefFG-1);
        if fInfoSSTrait.CoefMarg <> 0 then
        begin
          TOBOL.putvalue('BLO_COEFMARG',fInfoSSTrait.CoefMarg);
          TOBOL.PutValue('POURCENTMARG',Arrondi((TOBOL.GetValue('BLO_COEFMARG')-1)*100,2));
        end;
        RecalcPaPRPV(TOBOL,PrixBloque,FromExcel);
        if TOBOL.GetValue('BLO_PUHT') <> 0 then
        begin
          TOBOL.PutValue('POURCENTMARQ',Arrondi(((TOBOL.GetValue('BLO_PUHT')- TOBOL.GetValue('BLO_DPR'))/TOBOL.GetValue('BLO_PUHT'))*100,2));
        end else
        begin
          TOBOL.PutValue('POURCENTMARQ',0);
        end;
      end;
    end;
  end;
end;

function TPieceCotrait.AfficheSousDetail (TOBL : TOB) : boolean;
begin
	result := ((not fModifSsDetail) and (TOBL.GetValue('GL_TYPEPRESENT') <> DOU_AUCUN)) or
   					((fModifSsDetail) and (TOBL.GetValue('GLC_VOIRDETAIL')='X') );
end;

procedure  TPieceCotrait.AffecteCoTraitanceSousDetail(TOBPere,TOBOUV: TOB;IndiceNomen : Integer;Fournisseur,LibelleFou,NatureTravail : string; Var NextLig : integer; OkINc : boolean);
var TOBOL,TOBL : TOB;
		II,NextII,SS,CurrII : integer;
    IncrementLigne : boolean;
    valeurs : t_valeurs;
begin
	inc(NextLig);
	IncrementLigne := false;
  if AfficheSousDetail (TOBPere) then
  begin
		IncrementLigne := true;
  end;
	for II := 0 to TOBOUV.detail.count -1 do
  begin
		TOBOL := TOBOUV.detail[II];
    TOBOL.putValue('BLO_FOURNISSEUR',Fournisseur);
    TOBOL.putValue('GA_FOURNPRINC',Fournisseur);
    TOBOL.putValue('LIBELLEFOU',LibelleFou);
    TOBOL.putValue('BLO_NATURETRAVAIL',NatureTravail);
    TOBOL.putValue('BLO_DPA',TOBOL.GetValue('BLO_PUHT'));
    TOBOL.putValue('BLO_DPR',TOBOL.GetValue('BLO_PUHT'));
    TOBOL.putValue('BLO_COEFFR',0);
    TOBOL.putValue('BLO_COEFFG',0);
    TOBOL.putValue('BLO_COEFFC',0);
    TOBOL.putValue('BLO_COEFMARG',1);
    TOBOL.putValue('BLO_NONAPPLICFRAIS','X');
    TOBOL.putValue('BLO_NONAPPLICFC','X');
    TOBOL.putValue('BLO_NONAPPLICFG','X');
    if TOBOL.FieldExists('BLF_NATURETRAVAIL') then
    begin
      TOBOL.putValue('BLF_FOURNISSEUR',TOBOL.GetValue('BLO_FOURNISSEUR'));
      TOBOL.putValue('BLF_NATURETRAVAIL',TOBOL.GetValue('BLO_NATURETRAVAIL'));
    end;
    if TOBOL.detail.count > 0 then
    begin
    	AffecteCoTraitanceSousDetail1 (TOBOL,IndiceNomen,Fournisseur,LibelleFou,NatureTravail);
      InitTableau (Valeurs);
      CalculeOuvrageDoc (TOBOL,1,1,true,fDEV,valeurs,(TOBOL.getValue('BLO_FACTUREHT')='X'),true,true);
      SetValeursOuv(TOBOL,valeurs);
    end;
  end;
end;


procedure TPieceCotrait.AffecteSousTraitanceSousDetail(TOBPere,TOBOUV: TOB; IndiceNomen: Integer; Fournisseur, ModePaie,CodeMarche,LibelleFou,NatureTravail: string; var NextLig: integer; OkINc: boolean; PrixBloque,FromExcel : boolean);
var TOBOL : TOB;
		II: integer;
    IncrementLigne : boolean;
    valeurs : T_Valeurs;
begin
  NextLig := NextLig+1;
	IncrementLigne := false;
  if AfficheSousDetail (TOBpere) then
  begin
		IncrementLigne := true;
  end;
  // Calcul des différents sous détails d'ouvrage après affectation
	for II := 0 to TOBOUV.detail.count -1 do
  begin
		TOBOL := TOBOUV.detail[II];
    TOBOL.putValue('BLO_FOURNISSEUR',Fournisseur);
    TOBOL.putValue('GA_FOURNPRINC',Fournisseur);
    TOBOL.putValue('LIBELLEFOU',LibelleFou);
    TOBOL.putValue('BLO_NATURETRAVAIL',NatureTravail);
    TOBOL.putValue('BLO_CODEMARCHE',CodeMarche);
    if TOBOL.fieldExists('BLF_NATURETRAVAIL') then
    begin
      TOBOL.putValue('BLF_FOURNISSEUR',TOBOL.GetValue('BLO_FOURNISSEUR'));
      TOBOL.putValue('BLF_NATURETRAVAIL',TOBOL.GetValue('BLO_NATURETRAVAIL'));
      TOBOL.putValue('BLF_CODEMARCHE',TOBOL.GetValue('BLO_CODEMARCHE'));
    end;
    TOBOL.PutValue('BLO_TPSUNITAIRE',0);
    if TOBOL.detail.count > 0 then
    begin
      AffecteSousTraitanceSousDetail1 (TOBOL,IndiceNomen,Fournisseur,ModePaie,CodeMarche,LibelleFou,NatureTravail,PrixBloque,FromExcel);
      InitTableau (Valeurs);
      CalculeOuvrageDoc (TOBOL,1,1,true,fDEV,valeurs,(TOBOL.getValue('BLO_FACTUREHT')='X'),true,true);
      SetValeursOuv(TOBOL,valeurs);
    end else
    begin
      if TOBOL.getDouble('BLO_DPA')=0 then
      begin
        TOBOL.setDouble('BLO_DPA',TOBOL.getDOuble('BLO_PUHT'));
        TOBOL.setDouble('BLO_DPR',TOBOL.getDOuble('BLO_PUHT'));
        TOBOL.setDouble('BLO_COEFFG',0);
        TOBOL.setDouble('BLO_COEFFC',0);
        TOBOL.setDouble('BLO_COEFFR',0);
        TOBOL.setDouble('BLO_COEFMARG',0);
        TOBOL.PutValue('POURCENTMARQ',0);
      end else
      begin
        if fInfoSSTrait.CoefFG <> 0 then TOBOL.putvalue('BLO_COEFFG',fInfoSSTrait.CoefFG-1);
        if not PrixBloque then
        begin
          if fInfoSSTrait.CoefMarg <> 0 then
          begin
            TOBOL.putvalue('BLO_COEFMARG',fInfoSSTrait.CoefMarg);
            TOBOL.PutValue('POURCENTMARG',Arrondi((TOBOL.GetValue('BLO_COEFMARG')-1)*100,2));
          end;
        end;
        RecalcPaPRPV(TOBOL,prixBloque,FromExcel);
        if TOBOL.GetValue('BLO_PUHT') <> 0 then
        begin
          TOBOL.PutValue('POURCENTMARQ',Arrondi(((TOBOL.GetValue('BLO_PUHT')- TOBOL.GetValue('BLO_DPR'))/TOBOL.GetValue('BLO_PUHT'))*100,2));
        end else
        begin
          TOBOL.PutValue('POURCENTMARQ',0);
        end;
      end;
    end;
    if IncrementLigne then inc(nextLig);
  end;
end;

procedure TPieceCotrait.SetInfoCotraitancelig (TOBL: TOB; Fournisseur,LibelleFou : string; var Indice : integer);
var PrixBloque : boolean;
begin

  if (not IsVariante(TOBL)) and (TOBL.GetValue('GL_ARTICLE') <> '') then
  begin
  	if AddLigneCalcul (TOBL) then
    begin
    	TFfacture(FF).deduitLaLigne (TOBL);
    end;
  end;

  PrixBloque := (TOBL.getValue('GL_BLOQUETARIF')='X');

  TOBL.putvalue('GL_FOURNISSEUR',Fournisseur);
  TOBL.putvalue('LIBELLEFOU',LibelleFou);
  TOBL.putvalue('GLC_NATURETRAVAIL','001');
  if TOBL.getvalue('GL_TYPELIGNE')<> 'ART' then BEGIN inc(Indice); exit; END;
  TOBL.putvalue('GL_DPA',TOBL.getValue('GL_PUHT'));
  TOBL.putvalue('GL_DPR',TOBL.getValue('GL_PUHT'));
//  TOBL.putvalue('GL_BLOQUETARIF','X');
  TOBL.putvalue('GLC_NONAPPLICFC','X');
  TOBL.putvalue('GLC_NONAPPLICFG','X');
  TOBL.putvalue('GLC_NONAPPLICFRAIS','X');
  TOBL.putvalue('GL_COEFFC',0.0);
  TOBL.putvalue('GL_COEFFR',0.0);
  TOBL.putvalue('GL_COEFFG',0.0);
  TOBL.putvalue('GL_COEFMARG',1.0);
  TOBL.putvalue('GL_RECALCULER','X');
  fTOBPiece.putvalue('GP_RECALCULER','X');
  if TOBL.FieldExists('BLF_NATURETRAVAIL') then
  begin
  	TOBL.putvalue('BLF_FOURNISSEUR',TOBL.Getvalue('GL_FOURNISSEUR'));
    TOBL.putvalue('BLF_NATURETRAVAIL',TOBL.Getvalue('GLC_NATURETRAVAIL'));
  end;
  TOBL.PutValue('GL_TPSUNITAIRE',0);
  TFFActure(FF).AfficheLaLigne(TOBL.getIndex);
  if IsOuvrage(TOBL) then
  begin
    AffecteCoTraitanceOuvrage (Fournisseur,LibelleFou,'001',TOBL,Indice,true);
  end else Inc(Indice);
//  if (not IsVariante(TOBL)) and (TOBL.GetValue('GL_ARTICLE') <> '') then TFfacture(FF).AjouteLaLigneCotrait (TOBL);
end;

procedure TPieceCotrait.SetInfoSoustraitancelig(TOBL: TOB; Fournisseur,ModePaie,CodeMarche,LibelleFou: string; var Ligne: integer; PrixBloque,FromExcel : boolean);
var LigneInit : integer;
    CodeFamille2 : string;
    MtPaiementDir : double;
begin
  CodeFamille2 := '';
  LigneInit := Ligne;
  if (not IsVariante(TOBL)) and (TOBL.GetValue('GL_ARTICLE') <> '') and (TOBL.getvalue('GL_TYPELIGNE')= 'ART') then
  begin
  	if AddLigneCalcul (TOBL) then
    begin
    	TFfacture(FF).deduitLaLigne (TOBL);
    end;
  end;
  TOBL.putvalue('GL_FOURNISSEUR',Fournisseur);
  TOBL.putvalue('LIBELLEFOU',LibelleFou);
  TOBL.putvalue('GLC_NATURETRAVAIL','002');
  TOBL.putvalue('GL_CODEMARCHE',CodeMarche);
  if VH_GC.BTCODESPECIF = '001' then
  begin
    CodeFamille2 := GetInfoMarcheST (TOBL.GetString('GL_AFFAIRE'),Fournisseur,CodeMarche,'FAMILLENIV2');
    if CodeFamille2 <> #0 then TOBL.putvalue('GL_FAMILLENIV2',CodeFamille2);
    if TOBL.GetString('GL_NATUREPIECEG')<>'BCE' then
    begin
      MtPaiementDir := GetMtPaiementDir (TOBL.GetString('GL_AFFAIRE'),Fournisseur,CodeMarche);
      if (MtPaiementDir <> 0) then TOBL.SetDouble('GL_DPA',MtPaiementDir);
    end;
  end;

  if TOBL.fieldExists('BLF_NATURETRAVAIL') then
  begin
    TOBL.putvalue('BLF_FOURNISSEUR',TOBL.Getvalue('GL_FOURNISSEUR'));
    TOBL.putvalue('BLF_NATURETRAVAIL',TOBL.Getvalue('GLC_NATURETRAVAIL'));
    TOBL.putvalue('BLF_CODEMARCHE',TOBL.Getvalue('GL_CODEMARCHE'));
  end;

  if TOBL.getvalue('GL_TYPELIGNE')<> 'ART' then BEGIN inc(Ligne); exit; END;

  TOBL.PutValue('GL_TPSUNITAIRE',0);
  TOBL.putvalue('GL_RECALCULER','X');
  fTOBPiece.putvalue('GP_RECALCULER','X');
  begin
    if IsOuvrage(TOBL) then
    begin
      AffecteSousTraitanceOuvrage (TOBL,Fournisseur,LibelleFou,ModePaie,CodeMarche,'002',Ligne,true,PrixBloque,FromExcel);
    end else
    begin
      if TOBL.GetValue('GL_DPA') = 0 then
      begin
        TOBL.SetDouble('GL_DPA',TOBL.GetValue('GL_PUHT'));
        TOBL.SetDOuble('GL_COEFFG',0);
        TOBL.SetDOuble('GL_COEFFC',0);
        TOBL.SetDOuble('GL_COEFFR',0);
        TOBL.SetDOuble('GL_COEFMARG',0);
        TOBL.SetDouble('GL_DPR',TOBL.GetValue('GL_PUHT'));
        TOBL.SetDouble('POURCENTMARQ',0);
      end else
      begin
        if fInfoSSTrait.CoefFG <> 0 then TOBL.putvalue('GL_COEFFG',fInfoSSTrait.CoefFG-1);
        if not prixBloque then
        begin
          if fInfoSSTrait.CoefMarg <> 0 then
          begin
            TOBL.putvalue('GL_COEFMARG',fInfoSSTrait.CoefMarg);
            TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));
          end;
        end;
        RecalcPaPRPV(TOBL,prixBloque,FromExcel);
        if TOBL.GetValue('GL_PUHT') <> 0 then
        begin
          TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('GL_PUHT')- TOBL.GetValue('GL_DPR'))/TOBL.GetValue('GL_PUHT'))*100,2));
        end else
        begin
          TOBL.PutValue('POURCENTMARQ',0);
        end;
      end;
      Inc(Ligne);
    end;
  end;

  TOBL.putvalue('GL_HEURE',Arrondi(TOBL.getValue('GL_TPSUNITAIRE')*TOBL.GetValue('GL_QTEFACT'),V_PGI.okdecQ));
  TOBL.putvalue('GL_TOTALHEURE',TOBL.Getvalue('GL_HEURE'));
  TFFActure(FF).AfficheLaLigne(LigneInit);
end;

procedure TPieceCotrait.AffecteCoTraitanceLigne (Fournisseur: string; var Indice: integer) ;
var TOBL : TOB;
		LibelleFou : string;
begin
  if FF is TFFacture then
  begin
    if (Indice > TOBpiece.detail.count) then begin inc(Indice) ;exit; end;
    TOBL := fTOBPiece.detail[Indice-1]; if TOBL=nil then exit;
    LibelleFou := GetLibIntervenant (Fournisseur);
    if IsSousDetail (TOBL) then
    begin
      AffecteCoTraitanceLigSDetail (TOBL,Fournisseur,LibelleFou,'001');
      Inc(Indice);
      exit;
    end;
    SetInfoCotraitancelig (TOBL,Fournisseur,LibelleFou,Indice);
    //
  end;
end;


procedure TPieceCotrait.AffecteSousTraitanceLigne(Fournisseur,ModePaie,CodeMarche: string;var Ligne: integer);
var TOBL : TOB;
		LibelleFou : string;
  	PrixBloque,FromExcel : boolean;
begin
  if FF is TFFacture then
  begin
    if Ligne > TOBpiece.detail.count then begin inc(Ligne) ;exit; end;
    TOBL := fTOBPiece.detail[Ligne-1]; if TOBL=nil then exit;
    //
  	PrixBloque := (TOBL.getValue('GL_BLOQUETARIF')='X');
    FromExcel := IsFromExcel(TOBL);
    //
    LibelleFou := GetLibIntervenant (Fournisseur);
    if IsSousDetail (TOBL) then
    begin
      AffectesousTraitanceLigSDetail (TOBL,Fournisseur,ModePaie,CodeMarche,LibelleFou,'002');
      Inc(Ligne);
      exit;
    end;
    SetInfoSoustraitancelig (TOBL,Fournisseur,ModePaie,CodeMarche,LibelleFou,Ligne,PrixBloque,FromExcel);
    //
  end;
end;

procedure TPieceCotrait.GSAffectCotrait(Sender: tobject);
var retour : string;
		Indice,IndiceIni : integer;
	  TOBL : TOB;
  	Acol,Arow : integer;
begin
  fTOBligneArecalc.ClearDetail;
	if GS.nbSelected = 0 then BEGIN PGiInfo('Aucune ligne sélectionnée');exit; END;
  Arow := GS.row;
  Acol := GS.col;
  for Indice := 0 to GS.RowCount -1 do
  begin
    if GS.IsSelected(Indice) then
    begin
      TOBL := TOBpiece.detail[Indice-1];
      if IsLignefacture (TOBL) then
      begin
        PGIError('Impossible : Certaines des lignes sont déjà facturés');
        TFFActure(FF).AnnuleSelection;
        exit;
      end;
    end;
  end;
  // Sélection du cotraitant
  retour := AGLLanceFiche('BTP','BTMULAFFAIREINTER','BAI_AFFAIRE='+fTOBAffaire.getValue('AFF_AFFAIRE'),'','');
  // --
  if retour <> '' then
  begin
  	Indice := 1;
    GS.SynEnabled := false;
    GS.BeginUpdate;
  	repeat
      if (GS.IsSelected (Indice) ) and (not IsCommentaire(TOBpiece.detail[Indice-1])) then
      begin
        // controle déjà facturé
        TOBL := GetTOBLigne(TOBpiece,Indice-1);
        if IsLignefacture (TOBL) then
        begin
          TFFActure(FF).AnnuleSelection;
          exit;
        end;
      	IndiceIni := Indice;
        AffecteCoTraitanceLigne (Retour,Indice);
        if FF is TFFActure then
        begin
        	TFFActure(FF).AfficheLaLigne(IndiceIni)
        end else
        begin
        	TFLigNomen(FF).AfficheLaLigne(IndiceIni)
        end;
      end else inc(Indice);
    until Indice > GS.rowCount;

//		if TFFacture(FF).CalculPieceAutorise then
    begin
      for Indice := 0 to fTOBligneArecalc.Detail.count -1 do
      begin
        TOBL := TOB(fTOBligneArecalc.Detail[Indice].Data);
        ZeroLigneMontant (TOBL);
        if AfficheSousDetail (TOBL) then
        begin
          GestionDetailOuvrage(TFFacture(FF),fTOBpiece,TOBL.GetValue('GL_NUMLIGNE'));
        end;
      end;
    end;
    if FF is TFFActure then TFFacture(FF).CalculeLaSaisie(-1,-1,true);
		NettoiePieceTrait(TFFacture(FF).ThePieceTrait);
    SetSaisie;
    GS.EndUpdate;
    GS.SynEnabled := true;
    GS.Invalidate;
  end;
  TFFacture(FF).AnnuleSelection;
  TFFacture(FF).GoToLigne(Acol,Arow); 
end;

procedure TPieceCotrait.MenuEnabled(State: boolean);
var Indice : integer;
		lPOPY : TPopupMenu;
    SS : string;
begin
  lPOPY := Tpopupmenu(TFFacture(FF).Findcomponent('POPY'));
  for Indice := 0 to fMaxItems -1 do
  begin
    SS := fMenuItem [Indice].Name;
    if (VH_GC.BTCODESPECIF = '001')  then
    begin
      if ( Pos(fTOBpiece.GetString('GP_NATUREPIECEG'),'FBT;FBP')>0) or (fMenuItem [Indice].Name <> 'BGESTSOUSTRAIT') then fMenuItem[Indice].visible := State;
    end else
    begin
      fMenuItem[Indice].visible := State;
    end;
  end;
  if fPOPY <> nil then
  begin
    for Indice := 0 to fMaxpopyItems  -1 do
    begin
      SS := fPOPYItem [Indice].Name;
      if (VH_GC.BTCODESPECIF = '001')  then
      begin
        if (( Pos(fTOBpiece.GetString('GP_NATUREPIECEG'),'FBT;FBP')=0) and (fPOPYItem [Indice].Name <> 'PAIESTPOC')) or
           (( Pos(fTOBpiece.GetString('GP_NATUREPIECEG'),'FBT;FBP')>0) and (fPOPYItem [Indice].Name = 'PAIESTPOC')) then fPOPYItem[Indice].visible := State;
      end else
      begin
        if (fPOPYItem [Indice].Name <> 'PAIESTPOC') then fPOPYItem[Indice].visible := State;
      end;
      //
      if fPOPYItem[Indice].Name = 'BCOTRAITTABLEAU' then
      begin
        fPOPYItem[Indice].enabled := (ftobPieceTrait.detail.count > 1);
      end;
    end;
  end;
end;

procedure TPieceCotrait.SetGrilleSaisie (State : boolean);
begin
  TFFacture(FF).BouttonInVisible;
	if SG_NATURETRAVAIL <> -1 then
  begin
    GS.ColEditables[SG_NATURETRAVAIL] := False;
    if not state then
    begin
      GS.Colwidths[SG_NATURETRAVAIL] := -1;
    end else
    begin
      GS.Colwidths[SG_NATURETRAVAIL] := 20;
      GS.ColFormats[SG_NATURETRAVAIL] := 'CB=BTNATURETRAVAIL';
      GS.ColDrawingModes[SG_NATURETRAVAIL]:= 'IMAGE'
    end;
	end;
  if SG_FOURNISSEUR <> -1 then
  begin
  	GS.ColEditables[SG_FOURNISSEUR] := False;
    if not state then
    begin
      GS.Colwidths[SG_FOURNISSEUR] := -1;
    end else
    begin
      GS.Colwidths[SG_FOURNISSEUR] := 90;
    end;
  end;
  if SG_CODEMARCHE <> -1 then
  begin
  	GS.ColEditables[SG_CODEMARCHE] := False;
    if (not state) or (VH_GC.BTCODESPECIF <> '001') then
    begin
      GS.Colwidths[SG_CODEMARCHE] := -1;
    end else
    begin
      GS.Colwidths[SG_CODEMARCHE] := 90;
    end;
  end;
// NOPE  TFFacture(FF).HMTrad.ResizeGridColumns(GS);
  TFFActure(FF).GS.invalidate;
  TFFacture(FF).BouttonVisible(GS.row);
end;

procedure TPieceCotrait.SetSaisie;
var WithCotrait,WithSousTrait,GrilleVisible : Boolean;
begin
	fusable := true; // par défaut
  //
  WithCotrait := IsPieceGerableCoTraitance (fTOBpiece,fTOBaffaire);
  WithSousTrait := IsPieceGerableSousTraitance(TOBpiece);
  if (not WithCotrait ) and (not WithSousTrait) then
  begin
  	fusable := false;
    GrilleVisible := false;
  end else
  begin
    GrilleVisible :=(WithCotrait and FindCoTrait) or (WithSoustrait and FindSousTrait);
  end;
  //
//  PieceTraitUsable := fUsable;
  //
  SetGrilleSaisie(GrilleVisible);
  MenuEnabled (fusable);
  if fusable then Setpopup (WithCotrait,WithSousTrait);
end;

procedure TPieceCotrait.Setaffaire(const Value: TOB);
begin
  fTOBAffaire := Value;
end;

procedure TPieceCotrait.AffecteInterneSousDetail(TOBPere: TOB; IndiceNomen : integer; TOBOUV : TOB; var indice : integer; PrixBloque,FromExcel : boolean);
var TOBOL : TOB;
		II : integer;
    IncrementLigne : boolean;
    valeurs : t_valeurs;
begin
  Indice := Indice +1;
	IncrementLigne := false;
  if AfficheSousDetail (TOBpere) then
  begin
		IncrementLigne := true;
  end;
	for II := 0 to TOBOUV.detail.count -1 do
  begin
		TOBOL := TOBOUV.detail[II];
    TOBOL.putValue('BLO_FOURNISSEUR','');
    TOBOL.putValue('GA_FOURNPRINC','');
    TOBOL.putValue('LIBELLEFOU','');
    TOBOL.putValue('BLO_NATURETRAVAIL','');
    TOBOL.putValue('BLO_DPA',TOBOL.GetValue('BLO_PUHT'));
    TOBOL.putValue('BLO_DPR',TOBOL.GetValue('BLO_PUHT'));
    TOBOL.putValue('BLO_COEFFR',0);
    TOBOL.putValue('BLO_COEFFG',0);
    TOBOL.putValue('BLO_COEFFC',0);
    TOBOL.putValue('BLO_COEFMARG',1);
    TOBOL.putValue('BLO_NONAPPLICFRAIS','-');
    TOBOL.putValue('BLO_NONAPPLICFC','-');
    TOBOL.putValue('BLO_NONAPPLICFG','-');
    if TOBOL.FieldExists('BLF_NATURETRAVAIL') then
    begin
      TOBOL.putvalue('BLF_FOURNISSEUR',TOBOL.Getvalue('BLO_FOURNISSEUR'));
      TOBOL.putvalue('BLF_NATURETRAVAIL',TOBOL.Getvalue('BLO_NATURETRAVAIL'));
    end;
    if TOBOL.detail.count > 0 then
    begin
    	AffecteInterneSousDetail1 (TOBOL,IndiceNomen,PrixBloque,FromExcel);
      InitTableau (Valeurs);
      CalculeOuvrageDoc (TOBOL,1,1,true,fDEV,valeurs,(TOBOL.getValue('BLO_FACTUREHT')='X'),true,true);
      SetValeursOuv(TOBOL,valeurs);
    end else
    begin
    	SetInfoArt (TOBOL,PrixBloque,FromExcel); // on remet à jour via l'article
    end;
    if IncrementLigne then inc(Indice);
  end;
end;


procedure TPieceCotrait.AffecteInterneOuvrage (TOBL : TOB; var indice : integer);
var TOBOUV : TOB;
		IndiceNomen : integer;
    PrixBloque,FromExcel : boolean;
    valeurs : t_valeurs;
begin
	IndiceNomen := TOBL.getvalue('GL_INDICENOMEN');
  PrixBloque := (TOBL.getvalue('GL_BLOQUETARIF')='X');
  FromExcel := IsFromExcel(TOBL);
	if IndiceNomen > 0 then
  begin
    TOBOUV := TOBOuvrage.detail[IndiceNomen-1];
    AffecteInterneSousDetail(TOBL,IndiceNomen,TOBOUV,Indice,PrixBloque,FromExcel);
    InitTableau (Valeurs);
    CalculeOuvrageDoc (TOBOUV,1,1,true,fDEV,valeurs,(TOBL.getValue('GL_FACTUREHT')='X'),true,true);
    SetValeurs(TOBL,valeurs);
  end else Inc(Indice);
end;

procedure TPieceCotrait.InitLigneInterne(TOBL : TOB; var Indice : integer);
var prixBloque,FromExcel : boolean;
		LigneInit : integer;
    TT : TOB;
begin
  LigneInit := indice;
  if (not IsVariante(TOBL)) and (TOBL.GetValue('GL_ARTICLE') <> '') then
  begin
  	if AddLigneCalcul (TOBL) then TFfacture(FF).deduitLaLigne (TOBL);
  end;
  PrixBloque := (TOBL.getValue('GL_BLOQUETARIF')='X');
  FromExcel := IsFromExcel(TOBL);
  if VH_GC.BTCODESPECIF = '001' then
  begin
    TT := FindSousTraitantPOC(TOBL.Getvalue('GL_FOURNISSEUR'),TOBL.Getvalue('GL_CODEMARCHE'));
    if TT <> nil then
    begin
      TT.SetDouble ('NBUSED',TT.GetDouble ('NBUSED') -1);
    end;
  end;
  TOBL.putvalue('GLC_NATURETRAVAIL','');
  TOBL.putvalue('GL_FOURNISSEUR','');
  TOBL.putvalue('GL_CODEMARCHE','');
  if TOBL.FieldExists('BLF_NATURETRAVAIL') then
  begin
    TOBL.putvalue('BLF_FOURNISSEUR',TOBL.Getvalue('GL_FOURNISSEUR'));
    TOBL.putvalue('BLF_NATURETRAVAIL',TOBL.Getvalue('GLC_NATURETRAVAIL'));
    TOBL.putvalue('BLF_CODEMARCHE','');
  end;
  TOBL.putvalue('LIBELLEFOU','');
  if TOBL.getvalue('GL_TYPELIGNE')<> 'ART' then BEGIN Inc(Indice); exit; END;
  TOBL.putvalue('GL_DPA',TOBL.getValue('GL_PUHT'));
  TOBL.putvalue('GL_DPR',TOBL.getValue('GL_PUHT'));
//  TOBL.putvalue('GL_BLOQUETARIF','-');
  TOBL.putvalue('GLC_NONAPPLICFC','-');
  TOBL.putvalue('GLC_NONAPPLICFG','-');
  TOBL.putvalue('GLC_NONAPPLICFRAIS','-');
  TOBL.putvalue('GL_COEFFC',0.0);
  TOBL.putvalue('GL_COEFFR',0.0);
  TOBL.putvalue('GL_COEFFG',0.0);
  TOBL.putvalue('GL_CODEMARCHE','');
  if FromExcel then
  begin
  	TOBL.putvalue('GL_COEFMARG',0);
  end else
  begin
  	TOBL.putvalue('GL_COEFMARG',1.0);
  end;
  //

  if IsOuvrage(TOBL) then
  begin
    AffecteInterneOuvrage (TOBL,Indice);
  end else
  begin
		SetInfoArt (TOBL,PrixBloque,FromExcel); // on remet à jour via l'article
  	Inc(Indice);
  end;
  //
  TOBL.putvalue('GL_RECALCULER','X');
  fTOBPiece.putvalue('GP_RECALCULER','X');

  TOBL.putvalue('GL_HEURE',Arrondi(TOBL.getValue('GL_TPSUNITAIRE')*TOBL.GetValue('GL_QTEFACT'),V_PGI.okdecQ));
  TOBL.putvalue('GL_TOTALHEURE',TOBL.Getvalue('GL_HEURE'));
  TFFActure(FF).AfficheLaLigne(LigneInit);

end;

procedure TPieceCotrait.AffecteInterne (var Ligne : integer);
var TOBL : TOB;
		Fournisseur : string;
    PrixBloque,FromExcel : boolean;
begin
	Fournisseur := '';
  if FF is TFFacture then
  begin
    if Ligne > TOBpiece.detail.count then begin inc(Ligne) ;exit; end;
    TOBL := fTOBPiece.detail[Ligne-1]; if TOBL=nil then exit;
    PrixBloque := (TOBL.GetValue('GL_BLOQUETARIF')='X');
    FromExcel := IsFromExcel(TOBL);
    if IsSousDetail (TOBL) then
    begin
    	AffecteInterneSDetail (TOBL,PrixBloque,FromExcel);
      Inc(Ligne);
      exit;
    end;
    //
    InitLigneInterne(TOBL,Ligne);

  end;
end;

procedure TPieceCotrait.NettoiePieceTrait(TOBpieceTrait : TOB);
var i : integer;
		TOBPT : TOB;
begin
  i:=0;
  if TOBpieceTrait.detail.count = 0 then exit;
  repeat
    TOBPT := TOBPieceTrait.Detail[i];
    if (TOBPT.getvalue('BPE_TOTALHTDEV')=0) and (TOBPT.getvalue('BPE_MONTANTREGL')=0) then
    begin
    	TOBPT.free;
    end else inc(I);
  until i >= TOBPieceTrait.Detail.Count;
end;

procedure TPieceCotrait.GSAffectInterne(Sender: tobject);
var Indice,IndiceIni : integer;
		TOBL : TOB;
    Acol,Arow : integer;
begin
  Arow := GS.row;
  Acol := GS.col;
  fDEV :=  TFfacture(FF).DEV;
  fTOBligneArecalc.ClearDetail;
	if GS.nbSelected = 0 then BEGIN PGiInfo('Aucune ligne sélectionnée');exit; END;

  for Indice := 0 to GS.RowCount -1 do
  begin
    if GS.IsSelected(Indice) then
    begin
      TOBL := TOBpiece.detail[Indice-1];
      (*
      if IsLignefacture (TOBL) then
      begin
        PGIError('Impossible : Certaines des lignes sont déjà facturés');
        TFFActure(FF).AnnuleSelection;
        exit;
      end;
      *)
    end;
  end;

  if PgiAsk ('Désirez-vous réellement affecter ces taches en interne ?','cotraitance') = mryes then
  begin
    GS.SynEnabled := false;
    GS.BeginUpdate;
  	indice := 1;
    repeat
      if (GS.IsSelected (Indice) ) then
      begin
        //
      	IndiceIni := Indice;
        AffecteInterne (Indice);
        if FF is TFFacture then TFFActure(FF).AfficheLaLigne(IndiceIni);
      end else Inc(Indice);
    until indice >= GS.rowCount;

    for Indice := 0 to fTOBligneArecalc.Detail.count -1 do
    begin
      TOBL := TOB(fTOBligneArecalc.Detail[Indice].Data);
      TFfacture(FF).AjouteLaLigne (TOBL);
      if AfficheSousDetail (TOBL) then
      begin
        GestionDetailOuvrage(TFFacture(FF),fTOBpiece,TOBL.GetValue('GL_NUMLIGNE'));
      end;
    end;

    if FF is TFFacture then TFFacture(FF).CalculeLaSaisie(-1,-1,true);
    NettoiePieceTrait (TFFacture(FF).ThePieceTrait );
  	SetSaisie;
    GS.EndUpdate;
    GS.SynEnabled := true;
    GS.Invalidate;
  end;
  TFFacture(FF).AnnuleSelection;
  TFFacture(FF).GoToLigne(Acol,Arow); 
end;

procedure TPieceCotrait.GSCotraittableau(sender: TObject);
begin
	TheTOB := TFFacture(FF).ThePieceTrait;
	//AGLLanceFiche('BTP','BTMULAFFECTEXTE','','','COTRAITANCE;ACTION=MODIFICATION');

	AGLLanceFiche('BTP','BTMULAFFECTEXTE','','','ACTION=MODIFICATION');

  TheTOB := nil;
end;

procedure TPieceCotrait.ReinitDetailOuvrage(TOBO : TOB);
var Indice : Integer;
begin
  TOBO.PutValue('BLO_NATURETRAVAIL','');
  TOBO.PutValue('BLO_FOURNISSEUR','');
  TOBO.putvalue('LIBELLEFOU','');
  if TOBO.FieldExists('BLF_NATURETRAVAIL') then
  begin
    TOBO.putvalue('BLF_FOURNISSEUR',TOBO.Getvalue('BLO_FOURNISSEUR'));
    TOBO.putvalue('BLF_NATURETRAVAIL',TOBO.Getvalue('BLO_NATURETRAVAIL'));
  end;

  if TOBO.Detail.Count > 0 then
  begin
    for Indice := 0 to TOBO.detail.count -1 do
    begin
      ReinitDetailOuvrage(TOBO.detail[Indice]);
    end;
  end;
end;

procedure TPieceCotrait.ReinitOuvrage(TOBL : TOB);
var indice : Integer;
		IndiceNomen : Integer;
    TOBOUV : TOB;
begin
  IndiceNomen := TOBL.GetValue('Gl_INDICENOMEN');
  if IndiceNomen > 0 then
  begin
    TOBOUV := fTOBOuvrage.detail[IndiceNomen-1];
    for Indice := 0 to TOBOUV.Detail.count -1 do
    begin
      ReinitDetailOuvrage (TOBOUV.detail[Indice]);
    end;
  end;
end;

procedure TPieceCotrait.ReinitSaisie;
var indice : integer;
begin
  GS.BeginUpdate;
	for Indice := 0 to fTOBpiece.detail.count -1 do
  begin
  	fTOBpiece.detail[Indice].PutValue('GLC_NATURETRAVAIL','');
  	fTOBpiece.detail[Indice].PutValue('GL_FOURNISSEUR','');
  	fTOBpiece.detail[Indice].PutValue('GL_RECALCULER','');
  	fTOBpiece.detail[Indice].putvalue('LIBELLEFOU','');
    if fTOBpiece.detail[Indice].FieldExists('BLF_NATURETRAVAIL') then
    begin
      fTOBpiece.detail[Indice].putvalue('BLF_FOURNISSEUR',fTOBpiece.detail[Indice].Getvalue('GL_FOURNISSEUR'));
      fTOBpiece.detail[Indice].putvalue('BLF_NATURETRAVAIL',fTOBpiece.detail[Indice].Getvalue('GLC_NATURETRAVAIL'));
    end;

    ZeroLigneMontant (fTOBpiece.detail[Indice]);
    if (IsOuvrage(fTOBpiece.detail[Indice])) and (fTOBpiece.detail[Indice].GetValue('GL_INDICENOMEN')>0) then
    begin
      ReinitOuvrage(fTOBpiece.detail[Indice]);
    end;
  end;
  TFFacture(FF).ReinitPieceForCalc;
  GS.EndUpdate;
  fTOBpiece.putvalue('GP_RECALCULER','X');
  TFFacture(FF).CalculeLaSaisie(-1,-1,false); 
end;

procedure ClearPieceTrait(TOBpieceTrait : TOB);
var TOBS : TOB;
    II : Integer;
begin
  if VH_GC.BTCODESPECIF = '001' then
  begin
    for II := 0 to TOBPieceTrait.detail.count -1 do
    begin
      TOBS := TOBPieceTrait.detail[II];
      TOBS.SetDouble('BPE_TOTALHTDEV',0);
      TOBS.SetDouble('BPE_TOTALTTCDEV',0);
      TOBS.SetDouble('BPE_MONTANTPA',0);
      TOBS.SetDouble('BPE_MONTANTPR',0);
      TOBS.SetDouble('BPE_MONTANTFAC',0);
    end;
  end else
  begin
    TOBPieceTrait.cleardetail;
  end;
end;

procedure TPieceCotrait.GSAffectSoustrait(Sender: tobject);
var Fournisseur,ModePaie : string;
		Indice,IndiceIni : integer;
    TOBParam,TOBL : TOB;
    Acol,Arow : integer;
begin
  fDEv := TFFacture(FF).DEV;
  fTOBligneArecalc.ClearDetail;
	if GS.nbSelected = 0 then BEGIN PGiInfo('Aucune ligne sélectionnée');exit; END;
  Arow := GS.row;
  Acol := GS.col;
  for Indice := 0 to GS.RowCount -1 do
  begin
    if GS.IsSelected(Indice) then
    begin
      TOBL := TOBpiece.detail[Indice-1];
    end;
  end;
  // Sélection du Sous traitant
  TOBParam := TOB.Create ('LES PARAMS',nil,-1);
  TOBParam.AddChampSupValeur ('ACTION','SELECTION');
  TOBParam.AddChampSupValeur ('SOUSTRAIT','');
  TOBParam.AddChampSupValeur ('MODEPAIE','');
  TOBParam.AddChampSupValeur ('AFFAIRE',fTOBpiece.GetValue('GP_AFFAIRE'));
  TOBParam.Data := fTOBPieceInterv;
  fTOBPieceInterv.Data := TOBPieceTrait;
  TheTOB := TOBParam;
  Fournisseur := '';
  ModePaie := '';
  TRY
  	AGLLanceFiche('BTP','BTPIECEINTERVMUL','','','');
    // --
    Fournisseur := TOBParam.GetValue('SOUSTRAIT');
    ModePaie := TOBParam.GetValue('MODEPAIE');
  FINALLY
    fTOBPieceInterv.Data := nil;
    TheTOB := nil;
  	TOBParam.free;
  end;
  if  Fournisseur <> '' then
  begin
  	Indice := 1;
    GS.SynEnabled := false;
    GS.BeginUpdate;
  	repeat
      if (GS.IsSelected (Indice) ) then
      begin
        // controle déjà facturé
        TOBL := GetTOBLigne(TOBpiece,Indice-1);
        (*
        if IsLignefacture (TOBL) then
        begin
          TFFActure(FF).AnnuleSelection;
          exit;
        end;
        *)
      	IndiceIni := Indice;
        AffecteSousTraitanceLigne (Fournisseur ,ModePaie,'',Indice);
        if FF is TFFacture then TFFActure(FF).AfficheLaLigne(IndiceIni);
      end else inc(Indice);
    until Indice > GS.rowCount;

    for Indice := 0 to fTOBligneArecalc.Detail.count -1 do
    begin
      TOBL := TOB(fTOBligneArecalc.Detail[Indice].Data);
      ZeroLigneMontant (TOBL);
      if AfficheSousDetail (TOBL) then
      begin
        GestionDetailOuvrage(TFFacture(FF),fTOBpiece,TOBL.GetValue('GL_NUMLIGNE'));
      end;
    end;
//		NettoiePieceTrait(TFFacture(FF).ThePieceTrait);
    if FF is TFFActure then
    begin
      TFFacture(FF).TheEches.ClearDetail;
//      TFFacture(FF).TheTOBBases.ClearDetail;
    	TFFacture(FF).CalculeLaSaisie(-1,-1,true);

    end;
    SetSaisie;
    GS.EndUpdate;
    GS.SynEnabled := true;
    GS.Invalidate;

  end;
	TFFacture(FF).AnnuleSelection;
  TFFacture(FF).GoToLigne(Acol,Arow); 
end;

function TPieceCotrait.FindSousTraitantPOC(Fournisseur,CodeMarche : string) : TOB;
begin
  result := fTOBPieceInterv.FindFirst(['BPI_TYPEINTERV','BPI_TIERSFOU','BPI_CODEMARCHE'],['Y00',Fournisseur,CodeMarche],true);
end;

procedure TPieceCotrait.AddSousTraitant(Fournisseur,CodeMarche : string);
var TT : TOB;
begin
  TT := TOB.Create('PIECEINTERV',fTOBPieceInterv,-1);
  AddChampsSupTraitCreat (TT);
  AddChampsSupTrait (TT);
  TT.putValue('BPI_NATUREPIECEG',fTOBpiece.GetString('GP_NATUREPIECEG'));
  TT.putValue('BPI_SOUCHE',fTOBpiece.GetString('GP_SOUCHE'));
  TT.putValue('BPI_NUMERO',fTOBpiece.GetInteger('GP_NUMERO') );
  TT.putValue('BPI_INDICEG',fTOBpiece.GetInteger('GP_INDICEG'));
  TT.SetInteger('BPI_ORDRE',fTOBPieceInterv.detail.count -1);
  //
  TT.putValue('BPI_TYPEINTERV','Y00');
  TT.putValue('BPI_TYPEPAIE','001');
  TT.putValue('BPI_DATECONTRAT',V_PGI.DateEntree);
  TT.SetBoolean('BPI_AUTOLIQUID',true);
  TT.SetString('BPI_FAMILLETAXE',GetInfoMarcheST(fTOBpiece.GetString('GP_AFFAIRE'),Fournisseur,CodeMarche,'FAMILLETAXE1'));
  TT.SetString('BPI_TIERSFOU',Fournisseur);
  TT.SetString('BPI_CODEMARCHE',CodeMarche);
  TT.SetDouble('NBUSED',1);
end;


procedure TPieceCotrait.GSAffectSoustraitPOC(Sender: tobject);

  procedure  AjouteSousTraitantPOC(Fournisseur,CodeMarche : string);
  var TT : TOB;
  begin
    TT := fTOBPieceInterv.FindFirst(['BPI_TYPEINTERV','BPI_TIERSFOU','BPI_CODEMARCHE'],['Y00',Fournisseur,CodeMarche],true);
    if TT = nil then
    begin
      AddSousTraitant(Fournisseur,CodeMarche);
    end else
    begin
      TT.SetDouble('NBUSED',TT.GetDouble('NBUSED')+1);
    end;
  end;

var Fournisseur,ModePaie,CodeMarche : string;
		Indice,IndiceIni : integer;
    TOBParam,TOBL : TOB;
    Acol,Arow : integer;
begin
  fDEv := TFFacture(FF).DEV;
  fTOBligneArecalc.ClearDetail;
  ModePaie := '';
	if GS.nbSelected = 0 then BEGIN PGiInfo('Aucune ligne sélectionnée');exit; END;
  Arow := GS.row;
  Acol := GS.col;
  for Indice := 0 to GS.RowCount -1 do
  begin
    if GS.IsSelected(Indice) then
    begin
      TOBL := TOBpiece.detail[Indice-1];
    end;
  end;
  // Sélection du Sous traitant
  TOBParam := TOB.Create ('LES PARAMS',nil,-1);
  TOBParam.AddChampSupValeur ('ACTION','SELECTION');
  TOBParam.AddChampSupValeur ('SOUSTRAIT','');
  TOBParam.AddChampSupValeur ('CODEMARCHE','');
  TOBParam.AddChampSupValeur ('AFFAIRE',fTOBpiece.GetValue('GP_AFFAIRE'));
  TOBParam.Data := fTOBPieceInterv;
  fTOBPieceInterv.Data := TOBPieceTrait;
  TheTOB := TOBParam;
  Fournisseur := '';
  ModePaie := '001';
  TRY
  	AGLLanceFiche('BTP','BTMARCHEST_MUL','BM1_AFFAIRE='+fTOBpiece.GetValue('GP_AFFAIRE'),'','');
    // --
    Fournisseur := TOBParam.GetValue('SOUSTRAIT');
    CodeMarche := TOBParam.GetValue('CODEMARCHE');
  FINALLY
    fTOBPieceInterv.Data := nil;
    TheTOB := nil;
  	TOBParam.free;
  end;
  if  Fournisseur <> '' then
  begin
    AjouteSousTraitantPOC(Fournisseur,CodeMarche);
  	Indice := 1;
    GS.SynEnabled := false;
    GS.BeginUpdate;
  	repeat
      if (GS.IsSelected (Indice) ) then
      begin
        // controle déjà facturé
        TOBL := GetTOBLigne(TOBpiece,Indice-1);
      	IndiceIni := Indice;
        AffecteSousTraitanceLigne (Fournisseur ,ModePaie,CodeMarche,Indice);
        if FF is TFFacture then TFFActure(FF).AfficheLaLigne(IndiceIni);
      end else inc(Indice);
    until Indice > GS.rowCount;

    for Indice := 0 to fTOBligneArecalc.Detail.count -1 do
    begin
      TOBL := TOB(fTOBligneArecalc.Detail[Indice].Data);
      ZeroLigneMontant (TOBL);
      if AfficheSousDetail (TOBL) then
      begin
        GestionDetailOuvrage(TFFacture(FF),fTOBpiece,TOBL.GetValue('GL_NUMLIGNE'));
      end;
    end;
//		NettoiePieceTrait(TFFacture(FF).ThePieceTrait);
    if FF is TFFActure then
    begin
      TFFacture(FF).TheEches.ClearDetail;
//      TFFacture(FF).TheTOBBases.ClearDetail;
    	TFFacture(FF).CalculeLaSaisie(-1,-1,true);

    end;
    SetSaisie;
    GS.EndUpdate;
    GS.SynEnabled := true;
    GS.Invalidate;

  end;
	TFFacture(FF).AnnuleSelection;
  TFFacture(FF).GoToLigne(Acol,Arow); 
end;

procedure TPieceCotrait.Setpopup(WithCotrait, WithSousTrait: boolean);
var MnMenu : TMenuItem;
begin
	if not WithCotrait then
  begin
  	MnMenu := TmenuItem(TFFacture(FF).FindComponent ('BCOTRAITAFFECT'));
    if MnMenu <> nil then MnMenu.Visible := false;
  end;
end;

function  ExistsCotrait(Fournisseur,Affaire : string ) : Boolean;
begin
  Result := false;
  if affaire = '' then exit;
  if ExisteSQL('SELECT BAI_TIERSFOU FROM AFFAIREINTERV WHERE BAI_TIERSFOU="'+
  							Fournisseur+'" AND BAI_AFFAIRE="'+Affaire+'"') then
  begin
    Result := true;
  end;
end;

function ExistsSoustrait (Fournisseur,NaturePiece,Souche : string ; Numero,Indiceg : integer) : Boolean;
var Sql : string;
begin
  Result := false;
  if (fournisseur = '') or (NaturePiece  = '' ) or (Numero=0)  then exit;
  Sql := 'SELECT BPI_TIERSFOU FROM PIECEINTERV WHERE '+
  			 'BPI_TIERSFOU="'+Fournisseur+'" AND '+
  			 'BPI_NATUREPIECEG="'+NaturePiece+'" AND '+
         'BPI_SOUCHE="'+Souche +'" AND '+
         'BPI_NUMERO='+IntToStr(Numero)+' AND '+
         'BPI_INDICEG='+IntToStr(Indiceg);
  if ExisteSQL(SQL) then
  begin
    Result := true;
  end;
end;

function TPieceCotrait.FindSousTrait: Boolean;
var Indice : Integer;
begin
	result := false;
  for Indice := 0 to ftobPieceTrait.detail.count -1 do
  begin
    if ftobPieceTrait.detail[Indice].GetValue('BPE_TYPEINTERV')='Y00' then
    begin
      Result := True;
      break;
    end;
  end;
  if not result then
  begin
    if fTOBPieceInterv.detail.count > 0 then result := true;
  end;
end;

function TPieceCotrait.FindCoTrait: Boolean;
var Indice : Integer;
begin
	result := false;
  for Indice := 0 to ftobPieceTrait.detail.count -1 do
  begin
    if ftobPieceTrait.detail[Indice].GetValue('BPE_TYPEINTERV')='X01' then
    begin
      Result := True;
      break;
    end;
  end;
  if not result then
  begin
     if (fTOBaffaire.getvalue('AFf_AFFAIRE')<>'') and (ExisteCoTraitants(fTOBaffaire.getvalue('AFf_AFFAIRE'))) then
     begin
       result := true;
     end;
  end;
end;

function  TPieceCotrait.PieceTraitUsable : boolean;
var WithCotrait,WithSousTrait : Boolean;
begin
  WithCotrait := IsPieceGerableCoTraitance (fTOBpiece,fTOBaffaire);
  WithSousTrait := IsPieceGerableSousTraitance(fTOBpiece);
  Result := WithCotrait or WithSousTrait;
end;

function PieceTraitUsable (fTOBpiece ,fTOBAffaire : TOB) : Boolean;
var WithCotrait,WithSousTrait,GrilleVisible : Boolean;
		TheCodeAffaire : string;
    fTOBAffaireL : TOB;
begin
  if fTOBAffaire = nil then
  begin
    TheCodeAffaire := fTOBpiece.GetValue('GP_AFFAIRE');
    fTOBAffaireL := FindCetteAffaire (TheCodeAffaire);
    if fTOBAffaireL = nil then
    begin
      StockeCetteAffaire(TheCodeAffaire);
    	fTOBAffaireL := FindCetteAffaire (TheCodeAffaire);
    end;
  end else
  begin
    fTOBAffaireL := fTOBAffaire;
  end;
  WithCotrait := IsPieceGerableCoTraitance (fTOBpiece,fTOBaffaireL);
  WithSousTrait := IsPieceGerableSousTraitance(fTOBpiece);
  Result := WithCotrait or WithSousTrait;
end;

procedure  InitTraitAvoir(TOBSSTrait,TOBOUvrage,TOBL: TOB; ModeGestion : Tmodegestion);

	procedure  InitTraitDetailAV(TOBSSTrait,TOBO : TOB; ModeGestion : TModegestion);
  var Indice : Integer;
  		NatureTravail : double;
  begin
  	for Indice := 0 to TOBO.Detail.count -1 do
    begin
      NatureTravail := Valeur( TOBO.detail[Indice].GetValue('BLO_NATURETRAVAIL'));
      if ((Modegestion=TmgSousTraitance) and (NatureTravail = 2)) then
      begin
        if TOBSSTrait.findFirst(['BPI_TIERSFOU'],[TOBO.detail[Indice].getString('BLO_FOURNISSEUR')],true) = nil then
        begin
          TOBO.detail[Indice].putValue('BLO_NATURETRAVAIL','');
          TOBO.detail[Indice].putValue('BLO_FOURNISSEUR','');
        end;
        if TOBO.detail[Indice].FieldExists('BLF_NATURETRAVAIL') then
        begin
          TOBO.detail[Indice].putvalue('BLF_FOURNISSEUR',TOBO.detail[Indice].Getvalue('BLO_FOURNISSEUR'));
          TOBO.detail[Indice].putvalue('BLF_NATURETRAVAIL',TOBO.detail[Indice].Getvalue('BLO_NATURETRAVAIL'));
        end;

        if TOBO.detail[Indice].Detail.count > 0 then
        begin
          InitTraitDetailAV(TOBSSTrait,TOBO.detail[Indice],ModeGestion);
        end;
      end;
    end;
  end;

var IndiceNomen,Indice : Integer;
		TOBO,TOBOL : TOB;
    NatureTravail : double;
    Affect,WhoAffect : string;
    Multiple : boolean;
begin
	NatureTravail := Valeur( TOBL.GetValue('GLC_NATURETRAVAIL'));
  if ((Modegestion=TmgSousTraitance) and (NatureTravail >0) and ((NatureTravail = 2) or (NatureTravail >= 10))) then
  begin
    if IsOuvrage(TOBL) and (TOBL.GetValue('GL_INDICENOMEN')>0)then
    begin
      IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
      TOBO := TOBOUvrage.detail[IndiceNOmen-1];
      InitTraitDetailav(TOBSSTrait,TOBO,Modegestion);
      Affect := '';
      WhoAffect := '';
      Multiple := false;
      for Indice := 0 to TOBO.detail.count -1 do
      begin
        TOBOL := TOBO.detail[Indice];
        if Indice = 0 then
        begin
          Affect := TOBOL.getValue('BLO_NATURETRAVAIL');
          WhoAffect := TOBOL.getValue('BLO_FOURNISSEUR');
        end;
        if (TOBOL.getValue('BLO_NATURETRAVAIL')<>Affect) or (TOBOL.getValue('BLO_FOURNISSEUR')<>WhoAffect) then Multiple := true;
      end;
      //
      if (Multiple) then
      begin
        TOBL.putvalue('GLC_NATURETRAVAIL','010');
        TOBL.putvalue('GL_FOURNISSEUR','');
        TOBL.putvalue('LIBELLEFOU','');
      end else if (not Multiple) and (Affect='') then
      begin
        TOBL.putvalue('GLC_NATURETRAVAIL','');
        TOBL.putvalue('GL_FOURNISSEUR','');
        TOBL.putvalue('LIBELLEFOU','');
      end else if (not Multiple) and (Affect<>'') then
      begin
        TOBL.putvalue('GLC_NATURETRAVAIL',TOBO.detail[0].getValue('BLO_NATURETRAVAIL'));
        TOBL.putvalue('GL_FOURNISSEUR',TOBO.detail[0].getValue('BLO_FOURNISSEUR'));
        TOBL.putvalue('LIBELLEFOU',TOBO.detail[0].getValue('LIBELLEFOU'));
      end;
    end else
    begin
      if TOBSSTrait.findFirst(['BPI_TIERSFOU'],[TOBL.getString('GL_FOURNISSEUR')],true) = nil then
      begin
        TOBL.PutValue('GLC_NATURETRAVAIL','');
        TOBL.PutValue('GL_FOURNISSEUR','');
        TOBL.PutValue('LIBELLEFOU','');
      end;
    end;
    if TOBL.FieldExists('BLF_NATURETRAVAIL') then
    begin
      TOBL.putvalue('BLF_FOURNISSEUR',TOBL.Getvalue('GL_FOURNISSEUR'));
      TOBL.putvalue('BLF_NATURETRAVAIL',TOBL.Getvalue('GLC_NATURETRAVAIL'));
    end;
  end;
end;

procedure  InitTrait(TOBOUvrage,TOBL: TOB; ModeGestion : Tmodegestion);

	procedure  InitTraitDetail(TOBO : TOB; ModeGestion : TModegestion);
  var Indice : Integer;
  		NatureTravail : double;
  begin
  	for Indice := 0 to TOBO.Detail.count -1 do
    begin
      NatureTravail := Valeur( TOBO.detail[Indice].GetValue('BLO_NATURETRAVAIL'));
      if ((Modegestion=TmgCotraitance) and (NatureTravail = 1)) OR
         ((Modegestion=TmgSousTraitance) and (NatureTravail = 2)) then
      begin
        TOBO.detail[Indice].putValue('BLO_NATURETRAVAIL','');
        TOBO.detail[Indice].putValue('BLO_FOURNISSEUR','');
        if TOBO.detail[Indice].FieldExists('BLF_NATURETRAVAIL') then
        begin
          TOBO.detail[Indice].putvalue('BLF_FOURNISSEUR',TOBO.detail[Indice].Getvalue('BLO_FOURNISSEUR'));
          TOBO.detail[Indice].putvalue('BLF_NATURETRAVAIL',TOBO.detail[Indice].Getvalue('BLO_NATURETRAVAIL'));
        end;

        if TOBO.detail[Indice].Detail.count > 0 then
        begin
          InitTraitDetail(TOBO.detail[Indice],ModeGestion);
        end;
      end;
    end;
  end;

var IndiceNomen,Indice : Integer;
		TOBO,TOBOL : TOB;
    NatureTravail : double;
    Affect,WhoAffect : string;
    Multiple : boolean;
begin
	NatureTravail := Valeur( TOBL.GetValue('GLC_NATURETRAVAIL'));
  if ((Modegestion=TmgCotraitance) and (NatureTravail >0) and ((NatureTravail = 1) or (NatureTravail >= 10))) OR
		 ((Modegestion=TmgSousTraitance) and (NatureTravail >0) and ((NatureTravail = 2) or (NatureTravail >= 10))) then
  begin
    if IsOuvrage(TOBL) and (TOBL.GetValue('GL_INDICENOMEN')>0)then
    begin
      IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
      TOBO := TOBOUvrage.detail[IndiceNOmen-1];
      InitTraitDetail(TOBO,Modegestion);
      Affect := '';
      WhoAffect := '';
      Multiple := false;
      for Indice := 0 to TOBO.detail.count -1 do
      begin
        TOBOL := TOBO.detail[Indice];
        if Indice = 0 then
        begin
          Affect := TOBOL.getValue('BLO_NATURETRAVAIL');
          WhoAffect := TOBOL.getValue('BLO_FOURNISSEUR');
        end;
        if (TOBOL.getValue('BLO_NATURETRAVAIL')<>Affect) or (TOBOL.getValue('BLO_FOURNISSEUR')<>WhoAffect) then Multiple := true;
      end;
      //
      if (Multiple) then
      begin
        TOBL.putvalue('GLC_NATURETRAVAIL','010');
        TOBL.putvalue('GL_FOURNISSEUR','');
        TOBL.putvalue('LIBELLEFOU','');
      end else if (not Multiple) and (Affect='') then
      begin
        TOBL.putvalue('GLC_NATURETRAVAIL','');
        TOBL.putvalue('GL_FOURNISSEUR','');
        TOBL.putvalue('LIBELLEFOU','');
      end else if (not Multiple) and (Affect<>'') then
      begin
        TOBL.putvalue('GLC_NATURETRAVAIL',TOBO.detail[0].getValue('BLO_NATURETRAVAIL'));
        TOBL.putvalue('GL_FOURNISSEUR',TOBO.detail[0].getValue('BLO_FOURNISSEUR'));
        TOBL.putvalue('LIBELLEFOU',TOBO.detail[0].getValue('LIBELLEFOU'));
      end;
    end else
    begin
      TOBL.PutValue('GLC_NATURETRAVAIL','');
      TOBL.PutValue('GL_FOURNISSEUR','');
      TOBL.PutValue('LIBELLEFOU','');
    end;
    if TOBL.FieldExists('BLF_NATURETRAVAIL') then
    begin
      TOBL.putvalue('BLF_FOURNISSEUR',TOBL.Getvalue('GL_FOURNISSEUR'));
      TOBL.putvalue('BLF_NATURETRAVAIL',TOBL.Getvalue('GLC_NATURETRAVAIL'));
    end;
  end;
end;

function GetMontantEntreprisePDir (TOBpieceTrait : TOB) : double;
var TOBPT : TOB;
begin
  Result := 0;
	TOBPT := TOBpieceTrait.FindFirst(['BPE_FOURNISSEUR'],[''],true);
  if TOBPT = nil then Exit;
	Result := TOBPT.getDouble('BPE_MONTANTREGL');
end;

procedure AjoutePaiementDirectMode (var Mode: T_ModeRegl;TOBPiece,TOBTiers,TOBPieceTrait,TOBPieceRG,TOBAcomptes: TOB;DEV: RDevise);
var Indice,II : Integer;
		TOBPT : TOB;
    Mtreste,Xp,XD : double;
    IIENt : integer;
begin
  if not (GetParamSocSecur('SO_BTCPTAPAIEDIRECT',false)) then Exit;
  Mode.NbEche := 0;
  IIEnt := -1;
  for Indice := 0 to TOBPieceTrait.detail.count -1 do
  begin
    TOBPT := TOBPieceTrait.detail[Indice];
		if ((TOBPT.getValue('BPE_TYPEINTERV')='Y00') and(TOBPT.getValue('TYPEPAIE')='001')) or
    	 (TOBPT.getValue('BPE_FOURNISSEUR')='') then
    begin
      if TOBPT.getValue('BPE_FOURNISSEUR')='' then
      begin
        GetSommeAcomptes(TOBAcomptes, XP, XD);
        Mtreste := ARRONDI(TOBPT.GetDouble('BPE_MONTANTREGL') - XD,DEV.Decimale);
        if Mtreste = 0 then Continue;
        //
        if IIEnt = -1 then
        begin
          Inc(Mode.NbEche);
          II := Mode.NbEche;
          Mode.TabEche [II].PourQui := '';
          Mode.TabEche [II].PourQuilib := GetParamSocSecur('SO_LIBELLE','Notre Société');
          Mode.TabEche [II].Pourc := 100;
          Mode.TabEche [II].CodeLettre := 'AL';
          Mode.TabEche [II].MontantD := 0;
          Mode.TabEche [II].MontantP := 0;
          IIent := II;
        end else II := IIEnt;
        Mode.TabEche [II].MontantD := Mode.TabEche [II].MontantD + TOBPT.GetDouble('BPE_MONTANTREGL');
        Mode.TabEche [II].MontantP := Mode.TabEche [II].MontantP + DEVISETOPIVOTEx(TOBPT.GetDouble('BPE_MONTANTREGL'),DEV.Taux,DEV.Quotite,V_PGI.OkDecV);
      end else
      begin
        GetSommeAcomptes(TOBAcomptes, XP, XD,TOBPT.getValue('BPE_FOURNISSEUR'));
        Mtreste := ARRONDI(TOBPT.GetDouble('BPE_MONTANTREGL') - XD,DEV.Decimale);
        if Mtreste = 0 then Continue;
        // Sous traitant en paiement direct
        Inc(Mode.NbEche);
        II := Mode.NbEche;
        Mode.TabEche [II].DateEche := Mode.TabEche [1].DateEche;
        Mode.TabEche [II].ModePaie := GetParamSocSecur('SO_BTPAIESTDIRECTE','');
        Mode.TabEche [II].PourQui := TOBPT.getValue('BPE_FOURNISSEUR');
        Mode.TabEche [II].PourQuilib := TOBPT.getValue('LIBELLE');
        Mode.TabEche [II].MontantD := TOBPT.GetDouble('BPE_MONTANTREGL');
        Mode.TabEche [II].MontantP := DEVISETOPIVOTEx(TOBPT.GetDouble('BPE_MONTANTREGL'),DEV.Taux,DEV.Quotite,V_PGI.OkDecV);
        Mode.TabEche [II].Pourc := 100;
        Mode.TabEche [II].CodeLettre := 'AL';
      end;
    end else if ((TOBPT.getValue('BPE_TYPEINTERV')='Y00') and (TOBPT.getValue('TYPEPAIE')='002')) then
    begin
      if IIEnt = -1 then
      begin
        Inc(Mode.NbEche);
        II := Mode.NbEche;
        Mode.TabEche [II].PourQui := '';
        Mode.TabEche [II].PourQuilib := GetParamSocSecur('SO_LIBELLE','Notre Société');
        Mode.TabEche [II].Pourc := 100;
        Mode.TabEche [II].CodeLettre := 'AL';
        Mode.TabEche [II].MontantD := 0;
        Mode.TabEche [II].MontantP := 0;
        IIent := II;
      end else II := IIEnt;
      Mode.TabEche [II].MontantD := Mode.TabEche [II].MontantD + TOBPT.GetDouble('BPE_MONTANTREGL');
      Mode.TabEche [II].MontantP := Mode.TabEche [II].MontantP + DEVISETOPIVOTEx(TOBPT.GetDouble('BPE_MONTANTREGL'),DEV.Taux,DEV.Quotite,V_PGI.OkDecV);
    end;
  end;
end;


function GetCumulPaiementDirect(TOBPieceTrait : TOB) : double;
var Indice : Integer;
begin
	Result := 0;
  if not (GetParamSocSecur('SO_BTCPTAPAIEDIRECT',false)) then Exit;
  for Indice := 0 to TOBPieceTrait.detail.count -1 do
  begin
		if (TOBPieceTrait.detail[Indice].getValue('BPE_TYPEINTERV')='Y00') and
			 (TOBPieceTrait.detail[Indice].getValue('TYPEPAIE')='001') then
    begin
      Result := Result + TOBPieceTrait.detail[Indice].getDouble('BPE_MONTANTREGL');
    end;
  end;
end;

(*
procedure ConstitueEchesFromPieceTrait(Mode: T_ModeRegl; TOBPiece, TOBEches,TOBpieceTrait,TOBAcomptes : TOB; NbLigAcpte : integer);

	function GetMontantReliquat(TOBPT,TOBAcomptes : TOB) : Double;
  var indice : Integer;
  		MtREgl : double;
  begin
    MtREgl := 0;
    for Indice := 0 to TOBAcomptes.Detail.Count -1 do
    begin
      if (TOBAcomptes.detail[Indice].GetString('GAC_FOURNISSEUR')= TOBPT.GetString('BPE_FOURNISSEUR')) then
      begin
				MtREgl := MTRegl + TOBAcomptes.detail[Indice].GetDouble('GAC_MONTANTDEV');
      end;
    end;
    Result := TOBPT.GetDouble('BPE_MONTANTREGL') - MtRegl;
  end;
var Indice,II : Integer;
		TOBH,TOBPT : TOB;
    DEV : Rdevise;
begin

  if not (GetParamSocSecur('SO_BTCPTAPAIEDIRECT',false)) then Exit;

  DEV.Code := TOBpiece.GetString('GP_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := TOBPiece.GetDouble('GP_TAUXDEV');
  II := 1;
  for Indice := 0 to TOBPieceTrait.detail.count -1 do
  begin
  	TOBPT := TOBpieceTrait.detail[Indice];
		if ((TOBPT.getValue('BPE_TYPEINTERV')='Y00') and
			 (TOBPT.getValue('TYPEPAIE')='001')) or
       (TOBPT.getValue('BPE_FOURNISSEUR') ='') then
    begin
      TOBH := TOB.Create('PIEDECHE', TOBEches, -1);
      TOBH.AddChampSupValeur('LIBELLE','');
      PieceToEcheGC(TOBPiece, TOBH);
      TOBH.PutValue('GPE_NUMECHE', II);
      TOBH.PutValue('GPE_ACOMPTE', '-');
      TOBH.PutValue('GPE_FOURNISSEUR',TOBPT.getValue('BPE_FOURNISSEUR'));
      if TOBPT.getValue('BPE_FOURNISSEUR') ='' then
      begin
      	TOBH.PutValue('GPE_MODEPAIE', Mode.TabEche[1].ModePaie);
      end else
      begin
      	TOBH.PutValue('GPE_MODEPAIE', GetParamSocSecur('SO_BTPAIESTDIRECTE',''));
      	TOBH.SetString('LIBELLE',TOBPT.getValue('LIBELLE'));
      end;
      TOBH.PutValue('GPE_DATEECHE', Mode.TabEche[1].DateEche);
      TOBH.PutValue('GPE_MONTANTDEV', GetMontantReliquat(TOBPT,TOBAcomptes));
      TOBH.PutValue('GPE_MONTANTECHE', DeviseToPivotEx(TOBH.GetDouble('GPE_MONTANTDEV'),DEV.Taux,DEV.quotite,V_PGI.OkdecP));
      inc(II);
    end;
  end;

end;
*)

function GetNbEcheances (TOBpieceTrait : TOB) : integer;
var Indice : Integer;
begin
	Result := 0;
  if not (GetParamSocSecur('SO_BTCPTAPAIEDIRECT',false)) then Exit;
  for Indice := 0 to TOBPieceTrait.detail.count -1 do
  begin
		if (TOBPieceTrait.detail[Indice].getValue('BPE_TYPEINTERV')='Y00') and
			 (TOBPieceTrait.detail[Indice].getValue('TYPEPAIE')='001') then
    begin
      inc(result);
    end;
  end;
end;

function ExisteReglDirect (TOBPieceTrait : TOB) : boolean;
var Indice : Integer;
begin
	Result := false;
  if not (GetParamSocSecur('SO_BTCPTAPAIEDIRECT',false)) then Exit;
  for Indice := 0 to TOBPieceTrait.detail.count -1 do
  begin
		if (TOBPieceTrait.detail[Indice].getValue('BPE_TYPEINTERV')='Y00') and
			 (TOBPieceTrait.detail[Indice].getValue('TYPEPAIE')='001') then
    begin
      result := True;
      break;
    end;
  end;
end;

function IsComptabilisationInterv (TOBPieceTrait : TOB) : boolean;
var Indice : Integer;
begin
	Result := false;
  if TOBPieceTrait = nil then Exit;
  if TOBPieceTrait.detail.count = 0 then Exit;
  for Indice := 0 to TOBPieceTrait.detail.count -1 do
  begin
		if (TOBPieceTrait.detail[Indice].getValue('BPE_FOURNISSEUR') <> '') and
    	 (TOBPieceTrait.detail[Indice].getValue('BPE_TYPEINTERV')<>'X00') then
    begin
      result := True;
      break;
    end;
  end;
end;

function GetModePaie (TOBPiece: TOB; Affaire,Fournisseur : string) : string;
var TOBPieceInterv : TOB;
		TOBAffaireInterv : TOB;
    TOBT : TOB;
    QQ : TQuery;
    cledoc : R_CLEDOC;
begin
  Result := '000';
	if Fournisseur = '' then Exit;
  TOBPieceInterv := TOB.Create ('LES SSTRAITS',nil,-1);
  TOBAffaireInterv := TOB.Create ('LES COTRAITS',nil,-1);
  try
		CleDoc:=TOB2CleDoc(TOBPiece);
    QQ := OpenSQL('SELECT BPI_TIERSFOU,BPI_TYPEPAIE FROM PIECEINTERV WHERE '+WherePiece(cledoc,ttdPieceInterv,True),True,-1,'',true);
    if not QQ.eof then TOBPieceInterv.LoadDetailDB ('PIECEINTERV','','',QQ,false);
    ferme (QQ);
    if Affaire <> '' then
    begin
      QQ := OpenSQL('SELECT BAI_TIERSFOU,(SELECT AFF_TYPEPAIE FROM AFFAIRE WHERE AFF_AFFAIRE=BAI_AFFAIRE) AS TYPEPAIE FROM AFFAIREINTERV WHERE BAI_AFFAIRE="'+Affaire+'"',True,-1,'',True);
      if not QQ.eof then TOBAffaireInterv.LoadDetailDB ('AFFAIREINTERV','','',QQ,false);
      ferme (QQ);
    end;

    TOBT := TOBAffaireInterv.FindFirst(['BAI_TIERSOU'],[Fournisseur],true);
    if TOBT <> nil then
    begin
      result := TOBT.GetValue('TYPEPAIE');
    end else
    begin
    	TOBT := TOBPieceInterv.FindFirst(['BPI_TIERSOU'],[Fournisseur],true);
      if TOBT <> nil then result := TOBT.GetValue('BPI_TYPEPAIE');
    end;
  finally
    TOBPieceInterv.Free;
    TOBAffaireInterv.free;
  End;

end;

function GetPaiementSSTrait (TOBSStrait : TOB; Intervenant : string) : string;
var TOBT : TOB;
begin
  result := '000';
  if TOBSStrait = nil then Exit;
	TOBT := TOBSSTrait.findFirst(['BPI_TIERSFOU'],[Intervenant],true);
  if TOBT = nil then exit;
  result := TOBT.GetString('BPI_TYPEPAIE');
end;

function TPieceCotrait.ExisteCoTraitants(CodeAffaire: string): boolean;
begin
	result := ExisteSQL('SELECT BAi_AFFAIRE FROM AFFAIREINTERV WHERE BAI_AFFAIRE="'+CodeAffaire+'"');
end;

Function IsPieceIntervenant(TOBT : TOB; Cledoc : R_Cledoc) : Boolean;
Var indice : Integer;
begin

  Result := False;

  //contrôle si document soumis à cotraitance...
  if TobT = nil then
    Result := ExisteSQL('SELECT BPE_TYPEINTERV FROM PIECETRAIT WHERE BPE_TYPEINTERV="X01" AND ' +
              WherePiece(CleDoc, ttdPieceTrait, False))
  else
  begin
    for Indice := 0 to TobT.detail.count -1 do
    begin
      if TobT.detail[Indice].GetValue('BPE_TYPEINTERV')='X01' then
      begin
        Result := True;
        break;
      end;
    end;
  end;


end;



procedure TPieceCotrait.SetInfoArt(TOBL : TOB; PrixBloque,FromExcel : boolean);
var TOBART : TOB;
		prefixe,RefSais,VenteAchat : string;
    Qart : TQuery;
    DEV : Rdevise;
begin
  DEV.code := '';
  prefixe := GetPrefixeTable (TOBL);
  RefSais := TOBL.GetValue(prefixe+'_ARTICLE');
  VenteAchat := GetInfoParPiece(fTOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_VENTEACHAT');
  if IsOuvrage (TOBL) then exit;
	// recup des prix achat
  TOBART := TOBARticles.findfirst(['GA_ARTICLE'],[RefSais],true);
  if TOBART = nil then
  begin
    QArt := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                   'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                   'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+RefSais+'"',true,-1, '', True);
//
    if Not QArt.EOF then
    begin
    	TOBART := TOB.Create ('ARTICLE',TOBArticles,-1);
      TOBART.SelectDB('',QArt);
    end;
    Ferme(QArt);
  end;
  if TOBArt = nil then exit;
  if TOBArt.FieldExists ('DEJA CALCULE') then TOBArt.AddChampSupValeur ('DEJA CALCULE','-',false);
  InitValoArtNomen (TOBART,venteAchat); // histoire de récup les valorisation a jour
  if VenteAchat = 'VEN' then
  begin
    if prefixe = 'GL' then
    begin
      TOBL.putValue('GLC_NONAPPLICFRAIS','-');
      TOBL.putValue('GLC_NONAPPLICFC','-');
    end else if prefixe = 'BLO' then
    begin
      TOBL.putValue('BLO_NONAPPLICFRAIS','-');
      TOBL.putValue('BLO_NONAPPLICFC','-');
    end;
    //
//    TOBL.putValue('GLC_NONAPPLICFRAIS','-');
    if IsprestationST (TOBL) then
    begin
      if prefixe = 'GL' then
      begin
        if fTOBPIece.getValue('GP_APPLICFGST')='-' then TOBL.putValue('GLC_NONAPPLICFRAIS','X') else TOBL.putValue('GLC_NONAPPLICFRAIS','-');
        if fTOBPIece.getValue('GP_APPLICFCST')='-' then TOBL.putValue('GLC_NONAPPLICFC','X') else TOBL.putValue('GLC_NONAPPLICFC','-');
      end else if prefixe = 'BLO' then
      begin
        if fTOBPIece.getValue('GP_APPLICFGST')='-' then TOBL.putValue('BLO_NONAPPLICFRAIS','X') else TOBL.putValue('BLO_NONAPPLICFRAIS','-');
        if fTOBPIece.getValue('GP_APPLICFCST')='-' then TOBL.putValue('BLO_NONAPPLICFC','X') else TOBL.putValue('BLO_NONAPPLICFC','-');
      end;
    end;
  	TOBL.putValue(prefixe+'_DPA',TOBART.GetDouble('GA_PAHT'));
  end;
  if TOBART.GetValue('GA_DPRAUTO')='X' then
  begin
    TOBL.PutValue(prefixe+'_COEFFG',TobART.GetValue('GA_COEFFG')-1);
  end else TOBL.PutValue(prefixe+'_COEFFG',0);

  if TOBArt.GetValue('GA_CALCAUTOHT')='X' then
  begin
    TOBL.PutValue(prefixe+'_COEFMARG', TOBArt.GetValue('GA_COEFCALCHT'));
    TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue(prefixe+'_COEFMARG')-1)*100,2));
  end else TOBL.PutValue(prefixe+'_COEFMARG',0);

  // recup des coef si domaine d'activite
  if fTOBPiece.getValue('GP_DOMAINE')<>'' then
  begin
    TOBL.putvalue(prefixe+'_DOMAINE',fTOBPiece.getValue('GP_DOMAINE'));
  end;
  if TOBART.getValue('GA_PRIXPASMODIF')<>'X' Then
  begin
    if prefixe = 'GL' then AppliqueCoefDomaineLig (TOBL)
    else if prefixe = 'BLO' then AppliqueCoefDomaineActOuv (TOBL);
    if VH_GC.BTCODESPECIF = '001' then ApliqueCoeflignePOC (TOBL,DEV);
  end;
  RecalcPaPRPV(TOBL,PrixBloque,FromExcel);
  if TOBL.GetValue(prefixe+'_PUHT') <> 0 then
  begin
    TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue(prefixe+'_PUHT')- TOBL.GetValue(prefixe+'_DPR'))/TOBL.GetValue(prefixe+'_PUHT'))*100,2));
  end else
  begin
    TOBL.PutValue('POURCENTMARQ',0);
  end;
end;

procedure TPieceCotrait.RecalcPaPRPV(TOBL: TOB; PrixBloque,FromExcel : boolean);
var coeffg : double;
		prefixe : string;
begin
  prefixe := GetPrefixeTable (TOBL);

  if TOBL.Getvalue(prefixe+'_COEFMARG')=0 then TOBL.putvalue(prefixe+'_COEFMARG',1);
  TOBL.PutValue(prefixe+'_DPR',Arrondi(TOBL.getValue(prefixe+'_DPA')*(1+TOBL.Getvalue(prefixe+'_COEFFG')),V_PGI.okdecP));
  if prixBloque or FromExcel then
  begin
    if TOBL.getValue(prefixe+'_PUHTDEV') = 0 then
    begin
    	TOBL.Putvalue(prefixe+'_COEFMARG',0);
    end else
    begin
    	TOBL.Putvalue(prefixe+'_COEFMARG',arrondi(TOBL.getValue(prefixe+'_DPR')/TOBL.getValue(prefixe+'_PUHTDEV'),4));
      TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue(prefixe+'_COEFMARG')-1)*100,2));
    end;
  end else
  begin
  	TOBL.PutValue(prefixe+'_PUHT',Arrondi(TOBL.getValue(prefixe+'_DPR')*TOBL.Getvalue(prefixe+'_COEFMARG'),V_PGI.okdecP));
    TOBL.putValue(prefixe+'_PUHTDEV',Pivottodevise(TOBL.getValue(prefixe+'_PUHT'),fDEV.Taux,fDEV.Quotite,V_PGI.okdecP));
  end;
end;

procedure TPieceCotrait.AffecteInterneSousDetail1(TOBOUV: TOB;IndiceNomen: integer; PrixBloque,FromExcel: boolean);
var II : integer;
		TOBOL : TOB;
    valeurs : t_valeurs;
begin
	for II := 0 to TOBOUV.detail.count -1 do
  begin
		TOBOL := TOBOUV.detail[II];
    TOBOL.putValue('BLO_FOURNISSEUR','');
    TOBOL.putValue('BLO_CODEMARCHE','');
    TOBOL.putValue('GA_FOURNPRINC','');
    TOBOL.putValue('LIBELLEFOU','');
    TOBOL.putValue('BLO_NATURETRAVAIL','');
    TOBOL.putValue('BLO_DPA',TOBOL.GetValue('BLO_PUHT'));
    TOBOL.putValue('BLO_DPR',TOBOL.GetValue('BLO_PUHT'));
    TOBOL.putValue('BLO_COEFFR',0);
    TOBOL.putValue('BLO_COEFFG',0);
    TOBOL.putValue('BLO_COEFFC',0);
    TOBOL.putValue('BLO_COEFMARG',1);
    TOBOL.putValue('BLO_NONAPPLICFRAIS','-');
    TOBOL.putValue('BLO_NONAPPLICFC','-');
    TOBOL.putValue('BLO_NONAPPLICFG','-');
    if TOBOL.FieldExists('BLF_NATURETRAVAIL') then
    begin
      TOBOL.putvalue('BLF_FOURNISSEUR',TOBOL.Getvalue('BLO_FOURNISSEUR'));
      TOBOL.putvalue('BLF_NATURETRAVAIL',TOBOL.Getvalue('BLO_NATURETRAVAIL'));
      TOBOL.putvalue('BLF_CODEMARCHE','');
    end;
    if TOBOL.detail.count > 0 then
    begin
    	AffecteInterneSousDetail1 (TOBOL,IndiceNomen,PrixBloque,FromExcel);
      InitTableau (Valeurs);
      CalculeOuvrageDoc (TOBOL,1,1,true,fDEV,valeurs,(TOBOL.getValue('BLO_FACTUREHT')='X'),true,true);
      SetValeursOuv(TOBOL,valeurs);
    end else
    begin
    	SetInfoArt (TOBOL,PrixBloque,FromExcel); // on remet à jour via l'article
    end;
  end;

end;

procedure TPieceCotrait.GSGestionSoustrait(Sender: Tobject);

    procedure MiseAjourPieceTrait (TOBpieceTrait,TOBPieceInterv : TOB);
    var Indice : integer;
    		TOBPT : TOB;
    begin
			For Indice := 0 to TOBPieceInterv.detail.count -1 do
      begin
				TOBPT := TOBpieceTrait.FindFirst(['BPE_FOURNISSEUR'],[TOBPieceInterv.detail[Indice].getString('BPI_TIERSFOU')],true);
        if TOBPT <> nil then
        begin
//          TOBPT.SetString('BPI_TYPEPAIE',TOBPieceInterv.detail[Indice].getString('BPI_TYPEPAIE'));
          TOBPT.SetString('TYPEPAIE',TOBPieceInterv.detail[Indice].getString('BPI_TYPEPAIE'));
        end;
      end;
    end;

var TOBparam : TOB;
begin
  TOBParam := TOB.Create ('LES PARAMS',nil,-1);
  TOBParam.AddChampSupValeur ('ACTION','MODIFICATION');
  TOBParam.AddChampSupValeur ('SOUSTRAIT','');
  TOBParam.AddChampSupValeur ('MODEPAIE','');
  TOBParam.AddChampSupValeur ('AFFAIRE',fTOBpiece.GetValue('GP_AFFAIRE'));
  TOBParam.Data := fTOBPieceInterv;
  fTOBPieceInterv.Data := TOBPieceTrait;
  TheTOB := TOBParam;
  TRY
  	AGLLanceFiche('BTP','BTPIECEINTERVMUL','','','');
  FINALLY
    fTOBPieceInterv.Data := nil;
    TheTOB := nil;
  	TOBParam.free;
    MiseAjourPieceTrait (TOBpieceTrait,fTOBPieceInterv);
    TOBpiece.putvalue('GP_RECALCULER','X');
    if FF is TFFacture then
    begin
      TFFacture(FF).TheEches.ClearDetail;
      TFFacture(FF).CalculeLaSaisie(-1,-1,true);
    end;
    NettoiePieceTrait (TFFacture(FF).ThePieceTrait );
    CalculeReglementsIntervenants (TFFacture(FF).TheTOBSSTRAIT,
    															 TOBpiece,TFFacture(FF).TheTOBPieceRG,
                                   TFFacture(FF).TheAcomptes,
                                   TFFacture(FF).TheTOBPorcs,
                                   fTOBpiece.GetValue('GP_AFFAIRE'),
                                   TOBPieceTrait,fDEV,false);
  end;
end;

procedure TPieceCotrait.GSPaieSTPOC(Sender: TObject);
var TOBParam : TOB;
    ok : boolean;
begin
  TOBParam := TOB.Create('LES PARAMS',nil,-1);
  TRY
    TOBParam.AddChampSupValeur('MODIF','-');
    TOBParam.AddChampSupValeur('NETAPAYER',TFFActure(FF).NETAPAYERDEV.Value);
    TOBParam.AddChampSupValeur('NBDEC',TFFActure(FF).DEV.decimale);
    TheTOB := TOBParam;
    TOBParam.Data := ftobPieceTrait;
    TOBpieceTrait.data := fTOBPieceInterv;
    AGLLanceFiche('BTP','BTSAISPAIMENTPOC','','',ActionToString(TFFActure(FF).TheAction));
    TheTOB := nil;
    TOBpieceTrait.data := nil;
    if TOBParam.getstring('MODIF')='X' then
    begin
      TFFActure(FF).RecalculeEcheances;
    end;
  FINALLY
  TOBParam.free;
  end;
//
end;

{ TInfoSSTrait }

constructor TInfoSSTrait.create;
begin
  fCode :='';
  fCoefFG :=0;
  fCoefMarg :=0;
end;

destructor TInfoSSTrait.destroy;
begin
  inherited;
end;

procedure TInfoSSTrait.SetCode(const Value: string);
var QQ : Tquery;
		SQl : string;
begin
  if Value = '' then exit;
  fcode := Value;
  Sql := 'SELECT * FROM NATUREPREST WHERE BNP_NATUREPRES="'+fCode+'"';
  QQ := OpenSql (Sql,true,1,'',true);
  if not QQ.eof then
  begin
		fCoefFG := QQ.findField('BNP_COEFCALCPR').AsFloat;
		fCoefMarg  := QQ.findField('BNP_COEFCALCHT').AsFloat;
  end;
  ferme (QQ);
end;

function IsMultiIntervRg (TOBpieceTrait: TOB) : boolean;
var TOBT : TOB;
		TypeInterv : string;
    Indice : integer;
    Interv : string;
begin
  result := (TOBpieceTrait.detail.Count > 1);
  (*
	for Indice := 0 to TOBpieceTrait.detail.count -1 do
  begin
    TOBT := TOBpieceTrait.detail[Indice];
    Interv := TOBT.GetString ('BPE_FOURNISSEUR');
    if Interv <> '' then
    begin
      TypeInterv := TOBT.GetString('BPE_TYPEINTERV');
      if TypeInterv = 'Y00' then // sous traitance
      begin
				if TOBT.GetString('TYPEPAIE') = '001' then  // paiement direct
				begin
        	result := true;
          exit;
        end;
      end else
      begin
        result := true;
        exit;
      end;
    end;
  end;
  *)
end;

FUNCTION FindRecapCotrait(Affaire : string) : boolean;
var StSql : string;
begin
  StSQL := 'SELECT BPE_AFFAIRE '+
           'FROM PIECETRAIT  WHERE BPE_AFFAIRE="'+Affaire+'"';
  result := ExisteSql (StSql);
end;


procedure ReinitReglPieceTrait(TOBpiecetrait : TOB);
var ii : integer;
begin
  if TOBpieceTrait = nil then exit;
  if TOBPieceTrait.detail.count = 0 then exit;
  for II := 0 to TOBPieceTrait.detail.count -1 do
  begin
    TOBPieceTrait.detail[II].SetBoolean('BPE_REGLSAISIE',false);
  end;
end;


end.

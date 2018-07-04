unit PiecesRecalculs;

interface

Uses Classes,
     Windows,
     sysutils,
     ComCtrls,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$ELSE}
     MainEagl,
{$ENDIF}
     HCtrls,
     Hpanel,
     HEnt1,
     HMsgBox,
     Vierge,
     EntGc,
     BTPUtil,
     ExtCtrls,
     factTOB,
     NomenUtil,
     SaisUtil,
     UtilPgi,
     FactTvaMilliem,
     UTOB,uEntCommun;

type

  TRResult = (TrrOk,TrrErrBeforeCalc,TrrErrCalcPR,TrrErrCalcPV,TrrErrEcriture);
  TTypeInfo = (TTiQte,TTiPua,TtiPuv,TTiUnite);

  RInfoLigne = record
  	Ligne : integer;
    UniqueBlo : integer;
  end;

	TRecalculPiece = class
  	private
    	TOBPiece,TOBPorcs,TOBOuvrages,TOBOuvragesP : TOB;
      TOBArticles,TOBTiers,TOBPieceRG,TOBPieceRG_O,TOBBasesRG,TOBBases,TOBBasesL,TOBAffaireInterv,TOBVTECOLL : TOB;
      TOBaffaire : TOB;
      TOBPieceTrait : TOB;
      TOBSSTRAIT,TOBAcomptes : TOB;
      TOBEChes : TOB;
      TOBCPTA : TOB;
      TOBanaP,TOBanaS : TOB;
      MontantFraisChantier,CoeffR : double;
      CalculPv,fReajustePVOuv : boolean;
      EnHt : boolean;
      DEV : RDevise;
      ArtEcart : string;
      CodeRetour : TRResult;
    	ApplicFRPiece,ApplicFCPiece : string;
      Cledoc : r_cledoc;
      VenteAchat, Nature : string;
      CalculViaAchat : boolean;
      // Pour sauvegarde du contexte avant calcul
			PieceInterne,OuvrageInterne,PortInterne : TOB;
  		ArticlesInterne,TiersInterne,PieceRGInterne,BasesRGInterne,BasesInterne : TOB;
      TheRepartTva : TREPARTTVAMILL;
      fNewCodeTva : string;
      OldEcr, OldStk: RMVT;
      // ---
      procedure LoadLesTObs;
    	procedure ChargeLesOuvrages;
      procedure CreerLesTOBOuv(TOBGroupeNomen, TOBNomen: TOB; LaLig,idep: integer);
      procedure ReInitMontantPA(TOBL: TOB);
    	procedure ReInitMontantPiece;

    	procedure SetPiece(const Value: TOB);
      procedure RecupCoefFG;
      procedure TraitementOuvrage (TOBL : TOB;WithPuv : Boolean=false);
    	procedure SommationLignePiece(TOBL, TOBPiece: TOB);
      procedure ChargeLesArticles;
      procedure ChargeTiers;
      procedure detruitpiece ;
      procedure validelaPiece;
      procedure EcritLapiece;
   		procedure RecalculePAPRPV;
      //
      procedure RestitueContexte;
      procedure SauvegardeContexte;
    	procedure DefiniContexteCalcul(LaTOBPiece, LaTOBOuvrage, LesPorts,LesArticles, LeTiers, LEsPieceRG,
      															 LesBasesRG, LesBases: TOB);
      procedure InitPiece;
      function FindInPiece(InfoLigne: RinfoLigne): TOB;
    	procedure ReajusteSousDetailOuvrage(TOBL, TOBO: TOB;var LeMontant: double);
    	procedure ReajusteSousDetail(TOBL: TOB; var LeMontant: double);
    	function FindOuvrage(TOBL: TOB; InfoLigne: RinfoLigne): TOB;
      procedure ControleEtReajusteOuvrages;
			procedure ControleEtReajusteSousDetail (TOBL : TOB);
    	procedure ReajusteSousDetailPV(TOBL: TOB; var LeMontant : Double; MtBase: double);
      function IsDejaFacture (TOBL : TOB) : Boolean;
      procedure EcritSituation;
      //
      public

    	property Result : TRResult read CodeRetour;
      //
      property AffecteNewTva : string read fNewCodeTva write fNewCodeTva;
    	property LaPiece : TOB read TOBPiece write SetPiece;
      property RecalcPv : boolean write CalculPv;
      property ReajustePVOuv : boolean read fReajustePVOuv write fReajustePVOuv ;
    	constructor create;
      destructor destroy; override;
      procedure lanceTraitement;
      //
    	procedure RecalculPieceFromSaisie(LaTOBPiece, LaTOBOuvrage, LesPorts, LesArticles, LeTiers,
      																	LEsPieceRG, LesBasesRG, LesBases: TOB);
      // -- FONCTIONS PERMETTANT DE METTRE A JOUR LES INFOS LIGNES D'UNE PIECE AVANT DE LA RECALCULER ET MAJ
      procedure CreateEnv;
      procedure ReInitEnv;
      procedure DestroyEnv;
    	function ChargeLapiece(lacledoc: r_cledoc): integer;
      function MajInfoLigne (InfoLigne: RinfoLigne; TypeInfo : TTypeInfo; valeur : variant) : integer;
    	procedure CalculeLaPiece;
    	procedure EnregistreLaPiece;
    	procedure ReCalculeLaPiece;
      //
  end;

procedure GereTraitementOuvrage (TOBOuvrage : TOB; var Valeur : T_Valeurs; DEV : Rdevise; CalculPv : boolean=false; RazFg : boolean=false);
procedure AppliqueFraisPiece (TOBPiece,TOBOuvrages : TOB; CalculViaAchat,CalculPV,EnHt : boolean; DEV : Rdevise; WithCalcCoefFG : boolean=false);


function PieceNonOptimiseCalcul (TOBpiece,TOBouvrage,TOBBasesL : TOB) : boolean;
function PieceNonMiseAJourOptimise (TOBpiece,TOBBAsesL : TOB) : boolean;
function TraitementRecalculPiece (Cledoc : r_cledoc) : TRResult; overload;
function TraitementRecalculPiece (TOBPiece : TOB; CalculPv ,ReajustePvOuv: boolean) : TRResult; overload;
function RecalculePieceFromSaisie (LaTOBPiece,LaTOBOuvrage,LesPorts,LesArticles,LeTiers,LEsPieceRG,LesBasesRG,LesBases: TOB; CalculPV : boolean=false) : TRResult;
Function CalculMtFraisFromLigne (TOBpiece, TOBL : TOB; PUA, CoefFG : Double ; var MontantFG, MontantFR, MontantFC : Double; CalculPv : Boolean=True) : Double;
procedure CalculFraisFromLigne (TOBPiece,TOBL : TOB; CalculPv : boolean=true);                                                                                
procedure CalculeBaseFraisLigne (TOBL : TOB; CalculViaAchat : boolean=true);
procedure SupprimeFraisRepartisPiece (TOBPorcs : TOB);
procedure  SupprimeFraisDetaille (TOBpiece : TOB);
procedure CalculeLigneAc (TOBL : TOB ; DEV : Rdevise;WithPuv : boolean=true);
function AffecteTvaPiece (TOBPiece : TOB; NewTva : string) : TRResult;

implementation

uses factcomm,FactOuvrage,FactUtil,FactCalc,
		 FactSpec,FactVariante,UtilTOBPiece,UcoTraitance,
     FactureBTP,ParamSoc,UtilArticle,FactRg,factCpta,UCumulCollectifs
     ;


procedure CalculeLigneAc (TOBL : TOB ; DEV : Rdevise;WithPuv : boolean=true);
var MontantCharge,Qte,QteDudetail : double;
    prefixe : string;
begin
 // if TOBL.nomtable <> 'LIGNE' Then exit;
 prefixe := GetprefixeTable (TOBL);
 if Prefixe = '' then exit;
 if (GetInfoParPiece(TOBL.getValue(prefixe+'_NATUREPIECEG'), 'GPP_VENTEACHAT') <> 'VEN') then Exit; // uniquement pour les documents de vente
 Qte := TOBL.GetValue(prefixe+'_QTEFACT');
 QteDuDetail := TOBL.GetValue(prefixe+'_PRIXPOURQTE'); if QteDuDetail = 0 then QteDuDetail := 1;
 TOBL.putValue(prefixe+'_MONTANTPA',Arrondi((TOBL.GetValue(prefixe+'_QTEFACT')*TOBL.GetValue(prefixe+'_DPA'))/QteDudetail,V_PGI.okDecP));
 //
 TOBL.putValue(prefixe+'_MONTANTPAFG',TOBL.GetValue(prefixe+'_MONTANTPA'));
 TOBL.putValue(prefixe+'_MONTANTPAFC',TOBL.GetValue(prefixe+'_MONTANTPA'));
 TOBL.putValue(prefixe+'_MONTANTPAFR',TOBL.GetValue(prefixe+'_MONTANTPA'));
 TOBL.putValue(prefixe+'_MONTANTFG',0);
 TOBL.putValue(prefixe+'_MONTANTFC',0);
 TOBL.putValue(prefixe+'_MONTANTFR',0);
 //
 if Prefixe = 'GL' then
 begin
   TOBL.putValue(prefixe+'_MONTANTFG',arrondi(TOBL.GetValue(prefixe+'_MONTANTPA')*TOBL.GetValue(prefixe+'_COEFFG'),4));
   MontantCharge := TOBL.GetValue(prefixe+'_MONTANTPA')+TOBL.GetValue(prefixe+'_MONTANTFG');
   //
   if TOBL.GetValue('GLC_NONAPPLICFC')<>'X' then TOBL.putValue(prefixe+'_MONTANTFC',arrondi(MontantCHarge*TOBL.GetValue(prefixe+'_COEFFC'),4));
   MontantCharge := TOBL.GetValue(prefixe+'_MONTANTPA')+TOBL.GetValue(prefixe+'_MONTANTFG')+TOBL.GetValue(prefixe+'_MONTANTFC');
   //
   if TOBL.GetValue('GLC_NONAPPLICFRAIS')<>'X' then TOBL.putValue(prefixe+'_MONTANTFR',arrondi(MontantCHarge*TOBL.GetValue(prefixe+'_COEFFR'),4));
   MontantCharge := TOBL.GetValue(prefixe+'_MONTANTPA')+TOBL.GetValue(prefixe+'_MONTANTFG')+TOBL.GetValue(prefixe+'_MONTANTFC')+TOBL.GetValue(prefixe+'_MONTANTFR');
   //
   TOBL.putValue(prefixe+'_MONTANTPR',MontantCharge);
   if Qte <> 0 then TOBL.putValue(prefixe+'_DPR',arrondi((MontantCharge/Qte)*QteDudetail,V_PGI.okdecP)) else TOBL.PutValue(prefixe+'_DPR',TOBL.GetValue(prefixe+'_MONTANTPR')*QteDudetail);
   if WithPuv then
   begin
     if TOBL.GetValue(prefixe+'_COEFMARG') <> 0 then TOBL.putValue(prefixe+'_PUHT',arrondi(TOBL.GetValue(prefixe+'_DPR')*TOBL.GetValue(prefixe+'_COEFMARG')*QteDudetail,V_PGI.okdecP))
                                                else TOBL.putValue(prefixe+'_PUHT',TOBL.GetValue(prefixe+'_DPR'));
     TOBL.PutValue (prefixe+'_PUHTDEV',pivottodevise(TOBL.GetValue(prefixe+'_PUHT'),DEV.Taux,DEV.quotite,V_PGI.okdecP ));
   end else
   begin
     if TOBL.GetValue(prefixe+'_DPR') <> 0 then TOBL.putValue(prefixe+'_COEFMARG',arrondi(TOBL.GetValue(prefixe+'_PUHT')/TOBL.GetValue(prefixe+'_DPR'),4));
   end;
 end else
 begin
   TOBL.putValue(prefixe+'_MONTANTFG',arrondi(TOBL.GetValue(prefixe+'_MONTANTPA')*TOBL.GetValue(prefixe+'_COEFFG'),4));
   MontantCharge := TOBL.GetValue(prefixe+'_MONTANTPA')+TOBL.GetValue(prefixe+'_MONTANTFG');
   //
   if TOBL.GetValue('GLC_NONAPPLICFC')<>'X' then TOBL.putValue(prefixe+'_MONTANTFC',arrondi(MontantCHarge*TOBL.GetValue(prefixe+'_COEFFC'),4));
   MontantCharge := TOBL.GetValue(prefixe+'_MONTANTPA')+TOBL.GetValue(prefixe+'_MONTANTFG')+TOBL.GetValue(prefixe+'_MONTANTFC');
   //
   if TOBL.GetValue('GLC_NONAPPLICFRAIS')<>'X' then TOBL.putValue(prefixe+'_MONTANTFR',arrondi(MontantCHarge*TOBL.GetValue(prefixe+'_COEFFR'),4));
   MontantCharge := TOBL.GetValue(prefixe+'_MONTANTPA')+TOBL.GetValue(prefixe+'_MONTANTFG')+TOBL.GetValue(prefixe+'_MONTANTFC')+TOBL.GetValue(prefixe+'_MONTANTFR');
   //
   TOBL.putValue(prefixe+'_MONTANTPR',MontantCharge);
   if Qte <> 0 then TOBL.putValue(prefixe+'_DPR',arrondi((MontantCharge/Qte)*QteDudetail,V_PGI.okdecP)) else TOBL.PutValue(prefixe+'_DPR',TOBL.GetValue(prefixe+'_MONTANTPR')*QteDudetail);
   if TOBL.GetValue(prefixe+'_DPR') <> 0 then TOBL.putValue(prefixe+'_COEFMARG',arrondi(TOBL.GetValue(prefixe+'_PUHT')/TOBL.GetValue(prefixe+'_DPR'),4));
 end;
end;

procedure CalculeBaseFraisLigne (TOBL : TOB; CalculViaAchat : boolean=true);
begin
  // ligne standard
  TOBL.PutValue('GL_MONTANTPA',TOBL.GetValue('GL_DPA')*TOBL.GetValue('GL_QTEFACT'));
  if (TOBL.GetValue ('GLC_NONAPPLICFG')<>'X') and (CalculViaAchat) then
  begin
    TOBL.PutValue('GL_MONTANTPAFG',TOBL.GetValue('GL_MONTANTPA'));
  end;

  if (TOBL.GetValue ('GLC_NONAPPLICFRAIS')<>'X') and (CalculViaAchat) then
  begin
    TOBL.PutValue('GL_MONTANTPAFR',TOBL.GEtValue('GL_MONTANTPA'));
  end;

  if (TOBL.GetValue ('GLC_NONAPPLICFC')<>'X') and (CalculViaAchat) then
  begin
    TOBL.PutValue('GL_MONTANTPAFC',TOBL.GEtValue('GL_MONTANTPA'));
  end;
end;

function PieceNonMiseAJourOptimise (TOBpiece,TOBBasesL : TOB) : boolean;
var VenteAchat : string;
begin
	result := false;
	VenteAchat := GetInfoParPiece(TOBPiece.getValue('GP_NATUREPIECEG'), 'GPP_VENTEACHAT');
  if VenteAchat <> 'VEN' then exit;
	result := (TOBPiece.getValue('GP_MONTANTPA')=0) and (TOBPiece.getValue('GP_MONTANTPR')=0) and (TOBPiece.getValue('GP_TOTALHT')<>0) ;
  result := (result) and (TOBBasesL.detail.count =0);
end;

function PieceNonOptimiseCalcul (TOBpiece,TOBOuvrage,TOBBasesL : TOB) : boolean;
begin
  result := false;
  if (TOBPiece.FindFIrst(['GL_TYPELIGNE'],['ART'],true) <> nil) and
     {(TOBOuvrage.detail.count > 0) and}
     (TOBBasesL.detail.count =0) then result := true;
end;

function RecalculePieceFromSaisie (LaTOBPiece,LaTOBOuvrage,LesPorts,LesArticles,LeTiers,LEsPieceRG,LesBasesRG,LesBases: TOB; CalculPV : boolean=false) : TRResult;
var RecalculPiece : TrecalculPiece;
begin
   RecalculPiece := TRecalculPiece.create;
   RecalculPiece.RecalcPv := CalculPv;
   TRY
     RecalculPiece.RecalculPieceFromSaisie(LaTOBPiece,LaTOBOuvrage,LesPorts,LesArticles,LeTiers,LEsPieceRG,LesBasesRG,LesBases);
   FINALLY
     result := RecalculPiece.Result;
   	 REcalculPiece.free;
   END;
end;

function TraitementRecalculPiece (Cledoc : r_cledoc) : TRResult; overload;
var RecalculPiece : TrecalculPiece;
    TOBPiece : TOB;
    StSql : string;
    QQ : TQuery;
    okok : Boolean;
begin
  okok := false;
  StSQl   := 'SELECT * FROM PIECE WHERE '+ WherePiece (Cledoc,ttdpiece,false);
  TOBPiece := TOB.Create('PIECE', nil, -1);
  QQ := OpenSql(StSQL,false);
  if not QQ.eof then
  begin
    TOBPiece.selectDb('', QQ);
    okok := True;
  end;
  Ferme(QQ);
  if okok then
  begin
    RecalculPiece := TRecalculPiece.create;
    RecalculPiece.LaPiece := TOBPiece;
    RecalculPiece.RecalcPv := false;
    RecalculPiece.ReajustePVOuv := false;
    TRY
      RecalculPiece.lanceTraitement;
    FINALLY
      result := RecalculPiece.Result;
      REcalculPiece.free;
    END;
  end;
end;

function TraitementRecalculPiece (TOBPiece : TOB; CalculPv,ReajustePvOuv : boolean) : TRResult;
var RecalculPiece : TrecalculPiece;
begin
   RecalculPiece := TRecalculPiece.create;
   RecalculPiece.LaPiece := TOBPiece;
   RecalculPiece.RecalcPv := CalculPv;
   RecalculPiece.ReajustePVOuv := ReajustePVOuv;
   TRY
     RecalculPiece.lanceTraitement;
   FINALLY
     result := RecalculPiece.Result;
   	 REcalculPiece.free;
   END;
end;

function AffecteTvaPiece (TOBPiece : TOB; NewTva : string) : TRResult;
var RecalculPiece : TrecalculPiece;
begin
   RecalculPiece := TRecalculPiece.create;
   RecalculPiece.LaPiece := TOBPiece;
   RecalculPiece.RecalcPv := false;
   RecalculPiece.ReajustePVOuv := false;
   TRY
     RecalculPiece.AffecteNewTva:= NewTva;
     RecalculPiece.lanceTraitement;
   FINALLY
     result := RecalculPiece.Result;
   	 REcalculPiece.free;
   END;
end;

procedure SupprimeFraisRepartisPiece (TOBPorcs : TOB);
var Indice : integer;
begin
  indice := 0;
  if TOBPOrcs.detail.count = 0 then exit;
  repeat
    if TOBPorcs.detail[Indice].GetValue('GPT_FRAISREPARTIS')='X' then TOBPorcs.detail[Indice].free else Inc(Indice);
  until Indice >= TOBPorcs.detail.count -1;
end;

procedure  SupprimeFraisDetaille (TOBpiece : TOB);
begin

end;


{ TRecalculPiece }

constructor TRecalculPiece.create;
begin
  ReajustePVOuv := false;
  TheRepartTva := TREPARTTVAMILL.create (nil) ;
  TOBPorcs := TOB.Create('LES PORCS',nil,-1);
  TOBOuvrages := TOB.Create('LES OUVRAGES',nil,-1);
  TOBOuvragesP := TOB.Create('LES OUVRAGES PLAT',nil,-1);
  TOBArticles := TOB.create ('LES ARTICLES',nil,-1);
  TOBTiers := TOB.Create('TIERS', nil, -1);
  TOBPieceRG :=  TOB.create('PIECERRET', nil, -1);
  TOBPieceRG_O := TOB.create('PIECERRET', nil, -1);
  TOBBasesRG := TOB.create('BASESRG', nil, -1);
  TOBVTECOLL := TOB.Create ('LES COLLE', nil,-1);
  TOBBasesL := TOB.create('LES BASES LIGNE', nil, -1);
  TOBBases := TOB.create('LES BASES', nil, -1);
  TOBPieceTrait := TOB.create('LES REPART COTRAIT', nil, -1);
  TOBAFFAIRE := TOB.create ('AFFAIRE',nil,-1);
  TOBSSTRAIT := TOB.Create ('LES SOUS TRAIT',nil,-1);
  TOBAcomptes := TOB.Create ('LES ACOMPTES',nil,-1);
  TOBEChes := TOB.Create ('LES ECHEANCES',nil,-1);
  TOBAffaireInterv := TOB.Create ('LES INTERV DE lAFF',nil,-1);
  ArtEcart := trim(GetParamsoc('SO_BTECARTPMA'));
  TOBCPTA := CreerTOBCpta;
  TOBanaP:= TOB.Create ('LES ANAP',nil,-1);
  TOBanaS := TOB.Create ('LES ANAS',nil,-1);
  fNewCodeTva := '';
end;

procedure TRecalculPiece.InitPiece;
begin
  TheRepartTva.Initialise;
  TOBPorcs.clearDetail;
  TOBOuvrages.clearDetail;
  TOBOuvragesP.ClearDetail;
  TOBArticles.clearDetail;
  TOBTiers.InitValeurs ;
  TOBPieceRG.clearDetail;
  TOBPieceRG_O.clearDetail;
  TOBPieceRG.InitValeurs;
  TOBPieceRG_O.InitValeurs;
  TOBBasesRG.ClearDetail ;
  TOBVTECOLL.clearDetail;
  TOBBasesL.clearDetail;
  TOBBases.clearDetail;
  TOBPieceTrait.clearDetail;
  TOBAFFAIRE.InitValeurs;
  TOBSSTRAIT.clearDetail;
  TOBAcomptes.clearDetail;
  TOBEChes.ClearDetail;
  TOBAffaireInterv.ClearDetail;
  TOBCPTA.ClearDetail; TOBCPTA.SetString('LASTSQL','');
  TOBanaP.ClearDetail;
  TOBanaS.ClearDetail;
end;

destructor TRecalculPiece.destroy;
begin
  TOBAFFAIRE.free;
  TOBPieceTrait.free;
  TheRepartTva.free;
	TOBpiece.clearDetail;
  TOBPorcs.free;
  TOBOuvrages.free;
  TOBOuvragesP.free;
  TOBArticles.free;
  TOBTiers.free;
  TOBPieceRG.free;
  TOBPieceRG_O.free;
  TOBBasesRG.free;
  TOBVTECOLL.free;
  TOBBases.free;
  TOBBasesL.free;
  TOBSSTRAIT.free;
  TOBAcomptes.free;
  TOBEChes.Free;
  TOBAffaireInterv.Free;
  TOBCPTA.Free;
  TOBanaP.free;
  TOBanaS.free;
  inherited;
end;


procedure TRecalculPiece.ReInitMontantPA (TOBL : TOB);
var TypeLigne : string;
begin
  TypeLigne := TOBL.GEtVAlue('GP_TYPELIGNE');
  if (TypeLigne='TOT') or (Copy(TypeLigne,1,2)='TP') then
  begin
	  ZeroLigne (TOBL);
  end;
  TOBL.putValue('GL_MONTANTPA',0);
  TOBL.putValue('GL_MONTANTPAFG',0);
  TOBL.putValue('GL_MONTANTPAFC',0);
  TOBL.putValue('GL_MONTANTPAFR',0);
  TOBL.putValue('GL_MONTANTFG',0);
  TOBL.putValue('GL_MONTANTFR',0);
  TOBL.putValue('GL_MONTANTFC',0);
  TOBL.putValue('GL_COEFFR',0);
  TOBL.putValue('GL_COEFFC',0);
  TOBL.putValue('GL_MONTANTPR',0);
end;

procedure TRecalculPiece.ReInitMontantPiece;
begin
	TOBPIece.putValue('GP_MONTANTPAFG',0);
	TOBPIece.putValue('GP_MONTANTPAFC',0);
	TOBPIece.putValue('GP_MONTANTPAFR',0);
	TOBPIece.putValue('GP_MONTANTFG',0);
	TOBPIece.putValue('GP_MONTANTFC',0);
	TOBPIece.putValue('GP_MONTANTFR',0);
	TOBPIece.putValue('GP_MONTANTPA',0);
	TOBPIece.putValue('GP_MONTANTPR',0);
	TOBPIece.putValue('GP_COEFFR',0);
	TOBPIece.putValue('GP_COEFFC',0);
  if TOBPiece.getValue('GP_APPLICFGST') <> 'X' then TOBPiece.PutValue('GP_APPLICFGST','-') ;
  if TOBPiece.getValue('GP_APPLICFCST') <> 'X' then TOBPiece.PutValue('GP_APPLICFCST','-') ;
  ReinitMontantPieceTrait(TOBPiece,TOBaffaire,TOBpieceTrait);
end;

procedure TRecalculPiece.RecalculePAPRPV;
var Indice : integer;
		TOBL : TOB;
    RecalcCoefFG : boolean;
    ApplicFRLigne,ApplicFCLigne,ApplicFGLigne,NumOrdre : integer;
    MontantPAFC,MontantFG : double;
begin
  Nature := TOBPiece.getValue('GP_NATUREPIECEG');
  VenteAchat := GetInfoParPiece(Nature, 'GPP_VENTEACHAT');
  CalculViaAchat :=  isPieceGerableFraisDetail(Nature);
  ApplicFRPiece := TOBPiece.GetValue ('GP_APPLICFGST');
  ApplicFCPiece := TOBPiece.GetValue ('GP_APPLICFCST');
  TRY
    // Etape 1 recalcul de Montants PA
    for Indice := 0 to TOBPiece.detail.count -1 do
    begin

      TOBL := TOBPiece.detail[Indice];
      NumOrdre := TOBL.GetInteger('GL_NUMORDRE');

      if indice = 0 then
      begin
        ApplicFGLigne := TOBL.GetNumChamp ('GLC_NONAPPLICFG');
        ApplicFCLigne := TOBL.GetNumChamp ('GLC_NONAPPLICFC');
        ApplicFRLigne := TOBL.GetNumChamp ('GLC_NONAPPLICFRAIS');
      end;

      if not IsArticle (TOBL) then continue;

      ReInitMontantPA (TOBL);
      if ((TOBL.GetValue('GL_TYPEARTICLE') = 'OUV') or (TOBL.GetValue('GL_TYPEARTICLE') = 'ARP')) and
         (TOBL.GetValue('GL_INDICENOMEN')> 0) then
      begin
        // gestion d'un ouvrage
        TraitementOuvrage (TOBL);
      end else
      begin
        CalculeBaseFraisLigne (TOBL,CalculViaAchat);
      end;
      SommationLignePiece (TOBL,TOBPiece);
    end;
  EXCEPT
  	CodeRetour := TrrErrBeforeCalc;
  	Exit;
  END;

  TRY
    // Etape 2 calcul des Coef FG et Coef FC
    MontantPAFC := TOBPIece.getValue('GP_MONTANTPAFC');
    MontantFG := TOBPIece.getValue('GP_MONTANTFG');
    TOBPIece.putValue('GP_COEFFR',arrondi(CoeffR/100,6));
    if MontantFraisChantier <> 0 then TOBPIece.putValue('GP_COEFFC', arrondi(MontantFraisChantier/(MontantPAFC+MontantFg),6))
    														 else TOBPIece.putValue('GP_COEFFC', 0);
		RecalcCoefFG := false;
    // Applique et calcul PR
    AppliqueFraisPiece (TOBPiece,TOBOuvrages,CalculViaAchat,CalculPv,EnHt,DEV,RecalcCoefFG) ;
  EXCEPT
  	CodeRetour := TrrErrCalcPR;
  	exit;
  END;

  TRY
//    	AjouteLignesVirtuellesOuv ( TOBPiece,TOBOuvrages ,TOBArticles,nil,TOBTiers,nil,nil,nil,DEV) ;
    ZeroFacture (TOBpiece); ZeroMontantPorts(TOBporcs); ReinitMontantPieceTrait (TOBPiece,TOBaffaire,TOBpieceTrait);
    for Indice := 0 to TOBPiece.detail.count -1 do ZeroLigneMontant (TOBPiece.detail[Indice]);
    PutValueDetail (TOBpiece,'GP_RECALCULER','X');
    TOBBases.ClearDetail;
    TOBBasesL.ClearDetail;
    TOBVTECOLL.ClearDetail;
    CalculFacture( TOBaffaire,TOBPiece, TOBPieceTrait,TOBSSTRAIT,TOBouvrages, TOBOuvragesP, TOBBases, TOBBasesL,TOBTiers, TOBArticles, TOBPorcs, TOBPieceRG, TOBBasesRG,TOBVTECOLL, DEV, false, taModif,false,-1,false,true);
//			DetruitLignesVirtuellesOuv (TOBPIece,DEV);
      //
	EXCEPT
    CodeRetour := TrrErrcalcPV;
    exit;
  END;

  CalculeSousTotauxPiece(TOBPiece);
end;

procedure TRecalculPiece.CalculeLaPiece;
begin
  ReInitMontantPiece;
  BeforeTraitementCalculCoef (TOBArticles,TOBPiece,TOBPorcs,TOBOuvrages,DEV,False,fNewCodeTva); // histoire de recup les coef de frais generaux (COEFFG)
  TRY
    if ReajustePVOuv then
    begin
      ControleEtReajusteOuvrages;
    end;
  	RecalculePAPRPV;
  EXCEPT
  	Exit;
  END;
end;

procedure TRecalculPiece.ReCalculeLaPiece;
begin
  ReInitMontantPiece;
  TRY
  	RecalculePAPRPV;
  EXCEPT
  	Exit;
  END;
end;

procedure TRecalculPiece.EnregistreLaPiece;
begin
  TRY
    TRANSACTIONS (validelaPiece,0);
  EXCEPT
    CodeRetour := TrrErrEcriture;
    exit;
  END;
end;

procedure TRecalculPiece.lanceTraitement;
begin
	CodeRetour := TrrOk;
	LoadLesTObs ;

  TRY
  	CalculeLaPiece;
  EXCEPT
  	Exit;
  END;

  EnregistreLaPiece;
end;

procedure TRecalculPiece.SommationLignePiece (TOBL,TOBPiece : TOB);
begin
	if (TOBL.GetValue('GL_TYPELIGNE')<>'ART') and (TOBL.GetValue('GL_TYPELIGNE')<>'ARV') then exit;
	if (TOBL.GetValue('GL_TYPELIGNE')='ARV') and (TOBL.GetValue('GL_NATUREPIECEG')<>'PBT') then exit;
	TOBPiece.PutValue('GP_MONTANTPA',TOBPiece.GetValue('GP_MONTANTPA')+TOBL.GetValue ('GL_MONTANTPA'));
	TOBPiece.PutValue('GP_MONTANTPAFG',TOBPiece.GetValue('GP_MONTANTPAFG')+TOBL.GetValue ('GL_MONTANTPAFG'));
	TOBPiece.PutValue('GP_MONTANTPAFC',TOBPiece.GetValue('GP_MONTANTPAFC')+TOBL.GetValue ('GL_MONTANTPAFC'));
	TOBPiece.PutValue('GP_MONTANTPAFR',TOBPiece.GetValue('GP_MONTANTPAFR')+TOBL.GetValue ('GL_MONTANTPAFR'));
	TOBPiece.PutValue('GP_MONTANTFG',TOBPiece.GetValue('GP_MONTANTFG')+TOBL.GetValue ('GL_MONTANTFG'));
end;

procedure TRecalculPiece.CreerLesTOBOuv (TOBGroupeNomen,TOBNomen: TOB ; LaLig,idep : integer) ;
Var i,Lig : integer ;
    TOBLN,TOBPere,TOBLoc : TOB ;
    LigneN1,LigneN2,LigneN3,LigneN4,LigneN5: integer;
BEGIN
  for i:=idep to TOBNomen.Detail.Count-1 do
  BEGIN
    TOBLN:=TOBNomen.Detail[i] ;
    if I =0 then EnHt := (TOBLN.GetString('BLO_FACTUREHT')='X');
    Lig:=TOBLN.GetValue('BLO_NUMLIGNE') ;
    if lig <> LaLig then break;
    LigneN1 := TOBLN.GetValue('BLO_N1');
    LigneN2 := TOBLN.GetValue('BLO_N2');
    LigneN3 := TOBLN.GetValue('BLO_N3');
    LigneN4 := TOBLN.GetValue('BLO_N4');
    LigneN5 := TOBLN.GetValue('BLO_N5');
    if LigneN5 > 0 then
    begin
      // recherche du pere au niveau 4
      TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,LigneN2,LigneN3,LigneN4,0],True) ;
    end else if TOBLN.GetValue('BLO_N4') > 0 then
    begin
      // recherche du pere au niveau 3
      TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,LigneN2,LigneN3,0,0],True) ;
    end else if TOBLN.GetValue('BLO_N3') > 0 then
    begin
      // recherche du pere au niveau 2
      TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,LigneN2,0,0,0],True) ;
    end else if TOBLN.GetValue('BLO_N2') > 0 then
    begin
      // recherche du pere au niveau 1
      TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,0,0,0,0],True) ;
    end else TOBPere:=TOBGroupeNomen;

    if TOBPere<>Nil then
    BEGIN
      TOBLoc:=TOB.Create('LIGNEOUV',TOBPere,-1) ;
      TOBLoc.Dupliquer(TOBLN,False,True) ;
  		InsertionChampSupOuv (TOBLoc,false);
    	if EnHt then TOBLoc.putvalue ('ANCPV',TOBLoc.Getvalue ('BLO_PUHTDEV'))
    					else TOBLoc.putvalue ('ANCPV',TOBLoc.Getvalue ('BLO_PUTTCDEV'));
    END ;
  END ;
END ;

procedure TRecalculPiece.ChargeLesOuvrages ;
var TOBNomen,TOBLN,TOBL,TOBGroupeNomen : TOB;
		Q : Tquery;
    i,Lig,Oldl,IndiceNomen : integer;
begin
	Oldl := -1;
  TOBNomen:=TOB.Create('',Nil,-1) ;
  Q:=OpenSQL('SELECT O.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM LIGNEOUV O '+
             'LEFT JOIN NATUREPREST N ON BNP_NATUREPRES=(SELECT GA_NATUREPRES FROM ARTICLE WHERE GA_ARTICLE=O.BLO_ARTICLE) '+
             'WHERE '+WherePiece(CleDoc,ttdOuvrage,False)+
             ' ORDER BY BLO_NUMLIGNE,BLO_N1, BLO_N2, BLO_N3, BLO_N4,BLO_N5',True) ;
  TOBNomen.LoadDetailDB ('LIGNEOUV','','',Q,True,true);
  Ferme(Q) ;
  for i:=0 to TOBNomen.Detail.Count-1 do
  BEGIN
    TOBLN:=TOBNomen.Detail[i] ;
    Lig:=TOBLN.GetValue('BLO_NUMLIGNE') ;
    if OldL<>Lig then
    BEGIN
      TOBGroupeNomen:=TOB.Create('',TOBOuvrages,-1) ;
      IndiceNomen := TobOuvrages.detail.count;
      TobL := TobPiece.FindFirst(['GL_NUMLIGNE'], [Lig], False);
      if TOBL <> nil then
      begin
        TOBL.PutValue('GL_INDICENOMEN',IndiceNomen);
      end;
      CreerLesTOBOuv(TOBGroupeNomen,TOBNomen,Lig,i);
    END ;
    OldL:=Lig ;
  END ;
  TOBNomen.Free ;
end;

procedure TRecalculPiece.LoadLesTObs;
var	QQ : TQuery;
    indice,Ind : integer;
begin
	//
	if TOBPiece.getValue('GP_AFFAIRE')<> '' then
  begin
  	QQ := OpenSql ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE ="'+TOBPiece.getvalue('GP_AFFAIRE')+'"',TRue, 1,'',true);
    if not QQ.eof then TOBaffaire.SelectDB('',QQ);
    ferme (QQ);
  end;
  //
  cledoc := TOB2CleDoc (TOBPiece);
	LoadLignes (cledoc,tobpiece);
  PieceAjouteSousDetail(TOBPiece,true,false,true);

  // Lecture des affectations document
  LoadLaTOBPieceTrait (TOBPieceTrait,Cledoc,'');
  LoadLesSousTraitants(Cledoc,TOBSSTRAIT);
  //
  QQ := OpenSQL('SELECT * FROM PIEDPORT WHERE ' + WherePiece(CleDoc, ttdPorc, False), True);
  TOBPorcs.LoadDetailDB('PIEDPORT', '', '', QQ, False);
  Ferme(QQ);
  // protection sur les pièces comptabilisés
  if (CalculPv) and (GetInfoParPiece(cledoc.NaturePiece, 'GPP_TYPEECRCPTA') <> '') then CalculPv := false;
  for Indice := 0 to TOBPiece.detail.count -1 do
  BEGIN
    TOBPiece.detail[Indice].AddChampSupValeur ('GL_RECALCULER','-');
    TOBPiece.detail[Indice].AddChampSupValeur ('MONTANTSIT',0);
//
    TOBPiece.detail[Indice].AddChampSupValeur('TOTREMLIGNETTCDEV', 0);
    TOBPiece.detail[Indice].AddChampSupValeur('TOTREMLIGNETTC', 0);
    TOBPiece.detail[Indice].AddChampSupValeur('TOTREMPIEDTTCDEV', 0);
    TOBPiece.detail[Indice].AddChampSupValeur('TOTREMPIEDTTC', 0);
    TOBPiece.detail[Indice].AddChampSupValeur('TOTESCLIGNETTCDEV', 0);
    TOBPiece.detail[Indice].AddChampSupValeur('TOTESCLIGNETTC', 0);
//
    for Ind:= 0 to VH_GC.TOBParamTaxe.detail.count -1 do
    begin
      TOBPiece.detail[Indice].AddChampSupValeur ('MILLIEME'+IntToStr(Ind+1),0);
    end;
  END;
  TOBPiece.AddChampSupValeur ('GP_RECALCULER','-');
  TOBPiece.AddChampSupValeur ('MONTANTSIT',0);
  TOBPiece.AddChampsupValeur('RUPTMILLIEME', '');
  // --
  RecupCoefFG;
  ChargeLesOuvrages ;
  //
  ChargeLesArticles;
  ChargeTiers;
  LoadLesRetenues(TOBPiece, TOBPieceRG, TOBBasesRG,taModif);
  TOBPieceRG_O.Dupliquer(TOBPieceRG,True,true); 
  //
  TheRepartTva.Initialise;
  TheRepartTva.TOBBases := TOBBases;
  TheRepartTva.TOBpiece := TOBPiece;
  TheRepartTva.TOBOuvrages := TOBOuvrages;
  TheRepartTva.Charge;
  TheRepartTva.Applique ;
  //
  // Lecture ACompte
  LoadLesAcomptes(TOBPiece, TOBAcomptes, CleDoc, nil);
  // Lecture Analytiques
  LoadLesAna(TOBPiece, TOBAnaP, TOBAnaS);
end;

procedure TRecalculPiece.SetPiece(const Value: TOB);
var ExistFg : boolean;
begin
  TOBPiece := value;
  if Pos(TOBPiece.getValue('GP_NATUREPIECEG'),'ETU;DBT;BCE') > 0 then
  begin
  	MontantFraisChantier := GetMontantFraisDetail (TOBPiece,ExistFg);
  end;
  EnHt := (TOBPiece.getValue('GP_FACTUREHT')='X');
  DEV.Code := TOBPIECE.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := TOBPiece.getValue('GP_TAUXDEV');
end;

procedure TRecalculPiece.RecupCoefFG;
var i : integer;
begin
  CoeffR:=0;
  for i:=0 to TOBPorcs.Detail.Count-1 do
  BEGIN
    if TOBPorcs.Detail[i].GetValue('GPT_FRAISREPARTIS') = 'X' then
    Begin
      CoeffR := CoeffR + TOBPorcs.Detail[i].GetValue('GPT_POURCENT');
    End;
  END;
end;


procedure CalculeMontantLigneOuv(TOBOuvrage,TOBO: TOB);
begin
  if TOBO.GetVAlue('BLO_NONAPPLICFRAIS')<>'X' then TOBO.PutValue('BLO_NONAPPLICFRAIS','-');
  if TOBO.GetVAlue('BLO_NONAPPLICFC')<>'X' then TOBO.PutValue('BLO_NONAPPLICFC','-');
  if TOBO.GetVAlue('BLO_NONAPPLICFG')<>'X' then TOBO.PutValue('BLO_NONAPPLICFG','-');

  TOBO.putValue('BLO_MONTANTPA',Arrondi(TOBO.GetValue('BLO_DPA')*TOBO.GetValue('BLO_QTEFACT'),V_PGI.OkDecP ));

  if TOBO.GetVAlue('BLO_NONAPPLICFG')<>'X' then
  begin
  	TOBO.putValue('BLO_MONTANTPAFG',TOBO.GetValue('BLO_MONTANTPA'));
  end;
  //
  TOBO.putValue('BLO_MONTANTFG',TOBO.GetValue('BLO_MONTANTPA')*TOBO.GetValue('BLO_COEFFG'));

  if TOBO.GetVAlue('BLO_NONAPPLICFRAIS')<>'X' then
  begin
  	TOBO.putValue('BLO_MONTANTPAFR',TOBO.GetValue('BLO_MONTANTPA'));
  end;

  if TOBO.GetVAlue('BLO_NONAPPLICFC')<>'X' then
  begin
	  TOBO.putValue('BLO_MONTANTPAFC',TOBO.GetValue('BLO_MONTANTPA'));
  end;

  if (Pos (TOBO.getValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne) > 0) and (not IsLigneExternalise(TOBO)) then
  begin
  	TOBO.putValue('BLO_TPSUNITAIRE',1);
  end else
  begin
  	TOBO.putValue('BLO_TPSUNITAIRE',0);
  end;

end;

procedure ReinitMontantPALigneOuv(TOBO: TOB; CalculPv : boolean=false;RazFr : boolean = false);
begin
  TOBO.PutValue('BLO_MONTANTPA',0);
  TOBO.PutValue('BLO_MONTANTFG',0);
  if RazFR then
  begin
    TOBO.PutValue('BLO_MONTANTFR',0);
    TOBO.PutValue('BLO_MONTANTFC',0);
  end;
  TOBO.PutValue('BLO_MONTANTPR',0);
  TOBO.PutValue('BLO_MONTANTPAFG',0);
  TOBO.PutValue('BLO_MONTANTPAFR',0);
  TOBO.PutValue('BLO_MONTANTPAFC',0);
  if RazFR then
  begin
  	TOBO.putValue('BLO_COEFFR',0);
		TOBO.putValue('BLO_COEFFC',0);
  end;
  TOBO.putValue('BLO_TPSUNITAIRE',0);
(*
  if not CalculPv then
  begin
    TOBO.putValue('BLO_COEFMARG',0);
  end;
*)
end;

procedure GereTraitementOuvrage (TOBOuvrage : TOB; var Valeur : T_Valeurs; DEV : Rdevise;  CalculPv : boolean=false; RazFg : boolean=false);
var TOBO : TOB;
		Indice : integer;
    Valloc,ValPou : T_Valeurs;
    QteDuDetail,Qte : double;
    IndPou : Integer;
    ArticleOk : string;
begin
  QteDudetail := 1;

  ArticleOk := '';
  IndPou := -1 ;

  TOBO := TOBOuvrage.FindFirst(['BLO_TYPEARTICLE'],['POU'],false);
  if TOBO <> nil then
  begin
  	IndPou := TOBO.GetIndex;
		ArticleOk := TOBO.GetValue('BLO_LIBCOMPL');
  end else IndPou := -1;

	for Indice := 0 to TOBOuvrage.detail.count -1 do
  begin
  	TOBO := TOBOuvrage.detail[Indice];
    if TOBO.GetString ('BLO_TYPEARTICLE')='POU' Then continue;
  	ReinitMontantPALigneOuv (TOBO,CalculPV,razFg);
    Qte := TOBO.GetValue('BLO_QTEFACT');
    if Indice = 0 then
    begin
    	QteDuDetail := TOBO.GetValue('BLO_QTEDUDETAIL');
    	if QteDuDetail = 0  then QteDuDetail := 1;
    end;
    if TOBO.Detail.count > 0 then
    begin
		  InitTableau (Valloc);
    	GereTraitementOuvrage (TOBO,Valloc,DEV,CalculPv,RazFG);
      TOBO.PutValue('BLO_TPSUNITAIRE',arrondi(Valloc[9],4));
      TOBO.PutValue('BLO_MONTANTPAFG',arrondi(Qte*Valloc[10],4));
      TOBO.PutValue('BLO_MONTANTPAFR',arrondi(Qte*Valloc[11],4));
      TOBO.PutValue('BLO_MONTANTPAFC',arrondi(Qte*Valloc[12],4));
      TOBO.PutValue('BLO_MONTANTPA',arrondi(Qte*TOBO.GetValue('BLO_DPA'),V_PGI.okdecP));
      TOBO.PutValue('BLO_MONTANTFG',arrondi(Qte*Valloc[13],4));
      if CalculPv then
      begin
        TOBO.Putvalue('BLO_PUHTDEV',Valloc[2]);
        TOBO.Putvalue('BLO_PUTTCDEV',Valloc[3]);
        TOBO.Putvalue('BLO_PUHT',DeviseToPivotEx(TOBO.GetValue('BLO_PUHTDEV'),DEV.Taux,DEV.quotite,V_PGI.OkdecP));
        TOBO.Putvalue('BLO_PUTTC',DevisetoPivotEx(TOBO.GetValue('BLO_PUTTCDEV'),DEV.taux,DEV.quotite,V_PGI.okdecP));
				CalculMontantHtDevLigOuv (TOBO,DEV);
      end;
    end else
    begin
    	CalculeMontantLigneOuv (TOBOuvrage,TOBO);
      if CalculPv then
      begin
      	CalculMontantHtDevLigOuv (TOBO,DEV);
      end;
    end;
    Valeur[9] := Valeur[9]+TOBO.getValue('BLO_TPSUNITAIRE')*Qte;
    Valeur[10] := Valeur[10]+TOBO.getValue('BLO_MONTANTPAFG');
    Valeur[11] := Valeur[11]+TOBO.getValue('BLO_MONTANTPAFR');
    Valeur[12] := Valeur[12]+TOBO.getValue('BLO_MONTANTPAFC');
    Valeur[13] := Valeur[13]+TOBO.getValue('BLO_MONTANTFG');
    Valeur[16] := Valeur[16]+TOBO.getValue('BLO_MONTANTPA');
    if CalculPv then
    begin
      Valeur[2] := Valeur[2] + TOBO.GetValue ('BLO_MONTANTHTDEV');
      Valeur[3] := Valeur[3] + TOBO.GetValue ('BLO_MONTANTTTCDEV');
    end;

		// Calcul pour article pourcentage
    if (IndPou >=0) and (ArticleOKInPOUR (TOBO.GetValue('BLO_TYPEARTICLE'),ArticleOK)) then
    begin
      ValPou[0] := ValPou[0] + ((Qte/QteDuDetail) * TOBO.GetValue ('BLO_DPA'));
      ValPou[1] := ValPou[1] + ((Qte/QteDuDetail) * TOBO.GetValue ('BLO_DPR'));
      ValPou[6] := ValPou[6] + ((Qte/QteDuDetail) * TOBO.GetValue ('BLO_PMAP'));
      ValPou[7] := ValPou[7] + ((Qte/QteDuDetail) * TOBO.GetValue ('BLO_PMRP'));
      if CalculPv then
      begin
        ValPou[2] := ValPou[2] + ((Qte/QteDuDetail) * TOBO.GetValue ('BLO_PUHTDEV'));
        ValPou[3] := ValPou[3] + ((Qte/QteDuDetail) * TOBO.GetValue ('BLO_PUTTCDEV'));
      end;
    end;
  end;

  if IndPou >= 0 then
  begin
  	TOBO := TOBOuvrage.detail[IndPou];
    Qte := 1;
    TOBO.putvalue ('BLO_DPA',ValPou[0]);
    TOBO.putvalue ('BLO_DPR',ValPou[1]);
    if calculPv then
    begin
      TOBO.putvalue ('BLO_PUHTDEV',ValPou[2]);
      TOBO.putvalue ('BLO_PUTTCDEV',ValPou[3]);
      TOBO.putvalue ('BLO_PUHT',DevisetopivotEx(ValPou[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
      TOBO.putvalue ('BLO_PUHTBASE',TOBOuvrage.detail[IndPou].Getvalue ('BLO_PUHT'));
      TOBO.putvalue ('BLO_PUTTC',DevisetopivotEx(ValPou[3],DEV.Taux,DEV.quotite,V_PGI.okdecP));
      TOBO.putvalue ('BLO_PUTTCBASE',TOBOuvrage.detail[IndPou].Getvalue ('BLO_PUTTC'));
    end;
    TOBO.PutValue('BLO_PMAP',ValPou[6]);
    TOBO.PutValue('BLO_PMRP',ValPou[7]);
    CalculMontantHtDevLigOuv (TOBO,DEV);

    Valeur[9] := Valeur[9]+TOBO.getValue('BLO_TPSUNITAIRE')*Qte;
    Valeur[10] := Valeur[10]+TOBO.getValue('BLO_MONTANTPAFG');
    Valeur[11] := Valeur[11]+TOBO.getValue('BLO_MONTANTPAFR');
    Valeur[12] := Valeur[12]+TOBO.getValue('BLO_MONTANTPAFC');
    Valeur[13] := Valeur[13]+TOBO.getValue('BLO_MONTANTFG');
    Valeur[16] := Valeur[16]+TOBO.getValue('BLO_MONTANTPA');
    if CalculPv then
    begin
      Valeur[2] := Valeur[2] + TOBO.GetValue ('BLO_MONTANTHTDEV');
      Valeur[3] := Valeur[3] + TOBO.GetValue ('BLO_MONTANTTTCDEV');
    end;

  end;

  Valeur[10] := Arrondi(Valeur[10]/QteDudetail,V_PGI.Okdecp);
  Valeur[11] := ARRONDI(Valeur[11]/QteDudetail,V_PGI.OkDecP);
  Valeur[12] := ARRONDI(Valeur[12]/QteDudetail,V_PGI.OkDecP);
  Valeur[13] := Arrondi(Valeur[13]/QteDudetail,V_PGI.OkDecP);
  Valeur[16] := ARRONDI(Valeur[16]/QteDudetail,V_PGI.OkDecP);
  if CalculPv then
  begin
    Valeur[2] := ARRONDI(Valeur[2]/QteDudetail,V_PGI.OkDecP);
    Valeur[3] := ARRONDI(Valeur[3]/QteDudetail,V_PGI.OkDecP);
  end;


end;

procedure TRecalculPiece.TraitementOuvrage(TOBL: TOB;WithPuv : Boolean=false);
var IndiceNomen : integer;
		TOBOuvrage : TOB;
    Valeur : T_Valeurs;
    Qte : double;
begin
  Qte := TOBL.GetValue('GL_QTEFACT');
	CodeRetour := TrrOk;
  IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
  InitTableau (Valeur);
  if IndiceNomen = 0 then exit;
  TRY
  	TOBOuvrage := TOBOuvrages.detail[IndiceNomen-1];
    GereTraitementOuvrage (TOBOuvrage,valeur,DEV,(CalculPV or WithPuv));

    TOBL.PutValue('GL_TPSUNITAIRE', Valeur[9]);
    TOBL.PutValue('GL_MONTANTPAFG', Arrondi(Valeur[10] * Qte,4));
    TOBL.PutValue('GL_MONTANTPAFR', arrondi(Valeur[11] * Qte,4));
    TOBL.PutValue('GL_MONTANTPAFC', arrondi(Valeur[12] * Qte,4));
    TOBL.PutValue('GL_MONTANTFG',   Arrondi(Valeur[13] * Qte,4));
    TOBL.PutValue('GL_MONTANTPA',   Arrondi(Valeur[16] * Qte,DEV.DECIMALE));
    if WithPuv then
    begin
      TOBL.Putvalue('GL_PUHTDEV',Valeur[2]);
      TOBL.Putvalue('GL_PUTTCDEV',Valeur[3]);
      TOBL.Putvalue('GL_PUHT',DeviseToPivotEx(TOBL.GetValue('GL_PUHTDEV'),DEV.Taux,DEV.quotite,V_PGI.OkdecP));
      TOBL.Putvalue('GL_PUTTC',DevisetoPivotEx(TOBL.GetValue('GL_PUTTCDEV'),DEV.taux,DEV.quotite,V_PGI.okdecP));
      TOBL.Putvalue('GL_PUHTBASE',TOBL.GetValue('GL_PUHT'));
      TOBL.Putvalue('GL_PUTTCBASE',TOBL.GetValue('GL_PUTTC'));
    end;
  EXCEPT
  	raise Exception.create ('');
  END;
end;


procedure AppliqueFraisDetailOuv (TOBPiece,TOBOUV : TOB ; var Valeurs : T_Valeurs; PVBloque,CalculViaAchat,CalculPv,EnHt : boolean; DEV : RDevise; RecalcCoefFg : boolean);
var indice : integer;
		valloc : T_Valeurs;
    Qte,CoefMarg,CoefDEv,MontantPACharge : double;
    PrestationST,ApplicFRST,ApplicFCST : boolean;
    QteDuDetail : double;
 begin
  ApplicFRSt := (TOBPIece.getValue('GP_APPLICFGST')='X');
  ApplicFCSt := (TOBPIece.getValue('GP_APPLICFCST')='X');
  QteDuDetail := TOBOUV.GetValue('BLO_QTEDUDETAIL');
  if QteDuDetail = 0  then QteDuDetail := 1;

	if TOBOUV.detail.count > 0 then
  begin
  	coefdev := TOBOUV.GetValue('BLO_TAUXDEV');
    if coefdev=0 then coefdev:=1;
  	InitTableau (Valloc);
    Qte := TOBOUV.GetValue('BLO_QTEFACT');
    for Indice := 0 To TOBOUV.detail.count -1 do
    begin
    	AppliqueFraisDetailOuv (TOBPiece,TOBOUV.detail[Indice],valloc,PVBloque,CalculViaAchat,CalculPV,EnHt,DEV,RecalcCoefFg);
    end;

    if CalculViaAchat then
    begin
      TOBOUV.PutValue('BLO_MONTANTPAFG',Arrondi(Valloc[10]*Qte,4));
      TOBOUV.PutValue('BLO_MONTANTPAFR',Arrondi(Valloc[11]*Qte,4));
      TOBOUV.PutValue('BLO_MONTANTPAFC',Arrondi(Valloc[12]*Qte,4));
      TOBOUV.PutValue('BLO_MONTANTFG',Arrondi(Valloc[13]*Qte,4));
      TOBOUV.PutValue('BLO_MONTANTFR',Arrondi(Valloc[14]*Qte,4));
      TOBOUV.PutValue('BLO_MONTANTFC',Arrondi(Valloc[15]*Qte,4));
       //
      if TOBOUV.GetValue('BLO_MONTANTPAFG') <> 0 then
      begin
        TOBOUV.PutValue('BLO_COEFFG',arrondi(TOBOUV.GetValue('BLO_MONTANTFG')/TOBOUV.GetValue('BLO_MONTANTPAFG'),4));
      end;
      //
      if TOBOUV.GetValue('BLO_MONTANTPAFC') <> 0 then
      begin
        TOBOUV.PutValue('BLO_COEFFC',arrondi(TOBOUV.GetValue('BLO_MONTANTFC')/TOBOUV.GetValue('BLO_MONTANTPAFC'),4));
      end;

      if TOBOUV.GetValue('BLO_MONTANTPAFR') <> 0 then
      begin
        TOBOUV.PutValue('BLO_COEFFR',arrondi(TOBOUV.GetValue('BLO_MONTANTFR')/TOBOUV.GetValue('BLO_MONTANTPAFR'),4));
      end;
       //
      if TOBOUV.GetValue('BLO_MONTANTFR')+TOBOUV.GetValue('BLO_MONTANTFC')+TOBOUV.GetValue('BLO_MONTANTFG') <> 0 then
      begin
        TOBOUV.PutValue('BLO_MONTANTPR',TOBOUV.GetValue('BLO_MONTANTPA')+
        																TOBOUV.GetValue('BLO_MONTANTFR')+
                                        TOBOUV.GetValue('BLO_MONTANTFG')+
                                        TOBOUV.GetValue('BLO_MONTANTFC'));
        if TOBOUV.GetValue('BLO_QTEFACT') <> 0 then
        begin
          TOBOUV.PutValue('BLO_DPR',arrondi(TOBOUV.GetValue('BLO_MONTANTPR')/TOBOUV.GetValue('BLO_QTEFACT'),V_PGI.okdecP));
        end else
        begin
          TOBOUV.PutValue('BLO_DPR',0);
        end;
      end else
      begin
    		TOBOUV.PutValue('BLO_MONTANTPR',TOBOUV.GetValue('BLO_MONTANTPA'));
    		TOBOUV.PutValue('BLO_DPR',TOBOUV.GetValue('BLO_DPA'));
      end;
    end else
    begin
    	TOBOUV.PutValue('BLO_MONTANTPR',TOBOUV.GetValue('BLO_DPR')*TOBOUV.GetValue('BLO_QTEFACT'));
      if TOBOUV.GEtValue('BLO_MONTANTPA') <> 0 then
      begin
        TOBOUV.putValue('BLO_COEFFR',arrondi(TOBOUV.GetValue('BLO_MONTANTPR')/TOBOUV.GEtValue('BLO_MONTANTPA'),4));
      end;
      TOBOUV.putValue('BLO_MONTANTFR',Arrondi(TOBOUV.GetValue('BLO_MONTANTPAFR')*TOBOUV.GEtValue('BLO_COEFFR'),4));
    end;

    if (CalculPv) and (not PVBloque) then
    begin
    	TOBOUV.Putvalue('ANCPV',arrondi(Valloc[2],V_PGI.okdecP));
    	TOBOUV.PutValue('BLO_PUHTDEV',arrondi(Valloc[2],V_PGI.okdecP));
    	TOBOUV.PutValue('BLO_PUHT',arrondi(Valloc[2]/coefDev,V_PGI.okdecP));
    	TOBOUV.PutValue('BLO_PUTTCDEV',arrondi(Valloc[3],V_PGI.Okdecp));
      // calcul des montants Ht,HTDEv,TTC,TTCDEV
    end;
    TOBOUV.PutValue('BLO_PUHT',arrondi(TOBOUV.GetValue('BLO_PUHT'),V_PGI.okdecP));
    TOBOUV.PutValue('BLO_PUTTC',arrondi(TOBOUV.GetValue('BLO_PUTTC'),V_PGI.okdecP));
    TOBOUV.PutValue('BLO_PUHTDEV',arrondi(TOBOUV.GetValue('BLO_PUHTDEV'),V_PGI.okdecP));
    TOBOUV.PutValue('BLO_PUTTCDEV',arrondi(TOBOUV.GetValue('BLO_PUTTCDEV'),V_PGI.okdecP));
    if EnHt then CalculeLigneHTOuv (TOBOuv,TOBPiece,DEV)
            else CalculeLigneTTCOuv (TOBOUV,TOBPiece,DEV);
  end else
  begin
    if TOBOUV.GetValue('BLO_TYPEARTICLE') <> 'POU' then
    begin
      coefMarg := 0;
      CoefDev := 0;
      if (CalCulPv) and (not PVBloque) then
      begin
        CoefMarg := TOBOUV.GetValue('BLO_COEFMARG');
        (*
        if coefmarg = 0 then
        begin
          if TOBOUV.GetValue('BLO_DPR')<>0 then
          begin
            CoefMarg := TOBOUV.GetValue('BLO_PUHT')/TOBOUV.GetValue('BLO_DPR');
          end;
        end;
        *)
        // if TOBOUV.GetValue('BLO_DPR') <> 0 then CoefMarg := TOBOUV.GetValue('BLO_PUHT')/TOBOUV.GetValue('BLO_DPR');
        if TOBOUV.GetValue('BLO_PUHT') <> 0 then CoefDEV := TOBOUV.GetValue('BLO_TAUXDEV');
      end;

      PrestationST := IsPrestationST(TOBOUV);

      if CalculViaAchat then
      begin
        if TOBOUV.GetValue('BLO_NONAPPLICFG')<>'X' then
        begin
          TOBOUV.putValue('BLO_MONTANTFG',Arrondi(TOBOUV.GetValue('BLO_MONTANTPA')*TOBOUV.getValue('BLO_COEFFG'),4));
        end else
        begin
          TOBOUV.putValue('BLO_MONTANTFG',0);
        end;

        MontantPACharge := TOBOUV.GetValue('BLO_MONTANTPA')+ TOBOUV.GetValue('BLO_MONTANTFG');

        if TOBOUV.GetValue('BLO_NONAPPLICFC')<>'X' then
        begin
          TOBOUV.putValue('BLO_COEFFC',TOBPiece.getValue('GP_COEFFC'));
          TOBOUV.putValue('BLO_MONTANTFC',Arrondi(MontantPACharge*TOBOUV.getValue('BLO_COEFFC'),4));
        end else
        begin
          TOBOUV.putValue('BLO_MONTANTFC',0);
        end;

        MontantPACharge := TOBOUV.GetValue('BLO_MONTANTPA')+ TOBOUV.GetValue('BLO_MONTANTFG')+TOBOUV.GetValue('BLO_MONTANTFC');


        if TOBouv.GetValue('BLO_NONAPPLICFRAIS')<>'X' then
        begin
          TOBOUV.putValue('BLO_COEFFR',TOBPiece.getValue('GP_COEFFR'));
          TOBOUV.putValue('BLO_MONTANTFR',Arrondi(MontantPACharge*TOBOUV.getValue('BLO_COEFFR'),4));
        end else
        begin
          TOBOUV.putValue('BLO_MONTANTFR',0);
        end;

        if TOBOUV.GetValue('BLO_MONTANTFG')+TOBOUV.GetValue('BLO_MONTANTFR')+TOBOUV.GetValue('BLO_MONTANTFC') <> 0 then
        begin
          TOBOUV.PutValue('BLO_MONTANTPR',TOBOUV.GetValue('BLO_MONTANTPA')+
                                          TOBOUV.GetValue('BLO_MONTANTFR')+
                                          TOBOUV.GetValue('BLO_MONTANTFG')+
                                          TOBOUV.GetValue('BLO_MONTANTFC'));

          if TOBOUV.GetValue('BLO_QTEFACT') <> 0 then
          begin
            TOBOUV.PutValue('BLO_DPR',arrondi(TOBOUV.GetValue('BLO_MONTANTPR')/TOBOUV.GetValue('BLO_QTEFACT'),V_PGI.okdecP));
          end else
          begin
            TOBOUV.PutValue('BLO_DPR',0);
          end;
        end else
        begin
          TOBOUV.PutValue('BLO_MONTANTPR',TOBOUV.GetValue('BLO_MONTANTPA'));
          TOBOUV.PutValue('BLO_DPR',TOBOUV.GetValue('BLO_DPA'));
        end;
      end else
      begin
        TOBOUV.PutValue('BLO_MONTANTPR',Arrondi(TOBOUV.GetValue('BLO_DPR')* TOBOUV.GetValue('BLO_QTEFACT'),V_PGI.okdecP));
        TOBOUV.putValue('BLO_MONTANTFG',Arrondi(TOBOUV.GetValue('BLO_MONTANTPA')*TOBOUV.GEtValue('BLO_COEFFG'),4));
      end;

      if (CalculPv) and (not PVBloque) then
      begin
        if CoefMarg <> 0 then
        begin
        TOBOUV.Putvalue('ANCPV',Arrondi(TOBOUV.GetValue('BLO_DPR')*CoefMarg*coefdev,V_PGI.OkdecP));

          TOBOuv.putvalue('BLO_PUHT',Arrondi(TOBOUV.GetValue('BLO_DPR')*CoefMarg,V_PGI.OkdecP));
          TOBOuv.putvalue('BLO_PUHTDEV',Arrondi(TOBOUV.GetValue('BLO_DPR')*CoefMarg*CoefDEV,V_PGI.OkdecP));
        end else
        begin
          if TOBOUV.GetValue('BLO_DPR') <> 0 then
          begin
            TOBOUV.PutValue('BLO_COEFMARG',arrondi(TOBOUV.GetValue('BLO_PUHT')/TOBOUV.GetValue('BLO_DPR'),4));
            TOBOuv.PutValue('POURCENTMARG',Arrondi((TOBOuv.GetValue('BLO_COEFMARG')-1)*100,2));
          end else TOBOUV.PutValue('BLO_COEFMARG',0);
        end;
      end else
      begin
        if TOBOUV.GetValue('BLO_DPR') <> 0 then
        begin
          TOBOUV.PutValue('BLO_COEFMARG',arrondi(TOBOUV.GetValue('BLO_PUHT')/TOBOUV.GetValue('BLO_DPR'),4));
          TOBOuv.PutValue('POURCENTMARG',Arrondi((TOBOuv.GetValue('BLO_COEFMARG')-1)*100,2));
        end else TOBOUV.PutValue('BLO_COEFMARG',0);
      end;
      TOBOUV.PutValue('BLO_PUHT',arrondi(TOBOUV.GetValue('BLO_PUHT'),V_PGI.okdecP));
      TOBOUV.PutValue('BLO_PUTTC',arrondi(TOBOUV.GetValue('BLO_PUTTC'),V_PGI.okdecP));
      TOBOUV.PutValue('BLO_PUHTDEV',arrondi(TOBOUV.GetValue('BLO_PUHTDEV'),V_PGI.okdecP));
      TOBOUV.PutValue('BLO_PUTTCDEV',arrondi(TOBOUV.GetValue('BLO_PUTTCDEV'),V_PGI.okdecP));
    end;
    if TOBOUV.GetValue('BLO_PUHT') <> 0 then
    begin
      TOBOUV.PutValue('POURCENTMARQ',Arrondi(((TOBOUV.GetValue('BLO_PUHT')- TOBOUV.GetValue('BLO_DPR'))/TOBOUV.GetValue('BLO_PUHT'))*100,2));
    end else
    begin
      TOBOUV.PutValue('POURCENTMARQ',0);
    end;
    if EnHt then CalculeLigneHTOuv (TOBOuv,TOBPiece,DEV)
            else CalculeLigneTTCOuv (TOBOUV,TOBPiece,DEV);
  end;
  (*
  if CalculViaAchat then
  begin
    if TOBOUV.GetValue('BLO_DPR') <> 0 then TOBOUV.PutValue('BLO_COEFMARG',TOBOUV.GetValue('BLO_PUHT')/TOBOUV.GetValue('BLO_DPR'))
                                       else TOBOUV.PutValue('BLO_COEFMARG',0);
  end;
  *)
  if (CalculPv) then
  begin
  	Valeurs[2] := Valeurs[2]+Arrondi(TOBOUV.GetValue('BLO_MONTANTHTDEV')/QteDudetail,V_PGI.okdecP);
  	Valeurs[3] := Valeurs[3]+Arrondi(TOBOUV.GetValue('BLO_MONTANTTTCDEV')/QteDuDetail,V_PGI.okdecP);
  end;
  Valeurs[10] := Valeurs[10]+Arrondi(TOBOUV.GetValue('BLO_MONTANTPAFG')/QteDUDetail,V_PGI.okdecP);
  Valeurs[11] := Valeurs[11]+Arrondi(TOBOUV.GetValue('BLO_MONTANTPAFR')/QteDuDetail,V_PGI.okdecP);
  Valeurs[12] := Valeurs[12]+Arrondi(TOBOUV.GetValue('BLO_MONTANTPAFC')/QteDuDetail,V_PGI.okdecP);
  Valeurs[13] := Valeurs[13]+Arrondi(TOBOUV.GetValue('BLO_MONTANTFG')/QteDuDetail,V_PGI.okdecP);
  Valeurs[14] := Valeurs[14]+Arrondi(TOBOUV.GetValue('BLO_MONTANTFR')/QteDuDetail,V_PGI.okdecP);
  Valeurs[15] := Valeurs[15]+Arrondi(TOBOUV.GetValue('BLO_MONTANTFC')/QteDuDetail,V_PGI.okdecP);
end;

procedure AppliqueFraisPieceOuvrage(TOBPiece,TOBL,TOBOuvrages : TOB; CalculViaAchat,CalculPV,EnHT : boolean ; DEV : Rdevise; RecalcCoeffg : boolean);
var IndiceNomen,Indice : integer;
		TOBOuvrage : TOB;
    Valeurs : T_valeurs;
    Qte : double;
    ratioHtTTC : double;
begin
  IndiceNomen := TOBL.getValue('GL_INDICENOMEN');
  if IndiceNomen = 0 then exit;
  Qte := TOBL.GetValue('GL_QTEFACT');
  TOBOUvrage := TOBOuvrages.detail[IndiceNomen-1];
  InitTableau (Valeurs);
  if TOBL.GetValue('GL_PUHTDEV') <> 0 then RatioHtTTC := TOBL.GetValue('GL_PUTTCDEV')/TOBL.GetValue('GL_PUHTDEV')
  else RatioHtTTC := 1;
  for Indice := 0 to TOBOUVrage.detail.count -1 do
  begin
  	AppliqueFraisDetailOuv (TOBPiece,TOBOuvrage.detail[Indice],Valeurs,(TOBL.GetValue('GL_BLOQUETARIF')='X'),CalculViaAchat,CalculPv,EnHT,DEV,RecalcCoeffg);
  end;

  if (CalculPv) then
  begin
  	if (EnHt) then
    begin
    	if Valeurs[2] <> TOBL.GetValue('GL_PUHTDEV') then
      begin
      	TOBL.Putvalue('ANCPV', Valeurs[2]);
      	TOBL.PutValue('GL_PUHTDEV',Valeurs[2]);
      	TOBL.PutValue('GL_PUHT',DEVISETOPIVOT (Valeurs[2],DEV.Taux,Dev.Quotite));
    		TOBL.putValue('GL_RECALCULER','X');
      end;
    end else
    begin
    	if Valeurs[3] <> TOBL.GetValue('GL_PUTTCDEV') then
      begin
      	TOBL.Putvalue('ANCPV', Valeurs[3]);
      	TOBL.PutValue('GL_PUTTCDEV',Valeurs[3]);
      	TOBL.PutValue('GL_PUHT',DEVISETOPIVOT (Valeurs[2],DEV.Taux,Dev.Quotite));
    		TOBL.putValue('GL_RECALCULER','X');
      end;
    end;
  end;
  //
  if TOBL.GetValue('GL_DPR') <> 0 then
  begin
    TOBL.putValue('GL_COEFMARG',Arrondi(TOBL.GetValue('GL_PUHT')/TOBL.getValue('GL_DPR'),4));
    TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));
  end;

  if CalculViaAchat then
  begin
    TOBL.PutValue('GL_MONTANTPAFG',Arrondi(Valeurs[10]*Qte,4));
    TOBL.PutValue('GL_MONTANTPAFR',Arrondi(Valeurs[11]*Qte,4));
    TOBL.PutValue('GL_MONTANTPAFC',Arrondi(Valeurs[12]*Qte,4));
    TOBL.PutValue('GL_MONTANTFG',Arrondi(Valeurs[13]*Qte,4));
    TOBL.PutValue('GL_MONTANTFR',Arrondi(Valeurs[14]*Qte,4));
    TOBL.PutValue('GL_MONTANTFC',Arrondi(Valeurs[15]*Qte,4));

    if (RecalcCoefFG) and (TOBL.GetValue('GL_MONTANTPAFG') <> 0) then
    begin
      TOBL.PutValue('GL_COEFFG',Arrondi(TOBL.GetValue('GL_MONTANTFG')/TOBL.GetValue('GL_MONTANTPAFG'),4));
    end;

    if TOBL.GetValue('GL_MONTANTPAFC') <> 0 then
    begin
      TOBL.PutValue('GL_COEFFC',Arrondi(TOBL.GetValue('GL_MONTANTFC')/TOBL.GetValue('GL_MONTANTPAFC'),4));
    end;

    if TOBL.GetValue('GL_MONTANTPAFR') <> 0 then
    begin
      TOBL.PutValue('GL_COEFFR',Arrondi(TOBL.GetValue('GL_MONTANTFR')/TOBL.GetValue('GL_MONTANTPAFR'),4));
    end;

    if TOBL.GetValue('GL_MONTANTFG')+TOBL.GetValue('GL_MONTANTFR')+TOBL.GetValue('GL_MONTANTFC') <> 0 then
    begin
      TOBL.PutValue('GL_MONTANTPR',TOBL.GetValue('GL_MONTANTPA')+TOBL.GetValue('GL_MONTANTFG')+
      														 TOBL.GetValue('GL_MONTANTFR')+TOBL.GetValue('GL_MONTANTFC'));
      if TOBL.GetValue('GL_QTEFACT') <>  0 then
        TOBL.PutValue('GL_DPR',arrondi(TOBL.GetValue('GL_MONTANTPR')/TOBL.GetValue('GL_QTEFACT'),V_PGI.okdecP));
    end else
    begin
    	TOBL.PutValue('GL_MONTANTPR',TOBL.GEtValue('GL_MONTANTPA'));
    	TOBL.PutValue('GL_DPR',TOBL.GEtValue('GL_DPA'));
      TOBL.putValue('GL_COEFFR',0);
      TOBL.putValue('GL_MONTANTFR',0);
    end;
    //
    if TOBL.GetValue('GL_DPR') <> 0 then
    begin
      TOBL.putValue('GL_COEFMARG',Arrondi(TOBL.GetValue('GL_PUHT')/TOBL.getValue('GL_DPR'),4));
      TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));
    end;
    //
  end else
  begin
    TOBL.PutValue('GL_MONTANTPR',Arrondi(TOBL.GetValue('GL_DPR')*TOBL.GetValue('GL_QTEFACT'),4));
    if TOBL.GEtValue('GL_MONTANTPA') <> 0 then
    begin
      TOBL.putValue('GL_COEFFR',Arrondi(TOBL.GetValue('GL_MONTANTPR')/TOBL.GEtValue('GL_MONTANTPA'),9));
    end;
  end;

  if TOBL.GetValue('GL_PUHT') <> 0 then
  begin
    TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('GL_PUHT')- TOBL.GetValue('GL_DPR'))/TOBL.GetValue('GL_PUHT'))*100,2));
  end else
  begin
    TOBL.PutValue('POURCENTMARQ',0);
  end;

end;

procedure CumuleMontantPiece(TOBPiece,TOBL: TOB);
begin
	if TOBL.GetValue('GL_TYPELIGNE')<>'ART' then exit;
  TOBPiece.PutValue('GP_MONTANTPR',TOBPiece.GetValue('GP_MONTANTPR')+TOBL.GetValue ('GL_MONTANTPR'));
	TOBPiece.PutValue('GP_MONTANTFR',TOBPiece.GetValue('GP_MONTANTFR')+TOBL.GetValue ('GL_MONTANTFR'));
	TOBPiece.PutValue('GP_MONTANTFC',TOBPiece.GetValue('GP_MONTANTFC')+TOBL.GetValue ('GL_MONTANTFC'));
	TOBPiece.PutValue('GP_MONTANTFG',TOBPiece.GetValue('GP_MONTANTFG')+TOBL.GetValue ('GL_MONTANTFG'));
end;


Function CalculMtFraisFromLigne (TOBpiece, TOBL : TOB; PUA, CoefFG : Double ; var MontantFG, MontantFR, MontantFC : Double; CalculPv : Boolean=True) : Double;
var DEV             : Rdevise;
    EnHt            : boolean;
    MontantAchat    : Double;
    MontantCharge   : Double;
    Qte             : Double;
    QteDuDetail     : Double;
    CoefMarg        : Double;
    CoefDev         : Double;
    CoefFC          : Double;
    CoefFR          : Double;
begin

  Qte := TOBL.GetValue('GL_QTEFACT');
  if Qte = 0 then exit;

	DEV.Code := TOBPIECE.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);

  MontantFG := 0;
  MontantFR := 0;
  MontantFC := 0;

  CoefFC := TOBPiece.getValue('GP_COEFFC');
  CoefFR := TOBPiece.getValue('GP_COEFFR');

  MontantAchat := Arrondi((Qte*PUA), 4);

  if (CoefFg <> 0) then
  begin
    if TOBL.GetValue('GLC_NONAPPLICFG')<>'X' then MontantFG := Arrondi(MontantAchat * CoefFG, 4);
    MontantCharge := MontantAchat + MontantFG;
  end;

  if TOBL.GetValue('GLC_NONAPPLICFC')<>'X' then MontantFC := Arrondi(MontantCharge * CoefFC, 4);
  MontantCharge := MontantAchat + MontantFG + MontantFC;

  if TOBL.GetValue('GLC_NONAPPLICFR')<>'X' then MontantFR := Arrondi(MontantCharge * CoefFR, 4);
  MontantCharge := MontantAchat + MontantFG + MontantFC + MontantFR;

  Result := Arrondi(Montantcharge / Qte, 4);

end;


procedure CalculFraisFromLigne (TOBpiece,TOBL : TOB; CalculPv : boolean=true);
var CoefDev,CoefMarg,MontantPaCharge,COEFFG,COEFFC,CoefFgLig : double;
    EnHt : boolean;
    DEV : Rdevise;
begin
	DEV.Code := TOBPIECE.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);

  EnHt := (TOBPiece.getValue('GP_FACTUREHT')='X');
  if CalculPv then
  begin
    CoefMarg := TOBL.GetValue('GL_COEFMARG');
    Coefdev := TOBL.GetValue('GL_TAUXDEV');
  end;


  if TOBL.GetValue('GLC_NONAPPLICFG')<>'X' then
  begin
    TOBL.putValue('GL_MONTANTFG',Arrondi(TOBL.GetValue('GL_MONTANTPAFG')*TOBL.getValue('GL_COEFFG'),4));
  end else
  begin
    TOBL.putValue('GL_MONTANTFG',0);
  end;

  MontantPaCharge :=TOBL.GetValue('GL_MONTANTPAFC')+TOBL.GetValue('GL_MONTANTFG');

  if TOBL.GetValue('GLC_NONAPPLICFC')<>'X' then
  begin
    TOBL.putValue('GL_COEFFC',TOBPiece.getValue('GP_COEFFC'));
    TOBL.putValue('GL_MONTANTFC',Arrondi(MontantPaCharge*TOBL.getValue('GL_COEFFC'),4));
  end else
  begin
    TOBL.putValue('GL_COEFFC',0);
    TOBL.putValue('GL_MONTANTFC',0);
  end;

  MontantPaCharge :=TOBL.GetValue('GL_MONTANTPAFR')+TOBL.GetValue('GL_MONTANTFC')+TOBL.GetValue('GL_MONTANTFG');

  if TOBL.GetValue('GLC_NONAPPLICFRAIS')<>'X' then
  begin
    TOBL.putValue('GL_COEFFR',TOBPiece.getValue('GP_COEFFR'));
    TOBL.putValue('GL_MONTANTFR',Arrondi(MontantPaCharge*TOBL.getValue('GL_COEFFR'),4));
  end else
  begin
    TOBL.putValue('GL_COEFFR',0);
    TOBL.putValue('GL_MONTANTFR',0);
  end;

  TOBL.PutValue('GL_MONTANTPR',TOBL.GetValue('GL_MONTANTPA')+TOBL.GetValue('GL_MONTANTFG')+
                                 TOBL.GetValue('GL_MONTANTFR')+TOBL.GetValue('GL_MONTANTFC'));

  if TOBL.GetValue('GL_QTEFACT') <> 0 then
    TOBL.PutValue('GL_DPR',arrondi(TOBL.GetValue('GL_MONTANTPR')/TOBL.GetValue('GL_QTEFACT'),V_PGI.okdecP));
  if (CalculPv) and (TOBL.GetValue('GL_BLOQUETARIF')='-') then
  begin
    if CoefMarg <> 0 then
    begin
      TOBL.putValue('GL_PUHT',Arrondi(TOBL.getValue('GL_DPR')*CoefMarg,V_PGI.okdecP));
      TOBL.putValue('GL_PUHTDEV',Arrondi(TOBL.getValue('GL_DPR')*CoefMarg*CoefDev,V_PGI.okdecP));
      TOBL.putValue('GL_RECALCULER','X');
    end;
  end else
  begin
    if TOBL.GetValue('GL_DPR') <> 0 then
    begin
      TOBL.putValue('GL_COEFMARG',arrondi(TOBL.GetValue('GL_PUHT')/TOBL.getValue('GL_DPR'),4));
      TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));
    end;
  end;
  TOBL.PutValue('GL_PUHT',arrondi(TOBL.GetValue('GL_PUHT'),V_PGI.okdecP));
  TOBL.PutValue('GL_PUTTC',arrondi(TOBL.GetValue('GL_PUTTC'),V_PGI.okdecP));
  TOBL.PutValue('GL_PUHTDEV',arrondi(TOBL.GetValue('GL_PUHTDEV'),V_PGI.okdecP));
  TOBL.PutValue('GL_PUTTCDEV',arrondi(TOBL.GetValue('GL_PUTTCDEV'),V_PGI.okdecP));
  if TOBL.GetValue('GL_PUHT') <> 0 then
  begin
    TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('GL_PUHT')- TOBL.GetValue('GL_DPR'))/TOBL.GetValue('GL_PUHT'))*100,2));
  end else
  begin
    TOBL.PutValue('POURCENTMARQ',0);
  end;
end;

procedure AppliqueFraisPiece (TOBPiece,TOBOuvrages : TOB; CalculViaAchat,CalculPV,EnHt : boolean; DEV : Rdevise ;WithCalcCoefFG : boolean=false);
var indice : integer;
		TOBL : TOB;
    CoefDev,CoefMarg,MontantPaCharge,COEFFR,COEFFC,CoefFGLig : double;
    PrestationST,ApplicFrST,ApplicFcST : boolean;

begin
	ApplicFRSt := (TOBPIEce.getValue('GP_APPLICFGST')='X');
	ApplicFCSt := (TOBPIEce.getValue('GP_APPLICFCST')='X');
  CoeffC := TOBPIEce.getValue('GP_COEFFC');
  CoeffR := TOBPIEce.getValue('GP_COEFFR');

	for Indice := 0 to TOBPiece.detail.count -1 do
  begin
  	TOBL := TOBPIece.detail[Indice];
    if not IsArticle (TOBL) then continue;
    if TOBL.GetString('GL_TYPEARTICLE')='POU' then continue;

    if ((TOBL.GetValue('GL_TYPEARTICLE') = 'OUV') or (TOBL.GetValue('GL_TYPEARTICLE') = 'ARP')) and
    	 (TOBL.GetValue('GL_INDICENOMEN')> 0) then
    begin
    	AppliqueFraisPieceOuvrage (TOBPiece,TOBL,TOBOuvrages,CalculViaAchat,CalculPV,EnHt,DEV,WithCalcCoefFG);
    end else
    begin
      CalculFraisFromLigne (TOBpiece,TOBL,calculPv);
    end;
    CumuleMontantPiece(TOBPiece,TOBL);
  end;
end;

procedure TRecalculPiece.ChargeLesArticles;
var stSelect,StWhereLigne,stWhereLigneOUV : string;
		QArticle : TQuery;
    TOBArtOuv : TOB;
    Indice : integer;
begin
	TOBArtOuv := TOB.Create ('LES ARTICLES',nil,-1);

	stWhereLigne := WherePiece(CleDoc, ttdLigne, False);
	stWhereLigneOUV := WherePiece(CleDoc, ttdOuvrage, False);

  StSelect := 'SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
              'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
              'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE';
  QArticle := OpenSQL(StSelect + ' WHERE GA_ARTICLE IN (SELECT DISTINCT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne + ')', True);
  if not QArticle.EOF then TOBArticles.LoadDetailDB('ARTICLE', '', '', QArticle, True, True);
  Ferme(QArticle);

  Qarticle := OpenSQL(StSelect + ' WHERE GA_ARTICLE IN (SELECT DISTINCT BLO_ARTICLE FROM LIGNEOUV WHERE ' + stWhereLigneOUV + ')', True);
  if not QArticle.EOF then TOBArtOuv.LoadDetailDB('ARTICLE', '', '', QArticle, True, True);
  Ferme(QArticle);
  indice := 0;
  if TOBArtOuv.detail.count = 0 then exit;
  repeat
  	if TOBArticles.findFirst(['GA_ARTICLE'],[TOBArtOuv.detail[Indice].GetValue('GA_ARTICLE')],false) = nil then
    begin
       TOBArtOuv.detail[Indice].ChangeParent (TOBArticles,-1);
    end else Inc(Indice);
  until Indice > TOBArtOuv.detail.count-1;
end;

procedure TRecalculPiece.ChargeTiers;
var Q : TQuery;
begin
  Q := OpenSQL('SELECT * FROM TIERS LEFT JOIN TIERSCOMPL ON T_TIERS=YTC_TIERS WHERE T_TIERS="' + TOBPiece.GetValue('GP_TIERS') + '"', True);
  if not Q.EOF then TOBTiers.SelectDB('', Q);
end;

procedure TRecalculPiece.detruitpiece;
var SQl,RefA : string;
begin
  DetruitCompta(TOBPiece, NowH, OldEcr, OldStk,true);
  //
  if ExecuteSQL('DELETE FROM PIECE WHERE ' + WherePiece(CleDoc, ttdPiece, False)) < 0 then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
  ExecuteSQL('DELETE FROM PIECETRAIT WHERE ' + WherePiece(CleDoc, ttdPieceTrait, False));
  ExecuteSQL('DELETE FROM PIECEINTERV WHERE ' + WherePiece(CleDoc, ttdPieceInterv, False));
  if ExecuteSQL('DELETE FROM LIGNE WHERE ' + WherePiece(CleDoc, ttdLigne, False)) < 0 then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
  if ExecuteSQL('DELETE FROM LIGNECOMPL WHERE ' + WherePiece(CleDoc, ttdLigneCompl, False)) < 0 then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
  if ExecuteSQL('DELETE FROM LIGNEOUV WHERE ' + WherePiece(CleDoc, ttdOuvrage, False)) < 0 then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
  if ExecuteSQL('DELETE FROM LIGNEOUVPLAT WHERE ' + WherePiece(CleDoc, ttdOuvrageP, False)) < 0 then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
  if ExecuteSQL('DELETE FROM LIGNEBASE WHERE ' + WherePiece(CleDoc, ttdLigneBase, False)) < 0 then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
    if ExecuteSQL('DELETE FROM PIEDBASE WHERE ' + WherePiece(CleDoc, ttdPiedBase, False)) < 0 then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
    if ExecuteSQL('DELETE FROM PIEDPORT WHERE ' + WherePiece(CleDoc, ttdPorc, False)) < 0 then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
    if ExecuteSQL('DELETE FROM PIECERG WHERE ' + WherePiece(CleDoc, ttdRetenuG, False)) < 0 then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
  if ExecuteSQL('DELETE FROM PIEDECHE WHERE ' + WherePiece(CleDoc, ttdEche, False)) < 0 then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
  SQl := 'DELETE FROM PIEDBASERG WHERE ' + WherePiece(CleDoc, ttdBaseRG, False);
  if ExecuteSQL(SQL) < 0 then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
  RefA := EncodeRefPresqueCPGescom(TOBPiece);
  ExecuteSQL('DELETE FROM VENTANA WHERE YVA_TABLEANA="GL" AND YVA_IDENTIFIANT LIKE "' + RefA + '"');

end;

procedure TRecalculPiece.EcritSituation;
var Req : string;
    MontantRegl,MontantAcompte,XD,XP,TXD,TXP,TTC : double;
    TOBSIt : TOB;
    QQ : TQuery;
begin
  req := 'SELECT * FROM BSITUATIONS WHERE BST_NATUREPIECE="'+TOBPIece.GetValue('GP_NATUREPIECEG')+'" AND ';
  req := Req + 'BST_SOUCHE="'+TOBPIece.GetVAlue('GP_SOUCHE')+'" AND BST_NUMEROFAC="'+inttoStr(TOBPiece.GetValue('GP_NUMERO'))+'"';
  QQ := OpenSql(Req,true,-1,'',true);
  if not QQ.eof then
  begin
    GetMontantsAcomptes (TOBAcomptes,MontantAcompte,MontantRegl);
    GetMontantRG (TOBPieceRG,TOBBasesRG,XD,XP,DEV,True,True);
    GetcumultaxesRG (TOBBasesRG,TOBPieceRG,TXD,TXP,DEV);
    TOBSit := TOB.create ('BSITUATIONS',nil,-1);
    TOBSIt.SelectDB ('',QQ);
    TTC := TOBPiece.GetValue('GP_TOTALTTCDEV') - XD - TXD;
    TOBSIt.putvalue('BST_MONTANTHT',TOBPiece.GetValue('GP_TOTALHTDEV'));
    TOBSIt.putvalue('BST_MONTANTTVA',TTC - TOBPiece.GetValue('GP_TOTALHTDEV'));
    TOBSIt.putvalue('BST_MONTANTTTC',TTC);
    TOBSIT.PutValue('BST_MONTANTREGL',MontantRegl);
    TOBSIT.PutValue('BST_MONTANTACOMPTE',MontantAcompte);
    TOBSit.SetAllModifie (true);
    TOBSIT.UpdateDB (false);
    TOBSIT.free;
  end;
  ferme (QQ);
end;

procedure TRecalculPiece.validelaPiece;
begin
  detruitpiece;
  //
  ValideLesOuvPlat (TOBOuvragesP , TOBpiece);
  ValideLesBases(TOBPiece,TobBases,TOBBasesL);
  ValideLesPieceTrait(TOBPiece,TOBaffaire,TOBPieceTrait,TOBSSTRAIT,DEv);
  ValideLesSousTrait (TOBPiece,TOBSSTRAIT,DEv);
  ValideLesRetenues(TOBPiece, TOBPieceRG);
  ValideLesBasesRG(TOBPiece, TOBBasesRG);
  GereEcheancesGC(TOBPiece, TOBTiers, TOBEches, TOBAcomptes, TOBPieceRG,TOBPieceTrait,TOBPorcs, taModif, DEV, False);
  //
  if V_PGI.IoError = OeOk then EcritLapiece;
  if V_PGI.IOError = oeOk then EcritSituation;
end;

procedure TRecalculPiece.EcritLapiece;
begin
  TOBPiece.SetAllModifie(true);
  if not TOBPiece.InsertDBByNivel (true) then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
  ValideLesLignesCompl (TOBPiece,nil);
  if V_Pgi.IOError <> OeOk then exit;
  if not TOBOuvrages.InsertDBByNivel (true) then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
  if not TOBOuvragesP.InsertDBByNivel (true) then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
  if not TOBBasesL.InsertDB (nil,true) then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
  if not TOBBases.InsertDB (nil,true) then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
//  if not TOBPieceTrait.InsertDB (nil,true) then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
  if not TOBPorcs.InsertDB (nil,true) then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
  if not TOBEches.InsertDB(nil) then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;

//  if not TOBPieceRG.InsertDB(nil,true) then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
//  if not TOBBasesRG.InsertDB(nil,true) then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
  if V_PGI.IoError = oeOk then
  begin
    if not PassationComptable( TOBPiece,TOBOUvrages,TOBOuvragesP, TOBBases,TOBBasesL, TOBEches,TOBPieceTrait,
                               TOBAffaireInterv,TOBTiers, TOBArticles, TOBCpta, TOBAcomptes, TOBPorcs,
                               TOBPIECERG, TOBBASESRG, TOBanaP,TOBanaS,TOBSSTRAIT,TOBVTECOLL, DEV, OldEcr,
                               OldStk, True, false ) then
    begin
      V_PGI.Ioerror := OeSaisie; Exit;
    end;
  end;
  if (V_PGI.IoError = oeOk) and (TOBVTECOLL.Detail.count > 0) then
  begin
    PrepareInsertCollectif (TOBPiece,TOBVTECOLL);
    if not TOBVTECOLL.InsertDB(nil) then BEGIN V_PGI.Ioerror := OeSaisie; Exit; END;
  end;
end;


procedure TRecalculPiece.SauvegardeContexte;
begin
	OuvrageInterne  := TOBOuvrages;
  PortInterne     := TOBPorcs;
  ArticlesInterne := TOBArticles  ;
  TiersInterne    := TOBTiers;
  PieceRGInterne  := TOBPieceRG;
  BasesRGInterne  := TOBBasesRG ;
  BasesInterne    := TOBBases;
end;

procedure TRecalculPiece.RestitueContexte;
begin
	TOBOuvrages  := OuvrageInterne;
  TOBPorcs     := PortInterne;
  TOBArticles  := ArticlesInterne;
  TOBTiers     := TiersInterne;
  TOBPieceRG   := PieceRGInterne;
  TOBBasesRG   := BasesRGInterne;
  TOBBases     := BasesInterne;
end;

procedure TRecalculPiece.DefiniContexteCalcul (LaTOBPiece,LaTOBOuvrage,LesPorts,LesArticles,LeTiers,LEsPieceRG,LesBasesRG,LesBases : TOB);
begin
  LaPiece     := LaTOBPiece;
  TOBOuvrages  := LaTOBOuvrage;
  TOBPorcs     := LesPorts;
  TOBArticles  := LesArticles;
  TOBTiers     := LeTiers;
  TOBPieceRG   := LEsPieceRG;
  TOBBasesRG   := LesBasesRG;
  TOBBases     := LesBases;
end;

procedure TRecalculPiece.RecalculPieceFromSaisie(LaTOBPiece,LaTOBOuvrage,LesPorts,LesArticles,LeTiers,
																								 LEsPieceRG,LesBasesRG,LesBases: TOB);
begin
	// Sauvegarde des pointeurs
  SauvegardeContexte;
  TRY
    // Mise en place des tob à recalculer
    DefiniContexteCalcul (LaTOBPiece,LaTOBOuvrage,LesPorts,LesArticles,LeTiers,LEsPieceRG,LesBasesRG,LesBases);
    RecalculePAPRPV;
    //
  FINALLY
  	RestitueContexte;
  End;
end;

// -- FONCTIONS PERMETTANT DE METTRE A JOUR LES INFOS LIGNES D'UNE PIECE AVANT DE LA RECALCULER ET MAJ

function TRecalculPiece.ChargeLapiece(lacledoc: r_cledoc) : integer;
var existFg : boolean;
begin
  result := 0;
	InitPiece;
	cledoc := Lacledoc;
  if not LoadPiece(Cledoc,TOBPiece) then BEGIN result := -1; exit; END;
  LoadLesTObs;
  if OuvrageNonDifferencie (TOBpiece,TOBOuvrages) then   // au cas ou
  begin
    OuvrageDifferencie (TOBpiece,TOBOuvrages,false,DEV);
  end;
  if Pos(TOBPiece.getValue('GP_NATUREPIECEG'),'ETU;DBT;BCE') > 0 then
  begin
  	MontantFraisChantier := GetMontantFraisDetail (TOBPiece,ExistFg);
  end;
  EnHt := (TOBPiece.getValue('GP_FACTUREHT')='X');
  DEV.Code := TOBPIECE.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := TOBPiece.getValue('GP_TAUXDEV');
end;

function TRecalculPiece.FindInPiece ( InfoLigne: RinfoLigne ) : TOB;
var IIndNumOrdre,Indice : integer;
begin
  result := nil;
  if TOBPiece.detail.count = 0 then exit;
  IIndNumOrdre := TOBpiece.detail[0].GetNumChamp('GL_NUMORDRE');
	for Indice := 0 to TOBpiece.detail.count -1 do
  begin
  	if TOBPiece.detail[Indice].GetInteger(IIndNumOrdre) = InfoLigne.Ligne then
    result := TOBPiece.detail[Indice];
  end;
end;

procedure TRecalculPiece.ReajusteSousDetail (TOBL : TOB; var LeMontant : double);
var IndiceNomen : Integer;
		TOBOUV : TOB;
begin
  TraitePrixOuvrage(TOBPiece,TOBL,TOBBases,TOBBasesL,TOBOuvrages, EnHt, leMontant,0,DEV,true);
  if not CalculPv then ReinitCoefMarg (TOBL,TOBOuvrages);
end;

procedure TRecalculPiece.ReajusteSousDetailPV (TOBL : TOB; var LeMontant : double; MtBase : double);
var IndiceNomen : Integer;
		TOBOUV : TOB;
begin
  TraitePrixOuvrage(TOBPiece,TOBL,TOBBases,TOBBasesL,TOBOuvrages, EnHt, leMontant,MtBase,DEV,false,true);
  ReinitCoefMarg (TOBL,TOBOuvrages);
end;

procedure TRecalculPiece.ReajusteSousDetailOuvrage (TOBL,TOBO : TOB; var LeMontant : double);
var ValeurAnc ,valeurPr : double;
begin
  ValeurAnc := TOBO.GetValue('BLO_DPA');
  ValeurPr := 0;
	ReajusteMontantOuvrage (TOBArticles ,TOBpiece,TOBL,TOBO,ValeurAnc,ValeurPr,LeMontant,DEV,EnHt,true,false,true);
end;

function TRecalculPiece.FindOuvrage (TOBL : TOB; InfoLigne : RinfoLigne) : TOB;
var IndiceNomen,Indice,IIndUniqueBlo : integer;
		TOBOUV : TOB;
begin
	result := nil;
  IndiceNomen := TOBL.GetValue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
  TOBOUV := TOBOuvrages.detail[IndiceNomen-1]; if TOBOUV.detail.count = 0 then exit;
  IIndUniqueBlo := TOBOUV.detail[0].GetNumChamp('BLO_UNIQUEBLO');
  for Indice := 0 to TOBOUV.detail.count - 1 do
  begin
		if TOBOUV.detail[Indice].GetInteger (IIndUniqueBlo) = InfoLigne.UniqueBlo then
    begin
      result := TOBOUV.detail[Indice];
      break;
    end;
  end;
end;

function TRecalculPiece.MajInfoLigne(InfoLigne: RinfoLigne; TypeInfo: TTypeInfo; valeur: variant) : integer;
var TOBL,TOBO : TOB;
		TheDouble : double;
    TheString : string;
begin
  result := 0;
  TOBL := FindInPiece (infoLigne); // la ligne de document;
  if TOBL = nil then BEGIN result := -1; exit; END;
  if InfoLigne.UniqueBlo <> 0 then
  begin
    // mise à jour d'un sous détail d'ouvrage
    TOBO := FindOuvrage (TOBL, infoLigne);
    if TOBO = nil then BEGIN result := -2; Exit; END;
    TheDouble := Valeur;
    if TypeInfo = TTiPua then
    begin
      if IsOuvrage(TOBO) then ReajusteSousDetailOuvrage (TOBL,TOBO,TheDouble);
      TOBO.SetDouble('BLO_DPA',TheDouble);
    end else if TypeInfo = TTiQte then
    begin
      TOBO.SetDouble ('BLO_QTEFACT',TheDouble);
    end else if TypeInfo = TTiUnite then
    begin
			TheString := Valeur;
      TOBO.SetString ('BLO_QUALIFQTEVTE',TheString);
    end;
  end else
  begin
		// mise à jour d'une ligne de document
    if TypeInfo = TTiPua then
    begin
			TheDouble := Valeur;
      if IsOuvrage(TOBL) then ReajusteSousDetail (TOBL,TheDouble);
			TOBL.SetDouble('GL_DPA',TheDouble);
    end else if TypeInfo = TTiQte then
    begin
			TheDouble := Valeur;
			TOBL.SetDouble('GL_QTEFACT',TheDouble);
			TOBL.SetDouble('GL_QTESTOCK',TheDouble);
			TOBL.SetDouble('GL_QTERESTE',TheDouble);
    end else if TypeInfo = TTiUnite then
    begin
			TheString := Valeur;
			TOBL.SetString('GL_QUALIFQTEVTE',TheString);
    end;
  end;
end;

procedure TRecalculPiece.CreateEnv;
begin
	TOBPiece := TOB.Create ('PIECE',nil,-1);
end;

procedure TRecalculPiece.DestroyEnv;
begin
	TOBPiece.free;
end;

procedure TRecalculPiece.ReInitEnv;
begin
	TOBPiece.ClearDetail;
  TOBPiece.InitValeurs;
end;

procedure TRecalculPiece.ControleEtReajusteOuvrages;
var Indice : Integer;
begin
	for Indice := 0 to TOBPiece.detail.count -1 do
  begin
		if isOuvrage (TOBPiece.detail[Indice]) and (not IsDejaFacture(TOBPiece.detail[Indice])) then
    begin
      ControleEtReajusteSousDetail (TOBPiece.detail[Indice]);
    end;
  end;
end;

procedure TRecalculPiece.ControleEtReajusteSousDetail(TOBL: TOB);
var IndiceNomen : Integer;
		TOBOUv : TOB;
    Indice : Integer;
    OkFind,OkFindPou : Boolean;
    PuDoc : double;
begin
  OkFind := false;
  IndiceNomen := TOBL.GetValue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
  TOBOUV := TOBOuvrages.detail[IndiceNomen-1]; if TOBOUV.detail.count = 0 then exit;
  Indice := 0;
  PuDoc := TOBL.GetDouble('GL_PUHTDEV');

	repeat
  	if TOBOUv.detail[Indice].GetInteger('BLO_N1')=0 then
    begin
      // Traitement de l'ane au mali
      TOBOUv.detail[Indice].Free;
      OkFind := True;
      Break;
    end else if (ArtEcart <> '') and (TOBOUV.detail[Indice].GetString('BLO_ARTICLE')=ArtEcart) then
    begin
      TOBOUv.detail[Indice].Free;
      OkFind := True;
      Break;
    end else if (TOBOUV.detail[Indice].GetString('BLO_TYPEARTICLE')='POU') then
    begin
      OkFind := true;
      break;
    end else inc(Indice);
  until Indice >= TOBOUv.detail.Count;

  if OkFind then
  begin
    TraitementOuvrage (TOBL,true);
    ReajusteSousDetailPV (TOBL,PuDoc,TOBL.GetDouble('GL_PUHTDEV'));
  end;
end;

function TRecalculPiece.IsDejaFacture(TOBL: TOB): Boolean;
begin
	Result :=  ((TOBL.GetValue('GL_QTEPREVAVANC') <> 0) or
  						(TOBL.GetValue('GL_QTESIT') <> 0)) and (TOBPiece.GetValue('GP_NATUREPIECEG')='DBT');
end;

end.

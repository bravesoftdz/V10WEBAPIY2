unit BTFactImprTOB;

interface
uses

  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids,UTOB,
  SaisUtil, HEnt1, hmsgbox,Hctrls,
  Menus, EntGC,AGLInitGC,UtilPgi,Aglinit,DicoBTP,FactUtil,HRichOle,HTB97,
{$IFDEF EAGLCLIENT}
  MaineAGL,
{$ELSE}
  Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
  {$IFDEF V530}
    EdtDOC,
  {$ELSE}
    EdtRdoc,
  {$ENDIF}
{$ENDIF}

{$IFDEF BTP}
  CalcOleGenericBTP,
{$ENDIF}
  Splash,FactTvaMilliem,
  Math,ParamSoc
  ,uEntCommun
  ;

Type

TImprPieceViaTOB = class
	private
    fcledocFacCur : R_CLEDOC;
  	fcledoc       : r_cledoc;
    fNumeroSit  : integer;
    fIsDGD      : boolean;
    fisAVA      : boolean; // pour déclancher les traitement sur DGD avancement
  	fUsable     : boolean;
    fInternal   : boolean;
    fcreation   : boolean;
    fAvancement : boolean;
    WithRecap   : boolean;

    Symbole     : string;
  	venteAchat  : string;
    DEV : Rdevise;

  	fTOBPieceImp      : TOB;
    fTOBPieceRecap    : TOB;
    fTobPiece,fPreTOBPiece : TOB;
    ftobParPiece      : TOB;
    fTOBAdresses      : TOB;
    fTOBTiers         : TOB;
    fTOBLIENOLE       : TOB;
    fTOBArticles      : TOB;
    fTOBPieceRG       : TOB;
    fTOBBases         : TOB;
    fTOBEches         : TOB;
    fTOBPorcs         : TOB;
    fTOBConds         : TOB;
    fTOBComms         : TOB;
    fTOBOuvrage       : TOB;
    fTOBNomenclature  : TOB;
    fTOBLOTS          : TOB;
    fTOBSerie         : TOB;
    fTOBAcomptes      : TOB;
    fTOBAffaire       : TOB;
    fTOBAffaireDev    : TOB;
    fTOBBasesRG       : TOB;
    fTOBFormuleVar    : TOB;
    fTOBMetre         : TOB;
    fTOBRepart        : TOB;
    fTOBCodeTaxes     : TOB;
    fTOBmetres        : TOB;
    fTOBSituations    : TOB;
    fTOBTimbres       : TOB;
    fTOBSsTraitant    : TOB;
    fTobSsTraitantImp : TOB;
    fTOBpieceTrait    : TOB;
    ftobRetenueDiverse: TOB;
    //
    fLastNumordre   : integer;
    fTexte          : THRichEditOle;
    fInternalWindow : TToolWindow97;
    fForm           : TForm;
    TheRepartTva    : TREPARTTVAMILL;
    //
    fNbSTTraitDir : integer; // nb sous traitant en paiement direct
    //
    procedure AjouteRetenuesDiv (TOBMere,TOBAInsere : TOB);
    procedure AjouteChampsParamDoc(TOBMere, TOBAInsere: TOB);
    procedure AjouteChampsPiece(TOBMere, TOBAInsere: TOB);
		procedure AjouteChampsMillieme (TOBMere, TOBAInsere: TOB);
		procedure AjouteTOB (TOBMere : TOB; CodeTraitement : string; TOBAInsere : TOB; TOBPrec : TOB = nil);
    function  InsereTOB(NomTable: string; TOBmere: TOB; WithNewNumOrdre : boolean=true): TOB;
    procedure AjouteChampsDOCTEXTEDEBFIN(TOBmere, TOBAInsere: TOB);
    procedure AjouteChampsINTEXTEDEBFIN(TOBmere, TOBAInsere: TOB);
    procedure AjouteChampsTiers(TOBmere, TOBAInsere: TOB);
    procedure CumuleChampsAvancLigne(TOBmere, TOBAInsere: TOB);
    procedure AjouteChampsAvancementPiece (TOBMere : TOB);
    procedure AjouteChampsAffaire(TOBmere, TOBAInsere: TOB);
    procedure AjouteChampsAffaireDEVIS(TOBmere, TOBAInsere: TOB);
    procedure AjouteChampsPiedBase(TOBmere, TOBAInsere: TOB);
    procedure AjouteChampsPiedBaseRG(TOBmere, TOBAInsere: TOB);
    procedure AjouteChampsPieceAdresse(TOBmere, TOBAInsere: TOB);
    procedure AjouteChampsPiedEche(TOBmere, TOBAInsere: TOB);
    procedure AjouteChampsPiedPort(TOBmere, TOBAInsere: TOB);
    procedure AjouteChampsPieceRG(TOBmere, TOBAInsere: TOB);
    procedure AjouteChampsParPiece(TOBmere, TOBAInsere: TOB);
    procedure AjouteChampsDEVISE(TOBmere: TOB);
    procedure AjouteChampsLIGNE(TOBmere, TOBAINSERE: TOB; LigneFac : boolean=false);
    procedure AjouteChampsARTICLES(TOBmere, TOBAINSERE: TOB);
    procedure AjouteChampsOUVRAGE(TOBmere, TOBAINSERE,TOBLigne : TOB);
    procedure AjouteChampsMETRE(TOBmere: TOB);
    procedure prepareTVA;
    procedure PreparePiedbase (TOBMere : TOB);
    procedure PrepareSsTraite (TOBMere : TOB);

    function  FindLibre(Prefixe,Code : string) : string;
    procedure AddSupLignesEdt(TOBMere: TOB);
    procedure DefiniAffichageLigne(TOBmere: TOB);
//    procedure ReDefiniNumLig (ThePiece : TOB);
//    function GetLastNumOrdre : integer;
		function IsBlocNoteVide (BlocNote : String; Fichier : boolean=False) : boolean;
    procedure CreeAdresseFacturation(TobMere: TOB);
    procedure AjouteChampsPARAG(TOBmere, TOBAInsere: TOB);
    procedure AlimenteTotauxParag(ThePiece, TOBparag: TOB);
    function  FindLaFinParag(TheLigne: TOB; IndiceDep: integer): TOB;
    procedure PoselesTotaux(ThePiece,TheLigne, LaTotalisation: TOB);
    procedure AlimentePort(ThePiece: TOB);
    procedure AjoutelesPORTS(TOBmere, TOBAInsere: TOB);
    procedure InsereChampMere(TOBMere, LaTOBgrandMere: TOB);
    procedure AlloueTout;
    procedure DetruitTout;
    procedure loadlesTobs;
		procedure RazTout;
    procedure LoadLesArticles (ThePiece: TOB);
    procedure InitModePaiement (TOBMere : TOB);
    function GetInfoParams: boolean;
    procedure SetMode;
    procedure SetMontantMarche(TOBmere, LaLigne: TOB);
    procedure AlimenteTableauSituations(ThePiece: TOB);
    procedure InitChamps(TOBL: TOB; Nomchamps,TypeChamps : string);
    procedure AjouteChampsSituation(TOBMere: TOB);
    procedure DefiniLigneSituation(TOBMere, TOBSituation: TOB);
    procedure AjouteChampsAcomptesCpta(TOBMere: TOB);
    function SetQueryForOldSituations : Widestring;
    procedure SetLignesRecap(ThePiece: TOB);
    procedure DefiniRecapTotalSit(ThePiece: TOB);
    procedure DefiniLigneTotalHt(ThePiece: TOB);
    procedure DefiniLigneMontantHt(ThePiece: TOB; EscompteCum, RemiseCum,PortCum,CumPortSit,MTRDSITHT,MTRDPREHT: double);
    procedure DefiniLigneRemise (ThePiece : TOB;RemiseCum: double);
    procedure DefiniLigneEscompte (ThePiece : TOB;EscompteCum : double);
    procedure DefiniLignePortFrais(ThePiece: TOB; PortCum,CumPortSit: double);
    procedure ConstitueTOBTVA(ThePiece, TOBTVAPre, TOBTVAPRETTC: TOB; var MtHtPortTTCPre,MtHTPortTTC,MtPortTTCPre,MtPortTTC : double);
    procedure SetMontantregleSituations;
    procedure ConstitueTOBRGPrec(ThePiece, TOBRGPre: TOB);
    procedure DefiniLigneTVARD (ThePiece,TOBFDDET: TOB; var MTSIT,MTPRE : double);
    procedure DefiniLigneTVARG(ThePiece, TOBRGPRE: TOB;var TVARGSIt,TVARGPre : double);
    procedure DefiniLignesTVA(ThePiece, TOBTVAPre,TOBTVAPRETTC,TOBRDDET: TOB;var CumSitTva,CumPreTva: double);
    procedure DefiniLigneRDHT(ThePiece,TOBFraisDet :TOB; Var ValSit,VAlpre : double);
    procedure DefiniLigneRDTTC(ThePiece,TOBFraisDet :TOB; Var ValSit,VAlpre : double);
    procedure DefiniLigneRGHT(ThePiece, TOBRGPRE: TOB; var CumSitHt,
      CumpreHt, ReliquatRGSit, reliquatRGPre: double);
    procedure DefiniLigneRGTTC(ThePiece,TOBRGPRE : TOB; var ReliquatRgSit,ReliquatRgPre : double);
    procedure DefiniLigneTOTALTTC(ThePiece,TOBfDetaille: TOB; CumSitTVA, CumPreTVA,
      CumRGSit, CumRgPre,TVARGSIt,TVARGPre,MTRDSIT,MTRDPRE,TVARDSIT,TVARDPRE,MtHtPortTTCPre,MtHTPortTTC: double; var TTCSIt, TTCPRE: double);
    procedure DefiniLigneACOMPTES(ThePiece: TOB; var AcomptesSit,
      AcomptesPre: double);
    procedure DefiniLigneTOTALNET(ThePiece: TOB; TTCSit, TTCPre,
      ReliquatRGSit, reliquatRGPre, AcomptesSit, AcomptesPre,Timbres,MTRDSITTTC,MTRDPRETTC,MtPortTTCPre,MtPortTTC: double);
    procedure SetMontantMarcheOldVersion(TOBmere : TOB);
    function IsExistsMemoFacturation: boolean;
    function IsSamePiece(reference, cledoc: r_cledoc): boolean;
    procedure CumuleBudgetFromCledoc(Cledoc: R_cledoc; TOBMere: TOB);
    function isExistsMemo(LaLigne: TOB): boolean;
    procedure PositionneMarcheFromPiece(TheRecap, ThePiece: TOB);
    procedure GestionAnciennesfactureation (TOBL : TOB; Modegenere : string; DEV : RDevise);
    procedure AjoutelesMETRESLIG(TOBmere, TOBAInsere: TOB);
    procedure AlimentemetresLig (ThePiece,TheLigne: TOB; Numordre : integer; TOBmetres : TOB);
    procedure AjouteleMtTimbre(TOBmere, TOBAInsere: TOB);
    function DefiniLigneTimbres(ThePiece, TOBTimbres: TOB) : double;
    procedure AjouteChampsSsTraitant(Tobmere, TOBAinsere: TOB);
    procedure AjouteChampsSsTraitantInterne(TOBmere,TOBAInsere : TOB);
    function ImpDecodeRib (RibEncode : string) : string;
    procedure ConstitueFraisDetaille(ThePiece, TOBFD: TOB);
    procedure ConstitueRD (ThePiece,TOBFDDET : TOB; var NbFDHT : integer);
    procedure DefiniLignePortFraisDet(ThePiece, TOBFraisDet: TOB;TypePort: string;var NbrPorts: integer; Sens:string='+');
    procedure DefiniLigneTotalHtDet(ThePiece, TOBFDetaille: TOB);
    procedure DefiniLignesTVADet(ThePiece, TOBTVAPre, TOBFDetaille, TOBRDDET: TOB);
    procedure DefiniLigneTOTALTTCInterm(ThePiece,TOBFDetaille: TOB; CumSitTVA,
      CumPreTVA, CumRGSit, CumRgPre, TVARGSIt, TVARGPre,MTRDSIT,MTRDPRE,TVARDSIT,TVARDPRE,MtHtPortTTCPre,MtHTPortTTC: double; var TTCSIt, TTCPRE: double);
    procedure DefiniLignePIECEREPART(ThePiece: TOB; EscompteCum,RemiseCum,PortCum,CumPortSit,MTRDSITHT,MTRDPREHT : double);
    procedure AjouteSsTraitant(Tobmere, TOBAinsere: TOB);
    procedure AjouteTobSsTraitant(Tobmere, TOBAinsere: TOB);
    procedure AjouteSsTraitantInterne(Tobmere, TOBAinsere: TOB);
    procedure AjouteTobSsTraitantInterne(Tobmere, TOBAinsere: TOB);
    procedure AjouteRetenuesDiverses(TOBMere, TobAInsere: TOB);
    function NbSSTraitDir (TOBpieceTrait : TOB) : integer;
    procedure GetInfosTVAFromRD(CodeTva: string; TOBRDDET: TOB;
      var BaseTvaRDPre, MtTvaRDPre, BaseTvaRDSit, MTTVARDSIT: double);
    function GetDevisFromprepaFac(TOBL: TOB): string;
    procedure AjouteRetenuesTTC(TOBMere, TOBAInsere: TOB);

  public
  	property Usable : boolean read GetInfoParams write fusable;
  	property TOBPiece : TOB read fpreTOBPiece write fPreTOBPiece;
    property TOBAdresses : TOB read fTOBAdresses write fTOBAdresses;
    property TOBTiers : TOB read fTOBTiers write fTOBTiers;
    property TOBLIENOLE : TOB read fTOBLIENOLE write fTOBLIENOLE;
    property TOBArticles : TOB read fTOBArticles write fTOBArticles;
    property TOBPieceRG : TOB read fTOBPieceRG write fTOBPieceRG;
    property TOBBasesRG : TOB read fTOBBasesRG write fTOBBasesRG;
    property TOBBases : TOB read fTOBBases write fTOBBases;
    property TOBEches : TOB read fTOBEches write fTOBEches;
    property TOBPorcs : TOB read fTOBPorcs write fTOBPorcs;
    property TOBConds : TOB read fTOBConds write fTOBConds;
    property TOBComms : TOB read fTOBComms write fTOBComms;
    property TOBParPiece : TOB read fTOBParPiece write fTOBParPiece;
    property TOBOuvrage : TOB read fTOBOuvrage write fTOBOuvrage;
    property TOBNomenclature : TOB read fTOBNomenclature write fTOBNomenclature;
    property TOBLOTS : TOB read fTOBLOTS write fTOBLOTS;
    property TOBSerie : TOB read fTOBSerie write fTOBSerie;
    property TOBAcomptes : TOB read fTOBAcomptes write fTOBAcomptes;
    property TOBAffaire : TOB read fTOBAffaire write fTOBAffaire;
    property TOBFormuleVar : TOB read fTOBFormuleVar write fTOBFormuleVar;
    property TOBAIMPRIMER : TOB read fTOBPieceImp write fTOBPieceImp;
    property InCreation : boolean read fcreation write fcreation;
    property TOBRECAP : TOB read fTOBPieceRecap write fTOBPieceRecap;
    property TobSsTraitantImp : TOB read fTobSsTraitantImp  write fTobSsTraitantImp;
    //
    property TobSsTraitant    : TOB read fTOBSsTraitant     write fTobSsTraitant;
    property TobRetenueDiverse: TOB read ftobRetenueDiverse write ftobRetenueDiverse;
    property TOBtimbres : TOB read fTOBTimbres write fTOBTimbres;
    procedure AssocieTOBs (TOBPiece,TOBBases,TOBEches,TOBPorcs,TOBTiers,
                           TOBArticles,TOBComms,TOBAdresses,TOBNomenClature,
                           TOBDesLots,TOBSerie,TOBAcomptes,TOBAffaire,TOBOuvrage,TOBLienOle,
                           TOBPieceRg,TOBBasesRG,TOBFormuleVar,TOBMetre,TOBRepart,TOBmetres,TOBTimbres,TOBPieceTrait: TOB);
  	constructor create (TheFiche : TForm);
    destructor Destroy; override;
    //
    procedure PreparationEtat; // appel de baaaase..
    //
    procedure SetDocument(Cledoc : R_cledoc; IsDgd : Boolean);
    procedure ClearInternal;
    function GetModeleAssocie(Modele : string) : string;
    function YaMetre : boolean;
    procedure VirelesLignesAZero (TobFacture : TOB);
    procedure SuprimeParagraphesOrphelin (TOBFacture : TOB);
end;

function EditionViaTOB (cledoc : r_cledoc) : boolean;
function ImpressionAutoriseViaTOB (cledoc : r_cledoc):boolean;

implementation
uses
//UtilMetres,
      FactVariante,FactAdresse,FactNomen,FactTob,OptimizeOuv,FactComm,FactOuvrage,Affaireutil,
			Factcalc,FactPiece,STockUtil,FactCOmmBTP,UtilTOBPiece,BTSAISIEPAGEMETRE_TOF,
      FactTimbres,FactRG,UtilReglementAffaire,FactRetenues,UCotraitance,Ent1
      ,CbpMCD
      ,CbpEnumerator
      ;
{$IFDEF BTP}
//var TheMetre: TStockMetre;
{$ENDIF}

function ImpressionAutoriseViaTOB (cledoc : r_cledoc):boolean;
begin
	Result := (Cledoc.Naturepiece = 'DBT') or
  					(Cledoc.Naturepiece = 'ETU') or
            (Cledoc.Naturepiece = 'FBT') or
            (Cledoc.Naturepiece = 'FBP') or
            (Cledoc.Naturepiece = 'BAC') or
            (Cledoc.Naturepiece = 'BCE') or
            (Cledoc.Naturepiece = 'DAC') or
            (Cledoc.Naturepiece = 'ABT') or
            (Cledoc.Naturepiece = 'ABP') or
            (Cledoc.Naturepiece = 'DE')  or
            (Cledoc.Naturepiece = 'CC')  or
            (Cledoc.Naturepiece = 'BLC') or
            (Cledoc.Naturepiece = 'FBC') or
            (Cledoc.Naturepiece = 'B00') or
            (Cledoc.Naturepiece = 'ABC');
end;

function EditionViaTOB (cledoc : r_cledoc):boolean;
var Req : string;
		Q : TQuery;
    LaTOB : TOB;
    SoucheG : string;
begin
	result := false;
  SoucheG := GetSoucheG(Cledoc.NaturePiece,'','');
  //
  Req := 'SELECT BPD_NUMPIECE,BPD_IMPRVIATOB FROM BTPARDOC WHERE BPD_NATUREPIECE="'+Cledoc.naturepiece +'" AND '+
  			 '((BPD_SOUCHE="'+Cledoc.Souche +'") OR (BPD_SOUCHE="'+SoucheG+'") OR (BPD_SOUCHE="")) AND '+
         '((BPD_NUMPIECE=0) OR (BPD_NUMPIECE='+InttoStr(Cledoc.NumeroPiece)+')) ORDER BY BPD_NUMPIECE DESC';
  Q := OpenSql (Req,True,-1,'',true);

  if not Q.eof then
  begin
		LaTOB := TOB.Create ('LES PARAMS',nil,-1);
    laTOB.LoadDetailDB ('BTPARDOC','','',Q,false);
    ferme(Q);
    result := (LaTOB.Detail[0].GetValue('BPD_IMPRVIATOB')='X');
  	LaTOB.free;
  end;
end;


{ TImprPieceViaTOB }

procedure TImprPieceViaTOB.DefiniAffichageLigne(TOBmere : TOB);
var TypePresent : Integer;
		TOBParent : TOB;
begin
	TOBParent := TOBmere.parent;

  if (TOBMERE.GetValue('ISLIGNESTD')='X') OR (TOBMERE.GetValue('ISVARIANTE')='X')  or
     (TOBMEre.GetValue ('ISDETOUVVAR')='X') OR (TOBMEre.GetValue ('ISDETOUVSTD')='X') then
  begin
  	TOBMere.PutValue('DESCRIPTIF','X');

  	if (TOBParent.GetValue('BPD_IMPDESCRIPTIF')='S') then
    begin
    TOBMere.PutValue('GL_BLOCNOTE','');
    end else if (TOBParent.GetValue('BPD_IMPDESCRIPTIF')='T') then
    begin
			if (TOBMere.getValue('GL_BLOCNOTE') = '') or (IsBlocNoteVide(TOBMere.getValue('GL_BLOCNOTE'))) then
      begin
      	TOBMere.PutValue('DESCRIPTIF','X');
        if (TOBMere.getValue('GA_BLOCNOTE') <> '') and (not IsBlocNoteVide(TOBMere.getValue('GA_BLOCNOTE'))) then
        begin
        	TOBMere.PutValue('GL_BLOCNOTE',TOBMere.GetValue('GA_BLOCNOTE'));
        	TOBMere.PutValue('GA_BLOCNOTE',''); // evite de surcharger la memoire
        end;
      end;
    end else TOBMere.PutValue('DESCRIPTIF','X');

    if (TOBMere.getValue('GL_BLOCNOTE') = '') or (IsBlocNoteVide(TOBMere.getValue('GL_BLOCNOTE'))) then
    begin
      TOBMere.PutValue('GL_BLOCNOTE', '');
      TOBMere.PutValue('DESCRIPTIF','-');
    end;
  end;

  if (TOBMere.getValue('GL_BLOCNOTE') <> '') and (not IsBlocNoteVide(TOBMere.getValue('GL_BLOCNOTE'))) then
  begin
    TOBMERE.PutValue('WITHBLOB','X');
    if TOBMere.GetValue('ISCOMMVARIANTE')='X' then
    begin
    	TOBMERE.PutValue('ISCOMMENTAIRE','X');
    	TOBMERE.PutValue('ISCOMMVARIANTE','-');
    end;
  end;

	if (TOBmere.GetValue('ISDETOUVSTD')='X') OR (TOBmere.GetValue('ISDETOUVVAR')='X') then
  begin
  	TypePresent := TOBParent.GetValue('BPD_TYPESSD');
    if TypePresent = -1 Then TypePresent := DOU_AUCUN else
    if TypePresent = 0 then TypePresent := TOBMere.GetValue('GL_TYPEPRESENT');
  end else
  begin
  	TypePresent := TOBParent.GetValue('BPD_TYPEPRES');
  end;
// TODO Voir en fonction du parametrage de l'edition
//
  if (TypePresent AND DOU_CODE) <> DOU_CODE Then
  begin
    TObMere.putValue('GL_REFARTTIERS','');
    TObMere.putValue('GL_REFARTSAISIE','');
  end;
  if (TypePresent AND DOU_NUMP) = DOU_NUMP Then
  begin
    TObMere.putValue('GL_REFARTTIERS',TObMere.GetString('GLC_NUMEROTATION'));
    TObMere.putValue('GL_REFARTSAISIE',TObMere.GetString('GLC_NUMEROTATION'));
  end;
  if (TypePresent AND DOU_LIBELLE) <> DOU_LIBELLE Then
  begin
    if (TOBMERE.GetValue('ISFINPARAGRAPHE')='-') and ((TOBMERE.GetValue('ISLIGNESTD')='X') or (TOBMERE.GetValue('ISVARIANTE')='X')) then
    	TObMere.putValue('GL_LIBELLE','');
  end;
  TOBMere.PutValue('LIBELLE',TObMere.GetValue('GL_REFARTSAISIE')+' '+TObMere.GetValue('GL_LIBELLE'));
  if (TypePresent AND DOU_QTE) <> DOU_QTE Then
  begin
    TObMere.putValue('GL_QTEFACT',0);
  end else
  begin
  	TOBMere.PutValue('TITREQTE','Quantité');
  end;
  if (TypePresent AND DOU_UNITE) <> DOU_UNITE Then
  begin
    TObMere.putValue('GL_QUALIFQTEVTE','');
  end;
  if (TypePresent AND DOU_PU) <> DOU_PU Then
  begin
    TObMere.putValue('GL_PUHTDEV',0);
    TObMere.putValue('GL_PUTTCDEV',0);
  end else
  begin
  	TOBMere.PutValue('TITREPU','Prix Unit.');
  end;

  if (TypePresent AND DOU_MONTANT) <> DOU_MONTANT Then
  begin
		if not IsParagraphe (TOBMere) then
  	begin
    	TObMere.putValue('GL_MONTANTHTDEV',0);
    	TObMere.putValue('GL_MONTANTTTCDEV',0);
    end;
  end else
  begin
  	TOBMere.PutValue('TITREMONTANT','Montant');
  end;
end;

procedure TImprPieceViaTOB.AddSupLignesEdt (TOBMere : TOB);
begin
  //
  AjouteTOB (TOBMere,'ARTICLES',nil); // définition des champs dans la tob
  AjouteTOB (TOBMere,'LIGNES',nil);
  AjouteTOB (TOBMere,'LIGNEFAC',nil);
  AjouteTOB (TOBMere,'TABLEAUSITUATIONS',nil);
  AjouteTOB (TOBMere,'METRESLIG',nil);
  //
  TOBMEre.AddChampSupValeur ('ISCUMULABLE','-');
  TOBMEre.AddChampSupValeur ('MONTANTPOURREPORT',0);
  //
	TOBMere.AddChampSupValeur ('DEFINIMETRE','');
  TOBMEre.AddChampSupValeur ('ISDEBPARAGRAPHE','-');
  TOBMEre.AddChampSupValeur ('ISDEBPARVARIANT','-');
  TOBMEre.AddChampSupValeur ('ISFINPARAGRAPHE','-');
  TOBMEre.AddChampSupValeur ('ISFINPARVARIANT','-');
  TOBMEre.AddChampSupValeur ('ISPARAGRAPHE','-');
  TOBMEre.AddChampSupValeur ('ISPARVARIANT','-');

  TOBMEre.AddChampSupValeur ('ISDETOUVSTD','-');
  TOBMEre.AddChampSupValeur ('ISDETOUVVAR','-');
  // Commentaires standard
  TOBMEre.AddChampSupValeur ('ISCOMMENTAIRE','-');
  TOBMEre.AddChampSupValeur ('ISCOMMVARIANTE','-');
  // BLOB Ligne renseigne
  TOBMEre.AddChampSupValeur ('WITHBLOB','-');
  // ---
  TOBMEre.AddChampSupValeur ('ISLIGNESTD','-');
  TOBMEre.AddChampSupValeur ('ISLIGNEPOR','-'); // ligne de frais détaillé
  TOBMEre.AddChampSupValeur ('ISVARIANTE','-');
  TOBMEre.AddChampSupValeur ('LIBELLE','');
  TOBMEre.AddChampSupValeur ('TITREQTE','');
  TOBMEre.AddChampSupValeur ('TITREPU','');
  TOBMEre.AddChampSupValeur ('TITREMONTANT','');
	TobMere.addChampSupValeur ('TAX_CODETVA','');
	TobMere.addChampSupValeur ('TX1_CODETVA','');
	TobMere.addChampSupValeur ('TX2_CODETVA','');
	TobMere.addChampSupValeur ('TX3_CODETVA','');
	TobMere.addChampSupValeur ('TX4_CODETVA','');
	TobMere.addChampSupValeur ('TX5_CODETVA','');

  TOBMere.AddChampSupValeur ('DESCRIPTIF','-');
  TOBMere.AddChampSupValeur ('HASMETRE','-');
  TOBMEre.AddChampSupValeur ('ISDEBEDTPAR','-');
  TOBMEre.AddChampSupValeur ('ISFINEDTPAR','-');
  TOBMEre.AddChampSupValeur ('ISLIGEDTPAR','-');
  TOBMEre.AddChampSupValeur ('ISLIGPORT','-');
  TOBMere.AddChampSupValeur ('INDICEMETRE',0);
  TOBMere.AddChampSupValeur ('ISDETOUVVAR','-');
  TOBMere.AddChampSupValeur ('ISDETOUVSTD','-');
  // ------
  // recapitulatif des situations
  // ------
  TOBMere.AddChampSupValeur ('DEBUTRECAPSIT','-');
  TOBMere.AddChampSupValeur ('LIGNERECAPSIT','-');
  TOBMere.AddChampSupValeur ('FINRECAPSIT','-');
  TOBMere.AddChampSupValeur ('ACOMPTESIT',0);
  TOBMere.AddChampSupValeur ('REGLEMENTSIT',0);
  //
  TOBMEre.AddChampSupValeur ('MONTANTSIT',0);
  TOBMEre.AddChampSupValeur ('LIBELLESIT','');
  TOBMEre.AddChampSupValeur ('MONTANTPRE',0);
  TOBMEre.AddChampSupValeur ('LIBELLEPRE','');
  TOBMEre.AddChampSupValeur ('LIBELLECUM','');
  // -----
  // GESTION DES METRES LIGNES (ex Gamme II)
  // ----
  TOBMEre.AddChampSupValeur ('ISDEBMETRE','-');
  TOBMEre.AddChampSupValeur ('ISFINMETRE','-');
  TOBMEre.AddChampSupValeur ('ISLIGMETRE','-');
  // ----
  TOBMEre.AddChampSupValeur ('PUHTDEVNETORI',0.0);
  TOBMEre.AddChampSupValeur ('PUHTDEVORI',0.0);
  TOBMEre.AddChampSupValeur ('PUTTCDEVNETORI',0.0);
  TOBMEre.AddChampSupValeur ('PUTTCDEVORI',0.0);
  //
end;

procedure TImprPieceViaTOB.AjouteChampsARTICLES (TOBmere,TOBAINSERE : TOB);
var ITable,Indice,IndiceTOB : integer;
  Mcd : IMCDServiceCOM;
  FieldList : IEnumerator ;
  NomChamps,TypeChamps : string;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  FieldList := Mcd.getTable(PrefixeToTable('GA')).Fields;
  FieldList.Reset();
  While FieldList.MoveNext do
  begin
    NomChamps := (FieldList.Current as IFieldCOM).name;
    TypeChamps := (FieldList.Current as IFieldCOM).Tipe;
		if not TOBMere.FieldExists (NomChamps) then
    begin
      TOBMere.AddChampSup (NomChamps,false);
      InitChamps (TOBmere,NomChamps,TypeChamps);
    end;
    if TOBAInsere <> nil then
    begin
      TOBMere.PutValue (NomChamps, TOBAInsere.GetValue(NomChamps));
    end;
  end;
end;


procedure TImprPieceViaTOB.AjouteChampsLIGNE(TOBmere, TOBAINSERE: TOB; LigneFac : boolean=false);
var ITable,Indice,IndiceTOB: integer;
		QQ : TQuery;
    SQl : String;
    LCledoc : r_cledoc;
    Mcd : IMCDServiceCOM;
    FieldList : IEnumerator ;
    NomChamps,TypeChamps : string;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  if LigneFac then
  begin
    FieldList := Mcd.getTable(PrefixeToTable('BLF')).Fields;
    //FV1 - 26/09/2017 - FieldList.Reset();
    While FieldList.MoveNext do
    begin
      NomChamps := (FieldList.Current as IFieldCOM).name;
      TypeChamps := (FieldList.Current as IFieldCOM).Tipe;
      if not TOBMere.FieldExists (NomChamps) then
      begin
        TOBMere.AddChampSup (NomChamps,false);
        InitChamps (TOBmere,NomChamps,TypeChamps);
      end;
      if TOBAinsere <> nil then
      begin
        TOBMere.PutValue (NomChamps,TOBAInsere.GetValue(NomChamps));
      end;
    end;
    if TOBmere.GetValue('GL_TYPELIGNE')='ART' then
    begin
      TOBMere.PutValue ('MONTANTPOURREPORT',TOBMere.getValue('BLF_MTCUMULEFACT'));
    end;
  end else
  begin
    FieldList := Mcd.getTable(PrefixeToTable('GL')).Fields;
    if (TOBAINSERE <> nil) and (TOBAINSERE.GetValue('GL_PIECEPRECEDENTE')<>'') then
    begin
      DecodeRefPiece(TOBAINSERE.GetString('GL_PIECEPRECEDENTE'),Lcledoc);
			Sql := 'SELECT GL_PUHTDEV,GL_PUHTNETDEV,GL_PUTTCDEV,GL_PUTTCNETDEV FROM LIGNE WHERE '+WherePiece (Lcledoc,ttdLigne,True,True);
			QQ := OpenSQL(SQl,True,1,'',true);
      if Not QQ.Eof then
      begin
        TOBMEre.SetDouble ('PUHTDEVNETORI',QQ.findField('GL_PUHTNETDEV').AsFloat);
        TOBMEre.SetDouble ('PUHTDEVORI',QQ.findField('GL_PUHTDEV').AsFloat);
        TOBMEre.SetDouble ('PUTTCDEVNETORI',QQ.findField('GL_PUTTCNETDEV').AsFloat);
        TOBMEre.SetDouble ('PUTTCDEVORI',QQ.findField('GL_PUTTCDEV').AsFloat);
      end;
      ferme (QQ);
    end;

    FieldList.Reset();
    While FieldList.MoveNext do
    begin
      NomChamps := (FieldList.Current as IFieldCOM).name;
      TypeChamps := (FieldList.Current as IFieldCOM).Tipe;
  //  	if V_PGI.Dechamps[iTable, Indice].Nom = 'GL_NUMORDRE' then continue;
      if not TOBMere.FieldExists (NomChamps) then
      begin
        TOBMere.AddChampSup (NomChamps,false);
        InitChamps (TOBmere,NomChamps,TypeChamps);
      end;
      if TOBAinsere <> nil then
      begin
        TOBMere.PutValue (NomChamps,TOBAInsere.GetValue(NomChamps));
      end;
    end;
    if TOBAinsere <> nil then
    begin
      if not TOBMere.FieldExists ('GLC_NUMEROTATION') then
      begin
        TOBMere.AddChampSupvaleur ('GLC_NUMEROTATION',TOBAInsere.GetValue('GLC_NUMEROTATION'));
      end else
      begin
        TOBMere.PutValue ('GLC_NUMEROTATION',TOBAInsere.GetValue('GLC_NUMEROTATION'));
      end;
    	// METRES
      TOBMere.PutValue ('INDICEMETRE',TOBAInsere.GetValue('INDICEMETRE'));
      //
      if (IsVariante(TOBMERE)) then
      begin
        if IsArticle(TOBMere) then TOBMEre.PutValue ('ISVARIANTE','X') else
        if ISCommentaire(TOBMERE) then TOBMEre.PutValue ('ISCOMMVARIANTE','X') else
        if IsDebutParagraphe(TOBMERE) then TOBMEre.PutValue ('ISDEBPARVARIANT','X') else
        if IsFinParagraphe(TOBMERE) then TOBMEre.PutValue ('ISFINPARVARIANT','X');
      end else
      begin
        if ISArticle(TOBMERE) then BEGIN TOBMEre.PutValue ('ISLIGNESTD','X'); TOBMEre.PutValue ('ISCUMULABLE','X'); end else
        if ISSousTotal(TOBMERE) then TOBMEre.PutValue ('ISLIGNESTD','X') else
        if ISCommentaire(TOBMERE) then TOBMEre.PutValue ('ISCOMMENTAIRE','X') else
        if IsDebutParagraphe(TOBMERE) then TOBMEre.PutValue ('ISDEBPARAGRAPHE','X') else
        if IsFinParagraphe(TOBMERE) then TOBMEre.PutValue ('ISFINPARAGRAPHE','X');
      end;
      if TOBMere.GetValue('ISLIGNESTD')='X' then
      begin
      	if TOBMere.GetValue ('GL_FAMILLETAXE1') <> '*' then
        begin
        TOBMere.PutValue('TX1_CODETVA',findLibre('TX1',TOBMere.GetValue ('GL_FAMILLETAXE1')));
        end else
        begin
        TOBMere.PutValue('TX1_CODETVA',TOBMere.GetValue ('GL_FAMILLETAXE1'));
        end;
        TOBMere.PutValue('TAX_CODETVA',TOBMere.GetValue('TX1_CODETVA'));
        if TOBMere.GetValue ('GL_FAMILLETAXE2') <> '' then
        begin
          TOBMere.PutValue('TX2_CODETVA',findLibre('TX2',TOBMere.GetValue ('GL_FAMILLETAXE2')));
          TOBMere.PutValue('TAX_CODETVA','*');
        end;
        if TOBMere.GetValue ('GL_FAMILLETAXE3') <> '' then
        begin
          TOBMere.PutValue('TX3_CODETVA',findLibre('TX3',TOBMere.GetValue ('GL_FAMILLETAXE3')));
          TOBMere.PutValue('TAX_CODETVA','*');
        end;
        if TOBMere.GetValue ('GL_FAMILLETAXE4') <> '' then
        begin
          TOBMere.PutValue('TX4_CODETVA',findLibre('TX4',TOBMere.GetValue ('GL_FAMILLETAXE4')));
          TOBMere.PutValue('TAX_CODETVA','*');
        end;
        if TOBMere.GetValue ('GL_FAMILLETAXE5') <> '' then
        begin
          TOBMere.PutValue('TX5_CODETVA',findLibre('TX5',TOBMere.GetValue ('GL_FAMILLETAXE5')));
          TOBMere.PutValue('TAX_CODETVA','*');
        end;
      end;
      DefiniAffichageLigne(TOBmere);
    end;
  end;
end;

procedure TImprPieceViaTOB.AjouteChampsPARAG (TOBmere,TOBAInsere : TOB);
begin
  TobMere.dupliquer(TOBAInsere,false,true);
  if TOBMere.GetValue('ISDEBPARAGRAPHE')='X' then
  begin
    TOBMere.PutValue('ISDEBPARAGRAPHE','-');
    TOBMere.PutValue('ISPARAGRAPHE','X');
  end else if TOBMere.GetValue('ISDEBPARVARIANT')='X' then
  begin
    TOBMere.PutValue('ISDEBPARVARIANT','-');
    TOBMere.PutValue('ISPARVARIANT','X');
  end;
end;

procedure TImprPieceViaTOB.AjouteChampsDEVISE (TOBmere : TOB);
var QQ : Tquery;
		REQ : String;
begin
	Req := 'SELECT D_SYMBOLE FROM DEVISE WHERE D_DEVISE="'+ fTOBPiece.GEtVAlue('GP_DEVISE')+'"';
  QQ := OpenSql(Req,True,-1,'',true);
  TOBMere.AddChampSup ('D_SYMBOLE',false);
  if not QQ.eof then TOBMere.putValue ('D_SYMBOLE',QQ.findField('D_SYMBOLE').AsString);
  Ferme(QQ);
  Symbole := TOBMere.GetValue ('D_SYMBOLE');
end;

procedure TImprPieceViaTOB.AjouteChampsAffaire (TOBmere,TOBAInsere : TOB);
var ITable,Indice : integer;
  Mcd : IMCDServiceCOM;
  FieldList : IEnumerator ;
  NomChamps,TypeChamps : string;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  FieldList := Mcd.getTable(PrefixeToTable('AFF')).Fields;
  FieldList.Reset();
  While FieldList.MoveNext do
  begin
    NomChamps := (FieldList.Current as IFieldCOM).name;
    TypeChamps := (FieldList.Current as IFieldCOM).Tipe;
		if not TOBMere.FieldExists (NomChamps) then
    begin
      TOBMere.AddChampSupValeur (NomChamps,TOBAInsere.GetValue(NomChamps));
    end else
    begin
      TOBMere.PutValue (NomChamps,TOBAInsere.GetValue(NomChamps));
    end;
  end;
  {$IFDEF BTP}
  TOBMere.AddChampSupValeur ('AFF_AFFAIREFORMATE',BTPCodeAffaireAffiche(TobMere.GetValue('AFF_AFFAIRE')));
  {$ELSE}
  TOBMere.AddChampSupValeur ('AFF_AFFAIREFORMATE',CodeAffaireAffiche(TobMere.GetValue('AFF_AFFAIRE')));
  {$ENDIF}
end;

procedure TImprPieceViaTOB.AjouteChampsAffaireDEVIS (TOBmere,TOBAInsere : TOB);
var ITable,Indice : integer;
		Suffixe : string;
    TOBLOC : TOB;
    Mcd : IMCDServiceCOM;
    FieldList : IEnumerator ;
    NomChamps,TypeChamps : string;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  FieldList := Mcd.getTable(PrefixeToTable('AFF')).Fields;
  FieldList.Reset();

  if TOBAInsere = nil then
  begin
    // cas de figure  : Impression d'un document en cours de création
    TOBLOc := TOB.Create ('AFFAIRE',nil,-1);
  end;

  While FieldList.MoveNext do
  begin
    NomChamps := (FieldList.Current as IFieldCOM).name;
    TypeChamps := (FieldList.Current as IFieldCOM).Tipe;

    suffixe := Copy(NomChamps,Pos('_',NomChamps)+1,255);
    if TOBAInsere <> nil then
    begin
      TOBMere.AddChampSupValeur ('IAFF_'+Suffixe,
                                 TOBAInsere.GetValue(NomChamps));
    end else
    begin
      TOBMere.AddChampSupValeur ('IAFF_'+Suffixe,
                                 TOBLoc.GetValue(NomChamps));
    end;
  end;

  if TOBAInsere = nil then
  begin
    freeAndNil(TOBLOc);
  end;
end;

procedure TImprPieceViaTOB.AjoutelesMETRESLIG(TOBmere,TOBAInsere : TOB);
var ITable,Indice,IndiceTOB : integer;
    Mcd : IMCDServiceCOM;
    FieldList : IEnumerator ;
    NomChamps,TypeChamps : string;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  FieldList := Mcd.getTable(PrefixeToTable('BLM')).Fields;
  FieldList.Reset();

  While FieldList.MoveNext do
  begin
    NomChamps := (FieldList.Current as IFieldCOM).name;
    TypeChamps := (FieldList.Current as IFieldCOM).Tipe;
  	//
  	if Pos(NomChamps ,'BLM_NATUREPIECEG;BLM_SOUCHE,BLM_NUMERO,BLM_INDICEG,BLM_NUMORDRE,BLM_INDICE')>0 then continue;
    // --
		if not TOBMere.FieldExists (NomChamps) then
    begin
      TOBMere.AddChampSup (NomChamps,false);
      InitChamps (TOBmere,NomChamps,TypeChamps);
    end;
    if TOBAInsere <> nil then
    begin
      TOBMere.PutValue (NomChamps,TOBAInsere.GetValue(NomChamps));
    end;
  end;
end;


procedure TImprPieceViaTOB.AjouteleMtTimbre (TOBmere,TOBAInsere : TOB);
var Mttimbre : Double;
		indice : Integer;
begin
  Mttimbre := 0;
  for Indice := 0 to TOBAInsere.detail.count -1 do
  begin
    Mttimbre := Mttimbre + TOBAInsere.detail[Indice].getValue('BT0_VALEUR');
  end;
  if not TOBMere.FieldExists ('VALTIMBRES') then
  begin
    TOBMere.AddChampSupValeur ('VALTIMBRES',Mttimbre);
  end else
  begin
    TOBMere.PutValue ('VALTIMBRES',Mttimbre);
  end;
end;

procedure TImprPieceViaTOB.AjoutelesPORTS (TOBmere,TOBAInsere : TOB);
var iTable,Indice : integer;
		LaTOBgrandMere : TOB;
    prefixe,newChamp,TypePort : String;
    Mcd : IMCDServiceCOM;
    FieldList : IEnumerator ;
    NomChamps,TypeChamps : string;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  LaTOBGrandMere := Tobmere.Parent;
  TypePort := TOBAinsere.GetString('GPT_TYPEPORT');
  InsereChampMere(TOBMere,LaTOBgrandMere);

  FieldList := Mcd.getTable(PrefixeToTable('GPT')).Fields;
  FieldList.Reset();

  While FieldList.MoveNext do
  begin
    NomChamps := (FieldList.Current as IFieldCOM).name;
    TypeChamps := (FieldList.Current as IFieldCOM).Tipe;
  	Prefixe := 'GL';
    NewChamp := prefixe+copy(NomChamps,4,255);
		if TOBMere.FieldExists (NewChamp) then
    		TOBMere.PutValue (NewChamp,TOBAInsere.GetValue(NomChamps));
  end;
  if (TypePort ='HT') or (TypePort='MI') or (TypePort='MT') or (TypePort = '') then TOBMEre.PutValue('GL_FACTUREHT','X')
  else if (TypePort='PT') or (TypePort = 'MIC') or (TypePort ='MTC') then TOBMEre.PutValue('GL_FACTUREHT','-');
  TOBMEre.PutValue ('ISLIGPORT','X');
end;

procedure TImprPieceViaTOB.AjouteChampsMETRE (TOBmere : TOB);
var LaLigneOle : TOB;
begin
	if TOBmere = nil then exit;
	if TOBMere.GetValue('INDICEMETRE')=0 then exit;
	laLigneOle := fTOBMetre.findfirst (['INDICEMETRE'],[TOBMere.getValue('INDICEMETRE')],true);
  if (LaLigneOle <> nil) then TOBMere.PutValue('DEFINIMETRE',laLigneOle.getValue('LO_OBJET'));
  TOBMere.PutValue ('HASMETRE','X');
	if (TOBMere.GetValue('DEFINIMETRE')='') Or (IsBlocNoteVide(TOBMere.GetValue('DEFINIMETRE'),True)) then
  	TOBMere.PutValue ('HASMETRE','-');
end;
           
procedure TImprPieceViaTOB.AjouteChampsTiers (TOBmere,TOBAInsere : TOB);
var ITable,Indice : integer;
    Mcd : IMCDServiceCOM;
    FieldList : IEnumerator ;
    NomChamps,TypeChamps : string;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  FieldList := Mcd.getTable(PrefixeToTable('T')).Fields;
  FieldList.Reset();

  While FieldList.MoveNext do
  begin
    NomChamps := (FieldList.Current as IFieldCOM).name;
    TypeChamps := (FieldList.Current as IFieldCOM).Tipe;
  	//
  	if NomChamps = 'T_BLOCNOTE' then continue;
    // --
		if not TOBMere.FieldExists (NomChamps) then
    begin
      TOBMere.AddChampSupValeur (NomChamps,
      													 TOBAInsere.GetValue(NomChamps));
    end else
    begin
      TOBMere.PutValue (NomChamps,TOBAInsere.GetValue(NomChamps));
    end;
  end;
end;

procedure TImprPieceViaTOB.AjouteChampsPiece (TOBMere,TOBAInsere : TOB);
var 
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
  NomChamps : string;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  Table := Mcd.GetTable( Mcd.PrefixeToTable('GP'));

  FieldList := Table.Fields;
  FieldList.Reset();
  While FieldList.MoveNext do
  begin
		NomChamps := (FieldList.Current as IFieldCOM).name;
		if not TOBMere.FieldExists (NomChamps) then
    begin
      TOBMere.AddChampSupValeur (NomChamps,TOBAInsere.GetValue(NomChamps));
    end else
    begin
      TOBMere.PutValue (NomChamps,TOBAInsere.GetValue(NomChamps));
    end;
  end;
  TOBMere.putValue('GP_NUMERO',fcledocFacCur.NumeroPiece );
  TOBMere.AddChampSupValeur ('DECODERIB',ImpDecodeRib(fTOBPiece.getString('GP_RIB')));
  TOBMere.AddChampSupValeur ('RETENUESTTC',0);
  TOBMere.AddChampSupValeur ('RETENUESDIVTTC',0);
  TOBMere.AddChampSupValeur ('RETENUESDIVHT',0);
  TOBMere.AddChampSupValeur ('RETENUEHTPRE',0);
end;

procedure TImprPieceViaTOB.AjouteChampsPiedEche (TOBmere,TOBAInsere : TOB);
var ITable,Indice,Num : integer;
		ChampsOrig,ChampsDest,Suffixe : string;
    LaTOBOrig : TOB;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
	InitModePaiement (TOBmere);
  Table := Mcd.GetTable(Mcd.PrefixeToTable('GPE'));

  For Num := 0 To TOBAInsere.detail.count -1 do
  begin
  	LATOBOrig := TOBAInsere.detail[Num];
    Suffixe := IntToStr (LATOBOrig.GetValue('GPE_NUMECHE'));
    //FV1 - 26/09/2017
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
    begin
      ChampsOrig := (FieldList.Current as IFieldCOM).name;
      ChampsDest := ChampsOrig+Suffixe;
      if not TOBMere.FieldExists (ChampsDest) then
      begin
        TOBMere.AddChampSupValeur (ChampsDest,LATOBOrig.GetValue(ChampsOrig));
      end else
      begin
        TOBMere.PutValue (ChampsDest,LATOBOrig.GetValue(ChampsOrig));
      end;
      if TOBMere.getValue('GPE_MODEPAIE1') <> '' then
      	TOBMere.AddChampSupValeur ('GPE_LIBMODEPAIE1',RechDom('GCMODEPAIE',TOBMere.getValue('GPE_MODEPAIE1'),false));
    end;
  end;
end;

procedure TImprPieceViaTOB.AjouteChampsPiedBase (TOBmere,TOBAInsere : TOB);
var ITable,Indice,Num,IndiceTaxe,Cpt : integer;
		ChampsOrig,ChampsDest,Suffixe,Prefixe,LastPrefixe,Code,Chpt : string;
    LaTOBOrig,LaTOBRepart,LaTOBInterm,TOBI,TOBBRG : TOB ;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  //FV1 - 26/09/2017
  //FieldList := Mcd.getTable(PrefixeToTable('GPB')).Fields;
  //FieldList := Table.Fields;
  //FieldList.Reset();

  LaTOBInterm := TOB.Create ('LES BASEs',nil,-1);
  TOBBRG := TOB.Create ('LES BASES RG',nil,-1);
  if Pos(fTOBPiece.getValue('GP_NATUREPIECEG'),'FBT;DAC;FBP;BAC;')>0 then
  begin
    ConstitueTVARG (fTOBPieceRG,fTOBBasesRG,TOBBRG);
  end;

  For Num := 0 To TOBAInsere.detail.count -1 do
  begin
    TOBI := LaTOBInterm.FindFirst(['GPB_CATEGORIETAXE','GPB_FAMILLETAXE'],
    															[TOBAInsere.detail[Num].GetValue('GPB_CATEGORIETAXE'),
                                   TOBAInsere.detail[Num].GetValue('GPB_FAMILLETAXE')],True);
  	if TOBI = nil then
    begin
      TOBI := TOB.Create('PIEDBASE',LaTOBInterm,-1);
      TOBI.Dupliquer(TOBAInsere.detail[Num],False,true);
    end else
    begin
      TOBI.PutValue('GPB_BASETAXE',TOBI.GetValue('GPB_BASETAXE')+TOBAInsere.detail[Num].GetValue('GPB_BASETAXE'));
      TOBI.PutValue('GPB_VALEURTAXE',TOBI.GetValue('GPB_VALEURTAXE')+TOBAInsere.detail[Num].GetValue('GPB_VALEURTAXE'));
      TOBI.PutValue('GPB_BASEDEV',TOBI.GetValue('GPB_BASEDEV')+TOBAInsere.detail[Num].GetValue('GPB_BASEDEV'));
      TOBI.PutValue('GPB_VALEURDEV',TOBI.GetValue('GPB_VALEURDEV')+TOBAInsere.detail[Num].GetValue('GPB_VALEURDEV'));
      TOBI.PutValue('GPB_BASETAXETTC',TOBI.GetValue('GPB_BASETAXETTC')+TOBAInsere.detail[Num].GetValue('GPB_BASETAXETTC'));
      TOBI.PutValue('GPB_BASETTCDEV',TOBI.GetValue('GPB_BASETTCDEV')+TOBAInsere.detail[Num].GetValue('GPB_BASETTCDEV'));
      TOBI.PutValue('GPB_BASETTCDEV',TOBI.GetValue('GPB_BASETTCDEV')+TOBAInsere.detail[Num].GetValue('GPB_BASETTCDEV'));
    end;
  end;

  For Num := 0 To TOBBRG.detail.count -1 do
  begin
    TOBI := LaTOBInterm.FindFirst(['GPB_CATEGORIETAXE','GPB_FAMILLETAXE'],
    															[TOBBRG.detail[Num].GetValue('CATEGORIETAXE'),
                                   TOBBRG.detail[Num].GetValue('FAMILLETAXE')],True);
    if TOBI <> nil then
    begin
      TOBI.PutValue('GPB_BASETAXE',TOBI.GetValue('GPB_BASETAXE')-TOBBRG.detail[Num].GetValue('BASETAXE'));
      TOBI.PutValue('GPB_VALEURTAXE',TOBI.GetValue('GPB_VALEURTAXE')-TOBBRG.detail[Num].GetValue('VALEURTAXE'));
      TOBI.PutValue('GPB_BASEDEV',TOBI.GetValue('GPB_BASEDEV')-TOBBRG.detail[Num].GetValue('BASEDEV'));
      TOBI.PutValue('GPB_VALEURDEV',TOBI.GetValue('GPB_VALEURDEV')-TOBBRG.detail[Num].GetValue('VALEURDEV'));
    end;
  end;

	Tobmere.AddChampSupValeur('PRESENCEMILLIEME','');
  iTable := PrefixeToNum('GPB');
  IndiceTaxe := 0;
  Prefixe := '';
  LastPrefixe := '';
  Cpt := 1;
  if not TOBmere.fieldExists('TVA_AUTOLIQUIDST') then
  begin
  	TOBMEre.AddChampSupValeur ('TVA_AUTOLIQUIDST','NON');
  end else
  begin
  	TOBMEre.PutValue ('TVA_AUTOLIQUIDST','NON');
  end;

  TOBMEre.AddChampSupValeur ('TAX_LCODETVA'+inttoStr(Num),'');
  for Num := 1 to 5 do
  begin
  	if not TOBmere.fieldExists('TAX_LCODETVA'+inttostr(Num)) then
    begin
      TOBMEre.AddChampSupValeur ('TAX_LCODETVA'+inttoStr(Num),'');
      TOBMEre.AddChampSupValeur ('TAX_BASETVA'+inttoStr(Num),'');
      TOBMEre.AddChampSupValeur ('TAX_TAUXTAXE'+inttoStr(Num),'');
      TOBMEre.AddChampSupValeur ('TAX_VALEURTAXE'+inttoStr(Num),'');
  		Tobmere.AddChampSupValeur ('TAX_REPART'+inttoStr(Num),0);
    end else
    begin
      TOBMEre.PutValue ('TAX_LCODETVA'+inttoStr(Num),'');
      TOBMEre.PutValue ('TAX_BASETVA'+inttoStr(Num),'');
      TOBMEre.PutValue ('TAX_TAUXTAXE'+inttoStr(Num),'');
      TOBMEre.PutValue ('TAX_VALEURTAXE'+inttoStr(Num),'');
    end;
  end;
  For Num := 0 To LaTOBInterm.detail.count -1 do
  begin
  	LATOBOrig := LaTOBInterm.detail[Num];
    prefixe := LATOBOrig.getValue('GPB_CATEGORIETAXE');
    Code := LATOBOrig.getValue('GPB_FAMILLETAXE');
    //
//    if Code=VH_GC.AutoLiquiTVAST then
    if TOBPiece.getBoolean('GP_AUTOLIQUID') then
    begin
  		TOBMEre.PutValue ('TVA_AUTOLIQUIDST','OUI');
    end;
    // -----
    LaTOBrepart := fTOBrepart.findfirst(['BPM_CATEGORIETAXE','BPM_FAMILLETAXE'],[prefixe,code],true);
    if LastPrefixe <> Prefixe Then
    begin
    	IndiceTaxe := 1;
      LastPrefixe := prefixe;
    end else
    begin
    	inc (IndiceTaxe);
    end;
    Suffixe := IntToStr (Indicetaxe);
    //FV1 - 26/09/2017
    FieldList := Mcd.getTable(PrefixeToTable('GPB')).Fields;
    FieldList.Reset();
    //
		While FieldList.MoveNext do
    begin
      ChampsOrig := (FieldList.Current as IFieldCOM).name;
      ChampsDest := Prefixe+Copy(ChampsOrig,4,255)+Suffixe;
      Chpt := 'TAX'+Copy(ChampsOrig,4,255)+InttoStr(Cpt);
      if not TOBMere.FieldExists (ChampsDest) then
      begin
        TOBMere.AddChampSupValeur (ChampsDest,LATOBOrig.GetValue(ChampsOrig));
      end else
      begin
        TOBMere.PutValue (ChampsDest,LATOBOrig.GetValue(ChampsOrig));
      end;
      if not TOBMere.FieldExists (ChPt) then
      begin
        TOBMere.AddChampSupValeur (ChPt,LATOBOrig.GetValue(ChampsOrig));
      end else
      begin
        TOBMere.PutValue (ChPt,LATOBOrig.GetValue(ChampsOrig));
    end;
    end;
    //
    if LaTOBrepart <> nil then
    begin
  		Tobmere.PutValue('TAX_REPART'+inttoStr(Cpt),LaTOBrepart.getValue('BPM_MILLIEME'));
  		Tobmere.PutValue('PRESENCEMILLIEME','X');
    end;
    //
    if not TOBMere.FieldExists('TAX_LCODETVA'+inttostr(Cpt)) then
    begin
      TOBMere.AddChampSup  ('TAX_LCODETVA'+inttostr(Cpt),false);
    end;
    TOBMere.PutValue ('TAX_LCODETVA'+inttostr(Cpt),FindLibre('TX1',Code));
    if TOBMere.FieldExists(Prefixe+'_LCODETVA'+Suffixe) then
    begin
      TOBMere.PutValue (Prefixe+'_LCODETVA'+Suffixe,FindLibre('TX1',Code));
    end;
    inc(Cpt);
  end;
  FreeAndNil(LaTOBInterm);
  TOBBRG.Free;
end;

procedure TImprPieceViaTOB.AjouteChampsPiedBaseRG (TOBmere,TOBAInsere : TOB);
var Num,IndiceTaxe : integer;
		Suffixe,Prefixe,LastPrefixe : string;
    LaTOBOrig : TOB;
begin
  IndiceTaxe := 0;
  Prefixe := '';
  LastPrefixe := '';
  For Num := 0 To TOBAInsere.detail.count -1 do
  begin
  	LATOBOrig := TOBAInsere.detail[Num];
    prefixe := LATOBOrig.getValue('PBR_CATEGORIETAXE');
    if LastPrefixe <> Prefixe Then
    begin
    	IndiceTaxe := 1;
      LastPrefixe := prefixe;
    end else
    begin
    	inc (IndiceTaxe);
    end;
    Suffixe := IntToStr (Indicetaxe);
    if TOBMere.FieldExists (Prefixe+'_BASETAXE'+suffixe) and (TOBMere.getValue('GPP_APPLICRG')='X') then
    begin
    	TOBMere.putValue(Prefixe+'_VALEURTAXE'+suffixe,TOBMere.GetValue(Prefixe+'_VALEURTAXE'+suffixe) - LaTOBOrig.GetValue('PBR_VALEURTAXE'));
    	TOBMere.putValue(Prefixe+'_VALEURDEV'+suffixe,TOBMere.GetValue(Prefixe+'_VALEURDEV'+suffixe) - LaTOBOrig.GetValue('PBR_VALEURDEV'));
    end;
    if not TOBMere.FieldExists(Prefixe+'_LCODETVA'+Suffixe) then
    begin
      TOBMere.AddChampSupValeur (Prefixe+'_LCODETVA'+Suffixe,
                                 rechdom('GCFAMILLETAXE'+Suffixe,LATOBOrig.GetValue('PBR_FAMILLETAXE'),True));
    end;
  end;
end;

procedure TImprPieceViaTOB.AjouteChampsPiedPort (TOBmere,TOBAInsere : TOB);
var ITable,Indice,Num : integer;
		ChampsOrig,ChampsDest,Suffixe : string;
    LaTOBOrig : TOB;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  Table := Mcd.GetTable(Mcd.PrefixeToTable('GPT'));
  //FV1 - 26/09/2017
  //FieldList := Table.Fields;
	//FieldList.Reset();
  //
	TOBMere.AddChampSupValeur ('GPT_LIBELLE1','');
	TOBMere.AddChampSupValeur ('GPT_LIBELLE2','');
	TOBMere.AddChampSupValeur ('GPT_LIBELLE3','');
	TOBMere.AddChampSupValeur ('GPT_LIBELLE4','');
	TOBMere.AddChampSupValeur ('GPT_LIBELLE5','');
  For Num := 0 To TOBAInsere.detail.count -1 do
  begin
  	LATOBOrig := TOBAInsere.detail[Num];
    //
    if (Pos(LATOBOrig.GetString('GPT_TYPEPORT'),'PT;MIC;MTC;')>0) then continue;
    //
    Suffixe := IntToStr (LATOBOrig.getValue('GPT_NUMPORT'));
    //FV1 - 26/09/2017
    FieldList := Table.Fields;
    FieldList.Reset();
    //
		While FieldList.MoveNext do
    begin
      ChampsOrig := (FieldList.Current as IFieldCOM).name;
      ChampsDest := ChampsOrig+Suffixe;
      if not TOBMere.FieldExists (ChampsDest) then
      begin
        TOBMere.AddChampSupValeur (ChampsDest,LATOBOrig.GetValue(ChampsOrig));
      end else
      begin
        TOBMere.PutValue (ChampsDest,LATOBOrig.GetValue(ChampsOrig));
      end;
    end;
  end;
end;

procedure TImprPieceViaTOB.AjouteChampsParamDoc (TOBMere,TOBAInsere : TOB);
var ITable,Indice,Num : integer;
		ChampsOrig,ChampsDest : string;
    LaTOBOrig : TOB;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  table := Mcd.getTable(Mcd.PrefixeToTable('BPD'));
  //FV1 - 26/09/2017
  //FieldList := Table.Fields;
  //FieldList.Reset();
  // Compatibilité avec l'existant n'ayant pas de paramétrage doc
  TOBMEre.AddChampSupValeur ('BPD_TYPESSD',         0);
  TOBMEre.AddChampSupValeur ('BPD_TYPEPRES', DOU_TOUS);
  TOBMEre.AddChampSupValeur ('BPD_IMPTOTPAR',     'X');
  TOBMEre.AddChampSupValeur ('BPD_IMPTOTSPP',     'X');
  TOBMEre.AddChampSupValeur ('BPD_IMPCOLONNES',   '-');
  TOBMEre.AddChampSupValeur ('BPD_TYPBLOCNOTE',   'D');
  TOBMEre.AddChampSupValeur ('BPD_SAUTAPRTXTDEB', '-');
  TOBMEre.AddChampSupValeur ('BPD_SAUTAVTTXTFIN', '-');
  TOBMEre.AddChampSupValeur ('BPD_IMPRRECPAR',    '-');
  TOBMEre.AddChampSupValeur ('BPD_IMPDESCRIPTIF', 'I');
  TOBMEre.AddChampSupValeur ('BPD_IMPMETRE',      'S');
  TOBMEre.AddChampSupValeur ('BPD_IMPTATOTSIT',   '-');
  TOBMEre.AddChampSupValeur ('BPD_DESCREMPLACE',  '-');
  TOBMEre.AddChampSupValeur ('BPD_IMPTABTOTSIT',  '-'); // Tableau total des situations
  TOBMEre.AddChampSupValeur ('BPD_IMPRECSIT',     '-'); // recap des situations
  TOBMEre.AddChampSupValeur ('BPD_IMPDEVAVE',     '-'); // recap des situations
  TOBMEre.AddChampSupValeur ('BPD_IMPCUMULAVE',   '-'); // recap des situations

  // --
  For Num := 0 To TOBAInsere.detail.count -1 do
  begin
  	LATOBOrig := TOBAInsere.detail[Num];
    //FV1 - 26/09/2017
    FieldList := Table.Fields;
    FieldList.Reset();
    //
		While FieldList.MoveNext do
    begin
      ChampsOrig := (FieldList.Current as IFieldCOM).name;
      ChampsDest := ChampsOrig;
      if not TOBMere.FieldExists (ChampsDest) then
      begin
        TOBMere.AddChampSupValeur (ChampsDest,LATOBOrig.GetValue(ChampsOrig));
      end else
      begin
        TOBMere.PutValue (ChampsDest,LATOBOrig.GetValue(ChampsOrig));
      end;
    end;
  end;
  //

  //
end;

procedure TImprPieceViaTOB.CreeAdresseFacturation (TobMere : TOB);
begin
  if not TOBMere.FieldExists('FAC_ADRESSE1') then
  	TOBMere.AddChampSupValeur('FAC_ADRESSE1',TOBMere.GetValue('LIV_ADRESSE1'))
  else
  	TOBMere.PutValue('FAC_ADRESSE1',TOBMere.GetValue('LIV_ADRESSE1'));

  if not TOBMere.FieldExists('FAC_ADRESSE2') then
  	TOBMere.AddChampSupValeur('FAC_ADRESSE2',TOBMere.GetValue('LIV_ADRESSE2'))
  else
  	TOBMere.PutValue('FAC_ADRESSE2',TOBMere.GetValue('LIV_ADRESSE2'));

  if not TOBMere.FieldExists('FAC_ADRESSE3') then
  	TOBMere.AddChampSupValeur('FAC_ADRESSE3',TOBMere.GetValue('LIV_ADRESSE3'))
  else
  	TOBMere.PutValue('FAC_ADRESSE3',TOBMere.GetValue('LIV_ADRESSE3'));

  if not TOBMere.FieldExists('FAC_CODEPOSTAL') then
  	TOBMere.AddChampSupValeur('FAC_CODEPOSTAL',TOBMere.GetValue('LIV_CODEPOSTAL'))
  else
  	TOBMere.PutValue('FAC_CODEPOSTAL',TOBMere.GetValue('LIV_CODEPOSTAL'));

  if not TOBMere.FieldExists('FAC_VILLE') then
  	TOBMere.AddChampSupValeur('FAC_VILLE',TOBMere.GetValue('LIV_VILLE'))
  else
  	TOBMere.PutValue('FAC_VILLE',TOBMere.GetValue('LIV_VILLE'));

  if not TOBMere.FieldExists('FAC_LIBELLE') then
  	TOBMere.AddChampSupValeur('FAC_LIBELLE',TOBMere.GetValue('LIV_LIBELLE'))
  else
  	TOBMere.PutValue('FAC_LIBELLE',TOBMere.GetValue('LIV_LIBELLE'));

  if not TOBMere.FieldExists('FAC_LIBELLE2') then
  	TOBMere.AddChampSupValeur('FAC_LIBELLE2',TOBMere.GetValue('LIV_LIBELLE2'))
  else
  	TOBMere.PutValue('FAC_LIBELLE2',TOBMere.GetValue('LIV_LIBELLE2'));
end;

procedure TImprPieceViaTOB.AjouteChampsPieceAdresse (TOBmere,TOBAInsere : TOB);
var ITable,Indice,Num : integer;
		ChampsOrig,ChampsDest,Prefixe : string;
    LaTOBOrig : TOB;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  Table := Mcd.GetTable(Mcd.PrefixeToTable('GPA'));

  TOBMere.AddChampSupValeur('FAC_JURIDIQUE','');
  TOBMere.AddChampSupValeur('FAC_NIF','');
  TOBMere.AddChampSupValeur('FAC_ADRESSE1','');
  TOBMere.AddChampSupValeur('FAC_ADRESSE2','');
  TOBMere.AddChampSupValeur('FAC_ADRESSE3','');
  TOBMere.AddChampSupValeur('FAC_CODEPOSTAL','');
  TOBMere.AddChampSupValeur('FAC_VILLE','');
  TOBMere.AddChampSupValeur('FAC_LIBELLE','');
  TOBMere.AddChampSupValeur('FAC_LIBELLE2','');
  TOBMere.AddChampSupValeur('FAC_PAYS','');
  TOBMere.AddChampSupValeur('FAC_INCOTERM','');
  TOBMere.AddChampSupValeur('FAC_EAN','');
  TOBMere.AddChampSupValeur('FAC_LIEUDISPO','');
  TOBMere.AddChampSupValeur('FAC_REGION','');
  TOBMere.AddChampSupValeur('FAC_NUMEROCONTACT',0);
  //
  TOBMere.AddChampSupValeur('LIV_JURIDIQUE','');
  TOBMere.AddChampSupValeur('LIV_NIF','');
  TOBMere.AddChampSupValeur('LIV_ADRESSE1','');
  TOBMere.AddChampSupValeur('LIV_ADRESSE2','');
  TOBMere.AddChampSupValeur('LIV_ADRESSE3','');
  TOBMere.AddChampSupValeur('LIV_CODEPOSTAL','');
  TOBMere.AddChampSupValeur('LIV_VILLE','');
  TOBMere.AddChampSupValeur('LIV_LIBELLE','');
  TOBMere.AddChampSupValeur('LIV_LIBELLE2','');
  TOBMere.AddChampSupValeur('LIV_PAYS','');
  TOBMere.AddChampSupValeur('LIV_EAN','');
  TOBMere.AddChampSupValeur('LIV_LIEUDISPO','');
  TOBMere.AddChampSupValeur('LIV_REGION','');
  TOBMere.AddChampSupValeur('LIV_NUMEROCONTACT',0);

  For Num := 0 To TOBAInsere.detail.count -1 do
  begin

  	LATOBOrig := TOBAInsere.detail[Num];

    if LATOBOrig.getValue('GPA_TYPEPIECEADR')='001' Then
    begin
    	Prefixe := 'LIV'
    end else // MODIFBRL 22/07/04 if LATOBOrig.getValue('GPA_TYPEPIECEADR')='002' Then
    begin
    	prefixe := 'FAC';
    end;

    FieldList := Table.Fields;
    FieldList.Reset();
		While FieldList.MoveNext do
    begin
      ChampsOrig := (FieldList.Current as IFieldCOM).name;
      ChampsDest := Prefixe + Copy(ChampsOrig,4,255);
      if not TOBMere.FieldExists (ChampsDest) then
      begin
        TOBMere.AddChampSupValeur (ChampsDest,LATOBOrig.GetValue(ChampsOrig));
      end else
      begin
        TOBMere.PutValue (ChampsDest,LATOBOrig.GetValue(ChampsOrig));
      end;
    end;
  end;

  // Cas ou une adresse manque
  if TOBAInsere.detail.count = 1 then
  begin
  	LATOBOrig := TOBAInsere.detail[0];
    if LATOBOrig.getValue('GPA_TYPEPIECEADR')='001' Then
    begin
    	Prefixe := 'FAC' // Si l'adresse de livraison existe on ajoute celle de facturation
    end else if LATOBOrig.getValue('GPA_TYPEPIECEADR')='002' Then
    begin
    	prefixe := 'LIV'; // Si l'adresse de facturation existe on ajoute celle de livraison
    end;
    //FV1 - 26/09/2017
    //FieldList := Table.Fields;
    //FieldList.Reset();
		While FieldList.MoveNext do
    begin
      ChampsOrig := (FieldList.Current as IFieldCOM).name;
      ChampsDest := Prefixe + Copy(ChampsOrig,4,255);
      if not TOBMere.FieldExists (ChampsDest) then
      begin
        TOBMere.AddChampSupValeur (ChampsDest,LATOBOrig.GetValue(ChampsOrig));
      end else
      begin
        TOBMere.PutValue (ChampsDest,LATOBOrig.GetValue(ChampsOrig));
      end;
    end;
  end;

  If (TOBMere.getValue('FAC_ADRESSE1')='') AND
  	 (TOBMere.getValue('FAC_ADRESSE2')='') AND
  	 (TOBMere.getValue('FAC_ADRESSE3')='') AND
  	 (TOBMere.getValue('FAC_CODEPOSTAL')='') AND
  	 (TOBMere.getValue('FAC_VILLE')='') AND
  	 (TOBMere.getValue('FAC_LIBELLE')='') AND
  	 (TOBMere.getValue('FAC_LIBELLE2')='') then CreeAdresseFacturation (TobMere);

end;

procedure TImprPieceViaTOB.AjouteChampsPieceRG (TOBmere,TOBAInsere : TOB);
var Num : integer;
    LaTOBOrig : TOB;
    NumCaution : string;
    Xp,XD : double;
begin
  NumCaution := '11111';
  TOBMere.AddChampSupValeur ('PRG_TYPERG','HT');
  TOBMere.AddChampSupValeur ('PRG_TAUXRG',0);
  TOBMere.AddChampSupValeur ('PRG_MTHTRG',0);
  TOBMere.AddChampSupValeur ('PRG_MTHTRGDEV',0);
  TOBMere.AddChampSupValeur ('PRG_MTTTCRG',0);
  TOBMere.AddChampSupValeur ('PRG_MTTTCRGDEV',0);
  TOBMere.AddChampSupValeur ('PRG_CAUTIONMT',0);
  TOBMere.AddChampSupValeur ('PRG_CAUTIONMTDEV',0);
  TOBMere.AddChampSupValeur ('PRG_NUMCAUTION','');
	if TOBAInsere.detail.count > 0 then
  begin
    TOBMere.PutValue ('PRG_TYPERG',TOBAInsere.detail[0].GetValue('PRG_TYPERG'));
    TOBMere.PutValue ('PRG_TAUXRG',TOBAInsere.detail[0].GetValue('PRG_TAUXRG'));
  end;
  //
  getRG (TOBAInsere,True,(Pos(fTOBPiece.getValue('GP_NATUREPIECEG'),'FBT;DAC;FBP;BAC;')>0),Xp,XD,numcaution);
  TOBMere.PutValue ('PRG_MTHTRG',XP);
  TOBMere.PutValue ('PRG_MTHTRGDEV',XD);
  //
  getRG (TOBAInsere,False,(Pos(fTOBPiece.getValue('GP_NATUREPIECEG'),'FBT;DAC;FBP;BAC;')>0),Xp,XD,numcaution);
  TOBMere.PutValue ('PRG_NUMCAUTION',NumCaution);
  TOBMere.PutValue ('PRG_MTTTCRG',XP);
  TOBMere.PutValue ('PRG_MTTTCRGDEV',XD);
  if NumCaution <> '' then
  begin
    // pour etre sur du fonctionnement comme a l'ancienne
    TOBMere.PutValue ('PRG_CAUTIONMT',TOBMere.GetValue('PRG_MTTTCRG')+1);
    TOBMere.PutValue ('PRG_CAUTIONMTDEV',TOBMere.GetValue('PRG_MTTTCRGDEV')+1);
  end;
end;

procedure TImprPieceViaTOB.AjouteChampsDOCTEXTEDEBFIN (TOBmere,TOBAInsere : TOB);
var Num : integer;
		ChampsOrig,ChampsDest,Suffixe : string;
    LaTOBOrig : TOB;
begin
  TOBMere.AddChampSupValeur ('LO_OBJET_ENTETEGEN',#0);
  TOBMere.AddChampSupValeur ('LO_OBJET_PIEDGEN',#0);
    For Num := 0 To TOBAInsere.detail.count -1 do
    begin
      LATOBOrig := TOBAInsere.detail[Num];

      if LATOBOrig.getValue('LO_RANGBLOB')=1 Then Suffixe := '_ENTETEGEN'
                                             else Suffixe := '_PIEDGEN';

      ChampsOrig := 'LO_OBJET';
      ChampsDest := ChampsOrig+Suffixe;

      if not TOBMere.FieldExists (ChampsDest) then
      begin
        TOBMere.AddChampSupValeur (ChampsDest,LATOBOrig.GetValue(ChampsOrig));
      end else
      begin
        TOBMere.PutValue (ChampsDest,LATOBOrig.GetValue(ChampsOrig));
    end;
  end;
end;

procedure TImprPieceViaTOB.AjouteChampsINTEXTEDEBFIN (TOBmere,TOBAInsere : TOB);
var Num : integer;
		ChampsOrig,ChampsDest,Suffixe : string;
    LaTOBOrig : TOB;
begin
  TOBMere.AddChampSupValeur ('LO_OBJET_ENTETEDOC',#0);
  TOBMere.AddChampSupValeur ('LO_OBJET_PIEDDOC',#0);
  For Num := 0 To TOBAInsere.detail.count -1 do
  begin
  	LATOBOrig := TOBAInsere.detail[Num];

    if LATOBOrig.getValue('LO_RANGBLOB')=1 Then Suffixe := '_ENTETEDOC'
                                           else Suffixe := '_PIEDDOC';
    ChampsOrig := 'LO_OBJET';
    ChampsDest := ChampsOrig+Suffixe;

    if not TOBMere.FieldExists (ChampsDest) then
    begin
      TOBMere.AddChampSupValeur (ChampsDest,LATOBOrig.GetValue(ChampsOrig));
    end else
    begin
      TOBMere.PutValue (ChampsDest,LATOBOrig.GetValue(ChampsOrig));
    end;
  end;
end;

procedure TImprPieceViaTOB.AjouteChampsOUVRAGE (TOBmere,TOBAINSERE ,TOBLigne : TOB);
var QteDuDetail,QteDuPV,Montant : Double;
		EnHt : boolean;
    DEv : RDevise;
		ITable,Indice: integer;
    NewChamp : string;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;
    NomChamps : string;
begin
  if TOBAinsere = nil then exit;
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  Table := Mcd.GetTable(Mcd.PrefixeToTable('BLO'));
  FieldList := Table.Fields;
  FieldList.Reset();

	FillChar(DEV, Sizeof(DEV), #0);
  DEV.Code := TOBLigne.GetValue('GL_DEVISE');
  GetInfosDevise(DEV);
	EnHt := (TOBLigne.GetValue('GL_FACTUREHT')='X');

  While FieldList.MoveNext do
  begin
		NomChamps := (FieldList.Current as IFieldCOM).name;
  	if NomChamps = 'BLO_NUMORDRE' Then continue;
  	// ca va pas non...
    // ben en fait si ca va bien ..
  	// if V_PGI.Dechamps[iTable, Indice].Nom = 'BLO_BLOCNOTE' then continue;
    // -----------------
    // --
  	NewChamp := 'GL'+Copy(NomChamps,4,255);
    if TOBMere.FieldExists (newChamp) then
    begin
    	TOBMere.PutValue (NewChamp,TOBAInsere.GetValue(NomChamps));
    end;
  end;
  TOBMere.PutValue ('GL_TYPEPRESENT',TobLigne.getValue('GL_TYPEPRESENT'));
  //
  QteDudetail := TOBAINSERE.GetValue('BLO_QTEDUDETAIL');
  if QteDUdetail = 0 then qteduDetail := 1;
  QteDuPv := TOBAINSERE.GetValue('BLO_PRIXPOURQTE');
  if QteDUpv = 0 then qtedupv := 1;

  if (IsVariante(TOBAInsere)) or (IsVariante(TOBLigne)) then
  begin
  	TOBMEre.PutValue ('ISDETOUVVAR','X');
  end else
  begin
  	TOBMEre.PutValue ('ISDETOUVSTD','X');
  end;
//
  if TOBAINSERE.GetValue('BLO_TYPEARTICLE') <> 'POU' then
  begin
    if Enht then
      Montant := arrondi(TOBAINSERE.GetValue('BLO_PUHTDEV') * (TOBLigne.getValue('GL_QTEFACT')) * (TOBAINSERE.GetValue('BLO_QTEFACT') / (QteDuPv * QteDuDetail)),
        DEV.Decimale)
    else
      Montant := arrondi(TOBAINSERE.GetValue('BLO_PUTTCDEV') * (TOBLigne.getValue('GL_QTEFACT')) * (TOBAINSERE.GetValue('BLO_QTEFACT') / (QteDuPv * QteDuDetail)),
        DEV.Decimale);
  end;
//
  TOBmere.PutValue('GL_REFARTSAISIE', TOBAINSERE.GetValue('BLO_CODEARTICLE'));
  TOBmere.PutValue('GL_REFARTTIERS', TOBAINSERE.GetValue('BLO_ARTICLE'));
  TOBmere.PutValue('GL_ARTICLE', TOBAINSERE.GetValue('BLO_ARTICLE'));
  TOBmere.PutValue('GL_CODEARTICLE', TOBAINSERE.GetValue('BLO_CODEARTICLE'));
	TOBmere.PutValue('GL_LIBELLE', TOBAINSERE.GetValue('BLO_LIBELLE'));
	TOBmere.PutValue('GL_BLOCNOTE', TOBAINSERE.GetValue('BLO_BLOCNOTE'));
  if TOBAINSERE.GetValue('BLO_TYPEARTICLE') = 'POU' then
  	TOBmere.PutValue('GL_QTEFACT', TOBAINSERE.GetValue('BLO_QTEFACT') / QteDuDetail)
  else TOBmere.PutValue('GL_QTEFACT', (TOBAINSERE.GetValue('BLO_QTEFACT') * TOBLigne.GetValue('GL_QTEFACT')) / QteDuDetail);
  TOBmere.PutValue('GL_QUALIFQTEVTE', TOBAINSERE.GetValue('BLO_QUALIFQTEVTE'));
	if TOBAINSERE.GetValue('BLO_TYPEARTICLE') <> 'POU' then
  begin
  	if EnHt then TOBmere.PutValue('GL_PUHTDEV', TOBAINSERE.GetValue('BLO_PUHTDEV') / QteDupv)
    				else TOBmere.PutValue('GL_PUTTCDEV', TOBAINSERE.GetValue('BLO_PUTTCDEV') / QteDupv);
  end;
  if TOBAINSERE.GetValue('BLO_TYPEARTICLE') <> 'POU' then
  begin
    if EnHt then TOBmere.PUtValue('GL_MONTANTHTDEV', Montant)
    				else TOBmere.PUtValue('GL_MONTANTTTCDEV', Montant);
  end;
  if fAvancement then
  begin
    Table := Mcd.GetTable(Mcd.PrefixeToTable('BLF'));
    FieldList := Table.Fields;
    FieldList.Reset();

		While FieldList.MoveNext do
    begin
			NomChamps := (FieldList.Current as IFieldCOM).name;
      NewChamp := 'BLF'+Copy(NomChamps,4,255);
      if TOBMere.FieldExists (newChamp) then
      begin
        TOBMere.PutValue (NewChamp,TOBAInsere.GetValue(NomChamps));
      end;
    end;
  end;
	DefiniAffichageLigne (TOBMere);
end;

procedure TImprPieceViaTOB.AjouteTOB(TOBMere: TOB; CodeTraitement: string; TOBAInsere: TOB; TOBPrec : TOB);
begin
  if CodeTraitement = 'RETENUESTTC' then
  begin
  	AjouteRetenuesTTC (TOBMere,TOBAInsere);
  end else if CodeTraitement = 'RETENUESDIV' then
  begin
  	AjouteRetenuesDiv (TOBMere,TOBAInsere);
  end else if CodeTraitement = 'ACOMPTESCPTA' then
  begin
  	AjouteChampsAcomptesCpta (TOBMere);
  end else if CodeTraitement = 'AVANCEMENTG' then
  begin
  	AjouteChampsAvancementPiece (TOBMere);
  end else if CodeTraitement = 'TABLEAUSITUATIONS' then
  begin
  	AjouteChampsSituation (TOBMere);
  end else if CodeTraitement = 'AVANCLIGNE' then
  begin
  	CumuleChampsAvancLigne (TOBMere,TOBAInsere);
  end else if CodeTraitement = 'PIECE' then
  begin
  	AjouteChampsPiece (TOBMere,TOBAInsere);
  end else if CodeTraitement = 'BTPIECEMILIEME' then
  begin
  	AjouteChampsMillieme (TOBMere,fTOBRepart);
  end else if CodeTraitement = 'PARAMDOC' then
  begin
  	AjouteChampsParamDoc (TOBMere,TOBAInsere);
  end else if CodeTraitement = 'PARPIECE' then
  begin
  	AjouteChampsParPiece (TOBMere,TOBAInsere);
  end else if CodeTraitement = 'DOCTEXTEDEBFIN' then
  begin
  	AjouteChampsDOCTEXTEDEBFIN (TOBmere,TOBAInsere);
  end else if CodeTraitement = 'INTEXTEDEBFIN' then
  begin
  	AjouteChampsINTEXTEDEBFIN (TOBmere,TOBAInsere);
  end else if CodeTraitement = 'TIERS' then
  begin
  	AjouteChampsTiers (TOBmere,TOBAInsere);
  end else if CodeTraitement = 'AFFAIRE' then
  begin
  	AjouteChampsAffaire (TOBmere,TOBAInsere);
  end else if CodeTraitement = 'AFFAIREDEV' then
  begin
  	AjouteChampsAffaireDEVIS (TOBmere,TOBAInsere);
  end else if CodeTraitement = 'PIEDBASE' then
  begin
  	AjouteChampsPiedBase (TOBmere,TOBAInsere);
  end  else if CodeTraitement = 'SSTRAITANCE' then
  begin
    AjouteChampsSsTraitant(TOBmere,TOBAInsere);
  end  else if CodeTraitement = 'SSTRAITANCEINT' then
  begin
    AjouteChampsSsTraitantInterne(TOBmere,TOBAInsere);
  end else if CodeTraitement = 'PIEDBASERG' then
  begin
  	AjouteChampsPiedBaseRG (TOBmere,TOBAInsere);
  end else if CodeTraitement = 'PIEDECHE' then
  begin
  	AjouteChampsPiedEche (TOBmere,TOBAInsere);
  end else if CodeTraitement = 'PIEDPORT' then
  begin
  	AjouteChampsPiedPort (TOBmere,TOBAInsere);
  end else if CodeTraitement = 'PIECEADRESSE' then
  begin
  	AjouteChampsPieceAdresse (TOBmere,TOBAInsere);
  end else if CodeTraitement = 'PIECERG' then
  begin
  	AjouteChampsPieceRG (TOBmere,TOBAInsere);
  end else if CodeTraitement = 'DEVISES' then
  begin
  	AjouteChampsDEVISE (TOBmere);
  end else if CodeTraitement = 'LIGNES' then
  begin
  	AjouteChampsLIGNE (TOBmere,TOBAINSERE);
  end else if CodeTraitement = 'LIGNEFAC' then
  begin
  	AjouteChampsLIGNE (TOBmere,TOBAINSERE,True);
  end else if CodeTraitement = 'ARTICLES' then
  begin
  	AjouteChampsARTICLES (TOBmere,TOBAINSERE);
  end else if CodeTraitement = 'OUVRAGE' then
  begin
  	AjouteChampsOUVRAGE (TOBmere,TOBAINSERE,TOBPrec);
  end else if CodeTraitement = 'METRE' then
  begin
  	AjouteChampsMETRE (TOBmere);
  end else if CodeTraitement = 'PARAG' then
  begin
  	AjouteChampsPARAG (TOBmere,TOBAInsere);
  end else if CodeTraitement = 'LESPORTS' then
  begin
  	AjoutelesPORTS (TOBmere,TOBAInsere);
  end else if CodeTraitement = 'TIMBRES' then
  begin
  	AjouteleMtTimbre (TOBmere,TOBAInsere);
  end else if CodeTraitement = 'METRESLIG' then
  begin
  	AjoutelesMETRESLIG (TOBmere,TOBAInsere);
  end else if CodeTraitement = 'RETENUEDIVERSE' then
  begin
    AjouteRetenuesDiverses(TOBMere, TobAInsere);
  end;

end;

procedure TImprPieceViaTOB.associeTOBs(TOBPiece, TOBBases, TOBEches,
                                      TOBPorcs, TOBTiers, TOBArticles, TOBComms, TOBAdresses, TOBNomenClature,
                                      TOBDesLots, TOBSerie, TOBAcomptes, TOBAffaire, TOBOuvrage, TOBLienOle,
                                      TOBPieceRg, TOBBasesRG, TOBFormuleVar,TOBMEtre,TOBRepart,TOBMetres,TOBTimbres,TOBPieceTrait : TOB);
var QQ : TQuery;
begin
  fTOBpieceTrait := TOBpieceTrait;

	fPreTobPiece := TOBPiece;
  fTOBBases := TOBBases;
  fTOBEches := TOBEches;
  fTOBPorcs := TOBPorcs;
  fTOBTiers := TOBTiers;
  fTOBMetres := TOBmetres;
  fTOBArticles := TOBArticles;
  fTOBComms := TOBComms;
  fTOBAdresses := TOBAdresses;
  fTOBNomenclature := TOBNomenclature;
  fTOBLOTS := TOBDesLots;
  fTOBSerie := TOBSerie;
  fTOBAcomptes := TOBAcomptes;
  fTOBAffaire := TOBAffaire;
  if TOBPiece.getValue('GP_AFFAIREDEVIS') <> '' then
  begin
    fTOBAffaireDev := findCetteAffaire(TOBPiece.getValue('GP_AFFAIREDEVIS'));
    if fTOBAffaireDev = nil then StockeCetteAffaire (TOBPiece.getValue('GP_AFFAIREDEVIS'))
  end else
  begin
		fTOBAffaireDev := nil;
  end;
  fTOBOuvrage := TOBOuvrage;
  fTOBLIENOLE := TOBLienOle;
  fTOBPieceRG := TOBPieceRG;
  fTOBBasesRG := TOBBasesRG;
  fTOBFormuleVar := TOBFormuleVar;
  fTOBMetre := TOBmetre ;
  fTOBRepart := TOBRepart;
  fTOBTimbres  := TOBTimbres;

  FillChar(fCleDoc, Sizeof(fCleDoc), #0);
  fCleDoc.NaturePiece := TOBPiece.getValue('GP_NATUREPIECEG');
  fCleDoc.DatePiece := TOBPiece.getValue('GP_DATEPIECE');
  fCleDoc.Souche := TOBPiece.getValue('GP_SOUCHE');
  fCleDoc.NumeroPiece := TOBPiece.getValue('GP_NUMERO');
  fCleDoc.Indice := TOBPiece.getValue('GP_INDICEG');
  if fcledoc.Naturepiece <> '' then
  begin
  	QQ := OpenSql ('SELECT * FROM PARPIECE WHERE GPP_NATUREPIECEG="'+fCledoc.NaturePiece+'"',true,-1,'',true);
    ftobParPiece.SelectDB ('',QQ);
    ferme (QQ);
  end;
end;

constructor TImprPieceViaTOB.create (TheFiche : TForm);
begin
  TheRepartTva := TREPARTTVAMILL.create (TheFiche) ;
  fTOBPieceImp := TOB.create ('LA PIECE A IMPRIMER',nil,-1);
  fTOBPieceRecap := TOB.Create ('LE RECAP',nil,-1);
  fTobSsTraitantImp := TOB.Create ('LE RECAP DES SS TRAIT',nil,-1);
  fTOBParPiece := TOB.Create ('PARPIECE',nil,-1);
  fTOBCodeTaxes := TOB.create ('LES TAXES',nil,-1);
  fTOBSituations := TOB.Create ('LES SITUATIONS',nil,-1);
  //
  fInternalWindow := TToolWindow97.create(fform);
  fInternalWindow.Parent := fform;
  fTexte := THRichEditOLE.Create (fInternalWindow);
  ftexte.Parent := fInternalWindow;
  ftexte.text := '';
  ftexte.Visible := false;
  fInternal := false;
  WithRecap := false;
//  TheMetre:= TStockMetre.create; ;
end;

destructor TImprPieceViaTOB.Destroy;
begin
	if fInternal then
  begin
  	DetruitTout;
  end;
  if TheRepartTva <> nil then FreeAndNil(TheRepartTva);
  if fTOBParPiece <> nil then FreeAndNil(fTOBParPiece);
  if fTOBPieceImp <> nil then FreeAndNil(fTOBPieceImp);
  if fTOBPieceRecap <> nil then FreeAndNil(fTOBPieceRecap);
  if fTOBCodeTaxes <> nil then FreeAndNil(fTOBCodeTaxes);
  if fTOBSituations <> nil then FreeAndNil(fTOBSituations);
  if ftexte <> nil then FreeAndNil(ftexte);
  if fInternalWindow <> nil then FreeAndNil(fInternalWindow);
  if fTobSsTraitantImp <> nil then FreeAndNil(fTobSsTraitantImp);
//  TheMEtre.free;
  inherited;
end;

function TImprPieceViaTOB.InsereTOB (NomTable : string;TOBmere : TOB; WithNewNumOrdre : boolean=true) : TOB;
begin
	result := TOB.Create(NomTable,TOBMere,-1);
  if (NomTable = 'LIGNE') and (WithNewNumOrdre) Then
  begin
  	inc(flastNumordre);
  	AddLesSupLigne (result,false);
    AddSupLignesEdt (result);
  	result.AddChampSupValeur('NUMORDRE',flastNumOrdre);
  	result.AddChampSupValeur('SAUTDEPAGE','-');
  end;
end;

function TImprPieceViaTOB.SetQueryForOldSituations : Widestring;
begin
  // Fu le changement de numérotation
  Result := 'SELECT * FROM BSITUATIONS WHERE '+
  					'BST_SSAFFAIRE="'+fTOBPiece.getValue('GP_AFFAIREDEVIS')+'" AND '+
            'BST_VIVANTE="X" AND '+
            'BST_NATUREPIECE<>"B00" AND '+
            'BST_NUMEROSIT <= '+
            	'('+
              	'SELECT BST_NUMEROSIT FROM BSITUATIONS WHERE BST_NATUREPIECE="'+fTOBPiece.getValue('GP_NATUREPIECEG')+'" AND '+
            		'BST_SOUCHE="'+fTOBPiece.getValue('GP_SOUCHE')+'" AND BST_NUMEROFAC='+InttoStr(fTOBPiece.getValue('GP_NUMERO'))+
              ') ORDER BY BST_NUMEROSIT';
end;



function TImprPieceViaTOB.GetDevisFromprepaFac (TOBL : TOB) : string;

  function findDevisFromprepaFac (Cledoc : r_cledoc) : string;
  var QQ : Tquery;
      OneDoc : r_cledoc;
  begin
    QQ:= OpenSql ('SELECT GL_PIECEPRECEDENTE,GL_PIECEORIGINE FROM LIGNE WHERE '+WherePiece(Cledoc,ttdligne,true,true),true,1,'',true);
    if not QQ.eof then
    begin
      DecodeRefPiece (QQ.fields[1].AsString,OneDoc);
      if OneDoc.NaturePiece = 'ETU' then
      begin
        result := QQ.fields[0].AsString;
      end else
      begin
        result := QQ.fields[1].AsString;
      end;
    end;
    ferme (QQ);
  end;
  
var cledoc : r_cledoc;
    OneRefDevis,RefPiece : string;
begin
  DecodeRefPiece (TOBL.getValue('GL_PIECEORIGINE'),cledoc);
  if cledoc.NaturePiece = 'ETU' then
  begin
    DecodeRefPiece (TOBL.getValue('GL_PIECEPRECEDENTE'),cledoc);
    if Cledoc.NaturePiece = 'FBP' then
    begin
      OneRefDevis := findDevisFromprepaFac (cledoc);
      DecodeRefPiece (OneRefDevis,cledoc);
      RefPiece := EncoderefSel (cledoc);
    end else
    begin
      RefPiece := EncoderefSel (cledoc);
    end;
  end else
  begin
    RefPiece := EncoderefSel (cledoc);
  end;
  result := RefPiece;
end;

function TImprPieceViaTOB.isExistsMemo(LaLigne:TOB) : boolean;
var cledoc : r_cledoc;
    RefPiece,requete,OneRefDevis : string;
    QQ : TQuery;
begin
  result := false;
  RefPiece := GetDevisFromprepaFac (LaLigne);
  (*
  DecodeRefPiece (LaLigne.getValue('GL_PIECEORIGINE'),cledoc);
  if cledoc.NaturePiece = 'ETU' then
  begin
    DecodeRefPiece (LaLigne.getValue('GL_PIECEPRECEDENTE'),cledoc);
    if Cledoc.NaturePiece = 'FBP' then
    begin
      OneRefDevis := GetDevisFromprepaFac (cledoc);
      DecodeRefPiece (OneRefDevis,cledoc);
      RefPiece := EncoderefSel (cledoc);
    end else
    begin
      RefPiece := EncoderefSel (cledoc);
    end;
  end else
  begin
    RefPiece := EncoderefSel (cledoc);
  end;
  *)
  //
  if RefPiece = '' then exit;
  //
  Requete := 'SELECT BMF_NATUREPIECEG,BMF_SOUCHE,BMF_NUMERO,BMF_INDICEG FROM BTMEMOFACTURE '+
             'WHERE BMF_DEVISPRINC=(SELECT BMF_DEVISPRINC FROM BTMEMOFACTURE WHERE BMF_DEVIS="'+RefPiece+'")';
  QQ := OpenSql (requete,true,-1,'',true);
  result := not QQ.eof;
  ferme (QQ);
end;

function TImprPieceViaTOB.IsExistsMemoFacturation : boolean;
var Indice : integer;
    first : boolean;
    LaLIgne : TOB;
    TypeLigne : string;
begin
  result := false;
  if not FAvancement then exit;
  first := true;
  for Indice := 0 to fTobPiece.detail.count -1 do
  begin
    LaLigne := fTobPiece.detail[Indice];
    TypeLigne := LaLigne.GetValue('GL_TYPELIGNE');
    if (first) and (TypeLigne = 'ART') and (LaLIgne.GetValue('GL_PIECEORIGINE')<>'') then
    begin
      result := isExistsMemo(LaLigne);
      break;
    end;
  end;
end;

procedure TImprPieceViaTOB.PositionneMarcheFromPiece (TheRecap,ThePiece : TOB);
begin
  TheRecap.PutValue ('NUMSIT',        ThePiece.GetValue ('NUMSIT'));
  TheRecap.PutValue ('MAXNUMSIT',     ThePiece.GetValue ('MAXNUMSIT'));
  TheRecap.PutValue ('DATESIT',       ThePiece.GetValue ('DATESIT'));
  TheRecap.PutValue ('TOTSITPREC',    ThePiece.GetValue ('TOTSITPREC'));
  TheRecap.PutValue ('TOTSITPRECBRUT',ThePiece.GetValue ('TOTSITPRECBRUT'));
  TheRecap.PutValue ('TOTSITPRECTTC', ThePiece.GetValue ('TOTSITPRECTTC'));
  TheRecap.PutValue ('TOTREMPREC',    ThePiece.GetValue ('TOTREMPREC'));
  TheRecap.PutValue ('TOTESCPREC',    ThePiece.GetValue ('TOTESCPREC'));
  TheRecap.PutValue ('TOTPORTPREC',   ThePiece.GetValue ('TOTPORTPREC'));
  TheRecap.PutValue ('TOTPORTSIT',    ThePiece.GetValue ('TOTPORTSIT'));
  TheRecap.PutValue ('TOTMARCHE',     ThePiece.GetValue ('TOTMARCHE'));
  TheRecap.PutValue ('TOTMARCHENET',  ThePiece.GetValue ('TOTMARCHENET'));
  TheRecap.PutValue ('TOTMARCHETTC',  ThePiece.GetValue ('TOTMARCHETTC'));
  TheRecap.PutValue ('TOTCUMULEFACT', ThePiece.GetValue ('TOTCUMULEFACT'));
  TheRecap.PutValue ('TOTSITUATION',  ThePiece.GetValue ('TOTSITUATION'));
  TheRecap.PutValue ('TOTSOLDE',      ThePiece.GetValue ('TOTSOLDE'));
  TheRecap.PutValue ('TOTREGLE',      ThePiece.GetValue ('TOTREGLE'));
end;

function TImprPieceViaTOB.NbSSTraitDir (TOBpieceTrait : TOB) : integer;
var II : integer;
begin
  result := 0;
  for II := 0 to TOBpieceTrait.detail.count -1 do
  begin
    if (TOBpieceTrait.detail[II].GetString('BPE_TYPEINTERV') = 'Y00') and (TOBpieceTrait.detail[II].GetString('TYPEPAIE')='001') then inc(result);
  end;
end;


procedure TImprPieceViaTOB.PreparationEtat;
{$IFDEF BTP}
var Indice,IndiceOuv : integer;
		ThePiece,TheLigne,TheArticle,LaLigne,TheOuvrage,TheParag,LaTotalisation,TheRecap,TheSsTraitant, TheRetenuesDiverses : TOB;
    TOBParamEdt,TOBEntetePiedDoc,TOBOuv,TOBO,TOBPARAG : TOB;
    Req : string;
    QQ : Tquery;
    NbrLignes : integer;
    FromOuvrage,First : boolean;
    TypeLigne : string;
    Niveau : integer;
    MessAttente : TFsplashScreen;
    FromMemoFacturation : boolean;
{$ENDIF}
begin
	FillChar(DEV, Sizeof(DEV), #0);
  DEV.Code := fPreTobPiece.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := fPreTobPiece.GetDouble('GP_TAUXDEV');
  //
  if ((fPreTOBPiece.getString('GP_NATUREPIECEG')='FBT') or (fPreTOBPiece.getString('GP_NATUREPIECEG')='FBP')) and
     (GetParamSocSecur('SO_BTGESTVIDESUREDIT',false)) and
     (fPreTOBPiece.getvalue('AFF_OKSIZERO')='-') then
  begin
    fTobPiece := TOB.Create('PIECE',nil,-1);
    AddLesSupEntete(fTOBPiece);
    fTobPiece.Dupliquer(fPreTOBPiece,true,true);
    VirelesLignesAZero (fTobPiece);
    SuprimeParagraphesOrphelin (fTobPiece);
  end else
  begin
    fTobPiece := fPreTOBPiece;
  end;
  SetMode; // --> facture Avancement ou demande d'acompte ?? A LAISSER EN DEBUT DE FONCTION (BRL 01/03/2013)
  FromMemoFacturation := IsExistsMemoFacturation;
  first := true;
  //
	fTOBPieceImp.ClearDetail;
  fTOBPieceRecap.clearDetail;
  fTobSsTraitantImp.cleardetail;
  //
  NbrLignes := 0;
  fLastNumordre := 0;
  prepareTVA;
  //
  fillchar (fcledocFacCur,sizeof(fcledocfaccur),#0);
  fcledocFacCur.NaturePiece := fTOBpiece.GetValue('GP_NATUREPIECEG');
  fcledocFacCur.Souche := fTOBpiece.GetValue('GP_SOUCHE');
  if InCreation then fcledocFacCur.NumeroPiece := 0
                else fcledocFacCur.NumeroPiece := fTOBpiece.GetValue('GP_NUMERO');
  fcledocFacCur.Indice := fTOBpiece.GetValue('GP_INDICEG');
  //
  TOBParamEdt := TOB.Create ('LES PARAMS EDT',nil,-1);
  TOBEntetePiedDoc := TOB.Create ('LES ENTETES ET PIED',nil,-1);
  TOBParag := TOB.Create ('LES PARAGRAPHES',nil,-1);

  // Récupération du paramétrage d'édition
  Req := 'SELECT * FROM BTPARDOC WHERE BPD_NATUREPIECE="'+fcledocFacCur.NaturePiece+'" AND '+
  			 'BPD_SOUCHE="'+fcledocFacCur.Souche+'" AND '+
         '((BPD_NUMPIECE=0) OR (BPD_NUMPIECE='+InttoStr(fcledocFacCur.NumeroPiece)+'))';
  QQ := OpenSql (Req,True,-1,'',true);
  if QQ.eof then
  begin
    Ferme(QQ);
    Req := 'SELECT * FROM BTPARDOC WHERE BPD_NATUREPIECE="'+fcledocFacCur.NaturePiece+'" AND '+
           'BPD_SOUCHE="'+GetSoucheG(fcledocFacCur.NaturePiece,'','')+'" AND '+
           '((BPD_NUMPIECE=0) OR (BPD_NUMPIECE='+InttoStr(fcledocFacCur.NumeroPiece)+'))';
    QQ := OpenSql (Req,True,-1,'',true);
  end;
  if not QQ.eof Then
  begin
    TOBParamEdt.loadDetailDb ('BTPARDOC','','',QQ,True);
    TOBParamEdt.Putvalue('BPD_SOUCHE', fcledocFacCur.Souche);
  end;

  Ferme (QQ);

  fNbSTTraitDir := NbSSTraitDir(fTOBpieceTrait);

  // récupération des bloc-notes entete et pied
  req := 'SELECT LO_RANGBLOB,LO_OBJET FROM LIENSOLE WHERE LO_TABLEBLOB="GP" AND LO_IDENTIFIANT="'+
  				fcledocFacCur.NaturePiece+':'+fcledocFacCur.Souche+':0:0"';
  QQ := OpenSql (Req,True,-1,'',true);
  if not QQ.eof Then TOBEntetePiedDoc.loadDetailDb ('LIENSOLE','','',QQ,True);
  Ferme (QQ);
  if fAvancement then
  begin
    //FV1 : 22/12/2014 - FS#1121 - BERTHES : l'impression du DGD ne tient pas compte des règlements affectés aux situations antérieures
    CalcReglementSituations (fTOBpiece.getValue('GP_AFFAIRE'));
    //
    Req := SetQueryForOldSituations; // no country for old men ^^
    // récupération des éléments de chiffrage des situations incluant aussi la courante
    QQ := OpenSql (Req,True,-1,'',true);
    if not QQ.eof Then fTOBSituations.loadDetailDb ('BSITUATIONS','','',QQ,false);
    Ferme (QQ);
    SetMontantregleSituations;
  end;
	TRY
		if (GetInfoParPiece(fTOBpiece.GetValue('GP_NATUREPIECEG'), 'GPP_ESTAVOIR') = 'X') then
    begin
      InverseLesPieces(fTOBPiece, 'PIECE');
      InverseLesPieces(fTOBBases, 'PIEDBASE');
      InverseLesPieces(fTOBEches, 'PIEDECHE');
      InverseLesPieces(fTOBPorcs, 'PIEDPORT');
      InverseLesPieces(fTOBPieceRG, 'PIECERG');
      InverseLesPieces(fTOBBasesRG, 'PIEDBASERG');
    end;
		ThePiece := InsereTOB('PIECE',fTOBPieceImp);
    TheRecap :=InsereTOB('PIECE',fTOBPieceRecap);
    TheSsTraitant := InsereTOB('PIECE',fTobSsTraitantImp);
    //
    PreparePiedbase (ThePiece);
    PrepareSsTraite (ThePiece);
    //
    AjouteTOB (ThePiece,'PIECE',fTobPiece);
    AjouteTOB (ThePiece,'BTPIECEMILIEME',fTobrepart);
    AjouteTOB (ThePiece,'PARAMDOC',TOBParamEdt);
    AjouteTOB (ThePiece,'PARPIECE',TOBParPiece);
    AjouteTOB (ThePiece,'DOCTEXTEDEBFIN',TOBEntetePiedDoc);
    AjouteTOB (ThePiece,'INTEXTEDEBFIN',fTOBLIENOLE);
    AjouteTOB (ThePiece,'TIERS',fTOBTiers);
    AjouteTOB (ThePiece,'AFFAIRE',fTOBAffaire);
    AjouteTOB (ThePiece,'AFFAIREDEV',fTOBAffaireDev);
    AjouteTOB (ThePiece,'PIEDBASE',fTOBBases);
    AjouteTOB (ThePiece,'PIEDBASERG',fTOBBasesRG);
    AjouteTOB (ThePiece,'PIEDECHE',fTOBEches);
    AjouteTOB (ThePiece,'PIEDPORT',fTOBPorcs);
    AjouteTOB (ThePiece,'PIECEADRESSE',fTOBAdresses);
    AjouteTOB (ThePiece,'PIECERG',fTOBPieceRG);
    AjouteTOB (ThePiece,'DEVISES',nil);
    AjouteTOB (ThePiece,'ACOMPTESCPTA',nil);
    AjouteTOB (ThePiece,'TIMBRES',fTOBTimbres);
    //
    if fAvancement then
    begin
      AjouteTOB (ThePiece,'AVANCEMENTG',nil); // Mise en place des infos d'avancement général pièce
    end;
    AjouteTOB (ThePiece,'RETENUESDIV',fTOBPorcs);
    AjouteTOB (ThePiece,'RETENUESTTC',fTOBPorcs);
    //
    ThePiece.AddChampSupValeur('NBRLIGNES',0);
    //
    // -------------------------------------------------------
    if pos(fTOBpiece.GetValue('GP_NATUREPIECEG'),'FBT;DAC;FBP;BAC;')> 0 then
    begin
      if not fInternal then  
      begin
        if fNbSTTraitDir <= 5 then
        begin
          AjouteTOB (ThePiece,'SSTRAITANCEINT',fTOBpieceTrait );
        end else
        begin
          TheSsTraitant.Dupliquer(ThePiece, false, true);
          AjouteTOB (TheSsTraitant,'SSTRAITANCEINT',fTOBpieceTrait );
          TheSsTraitant.detail.sort('LIG');
        end;
      end else
      begin
        if fNbSTTraitDir <= 5 then
        begin
          AjouteTOB (ThePiece,'SSTRAITANCE',fTOBSsTraitant );
        end else
        begin
          TheSsTraitant.Dupliquer(ThePiece, false, true);
          AjouteTOB (TheSsTraitant,'SSTRAITANCE',fTOBSsTraitant );
          TheSsTraitant.detail.sort('LIG');
        end;
      end;
    end;
    // -------------------------------------------------------
    if (ThePiece.GetValue('BPD_IMPTABTOTSIT')='X') then //(fAvancement) and
    begin
      PreparePiedbase (TheRecap);
      PrepareSsTraite (TheRecap);

      AjouteTOB (TheRecap,'PIECE',fTobPiece);
      AjouteTOB (TheRecap,'BTPIECEMILIEME',fTobrepart);
      AjouteTOB (TheRecap,'PARAMDOC',TOBParamEdt);
      AjouteTOB (TheRecap,'PARPIECE',TOBParPiece);
      AjouteTOB (TheRecap,'DOCTEXTEDEBFIN',TOBEntetePiedDoc);
      AjouteTOB (TheRecap,'INTEXTEDEBFIN',fTOBLIENOLE);
      AjouteTOB (TheRecap,'TIERS',fTOBTiers);
      AjouteTOB (TheRecap,'AFFAIRE',fTOBAffaire);
    	AjouteTOB (TheRecap,'AFFAIREDEV',fTOBAffaireDev);
      AjouteTOB (TheRecap,'PIEDBASE',fTOBBases);
      //
      AjouteTOB (TheRecap,'PIEDBASERG',fTOBBasesRG);
      AjouteTOB (TheRecap,'PIEDECHE',fTOBEches);
      AjouteTOB (TheRecap,'PIEDPORT',fTOBPorcs);
      AjouteTOB (TheRecap,'PIECEADRESSE',fTOBAdresses);
      AjouteTOB (TheRecap,'PIECERG',fTOBPieceRG);
      AjouteTOB (TheRecap,'DEVISES',nil);
      AjouteTOB (TheRecap,'ACOMPTESCPTA',nil);
      AjouteTOB (TheRecap,'AVANCEMENTG',nil); // Mise en place des infos d'avancement général pièce
    	AjouteTOB (TheRecap,'TIMBRES',fTOBTimbres);
      //
      if pos(fTOBpiece.GetValue('GP_NATUREPIECEG'),'FBT;DAC;FBP;BAC;')> 0 then
      begin
        if not fInternal then
        begin
          if fNbSTTraitDir <= 5 then
          begin
            AjouteTOB (TheRecap,'SSTRAITANCEINT',fTOBpieceTrait );
          end;
        end else
        begin
          if fNbSTTraitDir <= 5 then
          begin
            AjouteTOB (TheRecap,'SSTRAITANCE',fTOBSsTraitant );
          end;
        end;
      end;

      TheRecap.AddChampSupValeur('NBRLIGNES',0);
    end;
    //
    if not FromMemoFacturation then
    begin
      if (fAvancement) then
      begin
        SetMontantMarcheOldVersion (ThePiece);
        PositionneMarcheFromPiece (TheRecap,ThePiece);
      end;
    end;
    //
    for Indice := 0 to fTobPiece.detail.count -1 do
    begin
      LaLigne := fTobPiece.detail[Indice];
      TypeLigne := LaLigne.GetValue('GL_TYPELIGNE');
      Niveau := LaLigne.GetValue('GL_NIVEAUIMBRIC');
      if (FromMemoFacturation) and (fAvancement)  and (first) and (TypeLigne = 'ART') and (LaLIgne.GetValue('GL_PIECEORIGINE')<>'') then
      begin
        if (fAvancement) then
        begin
          SetMontantMarche(ThePiece,LaLigne);
          PositionneMarcheFromPiece (TheRecap,ThePiece);
          first := false;
        end;
      end;
      //
      if LaLigne.GetValue('GL_NONIMPRIMABLE')='X' Then Continue;
      if IsFinParagraphe(LaLigne) and (Niveau = 1) and (ThePiece.GetValue('BPD_IMPTOTPAR')<>'X') then Continue;
      if IsFinParagraphe(LaLigne) and (Niveau > 1) and (ThePiece.GetValue('BPD_IMPTOTSSP')<>'X') then Continue;
      if IsLigneDetail(nil,LaLigne) then continue;  // enleve les lignes de commentaires de détail (affichage)
      if IsSousDetail(LaLigne) then Continue;       // enleve les lignes de sous detail d'ouvrage (modif sous detail)

      TheLigne := InsereTOB ('LIGNE',ThePiece);
      //
    	if (fLastNumOrdre = 1) and
//      	 (not IsBlocNoteVide(ThePiece.getValue('LO_OBJET_ENTETEDOC'))) and
         (ThePiece.getValue('BPD_SAUTAPRTXTDEB')='X') Then
         			TheLigne.PutValue('SAUTDEPAGE','X');

      TheArticle := fTOBArticles.FindFirst(['GA_ARTICLE'],[LaLigne.getValue('GL_ARTICLE')],true);
      AjouteTOB (TheLigne,'ARTICLES',TheArticle);
      AjouteTOB (TheLigne,'LIGNES',LaLigne);
      AjouteTOB (TheLigne,'LIGNEFAC',LaLigne);
//
    	if fAvancement then
    	begin
      	AjouteTOB (ThePiece,'AVANCLIGNE',LaLigne); // cumul sur entete de pièce
      end;
      if GetParamSocSecur('SO_METRESEXCEL',False) then AjouteTOB (TheLigne,'METRE',nil);
      // Gestion du recap des paragraphes
      if (ThePIece.getValue('BPD_IMPRECPAR')='X') then
      begin
      	if IsDebutParagraphe (LaLigne) Then
        begin
          TheParag := InsereTOB('LIGNE',TOBParag,False);
          LaTotalisation := FindLaFinParag (LaLigne,Indice);
          AjouteTOB (TheParag,'PARAG',TheLigne);
          PoselesTotaux (ThePiece,Theparag,LaTotalisation);
        end;
      end;
      //
      LaLigne.AddChampSupValeur('TYPEPRESENT',0);
      //
      LaLigne.putValue('TYPEPRESENT',LaLigne.GetValue('GL_TYPEPRESENT'));
      if ThePiece.getValue('BPD_TYPESSD') = -1 then
      begin
        LaLigne.PutValue('TYPEPRESENT',DOU_AUCUN);
      end else if ThePiece.GetValue('BPD_TYPESSD') > 0 then
      begin
        LaLigne.PutValue('TYPEPRESENT',ThePiece.GetValue('BPD_TYPESSD'));
      end;
      //
      // Ajout des lignes de métrés s'il y en a  et si demandé
      //
      if ThePiece.getValue('BPD_IMPMETRE')='T' then
      begin
      	AlimentemetresLig (ThePiece,Laligne,Laligne.getValue('GL_NUMORDRE'),fTOBmetres);
      end;
      //
      if (LaLigne.GetValue('GL_TYPEARTICLE') <> 'ARP') and
      	 (LaLigne.GetValue('GL_INDICENOMEN')<>0) and
      	 (LaLigne.GetValue('TYPEPRESENT')<>DOU_AUCUN) then
      begin
      	FromOuvrage := false;
        TOBOuv := fTOBOuvrage.Detail[LaLigne.GetValue('GL_INDICENOMEN')-1];
        if TOBOUv = nil then
        begin
        	TOBOuv := fTOBNomenclature.Detail[LaLigne.GetValue('GL_INDICENOMEN')-1];
        end else FromOuvrage := true;
        if TOBOUV <> nil then
        begin
          for IndiceOuv := 0 to TOBOUV.detail.count -1 do
          begin
          	TOBO := TOBOUV.detail[IndiceOuv];
            if FromOuvrage Then
            begin
            	TheOuvrage := InsereTOB ('LIGNE',ThePiece);
      				TheArticle := fTOBArticles.FindFirst(['GA_ARTICLE'],[TOBO.getValue('BLO_ARTICLE')],true);
      				AjouteTOB (TheOuvrage,'ARTICLES',TheArticle);
              AjouteTOB (TheOuvrage,'OUVRAGE',TOBO,LaLigne);
      				AjouteTOB (TheLigne,'METRE',nil);
            end else
            begin
            	TheOuvrage := InsereTOB ('LIGNE',ThePiece);
      				TheArticle := fTOBArticles.FindFirst(['GA_ARTICLE'],[TOBO.getValue('GLN_ARTICLE')],true);
      				AjouteTOB (TheOuvrage,'ARTICLES',TheArticle);
              AjouteTOB (TheOuvrage,'NOMENCLATURE',TOBO);
      				AjouteTOB (TheLigne,'METRE',nil);
            end;
          end;
        end;
      end;
    end;
    ThePiece.PutValue('NBRLIGNES',fLastNumordre);
    if TOBParag.detail.count > 0 then
    begin
    	AlimenteTotauxParag (ThePiece,TOBparag);
    end;
    //

    if (fAvancement) and (ThePiece.GetValue('BPD_IMPRECSIT')='X') then
    begin
      AlimenteTableauSituations (ThePiece);
    end;

    if (fAvancement) then
    begin
      if (ThePiece.GetValue('BPD_IMPTABTOTSIT')='X') then
      begin
        DefiniRecapTotalSit (TheRecap);
        if ThePiece.GetValue('BPD_IMPRECSIT')='X' then AlimenteTableauSituations (TheRecap);
      end;
    end
    //FV1 : 29/06/2016 - Edition Recap des Retenues Diverse Si Facture directe et pas de tableau recap de situation
    else
    begin
      //Mis en commentaire par BRL 3/08/2016 car inutile et génère une requête très longue dans ConstitueTOBRGPrec
      //DefiniRecapTotalSit (TheRecap);

      //FV1 ; Je L'enlève pas on sait jamais un jour on pourrait en avoir besoin
      if (ThePiece.GetValue('BPD_IMPTABTOTSIT')='-') then
      begin
        TheRetenuesDiverses := tob.create('PIECE', fTobRetenueDiverse, -1);
        TheRetenuesDiverses.Dupliquer(ThePiece, false, true);
        AjouteTOB (TheRetenuesDiverses,'RETENUEDIVERSE',fTOBPorcs);
      end else
      begin
        //FV1 : je le remets car recap sur facture direct ne marche plus...
        DefiniRecapTotalSit (TheRecap);
      end;
    end;
    //
    if fTOBPorcs.detail.count > 0 then
    begin
    	AlimentePort(ThePiece);
    end;

		if (GetInfoParPiece(fTOBpiece.GetValue('GP_NATUREPIECEG'), 'GPP_ESTAVOIR') = 'X') then
    begin
      InverseLesPieces(fTOBPiece, 'PIECE');
      InverseLesPieces(fTOBBases, 'PIEDBASE');
      InverseLesPieces(fTOBEches, 'PIEDECHE');
      InverseLesPieces(fTOBPorcs, 'PIEDPORT');
      InverseLesPieces(fTOBPieceRG, 'PIECERG');
      InverseLesPieces(fTOBBasesRG, 'PIEDBASERG');
    end;

  FINALLY
  	TOBParamEdt.free;
    TOBEntetePiedDoc.free;
    TOBParag.free;
    if ((fPreTOBPiece.getString('GP_NATUREPIECEG')='FBT') or (fPreTOBPiece.getString('GP_NATUREPIECEG')='FBP')) and
        (GetParamSocSecur('SO_BTGESTVIDESUREDIT',false)) and
        (fPreTOBPiece.getvalue('AFF_OKSIZERO')='-') then
    begin
      fTobPiece.Free;
    end;
  END;

end;

procedure TImprPieceViaTOB.AlimenteTotauxParag (ThePiece,TOBparag : TOB);
var Indice,TheOrdre : integer;
		TOBLigne,TOBL : TOB;
begin
	for Indice := 0 to TOBparag.detail.count -1 do
  begin
  	TOBLigne := TOBParag.detail[Indice];
    TOBL := InsereTOB ('LIGNE',ThePiece);
    TheOrdre := TOBL.GetValue('GL_NUMORDRE');
    TOBL.Dupliquer (TOBLigne,false,true);
    TOBL.PutValue('GL_NUMORDRE',TheOrdre); // retablissement
  	if Indice = 0 then TOBL.PutValue ('ISDEBEDTPAR','X');
  	if Indice = TOBparag.detail.count -1 then TOBL.PutValue ('ISFINEDTPAR','X');
  end;
end;

procedure TImprPieceViaTOB.prepareTVA;
var QQ: TQUERY;
begin
	fTOBCodeTaxes.clearDetail;
  QQ := OpenSql ('SELECT * FROM CHOIXCOD WHERE CC_TYPE IN ("TX1","TX2","GX3","GX4","GX5")',True,-1,'',true);
  if not QQ.eof then
  begin
    fTOBCodeTaxes.LoadDetailDB ('CHOIXCOD','','',QQ,false,true);
  end;
  ferme (QQ);
end;

procedure TImprPieceViaTOB.PreparePiedbase(TOBMere: TOB);
var Indice,IndSuiv : integer;
		UnChamp : String;
begin
  for Indice := 1 to 5 do
  begin
  	For indSuiv := 1 to 5 do
    begin
      UnChamp := 'TX'+IntTostr(Indice)+'_LCODETVA'+IntTostr(IndSUiv);
      TOBMERE.AddChampSupValeur (UnChamp,'');
      UnChamp := 'TX'+IntToStr(Indice)+'_BASETAXE'+IntTostr(IndSUiv);
      TOBMERE.AddChampSupValeur (UnChamp,0);
      UnChamp := 'TX'+IntToStr(Indice)+'_TAUXTAXE'+IntTostr(IndSUiv);
      TOBMERE.AddChampSupValeur (UnChamp,0);
      UnChamp := 'TX'+IntToStr(Indice)+'_VALEURTAXE'+IntTostr(IndSUiv);
      TOBMERE.AddChampSupValeur (UnChamp,0);
      UnChamp := 'TX'+IntToStr(Indice)+'_VALEURDEV'+IntTostr(IndSUiv);
      TOBMERE.AddChampSupValeur (UnChamp,0);
    end;
  end;
end;

procedure TImprPieceViaTOB.PrepareSsTraite(TOBMere: TOB);
var Indice  : integer;
    IndSuiv : integer;
		UnChamp : String;
begin

  TOBMERE.AddChampSupValeur ('SSTRAITANCE','-');
  TOBMERE.AddChampSupValeur ('TOTSSTRAIT',0.00);

  //for Indice := 1 to 5 do ==> ???????????? on crée 5 fois le même enreg au même endroit ???????????
  //begin
  	For indSuiv := 1 to 6 do
    begin
      UnChamp := 'SST'+'_FOURNISSEUR'+IntTostr(IndSUiv);
      TOBMERE.AddChampSupValeur (UnChamp,'');
      UnChamp := 'SST'+'_TOTALHTDEV' +IntTostr(IndSUiv);
      TOBMERE.AddChampSupValeur (UnChamp,0.00);
      UnChamp := 'SST'+'_TOTALTTCDEV'+IntTostr(IndSUiv);
      TOBMERE.AddChampSupValeur (UnChamp,0.00);
      UnChamp := 'SST'+'MONTANTPA'   +IntTostr(IndSUiv);
      TOBMERE.AddChampSupValeur (UnChamp,0.00);
      UnChamp := 'SST'+'MONTANTFAC'  +IntTostr(IndSUiv);
      TOBMERE.AddChampSupValeur (UnChamp,0.00);
      UnChamp := 'SST'+'_MONTANTREGL'+IntTostr(IndSUiv);
      TOBMERE.AddChampSupValeur (UnChamp,0.00);
    end;
  //end;

end;

function TImprPieceViaTOB.FindLibre(Prefixe, Code: string): string;
var ThisTOB : TOB;
begin
	result := '';
  ThisTob := fTOBCodeTaxes.findFirst(['CC_TYPE','CC_CODE'],[Prefixe,Code],true);
  if ThisTOB <> nil then
  begin
  	result := ThisTOB.GetValue('CC_LIBRE');
  end;
end;

(*
procedure TImprPieceViaTOB.ReDefiniNumLig(ThePiece: TOB);
var Indice : integer;
begin
	for Indice := 0 to ThePiece.detail.count -1 do
  begin
  	ThePiece.putValue('GL_NUMLIGNE',Indice+1);
  end;
end;
*)

function TImprPieceViaTOB.IsBlocNoteVide (BlocNote : String; Fichier : boolean=False) : boolean;
begin
	result := true;
	if (Fichier) and (Not FileExists(BlocNote)) then
  	Exit
  else
  	StringToRich(fTexte, BlocNote);
  if (Length(fTexte.Text) <> 0) and (ftexte.Text <> #$D#$A) then result := false;
end;

function TImprPieceViaTOB.FindLaFinParag (TheLigne : TOB; IndiceDep : integer) : TOB;
var Indice : integer;
		TOBL : TOB;
begin
	result := nil;
	for Indice := IndiceDep to fTOBpiece.detail.count -1 do
  begin
    TOBL := fTOBPiece.detail[Indice];
    if isFinParagraphe (TOBL,TheLigne.getValue('GL_NIVEAUIMBRIC')) then result := TOBL;
    if result <> nil then Exit; // la ligne de fin de paragraphe a été trouvée : on sort
  end;
end;

procedure TImprPieceViaTOB.PoselesTotaux (thePiece,TheLigne,LaTotalisation : TOB);
begin
	if LaTotalisation = nil then exit;
	if ((TheLigne.GetValue('GL_NIVEAUIMBRIC')=1)and (thePiece.GetValue('BPD_IMPTOTPAR')='X')) or
  	 ((TheLigne.GetValue('GL_NIVEAUIMBRIC')>1) and (thePiece.GetValue('BPD_IMPTOTSSP')='X')) then
  begin
    TheLIgne.putValue('GL_MONTANTHTDEV',LaTotalisation.GetValue('GL_MONTANTHTDEV'));
    TheLIgne.putValue('GL_MONTANTHT',LaTotalisation.GetValue('GL_MONTANTHT'));
    TheLIgne.putValue('GL_TOTALHTDEV',LaTotalisation.GetValue('GL_TOTALHTDEV'));
    TheLIgne.putValue('GL_TOTALHT',LaTotalisation.GetValue('GL_TOTALHT'));
    TheLIgne.putValue('GL_MONTANTTTCDEV',LaTotalisation.GetValue('GL_MONTANTTTCDEV'));
    TheLIgne.putValue('GL_MONTANTTTC',LaTotalisation.GetValue('GL_MONTANTTTC'));
    TheLIgne.putValue('GL_TOTALTTCDEV',LaTotalisation.GetValue('GL_TOTALTTCDEV'));
    TheLIgne.putValue('GL_TOTALTTC',LaTotalisation.GetValue('GL_TOTALTTC'));
    TheLIgne.putValue('BLF_MTCUMULEFACT',LaTotalisation.GetValue('BLF_MTCUMULEFACT'));
  end;
end;

procedure TImprPieceViaTOB.AlimentePort(ThePiece : TOB);
var Indice : integer;
		TOBLigne,TOBL : TOB;
begin
	for Indice := 0 to fTOBPorcs.detail.count -1 do
  begin
  	TOBLigne := fTOBPorcs.detail[Indice];
    if TOBLigne.getValue('GPT_FRAISREPARTIS')='X' then continue;
    if (Pos(TOBLigne.GetString('GPT_TYPEPORT'),'PT;MIC;MTC;')>0) then continue;
    //
    TOBL := InsereTOB ('LIGNE',ThePiece);
    AjouteTOB (TOBL,'LESPORTS',TOBLigne);
  end;
end;

procedure TImprPieceViaTOB.InsereChampMere(TOBMere,LaTOBgrandMere : TOB);
var iTable,Indice : integer;
    prefixe,newChamp : String;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;
    NomChamps : string;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  Table := Mcd.getTable(Mcd.PrefixeToTable('GP'));
  FieldList := Table.Fields;
  FieldList.Reset();

	While FieldList.MoveNext do
  begin
		NomChamps := (FieldList.Current as IFieldCOM).name;
  	Prefixe := 'GL';
    NewChamp := prefixe+copy(NomChamps,3,255);
		if TOBMere.FieldExists (NewChamp) then
    		TOBMere.PutValue (NewChamp,LaTOBgrandMere.GetValue(NomChamps));
  end;
end;

procedure TImprPieceViaTOB.SetDocument(Cledoc: R_cledoc; isDGD : boolean);
begin
	fIsDGD := IsDGD;
  fInternal := true;
  fcledoc := cledoc;
  VenteAchat := GetInfoParPiece(Cledoc.NaturePiece, 'GPP_VENTEACHAT');
  AlloueTout;
	loadlesTobs;
end;

procedure TImprPieceViaTOB.ClearInternal;
begin
	RazTout;
end;

procedure TImprPieceViaTOB.AlloueTout;
begin
  fpreTOBPiece := TOB.Create('PIECE', nil, -1);
  AddLesSupEntete(fpreTOBPiece);
  // ---
  fTOBBases := TOB.Create('BASES', nil, -1);
  fTOBEches := TOB.Create('Les ECHEANCES', nil, -1);
  fTOBPorcs := TOB.Create('PORCS', nil, -1);
  fTOBTiers := TOB.Create('TIERS', nil, -1);
  // ---
  fTOBSsTraitant    := TOB.Create('LES SSTRAITANTS', nil, -1);
  TobSsTraitantImp  := TOB.Create('IMP SSTRAITANTS', nil, -1);
  fTOBPieceTrait := TOB.Create ('LES LIGNES EXTRENALISE',nil,-1);
  //
  ftobRetenueDiverse:= TOB.Create('LES RETENUES DIVERSES', nil, -1);
  //
  TOBTiers.AddChampSup('RIB', False);
  fTOBArticles := TOB.Create('ARTICLES', nil, -1);
  fTOBConds := TOB.Create('CONDS', nil, -1);
  fTOBComms := TOB.Create('COMMERCIAUX', nil, -1);
  // Adresses
  fTOBAdresses := TOB.Create('LESADRESSES', nil, -1);
  if GetParamSoc('SO_GCPIECEADRESSE') then
  begin
    TOB.Create('PIECEADRESSE', TOBAdresses, -1); {Livraison}
    TOB.Create('PIECEADRESSE', TOBAdresses, -1); {Facturation}
  end else
  begin
    TOB.Create('ADRESSES', TOBAdresses, -1); {Livraison}
    TOB.Create('ADRESSES', TOBAdresses, -1); {Facturation}
  end;
  fTOBNomenclature := TOB.Create('NOMENCLATURES', nil, -1);
  fTOBSerie := TOB.Create('', nil, -1);
  fTOBAcomptes := TOB.Create('', nil, -1);
  // Affaires
  TOBAffaire := TOB.Create('AFFAIRE', nil, -1);
  //Saisie Code Barres
  // MODIF BTP
  // Ouvrages
  fTOBOuvrage := TOB.Create('OUVRAGES', nil, -1);
  // textes debut et fin
  fTOBLIENOLE := TOB.Create('LIENS', nil, -1);
  // retenues de garantie
  fTOBPieceRG := TOB.create('PIECERRET', nil, -1);
  // Bases de tva sur RG
  fTOBBasesRG := TOB.create('BASESRG', nil, -1);
  // --
  //Affaire-Formule des variables
  fTOBFormuleVar := TOB.Create('AFORMULEVAR', nil, -1);
  TheRepartTva := TREPARTTVAMILL.create (fForm);
  TheRepartTva.TOBBases := TOBBases;
  TheRepartTva.TOBpiece := fpreTOBPiece;
  TheRepartTva.TOBOuvrages := TOBOUvrage;
  TheRepartTva.Charge;
  fTOBRepart := theRepartTva.Tobrepart;
  fTOBLOTS := tOB.Create ('LES LOTS',nil,-1);
  fTOBMetre := TOB.Create('LES METRES',nil,-1);
  fTOBMetres := TOB.Create('LES METRES LIGNE',nil,-1);
  fTOBTimbres := TOB.Create('LES TIMBREs',nil,-1);


end;

procedure TImprPieceViaTOB.DetruitTout;
begin
  //
  if fpretobpiece        <> nil then FreeAndNil(fpreTOBPiece);
  if ftobbases        <> nil then FreeAndNil(fTOBBases);
  if ftobEches        <> nil then FreeAndNil(fTOBEches);
  if ftobporcs        <> nil then FreeAndNil(fTOBPorcs);
  if fTOBTiers        <> nil then FreeAndNil(fTOBTiers);
  if fTOBSsTraitant   <> nil then FreeAndNil(fTOBSsTraitant);
  if ftobRetenueDiverse<>nil then FreeAndNil(ftobRetenueDiverse);
  //
  if ftobArticles     <> nil then FreeAndNil(fTOBArticles);
  if ftobconds        <> nil then FreeAndNil(fTOBConds);
  if ftobComms        <> nil then FreeAndNil(fTOBComms);
  // Adresses
  if ftobAdresses     <> nil then FreeAndNil(fTOBAdresses);
  if fTOBNomenclature <> nil then FreeAndNil(fTOBNomenclature);
  if fTOBSerie        <> nil then FreeAndNil(fTOBSerie);
  if fTOBAcomptes     <> nil then FreeAndNil(fTOBAcomptes);
  //Saisie Code Barres
  // MODIF BTP
  // Ouvrages
  if fTOBOuvrage      <> nil then FreeAndNil(fTOBOuvrage);
  // textes debut et fin
  if fTOBLIENOLE      <> nil then FreeAndNil(fTOBLIENOLE);
  // retenues de garantie
  if fTOBPieceRG      <> nil then FreeAndNil(fTOBPieceRG);
  // Bases de tva sur RG
  if fTOBBasesRG      <> nil then FreeAndNil(fTOBBasesRG);
  // --
  //Affaire-Formule des variables
  if fTOBFormuleVar   <> nil then FreeAndNil(fTOBFormuleVar);

  if fTOBLOTS         <> nil then FreeAndNil(fTOBLOTS);
  if fTOBMetre        <> nil then FreeAndNil(fTOBMetre);
  if fTOBMetres       <> nil then FreeAndNil(fTOBMetres);
  if fTOBTimbres      <> nil then FreeAndNil(fTOBTimbres);

  FreeAndNil(TheRepartTva);
end;

procedure TImprPieceViaTOB.loadlesTobs;
var Q : Tquery;
		indiceOuv,Indice : integer;
		OptimizeOuv :TOptimizeOuv;
begin
	OptimizeOuv := TOptimizeOuv.create;
	IndiceOuv := 1;
  //
  LoadLesSsTraitants(fCleDoc, fTOBSsTraitant);
  //

  LoadPieceLignes(fCleDoc, fPreTobPiece);

	FillChar(DEV, Sizeof(DEV), #0);
  DEV.Code := fPreTobPiece.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  //
  loadlesMetres (fCleDoc,fTOBMetres);
  //
  AddLesSupEntete (fPreTobPiece);
  //
  for Indice := 0 to fPreTobPiece.detail.count -1 do
  begin
    AddLesSupLigne (fPreTobPiece.detail[indice],false);
    InitLesSupLigne (fPreTobPiece.detail[indice]);
    if (pos(fcledoc.Naturepiece ,'FBT;FBP')>0) then
      GestionAnciennesfactureation (fPreTobPiece.detail[indice],fPreTobPiece.GetValue('AFF_GENERAUTO'),DEV);
  end;
  // lecture affaire
  if fPreTobPiece.getValue('GP_AFFAIRE') <> '' then
  begin
    StockeCetteAffaire (fPreTobPiece.getValue('GP_AFFAIRE'));
    fTOBaffaire := FindCetteAffaire (fPreTobPiece.getValue('GP_AFFAIRE'));
  end;
  if fPreTobPiece.getValue('GP_AFFAIREDEVIS') <> '' then
  begin
    StockeCetteAffaire (fPreTobPiece.getValue('GP_AFFAIREDEVIS'));
    fTOBAffaireDev  := FindCetteAffaire (fPreTobPiece.getValue('GP_AFFAIREDEVIS'));
  end;
  // Lecture bases
  Q := OpenSQL('SELECT * FROM PIEDBASE WHERE ' + WherePiece(fCleDoc, ttdPiedBase, False), True,-1,'',true);
  fTOBBases.LoadDetailDB('PIEDBASE', '', '', Q, False);
  Ferme(Q);
 // Lecture Echéances
  Q := OpenSQL('SELECT * FROM PIEDECHE WHERE ' + WherePiece(fCleDoc, ttdEche, False), True,-1,'',true);
  fTOBEches.LoadDetailDB('PIEDECHE', '', '', Q, False);
  Ferme(Q);
  // Lecture Ports
  Q := OpenSQL('SELECT * FROM PIEDPORT WHERE ' + WherePiece(fCleDoc, ttdPorc, False), True,-1,'',true);
  fTOBPorcs.LoadDetailDB('PIEDPORT', '', '', Q, False);
  Ferme(Q);
  // Lecture Lignes Formule
//  LoadLesFormules(fTOBPiece,fTOBLigFormule,CleDocfCleDoc) ;
  // Lecture Adresses
  LoadLesAdresses(fPreTobPiece, fTOBAdresses);
  // Lecture Nomenclatures
  LoadLesNomen(fPreTobPiece, fTOBNomenclature, fTOBArticles, fCleDoc);
  // Lecture ACompte
  LoadLesAcomptes(fPreTobPiece, fTOBAcomptes, fCleDoc);
  // Modif BTP
  // chargement textes debut et fin
  LoadLesBlocNotes([tModeBlocNotes], fTOBLIENOLE, fCleDoc);
  // Chargement des retenues de garantie et Tva associe
  if GetParamSoc('SO_RETGARANTIE') then LoadLesRetenues(fPreTobPiece, fTOBPieceRG , fTOBBasesRG,taModif);
  {$IFDEF BTP}
  // Lecture Ouvrages
  LoadLesOuvrages(fPreTobPiece, fTOBOuvrage, fTOBArticles, fCleDoc, IndiceOuv,OptimizeOuv);
  {$ENDIF}
  LoadLesArticles(fPreTobPiece);
  CalculeSousTotauxPiece(fPreTobPiece);
  LoadlesTimbres (fCleDoc,fTOBTimbres);
(*
  TheMetre.Action := taConsult;
  TheMetre.Duplic := false;
  TheMetre.User := V_PGI.User;
  TheMetre.Ouvrages := fTOBOuvrage;
  TheMetre.Document := fTOBPiece;
*)
{$IFDEF BTP}
	OptimizeOuv.fusionneArticles (fTOBArticles);
{$ENDIF}
  OptimizeOUv.free;
  TheRepartTva.TOBBases := TOBBases;
  TheRepartTva.TOBpiece := fPreTobPiece;
  TheRepartTva.TOBOuvrages := TOBOUvrage;
  TheRepartTva.Charge; // voila voila
  fTOBRepart := TheRepartTva.Tobrepart;
  if fcledoc.Naturepiece <> '' then
  begin
  	Q := OpenSql ('SELECT * FROM PARPIECE WHERE GPP_NATUREPIECEG="'+fCledoc.NaturePiece+'"',true,-1,'',true);
    ftobParPiece.SelectDB ('',Q);
    ferme (Q);
  end;
  LoadLaTOBPieceTrait(fTOBpieceTrait,fCleDoc,'');
  fNbSTTraitDir := NbSSTraitDir(fTOBpieceTrait);
end;

procedure TImprPieceViaTOB.RazTout;
begin

  if fInternal then
  begin
    if fPreTOBPiece   <> nil then fPreTOBPiece.clearDetail;
    if fPreTOBPiece   <> nil then fPreTOBPiece.InitValeurs;
    if fTOBSsTraitant <> nil then fTOBSsTraitant.Cleardetail;
    if ftobRetenueDiverse <> nil then ftobRetenueDiverse.ClearDetail;
    if ftobAdresses   <> nil then fTOBAdresses.cleardetail;
    if ftobTiers      <> nil then fTOBTiers.InitValeurs;
    if ftobLienOle    <> nil then fTOBLIENOLE.clearDetail;
    if ftobArticles   <> nil then fTOBArticles.clearDetail;
    if ftobpieceRG    <> nil then fTOBPieceRG.cleardetail;
    if ftobBases      <> nil then fTOBBases.cleardetail;
    if ftobEches      <> nil then fTOBEches.cleardetail;
    if fTOBPorcs      <> nil then fTOBPorcs.cleardetail;
    if fTobConds      <> nil then fTOBConds.cleardetail;
    if ftobComms      <> nil then fTOBComms.cleardetail;
    if fTOBOuvrage    <> nil then fTOBOuvrage.cleardetail;
    if fTOBNomenclature <> nil then fTOBNomenclature.cleardetail;
    if fTOBLOTS       <> nil then fTOBLOTS.cleardetail;
    if fTOBSerie      <> nil then fTOBSerie.cleardetail;
    if fTOBAcomptes   <> nil then fTOBAcomptes.cleardetail;
    if fTOBBasesRG    <> nil then fTOBBasesRG.cleardetail;
    if fTOBFormuleVar <> nil then fTOBFormuleVar.cleardetail;
    if fTOBMetre      <> nil then fTOBMetre.cleardetail;
    if fTOBMetres     <> nil then fTOBMetres.cleardetail;
    if fTOBCodeTaxes  <> nil then fTOBCodeTaxes.cleardetail;
    if TheRepartTva   <> nil then TheRepartTva.Initialise;
    if fTOBAffaireDev <> nil then fTOBAffaireDev := nil;
    ReinitTOBAffaires;
  end else
  begin
		fTOBPieceImp.clearDetail;
    fTOBPieceRecap.clearDetail;
    fTobSsTraitantImp.clearDetail;
  end;
end;

procedure TImprPieceViaTOB.LoadLesArticles (ThePiece : TOB);
var NbArt : integer;
  StSelect, StSelectDepot, StSelectDepotLot, StSelectCon, {StArticle,}
  StSelectCatalogu, StSelectFormule, {StCodeArticle,} stWhereLigne, ListeDepot,
  VteAch {, RefGen, Statut}: string;
  TOBDispo, TOBDispoLot, TOBArt, TobDispoArt, TobDispoLotArt, TOBCata: TOB;
  QArticle, QDepot, QDepotLot, QCon: TQuery;
  {OkTab,} DoubleDepot: Boolean;
begin
  StSelect := 'SELECT * FROM ARTICLE';
  StSelectDepot := 'SELECT * FROM DISPO';
  StSelectDepotLot := 'SELECT * FROM DISPOLOT';
  StSelectCon := 'SELECT * FROM CONDITIONNEMENT';
  StSelectCatalogu := 'SELECT * FROM CATALOGU';
  StSelectFormule := 'SELECT * FROM ARTICLEQTE';

  ListeDepot := '';
  DoubleDepot := IsTransfert(ThePiece.GetValue('GP_NATUREPIECEG'));
  if DoubleDepot then
    ListeDepot := '"' + ThePiece.GetValue('GP_DEPOT') + '","' + ThePiece.GetValue('GP_DEPOTDEST') + '"'
  else ListeDepot := '"' + ThePiece.GetValue('GP_DEPOT') + '"';

  TOBDispo := TOB.CREATE('Les Dispos', nil, -1);
  TOBDispoLot := TOB.CREATE('Les Dispolots', nil, -1);
  TOBCata := TOB.CREATE('Les Catalogues', nil, -1);

{$IFDEF AFFAIRE}
// gestion de la saisie de facture depuis l'affaire avec reprise des lignes de l'affaire
// il faut charger les articles issus des lignes d'affaire  gm 01/07/03
    stWhereLigne := WherePiece(fCleDoc, ttdLigne, False);
(*
  if (CleDocAffaire.NumeroPiece = 0) then
    stWhereLigne := WherePiece(CleDoc, ttdLigne, False)
  else
    stWhereLigne := WherePiece(CleDocAffaire, ttdLigne, False);
*)
{$else}
  stWhereLigne := WherePiece(fCleDoc, ttdLigne, False) ;
{$endif}

  QArticle := OpenSQL(StSelect + ' WHERE GA_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne + ')', True,-1,'',true);
  if not QArticle.EOF then fTOBArticles.LoadDetailDB('ARTICLE', '', '', QArticle, True, True);
  Ferme(QArticle);

  if DoubleDepot then
    QDepot := OpenSQL(StSelectDepot + ' WHERE GQ_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne +
      ') AND GQ_DEPOT IN (' + ListeDepot + ') AND GQ_CLOTURE="-"', True,-1,'',true)
  else QDepot := OpenSQL(StSelectDepot + ' WHERE GQ_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne +
      ') AND GQ_DEPOT IN (SELECT DISTINCT GL_DEPOT FROM LIGNE WHERE ' + stWhereLigne + ') AND GQ_CLOTURE="-"', True,-1,'',true);
  if not (CtxMode in V_PGI.PGIContexte) then
  begin
    if DoubleDepot then
      QDepotLot := OpenSQL(StSelectDepotLot + ' WHERE GQL_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne +
        ') AND GQL_DEPOT IN (' + ListeDepot + ')', True,-1,'',true)
    else QDepotLot := OpenSQL(StSelectDepotLot + ' WHERE GQL_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne +
        ') AND GQL_DEPOT IN (SELECT DISTINCT GL_DEPOT FROM LIGNE WHERE ' + stWhereLigne + ')', True,-1,'',true);
    if not QDepotLot.EOF then TOBDispoLot.LoadDetailDB('DISPOLOT', '', '', QDepotLot, True, True);
    Ferme(QDepotLot);
  end;
  if not QDepot.EOF then TOBDispo.LoadDetailDB('DISPO', '', '', QDepot, True, True);
  Ferme(QDepot);
(*
  //Formules de qté  JS 30/10/2003
  QFormule := OpenSQL(StSelectFormule + ' WHERE GAF_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne + ')', True);
  if not QFormule.EOF then TOBAFormule.LoadDetailDB('ARTICLEQTE', '', '', QFormule, True, True);
  Ferme(QFormule);
*)
  QCon := OpenSQL(StSelectCon + ' WHERE GCO_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne +
    ') ORDER BY GCO_NBARTICLE', True,-1,'',true);
  if not QCon.EOF then TOBConds.LoadDetailDB('CONDITIONNEMENT', '', '', QCon, True, True);
  Ferme(QCon);
(*
  if (VenteAchat = 'ACH') then
  begin
    QCata := OpenSQL(StSelectCatalogu + ' WHERE GCA_TIERS IN (SELECT DISTINCT GL_TIERS FROM LIGNE WHERE ' + stWhereLigne +
      ') AND GCA_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne + ')', True);
    if not QCata.EOF then fTOBCata.LoadDetailDB('CATALOGU', '', '', QCata, True, True);
    Ferme(QCata);
  end;
*)
  //Affecte les stocks aux articles sélectionnés
  for NbArt := 0 to fTOBArticles.detail.Count - 1 do
  begin
    with fTOBArticles.detail[NbArt] do
    begin
      AddChampSupValeur('UTILISE', '-', False);
      AddChampSupValeur('REFARTSAISIE', '', False);
      AddChampSupvaleur('REFARTBARRE', '', False);
      AddChampSupValeur('REFARTTIERS', '', False);
    end;
    // FIN MODIF DC

    TOBArt := fTOBArticles.detail[NbArt];
    TobDispoArt := TOBDispo.FindFirst(['GQ_ARTICLE'], [TOBArt.GetValue('GA_ARTICLE')], False);
    //if TobDispoArt<>nil then //Modif du 18/07/02: Bug chez Negoce
    while TobDispoArt <> nil do
    begin
      TobDispoLotArt := TOBDispoLot.FindFirst(['GQL_ARTICLE', 'GQL_DEPOT'], [TobDispoArt.GetValue('GQ_ARTICLE'), TobDispoArt.GetValue('GQ_DEPOT')], False);
      while TobDispoLotArt <> nil do
      begin
        TobDispoLotArt.Changeparent(TobDispoArt, -1);
        TobDispoLotArt := TOBDispoLot.FindNext(['GQL_ARTICLE', 'GQL_DEPOT'], [TobDispoArt.GetValue('GQ_ARTICLE'), TobDispoArt.GetValue('GQ_DEPOT')], False);
      end;
      TobDispoArt.Changeparent(TOBArt, -1);
      TobDispoArt := TOBDispo.FindNext(['GQ_ARTICLE'], [TOBArt.GetValue('GA_ARTICLE')], False);
      if DoubleDepot then
      begin
        //TobDispoArt:=TOBDispo.FindNext(['GQ_ARTICLE'],[TOBArt.GetValue('GA_ARTICLE')],False);
        if TobDispoArt <> nil then
        begin
          TobDispoLotArt := TOBDispoLot.FindFirst(['GQL_ARTICLE', 'GQL_DEPOT'], [TobDispoArt.GetValue('GQ_ARTICLE'), fTOBPiece.GetValue('GP_DEPOTDEST')],
            False);
          while TobDispoLotArt <> nil do
          begin
            TobDispoLotArt.Changeparent(TobDispoArt, -1);
            TobDispoLotArt := TOBDispoLot.FindNext(['GQL_ARTICLE', 'GQL_DEPOT'], [TobDispoArt.GetValue('GQ_ARTICLE'), fTOBPiece.GetValue('GP_DEPOTDEST')],
              False);
          end;
          TobDispoArt.Changeparent(TOBArt, -1);
        end;
        break;
      end;
    end;
(*
    //Affecte les catalogu aux articles sélectionnés
    TobCataArt := fTOBCata.FindFirst(['GCA_ARTICLE'], [TOBArt.GetValue('GA_ARTICLE')], False);
    while TobCataArt <> nil do
    begin
      for i := 1 to TobCataArt.NbChamps do
      begin
        NomChamp := TobCataArt.GetNomChamp(i);
        if NomChamp = '' then continue;
        TOBArt.AddChampSup(NomChamp, False);
        TOBArt.PutValue(NomChamp, TobCataArt.GetValue(NomChamp));
      end;
      TOBArt.AddChampSup('DPANEW', False);
      TOBArt.PutValue('DPANEW', TOBArt.GetValue('GCA_DPA'));
      TobCataArt := fTOBCata.FindNext(['GCA_ARTICLE'], [TOBArt.GetValue('GA_ARTICLE')], False);
    end;
*)
  end;

  //Formules de qté  JS 30/10/2003
  if VenteAchat='VEN' then VteAch:='VTE'
  else VteAch := VenteAchat;
(*
  for i := 0 to TobAFormule.Detail.Count -1 do
  begin
    TobAF := fTobAFormule.Detail[i];
    QFormule := OpenSql('SELECT * FROM ARTFORMULEVAR WHERE GAV_ARTICLE="'+
                   TobAF.GetValue('GAF_ARTICLE')+'" AND GAV_VENTEACHAT="'+VteAch+'"',True);
    if not QFormule.Eof then TobAF.LoadDetailDB('ARTICLEQTE','','',QFormule,False) ;
    Ferme(QFormule);
  end;
*)
  TOBDispo.Free;
  TOBDispoLot.Free;
  TOBCata.Free;
  DispoChampSupp(TOBArticles);
end;

procedure TImprPieceViaTOB.InitModePaiement (TOBMere : TOB);
begin
	TOBMERE.AddChampSupValeur ('GPE_NUMECHE1',1);
	TOBMERE.AddChampSupValeur ('GPE_MODEPAIE1','');
	TOBMERE.AddChampSupValeur ('GPE_DATECHE1',TobMere.GetValue('GP_DATEPIECE'));
	TOBMERE.AddChampSupValeur ('GPE_MONTANTECHE1',TobMere.GetValue('GP_TOTALHTDEV'));
	TOBMERE.AddChampSupValeur ('GPE_MONTANTDEV1',TobMere.GetValue('GP_TOTALHTDEV'));
end;

function TImprPieceViaTOB.GetInfoParams: boolean;
begin
(*
{$IFDEF LINE}
  Result := (EditionViaTOB(fcledoc)) and (not YaMetre);
{$ELSE}
  Result := EditionViaTOB(fcledoc);
{$ENDIF}
*)
  Result := EditionViaTOB(fcledoc);
  fusable := result;
end;

function TImprPieceViaTOB.GetModeleAssocie(Modele: string): string;
var NaturePiece : string;
begin
	result := Modele;
  NaturePiece :=  fTOBPieceImp.detail[0].getValue('GP_NATUREPIECEG');
  if (NaturePiece = VH_GC.AFNatAffaire) or (NaturePiece = VH_GC.AFNatProposition) or (Pos(NaturePiece ,'FBT;ABT;ABP;FBP;BAC;')>0)  then
  begin
     result:=GetInfoParPiece(NaturePiece, 'GPP_IMPMODELE');
     if result = '' then result := 'DDT';
  end;
end;

procedure TImprPieceViaTOB.AjouteChampsMillieme(TOBMere, TOBAInsere: TOB);
begin
end;

procedure TImprPieceViaTOB.AjouteChampsParPiece(TOBmere, TOBAInsere: TOB);
var ITable,Indice : integer;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;
    NomChamps : string;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  Table := Mcd.getTable(Mcd.PrefixeToTable('GPP'));
  FieldList := Table.Fields;
  FieldList.Reset();

  While FieldList.MoveNext do
  begin
		NomChamps := (FieldList.Current as IFieldCOM).name;
  	//
  	if NomChamps = 'GPP_LIBELLE' then continue;
    // --
		if not TOBMere.FieldExists (NomChamps) then
    begin
      TOBMere.AddChampSupValeur (NomChamps,TOBAInsere.GetValue(NomChamps));
    end else
    begin
      TOBMere.PutValue (NomChamps,TOBAInsere.GetValue(NomChamps));
    end;
  end;
end;

function TImprPieceViaTOB.YaMetre: boolean;
begin
  result := (fTOBMetre.detail.count > 0);
end;

procedure TImprPieceViaTOB.SetMode;
var naturePiece : string;
begin

  // par defaut
  fAvancement := false;
  fisAVA := false;
  //
  Naturepiece := fTOBPiece.getValue('GP_NATUREPIECEG');
  if fTOBPiece.GetValue('AFF_GENERAUTO')='AVA' then fIsAVA := true;
  if (Pos(fTOBPiece.GetValue('AFF_GENERAUTO'),'AVA;DAC;')>0) and (Pos(NaturePiece,'FBT;DAC;FBP;BAC;')>0) then
  begin
    fAvancement := true;
  end;
end;

procedure TImprPieceViaTOB.AjouteChampsAvancementPiece(TOBMere: TOB);
var TOBSit  : TOB;
    Indice  : Integer;
    PosMax  : integer;
    req     : string;
    QQ      : TQuery;
    TotalA  : Double;
begin
  PosMax := 1; if (fIsDGD) then PosMax := 0;
  //
  if (not TOBMere.fieldExists ('NUMSIT')) then
  begin
    fNumeroSit := 0;
    TOBMere.AddChampSupValeur ('NUMSIT',0);
    TOBMere.AddChampSupValeur ('MAXNUMSIT',0);
    TOBMere.AddChampSupValeur ('DATESIT',idate1900);
    TOBMere.AddChampSupValeur ('TOTSITPREC',0); // montant situation précédente
    TOBMere.AddChampSupValeur ('TOTSITPRECBRUT',0); // montant situation précédente BRUT
    TOBMere.AddChampSupValeur ('TOTSITPRECTTC',0); // montant situation précédente TTC
    TOBMere.AddChampSupValeur ('TOTREMPREC',0); // montant remise / situation précédente
    TOBMere.AddChampSupValeur ('TOTESCPREC',0); // montant escompte / situation
    TOBMere.AddChampSupValeur ('TOTPORTPREC',0); // montant Port / situation précédente
    TOBMere.AddChampSupValeur ('TOTPORTSIT',0); // montant Port / situation précédente
    TOBMere.AddChampSupValeur ('TOTMARCHE',0);  // Total du Marché  BRUT
    TOBMere.AddChampSupValeur ('TOTMARCHENET',0);  // Total du Marché  NET
    TOBMere.AddChampSupValeur ('TOTMARCHETTC',0);  // Total du Marché  TTC
    TOBMere.AddChampSupValeur ('TOTCUMULEFACT',0);  // Total du cumulé facturé
    TOBMere.AddChampSupValeur ('TOTSITUATION',0); // Montant situation
    TOBMere.AddChampSupValeur ('TOTSOLDE',0); // Total du solde a regler
    TOBMere.AddChampSupValeur ('TOTREGLE',0);
  end;

//
  TOBMere.Putvalue ('TOTSITPREC',0); // montant situation précédente
  TOBMere.Putvalue ('TOTREMPREC',0); // montant remise / situation précédente
  TOBMere.Putvalue ('TOTESCPREC',0); // montant escompte / situation précédente
  TOBMere.Putvalue ('TOTPORTPREC',0); // montant Port / situation précédente
  TOBMere.Putvalue ('TOTPORTSIT',0); // montant Port / situation
  TOBMere.Putvalue ('TOTSITPRECBRUT',0); // montant situation précédente BRUT
  TOBMere.Putvalue ('TOTSITPRECTTC',0); // montant situation précédente TTC
  TOBMere.Putvalue ('TOTMARCHE',0);  // Total du Marché BRUT
  TOBMere.PutValue ('TOTMARCHENET',0);  // Total du Marché NET
  TOBMere.PutValue ('TOTMARCHETTC',0);  // Total du Marché TTC
  TOBMere.Putvalue ('TOTCUMULEFACT',0);  // Total du cumulé facturé
  TOBMere.Putvalue ('TOTSITUATION',0); // Montant situation
  TOBMere.Putvalue ('TOTSOLDE',0); // Total du solde a regler
  TOBMere.Putvalue ('TOTREGLE',0);
  //
  for Indice := 0 to fTOBPorcs.detail.count -1 do
  begin
    if (not fTOBPorcs.detail[Indice].getboolean('GPT_FRAISREPARTIS')) and (not fTOBPorcs.detail[Indice].getboolean('GPT_RETENUEDIVERSE')) then
    begin
      TOBMere.Putvalue ('TOTPORTSIT',TOBMere.Getvalue ('TOTPORTSIT') + fTOBPorcs.detail[Indice].getValue('GPT_TOTALHTDEV'));
    end;
  end;
  //
  if fTOBSituations.detail.count < 1 then exit;
  TOBSit := fTOBSituations.Detail[fTOBSituations.detail.count-1]; // forcement la derniére
  fNumeroSit := TOBSIT.GetInteger('BST_NUMEROSIT');
  TOBMere.Putvalue ('NUMSIT',TOBSIT.getValue('BST_NUMEROSIT'));
  TOBMere.Putvalue ('DATESIT',TOBSIT.getValue('BST_DATESIT'));
  //
  if fTOBSituations.detail.count <= PosMax then exit;
  for Indice := 0 to fTOBSituations.detail.count- (1+PosMax) do
  begin
    TOBSit := fTOBSituations.Detail[Indice];
    if (not fIsDGD) or ((fIsDGD) and (fIsAva)) then
    begin
      TOBMere.PutValue ('TOTSITPREC',TOBMere.GetValue ('TOTSITPREC')+TOBSit.GetValue('BST_MONTANTHT'));
      TOBMere.PutValue ('TOTSITPRECTTC',TOBMere.GetValue ('TOTSITPRECTTC')+TOBSit.GetValue('BST_MONTANTTTC'));
    end;
    TotalA := TOBSit.GetValue('ACOMPTESIT') + TOBSit.GetValue('REGLEMENTSIT');
    TOBMere.PutValue ('TOTSOLDE',TOBMere.GetValue ('TOTSOLDE')+(TOBSit.GetValue('BST_MONTANTTTC')- TotalA));
    TOBMere.PutValue ('TOTREGLE',TOBMere.GetValue ('TOTREGLE')+TotalA);
  end;
  if (not fIsDGD) or ((fIsDGD) and (fIsAva)) then
  begin
    req := 'SELECT SUM(GP_TOTALREMISEDEV) AS REMPREC, SUM(GP_TOTALESCDEV) AS ESCPREC FROM PIECE, BSITUATIONS WHERE '+
           'BST_SSAFFAIRE="'+TOBMere.GetValue('GP_AFFAIREDEVIS')+'" AND '+
           'BST_NATUREPIECE="'+TOBMere.GetValue('GP_NATUREPIECEG')+'" AND '+
           'BST_SOUCHE="'+TOBMere.GetValue('GP_SOUCHE')+'" AND '+
           'BST_NUMEROSIT < '+TOBMere.getString('NUMSIT')+' AND '+
           'BST_VIVANTE="X" AND '+
           'GP_NATUREPIECEG=BST_NATUREPIECE AND GP_SOUCHE=BST_SOUCHE AND GP_NUMERO=BST_NUMEROFAC';
    QQ := OpenSql (req,True,-1,'',true);
    if not QQ.eof then //
    begin
      TOBMere.Putvalue ('TOTREMPREC',QQ.findField('REMPREC').asFloat); // montant remise / situation précédente
      TOBMere.Putvalue ('TOTESCPREC',QQ.findField('ESCPREC').asFloat); // montant escompte / situation précédente
    end;
    ferme (QQ);
  //
    req := 'SELECT SUM(GPT_TOTALHTDEV) AS PORTPREC FROM PIEDPORT '+
    			 'WHERE GPT_NATUREPIECEG="'+TOBMere.GetValue('GP_NATUREPIECEG')+'" '+
    			 'AND GPT_SOUCHE="'+TOBMere.GetValue('GP_SOUCHE')+'" '+
           'AND GPT_NUMERO IN ('+
    			 'SELECT BST_NUMEROFAC FROM BSITUATIONS WHERE '+
           'BST_SSAFFAIRE="'+TOBMere.GetValue('GP_AFFAIREDEVIS')+'" AND '+
           'BST_NATUREPIECE="'+TOBMere.GetValue('GP_NATUREPIECEG')+'" AND '+
           'BST_SOUCHE="'+TOBMere.GetValue('GP_SOUCHE')+'" AND '+
           'BST_VIVANTE="X" AND '+
           'BST_NUMEROSIT < '+TOBMere.getString('NUMSIT')+') AND '+
           'GPT_FRAISREPARTIS="-" AND GPT_RETENUEDIVERSE="-"';
    QQ := OpenSql (req,True,-1,'',true);
    if not QQ.eof then TOBMere.Putvalue ('TOTPORTPREC',QQ.findField('PORTPREC').asFloat); // montant Port / situation précédente
    ferme (QQ);
  end;
  //
  req := 'SELECT MAX(BST_NUMEROSIT) AS MAXNUMSIT FROM BSITUATIONS WHERE BST_NATUREPIECE="'+
          TOBmere.getValue('GP_NATUREPIECEG')+'" AND BST_SSAFFAIRE="'+
          TOBmere.getValue('GP_AFFAIREDEVIS')+'"';
  QQ := OpenSql (req,true,-1,'',true);
  if not QQ.eof then TOBMere.putValue ('MAXNUMSIT',QQ.Findfield('MAXNUMSIT').asInteger);
  ferme (QQ);
  //
  TOBMere.putValue('TOTSITPRECBRUT',TOBMere.GetValue ('TOTSITPREC')-TOBMere.Getvalue ('TOTPORTPREC')+
  								 TOBMere.getvalue ('TOTREMPREC')+TOBMere.getvalue ('TOTESCPREC'));
end;

procedure TImprPieceViaTOB.CumuleChampsAvancLigne (TOBmere, TOBAInsere: TOB);
begin
  if TOBAInsere.GetValue('GL_TYPELIGNE')<>'ART' then exit;
  // présente pour faire les cumuls sur entete de pièce (CUMULE FAC ET SITUATIONS)
  TOBMere.PutValue ('TOTCUMULEFACT',TOBMere.GetValue ('TOTCUMULEFACT')+ TOBAInsere.GetValue('BLF_MTCUMULEFACT'));  // Total du cumulé facturé
  TOBMere.PutValue ('TOTSITUATION',TOBMere.GetValue ('TOTSITUATION')+ TOBAInsere.GetValue('BLF_MTSITUATION'));  // Total situation
end;

function TImprPieceViaTOB.IsSamePiece(reference, cledoc: r_cledoc): boolean;
begin
  result := false;
  if (reference.NaturePiece=cledoc.NaturePiece) and
     (reference.Souche=cledoc.Souche) and
     (reference.Indice=cledoc.Indice) and
     (reference.NumeroPiece=cledoc.NumeroPiece) then result := true;
end;


procedure TImprPieceViaTOB.SetMontantMarcheOldVersion (TOBmere : TOB);
var Indice : integer;
    TOBL :  TOB;
    RefCledoc,CurCledoc: R_CLEDOC;
    TypeLigne,TheRef  : string;
begin

  fillchar(refcledoc,sizeof(refcledoc),#0); // init
  for Indice := 0 to fTobPiece.detail.count -1 do
  begin
    TOBL := fTOBpiece.detail[Indice];
    TypeLigne := TOBL.GetValue('GL_TYPELIGNE');
    if (TypeLigne = 'ART') and (TOBL.GetValue('GL_PIECEORIGINE')<>'') then
    begin
      if Copy(TOBL.getValue('GL_PIECEORIGINE'),10,3)='ETU' then TheRef := TOBL.getValue('GL_PIECEPRECEDENTE')
                                                           else TheRef := TOBL.getValue('GL_PIECEORIGINE');
      DecodeRefPiece (TheRef,CurCledoc);
    if not IsSamePiece(CurCledoc ,RefCledoc) then
    begin
      CumuleBudgetFromCledoc (curcledoc,TOBMere);
      RefCledoc := CurCledoc;
    end;
  end;
  end;
end;

procedure TImprPieceViaTOB.CumuleBudgetFromCledoc (Cledoc: R_cledoc;TOBMere : TOB);
var requete : string;
    QQ : TQuery;
    TotalHt,Remise,escompte,TotalTTC : double;
begin
  Requete := 'SELECT GP_TOTALHTDEV,GP_TOTALTTCDEV,GP_TOTALREMISEDEV,GP_TOTALESCDEV FROM PIECE WHERE GP_NATUREPIECEG="'+
             Cledoc.NaturePiece +'" AND '+
             'GP_SOUCHE="'+Cledoc.Souche +'" AND '+
             'GP_NUMERO='+InttoStr(Cledoc.NumeroPiece)+' AND '+
             'GP_INDICEG='+InttoStr(Cledoc.Indice);
  QQ := OpenSql (Requete,True,-1,'',true);
  if not QQ.eof then
  begin
    TotalHt := QQ.findField('GP_TOTALHTDEV').asfloat;
    TotalTTC := QQ.findField('GP_TOTALTTCDEV').asfloat;
    Remise := QQ.findField('GP_TOTALREMISEDEV').asfloat;
    Escompte := QQ.findField('GP_TOTALESCDEV').asfloat;
    TOBMere.PutValue ('TOTMARCHE',TOBMere.GetValue ('TOTMARCHE')+ TotalHT+Remise+Escompte );  // Total du Marché brut
    TOBMere.PutValue ('TOTMARCHENET',TOBMere.GetValue ('TOTMARCHENET')+TotalHt );  // Total du Marché Net
    TOBMere.PutValue ('TOTMARCHETTC',TOBMere.GetValue ('TOTMARCHETTC')+TotalTTC );  // Total TTC du Marché
  end;
  ferme (QQ);
  //
  Requete := 'SELECT SUM(GPT_TOTALHTDEV) AS TOTPORT FROM PIEDPORT WHERE GPT_NATUREPIECEG="'+
             Cledoc.NaturePiece+'" AND '+
             'GPT_SOUCHE="'+Cledoc.Souche+'" AND '+
             'GPT_NUMERO='+InttoStr(Cledoc.NumeroPiece)+' AND '+
             'GPT_INDICEG='+InttoStr(Cledoc.Indice)+' AND '+
             'GPT_FRAISREPARTIS="-" AND GPT_RETENUEDIVERSE="-"';
  QQ := OpenSql (Requete,True,-1,'',true);
  if not QQ.eof then
  begin
    TOBMere.PutValue ('TOTMARCHE',TOBMere.GetValue ('TOTMARCHE') - QQ.findField('TOTPORT').asFloat );  // Total du Marché
  end;
  ferme (QQ);
end;

procedure TImprPieceViaTOB.SetMontantMarche(TOBmere,LaLigne : TOB);
var RefPiece : string;
    Cledoc : r_cledoc;
    TOBDetDevis : TOB;
    requete : string;
    Indice : integer;
    QQ : TQuery;
    TotalHt,Remise,escompte,TotalTTC : double;
begin
  TOBDetDevis := TOB.Create ('LES DEVIS FAC',nil,-1);
  RefPiece := GetDevisFromprepaFac (LaLigne);
  (*
  DecodeRefPiece (LaLigne.getValue('GL_PIECEORIGINE'),cledoc);
  RefPiece := EncoderefSel (cledoc);
  *)
  Requete := 'SELECT BMF_NATUREPIECEG,BMF_SOUCHE,BMF_NUMERO,BMF_INDICEG FROM BTMEMOFACTURE '+
             'WHERE BMF_DEVISPRINC=(SELECT BMF_DEVISPRINC FROM BTMEMOFACTURE WHERE BMF_DEVIS="'+RefPiece+'")';
  //
  TRY
    TOBDETDEVIS.LoadDetailDBFromSQL ('BTMEMOFACTURE',requete,false);
    for Indice := 0 to TOBDetDevis.detail.count -1 do
    begin
      Requete := 'SELECT GP_TOTALHTDEV,GP_TOTALTTCDEV,GP_TOTALREMISEDEV,GP_TOTALESCDEV FROM PIECE WHERE GP_NATUREPIECEG="'+
                 TOBDetDevis.detail[Indice].getValue('BMF_NATUREPIECEG')+'" AND '+
                 'GP_SOUCHE="'+TOBDetDevis.detail[Indice].getValue('BMF_SOUCHE')+'" AND '+
                 'GP_NUMERO='+InttoStr(TOBDetDevis.detail[Indice].getValue('BMF_NUMERO'))+' AND '+
                 'GP_INDICEG='+InttoStr(TOBDetDevis.detail[Indice].getValue('BMF_INDICEG'));
      QQ := OpenSql (Requete,True,-1,'',true);
      if not QQ.eof then
      begin
        TotalHt := QQ.findField('GP_TOTALHTDEV').asfloat;
        TotalTTC := QQ.findField('GP_TOTALTTCDEV').asfloat;
        Remise := QQ.findField('GP_TOTALREMISEDEV').asfloat;
        Escompte := QQ.findField('GP_TOTALESCDEV').asfloat;
        //
        TOBMere.PutValue ('TOTMARCHE',TOBMere.GetValue ('TOTMARCHE')+ TotalHT+Remise+Escompte );  // Total du Marché brut
        TOBMere.PutValue ('TOTMARCHENET',TOBMere.GetValue ('TOTMARCHENET')+TotalHt );  // Total du Marché Net
        TOBMere.PutValue ('TOTMARCHETTC',TOBMere.GetValue ('TOTMARCHETTC')+TotalTTC );  // Total TTC du Marché
      end;
      ferme (QQ);
      //
      Requete := 'SELECT SUM(GPT_TOTALHTDEV) AS TOTPORT FROM PIEDPORT WHERE GPT_NATUREPIECEG="'+
                 TOBDetDevis.detail[Indice].getValue('BMF_NATUREPIECEG')+'" AND '+
                 'GPT_SOUCHE="'+TOBDetDevis.detail[Indice].getValue('BMF_SOUCHE')+'" AND '+
                 'GPT_NUMERO='+InttoStr(TOBDetDevis.detail[Indice].getValue('BMF_NUMERO'))+' AND '+
                 'GPT_INDICEG='+InttoStr(TOBDetDevis.detail[Indice].getValue('BMF_INDICEG'))+' AND '+
                 'GPT_FRAISREPARTIS="-" AND GPT_RETENUEDIVERSE="-"';
      QQ := OpenSql (Requete,True,-1,'',true);
      if not QQ.eof then
      begin
        TOBMere.PutValue ('TOTMARCHE',TOBMere.GetValue ('TOTMARCHE') - QQ.findField('TOTPORT').asFloat );  // Total du Marché
      end;
      ferme (QQ);
    end;
  FINALLY
    TOBDETDEVIS.free;
  END;
end;

procedure TImprPieceViaTOB.AlimenteTableauSituations (ThePiece : TOB);
var Indice : integer;
    TOBS,TOBL : TOB;
    TheOrdre : integer;
    MaxC : integer;
begin
	if (fIsDgd) and (fisAVA) then MaxC := 0 else MaxC := 1;
  if fTOBSituations.detail.count <= MaxC then exit;
  for Indice := 0 to fTOBSituations.Detail.count - (1+MaxC) do
  begin
    TOBS := fTOBSituations.detail[Indice];
    TOBL := InsereTOB ('LIGNE',ThePiece);
    TheOrdre := TOBL.GetValue('GL_NUMORDRE');
    //
    DefiniLigneSituation (TOBL,TOBS);
    //
    TOBL.PutValue('GL_NUMORDRE',TheOrdre); // retablissement
  	if Indice = 0 then TOBL.PutValue ('DEBUTRECAPSIT','X');
  	if Indice = fTOBSituations.Detail.count -2 then TOBL.PutValue ('FINRECAPSIT','X');
  end;

end;

procedure TImprPieceViaTOB.SetMontantregleSituations;
var indice  : integer;
    TOBS    : TOB;
    TOBAC   : TOB;
    req     : string;
    QQ      : TQuery;
    TotalAc : Double;
    TotalRg : Double;
    i       : Integer;
begin

  for Indice := 0 to fTOBSituations.detail.count - 1 do
  begin
    TOBS := fTOBSituations.detail[Indice];
    //
    TOBS.AddChampSupValeur ('ACOMPTESIT',0);
    TOBS.AddChampSupValeur('REGLEMENTSIT',0);
    //
    TotalAC := 0;
    TotalRG := 0;
    //
    req := 'SELECT GAC_MONTANTDEV, GAC_ISREGLEMENT FROM ACOMPTES WHERE '+
            'GAC_NUMERO='+inttoStr(TOBS.getValue('BST_NUMEROFAC'))+' AND '+
            'GAC_NATUREPIECEG="'+TOBS.getValue('BST_NATUREPIECE')+'" AND '+
            'GAC_SOUCHE="'+TOBS.getValue('BST_SOUCHE')+'"';
    //
    TOBAC := TOB.Create('ACOMPTE',nil, -1);
    TOBAC.LoadDetailDBFromSQL('ACOMPTES', Req, False);
    //
    FOR I := 0 TO TOBAC.Detail.count - 1 do
    begin
     if TOBAC.Detail[i].GetValue('GAC_ISREGLEMENT') = 'X' then
        TotalRg := TotalRg + TOBAC.Detail[i].GetDouble('GAC_MONTANTDEV')
     else
        TotalAC := TotalAC + TOBAC.Detail[i].GetDouble('GAC_MONTANTDEV');
    end;
    //
    ferme (QQ);
    //
    TOBS.putValue('ACOMPTESIT',TotalAc);
    TOBS.putValue('REGLEMENTSIT',TotalRg);
    //
    FreeAndNil(TOBAC);
  end;

end;

procedure TImprPieceViaTOB.DefiniLigneSituation (TOBMere,TOBSituation : TOB);
var QQ : TQuery;
    req : string;
    Indice : integer;
begin
  TOBMere.PutValue('BST_NATUREPIECE', TOBSituation.GetValue('BST_NATUREPIECE'));
  TOBMere.PutValue('BST_NUMEROFAC',   TOBSituation.GetValue('BST_NUMEROFAC'));
  TOBMere.PutValue('BST_NUMEROSIT',   TOBSituation.GetValue('BST_NUMEROSIT'));
  TOBMere.PutValue('BST_DATESIT',     TOBSituation.GetValue('BST_DATESIT'));
  TOBMere.PutValue('BST_MONTANTHT',   TOBSituation.GetValue('BST_MONTANTHT'));
  TOBMere.PutValue('BST_MONTANTTTC',  TOBSituation.GetValue('BST_MONTANTTTC'));
  TOBMere.PutValue('BST_MONTANTREGL', TOBSituation.GetValue('BST_MONTANTREGL'));
  TOBMere.PutValue('ACOMPTESIT',      TOBSituation.GetValue('ACOMPTESIT'));
  TOBMere.PutValue('REGLEMENTSIT',    TOBSituation.GetValue('REGLEMENTSIT'));
  TOBMere.PutValue('LIGNERECAPSIT',   'X');
end;

procedure TImprPieceViaTOB.AjouteChampsSituation (TOBMere: TOB);
var Itable,Indice,IndiceTOB : integer;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
	NomChamps,TypeChamps : string;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  Table := Mcd.GetTable(Mcd.PrefixeToTable('BST'));
  FieldList := Table.Fields;
  FieldList.Reset();

  While FieldList.MoveNext do
  begin
		NomChamps := (FieldList.Current as IFieldCOM).name;
		TypeChamps := (FieldList.Current as IFieldCOM).Tipe;
    if not TOBMere.FieldExists (NomChamps) then
    begin
      TOBMere.AddChampSup (NomChamps,false);
      IndiceTOB := TOBMere.GetNumChamp (NomChamps);
      InitChamps (TOBmere,NomChamps,TypeChamps);
    end;
  end;
end;

procedure TImprPieceViaTOB.InitChamps(TOBL : TOB; NomChamps,TypeChamps : string);
var typeC : string;
begin
  if (TypeChamps = 'INTEGER') or (TypeChamps = 'SMALLINT') then TOBL.PutValue(NomChamps,0)
  else if (TypeChamps = 'DOUBLE') or (TypeChamps = 'EXTENDED') or (typeC = 'DOUBLE') then TOBL.Putvalue(NomChamps,0)
  else if (TypeChamps = 'DATE') then TOBL.Putvalue(NomChamps,iDate1900)
  else if (TypeChamps = 'BLOB') or (TypeChamps = 'DATA') then TOBL.Putvalue(NomChamps,'')
  else if (TypeChamps = 'BOOLEAN') then TOBL.Putvalue(NomChamps,'-')
  else TOBL.Putvalue(NomChamps,'')
end;


procedure TImprPieceViaTOB.AjouteChampsAcomptesCpta (TOBMere : TOB);
var Value : string;
begin
  if not TOBMere.FieldExists ('ACOMPTECPT') then
  begin
    TOBMere.AddChampSupValeur  ('ACOMPTECPT','-');
  end;
  if (GetParamSoc('SO_BTCOMPTAREGL')) then Value := 'X' else Value := '-';
  TOBMere.Putvalue  ('ACOMPTECPT',Value);
end;

procedure TImprPieceViaTOB.DefiniRecapTotalSit (ThePiece : TOB);
begin

  SetLignesRecap (ThePiece);
end;

procedure TImprPieceViaTOB.ConstitueRD (ThePiece,TOBFDDET : TOB; var NbFDHT : integer);

  function FindBaseDoc (TOBBASESPR : TOb;OneCledoc : r_cledoc) : TOB;
  var thePiece : string;
  begin
    ThePiece := EncodeRefPiece(OneCledoc);
    result := TOBBasesRG.findfirst(['CLEDOC'],[ThePiece],false);
  end;

  procedure AjouteBaseDoc (TOBBPR : TOB;OneCledoc : r_cledoc);
  var thePiece : string;
      QQ : TQuery;
  begin
    ThePiece := EncodeRefPiece(OneCledoc);
    TOBBPR.AddChampSupValeur('CLEDOC',ThePiece);
    TOBBPR.AddChampSupValeur('TOTALHTDEV',0);
    if favancement then
    begin
      QQ := OpenSql ('SELECT PIEDBASE.*,GP_TOTALHTDEV '+
                     'FROM PIEDBASE '+
                     'LEFT JOIN PIECE ON '+
                     'GPB_NATUREPIECEG=GP_NATUREPIECEG AND '+
                     'GPB_SOUCHE=GP_SOUCHE AND '+
                     'GPB_NUMERO=GP_NUMERO AND '+
                     'GPB_INDICEG=GP_INDICEG '+
                     'WHERE '+WherePiece(OneCledoc,ttdPiedBase,false),true,-1,'',true);
    end
    else
    begin
      QQ := OpenSql ('SELECT PIEDBASE.*,GP_TOTALHTDEV '+
                     'FROM PIEDBASE '+
                     'LEFT JOIN PIECE ON '+
                     'GPB_NATUREPIECEG=GP_NATUREPIECEG AND '+
                     'GPB_SOUCHE=GP_SOUCHE AND '+
                     'GPB_NUMERO=GP_NUMERO AND '+
                     'GPB_INDICEG=GP_INDICEG '+
                     'WHERE '+WherePiece(OneCledoc,ttdPiedBase,false),true,-1,'',true);
    end;

    if not QQ.eof then
    begin
      TOBBPR.LoadDetailDB('PIEDBASE','','',QQ,false);
      if TOBBPR.detail.count > 0 then
      begin
        TOBBPR.SetDouble('TOTALHTDEV',TOBBPR.Detail[0].getDouble('GP_TOTALHTDEV'));
      end;
    end;
    ferme (QQ);  
  end;

  procedure ConstitueBasesRD (TOBRD,TOBBPR : TOB);
  var ratio : double;
  begin
    Ratio := TOBRD.getDouble('GPT_TOTALHTDEV')/TOBBPR.GetDouble('TOTALHTDEV');
    RatioizeTVARetenue (TOBRD,TOBBPR,TOBRD,DEV,ratio);    
  end;

  
var TOBL,TOBF,TOBBASESPR,TOBBPR,TOBDET,TT,TB,TLB: TOB;
    II,JJ : Integer;
    req: string;
    QQ : TQuery;
    OneCledoc : r_cledoc;
    ratio : double;
begin
  for II := 0 to TOBPorcs.detail.count -1 do
  begin
    TOBPorcs.detail[II].Cleardetail;
  end;
  fillchar(OneCledoc,Sizeof(OneCledoc),#0);
  TOBBASESPR := TOB.Create ('LES BASES PREC',nil,-1);
  TOBDET := TOB.Create ('LES RD',nil,-1);
  // ATTENTION -- requete ... Pour les retenues diverse
  if favancement then
  begin
    req := 'SELECT PIEDPORT.* FROM BSITUATIONS LEFT JOIN PIEDPORT ON '+
           'GPT_NATUREPIECEG=BST_NATUREPIECE AND '+
           'GPT_SOUCHE=BST_SOUCHE AND '+
           'GPT_NUMERO=BST_NUMEROFAC ';
    Req := Req + 'WHERE BST_SSAFFAIRE = "'+ThePiece.GetString('GP_AFFAIREDEVIS') + '"';
    Req := Req + 'AND BST_NUMEROSIT < '+ ThePiece.getString('NUMSIT');
    req := Req + 'AND GPT_RETENUEDIVERSE="X" AND BST_VIVANTE="X" ';
  end
  else
  begin
    req := 'SELECT PIEDPORT.* FROM PIECE LEFT JOIN PIEDPORT ON '+
           'GPT_NATUREPIECEG=GP_NATUREPIECEG AND '+
           'GPT_SOUCHE=GP_SOUCHE AND '+
           'GPT_NUMERO=GP_NUMERO ';
    Req := Req + 'WHERE GP_AFFAIREDEVIS ="'+ ThePiece.GetString('GP_AFFAIREDEVIS') + '"';
    Req := Req + '  AND GP_NATUREPIECEG="' + ThePiece.GetString('GP_NATUREPIECEG') + '"';
    Req := Req + '  AND GP_SOUCHE="'       + ThePiece.GetString('GP_SOUCHE')       + '"';
    Req := Req + '  AND GP_NUMERO='        + ThePiece.GetString('GP_NUMERO');
    Req := Req + '  AND GP_INDICEG='       + ThePiece.GetString('GP_INDICEG');
    req := Req + '  AND GPT_RETENUEDIVERSE="X" AND GP_VIVANTE="X" ';
  end;
  //
  QQ := OpenSQL(req,True,-1,'',true);
  if not QQ.eof then
  begin
    TOBDET.LoadDetailDB('PIEDPORT','','',QQ,false);
  end;
  ferme (QQ);
  //
  for II := 0 to TOBDET.detail.count -1 do
  begin
    TT := TOBDET.detail[II];
    if (Pos(TT.GetString('GPT_TYPEPORT'),'HT;MI;MT')>0) then
    begin
      OneCledoc.NaturePiece := TT.GetString('GPT_NATUREPIECEG') ;
      OneCledoc.Souche := TT.GetString('GPT_SOUCHE');
      OneCledoc.NumeroPiece := TT.GetInteger('GPT_NUMERO');
      OneCledoc.Indice :=TT.GetInteger('GPT_INDICEG');
      TOBBPR :=  FindBaseDoc (TOBBASESPR,OneCledoc);
      if TOBBPR = nil then
      begin
        TOBBPR := TOB.Create('UNE BASE',TOBBASESPR,-1);
        AjoutebaseDoc (TOBBPR,OneCledoc);
      end;
      ConstitueBasesRD (TT,TOBBPR);
    end;
  end;
  //
  for II := 0 to TOBDet.detail.count -1 do
  begin
    TT := TOBDet.detail[II];
    TOBL :=  TOBFDDET.findfirst(['CODE'],[TT.getString('GPT_CODEPORT')],false);
    if TOBL = nil then
    begin
      inc(NbFDHT);
      TOBL := TOB.Create ('UNE RET DIV',TOBFDDET,-1);
      TOBL.AddChampSupValeur('CODE',TT.getString('GPT_CODEPORT'));
      TOBL.AddChampSupValeur('TYPE',TT.getString('GPT_TYPEPORT'));
      TOBL.AddChampSupValeur('LIBELLE',TT.getString('GPT_LIBELLE'));
      TOBL.AddChampSupValeur('FAMILLETAXE1','');
      TOBL.AddChampSupValeur('MONTANTHTPRE',0);
      TOBL.AddChampSupValeur('MONTANTTTCPRE',0);
      TOBL.AddChampSupValeur('TVAPRE',0);
      TOBL.AddChampSupValeur('MONTANTHTSIT',0);
      TOBL.AddChampSupValeur('MONTANTTTCSIT',0);
      TOBL.AddChampSupValeur('TVASIT',0);
    end;
    if favancement then TOBL.SetDouble('MONTANTHTPRE',TOBL.GetDouble('MONTANTHTPRE')+TT.getDouble('GPT_TOTALHTDEV'));
    if favancement then TOBL.SetDouble('MONTANTTTCPRE',TOBL.GetDouble('MONTANTTTCPRE')+TT.getDouble('GPT_TOTALTTCDEV'));
    if favancement then TOBL.SetDouble('TVAPRE',TOBL.GetDouble('MONTANTTTCPRE')-TOBL.GetDouble('MONTANTHTPRE'));
    // gestion de la tva associée
    for JJ := 0 to TT.detail.count -1 do
    begin
      TB := TT.detail[JJ];
      TLB := TOBL.findfirst(['FAMILLETAXE'],[TB.getString('GPB_FAMILLETAXE')],true);
      if TLB = nil then
      begin
        TLB := TOB.create ('UNE TVA',TOBL,-1);
        TLB.AddChampSupValeur ('FAMILLETAXE',TB.getValue('GPB_FAMILLETAXE'));
        TLB.AddChampSupValeur ('TAUXTAXE',TB.getValue('GPB_TAUXTAXE'));
        TLB.AddChampSupValeur ('BASETAXE',0);
        TLB.AddChampSupValeur ('BASETAXEPRE',0);
        TLB.AddChampSupValeur ('BASETAXECUM',0);
        TLB.AddChampSupValeur ('VALEURTAXEPRE',0);
        TLB.AddChampSupValeur ('VALEURTAXESIT',0);
      end;
      TLB.SetDouble('BASETAXECUM',TLB.GetDouble('BASETAXECUM')+TB.getDouble('GPB_BASETAXE'));
      if favancement then TLB.SetDouble('BASETAXEPRE',TLB.GetDouble('BASETAXEPRE')+TB.getDouble('GPB_BASETAXE'));
      if favancement then TLB.SetDouble('VALEURTAXEPRE',TLB.GetDouble('VALEURTAXEPRE')+TB.getDouble('GPB_VALEURTAXE'));
    end;
  end;
  //
  For ii := 0 to fTOBPorcs.detail.count -1 do
  begin
    TOBF := fTOBPorcs.detail[II];
    if TOBF.GetString('GPT_FRAISREPARTIS')<>'-' then Continue;
    if not TOBF.GetBoolean('GPT_RETENUEDIVERSE') then Continue;
    //
    Ratio := TOBF.GetDouble('GPT_TOTALHTDEV') / TOBPiece.geTDouble('GP_TOTALHTDEV');  // % du montant global
    RatioizeTVARetenue(TOBF,TOBBases,TOBF,DEV,ratio);
    //
    TOBL := TOBFDDET.FindFirst(['CODE'],[TOBF.getString('GPT_CODEPORT')],false);
    if TOBL = nil then
    begin
      TOBL := TOB.Create ('UNE RET DIV',TOBFDDET,-1);
      TOBL.AddChampSupValeur('CODE',TOBF.GetString('GPT_CODEPORT'));
      TOBL.AddChampSupValeur('TYPE',TOBF.GetString('GPT_TYPEPORT'));
      TOBL.AddChampSupValeur('LIBELLE',TOBF.GetString('GPT_LIBELLE'));
      TOBL.AddChampSupValeur('FAMILLETAXE1',TOBF.GetString('GPT_FAMILLETAXE1'));
      TOBL.AddChampSupValeur('MONTANTHTPRE',0);
      TOBL.AddChampSupValeur('MONTANTTTCPRE',0);
      TOBL.AddChampSupValeur('TVAPRE',0);
      TOBL.AddChampSupValeur('MONTANTHTSIT',0);
      TOBL.AddChampSupValeur('MONTANTTTCSIT',0);
      TOBL.AddChampSupValeur('TVASIT',0);
    end;
    TOBL.SetDouble('MONTANTHTSIT',TOBL.getDouble('MONTANTHTSIT')+TOBF.getDouble('GPT_TOTALHTDEV'));
    TOBL.SetDouble('MONTANTTTCSIT',TOBL.GetDouble('MONTANTTTCSIT')+TOBF.getDouble('GPT_TOTALTTCDEV'));
    TOBL.SetDouble('TVASIT',TOBL.GetDouble('MONTANTTTCSIT')-TOBL.GetDouble('MONTANTHTSIT'));
    // gestion de la tva associée
    for JJ := 0 to TOBF.detail.count -1 do
    begin
      TB := TOBF.detail[JJ];
      if TB.GetValue('GPB_TYPEINTERV') = 'Y00' then Continue;
      //
      if fAvancement then
        TLB := TOBL.findfirst(['FAMILLETAXE'],[TB.getString('GPB_FAMILLETAXE')],true)
      else
        TLB := TOBL.findfirst(['FAMILLETAXE'],[TOBF.getString('GPT_FAMILLETAXE1')],true);
      if TLB = nil then
      begin
        TLB := TOB.create ('UNE TVA',TOBL,-1);
        TLB.AddChampSupValeur ('FAMILLETAXE',TB.getValue('GPB_FAMILLETAXE'));
        TLB.AddChampSupValeur ('TAUXTAXE',TB.getValue('GPB_TAUXTAXE'));
        TLB.AddChampSupValeur ('BASETAXE',0);
        TLB.AddChampSupValeur ('BASETAXEPRE',0);
        TLB.AddChampSupValeur ('BASETAXECUM',0);
        TLB.AddChampSupValeur ('VALEURTAXEPRE',0);
        TLB.AddChampSupValeur ('VALEURTAXESIT',0);
      end;
      TLB.SetDouble('BASETAXE',TLB.GetDouble('BASETAXE')+TB.getDouble('GPB_BASETAXE'));
      TLB.SetDouble('BASETAXECUM',TLB.GetDouble('BASETAXECUM')+TB.getDouble('GPB_BASETAXE'));
      TLB.SetDouble('VALEURTAXESIT',TLB.GetDouble('VALEURTAXESIT')+TB.getDouble('GPB_VALEURTAXE'));
    end;
  end;
  TOBBASESPR.free;
  TOBDET.free;
end;

procedure TImprPieceViaTOB.ConstitueFraisDetaille (ThePiece : TOB; TOBFD : TOB);
var TOBL,TOBF : TOB;
    II : Integer;
    req: string;
    QQ : TQuery;
begin
  // ATTENTION -- requete ... Pour les frais et ports detaillés
  if fAvancement then
  Begin
    req := 'SELECT '+
           'GPT_TYPEPORT,'+
           'GPT_LIBELLE,'+
           'GPT_CODEPORT,'+
           'SUM(GPT_TOTALHTDEV) AS CUMHT ,'+
           'SUM(GPT_TOTALTTCDEV) AS CUMTTC,'+
           '(SELECT GPO_FRAISREPARTIS FROM PORT WHERE GPO_CODEPORT = GPT_CODEPORT) AS FRAISREPARTIS,'+
           'GPT_FAMILLETAXE1 '+
           'FROM BSITUATIONS '+
           'LEFT JOIN PIEDPORT ON '+
           'GPT_NATUREPIECEG=BST_NATUREPIECE AND '+
           'GPT_SOUCHE=BST_SOUCHE AND '+
           'GPT_NUMERO=BST_NUMEROFAC '+
           'WHERE '+
           'BST_SSAFFAIRE = "'+ThePiece.GetString('GP_AFFAIREDEVIS')+'" AND '+
           'BST_NUMEROSIT < ( SELECT S.BST_NUMEROSIT FROM BSITUATIONS S WHERE '+
           'S.BST_NATUREPIECE="'+ThePiece.getString('GP_NATUREPIECEG')+'" AND '+
           'S.BST_SOUCHE="'+ThePiece.GetString('GP_SOUCHE')+'" AND '+
           'S.BST_NUMEROFAC='+ThePiece.GetString('GP_NUMERO')+') AND '+
           'GPT_RETENUEDIVERSE<>"X" AND '+
           'BST_VIVANTE="X" '+
           'GROUP BY GPT_TYPEPORT,GPT_CODEPORT,GPT_LIBELLE,GPT_FAMILLETAXE1';
  end
  else
  begin
    req := 'SELECT '+
           'GPT_TYPEPORT,'+
           'GPT_LIBELLE,'+
           'GPT_CODEPORT,'+
           'SUM(GPT_TOTALHTDEV) AS CUMHT ,'+
           'SUM(GPT_TOTALTTCDEV) AS CUMTTC,'+
           '(SELECT GPO_FRAISREPARTIS FROM PORT WHERE GPO_CODEPORT = GPT_CODEPORT) AS FRAISREPARTIS,'+
           'GPT_FAMILLETAXE1 '+
           'FROM PIECE '+
           'LEFT JOIN PIEDPORT ON '+
           'GPT_NATUREPIECEG=GP_NATUREPIECEG AND '+
           'GPT_SOUCHE=GP_SOUCHE AND '+
           'GPT_NUMERO=GP_NUMERO '+
           'WHERE '+
           'GP_AFFAIRE = "'   + ThePiece.GetString('GP_AFFAIREDEVIS') + '" AND ' +
           'GP_NATUREPIECEG="'+ ThePiece.getString('GP_NATUREPIECEG') + '" AND ' +
           'GP_SOUCHE="'      + ThePiece.GetString('GP_SOUCHE')       + '" AND ' +
           'GP_NUMERO='       + ThePiece.GetString('GP_NUMERO')       + '  AND ' +
           'GPT_RETENUEDIVERSE<>"X" AND GP_VIVANTE="X" ' +
           'GROUP BY GPT_TYPEPORT,GPT_CODEPORT,GPT_LIBELLE,GPT_FAMILLETAXE1';
  end;
  //
  QQ := OpenSQL(req,True,-1,'',true);
  QQ.first;
  while not QQ.eof do
  begin
    if QQ.fields[5].AsString <> 'X' then
    begin
      TOBL := TOB.Create ('UN PORT',TOBFD,-1);
      TOBL.AddChampSupValeur('CODE',QQ.fields[2].asstring);
      TOBL.AddChampSupValeur('TYPE',QQ.fields[0].asstring);
      TOBL.AddChampSupValeur('LIBELLE',QQ.fields[1].asstring);
      TOBL.AddChampSupValeur('FAMILLETAXE1',QQ.fields[6].asstring);
      if fAvancement then TOBL.AddChampSupValeur('MONTANTHTPRE',QQ.fields[3].asfloat);
      if fAvancement then TOBL.AddChampSupValeur('MONTANTTTCPRE',QQ.fields[4].asfloat);
      if fAvancement then TOBL.AddChampSupValeur('TVAPRE',TOBL.GetDouble('MONTANTTTCPRE')-TOBL.GetDouble('MONTANTHTPRE'));
      TOBL.AddChampSupValeur('MONTANTHTSIT',0);
      TOBL.AddChampSupValeur('MONTANTTTCSIT',0);
      TOBL.AddChampSupValeur('TVASIT',0);
    end;
    QQ.next;
  end;
  ferme (QQ);
  For ii := 0 to fTOBPorcs.detail.count -1 do
  begin
    TOBF := fTOBPorcs.detail[II];
    if TOBF.GetString('GPT_FRAISREPARTIS')<>'-' then Continue;
    if TOBF.GetBoolean('GPT_RETENUEDIVERSE') then Continue;
    //
    TOBL := TOBFD.FindFirst(['CODE','LIBELLE','FAMILLETAXE1'],[TOBF.getString('GPT_CODEPORT'),TOBF.getString('GPT_LIBELLE'),TOBF.getString('GPT_FAMILLETAXE1')],True);
    if TOBL = nil then
    begin
      TOBL := TOB.Create ('UN PORT',TOBFD,-1);
      TOBL.AddChampSupValeur('CODE',TOBF.GetString('GPT_CODEPORT'));
      TOBL.AddChampSupValeur('TYPE',TOBF.GetString('GPT_TYPEPORT'));
      TOBL.AddChampSupValeur('LIBELLE',TOBF.GetString('GPT_LIBELLE'));
      TOBL.AddChampSupValeur('FAMILLETAXE1',TOBF.GetString('GPT_FAMILLETAXE1'));
      if fAvancement then TOBL.AddChampSupValeur('MONTANTHTPRE',0);
      if fAvancement then TOBL.AddChampSupValeur('MONTANTTTCPRE',0);
      if fAvancement then TOBL.AddChampSupValeur('TVAPRE',0);
      TOBL.AddChampSupValeur('MONTANTHTSIT',0);
      TOBL.AddChampSupValeur('MONTANTTTCSIT',0);
      TOBL.AddChampSupValeur('TVASIT',0);
    end;
    TOBL.SetDouble('MONTANTHTSIT',TOBL.getDouble('MONTANTHTSIT')+TOBF.getDouble('GPT_TOTALHTDEV'));
    TOBL.SetDouble('MONTANTTTCSIT',TOBL.GetDouble('MONTANTTTCSIT')+TOBF.getDouble('GPT_TOTALTTCDEV'));
    TOBL.SetDouble('TVASIT',TOBL.GetDouble('MONTANTTTCSIT')-TOBL.GetDouble('MONTANTHTSIT'));
  end;
end;

procedure TImprPieceViaTOB.SetLignesRecap (ThePiece : TOB);

  procedure   GetInfosHtRD (TOBRDDDET : TOB; var MTRDPREHT,MTRDSITHT : double);
  var II : integer;
  begin
    for II := 0 to TOBRDDDET.detail.count - 1 do
    begin
      if TOBRDDDET.detail[II].GetString('TYPE')<> 'HT' THEN continue;
      MTRDPREHT := MTRDPREHT + TOBRDDDET.detail[II].GetDouble('MONTANTHTPRE');
      MTRDSITHT := MTRDSITHT + TOBRDDDET.detail[II].GetDouble('MONTANTHTSIT');
    end;
  end;


var RemiseCum,EscompteCum,PortCum,CumPortSit : double;
    QQ : Tquery;
    req : string;
    Indice : integer;
    TOBTVAPre,TOBRGPre,TOBFDetaille : TOB;
    CumSitTVA,CumPreTVA,CumRGSit,CumRgPre,ReliquatRGSit,reliquatRGPre,TVARGSIt,TVARGPre,TTCPre,TTCSit,AcomptesSit,AcomptesPre,Timbres : double;
    MTRDSIT,MTRDPRE,TVARDSIT,TVARDPRE,MTRDSITTTC,MTRDPRETTC,MTRDSITHT,MTRDPREHT,MtHtPortTTCPre,MtHTPortTTC,MtPortTTCPre,MtPortTTC : double;
    NbrHT,NbFDHt,NbrTTC,NBrPOrts : Integer;
    TypePort : string;
    Repartition     : boolean;
    TOBPIECEREPART,TOBRDDET,TOBTVAPRETTC  : TOB;
begin
  NbFDHt := 0;
  Repartition     :=  GetParamSocSecur('SO_PIECEREPART', False);

  TOBTVAPre := TOB.Create ('LES TVA PREC',nil,-1);
  TOBRGPre := TOB.Create ('LES RG PREC',nil,-1);
  TOBFDetaille := TOB.Create ('LES FRAIS DETAILLE',nil,-1);
  TOBRDDET        := TOB.Create ('LES RETENUE DIV',nil,-1);
  TOBTVAPRETTC    := TOB.Create ('LES TVA SUR PORTS TTC',nil,-1);

  ConstitueFraisDetaille (ThePiece,TOBFDetaille);
  ConstitueRD (ThePiece,TOBRDDET,NbFDHt); // continue les retenues diverses ainsi que leurs tva associés (cumulés)

  MTRDSIT := 0;
  MTRDPRE := 0;
  TVARDSIT := 0;
  TVARDPRE := 0;
  // --- Montant des retenues divers en HT a rajouter au HT (gestion....)
  MTRDSITHT := 0;
  MTRDPREHT := 0;
  GetInfosHtRD (TOBRDDET,MTRDPREHT,MTRDSITHT);
  //

  CumPortSit := 0; NbrHT := 0; NBrTTC := 0;
  for Indice := 0 to fTOBPorcs.detail.count -1 do
  begin
    if (fTOBPorcs.detail[Indice].getValue('GPT_FRAISREPARTIS')<>'X') and (not fTOBPorcs.detail[Indice].getBoolean('GPT_RETENUEDIVERSE'))  then
    begin
      CumPortSit := CumPortSit + fTOBPorcs.detail[Indice].getValue('GPT_TOTALHTDEV');
      TypePort := fTOBPorcs.detail[Indice].getValue('GPT_TYPEPORT');
    end;
    if (TypePort='HT') or (TypePort='MI') or (TypePort='MT') then Inc(NbrHT);
    if (fTOBPorcs.detail[Indice].getBoolean('GPT_RETENUEDIVERSE')) and (TypePort='PT') or (TypePort='MIC') or (TypePort='MTC') then Inc(NbrTTC);
  end;
  //

  if Favancement then
  Begin
    Req := 'SELECT SUM(GP_TOTALREMISEDEV) AS REMCUM, SUM(GP_TOTALESCDEV) AS ESCCUM FROM PIECE, BSITUATIONS ';
    Req := Req + ' WHERE BST_SSAFFAIRE  ="' + ThePiece.GetValue('GP_AFFAIREDEVIS')  +'" ';
    Req := Req + '   AND BST_NATUREPIECE="' + ThePiece.GetValue('GP_NATUREPIECEG')  +'" ';
    Req := Req + '   AND BST_SOUCHE     ="' + ThePiece.GetValue('GP_SOUCHE')        +'" ';
    Req := Req + '   AND BST_VIVANTE    ="X"';
    Req := Req + '   AND BST_NUMEROSIT < '+ ThePiece.getString('NUMSIT');
    Req := Req + '   AND GP_NATUREPIECEG=BST_NATUREPIECE AND GP_SOUCHE=BST_SOUCHE AND GP_NUMERO=BST_NUMEROFAC';
  end
  else
  begin
    Req := 'SELECT SUM(GP_TOTALREMISEDEV) AS REMCUM, SUM(GP_TOTALESCDEV) AS ESCCUM FROM PIECE ';
    Req := Req + ' WHERE GP_AFFAIREDEVIS  ="' + ThePiece.GetValue('GP_AFFAIREDEVIS')  +'" ';
    Req := Req + '   AND GP_NATUREPIECEG="'   + ThePiece.GetValue('GP_NATUREPIECEG')  +'" ';
    Req := Req + '   AND GP_SOUCHE     ="'    + ThePiece.GetValue('GP_SOUCHE')        +'" ';
    Req := Req + '   AND GP_NUMERO     = '    + InttoStr(ThePiece.GetValue('GP_NUMERO'));
    Req := Req + '   AND GP_VIVANTE    ="X"';
  end;
  //
  QQ := OpenSql (Req,true,-1,'',true);
  if not QQ.eof then
  begin
    RemiseCum := QQ.findField('REMCUM').asFloat;
    EscompteCum := QQ.findField('ESCCUM').asFloat;
  end;
  ferme (QQ);
  //
  PortCum := 0;
  For Indice := 0 to TOBFDetaille.detail.count -1 do
  begin
    PortCum := PortCum + TOBFDetaille.Detail [Indice].GetDouble('MONTANTHTPRE')+ TOBFDetaille.Detail [Indice].GetDouble('MONTANTHTSIT');
  end;
  //
  ConstitueTOBTVA (ThePiece,TOBTVAPre,TOBTVAPRETTC,MtHtPortTTCPre,MtHTPortTTC,MtPortTTCPre,MtPortTTC);
  ConstitueTOBRGPrec (ThePiece,TOBRGPre);
  if Repartition then
    if ThePiece.GetValue('BPD_IMPDETAVE')='X' then
    begin
      DefiniLignePieceRepart(ThePiece,EscompteCum,RemiseCum,PortCum,CumPortSit,MTRDSITHT,MTRDPREHT);
    end
    else
      DefiniLigneMontantHt (ThePiece,EscompteCum,RemiseCum,PortCum,CumPortSit,MTRDSITHT,MTRDPREHT)
  else
    DefiniLigneMontantHt (ThePiece,EscompteCum,RemiseCum,PortCum,CumPortSit,MTRDSITHT,MTRDPREHT);
  //
  DefiniLigneRemise (ThePiece,RemiseCum);
  DefiniLigneEscompte (ThePiece,EscompteCum);
//  DefiniLignePortFrais (ThePiece,PortCum,CumPortSit);
  DefiniLignePortFraisDet (ThePiece,TOBFDetaille,'HT',NbrTTc);
  // Gestion des retenues diverses
  DefiniLigneRDHT (ThePiece,TOBRDDET,MTRDSIT,MTRDPRE);
  MTRDSIT :=0; MTRDPRE := 0; // ca four le souk
  if (EscompteCum<>0) or (RemiseCum<>0) or (nbrHt<>0) or (MTRDPREHT<>0) or (MTRDSITHT<>0) then
  begin
    DefiniLigneTotalHt(ThePiece);
    DefiniLigneTotalHtDet(ThePiece,TOBFDetaille);
  end;

  DefiniLignesTVA (ThePiece,TOBTVAPre,TOBTVAPRETTC,TOBRDDET,CumSitTVA,CumPreTVA);
  if fAvancement then DefiniLignesTVADet(ThePiece,TOBTVAPre,TOBFDetaille,TOBRDDET);
  DefiniLigneRGHT (ThePiece,TOBRGPRE,CumRGSit,CumRgPre,ReliquatRGSit, reliquatRGPre);
  //
  DefiniLigneTVARG(ThePiece,TOBRGPRE,TVARGSIt,TVARGPre);
  // Gestion des retenues diverses
  DefiniLigneTVARD (ThePiece,TOBRDDET,TVARDSIT,TVARDPRE);
  //
  DefiniLigneTOTALTTCInterm (ThePiece,TOBFDetaille,CumSitTVA,CumPreTVA,CumRGSit,CumRgPre,TVARGSIt,TVARGPre,MTRDSIT,MTRDPRE,TVARDSIT,TVARDPRE,MtHtPortTTCPre,MtHTPortTTC,TTCSit,TTCPre);
  DefiniLigneTOTALTTC (ThePiece,TOBfDetaille,CumSitTVA,CumPreTVA,CumRGSit,CumRgPre,TVARGSIt,TVARGPre,MTRDSIT,MTRDPRE,TVARDSIT,TVARDPRE,MtHtPortTTCPre,MtHTPortTTC,TTCSit,TTCPre);
  DefiniLigneRGTTC(ThePiece,TOBRGPRE,ReliquatRgSit,ReliquatRgpre);
  DefiniLignePortFraisDet (ThePiece,TOBFDetaille,'TTC',NBrPOrts);
  // Gestion des retenues diverses
  DefiniLigneRDTTC (ThePiece,TOBRDDET,MTRDSITTTC,MTRDPRETTC);
  //
  if (MTRDSITTTC<>0) or (MTRDPRETTC<>0) or (NBrPOrts<>0) then NbrTTC := 1;
  if (TOBtimbres.detail.count > 0) then
  begin
    Timbres := DefiniLigneTimbres(ThePiece,TOBTimbres);
  end;

  DefiniLigneACOMPTES(ThePiece,AcomptesSit,AcomptesPre);
  ThePiece.AddChampSupValeur ('TOTALNET', TTCSit); // défini systématiquement et initialisé au TTC car utilisé si solde sur situations précédentes
  if (Arrondi(Abs(AcomptesSit),2)<>0) or (Arrondi(Abs(AcomptesPre),2)<>0) or
     (Arrondi(Abs(ReliquatRGSit),2)<>0) or (Arrondi(Abs(reliquatRGPre),2)<>0) or
     (Timbres <> 0) or (NbrTTC>0)then
  BEGIN
    DefiniLigneTOTALNET (ThePiece,TTCSit,TTCPre,ReliquatRGSit,reliquatRGPre,AcomptesSit,AcomptesPre,Timbres,MTRDSITTTC,MTRDPRETTC,MtPortTTCPre,MtPortTTC);
  END;
  TOBTVAPre.free;
  TOBRGPre.free;
  TOBFDetaille.free;
  TOBRDDET.free;
  TOBTVAPRETTC.free;
end;

procedure TImprPieceViaTOB.DefiniLigneTOTALNET (ThePiece : TOB;TTCSit,TTCPre,ReliquatRGSit,reliquatRGPre,AcomptesSit,AcomptesPre,Timbres,MTRDSITTTC,MTRDPRETTC,MtPortTTCPre,MtPortTTC : double);
var TOBL : TOB;
begin
  TOBL := InsereTOB ('LIGNE',ThePiece);
  TOBL.putValue('LIBELLE','NET A PAYER');
  // reliquatRGSit est négatif si la RG n'est pas couverte par une caution sinon il est positif
  if (ThePiece.getValue('PRG_TYPERG')='TTC') and (ReliquatRGSit < 0) then TOBL.putValue('MONTANTSIT',TTCSit-AcomptesSit+ReliquatRGSit)
  else TOBL.putValue('MONTANTSIT',TTCSit-AcomptesSit+Timbres);
  // reliquatRGPre est négatif si la RG n'est pas couverte par une caution sinon il est positif
  if (ThePiece.getValue('PRG_TYPERG')='TTC') and (ReliquatRGPre < 0) then TOBL.putValue('MONTANTPRE',TTCPre-AcomptesPre+reliquatRGPre)
  else TOBL.putValue('MONTANTPRE',TTCPre-AcomptesPre);

  TOBL.putValue('MONTANTSIT',TOBL.GetDouble('MONTANTSIT')-MTRDSITTTC+MtPortTTC);
  TOBL.putValue('MONTANTPRE',TOBL.GetDouble('MONTANTPRE')-MTRDPRETTC+MtPortTTCPre);
  TOBL.putValue('ISLIGNESTD','X');
  TOBL.putValue('ISLIGNEPOR','X');
  ThePiece.Putvalue ('TOTALNET', TOBL.GetValue('MONTANTSIT'));
end;

procedure TImprPieceViaTOB.DefiniLigneACOMPTES(ThePiece : TOB; var AcomptesSit,AcomptesPre : double);
var TOBL,TOBS : TOB;
    Indice : integer;
begin

  AcomptesPre := 0;

  for Indice := 0 to fTOBSItuations.detail.count -1 do
  begin
    TOBS := fTOBSItuations.detail[Indice];
    if Indice = fTOBSItuations.detail.count -1 then
    begin
      AcomptesSit := TOBS.getValue('ACOMPTESIT');
    end else
    begin
      AcomptesPre := AcomptesPre + TOBS.getValue('ACOMPTESIT');
    end;
  end;

  if (AcomptesSit=0) and (AcomptesPre=0) then exit;

  TOBL := InsereTOB ('LIGNE',ThePiece);
  TOBL.putValue('LIBELLE','ACOMPTE/REGLEMENT');
  TOBL.putValue('MONTANTSIT',(AcomptesSit * -1));
  TOBL.putValue('MONTANTPRE',(AcomptesPre * -1));
  TOBL.putValue('ISLIGNESTD','X');
  TOBL.putValue('ISLIGNEPOR','X');

end;

function TImprPieceViaTOB.DefiniLigneTimbres(ThePiece,TOBTimbres : TOB) : double;
var Cumul : Double;
		Indice : Integer;
    TOBL : TOB;
begin
  Cumul := 0;
	for Indice := 0 to TOBTimbres.detail.count -1 do
  begin
    Cumul := Cumul + TOBTImbres.detail[Indice].getValue('BT0_VALEUR');
  end;
  if Cumul > 0 then
  begin
    TOBL := InsereTOB ('LIGNE',ThePiece);
    TOBL.putValue('LIBELLE','TIMBRES');
  	TOBL.putValue('MONTANTSIT',Cumul);
    TOBL.putValue('MONTANTPRE',0);
    TOBL.putValue('ISLIGNESTD','X');
  end;
  Result := Cumul;
end;

procedure TImprPieceViaTOB.DefiniLigneRGTTC(ThePiece,TOBRGPRE : TOB; var ReliquatRgSit,ReliquatRgPre : double);
var TOBL : TOB;
		XD,XP : double;
    NumCautionSit,NumCautionpre : string;
begin
  if (ThePiece.GetValue('PRG_TYPERG')<>'TTC') or (ThePiece.GetValue('PRG_TAUXRG')=0) OR (ThePiece.getValue('GPP_APPLICRG')<>'X') then exit;
  TOBL := InsereTOB ('LIGNE',ThePiece);
  TOBL.putValue('LIBELLE','RETENUE DE GARANTIE TTC '+FloatToStr(ThePiece.GetValue('PRG_TAUXRG'))+' %');
  //
  GetRg (TOBPieceRG,false,True,Xp,XD,NuMCautionSit);
  ReliquatRgSit := XD * (-1);
  //
  if TOBRGPre.detail.count > 0 then
  begin
    GetRg (TOBRGPRE,false,True,Xp,XD,NuMCautionPre);
    ReliquatRgPre := XD * (-1);
  end;
  //
  if numCautionSit <> '' then
  begin
    TOBL.putValue('LIBELLESIT','CAUTION');
  end else
  begin
    TOBL.putValue('MONTANTSIT',ReliquatRGSit);
  end;
  if NumCautionPre <> '' then
  begin
    if (fNumeroSit<>1) then  TOBL.putValue('LIBELLEPRE','CAUTION');
  end else
  begin
    TOBL.putValue('MONTANTPRE',ReliquatRGPRE);
  end;
  if (ReliquatRgPre+ReliquatRgSit=0) then
  begin
    if (fNumeroSit<>1) then TOBL.putValue('LIBELLECUM','CAUTION');
  end;
  TOBL.putValue('ISLIGNESTD','X');
  TOBL.putValue('ISLIGNEPOR','X');
end;


procedure TImprPieceViaTOB.DefiniLigneTOTALTTCInterm (ThePiece,TOBFDetaille : TOB; CumSitTVA,CumPreTVA,CumRGSit,CumRgPre,TVARGSIt,TVARGPre,MTRDSIT,MTRDPRE,TVARDSIT,TVARDPRE,MtHtPortTTCPre,MtHTPortTTC : double; var TTCSIt,TTCPRE : double);

  procedure CalculeMT (TOBFDetaille : TOB; var MTSIT,MTPRE : double);
  var II : Integer;
      TOBF : TOB;
  begin
    MTSIT := 0; MTPRE := 0;
    for II := 0 TO TOBFDetaille.Detail.count -1 do
    begin
      TOBF := TOBFDetaille.detail[II];
      if (TOBF.GetString('TYPE')='HT') or
         (TOBF.GetString('TYPE')='MI') or
         (TOBF.GetString('TYPE')='MT') then continue;
      MTSIT := MTSIT + TOBF.GetDouble('MONTANTTTCSIT');
      if fAvancement then MTPRE := MTPRE + TOBF.GetDouble('MONTANTTTCPRE');
    end;

  end;

var TOBL : TOB;
    CUMSIT,CUMPRE : Double;
    Valsit,valpre : double;
begin
  if TOBFDetaille.detail.count = 0 then Exit;
//  CalculeMT (TOBFDetaille,CUMSIT,CUMPRE);
  if (CUMSIT = 0) and (CUMPRE = 0) then Exit;
  TOBL := InsereTOB ('LIGNE',ThePiece);
  TOBL.putValue('LIBELLE','TOTAL T.T.C.');
  if ThePiece.getValue('PRG_TYPERG')='HT' then
  begin
    ValSit := ThePiece.getValue('GP_TOTALHTDEV')+CumSitTVA+CumRGSit-TVARGSit-CUMSIT;
    if fAvancement then ValPre := ThePiece.getValue('TOTSITPREC')+CumPreTVA+CumRGPre-TVARGPre-CUMPRE;
  end else
  begin
    ValSit := ThePiece.getValue('GP_TOTALHTDEV')+CumSitTVA-CUMSIT;
    if fAvancement then valPre := ThePiece.getValue('TOTSITPREC')+CumPreTVA-CUMPRE;
  end;
  valSit := valSit-MTRDSIT-TVARDSIT-MtHTPortTTC;
  if fAvancement then ValPre := ValPre-MTRDPRE-TVARDPRE-MtHtPortTTCPre;
  TOBL.putValue('MONTANTSIT',ValSit);
  if fAvancement then TOBL.putValue('MONTANTPRE',valpre);
  TOBL.putValue('ISLIGNEPOR','X');
end;

procedure TImprPieceViaTOB.DefiniLigneTOTALTTC (ThePiece,TOBfDetaille : TOB; CumSitTVA,CumPreTVA,CumRGSit,CumRgPre,TVARGSIt,TVARGPre,MTRDSIT,MTRDPRE,TVARDSIT,TVARDPRE,MtHtPortTTCPre,MtHTPortTTC : double; var TTCSIt,TTCPRE : double);
var TOBL : TOB;
    valSit,valpre : double;
begin
  TOBL := InsereTOB ('LIGNE',ThePiece);
  if TOBfDetaille.Detail.Count = 0 then TOBL.putValue('LIBELLE','TOTAL T.T.C.')
                                   else TOBL.putValue('LIBELLE','TOTAL T.T.C. NET');
  if ThePiece.getValue('PRG_TYPERG')='HT' then
  begin
    valSit := ThePiece.getValue('GP_TOTALHTDEV')+CumSitTVA+CumRGSit+TVARGSit;
    if fAvancement then Valpre := ThePiece.getValue('TOTSITPREC')+CumPreTVA+CumRGPre+TVARGPre;
  end else
  begin
    valSit := ThePiece.getValue('GP_TOTALHTDEV')+CumSitTVA;
    if fAvancement then valPre := ThePiece.getValue('TOTSITPREC')+CumPreTVA
  end;
  ValSit := ValSit-MTRDSIT-TVARDSIT;
  if fAvancement then valpre := valpre - MTRDPRE - TVARDPRE - MtHtPortTTCPre;
  TOBL.putValue('MONTANTSIT',valSit-MtHTPortTTC);
  if fAvancement then TOBL.putValue('MONTANTPRE',Valpre);
  TTCSit := TOBL.GetValue('MONTANTSIT');
  if fAvancement then TTCPre := TOBL.GetValue('MONTANTPRE');
  TOBL.putValue('ISLIGNESTD','X');
  TOBL.putValue('ISLIGNEPOR','X');
end;

procedure TImprPieceViaTOB.DefiniLigneTVARD (ThePiece,TOBFDDET: TOB; var MTSIT,MTPRE : double);
var II,JJ : integer;
    TOBL,TF,TFL : TOB;
begin
  for II := 0 to TOBFDDET.detail.count -1 do
  begin
    TF := TOBFDDET.detail[II];
    if (TF.GetString('TYPE')<>'HT') and
       (TF.GetString('TYPE')<>'MI') and
       (TF.GetString('TYPE')<>'MT') then continue;
    for JJ := 0 to TF.detail.count -1 do
    begin
      TFL := TF.detail[JJ];
      TOBL := InsereTOB ('LIGNE',ThePiece);
      TOBL.putValue('LIBELLE','TVA Sur '+TF.GetString('LIBELLE'));
      TOBL.putValue('MONTANTSIT',TFL.GetDouble('VALEURTAXESIT')* (-1));
      if fAvancement then TOBL.putValue('MONTANTPRE',TFL.GetDouble('VALEURTAXEPRE') * (-1));
      TOBL.putValue('ISLIGNESTD','X');
      TOBL.putValue('ISLIGNEPOR','X');
      MtSIT := MtSit + TFL.GetDouble('VALEURTAXESIT');
      if fAvancement then MtPRE := MtPRE + TFL.GetDouble('VALEURTAXEPRE');
    end;
  end;
end;

procedure TImprPieceViaTOB.DefiniLigneTVARG(ThePiece,TOBRGPRE : TOB; var TVARGSIt,TVARGPre : double);
var TOBL : TOB;
    Indice : integer;
    Xp,Xd : double;
begin
  if (ThePiece.GetValue('PRG_TYPERG')<>'HT') or (ThePiece.GetValue('PRG_TAUXRG')=0) OR (ThePiece.getValue('GPP_APPLICRG')<>'X') then exit;

  GetTvaRg (TOBPieceRG,True,True,Xp,XD);
  if XD <> 0 then TVARGSIt := XD * (-1) else TVARGSIt := 0;

  GetTvaRg (TOBRGPre,True,True,Xp,XD);
  if XD <> 0 then  TVARGPre := XD * (-1) else TVARGPre := 0;
  // --
  if (TvaRgPre = 0) and (TVARGSIt=0) then exit;
  // --
  TOBL := InsereTOB ('LIGNE',ThePiece);
  TOBL.putValue('LIBELLE','TVA SUR RETENUE DE GARANTIE');
  if TVARGSIt<>0 then
  begin
    TOBL.putValue('MONTANTSIT',arrondi(TVARGSIt,4));
  end else
  begin
    TOBL.putValue('LIBELLESIT','CAUTION');
  end;
  if TVARGPre<>0 then
  begin
    TOBL.putValue('MONTANTPRE',arrondi(TVARGPre,4));
  end else
  begin
    if (fNumeroSit<>1) then TOBL.putValue('LIBELLEPRE','CAUTION');
  end;
  if TVARGSIt+TVARGPre = 0 then if (fNumeroSit<>1) then TOBL.putValue('LIBELLECUM','CAUTION');

  TOBL.putValue('ISLIGNESTD','X');
  TOBL.putValue('ISLIGNEPOR','X');
end;


procedure TImprPieceViaTOB.DefiniLigneRDHT(ThePiece,TOBFraisDet :TOB; Var ValSit,VAlpre : double);
var TOBL,TOBF : TOB;
    II : Integer;
    ISens : integer;
begin
  if TOBFraisDet.detail.count = 0 then Exit;
  for II := 0 to TOBFraisDet.Detail.Count -1 do
  begin
    TOBF := TOBFraisDet.detail[II];
    if (TOBF.GetString('TYPE')<>'HT') and
       (TOBF.GetString('TYPE')<>'MI') and
       (TOBF.GetString('TYPE')<>'MT') then continue;

   //
    TOBL := InsereTOB ('LIGNE',ThePiece);
    TOBL.putValue('LIBELLE',TOBF.GetString('LIBELLE'));
    TOBL.putValue('MONTANTSIT',TOBF.GetDouble('MONTANTHTSIT')*(-1));
    TOBL.putValue('MONTANTPRE',TOBF.GetDouble('MONTANTHTPRE')*(-1));

    TOBL.putValue('ISLIGNEPOR','X');
    valSit := valSit + TOBF.GetDouble('MONTANTHTSIT');
    valpre := valpre + TOBF.GetDouble('MONTANTHTPRE');
  end;
end;


procedure TImprPieceViaTOB.DefiniLigneRDTTC(ThePiece,TOBFraisDet :TOB; Var ValSit,VAlpre : double);
var TOBL,TOBF : TOB;
    II : Integer;
    ISens : integer;
begin
  if TOBFraisDet.detail.count = 0 then Exit;
  for II := 0 to TOBFraisDet.Detail.Count -1 do
  begin
    TOBF := TOBFraisDet.detail[II];
    if (TOBF.GetString('TYPE')<>'PT') and
       (TOBF.GetString('TYPE')<>'MIC') and
       (TOBF.GetString('TYPE')<>'MTC') then continue;
   //
    TOBL := InsereTOB ('LIGNE',ThePiece);
    TOBL.putValue('LIBELLE',TOBF.GetString('LIBELLE'));
    TOBL.putValue('MONTANTSIT',TOBF.GetDouble('MONTANTTTCSIT')*(-1));
    TOBL.putValue('MONTANTPRE',TOBF.GetDouble('MONTANTTTCPRE')*(-1));

    TOBL.putValue('ISLIGNESTD','X');
    TOBL.putValue('ISLIGNEPOR','X');
    valSit := valSit + TOBF.GetDouble('MONTANTTTCSIT');
    valpre := valpre + TOBF.GetDouble('MONTANTTTCPRE');
  end;
end;

procedure TImprPieceViaTOB.DefiniLigneRGHT (ThePiece,TOBRGPRE : TOB; Var CumSitHt,CumpreHt,ReliquatRGSit,reliquatRGPre : double);
var TOBL : TOB;
    Indice : integer;
    RgSit,Rgpre,Caution,XD,XP,XX ,MT: double;
    NumCaution : string;
begin
  CumSitHt := 0; CumpreHt := 0;
  ReliquatRgSit := 0; reliquatRgPre := 0;
  // Situation
  GetRg (TOBPieceRG,True,True,Xp,XD,NuMCaution);
  if XD <> 0 then ReliquatRgSit := XD * (-1) else ReliquatRgSit := 0;
  for Indice := 0 to fTOBPieceRG.detail.count -1 do
  begin
    TOBL := fTOBPieceRG.detail[Indice];
    RgSit := RgSit+ TOBL.getValue('PRG_MTHTRGDEV');
  end;
  CumSitHt := RgSit * (-1);
  // Précédent
  GetRg (TOBRGPre,True,True,Xp,XD,NuMCaution);
  if XD <> 0 then  ReliquatRgpre := XD * (-1) else ReliquatRgpre := 0;
  for Indice := 0 to TOBRGPre.detail.count -1 do
  begin
    TOBL := TOBRGPre.detail[Indice];
    Rgpre := Rgpre+ TOBL.getValue('PRG_MTHTRGDEV');
  end;
  //
  CumpreHt := RgPre*(-1);
  //
  if (ThePiece.GetValue('PRG_TYPERG')<>'HT') or (ThePiece.GetValue('PRG_TAUXRG')=0) OR (ThePiece.getValue('GPP_APPLICRG')<>'X') then exit;
  TOBL := InsereTOB ('LIGNE',ThePiece);
  TOBL.putValue('LIBELLE','RETENUE DE GARANTIE HT '+FloatToStr(ThePiece.GetValue('PRG_TAUXRG'))+' %');
  Mt := Arrondi(ReliquatRgSit,V_PGI.okdecV);
  if Mt <>0 then
  begin
    TOBL.putValue('MONTANTSIT',ReliquatRgSit);
  end else
  begin
    TOBL.putValue('LIBELLESIT','CAUTION');
  end;
  Mt := ARRONDI(ReliquatRgpre,V_PGI.OkDecV);
  if Mt <>0 then
  begin
    TOBL.putValue('MONTANTPRE',ReliquatRgpre);
  end else
  begin
    if fNumeroSit <> 1 then TOBL.putValue('LIBELLEPRE','CAUTION');
  end;
  CumSitHt := ReliquatRgSit;
  CumpreHt := ReliquatRgpre;
  if (ReliquatRgSit+ReliquatRgpre = 0) and (fNumeroSit <> 1)then TOBL.putValue('LIBELLECUM','CAUTION');

  TOBL.putValue('ISLIGNESTD','X');
  TOBL.putValue('ISLIGNEPOR','X');
end;

procedure TImprPieceViaTOB.ConstitueTOBRGPrec (ThePiece,TOBRGPre : TOB);
var req : string;
		Indice : integer;
    CumXd,CumXp : double;
    TOBLOC,TOBL,TOBP : TOB;
begin
  TOBLOC := TOB.Create ('LES RG DETAILLES',nil,-1);
  TOBRGPre.clearDetail;
  CumXP := 0; CumXd := 0;
  if fAvancement then
  begin
    if (fIsDGD and fIsAVA) then
    begin
      Req := 'SELECT BST_NATUREPIECE,BST_SOUCHE,BST_NUMEROFAC,BST_NUMEROSIT,PRG_NUMLIGNE,PRG_FOURN,'+
             'PRG_MTHTRG,PRG_MTHTRGDEV,PRG_MTTTCRGDEV,PRG_MTTTCRG,PRG_NUMCAUTION,PRG_CAUTIONMT,PRG_CAUTIONMTDEV,'+
             'GL_PIECEORIGINE FROM BSITUATIONS '+
             'LEFT JOIN PIECERG ON PRG_NATUREPIECEG=BST_NATUREPIECE AND '+
             'PRG_SOUCHE=BST_SOUCHE AND PRG_NUMERO=BST_NUMEROFAC AND PRG_INDICEG=0 '+
             'LEFT JOIN LIGNE ON GL_NATUREPIECEG=BST_NATUREPIECE AND GL_SOUCHE=BST_SOUCHE AND '+
             'GL_NUMERO=BST_NUMEROFAC AND GL_INDICEG=0 AND GL_NUMORDRE=PRG_NUMLIGNE '+
             'WHERE BST_SSAFFAIRE="'+ThePiece.GetValue('GP_AFFAIREDEVIS')+'" AND '+
             'BST_VIVANTE="X" AND '+
             'BST_NUMEROSIT <= '+ThePiece.getString('NUMSIT')+' ORDER BY GL_PIECEORIGINE,PRG_FOURN,BST_NUMEROSIT';
    end else
    begin
      Req := 'SELECT BST_NATUREPIECE,BST_SOUCHE,BST_NUMEROFAC,BST_NUMEROSIT,PRG_NUMLIGNE,PRG_FOURN,'+
             'PRG_MTHTRG,PRG_MTHTRGDEV,PRG_MTTTCRGDEV,PRG_MTTTCRG,PRG_NUMCAUTION,PRG_CAUTIONMT,PRG_CAUTIONMTDEV,'+
             'GL_PIECEORIGINE FROM BSITUATIONS '+
             'LEFT JOIN PIECERG ON PRG_NATUREPIECEG=BST_NATUREPIECE AND '+
             'PRG_SOUCHE=BST_SOUCHE AND PRG_NUMERO=BST_NUMEROFAC AND PRG_INDICEG=0 '+
             'LEFT JOIN LIGNE ON GL_NATUREPIECEG=BST_NATUREPIECE AND GL_SOUCHE=BST_SOUCHE AND '+
             'GL_NUMERO=BST_NUMEROFAC AND GL_INDICEG=0 AND GL_NUMORDRE=PRG_NUMLIGNE '+
             'WHERE BST_SSAFFAIRE="'+ThePiece.GetValue('GP_AFFAIREDEVIS')+'" AND '+
             'BST_VIVANTE="X" AND '+
             'BST_NUMEROSIT < '+ThePiece.getString('NUMSIT')+' '+
             'ORDER BY GL_PIECEORIGINE,PRG_FOURN,BST_NUMEROSIT';
    end;
  end
  else
  begin
    Req := 'SELECT GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO,PRG_NUMLIGNE,PRG_FOURN,'+
           'PRG_MTHTRG,PRG_MTHTRGDEV,PRG_MTTTCRGDEV,PRG_MTTTCRG,PRG_NUMCAUTION,PRG_CAUTIONMT,PRG_CAUTIONMTDEV,'+
           'GL_PIECEORIGINE FROM PIECE '+
           'LEFT JOIN PIECERG ON PRG_NATUREPIECEG=GP_NATUREPIECEG AND '+
           'PRG_SOUCHE=GP_SOUCHE AND PRG_NUMERO=GP_NUMERO AND PRG_INDICEG=0 '+
           'LEFT JOIN LIGNE ON GL_NATUREPIECEG=GP_NATUREPIECEG AND GL_SOUCHE=GP_SOUCHE AND '+
           'GL_NUMERO=GP_NUMERO AND GL_INDICEG=0 AND GL_NUMORDRE=PRG_NUMLIGNE '+
           'WHERE GP_AFFAIREDEVIS="'+ThePiece.GetValue('GP_AFFAIREDEVIS')+'" AND '+
           'GP_VIVANTE="X" '+
           'ORDER BY GL_PIECEORIGINE,PRG_FOURN,PRG_NUMERO';
  end;

  TOBLOC.LoadDetailDBFromSQL ('PIECERG',req,false);
  for Indice := 0 to TOBLOC.detail.count -1 do
  begin
    TOBL := TOBLOC.detail[Indice];
    if (TOBL.getString('PRG_NUMCAUTION') <> '') and ( TOBL.getDouble('PRG_CAUTIONMT')<>0) then Continue; // mt rg couvert par caution
    TOBP := TOBRGPre.findFirst(['PIECEPRECEDENTE','PRG_FOURN'],[TOBL.GetValue('GL_PIECEORIGINE'),TOBL.GetValue('PRG_FOURN')],true);
    if TOBP = nil then
    begin
			TOBP := TOB.Create ('PIECERG',TOBRGPre,-1);
      TOBP.Dupliquer(TOBL,false,true);
      TOBP.AddChampSupValeur('PIECEPRECEDENTE',TOBL.GetValue('GL_PIECEORIGINE'));
      if TOBP.GetString('PRG_NUMCAUTION')='' then
      begin
      	TOBP.SetDouble('PRG_CAUTIONMT',0);
      	TOBP.SetDouble('PRG_CAUTIONMTDEV',0);
      end;
    end else
    begin
      TOBP.SetDouble ('PRG_MTHTRG',TOBP.GetDouble ('PRG_MTHTRG')+TOBL.GetDouble ('PRG_MTHTRG'));
      TOBP.SetDouble ('PRG_MTHTRGDEV',TOBP.GetDouble ('PRG_MTHTRGDEV')+TOBL.GetDouble ('PRG_MTHTRGDEV'));
      TOBP.SetDouble ('PRG_MTTTCRG',TOBP.GetDouble ('PRG_MTTTCRG')+TOBL.GetDouble ('PRG_MTTTCRG'));
      TOBP.SetDouble ('PRG_MTTTCRGDEV',TOBP.GetDouble ('PRG_MTTTCRGDEV')+TOBL.GetDouble ('PRG_MTTTCRGDEV'));
      if TOBP.GetString('PRG_NUMCAUTION')='' then
      begin
        TOBP.SetDouble('PRG_CAUTIONMT',0);
        TOBP.SetDouble('PRG_CAUTIONMTDEV',0);
      end;
    end;
  end;
  TOBLOC.Free;
end;

procedure TImprPieceViaTOB.ConstitueTOBTVA (ThePiece,TOBTVAPre,TOBTVAPRETTC : TOB; var MtHtPortTTCPre,MtHTPortTTC,MtPortTTCPre,MtPortTTC : double);
var Indice : integer;
    TvASit,TOBS,TVAC,TVAS,TT,TOBT : TOB;
    req : string;
begin
  MtHtPortTTCPre := 0;
  MtHTPortTTC := 0;
  MtPortTTCPre := 0;
  MtPortTTC := 0;
  TOBTVAPre.clearDetail;
  TVASit := TOB.Create ('TVA SITUATION',nil,-1);
//
  if fAvancement then
  begin
    if (fISDGD and fIsAVA) then
    begin
      req := 'SELECT * FROM PIEDBASE, BSITUATIONS WHERE '+
             'BST_SSAFFAIRE="'+ThePiece.GetValue('GP_AFFAIREDEVIS')+'" AND '+
             'BST_NATUREPIECE IN ("FBP","FBT") AND '+
             'BST_NUMEROSIT <= '+ThePiece.getString('NUMSIT')+' AND '+
             'BST_VIVANTE="X" AND '+
             'GPB_NATUREPIECEG=BST_NATUREPIECE AND GPB_SOUCHE=BST_SOUCHE AND GPB_NUMERO=BST_NUMEROFAC';
    end
    else
    begin
      req := 'SELECT * FROM PIEDBASE, BSITUATIONS WHERE '+
             'BST_SSAFFAIRE="'+ThePiece.GetValue('GP_AFFAIREDEVIS')+'" AND '+
             'BST_NATUREPIECE IN ("FBP","FBT") AND '+
             'BST_NUMEROSIT < '+ThePiece.getString('NUMSIT')+' AND '+
             'BST_VIVANTE="X" AND '+
             'GPB_NATUREPIECEG=BST_NATUREPIECE AND GPB_SOUCHE=BST_SOUCHE AND GPB_NUMERO=BST_NUMEROFAC';
    end;
  end
  else
  begin
    req := 'SELECT * FROM PIEDBASE, PIECE WHERE '+
           'GP_AFFAIREDEVIS="' + ThePiece.GetValue('GP_AFFAIREDEVIS') + '" AND ' +
           'GP_NATUREPIECE IN ("FBP","FBT") AND '+
           'GP_NUMERO      = ' + IntToStr(ThePiece.GetValue('GP_NUMERO'))       + '  AND ' +
           'GP_VIVANTE="X" AND '+
           'GPB_NATUREPIECEG=GP_NATUREPIECEG AND GPB_SOUCHE=GP_SOUCHE AND GPB_NUMERO=GP_NUMERO';
  end;
    //
  TVASIt.LoadDetailDBFromSQL ('PIEDBASE',Req,false);
  for Indice := 0 to TVASIt.detail.count -1 do
  begin
    TVAS :=TVASit.detail[indice];
    if (VH_GC.BTCODESPECIF = '001') and (TVAS.getString('GPB_FOURN')<> '') then continue;
    TVAC := TOBTVAPre.findFirst(['FAMILLETAXE'],[TVAS.getValue('GPB_FAMILLETAXE')],true);
    if TVAC = nil then
    begin
      TVAC := TOB.create ('UNE TVA',TOBTVAPRE,-1);
      TVAC.AddChampSupValeur ('FAMILLETAXE',TVAS.getValue('GPB_FAMILLETAXE'));
      TVAC.AddChampSupValeur ('TAUXTAXE',TVAS.getValue('GPB_TAUXTAXE'));
      TVAC.AddChampSupValeur ('BASETAXE',0);
      TVAC.AddChampSupValeur ('BASETAXEPRE',0);
      TVAC.AddChampSupValeur ('BASETAXECUM',0);
      TVAC.AddChampSupValeur ('VALEURTAXEPRE',0);
      TVAC.AddChampSupValeur ('VALEURTAXESIT',0);
    end;
    if fAvancement then TVAC.putValue('VALEURTAXEPRE',  TVAC.GetValue('VALEURTAXEPRE')+TVAS.GetValue('GPB_VALEURTAXE'));
    if fAvancement then TVAC.putValue('BASETAXEPRE',    TVAC.GetValue('BASETAXEPRE')+TVAS.getValue('GPB_BASEDEV'));
    TVAC.putValue('BASETAXECUM',TVAC.GetValue('BASETAXECUM')+TVAS.getValue('GPB_BASEDEV'));
  end;
  for Indice := 0 to fTOBBases.detail.count -1 do
  begin
    TVAS := fTOBBases.detail[Indice];
    if (VH_GC.BTCODESPECIF = '001') and (TVAS.getString('GPB_FOURN')<> '') then continue;
    TVAC := TOBTVAPre.findFirst(['FAMILLETAXE'],[TVAS.getValue('GPB_FAMILLETAXE')],true);
    if TVAC = nil then
    begin
      TVAC := TOB.create ('UNE TVA',TOBTVAPRE,-1);
      TVAC.AddChampSupValeur ('FAMILLETAXE',TVAS.getValue('GPB_FAMILLETAXE'));
      TVAC.AddChampSupValeur ('TAUXTAXE',TVAS.getValue('GPB_TAUXTAXE'));
      TVAC.AddChampSupValeur ('BASETAXE',0);
      TVAC.AddChampSupValeur ('BASETAXEPRE',0);
      TVAC.AddChampSupValeur ('BASETAXECUM',0);
      TVAC.AddChampSupValeur ('VALEURTAXEPRE',0);
      TVAC.AddChampSupValeur ('VALEURTAXESIT',0);
    end;
    TVAC.putValue('VALEURTAXESIT',TVAC.GetValue('VALEURTAXESIT')+TVAS.GetValue('GPB_VALEURTAXE'));
    TVAC.putValue('BASETAXE',TVAC.GetValue('BASETAXE')+TVAS.getValue('GPB_BASEDEV'));
    TVAC.putValue('BASETAXECUM',TVAC.GetValue('BASETAXECUM')+TVAS.getValue('GPB_BASEDEV'));
  end;
  TVASIT.clearDetail;
  //
  if VH_GC.BTCODESPECIF = '001' then
  begin
    TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],['TX1',ThePiece.GetValue('GP_REGIMETAXE'),'TN'],False) ;
    if fAvancement then
    begin
      req := 'SELECT * FROM BSITUATIONS WHERE '+
             'BST_SSAFFAIRE="'+ThePiece.GetValue('GP_AFFAIREDEVIS')+'" AND '+
             'BST_NATUREPIECE="'+ThePiece.GetValue('GP_NATUREPIECEG')+'" AND '+
             'BST_NUMEROSIT = 0 AND '+
             'BST_VIVANTE="X"';
      TVASIt.LoadDetailDBFromSQL ('BSITUATIONS',Req,false);
      for Indice := 0 to TvASit.detail.count -1 do
      begin
        TVAS := TvASit.detail[Indice];
        TVAC := TOBTVAPre.findFirst(['FAMILLETAXE'],['TN'],true);
        if TVAC = nil then
        begin
          TVAC := TOB.create ('UNE TVA',TOBTVAPRE,-1);
          TVAC.AddChampSupValeur ('FAMILLETAXE','TN');
          TVAC.AddChampSupValeur ('TAUXTAXE',TOBT.GetValue('TV_TAUXVTE'));
          TVAC.AddChampSupValeur ('BASETAXE',0);
          TVAC.AddChampSupValeur ('BASETAXEPRE',0);
          TVAC.AddChampSupValeur ('BASETAXECUM',0);
          TVAC.AddChampSupValeur ('VALEURTAXEPRE',0);
          TVAC.AddChampSupValeur ('VALEURTAXESIT',0);
        end;
        TVAC.putValue('VALEURTAXEPRE',TVAC.GetValue('VALEURTAXEPRE')+TVAS.GetValue('BST_MONTANTTVA'));
        TVAC.putValue('BASETAXEPRE',TVAC.GetValue('BASETAXEPRE')+TVAS.getValue('BST_MONTANTHT'));
        TVAC.putValue('BASETAXECUM',TVAC.GetValue('BASETAXECUM')+TVAS.getValue('BST_MONTANTHT'));
      end;

    end;
  end;
  TVASIT.clearDetail;
  //
  // Constitution de la TVA sur Port TTC
  //
  if fAvancement then
  begin
    if (fISDGD and fIsAVA) then
    begin
      req := 'SELECT * FROM PIEDPORT, BSITUATIONS WHERE '+
             'BST_SSAFFAIRE="'+ThePiece.GetValue('GP_AFFAIREDEVIS')+'" AND '+
             'BST_NATUREPIECE="'+ThePiece.GetValue('GP_NATUREPIECEG')+'" AND '+
             'BST_NUMEROSIT <= '+ThePiece.getString('NUMSIT')+' AND '+
             'BST_VIVANTE="X" AND '+
             'GPT_NATUREPIECEG=BST_NATUREPIECE AND GPT_SOUCHE=BST_SOUCHE AND GPT_NUMERO=BST_NUMEROFAC AND GPT_TYPEPORT IN ("MTC","MIC","PT") AND GPT_RETENUEDIVERSE<>"X"';
    end
    else
    begin
      req := 'SELECT * FROM PIEDPORT, BSITUATIONS WHERE '+
             'BST_SSAFFAIRE="'+ThePiece.GetValue('GP_AFFAIREDEVIS')+'" AND '+
             'BST_NATUREPIECE="'+ThePiece.GetValue('GP_NATUREPIECEG')+'" AND '+
             'BST_NUMEROSIT < '+ThePiece.getString('NUMSIT')+' AND '+
             'BST_VIVANTE="X" AND '+
             'GPT_NATUREPIECEG=BST_NATUREPIECE AND GPT_SOUCHE=BST_SOUCHE AND GPT_NUMERO=BST_NUMEROFAC AND GPT_TYPEPORT IN ("MTC","MIC","PT") AND GPT_RETENUEDIVERSE<>"X"';
    end;
  end
  else
  begin
    req := 'SELECT * FROM PIEDPORT, PIECE WHERE '+
           'GP_AFFAIREDEVIS="' + ThePiece.GetValue('GP_AFFAIREDEVIS') + '" AND ' +
           'GP_NATUREPIECEG="' + ThePiece.GetValue('GP_NATUREPIECEG') + '" AND ' +
           'GP_NUMERO      = ' + IntToStr(ThePiece.GetValue('GP_NUMERO'))       + '  AND ' +
           'GP_VIVANTE="X" AND '+
           'GPT_NATUREPIECEG=GP_NATUREPIECEG AND GPT_SOUCHE=GP_SOUCHE AND GPT_NUMERO=GP_NUMERO AND GPT_TYPEPORT IN ("MTC","MIC","PT") AND GPT_RETENUEDIVERSE<>"X"';
  end;
  TVASIt.LoadDetailDBFromSQL ('PIEDPORT',Req,false);
  for Indice := 0 to TVASIt.detail.count -1 do
  begin
    TVAS :=TVASit.detail[indice];
    TVAC := TOBTVAPRETTC.findFirst(['FAMILLETAXE'],[TVAS.getValue('GPT_FAMILLETAXE1')],true);
    if TVAC = nil then
    begin
      TVAC := TOB.create ('UNE TVA',TOBTVAPRETTC,-1);
      TVAC.AddChampSupValeur ('FAMILLETAXE',TVAS.getValue('GPT_FAMILLETAXE1'));
      TVAC.AddChampSupValeur ('TAUXTAXE',GetTauxTaxe(fTobPiece.GetString('GP_REGIMETAXE'),TVAS.getValue('GPT_FAMILLETAXE1')));
      TVAC.AddChampSupValeur ('BASETAXE',0);
      TVAC.AddChampSupValeur ('BASETAXEPRE',0);
      TVAC.AddChampSupValeur ('BASETAXECUM',0);
      TVAC.AddChampSupValeur ('VALEURTAXEPRE',0);
      TVAC.AddChampSupValeur ('VALEURTAXESIT',0);
    end;
    if fAvancement then TVAC.putValue('VALEURTAXEPRE',  TVAC.GetValue('VALEURTAXEPRE')+TVAS.GetValue('GPT_TOTALTAXEDEV1'));
    if fAvancement then TVAC.putValue('BASETAXEPRE',    TVAC.GetValue('BASETAXEPRE')+TVAS.getValue('GPT_TOTALHTDEV'));
    TVAC.putValue('BASETAXECUM',TVAC.GetValue('BASETAXECUM')+TVAS.getValue('GPT_TOTALHTDEV'));
    MtHtPortTTCPre := MtHtPortTTCPre + TVAS.getValue('GPT_TOTALHTDEV');
    MtPortTTCPre := MtPortTTCPre + TVAS.getValue('GPT_TOTALTTCDEV');
  end;
  for Indice := 0 to fTOBPorcs.detail.count -1 do
  begin
    TVAS := fTOBPorcs.detail[Indice];
    if (TVAS.GetBoolean('GPT_RETENUEDIVERSE')) then continue;
    if (TVAS.GeTString('GPT_TYPEPORT') <> 'MIC') and (TVAS.GeTString('GPT_TYPEPORT') <> 'MTC') and (TVAS.GeTString('GPT_TYPEPORT') <> 'PT') then continue;
    TVAC := TOBTVAPRETTC.findFirst(['FAMILLETAXE'],[TVAS.getValue('GPT_FAMILLETAXE1')],true);
    if TVAC = nil then
    begin
      TVAC := TOB.create ('UNE TVA',TOBTVAPRETTC,-1);
      TVAC.AddChampSupValeur ('FAMILLETAXE',TVAS.getValue('GPT_FAMILLETAXE1'));
      TVAC.AddChampSupValeur ('TAUXTAXE',GetTauxTaxe(fTobPiece.GetString('GP_REGIMETAXE'),TVAS.getValue('GPT_FAMILLETAXE1')));
      TVAC.AddChampSupValeur ('BASETAXE',0);
      TVAC.AddChampSupValeur ('BASETAXEPRE',0);
      TVAC.AddChampSupValeur ('BASETAXECUM',0);
      TVAC.AddChampSupValeur ('VALEURTAXEPRE',0);
      TVAC.AddChampSupValeur ('VALEURTAXESIT',0);
    end;
    TVAC.putValue('VALEURTAXESIT',TVAC.GetValue('VALEURTAXESIT')+TVAS.GetValue('GPT_TOTALTAXEDEV1'));
    TVAC.putValue('BASETAXE',TVAC.GetValue('BASETAXE')+TVAS.getValue('GPT_TOTALHTDEV'));
    TVAC.putValue('BASETAXECUM',TVAC.GetValue('BASETAXECUM')+TVAS.getValue('GPT_TOTALHTDEV'));
    MtHtPortTTC := MtHtPortTTC + TVAS.getValue('GPT_TOTALHTDEV');
    MtPortTTC := MtPortTTC + TVAS.getValue('GPT_TOTALTTCDEV');
  end;
  // phase 2 déduction des tvas PRE
  for Indice := 0 to TOBTVAPRETTC.detail.count -1 do
  begin
    TVAS :=TOBTVAPRETTC.detail[indice];
    TVAC := TOBTVAPre.findFirst(['FAMILLETAXE'],[TVAS.getValue('FAMILLETAXE')],true);
    if TVAC <> nil then
    begin
      if fAvancement then TVAC.putValue('VALEURTAXEPRE',  TVAC.GetValue('VALEURTAXEPRE')-TVAS.GetValue('VALEURTAXEPRE'));
      if fAvancement then TVAC.putValue('BASETAXEPRE',    TVAC.GetValue('BASETAXEPRE')-TVAS.getValue('BASETAXEPRE'));
      TVAC.putValue('BASETAXECUM',TVAC.GetValue('BASETAXECUM')-TVAS.getValue('BASETAXECUM'));
      TVAC.putValue('VALEURTAXESIT',TVAC.GetValue('VALEURTAXESIT')-TVAS.GetValue('VALEURTAXESIT'));
      TVAC.putValue('BASETAXE',TVAC.GetValue('BASETAXE')-TVAS.GetValue('BASETAXE'));
    end;
  end;
  Indice := 0;
  repeat
    TVAS :=  TOBTVAPre.detail[Indice];
    if (not fTOBPiece.getBoolean('GP_AUTOLIQUID')) and (TVAS.GetDouble('VALEURTAXESIT')=0) and (TVAS.GetDouble('VALEURTAXEPRE')=0) then TVAS.free else Inc(Indice); 
  until Indice >= TOBTVAPRE.detail.count;
  TVASit.free;
end;

procedure TImprPieceViaTOB.DefiniLignesTVADet (ThePiece,TOBTVAPre,TOBFDetaille,TOBRDDET : TOB);

  procedure CalculeTvaAdeduire (TOBFDetaille : TOB; CodeFamille : string; var MTTVA,MTTVAPRE,BASETVA,BASETVAPRE : double);
  var II : Integer;
      TOBF : TOB;
  begin
    MTTVA := 0; MTTVAPRe := 0;  BASETVA := 0; BASETVAPRE := 0;
    for II := 0 TO TOBFDetaille.Detail.count -1 do
    begin
      TOBF := TOBFDetaille.detail[II];
      if not ((TOBF.GetString('TYPE')='HT') or
         (TOBF.GetString('TYPE')='MI') or
         (TOBF.GetString('TYPE')='MT')) then continue;
      if TOBF.GetString('FAMILLETAXE1')<>CodeFamille then Continue;
      MTTVA := MTTVA + TOBF.GetDouble('TVASIT');
      MTTVAPRE := MTTVAPRE + TOBF.GetDouble('TVAPRE');
      BASETVA := BASETVA + TOBF.GetDouble('MONTANTHTSIT');
      BASETVAPRE := BASETVAPRE + TOBF.GetDouble('MONTANTHTPRE');
    end;

  end;

var TOBL,TOBT : TOB;
    Indice : integer;
    LibTva : string;
    MTTVA,MTTVAPRE,reliqTvaSit,reliqtvapre,Base,BasePre : double;
    BaseTvaRDPre,MtTvaRDPre,BaseTvaRDSit,MTTVARDSIT : double;
begin
  for Indice := 0 to TOBTVAPRE.detail.count -1 do
  begin
    TOBT := TOBTVAPre.detail[Indice];
    GetInfosTVAFromRD(TOBT.GetString('FAMILLETAXE'),TOBRDDET,BaseTvaRDPre,MtTvaRDPre,BaseTvaRDSit,MTTVARDSIT);
    TOBL := InsereTOB ('LIGNE',ThePiece);
//    if TOBT.GetString('FAMILLETAXE')=getparamSocSecur('SO_CODETVALIQUIDST','') then
    if fTOBPiece.Getboolean('GP_AUTOLIQUID') then
    begin
        LibTva := 'AUTOLIQUIDATION DE TVA';
        TOBL.putValue('LIBELLE',libTva);
        TOBL.putValue('ISLIGNEPOR','X');
    end else
    begin
//      CalculeTvaAdeduire (TOBFDetaille, TOBT.getValue('FAMILLETAXE'), MTTVA,MTTVAPRE,Base,basePre);
    LibTva := 'TVA à '+StrF00(TOBT.getValue('TAUXTAXE'),2)+'%';
    if (fIsDGD) then
    begin
        LibTva := LibTva + ' - Base : '+StrF00(TOBT.getValue('BASETAXEPRE')-Basepre+BaseTvaRDPre, V_PGI.OkDecV)+ ' '+Symbole ;
    end else
    begin
        if (TOBTVAPRE.detail.count > 1) and (TOBT.getValue('BASETAXE')<>0) then
      begin
        LibTva := LibTva + ' - Base : '+StrF00(TOBT.getValue('BASETAXEPRE')-Basepre+BaseTvaRDPre, V_PGI.OkDecV)+ ' '+Symbole ;
      end else
      begin
        if (TOBTVAPRE.detail.count > 1) and (TOBT.getValue('BASETAXE')<>0) then
        begin
            LibTva := LibTva + ' - Base Sit. : '+StrF00(TOBT.getValue('BASETAXE')-Base+BaseTvaRDSit, V_PGI.OkDecV)+ ' '+Symbole ;
        end;
      end;
      ReliqTvaSit := Arrondi(TOBT.GetValue('VALEURTAXESIT')+MTTVARDSIT {- MTTVA},2);
      ReliqTvaPre := ARRONDI(TOBT.GetValue('VALEURTAXEPRE')+MtTvaRDPre {- MTTVAPRE},2);
      if (reliqTvaSit<>0) or (reliqTvapre<>0) then
      begin
        TOBL.putValue('LIBELLE',libTva);
        TOBL.putValue('MONTANTSIT',TOBT.GetValue('VALEURTAXESIT') - MTTVA + MTTVARDSIT);
        if fAvancement then TOBL.putValue('MONTANTPRE',TOBT.GetValue('VALEURTAXEPRE') - MTTVAPRE+MtTvaRDPre);
        TOBL.putValue('ISLIGNEPOR','X');
      end;
    end;
  	end;
	end;
end;

procedure TImprPieceViaTOB.GetInfosTVAFromRD(CodeTva : string; TOBRDDET : TOB; var BaseTvaRDPre,MtTvaRDPre,BaseTvaRDSit,MTTVARDSIT : double);
var II,JJ : integer;
    TOBDG,TOBDD  :TOB;
begin
  BaseTvaRDPre := 0;
  MtTvaRDPre := 0;
  BaseTvaRDSit := 0;
  MTTVARDSIT := 0;
  //
  for II := 0 to TOBRDDET.detail.count -1 do
  begin
    TOBDG := TOBRDDET.detail[II];
    if TOBDG.GEtString('TYPE')<>'HT' then continue;
    for JJ := 0 to TOBDG.detail.count -1 do
    begin
      TOBDD := TOBDG.detail[JJ];
      if TOBDD.GetString('FAMILLETAXE')<> CodeTva then continue;
      BaseTvaRDPre := BaseTvaRDPre + TOBDD.GetDouble('BASETAXEPRE');
      MtTvaRDPre := MtTvaRDPre + TOBDD.GetDouble('VALEURTAXEPRE');
      BaseTvaRDSit := BaseTvaRDSit + TOBDD.GetDouble('BASETAXE');
      MTTVARDSIT := MTTVARDSIT + TOBDD.GetDouble('VALEURTAXESIT');
    end;
  end;
end;


procedure TImprPieceViaTOB.DefiniLignesTVA (ThePiece,TOBTVAPre,TOBTVAPRETTC,TOBRDDET : TOB; var CumSitTva,CumPreTva: double);

var TOBL,TOBT : TOB;
    Indice : integer;
    LibTva : string;
    BaseTvaRDPre,MtTvaRDPre,BaseTvaRDSit,MTTVARDSIT : double;
begin
  CumSitTva := 0; CumPreTva := 0;
  for Indice := 0 to TOBTVAPRE.detail.count -1 do
  begin
    TOBT := TOBTVAPre.detail[Indice];
    TOBL := InsereTOB ('LIGNE',ThePiece);
    GetInfosTVAFromRD(TOBT.GetString('FAMILLETAXE'),TOBRDDET,BaseTvaRDPre,MtTvaRDPre,BaseTvaRDSit,MTTVARDSIT);
//    if TOBT.GetString('FAMILLETAXE')=getparamSocSecur('SO_CODETVALIQUIDST','') then
    if fTOBPiece.getBoolean('GP_AUTOLIQUID') then
    begin
        LibTva := 'AUTOLIQUIDATION DE TVA';
    end else
    begin
      LibTva := 'TVA à '+StrF00(TOBT.getValue('TAUXTAXE'),2)+'%';
      if (fIsDGD) then
      begin
        LibTva := LibTva + ' - Base : '+StrF00(TOBT.getValue('BASETAXEPRE')+BaseTvaRDPre, V_PGI.OkDecV)+ ' '+Symbole ;
      end else
      begin
        if (TOBTVAPRE.detail.count > 1) and (TOBT.getValue('BASETAXE')<>0) then
        begin
            LibTva := LibTva + ' - Base Sit. : '+StrF00(TOBT.getValue('BASETAXE')+BaseTvaRDSit, V_PGI.OkDecV)+ ' '+Symbole ;
        end;
      end;
      TOBL.putValue('MONTANTSIT',TOBT.GetValue('VALEURTAXESIT')+MTTVARDSIT); CumSitTva := CumSitTva + TOBT.GetValue('VALEURTAXESIT')+MTTVARDSIT;
      if fAvancement then TOBL.putValue('MONTANTPRE',TOBT.GetValue('VALEURTAXEPRE')+MtTvaRDPre); CumPreTva := CumPreTva + TOBT.GetValue('VALEURTAXEPRE')+MtTvaRDPre;
    end;
    TOBL.putValue('LIBELLE',libTva);
    TOBL.putValue('ISLIGNESTD','X');
  end;
end;

procedure TImprPieceViaTOB.DefiniLigneMontantHt (ThePiece : TOB; EscompteCum,RemiseCum,PortCum,CumPortSit,MTRDSITHT,MTRDPREHT : double);
var TOBL : TOB;
begin
  TOBL := InsereTOB ('LIGNE',ThePiece);
  TOBL.putValue('LIBELLE','TOTAL H.T. DES TRAVAUX');
  TOBL.putValue('MONTANTSIT',ThePiece.getValue('GP_TOTALHTDEV') + ThePiece.getValue('GP_TOTALREMISEDEV')+
                             ThePiece.getValue('GP_TOTALESCDEV')- CumPortSit + MTRDSITHT);

  TOBL.putValue('MONTANTPRE',ThePiece.getValue('TOTSITPRECBRUT')+MTRDPREHT);
  TOBL.putValue('ISLIGNESTD','X');
  TOBL.putValue('ISLIGNEPOR','X');
end;

procedure TImprPieceViaTOB.DefiniLignePIECEREPART (ThePiece : Tob; EscompteCum,RemiseCum,PortCum,CumPortSit,MTRDSITHT,MTRDPREHT : double);
Var Nature  : String;
    Souche  : String;
    Numero  : Integer;
    IndiceG : Integer;
    TOBE    : TOB;
    TOBL    : TOB;
    StSQL   : string;
    QQ      : TQuery;
    ind     : Integer;
    CumulSit: Double;
    CumulPre: Double;
begin

  //Lecture du fichier Piece repart en fonction de la piece en cours
  StSQL := 'SELECT * FROM BTPIECEREPART   ';
  StSQL := StSQL + ' WHERE BPR_NATUREPIECEG ="' + thePiece.GetValue('GP_NATUREPIECEG') + '" ';
  StSQL := StSQL + '   AND BPR_SOUCHE ="'       + ThePiece.GetValue('GP_SOUCHE')      + '" ';
  StSQl := StSQL + '   AND BPR_NUMERO = '       + InttoStr(thePiece.GetValue('GP_NUMERO'))      + '  ';
  StSQL := StSQL + '   AND BPR_INDICEG= '       + IntToStr(ThePiece.GetValue('GP_INDICEG'))     + '  ';
  StSQL := StSQL + ' ORDER BY BPR_AVENANT';

  QQ := OpenSQL(STSQL, False);

  if QQ.eof then
  begin
    DefiniLigneMontantHt (ThePiece,EscompteCum,RemiseCum,PortCum,CumPortSit,MTRDSITHT,MTRDPREHT);
    Ferme(QQ);
    Exit;
  end;          

  CumulSit := 0;
  CumulPre := 0;

  TOBE := TOB.Create('Les Pieces Réparties',nil,-1);
  TOBE.LoadDetailDB('BTPIECEREPART', '','',QQ,false);

  For ind := 0 to TOBE.Detail.count-1 do
  begin
    if TOBE.detail[ind].GetValue('BPR_AVENANT') = 0 then
    begin
      TOBL := InsereTOB ('LIGNE',ThePiece);
      //
      TOBL.putValue('LIBELLE','TOTAL HT MARCHE DE BASE');
      TOBL.putValue('MONTANTSIT', TOBE.detail[ind].GetValue('BPR_SITENCOURS'));
      TOBL.putValue('MONTANTPRE', TOBE.detail[ind].GetValue('BPR_SITPRECEDENTE'));
      TOBL.putValue('ISLIGNESTD','X');
      TOBL.putValue('ISLIGNEPOR','X');
    end
    else
    begin
      if ThePiece.GetValue('BPD_IMPCUMULAVE')='X' then
      begin
        CumulSit := Cumulsit + TOBE.detail[ind].GetValue('BPR_SITENCOURS');
        CumulPre := CumulPre + TOBE.detail[ind].GetValue('BPR_SITPRECEDENTE');
      end
      else
      begin
        TOBL := InsereTOB ('LIGNE',ThePiece);
        //
        TOBL.putValue('LIBELLE','AVENANT N° ' + TOBE.detail[ind].GetValue('BPR_AVENANT'));
        TOBL.putValue('MONTANTSIT', TOBE.detail[ind].GetValue('BPR_SITENCOURS'));
        TOBL.putValue('MONTANTPRE', TOBE.detail[ind].GetValue('BPR_SITPRECEDENTE'));
        TOBL.putValue('ISLIGNESTD','X');
        TOBL.putValue('ISLIGNEPOR','X');
      end;
    end;
  end;

  if (CumulSit <> 0) Or (CumulPre <> 0) then
  begin
    TOBL := InsereTOB ('LIGNE',ThePiece);
    //
    TOBL.putValue('LIBELLE','CUMUL DES AVENANTS');
    TOBL.putValue('MONTANTSIT', CumulSit);
    TOBL.putValue('MONTANTPRE', CumulPre);
    TOBL.putValue('ISLIGNESTD','X');
    TOBL.putValue('ISLIGNEPOR','X');
  end;

  Ferme(QQ);

  FreeAndNil(TOBE);

end;



procedure TImprPieceViaTOB.DefiniLigneTotalHtDet(ThePiece,TOBFDetaille: TOB);

  procedure CalculeHTFromTTC (TOBFD: TOB; var HT: double ; var HTPRE : double);
  var II : Integer;
      TOBD : TOB;
  begin
    HT := 0;
    HTPRE := 0;
    for II := 0 to TOBFD.detail.count -1 do
    begin
      TOBD := TOBFD.detail[II];
      if (TOBD.GetString('TYPE')='HT') or
         (TOBD.GetString('TYPE')='MI') or
         (TOBD.GetString('TYPE')='MT') then continue;
      HT := HT + TOBD.GetDouble('MONTANTHTSIT');
      HTPRE := HTPRE + TOBD.GetDouble('MONTANTHTPRE');
    end;
  end;

var TOBL : TOB;
    HT,HTPRE : Double;
begin
  TOBL := InsereTOB ('LIGNE',ThePiece);
  CalculeHTFromTTC (TOBFDetaille,HT,HTPRE);
  TOBL.putValue('LIBELLE','TOTAL H.T.');
  TOBL.putValue('MONTANTSIT',ThePiece.getValue('GP_TOTALHTDEV')-HT);
  TOBL.putValue('MONTANTPRE',ThePiece.getValue('TOTSITPREC')-HTPRE);
  TOBL.putValue('ISLIGNEPOR','X');
end;

procedure TImprPieceViaTOB.DefiniLigneTotalHt (ThePiece : TOB);
var TOBL : TOB;
begin
  TOBL := InsereTOB ('LIGNE',ThePiece);
  TOBL.putValue('LIBELLE','TOTAL H.T.');
  TOBL.putValue('MONTANTSIT',ThePiece.getValue('GP_TOTALHTDEV'));
  TOBL.putValue('MONTANTPRE',ThePiece.getValue('TOTSITPREC'));
  TOBL.putValue('ISLIGNESTD','X');
end;

procedure TImprPieceViaTOB.DefiniLigneRemise(ThePiece: TOB; RemiseCum: double);
var TOBL : TOB;
begin
  if (remiseCum = 0) and (ThePiece.getValue('GP_TOTALREMISEDEV')=0) then exit;
  TOBL := InsereTOB ('LIGNE',ThePiece);
  TOBL.putValue('LIBELLE','REMISE SUR DEVIS DE BASE');
  TOBL.putValue('MONTANTSIT',ThePiece.getValue('GP_TOTALREMISEDEV'));
  TOBL.putValue('MONTANTPRE',ThePiece.getValue('TOTREMPREC'));
  TOBL.putValue('ISLIGNESTD','X');
  TOBL.putValue('ISLIGNEPOR','X');
end;

procedure TImprPieceViaTOB.DefiniLigneEscompte(ThePiece: TOB;EscompteCum: double);
var TOBL : TOB;
begin
  if (EscompteCum = 0) and (ThePiece.getValue('GP_TOTALESCDEV')=0) then exit;
  TOBL := InsereTOB ('LIGNE',ThePiece);
  TOBL.putValue('LIBELLE','ESCOMPTE SUR DEVIS DE BASE');
  TOBL.putValue('MONTANTSIT',ThePiece.getValue('GP_TOTALESCDEV'));
  if fAvancement then TOBL.putValue('MONTANTPRE',ThePiece.getValue('TOTESCPREC'));
  TOBL.putValue('ISLIGNESTD','X');
  TOBL.putValue('ISLIGNEPOR','X');
end;

procedure TImprPieceViaTOB.DefiniLignePortFrais(ThePiece: TOB;PortCum,CumPortSit: double);
var TOBL : TOB;
    Req : string;
    QQ : TQuery;
    CumHt,CumTTc : Double;
begin
  if (PortCum = 0) and (CumPortSit=0) then exit;
  TOBL := InsereTOB ('LIGNE',ThePiece);
  TOBL.putValue('LIBELLE','PORT & FRAIS SUR DEVIS DE BASE');
  TOBL.putValue('MONTANTSIT',CumPortSit);
  if fAvancement then TOBL.putValue('MONTANTPRE',ThePiece.getValue('TOTPORTPREC'));
  TOBL.putValue('ISLIGNESTD','X');
end;

procedure TImprPieceViaTOB.DefiniLignePortFraisDet(ThePiece,TOBFraisDet :TOB; TypePort : string ; var NbrPorts: integer; sens : string='+');
var TOBL,TOBF : TOB;
    II : Integer;
    ISens : integer;
begin
  NbrPorts := 0;
  if Sens = '+' then ISens := 1 else iSens := (-1);
  if TOBFraisDet.detail.count = 0 then Exit;
  for II := 0 to TOBFraisDet.Detail.Count -1 do
  begin
    TOBF := TOBFraisDet.detail[II];
    if (TOBF.GetString('TYPE')<>'HT') and
       (TOBF.GetString('TYPE')<>'MI') and
       (TOBF.GetString('TYPE')<>'MT') and
       (TypePort='HT') then continue;

    if (TOBF.GetString('TYPE')<>'PT') and
       (TOBF.GetString('TYPE')<>'MIC') and
       (TOBF.GetString('TYPE')<>'MTC') and
       (TypePort='TTC') then continue;
    //
    Inc(NbrPorts);
    TOBL := InsereTOB ('LIGNE',ThePiece);
    TOBL.putValue('LIBELLE',TOBF.GetString('LIBELLE'));
    if TypePort = 'HT' then
    begin
      TOBL.putValue('MONTANTSIT',TOBF.GetDouble('MONTANTHTSIT')*Isens);
      if fAvancement then TOBL.putValue('MONTANTPRE',TOBF.GetDouble('MONTANTHTPRE')*Isens);
    end else
    begin
      TOBL.putValue('MONTANTSIT',TOBF.GetDouble('MONTANTTTCSIT')*Isens);
      if fAvancement then TOBL.putValue('MONTANTPRE',TOBF.GetDouble('MONTANTTTCPRE')*ISens);
    end;
    TOBL.putValue('ISLIGNEPOR','X');
    TOBL.putValue('ISLIGNESTD','X');
    (*
    // Ligne de Tva associé
    if (TypePort = 'TTC') then
    begin
      TOBL := InsereTOB ('LIGNE',ThePiece);
      TOBL.putValue('LIBELLE','Dont TVA '+ RechDom('GCFAMILLETAXE1',TOBF.GetString('FAMILLETAXE1'),false));
      TOBL.putValue('MONTANTSIT',TOBF.GetDouble('MONTANTTTCSIT')-TOBF.GetDouble('MONTANTHTSIT'));
      TOBL.putValue('MONTANTPRE',TOBF.GetDouble('MONTANTTTCPRE')-TOBF.GetDouble('MONTANTHTPRE'));
      TOBL.putValue('ISLIGNEPOR','X');
    end;
    *)
  end;
end;

procedure TImprPieceViaTOB.GestionAnciennesfactureation(TOBL: TOB; Modegenere : string; DEV : RDevise);

	procedure  GetInfoMarche (TOBL : TOB; var Mtmarche,QteMarche :double);
  var QQ : TQuery;
  		Req : String;
      Cledoc : r_cledoc;
  begin
  	if TOBL.GetValue('GL_PIECEORIGINE')='' then exit;
  	DecodeRefPiece (TOBL.GetValue('GL_PIECEORIGINE'),Cledoc);
  	req := 'SELECT GL_MONTANTHTDEV,GL_QTEFACT FROM LIGNE WHERE '+
    			 WherePiece(Cledoc,ttdLigne ,true,true);
  	QQ := OpenSql (Req,true,1,'',true);
    if not QQ.Eof then
    begin
    	MtMarche := QQ.findField('GL_MONTANTHTDEV').AsFloat;
    	QteMarche := QQ.findField('GL_QTEFACT').AsFloat;
    end;
    Ferme (QQ);
  end;

var MtDejafac,QteDejaFact,MtHtDev,PQ,PourcentAvanc : double;
		Mtmarche,Qtemarche : double;
begin
  if TOBL.getValue('GL_TYPELIGNE')<>'ART' then exit;
  //
//  Modif BRL 18/07/2014 : if (TOBL.GetValue('BLF_MTCUMULEFACT')=0) and ((TOBL.GetValue('GL_QTEPREVAVANC')<>0) or (TOBL.GetValue('GL_QTESIT')<>0)) then
  if (TOBL.GetValue('BLF_MTCUMULEFACT')=0) and (TOBL.GetValue('GL_QTESIT')<>0) then
  begin
    PQ := TOBL.GEtValue('GL_PRIXPOURQTE'); if PQ = 0 then PQ := 1;
    MtHtDev := TOBL.getValue('GL_MONTANTHTDEV');
    if MtHtdev = 0 then MtHtdev := arrondi(TOBL.getValue('GL_QTEFACT')*TOBL.GEtValue('GL_PUHTDEV') / PQ,DEV.decimale);
    MtMarche := 0; Qtemarche := 0;
    GetInfoMarche (TOBL,Mtmarche,QteMarche);
    if Modegenere='AVA' then
    begin
      // avancement ou demande d'acompte
      QteDejaFact := TOBL.GetValue('GL_QTEPREVAVANC');
      MtDejaFac := Arrondi(QteDejaFact* TOBL.GEtValue('GL_PUHTDEV') / PQ,DEV.decimale);
      if Mtmarche <> 0 then PourcentAvanc := Arrondi(MtDejaFac / Mtmarche * 100,2) else PourcentAvanc := 0;
      TOBL.putValue('BLF_MTMARCHE',MtMarche);
      TOBL.putValue('BLF_MTDEJAFACT',MtDejaFac);
      TOBL.putValue('BLF_MTCUMULEFACT',MtDejaFac);
      TOBL.putValue('BLF_MTSITUATION',0);
      TOBL.putValue('BLF_QTEMARCHE',QteMarche);
      TOBL.putValue('BLF_QTEDEJAFACT',QteDejaFact);
      TOBL.putValue('BLF_QTECUMULEFACT',QteDejaFact);
      TOBL.putValue('BLF_QTESITUATION',0);
      TOBL.putValue('BLF_POURCENTAVANC',PourcentAvanc);
      if TOBL.FieldExists('POURCENTAVANCINIT') then TOBL.putValue ('POURCENTAVANCINIT',PourcentAvanc);
      if TOBL.FieldExists('MTCUMULEFACTINIT')  then TOBL.putValue ('MTCUMULEFACTINIT',MtDejaFac);
      if TOBL.FieldExists('QTECUMULEFACTINIT') then TOBL.putValue ('QTECUMULEFACTINIT',QteDejaFact);
    end else
    begin
      // facturation directe ou hors devis
      QteDejaFact := TOBL.GetValue('GL_QTESIT');
      MtDejaFac := Arrondi(QteDejaFact* TOBL.GEtValue('GL_PUHTDEV') / PQ,DEV.decimale);
      if MtMarche <> 0 then PourcentAvanc := Arrondi(MtDejaFac / MtMarche * 100,2) else PourcentAvanc := 0;
      TOBL.putValue('BLF_MTMARCHE.',MtMarche);
      TOBL.putValue('BLF_MTDEJAFACT',MtDejaFac);
      TOBL.putValue('BLF_MTCUMULEFACT',0);
      TOBL.putValue('BLF_MTSITUATION',0);
      TOBL.putValue('BLF_QTEMARCHE',QteMarche);
      TOBL.putValue('BLF_QTEDEJAFACT',QteDejaFact);
      TOBL.putValue('BLF_QTECUMULEFACT',0);
      TOBL.putValue('BLF_QTESITUATION',0);
      TOBL.putValue('BLF_POURCENTAVANC',0);
      TOBL.putValue ('POURCENTAVANCINIT',PourcentAvanc);
      TOBL.putValue ('MTCUMULEFACTINIT',MtDejaFac);
      TOBL.putValue ('QTECUMULEFACTINIT',QteDejaFact);
    end;
  end;
end;

procedure TImprPieceViaTOB.AlimentemetresLig(ThePiece,TheLigne: TOB; Numordre : integer; TOBmetres: TOB);
var TOBM,TOBMD,TOBL : TOB;
		indice : integer;
    TheOrdre : integer;
begin
  TOBM := TOBMetres.findFirst(['NUMORDRE'],[numordre],true);
  if TOBM <> nil then
  begin
  	TOBL := InsereTOB ('LIGNE',ThePiece);
    AjouteTOB (TOBL,'LIGNES',TheLigne);
    TOBL.PutValue ('ISDEBMETRE','X');
    TOBL.PutValue ('ISLIGNESTD','-');
  	TOBL.PutValue('DESCRIPTIF','-');
    TOBL.PutValue('GL_NUMORDRE',NumOrdre); // retablissement
    for Indice := 0 to TOBM.detail.count -1 do
    begin
      TOBMD := TOBM.detail[Indice];
      TOBL := InsereTOB ('LIGNE',ThePiece);
    	AjouteTOB (TOBL,'LIGNES',TheLigne);
      TOBL.PutValue('BLM_CHAINE',TOBMD.GetValue('BLM_CHAINE'));
      TOBL.PutValue('BLM_OPERATION',TOBMD.GetValue('BLM_OPERATION'));
      TOBL.PutValue('BLM_VALEUR',TOBMD.GetValue('BLM_VALEUR'));
      TOBL.PutValue('GL_NUMORDRE',NumOrdre); // retablissement
    	TOBL.PutValue ('ISLIGMETRE','X');
    	TOBL.PutValue ('ISLIGNESTD','-');
  		TOBL.PutValue('DESCRIPTIF','-');
    end;
  	TOBL := InsereTOB ('LIGNE',ThePiece);
    AjouteTOB (TOBL,'LIGNES',TheLigne);
    TOBL.PutValue ('ISFINMETRE','X');
    TOBL.PutValue ('ISLIGNESTD','-');
  	TOBL.PutValue('DESCRIPTIF','-');
    TOBL.PutValue('GL_NUMORDRE',NumOrdre); // retablissement
  end;
end;

function TImprPieceViaTOB.ImpDecodeRib(RibEncode: string): string;
var Etab,Guichet,Numero,Cle,Dom : String;
begin
  if RibEncode='' then Exit;
  DecodeRIB(Etab,Guichet,Numero,Cle,Dom,RibEncode);
  Result := 'Code Banque : '+Etab+' Domiciliation :'+trim(Dom)+' Guichet : '+Guichet+' Compte : '+Numero+' Clé : '+Cle;
end;

procedure TImprPieceViaTOB.VirelesLignesAZero(Tobfacture: TOB);
var II : integer;
    TOBL : TOB;
begin
  II := 0;
  Repeat
    TOBL := TOBFacture.detail[II];
    if TOBL.GetValue('GL_TYPELIGNE')='ART' then
    begin
      if (TOBL.GetValue('BLF_MTDEJAFACT')=0) and (TOBL.GetValue('BLF_MTSITUATION')=0) then
      begin
        TOBL.free;
        continue;
      end;
    end;
    inc(II);
  Until II>= TOBfacture.detail.count;
end;

procedure TImprPieceViaTOB.SuprimeParagraphesOrphelin(TOBFacture: TOB);

  procedure SuprimeThisParagraphe(TOBL : TOB; var IndDep: integer);
  var Niveau : integer;
      TOBI : TOB;
      StopItNow : boolean;
  begin
    Niveau := TOBL.GetValue('GL_NIVEAUIMBRIC');
    StopItNow := false;
    repeat
      TOBI := TOBFacture.detail[IndDep];
      if IsDebutParagraphe (TOBI,Niveau) then StopItNow := true;
      TOBI.free;
      Dec(IndDep);
    until (IndDep = 0 ) or (StopItNow);
  end;


  function IsDetailInside(IndDep: integer; TOBL: TOB): boolean;
  var indice : integer;
      Niveau : integer;
      TOBI : TOB;
  begin
    result := false;
    Niveau := TOBL.GetValue('GL_NIVEAUIMBRIC');
    for Indice := Inddep-1 downto 0 do
    begin
      TOBI := TOBFacture.detail[Indice];
      if IsArticle (TOBI) then BEGIN result := true; break; END;
      if IsDebutParagraphe (TOBI,Niveau) then break;
    end;
  end;

var Indice : integer;
    TOBL : TOB;
begin
  Indice := TOBFacture.detail.count -1;
  repeat
    if Indice >= 0 then
    begin
      TOBL := TOBFacture.detail[Indice];
      if IsFinParagraphe(TOBL) then
      begin
        if not IsDetailInside (Indice,TOBL) then
        begin
          SuprimeThisParagraphe (TOBL,Indice);
        end;
      end;
      dec(Indice);
    end;
  until indice <= 0;
end;
{***********A.G.L.***********************************************
Auteur  ...... : FV
Créé le ...... : 15/06/2016
Modifié le ... :   /  /    
Description .. : Impression recapitulatif paiement Ss-Traitance
Mots clefs ... :
*****************************************************************}
procedure TImprPieceViaTOB.AjouteChampsSsTraitantInterne(TOBmere, TOBAInsere: TOB);
begin

  if TOBAinsere = Nil Then Exit;

  if fNbSTTraitDir <= 5 then
    AjouteSsTraitantInterne(Tobmere, TOBAinsere)
  else
    AjouteTobSsTraitantInterne(Tobmere, TOBAinsere);

end;

Procedure TimprPieceViaTob.AjouteChampsSsTraitant(Tobmere, TOBAinsere : TOB);
begin

  if TOBAinsere = Nil Then Exit;

  if fNbSTTraitDir <= 5 then
    AjouteSsTraitant(Tobmere, TOBAinsere)
  else
    AjouteTobSsTraitant(Tobmere, TOBAinsere);

end;

Procedure TimprPieceViaTob.AjouteSsTraitantInterne(Tobmere, TOBAinsere : TOB);
var num       : integer;
    TOBI      : TOB;
    TOBL      : TOB;
    I         : Integer;
    TotSsTrait : Double;
begin

	TotSsTrait := 0;

  i := 1;

  For Num := 0 To TOBAinsere.detail.count -1 do
  begin
    TOBL := TOBAinsere.detail[Num];
    if (TOBL.GetString('BPE_TYPEINTERV') = 'Y00') and (TOBL.GetString('TYPEPAIE')='001')then
    begin
      TOBMEre.SetString ('SSTRAITANCE','X');
      inc(I);
      TOBMEre.PutValue ('SST_FOURNISSEUR'+inttoStr(I),TOBL.GetString('LIBELLE'));
      TOBMEre.PutValue ('SST_TOTALHTDEV' +inttoStr(I),TOBL.GetDouble('BPE_TOTALHTDEV'));
      TOBMEre.PutValue ('SST_TOTALTTCDEV'+inttoStr(I),TOBL.GetDouble('BPE_TOTALTTCDEV'));
      TOBMEre.PutValue ('SST_MONTANTPA'  +inttoStr(I),TOBL.GetDouble('BPE_MONTANTPA'));
      TOBMEre.PutValue ('SST_MONTANTFAC' +inttoStr(I),TOBL.GetDouble('BPE_MONTANTFAC'));
      TOBMEre.PutValue ('SST_MONTANTREGL'+inttoStr(I),TOBL.GetDouble('BPE_MONTANTREGL'));
      TotSsTrait  := TotSsTrait + TOBL.GetDouble('BPE_MONTANTREGL');
    end;
  end;

  TOBMEre.PutValue ('TOTSSTRAIT', TotSsTrait);

  if i > 1 then
  begin
    TOBI := TOBAInsere.findFirst(['BPE_FOURNISSEUR'],[''],True);
    if TOBI <> nil then
    begin
      TOBMEre.PutValue ('SST_FOURNISSEUR1',TOBI.GetString('LIBELLE'));
      TOBMEre.PutValue ('SST_TOTALHTDEV1' ,TOBI.GetDouble('BPE_TOTALHTDEV'));
      TOBMEre.PutValue ('SST_TOTALTTCDEV1',TOBI.GetDouble('BPE_TOTALTTCDEV'));
      TOBMEre.PutValue ('SST_MONTANTPA1'  ,TOBI.GetDouble('BPE_MONTANTPA'));
      TOBMEre.PutValue ('SST_MONTANTFAC1' ,TOBI.GetDouble('BPE_MONTANTFAC'));
      TOBMEre.PutValue ('SST_MONTANTREGL1',TOBI.GetDouble('BPE_MONTANTREGL'));
    end;
  end;

end;

Procedure TimprPieceViaTob.AjouteTobSsTraitantInterne(Tobmere, TOBAinsere : TOB);
Var TOBI      : TOB;
    TOBL      : TOB;
    num       : integer;
    NoLig     : Integer;
    TotSsTrait: Double;
begin

  TotSsTrait := 0;

  NoLig := 1;

  For Num := 0 To TOBAinsere.detail.count -1 do
  begin
    TOBL := TOBAinsere.Detail[Num];
    if (TOBL.GetString('BPE_TYPEINTERV') = 'Y00') and (TOBL.GetString('TYPEPAIE')='001') then
    begin
      TOBI := TOB.Create('Les SSTRAITANTS', Tobmere, -1);
      TOBI.Dupliquer(TOBL, True, true);
      Inc(NoLig);
      TOBI.AddChampSupValeur('LIG', NoLig);
      TOBI.AddChampSupValeur('LIBTIERS', TOBL.getString('LIBELLE'));
      TotSsTrait  := TotSsTrait + TOBL.GetDouble('BPE_MONTANTREGL');
    end;
  end;

  TOBMEre.AddChampSupValeur ('TOTSSTRAIT', TotSsTrait);

  if (TobMere.Detail.count > 1) then
  begin
    TOBL := TOBAinsere.FindFirst(['BPE_FOURNISSEUR'],[''],True);
    if (TOBL <> nil)  then
    begin
      TOBI := TOB.Create('Les SSTRAITANTS', TobMere, -1);
      TOBI.Dupliquer(TOBL, True, true);
      TOBI.AddChampSupValeur('LIG', 1);
      TOBI.AddChampSupValeur('LIBTIERS',' Paiement entreprise ');
//      TOBI.PutValue ('BPE_FOURNISSEUR',' Paiement entreprise ');
    end;
  end;

end;

Procedure TimprPieceViaTob.AjouteSsTraitant(Tobmere, TOBAinsere : TOB);
Var i         : Integer;
    TOBI      : TOB;
    num       : integer;
    TotSsTrait: Double;
begin

  i := 1;

  TotSsTrait := 0;

  For Num := 0 To TOBAinsere.detail.count -1 do
  begin
    TOBI := TOBAinsere.detail[Num];
    if (TOBI.GetString('BPE_TYPEINTERV') = 'Y00') and (TOBI.GetString('REGSSTRAIT')='001') then
    begin
      TOBMEre.AddChampSupValeur ('SSTRAITANCE','X');
      inc(I);
      TOBMEre.PutValue ('SST_FOURNISSEUR'+inttoStr(I),TOBI.GetString('LIBTIERS'));
      TOBMEre.PutValue ('SST_TOTALHTDEV' +inttoStr(I),TOBI.GetDouble('BPE_TOTALHTDEV'));
      TOBMEre.PutValue ('SST_TOTALTTCDEV'+inttoStr(I),TOBI.GetDouble('BPE_TOTALTTCDEV'));
      TOBMEre.PutValue ('SST_MONTANTPA'  +inttoStr(I),TOBI.GetDouble('BPE_MONTANTPA'));
      TOBMEre.PutValue ('SST_MONTANTFAC' +inttoStr(I),TOBI.GetDouble('BPE_MONTANTFAC'));
      TOBMEre.PutValue ('SST_MONTANTREGL'+inttoStr(I),TOBI.GetDouble('BPE_MONTANTREGL'));
      TotSsTrait  := TotSsTrait + TOBI.GetDouble('BPE_MONTANTREGL');
    end;
  end;

  TOBI := TOBAinsere.FindFirst(['BPE_FOURNISSEUR'],[''],True);
  if (TOBI <> nil) and (i>1) then
  begin
    TOBMEre.PutValue ('SST_FOURNISSEUR1',' Paiement Entreprise ');
    TOBMEre.PutValue ('SST_TOTALHTDEV1' ,TOBI.GetDouble('BPE_TOTALHTDEV'));
    TOBMEre.PutValue ('SST_TOTALTTCDEV1',TOBI.GetDouble('BPE_TOTALTTCDEV'));
    TOBMEre.PutValue ('SST_MONTANTPA1'  ,TOBI.GetDouble('BPE_MONTANTPA'));
    TOBMEre.PutValue ('SST_MONTANTFAC1' ,TOBI.GetDouble('BPE_MONTANTFAC'));
    TOBMEre.PutValue ('SST_MONTANTREGL1',TOBI.GetDouble('BPE_MONTANTREGL'));
    inc(I);
  end;

  TOBMEre.PutValue ('TOTSSTRAIT', TotSsTrait);

end;

Procedure TimprPieceViaTob.AjouteTobSsTraitant(Tobmere, TOBAinsere : TOB);
Var TOBI      : TOB;
    TOBL      : TOB;
    num       : integer;
    TotSsTrait: Double;
    NoLig     : Integer;
begin

  TotSsTrait := 0;

  NoLig := 1;

  For Num := 0 To TOBAinsere.detail.count -1 do
  begin
    TOBL := TOBAinsere.Detail[Num];
    if (TOBL.GetString('BPE_TYPEINTERV') = 'Y00') and (TOBL.GetString('REGSSTRAIT')='001') then
    begin
      TOBI := TOB.Create('Les SSTRAITANTS', Tobmere, -1);
      TOBI.Dupliquer(TOBL, True, true);
      inc(NoLig);
      TOBI.AddChampSupValeur('LIG', NoLig);
      TotSsTrait  := TotSsTrait + TOBL.GetDouble('BPE_MONTANTREGL');
    end;
  end;

  TOBMEre.AddChampSupValeur ('TOTSSTRAIT', TotSsTrait);

  if (TobMere.Detail.count > 1) then
  begin
    TOBL := TOBAinsere.FindFirst(['BPE_FOURNISSEUR'],[''],True);
    if (TOBL <> nil)  then
    begin
      TOBI := TOB.Create('Les SSTRAITANTS', Tobmere, -1);
      TOBI.Dupliquer(TOBL, True, true);
      TOBI.AddChampSupValeur('LIG', 1);
      TOBI.PutValue ('BPE_FOURNISSEUR',' Paiement Entreprise ');
    end;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : FV
Créé le ...... : 29/06/2016
Modifié le ... :   /  /
Description .. : Impression recapitulatif paiement Ss-Traitance
Mots clefs ... :
*****************************************************************}
procedure TimprPieceViaTob.AjouteRetenuesDiverses(TOBMere, TobAInsere : TOB);
Var II    : Integer;
    III   : Integer;
    IV    : Integer;
    TOBL  : Tob;
    TOB1  : Tob;
    TOB2  : Tob;
    TOBRD : Tob;
    TOBI  : TOB;
    NbFDHt: Integer;
begin

  if TOBAinsere = Nil Then Exit;

  NbFDHt := 0;

  For II := 0 to TobAInsere.Detail.count -1 do
  begin
    TOBL := TobAInsere.Detail[II];
    If TOBL.GetValue('GPT_RETENUEDIVERSE') = '-' then continue;
    //If TOBL.GetValue('GPT_TYPEINTER') = 'Y00' then continue;
    TOBRD:= TOB.Create('RETENUE DIVERSE', nil, -1);
    ConstitueRD (TOBMere,TOBRD,NbFDHt); // continue les retenues diverses ainsi que leurs tva associés (cumulés)
    For III := 0 To TOBRD.Detail.count -1 do
    begin
      TOB1 := TOBRD.Detail[III];
      TOBI := TOB.Create('RETENUE DIVERSE', Tobmere, -1);
      TOBI.Dupliquer(TOB1, False, true);
      For IV := 0 TO TOB1.Detail.count -1 do
      Begin
        TOB2 := TOB1.Detail[IV];
        TOBI := TOB.Create('RETENUE DIVERSE', Tobmere, -1);
        TOBI.Dupliquer(TOB1, False, true);
        TOBI.PutValue('CODE', 'TVA' + IntToStr(IV));
        TOBI.PutValue('TYPE', TOB1.GetValue('TYPE'));
        TOBI.PutValue('LIBELLE', 'TVA sur ' + TOB1.GetValue('LIBELLE'));
        TOBI.PutValue('FAMILLETAXE1', TOB2.GetValue('FAMILLETAXE'));
        TOBI.PutValue('MONTANTHTPRE', TOB2.GetValue('BASETAXEPRE'));
        TOBI.PutValue('MONTANTTTCPRE', 0);
        TOBI.PutValue('TVAPRE', TOB2.GetValue('VALEURPRE'));
        TOBI.PutValue('MONTANTHTSIT', TOB2.GetValue('VALEURTAXESIT'));
        TOBI.PutValue('MONTANTTTCSIT', 0);
        TOBI.PutValue('TVASIT', 0);
      end;
    end;    
  end;

end;


procedure TImprPieceViaTOB.AjouteRetenuesTTC(TOBMere, TOBAInsere: TOB);
var RD,RP : double;
    Req : string;
    QQ : Tquery;
begin
  GetRetenuesTTC(TOBAinsere,RP,RD);
  TOBMere.SetDouble('RETENUESTTC',RD);
end;

procedure TImprPieceViaTOB.AjouteRetenuesDiv(TOBMere, TOBAInsere: TOB);
var RD,RP : double;
    Req : string;
    QQ : Tquery;
begin
  GetRetenuesCollectif(TOBAinsere,RP,RD,'TTC');
  TOBMere.SetDouble('RETENUESDIVTTC',RD);
  GetRetenuesCollectif(TOBAinsere,RP,RD,'HT');
  TOBMere.SetDouble('RETENUESDIVHT',RD);
  if fAvancement then
  begin
    req := 'SELECT SUM(GPT_TOTALHTDEV) AS PORTPREC FROM PIEDPORT '+
         'WHERE GPT_NATUREPIECEG="'+TOBMere.GetValue('GP_NATUREPIECEG')+'" '+
         'AND GPT_SOUCHE="'+TOBMere.GetValue('GP_SOUCHE')+'" '+
         'AND GPT_NUMERO IN ('+
         'SELECT BST_NUMEROFAC FROM BSITUATIONS WHERE '+
         'BST_SSAFFAIRE="'+TOBMere.GetValue('GP_AFFAIREDEVIS')+'" AND '+
         'BST_NATUREPIECE="'+TOBMere.GetValue('GP_NATUREPIECEG')+'" AND '+
         'BST_SOUCHE="'+TOBMere.GetValue('GP_SOUCHE')+'" AND '+
         'BST_VIVANTE="X" AND '+
         'BST_NUMEROSIT < '+TOBMere.getString('NUMSIT')+') AND '+
         'GPT_RETENUEDIVERSE="X" AND GPT_TYPEPORT IN ("HT","MI","MT")';
    QQ := OpenSql (req,True,-1,'',true);
    if not QQ.eof then
    begin
      TOBMere.SetDouble('RETENUEHTPRE',QQ.fields[0].AsFloat);
    end;
    ferme (QQ);
  end;

end;

end.

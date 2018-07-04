unit GestRemplArticle;

interface
uses sysutils,classes,windows,messages,controls,forms,hmsgbox,stdCtrls,clipbrd,nomenUtil,
     HCtrls,SaisUtil,HEnt1,Ent1,EntGC,UtilPGI,UTOB,HTB97,FactUtil,FactComm,Menus,ParamSoc,
     AglInit,AglInitGc,FactTob,FactVariante,vierge,FactOuvrage,Graphics,FactLigneBase,
{$IFDEF EAGLCLIENT}
  maineagl,
{$ELSE}
  fe_main,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HRichOLE,UTOF,UEntCommun;

const MAXITEMS = 2;

type

  TOF_BTREPLACEART = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	TOBL : TOB;
    RefUniqueChg,RefUniqueAChg : string;
    procedure Recharticle (sender : Tobject);
    procedure QueSurLigne (Sender : TObject);
    procedure PasQueSurLigne (Sender : TObject);
    procedure DetermineEnabledFac;
    procedure DetermineEnabledOuv;
    procedure ChoixArtligne (Sender : TObject);
    procedure ChoixArticlePart (sender : Tobject);
    procedure ARecharticle(sender: Tobject);
    function VerifArticle(NomChamp, NOMChampLib: string;var LaRefUnique: string): boolean;
    procedure VerificationArtAreplace (Sender : TObject);
    procedure VerificationArtreplacement (Sender : Tobject);
  end ;


	Tgestremplarticle = class
  private
    TOBMemorisation : TOB;
    fCreatedPop : boolean;
    FF : TForm;
    fusable : boolean;
    fgrid : THGrid;
    POPGS : TPopupMenu;
    TOBPiece : TOB;
    TOBL : TOB;
    TOBBases,TOBBasesL : TOB;
    TOBTiers : TOB;
    TOBAffaire : TOB;
    TOBArticles : TOB;
    TOBOUvrage : TOB;
    TOBOUvragesP : TOB;
    TOBpieceTrait : TOB;
    fTOBSSTRAIT : TOB;
    DEV : RDevise;
    MesMenuItem: array[0..MAXITEMS] of TMenuItem;
    fMaxItems : integer;
    iWarnings : integer;
    procedure DefiniMenuPop(Parent: Tform);
    procedure EchangeLigne(TOBLS, TheTOBParam: TOB;WithCalcul: boolean=false);
    procedure EchangeParagraphe (TheTOBParam : TOB);
    procedure EchangeSelection (TheTOBParam : TOB);
    procedure EchangeDocument (TheTOBParam : TOB);
    procedure GSremplArticle (Sender : Tobject);
    procedure Setdocument(const Value: TOB);
    procedure SetGrid(const Value: THGrid); virtual;
    procedure SetTiers(const Value: TOB);
    procedure SetAffaire(const Value: TOB);
    procedure SetLigne(const Value: TOB);
//    procedure SetEnabled (const value : boolean);
    procedure AddChampSupEchg(UneTOB: TOB);
    procedure RemplacementArticle(UneTOBechg: TOB);
    procedure SetArticles(const Value: TOB);
    procedure ReinitLigneDoc(TOBLigne: TOB);
    function PositionneArticle(TheTOBParam: TOB): TOB;
    procedure RetrouveParagraphe (var Ligdepart,LigFin : integer ; TOBLigne : TOB);
    procedure SetOuvrage(const Value: TOB);
    procedure EchangeSurOuvrage(TheArticle : string; TOBLS, TOBOUV, TheTOBParam: TOB);
    procedure ReplaceArticleDetailOuvrage(TOBLS, TOBO, TheTOBParam: TOB);
    procedure Reinitialise(TOBO: TOB);
    procedure AffecteArticleSurOUv(TOBA,TOBO: TOB);
    procedure TraitementLigne(TOBLS, TheTOBparam: TOB);
    procedure AffichageWarnings;
    procedure MemoriseLigneOuv (TOBL : TOB);
    procedure RaffraichitDetailOuvrage;
    function GetTOBLigneViaOrdre (NumOrdre : integer) : TOB;
    procedure SetBases(const Value: TOB);
    procedure SetBasesL(const Value: TOB);
    procedure SSEnabled(etat: boolean);
    procedure DeduitlaLigne (TOBLS : TOB);
  public
    constructor create (TT : TForm);
    destructor  destroy ; override;
    property    Grid : THGrid write setGrid ;
    property    Piece : TOB read TobPiece write Setdocument;
    property    Ligne : TOB read TobL write SetLigne;
    property    Tiers : TOB read TobTiers write SetTiers;
    property    Affaire : TOB read Tobaffaire write SetAffaire;
    property    Articles : TOB read TobArticles write SetArticles;
    property    PiedBase : TOB read TobBases write SetBases;
    property    LigneBase : TOB read TobBasesL write SetBasesL;
    property    PieceTrait : TOB read TobPieceTrait write TobPieceTrait;
    property    Ouvrage : TOB read TobOuvrage write SetOuvrage;
    property    OuvrageP : TOB read TobOuvragesP write TOBOuvragesP;
    property    TOBSSTRAIT : TOB read fTOBSSTRAIT write fTOBSSTRAIT;
    procedure    SetEnabled;
end;

implementation
uses facture,UtilArticle,FactArticle,FactCalc,FactureBtp,UCotraitance, BTPUtil,FactPiece;

var FFactGestRempl : TFFacture;



function OkTYpeArticle(TypeArticle: string): boolean;
begin
	result := (Pos(TypeArticle,'MAR;PRE;POU')>0);
end;

function verifselection: boolean;
var Indice : integer;
		Article : string;
    TOBL : TOB;
begin
	Article := '';
  result := true;
  with FFactGestRempl do
  begin
    for Indice := 1 to GS.RowCount do
    begin
      if (GS.IsSelected (Indice) ) then
      begin
        TOBL := GetTOBLigne(LaPieceCourante,Indice);
        if not OkTYpeArticle(TOBL.getValue('GL_TYPEARTICLE')) then BEGIN result := false; break; END;
      end;
    end;
  end;
end;

{ Tgestremplarticle }

constructor Tgestremplarticle.create(TT: TForm);
var ThePop : Tcomponent;
begin
  TOBMemorisation := TOB.Create ('LES LIGNES',nil,-1);
  //
  FF := TT;
  fusable := true;
{$IFDEF BTP}
  if FF is TFFacture then ThePop := TFFacture(TT).Findcomponent  ('POPBTP') else thePop := nil;
  if ThePop = nil then
  BEGIN
    // pas de menu BTP trouve ..on le cree
    POPGS := TPopupMenu.Create(TT);
    POPGS.Name := 'POPBTP';
    fCreatedPop := true;
  END else
  BEGIN
    fCreatedPop := false;
    POPGS := TPopupMenu(thePop);
  END;
  DefiniMenuPop(TT);
{$ENDIF}
  if TT is TFFActure then
  begin
  	DEV := TFFacture(FF).DEV;
  	grid := TFFacture(TT).GS;
  end;
end;

procedure Tgestremplarticle.DefiniMenuPop(Parent: Tform);
var Indice : integer;
begin
  fmaxitems := 0;
  if not fcreatedPop then
  begin
    MesMenuItem[fmaxitems] := TmenuItem.Create (parent);
    with MesMenuItem[fmaxitems] do
      begin
      Caption := '-';
      end;
    inc (fmaxitems);
  end;
  MesMenuItem[fmaxitems] := TmenuItem.Create (parent);
  with MesMenuItem[fmaxitems] do
    begin
    Caption := TraduireMemoire ('Remplacement d''article');  // par défaut
    Name := 'BREMPLACEART';
    OnClick := GSremplArticle;
    enabled := true;
    end;
  MesMenuItem[fmaxitems].ShortCut := ShortCut( Word('R'), [ssCtrl,ssAlt]);
  inc (fmaxitems);

  for Indice := 0 to fmaxitems -1 do
    begin
      if MesMenuItem [Indice] <> nil then POPGS.Items.Add (MesMenuItem[Indice]);
    end;
end;

destructor Tgestremplarticle.destroy;
var indice : integer;
begin
  for Indice := 0 to fmaxitems -1 do
  begin
    MesMenuItem[Indice].Free;
  end;
  if fcreatedPop then POPGS.free;
  TOBMemorisation.free;
  inherited;
end;

procedure Tgestremplarticle.AddChampSupEchg (UneTOB : TOB);
begin
	UneTOB.AddChampSupValeur ('MODEREMPL',0,false);
  // article a remplacer
	UneTOB.AddChampSupValeur ('CODEARTAREMPL','',false);
	UneTOB.AddChampSupValeur ('ARTAREMPL','',false);
  // remplacement par
	UneTOB.AddChampSupValeur ('CODEARTICLE','',false);
	UneTOB.AddChampSupValeur ('ARTICLE','',false);
  //
	UneTOB.AddChampSupValeur ('MULTISEL','-',false);
	UneTOB.AddChampSupValeur ('GARDEVALO','-',false);
	UneTOB.AddChampSupValeur ('SUROUVRAGE','',false);
end;

procedure Tgestremplarticle.GSremplArticle(Sender: Tobject);
var UneTOBEchg : TOB;
		SavCol,SavRow : integer;
begin
  TOBMemorisation.clearDetail;
	SavCol := fgrid.col;
  Savrow := fGrid.row;
  iwarnings := 0; // initialisation
  fgrid.SynEnabled := False;
  UneTOBEchg := TOB.Create(' TOB ECHANGE', nil,-1);
  AddChampSupEchg (UneTOBechg);
  if fgrid.nbSelected > 0 then UneTOBEchg.putValue('MULTISEL','X');
  UneTOBEchg.data := TOBL;
  FFactGestRempl := TFFacture(FF);
  theTOB := UneTOBEchg;
  AGLLanceFiche ('BTP','BTREPLACEART','','','ACTION=MODIFICATION');
  FFactGestRempl := nil;
  if (theTOB <> NIL) and (TheTOB.GetValue('MODEREMPL') > 0) and (theTOB.GetValue('ARTICLE')<>'') then
  begin
  	RemplacementArticle (UneTOBechg);
    //
    if iWarnings > 0 then
    begin
    	AffichageWarnings;
    end;
    RaffraichitDetailOuvrage;
  end;
  UneTOBechg.free;
  TheTOB := nil;
  if (FF <> nil) and (TFFacture(FF).CopierColleObj <> nil) then
  begin
    TFFacture(FF).CopierColleObj.deselectionneRows;
  end else
  begin
  	TFFacture(FF).Gs.ClearSelected ;
  end;
  fgrid.SynEnabled := true;
  Tffacture(FF).GoToLigne (savrow,savcol);
end;

procedure Tgestremplarticle.RemplacementArticle (UneTOBechg : TOB);
begin
	if UneTOBEchg.GetValue ('MODEREMPL') = 1 then
  begin
  	// ligne courante
    if TOBL.GEtVAlue('GL_TYPELIGNE') <> 'ART' then exit;
		TraitementLigne (TOBL,UneTOBEchg);
    if TOBL.GetValue('GL_TYPEARTICLE')='OUV' then
    begin
      memoriseLigneOuv (TOBL);
//      GestionDetailOuvrage(FF,TobPiece,TOBL.GEtValue('GL_NUMLIGNE'));
    end;
  end else if UneTOBEchg.GetValue ('MODEREMPL') = 2 then
  begin
  	// selection de ligne courante
    EchangeSelection (UneTOBEchg);
  end else if UneTOBEchg.GetValue ('MODEREMPL') = 3 then
  begin
  	// sur le paragraphe courant
    EchangeParagraphe (UneTOBEchg);
  end else if UneTOBEchg.GetValue ('MODEREMPL') = 4 then
  begin
  	// sur la totalite du document
    EchangeDocument (UneTOBEchg);
  end;
end;

procedure Tgestremplarticle.SetAffaire(const Value: TOB);
begin
	TobAffaire := Value;
end;

procedure Tgestremplarticle.Setdocument(const Value: TOB);
begin
  TOBPiece := Value;
end;

(*
procedure Tgestremplarticle.SetEnabled(const value: boolean);
var Indice : integer;
begin
  for Indice := 0 to fMaxItems -1 do
  begin
    if (MesMenuItem[Indice].name = 'BREMPLACEART') then MesMenuItem[Indice].enabled := value;
  end;
end;
*)
procedure Tgestremplarticle.SetGrid(const Value: THGrid);
begin
	fgrid := value;
end;

procedure Tgestremplarticle.SetLigne(const Value: TOB);
begin
  TOBL := value;
end;

procedure Tgestremplarticle.SetTiers(const Value: TOB);
begin
  TOBTiers := value;
end;

procedure Tgestremplarticle.ReinitLigneDoc (TOBLigne : TOB); // Raz de la ligne sur les codes , valorisation, etc..
var Arow : integer;
begin
  Arow := TOBLIGne.getValue('GL_NUMLIGNE');
  TOBLigne.InitValeurs;
  InitLigneVide(TOBPiece, TOBLigne, TOBTiers, TOBAffaire, ARow, 1);
  TOBLIGne.PutValue('GL_ENCONTREMARQUE', '-');
//  if TOBPiece.FieldExists('_BLOQUETARIF') then TOBLIGne.PutValue('GL_BLOQUETARIF',TOBPiece.GetValue('_BLOQUETARIF'));

  // init domaine d'activité
  if ARow > 1 then TOBLIGne.PutValue('GL_DOMAINE', TOBPiece.Detail[ARow - 2].GetValue('GL_DOMAINE'));
end;

function Tgestremplarticle.PositionneArticle (TheTOBParam : TOB) : TOB;
var TOBA : TOB;
		refUnique : string;
    QQ : TQuery;
begin
	refUnique := TheTOBParam.getValue('ARTICLE');
  TOBA := FindTOBArtSais(TOBArticles, RefUnique);
  if TOBA = nil then
  begin
  	TOBA := CreerTOBArt(TOBArticles);
    QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                     'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                     'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+RefUnique+'"',true,-1, '', True);
//      TOBA.SelectDB('"' + RefUnique + '"', nil);
    TOBA.SelectDB('',QQ);
    InitChampsSupArticle (TOBA);
    ferme (QQ);
  end;
  result := TOBA;
end;

procedure Tgestremplarticle.EchangeLigne(TOBLS,TheTOBParam: TOB;WithCalcul : boolean=false);
var SavPVHT,SavPVTTC : double;
		SavQte : double;
		TOBArt : TOB;
    Arow : integer;
    RefUnique,SavRefArticle,SavLibelle,SavUnite : string;
begin

  if (TheTOBparam.GetValue('ARTAREMPL') <> '') then
  begin
  	RefUnique := TheTOBparam.GetValue('ARTAREMPL');
  end else
  begin
  	RefUnique := TOBLS.getValue('GL_ARTICLE');
  end;
  if TOBLS.getValue('GL_ARTICLE') <> RefUnique then exit;
  // FQ 11896 -- Les lignes remplacées ne sont pas déduites du total du document
  DeduitlaLigne (TOBLS);
  (*DeduitLigneModifie (TOBLS,TOBpiece,TOBBases,TOBBasesL,DEV.Decimale,tamodif);*)
  //

  if ((TOBPiece.getValue('GP_NATUREPIECEG')='ETU') and (TFFacture(FF).OrigineEXCEL)) or
  	 (TOBLS.GetValue('GLC_FROMBORDEREAU')='X') then
  begin
    SavRefArticle := TOBLS.GetValue('GL_REFARTSAISIE');
    SavLibelle :=TOBLS.GetValue('GL_LIBELLE');
    SavUnite :=TOBLS.GetValue('GL_QUALIFQTEVTE');
  end;

  SavPvHt := 0; SavPVTTC := 0;
  Arow := TOBLS.getValue('GL_NUMLIGNE');
  SavQte := TOBLS.getValue('GL_QTEFACT');
  if (TheTOBParam.GetValue('GARDEVALO')='X') or (TOBLS.GetValue('GLC_FROMBORDEREAU')='X') then
  begin
  	SavPvHt := TOBLS.GetValue('GL_PUHTDEV');
  	SavPvTTC := TOBLS.GetValue('GL_PUHTDEV');
  end;
  TOBART := PositionneArticle (TheTOBParam);
  if TOBART = nil then exit;
  // reinit ligne a vide
  ReinitLigneDoc (TOBLS); // Raz de la ligne sur les codes , valorisation, etc..
  if (TOBArt <> nil) then
  begin
  	TOBLS.PutValue('GL_CODEARTICLE',TOBART.GEtValue('GA_CODEARTICLE'));
  	TOBLS.PutValue('GL_ARTICLE',TOBART.GEtValue('GA_ARTICLE'));
  	TOBLS.PutValue('GL_REFARTSAISIE',TOBART.GEtValue('GA_CODEARTICLE'));
  	TOBLS.PutValue('GL_REFARTTIERS',TOBART.GEtValue('GA_CODEARTICLE'));
    fgrid.Cells[SG_RefArt,Arow] := TOBART.GEtValue('GA_CODEARTICLE');
    TFfacture(FF).StCellCur := fgrid.Cells[SG_RefArt,Arow];
    //  BBI, fiche correction 10410
    CalculePrixArticle(TOBArt, TobPiece.GetValue('GP_DEPOT'));
    //  BBI, fin fiche correction 10410
    TFFacture(FF).CodesArtToCodesLigne(TOBArt, ARow);
  end;
  TFFacture(FF).UpdateArtLigne (Arow,false,true,true,SavQte);
  if (SavPVHt <> 0) or (TOBLS.GetValue('GLC_FROMBORDEREAU')='X') then
  begin
  	TOBLS.putValue('GL_PUHTDEV',SavPvHt);
  	TOBLS.putValue('GL_PUTTCDEV',SavPvTTC);
  end;
  //
  TOBLS.PutValue('GL_QTEFACT',SavQte );
  TOBLS.PutValue('GL_QTERESTE',SavQte );
  TOBLS.PutValue('GL_QTESTOCK',SavQte );
  //
  TOBLS.PutValue('GL_MTRESTE',TOBLS.getValue('GL_MONTANTHTDEV'));
  //
  if ((TOBPiece.getValue('GP_NATUREPIECEG')='ETU') and (TFFacture(FF).OrigineEXCEL)) or
  	 (TOBLS.GetValue('GLC_FROMBORDEREAU')='X') then
  begin
    TOBLS.PutValue('GL_REFARTSAISIE',SavRefArticle);
    TOBLS.PutValue('GL_LIBELLE',SavLibelle);
    TOBLS.PutValue('GL_QUALIFQTEVTE',SavUnite);
  end;
  TOBLS.PutValue('GL_RECALCULER','X');
  TOBPiece.putvalue('GP_RECALCULER','X');
  TFFacture(FF).CalculeLaSaisie(-1, TOBLS.getValue('GL_NUMLIGNE'), False);
end;

procedure Tgestremplarticle.EchangeSelection(TheTOBParam: TOB);
var Indice : integer;
		TOBLS : TOB;
begin
  for indice := 1 to fgrid.RowCount do
  begin
    if (fgrid.IsSelected (Indice) ) then
    begin
    	TOBLS := GetTOBLigne (TOBPiece,Indice);
    	if TOBLS.GEtVAlue('GL_TYPELIGNE') <> 'ART' then continue;
			TraitementLigne (TOBLS,TheTOBParam);
      if TOBLS.GetValue('GL_TYPEARTICLE')='OUV' then
      begin
        memoriseLigneOuv (TOBLS);

        //GestionDetailOuvrage(FF,TobPiece,TOBLS.GEtValue('GL_NUMLIGNE'));
      end;
    end;
  end;
//  TOBPiece.PutValue('GP_RECALCULER', 'X');
//  TFFacture(FF).CalculeLaSaisie(-1, -1, False);
end;

procedure Tgestremplarticle.EchangeParagraphe(TheTOBParam: TOB);
var Ligdepart,LigFin : integer;
		TOBLS : TOB;
    Indice : integer;
    Article : string;
begin
  if (TheTOBparam.GetValue('ARTAREMPL') <> '') then
  begin
  	Article := TheTOBparam.GetValue('ARTAREMPL');
  end else
  begin
  	Article := TOBL.GetValue('GL_ARTICLE');
  end;
  RetrouveParagraphe (Ligdepart,LigFin,TOBL);
  if (LIgdepart = 0) and (ligfin = 0) then exit;
  for Indice := Ligdepart-1 to Ligfin-1 do
  begin
    TOBLS := TOBPiece.detail[Indice];
    if TOBLS.GEtVAlue('GL_TYPELIGNE') <> 'ART' then continue;
    TraitementLigne (TOBLS,TheTOBParam);
    if TOBLS.GetValue('GL_TYPEARTICLE')='OUV' then
    begin
      memoriseLigneOuv (TOBLS);
    	//GestionDetailOuvrage(FF,TobPiece,TOBLS.GEtValue('GL_NUMLIGNE'));
    end;
  end;
  TOBPiece.PutValue('GP_RECALCULER', 'X');
//  TFFacture(FF).CalculeLaSaisie(-1, -1, False);
end;

procedure Tgestremplarticle.Reinitialise (TOBO : TOB);
var LastIndice : integer;
begin
  LastIndice := TOBO.GetValue('BLO_NUMORDRE');
  TOBO.InitValeurs;
  TOBO.PutValue('BLO_NUMORDRE',LastIndice);
end;

procedure Tgestremplarticle.AffecteArticleSurOUv (TOBA,TOBO: TOB);
begin
  TOBO.putValue('BLO_ARTICLE',TOBA.GetValue('GA_ARTICLE'));
  TOBO.putValue('BLO_CODEARTICLE',TOBA.GetValue('GA_CODEARTICLE'));
  TOBO.putValue('BLO_REFARTSAISIE',TOBA.GetValue('GA_CODEARTICLE'));
  TOBO.putValue('BLO_LIBELLE',TOBA.GetValue('GA_LIBELLE'));
  TOBO.PutValue('BLO_TENUESTOCK',TOBA.GetValue('GA_TENUESTOCK'));
  TOBO.PutValue('BLO_QUALIFQTEVTE',TobA.GetValue('GA_QUALIFUNITEVTE'));
  TOBO.PutValue('BLO_QUALIFQTESTO',TobA.GetValue('GA_QUALIFUNITESTO'));
  TOBO.PutValue('BLO_PRIXPOURQTE',TobA.GetValue('GA_PRIXPOURQTE'));
  TOBO.PutValue('BLO_DPA',TobA.GetValue('GA_PAHT'));
  TOBO.PutValue('BLO_PMAP',TobA.GetValue('GA_PMAP'));
  if (TOBPiece.fieldExists ('COEFFGFORCE')) and (TOBPiece.GetValue ('COEFFGFORCE') <> 0) then
  begin
    TOBO.PutValue('BLO_DPR',TobA.GetValue('GA_DPA') * TOBPiece.GetValue ('COEFFGFORCE') );
    TOBO.PutValue('BLO_PMRP',TobA.GetValue('GA_DPR'));
  end else
  begin
    TOBO.PutValue('BLO_DPR',TobA.GetValue('GA_DPR'));
    TOBO.PutValue('BLO_PMRP',TobA.GetValue('GA_PMRP'));
  end;
  CopieOuvFromArt (TOBPiece,TOBO,TobA,DEV);

  TOBO.PutValue('BLO_PUHT',TOBA.GetValue('GA_PVHT'));
  TOBO.PutValue('BLO_PUHTBASE',TOBA.GetValue('GA_PVHT'));
  TOBO.PutValue('BLO_PUHTDEV',pivottodevise(TOBO.GetValue('BLO_PUHTBASE'),DEV.Taux,DEV.quotite,V_PGI.OkdecP ));
  TOBO.PutValue('ANCPV',TOBO.GetValue('BLO_PUHTDEV'));
  TOBO.PutValue('BLO_PUTTC',TOBA.GetValue('GA_PVTTC'));
  TOBO.PutValue('BLO_PUTTCBASE',TOBA.GetValue('GA_PVTTC'));
  TOBO.PutValue('BLO_PUTTCDEV',pivottodevise(TOBO.GetValue('BLO_PUTTCBASE'),DEV.Taux,DEV.quotite,V_PGI.OkdecP ));

  TOBO.PutValue('BLO_LIBELLE',TOBA.GetValue('GA_LIBELLE'));
  TOBO.putValue('BNP_TYPERESSOURCE',TOBA.getValue('BNP_TYPERESSOURCE'));
end;

procedure Tgestremplarticle.ReplaceArticleDetailOuvrage (TOBLS,TOBO,TheTOBParam : TOB);

  procedure SetInfoLigne (TOBDest,TOBsrc : TOB);
  begin
    TOBDest.putValue('BLO_NIVEAU',TOBSrc.GetValue('BLO_NIVEAU'));
    TOBDest.putValue('BLO_ORDRECOMPO',TOBSrc.GetValue('BLO_ORDRECOMPO'));
    //
    TOBDest.PutValue('BLO_NATUREPIECEG',TOBSrc.GetValue('BLO_NATUREPIECEG')) ;
    TOBDest.PutValue('BLO_REGIMETAXE',TOBSrc.GetValue('BLO_REGIMETAXE')) ;
    TOBDest.PutValue('BLO_SOUCHE',TOBSrc.GetValue('BLO_SOUCHE')) ;
    TOBDest.PutValue('BLO_NUMERO',TOBSrc.GetValue('BLO_NUMERO')) ;
    TOBDest.PutValue('BLO_INDICEG',TOBSrc.GetValue('BLO_INDICEG')) ;
    TOBDest.PutValue('BLO_NUMLIGNE',TOBSrc.GetValue('BLO_NUMLIGNE')) ;
  end;

var TOBART : TOB;
		LastQte,LastPuhtDev,LastPuttcdev : double;
    TOBOL : TOB;
    StInfg,StInFc : boolean;
begin
  StInFg := (TOBPiece.getValue('GP_APPLICFGST')='X');
  StInFC := (TOBPiece.getValue('GP_APPLICFCST')='X');
  //
  TOBOL := TOB.create ('LIGNEOUV',nil,-1);
  SetInfoLigne (TOBOL,TOBO); // sauvegarde des données
  TOBART := PositionneArticle (TheTOBParam);
  if TOBART = nil then exit;
  lastQTe := TOBO.GetValue('BLO_QTEFACT');
  LastPuHtdev := 0;
  LastPuTTCdev := 0;
  if TheTOBParam.GetValue('GARDEVALO')='X' then
  begin
  	LastPuhtDev := TOBO.GetValue('BLO_PUHTDEV');
  	LastPuttcdev := TOBO.GetValue('BLO_PUTTCDEV');
  end;

  Reinitialise (TOBO);
  CopieOuvFromLigne (TOBO,TOBLS);
  AffecteArticleSurOUv (TOBART,TOBO);
  //
  SetInfoLigne (TOBO,TOBOL);   // restauration des données
  //
  if (IsPrestationSt(TOBO)) and (not StInFg) then TOBO.PutValue('BLO_NONAPPLICFRAIS','X') ;
  if (IsPrestationSt(TOBO)) and (not StInFC) then TOBO.PutValue('BLO_NONAPPLICFC','X') ;
  if (Pos (TOBO.getValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne) > 0) and (not IsLigneExternalise(TOBO)) then
  begin
    TOBO.PutValue('BLO_TPSUNITAIRE',1);
  end else
  begin
    TOBO.PutValue('BLO_TPSUNITAIRE',0);
  end;
  //
  if TheTOBParam.GetValue('GARDEVALO')='X' then
  begin
  	TOBO.PutValue('BLO_PUHTDEV',LastPuhtDev);
  	TOBO.PutValue('BLO_PUTTCDEV',LastPuTTCDev);
  end;
  TOBO.PutValue('BLO_QTEFACT',lastQTe);
  if TOBPIece.getValue('GP_FACTUREHT')='X' then
  begin
  	CalculeLigneHTOuv  (TOBO,TOBPiece,DEV);
  end else
  begin
  	CalculeLigneTTCOuv  (TOBO,TOBPiece,DEV);
  end;
  TOBOL.free;
end;

procedure Tgestremplarticle.EchangeSurOuvrage (TheARticle : string; TOBLS,TOBOUV,TheTOBParam : TOB);
var Indice : integer;
		TOBO : TOB;
    Article : string;
begin
  if (TheTOBparam.GetValue('ARTAREMPL') <> '') then
  begin
  	Article := TheTOBparam.GetValue('ARTAREMPL');
  end else
  begin
		Article := TheARticle;
  end;

  for Indice := 0 to TOBOUV.detail.count -1 do
  begin
  	TOBO := TOBOUV.detail[Indice];
    if TOBO.GetValue('BLO_ARTICLE') = Article then
    begin
    	ReplaceArticleDetailOuvrage (TOBLS,TOBO,TheTOBParam);
    end;
    if Pos(TOBO.getValue ('BLO_TYPEARTICLE'),'ARP;OUV;') > 0 then
    begin
    	EchangeSurOuvrage (TheArticle,TOBLS,TOBO,TheTOBParam);
    end;
  end;
end;

procedure Tgestremplarticle.TraitementLigne (TOBLS,TheTOBparam : TOB);
var TOBOUV : TOB;
    TheValeurs : T_Valeurs;
    TheARticle : string;
    Qte : double;
begin
	// controle de base
  if CtrlOkReliquat(TOBLS, 'GL') then
  begin
    if (TOBLS.GetValue('GL_MONTANTHTDEV') <> TOBLS.GetValue('GL_MTRESTE')) or
       (TOBLS.GetValue('GL_SOLDERELIQUAT')='X') then
    begin
      iWarnings := 1;
      exit;
    end;

  end
  else
  begin
    if (TOBLS.GetValue('GL_QTEFACT') <> TOBLS.GetValue('GL_QTERESTE')) or
        (TOBLS.GetValue('GL_SOLDERELIQUAT')='X') then
    begin
      iWarnings := 1;
      exit;
    end;
  end;

  // --
  Qte := TOBLS.GetValue('GL_QTEFACT');
	TheArticle := TOBLS.GetValue('GL_ARTICLE');
  if (TheTOBParam.getValue('SUROUVRAGE') = '' ) then EchangeLigne (TOBLS,TheTOBparam);
  if (TOBLS.GetValue('GL_INDICENOMEN') > 0) and ( Pos(TOBLS.getValue('GL_TYPEARTICLE'),'OUV;ARP')>0) then
  begin
  	if (TheTOBParam.getValue('SUROUVRAGE') <> '' ) and (TheTOBParam.GetValue('SUROUVRAGE')<> TheArticle) then exit;
    if TOBOUvrage <> nil then
    begin
      //
      DeduitlaLigne (TOBLS);
      //
      TOBOUV := TOBOuvrage.detail[TOBLS.GetValue('GL_INDICENOMEN')-1];
      //
      EchangeSurOuvrage (TheArticle,TOBLS,TOBOUV,TheTOBParam);
      NumeroteLigneOuv (TOBOUV,TOBLS,1,1,0,0,0);
      //
      CalculeOuvrageDoc (TOBOUV,1,1,true,DEV,TheValeurs,(TOBPiece.getValue('GP_FACTUREHT')='X'));
      //
      if (Pos(TOBL.GetValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne)>0) and (not IsLigneExternalise(TOBL)) then
      // if (Pos(TOBL.GetValue('BNP_TYPERESSOURCE'),'SAL;INT;')>0) then
      begin
        TOBLS.Putvalue('BLO_TPSUNITAIRE',TheValeurs[9]);
      end;
      TOBLS.Putvalue('GL_MONTANTPAFG',TheValeurs[10]*Qte);
      TOBLS.Putvalue('GL_MONTANTPAFR',TheValeurs[11]*Qte);
      TOBLS.Putvalue('GL_MONTANTPAFC',TheValeurs[12]*Qte);
      TOBLS.Putvalue('GL_MONTANTFG',  TheValeurs[13]*Qte);
      TOBLS.Putvalue('GL_MONTANTFR',  TheValeurs[14]*Qte);
      TOBLS.Putvalue('GL_MONTANTFC',  TheValeurs[15]*Qte);
//      TOBLS.Putvalue('GL_MONTANTPA',TheValeurs[16]*Qte);
//      TOBLS.Putvalue('GL_MONTANTPR',TheValeurs[17]*Qte);
      TOBLS.Putvalue('GL_MONTANTPA',Arrondi((TOBLS.Getvalue('GL_QTEFACT') * TOBLS.GetValue('GL_DPA')),V_PGI.okdecV));
      TOBLS.Putvalue('GL_MONTANTPR',Arrondi((TOBLS.Getvalue('GL_QTEFACT') * TOBLS.GetValue('GL_DPR')),V_PGI.okdecV));

      TOBLS.Putvalue('GL_DPA',TheValeurs[16]);
      TOBLS.Putvalue('GL_DPR',TheValeurs[17]);
      //
      TOBLS.Putvalue('GL_PUHTDEV',TheValeurs[2]);
      TOBLS.Putvalue('GL_PUTTCDEV',TheValeurs[3]);
      TOBLS.Putvalue('GL_PUHT',DeviseToPivotEx(TOBLS.GetValue('GL_PUHTDEV'),DEV.Taux,DEV.quotite,V_PGI.OkdecP));
      TOBLS.Putvalue('GL_PUTTC',DevisetoPivotEx(TOBLS.GetValue('GL_PUTTCDEV'),DEV.taux,DEV.quotite,V_PGI.okdecP));
      TOBLS.Putvalue('GL_PUHTBASE',TOBLS.GetValue('GL_PUHT'));
      TOBLS.Putvalue('GL_PUTTCBASE',TOBLS.GetValue('GL_PUTTC'));
      TOBLS.Putvalue('GL_DPA',TheValeurs[0]);
      TOBLS.Putvalue('GL_DPR',TheValeurs[1]);
      TOBLS.Putvalue('GL_PMAP',TheValeurs[6]);
      TOBLS.Putvalue('GL_PMRP',TheValeurs[7]);
      TOBLS.putvalue('GL_TPSUNITAIRE',TheValeurs[9]);
      TOBLS.putvalue('GL_RECALCULER','X');
      StockeInfoTypeLigne (TOBL,TheValeurs);      
      TOBPiece.putvalue('GP_RECALCULER','X');
      TFFacture(FF).CalculeLaSaisie(-1, TOBLS.getValue('GL_NUMLIGNE'), False);
    end;
  end;
end;


procedure Tgestremplarticle.EchangeDocument(TheTOBParam: TOB);
var TOBLS : TOB;
    Indice : integer;
begin

  for Indice := 0 to TobPiece.detail.count -1 do
  begin
    TOBLS := TOBPiece.detail[Indice];
//    if TOBLS.GEtVAlue('GL_TYPELIGNE') <> 'ART' then continue;
    if not IsArticle(TOBLS) then continue;
    TraitementLigne (TOBLS,TheTOBparam);
    if TOBLS.GetValue('GL_TYPEARTICLE')='OUV' then
    begin
      memoriseLigneOuv (TOBLS);
      //GestionDetailOuvrage(FF,TobPiece,TOBLS.GEtValue('GL_NUMLIGNE'));
    end;
  end;
  TOBPiece.PutValue('GP_RECALCULER', 'X');
//  TFFacture(FF).CalculeLaSaisie(-1, -1, False);
end;

procedure Tgestremplarticle.SetArticles(const Value: TOB);
begin
	TOBArticles  := Value;
end;

procedure Tgestremplarticle.RetrouveParagraphe(var Ligdepart, LigFin: integer;TOBLigne: TOB);
var Ligne : integer;
    TOBLS : TOB;
    indice : integer;
begin
	Ligdepart := 0;
  LigFin := 0;
  if TOBLIgne.getValue('GL_NIVEAUIMBRIC') = 0 then exit; // on est pas dans un paragraphe
  Ligne := TOBLigne.getValue('GL_NUMLIGNE');
	for Indice := Ligne -1 downto 0 do
  begin
  	TOBLS := TOBPiece.detail[Indice];
    if IsDebutParagraphe (TOBLS) then
    begin
    	LigDepart := TOBLS.GetValue('GL_NUMLIGNE');
      break;
    end;
  end;
	for Indice := Ligne -1 to TOBPiece.detail.count -1 do
  begin
  	TOBLS := TOBPiece.detail[Indice];
    if IsFinParagraphe (TOBLS) then
    begin
    	Ligfin := TOBLS.GetValue('GL_NUMLIGNE');
      break;
    end;
  end;
end;

procedure Tgestremplarticle.SetOuvrage(const Value: TOB);
begin
  TOBOuvrage := Value;
end;

procedure Tgestremplarticle.AffichageWarnings;
begin
	PGIInfo ('Certains articles n''ont pus êtres traités',FF.Caption);
end;

procedure Tgestremplarticle.MemoriseLigneOuv(TOBL : TOB);
var TOBM : TOB;
begin
  //
  TOBM := TOB.Create ('UNE LIGNE',TOBMemorisation,-1);
  TOBM.AddChampSupValeur ('LIGNE',TOBL.GetValue('GL_NUMORDRE'));
end;

procedure Tgestremplarticle.RaffraichitDetailOuvrage;
var Indice : integer;
    TOBM : TOB;
begin
  for Indice := 0 to TOBMemorisation.detail.count -1 do
  begin
    TOBM := TOBMemorisation.detail[Indice];
    TOBL := GetTOBLigneViaOrdre (TOBM.GetValue('LIGNE'));
    if TOBL <> nil then GestionDetailOuvrage(FF,TobPiece,TOBL.GEtValue('GL_NUMLIGNE'));
  end;
end;

function Tgestremplarticle.GetTOBLigneViaOrdre(NumOrdre: integer): TOB;
var Indice : integer;
begin
  result := nil;
  for Indice := 0 to TOBPiece.Detail.count -1 do
  begin
    if TOBPiece.Detail[Indice].getValue('GL_NUMORDRE')=NumOrdre then
    begin
      result := TOBPiece.Detail[Indice];
    end;
  end;
end;

procedure Tgestremplarticle.SetBases(const Value: TOB);
begin
  TOBBases := value;
end;

procedure Tgestremplarticle.SetBasesL(const Value: TOB);
begin
  TOBBasesL := value;
end;

procedure Tgestremplarticle.SSEnabled (etat : boolean);
var indice : integer;
begin
  for Indice := 0 to fmaxitems -1 do
  begin
    if MesMenuItem [Indice].Name = 'BREMPLACEART' then MesMenuItem [Indice].Enabled := Etat;
  end;
end;

procedure Tgestremplarticle.SetEnabled;
begin
  if TFFacture(FF).IsDejaFacture then
  begin
    SSEnabled (false);
  end else
  begin
    SSEnabled (True);
  end;
end;

procedure Tgestremplarticle.DeduitlaLigne(TOBLS: TOB);
var TOBB : TOB;
    NumOrdre : integer;
begin
  AssigneDocumentTva2014(TOBPiece,TOBSSTRAIT);
  DeduitLigneModifie (TOBLS,TOBpiece,TOBpieceTrait,TOBOUvrage,TOBOUvragesP,TOBBases,TOBBasesL,TOBTiers,TFFacture(FF).DEV, taModif);
  ZeroLigneMontant (TOBLS);
  InitDocumentTva2014;
end;

{ TOF_BTREPLACEART }

procedure TOF_BTREPLACEART.DetermineEnabledFac;

begin
	// si dans un paragraphe
  // Selection de ligne courante
	if laTOB.GetValue('MULTISEL') = 'X'     then
  begin
  	TradioButton(GetControl('CURLIG')).enabled := false;
  	TradioButton(GetControl('CHLIGNE')).enabled := false;
    TradioButton(GetControl('CHARTICLE')).checked := true;
    TradioButton(GetControl('ELTSEL')).Enabled := true;
    TradioButton(GetControl('ELTSEL')).Checked := true;
    TradioButton(GetControl('REPLSURPAR')).Enabled := false;
    TradioButton(GetControl('REPLTOUTDOC')).Enabled := false;
    TradioButton(GetControl('MMOUVRAGE')).Enabled := false;
    ecran.caption := 'Remplacement de l''article de la sélection courante';
  end else
  begin
    if OkTYpeArticle (TOBL.getValue('GL_TYPEARTICLE')) then
    begin
    	ecran.caption := 'Remplacement de l''article '+TOBL.GetValue('GL_REFARTSAISIE');
    	TradioButton(GetControl('CHLIGNE')).checked := true;
			ChoixArtligne (self);
    	TradioButton(GetControl('CURLIG')).Checked := true;
      RefUniqueAchg := TOBL.GetValue('GL_ARTICLE');
    end else
    begin
    	ecran.caption := 'Remplacement';
      TradioButton(GetControl('CURLIG')).Enabled := false;
    	TradioButton(GetControl('CHLIGNE')).Enabled := false;
    	TradioButton(GetControl('CHARTICLE')).checked := true;
      ChoixArticlePart (self);
      if pos(TOBL.GetValue('GL_TYPEARTICLE'),'OUV;ARP;') > 0  then
      begin
    		ecran.caption := 'Remplacement dans ouvrage '+TOBL.GetValue('GL_REFARTSAISIE');
      	TradioButton(GetControl('CURLIG')).Enabled := true;
    		TradioButton(GetControl('CURLIG')).Checked := true;
        if TOBL.GetValue('GL_TYPEARTICLE')='OUV' then
        begin
      		TradioButton(GetControl('CURLIG')).Caption := 'Sur l''ouvrage';
        end;
      end else
      begin
    		TradioButton(GetControl('REPLTOUTDOC')).Checked := true;
      end;
    end;
  	TradioButton(GetControl('ELTSEL')).Enabled := false;
		if TOBL.GetValue('GL_NIVEAUIMBRIC') > 0 then TradioButton(GetControl('REPLSURPAR')).Enabled := true
  																				  else TradioButton(GetControl('REPLSURPAR')).Enabled := false;
  end;
end;

procedure TOF_BTREPLACEART.DetermineEnabledOuv;
begin

end;

procedure TOF_BTREPLACEART.OnArgument(S: String);
begin
  inherited;
  TOBL := TOB(laTOB.data);
  if TOBL.NomTable = 'LIGNE' then DetermineEnabledFac
  else if TOBL.NomTable = 'LIGNEOUV' then DetermineEnabledOuv
  else TFVierge(ecran).TypeAction := taConsult;
  THLabel(GetCOntrol('LREPLACEART')).caption := '';
  THLabel(GetCOntrol('LAREPLACEART')).caption := '';
  ThEdit(GetControl('AREPLACEART')).OnElipsisClick := ARecharticle;
  ThEdit(GetControl('AREPLACEART')).OnExit := VerificationArtAreplace;
  ThEdit(GetControl('REPLACEART')).OnElipsisClick := Recharticle;
  ThEdit(GetControl('REPLACEART')).OnExit := VerificationArtreplacement;
  TradioButton(GetControl('CHLIGNE')).OnClick := ChoixArtligne;
  TradioButton(GetControl('CHARTICLE')).OnClick := ChoixArticlePart;
  TradioButton(GetControl('CURLIG')).OnClick := QueSurLigne;
  TradioButton(GetControl('REPLSURPAR')).OnClick := PasQueSurLigne;
  TradioButton(GetControl('REPLTOUTDOC')).OnClick := PasQueSurLigne;

end;

procedure TOF_BTREPLACEART.OnCancel;
begin
  inherited;
end;

procedure TOF_BTREPLACEART.OnClose;
begin
  inherited;
end;

procedure TOF_BTREPLACEART.OnDelete;
begin
  inherited;

end;

procedure TOF_BTREPLACEART.OnDisplay;
begin
  inherited;

end;

procedure TOF_BTREPLACEART.OnLoad;
begin
  inherited;
  ecran.Update;
end;

procedure TOF_BTREPLACEART.OnNew;
begin
  inherited;

end;


function TOF_BTREPLACEART.VerifArticle (NomChamp,NOMChampLib : string; var LaRefUnique : string) : boolean;
var Article : string;
		QQ : Tquery;
begin
	result := true;
  Article := ThEdit(GetControl(NomCHamp)).Text;
  if (LaRefUnique <> '') or ( Article <> '') then
  begin
  	if LaRefUnique <> '' then
    begin
  		QQ := OpenSql('SELECT GA_ARTICLE,GA_LIBELLE,GA_STATUTART FROM ARTICLE WHERE GA_ARTICLE="'+LaRefUnique+'" AND GA_STATUTART="UNI"',true,-1, '', True);
    end else
    begin
  		QQ := OpenSql('SELECT GA_ARTICLE,GA_LIBELLE,GA_STATUTART FROM ARTICLE WHERE GA_CODEARTICLE="'+Article+'" AND GA_STATUTART IN ("UNI","GEN")',true,-1, '', True);
    end;
    if QQ.eof then
    begin
    	result := false;
    end else
    begin
     	if QQ.findField('GA_STATUTART').asString = 'GEN' then
      begin
      	LaRefUnique := SelectUneDimension (QQ.findField('GA_ARTICLE').asString);
        if LaRefUnique <> '' then
        begin
        	ferme (QQ);
          QQ := OpenSql('SELECT GA_LIBELLE,GA_STATUTART FROM ARTICLE WHERE GA_ARTICLE="'+LaRefUnique+'"',true,-1, '', True);
          ThLabel(GetControl(NOMChampLib)).Caption   := QQ.findField('GA_LIBELLE').asString;
        end else
        begin
          result := false;
        end;
      end else
      begin
      	LaRefUnique := QQ.findField('GA_ARTICLE').asString;
    		ThLabel(GetControl(NOMChampLib)).caption  := QQ.findField('GA_LIBELLE').asString;
      end;
    end;
  	ferme (QQ);
	end;
end;

procedure TOF_BTREPLACEART.OnUpdate;
begin
  inherited;
  if TradioButton(GetControl('CURLIG')).Checked  then laTOB.putValue('MODEREMPL',1)
  else if TradioButton(GetControl('ELTSEL')).checked then laTOB.putValue('MODEREMPL',2)
  else if TradioButton(GetControl('REPLSURPAR')).checked then laTOB.putValue('MODEREMPL',3)
  else if TradioButton(GetControl('REPLTOUTDOC')).checked then laTOB.putValue('MODEREMPL',4);

  if not VerifArticle ('AREPLACEART', 'LAREPLACEART', RefUniqueAchg ) then
  begin
    PgiBox ('L''article n''existe pas',ecran.caption);
    TFVierge(ecran).ModalResult := 0;
    exit;
  end;

  if not VerifArticle ('REPLACEART', 'LREPLACEART', RefUniqueChg ) then
  begin
    PgiBox ('L''article n''existe pas',ecran.caption);
    TFVierge(ecran).ModalResult := 0;
    exit;
  end;

  if RefUniqueChg <> '' then
  begin
    laTOB.putValue('CODEARTICLE',THEdit(GetCOntrol('REPLACEART')).Text);
    laTOB.putValue('ARTICLE',RefUniqueChg);
  end;

  if RefUniqueAChg <> '' then
  begin
    laTOB.putValue('CODEARTAREMPL',THEdit(GetCOntrol('AREPLACEART')).Text);
    laTOB.putValue('ARTAREMPL',RefUniqueAchg);
  end;

  if TCheckBox(GetControl('CGARDEVALO')).checked then LaTOB.PutValue ('GARDEVALO','X');
  if TCheckBox(GetControl('MMOUVRAGE')).checked then LaTOB.PutValue ('SUROUVRAGE',TOBL.GEtValue('GL_ARTICLE'));
  TheTOB := LaTOB;

end;

procedure TOF_BTREPLACEART.Recharticle(sender: Tobject);
var TheArt : string;
		sWhere : string;
    TheRange : string;
    TheResult : string;
begin
	TheArt := Trim(ThEdit(GetControl('REPLACEART')).Text);
  sWhere := ' GA_TYPEARTICLE="MAR" OR GA_TYPEARTICLE="PRE" OR GA_TYPEARTICLE="FRA" OR GA_TYPEARTICLE="POU"';
  TheRange := 'XX_WHERE= AND (GA_TYPEARTICLE="MAR" OR GA_TYPEARTICLE="PRE" OR GA_TYPEARTICLE="FRA" OR GA_TYPEARTICLE="POU")';
  if TheARt <> '' then TheRange := TheRange+';GA_CODEARTICLE=' +TheArt;
	Theresult := DispatchRechArt(ThEdit(GetControl('REPLACEART')), 1, sWhere,TheRange,'');
  if Theresult <> '' then
  begin
  	if VerifArticle ('REPLACEART', 'LREPLACEART', theresult ) then
    begin
    	RefUniqueChg  := TheREsult;
    	ThEdit(GetControl('REPLACEART')).Text  := Trim(Copy(RefUniqueChg,1,18));
    end else
    begin
    	ThEdit(GetControl('REPLACEART')).Text  := Trim(Copy(RefUniqueChg,1,18));
    end;
  end;
end;


procedure TOF_BTREPLACEART.ARecharticle(sender: Tobject);
var TheArt : string;
		sWhere : string;
    TheRange : string;
    TheREsult : string;
begin
	TheArt := Trim(ThEdit(GetControl('AREPLACEART')).Text);
  sWhere := ' GA_TYPEARTICLE="MAR" OR GA_TYPEARTICLE="PRE" OR GA_TYPEARTICLE="FRA" OR GA_TYPEARTICLE="POU"';
  TheRange := 'XX_WHERE= AND (GA_TYPEARTICLE="MAR" OR GA_TYPEARTICLE="PRE" OR GA_TYPEARTICLE="FRA" OR GA_TYPEARTICLE="POU")';
  if TheARt <> '' then TheRange := TheRange+';GA_CODEARTICLE=' +TheArt;
	TheREsult := DispatchRechArt(ThEdit(GetControl('AREPLACEART')), 1, sWhere,TheRange,'');
  if TheREsult <> '' then
  begin
  	if VerifArticle ('AREPLACEART', 'LAREPLACEART', Theresult ) then
    begin
    	RefUniqueAchg  := TheREsult;
    	ThEdit(GetControl('AREPLACEART')).Text  := Trim(Copy(RefUniqueAChg,1,18));
    end else
    begin
    	ThEdit(GetControl('AREPLACEART')).Text  := Trim(Copy(RefUniqueAChg,1,18));
    end;
  end;
end;

procedure TOF_BTREPLACEART.ChoixArtligne(Sender: TObject);
begin
  Thedit(GetCOntrol('AREPLACEART')).Enabled := false;
  Thedit(GetCOntrol('AREPLACEART')).Color := clInactiveCaptiontext;
  TradioButton(GetControl('CURLIG')).enabled := true;
  TradioButton(GetControl('CURLIG')).checked := true;
  TcheckBox(GetControl('MMOUVRAGE')).Enabled := false;
end;

procedure TOF_BTREPLACEART.ChoixArticlePart(sender: Tobject);
begin
  Thedit(GetCOntrol('AREPLACEART')).Enabled := True;
  Thedit(GetCOntrol('AREPLACEART')).Color := clWindow;
  Thedit(GetCOntrol('AREPLACEART')).Text := '';
  TradioButton(GetControl('CURLIG')).Checked := true;
end;

procedure TOF_BTREPLACEART.VerificationArtAreplace(Sender: TObject);
begin
  if not VerifArticle ('AREPLACEART', 'LAREPLACEART', RefUniqueAchg ) then
  begin
    PgiBox ('L''article n''existe pas',ecran.caption);
  	ThEdit(GetControl('AREPLACEART')).SetFocus;
  end;
end;

procedure TOF_BTREPLACEART.VerificationArtreplacement(Sender: Tobject);
begin
  if not VerifArticle ('REPLACEART', 'LREPLACEART', RefUniqueChg ) then
  begin
    PgiBox ('L''article n''existe pas',ecran.caption);
  	ThEdit(GetControl('REPLACEART')).SetFocus;
  end;
end;

procedure TOF_BTREPLACEART.PasQueSurLigne(Sender: TObject);
begin
	if TOBL.GetValue('GL_TYPEARTICLE')='OUV' then TcheckBox(GetControl('MMOUVRAGE')).Enabled := True;
end;

procedure TOF_BTREPLACEART.QueSurLigne(Sender: TObject);
begin
  TcheckBox(GetControl('MMOUVRAGE')).Enabled := false;
  TcheckBox(GetControl('MMOUVRAGE')).Checked := false;
end;

Initialization
  registerclasses ( [ TOF_BTREPLACEART ] ) ;
end.

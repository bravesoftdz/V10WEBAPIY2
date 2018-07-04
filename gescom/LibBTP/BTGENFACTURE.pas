unit BTGENFACTURE;

interface

Uses HEnt1, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,HCtrls,forms,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DB, Fe_Main,
      {$IFDEF V530} EdtDOC,{$ELSE} EdtRdoc, {$ENDIF}
{$ENDIF}
{$IFDEF BTP}
     // Modif BTP
     BTPUtil,FactAcompte, FactGrpBtp,
     (* CONSO *)
     UtilPhases,FactTobBtp,
     (* ------ *)
     FactAdresseBTP,Factrgpbesoin,FactTvaMilliem,
     // ----
{$ENDIF}
     SysUtils, Dialogs, SaisUtil, UtilPGI, AGLInit, FactUtil, UtilSais,chancel,
     EntGC, Classes, HMsgBox, SaisComm, FactComm, ed_tools, FactNomen, UtilGrp,
     FactCpta, FactCalc, ParamSoc, UtilArticle,  Stockutil,
{$IFDEF AFFAIRE}
      FactActivite, AffaireUtil, FactAffaire,
{$ENDIF}
     Factspec,FactRg, FactTOB, FactPiece, FactArticle, FactTiers, FactAdresse,FactureBtp,Splash,uEntCommun,UtilTOBPiece;

type
    TGenFacAvanc = class (Tobject)
      private
				MontantMarche,DejaFacture,MtHorsDevis : Double;
      	fCloture        : boolean;
        fDatefac        : TdateTime;
        Avancement      : boolean;
        NewNature       : string;
        NewSouche       : string;
        NewNumero       : integer;
        NewIndiceg      : integer;
        PiecesG         : TStrings ;
        DEV             : Rdevise;
        fCotraitance    : boolean;
        fLastfacture    : boolean;
        ModifSituation  : boolean;
        GenereAvoir : boolean;
        NumSituation : integer;
        PourcentAvanc : double; // pourcentage d'avancement global de la situation
        //
        TOBDevis,TOBFacture,TOBOptDevis,TOBTablette,TOBPieceTrait,TOBPieceTraitDevis,TOBAffaireInterv,TOBSousTrait,TOBAffaire : TOB;
        TOBArticles,TOBTiers,TOBEches,TOBAdresses,
        TOBCpta,TOBAnaP,TOBAnaS,TOBBases,TOBBasesL,TOBAcomptes,TOBPorcs,TOBComms,TOBLigneRG : TOB ;
        TOBOuvrage,TOBOuvragesP,TOBPIECERG,TOBBasesRg,TOBVTECOLL,TOBSituations : TOB;
        TOBAnalPiece  : TOB;
        TOBAcomptes_O : TOB;
        TOBtimbres    : TOB;
        IndiceRg      : integer;
        TheRepartTva  : TREPARTTVAMILL;
        OkCPta        : boolean;
        //
        procedure AppliqueInfoNewPiece;
        procedure AddChampNbUses(TOBPDOC : TOB);
        procedure AddReference(TOBPDoc: TOB);
        procedure ChargeLesAdresses(TOBDevis: TOB);
        procedure CompleteTobArticle(tOBL: TOB);
        procedure createTOBS;
        function  FindParamDoc (Cledoc : r_cledoc) : TOB;
        function  GetNumEvent : integer;
        procedure InitFacture (Multi : boolean);
        procedure InitToutModif;
        function  IsDetailInside (IndDep : integer;TOBL : TOB) : boolean;
        function  ISSameDoc(Courant, reference: r_cledoc): boolean;
        procedure LibereTobs;
        procedure NettoieFacture;
        function  OKLigneFacturable (TOBL,TOBPDoc : TOB) : boolean;
//        procedure RecupDetailOuvrage (TOBL: TOB);
        procedure RecupAnalytique(TOBL: TOB);
        procedure RenumeroteLaPiece;
        procedure SetAvancementFacture (TOBL : TOB);
        procedure SuprimeParagraphesOrphelin;
        procedure SuprimeThisParagraphe (TOBL : TOB; var IndDep : integer);
        procedure ChargeLesRG(TOBL: TOB);
        function  GetLeTauxDevise(CodeD: string; var DateTaux: TDateTime;DateP: TDateTime): Double;
        function  GetNumeroPiece: boolean;
        function  GetNewNumero(newNature,NewSouche: string): boolean;
        procedure SetInfoOuvrage (TOBPiece,TOBL : TOB;WithRenumOuv:boolean=false);
//        procedure TraiteArticleOuv (TOBL: TOB);
        procedure ValideThispiece;
        procedure ValideLesLignesFac;
        procedure AddChampsSupFac (TOBL : TOB);
        procedure GetAnalytiquePiece (TOBP : TOB);
        procedure GetArticlePieces(TOBPDoc : TOB);
        //
        function  ConstitueLafacture: boolean;
        function  RecalculeLapiece : boolean;
        function  ValidelaPiece : boolean;
        procedure GenereCompta;
        procedure AjouteEvent;
        procedure InitChampsSupFac(TOBL: TOB);
        procedure DupliqueLaLigne(NewTOBL, OldTOBL: TOB);
        procedure LoadLesGCS;
        procedure InitLigneFacturation(TOBL: TOB);
        procedure ZeroAchat(TOBfacture: TOB);
    		procedure InsereLigneremEsc(CledocR : R_CLEDOC);
        procedure GetlesPortsFromDevisprinc;
        procedure MajPiecePrecedente;
    		function  FindTOBArt(refArticle: string): TOB;
    		procedure SetInfoOuvrageFille(TOBPiece, TOBOUV, TOBL: TOB;Ligne: integer;WithRenumOuv : boolean=false);
    		procedure ValideLesOuv ( TOBOuv : TOB ; TOBPiece : TOB = nil) ;
    		procedure PositionneInfoPieceSousDetail(TOBpiece, TOBL: TOB);
    		procedure UpdateLesLignesDevis;
    		function ChargePieceTraitOrig(cledoc: R_CLEDOC;Fournisseur: string): TOB;
    		function FindPieceTraitOrig(Cledoc: r_cledoc;Fournisseur: string): TOB;
    		procedure SetLignesDevisOuvDetail(TOBL, TOBOuvrage, TOBLDEVIS: TOB;avancement: boolean);
        procedure ValideLesPieceTraitOrig;
        function PositionneEcartLastfacture ( var EcartConstate : boolean) : boolean;
    		function AjouteLigneEcart(MtEcart: double): boolean;
        procedure RestitueAvancementDevis;
        procedure SetPieceMorte(TOBFacture: TOB);

      public
        CodeErreur : string;
        property ClotureFacturation : boolean read fcloture write fcloture;
        property DateFac : TdateTime read fDateFac write fDateFac;
        constructor create;
        destructor  destroy; override;
        function    GenereLaFacture : boolean;
     end;


function GenereFactureFromAvanc (TOBSOUSTraits,TOBSource,TOBDetailDevis,TOBouvrages,TobAcomptes_O,TOBAcomptes : TOB; ClotureFacturation : boolean; var Erreur : string; var ClePiece : R_CLEDOC; DateFac : TdateTime; MontantMarche,DejaFacture,MtHorsDevis : Double; lastFacture : boolean=false; Cotraitance : boolean=false; ModifSituation : Boolean=false; GenereAvoir : boolean=false) : boolean;

implementation

uses factvariante,FactOuvrage,Ucotraitance,FactTimbres, Math, FactCommBTP, UCumulCollectifs,UspecifPOC,BTSAISPAIMENTPOC_TOF,factretenues
   ,CbpMCD
   ,CbpEnumerator
;

function GenereFactureFromAvanc (TOBSOUSTraits,TOBSource,TOBDetailDevis,TOBouvrages,TobAcomptes_O,TOBAcomptes : TOB; ClotureFacturation : boolean; var Erreur : string; var ClePiece : R_CLEDOC; DateFac : TdateTime; MontantMarche,DejaFacture,MtHorsDevis : Double; lastFacture : boolean=false; Cotraitance : boolean=false; ModifSituation : Boolean=false; GenereAvoir : boolean=false) : boolean;
var XX      : TgenFacAvanc;
begin
  XX := TgenFacAvanc.create;
  TRY
    XX.TOBDevis         := TOBSource;
    XX.TOBOptDevis      := TOBDetailDevis;
    XX.TOBAcomptes      := TOBAcomptes;
    XX.TOBAcomptes_O    := TobAcomptes_O;
    XX.TOBOuvrage       := TOBouvrages;
    XX.TOBSousTrait     :=  TOBSOUSTraits;
    XX.ClotureFacturation := ClotureFacturation;
    XX.DateFac          := DateFac;
    XX.fCotraitance     := Cotraitance;
    XX.MontantMarche    := MontantMarche;
    XX.MtHorsDevis      := MtHorsDevis;
    XX.DejaFacture      := DejaFacture;
    XX.Flastfacture     := lastFacture;
    XX.ModifSituation   := ModifSituation;
    XX.GenereAvoir      := GenereAvoir;
    result              := XX.GenereLaFacture;
    ClePiece.NaturePiece  := XX.NewNature;
    ClePiece.Souche       := XX.NewSouche;
    ClePiece.NumeroPiece  := XX.NewNumero;
    ClePiece.DatePiece    := DateFac;
  FINALLY
    Erreur := XX.CodeErreur;
    XX.free;
  end;
end;

{ TGenFacAvanc }

procedure TGenFacAvanc.AddReference(TOBPDoc: TOB);
Var NewL : TOB ;
    RefP : String ;
BEGIN
  if ModifSituation then Exit;
  NewL:=NewTOBLigne(TOBfacture,0) ;
  AddChampsSupFac (NewL);
  InitChampsSupFac (Newl);

  RefP:=RechDom('GCNATUREPIECEG',TOBPDoc.GetValue('GP_NATUREPIECEG'),False)
     +' N° '+IntToStr(TOBPDoc.GetValue('GP_NUMERO'))
     +' du '+DateToStr(TOBPDoc.GetValue('GP_DATEPIECE'))
     +'  '+TOBPDoc.GetValue('GP_REFINTERNE') ;
  RefP:=Copy(RefP,1,70) ;
  NewL.PutValue('GL_LIBELLE',RefP)    ; NewL.PutValue('GL_TYPELIGNE','COM') ;
  NewL.PutValue('GL_TYPEDIM','NOR')   ;
  NewL.PutValue('GL_TYPEARTICLE','EPO');
  // Pour pouvoir se pointer sur la piece d'origine
  NewL.PutValue('GL_NATUREPIECEG',TOBPDoc.getValue('GP_NATUREPIECEG'));
  NewL.PutValue('GL_SOUCHE',TOBPDoc.getValue('GP_SOUCHE'));
  NewL.PutValue('GL_NUMERO',TOBPDoc.getValue('GP_NUMERO'));
  NewL.PutValue('GL_INDICEG',TOBPDoc.getValue('GP_INDICEG'));
END ;

Procedure TGenFacAvanc.ChargeLesAdresses ( TOBDevis : TOB ) ;
var NumL, NumF : integer ;
BEGIN
TOBAdresses.cleardetail;

if GetParamSoc('SO_GCPIECEADRESSE') then
   BEGIN
   TOB.Create('PIECEADRESSE',TOBAdresses,-1) ; {Livraison}
   TOB.Create('PIECEADRESSE',TOBAdresses,-1) ; {Facturation}
   END else
   BEGIN
   TOB.Create('ADRESSES',TOBAdresses,-1) ; {Livraison}
   TOB.Create('ADRESSES',TOBAdresses,-1) ; {Facturation}
   END ;
LoadLesAdresses(TOBDevis,TOBAdresses) ;
TOBAdresses.SetAllModifie(True);
NumL:=TOBDevis.GetValue('GP_NUMADRESSELIVR'); NumF:=TOBDevis.GetValue('GP_NUMADRESSEFACT');
if GetParamSoc('SO_GCPIECEADRESSE') then
   BEGIN
   TOBFacture.PutValue('GP_NUMADRESSELIVR',+1) ;
   if NumF=NumL then TOBFacture.PutValue('GP_NUMADRESSEFACT',+1)
                else TOBFacture.PutValue('GP_NUMADRESSEFACT',+2);
   END else
   BEGIN
   TOBFacture.PutValue('GP_NUMADRESSELIVR',-1) ;
   if NumF=NumL then TOBFacture.PutValue('GP_NUMADRESSEFACT',-1)
                else TOBFacture.PutValue('GP_NUMADRESSEFACT',-2);
   END ;
END ;

function TGenFacAvanc.FindTOBArt(refArticle : string) : TOB;
var QQ : TQuery;
begin
  result:=TOBArticles.FindFirst(['GA_ARTICLE'],[RefArticle],False) ;
  if result=Nil then
  BEGIN
    result:=CreerTOBArt(TOBArticles) ;
    QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                   'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                   'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+RefArticle+'"',true);
    result.SelectDB('"'+RefArticle+'"',Nil) ;
    Ferme (QQ);
  END ;
end;

procedure TGenFacAvanc.CompleteTobArticle (tOBL : TOB);
var TOBART,TOBLAFFAIRE,TOBC : TOB;
    RefArticle : string;
begin
  TOBLaffaire := FindCetteAffaire (TOBL.GEtValue('GL_AFFAIRE')); // positionnement
  RefArticle := TOBL.GetValue('GL_ARTICLE');
  TOBART := FindTOBArt(refArticle);
  if (TOBArt<>nil) and (OkCpta) then
  begin
    TOBC:=ChargeAjouteCompta(TOBCpta,TOBfacture,TOBL,TOBArt,TOBTiers,nil,TOBLaffaire,True,Tobtablette) ;
    if TOBC <> nil then PreVentileLigne(TOBC,TOBAnaP,TOBAnaS,TOBL) ;
  end ;
end;

procedure  TGenFacAvanc.InsereLigneremEsc ( CledocR : R_CLEDOC);
var newl,TOBPDOC : TOB;
		LibRemEsc : string;
    rem,Esc : double;
begin
	TOBPDoc := FindParamDoc (CledocR);
  if TOBPdoc = nil then exit;
  Rem := TOBPDOC.GetValue('GP_REMISEPIED');
  Esc := TOBPDOC.GetValue('GP_ESCOMPTE');
  if (rem = 0) and (Esc = 0 ) then exit;
  //
  NewL := NewTOBLigne (TOBfacture,0,true);
  AddChampsSupFac (NewL);
  InitChampsSupFac (Newl);
  NewL.PutValue('GL_TYPELIGNE','COM') ;
  NewL.PutValue('GL_TYPEDIM','NOR')   ;
  NewL.PutValue('GL_TYPEARTICLE','RRR');
  // Pour pouvoir se pointer sur la piece d'origine
  NewL.PutValue('GL_NATUREPIECEG',TOBPDOC.getValue('GP_NATUREPIECEG'));
  NewL.PutValue('GL_SOUCHE',TOBPDOC.getValue('GP_SOUCHE'));
  NewL.PutValue('GL_NUMERO',TOBPDOC.getValue('GP_NUMERO'));
  NewL.PutValue('GL_INDICEG',TOBPDOC.getValue('GP_INDICEG'));
  if Rem <> 0 then
  begin
    LibremEsc :='Remise de '+FloatToStr(Rem)+'%'
  end;
  if Esc <> 0 then
  begin
  	if Rem = 0 then
    begin
    	LibremEsc := 'Escompte de '+FloatToStr(Esc)+'%'
    end else
    begin
    	LibremEsc := LibremEsc + ' Escompte de '+FloatToStr(Esc)+'%'
    end;
  end;
  NewL.PutValue('GL_LIBELLE',LibRemEsc)    ;
  NewL := NewTOBLigne (TOBfacture,0,true);
  AddChampsSupFac (NewL);
  InitChampsSupFac (Newl);
  // Pour pouvoir se pointer sur la piece d'origine
  NewL.PutValue('GL_NATUREPIECEG',TOBPDOC.getValue('GP_NATUREPIECEG'));
  NewL.PutValue('GL_SOUCHE',TOBPDOC.getValue('GP_SOUCHE'));
  NewL.PutValue('GL_NUMERO',TOBPDOC.getValue('GP_NUMERO'));
  NewL.PutValue('GL_INDICEG',TOBPDOC.getValue('GP_INDICEG'));
  NewL.PutValue('GL_TYPELIGNE','COM') ;
  NewL.PutValue('GL_TYPEDIM','NOR')   ;
end;

function TGenFacAvanc.ConstitueLafacture : boolean;
var CledocR,Cledoc : r_cledoc;
    TOBPDoc : TOB;
    Indice : integer;
    TOBL,TOBLF : TOB;
    TypeCom : string;
    Ouvrage : boolean;
    PassP : string;
    zCledoc : R_CLEDOC;
    lastPrec,CurrPiece : string;
    multiDoc : boolean;
    TouteLigne : Boolean;
    MontantCumule : double;
begin
  multiDoc := false;
	lastPrec := '';
  CurrPiece := '';
  TOBPDoc := nil;
  FillChar (cledoc,sizeof(cledoc),#0);
  FillChar (cledocR,sizeof(cledoc),#0);
  result := true;
  TouteLigne := (TOBdevis.GetValue('AFF_OKSIZERO')='X') or (GetParamSocSecur('SO_BTGESTVIDESUREDIT',false));
  CodeErreur := 'Erreur en cours de génération';
  if (Pos(TOBdevis.GetValue('AFF_GENERAUTO'),'AVA;DIR;')>0) or (TOBdevis.GetValue('AFF_GENERAUTO')='') then
  begin
    if GetParamSocSecur('SO_BTSITPROVISOIRE',false) then
    begin
      if GenereAvoir then NewNature := 'ABP'
                     else NewNature := 'FBP';
    end else
    begin
      if GenereAvoir then NewNature := 'ABT'
                     else NewNature := 'FBT';
    end;
    if GenereAvoir then NumSItuation := GetLastNumSituation (TOBDevis);

  end else
  begin
  	if (not ClotureFacturation) then
    begin
      NewNature := 'DAC';
    end else
    begin
      if GenereAvoir then NewNature := 'ABT'
                     else NewNature := 'FBT';
    end;
  end;
  //
  if not result then Exit;
  //
  //FV1 - 13/02/2018 : FS#2943 - CLOSSUR - Erreur de numéro de situation pour la génération d'avoir à partir de préparation
  if GenereAvoir then NumSItuation := NumSItuation + 1;
  //
  PassP := GetInfoParPiece(NewNature, 'GPP_TYPEECRCPTA');
  OkCpta := ((PassP <> '') and (PassP <> 'RIE'));
  //
   if RemplirTOBTiers(TOBTiers,TOBdevis.getValue('GP_TIERS'),NewNature,true)<>trtOk then
  begin
    CodeErreur := 'ATTENTION : Tiers inexistant ou fermé';
    Result := False;
    exit;
  end;
  LoadLesGCS;
  //
  TOBfacture.dupliquer (TOBDevis,false,true); AddLesSupEntete (TOBfacture);
  // et hop l'oublie qui tue
  GetlesPortsFromDevisprinc ;
  //
  InitFacture (TOBOptdevis.detail.count > 1);
  // TVA Millieme
  TheRepartTva.Initialise;
  TheRepartTva.TOBBases := TOBBases;
  TheRepartTva.TOBpiece := TOBFacture;
  TheRepartTva.TOBOuvrages := TOBOuvrage;
  TheRepartTva.Charge; // voila voila
  //
  ChargeLesAdresses (TOBdevis);
  TypeCom:=GetInfoParPiece(TOBDevis.GetValue('GP_NATUREPIECEG'), 'GPP_TYPECOMMERCIAL') ;
  TOBfacture.putValue('GP_RGMULTIPLE','-');

  //
  MontantCumule := 0;

  if TOBdevis.getValue('GP_AFFAIRE')<> '' then
  begin
    StockeCetteAffaire (TOBdevis.getString('GP_AFFAIRE'));
    TOBAffaire := FindCetteAffaire(TOBdevis.getString('GP_AFFAIRE'))
  end;


  for Indice := 0 to TOBdevis.detail.count -1 do
  begin
    TOBL := TOBDevis.detail[Indice];
    if TOBL.GetValue('GL_TYPELIGNE')='REF' then
    begin
      continue; // on saute les references de pieces de la saisie d'avancement
    end;
    cledoc := TOB2CleDoc (TOBL);
    if not ISSameDoc(cledoc,CledocR) then
    begin
    	if (Cledoc.NaturePiece <> '') and (TOBOptdevis.detail.count > 1) then InsereLigneremEsc (cledocR);
      // rupture  sur document... on va chercher les parametres de celui-ci
      TOBPDoc := FindParamDoc (Cledoc);
      // Ajout de la référence du devis
      if TOBPDoc <> nil then
      begin
        AddReference(TOBPDoc);
        AddChampNbUses(TOBPDoc);
        GetAnalytiquePiece (TOBPDoc);
        GetArticlePieces(TOBPDoc);
      end;
//      ChargeAcompte (Cledoc); // Acompte de la piece
      CledocR := Cledoc ;
    end;
    if not OKLigneFacturable (TOBL,TOBPDoc) then continue; // Ligne a zéro et param lignea zero sur situation a non
    //
    TOBLF := NewTOBLigne (TOBfacture,0,true);
//    AddChampsSupFac (TOBLF);
    InitChampsSupFac (TOBLF);
    DupliqueLaLigne (TOBLF,TOBL);
    // correction FQ 12030
    if (TOBOptdevis.detail.count > 1) then TOBLF.putValue('GL_ACTIONLIGNE','');
    TOBLF.putValue('GL_FACTUREHT',TOBL.getValue('GL_FACTUREHT'));
    //
    NewTOBLigneFille(TOBLF);

    //
    if TOBLF.getValue('GL_CREERPAR')<>'FAC' then
    begin
      if not ModifSituation then
      begin
        TOBLF.PutValue ('GL_PIECEPRECEDENTE',encoderefPiece(TOBLF,0));
        //
      end;
      if TOBLF.getValue('GL_PIECEPRECEDENTE') <> '' then
      begin
        DecodeRefPiece(TOBLF.getValue('GL_PIECEPRECEDENTE'),zCledoc);
        CurrPiece := zCledoc.NaturePiece+';'+zcledoc.Souche +';'+IntToStr(zCledoc.NumeroPiece)+';'+IntToStr(zCledoc.Indice);
        if lastPrec = '' then
        begin
          lastprec := CurrPiece;
        end;
      end;
      //
      if not ModifSituation then
      begin
        if TOBLF.GetValue('GL_PIECEORIGINE')='' then TOBLF.PutValue ('GL_PIECEORIGINE',encoderefPiece(TOBLF,0));
      end;
    end;
    //
    AjouteRepres(TOBLF.GetValue('GL_REPRESENTANT'),TypeCom,TOBComms) ;
    //
    Ouvrage := IsOuvrage (TOBLF) AND (not IsSousDetail (TOBLF));
    if  (Ouvrage) or ( IsArticle (TOBLF) and (TOBLF.GetValue('GL_TYPELIGNE')<>'RG') ) then
    begin
      RecupAnalytique (TOBLF);
      CompleteTobArticle (TOBLF);
      ZeroLigneMontant (TOBLF);  // reinit des montants de la ligne
      InitLigneFacturation (TOBLF);
      if (TOBLF.getValue('BLF_MTSITUATION')<> 0) or (TOBLF.getValue('BLF_MTDEJAFACT')<>0) or (TouteLigne) then
      begin
        TOBPDoc.PutValue ('NB LIGNES FACTURE',TOBPDoc.GetValue ('NB LIGNES FACTURE')+1);
        // Initialisation en fonction de l'avancement
        TOBLF.putValue('BLF_NATUREPIECEG',NewNature);
        SetAvancementFacture (TOBLF);
        if (TOBLF.getValue('BLF_MTSITUATION')<> 0) or (TouteLigne) then
        begin
          if Avancement then
          begin
            TOBLF.putValue('GL_QTESIT',TOBLF.GetValue('BLF_QTECUMULEFACT'));
            TOBLF.putValue('GL_QTEPREVAVANC',TOBLF.GetValue('BLF_QTEMARCHE'));
          end else
          begin
            TOBLF.putValue('GL_QTESIT',TOBLF.GetValue('BLF_QTESITUATION'));
            TOBLF.putValue('GL_QTEPREVAVANC',TOBLF.GetValue('BLF_QTESITUATION'));
          end;
        end;
      end;
      //
      if VH_GC.BTCODESPECIF = '001' then
      begin
        if not IsLigneSoustraite (TOBL) then MontantCumule := MontantCumule + TOBLF.getDouble('BLF_MTCUMULEFACT');
      end else
      begin
        MontantCumule := MontantCumule + TOBLF.getDouble('BLF_MTCUMULEFACT');
      end;
      //

      TheRepartTva.AppliqueLig (TOBLF);
      //
    end else if TOBLF.GetValue('GL_TYPELIGNE')='RG' then
    begin
      ChargeLesRG (TOBLF);
    end
  end;
  //
  if Avancement then PourcentAvanc := (MontantCumule / MontantMarche) * 100;
  //
  if (TOBOptdevis.detail.count > 1) then InsereLigneremEsc (cledoc);
  //
  NettoieFacture; // enleve les lignes qui ne sont pas valorisé depuis un devis ou aucune ligne facturée
  if not GetParamSocSecur('SO_BTGESTVIDESUREDIT',false) then
  begin
    SuprimeParagraphesOrphelin; // enleve les paragraphes vides
  end;
  //
  for Indice := 0 to TOBfacture.detail.count -1 do
  begin
    TOBLF := TOBfacture.detail[Indice];
    if TOBLF.getValue('GL_CREERPAR')<>'FAC' then
    begin
      if not ModifSituation then
      begin
        TOBLF.PutValue ('GL_PIECEPRECEDENTE',encoderefPiece(TOBLF,0));
      end;
      //
      if TOBLF.getValue('GL_PIECEPRECEDENTE') = '' then Continue;
      DecodeRefPiece(TOBLF.getValue('GL_PIECEPRECEDENTE'),zCledoc);
      CurrPiece := zCledoc.NaturePiece+';'+zcledoc.Souche +';'+IntToStr(zCledoc.NumeroPiece)+';'+IntToStr(zCledoc.Indice);
      if lastPrec = '' then
      begin
        lastprec := CurrPiece;
      end else if lastPrec <> CurrPiece then
      begin
        Multidoc := true;
      end;
      //
    end;
  end;
  //
  if (TOBFacture.findFirst(['GL_TYPELIGNE'],['RG'],true)<>nil) and (MultiDoc) then
  begin
  	TOBfacture.putValue('GP_RGMULTIPLE','X');
  end;

  if not GetNumeroPiece then
  begin
    CodeErreur := 'Erreur en récupération de numéro de pièce';
    Result := False;
    exit;
  end;
  //
  UG_MajLesComms (TOBfacture,TOBArticles,TOBComms); // mise à jour des commissions
  // phase 2
  AppliqueInfoNewPiece;
  //
  RenumeroteLaPiece;
  //
  //
  ValideAnalytiques (TOBfacture,TOBAnaP,TOBAnaS);
  //
  InitToutModif;
  result := true;
end;

constructor TGenFacAvanc.create;
begin
  NumSituation := -1;
  TheRepartTva := TREPARTTVAMILL.create;
  createTOBS;
  IndiceRg := 1;
  NewNature := '';
  ReinitTOBAffaires;
  PiecesG := TStringList.create;
  InitParamTimbres;
end;

procedure TGenFacAvanc.createTOBS;
begin
  TOBSituations := TOB.Create('LES SITUATIONS',nil,-1);
  TOBLigneRG := TOB.Create ('LES LIGNES RG',nil,-1); 
  TOBFacture:= TOB.Create ('PIECE',nil,-1); AddLesSupEntete (TOBfacture);
  TOBArticles:=TOB.Create('',Nil,-1) ;
  TOBBases:=TOB.Create('BASES',Nil,-1) ;
  TOBBasesL:=TOB.Create('LES BASES LIGNE',Nil,-1) ;
  TOBTiers:=TOB.Create('TIERS',Nil,-1) ; TOBTiers.AddChampSup('RIB',False) ;
  TOBEches:=TOB.Create('Les ECHEANCES',Nil,-1) ;
  TOBPorcs:=TOB.Create('',Nil,-1) ;
//  TOBOuvrage := TOB.create ('OUVRAGES',nil,-1);
  TOBOuvragesP := TOB.create ('OUVRAGES PLAT',nil,-1);
  TOBPIECERG := TOB.create ('RETENUES',nil,-1);
  TOBBasesRG := TOB.create ('BASESRG',nil,-1);
  TOBVTECOLL := TOB.create ('LES COLL',nil,-1);
  TOBComms:=TOB.Create('COMMERCIAUX',Nil,-1) ;
  TOBCpta:=CreerTOBCpta ;
  TOBAnaP:=TOB.Create('',Nil,-1) ; TOBAnaS:=TOB.Create('',Nil,-1) ;
  TOBAdresses:=TOB.Create('',Nil,-1);
  TOBAnalPiece := TOB.create ('ANALYTIQUES PIECES',nil,-1);
  TOBTablette:= TOB.Create ('LES TABLETTES',nil,-1);
  TOBpieceTrait:= TOB.Create ('LES PIECE EXTE',nil,-1);
  TOBAffaireInterv := TOB.Create ('LES Co-SOUSTRAITANTS',nil,-1);
  TOBPieceTraitDevis := TOB.Create ('LES PIECES EXTE ORIG',nil,-1);
  TOBtimbres := TOB.Create('LES TIMBRES',nil,-1);

end;

destructor TGenFacAvanc.destroy;
begin
	LibereParamTimbres;
  LibereTobs;
  ReinitTOBAffaires;
  PiecesG.free;
  TheRepartTva.free;
  TOBpieceTrait.free;
  TOBAffaireInterv.free;
  inherited;
end;

function TGenFacAvanc.FindParamDoc(Cledoc: r_cledoc): TOB;
begin
  result := TOBOptDevis.findfirst (['GP_NATUREPIECEG','GP_SOUCHE','GP_NUMERO','GP_INDICEG'],
                                   [cledoc.naturepiece,cledoc.souche, Cledoc.NumeroPiece,Cledoc.Indice],true);
end;

function TGenFacAvanc.GenereLaFacture: boolean;
var XX : TFsplashScreen;
		ecart : boolean;
    MtPaiement,XP,XD,TheMontantRGTTC,TOTALRETENUHT,TOTALRETENUTTC,TheMontantTimbres : double;
    ApplicRetenue : boolean;
begin
  //
  ecart := false;
  LoadLaTOBAffaireInterv (TOBAffaireInterv,TOBDevis.getvalue('GP_AFFAIRE'));
  XX := TFsplashScreen.create(application);
  if GenereAvoir then XX.Label1.caption := 'Génération de l''avoir en cours...'
                 else XX.Label1.caption := 'Génération de la facture en cours...';
  XX.Animate1.Active := true;
  XX.Show;
  XX.Refresh;
  TRY
    if (Pos(TOBdevis.GetValue('AFF_GENERAUTO'),'AVA;DAC;')>0) then
    begin
      Avancement := True;
    end else
    begin
    	Avancement := false;
    end;
    //
    result := ConstitueLafacture; if not result then exit;
    if result then
    begin
      ApplicRetenue := (GetInfoParPiece(TOBFacture.GetString('GP_NATUREPIECEG'), 'GPP_APPLICRG') = 'X');
      DEV.Code := TOBFacture.GetValue('GP_DEVISE');
      GetInfosDevise(DEV);
      DEV.Taux := GetLeTauxDevise(DEV.Code, DEV.DateTaux, TOBFacture.GetValue('GP_DATEPIECE'));
      ChargeLesTimbres (TOBFacture,TOBTimbres);
      result := RecalculeLapiece; if not result then exit;
    end;
    if result then
    begin
      if VH_GC.BTCODESPECIF = '001' then
      begin
        NumSituation := GetNextNumSituation(TOBfacture);
        // ---- Constitue la situation 0 si elle n'existe pas -----
        if NumSituation = -1 then
        begin
          // on est en train de générer la première situation donc c'est bon
          EnregistreAvancePOC (TOBFacture,TOBAffaire,TOBPorcs,TOBBases,DEV);
        end;
        // application de la retenue de garantie en provenance d ela fiche affaire
        AppliqueRGPOC (TOBfacture,TOBAffaire,TOBPorcs,TOBTiers,TOBPIECERG,DEV,NumSituation);
        RecalculeRG(TOBPorcs,TOBFacture, TOBPieceRG, TOBBases, TOBBasesRG,TOBpieceTrait, DEV);
        // ----
        // calcul de la restitution d'avance suivant le mode de calcul POC
        // permet de restituer 100 % de l'avance entre un % mini et un % maxi
        // Soit : (% Avancement - %Debut restitution) / (%Fin de restitution - %Debut restitution) = % de restitution sur cette facture
        if Avancement then
        begin
          Result := RestitueAvancePOC(TOBFacture,TOBAffaire,TOBPorcs,TOBBases,PourcentAvanc,DEV,NumSituation);
          if not result then exit;
        end;
        // ENFIN Recalcul global de la pièce
        result := RecalculeLapiece; if not result then exit;
        InsereLigneRG(TOBFacture, TOBPIeceRG, TOBBASES, TOBBASESRG, TOBTIers, TOBAffaire, TOBLigneRG, -1);
      end;

      (* Nouvel emplacement pour la gestion de l'écart *)
      if (fLastfacture) and (avancement) then
      begin
        Result := PositionneEcartLastfacture (ecart); if not result then exit;
        if (Result) and (ecart) then
        begin
          ReinitReglPieceTrait(TOBpiecetrait);
        	Result := RecalculeLapiece;
        end;
      end;
      
      if VH_GC.BTCODESPECIF = '001' then
      begin
        GetMontantRGReliquat(TOBPIeceRG, XD, XP,true);
        TheMontantRGTTC := XD;
        GetRetenuesCollectif (TOBPorcs,XD,XP,'HT');
        TOTALRETENUHT := Xd;
        GetRetenuesCollectif (TOBPorcs,XD,XP,'TTC');
        TOTALRETENUTTC := Xd;
    //    GetRetenuesCollectif (TOBPorcs,RXD,RXP);
        // Recup du montant de timbres
        TheMontantTimbres := TOBFacture.GetDouble('GP_TOTALTIMBRES');
        // Modif BTP
        if ((TheMontantRGTTC <> 0) or (TOTALRETENUTTC<>0)) and (ApplicRetenue) then
        begin
          MtPaiement := TOBfacture.GetValue('GP_TOTALTTCDEV') - TheMontantRGTTC + TheMontantTimbres - TOBfacture.GetValue('GP_ACOMPTEDEV') - TOTALRETENUTTC;
        end else
        begin
          // ---
          MtPaiement := TOBfacture.GetValue('GP_TOTALTTCDEV') + TheMontantTimbres - TOBfacture.GetValue('GP_ACOMPTEDEV');
        end;
        //
        result := DemandeReglementSousTrait(TOBfacture,TOBTiers,TOBEches,TOBAcomptes,TObpieceTrait,TOBPieceRG,TOBSousTrait,TOBPorcs, MtPaiement ,DEV );
        if not result then exit;
      end;

      (* ---------------------------------------------- *)
      if Result then  Result := ValidelaPiece;
      if not result then
      begin
        exit;
      end else
      begin
        NewNature := TOBfacture.getString('GP_NATUREPIECEG');
        NewSouche := TOBfacture.getString('GP_SOUCHE');
        NewNumero := TOBfacture.getInteger('GP_NUMERO');
        NewIndiceg := TOBfacture.getInteger('GP_INDICEG');
        AjouteEvent ;
    end;
    end;
    CodeErreur:='Le traitement s''est correctement effectué. #13#10 '+
                 RechDom('GCNATUREPIECEG',NewNature,False)+' N° '+IntToStr(NewNumero)+
                 ' générée avec succès';
  FINALLY
    XX.free;
    ReinitTOBAffaires;
//    PieceTraitUsable := false;
  END;
end;

procedure TGenFacAvanc.InitFacture (Multi : boolean);
begin
  if Multi then
  begin
    TOBfacture.PutValue ('GP_REMISEPIED',0);
    TOBfacture.PutValue ('GP_ESCOMPTE',0);
  end;
  ZeroFacture (TOBFacture);
  StockeCetteAffaire (string(TOBfacture.GetValue('GP_AFFAIRE')));
  StockeCetteAffaire (string(TOBfacture.GetValue('GP_AFFAIREDEVIS')));
//  TOBfacture.putvalue ('GP_FACTUREHT','X'); --> correction LS pour facture TTC
  TOBfacture.putvalue ('GP_PIECEFRAIS','');
  TOBfacture.putvalue ('GP_EDITEE','-');
  TOBfacture.putvalue ('GP_REFCOMPTABLE','');
  TOBfacture.putvalue ('GP_REFCPTASTOCK','');
  TOBfacture.putvalue ('GP_DEVENIRPIECE','');
  TOBfacture.putvalue ('GP_JALCOMPTABLE','');
  TOBfacture.putvalue ('GP_VIVANTE','X');
end;


function TGenFacAvanc.ISSameDoc (Courant,reference : r_cledoc) : boolean;
begin
  result := (Courant.NaturePiece  = reference.NaturePiece) and
            (Courant.Souche = Reference.Souche) and
            (Courant.NumeroPiece = Reference.NumeroPiece ) and
            (Courant.Indice  = Reference.Indice );
end;

procedure TGenFacAvanc.LibereTobs;
begin
  TOBSituations.Free;
  TOBFacture.free;
  TOBArticles.free;
  TOBBases.free;
  TOBBasesL.free;
  TOBTiers.free;
  TOBEches.free;
  TOBAdresses.free;
  TOBCpta.free;
  TOBAnaP.free;
  TOBAnaS.free;
  TOBPorcs.free;
  TOBComms.free;
//  TOBOuvrage.free;
  TOBOuvragesP.free;
  TOBPIECERG.free;
  TOBBasesRg.free;
  TOBVTECOLL.Free;
  TOBAnalPiece.free;
  TOBTablette.free;
  TOBPieceTraitDevis.free;
  TOBLigneRG.Free;
end;

function TGenFacAvanc.OKLigneFacturable(TOBL, TOBPDoc: TOB): boolean;
begin
  result := true;
  if TOBL.GetValue('GL_TYPELIGNE')<>'ART' then exit;
  if GetParamSocSecur('SO_BTGESTVIDESUREDIT',false) then exit; // si on gére les lignes vide sur édition --> on prend tout
  if (TOBL.GetValue('BLF_MTDEJAFACT')=0) and (TOBL.GetValue('BLF_MTSITUATION')=0) and (TOBPDOC.GetValue('AFF_OKSIZERO')='-') then result := false;
end;

procedure TGenFacAvanc.RecupAnalytique (TOBL : TOB);
Var RefA : String ;
    NumL : integer ;
    TOBAL,TOBLAL,TOBALF : TOB ;
BEGIN
  if TOBL=Nil then Exit ;
  if TOBL.detail.count = 0 then NewTOBLigneFille (TOBL);
  NumL:=TOBL.GetValue('GL_NUMLIGNE') ;
  RefA:=EncodeRefCPGescom(TOBL) ; if RefA='' then Exit ;
  TOBLAL := TOBAnalPiece.findFirst(['YVA_TABLEANA','YVA_IDENTIFIANT','YVA_IDENTLIGNE'],
                                   ['GL',RefA,FormatFloat('000',NumL)],true);
  if TOBLAL = nil then
  begin
  	
  end;

  if TOBLAL <> nil then
  begin
    if TOBL.Detail.Count<=0 then TOBAL:=TOB.Create('',TOBL,-1) else TOBAL:=TOBL.Detail[0] ;
    repeat
      TOBALF := TOB.create ('VENTANA',TOBAL,-1);
      TOBALF.dupliquer(TOBLAL,true,true);
      TOBLAL := TOBAnalPiece.findNext(['YVA_TABLEANA','YVA_IDENTIFIANT','YVA_IDENTLIGNE'],
                                        ['GL',RefA,FormatFloat('000',NumL)],true);
    until TOBLAL = nil;
  end;

END ;

procedure TGenFacAvanc.SetAvancementFacture(TOBL: TOB);
begin
  TOBL.putvalue('GL_QTEFACT',      TOBL.GEtValue('BLF_QTESITUATION'));
  TOBL.putvalue('GL_QTERESTE',     TOBL.GEtValue('BLF_QTESITUATION'));
  TOBL.putvalue('GL_QTESTOCK',     TOBL.GEtValue('BLF_QTESITUATION'));
  TOBL.putvalue('GL_QTERELIQUAT',0);
  TOBL.putvalue('GL_MONTANTHTDEV', TOBL.GEtValue('BLF_MTSITUATION'));
  TOBL.putvalue('GL_MTRESTE',      TOBL.GEtValue('BLF_MTSITUATION'));
  TOBL.putvalue('GL_MTRELIQUAT', 0);
  TOBL.putValue('GL_POURCENTAVANC',TOBL.GetValue('BLF_POURCENTAVANC'));
end;

procedure TGenFacAvanc.SuprimeParagraphesOrphelin;
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

function TGenFacAvanc.IsDetailInside(IndDep: integer; TOBL: TOB): boolean;
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

procedure TGenFacAvanc.SuprimeThisParagraphe(TOBL : TOB; var IndDep: integer);
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

procedure TGenFacAvanc.ChargeLesRG(TOBL:TOB) ;
Var Q : TQuery ;
    CD : R_CleDoc ;
    i  : integer ;
    TOBDocP,TOBInsere,TOBPiece : TOB ;
    CautionU,CautionUdev : double;
BEGIN
  if TOBL.getValue('GL_TYPELIGNE')<>'RG' then exit;
  TOBPiece := TOBL.Parent;
  TOBDocP := TOB.create ('LATOB',nil,-1);
  CD:=TOB2CleDoc(TOBL) ;
  Q:=OpenSQL('SELECT * FROM PIECERG WHERE '+WherePiece(CD,ttdRetenuG,False)+' AND PRG_NUMLIGNE='+inttostr(TOBL.GetValue('GL_NUMORDRE')),True) ;
  if Not Q.EOF then TOBDocP.LoadDetailDB('PIECERG','','',Q,True,False) ;
  Ferme(Q) ;
  for i:= 0 to TOBDocP.detail.count -1 do
  begin
    TOBInsere:=TOB.create ('PIECERG',TOBPieceRG,-1);
    TOBInsere.Dupliquer (TOBDocP.detail[i],true,true);
    TOBInsere.AddChampSupValeur ('INDICERG',IndiceRG);
    TOBInsere.AddChampSupValeur ('PIECEPRECEDENTE',encoderefPiece(TOBDOCP.detail[i],0));
    TOBInsere.AddChampSupValeur ('NUMORDRE',0,false);
    TOBInsere.putValue('PRG_CAUTIONMTU',0);
    TOBInsere.putValue('PRG_CAUTIONMTUDEV',0);
    //
    CautionU :=0; CautionUdev := 0;
    GetCautionAlreadyUsed (TOBPiece,TOBL,TOBL.GEtValue('GL_PIECEPRECEDENTE'),CautionU,CautionUdev,TOBInsere.getString('PRG_FOURN'),true);
    TOBInsere.AddChampSupValeur ('CAUTIONUTIL',CautionU);
    TOBInsere.AddChampSupValeur ('CAUTIONUTILDEV',CautionUDev);
    TOBInsere.AddChampSupValeur ('CAUTIONAPRES',0);
    TOBInsere.AddChampSupValeur ('CAUTIONAPRESDEV',0);
    //
    TOBL.PutValue ('INDICERG',IndiceRG);
  end;
  if TOBDocP.detail.count > 0 then inc(IndiceRg);
  TOBDocP.free;
end;

procedure TGenFacAvanc.AddChampNbUses(TOBPDOC: TOB);
begin
	if not TOBPDOC.FieldExists('NB LIGNES FACTURE') then
  begin
  	TOBPDOC.AddChampSupValeur ('NB LIGNES FACTURE',0);
  end;
end;

procedure TGenFacAvanc.NettoieFacture;
var TOBPDoc,TOBLF : TOB;
    cledoc,CledocR : r_cledoc;
    indice : integer;
begin
  TOBPDoc := nil;
  fillchar(cledoc,sizeof(cledoc),#0);
  fillchar(cledocR,sizeof(cledocR),#0);
  Indice := 0;
  repeat
    TOBLF := TOBFacture.detail[Indice];
    cledoc := TOB2CleDoc (TOBLF);
    if not ISSameDoc(cledoc,CledocR) then
    begin
      TOBPDoc := FindParamDoc (Cledoc);
      cledocR := cledoc;
    end;
    if (IsSousDetail(TOBLF)) or (IsLigneDetail(nil,TOBLF)) then // on enleve les sous détail et les lignes de detail d'ouvrages commentaires
    begin
      TOBLF.free;
    end else if (TOBPDoc.GetValue('NB LIGNES FACTURE')=0) and ( not IsQteZeroAutorise(TOBPDoc)) then
    begin
      if TOBLF.getValue('GL_TYPELIGNE')='RG' then
      begin
        SupprimeRg (TOBPIECERG,TOBBasesRg,TOBLF.GetValue('INDICERG'));
      end;
      TOBLF.free;
    end else Inc(Indice);
  until Indice >= TOBFacture.detail.count;
  //
  // Gestion des paragraphes
  //
end;

function TGenFacAvanc.GetnewNumero (NewNature,NewSouche : string) : boolean;
begin
//  result := false;
  NewNumero := 0;
  NewNumero := SetNumberAttribution(NewNature,newSouche, fDatefac,1);
  Result := True;
  (*
  repeat
    Q := OpenSQL('SELECT SH_NUMDEPART FROM SOUCHE WHERE SH_TYPE="GES" AND SH_SOUCHE="' + NewSouche + '"', True);
    if not Q.EOF then NewNumero := Q.Fields[0].AsInteger;
    Ferme(Q);
    Req := 'UPDATE SOUCHE SET SH_NUMDEPART='+
            IntToStr(NewNumero+1)+
            ' WHERE SH_TYPE="GES" AND SH_SOUCHE="' + NewSouche + '" AND SH_NUMDEPART='+
            IntToStr(NewNumero);
    if ExecuteSQL (Req) < 0 then
    begin
      inc(CptRetry);
    end else
    begin
      result := true;
    end;
  until (result) or (CptRetry=3);
  *)
end;

function TGenFacAvanc.GetNumeroPiece: boolean;
begin
  result := true;
  NewSouche := GetSoucheG (NewNature,TOBFacture.getValue('GP_ETABLISSEMENT'),TOBFacture.getValue('GP_DOMAINE'));
  if NewSouche = '' then
  begin
    CodeErreur := 'problème de souche de pièce';
    result := false;
    exit;
  end;
  if not GetnewNumero (NewNature,NewSouche) then
  begin
    CodeErreur := 'problème de numérotation de pièce';
    result := false;
    exit;
  end;
end;

procedure TGenFacAvanc.AppliqueInfoNewPiece;
var Indice : integer;
		TOBLF : TOB;
begin
  TOBFacture.putvalue('GP_NATUREPIECEG',NewNature);
  TOBFacture.putvalue('GP_SOUCHE',NewSouche);
  TOBFacture.putvalue('GP_NUMERO',NewNumero);
  TOBFacture.putvalue('GP_INDICEG',0);
  TobFacture.putvalue('GP_CODEORDRE',0);
  ValideLaPeriode (TOBFacture);
  TOBFacture.PutValue('GP_DATEPIECE',fDatefac) ;
  //
  if GetInfoParPiece(NewNature,'GPP_ACTIONFINI')='ENR' then
    TOBfacture.PutValue('GP_VIVANTE','-')
  else
    TOBfacture.PutValue('GP_VIVANTE','X') ;
  // FQ 12044 - reprise du mode de gestion (PA / Pv) de la facture générée
  TOBfacture.PutValue('GP_PIECEENPA',GetInfoParPiece(NewNature,'GPP_FORCEENPA'));
  //
  TOBFacture.PutValue('GP_DATECREATION',V_PGI.DateEntree) ;
  TOBFacture.PutValue('GP_CREATEUR',V_PGI.User);
  TOBFacture.PutValue('GP_DATEMODIF',V_PGI.DateEntree) ;
  TOBFacture.PutValue('GP_ETATVISA','NON') ;
  TOBFacture.PutValue('GP_EDITEE','-') ;
  TOBFacture.PutValue('GP_CREEPAR','GEN') ;

  ZeroFacture(TOBFacture) ;
  for Indice := 0 to TOBFacture.detail.count -1 do
  begin
    TOBLF := TOBFacture.detail[Indice];
  	if isOuvrage(TOBLF) then PositionneInfoPieceSousDetail (TOBFacture,TOBLF);
  end;
end;

procedure TGenFacAvanc.RenumeroteLaPiece;
var Indice : integer;
//    LastNumOrdre : integer;
    TypeLigne : string;
    TOBL,TOBRG : TOB;
    MaxNumOrdre : integer;
    IOrdre,ILigne,IndiceRG : integer;
    RefA : string;
begin
  TOBfacture.putValue('GP_UNIQUEBLO',0);
  MaxNumOrdre := 1;
  RefA:=EncodeRefCPGescom(TOBFacture) ;
  IOrdre := -1;
  ILigne := -1;
  //
  for Indice := 0 to TOBFacture.detail.count -1 do
  begin
    TOBL := TOBfacture.detail[Indice];
    if Indice = 0 then
    begin
      IOrdre := TOBL.GetNumChamp ('GL_NUMORDRE');
      ILigne := TOBL.GetNumChamp ('GL_NUMLIGNE');
    end;
//    LastNumOrdre := TOBL.GetValeur(Iordre);
    TypeLigne := TOBL.GetValue('GL_TYPELIGNE');
    TOBL.PutValeur(IOrdre,MaxNumordre);
    TOBL.PutValeur(Iligne,MaxNumordre);
    Inc(MaxNumOrdre);
    // on en profite pour lui coller subreptissement les infos de la pièce
    UG_PieceVersLigne (TOBfacture,TOBL,false);
    UG_MajAnalLigne (TOBL,RefA);
    //
    if (TypeLigne='RG') then
    begin
      IndiceRg := TOBL.getValue('INDICERG');
      if IndiceRg <>0 then
      begin
      	TOBRG:=TOBPieceRG.findfirst(['INDICERG'],[IndiceRg],true);
        repeat
          if TOBRG <> nil then
          begin
            TOBRG.putValue('NUMORDRE',TOBL.GetValeur(IOrdre));
      			TOBRG:=TOBPieceRG.findnext(['INDICERG'],[IndiceRg],true);
          end;
        until TOBRG = nil;
      end;
    end else if isOuvrage(TOBL) then
    begin
      SetInfoOuvrage (TOBFacture,TOBL,true);
    end;
  end;
  EcrireMaxNumOrdre(TOBFacture, MaxNumOrdre) ;
end;

function TGenFacAvanc.RecalculeLapiece: boolean;
begin
//  result := false;
  TOBBases.ClearDetail; TOBBasesL.ClearDetail; ZeroFacture (TOBFacture); ZeroAchat(TOBfacture);
  PutValueDetail(TOBFacture,'GP_RECALCULER','X');
  CalculeMontantsDoc (TOBfacture,TOBOuvrage,false);
  ZeroMontantPorts(TOBPorcs);
  TOBVTECOLL.ClearDetail;
  CalculFacture(nil, TOBFacture,TOBPieceTrait,TOBSousTrait,TOBOuvrage,TOBOuvragesP,TOBBases,TOBBasesL,TOBTiers,TOBArticles,TOBPorcs,TOBPieceRG,TOBBasesRG,TOBVTECOLL,DEV,false,tamodif,false,-1,True) ;
  {Sous totaux de la piece}
  CalculeSousTotauxPiece(TOBFacture) ;
  //
  if not RGMultiple(TOBFacture) then
  begin
  	RecalculeRG(TOBPORCS,TOBfACTURE, TOBPIECERG, TOBBASES, TOBBASESRG,TOBPieceTrait,DEV);
  end;
  //
  CalculeReglementsIntervenants(TOBSousTrait, TOBFacture,TOBPieceRG,TOBAcomptes,TOBPOrcs, TOBFacture.getValue('GP_AFFAIRE'),TOBPieceTrait,DEV,fLastfacture);
  {Echéances}
  TOBEches.ClearDetail;
  GereEcheancesGC(TOBFacture,TOBTiers,TOBEches,TOBAcomptes,TOBPieceRG,TOBPieceTrait,TOBPorcs,taCreat,DEV,False);
  result := true;
end;

procedure TGenFacAvanc.SetInfoOuvrageFille (TOBPiece,TOBOUV,TOBL : TOB; Ligne : integer; WithRenumOuv : boolean);
var Indice : integer;
    TOBO : TOB;
begin
  for Indice := 0 to TOBOuv.detail.count -1 do
  begin
    TOBO := TOBOUv.detail[Indice];
    //
    TOBO.putValue('BLO_NATUREPIECEG',newNature);
    TOBO.putValue('BLO_SOUCHE',NewSouche);
    TOBO.putValue('BLO_NUMERO',NewNumero);
    TOBO.putValue('BLO_INDICEG',0);
    TOBO.putValue('BLO_NUMLIGNE',Ligne);
    //
    if WithrenumOuv then
    begin
      TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
      TOBO.putValue('BLO_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));
    end;
    //
    if TOBO.detail.count> 0 then SetInfoOuvrageFille (TOBpiece,TOBO,TOBL,Ligne,WithRenumOuv);
  end;
end;

procedure TGenFacAvanc.SetInfoOuvrage(TOBPiece,TOBL: TOB;WithRenumOuv:boolean);
var IndiceNomen : integer;
    TOBOuv : TOB;
    Ligne : integer;
begin
  Ligne := TOBL.GetValue('GL_NUMLIGNE');
  IndiceNomen := TOBL.GetValue('GL_INDICENOMEN')-1;
  if IndiceNomen < 0 then exit;
  //FV1 : 28/07/2016
  if TOBOuvrage.detail.count = 0 then exit;
  if (Indicenomen) > (TOBOuvrage.detail.count - 1) then exit;
  //
  TOBOuv := TOBOuvrage.detail[IndiceNomen];
  TOBOuv.PutValue('UTILISE','X');
  SetInfoOuvrageFille (TOBPIECE,TOBOUV,TOBL,Ligne,WithRenumOuv);
end;

procedure TGenFacAvanc.PositionneInfoPieceSousDetail (TOBpiece,TOBL: TOB);

	procedure PositionneInfoPieceSousDetailDet (TOBPIECE,TOBOUV,TOBL : TOb; Ligne : integer; WithAvanc : boolean=true);
  var Indice : integer;
      TOBO : TOB;
  begin
    for Indice := 0 to TOBOuv.detail.count -1 do
    begin
      TOBO := TOBOUv.detail[Indice];
      //
      if not ModifSituation then
      begin
        if TOBO.Getvalue('BLO_PIECEORIGINE',) = '' then TOBO.putvalue('BLO_PIECEORIGINE',EncodeRefPieceOuv(TOBO));
        TOBO.putvalue('BLO_PIECEPRECEDENTE',EncodeRefPieceOuv(TOBO));
      end;
      //
      TOBO.putValue('BLO_NATUREPIECEG',TOBPiece.GetString('GP_NATUREPIECEG'));
      TOBO.putValue('BLO_SOUCHE',TOBPiece.GetString('GP_SOUCHE'));
      TOBO.putValue('BLO_NUMERO',TOBPiece.GetString('GP_NUMERO'));
      TOBO.putValue('BLO_INDICEG',TOBPiece.GetString('GP_INDICEG'));
      TOBO.putValue('BLO_NUMLIGNE',Ligne);
      //
      if TOBO.FieldExists('BLF_NATUREPIECEG') then
      begin
      	TOBO.putValue('BLF_NATUREPIECEG',TOBPiece.GetString('GP_NATUREPIECEG'));
      end;
//      if TOBO.detail.count> 0 then PositionneInfoPieceSousDetailDet (TOBpiece,TOBO,TOBL,Ligne,false);
    end;
  end;

var IndiceNomen : integer;
    TOBOuv : TOB;
    Ligne : integer;
begin
  Ligne := TOBL.GetValue('GL_NUMLIGNE');
  IndiceNomen := TOBL.GetValue('GL_INDICENOMEN')-1;
  if IndiceNomen < 0 then exit;
  //FV1 : 28/07/2016
  if TOBOuvrage.detail.count = 0 then exit;
  if (Indicenomen) > (TOBOuvrage.detail.count) then exit;
  //
  TOBOuv := TOBOuvrage.detail[IndiceNomen];
  TOBOuv.PutValue('UTILISE','X');
  PositionneInfoPieceSousDetailDet (TOBPIECE,TOBOUV,TOBL,Ligne);
end;


function TGenFacAvanc.GetLeTauxDevise(CodeD: string; var DateTaux: TDateTime; DateP: TDateTime): Double;
var Taux: Double;
  AvecT, ChoixT, Taux1: Boolean;
  ii: integer;
begin
  AvecT := True;
  ChoixT := False;
  Taux1 := False;
  if AvecT then
  begin
    Taux := GetTaux(CodeD, DateTaux, DateP);
    //
    // MODIF LM 02/04/2003
    // Saisie d'un document EUR sur un dossier en devise OUT
    // La table de chancellerie stocke le taux par rapport à l'EUR, par par rapport à la devise pivot du dossier.
    //
    //if ((CodeD<>V_PGI.DevisePivot) and (Not EstMonnaieIN(CodeD))) then
    //if ((CodeD <> 'EUR') and (not EstMonnaieIN(CodeD))) then
    if (CodeD <> GetParamSocSecur('SO_DEVISEPRINC','EUR')) then
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
  end else
  begin
    Taux := TOBfacture.GetValue('GP_TAUXDEV');
  end;
  Result := Taux;
end;

procedure TGenFacAvanc.InitToutModif ;
Var NowFutur : TDateTime ;
BEGIN
  NowFutur:=NowH ;
  TOBFacture.SetAllModifie(True) ; TOBFacture.SetDateModif(NowFutur) ;
  TOBBases.SetAllModifie(True)  ;
  TOBBasesL.SetAllModifie(True)  ;
  TOBEches.SetAllModifie(True)  ;
  TOBAcomptes.SetAllModifie(True)  ;
  TOBPorcs.SetAllModifie(True)  ;
  TOBTiers.SetAllModifie(True)  ;
  TOBAnaP.SetAllModifie(True)   ; TOBAnaS.SetAllModifie(True)   ;
  TOBPieceRG.SetAllModifie (true);
  TOBBasesRG.SetAllModifie (true);
  TOBVTECOLL.SetAllModifie(true);
END ;

function TGenFacAvanc.ValidelaPiece: boolean;
var IO : TIOErr;
begin
  IO := TRANSACTIONS (ValideThispiece,0);
  result := (IO=OeOk);
end;

Procedure TGenFacAvanc.ValideLesOuv ( TOBOuv : TOB ; TOBPiece : TOB = nil) ;
Var i,II : integer ;
    TOBL,TOBGO,TOBO : TOB;
BEGIN
  if TOBOUV = nil then exit;
  for i:=TOBOuv.Detail.Count-1 downto 0 do
  	if TOBOuv.Detail[i].GetValue('UTILISE')<>'X' then
  		TOBOuv.Detail[i].Free ;
  //
  if TOBPiece = nil then exit;

  for i := 0 to TOBOuv.detail.count -1 do
  begin
    TOBGO := TOBOUv.detail[I]; // groupe d'ouvrage;
    II := 0;     
    //FV1 : 27/07/2016 : Génération facturation avec des trucs qui manquent...
    if TOBGO.Detail.count <= 0 then continue;
    //
    repeat
      if TOBGO.detail[II].getValue('BLO_ARTICLE')='' then TOBGO.detail[II].Free else inc(II);
    until II >= TOBGO.detail.count;
    //
    TOBO := TOBGO.detail[0];
    TOBL := GetTOBLIgne (TOBPiece, TOBO.getValue('BLO_NUMLIGNE'));
    if TOBL <> nil then TOBL.putValue('GL_INDICENOMEN',I+1);
    //
  end;

  if Not TOBOuv.InsertDBByNivel(false) then
  begin
    V_PGI.IoError:=oeUnknown ;
    MessageValid := 'Erreur écriture des ouvrages';
  end;
END ;
//
procedure TGenFacAvanc.ValideThispiece;
begin
  V_PGI.IOError := OeOk;
  if ModifSituation then
  begin
    // On restitue l'avancement précédent avant la mise en place du nouvel avancement
    RestitueAvancementDevis;
    if V_PGI.IoError=oeOk then SetPieceMorte(TOBDevis)
  end;
  TOBPieceTraitDevis.ClearDetail;

  if TOBSituations.Detail.count > 0 then
  begin
    if not TOBSituations.InsertDB(nil) then V_PGI.IOError := oeUnknown;
  end;

  if V_PGI.IoError=oeOk then ValideLesArticlesFromOuv(TOBarticles,TOBOuvrage);
  if V_PGI.IoError=oeOk then ValideLesLignes(TOBfacture,TOBArticles,nil,nil,TOBOuvrage,TOBPieceRG,TOBBasesRG,false,True,false,true,nil,false) ;
  if V_PGI.IoError=oeOk then ValideLesLignesCompl(TOBfacture, nil);
  if V_PGI.IoError=oeOk then ValideLesLignesFac;
  if V_PGI.IoError=oeOk then ValideLesPieceTrait(TOBFacture, nil,TOBPieceTrait,TOBSousTrait,DEV);
  if V_PGI.IoError=oeOk then ValideLesSousTrait (TOBFacture,TOBSousTrait,DEV);
  if V_PGI.IoError=oeOk then ValideLesPieceTraitOrig;
  if V_PGI.IoError=oeOk then UpdateLesLignesDevis; // Mise à jour des quantités situations et cumulés (ancienne gestion)
  if V_PGI.IoError=oeOk then ValideLesAdresses(TOBfacture,nil,TOBAdresses) ;
  if V_PGI.IoError=oeOk then ValideLesArticles(TOBfacture,TOBArticles) ;
  if V_PGI.IoError=oeOk then ValideleTiers(TOBfacture,TOBTiers) ;
//  if V_PGI.IoError=oeOk then UG_ValideAnals(TOBfacture,TOBAnaP,TOBAnaS) ;
  if V_PGI.IoError=oeOk then GenereCompta ;
  if (V_PGI.IoError=oeOk) and (TOBVTECOLL.detail.count > 0) then
  begin
    PrepareInsertCollectif (TOBFacture,TOBVTECOLL);
    TOBVTECOLL.InsertDB(nil);
  end;
  if V_PGI.IoError=oeOk then ValideLesAcomptes(TOBfacture,TOBAcomptes,TOBAcomptes_O) ;
  if V_PGI.IoError=oeOk then ValideLesPorcs(TOBfacture,TOBPorcs) ;
  if V_PGI.IoError=oeOk then ValideLesOuv(TOBOuvrage,TOBfacture) ;
  if V_PGI.IoError=oeOk then ValideLesOuvPlat (TOBOuvragesP, TOBfacture);
  if V_PGI.IoError=oeOk then ValideLesBases(TOBfacture,TobBases,TOBBasesL);
  if V_PGI.IoError=oeOk then ValideLesTimbres(TOBfacture, TOBTimbres);
  //
  if V_PGI.IoError=oeOk then TOBfacture.InsertDBByNivel(False) ;
  if V_PGI.IoError=oeOk then TOBOuvragesP.InsertDBByNivel(false);
  if V_PGI.IoError=oeOk then TOBBasesL.InsertDB(nil);
  //--
  if V_PGI.IoError=oeOk then TOBBases.InsertDB(Nil) ;
  if V_PGI.IoError=oeOk then TOBEches.InsertDB(Nil) ;
  if V_PGI.IoError=oeOk then TOBAnaP.InsertDB(Nil) ;
  if V_PGI.IoError=oeOk then TOBAnaS.InsertDB(Nil) ;
  // Modif BTP
  if V_PGI.IoError=oeOk then ValideLesRetenues (TOBfacture,TOBPIECERG);
  if V_PGI.IoError=oeOk then ReajusteCautionDocOrigine (TOBfacture,TOBPIECERG,true);
  if V_PGI.IoError=oeOk then ValideLesBasesRG(TOBfacture,TOBBasesRG);
  // ----
  if (V_PGI.IoError=oeOk) then MajAffaireApresGeneration (TOBfacture,TOBPieceRG,TOBBasesRG,TOBAcomptes,DEV,false,NumSituation);
  if (V_PGI.IoError=oeOk) then TheRepartTva.Ecrit;
  if (V_PGI.IoError=oeOk) then MajPiecePrecedente;
  // --
  {Compte rendu}
end;

Procedure TGenFacAvanc.GenereCompta ;
Var OldEcr,OldStk : RMVT ;
BEGIN
  FillChar(OldEcr,Sizeof(OldEcr),#0) ; FillChar(OldStk,Sizeof(OldStk),#0) ;
  if Not PassationComptable(TOBFacture,TOBOuvrage ,TOBOuvragesP,TOBBases,TOBBasesL,TOBEches,TOBPieceTrait,TOBAffaireInterv,TOBTiers,TOBArticles,TOBCpta,TOBAcomptes,TOBPorcs,TOBPIECERG,TOBBasesRG,TOBAnaP,TOBAnaS,TOBSousTrait,TOBVTECOLL,DEV,OldEcr,OldStk,false)
     then V_PGI.IoError:=oeLettrage ;
END ;

Procedure TGenFacAvanc.AjouteEvent ;
Var St : String ;
    TOBE : TOB;
    OkOk : boolean;
    NumEvent,NbRetry : integer;
BEGIN
  TOBE := TOB.Create ('JNALEVENT',nil,-1);
  TRY
    St:='Pièce N° '+IntToStr(TOBfacture.GetValue('GP_NUMERO'))
        +', Tiers '+TOBfacture.GetValue('GP_TIERS')
        +', Total HT de '+StrfPoint(TOBfacture.GetValue('GP_TOTALHTDEV'))+' '+RechDom('TTDEVISETOUTES',TOBfacture.GetValue('GP_DEVISE'),False) ;
    PiecesG.Add(St) ;
    TOBE.putValue('GEV_TYPEEVENT','GEN');
    TOBE.putValue('GEV_LIBELLE',Copy('Génération de '+RechDom('GCNATUREPIECEG',NewNature,False),1,35));
    TOBE.putValue('GEV_DATEEVENT',Date);
    TOBE.putValue('GEV_UTILISATEUR',V_PGI.User);
    TOBE.putValue('GEV_ETATEVENT','OK');
    NumEvent := GetNumEvent;
    NbRetry := 0;
    repeat
      Inc(NumEvent);
      TOBE.putValue('GEV_NUMEVENT',NumEvent);
      TOBE.SetAllModifie(true);
      OkOk := TOBE.InsertDB (nil); if not OkOk then Inc(NbRetry);
    until (OkOk) or (Nbretry > 10);
  FINALLY
    TOBE.Free;
  END;
END ;

function TGenFacAvanc.GetNumEvent: integer;
var QQ : Tquery;
begin
  result := -1;
  QQ := OpenSQL ('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT',True) ;
  if not QQ.eof then result := QQ.Fields[0].AsInteger;
  ferme(QQ);
end;

procedure TGenFacAvanc.UpdateLesLignesDevis;

  procedure MajLigneDevis (TOBL,TOBDevis : TOB);
  var CD : r_cledoc;
      StPrec,Req : string;
      NewValeur : double;
      Q : Tquery;
      TOBD : TOB;
  begin
    StPrec:=TOBL.GetValue('GL_PIECEPRECEDENTE') ;
    DecodeRefPiece(StPrec,CD) ;
    Req:='SELECT * FROM LIGNE WHERE '+ WherePiece (CD,ttdLigne,true,True);
    Q := OpenSql (Req,false);
    if not Q.eof then
    begin
      TOBD := TOB.Create ('LIGNE',TOBDEVIS,-1);
      TOBD.SelectDB ('',Q);
      //
      if Avancement then
      begin
        NewValeur := TOBL.GetValue('BLF_QTECUMULEFACT');
      end else
      begin
        NewValeur := TOBL.GetValue('BLF_QTEDEJAFACT')+TOBL.GetValue('BLF_QTESITUATION');
      end;
      //
      TOBD.PutValue('GL_QTESIT',NewValeur);
      TOBD.PutValue('GL_QTEPREVAVANC',NewValeur);
      if Avancement then
      begin
        TOBD.PutValue('GL_POURCENTAVANC',TOBL.GetValue('BLF_POURCENTAVANC'));
      end;
    end;
    ferme (Q);
  end;

var TOBL : TOB;
    Indice : integer;
    TOBLDEVIS : TOB;
begin
  TOBLDEVIS := TOB.create('LES LIGNES FACTURES DEVIS',nil,-1);
  TRY
    for Indice := 0 to TOBFACTURE.detail.count -1 do
    begin
      TOBL := TOBFACTURE.detail[Indice];
      if TOBL.GetValue('GL_TYPELIGNE')<>'ART' then continue;
      if TOBL.GetValue('GL_PIECEPRECEDENTE')= '' then continue;
      if (TOBL.GetValue('GL_NUMORDRE')<>0) and ((TOBL.GetValue('BLF_MTSITUATION')<>0) or (ModifSituation)) then
      begin
        MajLigneDevis (TOBL,TOBLdevis);
      end;
    end;
    //
    if TOBLDEVIS.detail.count <> 0 then
    begin
      if not TOBLDEVIS.UpdateDB (false) then
      BEGIN
        CodeErreur := 'Erreur durant l''écriture de l''avancement dans le devis';
        V_PGI.IOError := OeUnknown;
      END;
    end;
  FINALLY
    TOBLDEVIS.Free;
  end;
end;

function TGenFacAvanc.FindPieceTraitOrig (Cledoc : r_cledoc;Fournisseur : string) : TOB;
begin
  Result := TOBPieceTraitDevis.findFirst(['BPE_NATUREPIECEG','BPE_SOUCHE','BPE_NUMERO','BPE_INDICEG','BPE_FOURNISSEUR'],
                                         [Cledoc.NaturePiece,Cledoc.Souche,cledoc.NumeroPiece,Cledoc.Indice,Fournisseur],True);
end;

function TGenFacAvanc.ChargePieceTraitOrig(cledoc : R_CLEDOC ; Fournisseur : string): TOB;
var Sql : string;
    QQ : TQuery;
begin
  Result := nil;
  Sql := 'SELECT *,BPI_TYPEPAIE AS TYPEPAIE FROM PIECETRAIT '+
         'LEFT JOIN PIECEINTERV ON BPI_NATUREPIECEG=BPE_NATUREPIECEG AND '+
         'BPI_SOUCHE=BPE_SOUCHE AND BPI_NUMERO=BPE_NUMERO AND BPI_INDICEG=BPE_INDICEG AND '+
         'BPI_TIERSFOU=BPE_FOURNISSEUR '+
  			 'WHERE '+WherePiece(cledoc,ttdPieceTrait,True)+' AND BPE_FOURNISSEUR="'+Fournisseur+'"';
  QQ := OpenSQL(Sql,True,1,'',True);
  if not QQ.eof then
  begin
    result := TOB.Create ('PIECETRAIT',TOBPieceTraitDevis,-1);
    result.SelectDB('',QQ);
  end;
  Ferme(QQ);
end;


procedure TGenFacAvanc.SetLignesDevisOuvDetail (TOBL,TOBOuvrage,TOBLDEVIS : TOB; avancement : boolean);
var indiceNomen,II : integer;
		TOBOUV,TOBOO,TOBD,TOBPTraitDevisL,TOBPLAT,TOBPP : TOB;
    cledoc : r_cledoc;
    NatureTravailOuv,NatureTravailDet : Double;
begin
	IndiceNomen      := TOBL.GetValue('GL_INDICENOMEN');
  NatureTravailOuv := VALEUR(TOBL.GetValue('GLC_NATURETRAVAIL'));
  //FV1 : 28/07/2016
  if IndiceNomen = 0 then exit;
  if TOBOuvrage.detail.count = 0 then exit;
  if (Indicenomen) > (TOBOuvrage.detail.count) then exit;
  //
  TOBOUV := TOBOuvrage.detail[IndiceNomen-1];
  if TOBOUV = nil then exit;
  for II := 0 to TOBOUV.detail.count -1 do
  begin
  	TOBOO := TOBOUV.detail[II];
  	if TOBOO.GEtValue('BLO_PIECEPRECEDENTE') = '' then continue;
    //
    if ((avancement) and (TOBOO.GEtValue('BLF_MTCUMULEFACT')=0)) OR
    	 ((not Avancement) and (TOBOO.GEtValue('BLF_MTSITUATION')=0)) then continue;
    //
    DecodeRefPieceOUv (TOBOO.GEtValue('BLO_PIECEPRECEDENTE'),cledoc);
    //
  	TOBD := TOB.Create ('LIGNEFAC',TOBLDEVIS,-1);
    TOBD.putValue('BLF_NATUREPIECEG',Cledoc.NaturePiece);
    TOBD.putValue('BLF_SOUCHE',Cledoc.souche);
    TOBD.putValue('BLF_DATEPIECE',Cledoc.DatePiece );
    TOBD.putValue('BLF_NUMERO',Cledoc.NumeroPiece);
    TOBD.putValue('BLF_INDICEG',Cledoc.Indice);
    TOBD.putValue('BLF_NUMORDRE',0);
    TOBD.putValue('BLF_UNIQUEBLO',Cledoc.UniqueBlo);
    TOBD.putValue('BLF_MTMARCHE',TOBOO.GetValue('BLF_MTMARCHE'));
    //
    TOBD.AddChampSupValeur('PIECEPRECEDENTE', TOBOO.GEtValue('BLO_PIECEPRECEDENTE'));
    //
    if Avancement then
    begin
      TOBD.putValue('BLF_MTDEJAFACT',TOBOO.GetValue('BLF_MTCUMULEFACT'));
      TOBD.putValue('BLF_MTCUMULEFACT',TOBOO.GetValue('BLF_MTCUMULEFACT'));
    end else
    begin
      TOBD.putValue('BLF_MTDEJAFACT',TOBOO.GetValue('BLF_MTDEJAFACT')+TOBOO.GetValue('BLF_MTSITUATION'));
    end;
    TOBD.putValue('BLF_MTSITUATION',0);
    TOBD.putValue('BLF_QTEMARCHE',TOBOO.GetValue('BLF_QTEMARCHE'));
    if Avancement then
    begin
      TOBD.putValue('BLF_QTECUMULEFACT',TOBOO.GetValue('BLF_QTECUMULEFACT'));
      TOBD.putValue('BLF_QTEDEJAFACT',TOBOO.GetValue('BLF_QTECUMULEFACT'));
    end else
    begin
      TOBD.putValue('BLF_QTEDEJAFACT',TOBOO.GetValue('BLF_QTEDEJAFACT')+TOBOO.GetValue('BLF_QTESITUATION'));
    end;
    TOBD.putValue('BLF_QTESITUATION',0);
    if Avancement then
    begin
      TOBD.putValue('BLF_POURCENTAVANC',TOBOO.GetValue('BLF_POURCENTAVANC'));
    end else
    begin
      TOBD.putValue('BLF_POURCENTAVANC',0);
    end;
    TOBD.putValue('BLF_NATURETRAVAIL',TOBOO.GetValue('BLF_NATURETRAVAIL'));
    TOBD.putValue('BLF_FOURNISSEUR',TOBOO.GetValue('BLF_FOURNISSEUR'));
    TOBD.SetAllModifie (true);
  end;
  //
  if (NatureTravailOuv>=10) then
  begin
    DecodeRefPiece (TOBL.GEtValue('GL_PIECEPRECEDENTE'),cledoc);
    TOBPlat := FindOuvragesPlat (TOBL,TOBOuvragesP);
    if TOBPlat = nil then exit;
    for II := 0 to TOBPlat.detail.count -1 do
    begin
      TOBPP := TOBplat.detail[II];
      NatureTravailDet := VALEUR(TOBPP.GetValue('BOP_NATURETRAVAIL'));
      if (natureTravailDet<10)  then
      begin
        TOBPTraitDevisL := FindPieceTraitOrig (cledoc,TOBPP.GetValue('BOP_FOURNISSEUR'));
        if TOBPTraitDevisL = nil then
        begin
          TOBPTraitDevisL := ChargePieceTraitOrig(cledoc,TOBPP.GetValue('BOP_FOURNISSEUR'));
        end;
        if TOBPTraitDevisL <> nil then
        begin
          TOBPTraitDevisL.PutValue('BPE_MONTANTFAC',TOBPTraitDevisL.GetValue('BPE_MONTANTFAC')+TOBPP.GetValue('BOP_TOTALTTCDEV'));
        end;
      end;
    end;
  end;
end;


procedure TGenFacAvanc.ValideLesLignesFac;

  procedure  SetLigneDevis (TOBL,TOBD : TOB);
  var cledoc : r_cledoc;
  		TOBPTraitDevisL : TOB;
  begin
  	if TOBL.GEtValue('GL_PIECEPRECEDENTE') = '' then exit;
    DecodeRefPiece (TOBL.GEtValue('GL_PIECEPRECEDENTE'),cledoc);

    TOBD.putValue('BLF_NATUREPIECEG',Cledoc.NaturePiece);
    TOBD.putValue('BLF_SOUCHE',Cledoc.souche);
    TOBD.putValue('BLF_DATEPIECE',Cledoc.DatePiece );
    TOBD.putValue('BLF_NUMERO',Cledoc.NumeroPiece);
    TOBD.putValue('BLF_INDICEG',Cledoc.Indice);
    TOBD.putValue('BLF_NUMORDRE',cledoc.numordre);
    TOBD.putValue('BLF_MTMARCHE',TOBL.GetValue('BLF_MTMARCHE'));
    //
    TOBD.AddChampSupValeur('PIECEPRECEDENTE', TOBL.GEtValue('BLO_PIECEPRECEDENTE'));
    //
    if Avancement then
    begin
      TOBD.putValue('BLF_MTDEJAFACT',TOBL.GetValue('BLF_MTCUMULEFACT'));
      TOBD.putValue('BLF_MTCUMULEFACT',TOBL.GetValue('BLF_MTCUMULEFACT'));
    end else
    begin
      TOBD.putValue('BLF_MTDEJAFACT',TOBL.GetValue('BLF_MTDEJAFACT')+TOBL.GetValue('BLF_MTSITUATION'));
    end;
    TOBD.putValue('BLF_MTSITUATION',0);
    TOBD.putValue('BLF_QTEMARCHE',TOBL.GetValue('BLF_QTEMARCHE'));
    if Avancement then
    begin
      TOBD.putValue('BLF_QTECUMULEFACT',TOBL.GetValue('BLF_QTECUMULEFACT'));
      TOBD.putValue('BLF_QTEDEJAFACT',TOBL.GetValue('BLF_QTECUMULEFACT'));
    end else
    begin
      TOBD.putValue('BLF_QTEDEJAFACT',TOBL.GetValue('BLF_QTEDEJAFACT')+TOBL.GetValue('BLF_QTESITUATION'));
    end;
    TOBD.putValue('BLF_QTESITUATION',0);
    if Avancement then
    begin
      TOBD.putValue('BLF_POURCENTAVANC',TOBL.GetValue('BLF_POURCENTAVANC'));
    end else
    begin
      TOBD.putValue('BLF_POURCENTAVANC',0);
    end;
    TOBD.putValue('BLF_NATURETRAVAIL',TOBL.GetValue('GLC_NATURETRAVAIL'));
    TOBD.putValue('BLF_FOURNISSEUR',TOBL.GetValue('GL_FOURNISSEUR'));
    TOBD.putValue('BLF_TOTALTTCDEV',TOBL.GetValue('BLF_TOTALTTCDEV'));

    TOBD.SetAllModifie (true);
    if (VALEUR(TOBL.GetValue('GLC_NATURETRAVAIL'))<10) then
    begin
      TOBPTraitDevisL := FindPieceTraitOrig (cledoc,TOBL.GetValue('GL_FOURNISSEUR'));
      if TOBPTraitDevisL = nil then
      begin
        TOBPTraitDevisL := ChargePieceTraitOrig(cledoc,TOBL.GetValue('GL_FOURNISSEUR'));
      end;
      if TOBPTraitDevisL <> nil then
      begin
      	TOBPTraitDevisL.PutValue('BPE_MONTANTFAC',TOBPTraitDevisL.GetValue('BPE_MONTANTFAC')+TOBL.GetValue('BLF_TOTALTTCDEV'));
      end;
    end;
  end;

var TOBL : TOB;
    Indice : integer;
    TOBLFactures,TOBLDEVIS,TOBD : TOB;
    Repartition : Boolean;
begin
  // plus valable depuis la facturation sur les sous détails
  // if not avancement then exit;
  TOBLFactures := TOB.create('LES LIGNES FACTURES',nil,-1);
  TOBLDEVIS := TOB.create('LES LIGNES FACTURES DEVIS',nil,-1);
  //
  Repartition := GetParamSocSecur('SO_PIECEREPART', False);
  //
  TRY
    for Indice := 0 to TOBFACTURE.detail.count -1 do
    begin
      TOBL := TOBFACTURE.detail[Indice];
      if TOBL.GetValue('GL_TYPELIGNE')<>'ART' then continue;
      if (TOBL.GetValue('GL_NUMORDRE')<>0) then
      begin
        TOBD := TOB.Create ('LIGNEFAC',TOBLFactures,-1);
        //
        TOBD.AddChampSupValeur('PIECEPRECEDENTE', TOBL.GetValue('GL_PIECEPRECEDENTE'));
        //
        SetLigneFacture (TOBL,TOBD);
        // MODIF
        if isOuvrage(TOBL) then
        begin
          //Pour un ouvrage il faut avoir dans le LigneFac une zone qui permet de l'identifier
          //pour pouvoir gérer la mise à jour des avenant
          if Repartition then TOBD.AddChampSupValeur('ISOUVRAGE', 'X');
          SetLignesFactureOuvDetail (TOBL,TOBOuvrage,TOBLFactures);
        end;
        //
        TOBD := TOB.Create ('LIGNEFAC',TOBLDEVIS,-1);
        SetLigneDevis (TOBL,TOBD);
        if isOuvrage(TOBL) then
        begin
          if Repartition then TOBD.AddChampSupValeur('ISOUVRAGE', 'X');
          SetLignesDevisOuvDetail (TOBL,TOBOuvrage,TOBLDEVIS,avancement);
        end;
      end;
    end;
    //
    if TOBLDEVIS.detail.count <> 0 then
    begin
      if not TOBLDEVIS.InsertOrUpdateDB (false) then
      BEGIN
        CodeErreur := 'Erreur durant l''écriture de l''avancement sur devis';
        V_PGI.IOError := OeUnknown;
      END;
    end;
    if TOBLFactures <> nil then
    begin
      TOBLFactures.Detail.Sort('BLF_NATUREPIECEG;BLF_SOUCHE;BLF_NUMERO;BLF_INDICEG;BLF_NUMORDRE;BLF_UNIQUEBLO');
      //FV1 - 27/07/2016
      if not TOBLFactures.InsertOrUpdateDB (false) then //InsertDB  (nil,false) then
      begin
      CodeErreur := 'Erreur durant l''écriture de l''avancement sur la facture';
      V_PGI.IOError := OeUnknown;
    end;
    end;

    //
  FINALLY
    //FV1 ! 01/12/2015 - FS#355 - DELABOUDINIERE : En impression situation, ajouter les avenants
    if Repartition then ChargePieceRepart(TOBFACTURE.GetString('GP_AFFAIREDEVIS'), TOBLFactures);
    TOBLFactures.Free;
    TOBLDEVIS.Free;
  end;
end;


procedure TGenFacAvanc.InitChampsSupFac(TOBL: TOB);
var indice,itable : integer;
    typeC,NomC : string;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  table := Mcd.getTable(mcd.PrefixetoTable('BLF'));
  FieldList := Table.Fields;
  FieldList.Reset();
  While FieldList.MoveNext do
  begin
    NomC := (FieldList.Current as IFieldCOM).name;
    TypeC := (FieldList.Current as IFieldCOM).Tipe;
    if (typeC = 'INTEGER') or (typeC = 'SMALLINT') then TOBL.putValue(NomC,0)
    else if (typeC = 'DOUBLE') or (typeC = 'EXTENDED') or (typeC = 'DOUBLE') then TOBL.putValue(NomC,0)
    else if (typeC = 'DATE') then TOBL.putValue(NomC,iDate1900)
    else if (typeC = 'BLOB') or (typeC = 'DATA') then TOBL.putValue(NomC,'')
    else if (typeC = 'BOOLEAN') then TOBL.putValue(NomC,'-')
    else TOBL.putValue(NomC,'')
  end;
end;

procedure TGenFacAvanc.AddChampsSupFac(TOBL: TOB);
var indice,itable : integer;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  table := Mcd.getTable(mcd.PrefixetoTable('BLF'));
  FieldList := Table.Fields;
  FieldList.Reset();
  While FieldList.MoveNext do
  begin
    if not TOBL.FieldExists ((FieldList.Current as IFieldCOM).name) then
    begin
      TOBL.AddChampSup ((FieldList.Current as IFieldCOM).name,false);
    end;
  end;
end;

procedure TGenFacAvanc.DupliqueLaLigne (NewTOBL,OldTOBL : TOB) ;
var Indice : integer;
    NomChamp : string;
begin
  for indice := 1 to NewTOBL.NbChamps do
  begin
    NomChamp := NewTOBL.GetNomChamp(indice);
    if OldTOBL.fieldExists (NomChamp) then
    begin
      NewTOBL.putValue(NomChamp,OldTOBL.GetValue(NomChamp));
    end;
  end;
  // Champs Suplémentaire
  for indice := 1000 to 1000 + NewTOBL.ChampsSup.Count - 1 do
  begin
    NomChamp := NewTOBL.GetNomChamp(indice);
    if OldTOBL.fieldExists (NomChamp) then
    begin
      NewTOBL.putValue(NomChamp,OldTOBL.GetValue(NomChamp));
    end;
  end;
  //
end;

procedure TGenFacAvanc.GetAnalytiquePiece(TOBP: TOB);
var RefA : string;
    Q : TQuery;
begin
  RefA:=EncodeRefCPGescom(TOBP) ; if RefA='' then Exit ;
  Q:=OpenSQL('SELECT * FROM VENTANA WHERE YVA_TABLEANA="GL" AND YVA_IDENTIFIANT="'+RefA+'" ORDER BY YVA_AXE, YVA_NUMVENTIL',True) ;
  if Not Q.EOF then TOBAnalPiece.LoadDetailDB('VENTANA','','',Q,true) ;
  Ferme(Q) ;
end;

procedure TGenFacAvanc.GetArticlePieces(TOBPDoc: TOB);
var stSelect : string;
    TOBART,TOBAA : TOB;
    QArticle : TQuery;
    stWhereLigne,StWhereLigneOuv  : string;
    Cledoc : R_cledoc;
    indice : integer;
begin
  TOBART := TOB.Create ('LES ARTICLES',nil,-1);
  Cledoc := TOB2CleDoc (TOBPDoc);
  //
  StSelect := 'SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
            'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
            'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE';
  stWhereLigne := WherePiece (cledoc,ttdligne,false);
  QArticle := OpenSQL(StSelect + ' WHERE GA_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne + ')', True);
  if not QArticle.EOF then TOBART.LoadDetailDB('ARTICLE', '', '', QArticle, True, True);
  Ferme(QArticle); // chargement des articles de la piecer (ligne)
  //
  stWhereLigneOuv := WherePiece (cledoc,ttdOuvrage,true);
  QArticle := OpenSQL(StSelect + ' WHERE GA_ARTICLE IN (SELECT DISTINCT BLO_ARTICLE FROM LIGNEOUV WHERE (' + stWhereLigneOuv + ') AND (BLO_ARTICLE <>""))', True);
  if not QArticle.EOF then TOBArt.LoadDetailDB('ARTICLE', '', '', QArticle, True, True);
  Ferme(QArticle); // chargement des articles de la piece (ouvrage)
  //
  indice := 0;
  repeat
    TOBAA := TOBArticles.findFirst(['GA_ARTICLE'],[TOBART.detail[Indice].GetValue('GA_ARTICLE')],true);
    if TOBAA <> nil then inc(indice) else TOBART.detail[Indice].ChangeParent (TOBArticles,-1);
  until indice >= TOBArt.detail.count;
  TOBART.free;
end;

procedure TGenFacAvanc.LoadLesGCS;
var Q : TQuery;
begin
  Q := OpenSQL('SELECT * FROM COMMUN WHERE CO_TYPE="GCS" AND CO_LIBRE LIKE "%'+VH_GC.GCMarcheVentilAna+'%"',True);
  if not Q.eof Then TOBTablette.LoadDetailDB ('COMMUN','','',Q,false,true);
  ferme (Q);
end;


procedure TGenFacAvanc.InitLigneFacturation (TOBL : TOB);
begin
  TOBL.PutValue('GL_QTEFACT', 0);
  TOBL.PutValue('GL_QTERESTE',0);
  TOBL.PutValue('GL_QTESTOCK',0);
  TOBL.PutValue('GL_MTRESTE', 0);
end;

procedure TGenFacAvanc.ZeroAchat(TOBfacture : TOB);
begin
  TOBfacture.putValue('GP_MONTANTPA',0);
  TOBfacture.putValue('GP_MONTANTPR',0);
  TOBfacture.putValue('GP_MONTANTPAFC',0);
  TOBfacture.putValue('GP_MONTANTPAFG',0);
  TOBfacture.putValue('GP_MONTANTPAFR',0);
  TOBfacture.putValue('GP_MONTANTFR',0);
  TOBfacture.putValue('GP_MONTANTFC',0);
  TOBfacture.putValue('GP_MONTANTFG',0);
end;

procedure TGenFacAvanc.GetlesPortsFromDevisprinc;
var SQL,refPiece : string;
		QQ : TQuery;
    cledoc : r_cledoc;
begin
  RefPiece := EncodeRefPiece (TOBFacture,0,false);
  DecodeRefPiece (Refpiece,cledoc);
  SQl := 'SELECT * FROM PIEDPORT WHERE '+WherePiece(Cledoc,ttdPorc ,true);
  QQ := OpenSql (SQl,true,-1,'',true);
  if not QQ.eof then TOBPorcs.LoadDetailDB('PIEDPORT','','',QQ,false);
  ferme (QQ);
end;

procedure TGenFacAvanc.MajPiecePrecedente;
var Indice : integer;
		TOBP : TOB;
    Req,RefFacture : string;
begin
	RefFacture := EncodeRefPiece (TOBFacture,0,false);
  for Indice := 0 to TOBOptDevis.detail.count -1 do
  begin
    TOBP := TOBOptDevis.detail[Indice];
    Req := 'UPDATE PIECE SET GP_DEVENIRPIECE="'+RefFacture+'" '+
    			 'WHERE GP_NATUREPIECEG="'+TOBP.getValue('GP_NATUREPIECEG')+'" AND '+
           'GP_SOUCHE="'+TOBP.getValue('GP_SOUCHE')+'" AND '+
           'GP_NUMERO='+IntToStr(TOBP.getValue('GP_NUMERO'))+' AND '+
           'GP_INDICEG='+IntToStr(TOBP.getValue('GP_INDICEG'));
    ExecuteSQL(REQ);
  end;
end;


procedure TGenFacAvanc.ValideLesPieceTraitOrig;
begin
	if TOBPieceTraitDevis.detail.count > 0 then
  begin
  	if not TOBPieceTraitDevis.UpdateDB  then
  	begin
    	MessageValid := 'Erreur mise à jour des PIECETRAIT ORIGINE';
	    V_PGI.IoError := oeUnknown;
  	end;
  end;
end;



function TGenFacAvanc.AjouteLigneEcart(MtEcart: double) : boolean;
var TOBL,TOBArt : TOB;
    Artecart : string;
    Qart : TQuery;
		TOBAffaire : TOB;
    QQ : Tquery;
begin
	result := true;
	TOBart := TOB.create ('ARTICLE',nil,-1);
  TOBAffaire := TOB.Create ('AFFAIRE',nil,-1);
  TRY
    if TOBFacture.GetString('GP_AFFAIRE') <> '' then
    begin
      QQ := OpenSql ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+TOBFacture.GetString('GP_AFFAIRE')+'"',true,-1, '', True);
      if not QQ.eof then TOBAffaire.selectdb ('',QQ);
      ferme (QQ);
    end;
    //
    ArtEcart := GetParamsoc('SO_ECARTFACTURATION');
    if ArtEcart = '' then
    begin
      PgiError (Traduirememoire('Impossible : l''article d''écart n''a pas été paramétré'));
      result := false;
    end;
    Qart := opensql('Select GA_ARTICLE from ARTICLE Where GA_CODEARTICLE="' + ArtEcart + '"', true,-1, '', True);
    if not Qart.eof then
    begin
      ArtEcart := Qart.fields[0].AsString;
      ferme (Qart);
    end else
    begin
      ferme (Qart);
      PgiError (Traduirememoire('Impossible : l''article d''écart n''a pas été paramétré'));
      result := false;
    end;
    QArt := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                   'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                   'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+ArtEcart+'"',true,-1, '', True);
    if not QArt.EOF then
    begin
      TOBArt.SelectDB('', QArt);
      InitChampsSupArticle (TOBART);
      ferme (Qart);
    end else
    begin
      ferme (Qart);
      PgiError (Traduirememoire('Impossible : l''article d''écart n''a pas été paramétré'));
      result := false;
    end;
    //
    TOBL := NewTOBLigne (TOBfacture,0,true);
    InitChampsSupFac (TOBL);
    InitialiseTOBLigne(TOBFacture,TOBTiers,TOBAffaire,TOBL.GetIndex +1,false);

    TOBL.putValue('GL_CREERPAR','FAC'); // Définition pour écart
//    EnHt := (TOBL.getValue('GL_FACTUREHT')='X');
    // -- Pour ne pas recalculer a partir des PA * coef ..etc..
    TOBL.putValue('GLC_NONAPPLICFRAIS','X');
    TOBL.putValue('GLC_NONAPPLICFG','X');
    TOBL.putValue('GLC_NONAPPLICFC','X');
    //
    TOBArt.PutValue('REFARTSAISIE', copy(TOBArt.GetValue('GA_ARTICLE'), 1, 18));
    TOBL.PutValue('GL_ARTICLE', TOBArt.GetValue('GA_ARTICLE'));
    TOBL.PutValue('GL_REFARTSAISIE', TOBArt.GetValue('REFARTSAISIE'));
    TOBL.PutValue('GL_CODEARTICLE', TOBArt.GetValue('GA_CODEARTICLE'));
    TOBL.PutValue('GL_REFARTBARRE', TOBArt.GetValue('REFARTBARRE'));
    TOBL.PutValue('GL_REFARTTIERS', TOBArt.GetValue('REFARTTIERS'));
    TOBL.PutValue('GL_TYPEREF', 'ART');
    ArticleVersLigne(TOBfacture, TOBArt, nil, TOBL, TOBTiers);
    TOBL.PutValue('GL_QTEFACT', 1);
    TOBL.PutValue('GL_QTESTOCK', 1);
    TOBL.PutValue('GL_QTERESTE', 1);
    TOBL.putValue('GL_MONTANTHTDEV',MtEcart);
    TOBL.putValue('GL_TOTALHTDEV',MtEcart);
    TOBL.PutValue('GL_MTRESTE', 0);
    //
    TOBL.PutValue('GL_PRIXPOURQTE', 1);
    TOBL.PutValue('GL_PUHTDEV', MtEcart);
    TOBL.putValue('GL_DPA',0);
    TOBL.putValue('GL_COEFFG',0);
    TOBL.putValue('GL_COEFFC',0);
    TOBL.putValue('GL_DPR',0);
    TOBL.putValue('GL_COEFMARG',0);
    TOBL.PutValue('GL_REMISABLELIGNE', '-');
    TOBL.PutValue('GL_REMISABLEPIED', '-');
    TOBL.putValue('BLF_MTMARCHE',0);
    TOBL.putValue('BLF_MTDEJAFACT',0);
    TOBL.putValue('BLF_MTCUMULEFACT',0);
    TOBL.putValue('BLF_MTSITUATION',MtEcart);
    TOBL.putValue('BLF_QTEMARCHE',0);
    TOBL.putValue('BLF_QTEDEJAFACT',0);
    TOBL.putValue('BLF_QTECUMULEFACT',0);
    TOBL.putValue('BLF_QTESITUATION',1);
    TOBL.putValue('BLF_POURCENTAVANC',0);
    TOBL.PutValue('GL_RECALCULER', 'X');
  FINALLY
  	TOBARt.free;
    TOBAffaire.Free;
  END;
end;

function TGenFacAvanc.PositionneEcartLastfacture ( var EcartConstate : boolean) : boolean;
var MtPorts,Ecart,MtSituationNet : Double;
begin
    Result := true;
    MtPorts := CalculPort(True,TOBporcs);
    MtSituationNet := TOBFacture.getDouble('GP_TOTALHTDEV');
    Ecart := Arrondi(MontantMarche-(MtSituationNet+MtPorts+DejaFacture-MtHorsDevis),Dev.decimale);
    if Ecart <> 0 then
    begin
      //
      EcartConstate := true;
      result := AjouteLigneEcart (Ecart);
      if Result then RenumeroteLaPiece;
      if not Result then CodeErreur := 'Erreur dans la gestion des écarts';
    end;
end;

procedure TGenFacAvanc.RestitueAvancementDevis ;

  procedure DeduitAvancLigne (TOBL : TOB);
  var cledoc : r_cledoc;
      QteDejaFactPrec,MtDejaFactPrec,PourcentAvanc,Mtmarche : double;
      SQL,SQLF : string;
      QQ : TQuery;
  begin
    DecodeRefPiece(TOBL.GetString('GL_PIECEORIGINE'),Cledoc);
    QteDejaFactPrec := TOBL.GetValue('QTECUMULEFACTINIT');
    MtDejaFactPrec := TOBL.GetValue('MTCUMULEFACTINIT');
    Mtmarche := TOBL.GetValue('BLF_MTMARCHE');
    if not Avancement then
    begin
      PourcentAvanc := 0
    end else
    begin
      if Mtmarche > 0 then
      begin
        PourcentAvanc := ARRONDI(MtDejaFactPrec/Mtmarche*100,2);
      end else
      begin
        if QteDejaFactPrec = 0 then PourcentAvanc := 0 else PourcentAvanc := 100;
      end;
    end;
    //
    SQL := 'SELECT BLF_MTMARCHE '+
           'FROM LIGNEFAC WHERE '+WherePiece(cledoc,ttdLigneFac,True,True);
    QQ := OpenSQL(SQL,True,1,'',true);
    if not QQ.eof then
    begin
      //
      // ---- LIGNE
      //
      SQL := 'UPDATE LIGNE SET '+
             'GL_QTEPREVAVANC='+Strfpoint(QteDejaFactPrec)+','+
             'GL_QTESIT='+Strfpoint(QteDejaFactPrec)+','+
             'GL_POURCENTAVANC='+Strfpoint(PourcentAvanc)+' '+
             'WHERE '+WherePiece(cledoc,ttdligne,True,True);
      //
      // -- LIGNEFAC
      //
      SQLF := 'UPDATE LIGNEFAC SET ';
      if Avancement then
      begin
         SQLF := SQLF + 'BLF_QTECUMULEFACT='+StrfPoint(QteDejaFactPrec)+',BLF_MTCUMULEFACT='+StrfPoint(MtDejaFactPrec)+',';
      end;
      SqlF := SqlF + 'BLF_QTEDEJAFACT='+StrfPoint(QteDejaFactPrec)+','+
             'BLF_MTDEJAFACT='+StrfPoint(PourcentAvanc)+','+
             'BLF_POURCENTAVANC='+StrfPoint(PourcentAvanc)+' ' +
             'WHERE '+WherePiece(cledoc,ttdligneFac,True,True);
    end;
    ferme (QQ);

    if ExecuteSQL(Sql) <=0 then
    begin
      CodeErreur := 'problème Enregistrement dans devis';
      V_PGI.Ioerror := OeUnknown;
      Exit;
    end;
    if ExecuteSQL(SqlF) <=0 then
    begin
      CodeErreur := 'problème Enregistrement dans devis';
      V_PGI.Ioerror := OeUnknown;
      Exit;
    end;
  end;

  procedure DeduitAvancLigneOUV (TOBL,TOBO : TOB);
  var cledoc : r_cledoc;
      QteDejaFactPrec,MtDejaFactPrec,Mtmarche,PourcentAvanc : double;
      SQL,SQLF : string;
      QQ : TQuery;
  begin
    DecodeRefPieceOUv(TOBO.GetString('BLO_PIECEORIGINE'),Cledoc);
    QteDejaFactPrec := TOBO.GetValue('QTECUMULEFACTINIT');
    MtDejaFactPrec := TOBO.GetValue('MTCUMULEFACTINIT');
    Mtmarche := TOBO.GetValue('BLF_MTMARCHE');
    //
    if not Avancement then
    begin
      PourcentAvanc := 0
    end else
    begin
      if Mtmarche > 0 then
      begin
        PourcentAvanc := ARRONDI(MtDejaFactPrec/Mtmarche*100,2);
      end else
      begin
        if QteDejaFactPrec = 0 then PourcentAvanc := 0 else PourcentAvanc := 100;
      end;
    end;

    SQL := 'SELECT BLF_MTMARCHE '+
           'FROM LIGNEFAC WHERE '+WherePiece(cledoc,ttdLigneFac,false)+' AND '+
           'BLF_UNIQUEBLO='+InttOStr(cledoc.UniqueBlo);
    QQ := OpenSQL(SQL,True,1,'',true);
    if not QQ.eof then
    begin
      //
      // -- LIGNEFAC
      //
      SQLF := 'UPDATE LIGNEFAC SET ';
      if Avancement then
      begin
         SQLF := SQLF + 'BLF_QTECUMULEFACT='+StrfPoint(QteDejaFactPrec)+
                 ',BLF_MTCUMULEFACT='+StrfPoint(MtDejaFactPrec)+',';
      end;
      SqlF := SqlF + 'BLF_QTEDEJAFACT='+StrfPoint(QteDejaFactPrec)+','+
             'BLF_MTDEJAFACT='+StrfPoint(MtDejaFactPrec)+',' +
             'BLF_POURCENTAVANC='+STRFPOINT (PourcentAvanc)+' '+
             'WHERE '+WherePiece(cledoc,ttdligneFac,false)+' AND '+
             'BLF_UNIQUEBLO='+InttOStr(cledoc.UniqueBlo);
      //
      if ExecuteSQL(SqlF)<=0 then
      begin
        CodeErreur := 'problème Enregistrement dans Ouvrage/devis';
        V_PGI.Ioerror := OeUnknown;
        Exit;
      end;
      //
    end;
    ferme (QQ);

  end;

  procedure RestitueAvancDevisOuv (TOBL,TOBOuvrage : TOB);
  var IndiceNomen,II : Integer;
      TOBOUV : TOB;
  begin
    IndiceNomen := TOBL.GetValue('GL_INDICENOMEN')-1;
    //FV1 : 28/07/2016
    if IndiceNomen < 0 then exit;
    if TOBOuvrage.detail.count = 0 then exit;
    if (Indicenomen) > (TOBOuvrage.detail.count ) then exit;
    //
    TOBOuv := TOBOuvrage.detail[IndiceNomen];
    for II := 0 to TOBOUV.Detail.Count -1 do
    begin
      DeduitAvancLigneOUV (TOBL,TOBOUV.detail[II]);
    end;
  end;

var II : Integer;
    TOBL : TOB;
begin
  for II := 0 to TOBDevis.Detail.count -1 do
  begin
    TOBL := TOBDevis.detail[II];
    //
    if TOBL.GetString('GL_PIECEORIGINE')='' then continue;
    if TOBL.GetValue('GL_TYPELIGNE')<>'ART' then continue;
    if TOBL.getInteger('OLD_QTESITUATION') = 0 then continue;
    //
    DeduitAvancLigne (TOBL);
    if isOuvrage(TOBL) then
    begin
      RestitueAvancDevisOuv (TOBL,TOBOuvrage);
    end;

  end;
end;

procedure TGenFacAvanc.SetPieceMorte (TOBFacture : TOB);
var OldEcr, OldStk: RMVT;
    SQl : string;
begin
  if TOBFacture.GetValue('GP_REFCOMPTABLE')<> '' then
  begin
    DetruitCompta(TOBFacture, Now, OldEcr, OldStk,true);
  end;
  if V_PGI.IOError = oeok then
  begin
    SQl := 'UPDATE PIECE SET GP_VIVANTE="-",GP_REFCOMPTABLE="" WHERE '+
           'GP_NATUREPIECEG="'+TOBFacture.GetString('GP_NATUREPIECEG')+'" AND '+
           'GP_SOUCHE="'+TOBFacture.GetString('GP_SOUCHE')+'" AND '+
           'GP_NUMERO='+IntToStr(TOBFacture.GetValue('GP_NUMERO'))+' AND '+
           'GP_INDICEG='+IntToStr(TOBFacture.GetValue('GP_INDICEG'));
    if ExecuteSQL(Sql) <= 0 then
    begin
      CodeErreur := 'problème réajustement précédente situation';
      V_PGI.IOError := oeUnknown;
    end;
    SQl := 'UPDATE LIGNE SET GL_VIVANTE="-" WHERE '+
           'GL_NATUREPIECEG="'+TOBFacture.GetString('GP_NATUREPIECEG')+'" AND '+
           'GL_SOUCHE="'+TOBFacture.GetString('GP_SOUCHE')+'" AND '+
           'GL_NUMERO='+IntToStr(TOBFacture.GetValue('GP_NUMERO'))+' AND '+
           'GL_INDICEG='+IntToStr(TOBFacture.GetValue('GP_INDICEG'));
    if ExecuteSQL(Sql) <= 0 then
    begin
      CodeErreur := 'problème réajustement précédente situation';
      V_PGI.IOError := oeUnknown;
    end;

    if (Pos(TOBFacture.GetString('GP_NATUREPIECEG'),'FBP;FBT;DAC;DAP')>0) then
    begin
      if (Pos(TOBFacture.GetValue('AFF_GENERAUTO'),'AVA;DAC;')>0) then
      begin
        SQL := 'UPDATE BSITUATIONS SET BST_VIVANTE="-" WHERE '+
               'BST_NATUREPIECE="'+TOBFacture.GetString('GP_NATUREPIECEG')+'" AND '+
               'BST_SOUCHE="'+TOBFacture.GetString('GP_SOUCHE')+'" AND '+
               'BST_NUMEROFAC='+IntToStr(TOBFacture.GetValue('GP_NUMERO'));
        if ExecuteSQL(SQl) <= 0 then
        begin
          CodeErreur := 'problème réajustement situation';
          V_PGI.IOError := oeUnknown;
        end;
      end;
    end;
  end;

end;


end.

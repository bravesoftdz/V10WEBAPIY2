unit FactSituations;

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
{$ENDIF}
     uTob,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     AglInit,
     UentCommun,
     SAISUTIL,
     ENTgc,
     OptimizeOuv,
     ParamSoc,
     FactTvaMilliem,
     Forms

    ;

type
  T_genereFacture = class (TObject)
  private
    DEV : Rdevise;
    NewNature : string;
    cledoc : R_cledoc;
    IndiceOuv: integer;
    TOBPiece :TOB;
    TOBPiece_O :TOB;
    TOBAffaireInterv :TOB;
    TOBPieceTrait:TOB;
    TOBmetres :TOB;
    TOBBasesL :TOB;
    TOBBases :TOB;
    TOBBases_O :TOB;
    TOBBasesSaisies :TOB;
    TOBEches :TOB;
    TOBEches_O :TOB;
    TOBPorcs :TOB;
    TOBPorcs_O :TOB;
    TOBTiers :TOB;
    TOBArticles :TOB;
    TOBConds :TOB;
    TOBTarif :TOB;
    TOBComms :TOB;
    TOBCatalogu :TOB;
    // Adresses
    TOBAdresses :TOB;
    TOBAdresses_O :TOB;
    TOBNomenclature :TOB;
    TOBN_O :TOB;
    TOBDim :TOB;
    TOBDesLots :TOB;
    TOBLOT_O :TOB;
    TOBAcomptes :TOB;
    TOBAcomptes_O :TOB;
    TOBAffaire :TOB;
    TOBCPTA :TOB;
    TOBANAP :TOB;
    TOBANAS :TOB;
    TOBOuvrage :TOB;
    TOBOuvrage_O :TOB;
    TOBOuvragesP :TOB;
    TOBLIENOLE :TOB;
    TOBLIENOLE_O :TOB;
    TOBPieceRG :TOB;
    TOBPieceRG_O :TOB;
    TOBBasesRG :TOB;
    TOBBASESRG_O :TOB;
    TOBLIGNERG :TOB;
    TOBLienDEVCHA :TOB;
    OptimizeOuv : TOptimizeOuv;
    TOBTablette :TOB;
    TOBTimbres :TOB;
    TOBSSTRAIT :TOB;
    TOBVTECOLL : TOB;
    TOBAFormule : TOB;
    lesAcomptes: TOB;
    TheRepartTva : TREPARTTVAMILL;
    //
    InAvancement : Boolean;
    TypeFacturation : string;
    //
    VenteAchat :string;
    GereStatPiece :Boolean;
    EstAvoir :Boolean;
    TypeCom :string;
    PassP : string;
    OkCpta :boolean;
    CompAnalP : string;
    CompAnalL :string;
    CompStockP :string;
    CompStockL :string;
    OkCptaStock :boolean;

    GereEche : String;
    GereAcompte :Boolean;
    ObligeRegle : Boolean;
    OuvreAutoPort :Boolean;
    ApplicRetenue :boolean;
    GppReliquat : Boolean; { NEWPIECE }
    MessageValid : string;
    NumSituation : Integer;
    //
    procedure LoadLesTOB;
    procedure ToutAllouer;
    procedure ToutLiberer;
    procedure LoadLesGCS;
    procedure ChargelaPiece(ElimineLigneSoldee: boolean = True);
    procedure SupLigneRgUnique;
    procedure LoadLesArticles;
    function TransformePiece(DatePiece: TDateTime) : integer;
    procedure InitDocument;
    procedure ValideLaPiece;
    procedure DelOrUpdateAncien;
    procedure ChargeFromNature(nature: string);
    procedure InitToutModif(Datemaj : TdateTime = 0);
    procedure ValideLaNumerotation;
    procedure ModifRefGC(TOBPiece, TOBAcomptes: TOB);
    procedure GestionRgUnique(TOBPIECE, TOBPIECERG, TOBBASES, TOBBASESRG,TOBTIers, TOBAffaire, TOBLigneRG: TOB);
    procedure GereLesReliquats;
    procedure ValideTiers;
    procedure ValideLaCompta(OldEcr, OldStk: RMVT);
    function GetMessage: string;
    procedure ValideNumeroPiece;
    procedure AjouteEvent(TOBPIECE_O, TOBPIECE: TOB; var MessEvent: string;OKPASOK: integer);
  public
    constructor create;
    destructor destroy; override;
    property LibRetour : string read GetMessage;
    function GenereFacDefinitif(Xcledoc: R_CLEDOC;DateSit: TdateTime): integer;
  end;

function DemandeDatesFacturation(var DateFac: TDateTime) : boolean;

implementation
uses FactGrp,UtilTOBPiece,FactTOB,BTStructChampSup,Affaireutil,
     UCotraitance,BTSAISIEPAGEMETRE_TOF,FactTimbres,factligneBase,
     FactAdresse,factNomen,FactLotSerie,FactUtil,FactCpta,FactOuvrage,
     BTPUtil,FactPiece,factRG,FactArticle,factComm,FactTiers,FactAdresseBTP,
     FactCalc,UCumulCollectifs
      ;

function DemandeDatesFacturation(var DateFac: TDateTime) : boolean;
var TobDates : TOB;
begin
  TOBDates := TOB.Create ('LES DATES', nil,-1);
  TOBDates.AddChampSupValeur('RETOUROK','-');
  TOBDates.AddChampSupValeur('DATFAC',Idate1900);
  TOBDates.AddChampSupValeur('DATESITUATION','-');
  TOBDates.AddChampSupValeur('TYPEDATE','Date de Facturation');
  TRY
    TheTOB := TOBDates;
    AGLLanceFiche('BTP','BTDEMANDEDATES','','','');
    TheTOB := nil;
    if TOBDates.getValue('RETOUROK')='X' then
    begin
    	DateFac := TOBDates.GetDateTime('DATFAC');
    end;
  FINALLY
  	result := (TOBDates.getValue('RETOUROK')='X');
  	freeAndNil(TOBDates);
  END;
end;

procedure T_genereFacture.ValideLaCompta(OldEcr, OldStk: RMVT);
begin
  if VH_GC.GCIfDefCEGID then
     RechLibTiersCegid(VenteAchat, TobTiers, TOBPiece, TobAdresses);
  if not PassationComptable( TOBPiece,TOBOUvrage,TOBOuvragesP, TOBBases,TOBBasesL, TOBEches,TOBPieceTrait,TOBAffaireInterv,TOBTiers, TOBArticles, TOBCpta, TOBAcomptes, TOBPorcs, TOBPIECERG, TOBBASESRG, TOBanaP,TOBanaS,TOBSSTRAIT,TOBVTECOLL, DEV, OldEcr,OldStk, True, true ) then
  begin
    MessageValid := 'Erreur / ecriture comptable';
    V_PGI.IoError := oeLettrage;
  end;
  LibereParamTimbres;
end;

procedure T_genereFacture.ChargeFromNature (nature : string);
var PassP : string;
begin
  {Caractéristiques nature}
  VenteAchat := GetInfoParPiece(nature, 'GPP_VENTEACHAT');
  GereStatPiece := GetInfoParPiece(nature, 'GPP_AFFPIECETABLE') = 'X';
  EstAvoir := (GetInfoParPiece(nature, 'GPP_ESTAVOIR') = 'X');
  {commercial}
  TypeCom := GetInfoParPiece(nature, 'GPP_TYPECOMMERCIAL');
  {Comptabilité}
  PassP := GetInfoParPiece(nature, 'GPP_TYPEECRCPTA');
  OkCpta := ((PassP <> '') and (PassP <> 'RIE'));
  CompAnalP := GetInfoParPiece(nature, 'GPP_COMPANALPIED');
  CompAnalL := GetInfoParPiece(nature, 'GPP_COMPANALLIGNE');
  CompStockP := GetInfoParPiece(nature, 'GPP_COMPSTOCKPIED');
  CompStockL := GetInfoParPiece(nature, 'GPP_COMPSTOCKLIGNE');
  OkCptaStock := IsComptaStock(nature);

  GereEche := GetInfoParPiece(nature, 'GPP_GEREECHEANCE');
  GereAcompte := (GetInfoParPiece(nature, 'GPP_ACOMPTE') = 'X') AND (NewNature <> 'DAC');
  ObligeRegle := (GetInfoParPiece(nature, 'GPP_OBLIGEREGLE') = 'X');
  OuvreAutoPort := (GetInfoParPiece(nature, 'GPP_OUVREAUTOPORT') = 'X');
  ApplicRetenue := (GetInfoParPiece(nature, 'GPP_APPLICRG') = 'X');
  GppReliquat := (GetInfoParPiece(nature, 'GPP_RELIQUAT') = 'X'); { NEWPIECE }
end;


procedure T_genereFacture.LoadLesArticles;
var i, iArt: integer;
  TOBL, TOBArt, TOBC, TOBCata, LocAnaP, LocAnaS: TOB;
  RefUnique, RefCata, RefFour: string;
  NaturePieceG: string;
  lcledoc : R_CLEDOC;
begin
  FillChar (lcledoc,SizeOf(lcledoc),#0);
  TOBCata := nil;
  iArt := 0;
  NaturePieceG := TOBPiece.GetValue('GP_NATUREPIECEG');
  RecupTOBArticle (TOBPiece,TOBArticles,TOBAFormule,TOBConds,lcledoc,cledoc,'VEN');
	if OptimizeOuv <> nil then OptimizeOuv.fusionneArticles (TOBArticles);
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[i];
    if i = 0 then
    begin
      iArt := TOBL.GetNumChamp('GL_ARTICLE');
    end;
    RefCata := '';
    RefFour := '';
    RefUnique := TOBL.GetValeur(iArt);
    if RefUnique <> '' then TobArt := TOBArticles.FindFirst(['GA_ARTICLE'], [RefUnique], False) else TOBArt := nil;
    if (refUnique <> '') and (TobArt <> nil) then
    begin
      InitLesSupLigne(TOBL,false);
      // MODIF AC CATALOGU
      //if (VenteAchat='ACH') and ((CataFourn) or (TOBL.GetValue('GL_REFARTTIERS')<>'')) then AjouteCatalogueArt(TOBL,TOBArt) ;
      CodesLigneToCodesArt(TOBL, TOBArt);
      TOBL.PutValue('BNP_TYPERESSOURCE',TOBArt.getValue('BNP_TYPERESSOURCE'));
      TOBL.PutValue('QTEORIG', TOBL.GetValue('GL_QTESTOCK'));
    end;
    if (RefUnique <> '') or ((RefCata <> '') and (RefFour <> '')) then
    begin
      if (TobArt = nil)  then
      begin
        // Erreur a gerer
      end else
      begin
        // Chargement du commercial
        AjouteRepres(TOBL.GetValue('GL_REPRESENTANT'), GetInfoParPiece(NewNature, 'GPP_TYPECOMMERCIAL'), TOBComms);
        if ((OkCpta) or (OkCptaStock)) then
        begin
          TOBC := ChargeAjouteCompta(TOBCpta, TOBPiece, TOBL, TOBArt, TOBTiers, TOBCata, TobAffaire, True,TOBTablette);
          if (TOBC <> nil) and (TOBL.Detail.Count > 0) then
          begin
            if ((TOBL.Detail[0].Detail.Count <= 0) and (OkCpta)) then LocAnaP := TOBAnaP else LocAnaP := nil;
            if ((TOBL.Detail[1].Detail.Count <= 0) and (OkCptaStock)) then LocAnaS := TOBAnaS else LocAnaS := nil;
            PreVentileLigne(TOBC, LocAnaP, LocAnaS, TOBL);
          end;
        end;
      end;
    end;
    // Modif BTP le 01/04/03 (Gestion Prix posés)
    if TOBL.GetValue('GL_TYPEARTICLE') = 'PRE' then
    begin
      if (Pos(TOBL.GetValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne)>0) and (not IsLigneExternalise(TOBL) ) then
      begin
      	TOBL.putvalue('GL_TPSUNITAIRE', 1)
      end else
      begin
      	TOBL.putvalue('GL_TPSUNITAIRE', 0);
      end;
    end;
    TOBL.PutValue('LIVREORIGINE',TOBL.getValue('GL_QTESTOCK')); // Qte initiale de sortie de stock
  end;
end;


procedure T_genereFacture.LoadLesGCS;
var Q : TQuery;
begin
  Q := OpenSQL('Select * from Commun where CO_TYPE="GCS" and CO_LIBRE Like "%'+VH_GC.GCMarcheVentilAna+'%"',True,-1, '', True);
  if not Q.eof Then TOBTablette.LoadDetailDB ('COMMUN','','',Q,false,true);
  ferme (Q);
end;

procedure T_genereFacture.ToutAllouer;
begin
  TOBVTECOLL := TOB.Create ('LES VTE COLL',nil,-1);
  // Pièce
  TOBPiece := TOB.Create('PIECE', nil, -1);
  TOBPiece_O := TOB.Create('', nil, -1);
  // Modif BTP
  AddLesSupEntete(TOBPiece);
  AddLesSupEntete(TOBPiece_O);
  // ---
  TOBAffaireInterv := TOB.Create ('LES CO-SOUSTRAITANTS',nil,-1);
  //
  TOBPieceTrait := TOB.Create ('LES LIGNES EXTRENALISE',nil,-1);
  TOBmetres := TOB.Create ('LES LIGNES METRES',nil,-1);
  TOBBasesL := TOB.Create('LES BASES LIGNES', nil, -1);
  TOBBases := TOB.Create('BASES', nil, -1);
  TOBBases_O := TOB.Create('', nil, -1);
  TOBBasesSaisies:=TOB.Create('BASES',Nil,-1);
  TOBEches := TOB.Create('Les ECHEANCES', nil, -1);
  TOBEches_O := TOB.Create('', nil, -1);
  TOBPorcs := TOB.Create('PORCS', nil, -1);
  TOBPorcs_O := TOB.Create('', nil, -1);
  // Fiches
  TOBTiers := TOB.Create('TIERS', nil, -1);
  TOBTiers.AddChampSup('RIB', False);
  TOBArticles := TOB.Create('ARTICLES', nil, -1);
  TOBConds := TOB.Create('CONDS', nil, -1);
  TOBTarif := TOB.Create('TARIF', nil, -1);
  TOBComms := TOB.Create('COMMERCIAUX', nil, -1);
  TOBCatalogu := TOB.Create('LECATALOGUE', nil, -1);
  // Adresses
  TOBAdresses := TOB.Create('LESADRESSES', nil, -1);
  TOBAdresses_O := TOB.Create('', nil, -1);
  if GetParamSoc('SO_GCPIECEADRESSE') then
  begin
    TOB.Create('PIECEADRESSE', TOBAdresses, -1); {Livraison}
    TOB.Create('PIECEADRESSE', TOBAdresses, -1); {Facturation}
  end else
  begin
    TOB.Create('ADRESSES', TOBAdresses, -1); {Livraison}
    TOB.Create('ADRESSES', TOBAdresses, -1); {Facturation}
  end;
  // Divers
  TOBNomenclature := TOB.Create('NOMENCLATURES', nil, -1);
  TOBN_O := TOB.Create('', nil, -1);
  TOBDim := TOB.Create('', nil, -1);
  TOBDesLots := TOB.Create('', nil, -1);
  TOBLOT_O := TOB.Create('', nil, -1);
  TOBAcomptes := TOB.Create('', nil, -1);
  TOBAcomptes_O := TOB.Create('', nil, -1);
  // Affaires
  TOBAffaire := TOB.Create('AFFAIRE', nil, -1);
  // Comptabilité
  TOBCPTA := CreerTOBCpta;
  TOBANAP := TOB.Create('', nil, -1);
  TOBANAS := TOB.Create('', nil, -1);
  // Ouvrages
  TOBOuvrage := TOB.Create('OUVRAGES', nil, -1);
  TOBOuvrage_O := TOB.Create('', nil, -1);
  // ouvrages Plat
  TOBOuvragesP := TOB.Create('LES OUVRAGES PLAT', nil, -1);
  // textes debut et fin
  TOBLIENOLE := TOB.Create('LIENS', nil, -1);
  TOBLIENOLE_O := TOB.Create('', nil, -1);
  // retenues de garantie
  TOBPieceRG := TOB.create('PIECERRET', nil, -1);
  TOBPieceRG_O := TOB.create('', nil, -1);
  // Bases de tva sur RG
  TOBBasesRG := TOB.create('BASESRG', nil, -1);
  TOBBASESRG_O := TOB.create('', nil, -1);
  // --
  TOBLIGNERG := nil;
  {$IFDEF BTP}
  TOBLienDEVCHA     := TOB.Create('LES LIENS DEVCHA', nil, -1);
  (* OPTIMIZATION *)
  OptimizeOuv       := TOptimizeOuv.create;
  (* ------------- *)
  {$ENDIF}
  // OPTIMISATION
	TOBTablette       := TOB.Create ('LA TABLETTE GCS',nil,-1);
  TOBTimbres        := TOB.Create ('LES TIMBRES',nil,-1);
  TOBSSTRAIT        := TOB.Create ('LES SOUS TRAITS',nil,-1);
  TOBAFormule := TOB.Create('LES FORMULES',nil,-1);
  TheRepartTva := TREPARTTVAMILL.create (application.MainForm) ;
  TheRepartTva.TOBBases := TOBBases;
  TheRepartTva.TOBpiece := TOBPiece;
  TheRepartTva.TOBOuvrages := TOBOUvrage;
end;

procedure T_genereFacture.ToutLiberer;
begin
  TOBVTECOLL.free;
  TheRepartTva.free;
  TOBPiece.Free;
  TOBPiece := nil;
  TOBPiece_O.Free;
  TOBPiece_O := nil;
  TOBTiers.Free;
  TOBTiers := nil;
  TOBArticles.Free;
  TOBArticles := nil;
  TOBCatalogu.Free;
  TOBCatalogu := nil;
  TOBConds.Free;
  TOBConds := nil;
  TOBComms.Free;
  TOBComms := nil;
  TOBEches.Free;
  TOBEches := nil;
  TOBEches_O.Free;
  TOBEches_O := nil;
  TOBBases.Free;
  TOBBases := nil;
  //
  FreeAndNil(TOBmetres);
  //
  TOBBasesL.free;
  TOBBasesL := nil;
  TOBBases_O.Free;
  TOBBases_O := nil;
  TOBBasesSaisies.Free;
  TOBBasesSaisies:=Nil;
  TOBPorcs.Free;
  TOBPorcs := nil;
  TOBPorcs_O.Free;
  TOBPorcs_O := nil;
  TOBTarif.Free;
  TOBTarif := nil;
  TOBAdresses.Free;
  TOBAdresses := nil;
  TOBAdresses_O.Free;
  TOBAdresses_O := nil;
  TOBAffaire.Free;
  TOBAffaire := nil;
  TOBDim.Free;
  TOBDim := nil;
  TOBCpta.Free;
  TOBCpta := nil;
  TOBAnaP.Free;
  TOBAnaP := nil;
  TOBAnaS.Free;
  TOBAnaS := nil;
  TOBDesLots.Free;
  TOBDesLots := nil;
  TOBLOT_O.Free;
  TOBLOT_O := nil;
  TOBNomenclature.Free;
  TOBNomenclature := nil;
  TOBN_O.Free;
  TOBN_O := nil;
  TOBAcomptes.Free;
  TOBAcomptes := nil;
  TOBAcomptes_O.Free;
  TOBAcomptes_O := nil;
  // MOdif BTP
  TOBOuvrage.free;
  TOBOuvrage := nil;
  TOBOuvragesP.free;
  TOBOuvragesP := nil;
  TOBOuvrage_O.free;
  TOBOuvrage_O := nil;
  TOBLIENOLE.free;
  TOBLIENOLE_O.free;
  TOBLIENOLE := nil;
  TOBLIENOLE := nil;
  TOBPIECERG.free;
  TOBPIECERG_O.free;
  TOBBASESRG.free;
  TOBBASESRG_O.free;
  // Modif BTP
  if TOBLigneRG <> nil then
  begin
    TOBLigneRG.free;
    TOBLigneRG := nil;
  end;
  {$IFDEF BTP}
  TOBLienDEVCHA.free;
  OptimizeOuv.free;
  {$ENDIF}
  // ---------
  // OPTIMISATION
	TOBTablette.free;
  // --
	FreeAndNil(TOBPieceTrait);
  FreeAndNil(TOBAffaireInterv);
  FreeAndNil(TOBTimbres);
  TOBSSTRAIT.Free;
  TOBAFormule.free;
end;

procedure T_genereFacture.InitDocument;
begin
  InAvancement := false;
  TypeFacturation := '';
  //
  TOBVTECOLL.ClearDetail;
  TOBPiece.ClearDetail;
  TOBPiece.InitValeurs(false);
  TOBPiece_O.ClearDetail;
  TOBPiece_O.InitValeurs(false);
  TOBTiers.InitValeurs(false);
  TOBArticles.ClearDetail;
  TOBCatalogu.ClearDetail;
  TOBConds.ClearDetail;
  TOBComms.ClearDetail;
  TOBEches.ClearDetail;
  TOBEches_O.ClearDetail;
  TOBBases.ClearDetail;
  TOBBases_O.ClearDetail;
  TOBmetres.ClearDetail;
  TOBBasesL.ClearDetail;
  TOBBasesSaisies.ClearDetail;
  TOBPorcs.ClearDetail;
  TOBPorcs_O.ClearDetail;
  TOBTarif.ClearDetail;
  TOBAdresses.ClearDetail;
  TOBAdresses_O.ClearDetail;
  TOBAffaire.ClearDetail;;
  TOBAffaire.InitValeurs(false);
  TOBDim.ClearDetail;
  TOBCpta.ClearDetail;
  TOBAnaP.ClearDetail;
  TOBAnaS.ClearDetail;
  TOBDesLots.ClearDetail;
  TOBLOT_O.ClearDetail;
  TOBNomenclature.ClearDetail;
  TOBN_O.ClearDetail;
  TOBAcomptes.ClearDetail;
  TOBAcomptes_O.ClearDetail;
  TOBOuvrage.ClearDetail;
  TOBOuvragesP.ClearDetail;
  TOBOuvrage_O.ClearDetail;
  TOBLIENOLE.ClearDetail;
  TOBLIENOLE.ClearDetail;
  TOBPIECERG.ClearDetail;
  TOBPIECERG_O.ClearDetail;
  TOBBASESRG.ClearDetail;
  TOBBASESRG_O.ClearDetail;
  if TOBLigneRG <> nil then TOBLigneRG.ClearDetail;
  TOBLienDEVCHA.ClearDetail;
  OptimizeOuv.initialisation;
	TOBTablette.ClearDetail;
	TOBPieceTrait.ClearDetail;
  TOBAffaireInterv.ClearDetail;
  TOBTimbres.ClearDetail;
  TOBSSTRAIT.ClearDetail;
  TOBAFormule.ClearDetail;
  //
  VenteAchat :='';
  GereStatPiece :=false;
  EstAvoir :=false;
  TypeCom :='';
  PassP :='';
  OkCpta :=false;
  CompAnalP :='';
  CompAnalL :='';
  CompStockP :='';
  CompStockL :='';
  OkCptaStock :=false;
  //
  GereEche :='';
  GereAcompte :=false;
  ObligeRegle :=false;
  OuvreAutoPort :=false;
  ApplicRetenue :=false;
  GppReliquat :=false; { NEWPIECE }
  MessageValid := '';
  NumSituation := 0;
end;

constructor T_genereFacture.create;
begin
  LesAcomptes := nil;
  ToutAllouer;
end;

destructor T_genereFacture.destroy;
begin
  ReinitTOBAffaires;
  ToutLiberer;
  inherited;
end;

procedure T_genereFacture.LoadLesTOB;
var Q: TQuery;
begin
  MemoriseChampsSupLigneETL (cledoc.NaturePiece,true);
  MemoriseChampsSupLigneOUV (cledoc.NaturePiece);
  //
  LoadPieceLignes(CleDoc, TobPiece,true,false,false,false);
  PieceAjouteSousDetail(TOBPiece,true,false,true);
  //
  LoadLesVTECOLLECTIF (CleDoc,TOBVTECOLL);
  // --
  if TOBPiece.getValue('GP_AFFAIREDEVIS')<> '' then
  begin
    StockeCetteAffaire (string(TOBPiece.getValue('GP_AFFAIREDEVIS')));
  end;
  if TOBPiece.getValue('GP_AFFAIRE')<> '' then
  begin
    StockeCetteAffaire (string(TOBPiece.getValue('GP_AFFAIRE')));
  end;
  // --
  // Lecture des affectations document
  LoadLaTOBPieceTrait(TOBpieceTrait,cledoc,'');
  LoadLesSousTraitants(cledoc,TOBSSTRAIT);

  // Lecture metres
  loadlesMetres (Cledoc,TOBMetres);
  LoadlesTimbres(cledoc,TOBTimbres);

  // Lecture bases Lignes
  Q := OpenSQL('SELECT * FROM LIGNEBASE WHERE ' + WherePiece(CleDoc, ttdLigneBase, False), True,-1, '', True);
  TOBBasesL.LoadDetailDB('LIGNEBASE', '', '', Q, False);
  Ferme(Q);
{$IFDEF BTP}
  OrdonnelignesBases (TOBBasesL);
{$ENDIF}

  // Lecture bases
  Q := OpenSQL('SELECT * FROM PIEDBASE WHERE ' + WherePiece(CleDoc, ttdPiedBase, False), True,-1, '', True);
  TOBBases.LoadDetailDB('PIEDBASE', '', '', Q, False);
  Ferme(Q);

  // Lecture Echéances
  Q := OpenSQL('SELECT *,(SELECT T_LIBELLE FROM TIERS WHERE T_TIERS=GPE_FOURNISSEUR AND T_NATUREAUXI="FOU") AS LIBELLE  FROM PIEDECHE WHERE ' + WherePiece(CleDoc, ttdEche, False), True,-1, '', True);
  TOBEches.LoadDetailDB('PIEDECHE', '', '', Q, False);
  Ferme(Q);

  // Lecture Ports
  Q := OpenSQL('SELECT * FROM PIEDPORT WHERE ' + WherePiece(CleDoc, ttdPorc, False), True,-1, '', True);
  TOBPorcs.LoadDetailDB('PIEDPORT', '', '', Q, False);
  Ferme(Q);

  // Lecture Adresses
  LoadLesAdresses(TOBPiece, TOBAdresses);

  // Lecture Nomenclatures
  LoadLesNomen(TOBPiece, TOBNomenclature, TOBArticles, CleDoc);

  // Lecture Lot
  LoadLesLots(TOBPiece, TOBDesLots, CleDoc);

  // Lecture ACompte
  LoadLesAcomptes(TOBPiece, TOBAcomptes, CleDoc, lesAcomptes);

  // Lecture Analytiques
  LoadLesAna(TOBPiece, TOBAnaP, TOBAnaS);

  // Modif BTP
  // chargement textes debut et fin
  LoadLesBlocNotes([], TOBLienOle, Cledoc);

  // Chargement des retenues de garantie et Tva associe
  if GetParamSoc('SO_RETGARANTIE') then LoadLesRetenues(TOBPiece, TOBPieceRG, TOBBasesRG, taModif);
  // Lecture Ouvrages
  LoadLesOuvrages(TOBPiece, TOBOuvrage, TOBArticles, Cledoc, IndiceOuv,OptimizeOuv);
  LoadLesOuvragesPlat(TOBPiece, TOBOuvragesP, Cledoc);
  TOBBasesSaisies.Dupliquer(TOBBases,True,True);
  // chargement des details ouvrage ou nomenclatures
  DEV.Code := TOBPIECE.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := TOBPiece.GetValue('GP_TAUXDEV');
  MemoriseNumLigne(TobPiece);
  LoadLesGCS;
  //
  LoadLaTOBAffaireInterv(TOBAffaireInterv,TOBPiece.getValue('GP_AFFAIRE'));
end;


procedure T_genereFacture.SupLigneRgUnique;
var TOBL: TOB;
begin
  if (not RGMultiple(TOBpiece)) then
  begin
    TOBL := TOBPIECE.findfirst(['GL_TYPELIGNE'], ['RG'], true);
    if TOBL <> nil then
    begin
      TOBL.free;
    end;
  end;
end;


procedure T_genereFacture.ChargelaPiece( ElimineLigneSoldee: boolean = True);
var
  Ind: integer;
  TOBL, TOBA: TOB;
  NaturePiece,lastPrec : string;
begin
  lastPreC := '';
  IndiceOuv := 1;
  NaturePiece := cledoc.NaturePiece;
  {$IFDEF CHR}
  RemplirTOBHrdossier(GP_HRDOSSIER.Text, TobHrdossier);
  {$ENDIF}
  LoadLesTOB;
  //
  if (TOBVTECOLL.Detail.count = 0) then
  begin
    ConstitueVteCollectif (TOBPiece,TOBSSTRAIT,TOBBases,TOBVTECOLL);
  end;
  TypeFacturation := RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS'));
  InAvancement := (Pos(Typefacturation,'AVA;DAC')>0);
  //
  EvaluerMaxNumOrdre(TOBPiece);
  if TOBPieceRG.detail.count > 0 then ReajusteNumOrdre (TOBPiece,TOBPieceRg);
  TOBPiece.putValue('GP_RGMULTIPLE','-');
  SupLigneRgUnique;
  TOBPiece_O.Dupliquer(TOBPiece, True, True);
  TOBBasesRG_O.Dupliquer(TOBBASESRG, True, True);

  //if TOBPieceRG.detail.count = 0 then TOBBasesRG.clearDetail;
  //if (Pos(NaturePiece,'FBT;FBP')>0) and (TypeFacturation='DIR') then
  //begin
    // modification d'une facture de type directe (ancienne gestion de facturation)
  //  DefiniDejaFacture(TOBPiece,TOBOUvrage,DEV,TypeFacturation);
  //end else
  //begin
  //  MajQtesAvantSaisie (TOBpiece);
  //end;
  RemplirTOBTiers(TOBTiers, TOBPiece.GetString('GP_TIERS'), TOBPiece.GetString('GP_NATUREPIECEG'), False);
  TOBTiers.PutValue('RIB', TOBPiece.GetValue('GP_RIB'));
  LoadLesArticles;
  if OuvrageNonDifferencie (TOBpiece,TOBouvrage) then
  begin
    OuvrageDifferencie (TOBpiece,TOBOuvrage,InAvancement,DEV);
  end;
  if InAvancement then CalculeSousTotauxPiece(TOBPiece);

  for ind := 0 to TOBPiece_O.Detail.Count - 1 do
  begin
    TOBL := TOBPiece_O.Detail[ind];
    if TOBL <> nil then
    begin
      if TOBL.FieldExists('ANCIENNUMLIGNE') then
        TOBL.PutValue('GL_NUMLIGNE', TOBL.GetValue('ANCIENNUMLIGNE'));
      // Deplacé
      //SommationAchatDoc(TOBPiece, TOBL);
      //
      if InAvancement then
      begin
        CalculValoAvanc(TOBPIece, Ind, DEV);
      end;
    end;
    if TOBL.Detail.Count > 0 then
    begin
      TOBA := TOBL.Detail[0];
      TOBA.ClearDetail;
    end;
  end;
    // Correction FQ 12256 -- Les adresses sont manquantes quand le document provient des interventions
  if TOBAdresses.detail.count = 0 then
  begin
    if GetParamSoc('SO_GCPIECEADRESSE') then
    begin
      TOB.Create('PIECEADRESSE', TOBAdresses, -1); {Livraison}
      TOB.Create('PIECEADRESSE', TOBAdresses, -1); {Facturation}
    end else
    begin
      TOB.Create('ADRESSES', TOBAdresses, -1); {Livraison}
      TOB.Create('ADRESSES', TOBAdresses, -1); {Facturation}
    end;
    //
    TiersVersAdresses(TOBTiers, TOBAdresses, TOBPiece);
    AffaireVersAdresses(TOBAffaire,TOBAdresses,TOBPiece); //mcd 29/09/03 déplacé pour résoudre cas clt fact # ds affaire et adresse fact
    if (VenteAchat = 'VEN')  then
    begin
      LivAffaireVersAdresses (TOBAffaire,TOBAdresses,TOBPiece);
    end;

  end;

  TOBBases_O.Dupliquer(TOBBases, True, True);
  TOBAdresses_O.Dupliquer(TOBAdresses, True, True);
  TOBEches_O.Dupliquer(TOBEches, True, True);
  TOBPorcs_O.Dupliquer(TOBPorcs, True, True);
  TOBN_O.Dupliquer(TOBNomenclature, True, True);
  TOBLOT_O.Dupliquer(TOBdesLots, True, True);
  TOBAcomptes_O.Dupliquer(TOBAcomptes, True, True);
  TOBOuvrage_O.Dupliquer(TOBOuvrage, True, True);
  TOBLIENOLE_O.Dupliquer(TOBLIENOLE, True, True);
  TOBPIECERG_O.Dupliquer(TOBPIECERG, True, True);
  TheRepartTva.Charge; // voila voila
end;

Function T_genereFacture.TransformePiece (DatePiece : TDateTime) : integer;
var i: integer;
  TOBL,  TOBB: TOB;
  ToblOld: Tob; { NEWPIECE }
  OldNat, EtatVC, RefPiece: string;
begin
  Result := 0;
  OldNat := TOBPiece_O.GetValue('GP_NATUREPIECEG');
  CleDoc.NaturePiece := NewNature;
  CleDoc.Souche := GetSoucheG(CleDoc.NaturePiece, TOBPiece.GetValue('GP_ETABLISSEMENT'), TOBPiece.GetValue('GP_DOMAINE'));
  CleDoc.Indice := 0;
  cledoc.DatePiece := DatePiece;
  //
  MajFromCleDoc(TOBPiece, CleDoc);
  TOBPiece.PutValue('GP_DEVENIRPIECE', '');  TOBPiece.PutValue('GP_CODEORDRE',0) ;
  TOBPiece.PutValue('GP_REFCOMPTABLE', '');
  TOBPiece.PutValue('GP_REFCPTASTOCK', '');
  TOBPiece.PutValue('GP_ETATVISA', 'NON');
  TOBPiece.PutValue('GP_DATEVISA', iDate1900);
  TOBPiece.PutValue('GP_EDITEE', '-');
  TOBPiece.PutValue('GP_DATECREATION', DatePiece);
  if GetParamSoc('SO_GCPIECEADRESSE') then
  begin
    // Laisser comme c'est
  end else
  begin
    TOBPiece.PutValue('GP_NUMADRESSELIVR', -1);
    TOBPiece.PutValue('GP_NUMADRESSEFACT', -2);
  end;

  // Provisoire : Reinitialisation des échéances
//  TOBEches.ClearDetail;
//  GereEcheancesGC(TOBPiece, TOBTiers, TOBEches, TOBAcomptes, TOBPieceRG,TOBPieceTrait, taModif, DEV, False);

  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[i];
    { Récupère la ligne précédente dans la TobPiece_O }{ NEWPIECE }
    ToblOld := FindTobLigInOldTobPiece(Tobl, TobPiece_O);
    //
    RefPiece := '';
    if Assigned(ToblOld) then RefPiece := EncodeRefPiece(ToblOld);
    //
    MajFromCleDoc(TOBL, CleDoc);
    { TP - NEWPIECE - BUG 10752 }
    TOBL.PutValue('GL_PIECEPRECEDENTE', RefPiece);
    if TOBL.GetValue('GL_PIECEORIGINE') = '' then TOBL.PutValue('GL_PIECEORIGINE', RefPiece);
    EtatVC := TOBL.GetValue('GL_VALIDECOM');
    if EtatVC = 'VAL' then TOBL.PutValue('GL_VALIDECOM', 'AFF') else
      if EtatVC = 'AFF' then CommVersLigne(TOBPiece, TOBArticles, TOBComms, i + 1, False);
  end;
  //
  SetInfoPrecOuv (TOBOuvrage); // memorisation piece precedente et origine 
  //
  for i := 0 to TOBmetres.Detail.Count - 1 do
  begin
    TOBB := TOBmetres.Detail[i];
    MajFromCleDoc(TOBB, CleDoc);
  end;
  for i := 0 to TOBBases.Detail.Count - 1 do
  begin
    TOBB := TOBBases.Detail[i];
    MajFromCleDoc(TOBB, CleDoc);
  end;
  for i := 0 to TOBBasesL.Detail.Count - 1 do
  begin
    TOBB := TOBBasesL.Detail[i];
    MajFromCleDoc(TOBB, CleDoc);
  end;
  for i := 0 to TOBEches.Detail.Count - 1 do
  begin
    TOBB := TOBEches.Detail[i];
    MajFromCleDoc(TOBB, CleDoc);
  end;
  for i := 0 to TOBPieceRG.Detail.Count - 1 do
  begin
    TOBB := TOBPieceRG.Detail[i];
    MajFromCleDoc(TOBB, CleDoc);
  end;
  for i := 0 to TOBBasesRG.Detail.Count - 1 do
  begin
    TOBB := TOBBasesRG.Detail[i];
    MajFromCleDoc(TOBB, CleDoc);
  end;
  //
  NumeroteLignesGC(nil,TOBpiece, False, false);
  //
end;

procedure T_genereFacture.ValideNumeroPiece;
var NewNum: integer;
begin
  NewNum := SetNumberAttribution(TOBPiece.GetValue('GP_NATUREPIECEG'),TOBPiece.GetValue('GP_SOUCHE'), TOBPiece.GetValue('GP_DATEPIECE'),1);
  if NewNum > 0 then
  begin
    CleDoc.NumeroPiece := NewNum
  end;
end;


function T_genereFacture.GenereFacDefinitif(Xcledoc: R_CLEDOC; DateSit : TdateTime) : integer;
begin
  V_PGI.IOError := oeOk;
  Result := 0;
  cledoc := Xcledoc;
  if cledoc.NaturePiece = 'FBP' then NewNature := 'FBT'
  else if cledoc.NaturePiece = 'ABP' then NewNature := 'ABT'
  else if cledoc.NaturePiece = 'BAC' then NewNature := 'DAC'
  else Exit;

  ReinitTOBAffaires;
  InitDocument;
  ChargeFromNature (Cledoc.naturepiece);
  ChargelaPiece(false);
  ChargeFromNature (NewNature);
  Result := TransformePiece (DateSit);
  if Result <> 0 then Exit;
  TRY
    BEGINTRANS;
    ValideNumeroPiece;
    COMMITTRANS;
  EXCEPT
    ROLLBACK;
    result := -1;
  END;
  if Result <> 0 then Exit;
  TRY
    BEGINTRANS;
    ValideLaPiece;
    COMMITTRANS;
  EXCEPT
    result := -1;
    ROLLBACK;
  END;
  AjouteEvent (TOBPiece_O,TOBPiece,MessageValid,Result);
end;

procedure T_genereFacture.AjouteEvent(TOBPIECE_O,TOBPIECE: TOB; var MessEvent: string; OKPASOK : integer);
var QQ: TQuery;
  MotifPiece: TStrings;
  NumEvent: integer;
  MessOkPASOK : string;
  LeLibelle: string;
begin
  LeLibelle := Copy('Validation de ' + RechDom('GCNATUREPIECEG', TOBPiece_O.GetString('GP_NATUREPIECEG'), False), 1, 35);

  MotifPiece := TStringList.Create;
  if OKPASOK = 0 then
  begin
    MessOkPASOK := 'Validation de ' + RechDom('GCNATUREPIECEG', TOBPiece_O.GetString('GP_NATUREPIECEG'), False)+
                   'N° '+ TOBPiece_O.GetString('GP_NUMERO')+' --> '+
                   RechDom('GCNATUREPIECEG', TOBPiece.GetString('GP_NATUREPIECEG'), False)+' N° '+
                   TOBPiece.GetString('GP_NUMERO');
  end else
  begin
    MessOkPASOK := 'Errreur Validation de ' + RechDom('GCNATUREPIECEG', TOBPiece_O.GetString('GP_NATUREPIECEG'), False)+
                   ' N° '+ TOBPiece_O.GetString('GP_NUMERO') + ' --> '+MessEvent;
  end;
  MessEvent := MessOkPASOK;
  MotifPiece.Add(MessEvent);
  NumEvent := 0;
  QQ := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', True,-1, '', True);
  if not QQ.EOF then NumEvent := QQ.Fields[0].AsInteger;
  Ferme(QQ);
  Inc(NumEvent);
  QQ := OpenSQL('SELECT * FROM JNALEVENT WHERE GEV_TYPEEVENT="GEN" AND GEV_NUMEVENT=-1', False);
  QQ.Insert;
  InitNew(QQ);
  QQ.FindField('GEV_NUMEVENT').AsInteger := NumEvent;
  QQ.FindField('GEV_TYPEEVENT').AsString := 'SPI';
  QQ.FindField('GEV_LIBELLE').AsString := LeLibelle;
  QQ.FindField('GEV_DATEEVENT').AsDateTime := Date;
  QQ.FindField('GEV_UTILISATEUR').AsString := V_PGI.User;
  if OKPASOK = 0 then QQ.FindField('GEV_ETATEVENT').AsString := 'OK'
                 else QQ.FindField('GEV_ETATEVENT').AsString := 'ERR';
  TMemoField(QQ.FindField('GEV_BLOCNOTE')).Assign(MotifPiece);
  QQ.Post;
  Ferme(QQ);
  MotifPiece.Free;
end;


procedure T_genereFacture.DelOrUpdateAncien ;
begin
  if not UpdateAncien(TOBPiece_O, TOBPiece, TOBAcomptes_O, TobTiers, False, CleDoc.NumeroPiece) then V_PGI.IoError := oeSaisie;
  if InAvancement then
  begin
    NumSituation := SetSituationVivante(TOBPiece_O,'-'); 
  end;
end;


procedure T_genereFacture.InitToutModif(Datemaj : TdateTime = 0);
var NowFutur: TDateTime;
begin
	if DateMaj = 0 then
  begin
  	NowFutur := NowH;
  end else
  begin
  	NowFutur := DateMaj;
  end;
  TOBPiece.SetAllModifie(True);
  TOBPiece.SetDateModif(NowFutur);
  TOBAdresses.SetAllModifie(True);
  TOBBases.SetAllModifie(True);
  TOBEches.SetAllModifie(True);
  TOBAcomptes.SetAllModifie(True);
  TOBPorcs.SetAllModifie(True);
  InvalideModifTiersPiece(TOBTiers);
  TOBAnaP.SetAllModifie(True);
  TOBAnaS.SetAllModifie(True);
  TOBNomenclature.SetAllModifie(True);
  // Modif BTP
	TOBmetres.setAllModifie(true);
  TOBBasesL.SetAllModifie(true);
  TOBOuvrage.SetAllModifie(True);
  TOBOuvragesP.SetAllModifie(True);
  TOBLIENOLE.SetAllModifie(true);
  TOBPieceRG.SetAllModifie(true);
  TOBBasesRG.SetAllModifie(True);
  TOBTimbres.SetAllModifie(True);
  TOBSSTRAIT.SetAllModifie(true);
  TOBPieceTrait.setAllModifie(true);
  TOBVTECOLL.SetAllModifie(true);
end;


procedure T_genereFacture.ValideLaNumerotation;
var  OldNum, NewNum: integer;
  NaturePieceG: string;
begin
  NaturePieceG := TOBPiece.GetValue('GP_NATUREPIECEG');
  NewNum := cledoc.NumeroPiece;
  //if (NaturePieceG='TRE') or (NaturePieceG='TRV') then Exit;
  OldNum := TOBPiece.GetValue('GP_NUMERO');
  if not SetDefinitiveNumber(TOBPiece, TOBBases, TOBBasesL,TOBEches, TOBNomenclature, TOBAcomptes, TOBPIeceRG, TOBBasesRg, nil, NewNum) then
  begin
    MessageValid := 'Erreur lors de la renumérotation de la pièce';
    V_PGI.IoError := oePointage;
  end;
  //if Not SetNumeroDefinitif(TOBPiece,TOBBases,TOBEches,TOBNomenclature,TOBAcomptes) then V_PGI.IoError:=oePointage ;
  if GetInfoParPiece(NaturePieceG, 'GPP_ACTIONFINI') = 'ENR' then TOBPiece.PutValue('GP_VIVANTE', '-');
  CleDoc.NumeroPiece := NewNum;
  if ((OldNum <> NewNum)) then MajAccRegleDiff(TOBPiece, TOBAcomptes, OldNum);
end;

procedure T_genereFacture.ModifRefGC(TOBPiece, TOBAcomptes : TOB); // JT eQualité 10246
var Cpt : integer;
begin
  TobAcomptes.DelChampSup('CHGTNUM',false);
  for Cpt := 0 to TobAcomptes.detail.count -1 do
    DetruitAcompteGCCpta(TOBPiece,TobAcomptes.Detail[Cpt],True,False);
end;


procedure T_genereFacture.GestionRgUnique(TOBPIECE, TOBPIECERG, TOBBASES, TOBBASESRG, TOBTIers, TOBAffaire, TOBLigneRG: TOB);
var TOBL, TOBRG: TOB;
  Indice: integer;
begin
  if TOBPieceRG = nil then exit;
  if TOBPiece.getValue('GP_RGMULTIPLE')='X' then exit;
  if TOBPIeceRG.detail.count > 0 then
  begin
    Indice := 0;
    repeat
      TOBRG := TOBPieceRG.detail[Indice];
      if TOBRG.GEtVAlue('PRG_TAUXRG') = 0 then TOBRG.free else Inc(Indice);
    until Indice >= TOBPieceRG.detail.count;
  end;
  if IsMultiIntervRg (TOBpieceTrait) then
  begin
    TOBL := TOBPIece.findfirst(['GL_TYPELIGNE'], ['RG'], true);
    if Assigned(TOBL) then
    begin
      TOBL.free;
    end;
    InsereLigneRGCot(TOBPIece, TOBPIeceRG, TOBBASES, TOBBASESRG, TOBTIers, TOBAffaire, TOBLigneRG, -1);
  end else if not RGMultiple (TOBPiece) then
  begin
    TOBL := TOBPIece.findfirst(['GL_TYPELIGNE'], ['RG'], true);
    if Assigned(TOBL) then
    begin
      TOBL.free;
    end;
    InsereLigneRG(TOBPIece, TOBPIeceRG, TOBBASES, TOBBASESRG, TOBTIers, TOBAffaire, TOBLigneRG, -1);
  end;
end;

procedure T_genereFacture.GereLesReliquats; { NEWPIECE }
begin
  RazResteALivrerAndKillPiece(TobPiece_O);
end;

procedure T_genereFacture.ValideTiers;
begin
  ValideleTiers(TOBPiece, TOBTiers);
end;


procedure T_genereFacture.ValideLaPiece;
var OldEcr, OldStk: RMVT;
  NowFutur : TDateTime;
begin
  NowFutur := NowH;
  {Traitement pièce d'origine}
  if V_PGI.IoError = oeOk then DetruitCompta(TOBPiece_O, NowFutur, OldEcr, OldStk,true);
  if V_PGI.IoError = oeOk then DelOrUpdateAncien; { NEWPIECE }
  if V_PGI.IoError = oeOk then InverseStockTransfo(TOBPiece_O, TobPiece, TOBArticles, TOBCatalogu, TOBN_O);
  {Enregistrement nouvelle pièce}
  if V_PGI.IoError = oeOk then
  begin
    InitToutModif (NowFutur);
    ValideLaNumerotation;
    ValideLaCotation(TOBPiece, TOBBases, TOBEches);
    ValideLaPeriode(TOBPiece);
    if GetInfoParPiece(NewNature, 'GPP_IMPIMMEDIATE') = 'X' then TOBPiece.PutValue('GP_EDITEE', 'X');
  end;
  // JT eQualité 10246
  if (V_PGI.IoError = oeOk) and (TOBAcomptes.detail.count > 0) and (TOBAcomptes.FieldExists('CHGTNUM') )then
    ModifRefGC(TOBPiece, TOBAcomptes);
  // Fin JT
  // Modif Btp
  if V_PGI.IoError = oeOk then GestionRgUnique(TOBPIECE, TOBPIECERG, TOBBASES, TOBBASESRG, TOBTIers, TOBAffaire, TOBLigneRG);
  if V_PGI.IoError = oeOk then ValideLesArticlesFromOuv(TOBarticles,TOBOuvrage);
  if V_PGI.IoError = oeOk then ValideLesLignes(TOBPiece, TOBArticles, TOBCatalogu, TOBNomenclature, TobOuvrage, TOBPieceRG, TOBBasesRG, False, False,false,false,nil,false);
  if V_PGI.IoError = oeOk then ValideLesLignesCompl(TOBPiece, TobPiece_O);
  if V_PGI.IoError = oeOk then ValideLesOuv(TOBOuvrage, TOBPiece);
  if V_PGI.IoError = oeOk then ValideLesOuvPlat (TOBOuvragesP, TOBPiece);
  if V_PGI.IoError = oeOk then ValideLesBases(TOBPiece,TobBases,TOBBasesL);
  if (V_PGI.Ioerror = OeOk) and (not GetParamSocSecur('SO_METRESEXCEL',true)) then ValidelesMetres (TOBPiece,TOBmetres,TobOuvrage);
  if V_PGI.Ioerror = OeOk then ValideLesLiensOle(TOBPiece, TOBPiece_O, TOBLienOLE);
  if V_PGI.IoError = oeOk then ValideLesAdresses(TOBPiece, TOBPiece_O, TOBAdresses);
  if V_PGI.IoError = oeOk then
  begin
    GereLesReliquats;
  end;
  { NEWPIECE }
  if (HistoPiece(TobPiece_O)) then
  begin
    if not TOBPiece_O.UpdateDB(False) then
    BEGIN
      MessageValid := 'Erreur mise à jour pièce précédente (UPDATE)';
      V_PGI.IoError := oeSaisie; { NEWPIECE }
    end;
  end;
  if V_PGI.IoError = oeOk then ValideLesArticles(TOBPiece, TOBArticles);
  if V_PGI.IoError = oeOk then ValideLesCatalogues(TOBPiece, TOBCatalogu);
  if V_PGI.IoError = oeOk then ValideAnalytiques(TOBPiece, TOBAnaP, TOBAnaS);
  if V_PGI.IoError = oeOk then ValideTiers;
  if V_PGI.IoError = oeOk then ValideLaCompta(OldEcr, OldStk);
  if V_PGI.IoError = oeOk then ValideLesAcomptes(TOBPiece, TOBAcomptes);
  if V_PGI.IoError = oeOk then ValideLesPorcs(TOBPiece, TOBPorcs);
  if V_PGI.IoError = oeOk then ValideLesPieceTrait(TOBPiece, TOBAffaire,TOBPieceTrait,TOBSSTrait,DEV);
  if V_PGI.IoError = oeOk then ValideLesSousTrait(TOBPiece,TOBSSTrait,DEV);
  if V_PGI.IoError = oeOk then ValideLesRetenues(TOBPiece, TOBPieceRG);
  if V_PGI.IoError = oeOk then ValideLesBasesRG(TOBPiece, TOBBasesRG);
  if V_PGI.IoError = oeOk then ValideLesTimbres(TOBPiece, TOBTimbres);
  // --
  if (V_PGI.IoError = oeOk) and (TOBVTECOLL.detail.count > 0)  then
  begin
    PrepareInsertCollectif(TOBPiece,TOBVTECOLL);
    if not TOBVTECOLL.InsertDB(nil) then
    begin
      MessageValid := 'Erreur mise à jour TTC/COLLECTIFS';
      V_PGI.IoError := oeUnknown;
    end;
  end;

  if V_PGI.IoError = oeOk then
  begin
    if (GetInfoParPiece(NewNature, 'GPP_ACTIONFINI') = 'TRA') and
    	 (ToutesLignesSoldees(TOBpiece)) then
    begin
    	TOBPiece.PutValue('GP_VIVANTE','-');
    end;
    //
    if not TOBPiece.InsertDBByNivel(False) then
    begin
      MessageValid := 'Erreur mise à jour LIGNE';
      V_PGI.IoError := oeUnknown;
    end;
  end;

  if V_PGI.IoError = oeOk then
  begin
    if not TOBOuvragesP.InsertDBByNivel(false) then
    begin
      MessageValid := 'Erreur mise à jour OUVRAGES PLAT';
      V_PGI.IoError := oeUnknown;
    end;
  end;
  if V_PGI.IoError = oeOk then
  begin
    if not TOBBases.InsertDB(nil) then
    begin
      MessageValid := 'Erreur mise à jour BASES';
      V_PGI.IoError := oeUnknown;
    end;
  end;
  if V_PGI.IoError = oeOk then
  begin
    if not TOBBasesL.InsertDB(nil) then
    begin
      MessageValid := 'Erreur mise à jour BASES LIGNE';
      V_PGI.IoError := oeUnknown;
    end;
  end;
  if V_PGI.IoError = oeOk then
  begin
    if not TOBEches.InsertDB(nil) then
    begin
      MessageValid := 'Erreur mise à jour ECHEANCES';
      V_PGI.IoError := oeUnknown;
    end;
  end;
  if V_PGI.IoError = oeOk then
  begin
    if not TOBAnaP.InsertDB(nil) then
    begin
      MessageValid := 'Erreur mise à jour ANALYTIQUE/PIECE';
      V_PGI.IoError := oeUnknown;
    end;
  end;
  if V_PGI.IoError = oeOk then
  begin
    if not TOBAnaS.InsertDB(nil) then
    begin
      MessageValid := 'Erreur mise à jour ANALYTIQUE/STOCK';
      V_PGI.IoError := oeUnknown;
    end;
  end;
  if (V_PGI.Ioerror = OeOk) and (not GetParamSocSecur('SO_METRESEXCEL',true)) then
  begin
  	if not TOBmetres.InsertDBByNivel(false) then
    begin
      MessageValid := 'Erreur mise à jour METRES';
      V_PGI.IoError := oeUnknown;
    end;
  end;
  if V_PGI.IoError = oeOk then ValideLesNomen(TOBNomenclature);
  if V_PGI.IoError = oeOk then ValideSituation(TOBPiece,TOBOuvrage);

  if (V_PGI.IoError = oeOk) and (InAvancement) then
  begin
    MajApresGeneration (TOBPiece,TOBPieceRG,TOBBasesRG,TOBAcomptes,DEV,NumSituation);
  end;
  if (V_PGI.IOError = OeOk) Then TheRepartTva.Ecrit;
end;

function T_genereFacture.GetMessage: string;
begin
  result :=MessageValid; 
end;

end.

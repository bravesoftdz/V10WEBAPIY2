{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 26/03/2008
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSAISIEAVANC ()
Mots clefs ... : TOF;BTSAISIEAVANC
*****************************************************************}
Unit BTSAISIEAVANC_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Menus,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     MainEagl,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     Hpanel,
     UTofAfBaseCodeAffaire,
     CalcOleGenericBTP,
     EntGC,
     AglInit,
     graphics,
     windows,
     grids,
     ImgList,
     ExtCtrls,
     Messages,
     SaisUtil,
     Splash,
     UTOF,
     vierge,
     BTCHOIXAPPLICAVAN_TOF,
     BTStructChampSup,uEntCommun,UtilTobPiece ;

Type
  TReference = class (Tobject)
  	private
    	fDoc : string;
    	fTOB : TOB;
      fNombre : integer;
    public
    	constructor create;
      destructor destroy; override;
      property TOBAssocie : TOB read fTOb write ftob;
      property nombre : integer read fNombre write fNombre;
      procedure Adddoc (TOBL: TOB);
      procedure IncNombre;
  end;

  Tlistref = class (Tlist)
  	private
      function Add(AObject: Treference): Integer;
      function GetItems(Indice: integer): TReference;
      procedure SetItems(Indice: integer; const Value: TReference);
    public
    	destructor destroy; override;
      property Items [Indice : integer] : TReference read GetItems write SetItems;
      function find (TOBL : TOB): Treference;
  end;

	TlisteTOB = class (Tlist)
  	private
      function Add(AObject: TOB): Integer;
      function GetItems(Indice: integer): TOB;
      procedure SetItems(Indice: integer; const Value: TOB);
    public
    	destructor destroy; override;
      property Items [Indice : integer] : TOB read GetItems write SetItems;
  end;

	TOnglet = class (Tobject)
  	private
    	fGeneral : boolean;
    	fname : string;
      fTVDevis : TTreeview;
      fTOBlignes : TOB;
      ffournisseur : string;
      fGS : Thgrid;
      fstPrev : string;
      fTabSheet : TTabSheet;
      fMtSituation : double;
      fSousTraitant : Boolean;
      fBOUVRIR : TToolbarButton97;
    public
    	constructor create (XX : TOF; Generale : boolean);
      destructor destroy; override;
      property TOBlignes : TOB read fTOBlignes write fTOBlignes;
      property GS : Thgrid read fGS write fGS;
      property Name : string read fName write fName;
      property fournisseur : string read ffournisseur write ffournisseur;
      property SousTraitant : boolean read fSousTraitant write fSousTraitant;
      property stPrev : string read fstPrev write fstPrev;
      property TabSheet : TTabSheet read ftabSheet write fTabSheet;
      property MtSituation : double read fMtSituation write fMtSituation;
  end;

  TlistOnglet = class (Tlist)
  	private
      ffreeAll : Boolean;
      function Add(AObject: TOnglet): Integer;
      function GetItems(Indice: integer): TOnglet;
      procedure SetItems(Indice: integer; const Value: Tonglet);
    public
      constructor create;
    	destructor destroy; override;
      property Items [Indice : integer] : TOnglet read GetItems write SetItems;
      property ClearWhenFree : Boolean read ffreeAll write ffreeAll;
      function findOnglet (NomOnglet : string ): TOnglet;
      procedure clear; override;
  end;

  TOF_BTSAISIEAVANC = Class (TOF_AFBASECODEAFFAIRE)
      procedure OnNew                    ; override ;
      procedure OnDelete                 ; override ;
      procedure OnUpdate                 ; override ;
      procedure OnLoad                   ; override ;
      procedure OnArgument (S : String ) ; override ;
      procedure OnDisplay                ; override ;
      procedure OnClose                  ; override ;
      procedure OnCancel                 ; override ;
      procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;
   private
      MontantMarche : double;    // Montant du marche global
      DejaFacture : double;      // Deja facture sur les devis sélectionné
      MtHorsDevis : double;
   		TOBEcart : TOB;
    	ArtEcart : string;
      ModifSousDetail : boolean;
      ReajusteDossier : boolean;
      ModifSituation : boolean;
      MaxNumOrdre : Integer;
      // Modified by f.vautrain 07/11/2017 15:53:27 - FS#2774 - GUINIER - Interdire la possibilité de facturer supérieur à 100 %
      {$IFDEF V10}
      CtrlSupCent : Boolean;
      {$ENDIF}
//      TOBPieceTrait : TOB;
      //
      fOnglets : TlistOnglet;
      fCurrentOnglet : TOnglet;
      FOngletGlobal  : TOnglet;
      flastfacture : boolean;
    	TOBOuvrages : TOB;
      TOBSOUSTraits : TOB;
      TOBInfoPieces : TOB;
      //
      ModeCloture : boolean;
      Avancement : boolean;
      ModeGeneration : string;
      NextPiece : string;
      ImTypeArticle,ImgListPlusMoins,ImgTypeTrait : TImageList;
      DEV: RDEVISE;
      VOIRQTE,VOIRMT : TCheckBox;
      POPACTIONS : TPopupMenu;
      MnAvancement,MnDefaireLig,MnDefairePiece,MnDefaireTous,MnDefaireParag : TMenuItem;
      //
      INDICE,TYPEARTICLE,CODEARTICLE,LIBELLE,QTEMARCHE,UNITE,VOIRDETAIL,CODEMARCHE : integer;
      QTEDEJAFACT,QTECUMULFACT,POURCENTAVANC,MTCUMULEFACT,MTDEJAFACT,MTSITUATION,MTMARCHE,QTESITUATION,PUSIT,MTPRODUCTION : integer;
      LNQTEDEJAFACT,LNQTECUMULFACT,LNQTESITUATION,LNQTEMARCHE,LNMTMARCHE,LNMTCUMULEFACT,LNMTDEJAFACT,LNUNITE,LPUSIT,LMTPRODUCTION : integer;
      dMtSItuation : double;
      FMTSITUATION : THNumEdit;
      //
      fCotraitance : boolean;
      fSoustraitance : boolean;
      //
      Action : TActionFiche;
      ListeChamps,stPrev : string;
      NbCols : integer;
      ToolWin : TToolWindow97;
      GS : THgrid;
      //
      PS : TPageControl;
      //
      BPF_MTSITUATION : THLabel;
      AFF_TIERS,AFF_AFFAIRE0,AFF_AFFAIRE1,AFF_AFFAIRE2,AFF_AFFAIRE3,AFF_AVENANT,AFF_AFFAIRE : THedit;
      AFF_DESIGNATION,LIBTIERS : THLABEL;
      BTDescriptif,Bferme,Bvalider : TToolbarButton97;
      PTVDevis : THPanel;
      TVDevis : TTreeview;
      LAFFAIRE : string;
      TOBSELECT,TOBLIGNES,LESTOBACOMPTES,TOBAcomptes_O,TOBPORCS : TOB;
      TOBTiers,TOBaffaire : TOB;
      TDEVISE : THlabel;
      BAVANC : TToolbarButton97;
      ErreurEcr : string;
      //
      procedure AddChampsSupAvanc (TOBL : TOB);
      procedure GetComponents;
      procedure CreateComposants;
      procedure DefinieLaTOB;
      procedure DefinieGrille (GS : Thgrid);
      procedure AddLeDocument(TOBS : TOB);
      procedure PositionneLigne (TOBL : TOB; Ligne : integer; EnAchat :Boolean=false);
      procedure DefiniAttributCol (GS : THgrid; Nom: string; Colonne: integer; LaLargeur,Lalignement,Letitre,LeNC : string);
      procedure AfficheInfos;
      procedure AfficheLaGrille (GS : THgrid; TOBlignes : TOB; Depart : integer =0) ;
      procedure AfficheLaLigne(GS : THgrid; TOBlignes: TOB; Ligne: integer);
      procedure SetEvents (GS : THgrid);
      procedure SetEventsGrid(GS : THgrid; Etat: boolean);
      function  GetTOBLigneAV (TOBlignes: TOB; Arow : integer) : TOB;
      procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
      procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer;  var Cancel: Boolean);
      procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
      procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
      procedure GSPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
      procedure GSGetCellCanvas(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
      procedure GSColumnsWithChanging (Sender : Tobject);
      function  RecupTypeGraph (TOBL : TOB) : integer;
      procedure ChargeListImage;
      procedure FreeComposants;
      procedure ChangeQteSituation (TOBL : TOB; CurrOnglet : Tonglet;  var Cancel : boolean);
      procedure ChangePourcentAvanc (TOBL : TOB ; CurrOnglet : Tonglet;  var Cancel : boolean);
      procedure ChangeMTSituationAvanc (Montant : Double; TOBL : TOB ; CurrOnglet : Tonglet; var Cancel : boolean);
      procedure ZoneSuivanteOuOk(Grille: THgrid; var ACol, ARow: integer;var Cancel: boolean);
      function  ZoneAccessible(Grille: THgrid; var ACol,ARow: integer): boolean;
      procedure CalculeLaLigneFromPourcent (TOBL : TOB;CurrOnglet : Tonglet);
      procedure CalculeLaLigneFromQteSit  (TOBL : TOB;CurrOnglet : Tonglet);
      procedure CalculeLaLigneFromMontant (Montant : Double; TOBL : TOB;CurrOnglet : Tonglet);
      procedure CumuleLigne(TOBL : TOB; Sens : string='+');
      procedure GetInfoDev (CodeDevise : string);
      procedure InitValeurZero(TOBL : TOB);
      procedure SetLigne(GS : Thgrid; Arow:integer;Acol : integer=-1);
      procedure ChangeMontantSituation (Sender : TOBject);
      procedure BAVANCClick (Sender : Tobject);
      function  RecupDebutParagraph(Ligne: integer) : integer;
      function  RecupFinParagraph(Ligne : integer): integer;
      function  RecupDebutPieceCur(Ligne : integer) : integer;
      function  RecupFinPieceCur(Ligne : integer) : integer;
      procedure AppliquePourcentAvanc(Pourcent: double;IndDepart, IndFin : integer);
      function  IsDebutPiece (TOBL : TOB; cledoc : r_cledoc) : boolean;
      function  IsPiece (TOBL : TOB) : boolean;
      function  IsSamePiece(TOBL: TOB; cledoc: r_cledoc): boolean;
      function  IsGerableAvanc (TOBL : TOB) : boolean;
      procedure CalculelesSousTotaux;
      procedure SommerLignesAV(ARow, Niv: integer);
      procedure VoirMtCLick (Sender : Tobject);
      procedure VoirQteCLick (Sender : Tobject);
      function  ChangeQteSituationOk (Ligne: integer; valeur : double) : boolean;
      procedure AddChampsSupMtINITAvanc(TOBL: TOB);
      function  ChangePourcentOk(Ligne: integer; valeur: double): boolean;
      function  ChangeMTSituationok(Ligne: integer; Valeur: double): boolean;
      procedure Constitue_TVDEVIS;
      procedure BTDescriptifClick (Sender : TObject);
      procedure TVDEVISClick(Sender: TObject);
      procedure ToolWinClose(Sender: TObject);
      procedure CalculeSousTotal(Ligne, IChampST: integer);
      procedure SetInfoFac(TOBL: TOB);
      procedure AddChampsFac;
      procedure GenereLafacture (DateFac : TdateTime ; Var Clefac : R_Cledoc; GenereAvoir : boolean=false);
      function  ExisteLigneSituation : boolean;
      procedure FormKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
      procedure DefaireParag (Sender : TObject);
      procedure DefaireLig (Sender : TObject);
      procedure defaireLigOuvSaisie (TOBO : TOB);
      procedure DefairePiece(Sender : TObject);
      procedure DefaireTous(Sender : TObject);
      procedure ReInitLigne(TOBL: TOB);
      procedure DefaireSegment(Debut, Fin: integer);
      procedure RecupAcomptePiece (TOBS : TOB);
      procedure InitChampsSupEntete(TOBS : TOB);
      procedure DefiniAvancementPrecPiece(TOBL : TOB);
      procedure SetAvancementPiece(TOBL: TOB; Sens: string='+');
      procedure SetAvancementAcompte;
      function  ChangePourcentOkMulti(Ligne: integer;valeur: double; var LibelleErreur : string): integer;
    	function  TraitementDerniereSituation : boolean;
    	function  AjouteLigneEcart(MtEcart: double): boolean;
    	procedure SetInfoFactEcart(TOBL: TOB);
      procedure ClotureFacturation;
      procedure GetPorts (TOBpiece : TOB);
    	function  DemandeDatesFacturation(var DateFac: TDateTime): boolean;
    	procedure ConstituelesOnglets;
      procedure CalculeTotauxSaisie;
      procedure ConstitueOngletInit;
      procedure AddLignefacturationFromOuv(TOBL,TOBLignes,TOBO: TOB; Indice, decalage,IndiceNomen: integer; var montant : double;AvancSousDetailPartiel:boolean; EnAchat : Boolean=false);
      procedure ChangeVue (Sender : Tobject);
      procedure BeforeChangeVue (Sender: TObject; var AllowChange: Boolean);
      procedure AppliqueVoirQte;
      procedure AppliqueVoirMt;
      procedure SetOnglet(Onglet: Tonglet);
      function  FindTOBligne(Onglet: Tonglet; NumPiece,numOrdre,uniqueBlo: integer): integer;
      function  FindlignesOuvSaisie(Onglet: Tonglet; TOBL: TOB): TlisteTOB;
    	function  FindLigneInOnglet(NumPiece,Numordre, uniqueBlo: integer;var ligne: integer): Tonglet;
      procedure FormResize (Sender : Tobject);
      procedure AppliqueAvancGeneral(TOBD: TOB; Avanc: double;WithDeduction:boolean=true;WithAjout: Boolean=True;WithCalcul:boolean=true);
      procedure AppliqueAvancFournisseur(TOBD: TOB; Avanc: double;WithDeduction:boolean=true;WithAjout:Boolean=True;WithCalcul:boolean=true);
      procedure AppliqueQteGeneral(TOBD: TOB; qte: double);
      procedure AppliqueQteFournisseur(TOBD: TOB; Qte: double);
      procedure AppliqueMtFournisseur(TOBD: TOB; Mt: double);
      procedure AppliqueMtGeneral(TOBD: TOB; Mt: double);
      procedure ReInitLigneGlobal(TOBD: TOB);
      function  ReInitLigneFournisseur(TOBD: TOB): Tonglet;
      procedure ReInitLigneOnglets (TOBD : TOB);

      procedure ReCalculelesOnglets(TheListe: TstringList);
      procedure ReCalculeglobal;
      procedure RecalculeOnglet(unOnglet: Tonglet; WithSaveBefore: boolean=false);
      procedure CumuleLigneGlobal(TOBD : TOB; SenS:string= '+');
      function  CumuleLigneFournisseur(TOBD: TOB; Sens:string='+'): Tonglet;
      procedure CumuleLigneOuv(TOBD: TOB; Sens: string= '+');
    	procedure AppliqueAvancSousDetail(TOBO: TOB; CurrOnglet : Tonglet;WithDeduction:Boolean=True;WithAjout:Boolean=True;WithCalcul:Boolean=true);
    	procedure AppliqueQteOuvSousDetail(TOBO: TOB; CurrOnglet : Tonglet);
      procedure EnlevepasConcerne(TOBlignes: TOB; TheFournisseur: string; SousTraitancePA : Boolean=false);
      function IsDetailInside(TOBLignes: TOB; IndDep: integer; TOBL: TOB): boolean;
      procedure Nettoieparagraphes(TOBLignes: TOB);
      procedure SuprimeThisParagraphe(TOBLignes, TOBL: TOB;var IndDep: integer);
    	function FindNextLigne(TOBLigne, TOBO: TOB; Indice: integer): integer;
    	procedure OrganiseOuvrage(TOBPiece, TOBNomen: TOB; TOBLL : TOB=nil);
    	procedure SetAvancSousDetail(TOBL: TOB; Onglet : Tonglet);
//    	procedure PositionneAvancOuv(TOBL: TOB);
      procedure SetAvancementDetailOuvrage(TOBgroupeNomen, TOBO, TOBL: TOB);
      function OnlySauvegarde: boolean;
      function EnregistreAvancement: boolean;
      procedure InitMontantSupfacturation(TOBL: TOB);
    	procedure ReinitLigneSD(TOBL : TOB);
    	procedure AppliqueAvancSDetailFournisseur(TOBD: TOB; Avanc: double;WithDeduction, WithAjout,WithCalcul: boolean);
//      procedure SaisieReglClick (Sender : TObject);
      function IsOuvragePartiel (TOBL : TOB) : boolean;
    	procedure AppliqueMtDetailGeneral(TOBD: TOB);
      procedure EnregistrelesSousTraits (TOBPiece : TOB);
      procedure ShowButton (Onglet : TOnglet ; Ligne : Integer; Visible : Boolean);
    	procedure CreateButtonOuvrir(Fonglet: TOnglet; indice: integer);
      procedure OuvrirClick (Sender : TObject);
      procedure AfficheSousDetail (Onglet : TOnglet; TOBLignes,TOBL : TOB; var Indice : integer; EnAchat : boolean=false);
      procedure SupprimeSousDetail(TOBLignes: TOB; Indice: integer);
      procedure FindLesOnglets(NumPiece, Numordre, uniqueBlo: integer;
      var ligne: integer; TheRef: TOnglet; LaListe: TstringList);
      procedure OuvrirDansOnglet(Onglet: TOnglet; ligne: integer);
      procedure SetAvancOuvrage(TOBL: TOB);
      procedure AppliqueAvancDetailGeneral(TOBD: TOB; Avanc: double; WithDeduction, WithAjout,WithCalcul: boolean);
      procedure AppliqueQteDetailGeneral(TOBD: TOB);
      procedure CumuleAvancDetailGeneral(TOBD: TOB; Sens: string);
      procedure NettoieOuvrageOrphelin(TOBlignes: TOB);
    	function IsAvancementDetailOuvrage (TOBL :TOB) : Boolean;
      procedure ReindiceOuvrages;
      //
      procedure ChangeQteFact (TOBL: TOB; CurrOnglet : Tonglet;  var Cancel: boolean);
      procedure CalculeLaLigneFromQteFact(TOBL: TOB;CurrOnglet : Tonglet );
      procedure AppliqueQteFacOuvSousDetail(TOBO : TOB; CurrOnglet : Tonglet);
      procedure AppliqueQteFact(TOBD: TOB; qte: double);
      procedure AppliqueQteFactFour(TOBD: TOB; Qte: double);
      function EnregistreDevis: Boolean;
      procedure RecalcMontantSupfacturation(TOBL: TOB);
      procedure REcalcMtSupfactOuv(TOBOUV: TOB);
      procedure MemoriseHighNumOrdre (TOBL: TOB);
      procedure SynchroniseDocumentWithDevis;
      procedure AppliqueMtSousDetail(Coef: Double; TOBO: TOB; CurrOnglet: Tonglet);
      procedure SetMtOuvrage(TOBL: TOB; Coef: double);
      function TraitementGestionAvoir : boolean ;
  end ;

Implementation
uses FactTOb,FactUtil,FactVariante,UTofSaisieAvanc,FactComm,FactOuvrage,FactGrp,BTGENFACTURE,FActCommBTP,FactureBTP,Factspec,
		 UtilArticle,Paramsoc,Factcalc,FactRG, uTOFComm,UCotraitance,Facture,UFonctionsCBP;

procedure TOF_BTSAISIEAVANC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEAVANC.OnDelete ;
begin
  Inherited ;
end ;

function TOF_BTSAISIEAVANC.TraitementDerniereSituation : boolean;
var indice : integer;
    TOBL : TOB;
    Atraiter : boolean;
    MtSituation,MtSItuationNet,Decart,Rempied,Escompte,MtPorcs,MtPorts : double;
    QQ : TQuery;
    Req : string;
begin
	result := true;
	MontantMarche := 0;
  MtSItuationNet := 0;
  DejaFacture := 0;
  MtHorsDevis := 0;
  MtPorcs := 0;
  Atraiter := true;
  flastfacture := false;
  //
  result := true;
	if not Avancement then exit;
  flastfacture := true;
  // ----------------------------------------------------
  for Indice := 0 to TOBSELECT.detail.count -1 do
  begin
    MontantMarche := MontantMarche + TOBSelect.detail[Indice].getValue('GP_TOTALHTDEV');
  end;
  For indice := 0 to TOBPORCS.detail.count -1 do
  begin
    if (TOBPorcs.detail[Indice].getValue('GPT_FRAISREPARTIS')<>'X') and (TOBPorcs.detail[Indice].getValue('GPT_RETENUEDIVERSE')<>'X') then
    begin
      MontantMarche := MontantMarche - TOBPorcs.detail[Indice].getValue('GPT_TOTALHT');
    end;
  end;
  MontantMarche := Arrondi(MontantMarche,DEV.decimale);
  // ----------------------------------------------------
  for Indice := 0 to TOBLignes.detail.count - 1 do
  begin
    if TOBLignes.detail[Indice].getValue('GL_TYPELIGNE')<>'ART' then continue;
    if (TOBLignes.detail[Indice].getValue('BLF_POURCENTAVANC')<>100) then
    begin
    	Atraiter := false;
      flastfacture := false;
      break;
    end;
  end;

  if Atraiter then
  begin
  	TOBL := TOBSELECT.detail[0]; // devis principal
    // Dans le cas ou l'on cloture-facture un devis en memoire final
    // il ne faut pas récupérer les éléments des situations précédentes
    // qui sont des demandes d'acomptes
    if ((ModeCloture) and (ModeGeneration<>'DAC')) or (not Modecloture) then
    begin
      // récupération du facture Hors devis initial (ajout sur facture)
      Req := 'SELECT BLF_MTSITUATION AS HORSDEVIS '+
                     'FROM BSITUATIONS '+
                     'LEFT JOIN LIGNEFAC ON '+
                     'BLF_NATUREPIECEG=BST_NATUREPIECE AND '+
                     'BLF_SOUCHE=BST_SOUCHE AND '+
                     'BLF_NUMERO=BST_NUMEROFAC '+
                     'WHERE BST_SSAFFAIRE="'+TOBL.getValue('GP_AFFAIREDEVIS')+'" AND BST_VIVANTE="X" AND '+
                     'BLF_UNIQUEBLO=0 AND BLF_MTMARCHE=0 AND BLF_MTSITUATION <>0';
      QQ := openSql (req,true,-1,'',true);
      if not QQ.eof then
      begin
        QQ.First;
        while not QQ.eof do
        begin
          MtHorsDevis := MtHorsDevis + QQ.findField('HORSDEVIS').asFloat;
          QQ.Next;
        end;
      end;
      ferme (QQ);
      //
      QQ := openSql ('SELECT SUM(BST_MONTANTHT) AS DEJAFACT '+
      							 'FROM BSITUATIONS '+
                     'WHERE '+
                     'BST_SSAFFAIRE="'+TOBL.getValue('GP_AFFAIREDEVIS')+'" AND '+
                     'BST_NATUREPIECE<>"B00" AND '+
                     'BST_VIVANTE="X"',true,-1,'',true);
      if not QQ.eof then
      begin
        DejaFacture := QQ.findField('DEJAFACT').AsFloat;
      end;
      ferme (QQ);
      // Récupération des ports facturés
      Req :='SELECT GPT_RETENUEDIVERSE,GPT_FRAISREPARTIS,GPT_TOTALHT FROM BSITUATIONS LEFT JOIN PIEDPORT ON '+
      			'GPT_NATUREPIECEG=BST_NATUREPIECE AND GPT_SOUCHE=BST_SOUCHE AND GPT_NUMERO=BST_NUMEROFAC '+
            'WHERE BST_SSAFFAIRE="'+TOBL.getValue('GP_AFFAIREDEVIS')+'" AND BST_VIVANTE="X"';
      QQ := openSql (Req,true,-1,'',true);
      if not QQ.eof then
      begin
        QQ.First;
        while not QQ.eof do
        begin
        	if (QQ.findField('GPT_FRAISREPARTIS').AsString <>'X') and (QQ.findField('GPT_RETENUEDIVERSE').AsString <>'X') then
          begin
          	MtPorcs := mtPorcs + QQ.findField('GPT_TOTALHT').asFloat;
          end;
          QQ.Next;
        end;
      end;
      ferme (QQ);
      Dejafacture := Dejafacture - MtPorcs;
    end;
    //
//    MtSituationNet := Arrondi(MtSItuationNet,DEV.decimale);
    (*
  	RecalculPiedPort (taModif,TOBSELECT.detail[0],TOBPorcs,MtSituationNet);
    MtPorts := CalculPort(True,TOBporcs);
    //
    if Arrondi(MontantMarche-(MtSituationNet+MtPorts+DejaFacture),Dev.decimale)<> 0 then
    begin
      Decart := arrondi(MontantMarche - (MtSituationNet+DejaFacture),2);
      if not AjouteLigneEcart (Decart) then result := false;
    end;
    *)
  end;
end;

function TOF_BTSAISIEAVANC.OnlySauvegarde : boolean ;
var OneTOB : TOB;
begin
	if (ModeCloture) or (ModifSituation) then
  begin
  	result := false;
    exit;
  end;
  OneTob := TOB.Create('UNE TOB',nil,-1);
  TRY
    OneTOB.AddChampSupValeur('RETOUR',1); // par défaut génération de la facture
    TheTOb := OneTOB;
    AGLLanceFiche('BTP','BTCHOIXFINSAISIE','','','');
    result := (OneTOB.getValue('RETOUR')=0);
  FINALLY
  	OneTOB.free;
  END;
end;

procedure TOF_BTSAISIEAVANC.OnUpdate ;
var IsModified : boolean;
		DateFac : TdateTime;
    Clefac : r_cledoc;
    GenereAvoir : boolean;
begin
  Inherited ;
  ErreurEcr := '';
  FillChar (Clefac,sizeof(CleFac),#0);
  TToolWindow97(GetControl('PBouton')).Enabled := false;
  isModified := ExisteLigneSituation;
  //
  if ReajusteDossier then
  begin
  	if not EnregistreDevis then PGIError('Erreur durant la sauvegarde de la saisie',ecran.caption);
  end else if OnlySauvegarde then
  begin
  	if not EnregistreAvancement then PGIError('Erreur durant la sauvegarde de la saisie',ecran.caption);
  end else
  begin

    if ModeCloture then
    begin
      if (ModeGeneration = 'AVA') or (ModeGeneration='DIR') then    // gestion des cloture de factures a l'avancement ou directe
      begin
        if IsModified then
        begin
          // on demande si l'utilisateur veux bien générer une nouvelle facturation pour ce dossier de facturation
          if PGIAsk ('ATTENTION : Vous avez modifié l''avancement.#13#10 Confirmez-vous la génération d''une facture pour ce dossier ?')=Mryes then
          begin
            if TraitementDerniereSituation then
            begin
              TRY
                if not DemandeDatesFacturation (DateFac) then
                begin
                  ecran.ModalResult := 0;
                  TToolWindow97(GetControl('PBouton')).Enabled := true;
                  ErreurEcr := ErreurEcr+'#13#10 la date de facturation est obligatoire';
                  V_PGI.Ioerror := oeUnknown;
                  exit;
                end;
                GenereLafacture(DateFac,Clefac);
              FINALLY
                if V_PGI.IOError = OeOK then
                begin
                  ClotureFacturation;
                  ErreurEcr := ErreurEcr+'#13#10 Le dossier de facturation à été cloturé avec succès';
                  PgiInfo(ErreurEcr);
                  if (CleFac.NumeroPiece <> 0) and (PgiAsk('Désirez-vous ouvrir la facture générée ?')=Mryes) then SaisiePiece(Clefac,taModif);
                end else  PGIError (ErreurEcr);
              END;
            end;
          end else
          begin
            ecran.ModalResult := 0;
            TToolWindow97(GetControl('PBouton')).Enabled := true;
          end;
        end else
        begin
          ClotureFacturation;
          ErreurEcr := 'Le dossier de facturation à été cloturé avec succès';
          PgiInfo(ErreurEcr);
        end;
      end else if ModeGeneration = 'DAC' then // gestion des cloture de demande d'acompte --> mémoire final
      begin
        if IsModified then
        begin
          if TraitementDerniereSituation then
          begin
            TRY
              if not DemandeDatesFacturation (DateFac) then
              begin
                ecran.ModalResult := 0;
                TToolWindow97(GetControl('PBouton')).Enabled := true;
                ErreurEcr := ErreurEcr+'#13#10 la date de facturation est obligatoire';
                V_PGI.Ioerror := oeUnknown;
                exit;
              end;
              GenereLafacture(DateFac,Clefac);
            FINALLY
              if V_PGI.IOError = OeOK then
              begin
                  ClotureFacturation;
                  ErreurEcr := ErreurEcr+'#13#10 Le dossier de facturation à été cloturé avec succès';
                  PgiInfo(ErreurEcr);
                if (CleFac.NumeroPiece <> 0) and (PgiAsk('Désirez-vous ouvrir la facture générée ?')=Mryes) then SaisiePiece(Clefac,taModif);
              end else  PGIError (ErreurEcr);
            END;
          end;
        end else
        begin
          PgiInfo('Rien à facturer pour l''instant');
          ecran.ModalResult := 0;
          TToolWindow97(GetControl('PBouton')).Enabled := true;
        end;
      end;
    end else
    begin
      if isModified then
      begin
        if TraitementDerniereSituation then
        begin
          TRY
            GenereAvoir := false;
            if not DemandeDatesFacturation (DateFac) then
            begin
              ecran.ModalResult := 0;
              TToolWindow97(GetControl('PBouton')).Enabled := true;
              ErreurEcr := ErreurEcr+'#13#10 la date de facturation est obligatoire';
              V_PGI.Ioerror := oeUnknown;
              exit;
            end;
            //
            // Controle si le montant du document est encore positif sinon proposition de générer un avoir a la place de la facture
            //
            // ----------------------------------------
            GenereAvoir := TraitementGestionAvoir;
            // ----------------------------------------
            GenereLafacture(DateFac,Clefac,GenereAvoir);
          FINALLY
            if V_PGI.IOError = OeOK then
            Begin
              PgiInfo(ErreurEcr);
              if (CleFac.NumeroPiece <> 0) and (PgiAsk('Désirez-vous ouvrir le document généré ?')=Mryes) then SaisiePiece(Clefac,taModif);
            end
                                    else  PGIError (ErreurEcr);
          END;
        end;
      end else
      begin
        PgiInfo('Rien à facturer pour l''instant');
        ecran.ModalResult := 0;
        TToolWindow97(GetControl('PBouton')).Enabled := true;
      end;
    end;
  end;
end ;

procedure TOF_BTSAISIEAVANC.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEAVANC.OnArgument (S : String ) ;
var QQ : Tquery;
begin
  Inherited ;
  ReajusteDossier := false;
  ModifSituation := false;
  modifSousDetail := GetParamSocSecur('SO_BTMODIFSDETAIL',false);
  fCotraitance := false;
  fSoustraitance := false;
  ArtEcart := trim(GetParamsoc('SO_BTECARTPMA'));
  // ----
  ModeCloture := false;
  if Pos('CLOTURE',S)>0 then ModeCloture := true;
  if Pos('REAJUSTEDOSSIER',S)>0 then ReajusteDossier := True;
  if Pos('MODIFSITUATION',S)>0 then ModifSituation := True;
  // ---
  Action := taModif;
  TOBSELECT :=LaTOB;
  ModeGeneration := TOBSELECT.GetValue('AFF_GENERAUTO');
  Avancement := (Pos(ModeGeneration,'DAC;AVA')>0);
  if Not avancement then THLabel(GetCOntrol('TMTSITUATION')).caption := 'Montant de la facture';
  
  if ReajusteDossier then
  begin
    Ecran.Caption := 'Réajustement dossier de facturation';
    UpdateCaption(ecran);
    SetControlVisible('TMTSITUATION',false);
    SetControlVisible('BPF_MTSITUATION',false);
    SetControlVisible('TDEVISE',false);
    SetControlVisible('BAVANC',false);
    SetControlVisible('BRECUPAVANC',true);
  end;
  //
  //
  if ModeGeneration='DAC' then
  begin
  	if ModeCloture then
    begin
	  	NextPiece := 'FBT'
    end else
    begin
	  	NextPiece := 'DAC'
    end;
  end else
  begin
    if ModifSituation then
    begin
      NextPiece := 'FBP';
    end else
    begin
      NextPiece := 'FBT';
    end;
  end;
  //
  MemoriseChampsSupLigneETL (NextPiece ,True);
  MemoriseChampsSupLigneOUV (NextPiece);
  MemoriseChampsSupPIECETRAIT;
  //
  // Modified by f.vautrain 07/11/2017 15:50:28 - FS#2774 - GUINIER - Interdire la possibilité de facturer supérieur à 100 %
  {$IFDEF V10}
  CtrlSupCent     := GetParamSocSecur('SO_AVCSUPCENT', False);
  {$ENDIF}
  //
  TOBLIGNES := TOB.Create ('PIECE',nil,-1);
  TOBTiers := TOB.Create ('TIERS',nil,-1);
  TOBECART := TOB.create ('ARTICLE',nil,-1);
  TOBOuvrages := TOB.Create ('LES DETAILS OUVRAGES',nil,-1);
  TOBaffaire := TOB.Create ('AFFAIRE',nil,-1);
  LESTOBACOMPTES := TOB.create ('LES ACOMPTES',nil,-1);
  TOBAcomptes_O := TOB.create ('LES ACOMPTES',nil,-1);
  TOBPORCS := TOB.create ('LES PORTS',nil,-1);
  TOBInfoPieces := TOB.Create ('LES DOCUMENTS',nil,-1);
  //
	TOBSOUSTraits := TOB.Create ('LES SOUSTRAITS',nil,-1);
  //
  QQ := OpenSql ('SELECT GA_CODEARTICLE,GA_LIBELLE FROM ARTICLE WHERE GA_CODEARTICLE="'+ArtEcart+'"',true,1,'',true);
  if not QQ.eof then
  begin
  	TOBEcart.SelectDB('',QQ,true);
  end;
  ferme (QQ);
  AddChampsFac;
  dMtSituation := 0;
  GetComponents;
  CreateComposants;
  ConstitueOngletInit;
  fCurrentOnglet := FOngletGlobal;
  DefinieLaTOB;
  Constitue_TVDEVIS;
  DefinieGrille (GS);
  //
  ConstituelesOnglets;
  //
  SetOnglet(FOngletGlobal);
  GS.setFocus;
  AfficheInfos;
  AfficheLaGrille(GS,fCurrentOnglet.TOBlignes);
  GS.SetFocus;
  SetLigne ( GS,1,LIBELLE);
  SetEvents (fCurrentOnglet.GS);
	PS.OnChange := ChangeVue;
  PS.OnChanging := BeforeChangeVue;
  Tfvierge(ecran).OnResize := FormResize;
  OnDisplay;
end ;

procedure TOF_BTSAISIEAVANC.ConstitueOngletInit;
begin
  // gestion onglet principal
  FOngletGlobal  := TOnglet.create (self,true);
 	FOngletGlobal.TOBlignes := TOBLIGNES;
  FOngletGlobal.GS := Thgrid(getControl('GS'));
  FOngletGlobal.Name := TTabSheet (getControl('TSGLOBALE')).Name;
  FonGletGlobal.tabSheet := TTabSheet (getControl('TSGLOBALE'));
  CreateButtonOuvrir (FonGletGlobal,9999);
  Fonglets.Add(FOngletGlobal);
  //
  fCurrentOnglet := FOngletGlobal;
  fCurrentOnglet.fTVDevis := TVDevis;
  //
  GS := fCurrentOnglet.GS;

end;

procedure TOF_BTSAISIEAVANC.ConstitueLesOnglets;

	function IsTOBLignesVide(TOBdetail : TOB) : boolean;
  var Indice : integer;
  begin
    result := true;
    if TOBdetail.Detail.count = 0 then exit;
    for Indice := 0 to TOBdetail.Detail.Count -1 do
    begin
      if Pos(TOBdetail.detail[Indice].GetValue('GL_TYPELIGNE'),'ART;SD;OUP')>0 then
      begin
        result := false;
        break;
      end;
    end;
  end;
  
var TOBAFFINTERV,TOBDetail : TOB;
		QQ : TQuery;
    indice,IndOnglet : integer;
    TS : TTabSheet;
    fGS : THgrid;
    fOnglet,LastOnglet : TOnglet;
    Sql : string;
    SousTraitantEnAchat : Boolean;
begin
  IndOnglet := 0;
	lastOnglet := fCurrentOnglet;

  //
  TOBAFFINTERV := TOB.create ('LES COTRAITS',nil,-1);
	if (TOBAffaire.getvalue('AFF_MANDATAIRE') = 'X') then
  begin
    // Co Traitants
    Sql := 'SELECT BAI_TIERSFOU,BAI_TYPEINTERV,'+
           '(SELECT T_LIBELLE FROM TIERS WHERE T_TIERS=BAI_TIERSFOU AND T_NATUREAUXI="FOU") AS LIBELLE '+
           'FROM AFFAIREINTERV '+
           'WHERE BAI_AFFAIRE="'+TOBAffaire.getvalue('AFF_AFFAIRE')+'"';
    QQ := OpenSql (Sql,True,-1,'',true);
    if not QQ.eof then
    begin
      TOBAFFINTERV.LoadDetailDB('AFFAIREINTERV','','',QQ,false);
    end;
    ferme (QQ);
  end;
  //
  fCotraitance := (TOBAFFINTERV.detail.count > 0);
  fSoustraitance := (TOBSOUSTraits.Detail.count >0);
  //
  if fCotraitance or fSoustraitance then
  begin
    TOBDetail := TOB.Create ('PIECE',nil,-1);
    TOBDetail.Dupliquer(TOBLIGNES ,true,true);
    EnlevepasConcerne (TOBdetail,'');
    if not IsTOBLignesVide(TOBdetail) then
    begin
      // il ya bien des intervenants --> on cree l'onglet entreprise
      TS := TTabSheet.Create(PS);
      TS.Parent := PS;
      TS.PageControl := PS;
      TS.ImageIndex := -1;
      TS.name := 'TS0';
      TS.caption := GetParamSocSecur('SO_LIBELLE','Notre entreprise');
      TS.TabVisible := true;
      //
      fGS := THgrid.Create(TS);
      fGS.Parent := TS;
      FGS.FixedCols := 1;
      FGS.FixedRows := 1;
      FGS.RowHeights [0] := GS.Rowheights[0];
      FGS.DefaultRowHeight  := GS.DefaultRowHeight ;
      FGS.Align := alclient;
      FGS.Visible := true;
      FGS.Couleur := GS.Couleur;
      FGS.Options := GS.options;
      FGS.EnabledBlueButton := GS.EnabledBlueButton;
      FGS.TypeSais := GS.TypeSais;
      FGS.Ctl3D := false;
      FGS.Color := GS.color;
      FGS.FixedColor := GS.FixedColor;
      FGS.PopupMenu  := GS.PopupMenu;

      //
      fOnglet := TOnglet.create(self,false);
      fOnglet.fournisseur := '';
      fOnglet.name := TS.name;
      fOnglet.GS := fGS;
      FOnglet.tabSheet := TS;
      Fonglet.TOBlignes := TOBdetail;
      CreateButtonOuvrir (Fonglet,0);
      //
      fCurrentOnglet :=fOnglet;
      GS := fCurrentOnglet.GS; // pour le cas ou
      //
      Constitue_TVDEVIS;
      //
      //
      Fonglets.Add(fOnglet);
      DefinieGrille(fGS);
      AfficheLaGrille(fGS,fOnglet.TOBlignes);
      SetLigne ( fGS,1,LIBELLE);
      CalculeTotauxSaisie;
      SetEvents (fGS);
    end else
    begin
      TOBdetail.free;
    end;
  end;

  for Indice := 0 to TOBAFFINTERV.detail.count -1 do
  begin
    //
    TOBDetail := TOB.Create ('PIECE',nil,-1);
    TOBDetail.Dupliquer(TOBLIGNES ,true,true);
    EnlevepasConcerne (TOBdetail,TOBAFFINTERV.detail[Indice].getString('BAI_TIERSFOU'));
    if TOBdetail.detail.count = 0 then
    begin
    	TOBdetail.free;
      continue;
    end;
    IndOnglet := indice +1;

    TS := TTabSheet.Create(PS);
    TS.Parent := PS;
    TS.PageControl := PS;
    TS.ImageIndex := 0;
    TS.name := 'TS'+inttostr(IndOnglet);
    TS.caption := TOBAFFINTERV.detail[Indice].getValue('LIBELLE');
    TS.TabVisible := true;
    //
    fGS := THgrid.Create(TS);
    fGS.Parent := TS;
    FGS.FixedCols := 1;
    FGS.FixedRows := 1;
    FGS.DefaultRowHeight  := GS.DefaultRowHeight ;
    FGS.Align := alclient;
    FGS.RowHeights [0] := GS.Rowheights[0];
    FGS.Options := GS.options;
    FGS.Visible := true;
    FGS.Couleur := GS.Couleur;
    FGS.EnabledBlueButton := GS.EnabledBlueButton;
    FGS.TypeSais := GS.TypeSais;
    FGS.Ctl3D := false;
    FGS.Color := GS.color;
    FGS.FixedColor := GS.FixedColor;
    FGS.PopupMenu  := GS.PopupMenu;
    //
    fOnglet := TOnglet.create(self,false);
    fOnglet.fournisseur := TOBAFFINTERV.detail[Indice].getValue('BAI_TIERSFOU');
    fOnglet.name := TS.name;
    fOnglet.GS := fGS;
    FOnglet.tabSheet := TS;
    Fonglet.TOBlignes := TOBdetail;
    CreateButtonOuvrir (Fonglet,IndOnglet);
    Fonglets.Add(fOnglet);
    //
    fCurrentOnglet :=fOnglet;
    GS := fCurrentOnglet.GS; // pour le cas ou
    //
    Constitue_TVDEVIS;
    //
    AfficheLaGrille(fGS,fOnglet.TOBlignes);
    DefinieGrille(fGS);
    AfficheLaGrille(fGS,fOnglet.TOBlignes);
    CalculeTotauxSaisie;
    SetLigne ( fGS,1,LIBELLE);
    SetEvents (fGS);
  end;
  TOBAFFINTERV.free;


  // -- GESTION DES SOUS TRAITANTS -- /
  for Indice := 0 to TOBSOUSTraits.Detail.count -1 do
  begin
    Inc(IndOnglet);
    SousTraitantEnAchat := (TOBSOUSTraits.detail[Indice].getValue('BPI_TYPEPAIE')='002');
    TOBDetail := TOB.Create ('PIECE',nil,-1);
    TOBDetail.Dupliquer(TOBLIGNES ,true,true);

    EnlevepasConcerne (TOBdetail,TOBSOUSTraits.detail[Indice].getValue('BPI_TIERSFOU'),SousTraitantEnAchat);
    if TOBdetail.detail.count = 0 then
    begin
    	TOBdetail.free;
      continue;
    end;
    TS := TTabSheet.Create(PS);
    TS.Parent := PS;
    TS.PageControl := PS;
    TS.ImageIndex := 1;
    TS.name := 'TS'+inttostr(IndOnglet);
    TS.caption := TOBSOUSTraits.detail[Indice].getValue('TLIBTIERSFOU');
    TS.TabVisible := true;
    //
    fGS := THgrid.Create(TS);
    fGS.Parent := TS;
    FGS.FixedCols := 1;
    FGS.FixedRows := 1;
    FGS.DefaultRowHeight  := GS.DefaultRowHeight ;
    FGS.Align := alclient;
    FGS.RowHeights [0] := GS.Rowheights[0];
    FGS.Options := GS.options;
    FGS.Visible := true;
    FGS.Couleur := GS.Couleur;
    FGS.EnabledBlueButton := GS.EnabledBlueButton;
    FGS.TypeSais := GS.TypeSais;
    FGS.Ctl3D := false;
    FGS.Color := GS.color;
    FGS.FixedColor := GS.FixedColor;
    FGS.PopupMenu  := GS.PopupMenu;
    //
    fOnglet := TOnglet.create(self,false);
    fOnglet.fournisseur := TOBSOUSTraits.detail[Indice].getValue('BPI_TIERSFOU');
    fOnglet.SousTraitant := True;
    fOnglet.name := TS.name;
    fOnglet.GS := fGS;
    FOnglet.tabSheet := TS;
    Fonglet.TOBlignes := TOBdetail;
    CreateButtonOuvrir (Fonglet,IndOnglet);
    Fonglets.Add(fOnglet);
    //
    fCurrentOnglet :=fOnglet;
    GS := fCurrentOnglet.GS; // pour le cas ou
    //
    Constitue_TVDEVIS;
    //
    AfficheLaGrille(fGS,fOnglet.TOBlignes);
    DefinieGrille(fGS);
    AfficheLaGrille(fGS,fOnglet.TOBlignes);
    CalculeTotauxSaisie;
    SetLigne ( fGS,1,LIBELLE);
    SetEvents (fGS);
  end;
  SetOnglet(LastOnglet);
//	fCurrentOnglet :=lastOnglet;
//  GS := fCurrentOnglet.GS; // pour le cas ou
end;

procedure TOF_BTSAISIEAVANC.OnClose ;
begin
	PS.ActivePage := FOngletGlobal.TabSheet;
  TOBOuvrages.free;
  fOnglets.free;
  TOBSOUSTraits.free;
  TOBPORCS.free;
  TOBLIGNES.free;
  TOBTiers.free;
  TOBaffaire.free;
  LESTOBACOMPTES.free;
  TOBAcomptes_O.free;
  TOBInfoPieces.Free;
  FreeComposants;
  Inherited ;
end ;

procedure TOF_BTSAISIEAVANC.OnDisplay () ;
var indice : integer;
begin
	for Indice := 1 to fOnglets.Count -1 do
  begin
  	Tonglet(fOnglets.Items [indice]).GS.RowHeights [0] :=  FOngletGlobal.GS.Rowheights[0];
  	Tonglet(fOnglets.Items [indice]).GS.ColWidths [0] :=  FOngletGlobal.GS.ColWidths [0];
    Tonglet(fOnglets.Items [indice]).GS.Ctl3D := FOngletGlobal.GS.Ctl3D;
    Tonglet(fOnglets.Items [indice]).GS.Color := FOngletGlobal.GS.color;
    Tonglet(fOnglets.Items [indice]).GS.FixedColor := FOngletGlobal.GS.FixedColor;
    Tonglet(fOnglets.Items [indice]).TabSheet.BorderWidth := FOngletGlobal.TabSheet.BorderWidth;
  end;
  GS.RowHeights [0] :=  FOngletGlobal.GS.Rowheights[0];
  Inherited ;
end ;

procedure TOF_BTSAISIEAVANC.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIEAVANC.GetComponents;
begin
	PS := TPageControl (GetControl('PS'));
//  GS := THgrid(GetControl ('GS'));
  AFF_TIERS := THEdit(GetCOntrol('AFF_TIERS'));
  AFF_AFFAIRE0 := THEdit(GetCOntrol('AFF_AFFAIRE0'));
  AFF_AFFAIRE1 := THEdit(GetCOntrol('AFF_AFFAIRE1'));
  AFF_AFFAIRE2 := THEdit(GetCOntrol('AFF_AFFAIRE2'));
  AFF_AFFAIRE3 := THEdit(GetCOntrol('AFF_AFFAIRE3'));
  AFF_AVENANT := THEdit(GetCOntrol('AFF_AVENANT'));
  AFF_AFFAIRE := THEdit(GetCOntrol('AFF_AFFAIRE'));
  BPF_MTSITUATION := Thlabel(GetCOntrol('BPF_MTSITUATION'));
  TDEVISE := THlabel(GetControl('TDEVISE'));
  FMTSITUATION := THNumEdit(GetCOntrol('FMTSITUATION'));
  AFF_DESIGNATION := THLabel(GetCOntrol('AFF_DESIGNATION'));
  LIBTIERS := THLabel (GetCOntrol('LIBTIERS'));
  BTDescriptif := TToolbarButton97 (GetControl('BTDESCRIPTIF'));
  PTVDevis := THPanel (GetControl('PTVDEVIS'));
  TVDevis := TTreeview (GetControl('TVDEVIS'));
  BAVANC := TToolbarButton97 (GetCOntrol('BAVANC'));
  //
  POPACTIONS := TPopupMenu (GetCOntrol('POPACTIONS'));
  MnDefaireLig := TMenuItem(GetControl('MnDefaireLig'));
  MnDefaireParag := TMenuItem(GetControl('MnDefaireParag'));
  MnDefairePiece := TMenuItem(GetControl('MnDefairePiece'));
  MnDefaireTous := TMenuItem(GetCOntrol('MnDefaireTous'));
  MnAvancement := TMenuItem(GetCOntrol('MnAvancement'));

  //
  VOIRQTE := TCheckBox(GetCOntrol('VOIRQTE'));
  VOIRMT := TCheckBox(GetCOntrol('VOIRMT'));
  Bferme := TToolbarButton97 (GetCOntrol('BFERME'));
  Bvalider := TToolbarButton97 (GetCOntrol('BVALIDER'));
  //
end;

procedure TOF_BTSAISIEAVANC.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,
  Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
  inherited;
  Aff0 := THEdit(GetControl('AFF_AFFAIRE0'));
  Aff := THEdit(GetControl('AFF_AFFAIRE'));
  Aff1 := THEdit(GetControl('AFF_AFFAIRE1'));
  Aff2 := THEdit(GetControl('AFF_AFFAIRE2'));
  Aff3 := THEdit(GetControl('AFF_AFFAIRE3'));
  Aff4 := THEdit(GetControl('AFF_AVENANT'));
end;

procedure TOF_BTSAISIEAVANC.CreateComposants;
begin
	fOnglets := TlistOnglet.create;
	//
  ImTypeArticle := TImageList.create(ecran);
  ImTypeArticle.ImageType := itImage;
  ImgListPlusMoins:= TImageList.create(ecran);
  ImgListPlusMoins.ImageType := itImage;
  ImgTypeTrait:= TImageList.create(ecran);
  ImgTypeTrait.ImageType := itImage;

  //
  ToolWin := TToolWindow97.Create (ecran);
  ToolWin.Caption := 'Structure de la facture';
  ToolWin.Parent := Ecran;
  ToolWin.ClientWidth := PTVDEVIS.Width+10;
  ToolWin.ClientHeight := PTVDEVIS.Height+10;
  ToolWin.Top := PTVDEVIS.Top;
  ToolWin.Left := PTVDEVIS.Left;
  ToolWin.Visible := false;
  //
  PTVDEVIS.parent := ToolWin;
  PTVDEVIS.Align := Alclient;
  //
  ChargeListImage;
  PS.Images := ImgTypeTrait;
end;

procedure TOF_BTSAISIEAVANC.DefinieLaTOB;
var Indice,iTot : integer;
    TOBL : TOB;
    QQ : TQuery;
    XX : TFsplashScreen;
    TypeArticle : string;
    NatureTravail : Double;
begin
  TOBaffaire.initValeurs;
  TOBTiers.InitValeurs;
  XX := TFsplashScreen.Create (application);
  XX.Label1.Caption  := TraduireMemoire('Constitution de la saisie en cours...');
  XX.Show;
  XX.BringToFront;
  XX.refresh;
  TRY
  	if TOBSelect.getString('GP_AFFAIRE') <> '' then
    begin
      QQ := OPenSql ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+ TOBSelect.getString('GP_AFFAIRE')+'"',True);
      TOBaffaire.SelectDB('',QQ);
      ferme (QQ);
    end;
    QQ := OPenSql ('SELECT * FROM TIERS WHERE T_TIERS="'+ TOBSelect.getValue('GP_TIERS')+'" AND T_NATUREAUXI="CLI"',True);
    TOBTiers.SelectDB ('',QQ);
    ferme (QQ);
    TOBLIGNES.ClearDetail;
    TOBOuvrages.ClearDetail;
    iTot := -1;
    for Indice := 0 to TOBSelect.detail.count -1 do
    begin
      TOBL := TOBSelect.detail[Indice];
      if Indice = 0 then
      begin
        SetInfoFac (TOBL);
        GetInfoDev (TOBL.getValue('GP_DEVISE'));
        GetPorts (TOBL);
      end;
      AddLeDocument(TOBL);
    end;
    //
    if ReajusteDossier then
    begin
//      ReindiceOuvrages;
    end;
    //
    Indice := 0;
    MaxNumOrdre := 0;
    repeat
      TOBL := TOBLIGNES.detail[Indice];
      //
      //  
      //PgiBox(IntToStr(Indice), 'DefinieLaTOB');
      //
      if (ModifSituation) and (TOBL.GetString('GL_PIECEORIGINE')<>'') then
      begin
        MemoriseHighNumOrdre (TOBL);
        if TOBL.GetInteger('GL_NUMORDRE')>MaxNumOrdre then MaxNumOrdre := TOBL.GetInteger('GL_NUMORDRE'); 
      end;
      if indice = 0 then iTot := TOBL.GetNumChamp('GL_TYPELIGNE');
      if Isvariante(TOBL) then begin TOBL.free; continue; END; // supression des variantes dans la facturation
      AddLesSupLigne(TOBL,false);
      AddChampsSupAvanc (TOBL);
      AddChampsSupMtINITAvanc (TOBL);
      PositionneLigne (TOBL,Indice+1);
      // NOUVEAU GESTION DES OUVRAGES
      TypeArticle := TOBL.getValue('GL_TYPEARTICLE');
			NatureTravail := valeur(TOBL.GetValue('GLC_NATURETRAVAIL'));
      if IsOuvrage(TOBL) and ((NatureTravail >= 10) or (ReajusteDossier)) then
      begin
        TOBL.PutValue('GLC_VOIRDETAIL','X');
      end;
      if IsOuvrage(TOBL) then
      begin
        if IsAvancementDetailOuvrage (TOBL) then TOBL.PutValue('GLC_VOIRDETAIL','X');
      end;
      if TOBL.GetValue('GL_TYPELIGNE')='ART' then CumuleLigne(TOBL);
      CalculeSousTotal (indice+1,iTot);
      if (isOuvrage (TOBL)) and  (pos(TypeArticle,'ARP;ARV;')=0) and (TOBL.GetValue('GLC_VOIRDETAIL')='X') then
      begin
        AfficheSousDetail (fCurrentOnglet,TOBLIGNES,TOBL,Indice);
      end;
      //
      inc(indice);
    until Indice >= TOBLignes.detail.count;
    indice := 0;

    if ModifSituation then
    begin
      XX.Label1.Caption  := TraduireMemoire('Synchronisation avec les devis constituants...');
      SynchroniseDocumentWithDevis;
    end;

    if not TOBLIGNES.FieldExists('GP_RGMULTIPLE') then
    begin
    	TOBLIGNES.AddChampSupValeur ('GP_RGMULTIPLE','-');
    end;
    TOBL := TOBLignes.findFirst (['GL_TYPELIGNE'],['RG'],true);
    repeat
    	if TOBL = nil then break;
    	TOBL := TOBLignes.findnext (['GL_TYPELIGNE'],['RG'],true);
      if TOBL <> nil then
      begin
        TOBLIGNES.putValue ('GP_RGMULTIPLE','X');
        break;
      end;
    until TOBL = nil;
    //
    BPF_MTSITUATION.caption := FMTSITUATION.Text;
  FINALLY
    XX.Free;
  END;
end;


procedure TOF_BTSAISIEAVANC.AddLignefacturationFromOuv (TOBL,TOBLignes,TOBO: TOB; Indice,decalage,IndiceNomen : integer;var montant : double;AvancSousDetailPartiel:boolean; EnAchat : Boolean=false);
var TOBN : TOB;
    cledoc : r_cledoc;
    NatureTravail,PQ : double;
begin
//  TOBL := TOBLignes.detail[Indice];
	NatureTravail := valeur(TOBL.GetValue('GLC_NATURETRAVAIL'));
	FillChar (cledoc,sizeof(cledoc),#0);
  cledoc.NaturePiece := TOBL.GetValue('GL_NATUREPIECEG');
  cledoc.souche := TOBL.GetValue('GL_SOUCHE');
  cledoc.NumeroPiece := TOBL.GetValue('GL_NUMERO');
  cledoc.Indice  := TOBL.GetValue('GL_INDICEG');
  //
  TOBN := NewTOBLigne (TOBlignes,Indice+decalage+1,true);
  TOBN.ClearDetail;
  AddLesSupLigne(TOBN,false);
  AddChampsSupAvanc (TOBN);
	AddChampsSupMtINITAvanc (TOBN);
  PQ := TOBO.GetValue('BLO_PRIXPOURQTE'); if PQ =0 then PQ := 1;
  //
  // cas de figure ou il y a un d'avancement au niveau détail d'ouvrage
  TOBN.putValue('BLF_NATUREPIECEG',TOBO.Getvalue('BLF_NATUREPIECEG'));
  TOBN.putValue('BLF_SOUCHE',TOBO.Getvalue('BLF_SOUCHE'));
  TOBN.putValue('BLF_NUMERO',TOBO.Getvalue('BLF_NUMERO'));
  TOBN.putValue('BLF_INDICEG',TOBO.Getvalue('BLF_INDICEG'));
  TOBN.putValue('BLF_NUMORDRE',0);
  TOBN.putValue('BLF_UNIQUEBLO',TOBO.Getvalue('BLF_UNIQUEBLO'));
  TOBN.putValue('BLF_MTPRODUCTION',0);
  if EnAchat then
  begin
  	TOBN.putValue('BLF_MTMARCHE',arrondi(TOBO.getValue('BLF_QTEMARCHE')*TOBO.GEtValue('BLO_DPA') / PQ,DEV.decimale));
    TOBN.putValue('BLF_MTCUMULEFACT',arrondi(TOBO.getValue('BLF_QTECUMULEFACT')*TOBO.GEtValue('BLO_DPA') / PQ,DEV.decimale));
    TOBN.putValue('BLF_MTDEJAFACT',arrondi(TOBO.getValue('BLF_QTEDEJAFACT')*TOBO.GEtValue('BLO_DPA') / PQ,DEV.decimale));
    TOBN.putValue('BLF_MTSITUATION',arrondi(TOBO.getValue('BLF_QTESITUATION')*TOBO.GEtValue('BLO_DPA') / PQ,DEV.decimale));
  end
  else begin
  	TOBN.putValue('BLF_MTMARCHE',TOBO.Getvalue('BLF_MTMARCHE'));
    TOBN.putValue('BLF_MTCUMULEFACT',TOBO.Getvalue('BLF_MTCUMULEFACT'));
    TOBN.putValue('BLF_MTDEJAFACT',TOBO.Getvalue('BLF_MTDEJAFACT'));
    TOBN.putValue('BLF_MTSITUATION',TOBO.Getvalue('BLF_MTSITUATION'));
  end;
  TOBN.putValue('BLF_QTEMARCHE',TOBO.Getvalue('BLF_QTEMARCHE'));
  TOBN.putValue('BLF_QTECUMULEFACT',TOBO.Getvalue('BLF_QTECUMULEFACT'));
  TOBN.putValue('BLF_QTEDEJAFACT',TOBO.Getvalue('BLF_QTEDEJAFACT'));
  TOBN.putValue('BLF_QTESITUATION',TOBO.Getvalue('BLF_QTESITUATION'));
  TOBN.putValue('BLF_POURCENTAVANC',TOBO.Getvalue('BLF_POURCENTAVANC'));
  TOBN.putValue('BLF_NATURETRAVAIL',TOBO.Getvalue('BLF_NATURETRAVAIL'));
  TOBN.putValue('BLF_FOURNISSEUR',TOBO.Getvalue('BLF_FOURNISSEUR'));
  if TOBO.FieldExists('POURCENTAVANCINIT') then TOBN.putValue ('POURCENTAVANCINIT',TOBO.getValue('POURCENTAVANCINIT'));
  if TOBO.FieldExists('MTCUMULEFACTINIT')  then TOBN.putValue ('MTCUMULEFACTINIT',TOBO.getValue('MTCUMULEFACTINIT'));
  if TOBO.FieldExists('QTECUMULEFACTINIT') then TOBN.putValue ('QTECUMULEFACTINIT',TOBO.getValue('QTECUMULEFACTINIT'));
  //
  TOBN.putValue('GL_NATUREPIECEG',TOBL.Getvalue('GL_NATUREPIECEG'));
  TOBN.putValue('GL_SOUCHE',TOBL.Getvalue('GL_SOUCHE'));
  TOBN.putValue('GL_NUMERO',TOBL.Getvalue('GL_NUMERO'));
  TOBN.putValue('GL_INDICEG',TOBL.Getvalue('GL_INDICEG'));
  TOBN.putvalue('GL_NUMLIGNE',TOBL.getValue('GL_NUMLIGNE'));
  TOBN.putvalue('GL_DPA',TOBO.getValue('BLO_DPA'));
  if (NatureTravail > 0) and (Naturetravail < 10) then
  begin
    TOBN.putValue('GL_FOURNISSEUR',TOBL.Getvalue('GL_FOURNISSEUR'));
    TOBN.putValue('BLF_FOURNISSEUR',TOBL.Getvalue('GL_FOURNISSEUR'));
    TOBN.putvalue('GLC_NATURETRAVAIL',TOBL.getValue('GLC_NATURETRAVAIL'));
  end else
  begin
    TOBN.putValue('GL_FOURNISSEUR',TOBO.Getvalue('BLO_FOURNISSEUR'));
    TOBN.putValue('BLF_FOURNISSEUR',TOBO.Getvalue('BLO_FOURNISSEUR'));
    TOBN.putvalue('GLC_NATURETRAVAIL',TOBO.getValue('BLO_NATURETRAVAIL'));
  end;
  TOBN.putvalue('GL_TYPELIGNE','SD');
  TOBN.putvalue('GL_INDICENOMEN',IndiceNomen);
  TOBN.putvalue('GL_TYPEARTICLE',TOBO.getValue('BLO_TYPEARTICLE'));
  TOBN.putvalue('GL_ARTICLE',TOBO.getValue('BLO_ARTICLE'));
  TOBN.putvalue('GL_CODEARTICLE',TOBO.getValue('BLO_CODEARTICLE'));
  TOBN.putvalue('GL_REFARTSAISIE',TOBO.getValue('BLO_CODEARTICLE'));
  TOBN.putvalue('GL_LIBELLE',TOBO.getValue('BLO_LIBELLE'));
  TOBN.putvalue('GL_QUALIFQTEVTE',TOBO.getValue('BLO_QUALIFQTEVTE'));
  TOBN.putvalue('UNIQUEBLO',TOBO.getValue('BLO_UNIQUEBLO'));
  TOBN.putvalue('GL_QTEFACT',TOBO.Getvalue('BLF_QTEMARCHE'));
  TOBN.putvalue('GL_QTEPREVAVANC',TOBO.Getvalue('BLF_QTECUMULEFACT'));
  TOBN.putvalue('GL_QTESIT',TOBO.Getvalue('BLF_QTECUMULEFACT'));
  TOBN.putvalue('GL_NUMORDRE',TOBL.getValue('GL_NUMORDRE'));
  TOBO.putvalue('BLO_NUMLIGNE',TOBL.getValue('GL_NUMORDRE'));
  PositionneLigne (TOBN,Indice+decalage+1,EnAchat);
  //
end;

procedure TOF_BTSAISIEAVANC.OrganiseOuvrage(TOBPiece,TOBNomen: TOB; TOBLL : TOB=nil) ;

  function Findpere (TOBGroupeNomen: TOB;Ligne1,Ligne2,Ligne3,Ligne4 : integer;niveau : integer) : TOB;
  begin
    if Niveau = 5 then result := TOBGroupeNomen.findfirst(['BLO_N1','BLO_N2','BLO_N3','BLO_N4'],[Ligne1,Ligne2,Ligne3,Ligne4],true)
    else if Niveau = 4 then result := TOBGroupeNomen.findfirst(['BLO_N1','BLO_N2','BLO_N3'],[Ligne1,Ligne2,Ligne3],true)
    else if Niveau = 3 then result := TOBGroupeNomen.findfirst(['BLO_N1','BLO_N2'],[Ligne1,Ligne2],true)
    else if Niveau = 2 then result := TOBGroupeNomen.findfirst(['BLO_N1'],[Ligne1],true)
    else result := nil;
    if (result = nil) and (Niveau > 1) then
    begin
      result := FindPere (TOBGroupeNomen,Ligne1,Ligne2,Ligne3,Ligne4,Niveau-1);
    end;
  end;

var i,II,IndNumLigne,IndN1,IndN2,IndN3,IndN4,IndN5,Lig : integer;
		TOBLN,TOBL,TOBGN : TOB;
    LaLig : integer;
    LigneN1,LigneN2,LigneN3,LigneN4,LigneN5 : integer;
    Tobpere,TOBpere1,TOBpere2,TOBpere3,TOBPere4,TOBGroupeNomen : TOB;
BEGIN
	TOBGN := TOB.Create ('LES OUVRAGES (WAITTT)',nil,-1);
  //
  i := 0;
  LaLig := -1;
  if TOBNomen.Detail.count = 0 then exit;
  IndNumLigne := TOBNomen.detail[0].GetNumChamp('BLO_NUMLIGNE') ;
  IndN1 := TOBNomen.detail[0].GetNumChamp('BLO_N1');
  IndN2 := TOBNomen.detail[0].GetNumChamp('BLO_N2');
  IndN3 := TOBNomen.detail[0].GetNumChamp('BLO_N3');
  IndN4 := TOBNomen.detail[0].GetNumChamp('BLO_N4');
  IndN5 := TOBNomen.detail[0].GetNumChamp('BLO_N5');
  repeat
    TOBLN:=TOBNomen.Detail[i] ;
    //
    Lig:=TOBLN.GetValeur(IndNumLigne) ;
    if lig <> LaLig then
    begin
      LaLig := Lig;
      TOBGroupeNomen:=TOB.Create('',TOBGN,-1) ;
      TOBGroupeNomen.AddChampSup('UTILISE',False) ;
      TOBGroupeNomen.PutValue('UTILISE','-') ;
      TOBGroupeNomen.AddChampSupValeur ('AVANCSOUSDET','-') ;
      if TOBLL = nil then TOBL := TOBPiece.detail[LaLig-1]
                     else TOBL := TOBLL; 
    end;
    //
    if (TOBL.getValue('BLF_POURCENTAVANC')<> TOBLN.Getvalue('BLF_POURCENTAVANC')) and
       (TOBLN.Getvalue('BLF_POURCENTAVANC')<>0) AND ((TOBL.Getvalue('BLF_POURCENTAVANC')<>0)) and
       (TOBLN.Getvalue('BLO_N2')=0)then
    begin
      TOBGroupeNomen.putValue('AVANCSOUSDET','X');
    end;
    LigneN1 := TOBLN.GetValeur(IndN1);
    LigneN2 := TOBLN.GetValeur(IndN2);
    LigneN3 := TOBLN.GetValeur(IndN3);
    LigneN4 := TOBLN.GetValeur(IndN4);
    LigneN5 := TOBLN.GetValeur(IndN5);
    //
    if LigneN2 = 0 then
    begin
    	TOBPere:=TOBGroupeNomen;
      // Stockage du pere de niveau 1
      TOBpere1 := TOBLN;
    end else if LigneN3 = 0 then
    begin
      // Stockage du pere de niveau 2
      TOBpere := TOBpere1;
      TOBpere2 := TOBLN;
    end else if LigneN4 = 0 then
    begin
      TOBpere := TOBpere2;
      // Stockage du pere de niveau 3
      TOBpere3 := TOBLN;
    end else if LigneN5=0 then
    begin
      TOBpere := TOBpere3;
      // Stockage du pere de niveau 4
      TOBpere4 := TOBLN;
    end else TOBpere := TOBPere4;

    if TOBPere = nil then
    begin
      if LigneN2 = 0 then Findpere (TOBGroupeNomen,LigneN1,LigneN2,LigneN3,LigneN4,2)
      else if LigneN3 = 0 then Findpere (TOBGroupeNomen,LigneN1,LigneN2,LigneN3,LigneN4,3)
      else if LigneN4 = 0 then Findpere (TOBGroupeNomen,LigneN1,LigneN2,LigneN3,LigneN4,4)
      else if LigneN5 = 0 then Findpere (TOBGroupeNomen,LigneN1,LigneN2,LigneN3,LigneN4,5);
      if TOBPere = nil then TOBpere := TOBGroupeNomen;
    end;

    if (TOBPere<>Nil) then
    BEGIN
      InsertionChampSupOuv (TOBLN,false);
      TOBLN.ChangeParent (TOBpere,-1);
      //
    END else
    begin
      TOBLN.free;
    end;
  until i>=TOBnomen.detail.count;
  i := 0;
  repeat
    TOBLN:=TOBGN.Detail[i] ;
    //FV1 - 26/07/2016 - Erreur sur préparation de facture
    if TOBLN.detail.count > 0 then
    begin
    Lig:=TOBLN.detail[0].GetValue('BLO_NUMLIGNE') ;
    end;
    //
    if TOBLL = nil then TOBL := TOBpiece.detail[Lig-1]
                   ELSE TOBL := TOBLL;
    //
    for II := 0 to TOBLN.detail.count - 1 do
    begin
      SetAvancementDetailOuvrage(TOBLN,TOBLN.detail[II],TOBL);
    end;
    TOBLN.ChangeParent(TOBOuvrages,-1);
    TOBL.putValue('GL_INDICENOMEN',TOBLN.GetIndex+1);
    //
  until I >= TOBGN.detail.count;
  //
  TOBGN.free;
END ;

procedure TOF_BTSAISIEAVANC.SetAvancementDetailOuvrage(TOBgroupeNomen,TOBO,TOBL: TOB);

procedure InitChampsSupBLF(TOBO : TOB);
begin
  TOBO.putValue('BLF_NATUREPIECEG','');
  TOBO.putValue('BLF_SOUCHE','');
  TOBO.putValue('BLF_NUMERO',0);
  TOBO.putValue('BLF_INDICEG',0);
  TOBO.putValue('BLF_NUMORDRE',0);
  TOBO.putValue('BLF_UNIQUEBLO',0);
  TOBO.putValue('BLF_NATURETRAVAIL','');
  TOBO.putValue('BLF_FOURNISSEUR','');
  TOBO.putValue('BLF_MTMARCHE',0);
  TOBO.putValue('BLF_MTCUMULEFACT',0);
  TOBO.putValue('BLF_MTDEJAFACT',0);
  TOBO.putValue('BLF_MTSITUATION',0);
  TOBO.putValue('BLF_QTEMARCHE',0);
  TOBO.putValue('BLF_QTECUMULEFACT',0);
  TOBO.putValue('BLF_QTEDEJAFACT',0);
  TOBO.putValue('BLF_QTESITUATION',0);
  TOBO.putValue('BLF_POURCENTAVANC',0);
  TOBO.putValue('BLF_MTPRODUCTION',0);
end;

var AvancSousDetailPartiel : boolean;
		QteDet,QteDejafact,MtDetDev,MtDejaFact,avancprec,Pourcentavanc : double;
begin
	AvancSousDetailPartiel := (TOBgroupeNomen.getvalue('AVANCSOUSDET')='X');
  Pourcentavanc := TOBL.geTDouble('BLF_POURCENTAVANC');
  if (TOBO.GetValue('BLF_NATUREPIECEG')= '') then
  begin
  	InitChampsSupBLF(TOBO);
    if TOBO.getValue('BLO_QTEDUDETAIL') = 0 then TOBO.SetDouble('BLO_QTEDUDETAIL',1);
    // cas de figure ou il n'y a pas d'avancement au niveau détail d'ouvrage --> mais bien au niveau ligne de piece
    if AvancSousDetailPartiel then
    begin
      QteDet := arrondi(TOBO.getValue('BLO_QTEFACT')*TOBL.getvalue('GL_QTEFACT') /TOBO.getValue('BLO_QTEDUDETAIL'),V_PGI.okdecQ);
      QteDejaFact := 0;
      MtDetDev := arrondi(QteDet*TOBO.getValue('BLO_PUHTDEV'),DEV.Decimale);
      MtDejaFact := 0;
    end else
    begin
      if (TOBL.GetString('BLF_NATUREPIECEG') = '')  then
      begin
        if TOBL.GetDouble('GL_QTEFACT')<>0 then AvancPrec := TOBL.GetDouble('GL_QTEPREVAVANC') / TOBL.GetDouble('GL_QTEFACT')
                                           else AvancPrec := 0;
        if PourcentAvanc = 0 then PourcentAvanc := Arrondi(Avancprec * 100,2);
      end else
      begin
      if TOBL.Getvalue('BLF_MTMARCHE')  > 0 then avancprec := TOBL.Getvalue('BLF_MTDEJAFACT')/TOBL.Getvalue('BLF_MTMARCHE')
                                            else avancPrec := 0;
        if PourcentAvanc = 0 then PourcentAvanc := Arrondi(Avancprec * 100,2);
      end;
      QteDet := arrondi(TOBO.getValue('BLO_QTEFACT')*TOBL.getvalue('GL_QTEFACT') /TOBO.getValue('BLO_QTEDUDETAIL'),V_PGI.okdecQ);
      QteDejaFact := arrondi(QteDet * avancprec,V_PGI.okdecQ);
      MtDetDev := arrondi(QteDet*TOBO.getValue('BLO_PUHTDEV'),DEV.Decimale);
      MtDejaFact := arrondi(QteDejaFact*TOBO.getValue('BLO_PUHTDEV'),DEV.Decimale);
    end;
    //
    TOBO.putValue('BLF_NATUREPIECEG',TOBL.Getvalue('GL_NATUREPIECEG'));
    TOBO.putValue('BLF_SOUCHE',TOBL.Getvalue('GL_SOUCHE'));
    TOBO.putValue('BLF_NUMERO',TOBL.Getvalue('GL_NUMERO'));
    TOBO.putValue('BLF_INDICEG',TOBL.Getvalue('GL_INDICEG'));
    TOBO.putValue('BLF_NUMORDRE',0);
    TOBO.putValue('BLF_UNIQUEBLO',TOBO.Getvalue('BLO_UNIQUEBLO'));
    TOBO.putValue('BLF_NATURETRAVAIL',TOBO.Getvalue('BLO_NATURETRAVAIL'));
    TOBO.putValue('BLF_FOURNISSEUR',TOBO.Getvalue('BLO_FOURNISSEUR'));
    if pos(ModeGeneration,'AVA;DAC')>0 then
    begin
    	if AvancSousDetailPartiel then
      begin
        TOBO.putValue('BLF_MTMARCHE',MtDetDev);
        TOBO.putValue('BLF_MTCUMULEFACT',0);
        TOBO.putValue('BLF_MTDEJAFACT',0);
        TOBO.putValue('BLF_MTSITUATION',0);
        TOBO.putValue('BLF_QTEMARCHE',QteDet);
        TOBO.putValue('BLF_QTECUMULEFACT',0);
        TOBO.putValue('BLF_QTEDEJAFACT',0);
        TOBO.putValue('BLF_QTESITUATION',0);
        TOBO.putValue('BLF_POURCENTAVANC',0);
      end else
      begin
        TOBO.putValue('BLF_MTMARCHE',MtDetDev);
        TOBO.putValue('BLF_MTCUMULEFACT',arrondi(MtDetDev*PourcentAvanc/100,DEV.Decimale ));
        TOBO.putValue('BLF_MTDEJAFACT',arrondi(MtDejafact,DEV.decimale));
        TOBO.putValue('BLF_MTSITUATION',0);
        TOBO.putValue('BLF_QTEMARCHE',QteDet);
        TOBO.putValue('BLF_QTECUMULEFACT',arrondi(QteDet*PourcentAvanc/100,V_PGI.okdecQ));
        TOBO.putValue('BLF_QTEDEJAFACT',arrondi(QteDejaFact,V_PGI.okdecQ));
        TOBO.putValue('BLF_QTESITUATION',0);
        TOBO.putValue('BLF_POURCENTAVANC',Arrondi(PourcentAvanc,2));
      end;
    end else
    begin
      TOBO.putValue('BLF_MTMARCHE',MtDetDev);
      TOBO.putValue('BLF_MTCUMULEFACT',0);
      TOBO.putValue('BLF_MTDEJAFACT',0);
      TOBO.putValue('BLF_MTSITUATION',0);
//      TOBO.putValue('BLF_QTEMARCHE',TOBL.Getvalue('BLF_QTEMARCHE'));
      TOBO.putValue('BLF_QTEMARCHE',QteDet);
      TOBO.putValue('BLF_QTECUMULEFACT',0);
      TOBO.putValue('BLF_QTEDEJAFACT',0);
      TOBO.putValue('BLF_QTESITUATION',0);
      TOBO.putValue('BLF_POURCENTAVANC',0);
    end;
    TOBO.putValue('BLF_MTPRODUCTION',0);
    AddChampsSupMtINITAvanc (TOBO);
    InitMontantSupfacturation(TOBO);
  end else
  begin
    if ModifSituation then
    begin
      REcalcMtSupfactOuv (TOBO);
    end else if TOBO.GetValue('BLF_MTSITUATION') <> 0 then
    begin
      MtDejaFact := TOBO.GetValue('BLF_MTCUMULEFACT');
      MtDetDev := TOBO.GetValue('BLF_MTMARCHE');
      if MtDetDev <> 0 then PourcentAvanc := Arrondi(MtDejaFact / MtDetDev * 100,2) else PourcentAvanc := 0;
    	AddChampsSupMtINITAvanc (TOBO);
      InitMontantSupfacturation (TOBO);
      TOBO.putValue('POURCENTAVANCINIT',Pourcentavanc);
      TOBO.putValue('BLF_MTCUMULEFACT',TOBO.GetValue('BLF_MTCUMULEFACT')+TOBO.GetValue('BLF_MTSITUATION'));
      TOBO.putValue('BLF_QTECUMULEFACT',TOBO.GetValue('BLF_QTECUMULEFACT')+TOBO.GetValue('BLF_QTESITUATION'));
    end else
    begin
    	AddChampsSupMtINITAvanc (TOBO);
      InitMontantSupfacturation (TOBO);
    end;
  end;
end;

procedure TOF_BTSAISIEAVANC.AddLeDocument(TOBS : TOB);
var TOBL,TOBLig,TOBNomen : TOB;
    CD : R_Cledoc;
    Sql : string;
    QQ : TQuery;
    Indice : integer;
begin
  TOBNomen := TOB.Create ('LES DETAILS OUV',nil,-1);
  TOBLig := TOB.Create ('LES DETAILS',nil,-1);
  //
  TRY
    RecupAcomptePiece (TOBS);
    InitChampsSupEntete(TOBS);
    if not ModifSituation then
    begin
      TOBL := TOB.Create ('LIGNE',TOBLIGNES,-1);
  //    AddLesSupLigne(TOBL,false,true);
    // Entete de reference
    TOBL.PutValue('GL_NATUREPIECEG',TOBS.GetValue('GP_NATUREPIECEG'));
    TOBL.PutValue('GL_SOUCHE',TOBS.GetValue('GP_SOUCHE'));
    TOBL.PutValue('GL_NUMERO',TOBS.GetValue('GP_NUMERO'));
    TOBL.PutValue('GL_INDICEG',TOBS.GetValue('GP_INDICEG'));
    //
    TOBL.putValue('GL_TYPELIGNE','REF');
    TOBL.PutValue('GL_LIBELLE','Devis N° '+InttoStr(TOBS.GetValue('GP_NUMERO'))+' du '+DateToStr (TOBS.GetValue('GP_DATEPIECE')));
    // -----
    end;
    CD := TOB2CleDoc (TOBS);
    Sql := MakeSelectLigneBtp(true,false,true)+' WHERE  '+
           WherePiece (CD,ttdLigne,false)+ ' '+
           'ORDER BY GL_NUMLIGNE';
    QQ := OpenSql (Sql,True);
    EnregistrelesSousTraits (TOBS);  // Stocke les sous traitants de la pièce
    TOBLig.LoadDetailDB ('LIGNE','','',QQ,true);
    ferme (QQ);
    // NOUVEAU -- Les details d'ouvrages
    Sql := MakeSelectLigneOuvBtp (True);
    Sql := Sql + ' WHERE '+WherePiece(CD,ttdOuvrage,true)+' ORDER BY BLO_NUMLIGNE,BLO_N1,BLO_N2,BLO_N3,BLO_N4,BLO_N5';
    QQ := OpenSql (Sql,true,-1,'',true);
    if not QQ.eof then
    begin
      TOBNomen.LoadDetailDB('LIGNEOUV','','',QQ,true,true);
      OrganiseOuvrage(TOBLIG,TOBNomen);
    end;
    ferme (QQ);
    //
    indice := 0;
    repeat
    	TOBL := TOBLig.detail[Indice];
      if ModifSituation then
      begin
        if (TOBL.GetString('GL_TYPEARTICLE')='EPO') and (TOBL.GetString('GL_TYPELIGNE')='COM') Then
        begin
          // c'est un entete de document --> on le passe comme tel dans la préparation
          TOBL.putValue('GL_TYPELIGNE','REF');
        end;
      end;
    	TOBL.ChangeParent(TOBLIGNES ,-1);
    until Indice >= TOBLig.detail.count;
  FINALLY
    TOBNomen.free;
    TOBLig.free;
  END;
end;

procedure TOF_BTSAISIEAVANC.REcalcMtSupfactOuv (TOBOUV : TOB);
begin
  AddChampsSupMtINITAvanc (TOBOUV);
  RecalcMontantSupfacturation (TOBOUV);
end;

procedure TOF_BTSAISIEAVANC.RecalcMontantSupfacturation(TOBL : TOB);
var LastQteCumulefact,LastMtCumuleFact,lastPourcentavanc : Double;
begin
  if avancement then
  begin
    LastQteCumulefact := TOBL.getValue('BLF_QTECUMULEFACT') - TOBL.getValue('BLF_QTESITUATION');
    LastMtCumulefact := TOBL.getValue('BLF_MTCUMULEFACT') - TOBL.getValue('BLF_MTSITUATION');
    if TOBL.GetDouble('BLF_MTMARCHE')<>0 then
    begin
      lastPourcentavanc := ARRONDI((LastMtCumulefact / TOBL.getValue('BLF_MTMARCHE'))*100,2);
    end else lastPourcentavanc := 100;
    if TOBL.FieldExists('POURCENTAVANCINIT') then TOBL.putValue ('POURCENTAVANCINIT',lastPourcentavanc);
    if TOBL.FieldExists('MTCUMULEFACTINIT') then TOBL.putValue ('MTCUMULEFACTINIT',LastMtCumulefact);
    if TOBL.FieldExists('QTECUMULEFACTINIT') then TOBL.putValue ('QTECUMULEFACTINIT',LastQteCumulefact);
  end else
  begin
    if TOBL.FieldExists('POURCENTAVANCINIT') then TOBL.putValue ('POURCENTAVANCINIT',TOBL.getValue('BLF_POURCENTAVANC'));
  end;
end;

procedure TOF_BTSAISIEAVANC.InitMontantSupfacturation(TOBL : TOB);
begin
  if Avancement then
  begin
  if TOBL.FieldExists('MTCUMULEFACTINIT') then TOBL.putValue ('MTCUMULEFACTINIT',TOBL.getValue('BLF_MTCUMULEFACT'));
  if TOBL.FieldExists('QTECUMULEFACTINIT') then TOBL.putValue ('QTECUMULEFACTINIT',TOBL.getValue('BLF_QTECUMULEFACT'));
  end;
  if TOBL.FieldExists('POURCENTAVANCINIT') then TOBL.putValue ('POURCENTAVANCINIT',TOBL.getValue('BLF_POURCENTAVANC'));
end;


procedure TOF_BTSAISIEAVANC.SetInfoFactEcart (TOBL : TOB);
begin
  TOBL.putValue('BLF_MTMARCHE',0);
  TOBL.putValue('BLF_MTDEJAFACT',0);
  TOBL.putValue('BLF_MTCUMULEFACT',0);
  TOBL.putValue('BLF_MTSITUATION',TOBL.getValue('GL_MONTANTHTDEV'));
  TOBL.putValue('BLF_QTEMARCHE',0);
  TOBL.putValue('BLF_QTEDEJAFACT',0);
  TOBL.putValue('BLF_QTECUMULEFACT',0);
  TOBL.putValue('BLF_QTESITUATION',TOBL.getValue('GL_QTEFACT'));
  TOBL.putValue('BLF_POURCENTAVANC',0);
  TOBL.putValue('BLF_MTPRODUCTION',0);
end;

procedure TOF_BTSAISIEAVANC.PositionneLigne(TOBL: TOB; Ligne : integer; EnAchat :Boolean=false);
var QteDejaFact,MtDejaFac,PourcentAvanc,MtHtDev,PQ: double;
		EnPA : Boolean;

  procedure ReinitAvancementOuvrage(TOBL : TOB);
  var IndiceNomen : integer;
  		TOBOUV : TOB;
      Indice : integer;
  begin
    IndiceNomen := TOBL.getValue('GL_INDICENOMEN') -1;
    if IndiceNomen < 0 then Exit;
    //FV1 : 28/07/2016
    if TobOuvrages.detail.count = 0 then exit;
    if Indicenomen > TobOuvrages.detail.count then exit;
    //
    TOBOUV := TOBOuvrages.Detail[IndiceNomen];
    for indice := 0 to TOBOUV.Detail.count -1 do
    begin
    	TOBOUV.detail[Indice].SetDouble('BLF_POURCENTAVANC',0);
      TOBOUV.detail[Indice].SetDouble('BLF_MTDEJAFACT',0);
      TOBOUV.detail[Indice].SetDouble('BLF_MTCUMULEFACT',0);
      TOBOUV.detail[Indice].SetDouble('BLF_QTEDEJAFACT',0);
      TOBOUV.detail[Indice].SetDouble('BLF_QTECUMULEFACT',0);
    end;
  end;

  procedure PositionneAncienneFacturation(TOBL : TOB);
  begin
    //
    PQ := TOBL.GEtValue('GL_PRIXPOURQTE'); if PQ = 0 then PQ := 1;
    if EnAchat then
    begin
    	MtHtdev := arrondi(TOBL.getValue('GL_QTEFACT')*TOBL.GEtValue('GL_DPA') / PQ,DEV.decimale);
    end else
    begin
    	MtHtDev := TOBL.getValue('GL_MONTANTHTDEV');
    	if MtHtdev = 0 then MtHtdev := arrondi(TOBL.getValue('GL_QTEFACT')*TOBL.GEtValue('GL_PUHTDEV') / PQ,DEV.decimale);
    end;
    if Avancement then
    begin
      // avancement ou demande d'acompte
      QteDejaFact := TOBL.GetValue('GL_QTEPREVAVANC');
      if EnAchat then
      begin
      	MtDejaFac := Arrondi(QteDejaFact* TOBL.GEtValue('GL_DPA') / PQ,DEV.decimale);
      end else
      begin
      	MtDejaFac := Arrondi(QteDejaFact* TOBL.GEtValue('GL_PUHTDEV') / PQ,DEV.decimale);
      end;
      if MtHtDev <> 0 then PourcentAvanc := Arrondi(MtDejaFac / MtHtDev * 100,2) else PourcentAvanc := 0;
      TOBL.putValue('BLF_MTMARCHE',MtHtDev);
      TOBL.putValue('BLF_MTDEJAFACT',MtDejaFac);
      TOBL.putValue('BLF_MTCUMULEFACT',MtDejaFac);
      TOBL.putValue('BLF_MTSITUATION',0);
      TOBL.putValue('BLF_QTEMARCHE',TOBL.getValue('GL_QTEFACT'));
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
      if EnAchat then
      begin
        MtDejaFac := Arrondi(QteDejaFact* TOBL.GEtValue('GL_DPA') / PQ,DEV.decimale);
      end else
      begin
        MtDejaFac := Arrondi(QteDejaFact* TOBL.GEtValue('GL_PUHTDEV') / PQ,DEV.decimale);
      end;
      if MtHtDev <> 0 then PourcentAvanc := Arrondi(MtDejaFac / MtHtDev * 100,2) else PourcentAvanc := 0;
      TOBL.putValue('BLF_MTMARCHE',MtHtDev);
      TOBL.putValue('BLF_MTDEJAFACT',MtDejaFac);
      TOBL.putValue('BLF_MTCUMULEFACT',0);
      TOBL.putValue('BLF_MTSITUATION',0);
      TOBL.putValue('BLF_QTEMARCHE',TOBL.getValue('GL_QTEFACT'));
      TOBL.putValue('BLF_QTEDEJAFACT',QteDejaFact);
      TOBL.putValue('BLF_QTECUMULEFACT',0);
      TOBL.putValue('BLF_QTESITUATION',0);
      TOBL.putValue('BLF_POURCENTAVANC',0);
      if TOBL.FieldExists('POURCENTAVANCINIT') then TOBL.putValue ('POURCENTAVANCINIT',PourcentAvanc);
      if TOBL.FieldExists('MTCUMULEFACTINIT') then TOBL.putValue ('MTCUMULEFACTINIT',MtDejaFac);
      if TOBL.FieldExists('QTECUMULEFACTINIT') then TOBL.putValue ('QTECUMULEFACTINIT',QteDejaFact);
    end;
    TOBL.putValue ('BLF_FOURNISSEUR',TOBL.GetValue('GL_FOURNISSEUR'));
    TOBL.putValue ('BLF_NATURETRAVAIL',TOBL.GetValue('GLC_NATURETRAVAIL'));
  end;

  procedure InitMontantfacturation(TOBL : TOB);
  begin
    if EnAchat then
    begin
    	TOBL.putValue('BLF_MTMARCHE',arrondi(TOBL.getValue('GL_QTEFACT')*TOBL.GEtValue('GL_DPA') / PQ,DEV.decimale));
    end else
    begin
    	TOBL.putValue('BLF_MTMARCHE',TOBL.getValue('GL_MONTANTHTDEV'));
    end;
    TOBL.putValue('BLF_MTDEJAFACT',0);
    TOBL.putValue('BLF_MTCUMULEFACT',0);
    TOBL.putValue('BLF_MTSITUATION',0);
    TOBL.putValue('BLF_QTEMARCHE',TOBL.getValue('GL_QTEFACT'));
    TOBL.putValue('BLF_QTEDEJAFACT',0);
    TOBL.putValue('BLF_QTECUMULEFACT',0);
    TOBL.putValue('BLF_QTESITUATION',0);
    TOBL.putValue('BLF_POURCENTAVANC',0);
    if TOBL.FieldExists('POURCENTAVANCINIT') then TOBL.putValue ('POURCENTAVANCINIT',0);
    if TOBL.FieldExists('MTCUMULEFACTINIT')  then TOBL.putValue ('MTCUMULEFACTINIT',0);
    if TOBL.FieldExists('QTECUMULEFACTINIT') then TOBL.putValue ('QTECUMULEFACTINIT',0);
    TOBL.putValue ('BLF_FOURNISSEUR',TOBL.GetValue('GL_FOURNISSEUR'));
    TOBL.putValue ('BLF_NATURETRAVAIL',TOBL.GetValue('GLC_NATURETRAVAIL'));
    TOBL.putValue('BLF_MTPRODUCTION',0);
  end;
begin
  TOBL.putValue('INDICE',Ligne); // numérotation de la ligne
  if TOBL.getValue('GL_TYPELIGNE')='SD' then exit;
  if TOBL.getValue('GL_TYPELIGNE')<>'ART' then BEGIN InitValeurZero(TOBL); Exit; END;
  if TOBL.getValue('BLF_NATUREPIECEG')=''  then
  begin
  	// gestion des anciennes facturations
  	if ModeCloture then
    begin
    	if ModeGeneration = 'AVA' then
      begin
        PositionneAncienneFacturation(TOBL);
      end else if ModeGeneration = 'DAC' then
      begin
        TOBL.putValue('GL_QTEPREVAVANC',0);
        TOBL.putValue('GL_QTESIT',0);
        TOBL.putValue('GL_POURCENTAVANC',0);
        PositionneAncienneFacturation(TOBL);
        ReinitAvancementOuvrage(TOBl);
      end else
      begin
        InitMontantfacturation(TOBL);
      end;
    end else
    begin
      if ((Avancement) and (TOBL.GetValue('GL_QTEPREVAVANC')<>0)) or ((not avancement) and (TOBL.GetValue('GL_QTESIT')<>0))then
      begin
        PositionneAncienneFacturation(TOBL);
      end else
      begin
        InitMontantfacturation(TOBL);
      end;
    end;
  	InitMontantSupfacturation (TOBL);
  end else
  begin
    if ModifSituation then
    begin
      REcalcMontantSupfacturation (TOBL);
    end else if ModeCloture then
    begin
      if ModeGeneration = 'DAC' then
      begin
        TOBL.putValue('GL_QTEPREVAVANC',0);
        TOBL.putValue('GL_QTESIT',0);
        TOBL.putValue('GL_POURCENTAVANC',0);
        PositionneAncienneFacturation(TOBL);
        ReinitAvancementOuvrage(TOBl);
      end;
  		InitMontantSupfacturation (TOBL);
    end else
    begin
    	if TOBL.getValue('BLF_MTSITUATION')<>0 then
      begin
      	MtDejaFac := TOBL.GetValue('BLF_MTCUMULEFACT');
        MtHTdev := TOBL.GetValue('BLF_MTMARCHE');
      	if MtHtDev <> 0 then PourcentAvanc := Arrondi(MtDejaFac / MtHtDev * 100,2) else PourcentAvanc := 0;
      	InitMontantSupfacturation (TOBL);
        if TOBL.FieldExists('POURCENTAVANCINIT') then TOBL.putValue('POURCENTAVANCINIT',Pourcentavanc);
        TOBL.putValue('BLF_MTCUMULEFACT',TOBL.GetValue('BLF_MTCUMULEFACT')+TOBL.GetValue('BLF_MTSITUATION'));
        TOBL.putValue('BLF_QTECUMULEFACT',TOBL.GetValue('BLF_QTECUMULEFACT')+TOBL.GetValue('BLF_QTESITUATION'));
      end else
      begin
      	InitMontantSupfacturation (TOBL);
      end;
    end;
  end;
  DefiniAvancementPrecPiece(TOBL);
end;

procedure TOF_BTSAISIEAVANC.DefinieGrille (GS : Thgrid);
var
    lesElements,Lelement,LaLargeur,Lalignement,LeTitre,LeNC : string;
    colonne : integer;
    flargeur,Falignement,FTitre,NC : string;
begin

	if Avancement then
  begin
    ListeChamps := 'INDICE;GLC_VOIRDETAIL;GL_CODEMARCHE;GL_TYPEARTICLE;GL_REFARTSAISIE;GL_LIBELLE;BLF_QTEMARCHE;GL_QUALIFQTEVTE;GL_PUHTDEV;BLF_MTMARCHE;'+
                   'BLF_QTEDEJAFACT;BLF_QTECUMULEFACT;BLF_POURCENTAVANC;BLF_QTESITUATION;'+
                   'BLF_MTCUMULEFACT;BLF_MTDEJAFACT;BLF_MTSITUATION;BLF_MTPRODUCTION;';
    flargeur := '2;2;13;3;8;19;9;5;9;12;9;9;9;9;12;12;12;12;';
    Falignement := 'G.0O X--;G.0O X--;G.0O X--;G.0  ---;G.0  ---;G.0  ---;D.3  -X-;C.0  ---;D.3  -X-;D.2  -X-;D.3  -X-;D.3  -X-;D.2  -X-;D.2  -X-;D.2  -X-;D.2  -X-;D.2  -X-;D.2  -X-';
    FTitre := 'N°; ;Marché;Type;Référence;Désignation;Qté Marché;Unité;P.U;Mt Marché;Qte déjà fact.;Qte Cum. fact.;% Avancement;Qte Sit.;Montant cumulé;MT Déjà fact;Montant sit.;Montant Prod.;';
    NC := '0;0;0;0;0;1;0;0;0;0;0;0;1;1;0;0;1;1;'; //definition des zones saisissable ou non

    NbCols := 18; // Nombre de colonnes dans la saisie
  end else
  begin
    ListeChamps := 'INDICE;GLC_VOIRDETAIL;GL_TYPEARTICLE;GL_REFARTSAISIE;GL_LIBELLE;BLF_QTEMARCHE;GL_QUALIFQTEVTE;GL_PUHTDEV;BLF_MTMARCHE;'+
                   'BLF_QTEDEJAFACT;BLF_POURCENTAVANC;BLF_QTESITUATION;'+
                   'BLF_MTDEJAFACT;BLF_MTSITUATION;BLF_MTPRODUCTION';
    flargeur := '2;2;3;8;19;9;5;9;12;9;9;9;12;12;12;';
    Falignement := 'G.0O X--;G.0O X--;G.0  ---;G.0  ---;G.0  ---;D.3  -X-;C.0  ---;D.3  -X-;D.2  -X-;D.3  -X-;D.2  -X-;D.2  -X-;D.2  -X-;D.2  -X-;D.2  -X-';
    FTitre := 'N°; ;Type;Référence;Désignation;Qté Marché;Unité;P.U;Mt Marché;Qte déjà fact.;% Facturation;Qte Facturation;Montant déjà fac.;Montant Facturation;Montant Prod.;';
    NC := '0;0;0;0;1;0;0;0;0;0;1;1;0;1;1;'; //definition des zones saisissable ou non

    NbCols := 15; // Nombre de colonnes dans la saisie
  end;
  GS.ColCount := NbCols;
  lesElements := ListeChamps;
  //
  Lelement := ReadtokenSt(lesElements);
  LaLargeur := ReadtokenSt(fLargeur);
  Lalignement := ReadtokenSt(Falignement);
  LeTItre := readTokenSt(Ftitre);
  LeNC := readTokenSt(NC);
  Colonne := 0;
  Repeat
    if Lelement = '' then break;
    if Lelement = 'INDICE' then
    begin
      INDICE := Colonne;
      DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
    end else if Lelement = 'GL_CODEMARCHE' then
    begin
      CODEMARCHE := Colonne;
      DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
    end else if Lelement = 'GL_TYPEARTICLE' then
    begin
      TYPEARTICLE := Colonne;
      DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
    end else if Lelement = 'GLC_VOIRDETAIL' then
    begin
      VOIRDETAIL := Colonne;
      DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
    end else if Lelement = 'GL_REFARTSAISIE' then
    begin
      CODEARTICLE := Colonne;
      DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
    end else if Lelement = 'GL_LIBELLE' then
    begin
      LIBELLE := Colonne;
      DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
    end else if Lelement = 'BLF_QTEMARCHE' then
    begin
      QTEMARCHE := Colonne;
      DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
      LNQTEMARCHE := StrToInt(LaLargeur);
    end else if Lelement = 'GL_QUALIFQTEVTE' then
    begin
      UNITE := Colonne;
      DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
      LNUNITE := StrToInt(LaLargeur);
    end else if Lelement = 'GL_PUHTDEV' then
    begin
      PUSIT := Colonne;
      DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
      LPUSIT := StrToInt(LaLargeur);
    end else if Lelement = 'BLF_MTMARCHE' then
    begin
      MTMARCHE := Colonne;
      DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
      LNMTMARCHE := StrToInt(LaLargeur);
    end else if Lelement = 'BLF_QTEDEJAFACT' then
    begin
      QTEDEJAFACT := Colonne;
      if ReajusteDossier then leNC := '1';
      DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
      LNQTEDEJAFACT := StrToInt(LaLargeur);
    end else if Lelement = 'BLF_QTECUMULEFACT' then
    begin
      if not Avancement then
      begin
        LaLargeur := '0';
      end;
      QTECUMULFACT := Colonne;
      DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
      LNQTECUMULFACT := StrToInt(LaLargeur);
    end else if Lelement = 'BLF_QTESITUATION' then
    begin
      QTESITUATION := Colonne;
      if ReajusteDossier then laLargeur := '0';
      DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
      LNQTESITUATION := StrToInt(LaLargeur);
    end else if Lelement = 'BLF_POURCENTAVANC' then
    begin
      if not Avancement then LeTitre := '% Facturation';
      POURCENTAVANC := Colonne;
      DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
    end else if Lelement = 'BLF_MTCUMULEFACT' then
    begin
      if not Avancement then LaLargeur := '0';
      MTCUMULEFACT := Colonne;
      DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
      LNMTCUMULEFACT := StrToInt(LaLargeur);
    end else if Lelement = 'BLF_MTDEJAFACT' then
    begin
      MTDEJAFACT := Colonne;
      if ReajusteDossier then leNC := '1';
      DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
      LNMTDEJAFACT := StrToInt(LaLargeur);
    end else if Lelement = 'BLF_MTSITUATION' then
    begin
      if not Avancement then LeTitre := 'Mont. Facturé';
      if ReajusteDossier then laLargeur := '0';
      MTSITUATION := Colonne;
      DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
    end else if Lelement = 'BLF_MTPRODUCTION' then
    begin
      if (ReajusteDossier) or (not ExJaiLeDroitConcept(521,false)) or (not GetparamSocSecur('SO_BTMTPRODUCTION',false))then laLargeur := '0';
      MTPRODUCTION := Colonne;
      DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
      LMTPRODUCTION := StrToInt(LaLargeur);
    end;
    //
    Lelement := ReadtokenSt(lesElements);
    LaLargeur := ReadtokenSt(fLargeur);
    Lalignement := ReadtokenSt(Falignement);
    LeTItre := readTokenSt(Ftitre);
    LeNC := readTokenSt(NC);
    inc(colonne);
    //
  until lelement = '';
  if not ModifSousDetail then
  begin
  	GS.ColWidths [VOIRDETAIL] := -1;
  end;
end;

procedure TOF_BTSAISIEAVANC.AddChampsSupAvanc(TOBL: TOB);
begin
  TOBL.AddChampSupValeur ('INDICE',0);
  TOBL.AddChampSupValeur ('UNITE',0);
  TOBL.AddChampSupValeur ('ENPA','-');
end;

procedure TOF_BTSAISIEAVANC.DefiniAttributCol(GS : THgrid; Nom: string; Colonne: integer; LaLargeur,Lalignement, Letitre, LeNC: string);
var FF,FFQ,FFM : string;
    Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep : boolean;
    dec,i : integer;

begin
  if V_PGI.OkDecQ > 0 then
  begin
    FFQ := '#0.';
    for i := 1 to V_PGI.OkDecQ - 1 do
    begin
      FFQ := FFQ + '0';
    end;
    FFQ := FFQ + '0';
  end;
  if DEV.decimale > 0 then
  begin
    FFM := '#0.';
    for i := 1 to Dev.decimale - 1 do
    begin
      FFM := FFM + '0';
    end;
    FFM := FFM + '0';
  end;
  TransAlign(Lalignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
  GS.cells[Colonne,0] := leTitre;

  // Alignement dans la colonne
  if copy(LAlignement,1,1)='G' then GS.ColAligns[Colonne] := taLeftJustify
  else if copy(LAlignement,1,1)='D' then GS.ColAligns[Colonne] := taRightJustify
  else if copy(LAlignement,1,1)='C' then GS.ColAligns[Colonne] := taCenter;
  GS.ColWidths[Colonne] := strtoint(lalargeur)*GS.Canvas.TextWidth('W');
  GS.Collengths[Colonne] := strtoint(lalargeur)*GS.Canvas.TextWidth('W');
  if (VH_GC.BTCODESPECIF <> '001') and (Nom='GL_CODEMARCHE') then
  begin
    GS.ColWidths[Colonne] := -1;
    GS.Collengths[Colonne] := 0;
    LeNC := 'O';
  end;
  if OkLib then
  begin
    GS.ColFormats[Colonne] := 'CB=' + Get_Join(Nom);
  end else if (Dec<>0) or (Sep) then
  begin
    if copy(Nom,1,7)='BLF_QTE' then
    begin
      GS.ColFormats[colonne] := FFQ+';-'+FFQ+'; ;'; { NEWPIECE }
    end else if copy(Nom,1,6)='BLF_MT' then
    begin
      GS.ColFormats[Colonne] := FFM+';-'+FFM+'; ;'; { NEWPIECE }
    end else
    begin
      GS.ColFormats[Colonne] := FF ;
    end;
  end;
  if LeNc = '0' then GS.ColEditables [Colonne]:= false;
end;

procedure TOF_BTSAISIEAVANC.AfficheLaLigne (GS : THGrid; TOBLignes : TOB; Ligne : integer);
var TOBL : TOB;
begin
  TOBL := GetTOBLigneAV (TOBlignes, Ligne); if TOBl = nil then exit;
  TOBL.PutLigneGrid (GS,Ligne,false,false,ListeChamps);
  GS.InvalidateRow(Ligne);
//  if (TOBL.GetValue('GL_TYPELIGNE') = 'RG') then GS.RowHeights [Ligne] := 0;
end;

procedure TOF_BTSAISIEAVANC.AfficheLaGrille (GS : Thgrid; TOBLignes : TOB; Depart : integer =0) ;
var Indice : integer;
begin
  GS.RowCount := TOBLignes.detail.count+1;
  For Indice := Depart to TOBLignes.detail.count -1 do
  begin
    AfficheLaLigne (GS, TOBLignes, Indice+1);
  end;
  GS.Invalidate;
end;

procedure TOF_BTSAISIEAVANC.AfficheInfos;
var Part0,Part1,Part2,Part3,Avenant : string;
begin
  LAFFAIRE := TOBSelect.getValue('GP_AFFAIRE');
  BTPCodeAffaireDecoupe (LAFFAIRE,Part0,Part1,Part2,Part3,Avenant,taModif,false);
  AFF_AFFAIRE0.text := Part0;
  AFF_AFFAIRE1.Text := Part1;
  AFF_AFFAIRE2.text := Part2;
  AFF_AFFAIRE3.text := Part3;
  AFF_AVENANT.text := Avenant;
  AFF_AFFAIRE.TEXT := Laffaire;
  //
  AFF_TIERS.Text := TOBSelect.getValue('GP_TIERS') ;
  AFF_DESIGNATION.Caption := TOBSelect.getValue('LIBELLEAFFAIRE');
  LIBTIERS.Caption := TOBSelect.getValue('LIBELLETIERS');
  //
//  BPF_MTSITUATION.caption := '0,00';
end;

procedure TOF_BTSAISIEAVANC.SetEventsGrid (GS : THgrid; Etat : boolean);
begin
  if Etat then
  begin
    GS.OnCellEnter := GSCellEnter;
    GS.OnCellExit := GSCellExit;
    GS.OnRowEnter := GSRowEnter;
    GS.OnRowExit := GSRowExit;
    GS.PostDrawCell := GSPostDrawCell;
    GS.GetCellCanvas := GSGetCellCanvas;
    GS.OnColumnWidthsChanged := GSColumnsWithChanging;
  end else
  begin
    GS.OnCellEnter := Nil;
    GS.OnCellExit := Nil;
    GS.OnRowEnter := Nil;
    GS.OnRowExit := Nil;
    GS.PostDrawCell := nil;
    GS.GetCellCanvas := nil;
    GS.OnColumnWidthsChanged := GSColumnsWithChanging;
  end;
end;

procedure TOF_BTSAISIEAVANC.SetEvents (GS : THgrid);
begin
  GS.OnKeyDown := FormKeyDown;
  BTDescriptif.onclick := BTDescriptifClick;
  FMTSITUATION.OnChange := ChangeMontantSituation;
  BAVANC.onclick := BAVANCClick;
  VOIRMT.OnClick := VoirMtCLick;
  VOIRQTE.OnClick := VoirQteClick;
  ToolWin.onclose := ToolWinClose;
  //
  MnDefaireLig.OnClick := DefaireLig;
  MnDefairePiece.onClick := DefairePiece;
  MnDefaireTous.OnClick := DefaireTous;
  MnAvancement.onclick := BAVANCClick;
  MnDefaireParag.onClick := DefaireParag;
  //


  SetEventsGrid (GS,True);
end;


procedure TOF_BTSAISIEAVANC.GSCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
  if Action = taConsult then Exit;
  ZoneSuivanteOuOk(GS,ACol, ARow, Cancel);
  if not Cancel then
  begin
    stprev := GS.Cells [Acol,Arow];
  	fCurrentOnglet.stPrev := stprev;
  end;
end;

function TOF_BTSAISIEAVANC.ZoneAccessible( Grille : THgrid; var ACol, ARow : integer) : boolean;
var TOBL : TOB;
begin
  result := true;
  if (Grille.ColWidths[acol] = 0) or (not Grille.ColEditables[acol] ) then BEGIN result := false; exit; END;
  if Arow < GS.FixedRows then exit;
  if Acol < GS.fixedCols then exit;
  if GS.RowHeights[ARow] <= 0 then begin result := false;Exit;  end;
  TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes, Arow); if TOBL = nil then BEGIN result := false; exit; END;
  if (not IsGerableAvanc (TOBL)) and (Acol <> LIBELLE) then BEGIN result := false; exit; END;
  if (TOBL.GetValue('BLF_MTSITUATION')=0) and (Acol = MTPRODUCTION) then BEGIN result := false; exit; END;
end;

procedure TOF_BTSAISIEAVANC.ZoneSuivanteOuOk(Grille : THgrid;var ACol, ARow : integer;var  Cancel : boolean);
var Sens, ii, Lim: integer;
  OldEna, ChgLig, ChgSens: boolean;
begin
  OldEna := Grille.SynEnabled;
  Grille.SynEnabled := False;
  Sens := -1;
  ChgLig := (Grille.Row <> ARow);
  ChgSens := False;
  if Grille.Row > ARow then Sens := 1 else if ((Grille.Row = ARow) and (ACol <= Grille.Col)) then Sens := 1;
  ACol := Grille.Col;
  ARow := Grille.Row;
  ii := 0;
  while not ZoneAccessible(Grille,ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if Sens = 1 then
    begin
      // Modif BTP
      Lim := Grille.RowCount ;
      // ---
      if ((ACol = Grille.ColCount - 1) and (ARow >= Lim)) then
      begin
        Sens := -1;
        continue;
      end;
      if ChgLig then
      begin
        ACol := Grille.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < Grille.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        ACol := Grille.FixedCols;
      end;
    end else
    begin
      if ((ACol = Grille.FixedCols) and (ARow = 1)) then
      begin
        if ChgSens then Break else
        begin
          Sens := 1;
          Continue;
        end;
      end;
      if ChgLig then
      begin
        ACol := Grille.ColCount;
        ChgLig := False;
      end;
      if ACol > Grille.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := Grille.ColCount - 1;
      end;
    end;
  end;
  Grille.SynEnabled := OldEna;
end;

function TOF_BTSAISIEAVANC.FindTOBligne (Onglet : Tonglet ; numPiece,numOrdre,uniqueBlo : integer) : integer;
var Indice : integer;
		IndnumOrdre,IndUniqueBLo,IndNumPiece : integer;
    TOBL : TOB;
begin
	IndNumpiece := 0;
	IndNumordre := 0;
	IndUniqueBlo := 0;
  result := -1;
  for Indice := 0 to Onglet.TOBlignes.detail.count -1 do
  begin
  	TOBl := Onglet.TOBlignes.detail[indice];
    if Indice = 0 then
    begin
    	IndNumPiece := TOBL.GetNumChamp('GL_NUMERO');
    	IndNumOrdre := TOBL.GetNumChamp('GL_NUMORDRE');
    	IndUniqueBLo := TOBL.GetNumChamp('UNIQUEBLO');
    end;
    if (TOBL.GetValeur(IndNumpiece) = NumPiece) and (TOBL.GetValeur(IndNumOrdre) = NumOrdre) and (TOBL.GetValeur(IndUniqueBLo) = uniqueBlo) then
    begin
    	result := Indice;
      break;
    end;
  end;
end;

function TOF_BTSAISIEAVANC.FindlignesOuvSaisie(Onglet: Tonglet; TOBL : TOB): TlisteTOB;
var IndiceDep,Indice : integer;
		IndiceNomen : integer;
    TOBD : TOB;
begin
	result := TlisteTOB.create;
  IndiceNomen := TOBL.getValue('GL_INDICENOMEN');
  IndiceDep := TOBL.GetIndex+1;
  for Indice := IndiceDep to Onglet.TOBlignes.detail.count -1 do
  begin
    TOBD := Onglet.TOBlignes.detail[Indice];
    if TOBD.getValue('GL_INDICENOMEN') <> IndiceNomen then break;
    result.add(TOBD);
  end;
end;

procedure TOF_BTSAISIEAVANC.SetOnglet (Onglet : Tonglet);
begin
  fCurrentOnglet := Onglet;
  GS := fCurrentOnglet.GS; // pour le cas ou
  stPrev := fCurrentOnglet.stPrev;
  dMtSItuation := fCurrentOnglet.MtSituation;
  TVDevis := fCurrentOnglet.fTVDevis;
  TVDevis.visible := true;
  FMTSITUATION.Value := dMtsituation;
  BPF_MTSITUATION.caption := FMTSITUATION.Text;
	BPF_MTSITUATION.Refresh;
  BPF_MTSITUATION.visible := (VH_GC.BTCODESPECIF<>'001') or (Pos(fCurrentOnglet.name,'TSGLOBALE;TS0;')>0);
  THLabel(GetControl('TMTSITUATION')).visible := (VH_GC.BTCODESPECIF<>'001') or (Pos(fCurrentOnglet.name,'TSGLOBALE;TS0;')>0);
  THLabel(GetControl('TDEVISE')).visible := (VH_GC.BTCODESPECIF<>'001') or (Pos(fCurrentOnglet.name,'TSGLOBALE;TS0;')>0);
end;

procedure TOF_BTSAISIEAVANC.AppliqueAvancGeneral (TOBD : TOB;Avanc : double;WithDeduction:boolean;WithAjout: Boolean;WithCalcul:boolean);
var NumOrdre,UniqueBlo,NumPiece : integer;
		LastOnglet : TOnglet;
    Indice : integer;
    cancel : boolean;
    TOBL : TOB;
begin
  if IsOuvragePartiel(TOBD) then
  begin
    AppliqueAvancDetailGeneral (TOBD,Avanc,WithDeduction,WithAjout,WithCalcul);
  end else
  begin
    NumPiece := TOBD.getvalue('GL_NUMERO');
    NumOrdre := TOBD.getvalue('GL_NUMORDRE');
    UniqueBlo := TOBD.getvalue('UNIQUEBLO');
    lastOnglet := fCurrentOnglet;
    SetOnglet(FOngletGlobal);
    indice := FindTOBligne (FongletGlobal,Numpiece,Numordre,uniqueBlo);
    if indice < 0 then exit;
    TOBL := GetTOBLigneAV(fCurrentOnglet.TOBlignes ,indice+1);
    if TOBL <> nil then
    begin
      if WithDeduction then CumuleLigne(TOBL,'-');
      TOBL.putValue('BLF_POURCENTAVANC',Avanc);
      ChangePourcentAvanc (TOBL,FOngletGlobal,Cancel);
      if WithAjout then CumuleLigne(TOBL);
      AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,Indice+1);
      if WithCalcul then CalculeLesSousTotaux;
    end;
    SetOnglet(lastOnglet);
  end;
end;

procedure TOF_BTSAISIEAVANC.AppliqueQteFact (TOBD : TOB;qte : double);
var NumOrdre,UniqueBlo,NumPiece : integer;
		LastOnglet : TOnglet;
    Indice : integer;
    cancel : boolean;
    TOBL : TOB;
begin
  if IsOuvragePartiel(TOBD) then
  begin
    AppliqueQteDetailGeneral (TOBD);
  end else
  begin
    NumOrdre := TOBD.getvalue('GL_NUMORDRE');
    UniqueBlo := TOBD.getvalue('UNIQUEBLO');
    NumPiece := TOBD.getvalue('GL_NUMERO');
    lastOnglet := fCurrentOnglet;
    SetOnglet(FOngletGlobal);
    indice := FindTOBligne (FongletGlobal,NumPiece,Numordre,uniqueBlo);
    if indice < 0 then exit;
    TOBL := GetTOBLigneAV(fCurrentOnglet.TOBlignes ,indice+1);
    if TOBL <> nil then
    begin
      TOBL.putValue('BLF_QTEDEJAFACT',Qte);
      ChangeQteFact (TOBL,FOngletGlobal,Cancel);
      AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,indice+1);
      CalculeLesSousTotaux;
    end;
    SetOnglet(lastOnglet);
  end;
end;

procedure TOF_BTSAISIEAVANC.AppliqueQteGeneral (TOBD : TOB;qte : double);
var NumOrdre,UniqueBlo,NumPiece : integer;
		LastOnglet : TOnglet;
    Indice : integer;
    cancel : boolean;
    TOBL : TOB;
begin
  if IsOuvragePartiel(TOBD) then
  begin
    AppliqueQteDetailGeneral (TOBD);
  end else
  begin
    NumOrdre := TOBD.getvalue('GL_NUMORDRE');
    UniqueBlo := TOBD.getvalue('UNIQUEBLO');
    NumPiece := TOBD.getvalue('GL_NUMERO');
    lastOnglet := fCurrentOnglet;
    SetOnglet(FOngletGlobal);
    indice := FindTOBligne (FongletGlobal,NumPiece,Numordre,uniqueBlo);
    if indice < 0 then exit;
    TOBL := GetTOBLigneAV(fCurrentOnglet.TOBlignes ,indice+1);
    if TOBL <> nil then
    begin
      CumuleLigne(TOBL,'-');
      TOBL.putValue('BLF_QTESITUATION',Qte);
      ChangeQteSituation (TOBL,FOngletGlobal,Cancel);
      CumuleLigne(TOBL);
      AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,indice+1);
      CalculeLesSousTotaux;
    end;
    SetOnglet(lastOnglet);
  end;
end;

procedure TOF_BTSAISIEAVANC.AppliqueQteDetailGeneral (TOBD : TOB);
var ListeDet : TlisteTOB;
    TOBL : TOB;
    Indice : integer;
    Avanc : Double;
begin
  Avanc := TOBD.GetValue('BLF_POURCENTAVANC');
	ListeDet := FindlignesOuvSaisie (fCurrentOnglet,TOBD);
  if ListeDet.count > 0 then   // lignes de détail trouvé
  begin
  	for Indice := 0 to ListeDet.count -1 do
    begin
    	TOBL := ListeDet.Items [Indice];
      AppliqueAvancGeneral (TOBL,Avanc);
      //
    	AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,TOBL.getvalue('INDICE'));
    end;
  end;
  ListeDet.free;
end;

procedure TOF_BTSAISIEAVANC.AppliqueAvancDetailGeneral(TOBD : TOB;Avanc : double;WithDeduction:boolean;WithAjout:Boolean;WithCalcul:boolean);
var ListeDet : TlisteTOB;
    TOBL : TOB;
    Indice : integer;
begin
	ListeDet := FindlignesOuvSaisie (fCurrentOnglet,TOBD);
  if ListeDet.count > 0 then   // lignes de détail trouvé
  begin
  	for Indice := 0 to ListeDet.count -1 do
    begin
    	TOBL := ListeDet.Items [Indice];
      AppliqueAvancGeneral (TOBL,Avanc,WithDeduction,WithAjout,WithCalcul);
      //
    	AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,TOBL.getvalue('INDICE'));
    end;
  end;
  ListeDet.free;
end;

procedure TOF_BTSAISIEAVANC.AppliqueMtDetailGeneral (TOBD : TOB);
var ListeDet : TlisteTOB;
    TOBL : TOB;
    Indice : integer;
    Avanc : double;
begin
  Avanc := TOBD.GetValue('BLF_POURCENTAVANC');
	ListeDet := FindlignesOuvSaisie (fCurrentOnglet,TOBD);
  if ListeDet.count > 0 then   // lignes de détail trouvé
  begin
  	for Indice := 0 to ListeDet.count -1 do
    begin
    	TOBL := ListeDet.Items [Indice];
      AppliqueAvancGeneral (TOBL,Avanc);
      //
    	AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,TOBL.getvalue('INDICE'));
    end;
  end;
  ListeDet.free;
end;

procedure TOF_BTSAISIEAVANC.AppliqueMtGeneral (TOBD : TOB;Mt : double);
var NumOrdre,UniqueBlo,NumPiece : integer;
		LastOnglet : TOnglet;
    Indice : integer;
    cancel : boolean;
    TOBL : TOB;
begin
  if IsOuvragePartiel(TOBD) then
  begin
    AppliqueMtDetailGeneral (TOBD);
  end else
  begin
    NumOrdre := TOBD.getvalue('GL_NUMORDRE');
    UniqueBlo := TOBD.getvalue('UNIQUEBLO');
    NumPiece := TOBD.getvalue('GL_NUMERO');
    lastOnglet := fCurrentOnglet;
    SetOnglet(FOngletGlobal);
    indice := FindTOBligne (FongletGlobal,NumPiece,Numordre,uniqueBlo);
    if indice < 0 then exit;
    TOBL := GetTOBLigneAV(fCurrentOnglet.TOBlignes ,indice+1);
    if TOBL <> nil then
    begin
      CumuleLigne(TOBL,'-');
      ChangeMTSituationAvanc (Mt,TOBL,FOngletGlobal,Cancel);
      CumuleLigne(TOBL);
      AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,Indice+1);
      CalculeLesSousTotaux;
    end;
    SetOnglet(lastOnglet);
  end;
end;

function TOF_BTSAISIEAVANC.FindLigneInOnglet (NumPiece,Numordre,uniqueBlo : integer; var ligne : integer) : Tonglet;
var Indice : integer;
begin
	result := nil;
  for Indice := 1 to fOnglets.count -1 do
  begin
  	if fOnglets.items[indice] = fCurrentOnglet then continue;
    ligne := FindTOBligne (FOnglets.Items[Indice],NumPiece,Numordre,uniqueBlo);
    if ligne >= 0 then
    begin
    	result := FOnglets.Items[Indice];
      break;
    end;
  end;
end;

procedure TOF_BTSAISIEAVANC.FindLesOnglets (NumPiece,Numordre,uniqueBlo : integer; var ligne : integer;TheRef : TOnglet; LaListe : TStringList);
var Indice : integer;
begin
  for Indice := 1 to fOnglets.count -1 do
  begin
  	if fOnglets.items[indice] = TheRef then continue;
    ligne := FindTOBligne (FOnglets.Items[Indice],NumPiece,Numordre,uniqueBlo);
    if ligne >= 0 then
    begin
    	LaListe.Add(FOnglets.Items[Indice].Name);
    end;
  end;
  if TheRef <> FOngletGlobal then
  begin
    ligne := FindTOBligne (FOngletGlobal,NumPiece,Numordre,uniqueBlo);
    if ligne >= 0 then
    begin
    	LaListe.Add(FOngletGlobal.Name);
    end;
  end;
end;

procedure TOF_BTSAISIEAVANC.AppliqueAvancSDetailFournisseur (TOBD : TOB;Avanc : double ;WithDeduction,WithAjout,WithCalcul : boolean);
var NumOrdre,UniqueBlo,NumPiece : integer;
		LastOnglet,TheOnglet : TOnglet;
    Indice,II,IndiceNomen : integer;
    cancel : boolean;
    TOBL,TOBO,TOBOD : TOB;
begin
  IndiceNomen := TOBD.getvalue('GL_INDICENOMEN');
  if IndiceNomen = 0 then exit;

  //FV1 : 28/07/2016
  if TOBOuvrages.detail.count = 0 then exit;
  if (indiceNomen-1 >= TOBOuvrages.detail.count - 1) then exit;

  TOBO := TOBOuvrages.Detail[indiceNomen-1];
  for II := 0 to TOBO.detail.count -1 do
  begin
 		TOBOD := TOBO.detail[II];
    NumOrdre := TOBOD.getvalue('BLO_NUMLIGNE');
    NumPiece := TOBOD.getvalue('BLO_NUMERO');
    UniqueBlo := TOBOD.getvalue('BLO_UNIQUEBLO');
    lastOnglet := fCurrentOnglet;
    indice := -1;
    TheOnglet := FindLigneInOnglet(NumPiece,NumOrdre,UniqueBlo,indice);
    if TheOnglet <> nil then
    begin
      SetOnglet(TheOnglet);
      TOBL := GetTOBLigneAV(fCurrentOnglet.TOBlignes ,indice+1);
      if TOBL <> nil then
      begin
        if WithDeduction then CumuleLigne(TOBL,'-');
        TOBL.putValue('BLF_POURCENTAVANC',avanc);
        ChangePourcentAvanc (TOBL,TheOnglet,Cancel);
        if WithAjout then CumuleLigne(TOBL);
        AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,indice+1);
        if WithCalcul then CalculeLesSousTotaux;
      end;
    end;
    SetOnglet(lastOnglet);
  end;
end;

procedure TOF_BTSAISIEAVANC.AppliqueAvancFournisseur (TOBD : TOB;Avanc : double;WithDeduction:boolean;WithAjout : Boolean;WithCalcul:boolean);
var NumOrdre,UniqueBlo,NumPiece : integer;
		LastOnglet,TheOnglet : TOnglet;
    Indice,IndInOnglet : integer;
    cancel : boolean;
    TOBL : TOB;
    NatureTravail : Double;
    laListe : TStringList;
begin
  NatureTravail := VALEUR(TOBD.GetValue('GLC_NATURETRAVAIL'));
  if IsOuvrage(TOBD) and (fCurrentOnglet=FOngletGlobal) and (NatureTravail > 10) then
  begin
  	AppliqueAvancSousDetail (TOBD,fCurrentOnglet,WithDeduction,WithAjout,WithCalcul );
    Exit;
  end;
  NumOrdre := TOBD.getvalue('GL_NUMORDRE');
  NumPiece := TOBD.getvalue('GL_NUMERO');
  UniqueBlo := TOBD.getvalue('UNIQUEBLO');
  lastOnglet := fCurrentOnglet;
  if IsOuvrage(TOBD) and (fCurrentOnglet=FOngletGlobal) and (NatureTravail = 10) then
  begin
  	laListe := TStringList.Create;

		indice := fCurrentOnglet.fGS.Row -1;
    FindLesOnglets(NumPiece,NumOrdre,0,indice,fCurrentOnglet,LaListe);
    if laListe.Count > 0 then
    begin
      for Indice := 0 to laListe.Count -1 do
      begin
        IndInOnglet := -1;
        TheOnglet := fOnglets.findOnglet(laListe.strings[Indice]);
        if TheOnglet <> nil then
        begin
          IndInOnglet := FindTOBligne(TheOnglet,Numpiece,Numordre,0);
          if IndInOnglet >0 then
          begin
      			SetOnglet(TheOnglet);
            TOBL := GetTOBLigneAV(TheOnglet.TOBlignes ,IndInOnglet+1);
            if TOBL <> nil then
            begin
              if WithDeduction then CumuleLigne(TOBL,'-');
              TOBL.putValue('BLF_POURCENTAVANC',avanc);
              ChangePourcentAvanc (TOBL,TheOnglet,Cancel);
              if WithAjout then CumuleLigne(TOBL);
              AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,IndInOnglet+1);
              if WithCalcul then CalculeLesSousTotaux;
            end;
          end;
        end;
      end;
    end;
    laListe.Free;
  end else
  begin
    indice := -1;
    TheOnglet := FindLigneInOnglet(NumPiece,NumOrdre,UniqueBlo,indice);
    if TheOnglet <> nil then
    begin
      SetOnglet(TheOnglet);
      TOBL := GetTOBLigneAV(fCurrentOnglet.TOBlignes ,indice+1);
      if TOBL <> nil then
      begin
        if WithDeduction then CumuleLigne(TOBL,'-');
        TOBL.putValue('BLF_POURCENTAVANC',avanc);
        ChangePourcentAvanc (TOBL,TheOnglet,Cancel);
        if WithAjout then CumuleLigne(TOBL);
        AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,indice+1);
        if WithCalcul then CalculeLesSousTotaux;
      end;
    end else
    begin
      // gestion ou l'on modifie l'avancement d'un ouvrage au niveau général --> application de la modification sur les sous détails
      AppliqueAvancSDetailFournisseur (TOBD,Avanc,WithDeduction,WithAjout,WithCalcul);
    end;
  end;

  SetOnglet(lastOnglet);
end;

procedure TOF_BTSAISIEAVANC.AppliqueQteFactFour (TOBD : TOB;Qte : double);
var NumOrdre,UniqueBlo ,NumPiece: integer;
		LastOnglet,TheOnglet : TOnglet;
    Indice : integer;
    cancel : boolean;
    TOBL : TOB;
    NatureTravail : double;
begin
  NatureTravail := VALEUR(TOBD.GetValue('GLC_NATURETRAVAIL'));
  if IsOuvrage(TOBD) and (fCurrentOnglet=FOngletGlobal) and (NatureTravail > 10) then
  begin
  	AppliqueAvancSousDetail (TOBD,fCurrentOnglet);
    Exit;
  end;
  NumOrdre := TOBD.getvalue('GL_NUMORDRE');
  UniqueBlo := TOBD.getvalue('UNIQUEBLO');
  NumPiece := TOBD.getvalue('GL_NUMERO');
  lastOnglet := fCurrentOnglet;
  indice := -1;
  TheOnglet := FindLigneInOnglet(NumPiece,NumOrdre,UniqueBlo,indice);
  if TheOnglet = nil then exit;
  SetOnglet(TheOnglet);
  TOBL := GetTOBLigneAV(fCurrentOnglet.TOBlignes ,indice+1);
  if TOBL <> nil then
  begin
    CumuleLigne(TOBL,'-');
  	NatureTravail := VALEUR(TOBL.GetValue('GLC_NATURETRAVAIL'));
    if (NatureTravail > 0) and (NatureTravail < 10 ) then
    begin
      TOBL.putValue('BLF_QTEDEJAFACT',Qte);
      ChangeQteFact (TOBL,TheOnglet,Cancel);
    end else
    begin
      TOBL.putValue('BLF_POURCENTAVANC',TOBD.GetValue('BLF_POURCENTAVANC'));
      ChangePourcentAvanc (TOBL,TheOnglet,Cancel);
    end;
    CumuleLigne(TOBL);
    AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,indice+1);
    CalculeLesSousTotaux;
  end;
  SetOnglet(lastOnglet);
end;

procedure TOF_BTSAISIEAVANC.AppliqueQteFournisseur (TOBD : TOB;Qte : double);
var NumOrdre,UniqueBlo ,NumPiece: integer;
		LastOnglet,TheOnglet : TOnglet;
    Indice : integer;
    cancel : boolean;
    TOBL : TOB;
    NatureTravail : double;
begin
  NatureTravail := VALEUR(TOBD.GetValue('GLC_NATURETRAVAIL'));
  if IsOuvrage(TOBD) and (fCurrentOnglet=FOngletGlobal) and (NatureTravail > 10) then
  begin
  	AppliqueAvancSousDetail (TOBD,fCurrentOnglet);
    Exit;
  end;
  NumOrdre := TOBD.getvalue('GL_NUMORDRE');
  UniqueBlo := TOBD.getvalue('UNIQUEBLO');
  NumPiece := TOBD.getvalue('GL_NUMERO');
  lastOnglet := fCurrentOnglet;
  indice := -1;
  TheOnglet := FindLigneInOnglet(NumPiece,NumOrdre,UniqueBlo,indice);
  if TheOnglet = nil then exit;
  SetOnglet(TheOnglet);
  TOBL := GetTOBLigneAV(fCurrentOnglet.TOBlignes ,indice+1);
  if TOBL <> nil then
  begin
    CumuleLigne(TOBL,'-');
  	NatureTravail := VALEUR(TOBL.GetValue('GLC_NATURETRAVAIL'));
    if (NatureTravail > 0) and (NatureTravail < 10 ) then
    begin
      TOBL.putValue('BLF_QTESITUATION',Qte);
      ChangeQteSituation (TOBL,TheOnglet,Cancel);
    end else
    begin
      TOBL.putValue('BLF_POURCENTAVANC',TOBD.GetValue('BLF_POURCENTAVANC'));
      ChangePourcentAvanc (TOBL,TheOnglet,Cancel);
    end;
    CumuleLigne(TOBL);
    AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,indice+1);
    CalculeLesSousTotaux;
  end;
  SetOnglet(lastOnglet);
end;

procedure TOF_BTSAISIEAVANC.AppliqueMtFournisseur (TOBD : TOB;Mt : double);
var NumOrdre,UniqueBlo,NumPiece : integer;
		LastOnglet,TheOnglet : TOnglet;
    Indice : integer;
    cancel : boolean;
    TOBL : TOB;
    NatureTravail : double;
begin
  NatureTravail := VALEUR(TOBD.GetValue('GLC_NATURETRAVAIL'));
  if IsOuvrage(TOBD) and (fCurrentOnglet=FOngletGlobal) and (NatureTravail > 10) then
  begin
  	AppliqueAvancSousDetail (TOBD,fCurrentOnglet);
    Exit;
  end;
  NumOrdre := TOBD.getvalue('GL_NUMORDRE');
  UniqueBlo := TOBD.getvalue('UNIQUEBLO');
  NumPiece := TOBD.getvalue('GL_NUMERO');
  lastOnglet := fCurrentOnglet;
  indice := -1;
  TheOnglet := FindLigneInOnglet(NumPiece,NumOrdre,UniqueBlo,indice);
  if TheOnglet = nil then exit;
  SetOnglet(TheOnglet);
  TOBL := GetTOBLigneAV(fCurrentOnglet.TOBlignes ,indice+1);
  if TOBL <> nil then
  begin
    CumuleLigne(TOBL,'-');
    {
    TOBL.putValue('BLF_MTSITUATION',Mt);
    ChangeMTSituationAvanc (TOBL,TheOnglet,Cancel);
    }
    TOBL.putValue('BLF_POURCENTAVANC',TOBD.GetValue('BLF_POURCENTAVANC'));
    ChangePourcentAvanc (TOBL,TheOnglet,Cancel);
    CumuleLigne(TOBL);
    AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,indice+1);
    CalculeLesSousTotaux;
  end;
  SetOnglet(lastOnglet);
end;

procedure TOF_BTSAISIEAVANC.GSCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var TOBL : TOB;
		TypeLigne : string;
    savcol,savrow : integer;
begin
  if Action = taConsult then Exit;
  SavCol := GS.col;
  SavRow := GS.row;
  if (stPrev <> GS.cells[Acol,Arow]) or (GS.cells[Acol,Arow]= '') then
  begin
    TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,Arow);
    //
    TypeLigne := TOBL.GetValue('GL_TYPELIGNE');
    if ACOL = QTEDEJAFACT then
    begin
      if pos(TypeLigne,'ART;SD;OUP')>0 then
      begin
        if ReajusteDossier then
        begin
          TOBL.putValue('BLF_QTEDEJAFACT',Valeur(GS.cells[Acol,Arow]));
          ChangeQteFact (TOBL,fCurrentOnglet,Cancel);
          AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,Arow);
          CalculeLesSousTotaux;
          GS.Invalidate;
          if fCurrentOnglet <> FOngletGlobal then
          begin
            AppliqueQteFact (TOBL,Valeur(GS.cells[Acol,Arow]));
          end else
          begin
            AppliqueQteFactFour(TOBL,Valeur(GS.cells[Acol,Arow]));
          end;
        end else
        begin
          // pour etre sur de ne rien faire
      	  GS.cells [Acol,Arow] := stprev;
        end;
      end else
      begin
      	GS.cells [Acol,Arow] := stprev;
      end;
    end else if ACol = QTESITUATION  then  (* Uniquement dans le cas de facturation d'avancement *)
    Begin
      if pos(TypeLigne,'ART;SD;OUP')>0 then
      begin
        if ReajusteDossier then
        begin
          //  Pour etre sur de ne rien faire dans ce cas
      	  GS.cells [Acol,Arow] := stprev;
        end else
        begin
        if ChangeQteSituationOk (Arow,Valeur(GS.cells[Acol,Arow])) then
        begin
          CumuleLigne(TOBL,'-');
          TOBL.putValue('BLF_QTESITUATION',Valeur(GS.cells[Acol,Arow]));
          ChangeQteSituation (TOBL,fCurrentOnglet,Cancel);
          CumuleLigne(TOBL);
          AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,Arow);
          CalculeLesSousTotaux;
          GS.Invalidate;
          if fCurrentOnglet <> FOngletGlobal then
          begin
          	AppliqueQteGeneral (TOBL,Valeur(GS.cells[Acol,Arow]));
          end else
          begin
          	AppliqueQteFournisseur (TOBL,Valeur(GS.cells[Acol,Arow]));
          end;
        end else
        begin
          GS.cells [Acol,Arow] := stprev;
          AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,Arow);
        end;
        end;
      end else
      begin
      	GS.cells [Acol,Arow] := stprev;
      end;
    end else if ACol = POURCENTAVANC  then
    Begin
      if pos(TypeLigne,'ART;SD;OUP')>0 then
      begin
        if ReajusteDossier then
        begin
          TOBL.putValue('BLF_POURCENTAVANC',Valeur(GS.cells[Acol,Arow]));
          ChangePourcentAvanc (TOBL,fCurrentOnglet,Cancel);
          AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,Arow);
          CalculeLesSousTotaux;
          GS.Invalidate;
          if fCurrentOnglet <> FOngletGlobal then
          begin
            AppliqueAvancGeneral (TOBL,Valeur(GS.cells[Acol,Arow]));
          end else
          begin
            AppliqueAvancFournisseur (TOBL,Valeur(GS.cells[Acol,Arow]));
          end;
        end else
        begin
        if ChangePourcentOk (Arow,Valeur(GS.cells[Acol,Arow])) then
        begin
          CumuleLigne(TOBL,'-');
          TOBL.putValue('BLF_POURCENTAVANC',Valeur(GS.cells[Acol,Arow]));
          ChangePourcentAvanc (TOBL,fCurrentOnglet,Cancel);
          CumuleLigne(TOBL);
          AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,Arow);
          CalculeLesSousTotaux;
          GS.Invalidate;
          if fCurrentOnglet <> FOngletGlobal then
          begin
          	AppliqueAvancGeneral (TOBL,Valeur(GS.cells[Acol,Arow]));
          end else
          begin
          	AppliqueAvancFournisseur (TOBL,Valeur(GS.cells[Acol,Arow]));
          end;
        end else
        begin
          GS.cells [Acol,Arow] := stprev;
          AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,Arow);
        end;
        end;
      end else
      begin
      	GS.cells [Acol,Arow] := stprev;
      end;
    end else if ACol = MTSITUATION  then
    Begin
      if pos(TypeLigne,'ART;SD;OUP')>0 then
      begin
        if ReajusteDossier then
        begin
          //
      	  GS.cells [Acol,Arow] := stprev;
        end else
        begin
        if ChangeMTSituationok (Arow,Valeur(GS.cells[Acol,Arow])) then
        begin
          CumuleLigne(TOBL,'-');
            ChangeMTSituationAvanc (Valeur(GS.cells[Acol,Arow]),TOBL,fCurrentOnglet,Cancel);
          CumuleLigne(TOBL);
          AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,Arow);
          CalculeLesSousTotaux;
          GS.Invalidate;
          if fCurrentOnglet <> FOngletGlobal then
          begin
          	AppliqueMtGeneral (TOBL,Valeur(GS.cells[Acol,Arow]));
          end else
          begin
          	AppliqueMtFournisseur (TOBL,Valeur(GS.cells[Acol,Arow]));
          end;
        end else
        begin
          GS.cells [Acol,Arow] := stprev;
          AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,Arow);
        end;
        end;
      end else
      begin
      	GS.cells [Acol,Arow] := stprev;
      end;
    end else if ACol = MTPRODUCTION  then
    Begin
      if (TypeLigne='ART') and (TOBL.GetValue('BLF_MTSITUATION')<>0) then
      begin
        if ReajusteDossier then
        begin
          //
      	  GS.cells [Acol,Arow] := stprev;
    end else
    begin
          TOBL.putValue('BLF_MTPRODUCTION',Valeur(GS.cells[Acol,Arow]));
          AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,Arow);
          CalculeLesSousTotaux;
          GS.Invalidate;
        end;
      end else
      begin
      GS.cells [Acol,Arow] := stprev;
      end;
    end else
    begin
      GS.cells [Acol,Arow] := stprev;
      AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,Arow);
    end;
    if not cancel then
    begin
			GS.col := SavCol;
      GS.row := SavRow;
      stPrev := GS.cells[Acol,Arow];
  		fCurrentOnglet.stPrev := stprev;
    end;
  end;

end;

procedure TOF_BTSAISIEAVANC.GSPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var ARect: TRect;
    TOBL: TOB;
    Numgraph : integer;
    thetext : string;
  	PosX,PosY : integer;
    OkOk : boolean;
    NatureTravail : double;
begin
  if GS.RowHeights[ARow] <= 0 then Exit;
  TOBL := nil;
  if (Arow >= GS.FixedRows) then
  begin
    TOBL := GetTOBLigneAV(fCurrentOnglet.TOBlignes,ARow);
  end;
  if TOBL = nil then Exit;
  Naturetravail := VALEUR(TOBL.GetValue('GLC_NATURETRAVAIL'));
  ARect := GS.CellRect(ACol, ARow);
  GS.Canvas.Pen.Style := psSolid;
  GS.Canvas.Pen.Color := clgray;
  GS.Canvas.Brush.Style := BsSolid;
  if (Acol = VOIRDETAIL) and (Arow >= GS.fixedRows) then
  begin
    Canvas.FillRect(ARect);
    if isOuvrage(TOBL) and (not IsSousDetail(TOBL))  then
    begin
      Okok := true;
    end;
    (*
  	Okok := (isOuvrage(TOBL)) and
    			  (NatureTravail < 10) and
    			  (not IsOuvragePartiel(TOBL)) and
            (not IsSousDetail(TOBL)) and
            (TOBL.GetValue('GL_TYPEARTICLE')<>'');
    *)
  	if (okok ) then
    begin
      PosX := ((Arect.left+Arect.Right  - ImgListPlusMoins.Width) div 2);
      posY := ((Arect.Top+Arect.Bottom - ImgListPlusMoins.Height) div 2)-1;
      if TOBL.GetValue('GLC_VOIRDETAIL')='X' then
      begin
      	ImgListPlusMoins.Draw(CanVas, PosX, posY, 1);
      end else
      begin
      	ImgListPlusMoins.Draw(CanVas, PosX, posY, 0);
      end;
    end;
  end else if (Acol = TYPEARTICLE) and (Arow >= GS.fixedRows) then
  begin
    if TOBL = nil then Exit;
    Canvas.FillRect(ARect);
    NumGraph := RecupTypeGraph(TOBL);
    if NumGraph >= 0 then
    begin
      ImTypeArticle.DrawingStyle := dsNormal;
    	if TOBL.getValue('GL_TYPELIGNE')<> 'SD' then
      begin
      	ImTypeArticle.Draw(CanVas, ARect.left, ARect.top, NumGraph);
      end else
      begin
      	ImTypeArticle.Draw(CanVas, ARect.left+10, ARect.top, NumGraph);
      end;
    end;
  end;
  if (Acol = CODEARTICLE) and (Arow >= GS.fixedRows) then
  begin
    if TOBL = nil then Exit;
    if TOBL.getValue('GL_TYPELIGNE')<> 'SD' then exit;
    TheText := TOBL.getvalue('GL_REFARTSAISIE');
    Canvas.FillRect(ARect);
    GS.Canvas.Brush.Style := bsSolid;
    GS.Canvas.TextOut (Arect.left,Arect.Top +2,'   '+Thetext);
  end;
  //
  if (Acol = LIBELLE) and (Arow >= GS.fixedRows) then
  begin
    if TOBL = nil then Exit;
    if TOBL.getValue('GL_TYPELIGNE')<> 'SD' then exit;
    TheText := TOBL.getvalue('GL_LIBELLE');
    Canvas.FillRect(ARect);
    GS.Canvas.Brush.Style := bsSolid;
    GS.Canvas.TextOut (Arect.left,Arect.Top +2,'   '+Thetext);
  end;
end;

procedure TOF_BTSAISIEAVANC.GSRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
var TOBL : TOB;
    TypeLigne : string;
    NiveauImbric : integer;
begin
  if Action = taConsult then Exit;
 	ShowButton(fCurrentOnglet, Ou,false);
  if Ou <= GS.fixedCols then exit;
  TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,Ou);
  if TOBL = nil then exit;
  TypeLigne := TOBL.GetValue('GL_TYPELIGNE');
  NiveauImbric := TOBL.GetValue('GL_NIVEAUIMBRIC');
  if ((copy(TypeLigne,1,2)='DP') or (copy(TypeLigne,1,2)='DV')) or (NiveauImbric > 0) then
  begin
    MnDefaireParag.Enabled := true;
  end else
  begin
    MnDefaireParag.Enabled := false;
  end;
 	ShowButton(fCurrentOnglet,Ou,true);
 end;

procedure TOF_BTSAISIEAVANC.GSRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  if Action = taConsult then Exit;
end;

function TOF_BTSAISIEAVANC.GetTOBLigneAV(TOBlignes: TOB; Arow: integer): TOB;
begin
  result := nil;
  if ARow > TOBLIgnes.detail.count then exit;
  result := TOBLIgnes.detail[Arow-1];
end;


function TOF_BTSAISIEAVANC.RecupTypeGraph(TOBL: TOB): integer;
var TypeArticle,TypeLigne : string;
    NiveauImbric,IndiceNOmen : integer;
begin
  TypeArticle := TOBL.GetValue('GL_TYPEARTICLE');
  NiveauImbric := TOBL.GetValue('GL_NIVEAUIMBRIC');
  TypeLigne := TOBL.GetValue('GL_TYPELIGNE');
  IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');

  if TypeArticle = 'MAR' Then Result := 7
  else if (TypeArticle = 'OUV') then Result := 1
  else if (TypeArticle = 'POU') then Result := 6
  else if (TypeArticle = 'COM') and (IndiceNomen = 0) then Result := 2
  // VARIANTE
  else if ((copy(TypeLigne,1,2)='DP') or (copy(TypeLigne,1,2)='DV')) and (NiveauImbric > 1) then Result := 4
  else if ((copy(TypeLigne,1,2)='DP') or (copy(TypeLigne,1,2)='DV')) and (NiveauImbric = 1 ) then Result := 3
  // --
  else if TypeArticle ='PRE' then Result := 5
  else if TypeArticle ='FRA' then Result := 8
  else if TypeArticle = 'ARP' then Result := 9
  else Result := -1;
end;


procedure TOF_BTSAISIEAVANC.ChargeListImage;
var UneImage : Timage;
    NomImg : string;
    indice : integer;
begin
  Indice := 1;
  NomImg := 'IMG'+IntToStr(Indice-1);
  UneImage := TImage(GetCOntrol(NomImg));
  repeat
    if UneImage <> nil then
    begin
      if ImTypeArticle.AddMasked  (UneImage.Picture.Bitmap,TColor($FF00FF)) < 0 then exit;
      //
      inc(Indice);
      //
      NomImg := 'IMG'+IntToStr(Indice-1);
      UneImage := TImage(GetCOntrol(NomImg));
    end;
  until UneImage = nil;
  Indice := 1;
  NomImg := 'IMGP'+IntToStr(Indice-1);
  UneImage := TImage(GetCOntrol(NomImg));
  repeat
    if UneImage <> nil then
    begin
      if ImgListPlusMoins.AddMasked  (UneImage.Picture.Bitmap,TColor($FF00FF)) < 0 then exit;
      //
      inc(Indice);
      //
      NomImg := 'IMGP'+IntToStr(Indice-1);
      UneImage := TImage(GetCOntrol(NomImg));
    end;
  until UneImage = nil;
  // ---
  Indice := 1;
  NomImg := 'TYPETRAIT'+IntToStr(Indice);
  UneImage := TImage(GetCOntrol(NomImg));
  repeat
    if UneImage <> nil then
    begin
      if ImgTypeTrait.AddMasked (UneImage.Picture.Bitmap,TColor($FF00FF)) < 0 then exit;
      //
      inc(Indice);
      //
      NomImg := 'TYPETRAIT'+IntToStr(Indice);
      UneImage := TImage(GetCOntrol(NomImg));
    end;
  until UneImage = nil;
end;

procedure TOF_BTSAISIEAVANC.FreeComposants;
begin
  ImTypeArticle.free;
  ImgTypeTrait.free;
  ImgListPlusMoins.free;
  ToolWin.free;
end;

procedure TOF_BTSAISIEAVANC.ChangeQteSituation(TOBL: TOB; CurrOnglet : Tonglet;  var Cancel: boolean);
begin
  CalculeLaLigneFromQteSit (TOBL,CurrOnglet);
end;

procedure TOF_BTSAISIEAVANC.ChangeQteFact (TOBL: TOB; CurrOnglet : Tonglet;  var Cancel: boolean);
begin
  CalculeLaLigneFromQteFact (TOBL,CurrOnglet);
end;


procedure TOF_BTSAISIEAVANC.ChangePourcentAvanc(TOBL: TOB;  CurrOnglet : Tonglet; var Cancel: boolean);
begin
  CalculeLaLigneFromPourcent (TOBL,CurrOnglet);
end;

procedure TOF_BTSAISIEAVANC.SetAvancSousDetail (TOBL : TOB;Onglet : Tonglet);
var IndiceNomen : Integer;
		II : integer;
    TOBOO,TOBOUV : TOB;
begin
  IndiceNomen := TOBL.getValue('GL_INDICENOMEN');
  if IndiceNomen = 0 then exit;

  //FV1 : 28/07/2016
  if TobOuvrages.detail.count = 0 then exit;
  if (Indicenomen-1) >= (TobOuvrages.detail.count-1) then exit;

  TOBOUV := TOBOuvrages.detail[IndiceNomen-1];
  for II := 0 to TOBOUV.detail.count -1 do
  begin
  	TOBOO := TOBOUV.detail[II];
    if TOBOO.getValue('BLO_UNIQUEBLO')=TOBL.GetValue('UNIQUEBLO') then
    begin
    	TOBOO.PutValue('BLF_MTCUMULEFACT',TOBL.getVAlue('BLF_MTCUMULEFACT'));
    	TOBOO.PutValue('BLF_MTSITUATION',TOBL.getVAlue('BLF_MTSITUATION'));
    	TOBOO.PutValue('BLF_QTECUMULEFACT',TOBL.getVAlue('BLF_QTECUMULEFACT'));
    	TOBOO.PutValue('BLF_QTESITUATION',TOBL.getVAlue('BLF_QTESITUATION'));
    	TOBOO.PutValue('BLF_POURCENTAVANC',TOBL.getVAlue('BLF_POURCENTAVANC'));
    	break;
    end;
  end;
end;
(*
procedure TOF_BTSAISIEAVANC.PositionneAvancOuv (TOBL : TOB);
var IndiceNomen : Integer;
		II : integer;
    TOBOO,TOBOUV : TOB;
begin
  IndiceNomen := TOBL.getValue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
  TOBOUV := TOBOuvrages.detail[IndiceNomen-1];
  for II := 0 to TOBOUV.detail.count -1 do
  begin
  	TOBOO := TOBOUV.detail[II];
    if Avancement then
    begin
      // Pour les montants
      TOBOO.PutValue('BLF_MTCUMULEFACT',Arrondi(TOBOO.GetValue('BLF_MTMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
      TOBOO.PutValue('BLF_MTSITUATION',TOBOO.GetValue('BLF_MTCUMULEFACT')-TOBOO.GetValue('BLF_MTDEJAFACT'));
      // Pour les Quantités
      TOBOO.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBOO.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
      TOBOO.PutValue('BLF_QTESITUATION',TOBOO.GetValue('BLF_QTECUMULEFACT')-TOBOO.GetValue('BLF_QTEDEJAFACT'));
    end else
    begin
      // Pour les montants
      TOBOO.PutValue('BLF_MTCUMULEFACT',0);
      TOBOO.PutValue('BLF_MTSITUATION',Arrondi(TOBOO.GetValue('BLF_MTMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
      // Pour les Quantités
      TOBOO.PutValue('BLF_QTECUMULEFACT',0);
      TOBOO.PutValue('BLF_QTESITUATION',Arrondi(TOBOO.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
    end;
  end;
end;
*)
procedure TOF_BTSAISIEAVANC.CalculeLaLigneFromPourcent(TOBL: TOB; CurrOnglet : Tonglet);
begin
  if Avancement then
  begin
    if ReajusteDossier then
    begin
      TOBL.PutValue('BLF_MTCUMULEFACT',Arrondi(TOBL.GetValue('BLF_MTMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
      TOBL.PutValue('BLF_MTDEJAFACT',TOBL.GetValue('BLF_MTCUMULEFACT'));
      TOBL.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
      TOBL.PutValue('BLF_QTEDEJAFACT',TOBL.GetValue('BLF_QTECUMULEFACT'));
    end else
    begin
    // Pour les montants
    TOBL.PutValue('BLF_MTCUMULEFACT',Arrondi(TOBL.GetValue('BLF_MTMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
    TOBL.PutValue('BLF_MTSITUATION',TOBL.GetValue('BLF_MTCUMULEFACT')-TOBL.GetValue('BLF_MTDEJAFACT'));
    // Pour les Quantités
    TOBL.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
    TOBL.PutValue('BLF_QTESITUATION',TOBL.GetValue('BLF_QTECUMULEFACT')-TOBL.GetValue('BLF_QTEDEJAFACT'));
    end;
  end else
  begin
    if ReajusteDossier then
    begin
      TOBL.PutValue('BLF_MTDEJAFACT',Arrondi(TOBL.GetValue('BLF_MTMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
      TOBL.PutValue('BLF_QTEDEJAFACT',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
      TOBL.PutValue('BLF_MTCUMULEFACT',0);
      TOBL.PutValue('BLF_QTECUMULEFACT',0);
    end else
    begin
    // Pour les montants
    TOBL.PutValue('BLF_MTCUMULEFACT',0);
    TOBL.PutValue('BLF_MTSITUATION',Arrondi(TOBL.GetValue('BLF_MTMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
    // Pour les Quantités
    TOBL.PutValue('BLF_QTECUMULEFACT',0);
    TOBL.PutValue('BLF_QTESITUATION',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
  end;
  end;
  TOBL.PutValue('BLF_MTPRODUCTION',TOBL.GetValue('BLF_MTSITUATION'));
  if isSousDetail (TOBL) {and (CurrOnglet = FOngletGlobal) } and (TOBL.GetValue('UNIQUEBLO')<>0) then
  begin
  	SetAvancSousDetail (TOBL,CurrOnglet);
  end;
  if (IsOuvrage(TOBL))  and (not IsSousDetail(TOBL)) and (TOBL.getValue('UNIQUEBLO')=0)  then
  begin
  	AppliqueAvancSousDetail (TOBL,CurrOnglet);
  end;
end;

procedure TOF_BTSAISIEAVANC.CumuleAvancDetailGeneral(TOBD : TOB;Sens : string);
var ListeDet : TlisteTOB;
    TOBL : TOB;
    Indice : integer;
begin
	ListeDet := FindlignesOuvSaisie (fCurrentOnglet,TOBD);
  if ListeDet.count > 0 then   // lignes de détail trouvé
  begin
  	for Indice := 0 to ListeDet.count -1 do
    begin
    	TOBL := ListeDet.Items [Indice];
      CumuleLigneGlobal (TOBL,Sens);
      //
    	AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,TOBL.getvalue('INDICE'));
    end;
  end;
  ListeDet.free;
end;

procedure TOF_BTSAISIEAVANC.CumuleLigneGlobal (TOBD : TOB; Sens : string);
var indice : integer;
		LastOnglet : Tonglet;
    numOrdre,uniqueBlo,NumPiece : integer;
    TOBL : TOB;
begin
  if IsOuvragePartiel(TOBD) then
  begin
    CumuleAvancDetailGeneral (TOBD,Sens);
  end else
  begin
    lastOnglet := fCurrentOnglet;
    NumOrdre := TOBD.getvalue('GL_NUMORDRE');
    UniqueBlo := TOBD.getvalue('UNIQUEBLO');
    NumPiece := TOBD.getvalue('GL_NUMERO');
    indice := FindTOBligne (FongletGlobal,NumPiece,Numordre,uniqueBlo);
    if indice < 0 then exit;
    SetOnglet(FOngletGlobal);
    TOBL := GetTOBLigneAV(fCurrentOnglet.TOBlignes ,indice+1);
    if TOBL <> nil then CumuleLigne(TOBL,Sens);
    SetOnglet(LastOnglet);
  end;
end;
{
procedure TOF_BTSAISIEAVANC.CumuleLigneFournisseurSousDetail (TOBO : TOB; CurrOnglet : Tonglet: Sens : string);
var ListeDet : TlisteTOB;
		avanc : double;
    TOBL : TOB;
    Indice : integer;
begin
	ListeDet := FindlignesOuvSaisie (CurrOnglet,TOBO);
  avanc := TOBO.getValue('BLF_POURCENTAVANC');
  if ListeDet.count > 0 then   // lignes de détail trouvé
  begin
  	for Indice := 0 to ListeDet.count -1 do
    begin
    	TOBL := ListeDet.Items [Indice];
      if Avancement then
      begin
      	TOBL.putValue('BLF_POURCENTAVANC',Avanc);
        // Pour les montants
        TOBL.PutValue('BLF_MTCUMULEFACT',Arrondi(TOBL.GetValue('BLF_MTMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
        TOBL.PutValue('BLF_MTSITUATION',TOBL.GetValue('BLF_MTCUMULEFACT')-TOBL.GetValue('BLF_MTDEJAFACT'));
        // Pour les Quantités
        TOBL.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
        TOBL.PutValue('BLF_QTESITUATION',TOBL.GetValue('BLF_QTECUMULEFACT')-TOBL.GetValue('BLF_QTEDEJAFACT'));
      end else
      begin
      	TOBL.putValue('BLF_POURCENTAVANC',Avanc);
        // Pour les montants
        TOBL.PutValue('BLF_MTCUMULEFACT',0);
        TOBL.PutValue('BLF_MTSITUATION',Arrondi(TOBL.GetValue('BLF_MTMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
        // Pour les Quantités
        TOBL.PutValue('BLF_QTECUMULEFACT',0);
        TOBL.PutValue('BLF_QTESITUATION',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
      end;
      //
      if (isSousDetail (TOBL)) and (CurrOnglet = FOngletGlobal) and (TOBL.GetValue('UNIQUEBLO')<>0) then
      begin
      	SetAvancSousDetail (TOBL,CurrOnglet);
        AppliqueAvancFournisseur (TOBL,avanc);
      end;
      //
    	AfficheLaLigne (CurrOnglet.GS,CurrOnglet.TOBlignes,TOBL.getvalue('INDICE'));
    end;
  end else
  begin
  	SetAvancOuvrage (TOBO);
  end;
  ListeDet.free;

end;
}
function TOF_BTSAISIEAVANC.CumuleLigneFournisseur (TOBD: TOB;Sens : string) : Tonglet;
var NumOrdre,UniqueBlo,NumPiece : integer;
		LastOnglet,TheOnglet : TOnglet;
    Indice : integer;
    TOBL : TOB;
    NatureTravail : double;
begin
  (*
  NatureTravail := VALEUR(TOBD.GetValue('GLC_NATURETRAVAIL'));
  if IsOuvrage(TOBD) and (fCurrentOnglet=FOngletGlobal) and (NatureTravail > 10) then
  begin
  	CumuleLigneFournisseurSousDetail (TOBD,fCurrentOnglet,Sens );
    Exit;
  end;
  *)
	result := nil;
  NumOrdre := TOBD.getvalue('GL_NUMORDRE');
  UniqueBlo := TOBD.getvalue('UNIQUEBLO');
  NumPiece := TOBD.getvalue('GL_NUMERO');
  lastOnglet := fCurrentOnglet;
  indice := -1;
  TheOnglet := FindLigneInOnglet(NumPiece,NumOrdre,UniqueBlo,indice);
  if TheOnglet = nil then exit;
  result := TheOnglet;
  SetOnglet(TheOnglet);
  TOBL := GetTOBLigneAV(fCurrentOnglet.TOBlignes ,indice+1);
  if TOBL <> nil then
  begin
    CumuleLigne(TOBL,sens);
  end;
  SetOnglet(lastOnglet);
end;

procedure TOF_BTSAISIEAVANC.CumuleLigneOuv(TOBD: TOB; Sens : string);
var TOBL : TOB;
		pourcent : double;
    indice : integer;
begin
	indice := FindTOBligne (fCurrentOnglet,TOBD.getvalue('GL_NUMERO'),TOBD.getValue('GL_NUMORDRE'),0);
  if indice < 0 then exit;
  TOBL := GetTOBLigneAV(fCurrentOnglet.TOBlignes ,indice+1);
  if TOBL <> nil then
  begin
  	if Sens = '-' then
    begin
      // Pour les montants
      TOBL.PutValue('BLF_MTSITUATION',Arrondi(TOBL.GetValue('BLF_MTSITUATION')-TOBD.GetValue('BLF_MTSITUATION'),DEV.decimale));
      TOBL.PutValue('BLF_MTCUMULEFACT',TOBL.GetValue('BLF_MTDEJAFACT')-TOBL.GetValue('BLF_MTSITUATION'));
      // Pour les Quantités
//      TOBL.PutValue('BLF_QTECUMULEFACT',TOBL.GetValue('BLF_QTEDEJAFACT')-TOBL.GetValue('BLF_QTESITUATION'));
      if Avancement then
      begin
        Pourcent := Arrondi(TOBL.GetValue('BLF_MTCUMULEFACT')/TOBL.GetValue('BLF_MTMARCHE')*100,2);
        TOBL.SetDouble('BLF_POURCENTAVANC',Pourcent);
      end else
      begin
        Pourcent := Arrondi(TOBL.GetValue('BLF_MTSITUATION')/TOBL.GetValue('BLF_MTMARCHE')*100,2);
        TOBL.SetDouble('BLF_POURCENTAVANC',Pourcent);
      end;
    end else
    begin
// Pour les montants
      TOBL.PutValue('BLF_MTSITUATION',Arrondi(TOBL.GetValue('BLF_MTSITUATION')+TOBD.GetValue('BLF_MTSITUATION'),DEV.decimale));
      TOBL.PutValue('BLF_MTCUMULEFACT',TOBL.GetValue('BLF_MTDEJAFACT')+TOBL.GetValue('BLF_MTSITUATION'));
      // Pour les Quantités
//      TOBL.PutValue('BLF_QTECUMULEFACT',TOBL.GetValue('BLF_QTEDEJAFACT')+TOBL.GetValue('BLF_QTESITUATION'));
      if Avancement then
      begin
        Pourcent := Arrondi(TOBL.GetValue('BLF_MTCUMULEFACT')/TOBL.GetValue('BLF_MTMARCHE')*100,2);
        TOBL.SetDouble('BLF_POURCENTAVANC',Pourcent);
      end else
      begin
        Pourcent := Arrondi(TOBL.GetValue('BLF_MTSITUATION')/TOBL.GetValue('BLF_MTMARCHE')*100,2);
        TOBL.SetDouble('BLF_POURCENTAVANC',Pourcent);
      end;
    end;
    TOBL.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,V_PGI.OkdecQ));
    TOBL.PutValue('BLF_QTESITUATION',Arrondi(TOBL.GetValue('BLF_QTECUMULEFACT')-TOBL.GetValue('BLF_QTEDEJAFACT')/100,V_PGI.OkdecQ));

    AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,Indice+1);
  end;
end;

procedure TOF_BTSAISIEAVANC.CumuleLigne(TOBL: TOB; Sens : string= '+');
begin
	if TOBL.getvalue('GL_TYPELIGNE') = 'SD' then
  begin
  	CumuleLigneOuv (TOBL,Sens);
  end;
  SetAvancementPiece (TOBL,Sens);
  if (Pos(TOBL.GetString('GL_NATUREPIECEG'),'DBT;FBP;FBT;ABP;ABT')=0) or (VH_GC.BTCODESPECIF <> '001') or (not IsLigneSoustraite (TOBL)) then
  begin
    if Sens = '-' then
    begin
      dMtSituation := dMtSItuation - TOBL.GetValue('BLF_MTSITUATION');
    end else
    begin
      dMtSituation := dMtSItuation + TOBL.GetValue('BLF_MTSITUATION');
    end;
  end;
  FMTSITUATION.Value := dMtSItuation;
  BPF_MTSITUATION.caption := FMTSITUATION.Text;
  fCurrentOnglet.MtSituation := dMtSituation;
end;

procedure TOF_BTSAISIEAVANC.GetInfoDev(CodeDevise: string);
begin
  DEV.Code := CodeDevise;
  GetInfosDevise(DEV);
  TDEVISE.Caption := DEV.Symbole;
end;

procedure TOF_BTSAISIEAVANC.GSGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var TOBL : TOB;
begin
  if ACol < GS.FixedCols then Exit;
  IF Arow < GS.fixedRows then exit;
  TOBL := GetTOBLigneAV( fCurrentOnglet.TOBlignes,ARow);
  if TOBL = nil then Exit;
  {Ligne de sous sétail}
  if (pos(TOBL.GetValue('GL_TYPELIGNE'),'SD;OUP')>0) and ((Action <> taConsult) or (ARow <> GS.Row)) then
  begin
    if (TOBL.GetValue('GL_TYPELIGNE')='SD') then
    begin
      Canvas.Font.Style := Canvas.Font.Style + [fsbold];
      Canvas.Font.Color := clActiveCaption;
    end;
    if (TOBL.GetValue('BLF_MTSITUATION') <> 0) then
    begin
      if Acol >= GS.FixedCols then
      begin
        Canvas.Font.Style := Canvas.Font.Style + [fsbold];
        Canvas.Font.Color := clBlue;
      end;
    end;
  end;
  {Ligne non imprimable}
  if TOBL.GetValue('GL_NONIMPRIMABLE') = 'X' then if ((Action <> taConsult) or (ARow <> GS.Row)) then Canvas.Font.Color := clBlue;
  {Lignes de sous-total}
  if TOBL.GetValue('GL_TYPELIGNE') = 'TOT' then Canvas.Font.Style := Canvas.Font.Style + [fsBold, fsItalic];
  {Lignes début de paragraphe}// Modif BTP
  if IsDebutParagraphe (TOBL) then
  begin
    if TOBL.GetValue('GL_NIVEAUIMBRIC') > 1 then Canvas.Font.Style := Canvas.Font.Style + [fsItalic]
    else Canvas.Font.Style := Canvas.Font.Style + [fsBold, fsItalic];
  end;
  {Lignes fin de paragraphe}// Modif BTP
  (* VARIANTE  -- if Copy(TOBL.GetValue('GL_TYPELIGNE'), 1, 2) = 'TP' then *)
  if IsFinParagraphe (TOBL) then
  begin
    if TOBL.GetValue('GL_NIVEAUIMBRIC') > 1 then Canvas.Font.Style := Canvas.Font.Style + [fsItalic]
    else Canvas.Font.Style := Canvas.Font.Style + [fsBold, fsItalic];
  end;
  {Ligne de commentaire rattachée}// Modif BTP
  if IsLigneDetail(nil, TOBL) then
  begin
    Canvas.Font.Style := Canvas.Font.Style + [fsbold, fsItalic];
    Canvas.Font.Color := clActiveCaption;
  end;
  //
  {Ligne de retenue de garantie}// Modif BTP
  if (TOBL.GetValue('GL_TYPELIGNE') = 'RG') then
  begin
    Canvas.Font.Style := Canvas.Font.Style + [fsbold, fsItalic];
    Canvas.Font.Color := clRed;
  end;
  // VARIANTE
  if (isVariante(TOBL)) then
  begin
    Canvas.Font.Style := Canvas.Font.Style + [fsbold,fsItalic];
    Canvas.Font.Color := clMaroon;
  end;
  // Ligne de référence de livraison ... Commande fournisseur ou réception
  if TOBL.GetValue('GL_TYPELIGNE') = 'REF' then
  begin
    if Acol = LIBELLE then
    begin
      Canvas.Font.Style := Canvas.Font.Style + [fsBold,fsUnderline];
      Canvas.Font.Color := clRed;
    end;
  end;
  if (TOBL.GetValue('BLF_MTSITUATION') <> 0) and (IsArticle(TOBL)) then
  begin
    if Acol >= GS.FixedCols then
    begin
      Canvas.Font.Style := Canvas.Font.Style + [fsbold];
      Canvas.Font.Color := clBlue;
    end;
  end;
  //
  if (Acol = MTMARCHE) or (Acol = MTDEJAFACT) then BEGIN canvas.Brush.Color := clInfoBk; END;
  IF (Acol= POURCENTAVANC) then BEGIN Canvas.font.color := $E6A008; Canvas.Font.Style := Canvas.Font.Style + [fsBold]; END;
  IF (Acol= MTSITUATION) then BEGIN canvas.Brush.Color := clInfoBk; Canvas.Font.Style := Canvas.Font.Style + [FsBold]; END;
  //
end;

procedure TOF_BTSAISIEAVANC.AddChampsSupMtINITAvanc(TOBL: TOB);
begin
  TOBL.AddChampSupValeur ('POURCENTAVANCINIT',0);
  TOBL.AddChampSupValeur ('MTCUMULEFACTINIT',0);
  TOBL.AddChampSupValeur ('QTECUMULEFACTINIT',0);
end;

procedure TOF_BTSAISIEAVANC.InitValeurZero(TOBL: TOB);
begin
  TOBL.PutValue ('BLF_NATUREPIECEG','');
  TOBL.PutValue ('BLF_MTMARCHE',0);
  TOBL.PutValue ('BLF_MTDEJAFACT',0);
  TOBL.PutValue ('BLF_MTCUMULEFACT',0);
  TOBL.PutValue ('BLF_MTSITUATION',0);
  TOBL.PutValue ('BLF_QTEMARCHE',0);
  TOBL.PutValue ('BLF_QTEDEJAFACT',0);
  TOBL.PutValue ('BLF_QTECUMULEFACT',0);
  TOBL.PutValue ('BLF_QTESITUATION',0);
  TOBL.PutValue ('BLF_POURCENTAVANC',0);
  TOBL.PutValue ('BLF_MTPRODUCTION',0);
end;

procedure TOF_BTSAISIEAVANC.SetLigne (GS : THgrid; Arow:integer;Acol : integer=-1);
var
    Cancel : boolean;
begin
  GS.synEnabled := false;
  GS.CacheEdit;
  SetEventsGrid (GS,false);
  Arow := Arow;
  Acol := 1;
  GS.Row := Arow;
  GS.Col := Acol;
  GSCellEnter (self,Acol,Arow,cancel);
  stPrev  :=GS.Cells[ACol,ARow] ;
  fCurrentOnglet.stPrev := stprev;
  GS.row := ARow; GS.Col := Acol;
  GS.SynEnabled := true;
  GS.MontreEdit;
  SetEventsGrid (GS,true);
end;

procedure TOF_BTSAISIEAVANC.ChangeMontantSituation(Sender: TOBject);
begin
  BPF_MTSITUATION.caption := FMTSITUATION.Text;
end;

procedure TOF_BTSAISIEAVANC.BAVANCClick(Sender: Tobject);
var
  TOBL: TOB;
  EtenduApplication: integer; // tout le document
  IndDepart, Indfin: integer;
  Pourcent: double;
  Parag: boolean;
  lastOnglet : Tonglet;
begin
	lastOnglet := fCurrentOnglet;
  EtenduApplication := 0;
  Parag := false;
  if (GS.row > 0) then
  begin
    TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,GS.row);
    if (TOBL.GetValue('GL_NIVEAUIMBRIC')<>0) then
    begin
      Parag := true;
    end;
  end;
  if AssistSaisieAvanc(EtenduApplication, Pourcent, parag,true) then
  begin
    if EtenduApplication = 1 then
    begin
      IndDepart := RecupDebutParagraph(GS.row );
      Indfin := RecupFinParagraph(GS.row);
    end else if EtenduApplication = 0 then
    begin
      indDepart := RecupDebutPieceCur(GS.row);
      IndFin := RecupFinPieceCur(GS.row);
    end else
    begin
      indDepart := 1;
      IndFin := GS.rowCount-1;
    end;
    AppliquePourcentAvanc( Pourcent, IndDepart, IndFin);
  end;
  SetOnglet(LastOnglet);
  GS.Invalidate;
end;

function TOF_BTSAISIEAVANC.RecupDebutParagraph(Ligne: integer) : integer;
var TOBL : TOB;
    StopIt : integer;
    NiveauImbric : integer;
begin
  result := 1;
  StopIt := Ligne;
  TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,StopIt);
  NiveauImbric := TOBL.GetValue('GL_NIVEAUIMBRIC');
  repeat
    if IsDebutParagraphe (TOBL,NiveauImbric) then
    begin
      result := StopIt+1;
      break;
    end else Dec(StopIt);
    if Stopit > 0 then TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,StopIt);
  until StopIt = 0;
end;

function TOF_BTSAISIEAVANC.RecupFinParagraph(Ligne: integer) : integer;
var TOBL : TOB;
    StopIt : integer;
    NiveauImbric : integer;
begin
  result := GS.rowCount-1;
  StopIt := Ligne;
  TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,StopIt);
  NiveauImbric := TOBL.GetValue('GL_NIVEAUIMBRIC');
  repeat
    if IsFinParagraphe (TOBL,NiveauImbric) then
    begin
      result := StopIt-1;
      break;
    end else inc(StopIt);
    if stopIt <= GS.RowCount then TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,StopIt);
  until StopIt = GS.rowCount;
end;

function TOF_BTSAISIEAVANC.RecupDebutPieceCur(Ligne: integer): integer;
var TOBL : TOB;
    StopIt : integer;
    lacle : string;
    cledoc : r_cledoc;
begin
  result := 1;
  StopIt := Ligne;
  TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,StopIt);
  laCle := EncodeRefPiece (TOBL);
  DecodeRefPiece (LaCle,Cledoc);
  repeat
    if IsDebutPiece (TOBL,cledoc) then
    begin
      result := StopIt+1;
      break;
    end else Dec(StopIt);
    if StopIt > 0 then TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,StopIt);
  until StopIt = 0;
end;

function TOF_BTSAISIEAVANC.RecupFinPieceCur(Ligne: integer): integer;
var TOBL : TOB;
    StopIt : integer;
    lacle : string;
    cledoc : r_cledoc;
begin
  result := GS.rowCount-1;
  StopIt := Ligne;
  TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,StopIt);
  laCle := EncodeRefPiece (TOBL);
  DecodeRefPiece (LaCle,Cledoc);
  repeat
    if not IsSamePiece (TOBL,cledoc) then
    begin
      result := StopIt -1; // on revient a la ligne précédente
      break;
    end else inc(StopIt);
    if Stopit < GS.rowCount  then TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,StopIt);
  until StopIt = GS.RowCount;
end;

procedure TOF_BTSAISIEAVANC.AppliquePourcentAvanc(Pourcent: double;IndDepart, IndFin: integer);
var TOBL,TOBDL : TOB;
    Indice,Index : integer;
    cancel : boolean;
    TOBDecisionL : TOB;
    Resultat : integer;
    LibelleErreur : string;
    ListeOnglet : TStringList;
    unOnglet : Tonglet;
begin
	listeOnglet := TstringList.Create;
  //
  TOBDecisionL := TOB.Create ('LES LIGNES A CONTROLER',nil,-1);
  TOBDECISIONL.AddChampSupValeur ('OKSAISIE','-');
  TOBDECISIONL.AddChampSupvaleur ('AVANCEMENT','-');

  if Avancement then TOBDECISIONL.AddChampSupvaleur ('AVANCEMENT','X');
  TRY
    for Indice := indDepart to IndFin do
    begin
      //
      TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,Indice);
      if TOBL.GetValue('BLF_MTMARCHE')= 0 then Continue;
      if Pos(TOBL.GetValue('GL_TYPELIGNE'),'ART;OUP') > 0 then
      begin
        Resultat := ChangePourcentOkMulti (Indice,POurcent,LibelleErreur);
        if  Resultat = 0 then
        begin
          if TOBL.GetValue('BLF_MTSITUATION')<>0 then
          begin
            ReInitLigne (TOBL);
            if fCurrentOnglet <> FOngletGlobal then
            begin
              ReInitLigneGlobal (TOBL);
            end else
            begin
              unOnglet := ReInitLigneFournisseur (TOBL);
              if unOnglet <> nil then
              begin
              	if not ListeOnglet.Find(UnOnglet.Name,Index) then
                	ListeOnglet.add(UnOnglet.name);
              end;
            end;
          end else
          begin
          	CumuleLigne(TOBL,'-');
            if fCurrentOnglet <> FOngletGlobal then
            begin
              CumuleLigneGlobal (TOBL,'-');
            end else
            begin
              unOnglet := CumuleLigneFournisseur (TOBL,'-');
              if unOnglet <> nil then
              begin
              	if not ListeOnglet.Find(UnOnglet.Name,Index) then
                	ListeOnglet.add(UnOnglet.name);
              end;
            end;
          end;
          TOBL.putValue('BLF_POURCENTAVANC',Pourcent);
          ChangePourcentAvanc (TOBL,fCurrentOnglet,Cancel);
          CumuleLigne(TOBL);
          if fCurrentOnglet <> FOngletGlobal then
          begin
          	AppliqueAvancGeneral (TOBL,Pourcent,false,true,false);
          end else
          begin
          	AppliqueAvancFournisseur (TOBL,Pourcent,false,true,false);
          end;
        end
        else
        begin
          // Modified by f.vautrain 07/11/2017 15:53:27 - FS#2774 - GUINIER - Interdire la possibilité de facturer supérieur à 100 %
          {$IFDEF V10}
          if (CtrlSupCent) And (resultat = 1) then
          begin
            PGIError(LibelleErreur, 'erreur avancement');
            Break;
          end
          else
          begin
          {$ENDIF}
            TOBDL := TOB.create ('UNE LIGNE',   TOBdecisionL,-1);
            TOBDL.AddChampSupValeur ('ANOMALIE',resultat);
            TOBDL.AddChampSupValeur ('LIGNE',   Indice);
            TOBDL.AddChampSupValeur ('LIBELLE', TOBL.GetValue('GL_LIBELLE'));
            TOBDL.AddChampSupValeur ('ANOMALIE',LibelleErreur);
            TOBDL.AddChampSupValeur ('OK',      '-');
          {$IFDEF V10}
          end;
          {$ENDIF}
        end;
      end;
    end;
    //
    if TOBDecisionL.detail.count > 0 then
    begin
      DefiniChoixApplication (TOBDecisionL);
      if TOBDecisionL.GetValue('OKSAISIE')='X' then
      begin
        for Indice := 0 to TOBDecisionL.detail.count - 1 do
        begin
          TOBDL := TOBDecisionL.detail[Indice];
          if TOBDL.GetValue('OK')='X' then
          begin
            TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,TOBDL.getValue('LIGNE'));
            CumuleLigne(TOBL,'-');
            if TOBL.GetValue('BLF_MTSITUATION')<>0 then  ReInitLigne (TOBL);
            TOBL.putValue('BLF_POURCENTAVANC',Pourcent);
            ChangePourcentAvanc (TOBL,fCurrentOnglet,Cancel);
            CumuleLigne(TOBL);
            if fCurrentOnglet <> FOngletGlobal then
            begin
              AppliqueAvancGeneral (TOBL,Pourcent,false,true,false);
            end else
            begin
              AppliqueAvancFournisseur (TOBL,Pourcent,false,true,false);
            end;
          end;
        end;
      end;
    end;
  FINALLY
    TOBDecisionL.free;
    CalculeLesSousTotaux;
    AfficheLaGrille (fCurrentOnglet.GS, fCurrentOnglet.TOBlignes);
    if fCurrentOnglet <> FOngletGlobal then
    begin
    	RecalculeOnglet (FOngletGlobal);
    end else
    begin
      ReCalculelesOnglets (ListeOnglet);
    end;
    listeonglet.Free;
  END;
end;

function TOF_BTSAISIEAVANC.IsDebutPiece(TOBL: TOB; cledoc: r_cledoc): boolean;
begin
  result := false;
  if (TOBL.GetValue('GL_NATUREPIECEG')=cledoc.NaturePiece) and
     (TOBL.GetValue('GL_SOUCHE')=cledoc.Souche) and
     (TOBL.GetValue('GL_NUMERO')=cledoc.NumeroPiece) and
     (TOBL.GetValue('GL_INDICEG')=cledoc.Indice) and
     (TOBL.GetValue('GL_TYPELIGNE')='REF') then result := true;
end;

function TOF_BTSAISIEAVANC.IsSamePiece(TOBL: TOB; cledoc: r_cledoc): boolean;
begin
  result := false;
  if (TOBL.GetValue('GL_NATUREPIECEG')=cledoc.NaturePiece) and
     (TOBL.GetValue('GL_SOUCHE')=cledoc.Souche) and
     (TOBL.GetValue('GL_INDICEG')=cledoc.Indice) and
     (TOBL.GetValue('GL_NUMERO')=cledoc.NumeroPiece) then result := true;
end;

function TOF_BTSAISIEAVANC.IsGerableAvanc(TOBL: TOB): boolean;
begin
//  result := false;
//  if (IsArticle (TOBL)) then result := true;
	result := true;
end;


procedure TOF_BTSAISIEAVANC.CalculeSousTotal (Ligne,IChampST : integer);
var Niv: integer;
  TOBL: TOB;
  TypeL: string;
begin
  TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,Ligne); if TOBL = nil then exit;
  TypeL := TOBL.GetValeur(IChampST);
  if (TypeL <> 'TOT') and (not IsFinParagraphe (TOBL)) then exit;

  if TypeL = 'TOT' then BEGIN SommerLignesAV( Ligne, 0); AfficheLaLigne(GS,fCurrentOnglet.TOBlignes,Ligne);END;
  // VARIANTE
  if ISFinParagraphe (TOBL) then
  begin
    Niv := StrToInt(Copy(TypeL, 3, 1));
    SommerLignesAV( Ligne, Niv);
    AfficheLaLigne(GS,fCurrentOnglet.TOBlignes,Ligne);
  end;
end;

procedure TOF_BTSAISIEAVANC.CalculelesSousTotaux;
var i, iTot: integer;
  TOBL: TOB;
begin
  iTot := 0;
  for i := 0 to fCurrentOnglet.TOBlignes.Detail.Count - 1 do
  begin
    TOBL := fCurrentOnglet.TOBLignes.Detail[i];
    if i = 0 then iTot := TOBL.GetNumChamp('GL_TYPELIGNE');
    CalculeSousTotal (i+1,iTot);
  end;
end;

procedure TOF_BTSAISIEAVANC.SommerLignesAV(ARow, Niv: integer);
var TOBL, TOBS: TOB;
  i : integer;
  TypL : string;
  // VARIANTE
  variante : boolean;
  Pourcent : double;
begin
  TOBS := GetTOBLigneAV(fCurrentOnglet.TOBlignes,ARow);
  if TOBS = nil then Exit;
  variante := IsVariante (TOBS);
  //
  TOBS.PutValue('BLF_MTMARCHE', 0);
  TOBS.PutValue('BLF_MTDEJAFACT', 0);
  TOBS.PutValue('BLF_MTCUMULEFACT', 0);
  TOBS.PutValue('BLF_MTSITUATION', 0);
  TOBS.PutValue('BLF_MTPRODUCTION', 0);
  //
  //TOBS.putvalue('MONTANTSIT', 0);

  // ---------
  for i := ARow - 1 downto 1 do
  begin
    TOBL := GetTOBLigneAV(fCurrentOnglet.TOBlignes,i);
    if TOBL = nil then Break;
    TypL := TOBL.GetValue('GL_TYPELIGNE');
    if (TOBL.GetValue('GL_NATUREPIECEG')<>'PBT') and (isVariante(TOBL)) and not variante then continue;
    if Niv = 0 then
    begin
      if ((TypL = 'TOT') or (TypL = 'DEB')) then Break;
    end
    else
    begin
      if IsParagraphe (TOBL,niv) then
      begin
        break;
      end;
    end;
    // Modif BTP
    if TYPL = 'RG' then continue;
    // ---
    if Pos(typL,'ART;OUP') <= 0 then Continue;
    {$IFDEF BTP}
    TOBS.PutValue('GL_TPSUNITAIRE', TOBS.GetValue('GL_TPSUNITAIRE') + (TOBL.GetValue('GL_TPSUNITAIRE') * TOBL.getValue('GL_QTEFACT')));
    TOBS.PutValue('BLF_MTMARCHE', TOBS.GetValue('BLF_MTMARCHE') + TOBL.GetValue('BLF_MTMARCHE'));
    TOBS.PutValue('BLF_MTDEJAFACT', TOBS.GetValue('BLF_MTDEJAFACT') + TOBL.GetValue('BLF_MTDEJAFACT'));
    TOBS.PutValue('BLF_MTCUMULEFACT', TOBS.GetValue('BLF_MTCUMULEFACT') + TOBL.GetValue('BLF_MTCUMULEFACT'));
    TOBS.PutValue('BLF_MTSITUATION', TOBS.GetValue('BLF_MTSITUATION') + TOBL.GetValue('BLF_MTSITUATION'));
    TOBS.PutValue('BLF_MTPRODUCTION', TOBS.GetValue('BLF_MTPRODUCTION') + TOBL.GetValue('BLF_MTPRODUCTION'));
    {$ENDIF}
  end;
  if TOBS.GetValue('BLF_MTMARCHE')<> 0 then
  begin
    if Avancement then
    begin
      Pourcent := Arrondi(TOBS.GetValue('BLF_MTCUMULEFACT')/TOBS.GetValue('BLF_MTMARCHE')*100,2);
    end else
    begin
      Pourcent := Arrondi(TOBS.GetValue('BLF_MTSITUATION')/TOBS.GetValue('BLF_MTMARCHE')*100,2);
    end;
    TOBS.PutValue('BLF_POURCENTAVANC', Pourcent);
  end;

  if Niv = 0 then
  begin
    TOBS.PutValue('GL_TYPELIGNE', 'TOT');
    TOBS.PutValue('GL_TYPEDIM', 'NOR');
    if TOBS.GetValue('GL_LIBELLE') = '' then TOBS.PutValue('GL_LIBELLE', TraduireMemoire('Sous-total'));
  end;
end;

procedure TOF_BTSAISIEAVANC.AppliqueVoirMt;
begin
  GS.BeginUpdate;
  if VOIRMT.Checked then
  begin
    GS.ColLengths [MTMARCHE] := LNMTMARCHE;
    GS.ColLengths [MTCUMULEFACT] := LNMTCUMULEFACT; GS.ColEditables  [MTCUMULEFACT] := false;
    GS.ColLengths [MTDEJAFACT] := LNMTDEJAFACT;
    //
    GS.ColWidths[MTMARCHE] := LNMTMARCHE* GS.Canvas.TextWidth('W');
    GS.ColWidths[MTCUMULEFACT] := LNMTCUMULEFACT* GS.Canvas.TextWidth('W');
    GS.ColWidths[MTDEJAFACT] := LNMTDEJAFACT* GS.Canvas.TextWidth('W');
  end else
  begin
    if (VOIRQTE.Checked) then
    begin
      GS.ColLengths [MTMARCHE] := -1;
      GS.ColWidths[MTMARCHE] := 0;
    end;
    GS.ColLengths [MTCUMULEFACT] := -1; GS.ColEditables  [MTCUMULEFACT] := false;
    GS.ColLengths [MTDEJAFACT] := -1;
    //
    GS.ColWidths[MTCUMULEFACT] := 0;
    GS.ColWidths[MTDEJAFACT] := 0;
  end;
  GS.VidePile(false);
  AfficheLaGrille(GS, fCurrentOnglet.TOBlignes);
  GS.EndUpdate;
  GS.Refresh;
end;

procedure TOF_BTSAISIEAVANC.VoirMtCLick(Sender: Tobject);
var Arow : integer;
		LastOnglet : TOnglet;
    Indice : integer;
begin
  Arow := GS.row;
  lastonglet := fCurrentOnglet;
  for Indice := 0 to fOnglets.Count -1 do
  begin
    fCurrentOnglet := fOnglets.items[Indice];
    GS := fCurrentOnglet.GS;
  	AppliqueVoirMt;
  end;
  fCurrentOnglet := LastOnglet;
  GS := fCurrentOnglet.GS;
  SetLigne (GS,Arow,LIBELLE);
//  TFVierge(Self).HMTrad.ResizeGridColumns (GS);
end;

procedure TOF_BTSAISIEAVANC.AppliqueVoirQte ;
begin
  GS.BeginUpdate;
  if VOIRQTE.Checked then
  begin
    GS.ColLengths [QTEMARCHE] := LNQTEMARCHE;
    GS.ColLengths [QTECUMULFACT] := LNQTECUMULFACT; GS.ColEditables  [QTECUMULFACT] := false;
    GS.ColLengths [QTEDEJAFACT] := LNQTEDEJAFACT;
    GS.ColLengths [QTESITUATION] := LNQTESITUATION;
    //
    GS.ColWidths[QTEMARCHE] := LNQTEMARCHE*GS.Canvas.TextWidth('W');
    GS.ColWidths[QTECUMULFACT] := LNQTECUMULFACT*GS.Canvas.TextWidth('W');
    GS.ColWidths[QTEDEJAFACT] := LNQTEDEJAFACT*GS.Canvas.TextWidth('W');
    GS.ColWidths[QTESITUATION] := LNQTESITUATION*GS.Canvas.TextWidth('W');
  end else
  begin
    if (not VOIRMT.Checked) then
    begin
      GS.ColLengths [MTMARCHE] := LNMTMARCHE;
      GS.ColWidths[MTMARCHE] := LNMTMARCHE*GS.Canvas.TextWidth('W');
    end;

    GS.ColLengths [QTEMARCHE] := -1;
    GS.ColLengths [QTECUMULFACT] := -1; GS.ColEditables  [QTECUMULFACT] := false;
    GS.ColLengths [QTEDEJAFACT] := -1;
    GS.ColLengths [QTESITUATION] := -1;
    //
    GS.ColWidths[QTEMARCHE] := 0;
    GS.ColWidths[QTECUMULFACT] := 0;
    GS.ColWidths[QTEDEJAFACT] := 0;
    GS.ColWidths[QTESITUATION] := 0;
  end;
  GS.VidePile(false);
  AfficheLaGrille(GS,fCurrentOnglet.TOBlignes);
  GS.EndUpdate;
  GS.Refresh;
end;

procedure TOF_BTSAISIEAVANC.VoirQteCLick(Sender: Tobject);
var Arow : integer;
		indice : integer;
    lastonglet : TOnglet;
begin
  Arow := GS.row;
  lastonglet := fCurrentOnglet;
  for Indice := 0 to fOnglets.Count -1 do
  begin
    fCurrentOnglet := fOnglets.items[Indice];
    GS := fCurrentOnglet.GS;
  	AppliqueVoirQte;
  end;
  fCurrentOnglet := LastOnglet;
  GS := fCurrentOnglet.GS;
  SetLigne (GS,Arow,LIBELLE);
//  TFVierge(Self).HMTrad.ResizeGridColumns (GS);
end;

function TOF_BTSAISIEAVANC.ChangeQteSituationOk(Ligne: integer;valeur: double): boolean;
var TOBL : TOB;
begin
  result := true;
  TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,ligne);
  if Avancement then
  begin
    if ModifSituation then
    begin
      if Valeur < 0 then
      begin
        result := (PgiAsk ('Etes-vous sur d''indiquer une quantité < 0 ?','Ligne : '+InttoStr(Ligne))=Mryes);
      end else if (Valeur + TOBL.GetValue('QTECUMULEFACTINIT'))> TOBL.GetValue('BLF_QTEMARCHE') then
      begin
        // Modified by f.vautrain 07/11/2017 17:07:19 - FS#2774 - GUINIER - Interdire la possibilité de facturer supérieur à 100 %
        {$IFDEF V10}
        if CtrlSupCent then
        begin
          PGIError('Quantité supérieure à la quantité du devis','Ligne : '+InttoStr(Ligne));
          Result := False;
        end
        else
        {$ENDIF}
          result := (PgiAsk ('Etes-vous sur d''indiquer une quantité supérieure à la quantité du devis ?','Ligne : '+InttoStr(Ligne))=Mryes);
      end;
    end else
    begin
    if Valeur < TOBL.GetValue('QTECUMULEFACTINIT') then
    begin
      result := (PgiAsk ('Etes-vous sur d''indiquer une quantité cumulée facturé inférieure à la précédente facturation ?','Ligne : '+InttoStr(Ligne))=Mryes);
    end else if Valeur > TOBL.GetValue('BLF_QTEMARCHE') then
    begin
        // Modified by f.vautrain 07/11/2017 17:07:19 - FS#2774 - GUINIER - Interdire la possibilité de facturer supérieur à 100 %
        {$IFDEF V10}
        if CtrlSupCent then
        begin
          PGIError('Quantité cumulée facturée supérieure à la quantité du devis','Ligne : '+InttoStr(Ligne));
          Result := False;
        end
        else
        {$ENDIF}
          result := (PgiAsk ('Etes-vous sur d''indiquer une quantité cumulée facturée supérieure à la quantité du devis ?','Ligne : '+InttoStr(Ligne))=Mryes);
    end;
    end;
  end
  else
  begin
    if (Valeur+TOBL.GetValue('BLF_QTEDEJAFACT')) > TOBL.GetValue('BLF_QTEMARCHE') then
    begin
      // Modified by f.vautrain 07/11/2017 17:07:19 - FS#2774 - GUINIER - Interdire la possibilité de facturer supérieur à 100 %
      {$IFDEF V10}
      if CtrlSupCent then
      begin
        PGIError('Quantité cumulée facturée supérieure à la quantité du devis','Ligne : '+InttoStr(Ligne));
        Result := False;
      end
      else
      {$ENDIF}
        result := (PgiAsk ('Etes-vous sur d''indiquer une quantité cumulée facturé supérieure à la quantité du devis ?','Ligne : '+InttoStr(Ligne))=Mryes);
    end;
  end;
end;

function TOF_BTSAISIEAVANC.ChangePourcentOk (Ligne : integer; valeur : double) : boolean;
var TOBL : TOB;
    MtFacturation,MtMarche,MtDejafact : double;
begin
  result := true;
  TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,ligne);
  if Avancement then
  begin
    if Valeur < TOBL.GetValue('POURCENTAVANCINIT') then
    begin
      result := (PgiAsk ('Etes-vous sur d''indiquer un pourcentage d''avancement inférieur à la précédente facturation ?','Ligne : '+InttoStr(Ligne))=Mryes);
    end else if Valeur > 100.0 then
    begin
      // Modified by f.vautrain 07/11/2017 15:53:27 - FS#2774 - GUINIER - Interdire la possibilité de facturer supérieur à 100 %
      {$IFDEF V10}
      if CtrlSupCent then
      begin
        PGIError('Pourcentage d''avancement supérieur à 100 %', 'Ligne : '+ InttoStr(Ligne));
        Result := False;
      end
      else
      {$ENDIF}
        result := (PgiAsk ('Etes-vous sur d''indiquer un pourcentage d''avancement supérieur à 100% ?','Ligne : '+InttoStr(Ligne))=Mryes);
    end;
  end else
  begin
    MtFacturation := Arrondi(TOBL.GetValue('BLF_MTMARCHE')*Valeur/100,DEV.Decimale);
    MtDejaFact := TOBL.GetValue('BLF_MTDEJAFACT');
    MtMarche := TOBL.GetValue('BLF_MTMARCHE');
    if (Valeur > 100.0) or (arrondi(MtDejaFact+MtFacturation-MtMarche,DEV.decimale) > 0) then
    begin
      // Modified by f.vautrain 07/11/2017 15:53:27 - FS#2774 - GUINIER - Interdire la possibilité de facturer supérieur à 100 %
      {$IFDEF V10}
      if CtrlSupCent then
      begin
        PGIError('Impossible de facturer plus que le montant du marché', 'Ligne : '+ InttoStr(Ligne));
        Result := False;
      end
      else
      {$ENDIF}
        result := (PgiAsk ('Etes-vous sur de vouloir facturer plus que le montant du marché ?','Ligne : '+InttoStr(Ligne))=Mryes);
    end;
  end;
end;

function TOF_BTSAISIEAVANC.ChangePourcentOkMulti (Ligne : integer; valeur : double; var LibelleErreur : string) : integer;
var TOBL : TOB;
    MtFacturation,MtMarche,MtDejafact : double;
begin
  result := 0;
  TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,ligne);
  if Avancement then
  begin
    if TOBL.GetValue('BLF_MTMARCHE')<>0 then
    begin
    if Valeur < TOBL.GetValue('POURCENTAVANCINIT') then
    begin
      result := -1;
      LibelleErreur := 'Avancement inférieur à la précédente facturation';
      end
      else if Valeur > 100.0 then
    begin
      result := 1;
      LibelleErreur := 'avancement supérieur à 100%';
    end;
    end;
  end
  else
  begin
    MtMarche := TOBL.GetValue('BLF_MTMARCHE');
    if MtMarche <> 0 then
    begin
    MtFacturation := Arrondi(TOBL.GetValue('BLF_MTMARCHE')*Valeur/100,DEV.Decimale);
    MtDejaFact := TOBL.GetValue('BLF_MTDEJAFACT');
    if (Valeur > 100.0) or (MtDejaFact+MtFacturation > MtMarche) then
    begin
      result := 1;
      LibelleErreur := 'facturation supérieur au montant du marché';
    end;
  end;
  end;
end;

procedure TOF_BTSAISIEAVANC.ChangeMTSituationAvanc(Montant : double; TOBL: TOB;CurrOnglet : Tonglet; var Cancel: boolean);
begin
  CalculeLaLigneFromMontant (Montant,TOBL,CurrOnglet);
end;

function TOF_BTSAISIEAVANC.ChangeMTSituationok (Ligne : integer;Valeur : double) : boolean;
var TOBL : TOB;
    MtMarche,MtDejafact : double;
begin
  result := true;
  TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,ligne);
  if Avancement then
  begin
    if Valeur < 0 then
    begin
      result := (PgiAsk ('Etes-vous sur d''indiquer un avancement inférieur à la précédente facturation ?','Ligne : '+InttoStr(Ligne))=Mryes);
    end
    else if (TOBL.getValue('BLF_MTMARCHE')) and (Valeur + TOBL.GetValue('BLF_MTDEJAFACT') > TOBL.GetValue('BLF_MTMARCHE')) then
    begin
      // Modified by f.vautrain 07/11/2017 17:03:12 - FS#2774 - GUINIER - Interdire la possibilité de facturer supérieur à 100 %
      {$IFDEF V10}
      if CtrlSupCent then
      begin
        PGIError('Montant de facturation supérieur à celui du marché', 'Ligne : '+ InttoStr(Ligne));
        result := false;
      end
      else
      {$ENDIF}
        result := (PgiAsk ('Etes-vous sur d''indiquer un montant de facturation supérieur à celui du marché ?','Ligne : '+InttoStr(Ligne))=Mryes);
    end;
  end else
  begin
    MtDejaFact := TOBL.GetValue('BLF_MTDEJAFACT');
    MtMarche := TOBL.GetValue('BLF_MTMARCHE');
    if (Valeur+MtDejaFact > MtMarche) then
    begin
      // Modified by f.vautrain 07/11/2017 17:02:59 -  FS#2774 - GUINIER - Interdire la possibilité de facturer supérieur à 100 %
      {$IFDEF V10}
      if CtrlSupCent then
      begin
        PGIError('Montant de facturation supérieur à celui du marché','Ligne : '+ InttoStr(Ligne));
        result := false;
      end
      else
      {$ENDIF}
        result := (PgiAsk ('Etes-vous sur d''indiquer un montant de facturation supérieur à celui du marché ?','Ligne : '+InttoStr(Ligne))=Mryes);
    end;
  end;
end;

procedure TOF_BTSAISIEAVANC.Constitue_TVDEVIS;
var TOBL: TOB;
    TOBC,TOBCL : TOB;
  i: Integer;
  Tn,RootTN: TTreeNode;
  Titre: string;
begin
  TOBC := TOB.Create ('LA TOB ', nil,-1);
  TRY
    // chargement treeview
    fCurrentOnglet.fTVDevis.items.clear;
    Titre := 'Préparation de facture';
    RootTN := fCurrentOnglet.fTVDevis.items.add(nil, Titre);
    TOBCL := TOB.Create ('UN ITEM',TOBC,-1);
    TOBCL.data := RootTN;

    for i := 0 to fCurrentOnglet.TOBLIGNES.Detail.Count - 1 do
    begin
      TOBL := fCurrentOnglet.TOBLignes.Detail[i];
      // VARIANTE
      if IsPiece (TOBL) then
      begin
        TN := fCurrentOnglet.fTVDEVIS.Items.AddChild (RootTN,TOBL.GetValue('GL_LIBELLE'));
        Tn.Data := TOBL;
        TOBC.clearDetail;
        TOBCL := TOB.Create ('UN ITEM',TOBC,-1);
        TOBCL.data := TN;
      end else
      begin
        if (IsDebutParagraphe (TOBL)) then
        begin
          //FV1 : 28/05/2014 - FS#1012 - VEODIS GROUPE BALIN - Message d'erreur indice de liste hors limite en saisie des avancements
          if TOBC.detail.count > 0 then
          begin
            Tn := TVDEVIS.Items.AddChild(TTreeNode(TOBC.detail[TOBC.detail.count-1].data), TOBL.GetValue('GL_LIBELLE'));
            Tn.Data := TOBL;
          end;
          TOBCL := TOB.Create ('UN ITEM',TOBC,-1);
          TOBCL.data := TN;
        end else if IsFinParagraphe (TOBL) then
        begin
          TOBC.detail[TOBC.detail.count-1].free;
        end;
      end;
    end;
  FINALLY
    TOBC.free;
  END;
  fCurrentOnglet.fTVDEVIS.FullExpand;
  fCurrentOnglet.fTVDEVIS.OnClick :=  TVDEVISClick;
end;


function TOF_BTSAISIEAVANC.IsPiece(TOBL: TOB): boolean;
begin
  result := false;
  if (TOBL.GetValue('GL_TYPELIGNE')='REF') then result := true;
end;

procedure TOF_BTSAISIEAVANC.BTDescriptifClick(Sender: TObject);
begin
  ToolWin.Visible := BTDescriptif.Down;
end;

procedure TOF_BTSAISIEAVANC.TVDEVISClick(Sender: TObject);
var Tn: TTreeNode;
  TOBL: TOB;
  Acol, Arow: integer;
begin
  Tn := TVDEVIS.selected;
  TOBL := TOB(Tn.data);
  if (TOBL <> nil) then
  begin
    ARow := TOBL.GetValue('INDICE');
    ACol := LIBELLE;
    GS.SetFocus;
    GS.Row := Arow;
    GS.col := Acol;
    stprev := GS.Cells[GS.Col, GS.Row];
  	fCurrentOnglet.stPrev := stprev;
  end;
end;

procedure TOF_BTSAISIEAVANC.ToolWinClose (Sender : TObject);
begin
  BTDescriptif.down := false;
end;

procedure TOF_BTSAISIEAVANC.AddChampsFac;
begin
//  TOBLIGNES.AddChampSupValeur ('AFF_GENERAUTO','');
end;

procedure TOF_BTSAISIEAVANC.SetInfoFac (TOBL : TOB);
begin
  TOBLIGNES.dupliquer (TOBL,false,true);
//  TOBLIGNES.PutValue('GP_AFFAIREDEVIS',TOBL.getValue('GP_AFFAIREDEVIS'));
//  TOBLIGNES.PutValue('AFF_GENERAUTO',TOBL.getValue('AFF_GENERAUTO'));
end;

procedure TOF_BTSAISIEAVANC.GenereLafacture (DateFac : TdateTime; Var Clefac : R_Cledoc; GenereAvoir : boolean=false);
begin
  if not GenereAvoir then
  begin
    if VH_GC.BTCODESPECIF = '001' then
    begin
      // Mise en place de la RG et des cautions bancaires
      // calcul de la restitution d'avances suivant les modalités POC
    end else
    begin
    // visu de l'avancement de l'acompte
      SetAvancementAcompte;
    end;
  end;
  if not GenereFactureFromAvanc (TOBSOUSTraits,TOBLignes,TOBSelect,TOBOuvrages,
  															 TOBAcomptes_O,LESTOBACOMPTES,ModeCloture,ErreurEcr,Clefac,DateFac,
      													 MontantMarche,DejaFacture,MtHorsDevis,
                                 flastfacture,fCotraitance,ModifSituation,GenereAvoir) then
    V_PGI.ioerror := OeUnknown;
end;

function TOF_BTSAISIEAVANC.ExisteLigneSituation: boolean;
var indice : integer;
begin
  result := false;
  for Indice := 0 to TOBLignes.detail.count - 1 do
  begin
    if TOBLignes.detail[Indice].getValue('BLF_MTSITUATION')<>0 then
    begin
      result := true;
      break;
    end;
  end;
end;

procedure TOF_BTSAISIEAVANC.FormKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
var OkG, Vide: Boolean;
begin
  OkG := (Screen.ActiveControl = GS);
  Vide := (Shift = []);
  case Key of
    VK_RETURN: if ((OkG) and (Vide)) then
      begin
        SendMessage(GS.Handle, WM_KEYDOWN , VK_TAB, 0)
      end;
    VK_F10: if Vide then
      begin
        Bvalider.OnClick (self);
      end;
    27:
      begin
        Bferme.OnClick (self);
      end;
  end;
end;

procedure TOF_BTSAISIEAVANC.ReinitLigneSD (TOBL : TOB);
var IndiceNomen : Integer;
		II : integer;
    TOBOO,TOBOUV : TOB;
    MtDejaFact,QteDejaFact,PourcentAvanc : double;
begin
  IndiceNomen := TOBL.getValue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
  if TOBL.GetValue('UNIQUEBLO') = 0 then exit;
  //FV1 : 28/07/2016
  if TobOuvrages.detail.count = 0 then exit;
  if (Indicenomen-1) > (TobOuvrages.detail.count - 1) then exit;
  //
  TOBOUV := TOBOuvrages.detail[IndiceNomen-1];
  for II := 0 to TOBOUV.detail.count -1 do
  begin
  	TOBOO := TOBOUV.detail[II];
    if TOBOO.getValue('BLO_UNIQUEBLO')=TOBL.GetValue('UNIQUEBLO') then
    begin
      MtDejaFact := TOBL.getValue('MTCUMULEFACTINIT');
      QteDejaFact := TOBL.GetValue ('QTECUMULEFACTINIT');
      PourcentAvanc := TOBL.GetValue ('POURCENTAVANCINIT');
      CumuleLigne(TOBL,'-');
      TOBL.PutValue('BLF_POURCENTAVANC',PourcentAvanc);
      TOBL.putValue('BLF_MTDEJAFACT',MtDejaFact);
      TOBL.putValue('BLF_MTCUMULEFACT',MtDejaFact);
      TOBL.putValue('BLF_MTSITUATION',0);
      TOBL.putValue('BLF_QTEDEJAFACT',QteDejaFact);
      TOBL.putValue('BLF_QTECUMULEFACT',QteDejaFact);
      TOBL.putValue('BLF_QTESITUATION',0);
    	break;
    end;
  end;
end;

procedure TOF_BTSAISIEAVANC.ReInitLigne(TOBL : TOB);
var  QteDejaFact,MtDejaFact,PourcentAvanc : double;
begin
  MtDejaFact    := TOBL.getValue('MTCUMULEFACTINIT');
  QteDejaFact   := TOBL.GetValue ('QTECUMULEFACTINIT');
  PourcentAvanc := TOBL.GetValue ('POURCENTAVANCINIT');
  CumuleLigne(TOBL,'-');
  TOBL.PutValue('BLF_POURCENTAVANC',PourcentAvanc);
  TOBL.putValue('BLF_MTDEJAFACT',MtDejaFact);
  TOBL.putValue('BLF_MTCUMULEFACT',MtDejaFact);
  TOBL.putValue('BLF_MTSITUATION',0);
  TOBL.putValue('BLF_QTEDEJAFACT',QteDejaFact);
  TOBL.putValue('BLF_QTECUMULEFACT',QteDejaFact);
  TOBL.putValue('BLF_QTESITUATION',0);
  TOBL.putValue('BLF_POURCENTAVANC',PourcentAvanc);
  if IsSousDetail(TOBL) and (fCurrentOnglet = FOngletGlobal) then
  begin
  	ReinitLigneSD (TOBL);
  end;
  AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,TOBL.getvalue('INDICE'));
  if (IsOuvrage(TOBL)) and (not IsSousDetail(TOBL)) and (TOBL.getValue('UNIQUEBLO')=0) then
  begin
  	defaireLigOuvSaisie (TOBL);
  end;
end;

procedure TOF_BTSAISIEAVANC.defaireLigOuvSaisie(TOBO: TOB);
var ListeDet : TlisteTOB;
    TOBL,TOBOUV : TOB;
    Indice,IndiceNomen : integer;
    MtDejafact,QteDejaFact,pourcentavanc : double;
begin
  IndiceNomen := TOBO.GetValue('GL_INDICENOMEN');
  //FV1 : 28/07/2016
  if TobOuvrages.detail.count = 0 then exit;
  if (Indicenomen-1) > (TobOuvrages.detail.count - 1) then exit;
  //
  if IndiceNOmen > 0 then
  begin
    //
    TOBOUV := TOBOuvrages.detail[IndiceNomen-1];
    for Indice:= 0 to TOBOUV.detail.count -1 do
    begin
			TOBL := TOBOUV.detail[Indice];
      MtDejaFact := TOBL.getValue('MTCUMULEFACTINIT');
      QteDejaFact := TOBL.GetValue ('QTECUMULEFACTINIT');
      PourcentAvanc := TOBL.GetValue ('POURCENTAVANCINIT');
      TOBL.PutValue('BLF_POURCENTAVANC',PourcentAvanc);
      TOBL.putValue('BLF_MTDEJAFACT',MtDejaFact);
      TOBL.putValue('BLF_MTCUMULEFACT',MtDejaFact);
      TOBL.putValue('BLF_MTSITUATION',0);
      TOBL.putValue('BLF_QTEDEJAFACT',QteDejaFact);
      TOBL.putValue('BLF_QTECUMULEFACT',QteDejaFact);
      TOBL.putValue('BLF_QTESITUATION',0);
      TOBL.putValue('BLF_POURCENTAVANC',PourcentAvanc);
    end;
  end;
  //
	ListeDet := FindlignesOuvSaisie (fCurrentOnglet,TOBO);
  if ListeDet.count > 0 then
  begin
  	for Indice := 0 to ListeDet.count -1 do
    begin
    	TOBL := ListeDet.Items [Indice];
      MtDejaFact := TOBL.getValue('MTCUMULEFACTINIT');
      QteDejaFact := TOBL.GetValue ('QTECUMULEFACTINIT');
      PourcentAvanc := TOBL.GetValue ('POURCENTAVANCINIT');
      TOBL.PutValue('BLF_POURCENTAVANC',PourcentAvanc);
      TOBL.putValue('BLF_MTDEJAFACT',MtDejaFact);
      TOBL.putValue('BLF_MTCUMULEFACT',MtDejaFact);
      TOBL.putValue('BLF_MTSITUATION',0);
      TOBL.putValue('BLF_QTEDEJAFACT',QteDejaFact);
      TOBL.putValue('BLF_QTECUMULEFACT',QteDejaFact);
      TOBL.putValue('BLF_QTESITUATION',0);
      TOBL.putValue('BLF_POURCENTAVANC',PourcentAvanc);
      if IsSousDetail(TOBL) and (fCurrentOnglet = FOngletGlobal) then
      begin
        ReinitLigneSD (TOBL);
      end;
    	AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,TOBL.getvalue('INDICE'));
    end;
  end;
  ListeDet.free;
end;

procedure TOF_BTSAISIEAVANC.DefaireLig(Sender: TObject);
var TOBL : TOB;
		UnOngletArecalc : Tonglet;
begin
	UnOngletArecalc := nil;
  TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,GS.row);
  ReInitLigne(TOBL);
  if fCurrentOnglet <> FOngletGlobal then
  begin
    ReInitLigneGlobal (TOBL);
  end else
  begin
    ReInitLigneOnglets (TOBL);
  end;
  AfficheLaLigne (fCurrentOnglet.GS,fCurrentOnglet.TOBlignes,GS.row);
  stPrev := GS.cells[Gs.col,GS.row];
  fCurrentOnglet.stPrev := stprev;
  CalculeLesSousTotaux;
  GS.Invalidate;
  if fCurrentOnglet <> FOngletGlobal then
  begin
    ReCalculeglobal;
  end else
  begin
  	if UnOngletArecalc <> nil then RecalculeOnglet (UnOngletArecalc,true);
  end;
end;

procedure TOF_BTSAISIEAVANC.DefairePiece(Sender: TObject);
var Debut,Fin : integer;
begin
  //
  Debut := RecupDebutPieceCur (GS.row);
  Fin := RecupFinPieceCur (GS.row);
  //
  DefaireSegment (Debut,Fin);
end;


procedure TOF_BTSAISIEAVANC.ReInitLigneGlobal (TOBD : TOB);
var NumOrdre,UniqueBlo,NumPiece : integer;
		LastOnglet : TOnglet;
    Indice : integer;
    TOBL : TOB;
    ListeDet : TlisteTOB;
begin
  if (TOBD.GetValue('GL_TYPELIGNE')='OUP') and (IsOuvrage(TOBD)) then
  begin
    ListeDet := FindlignesOuvSaisie (fCurrentOnglet,TOBD);
    if ListeDet.count > 0 then   // lignes de détail trouvé
    begin
      for Indice := 0 to ListeDet.count -1 do
      begin
        ReInitLigneGlobal (ListeDet.Items [Indice]);
      end;
    end;
    ListeDet.free;
    Exit;
  end;

  NumOrdre := TOBD.getvalue('GL_NUMORDRE');
  UniqueBlo := TOBD.getvalue('UNIQUEBLO');
  NumPiece := TOBD.getvalue('GL_NUMERO');
  lastOnglet := fCurrentOnglet;
  SetOnglet(FOngletGlobal);
  indice := FindTOBligne (FongletGlobal,NumPiece,Numordre,uniqueBlo);
  if indice < 0 then exit;
  TOBL := GetTOBLigneAV(fCurrentOnglet.TOBlignes ,indice+1);
  if TOBL <> nil then
  begin
    ReInitLigne(TOBL);
  end;
  SetOnglet(lastOnglet);
end;


function TOF_BTSAISIEAVANC.ReInitLigneFournisseur (TOBD : TOB) : Tonglet;
var NumOrdre,UniqueBlo,NumPiece : integer;
		LastOnglet,TheOnglet : TOnglet;
    Indice : integer;
    TOBL : TOB;
begin
	result := nil;
  NumOrdre := TOBD.getvalue('GL_NUMORDRE');
  UniqueBlo := TOBD.getvalue('UNIQUEBLO');
  NumPiece := TOBD.getvalue('GL_NUMERO');
  lastOnglet := fCurrentOnglet;
  indice := -1;
  TheOnglet := FindLigneInOnglet(NumPiece,NumOrdre,UniqueBlo,indice);
  if TheOnglet = nil then exit;
  result := TheOnglet;
  SetOnglet(TheOnglet);
  TOBL := GetTOBLigneAV(fCurrentOnglet.TOBlignes ,indice+1);
  if TOBL <> nil then
  begin
    ReInitLigne(TOBL);
  end;
  SetOnglet(lastOnglet);
end;

procedure TOF_BTSAISIEAVANC.ReInitLigneOnglets (TOBD : TOB);
var NumOrdre,UniqueBlo,NumPiece : integer;
		LastOnglet,TheOnglet : TOnglet;
    Indice : integer;
    TOBL : TOB;
    ListeDet : TlisteTOB;
begin
  if IsOuvrage(TOBD) and (TOBD.GetValue('GLC_VOIRDETAIL')='X') then
  begin
    ListeDet := FindlignesOuvSaisie (fCurrentOnglet,TOBD);
    if ListeDet.count > 0 then   // lignes de détail trouvé
    begin
      for Indice := 0 to ListeDet.count -1 do
      begin
        ReInitLigneOnglets (ListeDet.Items [Indice]);
      end;
    end;
    ListeDet.free;
    Exit;
  end;
  NumOrdre := TOBD.getvalue('GL_NUMORDRE');
  UniqueBlo := TOBD.getvalue('UNIQUEBLO');
  NumPiece := TOBD.getvalue('GL_NUMERO');
  lastOnglet := fCurrentOnglet;
  indice := -1;
  TheOnglet := FindLigneInOnglet(NumPiece,NumOrdre,UniqueBlo,indice);
  if TheOnglet = nil then exit;
  SetOnglet(TheOnglet);
  TOBL := GetTOBLigneAV(fCurrentOnglet.TOBlignes ,indice+1);
  if TOBL <> nil then
  begin
    ReInitLigne(TOBL);
  end;
  SetOnglet(lastOnglet);
end;


procedure TOF_BTSAISIEAVANC.RecalculeOnglet (unOnglet : Tonglet;WithSaveBefore : boolean=false);
var lastOnglet : Tonglet;
begin
	lastOnglet := nil;
	if WithSaveBefore then lastOnglet := fCurrentOnglet;
  SetOnglet(unOnglet);
  CalculeLesSousTotaux;
  AfficheLaGrille (GS,fCurrentOnglet.TOBlignes);
  stPrev := GS.cells[Gs.col,GS.row];
  fCurrentOnglet.stPrev := stprev;
  GS.Invalidate;
  if WithSaveBefore then SetOnglet(Lastonglet);
end;

procedure TOF_BTSAISIEAVANC.ReCalculelesOnglets (TheListe: TstringList);
var indice : integer;
		UnOnglet,LastOnglet : Tonglet;
begin
	lastOnglet := fCurrentOnglet;
	for Indice := 0 to TheListe.Count -1 do
  begin
    if TheListe.Strings [indice] <> '' then
    begin
    	UnOnglet := fOnglets.findOnglet(TheListe.strings[indice]);
      if UnOnglet <> nil then
      begin
      	RecalculeOnglet (unOnglet);
      end;
    end;
  end;
  SetOnglet(LastOnglet);
end;

procedure TOF_BTSAISIEAVANC.ReCalculeglobal;
var LastOnglet : Tonglet;
begin
	lastOnglet := fCurrentOnglet;
  SetOnglet(FOngletGlobal);
  CalculeLesSousTotaux;
  AfficheLaGrille (GS,fCurrentOnglet.TOBlignes);
  stPrev := GS.cells[Gs.col,GS.row];
  fCurrentOnglet.stPrev := stprev;
  GS.Invalidate;
  SetOnglet(LastOnglet);
end;

function TOF_BTSAISIEAVANC.FindNextLigne (TOBLigne ,TOBO : TOB; Indice : integer) : integer;
var numOrdre : integer;
		ind : integer;
    TOBL : TOB;
begin
	Numordre := TOBO.getValue('GL_NUMORDRE');
  Ind := indice;
  repeat
  	TOBL := TOBligne.detail[Ind];
    if numordre <> TOBL.getValue('GL_NUMORDRE') then break;
    inc(Ind);
  until ind >= TOBLigne.detail.count -1;
  if Ind = Indice then inc(ind); // cas d el'ouvrage sans sou sdétail
  result := ind;
end;

procedure TOF_BTSAISIEAVANC.DefaireSegment (Debut,Fin : integer);
var Indice : integer;
    TOBL : TOB;
    TheListetoRecalc : TStringList;
begin
	TheListetoRecalc := TStringList.Create;

  indice := debut;
  repeat
    TOBL := GetTOBLigneAV (fCurrentOnglet.TOBlignes,Indice);
    if (Pos(TOBL.GetValue('GL_TYPELIGNE'),'ART;OUP')>0) then
    begin
      ReInitLigne(TOBL);
      if (isOuvrage (TOBL)) and (not IsSousDetail(TOBL)) and (TOBL.getValue('UNIQUEBLO')=0) then
      begin
      	indice := FindNextLigne (fCurrentOnglet.TOBlignes,TOBL,Indice);
      end else
      begin
      	inc(Indice);
      end;
      if fCurrentOnglet <> FOngletGlobal then
      begin
        ReInitLigneGlobal (TOBL);
      end else
      begin
    		ReInitLigneOnglets (TOBL);
        (*
        UnOngletArecalc := ReInitLigneFournisseur (TOBL);
        if UnOngletArecalc <> nil then
        begin
        	if not TheListetoRecalc.Find (UnOngletArecalc.Name,index) then
          begin
          	TheListetoRecalc.Add(UnOngletArecalc.Name);
          end;
        end;
        *)
      end;
    end else inc(Indice);

  until indice >= Fin;
  CalculeLesSousTotaux;
  AfficheLaGrille (GS,fCurrentOnglet.TOBlignes);
  stPrev := GS.cells[Gs.col,GS.row];
  fCurrentOnglet.stPrev := stprev;
  GS.Invalidate;
  if fCurrentOnglet <> FOngletGlobal then
  begin
    ReCalculeglobal;
  end else
  begin
    ReCalculelesOnglets (TheListetoRecalc);
  end;
  TheListetoRecalc.free;
end;

procedure TOF_BTSAISIEAVANC.DefaireTous(Sender: TObject);
var Debut,Fin: integer;
begin
  Debut := 1;
  Fin := Gs.rowCount;
  DefaireSegment (Debut,Fin);
end;


procedure TOF_BTSAISIEAVANC.RecupAcomptePiece(TOBS: TOB);
var Sql : string;
begin
  Sql := 'SELECT * FROM ACOMPTES WHERE GAC_NATUREPIECEG="'+TOBS.GetValue('GP_NATUREPIECEG')+'" AND '+
         'GAC_SOUCHE="'+TOBS.GetValue('GP_SOUCHE')+'" AND '+
         'GAC_NUMERO='+IntToStr(TOBS.getValue('GP_NUMERO'))+' AND '+
         'GAC_INDICEG='+IntToStr(TOBS.GetValue('GP_INDICEG'));
  LESTOBACOMPTES.LoadDetailDBFromSQL ('ACOMPTES',Sql,True,false);
  TOBAcomptes_O.Dupliquer (LESTOBACOmptes,true,true);
end;

procedure TOF_BTSAISIEAVANC.InitChampsSupEntete(TOBS: TOB);
begin
  TobS.AddChampSupValeur ('AVANCSAISIE',0);
  TOBS.AddChampSupValeur ('AVANCPREC',0)
end;

procedure TOF_BTSAISIEAVANC.DefiniAvancementPrecPiece(TOBL: TOB);
var TOBP : TOB;
    cledoc : R_CLEDOC;
begin
  cledoc := TOB2Cledoc(TOBL);
  TOBP := TOBSELECT.findFirst(['GP_NATUREPIECEG','GP_SOUCHE','GP_NUMERO','GP_INDICEG'],
                              [Cledoc.NaturePiece,cledoc.Souche,cledoc.NumeroPiece,cledoc.Indice],true);
  if TOBP <> nil then
  begin
    TOBP.putValue('AVANCPREC',TOBP.GetValue('AVANCPREC')+TOBL.GetValue('MTCUMULEFACTINIT'));
  end;
end;

procedure TOF_BTSAISIEAVANC.SetAvancementPiece(TOBL: TOB; Sens : string='+');
var TOBP : TOB;
    cledoc : R_CLEDOC;
begin
  cledoc := TOB2Cledoc(TOBL);
  if fCurrentOnglet <> FOngletGlobal then exit;
  TOBP := TOBSELECT.findFirst(['GP_NATUREPIECEG','GP_SOUCHE','GP_NUMERO','GP_INDICEG'],
                              [Cledoc.NaturePiece,cledoc.Souche,cledoc.NumeroPiece,cledoc.Indice],true);
  if TOBP <> nil then
  begin
    if Sens = '+' then
    begin
      TOBP.putValue('AVANCSAISIE',TOBP.GetValue('AVANCSAISIE')+TOBL.GetValue('BLF_MTCUMULEFACT'));
    end else
    begin
      TOBP.putValue('AVANCSAISIE',TOBP.GetValue('AVANCSAISIE')-TOBL.GetValue('BLF_MTCUMULEFACT'));
    end;
  end;
end;


procedure TOF_BTSAISIEAVANC.SetAvancementAcompte;
var
  TOBAcc, TOBPieceLoc, TOBAcptes, TOBA,SavAcomptes: TOB;
  StRegle, stModeTraitement: string;
  Indice: integer;
begin

  if LESTOBACOMPTES.detail.count = 0 then exit; // si pas d'acompte poas la peine de continuer
  SavAcomptes := TOB.create('LES ACOMPTES',nil,-1);
  SavAcomptes.Dupliquer (LESTOBACOMPTES,true,true);

  TOBAcptes := TOB.Create('LACOMPTE', nil, -1);
  TOBPieceLoc := TOB.Create('PIECE', nil, -1);
  TOBPieceLoc.Dupliquer (TOBSELECT.detail[0],false,true);

  for Indice := 0 to LESTOBACOMPTES.detail.count - 1 do
  begin
    TOBA := TOB.Create('ACOMPTES', TOBAcptes, -1);
    TOBA.Dupliquer(LESTOBACOMPTES.detail[Indice], true, true);
  end;

  TobAcc := Tob.Create('Les acomptes', nil, -1);
  StRegle := '';
  Tob.Create('', TobAcc, -1);
  TobAcc.Detail[0].Dupliquer(TobTiers, False, TRUE, TRUE);
  Tob.Create('', TobAcc.Detail[0], -1);
  TobAcc.Detail[0].Detail[0].Dupliquer(TobPieceLoc, False, TRUE, TRUE);
  if TOBSelect.detail.count > 1 then
  begin
    TobAcc.Detail[0].Detail[0].data := TOBSelect;
  end;
  TheTob := TobAcc;
  TOBAcptes.ChangeParent(TobAcc.Detail[0].Detail[0], -1);
  //
  // ---
  if TOBSELECT.GetValue('GP_AFFAIREDEVIS') = '' then stModeTraitement := ';HDEVIS'
  else if TOBSELECT.GetValue('AFF_GENERAUTO') = 'AVA' then stModeTraitement := ';AVANCEMENT'
  else if TOBSELECT.GetValue('AFF_GENERAUTO') = 'DAC' then stModeTraitement := ';MEMOIRE'
  else stModeTraitement := ';DIRECTE';
  stModeTraitement := stModetraitement + ';CHOIXSEL;TRANSFOPIECE;';
  AGLLanceFiche('BTP', 'BTACOMPTES', '', '', ActionToString(Action) + StModeTraitement);
  LESTOBACOMPTES.ClearDetail;
  if TOBAcptes.detail.count = 0 then // cas de l'abandon d'application d'acompte
  begin
    for Indice := 0 to SavAcomptes.detail.count - 1 do
    begin
      TOBA := TOB.Create('ACOMPTES', LESTOBACOMPTES, -1);
      TOBA.Dupliquer(SavAcomptes.detail[Indice], true, true);
      TOBA.putValue('GAC_MONTANT', 0);
      TOBA.putValue('GAC_MONTANTDEV', 0);
    end;
  end else
  begin
    for Indice := 0 to TOBAcptes.detail.count - 1 do
    begin
      TOBA := TOB.Create('ACOMPTES', LESTOBACOMPTES, -1);
      TOBA.Dupliquer(TOBAcptes.detail[Indice], true, true);
    end;
  end;
  TobAcc.Free;
  TOBPieceLoc.free;
end;

procedure TOF_BTSAISIEAVANC.CalculeLaLigneFromMontant(Montant : Double; TOBL: TOB;CurrOnglet : Tonglet );
var Pourcent : double;
    Coef : double;
begin
  // Pour les montants
  if TOBL.GetValue('BLF_MTSITUATION') <> 0 then Coef := Montant / TOBL.GetValue('BLF_MTSITUATION') else coef := 1;
  TOBL.putValue('BLF_MTSITUATION',montant);

  TOBL.PutValue('BLF_MTCUMULEFACT',TOBL.GetValue('BLF_MTDEJAFACT')+TOBL.GetValue('BLF_MTSITUATION'));
  if Avancement then
  begin
    if TOBL.GetValue('BLF_MTMARCHE') <> 0 then
    begin
      Pourcent := Arrondi(TOBL.GetValue('BLF_MTCUMULEFACT')/TOBL.GetValue('BLF_MTMARCHE')*100,2);
      TOBL.SetDouble('BLF_POURCENTAVANC',Pourcent);
    //
      TOBL.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*(TOBL.GetValue('BLF_MTCUMULEFACT')/TOBL.GetValue('BLF_MTMARCHE')),V_PGI.okdecQ));
      TOBL.PutValue('BLF_QTESITUATION',TOBL.GetValue('BLF_QTECUMULEFACT')-TOBL.GetValue('BLF_QTEDEJAFACT'));
    end;
  end else
  begin
    if TOBL.GetValue('BLF_MTMARCHE') <> 0 then
    begin
      Pourcent := Arrondi(TOBL.GetValue('BLF_MTSITUATION')/TOBL.GetValue('BLF_MTMARCHE')*100,2);
      TOBL.SetDouble('BLF_POURCENTAVANC',Pourcent);
      // Pour les Quantités
      TOBL.PutValue('BLF_QTESITUATION',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*Pourcent/100,V_PGI.okdecQ));
      TOBL.PutValue('BLF_QTECUMULEFACT',TOBL.GetValue('BLF_QTEDEJAFACT')+TOBL.GetValue('BLF_QTESITUATION'));
  end;
  end;
  //
  if (TOBL.GetValue('BLF_MTMARCHE')=0) then
  begin
    TOBL.SetDouble('GL_PUHTDEV',Arrondi(Montant/TOBL.GetDouble('GL_QTEFACT'),V_PGI.okdecP));
    TOBL.SetDouble('GL_COEFMARG',0);
  end;
  TOBL.PutValue('BLF_MTPRODUCTION',TOBL.GetValue('BLF_MTSITUATION'));
  //
  if (isSousDetail (TOBL)) {and (CurrOnglet = FOngletGlobal)} and (TOBL.GetValue('UNIQUEBLO')<>0) then
  begin
  	SetAvancSousDetail (TOBL,CurrOnglet);
  end;
  //
  if (IsOuvrage(TOBL))and (not IsSousDetail(TOBL)) and (TOBL.getValue('UNIQUEBLO')=0) then
  begin
    IF TOBL.GetValue('BLF_MTMARCHE')<>0 then
    begin
  	  AppliqueAvancSousDetail (TOBL,CurrOnglet);
    end else
    begin
      AppliqueMtSousDetail (Coef,TOBL,CurrOnglet);
    end;
  end;
end;

procedure TOF_BTSAISIEAVANC.SetMtOuvrage (TOBL : TOB;Coef :double);
var Indice,IndiceNomen : Integer;
		TOBOUV,TOBO : TOB;
    MtCumule,MaxMontant,Ecart,Mtligne,MtMarcheL,MtMarcheC : double;
    imaxMontant : Integer;
begin
  if fCurrentOnglet <> FOngletGlobal then Exit;
  //
	IndiceNOmen := TOBL.GetValue('GL_INDICENOMEN')-1;
  MtCumule := 0;  MaxMontant := 0; imaxMontant := -1; MtMarcheL := 0;
  //
  if IndiceNOmen < 0 then Exit;
  //
  //FV1 : 28/07/2016
  if TobOuvrages.detail.count = 0 then exit;
  if (Indicenomen-1) > (TobOuvrages.detail.count - 1) then exit;
  //
  TOBOUV := TOBOuvrages.detail[IndiceNOmen];
  for Indice := 0 to TOBOUV.detail.count -1 do
  begin
    TOBO := TOBOUV.detail[Indice];
    TOBO.putValue('BLF_MTSITUATION',ARRONDI(TOBO.GetValue('BLF_MTSITUATION')*coef,V_PGI.OkdecV));
    TOBO.putValue('BLO_PUHTDEV',ARRONDI(TOBO.GetValue('BLF_MTSITUATION')/TOBO.getDouble('BLO_QTEFACT'),V_PGI.OkdecP));
    if TOBO.GetDouble ('BLF_MTSITUATION') > MaxMontant then
    begin
      imaxMontant := Indice;
      MaxMontant := TOBO.GetDouble ('BLF_MTSITUATION');
    end;
  end;
  // --
  Ecart := ARRONDI(TOBL.GetDouble('BLF_MTSITUATION'),DEV.Decimale) - ARRONDI(MtCumule,DEV.Decimale );
  if Ecart <> 0 then
  begin
    if IMaxMontant = -1 then imaxMontant := 0;
    TOBO := TOBOUV.detail[ImaxMontant];
    TOBO.PutValue('BLF_MTSITUATION',TOBO.GetDouble ('BLF_MTSITUATION')+Ecart);
    TOBO.putValue('BLO_PUHTDEV',ARRONDI(TOBO.GetValue('BLF_MTSITUATION')/TOBO.getDouble('BLO_QTEFACT'),V_PGI.OkdecP));
  end;
  // --
end;


procedure TOF_BTSAISIEAVANC.SetAvancOuvrage (TOBL : TOB);
var Indice,IndiceNomen : Integer;
		TOBOUV,TOBO : TOB;
    Avanc,MtCumule,MaxMontant,Ecart,Mtligne,MtMarcheL,MtMarcheC : double;
    imaxMontant : Integer;
begin
  if fCurrentOnglet <> FOngletGlobal then Exit;
  //
	IndiceNOmen := TOBL.GetValue('GL_INDICENOMEN')-1;
  Avanc := TOBL.getValue('BLF_POURCENTAVANC');
  MtCumule    := 0;
  MaxMontant  := 0;
  imaxMontant := -1;
  MtMarcheL   := 0;
  //
  //FV1 - 28/07/2016
  if TobOuvrages.detail.count = 0 then exit;
  if (IndiceNomen < 0) Or (IndiceNomen > TobOuvrages.detail.count) then Exit;

  TOBOUV := TOBOuvrages.detail[IndiceNOmen];
  for Indice := 0 to TOBOUV.detail.count -1 do
  begin
    TOBO := TOBOUV.detail[Indice];
    if Avancement then
    begin
      if ReajusteDossier then
      begin
      TOBO.putValue('BLF_POURCENTAVANC',Avanc);
        TOBO.PutValue('BLF_MTCUMULEFACT',Arrondi(TOBO.GetValue('BLF_MTMARCHE')*TOBO.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
        TOBO.PutValue('BLF_MTDEJAFACT',TOBO.GetValue('BLF_MTCUMULEFACT'));
        TOBO.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBO.GetValue('BLF_QTEMARCHE')*TOBO.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
        TOBO.PutValue('BLF_QTEDEJAFACT',TOBO.GetValue('BLF_QTECUMULEFACT'));
        MtCumule := MtCumule + TOBO.GetDouble ('BLF_MTDEJAFACT');
        MtMarcheL := MtMarcheL + TOBO.GetValue('BLF_MTMARCHE');
        if TOBO.GetDouble ('BLF_MTDEJAFACT') > MaxMontant then
        begin
          imaxMontant := Indice;
          MaxMontant := TOBO.GetDouble ('BLF_MTDEJAFACT');
        end;
      end else
      begin
        TOBO.putValue('BLF_POURCENTAVANC',Avanc);
      // Pour les montants
      TOBO.PutValue('BLF_MTCUMULEFACT',Arrondi(TOBO.GetValue('BLF_MTMARCHE')*TOBO.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
      TOBO.PutValue('BLF_MTSITUATION',TOBO.GetValue('BLF_MTCUMULEFACT')-TOBO.GetValue('BLF_MTDEJAFACT'));
      // Pour les Quantités
      TOBO.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBO.GetValue('BLF_QTEMARCHE')*TOBO.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
      TOBO.PutValue('BLF_QTESITUATION',TOBO.GetValue('BLF_QTECUMULEFACT')-TOBO.GetValue('BLF_QTEDEJAFACT'));
      MtCumule := MtCumule + TOBO.GetDouble ('BLF_MTSITUATION');
      if TOBO.GetDouble ('BLF_MTSITUATION') > MaxMontant then
      begin
      	imaxMontant := Indice;
        MaxMontant := TOBO.GetDouble ('BLF_MTSITUATION');
      end;
      end;
    end else
    begin
      if ReajusteDossier then
      begin
      TOBO.putValue('BLF_POURCENTAVANC',Avanc);
        TOBO.PutValue('BLF_MTDEJAFACT',Arrondi(TOBO.GetValue('BLF_MTMARCHE')*TOBO.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
        TOBO.PutValue('BLF_QTEDEJAFACT',Arrondi(TOBO.GetValue('BLF_QTEMARCHE')*TOBO.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
        TOBO.PutValue('BLF_MTCUMULEFACT',0);
        TOBO.PutValue('BLF_QTECUMULEFACT',0);
        MtCumule := MtCumule + TOBO.GetDouble ('BLF_MTDEJAFACT');
        MtMarcheL := MtMarcheL + TOBO.GetValue('BLF_MTMARCHE');
        if TOBO.GetDouble ('BLF_MTDEJAFACT') > MaxMontant then
        begin
          imaxMontant := Indice;
          MaxMontant := TOBO.GetDouble ('BLF_MTDEJAFACT');
        end;
      end else
      begin
        TOBO.putValue('BLF_POURCENTAVANC',Avanc);
      // Pour les montants
      TOBO.PutValue('BLF_MTCUMULEFACT',0);
      TOBO.PutValue('BLF_MTSITUATION',Arrondi(TOBO.GetValue('BLF_MTMARCHE')*TOBO.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
      // Pour les Quantités
      TOBO.PutValue('BLF_QTECUMULEFACT',0);
      TOBO.PutValue('BLF_QTESITUATION',Arrondi(TOBO.GetValue('BLF_QTEMARCHE')*TOBO.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
      MtCumule := MtCumule + TOBO.GetDouble ('BLF_MTSITUATION');
      if TOBO.GetDouble ('BLF_MTSITUATION') > MaxMontant then
      begin
      	imaxMontant := Indice;
        MaxMontant := TOBO.GetDouble ('BLF_MTSITUATION');
      end;
    end;
  end;
  end;
  // --
  if ReajusteDossier then
  begin
    Mtligne := ARRONDI(TOBL.GetDouble('BLF_MTDEJAFACT'),DEV.Decimale);
    Ecart :=  Mtligne - ARRONDI(MtCumule,DEV.Decimale );
    if Ecart <> 0 then
    begin
      if IMaxMontant = -1 then imaxMontant := 0;
      //FV1 : 28/07/2016
      if TOBOuv.detail.count > 0 then
      begin
      TOBO := TOBOUV.detail[ImaxMontant];
      if Avancement then
      begin
          TOBO.PutValue('BLF_MTCUMULEFACT', TOBO.GetDouble ('BLF_MTCUMULEFACT')+Ecart);
          TOBO.PutValue('BLF_MTDEJAFACT',   TOBO.GetDouble ('BLF_MTDEJAFACT')+Ecart);
      end else
      begin
          TOBO.PutValue('BLF_MTDEJAFACT',   TOBO.GetDouble ('BLF_MTDEJAFACT')+Ecart);
      end;
    end;
    end;

    MtMarcheC := ARRONDI(TOBL.GetDouble('BLF_MTMARCHE'),DEV.Decimale);
    Ecart :=  MtMarcheC - ARRONDI(MtMarcheL,DEV.Decimale );

    if Ecart <> 0 then
    begin
      if IMaxMontant = -1 then imaxMontant := 0;
      //FV1 : 28/07/2016
      if TOBOuv.detail.count > 0 then
      begin
      TOBO := TOBOUV.detail[ImaxMontant];
        TOBO.PutValue('BLF_MTMARCHE', TOBO.GetDouble ('BLF_MTMARCHE')+Ecart);
    end;
    end;


  end else
  begin
  Ecart := ARRONDI(TOBL.GetDouble('BLF_MTSITUATION'),DEV.Decimale) - ARRONDI(MtCumule,DEV.Decimale );
  if Ecart <> 0 then
  begin
    if IMaxMontant = -1 then imaxMontant := 0;
      //FV1 : 28/07/2016
      if TOBOuv.detail.count > 0 then
      begin
    TOBO := TOBOUV.detail[ImaxMontant];
    if Avancement then
    begin
          TOBO.PutValue('BLF_MTCUMULEFACT', TOBO.GetDouble ('BLF_MTCUMULEFACT')+Ecart);
          TOBO.PutValue('BLF_MTSITUATION',  TOBO.GetDouble ('BLF_MTSITUATION')+Ecart);
    end else
    begin
          TOBO.PutValue('BLF_MTSITUATION',  TOBO.GetDouble ('BLF_MTSITUATION')+Ecart);
    end;
  end;
  end;
  end;
  // --
end;

procedure TOF_BTSAISIEAVANC.AppliqueQteFacOuvSousDetail(TOBO : TOB; CurrOnglet : Tonglet);
var ListeDet : TlisteTOB;
		avanc : double;
    TOBL : TOB;
    Indice : integer;
begin
	ListeDet := FindlignesOuvSaisie (CurrOnglet,TOBO);
  avanc := TOBO.getValue('BLF_POURCENTAVANC');
  if ListeDet.count > 0 then   // lignes de détail trouvé
  begin
  	for Indice := 0 to ListeDet.count -1 do
    begin
    	TOBL := ListeDet.Items [Indice];
      if Avancement then
      begin
      	TOBL.putValue('BLF_POURCENTAVANC',Avanc);
        // Pour les montants
        TOBL.PutValue('BLF_MTCUMULEFACT',Arrondi(TOBL.GetValue('BLF_MTMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
        TOBL.PutValue('BLF_MTDEJAFACT',TOBL.GetValue('BLF_MTCUMULEFACT'));
        // Pour les Quantités
        TOBL.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
        TOBL.PutValue('BLF_QTEDEJAFACT',TOBL.GetValue('BLF_QTECUMULEFACT'));
      end else
      begin
      	TOBL.putValue('BLF_POURCENTAVANC',Avanc);
        // Pour les montants
        TOBL.PutValue('BLF_MTCUMULEFACT',0);
        TOBL.PutValue('BLF_MTDEJAFACT',Arrondi(TOBL.GetValue('BLF_MTMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
        // Pour les Quantités
        TOBL.PutValue('BLF_QTECUMULEFACT',0);
        TOBL.PutValue('BLF_QTEDEJAFACT',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
      end;
      //
      if (isSousDetail (TOBL)) and (CurrOnglet = FOngletGlobal) and (TOBL.GetValue('UNIQUEBLO')<>0) then
      begin
      	SetAvancSousDetail (TOBL,CurrOnglet);
        AppliqueAvancFournisseur (TOBL,avanc);
      end;
      //
    	AfficheLaLigne (CurrOnglet.GS,CurrOnglet.TOBlignes,TOBL.getvalue('INDICE'));
    end;
  end else
  begin
  	SetAvancOuvrage (TOBO);
  end;
  ListeDet.free;

end;

procedure TOF_BTSAISIEAVANC.AppliqueQteOuvSousDetail (TOBO : TOB; CurrOnglet : Tonglet);
var ListeDet : TlisteTOB;
		avanc : double;
    TOBL : TOB;
    Indice : integer;
begin
	ListeDet := FindlignesOuvSaisie (CurrOnglet,TOBO);
  avanc := TOBO.getValue('BLF_POURCENTAVANC');
  if ListeDet.count > 0 then   // lignes de détail trouvé
  begin
  	for Indice := 0 to ListeDet.count -1 do
    begin
    	TOBL := ListeDet.Items [Indice];
      if Avancement then
      begin
      	TOBL.putValue('BLF_POURCENTAVANC',Avanc);
        // Pour les montants
        TOBL.PutValue('BLF_MTCUMULEFACT',Arrondi(TOBL.GetValue('BLF_MTMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
        TOBL.PutValue('BLF_MTSITUATION',TOBL.GetValue('BLF_MTCUMULEFACT')-TOBL.GetValue('BLF_MTDEJAFACT'));
        // Pour les Quantités
        TOBL.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
        TOBL.PutValue('BLF_QTESITUATION',TOBL.GetValue('BLF_QTECUMULEFACT')-TOBL.GetValue('BLF_QTEDEJAFACT'));
      end else
      begin
      	TOBL.putValue('BLF_POURCENTAVANC',Avanc);
        // Pour les montants
        TOBL.PutValue('BLF_MTCUMULEFACT',0);
        TOBL.PutValue('BLF_MTSITUATION',Arrondi(TOBL.GetValue('BLF_MTMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
        // Pour les Quantités
        TOBL.PutValue('BLF_QTECUMULEFACT',0);
        TOBL.PutValue('BLF_QTESITUATION',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
      end;
      //
      if (isSousDetail (TOBL)) and (CurrOnglet = FOngletGlobal) and (TOBL.GetValue('UNIQUEBLO')<>0) then
      begin
      	SetAvancSousDetail (TOBL,CurrOnglet);
        AppliqueAvancFournisseur (TOBL,avanc);
      end;
      //
    	AfficheLaLigne (CurrOnglet.GS,CurrOnglet.TOBlignes,TOBL.getvalue('INDICE'));
    end;
  end else
  begin
  	SetAvancOuvrage (TOBO);
  end;
  ListeDet.free;

end;

procedure TOF_BTSAISIEAVANC.AppliqueMtSousDetail (Coef : Double; TOBO : TOB; CurrOnglet : Tonglet);
var ListeDet : TlisteTOB;
		max,summ,diff : double;
    TOBL : TOB;
    Indice,II,LastRow,TheROW : integer;
begin
  if coef = 1 then Exit;
  Max := 0;
  II := -1;
	ListeDet := FindlignesOuvSaisie (CurrOnglet,TOBO);
  if (ListeDet.count > 0) and (CurrOnglet <> fOngletGlobal) then   // lignes de détail trouvé
  begin
  	for Indice := 0 to ListeDet.count -1 do
    begin
    	TOBL := ListeDet.Items [Indice];
      if TOBL.GetValue('BLF_MTSITUATION')>max then
      begin
        Max := TOBL.GetValue('BLF_MTSITUATION');
        II := Indice;
      end;
      TOBL.PutValue('BLF_MTSITUATION',ARRONDI(TOBL.GetValue('BLF_MTSITUATION')*Coef,V_PGI.okdecV));
      TOBL.SetDouble('GL_PUHTDEV',ARRONDI(TOBL.GetValue('BLF_MTSITUATION')/TOBL.Getvalue('GL_QTEFACT'),V_PGI.okdecP));
      summ := summ + TOBL.GetValue('BLF_MTSITUATION');
      //
    	AfficheLaLigne (CurrOnglet.GS,CurrOnglet.TOBlignes,TOBL.getvalue('INDICE'));
    end;
    if summ <>  TOBO.GetValue('BLF_MTSITUATION') then
    begin
      diff:= TOBO.GetValue('BLF_MTSITUATION') - summ;
      TOBL := ListeDet.items[II];
      TOBL.PutValue('BLF_MTSITUATION',ARRONDI(TOBL.GetValue('BLF_MTSITUATION')+diff,V_PGI.okdecV));
      TOBL.SetDouble('GL_PUHTDEV',ARRONDI(TOBL.GetValue('BLF_MTSITUATION')/TOBL.Getvalue('GL_QTEFACT'),V_PGI.okdecP));
    	AfficheLaLigne (CurrOnglet.GS,CurrOnglet.TOBlignes,TOBL.getvalue('INDICE'));
    end;
  end else
  begin
    lastRow := GS.row;
    TheRow := TOBO.GetIndex+1;
  	SetMtOuvrage (TOBO,Coef);
  	SetLigne ( GS,TheRow,LIBELLE);
    if (ListeDet.count > 0) then
    begin
      OuvrirClick(self);
      OuvrirClick(self);
    end;
  	SetLigne ( GS,lastRow,LIBELLE);
  end;
  ListeDet.free;

end;

procedure TOF_BTSAISIEAVANC.AppliqueAvancSousDetail (TOBO : TOB; CurrOnglet : Tonglet;WithDeduction:Boolean=True;WithAjout:Boolean=True;WithCalcul:Boolean=true);
var ListeDet : TlisteTOB;
		avanc : double;
    TOBL : TOB;
    Indice,LastRow,TheRow : integer;
begin
	ListeDet := FindlignesOuvSaisie (CurrOnglet,TOBO);
  avanc := TOBO.getValue('BLF_POURCENTAVANC');
  if (ListeDet.count > 0) and (CurrOnglet <> fOngletGlobal) then   // lignes de détail trouvé
  begin
  	for Indice := 0 to ListeDet.count -1 do
    begin
    	TOBL := ListeDet.Items [Indice];
      if Avancement then
      begin
        if ReajusteDossier then
        begin
      	TOBL.putValue('BLF_POURCENTAVANC',Avanc);
          TOBL.PutValue('BLF_MTCUMULEFACT',Arrondi(TOBL.GetValue('BLF_MTMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
          TOBL.PutValue('BLF_MTDEJAFACT',TOBL.GetValue('BLF_MTCUMULEFACT'));
          TOBL.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
          TOBL.PutValue('BLF_QTEDEJAFACT',TOBL.GetValue('BLF_QTECUMULEFACT'));
        end else
        begin
          TOBL.putValue('BLF_POURCENTAVANC',Avanc);
        // Pour les montants
        TOBL.PutValue('BLF_MTCUMULEFACT',Arrondi(TOBL.GetValue('BLF_MTMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
        TOBL.PutValue('BLF_MTSITUATION',TOBL.GetValue('BLF_MTCUMULEFACT')-TOBL.GetValue('BLF_MTDEJAFACT'));
        // Pour les Quantités
        TOBL.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
        TOBL.PutValue('BLF_QTESITUATION',TOBL.GetValue('BLF_QTECUMULEFACT')-TOBL.GetValue('BLF_QTEDEJAFACT'));
        end;
      end else
      begin
        if ReajusteDossier then
        begin
      	TOBL.putValue('BLF_POURCENTAVANC',Avanc);
          TOBL.PutValue('BLF_MTDEJAFACT',Arrondi(TOBL.GetValue('BLF_MTMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
          TOBL.PutValue('BLF_QTEDEJAFACT',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
          TOBL.PutValue('BLF_MTCUMULEFACT',0);
          TOBL.PutValue('BLF_QTECUMULEFACT',0);
        end else
        begin
          TOBL.putValue('BLF_POURCENTAVANC',Avanc);
        // Pour les montants
        TOBL.PutValue('BLF_MTCUMULEFACT',0);
        TOBL.PutValue('BLF_MTSITUATION',Arrondi(TOBL.GetValue('BLF_MTMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.decimale));
        // Pour les Quantités
        TOBL.PutValue('BLF_QTECUMULEFACT',0);
        TOBL.PutValue('BLF_QTESITUATION',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
      end;
      end;
      //
      if (isSousDetail (TOBL)) and (CurrOnglet = FOngletGlobal) and (TOBL.GetValue('UNIQUEBLO')<>0) then
      begin
      	SetAvancSousDetail (TOBL,CurrOnglet);
        AppliqueAvancFournisseur (TOBL,avanc,WithDeduction,WithAjout,WithCalcul);
      end;
      //
    	AfficheLaLigne (CurrOnglet.GS,CurrOnglet.TOBlignes,TOBL.getvalue('INDICE'));
    end;
  end else
  begin
    lastRow := GS.row;
    TheRow := TOBO.GetIndex+1;
  	SetAvancOuvrage (TOBO);
  	SetLigne ( GS,TheRow,LIBELLE);
    if (ListeDet.count > 0) then
    begin
      OuvrirClick(self);
      OuvrirClick(self);
    end;
  	SetLigne ( GS,lastRow,LIBELLE);
  end;
  ListeDet.free;

end;

procedure TOF_BTSAISIEAVANC.CalculeLaLigneFromQteFact(TOBL: TOB;CurrOnglet : Tonglet );
var pourcent : double;
		PPQ : double;
begin
  PPQ := TOBL.GEtValue('GL_PRIXPOURQTE'); if PPQ = 0 then PPQ := 1;
	if Avancement then
  begin
    // Pour les Quantités
    TOBL.PutValue('BLF_QTECUMULEFACT',TOBL.GetValue('BLF_QTEDEJAFACT'));
    if TOBL.GetValue('BLF_QTEMARCHE') <> 0 then
    begin
      Pourcent := Arrondi(TOBL.GetValue('BLF_QTEDEJAFACT')/TOBL.GetValue('BLF_QTEMARCHE')*100,2);
    end else
    begin
      Pourcent := 100;
    end;
    // Pour les montants
    TOBL.PutValue('BLF_MTDEJAFACT',Arrondi(TOBL.GetValue('BLF_QTEDEJAFACT')*TOBL.GetValue('GL_PUHTDEV')/PPQ,DEV.decimale));
    TOBL.PutValue('BLF_MTCUMULEFACT',TOBL.GetValue('BLF_MTDEJAFACT'));
    TOBL.SetDouble('BLF_POURCENTAVANC',Pourcent);
  end else
  begin
    // Pour les Quantités
    TOBL.PutValue('BLF_QTECUMULEFACT',0);
    TOBL.PutValue('BLF_MTCUMULEFACT',0);
    if TOBL.GetValue('BLF_QTEMARCHE') <> 0 then
    begin
      Pourcent := Arrondi(TOBL.GetValue('BLF_QTEDEJAFACT')/TOBL.GetValue('BLF_QTEMARCHE')*100,2);
    end else
    begin
      pourcent := 100;
    end;
    // Pour les montants
    TOBL.PutValue('BLF_MTDEJAFACT',Arrondi(TOBL.GetValue('BLF_QTEDEJAFACT')*TOBL.GetValue('GL_PUHTDEV')/PPQ,DEV.decimale));
    TOBL.SetDouble('BLF_POURCENTAVANC',Pourcent);
  end;
  //
  if (isSousDetail (TOBL)) {and (CurrOnglet = FOngletGlobal)} and (TOBL.GetValue('UNIQUEBLO')<>0) then
  begin
  	SetAvancSousDetail (TOBL,CurrOnglet);
  end;
  //
  if (IsOuvrage(TOBL)) and (not IsSousDetail(TOBL)) and (TOBL.getValue('UNIQUEBLO')=0) then
  begin
  	AppliqueQteFacOuvSousDetail (TOBL,CurrOnglet);
  end;

end;

procedure TOF_BTSAISIEAVANC.CalculeLaLigneFromQteSit(TOBL: TOB;CurrOnglet : Tonglet );
var pourcent : double;
		QteCumule,PPQ : double;
begin
  PPQ := TOBL.GEtValue('GL_PRIXPOURQTE'); if PPQ = 0 then PPQ := 1;
	if Avancement then
  begin
    // Pour les montants
    TOBL.PutValue('BLF_MTSITUATION',Arrondi(TOBL.GetValue('BLF_QTESITUATION')*TOBL.GetValue('GL_PUHTDEV')/PPQ,DEV.decimale));
    TOBL.PutValue('BLF_MTCUMULEFACT',TOBL.GetValue('BLF_MTDEJAFACT')+TOBL.GetValue('BLF_MTSITUATION'));
    // Pour les Quantités
    TOBL.PutValue('BLF_QTECUMULEFACT',TOBL.GetValue('BLF_QTEDEJAFACT')+TOBL.GetValue('BLF_QTESITUATION'));
  	QteCumule := TOBL.GetValue('BLF_QTEDEJAFACT')+TOBL.GetValue('BLF_QTESITUATION');
    if TOBL.GetValue('BLF_QTEMARCHE') <> 0 then
    begin
    Pourcent := Arrondi(QteCumule/TOBL.GetValue('BLF_QTEMARCHE')*100,2);
    end else
    begin
      Pourcent := 100;
    end;
    TOBL.SetDouble('BLF_POURCENTAVANC',Pourcent);
  end else
  begin
    // Pour les montants
    TOBL.PutValue('BLF_MTCUMULEFACT',0);
    TOBL.PutValue('BLF_MTSITUATION',Arrondi(TOBL.GetValue('BLF_QTESITUATION')*TOBL.GetValue('GL_PUHTDEV')/PPQ,DEV.decimale));
    // Pour les Quantités
    TOBL.PutValue('BLF_QTECUMULEFACT',0);
    if TOBL.GetValue('BLF_QTEMARCHE') <> 0 then
    begin
    Pourcent := Arrondi(TOBL.GetValue('BLF_QTESITUATION')/TOBL.GetValue('BLF_QTEMARCHE')*100,2);
    end else
    begin
      pourcent := 100;
    end;
    TOBL.SetDouble('BLF_POURCENTAVANC',Pourcent);
  end;
  //
  TOBL.PutValue('BLF_MTPRODUCTION',TOBL.GetValue('BLF_MTSITUATION'));
  //
  if (isSousDetail (TOBL)) {and (CurrOnglet = FOngletGlobal)} and (TOBL.GetValue('UNIQUEBLO')<>0) then
  begin
  	SetAvancSousDetail (TOBL,CurrOnglet);
  end;
  //
  if (IsOuvrage(TOBL)) and (not IsSousDetail(TOBL)) and (TOBL.getValue('UNIQUEBLO')=0) then
  begin
  	AppliqueQteOuvSousDetail (TOBL,CurrOnglet);
  end;
end;

procedure TOF_BTSAISIEAVANC.DefaireParag(Sender: TObject);
var Debut,Fin : integer;
begin
  //
  Debut := RecupDebutParagraph (GS.row);
  Fin := RecupFinParagraph (GS.row);
  //
  DefaireSegment (Debut,Fin);
end;

function TOF_BTSAISIEAVANC.AjouteLigneEcart(MtEcart: double) : boolean;
var TOBL,TOBArt : TOB;
		Numl : integer;
    Artecart : string;
    Qart : TQuery;
//    EnHt : boolean;
    TOBtaxesL : TOB;
begin
	result := true;
	TOBart := TOB.create ('ARTICLE',nil,-1);
  TOBTaxesL := TOB.Create ('LES TAXES LIGNES',nil,-1);
  TRY
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
    TOBL := TOB.Create ('LIGNE',TOBLIGNES, -1); // ajout d'une ligne
    numl := TOBLignes.detail.count;
    AddLesSupLigne(TOBL,false);
    InitLesSupLigne(TOBL);
    AddChampsSupAvanc (TOBL);
    AddChampsSupMtINITAvanc (TOBL);
    PositionneLigne (TOBL,Numl);
    InitialiseTOBLigne(TOBLIGNES, TOBTiers, TOBAffaire, NumL);
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
    ArticleVersLigne(TOBLignes, TOBArt, nil, TOBL, TOBTiers);
    TOBL.PutValue('GL_QTEFACT', 1);
    TOBL.PutValue('GL_QTESTOCK', 1);
    TOBL.PutValue('GL_QTERESTE', 1);
    TOBL.PutValue('GL_MTRESTE',  0);

    TOBL.PutValue('GL_PRIXPOURQTE', 1);
    TOBL.PutValue('GL_PUHTDEV', MtEcart);
    TOBL.putValue('GL_DPA',0);
    TOBL.putValue('GL_COEFFG',0);
    TOBL.putValue('GL_COEFFC',0);
    TOBL.putValue('GL_DPR',0);
    TOBL.putValue('GL_COEFMARG',0);
    TOBL.PutValue('GL_RECALCULER', 'X');
    CalculeLigneHT(TOBL,nil,TOBLignes,DEV, DEV.Decimale,False,TOBTiers) ;
    SetInfoFactEcart (TOBL);
  FINALLY
  	TOBARt.free;
    TOBtaxesL.free;
  END;
end;

procedure TOF_BTSAISIEAVANC.ClotureFacturation;
var TOBPiece,TOBFAC : TOB;
		I,Indice,NbMois : integer;
    CodeAffaire,TypeGeneration,Sql : string;
    DateCloture,DateFinGarantie : TDateTime;
    QQ : TQuery;
    result : boolean;
begin
	TOBFAC := TOB.Create ('LES FACTURES DU DOSSIER',nil,-1);
	TypeGeneration:= TOBSelect.getValue('AFF_GENERAUTO');
	DateCloture:= TOBSelect.getValue('DATECLOTURE');
//
  NbMois := ValeurI(GetParamSocSecur('SO_AFALIMGARANTI', 0));
  if NbMois <> 0 then DateFinGarantie := PlusDate(DateCloture,NbMois,'M') else DateFinGarantie := PlusDate(DateCloture,12,'M');
//
  for I := 0 to TOBSelect.detail.count -1 do
  begin
    TobPiece := TOBSelect.detail[i];
    CodeAffaire := TOBPiece.GetValue('GP_AFFAIREDEVIS');
    sql := 'UPDATE AFFAIRE SET AFF_ETATAFFAIRE="TER",AFF_DATEFIN="'+usdatetime(DateCloture)+
    			 '",AFF_DATEGARANTIE="'+UsDateTime(DateFinGarantie)+'" WHERE AFF_AFFAIRE="'+CodeAffaire+'"';
    result := (ExecuteSql (Sql) > 0 );
    if (result) then
    begin
      sql := 'SELECT GP_NATUREPIECEG,GP_DATEPIECE,GP_SOUCHE,GP_NUMERO,GP_INDICEG FROM PIECE WHERE GP_NATUREPIECEG="FBT" '+
      'AND GP_AFFAIREDEVIS="'+CodeAffaire+'"';
      QQ := OpenSql (sql,true,-1,'',true);
      if not QQ.eof then
      begin
        TOBFAC.ClearDetail;
        TOBFAC.LoadDetailDB ('PIECE','','',QQ,false);
        //
        for Indice := 0 to TOBFAC.detail.count -1 do
        begin
          Sql := 'UPDATE PIECERG SET PRG_DATEECH="'+UsDateTime(DateFinGarantie)+'"'+
                  'WHERE PRG_NATUREPIECEG="'+TOBFAC.detail[Indice].GetValue('GP_NATUREPIECEG')+'" AND '+
//                  'PRG_DATEPIECE="'+UsDateTime(TOBFAC.detail[Indice].GetValue('GP_DATEPIECE'))+'" AND '+
                  'PRG_SOUCHE="'+TOBFAC.detail[Indice].GetValue('GP_SOUCHE')+'" AND '+
                  'PRG_NUMERO='+inttostr(TOBFAC.detail[Indice].GetValue('GP_NUMERO'))+' AND '+
                  'PRG_INDICEG='+inttostr(TOBFAC.detail[Indice].GetValue('GP_INDICEG'));
          ExecuteSql (Sql);
          //
          Sql := 'UPDATE PIECE SET GP_VIVANTE="-"'+
                  'WHERE GP_NATUREPIECEG="'+TOBFAC.detail[Indice].GetValue('GP_NATUREPIECEG')+'" AND '+
                  'GP_SOUCHE="'+TOBFAC.detail[Indice].GetValue('GP_SOUCHE')+'" AND '+
                  'GP_NUMERO='+inttostr(TOBFAC.detail[Indice].GetValue('GP_NUMERO'))+' AND '+
                  'GP_INDICEG='+inttostr(TOBFAC.detail[Indice].GetValue('GP_INDICEG'));
          ExecuteSql (Sql);
          //
          Sql := 'UPDATE LIGNE SET GL_VIVANTE="-"'+
                  'WHERE GL_NATUREPIECEG="'+TOBFAC.detail[Indice].GetValue('GP_NATUREPIECEG')+'" AND '+
                  'GL_SOUCHE="'+TOBFAC.detail[Indice].GetValue('GP_SOUCHE')+'" AND '+
                  'GL_NUMERO='+inttostr(TOBFAC.detail[Indice].GetValue('GP_NUMERO'))+' AND '+
                  'GL_INDICEG='+inttostr(TOBFAC.detail[Indice].GetValue('GP_INDICEG'));
          ExecuteSql (Sql);
        end;
        //
      end;
      ferme (QQ);
    end;
    if result then
    begin
    	Sql := 'UPDATE PIECE SET GP_VIVANTE="-" WHERE GP_NATUREPIECEG="'+TOBPiece.getValue('GP_NATUREPIECEG')+'" AND '+
      			 'GP_SOUCHE="'+TOBPiece.getValue('GP_SOUCHE')+'" AND GP_NUMERO='+IntToStr(TOBPiece.getValue('GP_NUMERO'))+' AND '+
             'GP_INDICEG='+IntToStr(TOBPiece.getValue('GP_INDICEG'));
      result := (ExecuteSql(Sql)>0);
    end;
    if result then
    begin
    	Sql := 'UPDATE LIGNE SET GL_VIVANTE="-" WHERE GL_NATUREPIECEG="'+TOBPiece.getValue('GP_NATUREPIECEG')+'" AND '+
      			 'GL_SOUCHE="'+TOBPiece.getValue('GP_SOUCHE')+'" AND GL_NUMERO='+IntToStr(TOBPiece.getValue('GP_NUMERO'))+' AND '+
             'GL_INDICEG='+IntToStr(TOBPiece.getValue('GP_INDICEG'));
      ExecuteSql(Sql);
    end;
  end;
  TOBFac.free;
end;

procedure TOF_BTSAISIEAVANC.GetPorts(TOBpiece: TOB);
var SQL,refPiece : string;
		QQ : TQuery;
    cledoc : r_cledoc;
begin
  RefPiece := EncodeRefPiece (TOBpiece,0,false);
  DecodeRefPiece (Refpiece,cledoc);
  SQl := 'SELECT * FROM PIEDPORT WHERE '+WherePiece(Cledoc,ttdPorc ,true);
  QQ := OpenSql (SQl,true,-1,'',true);
  if not QQ.eof then TOBPorcs.LoadDetailDB('PIEDPORT','','',QQ,false);
  ferme (QQ);
end;

function TOF_BTSAISIEAVANC.DemandeDatesFacturation(var DateFac: TDateTime ) : boolean;
var TobDates : TOB;
begin
  TOBDates := TOB.Create ('LES DATES', nil,-1);
  TOBDates.AddChampSupValeur('RETOUROK','-');
  TOBDates.AddChampSupValeur('DATFAC',iDate1900);
  TOBDates.AddChampSupValeur('DATESITUATION','-');
  TOBDates.AddChampSupValeur('TYPEDATE','Date de mise à jour');
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


{ TlistOnglet }

function TlistOnglet.Add(AObject: TOnglet): Integer;
begin
  result := Inherited Add(AObject);
end;

procedure TlistOnglet.clear;
var indice : integer;
begin
  if count > 0 then
  begin
    for Indice := count -1 downto 0 do
    begin
      if TOnglet(Items [Indice])<> nil then TOnglet(Items [Indice]).free;
    end;
  end;
  inherited;
end;

constructor TlistOnglet.create;
begin
	ffreeAll := True;
end;

destructor TlistOnglet.destroy;
begin
  if ffreeAll then clear;
  inherited;
end;

function TlistOnglet.findOnglet(NomOnglet: string): TOnglet;
var Indice : integer;
begin
  result := nil;
  for Indice := 0 to Count -1 do
  begin
    if Items[Indice].Name = NomOnglet then
    begin
      result:=Items[Indice];
      break;
    end;
  end;
end;

function TlistOnglet.GetItems(Indice: integer): TOnglet;
begin
  result := TOnglet (Inherited Items[Indice]);
end;


procedure TlistOnglet.SetItems(Indice: integer; const Value: TOnglet);
begin
  Inherited Items[Indice]:= Value;
end;

{ TOnglet }

constructor TOnglet.create (XX : TOF; Generale : boolean);
begin
  fGeneral := Generale;
  fSousTraitant := false;
  fTVDevis := TTreeView.create(TOF_BTSAISIEAVANC(XX).GetControl('PTVDEVIS'));
  fTVDevis.align := alclient;
  fTVDEVIS.Parent := THPanel(TOF_BTSAISIEAVANC(XX).GetControl('PTVDEVIS'));
  fTVDEVIS.Visible := false;
  fname := '';
  fGS := nil;
  fstPrev := '';
  fTabSheet := nil;
  fMtSituation := 0;
	ffournisseur := '';
  fBOUVRIR := TToolbarButton97.Create(Application);
end;

destructor TOnglet.destroy;
begin
  if fBOUVRIR<>nil then
  begin
  	fBOUVRIR.free;
  	fBOUVRIR := nil;
  end;
  if not fGeneral then
  begin
  	if fTOBLignes <> nil then fTOBlignes.free;
    if fGS <> nil then fGS.free;
    if fTVDEVIS <> nil then fTVDEVIS.free;
  end;
  inherited destroy;
end;

function TOF_BTSAISIEAVANC.EnregistreDevis : Boolean;

  procedure EcritLigneLoc (TOBL : TOB);
  var Sql : string;
  begin
    Sql := 'UPDATE LIGNE SET '+
           'GL_POURCENTAVANC='+STRFPOINT(TOBL.GetDouble('BLF_POURCENTAVANC'))+','+
           'GL_QTEPREVAVANC='+STRFPOINT(TOBL.GetDouble('GL_QTEPREVAVANC'))+','+
           'GL_QTESIT='+STRFPOINT(TOBL.GetDouble('GL_QTESIT'))+ ' WHERE '+
           'GL_NATUREPIECEG="'+TOBL.Getvalue ('GL_NATUREPIECEG')+'" AND '+
           'GL_SOUCHE="'+TOBL.Getvalue ('GL_SOUCHE')+'" AND '+
           'GL_NUMERO='+inttostr(TOBL.Getvalue ('GL_NUMERO'))+' AND '+
           'GL_INDICEG='+inttostr(TOBL.Getvalue ('GL_INDICEG'))+' AND '+
           'GL_NUMLIGNE='+inttostr(TOBL.Getvalue ('GL_NUMLIGNE'));
    ExecuteSQL(SQL);
  end;

  procedure SetLigneDevis (TOBL,TOBLdevis : TOB; avancement : boolean);
  var TOBD: TOB;
  begin
  	TOBD := TOB.Create ('LIGNEFAC',TOBLDEVIS,-1);
    TOBD.putValue('BLF_NATUREPIECEG',TOBL.getValue('GL_NATUREPIECEG'));
    TOBD.putValue('BLF_SOUCHE',TOBL.getValue('GL_SOUCHE'));
    TOBD.putValue('BLF_DATEPIECE',TOBL.getValue('GL_DATEPIECE'));
    TOBD.putValue('BLF_NUMERO',TOBL.getValue('GL_NUMERO'));
    TOBD.putValue('BLF_INDICEG',TOBL.getValue('GL_INDICEG'));
    TOBD.putValue('BLF_NUMORDRE',TOBL.getValue('GL_NUMORDRE'));
    TOBD.putValue('BLF_MTMARCHE',TOBL.GetValue('BLF_MTMARCHE'));
    TOBD.putValue('BLF_POURCENTAVANC',TOBL.GetValue('BLF_POURCENTAVANC'));
    TOBD.putValue('BLF_MTDEJAFACT',TOBL.GetValue('BLF_MTDEJAFACT'));
    TOBD.putValue('BLF_MTCUMULEFACT',TOBL.GetValue('BLF_MTCUMULEFACT'));
    TOBD.putValue('BLF_MTSITUATION',0);
    TOBD.putValue('BLF_UNIQUEBLO',0);
    TOBD.putValue('BLF_QTEMARCHE',TOBL.GetValue('BLF_QTEMARCHE'));
    TOBD.putValue('BLF_QTECUMULEFACT',TOBL.GetValue('BLF_QTECUMULEFACT'));
    TOBD.putValue('BLF_QTEDEJAFACT',TOBL.GetValue('BLF_QTEDEJAFACT'));
    TOBD.putValue('BLF_QTESITUATION',0);
    TOBD.putValue('BLF_NATURETRAVAIL',TOBL.GetValue('GLC_NATURETRAVAIL'));
    TOBD.putValue('BLF_FOURNISSEUR',TOBL.GetValue('GL_FOURNISSEUR'));
    //
    TOBL.putvalue('GL_QTEPREVAVANC',TOBL.Getvalue('BLF_QTECUMULEFACT'));
    TOBL.putvalue('GL_QTESIT',TOBL.Getvalue('BLF_QTECUMULEFACT'));
    TOBL.putvalue('GL_POURCENTAVANC',TOBL.Getvalue('BLF_POURCENTAVANC'));
    //
    EcritLigneLoc (TOBL);
    TOBD.SetAllModifie (true);
  end;

  procedure  EcritLigneOuv (TOBOO : TOB);
  var Sql : string;
      Ligne : Integer;
      ii,UniqueBLO : Integer;
  begin
    Ligne :=TOBOO.Getvalue ('BLO_NUMLIGNE');
    UniqueBlo := TOBOO.GetInteger('BLO_UNIQUEBLO');

    Sql :='UPDATE LIGNEOUV SET '+
           'BLO_UNIQUEBLO='+inttostr(TOBOO.GetValue('BLO_UNIQUEBLO'))+' '+
           'WHERE '+
           'BLO_NATUREPIECEG="'+TOBOO.Getvalue ('BLO_NATUREPIECEG')+'" AND '+
           'BLO_SOUCHE="'+TOBOO.Getvalue ('BLO_SOUCHE')+'" AND '+
           'BLO_NUMERO='+inttostr(TOBOO.Getvalue ('BLO_NUMERO'))+' AND '+
           'BLO_INDICEG='+inttostr(TOBOO.Getvalue ('BLO_INDICEG'))+' AND '+
           'BLO_NUMLIGNE='+inttostr(TOBOO.Getvalue ('BLO_NUMLIGNE'))+' AND '+
           'BLO_N1='+inttostr(TOBOO.Getvalue ('BLO_N1'))+' AND '+
           'BLO_N2='+inttostr(TOBOO.Getvalue ('BLO_N2'))+' AND '+
           'BLO_N3='+inttostr(TOBOO.Getvalue ('BLO_N3'))+' AND '+
           'BLO_N4='+inttostr(TOBOO.Getvalue ('BLO_N4'))+' AND '+
           'BLO_N5='+inttostr(TOBOO.Getvalue ('BLO_N5'));
    ExecuteSQL(Sql);
    if TOBOO.Detail.count > 0 then
    begin
      for II := 0 to TOBOO.detail.count -1 do
      begin
        EcritLigneOuv(TOBOO.detail[II]);
      end;
    end;
  end;

  procedure SetLignesDevisOuvDetail (TOBL,TOBOuvrage,TOBLDEVIS : TOB ; avancement : boolean);
  var indiceNomen,II,UniqueBlo : integer;
      TOBOUV,TOBOO,TOBD : TOB;
  begin
    IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
    if IndiceNomen = 0 then exit;
    //FV1 : 28/07/2016
    if TobOuvrage.detail.count = 0 then exit;
    if (Indicenomen-1) > (TobOuvrage.detail.count - 1) then exit;
    //
    TOBOUV := TOBOuvrage.detail[IndiceNomen-1];
    if TOBOUV = nil then exit;
    for II := 0 to TOBOUV.detail.count -1 do
    begin
      TOBOO := TOBOUV.detail[II];
      //
      UniqueBlo := TOBOO.GetInteger('BLO_UNIQUEBLO');
//      EcritLigneOuv (TOBOO);
      //
      TOBD := TOB.Create ('LIGNEFAC',TOBLDEVIS,-1);
      TOBD.putValue('BLF_NATUREPIECEG',TOBOO.getValue('BLO_NATUREPIECEG'));
      TOBD.putValue('BLF_SOUCHE',TOBOO.getValue('BLO_SOUCHE'));
      TOBD.putValue('BLF_DATEPIECE',TOBOO.getValue('BLO_DATEPIECE'));
      TOBD.putValue('BLF_NUMERO',TOBOO.getValue('BLO_NUMERO'));
      TOBD.putValue('BLF_INDICEG',TOBOO.getValue('BLO_INDICEG'));
      TOBD.putValue('BLF_NUMORDRE',0);
      TOBD.putValue('BLF_UNIQUEBLO',TOBOO.getValue('BLO_UNIQUEBLO'));
      TOBD.putValue('BLF_MTMARCHE',TOBOO.GetValue('BLF_MTMARCHE'));
    	TOBD.putValue('BLF_POURCENTAVANC',TOBOO.GetValue('BLF_POURCENTAVANC'));
      TOBD.putValue('BLF_MTDEJAFACT',TOBOO.GetValue('BLF_MTDEJAFACT'));
      TOBD.putValue('BLF_MTCUMULEFACT',TOBOO.GetValue('BLF_MTCUMULEFACT'));
      TOBD.putValue('BLF_MTSITUATION',0);
      TOBD.putValue('BLF_QTEMARCHE',TOBOO.GetValue('BLF_QTEMARCHE'));
      TOBD.putValue('BLF_QTECUMULEFACT',TOBOO.GetValue('BLF_QTECUMULEFACT'));
      TOBD.putValue('BLF_QTEDEJAFACT',TOBOO.GetValue('BLF_QTEDEJAFACT'));
      TOBD.putValue('BLF_QTESITUATION',0);
      TOBD.putValue('BLF_NATURETRAVAIL',TOBOO.GetValue('BLF_NATURETRAVAIL'));
      TOBD.putValue('BLF_FOURNISSEUR',TOBOO.GetValue('BLF_FOURNISSEUR'));
      TOBD.SetAllModifie (true);
    end;
  end;

var TOBLDevis,TOBL : TOB;
		Indice,II : integer;
    cledoc : r_cledoc;
    XX : TFsplashScreen;
begin
	result := true;
  XX := TFsplashScreen.Create (application);
  XX.Label1.Caption  := TraduireMemoire('Enregistrement du dossier de facturation en cours...');
  XX.Show;
  XX.BringToFront;
  XX.refresh;
  TOBLDEVIS := TOB.create('LES LIGNES FACTURES DEVIS',nil,-1);
  //
  TRY
    for Indice := 0 to TOBLIGNES.detail.count -1 do
    begin
      TOBL := TOBLIGNES.detail[Indice];
      if TOBL.GetValue('GL_TYPELIGNE')<>'ART' then continue;
      if (TOBL.GetValue('GL_NUMORDRE')<>0) then
      begin
        //
        SetLigneDevis (TOBL,TOBLdevis,Avancement);
        if isOuvrage(TOBL) then
        begin
          SetLignesDevisOuvDetail (TOBL,TOBOuvrages,TOBLDEVIS,avancement);
        end;
      end;
    end;
    //
    for II := 0 to TOBSELECT.Detail.count -1 do
    begin
      cledoc := TOB2CleDoc(TOBSELECT.detail[II]);
      ExecuteSql('DELETE FROM LIGNEFAC WHERE '+WherePiece(cledoc,ttdLignefac,false));
      //
      ExecuteSQL('UPDATE PIECE SET GP_UNIQUEBLO='+InttoStr(TOBSELECT.detail[II].GetValue('GP_UNIQUEBLO')));
    end;
    //
    if TOBLDEVIS.detail.count <> 0 then
    begin
    	for indice := 0 to TOBLDevis.detail.count -1 do
      begin
        TOBLDEVIS.detail[Indice].SetAllModifie(true);   
        if not TOBLDEVIS.detail[Indice].InsertDB (nil,false) then
        BEGIN
          result := false;
          break;
        END;
      end;
    end;
    //
  FINALLY
    TOBLDEVIS.Free;
    XX.free;
  end;

end;

function TOF_BTSAISIEAVANC.EnregistreAvancement : boolean;

  procedure SetLigneDevis (TOBL,TOBLdevis : TOB; avancement : boolean);
  var TOBD: TOB;
  begin
  	TOBD := TOB.Create ('LIGNEFAC',TOBLDEVIS,-1);
    TOBD.putValue('BLF_NATUREPIECEG',TOBL.getValue('GL_NATUREPIECEG'));
    TOBD.putValue('BLF_SOUCHE',TOBL.getValue('GL_SOUCHE'));
    TOBD.putValue('BLF_DATEPIECE',TOBL.getValue('GL_DATEPIECE'));
    TOBD.putValue('BLF_NUMERO',TOBL.getValue('GL_NUMERO'));
    TOBD.putValue('BLF_INDICEG',TOBL.getValue('GL_INDICEG'));
    TOBD.putValue('BLF_NUMORDRE',TOBL.getValue('GL_NUMORDRE'));
    TOBD.putValue('BLF_MTMARCHE',TOBL.GetValue('BLF_MTMARCHE'));
    TOBD.putValue('BLF_POURCENTAVANC',TOBL.GetValue('BLF_POURCENTAVANC'));
    if Avancement then
    begin
      TOBD.putValue('BLF_MTDEJAFACT',TOBL.GetValue('BLF_MTCUMULEFACT')-TOBL.GetValue('BLF_MTSITUATION'));
      TOBD.putValue('BLF_MTCUMULEFACT',TOBL.GetValue('BLF_MTCUMULEFACT')-TOBL.GetValue('BLF_MTSITUATION'));
    end else
    begin
      TOBD.putValue('BLF_MTDEJAFACT',TOBL.GetValue('BLF_MTDEJAFACT'));
    end;
    TOBD.putValue('BLF_MTSITUATION',TOBL.GetValue('BLF_MTSITUATION'));
    TOBD.putValue('BLF_QTEMARCHE',TOBL.GetValue('BLF_QTEMARCHE'));
    if Avancement then
    begin
      TOBD.putValue('BLF_QTECUMULEFACT',TOBL.GetValue('BLF_QTECUMULEFACT')-TOBL.GetValue('BLF_QTESITUATION'));
      TOBD.putValue('BLF_QTEDEJAFACT',TOBL.GetValue('BLF_QTECUMULEFACT')-TOBL.GetValue('BLF_QTESITUATION'));
    end else
    begin
      TOBD.putValue('BLF_QTEDEJAFACT',TOBL.GetValue('BLF_QTEDEJAFACT'));
    end;
    TOBD.putValue('BLF_QTESITUATION',TOBL.GetValue('BLF_QTESITUATION'));
    TOBD.putValue('BLF_NATURETRAVAIL',TOBL.GetValue('GLC_NATURETRAVAIL'));
    TOBD.putValue('BLF_FOURNISSEUR',TOBL.GetValue('GL_FOURNISSEUR'));

    TOBD.SetAllModifie (true);
  end;

  procedure SetLignesDevisOuvDetail (TOBL,TOBOuvrage,TOBLDEVIS : TOB ; avancement : boolean);
  var indiceNomen,II : integer;
      TOBOUV,TOBOO,TOBD : TOB;
  begin
    IndiceNomen := TOBL.GetValue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
    //FV1 : 28/07/2016
    if TobOuvrage.detail.count = 0 then exit;
    if (Indicenomen-1) > (TobOuvrage.detail.count - 1) then exit;
    //
    TOBOUV := TOBOuvrage.detail[IndiceNomen-1];
    if TOBOUV = nil then exit;
    for II := 0 to TOBOUV.detail.count -1 do
    begin
      TOBOO := TOBOUV.detail[II];
      TOBD := TOB.Create ('LIGNEFAC',TOBLDEVIS,-1);
      TOBD.putValue('BLF_NATUREPIECEG',TOBOO.getValue('BLO_NATUREPIECEG'));
      TOBD.putValue('BLF_SOUCHE',TOBOO.getValue('BLO_SOUCHE'));
      TOBD.putValue('BLF_DATEPIECE',TOBOO.getValue('BLO_DATEPIECE'));
      TOBD.putValue('BLF_NUMERO',TOBOO.getValue('BLO_NUMERO'));
      TOBD.putValue('BLF_INDICEG',TOBOO.getValue('BLO_INDICEG'));
      TOBD.putValue('BLF_NUMORDRE',0);
      TOBD.putValue('BLF_UNIQUEBLO',TOBOO.getValue('BLO_UNIQUEBLO'));
      TOBD.putValue('BLF_MTMARCHE',TOBOO.GetValue('BLF_MTMARCHE'));
    	TOBD.putValue('BLF_POURCENTAVANC',TOBOO.GetValue('BLF_POURCENTAVANC'));
      if Avancement then
      begin
        TOBD.putValue('BLF_MTDEJAFACT',TOBOO.GetValue('BLF_MTCUMULEFACT')-TOBOO.GetValue('BLF_MTSITUATION'));
        TOBD.putValue('BLF_MTCUMULEFACT',TOBOO.GetValue('BLF_MTCUMULEFACT')-TOBOO.GetValue('BLF_MTSITUATION'));
      end else
      begin
        TOBD.putValue('BLF_MTDEJAFACT',TOBOO.GetValue('BLF_MTDEJAFACT'));
      end;
      TOBD.putValue('BLF_MTSITUATION',TOBOO.GetValue('BLF_MTSITUATION'));
      TOBD.putValue('BLF_QTEMARCHE',TOBOO.GetValue('BLF_QTEMARCHE'));
      if Avancement then
      begin
        TOBD.putValue('BLF_QTECUMULEFACT',TOBOO.GetValue('BLF_QTECUMULEFACT')-TOBOO.GetValue('BLF_QTESITUATION'));
        TOBD.putValue('BLF_QTEDEJAFACT',TOBOO.GetValue('BLF_QTECUMULEFACT')-TOBOO.GetValue('BLF_QTESITUATION'));
      end else
      begin
        TOBD.putValue('BLF_QTEDEJAFACT',TOBOO.GetValue('BLF_QTEDEJAFACT'));
      end;
      TOBD.putValue('BLF_QTESITUATION',TOBOO.GetValue('BLF_QTESITUATION'));
      TOBD.putValue('BLF_NATURETRAVAIL',TOBOO.GetValue('BLF_NATURETRAVAIL'));
      TOBD.putValue('BLF_FOURNISSEUR',TOBOO.GetValue('BLF_FOURNISSEUR'));
      TOBD.SetAllModifie (true);
    end;
  end;

var TOBLDevis,TOBL : TOB;
		Indice : integer;
begin
	result := true;
  TOBLDEVIS := TOB.create('LES LIGNES FACTURES DEVIS',nil,-1);
  //
  TRY
    for Indice := 0 to TOBLIGNES.detail.count -1 do
    begin
      TOBL := TOBLIGNES.detail[Indice];
      if TOBL.GetValue('GL_TYPELIGNE')<>'ART' then continue;
      if (TOBL.GetValue('GL_NUMORDRE')<>0) then
      begin
        //
        SetLigneDevis (TOBL,TOBLdevis,Avancement);
        if isOuvrage(TOBL) then
        begin
          SetLignesDevisOuvDetail (TOBL,TOBOuvrages,TOBLDEVIS,avancement);
        end;
      end;
    end;
    //
    if TOBLDEVIS.detail.count <> 0 then
    begin
    	for indice := 0 to TOBLDevis.detail.count -1 do
      begin
        TOBLDEVIS.detail[Indice].DeleteDB (false);
        if TOBLDEVIS.detail[Indice].Getvalue('BLF_MTSITUATION')<>0 then
        begin
          if not TOBLDEVIS.detail[Indice].InsertDB (nil,false) then
          BEGIN
            result := false;
            break;
          END;
        end;
      end;
    end;
    //
  FINALLY
    TOBLDEVIS.Free;
  end;

end;

procedure TOF_BTSAISIEAVANC.EnlevepasConcerne (TOBlignes : TOB; TheFournisseur : string; SousTraitancePA : Boolean=false );

	procedure  RecalculeOuvEnPA (TOBL : TOB;IndiceNomen : integer);
  var Qteouv,Qte,QteDejaFact,PrixPourQte,QteDuDetail : Double;
  		MtMarche,MtDejaFact,PourcentAvanc : Double;
  		TOBOUV,TOBO : TOB;
      Indice : Integer;
  begin
  	// ceci n'est fait que pour les sous traitants en paiement direct
		if IndiceNomen = 0 then exit;
 		QteOuv := TOBL.GetValue('BLF_QTEMARCHE');
    // init
    TOBL.PutValue('BLF_MTMARCHE',0);
    TOBL.PutValue('BLF_QTEDEJAFACT',0);
    TOBL.PutValue('BLF_QTECUMULEFACT',0);
    TOBL.PutValue('BLF_MTCUMULEFACT',0);
    TOBL.PutValue('BLF_MTDEJAFACT',0);
    TOBL.PutValue('BLF_POURCENTAVANC',0);
    //
    //FV1 : 28/07/2016
    if TobOuvrages.detail.count = 0 then exit;
    if (Indicenomen-1) > (TobOuvrages.detail.count - 1) then exit;
    //
    TOBOUV := TOBOuvrages .detail[IndiceNomen-1];
    for Indice := 0 to TOBOUV.detail.count -1 do
    begin
			TOBO := TOBOUV.detail[Indice];
      Qte := TOBO.GetValue('BLF_QTEMARCHE');
      QteDejaFact := TOBO.GetValue('BLF_QTEDEJAFACT');
      PrixPourQte := TOBO.GetValue('BLO_PRIXPOURQTE'); if PrixPourQte = 0 then PrixPourQte := 1;
      MtMarche := ARRONDI(Qte*TOBO.GetValue('BLO_DPA')/prixPourQte,DEV.Decimale);
      MtDejaFact := ARRONDI(QteDejaFact*TOBO.GetValue('BLO_DPA')/prixPourQte,DEV.Decimale);
      TOBL.PutValue('BLF_MTMARCHE',ARRONDI(TOBL.GetValue('BLF_MTMARCHE')+MtMarche,DEV.Decimale));
      TOBL.PutValue('BLF_MTCUMULEFACT',ARRONDI(TOBL.GetValue('BLF_MTCUMULEFACT')+MtDejaFact,DEV.Decimale));
      TOBL.PutValue('BLF_MTDEJAFACT',ARRONDI(TOBL.GetValue('BLF_MTDEJAFACT')+MtDejaFact,DEV.decimale));
    end;
    if Avancement then
    begin
//      PourcentAvanc := Arrondi(TOBL.GetValue('BLF_MTMARCHE') - TOBL.GetValue('BLF_MTDEJAFACT') / TOBL.GetValue('BLF_MTMARCHE') * 100,2);
      PourcentAvanc := Arrondi(TOBL.GetValue('BLF_MTDEJAFACT') / TOBL.GetValue('BLF_MTMARCHE') * 100,2);
    	TOBL.PutValue('BLF_POURCENTAVANC',PourcentAvanc);
      TOBL.PutValue('BLF_QTEDEJAFACT', Arrondi(QteOuv*PourcentAvanc/100,V_PGI.okdecQ));
      TOBL.PutValue('BLF_QTECUMULEFACT', Arrondi(QteOuv*PourcentAvanc/100,V_PGI.okdecQ));
    end else
    begin
      PourcentAvanc := Arrondi(TOBL.GetValue('BLF_MTDEJAFACT') / TOBL.GetValue('BLF_MTMARCHE') * 100,2);
      TOBL.PutValue('BLF_QTEDEJAFACT', Arrondi(QteOuv*PourcentAvanc / 100,V_PGI.okdecQ));
      TOBL.PutValue('BLF_QTECUMULEFACT', Arrondi(TOBL.GetValue('BLF_QTEDEJAFACT'),V_PGI.okdecQ));
    end;
    //
    TOBL.PutValue('GL_PUHTDEV',TOBL.GetValue('BLF_MTMARCHE')/TOBL.GetValue('BLF_QTEMARCHE'));

  end;

	procedure  EnleveOuvrageEtDetail (TOBlignes: TOB; Indice,NumOrdre : integer);
  var TOBL : TOB;
  begin
  	repeat
      TOBL := TOBlignes.detail[Indice];
      if TOBL.getValue('GL_NUMORDRE')<>NumOrdre then break;
      TOBL.free;
    until Indice >= TOBlignes.detail.count;
  end;

  function FindOuvragePere (TOBlignes : TOB;  IndiceNomen, Indice : integer) : integer;
  var ind : integer;
  		TOBL : TOB;
  begin
  	result := -1;
  	for ind := Indice downto 0 do
    begin
    	TOBL := TOBLignes.detail[Ind];
      if isOuvrage(TOBL) and
      	(not IsSousDetail(TOBL)) and
        (TOBL.Getvalue('GL_INDICENOMEN')=IndiceNomen) and
        (TOBL.GetValue('UNIQUEBLO')= 0) then
      begin
      	result := Ind;
        break;
      end;
    end;
  end;

  procedure RecalculeOuv (TOBlignes : TOB; Indice : Integer; IndiceNomen : integer);
  var II : integer;
  		MtDejaFact,MtDetDev,PourcentAvanc,QteOuv : double;
      MtMarche   : Double;
      TOBLP,TOBL : TOB;
  begin
    TOBLP := TOBlignes.detail[Indice];
 		QteOuv := TOBLP.GetValue('BLF_QTEMARCHE');
    TOBLP.PutValue('BLF_MTMARCHE',0);
    TOBLP.PutValue('BLF_QTEDEJAFACT',0);
    TOBLP.PutValue('BLF_QTECUMULEFACT',0);
    TOBLP.PutValue('BLF_MTCUMULEFACT',0);
    TOBLP.PutValue('BLF_MTDEJAFACT',0);
    TOBLP.PutValue('BLF_POURCENTAVANC',0);
    For II := Indice to TOBlignes.Detail.Count -1 do
    begin
      TOBL := TOBLignes.detail[II];
      if TOBL.GetValue('GL_INDICENOMEN') <> IndiceNomen then break;
      TOBLP.PutValue('BLF_MTMARCHE',TOBLP.GetValue('BLF_MTMARCHE')+TOBL.GetValue('BLF_MTMARCHE'));
      TOBLP.PutValue('BLF_MTCUMULEFACT',TOBLP.GetValue('BLF_MTCUMULEFACT')+TOBL.GetValue('BLF_MTCUMULEFACT'));
      TOBLP.PutValue('BLF_MTDEJAFACT',TOBLP.GetValue('BLF_MTDEJAFACT')+TOBL.GetValue('BLF_MTDEJAFACT'));
      (*
      MtDejaFact := TOBLP.GetValue('BLF_MTCUMULEFACT');
      MtDetDev := TOBLP.GetValue('BLF_MTMARCHE');
      if MtDetDev <> 0 then PourcentAvanc := Arrondi(MtDejaFact / MtDetDev * 100,2) else PourcentAvanc := 0;
      TOBLP.PutValue('BLF_POURCENTAVANC',PourcentAvanc);
      TOBLP.PutValue('BLF_QTECUMULEFACT',ARRONDI(TOBLP.GetValue('BLF_QTECUMULEFACT')*PourcentAvanc/100,V_PGI.OkDecQ));
      TOBLP.PutValue('BLF_QTEDEJAFACT',ARRONDI(TOBLP.GetValue('BLF_QTEDEJAFACT')*PourcentAvanc/100,V_PGI.OkDecQ));
      *)
    end;
    if Avancement then
    begin
      if TOBLP.GetValue('BLF_MTMARCHE') <> 0 then
      begin
      	PourcentAvanc := Arrondi(TOBLP.GetValue('BLF_MTDEJAFACT') / TOBLP.GetValue('BLF_MTMARCHE') * 100,2);
      end else
      begin
      	PourcentAvanc := 0;
      end;

    	TOBLP.PutValue('BLF_POURCENTAVANC',PourcentAvanc);
      TOBLP.PutValue('BLF_QTEDEJAFACT', Arrondi(QteOuv*PourcentAvanc/100,V_PGI.okdecQ));
      TOBLP.PutValue('BLF_QTECUMULEFACT', Arrondi(QteOuv*PourcentAvanc/100,V_PGI.okdecQ));
    end else
    begin
      MtDejaFact  := TOBLP.GetValue('BLF_MTDEJAFACT');
      MtMarche    := TOBLP.GetValue('BLF_MTMARCHE');
      PourcentAvanc := 0;
      if MtMarche <> 0 then PourcentAvanc := Arrondi(MtDejaFact / MtMarche * 100,2);
      TOBLP.PutValue('BLF_QTEDEJAFACT', Arrondi(QteOuv*PourcentAvanc / 100,V_PGI.okdecQ));
      TOBLP.PutValue('BLF_QTECUMULEFACT', Arrondi(QteOuv*PourcentAvanc / 100,V_PGI.okdecQ));
    end;
    if TOBLP.GetValue('BLF_QTEMARCHE') <> 0 then
    begin
    	TOBLP.PutValue('GL_PUHTDEV',TOBLP.GetValue('BLF_MTMARCHE')/TOBLP.GetValue('BLF_QTEMARCHE'));
    end;
  end;

  procedure  RecalculeAvancEnPa (TOBL : TOB);
  var Pua : Double;
  		Qte : double;
      PrixPourQte : Double;
      MtMarche,MtCumuleFac : double;
  begin
 		Qte := TOBL.GetValue('BLF_QTEMARCHE');
    PrixPourQte := TOBL.GetValue('GL_PRIXPOURQTE'); if PrixPourQte = 0 then PrixPourQte := 1;
    MtMarche := ARRONDI(Qte*TOBL.GetValue('GL_DPA')/prixPourQte,DEV.Decimale);
    MtCumuleFac := 0;
    if Avancement then
    begin
    	MtCumulefac := ARRONDI(MTMARCHE * TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.Decimale);
    end;
		TOBL.PutValue('BLF_MTMARCHE',MtMarche);
		TOBL.PutValue('BLF_MTCUMULEFACT',MtCumulefac);
		TOBL.PutValue('BLF_MTDEJAFACT',MtCumulefac);
    TOBL.PutValue('GL_PUHTDEV',TOBL.GetValue('BLF_MTMARCHE')/TOBL.GetValue('BLF_QTEMARCHE'));
  end;


var Indice,II : integer;
		TOBl,TOBLP: TOB;
    Typetravail : double;
    TypeLigne,Fournisseur,NatureTravail : string;
    Numordre,IndiceNomen : integer;
begin
	indice := 0;
  if TOBlignes.detail.count = 0 then exit;

	repeat
    TOBL := TOBlignes.detail[indice];
    if TOBL = nil then exit;
    NatureTravail := TOBL.getvalue('GLC_NATURETRAVAIL');
    TypeTravail := valeur(TOBL.getvalue('GLC_NATURETRAVAIL'));
    TypeLigne := TOBL.getvalue('GL_TYPELIGNE');
    Numordre := TOBL.getvalue('GL_NUMORDRE');
    fournisseur := TOBL.getvalue('GL_FOURNISSEUR'); if (Typetravail = 0) then Fournisseur := '';
    IndiceNomen := TOBL.getvalue('GL_INDICENOMEN');
    if (NatureTravail = '') and (TheFournisseur='') then
    begin
      if (TypeLigne='SD') and (TheFournisseur= Fournisseur) then
      begin
        if FindOuvragePere (TOBLignes,IndiceNomen, Indice) < 0 then
        begin
          TOBL.PutValue('GL_TYPELIGNE','ART'); // --> transformation en ligne article pour la saisie
        end;
      end;
    	inc(Indice);
    	continue;
    end else if (isOuvrage(TOBL)) and (not IsSousDetail(TOBL)) and
       (TypeTravail>=0) and (TypeTravail < 10) and
       (TheFournisseur <> Fournisseur) then
    begin
      EnleveOuvrageEtDetail (TOBlignes,Indice,NumOrdre);
      continue;
    end else if (isOuvrage(TOBL)) and (not IsSousDetail(TOBL)) and
       (TypeTravail>0) and (TypeTravail < 10) and
       (TheFournisseur = Fournisseur) and (SousTraitancePA) then
    begin
      TOBL.PutValue('ENPA','X');
    	RecalculeOuvEnPA (TOBL,IndiceNomen);
    end else if (isOuvrage(TOBL)) and (not IsSousDetail(TOBL)) and
						{(TheFournisseur<>'') and} (Typetravail >= 10) then // partiellement cotraité ou sous traité
    begin
    	TOBL.PutValue('GL_TYPELIGNE','OUP'); // --> transformation de l'ouvrage en ouvrage partiel
    end else if (TypeLigne='SD') and (TheFournisseur= Fournisseur) then
    begin
    	if FindOuvragePere (TOBLignes,IndiceNomen, Indice) < 0 then
      begin
    		TOBL.PutValue('GL_TYPELIGNE','ART'); // --> transformation en ligne article pour la saisie
      end;
      if SousTraitancePA then
      begin
        RecalculeAvancEnPa (TOBL);
        II := FindOuvragePere (TOBLignes,IndiceNomen, Indice);
        if II >= 0 then
        begin
          TOBLP := TOBlignes.detail[Indice];
      		TOBLP.PutValue('ENPA','X');
        	RecalculeOuv (TOBlignes,II,IndiceNomen);
        end;
      end;
    end else if (TypeLigne='COM') and
             (TypeTravail>=0) and (TypeTravail < 10) and
             (TheFournisseur <> Fournisseur) then
    begin
      	TOBL.free;
        continue;
    end else
    begin
    	if (TypeLigne <> 'REF') and  (not IsParagraphe(TOBL)) and (TheFournisseur <> fournisseur) then
      begin
      	if IsSousDetail(TOBL) then
        begin
        	II := FindOuvragePere (TOBLignes,IndiceNomen, Indice);
          TOBL.free;
          if II >= 0 then RecalculeOuv (TOBlignes,II,IndiceNomen);
          continue;
        end;
      	TOBL.free;
        continue;
      end;
    end;
    inc(Indice);
  until Indice >= TOBlignes.detail.count;
  NettoieOuvrageorphelin (TOBlignes);
  Nettoieparagraphes (TOBlignes);
end;

procedure TOF_BTSAISIEAVANC.NettoieOuvrageOrphelin (TOBlignes : TOB);
var Indice,indSuiv : integer;
    TOBL : TOB;
begin
	Indice := 0;
  repeat
    TOBL := TOBlignes.detail[Indice];
    if TOBL.GetValue('GL_TYPELIGNE')='OUP' then
    begin
			indSuiv := Indice +1;
      if indSuiv < TOBlignes.detail.count then
      begin
      	if TOBlignes.detail[indSuiv].GetValue('GL_TYPELIGNE') <> 'SD' then
        begin
          TOBL.Free;
        end else
        begin
          inc(Indice);
        end;
      end else
      begin
        TOBL.Free;
      end;
    end else inc(indice);
  until Indice >= TOBlignes.detail.count;
end;

procedure TOF_BTSAISIEAVANC.Nettoieparagraphes (TOBLignes : TOB);
var Indice : integer;
    TOBL : TOB;
    ListeDoc : Tlistref;
    Doc : TReference;
begin
	ListeDoc := Tlistref.create;

  Indice := TOBLignes.detail.count -1;
  repeat
    if Indice >= 0 then
    begin
      TOBL := TOBLignes.detail[Indice];
      if IsFinParagraphe(TOBL) then
      begin
        if not IsDetailInside (TOBlignes,Indice,TOBL) then
        begin
          SuprimeThisParagraphe (TOBLignes,TOBL,Indice);
          continue;
        end;
      end;
      dec(Indice);
    end;
  until indice < 0;
  //
  for Indice := 0 to TOBLignes.detail.count -1 do
  begin
    TOBL := TOBLignes.detail[Indice];
    if TOBL.getValue('GL_TYPELIGNE')='REF' then
    begin
      Doc := Treference.create;
      Doc.Adddoc (TOBL);
      Doc.TOBAssocie := TOBL;
      ListeDoc.Add(Doc);
    end else
    begin
      if Pos(TOBL.getValue('GL_TYPELIGNE'),'ART;SD;COM') >0 then
      begin
        Doc := ListeDoc.find (TOBL);
        if Doc <> nil then Doc.IncNombre;
      end;
    end;
  end;

  for Indice := 0 to ListeDoc.Count -1 do
  begin
  	doc := ListeDoc.Items[Indice];
    if doc.nombre = 0 then doc.TOBAssocie.free;
  end;

  for Indice := 0 to TOBLignes.detail.count -1 do
  begin
    TOBLignes.detail[Indice].putValue('INDICE',indice+1); // numérotation de la ligne
	end;

  ListeDoc.free;
end;

function TOF_BTSAISIEAVANC.IsDetailInside(TOBLignes : TOB; IndDep: integer; TOBL: TOB): boolean;
var indice : integer;
    Niveau : integer;
    TOBI : TOB;
begin
  result := false;
  Niveau := TOBL.GetValue('GL_NIVEAUIMBRIC');
  for Indice := Inddep-1 downto 0 do
  begin
    TOBI := TOBLignes.detail[Indice];
    if (IsArticle (TOBI)) OR (TOBI.GetString('GL_TYPELIGNE')='OUP') then BEGIN result := true; break; END;
    if IsDebutParagraphe (TOBI,Niveau) then break;
  end;
end;

procedure TOF_BTSAISIEAVANC.SuprimeThisParagraphe(TOBLignes,TOBL : TOB; var IndDep: integer);
var Niveau : integer;
    TOBI : TOB;
    StopItNow : boolean;
begin
  Niveau := TOBL.GetValue('GL_NIVEAUIMBRIC');
  StopItNow := false;
  repeat
    TOBI := TOBLignes.detail[IndDep];
    if IsDebutParagraphe (TOBI,Niveau) then StopItNow := true;
    TOBI.free;
    Dec(IndDep);
  until (IndDep < 0 ) or (StopItNow);
end;

procedure TOF_BTSAISIEAVANC.BeforeChangeVue (Sender: TObject; var AllowChange: Boolean);
var isModified : Boolean;
begin
  ShowButton (fCurrentOnglet,0,false);
  if fCurrentOnglet.GS.InplaceEditor.Modified then
  begin
		SendMessage(fCurrentOnglet.fGS.Handle, WM_KeyDown, VK_TAB, 0);
  end;
end;

procedure TOF_BTSAISIEAVANC.ChangeVue(Sender: Tobject);
begin
  SetEventsGrid (GS,false);
  fCurrentOnglet.fTVDevis.visible := false;
  SetOnglet(fOnglets.findOnglet(PS.ActivePage.Name));
//  SetLigne ( GS,GS.row,LIBELLE);
(*  for Indice := 0 to FOngletGlobal.GS.ColCount -1 do
  begin
  	GS.ColWidths [indice] := FOngletGlobal.GS.ColWidths [indice];
  end;
*)
  fCurrentOnglet.GS.Invalidate;
  stPrev := fCurrentOnglet.GS.cells[fCurrentOnglet.GS.col,fCurrentOnglet.GS.row]; // evite les effets de bord lors du passage sur la grid
	ShowButton(fCurrentOnglet,fCurrentOnglet.GS.row,true);
  SetEventsGrid (fCurrentOnglet.GS,True);
end;

procedure TOF_BTSAISIEAVANC.GSColumnsWithChanging (Sender : Tobject);
var indice : integer;
		ind1 : integer;
begin
  inherited;
  for Ind1 := 0 to fOnglets.Count -1 do
  begin
  	if fOnglets.Items [ind1] <> fCurrentOnglet then
    begin
      for Indice := 0 to fCurrentOnglet.GS.ColCount -1 do
      begin
        fOnglets.Items [ind1].GS.ColWidths [indice] := fCurrentOnglet.GS.ColWidths [indice];
      end;
    end;
  end;
  GS.Invalidate;
END;


procedure TOF_BTSAISIEAVANC.FormResize(Sender: Tobject);
var indice : integer;
begin
  inherited;
  for Indice := 0 to FOngletGlobal.GS.ColCount -1 do
  begin
  	GS.ColWidths [indice] := FOngletGlobal.GS.ColWidths [indice];
  end;
  GS.Invalidate;
end;

{ TlisteTOB }

function TlisteTOB.Add(AObject: TOB): Integer;
begin
  result := Inherited Add(AObject);
end;

destructor TlisteTOB.destroy;
begin
  inherited;
end;


function TlisteTOB.GetItems(Indice: integer): TOB;
begin
  result := TOB (Inherited Items[Indice]);
end;

procedure TlisteTOB.SetItems(Indice: integer; const Value: TOB);
begin
  Inherited Items[Indice]:= Value;
end;


{ Tlistref }

function Tlistref.Add(AObject: TReference): Integer;
begin
	result := inherited add(Aobject);
end;

destructor Tlistref.destroy;
var indice : integer;
begin
  if count > 0 then
  begin
    for Indice := count -1 downto 0 do
    begin
      if TReference(Items [Indice])<> nil then TReference(Items [Indice]).free;
    end;
  end;
  inherited;
end;

function Tlistref.find(TOBL: TOB): Treference;
var Indice : integer;
		ref : string;
begin
  result := nil;
	Ref := TOBL.getValue('GL_NATUREPIECEG')+';'+TOBL.getValue('GL_SOUCHE')+';'+
  				IntToStr(TOBL.getValue('GL_NUMERO'))+';'+intToStr(TOBL.getValue('GL_INDICEG'));
  for Indice := 0 to Count -1 do
  begin
    if Items[Indice].fDoc = Ref then
    begin
      result:=Items[Indice];
      break;
    end;
  end;
end;

function Tlistref.GetItems(Indice: integer): TReference;
begin
  result := Treference (Inherited Items[Indice]);
end;

procedure Tlistref.SetItems(Indice: integer; const Value: TReference);
begin
  Inherited Items[Indice]:= Value;
end;

{ TReference }

procedure TReference.Adddoc(TOBL: TOB);
begin
	fdoc := TOBL.getValue('GL_NATUREPIECEG')+';'+TOBL.getValue('GL_SOUCHE')+';'+
  				IntToStr(TOBL.getValue('GL_NUMERO'))+';'+intToStr(TOBL.getValue('GL_INDICEG'));
end;

constructor TReference.create;
begin
	fTOB := nil;
  fNombre := 0;
end;

destructor TReference.destroy;
begin
  inherited;
end;

procedure TReference.IncNombre;
begin
	inc(fNombre);
end;


procedure TOF_BTSAISIEAVANC.CalculeTotauxSaisie;
var Indice : integer;
		TOBl : TOB;
    iTot : integer;
    TypeArticle : string;
begin
	fCurrentOnglet.MtSituation := 0;
  dMtSituation := 0;
  for Indice := 0 to fCurrentOnglet.TOBlignes.detail.count -1 do
  begin
    TOBL := fCurrentOnglet.TOBlignes.detail[Indice];
    if indice = 0 then iTot := TOBL.GetNumChamp('GL_TYPELIGNE');
    TypeArticle := TOBL.getValue('GL_TYPEARTICLE');
    if TOBL.GetValue('GL_TYPELIGNE')='ART' then CumuleLigne(TOBL);
    CalculeSousTotal (indice+1,iTot);
  end;
end;
(*
procedure TOF_BTSAISIEAVANC.SaisieReglClick(Sender: TObject);
//var LaTOBRegl : TOB;
begin
  //retour := AGLLanceFiche('BTP','BTSAISREGL','','','');
end;
*)
function TOF_BTSAISIEAVANC.IsOuvragePartiel(TOBL: TOB): boolean;
var prefixe : string;
begin
  result := false;
  Prefixe := GetPrefixeTable (TOBL);
  if prefixe = '' then exit;
  if (TOBL.GetValue(prefixe+'_TYPELIGNE')='OUP') then result := true;
end;

procedure TOF_BTSAISIEAVANC.EnregistrelesSousTraits(TOBPiece: TOB);

	function SSTExists (TOBSS : TOB) : Boolean;
  begin
    Result := false;
    if TOBSOUSTraits.FindFirst(['BPI_TIERSFOU'],[TOBSS.GetValue('BPI_TIERSFOU')],True)<>nil then
    	Result := True;
  end;

var TOBSST,TOBLST : TOB;
		refpiece : string;
    cledoc : R_CLEDOC;
    Indice : Integer;
begin
  TOBSST := TOB.Create ('LES SSTRAIT P',nil,-1);
  TRY
    refpiece := EncodeCleDoc(TOBPIECE);
    DecodeCleDoc(refpiece,Cledoc);
    LoadLesSousTraitants(cledoc,TOBSST);
    if TOBSST.detail.count = 0 then exit;
    indice := 0;
    repeat
      if not SSTExists (TOBSST.detail[Indice]) then
      begin
        TOBLST := TOBSST.detail[Indice];
        TOBLST.ChangeParent(TOBSOUSTraits,-1);
        TOBLST.SetInteger('BPI_ORDRE',TOBLST.GetIndex+1);
      end else inc(Indice);
    until Indice >= TOBSST.Detail.Count;
  FINALLY
  	TOBSST.Free;
  END;
end;

procedure TOF_BTSAISIEAVANC.ShowButton(Onglet : TOnglet ; Ligne: Integer; Visible: Boolean);
var TOBL : TOB;
		Arect : TRect;
    okok : boolean;
    NatureTravail : double;
begin
  if Onglet.fBOuvrir = nil then Exit;
  Onglet.fBOuvrir.visible := false;
  if not Visible then Exit;
  TOBL := GetTOBLigneAV(Onglet.TOBlignes,Ligne);
  Naturetravail := VALEUR(TOBL.GetValue('GLC_NATURETRAVAIL'));
  Okok := false;
  if isOuvrage(TOBL) and (not IsSousDetail(TOBL))  then
  begin
    Okok := true;
  end;
(*  Okok := (isOuvrage(TOBL)) and
  				(NatureTravail < 10) and
  				(not IsOuvragePartiel(TOBL)) and
          (not IsSousDetail(TOBL)) and
          (TOBL.GetValue('GL_TYPEARTICLE')<>'');
*)
  if (not okok ) then Exit;
  if (VOIRDETAIL <> -1) and (ModifSousDetail) and (TOBL <> nil) and (Action <> taConsult) then
  begin
    ARect := Onglet.GS.CellRect(VOIRDETAIL, Ligne);
    with Onglet do
    begin
      fBOuvrir.Top := Arect.top - GS.GridLineWidth;
      fBOuvrir.Left := Arect.Left;
      fBOuvrir.Width := Arect.Right - Arect.Left;
      fBOuvrir.Height := Arect.Bottom - Arect.Top;
      if TOBL.GetValue('GLC_VOIRDETAIL')='X' then
      begin
        fBOuvrir.ImageIndex := 1;
      end else
      begin
        fBOuvrir.ImageIndex := 0;
      end;
      fBOuvrir.visible := true;
    end;
  end;
end;

procedure TOF_BTSAISIEAVANC.CreateButtonOuvrir (Fonglet:TOnglet;indice : integer);
var NomImg : string;
		UneImage : TImage;
begin
  NomImg := 'IMGP0';
  UneImage := TImage(GetCOntrol(NomImg));
  //
  with fOnglet do
  begin
    fBOUVRIR.Name := 'BOUV'+inttostr(indice);
  	fBOUVRIR.Parent := fGS;
    fBOUVRIR.Width := UneImage.Width;
    fBOUVRIR.Height := UneImage.Height;
    fBOUVRIR.Images := ImgListPlusMoins;
    fBOUVRIR.Color   := $19a2fd;
    fBOUVRIR.Opaque := true;
    fBOUVRIR.Alignment  := taCenter;
    fBOUVRIR.Flat   := true;
    fBOUVRIR.visible   := false;
    fBOUVRIR.ParentColor   := false;
    fBOUVRIR.Mode2003 := True;
    fBOUVRIR.onClick := OuvrirClick;
  end;
end;

procedure TOF_BTSAISIEAVANC.OuvrirDansOnglet(Onglet : TOnglet;ligne : integer);
var Indice,IndiceDep: integer;
		TOBL : TOB;
    EnAchat : Boolean;
begin
  ShowButton (Onglet,ligne,false);
  //
	indice := Ligne -1;
  IndiceDep := Indice;
  TOBL := Onglet.fTOBlignes.detail[Indice];
  EnAchat := (TOBL.GetValue('ENPA')='X');
  //
  if TOBL.GetValue('GLC_VOIRDETAIL')='-' then
  begin
    TOBL.PutValue('GLC_VOIRDETAIL','X');
    AfficheSousDetail (Onglet,Onglet.fTOBlignes,TOBL,Indice,EnAchat);
  end else
  begin
    TOBL.PutValue('GLC_VOIRDETAIL','-');
    SupprimeSousDetail (Onglet.fTOBlignes,Indice);
  end;
  AfficheLaGrille (Onglet.fGS,Onglet.fTOBlignes,Ligne);
  Onglet.fGS.InvalidateRow(Ligne);
  ShowButton (Onglet,Ligne,True);
end;

procedure TOF_BTSAISIEAVANC.OuvrirClick(Sender: TObject);
var Indice,IndiceDep,Ligne,IndInOnglet : integer;
		TOBL : TOB;
    laListe : TStringList;
    TheOnglet,LastOnglet : TOnglet;
    NumOrdre,NumPiece: Integer;
begin
  if fCurrentOnglet.fGS.Row = 0 then Exit;
  //
  lastOnglet := fCurrentOnglet;
  Ligne := fCurrentOnglet.fGS.Row;
	indice := Ligne -1;
  TOBL := fCurrentOnglet.fTOBlignes.detail[Indice];
  if TOBL.GetValue('GL_INDICENOMEN')=0 then exit;
  laListe := TStringList.Create;
  // stockage
  NumOrdre := TOBL.getvalue('GL_NUMORDRE');
  NumPiece := TOBL.getvalue('GL_NUMERO');
  // --
  OuvrirDansOnglet(fCurrentOnglet,ligne);
  //
  FindLesOnglets(NumPiece,NumOrdre,0,indice,fCurrentOnglet,LaListe);
  if laListe.Count > 0 then
  begin
    for Indice := 0 to laListe.Count -1 do
    begin
      IndInOnglet := -1;
      TheOnglet := fOnglets.findOnglet(laListe.strings[Indice]);
      if TheOnglet <> nil then
      begin
        IndInOnglet := FindTOBligne(TheOnglet,Numpiece,Numordre,0);
        if IndInOnglet >0 then
        begin
          SetOnglet(TheOnglet);
          OuvrirDansOnglet(TheOnglet,IndInOnglet+1);;
        end;
      end else
      begin
        if laListe.strings[Indice] = FOngletGlobal.Name then
        begin
          IndInOnglet := FindTOBligne(FOngletGlobal,Numpiece,Numordre,0);
          if IndInOnglet >0 then
          begin
            SetOnglet(FOngletGlobal);
            OuvrirDansOnglet(FOngletGlobal,IndInOnglet+1);;
          end;
        end;
      end;
    end;
  end;
  //
  SetOnglet(lastOnglet);
  LaListe.free;
end;

procedure TOF_BTSAISIEAVANC.SupprimeSousDetail (TOBLignes : TOB; Indice : integer);
var ind : Integer;
		TOBL : TOB;
begin
  ind := Indice + 1;
	repeat
		TOBL := TOBLignes.detail[ind];
    if not IsSousDetail (TOBL) then break;
    if (IsSousDetail(TOBL)) and (TOBL.GetValue('UNIQUEBLO')<>0) then TOBL.Free;
  until Ind  >= TOBLignes.Detail.Count;
end;

procedure TOF_BTSAISIEAVANC.AfficheSousDetail(Onglet : TOnglet; TOBLignes,TOBL: TOB; var Indice : integer; EnAchat : boolean=false);

	function IsLigneAffichable (Onglet: TOnglet ;TOBDO : TOB) : Boolean;
  var Fournisseur : string;
  begin
    Result := true;
    if Onglet = FOngletGlobal then exit;
    Result :=  (Onglet.fournisseur = TOBDO.GetString('BLO_FOURNISSEUR'));
  end;

var ind,IndiceNomen,IndO: Integer;
		TOBO : TOB;
    MtLigne : Double;
begin
  Ind := 1;
  TOBO := TOBOuvrages.Detail[TOBL.getValue('GL_INDICENOMEN')-1];
  IndiceNomen := TOBL.getValue('GL_INDICENOMEN');
  For IndO := 0 to TOBO.detail.count -1 do
  begin
    if IsLigneAffichable (Onglet,TOBO.detail[IndO]) then
    begin
    	AddLignefacturationFromOuv (TOBL,TOBLignes,TOBO.detail[IndO],Indice,Ind,IndiceNomen,MtLigne,(TOBO.GetValue('AVANCSOUSDET')='X'),EnAchat);
    	inc(Ind);
    end;
  end;
  Indice := Indice + Ind -1 ;
end;


function TOF_BTSAISIEAVANC.IsAvancementDetailOuvrage(TOBL: TOB): Boolean;
var TOBOUV,TOBO : TOB;
		IndiceNomen,Indice : Integer;
begin
  Result := false;
	IndiceNomen := TOBL.getValue('GL_INDICENOMEN') -1;
  //FV1 : 27/07/2016
  if TobOuvrages.detail.count = 0 then exit;
  if (IndiceNomen < 0) Or (IndiceNomen > TOBOuvrages.Detail.count) then Exit;
  TOBOUV := TOBOuvrages.Detail[IndiceNomen];
  for INDICE := 0 to TOBOUV.Detail.count -1 do
  begin
		if TOBOUV.detail[Indice].GetValue('BLF_POURCENTAVANC')<>TOBL.GetValue('BLF_POURCENTAVANC') then
    begin
      Result := True;
      break;
    end;
  end;
end;


procedure TOF_BTSAISIEAVANC.ReindiceOuvrages;

  function GetTobPiece(TOBOUV : TOB) : TOB;
  var II : Integer;
    TOBP : TOB;
  begin
    Result := nil;
    for II := 0 to TOBSELECT.detail.Count -1 do
    begin
      TOBP := TOBSELECT.detail[II];
      if (TOBP.GetValue('GP_NATUREPIECEG')=TOBOUV.GetValue('BLO_NATUREPIECEG')) AND
         (TOBP.GetValue('GP_SOUCHE')=TOBOUV.GetValue('BLO_SOUCHE')) AND
         (TOBP.GetValue('GP_NUMERO')=TOBOUV.GetValue('BLO_NUMERO')) AND
         (TOBP.GetValue('GP_INDICEG')=TOBOUV.GetValue('BLO_INDICEG')) then
         begin
           Result := TOBSELECT.detail[II];
           break;
         end;
    end;
  end;

  function  NewUniqueBLO (TOBouv : TOB) : integer;
  var UniqueBlo : Integer;
      TOBP : TOB;
  begin
    TOBP := GetTobPiece(TOBOUV);
    if TOBP = NIL then
    begin
      result := 0;
      Exit;
    end;
    TOBP.putvalue('GP_UNIQUEBLO',TOBP.Getvalue('GP_UNIQUEBLO')+1);
    result := TOBP.Getvalue('GP_UNIQUEBLO');
  end;

	procedure  IndiceUniqueBlo (TOBouv : TOB);
  var Indice : integer;
  begin
    For Indice := 0 to TOBouv.detail.count -1 do
    begin
      TOBOUV.detail[Indice].putValue('BLO_UNIQUEBLO',0);
      TOBOUV.detail[Indice].putValue('BLO_UNIQUEBLO',NewUniqueBLO (TOBOUV.detail[Indice]));
      TOBOUV.detail[Indice].putValue('BLF_UNIQUEBLO',TOBOUV.detail[Indice].GetValue('BLO_UNIQUEBLO'));
      if TOBOUV.detail[Indice].detail.count> 0 then
      begin
      	IndiceUniqueBlo (TOBOUV.detail[Indice]);
      end;
    end;
  end;

var Indice : integer;
begin
	Indice := 0;
  for Indice := 0 to TOBSelect.detail.count -1 do
  begin
    TOBSelect.detail[indice].putvalue('GP_UNIQUEBLO',0);
  end;
  for Indice := 0 to TOBouvrages.detail.count -1 do
  begin
  	IndiceUniqueBlo (TOBouvrages.detail[Indice]);
  end;
end;

procedure TOF_BTSAISIEAVANC.MemoriseHighNumOrdre(TOBL: TOB);
var TOBT : TOB;
    cledoc : R_CLEDOC;
    fclef : string;
begin
  if TOBL.GetString('GL_PIECEORIGINE')='' then Exit;
  DecodeRefPiece(TOBL.GetString('GL_PIECEORIGINE'),cledoc);
  fclef := cledoc.NaturePiece+';'+cledoc.Souche+';'+InttoStr(cledoc.NumeroPiece)+';'+InttoStr(cledoc.Indice);
  TOBT := TOBInfoPieces.FindFirst(['DOCUMENT'],[fclef],true);
  if TOBT=nil then
  begin
    TOBT:= TOB.Create('UN DOC',TOBInfoPieces,-1);
    TOBT.AddChampSupValeur('DOCUMENT',fclef);
    TOBT.AddChampSupValeur('LASTMAXORDRE',0);
  end;
  if cledoc.NumOrdre > TOBT.GetInteger('LASTMAXORDRE') then TOBT.SetInteger('LASTMAXORDRE',cledoc.NumOrdre);
end;

procedure TOF_BTSAISIEAVANC.SynchroniseDocumentWithDevis;

  procedure DecodeRefInfo (TOBI : TOB; var Cledoc : R_CLEDOC);
  var stC : string;
  begin
    stC := TOBI.GetString('DOCUMENT');
    CleDoc.NaturePiece := ReadTokenSt(stC);
    CleDoc.Souche := ReadTokenSt(stC);
    CleDoc.NumeroPiece := StrToInt(ReadTokenSt(stC));
    CleDoc.Indice := StrToInt(ReadTokenSt(stC));
  end;

  function FindPrecNumOrdre (TOBL,TOBRefLignes : TOB) : TOB;
  begin
    result := TOBRefLignes.FindFirst(['GL_NUMORDRE'],[TOBL.GetInteger('GL_NUMORDRE')],true);
  end;

  function EncodeRefLigneAFind (TOBL,TOBR : TOB) : string;
  var std : string;
  begin
      StD := FormatDateTime('ddmmyyyy', TOBL.GetDateTime('GL_DATEPIECE'));
      Result := StD + ';' + TOBL.getString('GL_NATUREPIECEG')  + ';' + TOBL.getString('GL_SOUCHE')  + ';' +
                IntToStr(TOBL.getInteger('GL_NUMERO')) + ';' +
                IntToStr(TOBL.getInteger('GL_INDICEG')) + ';' +
                IntToStr(TOBR.getInteger('PRECNUMORDRE')) + ';';
  end;

  function FindLigneInDocument (TOBLignes: TOB;reference : string): integer;
  var II : Integer;
      TOBL : TOB;
  begin
    Result := -1;
    TOBL := TOBLignes.FindFirst(['GL_PIECEORIGINE'],[Reference],true);
    if TOBL <> nil then Result := TOBL.GetIndex+1;
  end;

  function FindRefInDocument (TOBLignes : TOB; NumDoc : Integer) : Integer;
  var II : Integer;
      cledoc : r_cledoc;
  begin
    Result := -1;
    for II := 0 TO TOBLignes.detail.count -1 do
    begin
      if TOBLIGNES.detail[II].GetString('GL_TYPELIGNE')<>'REF' then Continue;
      DecodeRefPiece (TOBLIGNES.detail[II].getString('GL_PIECEPRECEDENTE'),cledoc);
      if cledoc.NumeroPiece = Numdoc then
      begin
        Result := TOBLIGNES.detail[II].GetIndex+1;
        Break;
      end;
    end;
  end;

  function findFinDocument (TOBLignes : TOB; NumDoc : Integer) : Integer;
  var II,Depart : Integer;
      cledoc : r_cledoc;
  begin
    Result := -1;
    Depart := FindRefInDocument (TOBlignes,Numdoc);
    if Depart = -1 then exit;
    for II := Depart+1 TO TOBLignes.detail.count -1 do
    begin
      if TOBLIGNES.detail[II].GetString('GL_TYPELIGNE')='REF' then
      begin
        DecodeRefPiece (TOBLIGNES.detail[II].getString('GL_PIECEPRECEDENTE'),cledoc);
      end else
      begin
        DecodeRefPiece (TOBLIGNES.detail[II].getString('GL_PIECEORIGINE'),cledoc);
      end;
      if cledoc.NumeroPiece > Numdoc then
      begin
        Result := TOBLIGNES.detail[II].GetIndex;
        Break;
      end;
    end;
  end;

  procedure InsereDetailOuvrage (TOBLIG : TOB);
  var Sql : string;
      QQ : TQuery;
      TOBNomen : TOB;
      CD : R_CLEDOC;
  begin
    TOBNomen := TOB.Create('LES DETAIL',nil,-1);
    TRY
      Cd := TOB2Cledoc(TOBLIG);
      Sql := MakeSelectLigneOuvBtp (True);
      Sql := Sql + ' WHERE '+WherePiece(CD,ttdOuvrage,true)+' AND '+
            'BLO_NUMLIGNE='+InttoStr(cd.NumLigne)+
            ' ORDER BY BLO_NUMLIGNE,BLO_N1,BLO_N2,BLO_N3,BLO_N4,BLO_N5';
      QQ := OpenSql (Sql,true,-1,'',true);
      if not QQ.eof then
      begin
        TOBNomen.LoadDetailDB('LIGNEOUV','','',QQ,true,true);
        OrganiseOuvrage(TOBLIG.parent,TOBNomen,TOBLIG);
      end;
      ferme (QQ);
    FINALLY
      TOBNomen.free;
    END;
  end;

  procedure AjouteUneLigneOuvrage(TOBATRAITER,TOBLL : TOB);
  var TOBTT : TOB;
  begin
    TOBTT := TOB.Create('UN OUVRAGE',TOBATRAITER,-1);
    TOBTT.AddChampSupValeur ('LALIGNE',TOBLL.GetIndex);
    TOBTT.Data := TOBLL;
  end;


var II,JJ,IposLigne : Integer;
    TOBI,TOBNL,TOBOrdres,TOBLL,TOBTT,TOBATRAITER : TOB;
    cledoc : R_CLEDOC;
    baseSql,SQl,RefligneOrigine,reforigine : string;
    QQ : TQuery;
    TypeArticle : string;
    NatureTravail : double;

begin
  // -- Sécurité -- on ne synchronise pas sur une facture qui ne comporte pas toutes les lignes des devis..sinon catastrophe
  if not GetParamSocSecur('SO_BTGESTVIDESUREDIT',false) then Exit;
  //
  TOBATRAITER := TOB.Create('LES LIGNES OUV',nil,-1);
  TOBOrdres := TOB.create ('LES TRAITES',nil,-1);
  TOBNL := TOB.create ('LES TRAITES',nil,-1);
  TRY
    II := 0;
    for II := 0 to TOBInfoPieces.Detail.count -1 do
    begin
      TOBI := TOBInfoPieces.detail[II];
      DecodeRefInfo (TOBI,cledoc);
      //
      Sql := 'SELECT L1.GL_NUMORDRE,'+
              '(SELECT L2.GL_NUMORDRE FROM LIGNE L2 WHERE '+
              'L2.GL_NATUREPIECEG=L1.GL_NATUREPIECEG AND '+
              'L2.GL_SOUCHE=L1.GL_SOUCHE AND '+
              'L2.GL_NUMERO=L1.GL_NUMERO AND '+
              'L2.GL_INDICEG=L1.GL_INDICEG AND '+
              'L2.GL_NUMLIGNE=(L1.GL_NUMLIGNE-1)) AS PRECNUMORDRE,0 AS NEWNUMORDRE ';
      Sql := SQl+ 'FROM LIGNE L1 '+
                'WHERE '+
                'L1.GL_NATUREPIECEG="'+Cledoc.NaturePiece+'" AND '+
                'L1.GL_SOUCHE="'+Cledoc.souche+'" AND '+
                'L1.GL_NUMERO='+InttoStr(cledoc.NumeroPiece)+' AND '+
                'L1.GL_INDICEG='+IntToStr(cledoc.Indice)+' AND '+
                'L1.GL_NUMORDRE > '+TOBI.GetString('LASTMAXORDRE')+ ' AND '+
                'L1.GL_TYPELIGNE NOT IN ("ARV","COV","SDV") AND '+
                'SUBSTRING(L1.GL_TYPELIGNE,1,2) NOT IN ("DV","TV") '+
                'ORDER BY L1.GL_NUMLIGNE';
      QQ := OpenSQL(SQl,True,-1,'',true);
      if not QQ.eof then
      begin
        TOBOrdres.LoadDetailDB('THIS EMPLAC','','',QQ,false);
      end;
      Ferme(QQ);
      //
      if TOBOrdres.detail.count > 0 then
      begin
        Sql := MakeSelectLigneBtp(true,false,true)+' WHERE  '+
               WherePiece (cledoc,ttdLigne,false)+ ' AND ';
        Sql := Sql + 'GL_NUMORDRE IN (';
        Sql := SQL + 'SELECT L1.GL_NUMORDRE ';
        Sql := SQl + 'FROM LIGNE L1 '+
                'WHERE '+
                'L1.GL_NATUREPIECEG="'+Cledoc.NaturePiece+'" AND '+
                'L1.GL_SOUCHE="'+Cledoc.souche+'" AND '+
                'L1.GL_NUMERO='+InttoStr(cledoc.NumeroPiece)+' AND '+
                'L1.GL_INDICEG='+IntToStr(cledoc.Indice)+' AND '+
                'L1.GL_NUMORDRE > '+TOBI.GetString('LASTMAXORDRE')+ ' AND '+
                'L1.GL_TYPELIGNE NOT IN ("ARV","COV","SDV") AND '+
                'SUBSTRING(L1.GL_TYPELIGNE,1,2) NOT IN ("DV","TV")';
        SQL := Sql + ') ';
        Sql := Sql +
               'ORDER BY GL_NUMLIGNE';
        QQ := OpenSQL(SQl,True,-1,'',true);
        TOBNL.LoadDetailDB('LIGNE','','',QQ,false);
        Ferme(QQ);
        JJ := 0;
        repeat
          TOBLL := TOBNL.detail[JJ];
          if Isvariante(TOBLL) then begin TOBLL.free; continue; END;
          //
          TOBTT := FindPrecNumOrdre(TOBLL,TOBOrdres);
          if TOBTT = nil then begin TOBLL.free; continue; END;

          if TOBTT.GetInteger('PRECNUMORDRE')= 0 then
          begin
            // Insertion au début du document d'origine
            IposLigne := FindRefInDocument (TOBLIGNES,TOBLL.GetInteger('GL_NUMERO'));
          end else
          begin
            // recherche de l'emplacement dans la facture courante
            RefligneOrigine := EncodeRefLigneAFind (TOBLL,TOBTT);
            IposLigne := FindLigneInDocument (TOBLignes,RefLigneOrigine);
          end;
          // si on arrive avec un indice a -1 c'est que l'on pas réussi a retrouver la ligne de ref du devis
          // --> insertion en fin du devis dans la facture
          if IposLigne = -1 then
          begin
            iPosligne := findFinDocument (TOBlignes,TOBLL.GetInteger('GL_NUMERO'));
          end;
          Inc(MaxNumOrdre);
          TOBLL.ChangeParent(TOBLIGNES,iposLigne);
          RefligneOrigine := EncodeRefPiece(TOBLL);
          TOBLL.SetString('GL_PIECEORIGINE',RefligneOrigine);
          TOBLL.SetString('GL_PIECEPRECEDENTE',RefligneOrigine);
          //
          TOBLL.SetInteger('GL_NUMORDRE',MaxNumOrdre);
          if isOuvrage(TOBLL) then
          begin
            AjouteUneLigneOuvrage(TOBATRAITER,TOBLL);
            InsereDetailOuvrage (TOBLL);
          end;
          AddLesSupLigne(TOBLL,false);
          AddChampsSupAvanc (TOBLL);
          AddChampsSupMtINITAvanc (TOBLL);
          PositionneLigne (TOBLL,IposLigne+1);
          //
          TOBLL.SetString('GL_NATUREPIECEG',TOBLIGNES.GetString('GP_NATUREPIECEG'));
          TOBLL.SetString('GL_SOUCHE',TOBLIGNES.GetString('GP_SOUCHE'));
          TOBLL.SetInteger('GL_NUMERO',TOBLIGNES.GetInteger('GP_NUMERO'));
          TOBLL.SetInteger('GL_INDICEG',TOBLIGNES.GetInteger('GP_INDICEG'));
          TOBLL.PutValue('GL_DATEPIECE',TOBLIGNES.GetValue('GP_DATEPIECE'));
          //
          //--
          // NOUVEAU GESTION DES OUVRAGES
          TypeArticle := TOBLL.getValue('GL_TYPEARTICLE');
          NatureTravail := valeur(TOBLL.GetValue('GLC_NATURETRAVAIL'));
          if IsOuvrage(TOBLL) and ((NatureTravail >= 10) or (ReajusteDossier)) then
          begin
            TOBLL.PutValue('GLC_VOIRDETAIL','X');
          end;
          if IsOuvrage(TOBLL) then
          begin
            if IsAvancementDetailOuvrage (TOBLL) then TOBLL.PutValue('GLC_VOIRDETAIL','X');
          end;
          if TOBLL.GetValue('GL_TYPELIGNE')='ART' then CumuleLigne(TOBLL);
          (*
          if (isOuvrage (TOBLL)) and  (pos(TypeArticle,'ARP;ARV;')=0) and (TOBLL.GetValue('GLC_VOIRDETAIL')='X') then
          begin
            AfficheSousDetail (fCurrentOnglet,TOBLIGNES,TOBLL,iPosligne);
          end;
          *)
        until JJ >= TOBNL.detail.Count;
      end;
    end;

    for II := 0 to TOBATRAITER.Detail.count -1 do
    begin
      TOBLL := TOB(TOBATRAITER.detail[II].Data);
      TypeArticle := TOBLL.GetString('GL_TYPEARTICLE');
      if (isOuvrage (TOBLL)) and  (pos(TypeArticle,'ARP;ARV;')=0) and (TOBLL.GetValue('GLC_VOIRDETAIL')='X') then
      begin
        iPosligne := TOBLL.GetIndex;
        AfficheSousDetail (fCurrentOnglet,TOBLIGNES,TOBLL,iPosLigne);
      end;
    end;
  FINALLY
    TOBOrdres.free;
    TOBNL.free;
    TOBATRAITER.Free;
  END;
end;

function TOF_BTSAISIEAVANC.TraitementGestionAvoir: boolean;
var Indice : integer;
    MtSituation : double;
    XX : TFsplashScreen;
begin
  result := false;
  MtSituation := 0;
  XX := TFsplashScreen.Create (application);
  XX.Label1.Caption  := TraduireMemoire('Vérification de la saisie en cours...');
  XX.Show;
  XX.BringToFront;
  XX.refresh;
  TRY
    for Indice := 0 to TOBLignes.detail.count - 1 do
    begin
      if TOBLignes.detail[Indice].getValue('GL_TYPELIGNE')<>'ART' then continue;
      MtSituation := MtSituation + TOBLignes.detail[Indice].getValue('BLF_MTSITUATION')
    end;
    if MtSituation < 0 then
    begin
      if PgiAsk('ATTENTION : le montant de cette saisie est négatif.#13#10 Désirez-vous générer un avoir ?') = mrYes then result := true;
    end;
  FINALLY
    XX.Free;
  END;
end;

Initialization
  registerclasses ( [ TOF_BTSAISIEAVANC ] ) ;
end.


{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 02/05/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : DECISIONACHAT ()
Mots clefs ... : TOF;DECISIONACHAT
*****************************************************************}
Unit DECISIONACHAT_TOF ;

Interface

uses {$IFDEF VER150} variants,{$ENDIF} StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
		 fe_main,
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     EdtREtat,
{$else}
		 maineagl,
     eMul,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Hpanel,
     HMsgBox,
     UTOF,
     UTOB,
     Windows,
     Graphics,
     grids,
     ImgList,
     Menus,
     ExtCtrls,
     HTB97,
     FactTOB,
     EntGc,
     FactComm,
     DecisionAchatUtil,
     CalcOLEGenericBTP,
     UtofListeInv,
     UtilDispGc,
     UtilSuggestion,
     DECISIONACHNART_TOF,
     Vierge,
     HRichOLE,
     paramsoc,
     UtilConsultFour,
     Ent1,Messages,Dialogs,
     Aglinit,uEntCommun
     ;

const MAXNiveauAff = 4;
Type
  TOF_DECISIONACHAT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    LastSelect : integer;
    OkConsultFou : boolean;
    CONSULTFOU : TGestConsultFou;
  	ImStateConsFou : TImageList;
  	MnVoirArticle,MnVoirStock,MnVoirBesoin,NewBesoin,AjoutLigne,MnScinder,Mnremplace,MnDeplier,MnRefermer,MnOptions : TmenuItem;
  	MnSep1,MnSupLigne  : TmenuItem;
    FenetreSup : TToolWindow97;
    CHOIXFOUR : TPopupMenu;
  	Action : TactionFiche;
  	ImTypeArticle,ImOpenClose : TImageList;
  	NbLigne,NbCols,NiveauMax : integer;
  	TOBDecisionnel : TOB;
    MnCatalogue,MnFournisseur : TmenuItem;
  	GS : THgrid;
    StoredRow : integer;
    GSX,GSY : integer;
    CellCur : string;
    BAD_DATELIVRAISON,REMISES,ARTREMPLACE : THedit;
    PINFOSTARIF : TPanel;
    TLIBELLEART,TLIBELLECHANTIER,TPROVENANCE,TLIBELLEFOURNISSEUR : THLabel;
    QUANTITE_VTE,QUANTITE_ACH,QUANTITE_STO,PRIXACHATBRUT,PRIXACHATNET : THNumEdit;
    UNITE_VTE,UNITE_ACH,UNITE_STO : THLabel;
    BLOCNOTE : ThRichEditOle;
    BBouton,BDelete,BImprimer,BViewDetail,BRECHARTICLE,BFINDASSOCIATION : TToolBarButton97;
    PRISEENCOMPTE,GERESTOCK,LIVCHANTIER,ASUPPRIMER : TCheckBox;
    TDEPOT : THlabel;
    DEPOT : THValComboBox;
    BVOIRSTOCK : TToolBarButton97;
    NumDecisionnel : String;
    ListeSaisie : string;
  	GS_OPEN,GS_OPEN1,GS_OPEN2,GS_TYPEARTICLE,GS_CODEARTICLE,GS_AFFAIRE,GS_QUANTITESTK : integer;
    GS_DEPOT,GS_QUALIFQTESTK,GS_LIVCHANTIER,GS_QTEPHY,GS_FOURNISSEUR,GS_PRIXACHNET,GS_PRISENCOMPTE,GS_POSITIONCONSULT : integer;
    GS_TENUESTOCK,GS_NUMREMPLACE,GS_PABASE : integer;
    BOPTIONMULTISEL : TToolbarButton97;
    CurrentTOb : TOB;
    LastPrixNet : double;
    MaxRemplace : integer;
    FindDialog: TFindDialog;
    FirstFind : Boolean;
    TOBLAffiche : TOB;
    LastRow,NiveauAffiche : integer;
    fTexte : THRichEditOle;
    fInternalWindow : TToolWindow97;
    GOptions : TGroupBox;
    // consultation fournisseur
    // --
		procedure AfficheInfos(Ligne : integer);
    procedure ActiveEventGS;
    procedure ActiveEvents;
    procedure ASupprimerClick (Sender : Tobject);
    procedure BBoutonClick(Sender: TObject);
    procedure BdeleteClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure CreateControls;
    procedure ChargeControls;
  	procedure ChargeDecisionnel;
    procedure CumulStkPhysique (TOBdecisionnel : TOB);
  	procedure DefiniGrilleSaisie;
    procedure DefiniNiveauVisible(Niveau: integer);
    procedure DesactiveEventGS;
  	procedure DesactiveEvents;
    function  FindEtablissement : string;
    procedure FreeControls;
    procedure GetCellCanvas(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GetControls;
    function  GetRow: integer;
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
    procedure GSElipsisClick (Sender : TObject);
    procedure GSMouseMove(Sender: TObject; Shift: TShiftState; X,Y: Integer);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean;Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean;Chg: Boolean);
    procedure GSDblClick (Sender : TOBject);
    procedure InscritDecisionnel;
    procedure MajDecisionnel;
    procedure PositionneLigne(Arow: integer;Acol: integer =-1);
    procedure PostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure QUANTITE_ACHChange (Sender : TObject);
    procedure RecalculNiveauSup (TOBL : TOB);
    procedure REMISESChange (Sender : TOBject);
    procedure RenumeroteDecisonnel;
    procedure SetEcran;
    procedure SetInfos(arguments: string);
    procedure SetEvents;
    procedure SetRow(Arow: integer);
    procedure SetEnabled(Arow: integer);
    procedure SupprimeDecisionnelAch;
    procedure ShowButton(Ou: integer);
    function ZoneAccessible(ACol, ARow: Integer): boolean;
    procedure ZoneSuivanteOuOk(var ACol, ARow: Integer;var Cancel: boolean);
    procedure CalculAndAfficheNiveauSup(TOBL: TOB);
    procedure CumulThisInPapou(TOBD, TOBP: TOB;Sens : string='+');
    procedure ReinitNiveauSup(TOBL: TOB);
    procedure PRISEENCOMPTEChange (Sender : TObject);
    procedure SetpriseEncompte(TOBL : TOB;Actif : boolean);
    procedure ChangeQteStk(Arow: integer; NewValue : double);
    procedure ChangeFournisseur (Arow : integer; NewValue : String; var cancel : boolean);
    procedure ChangePrixNet (Arow : integer; NewValue : double);
    procedure GERESTOCKChange (Sender : TObject);
    procedure SetPrisEnStock(TOBL: TOB; Actif: boolean);
    procedure LIVCHANTIERChange (Sender : TOBject);
    procedure InitligneDecisionnel(TOBpere, TOBL: TOB; Niveau : integer);
    procedure MepLigneInDecisionnel(TOBL : TOB; ChangementParent : boolean=true);
    procedure SetLivSurChantier(TOBL: TOB; Actif: boolean);
    procedure NettoieParent(TheParent: TOB);
    procedure ReindiceDecisionnel;
    procedure SetNumligne(TOBPere: TOB; IndiceN1, IndiceN2, IndiceN3,
      IndiceN4, Niveau: integer);
    procedure SetFournisseur (TOBL : TOB);
    procedure SetPrixAchNet (TOBP : TOB);
    procedure EnleveDesPeres(TOBL, TOBP: TOB);
    procedure AjouteAuxPeres(TOBL, TOBP: TOB);
    procedure VoirArticleClick (Sender : TObject);
    procedure VoirStockClick(Sender: TObject);
		procedure VOIRESTOCKSClick (Sender : TObject);
    procedure VoirBesoinClick (Sender: TObject);
  	procedure NewbesoinClick (sender : TObject);
  	procedure AjoutligneClick (Sender : TOBject);
  	procedure MnScinderClick (Sender : Tobject);
		procedure MnSupLigneClick (Sender : Tobject);
    procedure MnRemplaceClick(Sender: Tobject);
    procedure SetMenuEnabled(TOBL: TOB);
    procedure InitAjoutLigne(TOBPere: TOB; Article, Depot,Livchantier: string; Niveau: integer);
    function FindParent(Article, Livchantier, Depot: string): TOB;
    function FindpreviousForDelete(TOBdecisionnel : TOB;Ligne : integer) : TOB;
    procedure BAD_DATELIVRAISONClick (Sender : TObject);
    procedure SetDateLivraison (TOBL : TOB);
		procedure SetInfosLivraison2Fils (TOBP : TOB);
    procedure BViewDetailClick (Sender : TOBject);
    procedure FenetreSupClose (Sender : TObject);
    procedure changeAffaire(Ligne: integer; TheAffaire: string; var cancel : boolean);
    function RechercheAffaire(var TheAffaire: string): boolean;
    function CalculPrixNet: double; overload;
    procedure PRIXACHATNETChange(Sender: Tobject);
    procedure PRIXACHATNETEnter(Sender: Tobject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
    procedure ReactiveLigneBase(NumRemplace: integer);
    procedure SetInfosFournisseur2Fils(TOBP: TOB; FUV, FUS, FUA, CoefuaUS,CoefUSUV: double);
		procedure BrechArticleClick (Sender : TOBject);
  	procedure BFindAssociationClick (Sender : Tobject);
    procedure FindDialogFind (Sender : TObject);
    procedure MnCatalogueClick(Sender: TObject);
    procedure MnFournisseurClick(Sender: TObject);
    procedure GSMouSeDown(Sender: TObject; Button: TMouseButton;Shift: TShiftState; X, Y: Integer);
    procedure SetPrisEnSTockpapou(TOBL: TOB);
    function NettoieFilles(TheParent: TOB): boolean;
    function IsBlocNoteVide(BlocNote: String): boolean;
    procedure MnDeplierClick (Sender : TObject);
    procedure MnRefermerClick (Sender : TObject);
    procedure DeplierLigne(TOBL: TOB; Open: boolean = true);
    function GetLigne: integer;
    procedure PositionneDeciFournisseur(const Value: String);
    procedure PositionneDeciPrix(const Value: double);
    procedure FlipThisLigne (Ligne : integer);
    procedure PositionneSelectionBranche (TOBL : TOB);
    procedure PositionneSelectBas(TOBL: TOB; Etat: boolean);
    procedure DeselectionneUp(TOBL: TOB);
    procedure DeSelectionne(TOBL: TOB);
    procedure SelectionneTout(Status: boolean);
    procedure PositionneSelectUp(TOBL: TOB);
    procedure PositionneOptions;
    procedure SetmenuOptions(Status: boolean);
    procedure MnOptionsClick(Sender: TObject);
    procedure AddlesZoneOptions(TOBO: TOB);
    function IsSelectionne(TOBCurr: TOB): boolean;
    procedure MultipleAffecteFournisseur (LeFournisseur : string);
    procedure AffecteFournisseur (TOBD: TOB;LeFournisseur: string);
    procedure MultipleAffecteDepot (LeDepot : string);
    procedure MultipleChangeOptions(LivraisonChantier, GereStock: string);
    procedure AffecteDepot (TOBD: TOB;LeDepot: string);
    procedure recalculelesNiveaux(TOBI: TOB);
    procedure ReinitNiveauInf(TOBI: tob);
    procedure SetChangeOptions(TOBD: TOB; LivraisonChantier,GereStock: string);
    procedure setcumul(TOBPere, TOBdetail: tob);
    procedure SetChangementOption(TOBD: TOB; LivChantier,GereStock: string; Position : integer);
    procedure NettoieApresModif(TOBD: TOB; var Indice: integer);
    procedure FlipOverThisLigne (Ligne : integer);
    procedure DEPOTCHANGE (Sender : Tobject);
    procedure SetDepot(TOBL: TOB; DEPOT: string);
    function GetPhysique(Article, Depot : string): double;
    function TransformeRemises(REMISES: string): string;
    function CalculPrixNet(PRIXACHATBRUT: double; REMISES: string): double; overload;

  public
    property TheTOBdecisionnel : TOB read TOBDecisionnel;
    property LaLigne : integer read GetLigne;
    property TheConsultFou : TGestConsultFou read CONSULTFOU;
    property TheSelectFournisseur : String write PositionneDeciFournisseur;
    property TheSelectPrix : double write PositionneDeciPrix;
    function FindSelected (TOBCurr : TOB) : TOB;
    //
  end ;


function GetSelectionne (TOBBrancheDecision : TOB) : TOB;
procedure Ouvredecisionnel(param : string);

Implementation
uses facture,FactOuvrage,BTUtilDecisionAch,FactUtil,TarifUtil
;


function GetSelectionne (TOBBrancheDecision : TOB) : TOB;
var TOBF : TOB;
    Indice : integer;
begin
  result := nil;
  if TOBBrancheDecision.GetValue('SELECT')='X' Then BEGIN result := TOBBrancheDecision; Exit; END;
  For Indice := 0 to TOBBrancheDecision.detail.count -1 do
  begin
    result := GetSelectionne(TOBBrancheDecision.detail[Indice]);
    if result <> nil then break;
  end;
end;

procedure Ouvredecisionnel(param : string);
begin
	AglLanceFiche ('BTP','BTDECISIONACH','','',Param);
end;

procedure TOF_DECISIONACHAT.ActiveEventGS;
begin
GS.OnCellEnter := GSCellEnter;
GS.OnCellExit := GSCellExit;
GS.OnRowEnter := GSRowEnter;
GS.OnRowExit := GSRowExit;
GS.OnMouseMove := GSMouseMove;
GS.OnMouseDown := GSMouseDown;
GS.OnDblClick := GSDblClick;
TFvierge(Ecran).OnKeyDown := FormKeyDown;

end;

procedure TOF_DECISIONACHAT.SetMenuEnabled (TOBL : TOB);
begin
	MnVoirStock.Enabled := (TOBL.GetValue('GA_TENUESTOCK') = 'X') AND (TOBL.GetValue('NIVEAU') > 2);
	MnVoirbesoin.Enabled := (TOBL.GetValue('BAD_REFGESCOM') <> '');
  //
	MnScinder.Enabled := (TOBL.GetValue('BAD_TENUESTOCK') = 'X') and
											 (TOBL.GetValue('NIVEAU') = 4) and
                       (TOBL.GetValue('BAD_SCINDEE')='-') and
  										 (TOBL.GetValue('BAD_REMPLACE')='-') and
  										 (TOBL.GetValue('BAD_NUMREMPLACE')=0);
  //
  Mnremplace.enabled := (TOBL.GetValue('NIVEAU')=4) and
  											(TOBL.GetValue('BAD_REFGESCOM')<>'') and
                        (TOBL.GetValue('BAD_SCINDEE')<>'X') and
  											(TOBL.GetValue('BAD_REMPLACE')='-') and
  											(TOBL.GetValue('BAD_NUMREMPLACE')=0);
  //
  AjoutLigne.enabled := TOBL.GEtValue('NIVEAU')=4;
  MnSep1.visible := (TOBL.GetValue('BAD_SCINDEE') = 'X');
  MnSupLigne.visible := (TOBL.GetValue('BAD_SCINDEE') = 'X') OR (TOBL.GetValue('BAD_REMPLACEANT')='X') OR
  											((TOBL.GetValue('BAD_REFGESCOM') = '') AND (TOBL.GetValue('NIVEAU')=4));
  Newbesoin.enabled := true;
end;

procedure TOF_DECISIONACHAT.AfficheInfos(Ligne: integer);
var TOBL : TOB;
		Cledoc : r_cledoc;
begin
  TOBL:=GetTOBDecisAch(TOBDecisionnel,Ligne) ; if TOBL=Nil then Exit ;
  SetMenuEnabled (TOBL);
  TProvenance.Caption := '';
  TLIBELLECHANTIER.caption := '';
  TLIBELLEART.Caption := '';
  TLIBELLEFOURNISSEUR.caption  := '';

  if not VarISNull (TOBL.GetValue('BAD_LIBELLE')) then
  begin
//  	TLIBELLEART.Caption := 'Désignation article '+TOBL.GetValue('GA_LIBELLE');
  	TLIBELLEART.Caption := 'Désignation article '+TOBL.GetValue('BAD_LIBELLE');
  end;
  if not VarIsNull (TOBL.GetValue('AFF_LIBELLE')) then TLIBELLECHANTIER.caption := 'Chantier ' + TOBL.GetValue('AFF_LIBELLE');
  if TOBL.GetValue('BAD_REFGESCOM') <> '' then
  begin
    DecodeRefPiece (TOBL.GetValue('BAD_REFGESCOM'),cledoc);
    TProvenance.caption := 'Provenance '+rechdom('GCNATUREPIECEG',cledoc.NaturePiece,false) + ' N° '+InttoStr(Cledoc.NumeroPiece);
//    +' Ligne '+InttoStr(cledoc.NumLigne);
  end;
  if not VarIsNull (TOBL.GetValue('T_LIBELLE')) then
  begin
  	TLIBELLEFOURNISSEUR.Caption := 'Fournisseur '+TOBL.GetValue('T_LIBELLE');
  end;

  QUANTITE_VTE.Value := TOBL.GetValue('BAD_QUANTITEVTE');
  UNITE_VTE.Caption := Rechdom('GCQUALUNITTOUS',TOBL.GetValue('BAD_QUALIFQTEVTE'),false);
  QUANTITE_ACH.Value := TOBL.GetValue('BAD_QUANTITEACH');
  UNITE_ACH.Caption := RechDom('GCQUALUNITTOUS',TOBL.GetValue('BAD_QUALIFQTEACH'),false);
  QUANTITE_STO.Value := TOBL.GetValue('BAD_QUANTITESTK');
  UNITE_STO.Caption := RechDom('GCQUALUNITTOUS',TOBL.GetValue('BAD_QUALIFQTESTO'),false);

  DesactiveEvents;
	BAD_DATELIVRAISON.Text := TOBL.GetValue('BAD_DATELIVRAISON');
  PRISEENCOMPTE.Checked := (TOBL.GEtVAlue('BAD_PRISENCOMPTE')='X');
  if TOBL.GEtVAlue('BAD_TENUESTOCK')='' then
  begin
  	GERESTOCK.State := cbGrayed;
  end else
  begin
  	GERESTOCK.Checked := (TOBL.GEtVAlue('BAD_TENUESTOCK')='X');
  end;
  TDEPOT.Enabled := (GERESTOCK.State = cbChecked) ;
  DEPOT.Enabled := (GERESTOCK.State = cbChecked) ;
  DEPOT.Value := TOBL.getValue('BAD_DEPOT');
  BVOIRSTOCK.Enabled := (GERESTOCK.State = cbChecked) ;
  LIVCHANTIER.Checked := (TOBL.GEtVAlue('BAD_LIVCHANTIER')='X');
  ASUPPRIMER.checked := (TOBL.GEtVAlue('BAD_SUPPRIME')='X');
  if (TOBL.GetValue('ARTICLEREMPLPAR') <> '') or (TOBL.GetValue('ARTICLEREMPLACE') <> '')  then
  begin
  	if TOBL.GetValue('ARTICLEREMPLPAR') <> '' then
    begin
    	ARTREMPLACE.Text := TOBL.GetValue('ARTICLEREMPLPAR');
  		THLabel(GetControl('TREPLACE')).Caption := 'Remplacé par';
    end else
    begin
    	ARTREMPLACE.Text := TOBL.GetValue('ARTICLEREMPLACE');
  		THLabel(GetControl('TREPLACE')).Caption := 'Remplace';
    end;
    ARTREMPLACE.Visible := True;
		BFINDASSOCIATION.visible := true;
    THLabel(GetControl('TREPLACE')).Visible  := True;
  end else
  begin
    ARTREMPLACE.Visible := false;
    THLabel(GetControl('TREPLACE')).Visible  := false;
		BFINDASSOCIATION.visible := false;
  end;
  PRIXACHATBRUT.Value := TOBL.GetValue('BAD_PRIXACH');
  PRIXACHATNET.Value := TOBL.GetValue('BAD_PRIXACHNET');
  REMISES.text := TOBL.GetValue('BAD_CALCULREMISE');
  StringToRich(BLOCNOTE, TOBL.GetValue('BAD_BLOCNOTE'));
  ActiveEvents;
end;

procedure TOF_DECISIONACHAT.CreateControls;
begin
  CONSULTFOU := TGestConsultFou.create (ecran);
  ImStateConsFou := TImageList.create (Tform(ecran));
  ImStateConsFou.height := 16;
  ImStateConsFou.Width := 16;
  ImTypeArticle := TImageList.create (Tform(ecran));
  ImTypeArticle.height := 16;
  ImTypeArticle.Width := 16;
  ImOpenClose := TImageList.create (Tform(ecran));
  ImOpenClose.height := 12;
  ImOpenClose.Width := 12;
  FenetreSup := TToolWindow97.Create (Ecran)  ;
  FenetreSup.Parent := Ecran;
  FenetreSup.ClientWidth := THPanel(GetControl('PINFOSTARIF')).Width+10;
  FenetreSup.ClientHeight := THPanel(GetControl('PINFOSTARIF')).Height+10;
  FenetreSup.Top := THPanel(GetControl('PINFOSTARIF')).Top;
  FenetreSup.Left := THPanel(GetControl('PINFOSTARIF')).Left;
  //
  THPanel(GetControl('PINFOSTARIF')).parent := FenetreSup;
  THPanel(GetControl('PINFOSTARIF')).Align := Alclient;
  //
  fenetreSup.Caption := TraduireMemoire('Information de la ligne');
  FenetreSup.visible := false;
  FindDialog := TFindDialog.Create (Ecran);
  FindDialog.OnFind := FindDialogFind;
  //
  fInternalWindow := TToolWindow97.create(ecran);
  fInternalWindow.Parent := ecran;
  fInternalWindow.Visible := false;
  fTexte := THRichEditOLE.Create (fInternalWindow);
  ftexte.Parent := fInternalWindow;
  ftexte.text := '';
  ftexte.Visible := false;
  //
end;

procedure TOF_DECISIONACHAT.ChargeControls;
var UneImage : TImage;
begin
	UneImage := TImage(getControl('IMG_PREST'));
  UneIMage.Transparent := true;
  ImTypeArticle.Add ( TBitMap(UneImage.Picture.Bitmap),nil);
	UneImage := TImage(getControl('IMG_ARTICLE'));
  UneIMage.Transparent := true;
  ImTypeArticle.Add ( TBitMap(UneImage.Picture.Bitmap),nil);
	UneImage := TImage(getControl('IMG_PRIXPOSE'));
  UneIMage.Transparent := true;
  ImTypeArticle.Add ( TBitMap(UneImage.Picture.Bitmap),nil);
  //
	UneImage := TImage(getControl('IMG_PLUS'));
  ImOpenClose.add (TBitMap(UneImage.Picture.Bitmap),nil);
	UneImage := TImage(getControl('IMG_MOINS'));
  ImOpenClose.add (TBitMap(UneImage.Picture.Bitmap),nil);
	UneImage := TImage(getControl('IMG_FAUCUN'));
  if UneImage <> nil then
  begin
    UneIMage.Transparent := true;
    ImStateConsFou.Add ( TBitMap(UneImage.Picture.Bitmap),nil);
    UneImage := TImage(getControl('IMG_FAPRES'));
    UneIMage.Transparent := true;
    ImStateConsFou.Add ( TBitMap(UneImage.Picture.Bitmap),nil);
    UneImage := TImage(getControl('IMG_FICI'));
    UneIMage.Transparent := true;
    ImStateConsFou.Add ( TBitMap(UneImage.Picture.Bitmap),nil);
    UneImage := TImage(getControl('IMG_FAVANT'));
    UneIMage.Transparent := true;
    ImStateConsFou.Add ( TBitMap(UneImage.Picture.Bitmap),nil);
  end else
  begin
    // prise en compte du fait que la consultation fournisseur n'a pas ete mise en place
    OkConsultFou := false;
  end;

end;

procedure TOF_DECISIONACHAT.ChargeDecisionnel;
var QQ : TQuery;
		req : string;
    TOBInterm : TOB;
begin
	MaxRemplace := 0;
	TOBInterm := TOB.Create ('LE DETAIL',nil,-1);
  TRY
    Req := 'SELECT * FROM DECISIONACH WHERE BAE_NUMERO='+NumDecisionnel;
    QQ := OpenSql (req,true,-1,'',true);
    TOBDecisionnel.SelectDB ('',QQ);
    ferme (QQ);
    Req := 'SELECT TIE.T_LIBELLE,AFF.AFF_LIBELLE,ART.GA_TENUESTOCK,ART.GA_LIBELLE,ACH.* FROM DECISIONACHLIG ACH '+
           'LEFT JOIN ARTICLE ART ON ART.GA_ARTICLE=ACH.BAD_ARTICLE '+
           'LEFT JOIN AFFAIRE AFF ON AFF.AFF_AFFAIRE=ACH.BAD_AFFAIRE '+
           'LEFT JOIN TIERS TIE ON (TIE.T_TIERS=ACH.BAD_FOURNISSEUR) AND (TIE.T_NATUREAUXI="FOU") '+
           'WHERE ACH.BAD_NUMERO='+NumDecisionnel +' ORDER BY ACH.BAD_NUMN1,ACH.BAD_NUMN2,ACH.BAD_NUMN3,ACH.BAD_NUMN4';
    QQ := OpenSql (req,true,-1,'',true);
    TOBInterm.loaddetailDB ('DECISIONACHLIG','','',QQ,false,true);
    ferme (QQ);
    if TOBInterm.detail.count > 0 then
    begin
      AddlesChampsSup(TOBInterm,MaxRemplace);
      SetStkPhysique (TOBInterm);
      ConstitueLaTOBGestion (TOBdecisionnel,TOBInterm);
      CumulStkPhysique (TOBdecisionnel);
    end;
  FINALLY
  	TOBInterm.free;
  END;
end;


procedure TOF_DECISIONACHAT.DefiniGrilleSaisie;
var Numero : integer;
		TheListe,Unchamp  : string;
begin
  TheListe := ListeSaisie;
  numero := 0;
  repeat
  	UnChamp := READTOKENST (TheListe);
    if UnChamp <> '' then
    begin
      if UnChamp = 'OPEN' then GS_OPEN := numero else
      if UnChamp = 'OPEN1' then GS_OPEN1 := numero else
      if UnChamp = 'OPEN2' then GS_OPEN2 := numero else
      if UnChamp = 'BAD_TYPEARTICLE' then GS_TYPEARTICLE := Numero else
      if UnChamp = 'BAD_CODEARTICLE' then GS_CODEARTICLE := Numero else
      if UnChamp = 'BAD_AFFAIRE' then GS_AFFAIRE := Numero else
      if UnChamp = 'BAD_QUANTITESTK' then GS_QUANTITESTK := Numero else
      if UnChamp = 'BAD_DEPOT' then GS_DEPOT := Numero else
      if UnChamp = 'BAD_QUALIFQTESTO' then GS_QUALIFQTESTK := Numero else
      if UnChamp = 'BAD_LIVCHANTIER' then GS_LIVCHANTIER := Numero else
      if UnChamp = 'BAD_PABASE' then GS_PABASE := Numero else
      if UnChamp = 'QTEPHY' then GS_QTEPHY := Numero else
      if UnChamp = 'BAD_FOURNISSEUR' then GS_FOURNISSEUR := Numero else
      if UnChamp = 'BAD_PRIXACHNET' then GS_PRIXACHNET := Numero else
      if UnChamp = 'BAD_PRISENCOMPTE' then GS_PRISENCOMPTE := Numero else
      if UnChamp = 'POSITIONCONSULT' then GS_POSITIONCONSULT := Numero else
      if UnChamp = 'BAD_TENUESTOCK' then GS_TENUESTOCK := Numero else
      if UnChamp = 'BAD_NUMREMPLACE' then GS_NUMREMPLACE := Numero;
      inc(numero);
    end;

  until UnChamp='';

  GS.ColCount := Numero;

  // Affichage en + ou en -
  GS.cells[GS_OPEN,0] := '';
  GS.ColWidths [GS_OPEN] := 16;

  GS.cells[GS_OPEN1,0] := '';
  GS.ColWidths [GS_OPEN1] := 16;

  GS.cells[GS_OPEN2,0] := '';
  GS.ColWidths [GS_OPEN2] := 16;

  // Affichage du type d'article
  GS.cells[GS_TYPEARTICLE,0] := 'Typ.';
  GS.ColWidths [GS_TYPEARTICLE] := 20;

  // Affichage du code article
  GS.cells[GS_CODEARTICLE,0] := 'Code';
  GS.ColWidths [GS_CODEARTICLE] := 120;
	GS.ColAligns[GS_CODEARTICLE]:=taLeftJustify;

  // affichage du code chantier
  GS.cells[GS_AFFAIRE,0] := 'Chantier';
  GS.ColWidths [GS_AFFAIRE] := 120;
  GS.Collengths [GS_AFFAIRE] := 120;
	GS.ColAligns[GS_AFFAIRE]:=taRightJustify;

  // Affichage de la quantité en UA
  GS.cells[GS_QUANTITESTK,0] := 'Qté';
  GS.ColWidths [GS_QUANTITESTK] := 80;
  GS.Collengths [GS_QUANTITESTK] := 80;
	GS.ColAligns[GS_QUANTITESTK]:=taRightJustify;
  GS.ColFormats[GS_QUANTITESTK]:='#,##0.00';
  GS.ColTypes[GS_QUANTITESTK]:='R';

	// Afffichage du dépot
  GS.cells[GS_DEPOT,0] := 'Depot';
  GS.ColWidths [GS_DEPOT] := 100;
	GS.ColAligns[GS_DEPOT]:=taCenter;

  // Affichage du qualifiant achat
  GS.cells[GS_QUALIFQTESTK,0] := 'US';
  GS.ColWidths [GS_QUALIFQTESTK] := 30;
	GS.ColAligns[GS_QUALIFQTESTK]:=taCenter;

  // Livre sur chantier
  GS.cells[GS_LIVCHANTIER,0] := 'LC';
  GS.ColWidths [GS_LIVCHANTIER] := 18;
	GS.ColAligns[GS_LIVCHANTIER]:=taCenter;
  GS.ColTypes[GS_LIVCHANTIER]   := 'B';                        //B = Boolean
  GS.ColFormats[GS_LIVCHANTIER] := IntToStr(integer(csCoche)); //Affichage mouette ou rien

  // Affichage du stock physique
  GS.cells[GS_QTEPHY,0] := 'Stock';
  GS.ColWidths [GS_QTEPHY] := 40;
	GS.ColAligns[GS_QTEPHY]:=taRightJustify;
  GS.ColFormats[GS_QTEPHY]:='#,##0.00';
  GS.ColTypes[GS_QTEPHY]:='R';

  // Affichage du prix budgété
  GS.cells[GS_PABASE,0] := 'PA Budget';
  GS.ColWidths [GS_PABASE] := 70;
  GS.Collengths [GS_PABASE] := 70;
	GS.ColAligns[GS_PABASE]:=taRightJustify;
  GS.ColFormats[GS_PABASE]:='#,##0.00;; ;';
  GS.ColTypes[GS_PABASE]:='R';

  // affichage du fournisseur
  GS.cells[GS_FOURNISSEUR,0] := 'Fournisseur';
  GS.ColWidths [GS_FOURNISSEUR] := 80;
  GS.Collengths [GS_FOURNISSEUR] := 80;
	GS.ColAligns[GS_FOURNISSEUR]:=taLeftJustify;

  // prix achat net
  GS.cells[GS_PRIXACHNET,0] := 'PUA';
  GS.ColWidths [GS_PRIXACHNET] := 70;
  GS.Collengths [GS_PRIXACHNET] := 70;
	GS.ColAligns[GS_PRIXACHNET]:=taRightJustify;
  GS.ColFormats[GS_PRIXACHNET]:='#,##0.00';
  GS.ColTypes[GS_PRIXACHNET]:='R';

  // pris en compte
  GS.cells[GS_PRISENCOMPTE,0] := 'OK';
  GS.ColWidths [GS_PRISENCOMPTE] := 18;
	GS.ColAligns[GS_PRISENCOMPTE]:=taLeftJustify;
  GS.ColTypes[GS_PRISENCOMPTE]   := 'B';                        //B = Boolean
  GS.ColFormats[GS_PRISENCOMPTE] := IntToStr(integer(csCoche)); //Affichage mouette ou rien

  // Tenue en stock
  GS.cells[GS_TENUESTOCK,0] := 'Stk';
  GS.ColWidths [GS_TENUESTOCK] := 18;
	GS.ColAligns[GS_TENUESTOCK]:=taCenter;
  GS.ColTypes[GS_TENUESTOCK]   := ' ';                        //B = Boolean
//  GS.ColFormats[GS_TENUESTOCK] := IntToStr(integer(csCoche)); //Affichage mouette ou rien
  // Consultation fournisseur
  GS.cells[GS_POSITIONCONSULT,0] := 'Cons.';
  GS.ColWidths [GS_POSITIONCONSULT] := 20;
	GS.ColAligns[GS_POSITIONCONSULT]:=taCenter;
  GS.ColTypes[GS_POSITIONCONSULT]   := ' ';
  // N° de remplacement
  GS.cells[GS_NUMREMPLACE,0] := 'N° rempl.';
  GS.ColWidths [GS_NUMREMPLACE] := 60;
	GS.ColAligns[GS_NUMREMPLACE]:=taRightjustify;
  GS.ColFormats[GS_NUMREMPLACE]:='#,##0';

end;


procedure TOF_DECISIONACHAT.DefiniNiveauVisible (Niveau : integer);
begin
  GS.colwidths [GS_OPEN] := 16;
  GS.colwidths [GS_OPEN1] := 16;
  GS.colwidths [GS_OPEN2] := 16;
  GS.colwidths [GS_TYPEARTICLE] := 20;
  GS.colwidths [GS_CODEARTICLE] := 120;
  GS.colwidths [GS_AFFAIRE] := 0;

  GS.colwidths [GS_QUANTITESTK] := 80;
  GS.colwidths [GS_QUALIFQTESTK] := 30;
  GS.colwidths [GS_QTEPHY] := 40;
  GS.colwidths [GS_PABASE] := 70;
  GS.colwidths [GS_FOURNISSEUR] := 80;
  GS.colwidths [GS_PRIXACHNET] := 70;
  GS.colwidths [GS_PRISENCOMPTE] := 18;
//  GS.colwidths [GS_TENUESTOCK] := 0;
	GS.colwidths [GS_TENUESTOCK] := 18;
  GS.colwidths [GS_NUMREMPLACE] := 0;

	if Niveau > 1 then // a partie du niveau Nature
  begin
		GS.colwidths [GS_LIVCHANTIER] := 18;
		if Niveau > 2 then // a partir du niveau Articles
    BEGIN
    	GS.colwidths [GS_DEPOT] := 100;
      if NIveau > 3 then
      begin
      	GS.colwidths [GS_AFFAIRE] := 120;
//  			GS.colwidths [GS_TENUESTOCK] := 18;
  			GS.colwidths [GS_NUMREMPLACE] := 60;
      end else
      begin
  			GS.colwidths [GS_OPEN1] := 18;
      	GS.colwidths [GS_AFFAIRE] := 0;
  			GS.colwidths [GS_NUMREMPLACE] := 0;
      end;
    END ELSE
    BEGIN
  		GS.colwidths [GS_OPEN2] := 18;
    	GS.colwidths [GS_DEPOT] := 0;
    END;
	end else
  begin
  	GS.colwidths [GS_OPEN1] := 18;
  	GS.colwidths [GS_AFFAIRE] := 0;
		GS.colwidths [GS_LIVCHANTIER] := 0;
    GS.colwidths [GS_DEPOT] := 0;
    GS.colwidths [GS_NUMREMPLACE] := 0;
  end;
end;

procedure TOF_DECISIONACHAT.DesactiveEventGS;
begin
GS.OnCellEnter := nil;
GS.OnCellExit := nil;
GS.OnRowEnter := nil;
GS.OnRowExit := nil;
end;

procedure TOF_DECISIONACHAT.FreeControls;
begin
  CONSULTFOU.free;
  ImStateConsFou.free;
  ImTypeArticle.free;
  ImOpenClose.free;
  FenetreSup.free;
	FindDialog.Free;
  ftexte.free;
  fInternalWindow.free;
end;

procedure TOF_DECISIONACHAT.GetCellCanvas(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var Niveau : integer;
    TOBL : TOB;
    ARect : TRect ;
    PrisEncompte,Asupprimer : boolean;
begin
  if ACol<0 then Exit ;
  if (Arow < GS.FixedRows) or (Arow > GS.Rowcount) then exit;
  ARect:=GS.CellRect(ACol,ARow) ;
  TOBL:=GetTOBDecisAch(TOBDecisionnel,ARow) ; if TOBL=Nil then Exit ;
  Canvas.Pen.Style:=psSolid ; Canvas.Pen.Color:=clgray ;
  Niveau := TOBL.GetValue('NIVEAU');
  PrisEncompte := (TOBL.GetValue('BAD_PRISENCOMPTE')='X');
  Asupprimer := (TOBL.GetValue('BAD_SUPPRIME')='X');
  if Asupprimer then BEGIN canvas.font.Style := [fsStrikeOut];canvas.font.Color := clblack END else
  if PrisEncompte then canvas.font.Color := clblack else canvas.font.Color := clGray;
  if Niveau = 1 then // Article
  BEGIN
    Canvas.Brush.Style:=bsSolid ;
    canvas.Brush.Color := $F2C7B6;
    Canvas.Font.Style:=Canvas.Font.Style+[fsItalic];
    canvas.DrawFocusRect (Arect);
  END else
  if Niveau = 2 then // Livraison chantier ou non
  BEGIN
  	if Acol > GS_CODEARTICLE Then
    BEGIN
      Canvas.Brush.Style:=bsSolid ;
      canvas.Brush.Color := $A2CAF2;
      Canvas.Font.Style:=Canvas.Font.Style+[fsItalic];
      canvas.DrawFocusRect (Arect);
    END;
  END else
  if Niveau = 3 then // Depot
  BEGIN
    if Acol > GS_LIVCHANTIER Then
    BEGIN
      Canvas.Brush.Style:=bsSolid ;
      canvas.Brush.Color := $F0C2DE;
      Canvas.Font.Style:=Canvas.Font.Style+[fsItalic];
      canvas.DrawFocusRect (Arect);
    END;
  END else
  if Niveau = 4 then // Article
  BEGIN
  	if TOBL.GetValue('BAD_REMPLACE')='X' then Canvas.Font.Style := Canvas.Font.Style + [fsStrikeOut];
   	if Acol > GS_DEPOT Then
    BEGIN
    	if (Action = TaConsult) and (gdselected in Astate) then
      begin
        Canvas.Brush.Style:=bsSolid ;
        canvas.font.color := clWhite;
     end;
    END;
  END;

  if (Acol = GS_TENUESTOCK) and (Arow >= GS.fixedRows) then
  begin
    if TOBL.GetValue('BAD_TENUESTOCK') = '' then
    begin
      Canvas.font.Color := clred;
    end;
  end;

  if (TOBL.GetValue('SELECT') = 'X')and(Arow >= GS.fixedRows) then
  begin
    Canvas.Brush.Style:=bsSolid ;
    canvas.Brush.Color := clHighlight;
    Canvas.Font.Style:=Canvas.Font.Style+[fsItalic];
    canvas.font.color := clHighlightText;
    canvas.DrawFocusRect (Arect);
  end;
end;

function TOF_DECISIONACHAT.GetRow : integer;
begin
  result := StoredRow;
end;

procedure TOF_DECISIONACHAT.GetControls;
begin
	CHOIXFOUR := TPopupMenu (GetCOntrol('CHOIXFOUR'));
  MnCatalogue := TmenuItem(GetCOntrol('Mncatalogue'));
  MnFournisseur := TmenuItem (GetCOntrol('MnFournisseur'));
  GOptions := TGroupBox (GetCOntrol('GOPTIONS'));

  GS := THgrid(getControl('GS'));
  TLIBELLEART := THlabel(GetControl('TLIBELLEART'));
  TLIBELLECHANTIER := THlabel (GetControl('TLIBELLECHANTIER'));
  TPROVENANCE := THlabel(GetControl('TPROVENANCE'));
  TLIBELLEFOURNISSEUR := THLabel (GetControl('TLIBELLEFOURNISSEUR'));
  QUANTITE_VTE := THNumEdit(GetCOntrol('QUANTITE_VTE'));
  QUANTITE_STO := THNumEdit(GetCOntrol('QUANTITE_STO'));
  QUANTITE_ACH := THNumEdit(GetCOntrol('QUANTITE_ACH'));
  BAD_DATELIVRAISON := THEdit(GetControl('BAD_DATELIVRAISON'));
  PRISEENCOMPTE := TCheckBox(GetCOntrol('PRISENCOMPTE'));
  ASUPPRIMER := TCheckBox(GetCOntrol('ASUPPRIME'));
  GERESTOCK := TCheckBox(GetCOntrol('GERESTOCK'));
	LIVCHANTIER := TCheckBox(GetCOntrol('LIVCHANTIER'));
  BBouton := TToolBarButton97(GetControl('BBUTTON'));
  UNITE_VTE := THLabel (GetControl('UNITE_VTE'));
  UNITE_ACH := THLabel (GetControl('UNITE_ACH'));
  UNITE_STO := THLabel (GetControl('UNITE_STO'));
  BDelete := TToolBarButton97(GetControl('BDelete'));
  BImprimer := TToolBarButton97(GetControl('BImprimer'));
  MnVoirArticle := TmenuItem(GetControl('MnVoirArticle'));
  MnVoirStock  := TmenuItem(GetControl('MnVoirStock'));
  MnVoirBesoin  := TmenuItem(GetControl('MnVoirBesoin'));
  NewBesoin  := TmenuItem(GetControl('NewBesoin'));
  AjoutLigne := TmenuItem(GetControl('AjoutLigne'));
  MnScinder := TmenuItem(GetControl('MnScinder'));
  MnRemplace := TmenuItem(GetControl('MnRemplace'));
  MnDeplier := TmenuItem(GetControl('MnDeplier'));
  MnRefermer := TmenuItem(GetControl('MnRefermer'));
  MnSep1:= TmenuItem(GetControl('MnSep1'));
  MnSupLigne:= TmenuItem(GetControl('MnSupLigne'));
  PINFOSTARIF := THPanel(GetControl('PINFOSTARIF'));
  PRIXACHATBRUT := THNumEdit(GetCOntrol('PRIXACHATBRUT'));
  PRIXACHATNET := THNumEdit(GetCOntrol('PRIXACHATNET'));
  BLOCNOTE := THRichEditOLE (GetCOntrol('BLOCNOTE'));
  ARTREMPLACE := THEdit(GetControl('ARTREMPLACE'));
  REMISES := THEdit(GetControl('REMISES'));
  BViewDetail := TToolBarButton97(getControl('BViewDetail'));
	BRECHARTICLE := TToolBarButton97(getControl('BRECHARTICLE'));
  BFINDASSOCIATION := TToolBarButton97(getControl('BFINDASSOCIATION'));
  MnOptions := TmenuItem(GetControl('MnOptions'));
  BOPTIONMULTISEL := TToolbarButton97 (GetCOntrol('BOPTIONMULTISEL'));
  TDEPOT := THlabel (GetCOntrol('LDEPOT'));
  DEPOT := THValComboBox (getControl('DEPOT'));
  BVOIRSTOCK := TToolbarButton97 (getControl('BVOIRSTOCK'));
end;

procedure TOF_DECISIONACHAT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACHAT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACHAT.OnUpdate ;
begin
  Inherited ;
  if TRANSACTIONS (MajDecisionnel,0) <> OeOK then
  begin
  	PGiBox (TraduireMemoire('Une erreur c''est produite durant la mise à jour..'),ecran.caption);
//    TFVierge(ecran).
  end;
end ;

procedure TOF_DECISIONACHAT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACHAT.OnArgument (S : String ) ;
begin
  Inherited ;
  LastSelect:=0;
  OkConsultFou := True;
  CreateControls;
  ChargeControls;
  NiveauMax := 1;
  NbCols := 13;
  if OkConsultFou then
  begin
    ListeSaisie := 'OPEN;BAD_TYPEARTICLE;BAD_CODEARTICLE;OPEN1;BAD_LIVCHANTIER;OPEN2;BAD_DEPOT;BAD_AFFAIRE;BAD_TENUESTOCK;BAD_QUANTITESTK;BAD_QUALIFQTESTO;QTEPHY;BAD_FOURNISSEUR;BAD_PABASE;BAD_PRIXACHNET;BAD_PRISENCOMPTE;POSITIONCONSULT;BAD_NUMREMPLACE';
  end else
  begin
    CONSULTFOU.NonUtilisable := true;
    ListeSaisie := 'OPEN;BAD_TYPEARTICLE;BAD_CODEARTICLE;OPEN1;BAD_LIVCHANTIER;OPEN2;BAD_DEPOT;BAD_AFFAIRE;BAD_TENUESTOCK;BAD_QUANTITESTK;BAD_QUALIFQTESTO;QTEPHY;BAD_FOURNISSEUR;BAD_PABASE;BAD_PRIXACHNET;BAD_PRISENCOMPTE;BAD_NUMREMPLACE';
  end;
  NbLigne := 1;
  TOBDecisionnel := TOB.create ('DECISIONACH',nil,-1);
  GetControls;  // gestion des zones de la fiche
  SetInfos (S); // recupere le numero de decisionnel achat
  SetEcran;
  ChargeDecisionnel; // chargement de la TOB
  //
  CONSULTFOU.NumDecisionnel := NumDecisionnel;
  //
  IndiceGrilleInit (TOBdecisionnel,NiveauMax,NbLigne);
  Inc(NiveauMax);
  DefiniGrilleSaisie;
  DefiniNiveauVisible (NiveauMax);
	AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,NiveauMax,NbLigne);
  PositionneLigne (1);
  if MnOptions <> nil then MnOptions.enabled := False;
  SetEvents;
  Lastrow := 0;
end ;

procedure TOF_DECISIONACHAT.OnClose ;
begin
  Inherited ;
	FreeControls;
  TOBDecisionnel.free;
end ;

procedure TOF_DECISIONACHAT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACHAT.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_DECISIONACHAT.PositionneLigne (Arow:integer;Acol : integer=-1);
var
    Cancel : boolean;
    Synchro : boolean;
begin
  Synchro := GS.synenabled;
  GS.synEnabled := false;
  GS.CacheEdit;
  DesactiveEventGS;
  Arow := Arow;
  Acol := GS_AFFAIRE;
  GS.Row := Arow;
  GS.Col := Acol;
  GSCellEnter (self,Acol,Arow,cancel);
  GS.ElipsisButton := (Acol=GS_FOURNISSEUR) Or (ACol = GS_AFFAIRE);
  CellCur:=GS.Cells[ACol,ARow] ;

  GS.row := ARow; GS.Col := Acol;
  GS.SynEnabled := synchro;
  GS.MontreEdit;
  AfficheInfos(Gs.Row);
  SetEnabled (GS.row);
	ActiveEventGS;
end;

procedure TOF_DECISIONACHAT.PostDrawCell(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var ARect: TRect;
    Numgraph : integer;
    Chaine : string;
    PosG : integer;

begin
	if (Arow < GS.FixedRows) or (Arow > GS.Rowcount) then exit;
//  if LastRow <> Arow then
//  begin
    TOBLAffiche  := GetTOBDecisAch(TOBdecisionnel, ARow);
    if TOBLAffiche = nil then Exit;
    LastRow := Arow;
    NiveauAffiche := TOBLAffiche.GEtValue('NIVEAU');
//  end;
  if TOBLAffiche = nil then exit;
  ARect := GS.CellRect(ACol, ARow);
  if (Acol = GS_OPEN) and (Arow >= GS.fixedRows) then
  begin
    Canvas.FillRect(ARect);
    if NiveauAffiche<>1 then exit;
    NumGraph := RecupTypeGraphOpenClose(TOBLAffiche);
    if NumGraph >= 0 then
    begin
      ImOpenClose.DrawingStyle := dsTransparent;
      ImOpenClose.Draw(CanVas, ARect.left, ARect.top, NumGraph);
    end;
  end else
  if (Acol = GS_OPEN1) and (Arow >= GS.fixedRows) then
  begin
    Canvas.FillRect(ARect);
    if NiveauAffiche <>2 then exit;
    NumGraph := RecupTypeGraphOpenClose(TOBLAffiche);
    if NumGraph >= 0 then
    begin
      ImOpenClose.DrawingStyle := dsTransparent;
      ImOpenClose.Draw(CanVas, ARect.left, ARect.top, NumGraph);
    end;
  end else
  if (Acol = GS_OPEN2) and (Arow >= GS.fixedRows) then
  begin
    Canvas.FillRect(ARect);
    if NiveauAffiche<>3 then exit;
    NumGraph := RecupTypeGraphOpenClose(TOBLAffiche);
    if NumGraph >= 0 then
    begin
      ImOpenClose.DrawingStyle := dsTransparent;
      ImOpenClose.Draw(CanVas, ARect.left, ARect.top, NumGraph);
    end;
  end else
  if (Acol = GS_CODEARTICLE) and (Arow >= GS.fixedRows) then
  begin
  	if NiveauAffiche > 1 then
    begin
    	Canvas.FillRect(ARect);
    end;
  end else
  if (Acol = GS_LIVCHANTIER) and (Arow >= GS.fixedRows) then
  begin
  	if NiveauAffiche > 2 then
    begin
    	Canvas.FillRect(ARect);
    end;
  end else
  if (Acol = GS_DEPOT) and (Arow >= GS.fixedRows) then
  begin
  	if NiveauAffiche > 3 then
    begin
    	Canvas.FillRect(ARect);
    end else
    begin
      Canvas.FillRect(ARect);
      if TOBLAffiche.GetValue('BAD_DEPOT') = '' then exit;
      Chaine := copy(rechdom ('GCDEPOT',TOBLAffiche.GetValue('BAD_DEPOT'),false),1,17);
      PosG := Arect.Left + ((ARect.Right - ARect.left - canvas.textwidth(chaine)) div 2);
      canvas.TextOut (PosG,ARect.top + 1,Chaine);
    end;
  end else
  if (Acol = GS_TENUESTOCK) and (Arow >= GS.fixedRows) then
  begin
    Canvas.FillRect(ARect);
    if TOBLAffiche.GetValue('BAD_TENUESTOCK') = 'X' then
    begin
    	Chaine := 'R';
      Canvas.font.name := 'Wingdings 2';
    end else if TOBLAffiche.GetValue('BAD_TENUESTOCK') = '' then
    begin
    	Chaine := 'R';
      Canvas.font.name := 'Wingdings 2';
    end else
    begin
    	chaine := ' ';
    end;
    PosG := Arect.Left + ((ARect.Right - ARect.left - canvas.textwidth(chaine)) div 2);
    canvas.TextOut (PosG,ARect.top + 1,Chaine);
  end else
  if (Acol = GS_POSITIONCONSULT) and (Arow >= GS.fixedRows) then
  begin
    Canvas.FillRect(ARect);
    NumGraph := RecupTypeConsFou(TOBLAffiche);
    PosG := Arect.Left + ((ARect.Right - ARect.left - 16) div 2);
    if (NumGraph >= 0)  then
    begin
      ImStateConsFou.Draw(CanVas, PosG, ARect.top, NumGraph);
    end;
  end else
  if (Acol = GS_TYPEARTICLE) and (Arow >= GS.fixedRows) then
  begin
    Canvas.FillRect(ARect);
    NumGraph := RecupTypeGraph(TOBLAffiche);
    if (NumGraph >= 0) and (NiveauAffiche < 2) then
    begin
//      ImTypeArticle.DrawingStyle := dsTransparent;
      ImTypeArticle.Draw(CanVas, ARect.left, ARect.top, NumGraph);
    end;
  end else
  if (Acol = GS_AFFAIRE) and (Arow >= GS.fixedRows) then
  begin
    Canvas.Brush.Style := BsSolid;
    Canvas.FillRect(ARect);
    Chaine := BTPCodeAffaireAffiche (TOBLAffiche.GetValue('BAD_AFFAIRE'));
    if (Trim(Chaine) = '') and (NiveauAffiche=4) then Chaine := '<<Aucun>>';
    PosG := Arect.Left + ((ARect.Right - ARect.left - canvas.textwidth(chaine)) div 2);
    canvas.TextOut (PosG,ARect.top + 1,Chaine);
  end else
  if (Acol = GS_QUANTITESTK) and (Arow >= GS.fixedRows) then
  begin
  	if TOBLAffiche.GetValue('BAD_QUANTITESTK')=0 then
    begin
      Canvas.Brush.Style := BsSolid;
      Canvas.FillRect(ARect);
    end;
  end else
  if (Acol = GS_QTEPHY) and (Arow >= GS.fixedRows) then
  begin
  	if TOBLAffiche.GetValue('QTEPHY')=0 then
    begin
      Canvas.Brush.Style := BsSolid;
      Canvas.FillRect(ARect);
    end;
  end else
  if (Acol = GS_NUMREMPLACE) and (Arow >= GS.fixedRows) then
  begin
  	if TOBLAffiche.GetValue('BAD_NUMREMPLACE')=0 then
    begin
      Canvas.Brush.Style := BsSolid;
      Canvas.FillRect(ARect);
    end;
  end else
  if (Acol = GS_PRIXACHNET) and (Arow >= GS.fixedRows) then
  begin
  	if TOBLAffiche.GetValue('BAD_PRIXACHNET')=0 then
    begin
      Canvas.Brush.Style := BsSolid;
      Canvas.FillRect(ARect);
    end;
  end;

end;

procedure TOF_DECISIONACHAT.ChangeQteStk (Arow : integer; NewValue : double);
var FUV,FUS,FUA : double;
		TOBL : TOB;
    RatioStkVte,RatioStkAch,CoefUaUs : double;
    Niveau : integer;
begin
	RatioStkVte := 1;
	RatioStkAch := 1;
  TOBL := GetTOBDecisAch (TOBdecisionnel,Arow);
  Niveau := TOBL.GetValue('NIVEAU');
  TOBL.putValue('BAD_QUANTITESTK',NewValue);
	if Niveau = 4 then
  begin
    FUV := RatioMesure('PIE', TobL.GetValue('BAD_QUALIFQTEVTE')); if FUV = 0 then FUV := 1;
    FUS := RatioMesure('PIE', TobL.GetValue('BAD_QUALIFQTESTO')); if FUS = 0 then FUS := 1;
    FUA := RatioMesure('PIE', TobL.GetValue('BAD_QUALIFQTEACH')); if FUA = 0 then FUA := 1;
    CoefuaUs := TobL.GetValue('BAD_COEFUAUS');
    RatioStkVte := FUS/FUV;   // partant de UA pour arriver a UV
    RatioStkAch := FUS/FUA;   // Partant de UV pour arriver a US
  end;
(*
   else
  begin
  	RatioAchVte := QteVte/QteAch;
  	RatioVteStk := QteStk/QteVte;
  end;
*)
	TOBL.putValue('BAD_QUANTITEVTE',TOBL.GetValue('BAD_QUANTITESTK')*RatioStkVte);
	if CoefuaUs <>  0 then
  begin
    TOBL.putValue('BAD_QUANTITEACH',TOBL.GetValue('BAD_QUANTITESTK')/ CoefuaUs );
  end else
  begin
    TOBL.putValue('BAD_QUANTITEACH',TOBL.GetValue('BAD_QUANTITESTK')*RatioStkAch);
  end;
  TOBL.PutLigneGrid (GS,Arow,false,false,ListeSaisie);
  if Niveau = 4 then RecalculNiveauSup (TOBL);
  AfficheInfos(Arow);
end;

procedure TOF_DECISIONACHAT.QUANTITE_ACHChange (Sender : TObject);
var FUV,FUS,FUA : double;
		TOBL : TOB;
    {QteAch,QteStk,QteVte,}RatioAchVte,RatioVteStk : double;
    Niveau : integer;
    CoefuaUs : double;
begin
  TOBL := GetTOBDecisAch (TOBdecisionnel,GS.row);
  Niveau := TOBL.GetValue('NIVEAU');
  RatioAchVte := 1;
  RatioVteStk := 1;
  CoefuaUs := 0;
  //
  (*
  QteAch := TOBL.GetValue('BAD_QUANTITEACH');
  QteVte := TOBL.GetValue('BAD_QUANTITEVTE');
  QteStk := TOBL.GetValue('BAD_QUANTITESTK');
  *)
  //
  TOBL.putValue('BAD_QUANTITEACH',QUANTITE_ACH.Value);
	if Niveau = 4 then
  begin
    FUV := RatioMesure('PIE', TobL.GetValue('BAD_QUALIFQTEVTE')); if FUV = 0 then FUV := 1;
    FUS := RatioMesure('PIE', TobL.GetValue('BAD_QUALIFQTESTO')); if FUS = 0 then FUS := 1;
    FUA := RatioMesure('PIE', TobL.GetValue('BAD_QUALIFQTEACH')); if FUA = 0 then FUA := 1;
    CoefuaUs := TobL.GetValue('BAD_COEFUAUS');
    RatioAchVte := FUA/FUV;   // partant de UA pour arriver a UV
    RatioVteStk := FUV/FUS;   // Partant de UV pour arriver a US
  end;
(*
   else
  begin
  	RatioAchVte := QteVte/QteAch;
  	RatioVteStk := QteStk/QteVte;
  end;
*)
  if CoefuaUs <> 0 then
  begin
  	TOBL.putValue('BAD_QUANTITEVTE',Arrondi(TOBL.GetValue('BAD_QUANTITEACH')*CoefUaUs/FUS*FUV,V_PGI.okdecQ));
  	TOBL.putValue('BAD_QUANTITESTK',Arrondi(TOBL.GetValue('BAD_QUANTITEACH')*CoefUaUs,V_PGI.okdecQ));
  end else
  begin
  	TOBL.putValue('BAD_QUANTITEVTE',Arrondi(TOBL.GetValue('BAD_QUANTITEACH')*RatioAchVte,V_PGI.OKDecQ));
  	TOBL.putValue('BAD_QUANTITESTK',Arrondi(TOBL.GetValue('BAD_QUANTITEVTE')*RatioVTeStk,V_PGI.okdecQ));
  end;
  TOBL.PutLigneGrid (GS,GS.row,false,false,ListeSaisie);
	AfficheInfos(Gs.Row);

  if Niveau = 4 then RecalculNiveauSup (TOBL);
end;


procedure TOF_DECISIONACHAT.ReinitNiveauSup (TOBL: TOB);
var TOBP : TOB;
begin
  TOBP := TOBL.Parent;
  if TOBP.FieldExists ('BAD_QUANTITEVTE') then
  begin
    TOBP.PutValue ('BAD_QUANTITEVTE',0);
    TOBP.PutValue ('BAD_QUANTITESTK',0);
    TOBP.PutValue ('BAD_QUANTITEACH',0);
    TOBP.PutValue ('BAD_DATELIVRAISON',iDate2099);
  end;
  if TOBP.Parent <> nil then
  begin
    ReInitNiveauSup (TOBP);
  end;
end;

procedure TOF_DECISIONACHAT.CumulThisInPapou(TOBD,TOBP : TOB;Sens : string='+');
begin
	if Sens = '+' then
  begin
    TOBP.putValue('BAD_QUANTITEVTE',TOBP.GetValue('BAD_QUANTITEVTE')+TOBD.GetValue('BAD_QUANTITEVTE'));
    TOBP.putValue('BAD_QUANTITEACH',TOBP.GetValue('BAD_QUANTITEACH')+TOBD.GetValue('BAD_QUANTITEACH'));
    TOBP.putValue('BAD_QUANTITESTK',TOBP.GetValue('BAD_QUANTITESTK')+TOBD.GetValue('BAD_QUANTITESTK'));
    if TOBP.Getvalue ('BAD_DATELIVRAISON') > TOBD.Getvalue ('BAD_DATELIVRAISON') then
  			TOBP.putvalue ('BAD_DATELIVRAISON',TOBD.Getvalue ('BAD_DATELIVRAISON'));
  end else
  begin
    TOBP.putValue('BAD_QUANTITEVTE',TOBP.GetValue('BAD_QUANTITEVTE')-TOBD.GetValue('BAD_QUANTITEVTE'));
    TOBP.putValue('BAD_QUANTITEACH',TOBP.GetValue('BAD_QUANTITEACH')-TOBD.GetValue('BAD_QUANTITEACH'));
    TOBP.putValue('BAD_QUANTITESTK',TOBP.GetValue('BAD_QUANTITESTK')-TOBD.GetValue('BAD_QUANTITESTK'));
  end;
end;

procedure TOF_DECISIONACHAT.EnleveDesPeres (TOBL,TOBP : TOB);
begin
  if not TOBP.FieldExists('BAD_QUANTITEVTE') then exit;
  CumulThisInPapou (TOBL,TOBP,'-');
  if TOBP.Parent <> nil then EnleveDesPeres (TOBL,TOBP.Parent);
end;

procedure TOF_DECISIONACHAT.AjouteAuxPeres (TOBL,TOBP : TOB);
begin
  if not TOBP.FieldExists('BAD_QUANTITEVTE') then exit;
  CumulThisInPapou (TOBL,TOBP,'+');
  if TOBP.Parent <> nil then AjouteAuxPeres (TOBL,TOBP.Parent);
end;

procedure TOF_DECISIONACHAT.CalculAndAfficheNiveauSup (TOBL : TOB);
var TOBP,TOBD : TOB;
    Indice,Ligne : integer;
begin
  TOBP := TOBL.Parent;
  if not TOBP.FieldExists ('BAD_QUANTITEVTE') then exit;

  for Indice := 0 to TOBP.detail.count -1 do
  begin
     TOBD := TOBP.detail[Indice];
     CumulThisInPapou(TOBD,TOBP);
  end;

  Ligne := TOBP.GetValue('NUMAFF');
  TOBP.PutLigneGrid (GS,Ligne,false,false,ListeSaisie);

  if TOBP.Parent <> nil then CalculAndAfficheNiveauSup (TOBP);
end;

procedure TOF_DECISIONACHAT.RecalculNiveauSup ( TOBL : TOB);
begin
  ReInitNiveauSup (TOBL);
  CalculAndAfficheNiveauSup (TOBL);
end;

procedure TOF_DECISIONACHAT.GSCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
//
  if Action=taConsult then Exit ;
  ZoneSuivanteOuOk(ACol,ARow,Cancel) ;
  if cancel then exit;
//  if Acol < GS_HRS then GS.CacheEdit else GS.ShowEditor ;
  GS.ElipsisButton := (Acol=GS_FOURNISSEUR) Or (ACol = GS_AFFAIRE);
  CellCur:=GS.Cells[ACol,ARow] ;
end;

procedure TOF_DECISIONACHAT.GSCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);

  function ZoneSaisissable(TOBL : TOB; Niveau,Acol : integer) : boolean;
  begin
  	result := false;
    if (Niveau=4) and ((TOBL.getvalue('BAD_REMPLACE')='X') OR (TOBL.GETVAlue('BAD_SUPPRIME')='X')) then exit;
    if (Niveau <> 4) and ((Acol = GS_FOURNISSEUR) OR (Acol = GS_PRIXACHNET)) then begin Result:= True; exit; end;
    if (Niveau = 4) and ((Acol = GS_FOURNISSEUR) OR (Acol = GS_PRIXACHNET)) then begin Result:= True; exit; end;
  	if (Niveau = 4) then begin result := True; exit; end;
  end;

var TOBL : TOB;
begin
	if Action = TaConsult then exit;
  if CellCur = GS.cells[Acol,Arow] then exit;
  TOBL := GetTOBDecisAch (TOBdecisionnel,Arow);
  if not ZoneSaisissable (TOBL,TOBL.GetValue('NIVEAU'),Acol) then
  begin
    GS.Cells[Acol,Arow] := CellCur;
    exit;
  end;
  if Acol = GS_QUANTITESTK then
  begin
  	ChangeQteStk (Arow,Valeur(GS.Cells[Acol,Arow]));
  end else
  if Acol = GS_FOURNISSEUR then
  begin
  	ChangeFournisseur (Arow,GS.Cells[Acol,Arow],Cancel);
  end else
  if Acol = GS_PRIXACHNET then
  begin
  	ChangePrixNet (Arow,Valeur(GS.Cells[Acol,Arow]));
  end else
  if Acol = GS_AFFAIRE Then
  begin
	  changeAffaire (Arow,GS.cells[Acol,Arow],cancel);
  end;

end;

procedure TOF_DECISIONACHAT.GSElipsisClick (Sender : TObject);
var TOBL : TOB;
		Fournisseur,TheAffaire : string;
    cancel : boolean;
    Coords,Rect : Trect;
    PointD,PointF : Tpoint;
begin
  TOBL := GetTOBDecisAch (TOBdecisionnel,GS.row);
	//
  if GS.Col = GS_FOURNISSEUR then
  begin
   Coords := GS.CellRect (GS.col,GS.row);
   PointD := GS.ClientToScreen ( Coords.Topleft)  ;
   PointF := GS.ClientToScreen ( Coords.BottomRight )  ;

   CHOIXFOUR.Popup (pointF.X ,pointD.y+10);
(*
  	Fournisseur := LookUpFournisseur (GS.Cells [GS.Col,GS.row],TOBL.GEtValue('BAD_ARTICLE'),TOBL.GEtValue('GA_LIBELLE'),true);
    if Fournisseur <> '' then
    begin
    	if Fournisseur <> TOBL.GEtValue('BAD_FOURNISSEUR') then
      begin
        ChangeFournisseur (GS.row,Fournisseur,cancel);
				CellCur := GS.Cells [GS.Col,GS.row];
      end;
    end;
*)
  end else if GS.Col = GS_AFFAIRE then
  begin
  	TheAffaire := TOBL.getValue('BAD_AFFAIRE');
  	if RechercheAffaire (TheAffaire) then
    begin
    	changeAffaire (GS.row,TheAffaire,Cancel);
    	CellCur := GS.Cells [GS.Col,GS.row];
    end;
  end;
end;

procedure TOF_DECISIONACHAT.changeAffaire (Ligne : integer;TheAffaire : string; var Cancel : boolean);

  function GetLibelleChantier(theAffaire : string; var LibelleChantier : string) : boolean;
  var QQ : TQuery;
  begin
  	result := false;
    QQ := OpenSql ('SELECT AFF_LIBELLE FROM AFFAIRE WHERE AFF_AFFAIRE="'+TheAffaire+'"',true,-1,'',true);
    if not QQ.eof then
    begin
    	LibelleChantier := QQ.findField ('AFF_LIBELLE').AsString;
      result := true;
    end;
    ferme(QQ);
  end;

var A0,A1,A2,A3,A4,LibelleChantier : string;
    TOBL : TOB;
begin
	cancel := false;
  if TheAffaire <> '' then
  begin
    if not GetLibelleChantier(theAffaire,LibelleChantier) then
    begin
      PGIBox(TraduireMemoire('Ce Chantier n''existe pas'),ecran.caption);
      Cancel := true;
      exit;
    end;
  	BTPCodeAffaireDecoupe(TheAffaire, A0, A1, A2, A3, A4, taCreat, false);
  end else
  begin
    A0 := ''; A1 := ''; A2 := '';
    A3 := ''; A4 := ''; LibelleChantier := '';
  end;

	TOBL := GetTOBDecisAch (TOBdecisionnel,Ligne);
  TOBL.PutValue('BAD_AFFAIRE', TheAffaire);
  TOBL.PutValue('BAD_AFFAIRE1', A1);
  TOBL.PutValue('BAD_AFFAIRE2', A2);
  TOBL.PutValue('BAD_AFFAIRE3', A3);
  TOBL.PutValue('BAD_AVENANT', A4);
  TOBL.PutValue('AFF_LIBELLE',LibelleChantier );
  TOBL.PutLigneGrid (GS,GS.row,false,false,ListeSaisie);
  AfficheInfos(Ligne);
end;

procedure TOF_DECISIONACHAT.GSMouseMove(Sender: TObject;Shift: TShiftState; X, Y: Integer);
var Acol,Arow : integer;
begin
  GS.MouseToCell (X,Y,Acol,Arow);
	if (Arow < GS.fixedRows) or (Arow > GS.rowCount) then exit;
  SetRow (Arow);
  ShowButton (Arow);
end;

procedure TOF_DECISIONACHAT.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean;Chg: Boolean);
var TOBL : TOB;
begin
  if Ou > Gs.rowCount then exit;
  TOBL := GetTOBDecisAch (TOBdecisionnel,Ou);
  if TOBL = nil then exit;
  CurrentTOb := TOBL;
  if TOBL.GetValue('NIVEAU') = 0 then exit;
  //AffichageValor (TOBL);
  if TOBL.GetValue('NIVEAU') > MAXNiveauAff then exit;
  ShowButton (Ou);
	AfficheInfos(Gs.Row);
  CONSULTFOU.RowChange;
  SetEnabled (GS.row);
end;

procedure TOF_DECISIONACHAT.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean;Chg: Boolean);
begin
end;

procedure TOF_DECISIONACHAT.SetEvents;
begin
  MnCatalogue.onClick := MnCatalogueClick;
  MnFournisseur.OnClick := MnFournisseurClick;
  GS.OnMouseDown := GSMouSeDown;
  GS.PostDrawCell := PostDrawCell;
  GS.GetCellCanvas := GetCellCanvas;
  GS.OnRowEnter := GsRowEnter;
  GS.OnRowExit := GsRowExit;
  GS.OnMouseMove := GSMouseMove;
  GS.OnCellExit := GSCellExit;
  GS.OnElipsisClick := GSElipsisClick;
  BBouton.onclick := BBoutonClick;
  Bdelete.onclick := BdeleteClick;
  BImprimer.onclick := BImprimerClick;
  QUANTITE_ACH.OnExit := QUANTITE_ACHChange;
  MnVoirArticle.OnClick := VoirArticleClick;
  MnVoirStock.onClick := VoirStockClick;
  MnVoirBesoin.onClick  := VoirBesoinClick;
  NewBesoin.onClick := NewbesoinClick;
  AjoutLigne.onclick := AjoutligneClick;
  MnScinder.onclick := MnScinderClick;
  MnRemplace.onClick := MnRemplaceClick;
  MnDeplier.onClick := MnDeplierClick;
  MnRefermer.onClick := MnRefermerClick;
  MnSupLigne.OnClick := MnSupLigneClick;
  BAD_DATELIVRAISON.OnExit := BAD_DATELIVRAISONClick;
  REMISES.OnExit := REMISESChange;
  PRIXACHATBRUT.onExit := REMISESChange;
  PRIXACHATNET.OnEnter := PRIXACHATNETEnter;
  PRIXACHATNET.OnExit := PRIXACHATNETChange;
  BViewDetail.onclick := BViewDetailClick;
  FenetreSup.OnClose := FenetreSupClose;
	BRECHARTICLE.OnClick := BrechArticleClick;
  BFINDASSOCIATION.OnClick := BFindAssociationClick;
  if MnOptions <> nil then MnOptions.onClick := MnOptionsClick;
  if BOPTIONMULTISEL <> nil then BOPTIONMULTISEL.onClick := MnOptionsClick;
  ActiveEvents;
end;

procedure TOF_DECISIONACHAT.SetInfos (arguments : string);
var critere,ChampMul,ValMul : string;
		x : integer;
begin
  repeat
    Critere := uppercase(Trim(ReadTokenSt(Arguments)));
    if Critere <> '' then
    begin
      x := pos('=', Critere); 
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'NUMDECISIONNEL' then NumDecisionnel := ValMul else
        if (ChampMul = 'ACTION') and (ValMul='CONSULTATION') then
        begin
        	TFVierge(ecran).TypeAction := TaConsult;
          Action := Taconsult;
        end else
        begin
        	TFVierge(ecran).TypeAction := TaModif;
          Action := TaModif;
        end;
      end;
    end;
  until Critere = '';
  Bdelete.visible := (Action <> taConsult);
  BImprimer.visible := True;
end;

procedure TOF_DECISIONACHAT.SetRow(Arow: integer);
begin
  StoredRow := Arow;
end;

procedure TOF_DECISIONACHAT.SetEnabled(Arow: integer);
var TOBL : TOB;
begin
  TOBL:=GetTOBDecisAch(TOBDecisionnel,Arow) ; if TOBL=Nil then Exit ;
  GERESTOCK.Enabled := (TOBL.GEtVAlue('GA_TENUESTOCK')='X') AND (TOBL.GEtVAlue('BAD_SCINDEE')<>'X') and (Action <> TaConsult);
  DEPOT.Enabled := (GERESTOCK.state=cbChecked );
  TDEPOT.Enabled := (GERESTOCK.state=cbChecked );
  BVOIRSTOCK.enabled := (GERESTOCK.state=cbChecked );
  ASUPPRIMER.Enabled := (Action <> TaConsult);
  if TOBL.getValue('NIVEAU') = 4 then
  begin
//  	GERESTOCK.Enabled := (TOBL.GEtVAlue('GA_TENUESTOCK')='X') AND (TOBL.GEtVAlue('BAD_SCINDEE')<>'X') and (Action <> TaConsult);
    QUANTITE_ACH.Enabled := (Action <> TaConsult);
    if Action <> TAConsult then QUANTITE_ACH.Color := ClWindow
    											 else QUANTITE_ACH.Color := clInactiveCaption;
//    ASUPPRIMER.Enabled := (TOBL.GetValue('BAD_REFGESCOM')<>'') and (Action <> TaConsult);
		BAD_DATELIVRAISON.enabled := true;
		BAD_DATELIVRAISON.Color  := clWindow;
  	LIVCHANTIER.Enabled := (Action <> TaConsult);
  end else
  begin
    LIVCHANTIER.Enabled := false;
//  	GERESTOCK.enabled := false;
    QUANTITE_ACH.Enabled := false;
    QUANTITE_ACH.Color := clInactiveCaption;
//    ASUPPRIMER.Enabled := false;
		BAD_DATELIVRAISON.enabled := False;
		BAD_DATELIVRAISON.Color  := clInactiveCaption;
  end;
end;

procedure TOF_DECISIONACHAT.ShowButton (Ou : integer);
var Arect : Trect;
		TOBL : TOB;
    Niveau : integer;
begin
	if (Ou < GS.FixedRows) or (Ou > GS.Rowcount) then exit;
  TOBL := GetTOBDecisAch (TOBdecisionnel,Ou);
  if TOBL = nil then exit;
  Niveau := TOBL.GetValue('NIVEAU');
  if Niveau = 0 then exit;
  if Niveau >= MAXNiveauAff then exit;
  if TOBL = nil then exit;
  CurrentTOb := TOBL;
  if Niveau = 1 then ARect:=GS.CellRect(GS_OPEN,Ou) else
  if NIveau = 2 then ARect:=GS.CellRect(GS_OPEN1,Ou) else
  if NIveau = 3 then ARect:=GS.CellRect(GS_OPEN2,Ou);
  BBouton.ImageIndex := RecupTypeGraphOpenClose (TOBL);
  With Bbouton do
  Begin
    Opaque := false;
    Top := Arect.top;
    Left := Arect.Left;
    Width := Arect.Right - Arect.Left ;
    Height  := Arect.Bottom - Arect.Top;
    Parent := GS;
    Visible := true;
  end;
end;


procedure TOF_DECISIONACHAT.BBoutonClick(Sender: TObject);
var TOBL : TOB;
    found : boolean;
    Arow ,Niveau, LigneRef: integer;
begin
  found := false;
  NbLigne := GetRow;
  Arow := GetRow;
  NiveauMax := 1;
  // --
  TOBL:=GetTOBDecisAch(TOBdecisionnel,Arow) ; if TOBL=Nil then Exit ;
  Niveau := TOBL.GEtValue('NIVEAU');
  Ligneref := TOBL.getValue('NUMAFF');
  if Niveau = 1 then
  begin
    if (TOBL.GetValue('OPEN')='-') and (TOBL.detail.count > 0)  then
    BEGIN
      if TOBL.GetValue ('NIVEAU') >= MAXNiveauAff then exit;
      OuvreBranche (GS,TOBdecisionnel,NbLigne,niveaumax,found,false);
      TOBL.PutValue('OPEN','X');
    END else if (TOBL.GetValue('OPEN')='X') and (TOBL.detail.count > 0)  then
    BEGIN
      FermeBranche (GS,TOBdecisionnel,NbLigne,niveauMax,found);
      inc(niveaumax);
      TOBL.PutValue('OPEN','-');
    END;
  end else if niveau = 2 then
  begin
    if (TOBL.GetValue('OPEN1')='-') and (TOBL.detail.count > 0)  then
    BEGIN
      if TOBL.GetValue ('NIVEAU') >= MAXNiveauAff then exit;
      OuvreBranche (GS,TOBdecisionnel,NbLigne,niveauMax,found,false);
      TOBL.PutValue('OPEN','X');
      TOBL.PutValue('OPEN1','X');
    END else if (TOBL.GetValue('OPEN1')='X') and (TOBL.detail.count > 0)  then
    BEGIN
      FermeBranche (GS,TOBdecisionnel,NbLigne,NiveauMax ,found);
      inc(niveaumax);
      TOBL.PutValue('OPEN1','-');
      TOBL.PutValue('OPEN2','-');
    END;
  end else if niveau = 3 then
  begin
    if (TOBL.GetValue('OPEN2')='-') and (TOBL.detail.count > 0)  then
    BEGIN
      if TOBL.GetValue ('NIVEAU') >= MAXNiveauAff then exit;
      OuvreBranche (GS,TOBdecisionnel,NbLigne,NiveauMax,found,false);
      TOBL.PutValue('OPEN','X');
      TOBL.PutValue('OPEN1','X');
      TOBL.PutValue('OPEN2','X');
    END else if (TOBL.GetValue('OPEN2')='X') and (TOBL.detail.count > 0)  then
    BEGIN
      FermeBranche (GS,TOBdecisionnel,NbLigne,NiveauMax,found);
      inc(niveaumax);
      TOBL.PutValue('OPEN2','-');
    END;
  end;

  DefiniNiveauVisible (NiveauMax);
  AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,NiveauMax,NbLigne,LigneRef);
  PositionneLigne (Arow);
end;

procedure TOF_DECISIONACHAT.BDeleteClick(Sender: TObject);
begin
	if PGiAsk (Traduirememoire('Etes-vous sur de vouloir supprimer ce décisionnel ?'),ecran.caption)=Mryes then
  begin
    SupprimeDecisionnelAch;
    TToolBarButton97(GetControl('Bferme')).Click;
  end;
end;

procedure TOF_DECISIONACHAT.BImprimerClick(Sender: TObject);
Var TobDemande,TOBL : TOB;
		indice : integer;
    EtatDemandePrix : string;
begin
//   stWhere := ' AND BAD_NUMERO=' + NumDecisionnel ;
//   LanceEtat( 'E', 'GPJ', 'BDP', True, False, False, nil, stWhere, '', False) ;
  EtatDemandePrix := GetParamSocSecur ('SO_BTIMPDEP','BDP');
  If CONSULTFOU.IsExisteConsult then
  BEGIN
    CONSULTFOU.SetEtat (EtatDemandePrix);
    CONSULTFOU.PrepareEdition;
  end else
  begin
    TobDemande := TOB.Create ('DEMANDE DE PRIX',nil,-1);
    // préparation TOB pour impression
    for Indice := 0 to TOBDecisionnel.detail.Count -1 do
    Begin
      // on imprime que les lignes à prendre en compte
      if TOBDecisionnel.detail[Indice].GetValue('BAD_PRISENCOMPTE') = 'X' then
      begin
        TOBL := TOB.Create ('DECISIONACHLIG',TobDemande,-1);
        TOBL.Dupliquer( TOBDecisionnel.detail[Indice], False, true);
        TOBL.AddChampSupValeur('DESCRIPTIF','-');
        if (TOBL.getValue('BAD_BLOCNOTE') <> '') and (not IsBlocNoteVide(TOBL.getValue('BAD_BLOCNOTE'))) then
        begin
          TOBL.PutValue('DESCRIPTIF','X');
        end else
        begin
          TOBL.PutValue('BAD_BLOCNOTE','');
        end;
      end;
    end;
    LanceEtatTob ('E','GPJ',EtatDemandePrix,TobDemande,true,false,false,nil,'','Edition demande de prix',false);
    TobDemande.Free;
  end;
end;

procedure TOF_DECISIONACHAT.ZoneSuivanteOuOk(var ACol, ARow: Integer;var Cancel: boolean);
Var Sens,ii,Lim : integer ;
    OldEna,ChgLig,ChgSens  : boolean ;
BEGIN
  OldEna:=GS.SynEnabled ; GS.SynEnabled:=False ;
//  Sens:=-1 ; ChgLig:=(GS.Row<>ARow) ; ChgSens:=False ;
  Sens:=1 ; ChgLig:=(GS.Row<>ARow) ; ChgSens:=False ;
//  if GS.Row>ARow then Sens:=1 else if ((GS.Row=ARow) and (ACol<=GS.Col)) then Sens:=1 ;
  ACol:=GS.Col ; ARow:=GS.Row ; ii:=0 ;
  While Not ZoneAccessible(ACol,ARow)  do
  BEGIN
    Cancel:=True ; inc(ii) ; if ii>500 then Break ;
    if Sens=1 then
    BEGIN
      Lim:=GS.rowCount-1;
      if ((ACol=GS.ColCount-1) and (ARow>=Lim)) then
      BEGIN
        if ChgSens then Break else BEGIN Sens:=-1 ; Continue ; ChgSens:=True ; END ;
      END ;
      if ChgLig then BEGIN ACol:=GS_AFFAIRE-1 ; ChgLig:=False ; END ;
      if ACol<GS.ColCount-1 then Inc(ACol) else BEGIN Inc(ARow) ; ACol:=GS_AFFAIRE ; END ;
    END else
    BEGIN
      if ((ACol=GS_AFFAIRE) and (ARow=1)) then
      BEGIN
        if ChgSens then Break else BEGIN Sens:=1 ; Continue ; END ;
      END ;
      if ChgLig then BEGIN ACol:=GS_AFFAIRE -1; ChgLig:=False ; END ;
      if ACol>GS_AFFAIRE then Dec(ACol) else BEGIN Dec(ARow) ; ACol:=GS_AFFAIRE ; END ;
    END ;
  END ;
  GS.SynEnabled:=OldEna ;
end;

function TOF_DECISIONACHAT.ZoneAccessible(ACol, ARow: Integer): boolean;
Var TOBL : TOB ;
BEGIN
  Result:=True ;
  TOBL:=GetTOBDecisAch(TOBdecisionnel,GS.row) ; if TOBL=Nil then Exit ;
  if GS.ColLengths [Acol] = 0 then BEGIN Result := false;exit;END;
  if (TOBL.GetValue('NIVEAU')=4) and
  	 ((Acol <> GS_QUANTITESTK) and (Acol <> GS_FOURNISSEUR) and (Acol  <> GS_PRIXACHNET)) then
  begin
    // on permet de saisir un chantier sur une ligne ajoutée
  	if (Acol = GS_AFFAIRE) and (TOBL.GetValue('BAD_REFGESCOM')='') then exit;
    //
  	result := false;
    exit;
  end;
  if (TOBL.GetValue('NIVEAU')<4) and (Acol <> GS_FOURNISSEUR) and (Acol  <> GS_PRIXACHNET) then
  begin
  	result := false;
    exit;
  end;
end;

procedure TOF_DECISIONACHAT.SupprimeDecisionnelAch;
begin
  BEGINTRANS;
  TRY
  	ExecuteSQL ('DELETE FROM DECISIONACH WHERE BAE_NUMERO='+NumDecisionnel);
  	ExecuteSQL ('DELETE FROM DECISIONACHLIG WHERE BAD_NUMERO='+NumDecisionnel);
    COMMITTRANS;
  EXCEPT
  	ROLLBACK;
  END;
end;

procedure TOF_DECISIONACHAT.GERESTOCKChange (Sender : TObject);
var TOBL : TOB;
begin
  TOBL:=GetTOBDecisAch(TOBdecisionnel,GS.row) ; if TOBL=Nil then Exit ;
  SetPrisEnStock(TOBL,GERESTOCK.checked);
  AfficheInfos(GS.row);
end;


procedure TOF_DECISIONACHAT.LIVCHANTIERChange(Sender: TOBject);
var TOBL : TOB;
begin
  TOBL:=GetTOBDecisAch(TOBdecisionnel,GS.row) ; if TOBL=Nil then Exit ;
  SetLivSurChantier(TOBL,LIVCHANTIER.checked);
  AfficheInfos(GS.row);
end;

procedure TOF_DECISIONACHAT.PRISEENCOMPTEChange(Sender: TObject);
var TOBL : TOB;
begin
  TOBL:=GetTOBDecisAch(TOBdecisionnel,GS.row) ; if TOBL=Nil then Exit ;
  SetpriseEncompte(TOBL,PriseEnCompte.checked);
  AfficheInfos(GS.row);
end;

procedure TOF_DECISIONACHAT.InitligneDecisionnel (TOBpere,TOBL : TOB; niveau : integer);
var LivChantier : string;
begin
	if Niveau < 4 then TOBpere.putValue('BAD_TYPEL','NI'+inttoStr(Niveau))
  							else TOBpere.putValue('BAD_TYPEL','ART');
	AddlesChampsSupDet (TOBPere);

	LivChantier := TOBL.GetValue('BAD_LIVCHANTIER');
  TOBpere.PutValue('BAD_LIVCHANTIER',LivChantier);

  if Niveau < 4 then
  begin
    if Niveau = 1 then
    begin
      TOBpere.PutValue('BAD_LIVCHANTIER','');
      TOBpere.PutValue('OPEN','X');
    end else if Niveau = 2 then
    begin
      TOBpere.PutValue('OPEN','X');
      TOBpere.PutValue('OPEN1','X');
    end else if Niveau = 3 then
    begin
      TOBpere.PutValue('OPEN','X');
      TOBpere.PutValue('OPEN1','X');
      TOBpere.PutValue('OPEN2','X');
    end;
  end;
  TOBpere.putValue('BAD_TYPEARTICLE',TOBL.GetValue('BAD_TYPEARTICLE'));
  TOBpere.putValue('BAD_CODEARTICLE',TOBL.GetValue('BAD_CODEARTICLE'));
  TOBpere.putValue('BAD_ARTICLE',TOBL.GetValue('BAD_ARTICLE'));
//  TOBpere.putValue('BAD_LIBELLE',TOBL.GetValue('BAD_LIBELLE'));
  TOBpere.putValue('BAD_PRISENCOMPTE','X');
  TOBpere.putValue('BAD_TENUESTOCK',TOBL.GetValue('BAD_TENUESTOCK'));
  TOBpere.putValue('BAD_QUALIFQTEVTE',TOBL.GetValue('BAD_QUALIFQTEVTE'));
  TOBpere.putValue('BAD_QUALIFQTESTO',TOBL.GetValue('BAD_QUALIFQTESTO'));
  TOBpere.PutValue('BAD_QUALIFQTEACH',TOBL.GetValue('BAD_QUALIFQTEACH') );
  TOBpere.putValue('BAD_DATELIVRAISON',TOBL.GetValue('BAD_DATELIVRAISON'));
  TOBpere.putValue('GA_LIBELLE',TOBL.GetValue('GA_LIBELLE'));
  TOBpere.putValue('GA_TENUESTOCK',TOBL.GetValue('GA_TENUESTOCK'));
  TOBpere.putValue('T_LIBELLE',TOBL.GetValue('T_LIBELLE'));
  TOBpere.putValue('AFF_LIBELLE',TOBL.GetValue('AFF_LIBELLE'));
  TOBpere.putValue('BAD_ADDLIGNE','X');

  if Niveau > 2 then
  begin
  	TOBpere.PutValue('BAD_DEPOT',TOBL.GEtValue('BAD_DEPOT'));
  end;

end;

procedure TOF_DECISIONACHAT.InitAjoutLigne (TOBPere : TOB; Article,Depot,Livchantier : string; Niveau : integer) ;
var TOBA: TOB;
    requete,Etab,LeFournisseur,Larticle : string;
    TarifFourn : TGinfostarifFour;
begin
	TOBA := TOB.Create('ARTICLE',nil,-1);
  TRY
    TOBA.SelectDB ('"'+Article+'"',nil);
		if Niveau < 4 then TOBpere.putValue('BAD_TYPEL','NI'+inttoStr(Niveau))
  								else TOBpere.putValue('BAD_TYPEL','ART');
    TOBpere.putValue('BAD_TYPEARTICLE',TOBA.GEtValue('GA_TYPEARTICLE'));
    AddlesChampsSupDet (TOBPere);

    TOBpere.PutValue('BAD_LIVCHANTIER',Livchantier);
    if Niveau < 4 then
    begin
      if Niveau = 1 then
      begin
        TOBpere.PutValue('BAD_LIVCHANTIER','');
        TOBpere.PutValue('OPEN','X');
        TOBpere.PutValue('NIVEAU',1);
    end else if Niveau = 2 then
      begin
        TOBpere.PutValue('OPEN','X');
        TOBpere.PutValue('OPEN1','X');
        TOBpere.PutValue('NIVEAU',2);
      end else if Niveau = 3 then
      begin
        TOBpere.PutValue('OPEN','X');
        TOBpere.PutValue('OPEN1','X');
        TOBpere.PutValue('OPEN2','X');
        TOBpere.PutValue('NIVEAU',3);
      end else
      begin
      	TOBpere.PutValue('NIVEAU',4);
      end;
    end;
    if TOBPere.GetValue('BAD_LIBELLE')='' then
    begin
    	TOBPere.PutValue('BAD_LIBELLE',TOBA.GetValue('GA_LIBELLE'));
    end;
    if TOBPere.GetValue('BAD_BLOCNOTE')='' then
    begin
    	TOBPere.PutValue('BAD_BLOCNOTE',TOBA.GetValue('GA_BLOCNOTE'));
    end;
    TOBpere.putValue('BAD_TYPEARTICLE',TOBA.GetValue('GA_TYPEARTICLE'));
    TOBpere.putValue('BAD_CODEARTICLE',Trim(Copy(Article,1,18)));
    TOBpere.putValue('BAD_ARTICLE',Article);
//    TOBpere.putValue('BAD_LIBELLE',TOBA.GetValue('GA_LIBELLE'));
    TOBpere.putValue('BAD_PRISENCOMPTE','X');
    TOBpere.putValue('BAD_QUALIFQTEVTE',TOBA.GetValue('GA_QUALIFUNITEVTE'));
    TOBpere.putValue('BAD_QUALIFQTESTO',TOBA.GetValue('GA_QUALIFUNITESTO'));
    TOBpere.putValue('BAD_TENUESTOCK',TOBA.GetValue('GA_TENUESTOCK'));
  	TOBpere.putValue('BAD_ADDLIGNE','X');
  	TOBpere.putValue('BAD_DATELIVRAISON',V_PGI.DateEntree);
  	TOBpere.putValue('BAD_SUPPRIME','-');
    TOBPere.PutValue('GA_LIBELLE',TOBA.GetValue('GA_LIBELLE'));
    TOBPere.PutValue('GA_TENUESTOCK',TOBA.GetValue('GA_TENUESTOCK'));
  	Etab := VH^.ProfilUserC[prEtablissement].Etablissement;
  	if Etab = '' then Etab := VH^.EtablisDefaut;
  	IF ETab = '' then Etab := FindEtablissement;
  	TOBpere.putValue('BAD_ETABLISSEMENT',Etab);

    if Niveau > 2 then
    begin
      TOBpere.PutValue('BAD_DEPOT',Depot);
      if (Niveau = 4) then
      begin
      	TOBPere.putValue('QTEPHY',GetPhysique(Article,Depot));
        //
        LeFournisseur := TOBPere.GetValue('BAD_FOURNISSEUR');
        Larticle := TOBPere.GetValue('BAD_ARTICLE');
        //
        TarifFourn := GetInfosTarifAch (LeFournisseur,Larticle,TurAchat,false,true,nil,TOBpere.GetValue('BAD_QUANTITEACH'),TOBPere.GetValue('BAD_DATELIVRAISON'));
        TOBPere.putValue('BAD_PRIXACH',TarifFourn.TarifAchBrut );
        TOBPere.putValue('BAD_CALCULREMISE',TarifFourn.Remise  );
        TOBPere.putValue('BAD_PRIXACHNET',TarifFourn.TarifAch );
        TOBPere.putValue('BAD_COEFUAUS',TarifFourn.CoefUAUs );
        TOBPere.putValue('BAD_QUALIFQTEACH',TarifFourn.UniteAchat);
      end;
    end;
  FINALLY
  	TOBA.free;
  END;
end;

function TOF_DECISIONACHAT.FindParent (Article,Livchantier,Depot : string) : TOB;
var LaTOBN1 , LaTOBN2,LaTOBN3,LaTOB : TOB;
begin
  LaTOBN1 := TOBdecisionnel.findFirst(['BAD_TYPEL','BAD_ARTICLE'],['NI1',Article],true);
  if LaTOBN1 = nil then
  begin
  	// y'en a po ..alors on la cree
    LaTOBN1 := TOB.Create('DECISIONACHLIG',TOBdecisionnel,-1);
    InitAjoutLigne (LaTOBN1,Article,Depot,LivChantier,1);
  end else LaTOBN1.putValue('OPEN','X');

  LaTOBN2 := LaTOBN1.findFirst(['BAD_TYPEL','BAD_ARTICLE','BAD_LIVCHANTIER'],['NI2',Article,LivChantier],True);
  if LaTOBN2 = nil then
  begin
  	// y'en a po ..alors on la cree
    LaTOBN2 := TOB.Create('DECISIONACHLIG',LaTOBN1,-1);
    InitAjoutLigne (LaTOBN2,Article,Depot,LivChantier,2);
  end else LaTOBN2.putValue('OPEN1','X');

  LaTOBN3 := LaTOBN2.findFirst(['BAD_TYPEL','BAD_ARTICLE','BAD_LIVCHANTIER','BAD_DEPOT'],['NI3',Article,LivChantier,Depot],True);
  if LaTOBN3 = nil then
  begin
  	// y'en a po ..alors on la cree
    LaTOBN3 := TOB.Create('DECISIONACHLIG',LaTOBN2,-1);
    InitAjoutLigne (LaTOBN3,Article,Depot,LivChantier,3);
  end else LaTOBN2.putValue('OPEN2','X');

  LaTOB := TOB.create('DECISIONACHLIG',LaTOBN3,-1);
  InitAjoutLigne (LaTOB,Article,Depot,LivChantier,4);
  result := LaTOB;
end;


procedure TOF_DECISIONACHAT.MepLigneInDecisionnel (TOBL : TOB; ChangementParent : boolean=true);
var LaTOBN1,LaTOBN2,LaTOBN3 : TOB;
		LivChantier : String;
    NewTOB : TOB;
begin
	LivChantier := TOBL.GetValue('BAD_LIVCHANTIER');
  LaTOBN1 := TOBdecisionnel.findFirst(['BAD_TYPEL','BAD_ARTICLE'],['NI1',TOBL.GetValue('BAD_ARTICLE')],true);
  if LaTOBN1 = nil then
  begin
  	// y'en a po ..alors on la cree
    LaTOBN1 := TOB.Create('DECISIONACHLIG',TOBdecisionnel,-1);
    InitligneDecisionnel (LaTOBN1,TOBL,1);
  end else LaTOBN1.putValue('OPEN','X');

  LaTOBN2 := LaTOBN1.findFirst(['BAD_TYPEL','BAD_ARTICLE','BAD_LIVCHANTIER'],['NI2',TOBL.GetValue('BAD_ARTICLE'),LivChantier],True);
  if LaTOBN2 = nil then
  begin
  	// y'en a po ..alors on la cree
    LaTOBN2 := TOB.Create('DECISIONACHLIG',LaTOBN1,-1);
    InitligneDecisionnel (LaTOBN2,TOBL,2);
  end else LaTOBN2.putValue('OPEN1','X');

  LaTOBN3 := LaTOBN2.findFirst(['BAD_TYPEL','BAD_ARTICLE','BAD_LIVCHANTIER','BAD_DEPOT'],['NI3',TOBL.GetValue('BAD_ARTICLE'),LivChantier,TOBL.GetValue('BAD_DEPOT')],True);
  if LaTOBN3 = nil then
  begin
  	// y'en a po ..alors on la cree
    LaTOBN3 := TOB.Create('DECISIONACHLIG',LaTOBN2,-1);
    InitligneDecisionnel (LaTOBN3,TOBL,3);
  end  else LaTOBN3.putValue('OPEN2','X');
  if ChangementParent then
  begin
    TOBL.ChangeParent (LaTOBN3,-1);
  end else
  begin
    NewTOB := TOB.Create('DECISIONACHLIG',laTOBN3,-1);
    AddlesChampsSup(NewTOB,MaxRemplace);
    NewTOB.Dupliquer (TOBL,true,true);
    NewTOB.putValue('_A_TRAITER','-');
  end;
end;

procedure TOF_DECISIONACHAT.NettoieParent (TheParent : TOB);
var TheGrandPere : TOB;
begin
	ThegrandPere := TheParent.Parent;
  if TheParent.detail.count = 0 then TheParent.free;
  if ThegrandPere <> nil then NettoieParent (ThegrandPere);
end;

function TOF_DECISIONACHAT.NettoieFilles (TheParent : TOB) : boolean;
var TheFille : TOB;
		Indice : integer;
begin
	result := True;
	Indice := 0 ;
  if TheParent.detail.count > 0 then
  begin
    repeat
    	TheFille := TheParent.detail[Indice];
      if TheFille.getValue('NIVEAU')=4 then break;
      result := NettoieFilles(TheFille);
      if result then inc(indice);
    until indice > TheParent.detail.count -1;
  end;
  if TheParent.detail.count = 0 then
  begin
  	TheParent.free;
    result := false;
  end;
end;


procedure TOF_DECISIONACHAT.SetNumligne(TOBPere: TOB; IndiceN1,IndiceN2,IndiceN3,IndiceN4,Niveau: integer);
var Indice : integer;
		TOBD : TOB;
    Cpt1,Cpt2,Cpt3,Cpt4 : integer;
begin
	if TOBPere.detail.count = 0 then exit;

  Cpt1 := IndiceN1;
  Cpt2 := IndiceN2;
  Cpt3 := IndiceN3;
  Cpt4 := IndiceN4;

	for Indice := 0 to TOBpere.detail.count -1 do
  begin
  	TOBD := TOBPere.detail[Indice];
    if Niveau = 2 then inc(Cpt2);
    if Niveau = 3 then inc(Cpt3);
    if Niveau = 4 then inc(Cpt4);
    //
    TOBD.putValue('BAD_NUMN1',Cpt1);
    TOBD.putValue('BAD_NUMN2',Cpt2);
    TOBD.putValue('BAD_NUMN3',Cpt3);
    TOBD.putValue('BAD_NUMN4',Cpt4);
    if TOBD.detail.count > 0 then SetNumligne (TOBD,cpt1,Cpt2,CPt3,CPt4,Niveau+1);
  end;
end;


procedure TOF_DECISIONACHAT.ReindiceDecisionnel;
var Indice : integer;
		TOBD : TOB;
begin
  for Indice := 0 to TOBdecisionnel.detail.count -1 do
  begin
    TOBD := TOBdecisionnel.detail[Indice];
    TOBD.putvalue('BAD_NUMN1',Indice+1);
    SetNumligne (TOBD,Indice+1,0,0,0,2);
  end;
end;



function TOF_DECISIONACHAT.GetPhysique (Article,Depot : string) : double;
  var QQ : TQuery;
      requete : string;
begin
  result := 0;
  requete := 'SELECT GQ_PHYSIQUE FROM DISPO WHERE GQ_ARTICLE="'+Article+'" AND '+
             'GQ_DEPOT="'+Depot+'" AND GQ_CLOTURE="-"';
  QQ := OpenSql (Requete,true,-1,'',true);
  if not QQ.eof then
  begin
    result := QQ.findField('GQ_PHYSIQUE').AsFloat;
  end;
  ferme (QQ);
end;

procedure TOF_DECISIONACHAT.SetDepot(TOBL : TOB;DEPOT : string);

  procedure SetDepotDet (TOBD : TOB;DEPOT : string; var changeparent : boolean);
  var TOBS,TOBPAPOU : TOB;
  		Indice,Ligne : integer;
  begin
  	if TOBD.GEtValue('NIVEAU') = 4 then
    begin
      Ligne := TOBD.GetValue('NUMAFF');
      TOBPAPOU := TOBD.parent;
      if DEPOT <> TOBD.getValue('BAD_DEPOT') then changeparent := true;
      TOBD.PutValue('BAD_DEPOT',DEPOT);
      TOBD.putValue('QTEPHY',Getphysique(TOBL.GetValue('BAD_ARTICLE'),DEPOT));
      MepLigneInDecisionnel (TOBD); // dans le cas ou le depot change
      if Ligne > 0 then TOBD.PutLigneGrid (GS,Ligne,false,false,ListeSaisie);
    end else
    begin
      if TOBD.Detail.count > 0 then
      begin
        indice := 0;
        repeat
          TOBS := TOBD.Detail[Indice];
          changeparent := false;
          SetDepotDet (TOBS,DEPOT,changeparent);
          if (not changeparent) then inc(Indice);
        until indice >= TOBD.detail.count ;
      end;
    end;
  end;

var Ligne : Integer;
		TheParentInit : TOB;
    Niveau : integer;
    Article,LivCha: string;
    chp : boolean;
begin
  Niveau := TOBL.GetValue('NIVEAU');
  Article := TOBL.GEtValue('BAD_ARTICLE');
  Livcha := TOBL.GEtValue('BAD_LIVCHANTIER');
  Ligne := TOBL.GetValue('NUMAFF');

  TheParentInit := FindFirstParent(TOBL);
  //
  if Niveau > 1 then ReinitNiveauInf (TheParentInit);
  Chp := False;
  SetDepotDet (TOBL,DEPOT,chp); // dans la descente
  //
  NettoieFilles (TheParentInit);
	ReinitNiveauInf (TheParentInit);
  recalculelesNiveaux (TheParentInit);
  ReindiceDecisionnel;
  //
  TheParentInit := TOBdecisionnel.findFirst(['BAD_TYPEL','BAD_ARTICLE'],['NI1',Article],true);
  TOBL := TheParentInit;
  CumulStkPhysique (TheParentInit);
  //
  NbLigne := 0;
  ReIndiceGrille (TOBdecisionnel,NbLigne);
  DefiniNiveauVisible (MaxNiveauAff);
  AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,5,NbLigne);
  PositionneLigne (Ligne);
  AfficheInfos(Ligne);
end;

procedure TOF_DECISIONACHAT.SetLivSurChantier(TOBL : TOB;Actif: boolean);
var Ligne : integer;
		TheParent,TheParentInit : TOB;
begin

  if LIVCHANTIER.Checked then
  begin
  	TOBL.putValue('BAD_LIVCHANTIER','X');
//  	if GERESTOCK.enabled then TOBL.putValue('BAD_TENUESTOCK','-');
  end else
  begin
  	TOBL.putValue('BAD_LIVCHANTIER','-');
//  	if GERESTOCK.enabled then TOBL.putValue('BAD_TENUESTOCK','X');
  end;
  TheParent := TOBL.Parent; // sauvegarde du parent
  TheParentInit := FindParentInit(TOBL);
	EnleveDesPeres (TOBL,TheParent);
	MepLigneInDecisionnel (TOBL);
  NettoieParent (TheParent);
  ReindiceDecisionnel;
  CumulStkPhysique (TheParentInit);
  AjouteAuxPeres (TOBL,TOBL.Parent);
  SetPrisEnSTockpapou (TOBL); // et on remonte
  NbLigne := 0;
  ReIndiceGrille (TOBdecisionnel,NbLigne);
	AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,5,NbLigne);
  Ligne := TOBL.GetValue('NUMAFF');
  PositionneLigne (Ligne);
  AfficheInfos(Ligne);
end;

procedure TOF_DECISIONACHAT.SetPrisEnSTockpapou (TOBL: TOB);

  function GetTenueDetail (TOBL : TOB) : String;
  var indice : integer;
  		TOBD : TOB;
  begin
  	result := '';
  	for indice := 0 to TOBL.detail.count -1 do
    begin
    	TOBD := TOBL.detail[Indice];
      if indice = 0 then result := TOBD.GetValue('BAD_TENUESTOCK');
      if result <> TOBD.GetValue('BAD_TENUESTOCK') then
      begin
      	result := '';
        break;
      end;
    end;
  end;

var TOBP : TOB;
    Ligne : integer;
begin
  TOBP := TOBL.Parent;
  if TOBP <> nil then
  begin
    if TOBP.fieldExists ('NUMAFF') then
    begin
      Ligne := TOBP.GetValue('NUMAFF');
      TOBP.PutValue('BAD_TENUESTOCK',GetTenueDetail(TOBP));
      if Ligne > 0 then TOBP.PutLigneGrid (GS,Ligne,false,false,ListeSaisie);
      SetPrisEnSTockpapou (TOBP);
    end;
  end;
end;
//
procedure TOF_DECISIONACHAT.ReinitNiveauInf (TOBI : tob);
var Indice : integer;
    TOBS : TOB;
begin
  if TOBI.GetValue('NIVEAU') = 4 then exit;
  TOBI.PutValue ('BAD_QUANTITEVTE',0);
  TOBI.PutValue ('BAD_QUANTITESTK',0);
  TOBI.PutValue ('BAD_QUANTITEACH',0);
  TOBI.PutValue ('BAD_QUANTITEINIT',0);
  TOBI.PutValue ('QTEPHY',0);
  TOBI.PutValue ('MULTIPABASE','-');
  TOBI.PutValue ('BAD_PABASE',0);
  for Indice := 0 to TOBI.detail.count -1 do
  begin
    TOBS := TOBI.detail[Indice];
    ReinitNiveauInf (TOBS);
  end;
end;

procedure TOF_DECISIONACHAT.setcumul (TOBPere,TOBdetail : tob);
var Indice,niveau : integer;
    TOBD : TOB;
begin
  if TOBDetail.detail.count > 0 then
  begin
    for Indice := 0 to TOBDetail.detail.count -1 do
    begin
      TOBD := TOBDetail.detail[Indice];
      SetCumul (TOBDetail,TOBD);
    end;
  end;
  if TOBPere <> nil then
  begin
    if TOBpere.getValue ('MULTIPABASE') = '-' then
    begin
      if (TOBpere.GetValue('BAD_PABASE')=0) then
      begin
         TOBPere.putValue('BAD_PABASE',TOBdetail.GetValue('BAD_PABASE'));
      end else if (TOBpere.GetValue('BAD_PABASE')<>TOBdetail.GetValue('BAD_PABASE')) then
      begin
         TOBPere.putValue('BAD_PABASE',0);
         TOBpere.putValue ('MULTIPABASE','X');
      end;
    end;
    TOBPere.putvalue ('BAD_QUANTITEACH',TOBPere.Getvalue ('BAD_QUANTITEACH')+TOBDetail.Getvalue ('BAD_QUANTITEACH'));
    TOBPere.putvalue ('BAD_QUANTITEVTE',TOBPere.Getvalue ('BAD_QUANTITEVTE')+TOBDetail.Getvalue ('BAD_QUANTITEVTE'));
    TOBPere.putvalue ('BAD_QUANTITEINIT',TOBPere.Getvalue ('BAD_QUANTITEINIT')+TOBDetail.Getvalue ('BAD_QUANTITEINIT'));
    TOBPere.putvalue ('BAD_QUANTITESTK',TOBPere.Getvalue ('BAD_QUANTITESTK')+TOBDetail.Getvalue ('BAD_QUANTITESTK'));
  end;
end;

procedure TOF_DECISIONACHAT.recalculelesNiveaux (TOBI : TOB);
var Indice : integer;
    TOBD : TOB;
begin
  if TOBI.detail.count > 0 then
  begin
    for Indice := 0 to TOBI.detail.count -1 do
    begin
      TOBD := TOBI.detail[Indice];
      SetCumul (TOBI,TOBD);
    end;
  end;
end;

//



procedure TOF_DECISIONACHAT.SetPrisEnStock(TOBL : TOB;Actif: boolean);

  procedure SetPrisEnStockDet (TOBD : TOB;Actif : boolean; var Newlivchantier : string);
  var TOBS,TOBPAPOU : TOB;
  		Indice,Ligne : integer;
      ATraiter : boolean;
      fNewLivChantier,OldLivChantier : string;
  begin
  	ATraiter := false;

  	if TOBD.GEtValue('NIVEAU') = 4 then
    begin
      Ligne := TOBD.GetValue('NUMAFF');
      if Actif then
      begin
      	TOBPAPOU := TOBD.parent;
        TOBD.PutValue('BAD_TENUESTOCK','X');
        if TOBD.GetValue('BAD_LIVCHANTIER') <> '-' then ATraiter := true;
        TOBD.PutValue('BAD_LIVCHANTIER','-');
        if Atraiter then MepLigneInDecisionnel (TOBD); // dans le cas ou le livre chantier change
      end else
      begin
        TOBD.PutValue('BAD_TENUESTOCK','-');
      end;
      NewLivChantier := TOBD.GetValue('BAD_LIVCHANTIER');
      if Ligne > 0 then TOBD.PutLigneGrid (GS,Ligne,false,false,ListeSaisie);
    end else
    begin
    	if not Actif then TOBD.PutValue('BAD_TENUESTOCK','-') else TOBD.PutValue('BAD_TENUESTOCK','X');
    end;

    if TOBD.Detail.count > 0 then
    begin
    	indice := 0;
    	repeat
      	TOBS := TOBD.Detail[Indice];
        OldLivChantier := TOBS.GetValue('BAD_LIVCHANTIER');
        SetPrisEnStockDet (TOBS,Actif,FNewLivChantier);
        if FnewlivChantier <> '' then NewLivChantier := FNewlivchantier;
        if (FnewLivChantier = OldLivChantier) or (TOBS.GetValue('NIVEAU')<>4) then inc(Indice);
      until indice > TOBD.detail.count -1;
    end;
  end;

var Ligne : Integer;
		TheParentInit : TOB;
    Niveau : integer;
    Article,LivChantier,Depot : string;
begin
	Niveau := TOBL.GetValue('NIVEAU');
  TheParentInit := FindFirstParent(TOBL);
  Article := TOBL.GEtValue('BAD_ARTICLE');
  Depot := TOBL.GEtValue('BAD_DEPOT');
  //
  if Niveau > 1 then ReinitNiveauInf (TheParentInit);

  SetPrisEnStockDet (TOBL,Actif,LivChantier); // dans la descente
  if Niveau > 1 then
  begin
  	// on a change les clefs donc il faut rechercher le nouvel emplacement
    if niveau < 4 then TOBL := TheParentInit.findFirst(['BAD_TYPEL','BAD_ARTICLE','BAD_LIVCHANTIER','BAD_DEPOT'],['NI'+InttoStr(Niveau),Article,LivChantier,depot],True);
  end;
  NettoieFilles (TheParentInit);
  SetPrisEnSTockpapou (TOBL); // et on remonte
  if Niveau > 1 then recalculelesNiveaux (TheParentInit);
  ReindiceDecisionnel;
  NbLigne := 0;
  ReIndiceGrille (TOBdecisionnel,NbLigne);
  AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,5,NbLigne);
  Ligne := TOBL.GetValue('NUMAFF');
  PositionneLigne (Ligne);
  AfficheInfos(Ligne);
end;

procedure TOF_DECISIONACHAT.SetpriseEncompte(TOBL : TOB;Actif: boolean);
var TOBD : TOB;
		Ligne,Indice : Integer;
begin
  if Actif then TOBL.PutValue('BAD_PRISENCOMPTE','X')
  				 else TOBL.PutValue('BAD_PRISENCOMPTE','-');
  Ligne := TOBL.GetValue('NUMAFF');
  if Ligne > 0 then TOBL.PutLigneGrid (GS,Ligne,false,false,ListeSaisie);

  if TOBL.Detail.count > 0 then
  begin
  	for Indice := 0 to TOBL.detail.count -1 do
    begin
    	TOBD := TOBL.detail[Indice];
      SetpriseEncompte (TOBD,Actif);
    end;
  end;
end;

procedure TOF_DECISIONACHAT.CumulStkPhysique (TOBdecisionnel : TOB);

	procedure CumulStkThisNiveau(TOBP : TOB);
  var Indice,Niveau : integer;
  		TOBL : TOB ;
  begin
  	for Indice := 0 TO TOBP.detail.count -1 do
    begin
    	TOBL := TOBP.detail[Indice];
      Niveau := TOBL.getValue('NIVEAU');
      if TOBL.detail.count > 0 then
      begin
        CumulStkThisNiveau (TOBL);
      end;
      if Niveau=1 then  TOBP.putValue('QTEPHY',TOBP.GetValue('QTEPHY')+TOBL.GetValue('QTEPHY'))
      						else  TOBP.putValue('QTEPHY',TOBL.GetValue('QTEPHY'));
    end;
  end;

var Indice : integer;
		Niveau : integer;
    TOBP : TOB;
begin
	if TOBDecisionnel = nil then exit;
  for Indice :=0 to  TOBdecisionnel.detail.count -1 do
  begin
  	TOBP := TOBdecisionnel.detail[Indice];
  	CumulStkThisNiveau(TOBP);
    if TOBDecisionnel.FieldExists ('QTEPHY') then
    begin
      Niveau := TOBP.getValue('NIVEAU');
      if Niveau=1 then  TOBdecisionnel.putValue('QTEPHY',TOBdecisionnel.GetValue('QTEPHY')+TOBP.GetValue('QTEPHY'))
      						else  TOBdecisionnel.putValue('QTEPHY',TOBP.GetValue('QTEPHY'));
    end;
  end;
end;

procedure TOF_DECISIONACHAT.ActiveEvents;
begin
  PRISEENCOMPTE.OnClick := PRISEENCOMPTEChange;
  GERESTOCK.OnClick := GERESTOCKChange;
  LIVCHANTIER.OnClick := LIVCHANTIERChange;
  BAD_DATELIVRAISON.OnExit := BAD_DATELIVRAISONClick;
  ASUPPRIMER.OnClick := ASupprimerClick;
  BVOIRSTOCK.onclick := VOIRESTOCKSClick;
  DEPOT.onchange := DEPOTCHANGE;
end;

procedure TOF_DECISIONACHAT.DesactiveEvents;
begin
  PRISEENCOMPTE.OnClick := nil;
  GERESTOCK.OnClick := nil;
  LIVCHANTIER.OnClick := nil;
  BAD_DATELIVRAISON.OnExit := nil;
  ASUPPRIMER.OnClick := nil;
  BVOIRSTOCK.onclick := nil;
  DEPOT.onchange := nil;
end;

procedure TOF_DECISIONACHAT.VoirArticleClick(Sender: TObject);
var TOBL : TOB;
begin
  TOBL:=GetTOBDecisAch(TOBdecisionnel,GS.row) ; if TOBL=Nil then Exit ;
  V_PGI.DispatchTT (7,TaConsult,TOBL.GetValue('BAD_ARTICLE'),'','TYPEARTICLE='+TOBL.GetValue('BAD_TYPEARTICLE'));
end;

procedure TOF_DECISIONACHAT.VoirStockClick(Sender: TObject);
var TOBL : TOB;
begin
  TOBL:=GetTOBDecisAch(TOBdecisionnel,GS.row) ; if TOBL=Nil then Exit ;
	AGLLanceFiche('BTP','BTDISPO','',TOBL.GetValue('BAD_ARTICLE')+';'+TOBL.GetValue('BAD_DEPOT')+';-','CONSULTATION') ;
end;

procedure TOF_DECISIONACHAT.VoirBesoinClick(Sender: TObject);
var TOBL : TOB;
		znum : string;
    cledoc : r_cledoc;
begin
  TOBL:=GetTOBDecisAch(TOBdecisionnel,GS.row) ; if TOBL=Nil then Exit ;
  DecodeRefPiece (TOBL.GetValue('BAD_REFGESCOM'),cledoc);
  znum :=CleDoc.NaturePiece+';'+FormatDateTime('dd/mm/yyyy',Cledoc.DatePiece)+';'+ CleDoc.Souche +';'
      +IntToStr(CleDoc.NumeroPiece)+';'+IntToStr(CleDoc.Indice)+';;';
  AppelPiece ([ZNum,'ACTION=CONSULTATION'],2);
end;

procedure TOF_DECISIONACHAT.AjoutligneClick(Sender: TOBject);
var TOBL,TOBA : TOB;
		NewTOBL : TOB;
    Ligne,NbLigne  : integer;

begin
	TOBA := TOB.Create('ARTICLE',nil,-1);

  TOBL:=GetTOBDecisAch(TOBdecisionnel,GS.row) ; if TOBL=Nil then Exit ;
  NewTOBL := TOB.Create('DECISIONACHLIG',TOBL.Parent,-1);
  NewTOBL.putValue('BAD_TYPEL',TOBL.GetValue('BAD_TYPEL'));
	AddlesChampsSupDet (NewTOBL);
  NewTOBL.Dupliquer (TOBL,true,true);
  TOBA.SelectDB ('"'+TOBL.GEtVALUE('BAD_ARTICLE')+'"',nil);

  NewTOBL.putValue('BAD_LIBELLE',TOBA.GEtVAlue('GA_LIBELLE'));
  NewTOBL.putValue('BAD_BLOCNOTE',TOBA.GEtVAlue('GA_BLOCNOTE'));
  NewTOBL.putValue('BAD_REFGESCOM','');
  NewTOBL.putValue('AFF_LIBELLE','');
  NewTOBL.putValue('BAD_AFFAIRE','');
  NewTOBL.putValue('BAD_AFFAIRE1','');
  NewTOBL.putValue('BAD_AFFAIRE2','');
  NewTOBL.putValue('BAD_AFFAIRE3','');
  NewTOBL.putValue('BAD_AVENANT','');
  NewTOBL.putValue('BAD_ADDLIGNE','X');
  newTOBL.putValue('BAD_SUPPRIME','-');
  //
  AjouteAuxPeres (TOBL,TOBL.Parent);
  //
  NbLigne := 0;
  ReIndiceGrille (TOBdecisionnel,NbLigne);
	AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,5,NbLigne);
  Ligne := NewTOBL.GetValue('NUMAFF');
  PositionneLigne (Ligne);
  AfficheInfos(Ligne);

end;

procedure TOF_DECISIONACHAT.MnScinderClick(Sender: Tobject);
var TOBL : TOB;
		NewTOBL : TOB;
    Ligne,NbLigne  : integer;

begin
  TOBL:=GetTOBDecisAch(TOBdecisionnel,GS.row) ; if TOBL=Nil then Exit ;
  NewTOBL := TOB.Create('DECISIONACHLIG',TOBL.Parent,-1);
  NewTOBL.putValue('BAD_TYPEL',TOBL.GetValue('BAD_TYPEL'));
	AddlesChampsSupDet (NewTOBL);
  NewTOBL.Dupliquer (TOBL,true,true);
  NewTOBL.putValue('GA_TENUESTOCK','-');
  NewTOBL.putValue('BAD_TENUESTOCK','-');
  NewTOBL.putValue('BAD_SCINDEE','X');
  NewTOBL.putValue('BAD_SUPPRIME','-');
  NewTOBL.putValue('BAD_ADDLIGNE','X');
  //
  TOBL.putValue('BAD_SCINDEE','X');
  //
  AjouteAuxPeres (TOBL,TOBL.Parent);
  //
  NbLigne := 0;
  ReIndiceGrille (TOBdecisionnel,NbLigne);
	AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,5,NbLigne);
  Ligne := NewTOBL.GetValue('NUMAFF');
  PositionneLigne (Ligne);
  AfficheInfos(Ligne);
end;


procedure TOF_DECISIONACHAT.MnRemplaceClick(Sender: Tobject);
  procedure SetRemplaceLigne (TOBL,TOBI : TOB);
  var FUV,FUS,FUA,CoefuaUs : double;

  begin
  	Inc(MaxRemplace);
  	FUV := RatioMesure('PIE', TobL.GetValue('BAD_QUALIFQTEVTE')); if FUV = 0 then FUV := 1;
    FUS := RatioMesure('PIE', TobL.GetValue('BAD_QUALIFQTESTO')); if FUS = 0 then FUS := 1;
    FUA := RatioMesure('PIE', TobL.GetValue('BAD_QUALIFQTEACH')); if FUA = 0 then FUA := 1;
    CoefuaUs := TobL.GetValue('BAD_COEFUAUS');
    TOBL.putValue('BAD_DATELIVRAISON',TOBI.GetValue('BAD_DATELIVRAISON'));
    TOBL.putValue('BAD_AFFAIRE',TOBI.GetValue('BAD_AFFAIRE'));
    TOBL.putValue('BAD_AFFAIRE1',TOBI.GetValue('BAD_AFFAIRE1'));
    TOBL.putValue('BAD_AFFAIRE2',TOBI.GetValue('BAD_AFFAIRE2'));
    TOBL.putValue('BAD_AFFAIRE3',TOBI.GetValue('BAD_AFFAIRE3'));
    TOBL.putValue('BAD_AVENANT',TOBI.GetValue('BAD_AVENANT'));
    TOBL.putValue('BAD_REFGESCOM',TOBI.GetValue('BAD_REFGESCOM'));
    TOBL.putValue('BAD_PABASE',TOBI.GetValue('BAD_PABASE'));
    //
    TOBL.putValue('BAD_QUANTITEVTE',TOBI.GetValue('BAD_QUANTITEVTE'));
    TOBL.putValue('BAD_QUANTITESTK',TOBI.GetValue('BAD_QUANTITEVTE')*FUV/FUS);
    if CoefuaUs <> 0 then
    begin
    	TOBL.putValue('BAD_QUANTITEACH',TOBI.GetValue('BAD_QUANTITEVTE')*FUV/FUS / CoefUAUS);
    end else
    begin
    	TOBL.putValue('BAD_QUANTITEACH',TOBI.GetValue('BAD_QUANTITEVTE')*FUV/FUA);
    end;
    //
    TOBL.putValue('BAD_LIVCHANTIER',TOBI.GetValue('BAD_LIVCHANTIER'));
    TOBL.putValue('AFF_LIBELLE',TOBI.GetValue('AFF_LIBELLE'));
    TOBL.putValue('BAD_REMPLACEANT','X');
    TOBL.putValue('BAD_NUMREMPLACE',MaxRemplace);
    TOBL.putValue ('ARTICLEREMPLACE',TOBI.GetValue('BAD_CODEARTICLE'));
    TOBI.putValue('BAD_REMPLACE','X');
    TOBI.putValue('BAD_NUMREMPLACE',MaxRemplace);
    TOBI.putValue ('ARTICLEREMPLPAR',TOBL.GetValue('BAD_CODEARTICLE'));
  end;

  procedure SetInfoLigtoPeres (TOBL : TOB);
  var TOBP : TOB;
  begin
    if TOBL.parent = nil then exit;
    TOBP := TOBL.Parent;
    if not TOBP.FieldExists('BAD_FOURNISSEUR') then exit;
    if TOBP.detail.count > 1 then exit;
    TOBP.putValue('BAD_FOURNISSEUR',TOBL.GetValue('BAD_FOURNISSEUR'));
    TOBP.putValue('BAD_PABASE',TOBL.GetValue('BAD_PABASE'));
    if not TOBP.FieldExists ('T_LIBELLE') then TOBP.AddChampSup ('T_LIBELLE',false);
    TOBP.PutValue('T_LIBELLE',TOBL.getValue('T_LIBELLE'));
    SetInfoLigtoPeres (TOBP);
  end;


var TOBL,TOBI,TheParentInit : TOB;
    Ligne,NbLigne,NiveauMax  : integer;
    Article,depot,fournprinc : string;
    TarifFourn : TGinfostarifFour;
begin
  TOBI:=GetTOBDecisAch(TOBdecisionnel,GS.row) ; if TOBI=Nil then Exit ;
  //
  if GetArticle (Article,depot,fournprinc) then
  begin
    EnleveDesPeres  (TOBI,TOBI.parent);
   	TOBL := FindParent (Article,TOBI.GetValue('BAD_LIVCHANTIER'),Depot);
    // On rajoute les infos de provenance
    SetRemplaceLigne (TOBL,TOBI);
    if fournprinc <> '' then
    begin
      TOBL.putValue('BAD_FOURNISSEUR',Fournprinc );
      TarifFourn := GetInfosTarifAch (fournprinc,Article,TurAchat,false,true,nil,TOBL.GetValue('BAD_QUANTITEACH'),TOBL.GetValue('BAD_DATELIVRAISON'));
      TOBL.PutValue('BAD_PRIXACH',TarifFourn.TarifAchBrut);
      TOBL.PutValue('BAD_PRIXACHNET',CalculPrixNet(TarifFourn.TarifAchbrut,TarifFourn.Remise));
      TOBL.PutValue('BAD_CALCULREMISE',TarifFourn.Remise);
      TOBL.PutValue('BAD_CASCADEREMISE',TarifFourn.cascade);
      TOBL.PutValue('BAD_REMISE',TarifFourn.remisesreelle );
      TOBL.PutValue('BAD_TARIF',TarifFourn.tarif );
      TOBL.PutValue('BAD_COEFUAUS',TarifFourn.CoefUAUs);
      TOBL.PutValue('BAD_QUALIFQTEACH',TarifFourn.UniteAchat);
      if not TOBL.FieldExists ('T_LIBELLE') then TOBL.AddChampSup ('T_LIBELLE',false);
      TOBL.PutValue('T_LIBELLE',TarifFourn.LibelleFournisseur);
      SetInfoLigtoPeres (TOBL);
    end;
    //
  	TheParentInit := FindParentInit(TOBL);
    ReindiceDecisionnel;
    CumulStkPhysique (TheParentInit);
    AjouteAuxPeres (TOBL,TOBL.Parent);
    NbLigne := 0;
    NiveauMax := 4;
  	DefiniNiveauVisible (NiveauMax);
    ReIndiceGrille (TOBdecisionnel,NbLigne);
    AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,5,NbLigne);
    Ligne := TOBL.GetValue('NUMAFF');
    PositionneLigne (Ligne);
    AfficheInfos(Ligne);
 end;
end;

procedure TOF_DECISIONACHAT.NewbesoinClick(sender: TObject);
var TOBL,TheParentInit : TOB;
    Ligne,NbLigne  : integer;
    Article,depot,Fournprinc : string;
    TarifFourn : TGinfostarifFour;
begin
  if GetArticle (Article,depot,Fournprinc) then
  begin
  	TOBL := FindParent (Article,'-',Depot);
  	TheParentInit := FindParentInit(TOBL);
    ReindiceDecisionnel;
    AjouteAuxPeres (TOBL,TOBL.Parent);
    if fournprinc <> '' then
    begin
      TOBL.putValue('BAD_FOURNISSEUR',Fournprinc );
      TarifFourn := GetInfosTarifAch (fournprinc,Article,TurAchat,false,true,nil,TOBL.GetValue('BAD_QUANTITEACH'),TOBL.GetValue('BAD_DATELIVRAISON'));
      TOBL.PutValue('BAD_PRIXACH',TarifFourn.TarifAchBrut);
      TOBL.PutValue('BAD_PRIXACHNET',CalculPrixNet(TarifFourn.TarifAchbrut,TarifFourn.Remise));
      TOBL.PutValue('BAD_CALCULREMISE',TarifFourn.Remise);
      TOBL.PutValue('BAD_CASCADEREMISE',TarifFourn.cascade);
      TOBL.PutValue('BAD_REMISE',TarifFourn.remisesreelle );
      TOBL.PutValue('BAD_COEFUAUS',TarifFourn.CoefUAUs );
      TOBL.PutValue('BAD_TARIF',TarifFourn.tarif );
      if not TOBL.FieldExists ('T_LIBELLE') then TOBL.AddChampSup ('T_LIBELLE',false);
      TOBL.PutValue('T_LIBELLE',TarifFourn.LibelleFournisseur);
    end;
    CONSULTFOU.PositionneTheConsult (TheParentInit);
    CumulStkPhysique (TheParentInit);
    NbLigne := 0;
    NiveauMax := 4;
  	DefiniNiveauVisible (NiveauMax);
    ReIndiceGrille (TOBdecisionnel,NbLigne);
    AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,5,NbLigne);
    Ligne := TOBL.GetValue('NUMAFF');
    PositionneLigne (Ligne);
    AfficheInfos(Ligne);
  end;
end;

procedure TOF_DECISIONACHAT.ReactiveLigneBase (NumRemplace : integer);
var TOBL : TOB;
begin
	TOBL := TOBDecisionnel.findFirst(['BAD_NUMREMPLACE'],[Numremplace],true);
  if TOBL = nil then exit;
  TOBL.putValue('BAD_NUMREMPLACE',0);
  TOBL.putValue('ARTICLEREMPLACE','');
  TOBL.putValue('ARTICLEREMPLPAR','');
  TOBL.putValue('BAD_REMPLACE','-');
  AjouteAuxPeres (TOBL,TOBL.Parent);
end;

procedure TOF_DECISIONACHAT.MnSupLigneClick(Sender: Tobject);
var TOBL,NextTOBL,TOBParentInit,TheParent,TheTOBScindee : TOB;
    Ligne,NbLigne  : integer;
    NumRemplace: integer;
    Thereference : string;
begin
  TOBL:=GetTOBDecisAch(TOBdecisionnel,GS.row) ; if TOBL=Nil then Exit ;
  NumRemplace := TOBL.getValue('BAD_NUMREMPLACE');
  TheReference := TOBL.GetValue('BAD_REFGESCOM');;
  if (TOBL.GetValue('BAD_SCINDEE')='X') and (Thereference <> '') then
  begin
  	TheTOBScindee := TOBDecisionnel.findFirst(['BAD_REFGESCOM'],[Thereference],true);
    repeat
    	if (TheTOBScindee <> nil) and (TheTOBScindee <> TOBL) then break;
      if (TheTOBScindee <> nil) then TheTOBScindee := TOBDecisionnel.findNext(['BAD_REFGESCOM'],[Thereference],true);
    until TheTOBScindee = nil;
    if TheTOBScindee <> nil then
    begin
    	TheTOBScindee.putValue('BAD_SCINDEE','-');
    end;
  end;
	TOBParentInit := FindParentInit(TOBL);
  TheParent := TOBL.Parent;
  NextTOBL:=FindpreviousForDelete(TOBdecisionnel,GS.row) ; if NextTOBL=Nil then Exit ;
  Ligne := NextTOBL.GetValue('NUMAFF');
  //
	EnleveDesPeres  (TOBL,TOBL.parent);
  CumulStkPhysique (TOBparentInit);
  //
  TOBL.free;
  NettoieParent (TheParent);
  //
  if Numremplace <> 0 then
  begin
  	ReactiveLigneBase (NumRemplace);
  end;
  NbLigne := 0;
  ReIndiceGrille (TOBdecisionnel,NbLigne);
	AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,5,NbLigne);
  PositionneLigne (Ligne);
  if NextTOBL.GetValue('NIVEAU') = 4 then SetPrisEnSTockpapou (nextTOBL); // et on remonte
  AfficheInfos(Ligne);
end;

function TOF_DECISIONACHAT.FindpreviousForDelete(TOBdecisionnel: TOB;Ligne: integer): TOB;

	function IsPapounet (TOBAINSPECTER,TOBL : TOB) : boolean;
  var TOBP : TOB;
  begin
  	result := false;
    TOBP := TOBL.parent;
    if TOBP = TOBAINSPECTER then result := true;
    if (not result) and (TOBP.parent <> nil) then result := IsPapounet (TOBAINSPECTER,TOBP);
  end;

var TOBL,TOBPrev : TOB;
		found : boolean;
    Depart : integer;
begin
	TOBPrev := nil;
	result := nil;
	Depart := GS.row;
  TOBL:=GetTOBDecisAch(TOBdecisionnel,Depart) ; if TOBL=Nil then Exit ;
  found := false;
  repeat
  	dec(Depart);
    if Depart > 0 then
    begin
      TOBPrev := GetTOBDecisAch(TOBdecisionnel,Depart);
      if TOBPrev.getValue('NIVEAU')=TOBL.GetValue('NIVEAU') then BEGIN Found := true; break; END;
      if IsPapounet (TOBPrev,TOBL) then continue else BEGIN Found := true; Break; End;
    end;
  until (Found) or (Depart=0);
  IF not Found then
  begin
    repeat
      Inc(Depart);
      if Depart > 0 then
      begin
        TOBPrev := GetTOBDecisAch(TOBdecisionnel,Depart);
        if TOBPrev = nil then break;
        if TOBPrev.getValue('NIVEAU')=TOBL.GetValue('NIVEAU') then BEGIN Found := true; break; END;
        if IsPapounet (TOBPrev,TOBL) then continue else BEGIN Found := true; Break; End;
      end;
    until (Found) or (Depart=0);
  end;
  if Found then Result := TOBPrev;
end;

procedure TOF_DECISIONACHAT.ChangeFournisseur(Arow: integer;NewValue: String; var Cancel : boolean);

	function ExistFournisseur (Valeur : string) : boolean;
  var QQ : TQuery;
  begin
  	QQ := OPenSql ('SELECT T_LIBELLE FROM TIERS WHERE T_TIERS="'+Valeur+'" AND T_NATUREAUXI="FOU"',True,-1,'',true);
    result := not QQ.eof;
    ferme (QQ);
  end;

var TOBL : TOB;
		NBligne , Ligne : integer;
begin
	cancel := false;
	if (not ExistFournisseur(newvalue)) and (NewValue <> '') then
  begin
  	PGIBox(TraduireMemoire('Ce fournisseur n''existe pas'),ecran.caption);
    cancel := true;
    exit;
  end;
  TOBL := GetTOBDecisAch (TOBdecisionnel,Arow);
  Ligne := TOBL.GetValue('NUMAFF');
  TOBL.PutValue('BAD_FOURNISSEUR',NewValue);
  SetFournisseur (TOBL);
	NbLigne := GS.rowCount -1;
	AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,5,NbLigne,Ligne);
  AfficheInfos(Ligne);
end;

procedure TOF_DECISIONACHAT.SetFournisseur(TOBL: TOB);
var resultat : TGinfostarifFour;
		FUA,FUV,FUS,CoefuaUs,CoefUSUv : double;
    QQ : TQuery;
begin
	if TOBL.GetValue('BAD_FOURNISSEUR') <> '' then
  begin
    resultat := GetInfosTarifAch (TOBL.GetValue('BAD_FOURNISSEUR'),TOBL.GEtVAlue('BAD_ARTICLE'),
                                  TurStock,true,true,nil,TOBL.GetValue('BAD_QUANTITEACH'),
                                  TOBL.GetValue('BAD_DATELIVRAISON'));
    TOBL.PutValue('BAD_PRIXACH',resultat.TarifAchBrut);
    TOBL.PutValue('BAD_PRIXACHNET',resultat.TarifAch);
    TOBL.PutValue('BAD_CALCULREMISE',resultat.Remise);
    TOBL.PutValue('BAD_CASCADEREMISE',resultat.cascade);
    TOBL.PutValue('BAD_REMISE',resultat.remisesreelle );
    TOBL.PutValue('BAD_TARIF',resultat.tarif );
    if resultat.CoefUaUs  <> 0 then
    begin
      TOBL.PutValue('BAD_COEFUAUS',resultat.CoefUaUs );
      TOBL.PutValue('BAD_QUALIFQTEACH',resultat.UniteAchat );
    end;
    if not TOBL.FieldExists ('T_LIBELLE') then TOBL.AddChampSup ('T_LIBELLE',false);
    TOBL.PutValue('T_LIBELLE',resultat.LibelleFournisseur);
  end else
  begin
  	resultat.uniteAchat := '';
    resultat.TarifAchBrut := 0;
    resultat.remise := '';
    resultat.remisesreelle := 0;
    resultat.tarif  := 0;
    resultat.CoefuaUs  := 0;
    TOBL.PutValue('BAD_PRIXACH',TOBL.GetValue('BAD_PRIXACHNET'));
    TOBL.PutValue('BAD_CALCULREMISE','');
    TOBL.PutValue('BAD_REMISE',0 );
    TOBL.PutValue('BAD_TARIF',0 );
  	QQ := OpenSql ('SELECT GA_COEFCONVQTEACH,GA_QUALIFUNITEACH FROM ARTICLE WHERE GA_ARTICLE="'+TOBL.GEtVAlue('BAD_ARTICLE')+'"',true,1,'',true);
    if not QQ.eof then
    begin
      if QQ.Fields[1].AsString <> '' then
      begin
        TOBL.PutValue('BAD_COEFUAUS',QQ.Fields[0].AsFloat);
        TOBL.PutValue('BAD_QUALIFQTEACH',QQ.Fields[1].AsString);
      end else
      begin
        TOBL.PutValue('BAD_COEFUAUS',0);
        TOBL.PutValue('BAD_QUALIFQTEACH','');
      end;
    end else
    begin
    	TOBL.PutValue('BAD_COEFUAUS',0);
      TOBL.PutValue('BAD_QUALIFQTEACH','');
    end;
    ferme (QQ);
    TOBL.PutValue('T_LIBELLE','Aucun');
  end;

  FUV := RatioMesure('PIE', TobL.GetValue('BAD_QUALIFQTEVTE')); if FUV = 0 then FUV := 1;
  FUS := RatioMesure('PIE', TobL.GetValue('BAD_QUALIFQTESTO')); if FUS = 0 then FUS := 1;
  FUA := RatioMesure('PIE', TobL.GetValue('BAD_QUALIFQTEACH')); if FUA = 0 then FUA := 1;
  CoefuaUs := TobL.GetValue('BAD_COEFUAUS') ;if CoefuaUs = 0 then CoefuaUS := FUS/FUA;
  CoefUsUV := TobL.GetValue('BAD_COEFUSUV') ; if CoefuSUV = 0 then CoefuSUV := FUS/FUV;
  //
  TOBL.putValue('BAD_QUANTITEACH',Arrondi(TOBL.GetValue('BAD_QUANTITEVTE')/CoefUSUV /CoefuaUs,V_PGI.okdecQ));
  TOBL.putValue('BAD_QUANTITESTK',Arrondi(TOBL.GetValue('BAD_QUANTITEVTE')/CoefUSUV,V_PGI.okdecQ));

  if TOBL.detail.count > 0 then
  begin
  	SetInfosFournisseur2Fils (TOBL,FUV,FUS,FUA,CoefuaUs,coefUsUv);
  end;

end;

procedure TOF_DECISIONACHAT.SetInfosFournisseur2Fils(TOBP: TOB;FUV,FUS,FUA, CoefuaUS,CoefUSUV : double);
var Indice : integer;
		TOBL : TOB;
begin
	for Indice := 0 to TOBP.detail.count -1 do
  begin
  	TOBL := TOBP.detail[Indice];
    TOBL.PutValue('BAD_FOURNISSEUR',TOBP.GetValue('BAD_FOURNISSEUR'));
    TOBL.PutValue('BAD_QUALIFQTEACH',TOBP.GetValue('BAD_QUALIFQTEACH'));
    TOBL.PutValue('BAD_PRIXACH',TOBP.GetValue('BAD_PRIXACH'));
    TOBL.PutValue('BAD_PRIXACHNET',TOBP.GetValue('BAD_PRIXACHNET'));
    TOBL.PutValue('BAD_CALCULREMISE',TOBP.GetValue('BAD_CALCULREMISE'));
    TOBL.PutValue('BAD_CASCADEREMISE',TOBP.GetValue('BAD_CASCADEREMISE'));
    TOBL.PutValue('BAD_REMISE',TOBP.GetValue('BAD_REMISE'));
    TOBL.PutValue('BAD_TARIF',TOBP.GetValue('BAD_TARIF'));
    TOBL.PutValue('BAD_COEFUAUS',CoefUaUs);
    //
    TOBL.putValue('BAD_QUANTITEACH',Arrondi(TOBL.GetValue('BAD_QUANTITEVTE')/CoefUSUV/CoefUaUs,V_PGI.okdecQ));
    TOBL.putValue('BAD_QUANTITESTK',Arrondi(TOBL.GetValue('BAD_QUANTITEVTE')/CoefUsUV,V_PGI.okdecQ));

    if not TOBL.FieldExists ('T_LIBELLE') then TOBL.AddChampSup ('T_LIBELLE',false);
    TOBL.PutValue('T_LIBELLE',TOBP.GetValue('T_LIBELLE'));
    if TOBL.detail.count > 0 then SetInfosFournisseur2Fils (TOBL,FUV,FUS,FUA,CoefuaUs,CoefUSUV);
  end;
//
end;

procedure TOF_DECISIONACHAT.BAD_DATELIVRAISONClick(Sender: TObject);
var TOBL : TOB;
		NBligne , Ligne : integer;
begin
  TOBL := GetTOBDecisAch (TOBdecisionnel,GS.row);
  Ligne := TOBL.GetValue('NUMAFF');
  TOBL.PutValue('BAD_DATELIVRAISON', StrToDate(BAD_DATELIVRAISON.Text));
  SetDateLivraison (TOBL);
  if TOBL.detail.count > 0 then
  begin
  	SetInfosLivraison2Fils (TOBL);
  end;
	NbLigne := GS.rowCount -1;
	AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,5,NbLigne);
  AfficheInfos(Ligne);
end;

procedure TOF_DECISIONACHAT.SetDateLivraison(TOBL: TOB);
var resultat : TGinfostarifFour;
begin
	resultat := GetInfosTarifAch (TOBL.GetValue('BAD_FOURNISSEUR'),TOBL.GEtVAlue('BAD_ARTICLE'),
  															TurAchat,true,true,nil,TOBL.GetValue('BAD_QUANTITEACH'),
                                TOBL.GetValue('BAD_DATELIVRAISON'));
  TOBL.PutValue('BAD_QUALIFQTEACH',resultat.UniteAchat);
  TOBL.PutValue('BAD_COEFUAUS',resultat.CoefUAUs);
  TOBL.PutValue('BAD_PRIXACH',resultat.TarifAchBrut);
  TOBL.PutValue('BAD_PRIXACHNET',resultat.TarifAch);
  TOBL.PutValue('BAD_CALCULREMISE',resultat.Remise);
  TOBL.PutValue('BAD_CASCADEREMISE',resultat.cascade );
  TOBL.PutValue('BAD_REMISE',resultat.remisesreelle);
  TOBL.PutValue('BAD_TARIF',resultat.Tarif);

end;

procedure TOF_DECISIONACHAT.SetInfosLivraison2Fils(TOBP: TOB);
var Indice : integer;
		TOBL : TOB;
begin
	for Indice := 0 to TOBP.detail.count -1 do
  begin
  	TOBL := TOBP.detail[Indice];
    TOBL.PutValue('BAD_FOURNISSEUR',TOBP.GetValue('BAD_FOURNISSEUR'));
    TOBL.PutValue('BAD_QUALIFQTEACH',TOBP.GetValue('BAD_QUALIFQTEACH'));
    TOBL.PutValue('BAD_COEFUAUS',TOBP.GetValue('BAD_COEFUAUS'));
    TOBL.PutValue('BAD_PRIXACH',TOBP.GetValue('BAD_PRIXACH'));
    TOBL.PutValue('BAD_PRIXACHNET',TOBP.GetValue('BAD_PRIXACHNET'));
    TOBL.PutValue('BAD_CALCULREMISE',TOBP.GetValue('BAD_CALCULREMISE'));
    TOBL.PutValue('BAD_DATELIVRAISON',TOBP.GetValue('BAD_DATELIVRAISON'));
    TOBL.PutValue('BAD_CASCADEREMISE',TOBP.GetValue('BAD_CASCADEREMISE'));
    TOBL.PutValue('BAD_REMISE',TOBP.GetValue('BAD_REMISE'));
  	TOBL.PutValue('BAD_TARIF',TOBP.GetValue('BAD_TARIF'));
  end;
//
end;

procedure TOF_DECISIONACHAT.BViewDetailClick(Sender: TOBject);
  Var SD: TSaveDialog;
begin
	FenetreSup.visible := BviewDetail.Down;
end;

procedure TOF_DECISIONACHAT.FenetreSupClose(Sender: TObject);
begin
	BviewDetail.Down := false;
end;

procedure TOF_DECISIONACHAT.MajDecisionnel;
begin
	SupprimeDecisionnelAch;
  CONSULTFOU.Supprime;
  RenumeroteDecisonnel;
  if V_PGI.IOError = OeOk then InscritDecisionnel;
  if V_PGI.IOError = OeOk then CONSULTFOU.ecrit;
end;

procedure TOF_DECISIONACHAT.ChangePrixNet(Arow: integer; NewValue: double);
var TOBL : TOB;
		NBligne , Ligne : integer;
begin
  TOBL := GetTOBDecisAch (TOBdecisionnel,Arow);
  Ligne := TOBL.GetValue('NUMAFF');
  TOBL.PutValue('BAD_PRIXACHNET',NewValue);
  TOBL.PutValue('BAD_CALCULREMISE','');
  TOBL.PutValue('BAD_PRIXACH',NewValue);
  TOBL.PutValue('BAD_CASCADEREMISE','');
  TOBL.PutValue('BAD_REMISE',0);
  TOBL.PutValue('BAD_TARIF',0);
  if TOBL.Detail.count > 0 then SetPrixAchNet (TOBL);
	NbLigne := GS.rowCount -1;
	AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,5,NbLigne,Arow);
  AfficheInfos(Ligne);
end;

procedure TOF_DECISIONACHAT.SetPrixAchNet(TOBP: TOB);
var Indice : integer;
		TOBL : TOB;
begin
	for Indice := 0 to TOBP.detail.count -1 do
  begin
  	TOBL := TOBP.detail[Indice];
    TOBL.PutValue('BAD_PRIXACHNET',TOBP.GetValue('BAD_PRIXACHNET'));
    TOBL.PutValue('BAD_CALCULREMISE','');
  	TOBL.PutValue('BAD_PRIXACH',TOBP.GetValue('BAD_PRIXACHNET'));
    TOBL.PutValue('BAD_CASCADEREMISE','');
    TOBL.PutValue('BAD_REMISE',0);
  	TOBL.PutValue('BAD_TARIF',0);
    if TOBL.detail.count > 0 then SetPrixAchNet (TOBL);
  end;
//
end;

procedure TOF_DECISIONACHAT.InscritDecisionnel;
begin
	TOBdecisionnel.SetAllModifie (true);
  if not TOBDecisionnel.InsertDB (nil,true) then V_PGI.IOError := OeUnknown;
end;

procedure TOF_DECISIONACHAT.RenumeroteDecisonnel;
begin
	NumeroteLigne(ecran,TOBDecisionnel);
	AppliqueNumerotationLigne (TOBDecisionnel,StrToInt(NumDecisionnel));
end;

function TOF_DECISIONACHAT.RechercheAffaire (var TheAffaire : string): boolean;
var Aff0,Aff1, Aff2, Aff3, Aff4,Tiers : THCritMaskEdit;
		A0,A1,A2,A3,A4 : string;
begin
  Aff0 := THCritMaskEdit.Create (ecran);
  Aff1 := THCritMaskEdit.Create (ecran);
  Aff2 := THCritMaskEdit.Create (ecran);
  Aff3 := THCritMaskEdit.Create (ecran);
  Aff4 := THCritMaskEdit.Create (ecran);
  Tiers := THCritMaskEdit.Create (ecran);
  TRY
  	BTPCodeAffaireDecoupe(TheAffaire, A0, A1, A2, A3, A4, taCreat, false);
    Aff0.Text := A0;
    Aff1.Text := A1;
    Aff2.Text := A2;
    Aff3.Text := A3;
    Aff4.Text := A4;
  	Tiers.Text := '';
  	result := GetAffaireEnteteSt(Aff0, Aff1, Aff2, Aff3, Aff4, Tiers, TheAffaire, false, false, false, true, true,'FAC');
  FINALLY
  	Tiers.free;
    Aff0.free;
    Aff1.free;
    Aff2.free;
    Aff3.free;
    Aff4.free;
  END;
end;

function TOF_DECISIONACHAT.TransformeRemises(REMISES: string) : string;
begin
  result := StringReplace (Remises,'%','',[rfReplaceAll]);
  result := StringReplace (result,'.',',',[rfReplaceAll]);
end;


function TOF_DECISIONACHAT.CalculPrixNet : double;
var RemiseResult : double;
begin
  RemiseResult := RemiseResultante(TransformeRemises(REMISES.text));
  result := arrondi(PRIXACHATBRUT.Value * (1-(RemiseResult/100)),V_PGI.OkDecP );
end;

procedure TOF_DECISIONACHAT.REMISESChange(Sender: TOBject);

  procedure SetPrixAchremises (TOBP : TOB);
  var Indice : integer;
      TOBL : TOB;
  begin
    for Indice := 0 to TOBP.detail.count -1 do
    begin
      TOBL := TOBP.detail[Indice];
      TOBL.PutValue('BAD_PRIXACHNET',TOBP.GetValue('BAD_PRIXACHNET'));
      TOBL.PutValue('BAD_CALCULREMISE',TOBP.GetValue('BAD_CALCULREMISE'));
      TOBL.PutValue('BAD_PRIXACH',TOBP.GetValue('BAD_PRIXACHNET'));
      TOBL.PutValue('BAD_CASCADEREMISE',TOBP.GetValue('BAD_CASCADEREMISE'));
      TOBL.PutValue('BAD_REMISE',TOBP.GetValue('BAD_REMISE'));
  		TOBL.PutValue('BAD_TARIF',TOBP.GetValue('BAD_TARIF'));
      if TOBL.detail.count > 0 then SetPrixAchremises (TOBL);
    end;
  end;
var
		TOBL : TOB;
    Ligne : integer;
begin
  TOBL := GetTOBDecisAch (TOBdecisionnel,GS.Row);
  Ligne := TOBL.GetValue('NUMAFF');
  PRIXACHATNET.Value := CalculPrixNet;
  TOBL.putValue('BAD_CALCULREMISE',REMISES.Text);
  TOBL.PutValue('BAD_PRIXACH',PRIXACHATBRUT.Value);
  TOBL.PutValue('BAD_PRIXACHNET',PRIXACHATNET.Value);
  TOBL.PutValue('BAD_REMISE',RemiseResultante(TransformeRemises(REMISES.text)));
  //
  LastPrixNet := PRIXACHATNET.Value;
  //
  if TOBL.Detail.count > 0 then SetPrixAchremises (TOBL);
	NbLigne := GS.rowCount -1;
	AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,5,NbLigne);
  AfficheInfos(Ligne);
  //
end;


Procedure TOF_DECISIONACHAT.PRIXACHATNETEnter (Sender : Tobject);
begin
	LastPrixNet := PRIXACHATNET.Value;
end;

Procedure TOF_DECISIONACHAT.PRIXACHATNETChange (Sender : Tobject);
begin
	if arrondi(LastPrixNet,v_PGI.okdecp) <> Arrondi(PRIXACHATNET.Value,V_PGI.okdecP) then
  begin
  	ChangePrixNet (GS.row,Valeur(PRIXACHATNET.Text));
  end;
end;

procedure TOF_DECISIONACHAT.SetEcran;
begin
  if action = taconsult then
  begin
  	GS.PopupMenu := nil;
    GS.options := GS.OPtions - [goediting,GoAlwaysShowEditor] + [GoRowSelect];
  end;
  if BOPTIONMULTISEL <> nil then BOPTIONMULTISEL.Visible := false;
end;

procedure TOF_DECISIONACHAT.ASupprimerClick(Sender: Tobject);

	procedure SupprimeLigne (TOBL : TOB);
  var TheLigne : integer;
  		Indice : integer;
      TOBS : TOB;
  begin
  	TheLigne := TOBL.GetValue('NUMAFF');
    if ASupprimer.Checked then TOBL.putValue('BAD_SUPPRIME','X')
                          else TOBL.putValue('BAD_SUPPRIME','-');
    if TheLIgne > 0 then
    begin
      TOBL.PutLigneGrid (GS,TheLigne,false,false,ListeSaisie);
    end;
    if TOBL.detail.count > 0 then
    begin
      for indice := 0 to TOBL.detail.count -1 do
      begin
        TOBS := TOBL.detail[Indice];
        SupprimeLigne (TOBS);
      end;
    end;
	end;

var TOBL : TOB;
begin
  TOBL := GetTOBDecisAch (TOBdecisionnel,GS.Row);
  SupprimeLigne (TOBL);
end;

function TOF_DECISIONACHAT.FindEtablissement: string;
var QQ : TQuery;
		TOBETAB : TOB;
begin
	TOBETAB := TOB.create ('ETABLISS',nil,-1);
  QQ := OpenSql ('SELECT * FROM ETABLISS',true,-1,'',true);
  if not QQ.eof then
  begin
    TOBETab.LoadDetailDB ('ETABLISS','','',QQ,false);
    result := TOBEtab.detail[0].getValue('ET_ETABLISSEMENT');
  end;
  Ferme (QQ);
  TOBEtab.free;
end;


procedure TOF_DECISIONACHAT.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = VK_RETURN) then
  begin
		if ecran.ActiveControl = GS then Key := VK_TAB;
  end;
  // Ctrl + z (deselectionne tout)
  if OkConsultFou then
  begin
    if (Key = 90) and ( ssCtrl in Shift) then
    begin
       SelectionneTout (false);
    end;
    // Ctrl + a (selectionne tout)
    if (Key = 65) and (ssCtrl in Shift) then
    begin
      SelectionneTout (True);
    end;
  end;
end;

procedure TOF_DECISIONACHAT.GSDblClick(Sender: TOBject);
begin
  if GS.Col = GS_FOURNISSEUR then
  begin
  	GSElipsisClick (self);
  end;
end;

procedure TOF_DECISIONACHAT.BFindAssociationClick(Sender: Tobject);
var TOBL,TOBR : TOB;
		Numero : integer;
    Ligne : integer;
    found : boolean;
begin
  TOBL := GetTOBDecisAch (TOBdecisionnel,GS.Row); if TOBL = nil then exit;
  Numero := TOBL.GetValue('BAD_NUMREMPLACE');
  if TOBL.GetValue('BAD_REMPLACE')='X' then
  begin
  	TOBR := TOBdecisionnel.findFirst(['BAD_NUMREMPLACE','BAD_REMPLACEANT'],[Numero,'X'],true);
  end else if TOBL.GetValue('BAD_REMPLACEANT')='X' then
  begin
  	TOBR := TOBdecisionnel.findFirst(['BAD_NUMREMPLACE','BAD_REMPLACE'],[Numero,'X'],true);
  end;
  if TOBR <> nil then
  begin
    Ligne := TOBR.GetValue('NUMAFF');
    if Ligne = 0 then
    begin
    	NiveauMax := 5;
			OuvreTouteLaBranche (GS,TOBR,TOBdecisionnel,nBLigne,niveauMax,found,True);
      Ligne := TOBR.GetValue('NUMAFF');
      DefiniNiveauVisible (NiveauMax);
      AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,NiveauMax,NbLigne,Ligne);
    end;
    PositionneLigne (Ligne);
  end;
end;

procedure TOF_DECISIONACHAT.BrechArticleClick(Sender: TOBject);
begin
  FirstFind := true;
  FindDialog.Execute;
end;

procedure TOF_DECISIONACHAT.FindDialogFind(Sender: TObject);
var TheString : string;
begin
  Rechercher(GS, FindDialog, FirstFind);
  gs.HideEditor;
  PositionneLigne (GS.Row);
  gs.ShowEditor;
end;

procedure TOF_DECISIONACHAT.MnFournisseurClick (Sender : TObject);
var TOBL : TOB;
		fournisseur,Article,MaCle : string;
    cancel : boolean;
begin
  TOBL := GetTOBDecisAch (TOBdecisionnel,GS.row);
(*  fournisseur := TOBL.GetValue('BAD_FOURNISSEUR');
  Article := TOBL.GetValue('BAD_ARTICLE');
*)
  MaCle := '';
  MaCle := AGLLanceFiche('GC', 'GCFOURNISSEUR_MUL', 'T_NATUREAUXI=FOU', '', 'SELECTION');
  if (MaCle<>'') and (MaCle <> TOBL.GEtValue('BAD_FOURNISSEUR')) then
  begin
    ChangeFournisseur (GS.row,MaCle,cancel);
    CellCur := GS.Cells [GS.Col,GS.row];
  end;
end;

procedure TOF_DECISIONACHAT.MnCatalogueClick (Sender : TObject);
var TOBL : TOB;
		fournisseur: string;
    cancel : boolean;
begin
  TOBL := GetTOBDecisAch (TOBdecisionnel,GS.row);
  Fournisseur := LookUpFournisseur (GS.Cells [GS.Col,GS.row],TOBL.GEtValue('BAD_ARTICLE'),TOBL.GEtValue('GA_LIBELLE'),true);
  if Fournisseur <> '' then
  begin
    if Fournisseur <> TOBL.GEtValue('BAD_FOURNISSEUR') then
    begin
      ChangeFournisseur (GS.row,Fournisseur,cancel);
      CellCur := GS.Cells [GS.Col,GS.row];
    end;
  end;
end;

procedure TOF_DECISIONACHAT.GSMouSeDown (Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
	GSX := X;
  GSY := Y;
  if OkConsultFou then
  begin
    if (Button = mbLeft) and (ssCtrl in Shift) and (not (ssShift in Shift)) and (action <> taconsult) then
    begin
      FlipThisLigne (GS.row);
    end;
    if (Button = mbLeft) and (ssCtrl in Shift) and (ssShift in Shift) and (action <> taconsult) then
    begin
      FlipOverThisLigne (GS.row);
    end;
  end;
end;

function TOF_DECISIONACHAT.IsBlocNoteVide(BlocNote: String): boolean;
begin
	result := true;
	StringToRich(fTexte, BlocNote);
  if (Length(fTexte.Text) <> 0) and (ftexte.Text <> #$D#$A) then result := false;
end;

procedure TOF_DECISIONACHAT.DeplierLigne (TOBL : TOB; Open : boolean=true);
var Niveau, Indice: integer;
		TOBS : TOB;
    Signe : string;
begin
	if Open then Signe := 'X' else Signe := '-';
  // --
  if TOBL=Nil then Exit ;
  Niveau := TOBL.GEtValue('NIVEAU');

  if Niveau = 1 then
  begin
  	TOBL.PutValue('OPEN',Signe);
  end else if niveau = 2 then
  begin
    TOBL.PutValue('OPEN',Signe);
    TOBL.PutValue('OPEN1',Signe);
  end else if niveau = 3 then
  begin
    TOBL.PutValue('OPEN',Signe);
    TOBL.PutValue('OPEN1',Signe);
    TOBL.PutValue('OPEN2',Signe);
  end ;
  //
  if not Open then TOBL.PutValue('NUMAFF',0);
  //
  if TOBL.detail.count > 0 then
  begin
  	for Indice :=0 to  TOBL.detail.count -1 do
    begin
    	TOBS := TOBL.detail[Indice];
    	DeplierLigne (TOBS,Open);
    end;
  end;

end;

procedure TOF_DECISIONACHAT.MnDeplierClick(Sender: TObject);
var indice : integer;
		Nbligne : integer;
    TOBL : TOB;
begin
	for Indice := 0 to TobDecisionnel.detail.count -1 do
  begin
    TOBL:=TOBdecisionnel.detail[Indice] ; if TOBL=Nil then break ;
    DeplierLigne (TOBL);
  end;
  NbLigne := 0;
  ReIndiceGrille (TOBdecisionnel,NbLigne);
  NiveauMax := MaxNiveauAff;
  DefiniNiveauVisible (NiveauMax);
	AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,NiveauMax,NbLigne);
  PositionneLigne (1);
end;

procedure TOF_DECISIONACHAT.MnRefermerClick(Sender: TObject);
var indice : integer;
		Nbligne : integer;
    TOBL : TOB;
begin
	NbLigne := 1;
	for Indice := 0 to TobDecisionnel.detail.count -1 do
  begin
    TOBL:=TOBdecisionnel.detail[Indice] ; if TOBL=Nil then break ;
    DeplierLigne (TOBL,false);
    TOBL.putValue('NUMAFF',NbLIgne); inc(NbLIgne);
  end;
  NiveauMax := 1;
  DefiniNiveauVisible (NiveauMax);
	AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,NiveauMax,NbLigne-1);
  PositionneLigne (1);
end;

function TOF_DECISIONACHAT.GetLigne: integer;
begin
  result := GS.Row;
end;

procedure TOF_DECISIONACHAT.PositionneDeciFournisseur(const Value: String);
var cancel : boolean;
begin
  ChangeFournisseur (GS.row,value,cancel);
end;

procedure TOF_DECISIONACHAT.PositionneDeciPrix(const Value: double);
var TOBL : TOB;
		FUA,FUS : double;
    CoefUaUs : double;
    ValueREs : double;
begin
  TOBL := GetTOBDecisAch (TOBdecisionnel,GS.row);
  CoefUaUs := TOBL.getValue('BAD_COEFUAUS');
  FUS := RatioMesure('PIE', TOBL.GetValue('BAD_QUALIFQTESTO')); if FUS = 0 then FUS := 1;
  FUA := RatioMesure('PIE', TOBL.GetValue('BAD_QUALIFQTEACH')); if FUA = 0 then FUA := 1;
  if CoefUaUs <> 0 then
  begin
  	ValueREs := Arrondi(Value / CoefUaUs,V_PGI.okdecP);
  end else
  begin
  	ValueREs := Arrondi(Value / FUA * FUS,V_PGI.okdecP);
  end;
  ChangePrixNet (GS.row,ValueREs);
end;

procedure TOF_DECISIONACHAT.PositionneSelectUp (TOBL : TOB);
var comptage,Indice : integer;
    TOBP,TOBLS : TOB;
begin
  TOBP := TOBL.Parent; if TOBP = nil then exit;
  if TOBP.NomTable = 'DECISIONACH' then exit;
  Comptage := 0;
  for Indice := 0 To TOBP.detail.count -1 do
  begin
    TOBLS := TOBP.detail[Indice];
    if (TOBLS <> TOBL) and (TOBLS.GetValue('SELECT')='-') Then Inc(Comptage);
  end;
  if Comptage = 0 then
  begin
    if TOBP.FieldExists ('SELECT') then
    begin
      TOBP.putValue('SELECT','X');
    end;
    PositionneSelectUp (TOBP);
  end;
end;

procedure TOF_DECISIONACHAT.FlipThisLigne(Ligne: integer);
var TOBL : TOB;
begin
  //
  TOBL := GetTOBDecisAch (TOBdecisionnel,Ligne);
  PositionneSelectionBranche (TOBL);
  GS.Invalidate; // rafraichissement
end;

procedure TOF_DECISIONACHAT.PositionneSelectBas (TOBL : TOB; Etat : boolean);
var Indice : integer;
    Status : string;
begin
  if etat then Status := 'X' else Status := '-';
  if TOBL.FieldExists ('SELECT') then
  BEGIN
    TOBL.putValue('SELECT',Status);
  END;

  if TOBL.detail.count > 0 then
  begin
    for Indice := 0 to TOBL.detail.count -1 do
    begin
      PositionneSelectBas (TOBL.detail[Indice],Etat);
    end;
  end;
end;

procedure TOF_DECISIONACHAT.DeselectionneUp (TOBL: TOB);
var comptage,Indice : integer;
    TOBP,TOBLS : TOB;
begin
  TOBP := TOBL.Parent; if TOBP = nil then exit;
  if TOBP.NomTable = 'DECISIONACH' then exit;
  TOBP.putValue('SELECT','-');
  DeselectionneUp (TOBP);
end;


procedure TOF_DECISIONACHAT.DeSelectionne (TOBL : TOB);
begin
  PositionneSelectBas (TOBL,False);
  DeselectionneUp (TOBL);
end;

procedure TOF_DECISIONACHAT.PositionneSelectionBranche(TOBL: TOB);
begin
  if TOBL.GetValue('SELECT')='-' then
  begin
    //
    LastSelect := TOBL.GetValue('NUMAFF');
    //
    PositionneSelectBas (TOBL,True);
    PositionneSelectUp (TOBL);
  end else
  begin
    DeSelectionne (TOBL);
  end;
  PositionneOptions;
end;

procedure TOF_DECISIONACHAT.SelectionneTout (Status : boolean);
begin
  PositionneSelectBas (TOBDecisionnel,Status);
  PositionneOptions;
  GS.Invalidate; // rafraichissement
end;

function TOF_DECISIONACHAT.IsSelectionne (TOBCurr : TOB) : boolean;
var TOBF : TOB;
    Indice : integer;
begin
  result := false;
  if TOBCurr.GetValue('SELECT')='X' Then BEGIN result := true; Exit; END;
  For Indice := 0 to TOBCurr.detail.count -1 do
  begin
    result := IsSelectionne(TOBCurr.detail[Indice]);
    if result then break;
  end;
end;

procedure TOF_DECISIONACHAT.SetmenuOptions(Status : boolean);
begin
  if MnOptions <> nil then MnOptions.enabled := status;
end;

procedure TOF_DECISIONACHAT.PositionneOptions;
var Indice : integer;
    OkEnabled : boolean;
begin
  OkEnabled := true;
  for Indice := 0 to TOBDecisionnel.detail.count -1 do
  begin
    If IsSelectionne(TOBDecisionnel.detail[Indice]) then BEGIN OkEnabled := false; break; END;
  end;
  GOptions.Enabled := OkEnabled;
  SetmenuOptions(not OkEnabled);
  if BOPTIONMULTISEL <> nil then BOPTIONMULTISEL.Visible := (not OkEnabled);
end;

procedure TOF_DECISIONACHAT.AddlesZoneOptions (TOBO : TOB);
begin
  TOBO.AddChampSupValeur ('GERESTOCK','');
  TOBO.AddChampSupValeur ('LIVCHANTIER','');
  TOBO.AddChampSupValeur ('FOURN','');
  TOBO.AddChampSupValeur ('CONSULTFOU','-');
  TOBO.AddChampSupValeur ('CHGOPT','-');
  TOBO.AddChampSupValeur ('DEPOT','');
end;

procedure TOF_DECISIONACHAT.MnOptionsClick (Sender : TObject);
var TOBOptions : TOB;
begin
  TOBOptions := TOB.Create ('LES OPTIONS',nil,-1);
  AddlesZoneOptions (TOBOptions);
  if OkConsultFou then TOBOptions.putValue('CONSULTFOU','X');
  TheTOB := TOBOptions;
  TheTOB.data := TOBDecisionnel;
  AGLLanceFiche ('BTP','BTDECISIONACHOPT','','','');
  if TheTOB <> nil then
  begin
    // traitement des Options
    if TOBOptions.detail.count > 0 then
    begin
      CONSULTFOU.MultipleAffecte (TOBOptions);
      GS.Invalidate;
    end else
    begin
      if TOBOptions.getValue('FOURN') <> '' then MultipleAffecteFournisseur (TOBOptions.getValue('FOURN'));
      if TOBOptions.getValue('DEPOT') <> '' then MultipleAffecteDepot (TOBOptions.getValue('DEPOT'));
      //
      if TOBOptions.getValue('CHGOPT') ='X' then
      begin
        MultipleChangeOptions (TOBOptions.getValue('LIVCHANTIER'),TOBOptions.getValue('GERESTOCK'));
      end;
      //
    end;
  end;
  TheTOB := nil;
  TOBOptions.free;
  SelectionneTout (false);  // deselectionne
end;

function TOF_DECISIONACHAT.FindSelected(TOBCurr: TOB): TOB;
var TOBF : TOB;
    Indice : integer;
begin
  result := GetSelectionne (TOBCurr);
end;

procedure TOF_DECISIONACHAT.MultipleAffecteFournisseur(LeFournisseur: string);
var Indice : integer;
    TOBD : TOB;
begin
  // provenance de la multi selection
  for Indice := 0 to TOBdecisionnel.detail.count -1 do
  begin
    TOBD:= FindSelected (TOBDecisionnel.detail[Indice]);
    if TOBD <> nil then
    begin
      AffecteFournisseur (TOBD,LeFournisseur);
    end;
  end;
	NbLigne := GS.rowCount -1;
	AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,5,NbLigne);
end;

procedure TOF_DECISIONACHAT.AffecteFournisseur(TOBD: TOB;LeFournisseur: string);
begin
  TOBD.PutValue('BAD_FOURNISSEUR',LeFournisseur);
  SetFournisseur (TOBD);
end;

procedure TOF_DECISIONACHAT.MultipleAffecteDepot(LeDepot: string);
var Indice : integer;
    TOBD : TOB;
begin
  // provenance de la multi selection
  for Indice := 0 to TOBdecisionnel.detail.count -1 do
  begin
    TOBD:= FindSelected (TOBDecisionnel.detail[Indice]);
    if TOBD <> nil then
    begin
      AffecteDepot (TOBD,LeDepot);
    end;
  end;
	NbLigne := GS.rowCount -1;
	AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,5,NbLigne);
end;

procedure TOF_DECISIONACHAT.AffecteDepot(TOBD: TOB;LeDepot: string);
begin
  SetDepot (TOBD, LeDepot);
end;

procedure TOF_DECISIONACHAT.SetChangementOption (TOBD : TOB;LivChantier,GereStock : string; Position : integer);
var Indice : integer;
    changed : boolean;
    TOBLocate,TOBP : TOB;
    Niveau : integer;
    Article,LivCha,Depot : string;

begin

  if TOBD.GetValue('NIVEAU') <> 4 then
  begin
    Indice := 0;
    For Indice := 0 to TOBD.detail.count -1 do
    begin
      SetChangementOption(TOBD.detail[Indice],LivChantier,GereStock,Indice);
    end;
    TOBD.putValue('SELECT','-');
  end else
  begin
    TOBLocate := TOBD.parent;     // parent dans la branche dupliquee
    Niveau := TOBLocate.GetValue('NIVEAU');
    Article := TOBLocate.GEtValue('BAD_ARTICLE');
    Depot := TOBLocate.GEtValue('BAD_DEPOT');
    Livcha := TOBLocate.GEtValue('BAD_LIVCHANTIER');
    TOBP := TrouveLigneinDecisionnel (TOBdecisionnel,Niveau,Article,Livcha,depot);  // on cherche dans le decisionnel la ligne concernee
    TOBLocate := TOBP.detail[Position];  // ligne qui se trouve dans le decisionnel

    changed := false;
    if Livchantier='X' then
    begin
      if TOBD.GetValue('BAD_LIVCHANTIER')<>'X' then changed := true;
      if changed then TOBD.putValue('BAD_LIVCHANTIER','X');
    end else if LivCHantier='-' then
    begin
      if TOBD.GetValue('BAD_LIVCHANTIER')<>'-' then changed := true;
      if changed then TOBD.putValue('BAD_LIVCHANTIER','-');
    end;
    //
    if GereStock='X' then
    begin
      if changed then TOBD.putValue('BAD_TENUESTOCK','X') else TOBLocate.putValue('BAD_TENUESTOCK','X');
    end else if gereStock='-' then
    begin
      if changed then TOBD.putValue('BAD_TENUESTOCK','-') else TOBlocate.putValue('BAD_TENUESTOCK','-');
    end;

    if changed then TOBD.Putvalue ('SELECT','X') else TOBlocate.putValue('SELECT','-');

    // comme on a changé la clef...
    if changed then
    begin
      MepLigneInDecisionnel (TOBD,false); // on duplique
      TOBLocate.PutValue ('_A_DELETE','X');
    end;
    //
  end;
end;

procedure TOF_DECISIONACHAT.NettoieApresModif (TOBD : TOB; var Indice : integer); // dans la descente
var LocIndice : integer;
    changed : boolean;
    TOBLocate : TOB;
begin
  if TOBD.GetValue('NIVEAU') <> 4 then
  begin
    LocIndice := 0;
    if TOBD.detail.count = 0 then exit;
    repeat
      NettoieApresModif(TOBD.detail[LocIndice],Locindice);
    until locIndice >= TOBD.detail.count;
    inc(indice);
  end else
  begin
    if TOBD.GetValue ('_A_DELETE')='X' then
    begin
      TOBD.free;
    end else
    begin
      if TOBD.GetValue('_A_TRAITER')='-' then
      begin
        TOBD.PutValue('_A_TRAITER','X'); // restitution
      end;
      Inc(Indice);
    end;
  end;
end;

procedure TOF_DECISIONACHAT.SetChangeOptions (TOBD : TOB;LivraisonChantier,GereStock : string);
var Ligne : Integer;
		TheParentInit,NewFirstParent : TOB;
    Niveau : integer;
    Article,LivChantier,Depot : string;
    Indice : integer;
    TOBReference : TOB;
begin
  Indice := 0;
	Niveau := TOBD.GetValue('NIVEAU');
  Article := TOBD.GEtValue('BAD_ARTICLE');
  Depot := TOBD.GEtValue('BAD_DEPOT');
  //
  TheParentInit := FindFirstParent(TOBD);
  if TheParentInit <> nil then ReinitNiveauInf (TheParentInit);
  //
  TOBReference := TOB.Create('DECISIONACHLIG',nil,-1);
  AddlesChampsSup(TOBReference,MaxRemplace);
  TOBReference.Dupliquer (TOBD,true,true);
  //
  SetChangementOption (TOBreference,LivraisonChantier,GereStock,0); // dans la descente
  TOBreference.free;
  //
  Indice := 0;
  NettoieApresModif (TOBD,Indice); // dans la descente
  //
  if TheParentInit <> nil then NettoieFilles (TheParentInit);
  if TheParentInit <> nil then recalculelesNiveaux (TheParentInit);
  if TheParentInit <> nil then CumulStkPhysique (TheParentInit);
end;

procedure TOF_DECISIONACHAT.MultipleChangeOptions (LivraisonChantier,GereStock: string);
var Indice : integer;
    First : boolean;
    Niveau : integer;
    Article,LivChantier : string;
    Depot : String;
    TOBD,TOBL: TOB;
    Ligne : integer;
begin
  if (LivraisonChantier='') and (GereStock='') then exit;
  first := true;
  for Indice := 0 to TOBdecisionnel.detail.count -1 do
  begin
    TOBD:= FindSelected (TOBDecisionnel.detail[Indice]);
    if TOBD <> nil then
    begin
      if First then
      BEGIN
        Niveau := TOBD.GetValue('NIVEAU');
        Article := TOBD.GEtValue('BAD_ARTICLE');
        Depot := TOBD.GEtValue('BAD_DEPOT');
      END;
      if LivraisonChantier='' then
      Begin
        LivCHantier := TOBD.GetValue('BAD_LIVCHANTIER');
      end else
      begin
        LivCHantier := LivraisonChantier;
      end;
      SetChangeOptions (TOBD,LivraisonChantier,GereStock);
      if FIrst then
      begin
        if Niveau < 4 then
        begin
          // on a peut etre change les clefs donc il faut rechercher le nouvel emplacement
          TOBL := TrouveLigneinDecisionnel (TOBDecisionnel,Niveau,Article,LivCHantier,Depot);
        end;
        first := false;
      end;
      CONSULTFOU.PositionneTheConsult (TOBD);

    end;
  end;
  ReindiceDecisionnel;
  NbLigne := 0;
  ReIndiceGrille (TOBdecisionnel,NbLigne);
  AfficheLaGrille (ecran,ListeSaisie,TOBdecisionnel,5,NbLigne);
  Ligne := TOBL.GetValue('NUMAFF');
  PositionneLigne (Ligne);
  AfficheInfos(Ligne);
end;

procedure TOF_DECISIONACHAT.FlipOverThisLigne(Ligne: integer);
var Indice : integer;
    TOBL : TOB;
begin
  if Ligne > LastSelect then
  begin
    for Indice:= LastSelect to Ligne do
    begin
      TOBL:=GetTOBDecisAch(TOBDecisionnel,Indice) ; if TOBL=Nil then Exit ;
      // Selectionnelaligne
      PositionneSelectBas (TOBL,True);
      PositionneSelectUp (TOBL);
    end;
    GS.Invalidate; // rafraichissement
  end;
end;

procedure TOF_DECISIONACHAT.VOIRESTOCKSClick(Sender: TObject);
var TOBL : TOB;
begin
  TOBL:=GetTOBDecisAch(TOBdecisionnel,GS.row) ; if TOBL=Nil then Exit ;
	AGLLanceFiche('BTP','BTDISPO_MUL','GA_CODEARTICLE='+TOBL.GetValue('BAD_CODEARTICLE'),'','CONSULTATION') ;

end;

function TOF_DECISIONACHAT.CalculPrixNet(PRIXACHATBRUT: double;REMISES: string): double;
var RemiseResult : double;
begin
  RemiseResult := RemiseResultante(TransformeRemises(REMISES));
  result := arrondi(PRIXACHATBRUT * (1-(RemiseResult/100)),V_PGI.OkDecP );
end;

procedure TOF_DECISIONACHAT.DEPOTCHANGE(Sender: Tobject);
var TOBL : TOB;
begin
  TOBL:=GetTOBDecisAch(TOBdecisionnel,GS.row) ; if TOBL=Nil then Exit ;
  SetDepot(TOBL,DEPOT.Value);
end;

Initialization
  registerclasses ( [ TOF_DECISIONACHAT ] ) ;
end.


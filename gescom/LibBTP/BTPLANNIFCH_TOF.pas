{***********UNITE*************************************************
Auteur  ...... : SANTUCCI Lionel
Créé le ...... : 13/01/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPLANNIFCH ()
Mots clefs ... : TOF;BTPLANNIFCH
*****************************************************************}
Unit BTPLANNIFCH_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$ELSE}
     MainEagl,
{$ENDIF}
     forms,
     menus,
     AglInit,
     saisUtil,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     UTOB,
     UTOF,
     UPlannifchUtil,
     FactComm,
     FactUtil,
     FactCalc,
     FactGrp,
     FactGrpBTP,
     paramSoc,
     EntGC,
     FactOuvrage,
     FactVariante,
     ExtCtrls,
     Graphics,
     FactTob,
     FactAdresse,
		 UtilPlannifchantier,
     hPanel,
     vierge,
     BTStructChampSup
     ,uEntCommun,UtilTOBPiece
     ;

const MAXPHASE = 9;

Type

	TTypeProv = (TTpNormal,TtpFrais);
  TTypeGestion = (TTGDevis,TTGCtrEtude);

  TOF_BTPLANNIFCH = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
  	fLivChantier : boolean;
  	EnHt : boolean;
    ModeTraitement : integer;
    gestion : TTypegestion;
    TOBDEVIS,TOBIntermediaire,TOBCHANTIER,TOBOuvrage,TOBGenere,TOBLiaison,TOBA,TOBFrais,TOBOuvrageFrais,TOBSST :TOB;
    TOBAdresses : TOB;
    TV_DEVIS,TV_CHANTIER : TTreeView;
    BValider,Bferme : TToolBarButton97;
    RootDevis,RootChantier : TTreeNode;
    CurrentOuv : TOB;
    ListPhase : Tlist;
    PopTrait : TpopupMenu;
    NewPhase,DelPhase,ChangeLivr : TMenuItem;
    DEV : Rdevise;
    TheArticle : string;
    TheListImage : TImageList;
    LivrFournChantier : TcheckBox;
    LivrFournDevis : TcheckBox;
    CoefFR : double;
    TOBResult : TOB;
    procedure formresize (sender : TObject);
    function  GetCoefFR (TOBDevis : TOB) : double;
    procedure getComponents;
    procedure SetMethods;
    procedure TV_DEVISDRAGOVER(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure TV_DEVISDRAGDROP(Sender, Source: TObject; X, Y: Integer);
    procedure TV_CHANTIERDRAGOVER(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure TV_CHANTIERDRAGDROP(Sender, Source: TObject; X, Y: Integer);
    procedure BFermeClick(Sender: Tobject);
    procedure BValiderClick(Sender: Tobject);
    procedure ReintegreDEVIS(ItemChantier: TTreeNode);
    function DefiniLibelle(TOBL: TOB): string;
    procedure CreeTob;
    procedure FreeTobs;
    procedure AddChampInterm(TOBD: TOB);
    function findlaTOB(TOBL, TOBDepart: TOB; ListPhase : Tlist): TOB;
    function InserePhaseInterm(TOBdevis,TOBL, TOBINTERMEDIAIRE: TOB;TheTypeLigne : TTypeProv=TtpNormal) : TOB;
    function InsereInterm (TOBDevis,TOBL,TOBDepart : TOB;ListPhase : Tlist;TheTypeLigne : TTypeProv=TtpNormal):TOB;
    procedure CumuleInterm(TOBDevis,TOBL, TOBI: TOB;TheTypeLigne : TTypeProv=TtpNormal);
    function InsereOuvrageInterm(TOBL, TOBDet: TOB;ListPhase : Tlist;TheTypeLigne : TTypeProv=TtpNormal): TOB;
    procedure ChargeComponents;
    procedure IndiceTobIntermediaire;
    procedure chargementTV_DEVIS(TV: TTreeView; Pere: TTreeNode; TobATraiter: TOB);
    function FindPere(SearchedItem: TTreeNode; TOBS: TOB): TTreeNode;
    procedure SupprimeFilleNor(ItemSel: TTreeNode);
    function AjouteFilleChantier(ItemSel, ItemDest: TTreeNode): TTreeNode;
    procedure IntegreChantier(ItemSel, ItemDest: TTreeNode);
    function FindLeNode(SearchedItem: TTreeNode; TOBL: TOB) : TTreeNode;
    procedure DelPhaseClick(Sender: TObject);
    procedure NewPhaseClick(Sender: TObject);
    procedure PopupNonUsable;
    procedure PopupUsable;
    procedure PopTraitPopup(Sender: TObject);
    function ControleIntegration(ItemSel, ItemDest: TTreeNode): boolean;
    function compteNiveau(ItemSel : TTreeNode; NiveauDepart : integer) : integer;
    procedure ReIndicePhase(NodeDepart: TTreeNode; IndiceDep: integer);
    procedure ItemMove(ItemSel, ItemDest: TTreeNode);
    procedure TV_CHANTIERDblClick(Sender: TOBject);
    function verifDevis(ITemDepart: TTreeNode): boolean;
    function ControleChantier: boolean;
    function ControleDevis : boolean;
    function DefiniLigneDocumentChantier(NodeDepart: TTreeNode;TOBLIGNE : TOB;NiveauImbric : integer; var Numligne : integer;DateLivraison : TDateTime = 0) : boolean;
    procedure EnregistreChantier;
    procedure AjouteDebutParagrapheChantier (TOBLigne ,TOBInterm : TOB;NiveauCurr,NumLigne : integer; DateLiv : TDateTime);
    procedure AjouteFinParagrapheChantier (TOBligne,TOBInterm : TOB;NiveauCurr,NumLigne : integer; Libelle : string;DateLivraison : TDateTime=0);
    procedure AjouteLigneChantier (TOBLIGNE,TOBINterm : TOB; NiveauCurr,NumLigne : integer; DateDevis : TdateTime);
//    function DefiniLiaisonDevisChantier (NodeDepart : TTreeNode) : boolean;
    procedure LigneDevisversChantier(TOBInterm,TOBL : TOB;NumLigne : integer;DateLiv : TDateTime);
    procedure CreerPieceChantier;
//    procedure AjouteLiaisondevcha (TOBCH : TOB);
    procedure RemplacePourcent(TOBL,TOBPiece: TOB);
    procedure InitTobs;
    function GenereTobIntermediaire: boolean;
    procedure ChargeReplPourcent;
    procedure MajSubLevels(ItemSel: TTreeNode);
    procedure ModifPhase(ItemSel: TTreeNode);
    procedure CalculeMontantPVLig(var MontantDev, Montant: double;TOBL: TOB);
    procedure TraiteCumuleEltChantier (NodeDepart : TTreeNode);
    procedure G_ChargeLesAdresses(TOBPiece: TOB);
    procedure AffecteLesAdresses(TOBPiece, TOBResult: TOB);
    procedure ChargeListImage;
    procedure ChangeLivrClick(Sender: TObject);
    function FindItemMenu(NomMenu: string): TMenuItem;
    procedure TV_DEVISCHANGE(Sender: TObject; Node: TTreeNode);
    procedure ChangeModeLivr (ItemSel : TTreeNode; Value : integer);
    procedure AffecteImage(NodeSuite: TTreeNode; TOBSuite: TOB);
    procedure GereDestLivr(Destination: TTreeNode);
		procedure TraitementModeGestion;
    procedure AppliqueCoeffRSurLigneFrais(TOBL: TOB; Coeff: double);
    procedure ReajusteLigneDevisPourCtrEtude(TOBL: TOB; CoefFr: double);

  end ;

  function SetPlannification (TOBDEVIS,TOBSST,TobOuvrage,TOBFrais,TOBOUvrageFrais : TOB) : boolean;


Implementation
uses BTPUtil,Facture, Factpiece;

function SetPlannification (TOBDEVIS,TOBSST,TobOuvrage,TOBFrais,TOBOuvrageFrais : TOB) : boolean;
var st,Param : string;
		leChoix : integer;
    OneDoc,cledocResult : R_Cledoc;
begin
  result := true;

  if TOBDevis.getValue('GP_NATUREPIECEG')=VH_GC.AFNatAffaire then
  begin
    OneDoc := TOB2Cledoc (TOBDEVIS);
//uniquement en line
//  lechoix := 1;
//  Param := 'MODE='+ inttostr(LeChoix)+';GESTION=CTRETUDE';
    lechoix := ChoixPlannifdevis (ttcChantier);
    Param := 'MODE='+ inttostr(LeChoix)+';GESTION=DEVIS';

    if LeChoix = -1 then exit;
    TheTob := TOBDEvIS;
    TheTob.Data := TOBOuvrage;
    TOBOuvrage.data := TOBFrais;
    TOBFrais.data := TOBOuvrageFrais;
    TOBOUvrageFrais.data := TOBSST;
    St:=AGLLanceFiche('BTP','BTPLANNIFCH','','',Param) ;
    if TheTOB <> nil then
    begin
      ExecuteSql ('UPDATE AFFAIRE SET AFF_PREPARE="X" '+
                  'WHERE '+
                  'AFF_AFFAIRE=('+
                  'SELECT GP_AFFAIREDEVIS FROM PIECE WHERE '+WherePiece(OneDoc,ttdPiece,false)+')');
//uniquement en line
{*
			if (PgiAsk (TraduireMemoire('Désirez-vous ouvrir la prévision de chantier ?'))=MrYes) then
      begin
      	CledocResult.NaturePiece := TheTOB.Detail[0].getValue('GP_NATUREPIECEG');
        CledocResult.Souche  		 := TheTOB.Detail[0].getValue('GP_SOUCHE');
        CledocResult.DatePiece 	 := TheTOB.Detail[0].getValue('GP_DATEPIECE');
        CledocResult.NumeroPiece := TheTOB.Detail[0].getValue('GP_NUMERO');
        CledocResult.Indice 	 	 := TheTOB.Detail[0].getValue('GP_INDICEG');
				SaisiePiece (CledocResult,TaModif);
      end;
*}
    	TheTOB.free;
    end;
    TheTob := nil;
  end else if TOBDevis.GetVAlue('GP_NATUREPIECEG')='BCE' then
  begin
  	{ ancienne gestion
    lechoix := ChoixPlannifCtrEtude;
    if LeChoix = -1 then exit;
    }
    OneDoc := TOB2Cledoc (TOBDEVIS);
    leChoix := 2;
    Param := 'MODE='+ inttostr(leChoix)+';GESTION=CTRETUDE';
    TheTob := TOBDEvIS;
    TheTob.Data := TOBOuvrage;
    TOBOuvrage.data := TOBFrais;
    TOBFrais.data := TOBOuvrageFrais;
    St:=AGLLanceFiche('BTP','BTPLANNIFCH','','',Param) ;
    if TheTOB <> nil then
    begin
      ExecuteSql ('UPDATE AFFAIRE SET AFF_PREPARE="X" '+
                  'WHERE '+
                  'AFF_AFFAIRE=('+
                  'SELECT GP_AFFAIREDEVIS FROM PIECE WHERE '+WherePiece(OneDoc,ttdPiece,false)+')');
    end;
    TheTob := nil;
  end;
end;

procedure TOF_BTPLANNIFCH.AddChampInterm (TOBD : TOB);
begin
TOBD.addChampsupValeur ('PROVENANCE','DEVIS',false);
TOBD.addChampsupValeur ('NUMLIGNE',0,false);
TOBD.addChampsupValeur ('PHASE','',false);
TOBD.addChampsupValeur ('LIGNEPHASE',0,false);
TOBD.addChampsupValeur ('OUVRAGE','',false);
TOBD.addChampsupValeur ('LIGNEOUV',0,false);
TOBD.addChampsupValeur ('ARTICLE','',false);
TOBD.addChampsupValeur ('TYPE','',false);
TOBD.addChampsupValeur ('LIBELLE','',false);
TOBD.addChampsupValeur ('QTE',0.0,false);
TOBD.addChampsupValeur ('PADRE',0,false);
TOBD.addChampsupValeur ('TRANSFERED','-',false);
TOBD.addChampsupValeur ('NIVEAUPHASE','0',false);
{
TOBD.addChampsupValeur ('DATEDEBUT',V_PGI.DateEntree,false);
TOBD.addChampsupValeur ('DATEDEBUT_',V_PGI.DateEntree,false);
}
TOBD.addChampsupValeur ('DATEDEBUT',iDate2099,false);
TOBD.addChampsupValeur ('DATEDEBUT_',iDate2099,false);
// --
TOBD.addChampsupvaleur ('NUMEROLIGNECH',0,false);
TOBD.addChampsupValeur ('DATEINIT',V_PGI.DateEntree,false);
TOBD.addChampSupValeur ('WOL ORIGINE',-2,false);
TOBD.addChampSupValeur ('IDENTIFIANTWOL',-2,false);
// --
TOBD.addChampsupValeur ('NATURETRAVAIL','',false);
TOBD.addChampsupValeur ('FOURNISSEUR','',false);
TOBD.addChampsupValeur ('DPA',0,false);
TOBD.addChampsupValeur ('DPR',0,false);
TOBD.addChampsupValeur ('PV',0,false);
TOBD.addChampsupValeur ('MONTANTPA',0,false);
TOBD.addChampsupValeur ('MONTANTPV',0,false);
TOBD.addChampsupValeur ('MONTANTPR',0,false);
TOBD.addChampsupValeur ('MONTANTFR',0,false);
TOBD.addChampsupValeur ('MONTANTFC',0,false);
TOBD.addChampsupValeur ('MONTANTFG',0,false);
end;

procedure TOF_BTPLANNIFCH.getComponents;
var Indice : integer;
    TheItem : TMenuItem;
begin
TV_DEVIS := TTreeView(GetControl('TV_DEVIS'));
TV_CHANTIER := TTreeView(GetControl('TV_CHANTIER'));
BValider := TToolBarButton97(GetCOntrol('BValider'));
BFerme := TToolBarButton97(GetCOntrol('Bferme'));
PopTrait := TPopupMenu (GetControl('POPDEF'));

for Indice := 0 to PopTrait.Items.Count -1 do
    begin
    TheItem := PopTrait.Items[Indice];
    if Uppercase (TheItem.Name) = 'DELPHASE' then DelPhase := TheItem
    else if Uppercase (TheItem.Name) = 'ADDPHASE' then NewPhase := TheItem
    else if Uppercase (TheItem.name) = 'LIVRCLIENT' then ChangeLivr := TheItem;
    end;
end;

procedure TOF_BTPLANNIFCH.TV_DEVISCHANGE (Sender: TObject; Node: TTreeNode);
var item: TTreeNode;
begin
item := TV_devis.Selected;  // Element selectionné
end;

procedure TOF_BTPLANNIFCH.TV_DEVISDRAGOVER (Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
var item: TTreeNode;
begin

	accept := true;

	if Source <> TV_CHANTIER then
  	 BEGIN
	   Accept := false;
  	 Exit;
	   END;

	item := TV_CHANTIER.GetNodeAt (X,Y);

	if Item = nil then
     BEGIN
     Accept := false;
     Exit;
     END;

end;

procedure TOF_BTPLANNIFCH.TV_DEVISDRAGDROP (Sender, Source: TObject; X, Y: Integer);
var item: TTreeNode;
		indice : integer;
begin
  for Indice := TV_CHANTIER.SelectionCount -1 downto 0 do
  begin
    item := TV_CHANTIER.Selections [indice];  // Element selectionné
    ReintegreDEVIS (Item);
    if Item <> RootChantier then
    begin
    	Item.Delete // enleve l'element selectionné de la liste de la commande de chantier
    end else
    begin
      if RootChantier.count > 0 then
      begin
        repeat
          item := RootChantier.item[0];
          item.delete;
        until Rootchantier.count = 0;
      end;
    end;
  end;
  TV_DEVIS.FullExpand;
end;

procedure TOF_BTPLANNIFCH.TV_CHANTIERDRAGOVER (Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
var itemSel,ItemDest: TTreeNode;
    TypeSource,TypeDest : string;
    indice : integer;
    okOk : boolean;
begin
  accept := false;
  OkOk := true;
  itemDest := TV_CHANTIER.GetNodeAt (X,Y);
  if ItemDest = nil then Exit;
  for Indice := 0 to TTreeView(Source).SelectionCount -1 do
  begin
    itemSel := TTreeView(Source).Selections [Indice];  // Element selectionné
    if ItemSel = nil then exit;
    if ItemSel.data <> nil then TypeSource := TOB(ItemSel.data).GetValue('TYPE')
                           else TypeSource := 'PHA';
    if (ItemDest <> RootChantier) then
    begin
      if ItemDest.data <> nil then TypeDest := TOB(ItemDest.data).GetValue('TYPE') else TypeDest := 'PHA';
      if (TypeSource <> 'NOR') and (ItemSel.count = 0) and (Source = TV_DEVIS) then
      begin
      	OkOK := false;
        break;
      end;
      if (TypeSource <> 'NOR') and (TypeDest='NOR') then
      begin
      	okOk := false;
      	break;
      end;
    end else
    begin
      if (Source = TV_CHANTIER) and (TypeSource = 'NOR') then
      begin
      	Okok := false;
        break;
      end;
      if (ItemSel.data <> nil) then TypeSource := TOB(ItemSel.data).GetValue('TYPE') else TypeSource := 'PHA';
    end;
  end;
  Accept := okok;
end;

procedure TOF_BTPLANNIFCH.SupprimeFilleNor (ItemSel : TTreeNode);
var Indice : integer;
    ItemFille : TTreeNode;
    TOBS : TOB;
begin
Indice := 0;
if (ItemSel.count = 0) and (TOB(ItemSel.data)<>nil) and (TOB(ItemSel.data).GetValue('TYPE')='NOR') then BEGIN ItemSel.delete; exit; END;
if (ItemSel.count = 0) then exit;
repeat
    ItemFille := ItemSel.Item [indice];
    TOBS := ItemFille.data;
    if TOBS <> nil then
       BEGIN
       if TOBS.GetValue('TYPE') <> 'NOR' Then BEGIN SupprimeFilleNor (ItemFille); inc(Indice); END
                                         else ItemFille.Delete;
       END else Itemfille.Delete;
until Indice > ItemSel.count -1;
end;

function TOF_BTPLANNIFCH.AjouteFilleChantier (ItemSel,ItemDest : TTreeNode):TTreeNode;
var TypeSource: string;
    Transfered : boolean;
    Destination : TTreeNode;
    NiveauPhase : integer;
    TheTOb : TOB;
begin

TypeSource := TOB(ItemSel.data).GetValue('TYPE');
Transfered := (TOB(ItemSel.data).GetValue('TRANSFERED')='X');

if (ItemDest = RootChantier) and (TypeSource = 'NOR') then
   BEGIN
   if ItemDest.count = 0 then
      begin
      TheTob := TOB.Create ('LA LIGNE',TOBINTERMEDIAIRE,-1);
      AddChampInterm (TheTob);
      TheTOB.PutValue('TYPE','PHA');
      TheTOB.PutValue('LIBELLE','Nouvelle Phase');
      if fLivChantier then TheTOB.PutValue('IDENTIFIANTWOL',-1)
      							  else TheTOB.PutValue('IDENTIFIANTWOL',0);
      Destination := TV_CHANTIER.Items.AddChild (Rootchantier,DefiniLibelle(TheTob));
      Destination.Data := TheTOB;
      AffecteImage (Destination,TheTOB);
      end else Destination := Itemdest.Item [Itemdest.count -1];
   END else Destination := ItemDest;

if Destination.Data = nil then
begin
  niveauPhase := 0
end else
begin
  NiveauPhase := TOB(Destination.data).getValue('NIVEAUPHASE');
end;

if (Destination.data <> nil) and (TOB(Destination.data).GetValue('TYPE') = 'NOR') and
   (TOB(ItemSel.data).GetValue('TYPE') = 'NOR') Then
   begin
   // cas du deplacement d'un article du devis sur chantier
   Destination := TV_Chantier.Items.Insert (Destination,ItemSel.Text);
   Destination.data := ITemSel.Data;
   if fLivChantier then TOB(ITemSel.Data).PutValue('IDENTIFIANTWOL',-1)
                   else TOB(ITemSel.Data).PutValue('IDENTIFIANTWOL',0);
   AffecteImage (Destination,TOB(ITemSel.Data));
   end else
   begin
   if not Transfered then
      begin
      Destination := TV_Chantier.Items.addChild(Destination,DefiniLibelle(TOB(ItemSel.data)));
      Destination.data := ITemSel.Data;
      if fLivChantier then TOB(ITemSel.Data).PutValue('IDENTIFIANTWOL',-1)
                      else TOB(ITemSel.Data).PutValue('IDENTIFIANTWOL',0);
      AffecteImage (Destination,TOB(ITemSel.Data));
      if (TypeSource <> 'NOR') then
         begin
         TOB(ItemSel.data).PutValue('TRANSFERED','X');
         TOB(ItemSel.data).PutValue('NIVEAUPHASE',NiveauPhase+1);
         end;
      if Destination.count > 0 then GereDestLivr (Destination);
      end;
   end;
result := destination;
end;

procedure TOF_BTPLANNIFCH.IntegreChantier (ItemSel,ItemDest : TTreeNode);
var Destination : TTreeNode;
    Indice : integer;
begin
Destination := ItemDest;
if ItemSel.data <> nil then
  begin
  Destination := AjouteFilleChantier (ItemSel,ItemDest);
  end;
if ItemSel.count > 0 then
  begin
  for Indice := 0 to ItemSel.count -1 do
      begin
      IntegreChantier (ItemSel.item[Indice],Destination);
      end;
  end;
end;

function TOF_BTPLANNIFCH.compteNiveau (ItemSel : TTreeNode; NiveauDepart : integer) : integer;
var Indice,NbrNiveauMax,NbrCompte : integer;
begin
NbrNiveauMax := NiveauDepart;
if ItemSel.count > 0 then
   begin
   inc(niveaudepart);
   for Indice := 0 to ItemSel.Count -1 do
       begin
       NbrCompte := CompteNiveau (ItemSel.Item[indice],niveaudepart);
       if NbrNiveauMax < NbrCompte then NbrNiveauMax := NbrCompte
       end;
   end;
if NiveauDepart = 0 then NbrNiveauMax := NbrNiveauMax + 1;
result := NbrNiveauMax;
end;

function TOF_BTPLANNIFCH.ControleIntegration (ItemSel,ItemDest : TTreeNode) : boolean;
var NivDepart,NbrCompte : integer;
begin
result := true;
if (ItemDest.data = nil) then NivDepart := 0 else NivDepart :=TOB(ItemDest.data).GetValue('NIVEAUPHASE');
NbrCompte := compteNiveau (ItemSel,0);
if NivDepart + NbrCompte > MAXPHASE then result := false;
if not result then PGIBox ('Nombre d''imbrication maximum de phases dépassé',ecran.Caption );
end;

procedure TOF_BTPLANNIFCH.GereDestLivr (Destination : TTreeNode);
var Indice : integer;
    Tsuiv : TTreeNode;
    ThisTob : TOB;
begin
  for Indice := 0 to Destination.Count -1 do
  begin
    TSuiv := Destination.Item [Indice];
    if TSuiv.data <> nil then
    begin
       ThisTob := TOB(TSuiv.data);
       if fLivChantier then ThisTob.PutValue('IDENTIFIANTWOL',-1)
                       else ThisTob.PutValue('IDENTIFIANTWOL',0);
      AffecteImage (TSuiv,ThisTob);
    end;
    if TSuiv.Count > 0 then GereDestLivr (Tsuiv);
  end;
end;

procedure TOF_BTPLANNIFCH.ItemMove (ItemSel,ItemDest : TTreeNode);
var TNAttach : TNodeAttachMode;
    Destination: TTreeNode;
    TypeDest : string;
begin
if (ItemDest = RootChantier) and (TOB(ItemSel.data).GetValue('TYPE')= 'NOR') then
   BEGIN
   TheTob := TOB.Create ('LA LIGNE',TOBINTERMEDIAIRE,-1);
   AddChampInterm (TheTob);
   TheTOB.PutValue('TYPE','PHA');
   TheTOB.PutValue('LIBELLE','Nouvelle Phase');
   Destination := TV_CHANTIER.Items.AddChild (Rootchantier,DefiniLibelle(TheTob));
   Destination.Data := TheTOB;
   if fLivChantier then TheTOB.PutValue('IDENTIFIANTWOL',-1)
                   else TheTOB.PutValue('IDENTIFIANTWOL',0);
   AffecteImage (Destination,TheTOB);
   END else Destination := ItemDest;

if Destination.data <> nil then
begin
  TypeDest := TOB(Destination.data).GetValue('TYPE')
end else
begin
  TypeDest := 'PHA';
end;
if TypeDest <> 'NOR' then TNAttach := naAddChildfirst else TNAttach := naInsert;
ITemSel.MoveTo (Destination,TNAttach);
end;

procedure TOF_BTPLANNIFCH.TV_CHANTIERDRAGDROP (Sender, Source: TObject; X, Y: Integer);
var itemSel,ItemDest,TTI: TTreeNode;
		indice,II : integer;
    TT : Tlist;
begin
	itemDest := TV_CHANTIER.GetNodeAt (X,Y);
  TT := Tlist.create;
  TTreeView(Source).GetSelections(TT);
  For Indice := 0 to TT.Count -1 do
  begin
    TTI := TT.Items[Indice];
    if TTI.Count > 0 then
    begin
    	For II := 0 to TTI.Count -1 do
      begin
      	TTreeView(source).Deselect ( TTI.Item [II]);
      end;
    end;
  end;
  //
  TT.Clear;
  TTreeView(Source).GetSelections(TT);
  for Indice := 0 to TT.count -1  do
  begin
    itemSel := TT.items [indice]; // Element selectionné

    if Source = TV_DEVIS Then // Provenance de tv devis
    begin
      if ControleIntegration (ItemSel,ItemDest) then
      begin
        IntegreChantier (ItemSel,ItemDest);
        SupprimeFilleNor (ItemSel);
      end else continue;
    end ELSE
    BEGIN
      if ItemSel.data = nil then continue;

      if ControleIntegration (ItemSel,ItemDest) then
      begin
        ItemMove (ItemSel,ItemDest); ItemDest.Selected := false;
      end;
    END;
  end;

  TT.free;
  TV_CHANTIER.FullExpand;
end;

function TOF_BTPLANNIFCH.verifDevis (ITemDepart : TTreeNode) : boolean;
var Indice : integer;
begin
Result := true;
if ItemDepart.data <> nil then
   begin
   if TOB(ItemDepart.data).GetValue ('TYPE') = 'NOR' then BEGIN ItemDepart.Selected := true; Result  := false; Exit; END;
   end;
if ItemDepart.count > 0 then
   begin
   for Indice := 0 to ItemDepart.count -1 do
       begin
       if not verifDevis(Itemdepart.item[Indice]) then BEGIN Result := false; break; END;
       end;
   end;
end;

function TOF_BTPLANNIFCH.ControleDevis  : boolean;
begin
result := verifDevis (RootDevis);
end;

function TOF_BTPLANNIFCH.ControleChantier : boolean;
begin
// Pour l'instant rien dans le controle chantier
result := true;
end;

procedure TOF_BTPLANNIFCH.BValiderClick (Sender : Tobject);
begin
if not ControleDevis then
   begin
   PGIBox (TraduireMemoire('Il reste des articles dans le devis'),ecran.caption);
   TForm(Ecran).ModalResult:=0;
   exit;
   end;
if not ControleChantier then
   begin
   exit;
   end;
TraiteCumuleEltChantier (RootChantier);
EnregistreChantier;
TheTOB := TOBresult;
//close;
end;

procedure TOF_BTPLANNIFCH.AjouteDebutParagrapheChantier (TOBLigne ,TOBInterm : TOB;NiveauCurr,NumLigne : integer; DateLiv : TDateTime);
var TOBL : TOB;
begin
{
TOBL := TOB.Create ('LIGNE',TOBLigne,-1);
InitTobLigne (TOBL);
}
TOBL:=NewTobLigne(TOBLigne,0);
InitialiseLigne (TOBL,NumLigne,0);
PieceVersLigne (TOBDevis,TOBL);
TOBL.PutValue('GL_DATELIVRAISON',Dateliv);
TOBL.putValue('GL_TYPELIGNE','DP'+inttoStr(NiveauCurr));
TOBL.PutValue('GL_NIVEAUIMBRIC',NiveauCurr) ;
TOBL.putValue('GL_NUMLIGNE',numligne);
TOBL.putValue('GL_NUMORDRE',numligne);
TOBL.putValue('GL_LIBELLE',TOBINterm.getValue('LIBELLE'));
TOBL.PutValue('GL_TYPEDIM','NOR') ;
TOBL.PutValue('GL_IDENTIFIANTWOL',TOBInterm.GetValue('IDENTIFIANTWOL'));
TOBL.AddChampSupValeur('GP_REFINTERNE',TOBdevis.GetValue('GP_REFINTERNE'));
TOBL.AddChampSupValeur('GP_REPRESENTANT',TOBdevis.GetValue('GP_REPRESENTANT'));
TOBL.AddChampSupValeur ('PIECEORIGINE',EncodeRefPiece (TOBDEVIS,0,false));
TOBL.AddChampSupValeur ('GP_NUMADRESSEFACT',TOBDevis.GetValue('GP_NUMADRESSEFACT'));
TOBL.AddChampSupValeur ('GP_NUMADRESSELIVR',TOBDevis.GetValue('GP_NUMADRESSELIVR'));
if TOBINterm.GetValue('PROVENANCE') <> 'DEVIS' then SetTypeLigne (TOBL,True);
end;

procedure TOF_BTPLANNIFCH.AjouteFinParagrapheChantier (TOBligne,TOBInterm : TOB;NiveauCurr,NumLigne : integer; Libelle : string;DateLivraison : TDateTime=0);
var TOBL : TOB;
begin
{
TOBL := TOB.Create ('LIGNE',TOBLigne,-1);
AddLesSupLigne (TOBL,false);
}
TOBL:=NewTobLigne(TOBLigne,0);
InitialiseLigne (TOBL,NumLigne,0);
PieceVersLigne (TOBDevis,TOBL);
TOBL.PutValue('GL_DATELIVRAISON',Datelivraison);
TOBL.putValue('GL_TYPELIGNE','TP'+inttoStr(NiveauCurr));
TOBL.PutValue('GL_NIVEAUIMBRIC',NiveauCurr) ;
TOBL.putValue('GL_NUMLIGNE',numligne);
TOBL.putValue('GL_NUMORDRE',numligne);
TOBL.putValue('GL_LIBELLE','TOTAL ' + Libelle);
TOBL.PutValue('GL_TYPEDIM','NOR') ;
TOBL.PutValue('GL_IDENTIFIANTWOL',TOBInterm.GetValue('IDENTIFIANTWOL'));
TOBL.AddChampSUpValeur('GP_REFINTERNE',TOBdevis.GetValue('GP_REFINTERNE'));
TOBL.AddChampSupValeur('GP_REPRESENTANT',TOBdevis.GetValue('GP_REPRESENTANT'));
TOBL.AddChampSupValeur ('PIECEORIGINE',EncodeRefPiece (TOBDEVIS,0,false));
TOBL.AddChampSupValeur ('GP_NUMADRESSEFACT',TOBDevis.GetValue('GP_NUMADRESSEFACT'));
TOBL.AddChampSupValeur ('GP_NUMADRESSELIVR',TOBDevis.GetValue('GP_NUMADRESSELIVR'));
if TOBINterm.GetValue('PROVENANCE') <> 'DEVIS' then SetTypeLigne (TOBL,True);
end;

procedure TOF_BTPLANNIFCH.AjouteLigneChantier (TOBLIGNE,TOBINterm : TOB; NiveauCurr,NumLigne : integer; DateDevis : TdateTime);
var TOBL : TOB;
begin
{
TOBL := TOB.Create ('LIGNE',TOBLigne,-1);
AddLesSupLigne (TOBL,false);
}
TOBL:=NewTobLigne(TOBLigne,0);
InitialiseLigne (TOBL,NumLigne,1);
LigneDevisversChantier (TOBInterm,TOBL,NumLigne,Datedevis);
PieceVersLigne (TOBDevis,TOBL);
if TOBL.GetValue('GL_TYPEARTICLE')='PRE' then
begin
	TOBL.PutValue('BNP_TYPERESSOURCE',RenvoieTypeRes(TOBL.GetValue('GL_ARTICLE')));
end;

TOBL.PutValue('GL_NIVEAUIMBRIC',NiveauCurr) ;
TOBL.PutValue('GL_INDICENOMEN',0) ;
TOBL.PutValue('GL_IDENTIFIANTWOL',TOBInterm.GetValue('IDENTIFIANTWOL'));

TOBL.PutValue('GL_QTESIT',0);
TOBL.PutValue('GL_POURCENTAVANC',0);
TOBL.PutValue('GL_QTEPREVAVANC',0);
TOBL.PutValue('GL_TOTPREVAVANC',0);
TOBL.PutValue('GL_TOTPREVDEVAVAN',0);

TOBL.AddChampSupValeur('GP_REFINTERNE',TOBdevis.GetValue('GP_REFINTERNE'));
TOBL.AddChampSupValeur('GP_REPRESENTANT',TOBdevis.GetValue('GP_REPRESENTANT'));
//
TOBL.AddChampSupValeur ('PIECEORIGINE',EncodeRefPiece (TOBDEVIS,0,false));
TOBL.AddChampSupValeur ('GP_NUMADRESSEFACT',TOBDevis.GetValue('GP_NUMADRESSEFACT'));
TOBL.AddChampSupValeur ('GP_NUMADRESSELIVR',TOBDevis.GetValue('GP_NUMADRESSELIVR'));
TOBL.Data := TOBInterm; // pour avoir la liste des lignes de devis associés
if TOBINterm.GetValue('PROVENANCE') <> 'DEVIS' then SetTypeLigne (TOBL,True);
//
end;

procedure TOF_BTPLANNIFCH.CalculeMontantPVLig(var MontantDev, Montant: double;TOBL: TOB);
var RemLigne,RemPied,Escompte,PQ : double;
begin
  RemLigne := 0 ; RemPied := 0; Escompte := 0;
  PQ:=TOBL.GetValue('GL_PRIXPOURQTE') ; if PQ<=0 then BEGIN PQ:=1.0 ; TOBL.PutValue('GL_PRIXPOURQTE',PQ) ; END ;

// Informations issues de l'article
  if TOBL.GetValue('GL_REMISABLELIGNE')='X' then RemLigne:=TOBL.GetValue('GL_REMISELIGNE')/100.0 ;
  if TOBL.GetValue('GL_REMISABLEPIED')='X' then RemPied:=TOBL.GetValue('GL_REMISEPIED')/100.0 ;
  if TOBL.GetValue('GL_ESCOMPTABLE')='X' then Escompte:=TOBL.GetValue('GL_ESCOMPTE')/100.0 ;

  MontantDev:=TOBL.GetValue('GL_PUHTDEV')*TOBL.GetValue('GL_QTEFACT')*(1.0-RemLigne)/PQ;
  MontantDev:=MontantDEV*(1.0-RemPied)*(1.0-Escompte) ;

  Montant := TOBL.GetValue('GL_PUHT')*TOBL.GetValue('GL_QTEFACT')*(1.0-RemLigne)/PQ;
  Montant:=Montant*(1.0-RemPied)*(1.0-Escompte) ;
end;

procedure TOF_BTPLANNIFCH.LigneDevisversChantier (TOBInterm,TOBL : TOB;NumLigne : integer;DateLiv : TDateTime);
var Montant,MontantDev,QteFact,Prix,PQ,RemLigne,MontantDloc,MontantLoc,MONTantDPA,MONTANTPMAP,MONTANTPMRP,MONTANTDPR: double;
    MtFR,MtFG,MTFC  : double;
    Indice: integer;
    ValDpr : double;
    CodeArticle : string;
begin
// initialisation
CodeArticle := Trim(copy(TOBInterm.GetValue('ARTICLE'),1,18));
Montant := 0; MontantDev := 0 ; QteFact := 0;
MontantDpa := 0; MontantDpr:= 0 ;MOntantPmap := 0; MontantPMRP := 0;
MtFR := 0; MtFg := 0; MTFC := 0;

// --
// /!\ ATTENTION, Le cumul est effectué sur la première ligne du TOBInterM
//
for Indice := 0 to TOBInterM.detail.count -1 do
    begin
    if Indice = 0 then
       BEGIN
       TOBL.Dupliquer (TOBInterm.detail[0],false,true);
       TobInterm.Putvalue('NUMEROLIGNECH',NumLigne);
       TOBL.PutValue('GL_LIBELLE',TOBInterm.getValue('LIBELLE'));
       // Au cas ou -- Vu dans les sous detail d'ouvrages
       if TOBL.getValue('GL_REFARTSAISIE') = '' then
       begin
       		TOBL.PutValue('GL_REFARTSAISIE',TOBL.GetValue('GL_CODEARTICLE'));
       end;
       // --
       TOBL.PutValue('GL_DATELIVRAISON',Dateliv);
       TOBL.PutValue('GL_TYPELIGNE','ART') ;
       TOBL.PutValue('GL_NUMEROSERIE','') ;
       TOBL.PutValue('GL_COLLECTION','') ;
       TOBL.PutValue('GL_NUMLIGNE',NumLigne) ;
       TOBL.PutValue('GL_NUMORDRE',Numligne) ;
       END;
    QteFact := QteFact + TOBInterm.detail[Indice].getValue('GL_QTEFACT');
    CalculeMontantPVLig (MontantDLoc,MontantLoc,TOBInterm.detail[Indice]);
    MontantDev := MontantDev +Arrondi(MontantDloc,V_PGI.OkDecV );
    Montant := Montant +Arrondi(MontantLoc,V_PGI.OkdecV);
    MontantDpa := MontantDpa + TOBInterm.detail[Indice].GetValue('GL_MONTANTPA');
    MTFG := MTFG + TOBInterm.detail[Indice].GetValue('GL_MONTANTFG');
    MTFC := MTFC + TOBInterm.detail[Indice].GetValue('GL_MONTANTFC');
    MTFR := MTFR + TOBInterm.detail[Indice].GetValue('GL_MONTANTFR');
//    MONTantDPA := MontantDpa + TOBInterm.detail[Indice].GetValue('GL_DPA') * TOBInterm.detail[Indice].getValue('GL_QTEFACT');
    MONTantPMAP := MontantPMAP + TOBInterm.detail[Indice].GetValue('GL_PMAP') * TOBInterm.detail[Indice].getValue('GL_QTEFACT');
    MONTantPMRP := MontantPMRP + TOBInterm.detail[Indice].GetValue('GL_PMRP') * TOBInterm.detail[Indice].getValue('GL_QTEFACT');
//    MONTantDPR := MontantDPR + TOBInterm.detail[Indice].GetValue('GL_DPR') * TOBInterm.detail[Indice].getValue('GL_QTEFACT');
    //
    MontantDpr := MontantDpR + TOBInterm.detail[Indice].GetValue('GL_MONTANTPR');
    //
    END;
PQ:=TOBL.GetValue('GL_PRIXPOURQTE') ; if PQ<=0 then BEGIN PQ:=1.0 ; TOBL.PutValue('GL_PRIXPOURQTE',PQ) ; END ;
if TOBL.GetValue('GL_REMISABLELIGNE')='X' then RemLigne:=TOBL.GetValue('GL_REMISELIGNE')/100.0
                                          else BEGIN RemLigne := 0; TOBL.PutValue('GL_REMISELIGNE',0) ; END;
TOBL.PutValue('GL_QTEFACT',QteFact) ;
{Autres prix}
if QteFact <> 0 then TOBL.PutValue('GL_DPA', MontantDPA / QteFact);
  TOBL.PutValue('GL_MONTANTPA',MONTantDPA);
  TOBL.PutValue('GL_MONTANTPAFG',MONTantDPA);
  TOBL.PutValue('GL_MONTANTPAFR',MONTantDPA);
  TOBL.PutValue('GL_MONTANTPAFC',MONTantDPA);
  TOBL.PutValue('GL_COEFFG',0);
  TOBL.PutValue('GL_COEFFC',0);
  TOBL.PutValue('GL_COEFFR',0);
  TOBL.PutValue('GL_MONTANTFG',MtFG);
  TOBL.PutValue('GL_MONTANTFC',MTFC);
  TOBL.PutValue('GL_MONTANTFR',MtFR);
  TOBL.PutValue('GL_TOTALHTDEV',MontantDev);

  if MontantDpa <> 0 then TOBL.PutValue('GL_COEFFR',MtFr/montantDpa);

  if MontantDpa <> 0 then
  begin
    TOBL.PutValue('GL_COEFFG',MtFG/MontantDpa);
  end;
  TOBL.PutValue('GL_MONTANTPR',MontantDpa + MtFR+ MTFC+MtFG);

  if QteFact <> 0 then TOBL.PutValue('GL_DPR', Arrondi((MontantDpa + MtFR+ MtFC+MtFG) / QteFact,V_PGI.OkdecP));

if QteFact <> 0 then
begin
  TOBL.PutValue('GL_PMAP', MontantPMAP / QteFact);
  TOBL.PutValue('GL_PMRP', MontantPMRP / QteFact);
  if TOBInterm.getValue('PROVENANCE')='DEVIS' then
  begin
    {Prix}
    Prix:=(Montant*PQ)/(QteFact*(1.0-RemLigne)) ;
    TOBL.PutValue('GL_PUHT',Arrondi(Prix,V_PGI.OkdecP)) ;
    Prix:=(MontantDev*PQ)/(QteFact*(1.0-RemLigne)) ;
    TOBL.PutValue('GL_PUHTDEV',Arrondi(Prix,V_PGI.OkDecP)) ;
    TOBL.pUTValue('GL_BLOQUETARIF','-');
//    if (TOBL.GEtValue('GL_BLOQUETARIF')='X') then
//    begin
//    	TOBL.PutValue('GL_TOTALHTDEV',Arrondi(TOBL.GetValue('GL_PUHTDEV')*QteFact,V_PGI.okdecV));
//    end;
    if TOBL.GetValue('GL_MONTANTPR') <> 0 then
    begin
      TOBL.PutValue('GL_COEFMARG',Arrondi(TOBL.GetValue('GL_TOTALHTDEV')/TOBL.GetValue('GL_MONTANTPR'),4));
    end;
  end else
  begin
    TOBL.PutValue('GL_PUHT',0) ;
    TOBL.PutValue('GL_PUHTDEV',0) ;
    TOBL.PutValue('GL_TOTALHTDEV',0);
  end;
end;

TOBL.PutValue('GL_INDICENOMEN',0);
end;

function TOF_BTPLANNIFCH.DefiniLigneDocumentChantier (NodeDepart: TTreeNode;TOBLIGNE : TOB;NiveauImbric : integer; var Numligne : integer;DateLivraison : TDateTime ) : boolean;
var Indice : integer;
    NiveauCurr : integer;
    DateLiv : TDateTime;
    NodeCurr : TTreeNode;
begin
result := true;
if DateLivraison <> 0 then DateLiv := DateLivraison else Dateliv := iDate2099;
for Indice := 0 to NodeDepart.count -1 do
    begin
    NodeCurr := NodeDepart.Item[Indice];
    if NodeCurr.data = nil then continue;
    if TOB(NodeCurr.data).GetValue('TYPE') <> 'NOR' then (* c'est forcement une phase *)
       begin
       // Creation de la ligne de paragraphe
       DateLiv := StrToDate(TOB(NodeCurr.data).GetValue('DATEDEBUT'));
       NiveauCurr := NiveauImbric + 1;
       AjouteDebutParagrapheChantier (TOBLIGNE,TOB(NodeCurr.data),NiveauCurr,NumLigne,DateLiv);
       inc(Numligne);
       if NodeCurr.count > 0 then
          begin
          DefiniLigneDocumentChantier (NodeCurr,TOBLIGNE,NiveauCurr,NumLigne,DateLiv);
          end;
       AjouteFinParagrapheChantier (TOBLIGNE,TOB(NodeCurr.data),NiveauCurr,NumLigne,TOB(NodeCurr.data).GetValue('LIBELLE'),DateLiv);
       inc(NumLigne);
       end else
       begin
       AjouteLigneChantier (TOBLIGNE,TOB(NodeCurr.data),NiveauImbric,NumLigne,DateLiv);
       inc(NumLigne);
       end;
    end;
end;

{
procedure TOF_BTPLANNIFCH.AjouteLiaisondevcha (TOBCH : TOB);
var Indice : integer;
    TOBL,TOBLI,TOBLD : TOB;
begin
  TOBL := TOB.Create ('LIGNE',nil,-1);
  TRY
  for Indice := 0 to TOBCH.detail.count -1 do
    begin
    TOBLD := TOBCH.detail[Indice];
    TOBLI := TOB.Create('LIENDEVCHA',TOBLiaison,-1);
    TOBLI.PutValue('BDA_REFD',EncodeLienDevCHA(TOBLD));
    TOBLI.PutValue('BDA_NUMLD',TOBLD.getValue('GL_NUMLIGNE'));
    if TOBLD.FieldExists ('N1') then
    begin
      TOBLI.PutValue('BDA_N1D',TOBLD.GetValue('N1'));
      TOBLI.PutValue('BDA_N2D',TOBLD.GetValue('N2'));
      TOBLI.PutValue('BDA_N3D',TOBLD.GetValue('N3'));
      TOBLI.PutValue('BDA_N4D',TOBLD.GetValue('N4'));
      TOBLI.PutValue('BDA_N5D',TOBLD.GetValue('N5'));
    end;
    TOBLI.PutValue('BDA_RANGD',Indice);
    TOBL.InitValeurs;
    TOBL.PutValue('GL_DATEPIECE',V_PGI.DateEntree);
    TOBL.PutValue('GL_NATUREPIECEG',GetParamSoc('SO_BTNATCHANTIER'));
    TOBL.PutValue('GL_SOUCHE',TOBLD.GetValue('GL_SOUCHE'));
    TOBL.PutValue('GL_NUMERO',TOBLD.GetValue('GL_NUMERO'));
    TOBL.Putvalue('GL_INDICEG',0);
    TOBLI.PutValue('BDA_REFC',EncodeLienDevCHA(TOBL));
    TOBLI.PutValue('BDA_NUMLC',TOBCH.GetValue('NUMEROLIGNECH'));
    end;
  FINALLY
    TOBL.free
  END;
end;

function TOF_BTPLANNIFCH.DefiniLiaisonDevisChantier (NodeDepart : TTreeNode) : boolean;
var indice : integer;
    NodeCurr : TTreeNode;
begin
result := true;
for Indice := 0 to NodeDepart.count -1 do
    begin
    NodeCurr := NodeDepart.Item[Indice];
    if NodeCurr.data = nil then continue;
    if TOB(NodeCurr.data).GetValue('TYPE') <> 'NOR' then (* c'est forcement une phase *)
       begin
       // Creation de la ligne de paragraphe
       if NodeCurr.count > 0 then
          begin
          DefiniLiaisonDevisChantier (NodeCurr);
          end;
       end else
       begin
       AjouteLiaisondevcha (TOB(NodeCurr.data));
       end;
    end;
end;
}

procedure TOF_BTPLANNIFCH.CreerPieceChantier;
var TOBPiece,TOBChantier : TOB;
    cledoc : R_CLEDOC;
    stChaine : string;
    Indice : integer;
    newPiece : boolean;
begin
  TOBChantier := TOB.Create ('PIECE',nil,-1);
  TOBPiece := TOB.create ('PIECE',nil,-1);
  TOBPiece.dupliquer (TOBDEVIS,false,true);
  TRY
    newPiece := ExistsChantier (TOBDevis,TOBChantier);
    if NewPiece then
    begin
      if not CreerPiecesFromLignes(TobGenere,'DEVTOCHAN',iDate1900,true,True,TOBResult) then exit;
      //FV1 - 12/06/2015 : On ne teste jamais si Tobresult est vide !!! dans ce cas erreur : index hors limites
      If Tobresult.detail.count = 0 then
        
      else
        cledoc := TOB2CleDoc (TOBresult.detail[0]);
    end else
    begin
      if not AjouteLignesToPiece (TOBGenere,TOBChantier,'DEVTOCHAN',True,true,TOBResult) then exit;
      //FV1 - 12/06/2015 : On ne teste jamais si Tobresult est vide !!! dans ce cas erreur : index hors limites
      cledoc := TOB2CleDoc (TOBresult.detail[0]);
    end;
    stchaine := '';
    stchaine := CleDocToString (cledoc);
//    TOBPiece.PutValue ('GP_DEVENIRPIECE',stchaine);
    if V_PGI.IoError=oeOk then if Not TOBPiece.UpdateDB (true) then V_PGI.IoError:=oeUnknown ;
  FINALLY
    TOBPiece.free;
    TOBChantier.free;
  END;
end;

procedure TOF_BTPLANNIFCH.EnregistreChantier;
var    NumLigne : integer;
begin
Numligne := 1;
if DefiniLigneDocumentChantier (RootChantier,TOBGenere,0,Numligne) then
   begin
//   if DefiniLiaisonDevisChantier (RootChantier) then
//      begin
      Transactions(CreerPieceChantier,1) ;
//      end;
   end;
end;

procedure TOF_BTPLANNIFCH.BFermeClick (Sender : Tobject);
begin
//close;
end;

procedure TOF_BTPLANNIFCH.SetMethods;
begin
TFVierge(ecran).OnResize := FormResize;
//
//TV_DEVIS.OnChange := TV_DEVISCHANGE;
TV_DEVIS.OnDragOver := TV_DEVISDRAGOVER;
TV_DEVIS.OnDragDrop := TV_DEVISDRAGDROP;
//
TV_CHANTIER.OnDragOver := TV_CHANTIERDRAGOVER;
TV_CHANTIER.OnDragDrop := TV_CHANTIERDRAGDROP;
TV_CHANTIER.RightClickSelect := true;
TV_CHANTIER.OnDblClick := TV_CHANTIERDblClick;

PopTrait.OnPopup  := PopTraitPopup;
if (gestion = TTGCtrEtude) then
begin
	TmenuItem(GetControl('AddPhase')).visible := false;
	TmenuItem(GetControl('DelPhase')).visible := false;
	PopTrait.OnPopup  := PopTraitPopup;
//	TV_CHANTIER.PopupMenu := nil;
	TV_CHANTIER.OnDragOver := nil;
	TV_CHANTIER.OnDragDrop := nil;
end;
//
BValider.OnClick := BValiderClick;
BFerme.Onclick := BFermeClick;

if DelPhase <> nil then DelPhase.OnClick := DelPhaseClick;
if NewPhase <> nil then NewPhase.OnClick := NewPhaseClick;
if ChangeLivr <> Nil then ChangeLivr.Onclick := ChangeLivrClick;
end;

procedure TOF_BTPLANNIFCH.ModifPhase (ItemSel : TTreeNode);
var TOBPassage : TOB;
    st : string;
begin
  TOBPassage := TOB.Create ('The Passage',nil,-1);
  AddChampInterm (TOBPassage);
  TOBPassage.Dupliquer (TOB(ItemSel.data),false,true);
  if AppelModifPhase (TOBpassage) then
  BEGIN
    TOB(ItemSel.data).PutValue('LIBELLE',TOBPassage.GetValue('LIBELLE'));
    TOB(ItemSel.data).PutValue('DATEDEBUT',TOBPassage.GetValue('DATEDEBUT'));
    TOB(ItemSel.data).PutValue('DATEDEBUT_',TOBPassage.GetValue('DATEDEBUT_'));
    ItemSel.Text :=  DefiniLibelle (TOB(ItemSel.data));
    MajSubLevels (ItemSel);
  END;
  TOBPassage.free;
end;

procedure TOF_BTPLANNIFCH.TV_CHANTIERDblClick (Sender : TOBject);
var ItemSel : TTreeNode;
begin
  ItemSel := TV_CHANTIER.Selected;
  if ItemSel.data = nil then exit;
  if TOB(ItemSel.data).getValue('TYPE')='NOR' then exit;
  TRY
    ModifPhase (ItemSel);
  FINALLY
    ItemSel.Expand (true);
    TV_CHANTIER.refresh;
//    TV_CHANTIER.FullExpand;
  end;
end;

procedure TOF_BTPLANNIFCH.MajSubLevels (ItemSel : TTreeNode);
var Indice : integer;
    ItemSuiv : TTreeNode;
    TOBItem,TOBSuiv : TOB;
begin
  TOBItem := TOB(ItemSel.data);
  for Indice := 0 to ItemSel.Count -1 do
  begin
    ItemSuiv := ItemSel.Item [Indice];
    if ItemSuiv.count > 0 then
    begin
      // c'est forcement une phase
      if ItemSuiv.data <> nil then
      begin
        TobSuiv := TOB(ItemSuiv.data);
        if (TOBSuiv.GetValue ('DATEDEBUT') < TOBItem.GetValue ('DATEDEBUT')) or
         		(TOBSuiv.GetValue ('DATEDEBUT') = Idate2099)  then
        begin
        	TOBSuiv.PutValue ('DATEDEBUT',TOBItem.getValue('DATEDEBUT'));
        end;
        TOBSuiv.PutValue ('DATEINIT',TOBItem.getValue('DATEDEBUT'));
        TOBSuiv.PutValue ('DATEDEBUT_',TOBItem.getValue('DATEDEBUT_'));
        ItemSuiv.Text := DefiniLibelle (TOBSuiv);
        MajSubLevels (ItemSuiv);
      end;
    end;
  end;
end;

Function TOF_BTPLANNIFCH.FindItemMenu (NomMenu : string) : TMenuItem;
var Indice : integer;
begin
  result := nil;
  for Indice := 0 to PopTrait.Items.Count -1 do
  begin
    if UpperCase (PopTrait.Items[Indice].Name) = Uppercase(NomMenu) then
    begin
       result := PopTrait.Items[indice];
    end;
  end;
end;

procedure TOF_BTPLANNIFCH.PopTraitPopup (Sender : TObject);
var itemSel : TTreeNode;
    theTob : TOB;
    TheItemMenu : TmenuItem;
begin
  ItemSel := TV_CHANTIER.Selected;

  if ItemSel.data <> nil then
  begin
    TheTOB := TOB(ItemSel.data);
    if (TheTOB.GetValue('IDENTIFIANTWOL') <> -1) then
    begin
      TheItemMenu := FindItemMenu ('LIVRCLIENT');
      TheItemMenu.Caption := TraduireMemoire('Livraison fournisseur au dépot');
    end else
    begin
      TheItemMenu := FindItemMenu ('LIVRCLIENT');
      TheItemMenu.Caption := TraduireMemoire('Livraison fournisseur sur chantier');
    end;
  end else
  begin
    if LivrFournChantier.Checked then
    begin
      TheItemMenu := FindItemMenu ('LIVRCLIENT');
      TheItemMenu.Caption := TraduireMemoire('Livraison fournisseur au dépot');
    end else
    begin
      TheItemMenu := FindItemMenu ('LIVRCLIENT');
      TheItemMenu.Caption := TraduireMemoire('Livraison fournisseur sur chantier');
    end;
  end;

  if (ItemSel.data <> nil) and (TheTob.GetValue('TYPE') = 'NOR') then
  begin
     PopupNonUsable;
     Exit;
  end;
  PopupUsable;
end;

procedure TOF_BTPLANNIFCH.PopupNonUsable;
var Indice : integer;
begin
  for Indice := 0 to PopTrait.Items.Count -1 do
  begin
    if UpperCase (PopTrait.Items[Indice].Name) <> 'LIVRCLIENT' then
    begin
       PopTrait.Items[indice].Enabled := false;
    end;
  end;
end;


procedure TOF_BTPLANNIFCH.PopupUsable;
var Indice : integer;
begin
for Indice := 0 to PopTrait.Items.Count -1 do
    begin
    PopTrait.Items[indice].Enabled := true;
    end;
end;

procedure TOF_BTPLANNIFCH.DelPhaseClick (Sender : TObject);
var ItemSel,ItemPrev : TTreeNode;
begin
ItemSel := TV_Chantier.Selected;
if ItemSel = RootChantier then exit;
ItemPrev :=  ItemSel.GetPrev;

if (ItemPrev.data <> nil) and (TOB(ItemPrev.data).GetValue('TYPE')='NOR') then
   begin
   repeat
   ItemPrev := ItemPrev.GetPrev;
   until (ItemPrev.data=nil) or ((ItemPrev.data<>nil) and (TOB(ItemPrev.data).GetValue('TYPE')<>'NOR'));
   end;
if (ItemPrev <> nil) and (ItemPrev <> RootChantier) then
   begin
   if ItemSel.count > 0 then
      begin
      repeat
          ItemMove (ItemSel.Item[0],ItemPrev);
      until ItemSel.count = 0;
      end;
   end;
ReintegreDEVIS (ItemSel);
ItemSel.Delete; // enleve l'element selectionné de la liste de la commande de chantier
ReIndicePhase (Rootchantier,0);
TV_CHANTIER.FullExpand;
end;

procedure TOF_BTPLANNIFCH.ChangeModeLivr (ItemSel : TTreeNode; Value : integer);
var Indice: integer;
    NodeSuite : TTreeNode;
begin
if ItemSel.data <> nil then
begin
  TOB(ItemSel.data).putvalue('IDENTIFIANTWOL',Value);
  AffecteImage (ItemSel,TOB(ItemSel.data));
end;
for Indice := 0 To ItemSel.count -1 do
    begin
    NodeSuite := ItemSel.Item [Indice];
    if (NodeSuite.data <> nil) then
       begin
       ChangeModeLivr(NodeSuite,Value);
       end;
    end;
end;

procedure TOF_BTPLANNIFCH.ChangeLivrClick (Sender : TObject);
var ItemSel : TTreeNode;
    NextValue : integer;
begin
ItemSel := TV_Chantier.Selected;
if ItemSel.data = nil then
begin
  if LivrFournChantier.checked then NextValue := 0 else NextValue := -1;
  LivrFournChantier.Checked := not LivrFournChantier.Checked;
end else
begin
  if TOB(ItemSel.data).GetValue('IDENTIFIANTWOL') = 0 then NextValue := -1  // livraison fournisseur chez client
                                                      else NextValue := 0;  // livraison fournisseur a la maison
end;
ChangeModeLivr (ItemSel,NextValue);
TV_CHANTIER.Invalidate;
TV_CHANTIER.Refresh;
end;

procedure TOF_BTPLANNIFCH.NewPhaseClick (Sender : TObject);
var ItemSel,Destination : TTreeNode;
    TheTOB : TOB;
    NivDepart,NbrCompte : integer;
begin
ItemSel := TV_Chantier.Selected;
if ItemSel.data <> nil then NivDepart := TOB(ItemSel.data).GetValue('NIVEAUPHASE') else NivDepart := 0;
NbrCompte := compteNiveau (ItemSel,0);
if NivDepart + NbrCompte + 1 > MAXPHASE then
   BEGIN
   PGIBox ('Nombre d''imbrication maximum de phases dépassé',ecran.Caption );
   Exit;
   END;
TheTob := TOB.Create ('LA LIGNE',TOBINTERMEDIAIRE,-1);
AddChampInterm (TheTob);
TheTOB.PutValue('TYPE','PHA');
TheTob.PutValue('LIBELLE','Nouvelle Phase');
TheTob.AddChampSupValeur('GP_REFINTERNE',TOBDevis.getValue('GP_REFINTERNE'));
Destination := TV_CHANTIER.Items.AddChild (ItemSel,DefiniLibelle(TheTOB));
Destination.data := TheTOB;
if fLivChantier  then TheTOB.PutValue('IDENTIFIANTWOL',-1)
                 else TheTOB.PutValue('IDENTIFIANTWOL',0);
AffecteImage (Destination,TheTOB);
ReIndicePhase (Rootchantier,0);
TV_CHANTIER.FullExpand;
end;

procedure TOF_BTPLANNIFCH.ReIndicePhase (NodeDepart: TTreeNode; IndiceDep : integer);
var Indice,IndiceSuite : integer;
    NodeSuite : TTreeNode;
begin
if (NodeDepart.data <> nil) and (TOB(NodeDepart.data).GetValue('TYPE')<>'NOR') then TOB(NodeDepart.data).putValue('NIVEAUPHASE',IndiceDep);
for Indice := 0 To NodeDepart.count -1 do
    begin
    NodeSuite := NodeDepart.Item [Indice];
    if (NodeSuite.data <> nil) and (TOB(NodeSuite.data).GetValue('TYPE')<>'NOR') then
       begin
       IndiceSuite := IndiceDep +1;
       ReIndicePhase(NodeSuite,IndiceSuite);
       end;
    end;
end;

function TOF_BTPLANNIFCH.FindPere (SearchedItem : TTreeNode;TOBS : TOB) : TTreeNode;
var TOBSEARCHED : TOB;
    Indice : integer;
    TheItem : TTreeNode;
begin
result := nil;
for Indice := 0 to SearchedItem.Count -1 do
    begin
    TheItem := SearchedItem.Item [indice];
    TOBSEARCHED := TheItem.data;
    if TOBSearched <> nil then
       begin
       if TOBSEARCHED.GetValue('NUMLIGNE') = TOBS.GetValue('PADRE') then
          begin
          result := TheItem;
          break;
          end;
       end;
    if TheItem.count > 0 then result := FindPere (TheItem,TOBS);
    if result <> nil then break;
    end;
end;

function TOF_BTPLANNIFCH.FindLeNode (SearchedItem: TTreeNode;TOBL : TOB) : TTreeNode;
var TOBSEARCHED : TOB;
    Indice : integer;
    TheItem : TTreeNode;
begin
result := nil;
for Indice := 0 to SearchedItem.Count -1 do
    begin
    TheItem := SearchedItem.Item [indice];
    TOBSEARCHED := TheItem.data;
    if TOBSearched <> nil then
       begin
       if TOBSEARCHED.GetValue('NUMLIGNE') = TOBL.GetValue('NUMLIGNE') then
          begin
          result := TheItem;
          break;
          end;
       end;
    if TheItem.count > 0 then result := FindLeNode (TheItem,TOBL);
    if result <> nil then break;
    end;
end;

procedure TOF_BTPLANNIFCH.ReintegreDEVIS (ItemChantier: TTreeNode);
var Inserted,ItemFille: TTreeNode;
    TOBS: TOB;
    Indice : integer;
    found : TTreeNode;
begin
TOBS := ItemChantier.data;
if TOBS <> nil then TOBS.PutValue('IDENTIFIANTWOL',TOBS.GetValue('WOL ORIGINE'));
if (TOBS <> nil) and (TOBS.GetValue('TYPE') = 'NOR')  then
   begin
   found := FindPere (RootDevis,TOBS);
   if found <> nil then
      begin
      Inserted := TV_DEVIS.Items.AddChildFirst (found,DefiniLibelle(TOBS));
      Inserted.Data := TOBS;
      end else
      begin
      Inserted := TV_DEVIS.Items.AddChild (RootDevis,DefiniLibelle(TOBS));
      Inserted.Data := TOBS;
      end;
   AffecteImage (Inserted,TOBS);
   end else
   BEGIN
   if TOBS <> nil then
      begin
      found := FindLeNode (RootDevis,TOBS);
      if found <> nil then
        begin
        TOBS.PutValue('TRANSFERED','-');
        AffecteImage (Found,TOBS);
        end;
      end;
   END;
// S'il a un détail associé ..on parcours le detail pour le rattacher
if ItemChantier.Count > 0 then
   begin
   for Indice := 0 to ItemChantier.count-1 do
       begin
       ItemFille := ItemCHantier.Item[indice];
       ReintegreDEVIS (ItemFille);
       end;
   end;
end;

function TOF_BTPLANNIFCH.DefiniLibelle(TOBL : TOB):string;
begin
  result := '';
  if TOBL.GetValue('TYPE') <> 'NOR' Then
  begin
    result := '(Phase) ';
    result := result + copy(TOBL.GetValue('LIBELLE'),1,25);
    result := result + ' ' + DateToStr (TOBL.GetValue('DATEDEBUT'));
  end else
  begin
    if TOBL.GetString('NATURETRAVAIL') <> '' then result := result + '[Sous traitance] ';
    result := result + copy(TOBL.GetValue('LIBELLE'),1,25);
    result := result + ' ('+ StrS0(TOBL.GetValue('QTE')) +')';
  end;
end;

function TOF_BTPLANNIFCH.InsereInterm (TOBDevis,TOBL,TOBDepart : TOB;ListPhase : Tlist;TheTypeLigne : TTypeProv):TOB;
var Reference,TOBS,TOBLIG : TOB;
    TypeArt : string;
    Qte : double;
    Fournisseur : string;
begin
  if TOBL.GetString('GLC_NATURETRAVAIl')='002' then
  begin
    Fournisseur := TOBL.GetString('GL_FOURNISSEUR');
  end else
  begin
    Fournisseur := '';
  end;
  TypeArt := TOBL.GetValue('GL_TYPEARTICLE');
  //Qte := TOBL.getValue('GL_QTEFACT')/TOBL.getValue('GL_QTEFACT')
  if (ListPhase = nil) or (ListPhase.Count = 0) then Reference := TOBDepart
                                                else Reference := TOB(ListPhase.Items[ListPhase.count-1]);
  TOBS := TOB.Create ('LA LIGNE',Reference,-1);
  // TODO AJOUTER LA MISE A JOUR DES ZONES SUP
  AddChampInterm (TOBS);
  if (ListPhase <> nil) and (ListPhase.count > 0) then
     BEGIN
     TOBS.PutValue ('PHASE',TOB(ListPhase.Items[ListPhase.count-1]).GetValue('PHASE'));
     TOBS.PutValue ('LIGNEPHASE',TOB(ListPhase.Items[ListPhase.count-1]).GetValue('LIGNEPHASE'));
     END;
  if TheTypeLigne = TTpNormal then TOBS.PutValue ('PROVENANCE','DEVIS')
                              else TOBS.PutValue ('PROVENANCE','FRAIS');
  if TOBL.GetValue('GLC_NATURETRAVAIL') = '002' then
  begin
    Fournisseur := TOBL.GetString('GL_FOURNISSEUR');
  end else
  begin
    Fournisseur := '';
  end;
  TOBS.PutValue ('OUVRAGE',CurrentOuv.GetValue('GL_ARTICLE'));
  TOBS.PutValue ('LIGNEOUV',CurrentOUV.GetValue('GL_NUMLIGNE'));
  TOBS.PutValue ('ARTICLE',TOBL.GetValue('GL_ARTICLE'));
  TOBS.PutValue ('TYPE','NOR');
  TOBS.putValue ('LIBELLE',TOBL.GetValue('GL_LIBELLE'));
  TOBS.PutValue ('QTE',TOBL.GetValue('GL_QTEFACT')+TOBS.GetValue('QTE'));
  //TOBS.PutValue ('DATEDEBUT',TOBL.GetValue('GL_DATEPIECE'));
  TOBS.PutValue ('DATEDEBUT',iDate2099);
  TOBS.PutValue ('NATURETRAVAIL',TOBL.GetValue('GLC_NATURETRAVAIL'));
  TOBS.PutValue ('FOURNISSEUR',Fournisseur);
  TOBS.PutValue ('DPA',TOBL.GetValue('GL_DPA'));
  TOBS.PutValue ('PV',TOBL.GetValue('GL_PUHTDEV'));
  TOBS.PutValue ('DPR',TOBL.GetValue('GL_DPR'));
  TOBS.PutValue ('MONTANTPV',TOBL.GetValue('GL_QTEFACT')*TOBL.GetValue('GL_PUHTDEV'));
  //TOBS.PutValue ('MONTANTPA',TOBL.GetValue('GL_QTEFACT')*TOBL.GetValue('GL_DPA'));
  //TOBS.PutValue ('MONTANTPR',TOBL.GetValue('GL_QTEFACT')*TOBL.GetValue('GL_DPR'));
  TOBS.PutValue ('MONTANTPA',TOBL.GetValue('GL_MONTANTPA'));
  TOBS.PutValue ('MONTANTPR',TOBL.GetValue('GL_MONTANTPR'));

  TOBS.PutValue ('MONTANTFR',TOBL.GetValue('GL_MONTANTFR'));
  TOBS.PutValue ('MONTANTFG',TOBL.GetValue('GL_MONTANTFG'));
  TOBS.PutValue ('MONTANTFC',TOBL.GetValue('GL_MONTANTFC'));

  if fLivChantier then TOBS.PutValue('IDENTIFIANTWOL',-1)
                  else TOBS.PutValue('IDENTIFIANTWOL',0);

  (*
  if (gestion <> TTGCtrEtude) then
  begin
    if fLivChantier then TOBS.PutValue('IDENTIFIANTWOL',-1) else TOBS.PutValue('IDENTIFIANTWOL',0);
  end else
  begin
    TOBS.PutValue('IDENTIFIANTWOL',TOBL.GetValue('GL_IDENTIFIANTWOL'))
  end;
  *)
  //if fLivChantier then TOBS.PutValue('IDENTIFIANTWOL',-1) else TOBS.PutValue('IDENTIFIANTWOL',0);
  //TOBS.PutValue ('IDENTIFIANTWOL',TOBL.GetValue('GL_IDENTIFIANTWOL'));
  TOBS.putValue ('WOL ORIGINE',TOBL.GetValue('GL_IDENTIFIANTWOL'));
  //
  TOBLIg:=NewTobLigne(TOBS,0);
  {
  TOBLIG := TOB.Create ('LIGNE',TOBS,-1);
  AddLesSupLigne (TOBLIG,false);
  }
  //AddLesReferences (TOBLig,TOBL);
  TOBLIG.dupliquer (TOBL,false,true);
  TOBS.AddChampSupValeur ('GP_REFINTERNE',TOBDevis.GetValue('GP_REFINTERNE'));

  if (TypeArt = 'OUV') or (TypeArt = 'ARP') then TOBLIG.PutValue('GL_INDICENOMEN',0);
  result := TOBS;
end;

function TOF_BTPLANNIFCH.InserePhaseInterm (TOBDevis,TOBL,TOBINTERMEDIAIRE : TOB;TheTypeLigne : TTypeProv ) : TOB;
var Reference,TOBS : TOB;
begin
if (ListPhase = nil) or (ListPhase.Count = 0) then Reference := TOBINTERMEDIAIRE
                                              else Reference := TOB(ListPhase.Items[ListPhase.count-1]);
TOBS := TOB.Create ('LA LIGNE',Reference,-1);
// TODO AJOUTER LA MISE A JOUR DES ZONES SUP
AddChampInterm (TOBS);
if TheTypeLigne = TTpNormal then TOBS.PutValue ('PROVENANCE','DEVIS')
														else TOBS.PutValue ('PROVENANCE','FRAIS');
TOBS.PutValue ('PHASE',TOBL.GetValue('GL_LIBELLE'));
TOBS.PutValue ('LIGNEPHASE',TOBL.GetValue('GL_NUMLIGNE'));
TOBS.PutValue ('TYPE','PHA');
TOBS.putValue ('LIBELLE',TOBL.GetValue('GL_LIBELLE'));
//if fLivChantier then TOBS.PutValue('IDENTIFIANTWOL',-1) else TOBS.PutValue('IDENTIFIANTWOL',0);
TOBS.PutValue ('WOL ORIGINE',TOBL.GetValue('GL_IDENTIFIANTWOL'));
if fLivChantier then TOBS.PutValue('IDENTIFIANTWOL',-1)
                else TOBS.PutValue('IDENTIFIANTWOL',0);
TOBS.AddChampSupValeur ('GP_REFINTERNE',TOBDevis.GetValue('GP_REFINTERNE'));
result := TOBS;
end;

function TOF_BTPLANNIFCH.InsereOuvrageInterm (TOBL,TOBDet : TOB; ListPhase : Tlist;TheTypeLigne : TTypeProv) : TOB;
var Reference,TOBS : TOB;
begin
if (ListPhase = nil) or (ListPhase.Count = 0) then Reference := TOBDet
                                              else Reference := TOB(ListPhase.Items[ListPhase.count-1]);
TOBS := TOB.Create ('LA LIGNE',Reference,-1);
// TODO AJOUTER LA MISE A JOUR DES ZONES SUP
AddChampInterm (TOBS);
if (ListPhase <> nil) and (ListPhase.count > 0) then
   BEGIN
   TOBS.PutValue ('PHASE',TOB(ListPhase.Items[ListPhase.count-1]).GetValue('PHASE'));
   TOBS.PutValue ('LIGNEPHASE',TOB(ListPhase.Items[ListPhase.count-1]).GetValue('LIGNEPHASE'));
   END;
if TheTypeLigne = TTpNormal then TOBS.PutValue ('PROVENANCE','DEVIS')
														else TOBS.PutValue ('PROVENANCE','FRAIS');
TOBS.PutValue ('TYPE','OUV');
TOBS.PutValue ('OUVRAGE',TOBL.GetValue('GL_ARTICLE'));
TOBS.PutValue ('LIGNEOUV',TOBL.GetValue('GL_NUMLIGNE'));
TOBS.putValue ('LIBELLE',TOBL.GetValue('GL_LIBELLE'));
(*
if (gestion <> TTGCtrEtude) then
begin
	if fLivChantier then TOBS.PutValue('IDENTIFIANTWOL',-1) else TOBS.PutValue('IDENTIFIANTWOL',0);
end else
begin
	TOBS.PutValue('IDENTIFIANTWOL',TOBL.GetValue('GL_IDENTIFIANTWOL'));
end;
*)
if fLivChantier then TOBS.PutValue('IDENTIFIANTWOL',-1)
                else TOBS.PutValue('IDENTIFIANTWOL',0);

//TOBS.putValue ('IDENTIFIANTWOL',TOBL.GetValue('GL_IDENTIFIANTWOL'));
TOBS.putValue ('WOL ORIGINE',TOBL.GetValue('GL_IDENTIFIANTWOL'));
result := TOBS;
end;

function TOF_BTPLANNIFCH.findlaTOB (TOBL,TOBDepart: TOB;ListPhase : Tlist): TOB;
var TOBDebut : TOB;
    Phase : string;
    LignePhase : integer;
    Fournisseur : string;
begin
	if (ListPhase = nil) or (ListPhase.count = 0) then
  BEGIN
    TOBDebut := TOBDepart;
    Phase := '';
    LignePhase := 0;
  END else
  BEGIN
    TOBDebut := TOB(ListPhase.Items [ListPhase.Count -1]);
    Phase := TOBDEBUT.GetValue('PHASE');
    LignePhase := TOBDebut.GetValue('LIGNEPHASE');
  END;
  if TOBL.GetValue('GLC_NATURETRAVAIL')='002' then
  begin
    Fournisseur := TOBL.GetValue('GL_FOURNISSEUR');
  end else
  begin
    Fournisseur := '';
  end;

  if TOBL.GetValue ('GL_PIECEPRECEDENTE') <> '' then
  Begin
    result := TOBDebut.findfirst (['PHASE','LIGNEPHASE','OUVRAGE','LIGNEOUV','ARTICLE','LIBELLE','NATURETRAVAIL','FOURNISSEUR'],
              [Phase,LignePhase,CurrentOuv.GetValue('GL_ARTICLE'),CurrentOuv.GetValue('GL_NUMLIGNE'),
              TOBL.GetValue('GL_ARTICLE'),TOBL.GetValue('GL_LIBELLE'),TOBL.GetValue('GLC_NATURETRAVAIL'),Fournisseur],true);
  end else
  Begin
    result := TOBDebut.findfirst (['PHASE','LIGNEPHASE','OUVRAGE','LIGNEOUV','ARTICLE','LIBELLE','NATURETRAVAIL','FOURNISSEUR'],
              [Phase,LignePhase,CurrentOuv.GetValue('GL_ARTICLE'),CurrentOuv.GetValue('GL_NUMLIGNE'),
              TOBL.GetValue('GL_ARTICLE'),TOBL.GetValue('GL_LIBELLE'),TOBL.GetValue('GLC_NATURETRAVAIL'),Fournisseur],true);
  end;
end;

procedure TOF_BTPLANNIFCH.CumuleInterm (TOBDevis,TOBL,TOBI : TOB;TheTypeLigne : TTypeProv);
var TOBLIG : TOB;
    TypeArt : string;
begin
TypeArt := TOBL.GetValue('GL_TYPEARTICLE');

TOBI.PutValue ('QTE',TOBL.GetValue('GL_QTEFACT')+TOBI.GetValue('QTE'));
TOBI.PutValue ('MONTANTPV',TOBI.GetValue('MONTANTPV')+(TOBL.GetValue('GL_QTEFACT')*TOBL.GetValue('GL_PUHTDEV')));
//TOBI.PutValue ('MONTANTPA',TOBI.GetValue('MONTANTPA') + (TOBL.GetValue('GL_QTEFACT')*TOBL.GetValue('GL_DPA')));
//TOBI.PutValue ('MONTANTPR',TOBI.GetValue('MONTANTPR')+(TOBL.GetValue('GL_QTEFACT')*TOBL.GetValue('GL_DPR')));
TOBI.PutValue ('MONTANTPA',TOBI.GetValue('MONTANTPA') + (TOBL.GetValue('GL_MONTANTPA')));
TOBI.PutValue ('MONTANTPR',TOBI.GetValue('MONTANTPR')+(TOBL.GetValue('GL_MONTANTPR')));
TOBI.PutValue ('MONTANTFR',TOBI.GetValue('MONTANTFR') + TOBL.GetValue('GL_MONTANTFR'));
TOBI.PutValue ('MONTANTFC',TOBI.GetValue('MONTANTFC') + TOBL.GetValue('GL_MONTANTFC'));
TOBI.PutValue ('MONTANTFG',TOBI.GetValue('MONTANTFG') + TOBL.GetValue('GL_MONTANTFG'));
if TOBI.GetValue('QTE') <> 0 then TOBI.PutValue ('DPA',Arrondi(TOBI.GetValue('MONTANTPA') / TOBI.GetValue('QTE'),V_PGI.okdecQ));
if TOBI.GetValue('QTE') <> 0 then TOBI.PutValue ('PV',Arrondi(TOBI.GetValue('MONTANTPV')/ TOBI.GetValue('QTE'),V_PGI.OkDecP));
if TOBI.GetValue('QTE') <> 0 then TOBI.PutValue ('DPR',Arrondi(TOBI.GetValue('MONTANTPR')/ TOBI.GetValue('QTE'),V_PGI.OkDecP));
{
TOBLIG := TOB.Create ('LIGNE',TOBI,-1);
AddLesSupLigne (TOBLIG,false);
}
TOBLIG:=NewTobLigne(TOBI,0);
//AddLesReferences (TOBLig,TOBL);
TOBLIG.dupliquer (TOBL,false,true);
TOBLIG.AddChampSUpValeur ('GP_REFINTERNE',TOBDevis.GetValue('GP_REFINTERNE'));
if (TypeArt = 'OUV') or (TypeArt = 'ARP') then TOBLIG.PutValue('GL_INDICENOMEN',0);
end;

procedure TOF_BTPLANNIFCH.ReajusteLigneDevisPourCtrEtude(TOBL : TOB; CoefFr : double);
var MontantAch,MtFG,MtFR,MTFC,QTe : double;
begin
  Qte := TOBL.GetValue('GL_QTEFACT')/ TOBL.GetValue('GL_PRIXPOURQTE');
  MontantAch := TOBL.getValue('GL_MONTANTPA');
  MtFG := TOBL.GetValue('GL_MONTANTFG');
  MtFC := 0;
  if TOBL.GetValue('GLC_NONAPPLICFRAIS')<>'X' then
  begin
    MtFR := Arrondi((MontantAch+MTFG+MTFC) * Coeffr,4);
  end;
  TOBL.putValue('GL_COEFFC',0);
  TOBL.putValue('GL_COEFFR',COEFFR);
  TOBL.putValue('GL_MONTANTPAFR',MontantAch);
  TOBL.putValue('GL_MONTANTPAFG',MontantAch);
  TOBL.putValue('GL_MONTANTPAFC',0);
  TOBL.putValue('GL_MONTANTFG',MTFG);
  TOBL.putValue('GL_MONTANTFR',MTFR);
  TOBL.putValue('GL_MONTANTFC',MTFC);
  TOBL.putValue('GL_MONTANTPR',MontantAch + MtFG+ MTFC+MtFR);
  if QTe <> 0 then TOBL.putValue('GL_DPR',Arrondi(TOBL.GetValue('GL_MONTANTPR')/Qte,V_PGI.okdecP));
  if TOBL.GetValue('GL_MONTANTPR')<>0 then TOBL.putValue('GL_COEFMARG',Arrondi(TOBL.GetValue('GL_MONTANTHT')/TOBL.GetValue('GL_MONTANTPR'),4));
end;

procedure TOF_BTPLANNIFCH.AppliqueCoeffRSurLigneFrais (TOBL : TOB;Coeff : double);
var MontantAch,CoeffA,Coeffg,MtFG,MtFR,QTe : double;
begin
  TOBL.putValue('GL_COEFFR',Coeff);
  Qte := TOBL.GetValue('GL_QTEFACT')/ TOBL.GetValue('GL_PRIXPOURQTE');
  MontantAch := Arrondi(TOBL.GetValue('GL_PUHTDEV') * Qte,4);
  MtFR := Arrondi(MontantAch * Coeff,4);
  MtFG := 0;
  TOBL.putValue('GL_MONTANTPA',MontantAch);
  TOBL.putValue('GL_MONTANTPAFR',MontantAch);
  TOBL.putValue('GL_MONTANTPAFG',MontantAch);
  TOBL.putValue('GL_MONTANTPAFC',0);
  TOBL.putValue('GL_MONTANTFG',MTFG);
  TOBL.putValue('GL_MONTANTFR',MtFR);
  TOBL.putValue('GL_MONTANTFC',0);
  TOBL.putValue('GL_COEFFR',Coeff);
  TOBL.putValue('GL_COEFFC',0);
  TOBL.putValue('GL_COEFFG',0);
  TOBL.putValue('GL_MONTANTPR',MontantAch + MtFG+ MtFR);
  TOBL.putValue('GL_DPA',Arrondi(TOBL.GetValue('GL_MONTANTPA')/Qte,V_PGI.okdecP));
  TOBL.putValue('GL_DPR',Arrondi(TOBL.GetValue('GL_MONTANTPR')/Qte,V_PGI.okdecP));
  TOBL.putValue('GL_PUHTDEV',0);
  TOBL.putValue('GL_COEFMARG',0);
end;

function TOF_BTPLANNIFCH.GenereTobIntermediaire: boolean;
var Indice,Indice1 : integer;
    TOBL,TOBI : TOB;
    TOBDet,TOBPlat : TOB;
    CodeArticle : string;
begin
Result := true;
//
MemoriseChampsSupLigneETL ('PBT' ,True);
MemoriseChampsSupLigneOUV ('PBT');
MemoriseChampsSupPIECETRAIT;
//

TOBPlat := TOB.Create ('LIGNE',nil,-1);
TRY
for Indice := 0 to TobDevis.detail.count -1 do
    begin
    TOBL := TOBDevis.detail[Indice];
    // VARIANTE
    if IsVariante (TOBL) then continue;
    //
    if (TOBL.GetString('GLC_NATURETRAVAIL')='001') then continue;
    // --
    TOBL.PutValue('GL_DATELIVRAISON',iDate2099);
    AddLesSupLigne (TOBL,false);
    if ((ModeTraitement and 1)=1) and (Indice = 0) then
    begin
      TOBI := TOB.Create ('LA LIGNE',TOBINTERMEDIAIRE,-1);
      AddChampInterm (Tobi);
      TOBI.PutValue('TYPE','PHA');
      TOBI.PutValue('LIBELLE','Détail des travaux');
      TOBI.PutValue ('PROVENANCE','DEVIS');
      if fLivChantier then TOBI.PutValue('IDENTIFIANTWOL',-1)
                      else TOBI.PutValue('IDENTIFIANTWOL',0);
      ListPhase.Add (TOBI);
    end;
    if (copy(TOBL.GetValue('GL_TYPELIGNE'), 1,2)='DP') and ((ModeTraitement and 2)=2) then
       BEGIN
       TOBI := InserePhaseInterm (TOBDevis,TOBL,TOBINTERMEDIAIRE);
       ListPhase.Add (TOBI);
       continue;
       END;
    if (copy(TOBL.GetValue('GL_TYPELIGNE'),1,2)='TP') and ((ModeTraitement and 1)<>1) and (ListPhase.count > 0) then
       BEGIN
       ListPhase.delete (ListPhase.Count-1);
       continue;
       END;
    if ((TOBL.GetValue('GL_TYPEARTICLE')='OUV') OR (TOBL.GetValue('GL_TYPEARTICLE')='ARP')) and (TOBL.GetValue('GL_INDICENOMEN')>0) then
       BEGIN
        TOBDet := TOBIntermediaire; // pointe sur le pere
        if ((ModeTraitement and 4)=4) then
        begin
          CurrentOuv.dupliquer(TOBL,false,true);
          TOBI := InsereOuvrageInterm (TOBL,TOBDet,ListPhase); // pointe la fille
          TobDet := TOBI;
        end;
        TRY
          MiseAplatouv (TOBdevis,TOBL,TOBOuvrage,TOBPLat,true,FALSE,true,true);
          for indice1 := 0 to TOBPlat.detail.count -1 do
          begin
           // VARIANTE
           if IsVariante (TOBPLat.detail[Indice1]) then continue;
       		 if TOBPLat.detail[Indice1].GetValue('GL_QTEFACT') = 0 then continue;
           //gestion de la non-prise en compte des lignes Cotraitantes
           if (TOBPLat.detail[Indice1].GetString('GLC_NATURETRAVAIL')='001') then continue;
           if (TOBPLat.detail[Indice1].GetValue('GL_TYPEARTICLE')='PRE') then
           begin
       		 	TOBPLat.detail[Indice1].PutValue('BNP_TYPERESSOURCE',RenvoieTypeRes(TOBPLat.detail[Indice1].GetValue('GL_ARTICLE')));
           end;
           if TOBL.GEtValue('GLC_DOCUMENTLIE')<> '' then
           begin
       		 	TOBPLat.detail[Indice1].PutValue('GLC_DOCUMENTLIE',TOBL.GEtValue('GLC_DOCUMENTLIE'));
           end;
           // -- Recalcul des elements de la ligne en enlevant les frais de chantier
           ReajusteLigneDevisPourCtrEtude(TOBPLat.detail[Indice1],TOBdevis.getValue('GP_COEFFR'));
           // --
           CodeArticle := TOBPlat.detail[Indice1].GetValue('GL_CODEARTICLE');
           if (TOBPLat.detail[Indice1].getvalue('GL_TYPEARTICLE')='POU') then RemplacePourcent(TOBPlat.detail[Indice1],TOBDevis);
           if ((ModeTraitement and 4)=4) then TOBI := findlaTOB (TOBPlat.detail[Indice1],TOBDet,nil)
                                         else TOBI := findlaTOB (TOBPlat.detail[Indice1],TOBDet,ListPhase);
           if TOBI = nil then
           BEGIN
            if ((ModeTraitement and 4)=4) then InsereInterm (TobDevis,TOBPlat.detail[indice1],TOBDet,nil)
                                          else InsereInterm (TobDevis,TOBPlat.detail[Indice1],TOBDet,ListPhase);

           END ELSE
           BEGIN
              CumuleInterm (TOBDevis,TOBPlat.detail[Indice1],TOBI);
           end;
          end;
        FINALLY
          TOBPlat.clearDetail;
          CurrentOuv.InitValeurs;
        END;
        continue;
       END;

    if (TOBL.GetValue('GL_TYPEARTICLE')='MAR') OR
    	 (TOBL.GetValue('GL_TYPEARTICLE')='PRE') or
       (TOBL.GetValue('GL_TYPEARTICLE')='FRA') or
       ((TOBL.GetValue('GL_TYPEARTICLE')='ARP') and (TOBL.GetValue('GL_INDICENOMEN')=0)) then
       BEGIN
       if TOBL.GetValue('GL_QTEFACT') = 0 then continue;
       ReajusteLigneDevisPourCtrEtude(TOBL,TOBdevis.getValue('GP_COEFFR'));
       TOBI := findlaTOB (TOBL,TOBIntermediaire,ListPhase);
       if TOBI = nil then
          BEGIN
          InsereInterm (TOBDevis,TOBL,TOBIntermediaire,ListPhase);
          END ELSE
          BEGIN
          CumuleInterm (TOBDevis,TOBL,TOBI);
          end;
       END;
    if (TOBL.GetValue('GL_TYPEARTICLE')='POU') then
       BEGIN
       // remplacement de l'article de type pourcentage en article std marchandise
       RemplacePourcent(TOBL,TOBDevis);
       ReajusteLigneDevisPourCtrEtude(TOBL,TOBdevis.getValue('GP_COEFFR'));
       TOBI := findlaTOB (TOBL,TOBIntermediaire,ListPhase);
       if TOBI = nil then
          BEGIN
          InsereInterm (TOBDevis,TOBL,TOBIntermediaire,ListPhase);
          END ELSE
          BEGIN
          CumuleInterm (TOBDevis,TOBL,TOBI);
          end;
       END;
    end;

    //paragraphe à blanc...
    SupprimeOuvrageVide(TOBIntermediaire);

    if ((ModeTraitement and 1)=1) then
    begin
       ListPhase.delete (ListPhase.Count-1);
    end;

// generation des frais de chantier
for Indice := 0 to TobFrais.detail.count -1 do
    begin
    TOBL := TOBFrais.detail[Indice];
    // VARIANTE
    if IsVariante (TOBL) then continue;
    // --
    TOBL.PutValue('GL_DATELIVRAISON',iDate2099);
    AddLesSupLigne (TOBL,false);
    if ((ModeTraitement and 1)=1) and (Indice = 0) then
    begin
      TOBI := TOB.Create ('LA LIGNE',TOBINTERMEDIAIRE,-1);
      AddChampInterm (Tobi);
      TOBI.PutValue('TYPE','PHA');
      TOBI.PutValue('LIBELLE','Frais de chantier');
      TOBI.PutValue ('PROVENANCE','FRAIS');
      if fLivChantier then TOBI.PutValue('IDENTIFIANTWOL',-1)
                      else TOBI.PutValue('IDENTIFIANTWOL',0);
      ListPhase.Add (TOBI);
    end;
    if (copy(TOBL.GetValue('GL_TYPELIGNE'), 1,2)='DP') and ((ModeTraitement and 2)=2) then
       BEGIN
       TOBI := InserePhaseInterm (TOBfrais,TOBL,TOBINTERMEDIAIRE,TtpFrais);
       ListPhase.Add (TOBI);
       continue;
       END;
    if (copy(TOBL.GetValue('GL_TYPELIGNE'),1,2)='TP') and ((ModeTraitement and 1)<>1) and (ListPhase.count > 0) then
       BEGIN
       ListPhase.delete (ListPhase.Count-1);
       continue;
       END;
    if (TOBL.GetValue('GL_TYPEARTICLE')='OUV') OR (TOBL.GetValue('GL_TYPEARTICLE')='ARP') then
       BEGIN
        TOBDet := TOBIntermediaire; // pointe sur le pere
        if ((ModeTraitement and 4)=4) then
        begin
          CurrentOuv.dupliquer(TOBL,false,true);
          TOBI := InsereOuvrageInterm (TOBL,TOBDet,ListPhase,TtpFrais); // pointe la fille
          TobDet := TOBI;
        end;
        TRY
          MiseAplatouv (TOBFrais,TOBL,TOBOuvrageFrais,TOBPLat,true,false,true,true);
          for indice1 := 0 to TOBPlat.detail.count -1 do
          begin
           // VARIANTE
           if IsVariante (TOBPLat.detail[Indice1]) then continue;
       		 if TOBPLat.detail[Indice1].GetValue('GL_QTEFACT') = 0 then continue;
           // --
          // pour mettre en place le coef de FR du devis
           AppliqueCoeffrSurLigneFrais (TOBPlat.detail[Indice1],TOBdevis.getValue('GP_COEFFR'));
           if (TOBPLat.detail[Indice1].getvalue('GL_TYPEARTICLE')='POU') then RemplacePourcent(TOBPlat.detail[Indice1],TOBDevis);
           if ((ModeTraitement and 4)=4) then TOBI := findlaTOB (TOBPlat.detail[Indice1],TOBDet,nil)
                                         else TOBI := findlaTOB (TOBPlat.detail[Indice1],TOBDet,ListPhase);
           if TOBI = nil then
           BEGIN
            if ((ModeTraitement and 4)=4) then InsereInterm (TobDevis,TOBPlat.detail[indice1],TOBDet,nil,TtpFrais)
                                          else InsereInterm (TobDevis,TOBPlat.detail[Indice1],TOBDet,ListPhase,TtpFrais);

           END ELSE
           BEGIN
              CumuleInterm (TOBDevis,TOBPlat.detail[Indice1],TOBI,TtpFrais);
           end;
          end;
        FINALLY
          TOBPlat.clearDetail;
          CurrentOuv.InitValeurs;
        END;
        continue;
       END;
    if (TOBL.GetValue('GL_TYPEARTICLE')='MAR') OR
    	 (TOBL.GetValue('GL_TYPEARTICLE')='PRE') OR
    	 (TOBL.GetValue('GL_TYPEARTICLE')='FRA') then
       BEGIN
       if TOBL.GetValue('GL_QTEFACT') = 0 then continue;
      // pour mettre en place le coef de FG du devis
       AppliqueCoeffRSurLigneFrais (TOBL,TOBdevis.getValue('GP_COEFFR'));
       TOBI := findlaTOB (TOBL,TOBIntermediaire,ListPhase);
       if TOBI = nil then
          BEGIN
          InsereInterm (TOBDevis,TOBL,TOBIntermediaire,ListPhase,TtpFrais);
          END ELSE
          BEGIN
          CumuleInterm (TOBDevis,TOBL,TOBI,TtpFrais);
          end;
       END;
    if (TOBL.GetValue('GL_TYPEARTICLE')='POU') then
       BEGIN
       // remplacement de l'article de type pourcentage en article std marchandise
       RemplacePourcent(TOBL,TOBDevis);
       TOBI := findlaTOB (TOBL,TOBIntermediaire,ListPhase);
       if TOBI = nil then
          BEGIN
          InsereInterm (TOBDevis,TOBL,TOBIntermediaire,ListPhase,TtpFrais);
          END ELSE
          BEGIN
          CumuleInterm (TOBDevis,TOBL,TOBI,TtpFrais);
          end;
       END;
    end;
FINALLY
TOBPLAT.free;
listPhase.Clear;
end;
end;

procedure IndiceLaTob(TOBTrait : TOB; var Numligne : integer;NumPere : integer);
var Indice: integer;
    TOBCurr : TOB;
    NumeroSuite : integer;
begin
Indice := 0;
if TOBTrait.detail.count = 0 then exit;
if TOBTrait.NomTable = 'LIGNE' then exit;  // c'est trop bas laaa..
repeat
    TOBCurr := TOBTrait.detail[Indice];
    TOBCurr.PutValue('NUMLIGNE',Numligne);
    NumeroSuite := NumLigne;
    TOBCurr.PutValue('PADRE',NumPere);
    inc(numligne);
    if (TOBCurr.GetValue('TYPE')='PHA') or (TOBCurr.getValue('TYPE')='OUV') then
       begin
       IndiceLaTob(TOBCurr,Numligne,NumeroSuite);
       end;
    inc(indice);
until Indice > TOBTrait.detail.count -1;
end;

procedure TOF_BTPLANNIFCH.IndiceTobIntermediaire;
var NumLigne: integer;
begin
Numligne := 1;
IndiceLaTob(TobIntermediaire,Numligne,0);
end;

procedure TOF_BTPLANNIFCH.InitTobs;
begin
TOBIntermediaire := nil;
TOBChantier := nil;
CurrentOuv := nil;
TOBGenere := nil;
TOBLiaison := nil;
end;

procedure TOF_BTPLANNIFCH.CreeTob;
begin
TOBresult := TOB.create ('LE RESULTAT',nil,-1);
TOBIntermediaire := TOB.Create ('PRESENTATION',nil,-1);
AddChampInterm (TOBIntermediaire);
TOBChantier := TOB.Create ('CHANTIER',nil,-1);
CurrentOuv := TOB.Create ('LIGNE',nil,-1);
TOBGenere := TOB.Create ('LES LIGNES',nil,-1);
TOBLiaison := TOB.Create ('LES LIENDEVCHA',nil,-1);
TOBA := TOB.Create ('ARTICLE',nil,-1);
TOBAdresses := TOB.create ('LES ADRESSES',nil,-1);
end;

procedure TOF_BTPLANNIFCH.FreeTobs;
begin
TOBIntermediaire.free;
TOBChantier.free;
CurrentOuv.free;
TOBGenere.free;
TOBLiaison.free;
TOBA.free;
TOBADresses.Free;
end;

procedure TOF_BTPLANNIFCH.AffecteImage (NodeSuite : TTreeNode ; TOBSuite : TOB);
begin
  if TOBSuite = nil then
  begin
    NodeSuite.ImageIndex := 0;
    NodeSuite.SelectedIndex := 0; 
  end
  else if TOBSUite.GetValue('IDENTIFIANTWOL') = 0 then
  begin
    NodeSuite.ImageIndex := 1;  // Chez nous (dépot...)
    NodeSuite.SelectedIndex := 1;  // Chez nous (dépot...)
  end
  else if TOBSUite.GetValue('IDENTIFIANTWOL') = -1 then
  begin
    NodeSuite.ImageIndex := 2; // chez client
    NodeSuite.SelectedIndex := 2;  // Chez nous (dépot...)
  end;
end;

procedure TOF_BTPLANNIFCH.chargementTV_DEVIS (TV : TTreeView;Pere:TTreeNode;TobATraiter : TOB);
VAR Indice : integer;
    TOBSuite : TOB;
    NodeSuite : TTreeNode;
begin
if TOBAtraiter.detail.count = 0 then exit;
if TOBATraiter.NomTable = 'LIGNE' then exit;
for Indice := 0 to TOBATraiter.detail.count -1 do
    begin
    TOBSuite := TOBATraiter.detail[Indice];
    NodeSuite := TV.items.Addchild (Pere, DefiniLibelle(TOBSuite));
    NodeSuite.data := ToBSuite;
    AffecteImage (NodeSuite,TOBSuite);
    if (TOBSUITE.GetValue('TYPE')='PHA') or (TOBSUITE.GetValue('TYPE')='OUV') then
       begin
       chargementTV_DEVIS (TV_DEVIS,NodeSuite,TobSuite);
       end;
    end;
end;

procedure TOF_BTPLANNIFCH.ChargeComponents;
begin
  // initialisation de base
  Rootdevis := TV_DEVIS.items.Add (nil,'DEVIS');
  RootDevis.ImageIndex := 0;
  RootDevis.SelectedIndex := 0;
  Rootchantier := TV_CHANTIER.items.Add (nil,'CHANTIER');
  Rootchantier.ImageIndex := 0;
  Rootchantier.SelectedIndex := 0;
  // chargement du TV_DEVIS
  IF gestion = TTGdevis then
  begin
    chargementTV_DEVIS (TV_DEVIS,Rootdevis,TobIntermediaire);
    TV_DEVIS.FullExpand;
    TV_DEVIS.Refresh;
  end else
  begin
    chargementTV_DEVIS (TV_CHANTIER,Rootchantier,TobIntermediaire);
    TV_CHANTIER.FullExpand;
    TV_CHANTIER.Refresh;
  end;
end;

procedure TOF_BTPLANNIFCH.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPLANNIFCH.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPLANNIFCH.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTPLANNIFCH.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTPLANNIFCH.ChargeListImage;
var UneImage : Timage;
begin
  if TImage(GetControl('LIVCHEZCLI')) <> nil then
  begin
    UneImage := TImage(GetControl('LIVAUCUN'));
    if TheListImage.add(TBitMap(UneImage.Picture.Bitmap ),nil) < 0 then exit;
    UneImage := TImage(GetControl('LIVCHEZNOUS'));
    if TheListImage.add(TBitMap(UneImage.Picture.Bitmap ),nil) < 0 then exit;
    UneImage := TImage(GetControl('LIVCHEZCLI'));
    if TheListImage.add(TBitMap(UneImage.Picture.Bitmap),nil) < 0 then exit;
    //
    TV_DEVIS.Images := TheListImage;
    TV_CHANTIER.Images := TheListImage;
  end;
end;

procedure TOF_BTPLANNIFCH.OnArgument (S : String ) ;
var Critere : string ;
    ChampMul,ValMul : string;
    x: integer ;
begin
Inherited ;
fLivChantier := false;
if GetParamSoc ('SO_BTLIVCHANTIER') then
begin
  fLivChantier := True;
end;
TheListImage := TImageList.Create (TForm(ecran));
InitTobs;
CreeTob;
TOBDEvis := Latob;
EnHt := (TOBDEVIS.GetValue('GP_FACTUREHT')='X') ;
G_ChargeLesAdresses (TOBDevis);
DEV.Code:=TOBDEVIS.GetValue('GP_DEVISE') ; GetInfosDevise(DEV) ;
CoefFR := GetCoefFR (TOBDevis);
LivrFournChantier := TCheckBox(GetControl('LIVRFOURNCH'));
LivrFournDevis := TCheckBox(GetControl('LIVRFOURNDEV'));
if TOBDevis.GetValue('GP_IDENTIFIANTWOT') = -1 then
begin
  LivrFournDevis.checked := true;
  LivrFournChantier.checked := true;
end;
TOBOuvrage := TOb(LaTob.data);
TOBFrais := TOB(TOBOuvrage.data);
TOBOuvrageFrais := TOB(TOBFrais.data);
TOBSST := TOB(TOBOuvrageFrais.data);
// Reinit
(*TOBFrais.data := nil;
TOBOuvrage.data := nil;*)
// --
ListPhase := Tlist.Create;
ChargeReplPourcent;
Repeat
 Critere:=uppercase(Trim(ReadTokenSt(S))) ;
 if Critere<>'' then
 begin
    x:=pos('=',Critere);
    if x<>0 then
    begin
      ChampMul:=copy(Critere,1,x-1);
      ValMul:=copy(Critere,x+1,length(Critere));
      if ChampMul='MODE' then ModeTraitement := strtoint(Valmul) else
      if ChampMul='GESTION' then
      begin
      	if Valmul = 'DEVIS' then Gestion := TTGDevis
        										else gestion := TTGCtrEtude;
      end;
    end;
 end;
until  Critere='';
GetComponents;
ChargeListImage;
SetMethods;
GenereTobIntermediaire;
IndiceTobIntermediaire;
ChargeComponents;
TraitementModeGestion;
end ;

procedure TOF_BTPLANNIFCH.ChargeReplPourcent;
var QQ : TQuery;
begin
  TheArticle := GetParamSoc ('SO_BTREPLPOURCENT');

  TOBA := TOB.Create ('ARTICLE',nil,-1);
  QQ := OpenSql ('SELECT * FROM ARTICLE WHERE GA_CODEARTICLE="'+TheArticle+'"',true,-1,'',true);
  TOBA.selectDb ('',QQ);
  ferme(QQ);
end;

procedure TOF_BTPLANNIFCH.OnClose ;
begin
  Inherited ;
TheListImage.Free;
if ListPhase <> nil then listPhase.Free;
FreeTobs;
end ;

procedure TOF_BTPLANNIFCH.RemplacePourcent(TOBL,TOBPiece: TOB);
begin
  TOBL.PutValue('GL_CODEARTICLE',TOBA.GetValue('GA_CODEARTICLE'));
  TOBL.PutValue('GL_ARTICLE',TOBA.GetValue('GA_ARTICLE'));
  TOBL.PutValue('GL_REFARTSAISIE',TOBA.GetValue('GA_CODEARTICLE'));
  TOBL.PutValue('GL_TYPEARTICLE',TOBA.GetValue('GA_TYPEARTICLE'));
  TOBL.PutValue('GL_QTEFACT',1);
  TOBL.PutValue('GL_QUALIFQTEVTE',TOBA.GetVAlue('GA_QUALIFUNITEVTE'));
  TOBL.PutValue('GL_PRIXPOURQTE',1);
  TOBL.Putvalue('GL_RECALCULER','X');
  if TOBPiece.GetValue('GP_FACTUREHT')='X' then
  begin
    TOBL.PutValue('GL_PUHTDEV',TOBL.GetValue('GL_MONTANTHTDEV'));
    CalculeLigneHT (TOBL,nil,TOBPiece,DEV,DEV.decimale);
  end else
  begin
    TOBL.PutValue('GL_PUTTCDEV',TOBL.GetValue('GL_MONTANTTTCDEV'));
    CalculeLigneTTC (TOBL,nil,TOBPiece,DEV,DEV.decimale);
  end;
end;

procedure TOF_BTPLANNIFCH.TraiteCumuleEltChantier(NodeDepart: TTreeNode);

  procedure CumuleThisTOB (ThisTOb,TOBDest : TOB);
  var indiceLoc : integer;
      TOBL,TOBIns : TOB;
      TypeArt : String;
  begin
    for IndiceLoc := 0 to ThisTOB.detail.count -1 do
    begin
      TOBL := ThisTOB.detail[IndiceLoc];
      TypeArt := TOBL.GetValue('GL_TYPEARTICLE');
      TOBDest.PutValue ('QTE',TOBL.GetValue('GL_QTEFACT')+TOBDest.GetValue('QTE'));
      TOBDest.PutValue ('MONTANTFR',TOBL.GetValue('GL_MONTANTFR')+TOBDest.GetValue('MONTANTFR'));
      TOBDest.PutValue ('MONTANTFC',TOBL.GetValue('GL_MONTANTFC')+TOBDest.GetValue('MONTANTFC'));
      TOBDest.PutValue ('MONTANTFG',TOBL.GetValue('GL_MONTANTFG')+TOBDest.GetValue('MONTANTFG'));
      TOBDest.PutValue ('MONTANTPA',TOBL.GetValue('GL_MONTANTPA')+TOBDest.GetValue('MONTANTPA'));
      TOBDest.PutValue ('MONTANTPR',TOBL.GetValue('GL_MONTANTPR')+TOBDest.GetValue('MONTANTPR'));
      TOBDest.PutValue ('MONTANTPV',TOBL.GetValue('GL_MONTANTHTDEV')+TOBDest.GetValue('MONTANTPV'));
{
      TOBIns := TOB.Create ('LIGNE',TOBDest,-1);
      AddLesSupLigne (TOBIns,false);
}
      TOBIns:=NewTobLigne(TOBDest,0);
//      AddLesReferences (TOBIns,TOBL);
      TOBIns.dupliquer (TOBL,false,true);
      // Correction dans le cas de TTC
      TOBIns.putValue ('GL_FACTUREHT','X');
      TOBIns.putValue ('GL_IDENTIFIANTWOL',ThisTob.getValue('IDENTIFIANTWOL'));
      // --
//			if fLivChantier then TOBIns.PutValue('GL_IDENTIFIANTWOL',-1) else TOBIns.PutValue('GL_IDENTIFIANTWOL',0);

      if (TypeArt = 'ARP') or (TypeArt = 'OUV') then TOBIns.PutValue('GL_INDICENOMEN',0);
    end;
    ThisTOB.free;
  end;

  function TOBEqual(TOBCurr,TOBSuiv : TOB; EnHt : boolean) : boolean;
  begin
  result := false;

  if (TOBCurr.GetValue('ARTICLE')=TOBSuiv.getValue('ARTICLE')) and
     (TOBCurr.GetValue('LIBELLE')=TOBSuiv.getValue('LIBELLE')) and
     (TOBCurr.GetValue('NATURETRAVAIL')=TOBSuiv.getValue('NATURETRAVAIL')) and
     (TOBCurr.GetValue('FOURNISSEUR')=TOBSuiv.getValue('FOURNISSEUR')) and
     (TOBCurr.GetValue('DPA')=TOBSuiv.getValue('DPA')) and
     (TOBCurr.GetValue('PV')=TOBSuiv.getValue('PV'))  then Result := true;
  end;

var indice,IndSuiv : integer;
    NodeCurr,NodeSuiv : TTreeNode;
    TOBCUrr,TOBSuiv : TOB;
begin
  Indice := 0;
  repeat
      NodeCurr := NodeDepart.Item[Indice];
      if NodeCurr.data = nil then BEGIN inc(Indice); continue; END;
      if TOB(NodeCurr.data).GetValue('TYPE') <> 'NOR' then (* c'est forcement une phase *)
         begin
         if NodeCurr.count > 0 then
            begin
            TraiteCumuleEltChantier (NodeCurr);
            end;
         inc(Indice);
         Continue;
         end;
      TOBCurr := TOB(NodeCurr.data);
      IndSuiv := Indice+1;
      if IndSuiv <= NodeDepart.count -1 then
      begin
        repeat
          NodeSuiv := NodeDepart.item[IndSuiv];
          if (NodeSuiv.Data = nil) or (TOB(NodeSuiv.data).GetValue('TYPE')<>'NOR') then BEGIN inc(IndSuiv); continue; END;
          TOBSUiv := TOB(NodeSuiv.data);
          if TOBEqual(TOBCurr,TOBSuiv,EnHt) then
          begin
            CumuleThisTOB ( TOBSuiv,TOBCurr); // TOBSUIV.detail -> TOBCurr.detail
            NodeSuiv.data := nil; // reinit a nil pour eviter de la traiter plus tard
          end;
          inc(IndSuiv);
        until IndSuiv > NodeDepart.Count -1;
      end;
      Inc(Indice);
  Until Indice > NodeDepart.count -1;
end;

Procedure TOF_BTPLANNIFCH.G_ChargeLesAdresses ( TOBPiece : TOB ) ;
var i_ind : integer ;
BEGIN
for i_ind:=TOBAdresses.Detail.Count-1 downto 0 do TOBAdresses.Detail[i_ind].Free ;
if GetParamSoc('SO_GCPIECEADRESSE') then
   BEGIN
   TOB.Create('PIECEADRESSE',TOBAdresses,-1) ; {Livraison}
   TOB.Create('PIECEADRESSE',TOBAdresses,-1) ; {Facturation}
   END else
   BEGIN
   TOB.Create('ADRESSES',TOBAdresses,-1) ; {Livraison}
   TOB.Create('ADRESSES',TOBAdresses,-1) ; {Facturation}
   END ;
LoadLesAdresses(TOBPiece,TOBAdresses) ;
TOBAdresses.SetAllModifie(True);
END ;


Procedure TOF_BTPLANNIFCH.AffecteLesAdresses ( TOBPiece,TOBResult : TOB ) ;
var NumL, NumF : integer ;
begin
NumL:=TOBPiece.GetValue('GP_NUMADRESSELIVR'); NumF:=TOBPiece.GetValue('GP_NUMADRESSEFACT');
if GetParamSoc('SO_GCPIECEADRESSE') then
   BEGIN
   TOBResult.PutValue('GP_NUMADRESSELIVR',+1) ;
   if NumF=NumL then TOBResult.PutValue('GP_NUMADRESSEFACT',+1)
                else TOBResult.PutValue('GP_NUMADRESSEFACT',+2);
   END else
   BEGIN
   TOBResult.PutValue('GP_NUMADRESSELIVR',-1) ;
   if NumF=NumL then TOBResult.PutValue('GP_NUMADRESSEFACT',-1)
                else TOBResult.PutValue('GP_NUMADRESSEFACT',-2);
   END ;
end;

function TOF_BTPLANNIFCH.GetCoefFR(TOBDevis: TOB): double;
var Q : TQuery;
		Cledoc : R_Cledoc;
    TOBPorcs,TOBP : TOB;
    Indice : integer;
begin
	result := 0;
	TOBPorcs := TOB.Create ('LES PORTS',nil,-1);
  TRY
    CleDoc := TOB2CleDoc(TOBDevis);
    // Lecture Ports
    Q := OpenSQL('SELECT * FROM PIEDPORT WHERE ' + WherePiece(CleDoc, ttdPorc, False), True,-1,'',true);
    TOBPorcs.LoadDetailDB('PIEDPORT', '', '', Q, False);
    Ferme(Q);
    if TOBPorcs.detail.count > 0 then
    begin
      for Indice := 0 to TOBPorcs.detail.count -1 do
      begin
        TOBP := TOBPorcs.detail[Indice];
        if TOBP.GetValue('GPT_FRAISREPARTIS')='X' then Result := result + (TOBP.getValue('GPT_POURCENT')/100);
      end;
      Result := Result + 1;
    end;
  FINALLY
//    Result := Result + 1;
    TOBPOrcs.free;
  END;
end;

procedure TOF_BTPLANNIFCH.TraitementModeGestion;
begin
   if Gestion = TTGctrEtude then
   begin
//   	THPanel(GetControl('PGAUCHE')).visible := false;
		TTreeView(GetControl('TV_DEVIS')).visible := false;
//   	THPanel(GetControl('PDROIT')).align := alClient;
   end;
end;

procedure TOF_BTPLANNIFCH.formresize(sender: TObject);
begin
	TV_CHANTIER.Invalidate;
end;


Initialization
  registerclasses ( [ TOF_BTPLANNIFCH , TOF_BTMODIFPHASE] ) ;
end.


Unit UTofGCContreMVTOA ;

Interface

uses  StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
      MaineAGL,eMul,
{$ELSE}
      db,dbTables,FE_main,Mul,
{$ENDIF}
      forms,sysutils,
      ComCtrls,Hpanel, Math,HCtrls,HEnt1,HMsgBox,UTOF,Vierge,UTOB,AglInit,LookUp,EntGC,SaisUtil,
      graphics,grids,windows,M3FP,HTB97,Dialogs, AGLInitGC, ExtCtrls, Hqry,LicUtil,HDimension,
      Facture,FactUtil,Menus,Messages,ENT1,FactComm,FactCalc,UTofOptionEdit,StockUtil,LigDispoLot,AssistContreMVtoA,
      FactContreM;

Type
  Tof_GCContreMVTOA = Class (TOF)

    procedure OnArgument (StArgument : String ) ; override ;
    procedure OnUpdate ; override;
    procedure OnClose  ; override ;

    private
    TOBLigne, TOBDispo, TOBDispoCM : TOB ;
    TOBArticles, TOBCatalogu : TOB ;
    TOBGrid : TOB;
    // Grille
    QteATraiterName : String;       // Nom du champ sup. quantité à commander
    QteATraiterTitle : String;     // Nom du champ sup. quantité à commander
    FormuleQteATraiter : String;    // formule permettant d'identifier la colonne champ sup dans la liste
    DefColonnesGrid : String ;
    DefLibelleGrid : String;
    DefColonnesSQL : String;
    GS : THGRID ;
    FSortList : String;
    ColName : string;
    //
    BOuvrir:TToolBarButton97;
    BSelectAll:TToolBarButton97;
    BRechercher:TToolBarButton97;
    BImprimer:TToolBarButton97;
    POPZ: TPopupMenu;
    BZoom : TToolbarButton97;
    BZoomArticle, BZoomCatalogue, BZoomPiece, BZoomClient, BZoomFournisseur : TToolbarButton97;
    TDimension:THLabel;
    FindLigne: TFindDialog;
    FindDebut: Boolean;
    TOTQTE_A_CDE : THNumEdit;
    Mode:String;
    //
    procedure ControlsFichInTof;
    procedure FormKeyDown(Sender: TObject; var Key: Word;  Shift: TShiftState);
    procedure BOuvrirClick(Sender: TObject);
    procedure BSelectAllClick(Sender: TObject);
    procedure BRechercherClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BZoomArticleClick(Sender: TObject);
    procedure BZoomCatalogueClick(Sender: TObject);
    procedure BZoomPieceClick(Sender: TObject);
    procedure BZoomClientClick(Sender: TObject);
    procedure BZoomFournisseurClick(Sender: TObject);
    //
    procedure CreateTOB ;
    procedure CreateTOBLigne ;
    procedure CreateTOBGrid ;
    procedure DestroyTOB ;
    procedure DestroyTOBLigne ;
    procedure DestroyTOBGrid ;
    procedure TOBLigneAddChampSupp;
    procedure TOBLigneConvQteStock;
    function  GetTOBLigne(Row:Integer):TOB ;
    procedure InitGrille ;
    procedure ChargeTOBArticles ;
    procedure ChargeTOBCatalogu ;
    procedure ChargeTOBDispo ;
    procedure ChargeTOBDispoCM ;
//    procedure AffichePied (ACol, ARow : integer);
//    Function  ColListeOk(iCol:Integer):Boolean;
    // Recherche
    procedure FindLigneFind(Sender: TObject);
    // Grille
    procedure DessineCellSEL (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSDblClick(Sender: TObject);
    // Impressions
    procedure Imprime ;
    function  PrepareImpression : integer ;
    // Zoom
    procedure BZoomMouseEnter(Sender: TObject);
    Procedure GereEnabled(ARow: Integer) ;
    procedure VoirPiece;
    procedure VoirClient;
    procedure VoirFournisseur;
    procedure VoirCatalogue;
    //
    procedure CumulQteATraiter;

  end ;

  function  IsArticle(TOBL:TOB):Boolean ;
  function  IsCatalogue(TOBL:TOB):Boolean ;
  Function  FsortByToTobSort(FsortBy:String):String;
  procedure CalculResteACderCM(TOBLigne,TOBDispo,TOBDispoCM:Tob;Champ_Ligne:String);
  procedure CalculPretALivrerCM(TOBLigne,TOBDispo,TOBDispoCM:Tob;Champ_Ligne:String);
  procedure CommandeTraitee (TobLigne : TOB);

Implementation

{===============================================================================}
{------------------------------ TOF --------------------------------------------}
{===============================================================================}
procedure Tof_GCContreMVTOA.OnArgument (StArgument : String ) ;
var Critere,ChampCritere,ValeurCritere:string;
    x:integer;
begin

Mode:='VTOA';

Repeat
 Critere:=Trim(ReadTokenSt(StArgument)) ;
 if Critere<>'' then
   begin
     x:=pos('=',Critere);
     if x<>0 then
       begin
         ChampCritere:=copy(Critere,1,x-1);
         ValeurCritere:=copy(Critere,x+1,length(Critere));
         if ChampCritere='MODE' then Mode := uppercase(ValeurCritere) ;
       end;
   end;
until  Critere='';

if Mode='VTOA' then
   begin
   QteATraiterName := 'QTE_A_CDE' ;              // Champ sup pour quantité à commander
   QteATraiterTitle:='Qté à commander';
   FormuleQteATraiter := ('(GL_QTESTOCK * 1)');  // Champ de la liste associé à la quantité à commander
   Ecran.Caption := 'Contremarque: reste à commander';
   end else if Mode='VTOBL' then
   Begin
   QteATraiterName := 'QTE_A_LIV' ;              // Champ sup pour quantité à commander
   QteATraiterTitle:='Qté livrable';
   FormuleQteATraiter := ('(GL_QTESTOCK * 1)');  // Champ de la liste associé à la quantité à commander
   Ecran.Caption := 'Contremarque: livrable';
   end;

if THValComboBox(GetControl('TMPDOMAINE')).Values.Count > 1 then
    BEGIN
    SetControlVisible('GL_DOMAINE',True);
    SetControlVisible('TGL_DOMAINE',True);
    SetControlProperty('GL_DOMAINE','TabStop',True);
    END;
// Colonnes obligatoires dans le Grid
DefColonnesGrid:='FIXED;GL_TIERS;GL_CODEARTICLE;GL_REFCATALOGUE;GL_LIBELLE;GL_FOURNISSEUR;GL_DEPOT;'+QteATraiterName+';';
DefLibelleGrid:=';Client;Article;Catalogue;Désignation article;Fournisseur;Dépôt;'+QteATraiterTitle+';';

// Champs obligatoires dans la requête SQL
DefColonnesSQL:='GL_TIERS;GL_CODEARTICLE;GL_ARTICLE;GL_LIBELLE;GL_REFCATALOGUE;GL_FOURNISSEUR;GL_DEPOT;'
               +'GL_QTESTOCK;GL_DATELIVRAISON;GL_ENCONTREMARQUE;GL_TYPELIGNE;GL_TYPEREF;'
               +'GL_NATUREPIECEG;GL_SOUCHE;GL_DATEPIECE;GL_NUMERO;GL_INDICEG;GL_NUMLIGNE;GL_NUMORDRE';

Inherited ;

ControlsFichInTof;
CreateTOB;
GS.SetFocus;
SetControlVisible('Fliste',False);
end ;


procedure Tof_GCContreMVTOA.OnUpDate ;
var Sql,Select,Where, ChampsCol,St:String;
    Name:String;
    i_pos : integer;
    QLig:TQuery;
BEGIN
inherited;
InitGrille;
ChampsCol:=Uppercase(DefColonnesSQL);
St:=Uppercase(ColName);
While St <> '' do
    BEGIN
    Name:=ReadTokenSt(St);
    if (Name='FIXED') or (Name=Uppercase(QteATraiterName)) then continue;
    i_pos:=Pos(Name,ChampsCol);
    if i_pos = 0 then ChampsCol:=ChampsCol+';'+Name;
    END;
// SELECT
Select := '';
While ChampsCol <> '' do
    BEGIN
    Name:=ReadTokenSt(ChampsCol);
    if Select='' then Select:='SELECT ' + Name
                 else Select:=Select+', '+Name;
    END;

if pos ('GL_QTERESTE', Select) = 0 then Select := Select + ', GL_QTERESTE'; // DBR NEWPIECE
if pos ('GL_QTEFACT', Select) = 0 then Select := Select + ', GL_QTEFACT'; // DBR NEWPIECE

// WHERE
Where := RecupWhereCritere(TFMul(Ecran).Pages) ;

// Charge la TOB Ligne
DestroyTOBligne ; CreateTOBLigne ;
Sql := Select + ' FROM LIGNE ' + Where + ' AND GL_QTERESTE<>0'; // DBR NEWPIECE Pour GL_QTERESTE <> 0
QLig:=OpenSQL(Sql,True) ;
if not QLig.Eof then TOBLigne.LoadDetailDB('LIGNE','','',QLig,False,True) ;
Ferme(QLig) ;
//
if GS.nbSelected > 0 then GS.ClearSelected;
GS.AllSelected:=False;
BSelectAll.Down := False;

GS.VidePile(False) ;
if TOBLigne.Detail.Count > 0 then
   BEGIN
   TOBLigneAddChampSupp;
   TOBLigneConvQteStock;
   ChargeTOBArticles;
   ChargeTOBCatalogu;
   ChargeTOBDispo;
   ChargeTOBDispoCM;

   If Mode='VTOA' then
   begin
    CommandeTraitee (TobLigne);
    CalculResteACderCM(TOBLigne,TOBDispo,TOBDispoCM,QteATraiterName)
   end else If Mode='VTOBL' then
        CalculPretALivrerCM(TOBLigne,TOBDispo,TOBDispoCM,QteATraiterName);

   //Ordre de tri
   TOBLigne.Detail.sort(FSortList);
   TOBLigne.PutGridDetail(GS, False, False, ColName, True);

   GS.Row := 1; GereEnabled(GS.Row) ;
   //
   end;
CumulQteATraiter;

end;

procedure Tof_GCContreMVTOA.OnClose ;
begin
Inherited ;
FindLigne.Free;
DestroyTOB ;
end ;

///////////////////////////////////////////////////
procedure CalculResteACderCM(TOBLigne,TOBDispo,TOBDispoCM:Tob;Champ_Ligne:String);
var TOBDCM,TOBDIS,TOBTMP,TOBL:TOB;
    Article,Cata,Tiers,Depot,Four:String;
    QteRES,QteATT:Double;
    i:integer;
begin
if TOBLigne.Detail.Count=0 then exit;

TOBDCM:=TOB.Create('',nil,-1);
TOBDIS:=TOB.Create('',nil,-1);
for i:=0 to TOBLigne.Detail.Count - 1 do
   BEGIN
   TOBL:=TOBLigne.Detail[i] ;
   if IsCatalogue(TOBL) then
      BEGIN
      Cata:=TOBL.GetValue('GL_REFCATALOGUE') ;
      Tiers:=TOBL.GetValue('GL_TIERS') ;
      Depot:=TOBL.GetValue('GL_DEPOT');
      Four:=TOBL.GetValue('GL_FOURNISSEUR');
      if TOBDCM.FindFirst(['CATA','TIERS','DEPOT','FOUR'],[Cata,Tiers,Depot,Four],False) = Nil then
         BEGIN
         TOBTMP:=TOB.Create('',TOBDCM,-1);
         TOBTMP.AddChampSupValeur ('CATA', Cata, False);
         TOBTMP.AddChampSupValeur ('TIERS', Tiers, False);
         TOBTMP.AddChampSupValeur ('DEPOT', Depot, False);
         TOBTMP.AddChampSupValeur ('FOUR', Four, False);
         END;
      END ELSE
      BEGIN
      Article:=TOBL.GetValue('GL_ARTICLE') ;
      Depot:=TOBL.GetValue('GL_DEPOT');
      if TOBDCM.FindFirst(['ARTICLE','DEPOT'],[Article,Depot],False) = Nil then
         BEGIN
         TOBTMP:=TOB.Create('',TOBDIS,-1);
         TOBTMP.AddChampSupValeur ('ARTICLE', Article, False);
         TOBTMP.AddChampSupValeur ('DEPOT', Depot, False);
         END;
      END ;
   END ;

for i:=0 to TOBDCM.Detail.Count-1 do
   BEGIN
   if (i=0) then TOBLigne.detail.sort('PIECESUIV;GL_REFCATALOGUE;GL_FOURNISSEUR;GL_DEPOT;GL_TIERS;GL_DATELIVRAISON');

   Cata:=TOBDCM.Detail[i].GetValue('CATA') ;
   Tiers:=TOBDCM.Detail[i].GetValue('TIERS') ;
   Depot:=TOBDCM.Detail[i].GetValue('DEPOT') ;
   Four:=TOBDCM.Detail[i].GetValue('FOUR') ;
   QteRES:=0;//QteATT:=0;
   TOBTMP:= TobDispoCM.FindFirst(['GQC_REFERENCE','GQC_CLIENT','GQC_DEPOT','GQC_FOURNISSEUR'],[Cata,Tiers,Depot,Four],False);
   if TOBTMP <> Nil then
      QteRES:=TOBTMP.getValue('GQC_PHYSIQUE')+TOBTMP.getValue('GQC_RESERVEFOU')-TOBTMP.getValue('GQC_PREPACLI');

   TOBL:= TOBLigne.FindFirst(['GL_REFCATALOGUE','GL_FOURNISSEUR','GL_DEPOT','GL_TIERS'],[Cata,Four,Depot,Tiers],False);
   While (TOBL<>nil) do
      BEGIN
      if (QteRES>0) then
         begin
         QteATT:=Min(TOBL.GetValue('GL_QTESTOCK'), QteRES);
         TOBL.putValue(Champ_Ligne,TOBL.GetValue('GL_QTESTOCK')-QteATT);
         QteRES:=QteRES-QteATT;
         end
      else
         TOBL.putValue(Champ_Ligne,TOBL.GetValue('GL_QTESTOCK'));

      TOBL:= TOBLigne.FindNext(['GL_REFCATALOGUE','GL_FOURNISSEUR','GL_DEPOT','GL_TIERS'],[Cata,Four,Depot,Tiers],False);
      END;

   END;
TOBDCM.Free;

for i:=0 to TOBDIS.Detail.Count-1 do
   BEGIN
   if (i=0) then TOBLigne.detail.sort('GL_ARTICLE;GL_DEPOT;GL_DATELIVRAISON');

   Article:=TOBDIS.Detail[i].GetValue('ARTICLE') ;
   Depot:=TOBDIS.Detail[i].GetValue('DEPOT') ;
   QteRES:=0;//QteATT:=0;
   TOBTMP:= TobDispo.FindFirst(['GQ_ARTICLE','GQ_DEPOT'],[Article,Depot],False);
   if TOBTMP <> Nil then
      QteRES:=TOBTMP.getValue('GQ_PHYSIQUE')+TOBTMP.getValue('GQ_RESERVEFOU')-TOBTMP.getValue('GQ_PREPACLI');

   TOBL:= TOBLigne.FindFirst(['GL_ARTICLE','GL_DEPOT'],[Article,Depot],False);
   While (TOBL<>nil) do
      BEGIN
      if (QteRES>0) then
         begin
         QteATT:=Min(TOBL.GetValue('GL_QTESTOCK'),QteRES);
         TOBL.putValue(Champ_Ligne,TOBL.GetValue('GL_QTESTOCK')-QteATT);
         QteRES:=QteRES-QteATT;
         end
      else
         TOBL.putValue(Champ_Ligne,TOBL.GetValue('GL_QTESTOCK'));

      TOBL:= TOBLigne.FindNext(['GL_ARTICLE','GL_DEPOT'],[Article,Depot],False);
      END;

   END;
TOBDIS.Free;

TOBLigne.detail.sort('-'+Champ_Ligne);
i:=TOBLigne.Detail.Count - 1;
While i>=0 do
   BEGIN
   TOBL:=TOBLigne.Detail[i];
   if (TOBL.GetValue(Champ_Ligne)=0) then
      TOBL.Free
   else if TOBL.GetValue(Champ_Ligne)>0 then
      i:=0;
   i:=i-1;
   END;
end;

procedure CalculPretALivrerCM(TOBLigne,TOBDispo,TOBDispoCM:Tob;Champ_Ligne:String);
var TOBDCM,TOBDIS,TOBTMP,TOBL:TOB;
    Article,Cata,Tiers,Depot,Four:String;
    QteRES,QteATT:Double;
    i:integer;
begin
if TOBLigne.Detail.Count=0 then exit;

TOBDCM:=TOB.Create('',nil,-1);
TOBDIS:=TOB.Create('',nil,-1);
for i:=0 to TOBLigne.Detail.Count - 1 do
   BEGIN
   TOBL:=TOBLigne.Detail[i] ;
   if IsCatalogue(TOBL) then
      BEGIN
      Cata:=TOBL.GetValue('GL_REFCATALOGUE') ;
      Tiers:=TOBL.GetValue('GL_TIERS') ;
      Depot:=TOBL.GetValue('GL_DEPOT');
      Four:=TOBL.GetValue('GL_FOURNISSEUR');
      if TOBDCM.FindFirst(['CATA','TIERS','DEPOT','FOUR'],[Cata,Tiers,Depot,Four],False) = Nil then
         BEGIN
         TOBTMP:=TOB.Create('',TOBDCM,-1);
         TOBTMP.AddChampSup('CATA',False);
         TOBTMP.PutValue('CATA', Cata);
         TOBTMP.AddChampSup('TIERS',False);
         TOBTMP.PutValue('TIERS', Tiers);
         TOBTMP.AddChampSup('DEPOT',False);
         TOBTMP.PutValue('DEPOT', Depot);
         TOBTMP.AddChampSup('FOUR',False);
         TOBTMP.PutValue('FOUR', Four);
         END;
      END ELSE
      BEGIN
      Article:=TOBL.GetValue('GL_ARTICLE') ;
      Depot:=TOBL.GetValue('GL_DEPOT');
      if TOBDCM.FindFirst(['ARTICLE','DEPOT'],[Article,Depot],False) = Nil then
         BEGIN
         TOBTMP:=TOB.Create('',TOBDIS,-1);
         TOBTMP.AddChampSup('ARTICLE',False);
         TOBTMP.PutValue('ARTICLE', Article);
         TOBTMP.AddChampSup('DEPOT',False);
         TOBTMP.PutValue('DEPOT', Depot);
         END;
      END ;
   END ;

for i:=0 to TOBDCM.Detail.Count-1 do
   BEGIN
   if (i=0) then TOBLigne.detail.sort('GL_REFCATALOGUE;GL_FOURNISSEUR;GL_DEPOT;GL_TIERS;GL_DATELIVRAISON');

   Cata:=TOBDCM.Detail[i].GetValue('CATA') ;
   Tiers:=TOBDCM.Detail[i].GetValue('TIERS') ;
   Depot:=TOBDCM.Detail[i].GetValue('DEPOT') ;
   Four:=TOBDCM.Detail[i].GetValue('FOUR') ;
   QteRES:=0;//QteATT:=0;
   TOBTMP:= TobDispoCM.FindFirst(['GQC_REFERENCE','GQC_CLIENT','GQC_DEPOT','GQC_FOURNISSEUR'],[Cata,Tiers,Depot,Four],False);
   if TOBTMP <> Nil then
      QteRES:=TOBTMP.getValue('GQC_PHYSIQUE')-TOBTMP.getValue('GQC_PREPACLI');

   TOBL:= TOBLigne.FindFirst(['GL_REFCATALOGUE','GL_FOURNISSEUR','GL_DEPOT','GL_TIERS'],[Cata,Four,Depot,Tiers],False);
   While (TOBL<>nil) do
      BEGIN
      if (QteRES>0) then
         begin
         QteATT:=Min(TOBL.GetValue('GL_QTESTOCK'),QteRES);
         TOBL.putValue(Champ_Ligne,QteATT);
         QteRES:=QteRES-QteATT;
         end
      else
         TOBL.putValue(Champ_Ligne,0);

      TOBL:= TOBLigne.FindNext(['GL_REFCATALOGUE','GL_FOURNISSEUR','GL_DEPOT','GL_TIERS'],[Cata,Four,Depot,Tiers],False);
      END;

   END;
TOBDCM.Free;

for i:=0 to TOBDIS.Detail.Count-1 do
   BEGIN
   if (i=0) then TOBLigne.detail.sort('GL_ARTICLE;GL_DEPOT;GL_DATELIVRAISON');

   Article:=TOBDIS.Detail[i].GetValue('ARTICLE') ;
   Depot:=TOBDIS.Detail[i].GetValue('DEPOT') ;
   QteRES:=0;//QteATT:=0;
   TOBTMP:= TobDispo.FindFirst(['GQ_ARTICLE','GQ_DEPOT'],[Article,Depot],False);
   if TOBTMP <> Nil then
      QteRES:=TOBTMP.getValue('GQ_PHYSIQUE')-TOBTMP.getValue('GQ_PREPACLI');

   TOBL:= TOBLigne.FindFirst(['GL_ARTICLE','GL_DEPOT'],[Article,Depot],False);
   While (TOBL<>nil) do
      BEGIN
      if (QteRES>0) then
         begin
         QteATT:=Min(TOBL.GetValue('GL_QTESTOCK'),QteRES);
         TOBL.putValue(Champ_Ligne,QteATT);
         QteRES:=QteRES-QteATT;
         end
      else
         TOBL.putValue(Champ_Ligne,0);

      TOBL:= TOBLigne.FindNext(['GL_ARTICLE','GL_DEPOT'],[Article,Depot],False);
      END;

   END;
TOBDIS.Free;
TOBLigne.detail.sort('-'+Champ_Ligne);
i:=TOBLigne.Detail.Count - 1;
While i>=0 do
   BEGIN
   TOBL:=TOBLigne.Detail[i];
   if (TOBL.GetValue(Champ_Ligne)=0) then
      TOBL.Free
   else if TOBL.GetValue(Champ_Ligne)>0 then
      i:=0;
   i:=i-1;
   END;
end;

procedure CommandeTraitee (TOBLigne : Tob);
var TSql : TQuery;
    iInd : integer;
    RefPiece : string;
begin
    for iInd := TobLigne.Detail.Count - 1 downto 0 do
    begin
        RefPiece := EncodeRefPiece(TobLigne.Detail[iInd]);
        if ExisteSql ('SELECT GL_FOURNISSEUR FROM LIGNE WHERE GL_PIECEORIGINE="' + RefPiece +
                      '" AND GL_NATUREPIECEG="CF" AND GL_ARTICLE="' +
                      TobLigne.Detail[iInd].GetValue ('GL_ARTICLE') + '" AND GL_DATEPIECE>="' +
                      USDateTime(TobLigne.Detail[iInd].GetValue ('GL_DATEPIECE')) + '"') then
        begin
            TobLigne.Detail[iInd].AddChampSupValeur ('PIECESUIV', 'A', False);
        end else
        begin
            TobLigne.Detail[iInd].AddChampSupValeur ('PIECESUIV', 'Z', False);
        end;
        Ferme (TSql);
    end;
end;

Function FsortByToTobSort(FsortBy:String):String;
var st,Sc:string;
begin
   result:='';
   St:=FsortBy;
   While St<>'' do
   Begin
      sc :=trim(ReadTokenPipe(St,','));
      if Pos(' DESC',sc)>0 then sc:='-' + FindEtReplace(trim(sc),' DESC','',true);
      if result='' then result := Sc else result:=result+';'+sC;
   end;
end;

{===============================================================================}
{------------------------------ Contrôles de la fiche --------------------------}
{===============================================================================}
Procedure Tof_GCContreMVTOA.ControlsFichInTof;
BEGIN
SetControlProperty('GL_REFCATALOGUE','ElipsisButton',False);
// Grille
GS:=THGRID(GetControl('GS')) ;
GS.PostDrawCell:=DessineCellSEL;
GS.OnRowEnter:=GSRowEnter ;
GS.OnDblClick:=GSDblClick ;
GS.OnRowExit:=GSRowExit ;
//
TFMul(Ecran).OnKeyDown:=FormKeyDown ;
// Boutons divers
BOuvrir := TToolBarButton97(GetControl('BOUVRIR'));
if (BOuvrir<>nil) then BOuvrir.OnClick:=BOuvrirClick;
BOuvrir.visible:=( Mode ='VTOA');
BSelectAll := TToolBarButton97(GetControl('BSELECTALL'));
if (BSelectAll<>nil) then BSelectAll.OnClick:=BSelectAllClick;
BSelectAll.visible:=( Mode ='VTOA');
BImprimer:=TToolBarButton97(GetControl('BIMPRIMER'));
if (BImprimer<>nil) then BImprimer.OnClick:=BImprimerClick;
// Zoom
POPZ:=TPopupMenu(GetControl('POPZ'));
BZoom:=TToolbarButton97(GetControl('BZOOM'));
BZoom.OnMouseEnter:=BZoomMouseEnter;
BZoomArticle:=TToolbarButton97(GetControl('BZOOMARTICLE'));
BZoomArticle.OnClick:=BZoomArticleClick;
BZoomCatalogue:=TToolbarButton97(GetControl('BZOOMCATALOGUE'));
BZoomCatalogue.OnClick:=BZoomCatalogueClick;
BZoomPiece:=TToolbarButton97(GetControl('BZOOMPIECE'));
BZoomPiece.OnClick:=BZoomPieceClick;
BZoomClient:=TToolbarButton97(GetControl('BZOOMCLIENT'));
BZoomClient.OnClick:=BZoomClientClick;
BZoomFournisseur:=TToolbarButton97(GetControl('BZOOMFOURNISSEUR'));
BZoomFournisseur.OnClick:=BZoomFournisseurClick;
// Panel des dimensions
TDimension:=THLabel(GetControl('TDIMENSION'));
// Total des quantités
TOTQTE_A_CDE:=THNumEdit(GetControl('TOTQTE_A_CDE'));
// Bouton de recherche dans le GRID
BRechercher := TToolBarButton97(GetControl('BRECHERCHER'));
if (BRechercher<>nil) then BRechercher.OnClick:=BRechercherClick;
FindLigne:=TFindDialog.Create(Ecran);
FindLigne.Name:='FindLigne';
FindLigne.OnFind:=FindLigneFind;
END;

procedure Tof_GCContreMVTOA.FormKeyDown(Sender: TObject; var Key: Word;  Shift: TShiftState);
BEGIN
END;

procedure Tof_GCContreMVTOA.BSelectAllClick(Sender: TObject);
BEGIN
GS.ClearSelected ;
GS.AllSelected:=BSELECTALL.Down ;
END ;

{===============================================================================}
{------------------------- Vos Impressions -------------------------------------}
{===============================================================================}
procedure Tof_GCContreMVTOA.BImprimerClick(Sender: TObject);
BEGIN
Imprime;
END;

procedure Tof_GCContreMVTOA.Imprime ;
BEGIN
PEtat.Tip:='E'; PEtat.Titre:=TFVierge(Ecran).Caption;
PEtat.Apercu:=True; PEtat.DeuxPages:=False; PEtat.First:=True;
PEtat.stSQL:='';

//Mettre les specificités de chaque Etat
if Mode='VTOA' then
   begin
   PEtat.Nat:='GEA';
   PEtat.Modele:='';
   End else if Mode='VTOBL' then
   begin
   PEtat.Nat:='GEA';
   PEtat.Modele:='';
   end;


EntreeOptionEdit(PEtat.Tip,PEtat.Nat,PEtat.Modele,PEtat.Apercu,PEtat.DeuxPages,PEtat.First,
                 TPageControl(GetControl('Pages')),PEtat.stSQL,PEtat.Titre, PrepareImpression);
ExecuteSQL('DELETE FROM GCTMPCONTREMVTOA WHERE GZJ_UTILISATEUR = "'+V_PGI.USer+'"');
END;

function Tof_GCContreMVTOA.PrepareImpression : integer ;
var i_ind : integer;
    TOBGZJ, TOBD, TOBL : TOB;
BEGIN
Result:=0;
ExecuteSQL('DELETE FROM GCTMPCONTREMVTOA WHERE GZJ_UTILISATEUR = "'+V_PGI.USer+'"');
TOBGZJ:=TOB.Create('',Nil,-1);
DestroyTOBGrid;
CreateTOBGrid;
for i_ind:=0 to TOBLigne.Detail.Count - 1 do
begin
   TOBL := TOBLigne.Detail[i_ind];
   TOBD := TOB.Create('',TOBGrid,-1) ;
   TOBD.Dupliquer(TOBL, False, True, False);
end;
for i_ind:=0 to TOBGrid.Detail.Count-1 do
    BEGIN
    TOBL:=TOBGrid.Detail[i_ind] ;
    TOBD:=TOB.Create('GCTMPCONTREMVTOA',TOBGZJ,-1);
    TOBD.PutValue('GZJ_UTILISATEUR',V_PGI.USer);
    TOBD.PutValue('GZJ_COMPTEUR',i_ind);
    TOBD.PutValue('GZJ_NATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG'));
    TOBD.PutValue('GZJ_SOUCHE',TOBL.GetValue('GL_SOUCHE'));
    TOBD.PutValue('GZJ_NUMERO',TOBL.GetValue('GL_NUMERO'));
    TOBD.PutValue('GZJ_INDICEG',TOBL.GetValue('GL_INDICEG'));
    TOBD.PutValue('GZJ_DATEPIECE',TOBL.GetValue('GL_DATEPIECE'));
    TOBD.PutValue('GZJ_ETABLISSEMENT',TOBL.GetValue('GL_ETABLISSEMENT'));
    TOBD.PutValue('GZJ_TIERS',TOBL.GetValue('GL_TIERS'));
    TOBD.PutValue('GZJ_ARTICLE',TOBL.GetValue('GL_ARTICLE'));
    TOBD.PutValue('GZJ_CODEARTICLE',TOBL.GetValue('GL_CODEARTICLE'));
    TOBD.PutValue('GZJ_LIBELLE',TOBL.GetValue('GL_LIBELLE'));
    TOBD.PutValue('GZJ_FOURNISSEUR',TOBL.GetValue('GL_FOURNISSEUR'));
    TOBD.PutValue('GZJ_REFCATALOGUE',TOBL.GetValue('GL_REFCATALOGUE'));
    TOBD.PutValue('GZJ_DEPOT',TOBL.GetValue('GL_DEPOT'));
    TOBD.PutValue('GZJ_ENCONTREMARQUE',TOBL.GetValue('GL_ENCONTREMARQUE'));
    TOBD.PutValue('GZJ_TENUESTOCK',TOBL.GetValue('GL_TENUESTOCK'));
    TOBD.PutValue('GZJ_TYPEREF',TOBL.GetValue('GL_TYPEREF'));
    TOBD.PutValue('GZJ_QTEACDE',TOBL.GetValue(QteATraiterName));
    END;
TOBGZJ.InsertOrUpdateDB(True);
TOBGZJ.Free;
END;

{===============================================================================}
{------------------------------ PopUp VOIRs ------------------------------------}
{===============================================================================}
procedure Tof_GCContreMVTOA.BZoomMouseEnter(Sender: TObject);
begin
if csDestroying in Ecran.ComponentState then Exit ;
PopZoom97(BZoom,POPZ) ;
end;

Procedure Tof_GCContreMVTOA.GereEnabled ( ARow : integer ) ;
var
   TOBL:TOB;
BEGIN
BZoomArticle.Enabled:=False;
BZoomCatalogue.Enabled:=False;
BZoomPiece.Enabled:=False;
BZoomClient.Enabled:=False;
BZoomFournisseur.Enabled:=False;
//
TOBL:=GetTOBLigne(GS.Row); if TOBL=nil then exit;
if IsArticle(TOBL) then BZoomArticle.Enabled := True;
BZoomCatalogue.Enabled:=True;
BZoomPiece.Enabled:=True;
BZoomClient.Enabled:=True;
BZoomFournisseur.Enabled:=True;
END ;

procedure Tof_GCContreMVTOA.BZoomArticleClick(Sender: TObject);
var TOBL,TOBA : TOB;
    RefUnique : string;
BEGIN
if not BZoomArticle.Enabled then Exit ;
TOBL := GetTOBLigne(GS.Row);
if TOBL=nil then exit;
RefUnique:=TOBL.GetValue('GL_ARTICLE');
TOBA := TOBArticles.FindFirst(['GA_ARTICLE'],[RefUnique],False) ;
if TOBA <> nil then ZoomArticle(RefUnique,TobA.GetValue('GA_TYPEARTICLE'),taConsult) ;
end;

procedure Tof_GCContreMVTOA.BZoomCatalogueClick(Sender: TObject);
BEGIN
if not BZoomCatalogue.Enabled then Exit ;
VoirCatalogue ;
end;

procedure Tof_GCContreMVTOA.BZoomPieceClick(Sender: TObject);
begin
if not BZoomPiece.Enabled then Exit ;
VoirPiece ;
end;

procedure Tof_GCContreMVTOA.BZoomClientClick(Sender: TObject);
begin
if not BZoomClient.Enabled then Exit ;
VoirClient ;
end;

procedure Tof_GCContreMVTOA.BZoomFournisseurClick(Sender: TObject);
begin
if not BZoomFournisseur.Enabled then Exit ;
VoirFournisseur ;
end;

procedure Tof_GCContreMVTOA.BOuvrirClick(Sender: TObject);
// Lance la génération des pièces
Var CM_Assist : TFCmdeContreM;
    i_ind : integer;
    TOBL:TOB;
begin
for i_ind:=0 to TOBLigne.Detail.Count - 1 do
   BEGIN
   TOBL := TOBLigne.Detail[i_ind];
   TOBL.PutValue('SELECT', '-') ;
   END;
if GS.nbSelected>0 then
    BEGIN
    for i_ind:=0 to GS.NbSelected-1 do
        BEGIN
        GS.GotoLeBOOKMARK(i_ind);
        TOBL := GetTOBLigne(GS.Row);
        if TOBL <> nil then TOBL.PutValue('SELECT', 'X');
        END;
    CM_Assist := TFCmdeContreM.Create (Application);
    Try
        CM_Assist.TOBLigne:=TOBLigne ;
        CM_Assist.TOBCatalogu:=TOBCatalogu ;
        CM_Assist.QteATraiterName:=QteATraiterName;
        CM_Assist.Mode:=Mode;
        CM_Assist.ShowModal;
        if CM_Assist.ModalResult=mrOk then Self.Update;
    Finally
        CM_Assist.free;
    End;
    END else
        PGIBox('Veuillez sélectionner les lignes à générer', '');
end ;

{===============================================================================}
{------------------------------ Divers------------------------------------------}
{===============================================================================}
function IsArticle(TOBL:TOB):Boolean;
begin
if TOBL <> Nil then
   Result := ((TOBL.GetValue('GL_TYPELIGNE')='ART') and (TOBL.GetValue('GL_TENUESTOCK')='X'))
else
   Result := False;
end;

function IsCatalogue(TOBL:TOB):Boolean;
begin
if TOBL <> Nil then
   Result := ((TOBL.GetValue('GL_TYPELIGNE')='ART') and (TOBL.GetValue('GL_TENUESTOCK')<>'X'))
else
   Result := False;
end;

Function Tof_GCContreMVTOA.GetTOBLigne(Row:Integer):TOB ;
// Retourne la TOB associée à une ligne du GRID
begin
Row:=Row-1;
if (Row >= 0) and (Row <= TOBLigne.Detail.Count - 1) then
   Result:=TOBLigne.Detail[Row]
else
   Result := nil;
end;

procedure Tof_GCContreMVTOA.VoirCatalogue;
// Visu du catalogue associé à une ligne
var TOBL : TOB;
    RefFour,RefCata : String;
BEGIN
TOBL := GetTOBLigne (GS.Row); if TOBL=nil then exit;
RefFour:=GetCodeFourDCM(TOBL);
RefCata:=TOBL.GetValue('GL_REFCATALOGUE');
AGLLanceFiche('GC','GCCATALOGU_SAISI3','',RefCata+';'+RefFour,'ACTION=CONSULTATION') ;
END;

procedure Tof_GCContreMVTOA.VoirPiece;
// Visu de la pièce d'une ligne du GRID
var TOBL : TOB;
    CleDoc : R_CleDoc;
BEGIN
TOBL := GetTOBLigne(GS.Row); if TOBL=nil then exit;
CleDoc.NaturePiece:=Uppercase(TOBL.GetValue('GL_NATUREPIECEG'));
CleDoc.Souche:=Uppercase(TOBL.GetValue('GL_SOUCHE'));
CleDoc.DatePiece:=TOBL.GetValue('GL_DATEPIECE');
CleDoc.NumeroPiece:=TOBL.GetValue('GL_NUMERO');
CleDoc.Indice:=TOBL.GetValue('GL_INDICEG');
SaisiePiece(CleDoc,taConsult)
END;

procedure Tof_GCContreMVTOA.VoirClient;
var TOBL : TOB;
    Client : String;
BEGIN
TOBL := GetTOBLigne(GS.Row); if TOBL=nil then exit;
Client:=TOBL.GetValue('GL_TIERS');
AGLLanceFiche('GC','GCTIERS','', Client,'ACTION=CONSULTATION');
END;

procedure Tof_GCContreMVTOA.VoirFournisseur;
var TOBL : TOB;
    Four : String;
BEGIN
TOBL := GetTOBLigne(GS.Row); if TOBL=nil then exit;
Four:=TOBL.GetValue('GL_FOURNISSEUR');
AGLLanceFiche('GC','GCTIERS','',Four,'ACTION=CONSULTATION') ;
END;

{procedure Tof_GCContreMVTOA.AffichePied (ACol, ARow : integer);
var TOBL,TOBArt : TOB;
    RefArt,St : string;
    Dim : Array of string;
    i_ind: integer;
    GrilleDim,CodeDim,LibDim : String ;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then Exit;
TDimension.Visible:=False;
if IsArticle(TOBL) then
   BEGIN
   RefArt:=TOBL.GetValue('GL_ARTICLE');
   TOBArt:=TOBArticles.FindFirst(['GA_ARTICLE'],[RefArt],False);
   if (TOBArt <> nil) and (TOBArt.GetValue('GA_STATUTART')='DIM') then
      BEGIN
      SetLength(dim,MaxDimension);
      St:='';
      for i_ind:=1 to MaxDimension do
         BEGIN
         GrilleDim:=TOBArt.GetValue ('GA_GRILLEDIM'+IntToStr(i_ind)) ;
         CodeDim:=TOBArt.GetValue ('GA_CODEDIM'+IntToStr(i_ind)) ;
         LibDim:=GCGetCodeDim(GrilleDim,CodeDim,i_ind) ;
         if LibDim<>'' then Dim[i_ind - 1]:=LibDim else Dim[i_ind - 1]:='' ;
         if Dim[i_ind - 1]<>'' then
            if St='' then St:=Dim[i_ind -1]
                     else St:=St + ' - ' + Dim[i_ind - 1];
         END;
      TDimension.Visible:=True;
      TDimension.Caption:=St;
      END;
   END;
end; }

Procedure Tof_GCContreMVTOA.BRechercherClick(Sender: TObject);
BEGIN
FindDebut:=True ; FindLigne.Execute ;
END;

procedure Tof_GCContreMVTOA.FindLigneFind(Sender: TObject);
BEGIN
Rechercher(GS, FindLigne, FindDebut) ;
END;

procedure Tof_GCContreMVTOA.CumulQteATraiter;
var i:Integer;
    Qte:Double;
    TOBL:TOB;
BEGIN
if TOTQTE_A_CDE <> nil then
   BEGIN
   Qte:=0;
   for i:=0 to TobLigne.Detail.count -1  do
      BEGIN
      TOBL := TobLigne.Detail[i];
      if TOBL <> nil then Qte := Qte + TOBL.GetValue(QteATraiterName);
      END;
   TOTQTE_A_CDE.Value := Qte;
   END;
END;

(*Function Tof_GCContreMVTOA.ColListeOk(iCol:Integer):Boolean;
BEGIN
{$IFDEF EAGLCLIENT}
    Result := iCol <= TFMul(Ecran).Fliste.RowCount - 1;
{$ELSE}
    Result := iCol <= TFMul(Ecran).Fliste.Columns.Count - 1;
{$ENDIF}
END; *)

{===============================================================================}
{------------------------------ Gestion des TOBs -------------------------------}
{===============================================================================}
procedure Tof_GCContreMVTOA.InitGrille ;
var Col,Dec, FixedWidth, Larg,QteLarg : integer ;
    tal,QteTa : TAlignment ;
    STitre,St,Ch,stA,FF,Typ,FPerso,QteFF : string ;
    NomList,FRecordSource,FLien,FSortBy,FFieldList,FTitre,FLargeur,FAlignement,FParams,tt,NC : string ;
    Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul,OkTri,OkNumCol : boolean ;
begin
FixedWidth:=10;
GS.ColCount:=60;
Col:=1; STitre:=' '; ColName:='Fixed';
QteTa:=taRightJustify ;
QteFF:='#,##0.000;; ;';
QteLarg:=100;

With TFMul(Ecran) do
    BEGIN
    NomList:=Q.Liste;
    ChargeHListe(NomList,FRecordSource,FLien,FSortBy,FFieldList,FTitre,FLargeur,FAlignement,FParams,tt,NC,FPerso,OkTri,OkNumCol);
    GS.Titres.Add(FFieldList) ;
    //FSortList:=FSortBy;
    FSortList:=FSortByToTobSort(FSortBy);
    While Ftitre<> '' do
        BEGIN
        StA:=ReadTokenSt(FAlignement);
        St:=ReadTokenSt(Ftitre);
        Ch:=ReadTokenSt(FFieldList);
        Larg:=ReadTokenI(FLargeur);
        tal:=TransAlign(StA,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
        if OkVisu then
            BEGIN
            GS.Cells[Col,0]:=St ;
            GS.ColAligns[Col]:=tal;
            GS.ColWidths[Col]:=Larg*GS.Canvas.TextWidth('W') ;;
            if OkLib then GS.ColFormats[Col]:='CB=' + Get_Join(Ch);
            Typ:=ChampToType(Ch) ;
            if (Typ='INTEGER') or (Typ='SMALLINT') or (Typ='DOUBLE') then GS.ColFormats[Col]:=FF ;
            STitre:=STitre+';'+St;
            ColName:=ColName+';'+Ch;
            if Ch='GL_QTESTOCK' then
                BEGIN
                QteTa:=tal;
                QteLarg:=GS.ColWidths[Col];
                QteFF:=FF;
                END;
            inc (Col);
            END;
        END;
    GS.ColAligns[Col]:=QteTa ;
    GS.ColWidths[Col] := QteLarg;
    GS.ColFormats[Col] := QteFF;
    GS.Cells[Col, 0] := QteATraiterTitle;
    ColName:=ColName+';'+QteATraiterName;
    inc (Col);
    GS.ColCount:=Col ;
    AffecteGrid(GS,taConsult) ;
    Hmtrad.ResizeGridColumns(GS) ;
    GS.ColWidths[0]:=FixedWidth;
    Hmtrad.ResizeGridColumns(GS) ;
    END ;
end;

procedure Tof_GCContreMVTOA.CreateTOBLigne ;
begin
TOBLigne:=TOB.Create('',Nil,-1) ;
end;

procedure Tof_GCContreMVTOA.CreateTOBGrid ;
begin
TOBGrid:=TOB.Create('',Nil,-1) ;
end;

procedure Tof_GCContreMVTOA.CreateTOB ;
begin
CreateTOBLigne ;
TOBArticles:=TOB.Create('',Nil,-1) ;
TOBCatalogu:=TOB.Create('',Nil,-1) ;
TOBDispo:=TOB.Create('',Nil,-1) ;
TOBDispoCM:=TOB.Create('',Nil,-1) ;
CreateTOBGrid ;
end;

procedure Tof_GCContreMVTOA.DestroyTOBLigne ;
begin
if TOBLigne <> Nil then begin TOBLigne.Free ; TOBLigne := Nil; end;
end;

procedure Tof_GCContreMVTOA.DestroyTOBGrid ;
begin
if TOBGrid <> Nil then TOBGrid.Free ; TOBGrid := Nil;
end;

procedure Tof_GCContreMVTOA.DestroyTOB ;
begin
DestroyTOBLigne ;
TOBArticles.Free ; TOBArticles := Nil;
TOBCatalogu.Free ; TOBCatalogu := Nil;
TOBDispo.Free ; TOBDispo := Nil;
TOBDispoCM.Free ; TOBDispoCM := Nil;
DestroyTOBGrid ;
end;

Procedure Tof_GCContreMVTOA.TOBLigneAddChampSupp;
var i:Integer;
    TOBL:TOB;
begin
for i:=0 to TOBLigne.Detail.Count - 1 do
   BEGIN
   TOBL := TOBLigne.Detail[i];
   TOBL.AddChampSup(QteATraiterName, False) ;
   TOBL.PutValue(QteATraiterName, 0) ;
   TOBL.AddChampSup('SELECT', False) ;
   TOBL.PutValue('SELECT', '-') ;
   END;
end;

Procedure Tof_GCContreMVTOA.TOBLigneConvQteStock;
var i:Integer;
    TOBL:TOB;
    RatioStk:double;
begin
for i:=0 to TOBLigne.Detail.Count - 1 do
   BEGIN
   TOBL := TOBLigne.Detail[i];
   RatioStk:=GetRatio(TOBL,Nil,trsStock);
   TOBL.putValue('GL_QTESTOCK',TOBL.getValue('GL_QTESTOCK')/RatioStk);
   TOBL.putValue('GL_QTERESTE', TOBL.getValue('GL_QTESTOCK'));  { NEWPIECE }
   TOBL.putValue('GL_QTEFACT',TOBL.getValue('GL_QTEFACT')/RatioStk);
   END;
end;

Procedure Tof_GCContreMVTOA.ChargeTOBArticles;
// Charge la TOBArticles pour les articles des lignes
var i:Integer;
    WhereArt,RefUnique:String;
    TOBL,TOBArt,TOBA:TOB;
    QArt:TQuery;
    QArtCol:String;
begin
TOBArt:=TOB.Create('',nil,-1);
for i:=0 to TOBLigne.Detail.Count-1 do
   BEGIN
   TOBL:=TOBLigne.Detail[i] ;
   if IsArticle(TOBL) then
      BEGIN
      RefUnique:=TOBL.GetValue('GL_ARTICLE');
      if TOBArt.FindFirst(['REFUNIQUE'],[RefUnique],False) = Nil then
         BEGIN
         TOBA:=TOB.Create('',TOBArt,-1);
         TOBA.AddChampSup('REFUNIQUE',False);
         TOBA.PutValue('REFUNIQUE', RefUnique);
         END;
      END;
   END;
WhereArt:='';
for i:=0 to TOBArt.Detail.Count-1 do
   BEGIN
   RefUnique:=TOBArt.Detail[i].GetValue('REFUNIQUE');
   if WhereArt <> '' then WhereArt:=WhereArt +' OR ';
   WhereArt:=WhereArt+'(GA_ARTICLE="'+RefUnique+'")';
   END;
TOBArt.Free;
if WhereArt<>'' then
   BEGIN
   QArtCol:='*';
   QArt:=OpenSQL('SELECT '+QArtCol+' FROM ARTICLE WHERE '+WhereArt,True) ;
   if not QArt.Eof then TOBArticles.LoadDetailDB('ARTICLE','','',QArt,False,True) ;
   Ferme(QArt) ;
   END;
end;

Procedure Tof_GCContreMVTOA.ChargeTOBCatalogu;
// Charge la TOBCatalogu pour les catalogues des lignes
Var i:integer ;
    TOBL, TOBCata, TOBCat:TOB ;
    RefCata,RefFour:String ;
    WhereCat:String;
    QCat:TQuery;
    QCatCol:String;
Begin
TOBCata:=TOB.Create('',nil,-1);
for i:=0 to TOBLigne.Detail.Count-1 do
   BEGIN
   TOBL:=TOBLigne.Detail[i] ;
   if IsCatalogue(TOBL) then
      BEGIN
      RefCata:=TOBL.GetValue('GL_REFCATALOGUE');
      RefFour:=GetCodeFourDCM(TOBL);
      if TOBCata.FindFirst(['REFCATA','REFFOUR'],[RefCata,RefFour],False) = Nil then
         BEGIN
         TOBCat:=TOB.Create('',TOBCata,-1);
         TOBCat.AddChampSup('REFCATA',False);
         TOBCat.PutValue('REFCATA', RefCata);
         TOBCat.AddChampSup('REFFOUR',False);
         TOBCat.PutValue('REFFOUR', RefFour);
         END;
      END;
   END;
WhereCat:='';
for i:=0 to TOBCata.Detail.Count-1 do
   BEGIN
   RefCata:=TOBCata.Detail[i].GetValue('REFCATA');
   RefFour:=TOBCata.Detail[i].GetValue('REFFOUR');
   if WhereCat <> '' then WhereCat:=WhereCat+' OR ';
   WhereCat:=WhereCat + '(GCA_REFERENCE="'+RefCata+'" AND GCA_TIERS="'+RefFour+'")';
   END;
TOBCata.Free;
if WhereCat<>'' then
   BEGIN
   QCatCol:='*';
   QCat:=OpenSQL('SELECT '+QCatCol+' FROM CATALOGU WHERE '+WhereCat,True) ;
   if not QCat.Eof then TOBCatalogu.LoadDetailDB('CATALOGU','','',QCat,False,True) ;
   Ferme(QCat) ;
   END ;
end;

procedure Tof_GCContreMVTOA.ChargeTOBDispo;
// Charge la TOBDispo avec les articles,depôts des lignes
var i:Integer ;
    TOBL,TOBDisp,TOBD:TOB ;
    RefUnique,RefDepot:String;
    WhereDispo:String;
    QDispo:TQuery;
    QDispoCol:String;
begin
TOBDisp:=TOB.Create('',nil,-1);
for i:=0 to TOBLigne.Detail.Count - 1 do
   BEGIN
   TOBL:=TOBLigne.Detail[i] ;
   if IsArticle(TOBL) then
      BEGIN
      RefUnique:=TOBL.GetValue('GL_ARTICLE');
      RefDepot:=TOBL.GetValue('GL_DEPOT');
      if TOBDisp.FindFirst(['REFUNIQUE','REFDEPOT'],[RefUnique,RefDepot],False) = Nil then
         BEGIN
         TOBD:=TOB.Create('',TOBDisp,-1);
         TOBD.AddChampSup('REFUNIQUE',False);
         TOBD.PutValue('REFUNIQUE', RefUnique);
         TOBD.AddChampSup('REFDEPOT',False);
         TOBD.PutValue('REFDEPOT', RefDepot);
         END;
      END ;
   END ;
WhereDispo:='';
for i:=0 to TOBDisp.Detail.Count-1 do
   BEGIN
   RefUnique:=TOBDisp.Detail[i].GetValue('REFUNIQUE') ;
   RefDepot:=TOBDisp.Detail[i].GetValue('REFDEPOT') ;
   if WhereDispo <> '' then WhereDispo:=WhereDispo+' OR ';
   WhereDispo:=WhereDispo + '(GQ_ARTICLE="'+RefUnique+'" AND GQ_DEPOT="'+RefDepot+'")';
   END;
TOBDisp.Free;
if WhereDispo <> '' then
   BEGIN
   WhereDispo:=WhereDispo+' AND GQ_CLOTURE="-"';
   QDispoCol:='*';
   QDispo:=OpenSQL('SELECT '+QDispoCol+' FROM DISPO WHERE '+WhereDispo,True) ;
   if not QDispo.Eof then TOBDispo.LoadDetailDB('DISPO','','',QDispo,False,True) ;
   Ferme(QDispo) ;
   END;
end;

procedure Tof_GCContreMVTOA.ChargeTOBDispoCM;
// Charge la TOBDispoCM avec les Depôts,Clients,Article et fournisseurs des lignes
var i:Integer ;
    TOBL,TOBDispCM,TOBDCM:TOB ;
    Cata,Tiers,Depot,Four:String;
    WhereDispoCM:String;
    QDispoCM:TQuery;
    QDispoCMCol:String;
begin
// Forme le SQL qui va bien
TOBDispCM:=TOB.Create('',nil,-1);
for i:=0 to TOBLigne.Detail.Count - 1 do
   BEGIN
   TOBL:=TOBLigne.Detail[i] ;
   if IsCatalogue(TOBL) then
      BEGIN
      Cata:=TOBL.GetValue('GL_REFCATALOGUE') ;
      Tiers:=TOBL.GetValue('GL_TIERS') ;
      Depot:=TOBL.GetValue('GL_DEPOT');
      Four:=TOBL.GetValue('GL_FOURNISSEUR');
      if TOBDispCM.FindFirst(['CATA','TIERS','DEPOT','FOUR'],[Cata,Tiers,Depot,Four],False) = Nil then
         BEGIN
         TOBDCM:=TOB.Create('',TOBDispCM,-1);
         TOBDCM.AddChampSup('CATA',False);
         TOBDCM.PutValue('CATA', Cata);
         TOBDCM.AddChampSup('TIERS',False);
         TOBDCM.PutValue('TIERS', Tiers);
         TOBDCM.AddChampSup('DEPOT',False);
         TOBDCM.PutValue('DEPOT', Depot);
         TOBDCM.AddChampSup('FOUR',False);
         TOBDCM.PutValue('FOUR', Four);
         END;
      END ;
   END ;
WhereDispoCM:='';
for i:=0 to TOBDispCM.Detail.Count-1 do
   BEGIN
   Cata:=TOBDispCM.Detail[i].GetValue('CATA') ;
   Tiers:=TOBDispCM.Detail[i].GetValue('TIERS') ;
   Depot:=TOBDispCM.Detail[i].GetValue('DEPOT') ;
   Four:=TOBDispCM.Detail[i].GetValue('FOUR') ;
   if WhereDispoCM <> '' then WhereDispoCM:=WhereDispoCM+' OR ';
   WhereDispoCM:=WhereDispoCM + '(GQC_DEPOT="'+Depot+'" AND GQC_CLIENT="'+Tiers+'" AND GQC_REFERENCE="'+Cata+'" AND GQC_FOURNISSEUR="'+Four+'")';
   END;
TOBDispCM.Free;
// Charge la TOB
if WhereDispoCM <> '' then
   BEGIN
   QDispoCMCol:='*';
   QDispoCM:=OpenSQL('SELECT '+QDispoCMCol+' FROM DISPOCONTREM WHERE '+WhereDispoCM,True) ;
   if not QDispoCM.Eof then TOBDispoCM.LoadDetailDB('DISPOCONTREM','','',QDispoCM,False,True) ;
   Ferme(QDispoCM) ;
   END;
end;

{===============================================================================}
{------------------------------ Gestion du grid --------------------------------}
{===============================================================================}
procedure Tof_GCContreMVTOA.DessineCellSEL (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
    ARect: Trect ;
BEGIN
if Arow < GS.FixedRows then exit ;
if (gdFixed in AState) and (ACol = 0) then
   begin
   ARect:=GS.CellRect(Acol,Arow) ;
   Canvas.Brush.Color := GS.FixedColor;
   Canvas.FillRect(ARect);
   if (ARow = GS.Row) then
      BEGIN
      Canvas.Brush.Color := clBlack ;
      Canvas.Pen.Color := clBlack ;
      Triangle[1].X:=((ARect.Left+ARect.Right) div 2) ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
      Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
      Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
      if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
      END;
   end;
END;

procedure Tof_GCContreMVTOA.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
BEGIN
GereEnabled(Ou) ;
//AffichePied(GS.Col,GS.Row);
GS.InvalidateRow(ou) ;
END;

procedure Tof_GCContreMVTOA.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
BEGIN
GS.InvalidateRow(ou) ;
END;

procedure Tof_GCContreMVTOA.GSDblClick(Sender: TObject);
// Double Click sur la grille
BEGIN
VoirPiece;
END;


Initialization
  registerclasses ( [ Tof_GCContreMVTOA ] ) ;
end.
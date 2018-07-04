unit UTofMouvStkEx;

interface

uses  StdCtrls,Controls,Classes,sysutils,forms,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      db,FE_main,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
      ComCtrls,Hpanel, Math,HCtrls,HEnt1,HMsgBox,UTOF,Vierge,UTOB,AglInit,LookUp,EntGC,SaisUtil,
      graphics,grids,windows,M3FP,HTB97,Dialogs, AGLInitGC, ExtCtrls, Hqry,LicUtil,HDimension,
{$IFNDEF CCS3}
      SaisieSerie_TOF,
{$ENDIF}
      Facture,FactUtil,Menus,Messages,ENT1,FactComm,FactCalc,FactTOB, FactArticle, FactPiece,
      FactLotSerie,
      UTofOptionEdit,StockUtil,LigDispoLot,UtilGrp,uEntCommun,UtilTOBPiece;

function GCLanceFiche_MouvStkEx(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
    TOF_MouvStkEx = Class (TOF)
    private
        GS : THGRID ;
        INFOSLIGNE : THGRID ;
        PENTETE, PPIED : THPanel;
        GP_NUMEROPIECE : THPanel;
        GP_DATEPIECE : THCritMaskEdit;
        GP_ETABLISSEMENT : THValComboBox;
        GP_DEPOT : THValComboBox;
        TGP_DEPOT : THLabel;
        GP_DEPOTDEST : THValComboBox;
        TGP_DEPOTDEST : THLabel;
        GP_DOMAINE : THValComboBox;
        TGP_DOMAINE : THLabel;
        BZOOM,BImprimer : TToolbarButton97;
        BZOOMARTICLE,BZOOMDEPOT : TToolbarButton97;
        BZOOMLOT,BZOOMSERIE,BZOOMSUIVANTE,BZOOMPRECEDENTE : TToolbarButton97;
        POPZ: TPopupMenu;
        PPInfosLigne : TStrings ;
        ForcerFerme,GeneCharge,GereLot,GereSerie,
        GereAcompte,CommentLigne,ForceRupt,EstAvoir,DejaRentre,ValideEnCours,
        OuvreAutoPort,PasBouclerCreat  : boolean ;
        StCellCur,LesColonnes,GereEche,CommentEnt,CommentPied,CalcRupt,
        VenteAchat,DimSaisie : String ;
        Nature   : string;
        CleDoc : R_CleDoc ;
        IdentCols : Array[0..19] of R_IdentCol ;
        NewLigneDim: Boolean ;//AC, nouvelle ligne d'un art GEN existant
        ANCIEN_TOBDimDetailCount : integer; // Sauvegarde du nb initial de dimensions
        // Objets mémoire
        TOBPiece,TOBArticles,TOBTarif,TOBConds : TOB ;
        TOBCXV, TOBDim,TOBDesLots,TOBSerie: TOB ;
        TobLigneTarif: Tob; 
        // Initialisation
        procedure ChargeFromNature ;
        procedure ToutAllouer ;
        procedure ToutLiberer ;
        procedure InitEnteteDefaut ( Totale : boolean ) ;
        procedure InitPieceCreation ;
        procedure GotoEntete ;
        procedure ReInitPiece ;
        procedure InitPasModif ;
        procedure InitToutModif ;
        procedure FlagDomaine ;
        procedure FlagDepot ;
        procedure ChargeLaPiece;
        //Actions liées au Grid
        procedure EtudieColsListe ;
        procedure FormateZoneSaisie (ACol,ARow : Longint ) ;
        Function  FormateZoneDivers ( St : String ; ACol : Longint ) : String ;
        //Evenement du grid
        procedure GSEnter(Sender: TObject);
        procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GSCellExit(Sender: TObject; var ACol,ARow : Integer; var Cancel : Boolean);
        procedure GSElipsisClick(Sender: TObject);
        Function  GereElipsis ( LaCol : integer ) : boolean ;
        procedure GSDblClick(Sender: TObject);
        //Manipulation des LIGNES
        Function  InitLaLigne ( ARow : integer ; NewQte : double ) : T_ActionTarifArt ;
        Procedure ArtVersLigne ( TOBPiece,TOBA,TOBL : TOB ) ;
        Function  PieceModifiee : boolean ;
        // Affichages
        Procedure ZoomOuChoixArt ( ACol,ARow : integer ) ;
        procedure AfficheLaLigne ( ARow : integer ) ;
        Procedure ShowDetail ( ARow : integer ) ;
        Procedure GereEnabled ( ARow : integer ) ;
        // Actions sur l'Entête
        procedure GP_DATEPIECEExit(Sender: TObject);
        Function  ExitDatePiece : Boolean ;
        procedure GP_DEPOTChange(Sender: TObject);
        procedure GP_DEPOTDESTChange(Sender: TObject);
        procedure GP_ETABLISSEMENTChange(Sender: TObject);
        //Manipulation des articles
        Procedure CodesArtToCodesLigne ( TOBArt : TOB ; ARow : integer ) ;
        Function  IdentifierArticle ( Var ACol,ARow : integer ; Var Cancel : boolean ; Click,FromMacro : Boolean ) : boolean ;
        Procedure ChargeTOBDispo ( ARow : integer ) ;
        Procedure UpdateArtLigne ( ARow : integer ; FromMacro,Calc : boolean ; NewQte : double ) ;
        Procedure TraiteArticle ( Var ACol,ARow : integer ; Var Cancel : boolean ; FromMacro,Calc : boolean ; NewQte : double ) ;
        // ARTICLES
        Procedure TraiteMotif ( ARow : integer ) ;
        Procedure TraiteLibelle ( ARow : integer ) ;
        Procedure GereLesLots ( ARow : integer ) ;
        Procedure GereLesSeries ( ARow : integer ) ;
        // DIMENSIONS
        Procedure RemplirTOBDim ( CodeArticle : String; Ligne: Integer ) ;
        Procedure AppelleDim ( ARow : integer ) ;
        Procedure TraiteLesDim ( var ARow : integer ; NewArt : boolean ) ;
        Procedure AffichageDim ;
        //Manipulation des Qtés
        Function  TraiteRupture ( ARow : integer ) : boolean ;
        Function  TraiteQte ( ACol,ARow : integer ) : Boolean ;
        //Manipulation des dépots
        Procedure AffecteTRE;
        //Actions, boutons
        procedure BZoomArticleClick(Sender: TObject);
        procedure BZoomDepotClick(Sender: TObject);
        procedure BZoomLotClick(Sender: TObject);
        procedure BZoomSerieClick(Sender: TObject);
        procedure BZoomSuivanteClick(Sender: TObject);
        procedure BZoomPrecedenteClick(Sender: TObject);
        procedure BZoomMouseEnter(Sender: TObject);
        procedure BImprimerClick(Sender: TObject);
        Procedure ClickDel ( ARow : integer ; AvecC,FromUser : boolean; SupDim: boolean=False ) ;
        Procedure ClickInsert ( ARow : integer ) ;
        // Validations
        procedure ValideLaNumerotation ;
        procedure ValideLesLots ;
        procedure ValideLesSeries ;
        procedure ValideLaPiece ;
        procedure ValideImpression ;
        Function  SortDeLaLigne : boolean ;
        procedure ClickValide ;
    public
        Action   : TActionFiche ;
        // Evenement de la Form
        procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
        // Evenement de la TOF
        procedure OnArgument (Arguments : String ) ; override;
        procedure OnUpdate ; override;
        procedure OnClose  ; override ;

    END ;

var DEV      : RDEVISE ;

const colRang=1 ;
      NbRowsInit = 50 ;
      NbRowsPlus = 20 ;

const
// libellés des messages
TexteMessage: array[1..11] of string 	= (
          {1}  '(référence de substitution possible :'
          {2} ,'Le stock disponible est insuffisant.'
          {3} ,'0;?caption?;Confirmez-vous l''abandon de la saisie ?;Q;YN;Y;N;'
          {4} ,'ATTENTION : Pièce non enregistrée !'
          {5} ,'ATTENTION : Cette pièce en cours de traitement par un autre utilisateur n''a pas été enregistrée !'
          {6} ,'ATTENTION : Cette pièce ne peut pas passer en comptabilité et n''a pas été enregistrée !'
          {7} ,'ATTENTION : La pièce présente un problème de numérotation et n''a pas été enregistrée !'
          {8} ,'ATTENTION. Le stock disponible est insuffisant pour certains articles.'
          {9} ,'ATTENTION : l''impression ne s''est pas correctement effectuée !'
         {10} ,'Non affecté'
         {11} ,''
              );
ErrorMessage: array[1..20] of string 	= (
          {1}  '0;?caption?;Saisie impossible. Cet article fermé n''est pas autorisé pour cette nature de document;E;O;O;O;'
          {2} ,'0;?caption?;Cet article est en rupture, confirmez-vous malgré tout la quantité ?;Q;YN;Y;N;'
          {3} ,'0;?caption?;Cet article est en rupture ;E;O;O;O;'
          {4} ,'0;?caption?;La date que vous avez renseignée n''est pas valide;E;O;O;O;'
          {5} ,'0;?caption?;La date que vous avez renseignée n''est pas dans un exercice ouvert;E;O;O;O;'
          {6} ,'0;?caption?;La date que vous avez renseignée est antérieure à une clôture;E;O;O;O;'
          {7} ,'0;?caption?;La date que vous avez renseignée est antérieure à une clôture;E;O;O;O;'
          {8} ,'0;?caption?;La date que vous avez renseignée est en dehors des limites autorisées;E;O;O;O;'
          {9} ,'0;?caption?;Vous ne pouvez pas saisir avant le ;E;O;O;O;'
         {10} ,'0;?caption?;La date est antérieure à celle de dernière clôture de stock;E;O;O;O;'
         {11} ,'0;?caption?;Vous ne pouvez pas enregistrer une pièce vide;E;O;O;O;'
         {12} ,'0;?caption?;Vous ne pouvez pas enregistrer une pièce sans articles;E;O;O;O;'
         {13} ,'0;?caption?;Vous ne pouvez pas enregistrer une pièce sans mouvements;E;O;O;O;'
         {14} ,'0;?caption?;Vous ne pouvez pas enregistrer une pièce à cette date;E;O;O;O;'
         {15} ,'0;?caption?;Enregistrement impossible : l''acompte est supérieur au total de la pièce;E;O;O;O;'
         {16} ,'0;?caption?;Enregistrement impossible : la devise est incorrecte;E;O;O;O;'
         {17} ,'0;?caption?;Les quantités négatives ne sont pas autorisées ;E;O;O;O;'
         {18} ,'0;?caption?;Voulez-vous affecter ce dépôt sur toutes les lignes concernées ?;Q;YN;Y;N;'
         {19} ,'0;?caption?;Les dépôts émetteur et destinataire doivent être renseignés;E;O;O;O;'
         {20} ,'0;?caption?;Le dépôt doit être renseigné;E;O;O;O;' // JT eQualité 10823
              );




implementation

function GCLanceFiche_MouvStkEx(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result:='';
if Nat='' then exit;
if Cod='' then exit;
Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

{==============================================================================================}
{================================= Evenement de la TOF ========================================}
{==============================================================================================}
Procedure TOF_MouvStkEx.OnArgument (Arguments : String ) ;
var St,StA,StM,StC : string;
    i_pos : integer ;
    CC : THValComboBox;
begin
inherited ;
St:=Arguments ; StA:=St;
i_pos:=Pos('ACTION=',St) ;
St:=Copy(St,i_pos,Length(St));
if i_pos>0 then
   BEGIN
   System.Delete(St,1,7) ;
   StM:=uppercase(ReadTokenSt(St)) ;
   if StM='CREATION' then BEGIN Action:=taCreat ; END ;
   if StM='MODIFICATION' then BEGIN Action:=taModif ; END ;
   if StM='CONSULTATION' then BEGIN Action:=taConsult ; END ;
   END ;
StC:=St;
LookLesDocks(Ecran) ;
InitLesCols ;
FillChar(IdentCols,Sizeof(IdentCols),#0) ;
if Action=taCreat then
   begin
   if PasCreerDateGC(V_PGI.DateEntree) then Exit ;
   Nature:=StC;
   FillChar(CleDoc,Sizeof(CleDoc),#0) ;
   CleDoc.NaturePiece:=Nature ; CleDoc.DatePiece:=V_PGI.DateEntree ;
   CleDoc.Souche:='' ; CleDoc.NumeroPiece:=0 ; CleDoc.Indice:=0 ;
   VH_GC.GCLastRefPiece:='' ;
   end else
if Action = taConsult then
   begin
   stringToCleDoc(stA,CleDoc);
   Nature:=CleDoc.NaturePiece;
   end;
if Nature='TEM' then
   begin
   TFVierge(Ecran).Caption:='Mouvements inter-dépôt';
   UpdateCaption(Ecran);
   Ecran.HelpContext:=110000313 ;
   end else
   if (Nature='EEX') or (Nature='SEX') then Ecran.HelpContext:=110000315 else
      if (Nature='RCC') or (Nature='RCF') or (Nature='RPR')then Ecran.HelpContext:=110000319 ;

PPInfosLigne:=TStringList.Create ;
// gestion etablissement par user
CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ; if CC<>Nil then PositionneEtabUser(CC) ;
//
PENTETE:=THPanel(GetControl('PENTETE'));
PPIED:=THPanel(GetControl('PPIED'));
GP_NUMEROPIECE:=THPanel(GetControl('GP_NUMEROPIECE'));
GP_DATEPIECE:=THCritMaskEdit(GetControl('GP_DATEPIECE'));
GP_DATEPIECE.OnExit:=GP_DATEPIECEExit;
GP_ETABLISSEMENT:=THValComboBox(GetControl('GP_ETABLISSEMENT'));
GP_ETABLISSEMENT.OnChange:=GP_ETABLISSEMENTChange ;
TGP_DEPOT:=THLabel(GetControl('TGP_DEPOT'));
GP_DEPOT:=THValComboBox(GetControl('GP_DEPOT'));
GP_DEPOT.OnChange:=GP_DEPOTChange ;
TGP_DEPOTDEST:=THLabel(GetControl('TGP_DEPOTDEST'));
GP_DEPOTDEST:=THValComboBox(GetControl('GP_DEPOTDEST'));
GP_DEPOTDEST.OnChange:=GP_DEPOTDESTChange ;
if Nature = 'TEM' then
    begin
    TGP_DEPOTDEST.Visible := True;
    GP_DEPOTDEST.Visible := True;
    end
    else
    begin
    TGP_DEPOTDEST.Visible := False;
    GP_DEPOTDEST.Visible := False;
    end;
TGP_DOMAINE:=THLabel(GetControl('TGP_DOMAINE'));
GP_DOMAINE:=THValComboBox(GetControl('GP_DOMAINE'));
//GP_DOMAINE.OnChange:=GP_DOMAINEChange ;

POPZ:=TPopupMenu(GetControl('POPZ'));
BZOOM:=TToolbarButton97(GetControl('BZOOM'));
BZOOM.OnMouseEnter:=BZoomMouseEnter;
BZOOMARTICLE:=TToolbarButton97(GetControl('BZOOMARTICLE'));
BZOOMARTICLE.OnClick:=BZoomArticleClick;
BZOOMDEPOT:=TToolbarButton97(GetControl('BZOOMDEPOT'));
BZOOMDEPOT.OnClick:=BZoomDepotClick;
BZOOMLOT:=TToolbarButton97(GetControl('BZOOMLOT'));
BZOOMLOT.OnClick:=BZoomLotClick;
BZOOMSERIE:=TToolbarButton97(GetControl('BZOOMSERIE'));
BZOOMSERIE.OnClick:=BZoomSerieClick;
BZOOMSUIVANTE:=TToolbarButton97(GetControl('BZOOMSUIVANTE'));
BZOOMSUIVANTE.OnClick:=BZoomSuivanteClick;
BZOOMPRECEDENTE:=TToolbarButton97(GetControl('BZOOMPRECEDENTE'));
BZOOMPRECEDENTE.OnClick:=BZoomPrecedenteClick;
BImprimer := TToolbarButton97(GetControl('BIMPRIMER'));
BImprimer.OnClick:=BImprimerClick;

GS:=THGRID(GetControl('GS'));
GS.OnRowEnter:=GSRowEnter ;
GS.OnRowExit:=GSRowExit ;
GS.OnCellEnter:=GSCellEnter ;
GS.OnCellExit:=GSCellExit ;
GS.OnElipsisClick:=GSElipsisClick ;
GS.OnDblClick:=GSDblClick ;
GS.OnEnter:=GSEnter;
INFOSLIGNE:=THGRID(GetControl('INFOSLIGNE'));
ChargeFromNature ; EtudieColsListe;
ToutAllouer ; DejaRentre:=False ; ValideEnCours:=False ;
PasBouclerCreat:=False ; if CleDoc.NoPersp>0 then PasBouclerCreat:=True ;
FlagDepot ; FlagDomaine ;
Case Action of
   taCreat   : BEGIN
               InitEnteteDefaut(True) ;
               InitPieceCreation ;
               END ;
   taConsult : BEGIN
               ChargeLaPiece;
               TobPiece.PutEcran(Ecran);
               TobPiece.PutGridDetail(GS,False,False,LesColonnes) ;
               GP_NUMEROPIECE.Caption := TobPiece.GetValue('GP_NUMERO');
               BImprimer.Visible := True;
               END;
   END ;
TFVierge(Ecran).OnKeyDown:=FormKeyDown ;
INFOSLIGNE.AutoResizeColumn(0,108);
INFOSLIGNE.AutoResizeColumn(1,108);
INFOSLIGNE.Width := 230;
//
TFVierge(Ecran).HMTrad.ResizeGridColumns(GS) ;
AffecteGrid(GS,Action) ;
InitPasModif ;
GereEnabled(1) ;
end;


procedure TOF_MouvStkEx.OnUpdate;
BEGIN
inherited;
ClickValide;
END;

procedure TOF_MouvStkEx.OnClose;
BEGIN
if Action<>taConsult then
   begin
   if ValideEnCours then
      begin
      AfficheError:=False;
      LastError:=1;
      end ;
   if ((Not ForcerFerme) and (PieceModifiee) and (DejaRentre) and (LastError=0)) then
      if HShowMessage(TexteMessage[3],TFVierge(Ecran).Caption,'')<>mrYes then
         begin
         AfficheError:=False;
         LastError:=1;
         end;
   end;
inherited;
if LastError<>0 then AfficheError:=true else
   begin
   ToutLiberer;
   PPInfosLigne.Clear ; PPInfosLigne.Free ; PPInfosLigne:=Nil ;
   end;
END;

procedure TOF_MouvStkEx.FormKeyDown(Sender: TObject; var Key: Word;  Shift: TShiftState);
Var OkG,Vide : Boolean ;
begin
OkG:=(Screen.ActiveControl=GS) ; Vide:=(Shift=[]) ;
Case Key of
   VK_RETURN : if (ctxMode in V_PGI.PGIContexte) then SendMessage(GS.Handle, WM_KeyDown, VK_DOWN, 0)
               else Key:=VK_TAB ;
   VK_F5     : if ((OkG) and (Vide)) then BEGIN Key:=0 ; ZoomOuChoixArt(GS.Col,GS.Row) ; END ;
   VK_INSERT : if ((OkG) and (Vide)) then BEGIN Key:=0 ; ClickInsert(GS.Row) ; END ;
   VK_DELETE : if ((OkG) and (Shift=[ssCtrl])) then BEGIN Key:=0 ; ClickDel(GS.Row,True,True,True) ; END ;
//   VK_F10    : if Vide then BEGIN Key:=0 ; ClickValide ; END ;
   VK_HOME   : if Shift=[ssCtrl] then BEGIN Key:=0 ; GotoEntete ; END ;
//   VK_END    : if Shift=[ssCtrl] then BEGIN Key:=0 ; GotoPied ; END ;
{Ctrl+D}  68 : if Shift=[ssCtrl] then BEGIN Key:=0 ; AppelleDim(GS.Row) ; END ;
  end ;
end;

{==============================================================================================}
{=================================== Initialisation ===========================================}
{==============================================================================================}
procedure TOF_MouvStkEx.ChargeFromNature ;
Var CodeL : String ;
    i : integer ;
BEGIN
if ctxAffaire in V_PGI.PGIContexte then
begin
{$IFDEF BTP}
  GS.ListeParam:= GetInfoParPiece(Nature,'GPP_LISTESAISIE') ;
{$ELSE}
  GS.ListeParam:= GetInfoParPiece(Nature,'GPP_LISTEAFFAIRE');
{$ENDIF}
end else GS.ListeParam:= GetInfoParPiece(Nature,'GPP_LISTESAISIE') ;
GereEche:=GetInfoParPiece(Nature,'GPP_GEREECHEANCE') ;
CommentEnt:=GetInfoParPiece(Nature,'GPP_COMMENTENT') ;
CommentPied:=GetInfoParPiece(Nature,'GPP_COMMENTPIED') ;
CalcRupt:=GetInfoParPiece(Nature,'GPP_CALCRUPTURE') ;
ForceRupt:=(GetInfoParPiece(Nature,'GPP_FORCERUPTURE')='X') ;
EstAvoir:=(GetInfoParPiece(Nature,'GPP_ESTAVOIR')='X') ;
OuvreAutoPort:=(GetInfoParPiece(Nature,'GPP_OUVREAUTOPORT')='X') ;
VenteAchat:=GetInfoParPiece(Nature,'GPP_VENTEACHAT') ;
DimSaisie:=GetInfoParPiece(Nature,'GPP_DIMSAISIE') ;
if Not(ctxAffaire in V_PGI.PGIContexte) then     // PA a voir pb plantage ?
GereLot:=(GetInfoParPiece(Nature,'GPP_LOT')='X') ;
{$IFDEF CCS3}
GereSerie:=False ;
{$else}
GereSerie:=(GetInfoParPiece(Nature,'GPP_NUMEROSERIE')='X') ;
{$ENDIF}
GereAcompte:=(GetInfoParPiece(Nature,'GPP_ACOMPTE')='X') ;
PPInfosLigne.Clear ;
for i:=1 to 8 do
    BEGIN
    CodeL:=GetInfoParPiece(Nature,'GPP_IFL'+IntToStr(i)) ;
    if CodeL<>'' then PPInfosLigne.Add(RechDom('GCINFOLIGNE',CodeL,False)+';'+RechDom('GCINFOLIGNE',CodeL,True)) ;
    END ;
SetControlText('T_TITREPIECE',GetInfoParPiece(Nature,'GPP_LIBELLE')) ;
END;

procedure TOF_MouvStkEx.ToutAllouer ;
BEGIN
GS.RowCount:=NbRowsInit ; StCellCur:='' ;
// Pièce
TOBPiece:=TOB.Create('PIECE',Nil,-1)     ;
// Fiches
TOBArticles:=TOB.Create('ARTICLES',Nil,-1) ;
TOBConds:=TOB.Create('CONDS',Nil,-1) ;
TOBTarif:=TOB.Create('TARIF',Nil,-1) ;
TobLigneTarif := Tob.Create('_LIGNETARIF', niL, -1); 
// Divers
TOBCXV:=TOB.Create('',Nil,-1) ;
TOBDim:=TOB.Create('',Nil,-1) ;
TOBDesLots:=TOB.Create('',Nil,-1) ;
TOBSerie:=TOB.Create('',Nil,-1) ;
END ;

procedure TOF_MouvStkEx.ToutLiberer ;
BEGIN
GS.VidePile(False) ;
PurgePop(POPZ) ;
TOBPiece.Free ; TOBPiece:=Nil ;
TOBArticles.Free ; TOBArticles:=Nil ;
TOBConds.Free ; TOBConds:=Nil ;
TOBTarif.Free ; TOBTarif:=Nil ;
TobLigneTarif.Free; TobLigneTarif := nil;
TOBCXV.Free ; TOBCXV:=Nil ;
TOBDim.Free ; TOBDim:=Nil ;
TOBDesLots.Free ; TOBDesLots:=Nil ;
TOBSerie.Free ; TOBSerie:=Nil ;
END ;

procedure TOF_MouvStkEx.InitEnteteDefaut ( Totale : boolean ) ;
var CC : THValComboBox;
BEGIN
if Totale then
   BEGIN
   CleDoc.DatePiece:=V_PGI.DateEntree ;
   GP_DATEPIECE.Text:=DateToStr(CleDoc.DatePiece) ;

   // A Garder ??
   if VH^.EtablisDefaut<>'' then GP_ETABLISSEMENT.Value:=VH^.EtablisDefaut else
    if GP_ETABLISSEMENT.Values.Count>0 then GP_ETABLISSEMENT.Value:=GP_ETABLISSEMENT.Values[0] ;
   if VH_GC.GCDepotDefaut<>'' then GP_DEPOT.Value:=VH_GC.GCDepotDefaut else
    if GP_DEPOT.Values.Count>0 then GP_DEPOT.Value:=GP_DEPOT.Values[0] ;

   CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ; if CC<>Nil then PositionneEtabUser(CC) ;
   END ;
GP_NUMEROPIECE.Caption:='' ;
CleDoc.NumeroPiece:=0 ;
TOBPiece.GetEcran(Ecran) ;
TOBPiece.PutValue('GP_NATUREPIECEG',Nature) ;
TOBPiece.PutValue('GP_VENTEACHAT',VenteAchat) ;
TOBPiece.PutValue('GP_DEVISE',V_PGI.DevisePivot) ;
TOBPiece.PutValue('GP_FACTUREHT', 'X') ;
TOBPiece.PutValue('GP_SAISIECONTRE', '-') ;
AddLesSupEntete (TOBPiece);
if GP_DEPOT.Value<>'' then BZoomDepot.Enabled:=True else BZoomDepot.Enabled:=False;
END ;

procedure TOF_MouvStkEx.InitPieceCreation ;
BEGIN
CleDoc.Souche:=GetSoucheG(CleDoc.NaturePiece,VH^.EtablisDefaut,'') ;
CleDoc.NumeroPiece:=GetNumSoucheG(CleDoc.Souche,StrToDate(GP_DATEPIECE.text)) ;
GP_NUMEROPIECE.Caption:=TexteMessage[10] ;
InitTOBPiece(TOBPiece) ;
TOBPiece.PutValue('GP_NUMERO',CleDoc.NumeroPiece) ; TOBPiece.PutValue('GP_SOUCHE',CleDoc.Souche) ;
DEV.Code := TobPiece.GetValue('GP_DEVISE');
GetInfosDevise(DEV);
GP_DATEPIECE.SetFocus ;
if (Nature='TEM') then
   begin
   GP_DEPOTDEST.Plus := 'GDE_DEPOT<>"'+ GP_DEPOT.Value + '"';
   end;
END ;

procedure TOF_MouvStkEx.GotoEntete ;
BEGIN
if GP_DATEPIECE.CanFocus then BEGIN GP_DATEPIECE.SetFocus ; Exit ; END ;
if GP_ETABLISSEMENT.CanFocus then BEGIN GP_ETABLISSEMENT.SetFocus ; Exit ; END ;
if GP_DEPOT.CanFocus then BEGIN GP_DEPOT.SetFocus ; Exit ; END ;
//if GP_REFINTERNE.CanFocus then BEGIN GP_REFINTERNE.SetFocus ; Exit ; END ;
END ;

procedure TOF_MouvStkEx.ReInitPiece ;
BEGIN
ForcerFerme:=False ; GeneCharge:=False ;
ToutLiberer ; ToutAllouer ;
//BlocageSaisie(False) ;
DejaRentre:=False ;
InitEnteteDefaut(False) ;
InitPieceCreation ;
AffecteGrid(GS,Action) ;
if Action=taConsult then GS.MultiSelect:=False else GS.Col:=SG_RefArt ;
END ;

procedure TOF_MouvStkEx.InitPasModif ;
BEGIN
TOBPiece.SetAllModifie(False) ;
TOBDesLots.SetAllModifie(False)  ;
TOBSerie.SetAllModifie(False)  ;
END ;

procedure TOF_MouvStkEx.InitToutModif ;
Var NowFutur : TDateTime ;
BEGIN
NowFutur:=NowH ;
TOBPiece.SetAllModifie(True)    ; TOBPiece.SetDateModif(NowFutur) ;
TOBDesLots.SetAllModifie(True)  ; TOBSerie.SetAllModifie(True)  ;
END ;

procedure TOF_MouvStkEx.FlagDomaine ;
BEGIN
if GP_DOMAINE.Values.Count>1 then BEGIN GP_DOMAINE.Visible:=True ; TGP_DOMAINE.Visible:=True ; END ;
if Action=taCreat then PositionneDomaineUser(GP_DOMAINE) ;
END ;

procedure TOF_MouvStkEx.FlagDepot ;
BEGIN
if (ctxMode in V_PGI.PGIContexte) then
  begin
  GP_ETABLISSEMENT.Plus:='ET_SURSITE="X"';
  GP_DEPOT.Plus:='GDE_SURSITE="X"';
  end;
if Nature <> 'TEM' then
  begin
  GP_DEPOTDEST.Visible:=False ; GP_DEPOTDEST.TabStop:=False ;
  TGP_DEPOTDEST.Visible:=False ;
  end;
if VH_GC.GCMultiDepots then Exit ;
GP_DEPOT.Visible:=False ; GP_DEPOT.TabStop:=False ;
TGP_DEPOT.Visible:=False ;
END ;

procedure TOF_MouvStkEx.ChargeLaPiece;
Var Q : TQuery ;
BEGIN
AddLesSupEntete (TOBPiece);
// Lecture Piece
Q:=OpenSQL('SELECT * FROM PIECE WHERE '+WherePiece(CleDoc,ttdPiece,False),True) ;
TOBPiece.SelectDB('',Q) ;
Ferme(Q) ;
// Lecture Lignes
Q:=OpenSQL('SELECT * FROM LIGNE WHERE '+WherePiece(CleDoc,ttdLigne,False)+' ORDER BY GL_NUMLIGNE',True) ;
TOBPiece.LoadDetailDB('LIGNE','','',Q,False,True) ;
Ferme(Q) ;
PieceAjouteSousDetail(TOBPiece);
UG_ConstruireTobArt(TobPiece,TOBArticles) ;
// Lecture Lot
LoadLesLots(TOBPiece,TOBDesLots,CleDoc) ;
{$IFNDEF CCS3}
// Lecture Serie
LoadLesSeries(TOBPiece,TOBSerie,CleDoc) ;
{$ENDIF}
DEV.Code := TobPiece.GetValue('GP_DEVISE');
GetInfosDevise(DEV);
END ;

{==============================================================================================}
{=============================== Actions liées au Grid ========================================}
{==============================================================================================}
procedure TOF_MouvStkEx.EtudieColsListe ;
Var i : integer ;
    Nam,St : String ;
BEGIN
LesColonnes:=GS.Titres[0] ;
EtudieColsGrid(GS) ; St:=LesColonnes ;
if SG_RefArt > 0 then GS.ColLengths[SG_RefArt]:=18;
if SG_lib > 0 then GS.ColLengths[SG_Lib]:=70;
for i:=0 to GS.ColCount-1 do
   BEGIN
   Nam:=ReadTokenSt(St) ;
   IdentCols[i].ColName:=Nam ;
   IdentCols[i].ColTyp:=ChampToType(Nam) ;
   if IdentCols[i].ColTyp='DATE' then
      BEGIN
      GS.ColTypes[i]:='D' ;
      GS.ColFormats[i]:=ShortdateFormat ;
      END ;
   if Nam='GL_MOTIFMVT' then GS.ColFormats[i]:='CB=GCMOTIFMOUVEMENT' ;
   if SG_Px > 0 then GS.ColLengths[SG_Px]:=-1 ;
   if SG_Rem > 0 then GS.ColLengths[SG_Rem]:=-1 ;
   if SG_Dep > 0 then GS.ColLengths[SG_Dep]:=-1 ;
   END ;
END ;

procedure TOF_MouvStkEx.FormateZoneSaisie (ACol,ARow : Longint ) ;
Var St,StC : String ;
BEGIN
St:=GS.Cells[ACol,ARow] ; StC:=St ;
if ((ACol=SG_RefArt) or (ACol=SG_Aff) or (ACol=SG_Rep) or (ACol=SG_Dep)) then StC:=uppercase(Trim(St)) else
 if ACol=SG_Px then StC:=StrF00(Valeur(St),V_PGI.OkDecP) else
  if ACol=SG_Rem then StC:=StrF00(Valeur(St),ADecimP) else
   if ((ACol=SG_QF) or (ACol=SG_QS)) then
      BEGIN
      StC:=StrF00(Valeur(St),V_PGI.OkDecQ) ;
      END else StC:=FormateZoneDivers(St,ACol) ;
GS.Cells[ACol,ARow]:=StC ;
END ;

Function TOF_MouvStkEx.FormateZoneDivers ( St : String ; ACol : Longint ) : String ;
Var ColName,Typ,StC : String ;
BEGIN
Result:=St ; StC:=St ;
ColName:=IdentCols[ACol].ColName ; if ColName='' then Exit ;
Typ:=IdentCols[ACol].ColTyp ;
if (Typ='INTEGER') or (Typ='SMALLINT') then StC:=InttoStr(ValeurI(St)) else
 if (Typ='DOUBLE') or (Typ='RATE')  then StC:=StrF00(Valeur(St),V_PGI.OkDecV) else
  if Typ='DATE' then
     BEGIN
     if St<>'' then if Not IsValidDate(St) then StC:=GP_DATEPIECE.Text ;
     END else
   if Typ='BOOLEAN' then BEGIN if Uppercase(St)<>'X' then St:='-' ; StC:=Uppercase(St) ; END else
    if ((Typ='COMBO') and (Copy(GS.ColFormats[ACol],1,3)<>'CB=')) then StC:=Uppercase(St)
       else StC:=St ;
Result:=StC ;
END ;

{==============================================================================================}
{=============================== Evènements de le Grid ========================================}
{==============================================================================================}
procedure TOF_MouvStkEx.GSEnter(Sender: TObject);
Var bc,Cancel : boolean ;
    ACol,ARow : integer ;
begin
if (Nature='TEM') then
   begin
   if ((GP_DEPOT.Text='') or (GP_DEPOTDEST.Text='')) then
      begin
      HShowMessage(ErrorMessage[19],TFVierge(Ecran).Caption,'') ;
      if GP_DEPOT.Text='' then GP_DEPOT.SetFocus else GP_DEPOTDEST.SetFocus;
      exit;
      end else
      begin
      GP_DEPOT.Enabled:=False;
      GP_DEPOTDEST.Enabled:=False;
      end;                  
   end else
// JT eQualité 10823
  if TestDepotOblig then
  begin
    if GP_DEPOT.Text = '' then
    begin
      HShowMessage(ErrorMessage[20],TFVierge(Ecran).Caption,'');
      GP_DEPOT.SetFocus;
      exit;
    end;
  end;
bc:=False ; Cancel:=False ; ACol:=GS.Col ; ARow:=GS.Row ;
GSRowEnter(GS,GS.Row,bc,False) ;
GSCellEnter(GS,ACol,ARow,Cancel) ;
//EnabledGrid ;
DejaRentre:=True ;
end;

procedure TOF_MouvStkEx.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
if Ou>=GS.RowCount-1 then GS.RowCount:=GS.RowCount+NbRowsPlus ;
if Action<>taConsult then CreerTOBLignes(GS,TOBPiece,Nil,Nil,Ou) ;
GereEnabled(Ou) ;
ShowDetail(Ou) ;
end;

procedure TOF_MouvStkEx.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
if csDestroying in Ecran.ComponentState then Exit ;
if Action=taConsult then Exit ;
DepileTOBLignes(GS,TOBPiece,Ou,GS.Row) ;
GereLesLots(Ou) ;
GereLesSeries(Ou);
end;

procedure TOF_MouvStkEx.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if Not Cancel then
   BEGIN
   GS.ElipsisButton:=(GS.Col=SG_RefArt) or (GS.Col=SG_Dep) ;
   if ((CommentLigne) and (GS.Col=SG_Lib)) then GS.ElipsisButton:=true ;
   StCellCur:=GS.Cells[GS.Col,GS.Row] ;
   END;
end;

procedure TOF_MouvStkEx.GSCellExit(Sender: TObject; var ACol,ARow : Integer; var Cancel : Boolean);
begin
if csDestroying in Ecran.ComponentState then Exit ;
if Action=taConsult then Exit ;
if GS.Cells[ACol,ARow]=StCellCur then Exit ;
FormateZoneSaisie(ACol,ARow) ;
if ACol=SG_RefArt then TraiteArticle(ACol,ARow,Cancel,False,False,1) else
  //if ACol=SG_Dep then TraiteDepot(ACol,ARow,Cancel) else
   if ((ACol=SG_QF) or (ACol=SG_QS)) then BEGIN if Not TraiteQte(ACol,ARow) then Cancel:=True ; END else
    if ACol=SG_Lib then TraiteLibelle(ARow) else
     if ACol=SG_Motif then TraiteMotif(ARow)
    ;
AfficheLaLigne(ARow);
end;

procedure TOF_MouvStkEx.GSElipsisClick(Sender: TObject);
begin
// Articles
if GS.Col=SG_RefArt then ZoomOuChoixArt(GS.Col,GS.Row) else
   ;
end;

Function TOF_MouvStkEx.GereElipsis ( LaCol : integer ) : boolean ;
Var StWhere : string;
BEGIN
Result:=False ;
  //FV1 - 28/02/2018 : FS#2970 - RESINA - Problème de relation articles / dépots dans l'affichage de la liste des articles
  if  GP_DEPOT.Value <> '' then
    StWhere := ' GQ_DEPOT="' + GP_DEPOT.Value + '" '
  else
    StWhere := '';
  if LaCol=SG_RefArt then Result:=GetArticleRecherche(GS,TexteMessage[1],CleDoc.NaturePiece,GP_DOMAINE.Value,StWhere) else
   ;
END ;

procedure TOF_MouvStkEx.GSDblClick(Sender: TObject);
begin
if GS.Col=SG_RefArt then ZoomOuChoixArt(GS.Col,GS.Row) else
 ;
end;

{==============================================================================================}
{=============================== Manipulation des LIGNES ======================================}
{==============================================================================================}
Function TOF_MouvStkEx.InitLaLigne ( ARow : integer ; NewQte : double ) : T_ActionTarifArt ;
Var TOBL,TOBA   : TOB ;
BEGIN
Result:=ataOk ;
TOBL:=GetTOBLigne(TOBPiece,ARow) ; if TOBL=Nil then Exit ;
TOBA:=FindTOBArtRow(TOBPiece,TOBArticles,ARow) ;
InitLigneVide(TOBPiece,TOBL,Nil,Nil,ARow,NewQte) ;
ArtVersLigne (TOBPiece,TOBA,TOBL);
AfficheLaLigne(ARow) ;
END ;

Procedure TOF_MouvStkEx.ArtVersLigne ( TOBPiece,TOBA,TOBL : TOB ) ;
Var NaturePiece,Depot : String ;
    i,j               : integer ;
    RefUnique         : String ;
BEGIN
if TOBA=Nil then Exit ;
RefUnique:=TOBA.GetValue('GA_ARTICLE') ;
if RefUnique<>'' then
   BEGIN
   TOBL.PutValue('GL_LIBELLE',TOBA.GetValue('GA_LIBELLE')) ;
   TOBL.PutValue('GL_PRIXPOURQTE',1) ;
   TOBL.PutValue('GL_ESCOMPTABLE','-') ;
   TOBL.PutValue('GL_REMISABLEPIED','-') ;
   TOBL.PutValue('GL_REMISABLELIGNE','-') ;
   TOBL.PutValue('GL_TENUESTOCK',TOBA.GetValue('GA_TENUESTOCK')) ;
   TOBL.PutValue('GL_TARIFARTICLE','') ;
   TOBL.PutValue('GL_TYPELIGNE','ART') ;
   {Famille}
   TOBL.PutValue('GL_FAMILLENIV1',TOBA.GetValue('GA_FAMILLENIV1')) ;
   TOBL.PutValue('GL_FAMILLENIV2',TOBA.GetValue('GA_FAMILLENIV2')) ;
   TOBL.PutValue('GL_FAMILLENIV3',TOBA.GetValue('GA_FAMILLENIV3')) ;
   {Nomenclature}
   TOBL.PutValue('GL_TYPEARTICLE',TOBA.GetValue('GA_TYPEARTICLE')) ;
   TOBL.PutValue('GL_TYPENOMENC',TOBA.GetValue('GA_TYPENOMENC')) ;
   NaturePiece:=TOBL.GetValue('GL_NATUREPIECEG') ; Depot:=TOBL.GetValue('GL_DEPOT') ;
   for i:=1 to 5 do TOBL.PutValue('GL_FAMILLETAXE'+IntToStr(i),TOBA.GetValue('GA_FAMILLETAXE'+IntToStr(i))) ;
   {Unités de mesure}
   TOBL.PutValue('GL_QUALIFSURFACE',TOBA.GetValue('GA_QUALIFSURFACE')) ;
   TOBL.PutValue('GL_QUALIFVOLUME',TOBA.GetValue('GA_QUALIFVOLUME')) ;
   TOBL.PutValue('GL_QUALIFPOIDS',TOBA.GetValue('GA_QUALIFPOIDS')) ;
   TOBL.PutValue('GL_QUALIFLINEAIRE',TOBA.GetValue('GA_QUALIFLINEAIRE')) ;
   TOBL.PutValue('GL_QUALIFHEURE',TOBA.GetValue('GA_QUALIFHEURE')) ;
   TOBL.PutValue('GL_SURFACE',TOBA.GetValue('GA_SURFACE')) ;
   TOBL.PutValue('GL_VOLUME',TOBA.GetValue('GA_VOLUME')) ;
   TOBL.PutValue('GL_POIDSBRUT',TOBA.GetValue('GA_POIDSBRUT')) ;
   TOBL.PutValue('GL_POIDSNET',TOBA.GetValue('GA_POIDSNET')) ;
   TOBL.PutValue('GL_POIDSDOUA',TOBA.GetValue('GA_POIDSDOUA')) ;
   TOBL.PutValue('GL_LINEAIRE',TOBA.GetValue('GA_LINEAIRE')) ;
   TOBL.PutValue('GL_HEURE',TOBA.GetValue('GA_HEURE')) ;
   TOBL.PutValue('GL_QUALIFQTESTO',TOBA.GetValue('GA_QUALIFUNITESTO')) ;
   TOBL.PutValue('GL_QUALIFQTEVTE',TOBA.GetValue('GA_QUALIFUNITEVTE')) ;
   {Tables libres}
   for j:=1 to 9 do TOBL.PutValue('GL_LIBREART'+IntToStr(j),TOBA.GetValue('GA_LIBREART'+IntToStr(j))) ;
   TOBL.PutValue('GL_LIBREARTA',TOBA.GetValue('GA_LIBREARTA')) ;
   AffectePrixValo(TOBL,TOBA) ;
   {Prix}
   if Nature='TEM' then
      begin
      TOBL.PutValue('GL_DPA',TOBL.GetValue('GL_PMAP')) ;
      TOBL.PutValue('GL_PUHT',TOBL.GetValue('GL_PMAP')) ; TOBL.PutValue('GL_PUHTDEV',TOBL.GetValue('GL_PMAP')) ;
      TOBL.PutValue('GL_PUTTC',TOBL.GetValue('GL_PMAP')) ; TOBL.PutValue('GL_PUTTCDEV',TOBL.GetValue('GL_PMAP')) ;
      TOBL.PutValue('GL_PUHTNET',TOBL.GetValue('GL_PMAP')) ; TOBL.PutValue('GL_PUHTNETDEV',TOBL.GetValue('GL_PMAP')) ;
      TOBL.PutValue('GL_PUTTCNET',TOBL.GetValue('GL_PMAP')) ; TOBL.PutValue('GL_PUTTCNETDEV',TOBL.GetValue('GL_PMAP')) ;
      end else
      begin
      TOBL.PutValue('GL_PUHT',0) ; TOBL.PutValue('GL_PUHTDEV',0) ;
      TOBL.PutValue('GL_PUTTC',0) ; TOBL.PutValue('GL_PUTTCDEV',0) ;
      TOBL.PutValue('GL_PUHTNET',0) ; TOBL.PutValue('GL_PUHTNETDEV',0) ;
      TOBL.PutValue('GL_PUTTCNET',0) ; TOBL.PutValue('GL_PUTTCNETDEV',0) ;
      end;
   {Conditionnements}
   {Bloc-Notes}
   if GetInfoParPiece(NaturePiece,'GPP_BLOBART')='X' then TOBL.PutValue('GL_BLOCNOTE',TOBA.GetValue('GA_BLOCNOTE')) ;
   {Divers}
   TOBL.PutValue('GL_DATELIVRAISON',TOBPiece.GetValue('GP_DATELIVRAISON')) ;
   TOBL.PutValue('GL_CAISSE',TOBPiece.GetValue('GP_CAISSE')) ;
   END ;
END ;

Function TOF_MouvStkEx.PieceModifiee : boolean ;
BEGIN
Result:=False ;
if Action=taConsult then Exit ;
Result:=TOBPiece.IsOneModifie ;
END ;

{==============================================================================================}
{======================================== Affichages ==========================================}
{==============================================================================================}
Procedure TOF_MouvStkEx.ZoomOuChoixArt ( ACol,ARow : integer ) ;
Var RefUnique : String ;
    Cancel, OkArt : boolean ;
    TOBA,TOBL : TOB;
BEGIN
if ACol<>SG_RefArt then Exit ;
RefUnique:=GetCodeArtUnique(TOBPiece,ARow) ;
if RefUnique<>'' then
   BEGIN
   TOBA:=FindTOBArtRow(TOBPiece,TOBArticles,ARow) ;
   ZoomArticle(RefUnique,TobA.GetValue('GA_TYPEARTICLE'),Action) ;
   END else if Action<>taConsult then
   BEGIN
   TOBL:=GetTOBLigne(TOBPiece,ARow) ;
   if TOBL=nil then exit;
   if TOBL.GetValue('GL_TYPEDIM')='GEN' then
      BEGIN
      AppelleDim(ARow) ;
      END else
      BEGIN
      Cancel:=False ; OkArt:=IdentifierArticle(ACol,ARow,Cancel,True,False) ;
       if ((OkArt) and (Not Cancel)) then
          BEGIN
          if Not ArticleAutorise(TOBPiece,TOBArticles,CleDoc.NaturePiece,ARow) then
             BEGIN
             HShowMessage(ErrorMessage[1],TFVierge(Ecran).Caption,'') ;
             VideCodesLigne(TOBPiece,ARow) ; InitialiseTOBLigne(TOBPiece,Nil,Nil,ARow) ;
             END else
             BEGIN
             UpdateArtLigne(ARow,False,False,1) ;
             StCellCur:=GS.Cells[ACol,ARow] ;
             END ;
          END ;
      END ;
   END ;
END ;

procedure TOF_MouvStkEx.AfficheLaLigne ( ARow : integer ) ;
Var TOBL : TOB ;
    i,QTE   : integer ;
    PX,RemTot,MontantHT: Double ;
BEGIN
TOBL:=GetTOBLigne(TOBPiece,ARow) ;
TOBL.PutLigneGrid(GS,ARow,False,False,LesColonnes) ;
if TOBL.GetValue('GL_TYPEDIM')='GEN' then
  begin
  GS.Cells[SG_RefArt,ARow]:=TOBL.GetValue('GL_CODESDIM') ;
  end ;

  QTE:=TOBL.GetValue('GL_QTEFACT') ;
  RemTot:=TOBL.GetValue('GL_REMISELIGNE') ;
  Px:=TOBL.GetValue('GL_PUHTDEV') ;
  //MontantHT:=MontantHT+(Px*Qte) ;
  MontantHT:=(Px*Qte) ;
  if SG_Montant<>-1 then
  begin
    GS.Cells[SG_Montant,Arow]:=FloatToStr(MontantHT) ;
    TOBL.PutValue('GL_MONTANTHTDEV',MontantHT) ;
  end;
  if SG_Rem<>-1 then GS.Cells[SG_Rem,Arow]:=FloatToStr(RemTot) ;

for i:=1 to GS.ColCount-1 do FormateZoneSaisie(i,ARow) ;
END ;

Procedure TOF_MouvStkEx.ShowDetail ( ARow : integer ) ;
Var TOBArt,TOBLig : TOB ;
    RefUnique     : String ;
BEGIN
if ARow<=TOBPiece.Detail.Count then
   BEGIN
   TOBLig:=GetTOBLigne(TOBPiece,ARow) ;
   RefUnique:=TOBLig.GetValue('GL_ARTICLE') ;
   END else
   BEGIN
   TOBLig:=Nil ; RefUnique:='';
   END ;
TOBArt:=TOBArticles.FindFirst(['GA_ARTICLE'],[RefUnique],False) ;
MontreInfosLigne(TOBLig,TOBArt,Nil,Nil,INFOSLIGNE,PPInfosLigne) ;
TFVierge(Ecran).HMTrad.ResizeGridColumns(INFOSLIGNE) ;
END ;

Procedure TOF_MouvStkEx.GereEnabled ( ARow : integer ) ;
Var TOBLig : TOB ;
BEGIN
BZoomArticle.Enabled:=(GetCodeArtUnique(TOBPiece,ARow)<>'') ;
TOBLig:=GetTOBLigne(TOBPiece,ARow) ;
if TOBLig=nil then
   begin
   BZoomLot.Enabled:=False;
   BZoomSerie.Enabled:=False;
   BZoomSuivante.Enabled:=False;
   BZoomPrecedente.Enabled:=False;
   exit;
   end;
BZoomLot.Enabled:=(TOBLig.GetValue('GL_INDICELOT')>0) ;
BZoomSerie.Enabled:=(TOBLig.GetValue('GL_INDICESERIE')>0) ;
BZoomSuivante.Enabled:=(TOBPiece.GetValue('GP_DEVENIRPIECE')<>'') ;
BZoomPrecedente.Enabled:=(TOBLig.GetValue('GL_PIECEORIGINE')<>'') ;
END ;

{==============================================================================================}
{=============================== Actions sur le Entête ========================================}
{==============================================================================================}
procedure TOF_MouvStkEx.GP_DATEPIECEExit(Sender: TObject);
BEGIN
if csDestroying in Ecran.ComponentState then Exit ;
ExitDatePiece ;
END ;

Function TOF_MouvStkEx.ExitDatePiece : Boolean ;
Var Err : integer ;
    DD : TDateTime ;
begin
Result:=False ;
Err:=ControleDate(GP_DATEPIECE.Text) ;
if Err>0 then
   BEGIN
   HShowMessage(ErrorMessage[3+Err],TFVierge(Ecran).Caption,'') ;
   if GP_DATEPIECE.CanFocus then GP_DATEPIECE.SetFocus else GotoEntete ;
   Exit ;
   END ;
DD:=StrToDate(GP_DATEPIECE.Text) ;
if DD<V_PGI.DateDebutEuro then
   BEGIN
   HShowMessage(ErrorMessage[9],TFVierge(Ecran).Caption,'') ;
   if GP_DATEPIECE.CanFocus then GP_DATEPIECE.SetFocus else GotoEntete ;
   Exit ;
   END ;
if ((DD<=VH_GC.GCDateClotureStock) and (VH_GC.GCDateClotureStock>100)) then
   BEGIN
   HShowMessage(ErrorMessage[10],TFVierge(Ecran).Caption,'') ;
   if GP_DATEPIECE.CanFocus then GP_DATEPIECE.SetFocus else GotoEntete ;
   Exit ;
   END ;
CleDoc.DatePiece:=StrToDate(GP_DATEPIECE.Text) ;
PutValueDetail(TOBPiece,'GP_DATEPIECE',CleDoc.DatePiece) ;
Result:=True ;
end;

procedure TOF_MouvStkEx.GP_DEPOTChange(Sender: TObject);
Var i,IndiceLot,NbSel : integer ;
    TOBL : TOB ;
    Dep,OldDep  : String ;
    Okok : boolean ;
begin
if (Action<>taCreat) then Exit ;
if GeneCharge then Exit ;
if (Nature='TEM') then
   begin
   GP_DEPOTDEST.Plus := 'GDE_DEPOT<>"'+ GP_DEPOT.Value + '"';
   end;
if GP_DEPOT.Value='' then Exit ;
OldDep:=TOBPiece.GetValue('GP_DEPOT') ;
NbSel:=GS.NbSelected ; Okok:=False ;
if ((OldDep<>GP_DEPOT.Value) and (TOBArticles.Detail.Count>0)) then
   BEGIN
   if HShowMessage(ErrorMessage[18],TFVierge(Ecran).Caption,'')=mrYes then
      BEGIN
      for i:=0 to TOBPiece.Detail.Count-1 do
          BEGIN
          if ((NbSel=0) or (GS.IsSelected(i+1))) then
             BEGIN
             TOBL:=TOBPiece.Detail[i] ;
             IndiceLot:=TOBL.GetValue('GL_INDICELOT') ; if IndiceLot>0 then Continue ;
             Dep:=TOBL.GetValue('GL_DEPOT') ; if Dep=OldDep then TOBL.PutValue('GL_DEPOT',GP_DEPOT.Value) ;
             END ;
          END ;
      GS.ClearSelected ; Okok:=True ;
      END ;
   END ;
if ((NbSel>0) and (Okok)) then
   BEGIN
   TOBPiece.PutValue('GP_DEPOT',OldDep) ;
   GP_DEPOT.Value:=OldDep ;
   END ;
TOBPiece.GetEcran(Ecran) ;
end;

procedure TOF_MouvStkEx.GP_DEPOTDESTChange(Sender: TObject);
begin
if (Action<>taCreat) then Exit ;
TOBPiece.GetEcran(Ecran) ;
end;

procedure TOF_MouvStkEx.GP_ETABLISSEMENTChange(Sender: TObject);
Var Etab : String ;
begin
if (Action<>taCreat) then Exit ;
Etab:=GP_ETABLISSEMENT.Value ; if Etab='' then Exit ;
if Etab=TOBPiece.GetValue('GP_ETABLISSEMENT') then Exit ;
PutValueDetail(TOBPiece,'GP_ETABLISSEMENT',Etab) ;
if not VH_GC.GCMultiDepots then GP_DEPOT.Value:=Etab ;
CleDoc.Souche:=GetSoucheG(CleDoc.NaturePiece,Etab,'') ; CleDoc.NumeroPiece:=GetNumSoucheG(CleDoc.Souche,CleDoc.DatePiece) ;
PutValueDetail(TOBPiece,'GP_SOUCHE',CleDoc.Souche) ;
PutValueDetail(TOBPiece,'GP_NUMERO',CleDoc.NumeroPiece) ;
TOBPiece.GetEcran(Ecran) ;
end;

{==============================================================================================}
{=============================== Manipulation des articles ====================================}
{==============================================================================================}
Procedure TOF_MouvStkEx.CodesArtToCodesLigne ( TOBArt : TOB ; ARow : integer ) ;
Var TOBL : TOB ;
    RefSais : String ;
BEGIN
TOBL:=GetTOBLigne(TOBPiece,ARow) ;
TOBL.InitValeurs ;
RefSais:=Trim(Copy(GS.Cells[SG_RefArt,ARow],1,18)) ; TOBArt.PutValue('REFARTSAISIE',RefSais) ;
TOBL.PutValue('GL_ARTICLE',TOBArt.GetValue('GA_ARTICLE')) ;
TOBL.PutValue('GL_REFARTSAISIE',TOBArt.GetValue('REFARTSAISIE')) ;
TOBL.PutValue('GL_CODEARTICLE',TOBArt.GetValue('GA_CODEARTICLE')) ;
TOBL.PutValue('GL_REFARTBARRE',TOBArt.GetValue('REFARTBARRE')) ;
TOBL.PutValue('GL_REFARTTIERS',TOBArt.GetValue('REFARTTIERS')) ;
END ;

Function TOF_MouvStkEx.IdentifierArticle ( Var ACol,ARow : integer ; Var Cancel : boolean ; Click,FromMacro : Boolean ) : boolean ;
Var RefSais,StatutArt,RefUnique,CodeArticle : String ;
    TOBArt  : TOB ;
    QQ,QArt : TQuery ;
    MultiDim,Okok : Boolean ;
    RechArt,EtatRech : T_RechArt ;
BEGIN
if Not Click Then
   BEGIN
   RefSais:=GS.Cells[SG_RefArt,ARow] ;
   if RefSais='' then BEGIN Result:=True ; Exit ; END ;
   END;
Result:=False ; RefUnique:='' ; CodeArticle:='' ; StatutArt:='' ;
EtatRech:=traOk ; MultiDim:=False ;
TOBArt:=FindTOBArtSais(TOBArticles,RefSais) ;
if TOBArt<>Nil then
   BEGIN
   RechArt:=traOk ;
   END else
   BEGIN
   TOBArt:=CreerTOBArt(TOBArticles) ;
   if FromMacro then
      BEGIN
      QArt:=OpenSQL('Select * from ARTICLE Where GA_ARTICLE="'+RefSais+'"',True) ;
      if Not QArt.EOF then BEGIN TOBArt.SelectDB('',QArt) ; RechArt:=traOk ; END
                      else RechArt:=traErreur ;
      Ferme(QArt) ;
      END else
      BEGIN
      RechArt:=TrouverArticleSQL(CleDoc.NaturePiece,RefSais,GP_DOMAINE.Value,'',CleDoc.DatePiece,TOBArt,Nil) ;
      END ;
   END ;
Case RechArt of
   traErreur : BEGIN
               if Not FromMAcro then ErreurCodeArticle(RefSais) ;
               CodeArticle:='' ; RefUnique:=''; StatutArt:='' ;
               END ;
      traSus : BEGIN
               CodeArticle:='' ; RefUnique:=''; StatutArt:='' ;
               END ;
       traOk : BEGIN
               CodeArticle:=TOBArt.GetValue('GA_CODEARTICLE') ;
               RefUnique:=TOBArt.GetValue('GA_ARTICLE') ;
               StatutArt:=TOBArt.GetValue('GA_STATUTART') ;
               END ;
   traOkGene : BEGIN
               CodeArticle:=TOBArt.GetValue('GA_CODEARTICLE') ;
               RefUnique:=TOBArt.GetValue('GA_ARTICLE') ;
               StatutArt:='UNI' ;
               END ;
   traGrille : BEGIN
               CodeArticle:=TOBArt.GetValue('GA_CODEARTICLE') ;
               RefUnique:=TOBArt.GetValue('GA_ARTICLE') ;
               StatutArt:=TOBArt.GetValue('GA_STATUTART') ;
               GS.Cells[SG_RefArt,ARow]:=CodeArticle ;
               END ;
    traAucun : BEGIN
               GS.Col:=SG_RefArt ; GS.Row:=ARow ;
               if GereElipsis(SG_RefArt) then
                  BEGIN
                  RefUnique:=GS.Cells[SG_RefArt,ARow] ;
                  TOBArt:=FindTOBArtSais(TOBArticles,RefUnique) ;
                  if TOBArt<>Nil then
                     BEGIN
                     CodeArticle:=TOBArt.GetValue('GA_CODEARTICLE') ;
                     StatutArt:=TOBArt.GetValue('GA_STATUTART') ;
                     END else
                     BEGIN
                     QQ:=OpenSQL('Select * from ARTICLE Where GA_ARTICLE="'+RefUnique+'" ',True) ;
                     if Not QQ.EOF then
                        BEGIN
                        TOBArt:=CreerTOBArt(TOBArticles) ; TOBArt.SelectDB('',QQ) ;
                        CodeArticle:=QQ.FindField('GA_CODEARTICLE').AsString ;
                        StatutArt:=QQ.FindField('GA_STATUTART').AsString ;
                        END;
                     Ferme(QQ) ;
                     END ;
                  EtatRech:=EtudieEtat(TOBArt,CleDoc.NaturePiece) ;
                  if EtatRech=traGrille then StatutArt:='GEN' else StatutArt:='UNI' ;
                  GS.Cells[SG_RefArt,ARow]:=CodeArticle ;
                  END ;
               END ;
   END ;
if ((StatutArt='') or (CodeArticle='') or (RefUnique='')) then
   BEGIN
   Cancel:=True ; GS.Col:=ACol ; GS.Row:=ARow ;
   TOBArt.Free ; Exit ;
   END ;
if StatutArt='GEN' then
   BEGIN
   if (ctxMode in V_PGI.PGIContexte) then
      BEGIN
      if GS.Cells[SG_Lib,Arow]='' then NewLigneDim:=True;
      RemplirTOBDim(CodeArticle,ARow) ;
      if SelectMultiDimensions(CodeArticle,GS,TobPiece.GetValue('GP_DEPOT'),Action) then  //AC
         BEGIN
         MultiDim:=True ;
         TOBDim:=TheTOB ; TheTOB:=Nil ;
         END else
         BEGIN
         RefUnique:='' ; TOBDim.ClearDetail ; TheTOB:=Nil ;
         END ;
      END else
      BEGIN
      RefUnique:=SelectUneDimension(RefUnique) ;
      END ;
   if RefUnique='' then
      BEGIN
      Cancel:=True ; GS.Col:=ACol ; GS.Row:=ARow ;
      TOBArt.Free ; Exit ;
      END ;
   END ;
Okok:=(StatutArt='UNI') or ((StatutArt='GEN') and (MultiDim)) ;
if Not Okok then
   BEGIN
   TOBArt.Free ; TOBArt:=CreerTOBArt(TOBArticles) ; TOBArt.SelectDB('"'+RefUnique+'"',Nil) ;
   END ;
CodesArtToCodesLigne(TOBArt,ARow) ;
Result:=True ;
END ;

Procedure TOF_MouvStkEx.ChargeTOBDispo ( ARow : integer ) ;
Var TOBA : TOB ;
BEGIN
TOBA:=FindTOBArtRow(TOBPiece,TOBArticles,ARow) ; if TOBA=Nil then Exit ;
LoadTOBDispo(TOBA,False) ;
END ;

Procedure TOF_MouvStkEx.UpdateArtLigne ( ARow : integer ; FromMacro,Calc : boolean ; NewQte : double ) ;
BEGIN
ChargeTOBDispo(ARow) ;
InitLaLigne(ARow,NewQte) ;
if Not FromMacro then
   BEGIN
   TraiteLesDim(ARow,True) ;
   END ;
ShowDetail(ARow) ;
GP_DOMAINE.Enabled:=False ;
END ;

Procedure TOF_MouvStkEx.TraiteArticle ( Var ACol,ARow : integer ; Var Cancel : boolean ; FromMacro,Calc : boolean ; NewQte : double ) ;
Var OkArt : Boolean ;
    NewCol,NewRow : integer ;
    TypeL,TypeD : String ;
BEGIN
TypeD:=GetChampLigne(TOBPiece,'GL_TYPEDIM',ARow) ;
if GS.Cells[ACol,ARow]='' then
   BEGIN
   TypeL:=GetChampLigne(TOBPiece,'GL_TYPELIGNE',ARow) ;
   if ((TypeL<>'TOT') and (TypeD<>'GEN')) then
      BEGIN
      VideCodesLigne(TOBPiece,ARow) ;
      InitialiseTOBLigne(TOBPiece,Nil,Nil,ARow) ;
      END ;
   Exit ;
   END ;
if TypeD='GEN' then
   BEGIN
   if GS.Cells[ACol,ARow]=GetChampLigne(TOBPiece,'GL_CODESDIM',ARow) then Exit ;
      END ;
NewCol:=GS.Col ; NewRow:=GS.Row ;
OkArt:=IdentifierArticle(ACol,ARow,Cancel,False,FromMacro) ;
if ((OkArt) and (Not Cancel)) then
   BEGIN
   GS.Col:=NewCol ; GS.Row:=NewRow ;
   if Not ArticleAutorise(TOBPiece,TOBArticles,CleDoc.NaturePiece,ARow) then
      BEGIN
      HShowMessage(ErrorMessage[1],TFVierge(Ecran).Caption,'') ;
      VideCodesLigne(TOBPiece,ARow) ; InitialiseTOBLigne(TOBPiece,Nil,Nil,ARow) ;
      END else
      BEGIN
      UpdateArtLigne(ARow,FromMacro,Calc,NewQte) ;
      if Not TraiteRupture(ARow) then
         BEGIN
         if SG_QF>=0 then BEGIN GS.Cells[SG_QF,ARow]:='' ; TraiteQte(SG_QF,ARow) ; END ;
         if SG_QS>=0 then BEGIN GS.Cells[SG_QS,ARow]:='' ; TraiteQte(SG_QS,ARow) ; END ;
         END ;
      END ;
   END ;
END ;

Procedure TOF_MouvStkEx.TraiteMotif ( ARow : integer ) ;
Var TOBL : TOB ;
BEGIN
TOBL:=GetTOBLigne(TOBPiece,ARow) ; if TOBL=Nil then Exit ;
TOBL.PutValue('GL_MOTIFMVT',GS.CellValues[SG_Motif,ARow]) ;
END ;

Procedure TOF_MouvStkEx.TraiteLibelle ( ARow : integer ) ;
Var TOBL : TOB ;
BEGIN
TOBL:=GetTOBLigne(TOBPiece,ARow) ; if TOBL=Nil then Exit ;
TOBL.PutValue('GL_LIBELLE',GS.Cells[SG_Lib,ARow]) ;
END;

Procedure TOF_MouvStkEx.GereLesLots ( ARow : integer ) ;
Var TOBA,TOBDepot,TOBL,TOBNoeud : TOB ;
    Depot,RefUnique,SensPiece   : String ;
    IndiceLot           : integer ;
    Qte,NewQte          : Double ;
    Valid,OkSerie       : boolean ;
BEGIN
if Not GereLot then Exit ;
TOBA:=FindTOBArtRow(TOBPiece,TOBArticles,ARow) ; if TOBA=Nil then Exit ;
if TOBA.GetValue('GA_LOT')<>'X' then Exit ;
TOBL:=GetTOBLigne(TOBPiece,ARow) ; if TOBL=Nil then Exit ;
if Not TOBL.FieldExists('QTECHANGE') then BEGIN TOBL.AddChampSup('QTECHANGE',False) ; TOBL.PutValue('QTECHANGE','X') ; END ;
if TOBL.GetValue('QTECHANGE')<>'X' then Exit ;
Depot:=TOBL.GetValue('GL_DEPOT') ; RefUnique:=TOBL.GetValue('GL_ARTICLE') ; Qte:=TOBL.GetValue('GL_QTEFACT') ;
TOBDepot:=TOBA.FindFirst(['GQ_DEPOT'],[Depot],False) ;
if TOBDepot=Nil then
   BEGIN
   TOBDepot:=TOB.Create('DISPO',TOBA,-1) ;
   TOBDepot.PutValue('GQ_ARTICLE',RefUnique) ;
   TOBDepot.PutValue('GQ_DEPOT',Depot) ;
   TOBDepot.PutValue('GQ_CLOTURE','-') ;
   END ;
IndiceLot:=TOBL.GetValue('GL_INDICELOT') ;
if IndiceLot>0 then TOBNoeud:=TOBDesLots.Detail[IndiceLot-1] else
   BEGIN
   TOBNoeud:=CreerNoeudLot(TOBPiece,TOBDesLots) ;
   IndiceLot:=TOBDesLots.Detail.Count ;
   TOBL.PutValue('GL_INDICELOT',IndiceLot) ;
   END ;
Valid:=False ;
OkSerie := (GereSerie) and (TOBA.GetValue('GA_NUMEROSERIE')='X');
SensPiece:=GetInfoParPiece(Nature,'GPP_SENSPIECE') ;
Valid:=Entree_LigDispolot(TOBDepot,TOBNoeud,TOBL,TobSerie,OkSerie,Action);
if Action=taConsult then Exit ;
TOBL.PutValue('QTECHANGE','-') ;
if Not Valid then
   BEGIN
   NewQte:=Arrondi(TOBNoeud.GetValue('QUANTITE'),6) ;
   if NewQte<>Qte then
      BEGIN
      if SG_QS>0 then BEGIN GS.Cells[SG_QS,ARow]:=StrF00(NewQte,V_PGI.OkDecQ) ; TraiteQte(SG_QS,ARow) ; END ;
      if SG_QF>0 then BEGIN GS.Cells[SG_QF,ARow]:=StrF00(NewQte,V_PGI.OkDecQ) ; TraiteQte(SG_QF,ARow) ; END ;
      END ;
   END ;
END ;

Procedure TOF_MouvStkEx.GereLesSeries ( ARow : integer ) ;
Var TOBA,TOBL,TOBM : TOB ;
    QualQte : String ;
    IndiceSerie : integer ;
    Qte,NewQte,UnitePiece : Double ;
    Valid : boolean ;
BEGIN
{$IFDEF CCS3}
Exit ;
{$ELSE}
if Not GereSerie then Exit ;
TOBL:=GetTOBLigne(TOBPiece,ARow) ; if TOBL=Nil then Exit ;
if TOBL.GetValue('GL_TYPELIGNE')<>'ART' then Exit ;
if Not TOBL.FieldExists('QTECHANGE') then BEGIN TOBL.AddChampSup('QTECHANGE',False) ; TOBL.PutValue('QTECHANGE','X') ; END ;
if TOBL.GetValue('QTECHANGE')<>'X' then Exit ;
Qte:=TOBL.GetValue('GL_QTESTOCK') ;
TOBA:=FindTOBArtRow(TOBPiece,TOBArticles,ARow) ; if TOBA=Nil then Exit ;
if TOBA.GetValue('GA_NUMEROSERIE')<>'X' then Exit ;
TOBL.PutValue('QTECHANGE','-') ;
Valid:=False ;
GS.CacheEdit ; GS.SynEnabled:=False ;
IndiceSerie:=TOBL.GetValue('GL_INDICESERIE') ;
if IndiceSerie<=0 then
   BEGIN
   IndiceSerie:=TOBSerie.Detail.Count+1 ;
   TOBL.PutValue('GL_INDICESERIE',IndiceSerie) ;
   END ;
Valid:=Entree_SaisiSerie(TOBL,Nil,TOBSerie,Nil,Nil,Action) ;
GS.MontreEdit ; GS.SynEnabled:=True ;
if Action=taConsult then Exit ;
if Not Valid then
   BEGIN
   if (IndiceSerie > TOBSerie.Detail.Count ) then
      begin
      TOBL.PutValue('GL_INDICESERIE',0) ;
      NewQte:=0;
      end else
      begin
      NewQte:=TOBSerie.Detail[IndiceSerie-1].Detail.Count;
      QualQte:=TOBL.GetValue('GL_QUALIFQTESTO');
      TOBM:=VH_GC.MTOBMEA.FindFirst(['GME_QUALIFMESURE','GME_MESURE'],['PIE',QualQte],False) ;
      UnitePiece:=0 ; if TOBM<>Nil then UnitePiece:=TOBM.GetValue('GME_QUOTITE') ;
      if (UnitePiece=0) then UnitePiece:=1.0 ;
      NewQte:=Arrondi(NewQte/UnitePiece, 6);
      end;
   if NewQte<>Qte then
      BEGIN
      if SG_QS>0 then BEGIN GS.Cells[SG_QS,ARow]:=StrF00(NewQte,V_PGI.OkDecQ) ; TraiteQte(SG_QS,ARow) ; END ;
      if SG_QF>0 then BEGIN GS.Cells[SG_QF,ARow]:=StrF00(NewQte,V_PGI.OkDecQ) ; TraiteQte(SG_QF,ARow) ; END ;
      END ;
   TOBL.PutValue('QTECHANGE','-') ;
   END;
{$ENDIF}
END;

{==============================================================================================}
{====================================== DIMENSIONS ============================================}
{==============================================================================================}
Procedure TOF_MouvStkEx.RemplirTOBDim ( CodeArticle : String ; Ligne: Integer) ;
Var iLigne,jDim : integer ;
    TOBL,TOBD : TOB ;
    Q: TQuery ;
    QteDejaSaisi,ReliqDejaSaisi: Double ;
BEGIN
TOBDim.ClearDetail ;
ANCIEN_TOBDimDetailCount:=0 ;
iLigne:=0 ; ReliqDejaSaisi:=0;
Q:=OpenSQL('Select GA_Article from Article where GA_CodeArticle="'+CodeArticle+'" AND GA_STATUTART="DIM" order by GA_ARTICLE' ,True) ;
While not Q.EOF do
  begin
  TOBD:=TOB.Create('',TOBDim,-1) ;
  QteDejaSaisi:=CalcQteDejaSaisie(TOBPiece,Q.FindField('GA_ARTICLE').AsString) ;
  //ReliqDejaSaisi:=CalcQteDejaSaisie(TOBPiece_O,Q.FindField('GA_ARTICLE').AsString) ;
  InitDim(TOBD,Q.FindField('GA_ARTICLE').AsString,QteDejaSaisi,ReliqDejaSaisi) ;
  Q.Next ;
  end ;
ferme(Q) ;
if not NewLigneDim then
    Begin
    for jDim:=0 to TOBDim.Detail.Count-1 do
        begin
        TOBD:=TOBDim.Detail[jDim] ;
        iLigne:=Ligne ;
        TOBL:=TOBPiece.Detail[iLigne] ;
        While TOBL.GetValue('GL_CODEARTICLE')=CodeArticle do
          begin
          if TOBL.GetValue('GL_ARTICLE')=TOBD.GetValue('GA_ARTICLE') then
            begin
            QteDejaSaisi:=CalcQteDejaSaisie(TOBPiece,TOBD.GetValue('GA_ARTICLE')) ;
            //ReliqDejaSaisi:=CalcQteDejaSaisie(TOBPiece_O,TOBD.GetValue('GA_ARTICLE')) ;
            LigneVersDim(TOBL,TOBD,QteDejaSaisi,ReliqDejaSaisi) ;
            ANCIEN_TOBDimDetailCount:=ANCIEN_TOBDimDetailCount+1 ;
            Break ;
            end else
            begin
            iLigne:=iLigne+1 ;
            if iLigne>TOBPiece.Detail.Count-1 then break ;
            TOBL:=TOBPiece.Detail[iLigne] ;
            end ;
          end ;
        end ;
    end ;
if Not TOBDim.FieldExists('GP_DEVISE') then TOBDim.AddChampSup('GP_DEVISE',False) ;
if Not TOBDim.FieldExists('GP_TAUXDEV') then TOBDim.AddChampSup('GP_TAUXDEV',False) ;
if Not TOBDim.FieldExists('GP_SAISIECONTRE') then TOBDim.AddChampSup('GP_SAISIECONTRE',False) ;
if Not TOBDim.FieldExists('GP_FACTUREHT') then TOBDim.AddChampSup('GP_FACTUREHT',False) ;
if Not TOBDim.FieldExists('GP_VENTEACHAT') then TOBDim.AddChampSup('GP_VENTEACHAT',False) ;
if Not TOBDim.FieldExists('GP_NATUREPIECEG') then TOBDim.AddChampSup('GP_NATUREPIECEG',False) ;
if Not TOBDim.FieldExists('GP_DATEPIECE') then TOBDim.AddChampSup('GP_DATEPIECE',False) ;// recup tarif
if Not TOBDim.FieldExists('GP_TIERS') then TOBDim.AddChampSup('GP_TIERS',False) ;// recup tarif achat
if Not TOBDim.FieldExists('GP_TARIFTIERS') then TOBDim.AddChampSup('GP_TARIFTIERS',False) ;// recup tarif tiers
TOBDim.PutValue('GP_DEVISE',TOBPiece.GetValue('GP_DEVISE')) ;
TOBDim.PutValue('GP_TAUXDEV',TOBPiece.GetValue('GP_TAUXDEV')) ;
TOBDim.PutValue('GP_SAISIECONTRE',TOBPiece.GetValue('GP_SAISIECONTRE')) ;
TOBDim.PutValue('GP_FACTUREHT',TOBPiece.GetValue('GP_FACTUREHT')) ;
TOBDim.PutValue('GP_VENTEACHAT',TOBPiece.GetValue('GP_VENTEACHAT')) ;
TOBDim.PutValue('GP_NATUREPIECEG',TOBPiece.GetValue('GP_NATUREPIECEG')) ;
TOBDim.PutValue('GP_DATEPIECE',TOBPiece.GetValue('GP_DATEPIECE')) ;
TOBDim.PutValue('GP_TIERS',TOBPiece.GetValue('GP_TIERS')) ;
TOBDim.PutValue('GP_TARIFTIERS',TOBPiece.GetValue('GP_TARIFTIERS')) ;
TheTOB:=TOBDim ;
END ;

Procedure TOF_MouvStkEx.AppelleDim ( ARow : integer ) ;
Var TypeDim,CodeArticle : String ;
    TOBL,TOBGEN : TOB ;
    LigGen : integer ;
BEGIN
NewLigneDim:=False ;
TOBL:=GetTOBLigne(TOBPiece,ARow) ; if TOBL=Nil then Exit ;
TypeDim:=TOBL.GetValue('GL_TYPEDIM') ; if ((TypeDim<>'GEN') and (TypeDim<>'DIM')) then Exit ;
if TypeDim='DIM' then BEGIN CodeArticle:=TOBL.GetValue('GL_CODEARTICLE') ; LigGen:=-1 ; END
                 else BEGIN CodeArticle:=TOBL.GetValue('GL_CODESDIM') ; LigGen:=ARow ; END ;
if CodeArticle='' then Exit ;
RemplirTOBDim(CodeArticle, ARow) ;
if SelectMultiDimensions(CodeArticle,GS,TobPiece.GetValue('GP_DEPOT'),Action) then TOBDim:=TheTOB else TOBDim.ClearDetail ; //AC
TheTOB:=Nil ;
if LigGen<0 then
   BEGIN
   TOBGEN:=TOBPiece.FindFirst(['GL_CODESDIM','GL_TYPEDIM'],[CodeArticle,'GEN'],False) ;
   if TOBGen=Nil then Exit ; LigGen:=TOBGEN.GetValue('GL_NUMLIGNE') ;
   END ;
TraiteLesDim(LigGen,False) ;
END ;

Procedure TOF_MouvStkEx.TraiteLesDim (var ARow : integer ; NewArt : boolean ) ;
Var RefUnique,CodeArticle,LeCom,TypeDim : String ;
    TOBL,TOBD,TOBA : TOB ;
    i,NbD,LaLig,ACol,NbDimInsert,j,NbSup: integer ;
    Cancel,Premier,OkPxUnique,RemA   : boolean ;
    Qte,TotQte,Prix,PrixUnique,TotPrix,Remise,TotRem : Double ;
BEGIN
TOBL:=GetTOBLigne(TOBPiece,ARow) ; if TOBL=Nil then Exit ;
if TOBDim.Detail.Count<=0 then Exit ;
RemA:=False; PrixUnique:=0;
if NewArt then
   BEGIN
   TOBA:=FindTOBArtRow(TOBPiece,TOBArticles,ARow) ; if TOBA.GetValue('GA_STATUTART')<>'GEN' then Exit ;
   CodeArticle:=TOBL.GetValue('GL_CODEARTICLE') ;
   RemA:=(TOBA.GetValue('GA_REMISELIGNE')='X') ;
   END else
   BEGIN
   TypeDim:=TOBL.GetValue('GL_TYPEDIM') ; if ((TypeDim<>'GEN') and (TypeDim<>'DIM')) then Exit ;
   if TypeDim='DIM' then CodeArticle:=TOBL.GetValue('GL_CODEARTICLE') else CodeArticle:=TOBL.GetValue('GL_CODESDIM') ;
   END ;
if CodeArticle='' then Exit ;
if ANCIEN_TOBDimDetailCount<>0 then NbSup:=SupDimAZero(GS,CodeArticle,TOBDim,TOBPiece,ARow) else NbSup:=0 ; // Supprime si QTE=0 dans THDim
NbD:=TOBDim.Detail.Count-(ANCIEN_TOBDimDetailCount-NbSup) ; //ac
NbDimInsert:=NbD;
GS.RowCount:=GS.RowCount+NbD ;
CreerTOBLignes(GS,TOBPiece,Nil,Nil,TOBPiece.Detail.Count+NbD) ;
NbD:=0 ; TotQte:=0 ;
Premier:=True ; OkPxUnique:=True ; TotPrix:=0 ;
Remise:=0 ; TotRem:=0 ;

for i:=0 to TOBDim.Detail.Count-1 do
    BEGIN
    TOBD:=TOBDim.Detail[i] ; Qte:=0 ;
    {Infos depuis TOBDim}
    RefUnique:=TOBD.GetValue('GA_ARTICLE') ; CodeArticle:=Trim(Copy(RefUnique,1,18)) ;
    Qte:=Arrondi(TOBD.GetValue('GL_QTEFACT'),6) ; if Qte=0 then Continue ;
    if RemA then Remise:=Valeur(TOBD.GetValue('GL_REMISELIGNE')) ;//Si remiseligne autorisé sur l'article
    Prix:=TOBD.GetValue('GL_PUHTDEV');
    for J:=ARow to ARow+TOBDim.Detail.Count-1 do
       Begin
       TOBL:=TOBPiece.Detail[j] ;
       if TOBL.GetValue('GL_ARTICLE')=RefUnique then Break
          else TOBL:=Nil ;
       end ;
    if TOBL<>Nil then
       BEGIN
       {Modification ligne}
       LaLig:=TOBL.GetValue('GL_NUMLIGNE') ;
       END else
       BEGIN
       {Insertion nouvelle ligne}
       NbD:= i ;
       Inc(NbD) ; LaLig:=ARow+NbD ;
       GS.Cells[SG_RefArt,LaLig]:=RefUnique ;
       ACol:=SG_RefArt ; Cancel:=False ;
       if (NbDimInsert>0) and (ARow<TobPiece.Detail.Count) then
         begin
         InsertTOBLigne(TOBPiece,Lalig) ;
         TOBPiece.Detail[TobPiece.Detail.count-1].Free;
         GS.DeleteRow (TobPiece.Detail.count+1) ;
         NbDimInsert:=NbDimInsert-1;
         end;
       TraiteArticle(ACol,LaLig,Cancel,True,False,Qte) ;
       if Cancel then BEGIN Dec(NbD) ; Continue ; END ;
       GS.Cells[SG_RefArt,LaLig]:=CodeArticle ;
       END ;
    TOBL:=GetTOBLigne(TOBPiece,LaLig) ;
    TOBL.PutValue('GL_TYPEDIM','DIM') ;
    TOBL.PutValue('GL_QTEFACT',Qte) ; TOBL.PutValue('GL_QTESTOCK',Qte) ;
    if RemA then TOBL.PutValue('GL_REMISELIGNE',Remise) ; //Si remiseligne autorisé sur l'article
    TOBL.PutValue('GL_PUHTDEV',Prix) ;
    AfficheLaLigne(LaLig) ; TotQte:=TotQte+Qte ;TotRem:=TotRem+Remise; TotPrix:=Arrondi(TotPrix+Qte*Prix,DEV.Decimale) ;
    if Premier then PrixUnique:=Prix else if Prix<>PrixUnique then OkPxUnique:=False ;
    Premier:=False ;
    END ;
//CalculeLaSaisie(-1,-1,False) ;
{Article générique passe en commentaire}
TOBL:=GetTOBLigne(TOBPiece,ARow) ; LeCom:=TOBL.GetValue('GL_LIBELLE') ;
ClickDel(ARow,True,False) ;
ClickInsert(ARow) ;
TOBL:=GetTOBLigne(TOBPiece,ARow) ; TOBL.PutValue('GL_LIBELLE',LeCom) ;
TOBL.PutValue('GL_TYPEDIM','GEN') ;
TOBL.PutValue('GL_QTEFACT',TotQte) ; TOBL.PutValue('GL_QTESTOCK',TotQte) ;
if RemA then TOBL.PutValue('GL_REMISELIGNE',TotRem) ; //Si remiseligne autorisé sur l'article
TOBL.PutValue('GL_CODESDIM',CodeArticle) ;
if Not OkPxUnique then
   BEGIN
   PrixUnique:=0 ;
   if ((ctxMode in V_PGI.PGIContexte) and (TotQte<>0)) then PrixUnique:=Arrondi(TotPrix/TotQte,DEV.Decimale) ;
   END ;
TOBL.PutValue('GL_PUHTDEV',PrixUnique) ;
//AfficheLaLigne(ARow);  if DimSaisie='DIM' then GS.RowHeights[ARow]:=0 ;
{Rebalayage des lignes dimensionnées}
GS.RowHeights[TOBPiece.Detail.Count+1]:=GS.DefaultRowHeight ;
AffichageDim ;
TOBDim.ClearDetail ;
END ;

Procedure TOF_MouvStkEx.AffichageDim ;  // Pour gérer l'affichage par rapport à DimSaisie
Var TOBA,TOBL: Tob ;
Grille,CodeDim,LibDim: String ;
k,i: integer ;
Begin
for i:=0 to TOBPiece.Detail.Count-1 do
    BEGIN
    if DimSaisie='GEN' then GS.RowHeights[i+1]:=GS.DefaultRowHeight  ;
    TOBL:=TOBPiece.Detail[i] ; //if TOBL.GetValue('GL_CODEARTICLE')<>CodeArticle then Continue ;
    if TOBL.GetValue('GL_TYPEDIM')='DIM' then
       BEGIN
       if DimSaisie='GEN' then GS.RowHeights[i+1]:=0 else
          BEGIN
          TOBA:=FindTOBArtRow(TOBPiece,TOBArticles,i+1) ;
          if TOBA<>Nil then for k:=1 to 5 do
             BEGIN
             Grille:=TOBA.GetValue('GA_GRILLEDIM'+IntToStr(k)) ;
             CodeDim:=TOBA.GetValue('GA_CODEDIM'+IntToStr(k)) ;
             if ((Grille<>'') and (CodeDim<>'')) then
                BEGIN
                LibDim:=RechDom('GCGRILLEDIM'+IntToStr(k),Grille,True)+' '+GCGetCodeDim(Grille,CodeDim,k) ;
                LibDim:=Copy(TOBL.GetValue('GL_LIBELLE')+'  '+LibDim,1,70) ;
                TOBL.PutValue('GL_LIBELLE',LibDim) ;
                END ;
             END ;
          END ;
       END else if DimSaisie='DIM' then GS.RowHeights[i+1]:=0 ;
    AfficheLaLigne(i+1) ;
    END ;
End ;

{==============================================================================================}
{=============================== Manipulation des Qtés ========================================}
{==============================================================================================}
Function TOF_MouvStkEx.TraiteRupture ( ARow : integer ) : boolean ;
Var TOBL,TOBA,TOBD : TOB ;
    RefU,Depot,ArtSub,sComp : String ;
    QteSais,QteDisp,RatioVA,QteAdd,RatioAdd : Double ;
    AnnuArt : Boolean ;
BEGIN
Result:=True ;
if Action=taConsult then Exit ;
if ((CalcRupt='') or (CalcRupt='AUC')) then Exit ;
TOBL:=GetTOBLigne(TOBPiece,ARow) ; if TOBL=Nil then Exit ;
RefU:=TOBL.GetValue('GL_ARTICLE') ; if RefU='' then Exit ;
QteSais:=TOBL.GetValue('GL_QTESTOCK') ; if QteSais<=0 then Exit ;
if TOBL.GetValue('GL_TENUESTOCK')<>'X' then Exit ;
Depot:=TOBL.GetValue('GL_DEPOT') ; if Depot='' then Exit ;
TOBA:=FindTOBArtRow(TOBPiece,TOBArticles,ARow) ; if TOBA=Nil then Exit ;
TOBD:=TOBA.FindFirst(['GQ_DEPOT'],[Depot],False) ;
RatioAdd:=1 ;
RatioVA:=GetRatio(TOBL,Nil,trsStock) ; if RatioVA=0 then RatioVA:=1 ;
if TOBD=Nil then
   BEGIN
   QteDisp:=0 ; QteAdd:=0 ;
   END else
   BEGIN
   QteDisp:=TOBD.GetValue('GQ_PHYSIQUE') ; QteAdd:=0 ;
   if CalcRupt='DIS' then QteDisp:=QteDisp-TOBD.GetValue('GQ_RESERVECLI')-TOBD.GetValue('GQ_PREPACLI') ;
   END ;
QteSais:=(QteSais/RatioVA)-(QteAdd/RatioAdd) ;
if QteSais<=QteDisp then Exit ;
ArtSub:=Trim(Copy(TOBA.GetValue('GA_SUBSTITUTION'),1,18)) ;
if ArtSub<>'' then sComp:=TexteMessage[1]+' '+ArtSub+')' else sComp:='' ;
AnnuArt:=True ;
if ForceRupt then
   BEGIN
   if HShowMessage(ErrorMessage[2],TFVierge(Ecran).Caption,'')=mrYes then AnnuArt:=False ;
   END else
   BEGIN
   HShowMessage(ErrorMessage[3],TFVierge(Ecran).Caption,'') ;
   END ;
if AnnuArt then Result:=False ;
END ;

Function TOF_MouvStkEx.TraiteQte ( ACol,ARow : integer ) : Boolean ;
Var TOBL : TOB ;
    Qte : Double ;
    RefUnique : String ;
BEGIN
Result:=True ;
TOBL:=GetTOBLigne(TOBPiece,ARow) ; if TOBL=Nil then Exit ;
Qte:=Valeur(GS.Cells[ACol,ARow]) ;
if (Qte < 0) and ((Nature='EEX') or (Nature='SEX')) then
   BEGIN
   HShowMessage(ErrorMessage[17],TFVierge(Ecran).Caption,'') ;
   if SG_QF>=0 then GS.Cells[SG_QF,ARow]:='' ;
   if SG_QS>=0 then GS.Cells[SG_QS,ARow]:='' ;
   Result:=False ;
   Exit ;
   END ;
RefUnique:=TOBL.GetValue('GL_ARTICLE') ;
if ACol=SG_QF then
   BEGIN
   TOBL.PutValue('GL_QTEFACT',Valeur(GS.Cells[ACol,ARow])) ;
   if SG_QS<0 then TOBL.PutValue('GL_QTESTOCK',Valeur(GS.Cells[ACol,ARow])) ;
   END else
   BEGIN
   TOBL.PutValue('GL_QTESTOCK',Valeur(GS.Cells[ACol,ARow])) ;
   if SG_QF<0 then TOBL.PutValue('GL_QTEFACT',Valeur(GS.Cells[ACol,ARow])) ;
   END ;
Tobl.PutValue('GL_QTERESTE', Tobl.GetValue('GL_QTESTOCK')); { NEWPIECE }
if Not TraiteRupture(ARow) then
   BEGIN
   if SG_QF>=0 then GS.Cells[SG_QF,ARow]:='' ;
   if SG_QS>=0 then GS.Cells[SG_QS,ARow]:='' ;
   Result:=False ;
   END ;
if Result then TOBL.PutValue('QTECHANGE','X') ;
END ;

{==============================================================================================}
{=============================== Manipulation des Depots=======================================}
{==============================================================================================}
procedure TOF_MouvStkEx.AffecteTRE;
Var ind1 : integer;
    TobPieceDest, TobLotDest, TobSerieDest, TOBL, TOBTmp : TOB;
    RefPreced,RefSuiv : string;
begin
if TobPiece.GetValue('GP_NATUREPIECEG') <> 'TEM' then Exit;
TOBTmp:=TOB.Create('',Nil,-1);
TobPieceDest := TOB.Create('PIECE', nil, -1);
TobPieceDest.Dupliquer(TOBPiece, True, True, True);
TobLotDest := TOB.Create('',Nil,-1) ;
TobLotDest.Dupliquer(TOBDesLots, True, True, True);
TobSerieDest := TOB.Create('',Nil,-1) ;
TobSerieDest.Dupliquer(TOBSerie, True, True, True);

TobPiece.PutValue('GP_NATUREPIECEG', 'TRE');
TobPiece.PutValue('GP_SOUCHE', GetSoucheG(TobPiece.GetValue('GP_NATUREPIECEG'),
                                          TobPiece.GetValue('GP_ETABLISSEMENT'),
                                          TobPiece.GetValue('GP_DOMAINE')));
TobPiece.PutValue('GP_INDICEG', 0);
TobPiece.PutValue('GP_DEPOTDEST',TOBPiece.GetValue('GP_DEPOT'));
TobPiece.PutValue('GP_DEPOT', GP_DEPOTDEST.Value);

for ind1 := 0 to TobPiece.Detail.Count - 1 do
    begin
    TOBL:=TobPiece.Detail[ind1] ;
    RefPreced := EncodeRefPiece(TOBL);
    TOBL.PutValue('GL_PIECEORIGINE',RefPreced);
    TOBL.PutValue('GL_NATUREPIECEG', TOBPiece.GetValue('GP_NATUREPIECEG'));
    TOBL.PutValue('GL_SOUCHE',TOBPiece.GetValue('GP_SOUCHE'));
    TOBL.PutValue('GL_INDICEG',TOBPiece.GetValue('GP_INDICEG'));
    TOBL.PutValue('GL_DEPOT',TOBPiece.GetValue('GP_DEPOT'));
    end;

{Enregistrement nouvelle pièce}
if V_PGI.IoError=oeOk then
    BEGIN
    InitToutModif ;
    ValideLaNumerotation ;
    if GetInfoParPiece(Nature,'GPP_IMPIMMEDIATE')='X' then TOBPiece.PutValue('GP_EDITEE','X') ;
    END ;

RefSuiv   := EncodeRefPiece(TOBPiece);
TOBPieceDest.PutValue('GP_DEVENIRPIECE',RefSuiv);

{for ind1 := 0 to TobDesLots.Detail.Count - 1 do
    begin
    TOBL:=TobDesLots.Detail[ind1] ;
    TOBL.PutValue('GLL_NATUREPIECEG',TOBPiece.GetValue('GP_NATUREPIECEG'));
    TOBL.PutValue('GLL_SOUCHE',TOBPiece.GetValue('GP_SOUCHE'));
    TOBL.PutValue('GLL_INDICEG',TOBPiece.GetValue('GP_INDICEG'));
    TOBL.PutValue('GLL_NUMERO',TOBPiece.GetValue('GP_NUMERO'));
    end;
for ind1 := 0 to TobSerie.Detail.Count - 1 do
    begin
    TOBL := TobSerie.Detail[ind1] ;
    for ind2 := 0 to TOBL.Detail.Count - 1 do
        begin
        TOBL2 := TOBL.Detail[ind2] ;
        TOBL2.PutValue('GLS_NATUREPIECEG',TOBPiece.GetValue('GP_NATUREPIECEG'));
        TOBL2.PutValue('GLS_SOUCHE',TOBPiece.GetValue('GP_SOUCHE'));
        TOBL2.PutValue('GLS_INDICEG',TOBPiece.GetValue('GP_INDICEG'));
        TOBL2.PutValue('GLS_NUMERO',TOBPiece.GetValue('GP_NUMERO'));
        end;
    end;  }

if V_PGI.IoError=oeOk then ValideLesLignes(TOBPiece,TOBArticles,Nil,TOBTmp,Nil,Nil,Nil,False,False) ;
if V_PGI.IoError=oeOk then ValideLesLots ;
if V_PGI.IoError=oeOk then ValideLesArticles(TOBPiece,TOBArticles) ;
if V_PGI.IoError=oeOk then ValideLesSeries ;

if V_PGI.IoError=oeOk then if Not TobPiece.InsertDBByNivel(False) then V_PGI.IoError:=oeUnknown ;

TOBPiece.Dupliquer(TobPieceDest, True, True, True);
TOBDesLots.Dupliquer(TobLotDest, True, True, True);
TOBSerie.Dupliquer(TobSerieDest, True, True, True);

TobPieceDest.Free; TobPieceDest := nil;
TobLotDest.Free; TobLotDest := nil;
TobSerieDest.Free; TobSerieDest := nil;
end;

{==============================================================================================}
{===================================== Actions, boutons =======================================}
{==============================================================================================}
procedure TOF_MouvStkEx.BZoomArticleClick(Sender: TObject);
Var RefUnique : String ;
    TOBA : TOB;
begin
if Not BZoomArticle.Enabled then Exit ;
RefUnique:=GetCodeArtUnique(TOBPiece,GS.Row) ;
TOBA:=FindTOBArtRow(TOBPiece,TOBArticles,GS.Row) ;
ZoomArticle(RefUnique,TobA.GetValue('GA_TYPEARTICLE'),taConsult) ;
end;

procedure TOF_MouvStkEx.BZoomDepotClick(Sender: TObject);
Var Depot : String ;
begin
if Not BZoomDepot.Enabled then Exit ;
Depot:=GP_DEPOT.Value;
AGLLanceFiche('GC','GCDEPOT',Depot,'','ACTION=CONSULTATION') ;
end;

procedure TOF_MouvStkEx.BZoomLotClick(Sender: TObject);
Var TOBLig : TOB ;
    ARow : integer;
BEGIN
ARow := GS.Row;
TOBLig:=GetTOBLigne(TOBPiece,ARow) ;
if TOBLig=nil then exit;
TOBLig.PutValue('QTECHANGE','X') ;
GereLesLots(ARow);
end;

procedure TOF_MouvStkEx.BZoomSerieClick(Sender: TObject);
Var TOBLig : TOB ;
    ARow : integer;
BEGIN
ARow := GS.Row;
TOBLig:=GetTOBLigne(TOBPiece,ARow) ;
if TOBLig=nil then exit;
TOBLig.PutValue('QTECHANGE','X') ;
GereLesSeries(ARow);
end;

procedure TOF_MouvStkEx.BZoomSuivanteClick(Sender: TObject);
Var StSuiv : string;
    CD     : R_CleDoc ;
BEGIN
StSuiv:=TOBPiece.GetValue('GP_DEVENIRPIECE');
if StSuiv='' then exit;
DecodeRefPiece(StSuiv,CD) ;
if ExistePiece(CD) then AGLLanceFiche('GC','GCMOUVSTKEX','','', CD.NaturePiece + ';' +
        DateToStr(CD.Datepiece) + ';' + CD.Souche  + ';' +
        intToStr(CD.NumeroPiece) + ';' + intToStr(CD.Indice) + ';ACTION=CONSULTATION');
end;

procedure TOF_MouvStkEx.BZoomPrecedenteClick(Sender: TObject);
Var TOBLig : TOB ;
    StPrec : string;
    CD     : R_CleDoc ;
BEGIN
TOBLig:=GetTOBLigne(TOBPiece,GS.Row) ;
if TOBLig=nil then exit;
StPrec:=TOBLig.GetValue('GL_PIECEORIGINE');
if StPrec='' then exit;
DecodeRefPiece(StPrec,CD) ;
if ExistePiece(CD) then AGLLanceFiche('GC','GCMOUVSTKEX','','', CD.NaturePiece + ';' +
        DateToStr(CD.Datepiece) + ';' + CD.Souche  + ';' +
        intToStr(CD.NumeroPiece) + ';' + intToStr(CD.Indice) + ';ACTION=CONSULTATION');
end;

procedure TOF_MouvStkEx.BZoomMouseEnter(Sender: TObject);
begin
if csDestroying in Ecran.ComponentState then Exit ;
PopZoom97(BZoom,POPZ) ;
end;

procedure TOF_MouvStkEx.BImprimerClick(Sender: TObject);
BEGIN
if Action = taConsult then ValideImpression;
END;

procedure TOF_MouvStkEx.ClickDel ( ARow : integer ; AvecC,FromUser : Boolean; SupDim:Boolean=False ) ;
Var RefUnique,CodeArticle,TypeDim : String ;
    bc,cancel : boolean ;
    ACol,LaLig,i : integer ;
    TOBL : TOB ;
    QTE: Double ;
BEGIN
if Action=taConsult then Exit ;
if ((ARow<1) or (ARow>TOBPiece.Detail.Count)) then Exit ;
TOBL:=GetTOBLigne(TOBPiece,ARow) ; if TOBL=Nil then Exit ;
RefUnique:=TOBL.Getvalue('GL_ARTICLE') ; CodeArticle:=TOBL.GetValue('GL_CODESDIM') ; TypeDim:=TOBL.GetValue('GL_TYPEDIM') ;
QTE:=TOBL.GetValue('GL_QTEFACT') ;
//DetruitLot(ARow) ;
GS.CacheEdit ; GS.SynEnabled:=False ;
GS.DeleteRow(ARow) ;
if (TOBPiece.Detail.Count>0) then GS.RowHeights[TOBPiece.Detail.Count]:=GS.DefaultRowHeight  ;//ac
if ((TOBPiece.Detail.Count>1) and (ARow<>TOBPiece.Detail.Count)) then TOBPiece.Detail[ARow-1].Free else TOBPiece.Detail[ARow-1].InitValeurs ;
if GS.RowCount<NbRowsInit then GS.RowCount:=GS.RowCount+NbRowsPlus ;
GS.MontreEdit ; GS.SynEnabled:=True ;
if SupDim then MiseAJourGrid(GS,TOBPiece,Trim(Copy(RefUnique,1,18)),QTE,ARow) ;
NumeroteLignesGC(GS,TOBPiece) ;
bc:=False ; Cancel:=False ; ACol:=GS.Col ;
GSRowEnter(GS,ARow,bc,False) ;
GSCellEnter(GS,ACol,ARow,Cancel) ;
ShowDetail(ARow) ;
if ((CodeArticle<>'') and (TypeDim='GEN') and (FromUser)) then
   BEGIN
   Repeat
    TOBL:=TOBPiece.FindFirst(['GL_CODEARTICLE','GL_TYPEDIM'],[CodeArticle,'DIM'],False) ;
    if TOBL<>Nil then
       BEGIN
       LaLig:=TOBL.GetValue('GL_NUMLIGNE') ;
       ClickDel(LaLig,False,False) ;
       END ;
   Until TOBL=Nil ;
   END ;
if DimSaisie='DIM' then  //ac
  for i:=0 to TOBPiece.Detail.Count-1 do
  begin
  TOBL:=TOBPiece.Detail[i] ;
  if TOBL.GetValue('GL_TYPEDIM')='GEN' then GS.RowHeights[i+1]:=0 else GS.RowHeights[i+1]:=GS.DefaultRowHeight ;
  end ;
END ;

procedure TOF_MouvStkEx.ClickInsert ( ARow : integer ) ;
BEGIN
if Action=taConsult then Exit ;
if ((ARow<1) or (ARow>TOBPiece.Detail.Count)) then Exit ;
GS.CacheEdit ; GS.SynEnabled:=False ;
InsertTOBLigne(TOBPiece,ARow) ;
GS.InsertRow(ARow) ;
InitialiseTOBLigne(TOBPiece,Nil,Nil,ARow) ;
GS.Row:=ARow ; GS.MontreEdit ; GS.SynEnabled:=True ;
NumeroteLignesGC(GS,TOBPiece) ;
ShowDetail(ARow) ;
END ;

{=======================================================================================}
{======================================= VALIDATIONS ===================================}
{=======================================================================================}
procedure TOF_MouvStkEx.ValideLaNumerotation ;
Var ind : integer ;
    NaturePieceG : String ;
    TOB1, TOB2, TOB3 : TOB;
BEGIN
if (Action=taCreat) then
    BEGIN
    NaturePieceG:=TOBPiece.GetValue('GP_NATUREPIECEG') ;
    TOB1:=TOB.Create('',Nil,-1);
    TOB2:=TOB.Create('',Nil,-1);
    TOB3:=TOB.Create('',Nil,-1);
    if Not SetNumeroDefinitif(TOBPiece,TOB1,nil,TOB2,TOB3,Nil,Nil,Nil,nil) then V_PGI.IoError:=oePointage ;
    TOB1.Free; TOB2.Free; TOB3.Free;
    if GetInfoParPiece(NaturePieceG,'GPP_ACTIONFINI')='ENR' then TOBPiece.PutValue('GP_VIVANTE','-') ;
    GP_NUMEROPIECE.Caption:=IntToStr(TOBPiece.GetValue('GP_NUMERO')) ;
    Ind:=TOBPiece.GetValue('GP_INDICEG') ; if Ind>0 then GP_NUMEROPIECE.Caption:=GP_NUMEROPIECE.Caption+' '+IntToStr(Ind) ;
    END ;
END ;

procedure TOF_MouvStkEx.ValideLesLots ;
Var i,IndiceLot : integer ;
    TOBDL,TOBL : TOB ;
BEGIN
{$IFDEF CCS3}
Exit;
{$ELSE}
if Not GereLot then Exit ;
if TOBDesLots.Detail.Count<=0 then Exit ;
TOBDesLots.Detail[0].AddChampSup('UTILISE',True) ;
for i:=0 to TOBPiece.Detail.Count-1 do
    BEGIN
    IndiceLot:=GetChampLigne(TOBPiece,'GL_INDICELOT',i+1) ;
    if IndiceLot>0 then
       BEGIN
       TOBL:=TOBPiece.Detail[i] ; TOBDL:=TOBDesLots.Detail[IndiceLot-1] ;
       TOBDL.PutValue('UTILISE','X') ;
       UpdateNoeudLot(TOBL,TOBDL) ;
	   MAJDispoLot(TOBArticles,TOBL,TOBDL,VenteAchat,TOBPiece.GetValue('GP_DEPOT'));
       END ;
    END ;
for i:=TOBDesLots.Detail.Count-1 downto 0 do
    BEGIN
    TOBDL:=TOBDesLots.Detail[i] ;
    if TOBDL.GetValue('UTILISE')<>'X' then TOBDL.Free ;
    END ;
if TOBDesLots.Detail.Count>0 then TOBPiece.PutValue('GP_ARTICLESLOT','X') ;
if Not TOBDesLots.InsertDB(Nil) then V_PGI.IoError:=oeUnknown ;
{$ENDIF}
END ;

procedure TOF_MouvStkEx.ValideLesSeries ;
Var i,IndiceSerie : integer ;
    TOBSerLig,TOBL : TOB ;
BEGIN
{$IFDEF CCS3}
Exit;
{$ELSE}
if Not GereSerie then Exit ;
if TOBSerie.Detail.Count<=0 then Exit ;
TOBSerie.Detail[0].AddChampSup('UTILISE',True) ;
for i:=0 to TOBPiece.Detail.Count-1 do
    BEGIN
    IndiceSerie:=GetChampLigne(TOBPiece,'GL_INDICESERIE',i+1) ;
    if IndiceSerie>0 then
       BEGIN
       TOBL:=TOBPiece.Detail[i] ; TOBSerLig:=TOBSerie.Detail[IndiceSerie-1] ;
       TOBSerLig.PutValue('UTILISE','X') ;
       UpdateLesSeries(TOBL,TOBSerLig) ;
       MAJDispoSerie(TOBL,TOBSerLig,False);
       END ;
    END ;
for i:=TOBSerie.Detail.Count-1 downto 0 do
    BEGIN
    TOBSerLig:=TOBSerie.Detail[i] ;
    if TOBSerLig.GetValue('UTILISE')<>'X' then TOBSerLig.Free ;
    END ;
if Not TOBSerie.InsertDB(Nil) then V_PGI.IoError:=oeUnknown ;
{$ENDIF}
END ;

procedure TOF_MouvStkEx.ValideLaPiece ;
Var TOBTmp : TOB;
BEGIN
if Action=taConsult then BEGIN Close ; Exit ; END ;
//if EstAvoir then InverseAvoir ;
TOBTmp:=TOB.Create('',Nil,-1);
//ReaffecteDispo(TOBTmp,TOBPiece,TOBarticles);
ReaffecteDispoDiff(TOBarticles);
{Enregistrement nouvelle pièce}
if V_PGI.IoError=oeOk then
   BEGIN
   InitToutModif ;
   ValideLaNumerotation ;
   //ValideLaCotation(TOBPiece,TOBBases,TOBEches) ;
   ValideLaPeriode(TOBPiece) ;
   if GetInfoParPiece(Nature,'GPP_IMPIMMEDIATE')='X' then TOBPiece.PutValue('GP_EDITEE','X') ;
   END ;
if V_PGI.IoError=oeOk then ValideLesLignes(TOBPiece,TOBArticles,Nil,TOBTmp,Nil,Nil,Nil,False,False) ;
if V_PGI.IoError=oeOk then ValideLesLots ;
if V_PGI.IoError=oeOk then ValideLesArticles(TOBPiece,TOBArticles) ;
if V_PGI.IoError=oeOk then ValideLesSeries ;
if V_PGI.IoError=oeOk then AffecteTRE ;
if V_PGI.IoError=oeOk then TOBPiece.InsertDBByNivel(False) ;
TOBTmp.Free;
END ;

procedure TOF_MouvStkEx.ValideImpression ;
var TOBT : TOB;
BEGIN
TOBT:=TOB.Create('TIERS',Nil,-1);
ImprimerLaPiece(TOBPiece,TOBT,CleDoc) ;
TOBT.Free ;
END ;

Function TOF_MouvStkEx.SortDeLaLigne : boolean ;
Var ACol,ARow : integer ;
    Cancel : boolean ;
BEGIN
Result:=False ;
if GP_DATEPIECE.Focused then
   BEGIN
   if Not ExitDatePiece then Exit ;
   if GP_ETABLISSEMENT.CanFocus then GP_ETABLISSEMENT.SetFocus ;
   END ;
if Not DejaRentre then BEGIN Result:=True ; Exit ; END ;
//if Screen.ActiveControl<>GS then BEGIN Result:=True ; Exit ; END ;  JS le 20/10/03 FI10915
Result:=False ;
ACol:=GS.Col ; ARow:=GS.Row ; Cancel:=False ;
GSCellExit(Nil,ACol,ARow,Cancel) ; if Cancel then Exit ;
GSRowExit(Nil,ARow,Cancel,False) ; if Cancel then Exit ;
Result:=True ;
END ;

procedure TOF_MouvStkEx.ClickValide ;
Var io : TIOErr ;
    ResGC : integer ;
    SomQte : Double;
begin
// Tests et actions préalables
if Action=taConsult then Exit ;
if Not SortDeLaLigne then Exit ;
if Not PieceModifiee then Exit ;
DepileTOBLignes(GS,TOBPiece,GS.Row,1) ;
// Contrôle intégrité
ResGC:=GCPieceCorrecte(TOBPiece,TOBArticles,Nil,Ecran) ;
if ResGC=3 then ResGC:=0;
SomQte:=TOBPiece.Somme('GL_QTEFACT',['GL_TYPELIGNE'],['ART'],TRUE);
if (SomQte <=0) and ((Nature='EEX') or (Nature='SEX')) then ResGC:=3;
if ResGC>0 then BEGIN HShowMessage(ErrorMessage[10+ResGC],TFVierge(Ecran).Caption,''); Exit ;END ;
// Appels automatiques en fin de saisie
//if Not BeforeValide(Self,TOBPiece,TOBBases,TOBTiers,TOBArticles,DEV) then Exit ;
// Enregistrement de la saisie
ValideEnCours:=True ;
io:=Transactions(ValideLaPiece,0) ;
if io<>oeOk then
   begin
   if Not TOBPiece.FieldExists('REVALIDATION') then TOBPiece.AddChampSup('REVALIDATION',False) ;
   TOBPiece.PutValue('REVALIDATION', 'X') ;
   end ;
Case io of
        oeOk : BEGIN
               ForcerFerme:=True ;
               //AfterValide(Self,TOBPiece,TOBBases,TOBTiers,TOBArticles,DEV) ;
               END ;
   oeUnknown : BEGIN MessageAlerte(TexteMessage[4]) ; ValideEnCours:=False ; Exit ; END ;
    oeSaisie : BEGIN MessageAlerte(TexteMessage[5]) ; ValideEnCours:=False ; Exit ; END ;
  oePointage : BEGIN MessageAlerte(TexteMessage[7]) ; ValideEnCours:=False ; Exit ; END ;
  oeLettrage : BEGIN MessageAlerte(TexteMessage[6]) ; ValideEnCours:=False ; Exit ; END ;
  oeStock    : BEGIN MessageAlerte(TexteMessage[8]) ; ValideEnCours:=False ; Exit ; END ;
   END ;
if GetInfoParPiece(Nature,'GPP_IMPIMMEDIATE')='X' then
   BEGIN
   if ((Action=taCreat) or (Action=taModif)) then
      BEGIN
      io:=Transactions(ValideImpression,1) ;
      if io<>oeOk then MessageAlerte(TexteMessage[9]) ;
      END ;
   END ;
ValideEnCours:=False ;
if Action<>taCreat then
   BEGIN
   Close ;
   END else
   BEGIN
   VH_GC.GCLastRefPiece:=EncodeRefPiece(TOBPiece) ;
   MontreNumero(TOBPiece) ;
   if PasBouclerCreat then Close else ReInitPiece ;
   END ;
end;

Initialization
registerclasses([TOF_MouvStkEx]);

end.


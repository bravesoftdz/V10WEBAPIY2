unit TarifArticle;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, Grids, Hctrls, ComCtrls, ExtCtrls, HPanel, StdCtrls, Mask, UIUtil,
  Hent1, TarifUtil, SaisUtil, HSysMenu,UtilPGI,Hqry,
{$IFDEF EAGLCLIENT}
  eFiche,MaineAGL,
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_main,Fiche,
{$ENDIF}
  UTOB, math, hmsgbox, Ent1,AglInit, HDimension, Menus, TarifRapide,
{$IFNDEF CCS3}
  TarifCond,
{$ENDIF}
  M3FP,HRichEdt, HRichOLE, AglInitGc,utilArticle, EntGC, TntGrids,
  TntStdCtrls, TntComCtrls, TntExtCtrls;

Function EntreeTarifArticle (Action : TActionFiche; TarifTTC : Boolean=False) : boolean ;
Function SaisieTarifArticle (CodeArt: string; Action : TActionFiche; TarifTTC : Boolean=False) : boolean ;

type
  TFTarifArticle = class(TForm)
    PDIMENSION: THPanel;
    Dock971: TDock97;
    Toolbar972: TToolWindow97;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    HMTrad: THSystemMenu;
    MsgBox: THMsgBox;
    GF_CODEDIM1: THCritMaskEdit;
    TGF_GRILLEDIM1: THLabel;
    TGF_GRILLEDIM2: THLabel;
    GF_CODEDIM2: THCritMaskEdit;
    GF_CODEDIM3: THCritMaskEdit;
    TGF_GRILLEDIM3: THLabel;
    GF_CODEDIM4: THCritMaskEdit;
    GF_CODEDIM5: THCritMaskEdit;
    TGF_GRILLEDIM4: THLabel;
    TGF_GRILLEDIM5: THLabel;
    BChercher: TToolbarButton97;
    FindLigne: TFindDialog;
    BInfos: TToolbarButton97;
    POPZ: TPopupMenu;
    InfArticle: TMenuItem;
    PPIED: THPanel;
    GF_REMISE: THCritMaskEdit;
    TGF_REMISE: THLabel;
    GF_CASCADEREMISE: THValComboBox;
    TGF_CASCADEREMISE: THLabel;
    BSaisieRapide: TToolbarButton97;
    HMess: THMsgBox;
    PDEVISE: THPanel;
    TGF_DEVISE: THLabel;
    GF_DEVISE: THValComboBox;
    PARTICLE: THPanel;
    GF_CODEARTICLE: THCritMaskEdit;
    TGF_ARTICLE: THLabel;
    TGF_LIBELLE: TLabel;
    PQUANTITATIF: THPanel;
    G_Qte: THGrid;
    PQTECAT: THPanel;
    G_QCa: THGrid;
    PCATEGORIE: THPanel;
    G_Cat: THGrid;
    PTITRE: THPanel;
    HTitre: THMsgBox;
    CBQUANTITAIF: TCheckBox;
    CBCATTIERS: TCheckBox;
    CONDAPPLIC: THRichEditOLE;
    BVoirCond: TToolbarButton97;
    TCONDTARF: TToolWindow97;
    G_COND: THGrid;
    FComboTIE: THValComboBox;
    FComboART: THValComboBox;
    FComboLIG: THValComboBox;
    FComboPIE: THValComboBox;
    GF_CONDAPPLIC: THRichEditOLE;
    BCondAplli: TToolbarButton97;
    BCopierCond: TToolbarButton97;
    BCollerCond: TToolbarButton97;
    FTable: THValComboBox;
    FOpe: THValComboBox;
    TTYPETARIF: THLabel;
    PART: THPanel;
    G_ART: THGrid;
    ISigneEuro: TImage;
    TGF_PRIXCON: THLabel;
    ISigneFranc: TImage;
    GF_PRIXCON: THNumEdit;
    GA_PRIXPOURQTE: THNumEdit;
    TGA_PRIXPOURQTE: THLabel;
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject) ;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
    procedure GF_CODEARTICLEExit(Sender: TObject);
    procedure G_QteCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GF_CODEARTICLEElipsisClick(Sender: TObject);
    procedure G_QteElipsisClick(Sender: TObject);
    procedure G_QteRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_QteRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_QteEnter(Sender: TObject);
    procedure G_QteCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure BValiderClick(Sender: TObject);
    procedure G_QCaCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_QCaEnter(Sender: TObject);
    procedure G_QCaRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_QCaRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_QCaElipsisClick(Sender: TObject);
    procedure G_QCaCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_CatCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_CatRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_CatEnter(Sender: TObject);
    procedure G_CatCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_CatRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_CatElipsisClick(Sender: TObject);
    procedure GF_DEVISEChange(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure FindLigneFind(Sender: TObject);
    procedure InfArticleClick(Sender: TObject);
    procedure GF_REMISEChange(Sender: TObject);
    procedure BSaisieRapideClick(Sender: TObject);
    procedure CBQUANTITAIFClick(Sender: TObject);
    procedure CBCATTIERSClick(Sender: TObject);
    procedure BCondAplliClick(Sender: TObject);
    procedure BCopierCondClick(Sender: TObject);
    procedure BCollerCondClick(Sender: TObject);
    procedure BVoirCondClick(Sender: TObject);
    procedure TCONDTARFClose(Sender: TObject);
    procedure G_ARTCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_ARTCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_ARTElipsisClick(Sender: TObject);
    procedure G_ARTEnter(Sender: TObject);
    procedure G_ARTRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_ARTRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GF_CASCADEREMISEChange(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  protected
    { Déclarations privées }
    iTableLigne : integer ;
    FindDebut,FClosing : Boolean;
    StCellCur,LesColArt, LesColQtes, LesColQCa, LesColCat : String ;
    ColsInter : Array of boolean ;
    DEV       : RDEVISE ;
    TarifTTC  : Boolean ;
    PrixPourQte : Double ;
// Objets mémoire
    TOBTarif, TOBTarfArt, TOBTarfQte, TOBTarfQCa, TOBTarfCat, TOBTarifDel, TOBArt  : TOB;
// Menu
    procedure AffectMenuCondApplic (G_GRID : THGrid; ttd : T_TableTarif) ;
// Actions liées au Grid
    procedure EtudieColsListe ;  dynamic;
    Procedure InitLesCols ;
    procedure PostDrawCell(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    procedure FormateZoneSaisie (ACol,ARow : Longint; ttd : T_TableTarif) ; dynamic;
    procedure InsertLigne (ARow : Longint; ttd : T_TableTarif) ;
    procedure SupprimeLigne (ARow : Longint; ttd : T_TableTarif) ;
    procedure SupprimeTOBTarif (ARow : Longint; ttd : T_TableTarif) ;
    Function  GrilleModifie : Boolean; Dynamic ;
    Function  SortDeLaLigne (ttd : T_TableTarif) : boolean ;
// Initialisations
    procedure LoadLesTOB ; Dynamic ;
    procedure ErgoGCS3 ;
    procedure ChargeTarif ;
// ENTETE
    Procedure PrepareEntete ;
    Procedure InitialiseEntete ; Dynamic ;
    Procedure AffecteEntete ;
    Procedure AffecteDimension;
    Procedure ChercherArticle ;
    Procedure TraiterArticle ;
    Function  QuestionTarifEnCours : Integer;
// LIGNES
    Procedure InitLaLigne (ARow : integer; ttd : T_TableTarif) ; Dynamic ;
    //procedure VideCodesLigne ( ARow : integer ) ;
    Function  GetTOBLigne (ARow : integer;  ttd : T_TableTarif) : TOB ;
    {$IFDEF MONCODE}
    procedure AffichePrix (TOBL : TOB ; ARow : integer; ttd : T_TableTarif) ;
    {$ENDIF}
    procedure AfficheLaLigne (ARow : integer;  ttd : T_TableTarif) ;
    procedure InitialiseGrille ;
    procedure InitialiseLigne (ARow : integer; ttd : T_TableTarif) ;
    Procedure DepileTOBLigne;
    Procedure CreerTOBLigne (ARow : integer; ttd : T_TableTarif);
    Function  LigneVide ( ARow : integer; ttd : T_TableTarif) : Boolean; Dynamic ;
    Procedure PreAffecteLigne (ARow : integer; ttd : T_TableTarif); Dynamic ;
// CHAMPS LIGNES
    procedure TraiterDepot (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterCatTiers (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterLibelle (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterBorneInf (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterBorneSup (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterPrix (ACol, ARow : integer; ttd : T_TableTarif); Dynamic ;
    procedure TraiterRemise (ACol, ARow : integer; ttd : T_TableTarif); Dynamic ;
    procedure TraiterDateDeb (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterDateFin (ACol, ARow : integer; ttd : T_TableTarif);
// PIED
    Procedure InitialisePied ;
    Procedure AffichePied (ttd : T_TableTarif) ;
// Boutons
    procedure VoirFicheArticle;
// Validations
    procedure ValideTarif; Dynamic ;
    procedure VerifLesTOB; Dynamic ;
// Conditions tarifaires
    procedure InitComboChamps ;
    procedure RemplitComboChamps(NomTable : String ; FCombo : THValComboBox) ;
    procedure EffaceGrid ;
    Procedure AfficheCondTarf (ARow : Longint; ttd : T_TableTarif) ;
    function  ValueToItem(CC : THValComboBox ; St : String) : String ;
    procedure GetConditions (TOBL : TOB) ;
  public
    { Déclarations publiques }
    CodeArticle, CodeDevise : string;
    FicheAncetre : Boolean;
    Action      : TActionFiche ;
  end;


Const SA_Depot    : integer = 0 ;
      SA_Lib      : integer = 0 ;
      SA_Px       : integer = 0 ;
      SA_Rem      : integer = 0 ;
      SA_Datedeb  : integer = 0 ;
      SA_Datefin  : integer = 0 ;
      SA_Article  : integer = 0 ; //AC
      SA_TarifArticle : integer = 0 ; // AC
      SA_Arrondi : integer = 0 ; //AC

Const SG_Depot    : integer = 0 ;
      SG_Lib      : integer = 0 ;
      SG_QInf     : integer = 0 ;
      SG_QSup     : integer = 0 ;
      SG_Px       : integer = 0 ;
      SG_Rem      : integer = 0 ;
      SG_Datedeb  : integer = 0 ;
      SG_Datefin  : integer = 0 ;

Const SC_Depot    : integer = 0 ;
      SC_Cat      : integer = 0 ;
      SC_Lib      : integer = 0 ;
      SC_QInf     : integer = 0 ;
      SC_QSup     : integer = 0 ;
      SC_Px       : integer = 0 ;
      SC_Rem      : integer = 0 ;
      SC_Datedeb  : integer = 0 ;
      SC_Datefin  : integer = 0 ;

Const SC2_Depot   : integer = 0 ;
      SC2_Cat     : integer = 0 ;
      SC2_Lib     : integer = 0 ;
      SC2_Px      : integer = 0 ;
      SC2_Rem     : integer = 0 ;
      SC2_Datedeb : integer = 0 ;
      SC2_Datefin : integer = 0 ;


implementation

{$R *.DFM}

uses           
   ParamSoc,   
   Tarifs
	,CbpMCD
  ,CbpEnumerator

   ;

{***********A.G.L.***********************************************
Auteur  ...... : Michel Richaud
Créé le ...... : 11/05/2000
Modifié le ... : 11/05/2000
Description .. : Saisie des tarifs par articles
Mots clefs ... : TARIF; ARTICLE
*****************************************************************}
Procedure AppelTarifArticle ( Parms : array of variant ; nb : integer ) ;
var
  StArticle, StTarifHTouTTC, StAction, stNatureAuxi, stTarifArticle : string ;
  i_ind    : integer ;
    Action : TActionFiche ;
    TarifTTC : Boolean ;
    F : TFFiche;
  sFonctionnalite : string;

BEGIN
F := TFFiche (Longint (Parms[0]));
Action:=F.TypeAction;

  stArticle      := string(Parms[1]);         
  stTarifHTouTTC := string(Parms[2]);         
  stAction       := string(Parms[3]);         
  stNatureAuxi   := string(Parms[4]);         
  stTarifArticle := string(Parms[5]);         

  if (StArticle='') then Action:=taConsult ;  
  TarifTTC:=(stTarifHTouTTC = 'TTC');

  if GetParamSoc('SO_PREFSYSTTARIF') then
  begin
    sFonctionnalite:='';
    if       (stNatureAuxi='FOU') then sFonctionnalite:=sTarifFournisseur
    else if  (sTNatureAuxi='CLI') then sFonctionnalite:=sTarifClient
    else                               PGIError('Fonctionnalité non disponible, opération abandonnée');
    if (sFonctionnalite<>'') then
    begin
      AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE='+sFonctionnalite,'','ACTION=CREATION    ;DROIT=CDAMV;MONOFICHE;YTA_FONCTIONNALITE='+sFonctionnalite+';APPEL=ARTICLE'+';YTQ_ARTICLE='+stArticle+';YTQ_TARIFARTICLE='+stTarifArticle);
    end;
  end
  else
  begin
    if Action=tacreat then
    begin
      EntreeTarifArticle (Action, TarifTTC) ;
    end
    else
    begin
      SaisieTarifArticle (StArticle, Action, TarifTTC);
    end;
  end;
END ;

Function EntreeTarifArticle (Action : TActionFiche; TarifTTC : Boolean=False) : boolean ;
BEGIN
result:=SaisieTarifArticle ('', Action, TarifTTC) ;
END;

Function SaisieTarifArticle (CodeArt: string; Action : TActionFiche; TarifTTC : Boolean=False) : boolean ;
var FF : TFTarifArticle ;
    PPANEL  : THPanel ;
begin
SourisSablier;
FF := TFTarifArticle.Create(Application) ;
FF.Action:=Action ;
FF.TarifTTC:=TarifTTC ;
FF.CodeArticle:=CodeArt ;
if CodeArt<>'' then FF.FicheAncetre:=True else FF.FicheAncetre:=False ;
PPANEL := FindInsidePanel ;
if PPANEL = Nil then
   BEGIN
    try
      FF.BorderStyle:=bsSingle ;
      FF.ShowModal ;
    finally
      FF.Free ;
    end ;
   SourisNormale ;
   END else
   BEGIN
   InitInside (FF, PPANEL) ;
   FF.Show ;
   END ;
END ;

procedure TFTarifArticle.PostDrawCell(ACol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
begin
//
end ;

{==============================================================================================}
{======================================= Initialisations ======================================}
{==============================================================================================}
procedure TFTarifArticle.LoadLesTOB ;
Var Q : TQuery ;
    i_ind : integer;
    WhereTTC : string ;
BEGIN
for i_ind := TOBTarfArt.Detail.Count - 1 downto 0 do
    BEGIN
    TOBTarfArt.Detail[i_ind].Free;
    END;
for i_ind := TOBTarfQte.Detail.Count - 1 downto 0 do
    BEGIN
    TOBTarfQte.Detail[i_ind].Free;
    END;
for i_ind := TOBTarfQCa.Detail.Count - 1 downto 0 do
    BEGIN
    TOBTarfQCa.Detail[i_ind].Free;
    END;
for i_ind := TOBTarfCat.Detail.Count - 1 downto 0 do
    BEGIN
    TOBTarfCat.Detail[i_ind].Free;
    END;

if TarifTTC then WhereTTC := ' AND GF_REGIMEPRIX = "TTC" '
            else WhereTTC := ' AND GF_REGIMEPRIX <> "TTC" ' ;
// Lecture Quantitatif

Q := OpenSQL('SELECT * FROM TARIF WHERE '+ WhereTarifArt (CodeArticle, CodeDevise, ttdArt,False)+
             WhereTTC + ' ORDER BY GF_DEPOT, GF_BORNEINF',True,-1,'',true) ;
TOBTarfArt.LoadDetailDB('TARIF','','',Q,False) ;
Ferme(Q) ;
Q := OpenSQL('SELECT * FROM TARIF WHERE '+ WhereTarifArt (CodeArticle, CodeDevise, ttdArtQte,False)+
             WhereTTC + ' ORDER BY GF_DEPOT, GF_BORNEINF',True,-1,'',true) ;
TOBTarfQte.LoadDetailDB('TARIF','','',Q,False) ;
Ferme(Q) ;
Q := OpenSQL('SELECT * FROM TARIF WHERE '+ WhereTarifArt (CodeArticle, CodeDevise, ttdArtQCa,False)+
             WhereTTC + ' ORDER BY GF_DEPOT, GF_TARIFTIERS, GF_BORNEINF',True,-1,'',true) ;
TOBTarfQCa.LoadDetailDB('TARIF','','',Q,False) ;
Ferme(Q) ;
Q := OpenSQL('SELECT * FROM TARIF WHERE '+ WhereTarifArt (CodeArticle, CodeDevise, ttdArtCat,False)+
             WhereTTC + ' ORDER BY GF_DEPOT, GF_TARIFTIERS',True,-1,'',true) ;
TOBTarfCat.LoadDetailDB('TARIF','','',Q,False) ;
Ferme(Q) ;
END ;

procedure TFTarifArticle.ChargeTarif ;
var i_ind : integer;
BEGIN
LoadLesTOB ;
for i_ind:=0 to TOBTarfArt.Detail.Count-1 do
    BEGIN
    // Affichage
    AfficheLaLigne (i_ind + 1, ttdArt) ;
    END ;
for i_ind:=0 to TOBTarfQte.Detail.Count-1 do
    BEGIN
    // Affichage
    AfficheLaLigne (i_ind + 1, ttdArtQte) ;
    END ;
for i_ind:=0 to TOBTarfQCa.Detail.Count-1 do
    BEGIN
    // Affichage
    AfficheLaLigne (i_ind + 1, ttdArtQCa) ;
    END ;
for i_ind:=0 to TOBTarfCat.Detail.Count-1 do
    BEGIN
    // Affichage
    AfficheLaLigne (i_ind + 1, ttdArtCat) ;
    END ;
END ;


{==============================================================================================}
{=============================== Evènements de la Form ========================================}
{==============================================================================================}
procedure TFTarifArticle.ErgoGCS3 ;
BEGIN
{$IFDEF CCS3}
BVoirCond.Visible:=False   ; BCondAplli.Visible:=False ;
BCopierCond.Visible:=False ; BCollerCond.Visible:=False ;
{$ENDIF}
END ;

procedure TFTarifArticle.FormShow(Sender: TObject);
begin
if Action = taConsult then BValider.Visible:=False;

If (ctxAffaire in V_PGI.PGIContexte) then G_ART.ListeParam:='AFTARIFPRIX'
else G_ART.ListeParam:='GCTARIFPRIX' ;
G_ART.PostDrawCell:=PostDrawCell ;
If (ctxAffaire in V_PGI.PGIContexte) then  G_Qte.ListeParam:='AFTARIFQTEPRIX'
else   G_Qte.ListeParam:='GCTARIFQTEPRIX' ;
G_Qte.PostDrawCell:=PostDrawCell ;
If (ctxAffaire in V_PGI.PGIContexte) then G_QCa.ListeParam:='AFTARIFCTQTEPRIX'
else G_QCa.ListeParam:='GCTARIFCTQTEPRIX' ;
G_QCa.PostDrawCell:=PostDrawCell ;
If (ctxAffaire in V_PGI.PGIContexte) then  G_Cat.ListeParam:='AFTARIFCTPRIX'
else G_Cat.ListeParam:='GCTARIFCTPRIX' ;
G_Cat.PostDrawCell:=PostDrawCell ;
EtudieColsListe ;
HMTrad.ResizeGridColumns (G_Art) ;
HMTrad.ResizeGridColumns (G_Qte) ;
HMTrad.ResizeGridColumns (G_QCa) ;
HMTrad.ResizeGridColumns (G_Cat) ;
AffecteGrid (G_Art,Action) ;
AffecteGrid (G_Qte,Action) ;
AffecteGrid (G_QCa,Action) ;
AffecteGrid (G_Cat,Action) ;
PART.Visible := True;
PQUANTITATIF.Visible := False;
PQTECAT.Visible := False;
PCATEGORIE.Visible := FAlse;
PTITRE.Caption := HTitre.Mess[3] ;
if TarifTTC then TTYPETARIF.Caption := HTitre.Mess[6]
            else TTYPETARIF.Caption := HTitre.Mess[5] ;
CBQUANTITAIF.Checked := False ;
CBCATTIERS.Checked := False ;
CodeDevise := V_PGI.DevisePivot;
DEV.Code := CodeDevise ; GetInfosDevise (DEV) ;
if CodeArticle <> '' then PrepareEntete else InitialiseEntete ;
InitComboChamps ;
ErgoGCS3 ;
end;

procedure TFTarifArticle.FormCreate(Sender: TObject);
begin
G_Art.RowCount := NbRowsInit ;
G_Qte.RowCount := NbRowsInit ;
G_QCa.RowCount := NbRowsInit ;
G_Cat.RowCount := NbRowsInit ;
StCellCur := '' ;
iTableLigne := PrefixeToNum('GF') ;
TOBTarif := TOB.Create ('', Nil, -1) ;
TOBTarfArt := TOB.Create ('', TOBTarif, 0) ;
TOBTarfQte := TOB.Create ('', TOBTarif, 1) ;
TOBTarfQCa := TOB.Create ('', TOBTarif, 2) ;
TOBTarfCat := TOB.Create ('', TOBTarif, 3) ;
TOBTarifDel := TOB.Create ('', Nil, -1) ;
TOBArt := TOB.Create ('ARTICLE', Nil, -1) ;
InitLesCols ;
FClosing:=False ;
end;

procedure TFTarifArticle.FormDestroy(Sender: TObject) ;
begin
TOBTarif.Free ;
TOBTarifDel.free ;
TobArt.Free ;
end ;

procedure TFTarifArticle.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
if Action=taConsult then Exit ;
if GrilleModifie then
   BEGIN
   if MsgBox.Execute(6,Caption,'')<>mrYes then CanClose:=False ;
   END ;
end;

procedure TFTarifArticle.FormClose(Sender: TObject; var Action: TCloseAction);
begin
G_Art.VidePile(True) ;
G_Qte.VidePile(True) ;
G_QCa.VidePile(True) ;
G_Cat.VidePile(True) ;
TOBTarif.Free ; TOBTarif:=Nil ;
TOBArt.Free ; TOBArt:=Nil ;
TOBTarifDel.Free; TOBTarifDel:=nil ;
if IsInside(Self) then Action:=caFree ;
FClosing:=True ;
end;

procedure TFTarifArticle.FormKeyDown(Sender: TObject; var Key: Word;
                                     Shift: TShiftState);
var FocusGrid : Boolean;
    ttd : T_TableTarif;
    ARow : Longint;
BEGIN
FocusGrid := False;
if(Screen.ActiveControl = G_Art) then
    BEGIN
    FocusGrid := True;
    ttd := ttdArt;
    ARow := G_Art.Row;
    END else
    if(Screen.ActiveControl = G_Qte) then
        BEGIN
        FocusGrid := True;
        ttd := ttdArtQte;
        ARow := G_Qte.Row;
        END else
        if (Screen.ActiveControl = G_QCa) then
            BEGIN
            FocusGrid := True;
            ttd := ttdArtQCa;
            ARow := G_QCa.Row;
            END else
            if (Screen.ActiveControl = G_Cat) then
                BEGIN
                FocusGrid := True;
                ttd := ttdArtCat;
                ARow := G_Cat.Row;
                END;
Case Key of
    VK_RETURN : Key:=VK_TAB ;
    VK_INSERT : BEGIN
                if FocusGrid then
                    BEGIN
                    Key := 0;
                    InsertLigne (ARow, ttd);
                    END;
                END;
    VK_DELETE : BEGIN
                if ((FocusGrid) and (Shift=[ssCtrl])) then
                    BEGIN
                    Key := 0 ;
                    SupprimeLigne (ARow, ttd) ;
                    END ;
                END;
    END;
END;


{==============================================================================================}
{============================= Manipulation liées au Menu =====================================}
{==============================================================================================}
procedure TFTarifArticle.AffectMenuCondApplic (G_GRID : THGrid; ttd : T_TableTarif) ;
BEGIN
if (LigneVide (G_GRID.Row, ttd))or (Action = taConsult) then
    BEGIN
    BCondAplli.Enabled := False ; BCopierCond.Enabled := False; BCollerCond.Enabled := False ;
    END else
    BEGIN
    G_GRID.SetFocus;
    BCondAplli.Enabled := True ; BCopierCond.Enabled := True;
    if CONDAPPLIC.Text = '' then BCollerCond.Enabled := False
                            else BCollerCond.Enabled := True ;
    END;
END;

{==============================================================================================}
{=============================== Actions liées au Grid ========================================}
{==============================================================================================}
Procedure TFTarifArticle.InitLesCols ;
BEGIN
SG_Depot:=-1 ; SG_Lib:=-1 ; SG_QInf:=-1 ; SG_QSup:=-1 ;
SG_Px:=-1 ; SG_Rem:=-1 ; SG_Datedeb:=-1 ;  SG_Datefin:=-1 ;
SC_Depot:=-1 ; SC_Lib:=-1 ; SC_QInf:=-1 ; SC_QSup:=-1 ;
SC_Px:=-1 ; SC_Rem:=-1 ; SC_Datedeb:=-1 ;  SC_Datefin:=-1 ;
SC2_Depot:=-1 ; SC2_Lib:=-1 ;
SC2_Px:=-1 ; SC2_Rem:=-1 ; SC2_Datedeb:=-1 ;  SC2_Datefin:=-1 ;
END ;

procedure TFTarifArticle.EtudieColsListe ;
Var NomCol,LesCols : String ;
    icol,ichamp, i_ind : integer ;
Mcd : IMCDServiceCOM;
BEGIN
MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();
G_Art.ColWidths[0]:=0 ;
G_Qte.ColWidths[0]:=0 ;
G_QCa.ColWidths[0]:=0 ;
G_Cat.ColWidths[0]:=0 ;
SetLength(ColsInter,G_Qte.ColCount) ;
for i_ind:=Low(ColsInter)to High(ColsInter) do ColsInter[i_ind]:=False ;
LesCols:=G_Art.Titres[0] ; LesColArt:=LesCols ; icol:=0 ;
Repeat
 NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
 if NomCol<>'' then
    if assigned(mcd.getField(Nomcol)) then
    BEGIN
       if Pos('X',mcd.getField(Nomcol).control)>0 then ColsInter[icol]:=True ;
       if NomCol='GF_DEPOT'        then SA_Depot:=icol else
       if NomCol='GF_LIBELLE'      then SA_Lib:=icol else
       if NomCol='GF_PRIXUNITAIRE' then SA_Px:=icol else
       if NomCol='GF_CALCULREMISE' then SA_Rem:=icol else
       if NomCol='GF_DATEDEBUT'    then SA_Datedeb:=icol else
       if NomCol='GF_DATEFIN'      then SA_Datefin:=icol ;
    END ;
 Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;

LesCols:=G_Qte.Titres[0] ; LesColQtes:=LesCols ; icol:=0 ;
Repeat
 NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
 if NomCol<>'' then
    BEGIN
    if assigned(mcd.getField(Nomcol)) then
       BEGIN
       if Pos('X',mcd.getField(Nomcol).control)>0 then ColsInter[icol]:=True ;
       if NomCol='GF_DEPOT'        then SG_Depot:=icol else
       if NomCol='GF_LIBELLE'      then SG_Lib:=icol else
       if NomCol='GF_BORNEINF'     then SG_QInf:=icol else
       if NomCol='GF_BORNESUP'     then SG_QSup:=icol else
       if NomCol='GF_PRIXUNITAIRE' then SG_Px:=icol else
       if NomCol='GF_CALCULREMISE' then SG_Rem:=icol else
       if NomCol='GF_DATEDEBUT'    then SG_Datedeb:=icol else
       if NomCol='GF_DATEFIN'      then SG_Datefin:=icol ;
       END ;
    END ;
 Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;

LesCols:=G_QCa.Titres[0] ; LesColQCa := LesCols ; icol:=0 ;
Repeat
 NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
 if NomCol<>'' then
    BEGIN
    if assigned(mcd.getField(Nomcol)) then
       BEGIN
       if NomCol='GF_DEPOT'        then SC_Depot:=icol else
       if NomCol='GF_TARIFTIERS'   then SC_Cat:=icol else
       if NomCol='GF_LIBELLE'      then SC_Lib:=icol else
       if NomCol='GF_BORNEINF'     then SC_QInf:=icol else
       if NomCol='GF_BORNESUP'     then SC_QSup:=icol else
       if NomCol='GF_PRIXUNITAIRE' then SC_Px:=icol else
       if NomCol='GF_CALCULREMISE' then SC_Rem:=icol else
       if NomCol='GF_DATEDEBUT'    then SC_Datedeb:=icol else
       if NomCol='GF_DATEFIN'      then SC_Datefin:=icol ;
       END ;
    END ;
 Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;

LesCols:=G_Cat.Titres[0] ; LesColCat := LesCols ; icol:=0 ;
Repeat
 NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
 if NomCol<>'' then
    BEGIN
    if assigned(mcd.getField(Nomcol)) then
       BEGIN
       if NomCol='GF_DEPOT'        then SC2_Depot:=icol else
       if NomCol='GF_TARIFTIERS'   then SC2_Cat:=icol else
       if NomCol='GF_LIBELLE'      then SC2_Lib:=icol else
       if NomCol='GF_PRIXUNITAIRE' then SC2_Px:=icol else
       if NomCol='GF_CALCULREMISE' then SC2_Rem:=icol else
       if NomCol='GF_DATEDEBUT'    then SC2_Datedeb:=icol else
       if NomCol='GF_DATEFIN'      then SC2_Datefin:=icol ;
       END ;
    END ;
 Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;
END ;

procedure TFTarifArticle.FormateZoneSaisie (ACol,ARow : Longint; ttd : T_TableTarif) ;
Var St,StC : String ;
BEGIN
Case ttd of
    ttdArt    : BEGIN
                St:=G_Art.Cells[ACol,ARow] ; StC:=St ;
                if ACol=SA_Depot then StC:=uppercase(Trim(St)) else
                {$IFDEF MONCODE}
                    if ACol=SA_Px then StC:=StrF00(Valeur(St),DEV.Decimale+NbreDecimalSuppPrix('CLI',TOBArt)) ;
                {$ELSE}
                    if ACol=SA_Px then StC:=StrF00(Valeur(St),V_PGI.OkDecP) ;
                {$ENDIF}
                G_Art.Cells[ACol,ARow]:=StC ;
                END;
    ttdArtQte : BEGIN
                St:=G_Qte.Cells[ACol,ARow] ; StC:=St ;
                if ACol=SG_Depot then StC:=uppercase(Trim(St)) else
                {$IFDEF MONCODE}
                    if ACol=SG_Px then StC:=StrF00(Valeur(St),DEV.Decimale+NbreDecimalSuppPrix('CLI',TOBArt)) else
                {$ELSE}
                    if ACol=SG_Px then StC:=StrF00(Valeur(St),V_PGI.OkDecP) else
                {$ENDIF}
                        if ((ACol=SG_QInf) or (ACol=SG_QSup)) then StC:=StrF00(Valeur(St),V_PGI.OkDecQ);
                G_Qte.Cells[ACol,ARow]:=StC ;
                END;
    ttdArtQCa : BEGIN
                St:=G_QCa.Cells[ACol,ARow] ; StC:=St ;
                if ((ACol=SC_Depot) or (ACol=SC_Cat)) then StC:=uppercase(Trim(St)) else
                {$IFDEF MONCODE}
                    if ACol=SC_Px then StC:=StrF00(Valeur(St),DEV.Decimale+NbreDecimalSuppPrix('CLI',TOBArt)) else
                {$ELSE}
                    if ACol=SC_Px then StC:=StrF00(Valeur(St),V_PGI.OkDecP) else
                {$ENDIF}
                        if ((ACol=SC_QInf) or (ACol=SC_QSup)) then StC:=StrF00(Valeur(St),V_PGI.OkDecQ);
                G_QCa.Cells[ACol,ARow]:=StC ;
                END;
    ttdArtCat : BEGIN
                St:=G_Cat.Cells[ACol,ARow] ; StC:=St ;
                if ((ACol=SC2_Depot) or (ACol=SC2_Cat)) then StC:=uppercase(Trim(St)) else
                {$IFDEF MONCODE}
                    if ACol=SC2_Px then StC:=StrF00(Valeur(St),DEV.Decimale+NbreDecimalSuppPrix('CLI',TOBArt));
                {$ELSE}
                    if ACol=SC2_Px then StC:=StrF00(Valeur(St),DEV.Decimale);
                {$ENDIF}
                G_Cat.Cells[ACol,ARow]:=StC ;
                END;
    END;
END ;

procedure TFTarifArticle.InsertLigne (ARow : Longint; ttd : T_TableTarif) ;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if LigneVide (ARow, ttd) then exit;
Case ttd of
    ttdArt    : BEGIN
                if (ARow > TOBTarfArt.Detail.Count) then Exit;
                G_Art.CacheEdit; G_Art.SynEnabled := False;
                TOB.Create ('TARIF', TOBTarfArt, ARow-1) ;
                G_Art.InsertRow (ARow); G_Art.Row := ARow;
                InitialiseLigne (ARow, ttd) ;
                PreAffecteLigne (ARow, ttd);
                G_Art.MontreEdit; G_Art.SynEnabled := True;
                AffectMenuCondApplic (G_Art, ttd);
                AfficheCondTarf (G_Art.Row, ttd);
                END;
    ttdArtQte : BEGIN
                if (ARow > TOBTarfQte.Detail.Count) then Exit;
                G_Qte.CacheEdit; G_Qte.SynEnabled := False;
                TOB.Create ('TARIF', TOBTarfQte, ARow-1) ;
                G_Qte.InsertRow (ARow); G_Qte.Row := ARow;
                InitialiseLigne (ARow, ttd) ;
                PreAffecteLigne (ARow, ttd);
                G_Qte.MontreEdit; G_Qte.SynEnabled := True;
                AffectMenuCondApplic (G_Qte, ttd);
                AfficheCondTarf (G_Qte.Row, ttd);
                END;
    ttdArtQCa : BEGIN
                if (ARow > TOBTarfQCa.Detail.Count) then Exit;
                G_QCa.CacheEdit; G_QCa.SynEnabled := False;
                TOB.Create ('TARIF', TOBTarfQCa, ARow-1) ;
                G_QCa.InsertRow (ARow); G_QCa.Row := ARow;
                InitialiseLigne (ARow, ttd) ;
                PreAffecteLigne (ARow, ttd);
                G_QCa.MontreEdit; G_QCa.SynEnabled := True;
                AffectMenuCondApplic (G_QCa, ttd);
                AfficheCondTarf (G_QCa.Row, ttd);
                END;
    ttdArtCat : BEGIN
                if (ARow > TOBTarfCat.Detail.Count) then Exit;
                G_Cat.CacheEdit; G_Cat.SynEnabled := False;
                TOB.Create ('TARIF', TOBTarfCat, ARow-1) ;
                G_Cat.InsertRow (ARow); G_Cat.Row := ARow;
                InitialiseLigne (ARow, ttd) ;
                PreAffecteLigne (ARow, ttd);
                G_Cat.MontreEdit; G_Cat.SynEnabled := True;
                AffectMenuCondApplic (G_Cat, ttd);
                AfficheCondTarf (G_Cat.Row, ttd);
                END;
    END;
END;

procedure TFTarifArticle.SupprimeLigne (ARow : Longint; ttd : T_TableTarif) ;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
Case ttd of
    ttdArt    : BEGIN
                if (ARow > TOBTarfArt.Detail.Count) then Exit;
                G_Art.CacheEdit; G_Art.SynEnabled := False;
                G_Art.DeleteRow (ARow);
                if (ARow = TOBTarfQte.Detail.Count) then
                    CreerTOBLigne (ARow + 1, ttd);
                SupprimeTOBTarif (ARow, ttd);
                if G_Art.RowCount < NbRowsInit then G_Art.RowCount := NbRowsInit;
                G_Art.MontreEdit; G_Art.SynEnabled := True;
                AffectMenuCondApplic (G_Art, ttd);
                AfficheCondTarf (G_Art.Row, ttd);
                END;
    ttdArtQte : BEGIN
                if (ARow > TOBTarfQte.Detail.Count) then Exit;
                G_Qte.CacheEdit; G_Qte.SynEnabled := False;
                G_Qte.DeleteRow (ARow);
                if (ARow = TOBTarfQte.Detail.Count) then
                    CreerTOBLigne (ARow + 1, ttd);
                SupprimeTOBTarif (ARow, ttd);
                if G_Qte.RowCount < NbRowsInit then G_Qte.RowCount := NbRowsInit;
                G_Qte.MontreEdit; G_Qte.SynEnabled := True;
                AffectMenuCondApplic (G_Qte, ttd);
                AfficheCondTarf (G_Qte.Row, ttd);
                END;
    ttdArtQCa : BEGIN
                if (ARow > TOBTarfQCa.Detail.Count) then Exit;
                G_QCa.CacheEdit; G_QCa.SynEnabled := False;
                G_QCa.DeleteRow (ARow);
                if (ARow = TOBTarfQCa.Detail.Count) then
                    CreerTOBLigne (ARow + 1, ttd);
                SupprimeTOBTarif (ARow, ttd);
                if G_QCa.RowCount < NbRowsInit then G_QCa.RowCount := NbRowsInit;
                G_QCa.MontreEdit; G_QCa.SynEnabled := True;
                AffectMenuCondApplic (G_QCa, ttd);
                AfficheCondTarf (G_QCa.Row, ttd);
                END;
    ttdArtCat : BEGIN
                if (ARow > TOBTarfCat.Detail.Count) then Exit;
                G_Cat.CacheEdit; G_Cat.SynEnabled := False;
                G_Cat.DeleteRow (ARow);
                if (ARow = TOBTarfCat.Detail.Count) then
                    CreerTOBLigne (ARow + 1, ttd);
                SupprimeTOBTarif (ARow, ttd);
                if G_Cat.RowCount < NbRowsInit then G_Cat.RowCount := NbRowsInit;
                G_Cat.MontreEdit; G_Cat.SynEnabled := True;
                AffectMenuCondApplic (G_Cat, ttd);
                AfficheCondTarf (G_Cat.Row, ttd);
                END;
    END;
END;

procedure TFTarifArticle.SupprimeTOBTarif (ARow : Longint; ttd : T_TableTarif) ;
Var i_ind: integer;
BEGIN
Case ttd of
    ttdArt    : BEGIN
                if TOBTarfArt.Detail[ARow-1].GetValue ('GF_TARIF') <> 0 then
                    BEGIN
                    i_ind := TOBTarifDel.Detail.Count;
                    TOB.Create ('TARIF', TOBTarifDel, i_ind) ;
                    TOBTarifDel.Detail[i_ind].Dupliquer (TOBTarfArt.Detail[ARow-1], False, True);
                    END;
                TOBTarfArt.Detail[ARow-1].Free;
                END;
    ttdArtQte : BEGIN
                if TOBTarfQte.Detail[ARow-1].GetValue ('GF_TARIF') <> 0 then
                    BEGIN
                    i_ind := TOBTarifDel.Detail.Count;
                    TOB.Create ('TARIF', TOBTarifDel, i_ind) ;
                    TOBTarifDel.Detail[i_ind].Dupliquer (TOBTarfQte.Detail[ARow-1], False, True);
                    END;
                TOBTarfQte.Detail[ARow-1].Free;
                END;
    ttdArtQCa : BEGIN
                if TOBTarfQCa.Detail[ARow-1].GetValue ('GF_TARIF') <> 0 then
                    BEGIN
                    i_ind := TOBTarifDel.Detail.Count;
                    TOB.Create ('TARIF', TOBTarifDel, i_ind) ;
                    TOBTarifDel.Detail[i_ind].Dupliquer (TOBTarfQCa.Detail[ARow-1], False, True);
                    END;
                TOBTarfQCa.Detail[ARow-1].Free;
                END;
    ttdArtCat : BEGIN
                if TOBTarfCat.Detail[ARow-1].GetValue ('GF_TARIF') <> 0 then
                    BEGIN
                    i_ind := TOBTarifDel.Detail.Count;
                    TOB.Create ('TARIF', TOBTarifDel, i_ind) ;
                    TOBTarifDel.Detail[i_ind].Dupliquer (TOBTarfCat.Detail[ARow-1], False, True);
                    END;
                TOBTarfCat.Detail[ARow-1].Free;
                END;
    END;
END;

Function TFTarifArticle.GrilleModifie : Boolean;
BEGIN
Result:=False ;
if Action=taConsult then Exit ;
Result:=(TOBTarfArt.IsOneModifie) or (TOBTarfQte.IsOneModifie) or (TOBTarfQCa.IsOneModifie) or
        (TOBTarfCat.IsOneModifie) or (TOBTarifDel.IsOneModifie);
END;

Function TFTarifArticle.SortDeLaLigne (ttd : T_TableTarif) : boolean ;
Var ACol,ARow : integer ;
    Cancel : boolean ;
BEGIN
Result:=False ;
Case ttd of
    ttdArt    : BEGIN
                ACol:=G_Art.Col ; ARow:=G_Art.Row ; Cancel:=False ;
                G_ArtCellExit(Nil,ACol,ARow,Cancel) ; if Cancel then Exit ;
                G_ArtRowExit(Nil,ACol,Cancel,False) ; if Cancel then Exit ;
                END;
    ttdArtQte : BEGIN
                ACol:=G_Qte.Col ; ARow:=G_Qte.Row ; Cancel:=False ;
                G_QteCellExit(Nil,ACol,ARow,Cancel) ; if Cancel then Exit ;
                G_QteRowExit(Nil,ACol,Cancel,False) ; if Cancel then Exit ;
                END;
    ttdArtQCa : BEGIN
                ACol:=G_QCa.Col ; ARow:=G_QCa.Row ; Cancel:=False ;
                G_QCaCellExit(Nil,ACol,ARow,Cancel) ; if Cancel then Exit ;
                G_QCaRowExit(Nil,ACol,Cancel,False) ; if Cancel then Exit ;
              END;
    ttdArtCat : BEGIN
                ACol:=G_Cat.Col ; ARow:=G_Cat.Row ; Cancel:=False ;
                G_CatCellExit(Nil,ACol,ARow,Cancel) ; if Cancel then Exit ;
                G_CatRowExit(Nil,ACol,Cancel,False) ; if Cancel then Exit ;
              END;
    END;
Result:=True ;
END ;

{==============================================================================================}
{=============================== Evènements de la Grid ========================================}
{==============================================================================================}
procedure TFTarifArticle.G_ArtEnter(Sender: TObject);
var Cancel, Chg : Boolean;
    ACol, ARow : integer;
BEGIN
Cancel := False; Chg := False;
G_ArtRowEnter (Sender, G_Art.Row, Cancel, Chg);
ACol := G_Art.Col ; ARow := G_Art.Row ;
G_ArtCellEnter (Sender, ACol, ARow, Cancel);
END;

procedure TFTarifArticle.G_ArtRowEnter(Sender: TObject; Ou: Integer;
                                       var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
BEGIN
if Ou >= G_Art.RowCount - 1 then G_Art.RowCount := G_Art.RowCount + NbRowsPlus ;;
ARow := Min (Ou, TOBTarfArt.detail.count + 1);
if (ARow = TOBTarfArt.detail.count + 1) AND (not LigneVide (ARow - 1, ttdArt)) then
    BEGIN
    CreerTOBligne (ARow, ttdArt);
    END;
if (LigneVide (ARow, ttdArt)) AND (not LigneVide (ARow - 1, ttdArt))then
    PreAffecteLigne (ARow, ttdArt);
if Ou > TOBTarfArt.detail.count then
    BEGIN
    G_Art.Row := TOBTarfArt.detail.count;
    END;
AffichePied (ttdArt);
AfficheCondTarf (ARow, ttdArt);
END;

procedure TFTarifArticle.G_ArtRowExit(Sender: TObject; Ou: Integer;
                                      var Cancel: Boolean; Chg: Boolean);
BEGIN
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if LigneVide (Ou, ttdArt) Then G_Art.Row := Min (G_Art.Row,Ou);
END;

procedure TFTarifArticle.G_ArtCellEnter(Sender: TObject; var ACol,
                                        ARow: Integer; var Cancel: Boolean);
begin
If Action=taConsult then Exit ;
if Not Cancel then
    BEGIN
    if (G_Art.Col <> SA_Depot) AND (G_Art.Col <> SA_Lib) AND
       (G_Art.Cells [SA_Lib,G_Art.Row] = '') then G_Art.Col := SA_Lib;
    G_Art.ElipsisButton := ((G_Art.Col = SA_Depot) or (G_Art.col = SA_Datedeb) or
                            (G_Art.col = SA_Datefin)) ;
    StCellCur := G_Art.Cells [G_Art.Col,G_Art.Row] ;
    AffectMenuCondApplic (G_Art, ttdArt);
    END ;
end;

procedure TFTarifArticle.G_ArtCellExit(Sender: TObject; var ACol,
                                       ARow: Integer; var Cancel: Boolean);
BEGIN
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
FormateZoneSaisie (ACol,ARow, ttdArt);
if ACol = SA_Depot then TraiterDepot (ACol, ARow, ttdArt) else
    if ACol = SA_Lib then TraiterLibelle (ACol, ARow, ttdArt) else
      if ACol = SA_Px then TraiterPrix (ACol, ARow, ttdArt) else
         if ACol = SA_Rem then TraiterRemise (ACol, ARow, ttdArt) else
              if ACol = SA_Datedeb then TraiterDateDeb (ACol, ARow, ttdArt) else
                  if ACol = SA_Datefin then TraiterDateFin (ACol, ARow, ttdArt);
if Not Cancel then
    BEGIN
    END;
END;

procedure TFTarifArticle.G_ArtElipsisClick(Sender: TObject);
Var DEPOT, DATE : THCritMaskEdit;
    Coord : TRect;
begin
if G_Art.Col = SA_Depot then
    BEGIN
    Coord := G_Art.CellRect (G_Art.Col, G_Art.Row);
    DEPOT := THCritMaskEdit.Create (Self);
    DEPOT.Parent := G_Art;
    DEPOT.Top := Coord.Top;
    DEPOT.Left := Coord.Left;
    DEPOT.Width := 3; DEPOT.Visible := False;
    DEPOT.DataType := 'GCDEPOT';
    GetDepotRecherche (DEPOT) ;
    if DEPOT.Text <> '' then G_Art.Cells[G_Art.Col,G_Art.Row]:= DEPOT.Text;
    DEPOT.Destroy;
    END ;
if (G_Art.Col = SA_Datedeb) or (G_Art.Col = SA_Datefin) then
    BEGIN
    Coord := G_Art.CellRect (G_Art.Col, G_Art.Row);
    DATE := THCritMaskEdit.Create (Self);
    DATE.Parent := G_Art;
    DATE.Top := Coord.Top;
    DATE.left := Coord.Left;
    DATE.Width := 3; DATE.Visible := False;
    DATE.OpeType:=otDate;
    GetDateRecherche (TForm(DATE.Owner), DATE) ;
    if DATE.Text <> '' then G_Art.Cells[G_Art.Col,G_Art.Row]:= DATE.Text;
    DATE.Destroy;
    END;
END;

procedure TFTarifArticle.G_QteEnter(Sender: TObject);
var Cancel, Chg : Boolean;
    ACol, ARow : integer;
BEGIN
Cancel := False; Chg := False;
G_QteRowEnter (Sender, G_Qte.Row, Cancel, Chg);
ACol := G_Qte.Col ; ARow := G_Qte.Row ;
G_QteCellEnter (Sender, ACol, ARow, Cancel);
END;

procedure TFTarifArticle.G_QteRowEnter(Sender: TObject; Ou: Integer;
                                       var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
BEGIN
if Ou >= G_Qte.RowCount - 1 then G_Qte.RowCount := G_Qte.RowCount + NbRowsPlus ;;
ARow := Min (Ou, TOBTarfQte.detail.count + 1);
if (ARow = TOBTarfQte.detail.count + 1) AND (not LigneVide (ARow - 1, ttdArtQte)) then
    BEGIN
    CreerTOBligne (ARow, ttdArtQte);
    END;
if (LigneVide (ARow, ttdArtQte)) AND (not LigneVide (ARow - 1, ttdArtQte))then
    PreAffecteLigne (ARow, ttdArtQte);
if Ou > TOBTarfQte.detail.count then
    BEGIN
    G_Qte.Row := TOBTarfQte.detail.count;
    END;
AffichePied (ttdArtQte);
AfficheCondTarf (ARow, ttdArtQte);
END;

procedure TFTarifArticle.G_QteRowExit(Sender: TObject; Ou: Integer;
                                      var Cancel: Boolean; Chg: Boolean);
BEGIN
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if LigneVide (Ou, ttdArtQte) Then G_Qte.Row := Min (G_Qte.Row,Ou);
END;

procedure TFTarifArticle.G_QteCellEnter(Sender: TObject; var ACol,
                                        ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if Not Cancel then
    BEGIN
    if (G_Qte.Col <> SG_Depot) AND (G_Qte.Col <> SG_Lib) AND
       (G_Qte.Cells [SG_Lib,G_Qte.Row] = '') then G_Qte.Col := SG_Lib;
    G_Qte.ElipsisButton := ((G_Qte.Col = SG_Depot) or (G_Qte.col = SG_Datedeb) or
                            (G_Qte.col = SG_Datefin)) ;
    StCellCur := G_Qte.Cells [G_Qte.Col,G_Qte.Row] ;
    AffectMenuCondApplic (G_Qte, ttdArtQte);
    END ;
end;

procedure TFTarifArticle.G_QteCellExit(Sender: TObject; var ACol,
                                       ARow: Integer; var Cancel: Boolean);
BEGIN
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
FormateZoneSaisie (ACol,ARow, ttdArtQte);
if ACol = SG_Depot then TraiterDepot (ACol, ARow, ttdArtQte) else
    if ACol = SG_Lib then TraiterLibelle (ACol, ARow, ttdArtQte) else
        if ACol = SG_QInf then TraiterBorneInf (ACol, ARow, ttdArtQte) else
            if ACol = SG_QSup then TraiterBorneSup (ACol, ARow, ttdArtQte) else
                if ACol = SG_Px then TraiterPrix (ACol, ARow, ttdArtQte) else
                    if ACol = SG_Rem then TraiterRemise (ACol, ARow, ttdArtQte) else
                        if ACol = SG_Datedeb then TraiterDateDeb (ACol, ARow, ttdArtQte) else
                            if ACol = SG_Datefin then TraiterDateFin (ACol, ARow, ttdArtQte);
if Not Cancel then
    BEGIN
    END;
END;

procedure TFTarifArticle.G_QteElipsisClick(Sender: TObject);
Var DEPOT, DATE : THCritMaskEdit;
    Coord : TRect;
begin
if G_Qte.Col = SG_Depot then
    BEGIN
    Coord := G_Qte.CellRect (G_Qte.Col, G_Qte.Row);
    DEPOT := THCritMaskEdit.Create (Self);
    DEPOT.Parent := G_Qte;
    DEPOT.Top := Coord.Top;
    DEPOT.Left := Coord.Left;
    DEPOT.Width := 3; DEPOT.Visible := False;
    DEPOT.DataType := 'GCDEPOT';
    GetDepotRecherche (DEPOT) ;
    if DEPOT.Text <> '' then G_Qte.Cells[G_Qte.Col,G_Qte.Row]:= DEPOT.Text;
    DEPOT.Destroy;
    END ;
if (G_Qte.Col = SG_Datedeb) or (G_Qte.Col = SG_Datefin) then
    BEGIN
    Coord := G_Qte.CellRect (G_Qte.Col, G_Qte.Row);
    DATE := THCritMaskEdit.Create (Self);
    DATE.Parent := G_Qte;
    DATE.Top := Coord.Top;
    DATE.left := Coord.Left;
    DATE.Width := 3; DATE.Visible := False;
    DATE.OpeType:=otDate;
    GetDateRecherche (TForm(DATE.Owner), DATE) ;
    if DATE.Text <> '' then G_Qte.Cells[G_Qte.Col,G_Qte.Row]:= DATE.Text;
    DATE.Destroy;
    END;
END;

procedure TFTarifArticle.G_QCaEnter(Sender: TObject);
var Cancel, Chg : Boolean;
    ACol, ARow : integer;
BEGIN
Cancel := False; Chg := False;
G_QCaRowEnter (Sender, G_QCa.Row, Cancel, Chg);
ACol := G_QCa.Col ; ARow := G_QCa.Row ;
G_QCaCellEnter (Sender, ACol, ARow, Cancel);
end;

procedure TFTarifArticle.G_QCaRowEnter(Sender: TObject; Ou: Integer;
                                       var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
BEGIN
if Ou >= G_QCa.RowCount - 1 then G_QCa.RowCount := G_QCa.RowCount + NbRowsPlus ;;
ARow := Min (Ou, TOBTarfQCa.detail.count + 1);
if (ARow = TOBTarfQCa.detail.count + 1) AND (not LigneVide (ARow - 1, ttdArtQCa)) then
    BEGIN
    CreerTOBligne (ARow, ttdArtQCa);
    END;
if (LigneVide (ARow, ttdArtQCa)) AND (not LigneVide (ARow - 1, ttdArtQCa))then
    PreAffecteLigne (ARow, ttdArtQCa);
if Ou > TOBTarfQCa.detail.count then
    BEGIN
    G_QCa.Row := TOBTarfQCa.detail.count;
    END;
AffichePied (ttdArtQCa);
AfficheCondTarf (ARow, ttdArtQCa);
end;

procedure TFTarifArticle.G_QCaRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if LigneVide (Ou, ttdArtQCa) Then G_QCa.Row := Min (G_QCa.Row,Ou);
end;

procedure TFTarifArticle.G_QCaCellEnter(Sender: TObject; var ACol,
                                        ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if Not Cancel then
    BEGIN
    if (G_QCa.Col <> SC_Depot) AND (G_QCa.Col <> SC_Lib) AND (G_QCa.Col <> SC_Cat) then
        BEGIN
        if (G_QCa.Cells [SC_Lib,G_QCa.Row] = '') then G_QCa.Col := SC_Lib;
        if (G_QCa.Cells [SC_Cat,G_QCa.Row] = '') then G_QCa.Col := SC_Cat;
        END;
    G_QCa.ElipsisButton := ((G_QCa.Col = SC_Depot) or (G_QCa.Col = SC_Datedeb) or
                            (G_QCa.Col = SC_Datefin) or (G_QCa.Col = SC_Cat)) ;
    StCellCur := G_QCa.Cells [G_QCa.Col,G_QCa.Row] ;
    AffectMenuCondApplic (G_QCa, ttdArtQCa);
    END ;
end;

procedure TFTarifArticle.G_QCaCellExit(Sender: TObject; var ACol,
                                       ARow: Integer; var Cancel: Boolean);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
FormateZoneSaisie (ACol,ARow, ttdArtQCa);
if ACol = SC_Depot then TraiterDepot (ACol, ARow, ttdArtQCa) else
    if ACol = SC_Cat then TraiterCatTiers (ACol, ARow, ttdArtQCa) else
        if ACol = SC_Lib then TraiterLibelle (ACol, ARow, ttdArtQCa) else
            if ACol = SC_QInf then TraiterBorneInf (ACol, ARow, ttdArtQCa) else
                if ACol = SC_QSup then TraiterBorneSup (ACol, ARow, ttdArtQCa) else
                    if ACol = SC_Px then TraiterPrix (ACol, ARow, ttdArtQCa) else
                        if ACol = SC_Rem then TraiterRemise (ACol, ARow, ttdArtQCa) else
                            if ACol = SC_Datedeb then TraiterDateDeb (ACol, ARow, ttdArtQCa) else
                                if ACol = SC_Datefin then TraiterDateFin (ACol, ARow, ttdArtQCa);
end;

procedure TFTarifArticle.G_QCaElipsisClick(Sender: TObject);
Var DEPOT, DATE, CAT : THCritMaskEdit;
    Coord : TRect;
BEGIN
if G_QCa.Col = SC_Depot then
    BEGIN
    Coord := G_QCa.CellRect (G_QCa.Col, G_QCa.Row);
    DEPOT := THCritMaskEdit.Create (Self);
    DEPOT.Parent := G_QCa;
    DEPOT.Top := Coord.Top;
    DEPOT.Left := Coord.Left;
    DEPOT.Width := 3; DEPOT.Visible := False;
    DEPOT.DataType := 'GCDEPOT';
    GetDepotRecherche (DEPOT) ;
    if DEPOT.Text <> '' then G_QCa.Cells[G_QCa.Col,G_QCa.Row]:= DEPOT.Text;
    DEPOT.Destroy;
    END ;
if G_QCa.Col = SC_Cat then
    BEGIN
    Coord := G_QCa.CellRect (G_QCa.Col, G_QCa.Row);
    CAT := THCritMaskEdit.Create (Self);
    CAT.Parent := G_QCa;
    CAT.Top := Coord.Top;
    CAT.Left := Coord.Left;
    CAT.Width := 3; CAT.Visible := False;
    CAT.DataType := 'TTTARIFCLIENT';
    GetCategorieRecherche (CAT) ;
    if CAT.Text <> '' then G_QCa.Cells[G_QCa.Col,G_QCa.Row]:= CAT.Text;
    CAT.Destroy;
    END ;
if (G_QCa.Col = SC_Datedeb) or (G_QCa.Col = SC_Datefin) then
    BEGIN
    Coord := G_QCa.CellRect (G_QCa.Col, G_QCa.Row);
    DATE := THCritMaskEdit.Create (Self);
    DATE.Parent := G_QCa;
    DATE.Top := Coord.Top;
    DATE.left := Coord.Left;
    DATE.Width := 3; DATE.Visible := False;
    DATE.OpeType:=otDate;
    GetDateRecherche (TForm(DATE.Owner), DATE) ;
    if DATE.Text <> '' then G_QCa.Cells[G_QCa.Col,G_QCa.Row]:= DATE.Text;
    DATE.Destroy;
    END;
end;

procedure TFTarifArticle.G_CatEnter(Sender: TObject);
var Cancel, Chg : Boolean;
    ACol, ARow : integer;
BEGIN
Cancel := False; Chg := False;
G_CatRowEnter (Sender, G_Cat.Row, Cancel, Chg);
ACol := G_Cat.Col ; ARow := G_Cat.Row ;
G_CatCellEnter (Sender, ACol, ARow, Cancel);
end;

procedure TFTarifArticle.G_CatRowEnter(Sender: TObject; Ou: Integer;
                                       var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
BEGIN
if Ou >= G_Cat.RowCount - 1 then G_Cat.RowCount := G_Cat.RowCount + NbRowsPlus ;;
ARow := Min (Ou, TOBTarfCat.detail.count + 1);
if (ARow = TOBTarfCat.detail.count + 1) AND (not LigneVide (ARow - 1, ttdArtCat)) then
    BEGIN
    CreerTOBligne (ARow, ttdArtCat);
    END;
if (LigneVide (ARow, ttdArtCat)) AND (not LigneVide (ARow - 1, ttdArtCat))then
    PreAffecteLigne (ARow, ttdArtCat);
if Ou > TOBTarfCat.detail.count then
    BEGIN
    G_Cat.Row := TOBTarfCat.detail.count;
    END;
AffichePied (ttdArtCat);
AfficheCondTarf (ARow, ttdArtCat);
end;

procedure TFTarifArticle.G_CatRowExit(Sender: TObject; Ou: Integer;
                                      var Cancel: Boolean; Chg: Boolean);
BEGIN
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if LigneVide (Ou, ttdArtCat) Then G_Cat.Row := Min (G_Cat.Row,Ou);
END;

procedure TFTarifArticle.G_CatCellEnter(Sender: TObject; var ACol,
                                        ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if Not Cancel then
    BEGIN
    if (G_Cat.Col <> SC2_Depot) AND (G_Cat.Col <> SC2_Lib) AND (G_Cat.Col <> SC2_Cat) then
        BEGIN
        if (G_Cat.Cells [SC2_Lib,G_Cat.Row] = '') then G_Cat.Col := SC2_Lib;
        if (G_Cat.Cells [SC2_Cat,G_Cat.Row] = '') then G_Cat.Col := SC2_Cat;
        END;
    G_Cat.ElipsisButton := ((G_Cat.Col = SC2_Depot) or (G_Cat.Col = SC2_Datedeb) or
                            (G_Cat.Col = SC2_Datefin) or (G_Cat.Col = SC2_Cat)) ;
    StCellCur := G_Cat.Cells [G_Cat.Col,G_Cat.Row] ;
    AffectMenuCondApplic (G_Cat, ttdArtCat);
    END ;
end;

procedure TFTarifArticle.G_CatCellExit(Sender: TObject; var ACol,
                                       ARow: Integer; var Cancel: Boolean);
BEGIN
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
FormateZoneSaisie (ACol,ARow, ttdArtCat);
if ACol = SC2_Depot then TraiterDepot (ACol, ARow, ttdArtCat) else
    if ACol = SC2_Cat then TraiterCatTiers (ACol, ARow, ttdArtCat) else
        if ACol = SC2_Lib then TraiterLibelle (ACol, ARow, ttdArtCat) else
            if ACol = SC2_Px then TraiterPrix (ACol, ARow, ttdArtCat) else
                if ACol = SC2_Rem then TraiterRemise (ACol, ARow, ttdArtCat) else
                    if ACol = SC2_Datedeb then TraiterDateDeb (ACol, ARow, ttdArtCat) else
                        if ACol = SC2_Datefin then TraiterDateFin (ACol, ARow, ttdArtCat);
END;

procedure TFTarifArticle.G_CatElipsisClick(Sender: TObject);
Var DEPOT, DATE, CAT : THCritMaskEdit;
    Coord : TRect;
BEGIN
if G_Cat.Col = SC2_Depot then
    BEGIN
    Coord := G_Cat.CellRect (G_Cat.Col, G_Cat.Row);
    DEPOT := THCritMaskEdit.Create (Self);
    DEPOT.Parent := G_Cat;
    DEPOT.Top := Coord.Top;
    DEPOT.Left := Coord.Left;
    DEPOT.Width := 3; DEPOT.Visible := False;
    DEPOT.DataType := 'GCDEPOT';
    GetDepotRecherche (DEPOT) ;
    if DEPOT.Text <> '' then G_Cat.Cells[G_Cat.Col,G_Cat.Row]:= DEPOT.Text;
    DEPOT.Destroy;
    END ;
if G_Cat.Col = SC2_Cat then
    BEGIN
    Coord := G_Cat.CellRect (G_Cat.Col, G_Cat.Row);
    CAT := THCritMaskEdit.Create (Self);
    CAT.Parent := G_Cat;
    CAT.Top := Coord.Top;
    CAT.Left := Coord.Left;
    CAT.Width := 3; CAT.Visible := False;
    CAT.DataType := 'TTTARIFCLIENT';
    GetCategorieRecherche (CAT) ;
    if CAT.Text <> '' then G_Cat.Cells[G_Cat.Col,G_Cat.Row]:= CAT.Text;
    CAT.Destroy;
    END ;
if (G_Cat.Col = SC2_Datedeb) or (G_Cat.Col = SC2_Datefin) then
    BEGIN
    Coord := G_Cat.CellRect (G_Cat.Col, G_Cat.Row);
    DATE := THCritMaskEdit.Create (Self);
    DATE.Parent := G_Cat;
    DATE.Top := Coord.Top;
    DATE.left := Coord.Left;
    DATE.Width := 3; DATE.Visible := False;
    DATE.OpeType:=otDate;
    GetDateRecherche (TForm(DATE.Owner), DATE) ;
    if DATE.Text <> '' then G_Cat.Cells[G_Cat.Col,G_Cat.Row]:= DATE.Text;
    DATE.Destroy;
    END;
end;

{==============================================================================================}
{========================= Manipulation des LIGNES Quantitatif ================================}
{==============================================================================================}
Procedure TFTarifArticle.InitLaLigne (ARow : integer;  ttd : T_TableTarif);
Var TOBL : TOB;
BEGIN
TOBL:=GetTOBLigne(ARow, ttd) ; if TOBL=Nil then Exit ;
TOBL.PutValue ('GF_ARTICLE', CodeArticle);
TOBL.PutValue ('GF_BORNESUP', 999999);
TOBL.PutValue ('GF_PRIXUNITAIRE', 0);
TOBL.PutValue ('GF_REMISE', 0);
TOBL.PutValue ('GF_CALCULREMISE', '');
TOBL.PutValue ('GF_DATEDEBUT', V_PGI.DateEntree);
TOBL.PutValue ('GF_DATEFIN', IDate2099);
TOBL.PutValue ('GF_MODECREATION', 'MAN');
if GF_CASCADEREMISE.Value <> '' then
    TOBL.PutValue ('GF_CASCADEREMISE', GF_CASCADEREMISE.Value)
    else TOBL.PutValue ('GF_CASCADEREMISE', 'MIE');
TOBL.PutValue ('GF_DEVISE', GF_DEVISE.Value);
TOBL.PutValue ('GF_QUALIFPRIX', 'GRP');
TOBL.PutValue ('GF_FERME', '-');
TOBL.PutValue ('GF_SOCIETE', V_PGI.CodeSociete) ;
if TarifTTC then TOBL.PutValue ('GF_REGIMEPRIX', 'TTC')
            else TOBL.PutValue ('GF_REGIMEPRIX', 'HT') ;
Case ttd of
    ttdArt    : BEGIN
                TOBL.PutValue ('GF_BORNEINF', -999999);
                TOBL.PutValue ('GF_QUANTITATIF', '-');
                END;
    ttdArtQte : BEGIN
                TOBL.PutValue ('GF_BORNEINF', 1);
                TOBL.PutValue ('GF_QUANTITATIF', 'X');
                END;
    ttdArtQCa : BEGIN
                TOBL.PutValue ('GF_BORNEINF', 1);
                TOBL.PutValue ('GF_QUANTITATIF', 'X');
                END;
    ttdArtCat : BEGIN
                TOBL.PutValue ('GF_BORNEINF', -999999);
                TOBL.PutValue ('GF_QUANTITATIF', '-');
                END;
    END;

AfficheLAligne (ARow, ttd);
END;

Function TFTarifArticle.GetTOBLigne ( ARow : integer;  ttd : T_TableTarif) : TOB ;
BEGIN
Result:=Nil ;
case ttd of
    ttdArt    : BEGIN
                if ((ARow<=0) or (ARow>TOBTarfArt.Detail.Count)) then Exit ;
                Result:=TOBTarfArt.Detail[ARow-1] ;
                END;
    ttdArtQte : BEGIN
                if ((ARow<=0) or (ARow>TOBTarfQte.Detail.Count)) then Exit ;
                Result:=TOBTarfQte.Detail[ARow-1] ;
                END;
    ttdArtQCa : BEGIN
                if ((ARow<=0) or (ARow>TOBTarfQCa.Detail.Count)) then Exit ;
                Result:=TOBTarfQCa.Detail[ARow-1] ;
                END;
    ttdArtCat : BEGIN
                if ((ARow<=0) or (ARow>TOBTarfCat.Detail.Count)) then Exit ;
                Result:=TOBTarfCat.Detail[ARow-1] ;
                END;
    END;
END ;

procedure TFTarifArticle.InitialiseGrille ;
BEGIN
G_Art.VidePile(False) ;
G_Art.RowCount:= NbRowsInit ;
G_Qte.VidePile(False) ;
G_Qte.RowCount:= NbRowsInit ;
G_QCa.VidePile(False) ;
G_QCa.RowCount:= NbRowsInit ;
G_Cat.VidePile(False) ;
G_Cat.RowCount:= NbRowsInit ;
END;

procedure TFTarifArticle.InitialiseLigne (ARow : integer; ttd : T_TableTarif) ;
Var TOBL : TOB ;
    i_ind : integer ;
BEGIN
TOBL:=GetTOBLigne(ARow, ttd) ; if TOBL<>Nil then TOBL.InitValeurs ;
case ttd of
    ttdArt    : BEGIN
                for i_ind := 1 to G_Art.ColCount-1 do
                    BEGIN
                    G_Art.Cells [i_ind, ARow]:='' ;
                    END;
                END;
    ttdArtQte : BEGIN
                for i_ind := 1 to G_Qte.ColCount-1 do
                    BEGIN
                    G_Qte.Cells [i_ind, ARow]:='' ;
                    END;
                END;
    ttdArtQCa : BEGIN
                for i_ind := 1 to G_QCa.ColCount-1 do
                    BEGIN
                    G_QCa.Cells [i_ind, ARow]:='' ;
                    END;
                END;
    ttdArtCat : BEGIN
                for i_ind := 1 to G_Cat.ColCount-1 do
                    BEGIN
                    G_Cat.Cells [i_ind, ARow]:='' ;
                    END;
                END;
    END;
END;

{$IFDEF MONCODE}
procedure TFTarifArticle.AffichePrix (TOBL : TOB ; ARow : integer; ttd : T_TableTarif) ;
var PuFix : double;
begin
if (TOBArt.GetValue('GA_DECIMALPRIX')='X') then
  begin
  PuFix:=TOBL.GetValue('GF_PRIXUNITAIRE');
  if PrixPourQte<=0 then PrixPourQte:=1;
  PuFix:=PuFix/PrixPourQte;
Case ttd of
  ttdArt    : G_Art.cells[SA_Px,Arow] := strf00 (PuFix,DEV.Decimale+NbreDecimalSuppPrix('CLI',TOBArt));
  ttdArtQte : G_QTE.cells[SG_Px,Arow] := strf00 (PuFix,DEV.Decimale+NbreDecimalSuppPrix('CLI',TOBArt));
  ttdArtQCa : G_QCa.cells[SC_Px,Arow] := strf00 (PuFix,DEV.Decimale+NbreDecimalSuppPrix('CLI',TOBArt));
  ttdArtCat : G_Cat.cells[SC2_Px,Arow] := strf00 (PuFix,DEV.Decimale+NbreDecimalSuppPrix('CLI',TOBArt));
  END;
  end;
end;
{$ENDIF}

procedure TFTarifArticle.AfficheLaLigne (ARow : integer;  ttd : T_TableTarif) ;
Var TOBL : TOB ;
    i_ind,i_ind2 : integer ;
    nbDec : string;
BEGIN
TOBL:=GetTOBLigne(ARow, ttd) ; if TOBL = Nil then exit;
nbDec := '0.';
Case ttd of
    ttdArt    : BEGIN
                for i_ind := 1 to G_Art.ColCount-1 do
                    BEGIN
                    if ((i_ind=SA_Datedeb) or (i_ind=SA_Datefin)) then
                        G_Art.ColFormats [i_ind] := '';
                    if i_ind=SA_Px then
                       begin
                      {$IFDEF MONCODE}
                       for i_ind2 := 1 to DEV.Decimale+NbreDecimalSuppPrix('CLI',TOBArt) do nbDec:=nbDec + '0';
                      {$ELSE}
                       for i_ind2 := 1 to V_PGI.OkDecP do nbDec:=nbDec + '0';
                      {$ENDIF}
                       G_Art.ColFormats [i_ind] := '# ##' + nbDec;
                       end;
                    END;
                TOBL.PutLigneGrid(G_Art,ARow,False,False,LesColArt) ;
                {$IFDEF MONCODE}
                AffichePrix(TOBL,ARow,ttd) ;
                {$ENDIF}
                if G_Art.Cells [SA_Datedeb, ARow] = '' then
                    G_Art.Cells [SA_Datedeb, ARow] := DateToStr (V_PGI.DateEntree);
                for i_ind := 1 to G_Art.ColCount-1 do
                    BEGIN
                    FormateZoneSaisie(i_ind, ARow, ttd) ;
                    END;
                END;
    ttdArtQte : BEGIN
                for i_ind := 1 to G_Qte.ColCount-1 do
                    BEGIN
                    if ((i_ind=SG_Datedeb) or (i_ind=SG_Datefin)) then
                        G_QTE.ColFormats [i_ind] := '';
                    if i_ind=SG_Px then
                       begin
                      {$IFDEF MONCODE}
                       for i_ind2 := 1 to DEV.Decimale+NbreDecimalSuppPrix('CLI',TOBArt) do nbDec:=nbDec + '0';
                      {$ELSE}
                       for i_ind2 := 1 to V_PGI.OkDecP do nbDec:=nbDec + '0';
                      {$ENDIF}
                       G_QTE.ColFormats [i_ind] := '# ##' + nbDec;
                       end;
                    END;
                TOBL.PutLigneGrid(G_Qte,ARow,False,False,LesColQtes) ;
                {$IFDEF MONCODE}
                AffichePrix(TOBL,ARow,ttd) ;
                {$ENDIF}
                if G_Qte.Cells [SG_Datedeb, ARow] = '' then
                    G_Qte.Cells [SG_Datedeb, ARow] := DateToStr (V_PGI.DateEntree);
                for i_ind := 1 to G_Qte.ColCount-1 do
                    BEGIN
                    FormateZoneSaisie(i_ind, ARow, ttd) ;
                    END;
                END;
    ttdArtQCa : BEGIN
                for i_ind := 1 to G_QCa.ColCount-1 do
                    BEGIN
                    if ((i_ind=SC_Datedeb) or (i_ind=SC_Datefin)) then
                        G_QCa.ColFormats [i_ind] := '';
                    if i_ind=SC_Px then
                       begin
                      {$IFDEF MONCODE}
                       for i_ind2 := 1 to DEV.Decimale+NbreDecimalSuppPrix('CLI',TOBArt) do nbDec:=nbDec + '0';
                      {$ELSE}
                       for i_ind2 := 1 to V_PGI.OkDecP do nbDec:=nbDec + '0';
                      {$ENDIF}
                       G_QCa.ColFormats [i_ind] := '# ##' + nbDec;
                       end;
                    END;
                TOBL.PutLigneGrid(G_QCa,ARow,False,False,LesColQCa) ;
                {$IFDEF MONCODE}
                AffichePrix(TOBL,ARow,ttd) ;
                {$ENDIF}
                if G_QCa.Cells [SC_Datedeb, ARow] = '' then
                    G_QCa.Cells [SC_Datedeb, ARow] := DateToStr (V_PGI.DateEntree);
                for i_ind := 1 to G_QCa.ColCount-1 do
                    BEGIN
                    FormateZoneSaisie(i_ind, ARow, ttd) ;
                    END;
                END;
    ttdArtCat : BEGIN
                for i_ind := 1 to G_Cat.ColCount-1 do
                    BEGIN
                    if ((i_ind=SC2_Datedeb) or (i_ind=SC2_Datefin)) then
                        G_Cat.ColFormats [i_ind] := '';
                    if i_ind=SC2_Px then
                       begin
                      {$IFDEF MONCODE}
                       for i_ind2 := 1 to DEV.Decimale+NbreDecimalSuppPrix('CLI',TOBArt) do nbDec:=nbDec + '0';
                      {$ELSE}
                       for i_ind2 := 1 to V_PGI.OkDecP do nbDec:=nbDec + '0';
                      {$ENDIF}
                       G_Cat.ColFormats [i_ind] := '# ##' + nbDec;
                       end;
                    END;
                TOBL.PutLigneGrid(G_Cat,ARow,False,False,LesColCat) ;
                {$IFDEF MONCODE}
                AffichePrix(TOBL,ARow,ttd) ;
                {$ENDIF}
                if G_Cat.Cells [SC2_Datedeb, ARow] = '' then
                    G_Cat.Cells [SC2_Datedeb, ARow] := DateToStr (V_PGI.DateEntree);
                for i_ind := 1 to G_Cat.ColCount-1 do
                    BEGIN
                    FormateZoneSaisie(i_ind, ARow, ttd) ;
                    END;
                END;
    END;
END ;

Procedure TFTarifArticle.DepileTOBLigne ;
var i_ind : integer;
BEGIN
for i_ind := TOBTarfArt.Detail.Count - 1 Downto 0 do
    BEGIN
    TOBTarfArt.Detail[i_ind].Free ;
    END;
for i_ind := TOBTarfQte.Detail.Count - 1 Downto 0 do
    BEGIN
    TOBTarfQte.Detail[i_ind].Free ;
    END;
for i_ind := TOBTarfQCa.Detail.Count - 1 Downto 0 do
    BEGIN
    TOBTarfQCa.Detail[i_ind].Free ;
    END;
for i_ind := TOBTarfCat.Detail.Count - 1 Downto 0 do
    BEGIN
    TOBTarfCat.Detail[i_ind].Free ;
    END;
for i_ind := TOBTarifDel.Detail.Count - 1 Downto 0 do
    BEGIN
    TOBTarifDel.Detail[i_ind].Free ;
    END;
END;

Procedure TFTarifArticle.CreerTOBLigne (ARow : integer; ttd : T_TableTarif);
BEGIN
Case ttd of
    ttdArt    : BEGIN
                if ARow <> TOBTarfArt.Detail.Count + 1 then exit;
                TOB.Create ('TARIF', TOBTarfArt, ARow-1) ;
                InitialiseLigne (ARow, ttd) ;
                END;
    ttdArtQte : BEGIN
                if ARow <> TOBTarfQte.Detail.Count + 1 then exit;
                TOB.Create ('TARIF', TOBTarfQte, ARow-1) ;
                InitialiseLigne (ARow, ttd) ;
                END;
    ttdArtQCa : BEGIN
                if ARow <> TOBTarfQCa.Detail.Count + 1 then exit;
                TOB.Create ('TARIF', TOBTarfQCa, ARow-1) ;
                InitialiseLigne (ARow, ttd) ;
                END;
    ttdArtCat : BEGIN
                if ARow <> TOBTarfCat.Detail.Count + 1 then exit;
                TOB.Create ('TARIF', TOBTarfCat, ARow-1) ;
                InitialiseLigne (ARow, ttd) ;
                END;
    END ;
END;

Function TFTarifArticle.LigneVide (ARow : integer; ttd : T_TableTarif) : Boolean;
BEGIN
Result := True;
Case ttd of
    ttdArt    : BEGIN
                if G_Art.Cells [SG_Lib, ARow] <> '' then
                    BEGIN
                    Result := False;
                    Exit;
                    END;
                END;
    ttdArtQte : BEGIN
                if G_Qte.Cells [SG_Lib, ARow] <> '' then
                    BEGIN
                    Result := False;
                    Exit;
                    END;
                END;
    ttdArtQCa : BEGIN
                if (G_QCa.Cells [SC_Lib, ARow] <> '') And (G_QCa.Cells [SC_Cat, ARow] <> '') then
                    BEGIN
                    Result := False;
                    Exit;
                    END;
                END;
    ttdArtCat : BEGIN
                if (G_Cat.Cells [SC2_Lib, ARow] <> '') And (G_Cat.Cells [SC2_Cat, ARow] <> '') then
                    BEGIN
                    Result := False;
                    Exit;
                    END;
                END;
    END ;
END;

Procedure TFTarifArticle.PreAffecteLigne (ARow : integer; ttd : T_TableTarif);
var TOBL, TOBLPrec : TOB;
BEGIN
TOBLPrec := GetTOBLigne (ARow - 1, ttd); if TOBLPrec = nil then exit;
if TOBLPrec.GetValue ('GF_BORNESUP') < 999999 then
    BEGIN
    Case ttd of
        ttdArt    : TOBTarfArt.Detail[ARow - 1].Dupliquer (TOBLPrec, False, True);
        ttdArtQte : TOBTarfQte.Detail[ARow - 1].Dupliquer (TOBLPrec, False, True);
        ttdArtQCa : TOBTarfQCa.Detail[ARow - 1].Dupliquer (TOBLPrec, False, True);
        ttdArtCat : TOBTarfCat.Detail[ARow - 1].Dupliquer (TOBLPrec, False, True);
        END;
    TOBL := GetTOBLigne (ARow, ttd);
    TOBL.PutValue ('GF_BORNEINF', TOBLPrec.GetValue ('GF_BORNESUP') + 1);
    TOBL.PutValue ('GF_BORNESUP', 999999);
    TOBL.PutValue ('GF_TARIF', 0);
    AfficheLAligne (ARow, ttd);
    END;
END;

{==============================================================================================}
{===================== Manipulation des Champs LIGNES Quantitatif =============================}
{==============================================================================================}
procedure TFTarifArticle.TraiterDepot (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    St : string;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
Case ttd of
    ttdArt    : St := G_Art.Cells [ACol, ARow];
    ttdArtQte : St := G_Qte.Cells [ACol, ARow];
    ttdArtQCa : St := G_QCa.Cells [ACol, ARow];
    ttdArtCat : St := G_Cat.Cells [ACol, ARow];
    END ;
if ExisteDepot ('GCDEPOT', St) then
    BEGIN
    TOBL.PutValue ('GF_DEPOT', St);
    END else
    BEGIN
    // message dépôt inexistant
    MsgBox.Execute (4,Caption,'') ;
    Case ttd of
        ttdArt    : G_Art.Cells [ACol, ARow] := TOBL.GetValue ('GF_DEPOT');
        ttdArtQte : G_Qte.Cells [ACol, ARow] := TOBL.GetValue ('GF_DEPOT');
        ttdArtQCa : G_QCa.Cells [ACol, ARow] := TOBL.GetValue ('GF_DEPOT');
        ttdArtCat : G_Cat.Cells [ACol, ARow] := TOBL.GetValue ('GF_DEPOT');
        END ;
    END;
END;

procedure TFTarifArticle.TraiterCatTiers (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    B_NewLine : Boolean;
    St : string;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
if TOBL.GetValue ('GF_TARIFTIERS') = '' then B_NewLine :=True else B_NewLine := False;
Case ttd of
    ttdArtQCa : St := G_QCa.Cells [ACol, ARow];
    ttdArtCat : St := G_Cat.Cells [ACol, ARow];
    END;
if St <> '' then
    BEGIN
    TOBL.PutValue ('GF_TARIFTIERS', St);
    if ExisteCategorie ('TTTARIFCLIENT', St) then
        BEGIN
        TOBL.PutValue ('GF_TARIFTIERS', St);
        END else
        BEGIN
        // message Catégorie Tiers inexistante
        MsgBox.Execute (5,Caption,'') ;
        Case ttd of
            ttdArtQCa : G_QCa.Cells [ACol, ARow] := TOBL.GetValue ('GF_TARIFTIERS');
            ttdArtCat : G_Cat.Cells [ACol, ARow] := TOBL.GetValue ('GF_TARIFTIERS');
            END;
        END;
    if (TOBL.GetValue ('GF_LIBELLE') <> '') AND (TOBL.GetValue ('GF_TARIFTIERS') <> '') AND
       (B_NewLine) then
        BEGIN
        InitLaLigne (ARow, ttd);
        END;
    END else if Not B_Newline then
        BEGIN
        Case ttd of
            ttdArtQCa : G_QCa.Cells [ACol, ARow] := TOBL.GetValue ('GF_TARIFTIERS');
            ttdArtCat : G_Cat.Cells [ACol, ARow] := TOBL.GetValue ('GF_TARIFTIERS');
            END;
        END;
END;

procedure TFTarifArticle.TraiterLibelle (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    B_NewLine : Boolean;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
if TOBL.GetValue ('GF_LIBELLE') = '' then B_NewLine :=True else B_NewLine := False;
Case ttd of
    ttdArt    : BEGIN
                if G_Art.Cells [ACol, ARow] <> '' then
                    BEGIN
                    TOBL.PutValue ('GF_LIBELLE', G_Art.Cells [ACol, ARow]);
                    if (TOBL.GetValue ('GF_LIBELLE') <> '') AND (B_NewLine) then
                        BEGIN
                        InitLaLigne (ARow, ttd);
                        END;
                    END else if Not B_Newline then
                        G_Art.Cells [ACol, ARow] := TOBL.GetValue ('GF_LIBELLE');
                 END;
    ttdArtQte : BEGIN
                if G_Qte.Cells [ACol, ARow] <> '' then
                    BEGIN
                    TOBL.PutValue ('GF_LIBELLE', G_Qte.Cells [ACol, ARow]);
                    if (TOBL.GetValue ('GF_LIBELLE') <> '') AND (B_NewLine) then
                        BEGIN
                        InitLaLigne (ARow, ttd);
                        END;
                    END else if Not B_Newline then
                        G_Qte.Cells [ACol, ARow] := TOBL.GetValue ('GF_LIBELLE');
                 END;
    ttdArtQCa : BEGIN
                if G_QCa.Cells [ACol, ARow] <> '' then
                    BEGIN
                    TOBL.PutValue ('GF_LIBELLE', G_QCa.Cells [ACol, ARow]);
                    if (TOBL.GetValue ('GF_LIBELLE') <> '') AND
                       (TOBL.GetValue ('GF_TARIFTIERS') <> '') AND
                       (B_NewLine) then
                        BEGIN
                        InitLaLigne (ARow, ttd);
                        END;
                    END else if Not B_Newline then
                        G_QCa.Cells [ACol, ARow] := TOBL.GetValue ('GF_LIBELLE');
                END;
    ttdArtCat : BEGIN
                if G_Cat.Cells [ACol, ARow] <> '' then
                    BEGIN
                    TOBL.PutValue ('GF_LIBELLE', G_Cat.Cells [ACol, ARow]);
                    if (TOBL.GetValue ('GF_LIBELLE') <> '') AND
                       (TOBL.GetValue ('GF_TARIFTIERS') <> '') AND
                       (B_NewLine) then
                        BEGIN
                        InitLaLigne (ARow, ttd);
                        END;
                    END else if Not B_Newline then
                        G_Cat.Cells [ACol, ARow] := TOBL.GetValue ('GF_LIBELLE');
                END;
    END ;
END;

procedure TFTarifArticle.TraiterBorneInf (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    f_QteInf : Extended;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
f_QteInf := 0;
Case ttd of
    ttdArtQte : f_QteInf := Valeur (G_Qte.Cells [ACol, ARow]);
    ttdArtQCa : f_QteInf := Valeur (G_QCa.Cells [ACol, ARow]);
    END ;
if f_QteInf > TOBL.GetValue ('GF_BORNESUP') then
    BEGIN
    TOBL.PutValue ('GF_BORNESUP', f_QteInf);
    Case ttd of
        ttdArtQte : G_Qte.Cells [SG_QSup, ARow] := floattostr (TOBL.GetValue ('GF_BORNESUP'));
        ttdArtQCa : G_QCa.Cells [SC_QSup, ARow] := floattostr (TOBL.GetValue ('GF_BORNESUP'));
        END ;
    END;
if f_QteInf < 1 then
    BEGIN
    MsgBox.Execute (2,Caption,'') ;
    Case ttd of
        ttdArtQte : BEGIN
                    G_Qte.Cells [ACol, ARow] := floattostr (TOBL.GetValue ('GF_BORNEINF'));
                    G_Qte.Col := ACol; G_Qte.Row := ARow;
                    END;
        ttdArtQCa : BEGIN
                    G_QCa.Cells [ACol, ARow] := floattostr (TOBL.GetValue ('GF_BORNEINF'));
                    G_QCa.Col := ACol; G_QCa.Row := ARow;
                    END;
        END ;
    END else TOBL.PutValue ('GF_BORNEINF', f_QteInf);
END;

procedure TFTarifArticle.TraiterBorneSup (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    f_QteSup : Extended;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
f_QteSup := 999999;
Case ttd of
    ttdArtQte : f_QteSup := Valeur (G_Qte.Cells [ACol, ARow]);
    ttdArtQCa : f_QteSup := Valeur (G_QCa.Cells [ACol, ARow]);
    END ;
if f_QteSup < TOBL.GetValue ('GF_BORNEINF') then
    BEGIN
    MsgBox.Execute (2,Caption,'') ;
    Case ttd of
        ttdArtQte : BEGIN
                    G_Qte.Cells [ACol, ARow] := floattostr (TOBL.GetValue ('GF_BORNESUP'));
                    G_Qte.Col := ACol; G_Qte.Row := ARow;
                    END;
        ttdArtQCa : BEGIN
                    G_QCa.Cells [ACol, ARow] := floattostr (TOBL.GetValue ('GF_BORNESUP'));
                    G_QCa.Col := ACol; G_QCa.Row := ARow;
                    END;
        END ;
    END else
    BEGIN
    TOBL.PutValue ('GF_BORNESUP', f_QteSup);
    END;
END;

procedure TFTarifArticle.TraiterPrix (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    PuFix : double;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
Case ttd of
    ttdArt    : PuFix := Valeur (G_Art.Cells [ACol, ARow]);
    ttdArtQte : PuFix := Valeur (G_Qte.Cells [ACol, ARow]);
    ttdArtQCa : PuFix := Valeur (G_QCa.Cells [ACol, ARow]);
    ttdArtCat : PuFix := Valeur (G_Cat.Cells [ACol, ARow]);
    END ;
{$IFDEF MONCODE}
if TOBArt.GetValue('GA_DECIMALPRIX')='X' then
  begin
  if PrixPourQte<=0 then PrixPourQte:=1;
  PuFix:=PuFix*PrixPourQte;
  end;
{$ENDIF}
TOBL.PutValue ('GF_PRIXUNITAIRE',PuFix);
END;

procedure TFTarifArticle.TraiterRemise(ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    St : string;
BEGIN
TOBL := GetTOBLigne (ARow, ttd);
Case ttd of
    ttdArt    :  begin
                 G_Art.Cells [ACol, ARow]:=ModifFormat(G_Art.Cells [ACol, ARow]);
                 St := G_Art.Cells [ACol, ARow];
                 end;
    ttdArtQte :  begin
                 G_Qte.Cells [ACol, ARow]:=ModifFormat(G_Qte.Cells [ACol, ARow]);
                 St := G_Qte.Cells [ACol, ARow];
                 end;
    ttdArtQCa :  begin
                 G_QCa.Cells [ACol, ARow]:=ModifFormat(G_QCa.Cells [ACol, ARow]);
                 St := G_QCa.Cells [ACol, ARow];
                 end;
    ttdArtCat :  begin
                 G_Cat.Cells [ACol, ARow]:=ModifFormat(G_Cat.Cells [ACol, ARow]);
                 St := G_Cat.Cells [ACol, ARow];
                 end;
    END ;
TOBL.PutValue ('GF_CALCULREMISE', St);
TOBL.PutValue ('GF_REMISE', RemiseResultante (St));
AffichePied (ttd);
END;

procedure TFTarifArticle.TraiterDateDeb (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    St_Date : string;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
Case ttd of
    ttdArt    : St_Date := G_Art.Cells [ACol, ARow] ;
    ttdArtQte : St_Date := G_Qte.Cells [ACol, ARow] ;
    ttdArtQCa : St_Date := G_QCa.Cells [ACol, ARow] ;
    ttdArtCat : St_Date := G_Cat.Cells [ACol, ARow] ;
    END ;
if IsValidDate (st_Date) then
    BEGIN
    if StrToDate (St_Date) > TOBL.GetValue ('GF_DATEFIN') then
        BEGIN
            MsgBox.Execute (3,Caption,'') ;
            Case ttd of
                ttdArt    : BEGIN
                            G_Art.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
                            G_Art.Col := ACol; G_Art.Row := ARow;
                            END;
                ttdArtQte : BEGIN
                            G_Qte.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
                            G_Qte.Col := ACol; G_Qte.Row := ARow;
                            END;
                ttdArtQCa : BEGIN
                            G_QCa.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
                            G_QCa.Col := ACol; G_QCa.Row := ARow;
                            END;
                ttdArtCat : BEGIN
                            G_Cat.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
                            G_Cat.Col := ACol; G_Cat.Row := ARow;
                            END;
                END ;
        END else
        BEGIN
        TOBL.PutValue ('GF_DATEDEBUT', StrToDate (St_Date));
        END;
    END else
    BEGIN
    if TOBL.GetValue ('GF_LIBELLE') <> '' then
        BEGIN
        Case ttd of
            ttdArt    : G_Art.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
            ttdArtQte : G_Qte.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
            ttdArtQCa : G_QCa.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
            ttdArtCat : G_Cat.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
            END ;
        END;
    END;
END;

procedure TFTarifArticle.TraiterDateFin (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    St_Date : string;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
Case ttd of
    ttdArt    : St_Date := G_Art.Cells [ACol, ARow] ;
    ttdArtQte : St_Date := G_Qte.Cells [ACol, ARow] ;
    ttdArtQCa : St_Date := G_QCa.Cells [ACol, ARow] ;
    ttdArtCat : St_Date := G_Cat.Cells [ACol, ARow] ;
    END ;
if IsValidDate (st_Date) then
    BEGIN
    if StrToDate (St_Date) < TOBL.GetValue ('GF_DATEDEBUT') then
        BEGIN
            MsgBox.Execute (3,Caption,'') ;
            Case ttd of
                ttdArt    : BEGIN
                            G_Art.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
                            G_Art.Col := ACol; G_Art.Row := ARow;
                            END;
                ttdArtQte : BEGIN
                            G_Qte.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
                            G_Qte.Col := ACol; G_Qte.Row := ARow;
                            END;
                ttdArtQCa : BEGIN
                            G_QCa.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
                            G_QCa.Col := ACol; G_QCa.Row := ARow;
                            END;
                ttdArtCat : BEGIN
                            G_Cat.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
                            G_Cat.Col := ACol; G_Cat.Row := ARow;
                            END;
                END ;
        END else
        BEGIN
        TOBL.PutValue ('GF_DATEFIN', StrToDate (St_Date));
        END;
    END else
    BEGIN
    if TOBL.GetValue ('GF_LIBELLE') <> '' then
        BEGIN
        Case ttd of
            ttdArt    : G_Art.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
            ttdArtQte : G_Qte.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
            ttdArtQCa : G_QCa.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
            ttdArtCat : G_Cat.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
            END ;
        END;
    END;
END;

{==============================================================================================}
{========================= Evenement de l'Entete  =============================================}
{==============================================================================================}

procedure TFTarifArticle.GF_CODEARTICLEExit(Sender: TObject);
BEGIN
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if GF_CODEARTICLE.Text <> TOBArt.GetValue ('GA_CODEARTICLE') then TraiterArticle ;
END;


procedure TFTarifArticle.GF_CODEARTICLEElipsisClick(Sender: TObject);
begin
ChercherArticle;
end;

procedure TFTarifArticle.GF_DEVISEChange(Sender: TObject);
var i_Rep : integer;
    ioerr : TIOErr ;
begin
if CodeDevise = GF_DEVISE.Value then Exit;
i_Rep := QuestionTarifEnCours;
Case i_Rep of
    mrYes    : BEGIN
               ioerr := Transactions (ValideTarif, 2);;
               Case ioerr of
                    oeOk      : ;
                    oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
                    oeSaisie  : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
                    END ;
               CodeDevise := GF_DEVISE.Value;
               DEV.Code := GF_DEVISE.Value ; GetInfosDevise (DEV) ;
               Transactions (ChargeTarif, 1) ;
               END;
    mrNo     : BEGIN
               InitialiseGrille;
               DepileTOBLigne;
               CodeDevise := GF_DEVISE.Value;
               DEV.Code := GF_DEVISE.Value ; GetInfosDevise (DEV) ;
               Transactions (ChargeTarif, 1) ;
               END;
    mrCancel : GF_DEVISE.Value := CodeDevise;
    END ;
InitialisePied ;
end;

procedure TFTarifArticle.CBQUANTITAIFClick(Sender: TObject);
begin
InitialisePied ;
if CBQUANTITAIF.Checked then
    BEGIN
    if CBCATTIERS.Checked then
        BEGIN
        PART.Visible := False;
        PQUANTITATIF.Visible := False;
        PQTECAT.Visible := True;
        PCATEGORIE.Visible := False;
        PTITRE.Caption := HTitre.Mess[1] ;
        AffectMenuCondApplic (G_QCa, ttdArtQCa);
        END else
        BEGIN
        PART.Visible := False;
        PQUANTITATIF.Visible := True;
        PQTECAT.Visible := False;
        PCATEGORIE.Visible := False;
        PTITRE.Caption := HTitre.Mess[0] ;
        AffectMenuCondApplic (G_Qte, ttdArtQte);
        END;
    END else
    BEGIN
    if CBCATTIERS.Checked then
        BEGIN
        PART.Visible := False;
        PQUANTITATIF.Visible := False;
        PQTECAT.Visible := False;
        PCATEGORIE.Visible := True;
        PTITRE.Caption := HTitre.Mess[2] ;
        AffectMenuCondApplic (G_Cat, ttdArtCat);
        END else
        BEGIN
        PART.Visible := True;
        PQUANTITATIF.Visible := False;
        PQTECAT.Visible := False;
        PCATEGORIE.Visible := False;
        PTITRE.Caption := HTitre.Mess[3] ;
        AffectMenuCondApplic (G_Art, ttdArt);
        END;
    END;
HMTrad.ResizeGridColumns (G_Art) ;
HMTrad.ResizeGridColumns (G_Qte) ;
HMTrad.ResizeGridColumns (G_QCa) ;
HMTrad.ResizeGridColumns (G_Cat) ;
end;

procedure TFTarifArticle.CBCATTIERSClick(Sender: TObject);
begin
InitialisePied ;
if CBCATTIERS.Checked then
    BEGIN
    if CBQUANTITAIF.Checked then
        BEGIN
        PART.Visible := FAlse;
        PQUANTITATIF.Visible := FAlse;
        PQTECAT.Visible := True;
        PCATEGORIE.Visible := False;
        PTITRE.Caption := HTitre.Mess[1] ;
        AffectMenuCondApplic (G_QCa, ttdArtQCa);
        END else
        BEGIN
        PART.Visible := FAlse;
        PQUANTITATIF.Visible := False;
        PQTECAT.Visible := False;
        PCATEGORIE.Visible := True;
        PTITRE.Caption := HTitre.Mess[2] ;
        AffectMenuCondApplic (G_Cat, ttdArtCat);
        END;
    END else
    BEGIN
    if CBQUANTITAIF.Checked then
        BEGIN
        PART.Visible := FAlse;
        PQUANTITATIF.Visible := True;
        PQTECAT.Visible := False;
        PCATEGORIE.Visible := False;
        PTITRE.Caption := HTitre.Mess[0];
        AffectMenuCondApplic (G_Qte, ttdArtQte);
        END else
        BEGIN
        PART.Visible := True;
        PQUANTITATIF.Visible := False;
        PQTECAT.Visible := False;
        PCATEGORIE.Visible := False;
        PTITRE.Caption := HTitre.Mess[3] ;
        AffectMenuCondApplic (G_Art, ttdArt);
        END;
    END;
HMTrad.ResizeGridColumns (G_Art) ;
HMTrad.ResizeGridColumns (G_Qte) ;
HMTrad.ResizeGridColumns (G_QCa) ;
HMTrad.ResizeGridColumns (G_Cat) ;
end;

{==============================================================================================}
{========================= Manipulation de l'entête ===========================================}
{==============================================================================================}
Procedure TFTarifArticle.PrepareEntete ;
BEGIN
GF_DEVISE.Value := CodeDevise;
TOBArt.InitValeurs ;
TOBArt.SelectDB ('"' + CodeArticle + '"', Nil) ;
GF_CODEARTICLE.Text := TOBArt.GetValue ('GA_CODEARTICLE') ;
if GF_CODEARTICLE.Text <> '' then
    BEGIN
    AffecteEntete;
    G_Art.Enabled := True ; G_Qte.Enabled := True ; G_QCa.Enabled := True ; G_Cat.Enabled := True ;
    InitialiseGrille;
    DepileTOBLigne;
    CodeArticle := TOBArt.GetValue ('GA_ARTICLE');
    Transactions (ChargeTarif, 1) ;
    END else
    BEGIN
    TOBArt.InitValeurs ;
    AffecteEntete;
    G_Art.Enabled := False ; G_Qte.Enabled := False ; G_QCa.Enabled := False ; G_Cat.Enabled := False ;
    END;
GF_CODEARTICLE.Enabled:=False ; GF_CASCADEREMISE.Enabled:=False; GF_CASCADEREMISE.Style:=csSimple;
G_Art.SetFocus ;
BCondAplli.Enabled := False ; BCopierCond.Enabled := False; BCollerCond.Enabled := False ;
CONDAPPLIC.Text := '' ;
InitialisePied;
END;

Procedure TFTarifArticle.InitialiseEntete ;
BEGIN
GF_DEVISE.Value := CodeDevise;
TOBArt.InitValeurs ; PrixPourQte:=1;
AffecteEntete;
CodeArticle := '';
GF_CODEARTICLE.SetFocus ;
G_Art.Enabled := False ; G_Qte.Enabled := False ; G_QCa.Enabled := False ; G_Cat.Enabled := False ;
BCondAplli.Enabled := False ; BCopierCond.Enabled := False; BCollerCond.Enabled := False ;
CONDAPPLIC.Text := '' ;
GA_PRIXPOURQTE.Value:=0;
GA_PRIXPOURQTE.Visible:=False;
TGA_PRIXPOURQTE.Visible:=False;
InitialisePied;
END;

Procedure TFTarifArticle.AffecteEntete ;
BEGIN
GF_CODEARTICLE.Text := TOBArt.GetValue ('GA_CODEARTICLE');
TGF_LIBELLE.Caption := TOBArt.GetValue ('GA_LIBELLE');
// Dimensions
AffecteDimension;
END;


Procedure TFTarifArticle.AffecteDimension ;
var TLIB,TLIB2 : THLabel;
    CHPS,CHPS2 : THCritMaskEdit;
    i_ind, i_Dim : integer;
    b_Dim : Boolean ;
    GrilleDim,CodeDim,LibDim : String ;
BEGIN
i_Dim := 1; b_Dim := False;
for i_ind := 1 to MaxDimension do
    BEGIN
    TLIB := THLabel (FindComponent ('TGF_GRILLEDIM' + IntToStr (i_Dim)));
    CHPS := THCritMaskEdit (FindComponent ('GF_CODEDIM' + IntToStr (i_Dim)));
    TLIB2 := THLabel (FindComponent ('TGF_GRILLEDIM' + IntToStr (i_ind)));
    CHPS2 := THCritMaskEdit (FindComponent ('GF_CODEDIM' + IntToStr (i_ind)));
    TLIB2.Caption := '';
    CHPS2.Text := '';
    CHPS2.Visible := False;
    GrilleDim:=TOBArt.GetValue('GA_GRILLEDIM'+IntToStr(i_ind)) ;
    CodeDim:=TOBArt.GetValue('GA_CODEDIM'+IntToStr(i_ind)) ;
    if  GrilleDim<>'' then
        BEGIN
        TLIB.Caption:=RechDom('GCGRILLEDIM'+IntToStr(i_ind),GrilleDim,FALSE) ;
        LibDim:=GCGetCodeDim(GrilleDim,CodeDim,i_ind) ;
        if LibDim<>'' then
            BEGIN
            CHPS.Text:=LibDim ; CHPS.Visible:=True ; b_Dim:=True ; Inc(i_Dim) ;
            END else
            BEGIN
            TLIB.Caption:='' ; CHPS.Text:='' ; CHPS.Visible:=False ;
            END;
        END ;
    END;
if Not b_Dim then
    BEGIN
    PDIMENSION.Visible := False ;
    END else
    BEGIN
    PDIMENSION.Visible := True ;
    END;
END;

Procedure TFTarifArticle.ChercherArticle ;
BEGIN
  //mcd 06/05/03 ajout du test, car en GI/GA seule, il ne faut pas toutes les catégorie article
if (ctxAffaire in V_PGI.PGIContexte) and  not(ctxGCAFF in V_PGI.PGIContexte) then
  DispatchRechArt (GF_CODEARTICLE, 1, '',
                   'GA_CODEARTICLE=' + Trim (Copy (GF_CODEARTICLE.Text, 1, 18)), '')
else
  DispatchRechArt (GF_CODEARTICLE, 1, '',
                   'GA_CODEARTICLE=' + Trim (Copy (GF_CODEARTICLE.Text, 1, 18))+';XX_WHERE= AND GA_TYPEARTICLE<>"POU" AND GA_TYPEARTICLE <> "OUV" AND GA_TYPEARTICLE <> "EPO"', '');
if GF_CODEARTICLE.Text <> '' then
    BEGIN
    GF_CODEARTICLE.Text:=Format('%-33.33sX',[GF_CODEARTICLE.Text]);
    TOBArt.SelectDB ('"' + GF_CODEARTICLE.Text + '"', Nil) ;
    GF_CODEARTICLE.Text := TOBArt.GetValue ('GA_CODEARTICLE') ;
    TraiterArticle;
    END;
END;

Procedure TFTarifArticle.TraiterArticle ;
Var RechArt : T_RechArt ;
    OkArt   : Boolean ;
    i_Rep: Integer;
    ioerr : TIOErr ;
BEGIN
if GF_CODEARTICLE.Text = '' then
    BEGIN
    InitialiseEntete ;
    InitialiseGrille;
    DepileTOBLigne;
    exit
    END;

OkArt:=False ;
RechArt := TrouverArticle (GF_CODEARTICLE, TOBArt);
Case RechArt of
        traOk : OkArt:=True ;
     traAucun : BEGIN
                // Recherche sur code via LookUp ou Recherche avancée
                DispatchRechArt (GF_CODEARTICLE, 1, '',
                                   'GA_CODEARTICLE=' + Trim (Copy (GF_CODEARTICLE.Text, 1, 18)), '');
                if GF_CODEARTICLE.Text <> '' then
                    GF_CODEARTICLE.Text:=Format('%-33.33sX',[GF_CODEARTICLE.Text]);
                    BEGIN
                    Okart := TOBArt.SelectDB ('"' + GF_CODEARTICLE.Text + '"', nil);
                    END;
                END ;
    traGrille : BEGIN
                // Forcement objet dimension avec saisie obligatoire
                if ChoisirDimension (TOBArt.GetValue ('GA_ARTICLE'), TOBArt) then
                    BEGIN
                    Okart := True;
                    END;
                END;
     End ; // Case

if (Okart) then
    BEGIN
    AffecteEntete;
    G_Art.Enabled := True ; G_Qte.Enabled := True ; G_QCa.Enabled := True ; G_Cat.Enabled := True ;
    PrixPourQte := TOBArt.GetValue('GA_PRIXPOURQTE');
    GA_PRIXPOURQTE.Value:=PrixPourQte;
    if GA_PRIXPOURQTE.Value>1 then
       begin
       {$IFDEF MONCODE}
       if (TOBArt.GetValue('GA_DECIMALPRIX')='X') then
         begin
         GA_PRIXPOURQTE.Value:=DEV.Decimale+NbreDecimalSuppPrix('CLI',TOBArt);
         TGA_PRIXPOURQTE.Caption := HTitre.Mess[7] ;
         end;
       {$ENDIF}
       GA_PRIXPOURQTE.Visible:=True;
       TGA_PRIXPOURQTE.Visible:=True;
       end else
       begin
       GA_PRIXPOURQTE.Visible:=False;
       TGA_PRIXPOURQTE.Visible:=False;
       end;
    if CodeArticle <> TOBArt.GetValue ('GA_ARTICLE') then
        BEGIN
        i_Rep := QuestionTarifEnCours;
        Case i_Rep of
            mrYes    : BEGIN
                       ioerr := Transactions (ValideTarif, 2);;
                       Case ioerr of
                               oeOk : ;
                          oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
                           oeSaisie : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
                          END ;
                       CodeArticle := TOBArt.GetValue ('GA_ARTICLE');
                       Transactions (ChargeTarif, 1) ;
                       END;
            mrNo     : BEGIN
                       InitialiseGrille;
                       DepileTOBLigne;
                       CodeArticle := TOBArt.GetValue ('GA_ARTICLE');
                       Transactions (ChargeTarif, 1) ;
                       END;
            mrCancel : BEGIN
                       TOBArt.SelectDB ('"' + CodeArticle + '"', Nil) ;
                       AffecteEntete;
                       END;
            END ;
        END;
    END else
    BEGIN
    InitialiseEntete ;
    END;
END;

Function TFTarifArticle.QuestionTarifEnCours : Integer;
BEGIN
Result := mrNo;
if Action = taConsult then Exit;
if GrilleModifie then
    BEGIN
    Result := MsgBox.Execute (0, Caption, '');
    END;
END;

{==============================================================================================}
{=============================== Evenements du pied ===========================================}
{==============================================================================================}
procedure TFTarifArticle.GF_REMISEChange(Sender: TObject);
var St : string;
begin
St := GF_REMISE.Text;
St:=StrF00(Valeur(St),ADecimP);
GF_REMISE.Text := St;
end;

procedure TFTarifArticle.GF_CASCADEREMISEChange(Sender: TObject);
Var Row : integer;
    ttd : T_TableTarif;
    TOBL : TOB ;
begin
if PART.Visible=True then
    BEGIN
    Row := G_Art.Row;
    ttd := ttdArt;
    END else if PQUANTITATIF.Visible=True then
    BEGIN
    Row := G_Qte.Row;
    ttd := ttdArtQte;
    END else if PQTECAT.Visible=True then
    BEGIN
    Row := G_QCa.Row;
    ttd := ttdArtQCa;
    END else if PCATEGORIE.Visible=True then
    BEGIN
    Row := G_Cat.Row;
    ttd := ttdArtCat;
    END else exit;
if Row > 0 then
    BEGIN
    TOBL := GetTOBLigne (Row, ttd); if TOBL=nil then exit;
    TOBL.PutValue ('GF_CASCADEREMISE', GF_CASCADEREMISE.Value);
    END;

end;

{==============================================================================================}
{============================= Manipulation du pied ===========================================}
{==============================================================================================}
Procedure TFTarifArticle.InitialisePied ;
BEGIN
GF_CASCADEREMISE.Text := '';
GF_REMISE.Text := '';
GF_PRIXCON.Text := '';
if ((GF_DEVISE.Value<>V_PGI.DevisePivot) or (ForceEuroGC)) then
    BEGIN
    GF_PRIXCON.Visible:=False;
    TGF_PRIXCON.Visible:=False;
    ISigneEuro.Visible:=False;
    ISigneFranc.Visible:=False;
    END else
    BEGIN
    GF_PRIXCON.Visible:=True;
    TGF_PRIXCON.Visible:=True;
    if VH^.TenueEuro then
        BEGIN
        ISigneEuro.Visible:=False;
        ISigneFranc.Visible:=True;
        END else
        BEGIN
        ISigneEuro.Visible:=True;
        ISigneFranc.Visible:=False;
        END;
    END;
END;

Procedure TFTarifArticle.AffichePied (ttd : T_TableTarif) ;
var TOBL : TOB;
    FF   : TForm;
    ARow : Longint ;
BEGIN
if ttd = ttdArt then  ARow := G_Art.Row;
if ttd = ttdArtQte then  ARow := G_Qte.Row;
if ttd = ttdArtQCa then  ARow := G_QCa.Row;
if ttd = ttdArtCat then  ARow := G_Cat.Row;
TOBL:=GetTOBLigne(ARow, ttd) ; if TOBL=Nil then Exit ;
FF := TForm (PPIED.Owner);
TOBL.PutEcran (FF, PPIED);
END;

{==============================================================================================}
{============================ Evenement lié aux Boutons =======================================}
{==============================================================================================}
procedure TFTarifArticle.BChercherClick(Sender: TObject);
begin
if PART.Visible then
    if G_Art.RowCount < 3 then Exit;
if PQUANTITATIF.Visible then
    if G_Qte.RowCount < 3 then Exit;
if PQTECAT.Visible then
    if G_QCa.RowCount < 3 then Exit;
if PCATEGORIE.Visible then
    if G_Cat.RowCount < 3 then Exit;
FindDebut:=True ; FindLigne.Execute ;
end;

procedure TFTarifArticle.InfArticleClick(Sender: TObject);
begin
if GF_CODEARTICLE.Text = '' then Exit;
VoirFicheArticle;
end;


procedure TFTarifArticle.BVoirCondClick(Sender: TObject);
Var TOBL : TOB;
begin
TCONDTARF.Visible := BVoirCond.Down ;
end;

procedure TFTarifArticle.BCondAplliClick(Sender: TObject);
Var TOBL : TOB;
begin
{$IFNDEF CCS3}
if GF_CODEARTICLE.Text = '' then exit;
if PART.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_Art.Row, ttdArt) ; if TOBL=Nil then Exit ;
    EntreeTarifCond (Action, TOBL);
    AfficheCondTarf (G_Art.Row, ttdArt);
    END;
if PQUANTITATIF.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_Qte.Row, ttdArtQte) ; if TOBL=Nil then Exit ;
    EntreeTarifCond (Action, TOBL);
    AfficheCondTarf (G_Qte.Row, ttdArtQte);
    END;
if PQTECAT.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_QCa.Row, ttdArtQCa) ; if TOBL=Nil then Exit ;
    EntreeTarifCond (Action, TOBL);
    AfficheCondTarf (G_QCa.Row, ttdArtQCa);
    END;
if PCATEGORIE.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_Cat.Row, ttdArtCat) ; if TOBL=Nil then Exit ;
    EntreeTarifCond (Action, TOBL);
    AfficheCondTarf (G_Cat.Row, ttdArtCat);
    END;
{$ENDIF}
end;

procedure TFTarifArticle.BCopierCondClick(Sender: TObject);
Var TOBL : TOB;
begin
if GF_CODEARTICLE.Text = '' then exit;
if PART.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_Art.Row, ttdArt) ; if TOBL=Nil then Exit ;
    CONDAPPLIC.Text := TOBL.GetValue ('GF_CONDAPPLIC') ;
    END;
if PQUANTITATIF.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_Qte.Row, ttdArtQte) ; if TOBL=Nil then Exit ;
    CONDAPPLIC.Text := TOBL.GetValue ('GF_CONDAPPLIC') ;
    END;
if PQTECAT.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_QCa.Row, ttdArtQCa) ; if TOBL=Nil then Exit ;
    CONDAPPLIC.Text := TOBL.GetValue ('GF_CONDAPPLIC') ;
    END;
if PCATEGORIE.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_Cat.Row, ttdArtCat) ; if TOBL=Nil then Exit ;
    CONDAPPLIC.Text := TOBL.GetValue ('GF_CONDAPPLIC') ;
    END;
end;

procedure TFTarifArticle.BCollerCondClick(Sender: TObject);
Var TOBL : TOB;
begin
if GF_CODEARTICLE.Text = '' then exit;
if PART.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_Art.Row, ttdArt) ; if TOBL=Nil then Exit ;
    TOBL.PutValue ('GF_CONDAPPLIC', CONDAPPLIC.Text) ;
    AfficheCondTarf (G_Art.Row, ttdArt);
    END;
if PQUANTITATIF.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_Qte.Row, ttdArtQte) ; if TOBL=Nil then Exit ;
    TOBL.PutValue ('GF_CONDAPPLIC', CONDAPPLIC.Text) ;
    AfficheCondTarf (G_Qte.Row, ttdArtQte);
    END;
if PQTECAT.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_QCa.Row, ttdArtQCa) ; if TOBL=Nil then Exit ;
    TOBL.PutValue ('GF_CONDAPPLIC', CONDAPPLIC.Text) ;
    AfficheCondTarf (G_QCa.Row, ttdArtQCa);
    END;
if PCATEGORIE.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_Cat.Row, ttdArtCat) ; if TOBL=Nil then Exit ;
    TOBL.PutValue ('GF_CONDAPPLIC', CONDAPPLIC.Text) ;
    AfficheCondTarf (G_Cat.Row, ttdArtCat);
    END;
end;

{==============================================================================================}
{============================== Action lié aux Boutons ========================================}
{==============================================================================================}
procedure TFTarifArticle.FindLigneFind(Sender: TObject);
begin
if PART.Visible then
    Rechercher (G_Art, FindLigne, FindDebut) ;
if PQUANTITATIF.Visible then
    Rechercher (G_Qte, FindLigne, FindDebut) ;
if PQTECAT.Visible then
    Rechercher (G_QCa, FindLigne, FindDebut) ;
if PCATEGORIE.Visible then
    Rechercher (G_Cat, FindLigne, FindDebut) ;
end;

procedure TFTarifArticle.VoirFicheArticle;
BEGIN
{$IFNDEF GPAO}
	if ctxAffaire in V_PGI.PGIContexte then V_PGI.dispatchTT (7,taConsult,CodeArticle,'','')  //mcd 10/12/02
	else AglLanceFiche ('GC', 'GCARTICLE', '', CodeArticle, 'ACTION=CONSULTATION;TARIF=N');
{$ELSE GPAO }
  V_PGI.DispatchTT(7, taConsult, CodeArticle, 'TARIF=N', '');
{$ENDIF GPAO }
END;

procedure TFTarifArticle.BSaisieRapideClick(Sender: TObject);
var i_ind : integer;
begin
if Action = taConsult then exit;
if GF_CODEARTICLE.Text = '' then exit;
if PQTECAT.Visible then
    begin
    EntreeTarifRapide (taModif, TOBart, '', '', CodeDevise, ttdArtQCa, TOBTarfQCa, tstArticle);
    for i_ind:=0 to TOBTarfQCa.Detail.Count-1 do
        BEGIN
        AfficheLaLigne (i_ind + 1, ttdArtQCa) ;
        END ;
    end;
if PQUANTITATIF.Visible then
    begin
    EntreeTarifRapide (taModif, TOBart, '', '', CodeDevise, ttdArtQte, TOBTarfQte, tstArticle);
    for i_ind:=0 to TOBTarfQte.Detail.Count-1 do
        BEGIN
        AfficheLaLigne (i_ind + 1, ttdArtQte) ;
        END ;
    end;
END;

{==============================================================================================}
{================================= Validation =================================================}
{==============================================================================================}
procedure TFTarifArticle.BValiderClick(Sender: TObject);
Var ioerr : TIOErr ;
begin
if Action = taConsult then exit;
// validation
ioerr := Transactions (ValideTarif, 2);
Case ioerr of
        oeOk  : ;
    oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
    oeSaisie  : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
   END ;
if Not FicheAncetre then InitialiseEntete;
end;

procedure TFTarifArticle.ValideTarif;
begin
if Not SortDeLaLigne(ttdArt) then Exit ;
if Not SortDeLaLigne(ttdArtQte) then Exit ;
if Not SortDeLaLigne(ttdArtQCa) then Exit ;
if Not SortDeLaLigne(ttdArtCat) then Exit ;
TOBTarifDel.DeleteDB (False);
VerifLesTOB;
TOBTarif.InsertOrUpdateDB(False) ;
InitialiseGrille;
DepileTOBLigne;
end;

procedure TFTarifArticle.VerifLesTOB;
var i_ind : integer;
    Q : TQuery ;
    MaxTarif : Longint ;
BEGIN
Q := OpenSQL ('SELECT MAX(GF_TARIF) FROM TARIF', TRUE,-1,'',true) ;
if Q.EOF then MaxTarif := 1 else MaxTarif := Q.Fields[0].AsInteger + 1 ;
Ferme(Q) ;
for i_ind := TOBTarfArt.Detail.count - 1 Downto 0 do
    BEGIN
    if LigneVide (i_ind + 1, ttdArt) then
        BEGIN
        TOBTarfArt.Detail[i_ind].Free ;
        END else
        BEGIN
        if TOBTarfArt.Detail[i_ind].GetValue ('GF_TARIF') = 0 then
            BEGIN
            TOBTarfArt.Detail[i_ind].PutValue ('GF_TARIF', MaxTarif);
            Inc (MaxTarif);
            END;
        CalcPriorite (TOBTarfArt.Detail[i_ind]);
        END;
    END;
for i_ind := TOBTarfQte.Detail.count - 1 Downto 0 do
    BEGIN
    if LigneVide (i_ind + 1, ttdArtQte) then
        BEGIN
        TOBTarfQte.Detail[i_ind].Free ;
        END else
        BEGIN
        if TOBTarfQte.Detail[i_ind].GetValue ('GF_TARIF') = 0 then
            BEGIN
            TOBTarfQte.Detail[i_ind].PutValue ('GF_TARIF', MaxTarif);
            Inc (MaxTarif);
            END;
        CalcPriorite (TOBTarfQte.Detail[i_ind]);
        END;
    END;
for i_ind := TOBTarfQCa.Detail.count - 1 Downto 0 do
    BEGIN
    if LigneVide (i_ind + 1, ttdArtQCa) then
        BEGIN
        TOBTarfQCa.Detail[i_ind].Free ;
        END else
        BEGIN
        if TOBTarfQCa.Detail[i_ind].GetValue ('GF_TARIF') = 0 then
            BEGIN
            TOBTarfQCa.Detail[i_ind].PutValue ('GF_TARIF', MaxTarif);
            Inc (MaxTarif);
            END;
        CalcPriorite (TOBTarfQCa.Detail[i_ind]);
        END;
    END;
for i_ind := TOBTarfCat.Detail.count - 1 Downto 0 do
    BEGIN
    if LigneVide (i_ind + 1, ttdArtCat) then
        BEGIN
        TOBTarfCat.Detail[i_ind].Free ;
        END else
        BEGIN
        if TOBTarfCat.Detail[i_ind].GetValue ('GF_TARIF') = 0 then
            BEGIN
            TOBTarfCat.Detail[i_ind].PutValue ('GF_TARIF', MaxTarif);
            Inc (MaxTarif);
            END;
        CalcPriorite (TOBTarfCat.Detail[i_ind]);
        END;
    END;
END;

{==============================================================================================}
{=============================== Conditions tarifaires ========================================}
{==============================================================================================}
procedure TFTarifArticle.InitComboChamps ;
begin
SourisSablier ;
RemplitComboChamps('TIERS',FComboTIE) ;
RemplitComboChamps('ARTICLE',FComboART) ;
RemplitComboChamps('LIGNE',FComboLIG) ;
RemplitComboChamps('PIECE',FComboPIE) ;
SourisNormale ;
end ;

procedure TFTarifArticle.RemplitComboChamps(NomTable : String ; FCombo : THValComboBox) ;
begin
{Pref := TableToPrefixe(NomTable) ;
Q:=OpenSQL('SELECT DH_NOMCHAMP,DH_LIBELLE FROM DECHAMPS WHERE DH_PREFIXE="'+
           Pref + '" AND DH_CONTROLE LIKE "%T%"',True) ;
While not Q.EOF do
   begin
   if Trim(Q.FindField('DH_LIBELLE').AsString)='' then FCombo.Items.Add(Q.FindField('DH_NOMCHAMP').AsString)
                                                  else FCombo.Items.Add(Q.FindField('DH_LIBELLE').AsString) ;
   FCombo.Values.Add(Q.FindField('DH_NOMCHAMP').AsString) ;
   Q.Next ;
   end ;
Ferme(Q) ;}
ExtractFields(NomTable,'T',FCombo.Items,FCombo.Values,Nil,False);
end ;

procedure TFTarifArticle.EffaceGrid ;
var lig, col : Integer ;
begin
for lig := 1 to G_COND.RowCount - 1 do
   begin
   for col := 0 to G_COND.ColCount - 1 do G_COND.Cells[col, lig] := '' ;
   end ;
G_COND.Row := 1 ;
end ;

Procedure TFTarifArticle.AfficheCondTarf (ARow : Longint; ttd : T_TableTarif) ;
var TOBL : TOB;
    FF   : TForm;
BEGIN
TOBL:=GetTOBLigne(ARow, ttd) ; if TOBL=Nil then Exit ;
GetConditions (TOBL) ;
END;

function TFTarifArticle.ValueToItem(CC : THValComboBox ; St : String) : String ;
var i_ind : Integer ;
begin
i_ind := CC.Values.IndexOf(St) ;
if i_ind >= 0 then result := CC.Items[i_ind] else result := '' ;
end ;

{Charge dans le grid les conditions stockées dans le TMemoField}
procedure TFTarifArticle.GetConditions (TOBL : TOB) ;
var i_ind   : Integer ;
    st, NomTable, s1,s2 : String ;
begin
EffaceGrid ;
GF_CONDAPPLIC.Text := TOBL.GetValue ('GF_CONDAPPLIC');
for i_ind := 0 to GF_CONDAPPLIC.Lines.Count-1 do
   begin
   st := GF_CONDAPPLIC.Lines[i_ind] ;
   s1 := ReadTokenSt(st) ; // table
   NomTable := ValueToItem(FTable,s1) ;
   If NomTable='Client' then Nomtable :='Tiers'; //mcd 26/08/03 .. traduction à tort du nom de la table dans la fiche
   //ChargeComboChamps(s1) ; // Charge ponctuellement les champs de la table s1
   G_COND.Cells[0, i_ind + 1] := NomTable ;
   s1 := ReadTokenSt(st) ; // champ
   s2 := ValueToItem(THValComboBox(FindComponent('FCombo'+Copy(NomTable,1,3))),s1) ;
   G_COND.Cells[1, i_ind + 1] := s2 ;
   s1 := ReadTokenSt(st) ; // opérateur
   s2 := ValueToItem(FOpe,s1) ;
   G_COND.Cells[2, i_ind + 1] := s2 ;
   s1 := ReadTokenSt(st) ; // valeur
   G_COND.Cells[3, i_ind + 1] := s1 ;
   // Appel de la première ligne
   G_COND.Row := 1 ;
   //ChangeLigneGrid(G_COND.Row) ;
   end ;
end ;

procedure TFTarifArticle.TCONDTARFClose(Sender: TObject);
begin
BVoirCond.Down := False ;
end;

procedure TFTarifArticle.BAbandonClick(Sender: TObject);
begin
Close ;
if FClosing and IsInside(Self) then THPanel(parent).CloseInside ;
end;

Procedure AGLEntreeTarifArticle ( Parms : array of variant ; nb : integer ) ;
Var Action : TActionFiche ;
    St     : String ;
    OkTTC : Boolean ;
BEGIN
Action:=StringToAction(String(Parms[1])) ;
OkTTC:=False ;
St:=String(Parms[2]) ; if St='TTC' then OkTTC:=True ;
EntreeTarifArticle(Action,OkTTC) ;
END ;

procedure InitTarifArticle();
begin
RegisterAglProc('AppelTarifArticle',True,3,AppelTarifArticle) ;
RegisterAglProc('EntreeTarifArticle',True,2,AGLEntreeTarifArticle) ;
end;

procedure TFTarifArticle.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self);
end;

Initialization
InitTarifArticle();

end.

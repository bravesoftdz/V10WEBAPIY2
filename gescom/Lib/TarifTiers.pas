unit TarifTiers;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, Grids, Hctrls, ComCtrls, ExtCtrls, HPanel, StdCtrls, Mask, UIUtil,
  Hent1, TarifUtil, SaisUtil, HSysMenu,
{$IFDEF EAGLCLIENT}
  MaineAGL,                         
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_main,
{$ENDIF}
  UTOB, math, hmsgbox, Ent1,
  AglInit,HDimension, Menus, TarifRapide, HRichEdt, HRichOLE,
{$IFNDEF CCS3}
  TarifCond,
{$ENDIF}
  AglInitGC, UtilArticle, EntGC, M3FP, TntGrids, TntStdCtrls, TntComCtrls,
  TntExtCtrls ;

Function EntreeTarifTiers (Action : TActionFiche; TarifTTC : Boolean=False) : boolean ;

Type T_TableTarif = (ttdCatArt, ttdCatQArt, ttdCatFam) ;

type
  TFTarifTiers = class(TForm)
    Panel1: TPanel;
    PPIED: TPanel;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    GF_DEVISE: THValComboBox;
    TGF_LIBELLE: TLabel;
    MsgBox: THMsgBox;
    HMTrad: THSystemMenu;
    TGF_DEVISE: TLabel;
    GF_CASCADEREMISE: THValComboBox;
    TGF_CASCADEREMISE: THLabel;
    TGF_REMISE: THLabel;
    GF_REMISE: THCritMaskEdit;
    BChercher: TToolbarButton97;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    FindLigne: TFindDialog;
    BInfos: TToolbarButton97;
    POPZ: TPopupMenu;
    InfArticle: TMenuItem;
    TGF_TARIFTIERS: TLabel;
    cbValider: TCheckBox;
    GF_TARIFTIERS: THValComboBox;
    HMess: THMsgBox;
    BSaisieRapide: TToolbarButton97;
    PFAMARTICLE: THPanel;
    PTITRE: THPanel;
    G_FAM: THGrid;
    HTitre: THMsgBox;
    PTARIFART: THPanel;
    PARTICLE: THPanel;
    TGF_ARTICLE: TLabel;
    TGF_LIBARTICLE: TLabel;
    GF_ARTICLE: THCritMaskEdit;
    PDIMENSION: THPanel;
    TGF_GRILLEDIM5: THLabel;
    TGF_GRILLEDIM4: THLabel;
    TGF_GRILLEDIM3: THLabel;
    TGF_GRILLEDIM2: THLabel;
    TGF_GRILLEDIM1: THLabel;
    GF_CODEDIM5: THCritMaskEdit;
    GF_CODEDIM4: THCritMaskEdit;
    GF_CODEDIM3: THCritMaskEdit;
    GF_CODEDIM2: THCritMaskEdit;
    GF_CODEDIM1: THCritMaskEdit;
    PQTEART: THPanel;
    PTART: THPanel;
    G_QART: THGrid;
    G_ART: THGrid;
    CBQUANTITATIF: TCheckBox;
    CONDAPPLIC: THRichEditOLE;
    CBARTICLE: TCheckBox;
    BCollerCond: TToolbarButton97;
    BCopierCond: TToolbarButton97;
    BCondAplli: TToolbarButton97;
    BVoirCond: TToolbarButton97;
    TCONDTARF: TToolWindow97;
    G_COND: THGrid;
    FComboTIE: THValComboBox;
    FComboART: THValComboBox;
    FComboLIG: THValComboBox;
    FComboPIE: THValComboBox;
    GF_CONDAPPLIC: THRichEditOLE;
    FTable: THValComboBox;
    FOpe: THValComboBox;
    TTYPETARIF: THLabel;
    ISigneEuro: TImage;
    TGF_PRIXCON: THLabel;
    ISigneFranc: TImage;
    GF_PRIXCON: THNumEdit;
    GA_PRIXPOURQTE: THNumEdit;
    TGA_PRIXPOURQTE: THLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
//
    procedure GF_TARIFTIERSChange(Sender: TObject);
//
    procedure GF_DEVISEChange(Sender: TObject);
//
    procedure GF_ARTICLEExit(Sender: TObject);
    procedure GF_ARTICLEElipsisClick(Sender: TObject);
    procedure GF_ARTICLEChange(Sender: TObject);
//
    procedure G_ARTEnter(Sender: TObject);
//
    procedure G_QARTEnter(Sender: TObject);
//
    procedure G_FAMEnter(Sender: TObject);
//
    procedure BValiderClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure FindLigneFind(Sender: TObject);
    procedure G_ACTIVEEnter(Sender: TObject);
    procedure G_ACTIVERowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_ACTIVERowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_ACTIVECellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_ACTIVECellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_ACTIVEElipsisClick(Sender: TObject);
    procedure BSaisieRapideClick(Sender: TObject);
    procedure CBQUANTITATIFClick(Sender: TObject);
    procedure CBARTICLEClick(Sender: TObject);
    procedure BCondAplliClick(Sender: TObject);
    procedure BCopierCondClick(Sender: TObject);
    procedure BCollerCondClick(Sender: TObject);
    procedure BVoirCondClick(Sender: TObject);
    procedure InfArticleClick(Sender: TObject);
    procedure TCONDTARFClose(Sender: TObject);
    procedure GF_CASCADEREMISEChange(Sender: TObject);
    procedure GF_REMISEChange(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    { Déclarations privées }
    iTableLigne : integer ;
    FindDebut,FClosing : Boolean;
    StCellCur,LesColART, LesColQART, LesColFAM : String ;
    ColsInter : Array of boolean ;
    DEV       : RDEVISE ;
    TarifTTC  : Boolean ;
    PrixPourQte : Double;
// Objets mémoire
    TOBTarif, TOBTarfArt, TOBTarfFam, TOBTarfQArt, TOBTarifDel, TOBCatTiers, TOBArt  : TOB;
// Menu
    procedure AffectMenuCondApplic (G_GRID : THGrid) ;
// Actions liées au Grid
    procedure EtudieColsListe ;
    procedure PostDrawCell(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    procedure FormateZoneSaisie (ACol,ARow : Longint) ;
    procedure InsertLigne (ARow : Longint) ;
    procedure SupprimeLigne (G_DEL : THGrid; TobEnCours : TOB; ARow : Longint) ;
    procedure SupprimeTOBTarif (TOBEnCours : TOB; ARow : Longint) ;
    procedure ActiveART;
    procedure ActiveQART;
    procedure ActiveFAM;
    Function  GrilleModifie : Boolean;
    Function  SortDeLaLigne (ttd : T_TableTarif) : boolean ;
// Initialisations liées au tiers et/ou la devise
    Procedure TraiterCatTiers ;
    Function  QuestionTarifEnCours : Integer;
    procedure InitialiseEntete ;
    procedure AffecteEntete ;
    procedure LoadLesTOB ;
    procedure ChargeTarif ;
    procedure VerifLesTOB;
    Function  WhereTarifTiers (CodeCatTiers, CodeDevise, CodeArticle : String; ttd : T_TableTarif ;
                               Totale : boolean ) : String ;
// Initialisations liées à l'article et/ou aux dimensions
    Procedure TraiterArticle ;
    Function  QuestionTarifArticleEnCours : Integer;
    procedure InitialiseEnteteArticle ;
    procedure AffecteEnteteArticle ;
    Function  ChoisirDimension (St_CodeArt: string) : Boolean;
    Procedure AffecteDimension ;
    Procedure ChercherArticle ;
    procedure LoadLesTOBArticle ;
    procedure ChargeTarifArticle ;
    procedure VerifLesTOBArticle;
    procedure ErgoGCS3 ;
// LIGNES
    Procedure InitLaLigne (ARow : integer) ;
    //procedure VideCodesLigne ( ARow : integer ) ;
    Function  GetTOBLigne (ARow : integer) : TOB ;
    {$IFDEF MONCODE}
    procedure AffichePrix (TOBL : TOB ; ARow : integer) ;
    {$ENDIF}
    procedure AfficheLaLigne (ARow : integer) ;
    procedure InitialiseGrille ;
    procedure InitialiseGrilleArticle ;
    procedure InitialiseLigne (G_ACTIVE: THGrid; ARow : integer) ;
    Procedure DepileTOBLigne;
    Procedure DepileTOBLigneArticle ;
    Procedure CreerTOBLigne (ARow : integer);
    function  LigneVide ( ARow : integer) : Boolean;
    procedure PreAffecteLigne (ARow : integer);
    procedure ValideTarif;
    procedure ValideTarifArticle;
// CHAMPS LIGNES
    procedure TraiterDepot (ACol, ARow : integer);
    procedure TraiterCatArticle (ACol, ARow : integer);
    procedure TraiterLibelle (ACol, ARow : integer);
    procedure TraiterBorneInf (ACol, ARow : integer);
    procedure TraiterBorneSup (ACol, ARow : integer);
    procedure TraiterPrix (ACol, ARow : integer);
    procedure TraiterRemise (ACol, ARow : integer);
    procedure TraiterDateDeb (ACol, ARow : integer);
    procedure TraiterDateFin (ACol, ARow : integer);
// PIED
    Procedure InitialisePied ;
    Procedure AffichePrixCon ;
    Procedure AffichePied (ARow : Longint) ;
// Boutons
    procedure VoirFicheArticle;
// Conditions tarifaires
    procedure InitComboChamps ;
    procedure RemplitComboChamps(NomTable : String ; FCombo : THValComboBox) ;
    procedure EffaceGrid ;
    Procedure AfficheCondTarf (ARow : Longint) ;
    function  ValueToItem(CC : THValComboBox ; St : String) : String ;
    procedure GetConditions (TOBL : TOB) ;
  public
    { Déclarations publiques }
    CodeCatTiers   : string;
    CodeArticle   : string;
    CodeDevise  : string;
    CodeDeviseOld : string;
    Action      : TActionFiche ;
    TOB_ACTIVE : TOB;
    G_ACTIVE : THGrid;
    Cols_Actives : string;
    ACT_Depot   : integer;
    ACT_Cat     : integer;
    ACT_Lib     : integer;
    ACT_QInf    : integer;
    ACT_QSup    : integer;
    ACT_Px      : integer;
    ACT_Rem     : integer;
    ACT_Datedeb : integer;
    ACT_Datefin : integer;
  end;

var
  lModif : boolean;
  SGA_Depot   : integer;
  SGA_Lib     : integer;
  SGA_Px      : integer;
  SGA_Rem     : integer;
  SGA_Datedeb : integer;
  SGA_Datefin : integer;

  SFA_Depot   : integer;
  SFA_Cat     : integer;
  SFA_Lib     : integer;
//  SFA_QInf    : integer;
//  SFA_QSup    : integer;
  SFA_Px      : integer;
  SFA_Rem     : integer;
  SFA_Datedeb : integer;
  SFA_Datefin : integer;

  SQA_Depot   : integer;
  SQA_Lib     : integer;
  SQA_QInf    : integer;
  SQA_QSup    : integer;
  SQA_Px      : integer;
  SQA_Rem     : integer;
  SQA_Datedeb : integer;
  SQA_Datefin : integer;

implementation
uses
   CbpMCD
   ,CbpEnumerator
;

{$R *.DFM}

Function EntreeTarifTiers (Action : TActionFiche; TarifTTC : Boolean=False) : boolean ;
var FF : TFTarifTiers ;
    PPANEL  : THPanel ;
begin
SourisSablier;
FF := TFTarifTiers.Create(Application) ;
FF.Action:=Action ;
FF.TarifTTC:=TarifTTC ;
PPANEL := FindInsidePanel ;
if PPANEL = Nil then
   BEGIN
    try
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
Result := True;
END ;

//==============================================================================
//  Action liées à la forme
//==============================================================================
procedure TFTarifTiers.FormCreate(Sender: TObject);
begin
G_ART.RowCount := NbRowsInit ;
G_FAM.RowCount := NbRowsInit ;
G_QART.RowCount := NbRowsInit ;
StCellCur := '' ;
iTableLigne := PrefixeToNum('GF') ;
TOBTarif := TOB.Create ('', Nil, -1) ;
TOBTarfArt := TOB.Create ('', TOBTarif, 0) ;
TOBTarfFam := TOB.Create ('', TOBTarif, 1) ;
TOBTarfQArt := TOB.Create ('', TOBTarif, 2) ;
TOBTarifDel := TOB.Create ('', Nil, -1) ;
TOBCatTiers := TOB.Create ('CHOIXCOD', Nil, -1) ;
TOBArt := TOB.Create ('ARTICLE', Nil, -1) ;
FClosing:=False ;
end;

procedure TFTarifTiers.ErgoGCS3 ;
BEGIN
{$IFDEF CCS3}
BVoirCond.Visible:=False   ; BCondAplli.Visible:=False ;
BCopierCond.Visible:=False ; BCollerCond.Visible:=False ;
{$ENDIF}
END ;

procedure TFTarifTiers.FormShow(Sender: TObject);
begin
If (ctxAffaire in V_PGI.PGIContexte) then G_ART.ListeParam:='AFTARIFPRIX'
else G_ART.ListeParam:='GCTARIFPRIX';
G_ART.PostDrawCell:=PostDrawCell ;
If (ctxAffaire in V_PGI.PGIContexte) then   G_FAM.ListeParam:='AFTARIFCA'
else G_FAM.ListeParam:='GCTARIFCA';
G_FAM.PostDrawCell:=PostDrawCell ;
If (ctxAffaire in V_PGI.PGIContexte) then G_QART.ListeParam:='AFTARIFQTEPRIX'
else G_QART.ListeParam:='GCTARIFQTEPRIX';
G_QART.PostDrawCell:=PostDrawCell ;
EtudieColsListe ;
HMTrad.ResizeGridColumns (G_ART) ;
HMTrad.ResizeGridColumns (G_FAM) ;
HMTrad.ResizeGridColumns (G_QART) ;
AffecteGrid (G_ART,Action) ;
AffecteGrid (G_FAM,Action) ;
AffecteGrid (G_QART,Action) ;
PFAMARTICLE.Visible := True;
PTARIFART.Visible := False;
PQTEART.Visible := False ;
PTART.Visible := False ;
PTITRE.Caption := HTitre.Mess[0] ;
TTYPETARIF.Caption := HTitre.Mess[4] ;
CBARTICLE.Checked := False;
cbValider.Visible := CBARTICLE.Checked;
CBQUANTITATIF.Checked := False ;
CodeDevise := V_PGI.DevisePivot;
DEV.Code := CodeDevise ; GetInfosDevise (DEV) ;
InitialiseEntete ;
lModif := True;
InitComboChamps ;
ErgoGCS3 ;
end;

procedure TFTarifTiers.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
if Action=taConsult then Exit ;
if GrilleModifie then
   BEGIN
   if MsgBox.Execute(6,Caption,'')<>mrYes then CanClose:=False ;
   END ;
end;

procedure TFTarifTiers.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
G_ART.VidePile(True) ;
G_FAM.VidePile(True) ;
G_QART.VidePile(True) ;
TOBTarif.Free ; TOBTarif:=Nil ;
TOBCatTiers.Free ; TOBCatTiers:=Nil ;
TOBArt.Free ; TOBArt:=Nil ;
if IsInside(Self) then Action:=caFree ;
FClosing:=True ;
end;

procedure TFTarifTiers.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var FocusGrid : Boolean;
    G_DEL : THGrid;
    ARow : Longint;
BEGIN
FocusGrid := False;
ARow := -1;
G_DEL := nil;
if(Screen.ActiveControl = G_ART) then
    BEGIN
    FocusGrid := True;
    TOB_ACTIVE := TOBTarfArt;
    G_DEL := G_ACTIVE;
    ARow := G_ACTIVE.Row;
    END else
    if (Screen.ActiveControl = G_QART) then
        BEGIN
        FocusGrid := True;
        TOB_ACTIVE := TOBTarfQArt;
        G_DEL := G_QART;
        ARow := G_QART.Row;
        END else
        if (Screen.ActiveControl = G_FAM) then
            BEGIN
            FocusGrid := True;
            TOB_ACTIVE := TOBTarfFam;
            G_DEL := G_FAM;
            ARow := G_FAM.Row;
            END;
Case Key of
    VK_RETURN : Key:=VK_TAB ;
    VK_INSERT : BEGIN
                if (FocusGrid) and (ARow <> -1) then
                    BEGIN
                    Key := 0;
                    InsertLigne (ARow);
                    END;
                END;
    VK_DELETE : BEGIN
                if ((FocusGrid) and (Shift=[ssCtrl])) and (G_DEL <> nil) then
                    BEGIN
                    Key := 0 ;
                    SupprimeLigne (G_DEL, TOB_ACTIVE, ARow) ;
                    END ;
                END;
    END;
end;

{==============================================================================================}
{============================= Manipulation liées au Menu =====================================}
{==============================================================================================}
procedure TFTarifTiers.AffectMenuCondApplic (G_GRID : THGrid) ;
BEGIN
if (LigneVide (G_GRID.Row))then
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
procedure TFTarifTiers.PostDrawCell(ACol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
begin
//
end ;

procedure TFTarifTiers.EtudieColsListe ;
Var NomCol,LesCols : String ;
    icol,ichamp, i_ind : integer ;
		Mcd : IMCDServiceCOM;
BEGIN
MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();
//
G_ART.ColWidths[0]:=0 ;
G_FAM.ColWidths[0]:=0 ;
G_QART.ColWidths[0]:=0 ;
SetLength(ColsInter,G_ART.ColCount) ;
for i_ind:=Low(ColsInter)to High(ColsInter) do ColsInter[i_ind]:=False ;
LesCols:=G_ART.Titres[0] ; LesColART:=LesCols ; icol:=0 ;
Repeat
 NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
 if NomCol<>'' then
    BEGIN
		if assigned(mcd.getField(Nomcol)) then
       BEGIN
       if Pos('X',mcd.getField(Nomcol).Control)>0 then ColsInter[icol]:=True ;
       if NomCol='GF_DEPOT'        then SGA_Depot:=icol else
       if NomCol='GF_LIBELLE'      then SGA_Lib:=icol else
//       if NomCol='GF_BORNEINF'     then SGA_QInf:=icol else
//       if NomCol='GF_BORNESUP'     then SGA_QSup:=icol else
       if NomCol='GF_PRIXUNITAIRE' then SGA_Px:=icol else
       if NomCol='GF_CALCULREMISE' then SGA_Rem:=icol else
       if NomCol='GF_DATEDEBUT'    then SGA_Datedeb:=icol else
       if NomCol='GF_DATEFIN'      then SGA_Datefin:=icol ;
       END ;
    END ;
 Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;

LesCols:=G_FAM.Titres[0] ; LesColFAM := LesCols ; icol:=0 ;
Repeat
 NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
 if NomCol<>'' then
    BEGIN
		if assigned(mcd.getField(Nomcol)) then
       BEGIN
       if NomCol='GF_DEPOT'        then SFA_Depot:=icol else
       if NomCol='GF_TARIFARTICLE' then SFA_Cat:=icol else
       if NomCol='GF_LIBELLE'      then SFA_Lib:=icol else
//       if NomCol='GF_BORNEINF'     then SFA_QInf:=icol else
//       if NomCol='GF_BORNESUP'     then SFA_QSup:=icol else
       if NomCol='GF_PRIXUNITAIRE' then SFA_Px:=icol else
       if NomCol='GF_CALCULREMISE' then SFA_Rem:=icol else
       if NomCol='GF_DATEDEBUT'    then SFA_Datedeb:=icol else
       if NomCol='GF_DATEFIN'      then SFA_Datefin:=icol ;
       END ;
    END ;
 Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;

SetLength(ColsInter,G_QART.ColCount) ;
for i_ind:=Low(ColsInter)to High(ColsInter) do ColsInter[i_ind]:=False ;
LesCols:=G_QART.Titres[0] ; LesColQART:=LesCols ; icol:=0 ;
Repeat
 NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
 if NomCol<>'' then
    BEGIN
		if assigned(mcd.getField(Nomcol)) then
       BEGIN
       if Pos('X',mcd.getField(Nomcol).Control)>0 then ColsInter[icol]:=True ;
       if NomCol='GF_DEPOT'        then SQA_Depot:=icol else
       if NomCol='GF_LIBELLE'      then SQA_Lib:=icol else
       if NomCol='GF_BORNEINF'     then SQA_QInf:=icol else
       if NomCol='GF_BORNESUP'     then SQA_QSup:=icol else
       if NomCol='GF_PRIXUNITAIRE' then SQA_Px:=icol else
       if NomCol='GF_CALCULREMISE' then SQA_Rem:=icol else
       if NomCol='GF_DATEDEBUT'    then SQA_Datedeb:=icol else
       if NomCol='GF_DATEFIN'      then SQA_Datefin:=icol ;
       END ;
    END ;
 Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;

END ;

procedure TFTarifTiers.FormateZoneSaisie (ACol,ARow : Longint) ;
Var St,StC : String ;
BEGIN

St:=G_ACTIVE.Cells[ACol,ARow] ; StC:=St ;
if G_ACTIVE = G_ART then
    begin
    if ACol=SGA_Depot then StC:=uppercase(Trim(St)) else
        {$IFDEF MONCODE}
        if ACol=SGA_Px then StC:=StrF00(Valeur(St),DEV.Decimale+NbreDecimalSuppPrix('CLI',TOBArt)) ;
        {$ELSE}
        if ACol=SGA_Px then StC:=StrF00(Valeur(St),V_PGI.OkDecP);
        {$ENDIF}
    end;
if G_ACTIVE = G_QART then
    begin
    if (ACol=SQA_Depot) then StC:=uppercase(Trim(St)) else
        {$IFDEF MONCODE}
        if ACol=SQA_Px then StC:=StrF00(Valeur(St),DEV.Decimale+NbreDecimalSuppPrix('CLI',TOBArt)) else
        {$ELSE}
        if ACol=SQA_Px then StC:=StrF00(Valeur(St),V_PGI.OkDecP) else
        {$ENDIF}
            if ((ACol=SQA_QInf) or (ACol=SQA_QSup)) then StC:=StrF00(Valeur(St),0);
    end;
if G_ACTIVE = G_FAM then
    begin
    if ((ACol=SFA_Depot) or (ACol=SFA_Cat)) then StC:=uppercase(Trim(St)) else
        if ACol=SFA_Px then StC:=StrF00(Valeur(St),V_PGI.OkDecP) //else
//            if ((ACol=SFA_QInf) or (ACol=SFA_QSup)) then StC:=StrF00(Valeur(St),0);
    end;
G_ACTIVE.Cells[ACol,ARow]:=StC ;
END ;

procedure TFTarifTiers.InsertLigne (ARow : Longint) ;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if LigneVide (ARow) then exit;
if (ARow > TOB_ACTIVE.Detail.Count) then Exit;
G_ACTIVE.CacheEdit; G_ACTIVE.SynEnabled := False;
TOB.Create ('TARIF', TOB_ACTIVE, ARow-1) ;
G_ACTIVE.InsertRow (ARow); G_ACTIVE.Row := ARow;
InitialiseLigne (G_ACTIVE, ARow) ;
PreAffecteLigne (ARow);
G_ACTIVE.MontreEdit; G_ACTIVE.SynEnabled := True;
AffectMenuCondApplic (G_ACTIVE);
AfficheCondTarf (G_ACTIVE.Row);
END;

procedure TFTarifTiers.SupprimeLigne (G_DEL : THGrid; TobEnCours : TOB; ARow : Longint) ;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if (ARow > TOBEnCours.Detail.Count) then Exit;
G_DEL.CacheEdit; G_DEL.SynEnabled := False;
G_DEL.DeleteRow (ARow);
if (ARow = TOBEnCours.Detail.Count) then
    CreerTOBLigne (ARow + 1);
SupprimeTOBTarif (TOBEncours, ARow);
if G_DEL.RowCount < NbRowsInit then G_DEL.RowCount := NbRowsInit;
G_DEL.MontreEdit; G_DEL.SynEnabled := True;
AffectMenuCondApplic (G_ACTIVE);
AfficheCondTarf (G_ACTIVE.Row);
END;

procedure TFTarifTiers.SupprimeTOBTarif (TOBEnCours : TOB; ARow : Longint) ;
Var i_ind: integer;
BEGIN
if TOBEnCours.Detail[ARow-1].GetValue ('GF_TARIF') <> 0 then
    BEGIN
    i_ind := TOBTarifDel.Detail.Count;
    TOB.Create ('TARIF', TOBTarifDel, i_ind) ;
    TOBTarifDel.Detail[i_ind].Dupliquer (TOBEnCours.Detail[ARow-1], False, True);
    END;
TOBEnCours.Detail[ARow-1].Free;
TOBEnCours.SetAllModifie(True);
END;

procedure TFTarifTiers.ActiveART;
begin
TOB_ACTIVE := TOBTarfArt;
G_ACTIVE := G_ART;
Cols_Actives := LesColART;
ACT_Depot   := SGA_Depot;
ACT_Cat     := -1;
ACT_Lib     := SGA_Lib;
ACT_QInf    := -1;
ACT_QSup    := -1;
ACT_Px      := SGA_Px;
ACT_Rem     := SGA_Rem;
ACT_Datedeb := SGA_Datedeb;
ACT_Datefin := SGA_Datefin;
end;

procedure TFTarifTiers.ActiveQART;
begin
TOB_ACTIVE := TOBTarfQArt;
G_ACTIVE := G_QART;
Cols_Actives := LesColQART;
ACT_Depot   := SQA_Depot;
ACT_Cat     := -1;
ACT_Lib     := SQA_Lib;
ACT_QInf    := SQA_QInf;
ACT_QSup    := SQA_QSup;
ACT_Px      := SQA_Px;
ACT_Rem     := SQA_Rem;
ACT_Datedeb := SQA_Datedeb;
ACT_Datefin := SQA_Datefin;
end;

procedure TFTarifTiers.ActiveFAM;
begin
TOB_ACTIVE := TOBTarfFam;
G_ACTIVE := G_FAM;
Cols_Actives := LesColFAM;
ACT_Depot   := SFA_Depot;
ACT_Cat     := SFA_Cat;
ACT_Lib     := SFA_Lib;
ACT_QInf    := -1;
ACT_QSup    := -1;
ACT_Px      := SFA_Px;
ACT_Rem     := SFA_Rem;
ACT_Datedeb := SFA_Datedeb;
ACT_Datefin := SFA_Datefin;
end;

Function TFTarifTiers.GrilleModifie : Boolean;
BEGIN
Result:=False ;
if Action=taConsult then Exit ;
Result:=(TOBTarfArt.IsOneModifie) or (TOBTarfQArt.IsOneModifie) or
        (TOBTarfFam.IsOneModifie) or (TOBTarifDel.IsOneModifie);
END;

Function TFTarifTiers.SortDeLaLigne (ttd : T_TableTarif) : boolean ;
Var ACol,ARow : integer ;
    Cancel : boolean ;
BEGIN
Result:=False ;
Case ttd of
    ttdCatArt : ActiveART ;
    ttdCatQArt : ActiveQART ;
    ttdCatFam : ActiveFAM
    END;
ACol:=G_ACTIVE.Col ; ARow:=G_ACTIVE.Row ; Cancel:=False ;
G_ACTIVECellExit(Nil,ACol,ARow,Cancel) ; if Cancel then Exit ;
G_ACTIVERowExit(Nil,ACol,Cancel,False) ; if Cancel then Exit ;
Result:=True ;
END ;

{==============================================================================================}
{=============================== Evenements du pied ===========================================}
{==============================================================================================}
procedure TFTarifTiers.GF_REMISEChange(Sender: TObject);
var St : string;
begin
St := GF_REMISE.Text;
St:=StrF00(Valeur(St),ADecimP);
GF_REMISE.Text := St;
end;

procedure TFTarifTiers.GF_CASCADEREMISEChange(Sender: TObject);
Var Row : integer;
    TOBL : TOB ;
begin
if PFAMARTICLE.Visible=True then ActiveFAM
    else if PQTEART.Visible=True then ActiveQART
        else if PTART.Visible=True then ActiveART
            else exit;
Row := G_ACTIVE.Row;
if Row > 0 then
    BEGIN
    TOBL := GetTOBLigne (Row); if TOBL=nil then exit;
    TOBL.PutValue ('GF_CASCADEREMISE', GF_CASCADEREMISE.Value);
    END;
end;

{==============================================================================================}
{============================= Manipulation du pied ===========================================}
{==============================================================================================}
Procedure TFTarifTiers.InitialisePied ;
BEGIN
GF_CASCADEREMISE.Text := '';
GF_REMISE.Text := '';
GF_PRIXCON.Text := '';
AffichePrixCon;
END;

Procedure TFTarifTiers.AffichePrixCon ;
BEGIN
if (GF_DEVISE.Value<>V_PGI.DevisePivot) or (PFAMARTICLE.Visible) or (ForceEuroGC) then
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

Procedure TFTarifTiers.AffichePied (ARow : Longint) ;
var TOBL : TOB;
    FF   : TForm;
BEGIN
TOBL:=GetTOBLigne(ARow) ; if TOBL=Nil then Exit ;
FF := TForm (PPIED.Owner);
TOBL.PutEcran (FF, PPIED);
END;

{==============================================================================================}
{========================= Manipulation des LIGNES Quantitatif ================================}
{==============================================================================================}
Procedure TFTarifTiers.InitLaLigne (ARow : integer);
Var TOBL : TOB;
BEGIN
TOBL:=GetTOBLigne(ARow) ; if TOBL=Nil then Exit ;
TOBL.PutValue ('GF_NATUREAUXI', 'CLI'); // DBR a voir
TOBL.PutValue ('GF_TARIFTIERS', CodeCatTiers);
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
TOBL.PutValue ('GF_SOCIETE', V_PGI.CodeSociete);
if G_ACTIVE = G_FAM then
    BEGIN
    TOBL.PutValue ('GF_QUANTITATIF', '-');
    TOBL.PutValue ('GF_BORNEINF', -999999);
    TOBL.PutValue ('GF_REGIMEPRIX', 'GLO') ;
    END else if G_ACTIVE = G_ART then
    BEGIN
    TOBL.PutValue ('GF_QUANTITATIF', '-');
    TOBL.PutValue ('GF_BORNEINF', -999999);
    if TarifTTC then TOBL.PutValue ('GF_REGIMEPRIX', 'TTC')
                else TOBL.PutValue ('GF_REGIMEPRIX', 'HT') ;
    END else
    BEGIN
    TOBL.PutValue ('GF_QUANTITATIF', 'X');
    TOBL.PutValue ('GF_BORNEINF', 1);
    if TarifTTC then TOBL.PutValue ('GF_REGIMEPRIX', 'TTC')
                else TOBL.PutValue ('GF_REGIMEPRIX', 'HT') ;
    END;
AfficheLAligne (ARow);
END;

Function TFTarifTiers.GetTOBLigne ( ARow : integer) : TOB ;
BEGIN
Result:=Nil ;
if ((TOB_ACTIVE = nil) or (ARow<=0) or (ARow>TOB_ACTIVE.Detail.Count)) then Exit ;
Result:=TOB_ACTIVE.Detail[ARow-1] ;
END ;

procedure TFTarifTiers.InitialiseGrille ;
BEGIN
InitialiseGrilleArticle ;
G_FAM.VidePile(True) ;
G_FAM.RowCount:= NbRowsInit ;
END;

procedure TFTarifTiers.InitialiseGrilleArticle ;
BEGIN
G_ART.VidePile(True) ;
G_ART.RowCount:= NbRowsInit ;
G_QART.VidePile(True) ;
G_QART.RowCount:= NbRowsInit ;
END;

procedure TFTarifTiers.InitialiseLigne (G_ACTIVE: THGrid; ARow : integer) ;
Var TOBL : TOB ;
    i_ind : integer ;
BEGIN
TOBL:=GetTOBLigne(ARow) ; if TOBL<>Nil then TOBL.InitValeurs ;
TOBL.PutValue('GF_TARIFTIERS',GF_TARIFTIERS.Text);
TOBL.PutValue('GF_ARTICLE',TOBArt.GetValue('GA_ARTICLE'));
for i_ind := 1 to G_ACTIVE.ColCount-1 do
    BEGIN
    G_ACTIVE.Cells [i_ind, ARow]:='' ;
    END;
END;

{$IFDEF MONCODE}
procedure TFTarifTiers.AffichePrix (TOBL : TOB ; ARow : integer) ;
var PuFix : double;
begin
if (TOBArt.GetValue('GA_DECIMALPRIX')='X') then
  begin
  PuFix:=TOBL.GetValue('GF_PRIXUNITAIRE');
  if PrixPourQte<=0 then PrixPourQte:=1;
  PuFix:=PuFix/PrixPourQte;
  G_ACTIVE.cells[ACT_Px,Arow] := strf00 (PuFix,DEV.Decimale+NbreDecimalSuppPrix('CLI',TOBArt));
  end;
end;
{$ENDIF}

procedure TFTarifTiers.AfficheLaLigne (ARow : integer) ;
Var TOBL : TOB ;
    i_ind,i_ind2 : integer ;
    nbDec : string;
BEGIN
TOBL:=GetTOBLigne(ARow) ; if TOBL = Nil then exit;
nbDec := '0.';
for i_ind := 1 to G_ACTIVE.ColCount-1 do
    BEGIN
    if ((i_ind=ACT_Datedeb) or (i_ind=ACT_Datefin)) then
        G_ACTIVE.ColFormats [i_ind] := '';
    if i_ind=ACT_Px then
       begin
       {$IFDEF MONCODE}
       for i_ind2 := 1 to DEV.Decimale+NbreDecimalSuppPrix('CLI',TOBArt) do nbDec:=nbDec + '0';
       {$ELSE}
       for i_ind2 := 1 to V_PGI.OkDecP do nbDec:=nbDec + '0';
       {$ENDIF}
       G_ACTIVE.ColFormats [i_ind] := '# ##' + nbDec;
       end;
    END;
TOBL.PutLigneGrid(G_ACTIVE,ARow,False,False,Cols_Actives) ;
{$IFDEF MONCODE}
AffichePrix(TOBL,ARow) ;
{$ENDIF}
if G_ACTIVE.Cells [ACT_Datedeb, ARow] = '' then
    G_ACTIVE.Cells [ACT_Datedeb, ARow] := DateToStr (IDate1900);
for i_ind := 1 to G_ACTIVE.ColCount-1 do
    FormateZoneSaisie(i_ind, ARow) ;
END ;

Procedure TFTarifTiers.DepileTOBLigne ;
var i_ind : integer;
BEGIN
for i_ind := TOBTarfFam.Detail.Count - 1 Downto 0 do
    BEGIN
    TOBTarfFam.Detail[i_ind].Free ;
    END;
for i_ind := TOBTarifDel.Detail.Count - 1 Downto 0 do
    BEGIN
    TOBTarifDel.Detail[i_ind].Free ;
    END;
DepileTOBLigneArticle;
END;

Procedure TFTarifTiers.DepileTOBLigneArticle ;
var i_ind : integer;
BEGIN
for i_ind := TOBTarfQArt.Detail.Count - 1 Downto 0 do
    BEGIN
    TOBTarfQArt.Detail[i_ind].Free ;
    END;
for i_ind := TOBTarfArt.Detail.Count - 1 Downto 0 do
    BEGIN
    TOBTarfArt.Detail[i_ind].Free ;
    END;
END;

Procedure TFTarifTiers.CreerTOBLigne (ARow : integer);
BEGIN
if ARow <> TOB_ACTIVE.Detail.Count + 1 then exit;
TOB.Create ('TARIF', TOB_ACTIVE, ARow-1) ;
InitialiseLigne (G_Active, ARow) ;
END;

Function TFTarifTiers.LigneVide (ARow : integer) : Boolean;
BEGIN
Result := True;
if (G_ACTIVE = G_FAM) then
    BEGIN
        if (G_ACTIVE.Cells [ACT_Lib, ARow] <> '')  and (G_ACTIVE.Cells [ACT_Cat, ARow] <> '') then
        BEGIN
        Result := False;
        END ;
    END else
        if (G_ACTIVE.Cells [ACT_Lib, ARow] <> '') then
            BEGIN
            Result := False;
            END;
END;

Procedure TFTarifTiers.PreAffecteLigne (ARow : integer);
var TOBL, TOBLPrec : TOB;
BEGIN
TOBLPrec := GetTOBLigne (ARow - 1); if TOBLPrec = nil then exit;
if TOBLPrec.GetValue ('GF_BORNESUP') < 999999 then
    BEGIN
    TOB_ACTIVE.Detail[ARow - 1].Dupliquer (TOBLPrec, False, True);
    TOBL := GetTOBLigne (ARow);
    TOBL.PutValue ('GF_BORNEINF', TOBLPrec.GetValue ('GF_BORNESUP') + 1);
    TOBL.PutValue ('GF_BORNESUP', 999999);
    TOBL.PutValue ('GF_TARIF', 0);
    AfficheLAligne (ARow);
    END;
END;

procedure TFTarifTiers.LoadLesTOB ;
Var Q : TQuery ;
    i_ind : integer;
    Select : string;
BEGIN
for i_ind := TOBTarfFam.Detail.Count - 1 downto 0 do
    BEGIN
    TOBTarfFam.Detail[i_ind].Free;
    END;

// Lecture Quantitatif

TOB_ACTIVE := TOBTarfFam;
Select := 'SELECT * FROM TARIF WHERE ' +
          WhereTarifTiers (CodeCatTiers, CodeDevise, '', ttdCatFam, False) +
          ' ORDER BY GF_DEPOT, GF_TARIFARTICLE, GF_BORNEINF';
Q := OpenSQL(Select,True,-1,'',true) ;
TOBTarfFam.LoadDetailDB('TARIF','','',Q,False) ;
Ferme(Q) ;
TOBTarfFam.SetAllModifie(False);
END ;

procedure TFTarifTiers.ChargeTarif ;
var i_ind : integer;
BEGIN
LoadLesTOB ;
// Affichage
ActiveFAM;
for i_ind:=0 to TOB_ACTIVE.Detail.Count-1 do
    BEGIN
    AfficheLaLigne (i_ind + 1) ;
    END ;
END ;

Function TFTarifTiers.WhereTarifTiers (CodeCatTiers, CodeDevise, CodeArticle : String; ttd : T_TableTarif ;
                                       Totale : boolean ) : String ;
Var St : String ;
BEGIN
St:='' ;
if TOB_ACTIVE = TOBTarfArt then
    St:= ' GF_TARIFTIERS="'+CodeCatTiers+'" AND GF_ARTICLE="'+CodeArticle+'" AND ' +
         ' GF_TARIFARTICLE="" AND GF_QUANTITATIF="-" AND GF_DEVISE="' + CodeDevise + '" ';
if TOB_ACTIVE = TOBTarfQArt then
    St:= ' GF_TARIFTIERS="'+CodeCatTiers+'" AND GF_ARTICLE="'+CodeArticle+'" AND ' +
         ' GF_TARIFARTICLE="" AND GF_QUANTITATIF="X" AND GF_DEVISE="' + CodeDevise + '" ';
if TOB_ACTIVE = TOBTarfFam then
    St:= ' GF_TARIFTIERS="'+CodeCatTiers+'" AND GF_ARTICLE="" AND ' +
         ' GF_TARIFARTICLE<>"" AND GF_QUANTITATIF="-" AND GF_DEVISE="' + CodeDevise + '" ';
Result:=St ;
END ;

{==============================================================================================}
{========================= Manipulation de l'entête Tiers =====================================}
{==============================================================================================}
Procedure TFTarifTiers.TraiterCatTiers ;
Var i_Rep: Integer;
    ioerr : TIOErr;
BEGIN
if GF_TARIFTIERS.Text = '' then
  BEGIN
  InitialiseEntete ;
  InitialiseGrille;
  DepileTOBLigne;
  exit
  END;

AffecteEntete;
// G_ART.Enabled := True ;
// G_QART.Enabled := True ;
G_FAM.Enabled := True ;
if ((CodeCatTiers <> GF_TARIFTIERS.Value) or (GF_DEVISE.Value <> CodeDevise)) then
    BEGIN
    if lModif then
        begin
        i_Rep := QuestionTarifEnCours;
        Case i_Rep of
            mrYes    : BEGIN
                       ioerr := Transactions (ValideTarif, 2);
                       Case ioerr of
                          oeOk      : ;
                          oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
                          oeSaisie  : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
                          END ;
                       CodeCatTiers := GF_TARIFTIERS.Value;
                       CodeDevise := GF_DEVISE.Value;
                       DEV.Code := GF_DEVISE.Value ; GetInfosDevise (DEV) ;
                       Transactions (ChargeTarif, 1) ;
//                       Transactions (ChargeTarifArticle, 1);
                       lModif := False;
                       END;
            mrNo     : BEGIN
                       InitialiseGrille;
                       DepileTOBLigne;
                       CodeCatTiers := GF_TARIFTIERS.Value;
                       CodeDevise := GF_DEVISE.Value;
                       DEV.Code := GF_DEVISE.Value ; GetInfosDevise (DEV) ;
                       Transactions (ChargeTarif, 1) ;
//                       Transactions (ChargeTarifArticle, 1);
                       lModif := False;
                       END;
            mrCancel : BEGIN
                       TOBCatTiers.SelectDB ('"' + CodeCatTiers + '"', Nil) ;
                       GF_DEVISE.Value := CodeDevise;
                       AffecteEntete;
                       END;
            END ;
        end else
        begin
        InitialiseGrille;
        DepileTOBLigne;
        CodeCatTiers := GF_TARIFTIERS.Value;
        CodeDevise := GF_DEVISE.Value;
        DEV.Code := GF_DEVISE.Value ; GetInfosDevise (DEV) ;
        Transactions (ChargeTarif, 1) ;
//        Transactions (ChargeTarifArticle, 1);
        end;
        GF_ARTICLE.Enabled := True;
        CodeArticle := '';
        GF_ARTICLE.Text := '';
        G_ART.Enabled := False;
        G_QART.Enabled := False;
    END;
END;

Procedure TFTarifTiers.InitialiseEntete ;
BEGIN
GF_DEVISE.Value := CodeDevise;
TOBCatTiers.InitValeurs ;
AffecteEntete;
CodeCatTiers := '';
CodeArticle := '';
GF_TARIFTIERS.Value := '' ;
GF_TARIFTIERS.SetFocus ;
GF_ARTICLE.Text := '';
GF_ARTICLE.Enabled := False;
G_ART.Enabled := False ; G_QART.Enabled := False ; G_FAM.Enabled := False ;
BCondAplli.Enabled := False ; BCopierCond.Enabled := False; BCollerCond.Enabled := False ;
CONDAPPLIC.Text := '' ;
InitialisePied ;
InitialiseEnteteArticle;
if TarifTTC then
begin
  CBArticle.Checked := True;
  CBArticle.Enabled := False;
end;
END;

Procedure TFTarifTiers.AffecteEntete ;
BEGIN
GF_TARIFTIERS.Text := TOBCatTiers.GetValue ('CC_CODE');
TGF_LIBELLE.Caption := TOBCatTiers.GetValue ('CC_LIBELLE');
END;

Function TFTarifTiers.QuestionTarifEnCours : Integer;
BEGIN
Result := mrNo;
if (TOBTarfArt.IsOneModifie) or (TOBTarfQArt.IsOneModifie) or
   (TOBTarfFam.IsOneModifie) then
    BEGIN
    Result := MsgBox.Execute (0, Caption, '');
    END;
END;

{==============================================================================================}
{========================= Manipulation de l'entête Article =====================================}
{==============================================================================================}
Procedure TFTarifTiers.TraiterArticle ;
Var RechArt : T_RechArt ;
    OkArt   : Boolean ;
    i_Rep: Integer;
    ioerr : TIOErr;
    CodeArt : string;
BEGIN
OkArt:=False ;
CodeArt := CodeArticle ;
if GF_ARTICLE.Text <> '' then
    BEGIN
    RechArt := TrouverArticle (GF_ARTICLE, TOBArt);
    Case RechArt of
            traOk : OkArt:=True ;
         traAucun : BEGIN
                    // Recherche sur code via LookUp ou Recherche avancée
                    DispatchRechArt (GF_ARTICLE, 1, '',
                                   'GA_CODEARTICLE=' + Trim (Copy (GF_ARTICLE.Text, 1, 18)), '');
                    if GF_ARTICLE.Text <> '' then
                        BEGIN
                        GF_ARTICLE.Text:=Format('%-33.33sX',[GF_ARTICLE.Text]);
                        Okart := TOBArt.SelectDB ('"' + GF_ARTICLE.Text + '"', nil);
                        END;
                    END ;
        traGrille : BEGIN
                    // Forcement objet dimension avec saisie obligatoire
                    //if GetArticleLookUp (GF_ARTICLE, 'GA_STATUTART = "DIM"') then
                    if ChoisirDimension (TOBArt.GetValue ('GA_ARTICLE')) then
                        BEGIN
                        Okart := True;
                        END;
                    END;
        End ; // Case
    END;

if (Okart) then
    BEGIN
    AffecteEnteteArticle;
    G_ART.Enabled := True ; G_QART.Enabled := True ;
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
    END else
    BEGIN
    InitialiseEnteteArticle
    END ;
if CodeArt <> TOBArt.GetValue ('GA_ARTICLE') then
    BEGIN
    if cbValider.Checked then i_Rep := mrYes else i_Rep := QuestionTarifArticleEnCours;
    Case i_Rep of
        mrYes    : BEGIN
                   ioerr := Transactions (ValideTarifArticle, 2);
                   Case ioerr of
                      oeOk      : ;
                      oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
                      oeSaisie  : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
                      END ;
                   CodeArticle := TOBArt.GetValue ('GA_ARTICLE');
                   Transactions (ChargeTarifArticle, 1) ;
                   END;
        mrNo     : BEGIN
                   TOBTarfArt.SetAllModifie(False) ;
                   TOBTarfQArt.SetAllModifie(False) ;
                   InitialiseGrilleArticle;
                   DepileTOBLigneArticle;
                   CodeArticle := TOBArt.GetValue ('GA_ARTICLE');
                   Transactions (ChargeTarifArticle, 1) ;
                   END;
        mrCancel : BEGIN
                   CodeArticle:=CodeArt ;
                   if TOBArt.SelectDB ('"' + CodeArticle + '"', Nil) then
                      BEGIN
                      AffecteEnteteArticle;
                      G_ART.Enabled := True ; G_QART.Enabled := True ;
                      END;
                   END;
        END ;
    END;
END;

Function TFTarifTiers.QuestionTarifArticleEnCours : Integer;
BEGIN
Result := mrNo;
if (TOBTarfArt.IsOneModifie) or (TOBTarfQArt.IsOneModifie) then
    BEGIN
    Result := MsgBox.Execute (0, Caption, '');
    END;
END;

Procedure TFTarifTiers.InitialiseEnteteArticle ;
BEGIN
TOBArt.InitValeurs ; PrixPourQte:=1;
AffecteEnteteArticle;
CodeArticle := '';
if GF_ARTICLE.Enabled then GF_ARTICLE.SetFocus ;
G_ART.Enabled := False ; G_QART.Enabled := False ;
GA_PRIXPOURQTE.Value:=0;
GA_PRIXPOURQTE.Visible:=False;
TGA_PRIXPOURQTE.Visible:=False;
END;

Procedure TFTarifTiers.AffecteEnteteArticle ;
BEGIN
GF_ARTICLE.Text := TOBArt.GetValue ('GA_CODEARTICLE');
TGF_LIBARTICLE.Caption := TOBArt.GetValue ('GA_LIBELLE');
// Dimensions
AffecteDimension;
END;

Procedure TFTarifTiers.ChercherArticle ;
BEGIN
DispatchRechArt (GF_ARTICLE, 1, '',
                 'GA_CODEARTICLE=' + Trim (Copy (GF_ARTICLE.Text, 1, 18)), '');
if GF_ARTICLE.Text <> '' then
    BEGIN
    GF_ARTICLE.Text:=Format('%-33.33sX',[GF_ARTICLE.Text]);
    TOBArt.SelectDB ('"' + GF_ARTICLE.Text + '"', Nil) ;
    GF_ARTICLE.Text := TOBArt.GetValue ('GA_CODEARTICLE') ;
    TraiterArticle;
    END;
END;

procedure TFTarifTiers.VerifLesTOBArticle;
var i_ind : integer;
    Q : TQuery ;
    MaxTarif : Longint ;
BEGIN
Q := OpenSQL ('SELECT MAX(GF_TARIF) FROM TARIF', TRUE,-1,'',true) ;
if Q.EOF then MaxTarif := 1 else MaxTarif := Q.Fields[0].AsInteger + 1 ;
Ferme(Q) ;
//  reinit des variables ACTIVES pour la fonction LigneVide
ActiveART;
for i_ind := TOBTarfArt.Detail.count - 1 Downto 0 do
    BEGIN
    if LigneVide (i_ind + 1) then
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
//  reinit des variables ACTIVES pour la fonction LigneVide
ActiveQART;
for i_ind := TOBTarfQArt.Detail.count - 1 Downto 0 do
    BEGIN
    if LigneVide (i_ind + 1) then
        BEGIN
        TOBTarfQArt.Detail[i_ind].Free ;
        END else
        BEGIN
        if TOBTarfQArt.Detail[i_ind].GetValue ('GF_TARIF') = 0 then
            BEGIN
            TOBTarfQArt.Detail[i_ind].PutValue ('GF_TARIF', MaxTarif);
            Inc (MaxTarif);
            END;
        CalcPriorite (TOBTarfQArt.Detail[i_ind]);
        END;
    END;
END;

procedure TFTarifTiers.ChargeTarifArticle ;
var i_ind : integer;
BEGIN
LoadLesTOBArticle ;
// Affichage
ActiveART;
for i_ind:=0 to TOB_ACTIVE.Detail.Count-1 do
    BEGIN
    AfficheLaLigne (i_ind + 1) ;
    END ;
ActiveQART;
for i_ind:=0 to TOB_ACTIVE.Detail.Count-1 do
    BEGIN
    AfficheLaLigne (i_ind + 1) ;
    END ;
END ;

procedure TFTarifTiers.LoadLesTOBArticle ;
Var Q : TQuery ;
    i_ind : integer;
    Select, WhereTTC : string;
BEGIN
for i_ind := TOBTarfArt.Detail.Count - 1 downto 0 do
    BEGIN
    TOBTarfArt.Detail[i_ind].Free;
    END;
for i_ind := TOBTarfQArt.Detail.Count - 1 downto 0 do
    BEGIN
    TOBTarfQArt.Detail[i_ind].Free;
    END;

if TarifTTC then WhereTTC := ' AND GF_REGIMEPRIX = "TTC" '
            else WhereTTC := ' AND GF_REGIMEPRIX <> "TTC" ' ;
// Lecture Quantitatif

TOB_ACTIVE := TOBTarfArt;
Select := 'SELECT * FROM TARIF WHERE ' +
          WhereTarifTiers (CodeCatTiers, CodeDevise, CodeArticle, ttdCatArt, False) +
          WhereTTC + ' ORDER BY GF_DEPOT, GF_DATEDEBUT';
Q := OpenSQL(Select,True,-1,'',true) ;
TOBTarfArt.LoadDetailDB('TARIF','','',Q,False) ;
Ferme(Q) ;
TOBTarfArt.SetAllModifie(False);
TOB_ACTIVE := TOBTarfQArt;
Select := 'SELECT * FROM TARIF WHERE ' +
          WhereTarifTiers (CodeCatTiers, CodeDevise, CodeArticle, ttdCatQArt, False) +
          WhereTTC + ' ORDER BY GF_DEPOT, GF_BORNEINF';
Q := OpenSQL(Select,True,-1,'',true) ;
TOBTarfQArt.LoadDetailDB('TARIF','','',Q,False) ;
Ferme(Q) ;
TOBTarfQArt.SetAllModifie(False);
END ;

Procedure TFTarifTiers.AffecteDimension ;
var TLIB, TLIB2 : THLabel;
    CHPS, CHPS2 : THCritMaskEdit;
    i_ind, i_Dim : integer;
    b_Dim : Boolean ;
    GrilleDim,CodeDim,LibDim : String ;
BEGIN
i_Dim := 1;
b_Dim := False;
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

Function TFTarifTiers.ChoisirDimension (St_CodeArt : string) : Boolean;
var QQ   :TQuery;
BEGIN
//TheTOB:=TOB.Create('ARTICLE', nil, -1);
//TheTOB.Dupliquer(TOBArt,True,True);
TheTOB:=TOB.Create('', nil, -1);
AglLanceFiche ('GC','GCSELECTDIM','','', 'GA_ARTICLE='+ St_CodeArt +
	           ';ACTION=SELECT;CHAMP= ') ;
if TheTOB = Nil then
    BEGIN
    Result := False;
    END else
    BEGIN
    Result := True;
    QQ := OpenSQL('Select * from ARTICLE Where GA_ARTICLE="' +
             TheTOB.Detail[0].GetValue ('GA_ARTICLE') +'" ',True,-1,'',true) ;
    if Not QQ.EOF then
        BEGIN
        TOBArt.SelectDB ('', QQ);
        Ferme(QQ) ;
        END;
    TheTOB := Nil;
    END;
END;

//=============================================================================
procedure TFTarifTiers.ValideTarif;
begin
if Not SortDeLaLigne(ttdCatArt) then Exit ;
if Not SortDeLaLigne(ttdCatQArt) then Exit ;
if Not SortDeLaLigne(ttdCatFam) then Exit ;
TOBTarifDel.DeleteDB (False);
VerifLesTOB;
TOBTarif.InsertOrUpdateDB(False) ;
InitialiseGrille;
DepileTOBLigne;
end;

procedure TFTarifTiers.VerifLesTOB;
var i_ind : integer;
    Q : TQuery ;
    MaxTarif : Longint ;
BEGIN
Q := OpenSQL ('SELECT MAX(GF_TARIF) FROM TARIF', TRUE,-1,'',true) ;
if Q.EOF then MaxTarif := 1 else MaxTarif := Q.Fields[0].AsInteger + 1 ;
Ferme(Q) ;
ActiveFAM;
for i_ind := TOBTarfFam.Detail.count - 1 Downto 0 do
    BEGIN
    if LigneVide (i_ind + 1) then
        BEGIN
        TOBTarfFam.Detail[i_ind].Free ;
        END else
        BEGIN
        if TOBTarfFam.Detail[i_ind].GetValue ('GF_ARTICLE') <> '' then
            TOBTarfFam.Detail[i_ind].PutValue ('GF_ARTICLE', '');
        if TOBTarfFam.Detail[i_ind].GetValue ('GF_TARIF') = 0 then
            BEGIN
            TOBTarfFam.Detail[i_ind].PutValue ('GF_TARIF', MaxTarif);
            Inc (MaxTarif);
            END;
        CalcPriorite (TOBTarfFam.Detail[i_ind]);
        END;
    END;
VerifLesTOBArticle;
END;

procedure TFTarifTiers.ValideTarifArticle;
begin
if Not SortDeLaLigne(ttdCatArt) then Exit ;
if Not SortDeLaLigne(ttdCatQArt) then Exit ;
TOBTarifDel.DeleteDB (False);
VerifLesTOBArticle;
TOBTarfArt.InsertOrUpdateDB(False) ;
TOBTarfQArt.InsertOrUpdateDB(False) ;
InitialiseGrilleArticle;
DepileTOBLigneArticle;
end;

{==============================================================================================}
{===================== Manipulation des Champs LIGNES Quantitatif =============================}
{==============================================================================================}
procedure TFTarifTiers.TraiterDepot (ACol, ARow : integer);
var TOBL : TOB;
    St : string;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
St := G_ACTIVE.Cells [ACol, ARow];
if ExisteDepot ('GCDEPOT', St) then
    BEGIN
    TOBL.PutValue ('GF_DEPOT', St);
    lModif := True;
    END else
    BEGIN
    // message dépôt inexistant
    MsgBox.Execute (4,Caption,'') ;
    G_ACTIVE.Cells [ACol, ARow] := TOBL.GetValue ('GF_DEPOT');
    END;
END;

procedure TFTarifTiers.TraiterCatArticle (ACol, ARow : integer);
var TOBL : TOB;
    B_NewLine : Boolean;
    St : string;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
if TOBL.GetValue ('GF_TARIFARTICLE') = '' then B_NewLine :=True else B_NewLine := False;
St := G_ACTIVE.Cells [ACol, ARow];
if St <> '' then
    BEGIN
    if ExisteCategorie ('GCTARIFARTICLE', St) then
        BEGIN
       TOBL.PutValue ('GF_TARIFARTICLE', St);
        lModif := True;
        END else
        BEGIN
        // message Catégorie Tiers inexistante
        MsgBox.Execute (4,Caption,'') ;
        G_ACTIVE.Cells [ACol, ARow] := TOBL.GetValue ('GF_TARIFARTICLE');
        END;
    if (TOBL.GetValue ('GF_TARIFARTICLE') <> '') AND (B_NewLine) then
        BEGIN
        InitLaLigne (ARow);
        END;
    END else if Not B_Newline then
            G_ACTIVE.Cells [ACol, ARow] := TOBL.GetValue ('GF_TARIFARTICLE');
END;

procedure TFTarifTiers.TraiterLibelle (ACol, ARow : integer);
var TOBL : TOB;
    B_NewLine : Boolean;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
if TOBL.GetValue ('GF_LIBELLE') = '' then B_NewLine :=True else B_NewLine := False;
if G_ACTIVE.Cells [ACol, ARow] <> '' then
    BEGIN
    TOBL.PutValue ('GF_LIBELLE', G_ACTIVE.Cells [ACol, ARow]);
    lModif := True;
    if (TOBL.GetValue ('GF_LIBELLE') <> '') AND (B_NewLine) then
        BEGIN
        if (G_ACTIVE = G_ART) or (G_ACTIVE = G_QART) or
           ((G_ACTIVE = G_FAM) and (TOBL.GetValue ('GF_TARIFTIERS') <> '')) then
           InitLaLigne (ARow);
        end;
    END else if Not B_Newline then
        G_ACTIVE.Cells [ACol, ARow] := TOBL.GetValue ('GF_LIBELLE');
END;

procedure TFTarifTiers.TraiterBorneInf (ACol, ARow : integer);
var TOBL : TOB;
    f_QteInf : Extended;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
f_QteInf := Valeur (G_ACTIVE.Cells [ACol, ARow]);
if f_QteInf > TOBL.GetValue ('GF_BORNESUP') then
    BEGIN
    TOBL.PutValue ('GF_BORNESUP', f_QteInf);
    lModif := True;
    G_ACTIVE.Cells [ACT_QSup, ARow] := floattostr (TOBL.GetValue ('GF_BORNESUP'));
    END;
if f_QteInf < 1 then
    BEGIN
    MsgBox.Execute (2,Caption,'') ;
    G_ACTIVE.Cells [ACol, ARow] := floattostr (TOBL.GetValue ('GF_BORNEINF'));
    G_ACTIVE.Col := ACol; G_ACTIVE.Row := ARow;
    END else
    begin
    TOBL.PutValue ('GF_BORNEINF', f_QteInf);
    lModif := True;
    end;
END;

procedure TFTarifTiers.TraiterBorneSup (ACol, ARow : integer);
var TOBL : TOB;
    f_QteSup : Extended;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
f_QteSup := Valeur (G_ACTIVE.Cells [ACol, ARow]);
if f_QteSup < TOBL.GetValue ('GF_BORNEINF') then
    BEGIN
    MsgBox.Execute (2,Caption,'') ;
    G_ACTIVE.Cells [ACol, ARow] := floattostr (TOBL.GetValue ('GF_BORNESUP'));
    G_ACTIVE.Col := ACol; G_ACTIVE.Row := ARow;
    END else
    BEGIN
    TOBL.PutValue ('GF_BORNESUP', f_QteSup);
    lModif := True;
    END;
END;

procedure TFTarifTiers.TraiterPrix (ACol, ARow : integer);
var TOBL : TOB;
    PuFix : double;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
PuFix := Valeur (G_ACTIVE.Cells [ACol, ARow]);
{$IFDEF MONCODE}
if TOBArt.GetValue('GA_DECIMALPRIX')='X' then
  begin
  if PrixPourQte<=0 then PrixPourQte:=1;
  PuFix:=PuFix*PrixPourQte;
  end;
{$ENDIF}
TOBL.PutValue ('GF_PRIXUNITAIRE',PuFix);
END;

procedure TFTarifTiers.TraiterRemise(ACol, ARow : integer);
var TOBL : TOB;
    St : string;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
G_ACTIVE.Cells [ACol, ARow]:=ModifFormat(G_ACTIVE.Cells [ACol, ARow]);
St := G_ACTIVE.Cells [ACol, ARow];
TOBL.PutValue ('GF_CALCULREMISE', St);
TOBL.PutValue ('GF_REMISE', RemiseResultante (St));
lModif := True;
AffichePied (ARow);
END;

procedure TFTarifTiers.TraiterDateDeb (ACol, ARow : integer);
var TOBL : TOB;
    St_Date : string;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
St_Date := G_ACTIVE.Cells [ACol, ARow] ;
if IsValidDate (st_Date) then
    BEGIN
    if StrToDate (St_Date) > TOBL.GetValue ('GF_DATEFIN') then
        BEGIN
            MsgBox.Execute (3,Caption,'') ;
            G_ACTIVE.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
            G_ACTIVE.Col := ACol; G_ACTIVE.Row := ARow;
        END else
        BEGIN
        TOBL.PutValue ('GF_DATEDEBUT', StrToDate (St_Date));
        lModif := True;
        END;
    END else
    BEGIN
    if TOBL.GetValue ('GF_LIBELLE') <> '' then
        G_ACTIVE.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
    END;
END;

procedure TFTarifTiers.TraiterDateFin (ACol, ARow : integer);
var TOBL : TOB;
    St_Date : string;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
St_Date := G_ACTIVE.Cells [ACol, ARow] ;
if IsValidDate (st_Date) then
    BEGIN
    if StrToDate (St_Date) < TOBL.GetValue ('GF_DATEDEBUT') then
        BEGIN
            MsgBox.Execute (3,Caption,'') ;
            G_ACTIVE.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
            G_ACTIVE.Col := ACol; G_ACTIVE.Row := ARow;
        END else
        BEGIN
        TOBL.PutValue ('GF_DATEFIN', StrToDate (St_Date));
        lModif := True;
        END;
    END else
    BEGIN
    if TOBL.GetValue ('GF_LIBELLE') <> '' then
        G_ACTIVE.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
    END;
END;

//==============================================================================
//  Evenements génériques sur HGrids
//==============================================================================
//==============================================================================
//  Evenement OnEnter
//==============================================================================
procedure TFTarifTiers.G_ACTIVEEnter(Sender: TObject);
var Cancel, Chg : Boolean;
    ACol, ARow : integer;
BEGIN
Cancel := False; Chg := False;
G_ACTIVERowEnter (Sender, G_ACTIVE.Row, Cancel, Chg);
ACol := G_ACTIVE.Col ; ARow := G_ACTIVE.Row ;
G_ACTIVECellEnter (Sender, ACol, ARow, Cancel);
end;

//==============================================================================
//  Evenement OnRowEnter
//==============================================================================
procedure TFTarifTiers.G_ACTIVERowEnter(Sender: TObject; Ou: Integer;
                                        var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
BEGIN
if Ou >= G_ACTIVE.RowCount - 1 then G_ACTIVE.RowCount := G_ACTIVE.RowCount + NbRowsPlus ;
ARow := Min (Ou, TOB_ACTIVE.detail.count + 1);
if (ARow = TOB_ACTIVE.detail.count + 1) AND (not LigneVide (ARow - 1)) then
    BEGIN
    CreerTOBligne (ARow);
    END;
if (LigneVide (ARow)) AND (not LigneVide (ARow - 1))then
    PreAffecteLigne (ARow);
if Ou > TOB_ACTIVE.detail.count then
    BEGIN
    G_ACTIVE.Row := TOB_ACTIVE.detail.count;
    END;
AffichePied (ARow);
AfficheCondTarf (ARow);
end;

//==============================================================================
//  Evenement OnRowExit
//==============================================================================
procedure TFTarifTiers.G_ACTIVERowExit(Sender: TObject; Ou: Integer;
                                       var Cancel: Boolean; Chg: Boolean);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if LigneVide (Ou) Then G_ACTIVE.Row := Min (G_ACTIVE.Row,Ou);
end;

//==============================================================================
//  Evenement OnCellEnter
//==============================================================================
procedure TFTarifTiers.G_ACTIVECellEnter(Sender: TObject; var ACol,
                                         ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if Not Cancel then
    BEGIN
    if (G_ACTIVE.Col <> ACT_Depot) AND (G_ACTIVE.Col <> ACT_Lib) then
        begin
        if (G_ACTIVE.Cells [ACT_Lib,G_ACTIVE.Row] = '') then
            G_ACTIVE.Col := ACT_Lib;
        if (G_ACTIVE = G_FAM) and (G_ACTIVE.Col <> ACT_Cat) then
            if (G_ACTIVE.Cells [ACT_Cat,G_ACTIVE.Row] = '') then G_ACTIVE.Col := ACT_Cat;
        end;
    G_ACTIVE.ElipsisButton := ((G_ACTIVE.Col = ACT_Depot) or (G_ACTIVE.col = ACT_Datedeb) or
                               (G_ACTIVE.col = ACT_Datefin) or (G_ACTIVE.col = ACT_Cat)) ;
    StCellCur := G_ACTIVE.Cells [G_ACTIVE.Col,G_ACTIVE.Row] ;
    AffectMenuCondApplic (G_ACTIVE);
    END ;
end;

//==============================================================================
//  Evenement OnCellExit
//==============================================================================
procedure TFTarifTiers.G_ACTIVECellExit(Sender: TObject; var ACol,
                                        ARow: Integer; var Cancel: Boolean);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
FormateZoneSaisie (ACol,ARow);
if ACol = ACT_Depot then TraiterDepot (ACol, ARow) else
    if ACol = ACT_Cat then TraiterCatArticle (ACol, ARow) else
        if ACol = ACT_Lib then TraiterLibelle (ACol, ARow) else
            if ACol = ACT_QInf then TraiterBorneInf (ACol, ARow) else
                if ACol = ACT_QSup then TraiterBorneSup (ACol, ARow) else
                    if ACol = ACT_Px then TraiterPrix (ACol, ARow) else
                        if ACol = ACT_Rem then TraiterRemise (ACol, ARow) else
                            if ACol = ACT_Datedeb then TraiterDateDeb (ACol, ARow) else
                                if ACol = ACT_Datefin then TraiterDateFin (ACol, ARow);
if Not Cancel then
    BEGIN
    END;
end;

procedure TFTarifTiers.G_ACTIVEElipsisClick(Sender: TObject);
Var DEPOT, CAT, DATE : THCritMaskEdit;
    Coord : TRect;
begin
if (G_ACTIVE.Col = ACT_Depot) then
    BEGIN
    Coord := G_ACTIVE.CellRect (G_ACTIVE.Col, G_ACTIVE.Row);
    DEPOT := THCritMaskEdit.Create (Self);
    DEPOT.Parent := G_ACTIVE;
    DEPOT.Top := Coord.Top;
    DEPOT.Left := Coord.Left;
    DEPOT.Width := 3; DEPOT.Visible := False;
    DEPOT.DataType := 'GCDEPOT';
    GetDepotRecherche (DEPOT) ;
    if DEPOT.Text <> '' then G_ACTIVE.Cells[G_ACTIVE.Col,G_ACTIVE.Row]:= DEPOT.Text;
    DEPOT.Destroy;
    END ;
if (G_ACTIVE.Col = ACT_Datedeb) or (G_ACTIVE.Col = ACT_Datefin) then
    BEGIN
    Coord := G_ACTIVE.CellRect (G_ACTIVE.Col, G_ACTIVE.Row);
    DATE := THCritMaskEdit.Create (Self);
    DATE.Parent := G_ACTIVE;
    DATE.Top := Coord.Top;
    DATE.left := Coord.Left;
    DATE.Width := 3; DATE.Visible := False;
    DATE.OpeType:=otDate;
    GetDateRecherche (TForm(DATE.Owner), DATE) ;
    if DATE.Text <> '' then G_ACTIVE.Cells[G_ACTIVE.Col,G_ACTIVE.Row]:= DATE.Text;
    DATE.Destroy;
    END;
if (G_ACTIVE.Col = ACT_Cat) then
    BEGIN
    Coord := G_ACTIVE.CellRect (G_ACTIVE.Col, G_ACTIVE.Row);
    CAT := THCritMaskEdit.Create (Self);
    CAT.Parent := G_ACTIVE;
    CAT.Top := Coord.Top;
    CAT.Left := Coord.Left;
    CAT.Width := 3; CAT.Visible := False;
    CAT.DataType := 'GCTARIFARTICLE';
    GetCategorieRecherche (CAT) ;
    if CAT.Text <> '' then G_ACTIVE.Cells[G_ACTIVE.Col,G_ACTIVE.Row]:= CAT.Text;
    CAT.Destroy;
    END ;
end;

//==============================================================================
//  Evenements sur GF_TARIFTIERS
//==============================================================================
procedure TFTarifTiers.GF_TARIFTIERSChange(Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if GF_TARIFTIERS.Text <> TOBCatTiers.GetValue ('CC_CODE') then TraiterCatTiers ;
end;

//==============================================================================
//  Evenements sur GF_DEVISE
//==============================================================================
procedure TFTarifTiers.GF_DEVISEChange(Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if GF_DEVISE.Value <> CodeDevise then TraiterCatTiers ;
InitialisePied ;
end;

//==============================================================================
//  Evenements sur GF_ARTICLE
//==============================================================================
procedure TFTarifTiers.GF_ARTICLEExit(Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if GF_ARTICLE.Text <> TOBArt.GetValue('GA_CODEARTICLE') then TraiterArticle ;
end;

procedure TFTarifTiers.GF_ARTICLEElipsisClick(Sender: TObject);
begin
ChercherArticle;
end;

procedure TFTarifTiers.GF_ARTICLEChange(Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
//if GF_ARTICLE.Text <> TOBArt.GetValue('GA_CODEARTICLE') then TraiterArticle ;
end;

//==============================================================================
//  Evenements sur G_ART
//==============================================================================
procedure TFTarifTiers.G_ARTEnter(Sender: TObject);
BEGIN
ActiveART;
G_ACTIVEEnter(Sender);
end;

//==============================================================================
//  Evenements sur G_FAM
//==============================================================================
procedure TFTarifTiers.G_FAMEnter(Sender: TObject);
BEGIN
ActiveFAM;
G_ACTIVEEnter(Sender);
end;

//==============================================================================
//  Evenements sur G_QART
//==============================================================================
procedure TFTarifTiers.G_QARTEnter(Sender: TObject);
begin
ActiveQART;
G_ACTIVEEnter(Sender);
end;

//==============================================================================
//  Gestion des clicks boutons
//==============================================================================

procedure TFTarifTiers.BSaisieRapideClick(Sender: TObject);
var i_ind : integer;
begin
if Action = taConsult then exit;
if GF_ARTICLE.Text = '' then exit;
if (PTARIFART.Visible) and (PQTEART.Visible) then
    begin
    ActiveQArt;
    EntreeTarifRapide (taModif, TOBart, CodeCatTiers, '', CodeDevise, ttdTiersQArt, TOBTarfQArt, tstTiers);
    // TobTarfQArt.PutGridDetail (G_QART, True, True, LesColQArt); des que ca marche avec les dates !!
    for i_ind:=0 to TOBTarfQArt.Detail.Count-1 do
        BEGIN
        AfficheLaLigne (i_ind + 1) ;
        END ;
    end;
end;

procedure TFTarifTiers.BValiderClick(Sender: TObject);
Var ioerr : TIOErr ;
begin
ioerr := Transactions (ValideTarif, 2);
Case ioerr of
        oeOk : ;
   oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
    oeSaisie : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
   END ;
InitialiseEntete;
end;

{==============================================================================================}
{============================ Evenement lié aux Boutons =======================================}
{==============================================================================================}
procedure TFTarifTiers.BChercherClick(Sender: TObject);
begin
if PTARIFART.Visible then
    BEGIN
    if PQTEART.Visible then
        if G_QART.RowCount < 3 then Exit;
    if PTART.Visible then
        if G_ART.RowCount < 3 then Exit;
    END;
if PFAMARTICLE.Visible then
    if G_FAM.RowCount < 3 then Exit;
FindDebut:=True ; FindLigne.Execute ;
end;

procedure TFTarifTiers.FindLigneFind(Sender: TObject);
begin
if PTARIFART.Visible then
    BEGIN
    if PQTEART.Visible then
        Rechercher (G_QART, FindLigne, FindDebut) ;
    if PTART.Visible then
        Rechercher (G_ART, FindLigne, FindDebut) ;
    END;
if PFAMARTICLE.Visible then
    Rechercher (G_FAM, FindLigne, FindDebut) ;
end;

procedure TFTarifTiers.InfArticleClick(Sender: TObject);
begin
if GF_ARTICLE.Text = '' then Exit;
VoirFicheArticle;
end;

procedure TFTarifTiers.CBQUANTITATIFClick(Sender: TObject);
begin
InitialisePied ;
if CBQUANTITATIF.Checked then
    BEGIN
    PQTEART.Visible := True ;
    PTART.Visible := False ;
    PTITRE.Caption := HTitre.Mess[1] ;
    ActiveQART;
    AffectMenuCondApplic (G_ACTIVE);
    END else
    BEGIN
    PQTEART.Visible := False ;
    PTART.Visible := True ;
    PTITRE.Caption := HTitre.Mess[2] ;
    ActiveART;
    AffectMenuCondApplic (G_ACTIVE);
    END;
HMTrad.ResizeGridColumns (G_ART) ;
HMTrad.ResizeGridColumns (G_FAM) ;
HMTrad.ResizeGridColumns (G_QART) ;
if TarifTTC then TTYPETARIF.Caption := HTitre.Mess[6]
            else TTYPETARIF.Caption := HTitre.Mess[5] ;
end;

procedure TFTarifTiers.CBARTICLEClick(Sender: TObject);
begin
InitialisePied ;
if CBARTICLE.Checked then
    BEGIN
    PFAMARTICLE.Visible := False;
    PTARIFART.Visible := True;
    if PrixPourQte>1 then
       begin
       GA_PRIXPOURQTE.Visible:=True;
       TGA_PRIXPOURQTE.Visible:=True;
       end;
    if CBQUANTITATIF.Checked then
        BEGIN
        PQTEART.Visible := True ;
        PTART.Visible := False ;
        PTITRE.Caption := HTitre.Mess[1] ;
        ActiveQART;
        AffectMenuCondApplic (G_ACTIVE);
        END else
        BEGIN
        PQTEART.Visible := False ;
        PTART.Visible := True ;
        PTITRE.Caption := HTitre.Mess[2] ;
        ActiveART;
        AffectMenuCondApplic (G_ACTIVE);
        END;
    if TarifTTC then TTYPETARIF.Caption := HTitre.Mess[6]
                else TTYPETARIF.Caption := HTitre.Mess[5] ;
    END else
    BEGIN
    PFAMARTICLE.Visible := True;
    PTARIFART.Visible := False;
    PQTEART.Visible := False ;
    PTART.Visible := False ;
    GA_PRIXPOURQTE.Visible:=False;
    TGA_PRIXPOURQTE.Visible:=False;
    PTITRE.Caption := HTitre.Mess[0] ;
    TTYPETARIF.Caption := HTitre.Mess[4] ;
    ActiveFAM;
    AffectMenuCondApplic (G_ACTIVE);
    END;
HMTrad.ResizeGridColumns (G_ART) ;
HMTrad.ResizeGridColumns (G_FAM) ;
HMTrad.ResizeGridColumns (G_QART) ;
AffichePrixCon ;
cbValider.Visible := CBARTICLE.Checked;
end;

procedure TFTarifTiers.BVoirCondClick(Sender: TObject);
begin
TCONDTARF.Visible := BVoirCond.Down ;
end;

procedure TFTarifTiers.BCondAplliClick(Sender: TObject);
Var TOBL : TOB;
begin
{$IFNDEF CCS3}
if GF_TARIFTIERS.Value = '' then exit;
if PTARIFART.Visible then
    BEGIN
    if GF_ARTICLE.Text = '' then Exit;
    if PQTEART.Visible then
        BEGIN
        ActiveQART;
        TOBL:=GetTOBLigne(G_ACTIVE.Row) ; if TOBL=Nil then Exit ;
        EntreeTarifCond (Action, TOBL);
        AfficheCondTarf (G_ACTIVE.Row);
        END;
    if PTART.Visible then
        BEGIN
        ActiveART;
        TOBL:=GetTOBLigne(G_ACTIVE.Row) ; if TOBL=Nil then Exit ;
        EntreeTarifCond (Action, TOBL);
        AfficheCondTarf (G_ACTIVE.Row);
        END;
    END;
if PFAMARTICLE.Visible then
    BEGIN
    ActiveFAM;
    TOBL:=GetTOBLigne(G_ACTIVE.Row) ; if TOBL=Nil then Exit ;
    EntreeTarifCond (Action, TOBL);
    AfficheCondTarf (G_ACTIVE.Row);
    END;
{$ENDIF}
end;

procedure TFTarifTiers.BCopierCondClick(Sender: TObject);
Var TOBL : TOB;
begin
if GF_TARIFTIERS.Value = '' then exit;
if PTARIFART.Visible then
    BEGIN
    if GF_ARTICLE.Text = '' then Exit;
    if PQTEART.Visible then
        BEGIN
        ActiveQART;
        TOBL:=GetTOBLigne(G_ACTIVE.Row) ; if TOBL=Nil then Exit ;
        CONDAPPLIC.Text := TOBL.GetValue ('GF_CONDAPPLIC') ;
        END;
    if PTART.Visible then
        BEGIN
        ActiveART;
        TOBL:=GetTOBLigne(G_ACTIVE.Row) ; if TOBL=Nil then Exit ;
        CONDAPPLIC.Text := TOBL.GetValue ('GF_CONDAPPLIC') ;
        END;
    END;
if PFAMARTICLE.Visible then
    BEGIN
    ActiveFAM;
    TOBL:=GetTOBLigne(G_ACTIVE.Row) ; if TOBL=Nil then Exit ;
    CONDAPPLIC.Text := TOBL.GetValue ('GF_CONDAPPLIC') ;
    END;
end;

procedure TFTarifTiers.BCollerCondClick(Sender: TObject);
Var TOBL : TOB;
begin
if GF_TARIFTIERS.Value = '' then exit;
if PTARIFART.Visible then
    BEGIN
    if GF_ARTICLE.Text = '' then Exit;
    if PQTEART.Visible then
        BEGIN
        ActiveQART;
        TOBL:=GetTOBLigne(G_ACTIVE.Row) ; if TOBL=Nil then Exit ;
        TOBL.PutValue ('GF_CONDAPPLIC', CONDAPPLIC.Text) ;
        AfficheCondTarf (G_ACTIVE.Row);
        END;
    if PTART.Visible then
        BEGIN
        ActiveART;
        TOBL:=GetTOBLigne(G_ACTIVE.Row) ; if TOBL=Nil then Exit ;
        TOBL.PutValue ('GF_CONDAPPLIC', CONDAPPLIC.Text) ;
        AfficheCondTarf (G_ACTIVE.Row);
        END;
    END;
if PFAMARTICLE.Visible then
    BEGIN
    ActiveFAM;
    TOBL:=GetTOBLigne(G_ACTIVE.Row) ; if TOBL=Nil then Exit ;
    TOBL.PutValue ('GF_CONDAPPLIC', CONDAPPLIC.Text) ;
    AfficheCondTarf (G_ACTIVE.Row);
    END;
end;

{==============================================================================================}
{============================== Action lié aux Boutons ========================================}
{==============================================================================================}
procedure TFTarifTiers.VoirFicheArticle;
BEGIN
{$IFDEF BTP}
  V_PGI.dispatchTT (7,taConsult,CodeArticle,'','');
{$ELSE BTP}
  {$IFNDEF GPAO}
	  if ctxAffaire in V_PGI.PGIContexte then V_PGI.dispatchTT (7,taConsult,CodeArticle,'','')
	  else AglLanceFiche ('GC', 'GCARTICLE', '', CodeArticle, 'ACTION=CONSULTATION;TARIF=N');
  {$ELSE GPAO}
    V_PGI.dispatchTT(7, taConsult, CodeArticle, 'TARIF=N', '');
  {$ENDIF GPAO}
{$ENDIF BTP}
END;

{==============================================================================================}
{=============================== Conditions tarifaires ========================================}
{==============================================================================================}
procedure TFTarifTiers.InitComboChamps ;
begin
SourisSablier ;
RemplitComboChamps('TIERS',FComboTIE) ;
RemplitComboChamps('ARTICLE',FComboART) ;
RemplitComboChamps('LIGNE',FComboLIG) ;
RemplitComboChamps('PIECE',FComboPIE) ;
SourisNormale ;
end ;

procedure TFTarifTiers.RemplitComboChamps(NomTable : String ; FCombo : THValComboBox) ;
var Pref : String ;
    Q    : TQuery ;
begin
Pref := TableToPrefixe(NomTable) ;
Q:=OpenSQL('SELECT DH_NOMCHAMP,DH_LIBELLE FROM DECHAMPS WHERE DH_PREFIXE="'+
           Pref + '" AND DH_CONTROLE LIKE "%T%"',True,-1,'',true) ;
While not Q.EOF do
   begin
   if Trim(Q.FindField('DH_LIBELLE').AsString)='' then FCombo.Items.Add(Q.FindField('DH_NOMCHAMP').AsString)
                                                  else FCombo.Items.Add(Q.FindField('DH_LIBELLE').AsString) ;
   FCombo.Values.Add(Q.FindField('DH_NOMCHAMP').AsString) ;
   Q.Next ;
   end ;
Ferme(Q) ;
end ;

procedure TFTarifTiers.EffaceGrid ;
var lig, col : Integer ;
begin
for lig := 1 to G_COND.RowCount - 1 do
   begin
   for col := 0 to G_COND.ColCount - 1 do G_COND.Cells[col, lig] := '' ;
   end ;
G_COND.Row := 1 ;
end ;

Procedure TFTarifTiers.AfficheCondTarf (ARow : Longint) ;
var TOBL : TOB;
BEGIN
TOBL:=GetTOBLigne(ARow) ; if TOBL=Nil then Exit ;
GetConditions (TOBL) ;
END;

function TFTarifTiers.ValueToItem(CC : THValComboBox ; St : String) : String ;
var i_ind : Integer ;
begin
i_ind := CC.Values.IndexOf(St) ;
if i_ind >= 0 then result := CC.Items[i_ind] else result := '' ;
end ;

{Charge dans le grid les conditions stockées dans le TMemoField}
procedure TFTarifTiers.GetConditions (TOBL : TOB) ;
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

procedure TFTarifTiers.TCONDTARFClose(Sender: TObject);
begin
BVoirCond.Down := False ;
end;

Procedure AGLEntreeTarifTiers ( Parms : array of variant ; nb : integer ) ;
Var Action : TActionFiche ;
    St     : String ;
    OkTTC : Boolean ;
BEGIN
Action:=StringToAction(String(Parms[1])) ;
OkTTC:=False ;
St:=String(Parms[2]) ; if St='TTC' then OkTTC:=True ;
EntreeTarifTiers(Action,OkTTC) ;
END ;

procedure TFTarifTiers.BAbandonClick(Sender: TObject);
begin
Close ;
if FClosing and IsInside(Self) then THPanel(parent).CloseInside ;
end;

procedure TFTarifTiers.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self);
end;

Initialization
RegisterAglProc('EntreeTarifTiers',True,2,AGLEntreeTarifTiers) ;

end.

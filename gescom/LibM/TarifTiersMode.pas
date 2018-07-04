{***********UNITE*************************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 30/07/2001
Modifié le ... :   /  /    
Description .. : Spécifique mode: Tarif détail
Suite ........ : Saisie des tarifs clients et catégories client
Mots clefs ... : TARIF;CLIENT
*****************************************************************}
unit TarifTiersMode;                         

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, Grids, Hctrls, ComCtrls, ExtCtrls, HPanel, StdCtrls, Mask, UIUtil,
  Hent1, TarifUtil, SaisUtil, HSysMenu, UTOB, math, hmsgbox, Ent1,
  AglInit, HDimension, Menus, TarifRapide, HRichEdt, HRichOLE,ParamSoc,
{$IFDEF EAGLCLIENT}
   Maineagl,
{$ELSE}
   dbctrls,dbTables,Fe_Main,
{$ENDIF}
   M3FP,TarifCond, AglInitGC, UtilArticle, EntGC,LookUp, HDB;

Function EntreeTarifTiersMode (Action : TActionFiche; TarifTTC : Boolean=False) : boolean ;
Function SaisieTarifTiersMode (NatTiers, CodeT : string; Action : TActionFiche; TarifTTC : Boolean=False) : boolean ;
//Type T_TableTarif = (ttdCatArt, ttdCatQArt, ttdCatFam) ;

type
  TFTarifTiersMode = class(TForm)
    Panel1: TPanel;
    PPIED: TPanel;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    MsgBox: THMsgBox;
    BChercher: TToolbarButton97;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    FindLigne: TFindDialog;
    BInfos: TToolbarButton97;
    POPZ: TPopupMenu;
    InfTiers: TMenuItem;
    HMTrad: THSystemMenu;
    HMess: THMsgBox;
    GF_CASCADEREMISE: THValComboBox;
    TGF_CASCADEREMISE: THLabel;
    PFAMARTICLE: THPanel;
    PTITRE: THPanel;
    HTitre: THMsgBox;
    G_Fam: THGrid;
    CONDAPPLIC: THRichEditOLE ;
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
    GF_REMISE: THCritMaskEdit ;
    PMode: TPanel;
    TGF_TYPTARIF_MODE: THLabel;
    _TypTarifMode: THValComboBox;
    TGF_DEVISE: THLabel;
    TGF_PerTarif_Mode: THLabel;
    _PerTarifMode: THValComboBox;
    TGF_DATEDEBUT: THLabel;
    TGF_DATEFIN: THLabel;
    TGF_DEPOT: THLabel;
    GF_DEPOT: THValComboBox;
    GF_DATEDEBUT: THCritMaskEdit;
    GF_DATEFIN: THCritMaskEdit;
    HTarif: THMsgBox;
    GF_DEVISE: TEdit;
    BnvType: TToolbarButton97;
    BnvPeriode: TToolbarButton97;

    procedure FormCreate(Sender:TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormDestroy(Sender: TObject) ;

//Evènement de l'entête
    procedure _TYPTARIFMODEChange(Sender:TObject) ;
    procedure _TYPTARIFMODEExit(Sender:TObject) ;
    procedure _PERTARIFMODEChange(Sender:TObject) ;
    procedure _PERTARIFMODEExit(Sender:TObject) ;
    procedure GF_DEPOTExit(Sender: TObject);
    procedure GF_DEPOTChange(Sender: TObject) ;

//Evènements de la Grid
    procedure G_FamEnter(Sender: TObject);
    procedure G_FamCellEnter(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure G_FamElipsisClick(Sender: TObject);
    procedure G_FamCellExit(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure G_FamRowEnter(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean) ;
    procedure G_FamRowExit(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);

// Boutons
    procedure BValiderClick(Sender: TObject) ;
    procedure BChercherClick(Sender: TObject) ;
    procedure FindLigneFind(Sender: TObject) ;
    procedure InfTiersClick(Sender: TObject) ;
    procedure BCondAplliClick(Sender: TObject);
    procedure BCopierCondClick(Sender: TObject);
    procedure BCollerCondClick(Sender: TObject);
    procedure BVoirCondClick(Sender: TObject);
    procedure BNvPeriodeClick(Sender: TObject);
    procedure BNvTypeClick(Sender: TObject);
    procedure TCONDTARFClose(Sender: TObject);
    procedure GF_CASCADEREMISEChange(Sender: TObject);
    procedure VoirFicheTiers ;
    procedure BAbandonClick(Sender: TObject) ;

  private
    iTableLigne : integer ;
    FindDebut,FClosing : Boolean;
    StCellCur, LesColFAM : String ;
    ColsInter : Array of boolean ;
    DEV       : RDEVISE ;
    NatureTiers : string ;
    TarifTTC,SavEntete,InitEntete,TarifExistant,SaisieAutomatique,ModifMul: Boolean ;
// Objets mémoire
    TOBTarif, TOBTarfFam, TOBTiers, TOBTarifDel, TOBCatTiers,TobTarifMode: TOB;
// Menu
    procedure AffectMenuCondApplic (G_GRID : THGrid) ;
// Actions liées au Grid
    procedure EtudieColsListe ; 
    procedure PostDrawCell(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    procedure FormateZoneSaisie (ACol,ARow : Longint) ;
    procedure InsertLigne (ARow : Longint) ;
    procedure SupprimeLigne (ARow : Longint) ;
    procedure SupprimeTOBTarif (ARow : Longint) ;
    Function  GrilleModifie : Boolean ;
    Function  SortDeLaLigne : Boolean ;

// ENTETE
    Procedure InitialiseEntete ;
    procedure AffecteEntete ;

// Initialisation
    Procedure LoadLesTOB ;
    Procedure ChargeTarif ;

    procedure LibereTOBMode ;

// Ligne
    procedure InitialiseGrille ;
    procedure DepileTOBLigne ;
    procedure AfficheLaLigne(Arow: integer) ;
    procedure InitLaLigne(ARow: Integer) ;
    function  GetTOBLigne(ARow: integer) : TOB;
    procedure InitialiseLigne (ARow: integer) ;
    procedure CreerTOBLigne (ARow: integer) ;
    function  LigneVide (ARow: integer) : Boolean ;
    procedure PreaffecteLigne (ARow: Integer) ;

// CHAMPS LIGNES
    procedure TraiterTiers (ACol, ARow:Integer ) ;
    procedure TraiterCatTiers (ACol, ARow:Integer ) ;
    procedure TraiterCatArticle (ACol, ARow:Integer ) ;
    procedure TraiterRemise (ACol, ARow:Integer ) ;
    Function  QuestionTarifEnCours : Integer;
    procedure TraiterEtablissement (ARow : integer);

// MAJ Bouleen dernier utilise pour entete
    procedure MAJChampEntete ;

// Pied
   Procedure InitialisePied ;
   //procedure AffichePrixCon ;
   procedure AffichePied (ARow : Longint) ;


// Validation
    procedure ValideTarif ;
    procedure VerifLesTOB ;

// Conditions tarifaires
   procedure InitComboChamps ;
   procedure RemplitComboChamps(NomTable : string ; FCombo : THValComboBox) ;
   procedure EffaceGrid ;
   procedure AfficheCondTarf (ARow : Longint) ;
   function  ValueToItem(CC: THValcomboBox ; St : String) : String ;
   procedure GetConditions (TOBL : TOB) ;
// Visualiser un tarif existant à partir de la consultation
   procedure ChargeTarifDepuisMul ;


  public
    { Déclarations publiques }
    CodeDevise  : string ;
    CodeTiers   : string ;
    CodeCatTiers : string ;
    CodeDepot : string ;
    CodeTarifArticle : string ;
    CodetarifMode, CodePeriode, CodeType,IdTarif,CodeArrondi : string ;
    Action : TActionFiche ;
    Existant: Boolean;
  end;

var
  FTarifTiersMode: TFTarifTiersMode;
  lModif : boolean;
  SFA_Depot    : integer;
  SFA_Tiers    : integer ;
  SFA_CatTiers : integer ;
  SFA_CatArt   : integer;
  SFA_Lib      : integer;
  //SFA_Px      : integer;
  SFA_Rem      : integer;
  //SFA_Datedeb  : integer;
  //SFA_Datefin  : integer;


Function TrouverTiers(CodeTiers:String ; TOBTiers: TOB): T_RechArt ;

implementation

{$R *.DFM}

Function EntreeTarifTiersMode (Action : TActionFiche; TarifTTC : Boolean=False) : boolean ;
BEGIN
Result := SaisieTarifTiersMode ('CLI', '', Action, TarifTTC) ;
END;

Function SaisieTarifTiersMode (NatTiers, CodeT : string; Action : TActionFiche; TarifTTC : Boolean=False) : boolean ;
var FF : TFTarifTiersMode ;
    PPANEL  : THPanel ;
begin
Result:=True ;
SourisSablier;
FF := TFTarifTiersMode.Create(Application) ;
FF.Action:=Action ;
FF.TarifTTC:=TarifTTC ;
FF.NatureTiers:=NatTiers;
FF.CodeTiers:=CodeT ;
FF.PMode.Visible:=True ;
PPANEL := FindInsidePanel ;
if PPANEL = Nil then
   BEGIN
    try
      FF.ShowModal ;
    finally
      FF.Free ;
      Result:=False ;
    end ;
   SourisNormale ;
   END else
   BEGIN
   InitInside (FF, PPANEL) ;
   FF.Show ;
   END ;
END ;

{Procedure AppelTarifTiersMode ( Parms : array of variant ; nb : integer ) ;
begin
end ; }


{==============================================================================================}
{=============================== Action liées à la forme ========================================}
{==============================================================================================}
procedure TFTarifTiersMode.FormCreate(Sender: TObject);
begin
G_FAM.RowCount := NbRowsInit ;
StCellCur := '' ;
iTableLigne := PrefixeToNum('GF') ;
TOBTarif := TOB.Create ('', Nil, -1) ;
TOBTarfFam := TOB.Create ('', TOBTarif, 0) ;
TOBTarifDel := TOB.Create ('', Nil, -1) ;
TOBCatTiers := TOB.Create ('CHOIXCOD', Nil, -1) ;
TOBTiers := TOB.Create ('TIERS', Nil, -1) ;
//TobTarifMode:=TOB.Create('TARIFMODE',NIL,-1) ;
FClosing:=False ;
end;

procedure TFTarifTiersMode.FormDestroy(Sender: TObject) ;
begin
TOBTarif.Free ;
TOBCatTiers.free ;
TOBTarifDel.free ;
TOBTiers.free ;
TobTarifMode.Free ;
AglDepileFiche ;
end ;

procedure TFTarifTiersMode.FormShow(Sender: TObject);
begin
G_FAM.ListeParam:='GCTARIFCLIENTMODE';
G_FAM.PostDrawCell:=PostDrawCell ;
EtudieColsListe ;
HMTrad.ResizeGridColumns (G_FAM) ;
AffecteGrid (G_FAM,Action) ;
PFAMARTICLE.Visible := True;
if TarifTTC then _TYPTARIFMODE.DataType := 'GCTARIFTYPE1VTE' ;
_TYPTARIFMODE.Reload ;
if TheTob<>nil then
   begin
   ChargeTarifDepuisMul ;
   end else
   begin
   InitialiseEntete ;
  end ;
PTITRE.Caption := HTitre.Mess[0] ;
TTYPETARIF.Caption := HTitre.Mess[1]+RechDom('TTDEVISE',CodeDevise,False) ;
DEV.Code := CodeDevise ; GetInfosDevise (DEV) ;
InitComboChamps ;
AglEmpileFiche(Self) ;
end;

procedure TFTarifTiersMode.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
if Action=taConsult then Exit ;
if GrilleModifie then
   BEGIN
   if MsgBox.Execute(6,Caption,'')<>mrYes then CanClose:=False ;
   END ;
end ;

procedure TFTarifTiersMode.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
G_FAM.VidePile(True) ;
TOBTarif.Free ; TOBTarif:=Nil ;
TOBCatTiers.Free ; TOBCatTiers:=Nil ;
TOBTiers.Free ; TOBTiers:=Nil ;
TobTarifMode.Free; TobTarifMode:=Nil ;
if IsInside(Self) then Action:=caFree ;
FClosing:=True ;
end;

procedure TFTarifTiersMode.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
Case Key of
    VK_RETURN : Key:=VK_TAB ;
    VK_INSERT : BEGIN
                if (Screen.ActiveControl = G_FAM) then
                    BEGIN
                    Key := 0;
                    InsertLigne (G_FAM.Row);
                    END;
                END;
    VK_DELETE : BEGIN
                if ((Screen.ActiveControl = G_FAM) and (Shift=[ssCtrl])) then
                    BEGIN
                    Key := 0 ;
                    SupprimeLigne (G_FAM.Row) ;
                    END ;
                END;
    END;
end ;

//==============================================================================
//  Evenements sur G_FAM
//==============================================================================
procedure TFTarifTiersMode.G_FamEnter(Sender: TObject);
var Cancel, Chg : Boolean;
    ACol, ARow : integer;
begin

{if (_TypTarifMode.Value<>'...')  and (GF_DEPOT.Value='') then
   begin
   PGIBox('L''établissement est obligatoire pour un type de tarif','Tarifs par article') ;
   IF CodeDepot<>'' then GF_DEPOT.Value:=CodeDepot else GF_DEPOT.SetFocus ;
   end ;   }
if _PERTARIFMODE.Value='' then
   BEGIN
   HTarif.Execute(1,Caption,'') ;
   _PERTARIFMODE.SetFocus  ;
   Exit ;
   END ;
if (_PERTARIFMODE.Value<>'') then
  begin
  TobTarifMode:=TraiterTableTarifMode(_TYPTARIFMODE.Value,_PERTARIFMODE.Value,'VTE',TobTarifMode) ;
  CodeTarifMode:=TobTarifMode.Getvalue('GFM_TARFMODE') ;
  end ;
Cancel := False; Chg := False;
G_FamRowEnter(Sender, G_Fam.Row, Cancel, Chg);
ACol := G_Fam.Col ; ARow := G_Fam.Row ;
G_FamCellEnter (Sender, ACol, ARow, Cancel);
end;


procedure TFTarifTiersMode.G_FamRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean) ;
var ARow : Integer;
begin
//if (ModifMul) and (LigneVide(Ou)) then BAbandonClick(Nil) ;
if Ou >= G_Fam.RowCount - 1 then G_Fam.RowCount := G_Fam.RowCount + NbRowsPlus ;;
ARow := Min (Ou, TOBTarfFam.detail.count + 1);
if (ARow = TOBTarfFam.detail.count + 1) AND (not LigneVide (ARow - 1)) then
    BEGIN
    CreerTOBligne (ARow);
    END;
if (LigneVide (ARow)) AND (not LigneVide (ARow - 1))then
    PreAffecteLigne (ARow);
if Ou > TOBTarfFam.detail.count then
    BEGIN
    G_Fam.Row := TOBTarfFam.detail.count;
    END;
AffichePied (ARow);
AfficheCondTarf (ARow);
end;

procedure TFTarifTiersMode.G_FamRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if LigneVide (Ou) Then G_Fam.Row := Min (G_Fam.Row,Ou);
end ;

//==============================================================================
//  Gestion des clicks boutons
//==============================================================================
procedure TFTarifTiersMode.BValiderClick(Sender: TObject);
Var ioerr : TIOErr ;
begin
ioerr := Transactions (ValideTarif, 2);
Case ioerr of
        oeOk : ;
   oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
    oeSaisie : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
   END ;
MAJChampEntete ;
InitialiseEntete;
end;

procedure TFTarifTiersMode.BChercherClick(Sender: TObject) ;
Begin
if G_FAM.RowCount < 3 then Exit;
FindDebut:=True ; FindLigne.Execute ;
end ;

procedure TFTarifTiersMode.BNvPeriodeClick(Sender: TObject);
begin
AglLanceFiche ('MBO','TARIFPER','','', 'ACTION=CREATION') ;
end ;

procedure TFTarifTiersMode.BNvTypeClick(Sender: TObject);
begin
AglLanceFiche('MBO', 'TARIFTYPE', '', '', 'TYPE=VTE;ACTION=CREATION');
end ;

procedure TFTarifTiersMode.FindLigneFind(Sender: TObject);
Begin
Rechercher (G_FAM, FindLigne, FindDebut) ;
end ;

procedure TFTarifTiersMode.InfTiersClick(Sender: TObject);
begin
if G_FAM.Cells[SFA_Tiers,G_FAM.Row]='' then exit ;
VoirFicheTiers;
end;

procedure TFTarifTiersMode.VoirFicheTiers;
BEGIN
AglLanceFiche ('GC', 'GCTIERS', '' , TOBTiers.GetValue('T_AUXILIAIRE'), 'ACTION=CONSULTATION;TARIF=N;MONOFICHE');
END;

procedure TFTarifTiersMode.BVoirCondClick(Sender: TObject);
begin
TCONDTARF.Visible := BVoirCond.Down ;
end;

procedure TFTarifTiersMode.BCondAplliClick(Sender: TObject);
Var TOBL : TOB;
begin
if GF_DEPOT.Text = '' then exit;
TOBL:=GetTOBLigne(G_FAM.Row) ; if TOBL=Nil then Exit ;
EntreeTarifCond (Action, TOBL);
AfficheCondTarf (G_FAM.Row);
end;

procedure TFTarifTiersMode.BCopierCondClick(Sender: TObject);
Var TOBL : TOB;
begin
if GF_Depot.Text = '' then exit;
TOBL:=GetTOBLigne(G_FAM.Row) ; if TOBL=Nil then Exit ;
CONDAPPLIC.Text := TOBL.GetValue ('GF_CONDAPPLIC') ;
end;

procedure TFTarifTiersMode.BCollerCondClick(Sender: TObject);
Var TOBL : TOB;
begin
if GF_DEPOT.Text = '' then exit;
TOBL:=GetTOBLigne(G_FAM.Row) ; if TOBL=Nil then Exit ;
TOBL.PutValue ('GF_CONDAPPLIC', CONDAPPLIC.Text) ;
AfficheCondTarf (G_FAM.Row);
end;

{==============================================================================================}
{============================= Manipulation liées au Menu =====================================}
{==============================================================================================}
procedure TFTarifTiersMode.AffectMenuCondApplic (G_GRID : THGrid) ;
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
procedure TFTarifTiersMode.PostDrawCell(ACol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
begin
//
end ;

procedure TFTarifTiersMode.EtudieColsListe ;
Var NomCol,LesCols : String ;
    icol,ichamp, i_ind : integer ;
BEGIN
G_FAM.ColWidths[0]:=0 ;
for i_ind:=Low(ColsInter)to High(ColsInter) do ColsInter[i_ind]:=False ;
LesCols:=G_FAM.Titres[0] ; LesColFAM := LesCols ; icol:=0 ;
Repeat
 NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
 if NomCol<>'' then
    BEGIN
    ichamp:=ChampToNum(NomCol) ;
    if ichamp>=0 then
       BEGIN
       if NomCol='GF_DEPOT'        then SFA_Depot:=icol else
       if NomCol='GF_TIERS'        then SFA_Tiers:=icol else
       if NomCol='GF_TARIFTIERS'   then SFA_CatTiers:=icol else
       if NomCol='GF_TARIFARTICLE' then SFA_CatArt:=icol else
       //if NomCol='GF_LIBELLE'      then SFA_Lib:=icol else
       if NomCol='GF_CALCULREMISE' then SFA_Rem:=icol else
       //if NomCol='GF_DATEDEBUT'    then SFA_Datedeb:=icol else
       //if NomCol='GF_DATEFIN'      then SFA_Datefin:=icol ;
       END ;
    END ;
 Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;
END ;

procedure TFTarifTiersMode.FormateZoneSaisie (ACol,ARow : Longint) ;
Var St,StC : String ;
BEGIN
St:=G_Fam.Cells[ACol,ARow] ; StC:=St ;
if ((ACol=SFA_CatArt) or (ACol=SFA_CatTiers) or (ACol=SFA_Tiers)) then StC:=uppercase(Trim(St));
   //if ACol=SFA_Px then StC:=StrF00(Valeur(St),DEV.Decimale) //else
//            if ((ACol=SFA_QInf) or (ACol=SFA_QSup)) then StC:=StrF00(Valeur(St),0);

G_Fam.Cells[ACol,ARow]:=StC ;
END ;

procedure TFTarifTiersMode.InsertLigne (ARow : Longint) ;
Begin
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if LigneVide (ARow) then exit;
if (ARow > TOBTarfFam.Detail.Count) then Exit;
G_FAM.CacheEdit; G_FAM.SynEnabled := False;
TOB.Create ('TARIF', TOBTarfFam, ARow-1) ;
G_FAM.InsertRow (ARow); G_FAM.Row := ARow;
InitialiseLigne (ARow) ;
PreAffecteLigne (ARow);
G_FAM.MontreEdit; G_FAM.SynEnabled := True;
AffectMenuCondApplic (G_FAM);
AfficheCondTarf (G_FAM.Row);
END;

procedure TFTarifTiersMode.SupprimeLigne (ARow : Longint) ;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if (ARow > TOBTarfFam.Detail.Count) then Exit;
G_FAM.CacheEdit; G_FAM.SynEnabled := False;
G_FAM.DeleteRow (ARow);
if (ARow = TOBTarfFam.Detail.Count) then
    CreerTOBLigne (ARow + 1);
SupprimeTOBTarif (ARow);
if G_FAM.RowCount < NbRowsInit then G_FAM.RowCount := NbRowsInit;
G_FAM.MontreEdit; G_FAM.SynEnabled := True;
AffectMenuCondApplic (G_FAM);
AfficheCondTarf (G_FAM.Row);
END;

procedure TFTarifTiersMode.SupprimeTOBTarif ( ARow : Longint) ;
Var i_ind: integer;
BEGIN
if TOBTarfFam.Detail[ARow-1].GetValue ('GF_TARIF') <> 0 then
    BEGIN
    i_ind := TOBTarifDel.Detail.Count;
    TOB.Create ('TARIF', TOBTarifDel, i_ind) ;
    TOBTarifDel.Detail[i_ind].Dupliquer (TOBTarfFam.Detail[ARow-1], False, True);
    END;
TOBTarfFam.Detail[ARow-1].Free;
TOBTarfFam.SetAllModifie(True);
END;

Function TFTarifTiersMode.GrilleModifie : Boolean ;
BEGIN
Result:=False ;
if Action=taConsult then Exit ;
Result:=(TOBTarfFam.IsOneModifie) or (TOBTarifDel.IsOneModifie);
END;

Function  TFTarifTiersMode.SortDeLaLigne : Boolean ;
Var ACol,ARow : integer ;
    Cancel : boolean ;
BEGIN
Result:=False ;
ACol:=G_FAM.Col ; ARow:=G_FAM.Row ; Cancel:=False ;
G_FAMCellExit(Nil,ACol,ARow,Cancel) ; if Cancel then Exit ;
G_FAMRowExit(Nil,ACol,Cancel,False) ; if Cancel then Exit ;
Result:=True ;
END ;

{==============================================================================================}
{========================= Evenement de l'Entete  =============================================}
{==============================================================================================}
procedure TFTarifTiersMode.GF_DEPOTExit(Sender: TObject);
var Info: String;
begin
CodeDepot:=GF_DEPOT.Value ;
Info:=RecupInfoPeriode(_PERTARIFMODE.Value,CodeDepot) ;
GF_DATEDEBUT.Text:=ReadTokenSt(Info) ;
GF_DATEFIN.Text:=ReadTokenSt(Info) ;
CodeArrondi:=ReadTokenSt(Info) ;
end;

procedure TFTarifTiersMode._TYPTARIFMODEExit(Sender: TObject) ;
Var QQ: TQuery ;
Etablissement: String ;
begin
CodeType:=_TYPTARIFMODE.Value ;
if CodeType='...' then
  begin
  GF_DEPOT.Visible:=False ;
  TGF_DEPOT.Visible:=False ;
  GF_DEPOT.Value:='' ;
  CodeDepot:='' ;
  end else
  begin
  GF_DEPOT.Visible:=True ;
  TGF_DEPOT.Visible:=True ;
  end ;
QQ:=OpenSql('Select GFT_DEVISE,GFT_ETABLISREF from TARIFTYPMODE where GFT_CODETYPE="'+_TYPTARIFMODE.Value+'"',True) ;
if not QQ.EOF then
  begin
  CodeDevise:=QQ.FindField('GFT_DEVISE').AsString ;
  GF_DEVISE.Text:=RechDom('TTDEVISE',CodeDevise,False) ;
  DEV.Code := CodeDevise ; GetInfosDevise (DEV) ;
  //GF_DEPOT.Value:=QQ.FindField('GFT_ETABLISREF').AsString ;
  Etablissement:=QQ.FindField('GFT_ETABLISREF').AsString ;
  end else
  GF_DEVISE.Text:=RechDom('TTDEVISE',GetParamSoc('SO_DEVISEPRINC'),False) ;
  CodeDevise:=GetParamSoc('SO_DEVISEPRINC') ;
ferme(QQ) ;
GF_DEPOT.Plus:='(ET_TYPETARIF = "'+CodeType+'")';
GF_DEPOT.Value:=ReadTokenSt(Etablissement) ;
TTYPETARIF.Caption := HTitre.Mess[1]+RechDom('TTDEVISE',CodeDevise,False) ;
G_Fam.Enabled:=True ;
end ;

procedure TFTarifTiersMode._PERTARIFMODEExit(Sender: TOBject) ;
Var Info: String;
begin
CodePeriode:=_PERTARIFMODE.Value ;
Info:=RecupInfoPeriode(_PERTARIFMODE.Value,GF_DEPOT.Value) ;
GF_DATEDEBUT.Text:=ReadTokenSt(Info) ;
GF_DATEFIN.Text:=ReadTokenSt(Info) ;
CodeArrondi:=ReadTokenSt(Info) ;
end ;

procedure TFTarifTiersMode._TYPTARIFMODEChange(Sender: TObject);
var i_Rep : integer;
    ioerr : TIOErr ;
begin
if InitEntete then exit ;
if (CodeType = _TYPTARIFMODE.Value) or (CodeType='') or (CodeType='...') then Exit;
i_Rep := QuestionTarifEnCours;
Case i_Rep of
    mrYes    : BEGIN
               ioerr := Transactions (ValideTarif, 2);;
               Case ioerr of
                    oeOk      : ;
                    oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
                    oeSaisie  : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
                    END ;
               //InitTarif ;
               CodeType := _TYPTARIFMODE.Value;
               if CodeType='' then CodeType:='...' ;
               Transactions (ChargeTarif, 1) ;
               END;
    mrNo     : BEGIN
               InitialiseGrille;
               DepileTOBLigne;
               LibereTOBMode ;
               CodetarifMode:='' ;  CodeTarifArticle:='' ;
               CodeType := _TYPTARIFMODE.Value ;
               if CodeType='' then CodeType:='...' ;
               Transactions (ChargeTarif, 1) ;
               END;
    mrCancel : BEGIN
               _TYPTARIFMODE.Value := CodeType;
               END ;
    END ;
//DEV.Code := GF_DEVISE.Value ; GetInfosDevise (DEV) ;
InitialisePied ;
end;

procedure TFTarifTiersMode._PERTARIFMODEChange(Sender: TObject);
var i_Rep : integer;
    ioerr : TIOErr ;
begin
if InitEntete then exit ;
if (CodePeriode = _PERTARIFMODE.Value) or (CodePeriode='') then Exit;
i_Rep := QuestionTarifEnCours;
Case i_Rep of
    mrYes    : BEGIN
               ioerr := Transactions (ValideTarif, 2);;
               Case ioerr of
                    oeOk      : ;
                    oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
                    oeSaisie  : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
                    END ;
               //InitTarif ;
               CodePeriode := _PERTARIFMODE.Value;
               Transactions (ChargeTarif, 1) ;
               END;
    mrNo     : BEGIN
               InitialiseGrille;
               DepileTOBLigne;
               LibereTOBMode ;
               CodetarifMode:='' ; CodeTarifArticle:='' ;
               CodePeriode := _PERTARIFMODE.Value ;
               Transactions (ChargeTarif, 1) ;
               END;
    mrCancel : _PERTARIFMODE.Value := CodePeriode;
    END ;
InitialisePied ;
end;

procedure TFTarifTiersMode.GF_DEPOTChange(Sender: TObject) ;
var i_Rep : integer;
    ioerr : TIOErr ;
begin
if InitEntete then exit ;
if (CodeDepot = GF_DEPOT.Value) then Exit;
i_Rep := QuestionTarifEnCours;
Case i_Rep of
    mrYes    : BEGIN
               ioerr := Transactions (ValideTarif, 2);;
               Case ioerr of
                    oeOk      : ;
                    oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
                    oeSaisie  : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
                    END ;
               //InitTarif ;
               CodeDepot := GF_DEPOT.Value;
               Transactions (ChargeTarif, 1) ;
               END;
    mrNo     : BEGIN
               InitialiseGrille;
               DepileTOBLigne;
               LibereTOBMode ;
               CodeDepot := GF_DEPOT.Value ;
               Transactions (ChargeTarif, 1) ;
               END;
    mrCancel : GF_DEPOT.Value := CodeDepot;
    END ;
InitialisePied ;
End ;

procedure TFTarifTiersMode.AffecteEntete ;
Begin
//GF_DEPOT.Text:= CodeDepot ;
//TGF_LIBDEPOT.Caption:=RechDom(GF_DEPOT.DataType,CodeDepot,FALSE) ;
end ;

procedure TFTarifTiersMode.initialiseEntete ;
Var QPer,QTyp: TQuery ;
Info,Etablissement :String ;
BEGIN
inherited ;
InitEntete:=True ;
TOBCatTiers.InitValeurs ;
TOBTiers.InitValeurs ;
CodeCatTiers := '';
CodeTiers := '';
SavEntete:=False ;
// modif le 13/08/02 on parle d'établissement et non de dépôt
GF_DEPOT.Value:=VH^.EtablisDefaut; //VH_GC.GCDepotDefaut ;
CodeDepot:=GF_DEPOT.Value ;
QTyp:=OpenSQL('Select GFT_DEVISE,GFT_CODETYPE,GFT_ETABLISREF from TarifTypMode where GFT_DERUTILISE="X"',True) ;
if Not QTyp.EOF then
   begin
   CodeType:=QTyp.FindField('GFT_CODETYPE').AsString ;
   if CodeType='...' then
    begin
    GF_DEPOT.Visible:=False ;
    TGF_DEPOT.Visible:=False ;
    GF_DEPOT.Value:='' ;
     _TYPTARIFMODE.Value:='...' ;
    end ;
   _TYPTARIFMODE.Value:=CodeType ;
   CodeDevise:=QTyp.FindField('GFT_DEVISE').AsString ;
   GF_DEVISE.Text:=RechDom('TTDEVISE',CodeDevise,False) ;
   TTYPETARIF.Caption := HTitre.Mess[1]+RechDom('TTDEVISE',CodeDevise,False) ;
   //GF_DEPOT.Value:=QTyp.FindField('GFT_ETABLISREF').AsString ;
   Etablissement:=QTyp.FindField('GFT_ETABLISREF').AsString ;
   //CodeDepot:=GF_DEPOT.Value ;
   end else
   begin
   CodeType:='...';   
   GF_DEVISE.Text:=RechDom('TTDEVISE',GetParamSoc('SO_DEVISEPRINC'),False) ;
   CodeDevise:=GetParamSoc('SO_DEVISEPRINC') ;
   GF_DEPOT.Visible:=False ;
   TGF_DEPOT.Visible:=False ;
   GF_DEPOT.Value:='' ;
   CodeDepot:=GF_DEPOT.Value ;
   end ;
QPer:=OpenSQL('Select * from TarifPer where GFP_DERUTILISE="X"',True) ;
if Not QPer.EOF then
   begin
   CodePeriode:=QPer.FindField('GFP_CODEPERIODE').AsString ;
   _PERTARIFMODE.Value:=CodePeriode ;
   GF_DATEDEBUT.Text:=QPer.FindField('GFP_DATEDEBUT').AsString ;
   GF_DATEFIN.Text:=QPer.FindField('GFP_DATEFIN').AsString ;
   CodeArrondi:=QPer.FindField('GFP_ARRONDI').AsString ;
   //if QPer.FindField('GFP_DEMARQUE').AsString<>'' then GF_DEMARQUE.Text:=RechDom('GCTYPEREMISE',QPer.FindField('GFP_DEMARQUE').AsString,False) ;
   end else
   begin
   CodePeriode:='' ;
   _PERTARIFMODE.Value:='' ;
   GF_DATEDEBUT.Text:='01/01/1900' ; GF_DATEFIN.Text:='31/12/2099' ;
   end ;
Ferme(QTyp) ;
Ferme(QPer) ;
GF_DEPOT.Plus:='(ET_TYPETARIF = "'+CodeType+'")';
GF_DEPOT.Value:=ReadTokenSt(Etablissement) ;
CodeDepot:=GF_DEPOT.Value ;
AffecteEntete;
BCondAplli.Enabled := False ; BCopierCond.Enabled := False; BCollerCond.Enabled := False ;
CONDAPPLIC.Text := '' ;
InitialisePied ;
InitEntete:=False ;
Info:=RecupInfoPeriode(_PERTARIFMODE.Value,CodeDepot) ;
GF_DATEDEBUT.Text:=ReadTokenSt(Info) ;
GF_DATEFIN.Text:=ReadTokenSt(Info) ;
CodeArrondi:=ReadTokenSt(Info) ;
END;

Procedure TFTarifTiersMode.LoadLesTOB ;
Var Q : TQuery ;
    i_ind : integer;
    Select : string;
Begin
if G_Fam.Cells[SFA_TIERS,G_Fam.Row]<>'' then
  begin
  for i_ind := TOBTarfFam.Detail.Count - 1 downto 0 do
      BEGIN
      TOBTarfFam.Detail[i_ind].Free;
      END;

  // Lecture Quantitatif
  Select := 'SELECT * FROM TARIF WHERE GF_TIERS="'+CodeTiers+'" AND GF_DEVISE="' + CodeDevise + '" '+
            'AND GF_DEPOT="'+CodeDepot+'" ORDER BY GF_BORNEINF';
            //'ORDER BY GF_DEPOT, GF_TARIFARTICLE, GF_BORNEINF';
  Q := OpenSQL(Select,True) ;
  TOBTarfFam.LoadDetailDB('TARIF','','',Q,False) ;
  TOBTarfFam.SetAllModifie(False);
  Ferme(Q) ;
  end
else if G_Fam.Cells[SFA_CATTIERS,G_Fam.Row]<>'' then
  begin
  for i_ind := TOBTarfFam.Detail.Count - 1 downto 0 do
      BEGIN
      TOBTarfFam.Detail[i_ind].Free;
      END;

  // Lecture Quantitatif
  Select := 'SELECT * FROM TARIF WHERE GF_TARIFTIERS="'+CodeCatTiers+'" AND GF_DEVISE="' + CodeDevise + '" ' +
            'AND GF_DEPOT="'+CodeDepot+'" ORDER BY GF_BORNEINF';
            //'ORDER BY GF_DEPOT, GF_TARIFARTICLE, GF_BORNEINF';
  Q := OpenSQL(Select,True) ;
  TOBTarfFam.LoadDetailDB('TARIF','','',Q,False) ;
  TOBTarfFam.SetAllModifie(False);
  Ferme(Q) ;
  end ;
end ;

Procedure TFTarifTiersMode.ChargeTarif ;
var i_ind : integer;
BEGIN
LoadLesTOB ;
// Affichage
for i_ind:=0 to TOBTarfFam.Detail.Count-1 do
    BEGIN
    AfficheLaLigne (i_ind + 1) ;
    END ;
END ;

//==============================================================================
//  Evenements sur GF_DEVISE
//==============================================================================
Function  TFTarifTiersMode.QuestionTarifEnCours : Integer;
BEGIN
Result := mrNo;
if (TOBTarfFam.IsOneModifie) or (TOBTarifDel.IsOneModifie) then
    BEGIN
    Result := MsgBox.Execute (0, Caption, '');
    END;
END;


{==============================================================================================}
{=============================== Ligne ========================================================}
{==============================================================================================}
procedure TFTarifTiersMode.InitialiseGrille ;
BEGIN
G_FAM.VidePile(True) ;
G_FAM.RowCount:= NbRowsInit ;
END;

Procedure TFTarifTiersMode.DepileTOBLigne ;
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
END;

procedure TFTarifTiersMode.LibereTOBMode ;
begin
TobTarifMode.Free ; TobTarifMode:=nil ;
end ;


procedure TFTarifTiersMode.AfficheLaLigne (ARow : integer) ;
Var TOBL : TOB ;
    i_ind : integer ;
BEGIN
TOBL:=GetTOBLigne(ARow) ; if TOBL = Nil then exit;
TOBL.PutLigneGrid(G_FAM,ARow,False,False,LesColFAM) ;
for i_ind := 1 to G_FAM.ColCount-1 do
    FormateZoneSaisie(i_ind, ARow) ;
END ;

procedure TFTarifTiersMode.InitLaLigne (ARow: Integer) ;
Var TOBL : TOB;
BEGIN
TOBL:=GetTOBLigne(ARow) ; if TOBL=Nil then Exit ;
TarifExistant:=False ;
TOBL.PutValue ('GF_NATUREAUXI', NatureTiers);
TOBL.PutValue ('GF_TIERS', G_Fam.Cells[SFA_Tiers,Arow]);
TOBL.PutValue ('GF_TARIFTIERS', G_Fam.Cells[SFA_CatTiers,Arow]);
if G_Fam.Cells[SFA_Tiers,Arow]<>'' then TOBL.PutValue ('GF_LIBELLE',copy(RechDom('GCTIERSCLI',G_Fam.Cells[SFA_Tiers,Arow],FALSE),1,35))
else TOBL.PutValue ('GF_LIBELLE',RechDom('TTTARIFCLIENT',G_Fam.Cells[SFA_CatTiers,Arow],FALSE)) ;
TOBL.PutValue ('GF_BORNESUP', 999999);
TOBL.PutValue ('GF_PRIXUNITAIRE', 0);
TOBL.PutValue ('GF_REMISE', 0);
TOBL.PutValue ('GF_CALCULREMISE', '');
TOBL.PutValue ('GF_DATEDEBUT', StrToDate(GF_DATEDEBUT.Text));
TOBL.PutValue ('GF_DATEFIN', StrToDate(GF_DATEFIN.Text));
TOBL.PutValue ('GF_TARFMODE',CodeTarifMode) ;
TOBL.PutValue ('GF_MODECREATION', 'MAN');
TOBL.PutValue('GF_ARRONDI',CodeArrondi) ;
TOBL.PutValue ('GF_CASCADEREMISE','MIE');
TOBL.PutValue ('GF_DEVISE', CodeDevise);
TOBL.PutValue ('GF_QUALIFPRIX', 'GRP');
TOBL.PutValue ('GF_FERME', '-');
//TOBL.PutValue ('GF_DEPOT', VH_GC.GCDepotDefaut) ;
TraiterEtablissement(ARow) ;
//TOBL.PutValue ('GF_DEPOT', GF_DEPOT.Value) ;
//TOBL.PutValue ('GF_SOCIETE', TOBL.GetValue('GF_DEPOT')) ;
TOBL.PutValue ('GF_QUANTITATIF', '-');
TOBL.PutValue ('GF_BORNEINF', -999999);
if TarifTTC then begin TOBL.PutValue ('GF_REGIMEPRIX', 'TTC'); TOBL.PutValue ('GF_NATUREAUXI', 'CLI'); end
            else begin TOBL.PutValue ('GF_REGIMEPRIX', 'HT') ;TOBL.PutValue ('GF_NATUREAUXI', 'FOU'); end ;
AfficheLAligne (ARow);
END;

function TFTarifTiersMode.GetTOBLigne(ARow: integer) : TOB ;
BEGIN
Result:=Nil ;
if ((TOBTarfFam = nil) or (ARow<=0) or (ARow>TOBTarfFam.Detail.Count)) then Exit ;
Result:=TOBTarfFam.Detail[ARow-1] ;
END ;

procedure TFTarifTiersMode.InitialiseLigne(ARow : integer) ;
Var TOBL : TOB ;
    i_ind : integer ;
BEGIN
TOBL:=GetTOBLigne(ARow) ; if TOBL<>Nil then TOBL.InitValeurs ;
for i_ind := 1 to G_FAM.ColCount-1 do
    BEGIN
    G_FAM.Cells [i_ind, ARow]:='' ;
    END;
END;

procedure TFTarifTiersMode.CreerTOBLigne(ARow : integer) ;
BEGIN
if ARow <> TOBTarfFam.Detail.Count + 1 then exit;
TOB.Create ('TARIF', TOBTarfFam, ARow-1) ;
InitialiseLigne (ARow) ;
END;

function TFTarifTiersMode.LigneVide (ARow: integer) : Boolean ;
BEGIN
Result := True;
if (G_Fam.Cells [SFA_Tiers, ARow] <> '') Or (G_Fam.Cells [SFA_CatTiers, ARow] <> '') then
    BEGIN
    Result := False;
    Exit;
    END;
END;

procedure TFTarifTiersMode.PreaffecteLigne (ARow: Integer) ;
var TOBL, TOBLPrec : TOB;
BEGIN
TOBLPrec := GetTOBLigne (ARow - 1); if TOBLPrec = nil then exit;
if TOBLPrec.GetValue ('GF_BORNESUP') < 999999 then
    BEGIN
    TOBTarfFam.Detail[ARow - 1].Dupliquer (TOBLPrec, False, True);
    TOBL := GetTOBLigne (ARow);
    TOBL.PutValue ('GF_BORNEINF', TOBLPrec.GetValue ('GF_BORNESUP') + 1);
    TOBL.PutValue ('GF_BORNESUP', 999999);
    TOBL.PutValue ('GF_TARIF', 0);
    AfficheLAligne (ARow);
    END;
END;

{==============================================================================================}
{=============================== Evènements de la Grid ========================================}
{==============================================================================================}
procedure TFTarifTiersMode.G_FamCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if Not Cancel then
    BEGIN
    if (G_FAM.Col = SFA_CatArt) then
            if (G_FAM.Cells [SFA_Tiers,G_FAM.Row] = '') And (G_FAM.Cells [SFA_CatTiers,G_FAM.Row] = '')
               then G_FAM.Col := SFA_Tiers;
    if (G_FAM.Col = SFA_CatTiers) then
       if  (G_FAM.Cells [SFA_Tiers,G_FAM.Row] <> '') then G_FAM.Col := SFA_CatArt;
    if (G_FAM.Col = SFA_Tiers) then
       if  (G_FAM.Cells [SFA_CatTiers,G_FAM.Row] <> '') then G_FAM.Col := SFA_CatArt;

    G_FAM.ElipsisButton := (( G_FAM.col = SFA_Tiers) or (G_FAM.col = SFA_CatTiers) or
                               (G_FAM.col = SFA_CatArt)) ;
    StCellCur := G_FAM.Cells [G_FAM.Col,G_FAM.Row] ;
    AffectMenuCondApplic (G_FAM);
    END ;
end;

procedure TFTarifTiersMode.G_FamElipsisClick(Sender: TObject);
Var TIERS,CATTIERS,FAM : THCritMaskEdit;
    Coord : TRect;
    Cancel: Boolean ;
    ACol,ARow: Integer ;
begin
SaisieAutomatique:=False ;
Cancel:= False;
ARow:=G_FAM.Row ;
if (G_FAM.Col = SFA_Tiers) then
    BEGIN
    Coord := G_FAM.CellRect (G_FAM.Col, G_FAM.Row) ;
    TIERS := THCritMaskEdit.Create (Self) ;
    TIERS.Parent := G_FAM ;
    TIERS.Top := Coord.Top ;
    TIERS.Left := Coord.Left ;
    TIERS.Width := 3; TIERS.Visible := False ;
    TIERS.DataType := 'GCTIERSCLI';
    LookUpCombo (TIERS) ;
    if TIERS.Text <> '' then
    begin
    G_FAM.Cells[G_FAM.Col,G_FAM.Row]:= TIERS.Text ;
    ACol:=SFA_Tiers ;
    G_FAMCellExit(Sender,ACol,ARow,Cancel);
    G_FAM.Col:=SFA_CatArt ;
    end ;
    TIERS.Destroy;
    END ;
if (G_FAM.Col = SFA_CatTiers) then
    BEGIN
    Coord := G_FAM.CellRect (G_FAM.Col, G_FAM.Row);
    CATTIERS := THCritMaskEdit.Create (Self);
    CATTIERS.Parent := G_FAM;
    CATTIERS.Top := Coord.Top;
    CATTIERS.Left := Coord.Left;
    CATTIERS.Width := 3; CATTIERS.Visible := False;
    CATTIERS.DataType := 'TTTARIFCLIENT';
    GetCategorieRecherche (CATTIERS) ;
    if CATTIERS.Text <> '' then
    begin
    G_FAM.Cells[G_FAM.Col,G_FAM.Row]:= CATTIERS.Text;
    ACol:=SFA_CatTiers ;
    G_FAMCellExit(Sender,ACol,ARow,Cancel);
    G_FAM.Col:=SFA_CatArt ;
    end ;
    CATTIERS.Destroy;
    END ;
if G_FAM.Col = SFA_CatArt then
    BEGIN
    Coord := G_FAM.CellRect (G_FAM.Col, G_FAM.Row);
    FAM := THCritMaskEdit.Create (Self);
    FAM.Parent := G_FAM;
    FAM.Top := Coord.Top;
    FAM.Left := Coord.Left;
    FAM.Width := 3; FAM.Visible := False;
    FAM.DataType := 'GCTARIFARTICLE';
    GetFamilleRecherche (FAM) ;
    if FAM.Text <> '' then G_Fam.Cells[G_FAM.Col,G_FAM.Row]:= FAM.Text;
    FAM.Destroy;
    END ;
end;

procedure TFTarifTiersMode.G_FamCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
FormateZoneSaisie (ACol,ARow);
    if ACol = SFA_Tiers then TraiterTiers (ACol, ARow) else
       if ACol = SFA_CatTiers then TraiterCatTiers (ACol, ARow) else
          if ACol = SFA_CatArt then TraiterCatArticle (ACol, ARow) else
             //if ACol = SFA_Lib then TraiterLibelle (ACol, ARow) else
                //if ACol = SFA_QInf then TraiterBorneInf (ACol, ARow) else
                   //if ACol = SFA_QSup then TraiterBorneSup (ACol, ARow) else
                      if ACol = SFA_Rem then TraiterRemise (ACol, ARow) ;
end;


{==============================================================================================}
{=============================== Champs ligne =================================================}
{==============================================================================================}
procedure TFTarifTiersMode.TraiterTiers (ACol, ARow:Integer ) ;
Var RechTiers : T_RechArt ;
    OkTiers   : Boolean ;
    QTiers : TQuery;
    i_Rep: Integer;
    ioerr : TIOErr ;
    SQL: String ;
BEGIN
if G_FAM.Cells[ACol,ARow] = '' then exit ;
OkTiers:=False ;
RechTiers := TrouverTiers (G_FAM.Cells[ACol,ARow], TOBTiers);
Case RechTiers of
        traOk : OkTiers:=True ;
     traAucun : BEGIN
                // Recherche sur code via LookUp ou Recherche avancée
                    //DispatchRecherche (GF_TIERS, 2, 'T_NATUREAUXI="' + NatureTiers + '"',
                   //                    'T_TIERS=' + Trim (Copy (G_FAM.Cells[ACol,ARow], 1, 17)), '');
                    if G_FAM.Cells[ACol,ARow] <> '' then
                    BEGIN
                    QTiers:=OpenSQL('SELECT * FROM TIERS WHERE T_TIERS="'+G_FAM.Cells[ACol,ARow]+'"',true);
                    if not QTiers.eof then TOBTiers.SelectDB('',QTiers);
                    Ferme(QTiers);
                    OkTiers := True;
                    END;
                END ;
     End ; // Case

if (OkTiers) then
    BEGIN
    //AffecteEntete;
    if CodeTiers <> G_FAM.Cells[ACol,ARow] then
      BEGIN
      //if lModif then
      //   begin
          if (G_Fam.Cells[SFA_Rem,ARow]='') then i_Rep:=7
          else i_Rep := QuestionTarifEnCours ;
          Case i_Rep of
              mrYes    : BEGIN
                         ioerr := Transactions (ValideTarif, 2);
                         Case ioerr of
                             oeOk      : ;
                             oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
                             oeSaisie  : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
                             END ;
                         CodeTiers := G_FAM.Cells[ACol,ARow];
                         //CodeDevise := GF_DEVISE.Value;
                         Transactions (ChargeTarif, 1);
                         //lModif:=False ;
                         END;
              mrNo     : BEGIN
                         if G_Fam.Cells[SFA_Tiers,Arow]<>'' then
                          SQL:='SELECT * FROM TARIF WHERE GF_TIERS="'+G_Fam.Cells[SFA_Tiers,Arow]+'" AND GF_DEVISE="' + CodeDevise + '" '+
                          'AND GF_TARFMODE="'+CodeTarifMode+'" AND GF_DEPOT="'+CodeDepot+'" ORDER BY GF_BORNEINF';
                          If ExisteSQL(SQL) then //AND not (G_Fam.Cells[SFA_Tiers,Arow-1]<> CodeTiers) And (ARow>1) then
                             Begin
                               TarifExistant:=True ;
                               CodeTiers := G_FAM.Cells[ACol,ARow];
                               //CodeDevise := GF_DEVISE.Value;
                               Transactions (ChargeTarif, 1);
                               TOBTarfFam.SetAllModifie (False);
                               TOBTarifDel.SetAllModifie (False);
                               //lModif:=False ;
                             End else
                             begin
                               if ExisteCategorie ('GCTIERSCLI', G_Fam.Cells [ACol, ARow]) then
                               InitLaLigne (ARow)
                                 else Begin
                                 // message Catégorie Tiers inexistante
                                 MsgBox.Execute (8,Caption,'') ;
                                 G_Fam.Cells [ACol, ARow] := '';
                                 G_Fam.Col:=SFA_Tiers ;
                                 end ;
                               InitLaLigne (ARow) ;
                               end ;
                         END;
              mrCancel : BEGIN
                         QTiers:=OpenSQL('SELECT * FROM TIERS WHERE T_TIERS="'+G_FAM.Cells[ACol,ARow]+'"',true);
                         if not QTiers.eof then TOBTiers.SelectDB('',QTiers);
                         Ferme(QTiers);
                         AffecteEntete;
                         END;
              END ;
      End else if G_Fam.Cells[SFA_Rem,ARow]='' then InitLaLigne (ARow) ;
    END ;
END;


procedure TFTarifTiersMode.TraiterCatTiers (ACol, ARow:Integer ) ;
Var i_Rep: Integer;
    ioerr : TIOErr;
    SQL: String ;
BEGIN
if G_FAM.Cells[Acol,Arow] = '' then  exit ;
if (CodeCatTiers <> G_FAM.Cells[Acol,Arow]) then //or (GF_DEVISE.Value <> CodeDevise)) then
    BEGIN
    //if lModif then
     //   begin
        if (G_Fam.Cells[SFA_Rem,ARow]='') then i_Rep:=7
        else i_Rep := QuestionTarifEnCours;
        Case i_Rep of
            mrYes    : BEGIN
                       ioerr := Transactions (ValideTarif, 2);
                       Case ioerr of
                          oeOk      : ;
                          oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
                          oeSaisie  : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
                          END ;
                       CodeCatTiers := G_FAM.Cells[Acol,Arow];
                       //CodeDevise := GF_DEVISE.Value;
                       Transactions (ChargeTarif, 1) ;
                       lModif := False;
                       END;
            mrNo     : BEGIN
                       if G_Fam.Cells[SFA_CatTiers,Arow]<>'' then
                        SQL:='SELECT * FROM TARIF WHERE GF_TARIFTIERS="'+G_Fam.Cells[SFA_CatTiers,Arow]+'" AND GF_TARIFARTICLE<>"" AND GF_DEVISE="' + CodeDevise + '" '+
                        'AND GF_TARFMODE="'+CodeTarifMode+'" AND GF_DEPOT="'+CodeDepot+'" ORDER BY GF_BORNEINF';
                       If ExisteSQL(SQL) then
                         Begin
                         TarifExistant:=True ;
                         CodeCatTiers := G_FAM.Cells[ACol,ARow];
                         //CodeDevise := GF_DEVISE.Value;
                         Transactions (ChargeTarif, 1);
                         TOBTarfFam.SetAllModifie (False);
                         TOBTarifDel.SetAllModifie (False);
                         lModif := False;
                         End else
                         begin
                         if ExisteCategorie ('TTTARIFCLIENT', G_Fam.Cells [ACol, ARow]) then
                            InitLaLigne (ARow)
                         else Begin
                           // message Catégorie Tiers inexistante
                           MsgBox.Execute (7,Caption,'') ;
                           G_Fam.Cells [ACol, ARow] := '';
                           G_Fam.Col:=SFA_CatTiers ;
                           end ;
                         end ;
                       END;
            mrCancel : BEGIN
                       TOBCatTiers.SelectDB ('"' + CodeCatTiers + '"', Nil) ;
                       //GF_DEVISE.Value := CodeDevise;
                       AffecteEntete;
                       END;
            END ;
    END;
END;

procedure TFTarifTiersMode.TraiterCatArticle (ACol, ARow:Integer ) ;
var TOBL : TOB;
    B_NewLine : Boolean;
    St,SQL : string;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
if TOBL.GetValue ('GF_TARIFARTICLE') = '' then B_NewLine :=True else B_NewLine := False;
St := G_Fam.Cells [ACol, ARow];
if St <> '' then
    BEGIN
    if ExisteCategorie ('GCTARIFARTICLE', St) then
        BEGIN
        TOBL.PutValue ('GF_TARIFARTICLE', St);
        SQL:='SELECT * FROM TARIF WHERE GF_TARIFTIERS="'+G_Fam.Cells[SFA_CatTiers,Arow]+'" AND GF_TIERS="'+G_Fam.Cells[SFA_Tiers,Arow]+'" AND GF_DEVISE="' + CodeDevise + '" '+
        'AND GF_TARFMODE="'+CodeTarifMode+'" AND GF_DEPOT="'+CodeDepot+'" AND GF_TARIFARTICLE="'+St+'" ORDER BY GF_BORNEINF';
        If (ExisteSQL(SQL)) and (not TarifExistant) then
          begin
          MsgBox.Execute (10,Caption,'') ;
          G_Fam.Col:=SFA_CatArt ;
          exit ;
        end ;
        //lModif := True;
        END else
        BEGIN
        // message Catégorie Article inexistante
        MsgBox.Execute (5,Caption,'') ;
        G_Fam.Cells [ACol, ARow] := TOBL.GetValue ('GF_TARIFARTICLE');
        G_Fam.Col:=SFA_CatArt ;
        END;
    END else if Not B_Newline then
            G_Fam.Cells [ACol, ARow] := TOBL.GetValue ('GF_TARIFARTICLE');
END;

procedure TFTarifTiersMode.TraiterRemise (ACol, ARow:Integer ) ;
var TOBL : TOB;
    St : string;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
G_Fam.Cells [ACol, ARow]:=ModifFormat(G_Fam.Cells [ACol, ARow]);
St := G_Fam.Cells [ACol, ARow];
TOBL.PutValue ('GF_CALCULREMISE', St);
if St<>'' then TOBL.PutValue ('GF_REMISE', StrToFloat(St))
   else TOBL.PutValue ('GF_REMISE', 0) ;
AffichePied (ARow);
END;

{==============================================================================================}
{=============================== Validation =================================================}
{==============================================================================================}
procedure TFTarifTiersMode.ValideTarif ;
var i: Integer ;
begin
i:=1 ;
while not LigneVide(i) do
  begin
    if valeur(G_Fam.Cells[SFA_Rem,i])=0 then
      begin
      MsgBox.Execute (9,Caption,'') ;
      G_Fam.Col:=SFA_Rem ;
      G_Fam.Row:=i ;
      exit ;
      end else
      i:=i+1 ;
  end ;
if Not SortDeLaLigne then Exit ;
TOBTarifDel.DeleteDB (False);
VerifLesTOB;
TOBTarif.InsertOrUpdateDB(False) ;
TOBTarifMode.InsertOrUpdateDB(False) ;
InitialiseGrille;
DepileTOBLigne;
LibereTOBMode ;
end;

procedure TFTarifTiersMode.VerifLesTOB ;
var i_ind : integer;
    Q : TQuery ;
    MaxTarif : Longint ;
BEGIN
Q := OpenSQL ('SELECT MAX(GF_TARIF) FROM TARIF', TRUE) ;
if Q.EOF then MaxTarif := 1 else MaxTarif := Q.Fields[0].AsInteger + 1 ;
Ferme(Q) ;
for i_ind := TOBTarfFam.Detail.count - 1 Downto 0 do
    BEGIN
    if LigneVide (i_ind + 1) then
        BEGIN
        TOBTarfFam.Detail[i_ind].Free ;
        END else
        BEGIN
        if TOBTarfFam.Detail[i_ind].GetValue ('GF_TARIF') = 0 then
            BEGIN
            TOBTarfFam.Detail[i_ind].PutValue ('GF_TARIF', MaxTarif);
            Inc (MaxTarif);
            END;
        CalcPriorite (TOBTarfFam.Detail[i_ind]);
        END;
    END;
END;

{==============================================================================================}
{=============================== Evenements du pied ===========================================}
{==============================================================================================}
procedure TFTarifTiersMode.GF_CASCADEREMISEChange(Sender: TObject);
var Row : integer;
TOBL : TOB ;
begin
Row := G_FAM.Row;
if Row > 0 then
    BEGIN
    TOBL := GetTOBLigne (Row); if TOBL=nil then exit;
    TOBL.PutValue ('GF_CASCADEREMISE', GF_CASCADEREMISE.Value);
    END;
end ;

{==============================================================================================}
{============================= Manipulation du pied ===========================================}
{==============================================================================================}
procedure TFTarifTiersMode.InitialisePied ;
BEGIN
GF_CASCADEREMISE.Text := '';
GF_REMISE.Text := '';
//AffichePrixCon;
END;

{procedure TFTarifTiersMode.AffichePrixCon ;
BEGIN
if (CodeDevise<>V_PGI.DevisePivot) or (PFAMARTICLE.Visible) then
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
END;   }

procedure TFTarifTiersMode.AffichePied (ARow : Longint) ;
var TOBL : TOB;
    FF   : TForm;
BEGIN
TOBL:=GetTOBLigne(ARow) ; if TOBL=Nil then Exit ;
FF := TForm (PPIED.Owner);
TOBL.PutEcran (FF, PPIED);
END;

{==============================================================================================}
{=============================== Conditions tarifaires ========================================}
{==============================================================================================}
procedure TFTarifTiersMode.InitComboChamps ;
begin
SourisSablier ;
RemplitComboChamps('TIERS',FComboTIE) ;
RemplitComboChamps('ARTICLE',FComboART) ;
RemplitComboChamps('LIGNE',FComboLIG) ;
RemplitComboChamps('PIECE',FComboPIE) ;
SourisNormale ;
end ;

procedure TFTarifTiersMode.RemplitComboChamps(NomTable : string ; FCombo : THValComboBox) ;
var Pref : String ;
    Q    : TQuery ;
begin
Pref := TableToPrefixe(NomTable) ;
Q:=OpenSQL('SELECT DH_NOMCHAMP,DH_LIBELLE FROM DECHAMPS WHERE DH_PREFIXE="'+
           Pref + '" AND DH_CONTROLE LIKE "%T%"',True) ;
While not Q.EOF do
   begin
   if Trim(Q.FindField('DH_LIBELLE').AsString)='' then FCombo.Items.Add(Q.FindField('DH_NOMCHAMP').AsString)
                                                  else FCombo.Items.Add(Q.FindField('DH_LIBELLE').AsString) ;
   FCombo.Values.Add(Q.FindField('DH_NOMCHAMP').AsString) ;
   Q.Next ;
   end ;
Ferme(Q) ;
end ;

procedure TFTarifTiersMode.EffaceGrid ;
var lig, col : Integer ;
begin
for lig := 1 to G_COND.RowCount - 1 do
   begin
   for col := 0 to G_COND.ColCount - 1 do G_COND.Cells[col, lig] := '' ;
   end ;
G_COND.Row := 1 ;
end ;

procedure TFTarifTiersMode.AfficheCondTarf (ARow : Longint) ;
var TOBL : TOB;
BEGIN
TOBL:=GetTOBLigne(ARow) ; if TOBL=Nil then Exit ;
GetConditions (TOBL) ;
END;

function TFTarifTiersMode.ValueToItem(CC: THValcomboBox ; St : String) : String ;
var i_ind : Integer ;
begin
i_ind := CC.Values.IndexOf(St) ;
if i_ind >= 0 then result := CC.Items[i_ind] else result := '' ;
end ;

procedure TFTarifTiersMode.GetConditions (TOBL : TOB) ;
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
   end ;
end ;

Function TrouverTiers(CodeTiers:String ; TOBTiers: TOB): T_RechArt ;
var Q    :TQuery;
    Etat : string ;
BEGIN
Result := traAucun;
if CodeTiers = '' Then exit;
Q := OpenSQL('Select * from Tiers Where T_TIERS="' +
             CodeTiers + '" AND T_NATUREAUXI = "CLI" ',True) ;
if Not Q.EOF then
    BEGIN
    TOBTiers.SelectDB('',Q);
    Etat:=TOBTiers.GetValue('T_NATUREAUXI') ;
    if Etat='CLI' then Result:=traOk;
    END ;
Ferme(Q) ;
END;

procedure TFTarifTiersMode.TCONDTARFClose(Sender: TObject);
begin
BVoirCond.Down := False ;
end;

procedure TFTarifTiersMode.TraiterEtablissement (ARow : integer);
var TOBL : TOB;
    TempDepot : string;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
TempDepot:=GF_DEPOT.Value ;
//if (TempDepot='') or (TempDepot='<Aucun>') then
//   begin
//   CodeDepot:='...' ;
//   TOBL.PutValue ('GF_DEPOT','...') ;
//   end else TOBL.PutValue ('GF_DEPOT',TempDepot) ;
TOBL.PutValue ('GF_DEPOT',TempDepot) ;
TOBL.PutValue ('GF_SOCIETE', TempDepot) ;
TOBL.SetAllModifie(True) ;
//TOBL.PutValue ('GF_SOCIETE', TOBL.GetValue('GF_DEPOT')) ;
END ;

procedure TFTarifTiersMode.MAJChampEntete ;
Begin
If SavEntete then exit ;
ExecuteSQL('Update TarifTypMode SET GFT_DERUTILISE="-" Where GFT_DERUTILISE="X"') ;
ExecuteSQL('Update TarifTypMode SET GFT_DERUTILISE="X" Where GFT_CODETYPE="'+CodeType+'"') ;
ExecuteSQL('Update TarifPer SET GFP_DERUTILISE="-" Where GFP_DERUTILISE="X"') ;
ExecuteSQL('Update TarifPer SET GFP_DERUTILISE="X" Where GFP_CODEPERIODE="'+CodePeriode+'"') ;
SavEntete:=True;
end ;

procedure TFTarifTiersMode.ChargeTarifDepuisMul ;
Var Q: TQuery ;
    i_ind : integer;
    Select : string;
    TobInit: TOB ;
Begin
TobTarifMode:=TOB.Create('TARIFMODE',NIL,-1) ;
TobInit:=TheTob ;
TheTob:=nil ;
TOBCatTiers.InitValeurs ;
TOBTiers.InitValeurs ;
IdTarif:=TobInit.GetValue('_CodeTarif') ;
CodeTiers:=TobInit.GetValue('_CodeTiers') ;
CodeCatTiers:=TobInit.GetValue('_TarifTiers') ;
if CodeTiers<>'' then
  begin
  // Charge tiers
  TOBTiers.SelectDB('"' + CodeTiers + '"', Nil) ;

  // Lecture Quantitatif
  Select := 'SELECT * FROM TARIF WHERE GF_TIERS="'+CodeTiers+'" AND GF_TARIF="' + IdTarif + '" '+
            ' ORDER BY GF_BORNEINF';
            //'ORDER BY GF_DEPOT, GF_TARIFARTICLE, GF_BORNEINF';
  Q := OpenSQL(Select,True) ;
  TOBTarfFam.LoadDetailDB('TARIF','','',Q,False) ;
  TOBTarfFam.SetAllModifie(False);
  CodeTarifMode:=Q.FindField('GF_TARFMODE').AsString ;
  CodeDepot:=Q.FindField('GF_DEPOT').AsString ;
  Ferme(Q) ;
  end
else if CodeCatTiers<>'' then
  begin
  // Charge catégorie tiers
  TOBCatTiers.SelectDB ('"' + CodeCatTiers + '"', Nil) ;
  // Lecture Quantitatif
  Select := 'SELECT * FROM TARIF WHERE GF_TARIFTIERS="'+CodeCatTiers+'" AND GF_TARIF="' + IdTarif + '" ' +
            ' ORDER BY GF_BORNEINF';
            //'ORDER BY GF_DEPOT, GF_TARIFARTICLE, GF_BORNEINF';
  Q := OpenSQL(Select,True) ;
  TOBTarfFam.LoadDetailDB('TARIF','','',Q,False) ;
  TOBTarfFam.SetAllModifie(False);
  CodeTarifMode:=Q.FindField('GF_TARFMODE').AsString ;
  CodeDepot:=Q.FindField('GF_DEPOT').AsString ;
  Ferme(Q) ;
  end ;
// Affichage entête
Q:=OpenSql('Select * from TARIFMODE where GFM_TARFMODE="'+CodeTarifMode+'"',True) ;
if not Q.EOF then
  begin
  TobTarifMode.SelectDB('',Q) ;
  ferme(Q) ;
  end ;
CodeType:=TobTarifMode.GetValue('GFM_TYPETARIF') ;
if CodeType='...' then
  begin
  GF_DEPOT.Visible:=False ;
  TGF_DEPOT.Visible:=False ;
  end else
  begin
  GF_DEPOT.Visible:=True ;   
  TGF_DEPOT.Visible:=True ;
  end ;
_TypTarifMode.Value:=CodeType ;
CodePeriode:=TobTarifMode.GetValue('GFM_PERTARIF') ;
_PerTarifMode.Value:=CodePeriode ;
GF_DEPOT.Value:=CodeDepot ;
GF_DATEDEBUT.Text:=TobTarifMode.GetValue('GFM_DATEDEBUT') ;
GF_DATEFIN.Text:=TobTarifMode.GetValue('GFM_DATEFIN') ;
CodeDevise:=TobTarifMode.GetValue('GFM_DEVISE') ;
GF_DEVISE.Text:=RechDom('TTDEVISE',CodeDevise,False) ;
CodeArrondi:=TobTarifMode.GetValue('GFM_ARRONDI') ;
//Interdire la modification de l'entête
_TYPTARIFMODE.Enabled:=False ;
_PERTARIFMODE.Enabled:=False ;
GF_DEPOT.Enabled:=False ;
//
// Affichage
for i_ind:=0 to TOBTarfFam.Detail.Count-1 do
    BEGIN
    AfficheLaLigne (i_ind + 1) ;
    END ;
ModifMul:=True ;
end ;

procedure TFTarifTiersMode.BAbandonClick(Sender: TObject);
begin
Close ;
if FClosing and IsInside(Self) then THPanel(parent).CloseInside ;
end;

Procedure AGLEntreeTarifTiersMode ( Parms : array of variant ; nb : integer ) ;
BEGIN
EntreeTarifTiersMode(taModif,TRUE);
END ;

Initialization
RegisterAglProc('EntreeTarifTiersMode',False,0,AGLEntreeTarifTiersMode) ;

end.

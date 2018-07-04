unit TarifCatArt;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, ComCtrls, ExtCtrls, HPanel, HSysMenu, hmsgbox, Menus, StdCtrls,
  Mask, Hctrls, Grids, UIUtil, Hent1, TarifUtil, SaisUtil,UtilPGI,AGLInit,Hqry,
{$IFDEF EAGLCLIENT}
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  UTOB,math, HRichEdt, HRichOLE,
{$IFNDEF CCS3}
  TarifCond,
{$ENDIF}
  M3FP, TntGrids, TntComCtrls, TntStdCtrls, TntExtCtrls ;

Function EntreeTarifCatArt (Action : TActionFiche) : boolean ;

type
  TFTarifCatArt = class(TForm)
    PENTETE: THPanel;
    Dock971: TDock97;
    Toolbar972: TToolWindow97;
    BAide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BValider: TToolbarButton97;
    POPZ: TPopupMenu;
    InfArticle: TMenuItem;
    MsgBox: THMsgBox;
    HMTrad: THSystemMenu;
    FindLigne: TFindDialog;
    BChercher: TToolbarButton97;
    PPIED: THPanel;
    GF_TARIFARTICLE: THValComboBox;
    TGF_TARIFARTICLE: THLabel;
    GF_DEVISE: THValComboBox;
    TGF_DEVISE: THLabel;
    GF_CASCADEREMISE: THValComboBox;
    TGF_CASCADEREMISE: THLabel;
    TGF_REMISE: THLabel;
    GF_REMISE: THCritMaskEdit;
    HMess: THMsgBox;
    PTITRE: THPanel;
    PTARIF: THPanel;
    PCATEGORIE: THPanel;
    HTitre: THMsgBox;
    G_CatA: THGrid;
    G_CatC: THGrid;
    CONDAPPLIC: THRichEditOLE;
    CBCATCLI: TCheckBox;
    TCONDTARF: TToolWindow97;
    G_COND: THGrid;
    FComboTIE: THValComboBox;
    FComboART: THValComboBox;
    FComboLIG: THValComboBox;
    FComboPIE: THValComboBox;
    GF_CONDAPPLIC: THRichEditOLE;
    FTable: THValComboBox;
    FOpe: THValComboBox;
    BCollerCond: TToolbarButton97;
    BCopierCond: TToolbarButton97;
    BCondAplli: TToolbarButton97;
    BVoirCond: TToolbarButton97;
    TTYPETARIF: THLabel;
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure G_CatAEnter(Sender: TObject);
    procedure G_CatARowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_CatARowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_CatACellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_CatACellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_CatAElipsisClick(Sender: TObject);
    procedure GF_TARIFARTICLEChange(Sender: TObject);
    procedure GF_DEVISEChange(Sender: TObject);
    procedure GF_REMISEChange(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure FindLigneFind(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure G_CatCEnter(Sender: TObject);
    procedure G_CatCRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_CatCRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_CatCCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_CatCCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_CatCElipsisClick(Sender: TObject);
    procedure CBCATCLIClick(Sender: TObject);
    procedure BCondAplliClick(Sender: TObject);
    procedure BCopierCondClick(Sender: TObject);
    procedure BCollerCondClick(Sender: TObject);
    procedure BVoirCondClick(Sender: TObject);
    procedure TCONDTARFClose(Sender: TObject);
    procedure GF_CASCADEREMISEChange(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    { Déclarations privées }
    iTableLigne : integer ;
    FindDebut,FClosing : Boolean;
    StCellCur,LesColCatA, LesColCatC : String ;
    ColsInter : Array of boolean ;
    DEV       : RDEVISE ;
// Objets mémoire
    TOBTarif, TOBTarfCatA, TOBTarfCatC, TOBTarifDel, TOBArt  : TOB;
// Menu
    procedure AffectMenuCondApplic (G_GRID : THGrid; ttd : T_TableTarif) ;
// Actions liées au Grid
    procedure EtudieColsListe ;
    procedure PostDrawCell(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    procedure FormateZoneSaisie (ACol,ARow : Longint; ttd : T_TableTarif) ;
    procedure InsertLigne (ARow : Longint; ttd : T_TableTarif) ;
    procedure SupprimeLigne (ARow : Longint; ttd : T_TableTarif) ;
    procedure SupprimeTOBTarif (ARow : Longint; ttd : T_TableTarif) ;
    Function  GrilleModifie : Boolean;
    Function  SortDeLaLigne (ttd : T_TableTarif) : boolean ;
// Initialisations
    Procedure InitialiseCols ;
    Function  WhereTarifCatArt (ttd : T_TableTarif) : String ;
    procedure LoadLesTOB ;
    procedure ChargeTarif ;
    procedure ErgoGCS3 ;
// ENTETE
    Procedure InitialiseEntete ;
    Procedure AffecteEntete ;
    Function  QuestionTarifEnCours : Integer;
// LIGNES
    Procedure InitLaLigne (ARow : integer; ttd : T_TableTarif) ;
    Function  GetTOBLigne (ARow : integer;  ttd : T_TableTarif) : TOB ;
    procedure AfficheLaLigne (ARow : integer;  ttd : T_TableTarif) ;
    procedure InitialiseGrille ;
    procedure InitialiseLigne (ARow : integer; ttd : T_TableTarif) ;
    Procedure DepileTOBLigne;
    Procedure CreerTOBLigne (ARow : integer; ttd : T_TableTarif);
    Function  LigneVide ( ARow : integer; ttd : T_TableTarif) : Boolean;
    Procedure PreAffecteLigne (ARow : integer; ttd : T_TableTarif);
// CHAMPS LIGNES
    procedure TraiterDepot (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterCatTiers (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterLibelle (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterPrix (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterRemise (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterDateDeb (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterDateFin (ACol, ARow : integer; ttd : T_TableTarif);
// PIED
    Procedure AffichePied (ttd : T_TableTarif) ;
// Boutons
// Validations
    procedure ValideTarif;
    procedure VerifLesTOB;
// Conditions tarifaires
    procedure InitComboChamps ;
    procedure RemplitComboChamps(NomTable : String ; FCombo : THValComboBox) ;
    procedure EffaceGrid ;
    Procedure AfficheCondTarf (ARow : Longint; ttd : T_TableTarif) ;
    function  ValueToItem(CC : THValComboBox ; St : String) : String ;
    procedure GetConditions (TOBL : TOB) ;
  public
    { Déclarations publiques }
    CodeTarifArticle : string;
    CodeDevise       : string;
    Action           : TActionFiche ;
  end;

Const SCA_Depot     : integer = 0 ;
      SCA_Lib       : integer = 0 ;
      SCA_Px        : integer = 0 ;
      SCA_Rem       : integer = 0 ;
      SCA_Datedeb   : integer = 0 ;
      SCA_Datefin   : integer = 0 ;

Const SCC_Depot     : integer = 0 ;
      SCC_Lib       : integer = 0 ;
      SCC_Cat       : integer = 0 ;
      SCC_Px        : integer = 0 ;
      SCC_Rem       : integer = 0 ;
      SCC_Datedeb   : integer = 0 ;
      SCC_Datefin   : integer = 0 ;

implementation
uses
   CbpMCD
   ,CbpEnumerator
;

{$R *.DFM}

Function EntreeTarifCatArt (Action : TActionFiche) : boolean ;
var FF : TFTarifCatArt ;
    PPANEL  : THPanel ;
begin
Result:=True;
SourisSablier;
FF := TFTarifCatArt.Create(Application) ;
FF.Action:=Action ;
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
END ;

procedure TFTarifCatArt.PostDrawCell(ACol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
begin
//
end ;

{==============================================================================================}
{======================================= Initialisations ======================================}
{==============================================================================================}
Procedure TFTarifCatArt.InitialiseCols ;
BEGIN
SCA_Depot:=-1 ; SCA_Lib:=-1 ; SCA_Px:=-1 ;
SCA_Rem:=-1 ; SCA_Datedeb:=-1 ;  SCA_Datefin:=-1 ;
SCC_Depot:=-1 ; SCC_Lib:=-1 ; SCC_Cat:=-1 ; SCC_Px:=-1 ;
SCC_Rem:=-1 ; SCC_Datedeb:=-1 ;  SCC_Datefin:=-1 ;
END ;

Function TFTarifCatArt.WhereTarifCatArt (ttd : T_TableTarif) : String ;
Var St : String ;
BEGIN
St:='' ;
Case ttd of
   ttdCatA : St:= ' GF_TARIFARTICLE="'+CodeTarifArticle+'" AND GF_TIERS="" AND ' +
                    ' GF_TARIFTIERS="" AND GF_ARTICLE="" ' +
                    ' AND GF_QUANTITATIF="-" AND GF_DEVISE="' + CodeDevise + '" ';
   ttdCatC : St:= ' GF_TARIFARTICLE="'+CodeTarifArticle+'" AND GF_TIERS="" AND ' +
                    ' GF_TARIFTIERS<>"" AND GF_ARTICLE="" ' +
                    ' AND GF_QUANTITATIF="-" AND GF_DEVISE="' + CodeDevise + '" ';
   END ;
Result:=St ;
END ;

procedure TFTarifCatArt.LoadLesTOB ;
Var Q : TQuery ;
    i_ind : integer;
BEGIN
for i_ind := TOBTarfCatA.Detail.Count - 1 downto 0 do
    BEGIN
    TOBTarfCatA.Detail[i_ind].Free;
    END;
for i_ind := TOBTarfCatC.Detail.Count - 1 downto 0 do
    BEGIN
    TOBTarfCatC.Detail[i_ind].Free;
    END;

// Lecture Quantitatif

Q := OpenSQL('SELECT * FROM TARIF WHERE '+ WhereTarifCatArt (ttdCatA)+
             ' ORDER BY GF_DEPOT',True,-1,'',true) ;
TOBTarfCatA.LoadDetailDB('TARIF','','',Q,False) ;
Ferme(Q) ;
Q := OpenSQL('SELECT * FROM TARIF WHERE '+ WhereTarifCatArt (ttdCatC)+
             ' ORDER BY GF_DEPOT, GF_TARIFTIERS',True,-1,'',true) ;
TOBTarfCatC.LoadDetailDB('TARIF','','',Q,False) ;
Ferme(Q) ;
END ;

procedure TFTarifCatArt.ChargeTarif ;
var i_ind : integer;
BEGIN
LoadLesTOB ;
for i_ind:=0 to TOBTarfCatA.Detail.Count-1 do
    BEGIN
    // Affichage
    AfficheLaLigne (i_ind + 1, ttdCatA) ;
    END ;
for i_ind:=0 to TOBTarfCatC.Detail.Count-1 do
    BEGIN
    // Affichage
    AfficheLaLigne (i_ind + 1, ttdCatC) ;
    END ;
END ;

{==============================================================================================}
{=============================== Evènements de la Form ========================================}
{==============================================================================================}
procedure TFTarifCatArt.ErgoGCS3 ;
BEGIN
{$IFDEF CCS3}
BVoirCond.Visible:=False   ; BCondAplli.Visible:=False ;
BCopierCond.Visible:=False ; BCollerCond.Visible:=False ;
{$ENDIF}
END ;

procedure TFTarifCatArt.FormShow(Sender: TObject);
begin
If (ctxAffaire in V_PGI.PGIContexte) then  G_CatA.ListeParam:='AFTARIFREM' 
else G_CatA.ListeParam:='GCTARIFREM' ;
G_CatA.PostDrawCell:=PostDrawCell ;
If (ctxAffaire in V_PGI.PGIContexte) then  G_CatC.ListeParam:='AFTARIFCT'
else G_CatC.ListeParam:='GCTARIFCT' ;
G_CatC.PostDrawCell:=PostDrawCell ;
EtudieColsListe ;
HMTrad.ResizeGridColumns (G_CatA) ;
HMTrad.ResizeGridColumns (G_CatC) ;
AffecteGrid (G_CatA,Action) ;
AffecteGrid (G_CatC,Action) ;
PTARIF.Visible := True;
PCATEGORIE.Visible := False;
PTITRE.Caption := HTitre.Mess[0] ;
TTYPETARIF.Caption := HTitre.Mess[3] ;
CBCATCLI.Checked := False ;
CodeDevise := V_PGI.DevisePivot;
DEV.Code := CodeDevise ; GetInfosDevise (DEV) ;
InitialiseEntete ;
InitComboChamps ;
ErgoGCS3 ;
end;

procedure TFTarifCatArt.FormCreate(Sender: TObject);
begin
G_CatA.RowCount := NbRowsInit ;
G_CatC.RowCount := NbRowsInit ;
StCellCur := '' ;
iTableLigne := PrefixeToNum('GF') ;
TOBTarif := TOB.Create ('', Nil, -1) ;
TOBTarfCatA := TOB.Create ('', TOBTarif, 0) ;
TOBTarfCatC := TOB.Create ('', TOBTarif, 1) ;
TOBTarifDel := TOB.Create ('', Nil, -1) ;
TOBArt := TOB.Create ('ARTICLE', Nil, -1) ;
InitialiseCols;
FClosing:=False ;
end;

procedure TFTarifCatArt.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
if Action=taConsult then Exit ;
if GrilleModifie then
   BEGIN
   if MsgBox.Execute(6,Caption,'')<>mrYes then CanClose:=False ;
   END ;
end;

procedure TFTarifCatArt.FormClose(Sender: TObject; var Action: TCloseAction);
begin
G_CatA.VidePile(True) ;
G_CatC.VidePile(True) ;
TOBTarif.Free ; TOBTarif:=Nil ;
TOBArt.Free ; TOBArt:=Nil ;
TOBTarifDel.Free; TOBTarifDel:=Nil ;
if IsInside(Self) then Action:=caFree ;
FClosing:=True ; 
end;

procedure TFTarifCatArt.FormKeyDown(Sender: TObject; var Key: Word;
                                     Shift: TShiftState);
var FocusGrid : Boolean;
    ttd : T_TableTarif;
    ARow : Longint;
BEGIN
FocusGrid := False;
if(Screen.ActiveControl = G_CatA) then
    BEGIN
    FocusGrid := True;
    ttd := ttdCatA;
    ARow := G_CatA.Row;
    END else
    if (Screen.ActiveControl = G_CatC) then
        BEGIN
        FocusGrid := True;
        ttd := ttdCatC;
        ARow := G_CatC.Row;
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
procedure TFTarifCatArt.AffectMenuCondApplic (G_GRID : THGrid; ttd : T_TableTarif) ;
BEGIN
if (LigneVide (G_GRID.Row, ttd))then
    BEGIN
    BCondAplli.Enabled := False ; BCopierCond.Enabled := False; BCollerCond.Enabled := False ;
    END else
    BEGIN
    BCondAplli.Enabled := True ; BCopierCond.Enabled := True;
    if CONDAPPLIC.Text = '' then BCollerCond.Enabled := False
                            else BCollerCond.Enabled := True ;
    END;
END;

{==============================================================================================}
{=============================== Actions liées au Grid ========================================}
{==============================================================================================}
procedure TFTarifCatArt.EtudieColsListe ;
Var NomCol,LesCols : String ;
    icol,ichamp, i_ind : integer ;
	Mcd : IMCDServiceCOM;
BEGIN
MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();
//
G_CatA.ColWidths[0]:=0 ;
G_CatC.ColWidths[0]:=0 ;
SetLength(ColsInter,G_CatA.ColCount) ;
for i_ind:=Low(ColsInter)to High(ColsInter) do ColsInter[i_ind]:=False ;
LesCols:=G_CatA.Titres[0] ; LesColCatA:=LesCols ; icol:=0 ;
Repeat
 NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
 if NomCol<>'' then
    BEGIN
		if assigned(mcd.getField(Nomcol)) then
       BEGIN
       if Pos('X',mcd.getField(NomCol).Control)>0 then ColsInter[icol]:=True ;
       if NomCol='GF_DEPOT'        then SCA_Depot:=icol else
       if NomCol='GF_LIBELLE'      then SCA_Lib:=icol else
       if NomCol='GF_PRIXUNITAIRE' then SCA_Px:=icol else
       if NomCol='GF_CALCULREMISE' then SCA_Rem:=icol else
       if NomCol='GF_DATEDEBUT'    then SCA_Datedeb:=icol else
       if NomCol='GF_DATEFIN'      then SCA_Datefin:=icol ;
       END ;
    END ;
 Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;

LesCols:=G_CatC.Titres[0] ; LesColCatC := LesCols ; icol:=0 ;
Repeat
 NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
 if NomCol<>'' then
    BEGIN
		if assigned(mcd.getField(Nomcol)) then
       BEGIN
       if NomCol='GF_DEPOT'        then SCC_Depot:=icol else
       if NomCol='GF_TARIFTIERS'   then SCC_Cat:=icol else
       if NomCol='GF_LIBELLE'      then SCC_Lib:=icol else
       if NomCol='GF_PRIXUNITAIRE' then SCC_Px:=icol else
       if NomCol='GF_CALCULREMISE' then SCC_Rem:=icol else
       if NomCol='GF_DATEDEBUT'    then SCC_Datedeb:=icol else
       if NomCol='GF_DATEFIN'      then SCC_Datefin:=icol ;
       END ;
    END ;
 Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;

END ;

procedure TFTarifCatArt.FormateZoneSaisie (ACol,ARow : Longint; ttd : T_TableTarif) ;
Var St,StC : String ;
BEGIN
Case ttd of
    ttdCatA : BEGIN
              St:=G_CatA.Cells[ACol,ARow] ; StC:=St ;
              if ACol=SCA_Depot then StC:=uppercase(Trim(St)) else
                if ACol=SCA_Px then StC:=StrF00(Valeur(St),V_PGI.OkDecP);
              G_CatA.Cells[ACol,ARow]:=StC ;
              END;
    ttdCatC : BEGIN
              St:=G_CatC.Cells[ACol,ARow] ; StC:=St ;
              if ((ACol=SCC_Depot) or (ACol=SCC_Cat)) then StC:=uppercase(Trim(St)) else
                if ACol=SCC_Px then StC:=StrF00(Valeur(St),V_PGI.OkDecP);
              G_CatC.Cells[ACol,ARow]:=StC ;
              END;
    END;
END ;

procedure TFTarifCatArt.InsertLigne (ARow : Longint; ttd : T_TableTarif) ;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if LigneVide (ARow, ttd) then exit;
Case ttd of
    ttdCatA : BEGIN
              if (ARow > TOBTarfCatA.Detail.Count) then Exit;
              G_CatA.CacheEdit; G_CatA.SynEnabled := False;
              TOB.Create ('TARIF', TOBTarfCatA, ARow-1) ;
              G_CatA.InsertRow (ARow); G_CatA.Row := ARow;
              InitialiseLigne (ARow, ttd) ;
              PreAffecteLigne (ARow, ttd);
              G_CatA.MontreEdit; G_CatA.SynEnabled := True;
              AffectMenuCondApplic (G_CatA, ttd);
              AfficheCondTarf (G_CatA.Row, ttd);
              END;
    ttdCatC : BEGIN
              if (ARow > TOBTarfCatC.Detail.Count) then Exit;
              G_CatC.CacheEdit; G_CatC.SynEnabled := False;
              TOB.Create ('TARIF', TOBTarfCatC, ARow-1) ;
              G_CatC.InsertRow (ARow); G_CatC.Row := ARow;
              InitialiseLigne (ARow, ttd) ;
              PreAffecteLigne (ARow, ttd);
              G_CatC.MontreEdit; G_CatC.SynEnabled := True;
              AffectMenuCondApplic (G_CatC, ttd);
              AfficheCondTarf (G_CatC.Row, ttd);
              END;
    END;
END;

procedure TFTarifCatArt.SupprimeLigne (ARow : Longint; ttd : T_TableTarif) ;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
Case ttd of
    ttdCatA : BEGIN
              if (ARow > TOBTarfCatA.Detail.Count) then Exit;
              G_CatA.CacheEdit; G_CatA.SynEnabled := False;
              G_CatA.DeleteRow (ARow);
              if (ARow = TOBTarfCatA.Detail.Count) then
                  CreerTOBLigne (ARow + 1, ttd);
              SupprimeTOBTarif (ARow, ttd);
              if G_CatA.RowCount < NbRowsInit then G_CatA.RowCount := NbRowsInit;
              G_CatA.MontreEdit; G_CatA.SynEnabled := True;
              AffectMenuCondApplic (G_CatA, ttd);
              AfficheCondTarf (G_CatA.Row, ttd);
              END;
    ttdCatC : BEGIN
              if (ARow > TOBTarfCatC.Detail.Count) then Exit;
              G_CatC.CacheEdit; G_CatC.SynEnabled := False;
              G_CatC.DeleteRow (ARow);
              if (ARow = TOBTarfCatC.Detail.Count) then
                  CreerTOBLigne (ARow + 1, ttd);
              SupprimeTOBTarif (ARow, ttd);
              if G_CatC.RowCount < NbRowsInit then G_CatC.RowCount := NbRowsInit;
              G_CatC.MontreEdit; G_CatC.SynEnabled := True;
              AffectMenuCondApplic (G_CatC, ttd);
              AfficheCondTarf (G_CatC.Row, ttd);
              END;
    END;
END;

procedure TFTarifCatArt.SupprimeTOBTarif (ARow : Longint; ttd : T_TableTarif) ;
Var i_ind: integer;
BEGIN
Case ttd of
    ttdCatA : BEGIN
              if TOBTarfCatA.Detail[ARow-1].GetValue ('GF_TARIF') <> 0 then
                  BEGIN
                  i_ind := TOBTarifDel.Detail.Count;
                  TOB.Create ('TARIF', TOBTarifDel, i_ind) ;
                  TOBTarifDel.Detail[i_ind].Dupliquer (TOBTarfCatA.Detail[ARow-1], False, True);
                  END;
              TOBTarfCatA.Detail[ARow-1].Free;
              END;
    ttdCatC : BEGIN
              if TOBTarfCatC.Detail[ARow-1].GetValue ('GF_TARIF') <> 0 then
                  BEGIN
                  i_ind := TOBTarifDel.Detail.Count;
                  TOB.Create ('TARIF', TOBTarifDel, i_ind) ;
                  TOBTarifDel.Detail[i_ind].Dupliquer (TOBTarfCatC.Detail[ARow-1], False, True);
                  END;
              TOBTarfCatC.Detail[ARow-1].Free;
              END;
    END;
END;

Function TFTarifCatArt.GrilleModifie : Boolean;
BEGIN
Result:=False ;
if Action=taConsult then Exit ;
Result:=(TOBTarfCatA.IsOneModifie) or (TOBTarfCatC.IsOneModifie) or (TOBTarifDel.IsOneModifie);
END;

Function TFTarifCatArt.SortDeLaLigne (ttd : T_TableTarif) : boolean ;
Var ACol,ARow : integer ;
    Cancel : boolean ;
BEGIN
Result:=False ;
Case ttd of
    ttdCatA : BEGIN
              ACol:=G_CatA.Col ; ARow:=G_CatA.Row ; Cancel:=False ;
              G_CatACellExit(Nil,ACol,ARow,Cancel) ; if Cancel then Exit ;
              G_CatARowExit(Nil,ACol,Cancel,False) ; if Cancel then Exit ;
              END;
    ttdCatC : BEGIN
              ACol:=G_CatC.Col ; ARow:=G_CatC.Row ; Cancel:=False ;
              G_CatCCellExit(Nil,ACol,ARow,Cancel) ; if Cancel then Exit ;
              G_CatCRowExit(Nil,ACol,Cancel,False) ; if Cancel then Exit ;
              END;
    END;
Result:=True ;
END ;

{==============================================================================================}
{=============================== Evènements de la Grid ========================================}
{==============================================================================================}
procedure TFTarifCatArt.G_CatAEnter(Sender: TObject);
var Cancel, Chg : Boolean;
    ACol, ARow : integer;
begin
Cancel := False; Chg := False;
G_CatARowEnter (Sender, G_CatA.Row, Cancel, Chg);
ACol := G_CatA.Col ; ARow := G_CatA.Row ;
G_CatACellEnter (Sender, ACol, ARow, Cancel);
end;

procedure TFTarifCatArt.G_CatARowEnter(Sender: TObject; Ou: Integer;
                                       var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
BEGIN
if Ou >= G_CatA.RowCount - 1 then G_CatA.RowCount := G_CatA.RowCount + NbRowsPlus ;;
ARow := Min (Ou, TOBTarfCatA.detail.count + 1);
if (ARow = TOBTarfCatA.detail.count + 1) AND (not LigneVide (ARow - 1, ttdCatA)) then
    BEGIN
    CreerTOBligne (ARow, ttdCatA);
    END;
if (LigneVide (ARow, ttdCatA)) AND (not LigneVide (ARow - 1, ttdCatA))then
    PreAffecteLigne (ARow, ttdCatA);
if Ou > TOBTarfCatA.detail.count then
    BEGIN
    G_CatA.Row := TOBTarfCatA.detail.count;
    END;
AffichePied (ttdCatA);
AfficheCondTarf (ARow, ttdCatA);
END;

procedure TFTarifCatArt.G_CatARowExit(Sender: TObject; Ou: Integer;
                                      var Cancel: Boolean; Chg: Boolean);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if LigneVide (Ou, ttdCatA) Then G_CatA.Row := Min (G_CatA.Row,Ou);
end;

procedure TFTarifCatArt.G_CatACellEnter(Sender: TObject; var ACol,
                                        ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if Not Cancel then
    BEGIN
    if (G_CatA.Col <> SCA_Depot) AND (G_CatA.Col <> SCA_Lib) AND
       (G_CatA.Cells [SCA_Lib,G_CatA.Row] = '') then G_CatA.Col := SCA_Lib;
    G_CatA.ElipsisButton := ((G_CatA.Col = SCA_Depot) or (G_CatA.col = SCA_Datedeb) or
                            (G_CatA.col = SCA_Datefin)) ;
    StCellCur := G_CatA.Cells [G_CatA.Col, G_CatA.Row] ;
    AffectMenuCondApplic (G_CatA, ttdCatA);
    END ;
end;

procedure TFTarifCatArt.G_CatACellExit(Sender: TObject; var ACol,
                                       ARow: Integer; var Cancel: Boolean);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
FormateZoneSaisie (ACol,ARow, ttdCatA);
if ACol = SCA_Depot then TraiterDepot (ACol, ARow, ttdCatA) else
    if ACol = SCA_Lib then TraiterLibelle (ACol, ARow, ttdCatA) else
        if ACol = SCA_Px then TraiterPrix (ACol, ARow, ttdCatA) else
            if ACol = SCA_Rem then TraiterRemise (ACol, ARow, ttdCatA) else
                if ACol = SCA_Datedeb then TraiterDateDeb (ACol, ARow, ttdCatA) else
                    if ACol = SCA_Datefin then TraiterDateFin (ACol, ARow, ttdCatA);
if Not Cancel then
    BEGIN
    END;
end;

procedure TFTarifCatArt.G_CatAElipsisClick(Sender: TObject);
Var DEPOT, DATE : THCritMaskEdit;
    Coord : TRect;
begin
if G_CatA.Col = SCA_Depot then
    BEGIN
    Coord := G_CatA.CellRect (G_CatA.Col, G_CatA.Row);
    DEPOT := THCritMaskEdit.Create (Self);
    DEPOT.Parent := G_CatA;
    DEPOT.Top := Coord.Top;
    DEPOT.Left := Coord.Left;
    DEPOT.Width := 3; DEPOT.Visible := False;
    DEPOT.DataType := 'GCDEPOT';
    GetDepotRecherche (DEPOT) ;
    if DEPOT.Text <> '' then G_CatA.Cells[G_CatA.Col,G_CatA.Row]:= DEPOT.Text;
    DEPOT.Destroy;
    END ;
if (G_CatA.Col = SCA_Datedeb) or (G_CatA.Col = SCA_Datefin) then
    BEGIN
    Coord := G_CatA.CellRect (G_CatA.Col, G_CatA.Row);
    DATE := THCritMaskEdit.Create (Self);
    DATE.Parent := G_CatA;
    DATE.Top := Coord.Top;
    DATE.left := Coord.Left;
    DATE.Width := 3; DATE.Visible := False;
    DATE.OpeType:=otDate;
    GetDateRecherche (TForm(DATE.Owner), DATE) ;
    if DATE.Text <> '' then G_CatA.Cells[G_CatA.Col,G_CatA.Row]:= DATE.Text;
    DATE.Destroy;
    END;
end;

procedure TFTarifCatArt.G_CatCEnter(Sender: TObject);
var Cancel, Chg : Boolean;
    ACol, ARow : integer;
begin
Cancel := False; Chg := False;
G_CatCRowEnter (Sender, G_CatC.Row, Cancel, Chg);
ACol := G_CatC.Col ; ARow := G_CatC.Row ;
G_CatCCellEnter (Sender, ACol, ARow, Cancel);
end;

procedure TFTarifCatArt.G_CatCRowEnter(Sender: TObject; Ou: Integer;
                                       var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
BEGIN
if Ou >= G_CatC.RowCount - 1 then G_CatC.RowCount := G_CatC.RowCount + NbRowsPlus ;;
ARow := Min (Ou, TOBTarfCatC.detail.count + 1);
if (ARow = TOBTarfCatC.detail.count + 1) AND (not LigneVide (ARow - 1, ttdCatC)) then
    BEGIN
    CreerTOBligne (ARow, ttdCatC);
    END;
if (LigneVide (ARow, ttdCatC)) AND (not LigneVide (ARow - 1, ttdCatC))then
    PreAffecteLigne (ARow, ttdCatC);
if Ou > TOBTarfCatC.detail.count then
    BEGIN
    G_CatC.Row := TOBTarfCatC.detail.count;
    END;
AffichePied (ttdCatC);
AfficheCondTarf (ARow, ttdCatC);
end;

procedure TFTarifCatArt.G_CatCRowExit(Sender: TObject; Ou: Integer;
                                      var Cancel: Boolean; Chg: Boolean);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if LigneVide (Ou, ttdCatC) Then G_CatC.Row := Min (G_CatC.Row,Ou);
end;

procedure TFTarifCatArt.G_CatCCellEnter(Sender: TObject; var ACol,
                                        ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if Not Cancel then
    BEGIN
    if (G_CatC.Col <> SCC_Depot) AND (G_CatC.Col <> SCC_Lib) AND (G_CatC.Col <> SCC_Cat) then
        BEGIN
        if G_CatC.Cells [SCC_Lib,G_CatC.Row] = '' then G_CatC.Col := SCC_Lib;
        if G_CatC.Cells [SCC_Cat,G_CatC.Row] = '' then G_CatC.Col := SCC_Cat;
        END;
    G_CatC.ElipsisButton := ((G_CatC.Col = SCC_Depot) or (G_CatC.col = SCC_Datedeb) or
                             (G_CatC.col = SCC_Datefin) or (G_CatC.Col = SCC_Cat)) ;
    StCellCur := G_CatC.Cells [G_CatC.Col, G_CatC.Row] ;
    AffectMenuCondApplic (G_CatC, ttdCatC);
    END ;
end;

procedure TFTarifCatArt.G_CatCCellExit(Sender: TObject; var ACol,
                                       ARow: Integer; var Cancel: Boolean);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
FormateZoneSaisie (ACol,ARow, ttdCatC);
if ACol = SCC_Depot then TraiterDepot (ACol, ARow, ttdCatC) else
    if ACol = SCC_Lib then TraiterLibelle (ACol, ARow, ttdCatC) else
        if ACol = SCC_Cat then TraiterCatTiers (ACol, ARow, ttdCatC) else
            if ACol = SCC_Px then TraiterPrix (ACol, ARow, ttdCatC) else
                if ACol = SCC_Rem then TraiterRemise (ACol, ARow, ttdCatC) else
                    if ACol = SCC_Datedeb then TraiterDateDeb (ACol, ARow, ttdCatC) else
                        if ACol = SCC_Datefin then TraiterDateFin (ACol, ARow, ttdCatC);
if Not Cancel then
    BEGIN
    END;
end;

procedure TFTarifCatArt.G_CatCElipsisClick(Sender: TObject);
Var DEPOT, DATE, CAT : THCritMaskEdit;
    Coord : TRect;
begin
if G_CatC.Col = SCC_Depot then
    BEGIN
    Coord := G_CatC.CellRect (G_CatC.Col, G_CatC.Row);
    DEPOT := THCritMaskEdit.Create (Self);
    DEPOT.Parent := G_CatC;
    DEPOT.Top := Coord.Top;
    DEPOT.Left := Coord.Left;
    DEPOT.Width := 3; DEPOT.Visible := False;
    DEPOT.DataType := 'GCDEPOT';
    GetDepotRecherche (DEPOT) ;
    if DEPOT.Text <> '' then G_CatC.Cells[G_CatC.Col,G_CatC.Row]:= DEPOT.Text;
    DEPOT.Destroy;
    END ;
if G_CatC.Col = SCC_Cat then
    BEGIN
    Coord := G_CatC.CellRect (G_CatC.Col, G_CatC.Row);
    CAT := THCritMaskEdit.Create (Self);
    CAT.Parent := G_CatC;
    CAT.Top := Coord.Top;
    CAT.Left := Coord.Left;
    CAT.Width := 3; CAT.Visible := False;
    CAT.DataType := 'TTTARIFCLIENT';
    GetCategorieRecherche (CAT) ;
    if CAT.Text <> '' then G_CatC.Cells[G_CatC.Col,G_CatC.Row]:= CAT.Text;
    CAT.Destroy;
    END ;
if (G_CatC.Col = SCC_Datedeb) or (G_CatC.Col = SCC_Datefin) then
    BEGIN
    Coord := G_CatC.CellRect (G_CatC.Col, G_CatC.Row);
    DATE := THCritMaskEdit.Create (Self);
    DATE.Parent := G_CatC;
    DATE.Top := Coord.Top;
    DATE.left := Coord.Left;
    DATE.Width := 3; DATE.Visible := False;
    DATE.OpeType:=otDate;
    GetDateRecherche (TForm(DATE.Owner), DATE) ;
    if DATE.Text <> '' then G_CatC.Cells[G_CatC.Col,G_CatC.Row]:= DATE.Text;
    DATE.Destroy;
    END;
end;

{==============================================================================================}
{========================= Manipulation des LIGNES Quantitatif ================================}
{==============================================================================================}
Procedure TFTarifCatArt.InitLaLigne (ARow : integer;  ttd : T_TableTarif);
Var TOBL : TOB;
BEGIN
TOBL:=GetTOBLigne(ARow, ttd) ; if TOBL=Nil then Exit ;
TOBL.PutValue ('GF_ARTICLE', '');
TOBL.PutValue ('GF_TARIFARTICLE', CodeTarifArticle);
TOBL.PutValue ('GF_BORNEINF', -999999);
TOBL.PutValue ('GF_QUANTITATIF', '-');
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
TOBL.PutValue ('GF_REGIMEPRIX', 'GLO');
AfficheLAligne (ARow, ttd);
END;

Function TFTarifCatArt.GetTOBLigne ( ARow : integer;  ttd : T_TableTarif) : TOB ;
BEGIN
Result:=Nil ;
case ttd of
    ttdCatA : BEGIN
              if ((ARow<=0) or (ARow>TOBTarfCatA.Detail.Count)) then Exit ;
              Result:=TOBTarfCatA.Detail[ARow-1] ;
              END;
    ttdCatC : BEGIN
              if ((ARow<=0) or (ARow>TOBTarfCatC.Detail.Count)) then Exit ;
              Result:=TOBTarfCatC.Detail[ARow-1] ;
              END;
    END;
END ;

procedure TFTarifCatArt.InitialiseGrille ;
BEGIN
G_CatA.VidePile(True) ;
G_CatA.RowCount:= NbRowsInit ;
G_CatC.VidePile(True) ;
G_CatC.RowCount:= NbRowsInit ;
END;

procedure TFTarifCatArt.InitialiseLigne (ARow : integer; ttd : T_TableTarif) ;
Var TOBL : TOB ;
    i_ind : integer ;
BEGIN
TOBL:=GetTOBLigne(ARow, ttd) ; if TOBL<>Nil then TOBL.InitValeurs ;
case ttd of
    ttdCatA : BEGIN
              for i_ind := 1 to G_CatA.ColCount-1 do
                  BEGIN
                  G_CatA.Cells [i_ind, ARow]:='' ;
                  END;
              END;
    ttdCatC : BEGIN
              for i_ind := 1 to G_CatC.ColCount-1 do
                  BEGIN
                  G_CatC.Cells [i_ind, ARow]:='' ;
                  END;
              END;
    END;
END;

procedure TFTarifCatArt.AfficheLaLigne (ARow : integer;  ttd : T_TableTarif) ;
Var TOBL : TOB ;
    i_ind : integer ;
BEGIN
TOBL:=GetTOBLigne(ARow, ttd) ; if TOBL = Nil then exit;
Case ttd of
    ttdCatA : BEGIN
              for i_ind := 1 to G_CatA.ColCount-1 do
                  BEGIN
                  if ((i_ind=SCA_Datedeb) or (i_ind=SCA_Datefin)) then
                      G_CatA.ColFormats [i_ind] := '';
                  END;
              TOBL.PutLigneGrid(G_CatA,ARow,False,False,LesColCatA) ;
              if G_CatA.Cells [SCA_Datedeb, ARow] = '' then
                  G_CatA.Cells [SCA_Datedeb, ARow] := DateToStr (IDate1900);
              for i_ind := 1 to G_CatA.ColCount-1 do
                  BEGIN
                  FormateZoneSaisie(i_ind, ARow, ttd) ;
                  END;
              END;
    ttdCatC : BEGIN
              for i_ind := 1 to G_CatC.ColCount-1 do
                  BEGIN
                  if ((i_ind=SCC_Datedeb) or (i_ind=SCC_Datefin)) then
                      G_CatC.ColFormats [i_ind] := '';
                  END;
              TOBL.PutLigneGrid(G_CatC,ARow,False,False,LesColCatC) ;
              if G_CatC.Cells [SCC_Datedeb, ARow] = '' then
                  G_CatC.Cells [SCC_Datedeb, ARow] := DateToStr (IDate1900);
              for i_ind := 1 to G_CatC.ColCount-1 do
                  BEGIN
                  FormateZoneSaisie(i_ind, ARow, ttd) ;
                  END;
              END;
    END;
END ;

Procedure TFTarifCatArt.DepileTOBLigne ;
var i_ind : integer;
BEGIN
for i_ind := TOBTarfCatA.Detail.Count - 1 Downto 0 do
    BEGIN
    TOBTarfCatA.Detail[i_ind].Free ;
    END;
for i_ind := TOBTarfCatC.Detail.Count - 1 Downto 0 do
    BEGIN
    TOBTarfCatC.Detail[i_ind].Free ;
    END;
for i_ind := TOBTarifDel.Detail.Count - 1 Downto 0 do
    BEGIN
    TOBTarifDel.Detail[i_ind].Free ;
    END;
END;

Procedure TFTarifCatArt.CreerTOBLigne (ARow : integer; ttd : T_TableTarif);
BEGIN
Case ttd of
    ttdCatA : BEGIN
              if ARow <> TOBTarfCatA.Detail.Count + 1 then exit;
              TOB.Create ('TARIF', TOBTarfCatA, ARow-1) ;
              InitialiseLigne (ARow, ttd) ;
              END;
    ttdCatC : BEGIN
              if ARow <> TOBTarfCatC.Detail.Count + 1 then exit;
              TOB.Create ('TARIF', TOBTarfCatC, ARow-1) ;
              InitialiseLigne (ARow, ttd) ;
              END;
    END ;
END;

Function TFTarifCatArt.LigneVide (ARow : integer; ttd : T_TableTarif) : Boolean;
BEGIN
Result := True;
Case ttd of
    ttdCatA : BEGIN
              if G_CatA.Cells [SCA_Lib, ARow] <> '' then
                  BEGIN
                  Result := False;
                  Exit;
                  END;
              END;
    ttdCatC : BEGIN
              if (G_CatC.Cells [SCC_Lib, ARow] <> '') And (G_CatC.Cells [SCC_Cat, ARow] <> '') then
                  BEGIN
                  Result := False;
                  Exit;
                  END;
              END;
    END ;
END;

Procedure TFTarifCatArt.PreAffecteLigne (ARow : integer; ttd : T_TableTarif);
BEGIN
;
END;

{==============================================================================================}
{===================== Manipulation des Champs LIGNES Quantitatif =============================}
{==============================================================================================}
procedure TFTarifCatArt.TraiterDepot (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    St : string;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
Case ttd of
    ttdCatA : St := G_CatA.Cells [ACol, ARow];
    ttdCatC : St := G_CatC.Cells [ACol, ARow];
    END ;
if ExisteDepot ('GCDEPOT', St) then
    BEGIN
    TOBL.PutValue ('GF_DEPOT', St);
    END else
    BEGIN
    // message dépôt inexistant
    MsgBox.Execute (4,Caption,'') ;
    Case ttd of
        ttdCatA : G_CatA.Cells [ACol, ARow] := TOBL.GetValue ('GF_DEPOT');
        ttdCatC : G_CatC.Cells [ACol, ARow] := TOBL.GetValue ('GF_DEPOT');
        END ;
    END;
END;

procedure TFTarifCatArt.TraiterCatTiers (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    B_NewLine : Boolean;
    St : string;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
if TOBL.GetValue ('GF_TARIFTIERS') = '' then B_NewLine :=True else B_NewLine := False;
St := G_CatC.Cells [ACol, ARow];
if St <> '' then
    BEGIN
    if ExisteCategorie ('TTTARIFCLIENT', St) then
        BEGIN
        TOBL.PutValue ('GF_TARIFTIERS', St);
        END else
        BEGIN
        // message Catégorie Tiers inexistante
        MsgBox.Execute (5,Caption,'') ;
        G_CatC.Cells [ACol, ARow] := TOBL.GetValue ('GF_TARIFTIERS');
        END;
    if (TOBL.GetValue ('GF_LIBELLE') <> '') AND (TOBL.GetValue ('GF_TARIFTIERS') <> '') AND
       (B_NewLine) then
        BEGIN
        InitLaLigne (ARow, ttd);
        END;
    END else if Not B_Newline then
            G_CatC.Cells [ACol, ARow] := TOBL.GetValue ('GF_TARIFTIERS');
END;

procedure TFTarifCatArt.TraiterLibelle (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    B_NewLine : Boolean;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
if TOBL.GetValue ('GF_LIBELLE') = '' then B_NewLine :=True else B_NewLine := False;
Case ttd of
    ttdCatA : BEGIN
              if G_CatA.Cells [ACol, ARow] <> '' then
                  BEGIN
                  TOBL.PutValue ('GF_LIBELLE', G_CatA.Cells [ACol, ARow]);
                  if (TOBL.GetValue ('GF_LIBELLE') <> '') AND (B_NewLine) then
                      BEGIN
                      InitLaLigne (ARow, ttd);
                      END;
                  END else if Not B_Newline then
                      G_CatA.Cells [ACol, ARow] := TOBL.GetValue ('GF_LIBELLE');
              END;
    ttdCatC : BEGIN
              if G_CatC.Cells [ACol, ARow] <> '' then
                  BEGIN
                  TOBL.PutValue ('GF_LIBELLE', G_CatC.Cells [ACol, ARow]);
                  if (TOBL.GetValue ('GF_LIBELLE') <> '') AND
                     (TOBL.GetValue ('GF_TARIFTIERS') <> '') AND
                     (B_NewLine) then
                      BEGIN
                      InitLaLigne (ARow, ttd);
                      END;
                  END else if Not B_Newline then
                      G_CatC.Cells [ACol, ARow] := TOBL.GetValue ('GF_LIBELLE');
              END;
    END ;
END;

procedure TFTarifCatArt.TraiterPrix (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
Case ttd of
    ttdCatA : TOBL.PutValue ('GF_PRIXUNITAIRE', Valeur (G_CatA.Cells [ACol, ARow]));
    ttdCatC : TOBL.PutValue ('GF_PRIXUNITAIRE', Valeur (G_CatC.Cells [ACol, ARow]));
    END ;
END;

procedure TFTarifCatArt.TraiterRemise(ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    St : string;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
Case ttd of
    ttdCatA : begin
              G_CatA.Cells [ACol, ARow]:=ModifFormat(G_CatA.Cells [ACol, ARow]);
              St := G_CatA.Cells [ACol, ARow];
              end;
    ttdCatC : begin
              G_CatC.Cells [ACol, ARow]:= ModifFormat(G_CatC.Cells [ACol, ARow]);
              St := G_CatC.Cells [ACol, ARow];
              end;
    END ;
TOBL.PutValue ('GF_CALCULREMISE', St);
TOBL.PutValue ('GF_REMISE', RemiseResultante (St));
AffichePied (ttd);
END;

procedure TFTarifCatArt.TraiterDateDeb (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    St_Date : string;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
Case ttd of
    ttdCatA : St_Date := G_CatA.Cells [ACol, ARow] ;
    ttdCatC : St_Date := G_CatC.Cells [ACol, ARow] ;
    END ;
if IsValidDate (st_Date) then
    BEGIN
    if StrToDate (St_Date) > TOBL.GetValue ('GF_DATEFIN') then
        BEGIN
            MsgBox.Execute (3,Caption,'') ;
            Case ttd of
                ttdCatA : BEGIN
                          G_CatA.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
                          G_CatA.Col := ACol; G_CatA.Row := ARow;
                          END;
                ttdCatC : BEGIN
                          G_CatC.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
                          G_CatC.Col := ACol; G_CatC.Row := ARow;
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
            ttdCatA : G_CatA.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
            ttdCatC : G_CatC.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
            END ;
        END;
    END;
END;

procedure TFTarifCatArt.TraiterDateFin (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    St_Date : string;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
Case ttd of
    ttdCatA : St_Date := G_CatA.Cells [ACol, ARow] ;
    ttdCatC : St_Date := G_CatC.Cells [ACol, ARow] ;
    END ;
if IsValidDate (st_Date) then
    BEGIN
    if StrToDate (St_Date) < TOBL.GetValue ('GF_DATEDEBUT') then
        BEGIN
            MsgBox.Execute (3,Caption,'') ;
            Case ttd of
                ttdCatA : BEGIN
                          G_CatA.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
                          G_CatA.Col := ACol; G_CatA.Row := ARow;
                          END;
                ttdCatC : BEGIN
                          G_CatC.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
                          G_CatC.Col := ACol; G_CatC.Row := ARow;
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
            ttdCatA : G_CatA.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
            ttdCatC : G_CatC.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
            END ;
        END;
    END;
END;

{==============================================================================================}
{========================= Evenement de l'Entete  =============================================}
{==============================================================================================}
procedure TFTarifCatArt.GF_TARIFARTICLEChange(Sender: TObject);
var i_Rep : integer;
    ioerr : TIOErr ;
begin
if CodeTarifArticle = GF_TARIFARTICLE.Value then Exit;
if GF_TARIFARTICLE.Value <> '' then
    BEGIN
    G_CatA.Enabled := True ; G_CatC.Enabled := True ;
    END;
i_Rep := QuestionTarifEnCours;
Case i_Rep of
    mrYes    : BEGIN
               ioerr := Transactions (ValideTarif, 2);;
               Case ioerr of
                    oeOk      : ;
                    oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
                    oeSaisie  : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
                    END ;
               CodeTarifArticle := GF_TARIFARTICLE.Value;
               Transactions (ChargeTarif, 1) ;
               END;
    mrNo     : BEGIN
               InitialiseGrille;
               DepileTOBLigne;
               CodeTarifArticle := GF_TARIFARTICLE.Value;
               Transactions (ChargeTarif, 1) ;
               END;
    mrCancel : GF_TARIFARTICLE.Value := CodeTarifArticle;
    END ;
end;

procedure TFTarifCatArt.GF_DEVISEChange(Sender: TObject);
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
               Transactions (ChargeTarif, 1) ;
               END;
    mrNo     : BEGIN
               InitialiseGrille;
               DepileTOBLigne;
               CodeDevise := GF_DEVISE.Value;
               Transactions (ChargeTarif, 1) ;
               END;
    mrCancel : GF_DEVISE.Value := CodeDevise;
    END ;
DEV.Code := GF_DEVISE.Value ; GetInfosDevise (DEV) ;
end;

procedure TFTarifCatArt.CBCATCLIClick(Sender: TObject);
begin
if CBCATCLI.Checked then
    BEGIN
    PTARIF.Visible := False;
    PCATEGORIE.Visible := True;
    PTITRE.Caption := HTitre.Mess[1] ;
    AffectMenuCondApplic (G_CatC, ttdCatC);
    END else
    BEGIN
    PTARIF.Visible := True;
    PCATEGORIE.Visible := False;
    PTITRE.Caption := HTitre.Mess[0] ;
    AffectMenuCondApplic (G_CatA, ttdCatA);
    END;
HMTrad.ResizeGridColumns (G_CatA) ;
HMTrad.ResizeGridColumns (G_CatC) ;
end;

{==============================================================================================}
{========================= Manipulation de l'entête ===========================================}
{==============================================================================================}
Procedure TFTarifCatArt.InitialiseEntete ;
BEGIN
GF_DEVISE.Value := CodeDevise;
CodeTarifArticle := '';
AffecteEntete;
GF_TARIFARTICLE.SetFocus ;
G_CatA.Enabled := False ; G_CatC.Enabled := False ;
BCondAplli.Enabled := False ; BCopierCond.Enabled := False; BCollerCond.Enabled := False ;
CONDAPPLIC.Text := '' ;
END;

Procedure TFTarifCatArt.AffecteEntete ;
BEGIN
GF_TARIFARTICLE.Value := CodeTarifArticle;
END;

Function TFTarifCatArt.QuestionTarifEnCours : Integer;
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

procedure TFTarifCatArt.GF_REMISEChange(Sender: TObject);
var St : string;
begin
St := GF_REMISE.Text;
St:=StrF00(Valeur(St),ADecimP);
GF_REMISE.Text := St;
end;

procedure TFTarifCatArt.GF_CASCADEREMISEChange(Sender: TObject);
Var Row : integer;
    ttd : T_TableTarif;
    TOBL : TOB ;
begin
if PTARIF.Visible=True then
    BEGIN
    Row := G_CatA.Row;
    ttd := ttdCatA;
    END else if PCATEGORIE.Visible=True then
    BEGIN
    Row := G_CatC.Row;
    ttd := ttdCatC;
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
Procedure TFTarifCatArt.AffichePied (ttd : T_TableTarif) ;
var TOBL : TOB;
    FF   : TForm;
    ARow : Longint ;
BEGIN
if ttd = ttdCatC then  ARow := G_CatC.Row;
if ttd = ttdCatA then  ARow := G_CatA.Row;
TOBL:=GetTOBLigne(ARow, ttd) ; if TOBL=Nil then Exit ;
FF := TForm (PPIED.Owner);
TOBL.PutEcran (FF, PPIED);
END;

{==============================================================================================}
{============================ Evenement lié aux Boutons =======================================}
{==============================================================================================}

procedure TFTarifCatArt.BChercherClick(Sender: TObject);
begin
if PTARIF.Visible then
    if G_CatA.RowCount < 3 then Exit;
if PCATEGORIE.Visible then
    if G_CatC.RowCount < 3 then Exit;
FindDebut:=True ; FindLigne.Execute ;
end;

procedure TFTarifCatArt.BVoirCondClick(Sender: TObject);
Var TOBL : TOB;
begin
TCONDTARF.Visible := BVoirCond.Down ;
end;

procedure TFTarifCatArt.BCondAplliClick(Sender: TObject);
Var TOBL : TOB;
begin
{$IFNDEF CCS3}
if GF_TARIFARTICLE.Value = '' then exit;
if PTARIF.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_CatA.Row, ttdCatA) ; if TOBL=Nil then Exit ;
    EntreeTarifCond (Action, TOBL);
    AfficheCondTarf (G_CatA.Row, ttdCatA);
    END;
if PCATEGORIE.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_CatC.Row, ttdCatC) ; if TOBL=Nil then Exit ;
    EntreeTarifCond (Action, TOBL);
    AfficheCondTarf (G_CatC.Row, ttdCatC);
    END;
{$ENDIF}
end;

procedure TFTarifCatArt.BCopierCondClick(Sender: TObject);
Var TOBL : TOB;
begin
if GF_TARIFARTICLE.Value = '' then exit;
if PTARIF.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_CatA.Row, ttdCatA) ; if TOBL=Nil then Exit ;
    CONDAPPLIC.Text := TOBL.GetValue ('GF_CONDAPPLIC') ;
    END;
if PCATEGORIE.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_CatC.Row, ttdCatC) ; if TOBL=Nil then Exit ;
    CONDAPPLIC.Text := TOBL.GetValue ('GF_CONDAPPLIC') ;
    END;
end;

procedure TFTarifCatArt.BCollerCondClick(Sender: TObject);
Var TOBL : TOB;
begin
if GF_TARIFARTICLE.Value = '' then exit;
if PTARIF.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_CatA.Row, ttdCatA) ; if TOBL=Nil then Exit ;
    TOBL.PutValue ('GF_CONDAPPLIC', CONDAPPLIC.Text) ;
    AfficheCondTarf (G_CatA.Row, ttdCatA);
    END;
if PCATEGORIE.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_CatC.Row, ttdCatC) ; if TOBL=Nil then Exit ;
    TOBL.PutValue ('GF_CONDAPPLIC', CONDAPPLIC.Text) ;
    AfficheCondTarf (G_CatC.Row, ttdCatC);
    END;
end;

{==============================================================================================}
{============================== Action lié aux Boutons ========================================}
{==============================================================================================}
procedure TFTarifCatArt.FindLigneFind(Sender: TObject);
begin
if PTARIF.Visible then
    Rechercher (G_CatA, FindLigne, FindDebut) ;
if PCATEGORIE.Visible then
    Rechercher (G_CatC, FindLigne, FindDebut) ;
end;

{==============================================================================================}
{================================= Validation =================================================}
{==============================================================================================}
procedure TFTarifCatArt.BValiderClick(Sender: TObject);
Var ioerr : TIOErr ;
begin
// validation
ioerr := Transactions (ValideTarif, 2);
Case ioerr of
        oeOk  : ;
    oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
    oeSaisie  : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
   END ;
InitialiseEntete;
end;

procedure TFTarifCatArt.ValideTarif;
begin
if Not SortDeLaLigne(ttdCatA) then Exit ;
if Not SortDeLaLigne(ttdCatC) then Exit ;
TOBTarifDel.DeleteDB (False);
VerifLesTOB;
TOBTarif.InsertOrUpdateDB(False) ;
InitialiseGrille;
DepileTOBLigne;
end;

procedure TFTarifCatArt.VerifLesTOB;
var i_ind : integer;
    Q : TQuery ;
    MaxTarif : Longint ;
BEGIN
Q := OpenSQL ('SELECT MAX(GF_TARIF) FROM TARIF', TRUE,-1,'',true) ;
if Q.EOF then MaxTarif := 1 else MaxTarif := Q.Fields[0].AsInteger + 1 ;
Ferme(Q) ;
for i_ind := TOBTarfCatA.Detail.count - 1 Downto 0 do
    BEGIN
    if LigneVide (i_ind + 1, ttdCatA) then
        BEGIN
        TOBTarfCatA.Detail[i_ind].Free ;
        END else
        BEGIN
        if TOBTarfCatA.Detail[i_ind].GetValue ('GF_TARIF') = 0 then
            BEGIN
            TOBTarfCatA.Detail[i_ind].PutValue ('GF_TARIF', MaxTarif);
            Inc (MaxTarif);
            END;
        CalcPriorite (TOBTarfCatA.Detail[i_ind]);
        END;
    END;
for i_ind := TOBTarfCatC.Detail.count - 1 Downto 0 do
    BEGIN
    if LigneVide (i_ind + 1, ttdCatC) then
        BEGIN
        TOBTarfCatC.Detail[i_ind].Free ;
        END else
        BEGIN
        if TOBTarfCatC.Detail[i_ind].GetValue ('GF_TARIF') = 0 then
            BEGIN
            TOBTarfCatC.Detail[i_ind].PutValue ('GF_TARIF', MaxTarif);
            Inc (MaxTarif);
            END;
        CalcPriorite (TOBTarfCatC.Detail[i_ind]);
        END;
    END;
END;

{==============================================================================================}
{=============================== Conditions tarifaires ========================================}
{==============================================================================================}
procedure TFTarifCatArt.InitComboChamps ;
begin
SourisSablier ;
RemplitComboChamps('TIERS',FComboTIE) ;
RemplitComboChamps('ARTICLE',FComboART) ;
RemplitComboChamps('LIGNE',FComboLIG) ;
RemplitComboChamps('PIECE',FComboPIE) ;
SourisNormale ;
end ;

procedure TFTarifCatArt.RemplitComboChamps(NomTable : String ; FCombo : THValComboBox) ;
begin
ExtractFields(NomTable,'T',FCombo.Items,FCombo.Values,Nil,False);
end ;

procedure TFTarifCatArt.EffaceGrid ;
var lig, col : Integer ;
begin
for lig := 1 to G_COND.RowCount - 1 do
   begin
   for col := 0 to G_COND.ColCount - 1 do G_COND.Cells[col, lig] := '' ;
   end ;
G_COND.Row := 1 ;
end ;

Procedure TFTarifCatArt.AfficheCondTarf (ARow : Longint; ttd : T_TableTarif) ;
var TOBL : TOB;
    FF   : TForm;
BEGIN
TOBL:=GetTOBLigne(ARow, ttd) ; if TOBL=Nil then Exit ;
GetConditions (TOBL) ;
END;

function TFTarifCatArt.ValueToItem(CC : THValComboBox ; St : String) : String ;
var i_ind : Integer ;
begin
i_ind := CC.Values.IndexOf(St) ;
if i_ind >= 0 then result := CC.Items[i_ind] else result := '' ;
end ;

{Charge dans le grid les conditions stockées dans le TMemoField}
procedure TFTarifCatArt.GetConditions (TOBL : TOB) ;
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

procedure TFTarifCatArt.TCONDTARFClose(Sender: TObject);
begin
BVoirCond.Down := False ;
end;

procedure TFTarifCatArt.BAbandonClick(Sender: TObject);
begin
Close ;
if FClosing and IsInside(Self) then THPanel(parent).CloseInside ;
end;

Procedure AGLEntreeTarifCatArt ( Parms : array of variant ; nb : integer ) ;
Var Action : TActionFiche ;
BEGIN
Action:=StringToAction(String(Parms[1])) ;
EntreeTarifCatArt(Action) ;
END ;

procedure InitTarifCatArt();
begin
RegisterAglProc('EntreeTarifCatArt',True,1,AGLEntreeTarifCatArt) ;
end;

procedure TFTarifCatArt.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self);
end;

Initialization
InitTarifCatArt();

end.

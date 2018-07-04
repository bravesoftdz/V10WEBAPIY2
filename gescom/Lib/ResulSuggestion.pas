unit ResulSuggestion;

interface

uses Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
     StdCtrls, Hctrls, Mask, ExtCtrls, Grids, hmsgbox, ColMemo,AglInit, UtilGC,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$IFDEF V530}
      EdtEtat,EdtDoc,
{$ELSE}
      EdtREtat, EdtRDoc,
{$ENDIF}
{$ENDIF}
{$IFDEF BTP}
      FactOuvrage,
      UtilSuggestion,
      UPlannifChUtil,
      UtilArticle,
{$ENDIF}
     UTofListeInv,
     UIUtil, UTOB, HTB97, M3FP, CalculSuggestion, TeEngine, TeeFunci, Series,
     TeeProcs, Chart, Math, ComCtrls, HEnt1, EntGC, FactUtil, FactComm,Ent1, FactPiece, FactTOB,
     ParamSoc, SaisUtil, HPanel, HSysMenu, NomenUtil, Menus, TiersUtil,uEntCommun,UtilTOBPiece,
  TntStdCtrls, TntGrids;

procedure Entree_ResulSuggestion(Parms: array of variant; nb: integer);

type
  TFResulSuggestion = class(TForm)
    G_LIG: THGrid;
    PANTETE: TPanel;
    MsgBox: THMsgBox;
    Recap: TColorMemo;
    DetailEvol: TToolWindow97;
    G_DET: THGrid;
    DataLigne: TMemo;
    GZZ_CREATION: THCritMaskEdit;
    TSR_CREATION: THLabel;
    G_Mvts: THGrid;
    G_Evol: THGrid;
    G_Fin: THGrid;
    DetailFourni: TToolWindow97;
    G_FOU: THGrid;
    HMTrad: THSystemMenu;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    bDetail: TToolbarButton97;
    bTri: TToolbarButton97;
    bRecalcul: TToolbarButton97;
    bGenPiece: TToolbarButton97;
    bPrint: TToolbarButton97;
    BAide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BValider: TToolbarButton97;
    GZZ_SUGGEST: THCritMaskEdit;
    BMenuZoom: TToolbarButton97;
    BZoomArticle: TToolbarButton97;
    BZoomTiers: TToolbarButton97;
    POPZ: TPopupMenu;
    // MODIF BTP
    BDelete: TToolbarButton97;
    HPiece: THMsgBox;
    HTitres: THMsgBox;
    BZoomLigne: TToolbarButton97;
    Panel1: TPanel;
    LStockMin: TLabel;
    LStockMax: TLabel;
    PChoixFour: TPopupMenu;
    Catalogue: TMenuItem;
    Listefournisseurs: TMenuItem;
    BChercher: TToolbarButton97;
    FindLigne: TFindDialog;
    Pbas: TPanel;
    PANDIM: TPanel;
    TGA_CODEDIM5: TLabel;
    TGA_CODEDIM4: TLabel;
    TGA_CODEDIM3: TLabel;
    TGA_CODEDIM2: TLabel;
    TGA_CODEDIM1: TLabel;
    Ptotaux: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    TmontantBudget: THNumEdit;
    TmontantAch: THNumEdit;
    //--------
    procedure FormShow(Sender: TObject);
    procedure G_LIGRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GZZ_SUGGESTExit(Sender: TObject);
    procedure DataLigneChange(Sender: TObject);
    procedure bDetailClick(Sender: TObject);
    procedure bPrintClick(Sender: TObject);
    procedure G_LIGCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure BValiderClick(Sender: TObject);
    procedure G_LIGCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_LIGElipsisClick(Sender: TObject);
    procedure bTriClick(Sender: TObject);
    procedure G_FOUSorted(Sender: TObject);
    procedure bRecalculClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure G_FOUDblClick(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure BZoomArticleClick(Sender: TObject);
    procedure BZoomTiersClick(Sender: TObject);
    // MODIF BTP
    procedure BDeleteClick(Sender: TObject);
    procedure BZoomLigneClick(Sender: TObject);
    procedure ListefournisseursClick(Sender: TObject);
    procedure CatalogueClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure FindLigneFind(Sender: TObject);
    //----------
//    procedure bCourbesClick(Sender: TObject);
  private
    { Déclarations privées }
//    TobSuggest : TOB;
    LesCols : string;
    stcell : string;
//    sPrefixe : string;
    TousDepot : boolean;
    DateDebut : string;
    Prio1, Prio2, Prio3, Prio4 : string;
    cledoc : R_CLEDOC; // Ajout BTP
    FindFirst : boolean;
    procedure EnregistrePiece;
    // Modif BTP
    function TraiteDonneeEntree : boolean;
    procedure loadlestobs;
    procedure DetruitLaPiece;
    procedure AjouteEvent(NatureP, MessEvent: string);
    procedure MiseEnPlaceRefFournisseur;
    procedure ModifiePiece;
    procedure AjouteUniteAchat (TOBL : TOB;ModeCreation : boolean);
    procedure TraiteChangeFour(TOBL: TOB; Arow : integer);
    procedure PositionneEnAchat(TOBPiece: TOB; ModeCreation : boolean=false);
    procedure CumuleMontants (TOBL : TOB);
		procedure AjouteMontantBudget(TOBL : TOB);
		procedure AjouteMontantAchat(TOBL: TOB; Sens : string='+');

{$IFDEF BTP}
    function ControleQteBesoin (TOBL : TOB ; Quantite : double) : boolean;
{$ENDIF}
    // --

  public
    { Déclarations publiques }
    Action      : TActionFiche ;
  end;

procedure Entree2_ResulSuggestion;
procedure ModifSuggestion (Cledoc : R_CLEDOC ; Action : TActionFiche );

var
  TobPiece,TOBLien : TOB;
  CleDoc      : R_CleDoc ;
  FResulSuggestion: TFResulSuggestion;

implementation
uses UtilBTPgestChantier;
{$R *.DFM}

procedure Entree_ResulSuggestion(Parms: array of variant; nb: integer);
var
    st1, st2 : string;

begin
  st1 := string(Parms[0]);
  st2 := string(Parms[1]);
  if Pos(';', st1) = 0 then Exit;
  FResulSuggestion := TFResulSuggestion.Create(Application);
  TRY
    FResulSuggestion.bPrint.Visible := False;
    TobPiece.PutValue('GP_BLOCNOTE', st2);
    if Calcul_Suggestion(st1, TobPiece, TOBLien) then
    begin
      FResulSuggestion.MiseEnPlaceRefFournisseur;
			FResulSuggestion.PositionneEnAchat (TOBPiece,True);
    	FResulSuggestion.ShowModal;
    end;
  FINALLY
    FResulSuggestion.free;
  END;
end;

procedure Entree2_ResulSuggestion;
var
    FRS : TFResulSuggestion;
    PP  : THPanel ;

begin
FRS := TFResulSuggestion.Create(Application);
FRS.Action := taModif;
SourisSablier;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
      FRS.ShowModal ;
    finally
      FRS.Free ;
    end ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(FRS,PP) ;
   FRS.Show ;
   END ;
end;

procedure ModifSuggestion (Cledoc : R_CLEDOC ; Action : TActionFiche );
var
    FRS : TFResulSuggestion;
    PP  : THPanel ;

begin
FRS := TFResulSuggestion.Create(Application);
FRS.Action := Action;
FRS.Cledoc := Cledoc;
SourisSablier;
TRY
  FRS.TraiteDonneeEntree;
FINALLY
  SourisNormale ;
END;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
      FRS.ShowModal ;
    finally
      FRS.Free ;
    end ;
   END else
   BEGIN
   InitInside(FRS,PP) ;
   FRS.Show ;
   END ;
end;

procedure TFResulSuggestion.FormShow(Sender: TObject);
var i_ind1 : integer;
    Col7 : boolean;
    FF, FQ {, FFPU} : string;
begin
//sPrefixe := TableToPrefixe('GCTMPREAPPRO');
//TobSuggest := TOB.Create('GCTMPREAPPRO', nil, -1);
TGA_CODEDIM1.Caption := '';
TGA_CODEDIM2.Caption := '';
TGA_CODEDIM3.Caption := '';
TGA_CODEDIM4.Caption := '';
TGA_CODEDIM5.Caption := '';
FF := '';
for i_ind1 := 1 to GetParamSoc('SO_DECPRIX') do FF := FF + '0';
FQ := '';
for i_ind1 := 1 to GetParamSoc('SO_DECQTE') do FQ := FQ + '0';

Col7 := False;

TmontantBudget.Value  := 0; TmontantAch.Value := 0;
G_LIG.RowCount := 50;
G_LIG.ColCount := 9;
//PositionneEnAchat (TOBPiece);
for i_ind1 := 0 to TobPiece.Detail.Count - 1 do
begin
  CumuleMontants(TOBPiece.detail[i_ind1]);
  if TobPiece.Detail[i_ind1].FieldExists('STOCKMAX') then Col7 := True;
end;
LesCols := 'GL_CODEARTICLE;GL_LIBELLE;GL_DEPOT;QTEACHAT;GL_QUALIFQTEACH;PUACHAT;GL_PUHTBASE;GL_TIERS;LIBTIERS';
G_LIG.Cells[0, 0] := 'Article';
G_LIG.ColWidths[0] := 100;
G_LIG.ColLengths[0] := -1;

G_LIG.Cells[1, 0] := 'Libellé';
G_LIG.ColWidths[1] := 280;
G_LIG.ColLengths[1] := -1;

G_LIG.Cells[2, 0] := 'Dépôt';
G_LIG.ColWidths[2] := 30;
G_LIG.ColLengths[2] := -1;

G_LIG.Cells[3, 0] := 'Quantité';
G_LIG.ColWidths[3] := 50;
G_LIG.ColAligns[3] := taRightJustify;
G_LIG.ColFormats[3] := '#,##0.' + FQ;
//
G_LIG.Cells[4, 0] := 'Unité';
G_LIG.ColWidths[4] := 30;
G_LIG.ColLengths[4] := -1;
G_LIG.ColAligns[4] := taCenter;
//
G_LIG.Cells[5, 0] := 'PU HT';
G_LIG.ColWidths[5] := 50;
G_LIG.ColAligns[5] := taRightJustify;
G_LIG.ColFormats[5] := '#,##0.' + FF;
//
G_LIG.Cells[6, 0] := 'Prix Budget';
G_LIG.ColWidths[6] := 50;
G_LIG.ColLengths[6] := -1;
G_LIG.ColAligns[6] := taRightJustify;
G_LIG.ColFormats[6] := '#,##0.' + FF;
//
G_LIG.Cells[7, 0] := 'Fournisseur';
G_LIG.ColWidths[7] := 50;

G_LIG.Cells[8, 0] := 'Libellé';
G_LIG.ColLengths[8] := -1;
G_LIG.ColWidths[8] := 100;

G_DET.ColAligns[2] := taRightJustify;
G_DET.ColFormats[2] := '#,##0.' + FQ;
G_DET.ColAligns[3] := taRightJustify;
G_DET.ColFormats[3] := '#,##0.' + FQ;
G_DET.ColAligns[4] := taRightJustify;
G_DET.ColFormats[4] := '#,##0.' + FQ;
G_DET.ColAligns[5] := taRightJustify;
G_DET.ColFormats[5] := '#,##0.' + FQ;
G_FOU.Cells[2, 0] := RechDom('GCAPPELPRIX', GetInfoParPiece('REA', 'GPP_APPELPRIX'), False);
G_FOU.ColAligns[2] := taRightJustify;
G_FOU.ColFormats[2] := '#,##0.' + FF;
G_Mvts.Visible:=False;
G_Evol.Visible:=False;
G_Fin.Visible:=False;
HMTrad.ResizeGridColumns (G_LIG) ;
if TobPiece.Detail.Count=0 then
   begin
   BValider.Visible:=False;
   BPrint.Visible:=False;
   BDetail.Visible:=False;
   Btri.Visible:=False;
   BRecalcul.Visible:=False;
   G_LIG.Enabled:=False
   end ;
if Action = taCreat then BDelete.Visible := false;
if TobPiece.Detail.Count > G_LIG.RowCount then G_LIG.RowCount:=TobPiece.Detail.Count ;
TobPiece.PutGridDetail(G_LIG,False,True,LesCols);
if GZZ_SUGGEST.Text <> '' then
    begin
    GZZ_SUGGESTExit(Sender);
    if G_LIG.Objects[0, 1] <> nil then
        DataLigne.Lines.Text := TOB(G_LIG.Objects[0, 1]).GetValue('GL_BLOCNOTE');
    end;
end;

procedure TFResulSuggestion.GZZ_SUGGESTExit(Sender: TObject);
var
    TobArt, TobTemp : TOB;
//    Q : TQuery;
    Select : string;
    i_ind1 : integer;
    bTemp1, bTemp2 : boolean;

begin
Btemp1 := false;
Btemp2 := false;
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
{Select := 'Select * from GCTMPREAPPRO where '+sPrefixe+'_SUGGEST="' + GZZ_SUGGEST.Text + '" and (' +
          sPrefixe+'_ARTICLE is Null or ' + sPrefixe+'_ARTICLE="")';
Q := OpenSQL(Select, True);
if Q.EOF then
    begin
    MsgBox.Execute(0, Caption, '');
    Ferme(Q);
    Exit;
    end;
TobSuggest.SelectDB('', Q);
Ferme(Q); }
Recap.Lines.Text := TobPiece.GetValue('GP_BLOCNOTE');
i_ind1 := Recap.Lines.IndexOf('Choix du fournisseur dans l''ordre :') + 1;
if Trim(Recap.Lines.Strings[i_ind1]) = 'Délai d''approvisionnement' then Prio1 := 'APP';
if Trim(Recap.Lines.Strings[i_ind1]) = 'Fournisseur principal'      then Prio1 := 'FPR';
if Trim(Recap.Lines.Strings[i_ind1]) = 'Côte du fournisseur'        then Prio1 := 'COT';
if Trim(Recap.Lines.Strings[i_ind1]) = 'Meilleur prix tarif'        then Prio1 := 'MPT';
Inc(i_ind1);
if Trim(Recap.Lines.Strings[i_ind1]) = 'Délai d''approvisionnement' then Prio2 := 'APP';
if Trim(Recap.Lines.Strings[i_ind1]) = 'Fournisseur principal'      then Prio2 := 'FPR';
if Trim(Recap.Lines.Strings[i_ind1]) = 'Côte du fournisseur'        then Prio2 := 'COT';
if Trim(Recap.Lines.Strings[i_ind1]) = 'Meilleur prix tarif'        then Prio2 := 'MPT';
Inc(i_ind1);
if Trim(Recap.Lines.Strings[i_ind1]) = 'Délai d''approvisionnement' then Prio3 := 'APP';
if Trim(Recap.Lines.Strings[i_ind1]) = 'Fournisseur principal'      then Prio3 := 'FPR';
if Trim(Recap.Lines.Strings[i_ind1]) = 'Côte du fournisseur'        then Prio3 := 'COT';
if Trim(Recap.Lines.Strings[i_ind1]) = 'Meilleur prix tarif'        then Prio3 := 'MPT';
Inc(i_ind1);
if Trim(Recap.Lines.Strings[i_ind1]) = 'Délai d''approvisionnement' then Prio4 := 'APP';
if Trim(Recap.Lines.Strings[i_ind1]) = 'Fournisseur principal'      then Prio4 := 'FPR';
if Trim(Recap.Lines.Strings[i_ind1]) = 'Côte du fournisseur'        then Prio4 := 'COT';
if Trim(Recap.Lines.Strings[i_ind1]) = 'Meilleur prix tarif'        then Prio4 := 'MPT';

GZZ_CREATION.Text := TobPiece.GetValue('GP_DATEPIECE');
GZZ_SUGGEST.Text := TobPiece.GetValue('GP_NUMERO');
TousDepot := False;
DateDebut := '';
if Recap.Lines.IndexOf('Pas de selection sur les dépôts') >= 0 then TousDepot := True;
for i_ind1 := 0 to Recap.Lines.Count - 1 do
    if Copy(Recap.Lines[i_ind1], 0, 16) = 'Début de période' then DateDebut := Recap.Lines[i_ind1];
Select := ReadTokenPipe(DateDebut,':');
DateDebut := Trim(DateDebut);
// MOdif BTP
if DateDebut = '' then DateDebut := DateToStr(Date);
// --
//
//Select := 'Select * from GCTMPREAPPRO where '+sPrefixe+'_SUGGEST="' + GZZ_SUGGEST.Text + '" and (' +
//          sPrefixe+'_ARTICLE is not Null and ' + sPrefixe+'_ARTICLE<>"")';
//Q := OpenSQL(Select, True);
//if not Q.EOF then TobSuggest.LoadDetailDB('GCTMPREAPPRO', '', '', Q, False);
//Ferme(Q);

{ Mis en commentaire par BRL 13/08/2015 car inutile (GL_CODEARTICLE déjà chargé) et écrase le libellé de la pièce d'origine
TobArt := TOB.Create('ARTICLE', nil, -1);
for i_ind1 := 0 to TobPiece.Detail.Count - 1 do
    begin
    TobTemp := TobPiece.Detail[i_ind1];
    TobArt.PutValue('GA_ARTICLE', TobTemp.GetValue('GL_ARTICLE'));
    if TobArt.LoadDB then
        begin
        TobTemp.PutValue('GL_CODEARTICLE', TobArt.GetValue('GA_CODEARTICLE'));
        TobTemp.PutValue('GL_LIBELLE', TobArt.GetValue('GA_LIBELLE'));
        end;
    end;
TOBArt.Free;
}
TobPiece.PutGridDetail(G_LIG, False, False, LesCols);
G_LIG.Row := 1;
G_LIGRowEnter(Sender, 1, bTemp1, bTemp2);
end;

procedure TFResulSuggestion.G_LIGRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var i_ind1, i_ind2 : integer;
    LibelleDim, GrilleDim : string;
    TheText : String;
begin
if G_LIG.Objects[0, G_LIG.Row] = nil then begin Cancel := True; Exit; end;
BZoomArticle.Enabled:=True;
BZoomLigne.Enabled := true;
if TOB(G_LIG.Objects[0, G_LIG.Row]).GetValue('GL_TIERS')<>'' then BZoomTiers.Enabled:=True else BZoomTiers.Enabled:=False;
TheText := TOB(G_LIG.Objects[0, G_LIG.Row]).GetValue('GL_BLOCNOTE');
DataLigne.Lines.Text := TheText;
//bRecalcul.Enabled := False;
TGA_CODEDIM1.Caption := '';
TGA_CODEDIM2.Caption := '';
TGA_CODEDIM3.Caption := '';
TGA_CODEDIM4.Caption := '';
TGA_CODEDIM5.Caption := '';
if (Trim(Copy(TOB(G_LIG.Objects[0, G_LIG.Row]).GetValue('GL_ARTICLE'), 19, 15)) <> '') then
    begin
    LibelleDim := '';
    GrilleDim := '';
    LibelleDimensions(TOB(G_LIG.Objects[0, G_LIG.Row]).GetValue('GL_ARTICLE'), nil, LibelleDim, GrilleDim);
    TGA_CODEDIM1.Caption := ReadTokenSt(LibelleDim);
    TGA_CODEDIM2.Caption := ReadTokenSt(LibelleDim);
    TGA_CODEDIM3.Caption := ReadTokenSt(LibelleDim);
    TGA_CODEDIM4.Caption := ReadTokenSt(LibelleDim);
    TGA_CODEDIM5.Caption := ReadTokenSt(LibelleDim);
    if TGA_CODEDIM1.Caption <> '' then
        begin
        TGA_CODEDIM1.Caption := TGA_CODEDIM1.Caption + ' : ' + ReadTokenSt(GrilleDim);
        TGA_CODEDIM1.Visible := True;
        end;
    if TGA_CODEDIM2.Caption <> '' then
        begin
        TGA_CODEDIM2.Caption := TGA_CODEDIM2.Caption + ' : ' + ReadTokenSt(GrilleDim);
        TGA_CODEDIM2.Visible := True;
        end;
    if TGA_CODEDIM3.Caption <> '' then
        begin
        TGA_CODEDIM3.Caption := TGA_CODEDIM3.Caption + ' : ' + ReadTokenSt(GrilleDim);
        TGA_CODEDIM3.Visible := True;
        end;
    if TGA_CODEDIM4.Caption <> '' then
        begin
        TGA_CODEDIM4.Caption := TGA_CODEDIM4.Caption + ' : ' + ReadTokenSt(GrilleDim);
        TGA_CODEDIM4.Visible := True;
        end;
    if TGA_CODEDIM5.Caption <> '' then
        begin
        TGA_CODEDIM5.Caption := TGA_CODEDIM5.Caption + ' : ' + ReadTokenSt(GrilleDim);
        TGA_CODEDIM5.Visible := True;
        end;
    i_ind1 := TGA_CODEDIM1.Width + TGA_CODEDIM2.Width + TGA_CODEDIM3.Width +
              TGA_CODEDIM4.Width + TGA_CODEDIM5.Width;
    i_ind2 := PANDIM.Width - i_ind1 - 20;
    i_ind1 := i_ind2 div 4;
    TGA_CODEDIM1.Left := 10;
    TGA_CODEDIM2.Left := TGA_CODEDIM1.Left + TGA_CODEDIM1.Width + i_ind1;
    TGA_CODEDIM3.Left := TGA_CODEDIM2.Left + TGA_CODEDIM2.Width + i_ind1;
    TGA_CODEDIM4.Left := TGA_CODEDIM3.Left + TGA_CODEDIM3.Width + i_ind1;
    TGA_CODEDIM5.Left := TGA_CODEDIM4.Left + TGA_CODEDIM4.Width + i_ind1;
    end;
end;

procedure TFResulSuggestion.DataLigneChange(Sender: TObject);
var
    st1, st2, sDate, sPiece, sQPhy, sCli, sFou : string;
    sDaten1, sClin1 : string;
    sDaten2, sClin2 : string;
    Annee, Mois, Jour : string;
    i_ind1, i_ind2 : integer;
    {stkmin,} stkphy, stkphydeb, rcli, afou : double;
    TobDisp : TOB;
//    STemp : TChartSeries;
//    Q : TQuery;

begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
TobDisp := TOB.Create('Le Dispo', nil, -1);
st1 := 'Select GQ_STOCKMIN, GQ_STOCKMAX from DISPO where GQ_ARTICLE="' +
       TOB(G_LIG.Objects[0, G_LIG.Row]).GetValue('GL_ARTICLE') +
       '" and GQ_DEPOT="' + TOB(G_LIG.Objects[0, G_LIG.Row]).GetValue('GL_DEPOT') + '"';
TobDisp.LoadDetailDBFromSQL('DISPO', st1);

sDate    := DataLigne.Lines.Strings[0];
sPiece   := DataLigne.Lines.Strings[1];
sQPhy    := DataLigne.Lines.Strings[2];
sCli     := DataLigne.Lines.Strings[3];
sFou     := DataLigne.Lines.Strings[4];
sDaten1  := DataLigne.Lines.Strings[5];
sClin1   := DataLigne.Lines.Strings[6];
sDaten2  := DataLigne.Lines.Strings[7];
sClin2   := DataLigne.Lines.Strings[8];
i_ind2 := 1;
for i_ind1 := 0 to Dataligne.Lines.Count - 1 do
    begin
    st1 := DataLigne.Lines.Strings[i_ind1];
    if Copy(st1, 0, 4) <> 'Four' then Continue;
    st1 := Copy(st1, 8, 255);
    G_FOU.Cells[0, i_ind2] := ReadTokenST(st1);
    G_FOU.Cells[1, i_ind2] := ReadTokenST(st1);
    st2 := ReadTokenST(st1);
    if st2 = '' then st2 := '0';
    G_FOU.Cells[2, i_ind2] := StrF00(valeur(st2), GetParamSoc('SO_DECPRIX'));
    G_FOU.Cells[3, i_ind2] := ReadTokenST(st1);
    G_FOU.Cells[4, i_ind2] := ReadTokenST(st1);
    Inc(i_ind2);
    end;
G_FOU.RowCount := i_ind2;
// G_DET.SynEnabled := False; G_DET.CacheEdit;
for i_ind1 := 1 to G_DET.RowCount do G_DET.Rows[i_ind1].Clear;
{for i_ind1 := 0 to Graph1.SeriesList.Count - 1 do
    begin
    Graph1.Series[i_ind1].Clear;
    Graph1.Series[i_ind1].Active := True;
    end;
Graph1.LeftAxis.Minimum := 0;
Graph1.LeftAxis.Maximum := 0;  }
EclateDate(StrToDate(DateDebut), Annee, Mois, Jour);
if VH_GC.GCPeriodeStock = 'QUI' then
    if StrToInt(Jour) < 15 then Mois := Mois + '1' else Mois := Mois + '2'
else if VH_GC.GCPeriodeStock = 'SEM' then
    Mois := Format('%.2d', [NumSemaine(StrToDate(DateDebut))]);
st1 := Annee + '/' + Mois;
//stkmin := TOB(G_LIG.Objects[0, G_LIG.Row]).GetValue('GL_STOCKMIN');
//stkmin := 0;
i_ind1 := 1;
i_ind2 := 0;
while sDate <> '' do
    begin
    Inc (i_ind2);
    G_DET.Cells[0, i_ind1] := ReadTokenSt(sDate);
    G_DET.Cells[1, i_ind1] := ReadTokenSt(sPiece);
    st2 := ReadTokenST(sQPhy);
    if st2 = '' then st2 := '0';
    stkphy := valeur(st2);
    st2 := ReadTokenST(sCli);
    if st2 = '' then st2 := '0';
    rcli   := valeur(st2);
    st2 := ReadTokenST(sFou);
    if st2 = '' then st2 := '0';
    afou   := valeur(st2);
    stkphydeb := 0;
    if i_ind2 = 1 then stkphydeb := stkphy;
    if (DateDebut <> '') and (G_DET.Cells[0, i_ind1] < st1) then
        begin
        stkphydeb := stkphydeb - rcli + afou;
        Continue;
        end;
    if i_ind1 = 1 then
        begin
        G_DET.Cells[2, i_ind1] := StrF00(stkphydeb, GetParamSoc('SO_DECQTE'));
        stkphy := stkphydeb;
        end
        else
        G_DET.Cells[2, i_ind1] := StrF00(stkphy, GetParamSoc('SO_DECQTE'));
    G_DET.Cells[3, i_ind1] := StrF00(rcli, GetParamSoc('SO_DECQTE'));
    G_DET.Cells[4, i_ind1] := StrF00(afou, GetParamSoc('SO_DECQTE'));
    if i_ind1 = 1 then
        G_DET.Cells[6, i_ind1] := StrF00(stkphy - rcli + afou, GetParamSoc('SO_DECQTE'))
        else
        G_DET.Cells[6, i_ind1] := StrF00(valeur(G_DET.Cells[5, i_ind1 - 1]) - rcli + afou, GetParamSoc('SO_DECQTE'));
//    if StrToFloat(G_DET.Cells[5, i_ind1]) < stkmin then G_DET.Cells.Font.Color := clRed;
    Inc (i_ind1);
    end;

if (DateDebut <> '') and (Pos('.', DetailEvol.Caption) = 0) then
    DetailEvol.Caption := DetailEvol.Caption + '. Stock physique théorique au ' + DateDebut;
if TOBDisp.detail.count > 0 then
begin
  if TobDisp.Detail[0].GetValue('GQ_STOCKMIN') > 0 then
      LStockMin.Caption := 'Stock mini : ' + StrF00(TobDisp.Detail[0].GetValue('GQ_STOCKMIN'), GetParamSoc('SO_DECQTE'));
  if TobDisp.Detail[0].GetValue('GQ_STOCKMAX') > 0 then
      LStockMax.Caption := 'Stock Maxi : ' + StrF00(TobDisp.Detail[0].GetValue('GQ_STOCKMAX'), GetParamSoc('SO_DECQTE'));
  Panel1.Visible := True;
  if (TobDisp.Detail[0].GetValue('GQ_STOCKMIN') = 0) and
     (TobDisp.Detail[0].GetValue('GQ_STOCKMAX') = 0) then Panel1.Visible := False;
end else
begin
  LStockMin.Caption := 'Stock mini : Non géré en stock';
  LStockMax.Caption := 'Stock Maxi : Non géré en stock';
  Panel1.Visible := True;
end;
TobDisp.Free;
{for i_ind1 := 0 to Graph2.SeriesList.Count - 1 do
    begin
    Graph2.Series[i_ind1].Clear;
    Graph2.Series[i_ind1].Active := True;
    end;
Graph2.LeftAxis.Minimum := 0;
Graph2.LeftAxis.Maximum := 0;
i_ind1 := 1;
while sDateN1 <> '' do
    begin
    if i_ind1 > G_MvtsN1.RowCount - 1 then G_MvtsN1.RowCount := G_MvtsN1.RowCount + 25;
    st1 := ReadTokenSt(sDepotN1);
    if st1 <> '' then G_MvtsN1.Cells[0, i_ind1] := RechDom('GCDEPOT', st1, False);
    G_MvtsN1.Cells[1, i_ind1] := ReadTokenSt(sDateN1);
    G_MvtsN1.Cells[2, i_ind1] := ReadTokenSt(sCliN1);
    Inc (i_ind1);
    end;
for i_ind1 := 0 to Graph3.SeriesList.Count - 1 do
    begin
    Graph3.Series[i_ind1].Clear;
    Graph3.Series[i_ind1].Active := True;
    end;
Graph3.LeftAxis.Minimum := 0;
Graph3.LeftAxis.Maximum := 0;
for i_ind1 := 0 to Graph4.SeriesList.Count - 1 do
    begin
    Graph4.Series[i_ind1].Clear;
    Graph4.Series[i_ind1].Active := True;
    end;
Graph4.LeftAxis.Minimum := 0;
Graph4.LeftAxis.Maximum := 0;
i_ind1 := 1;
while sDateN2 <> '' do
    begin
    if i_ind1 > G_MvtsN2.RowCount - 1 then G_MvtsN2.RowCount := G_MvtsN2.RowCount + 25;
    st1 := ReadTokenSt(sDepotN2);
    if st1 <> '' then G_MvtsN2.Cells[0, i_ind1] := RechDom('GCDEPOT', st1, False);
    G_MvtsN2.Cells[1, i_ind1] := ReadTokenSt(sDateN2);
    G_MvtsN2.Cells[2, i_ind1] := ReadTokenSt(sCliN2);
    Inc (i_ind1);
    end;
// G_DET.SynEnabled := True; G_DET.MontreEdit;

if i_nbdep = 1 then
    begin
    for i_ind1 := 1 to G_DET.RowCount - 1 do
        begin
        if G_DET.Cells[6, i_ind1] = '' then break;
        Graph1.Series[0].Add(StrToFloat(G_DET.Cells[6, i_ind1]), G_DET.Cells[1, i_ind1], clTeeColor);
        Graph4.Series[0].Add(StrToFloat(G_DET.Cells[6, i_ind1]), G_DET.Cells[1, i_ind1], clTeeColor);
        end;
    Graph1.LeftAxis.Maximum := Graph1.Series[0].MaxYValue + 10;
    Graph1.LeftAxis.Minimum := Graph1.Series[0].MinYValue - 10;
    for i_ind1 := 1 to G_MvtsN1.RowCount - 1 do
        begin
        if G_MvtsN1.Cells[1, i_ind1] = '' then break;
        Graph2.Series[0].Add(StrToFloat(G_MvtsN1.Cells[2, i_ind1]), G_MvtsN1.Cells[0, i_ind1], clTeeColor);
        Graph4.Series[1].Add(StrToFloat(G_MvtsN1.Cells[2, i_ind1]), G_MvtsN1.Cells[0, i_ind1], clTeeColor);
        end;
    Graph2.LeftAxis.Maximum := Graph2.Series[0].MaxYValue + 10;
    Graph2.LeftAxis.Minimum := Graph2.Series[0].MinYValue - 10;
    for i_ind1 := 1 to G_MvtsN2.RowCount - 1 do
        begin
        if G_MvtsN2.Cells[1, i_ind1] = '' then break;
        Graph3.Series[0].Add(StrToFloat(G_MvtsN2.Cells[2, i_ind1]), G_MvtsN2.Cells[0, i_ind1], clTeeColor);
        Graph4.Series[2].Add(StrToFloat(G_MvtsN2.Cells[2, i_ind1]), G_MvtsN2.Cells[0, i_ind1], clTeeColor);
        end;
    Graph3.LeftAxis.Maximum := Graph3.Series[0].MaxYValue + 10;
    Graph3.LeftAxis.Minimum := Graph3.Series[0].MinYValue - 10;
    Graph4.LeftAxis.Maximum := Max(Max(Graph4.Series[0].MaxYValue, Graph4.Series[1].MaxYValue),
                                   Graph4.Series[2].MaxYValue) + 10;
    Graph4.LeftAxis.Minimum := Min(Min(Graph4.Series[0].MinYValue, Graph4.Series[1].MinYValue),
                                   Graph4.Series[2].MinYValue) + 10;
    RG_Depot.Visible := False;
    end
    else
    begin
    for i_ind1 := i_nbdep to Graph1.SeriesList.Count - 1 do Graph1.Series[i_ind1].Active := False;
    for i_ind1 := i_nbdep to Graph2.SeriesList.Count - 1 do Graph2.Series[i_ind1].Active := False;
    for i_ind1 := i_nbdep to Graph3.SeriesList.Count - 1 do Graph3.Series[i_ind1].Active := False;
    i_nbdep := 0;
    stRuptDepot := G_DET.Cells[0, 1];
    RG_Depot.Items.Text := G_DET.Cells[0, 1];
    for i_ind1 := 1 to G_DET.RowCount - 1 do
        begin
        if G_DET.Cells[6, i_ind1] = '' then break;
        if stRuptDepot <> G_DET.Cells[0, i_ind1] then
            begin
            Graph1.LeftAxis.Maximum := Max(Graph1.LeftAxis.Maximum, Graph1.Series[i_nbdep].MaxYValue + 10);
            Graph1.LeftAxis.Minimum := Min(Graph1.LeftAxis.Minimum, Graph1.Series[i_nbdep].MinYValue - 10);
            Graph1.Series[i_nbdep].Title := stRuptDepot;
            Inc(i_nbdep);
            stRuptDepot := G_DET.Cells[0, i_ind1];
            RG_Depot.Items.Add(G_DET.Cells[0, i_ind1]);
            end;
        Graph1.Series[i_nbdep].Add(StrToFloat(G_DET.Cells[6, i_ind1]), G_DET.Cells[1, i_ind1], clTeeColor);
        end;
    Graph1.LeftAxis.Maximum := Max(Graph1.LeftAxis.Maximum, Graph1.Series[i_nbdep].MaxYValue + 10);
    Graph1.LeftAxis.Minimum := Min(Graph1.LeftAxis.Minimum, Graph1.Series[i_nbdep].MinYValue - 10);
    Graph1.Series[i_nbdep].Title := stRuptDepot;
    Graph1.LeftAxis.AdjustMaxMin;
    Graph1.Invalidate;
//
    i_nbdep := 0;
    stRuptDepot := G_MvtsN1.Cells[0, 1];
    for i_ind1 := 1 to G_MvtsN1.RowCount - 1 do
        begin
        if G_MvtsN1.Cells[1, i_ind1] = '' then break;
        if stRuptDepot <> G_MvtsN1.Cells[0, i_ind1] then
            begin
            Graph2.LeftAxis.Maximum := Max(Graph2.LeftAxis.Maximum, Graph2.Series[i_nbdep].MaxYValue + 10);
            Graph2.LeftAxis.Minimum := Min(Graph2.LeftAxis.Minimum, Graph2.Series[i_nbdep].MinYValue - 10);
            Graph2.Series[i_nbdep].Title := stRuptDepot;
            Inc(i_nbdep);
            stRuptDepot := G_MvtsN1.Cells[0, i_ind1];
            end;
        Graph2.Series[i_nbdep].Add(StrToFloat(G_MvtsN1.Cells[2, i_ind1]), G_MvtsN1.Cells[1, i_ind1], clTeeColor);
        end;
    Graph2.LeftAxis.Maximum := Max(Graph2.LeftAxis.Maximum, Graph2.Series[i_nbdep].MaxYValue + 10);
    Graph2.LeftAxis.Minimum := Min(Graph2.LeftAxis.Minimum, Graph2.Series[i_nbdep].MinYValue - 10);
    Graph2.Series[i_nbdep].Title := stRuptDepot;
    Graph2.LeftAxis.AdjustMaxMin;
    Graph2.Invalidate;
//
    i_nbdep := 0;
    stRuptDepot := G_MvtsN2.Cells[0, 1];
    for i_ind1 := 1 to G_MvtsN2.RowCount - 1 do
        begin
        if G_MvtsN2.Cells[1, i_ind1] = '' then break;
        if stRuptDepot <> G_MvtsN2.Cells[0, i_ind1] then
            begin
            Graph3.LeftAxis.Maximum := Max(Graph3.LeftAxis.Maximum, Graph3.Series[i_nbdep].MaxYValue + 10);
            Graph3.LeftAxis.Minimum := Min(Graph3.LeftAxis.Minimum, Graph3.Series[i_nbdep].MinYValue - 10);
            Graph3.Series[i_nbdep].Title := stRuptDepot;
            Inc(i_nbdep);
            stRuptDepot := G_MvtsN2.Cells[0, i_ind1];
            end;
        Graph3.Series[i_nbdep].Add(StrToFloat(G_MvtsN2.Cells[2, i_ind1]), G_MvtsN2.Cells[1, i_ind1], clTeeColor);
        end;
    Graph3.LeftAxis.Maximum := Max(Graph3.LeftAxis.Maximum, Graph3.Series[i_nbdep].MaxYValue + 10);
    Graph3.LeftAxis.Minimum := Min(Graph3.LeftAxis.Minimum, Graph3.Series[i_nbdep].MinYValue - 10);
    Graph3.Series[i_nbdep].Title := stRuptDepot;
    Graph3.LeftAxis.AdjustMaxMin;
    Graph3.Invalidate;
//    for i_ind1 := 0 to Graph1.SeriesList.Count - 1 do
//        Graph1.Series[i_ind1].RePaint;
    RG_Depot.Visible := True;
    RG_Depot.ItemIndex := 0;
    end;  }
end;

procedure TFResulSuggestion.bDetailClick(Sender: TObject);
begin
DetailEvol.Visible := not DetailEvol.Visible;
end;

procedure TFResulSuggestion.bTriClick(Sender: TObject);
begin
DetailFourni.Visible := not DetailFourni.Visible;
end;

{procedure TFResulSuggestion.bCourbesClick(Sender: TObject);
begin
Courbes.Visible := not Courbes.Visible;
end;

procedure TFResulSuggestion.RG_DepotClick(Sender: TObject);
var
    i_ind1, i_ind2 : integer;

begin
i_ind2 := RG_Depot.ItemIndex;
for i_ind1 := 0 to Graph4.SeriesList.Count - 1 do
    begin
    Graph4.Series[i_ind1].Active := True;
    Graph4.Series[i_ind1].Clear;
    end;
Graph4.Series[0].Title := 'Année en cours';
Graph4.Series[1].Title := 'Année N - 1';
Graph4.Series[2].Title := 'Année N - 2';
for i_ind1 := 1 to G_DET.RowCount - 1 do
    begin
    if G_DET.Cells[6, i_ind1] = '' then break;
    if G_DET.Cells[0, i_ind1] = RG_Depot.Items.Strings[i_ind2] then
        Graph4.Series[0].Add(StrToFloat(G_DET.Cells[6, i_ind1]),
                             Copy(G_DET.Cells[1, i_ind1], Pos('/', G_DET.Cells[1, i_ind1]) + 1, 255),
                             clTeeColor);
    end;
for i_ind1 := 1 to G_MvtsN1.RowCount - 1 do
    begin
    if G_MvtsN1.Cells[1, i_ind1] = '' then break;
    if G_MvtsN1.Cells[0, i_ind1] = RG_Depot.Items.Strings[i_ind2] then
        Graph4.Series[1].Add(StrToFloat(G_MvtsN1.Cells[2, i_ind1]),
                             Copy(G_MvtsN1.Cells[1, i_ind1], Pos('/', G_MvtsN1.Cells[1, i_ind1]) + 1, 255),
                             clTeeColor);
    end;
for i_ind1 := 1 to G_MvtsN2.RowCount - 1 do
    begin
    if G_MvtsN2.Cells[1, i_ind1] = '' then break;
    if G_MvtsN2.Cells[0, i_ind1] = RG_Depot.Items.Strings[i_ind2] then
        Graph4.Series[2].Add(StrToFloat(G_MvtsN2.Cells[2, i_ind1]),
                             Copy(G_MvtsN2.Cells[1, i_ind1], Pos('/', G_MvtsN2.Cells[1, i_ind1]) + 1, 255),
                             clTeeColor);
    end;
Graph4.LeftAxis.Maximum := Max(Max(Graph4.Series[0].MaxYValue, Graph4.Series[1].MaxYValue),
                               Graph4.Series[2].MaxYValue) + 10;
Graph4.LeftAxis.Minimum := Min(Min(Graph4.Series[0].MinYValue, Graph4.Series[1].MinYValue),
                               Graph4.Series[2].MinYValue) + 10;
end;  }

/////////////////////////////////////////////////////////////////////////////
procedure TFResulSuggestion.bPrintClick(Sender: TObject);
Var
    TobTiers : TOB;
begin
CleDoc := TOB2CleDoc(TobPiece);
TobTiers := Tob.Create('TIERS', nil, -1);
ImprimerLaPiece(TOBPiece, TobTiers, CleDoc) ;
end;

procedure TFResulSuggestion.G_LIGCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
    Q : TQuery;
    TOBL : TOB;
    FUA,FUS,CoefUaUs : double;
//    Select : string;

begin
  if G_LIG.cells[Acol,Arow]= stcell then exit;
  TOBL := TOB(G_LIG.Objects[0, ARow]);
  FUS := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTESTO'));
  FUA := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTEACH'));
  CoefuaUs := TobL.GetValue('GL_COEFCONVQTE');
{$IFDEF BTP}
	if (ACol = 3) or (ACol = 5) then
{$ELSE}
	if (ACol = 3) or (ACol = 4) then
{$ENDIF}
  begin
    if not IsNumeric(G_LIG.Cells[ACol, ARow]) then
    begin
      Cancel := True;
      Exit;
    end;
  {$IFDEF BTP}
    if (ACol=3) then
    begin
      if not ControleQteBesoin (TOBL,Valeur(G_LIG.Cells[ACol, ARow])) then
      begin
        Hpiece.Execute (3,caption,'');
        cancel := true;
      end;
    	AjouteMontantAchat (TOBL,'-');
      TOBL.PutValue('QTEACHAT', Valeur(G_LIG.Cells[ACol, ARow]));
      if CoefUaUs <> 0 then
      begin
      	TOBL.PutValue('GL_QTESTOCK', arrondi(TOBL.GetValue('QTEACHAT')*CoefUaUs,V_PGI.okDecQ));
      end else
      begin
      	TOBL.PutValue('GL_QTESTOCK', arrondi(TOBL.GetValue('QTEACHAT')*FUA/FUS,V_PGI.okdecQ));
      end;
      TOBL.PutValue('GL_QTEFACT', TOBL.GetValue('GL_QTESTOCK'));
      TOBL.PutValue('GL_QTERESTE', TOBL.GetValue('GL_QTESTOCK')); { NEWPIECE }
    	AjouteMontantAchat (TOBL,'+');
    	TOBL.PutLigneGrid (G_LIG,Arow,false,false,LesCols);
    end;
  {$ELSE}
    if (ACol=3) then
    begin
      TOB(G_LIG.Objects[0, ARow]).PutValue('GL_QTESTOCK', Valeur(G_LIG.Cells[ACol, ARow]));
      TOB(G_LIG.Objects[0, ARow]).PutValue('GL_QTEFACT', Valeur(G_LIG.Cells[ACol, ARow]));
      TOB(G_LIG.Objects[0, ARow]).PutValue('GL_QTERESTE', Valeur(G_LIG.Cells[ACol, ARow])); { NEWPIECE }
    end;
  {$ENDIF}

  {$IFDEF BTP}
    if (ACol=5) then
    begin
    	AjouteMontantAchat (TOBL,'-');
      TOBL.PutValue('PUACHAT', Valeur(G_LIG.Cells[ACol, ARow]));
      if CoefuaUs <> 0 then
      begin
      	TOBL.PutValue('GL_PUHT',Arrondi(TOBL.GetValue('PUACHAT')/ CoefUaUs,V_PGI.okdecP));
      end else
      begin
      	TOBL.PutValue('GL_PUHT',Arrondi(TOBL.GetValue('PUACHAT')/FUA*FUS,V_PGI.okdecP));
      end;
      TOBL.PutValue('GL_PUHTDEV', TOBL.GetValue('GL_PUHT'));
    	AjouteMontantAchat (TOBL,'+');
      TOB(G_LIG.Objects[0, ARow]).PutValue('GL_MTRESTE', TOBL.GetValue('GL_QTEFACT') * TOBL.GetValue('GL_PUHTDEV')); { NEWPIECE }
    	TOBL.PutLigneGrid (G_LIG,Arow,false,false,LesCols);
    end;
  {$ELSE}
    if (ACol=4) then
    begin
      TOB(G_LIG.Objects[0, ARow]).PutValue('GL_PUHT', Valeur(G_LIG.Cells[ACol, ARow]));
      TOB(G_LIG.Objects[0, ARow]).PutValue('GL_PUHTDEV', Valeur(G_LIG.Cells[ACol, ARow]));
//      TOB(G_LIG.Objects[0, ARow]).PutValue('GL_PUHTBASE', Valeur(G_LIG.Cells[ACol, ARow]));
    end;
  {$ENDIF}
  end;
  {$IFDEF BTP}
  if (ACol = 7)  then
  {$ELSE}
  if (ACol = 5) and (G_LIG.Cells[ACol, ARow] <> '') then
  {$ENDIF}
  begin
    if (G_LIG.Cells[ACol, ARow] <> '') then
    begin
    	if TOB(G_LIG.Objects[0, ARow]).GetValue('GL_TIERS') <> G_LIG.Cells[ACol, ARow] then
      begin
        if not ControleFournisseurOk (G_LIG.Cells[ACol, ARow]) then
        begin
          Cancel := True;
          HPiece.Execute(4,caption,'');
          Exit;
        end;
        G_LIG.Cells[ACol, ARow] := UpperCase (G_LIG.Cells[ACol, ARow]);
        TOB(G_LIG.Objects[0, ARow]).PutValue('GL_TIERS', G_LIG.Cells[ACol, ARow]);
        // TODO
        // recupere prix catalogue et UA
        TraiteChangeFour (TOB(G_LIG.Objects[0, ARow]),Arow);
        //
        Ferme(Q);
      end;
    end else
    begin
      TOB(G_LIG.Objects[0, ARow]).PutValue('GL_TIERS', G_LIG.Cells[ACol, ARow]);
      TOB(G_LIG.Objects[0, ARow]).PutValue('LIBTIERS', '');
      G_LIG.Cells[ACol+1, ARow] := '';
    end;
  end;
end;

procedure TFResulSuggestion.BValiderClick(Sender: TObject);
Var i_ind1 : integer;
    TobTemp : TOB;
    cancel : boolean;
    Arow,Acol : integer;
begin
cancel := false;
Acol := G_LIG.Col;
Arow := G_LIG.Row;
G_LIGCellExit(self,Acol,Arow,cancel);
for i_ind1 := 0 to TobPiece.Detail.Count - 1 do
    begin
    TobTemp := TobPiece.Detail[i_ind1];
//    if TobTemp.GetValue('GL_TIERS') <> '' then
//        begin
        TobTemp.PutValue('GL_FOURNISSEUR', TobTemp.GetValue('GL_TIERS'));
        TobTemp.PutValue('GL_TIERS', '');
//        end;
    TobTemp.PutValue('GL_TYPEDIM', 'NOR');
    end;
if Action = TaModif then Transactions(ModifiePiece,5)
else if Action <> taConsult then
begin
    if Transactions(EnregistrePiece,5) <> oeOk then HPiece.Execute(2,caption,'')
                                               else MontreNumero(TOBPiece);
end;
Close;
end;

procedure TFResulSuggestion.ModifiePiece;
begin
TOBPiece.SetAllModifie(true);
if V_PGI.IoError = OeOk then TobPiece.UpdateDB(true);
end;

procedure TFResulSuggestion.EnregistrePiece;
var st1 : string;
    i_ind1 : integer;
    Depot : string ;
begin
st1 := GetInfoParPiece('REA','GPP_ACTIONFINI');
if st1 = 'TRA' then TOBPiece.PutValue('GP_VIVANTE', 'X') else TOBPiece.PutValue('GP_VIVANTE', '-');
if (GetInfoParPiece('REA','GPP_HISTORIQUE') = 'X') or
   (TOBPiece.GetValue('GP_VIVANTE') = 'X') then
    begin
    SetNumeroDefinitif(TOBPiece, nil, nil, nil, nil, nil, nil, nil, nil);
    for i_ind1 := 0 to TobPiece.Detail.Count - 1 do
        BEGIN
        Depot:=TobPiece.Detail[i_ind1].GetValue('GL_DEPOT');
        PieceVersLigne(TOBPiece, TobPiece.Detail[i_ind1]);
        TobPiece.Detail[i_ind1].PutValue('GL_DEPOT',Depot);
        TobPiece.Detail[i_ind1].PutValue('GL_NUMORDRE',TobPiece.Detail[i_ind1].GetValue('GL_NUMLIGNE'));
        END;
    if not TobPiece.InsertDB(nil) then V_PGI.ioError := oeUnknown;
{$IFDEF BTP}
    if V_PGI.Ioerror = OeOk then MiseEnPlaceLienReappro (TOBPiece,TOBLien);
    if V_PGI.Ioerror = OeOk then PurgeLesLiens (TOBLIen);
    if (TOBLien.detail.count >0) and ( V_PGI.Ioerror = OeOk ) then
    begin
    	TOBLien.SetAllModifie (True);
      if not TOBLien.InsertDb(nil) then V_PGI.IoError := OeUnknown;
    end;
{$ENDIF}
    if V_PGI.Ioerror = OeOk then NumeroteLignesGC(Nil,TOBPiece) ; // ?? Huuuuuuu ?? ca sert a quoi puisque déjà enregistré
    end;
end;

procedure TFResulSuggestion.G_LIGCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
{$IFDEF BTP}
G_LIG.ElipsisButton := (G_LIG.Col = 7);
{$ELSE}
G_LIG.ElipsisButton := (G_LIG.Col = 5);
{$ENDIF}
stcell := G_LIG.Cells [G_LIG.col,G_LIG.row];
end;

procedure TFResulSuggestion.G_LIGElipsisClick(Sender: TObject);
{$IFDEF BTP}
var
//    st1 : string;
//    Q : TQuery;
//    Nb : integer;
    Coords,Rect : Trect;
//    X,Y : integer;
    TOBL : TOB;
    FUS,FUV : double;
{$ENDIF}
begin
  TOBL := TOB(G_LIG.Objects[0, G_LIG.Row]);
  FUS := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTESTO'));
  FUV := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTEVTE'));
  if FUS = 0 then FUS := 1;
  if FUV = 0 then FUV := 1;
  if (FUS <> 1) or (FUV <> 1) then
  begin
    CatalogueClick (self); // on autorise pas la sélection sur fournisseur si US ou UV > 0
    exit;
  end;
  Coords := G_LIG.BoundsRect;
  Rect := G_LIG.CellRect (G_LIG.col, G_LIG.row);
  PChoixFour.Popup (Coords.left+Rect.Left ,Coords.Top+Rect.Top );
end;

procedure TFResulSuggestion.G_FOUSorted(Sender: TObject);
begin
;
end;

procedure TFResulSuggestion.bRecalculClick(Sender: TObject);
begin
AjouteMontantAchat(TOB(G_LIG.Objects[0, G_LIG.Row]),'-');
Calcul_Fournisseur(TOB(G_LIG.Objects[0, G_LIG.Row]), Prio1+';'+Prio2+';'+Prio3+';'+Prio4);
AjouteMontantAchat(TOB(G_LIG.Objects[0, G_LIG.Row]),'+');
TobPiece.PutGridDetail(G_LIG, False, False, LesCols);
end;

procedure TFResulSuggestion.FormClose(Sender: TObject; var Action: TCloseAction);
begin
PurgePop(POPZ) ;
TobPiece.Free; TobPiece := nil;
TOBLien.free; TOBLien := nil;
end;

procedure TFResulSuggestion.FormCreate(Sender: TObject);
begin
TobPiece := TOB.Create('PIECE', nil, -1);
// Ajout BTP
AddLesSupEntete (TOBPiece);
//TobLigne := TOB.Create('', nil, -1);
TOBLien := TOB.Create ('LES LIENS', nil ,-1);
end;

procedure TFResulSuggestion.G_FOUDblClick(Sender: TObject);
var st1 : string;
{$IFDEF BTP}
  libart, CodeArtic : string;
{$ENDIF}
begin
if G_FOU.Cells[1, G_FOU.Row] = '' then Exit;
st1 := G_FOU.Cells[1, G_FOU.Row]+';'+G_FOU.Cells[0, G_FOU.Row];
{$IFDEF BTP}
CodeArtic := TOB(G_LIG.Objects[0, G_LIG.Row]).GetValue('GL_ARTICLE');
libart := TOB(G_LIG.Objects[0, G_LIG.Row]).GetValue('GL_LIBELLE');
AGLLanceFiche('GC', 'GCCATALOGU_SAISI3', '', st1, 'LIB='+libart+';ARTICLE='+CodeArtic);
{$ELSE}
AGLLanceFiche('GC', 'GCCATALOGU_SAISI3', '', st1, '');
{$ENDIF}
bRecalculClick(Sender);
DataLigne.Lines.Text := TOB(G_LIG.Objects[0, G_LIG.Row]).GetValue('GL_BLOCNOTE');
end;

procedure TFResulSuggestion.BMenuZoomMouseEnter(Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;
PopZoom97(BMenuZoom,POPZ) ;
end;

procedure TFResulSuggestion.BZoomArticleClick(Sender: TObject);
Var RefUnique : string ;
begin
if Not BZoomArticle.Enabled then Exit ;
RefUnique:=TOB(G_LIG.Objects[0, G_LIG.Row]).GetValue('GL_ARTICLE') ;
{$IFNDEF GPAO}
{$IFDEF BTP}
V_PGI.DispatchTT(7, taConsult, RefUnique, '', 'MONOFICHE');
{$ELSE}
AGLLanceFiche('GC','GCARTICLE','',RefUnique,'ACTION=CONSULTATION;MONOFICHE');
{$ENDIF}
{$ELSE}
V_PGI.DispatchTT(7, taConsult, RefUnique, '', 'MONOFICHE');
{$ENDIF}
end;

procedure TFResulSuggestion.BZoomTiersClick(Sender: TObject);
Var Tiers,Auxiliaire : string ;
begin
if Not BZoomTiers.Enabled then Exit ;
Tiers:=TOB(G_LIG.Objects[0, G_LIG.Row]).GetValue('GL_TIERS') ;
Auxiliaire := TiersAuxiliaire (Tiers,false,'FOU');
//V_PGI.DispatchTT(12, taConsult, Tiers, '', '') ;
V_PGI.DispatchTT(12, taConsult, Auxiliaire, '', '') ;
end;

// AJOUTS BTP 20/06
Procedure AppelPieceReappro ( Parms : array of variant ; nb : integer ) ;
var CleDoc : R_CleDoc;
    StA,StM : string ;
    i_ind  : integer ;
    Action : TActionFiche ;
BEGIN
FillChar(CleDoc,Sizeof(CleDoc),#0) ;
StA:=String(Parms[0]) ; StM:=String(Parms[1]) ;
Action:=taModif ;
i_ind:=Pos('ACTION=',StM) ;
if i_ind>0 then
   BEGIN
   Delete(StM,1,i_ind+6) ; StM:=uppercase(ReadTokenSt(StM)) ;
   if StM='CREATION' then BEGIN Action:=taCreat ; END ;
   if StM='MODIFICATION' then BEGIN Action:=taModif ; END ;
   if StM='CONSULTATION' then BEGIN Action:=taConsult ; END ;
   END ;
   StringToCleDoc(StA,CleDoc) ;
   if CleDoc.NaturePiece<>'' then ModifSuggestion(CleDoc,Action) ;
END ;

procedure TFResulSuggestion.MiseEnPlaceRefFournisseur;
var indice  : integer;
    Q       : TQuery;
begin

  for indice := 0 to tobpiece.detail.count -1 do
  begin
    if TobPiece.detail[indice].GetValue('GL_FOURNISSEUR') <> '' then
      TobPiece.detail[indice].putValue('GL_TIERS',TobPiece.detail[indice].GetValue('GL_FOURNISSEUR'));
    TobPiece.detail[indice].AddChampSupValeur ('LIBTIERS','');
    Q := OpenSQL('Select T_LIBELLE from TIERS where T_TIERS="'+TobPiece.detail[indice].GetValue('GL_TIERS')+'"', True,-1,'',true);
    if Not Q.EOF then TobPiece.detail[indice].putValue('LIBTIERS',Q.Fields[0].AsString);
    ferme(Q);
  end;

end;

function TFResulSuggestion.TraiteDonneeEntree : boolean;
begin
  result := false;
  if cledoc.NaturePiece = '' then exit;
  loadlestobs;
  MiseEnPlaceRefFournisseur;
  GZZ_SUGGEST.Text := TOBPiece.GetValue('GP_NUMERO');
  PositionneEnAchat (TOBPiece);
  result := true;
end;

procedure TFResulSuggestion.loadlestobs;
var Q : Tquery;
begin
  Q:=OpenSQL('SELECT * FROM PIECE WHERE '+WherePiece(CleDoc,ttdPiece,False),True,-1,'',true) ;
  TOBPiece.SelectDB('',Q) ;
  Ferme(Q) ;
  // Lecture Lignes
  Q:=OpenSQL('SELECT * FROM LIGNE WHERE '+WherePiece(CleDoc,ttdLigne,False)+' ORDER BY GL_NUMLIGNE',True,-1,'',true) ;
  TOBPiece.LoadDetailDB('LIGNE','','',Q,False,True) ;
  Ferme(Q) ;
end;

procedure InitResulSuggestion();
begin
 RegisterAglProc( 'Entree_ResulSuggest', False , 2, Entree_ResulSuggestion);
 RegisterAglProc( 'AppelPieceReappro',False,2,AppelPieceReappro) ;
end;

procedure TFResulSuggestion.BDeleteClick(Sender: TObject);
Var St : String ;
begin
if HPiece.Execute(0,caption,'')<>mrYes then Exit ;
if Transactions(DetruitLaPiece,1)<>oeOk then
   BEGIN
   MessageAlerte(HTitres.Mess[0]) ;
   END else
   BEGIN
   St:='Pièce N° '+IntToStr(TOBPiece.GetValue('GP_NUMERO'))
   +', Indice '+IntToStr(TOBPiece.GetValue('GP_INDICEG'))
   +', Date '+DateToStr(TOBPiece.GetValue('GP_DATEPIECE'))
   +', Tiers '+TOBPiece.GetValue('GP_TIERS')
   +', Total HT de '+StrfPoint(TOBPiece.GetValue('GP_TOTALHTDEV'))+' '+RechDom('TTDEVISETOUTES',TOBPiece.GetValue('GP_DEVISE'),False) ;
   AjouteEvent(TOBPiece.GetValue('GP_NATUREPIECEG'),St);
   Close;
   END ;
end;

procedure TFResulSuggestion.DetruitLaPiece;
var nb : integer;
    st : string;
begin
st := 'DELETE FROM PIECE WHERE '+WherePiece(CleDoc,ttdPiece,False);
Nb:=ExecuteSQL(st) ;
if Nb<=0 then BEGIN V_PGI.IoError:=oeSaisie ; Exit ; END ;
if TOBPiece<>Nil then if TOBPiece.Detail.Count>0 then
   BEGIN
   Nb:=ExecuteSQL('DELETE FROM LIGNE WHERE '+WherePiece(CleDoc,ttdLigne,False)) ;
   if Nb<=0 then BEGIN V_PGI.IoError:=oeUnknown ; Exit ; END ;
   END ;
{$IFDEF BTP}
if V_PGI.Ioerror = oeOk then DetruitLienDevisChantier (TOBPiece);
{$ENDIF}
end;


Procedure TFResulSuggestion.AjouteEvent(NatureP, MessEvent : string) ;
Var QQ : TQuery ;
    MotifPiece  : TStrings ;
    NumEvent : integer ;
BEGIN
MotifPiece:=TStringList.Create ;
MotifPiece.Add(MessEvent) ;
NumEvent:=0;
QQ:=OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT',True,-1,'',true) ;
if Not QQ.EOF then NumEvent:=QQ.Fields[0].AsInteger ;
Ferme(QQ) ;
Inc(NumEvent) ;
QQ:=OpenSQL('SELECT * FROM JNALEVENT WHERE GEV_TYPEEVENT="GEN" AND GEV_NUMEVENT=-1',False) ;
QQ.Insert ; InitNew(QQ) ;
QQ.FindField('GEV_NUMEVENT').AsInteger:=NumEvent ;
QQ.FindField('GEV_TYPEEVENT').AsString:='SPI' ;
QQ.FindField('GEV_LIBELLE').AsString:=Copy('Suppression de '+RechDom('GCNATUREPIECEG',NatureP,False),1,35) ;
QQ.FindField('GEV_DATEEVENT').AsDateTime:=Date ;
QQ.FindField('GEV_UTILISATEUR').AsString:=V_PGI.User ;
QQ.FindField('GEV_ETATEVENT').AsString:='OK' ;
TMemoField(QQ.FindField('GEV_BLOCNOTE')).Assign(MotifPiece) ;
QQ.Post ;
Ferme(QQ) ;
MotifPiece.Free ;
END ;


procedure TFResulSuggestion.BZoomLigneClick(Sender: TObject);
{$IFNDEF BTP}
var starg : string;
{$ENDIF}
begin
TheTOB:=GetTOBLigne(TOBPiece,G_LIG.Row);
// Modif Btp
{$IFDEF BTP}
AGLLanceFiche ('BTP','BTCOMPLLIGNE','','','ACTION=CONSULTATION');
{$ELSE}
stArg:='ACTION=CONSULTATION';
if VH_GC.GCIfDefCEGID then
   if sender=nil then stArg:=stArg+';FORCEINFO' ;
if (ctxAffaire in V_PGI.PGIContexte) then
   AGLLanceFiche ('AFF','AFCOMPLLIGNE','','',stArg)
else
   AGLLanceFiche ('GC','GCCOMPLLIGNE','','',stArg) ;
{$ENDIF}
TheTob := Nil;
end;
// FIN AJOUTS BTP 20/06

procedure TFResulSuggestion.ListefournisseursClick(Sender: TObject);
var MaCle,Fournisseur,Article : string;
    TOBL : TOB;
    MtPa,CoefuaUs : double;
    Ua : string;
begin
	Ua := '';
  TOBL := TOB(G_LIG.Objects[0, G_LIG.Row]);
  fournisseur := TOBL.GetValue('GL_TIERS');
  Article := TOBL.GetValue('GL_ARTICLE');
  MaCle := '';
  MaCle := AGLLanceFiche('GC', 'GCFOURNISSEUR_MUL', 'T_NATUREAUXI=FOU;T_FERME=-', '', 'SELECTION');
  if MaCle <> '' then
  begin
    if ControleFournisseurOk (MaCle) then
    begin
      TOBL.PutValue('GL_TIERS', MaCle);
      TraiteChangeFour (TOBL,G_LIG.Row);
    end else
    begin
      PgiError ('Impossible : ce fournisseur est fermé');
    end;
  end;
end;

procedure TFResulSuggestion.CatalogueClick(Sender: TObject);
var TOBL : TOB;
    Fournisseur,Article,St1 : string;
begin
  TOBL := TOB(G_LIG.Objects[0, G_LIG.Row]);
  fournisseur := TOBL.GetValue('GL_TIERS');
  Article := TOBL.GetValue('GL_ARTICLE');
{$IFDEF BTP}
  st1 := LookupFournisseur (Fournisseur,Article,G_LIG.Cells[1, G_LIG.Row],false);
{$ENDIF}
  if (st1 <> '') then
  begin
    if ControleFournisseurOk (St1) then
    begin
      TOBL.PutValue('GL_TIERS', st1);
      TraiteChangeFour (TOBL,G_LIG.row);
    end else
    begin
      PgiError ('Impossible : ce fournisseur est fermé');
    end;
  end;
end;

procedure TFResulSuggestion.BChercherClick(Sender: TObject);
begin
  if G_LIG.RowCount <= 2 then Exit;
  FindFirst := True;
  FindLigne.Execute;
end;

procedure TFResulSuggestion.FindLigneFind(Sender: TObject);
begin
  Rechercher(G_LIG, FindLigne, FindFirst);
end;

procedure TFResulSuggestion.PositionneEnAchat (TOBPiece : TOB; ModeCreation : boolean=false);
var Indice : integer;
begin

  For Indice := 0 to TOBPiece.detail.count -1 do
  begin
    AjouteUniteAchat (TOBPiece.detail[Indice],ModeCreation);
  end;

end;

procedure TFResulSuggestion.AjouteUniteAchat(TOBL: TOB; ModeCreation : boolean);
var FUS     : Double;
    FUA     : Double;
    CoefuaUs: Double;
    CoefUSUv: Double;
    PQQ     : Double;
    //
		Tiers   : String;
    Article : String;
    UA      : String;
begin

    Tiers   := TOBL.GetValue('GL_TIERS');
    Article := TOBL.GetValue('GL_ARTICLE');

    GetInfoFromCatalog(Tiers,Article,UA,PQQ,CoefUaUs);

    if UA = '' then GetInfoFromArticle(Article,UA,PQQ,CoefUaUs);

    TOBL.putValue('GL_COEFCONVQTE',CoefUaUs);

    CoefUSUv := TOBL.GetDouble('GL_COEFCONVQTEVTE');

    if PQQ = 0 then PQQ := 1;

    TOBL.PutValue('GL_QUALIFQTEACH',UA);
    TOBL.putValue('GL_QUALIFQTEVTE',UA);

    FUS := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTESTO'));
    FUA := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTEACH'));

    if CoefUaUs <> 0 then
    begin
      TOBL.AddChampSupValeur ('QTEACHAT',Arrondi(TOBL.GEtVAlue('GL_QTESTOCK')/ CoefUaUs,V_PGI.okdecP) ,false);
      TOBL.AddChampSupValeur ('PUACHAT', Arrondi(TOBL.GEtVAlue('GL_PUHT')* CoefUaUs,V_PGI.okdecP),false);
    end
    else
    begin
      TOBL.AddChampSupValeur ('QTEACHAT',Arrondi(TOBL.GEtVAlue('GL_QTESTOCK')*FUS/FUA,V_PGI.okdecP) ,false);
      TOBL.AddChampSupValeur ('PUACHAT', Arrondi(TOBL.GEtVAlue('GL_PUHT')/FUS*FUA,V_PGI.OkdecP),false);
    end;

end;

procedure TFResulSuggestion.TraiteChangeFour (TOBL : TOB; Arow : integer);
var MtPa : double;
    UA : string;
    PQQ : double;
    FUS,FUA,CoefUaUs : double;
    Tiers,Article : string;
    Q : Tquery;
begin

    Tiers := TOBL.GetValue('GL_TIERS');

    Q := OpenSQL('Select T_LIBELLE from TIERS where T_TIERS="'+Tiers+'"', True,-1,'',true);
    if Not Q.EOF then TOBL.putValue('LIBTIERS',Q.Fields[0].AsString);
    ferme(Q);

    Article := TOBL.GetValue('GL_ARTICLE');

    FUS := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTESTO'));
    FUA := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTEACH'));

{$IFDEF BTP}
		AjouteMontantAchat(TOBL,'-');
    MtPa := RecupTarifAch (Tiers,Article,Ua,CoefuaUs,TurAchat);
    TOBL.PutValue('GL_QUALIFQTEACH',UA);
		TOBL.putValue('GL_QUALIFQTEVTE',UA);
    TobL.PutValue('GL_COEFCONVQTE',CoefUaUs);
    TOBL.PutValue ('PUACHAT',MtPa);
    //
    TOBL.PutValue('GL_COEFCONVQTE', CoefUaUs);
    //
    if CoefuaUs <> 0 then
    begin
      TOBL.PutValue('GL_PUHT',Arrondi(TOBL.GetValue('PUACHAT')/ CoefUaUs,V_PGI.okdecP));
    end else
    begin
      TOBL.PutValue('GL_PUHT',Arrondi(TOBL.GetValue('PUACHAT')*FUA/FUS,V_PGI.okdecP));
    end;
    //
    TOBL.PutValue('GL_PUHTDEV',TOBL.GetValue('GL_PUHT'));
    // Quantite
    if CoefuaUs <> 0 then
    begin
      TOBL.PutValue ('QTEACHAT',Arrondi(TOBL.GetValue('GL_QTESTOCK')/ CoefuaUs,V_PGI.okdecQ));
    end else
    begin
      TOBL.PutValue ('QTEACHAT',Arrondi(TOBL.GetValue('GL_QTESTOCK')*FUS/FUA,V_PGI.OkdecQ));
    end;
{$ELSE}
{$IFNDEF PGIMAJVER}
    MtPa := RecupTarifAch (Tiers,Article,TurStock);
    TOBL.PutValue('GL_PUHT', MtPa);
    TOBL.PutValue('GL_PUHTDEV', MtPa);
{$ENDIF}
{$ENDIF}
		AjouteMontantAchat(TOBL,'+');
    TOBL.PutLigneGrid (G_LIG,Arow,false,false,LesCols);
end;

{$IFDEF BTP}
function TFResulSuggestion.ControleQteBesoin(TOBL: TOB; Quantite: double) : boolean;
var QQ : Tquery;
    TOBLien : TOB;
    Req : string;
    TheClef : CleligneDevCha;
    CumulQte : double;
    Indice : integer;
begin
  result := true;
  CumulQte := 0;
  TOBLien := TOB.Create ('LES LINES',nil,-1);
  Req := 'SELECT BDA_REFD,BDA_NUMLD FROM LIENDEVCHA '+
         'WHERE BDA_REFC="'+EncodeLienDevCHA (TOBL)+'" AND ' +
         'BDA_NUMLC='+IntToStr(TOBL.getValue('GL_NUMLIGNE'));
  QQ := OpenSql (Req,true,-1,'',true);
  TRY
    if not QQ.eof then
    begin
      TOBLien.LoadDetailDB ('LIENDEVCHA','','',QQ,false);
      for Indice := 0 to TOBLien.detail.count -1 do
      begin
        theClef := DecodeLienDevCHA (TOBLien.detail[Indice].getValue('BDA_REFD'));
        ferme (QQ);
        req := 'SELECT GL_QTERESTE FROM LIGNE WHERE '+
               ' GL_NATUREPIECEG="'+TheClef.NaturePiece+'" AND GL_SOUCHE="'+TheClef.Souche+'"' +
               ' AND GL_NUMERO='+IntToStr(TheClef.NumeroPiece)+
               ' AND GL_INDICEG='+intToStr(TheClef.Indice) +
               ' AND GL_NUMLIGNE='+InttoStr(TOBLien.detail[Indice].getValue('BDA_NUMLD'))  ;
        QQ := OpenSql (Req,True,-1,'',true);
        if not QQ.eof then
        begin
          CumulQte := CumulQte + QQ.findField('GL_QTERESTE').asfloat;
        end;
      end;
      if CumulQte > Quantite then result := false;
    end;
  FINALLY
    Ferme (QQ);
    TOBLien.free;
  END;
end;
{$ENDIF}
procedure TFResulSuggestion.CumuleMontants(TOBL: TOB);
begin
	AjouteMontantBudget(TOBL);
	AjouteMontantAchat(TOBL,'+');
end;

procedure TFResulSuggestion.AjouteMontantAchat(TOBL: TOB; Sens: string);
begin
	if Sens = '+' then
  begin
  	TmontantAch.value  := TmontantAch.Value + Arrondi(TOBL.getValue('QTEACHAT')*TOBL.getValue('PUACHAT'),V_PGI.OkDecV);
  end else
  begin
  	TmontantAch.value  := TmontantAch.value - Arrondi(TOBL.getValue('QTEACHAT')*TOBL.getValue('PUACHAT'),V_PGI.OkDecV);
  end;
end;

procedure TFResulSuggestion.AjouteMontantBudget(TOBL: TOB);
begin
  TmontantBudget.value := TMontantBudget.Value + Arrondi(TOBL.getValue('GL_QTEPREVAVANC')*TOBL.getValue('GL_PUHTBASE'),V_PGI.OkDecV);
end;


Initialization
InitResulSuggestion();

end.

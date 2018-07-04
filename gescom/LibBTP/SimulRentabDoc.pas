unit SimulRentabDoc;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,DicoBTP,
{$IFDEF EAGLCLIENT}
  maineagl,
{$ELSE}
  Doc_Parser,DBCtrls, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} fe_main,
{$ENDIF}
  Hctrls, Grids, Mask, StdCtrls, HTB97, ExtCtrls,UTOB,SimulRentabUtil,ParamSoc,
  ImgList,Saisutil,HEnt1, ComCtrls,FactUtil,BTPUtil, HImgList,
  hmsgbox,M3FP, TntStdCtrls, TntComCtrls, TntGrids,UEntCommun;

type

  TFSimulRentab = class(TForm)
    PPIED     : TPanel;
    DockBottom: TDock97;
    Outils97  : TToolbar97;
    BDetailVal: TToolbarButton97;
    Valide97  : TToolbar97;
    BValider  : TToolbarButton97;
    BAbandon  : TToolbarButton97;
    BAide     : TToolbarButton97;
    HLabel1   : THLabel;
    IListeDeploie: THImageList;
    BAction   : TToolbarButton97;
    GS        : THGrid;
    TDetailVal: TToolWindow97;
    PHeures   : TPanel;
    QTE       : THNumEdit;
    LHEURES   : THLabel;
    MPA       : THNumEdit;
    MPR       : THNumEdit;
    MPV       : THNumEdit;
    COEFFR    : THNumEdit;
    COEFMARG  : THNumEdit;
    PVMOY     : THNumEdit;
    PRMOY     : THNumEdit;
    PAMOY     : THNumEdit;
    HLabel2   : THLabel;
    HLabel3   : THLabel;
    HLabel4   : THLabel;
    HLabel5   : THLabel;
    HLabel6   : THLabel;
    HLabel7   : THLabel;
    HLabel8   : THLabel;
    HLabel9   : THLabel;
    ApplicPAH : TToolbarButton97;
    Barbo     : TToolbarButton97;
    ApplicPRH : TToolbarButton97;
    ApplicPVH : TToolbarButton97;
    TTarborescence: TToolWindow97;
    TV        : THTreeView;
    TLIBMTTTC : THLabel;
    MPVTTC    : THNumEdit;
    TLIBPVMOYTTC: THLabel;
    PVMOYTTC  : THNumEdit;
    ApplicPVTTC: TToolbarButton97;
    HLabel10  : THLabel;
    COEFFC    : THNumEdit;
    COEFFG    : THNumEdit;
    HLabel11  : THLabel;

    procedure FormCreate(Sender: TObject);
    procedure FormClose (Sender: TObject; var Action: TCloseAction);
    procedure FormShow  (Sender: TObject);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowEnter (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit  (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure BActionClick (Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormCloseQuery (Sender: TObject; var CanClose: Boolean);
    procedure BDetailValClick(Sender: TObject);
    procedure TDetailValClose(Sender: TObject);
    procedure QTEChange(Sender: TObject);
    procedure MPAChange(Sender: TObject);
    procedure MPRChange(Sender: TObject);
    procedure MPVChange(Sender: TObject);
    procedure MPVTTCExit(Sender: TObject);
    procedure COEFFGChange(Sender: TObject);
    procedure COEFMARGChange(Sender: TObject);
    procedure PAMOYChange(Sender: TObject);
    procedure PRMOYChange(Sender: TObject);
    procedure PVMOYChange(Sender: TObject);
    procedure PVMOYTTCExit(Sender: TObject);
    procedure ApplicPAHClick(Sender: TObject);
    procedure ApplicPRHClick(Sender: TObject);
    procedure ApplicPVTTCClick(Sender: TObject);
    procedure BarboClick(Sender: TObject);
    procedure TTarborescenceClose(Sender: TObject);
    procedure TVClick(Sender: TObject);
    procedure GSMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure ApplicPVHClick(Sender: TObject);
    procedure QTEEnter(Sender: TObject);
    procedure MPAEnter(Sender: TObject);
    procedure MPREnter(Sender: TObject);
    procedure COEFMARGEnter(Sender: TObject);
    procedure MPVEnter(Sender: TObject);
    procedure MPVTTCEnter(Sender: TObject);
    procedure COEFFGEnter(Sender: TObject);
    procedure PAMOYEnter(Sender: TObject);
    procedure PRMOYEnter(Sender: TObject);
    procedure PVMOYEnter(Sender: TObject);
    procedure PVMOYTTCEnter(Sender: TObject);
  private
    { Déclarations privées }
		TOBPieceFrais : TOB;
    gestionInterdit : TSGestionInterdit;
    MontantFD : double; // Frais détaillé de chantier
//    BloquePv : boolean;
    TheQte,TheMpa,TheMPR,TheCoefFG,TheCoefMarg,TheMPV,TheMPVTTC,ThePAMOY,ThePRMOY,ThePVMOY,ThePVMOYTTC : double;
    StoredRow : integer;
    EnHt : boolean;
    DEV : Rdevise;
    CellCur : string;
    WhantVisible : boolean;
    Action : TActionFiche ;
    TOBSimulPres,TOBSimulPres_O: TOB;
    TOBPiece,TOBOuvrage,TOBPorcs : TOB;
    TOBBases,TOBBasesL : TOB;
    TOBArticles : TOB;
    Paragraphe : boolean;
    Cotraitance: Boolean;
    ListeSaisie : string;
    ValidateOk : boolean;
    IsModifie : boolean;
    CurrentTOb : TOB;
    Niveau,NBLigne : integer;
    GS_NUMLIG,GS_OPEN,GS_PARAG,GS_CATEGORIE,GS_NATURE,GS_CODE,GS_LIBELLE,GS_HRS,GS_VALACHAT : integer;
    GS_COEFFR,GS_COEFFG,GS_COEFFC,GS_VALPR,GS_COEFMARG,GS_VALPV : integer;
    CoefMargGlobalModif : boolean;  // permet de savoir si le coef de fg global document à été modifié ou pas
    ValCoefMarg : double; // valeur du coef de marge global saisie
    procedure definiGrilleSaisie;
    procedure DefiniNiveauVisible(Niveau: integer);
    procedure GSPostDrawcell(ACol, ARow: Integer; Canvas: TCanvas;
      AState: TGridDrawState);
    procedure ZoneSuivanteOuOk(var ACol, ARow: Integer;
      var Cancel: boolean);
    function ZoneAccessible(ACol, ARow: Integer): boolean;
    procedure PositionneLigne(Arow:integer;Acol : integer=-1);
    procedure DesactiveEventGS;
    procedure ActiveEventGS;
    procedure GSGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas;
      AState: TGridDrawState);
    procedure ModificationHrs (Arow : integer; NewQte : double = -1);
    procedure ModificationAchat(Arow: integer; NewQte : double = -1);
    procedure ModificationCoefFG(Arow: Integer; NewQte : double = -1);
    procedure ModificationCoefMARG(Arow: Integer; NewQte : double = -1);
    procedure ModificationRevient(Arow: Integer;NewQte : double = -1);
    procedure ModificationPV(Arow: Integer; NewQte : double = -1);
    procedure ModificationPVTTC(Arow: Integer; NewQte : double = -1);
    procedure ModificationPAMOY(Arow: Integer; NewQte : double = -1);
    procedure AffichageDetailValor(OU : Integer);
    procedure ModificationQte(Arow: integer; NewQte: double);
    procedure ModificationPRMOY(Arow: Integer; NewQte: double);
    procedure ModificationPVMOY(Arow: Integer; NewQte: double);
    procedure CreationTreeView;
    procedure CreeArboSimul(TOBSimulPres: TOB; Racine: HTreeNode);
    procedure DefinieListeSaisie (CurElt : integer);
    procedure ShowButton(Ou: integer);
    function GetRow: integer;
    procedure SetRow(Arow: integer);
    procedure Enabled(Zone: THNumEdit; State: boolean);
    procedure ModificationPVMOYTTC(Arow: Integer; NewQte: double);
    procedure PVMOYTTCChange(Sender: TObject);
    procedure SetDecimals;
    procedure GestionCoefFrais (TOBL : TOB);
  public
    { Déclarations publiques }
  end;

const MAXNiveauAff = 4;
var
  FSimulRentab: TFSimulRentab;

function Entree_Simulation (TOBArticles,TOBPiece,TOBOuvrage,TOBPorcs,TOBBases,TOBBasesL : TOB; Caption : string;Action:TactionFiche; DEV: Rdevise;Paragraphe : boolean=true; Cotraitance : boolean=False):boolean;

implementation
uses FactBordereau,NomenUtil,FactureBTP,UtilPgi,UFonctionsCBP;
{$R *.DFM}

function Entree_Simulation (TOBArticles,TOBPiece,TOBOuvrage,TOBPorcs,TOBBases,TOBBasesL : TOB; Caption : string;Action:TactionFiche; DEV: Rdevise;Paragraphe : boolean=true; Cotraitance : boolean=False):boolean;
var XX : TFSimulRenTab;
begin
  Result := false;
  XX := TFSimulRenTab.create(application);
  XX.TobArticles  := TOBArticles;
  XX.TobPiece     := TOBPiece;
  XX.TOBOuvrage   := TOBOuvrage;
  XX.TOBPorcs     := TOBPorcs;
  XX.TOBBases     := TOBBases;
  XX.TOBBasesL    := TOBBasesL;
  XX.Caption      := Caption;
  XX.Action       := Action;
  XX.paragraphe   := Paragraphe;
  XX.Cotraitance  := Cotraitance;
  XX.DEV          := DEV;
  XX.ShowModal;
  if XX.ValidateOk then Result := true;
  XX.Free;
end;

procedure TFSimulRentab.DefinieListeSaisie (CurElt : integer);
begin
  if Paragraphe then
  BEGIN
    GS_PARAG := CurElt;
    if EnHt then
    begin
    	ListeSaisie := 'NUMAFF;OPEN;LIBPARAGRA;LIBCATEGORIE;LIBNATURE;ARTICLE;LIBELLE;HEURES;MONTANTPA;COEFFG;COEFFC;COEFFR;MONTANTPR;COEFMARG;MONTANTPV';
    end else
    begin
    	ListeSaisie := 'NUMAFF;OPEN;LIBPARAGRA;LIBCATEGORIE;LIBNATURE;ARTICLE;LIBELLE;HEURES;MONTANTPA;COEFFG;COEFFC;COEFFR;MONTANTPR;COEFMARG;MONTANTPVTTC';
    end;
  END else
  BEGIN
    GS_PARAG := -1;
    if EnHt then
    begin
    	ListeSaisie := 'NUMAFF;OPEN;LIBCATEGORIE;LIBNATURE;ARTICLE;LIBELLE;HEURES;MONTANTPA;COEFFG;COEFFC;COEFFR;MONTANTPR;COEFMARG;MONTANTPV';
    end else
    begin
    	ListeSaisie := 'NUMAFF;OPEN;LIBCATEGORIE;LIBNATURE;ARTICLE;LIBELLE;HEURES;MONTANTPA;COEFFG;COEFFC;COEFFR;MONTANTPR;COEFMARG;MONTANTPVTTC';
    end
  END;
end;

procedure TFSimulRentab.definiGrilleSaisie ;
var CurElt : integer;
begin
  CurElt := 0;
  GS_NUMLIG := CurElt;
  inc(CurElt);
  GS_OPEN := CurElt;
  inc(CurElt);
  DefinieListeSaisie (CurElt);
  if GS_PARAG <> -1 then inc(CurElt);
  GS_CATEGORIE := CurElt;
  inc(CurElt);
  GS_NATURE := CurElt;
  inc(CurElt);
  GS_CODE := CurElt;
  inc(CurElt);
  GS_LIBELLE:=CurElt;
  inc(CurElt);
  GS_HRS := CurElt;
  inc(CurElt);
  GS_VALACHAT := CurElt;
  inc(CurElt);
  GS_COEFFG := CurElt;
  inc(CurElt);
  GS_COEFFC := CurElt;
  inc(CurElt);
  GS_COEFFR := CurElt;
  inc(CurElt);
  GS_VALPR := CurElt;
  inc(CurElt);
  GS_COEFMARG := CurElt;
  inc(CurElt);
  GS_VALPV := CurElt;
  inc(CurElt);
  //
  GS.colcount := CurElt;
  //
  GS.cells[GS_NUMLIG,0] := 'Num';
  GS.ColAligns[GS_NUMLIG]:=taRightJustify;
  GS.ColWidths [GS_NUMLIG] := 30;
  GS.ColLengths  [GS_NUMLIG] := 30;
  //
  GS.cells[GS_OPEN,0] := 'D';
  GS.ColWidths [GS_OPEN] := 18;
  GS.Collengths [GS_OPEN] := 18;
  //
  if GS_PARAG <> -1 then
     begin
     GS.cells [GS_PARAG,0] := 'Paragraphe';
     GS.ColWidths [GS_PARAG] := 300;
     GS.ColLengths [GS_PARAG] := 300;
     GS.ColAligns[GS_PARAG]:=taLeftJustify;
     end;
  //
  GS.Cells[GS_CATEGORIE,0] := 'Catégorie';
  GS.ColWidths [GS_CATEGORIE] := 100;
  GS.ColLengths [GS_CATEGORIE] := 100;
  GS.ColAligns[GS_CATEGORIE]:=taLeftJustify;
  //
  GS.Cells[GS_NATURE,0] := 'Nature';
  GS.ColWidths [GS_NATURE] := 200;
  GS.ColLengths [GS_NATURE] := 200;
  GS.ColAligns[GS_NATURE]:=taLeftJustify;
  //
  GS.Cells[GS_CODE,0] := 'Code';
  GS.ColWidths [GS_CODE] := 100;
  GS.ColLengths [GS_CODE] := 100;
  GS.ColAligns[GS_CODE]:=taLeftJustify;
  //
  GS.Cells[GS_LIBELLE,0] := 'Libellé';
  GS.ColWidths [GS_LIBELLE] := 300;
  GS.ColLengths [GS_LIBELLE] := 200;
  GS.ColAligns[GS_LIBELLE]:=taLeftJustify;
  //
  GS.Cells[GS_HRS,0] := 'Heures';
  GS.ColWidths [GS_HRS] := 90;
  GS.ColLengths [GS_HRS] := 90;
  GS.ColAligns[GS_HRS]:=taRightJustify;
  GS.ColFormats[GS_HRS]:='##0.000;##0.000; ;';
  GS.ColTypes[GS_HRS]:='R';
  //
  GS.Cells[GS_VALACHAT,0] := 'Achat';
  GS.ColWidths [GS_VALACHAT] := 90;
  GS.ColLengths [GS_VALACHAT] := 90;
  GS.ColAligns[GS_VALACHAT]:=taRightJustify;
  GS.ColFormats[GS_VALACHAT]:='# ##0.00;# ##0.00; ;';
  GS.ColTypes[GS_VALACHAT]:='R';
  //
  GS.Cells[GS_COEFFG,0] := 'Coef FG';
  GS.ColWidths [GS_COEFFG] := 80;
  GS.Collengths [GS_COEFFG] := 80;
  GS.ColAligns[GS_COEFFG]:=taRightJustify;
  GS.ColFormats[GS_COEFFG]:='#0.0000;#0.0000; ;';
  GS.ColTypes[GS_COEFFG]:='R';
  //
  GS.Cells[GS_COEFFC,0] := 'Coef FC';
  GS.ColWidths [GS_COEFFC] := 80;
  GS.Collengths [GS_COEFFC] := 80;
  GS.ColAligns[GS_COEFFC]:=taRightJustify;
  GS.ColFormats[GS_COEFFC]:='#0.0000;#0.0000; ;';
  GS.ColTypes[GS_COEFFC]:='R';
  //
  GS.Cells[GS_COEFFR,0] := 'Coef FR';
  GS.ColWidths [GS_COEFFR] := 80;
  GS.Collengths [GS_COEFFR] := 80;
  GS.ColAligns[GS_COEFFR]:=taRightJustify;
  GS.ColFormats[GS_COEFFR]:='#0.0000;#0.0000; ;';
  GS.ColTypes[GS_COEFFR]:='R';
  //
  GS.Cells[GS_VALPR,0] := 'Revient';
  GS.ColWidths [GS_VALPR] := 90;
  GS.ColLengths [GS_VALPR] := 90;
  GS.ColAligns[GS_VALPR]:=taRightJustify;
  GS.ColFormats[GS_VALPR]:='# ###.00;# ###.00; ;';
  GS.ColTypes[GS_VALPR]:='R';
  //
  GS.Cells[GS_COEFMARG,0] := 'Coef Marge';
  GS.ColWidths [GS_COEFMARG] := 80;
  GS.Collengths [GS_COEFMARG] := 80;
  GS.ColAligns[GS_COEFMARG]:=taRightJustify;
  GS.ColFormats[GS_COEFMARG]:='#0.0000;#0.0000; ;';
  GS.ColTypes[GS_COEFMARG]:='R';
  //
  GS.Cells[GS_VALPV,0] := 'Vente';
  GS.ColWidths [GS_VALPV] := 90;
  GS.ColLengths [GS_VALPV] := 90;
  GS.ColAligns[GS_VALPV]:=taRightJustify;
  GS.ColFormats[GS_VALPV]:='# ##0.00;# ##0.00; ;';
  GS.ColTypes[GS_VALPV]:='R';
  //
end;

procedure TFSimulRentab.DefiniNiveauVisible (Niveau : integer);
begin
GS.Collengths [GS_NUMLIG] := 30;
GS.Collengths [GS_OPEN] := 18;
if GS_PARAG <> -1 then GS.ColLengths [GS_PARAG] := 300;
GS.ColLengths [GS_CATEGORIE] := 100;
if Niveau > 2 then // a partie du niveau Nature
   begin
   GS.ColWidths [GS_NATURE] := 200;
   if Niveau > 3 then // a partir du niveau Articles
      BEGIN
      GS.ColWidths [GS_CODE] := 100;
      GS.ColWidths [GS_LIBELLE] := 300;
      END ELSE
      BEGIN
      GS.ColWidths [GS_CODE] := 0;
      GS.ColWidths [GS_LIBELLE] := 0;
      END;
   end else
   begin
   GS.ColWidths [GS_NATURE] := 0;
   GS.ColWidths [GS_CODE] := 0;
   GS.ColWidths [GS_LIBELLE] := 0;
   end;
GS.ColLengths [GS_HRS] := 90;
GS.ColLengths [GS_VALACHAT] := 90;
GS.Collengths [GS_COEFFG] := 80;
GS.Collengths [GS_COEFFC] := 80;
GS.Collengths [GS_COEFFR] := 80;
GS.ColLengths [GS_VALPR] := 90;
GS.Collengths [GS_COEFMARG] := 80;
GS.ColLengths [GS_VALPV] := 90;
end;

procedure TFSimulRentab.FormCreate(Sender: TObject);
begin
  TOBPieceFrais := TOB.Create ('LES FRAIS PIECE',nil,-1);
  AddChampsFrais (TOBPieceFrais);
  TOBSimulPres := TOB.create ('PRESENTATION',nil,-1);
  AddchampsupTSimul (TOBSimulPres);
  TOBSimulPres_O := TOB.create ('PRESENTATION',nil,-1);
  Niveau := 0;
  GS.PostDrawCell := GSPostDrawcell;
  GS.GetCellCanvas := GSGetCellCanvas;
  ValidateOk := false;
  IsModifie := false;
  WhantVisible := false;
  CoefMargGlobalModif := false;
end;

procedure TFSimulRentab.GSGetCellCanvas ( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) ;
var Niveau : integer;
    TOBL : TOB;
    ARect : TRect ;
begin
if ACol<GS.FixedCols then Exit ;
if Arow < GS.fixedRows then exit;
ARect:=GS.CellRect(ACol,ARow) ;
TOBL:=GetTOBLigne(TOBSimulPres,ARow) ; if TOBL=Nil then Exit ;
Canvas.Pen.Style:=psSolid ; Canvas.Pen.Color:=clgray ;
Niveau := TOBL.GetValue('NIVEAU');
if Niveau = 0 then // Totalisation
   BEGIN
   Canvas.Brush.Style:=bsSolid ;
   canvas.Brush.Color := clActiveCaption;  (* couleur de la caption active *)
   Canvas.Font.Style:=Canvas.Font.Style+[fsbold];
   canvas.font.color := clwhite;
   canvas.DrawFocusRect (Arect);
   END;
if Niveau = 1 then // paragraphe
   BEGIN
   Canvas.Brush.Style:=bsSolid ;
   canvas.Brush.Color := $F2C7B6; (* clRed; *)
   //Canvas.Font.Style:=Canvas.Font.Style+[fsItalic];
   canvas.font.Color := clblack;
   canvas.DrawFocusRect (Arect);
   END;
if Niveau = 2 then // categorie
   BEGIN
   if Acol >= GS_CATEGORIE Then
      BEGIN
      Canvas.Brush.Style:=bsSolid ;
      canvas.Brush.Color := clBtnFace;
      //Canvas.Font.Style:=Canvas.Font.Style+[fsItalic];
      canvas.font.color := clBlack;
      canvas.DrawFocusRect (Arect);
      END;
   END;
if Niveau = 3 then // Nature
   BEGIN
   if Acol >= GS_NATURE Then
      BEGIN
      Canvas.Brush.Style:=bsSolid ;
      canvas.Brush.Color := clInfoBk;
      //Canvas.Font.Style:=Canvas.Font.Style+[fsItalic];
      canvas.font.color := clBlack;
      canvas.DrawFocusRect (Arect);
      END;
   END;
if Niveau = 4 then // Article
   BEGIN
   if Acol >= GS_CODE Then
      BEGIN
      canvas.font.color := clBlack;
      END;
   END;
end;

procedure TFSimulRentab.FormClose(Sender: TObject;  var Action: TCloseAction);
begin
  TOBSimulPres.free;
  TOBSimulPres_O.free;
  TOBPieceFrais.free;
end;

procedure TFSimulRentab.SetDecimals;
begin
	QTe.Decimals := V_PGI.okdecQ;
	QTe.NumericType := ntDecimal;
	MPA.Decimals := V_PGI.okdecV;
	MPA.NumericType := ntDecimal;
	MPR.Decimals := V_PGI.okdecV;
	MPR.NumericType := ntDecimal;
	MPV.Decimals := V_PGI.okdecV;
	MPV.NumericType := ntDecimal;
	MPVTTC.Decimals := V_PGI.okdecV;
	MPVTTC.NumericType := ntDecimal;
	PAMOY.Decimals := V_PGI.okdecP;
	PAMOY.NumericType := ntDecimal;
	PRMOY.Decimals := V_PGI.okdecP;
	PRMOY.NumericType := ntDecimal;
	PVMOY.Decimals := V_PGI.okdecP;
	PVMOY.NumericType := ntDecimal;
	PVMOYTTC.Decimals := V_PGI.okdecP;
	PVMOYTTC.NumericType := ntDecimal;
end;

procedure TFSimulRentab.FormShow(Sender: TObject);
var cancel : boolean;
    Acol , Arow : integer;
    bBof : boolean;
begin
  EnHt :=  (TOBPiece.getValue('GP_FACTUREHT')='X');

  NbLigne := 0;
  Niveau := 2;

  definiGrilleSaisie;
  //
  TOBPieceFrais.PutValue ('MONTANTFC',GetMontantFraisDetail (TOBPiece, bBof));
  TOBPieceFrais.PutValue ('COEFFR',GetTauxFg (TOBporcs)/100);
  //
  CreationTOBSimul(TobSimulPres,TOBPIECE,TOBOuvrage,TOBArticles,Paragraphe,GestionInterdit,Cotraitance);
  CreationTreeView;
  TOBSimulPres.SetAllModifie (false);
  TOBSimulPres_O.dupliquer (TobSimulPres,true,true);
  IndiceGrilleInit (TOBSimulPres,Niveau,NbLigne);
  inc(nbligne);
  DefiniNiveauVisible (Niveau);
  AfficheLaGrille (GS,ListeSaisie,TobSimulPres,Niveau,NbLigne,Paragraphe);
  Arow := 1; Acol := 1; Cancel := false;
  GSRowEnter (self,GS.row,cancel,false);
  GSCellEnter (self,Acol,Arow,cancel);
  GS.row := Arow;
  GS.col := Acol;
  CellCur:=GS.Cells[GS.Col,GS.Row] ;
  affecteGrid(GS,Action);
  SetDecimals;
  //
end;

procedure TFSimulRentab.GSPostDrawcell ( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) ;
Var ARect : TRect ;
    // modif BTP
    TOBL : TOB;
    NumGraph : integer;
begin
ARect:=GS.CellRect(ACol,ARow) ;
GS.Canvas.Pen.Style:=psSolid ; GS.Canvas.Pen.Color:=clgray ;
GS.Canvas.Brush.Style:=BsSolid ;
TOBL:=GetTOBLigne(TOBSimulPres,ARow) ; if TOBL=Nil then Exit ;
if (Acol = GS_OPEN) and (Arow >= GS.fixedRows) then
   BEGIN
   Canvas.FillRect (ARect);
   if (TOBL.GetValue ('NIVEAU') >= MAXNiveauAff) or (TOBL.GetValue ('NIVEAU') = 0) then exit;
   NumGraph := RecupTypeGraph (TOBL);
   if NumGraph >= 0 then
      begin
      IListeDeploie.DrawingStyle := dsTransparent;
      IListeDeploie.Draw (CanVas,ARect.left,ARect.top,NumGraph);
      end;
   END;
if (Acol = GS_HRS) and (Arow >= GS.fixedRows) then
   BEGIN
   if TOBL.GetValue('HEURES') = 0 then Canvas.FillRect (ARect);
   END;

end;

procedure TFSimulRentab.GSCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
//
if Action=taConsult then Exit ;
ZoneSuivanteOuOk(ACol,ARow,Cancel) ;
if cancel then exit;
if Acol < GS_HRS then GS.CacheEdit else GS.ShowEditor ;
CellCur:=GS.Cells[GS.Col,GS.Row] ;
end;

procedure TFSimulRentab.GSCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
  if CellCur = GS.cells[Acol,Arow] then exit;
  if Acol < GS_HRS then GS.cells [Acol,Arow] := Cellcur else
  if (Acol = GS_HRS) then ModificationHrs (Arow) else
  if (Acol = GS_VALACHAT) then ModificationAchat (Arow);
  if (Acol = GS_COEFFG) then ModificationCoefFG (Arow) else
  if (Acol = GS_VALPR) then ModificationRevient (Arow) else
  if (Acol = GS_COEFMARG) then ModificationCoefMARG (Arow) else
  if (Acol = GS_VALPV) and (EnHt) then ModificationPV (Arow) else
  if (Acol = GS_VALPV) and (not EnHt) then ModificationPVTTC (Arow);
  CellCur := GS.cells[ACol,Arow];
end;

procedure TFSimulRentab.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var TOBL : TOB;
begin
TOBL := GetTOBLigne (TOBSimulPres,Ou);
if TOBL = nil then exit;
CurrentTOb := TOBL;
if TOBL.GetValue('NIVEAU') = 0 then exit;
AffichageDetailValor (Ou);
if TOBL.GetValue('NIVEAU') >= MAXNiveauAff then exit;
ShowButton (Ou);
(* il etait laaa avant
ARect:=GS.CellRect(GS_OPEN,Ou) ;
Baction.ImageIndex := RecupTypeGraph (TOBL);
BAction.Opaque := false;
With BAction do
    Begin
    Top := Arect.top;
    Left := Arect.Left;
    Width := Arect.Right - Arect.Left ;
    Height  := Arect.Bottom - Arect.Top;
    Parent := GS;
    Visible := true;
    end;
*)
end;

procedure TFSimulRentab.ShowButton (Ou : integer);
var Arect : Trect;
		TOBL : TOB;
begin
TOBL := GetTOBLigne (TOBSimulPres,Ou);
if TOBL.GetValue('NIVEAU') = 0 then exit;
if TOBL.GetValue('NIVEAU') >= MAXNiveauAff then exit;
if TOBL = nil then exit;
CurrentTOb := TOBL;
ARect:=GS.CellRect(GS_OPEN,Ou) ;
Baction.ImageIndex := RecupTypeGraph (TOBL);
//BAction.Opaque := false;
With BAction do
    Begin
    Top := Arect.top;
    Left := Arect.Left;
    Width := Arect.Right - Arect.Left ;
    Height  := Arect.Bottom - Arect.Top;
    Parent := GS;
    Visible := true;
    end;
end;

procedure TFSimulRentab.GSRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
BAction.Visible := false;
end;


procedure TFSimulRentab.ZoneSuivanteOuOk ( Var ACol,ARow : Longint ; Var Cancel : boolean ) ;
Var Sens,ii,Lim : integer ;
    OldEna,ChgLig,ChgSens  : boolean ;
BEGIN
//Lim := GS.rowCount-1;
OldEna:=GS.SynEnabled ; GS.SynEnabled:=False ;
Sens:=-1 ; ChgLig:=(GS.Row<>ARow) ; ChgSens:=False ;
if GS.Row>ARow then Sens:=1 else if ((GS.Row=ARow) and (ACol<=GS.Col)) then Sens:=1 ;
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
      if ChgLig then BEGIN ACol:=GS.FixedCols-1 ; ChgLig:=False ; END ;
      if ACol<GS.ColCount-1 then Inc(ACol) else BEGIN Inc(ARow) ; ACol:=GS.FixedCols ; END ;
      END else
      BEGIN
      if ((ACol=GS.FixedCols) and (ARow=1)) then
         BEGIN
         if ChgSens then Break else BEGIN Sens:=1 ; Continue ; END ;
         END ;
      if ChgLig then BEGIN ACol:=GS.FixedCols ; ChgLig:=False ; END ;
      if ACol>GS.FixedCols then
      BEGIN
      	Dec(ACol);
        if Acol < 0 then Acol := GS.FixedCols;
      end else BEGIN Dec(ARow) ; ACol:=GS.Fixedcols ; END ;
      END ;
   END ;
GS.SynEnabled:=OldEna ;
if Arow < 0 then Arow := GS.fixedRows;
END ;

Function TFSimulRentab.ZoneAccessible ( ACol,ARow : Longint) : boolean ;
Var TOBL : TOB ;
BEGIN
  Result:=True ;

  if GS.ColLengths [Acol] = 0 then BEGIN Result := false;exit;END;

  TOBL := GetTOBLigne (TOBSimulPres,Arow);  if TOBL = nil then exit;

  if Acol = GS_OPEN then BEGIN REsult := false; Exit; END;

  //FV1 : 30/10/2013 -FS#731 - POUCHAIN : simulation de rentabilité, ajout autorisation d'accès au coef de FG.
  //if (Acol=GS_COEFFG) or (Acol=GS_COEFFR) or (Acol=GS_COEFFC) then BEGIN result := false; Exit; END;
  if not (ExJaiLeDroitConcept(TConcept(bt517),False)) then
  begin
    if (Acol=GS_COEFFG) Then
    begin
      Result := false;
      exit;
    end;
  end;

  if (Acol=GS_COEFFR) or (Acol=GS_COEFFC) then BEGIN result := false; Exit; END;
  //if Acol = GS_VALACHAT then BEGIN Result := false;Exit; END;

  if (Acol = GS_HRS) Then
     begin
     if TOBL.GetValue('TYPE') = 'F' then BEGIN Result := false;Exit; END;
     end;
  if ((Acol = GS_VALPV) or (Acol= GS_COEFMARG)) and (TOBL.GetValue('BLOQUEPV')='X') then BEGIN Result := false; exit; end;
  if (Acol = GS_VALPR) and  (TOBL.GetValue('BLOQUEPR')='X') then BEGIN Result := false; exit; end;

END ;

procedure TFSimulRentab.BActionClick(Sender: TObject);
var TOBL : TOB;
    found : boolean;
    Arow : integer;
begin
found := false;
(* -- AVANT
NbLigne := GS.row;
Arow := GS.row;
*)
NbLigne := GetRow;
Arow := GetRow;
// --
TOBL:=GetTOBLigne(TOBSimulPres,Arow) ; if TOBL=Nil then Exit ;
if (TOBL.GetValue('OPEN')='-') and (TOBL.detail.count > 0)  then
   BEGIN
   if TOBL.GetValue ('NIVEAU') >= MAXNiveauAff then exit;
   OuvreBranche (GS,TobSimulPres,NbLigne,Niveau,found,false);
   TOBL.PutValue('OPEN','X');
   END else
   if (TOBL.GetValue('OPEN')='X') and (TOBL.detail.count > 0)  then
      BEGIN
      Niveau := 1;
      FermeBranche (GS,TobSimulPres,NbLigne,niveau,found);
      TOBL.PutValue('OPEN','-');
      END;
inc(Nbligne);
DefiniNiveauVisible (Niveau);
AfficheLaGrille (GS,ListeSaisie,TobSimulPres,Niveau,NbLigne,Paragraphe);
PositionneLigne (Arow);
end;

procedure TFSimulRentab.PositionneLigne (Arow:integer;Acol : integer=-1);
var GRow,GCol : integer;
    Cancel : boolean;
    Synchro : boolean;
begin
Synchro := GS.synenabled;
GS.synEnabled := false;
Grow := Arow;
if Acol <> -1 then GCol := Acol else Gcol := GS.Fixedcols ;
GS.col := Gcol;
GS.row:= Grow;
cancel := false;
GS.CacheEdit;
GSRowenter (Self,Grow,cancel,false);
GSCellenter (Self,Gcol,Grow,Cancel);
GS.col := Gcol;
GS.row:= Grow;
CellCur := GS.cells[Gcol,Grow];
GS.SynEnabled := synchro;
GS.MontreEdit;
end;

procedure TFSimulRentab.DesactiveEventGS;
begin
GS.OnCellEnter := nil;
GS.OnCellExit := nil;
GS.OnRowEnter := nil;
GS.OnRowExit := nil;
end;

procedure TFSimulRentab.ActiveEventGS;
begin
GS.OnCellEnter := GSCellEnter;
GS.OnCellExit := GSCellExit;
GS.OnRowEnter := GSRowEnter;
GS.OnRowExit := GSRowExit;
GS.OnMouseMove := GSMouseMove;
end;

procedure TFSimulRentab.ModificationHrs (Arow : integer; NewQte : double);
var TOBL : TOB;
    NewHrs,AncHrs: double;
    SavRow,SavCol : integer;
begin
SavRow := GS.Row;
SavCol := GS.col;
TOBL := GetTOBLigne (TOBSimulPres,Arow); if TOBL = nil then exit;
if NewQte = -1 then NewHrs := valeur(GS.cells[GS_HRS,Arow]) else NewHrs := NewQte;
AncHrs := TOBL.GetValue('HEURES');
ReinitValeur (TscHrs,TOBSimulPres,gestionInterdit);
changeValue(TscHrs,AncHrs,NewHrs,TOBL,gestionInterdit);
//
GestionCoefFrais (TOBL);
//
FinaliseCalcul (TOBSimulPres);
DesactiveEventGS;
AfficheLaGrille (GS,ListeSaisie,TobSimulPres,Niveau,NbLigne,Paragraphe);
GS.row := SavRow;
GS.col := SavCol;
ActiveEventGS;
IsModifie := true;
AffichageDetailValor (Savrow);
CellCur := GS.cells[SAvCol,SAvrow];
end;

procedure TFSimulRentab.ModificationQte (Arow : integer;NewQte : double);
var TOBL : TOB;
    AncHrs: double;
    SavRow,SavCol : integer;
begin
SavRow := GS.Row;
SavCol := GS.col;
TOBL := GetTOBLigne (TOBSimulPres,Arow); if TOBL = nil then exit;
AncHrs := TOBL.GetValue('QTE');
ReinitValeur (TscHrs,TOBSimulPres,gestionInterdit);
changeValue(TscQte,AncHrs,NewQte,TOBL,gestionInterdit);
FinaliseCalcul (TOBSimulPres);
DesactiveEventGS;
AfficheLaGrille (GS,ListeSaisie,TobSimulPres,Niveau,NbLigne,Paragraphe);
GS.row := SavRow;
GS.col := SavCol;
ActiveEventGS;
IsModifie := true;
AffichageDetailValor (SavRow);
end;

procedure TFSimulRentab.ModificationAchat  (Arow : integer; NewQte : double = -1);
var TOBL : TOB;
    NewHrs,AncHrs: double;
    SavRow,SavCol : integer;
begin
SavRow := GS.Row;
SavCol := GS.col;
TOBL := GetTOBLigne (TOBSimulPres,Arow); if TOBL = nil then exit;
if NewQte = -1 then NewHrs := valeur(GS.cells[GS_VALACHAT,Arow]) else NewHrs := NewQte;
AncHrs := TOBL.GetValue('MONTANTPA');
ReinitValeur (TscPa,TOBSimulPres,gestionInterdit);
changeValue(TscPa,AncHrs,NewHrs,TOBL,gestionInterdit);
GestionCoefFrais (TOBL);
FinaliseCalcul (TOBSimulPres);
DesactiveEventGS;
AfficheLaGrille (GS,ListeSaisie,TobSimulPres,Niveau,NbLigne,Paragraphe);
GS.row := SavRow;
GS.col := SavCol;
ActiveEventGS;
IsModifie := true;
AffichageDetailValor (SavROw);
CellCur := GS.cells[SAvCol,SAvrow];
end;

procedure TFSimulRentab.ModificationCoefFG (Arow : Integer; NewQte : double);
var TOBL : TOB;
    NewHrs,AncHrs: double;
    SavRow,SavCol : integer;
begin
SavRow := GS.Row;
SavCol := GS.col;
// phase 1
TOBL := GetTOBLigne (TOBSimulPres,Arow); if TOBL = nil then exit;
if NewQte = -1 then NewHrs := valeur(GS.cells[GS_COEFFG,Arow]) else NewHrs := NewQte;
AncHrs := TOBL.GetValue('COEFFG');
ReinitValeur (TscCoefFG,TOBSimulPres,gestionInterdit);
changeValue(TscCoefFG,AncHrs,NewHrs,TOBL,gestionInterdit);
//
GestionCoefFrais (TOBL);
//
FinaliseCalcul (TOBSimulPres);
DesactiveEventGS;
AfficheLaGrille (GS,ListeSaisie,TobSimulPres,Niveau,NbLigne,Paragraphe);
GS.row := SavRow;
GS.col := SavCol;
ActiveEventGS;
IsModifie := true;
AffichageDetailValor (SAvRow);
CellCur := GS.cells[SavCol,Savrow];
end;

procedure TFSimulRentab.ModificationCoefMARG (Arow: Integer; NewQte : double = -1);
var TOBL : TOB;
    NewHrs,AncHrs: double;
    SavRow,SavCol : integer;
    Niveau : Integer;
begin
SavRow := GS.Row;
SavCol := GS.col;
TOBL := GetTOBLigne (TOBSimulPres,Arow); if TOBL = nil then exit;
if NewQte = -1 then NewHrs := valeur(GS.cells[GS_COEFMARG,Arow]) else NewHrs := NewQte ;
Niveau := TOBL.GetValue('NIVEAU');
AncHrs := TOBL.GetValue('COEFMARG');
ReinitValeur (TscCoefMA,TOBSimulPres,gestionInterdit);
changeValue(TscCoefMA,AncHrs,NewHrs,TOBL,gestionInterdit,(Niveau=0));
if Niveau=0 then
begin
	CoefMargGlobalModif := True;
  ValCoefMarg := NewHrs;
end;
FinaliseCalcul (TOBSimulPres);
DesactiveEventGS;
AfficheLaGrille (GS,ListeSaisie,TobSimulPres,Niveau,NbLigne,Paragraphe);
GS.row := SavRow;
GS.col := SavCol;
ActiveEventGS;
IsModifie := true;
AffichageDetailValor (SavRow);
end;

procedure TFSimulRentab.ModificationRevient(Arow: Integer;NewQte : double = -1);
var TOBL : TOB;
    NewHrs,AncHrs,PA: double;
    SavRow,SavCol : integer;
begin
SavRow := GS.Row;
SavCol := GS.col;
TOBL := GetTOBLigne (TOBSimulPres,Arow); if TOBL = nil then exit;
if NewQte = -1 then NewHrs := valeur(GS.cells[GS_VALPR,Arow]) else NewHrs := NewQte;
AncHrs := TOBL.GetValue('COEFFG');
PA := TOBL.GetValue('MONTANTPA'); NewHrs := Arrondi(NewHrs/Pa,4);
ReinitValeur (TscCoefFG,TOBSimulPres,gestionInterdit);
changeValue(TscCoefFG,AncHrs,NewHrs,TOBL,gestionInterdit);
FinaliseCalcul (TOBSimulPres);
DesactiveEventGS;
AfficheLaGrille (GS,ListeSaisie,TobSimulPres,Niveau,NbLigne,Paragraphe);
GS.row := SavRow;
GS.col := SavCol;
ActiveEventGS;
IsModifie := true;
AffichageDetailValor (SavRow);
CellCur := GS.cells[SAvCol,SAvrow];
end;


procedure TFSimulRentab.ModificationPV (Arow : Integer; NewQte : double );
var TOBL : TOB;
    NewHrs,{AncHrs,}VALPR: double;
    SavRow,SavCol : integer;
begin
SavRow := GS.Row;
SavCol := GS.col;
TOBL := GetTOBLigne (TOBSimulPres,Arow); if TOBL = nil then exit;
//
if NewQte = -1 then NewHrs := valeur(GS.cells[GS_VALPV,Arow]) else NewHrs := NewQte;
ValPr := TOBL.GetValue('MONTANTPR'); NewHrs := arrondi(NewHrs/ValPr,4); //ANcHrs := 1;
//AncHrs := TOBL.GetValue('COEFMARG');
//
ReinitValeur (TscCoefMA,TOBSimulPres,gestionInterdit); // c'est normal ..laisser le coefma..svp
changeValue(TscCoefMA,{AncHrs}TOBL.GetValue('COEFMARG'),NewHrs,TOBL,gestionInterdit);
//
//
FinaliseCalcul (TOBSimulPres);
DesactiveEventGS;
AfficheLaGrille (GS,ListeSaisie,TobSimulPres,Niveau,NbLigne,Paragraphe);
GS.row := SavRow;
GS.col := SavCol;
ActiveEventGS;
IsModifie := true;
AffichageDetailValor (SavRow);
end;

procedure TFSimulRentab.ModificationPVTTC(Arow: Integer; NewQte: double);
var TOBL : TOB;
    NewVal,AncVal: double;
    SavRow,SavCol : integer;
begin
SavRow := GS.Row;
SavCol := GS.col;
TOBL := GetTOBLigne (TOBSimulPres,Arow); if TOBL = nil then exit;
//
	if NewQte = -1 then NewVal := valeur(GS.cells[GS_VALPV,Arow]) else NewVal := NewQte;
AncVal := TOBL.GetValue('MONTANTPVTTC');
//
ReinitValeur (TscPVTTC,TOBSimulPres,gestionInterdit);
changeValue(TscPVTTC,AncVal,NewVal,TOBL,gestionInterdit);
//
//
FinaliseCalcul (TOBSimulPres);
DesactiveEventGS;
AfficheLaGrille (GS,ListeSaisie,TobSimulPres,Niveau,NbLigne,Paragraphe);
GS.row := SavRow;
GS.col := SavCol;
ActiveEventGS;
IsModifie := true;
AffichageDetailValor (SavRow);
end;

procedure TFSimulRentab.BAbandonClick(Sender: TObject);
begin
close;
end;

procedure TFSimulRentab.BValiderClick(Sender: TObject);
var Indice : Integer;
    TOBS,TOBL  : TOB;
    TOBMOD: TOB;
begin
  TOBMOd := TOB.Create ('LES LIGNES MODIF',nil,-1);
  ValidateOk := True;
  if CoefMargGlobalModif then
  begin
    TOBPiece.PutValue('GP_RECALCULER','X');
		for Indice := 0 to TOBPiece.detail.count -1 do
    begin
			TOBL := TOBPiece.detail[Indice];
  		if TOBL.getValue('GL_TYPELIGNE')<>'ART' Then Continue;
      if TOBL.getValue('GL_BLOQUETARIF')<> 'X' then
      begin
        TOBL.PutValue('GL_COEFMARG',ValCoefMarg);
        TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));
        if IsOuvrage(TOBL) then
        begin
        	AppliqueChangeCoefMargOuvrage(TOBL,TOBOuvrage,DEV);
          RecalculeOuvrage (TOBPiece,TOBL,TOBOuvrage,TOBBases,TOBBasesL,DEV) ;
        end else
        begin
          if TOBL.GetValue('GL_DPR') <> 0 then
          begin
            TOBL.PutValue('GL_PUHT', Arrondi(TOBL.GetValue('GL_DPR') * TOBL.GetValue('GL_COEFMARG'),V_PGI.OkdecP));
            TOBL.PutValue('GL_PUHTDEV',pivottodevise(TobL.GetValue('GL_PUHT'),DEV.Taux,DEV.quotite,V_PGI.okdecP ));
          end;
        end;
        if TOBL.GetValue('GL_PUHT') <> 0 then
        begin
          TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('GL_PUHT')- TOBL.GetValue('GL_DPR'))/TOBL.GetValue('GL_PUHT'))*100,2));
        end else
        begin
          TOBL.PutValue('POURCENTMARQ',0);
        end;
        TOBL.PutValue('GL_RECALCULER','X');
      end;
    end;
  end;
  for Indice := 0 to TOBSimulPres.detail.count -1 do
  begin
    TOBS := TOBSimulPres.detail[Indice];
    TraiteModification (TOBS,TOBPIECE,TOBOuvrage,TOBMOD,DEV,gestionInterdit);
  end;
  RecalcSuiteModif (TOBPiece,TOBOuvrage,TOBMod,TOBBases,TOBBAsesL,DEV);
  if (TOBMod.detail.count = 0 ) and (not CoefMargGlobalModif) then ValidateOk := False;
  TOBMod.free;
  close;
end;


procedure TFSimulRentab.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
if Validateok then Exit;
if IsModifie then
   begin
   if PGIAskAF ('Etes-vous sur de vouloir annuler les modifications ?',caption) <> mryes then
      begin
      canclose := false;
      end;
   end;
end;

procedure TFSimulRentab.BDetailValClick(Sender: TObject);
var TOBL : TOB;
begin
TOBL := GetTOBLigne (TOBSimulPres,GS.row);
if TOBL = nil then exit;
CurrentTOb := TOBL;
if WhantVisible then BEGIN WhantVisible := false; TdetailVal.Visible := false; BdetailVal.Down := false; Exit; END
                else BEGIN WhantVisible := True; BdetailVal.Down := True;END;
if CurrentTOb.GetValue('NIVEAU')=4 then TdetailVal.visible := true;
end;

procedure TFSimulRentab.Enabled (Zone : THNumEdit ; State : boolean);
begin
  Zone.enabled := State;
  if State = true then
  begin
    Zone.color := clWindow;
    Zone.Font.Style := Zone.Font.Style -[fsBold];
  end else
  begin
    Zone.color := clBtnFace;
    Zone.Font.Style := Zone.Font.Style +[fsBold];
  end;
end;

procedure TFSimulRentab.AffichageDetailValor (Ou : Integer);
var TOBL : TOB;
begin

  if Ou < 1 then exit;

  TOBL := GetTOBLigne (TOBSimulPres,Ou);
  //
  if TOBL = nil then exit;

  if TOBL.GetValue('NIVEAU') <> 4 then
  BEGIN
    TdetailVal.Visible := False;
    Exit;
  END;

  TDetailVal.Caption := 'Détail des valorisations '+TOBL.GetValue('LIBELLE');

  if TOBL.GetValue('TYPE') = 'F' then
  BEGIN
    LHeures.Caption := 'Quantités';
    Qte.Value := TOBL.GetValue('QTE');
    Mpa.Value := TOBL.GetVAlue('MONTANTPAQ');
    Mpr.Value := TOBL.GetVAlue('MONTANTPRQ');
    Mpv.Value := TOBL.GetVAlue('MONTANTPVQ'); if not EnHT then Enabled(MPV,false) else Enabled(MPV,true);
    MpvTTC.Value := TOBL.GetVAlue('MONTANTPVTTCQ'); if EnHT then Enabled(MpvTTC,false) else Enabled(MpvTTC,true);
    CoeffG.Value := TOBL.GetValue('COEFFGQ');
    CoeffC.Value := TOBL.GetValue('COEFFCQ');
    CoeffR.Value := TOBL.GetValue('COEFFRQ');
    CoefMarg.Value := TOBL.GetValue('COEFMARGQ');
    PAMOY.Value := TOBL.GetValue('PAMOYENQ');
    PRMOY.Value := TOBL.GetValue('PRMOYENQ');
    PVMOY.Value := TOBL.GetValue('PVMOYENQ'); if not EnHT then Enabled(PVMOY,false) else Enabled(PVMOY,True);
    PVMOYTTC.Value := TOBL.GetValue('PVMOYENTTCQ'); if EnHT then Enabled(PVMOYTTC,false) else Enabled(PVMOYTTC,true);
  END
  ELSE
  BEGIN
    LHeures.Caption := 'Heures';
    Qte.Value := TOBL.GetValue('HEURES');
    Mpa.Value := TOBL.GetVAlue('MONTANTPAH');
    Mpr.Value := TOBL.GetVAlue('MONTANTPRH');
    Mpv.Value := TOBL.GetVAlue('MONTANTPVH'); if not EnHT then Enabled(MPV,false) else Enabled(MPV,true);
    MpvTTC.Value := TOBL.GetVAlue('MONTANTPVTTCH'); if EnHT then Enabled(MpvTTC,false) else Enabled(MpvTTC,true);
    CoeffG.Value := TOBL.GetValue('COEFFGH');
    CoeffC.Value := TOBL.GetValue('COEFFCH');
    CoeffR.Value := TOBL.GetValue('COEFFRH');
    CoefMarg.Value := TOBL.GetValue('COEFMARGH');
    PAMOY.Value := TOBL.GetValue('PAMOYENH');
    PRMOY.Value := TOBL.GetValue('PRMOYENH');
    PVMOY.Value := TOBL.GetValue('PVMOYENH');  if not EnHT then Enabled(PVMOY,false) else Enabled(PVMOY,True);
    PVMOYTTC.Value := TOBL.GetValue('PVMOYENTTCH'); if EnHT then Enabled(PVMOYTTC,false) else Enabled(PVMOYTTC,true);
  END;

  Enabled(COEFFG,false);
  Enabled(COEFFR,false);
  Enabled(COEFFC,false);
  //
//
	if TOBL.GetValue('BLOQUEPR')='X' then
  begin
  	Enabled(MPR,false);
  	Enabled(PRMOY,false);
  end else
  begin
  	Enabled(MPR,true);
  	Enabled(PRMOY,true);
  end;

	if TOBL.GetValue('BLOQUEPV')='X' then
  begin
  	Enabled(COEFMARG,false);
  	Enabled(MPV,false);
  	Enabled(MPVTTC,false);
  	Enabled(PVMOY,false);
  	Enabled(PVMOYTTC,false);
  end else
  begin
  	Enabled(COEFMARG,true);
  	Enabled(MPV,true);
  	Enabled(MPVTTC,true);
  	Enabled(PVMOY,true);
  	Enabled(PVMOYTTC,true);
  end;
  
  //FV1 : 30/10/2013 -FS#731 - POUCHAIN : simulation de rentabilité, ajout autorisation d'accès au coef de FG.
  if (ExJaiLeDroitConcept(TConcept(bt517),False)) then
    Enabled(COEFFG,True)
  else
    Enabled(COEFFG,False);
  //
  if not WhantVisible then exit;

  TDetailVal.Visible := true;

end;


procedure TFSimulRentab.TDetailValClose(Sender: TObject);
begin
WhantVisible := false; BdetailVal.Down := false;
end;

procedure TFSimulRentab.QTEChange(Sender: TObject);
begin
if THEQTE = QTE.Value then exit;
if CurrentTOb.GetValue('TYPE')='H' Then ModificationHrs (CurrentTob.GetValue('NUMAFF'),Qte.Value)
                                   Else ModificationQTe (CurrentTob.GetValue('NUMAFF'),Qte.Value);
AffichageDetailValor (GS.row);
end;

procedure TFSimulRentab.MPAChange(Sender: TObject);
var TOBL : TOB;
begin
if TheMPA = MPA.Value then exit;
TOBL := GetTOBLigne (TOBSimulPres,GS.row); if TOBL = nil then exit;
ModificationAchat (TOBL.GetValue('NUMAFF'),Mpa.Value);
AffichageDetailValor (GS.Row);
end;

procedure TFSimulRentab.MPRChange(Sender: TObject);
var TOBL : TOB;
begin
if TheMPR = MPR.Value then exit;
TOBL := GetTOBLigne (TOBSimulPres,GS.row); if TOBL = nil then exit;
ModificationRevient (TOBL.GetValue('NUMAFF'),MPR.Value);
AffichageDetailValor (GS.Row);
end;

procedure TFSimulRentab.MPVChange(Sender: TObject);
var TOBL : TOB;
begin
if TheMPV = MPV.Value then exit;
TOBL := GetTOBLigne (TOBSimulPres,GS.row); if TOBL = nil then exit;
ModificationPV (TOBL.GetValue('NUMAFF'),MPV.Value);
AffichageDetailValor (GS.Row);
end;

procedure TFSimulRentab.MPVTTCExit(Sender: TObject);
var TOBL : TOB;
begin
if TheMPVTTC = MPVTTC.Value then exit;
TOBL := GetTOBLigne (TOBSimulPres,GS.row); if TOBL = nil then exit;
ModificationPVTTC (TOBL.GetValue('NUMAFF'),MPVTTC.Value);
AffichageDetailValor (GS.Row);
end;

procedure TFSimulRentab.COEFFGChange(Sender: TObject);
var TOBL : TOB;
begin
if TheCOEFFG = COEFFG.Value then exit;
TOBL := GetTOBLigne (TOBSimulPres,GS.row); if TOBL = nil then exit;
ModificationCoefFG (TOBL.GetValue('NUMAFF'),COEFFG.Value);
AffichageDetailValor (GS.Row);
end;

procedure TFSimulRentab.COEFMARGChange(Sender: TObject);
var TOBL : TOB;
begin
if TheCOEFMARG = COEFMARG.Value then exit;
TOBL := GetTOBLigne (TOBSimulPres,GS.row); if TOBL = nil then exit;
ModificationCoefMARG (TOBL.GetValue('NUMAFF'),COEFMARG.Value);
AffichageDetailValor (GS.Row);
end;

procedure TFSimulRentab.PAMOYChange(Sender: TObject);
var TOBL : TOB;
begin
if ThePAMOY = PAMOY.Value then exit;
TOBL := GetTOBLigne (TOBSimulPres,GS.row); if TOBL = nil then exit;
ModificationPAMOY (TOBL.GetValue('NUMAFF'),PAMOY.Value);
AffichageDetailValor (GS.Row);
end;

procedure TFSimulRentab.PRMOYChange(Sender: TObject);
var TOBL : TOB;
begin
if ThePRMOY = PRMOY.Value then exit;
TOBL := GetTOBLigne (TOBSimulPres,GS.row); if TOBL = nil then exit;
ModificationPRMOY (TOBL.GetValue('NUMAFF'),PRMOY.Value);
AffichageDetailValor (GS.Row);
end;

procedure TFSimulRentab.PVMOYChange(Sender: TObject);
var TOBL : TOB;
begin
if ThePVMOY = PVMOY.Value then exit;
TOBL := GetTOBLigne (TOBSimulPres,GS.row); if TOBL = nil then exit;
ModificationPVMOY (TOBL.GetValue('NUMAFF'),PVMOY.Value);
AffichageDetailValor (GS.Row);
end;

procedure TFSimulRentab.PVMOYTTCChange(Sender: TObject);
var TOBL : TOB;
begin
if ThePVMOYTTC = PVMOYTTC.Value then exit;
TOBL := GetTOBLigne (TOBSimulPres,GS.row); if TOBL = nil then exit;
ModificationPVMOYTTC (TOBL.GetValue('NUMAFF'),PVMOYTTC.Value);
AffichageDetailValor (GS.Row);
end;

procedure TFSimulRentab.PVMOYTTCExit(Sender: TObject);
var TOBL : TOB;
begin
PVMOYTTCChange (self);
end;

procedure TFSimulRentab.ModificationPAMOY(Arow: Integer; NewQte: double);
var TOBL : TOB;
    SavRow,SavCol : integer;
begin
SavRow := GS.Row;
SavCol := GS.col;
TOBL := GetTOBLigne (TOBSimulPres,Arow); if TOBL = nil then exit;
ReinitValeur (TscPUA,TOBSimulPres,gestionInterdit);
changeValue(TscPUA,0,NewQte,TOBL,gestionInterdit);
//
GestionCoefFrais (TOBL);
//
FinaliseCalcul (TOBSimulPres);
DesactiveEventGS;
AfficheLaGrille (GS,ListeSaisie,TobSimulPres,Niveau,NbLigne,Paragraphe);
GS.row := SavRow;
GS.col := SavCol;
ActiveEventGS;
IsModifie := true;
end;

procedure TFSimulRentab.ModificationPRMOY(Arow: Integer; NewQte: double);
var TOBL : TOB;
    SavRow,SavCol : integer;
begin
SavRow := GS.Row;
SavCol := GS.col;
TOBL := GetTOBLigne (TOBSimulPres,Arow); if TOBL = nil then exit;
ReinitValeur (TscPUR,TOBSimulPres,gestionInterdit);
changeValue(TscPUR,0,NewQte,TOBL,gestionInterdit);
FinaliseCalcul (TOBSimulPres);
DesactiveEventGS;
AfficheLaGrille (GS,ListeSaisie,TobSimulPres,Niveau,NbLigne,Paragraphe);
GS.row := SavRow;
GS.col := SavCol;
ActiveEventGS;
IsModifie := true;
end;

procedure TFSimulRentab.ModificationPVMOYTTC(Arow: Integer; NewQte: double);
var TOBL : TOB;
    SavRow,SavCol : integer;
begin
SavRow := GS.Row;
SavCol := GS.col;
TOBL := GetTOBLigne (TOBSimulPres,Arow); if TOBL = nil then exit;
ReinitValeur (TscPUVTTC,TOBSimulPres,gestionInterdit);
changeValue(TscPUVTTC,0,NewQte,TOBL,gestionInterdit);
FinaliseCalcul (TOBSimulPres);
DesactiveEventGS;
AfficheLaGrille (GS,ListeSaisie,TobSimulPres,Niveau,NbLigne,Paragraphe);
GS.row := SavRow;
GS.col := SavCol;
ActiveEventGS;
IsModifie := true;
end;

procedure TFSimulRentab.ModificationPVMOY(Arow: Integer; NewQte: double);
var TOBL : TOB;
    SavRow,SavCol : integer;
begin
SavRow := GS.Row;
SavCol := GS.col;
TOBL := GetTOBLigne (TOBSimulPres,Arow); if TOBL = nil then exit;
ReinitValeur (TscPUV,TOBSimulPres,gestionInterdit);
changeValue(TscPUV,0,NewQte,TOBL,gestionInterdit);
FinaliseCalcul (TOBSimulPres);
DesactiveEventGS;
AfficheLaGrille (GS,ListeSaisie,TobSimulPres,Niveau,NbLigne,Paragraphe);
GS.row := SavRow;
GS.col := SavCol;
ActiveEventGS;
IsModifie := true;
end;

procedure TFSimulRentab.ApplicPAHClick(Sender: TObject);
begin
PAMOYChange (self);
end;

procedure TFSimulRentab.ApplicPRHClick(Sender: TObject);
begin
PRMOYChange (self);
end;

procedure TFSimulRentab.ApplicPVTTCClick(Sender: TObject);
begin
PVMOYTTCChange (self);
end;

procedure TFSimulRentab.CreationTreeView;
var Racine : HTreeNode ;
    Titre : HSTring ;
begin
TV.items.clear ;
Titre := 'Arborescence de l''étude de rentabilité';
Racine := TV.items.add(Nil, Titre);
CreeArboSimul (TOBSimulPres,Racine);
end;

procedure TFSimulRentab.CreeArboSimul (TOBSimulPres : TOB;Racine : HTreeNode);
var TOBL : TOB;
    Titre : hSTring ;
    Fille : HTreeNode ;
    Indice,NIVEAU : integer;
begin
for Indice := 0 to TOBSimulPres.detail.count -1 do
    begin
    TOBL := TOBSimulPres.detail[Indice];
    Niveau := TOBL.GetValue('NIVEAU');
    if Niveau > 4 then break;
    if Niveau = 1 then Titre := TOBL.GetValue('LIBPARAGRA');
    if Niveau = 2 then Titre := TOBL.GetValue('LIBCATEGORIE');
    if Niveau = 3 then Titre := TOBL.GetValue('LIBNATURE');
    if Niveau = 4 then Titre := TOBL.GetValue('LIBELLE');
    Fille := TV.Items.AddChild(Racine, Titre);
    Fille.Data := TOBL;
    if TOBL.detail.count > 0 then CreeArboSimul (TOBL,Fille);
    end;
end;

procedure TFSimulRentab.BarboClick(Sender: TObject);
begin
TTarborescence.Visible := true; Barbo.Down := true;
end;

procedure TFSimulRentab.TTarborescenceClose(Sender: TObject);
begin
Barbo.down := false;
end;

procedure TFSimulRentab.TVClick(Sender: TObject);
var Tn : TTreeNode;
    TOBL : TOB;
    Acol,Arow : integer;
    nivInterne : integer;
begin
tn := TV.Selected;
TOBL := tn.data;
if TOBL <> nil then
   begin
   GS.CacheEdit;
   DefiniOuvert (TOBL.parent);
   NbLigne := 0;
   NivInterne := 1;
   Renumerote (TOBSimulPres,Nbligne,NivInterne); inc(Nbligne); Niveau := NivInterne;
   DefiniNiveauVisible (Niveau);
   AfficheLaGrille (GS,ListeSaisie,TobSimulPres,Niveau,Nbligne,Paragraphe);
   Arow := TOBL.GetValue('NUMAFF');
   Acol := -1;
   if TOBL.GetValue('NIVEAU') = 1 then Acol := GS_PARAG;
   if TOBL.GetValue('NIVEAU') = 2 then Acol := GS_CATEGORIE;
   if TOBL.GetValue('NIVEAU') = 3 then Acol := GS_NATURE;
   if TOBL.GetValue('NIVEAU') = 4 then Acol := GS_CODE;
   if Acol >= 0 then PositionneLigne (Arow,Acol);
   gs.MontreEdit;
   end;
end;

procedure TFSimulRentab.GSMouseMove(Sender: TObject; Shift: TShiftState; X,Y: Integer);
var Acol,Arow : integer;
begin
  GS.MouseToCell (X,Y,Acol,Arow);
	if (Arow < GS.fixedRows) then exit;
  SetRow (Arow);
  ShowButton (Arow);
end;

procedure TFSimulRentab.SetRow (Arow : integer);
begin
  StoredRow := Arow;
end;

function TFSimulRentab.GetRow : integer;
begin
  result := StoredRow;
end;

procedure TFSimulRentab.ApplicPVHClick(Sender: TObject);
begin
PVMOYChange (self);
end;

procedure TFSimulRentab.QTEEnter(Sender: TObject);
begin
TheQte := Qte.Value;
end;

procedure TFSimulRentab.MPAEnter(Sender: TObject);
begin
TheMPA := MPA.value;
end;

procedure TFSimulRentab.MPREnter(Sender: TObject);
begin
TheMPR := MPR.Value;
end;

procedure TFSimulRentab.COEFMARGEnter(Sender: TObject);
begin
	TheCOEFMARG := COEFMARG.Value;
end;

procedure TFSimulRentab.MPVEnter(Sender: TObject);
begin
TheMPV := MPV.Value;
end;

procedure TFSimulRentab.MPVTTCEnter(Sender: TObject);
begin
TheMPVTTC := MPVTTC.Value;
end;

procedure TFSimulRentab.COEFFGEnter(Sender: TObject);
begin
TheCOEFFG := COEFFG.Value;
end;

procedure TFSimulRentab.PAMOYEnter(Sender: TObject);
begin
THEPAMOY :=  PAMOY.Value;
end;

procedure TFSimulRentab.PRMOYEnter(Sender: TObject);
begin
ThePRMOY := PRMOY.Value;
end;

procedure TFSimulRentab.PVMOYEnter(Sender: TObject);
begin
THEPVMOY := PVMOY.Value;
end;

procedure TFSimulRentab.PVMOYTTCEnter(Sender: TObject);
begin
THEPVMOYTTC := PVMOYTTC.Value;
end;

procedure TFSimulRentab.GestionCoefFrais (TOBL : TOB);
var Valor : T_Valeurs;
		newCoefFc : Double;
begin
// il faut recalculer le montant global PA pour recalculer le coef FC et ensuite appliquer les coefs de frais
InitTableau (valor);
CalculePAGlobal (TOBSImulPres,valor);
// On a ici le déboursé global du document ..
NewCoefFc := GetCoefFc(TOBSimulPres.GetValue('MONTANTPAFG')+TOBSimulPres.GetValue('MONTANTFG'),TOBPieceFrais.getValue('MONTANTFC'));
AppliqueCoefFC (TOBSimulPres,NewCoefFC);
end;

end.

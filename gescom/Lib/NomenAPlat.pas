unit NomenAPlat;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ComCtrls, Grids, Hctrls, ExtCtrls, UTOB, NomenUtil,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      MPlayer, Fe_Main,
{$ENDIF}
  HSysMenu, Menus, CasEmplois, HTB97, UtilArticle, Mask,UentCOmmun,
  TntGrids;

type
  TFNomenAPLat = class(TForm)
    PANNOMENENT: TPanel;
    TGNE_ARTICLE: TLabel;
    TGNE_NOMENCLATURE: TLabel;
    TGNE_LIBELLE: TLabel;
    GNE_ARTICLE: TEdit;
    GNE_NOMENCLATURE: TEdit;
    GNE_LIBELLE: TEdit;
    PANPIED: TPanel;
    HMTrad: THSystemMenu;
    PopupG_NLIG: TPopupMenu;
    PopG_NLIG_C: TMenuItem;
    PopG_NLIG_A: TMenuItem;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BAide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BValider: TToolbarButton97;
    PANDIM: TPanel;
    TGA_CODEDIM1: TLabel;
    TGA_CODEDIM2: TLabel;
    TGA_CODEDIM3: TLabel;
    TGA_CODEDIM4: TLabel;
    TGA_CODEDIM5: TLabel;
    G_NLIG: THGrid;
    TV_NLIG: TTreeView;
    PopG_NLIG_D: TMenuItem;
    N2: TMenuItem;
    PopG_NLIG_P: TMenuItem;
    BInfos: TToolbarButton97;
    GNE_DPA: THCritMaskEdit;
    GNE_DPR: THCritMaskEdit;
    GNE_PVHT: THCritMaskEdit;
    GNE_PVTTC: THCritMaskEdit;
    GNE_PRHT: THCritMaskEdit;
    GNE_PMAP: THCritMaskEdit;
    GNE_PMRP: THCritMaskEdit;
    GNE_PAHT: THCritMaskEdit;
    GNL_DPA: THCritMaskEdit;
    GNL_DPR: THCritMaskEdit;
    GNL_PVHT: THCritMaskEdit;
    GNL_PVTTC: THCritMaskEdit;
    GNL_PAHT: THCritMaskEdit;
    GNL_PRHT: THCritMaskEdit;
    GNL_PMAP: THCritMaskEdit;
    GNL_PMRP: THCritMaskEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure G_NLIGRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_NLIGDblClick(Sender: TObject);
    procedure TV_NLIGMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure PopG_NLIG_AClick(Sender: TObject);
    procedure PopG_NLIG_CClick(Sender: TObject);
    procedure PopG_NLIG_PClick(Sender: TObject);
    procedure PopG_NLIG_DClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    { Déclarations privées }
    TobNomen : TOB;
    TobResul : TOB;
//    TobArt : TOB;
    procedure AffectePied (ARow : Longint) ;
    procedure EffaceDimensions;
    procedure ChargeTobArticles;
//    function  ChercheLigne (Valeur : string) : integer; // DBR Fiche 10995
    function  ChercheLigne (tn1 : TTreeNode) : integer; // DBR Fiche 10995
  public
    { Déclarations publiques }
  end;

var
  FNomenAPLat: TFNomenAPLat;
  SGN_Comp    : integer;
  SGN_IdComp  : integer;
  SGN_Lib     : integer;
  SGN_Qte     : integer;
  Fmt_Prix    : string;

implementation

{$R *.DFM}

//=============================================================================
procedure TFNomenAPLat.FormCreate(Sender: TObject);
begin
TobNomen := TOB.Create('', nil, -1);
TobResul := TOB.Create('', nil, -1);
end;

//=============================================================================
procedure TFNomenAPLat.FormShow(Sender: TObject);
var
    i_ind1 : integer;
    Cancel : boolean;
    st1 : string;
    Valeurs : T_Valeurs;
begin
PANDIM.Visible := PopG_NLIG_D.Checked;
PANPIED.Visible := PopG_NLIG_P.Checked;
G_NLIG.Cells[0, 0] := 'Code Article';
G_NLIG.Cells[1, 0] := 'Libellé';
G_NLIG.Cells[2, 0] := 'Quantité';
G_NLIG.Cells[3, 0] := '';
SGN_Comp := 0;
SGN_Lib := 1;
SGN_Qte := 2;
SGN_IdComp := 3;
G_NLIG.ColWidths[3] := 0;
G_NLIG.ColAligns[2] := taRightJustify;
HMTrad.ResizeGridColumns (G_NLIG) ;
st1 := GNE_ARTICLE.Text;
GNE_ARTICLE.Text := CodeArticleGenerique(st1, st1, st1, st1, st1, st1);
MiseAPlat(GNE_NOMENCLATURE.Text, nil, TobResul);
Valorise('', TobResul, Valeurs);
Fmt_Prix := '%10.3f';
for i_ind1 := 0 to 7 do
    case i_ind1 of
    0 : GNE_DPA.Text := FloatToStr(Valeurs[0]);
    1 : GNE_DPR.Text := FloatToStr(Valeurs[1]);
    2 : GNE_PVHT.Text := FloatToStr(Valeurs[2]);
    3 : GNE_PVTTC.Text := FloatToStr(Valeurs[3]);
    4 : GNE_PAHT.Text := FloatToStr(Valeurs[4]);
    5 : GNE_PRHT.Text := FloatToStr(Valeurs[5]);
    6 : GNE_PMAP.Text := FloatToStr(Valeurs[6]);
    7 : GNE_PMRP.Text := FloatToStr(Valeurs[7]);
    end;
AffecteLibelle(TobResul, 'GNL_CODEARTICLE');
TobResul.PutGridDetail(G_NLIG, False, False, 'GNL_CODEARTICLE;GNL_LIBELLE;GNL_QTE;GNL_ARTICLE');
AfficheTreeView(TV_NLIG, TobResul, 'LIBELLE');
ChargeTobArticles;
TV_NLIG.TopItem.Text := GNE_NOMENCLATURE.Text;
TV_NLIG.FullExpand;
Cancel := False;
G_NLIGRowEnter(Sender, 1, Cancel, False);
// AffectePied(G_NLIG.Row);
end;

procedure TFNomenAPLat.BitBtn1Click(Sender: TObject);
begin
Close;
end;

procedure TFNomenAPLat.G_NLIGRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
TV_NLIG.Selected := TTreeNode(G_NLIG.Objects[SGN_Lib, Ou]);
AffectePied(Ou);
end;

procedure TFNomenAPLat.G_NLIGDblClick(Sender: TObject);
begin
if (G_NLIG.Cells[SGN_Comp, G_NLIG.Row] = '') then Exit;
AGLLanceFiche('GC','GCARTICLE', '', TOB(G_NLIG.Objects[SGN_Comp, G_NLIG.Row]).GetValue('GNL_ARTICLE'),
              'ACTION=CONSULTATION')
end;

procedure TFNomenAPLat.PopG_NLIG_AClick(Sender: TObject);
begin
G_NLIGDblClick(Sender);
end;

procedure TFNomenAPLat.PopG_NLIG_CClick(Sender: TObject);
begin
Entree_CasEmploi(['NOMENCLATURE=NON', G_NLIG.Cells[0, G_NLIG.Row]], 1);
end;

procedure TFNomenAPLat.PopG_NLIG_DClick(Sender: TObject);
begin
PopG_NLIG_D.Checked := not (PopG_NLIG_D.Checked);
PANDIM.Visible := PopG_NLIG_D.Checked;
end;

procedure TFNomenAPLat.PopG_NLIG_PClick(Sender: TObject);
begin
PopG_NLIG_P.Checked := not (PopG_NLIG_P.Checked);
PANPIED.Visible := PopG_NLIG_P.Checked;
end;

procedure TFNomenAPlat.AffectePied (ARow : Longint) ;
var
//    TOBArt : TOB;
    i_ind1, i_ind2 : integer;
//    st1, Dim1, Dim2, Dim3, Dim4, Dim5 : string; *)
    LibelleDim, GrilleDim : string;

begin
if G_NLIG.Cells[0, ARow] = '' then
    begin
    EffaceDimensions;
    GNL_DPA.Text := '';
    GNL_DPR.Text := '';
    GNL_PVHT.Text := '';
    GNL_PVTTC.Text := '';
    GNL_PAHT.Text := '';
    GNL_PRHT.Text := '';
    GNL_PMAP.Text := '';
    GNL_PMRP.Text := '';
    Exit;
    end;

EffaceDimensions;
GNL_DPA.Text := FloatToStr(TOB(G_NLIG.Objects[SGN_Qte, ARow]).GetValue('GA_DPA'));
GNL_DPR.Text := FloatToStr(TOB(G_NLIG.Objects[SGN_Qte, ARow]).GetValue('GA_DPR'));
GNL_PVHT.Text := FloatToStr(TOB(G_NLIG.Objects[SGN_Qte, ARow]).GetValue('GA_PVHT'));
GNL_PVTTC.Text := FloatToStr(TOB(G_NLIG.Objects[SGN_Qte, ARow]).GetValue('GA_PVTTC'));
GNL_PAHT.Text := FloatToStr(TOB(G_NLIG.Objects[SGN_Qte, ARow]).GetValue('GA_PAHT'));
GNL_PRHT.Text := FloatToStr(TOB(G_NLIG.Objects[SGN_Qte, ARow]).GetValue('GA_PRHT'));
GNL_PMAP.Text := FloatToStr(TOB(G_NLIG.Objects[SGN_Qte, ARow]).GetValue('GA_PMAP'));
GNL_PMRP.Text := FloatToStr(TOB(G_NLIG.Objects[SGN_Qte, ARow]).GetValue('GA_PMRP'));
if TOB(G_NLIG.Objects[SGN_Qte, ARow]).GetValue('GA_STATUTART') = 'DIM' then
    begin
    LibelleDim := '';
    GrilleDim := '';
    LibelleDimensions('', TOB(G_NLIG.Objects[SGN_Qte, ARow]), LibelleDim, GrilleDim);
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
    i_ind2 := PANPIED.Width - i_ind1 - 20;
    i_ind1 := i_ind2 div 4;
    TGA_CODEDIM1.Left := 10;
    TGA_CODEDIM2.Left := TGA_CODEDIM1.Left + TGA_CODEDIM1.Width + i_ind1;
    TGA_CODEDIM3.Left := TGA_CODEDIM2.Left + TGA_CODEDIM2.Width + i_ind1;
    TGA_CODEDIM4.Left := TGA_CODEDIM3.Left + TGA_CODEDIM3.Width + i_ind1;
    TGA_CODEDIM5.Left := TGA_CODEDIM4.Left + TGA_CODEDIM4.Width + i_ind1;
    end;
end;

procedure TFNomenAPlat.EffaceDimensions;
begin
TGA_CODEDIM1.Caption := '';
TGA_CODEDIM2.Caption := '';
TGA_CODEDIM3.Caption := '';
TGA_CODEDIM4.Caption := '';
TGA_CODEDIM5.Caption := '';
TGA_CODEDIM1.Visible := False;
TGA_CODEDIM2.Visible := False;
TGA_CODEDIM3.Visible := False;
TGA_CODEDIM4.Visible := False;
TGA_CODEDIM5.Visible := False;
end;

procedure TFNomenAPlat.ChargeTobArticles;
var
    i_ind1 : integer;
    st1 : string;

begin
for i_ind1 := 1 to G_NLIG.RowCount - 1 do
    begin
    if G_NLIG.Cells[SGN_Comp, i_ind1] = '' then Exit;
    G_NLIG.Objects[SGN_Qte, i_ind1] := TOB.Create('ARTICLE', nil, -1);
    st1 := G_NLIG.Cells[SGN_IdComp, i_ind1];
    Tob(G_NLIG.Objects[SGN_Qte, i_ind1]).SelectDB('"' + st1 + '"', nil);
//    G_NLIG.Objects[SGN_Lib, i_ind1] := ChercheNode(TV_NLIG, nil, G_NLIG.Cells[SGN_Comp, i_ind1]); // DBR Fiche 10995
    G_NLIG.Objects[SGN_Lib, i_ind1] := ChercheNode(TV_NLIG, nil, i_ind1 + 1); // DBR Fiche 10995
    end;
end;

//=============================================================================
//  Recherche d'une ligne de la grille en fonction d'une valeur chaine de caracteres
//=============================================================================
//function TFNomenAPlat.ChercheLigne (Valeur : string) : integer; // DBR Fiche 10995
function TFNomenAPlat.ChercheLigne (tn1 : TTreeNode) : integer; // DBR Fiche 10995
var
    i_ind1 : integer ;
//    st1, st2 : string; // DBR Fiche 10995
begin
Result := 0;
for i_ind1 := 0 to G_NLIG.RowCount - 1 do
    begin
(*    st1 := G_NLIG.Cells[0, i_ind1]; // DBR Fiche 10995
    st2 := Trim(Copy(Valeur, 0, Pos('(', Valeur) - 1));
    if Pos(st2, st1) <> 0 then *)
    if tn1 = TTreeNode (G_NLIG.Objects[SGN_Lib, i_ind1]) then // DBR Fiche 10995
        begin
        Result := i_ind1;
        Exit;
        end;
    end;
end;

procedure TFNomenAPLat.TV_NLIGMouseUp(Sender: TObject; Button: TMouseButton;
                                      Shift: TShiftState; X, Y: Integer);
var
    tn1 : TTreeNode;
    i_ind1 : integer;
begin
tn1 := TV_NLIG.GetNodeAt(X, Y);
if tn1 = nil then Exit;
//--------------------------------------------------------------------------
//  affichage synchrone de la grid avec le treeview
//--------------------------------------------------------------------------
//i_ind1 := ChercheLigne(tn1.text); // DBR Fiche 10995
i_ind1 := ChercheLigne(TV_NLIG.Selected); // DBR Fiche 10995
if i_ind1 <> 0 then
    begin
    G_NLIG.Col := 0;
    G_NLIG.Row := i_ind1;
    G_NLIG.SetFocus;
    end;
end;

procedure TFNomenAPLat.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.

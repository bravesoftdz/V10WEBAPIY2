unit RecupGBParamMvt;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, FE_Main, ExtCtrls, DicoAf, Spin, Hent1, EntGC, UTOB, HTB97,
  Mask, DBCtrls, HDB, HSysMenu;

type
  TFRecupParamMvt = class(TForm)
    FermerParamFicBase: TButton;
    GroupCde: TGroupBox;
    HLabel10: THLabel;
    NatCde: THValComboBox;
    GroupAnn: TGroupBox;
    HLabel7: THLabel;
    GroupRec: TGroupBox;
    HLabel1: THLabel;
    NatRec: THValComboBox;
    GroupFac: TGroupBox;
    GroupEnt: TGroupBox;
    HLabel3: THLabel;
    NatEnt: THValComboBox;
    GroupVte: TGroupBox;
    HLabel4: THLabel;
    NatVte: THValComboBox;
    GroupTrf: TGroupBox;
    HLabel5: THLabel;
    NATTRFRE: THValComboBox;
    GroupSor: TGroupBox;
    HLabel6: THLabel;
    NatSor: THValComboBox;
    ModeReglVte: THValComboBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    GroupBox1: TGroupBox;
    ValStock: TCheckBox;
    Tarif: TCheckBox;
    TypeTarifTTC: THValComboBox;
    Label5: TLabel;
    Label7: TLabel;
    HLabel8: THLabel;
    NATTRFEM: THValComboBox;
    PeriodeTarifTTC: THValComboBox;
    TypeTarifsolde: THValComboBox;
    PeriodeTarifSolde: THValComboBox;
    TarifHT: TCheckBox;
    GroupBox2: TGroupBox;
    HLabel9: THLabel;
    HLabel11: THLabel;
    NATNEGFAC: THValComboBox;
    NATNEGCDE: THValComboBox;
    HLabel12: THLabel;
    NATNEGLIV: THValComboBox;
    HLabel13: THLabel;
    NATNEGAVO: THValComboBox;
    StockCloture: TCheckBox;
    Label4: TLabel;
    DateCloture: THCritMaskEdit;
    HLabel14: THLabel;
    REPPAARTICLE: TCheckBox;
    REPPVARTICLE: TCheckBox;
    REPPVHTARTICLE: TCheckBox;
    REPPRARTICLE: TCheckBox;
    Label6: TLabel;
    Codeartvente: THCritMaskEdit;
    HMTrad: THSystemMenu;
    CodeCliVte: THCritMaskEdit;
    TarifAch: TCheckBox;
    TypeTarifAch: THValComboBox;
    PeriodeTarifAch: THValComboBox;
    procedure FermerParamFicBaseClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure TarifClick(Sender: TObject);
    procedure ValStockClick(Sender: TObject);
    procedure StockClotureClick(Sender: TObject);
    procedure TarifAchClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    Tob_parametrage : TOB;
    { Déclarations publiques }
  end;

Procedure AppelRecupGBParamMvt ( var TOBParam : TOB; var DteCloture : TDateTime; var CodeArticle : string ) ;

implementation

{$R *.DFM}

Procedure AppelRecupGBParamMvt ( var TOBParam : TOB; var DteCloture : TDateTime; var CodeArticle : string ) ;
var X  : TFRecupParamMvt ;
BEGIN
SourisSablier;
X:=TFRecupParamMvt.Create(Application) ;
X.Tob_Parametrage:=TOBParam ;
X.DateCloture.Text := DateToStr (DteCloture) ;
X.CodeArtVente.Text := CodeArticle ;
X.Tob_Parametrage.PutEcran(X) ;
try
 X.ShowModal ;
 DteCloture := StrToDate (X.DateCloture.Text);
 CodeArticle := X.Codeartvente.text;
 finally
 X.Free ;
 end ;
SourisNormale ;
END ;

procedure TFRecupParamMvt.FermerParamFicBaseClick(Sender: TObject);
begin
Tob_Parametrage.GetEcran(Self) ;
Self.Close;
end;

procedure TFRecupParamMvt.FormShow(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  AglEmpileFiche(Self) ;
end;

procedure TFRecupParamMvt.TarifClick(Sender: TObject);
begin
 if Tarif.Checked then
 begin
  TypeTarifTTC.enabled      := TRUE;
  PeriodeTarifTTC.enabled   := TRUE;
  TypeTarifSolde.enabled    := TRUE;
  PeriodeTarifSolde.enabled := TRUE;
 end else
 begin
  TypeTarifTTC.enabled      := False ;
  PeriodeTarifTTC.enabled   := False;
  TypeTarifSolde.enabled    := False;
  PeriodeTarifSolde.enabled := False;
 end;
end;

procedure TFRecupParamMvt.ValStockClick(Sender: TObject);
begin
 if ValStock.Checked then
 begin
  StockCloture.Enabled:=TRUE ;
  DateCloture.Enabled :=TRUE ;
 end else
 begin
   StockCloture.Checked := False;
   StockCloture.Enabled :=False ;
   DateCloture.Enabled  :=False ;
 end;
end;

procedure TFRecupParamMvt.StockClotureClick(Sender: TObject);
begin
  if StockCloture.Checked then
  begin
    DateCloture.Enabled :=TRUE ;
  end else
  begin
    DateCloture.Enabled :=False ;
  end;
end;

procedure TFRecupParamMvt.FormDestroy(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche ;
end;

procedure TFRecupParamMvt.TarifAchClick(Sender: TObject);
begin
 if TarifAch.Checked then
 begin
  TypeTarifAch.enabled      := TRUE;
  PeriodeTarifAch.enabled   := TRUE;
 end else
 begin
  TypeTarifAch.enabled      := False ;
  PeriodeTarifTTC.enabled   := False;
 end;
end;

end.

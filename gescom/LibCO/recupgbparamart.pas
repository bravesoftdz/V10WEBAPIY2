unit RecupGBParamArt;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, HTB97, ExtCtrls, DicoAf, Hctrls, Spin, ComCtrls, Hent1, UTOB,
  Mask, HSysMenu;

type
  TFRecupParamArt = class(TForm)
    FermerParamFicBase: TButton;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    GroupBox4: TGroupBox;
    HLabel1: THLabel;
    HLabel2: THLabel;
    HLabel3: THLabel;
    HLabel4: THLabel;
    FamNiv1: TSpinEdit;
    FamNiv2: TSpinEdit;
    FamNiv3: TSpinEdit;
    GroupBox1: TGroupBox;
    HLabel5: THLabel;
    HLabel6: THLabel;
    HLabel7: THLabel;
    HLabel8: THLabel;
    HLabel9: THLabel;
    HLabel10: THLabel;
    HLabel11: THLabel;
    HLabel12: THLabel;
    HLabel13: THLabel;
    HLabel14: THLabel;
    HLabel15: THLabel;
    RepMod: THValComboBox;
    RepCat: THValComboBox;
    RepMat: THValComboBox;
    RepCol: THValComboBox;
    RepFin: THValComboBox;
    RepStr: THValComboBox;
    RepFor: THValComboBox;
    RepCp1: THValComboBox;
    RepCp2: THValComboBox;
    RepTyp: THValComboBox;
    RepLibCol: THValComboBox;
    GroupBox2: TGroupBox;
    HLabel16: THLabel;
    CategorieDim: THValComboBox;
    GroupBox3: TGroupBox;
    Label1: TLabel;
    RepTVA: THValComboBox;
    GroupBox5: TGroupBox;
    HLabel17: THLabel;
    HLabel18: THLabel;
    HLabel19: THLabel;
    HLabel25: THLabel;
    HLabel26: THLabel;
    HLabel27: THLabel;
    HLabel28: THLabel;
    HLabel29: THLabel;
    HLabel30: THLabel;
    HLabel31: THLabel;
    HLabel32: THLabel;
    HLabel33: THLabel;
    HLabel34: THLabel;
    HLabel35: THLabel;
    HLabel36: THLabel;
    HLabel37: THLabel;
    HLabel38: THLabel;
    HLabel39: THLabel;
    HLabel40: THLabel;
    HLabel41: THLabel;
    RepStatA01: THValComboBox;
    RepStatA02: THValComboBox;
    RepStatA03: THValComboBox;
    RepStatA04: THValComboBox;
    RepStatA05: THValComboBox;
    RepStatA06: THValComboBox;
    RepStatA07: THValComboBox;
    RepStatA08: THValComboBox;
    RepStatA09: THValComboBox;
    RepStatA10: THValComboBox;
    RepStatA11: THValComboBox;
    RepStatA12: THValComboBox;
    RepStatA13: THValComboBox;
    RepStatA14: THValComboBox;
    RepStatA15: THValComboBox;
    RepStatA16: THValComboBox;
    RepStatA17: THValComboBox;
    RepStatA18: THValComboBox;
    RepStatA19: THValComboBox;
    RepStatA20: THValComboBox;
    Repreg: THValComboBox;
    Label2: TLabel;
    GroupBox6: TGroupBox;
    HLabel20: THLabel;
    RepCodeBarre: THValComboBox;
    GroupBox7: TGroupBox;
    HLabel21: THLabel;
    RepDevise: THValComboBox;
    RecodifieFamille: TCheckBox;
    GroupBox8: TGroupBox;
    OptimisationArticle: TCheckBox;
    Label3: TLabel;
    GRILLEUNI: THCritMaskEdit;
    LGRILLEUNI: THLabel;
    HLabel22: THLabel;
    RepFamille: THValComboBox;
    HMTrad: THSystemMenu;
    procedure FermerParamFicBaseClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FamNivEnter(Sender: TObject);
    procedure Controle_ChampLibre(Sender: TObject);
    procedure RecodifieFamilleClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    Tob_parametrage : TOB;
    { Déclarations publiques }
  end;

  
Procedure AppelRecupGBParamArt ( var TOBParam : TOB ) ;

implementation

{$R *.DFM}

Procedure AppelRecupGBParamArt ( var TOBParam : TOB ) ;
var X  : TFRecupParamArt ;
BEGIN
SourisSablier;
X:=TFRecupParamArt.Create(Application) ;
X.Tob_Parametrage:=TOBParam ;
X.Tob_Parametrage.PutEcran(X) ;
try
 X.ShowModal ;
 finally
 X.Free ;
 end ;
SourisNormale ;
END ;

procedure TFRecupParamArt.FermerParamFicBaseClick(Sender: TObject);
begin
Tob_Parametrage.GetEcran(Self) ;
Self.Close;
end;

procedure TFRecupParamArt.FormShow(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  AglEmpileFiche(Self) ;
  if RecodifieFamille.Checked then
  begin
 	  FamNiv1.enabled      := FALSE;
    FamNiv2.enabled      := FALSE;
    FamNiv3.enabled      := FALSE;
  end else
  begin
    FamNiv1.enabled      := TRUE;
    FamNiv2.enabled      := TRUE;
    FamNiv3.enabled      := TRUE;
  end;
end;

procedure TFRecupParamArt.FamNivEnter(Sender: TObject);
begin
if (FamNiv1.Value <> 0) then FamNiv2.Enabled := True else
   begin
   FamNiv2.Value := 0;
   FamNiv2.Enabled := False;
   end;
if (FamNiv2.Value <> 0) then FamNiv3.Enabled := True else
   begin
   FamNiv3.Value := 0;
   FamNiv3.Enabled := False;
   end;
if (FamNiv1.Value+FamNiv2.Value+FamNiv3.Value > 6) then FamNiv3.Value := 6-FamNiv1.Value-FamNiv2.Value;
end;

procedure TFRecupParamArt.Controle_ChampLibre(Sender: TObject);
var Ctrl : boolean;
begin
inherited;
  Ctrl := True;
  if (THValComboBox(Sender).Value = '') then exit;
  if (THValComboBox(Sender) <> RepMod) and (THValComboBox(Sender).Value = RepMod.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepMat) and (THValComboBox(Sender).Value = RepMat.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepCol) and (THValComboBox(Sender).Value = RepCol.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepFin) and (THValComboBox(Sender).Value = RepFin.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepCat) and (THValComboBox(Sender).Value = RepCat.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepFor) and (THValComboBox(Sender).Value = RepFor.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepCp1) and (THValComboBox(Sender).Value = RepCp1.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepCp2) and (THValComboBox(Sender).Value = RepCp2.Value) then Ctrl := False
  // Zones gérées par la table GCCODELIBREART
  else if (THValComboBox(Sender) <> RepStr) and (THValComboBox(Sender).Value = RepStr.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA01) and (THValComboBox(Sender).Value = RepStatA01.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA02) and (THValComboBox(Sender).Value = RepStatA02.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA03) and (THValComboBox(Sender).Value = RepStatA03.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA04) and (THValComboBox(Sender).Value = RepStatA04.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA05) and (THValComboBox(Sender).Value = RepStatA05.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA06) and (THValComboBox(Sender).Value = RepStatA06.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA07) and (THValComboBox(Sender).Value = RepStatA07.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA08) and (THValComboBox(Sender).Value = RepStatA08.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA09) and (THValComboBox(Sender).Value = RepStatA09.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA10) and (THValComboBox(Sender).Value = RepStatA10.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA11) and (THValComboBox(Sender).Value = RepStatA11.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA12) and (THValComboBox(Sender).Value = RepStatA12.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA13) and (THValComboBox(Sender).Value = RepStatA13.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA14) and (THValComboBox(Sender).Value = RepStatA14.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA15) and (THValComboBox(Sender).Value = RepStatA15.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA16) and (THValComboBox(Sender).Value = RepStatA16.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA17) and (THValComboBox(Sender).Value = RepStatA17.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA18) and (THValComboBox(Sender).Value = RepStatA18.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA19) and (THValComboBox(Sender).Value = RepStatA19.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatA20) and (THValComboBox(Sender).Value = RepStatA20.Value) then Ctrl := False
  // Zones additionnelles gérées
  else if (THValComboBox(Sender) <> RepTyp) and (THValComboBox(Sender).Value = RepTyp.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepLibCol) and (THValComboBox(Sender).Value = RepLibCol.Value) then Ctrl := False;

  if Ctrl = False then
  begin
   ShowMessage (TraduireMemoire ('Affectation impossible, cette zone est déjà utilisée !'));
   THValComboBox(Sender).Value := '';
  end;
end;

procedure TFRecupParamArt.RecodifieFamilleClick(Sender: TObject);
begin
 if RecodifieFamille.Checked then
 begin
 	FamNiv1.enabled      := FALSE;
  FamNiv2.enabled      := FALSE;
  FamNiv3.enabled      := FALSE;
 end else
 begin
  FamNiv1.enabled      := TRUE;
  FamNiv2.enabled      := TRUE;
  FamNiv3.enabled      := TRUE;
 end;
end;

procedure TFRecupParamArt.FormDestroy(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche ;
end;

end.

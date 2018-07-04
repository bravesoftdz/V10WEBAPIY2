unit RecupGBParamGen;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, HTB97, ExtCtrls, DicoAf, Hctrls, Spin, Hent1, UTOB, Grids,
  HSysMenu ;

type
  TFRecupParamGen = class(TForm)
    FermerParamFicBase: TButton;
    GroupBox2: TGroupBox;
    HLabel16: THLabel;
    HLabel17: THLabel;
    HLabel18: THLabel;
    HLabel19: THLabel;
    HLabel21: THLabel;
    HLabel22: THLabel;
    HLabel23: THLabel;
    HLabel24: THLabel;
    RepMod: THValComboBox;
    RepCat: THValComboBox;
    RepMat: THValComboBox;
    RepCol: THValComboBox;
    RepFin: THValComboBox;
    RepFor: THValComboBox;
    RepCp1: THValComboBox;
    RepCp2: THValComboBox;
    GroupBox1: TGroupBox;
    HLabel13: THLabel;
    HLabel10: THLabel;
    FamNiv1: TSpinEdit;
    HLabel11: THLabel;
    FamNiv2: TSpinEdit;
    HLabel12: THLabel;
    FamNiv3: TSpinEdit;
    GroupBox3: TGroupBox;
    HLabel45: THLabel;
    CategorieDim: THValComboBox;
    RecodifieFamille: TCheckBox;
    Label3: TLabel;
    HLabel1: THLabel;
    RepFamille: THValComboBox;
    HMTrad: THSystemMenu;
    procedure FermerParamFicBaseClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy (Sender: TObject);
    procedure FamNivEnter(Sender: TObject);
    procedure Controle_ChampLibre(Sender: TObject);
    procedure RecodifieFamilleClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    Tob_parametrage : TOB;
    { Déclarations publiques }
  end;

Procedure AppelRecupGBParamGen ( var TOBParam : TOB ) ;

implementation

{$R *.DFM}

Procedure AppelRecupGBParamGen (  var TOBParam : TOB ) ;
var X  : TFRecupParamGen ;
BEGIN
SourisSablier;
X:=TFRecupParamGen.Create(Application) ;
// Récupération des 3 TOB
X.Tob_Parametrage := TOBParam ;
X.Tob_Parametrage.PutEcran(X) ;

try
 X.ShowModal ;
 finally
 X.Free ;
 end ;
SourisNormale ;
END ;

procedure TFRecupParamGen.FermerParamFicBaseClick(Sender: TObject);
begin
Tob_Parametrage.GetEcran(Self) ;
Self.Close;
end;

procedure TFRecupParamGen.FormShow(Sender: TObject);
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

procedure TFRecupParamGen.FamNivEnter(Sender: TObject);
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

procedure TFRecupParamGen.Controle_ChampLibre(Sender: TObject);
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
  else if (THValComboBox(Sender) <> RepCp2) and (THValComboBox(Sender).Value = RepCp2.Value) then Ctrl := False;
  
  if Ctrl = False then
  begin
   ShowMessage (TraduireMemoire ('Affectation impossible, cette zone est déjà utilisée !'));
   THValComboBox(Sender).Value := '';
  end;

  // Test des zones renseignées dans les fournisseurs
  if Ctrl = False then exit ;
  // Recherche si une affectation FOURNISSEUR utilise une zone affectee aux CLIENTS
  if THValComboBox(Sender).Value =  Tob_Parametrage.GetValue ('REPTYP') then
  begin
    ShowMessage (TraduireMemoire ('Attention, cette zone est déjà utilisée pour le type d''article !'));
    THValComboBox(Sender).Value := '';
  end;
end;

procedure TFRecupParamGen.RecodifieFamilleClick(Sender: TObject);
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

procedure TFRecupParamGen.FormDestroy(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche ;
end;

end.

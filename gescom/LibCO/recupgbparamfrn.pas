unit RecupGBParamFrn;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, HTB97, ExtCtrls, DicoAf, Hctrls, Spin, Hent1, UTOB, HSysMenu;

type
  TFRecupParamFrn = class(TForm)
    FermerParamFicBase: TButton;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    HLabel65: THLabel;
    HLabel66: THLabel;
    HLabel67: THLabel;
    HLabel68: THLabel;
    HLabel69: THLabel;
    HLabel70: THLabel;
    HLabel71: THLabel;
    HLabel72: THLabel;
    HLabel73: THLabel;
    HLabel74: THLabel;
    HLabel75: THLabel;
    HLabel76: THLabel;
    HLabel77: THLabel;
    HLabel78: THLabel;
    HLabel79: THLabel;
    HLabel80: THLabel;
    HLabel81: THLabel;
    HLabel82: THLabel;
    HLabel83: THLabel;
    HLabel84: THLabel;
    RepStatF01: THValComboBox;
    RepStatF02: THValComboBox;
    RepStatF03: THValComboBox;
    RepStatF04: THValComboBox;
    RepStatF05: THValComboBox;
    RepStatF06: THValComboBox;
    RepStatF07: THValComboBox;
    RepStatF08: THValComboBox;
    RepStatF09: THValComboBox;
    RepStatF10: THValComboBox;
    RepStatF11: THValComboBox;
    RepStatF12: THValComboBox;
    RepStatF13: THValComboBox;
    RepStatF14: THValComboBox;
    RepStatF15: THValComboBox;
    RepStatF16: THValComboBox;
    RepStatF17: THValComboBox;
    RepStatF18: THValComboBox;
    RepStatF19: THValComboBox;
    RepStatF20: THValComboBox;
    Label1: TLabel;
    FrnRegimeTVA: THValComboBox;
    Label2: TLabel;
    FRNFAMCOMPTA: THValComboBox;
    Label3: TLabel;
    HMTrad: THSystemMenu;
    procedure FermerParamFicBaseClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Controle_ChampLibre(Sender: TObject);
  private
    { Déclarations privées }
  public
    Tob_Parametrage : TOB;
    { Déclarations publiques }
  end;

  
Procedure AppelRecupGBParamFrn ( var TOBParam : TOB );

implementation

{$R *.DFM}

Procedure AppelRecupGBParamFrn ( var TOBParam : TOB ) ;
var X  : TFRecupParamFrn ;
BEGIN
SourisSablier;
X:=TFRecupParamFrn.Create(Application) ;
X.Tob_Parametrage:=TOBParam ;
X.Tob_Parametrage.PutEcran(X) ;
try
 X.ShowModal ;
 finally
 X.Free ;
 end ;
SourisNormale ;
END ;

procedure TFRecupParamFrn.FermerParamFicBaseClick(Sender: TObject);
begin
Tob_Parametrage.GetEcran(Self) ;

Self.Close;
end;

procedure TFRecupParamFrn.FormShow(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  AglEmpileFiche(Self) ;
end;

procedure TFRecupParamFrn.Controle_ChampLibre(Sender: TObject);
var Ctrl : boolean;
   st:string;
begin
inherited;
  Ctrl := True;
  st:=Tob_Parametrage.GetValue ('REPSTATC01');

  if (THValComboBox(Sender).Value = '') then exit;
  if (THValComboBox(Sender) <> RepStatF01) and (THValComboBox(Sender).Value = RepStatF01.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF02) and (THValComboBox(Sender).Value = RepStatF02.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF03) and (THValComboBox(Sender).Value = RepStatF03.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF04) and (THValComboBox(Sender).Value = RepStatF04.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF05) and (THValComboBox(Sender).Value = RepStatF05.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF06) and (THValComboBox(Sender).Value = RepStatF06.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF07) and (THValComboBox(Sender).Value = RepStatF07.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF08) and (THValComboBox(Sender).Value = RepStatF08.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF09) and (THValComboBox(Sender).Value = RepStatF09.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF10) and (THValComboBox(Sender).Value = RepStatF10.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF11) and (THValComboBox(Sender).Value = RepStatF11.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF12) and (THValComboBox(Sender).Value = RepStatF12.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF13) and (THValComboBox(Sender).Value = RepStatF13.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF14) and (THValComboBox(Sender).Value = RepStatF14.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF15) and (THValComboBox(Sender).Value = RepStatF15.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF16) and (THValComboBox(Sender).Value = RepStatF16.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF17) and (THValComboBox(Sender).Value = RepStatF17.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF18) and (THValComboBox(Sender).Value = RepStatF18.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF19) and (THValComboBox(Sender).Value = RepStatF19.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatF20) and (THValComboBox(Sender).Value = RepStatF20.Value) then Ctrl := False;

  if Ctrl = False then
  begin
    ShowMessage (TraduireMemoire ('Affectation impossible, cette zone est déjà utilisée !'));
    THValComboBox(Sender).Value := '';
  end
  
end;

procedure TFRecupParamFrn.FormDestroy(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche ;
end;

end.

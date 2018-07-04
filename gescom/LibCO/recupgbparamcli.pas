unit RecupGBParamCLi;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, HTB97, ExtCtrls, DicoAf, Hctrls, Spin, Hent1, UTOB, HSysMenu;

type
  TFRecupParamCli = class(TForm)
    FermerParamFicBase: TButton;
    GroupBox5: TGroupBox;
    HLabel6: THLabel;
    HLabel44: THLabel;
    HLabel47: THLabel;
    HLabel48: THLabel;
    HLabel49: THLabel;
    HLabel50: THLabel;
    HLabel51: THLabel;
    HLabel52: THLabel;
    HLabel53: THLabel;
    HLabel54: THLabel;
    HLabel55: THLabel;
    HLabel56: THLabel;
    HLabel57: THLabel;
    HLabel58: THLabel;
    HLabel59: THLabel;
    HLabel60: THLabel;
    HLabel61: THLabel;
    HLabel62: THLabel;
    HLabel63: THLabel;
    HLabel64: THLabel;
    RepStatC01: THValComboBox;
    RepStatC02: THValComboBox;
    RepStatC03: THValComboBox;
    RepStatC04: THValComboBox;
    RepStatC05: THValComboBox;
    RepStatC06: THValComboBox;
    RepStatC07: THValComboBox;
    RepStatC08: THValComboBox;
    RepStatC09: THValComboBox;
    RepStatC10: THValComboBox;
    RepStatC11: THValComboBox;
    RepStatC12: THValComboBox;
    RepStatC13: THValComboBox;
    RepStatC14: THValComboBox;
    RepStatC15: THValComboBox;
    RepStatC16: THValComboBox;
    RepStatC17: THValComboBox;
    RepStatC18: THValComboBox;
    RepStatC19: THValComboBox;
    RepStatC20: THValComboBox;
    Panel1: TPanel;
    CodeCli: TRadioGroup;
    Label7: TLabel;
    RepCliProMail: TComboBox;
    Label8: TLabel;
    CliExport: THValComboBox;
    RepEtaCpta: THValComboBox;
    Label10: TLabel;
    HMTrad: THSystemMenu;
    procedure FermerParamFicBaseClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Controle_ChampLibre(Sender: TObject);
  private
    { Déclarations privées }
  public
    Tob_parametrage : TOB;
    { Déclarations publiques }
  end;

Procedure AppelRecupGBParamCli ( var TOBParam : TOB ) ;

implementation

{$R *.DFM}

Procedure AppelRecupGBParamCli ( var TOBParam : TOB ) ;
var X  : TFRecupParamCli ;
BEGIN
SourisSablier;
X:=TFRecupParamCli.Create(Application) ;
X.Tob_Parametrage:=TOBParam ;

// MODIF LM 26/09/00       Sur Nouvelle version, perte de CODECLI ????
x.CodeCli.ItemIndex := x.Tob_Parametrage.GetValue ('CODECLI');

X.Tob_Parametrage.PutEcran(X) ;
try
 X.ShowModal ;
 finally
 X.Free ;
 end ;
SourisNormale ;
END ;

procedure TFRecupParamCli.FermerParamFicBaseClick(Sender: TObject);
begin

Tob_Parametrage.GetEcran(Self) ;

// MODIF LM 26/09/00           CODECLI n'est plus récupéré sur la nouvelle version ????
Tob_Parametrage.PutValue ('CODECLI', Codecli.ItemIndex );

Self.Close;
end;

procedure TFRecupParamCli.FormShow(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  AglEmpileFiche(Self) ;
end;

procedure TFRecupParamCli.Controle_ChampLibre(Sender: TObject);
var Ctrl : boolean;
begin
inherited;
  Ctrl := True;
  if (THValComboBox(Sender).Value = '') then exit;
  if (THValComboBox(Sender) <> RepStatC01) and (THValComboBox(Sender).Value = RepStatC01.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC02) and (THValComboBox(Sender).Value = RepStatC02.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC03) and (THValComboBox(Sender).Value = RepStatC03.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC04) and (THValComboBox(Sender).Value = RepStatC04.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC05) and (THValComboBox(Sender).Value = RepStatC05.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC06) and (THValComboBox(Sender).Value = RepStatC06.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC07) and (THValComboBox(Sender).Value = RepStatC07.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC08) and (THValComboBox(Sender).Value = RepStatC08.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC09) and (THValComboBox(Sender).Value = RepStatC09.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC10) and (THValComboBox(Sender).Value = RepStatC10.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC11) and (THValComboBox(Sender).Value = RepStatC11.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC12) and (THValComboBox(Sender).Value = RepStatC12.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC13) and (THValComboBox(Sender).Value = RepStatC13.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC14) and (THValComboBox(Sender).Value = RepStatC14.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC15) and (THValComboBox(Sender).Value = RepStatC15.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC16) and (THValComboBox(Sender).Value = RepStatC16.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC17) and (THValComboBox(Sender).Value = RepStatC17.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC18) and (THValComboBox(Sender).Value = RepStatC18.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC19) and (THValComboBox(Sender).Value = RepStatC19.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepStatC20) and (THValComboBox(Sender).Value = RepStatC20.Value) then Ctrl := False
  else if (THValComboBox(Sender) <> RepEtaCpta) and (THValComboBox(Sender).Value = RepEtaCpta.Value) then Ctrl := False;

  if Ctrl = False then
  begin
    ShowMessage (TraduireMemoire ('Affectation impossible, cette zone est déjà utilisée !'));
    THValComboBox(Sender).Value := '';
  end else
  begin
    if (THValComboBox(Sender).Value='YTC_BOOLLIBRE1') and (RepCliProMail.Text='Booléen libre n° 1') then Ctrl := False
    else if (THValComboBox(Sender).Value='YTC_BOOLLIBRE2') and (RepCliProMail.Text='Booléen libre n° 2') then Ctrl := False
    else if (THValComboBox(Sender).Value='YTC_BOOLLIBRE3') and (RepCliProMail.Text='Booléen libre n° 3') then Ctrl := False;
    if Ctrl = False then
    begin
      ShowMessage (TraduireMemoire ('Affectation impossible, valeur est déjà utilisée pour la zone prochain mailing !'));
      THValComboBox(Sender).Value := '';
    end;
  end;

end;

procedure TFRecupParamCli.FormDestroy(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche ;
end;

end.

unit LC_copie;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, StdCtrls, Hctrls, ExtCtrls, HPanel, Math, HEnt1;

function OkRecopie(DepuisCaisse: string): string;

type
  TFRecopieLC = class(TForm)
    PnlFonsRecopie: THPanel;
    TFDepuisCaisse: THLabel;
    FDepuisCaisse: THValComboBox;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BValider: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    BFerme: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
  public
    ACaisse: string;
    DepuisCaisse: string;
  end;

implementation

{$R *.DFM}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function OkRecopie(DepuisCaisse: string): string;
var X: TFRecopieLC;
begin
  Result := '';
  X := TFRecopieLC.Create(Application);
  try
    X.ACaisse := DepuisCaisse;
    if X.ShowModal = mrOk then Result := X.DepuisCaisse;
  finally
    X.Free;
  end;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TFRecopieLC.FormShow(Sender: TObject);
var i, j: Integer;
begin
  j := -1;
  i := 0;
  while (i <= FDepuisCaisse.Values.count - 1) and (J = -1) do
    if FDepuisCaisse.Values[i] = ACaisse then j := i else inc(i);
  if j <> -1 then
  begin
    FDepuisCaisse.Values.Delete(j);
    FDepuisCaisse.Items.delete(j);
  end;
  FDepuisCaisse.ItemIndex := minintvalue([0, FDepuisCaisse.values.count - 1]);
  // Appel de la fonction d'empilage dans la liste des fiches
  AglEmpileFiche(Self);
  Self.HelpContext := 113000200; // JTR - eQualité 11771
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TFRecopieLC.FormDestroy(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TFRecopieLC.BValiderClick(Sender: TObject);
begin
  DepuisCaisse := FDepuisCaisse.Value;
  ModalResult := mrOk;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TFRecopieLC.HelpBtnClick(Sender: TObject);
begin
  Callhelptopic(Self);
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
end.

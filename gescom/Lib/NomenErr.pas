unit NomenErr;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ColMemo, ComCtrls, Buttons, ExtCtrls, ImgList, HTB97, HImgList;

type
  TFNomenErr = class(TForm)
    TV_ERREUR: TTreeView;
    ColorMemo1: TColorMemo;
    Panel1: TPanel;
    Image1: TImage;
    Label1: TLabel;
    Label2: TLabel;
    ImageList1: THImageList;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BAide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BValider: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Déclarations privées }
    procedure Parcours(TN : TTreeNode; Valeur : string; Level : integer);
  public
    { Déclarations publiques }
  end;

var
  FNomenErr: TFNomenErr;

implementation

{$R *.DFM}

procedure TFNomenErr.FormShow(Sender: TObject);
var
    tn1 : TTreeNode;
begin
TV_ERREUR.FullExpand;
ColorMemo1.Lines.Clear;
ColorMemo1.Lines.Text := 'La décomposition de votre nomenclature contient';
ColorMemo1.Lines.Append('des références circulaires dont voici les données :');
ColorMemo1.Lines.Append('');
ColorMemo1.Lines.Append('');
tn1 := TV_ERREUR.TopItem.GetFirstChild;
while tn1 <> nil do
    begin
    tn1.ImageIndex := 1;
    Parcours(tn1, TV_ERREUR.TopItem.Text, TV_ERREUR.TopItem.Level);
    tn1 := tn1.GetNextSibling;
    end;
end;

procedure TFNomenErr.Parcours(TN : TTreeNode; Valeur : string; Level : integer);
var
    tn1 : TTreeNode;
    st1 : string;
begin
    tn1 := TN.GetFirstChild;
while tn1 <> nil do
    begin
    tn1.ImageIndex := 1;
    if tn1.Text = Valeur then
        begin
        tn1.ImageIndex := 0;
        st1 := 'La nomenclature ' + Valeur + ' de niveau ' + IntToStr(Level) +
               ' se référence elle même au niveau ' + IntToStr(tn1.Level) + ' sous ' +
               'la nomenclature ' + tn1.Parent.Text;
        ColorMemo1.Lines.Append(st1);
        end
        else
        Parcours(tn1, Valeur, Level);
    tn1 := tn1.GetNextSibling;
    end;

end;

procedure TFNomenErr.BitBtn1Click(Sender: TObject);
begin
Close;
end;

end.

unit TeletransFichierLB;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, StdCtrls, Mask, DBCtrls, HDB, Hctrls, HEnt1;

  Procedure LanceSaisieFichierLB (var S: string; Titre : string ; Sens : integer);

type
  TFTeletransFichierLB = class(TForm)
    HlblTeletrans: THLabel;
    BAnnuler: TToolbarButton97;
    BOuvrir: TToolbarButton97;
    SaveDialog1: TSaveDialog;
    OpenDialog1: TOpenDialog;
    FichierLB: THCritMaskEdit;
    procedure BAnnulerClick(Sender: TObject);
    procedure FichierLBElipsisClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
   private
    { Déclarations privées }
    Sens : integer;
  public
    { Déclarations publiques }
  end;

Const Reception = 0;
      Emission = 1;


implementation

{$R *.DFM}
Procedure LanceSaisieFichierLB (var S : string ; Titre : string; Sens : integer);
var  X : TFTeletransFichierLB;
BEGIN
SourisSablier;
X:=TFTeletransFichierLB.Create(Application) ;
  try
  X.FichierLb.text := S ;
  X.Sens := Sens;
  X.HlblTeletrans.caption := Titre ;
  X.ShowModal ;
  if  X.FichierLb.Text <> '' then S:= X.FichierLb.Text else S:= '' ;
  Finally
  X.Free ;
  end ;
SourisNormale ;

END ;

procedure TFTeletransFichierLB.FormShow(Sender: TObject);
begin
 Case Sens of
    Reception : HlblTeletrans.caption:='Nom du fichier à réceptionner ' ;
    Emission  : HlblTeletrans.caption:='Nom du fichier à envoyer ' ;
    end;
end;

procedure TFTeletransFichierLB.BAnnulerClick(Sender: TObject);
begin
FichierLb.Text := '';
end;

procedure TFTeletransFichierLB.FichierLBElipsisClick(Sender: TObject);
begin
Case Sens of
    Reception : if SaveDialog1.Execute then  FichierLB.Text := SaveDialog1.Filename ;
    Emission  : if OpenDialog1.Execute then  FichierLB.Text := OpenDialog1.Filename ;
    end;
end;

end.





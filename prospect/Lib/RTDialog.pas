unit RTDialog;

interface

uses
  Classes, Graphics, Controls, Forms, Dialogs,
  ExtDlgs;

type
  TRT_Dialog = class(TForm)
    CD_Param: TColorDialog;
    FD_Param: TFontDialog;
    SD_Param: TSavePictureDialog;

  public
    { Déclarations publiques }
  end;

function execRTColorDialog (Couleur : TColor) : TColor;
function execRTFontDialog (Fonte : TFont) : TFont;

implementation

{$R *.DFM}

function execRTColorDialog (Couleur : TColor) : TColor;
var
  vForm : TRT_Dialog;

begin
  vForm :=  TRT_Dialog.Create(Application);
  vForm.CD_Param.Color := Couleur;
  try
    if vForm.CD_Param.execute then
      result := vForm.CD_Param.Color
    else
      result := 0;
  finally
    vForm.Free;
  end;
end;

function execRTFontDialog (Fonte : TFont) : TFont;
var
  vForm : TRT_Dialog;
  vFont : TFont;
begin
  vForm :=  TRT_Dialog.Create(Application);
  vFont := TFont.Create;
  vForm.FD_Param.Font := Fonte;
  vForm.Font := Fonte;
  try
    if vForm.FD_Param.execute then
    begin
      vFont.name := vForm.FD_Param.Font.name;
      vFont.size := vForm.FD_Param.Font.size;
      vFont.style := vForm.FD_Param.Font.style;
      vFont.color := vForm.FD_Param.Font.color;
    end
    else
    begin
      vFont.name := vForm.Font.name;
      vFont.size := vForm.Font.size;
      vFont.style := vForm.Font.style;
      vFont.color := vForm.Font.color;
    end;
    result := vFont;
  finally
    vForm.Free;
  end;
end;


end.

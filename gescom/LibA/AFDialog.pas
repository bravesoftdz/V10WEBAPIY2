unit AFDialog;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtDlgs, extctrls;

type
  TAF_Dialog = class(TForm)
    CD_Param: TColorDialog;
    FD_Param: TFontDialog;
    SD_Param: TSavePictureDialog;

  public
    { Déclarations publiques }
  end;

function execAFColorDialog : TColor;
function execAFFontDialog : TFont;
function execAFSavePictureDialog : String;

implementation

{$R *.DFM}

function execAFColorDialog : TColor;
var
  vForm : TAF_Dialog;

begin
  vForm :=  TAF_Dialog.Create(Application);
  try
    if vForm.CD_Param.execute then
      result := vForm.CD_Param.Color
    else
      result := 0;
  finally
    vForm.Free;
  end;
end;

function execAFFontDialog : TFont;
var
  vForm : TAF_Dialog;
  vFont : TFont;
begin
  vForm :=  TAF_Dialog.Create(Application);
  vFont := TFont.Create;
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

function execAFSavePictureDialog : String;
var
  vForm : TAF_Dialog;
  //vPicture : TPicture;

begin
  vForm :=  TAF_Dialog.Create(Application);
  try
    if vForm.SD_Param.execute then
      result := vForm.SD_Param.FileName;
  finally
    vForm.Free;
  end;
end;

end.

object AF_Dialog: TAF_Dialog
  Left = 575
  Top = 325
  Width = 410
  Height = 231
  Caption = 'AF_Dialog'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object CD_Param: TColorDialog
    Ctl3D = True
    Left = 160
    Top = 40
  end
  object FD_Param: TFontDialog
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    MinFontSize = 0
    MaxFontSize = 0
    Left = 160
    Top = 92
  end
  object SD_Param: TSavePictureDialog
    Filter = 'Bitmaps (*.bmp)|*.bmp|Icônes (*.ico)|*.ico'
    Left = 216
    Top = 40
  end
end

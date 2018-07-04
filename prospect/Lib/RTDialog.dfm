object RT_Dialog: TRT_Dialog
  Left = 575
  Top = 325
  Width = 410
  Height = 231
  Caption = 'RT_Dialog'
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
    Left = 160
    Top = 40
  end
  object FD_Param: TFontDialog
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Left = 160
    Top = 92
  end
  object SD_Param: TSavePictureDialog
    Filter = 'Bitmaps (*.bmp)|*.bmp|Ic'#244'nes (*.ico)|*.ico'
    Left = 216
    Top = 40
  end
end

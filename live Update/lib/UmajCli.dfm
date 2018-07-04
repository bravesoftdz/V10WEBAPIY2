object FmajCli: TFmajCli
  Left = 457
  Top = 355
  Width = 587
  Height = 213
  Caption = 'Mise '#224' jour du poste'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 3
    Top = 16
    Width = 574
    Height = 20
    Alignment = taCenter
    AutoSize = False
    Caption = 'Progression de la copie'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object LProduct: TLabel
    Left = 0
    Top = 72
    Width = 577
    Height = 13
    Alignment = taCenter
    AutoSize = False
    Caption = 'XXX'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object ProgressCopie: TProgressBar
    Left = 40
    Top = 104
    Width = 505
    Height = 17
    Smooth = True
    TabOrder = 0
  end
  object XPManifest1: TXPManifest
    Left = 528
    Top = 48
  end
end

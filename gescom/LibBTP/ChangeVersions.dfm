object FMajVerBtp: TFMajVerBtp
  Left = 332
  Top = 264
  Width = 316
  Height = 161
  BorderIcons = [biHelp]
  Caption = 'V'#233'rification m'#233'tier'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  Visible = True
  PixelsPerInch = 96
  TextHeight = 13
  object Titre: THLabel
    Left = 0
    Top = 0
    Width = 300
    Height = 13
    Align = alTop
    Alignment = taCenter
    AutoSize = False
    Caption = 'XXX'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    Visible = False
    WordWrap = True
  end
  object PB: TEnhancedGauge
    Left = 23
    Top = 58
    Width = 265
    Height = 25
    Caption = 'PB'
    TabOrder = 0
    ForeColor = clNavy
    BackColor = clBtnFace
    Progress = 0
  end
end

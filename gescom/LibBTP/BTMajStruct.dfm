object fbtMajStructure: TfbtMajStructure
  Left = 526
  Top = 328
  Width = 323
  Height = 151
  Caption = 'Mise '#224' jour de structure'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Titre: THLabel
    Left = 0
    Top = 0
    Width = 307
    Height = 19
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
  object CURVUE: THLabel
    Left = 0
    Top = 99
    Width = 307
    Height = 13
    Align = alBottom
    Alignment = taCenter
    Caption = 'CURVUE'
    Visible = False
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

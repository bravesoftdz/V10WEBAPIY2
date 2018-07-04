object FBTRapport: TFBTRapport
  Left = 432
  Top = 241
  Width = 769
  Height = 468
  Caption = 'Rapport '
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object PBAS: TPanel
    Left = 0
    Top = 388
    Width = 753
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    object BClose: TToolbarButton97
      Left = 712
      Top = 8
      Width = 24
      Height = 24
      OnClick = BCloseClick
      GlobalIndexImage = 'M0002_S24G1'
    end
  end
  object MemoRapport: THMemo
    Left = 0
    Top = 0
    Width = 753
    Height = 388
    Align = alClient
    TabOrder = 1
  end
end

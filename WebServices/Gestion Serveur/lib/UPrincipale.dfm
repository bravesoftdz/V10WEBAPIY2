object Fprincipale: TFprincipale
  Left = 318
  Top = 228
  Width = 600
  Height = 425
  Caption = 'Autorisation d'#39'appels des WebServices L.S.E'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PBAS: TPanel
    Left = 0
    Top = 353
    Width = 584
    Height = 34
    Align = alBottom
    BevelOuter = bvNone
    Ctl3D = False
    ParentCtl3D = False
    TabOrder = 0
    DesignSize = (
      584
      34)
    object BANNULE: TToolbarButton97
      Left = 541
      Top = 2
      Width = 27
      Height = 27
      Anchors = [akRight, akBottom]
      OnClick = BANNULEClick
      GlobalIndexImage = 'Z0021_S24G1'
    end
    object BNEW: TToolbarButton97
      Left = 32
      Top = 3
      Width = 27
      Height = 27
      OnClick = BNEWClick
      GlobalIndexImage = 'Z0053_S24G1'
    end
  end
  object GS: THGrid
    Left = 0
    Top = 0
    Width = 584
    Height = 353
    Align = alClient
    DefaultRowHeight = 18
    RowCount = 3
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
    TabOrder = 1
    OnDblClick = GSDblClick
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = clSilver
  end
  object HmTrad: THSystemMenu
    Separator = True
    Traduction = False
    Left = 424
    Top = 40
  end
end

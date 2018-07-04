object FEncTiers: TFEncTiers
  Left = 256
  Top = 174
  Width = 466
  Height = 279
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Liste des '#233'ch'#233'ances du tiers'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object G: THGrid
    Tag = 1
    Left = 0
    Top = 0
    Width = 458
    Height = 219
    Align = alClient
    Ctl3D = True
    DefaultColWidth = 50
    DefaultRowHeight = 18
    FixedCols = 0
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
    ParentCtl3D = False
    TabOrder = 0
    OnDblClick = GDblClick
    OnMouseUp = GMouseUp
    SortedCol = -1
    Couleur = True
    MultiSelect = True
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = 13224395
  end
  object P: TPanel
    Left = 0
    Top = 219
    Width = 458
    Height = 33
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 1
    object H_TOTDEVISE: TLabel
      Left = 169
      Top = 10
      Width = 24
      Height = 13
      Caption = 'Total'
    end
    object BValider: THBitBtn
      Left = 364
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Valider'
      Default = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BValiderClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0003_S16G2'
      IsControl = True
    end
    object BFerme: THBitBtn
      Left = 396
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      ModalResult = 2
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      GlobalIndexImage = 'Z0021_S16G1'
    end
    object BAide: THBitBtn
      Left = 428
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Aide'
      HelpContext = 1710
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z1117_S16G1'
      IsControl = True
    end
    object ER_SOLDED: THNumEdit
      Tag = 2
      Left = 200
      Top = 7
      Width = 121
      Height = 21
      Color = clBtnFace
      Ctl3D = True
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      NumericType = ntDC
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 3
      UseRounding = True
      Validate = False
    end
    object bZoomPiece: THBitBtn
      Left = 332
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Voir '#233'criture'
      HelpContext = 1710
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 4
      OnClick = bZoomPieceClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0061_S16G1'
      IsControl = True
    end
    object bModifEche: THBitBtn
      Left = 3
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Modification de l'#39#233'ch'#233'ance'
      HelpContext = 1710
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 5
      Visible = False
      OnClick = bModifEcheClick
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      GlobalIndexImage = 'Z0357_S16G2'
      IsControl = True
    end
  end
end

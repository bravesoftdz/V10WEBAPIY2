object FSaisBase: TFSaisBase
  Left = 210
  Top = 188
  Width = 469
  Height = 323
  HelpContext = 7646100
  BorderIcons = [biSystemMenu]
  Caption = 'Bases Hors-taxes'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 328
    Top = 148
    Width = 42
    Height = 13
    AutoSize = False
    Caption = 'Total HT'
  end
  object Label2: TLabel
    Left = 328
    Top = 180
    Width = 48
    Height = 13
    AutoSize = False
    Caption = 'Total TTC'
  end
  object BCalcHT: TToolbarButton97
    Left = 338
    Top = 112
    Width = 23
    Height = 23
    Hint = 'Calcul du HT'
    Layout = blGlyphRight
    NumGlyphs = 2
    ParentShowHint = False
    ShowHint = True
    OnClick = BCalcHTClick
    GlobalIndexImage = 'Z0051_S16G2'
  end
  object H_TVA: THLabel
    Left = 328
    Top = 212
    Width = 49
    Height = 13
    Caption = 'Code TVA'
    FocusControl = E_TVA
    Visible = False
  end
  object POutils: TPanel
    Left = 0
    Top = 256
    Width = 461
    Height = 33
    Align = alBottom
    BevelInner = bvLowered
    BevelOuter = bvNone
    TabOrder = 0
    object BValide: THBitBtn
      Left = 364
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Valider'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BValideClick
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      GlobalIndexImage = 'Z0003_S16G2'
      IsControl = True
    end
    object BAbandon: THBitBtn
      Left = 396
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BAbandonClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0021_S16G1'
      IsControl = True
    end
    object BAide: THBitBtn
      Left = 428
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Aide'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = BAideClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z1117_S16G1'
      IsControl = True
    end
  end
  object G: THGrid
    Left = 12
    Top = 112
    Width = 313
    Height = 137
    ColCount = 4
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 6
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goEditing, goTabs]
    ScrollBars = ssNone
    TabOrder = 1
    OnExit = GExit
    SortedCol = -1
    Titres.Strings = (
      'Base'
      'Code TVA'
      'Taux TVA'
      'Base HT')
    Couleur = False
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    OnCellEnter = GCellEnter
    OnCellExit = GCellExit
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = 13224395
    ColWidths = (
      85
      61
      67
      93)
  end
  object GEche: THGrid
    Left = 12
    Top = 16
    Width = 397
    Height = 79
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 10
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goTabs, goRowSelect]
    ScrollBars = ssVertical
    TabOrder = 2
    SortedCol = -1
    Titres.Strings = (
      'N'#176
      'Ech'#233'ance'
      'Mode paiement'
      'Montant TTC'
      'Total HT')
    Couleur = False
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    OnRowEnter = GEcheRowEnter
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = 13224395
    ColWidths = (
      22
      69
      110
      86
      85)
  end
  object TotHT: THNumEdit
    Left = 380
    Top = 144
    Width = 75
    Height = 21
    Decimals = 2
    Digits = 12
    Enabled = False
    Masks.PositiveMask = '#,##0'
    Debit = False
    TabOrder = 3
    UseRounding = True
    Validate = False
  end
  object TotTTC: THNumEdit
    Left = 380
    Top = 176
    Width = 75
    Height = 21
    Decimals = 2
    Digits = 12
    Enabled = False
    Masks.PositiveMask = '#,##0'
    Debit = False
    TabOrder = 4
    UseRounding = True
    Validate = False
  end
  object E_TVA: THValComboBox
    Left = 380
    Top = 208
    Width = 75
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 5
    Visible = False
    TagDispatch = 0
    DataType = 'TTTVA'
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 184
    Top = 168
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Encais 1'
      'Encais 2'
      'Encais 3'
      'Encais 4'
      'D'#233'bit'
      
        '5;Bases Hors-taxes;Le total HT d'#233'passe le total TTC. Confirmez-v' +
        'ous la modification ?;W;YN;N;N;')
    Left = 135
    Top = 171
  end
end

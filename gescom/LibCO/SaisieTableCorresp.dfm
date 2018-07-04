object SaisieTable: TSaisieTable
  Left = 162
  Top = 147
  Width = 538
  Height = 610
  Caption = 'Saisie des tables de correspondances'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object FermerParamFicBase: TButton
    Left = 447
    Top = 554
    Width = 75
    Height = 25
    Caption = '&Fermer'
    TabOrder = 0
    OnClick = FermerParamFicBaseClick
  end
  object GroupBox4: TGroupBox
    Left = 5
    Top = 139
    Width = 516
    Height = 143
    Caption = 'Modes de règlement '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
    object GridReg: THGrid
      Tag = 1
      Left = 2
      Top = 18
      Width = 512
      Height = 123
      Align = alBottom
      ColCount = 3
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 30
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goEditing, goTabs, goAlwaysShowEditor]
      TabOrder = 0
      SortedCol = -1
      Titres.Strings = (
        'Code GB 2000'
        'Libellé'
        'Code PGI'
        '')
      Couleur = False
      MultiSelect = False
      TitleBold = True
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = 13224395
      ColWidths = (
        116
        183
        190)
      RowHeights = (
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18)
    end
  end
  object GroupBox5: TGroupBox
    Left = 3
    Top = 420
    Width = 520
    Height = 127
    Caption = 'Pays'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 2
    object GridPays: THGrid
      Tag = 1
      Left = 2
      Top = 16
      Width = 516
      Height = 109
      Align = alBottom
      ColCount = 3
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 30
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goEditing, goTabs, goAlwaysShowEditor]
      TabOrder = 0
      SortedCol = -1
      Titres.Strings = (
        'Code GB 2000'
        'Libellé '
        'Code PGI'
        '')
      Couleur = False
      MultiSelect = False
      TitleBold = True
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = 13224395
      ColWidths = (
        116
        187
        190)
      RowHeights = (
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18)
    end
  end
  object GroupBox1: TGroupBox
    Left = 5
    Top = 286
    Width = 519
    Height = 127
    Caption = 'Devises '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 3
    object GridDevise: THGrid
      Tag = 1
      Left = 2
      Top = 18
      Width = 515
      Height = 107
      Align = alBottom
      ColCount = 3
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 30
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goEditing, goTabs, goAlwaysShowEditor]
      TabOrder = 0
      SortedCol = -1
      Titres.Strings = (
        'Code GB 2000'
        'Libellé '
        'Code PGI'
        '')
      Couleur = False
      MultiSelect = False
      TitleBold = True
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = 13224395
      ColWidths = (
        116
        187
        189)
      RowHeights = (
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18)
    end
  end
  object GroupBox2: TGroupBox
    Left = 3
    Top = 6
    Width = 518
    Height = 131
    Caption = 'Familles '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 4
    object GridFamille: THGrid
      Tag = 1
      Left = 2
      Top = 15
      Width = 514
      Height = 114
      Align = alBottom
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 30
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goEditing, goTabs, goAlwaysShowEditor]
      TabOrder = 0
      SortedCol = -1
      Titres.Strings = (
        'Code GB 2000'
        'Libellé '
        'FAM 1'
        'FAM 2'
        'FAM 3'
        '')
      Couleur = False
      MultiSelect = False
      TitleBold = True
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = 13224395
      ColWidths = (
        116
        185
        67
        62
        58)
      RowHeights = (
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18)
    end
  end
end

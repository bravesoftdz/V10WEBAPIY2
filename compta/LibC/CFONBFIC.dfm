object FCFONBFIC: TFCFONBFIC
  Left = 188
  Top = 135
  Width = 625
  Height = 465
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Visualisation du fichier d'#39'export CFONB'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object pnl_Main: THPanel
    Left = 0
    Top = 41
    Width = 617
    Height = 362
    Align = alClient
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object G: THGrid
      Left = 1
      Top = 1
      Width = 615
      Height = 360
      Align = alClient
      Ctl3D = True
      DefaultColWidth = 18
      DefaultRowHeight = 16
      RowCount = 2
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'Courier'
      Font.Style = []
      Options = [goFixedVertLine, goFixedHorzLine, goHorzLine, goRowSizing, goColSizing, goRowSelect, goThumbTracking]
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 0
      SortedCol = -1
      Titres.Strings = (
        'N'#176
        'Informations 1'
        'Informations 2'
        'Informations 3'
        'Informations 4')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = 13224395
      ColWidths = (
        18
        127
        116
        116
        130)
    end
  end
  object Outils: TPanel
    Left = 0
    Top = 403
    Width = 617
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    Ctl3D = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentCtl3D = False
    ParentFont = False
    ParentShowHint = False
    ShowHint = False
    TabOrder = 1
    object Panel1: TPanel
      Left = 484
      Top = 2
      Width = 131
      Height = 31
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 0
      object BAide: THBitBtn
        Left = 100
        Top = 2
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
        TabOrder = 0
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BFerme: THBitBtn
        Left = 68
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = BFermeClick
        Glyph.Data = {
          CE070000424DCE07000000000000360000002800000024000000120000000100
          1800000000009807000000000000000000000000000000000000008284008284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284008284008284008284008284008284008284008284
          0082840082840082840082840082840082848482848482840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284FFFFFF008284008284008284008284008284008284
          0082840082840082840082840082840082840082840082840082840000FF0000
          840000848482840082840082840082840082840082840000FF84828400828400
          8284008284008284008284008284008284008284848284848284FFFFFF008284
          008284008284008284008284008284FFFFFF0082840082840082840082840082
          840082840082840000FF00008400008400008484828400828400828400828400
          00FF000084000084848284008284008284008284008284008284008284848284
          FFFFFF008284848284FFFFFF008284008284008284FFFFFF848284848284FFFF
          FF0082840082840082840082840082840082840000FF00008400008400008400
          00848482840082840000FF000084000084000084000084848284008284008284
          008284008284008284848284FFFFFF008284008284848284FFFFFF008284FFFF
          FF848284008284008284848284FFFFFF00828400828400828400828400828400
          82840000FF000084000084000084000084848284000084000084000084000084
          000084848284008284008284008284008284008284848284FFFFFF0082840082
          84008284848284FFFFFF848284008284008284008284008284848284FFFFFF00
          82840082840082840082840082840082840000FF000084000084000084000084
          0000840000840000840000848482840082840082840082840082840082840082
          84008284848284FFFFFF00828400828400828484828400828400828400828400
          8284FFFFFF848284008284008284008284008284008284008284008284008284
          0000FF0000840000840000840000840000840000848482840082840082840082
          84008284008284008284008284008284008284848284FFFFFF00828400828400
          8284008284008284008284FFFFFF848284008284008284008284008284008284
          0082840082840082840082840082840000840000840000840000840000848482
          8400828400828400828400828400828400828400828400828400828400828400
          8284848284FFFFFF008284008284008284008284008284848284008284008284
          0082840082840082840082840082840082840082840082840082840000FF0000
          8400008400008400008484828400828400828400828400828400828400828400
          8284008284008284008284008284008284848284FFFFFF008284008284008284
          8482840082840082840082840082840082840082840082840082840082840082
          840082840000FF00008400008400008400008400008484828400828400828400
          8284008284008284008284008284008284008284008284008284008284848284
          008284008284008284008284848284FFFFFF0082840082840082840082840082
          840082840082840082840082840000FF00008400008400008484828400008400
          0084000084848284008284008284008284008284008284008284008284008284
          008284008284848284008284008284008284008284008284848284FFFFFF0082
          840082840082840082840082840082840082840082840000FF00008400008400
          00848482840082840000FF000084000084000084848284008284008284008284
          008284008284008284008284008284848284008284008284008284848284FFFF
          FF008284008284848284FFFFFF00828400828400828400828400828400828400
          82840000FF0000840000848482840082840082840082840000FF000084000084
          000084848284008284008284008284008284008284008284848284FFFFFF0082
          84008284848284008284848284FFFFFF008284008284848284FFFFFF00828400
          82840082840082840082840082840082840000FF000084008284008284008284
          0082840082840000FF0000840000840000840082840082840082840082840082
          84008284848284FFFFFFFFFFFF848284008284008284008284848284FFFFFF00
          8284008284848284FFFFFF008284008284008284008284008284008284008284
          0082840082840082840082840082840082840082840000FF0000840000FF0082
          8400828400828400828400828400828400828484828484828400828400828400
          8284008284008284848284FFFFFFFFFFFFFFFFFF848284008284008284008284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284008284008284008284008284848284848284848284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284008284008284008284008284008284008284008284
          008284008284008284008284008284008284}
        NumGlyphs = 2
      end
      object BValider: THBitBtn
        Left = 36
        Top = 2
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
        TabOrder = 2
        OnClick = BValiderClick
        Margin = 2
        NumGlyphs = 2
        Spacing = -1
        GlobalIndexImage = 'Z0127_S16G1'
        IsControl = True
      end
    end
  end
  object PEntete: TPanel
    Left = 0
    Top = 0
    Width = 617
    Height = 41
    Align = alTop
    TabOrder = 2
    object HTitre: TLabel
      Left = 8
      Top = 12
      Width = 271
      Height = 16
      Caption = 'Compte rendu d'#39'exportation CFONB.    Fichier '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
  end
  object pnl_IBAN: THPanel
    Left = 0
    Top = 41
    Width = 617
    Height = 362
    Align = alClient
    FullRepaint = False
    TabOrder = 3
    Visible = False
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object HSplitter1: THSplitter
      Left = 1
      Top = 86
      Width = 615
      Height = 3
      Cursor = crVSplit
      Align = alTop
      Visible = True
    end
    object HSplitter2: THSplitter
      Left = 1
      Top = 174
      Width = 615
      Height = 3
      Cursor = crVSplit
      Align = alTop
      Visible = True
    end
    object HSplitter3: THSplitter
      Left = 1
      Top = 258
      Width = 615
      Height = 3
      Cursor = crVSplit
      Align = alTop
      Visible = True
    end
    object G2: THGrid
      Left = 1
      Top = 89
      Width = 615
      Height = 85
      Align = alTop
      ColCount = 26
      DefaultColWidth = 18
      DefaultRowHeight = 16
      RowCount = 2
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Options = [goFixedVertLine, goFixedHorzLine, goHorzLine, goRowSizing, goColSizing, goRowSelect, goThumbTracking]
      ParentFont = False
      TabOrder = 0
      SortedCol = -1
      Titres.Strings = (
        ''
        'Code'
        'Op'#233'ration'
        'N'#176
        'Identifiant'
        'Compte b'#233'n'#233'ficiaire'
        'Nom b'#233'n'#233'ficiaire'
        'Adresse'
        'Ident. Nat. b'#233'n'#233'ficiaire'
        'Pays'
        'R'#233'f'#233'rence op'#233'ration'
        'Qualifiant'
        'R'#233'serv'#233
        'Montant'
        'D'#233'cimales'
        'R'#233'serv'#233
        'Motif '#233'conomique'
        'Code pays BDF'
        'Mode r'#233'glement'
        'Code imputation'
        'Identifiant compte de frais'
        'Compte de frais'
        'Code devise compte de frais'
        'R'#233'serv'#233
        'Date ex'#233'cution'
        'Code devise transfert')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = 13224395
      ColWidths = (
        18
        41
        58
        24
        65
        99
        105
        60
        115
        42
        120
        77
        65
        56
        70
        63
        103
        96
        94
        92
        128
        99
        153
        70
        106
        118)
    end
    object G3: THGrid
      Left = 1
      Top = 177
      Width = 615
      Height = 81
      Align = alTop
      ColCount = 9
      DefaultColWidth = 18
      DefaultRowHeight = 16
      RowCount = 2
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Options = [goFixedVertLine, goFixedHorzLine, goHorzLine, goRowSizing, goColSizing, goRowSelect, goThumbTracking]
      ParentFont = False
      TabOrder = 1
      SortedCol = -1
      Titres.Strings = (
        ''
        'Code'
        'Op'#233'ration'
        'N'#176
        'Banque b'#233'n'#233'ficiaire'
        'Agence'
        'Code BIC'
        'Code pays'
        'R'#233'serv'#233)
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = 13224395
      ColWidths = (
        18
        41
        58
        25
        118
        79
        103
        111
        70)
    end
    object G1: THGrid
      Left = 1
      Top = 1
      Width = 615
      Height = 85
      Align = alTop
      ColCount = 22
      DefaultColWidth = 18
      DefaultRowHeight = 16
      RowCount = 2
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Options = [goFixedVertLine, goFixedHorzLine, goHorzLine, goRowSizing, goColSizing, goRowSelect, goThumbTracking]
      ParentFont = False
      TabOrder = 2
      SortedCol = -1
      Titres.Strings = (
        ''
        'Code'
        'Op'#233'ration'
        'N'#176
        'Date cr'#233'ation'
        'Raison sociale'
        'Adresse'
        'N'#176' SIRET'
        'R'#233'f'#233'rence remise'
        'Code BIC'
        'Identifiant'
        'Compte '#224' d'#233'biter'
        'Code devise'
        'Identifiant client'
        'Identifiant compte de frais'
        'Compte de frais'
        'Code devise compte de frais'
        'R'#233'serv'#233
        'Indice type d'#233'bit remise'
        'Indice type remise'
        'Date ex'#233'cution'
        'Code devise transfert ')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = 13224395
      ColWidths = (
        18
        41
        58
        24
        75
        99
        105
        60
        115
        67
        120
        97
        77
        99
        143
        122
        162
        96
        94
        92
        128
        116)
    end
    object G4: THGrid
      Left = 1
      Top = 261
      Width = 615
      Height = 100
      Align = alClient
      ColCount = 15
      DefaultColWidth = 18
      DefaultRowHeight = 16
      RowCount = 2
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Options = [goFixedVertLine, goFixedHorzLine, goHorzLine, goRowSizing, goColSizing, goRowSelect, goThumbTracking]
      ParentFont = False
      TabOrder = 3
      SortedCol = -1
      Titres.Strings = (
        ''
        'Code'
        'Op'#233'ration'
        'N'#176
        'Date cr'#233'ation'
        'R'#233'serv'#233
        'N'#176' SIRET '#233'metteur'
        'R'#233'f'#233'rence remise'
        'R'#233'serv'#233
        'Identifiant'
        'Compte '#224' d'#233'biter'
        'Code devise'
        'Identifiant client'
        'Total de contr'#244'le'
        'R'#233'serv'#233)
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = 13224395
      ColWidths = (
        18
        41
        58
        25
        78
        72
        103
        103
        70
        71
        98
        87
        101
        104
        99)
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 514
    Top = 72
  end
end

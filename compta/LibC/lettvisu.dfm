object FLettVisu: TFLettVisu
  Left = 246
  Top = 163
  HelpContext = 7508200
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Visualisation des solutions du rapprochement'
  ClientHeight = 365
  ClientWidth = 682
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  PopupMenu = POPS
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object POutils: TPanel
    Left = 0
    Top = 333
    Width = 682
    Height = 32
    Align = alBottom
    TabOrder = 0
    object H_NBSOL: TLabel
      Left = 8
      Top = 8
      Width = 160
      Height = 13
      Caption = 'Solutions trouv'#233'es (10 maximum) :'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object H_NUMSOL: TLabel
      Left = 220
      Top = 8
      Width = 56
      Height = 13
      Caption = 'Solution N'#176' '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object E_NBSOL: TLabel
      Left = 176
      Top = 8
      Width = 8
      Height = 13
      Caption = '1'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object E_NUMSOL: TLabel
      Left = 280
      Top = 8
      Width = 8
      Height = 13
      Caption = '1'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object BValideSelect: TToolbarButton97
      Tag = 1
      Left = 512
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Choisir la solution courante'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Glyph.Data = {
        BE060000424DBE06000000000000360400002800000024000000120000000100
        0800000000008802000000000000000000000001000000010000000000000000
        80000080000000808000800000008000800080800000C0C0C000C0DCC000F0C8
        A400000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        000000000000000000000000000000000000F0FBFF00A4A0A000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00030303030303
        0303030303030303030303030303030303030303030303030303030303030303
        03030303030303030303030303030303030303030303FF030303030303030303
        03030303030303040403030303030303030303030303030303F8F8FF03030303
        03030303030303030303040202040303030303030303030303030303F80303F8
        FF030303030303030303030303040202020204030303030303030303030303F8
        03030303F8FF0303030303030303030304020202020202040303030303030303
        0303F8030303030303F8FF030303030303030304020202FA0202020204030303
        0303030303F8FF0303F8FF030303F8FF03030303030303020202FA03FA020202
        040303030303030303F8FF03F803F8FF0303F8FF03030303030303FA02FA0303
        03FA0202020403030303030303F8FFF8030303F8FF0303F8FF03030303030303
        FA0303030303FA0202020403030303030303F80303030303F8FF0303F8FF0303
        0303030303030303030303FA0202020403030303030303030303030303F8FF03
        03F8FF03030303030303030303030303FA020202040303030303030303030303
        0303F8FF0303F8FF03030303030303030303030303FA02020204030303030303
        03030303030303F8FF0303F8FF03030303030303030303030303FA0202020403
        030303030303030303030303F8FF0303F8FF03030303030303030303030303FA
        0202040303030303030303030303030303F8FF03F8FF03030303030303030303
        03030303FA0202030303030303030303030303030303F8FFF803030303030303
        030303030303030303FA0303030303030303030303030303030303F803030303
        0303030303030303030303030303030303030303030303030303030303030303
        0303}
      GlyphMask.Data = {00000000}
      Margin = 2
      NumGlyphs = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      OnClick = BValideSelectClick
      IsControl = True
    end
    object BAbandon: TToolbarButton97
      Tag = 1
      Left = 540
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Margin = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      OnClick = BAbandonClick
      GlobalIndexImage = 'Z0021_S16G1'
      IsControl = True
    end
    object BAide: TToolbarButton97
      Tag = 1
      Left = 568
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Aide'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Margin = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      OnClick = BAideClick
      GlobalIndexImage = 'Z1117_S16G1'
      IsControl = True
    end
    object BFirst: TToolbarButton97
      Tag = 1
      Left = 324
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Premi'#232're solution'
      ParentShowHint = False
      ShowHint = True
      OnClick = BFirstClick
      GlobalIndexImage = 'Z0301_S16G1'
    end
    object BPrev: TToolbarButton97
      Tag = 1
      Left = 352
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Solution pr'#233'c'#233'dente'
      ParentShowHint = False
      ShowHint = True
      OnClick = BPrevClick
      GlobalIndexImage = 'Z0017_S16G1'
    end
    object BNext: TToolbarButton97
      Tag = 1
      Left = 380
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Solution suivante'
      ParentShowHint = False
      ShowHint = True
      OnClick = BNextClick
      GlobalIndexImage = 'Z0031_S16G1'
    end
    object BLast: TToolbarButton97
      Tag = 1
      Left = 408
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Derni'#232're solution'
      ParentShowHint = False
      ShowHint = True
      OnClick = BLastClick
      GlobalIndexImage = 'Z0264_S16G1'
    end
    object BZoomPiece: TToolbarButton97
      Tag = 1
      Left = 440
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Voir '#233'criture'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Margin = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      OnClick = BZoomPieceClick
      GlobalIndexImage = 'Z0061_S16G1'
      IsControl = True
    end
  end
  object PEntete: TPanel
    Left = 0
    Top = 0
    Width = 682
    Height = 23
    Align = alTop
    TabOrder = 1
    object H_COMPTE: THLabel
      Left = 4
      Top = 3
      Width = 589
      Height = 18
      Alignment = taCenter
      AutoSize = False
      Caption = 'Comptes et Intitul'#233's'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object GD: THGrid
    Tag = 1
    Left = 0
    Top = 23
    Width = 305
    Height = 197
    Align = alLeft
    Ctl3D = True
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 11
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
    ParentCtl3D = False
    ParentFont = False
    ScrollBars = ssVertical
    TabOrder = 2
    OnEnter = GDEnter
    SortedCol = -1
    Couleur = True
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    OnRowEnter = MontreDetail
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = 13224395
    ColWidths = (
      64
      64
      64
      64
      64)
  end
  object PSepare: TPanel
    Left = 305
    Top = 23
    Width = 8
    Height = 197
    Align = alLeft
    BevelOuter = bvNone
    Color = clWhite
    Enabled = False
    TabOrder = 3
  end
  object GC: THGrid
    Tag = 1
    Left = 313
    Top = 23
    Width = 369
    Height = 197
    Align = alClient
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 11
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
    ScrollBars = ssVertical
    TabOrder = 4
    OnEnter = GDEnter
    SortedCol = -1
    Couleur = True
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    OnRowEnter = MontreDetail
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = 13224395
  end
  object CJOURNAL: THValComboBox
    Left = 8
    Top = 112
    Width = 57
    Height = 21
    ItemHeight = 13
    TabOrder = 5
    Visible = False
    TagDispatch = 0
    DataType = 'TTJOURNAL'
  end
  object CNATUREPIECE: THValComboBox
    Left = 68
    Top = 112
    Width = 57
    Height = 21
    ItemHeight = 13
    TabOrder = 6
    Visible = False
    TagDispatch = 0
    DataType = 'TTNATUREPIECE'
  end
  object CMODEPAIE: THValComboBox
    Left = 128
    Top = 112
    Width = 57
    Height = 21
    ItemHeight = 13
    TabOrder = 7
    Visible = False
    TagDispatch = 0
    DataType = 'TTMODEPAIE'
  end
  object CDevise: THValComboBox
    Left = 192
    Top = 112
    Width = 57
    Height = 21
    ItemHeight = 13
    TabOrder = 8
    Visible = False
    TagDispatch = 0
  end
  object PCDetail: TPageControl
    Left = 0
    Top = 220
    Width = 682
    Height = 113
    ActivePage = TSDetail
    Align = alBottom
    TabOrder = 9
    object TSDetail: TTabSheet
      Caption = 'D'#233'tail'
      object H_JOURNAL: TLabel
        Left = 4
        Top = 4
        Width = 34
        Height = 13
        AutoSize = False
        Caption = 'Journal'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object E_JOURNAL: TLabel
        Tag = 1
        Left = 56
        Top = 4
        Width = 96
        Height = 13
        AutoSize = False
        Caption = 'E_JOURNAL'
      end
      object H_NUMEROPIECE: TLabel
        Left = 4
        Top = 20
        Width = 37
        Height = 13
        AutoSize = False
        Caption = 'Num'#233'ro'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object E_NUMEROPIECE: TLabel
        Tag = 1
        Left = 56
        Top = 20
        Width = 96
        Height = 13
        AutoSize = False
        Caption = 'E_NUMEROPIECE'
      end
      object H_DATEECHEANCE: TLabel
        Left = 4
        Top = 36
        Width = 49
        Height = 13
        AutoSize = False
        Caption = 'Ech'#233'ance'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object E_DATEECHEANCE: TLabel
        Tag = 1
        Left = 56
        Top = 36
        Width = 96
        Height = 13
        AutoSize = False
        Caption = 'E_DATEECHEANCE'
      end
      object H_REFLIBRE: TLabel
        Left = 4
        Top = 52
        Width = 42
        Height = 13
        AutoSize = False
        Caption = 'Ref. libre'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object E_REFLIBRE: TLabel
        Tag = 1
        Left = 56
        Top = 52
        Width = 96
        Height = 13
        AutoSize = False
        Caption = 'E_REFLIBRE'
      end
      object H_NATUREPIECE: TLabel
        Left = 156
        Top = 4
        Width = 32
        Height = 13
        AutoSize = False
        Caption = 'Nature'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object E_NATUREPIECE: TLabel
        Tag = 1
        Left = 200
        Top = 4
        Width = 96
        Height = 13
        AutoSize = False
        Caption = 'E_NATUREPIECE'
      end
      object E_LIBELLE: TLabel
        Tag = 1
        Left = 200
        Top = 20
        Width = 96
        Height = 13
        AutoSize = False
        Caption = 'E_LIBELLE'
      end
      object H_LIBELLE: TLabel
        Left = 156
        Top = 20
        Width = 30
        Height = 13
        AutoSize = False
        Caption = 'Libell'#233
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object H_MODEPAIE: TLabel
        Left = 156
        Top = 36
        Width = 24
        Height = 13
        AutoSize = False
        Caption = 'Paie.'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object E_MODEPAIE: TLabel
        Tag = 1
        Left = 200
        Top = 36
        Width = 96
        Height = 13
        AutoSize = False
        Caption = 'E_MODEPAIE'
      end
      object H_TOTAL: TLabel
        Left = 156
        Top = 52
        Width = 39
        Height = 13
        AutoSize = False
        Caption = 'Montant'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object BSepare: TBevel
        Left = 300
        Top = 1
        Width = 2
        Height = 89
        Shape = bsLeftLine
      end
      object H_JOURNAL_: TLabel
        Left = 304
        Top = 4
        Width = 34
        Height = 13
        AutoSize = False
        Caption = 'Journal'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object H_NUMEROPIECE_: TLabel
        Left = 304
        Top = 20
        Width = 37
        Height = 13
        AutoSize = False
        Caption = 'Num'#233'ro'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object H_DATEECHEANCE_: TLabel
        Left = 304
        Top = 36
        Width = 49
        Height = 13
        AutoSize = False
        Caption = 'Ech'#233'ance'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object H_REFLIBRE_: TLabel
        Left = 304
        Top = 52
        Width = 42
        Height = 13
        AutoSize = False
        Caption = 'Ref. libre'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object E_REFLIBRE_: TLabel
        Tag = 1
        Left = 356
        Top = 52
        Width = 96
        Height = 13
        AutoSize = False
        Caption = 'E_REFLIBRE'
      end
      object E_DATEECHEANCE_: TLabel
        Tag = 1
        Left = 356
        Top = 36
        Width = 96
        Height = 13
        AutoSize = False
        Caption = 'E_DATEECHEANCE'
      end
      object E_NUMEROPIECE_: TLabel
        Tag = 1
        Left = 356
        Top = 20
        Width = 96
        Height = 13
        AutoSize = False
        Caption = 'E_NUMEROPIECE'
      end
      object E_JOURNAL_: TLabel
        Tag = 1
        Left = 356
        Top = 4
        Width = 96
        Height = 13
        AutoSize = False
        Caption = 'E_JOURNAL'
      end
      object H_NATUREPIECE_: TLabel
        Left = 456
        Top = 4
        Width = 32
        Height = 13
        AutoSize = False
        Caption = 'Nature'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object H_LIBELLE_: TLabel
        Left = 456
        Top = 20
        Width = 30
        Height = 13
        AutoSize = False
        Caption = 'Libell'#233
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object H_MODEPAIE_: TLabel
        Left = 456
        Top = 36
        Width = 24
        Height = 13
        AutoSize = False
        Caption = 'Paie.'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object H_TOTAL_: TLabel
        Left = 456
        Top = 52
        Width = 39
        Height = 13
        AutoSize = False
        Caption = 'Montant'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object E_MODEPAIE_: TLabel
        Tag = 1
        Left = 500
        Top = 36
        Width = 96
        Height = 13
        AutoSize = False
        Caption = 'E_MODEPAIE'
      end
      object E_LIBELLE_: TLabel
        Tag = 1
        Left = 500
        Top = 20
        Width = 96
        Height = 13
        AutoSize = False
        Caption = 'E_LIBELLE'
      end
      object E_NATUREPIECE_: TLabel
        Tag = 1
        Left = 500
        Top = 4
        Width = 96
        Height = 13
        AutoSize = False
        Caption = 'E_NATUREPIECE'
      end
      object E_TOTAL: THNumEdit
        Tag = 1
        Left = 200
        Top = 50
        Width = 97
        Height = 21
        Color = clBtnFace
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Masks.PositiveMask = '#,##0'
        Debit = False
        ParentFont = False
        TabOrder = 0
        UseRounding = True
        Validate = False
      end
      object E_TOTAL_: THNumEdit
        Tag = 1
        Left = 500
        Top = 50
        Width = 97
        Height = 21
        Color = clBtnFace
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Masks.PositiveMask = '#,##0'
        Debit = False
        ParentFont = False
        TabOrder = 1
        UseRounding = True
        Validate = False
      end
    end
    object TSDetail1: TTabSheet
      Caption = 'D'#233'tail'
      ImageIndex = 1
      object Label1: TLabel
        Left = 3
        Top = 4
        Width = 34
        Height = 13
        AutoSize = False
        Caption = 'Nature'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object e_naturepcl: TLabel
        Tag = 1
        Left = 48
        Top = 4
        Width = 113
        Height = 13
        AutoSize = False
        Caption = 'E_JOURNAL'
      end
      object Label3: TLabel
        Left = 3
        Top = 19
        Width = 42
        Height = 13
        AutoSize = False
        Caption = 'Relev'#233
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object e_refrelevepcl: TLabel
        Tag = 1
        Left = 48
        Top = 19
        Width = 105
        Height = 13
        AutoSize = False
        Caption = 'E_JOURNAL'
      end
      object Label7: TLabel
        Left = 3
        Top = 35
        Width = 73
        Height = 14
        AutoSize = False
        Caption = 'Affaire'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object e_affairepcl: TLabel
        Tag = 1
        Left = 48
        Top = 35
        Width = 113
        Height = 13
        AutoSize = False
        Caption = 'E_JOURNAL'
      end
      object Label2: TLabel
        Left = 170
        Top = 4
        Width = 49
        Height = 13
        AutoSize = False
        Caption = 'Ech'#233'ance'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object e_dateecheancepcl: TLabel
        Tag = 1
        Left = 267
        Top = 4
        Width = 130
        Height = 13
        AutoSize = False
        Caption = 'E_JOURNAL'
      end
      object Label4: TLabel
        Left = 170
        Top = 19
        Width = 89
        Height = 13
        AutoSize = False
        Caption = 'R'#233'f'#233'rence libre'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object e_reflibrepcl: TLabel
        Tag = 1
        Left = 267
        Top = 19
        Width = 130
        Height = 13
        AutoSize = False
        Caption = 'E_JOURNAL'
      end
      object Label8: TLabel
        Left = 170
        Top = 35
        Width = 97
        Height = 14
        AutoSize = False
        Caption = 'R'#233'f'#233'rence externe'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object e_refexternepcl: TLabel
        Tag = 1
        Left = 267
        Top = 35
        Width = 138
        Height = 13
        AutoSize = False
        Caption = 'E_JOURNAL'
      end
      object Label5: TLabel
        Left = 416
        Top = 4
        Width = 121
        Height = 13
        AutoSize = False
        Caption = 'Compte de contrepartie'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object E_CONTREPARTIEGENPCL: TLabel
        Tag = 1
        Left = 555
        Top = 4
        Width = 113
        Height = 13
        AutoSize = False
        Caption = 'E_JOURNAL'
      end
      object E_ETABLISSEMENTPCL: TLabel
        Tag = 1
        Left = 555
        Top = 19
        Width = 113
        Height = 13
        AutoSize = False
        Caption = 'E_JOURNAL'
      end
      object Label6: TLabel
        Left = 416
        Top = 19
        Width = 121
        Height = 13
        AutoSize = False
        Caption = 'Etablissement'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
    end
  end
  object HMessVisuL: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Visualisation des solutions du rapprochement;Validez-vous le c' +
        'hoix de cette solution ?;Q;YN;Y;Y;')
    Left = 380
    Top = 112
  end
  object POPS: TPopupMenu
    OnPopup = POPSPopup
    Left = 324
    Top = 112
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 444
    Top = 112
  end
end

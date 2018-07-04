object FSaisBud: TFSaisBud
  Left = 164
  Top = 125
  Width = 634
  Height = 403
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Saisie du budget'
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
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object DockTop: TDock97
    Left = 0
    Top = 0
    Width = 626
    Height = 9
  end
  object DockRight: TDock97
    Left = 617
    Top = 9
    Width = 9
    Height = 332
    Position = dpRight
  end
  object DockLeft: TDock97
    Left = 0
    Top = 9
    Width = 9
    Height = 332
    Position = dpLeft
  end
  object DockBottom: TDock97
    Left = 0
    Top = 341
    Width = 626
    Height = 28
    Position = dpBottom
    object Valide97: TToolbar97
      Left = 522
      Top = 0
      Caption = 'Validation'
      CloseButton = False
      DockPos = 522
      TabOrder = 0
      object BValide: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Valider la saisie'
        DisplayMode = dmGlyphOnly
        Caption = 'Valider'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValideClick
        GlobalIndexImage = 'Z0184_S16G1'
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 27
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Fermer'
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
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
        Left = 54
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
    end
    object Outils97: TToolbar97
      Left = 0
      Top = 0
      Caption = 'Outils'
      CloseButton = False
      DockPos = 0
      TabOrder = 1
      object BVentil: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Ventilations'
        DisplayMode = dmGlyphOnly
        Caption = 'Ventilations'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BVentilClick
        GlobalIndexImage = 'Z0133_S16G1'
        IsControl = True
      end
      object BComplement: TToolbarButton97
        Tag = 1
        Left = 27
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Informations compl'#233'mentaires'
        DisplayMode = dmGlyphOnly
        Caption = 'Compl'#233'ments'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 0
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BComplementClick
        GlobalIndexImage = 'Z0105_S16G2'
        IsControl = True
      end
      object BChercher: TToolbarButton97
        Tag = 1
        Left = 54
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Rechercher dans la pi'#232'ce'
        DisplayMode = dmGlyphOnly
        Caption = 'Chercher'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BChercherClick
        GlobalIndexImage = 'Z0077_S16G1'
        IsControl = True
      end
      object BMenuZoom: TToolbarButton97
        Tag = -100
        Left = 184
        Top = 0
        Width = 37
        Height = 24
        Hint = 'Menu zoom'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = POPZ
        Caption = 'Zooms'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnMouseEnter = BMenuZoomMouseEnter
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
      object BAction: TToolbarButton97
        Tag = -101
        Left = 141
        Top = 0
        Width = 37
        Height = 24
        Hint = 'Menu actions'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = POPZ
        Caption = 'Actions'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnMouseEnter = BActionMouseEnter
        GlobalIndexImage = 'Z0046_S16G1'
        IsControl = True
      end
      object ToolbarSep972: TToolbarSep97
        Left = 178
        Top = 0
      end
      object BCopyCol: TToolbarButton97
        Tag = 1
        Left = 81
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Copier une colonne'
        DisplayMode = dmGlyphOnly
        Caption = 'Copier colonne'
        Layout = blGlyphTop
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BCopyColClick
        GlobalIndexImage = 'Z1016_S16G1'
      end
      object BPasteCol: TToolbarButton97
        Tag = 101
        Left = 108
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Coller une colonne'
        DisplayMode = dmGlyphOnly
        Caption = 'Coller colonne'
        Layout = blGlyphTop
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BPasteColClick
        GlobalIndexImage = 'Z1175_S16G1'
      end
      object ToolbarSep973: TToolbarSep97
        Left = 135
        Top = 0
      end
    end
  end
  object PFEN: TPanel
    Left = 9
    Top = 9
    Width = 608
    Height = 332
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object PG: TPanel
      Left = 0
      Top = 107
      Width = 588
      Height = 225
      Align = alClient
      BevelOuter = bvNone
      TabOrder = 7
      object GDroit: THGrid
        Left = 501
        Top = 0
        Width = 87
        Height = 180
        TabStop = False
        Align = alRight
        ColCount = 1
        Ctl3D = True
        DefaultColWidth = 82
        DefaultRowHeight = 16
        Enabled = False
        FixedCols = 0
        RowCount = 25
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine]
        ParentCtl3D = False
        ParentFont = False
        ScrollBars = ssNone
        TabOrder = 1
        SortedCol = -1
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = 13224395
      end
      object Panel5: THPanel
        Left = 0
        Top = 180
        Width = 588
        Height = 45
        Align = alBottom
        BevelOuter = bvNone
        Caption = ' '
        FullRepaint = False
        TabOrder = 2
        BackGroundEffect = bdFond
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        ColorNb = 2
        TextEffect = tenone
        object GBas: THGrid
          Left = 0
          Top = 1
          Width = 521
          Height = 20
          TabStop = False
          ColCount = 14
          Ctl3D = True
          DefaultColWidth = 80
          DefaultRowHeight = 16
          Enabled = False
          RowCount = 1
          FixedRows = 0
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine]
          ParentCtl3D = False
          ParentFont = False
          ScrollBars = ssNone
          TabOrder = 0
          SortedCol = -1
          Couleur = False
          MultiSelect = False
          TitleBold = True
          TitleCenter = True
          ColCombo = 0
          SortEnabled = False
          SortRowExclude = 0
          TwoColors = False
          AlternateColor = 13224395
        end
        object TG: TPanel
          Left = 521
          Top = 0
          Width = 86
          Height = 21
          Alignment = taRightJustify
          BevelOuter = bvNone
          BorderStyle = bsSingle
          Color = clWindow
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 1
        end
        object SB: TScrollBar
          Left = 166
          Top = 24
          Width = 351
          Height = 17
          LargeChange = 5
          Min = 1
          PageSize = 0
          Position = 1
          TabOrder = 2
          OnChange = SBChange
        end
      end
      object GB: THGrid
        Left = 0
        Top = 0
        Width = 501
        Height = 180
        Align = alClient
        ColCount = 14
        Ctl3D = True
        DefaultColWidth = 80
        DefaultRowHeight = 16
        RowCount = 25
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goThumbTracking]
        ParentCtl3D = False
        PopupMenu = POPS
        ScrollBars = ssNone
        TabOrder = 0
        OnEnter = GBEnter
        OnExit = GBExit
        OnTopLeftChanged = GBTopLeftChanged
        SortedCol = -1
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnCellEnter = GBCellEnter
        OnCellExit = GBCellExit
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = 13224395
        RowHeights = (
          16
          16
          16
          16
          16
          16
          16
          16
          16
          16
          16
          16
          16
          16
          16
          16
          16
          16
          16
          16
          16
          16
          16
          16
          16)
      end
      object BE_EXERCICE: THValComboBox
        Left = 427
        Top = 147
        Width = 20
        Height = 21
        Style = csSimple
        Color = clYellow
        Ctl3D = True
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 3
        Visible = False
        OnChange = BE_EXERCICEChange
        TagDispatch = 0
        DataType = 'TTEXERCICEBUDGET'
      end
      object BE_EXERCICE_: THValComboBox
        Left = 450
        Top = 147
        Width = 20
        Height = 21
        Style = csSimple
        Color = clYellow
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 4
        Visible = False
        OnChange = BE_EXERCICE_Change
        TagDispatch = 0
        DataType = 'TTEXERCICEBUDGET'
      end
      object BPasteAll: THBitBtn
        Tag = 101
        Left = 297
        Top = 120
        Width = 27
        Height = 27
        Hint = 'Coller sur les colonnes vides'
        Caption = 'Paste'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 5
        Visible = False
        OnClick = BPasteAllClick
        Layout = blGlyphTop
        Spacing = -1
      end
      object BRepart: THBitBtn
        Tag = 101
        Left = 329
        Top = 120
        Width = 27
        Height = 27
        Hint = 'R'#233'partition par cl'#233's budg'#233'taires'
        Caption = 'Repart'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 6
        Visible = False
        OnClick = BRepartClick
        Layout = blGlyphTop
        Spacing = -1
      end
      object BCreerRepart: THBitBtn
        Tag = 101
        Left = 365
        Top = 120
        Width = 27
        Height = 27
        Hint = 'Cr'#233'ation de cl'#233's de r'#233'partitions'
        Caption = 'Repart'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 7
        Visible = False
        OnClick = BCreerRepartClick
        Layout = blGlyphTop
        Spacing = -1
      end
      object EXOSINTER: THValComboBox
        Left = 427
        Top = 172
        Width = 22
        Height = 21
        Style = csSimple
        Color = clYellow
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 8
        Visible = False
        TagDispatch = 0
      end
      object ExoCalcul: THValComboBox
        Left = 451
        Top = 172
        Width = 22
        Height = 21
        Style = csSimple
        Color = clYellow
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 9
        Visible = False
        TagDispatch = 0
      end
      object BSwapSaisie: THBitBtn
        Tag = 101
        Left = 93
        Top = 152
        Width = 27
        Height = 27
        Hint = 'Basculer la saisie mono-colonne /  multi-colonnes'
        Caption = 'Mono'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 10
        Visible = False
        OnClick = BSwapSaisieClick
        Layout = blGlyphTop
        Spacing = -1
      end
    end
    object PD: THPanel
      Left = 588
      Top = 107
      Width = 20
      Height = 225
      Align = alRight
      Caption = ' '
      FullRepaint = False
      TabOrder = 8
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
      object SD: TScrollBar
        Left = 1
        Top = 4
        Width = 17
        Height = 197
        Kind = sbVertical
        LargeChange = 5
        Min = 1
        PageSize = 0
        Position = 1
        TabOrder = 0
        OnChange = SDChange
      end
    end
    object BZoomJournal: THBitBtn
      Tag = 100
      Left = 94
      Top = 200
      Width = 27
      Height = 27
      Hint = 'Voir budget'
      Caption = 'Jal'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      Visible = False
      OnClick = BZoomJournalClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BZoomBudget: THBitBtn
      Tag = 100
      Left = 127
      Top = 200
      Width = 27
      Height = 27
      Hint = 'Voir compte budg'#233'taire'
      Caption = 'Bud'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      Visible = False
      OnClick = BZoomBudgetClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BZoomSection: THBitBtn
      Tag = 100
      Left = 161
      Top = 200
      Width = 27
      Height = 27
      Hint = 'Voir section budg'#233'taire'
      Caption = 'Sect'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      Visible = False
      OnClick = BZoomSectionClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BRazCol: THBitBtn
      Tag = 101
      Left = 194
      Top = 200
      Width = 27
      Height = 27
      Hint = 'RAZ colonne'
      Caption = 'RazC'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 4
      Visible = False
      OnClick = BRazColClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BRazLigne: THBitBtn
      Tag = 101
      Left = 228
      Top = 200
      Width = 27
      Height = 27
      Hint = 'RAZ ligne'
      Caption = 'RazL'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 5
      Visible = False
      OnClick = BRazLigneClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BRazAll: THBitBtn
      Tag = 101
      Left = 261
      Top = 200
      Width = 27
      Height = 27
      Hint = 'Effacer la grille'
      Caption = 'RazAll'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 6
      Visible = False
      OnClick = BRazAllClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object PEntete: THPanel
      Left = 0
      Top = 0
      Width = 608
      Height = 107
      Align = alTop
      FullRepaint = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
      object BevelNum: TBevel
        Left = 315
        Top = 54
        Width = 137
        Height = 20
      end
      object H_PERFIN: THLabel
        Left = 469
        Top = 9
        Width = 6
        Height = 13
        Caption = #224
        Transparent = True
      end
      object H_JOURNAL: THLabel
        Left = 12
        Top = 8
        Width = 34
        Height = 13
        Caption = 'Budget'
        FocusControl = BE_BUDJAL
        Transparent = True
      end
      object H_NATUREBUD: THLabel
        Left = 12
        Top = 33
        Width = 32
        Height = 13
        Caption = 'Nature'
        FocusControl = BE_NATUREBUD
        Transparent = True
      end
      object H_AXE: THLabel
        Left = 12
        Top = 58
        Width = 18
        Height = 13
        Caption = 'Axe'
        FocusControl = BE_AXE
        Transparent = True
      end
      object H_BUDSECT: THLabel
        Left = 12
        Top = 58
        Width = 36
        Height = 13
        Caption = 'Section'
        FocusControl = BE_BUDGENE
        Transparent = True
      end
      object H_BUDGENE: THLabel
        Left = 12
        Top = 58
        Width = 58
        Height = 13
        Caption = 'Cpte budget'
        FocusControl = BE_BUDGENE
        Transparent = True
      end
      object H_PERDEB: THLabel
        Left = 246
        Top = 9
        Width = 51
        Height = 13
        Caption = 'P'#233'riode de'
        FocusControl = BE_BUDJAL
        Transparent = True
      end
      object H_NUMEROPIECE: THLabel
        Left = 246
        Top = 58
        Width = 37
        Height = 13
        Caption = 'Num'#233'ro'
        Transparent = True
      end
      object BE_NUMEROPIECE: THLabel
        Left = 317
        Top = 56
        Width = 134
        Height = 16
        Alignment = taCenter
        AutoSize = False
        Caption = 'Numpi'#232'ce'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object H_ETABLISSEMENT: THLabel
        Left = 246
        Top = 33
        Width = 65
        Height = 13
        Caption = 'Etablissement'
        FocusControl = BE_ETABLISSEMENT
        Transparent = True
      end
      object BZoomSousPlan: TToolbarButton97
        Left = 227
        Top = 4
        Width = 17
        Height = 21
        Caption = '...'
        Visible = False
        OnClick = BZoomSousPlanClick
      end
      object HLabel3: THLabel
        Left = 12
        Top = 84
        Width = 19
        Height = 13
        Caption = 'Vue'
        FocusControl = VUE
        Transparent = True
      end
      object HLabel4: THLabel
        Left = 246
        Top = 84
        Width = 25
        Height = 13
        Caption = 'Unit'#233
        FocusControl = RESOL
        Transparent = True
      end
      object H_CSens: THLabel
        Left = 469
        Top = 33
        Width = 24
        Height = 13
        Caption = '&Sens'
        FocusControl = CSens
        Transparent = True
      end
      object BE_BUDGENE: THCpteEdit
        Left = 82
        Top = 54
        Width = 145
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 3
        OnExit = BE_BUDGENEExit
        ZoomTable = tzBudgen
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object BE_AXE: THValComboBox
        Left = 82
        Top = 54
        Width = 145
        Height = 21
        Style = csSimple
        Enabled = False
        ItemHeight = 13
        TabOrder = 2
        Visible = False
        OnChange = BE_AXEChange
        TagDispatch = 0
        DataType = 'TTAXE'
      end
      object BE_BUDSECT: THCpteEdit
        Left = 82
        Top = 54
        Width = 145
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 4
        OnExit = BE_BUDSECTExit
        ZoomTable = tzBudSec1
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object BE_NATUREBUD: THValComboBox
        Left = 82
        Top = 29
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 1
        TagDispatch = 0
        DataType = 'TTNATECRBUD'
      end
      object BE_BUDJAL: THValComboBox
        Left = 82
        Top = 4
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = BE_BUDJALChange
        TagDispatch = 0
        DataType = 'TTBUDJAL'
      end
      object PERDEB: THValComboBox
        Left = 315
        Top = 4
        Width = 137
        Height = 21
        Style = csSimple
        Enabled = False
        ItemHeight = 13
        TabOrder = 5
        TagDispatch = 0
      end
      object PERFIN: THValComboBox
        Left = 497
        Top = 4
        Width = 137
        Height = 21
        Style = csSimple
        Enabled = False
        ItemHeight = 13
        TabOrder = 6
        TagDispatch = 0
      end
      object BE_ETABLISSEMENT: THValComboBox
        Left = 315
        Top = 29
        Width = 137
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 7
        OnChange = BE_ETABLISSEMENTChange
        TagDispatch = 0
        DataType = 'TTETABLISSEMENT'
      end
      object VUE: THValComboBox
        Left = 82
        Top = 80
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 8
        OnChange = VUEChange
        Items.Strings = (
          'Comptes budg'#233'taires'
          'Sections budg'#233'taires'
          'Comptes / Section'
          'Sections / Compte')
        TagDispatch = 0
        Values.Strings = (
          'G'
          'S'
          'GS'
          'SG')
      end
      object RESOL: THValComboBox
        Left = 315
        Top = 80
        Width = 137
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 9
        OnChange = RESOLChange
        Items.Strings = (
          'Avec d'#233'cimales'
          'Sans d'#233'cimales'
          'en K(ilo)'
          'en M('#233'ga)')
        TagDispatch = 0
        Values.Strings = (
          'C'
          'F'
          'K'
          'M')
      end
      object CSens: TComboBox
        Left = 497
        Top = 29
        Width = 126
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 10
        OnChange = CSensChange
        Items.Strings = (
          'D'#233'bit'
          'Cr'#233'dit')
      end
    end
  end
  object POPS: TPopupMenu
    OnPopup = POPSPopup
    Left = 150
    Top = 153
  end
  object POPZ: TPopupMenu
    AutoPopup = False
    Left = 104
    Top = 152
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Saisie budg'#233'taire;Confirmez-vous l'#39'abandon de la saisie ?;Q;YN' +
        ';N;N;'
      
        '1;Saisie budg'#233'taire;Confirmez-vous l'#39'effacement de la grille ?;Q' +
        ';YN;N;N;'
      
        '2;Saisie budg'#233'taire;Le budget est valid'#233' ! Seule la consultation' +
        ' est autoris'#233'e;W;O;O;O;'
      
        '3;Saisie budg'#233'taire;La colonne contient d'#233'j'#224' des donn'#233'es. Voulez' +
        '-vous '#233'craser les d'#233'bits et cr'#233'dits ?;Q;YN;Y;Y;'
      
        '4;Saisie budg'#233'taire;Le budget n'#39'est pas renseign'#233' . Saisie oblig' +
        'atoire !;W;O;O;O;'
      
        '5;Saisie budg'#233'taire;La nature de l'#39#233'criture est non renseign'#233'e .' +
        ' Saisie obligatoire !;W;O;O;O;'
      
        '6;Saisie budg'#233'taire;La section renseign'#233'e n'#39'existe pas pour ce b' +
        'udget. Saisie obligatoire !;W;O;O;O;'
      
        '7;?caption?;Vous devez renseigner un compte existant pour ce bud' +
        'get;W;O;O;O;'
      
        '8;Saisie budg'#233'taire;Certains montants sur des comptes et section' +
        's ont des sens incoh'#233'rents. Confirmez-vous la validation ?;Q;YN;' +
        'N;N;')
    Left = 195
    Top = 152
  end
  object FindSais: TFindDialog
    OnFind = FindSaisFind
    Left = 286
    Top = 153
  end
  object HDiv: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Total'
      'Sections'
      'Comptes'
      'Total'
      
        'Validation non effectu'#233'e ! Ce budget est en cours de modificatio' +
        'n par un autre utilisateur'
      'Budget non enregistr'#233' !'
      'Ventilations analytiques'
      'Ventilations budg'#233'taires'
      'Consultation du budget'
      'Modification du budget'
      'Saisie du budget par comptes'
      'Vision consolid'#233'e comptes / sections'
      'Vision consolid'#233'e sections / comptes'
      'Choix d'#39'une cl'#233' de r'#233'partition'
      'Saisie du budget par sections'
      'Saisie du budget par comptes avec section implicite'
      'Saisie du budget par sections avec compte implicite'
      'Budget saisi en'
      '18')
    Left = 332
    Top = 153
  end
  object HMTrad: THSystemMenu
    LockedCtrls.Strings = (
      'GB'
      'GBas'
      'GDroit'
      'TG')
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 376
    Top = 154
  end
end

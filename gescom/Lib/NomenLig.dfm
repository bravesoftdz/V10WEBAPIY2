object FNomenLig: TFNomenLig
  Left = 258
  Top = 222
  Width = 688
  Height = 538
  HelpContext = 110000239
  Caption = 'Saisie de la nomenclature'
  Color = clBtnFace
  Constraints.MinHeight = 440
  Constraints.MinWidth = 620
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Dock971: TDock97
    Left = 0
    Top = 459
    Width = 672
    Height = 40
    AllowDrag = False
    Position = dpBottom
    object ToolWindow971: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 36
      ClientWidth = 672
      Caption = 'ToolWindow971'
      ClientAreaHeight = 36
      ClientAreaWidth = 672
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BAide: TToolbarButton97
        Tag = 1
        Left = 569
        Top = 5
        Width = 28
        Height = 27
        Hint = 'Aide'
        Caption = 'Aide'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 539
        Top = 5
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Caption = 'Annuler'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ModalResult = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BValider: TToolbarButton97
        Left = 509
        Top = 5
        Width = 28
        Height = 27
        Hint = 'Enregistrer la nomenclature'
        Caption = 'Enregistrer'
        Default = True
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ModalResult = 1
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0127_S16G1'
        IsControl = True
      end
      object bNewligne: TToolbarButton97
        Left = 2
        Top = 5
        Width = 28
        Height = 27
        Hint = 'Ins'#233'rer une ligne'
        Caption = 'Ins'#233'rer'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = bNewligneClick
        GlobalIndexImage = 'Z0074_S16G1'
      end
      object bDelLigne: TToolbarButton97
        Left = 32
        Top = 5
        Width = 28
        Height = 27
        Hint = 'Supprimer la ligne en cours'
        Caption = 'Effacer'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = bDelLigneClick
        GlobalIndexImage = 'Z0072_S16G1'
      end
      object bNomenc: TToolbarButton97
        Left = 62
        Top = 5
        Width = 28
        Height = 27
        Hint = 'D'#233'tail de la nomenclature'
        Flat = False
        OnClick = bNomencClick
        GlobalIndexImage = 'Z0678_S24G1'
      end
      object bPrint: TToolbarButton97
        Left = 479
        Top = 5
        Width = 28
        Height = 27
        Flat = False
        OnClick = bPrintClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BInfos: TToolbarButton97
        Tag = 1
        Left = 179
        Top = 5
        Width = 40
        Height = 27
        Hint = 'Compl'#233'ments d'#39'informations'
        Caption = 'Compl'#233'ments'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = PopupG_NLIG
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        Spacing = -1
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
      object BTypeArticle: TToolbarButton97
        Left = 224
        Top = 5
        Width = 13
        Height = 22
        Alignment = taLeftJustify
        Opaque = False
        OnClick = BTypeArticleClick
      end
      object BArborescence: TToolbarButton97
        Tag = 1
        Left = 388
        Top = 5
        Width = 28
        Height = 27
        Hint = 'Structure de l'#39'ouvrage'
        Caption = 'Structure de l'#39'ouvrage'
        AllowAllUp = True
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        GlobalIndexImage = 'Z0189_S16G0'
        IsControl = True
      end
      object BInfosLigne: TToolbarButton97
        Tag = 1
        Left = 418
        Top = 5
        Width = 28
        Height = 27
        Hint = 'Infos ligne'
        Caption = 'Infos ligne'
        AllowAllUp = True
        GroupIndex = 2
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        GlobalIndexImage = 'Z0029_S16G1'
        IsControl = True
      end
      object Variables_Lig: TToolbarButton97
        Left = 110
        Top = 5
        Width = 28
        Height = 27
        Hint = 'Variables de la ligne sous-d'#233'tail'
        Flat = False
        OnClick = Variables_LigClick
        GlobalIndexImage = 'Z0589_S16G1'
      end
      object AppelXLS: TToolbarButton97
        Left = 142
        Top = 5
        Width = 28
        Height = 27
        Hint = 'Appel fichier XLS sous-d'#233'tail'
        Flat = False
        OnClick = AppelXLSClick
        GlobalIndexImage = 'Z2263_S16G1'
      end
    end
  end
  object G_NLIG: THGrid
    Tag = 1
    Left = 0
    Top = 83
    Width = 432
    Height = 235
    Align = alClient
    DefaultRowHeight = 18
    RowCount = 50
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goRowMoving, goEditing, goRowSelect]
    PopupMenu = PopupG_NLIG
    TabOrder = 2
    OnDblClick = G_NLIGDblClick
    OnMouseDown = G_NLIGMouseDown
    OnRowMoved = G_NLIGRowMoved
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    OnRowEnter = G_NLIGRowEnter
    OnRowExit = G_NLIGRowExit
    OnCellEnter = G_NLIGCellEnter
    OnCellExit = G_NLIGCellExit
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = clSilver
    ElipsisButton = True
    OnElipsisClick = G_NLIGElipsisClick
    ColWidths = (
      64
      64
      64
      64
      64)
  end
  object PANNOMENENT: TPanel
    Left = 0
    Top = 0
    Width = 672
    Height = 57
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    object TGNE_ARTICLE: TLabel
      Left = 316
      Top = 12
      Width = 103
      Height = 13
      Caption = 'Code Article compos'#233
      FocusControl = GNE_ARTICLE
    end
    object TGNE_NOMENCLATURE: TLabel
      Left = 16
      Top = 12
      Width = 94
      Height = 13
      Caption = 'Code Nomenclature'
      FocusControl = GNE_NOMENCLATURE
    end
    object TGNE_LIBELLE: TLabel
      Left = 16
      Top = 36
      Width = 123
      Height = 13
      Caption = 'Libell'#233' de la nomenclature'
      FocusControl = GNE_LIBELLE
    end
    object GNE_ARTICLE: TEdit
      Left = 428
      Top = 8
      Width = 121
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 0
    end
    object GNE_NOMENCLATURE: TEdit
      Left = 144
      Top = 8
      Width = 121
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 1
    end
    object GNE_LIBELLE: TEdit
      Left = 144
      Top = 32
      Width = 405
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 2
    end
  end
  object PANPIED: TPanel
    Left = 0
    Top = 433
    Width = 672
    Height = 26
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 6
    object TGNL_SOUSNOMEN: TLabel
      Left = 12
      Top = 7
      Width = 66
      Height = 13
      Caption = 'Nomenclature'
      FocusControl = GNL_SOUSNOMEN
      Visible = False
    end
    object GNL_SOUSNOMEN: TEdit
      Left = 84
      Top = 3
      Width = 185
      Height = 21
      Hint = 'Code de la sous nomenclature'
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 0
      Visible = False
    end
  end
  object PANDIM: TPanel
    Left = 0
    Top = 420
    Width = 672
    Height = 13
    Align = alBottom
    AutoSize = True
    BevelOuter = bvNone
    TabOrder = 4
    object TGA_CODEDIM5: TLabel
      Left = 500
      Top = 0
      Width = 32
      Height = 13
      Caption = 'Label1'
    end
    object TGA_CODEDIM4: TLabel
      Left = 376
      Top = 0
      Width = 32
      Height = 13
      Caption = 'Label1'
    end
    object TGA_CODEDIM3: TLabel
      Left = 252
      Top = 0
      Width = 32
      Height = 13
      Caption = 'Label1'
    end
    object TGA_CODEDIM2: TLabel
      Left = 132
      Top = 0
      Width = 32
      Height = 13
      Caption = 'Label1'
    end
    object TGA_CODEDIM1: TLabel
      Left = 12
      Top = 0
      Width = 32
      Height = 13
      Caption = 'Label1'
    end
  end
  object BTPpanel: THPanel
    Left = 432
    Top = 83
    Width = 240
    Height = 235
    Align = alRight
    BevelOuter = bvNone
    BorderStyle = bsSingle
    FullRepaint = False
    TabOrder = 3
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object TV_NLIG: TTreeView
      Left = 0
      Top = 0
      Width = 236
      Height = 231
      Align = alClient
      HideSelection = False
      Images = ImageList1
      Indent = 19
      PopupMenu = PopupTVNLIG
      StateImages = ImageList1
      TabOrder = 0
      OnCollapsing = TV_NLIGCollapsing
      OnDblClick = TV_NLIGDblClick
      OnExpanding = TV_NLIGExpanding
      OnMouseUp = TV_NLIGMouseUp
    end
  end
  object POuvrage: THPanel
    Left = 0
    Top = 318
    Width = 672
    Height = 102
    Align = alBottom
    AutoSize = True
    BevelOuter = bvNone
    BorderStyle = bsSingle
    FullRepaint = False
    TabOrder = 5
    Visible = False
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    AutoResize = True
    object HPanel3: THPanel
      Left = 0
      Top = 0
      Width = 668
      Height = 25
      Align = alTop
      FullRepaint = False
      TabOrder = 1
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
      object HLabel12: THLabel
        Left = 5
        Top = 4
        Width = 39
        Height = 13
        Caption = 'Coef FG'
      end
      object HLabel13: THLabel
        Left = 162
        Top = 4
        Width = 55
        Height = 13
        Caption = 'Coef Marge'
      end
      object HLabel17: THLabel
        Left = 321
        Top = 4
        Width = 59
        Height = 13
        Caption = 'Temps Total'
      end
      object HLabel21: THLabel
        Left = 472
        Top = 5
        Width = 69
        Height = 13
        Caption = 'Temps unitaire'
      end
      object COEFFG: THNumEdit
        Left = 80
        Top = 1
        Width = 74
        Height = 21
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.0000'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object COEFMARG: THNumEdit
        Left = 241
        Top = 1
        Width = 68
        Height = 21
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 1
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.0000'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object TOTALTPS: THNumEdit
        Left = 386
        Top = 1
        Width = 68
        Height = 21
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 2
        OnExit = ExitPVForfait
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object TPSUNITAIRE: THNumEdit
        Left = 554
        Top = 1
        Width = 68
        Height = 21
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 3
        OnExit = ExitPVForfait
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
    end
    object HPanel1: THPanel
      Left = 0
      Top = 25
      Width = 668
      Height = 73
      Align = alTop
      AutoSize = True
      FullRepaint = False
      TabOrder = 0
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
      AutoResize = True
      object PValo: THPanel
        Left = 1
        Top = 1
        Width = 382
        Height = 71
        Align = alLeft
        Alignment = taLeftJustify
        AutoSize = True
        FullRepaint = False
        TabOrder = 1
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        AutoResize = True
        object HLabel8: THLabel
          Left = 3
          Top = 25
          Width = 68
          Height = 13
          Caption = 'D'#233'bours'#233' Sec'
        end
        object HLabel9: THLabel
          Left = 147
          Top = 25
          Width = 18
          Height = 13
          Caption = 'P.R'
        end
        object HLabel10: THLabel
          Left = 258
          Top = 24
          Width = 38
          Height = 13
          Caption = 'P.V H.T'
        end
        object HLabel11: THLabel
          Left = 257
          Top = 49
          Width = 47
          Height = 13
          Caption = 'P.V T.T.C'
        end
        object HLabel14: THLabel
          Left = 3
          Top = 49
          Width = 59
          Height = 13
          Caption = 'Montant F.G'
        end
        object HLabel15: THLabel
          Left = 147
          Top = 49
          Width = 30
          Height = 13
          Caption = 'Marge'
        end
        object HLabel28: THLabel
          Left = 1
          Top = 1
          Width = 380
          Height = 13
          Align = alTop
          Alignment = taCenter
          Caption = 'Valorisations'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object MONTANTPA: THNumEdit
          Left = 74
          Top = 21
          Width = 68
          Height = 21
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 0
          OnExit = ExitPVForfait
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object MONTANTPR: THNumEdit
          Left = 184
          Top = 21
          Width = 68
          Height = 21
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 1
          OnExit = ExitPVForfait
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object MONTANTHT: THNumEdit
          Left = 308
          Top = 21
          Width = 68
          Height = 21
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 2
          OnExit = ExitPVForfait
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object MONTANTTTC: THNumEdit
          Left = 308
          Top = 46
          Width = 68
          Height = 21
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 3
          OnExit = ExitPVForfait
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object MONTANTFG: THNumEdit
          Left = 74
          Top = 46
          Width = 68
          Height = 21
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 4
          OnExit = ExitPVForfait
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object MONTANTMARG: THNumEdit
          Left = 184
          Top = 46
          Width = 68
          Height = 21
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 5
          OnExit = ExitPVForfait
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          UseRounding = True
          Validate = False
        end
      end
      object PValoUnitaire: THPanel
        Left = 383
        Top = 1
        Width = 382
        Height = 71
        Align = alLeft
        Alignment = taLeftJustify
        AutoSize = True
        FullRepaint = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        AutoResize = True
        object HLabel18: THLabel
          Left = 3
          Top = 24
          Width = 68
          Height = 13
          Caption = 'D'#233'bours'#233' Sec'
        end
        object HLabel19: THLabel
          Left = 146
          Top = 24
          Width = 18
          Height = 13
          Caption = 'P.R'
        end
        object HLabel20: THLabel
          Left = 255
          Top = 23
          Width = 38
          Height = 13
          Caption = 'P.V H.T'
        end
        object HLabel23: THLabel
          Left = 254
          Top = 50
          Width = 47
          Height = 13
          Caption = 'P.V T.T.C'
        end
        object HLabel24: THLabel
          Left = 3
          Top = 49
          Width = 59
          Height = 13
          Caption = 'Montant F.G'
        end
        object HLabel25: THLabel
          Left = 146
          Top = 49
          Width = 30
          Height = 13
          Caption = 'Marge'
        end
        object HLabel27: THLabel
          Left = 1
          Top = 1
          Width = 380
          Height = 13
          Align = alTop
          Alignment = taCenter
          Caption = 'Valorisation unitaire'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object MONTANTPAU: THNumEdit
          Left = 74
          Top = 20
          Width = 68
          Height = 21
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 0
          OnExit = ExitPVForfait
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object MONTANTPRU: THNumEdit
          Left = 179
          Top = 20
          Width = 68
          Height = 21
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 1
          OnExit = ExitPVForfait
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object MONTANTHTU: THNumEdit
          Left = 310
          Top = 20
          Width = 68
          Height = 21
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 2
          OnExit = ExitPVForfait
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object MONTANTTTCU: THNumEdit
          Left = 310
          Top = 46
          Width = 68
          Height = 21
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 3
          OnExit = ExitPVForfait
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object MONTANTFGU: THNumEdit
          Left = 74
          Top = 46
          Width = 68
          Height = 21
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 4
          OnExit = ExitPVForfait
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object MONTANTMARGU: THNumEdit
          Left = 179
          Top = 46
          Width = 68
          Height = 21
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 5
          OnExit = ExitPVForfait
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          UseRounding = True
          Validate = False
        end
      end
    end
  end
  object PANDetailBtp: TPanel
    Left = 0
    Top = 57
    Width = 672
    Height = 26
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    Visible = False
    object TGNE_QTEDUDETAIL: TLabel
      Left = 16
      Top = 7
      Width = 90
      Height = 13
      Caption = 'D'#233'tail exprim'#233' pour'
      Visible = False
    end
    object TUNITE: THLabel
      Left = 204
      Top = 7
      Width = 24
      Height = 13
      Alignment = taCenter
      Caption = 'UUU'
    end
    object GNE_QTEDUDETAIL: THNumEdit
      Left = 144
      Top = 2
      Width = 49
      Height = 21
      TabOrder = 0
      OnExit = GNE_QTEDUDETAILExit
      Decimals = 0
      Digits = 12
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      NumericType = ntDecimal
      UseRounding = True
      Validate = False
    end
  end
  object ToolBarDesc: TToolWindow97
    Left = 502
    Top = 6
    ClientHeight = 245
    ClientWidth = 265
    Caption = 'Infos ligne'
    ClientAreaHeight = 245
    ClientAreaWidth = 265
    TabOrder = 8
    Visible = False
    OnClose = ToolBarDescClose
    object Pligne: THPanel
      Left = 0
      Top = 0
      Width = 265
      Height = 245
      Align = alClient
      BevelOuter = bvLowered
      FullRepaint = False
      TabOrder = 0
      Visible = False
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
      object HLabel1: THLabel
        Left = 4
        Top = 53
        Width = 48
        Height = 13
        Caption = 'Prix Achat'
      end
      object HLabel3: THLabel
        Left = 4
        Top = 101
        Width = 67
        Height = 13
        Caption = 'Prix de revient'
      end
      object HLabel2: THLabel
        Left = 4
        Top = 77
        Width = 134
        Height = 13
        Caption = 'Coefficient de frais g'#233'n'#233'raux'
      end
      object HLabel4: THLabel
        Left = 4
        Top = 125
        Width = 96
        Height = 13
        Caption = 'coefficient de marge'
      end
      object HLabel5: THLabel
        Left = 4
        Top = 149
        Width = 79
        Height = 13
        Caption = 'P.V H.T de base'
      end
      object HLabel6: THLabel
        Left = 4
        Top = 173
        Width = 84
        Height = 13
        Caption = 'P.V H.T forfaitaire'
      end
      object HLabel7: THLabel
        Left = 4
        Top = 197
        Width = 92
        Height = 13
        Caption = 'Prix de vente T.T.C'
      end
      object TGA_FOURNPRINC: TLabel
        Left = 104
        Top = 8
        Width = 114
        Height = 13
        Caption = 'TGA_FOURNPRINC'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label1: TLabel
        Left = 4
        Top = 28
        Width = 92
        Height = 13
        Caption = 'Prestation associ'#233'e'
      end
      object TGA_PRESTATION: TLabel
        Left = 104
        Top = 28
        Width = 112
        Height = 13
        Caption = 'TGA_PRESTATION'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object TLIB_PREST: TLabel
        Left = 4
        Top = 8
        Width = 54
        Height = 13
        Caption = 'Fournisseur'
      end
      object HLabel16: THLabel
        Left = 4
        Top = 221
        Width = 73
        Height = 13
        Caption = 'Temps de pose'
      end
      object NEPA: THNumEdit
        Left = 148
        Top = 49
        Width = 109
        Height = 21
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object NECOEFPAPR: THNumEdit
        Left = 182
        Top = 73
        Width = 75
        Height = 21
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 1
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.0000'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object NEPR: THNumEdit
        Left = 148
        Top = 97
        Width = 109
        Height = 21
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 2
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object NECOEFPRPV: THNumEdit
        Left = 182
        Top = 121
        Width = 75
        Height = 21
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 3
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.0000'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object NEPVHT: THNumEdit
        Left = 148
        Top = 145
        Width = 109
        Height = 21
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 4
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object NEPVFORFAIT: THNumEdit
        Left = 148
        Top = 169
        Width = 109
        Height = 21
        TabOrder = 5
        OnExit = ExitPVForfait
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object NEPVTTC: THNumEdit
        Left = 148
        Top = 193
        Width = 109
        Height = 21
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 6
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object NETPS: THNumEdit
        Left = 196
        Top = 217
        Width = 61
        Height = 21
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 7
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        UseRounding = True
        Validate = False
      end
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Voulez-vous enregistrer les modifications ?;Q;YNC;Y;' +
        'C;'
      '1;?caption?;Voulez-vous supprimer l'#39'enregistrement ?;Q;YNC;N;C;'
      
        '2;?caption?;La quantit'#233' doit '#234'tre sup'#233'rieure ou '#233'gale '#224' z'#233'ro.;W;' +
        'O;O;O;'
      '3;?caption?;Le composant n'#39'existe pas.;W;O;O;O;'
      '4;?caption?;La valeur doit '#234'tre O ou N.;W;O;O;O;'
      '5;?caption?;Le composant est obligatoire.;W;O;O;O;'
      '6;?caption?;Le composant doit '#234'tre unique.;W;O;O;O;'
      
        '7;?caption?;La nomenclature contient au moins une r'#233'f'#233'rence circ' +
        'ulaire.;W;O;O;O;'
      '8;?caption?;Cette nomenclature n'#39'existe pas.;W;O;O;O;'
      '9;?caption?;Cet article est ferm'#233';W;O;O;O;'
      ''
      '')
    Left = 488
    Top = 136
  end
  object ImageList1: THImageList
    GlobalIndexImages.Strings = (
      'Z0172_S16G1'
      'Z0368_S16G1'
      'Z0238_S16G1')
    Left = 536
    Top = 84
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 488
    Top = 80
  end
  object PopupTVNLIG: TPopupMenu
    Left = 424
    Top = 128
    object PopTVNLIG_T: TMenuItem
      Caption = '&Tout ouvrir'
      ShortCut = 16468
      OnClick = PopTVNLIG_TClick
    end
    object PopTVNLIG_F: TMenuItem
      Caption = '&Fermer sous niveaux'
      ShortCut = 16454
      OnClick = PopTVNLIG_FClick
    end
    object N2: TMenuItem
      Caption = '-'
    end
    object PopTV_NLIG_O: TMenuItem
      Caption = '&Ouverture niveau'
      Checked = True
      ShortCut = 16463
      OnClick = PopTV_NLIG_OClick
    end
  end
  object PopupG_NLIG: TPopupMenu
    Left = 456
    Top = 84
    object PopG_NLIG_A: TMenuItem
      Caption = 'Fiche &Article'
      ShortCut = 16449
      OnClick = PopG_NLIG_AClick
    end
    object PopG_NLIG_V: TMenuItem
      Caption = '&Variables'
      ShortCut = 16470
      OnClick = Variables_LigClick
    end
    object PopG_NLIG_E: TMenuItem
      Caption = 'Fichier &Excel'
      ShortCut = 16453
      OnClick = AppelXLSClick
    end
    object PopG_NLIG_C: TMenuItem
      Caption = '&Cas d'#39'emploi'
      ShortCut = 16451
      OnClick = PopG_NLIG_CClick
    end
    object PopG_NLIG_N: TMenuItem
      Caption = '&Nomenclature'
      ShortCut = 16462
      OnClick = PopG_NLIG_NClick
    end
    object PopG_NLIG_O: TMenuItem
      Caption = 'Analyse &Ouvrage'
      ShortCut = 16463
      Visible = False
      OnClick = PopG_NLIG_OClick
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object PopG_NLIG_I: TMenuItem
      Caption = '&Inserer'
      ShortCut = 16457
      OnClick = PopG_NLIG_IClick
    end
    object PopG_NLIG_S: TMenuItem
      Caption = '&Supprimer'
      ShortCut = 16467
      OnClick = PopG_NLIG_SClick
    end
    object N3: TMenuItem
      Caption = '-'
    end
    object PopG_NLIG_D: TMenuItem
      Caption = 'Voir &Dimensions '
      Checked = True
      ShortCut = 16452
      OnClick = PopG_NLIG_DClick
    end
  end
  object ImTypeArticle: THImageList
    GlobalIndexImages.Strings = (
      'Z0054_S16G1'
      'Z0153_S16G1'
      'Z0507_S16G1'
      'Z0146_S16G1'
      'Z0160_S16G1'
      'Z0150_S16G1'
      'Z0415_S16G1'
      'Z0885_S16G1'
      'Z1439_S16G1'
      'Z0510_S16G1')
    Height = 18
    Width = 18
    Left = 679
    Top = 325
  end
end

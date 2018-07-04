object FSaisie: TFSaisie
  Left = 302
  Top = 224
  Width = 617
  Height = 388
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Ecriture'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clBlack
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
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object DockTop: TDock97
    Left = 0
    Top = 0
    Width = 601
    Height = 9
  end
  object DockRight: TDock97
    Left = 592
    Top = 9
    Width = 9
    Height = 312
    Position = dpRight
  end
  object DockLeft: TDock97
    Left = 0
    Top = 9
    Width = 9
    Height = 312
    Position = dpLeft
  end
  object DockBottom: TDock97
    Left = 0
    Top = 321
    Width = 601
    Height = 28
    Position = dpBottom
    object Valide97: TToolbar97
      Left = 506
      Top = 0
      Caption = 'Validation'
      CloseButton = False
      DockPos = 506
      TabOrder = 0
      object BValide: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Valider la saisie'
        Caption = 'Valider'
        DisplayMode = dmGlyphOnly
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
        OnMouseEnter = BValideMouseEnter
        OnMouseExit = BValideMouseExit
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
        Caption = 'Fermer'
        DisplayMode = dmGlyphOnly
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
        Caption = 'Aide'
        DisplayMode = dmGlyphOnly
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
    object Outils: TToolbar97
      Left = 0
      Top = 0
      Caption = 'Actions'
      CloseButton = False
      DockPos = 0
      TabOrder = 1
      object BSolde: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Calcul du solde'
        Caption = 'Solde'
        DisplayMode = dmGlyphOnly
        Layout = blGlyphTop
        NumGlyphs = 2
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BSoldeClick
        GlobalIndexImage = 'Z0051_S16G2'
      end
      object BEche: TToolbarButton97
        Tag = 1
        Left = 27
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Ech'#233'ances'
        Caption = 'Ech'#233'ances'
        DisplayMode = dmGlyphOnly
        Layout = blGlyphTop
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BEcheClick
        OnMouseEnter = GereTagEnter
        OnMouseExit = GereTagExit
        GlobalIndexImage = 'Z0041_S16G2'
      end
      object BGenereTva: TToolbarButton97
        Tag = 1
        Left = 87
        Top = 0
        Width = 27
        Height = 24
        Hint = 'G'#233'n'#233'ration TVA'
        Caption = 'G'#233'n'#233'ration TVA'
        DisplayMode = dmGlyphOnly
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BGenereTvaClick
        OnMouseEnter = GereTagEnter
        OnMouseExit = GereTagExit
        GlobalIndexImage = 'Z0775_S16G1'
        IsControl = True
      end
      object BControleTVA: TToolbarButton97
        Tag = 1
        Left = 114
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Contr'#244'le de TVA'
        Caption = 'Contr'#244'le TVA'
        DisplayMode = dmGlyphOnly
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
        OnClick = BControleTVAClick
        OnMouseEnter = GereTagEnter
        OnMouseExit = GereTagExit
        GlobalIndexImage = 'Z0584_S16G1'
        IsControl = True
      end
      object BComplement: TToolbarButton97
        Tag = 1
        Left = 168
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Informations compl'#233'mentaires'
        Caption = 'Compl'#233'ments'
        DisplayMode = dmGlyphOnly
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
        OnMouseEnter = GereTagEnter
        OnMouseExit = GereTagExit
        GlobalIndexImage = 'Z0105_S16G2'
        IsControl = True
      end
      object BChercher: TToolbarButton97
        Tag = 1
        Left = 195
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Rechercher dans la pi'#232'ce'
        Caption = 'Chercher'
        DisplayMode = dmGlyphOnly
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
        OnMouseEnter = GereTagEnter
        OnMouseExit = GereTagExit
        GlobalIndexImage = 'Z0077_S16G1'
        IsControl = True
      end
      object BMenuZoom: TToolbarButton97
        Tag = -100
        Left = 228
        Top = 0
        Width = 37
        Height = 24
        Hint = 'Zooms'
        Caption = 'Zooms'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = POPZ
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
      object BPrev: TToolbarButton97
        Left = 271
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Ecriture/Ligne pr'#233'c'#233'dente'
        Caption = 'Pr'#233'c'#233'dent'
        DisplayMode = dmGlyphOnly
        Layout = blGlyphTop
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BPrevClick
        OnMouseEnter = GereTagEnter
        OnMouseExit = GereTagExit
        GlobalIndexImage = 'Z0171_S16G1'
      end
      object BNext: TToolbarButton97
        Left = 298
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Ecriture/Ligne suivante'
        Caption = 'Suivant'
        DisplayMode = dmGlyphOnly
        Layout = blGlyphTop
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BNextClick
        OnMouseEnter = GereTagEnter
        OnMouseExit = GereTagExit
        GlobalIndexImage = 'Z0319_S16G1'
      end
      object ToolbarSep971: TToolbarSep97
        Left = 81
        Top = 0
      end
      object ToolbarSep972: TToolbarSep97
        Left = 222
        Top = 0
      end
      object Sep97: TToolbarSep97
        Left = 265
        Top = 0
      end
      object BVentil: TToolbarButton97
        Tag = 1
        Left = 54
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Ventilation analytique'
        Caption = 'Analytiques'
        DisplayMode = dmGlyphOnly
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
        OnMouseEnter = GereTagEnter
        OnMouseExit = GereTagExit
        GlobalIndexImage = 'Z0133_S16G1'
        IsControl = True
      end
      object BModifTva: TToolbarButton97
        Tag = 1
        Left = 141
        Top = 0
        Width = 27
        Height = 24
        Hint = 'El'#233'ments de TVA / Bases Hors-taxes'
        Caption = 'TVA'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BModifTvaClick
        OnMouseEnter = GereTagEnter
        OnMouseExit = GereTagExit
        GlobalIndexImage = 'Z0543_S16G2'
        IsControl = True
      end
      object BScan: TToolbarButton97
        Left = 325
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Visualiser le document associ'#233
        DisplayMode = dmGlyphOnly
        Enabled = False
        Layout = blGlyphTop
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BScanClick
        GlobalIndexImage = 'Z0864_S16G1'
      end
    end
  end
  object PFEN: TPanel
    Left = 9
    Top = 9
    Width = 583
    Height = 312
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object GS: THGrid
      Left = 0
      Top = 53
      Width = 583
      Height = 209
      Align = alClient
      BorderStyle = bsNone
      ColCount = 10
      Ctl3D = True
      DefaultRowHeight = 18
      FixedCols = 4
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing]
      ParentCtl3D = False
      TabOrder = 1
      OnDblClick = GSDblClick
      OnEnter = GSEnter
      OnExit = GSExit
      OnKeyPress = GSKeyPress
      OnMouseDown = GSMouseDown
      OnSetEditText = GSSetEditText
      SortedCol = -1
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      OnRowEnter = GSRowEnter
      OnRowExit = GSRowExit
      OnCellEnter = GSCellEnter
      OnCellExit = GSCellExit
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
      ColWidths = (
        0
        0
        0
        22
        66
        77
        99
        154
        77
        77)
      RowHeights = (
        18
        18)
    end
    object PForce: TPanel
      Left = 4
      Top = 272
      Width = 577
      Height = 54
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 3
      Visible = False
      object Label2: TLabel
        Left = 70
        Top = 12
        Width = 331
        Height = 13
        Caption = 'La pi'#232'ce est d'#233's'#233'quilibr'#233'e en devise pivot : l'#39#233'cart est de '
      end
      object Label3: TLabel
        Left = 43
        Top = 32
        Width = 500
        Height = 13
        Caption = 
          'Pour l'#39#233'quilibrer, choisissez une ligne et validez par <<ENTREE>' +
          '> ou  <<Double-Click>>'
      end
      object Image1: TImage
        Left = 4
        Top = 8
        Width = 34
        Height = 37
        Center = True
        Picture.Data = {
          055449636F6E0000010001002020100000000000E80200001600000028000000
          2000000040000000010004000000000080020000000000000000000000000000
          0000000000000000000080000080000000808000800000008000800080800000
          C0C0C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000
          FFFFFF0000000888888888888888888888888800000088888888888888888888
          888888800030000000000000000000000008888803BBBBBBBBBBBBBBBBBBBBBB
          BB7088883BBBBBBBBBBBBBBBBBBBBBBBBBB708883BBBBBBBBBBBBBBBBBBBBBBB
          BBBB08883BBBBBBBBBBBB7007BBBBBBBBBBB08803BBBBBBBBBBBB0000BBBBBBB
          BBB7088003BBBBBBBBBBB0000BBBBBBBBBB0880003BBBBBBBBBBB7007BBBBBBB
          BB708800003BBBBBBBBBBBBBBBBBBBBBBB088000003BBBBBBBBBBB0BBBBBBBBB
          B70880000003BBBBBBBBB707BBBBBBBBB08800000003BBBBBBBBB303BBBBBBBB
          7088000000003BBBBBBBB000BBBBBBBB0880000000003BBBBBBB70007BBBBBB7
          08800000000003BBBBBB30003BBBBBB088000000000003BBBBBB00000BBBBB70
          880000000000003BBBBB00000BBBBB08800000000000003BBBBB00000BBBB708
          8000000000000003BBBB00000BBBB0880000000000000003BBBB00000BBB7088
          00000000000000003BBB70007BBB088000000000000000003BBBBBBBBBB70880
          000000000000000003BBBBBBBBB08800000000000000000003BBBBBBBB708800
          0000000000000000003BBBBBBB0880000000000000000000003BBBBBB7088000
          00000000000000000003BBBBB088000000000000000000000003BBBB70800000
          000000000000000000003BB70000000000000000000000000000033300000000
          00000000F8000003F0000001C000000080000000000000000000000000000001
          000000018000000380000003C0000007C0000007E000000FE000000FF000001F
          F000001FF800003FF800003FFC00007FFC00007FFE0000FFFE0000FFFF0001FF
          FF0001FFFF8003FFFF8003FFFFC007FFFFC007FFFFE00FFFFFE01FFFFFF07FFF
          FFF8FFFF}
        Stretch = True
      end
      object HForce: THNumEdit
        Tag = 2
        Left = 410
        Top = 8
        Width = 97
        Height = 21
        Color = clBtnFace
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        NumericType = ntDC
        UseRounding = True
        Validate = False
      end
    end
    object PEntete: THPanel
      Left = 0
      Top = 0
      Width = 583
      Height = 53
      Align = alTop
      FullRepaint = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      OnEnter = PEnteteEnter
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
      object H_NUMEROPIECE: THLabel
        Left = 397
        Top = 30
        Width = 48
        Height = 13
        Caption = 'N'#176' d'#233'finitif'
        Transparent = True
        Visible = False
      end
      object Bevel1: TBevel
        Left = 477
        Top = 28
        Width = 81
        Height = 20
      end
      object H_JOURNAL: THLabel
        Left = 4
        Top = 8
        Width = 34
        Height = 13
        Caption = 'Journal'
        FocusControl = E_JOURNAL
        Transparent = True
        OnDblClick = H_JOURNALDblClick
      end
      object H_DATECOMPTABLE: THLabel
        Left = 396
        Top = 8
        Width = 75
        Height = 13
        Caption = 'Date comptable'
        Transparent = True
      end
      object H_NATUREPIECE: THLabel
        Left = 193
        Top = 8
        Width = 61
        Height = 13
        Caption = 'Nature pi'#232'ce'
        FocusControl = E_NATUREPIECE
        Transparent = True
      end
      object H_DEVISE: THLabel
        Left = 4
        Top = 30
        Width = 33
        Height = 13
        Caption = 'Devise'
        FocusControl = E_DEVISE
        Transparent = True
      end
      object E_NUMEROPIECE: THLabel
        Left = 478
        Top = 30
        Width = 79
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
        Left = 193
        Top = 30
        Width = 65
        Height = 13
        Caption = 'Etablissement'
        FocusControl = E_ETABLISSEMENT
        Transparent = True
        OnDblClick = H_ETABLISSEMENTDblClick
      end
      object H_EXERCICE: THLabel
        Left = 563
        Top = 6
        Width = 14
        Height = 13
        Caption = '(N)'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Transparent = True
      end
      object BPopTva: TPopButton97
        Left = 562
        Top = 27
        Width = 31
        Height = 22
        Hint = 'Exigibilit'#233' de la TVA'
        GroupIndex = 1
        Glyph.Data = {
          36030000424D3603000000000000360000002800000010000000100000000100
          1800000000000003000000000000000000000000000000000000FF00FFFF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
          00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
          FF00FFFF00FFFF00FF7F7F7F7F7F7F7F7F7F7F7F7F7F7F7F7F7F7FFF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF7F7F7F00000000
          0000000000000000000000000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
          FF00FFFF00FFFF00FF7F7F7F000000000000FF00FFFF00FFFF00FF0000000000
          00FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF7F7F7F00000000
          0000FF00FFFF00FFFF00FFFF00FF000000000000FF00FFFF00FFFF00FFFF00FF
          FF00FFFF00FFFF00FF7F7F7F000000000000FF00FFFF00FFFF00FFFF00FF0000
          00000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF7F7F7F00000000
          0000FF00FFFF00FFFF00FFFF00FF000000000000FF00FFFF00FFFF00FFFF00FF
          FF00FFFF00FFFF00FF7F7F7F000000000000FF00FFFF00FFFF00FFFF00FF0000
          00000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF7F7F7F00000000
          0000FF00FFFF00FFFF00FFFF00FF000000000000FF00FFFF00FFFF00FFFF00FF
          FF00FFFF00FFFF00FF7F7F7F000000000000FF00FFFF00FFFF00FF0000000000
          00FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF00000000
          0000000000000000000000000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
          00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
          00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
        GlyphMask.Data = {00000000}
        Images = TVAImages
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        Items.Strings = (
          'Tva sur les d'#233'bits'
          'Tva sur les encaissements'
          'Tva mixte')
        ItemIndex = 1
        OnChange = BPopTvaChange
      end
      object H_NUMEROPIECE_: THLabel
        Left = 396
        Top = 30
        Width = 60
        Height = 13
        Caption = 'N'#176' provisoire'
        Transparent = True
      end
      object E_DEVISE_: THValComboBox
        Left = 43
        Top = 26
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 4
        Visible = False
        TagDispatch = 0
        DataType = 'TTDEVISEETAT'
      end
      object E_DEVISE: THValComboBox
        Left = 43
        Top = 26
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        OnChange = E_DEVISEChange
        TagDispatch = 0
        DataType = 'TTDEVISE'
      end
      object E_JOURNAL: THValComboBox
        Left = 43
        Top = 4
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = E_JOURNALChange
        TagDispatch = 0
        DataType = 'TTJALSAISIE'
      end
      object E_NATUREPIECE: THValComboBox
        Left = 261
        Top = 4
        Width = 129
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 1
        OnChange = E_NATUREPIECEChange
        TagDispatch = 0
      end
      object E_ETABLISSEMENT: THValComboBox
        Left = 261
        Top = 26
        Width = 129
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 5
        OnChange = E_ETABLISSEMENTChange
        TagDispatch = 0
        DataType = 'TTETABLISSEMENT'
      end
      object EURO: TEdit
        Left = 81
        Top = 14
        Width = 13
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 6
        Text = 'ECRITURE EN EURO'
        Visible = False
      end
      object compte: THEdit
        Left = 97
        Top = 14
        Width = 48
        Height = 21
        TabOrder = 7
        Visible = False
        TagDispatch = 0
        DataType = 'TZGENERAL'
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 477
        Top = 4
        Width = 80
        Height = 21
        EditMask = '!00/00/0000;1;_'
        MaxLength = 10
        TabOrder = 2
        Text = '  /  /    '
        OnChange = E_DATECOMPTABLEChange
        OnEnter = E_DATECOMPTABLEEnter
        OnExit = E_DATECOMPTABLEExit
        TagDispatch = 0
        OpeType = otDate
        ControlerDate = True
      end
    end
    object PCreerGuide: TPanel
      Left = 252
      Top = 172
      Width = 321
      Height = 97
      TabOrder = 4
      Visible = False
      object H_TitreGuide: TLabel
        Left = 4
        Top = 15
        Width = 313
        Height = 18
        Alignment = taCenter
        AutoSize = False
        Caption = 'Enregistrement de la pi'#232'ce en tant que guide'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
      end
      object H_NOMGUIDE: THLabel
        Left = 8
        Top = 40
        Width = 28
        Height = 13
        Caption = 'Guide'
        FocusControl = FNOMGUIDE
      end
      object PFenGuide: TPanel
        Left = 1
        Top = 1
        Width = 319
        Height = 9
        Align = alTop
        BevelInner = bvLowered
        BevelOuter = bvNone
        Color = clActiveCaption
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
      end
      object BCValide: THBitBtn
        Tag = 1
        Left = 253
        Top = 62
        Width = 28
        Height = 27
        Hint = 'Enregistrer le guide'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = BCValideClick
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
        Margin = 2
        NumGlyphs = 2
        Spacing = -1
        IsControl = True
      end
      object BCAbandon: THBitBtn
        Tag = 1
        Left = 284
        Top = 62
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnClick = BCAbandonClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object CMontantsGuide: TCheckBox
        Left = 6
        Top = 64
        Width = 131
        Height = 22
        Alignment = taLeftJustify
        Caption = 'Enregistrer les montants '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clGray
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
      end
      object FNOMGUIDE: TEdit
        Left = 44
        Top = 36
        Width = 269
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 35
        TabOrder = 0
      end
    end
    object BZoom: THBitBtn
      Tag = 100
      Left = 52
      Top = 96
      Width = 28
      Height = 27
      Hint = 'Voir compte'
      Caption = 'ZCpt'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 5
      Visible = False
      OnClick = BZoomClick
      Margin = 2
      Spacing = -1
      IsControl = True
    end
    object BZoomJournal: THBitBtn
      Tag = 100
      Left = 86
      Top = 96
      Width = 28
      Height = 27
      Hint = 'Voir journal'
      Caption = 'Jal'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 6
      Visible = False
      OnClick = BZoomJournalClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BZoomDevise: THBitBtn
      Tag = 100
      Left = 119
      Top = 96
      Width = 28
      Height = 27
      Hint = 'Voir devise'
      Caption = 'Dev'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 7
      Visible = False
      OnClick = BZoomDeviseClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BZoomEtabl: THBitBtn
      Tag = 100
      Left = 153
      Top = 96
      Width = 28
      Height = 27
      Hint = 'Voir '#233'tablissement'
      Caption = 'Etab'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 8
      Visible = False
      OnClick = BZoomEtablClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BScenario: THBitBtn
      Tag = 100
      Left = 187
      Top = 96
      Width = 28
      Height = 27
      Hint = 'Voir sc'#233'nario'
      Caption = 'Scen'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 9
      Visible = False
      OnClick = BScenarioClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BDernPieces: THBitBtn
      Tag = 100
      Left = 221
      Top = 96
      Width = 28
      Height = 27
      Hint = 'Liste des derni'#232'res '#233'critures'
      Caption = 'Dern'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 10
      Visible = False
      OnClick = BDernPiecesClick
      Margin = 0
      Spacing = -1
      IsControl = True
    end
    object BChancel: THBitBtn
      Tag = 100
      Left = 254
      Top = 96
      Width = 28
      Height = 27
      Hint = 'Historique des taux'
      Caption = 'Chan'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 11
      Visible = False
      OnClick = BChancelClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object PPied: THPanel
      Left = 0
      Top = 262
      Width = 583
      Height = 50
      Align = alBottom
      BevelOuter = bvNone
      BorderStyle = bsSingle
      Enabled = False
      FullRepaint = False
      TabOrder = 2
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
      object H_SOLDE: THLabel
        Left = 439
        Top = 28
        Width = 33
        Height = 13
        Caption = 'Solde'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object FlashGuide: TFlashingLabel
        Left = 343
        Top = 28
        Width = 40
        Height = 13
        Caption = 'GUIDE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        Visible = False
      end
      object Bevel2: TBevel
        Left = 356
        Top = 4
        Width = 116
        Height = 17
      end
      object Bevel3: TBevel
        Left = 478
        Top = 4
        Width = 116
        Height = 17
      end
      object Bevel4: TBevel
        Left = 478
        Top = 27
        Width = 116
        Height = 17
      end
      object HConf: TToolbarButton97
        Left = 386
        Top = 22
        Width = 25
        Height = 24
        Hint = 'Ecriture confidentielle'
        Enabled = False
        NoBorder = True
        ParentShowHint = False
        ShowHint = True
        Visible = False
        GlobalIndexImage = 'Z0141_S16G2'
      end
      object ISigneEuro: TImage
        Left = 416
        Top = 27
        Width = 16
        Height = 16
        AutoSize = True
        Picture.Data = {
          07544269746D6170F6000000424DF60000000000000076000000280000001000
          0000100000000100040000000000800000000000000000000000100000001000
          0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
          C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
          FF00FFFFFFFFFFFFFFFFFFFFFFF44444444FFFFFF44444444444FFFF4444FFFF
          FFF4FFFF444FFFFFFFFFFFFF44FFFFFFFFFFF44444444444FFFFFF4444444444
          4FFFFFF44FFFFFFFFFFFF444444444444FFFFF444444444444FFFFFF44FFFFFF
          FFF4FFFF444FFFFFFF44FFFFF444FFFFF444FFFFFF4444444444FFFFFFF44444
          4FF4}
        Stretch = True
        Transparent = True
        Visible = False
      end
      object G_LIBELLE: THLabel
        Left = 4
        Top = 6
        Width = 56
        Height = 13
        Caption = 'G_LIBELLE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Transparent = True
      end
      object T_LIBELLE: THLabel
        Left = 4
        Top = 27
        Width = 55
        Height = 13
        Caption = 'T_LIBELLE'
        Transparent = True
      end
      object LSA_SoldeG: THLabel
        Left = 66
        Top = 10
        Width = 61
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_SoldeG'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Transparent = True
      end
      object LSA_SoldeT: THLabel
        Left = 133
        Top = 10
        Width = 60
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_SoldeT'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Transparent = True
      end
      object LSA_Solde: THLabel
        Left = 201
        Top = 10
        Width = 63
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_Solde'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object LSA_TotalCredit: THLabel
        Left = 150
        Top = 27
        Width = 93
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_TotalCredit'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object LSA_TotalDebit: THLabel
        Left = 68
        Top = 26
        Width = 90
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_TotalDebit'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object SA_TOTALDEBIT: THNumEdit
        Tag = 1
        Left = 357
        Top = 5
        Width = 114
        Height = 15
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 4
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object SA_SOLDEG: THNumEdit
        Tag = 2
        Left = 238
        Top = 6
        Width = 101
        Height = 14
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        ParentCtl3D = False
        TabOrder = 0
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        NumericType = ntDC
        UseRounding = True
        Validate = False
      end
      object SA_SOLDET: THNumEdit
        Tag = 2
        Left = 238
        Top = 27
        Width = 101
        Height = 15
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        ParentCtl3D = False
        TabOrder = 2
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        NumericType = ntDC
        UseRounding = True
        Validate = False
      end
      object SA_TOTALCREDIT: THNumEdit
        Tag = 1
        Left = 479
        Top = 5
        Width = 114
        Height = 15
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object SA_SOLDE: THNumEdit
        Tag = 1
        Left = 479
        Top = 28
        Width = 114
        Height = 15
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 3
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        NumericType = ntDC
        UseRounding = True
        Validate = False
      end
    end
    object BInsert: THBitBtn
      Tag = 102
      Left = 360
      Top = 96
      Width = 28
      Height = 27
      Hint = 'Ins'#233'rer ligne'
      Caption = 'Ins'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 12
      Visible = False
      OnClick = BInsertClick
      Margin = 2
      NumGlyphs = 2
    end
    object BSDel: THBitBtn
      Tag = 102
      Left = 392
      Top = 96
      Width = 28
      Height = 27
      Hint = 'Supprimer ligne'
      Caption = 'Supprimer ligne'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 13
      Visible = False
      OnClick = BSDelClick
      Margin = 2
    end
    object BProrata: THBitBtn
      Tag = 1000
      Left = 324
      Top = 96
      Width = 28
      Height = 27
      Hint = 'Recalcul au prorata'
      Caption = 'Pror'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 14
      Visible = False
      OnClick = BProrataClick
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object BGuide: THBitBtn
      Tag = 101
      Left = 512
      Top = 96
      Width = 28
      Height = 27
      Hint = 'Appeler un guide'
      Caption = 'AGui'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 15
      Visible = False
      OnClick = BGuideClick
      Margin = 0
      Spacing = -1
      IsControl = True
    end
    object BCreerGuide: THBitBtn
      Tag = 101
      Left = 544
      Top = 96
      Width = 28
      Height = 27
      Hint = 'Cr'#233'er un guide'
      Caption = 'CGui'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 16
      Visible = False
      OnClick = BCreerGuideClick
      Margin = 0
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object BMenuGuide: THBitBtn
      Tag = -101
      Left = 464
      Top = 96
      Width = 45
      Height = 27
      Hint = 'Guides'
      Caption = 'Guides'
      TabOrder = 17
      Visible = False
    end
    object BModifSerie: THBitBtn
      Tag = 103
      Left = 48
      Top = 128
      Width = 29
      Height = 25
      Hint = 'Modification en s'#233'rie libell'#233's-ref.'
      Caption = 'ModZ'
      TabOrder = 18
      Visible = False
      OnClick = BModifSerieClick
    end
    object BSwapPivot: THBitBtn
      Tag = 100
      Left = 288
      Top = 96
      Width = 28
      Height = 27
      Hint = 'Basculement Devise saisie / Monnaie pivot'
      Caption = 'Swap'
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 19
      Visible = False
      OnClick = BSwapPivotClick
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object BSaisTaux: THBitBtn
      Tag = 1000
      Left = 84
      Top = 128
      Width = 29
      Height = 25
      Hint = 'Saisie du taux de la devise'
      Caption = 'Taux'
      Enabled = False
      TabOrder = 20
      Visible = False
      OnClick = BSaisTauxClick
    end
    object BModifRIB: THBitBtn
      Tag = 103
      Left = 152
      Top = 128
      Width = 28
      Height = 27
      Hint = 'Affectation du RIB'
      Caption = 'RIB'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 21
      Visible = False
      OnClick = BModifRIBClick
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object BCheque: THBitBtn
      Left = 250
      Top = 130
      Width = 27
      Height = 24
      Hint = 'Editer les lettres-ch'#232'ques'
      Caption = 'Chq'
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 22
      Visible = False
      OnClick = BChequeClick
    end
    object BChoixRegime: THBitBtn
      Tag = 103
      Left = 347
      Top = 129
      Width = 29
      Height = 25
      Hint = 'Changement de r'#233'gime fiscal'
      Caption = 'R'#233'gime fiscal'
      TabOrder = 23
      Visible = False
      OnClick = BChoixRegimeClick
    end
    object BZoomTP: THBitBtn
      Tag = 100
      Left = 546
      Top = 128
      Width = 28
      Height = 27
      Hint = 'Voir pi'#232'ce associ'#233'e du factur'#233'/payeur'
      Caption = 'Pi'#232'ce payeur'
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 24
      Visible = False
      OnClick = BZoomTPClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BMenuAction: THBitBtn
      Tag = -102
      Left = 384
      Top = 144
      Width = 41
      Height = 27
      Hint = 'Actions lignes'
      Caption = 'Actions'
      TabOrder = 25
      Visible = False
    end
    object BModifs: THBitBtn
      Tag = -103
      Left = 432
      Top = 144
      Width = 41
      Height = 27
      Hint = 'Modifications'
      Caption = 'Modifications'
      TabOrder = 26
      Visible = False
    end
    object BPieceGC: THBitBtn
      Tag = 100
      Left = 86
      Top = 156
      Width = 28
      Height = 27
      Hint = 'Voir pi'#232'ce commerciale'
      Caption = 'Pi'#232'ce'
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 27
      Visible = False
      OnClick = BPieceGCClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object bLignesGC: THBitBtn
      Tag = 100
      Left = 54
      Top = 156
      Width = 28
      Height = 27
      Hint = 'Voir lignes pi'#232'ces commerciales'
      Caption = 'Lignes'
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 28
      Visible = False
      OnClick = BLignesGCClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BSCANPDF: THBitBtn
      Tag = 102
      Left = 8
      Top = 184
      Width = 28
      Height = 27
      Hint = 'Associer un fichier Pdf'
      Caption = 'BSCANPDF'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 29
      Visible = False
      OnClick = BSCANPDFClick
      Margin = 2
      NumGlyphs = 2
    end
  end
  object BZoomImmo: THBitBtn
    Tag = 100
    Left = 526
    Top = 140
    Width = 28
    Height = 27
    Hint = 'Voir immobilisation'
    Caption = 'Immo'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 5
    Visible = False
    OnClick = BZoomImmoClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BLibAuto: THBitBtn
    Tag = 102
    Left = 487
    Top = 152
    Width = 27
    Height = 24
    Hint = 'Libell'#233's automatiques'
    Caption = 'Libell'#233's'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 1
    Visible = False
    OnClick = BLibAutoClick
    Layout = blGlyphTop
    Margin = 0
    NumGlyphs = 2
    Spacing = -1
    IsControl = True
  end
  object BVidePiece: THBitBtn
    Tag = 102
    Left = 16
    Top = 156
    Width = 28
    Height = 27
    Hint = 'Supprimer toutes les lignes'
    Caption = 'Supp. toutes les lignes'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 7
    Visible = False
    OnClick = BVidePieceClick
    Margin = 2
  end
  object HLigne: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Vous devez renseigner un compte g'#233'n'#233'ral existant.;W;' +
        'O;O;O;'
      
        '1;?caption?;Vous devez renseigner un compte g'#233'n'#233'ral collectif;W;' +
        'O;O;O;'
      
        '2;?caption?;Vous devez renseigner un compte auxiliaire existant;' +
        'W;O;O;O;'
      '3;?caption?;Ce compte g'#233'n'#233'ral interdit sur ce journal. ;W;O;O;O;'
      '4;?caption?;Vous devez saisir un montant;W;O;O;O;'
      '5;?caption?;Le montant en devise pivot est nul;W;O;O;O;'
      
        '6;?caption?;Les natures des comptes renseign'#233's sont incompatible' +
        's;W;O;O;O;'
      
        '7;?caption?;Ce compte g'#233'n'#233'ral est interdit pour cette nature de ' +
        'pi'#232'ce;W;O;O;O;'
      
        '8;?caption?;Vous ne pouvez pas saisir des montants n'#233'gatifs;W;O;' +
        'O;O;'
      
        '9;?caption?;Le total des '#233'ch'#233'ances est incorrect. Vous devez les' +
        ' modifier.;W;O;O;O;'
      
        '10;?caption?;Les ventilations analytiques sont incorrectes. Vous' +
        ' devez les modifier.;W;O;O;O;'
      
        '11;?caption?;Les ventilations analytiques sur l'#39'axe 1 sont incor' +
        'rectes. Vous devez les modifier.;W;O;O;O;'
      
        '12;?caption?;Les ventilations analytiques sur l'#39'axe 2 sont incor' +
        'rectes. Vous devez les modifier.;W;O;O;O;'
      
        '13;?caption?;Les ventilations analytiques sur l'#39'axe 3 sont incor' +
        'rectes. Vous devez les modifier.;W;O;O;O;'
      
        '14;?caption?;Les ventilations analytiques sur l'#39'axe 4 sont incor' +
        'rectes. Vous devez les modifier.;W;O;O;O;'
      
        '15;?caption?;Les ventilations analytiques sur l'#39'axe 5 sont incor' +
        'rectes. Vous devez les modifier.;W;O;O;O;'
      '16;?caption?;Attention : ce compte g'#233'n'#233'ral est ferm'#233';E;O;O;O;'
      
        '17;?caption?;Ce compte est interdit au d'#233'bit sur ce journal;W;O;' +
        'O;O;'
      
        '18;?caption?;Ce compte est interdit au cr'#233'dit sur ce journal;W;O' +
        ';O;O;'
      
        '19;?caption?;Vous ne pouvez pas saisir sur le compte d'#39'ouverture' +
        ' de bilan;W;O;O;O;'
      
        '20;?caption?;Vous ne pouvez pas saisir sur un compte de charge o' +
        'u de produit en pi'#232'ce d'#39'A-Nouveau;W;O;O;O;'
      
        '21;?caption?;Certaines lignes de la pi'#232'ce sont incorrectes;W;O;O' +
        ';O;'
      
        '22;?caption?;Le montant que vous avez saisi est en dehors de la ' +
        'fourchette autoris'#233'e;W;O;O;O;'
      
        '23;?caption?;Le compte g'#233'n'#233'ral collectif associ'#233' n'#39'est pas autor' +
        'is'#233';W;O;O;O;'
      
        '24;?caption?;Ce compte de banque ne supporte pas la devise de la' +
        ' pi'#232'ce;W;O;O;O;'
      
        '25;?caption?;Vous ne pouvez pas modifier ce mouvement : il a '#233't'#233 +
        ' lettr'#233' ou point'#233';W;O;O;O;'
      
        '26;?caption?;Attention : le compte auxiliaire est en sommeil;E;O' +
        ';O;O;'
      
        '27;?caption?;Attention : la devise du tiers ne correspond pas '#224' ' +
        'celle de l'#39#233'criture saisie;E;O;O;O;'
      
        '28;?caption?;Attention : le sens de la ligne d'#39#233'criture ne corre' +
        'spond pas au sens du compte;W;O;O;O;'
      
        '29;?caption?;Attention : ce compte est ferm'#233'. Vous ne pouvez plu' +
        's l'#39'utiliser en saisie;E;O;O;O;'
      
        '30;?caption?;Attention : ce compte est vis'#233'. Vous ne pouvez plus' +
        ' l'#39'utiliser en saisie;E;O;O;O;'
      
        '31;?caption?;Les dates d'#39#233'ch'#233'ance doivent respecter la plage de ' +
        'saisie autoris'#233'e;W;O;O;O;'
      ' '
      ' '
      ' ')
    Left = 165
    Top = 233
  end
  object HPiece: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Votre pi'#232'ce est incorrecte : elle n'#39'est pas '#233'quilibr' +
        #233'e en devise;W;O;O;O;'
      
        '1;?caption?;Votre pi'#232'ce est incorrecte : vous devez saisir au mo' +
        'ins deux lignes;W;O;O;O;'
      
        '2;?caption?;Votre pi'#232'ce est incorrecte : elle n'#39'est pas '#233'quilibr' +
        #233'e en devise pivot;W;O;O;O;'
      '3;?caption?;Confirmez-vous l'#39'abandon de la saisie ?;Q;YN;N;N;'
      
        '4;?caption?;Confirmez-vous la modification de toutes les lignes ' +
        'de la pi'#232'ce ?;Q;YN;Y;Y;'
      
        '5;?caption?;Votre pi'#232'ce est incorrecte : la date saisie n'#39'est pa' +
        's valide;W;O;O;O;'
      
        '6;?caption?;Votre pi'#232'ce est incorrecte : la date saisie n'#39'est pa' +
        's sur un exercice ouvert;W;O;O;O;'
      
        '7;?caption?;Votre pi'#232'ce est incorrecte : la date saisie est inf'#233 +
        'rieure '#224' la date de cl'#244'ture p'#233'riodique;W;O;O;O;'
      
        '8;?caption?;Votre pi'#232'ce est incorrecte : la date saisie est inf'#233 +
        'rieure '#224' la date de cl'#244'ture d'#233'finitive;W;O;O;O;'
      
        '9;?caption?;Votre pi'#232'ce est incorrecte : la date saisie est inf'#233 +
        'rieure '#224' la limite autoris'#233'e;W;O;O;O;'
      
        '10;?caption?;Votre pi'#232'ce est incorrecte : la date saisie est sup' +
        #233'rieure '#224' la limite autoris'#233'e;W;O;O;O;'
      
        '11;?caption?;La pi'#232'ce est valid'#233'e : seule la consultation est au' +
        'toris'#233'e;E;O;O;O;'
      
        '12;?caption?;Abandon interdit : vous devez valider la pi'#232'ce;W;O;' +
        'O;O;'
      
        '13;?caption?;ATTENTION : certaines '#233'critures point'#233'es ou lettr'#233'e' +
        's n'#39'ont pas '#233't'#233' modifi'#233'es;E;O;O;O;'
      
        '14;?caption?;Confirmez-vous la suppression de toutes les lignes?' +
        ';Q;YN;N;N;'
      '15;?caption?;Aucun libell'#233' automatique n'#39'a '#233't'#233' choisi.;E;O;O;O;'
      
        '16;?caption?;La date que vous avez renseign'#233'e n'#39'est pas valide;W' +
        ';O;O;O;'
      
        '17;?caption?;La date que vous avez renseign'#233'e n'#39'est pas dans un ' +
        'exercice ouvert;W;O;O;O;'
      
        '18;?caption?;La date que vous avez renseign'#233'e est ant'#233'rieure '#224' u' +
        'ne cl'#244'ture;W;O;O;O;'
      
        '19;?caption?;La date que vous avez renseign'#233'e est ant'#233'rieure '#224' u' +
        'ne cl'#244'ture;W;O;O;O;'
      
        '20;?caption?;La date que vous avez renseign'#233'e est en dehors des ' +
        'limites autoris'#233'es;W;O;O;O;'
      '21;?caption?;Vous devez renseigner un intitul'#233' du guide;W;O;O;O;'
      '22;?caption?;Confirmez-vous la cr'#233'ation du guide ?;Q;YN;Y;Y;'
      
        '23;?caption?;Ce guide existe d'#233'j'#224'. Voulez-vous l'#39#233'craser ?;Q;YN;' +
        'N;N;'
      '24;?caption?;Voulez-vous lettrer vos '#233'ch'#233'ances ?;Q;YN;Y;Y;'
      
        '25;?caption?;Votre pi'#232'ce est incorrecte : elle doit comporter un' +
        'e ligne sur le compte de tr'#233'sorerie du journal;W;O;O;O;'
      
        '26;?caption?;Vous n'#39'avez pas le droit de saisir sur ce journal;W' +
        ';O;O;O;'
      
        '27;?caption?;Journal incorrect : vous n'#39'avez pas d'#233'fini de compt' +
        'eur de num'#233'rotation pour ce journal;W;O;O;O;'
      
        '28;?caption?;Voulez-vous enregistrer les modifications ? ;Q;YNC;' +
        'Y;Y;'
      
        '29;?caption?;Vous ne pouvez pas cr'#233'er de guide pour cette nature' +
        ' de journal ;W;O;O;O;'
      '30;?caption?;ATTENTION : Le journal est ferm'#233' (;E;O;O;O;'
      
        '31;?caption?;Recalcul au pro-rata interdit : il existe des ligne' +
        's d'#39#233'criture non modifiables;W;O;O;O;'
      
        '32;?caption?;ATTENTION : seule la consultation est autoris'#233'e. Ce' +
        'tte pi'#232'ce poss'#232'de des '#233'ch'#233'ances en cours de lettrage en saisie d' +
        'e tr'#233'sorerie;E;O;O;O;'
      
        '33;?caption?;Vous devez saisir une pi'#232'ce d'#39#233'cart de change sur u' +
        'n journal multi-devises;W;O;O;O;'
      
        '34;?caption?;Votre pi'#232'ce est incorrecte : l'#39#233'tablissement est su' +
        'pprim'#233' ou erron'#233';W;O;O;O;'
      '35;?caption?;Voulez-vous faire un export CFONB ?;Q;YN;Y;Y;'
      
        '36;?caption?;Certaines '#233'ch'#233'ances n'#39'ont pas de RIB renseign'#233'. L'#39'e' +
        'xport CFONB sera impossible ou incomplet. Confirmez-vous malgr'#233' ' +
        'tout la validation ?;Q;YN;Y;Y;'
      
        '37;?caption?;Voulez-vous faire une '#233'dition de bordereau d'#39'accomp' +
        'agnement ?;Q;YN;Y;Y;'
      
        '38;?caption?;La pi'#232'ce est sur une p'#233'riode cl'#244'tur'#233'e : seule la co' +
        'nsultation est autoris'#233'e;E;O;O;O;'
      
        '39;?caption?;Vous ne pouvez pas saisir en Euro sur une date ant'#233 +
        'rieure '#224' celle d'#39'entr'#233'e en vigueur;W;O;O;O;'
      
        '40;?caption?;Vous n'#39'avez pas le droit de saisir sur ce journal :' +
        ' seule la consultation est autoris'#233'e;E;O;O;O;'
      
        '41;?caption?;Confirmez-vous la date comptable de l'#39#233'criture ?;Q;' +
        'YN;Y;Y;'
      
        '42;?caption?;Les TVA g'#233'n'#233'r'#233'es ou calcul'#233'es peuvent devenir incor' +
        'rectes. Confirmez-vous l'#39'op'#233'ration ?;Q;YN;Y;Y;'
      
        '43;?caption?;Vous ne pouvez pas saisir sur les journaux d'#233'di'#233's a' +
        'ux tiers payeurs;W;O;O;O;'
      
        '44;?caption?;Il existe des doublons. Confirmez-vous la validatio' +
        'n de la pi'#232'ce ?;Q;YN;N;N;'
      
        '45;?caption?;Voulez-vous visualiser le premier doublon d'#233'tect'#233' ?' +
        ';Q;YN;N;N;'
      
        '46;?caption?;Voulez-vous cr'#233'er une fiche d'#39'immobilisation ?;Q;YN' +
        ';Y;N;'
      
        '47;?caption?;Voulez-vous g'#233'n'#233'rer les '#233'critures sur le tiers paye' +
        'ur ?;Q;YN;Y;N;'
      
        '48;?caption?;Vous n'#39'avez pas le droit de modifier les '#233'critures ' +
        ': seule la consultation est autoris'#233'e;E;O;O;O;'
      
        '49;?caption?;Vous n'#39'avez pas param'#233'tr'#233' de journal pour les tiers' +
        ' payeurs, le m'#233'canisme ne sera pas activ'#233';E;O;O;O;'
      
        '50;?caption?;Votre pi'#232'ce est incorrecte : elle doit comporter un' +
        'e ligne sur le compte de contrepartie du journal;W;O;O;O;'
      
        '51;?caption?;La pi'#232'ce contient des lignes '#224' z'#233'ro. Voulez-vous le' +
        's d'#233'truire?;Q;YN;N;N;'
      
        '52;?caption?;La pi'#232'ce est bloqu'#233'e (suite '#224' synchronisation de do' +
        'ssier): seule la consultation est autoris'#233'e;E;O;O;O;'
      
        '53;?caption?;Vous ne pouvez pas saisir sur cette devise '#224' cette ' +
        'date;W;O;O;O;'
      '')
    Left = 25
    Top = 233
  end
  object POPS: TPopupMenu
    AutoPopup = False
    OnPopup = POPSPopup
    Left = 168
    Top = 185
  end
  object HTitres: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Visualisation des '#233'critures courantes'
      'Visualisation des '#233'critures de simulation'
      'Visualisation des '#233'critures de pr'#233'vision'
      'Visualisation des '#233'critures de situation'
      'Visualisation des '#233'critures de r'#233'vision'
      'Visualisation des '#233'critures IFRS'
      'Visualisation des '#233'critures d'#39#224'-nouveaux'
      ''
      ''
      ''
      'Modification des '#233'critures courantes'
      'Modification des '#233'critures de simulation'
      'Modification des '#233'critures de pr'#233'vision'
      'Modification des '#233'critures de situation'
      'Modification des '#233'critures IFRS'
      'Modification des '#233'critures de r'#233'vision'
      ''
      ''
      ''
      ''
      'Saisie des '#233'critures courantes'
      'Saisie des '#233'critures de simulation'
      'Saisie des '#233'critures de pr'#233'vision'
      'Saisie des '#233'critures de situation'
      'Saisie des '#233'critures de r'#233'vision'
      'Saisie des '#233'critures IFRS'
      'Saisie des '#233'critures d'#39#224'-nouveaux'
      'G'#233'n'#233'ration des '#233'critures'
      '28;')
    Left = 71
    Top = 233
  end
  object FindSais: TFindDialog
    OnFind = FindSaisFind
    Left = 216
    Top = 185
  end
  object HDiv: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Choix d'#39'un guide de saisie'
      'Choix des libell'#233's automatiques'
      
        '2;?caption?;Voulez-vous saisir ce taux dans la table de chancell' +
        'erie ?;Q;YN;Y;N;O;'
      'ATTENTION : Pi'#232'ce non valid'#233'e'
      
        'Validation impossible! Des '#233'critures ou des comptes ont '#233't'#233' modi' +
        'fi'#233's par un autre utilisateur'
      
        'ATTENTION! Le num'#233'ro de pi'#232'ce n'#39'est pas d'#233'finitivement attribu'#233'.' +
        ' Abandonnez la saisie et r'#233'-essayez de nouveau'
      
        'ATTENTION! Probl'#232'me de num'#233'rotation. V'#233'rifiez votre compteur jou' +
        'rnal'
      'Mouvement non modifiable.'
      
        '8;?caption?;ATTENTION : Le taux en cours est de 1. Voulez-vous s' +
        'aisir ce taux dans la table de chancellerie;Q;YN;Y;N;O;'
      
        '9;Editions de r'#232'glements;Certaines lignes d'#39#233'critures sont multi' +
        '-ech'#233'ances, l'#39#233'dition est impossible.;W;O;O;O;'
      
        '10;Editions de r'#232'glements;Les lignes d'#39#233'ch'#233'ances portent des mod' +
        'es de paiement diff'#233'rents, l'#39#233'dition est impossible;W;O;O;O;'
      
        '11;?caption?;Voulez-vous '#233'diter le document interne des '#233'criture' +
        's cr'#233#233'es ?;Q;YN;Y;N;O;'
      
        '12;Emission de lettres-ch'#232'que;Vous n'#39'avez '#233'dit'#233' aucun ch'#232'que. L'#39 +
        'enregistrement de l'#39#233'criture est impossible;W;O;O;O;'
      'EURO'
      
        '14;?caption?;La parit'#233' fixe est mal renseign'#233'e. Voulez-vous la s' +
        'aisir ?;Q;YN;Y;N;O;'
      
        '15;?caption?;ATTENTION. La parit'#233' est incorrecte. Voulez-vous la' +
        ' renseigner ?;Q;YN;Y;N;O;'
      'Ecart de conversion'
      
        'ATTENTION : L'#39#233'criture est valid'#233'e sans g'#233'n'#233'ration de pi'#232'ces sur' +
        ' le tiers payeur.'
      'Choix du r'#233'gime (Actuel'
      
        'Lettrage impossible! Des '#233'critures ont '#233't'#233' modifi'#233'es ou lettr'#233'es' +
        ' par un autre utilisateur'
      '20;?caption?;L'#39'immobilisation n'#39'existe plus;W;O;O;O;'
      
        '21;Edition de lettres-traite;Certaines r'#233'f'#233'rences n'#39'ont pas '#233't'#233' ' +
        'mises '#224' jour;E;O;O;O;'
      'ATTENTION : Des conflits d'#39'acc'#232's ont '#233't'#233' d'#233'tect'#233's ! #13'
      
        'La pi'#232'ce est valid'#233'e, mais vous devez demander '#224' l'#39'administrateu' +
        'r de lancer un recalcul du solde des comptes'
      
        'ATTENTION : L'#39#233'criture est valid'#233'e sans suppression des pi'#232'ces s' +
        'ur le tiers payeur.'
      
        'ATTENTION : L'#39#233'criture est valid'#233'e sans g'#233'n'#233'ration des pi'#232'ces d'#39 +
        'escompte.'
      
        '26;?caption?;Validation impossible : Le guide est en cours de sa' +
        'isie;W;O;O;O;'
      '27')
    Left = 122
    Top = 229
  end
  object HTVA: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;G'#233'n'#233'ration de TVA impossible : la pi'#232'ce et le journa' +
        'l doivent '#234'tre de nature Vente ou Achat;E;O;O;O;'
      
        '1;?caption?;G'#233'n'#233'ration de TVA impossible : la pi'#232'ce doit poss'#233'de' +
        'r une ligne sur un compte auxiliaire;E;O;O;O;'
      
        '2;?caption?;G'#233'n'#233'ration de TVA impossible : la pi'#232'ce ne doit poss' +
        #233'der qu'#39'une seule ligne de Tiers;E;O;O;O;'
      
        '3;?caption?;G'#233'n'#233'ration de TVA impossible : la pi'#232'ce ne poss'#232'de p' +
        'as de ligne sur un compte HT;E;O;O;O;'
      
        '4;?caption?;Les lignes g'#233'n'#233'r'#233'es sont ventil'#233'es sur les sections ' +
        'd'#39'attente;A;O;O;O;'
      
        '5;?caption?;G'#233'n'#233'ration de TVA incoh'#233'rente : certains comptes HT ' +
        'sont incompatibles avec le param'#233'trage de TVA;E;O;O;O;'
      
        '6;?caption?;G'#233'n'#233'ration de TVA impossible : le r'#233'gime n'#39'est pas r' +
        'enseign'#233' pour le tiers ;E;O;O;O;'
      
        '7;?caption?;G'#233'n'#233'ration de TVA incompl'#232'te : certains comptes de T' +
        'VA ou TPF ne sont pas renseign'#233's;E;O;O;O;'
      
        '8;?caption?;Contr'#244'le de TVA impossible : vous devez '#234'tre positio' +
        'nn'#233' sur une ligne avec un compte HT;E;O;O;O;'
      
        '9;?caption?;Contr'#244'le de TVA impossible : la pi'#232'ce doit poss'#233'der ' +
        'une ligne sur un compte de Tiers;E;O;O;O;'
      
        '10;?caption?;Contr'#244'le de TVA impossible : certains comptes de TV' +
        'A ou TPF ne sont pas renseign'#233's;E;O;O;O;'
      
        '11;?caption?;La TVA de la pi'#232'ce est incorrecte. Confirmez-vous l' +
        'a validation de cette pi'#232'ce ?;Q;YNC;N;N;'
      ' ')
    Left = 213
    Top = 233
  end
  object POPZ: TPopupMenu
    AutoPopup = False
    Left = 24
    Top = 104
  end
  object HMTrad: THSystemMenu
    LockedCtrls.Strings = (
      'PCreerGuide')
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 432
    Top = 104
  end
  object LeTimer: TTimer
    Enabled = False
    OnTimer = LeTimerTimer
    Left = 217
    Top = 137
  end
  object TVAImages: THImageList
    GlobalIndexImages.Strings = (
      'Z0426_S16G1'
      'Z0378_S16G1'
      'Z0457_S16G1')
    Left = 298
    Top = 138
  end
end

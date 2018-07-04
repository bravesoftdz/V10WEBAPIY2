object FSaisieEff: TFSaisieEff
  Left = 254
  Top = 201
  Width = 612
  Height = 398
  HelpContext = 7245000
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = '-'
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
  OnKeyPress = FormKeyPress
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object DockTop: TDock97
    Left = 0
    Top = 0
    Width = 596
    Height = 9
  end
  object DockRight: TDock97
    Left = 587
    Top = 9
    Width = 9
    Height = 322
    Position = dpRight
  end
  object DockLeft: TDock97
    Left = 0
    Top = 9
    Width = 9
    Height = 322
    Position = dpLeft
  end
  object DockBottom: TDock97
    Left = 0
    Top = 331
    Width = 596
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
      Caption = 'Outils'
      CloseButton = False
      DockPos = 0
      TabOrder = 1
      object BEche: TToolbarButton97
        Tag = 1
        Left = 0
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
      object BComplement: TToolbarButton97
        Tag = 1
        Left = 54
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
        Left = 180
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
      object BRef: TToolbarButton97
        Left = 87
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Automate r'#233'f'#233'rence'
        Caption = 'R'#233'f'#233'rences'
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
        OnClick = BRefClick
        GlobalIndexImage = 'Z0165_S16G1'
        IsControl = True
      end
      object BLig: TToolbarButton97
        Left = 114
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Voir/Cacher Contrepartie '
        Caption = 'Contrepartie'
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
        OnClick = BLigClick
        GlobalIndexImage = 'Z0472_S16G1'
        IsControl = True
      end
      object BVentil: TToolbarButton97
        Tag = 1
        Left = 27
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Analytiques'
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
      object ToolbarSep971: TToolbarSep97
        Left = 300
        Top = 0
      end
      object ToolbarSep972: TToolbarSep97
        Left = 81
        Top = 0
      end
      object BParamListeSaisieEff: TToolbarButton97
        Left = 207
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Param'#233'trer la grille de saisie'
        Caption = 'Liste'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BParamListeSaisieEffClick
        GlobalIndexImage = 'Z0241_S16G1'
      end
      object BZoomMvtEff: TToolbarButton97
        Tag = 1
        Left = 141
        Top = 0
        Width = 39
        Height = 24
        Hint = 'Voir/S'#233'lectionner les factures du compte'
        Caption = 'Voir/S'#233'lectionner factures'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = POPMVT
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
        OnMouseEnter = GereTagEnter
        OnMouseExit = GereTagExit
        GlobalIndexImage = 'Z0795_S16G1'
        IsControl = True
      end
      object BVentilCtr: TToolbarButton97
        Tag = 1
        Left = 234
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Analytiques de la ligne de contrepartie'
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
        OnClick = BVentilCtrClick
        GlobalIndexImage = 'Z0133_S16G1'
        IsControl = True
      end
      object BMenuZoom: TToolbarButton97
        Tag = -100
        Left = 261
        Top = 0
        Width = 39
        Height = 24
        Hint = 'Menu zoom'
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
        OnMouseExit = BMenuZoomMouseExit
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
    end
  end
  object PFEN: TPanel
    Left = 9
    Top = 9
    Width = 578
    Height = 322
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
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
    object GS: THGrid
      Tag = 1
      Left = 0
      Top = 81
      Width = 578
      Height = 158
      Align = alClient
      BorderStyle = bsNone
      ColCount = 10
      DefaultRowHeight = 19
      FixedCols = 4
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing]
      TabOrder = 1
      OnDblClick = GSDblClick
      OnEnter = GSEnter
      OnExit = GSExit
      OnKeyPress = GSKeyPress
      OnMouseDown = GSMouseDown
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
      TwoColors = False
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
        64)
    end
    object PPied: THPanel
      Left = 0
      Top = 268
      Width = 578
      Height = 54
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
      object H_SOLDET: THLabel
        Left = 388
        Top = 33
        Width = 77
        Height = 13
        Caption = 'Solde Tr'#233'sorerie'
        Transparent = True
      end
      object Bevel2: TBevel
        Left = 360
        Top = 4
        Width = 120
        Height = 21
      end
      object Bevel1: TBevel
        Left = 482
        Top = 4
        Width = 120
        Height = 21
      end
      object Bevel7: TBevel
        Left = 268
        Top = 0
        Width = 81
        Height = 50
      end
      object HE_MODEP: THLabel
        Left = 270
        Top = 6
        Width = 75
        Height = 16
        Alignment = taCenter
        AutoSize = False
        Caption = 'HE_MODEP'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Transparent = True
      end
      object HE_DATEECHE: THLabel
        Left = 270
        Top = 27
        Width = 75
        Height = 16
        Alignment = taCenter
        AutoSize = False
        Caption = 'HE_DATEECHE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Transparent = True
      end
      object HConf: TToolbarButton97
        Left = 350
        Top = 25
        Width = 25
        Height = 24
        Hint = 'Ecriture confidentielle'
        Enabled = False
        NoBorder = True
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        GlobalIndexImage = 'Z0141_S16G2'
      end
      object ISigneEuro: TImage
        Left = 366
        Top = 32
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
      object Bevel3: TBevel
        Left = 482
        Top = 28
        Width = 120
        Height = 21
      end
      object G_LIBELLE: THLabel
        Left = 1
        Top = 2
        Width = 3
        Height = 13
        Caption = ' '
        Transparent = True
      end
      object T_LIBELLE: THLabel
        Left = 1
        Top = 17
        Width = 3
        Height = 13
        Caption = ' '
        Transparent = True
      end
      object E_RIB: THLabel
        Left = 1
        Top = 33
        Width = 3
        Height = 13
        Caption = ' '
        Transparent = True
      end
      object LSA_SOLDEG: THLabel
        Left = 148
        Top = 3
        Width = 70
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_SOLDEG'
        Transparent = True
        OnDblClick = H_JOURNALDblClick
      end
      object LSA_SOLDET: THLabel
        Left = 149
        Top = 19
        Width = 69
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_SOLDET'
        Transparent = True
        OnDblClick = H_JOURNALDblClick
      end
      object LSA_TOTALDEBIT: THLabel
        Left = 306
        Top = 3
        Width = 108
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_TOTALDEBIT'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        OnDblClick = H_JOURNALDblClick
      end
      object LSA_TOTALCREDIT: THLabel
        Left = 305
        Top = 15
        Width = 117
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_TOTALCREDIT'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        OnDblClick = H_JOURNALDblClick
      end
      object LSA_SOLDETR: THLabel
        Left = 288
        Top = 37
        Width = 89
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_SOLDETR'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        OnDblClick = H_JOURNALDblClick
      end
      object SA_SOLDETR: THNumEdit
        Tag = 2
        Left = 484
        Top = 31
        Width = 114
        Height = 13
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Enabled = False
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
      object SA_SOLDE: THNumEdit
        Tag = 1
        Left = 484
        Top = 30
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
        TabOrder = 5
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        NumericType = ntDC
        UseRounding = True
        Validate = False
      end
      object SA_SOLDEG: THNumEdit
        Tag = 2
        Left = 164
        Top = 2
        Width = 101
        Height = 17
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        ParentCtl3D = False
        TabOrder = 0
        Visible = False
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
        Left = 164
        Top = 17
        Width = 101
        Height = 18
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        ParentCtl3D = False
        TabOrder = 1
        Visible = False
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
        Left = 484
        Top = 8
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
        Visible = False
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object SA_TOTALDEBIT: THNumEdit
        Tag = 1
        Left = 362
        Top = 8
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
        TabOrder = 2
        Visible = False
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
    end
    object PEntete: THPanel
      Left = 0
      Top = 0
      Width = 578
      Height = 81
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
      object Bevel6: TBevel
        Left = 480
        Top = 56
        Width = 128
        Height = 20
      end
      object H_CONTREPARTIEGEN: THLabel
        Left = 200
        Top = 8
        Width = 39
        Height = 13
        Caption = 'Compte '
        Transparent = True
      end
      object E_NUMEROPIECE: THLabel
        Left = 484
        Top = 57
        Width = 119
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
      object H_NUMEROPIECE: THLabel
        Left = 420
        Top = 59
        Width = 73
        Height = 13
        AutoSize = False
        Caption = 'N'#176' d'#233'finitif'
        Transparent = True
        Visible = False
      end
      object H_NUMEROPIECE_: THLabel
        Left = 417
        Top = 58
        Width = 61
        Height = 13
        AutoSize = False
        Caption = 'N'#176' provisoire'
        Transparent = True
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
        Left = 417
        Top = 8
        Width = 23
        Height = 13
        Caption = 'Date'
        Transparent = True
      end
      object H_DEVISE: THLabel
        Left = 4
        Top = 58
        Width = 33
        Height = 13
        Caption = 'Devise'
        FocusControl = E_DEVISE
        Transparent = True
      end
      object H_ETABLISSEMENT: THLabel
        Left = 200
        Top = 58
        Width = 65
        Height = 13
        Caption = 'Etablissement'
        FocusControl = E_ETABLISSEMENT
        Transparent = True
      end
      object H_EXERCICE: THLabel
        Left = 552
        Top = 8
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
      object H_MODEPAIE: THLabel
        Left = 4
        Top = 32
        Width = 44
        Height = 13
        Caption = 'Paiement'
        FocusControl = E_MODEPAIE
        Transparent = True
        OnDblClick = H_MODEPAIEDblClick
      end
      object H_DATEECHEANCE: THLabel
        Left = 417
        Top = 32
        Width = 49
        Height = 13
        Caption = 'Ech'#233'ance'
        Transparent = True
      end
      object H_NATUREPIECE: THLabel
        Left = 200
        Top = 8
        Width = 32
        Height = 13
        Caption = 'Nature'
        Transparent = True
      end
      object H_TYPECTR: THLabel
        Left = 200
        Top = 32
        Width = 57
        Height = 13
        Caption = 'Contrepartie'
        FocusControl = E_TYPECTR
        Transparent = True
      end
      object BPopTva: TPopButton97
        Left = 552
        Top = 28
        Width = 31
        Height = 22
        Hint = 'Exigibilit'#233' de la Tva'
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
      object E_CONTREPARTIEGEN: THCpteEdit
        Left = 268
        Top = 4
        Width = 145
        Height = 21
        TabOrder = 1
        OnChange = E_CONTREPARTIEGENChange
        OnExit = E_CONTREPARTIEGENExit
        ZoomTable = tzGEffet
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object E_NATUREPIECE: THValComboBox
        Left = 268
        Top = 4
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
        OnChange = E_NATUREPIECEChange
        OnExit = E_NATUREPIECEExit
        TagDispatch = 0
      end
      object E_DEVISE: THValComboBox
        Left = 52
        Top = 54
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 8
        OnChange = E_DEVISEChange
        OnExit = E_DEVISEExit
        TagDispatch = 0
        DataType = 'TTDEVISE'
      end
      object E_JOURNAL: THValComboBox
        Left = 52
        Top = 4
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = E_JOURNALChange
        OnClick = E_JOURNALClick
        OnExit = E_MODEPAIEExit
        TagDispatch = 0
        DataType = 'TTJALEFFET'
      end
      object Cache: THCpteEdit
        Left = 568
        Top = 4
        Width = 16
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 10
        Text = '!!!'
        Visible = False
        ZoomTable = tzGeneral
        Vide = False
        Bourre = True
        okLocate = False
        SynJoker = False
      end
      object E_ETABLISSEMENT: THValComboBox
        Left = 268
        Top = 54
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 9
        OnExit = E_MODEPAIEExit
        TagDispatch = 0
        DataType = 'TTETABLISSEMENT'
      end
      object E_MODEPAIE: THValComboBox
        Left = 52
        Top = 29
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 4
        OnExit = E_MODEPAIEExit
        TagDispatch = 0
        Plus = 'MP_ENCAISSEMENT="MIX" OR MP_ENCAISSEMENT="ENC"'
        DataType = 'TTMODEPAIE'
      end
      object E_NUMREF: THNumEdit
        Left = 584
        Top = 4
        Width = 17
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 11
        Decimals = 0
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
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
        TabOrder = 7
        Text = 'ECRITURE EN EURO'
        Visible = False
      end
      object E_TYPECTR: THValComboBox
        Left = 268
        Top = 29
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 5
        OnChange = E_TYPECTRChange
        TagDispatch = 0
        DataType = 'TTTYPECONTREPARTIE'
      end
      object compte: THEdit
        Left = 105
        Top = 22
        Width = 48
        Height = 21
        TabOrder = 12
        Visible = False
        TagDispatch = 0
        DataType = 'TZGENERAL'
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 480
        Top = 4
        Width = 64
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 3
        Text = '  /  /    '
        OnEnter = E_DATECOMPTABLEEnter
        OnExit = E_DATECOMPTABLEExit
        TagDispatch = 0
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATEECHEANCE: THCritMaskEdit
        Left = 480
        Top = 29
        Width = 64
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 6
        Text = '  /  /    '
        OnExit = E_DATEECHEANCEExit
        TagDispatch = 0
        OpeType = otDate
        ControlerDate = True
      end
    end
    object BZoom: THBitBtn
      Tag = 100
      Left = 136
      Top = 144
      Width = 28
      Height = 27
      Hint = 'Voir compte'
      Caption = 'Cpte'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 4
      Visible = False
      OnClick = BZoomClick
      Margin = 2
      Spacing = -1
      IsControl = True
    end
    object BZoomJournal: THBitBtn
      Tag = 100
      Left = 170
      Top = 144
      Width = 28
      Height = 27
      Hint = 'Voir journal'
      Caption = 'Jal'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 5
      Visible = False
      OnClick = BZoomJournalClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BZoomEtabl: THBitBtn
      Tag = 100
      Left = 237
      Top = 144
      Width = 28
      Height = 27
      Hint = 'Voir '#233'tablissement'
      Caption = 'Etab'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 6
      Visible = False
      OnClick = BZoomEtablClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BScenario: THBitBtn
      Tag = 100
      Left = 271
      Top = 144
      Width = 28
      Height = 27
      Hint = 'Voir sc'#233'nario'
      Caption = 'Scen'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 7
      Visible = False
      OnClick = BScenarioClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object PPiedTreso: THPanel
      Left = 0
      Top = 239
      Width = 578
      Height = 29
      Align = alBottom
      BevelOuter = bvNone
      BorderStyle = bsSingle
      Caption = ' '
      FullRepaint = False
      TabOrder = 8
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
      object Bevel4: TBevel
        Left = 360
        Top = 2
        Width = 120
        Height = 21
      end
      object Bevel5: TBevel
        Left = 482
        Top = 2
        Width = 120
        Height = 21
      end
      object LE_DEBITTRESO: THLabel
        Left = 353
        Top = 7
        Width = 88
        Height = 13
        Alignment = taRightJustify
        Caption = 'LE_DEBITTRESO'
        Transparent = True
        OnDblClick = H_JOURNALDblClick
      end
      object LE_CREDITTRESO: THLabel
        Left = 420
        Top = 6
        Width = 96
        Height = 13
        Alignment = taRightJustify
        Caption = 'LE_CREDITTRESO'
        Transparent = True
        OnDblClick = H_JOURNALDblClick
      end
      object BMenuZoomT: TToolbarButton97
        Tag = -101
        Left = 84
        Top = 0
        Width = 40
        Height = 24
        Hint = 'Zoom Tr'#233'sorerie'
        DropdownArrow = True
        DropdownMenu = POPZT
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphRight
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnMouseEnter = BMenuZoomTMouseEnter
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
      object E_CREDITTRESO: THNumEdit
        Tag = 1
        Left = 484
        Top = 5
        Width = 114
        Height = 16
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 4
        Visible = False
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object E_DEBITTRESO: THNumEdit
        Tag = 1
        Left = 362
        Top = 5
        Width = 114
        Height = 16
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 3
        Visible = False
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object E_LIBTRESO: TEdit
        Left = 216
        Top = 2
        Width = 133
        Height = 21
        Color = clWhite
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentFont = False
        TabOrder = 2
        OnExit = E_LIBTRESOExit
      end
      object E_REFTRESO: TEdit
        Left = 128
        Top = 2
        Width = 85
        Height = 21
        Color = clWhite
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentFont = False
        TabOrder = 1
        OnExit = E_REFTRESOExit
      end
      object E_CPTTRESO: THPanel
        Left = 4
        Top = 2
        Width = 77
        Height = 21
        BevelOuter = bvLowered
        Enabled = False
        FullRepaint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
        BackGroundEffect = bdFond
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        ColorNb = 2
        TextEffect = tenone
      end
    end
    object BSwapPivot: THBitBtn
      Tag = 100
      Left = 372
      Top = 144
      Width = 28
      Height = 27
      Hint = 'Basculement devise/pivot'
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
      TabOrder = 9
      Visible = False
      OnClick = BSwapPivotClick
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object BModifSerie: THBitBtn
      Tag = 103
      Left = 404
      Top = 144
      Width = 29
      Height = 25
      Hint = 'Modification en s'#233'rie libell'#233's-ref.'
      Caption = 'ModZ'
      TabOrder = 10
      Visible = False
      OnClick = BModifSerieClick
    end
    object BZoomT: THBitBtn
      Tag = 101
      Left = 440
      Top = 144
      Width = 28
      Height = 27
      Hint = 'Voir compte'
      Caption = 'Treso'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 11
      Visible = False
      OnClick = BZoomTClick
      Margin = 2
      Spacing = -1
      IsControl = True
    end
    object BSwapEuro: THBitBtn
      Tag = 100
      Left = 542
      Top = 144
      Width = 28
      Height = 27
      Hint = 'Visualisation pluri-mon'#233'taire des montants'
      Caption = 'Swap'
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
      OnClick = BSwapEuroClick
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object BSaisTaux: THBitBtn
      Tag = 1000
      Left = 508
      Top = 144
      Width = 29
      Height = 25
      Hint = 'Saisie du taux de la devise'
      Caption = 'Taux'
      Enabled = False
      TabOrder = 13
      Visible = False
      OnClick = BSaisTauxClick
    end
    object BComplementT: THBitBtn
      Tag = 101
      Left = 476
      Top = 144
      Width = 28
      Height = 27
      Hint = 'Informations compl'#233'mentaires'
      Caption = 'T Infos'
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
      OnClick = BComplementTClick
      Margin = 2
      Spacing = -1
      IsControl = True
    end
  end
  object BModifRIB: THBitBtn
    Tag = 103
    Left = 110
    Top = 156
    Width = 27
    Height = 25
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
    TabOrder = 5
    Visible = False
    OnClick = BModifRIBClick
    Margin = 2
    NumGlyphs = 2
    Spacing = -1
    IsControl = True
  end
  object PRefAuto: TPanel
    Left = 200
    Top = 49
    Width = 213
    Height = 97
    TabOrder = 6
    Visible = False
    object H_TitreRefAuto: TLabel
      Left = 5
      Top = 11
      Width = 204
      Height = 18
      Alignment = taCenter
      AutoSize = False
      Caption = 'R'#233'f'#233'rence automatique'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = False
      ParentFont = False
    end
    object H_REF: THLabel
      Left = 9
      Top = 40
      Width = 50
      Height = 13
      Caption = 'R'#233'f'#233'rence'
      FocusControl = E_REF
    end
    object PFenGuide: TPanel
      Left = 1
      Top = 1
      Width = 211
      Height = 9
      Align = alTop
      BevelInner = bvLowered
      BevelOuter = bvNone
      Color = clTeal
      Enabled = False
      TabOrder = 2
    end
    object BRAValide: THBitBtn
      Tag = 1
      Left = 148
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
      TabOrder = 0
      OnClick = BRAValideClick
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
    object BRAAbandon: THBitBtn
      Tag = 1
      Left = 179
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
      TabOrder = 1
      OnClick = BRAAbandonClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0021_S16G1'
      IsControl = True
    end
    object E_REF: TEdit
      Left = 66
      Top = 36
      Width = 64
      Height = 21
      MaxLength = 35
      TabOrder = 3
    end
    object E_NUMDEP: TMaskEdit
      Left = 140
      Top = 36
      Width = 62
      Height = 21
      EditMask = '9999999999999;0; '
      MaxLength = 13
      TabOrder = 4
    end
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
      
        '1;?caption?;Vous devez renseigner un compte g'#233'n'#233'ral collectif.;W' +
        ';O;O;O;'
      
        '2;?caption?;Vous devez renseigner un compte auxiliaire existant.' +
        ';W;O;O;O;'
      '3;?caption?;Ce compte g'#233'n'#233'ral interdit sur ce journal.;W;O;O;O;'
      '4;?caption?;Vous devez saisir un montant.;W;O;O;O;'
      '5;?caption?;Le montant en devise pivot est nul.;W;O;O;O;'
      
        '6;?caption?;Les natures des comptes renseign'#233's sont incompatible' +
        's.;W;O;O;O;'
      
        '7;?caption?;Ce compte g'#233'n'#233'ral est interdit pour cette nature de ' +
        'pi'#232'ce.;W;O;O;O;'
      
        '8;?caption?;Vous ne pouvez pas saisir des montants n'#233'gatifs.;W;O' +
        ';O;O;'
      
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
      '16;?caption?;Attention : ce compte est ferm'#233'.;E;O;O;O;'
      
        '17;?caption?;Ce compte est interdit au d'#233'bit sur ce journal.;W;O' +
        ';O;O;'
      
        '18;?caption?;Ce compte est interdit au cr'#233'dit sur ce journal.;W;' +
        'O;O;O;'
      
        '19;?caption?;Vous ne pouvez pas saisir sur le compte d'#39'ouverture' +
        ' de bilan.;W;O;O;O;'
      
        '20;?caption?;Vous ne pouvez pas saisir sur un compte de charge o' +
        'u de produit en pi'#232'ce d'#39'A-Nouveau.;W;O;O;O;'
      
        '21;?caption?;Vous ne pouvez pas saisir sur le compte de contrepa' +
        'rtie de tr'#233'sorerie.;W;O;O;O;'
      
        '22;?caption?;Le montant que vous avez saisi est en dehors de la ' +
        'fourchette autoris'#233'e.;W;O;O;O;'
      
        '23;?caption?;Le compte g'#233'n'#233'ral collectif associ'#233' n'#39'est pas autor' +
        'is'#233'.;W;O;O;O;'
      
        '24;?caption?;Ce compte de banque ne supporte pas la devise de la' +
        ' pi'#232'ce.;W;O;O;O;'
      
        '25;?caption?;Vous ne pouvez pas modifier ce mouvement : il a '#233't'#233 +
        ' lettr'#233' ou point'#233'.;W;O;O;O;'
      
        '26;?caption?;Attention : le compte auxiliaire est en sommeil.;E;' +
        'O;O;O;'
      
        '27;?caption?;Attention : la devise du tiers ne correspond pas '#224' ' +
        'celle de l'#39#233'criture saisie.;E;O;O;O;'
      
        '28;?caption?;Vous ne pouvez pas saisir sur un compte de tiers no' +
        'n lettrable.;W;O;O;O;'
      
        '29;?caption?;Vous ne pouvez pas saisir sur le compte de contrepa' +
        'rtie.;W;O;O;O;'
      
        '30;?caption?;Attention : ce compte est ferm'#233' et sold'#233'. Vous ne p' +
        'ouvez plus l'#39'utiliser en saisie.;E;O;O;O;'
      '31;?caption?;Attention : Le RIB n'#39'est pas renseign'#233'.;E;O;O;O;'
      
        '32;?caption?;Vous ne pouvez pas saisir un compte fournisseur.;E;' +
        'O;O;O; '
      ' ')
    Left = 553
    Top = 188
  end
  object HPiece: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Votre pi'#232'ce est incorrecte : elle n'#39'est pas '#233'quilibr' +
        #233'e en devise.;W;O;O;O;'
      
        '1;?caption?;Votre pi'#232'ce est incorrecte : vous devez saisir au mo' +
        'ins deux lignes.;W;O;O;O;'
      
        '2;?caption?;Votre pi'#232'ce est incorrecte : elle n'#39'est pas '#233'quilibr' +
        #233'e en devise pivot.;W;O;O;O;'
      '3;?caption?;Confirmez-vous l'#39'abandon de la saisie ?;Q;YN;N;N;'
      
        '4;?caption?;L'#39'abandon de la saisie est impossible. Vous devez va' +
        'lider la pi'#232'ce.;W;O;O;O;'
      
        '5;?caption?;Votre pi'#232'ce est incorrecte : la date saisie n'#39'est pa' +
        's valide.;W;O;O;O;'
      
        '6;?caption?;Votre pi'#232'ce est incorrecte : la date saisie n'#39'est pa' +
        's sur un exercice ouvert.;W;O;O;O;'
      
        '7;?caption?;Votre pi'#232'ce est incorrecte : la date saisie est inf'#233 +
        'rieure '#224' la date de cl'#244'ture p'#233'riodique.;W;O;O;O;'
      
        '8;?caption?;Votre pi'#232'ce est incorrecte : la date saisie est inf'#233 +
        'rieure '#224' la date de cl'#244'ture d'#233'finitive.;W;O;O;O;'
      
        '9;?caption?;Votre pi'#232'ce est incorrecte : la date saisie est inf'#233 +
        'rieure '#224' la limite autoris'#233'e.;W;O;O;O;'
      
        '10;?caption?;Votre pi'#232'ce est incorrecte : la date saisie est sup' +
        #233'rieure '#224' la limite autoris'#233'e.;W;O;O;O;'
      
        '11;?caption?;La pi'#232'ce est valid'#233'e : seule la consultation est au' +
        'toris'#233'e.;E;O;O;O;'
      
        '12;?caption?;Abandon interdit : vous devez valider la pi'#232'ce.;W;O' +
        ';O;O;'
      '13;=========== LIBRE LIBRE LIBRE ========================='
      
        '14;?caption?;Confirmez-vous la suppression de toutes les lignes ' +
        '?;Q;YN;N;N;'
      '15;?caption?;Aucun libell'#233' automatique n'#39'a '#233't'#233' choisi.;E;O;O;O;'
      
        '16;?caption?;La date que vous avez renseign'#233'e n'#39'est pas valide.;' +
        'W;O;O;O;'
      
        '17;?caption?;La date que vous avez renseign'#233'e n'#39'est pas dans un ' +
        'exercice ouvert.;W;O;O;O;'
      
        '18;?caption?;La date que vous avez renseign'#233'e est ant'#233'rieure '#224' u' +
        'ne cl'#244'ture.;W;O;O;O;'
      
        '19;?caption?;La date que vous avez renseign'#233'e est ant'#233'rieure '#224' u' +
        'ne cl'#244'ture.;W;O;O;O;'
      
        '20;?caption?;La date que vous avez renseign'#233'e est en dehors des ' +
        'limites autoris'#233'es.;W;O;O;O;'
      
        '21;?caption?;Vous devez renseigner un intitul'#233' du guide.;W;O;O;O' +
        ';'
      '22;?caption?;Confirmez-vous la cr'#233'ation du guide ?;Q;YN;Y;Y;'
      
        '23;?caption?;Ce guide existe d'#233'j'#224'. Voulez-vous l'#39#233'craser ?;Q;YN;' +
        'N;N;'
      '24;?caption?;Voulez-vous lettrer vos '#233'ch'#233'ances ?;Q;YN;Y;Y;'
      ';;;'
      
        '26;?caption?;Vous n'#39'avez pas le droit de saisir sur ce journal.;' +
        'W;O;O;O;'
      
        '27;?caption?;Vous n'#39'avez pas d'#233'fini de compteur de num'#233'rotation ' +
        'pour ce journal.;W;O;O;O;'
      
        '28;?caption?;Confirmez-vous la modification de toutes les lignes' +
        ' de la pi'#232'ce ?;Q;YN;Y;Y;'
      '29;?caption?;Voulez-vous d'#233'lettrer vos '#233'ch'#233'ances ?;Q;YN;Y;Y;'
      
        '30;?caption?;Vous ne pouvez pas saisir en Euro sur une date ant'#233 +
        'rieure '#224' celle d'#39'entr'#233'e en vigueur.;W;O;O;O;'
      
        '31;?caption?;Confirmez-vous la date comptable de l'#39#233'criture ?;Q;' +
        'YN;Y;Y;'
      
        '32;?caption?;Les TVA g'#233'n'#233'r'#233'es ou calcul'#233'es peuvent devenir incor' +
        'rectes. Confirmez-vous l'#39'op'#233'ration ?;Q;YN;Y;Y;'
      
        '33;?caption?;Vous ne pouvez pas saisir sur les journaux d'#233'di'#233's a' +
        'ux tiers payeurs.;W;O;O;O;'
      ' '
      ' ')
    Left = 412
    Top = 188
  end
  object POPS: TPopupMenu
    OnPopup = POPSPopup
    Left = 140
    Top = 188
  end
  object HTitres: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Visualisation des '#233'critures courantes'
      'Visualisation des '#233'critures de simulation'
      'Visualisation des '#233'critures de pr'#233'vision'
      'Visualisation des '#233'critures de situation'
      ''
      ''
      ''
      ''
      ''
      ''
      'Modification des '#233'critures courantes'
      'Modification des '#233'critures de simulation'
      'Modification des '#233'critures de pr'#233'vision'
      'Modification des '#233'critures de situation'
      ''
      ''
      ''
      ''
      ''
      ''
      'Saisie des '#233'critures courantes'
      'Saisie des '#233'critures de simulation'
      'Saisie des '#233'critures de pr'#233'vision'
      'Saisie des '#233'critures de situation'
      ''
      ''
      ''
      ''
      ''
      ''
      'Saisie des '#233'critures de r'#232'glement'
      'Saisie des traites en retour d'#39'acceptation'
      'Saisie des ch'#232'ques'
      'Saisie des cartes bleues')
    Left = 459
    Top = 188
  end
  object FindSais: TFindDialog
    OnFind = FindSaisFind
    Left = 188
    Top = 188
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
      
        'Validation impossible. Certaines '#233'critures ont '#233't'#233' modifi'#233'es par' +
        ' un autre utilisateur.'
      
        'ATTENTION : Le num'#233'ro de pi'#232'ce n'#39'est pas d'#233'finitivement attribu'#233 +
        '. Abandonnez la saisie et r'#233'-essayez de nouveau.'
      
        'ATTENTION : Probl'#232'me de num'#233'rotation. V'#233'rifiez votre compteur jo' +
        'urnal.'
      'Mouvement lettr'#233' non modifiable.'
      '8;?caption?;Vous devez indiquer un mode de paiement.;W;O;O;O;O;'
      
        '9;?caption?;Vous devez indiquer un type de contrepartie (Pensez ' +
        'aussi '#224' mettre le journal '#224' jour).;W;O;O;O;O;'
      
        '10;?caption?;Le compte de tr'#233'sorerie est ventilable. Vous devez ' +
        'utiliser la saisie courante pour saisir sur ce compte.;W;O;O;O;O' +
        ';'
      
        '11;?caption?;Cette devise ne correspond pas '#224' la devise du compt' +
        'e de banque. Vous devez la modifier.;W;O;O;O;O;'
      
        'ATTENTION : Probl'#232'me fichier. Ressayez ou bien validez votre pi'#232 +
        'ce.'
      
        'ATTENTION : Probl'#232'me fichier. Veuillez revenir au menu, puis rel' +
        'ancer la fonction.'
      'RIB :'
      
        '15;?caption?;Vous devez indiquer un compte de contrepartie.;W;O;' +
        'O;O;O;'
      
        '16;?caption?;Le compte de contrepartie est collectif. Vous ne po' +
        'uvez pas utiliser un mode de contrepartie "Pied".;W;O;O;O;O;'
      '17;?caption?;Le compte de contrepartie n'#39'existe pas.;W;O;O;O;O;'
      'Solde Tr'#233'sorerie'
      'Solde Contrepartie'
      
        '20;?caption?;Le compte de contrepartie est ventilable. Vous deve' +
        'z utiliser la saisie courante pour saisir sur ce compte.;W;O;O;O' +
        ';O;'
      'EURO'
      
        '22;?caption?;Voulez-vous '#233'diter le document interne des '#233'criture' +
        's cr'#233#233'es ?;Q;YN;Y;N;O;'
      
        '23;?caption?;La parit'#233' fixe est mal renseign'#233'e. Voulez-vous la s' +
        'aisir ?;Q;YN;Y;N;O;'
      
        '24;?caption?;ATTENTION : La parit'#233' est incorrecte. Voulez-vous l' +
        'a renseigner ?;Q;YN;Y;N;O;'
      
        '25;?caption?;ATTENTION : Le taux en cours est de 1. Voulez-vous ' +
        'saisir ce taux dans la table de chancellerie ?;Q;YN;Y;N;O;'
      'Choix du r'#233'gime (Actuel'
      'ATTENTION : Des conflits d'#39'acc'#232's ont '#233't'#233' d'#233'tect'#233's. #13'
      
        'La pi'#232'ce est valid'#233'e, mais vous devez demander '#224' l'#39'administrateu' +
        'r de lancer un recalcul du solde des comptes.'
      
        '29;?caption?;Vous devez vous positionner sur une ligne vide pour' +
        ' acc'#233'der au param'#233'trage de la grille de saisie.;W;O;O;O;O;'
      
        '30;?caption?;Parmi les '#233'l'#233'ments s'#233'lectionn'#233's, certains ont d'#233'j'#224' ' +
        #233't'#233' pris en compte sur une autre ligne.;W;O;O;O;O;'
      
        '31;?caption?;ATTENTION : vous avez d'#233'j'#224' sel'#233'ctionn'#233' des '#233'criture' +
        's pour cette ligne (Ctrl-F6 pour voir la s'#233'lection).;W;O;O;O;O;'
      
        '32;?caption?;ATTENTION : La s'#233'lection va '#234'tre annul'#233'e. Confirmez' +
        '-vous ?;Q;YN;Y;N;O;'
      '33;')
    Left = 505
    Top = 188
  end
  object POPZ: TPopupMenu
    AutoPopup = False
    Left = 232
    Top = 188
  end
  object POPZT: TPopupMenu
    AutoPopup = False
    Left = 272
    Top = 188
  end
  object HMTrad: THSystemMenu
    LockedCtrls.Strings = (
      'PRefAuto')
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 364
    Top = 188
  end
  object POPMVT: TPopupMenu
    Left = 89
    Top = 189
    object POPZOOMFACT: TMenuItem
      Caption = 'S'#233'lectionner les factures du compte'
      OnClick = POPZOOMFACTClick
    end
    object BZOOMFACT2: TMenuItem
      Caption = 'Voir les factures du compte'
      OnClick = BZOOMFACT2Click
    end
  end
end

object FConsultCaisse: TFConsultCaisse
  Left = 219
  Top = 158
  HelpContext = 301100270
  Align = alClient
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Situation Flash'
  ClientHeight = 543
  ClientWidth = 727
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poDesktopCenter
  WindowState = wsMaximized
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnKeyDown = FormKeyDown
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object DockBottom: TDock97
    Left = 0
    Top = 0
    Width = 727
    Height = 36
    AllowDrag = False
    object ToolWindow971: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 727
      ClientAreaHeight = 32
      ClientAreaWidth = 727
      DockPos = 0
      DragHandleStyle = dhNone
      FullSize = True
      TabOrder = 0
      object BImprimer: TToolbarButton97
        Left = 548
        Top = 3
        Width = 56
        Height = 27
        Hint = 'Imprimer'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Imprimer'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          4E010000424D4E01000000000000760000002800000012000000120000000100
          040000000000D800000000000000000000001000000010000000000000000000
          BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00DDDDDDDDDDDD
          DDDDDD000000DDD00000000000DDDD000000DD0777777777070DDD000000D000
          000000000070DD000000D0777777FFF77000DD000000D077777799977070DD00
          0000D0000000000000770D000000D0777777777707070D000000DD0000000000
          70700D000000DDD0FFFFFFFF07070D000000DDDD0FCCCCCF0000DD000000DDDD
          0FFFFFFFF0DDDD000000DDDDD0FCCCCCF0DDDD000000DDDDD0FFFFFFFF0DDD00
          0000DDDDDD000000000DDD000000DDDDDDDDDDDDDDDDDD000000DDDDDDDDDDDD
          DDDDDD000000DDDDDDDDDDDDDDDDDD000000}
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BImprimerClick
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 668
        Top = 3
        Width = 56
        Height = 27
        Hint = 'Aide'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
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
          03030606030303030303030303030303030303FFFF0303030303030303030303
          0303030303060404060303030303030303030303030303F8F8FF030303030303
          030303030303030303FE06060403030303030303030303030303F8FF03F8FF03
          0303030303030303030303030303FE060603030303030303030303030303F8FF
          FFF8FF0303030303030303030303030303030303030303030303030303030303
          030303F8F8030303030303030303030303030303030304040603030303030303
          0303030303030303FFFF03030303030303030303030303030306060604030303
          0303030303030303030303F8F8F8FF0303030303030303030303030303FE0606
          0403030303030303030303030303F8FF03F8FF03030303030303030303030303
          03FE06060604030303030303030303030303F8FF03F8FF030303030303030303
          030303030303FE060606040303030303030303030303F8FF0303F8FF03030303
          0303030303030303030303FE060606040303030303030303030303F8FF0303F8
          FF030303030303030303030404030303FE060606040303030303030303FF0303
          F8FF0303F8FF030303030303030306060604030303FE06060403030303030303
          F8F8FF0303F8FF0303F8FF03030303030303FE06060604040406060604030303
          030303F8FF03F8FFFFFFF80303F8FF0303030303030303FE0606060606060606
          06030303030303F8FF0303F8F8F8030303F8FF030303030303030303FEFE0606
          060606060303030303030303F8FFFF030303030303F803030303030303030303
          0303FEFEFEFEFE03030303030303030303F8F8FFFFFFFFFFF803030303030303
          0303030303030303030303030303030303030303030303F8F8F8F8F803030303
          0303}
        Layout = blGlyphTop
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAideClick
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 608
        Top = 3
        Width = 56
        Height = 27
        Hint = 'Fermer'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Annuler'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          CE070000424DCE07000000000000360000002800000024000000120000000100
          1800000000009807000000000000000000000000000000000000008080008080
          0080800080800080800080800080800080800080800080800080800080800080
          8000808000808000808000808000808000808000808000808000808000808000
          8080008080008080008080008080008080008080008080008080008080008080
          0080800080800080800080800080800080808080808080800080800080800080
          8000808000808000808000808000808000808000808000808000808000808000
          8080008080008080008080FFFFFF008080008080008080008080008080008080
          0080800080800080800080800080800080800080800080800080800000FF0000
          800000808080800080800080800080800080800080800000FF80808000808000
          8080008080008080008080008080008080008080808080808080FFFFFF008080
          008080008080008080008080008080FFFFFF0080800080800080800080800080
          800080800080800000FF00008000008000008080808000808000808000808000
          00FF000080000080808080008080008080008080008080008080008080808080
          FFFFFF008080808080FFFFFF008080008080008080FFFFFF808080808080FFFF
          FF0080800080800080800080800080800080800000FF00008000008000008000
          00808080800080800000FF000080000080000080000080808080008080008080
          008080008080008080808080FFFFFF008080008080808080FFFFFF008080FFFF
          FF808080008080008080808080FFFFFF00808000808000808000808000808000
          80800000FF000080000080000080000080808080000080000080000080000080
          000080808080008080008080008080008080008080808080FFFFFF0080800080
          80008080808080FFFFFF808080008080008080008080008080808080FFFFFF00
          80800080800080800080800080800080800000FF000080000080000080000080
          0000800000800000800000808080800080800080800080800080800080800080
          80008080808080FFFFFF00808000808000808080808000808000808000808000
          8080FFFFFF808080008080008080008080008080008080008080008080008080
          0000FF0000800000800000800000800000800000808080800080800080800080
          80008080008080008080008080008080008080808080FFFFFF00808000808000
          8080008080008080008080FFFFFF808080008080008080008080008080008080
          0080800080800080800080800080800000800000800000800000800000808080
          8000808000808000808000808000808000808000808000808000808000808000
          8080808080FFFFFF008080008080008080008080008080808080008080008080
          0080800080800080800080800080800080800080800080800080800000FF0000
          8000008000008000008080808000808000808000808000808000808000808000
          8080008080008080008080008080008080808080FFFFFF008080008080008080
          8080800080800080800080800080800080800080800080800080800080800080
          800080800000FF00008000008000008000008000008080808000808000808000
          8080008080008080008080008080008080008080008080008080008080808080
          008080008080008080008080808080FFFFFF0080800080800080800080800080
          800080800080800080800080800000FF00008000008000008080808000008000
          0080000080808080008080008080008080008080008080008080008080008080
          008080008080808080008080008080008080008080008080808080FFFFFF0080
          800080800080800080800080800080800080800080800000FF00008000008000
          00808080800080800000FF000080000080000080808080008080008080008080
          008080008080008080008080008080808080008080008080008080808080FFFF
          FF008080008080808080FFFFFF00808000808000808000808000808000808000
          80800000FF0000800000808080800080800080800080800000FF000080000080
          000080808080008080008080008080008080008080008080808080FFFFFF0080
          80008080808080008080808080FFFFFF008080008080808080FFFFFF00808000
          80800080800080800080800080800080800000FF000080008080008080008080
          0080800080800000FF0000800000800000800080800080800080800080800080
          80008080808080FFFFFFFFFFFF808080008080008080008080808080FFFFFF00
          8080008080808080FFFFFF008080008080008080008080008080008080008080
          0080800080800080800080800080800080800080800000FF0000800000FF0080
          8000808000808000808000808000808000808080808080808000808000808000
          8080008080008080808080FFFFFFFFFFFFFFFFFF808080008080008080008080
          0080800080800080800080800080800080800080800080800080800080800080
          8000808000808000808000808000808000808000808000808000808000808000
          8080008080008080008080008080008080008080008080808080808080808080
          0080800080800080800080800080800080800080800080800080800080800080
          8000808000808000808000808000808000808000808000808000808000808000
          8080008080008080008080008080008080008080008080008080008080008080
          008080008080008080008080008080008080}
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAbandonClick
        IsControl = True
      end
      object BParametrer: TToolbarButton97
        Left = 2
        Top = 3
        Width = 56
        Height = 27
        Hint = 'Paramétrer'
        DisplayMode = dmGlyphOnly
        Caption = 'Paramétrer'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Glyph.Data = {
          42010000424D4201000000000000760000002800000011000000110000000100
          040000000000CC000000CE0E0000D80E00001000000000000000000000000000
          BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          77777000000070000000000000077000000070FFF0FFFFFF0F077000000070F0
          F0F0FF0F0F077000000070000000000000077000000077709999999077777000
          0000777090000907777770000000777090709007777770000000777090099700
          77777000000077709099070007777000000077709990770BB077700000007770
          9907770BB07770000000777090777770BB0770000000777007777770B0077000
          00007770777777770BB070000000777777777777000770000000777777777777
          777770000000}
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BParametrerClick
      end
      object LNumCaisse: THLabel
        Left = 72
        Top = 8
        Width = 42
        Height = 16
        Caption = 'Caisse'
        FocusControl = NUMCAISSE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Visible = False
      end
      object Famille: THValComboBox
        Left = 332
        Top = 5
        Width = 61
        Height = 21
        Style = csDropDownList
        Color = clYellow
        ItemHeight = 13
        TabOrder = 0
        Visible = False
        OnChange = OnChangeFamille
        TagDispatch = 0
        DataType = 'GCLIBFAMILLE'
      end
      object NUMCAISSE: THValComboBox
        Left = 120
        Top = 4
        Width = 209
        Height = 24
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 16
        ParentFont = False
        TabOrder = 1
        Visible = False
        OnChange = NUMCAISSEChange
        TagDispatch = 0
        Vide = True
        DataType = 'GCPCAISSE'
      end
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 36
    Width = 727
    Height = 507
    Align = alClient
    TabOrder = 1
    object FNomCaisse: THLabel
      Left = 6
      Top = 7
      Width = 61
      Height = 13
      Caption = 'Caisse n°1'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Ombre = True
      OmbreDecalX = 1
      OmbreDecalY = 1
      OmbreColor = clWhite
    end
    object Ouvertele: THLabel
      Left = 482
      Top = 7
      Width = 237
      Height = 13
      Alignment = taRightJustify
      Caption = 'Dernière ouverture : 29/05/2000 à 08h00'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Ombre = True
      OmbreDecalX = 1
      OmbreDecalY = 1
      OmbreColor = clWhite
    end
    object Bevel1: TBevel
      Left = 1
      Top = 14
      Width = 725
      Height = 14
      Shape = bsBottomLine
    end
    object PG1: THPanel
      Left = 12
      Top = 311
      Width = 348
      Height = 175
      BevelOuter = bvLowered
      FullRepaint = False
      TabOrder = 3
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
      object GRV: TChart
        Left = 1
        Top = 1
        Width = 346
        Height = 173
        AllowPanning = pmNone
        AllowZoom = False
        BackWall.Brush.Color = clWhite
        BackWall.Brush.Style = bsClear
        BackWall.Pen.Visible = False
        Gradient.Direction = gdBottomTop
        Gradient.EndColor = clSilver
        Gradient.Visible = True
        Title.AdjustFrame = False
        Title.Font.Charset = ANSI_CHARSET
        Title.Font.Color = clBlack
        Title.Font.Height = -11
        Title.Font.Name = 'MS Sans Serif'
        Title.Font.Style = []
        Title.Frame.Color = clSilver
        Title.Frame.Visible = True
        Title.Text.Strings = (
          'Statistiques par vendeur')
        AxisVisible = False
        ClipPoints = False
        Frame.Visible = False
        LeftAxis.Automatic = False
        LeftAxis.AutomaticMaximum = False
        LeftAxis.AutomaticMinimum = False
        LeftAxis.LabelsSeparation = 20
        LeftAxis.Visible = False
        Legend.Alignment = laBottom
        Legend.ColorWidth = 10
        Legend.LegendStyle = lsValues
        Legend.ResizeChart = False
        Legend.ShadowSize = 2
        Legend.TextStyle = ltsLeftPercent
        Legend.Visible = False
        View3DOptions.Elevation = 315
        View3DOptions.Orthogonal = False
        View3DOptions.Perspective = 0
        View3DOptions.Rotation = 360
        View3DWalls = False
        Align = alClient
        BevelOuter = bvNone
        TabOrder = 0
        object Series1: TPieSeries
          Marks.ArrowLength = 0
          Marks.Style = smsLabelPercent
          Marks.Visible = True
          SeriesColor = clRed
          Title = 'Vendeurs'
          OtherSlice.Text = 'Other'
          PieValues.DateTime = False
          PieValues.Name = 'Pie'
          PieValues.Multiplier = 1
          PieValues.Order = loNone
          Left = 496
          Top = 5
        end
      end
    end
    object PG2: THPanel
      Left = 369
      Top = 311
      Width = 348
      Height = 175
      BevelOuter = bvLowered
      FullRepaint = False
      TabOrder = 4
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
      object GRF: TChart
        Left = 1
        Top = 1
        Width = 346
        Height = 173
        AllowPanning = pmNone
        AllowZoom = False
        BackWall.Brush.Color = clWhite
        BackWall.Brush.Style = bsClear
        BackWall.Pen.Visible = False
        Gradient.Direction = gdBottomTop
        Gradient.EndColor = clSilver
        Gradient.Visible = True
        Title.AdjustFrame = False
        Title.Font.Charset = ANSI_CHARSET
        Title.Font.Color = clBlack
        Title.Font.Height = -11
        Title.Font.Name = 'MS Sans Serif'
        Title.Font.Style = []
        Title.Frame.Color = clSilver
        Title.Frame.Visible = True
        Title.Text.Strings = (
          'Statistiques par famille')
        AxisVisible = False
        ClipPoints = False
        Frame.Visible = False
        LeftAxis.Automatic = False
        LeftAxis.AutomaticMaximum = False
        LeftAxis.AutomaticMinimum = False
        Legend.ResizeChart = False
        Legend.Visible = False
        View3DOptions.Elevation = 315
        View3DOptions.Orthogonal = False
        View3DOptions.Perspective = 0
        View3DOptions.Rotation = 360
        View3DWalls = False
        Align = alClient
        BevelOuter = bvNone
        TabOrder = 0
        object PieSeries1: TPieSeries
          Marks.ArrowLength = 0
          Marks.Style = smsLabelPercent
          Marks.Visible = True
          SeriesColor = clRed
          Title = 'Familles'
          OtherSlice.Text = 'Other'
          PieValues.DateTime = False
          PieValues.Name = 'Pie'
          PieValues.Multiplier = 1
          PieValues.Order = loNone
          Left = 466
          Top = 5
        end
      end
    end
    object GSS: THGrid
      Left = 12
      Top = 193
      Width = 705
      Height = 113
      ColCount = 9
      DefaultRowHeight = 18
      FixedCols = 3
      Options = [goFixedVertLine, goVertLine, goHorzLine]
      ScrollBars = ssVertical
      TabOrder = 2
      SortedCol = -1
      Titres.Strings = (
        ''
        ''
        ''
        'C.A. T.T.C.'
        'Nb. d'#39'articles'
        'Nb. de tickets'
        'Remise consentie'
        'Panier moyen (mtt)'
        'Panier moyen (qté)')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      OnCellEnter = GSSCellEnter
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = 13224395
      ColWidths = (
        6
        6
        150
        75
        75
        75
        75
        75
        75)
      RowHeights = (
        18
        18
        18
        18
        18)
    end
    object GSV: THGrid
      Left = 12
      Top = 31
      Width = 348
      Height = 156
      ColCount = 4
      DefaultRowHeight = 18
      RowCount = 7
      Options = [goFixedVertLine, goVertLine, goHorzLine, goRowSelect]
      ScrollBars = ssVertical
      TabOrder = 0
      SortedCol = -1
      Titres.Strings = (
        ''
        'Ventes'
        'Nombre'
        'Montant')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = 13224395
      ColWidths = (
        9
        150
        64
        84)
      RowHeights = (
        18
        18
        18
        18
        18
        18
        18)
    end
    object GSR: THGrid
      Left = 369
      Top = 31
      Width = 348
      Height = 156
      ColCount = 6
      DefaultRowHeight = 18
      FixedCols = 2
      RowCount = 7
      Options = [goFixedVertLine, goVertLine, goHorzLine, goRowSelect]
      ScrollBars = ssVertical
      TabOrder = 1
      SortedCol = -1
      Titres.Strings = (
        'Type'
        'Code'
        'Règlements'
        'Mtt encaissé'
        'Mtt en devise'
        'Fonds caisse')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = 13224395
      ColWidths = (
        8
        8
        150
        65
        65
        65)
      RowHeights = (
        18
        18
        18
        18
        18
        18
        18)
    end
  end
  object HMTrad: THSystemMenu
    MaxInsideSize = False
    ResizeDBGrid = True
    LockedCtrls.Strings = (
      'GRAutres'
      'GRCADEAU')
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 435
    Top = 5
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Clôture'
      'n°'
      'du'
      'Caisses'
      'Statistiques par'
      'et par'
      'Total ventes'
      'Retours/Annulations'
      'TOTAL GENERAL'
      'Panier moyen'
      'Total opérations de caisse'
      'Prestations Diverses'
      'Vendeur'
      'Clôtures'
      'Total ventes et prestations'
      'Ventes'
      'Total encaissements'
      'Total autres règlements'
      'Reste dû du jour encaissé'
      'Reste dû non encaissé '
      'Encaissement de Reste dû antérieurs'
      'Reste dû antérieurs'
      'Total à encaisser')
    Left = 405
    Top = 5
  end
end

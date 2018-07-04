object FCumulGestionAssociee: TFCumulGestionAssociee
  Left = 49
  Top = 145
  Width = 613
  Height = 528
  HelpContext = 41120010
  BorderIcons = [biSystemMenu]
  Caption = 'Gestion associ'#233'e'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCloseQuery = FormCloseQuery
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object HPanel1: THPanel
    Left = 0
    Top = 0
    Width = 605
    Height = 460
    Align = alClient
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFond
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    ColorNb = 1
    TextEffect = tenone
    object Label1: TLabel
      Left = 48
      Top = 36
      Width = 116
      Height = 13
      Caption = 'Cumuls non affect'#233's'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Transparent = True
    end
    object Label2: TLabel
      Left = 408
      Top = 36
      Width = 137
      Height = 13
      Caption = 'Cumuls affect'#233's en gain'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Transparent = True
    end
    object Label3: TLabel
      Left = 412
      Top = 268
      Width = 156
      Height = 13
      Caption = 'Cumuls affect'#233's en retenue'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Transparent = True
    end
    object Label4: TLabel
      Left = 8
      Top = 8
      Width = 43
      Height = 13
      Caption = 'Rubrique'
    end
    object BNONAFFECTE_GAIN: THSpeedButton
      Left = 292
      Top = 138
      Width = 23
      Height = 22
      NumGlyphs = 2
      OnClick = BNONAFFECTE_GAINClick
      GlobalIndexImage = 'Z0056_S16G2'
    end
    object BGAIN_NONAFFECTE: THSpeedButton
      Left = 292
      Top = 162
      Width = 23
      Height = 22
      NumGlyphs = 2
      OnClick = BGAIN_NONAFFECTEClick
      GlobalIndexImage = 'Z0077_S16G2'
    end
    object BGAIN_RETENU: THSpeedButton
      Left = 463
      Top = 243
      Width = 23
      Height = 22
      NumGlyphs = 2
      OnClick = BGAIN_RETENUClick
      GlobalIndexImage = 'Z0165_S16G2'
    end
    object BRETENU_GAIN: THSpeedButton
      Left = 487
      Top = 243
      Width = 23
      Height = 22
      NumGlyphs = 2
      OnClick = BRETENU_GAINClick
      GlobalIndexImage = 'Z0104_S16G2'
    end
    object BNONAFFECTE_RETENU: THSpeedButton
      Left = 294
      Top = 366
      Width = 23
      Height = 22
      NumGlyphs = 2
      OnClick = BNONAFFECTE_RETENUClick
      GlobalIndexImage = 'Z0056_S16G2'
    end
    object BRETENU_NONAFFECTE: THSpeedButton
      Left = 294
      Top = 390
      Width = 23
      Height = 22
      NumGlyphs = 2
      OnClick = BRETENU_NONAFFECTEClick
      GlobalIndexImage = 'Z0077_S16G2'
    end
    object LRUBRIQUE: THLabel
      Left = 72
      Top = 8
      Width = 63
      Height = 13
      Caption = 'LRUBRIQUE'
    end
    object GRID_GAIN: THGrid
      Left = 322
      Top = 56
      Width = 276
      Height = 177
      ColCount = 2
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing]
      ScrollBars = ssVertical
      TabOrder = 0
      OnDblClick = GRID_GAINDblClick
      OnDragDrop = DEPOSE_OBJET
      OnDragOver = TEST_DEPOSE_OBJET
      OnMouseDown = GRID_CUMULSMouseDown
      SortedCol = 1
      Titres.Strings = (
        'N'#176
        'Cumul')
      Couleur = False
      MultiSelect = True
      TitleBold = True
      TitleCenter = True
      ColCombo = 0
      SortEnabled = True
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = 13224395
      ColWidths = (
        18
        251)
    end
    object GRID_CUMULS: THGrid
      Left = 8
      Top = 52
      Width = 273
      Height = 405
      ColCount = 2
      DefaultRowHeight = 18
      DragCursor = crMultiDrag
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing]
      ScrollBars = ssVertical
      TabOrder = 1
      OnDblClick = GRID_CUMULSDblClick
      OnDragDrop = DEPOSE_OBJET
      OnDragOver = TEST_DEPOSE_OBJET
      OnMouseDown = GRID_CUMULSMouseDown
      SortedCol = 1
      Titres.Strings = (
        'N'#176
        'Cumul')
      Couleur = False
      MultiSelect = True
      TitleBold = True
      TitleCenter = True
      ColCombo = 0
      SortEnabled = True
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = 13224395
      ColWidths = (
        27
        233)
    end
    object GRID_RETENU: THGrid
      Left = 322
      Top = 284
      Width = 276
      Height = 169
      ColCount = 2
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing]
      ScrollBars = ssVertical
      TabOrder = 2
      OnDblClick = GRID_RETENUDblClick
      OnDragDrop = DEPOSE_OBJET
      OnDragOver = TEST_DEPOSE_OBJET
      OnMouseDown = GRID_CUMULSMouseDown
      SortedCol = 1
      Titres.Strings = (
        'N'#176
        'Cumul')
      Couleur = False
      MultiSelect = True
      TitleBold = True
      TitleCenter = True
      ColCombo = 0
      SortEnabled = True
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = 13224395
      ColWidths = (
        20
        249)
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 460
    Width = 605
    Height = 34
    BackgroundOnToolbars = False
    BackgroundTransparent = True
    Position = dpBottom
    object ToolWindow971: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 30
      ClientWidth = 605
      Caption = 'ToolWindow971'
      ClientAreaHeight = 30
      ClientAreaWidth = 605
      DockableTo = [dpBottom, dpLeft]
      DockPos = 0
      DragHandleStyle = dhNone
      FullSize = True
      Resizable = False
      TabOrder = 0
      object BVALIDER: TToolbarButton97
        Left = 503
        Top = 2
        Width = 28
        Height = 27
        Flat = False
        NumGlyphs = 2
        OnClick = BVALIDERClick
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object BFERMER: TToolbarButton97
        Left = 535
        Top = 2
        Width = 28
        Height = 27
        Flat = False
        ModalResult = 2
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BHELP: TToolbarButton97
        Left = 567
        Top = 2
        Width = 28
        Height = 27
        Flat = False
        OnClick = BHELPClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
    end
  end
  object Msg: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Voulez-vous sauvegarder les modifications ?;Q;YNC;Y;' +
        'C')
    Left = 236
    Top = 8
  end
end

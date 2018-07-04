object FTarifCond: TFTarifCond
  Left = 316
  Top = 180
  HelpContext = 110000064
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Condition d'#39'application tarifaire'
  ClientHeight = 348
  ClientWidth = 521
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PENTETE: THPanel
    Left = 0
    Top = 0
    Width = 521
    Height = 57
    Align = alTop
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object TGF_DATEDEBUT: THLabel
      Left = 5
      Top = 34
      Width = 14
      Height = 13
      Caption = 'De'
      FocusControl = GF_DATEDEBUT
    end
    object TGF_DATEFIN: THLabel
      Left = 170
      Top = 34
      Width = 6
      Height = 13
      Caption = #224
      FocusControl = GF_DATEFIN
    end
    object TGF_DEPOT: THLabel
      Left = 336
      Top = 7
      Width = 29
      Height = 13
      Caption = 'D'#233'p'#244't'
      FocusControl = GF_DEPOT
    end
    object GF_LIBELLE: THCritMaskEdit
      Left = 33
      Top = 7
      Width = 283
      Height = 21
      BorderStyle = bsNone
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 0
      TagDispatch = 0
    end
    object GF_DATEDEBUT: THCritMaskEdit
      Left = 33
      Top = 34
      Width = 121
      Height = 21
      BorderStyle = bsNone
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 1
      TagDispatch = 0
      OpeType = otDate
      ControlerDate = True
    end
    object GF_DATEFIN: THCritMaskEdit
      Left = 195
      Top = 34
      Width = 121
      Height = 21
      BorderStyle = bsNone
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 2
      TagDispatch = 0
      OpeType = otDate
      ControlerDate = True
    end
    object GF_DEPOT: THValComboBox
      Left = 379
      Top = 4
      Width = 122
      Height = 22
      Style = csSimple
      Color = clBtnFace
      ItemHeight = 13
      TabOrder = 3
      TagDispatch = 0
      DataType = 'GCDEPOT'
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 312
    Width = 521
    Height = 36
    AllowDrag = False
    FixAlign = True
    Position = dpBottom
    object Toolbar972: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 521
      Caption = 'Outils Tarif'
      ClientAreaHeight = 32
      ClientAreaWidth = 521
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BValider: TToolbarButton97
        Left = 407
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Enregistrer les tarifs'
        Default = True
        DisplayMode = dmGlyphOnly
        Caption = 'Enregistrer'
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
        Layout = blGlyphTop
        ModalResult = 1
        NumGlyphs = 2
        ParentFont = False
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 439
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        DisplayMode = dmGlyphOnly
        Caption = 'Annuler'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ModalResult = 2
        ParentFont = False
        Spacing = -1
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 470
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
    end
  end
  object PPIED: THPanel
    Left = 0
    Top = 272
    Width = 521
    Height = 40
    Align = alBottom
    FullRepaint = False
    TabOrder = 2
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object FTable: THValComboBox
      Left = 3
      Top = 4
      Width = 79
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = FTableChange
      Items.Strings = (
        'Piece'
        'Ligne'
        'Tiers'
        'Article')
      TagDispatch = 0
      Values.Strings = (
        'PIECE'
        'LIGNE'
        'TIERS'
        'ARTICLE')
    end
    object FChamp: THValComboBox
      Left = 80
      Top = 4
      Width = 161
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 1
      OnChange = FChampChange
      TagDispatch = 0
    end
    object FOpe: THValComboBox
      Left = 240
      Top = 4
      Width = 108
      Height = 21
      Style = csDropDownList
      Ctl3D = False
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 2
      OnChange = FOpeChange
      TagDispatch = 0
      DataType = 'TTCOMPARE'
    end
    object FVal: TEdit
      Left = 349
      Top = 4
      Width = 151
      Height = 21
      Ctl3D = True
      ParentCtl3D = False
      TabOrder = 3
      OnChange = FValChange
    end
  end
  object PGRILLE: THPanel
    Left = 0
    Top = 57
    Width = 521
    Height = 215
    Align = alClient
    FullRepaint = False
    TabOrder = 3
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object G_COND: THGrid
      Left = 1
      Top = 1
      Width = 519
      Height = 213
      Align = alClient
      ColCount = 4
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 25
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
      TabOrder = 0
      OnKeyDown = G_CONDKeyDown
      SortedCol = -1
      Titres.Strings = (
        'Table'
        'Champ'
        'Condition'
        'Valeur')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      OnRowEnter = G_CONDRowEnter
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = clSilver
      ColWidths = (
        73
        162
        104
        154)
    end
    object FComboTIE: THValComboBox
      Left = 207
      Top = 52
      Width = 56
      Height = 21
      Style = csDropDownList
      Color = clYellow
      ItemHeight = 13
      TabOrder = 1
      Visible = False
      TagDispatch = 0
    end
    object FComboART: THValComboBox
      Left = 207
      Top = 77
      Width = 56
      Height = 21
      Style = csDropDownList
      Color = clYellow
      ItemHeight = 13
      TabOrder = 2
      Visible = False
      TagDispatch = 0
    end
    object FComboLIG: THValComboBox
      Left = 207
      Top = 103
      Width = 56
      Height = 21
      Style = csDropDownList
      Color = clYellow
      ItemHeight = 13
      TabOrder = 3
      Visible = False
      TagDispatch = 0
    end
    object FComboPIE: THValComboBox
      Left = 207
      Top = 128
      Width = 56
      Height = 21
      Style = csDropDownList
      Color = clYellow
      ItemHeight = 13
      TabOrder = 4
      Visible = False
      TagDispatch = 0
    end
    object GF_CONDAPPLIC: THRichEditOLE
      Left = 272
      Top = 56
      Width = 185
      Height = 89
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      TabOrder = 5
      Visible = False
      WordWrap = False
      Margins.Top = 0
      Margins.Bottom = 0
      Margins.Left = 0
      Margins.Right = 0
      ContainerName = 'Document'
      ObjectMenuPrefix = '&Object'
      LinesRTF.Strings = (
        
          '{\rtf1\ansi\ansicpg1252\deff0\deflang1036{\fonttbl{\f0\fswiss Ar' +
          'ial;}}'
        
          '{\*\generator Msftedit 5.41.15.1507;}\viewkind4\uc1\pard\f0\fs16' +
          ' GF_CONDAPPLIC'
        '\par }')
    end
  end
end

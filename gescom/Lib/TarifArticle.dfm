object FTarifArticle: TFTarifArticle
  Left = 251
  Top = 150
  Width = 613
  Height = 441
  HelpContext = 110000058
  BorderIcons = [biSystemMenu]
  Caption = 'Tarifs par article'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  ShowHint = True
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PDIMENSION: THPanel
    Left = 0
    Top = 27
    Width = 597
    Height = 43
    Align = alTop
    BevelOuter = bvNone
    FullRepaint = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object TGF_GRILLEDIM1: THLabel
      Left = 12
      Top = 1
      Width = 47
      Height = 13
      Caption = 'GrilledIM1'
    end
    object TGF_GRILLEDIM2: THLabel
      Left = 131
      Top = 1
      Width = 47
      Height = 13
      Caption = 'GrilledIM2'
    end
    object TGF_GRILLEDIM3: THLabel
      Left = 250
      Top = 1
      Width = 47
      Height = 13
      Caption = 'GrilledIM3'
    end
    object TGF_GRILLEDIM4: THLabel
      Left = 368
      Top = 1
      Width = 47
      Height = 13
      Caption = 'GrilledIM4'
    end
    object TGF_GRILLEDIM5: THLabel
      Left = 487
      Top = 1
      Width = 47
      Height = 13
      Caption = 'GrilledIM5'
    end
    object GF_CODEDIM1: THCritMaskEdit
      Left = 12
      Top = 16
      Width = 113
      Height = 21
      ReadOnly = True
      TabOrder = 0
      TagDispatch = 0
    end
    object GF_CODEDIM2: THCritMaskEdit
      Left = 131
      Top = 16
      Width = 113
      Height = 21
      ReadOnly = True
      TabOrder = 1
      TagDispatch = 0
    end
    object GF_CODEDIM3: THCritMaskEdit
      Left = 250
      Top = 16
      Width = 113
      Height = 21
      ReadOnly = True
      TabOrder = 2
      TagDispatch = 0
    end
    object GF_CODEDIM4: THCritMaskEdit
      Left = 368
      Top = 16
      Width = 113
      Height = 21
      ReadOnly = True
      TabOrder = 3
      TagDispatch = 0
    end
    object GF_CODEDIM5: THCritMaskEdit
      Left = 487
      Top = 16
      Width = 113
      Height = 21
      ReadOnly = True
      TabOrder = 4
      TagDispatch = 0
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 367
    Width = 597
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object Toolbar972: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 597
      Caption = 'Outils Tarif'
      ClientAreaHeight = 32
      ClientAreaWidth = 597
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BValider: TToolbarButton97
        Left = 498
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Enregistrer les tarifs'
        Caption = 'Enregistrer'
        Default = True
        DisplayMode = dmGlyphOnly
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
        Left = 530
        Top = 3
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
        Spacing = -1
        OnClick = BAbandonClick
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 561
        Top = 3
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
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BChercher: TToolbarButton97
        Tag = 1
        Left = 53
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Rechercher '
        Caption = 'Rechercher'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        Spacing = -1
        OnClick = BChercherClick
        GlobalIndexImage = 'Z0077_S16G1'
        IsControl = True
      end
      object BInfos: TToolbarButton97
        Tag = 1
        Left = 10
        Top = 3
        Width = 40
        Height = 27
        Hint = 'Compl'#233'ments d'#39'informations'
        Caption = 'Compl'#233'ments'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = POPZ
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
      object BSaisieRapide: TToolbarButton97
        Tag = 1
        Left = 84
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Saisie Rapide'
        Caption = 'Saisie Rapide'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        Spacing = -1
        OnClick = BSaisieRapideClick
        GlobalIndexImage = 'Z0298_S16G1'
        IsControl = True
      end
      object BVoirCond: TToolbarButton97
        Tag = 1
        Left = 136
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Voir conditions tarifaires'
        Caption = 'Voir condition'
        AllowAllUp = True
        GroupIndex = 2
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Spacing = -1
        OnClick = BVoirCondClick
        GlobalIndexImage = 'Z0206_S16G1'
        IsControl = True
      end
      object BCondAplli: TToolbarButton97
        Left = 167
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Modifier conditions'
        Caption = 'Modifier'
        DisplayMode = dmGlyphOnly
        Flat = False
        NumGlyphs = 2
        OnClick = BCondAplliClick
        GlobalIndexImage = 'Z0105_S16G2'
      end
      object BCopierCond: TToolbarButton97
        Left = 198
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Copier conditions'
        Caption = 'Copier'
        DisplayMode = dmGlyphOnly
        Flat = False
        OnClick = BCopierCondClick
        GlobalIndexImage = 'Z0342_S16G1'
      end
      object BCollerCond: TToolbarButton97
        Left = 229
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Coller conditions'
        Caption = 'Coller'
        DisplayMode = dmGlyphOnly
        Flat = False
        OnClick = BCollerCondClick
        GlobalIndexImage = 'Z0175_S16G1'
      end
    end
  end
  object PPIED: THPanel
    Left = 0
    Top = 316
    Width = 597
    Height = 51
    Align = alBottom
    FullRepaint = False
    TabOrder = 7
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object TGF_REMISE: THLabel
      Left = 11
      Top = 29
      Width = 84
      Height = 13
      Caption = 'Remise r'#233'sultante'
      FocusControl = GF_REMISE
    end
    object TGF_CASCADEREMISE: THLabel
      Left = 11
      Top = 6
      Width = 60
      Height = 13
      Caption = '&Mode remise'
      FocusControl = GF_CASCADEREMISE
    end
    object ISigneEuro: TImage
      Left = 364
      Top = 6
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
    object TGF_PRIXCON: THLabel
      Left = 306
      Top = 6
      Width = 54
      Height = 13
      Caption = 'Prix unitaire'
    end
    object ISigneFranc: TImage
      Left = 364
      Top = 6
      Width = 16
      Height = 16
      AutoSize = True
      Picture.Data = {
        07544269746D6170F6000000424DF60000000000000076000000280000001000
        0000100000000100040000000000800000000000000000000000100000001000
        0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
        C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
        FFFFF44FFFF44FF444FFF44FFFF44F44F44FF44FFFF44FFF44FFF44FFFF44FF4
        4FFFF44444F44F44F44FF44FFFF444F444FFF44FFFFFFFFFFFFFF44FFFFFFFFF
        FFFFF444444FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
        FFFF}
      Stretch = True
      Transparent = True
    end
    object TGA_PRIXPOURQTE: THLabel
      Left = 250
      Top = 29
      Width = 132
      Height = 13
      Caption = 'Prix indiqu'#233' pour une qt'#233' de'
      Visible = False
    end
    object GF_REMISE: THCritMaskEdit
      Left = 105
      Top = 26
      Width = 126
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 0
      OnChange = GF_REMISEChange
      TagDispatch = 0
      OpeType = otReel
    end
    object GF_CASCADEREMISE: THValComboBox
      Left = 105
      Top = 3
      Width = 126
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 1
      OnChange = GF_CASCADEREMISEChange
      TagDispatch = 0
      DataType = 'GCCASCADEREMISE'
    end
    object CONDAPPLIC: THRichEditOLE
      Left = 532
      Top = 28
      Width = 73
      Height = 21
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
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
        '{\colortbl ;\red255\green0\blue0;}'
        
          '{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\pard\cf1\f0\' +
          'fs16 CONDAPPLIC'
        '\par }')
    end
    object GF_PRIXCON: THNumEdit
      Left = 388
      Top = 4
      Width = 85
      Height = 21
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 3
      Decimals = 2
      Digits = 12
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      UseRounding = True
      Validate = False
    end
    object GA_PRIXPOURQTE: THNumEdit
      Left = 388
      Top = 26
      Width = 85
      Height = 21
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 4
      Visible = False
      Decimals = 0
      Digits = 12
      Masks.PositiveMask = '#,##0'
      Debit = False
      UseRounding = True
      Validate = False
    end
  end
  object PDEVISE: THPanel
    Left = 0
    Top = 70
    Width = 597
    Height = 51
    Align = alTop
    BevelOuter = bvNone
    FullRepaint = False
    TabOrder = 2
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object TGF_DEVISE: THLabel
      Left = 12
      Top = 5
      Width = 59
      Height = 13
      Caption = 'Code &devise'
      FocusControl = GF_DEVISE
    end
    object GF_DEVISE: THValComboBox
      Left = 82
      Top = 0
      Width = 145
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = GF_DEVISEChange
      TagDispatch = 0
      DataType = 'TTDEVISE'
    end
    object PTITRE: THPanel
      Left = 0
      Top = 26
      Width = 605
      Height = 25
      Caption = 'PTITRE'
      FullRepaint = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 1
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
      object TTYPETARIF: THLabel
        Left = 8
        Top = 6
        Width = 66
        Height = 13
        Caption = 'TTYPETARIF'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
    end
    object CBQUANTITAIF: TCheckBox
      Left = 266
      Top = 5
      Width = 97
      Height = 17
      Caption = 'Tarif &quantitatif'
      TabOrder = 2
      OnClick = CBQUANTITAIFClick
    end
    object CBCATTIERS: TCheckBox
      Left = 384
      Top = 5
      Width = 145
      Height = 17
      Caption = '&Tarif cat'#233'gorie client'
      TabOrder = 3
      OnClick = CBCATTIERSClick
    end
  end
  object PARTICLE: THPanel
    Left = 0
    Top = 0
    Width = 597
    Height = 27
    Align = alTop
    BevelOuter = bvNone
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object TGF_ARTICLE: THLabel
      Left = 12
      Top = 8
      Width = 56
      Height = 13
      Caption = 'Code &article'
      FocusControl = GF_CODEARTICLE
    end
    object TGF_LIBELLE: TLabel
      Left = 242
      Top = 8
      Width = 3
      Height = 13
    end
    object GF_CODEARTICLE: THCritMaskEdit
      Left = 82
      Top = 4
      Width = 145
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 0
      OnExit = GF_CODEARTICLEExit
      TagDispatch = 0
      DataType = 'GCARTICLEGENERIQUE'
      ElipsisButton = True
      OnElipsisClick = GF_CODEARTICLEElipsisClick
    end
  end
  object PQUANTITATIF: THPanel
    Left = 0
    Top = 121
    Width = 597
    Height = 195
    Align = alClient
    BevelOuter = bvLowered
    FullRepaint = False
    TabOrder = 4
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object G_Qte: THGrid
      Tag = 1
      Left = 1
      Top = 1
      Width = 595
      Height = 193
      Align = alClient
      ColCount = 10
      DefaultColWidth = 50
      DefaultRowHeight = 18
      RowCount = 3
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goTabs]
      TabOrder = 0
      OnEnter = G_QteEnter
      SortedCol = 0
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      OnRowEnter = G_QteRowEnter
      OnRowExit = G_QteRowExit
      OnCellEnter = G_QteCellEnter
      OnCellExit = G_QteCellExit
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
      ElipsisButton = True
      OnElipsisClick = G_QteElipsisClick
      ColWidths = (
        50
        50
        133
        75
        88
        67
        80
        50
        50
        50)
    end
  end
  object PQTECAT: THPanel
    Left = 0
    Top = 121
    Width = 597
    Height = 195
    Align = alClient
    FullRepaint = False
    TabOrder = 5
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object G_QCa: THGrid
      Tag = 1
      Left = 1
      Top = 1
      Width = 595
      Height = 193
      Align = alClient
      ColCount = 10
      DefaultColWidth = 50
      DefaultRowHeight = 18
      RowCount = 3
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goTabs]
      TabOrder = 0
      OnEnter = G_QCaEnter
      SortedCol = 0
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      OnRowEnter = G_QCaRowEnter
      OnRowExit = G_QCaRowExit
      OnCellEnter = G_QCaCellEnter
      OnCellExit = G_QCaCellExit
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
      ElipsisButton = True
      OnElipsisClick = G_QCaElipsisClick
      ColWidths = (
        50
        50
        133
        75
        88
        67
        80
        50
        50
        50)
    end
  end
  object PCATEGORIE: THPanel
    Left = 0
    Top = 121
    Width = 597
    Height = 195
    Align = alClient
    FullRepaint = False
    TabOrder = 6
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object G_Cat: THGrid
      Tag = 1
      Left = 1
      Top = 1
      Width = 595
      Height = 193
      Align = alClient
      ColCount = 10
      DefaultColWidth = 50
      DefaultRowHeight = 18
      RowCount = 3
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goTabs]
      TabOrder = 0
      OnEnter = G_CatEnter
      SortedCol = 0
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      OnRowEnter = G_CatRowEnter
      OnRowExit = G_CatRowExit
      OnCellEnter = G_CatCellEnter
      OnCellExit = G_CatCellExit
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
      ElipsisButton = True
      OnElipsisClick = G_CatElipsisClick
      ColWidths = (
        50
        50
        133
        75
        88
        67
        80
        50
        50
        50)
      RowHeights = (
        18
        18
        18)
    end
  end
  object PART: THPanel
    Left = 0
    Top = 121
    Width = 597
    Height = 195
    Align = alClient
    FullRepaint = False
    TabOrder = 3
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object G_ART: THGrid
      Tag = 1
      Left = 1
      Top = 1
      Width = 595
      Height = 193
      Align = alClient
      ColCount = 10
      DefaultColWidth = 50
      DefaultRowHeight = 18
      RowCount = 3
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goTabs]
      TabOrder = 0
      OnEnter = G_ARTEnter
      SortedCol = 0
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      OnRowEnter = G_ARTRowEnter
      OnRowExit = G_ARTRowExit
      OnCellEnter = G_ARTCellEnter
      OnCellExit = G_ARTCellExit
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
      ElipsisButton = True
      OnElipsisClick = G_ARTElipsisClick
      ColWidths = (
        50
        50
        133
        75
        88
        67
        80
        50
        50
        50)
      RowHeights = (
        18
        18
        18)
    end
  end
  object TCONDTARF: TToolWindow97
    Left = 64
    Top = 112
    ClientHeight = 175
    ClientWidth = 516
    Caption = 'Conditions tarifaires'
    ClientAreaHeight = 175
    ClientAreaWidth = 516
    Resizable = False
    TabOrder = 9
    Visible = False
    OnClose = TCONDTARFClose
    object G_COND: THGrid
      Left = 0
      Top = 0
      Width = 516
      Height = 175
      Align = alClient
      ColCount = 4
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 25
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
      TabOrder = 0
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
      Top = 21
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
      Top = 46
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
      Top = 72
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
      Top = 97
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
        
          '{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\pard\f0\fs16' +
          ' GF_CONDAPPLIC'
        '\par }')
    end
    object FTable: THValComboBox
      Left = 207
      Top = 121
      Width = 56
      Height = 21
      Style = csDropDownList
      Color = clYellow
      ItemHeight = 13
      TabOrder = 6
      Visible = False
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
    object FOpe: THValComboBox
      Left = 207
      Top = 146
      Width = 56
      Height = 21
      Style = csDropDownList
      Color = clYellow
      Ctl3D = False
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 7
      Visible = False
      TagDispatch = 0
      DataType = 'TTCOMPARE'
    end
  end
  object HMTrad: THSystemMenu
    ResizeDBGrid = True
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 517
    Top = 204
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
      '2;?caption?;Les bornes de quantit'#233's sont incoh'#233'rentes.;W;O;O;O;'
      
        '3;?caption?;Les bornes de dates de validit'#233' sont incoh'#233'rentes.;W' +
        ';O;O;O;'
      '4;?caption?;Ce code d'#233'p'#244't n'#39'existe pas.;W;O;O;O;'
      '5;?caption?;Cette cat'#233'gorie de tiers n'#39'existe pas.;W;O;O;O;'
      '6;?caption?;Confirmez-vous l'#39'abandon de la saisie ?;Q;YN;Y;N;')
    Left = 466
    Top = 204
  end
  object FindLigne: TFindDialog
    OnFind = FindLigneFind
    Left = 568
    Top = 204
  end
  object POPZ: TPopupMenu
    Left = 415
    Top = 204
    object InfArticle: TMenuItem
      Caption = 'Article'
      OnClick = InfArticleClick
    end
  end
  object HMess: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'ATTENTION : Tarif non enregistr'#233' !'
      
        'ATTENTION : Ce tarif, en cours de traitement par un autre utilis' +
        'ateur, n'#39'a pas '#233't'#233' enregistr'#233'e !')
    Left = 364
    Top = 204
  end
  object HTitre: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Quantitatif article'
      'Quantitatif par cat'#233'gorie de client'
      'Cat'#233'gorie client'
      'Article'
      'Tarif global'
      'Tarif HT'
      'Tarif TTC'
      'Nombre de d'#233'cimales du prix')
    Left = 312
    Top = 204
  end
end

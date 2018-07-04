object FCorrespGPAO: TFCorrespGPAO
  Left = 24
  Top = 136
  Width = 696
  Height = 480
  Caption = 'Saisie des tables de correspondances'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnDestroy = FormDestroy
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PENTETE: TPanel
    Left = 0
    Top = 0
    Width = 688
    Height = 129
    Align = alTop
    TabOrder = 0
    object LGRO_TYPEREPRISE: TLabel
      Left = 24
      Top = 8
      Width = 73
      Height = 13
      Caption = 'Type de reprise'
      FocusControl = GRO_TYPEREPRISE
    end
    object LibNomTable: TLabel
      Left = 427
      Top = 12
      Width = 3
      Height = 13
    end
    object LGRO_NOMCHAMP: TLabel
      Left = 24
      Top = 32
      Width = 72
      Height = 13
      Caption = 'Nom du champ'
      FocusControl = GRO_NOMCHAMP
    end
    object LVALDEFAUTGPAO: TLabel
      Left = 24
      Top = 56
      Width = 114
      Height = 13
      Caption = 'Valeur par d'#233'faut GPAO'
      FocusControl = VALDEFAUTGPAO
    end
    object LVALDEFAUTPGI: TLabel
      Left = 24
      Top = 80
      Width = 107
      Height = 13
      Caption = 'Valeurs par d'#233'faut PGI'
      FocusControl = VALDEFAUTPGI1
    end
    object GRO_TYPEREPRISE: THValComboBox
      Left = 176
      Top = 4
      Width = 248
      Height = 21
      Style = csDropDownList
      Enabled = False
      ItemHeight = 13
      TabOrder = 0
      TagDispatch = 0
      DataType = 'GCTYPEREPRISE'
    end
    object GRO_NOMCHAMP: THCritMaskEdit
      Left = 176
      Top = 28
      Width = 248
      Height = 21
      CharCase = ecUpperCase
      ReadOnly = True
      TabOrder = 1
      TagDispatch = 0
    end
    object VALDEFAUTGPAO: THCritMaskEdit
      Left = 176
      Top = 52
      Width = 248
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 35
      TabOrder = 2
      OnChange = VALDEFAUTChange
      TagDispatch = 0
    end
    object VALDEFAUTPGI1: THCritMaskEdit
      Left = 176
      Top = 76
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 35
      TabOrder = 4
      Visible = False
      OnChange = VALDEFAUTChange
      TagDispatch = 0
      ElipsisButton = True
    end
    object VALDEFAUTPGI5: THCritMaskEdit
      Left = 176
      Top = 100
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 35
      TabOrder = 8
      Visible = False
      OnChange = VALDEFAUTChange
      TagDispatch = 0
      ElipsisButton = True
    end
    object VALDEFAUTPGI2: THCritMaskEdit
      Left = 303
      Top = 76
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 35
      TabOrder = 5
      Visible = False
      OnChange = VALDEFAUTChange
      TagDispatch = 0
      ElipsisButton = True
    end
    object VALDEFAUTPGI3: THCritMaskEdit
      Left = 431
      Top = 76
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 35
      TabOrder = 6
      Visible = False
      OnChange = VALDEFAUTChange
      TagDispatch = 0
      ElipsisButton = True
    end
    object VALDEFAUTPGI4: THCritMaskEdit
      Left = 558
      Top = 76
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 35
      TabOrder = 7
      Visible = False
      OnChange = VALDEFAUTChange
      TagDispatch = 0
      ElipsisButton = True
    end
    object VALDEFAUTPGI6: THCritMaskEdit
      Left = 303
      Top = 100
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 35
      TabOrder = 9
      Visible = False
      OnChange = VALDEFAUTChange
      TagDispatch = 0
      ElipsisButton = True
    end
    object VALDEFAUTPGI7: THCritMaskEdit
      Left = 431
      Top = 100
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 35
      TabOrder = 10
      Visible = False
      OnChange = VALDEFAUTChange
      TagDispatch = 0
      ElipsisButton = True
    end
    object VALDEFAUTPGI8: THCritMaskEdit
      Left = 558
      Top = 100
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 35
      TabOrder = 11
      Visible = False
      OnChange = VALDEFAUTChange
      TagDispatch = 0
      ElipsisButton = True
    end
    object LIBDEFAUTGPAO: THCritMaskEdit
      Left = 431
      Top = 52
      Width = 248
      Height = 21
      MaxLength = 35
      TabOrder = 3
      OnChange = VALDEFAUTChange
      TagDispatch = 0
    end
  end
  object Panel4: TPanel
    Left = 0
    Top = 405
    Width = 688
    Height = 41
    Align = alBottom
    TabOrder = 1
    object BValider: TToolbarButton97
      Left = 627
      Top = 7
      Width = 28
      Height = 27
      Hint = 'Enregistrer le param'#233'trage'
      Default = True
      DisplayMode = dmGlyphOnly
      Caption = 'Enregistrer'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
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
      NumGlyphs = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      OnClick = BValiderClick
      IsControl = True
    end
    object BFerme: TToolbarButton97
      Left = 657
      Top = 7
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      ModalResult = 2
      ParentShowHint = False
      ShowHint = True
      GlobalIndexImage = 'Z1770_S16G1'
    end
    object BChercher: TToolbarButton97
      Tag = 1
      Left = 255
      Top = 7
      Width = 28
      Height = 27
      Hint = 'Rechercher dans la liste des valeurs'
      DisplayMode = dmGlyphOnly
      Caption = 'Rechercher'
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
      OnClick = BChercherClick
      GlobalIndexImage = 'Z0077_S16G1'
      IsControl = True
    end
    object bNewligne: TToolbarButton97
      Left = 2
      Top = 7
      Width = 28
      Height = 27
      Hint = 'Ins'#233'rer une ligne'
      DisplayMode = dmGlyphOnly
      Caption = 'Ins'#233'rer'
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
      OnClick = bNewligneClick
      GlobalIndexImage = 'Z0074_S16G1'
    end
    object bDelLigne: TToolbarButton97
      Left = 32
      Top = 7
      Width = 28
      Height = 27
      Hint = 'Supprimer la ligne en cours'
      DisplayMode = dmGlyphOnly
      Caption = 'Effacer'
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
      OnClick = bDelLigneClick
      GlobalIndexImage = 'Z0072_S16G1'
    end
    object BtestGPAOenPGI: TToolbarButton97
      Left = 315
      Top = 7
      Width = 28
      Height = 27
      Hint = 'Test correspondance GPAO vers PGI'
      DisplayMode = dmGlyphOnly
      Caption = 'Ins'#233'rer'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Layout = blGlyphTop
      NumGlyphs = 2
      Opaque = False
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      OnClick = BtestGPAOenPGIClick
      GlobalIndexImage = 'Z0822_S16G2'
    end
    object BtestPGIenGPAO: TToolbarButton97
      Left = 345
      Top = 7
      Width = 28
      Height = 27
      Hint = 'Test correspondance PGI vers GPAO'
      DisplayMode = dmGlyphOnly
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Layout = blGlyphTop
      NumGlyphs = 2
      Opaque = False
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      OnClick = BtestPGIenGPAOClick
      GlobalIndexImage = 'Z0190_S16G2'
    end
    object BGenere: TToolbarButton97
      Left = 405
      Top = 7
      Width = 27
      Height = 27
      Hint = 'G'#233'n'#233'rer les valeurs PGI sans correspondance'
      DisplayMode = dmGlyphOnly
      Caption = 'Ins'#233'rer'
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
      OnClick = BGenereClick
      GlobalIndexImage = 'Z1028_S16G1'
    end
    object BVoirTob: TToolbarButton97
      Tag = 1
      Left = 285
      Top = 7
      Width = 28
      Height = 27
      Hint = 'Voir la TOB'
      DisplayMode = dmGlyphOnly
      Caption = 'Rechercher'
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
      OnClick = BVoirTobClick
      GlobalIndexImage = 'Z0061_S16G1'
      IsControl = True
    end
    object BImprime: TToolbarButton97
      Left = 225
      Top = 7
      Width = 28
      Height = 27
      Hint = 'Imprimer la liste des valeurs'
      ParentShowHint = False
      ShowHint = True
      OnClick = BImprimeClick
      GlobalIndexImage = 'Z0369_S16G1'
    end
    object FAutoSave: TCheckBox
      Left = 117
      Top = 8
      Width = 44
      Height = 17
      Caption = 'Auto'
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      TabOrder = 0
      Visible = False
    end
  end
  object GS: THGrid
    Left = 0
    Top = 129
    Width = 688
    Height = 276
    Align = alClient
    ColCount = 10
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goTabs]
    TabOrder = 2
    OnDblClick = GSElipsisClick
    OnEnter = GSEnter
    SortedCol = -1
    Titres.Strings = (
      ''
      'Valeur du champ repris'
      'Libell'#233' du champ repris'
      'Valeur correspondante')
    Couleur = False
    MultiSelect = False
    TitleBold = True
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
    OnElipsisClick = GSElipsisClick
    DBIndicator = True
    ColWidths = (
      10
      64
      64
      64
      64
      64
      64
      64
      64
      64)
  end
  object TWTest: TToolWindow97
    Left = 80
    Top = 160
    ClientHeight = 133
    ClientWidth = 577
    Caption = 'Test d'#39'utilisation de la table de correspondance :'
    ClientAreaHeight = 133
    ClientAreaWidth = 577
    DockableTo = []
    DockPos = 20
    Resizable = False
    TabOrder = 3
    Visible = False
    OnClose = TWTestClose
    object LVALTESTGPAO: TLabel
      Left = 8
      Top = 16
      Width = 63
      Height = 13
      Caption = 'Valeur GPAO'
      FocusControl = VALTESTGPAO
    end
    object LVALTESTPGI1: TLabel
      Left = 8
      Top = 40
      Width = 56
      Height = 13
      Caption = 'Valeurs PGI'
      FocusControl = VALTESTPGI1
    end
    object TVALTESTGPAO: THLabel
      Left = 326
      Top = 16
      Width = 85
      Height = 13
      Caption = 'TVALTESTGPAO'
    end
    object HPanel5: THPanel
      Left = 0
      Top = 95
      Width = 577
      Height = 38
      Align = alBottom
      BevelInner = bvRaised
      BevelOuter = bvSpace
      FullRepaint = False
      TabOrder = 9
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
      object BValideTest: TToolbarButton97
        Left = 504
        Top = 3
        Width = 32
        Height = 32
        DisplayMode = dmGlyphOnly
        Caption = 'Valider'
        OnClick = BValideTestClick
        GlobalIndexImage = 'Z0127_S16G1'
      end
      object BFermeTest: TToolbarButton97
        Left = 540
        Top = 3
        Width = 32
        Height = 32
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        OnClick = BFermeTestClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
    end
    object VALTESTGPAO: THCritMaskEdit
      Left = 78
      Top = 12
      Width = 244
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 35
      TabOrder = 0
      TagDispatch = 0
    end
    object VALTESTPGI1: THCritMaskEdit
      Left = 78
      Top = 36
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 35
      TabOrder = 1
      Visible = False
      TagDispatch = 0
    end
    object VALTESTPGI2: THCritMaskEdit
      Left = 202
      Top = 36
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 35
      TabOrder = 2
      Visible = False
      TagDispatch = 0
    end
    object VALTESTPGI3: THCritMaskEdit
      Left = 326
      Top = 36
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 35
      TabOrder = 3
      Visible = False
      TagDispatch = 0
    end
    object VALTESTPGI4: THCritMaskEdit
      Left = 450
      Top = 36
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 35
      TabOrder = 4
      Visible = False
      TagDispatch = 0
    end
    object VALTESTPGI8: THCritMaskEdit
      Left = 450
      Top = 60
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 35
      TabOrder = 8
      Visible = False
      TagDispatch = 0
    end
    object VALTESTPGI7: THCritMaskEdit
      Left = 326
      Top = 60
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 35
      TabOrder = 7
      Visible = False
      TagDispatch = 0
    end
    object VALTESTPGI6: THCritMaskEdit
      Left = 202
      Top = 60
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 35
      TabOrder = 6
      Visible = False
      TagDispatch = 0
    end
    object VALTESTPGI5: THCritMaskEdit
      Left = 78
      Top = 60
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 35
      TabOrder = 5
      Visible = False
      TagDispatch = 0
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 546
    Top = 360
  end
  object FindLigne: TFindDialog
    OnFind = FindLigneFind
    Left = 612
    Top = 359
  end
end

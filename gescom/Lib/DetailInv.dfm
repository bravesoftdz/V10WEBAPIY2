object FDetailInv: TFDetailInv
  Left = 240
  Top = 106
  Width = 548
  Height = 320
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Inventaire : d'#233'tail des lots'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PENTETE: THPanel
    Left = 0
    Top = 0
    Width = 540
    Height = 57
    Align = alTop
    BevelOuter = bvNone
    FullRepaint = False
    TabOrder = 2
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object TCodeListe: THLabel
      Left = 4
      Top = 4
      Width = 54
      Height = 13
      Caption = 'TCodeListe'
    end
    object TLibelle: THLabel
      Left = 112
      Top = 4
      Width = 37
      Height = 13
      Caption = 'TLibelle'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TDepot: THLabel
      Left = 420
      Top = 4
      Width = 36
      Height = 13
      Caption = 'TDepot'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HLabel1: THLabel
      Left = 340
      Top = 4
      Width = 35
      Height = 13
      Caption = 'D'#233'p'#244't :'
    end
    object TLibelleArticle: THLabel
      Left = 4
      Top = 40
      Width = 66
      Height = 13
      Caption = 'TLibelleArticle'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HLabel2: THLabel
      Left = 340
      Top = 40
      Width = 75
      Height = 13
      Caption = 'Unit'#233' de stock :'
    end
    object TUStock: THLabel
      Left = 420
      Top = 40
      Width = 43
      Height = 13
      Caption = 'TUStock'
    end
    object TGIL_EMPLACEMENT: THLabel
      Left = 340
      Top = 20
      Width = 70
      Height = 13
      Caption = 'Emplacement :'
    end
    object TEmplacement: THLabel
      Left = 420
      Top = 20
      Width = 71
      Height = 13
      Caption = 'TEmplacement'
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 250
    Width = 540
    Height = 36
    Position = dpBottom
    object ToolWindow971: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 531
      Caption = 'Outils inventaire'
      ClientAreaHeight = 32
      ClientAreaWidth = 531
      DockPos = 0
      FullSize = True
      TabOrder = 0
      DesignSize = (
        531
        32)
      object BChercher: TToolbarButton97
        Tag = 1
        Left = 99
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Rechercher '
        DisplayMode = dmGlyphOnly
        Caption = 'Rechercher'
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
        OnClick = BChercherClick
        GlobalIndexImage = 'Z0077_S16G1'
        IsControl = True
      end
      object BValider: TToolbarButton97
        Left = 434
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Valider'
        Anchors = [akTop, akRight]
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
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 466
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Annuler'
        Anchors = [akTop, akRight]
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
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
      object BAide: TToolbarButton97
        Tag = 1
        Left = 498
        Top = 3
        Width = 28
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
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BInfos: TToolbarButton97
        Tag = 1
        Left = 54
        Top = 3
        Width = 40
        Height = 27
        Hint = 'Fonctions suppl'#233'mentaires'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = POPU
        Caption = 'Compl'#233'ments'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        GlobalIndexImage = 'Z0105_S16G2'
        IsControl = True
      end
      object BZoom: TToolbarButton97
        Tag = 1
        Left = 9
        Top = 3
        Width = 40
        Height = 27
        Hint = 'Compl'#233'ments d'#39'informations'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = POPZ
        Caption = 'Zoom'
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
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
      object BDelLigne: TToolbarButton97
        Tag = 1
        Left = 167
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Supprimer le lot'
        DisplayMode = dmGlyphOnly
        Caption = 'Suppression'
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
        OnClick = BDelLigneClick
        GlobalIndexImage = 'Z0367_S16G1'
        IsControl = True
      end
      object BAddLigne: TToolbarButton97
        Tag = 1
        Left = 135
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Ins'#233'rer un lot'
        DisplayMode = dmGlyphOnly
        Caption = 'Ajout'
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
        OnClick = BAddLigneClick
        GlobalIndexImage = 'Z0661_S16G1'
        IsControl = True
      end
    end
  end
  object G_Dtl: THGrid
    Tag = 1
    Left = 0
    Top = 101
    Width = 540
    Height = 149
    Align = alClient
    ColCount = 3
    DefaultRowHeight = 18
    RowCount = 3
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected, goEditing, goTabs, goAlwaysShowEditor, goThumbTracking]
    TabOrder = 0
    OnKeyDown = G_DtlKeyDown
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    OnRowEnter = G_DtlRowEnter
    OnCellEnter = G_DtlCellEnter
    OnCellExit = G_DtlCellExit
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = clSilver
    ColWidths = (
      64
      64
      64)
  end
  object PDim: THPanel
    Left = 0
    Top = 57
    Width = 540
    Height = 44
    Align = alTop
    BevelOuter = bvNone
    FullRepaint = False
    TabOrder = 3
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object TGF_GRILLEDIM1: THLabel
      Left = 4
      Top = 1
      Width = 47
      Height = 13
      Caption = 'GrilledIM1'
    end
    object TGF_GRILLEDIM2: THLabel
      Left = 112
      Top = 1
      Width = 47
      Height = 13
      Caption = 'GrilledIM2'
    end
    object TGF_GRILLEDIM3: THLabel
      Left = 220
      Top = 1
      Width = 47
      Height = 13
      Caption = 'GrilledIM3'
    end
    object TGF_GRILLEDIM4: THLabel
      Left = 328
      Top = 1
      Width = 47
      Height = 13
      Caption = 'GrilledIM4'
    end
    object TGF_GRILLEDIM5: THLabel
      Left = 436
      Top = 1
      Width = 47
      Height = 13
      Caption = 'GrilledIM5'
    end
    object GF_CODEDIM1: THCritMaskEdit
      Left = 4
      Top = 15
      Width = 100
      Height = 21
      ReadOnly = True
      TabOrder = 0
      Text = 'GF_CODEDIM1'
      TagDispatch = 0
    end
    object GF_CODEDIM2: THCritMaskEdit
      Left = 112
      Top = 15
      Width = 100
      Height = 21
      ReadOnly = True
      TabOrder = 1
      Text = 'GF_CODEDIM2'
      TagDispatch = 0
    end
    object GF_CODEDIM3: THCritMaskEdit
      Left = 220
      Top = 15
      Width = 100
      Height = 21
      ReadOnly = True
      TabOrder = 2
      Text = 'GF_CODEDIM3'
      TagDispatch = 0
    end
    object GF_CODEDIM4: THCritMaskEdit
      Left = 328
      Top = 15
      Width = 100
      Height = 21
      ReadOnly = True
      TabOrder = 3
      Text = 'GF_CODEDIM4'
      TagDispatch = 0
    end
    object GF_CODEDIM5: THCritMaskEdit
      Left = 436
      Top = 15
      Width = 100
      Height = 21
      ReadOnly = True
      TabOrder = 4
      Text = 'GF_CODEDIM5'
      TagDispatch = 0
    end
  end
  object POPZ: TPopupMenu
    Left = 12
    Top = 132
    object InfArticle: TMenuItem
      Caption = '&Voir l'#39'article'
      OnClick = InfArticleClick
    end
  end
  object HMTrad: THSystemMenu
    ResizeDBGrid = True
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 48
    Top = 132
  end
  object FindLigne: TFindDialog
    OnFind = FindLigneFind
    Left = 120
    Top = 132
  end
  object POPU: TPopupMenu
    OnPopup = POPUPopup
    Left = 12
    Top = 164
    object MenuItem3: TMenuItem
      Caption = 'Aller '#224' la &premi'#232're ligne non saisie'
      OnClick = Allerlapremirelignenonsaisie1Click
    end
    object SetSaisi: TMenuItem
      Caption = '&Marquer la ligne comme saisie'
      OnClick = SetSaisiClick
    end
    object SetNonSaisi: TMenuItem
      Caption = '&Marquer la ligne comme non saisie'
      OnClick = SetSaisiClick
    end
  end
end

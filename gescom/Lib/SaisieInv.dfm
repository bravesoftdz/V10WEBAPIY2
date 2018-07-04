object FSaisieInv: TFSaisieInv
  Left = 147
  Top = 204
  Width = 849
  Height = 511
  HelpContext = 110000340
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Saisie Inventaire'
  Color = clBtnFace
  Constraints.MinHeight = 400
  Constraints.MinWidth = 600
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  WindowState = wsMaximized
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object HLabel6: THLabel
    Left = 28
    Top = 20
    Width = 89
    Height = 13
    Caption = 'Dernier prix achat :'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object HLabel7: THLabel
    Left = 28
    Top = 36
    Width = 129
    Height = 13
    Caption = 'Prix moyen achat pond'#233'r'#233' :'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object PENTETE: THPanel
    Left = 0
    Top = 0
    Width = 833
    Height = 69
    Align = alTop
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    DesignSize = (
      833
      69)
    object TGIL_EMPLACEMENT: THLabel
      Left = 8
      Top = 44
      Width = 64
      Height = 13
      Caption = '&Emplacement'
      FocusControl = GIL_EMPLACEMENT
    end
    object TCodeListe: THLabel
      Left = 8
      Top = 8
      Width = 54
      Height = 13
      Caption = 'TCodeListe'
    end
    object TLibelle: THLabel
      Left = 112
      Top = 8
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
      Left = 778
      Top = 24
      Width = 36
      Height = 13
      Alignment = taRightJustify
      Anchors = [akTop, akRight]
      Caption = 'TDepot'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object T_Depot: THLabel
      Left = 779
      Top = 8
      Width = 35
      Height = 13
      Alignment = taRightJustify
      Anchors = [akTop, akRight]
      Caption = 'D'#233'p'#244't :'
    end
    object TDate: THLabel
      Left = 112
      Top = 24
      Width = 30
      Height = 13
      Caption = 'TDate'
    end
    object LConsult: TLabel
      Left = 709
      Top = 47
      Width = 109
      Height = 13
      Anchors = [akTop, akRight]
      Caption = 'Consultation seulement'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsUnderline]
      ParentColor = False
      ParentFont = False
      Visible = False
    end
    object GIL_EMPLACEMENT: THValComboBox
      Left = 112
      Top = 40
      Width = 145
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = GIL_EMPLACEMENTChange
      TagDispatch = 0
      Vide = True
      VideString = 'Tous'
      DataType = 'GCEMPLACEMENT'
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 432
    Width = 833
    Height = 40
    Position = dpBottom
    object ToolWindow971: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 36
      ClientWidth = 824
      Caption = 'Outils inventaire'
      ClientAreaHeight = 36
      ClientAreaWidth = 824
      DockPos = 0
      FullSize = True
      TabOrder = 0
      DesignSize = (
        824
        36)
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
        Caption = 'Fonctions'
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
      object BValider: TToolbarButton97
        Left = 727
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Valider'
        Anchors = [akTop]
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
        Left = 759
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Annuler'
        Anchors = [akTop]
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
        OnClick = BAbandonClick
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 791
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        Anchors = [akTop]
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
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BDetailPrix: TToolbarButton97
        Tag = 1
        Left = 133
        Top = 3
        Width = 28
        Height = 27
        Hint = 'D'#233'tail des prix'
        AllowAllUp = True
        GroupIndex = 2
        DisplayMode = dmGlyphOnly
        Caption = 'D'#233'tail prix'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BDetailPrixClick
        GlobalIndexImage = 'Z0206_S16G1'
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
      object BAddLigne: TToolbarButton97
        Tag = 1
        Left = 175
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Ajouter un article'
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
        GlobalIndexImage = 'Z0074_S16G1'
        IsControl = True
      end
      object BDelLigne: TToolbarButton97
        Tag = 1
        Left = 207
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Supprimer l'#39'article'
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
        GlobalIndexImage = 'Z0072_S16G1'
        IsControl = True
      end
      object bSelectAll: TToolbarButton97
        Tag = 1
        Left = 240
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Tout s'#233'lectionner'
        AllowAllUp = True
        GroupIndex = 2
        DisplayMode = dmGlyphOnly
        Caption = 'D'#233'tail prix'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = bSelectAllClick
        GlobalIndexImage = 'Z0205_S16G1'
        IsControl = True
      end
      object BActualise: TToolbarButton97
        Left = 272
        Top = 3
        Width = 28
        Height = 27
        Hint = 'R'#233'actualisation du stock ordinateur'
        ParentShowHint = False
        ShowHint = True
        OnClick = BActualiseClick
        GlobalIndexImage = 'M0063_S16G1'
      end
      object PageScroller1: TPageScroller
        Left = 307
        Top = -4
        Width = 150
        Height = 45
        TabOrder = 0
      end
    end
  end
  object PPIED: THPanel
    Left = 0
    Top = 403
    Width = 833
    Height = 29
    HelpContext = 110000340
    Align = alBottom
    FullRepaint = False
    TabOrder = 2
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    DesignSize = (
      833
      29)
    object HLabel2: THLabel
      Left = 8
      Top = 8
      Width = 75
      Height = 13
      Caption = 'Unit'#233' de stock :'
    end
    object TUStock: THLabel
      Left = 88
      Top = 8
      Width = 43
      Height = 13
      Caption = 'TUStock'
    end
    object T_Emplacement: THLabel
      Left = 160
      Top = 8
      Width = 70
      Height = 13
      Caption = 'Emplacement :'
    end
    object TEmplacement: THLabel
      Left = 236
      Top = 8
      Width = 71
      Height = 13
      Caption = 'TEmplacement'
    end
    object LTOTQTE: TLabel
      Left = 624
      Top = 8
      Width = 30
      Height = 13
      Anchors = [akTop, akRight]
      Caption = 'Total :'
    end
    object TOTQTESTO: THNumEdit
      Left = 661
      Top = 4
      Width = 56
      Height = 21
      Ctl3D = True
      Decimals = 2
      Digits = 15
      Enabled = False
      Masks.PositiveMask = '#,##0'
      Debit = False
      NumericType = ntDecimal
      ParentCtl3D = False
      TabOrder = 0
      UseRounding = True
      Validate = False
      Anchors = [akTop, akRight]
    end
    object TOTQTEINV: THNumEdit
      Left = 717
      Top = 4
      Width = 50
      Height = 21
      Ctl3D = True
      Decimals = 2
      Digits = 15
      Enabled = False
      Masks.PositiveMask = '#,##0'
      Debit = False
      NumericType = ntDecimal
      ParentCtl3D = False
      TabOrder = 1
      UseRounding = True
      Validate = False
      Anchors = [akTop, akRight]
    end
    object TOTECART: THNumEdit
      Left = 766
      Top = 4
      Width = 52
      Height = 21
      Ctl3D = True
      Decimals = 2
      Digits = 15
      Enabled = False
      Masks.PositiveMask = '#,##0'
      Debit = False
      NumericType = ntDecimal
      ParentCtl3D = False
      TabOrder = 2
      UseRounding = True
      Validate = False
      Anchors = [akTop, akRight]
    end
  end
  object G_Inv: THGrid
    Tag = 1
    Left = 0
    Top = 69
    Width = 833
    Height = 334
    Align = alClient
    ColCount = 9
    DefaultRowHeight = 18
    RowCount = 3
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected, goColSizing, goEditing, goTabs, goThumbTracking]
    TabOrder = 3
    OnDblClick = G_InvDblClick
    OnKeyDown = G_InvKeyDown
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    OnRowEnter = G_InvRowEnter
    OnCellEnter = G_InvCellEnter
    OnCellExit = G_InvCellExit
    ColCombo = 0
    SortEnabled = True
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = clSilver
    ElipsisButton = True
    OnElipsisClick = G_InvElipsisClick
    OnSorted = G_InvSorted
    ColWidths = (
      64
      64
      64
      64
      64
      64
      64
      64
      64)
    RowHeights = (
      18
      18
      18)
  end
  object TDetailPrix: TToolWindow97
    Left = 212
    Top = 156
    ClientHeight = 167
    ClientWidth = 413
    Caption = 'D'#233'tail des prix'
    ClientAreaHeight = 167
    ClientAreaWidth = 413
    DockableTo = []
    Resizable = False
    TabOrder = 4
    Visible = False
    OnClose = TDetailPrixClose
    object grp_depot: TGroupBox
      Left = 2
      Top = 0
      Width = 409
      Height = 53
      Caption = 'D'#233'p'#244't'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      object HLabel4: THLabel
        Left = 8
        Top = 16
        Width = 89
        Height = 13
        Caption = 'Dernier prix achat :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel5: THLabel
        Left = 8
        Top = 32
        Width = 129
        Height = 13
        Caption = 'Prix moyen achat pond'#233'r'#233' :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel9: THLabel
        Left = 208
        Top = 32
        Width = 134
        Height = 13
        Caption = 'Prix moyen revient pond'#233'r'#233' :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel8: THLabel
        Left = 208
        Top = 16
        Width = 94
        Height = 13
        Caption = 'Dernier prix revient :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object GIL_DPA: THNumEdit
        Left = 104
        Top = 16
        Width = 97
        Height = 13
        BorderStyle = bsNone
        Color = clBtnFace
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        ParentFont = False
        ReadOnly = True
        TabOrder = 0
        UseRounding = True
        Validate = False
      end
      object GIL_PMAP: THNumEdit
        Left = 140
        Top = 32
        Width = 61
        Height = 13
        BorderStyle = bsNone
        Color = clBtnFace
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        ParentFont = False
        ReadOnly = True
        TabOrder = 1
        UseRounding = True
        Validate = False
      end
      object GIL_DPR: THNumEdit
        Left = 308
        Top = 16
        Width = 97
        Height = 13
        BorderStyle = bsNone
        Color = clBtnFace
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        ParentFont = False
        ReadOnly = True
        TabOrder = 2
        UseRounding = True
        Validate = False
      end
      object GIL_PMRP: THNumEdit
        Left = 344
        Top = 32
        Width = 61
        Height = 13
        BorderStyle = bsNone
        Color = clBtnFace
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        ParentFont = False
        ReadOnly = True
        TabOrder = 3
        UseRounding = True
        Validate = False
      end
    end
    object grp_article: TGroupBox
      Left = 2
      Top = 56
      Width = 409
      Height = 53
      Caption = 'Article'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      object HLabel10: THLabel
        Left = 8
        Top = 16
        Width = 89
        Height = 13
        Caption = 'Dernier prix achat :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel11: THLabel
        Left = 8
        Top = 32
        Width = 129
        Height = 13
        Caption = 'Prix moyen achat pond'#233'r'#233' :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel12: THLabel
        Left = 208
        Top = 16
        Width = 94
        Height = 13
        Caption = 'Dernier prix revient :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel13: THLabel
        Left = 208
        Top = 32
        Width = 134
        Height = 13
        Caption = 'Prix moyen revient pond'#233'r'#233' :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object GIL_PMAPART: THNumEdit
        Left = 140
        Top = 32
        Width = 61
        Height = 13
        BorderStyle = bsNone
        Color = clBtnFace
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        ParentFont = False
        ReadOnly = True
        TabOrder = 0
        UseRounding = True
        Validate = False
      end
      object GIL_PMRPART: THNumEdit
        Left = 344
        Top = 32
        Width = 61
        Height = 13
        BorderStyle = bsNone
        Color = clBtnFace
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        ParentFont = False
        ReadOnly = True
        TabOrder = 1
        UseRounding = True
        Validate = False
      end
      object GIL_DPRART: THNumEdit
        Left = 308
        Top = 16
        Width = 97
        Height = 13
        BorderStyle = bsNone
        Color = clBtnFace
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        ParentFont = False
        ReadOnly = True
        TabOrder = 2
        UseRounding = True
        Validate = False
      end
      object GIL_DPAART: THNumEdit
        Left = 104
        Top = 16
        Width = 97
        Height = 13
        BorderStyle = bsNone
        Color = clBtnFace
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        ParentFont = False
        ReadOnly = True
        TabOrder = 3
        UseRounding = True
        Validate = False
      end
    end
    object grp_saisi: TGroupBox
      Left = 2
      Top = 112
      Width = 409
      Height = 53
      Caption = 'Saisi'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      object HLabel14: THLabel
        Left = 8
        Top = 16
        Width = 89
        Height = 13
        Caption = 'Dernier prix achat :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel15: THLabel
        Left = 8
        Top = 32
        Width = 129
        Height = 13
        Caption = 'Prix moyen achat pond'#233'r'#233' :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel17: THLabel
        Left = 208
        Top = 32
        Width = 134
        Height = 13
        Caption = 'Prix moyen revient pond'#233'r'#233' :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel16: THLabel
        Left = 208
        Top = 16
        Width = 94
        Height = 13
        Caption = 'Dernier prix revient :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object GIL_PMAPSAIS: THNumEdit
        Left = 140
        Top = 32
        Width = 61
        Height = 13
        BorderStyle = bsNone
        Color = clBtnFace
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        ParentFont = False
        ReadOnly = True
        TabOrder = 0
        UseRounding = True
        Validate = False
      end
      object GIL_PMRPSAIS: THNumEdit
        Left = 344
        Top = 32
        Width = 61
        Height = 13
        BorderStyle = bsNone
        Color = clBtnFace
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        ParentFont = False
        ReadOnly = True
        TabOrder = 1
        UseRounding = True
        Validate = False
      end
      object GIL_DPRSAIS: THNumEdit
        Left = 308
        Top = 16
        Width = 97
        Height = 13
        BorderStyle = bsNone
        Color = clBtnFace
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        ParentFont = False
        ReadOnly = True
        TabOrder = 2
        UseRounding = True
        Validate = False
      end
      object GIL_DPASAIS: THNumEdit
        Left = 104
        Top = 16
        Width = 97
        Height = 13
        BorderStyle = bsNone
        Color = clBtnFace
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        ParentFont = False
        ReadOnly = True
        TabOrder = 3
        UseRounding = True
        Validate = False
      end
    end
  end
  object HMTrad: THSystemMenu
    ResizeDBGrid = True
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 56
    Top = 212
  end
  object POPZ: TPopupMenu
    Left = 8
    Top = 268
    object InfArticle: TMenuItem
      Caption = '&Voir l'#39'article'
      OnClick = InfArticleClick
    end
  end
  object FindLigne: TFindDialog
    OnFind = FindLigneFind
    Left = 88
    Top = 268
  end
  object POPU: TPopupMenu
    OnPopup = POPUPopup
    Left = 12
    Top = 232
    object ModifPrix: TMenuItem
      Caption = '&D'#233'tail des prix'
      OnClick = Voirlesprix1Click
    end
    object JumpNonSaisi: TMenuItem
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
    object SetAllSaisie: TMenuItem
      Caption = 'Marquer &toutes les lignes comme saisies'
      OnClick = SetAllSaisieClick
    end
  end
end

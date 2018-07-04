object FActivite: TFActivite
  Left = 103
  Top = 68
  Width = 942
  Height = 591
  HelpContext = 120000501
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Saisie d'#39'activit'#233
  Color = clBtnFace
  Constraints.MinHeight = 400
  Constraints.MinWidth = 600
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnActivate = FormActivate
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDeactivate = FormDeactivate
  OnKeyDown = FormKeyDown
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Dock971: TDock97
    Left = 0
    Top = 520
    Width = 926
    Height = 33
    BackgroundTransparent = True
    BoundLines = [blTop, blBottom, blLeft, blRight]
    Position = dpBottom
    object Toolbar973: TToolbar97
      Left = 333
      Top = 0
      Hint = 'Valider'
      Caption = 'Valider'
      CloseButton = False
      DockPos = 333
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      DesignSize = (
        140
        27)
      object BAide: TToolbarButton97
        Tag = 1
        Left = 112
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Aide'
        HelpContext = 120000501
        Caption = 'Aide'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 84
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Caption = 'Annuler'
        Anchors = [akTop, akRight]
        Cancel = True
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAbandonClick
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BValider: TToolbarButton97
        Left = 56
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Enregistrer l'#39'activit'#233
        Caption = 'Enregistrer'
        Anchors = [akTop, akRight]
        Default = True
        DisplayMode = dmGlyphOnly
        Flat = False
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
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        OnClick = BValiderClick
        IsControl = True
      end
      object BImprimer: TToolbarButton97
        Left = 0
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Caption = 'Imprimer'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = BImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object bRafraichit: TToolbarButton97
        Left = 28
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Rafraichir la s'#233'lection'
        Caption = 'Rafraichir'
        Anchors = [akTop, akRight]
        Default = True
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = bRafraichitClick
        GlobalIndexImage = 'Z0338_S16G1'
        IsControl = True
      end
    end
    object Toolbar974: TToolbar97
      Left = 0
      Top = 0
      Hint = 'Actions'
      Caption = 'Actions'
      CloseButton = False
      DockPos = 0
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      DesignSize = (
        320
        27)
      object bNewligne: TToolbarButton97
        Left = 68
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Ins'#233'rer une ligne'
        Caption = 'Ins'#233'rer'
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = bNewligneClick
        GlobalIndexImage = 'Z0074_S16G1'
      end
      object bDelLigne: TToolbarButton97
        Left = 96
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Supprimer la ligne en cours'
        Caption = 'Effacer'
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = bDelLigneClick
        GlobalIndexImage = 'Z0072_S16G1'
      end
      object BChercher: TToolbarButton97
        Tag = 1
        Left = 152
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Rechercher dans les lignes'
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
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BChercherClick
        GlobalIndexImage = 'Z0077_S16G1'
        IsControl = True
      end
      object bSupplement: TToolbarButton97
        Left = 180
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Afficher les compl'#233'ments'
        Caption = 'Compl'#233'ments'
        AllowAllUp = True
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = bSupplementClick
        GlobalIndexImage = 'Z0661_S16G1'
      end
      object bDescriptif: TToolbarButton97
        Left = 208
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Afficher le descriptif'
        Caption = 'Descriptif'
        AllowAllUp = True
        GroupIndex = 2
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = bDescriptifClick
        GlobalIndexImage = 'Z0029_S16G1'
      end
      object BMenuZoom: TToolbarButton97
        Tag = -100
        Left = 0
        Top = 0
        Width = 40
        Height = 27
        Hint = 'Menu zoom'
        Caption = 'Zoom'
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
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnMouseEnter = BMenuZoomMouseEnter
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
      object bSelectAll: TToolbarButton97
        Left = 124
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Tout s'#233'lectionner'
        Caption = 'Tout s'#233'lectionner'
        AllowAllUp = True
        GroupIndex = 2
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = bSelectAllClick
        GlobalIndexImage = 'Z0205_S16G1'
      end
      object bConsult: TToolbarButton97
        Left = 236
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Mode consultation'
        Caption = 'Consultation'
        AllowAllUp = True
        GroupIndex = 3
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = bConsultClick
        GlobalIndexImage = 'Z0099_S16G1'
      end
      object b35heures: TToolbarButton97
        Left = 264
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Contr'#244'le 35 heures'
        Caption = 'Contr'#244'le 35 heures'
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = b35heuresClick
        GlobalIndexImage = 'Z1463_S16G1'
      end
      object bPlanning: TToolbarButton97
        Left = 292
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Suivi d'#39'activit'#233
        Caption = 'Suivi d'#39'activit'#233
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = bPlanningClick
        GlobalIndexImage = 'Z1463_S16G1'
      end
      object bInsererLgnAff: TToolbarButton97
        Left = 40
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Ins'#233'rer les lignes d'#39'affaire'
        Caption = 'Ins'#233'rer les lignes d'#39'affaire'
        Anchors = [akTop, akRight]
        Default = True
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = bInsererLgnAffClick
        GlobalIndexImage = 'Z0822_S16G2'
        IsControl = True
      end
    end
  end
  object HPanel1: THPanel
    Left = 0
    Top = 123
    Width = 926
    Height = 295
    Align = alClient
    BevelOuter = bvNone
    FullRepaint = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object ToolbarButton971: TToolbarButton97
      Left = 311
      Top = 140
      Width = 23
      Height = 22
    end
    object HPanel2: THPanel
      Left = 0
      Top = 276
      Width = 926
      Height = 19
      Align = alBottom
      BevelOuter = bvLowered
      FullRepaint = False
      TabOrder = 0
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
      object HLabelUniteBase: THLabel
        Left = 323
        Top = 3
        Width = 46
        Height = 16
        Alignment = taRightJustify
        Caption = '             '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsItalic]
        ParentFont = False
        Transparent = True
      end
      object LHNumEditTOTQTE: THLabel
        Left = 471
        Top = 3
        Width = 8
        Height = 13
        Alignment = taRightJustify
        Caption = '0'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object LHNumEditTOTPR: THLabel
        Left = 577
        Top = 3
        Width = 8
        Height = 13
        Alignment = taRightJustify
        Caption = '0'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object LHNumEditTOTPV: THLabel
        Left = 673
        Top = 3
        Width = 8
        Height = 13
        Alignment = taRightJustify
        Caption = '0'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
    end
    object GS: THGrid
      Tag = 1
      Left = 0
      Top = 0
      Width = 926
      Height = 276
      Align = alClient
      BorderStyle = bsNone
      ColCount = 30
      DefaultColWidth = 50
      DefaultRowHeight = 18
      RowCount = 3
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goEditing, goTabs, goAlwaysShowEditor]
      ParentFont = False
      PopupMenu = POPY
      TabOrder = 1
      OnDblClick = GSDblClick
      OnEnter = GSEnter
      OnExit = GSExit
      OnGetEditMask = GSGetEditMask
      OnMouseDown = GSMouseDown
      OnMouseMove = GSMouseMove
      OnMouseUp = GSMouseUp
      OnSelectCell = GSSelectCell
      OnTopLeftChanged = GSTopLeftChanged
      OnBeforeFlip = GSBeforeFlip
      SortedCol = -1
      Couleur = False
      MultiSelect = True
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
      OnElipsisClick = GSElipsisClick
      ColWidths = (
        19
        16
        108
        209
        78
        68
        55
        82
        50
        50
        50
        50
        50
        50
        50
        50
        50
        50
        50
        50
        50
        50
        50
        50
        50
        50
        50
        50
        50
        50)
    end
    object HNumEditTOTQTE: THNumEdit
      Left = 421
      Top = 241
      Width = 85
      Height = 19
      TabStop = False
      AutoSize = False
      Color = clYellow
      Ctl3D = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentCtl3D = False
      ParentFont = False
      ReadOnly = True
      TabOrder = 2
      Visible = False
      OnChange = HNumEditTOTPRChange
      Decimals = 2
      Digits = 12
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      UseRounding = True
      Value = 1.000000000000000000
      Validate = False
    end
    object HNumEditTOTPR: THNumEdit
      Left = 510
      Top = 242
      Width = 85
      Height = 19
      TabStop = False
      AutoSize = False
      Color = clYellow
      Ctl3D = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentCtl3D = False
      ParentFont = False
      ReadOnly = True
      TabOrder = 3
      Visible = False
      OnChange = HNumEditTOTPRChange
      Decimals = 2
      Digits = 12
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      UseRounding = True
      Value = 1.000000000000000000
      Validate = False
    end
    object HNumEditTOTPV: THNumEdit
      Left = 598
      Top = 242
      Width = 85
      Height = 19
      TabStop = False
      AutoSize = False
      Color = clYellow
      Ctl3D = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentCtl3D = False
      ParentFont = False
      ReadOnly = True
      TabOrder = 4
      Visible = False
      OnChange = HNumEditTOTPRChange
      Decimals = 2
      Digits = 12
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      UseRounding = True
      Value = 1.000000000000000000
      Validate = False
    end
    object BZoomCalendrierAssistant: THBitBtn
      Tag = 100
      Left = 115
      Top = 22
      Width = 50
      Height = 27
      Hint = 'Voir calendrier ressource'
      Caption = 'Calendrier'
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 5
      Visible = False
      OnClick = BZoomCalendrierAssistantClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BZoomPiece: THBitBtn
      Tag = 100
      Left = 9
      Top = 105
      Width = 44
      Height = 27
      Hint = 'Voir pi'#232'ce vente'
      Caption = 'piece '
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 6
      Visible = False
      OnClick = BZoomPieceClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object Bzoompieceach: THBitBtn
      Tag = 100
      Left = 129
      Top = 104
      Width = 105
      Height = 27
      Hint = 'Voir pi'#232'ce achat'
      Caption = 'piece achat'
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 7
      Visible = False
      OnClick = BZoomPieceAchClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BZoomLigneAch: THBitBtn
      Tag = 100
      Left = 55
      Top = 104
      Width = 71
      Height = 27
      Hint = 'Voir ligne achat'
      Caption = 'ligne achat'
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 8
      Visible = False
      OnClick = BZoomLigneAchClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BZoomTiers: THBitBtn
      Tag = 100
      Left = 229
      Top = 22
      Width = 39
      Height = 27
      Hint = 'Voir tiers'
      Caption = 'Tiers'
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 9
      Visible = False
      OnClick = BZoomTiersClick
      Layout = blGlyphTop
      Spacing = -1
    end
  end
  object BZoomArticle: THBitBtn
    Tag = 100
    Left = 6
    Top = 143
    Width = 37
    Height = 27
    Hint = 'Voir article'
    Caption = 'Art'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 3
    Visible = False
    OnClick = BZoomArticleClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BZoomAssistant: THBitBtn
    Tag = 100
    Left = 48
    Top = 143
    Width = 61
    Height = 27
    Hint = 'Voir ressource'
    Caption = 'Ressource'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 4
    Visible = False
    OnClick = BZoomAssistantClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object GP_FACTUREHT: TCheckBox
    Left = 115
    Top = 194
    Width = 108
    Height = 17
    Caption = 'GP_FACTUREHT'
    Checked = True
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentColor = False
    ParentFont = False
    State = cbChecked
    TabOrder = 5
    Visible = False
  end
  object BZoomAffaire: THBitBtn
    Tag = 100
    Left = 11
    Top = 178
    Width = 39
    Height = 27
    Hint = 'Voir affaire'
    Caption = 'Affaire'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 6
    Visible = False
    OnClick = BZoomAffaireClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object TDescriptif: TToolWindow97
    Left = 388
    Top = 153
    ClientHeight = 119
    ClientWidth = 265
    Caption = 'Descriptif compl'#233'mentaire de la ligne'
    ClientAreaHeight = 119
    ClientAreaWidth = 265
    TabOrder = 7
    Visible = False
    OnClose = TDescriptifClose
    object ACT_DESCRIPTIF: THRichEditOLE
      Left = 0
      Top = 0
      Width = 265
      Height = 119
      Align = alClient
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ScrollBars = ssBoth
      TabOrder = 0
      Margins.Top = 0
      Margins.Bottom = 0
      Margins.Left = 0
      Margins.Right = 0
      ContainerName = 'Document'
      ObjectMenuPrefix = '&Object'
      LinesRTF.Strings = (
        
          '{\rtf1\ansi\ansicpg1252\deff0\deflang1036{\fonttbl{\f0\fnil MS S' +
          'ans Serif;}}'
        
          '{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\pard\f0\fs16' +
          ' '
        '\par '
        '\par '
        '\par '
        '\par '
        '\par }')
    end
  end
  object panel3: THPanel
    Left = 0
    Top = 0
    Width = 926
    Height = 123
    Align = alTop
    BevelInner = bvRaised
    BevelOuter = bvLowered
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object PEntete: THPanel
      Left = 2
      Top = 25
      Width = 922
      Height = 95
      Align = alTop
      BevelOuter = bvNone
      FullRepaint = False
      TabOrder = 1
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
      object LabelMois: TLabel
        Left = 3
        Top = 53
        Width = 22
        Height = 13
        Caption = 'Mois'
        Transparent = True
      end
      object LabelSemaine: TLabel
        Left = 3
        Top = 77
        Width = 41
        Height = 13
        Caption = 'Semaine'
        Transparent = True
      end
      object LabelAnnee: TLabel
        Left = 3
        Top = 30
        Width = 31
        Height = 13
        Caption = 'Ann'#233'e'
        FocusControl = ACT_ANNEE
        Transparent = True
      end
      object LblIntervalle: THLabel
        Left = 130
        Top = 54
        Width = 50
        Height = 13
        Constraints.MaxWidth = 200
        Constraints.MinWidth = 50
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Transparent = True
      end
      object LblFolio: THLabel
        Left = 127
        Top = 77
        Width = 22
        Height = 13
        Caption = 'Folio'
        Transparent = True
      end
      object LblTri: THLabel
        Left = 202
        Top = 77
        Width = 12
        Height = 13
        Caption = '&Tri'
        FocusControl = HValComboBoxTri
        Transparent = True
      end
      object LabelAffaire: THLabel
        Left = 3
        Top = 6
        Width = 30
        Height = 13
        Caption = 'Affaire'
        FocusControl = ACT_AFFAIRE1
        Transparent = True
      end
      object BRechAffaire: TToolbarButton97
        Left = 45
        Top = 1
        Width = 23
        Height = 22
        Hint = 'Recherche Affaire'
        Color = clNone
        Opaque = False
        ShowBorderWhenInactive = True
        OnClick = BRechAffaireClick
        GlobalIndexImage = 'Z0002_S16G1'
      end
      object LibelleAffaire: THLabel
        Left = 241
        Top = 4
        Width = 89
        Height = 16
        Caption = '                      '
        Constraints.MaxWidth = 190
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object LabelMoisClotClient: TLabel
        Left = 136
        Top = 30
        Width = 100
        Height = 13
        Caption = 'Mois de cl'#244'ture client'
        Transparent = True
      end
      object LblIntervalle2: THLabel
        Left = 142
        Top = 54
        Width = 50
        Height = 13
        Constraints.MaxWidth = 200
        Constraints.MinWidth = 50
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Transparent = True
      end
      object GCalendrier: THGrid
        Left = 623
        Top = 0
        Width = 299
        Height = 95
        Align = alRight
        Color = clWhite
        ColCount = 9
        Ctl3D = False
        DefaultColWidth = 40
        DefaultRowHeight = 12
        DefaultDrawing = False
        FixedCols = 0
        RowCount = 7
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -9
        Font.Name = 'Small Fonts'
        Font.Style = []
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected]
        ParentCtl3D = False
        ParentFont = False
        ParentShowHint = False
        ScrollBars = ssVertical
        ShowHint = True
        TabOrder = 10
        OnExit = GCalendrierExit
        OnMouseDown = GCalendrierMouseDown
        SortedCol = -1
        Couleur = False
        MultiSelect = True
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = clSilver
      end
      object ACT_ANNEE: TSpinEdit
        Left = 68
        Top = 25
        Width = 59
        Height = 22
        AutoSize = False
        Constraints.MaxWidth = 59
        Constraints.MinWidth = 59
        MaxValue = 2100
        MinValue = 1900
        TabOrder = 5
        Value = 2000
      end
      object ACT_MOIS: TSpinEdit
        Left = 68
        Top = 49
        Width = 42
        Height = 22
        AutoSize = False
        Constraints.MaxWidth = 42
        Constraints.MinWidth = 42
        MaxValue = 13
        MinValue = 0
        TabOrder = 6
        Value = 1
        OnChange = ACT_MOISChange
      end
      object ACT_SEMAINE: TSpinEdit
        Left = 68
        Top = 74
        Width = 42
        Height = 22
        AutoSize = False
        Constraints.MaxWidth = 42
        Constraints.MinWidth = 42
        MaxValue = 52
        MinValue = 1
        TabOrder = 7
        Value = 1
      end
      object ACT_FOLIO: TSpinEdit
        Left = 155
        Top = 74
        Width = 42
        Height = 22
        AutoSize = False
        Constraints.MaxWidth = 42
        Constraints.MinWidth = 42
        MaxValue = 999
        MinValue = 1
        TabOrder = 8
        Value = 1
      end
      object HValComboBoxTri: THValComboBox
        Left = 220
        Top = 75
        Width = 119
        Height = 21
        Style = csDropDownList
        Constraints.MaxWidth = 119
        Constraints.MinWidth = 119
        ItemHeight = 13
        TabOrder = 9
        TagDispatch = 0
        DisableTab = True
      end
      object ACT_AFFAIRE3: THCritMaskEdit
        Left = 195
        Top = 2
        Width = 20
        Height = 21
        AutoSize = False
        CharCase = ecUpperCase
        TabOrder = 3
        OnDblClick = ACT_AFFAIRE3DblClick
        OnExit = ACT_AFFAIRE3Exit
        TagDispatch = 0
      end
      object ACT_AFFAIRE2: THCritMaskEdit
        Left = 141
        Top = 2
        Width = 54
        Height = 21
        AutoSize = False
        CharCase = ecUpperCase
        TabOrder = 2
        OnChange = ACT_AFFAIRE2Change
        OnDblClick = ACT_AFFAIRE2DblClick
        OnExit = ACT_AFFAIRE2Exit
        TagDispatch = 0
      end
      object ACT_AFFAIRE1: THCritMaskEdit
        Left = 101
        Top = 2
        Width = 42
        Height = 21
        AutoSize = False
        CharCase = ecUpperCase
        TabOrder = 1
        OnChange = ACT_AFFAIRE1Change
        OnDblClick = ACT_AFFAIRE1DblClick
        OnExit = ACT_AFFAIRE1Exit
        TagDispatch = 0
      end
      object ACT_AFFAIRE0: THCritMaskEdit
        Left = 68
        Top = 2
        Width = 31
        Height = 21
        AutoSize = False
        CharCase = ecUpperCase
        TabOrder = 0
        OnChange = ACT_AFFAIRE0Change
        OnDblClick = ACT_AFFAIRE0DblClick
        OnExit = ACT_AFFAIRE0Exit
        TagDispatch = 0
      end
      object ACT_AVENANT: THCritMaskEdit
        Left = 215
        Top = 2
        Width = 22
        Height = 21
        AutoSize = False
        CharCase = ecUpperCase
        TabOrder = 4
        Visible = False
        OnDblClick = ACT_AFFAIRE3DblClick
        OnExit = ACT_AFFAIRE3Exit
        TagDispatch = 0
      end
      object MOIS_CLOT_CLIENT: THCritMaskEdit
        Tag = -1
        Left = 244
        Top = 26
        Width = 22
        Height = 21
        Hint = 'Mois de cl'#244'ture client'
        TabStop = False
        AutoSize = False
        CharCase = ecUpperCase
        Color = clBtnFace
        ParentShowHint = False
        ReadOnly = True
        ShowHint = True
        TabOrder = 11
        Visible = False
        TagDispatch = 0
      end
    end
    object PEntete2: THPanel
      Left = 2
      Top = 2
      Width = 922
      Height = 23
      Align = alTop
      BevelOuter = bvNone
      FullRepaint = False
      TabOrder = 0
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
      DesignSize = (
        922
        23)
      object LabelAssistant: TLabel
        Left = 4
        Top = 5
        Width = 51
        Height = 13
        Caption = 'Ressource'
        FocusControl = ACT_RESSOURCE
        Transparent = True
      end
      object LIBELLEASSISTANT: THLabel
        Left = 211
        Top = 4
        Width = 117
        Height = 16
        Anchors = [akLeft]
        Caption = '                             '
        Color = clInfoBk
        Constraints.MaxWidth = 210
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
        Transparent = True
      end
      object LblMois: THLabel
        Left = 857
        Top = 0
        Width = 63
        Height = 20
        Alignment = taCenter
        Caption = 'LblMois'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -16
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsItalic]
        ParentFont = False
        Transparent = True
      end
      object LIBELLETIERS: THLabel
        Left = 209
        Top = 3
        Width = 117
        Height = 16
        Anchors = [akLeft]
        Caption = '                             '
        Color = clInfoBk
        Constraints.MaxWidth = 210
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
        Transparent = True
      end
      object bEffaceAffaireTiers: TToolbarButton97
        Left = 45
        Top = 1
        Width = 23
        Height = 22
        Hint = 'Nouvelle s'#233'lection'
        Opaque = False
        ShowBorderWhenInactive = True
        OnClick = bEffaceAffaireTiersClick
        GlobalIndexImage = 'Z0080_S16G1'
      end
      object BMoisMoins: TToolbarButton97
        Left = 452
        Top = 0
        Width = 23
        Height = 22
        Hint = 'Mois pr'#233'c'#233'dent'
        NumGlyphs = 2
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BMoisMoinsClick
        GlobalIndexImage = 'Z0077_S16G2'
      end
      object BMoisPlus: TToolbarButton97
        Left = 662
        Top = 0
        Width = 23
        Height = 22
        Hint = 'Mois suivant'
        NumGlyphs = 2
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BMoisPlusClick
        GlobalIndexImage = 'Z0056_S16G2'
      end
      object LIBELLEUSER: THLabel
        Left = 222
        Top = 7
        Width = 117
        Height = 16
        Anchors = [akLeft]
        Caption = '                             '
        Color = clInfoBk
        Constraints.MaxWidth = 210
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
        Transparent = True
        Visible = False
      end
      object ACT_RESSOURCE: THCritMaskEdit
        Left = 68
        Top = 1
        Width = 136
        Height = 21
        AutoSize = False
        CharCase = ecUpperCase
        Constraints.MaxWidth = 136
        Constraints.MinWidth = 136
        TabOrder = 0
        OnChange = ACT_RESSOURCEChange
        OnDblClick = ACT_RESSOURCEDblClick
        OnExit = ACT_RESSOURCEExit
        TagDispatch = 0
        ElipsisButton = True
        OnElipsisClick = ACT_RESSOURCEElipsisClick
      end
      object ACT_AFFAIRE: THCritMaskEdit
        Left = 725
        Top = 1
        Width = 27
        Height = 21
        CharCase = ecUpperCase
        Color = clYellow
        TabOrder = 1
        Visible = False
        OnChange = ACT_AFFAIREChange
        TagDispatch = 0
        ElipsisAutoHide = True
      end
      object ACT_TIERS: THCritMaskEdit
        Left = 72
        Top = 5
        Width = 136
        Height = 21
        AutoSize = False
        CharCase = ecUpperCase
        Constraints.MaxWidth = 136
        Constraints.MinWidth = 136
        TabOrder = 2
        OnChange = ACT_TIERSChange
        OnDblClick = ACT_TIERSDblClick
        OnExit = ACT_TIERSExit
        TagDispatch = 0
        DataType = 'GCTIERSSAISIE'
        ElipsisButton = True
        OnElipsisClick = ACT_TIERSElipsisClick
      end
      object USER: THCritMaskEdit
        Left = 76
        Top = 10
        Width = 136
        Height = 21
        TabStop = False
        AutoSize = False
        CharCase = ecUpperCase
        Constraints.MaxWidth = 136
        Constraints.MinWidth = 136
        TabOrder = 3
        Visible = False
        TagDispatch = 0
      end
    end
  end
  object HPBas: THPanel
    Left = 0
    Top = 418
    Width = 926
    Height = 102
    Align = alBottom
    BevelOuter = bvNone
    FullRepaint = False
    TabOrder = 8
    Visible = False
    BackGroundEffect = bdFond
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    ColorNb = 2
    TextEffect = tenone
    object PagePied: TPageControl
      Left = 0
      Top = 0
      Width = 926
      Height = 102
      ActivePage = SheetPied
      Align = alClient
      MultiLine = True
      TabOrder = 0
      TabPosition = tpLeft
      object SheetPied: TTabSheet
        Caption = 'D'#233'tail'
        object PPied: THPanel
          Left = 0
          Top = 0
          Width = 899
          Height = 94
          Align = alClient
          BevelOuter = bvNone
          FullRepaint = False
          TabOrder = 0
          BackGroundEffect = bdFond
          ColorShadow = clWindowText
          ColorStart = clBtnFace
          ColorNb = 2
          TextEffect = tenone
          object PTotaux1: THPanel
            Left = 334
            Top = 0
            Width = 300
            Height = 94
            Align = alRight
            Anchors = [akTop, akBottom]
            AutoSize = True
            BevelInner = bvRaised
            BevelOuter = bvLowered
            Caption = ' '
            Constraints.MinWidth = 300
            FullRepaint = False
            TabOrder = 1
            OnEnter = PTotaux1Enter
            BackGroundEffect = bdFond
            ColorShadow = clWindowText
            ColorStart = clBtnFace
            ColorNb = 2
            TextEffect = tenone
            object LblPrixUnit: THLabel
              Left = 147
              Top = 30
              Width = 54
              Height = 13
              Caption = 'Prix unitaire'
              Transparent = True
            end
            object LblMontantTot: THLabel
              Left = 220
              Top = 30
              Width = 62
              Height = 13
              Caption = 'Montant total'
              Transparent = True
            end
            object LblPRCharge: THLabel
              Left = 122
              Top = 49
              Width = 15
              Height = 13
              Caption = 'PR'
              Transparent = True
            end
            object LblPV: THLabel
              Left = 122
              Top = 71
              Width = 14
              Height = 13
              Caption = 'PV'
              Transparent = True
            end
            object LblUnite: THLabel
              Left = 2
              Top = 49
              Width = 25
              Height = 13
              Caption = 'Unit'#233
              Transparent = True
            end
            object LblQte: THLabel
              Left = 2
              Top = 71
              Width = 40
              Height = 13
              Caption = 'Quantit'#233
              Transparent = True
            end
            object LblDate: THLabel
              Left = 2
              Top = 5
              Width = 23
              Height = 13
              Caption = 'Date'
              Color = clInfoBk
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = []
              ParentColor = False
              ParentFont = False
              Transparent = True
            end
            object LblActiviteReprise: THLabel
              Left = 175
              Top = 6
              Width = 69
              Height = 13
              Caption = 'Activit'#233' reprise'
              Transparent = True
            end
            object LibTypeHeure: THLabel
              Left = 2
              Top = 28
              Width = 54
              Height = 13
              Caption = 'Heures sup'
              Transparent = True
            end
            object ACT_PUPRCHARGE: THNumEdit
              Left = 141
              Top = 45
              Width = 69
              Height = 21
              AutoSize = False
              TabOrder = 3
              OnEnter = ACT_PUPRCHARGEEnter
              OnExit = ACT_PUPRCHARGEExit
              Decimals = 2
              Digits = 12
              Masks.PositiveMask = '#,##0.00'
              Debit = False
              UseRounding = True
              Validate = False
            end
            object ACT_TOTPRCHARGE: THNumEdit
              Left = 213
              Top = 45
              Width = 81
              Height = 21
              AutoSize = False
              TabOrder = 4
              OnEnter = ACT_TOTPRCHARGEEnter
              OnExit = ACT_TOTPRCHARGEExit
              Decimals = 2
              Digits = 12
              Masks.PositiveMask = '#,##0.00'
              Debit = False
              UseRounding = True
              Validate = False
            end
            object ACT_PUVENTE: THNumEdit
              Left = 141
              Top = 67
              Width = 69
              Height = 21
              AutoSize = False
              TabOrder = 5
              OnEnter = ACT_PUVENTEEnter
              OnExit = ACT_PUVENTEExit
              Decimals = 2
              Digits = 12
              Masks.PositiveMask = '#,##0.00'
              Debit = False
              UseRounding = True
              Validate = False
            end
            object ACT_TOTVENTE: THNumEdit
              Left = 213
              Top = 67
              Width = 81
              Height = 21
              AutoSize = False
              TabOrder = 6
              OnEnter = ACT_TOTVENTEEnter
              OnExit = ACT_TOTVENTEExit
              Decimals = 2
              Digits = 12
              Masks.PositiveMask = '#,##0.00'
              Debit = False
              UseRounding = True
              Validate = False
            end
            object ACT_UNITE: THCritMaskEdit
              Tag = -1
              Left = 61
              Top = 45
              Width = 49
              Height = 21
              TabStop = False
              AutoSize = False
              Color = clBtnFace
              ReadOnly = True
              TabOrder = 1
              OnEnter = ACT_UNITEEnter
              OnExit = ACT_UNITEExit
              TagDispatch = 0
              OnElipsisClick = ACT_UNITEElipsisClick
            end
            object ACT_QTE: THNumEdit
              Left = 61
              Top = 67
              Width = 49
              Height = 21
              AutoSize = False
              TabOrder = 2
              OnEnter = ACT_QTEEnter
              OnExit = ACT_QTEExit
              Decimals = 2
              Digits = 12
              Masks.PositiveMask = '#,##0.00'
              Debit = False
              UseRounding = True
              Validate = False
            end
            object ACT_DATEACTIVITE: THCritMaskEdit
              Tag = -1
              Left = 61
              Top = 2
              Width = 75
              Height = 21
              TabStop = False
              AutoSize = False
              Color = clBtnFace
              EditMask = '!99/99/0000;1;_'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = []
              MaxLength = 10
              ParentFont = False
              ReadOnly = True
              TabOrder = 7
              Text = '  /  /    '
              TagDispatch = 0
              OpeType = otDate
              ControlerDate = True
            end
            object ACT_ACTIVITEREPRIS: THCritMaskEdit
              Left = 258
              Top = 2
              Width = 36
              Height = 21
              AutoSize = False
              TabOrder = 0
              OnEnter = ACT_ACTIVITEREPRISEnter
              OnExit = ACT_ACTIVITEREPRISExit
              TagDispatch = 0
              Plus = 'AND (CO_LIBRE="ACT" OR CO_LIBRE="ART")'
              DataType = 'AFTACTIVITEREPRISE'
              ElipsisButton = True
            end
            object ACT_TYPEHEURE: THCritMaskEdit
              Left = 61
              Top = 23
              Width = 49
              Height = 21
              AutoSize = False
              CharCase = ecUpperCase
              TabOrder = 8
              OnEnter = ACT_TYPEHEUREEnter
              OnExit = ACT_TYPEHEUREExit
              TagDispatch = 0
              DataType = 'AFTTYPEHEURE'
              ElipsisButton = True
            end
          end
          object pInfo1: THPanel
            Left = 0
            Top = 0
            Width = 334
            Height = 94
            Align = alClient
            Anchors = [akLeft, akTop, akBottom]
            AutoSize = True
            BevelInner = bvRaised
            BevelOuter = bvLowered
            Caption = ' '
            Constraints.MinWidth = 334
            FullRepaint = False
            ParentCtl3D = False
            TabOrder = 0
            OnEnter = pInfo1Enter
            BackGroundEffect = bdFond
            ColorShadow = clWindowText
            ColorStart = clBtnFace
            ColorNb = 2
            TextEffect = tenone
            DesignSize = (
              334
              94)
            object LblAffaire: THLabel
              Left = 4
              Top = 5
              Width = 30
              Height = 13
              Anchors = [akLeft]
              Caption = 'Affaire'
              Color = clInfoBk
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = []
              ParentColor = False
              ParentFont = False
              Transparent = True
            end
            object LblArticle: THLabel
              Left = 4
              Top = 48
              Width = 29
              Height = 13
              Anchors = [akLeft]
              Caption = 'Article'
              Color = clInfoBk
              ParentColor = False
              Transparent = True
            end
            object HLabelLibelle: THLabel
              Left = 4
              Top = 71
              Width = 30
              Height = 13
              Anchors = [akLeft]
              Caption = 'Libell'#233
              Transparent = True
            end
            object ACT_AFFAIREL: THLabel
              Left = 224
              Top = 6
              Width = 85
              Height = 16
              Anchors = [akLeft]
              Caption = '                     '
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -13
              Font.Name = 'MS Sans Serif'
              Font.Style = [fsBold]
              ParentFont = False
              Transparent = True
            end
            object ACT_ARTICLEL: THLabel
              Left = 181
              Top = 48
              Width = 73
              Height = 16
              Anchors = [akLeft]
              Caption = '                  '
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -13
              Font.Name = 'MS Sans Serif'
              Font.Style = [fsBold]
              ParentFont = False
              Transparent = True
            end
            object LblClient: THLabel
              Left = 4
              Top = 26
              Width = 26
              Height = 13
              Anchors = [akLeft]
              Caption = 'Client'
              Color = clInfoBk
              ParentColor = False
              Transparent = True
            end
            object ACT_TIERSL: THLabel
              Left = 202
              Top = 27
              Width = 69
              Height = 16
              Anchors = [akLeft]
              Caption = '                 '
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -13
              Font.Name = 'MS Sans Serif'
              Font.Style = [fsBold]
              ParentFont = False
              Transparent = True
            end
            object ACT_RESSOURCEL: THLabel
              Left = 178
              Top = 10
              Width = 85
              Height = 16
              Anchors = [akLeft]
              Caption = '                     '
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -13
              Font.Name = 'MS Sans Serif'
              Font.Style = [fsBold]
              ParentFont = False
              Transparent = True
            end
            object LblActiviteEffect: THLabel
              Left = 286
              Top = 71
              Width = 67
              Height = 13
              Anchors = [akLeft]
              Caption = 'Travail effectif'
              Transparent = True
            end
            object ACT_AFFAIRE3_: THCritMaskEdit
              Tag = -1
              Left = 179
              Top = 3
              Width = 24
              Height = 21
              TabStop = False
              Anchors = [akLeft]
              AutoSize = False
              CharCase = ecUpperCase
              Color = clBtnFace
              ReadOnly = True
              TabOrder = 8
              Visible = False
              TagDispatch = 0
            end
            object ACT_AFFAIRE2_: THCritMaskEdit
              Tag = -1
              Left = 128
              Top = 3
              Width = 51
              Height = 21
              TabStop = False
              Anchors = [akLeft]
              AutoSize = False
              CharCase = ecUpperCase
              Color = clBtnFace
              ReadOnly = True
              TabOrder = 7
              Visible = False
              TagDispatch = 0
            end
            object ACT_AFFAIRE1_: THCritMaskEdit
              Tag = -1
              Left = 85
              Top = 3
              Width = 42
              Height = 21
              TabStop = False
              Anchors = [akLeft]
              AutoSize = False
              CharCase = ecUpperCase
              Color = clBtnFace
              ReadOnly = True
              TabOrder = 6
              Visible = False
              TagDispatch = 0
            end
            object ACT_AFFAIRE_: THCritMaskEdit
              Tag = -1
              Left = 268
              Top = 16
              Width = 17
              Height = 21
              TabStop = False
              AutoSize = False
              Color = clYellow
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = [fsBold]
              ParentFont = False
              ReadOnly = True
              TabOrder = 0
              Visible = False
              OnChange = ACT_AFFAIRE_Change
              TagDispatch = 0
            end
            object ACT_CODEARTICLE: THCritMaskEdit
              Left = 55
              Top = 46
              Width = 121
              Height = 21
              Anchors = [akLeft]
              AutoSize = False
              TabOrder = 3
              OnChange = ACT_CODEARTICLEChange
              OnEnter = ACT_CODEARTICLEEnter
              OnExit = ACT_CODEARTICLEExit
              TagDispatch = 0
              ElipsisButton = True
              OnElipsisClick = ACT_CODEARTICLEElipsisClick
            end
            object ACT_LIBELLE: THCritMaskEdit
              Left = 55
              Top = 67
              Width = 207
              Height = 21
              Anchors = [akLeft]
              AutoSize = False
              TabOrder = 4
              OnEnter = ACT_LIBELLEEnter
              OnExit = ACT_LIBELLEExit
              TagDispatch = 0
            end
            object ACT_TIERS_: THCritMaskEdit
              Tag = -1
              Left = 79
              Top = 25
              Width = 121
              Height = 21
              TabStop = False
              Anchors = [akLeft]
              AutoSize = False
              Color = clBtnFace
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = []
              ParentFont = False
              ReadOnly = True
              TabOrder = 1
              OnChange = ACT_TIERS_Change
              TagDispatch = 0
            end
            object ACT_ARTICLE: THCritMaskEdit
              Tag = -1
              Left = 269
              Top = 38
              Width = 16
              Height = 21
              TabStop = False
              AutoSize = False
              Color = clYellow
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clRed
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = []
              ParentFont = False
              TabOrder = 5
              Visible = False
              OnExit = ACT_CODEARTICLEExit
              TagDispatch = 0
              ElipsisAutoHide = True
            end
            object ACT_AFFAIRE0_: THCritMaskEdit
              Tag = -1
              Left = 55
              Top = 3
              Width = 29
              Height = 21
              TabStop = False
              Anchors = [akLeft]
              AutoSize = False
              CharCase = ecUpperCase
              Color = clBtnFace
              ReadOnly = True
              TabOrder = 9
              Text = 'LIDE'
              Visible = False
              TagDispatch = 0
            end
            object ACT_RESSOURCE_: THCritMaskEdit
              Tag = -1
              Left = 59
              Top = 6
              Width = 121
              Height = 21
              TabStop = False
              Anchors = [akLeft]
              AutoSize = False
              Color = clBtnFace
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = []
              ParentFont = False
              ReadOnly = True
              TabOrder = 2
              Visible = False
              OnChange = ACT_RESSOURCE_Change
              TagDispatch = 0
            end
            object ACT_AVENANT_: THCritMaskEdit
              Tag = -1
              Left = 203
              Top = 3
              Width = 20
              Height = 21
              TabStop = False
              Anchors = [akLeft]
              AutoSize = False
              CharCase = ecUpperCase
              Color = clBtnFace
              ReadOnly = True
              TabOrder = 10
              Visible = False
              TagDispatch = 0
            end
            object MOIS_CLOT_CLIENT_: THCritMaskEdit
              Tag = -1
              Left = 55
              Top = 25
              Width = 22
              Height = 21
              Hint = 'Mois de cl'#244'ture client'
              TabStop = False
              Anchors = [akLeft]
              AutoSize = False
              CharCase = ecUpperCase
              Color = clBtnFace
              ParentShowHint = False
              ReadOnly = True
              ShowHint = True
              TabOrder = 11
              TagDispatch = 0
            end
            object ACT_ACTIVITEEFFECT: TCheckBox
              Left = 267
              Top = 69
              Width = 15
              Height = 17
              Anchors = [akLeft]
              Enabled = False
              TabOrder = 12
            end
          end
          object PTotaux2: THPanel
            Left = 634
            Top = 0
            Width = 265
            Height = 94
            Align = alRight
            AutoSize = True
            BevelInner = bvRaised
            BevelOuter = bvLowered
            Constraints.MinWidth = 265
            FullRepaint = False
            TabOrder = 2
            OnEnter = PTotaux2Enter
            BackGroundEffect = bdFond
            ColorShadow = clWindowText
            ColorStart = clBtnFace
            ColorNb = 2
            TextEffect = tenone
            object LblMoisTotaux: THLabel
              Left = 124
              Top = 4
              Width = 8
              Height = 13
              Caption = '_'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = [fsBold]
              ParentFont = False
              Transparent = True
            end
            object lblTitreTotaux2: THLabel
              Left = 2
              Top = 4
              Width = 97
              Height = 13
              Caption = 'Cumuls mensuels'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = [fsBold]
              ParentFont = False
              Transparent = True
            end
            object LblTemps: THLabel
              Left = 6
              Top = 20
              Width = 50
              Height = 13
              Caption = 'Activit'#233' en'
              Transparent = True
            end
            object lblFacturable: THLabel
              Left = 19
              Top = 36
              Width = 81
              Height = 13
              Caption = 'Dont Facturables'
              Transparent = True
            end
            object lblNonFacturable: THLabel
              Left = 19
              Top = 52
              Width = 104
              Height = 13
              Caption = 'Dont Non Facturables'
              Transparent = True
            end
            object lblFrais: THLabel
              Left = 5
              Top = 70
              Width = 99
              Height = 13
              Caption = 'Montant des Frais en'
              Transparent = True
            end
            object CumulTemps: THLabel
              Left = 194
              Top = 21
              Width = 8
              Height = 13
              Alignment = taRightJustify
              Caption = '0'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = [fsBold]
              ParentFont = False
              Transparent = True
            end
            object CumulFac: THLabel
              Left = 194
              Top = 36
              Width = 8
              Height = 13
              Alignment = taRightJustify
              Caption = '0'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = [fsBold]
              ParentFont = False
              Transparent = True
            end
            object CumulNonFac: THLabel
              Left = 194
              Top = 52
              Width = 8
              Height = 13
              Alignment = taRightJustify
              Caption = '0'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = [fsBold]
              ParentFont = False
              Transparent = True
            end
            object CumulFrais: THLabel
              Left = 194
              Top = 70
              Width = 8
              Height = 13
              Alignment = taRightJustify
              Caption = '0'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = [fsBold]
              ParentFont = False
              Transparent = True
            end
            object PourcFac: THLabel
              Left = 236
              Top = 36
              Width = 20
              Height = 13
              Alignment = taRightJustify
              Caption = '(0%)'
              Transparent = True
            end
            object PourcNonFac: THLabel
              Left = 236
              Top = 52
              Width = 20
              Height = 13
              Alignment = taRightJustify
              Caption = '(0%)'
              Transparent = True
            end
          end
        end
      end
      object SheetPiedCompl: TTabSheet
        Caption = 'Avanc'#233
        ImageIndex = 1
        object PPiedAvance: THPanel
          Left = 0
          Top = 0
          Width = 899
          Height = 94
          Align = alClient
          Anchors = [akLeft, akRight, akBottom]
          BevelOuter = bvNone
          FullRepaint = False
          TabOrder = 0
          BackGroundEffect = bdFond
          ColorShadow = clWindowFrame
          ColorStart = clBtnFace
          ColorNb = 0
          TextEffect = tenone
          object Ptotaux3: THPanel
            Left = 599
            Top = 0
            Width = 300
            Height = 94
            Align = alRight
            AutoSize = True
            BevelInner = bvRaised
            BevelOuter = bvLowered
            Caption = ' '
            Constraints.MinWidth = 300
            FullRepaint = False
            TabOrder = 0
            OnEnter = PTotaux1Enter
            BackGroundEffect = bdFond
            ColorShadow = clWindowText
            ColorStart = clBtnFace
            ColorNb = 2
            TextEffect = tenone
            object lblprixunit_: THLabel
              Left = 147
              Top = 30
              Width = 54
              Height = 13
              Caption = 'Prix unitaire'
              Transparent = True
            end
            object lblmontanttot_: THLabel
              Left = 222
              Top = 30
              Width = 62
              Height = 13
              Caption = 'Montant total'
              Transparent = True
            end
            object lblprcharge_: THLabel
              Left = 122
              Top = 49
              Width = 15
              Height = 13
              Caption = 'PR'
              Transparent = True
            end
            object lblpv_: THLabel
              Left = 122
              Top = 71
              Width = 14
              Height = 13
              Caption = 'PV'
              Transparent = True
            end
            object lblunite_: THLabel
              Left = 2
              Top = 49
              Width = 25
              Height = 13
              Caption = 'Unit'#233
              Transparent = True
            end
            object lblqte_: THLabel
              Left = 2
              Top = 71
              Width = 40
              Height = 13
              Caption = 'Quantit'#233
              Transparent = True
            end
            object lbldate_: THLabel
              Left = 2
              Top = 5
              Width = 23
              Height = 13
              Caption = 'Date'
              Color = clInfoBk
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = []
              ParentColor = False
              ParentFont = False
              Transparent = True
            end
            object lblactivitereprise_: THLabel
              Left = 175
              Top = 6
              Width = 69
              Height = 13
              Caption = 'Activit'#233' reprise'
              Transparent = True
            end
            object lbltypeheure_: THLabel
              Left = 2
              Top = 28
              Width = 54
              Height = 13
              Caption = 'Heures sup'
              Transparent = True
            end
            object act_puprcharge_: THNumEdit
              Left = 141
              Top = 45
              Width = 69
              Height = 21
              AutoSize = False
              TabOrder = 3
              OnEnter = ACT_PUPRCHARGEEnter
              OnExit = ACT_PUPRCHARGEExit
              Decimals = 2
              Digits = 12
              Masks.PositiveMask = '#,##0.00'
              Debit = False
              UseRounding = True
              Validate = False
            end
            object act_totprcharge_: THNumEdit
              Left = 215
              Top = 45
              Width = 81
              Height = 21
              AutoSize = False
              TabOrder = 4
              OnEnter = ACT_TOTPRCHARGEEnter
              OnExit = ACT_TOTPRCHARGEExit
              Decimals = 2
              Digits = 12
              Masks.PositiveMask = '#,##0.00'
              Debit = False
              UseRounding = True
              Validate = False
            end
            object act_puvente_: THNumEdit
              Left = 141
              Top = 67
              Width = 69
              Height = 21
              AutoSize = False
              TabOrder = 5
              OnEnter = ACT_PUVENTEEnter
              OnExit = ACT_PUVENTEExit
              Decimals = 2
              Digits = 12
              Masks.PositiveMask = '#,##0.00'
              Debit = False
              UseRounding = True
              Validate = False
            end
            object act_totvente_: THNumEdit
              Left = 215
              Top = 67
              Width = 81
              Height = 21
              AutoSize = False
              TabOrder = 6
              OnEnter = ACT_TOTVENTEEnter
              OnExit = ACT_TOTVENTEExit
              Decimals = 2
              Digits = 12
              Masks.PositiveMask = '#,##0.00'
              Debit = False
              UseRounding = True
              Validate = False
            end
            object act_unite_: THCritMaskEdit
              Tag = -1
              Left = 61
              Top = 45
              Width = 49
              Height = 21
              TabStop = False
              AutoSize = False
              Color = clBtnFace
              ReadOnly = True
              TabOrder = 1
              OnEnter = ACT_UNITEEnter
              OnExit = ACT_UNITEExit
              TagDispatch = 0
              OnElipsisClick = ACT_UNITEElipsisClick
            end
            object act_qte_: THNumEdit
              Left = 61
              Top = 67
              Width = 49
              Height = 21
              AutoSize = False
              TabOrder = 2
              OnEnter = ACT_QTEEnter
              OnExit = ACT_QTEExit
              Decimals = 2
              Digits = 12
              Masks.PositiveMask = '#,##0.00'
              Debit = False
              UseRounding = True
              Validate = False
            end
            object act_dateactivite_: THCritMaskEdit
              Tag = -1
              Left = 61
              Top = 2
              Width = 75
              Height = 21
              TabStop = False
              AutoSize = False
              Color = clBtnFace
              EditMask = '!99/99/0000;1;_'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = []
              MaxLength = 10
              ParentFont = False
              ReadOnly = True
              TabOrder = 7
              Text = '  /  /    '
              TagDispatch = 0
              OpeType = otDate
              ControlerDate = True
            end
            object act_activiterepris_: THCritMaskEdit
              Left = 260
              Top = 2
              Width = 36
              Height = 21
              AutoSize = False
              TabOrder = 0
              OnEnter = ACT_ACTIVITEREPRISEnter
              OnExit = ACT_ACTIVITEREPRISExit
              TagDispatch = 0
              Plus = 'AND (CO_LIBRE="ACT" OR CO_LIBRE="ART")'
              DataType = 'AFTACTIVITEREPRISE'
              ElipsisButton = True
            end
            object act_typeheure_: THCritMaskEdit
              Left = 61
              Top = 23
              Width = 49
              Height = 21
              AutoSize = False
              CharCase = ecUpperCase
              TabOrder = 8
              OnEnter = ACT_TYPEHEUREEnter
              OnExit = ACT_TYPEHEUREExit
              TagDispatch = 0
              DataType = 'AFTTYPEHEURE'
              ElipsisButton = True
            end
          end
          object HPanel3: THPanel
            Left = 0
            Top = 0
            Width = 413
            Height = 94
            Align = alClient
            BevelOuter = bvNone
            FullRepaint = False
            TabOrder = 1
            BackGroundEffect = bdFond
            ColorShadow = clWindowText
            ColorStart = clBtnFace
            ColorNb = 0
            TextEffect = tenone
            object HPanel4: THPanel
              Left = 0
              Top = 59
              Width = 413
              Height = 33
              Align = alBottom
              AutoSize = True
              BevelInner = bvRaised
              BevelOuter = bvLowered
              Constraints.MaxHeight = 33
              Constraints.MinHeight = 33
              FullRepaint = False
              TabOrder = 0
              BackGroundEffect = bdFond
              ColorShadow = clWindowText
              ColorStart = clBtnFace
              ColorNb = 2
              TextEffect = tenone
              object LibDateCreation: THLabel
                Left = 4
                Top = 2
                Width = 79
                Height = 13
                Caption = 'Date de cr'#233'ation'
                Color = clInfoBk
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = []
                ParentColor = False
                ParentFont = False
                Transparent = True
              end
              object LblDateModification: THLabel
                Left = 4
                Top = 15
                Width = 97
                Height = 13
                Caption = 'Date de modification'
                Color = clInfoBk
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = []
                ParentColor = False
                ParentFont = False
                Transparent = True
              end
              object Datecreat: THLabel
                Left = 116
                Top = 2
                Width = 8
                Height = 13
                Caption = '_'
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = [fsBold]
                ParentFont = False
                Transparent = True
              end
              object DateModif: THLabel
                Left = 116
                Top = 15
                Width = 8
                Height = 13
                Caption = '_'
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = [fsBold]
                ParentFont = False
                Transparent = True
              end
              object LblCreateur: THLabel
                Left = 202
                Top = 2
                Width = 40
                Height = 13
                Caption = 'Cr'#233'ateur'
                Color = clInfoBk
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = []
                ParentColor = False
                ParentFont = False
                Transparent = True
              end
              object LblUtilisateur: THLabel
                Left = 202
                Top = 15
                Width = 81
                Height = 13
                Caption = 'Dernier utilisateur'
                Color = clInfoBk
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = []
                ParentColor = False
                ParentFont = False
                Transparent = True
              end
              object Utilisateur: THLabel
                Left = 302
                Top = 15
                Width = 8
                Height = 13
                Caption = '_'
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = [fsBold]
                ParentFont = False
                Transparent = True
              end
              object Createur: THLabel
                Left = 302
                Top = 2
                Width = 8
                Height = 13
                Caption = '_'
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = [fsBold]
                ParentFont = False
                Transparent = True
              end
              object ACT_DATECREATION: THCritMaskEdit
                Left = 141
                Top = 6
                Width = 47
                Height = 17
                TabStop = False
                AutoSize = False
                BorderStyle = bsNone
                EditMask = '!99/99/0000;1;_'
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = []
                MaxLength = 10
                ParentColor = True
                ParentFont = False
                ReadOnly = True
                TabOrder = 0
                Text = '  /  /    '
                Visible = False
                OnChange = ACT_DATECREATIONChange
                TagDispatch = 0
                Operateur = Egal
                OpeType = otDate
                ControlerDate = True
              end
              object ACT_DATEMODIF: THCritMaskEdit
                Left = 143
                Top = 14
                Width = 40
                Height = 13
                TabStop = False
                AutoSize = False
                BorderStyle = bsNone
                EditMask = '!99/99/0000;1;_'
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = []
                MaxLength = 10
                ParentColor = True
                ParentFont = False
                ReadOnly = True
                TabOrder = 1
                Text = '  /  /    '
                Visible = False
                OnChange = ACT_DATEMODIFChange
                TagDispatch = 0
                Operateur = Egal
                OpeType = otDate
                ControlerDate = True
              end
              object ACT_CREATEUR: THCritMaskEdit
                Left = 341
                Top = 4
                Width = 15
                Height = 21
                AutoSize = False
                CharCase = ecUpperCase
                Color = clYellow
                TabOrder = 2
                Visible = False
                OnChange = ACT_CREATEURChange
                TagDispatch = 0
                DataType = 'AFTTYPEHEURE'
              end
              object ACT_UTILISATEUR: THCritMaskEdit
                Left = 368
                Top = 4
                Width = 17
                Height = 21
                AutoSize = False
                CharCase = ecUpperCase
                Color = clYellow
                TabOrder = 3
                Visible = False
                OnChange = ACT_UTILISATEURChange
                TagDispatch = 0
                DataType = 'AFTTYPEHEURE'
              end
            end
            object HPanel6: THPanel
              Left = 0
              Top = 0
              Width = 259
              Height = 59
              Align = alLeft
              AutoSize = True
              BevelInner = bvRaised
              BevelOuter = bvLowered
              Constraints.MaxWidth = 266
              Constraints.MinWidth = 155
              FullRepaint = False
              TabOrder = 1
              BackGroundEffect = bdFond
              ColorShadow = clWindowText
              ColorStart = clBtnFace
              ColorNb = 2
              TextEffect = tenone
              object LibPiecevente: THLabel
                Left = 3
                Top = 22
                Width = 120
                Height = 13
                Caption = 'Pi'#232'ce de vente associ'#233'e '
                Color = clInfoBk
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = []
                ParentColor = False
                ParentFont = False
                Transparent = True
              end
              object NUMPIECEVENTE: THLabel
                Left = 130
                Top = 22
                Width = 6
                Height = 13
                Caption = '0'
                Color = clInfoBk
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = []
                ParentColor = False
                ParentFont = False
                Transparent = True
              end
              object LibCreationFactAff: THLabel
                Left = 3
                Top = 39
                Width = 158
                Height = 13
                Caption = 'Cr'#233'ation '#233'ch'#233'ance de facturation'
                Color = clInfoBk
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = []
                ParentColor = False
                ParentFont = False
                Transparent = True
              end
              object BDATEFACTAFF: TToolbarButton97
                Left = 234
                Top = 35
                Width = 23
                Height = 22
                Hint = 'Cr'#233'ation automatique d'#39'une date de facturation'
                Color = clNone
                NumGlyphs = 2
                Opaque = False
                ShowBorderWhenInactive = True
                OnClick = BDATEFACTAFFClick
                GlobalIndexImage = 'Z0243_S16G2'
              end
              object LIBFACTURATION: THLabel
                Left = 2
                Top = 5
                Width = 69
                Height = 13
                Caption = 'Facturation '
                Color = clInfoBk
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = [fsBold]
                ParentColor = False
                ParentFont = False
                Transparent = True
              end
              object DATEFACTAFF: THCritMaskEdit
                Left = 166
                Top = 36
                Width = 67
                Height = 21
                TabStop = False
                AutoSize = False
                EditMask = '!99/99/0000;1;_'
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = []
                MaxLength = 10
                ParentFont = False
                TabOrder = 0
                Text = '  /  /    '
                TagDispatch = 0
                OpeType = otDate
                DefaultDate = od1900
                ControlerDate = True
              end
            end
            object HPanel7: THPanel
              Left = 259
              Top = 0
              Width = 154
              Height = 59
              Align = alClient
              AutoSize = True
              BevelInner = bvRaised
              BevelOuter = bvLowered
              Caption = ' '
              Constraints.MinWidth = 152
              FullRepaint = False
              TabOrder = 2
              OnEnter = PTotaux1Enter
              BackGroundEffect = bdFond
              ColorShadow = clWindowText
              ColorStart = clBtnFace
              ColorNb = 2
              TextEffect = tenone
              object LibPieceAchat: THLabel
                Left = 7
                Top = 22
                Width = 113
                Height = 13
                Caption = 'Pi'#232'ce d'#39'achat associ'#233'e '
                Color = clInfoBk
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = []
                ParentColor = False
                ParentFont = False
                Transparent = True
              end
              object NUMPIECEACHAT: THLabel
                Left = 134
                Top = 23
                Width = 6
                Height = 13
                Caption = '0'
                Color = clInfoBk
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = []
                ParentColor = False
                ParentFont = False
                Transparent = True
              end
              object PrixAchat: THLabel
                Left = 7
                Top = 39
                Width = 95
                Height = 13
                Caption = 'Prix d'#39'achat unitaire '
                Color = clInfoBk
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = []
                ParentColor = False
                ParentFont = False
                Transparent = True
              end
              object LIBACHAT: THLabel
                Left = 6
                Top = 5
                Width = 136
                Height = 13
                Caption = 'Ligne d'#39'achat associ'#233'e '
                Color = clInfoBk
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clWindowText
                Font.Height = -11
                Font.Name = 'MS Sans Serif'
                Font.Style = [fsBold]
                ParentColor = False
                ParentFont = False
                Transparent = True
              end
            end
          end
          object HPanel5: THPanel
            Left = 413
            Top = 0
            Width = 186
            Height = 94
            Align = alRight
            AutoSize = True
            BevelInner = bvRaised
            BevelOuter = bvLowered
            Caption = ' '
            Constraints.MinWidth = 100
            FullRepaint = False
            TabOrder = 2
            Visible = False
            OnEnter = PTotaux1Enter
            BackGroundEffect = bdFond
            ColorShadow = clWindowText
            ColorStart = clBtnFace
            ColorNb = 2
            TextEffect = tenone
            object LblOrigine: THLabel
              Left = 2
              Top = 10
              Width = 33
              Height = 13
              Caption = 'Origine'
              Transparent = True
            end
            object ACT_ACTORIGINE: THValComboBox
              Left = 45
              Top = 6
              Width = 139
              Height = 21
              ItemHeight = 0
              TabOrder = 0
              OnExit = ACT_ACTORIGINEExit
              TagDispatch = 0
              DataType = 'AFTACTIVITEORIGINE'
            end
          end
        end
      end
    end
  end
  object BZoomTableauBord: THBitBtn
    Tag = 100
    Left = 172
    Top = 145
    Width = 50
    Height = 27
    Hint = 'Tableau de bord Affaire'
    Caption = 'Tableau de bord Mission'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 9
    Visible = False
    OnClick = BZoomTableauBordClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object FindLigne: TFindDialog
    OnFind = FindLigneFind
    Left = 15
    Top = 267
  end
  object HTitres: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Liste des articles'
      'Articles'
      'Tiers'
      'Liste des affaires'
      'Affaires'
      'ATTENTION : Activit'#233' non enregistr'#233'e.'
      
        'ATTENTION : Cette activit'#233' en cours de traitement par un autre u' +
        'tilisateur n'#39'a pas '#233't'#233' enregistr'#233'e.'
      '7;'
      'La date saisie n'#39'est pas correcte.'
      'La date saisie n'#39'est pas dans la semaine courante.'
      'Liste des ressources'
      'Ressources'
      'Liste des unit'#233's'
      'Unit'#233's'
      'Choix Activit'#233' reprise'
      'Activit'#233' reprise'
      'Liste des Types d'#39'articles'
      'Type Article'
      
        'Il existe une incoh'#233'rence entre le client et l'#39'affaire de la lig' +
        'ne d'#39'activit'#233' courante.'
      
        'La date choisie n'#39'est pas dans l'#39'intervalle des dates de saisie ' +
        'de l'#39'activit'#233'.'
      
        'Veuillez s'#233'lectionner une affaire pour acc'#233'der '#224' la fonctionnali' +
        't'#233' demand'#233'e.'
      
        'La ligne n'#39'est pas compl'#232'te, veuillez saisir les '#233'l'#233'ments manqua' +
        'nts ou la supprimer.'
      'Liste des clients'
      
        'La date choisie n'#39'est pas dans l'#39'intervalle des dates de l'#39'affai' +
        're.'
      'Types d'#39'heures suppl'#233'mentaires'
      'Heures suppl'#233'mentaires'
      'Choix impossible. La ressource est ferm'#233'.'
      'Tri invalide dans ce contexte'
      'Voulez-vous annuler les modifications de la ligne ?'
      'L'#39'article saisi est ferm'#233', veuillez en saisir un autre.'
      
        'Attention, l'#39'action de couper supprimera les lignes. Elles ne po' +
        'urront '#234'tre r'#233'cup'#233'r'#233'es que par l'#39'action coller. Veuillez confirm' +
        'er cette action.'
      'Confirmez-vous la suppression de la ligne courante ?')
    Left = 316
    Top = 287
  end
  object POPZ: TPopupMenu
    AutoPopup = False
    Left = 58
    Top = 178
  end
  object POPY: TPopupMenu
    OnPopup = POPYPopup
    Left = 371
    Top = 221
    object Ajoutduneligne1: TMenuItem
      Caption = '&Ajouter une ligne'
      OnClick = Ajoutduneligne1Click
    end
    object Supprimeuneligne1: TMenuItem
      Caption = '&Supprimer une ligne'
      OnClick = Supprimeuneligne1Click
    end
    object Slectionneruneligne1: TMenuItem
      Caption = 'S'#233'l&ectionner une ligne'
      OnClick = Slectionneruneligne1Click
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object Copier1: TMenuItem
      Caption = '&Copier'
      OnClick = Copier1Click
    end
    object Couper1: TMenuItem
      Caption = 'Cou&per'
      OnClick = Couper1Click
    end
    object Coller1: TMenuItem
      Caption = 'Co&ller'
      Enabled = False
      OnClick = Coller1Click
    end
    object Inverserslection1: TMenuItem
      Caption = '&Inverser la s'#233'lection'
      OnClick = Inverserslection1Click
    end
    object N2: TMenuItem
      Caption = '-'
    end
    object CpltLigne: TMenuItem
      Caption = 'C&ompl'#233'ment ligne'
      OnClick = CpltLigneClick
    end
    object Afficherlemmo1: TMenuItem
      Caption = 'Afficher le &m'#233'mo'
      OnClick = Afficherlemmo1Click
    end
  end
  object HMTrad: THSystemMenu
    LockedCtrls.Strings = (
      'GCalendrier')
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 68
    Top = 268
  end
  object HActivite: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Il n'#39'existe pas de tarif pour cet article correspond' +
        'ant '#224' la devise de la pi'#232'ce !;W;O;O;O;'
      
        '1;?caption?;Saisie impossible. Il n'#39'existe pas de tarif pour cet' +
        ' article correspondant '#224' la devise de la pi'#232'ce !;W;O;O;O;'
      
        '2;?caption?;Saisie impossible. Cet article ferm'#233' n'#39'est pas autor' +
        'is'#233' pour cette nature de document;W;O;O;O;'
      
        '3;?caption?;Choix impossible. Ce  tiers ferm'#233' ne peut pas '#234'tre a' +
        'ppel'#233' pour cette nature de document;W;O;O;O;'
      
        '4;?caption?;Choix impossible. Le tiers et l'#39'affaire ne sont pas ' +
        'compatibles;W;O;O;O;'
      '5;?caption?;Vous devez renseigner une ressource valide;W;O;O;O;'
      '6;?caption?;Confirmez-vous l'#39'abandon de la saisie ?;Q;YN;Y;Y;'
      
        '7;?caption?;Voulez-vous affecter ce d'#233'p'#244't sur toutes les lignes ' +
        'concern'#233'es ?;Q;YN;Y;Y;'
      
        '8;?caption?;Voulez-vous enregistrer les modifications ?;Q;YNA;Y;' +
        'Y;'
      '9;?caption?;Vous devez renseigner une affaire valide;W;O;O;O;'
      
        '10;?caption?;Choix impossible. Ce  code tiers est incorrect;W;O;' +
        'O;O;'
      
        '11;?caption?;Choix impossible. Ce  code unit'#233' est incorrect;W;O;' +
        'O;O;'
      
        '12;?caption?;Saisie impossible. Le User courant n'#39'a pas les droi' +
        'ts ou pas de code associ'#233';W;O;O;O;'
      
        '13;?caption?;Saisie impossible. Le User courant n'#39'a pas les droi' +
        'ts;W;O;O;O;')
    Left = 298
    Top = 230
  end
end

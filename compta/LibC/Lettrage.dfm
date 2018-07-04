object FLettrage: TFLettrage
  Left = 289
  Top = 142
  Width = 647
  Height = 512
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Lettrage manuel'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  PopupMenu = POPS
  Position = poScreenCenter
  OnActivate = FormActivate
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object BValide: TToolbarButton97
    Tag = 1
    Left = 81
    Top = 0
    Width = 27
    Height = 24
    Hint = 'Valider le lettrage'
    DisplayMode = dmTextOnly
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    Margin = 2
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    Spacing = -1
    Visible = False
    OnClick = BValideClick
    GlobalIndexImage = 'Z0184_S16G1'
    IsControl = True
  end
  object BValideSelect: TToolbarButton97
    Tag = 1
    Left = 85
    Top = 4
    Width = 27
    Height = 24
    Hint = 'Pr'#233'-lettrer la s'#233'lection'
    DisplayMode = dmTextOnly
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
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
    Margin = 2
    NumGlyphs = 2
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    Spacing = -1
    Visible = False
    OnClick = BValideSelectClick
    IsControl = True
  end
  object DockTop: TDock97
    Left = 0
    Top = 0
    Width = 631
    Height = 9
  end
  object DockRight: TDock97
    Left = 622
    Top = 9
    Width = 9
    Height = 436
    Position = dpRight
  end
  object DockLeft: TDock97
    Left = 0
    Top = 9
    Width = 9
    Height = 436
    Position = dpLeft
  end
  object DockBottom: TDock97
    Left = 0
    Top = 445
    Width = 631
    Height = 28
    AllowDrag = False
    Position = dpBottom
    object Valide97: TToolbar97
      Left = 562
      Top = 0
      Caption = 'Validation'
      CloseButton = False
      DockPos = 562
      TabOrder = 0
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
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object GeneValide: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Valider le lettrage'
        Caption = 'Valider'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = GeneValideClick
        GlobalIndexImage = 'Z0184_S16G1'
        IsControl = True
      end
    end
    object POutils: TToolbar97
      Left = 0
      Top = 0
      Caption = 'Outils'
      CloseButton = False
      DockPos = 0
      TabOrder = 1
      object BAnnulerSel: TToolbarButton97
        Tag = 1
        Left = 81
        Top = 0
        Width = 27
        Height = 24
        Hint = 'D'#233'faire la s'#233'lection en cours'
        Caption = 'D'#233'faire'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ImageIndex = 1
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAnnulerSelClick
        GlobalIndexImage = 'Z0075_S16G1'
        IsControl = True
      end
      object BRajoute: TToolbarButton97
        Tag = 1
        Left = 387
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Remettre les totalement lettr'#233'es'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 0
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        OnClick = BRajouteClick
        GlobalIndexImage = 'Z0344_S16G2'
        IsControl = True
      end
      object BCombi: TToolbarButton97
        Tag = 1
        Left = 195
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Rapprochement combinatoire'
        Caption = 'Combinatoire'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 0
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BCombiClick
        GlobalIndexImage = 'Z0190_S16G2'
        IsControl = True
      end
      object BUnique: TToolbarButton97
        Tag = 1
        Left = 222
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Rapprochement simple'
        Caption = 'Unitaire'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 0
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BUniqueClick
        GlobalIndexImage = 'Z0192_S16G1'
        IsControl = True
      end
      object BChercher: TToolbarButton97
        Tag = 1
        Left = 249
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Rechercher dans la liste'
        Caption = 'Chercher'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ImageIndex = 1
        Layout = blGlyphTop
        Margin = 3
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BChercherClick
        GlobalIndexImage = 'Z0077_S16G1'
        IsControl = True
      end
      object BMenuZoom: TToolbarButton97
        Tag = -100
        Left = 282
        Top = 0
        Width = 37
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
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -100
        OnClick = BMenuZoomClick
        OnMouseEnter = BMenuZoomMouseEnter
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
      object BEnleve: TToolbarButton97
        Tag = 1
        Left = 414
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Enlever les totalement lettr'#233'es'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 0
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        OnClick = BEnleveClick
        GlobalIndexImage = 'Z0078_S16G2'
        IsControl = True
      end
      object ToolbarSep972: TToolbarSep97
        Left = 276
        Top = 0
      end
      object BAjouteEnleve: TToolbarButton97
        Tag = 1
        Left = 108
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Remettre les totalement lettr'#233'es'
        Caption = 'Basculer'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 0
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAjouteEnleveClick
        GlobalIndexImage = 'Z0344_S16G2'
        IsControl = True
      end
      object BTF: TToolbarButton97
        Left = 319
        Top = 0
        Width = 28
        Height = 24
        Hint = 'Voir les tiers factur'#233's'
        Caption = 'Factur'#233's'
        AllowAllUp = True
        GroupIndex = 100
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -16
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Margin = 3
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BTFClick
        GlobalIndexImage = 'Z0040_S16G1'
        IsControl = True
      end
      object BDelettre: TToolbarButton97
        Left = 347
        Top = 0
        Width = 40
        Height = 24
        Hint = 'D'#233'lettrage'
        Caption = 'D'#233'lettrage'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = POPDeLett
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -16
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ImageIndex = -100
        Margin = 3
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        GlobalIndexImage = 'Z1547_S16G1'
        IsControl = True
      end
      object BAgrandir: TToolbarButton97
        Tag = 1
        Left = 27
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Agrandir la liste'
        ImageIndex = 1
        ParentShowHint = False
        ShowHint = True
        OnClick = BAgrandirClick
        GlobalIndexImage = 'Z0270_S16G1'
      end
      object BReduire: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 27
        Height = 24
        Hint = 'R'#233'duire la liste'
        ImageIndex = 1
        ParentShowHint = False
        ShowHint = True
        OnClick = BReduireClick
        GlobalIndexImage = 'Z0910_S16G1'
      end
      object BCalcul: TToolbarButton97
        Tag = -101
        Left = 158
        Top = 0
        Width = 37
        Height = 24
        Hint = 'M'#233'thodes de lettrage'
        Caption = 'Combinatoire'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = PopCalcul
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ImageIndex = -100
        Layout = blGlyphTop
        Margin = 0
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnMouseEnter = BCalculMouseEnter
        GlobalIndexImage = 'Z0190_S16G2'
        IsControl = True
      end
      object BForceVadide: TToolbarButton97
        Left = 469
        Top = 0
        Width = 37
        Height = 24
        Hint = 'Action de lettrage'
        DropdownArrow = True
        DropdownMenu = PopForceValide
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z0123_S16G1'
      end
      object BTri: TToolbarButton97
        Tag = 1
        Left = 441
        Top = 0
        Width = 28
        Height = 24
        Hint = 'Trier par code lettre'
        ParentShowHint = False
        ShowHint = True
        OnClick = BForceVadideClick
        GlobalIndexImage = 'Z1157_S16G1'
      end
      object BPresentationPGE: TToolbarButton97
        Left = 54
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Mode d'#233'bit cr'#233'dit'
        ParentShowHint = False
        ShowHint = True
        OnClick = BPresentationPGEClick
        GlobalIndexImage = 'Z0038_S16G1'
      end
      object BPresentationPCL: TToolbarButton97
        Left = 135
        Top = 0
        Width = 23
        Height = 24
        Hint = 'Mode solde progressif'
        ParentShowHint = False
        ShowHint = True
        OnClick = BPresentationPCLClick
        GlobalIndexImage = 'Z0285_S16G1'
      end
      object BTriDate: TToolbarButton97
        Tag = 1
        Left = 534
        Top = 0
        Width = 28
        Height = 24
        Hint = 'Trier par date comptable'
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BForceVadideClick
        GlobalIndexImage = 'Z1536_S16G1'
      end
      object BTriCode: TToolbarButton97
        Tag = 1
        Left = 506
        Top = 0
        Width = 28
        Height = 24
        Hint = 'Trier par code lettre'
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BForceVadideClick
        GlobalIndexImage = 'Z1157_S16G1'
      end
    end
  end
  object PFEN: TPanel
    Left = 9
    Top = 9
    Width = 613
    Height = 436
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object GC: THGrid
      Tag = 1
      Left = 313
      Top = 93
      Width = 300
      Height = 154
      Align = alRight
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 11
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
      ScrollBars = ssVertical
      TabOrder = 1
      OnDblClick = ClickSelect
      OnEnter = ClickActive
      OnMouseUp = GDMouseUp
      SortedCol = -1
      Couleur = True
      MultiSelect = False
      TitleBold = True
      TitleCenter = True
      OnRowEnter = MontreDetail
      OnCellEnter = SauteZero
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = clSilver
    end
    object GD: THGrid
      Tag = 1
      Left = 0
      Top = 93
      Width = 313
      Height = 154
      Align = alClient
      Ctl3D = True
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 11
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
      ParentCtl3D = False
      ParentFont = False
      ScrollBars = ssVertical
      TabOrder = 0
      OnDblClick = ClickSelect
      OnEnter = ClickActive
      OnMouseUp = GDMouseUp
      SortedCol = -1
      Couleur = True
      MultiSelect = False
      TitleBold = True
      TitleCenter = True
      OnRowEnter = MontreDetail
      OnCellEnter = SauteZero
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = clSilver
      ColWidths = (
        64
        64
        64
        64
        64)
    end
    object CJOURNAL: THValComboBox
      Left = 16
      Top = 92
      Width = 24
      Height = 21
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 2
      Visible = False
      TagDispatch = 0
      DataType = 'TTJOURNAUX'
    end
    object CNATUREPIECE: THValComboBox
      Left = 40
      Top = 92
      Width = 24
      Height = 21
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 3
      Visible = False
      TagDispatch = 0
      DataType = 'TTNATUREPIECE'
    end
    object CMODEPAIE: THValComboBox
      Left = 64
      Top = 92
      Width = 24
      Height = 21
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 4
      Visible = False
      TagDispatch = 0
      DataType = 'TTMODEPAIE'
    end
    object CDevise: THValComboBox
      Left = 88
      Top = 92
      Width = 24
      Height = 21
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 5
      Visible = False
      TagDispatch = 0
      DisableTab = True
    end
    object BZoomEcriture: THBitBtn
      Tag = 100
      Left = 104
      Top = 1
      Width = 22
      Height = 22
      Hint = 'Voir l'#39#233'criture'
      Caption = 'Piec'
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
      OnClick = BZoomEcritureClick
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object BZoomCpte: THBitBtn
      Tag = 100
      Left = 128
      Top = 1
      Width = 22
      Height = 22
      Hint = 'Voir Compte'
      Caption = 'Cpte'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 8
      Visible = False
      OnClick = BZoomCpteClick
      Margin = 2
      Spacing = -1
      IsControl = True
    end
    object BChancel: THBitBtn
      Tag = 100
      Left = 152
      Top = 1
      Width = 22
      Height = 22
      Hint = 'Table de chancellerie'
      Caption = 'Dev'
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
      OnClick = BChancelClick
      Margin = 4
      Spacing = -1
      IsControl = True
    end
    object BJustif: THBitBtn
      Tag = 100
      Left = 176
      Top = 1
      Width = 22
      Height = 22
      Hint = 'D'#233'tail des '#233'ch'#233'ances'
      Caption = 'Just.'
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
      OnClick = BJustifClick
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object PCouverture: TPanel
      Left = 120
      Top = 112
      Width = 497
      Height = 108
      BevelInner = bvRaised
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 6
      Visible = False
      object H_TitreCouv: TLabel
        Left = 2
        Top = 13
        Width = 493
        Height = 13
        Align = alTop
        Alignment = taCenter
        AutoSize = False
        Caption = 'Saisie des montants de couverture'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
      end
      object H_TOTCOUVDD: TLabel
        Left = 8
        Top = 78
        Width = 106
        Height = 13
        Caption = 'Total couverture D'#233'bit'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object H_TOTCOUVDC: TLabel
        Left = 220
        Top = 78
        Width = 27
        Height = 13
        Caption = 'Cr'#233'dit'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object BCValide: THBitBtn
        Tag = 1
        Left = 425
        Top = 74
        Width = 28
        Height = 27
        Hint = 'Valider la couverture'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
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
        Left = 460
        Top = 74
        Width = 28
        Height = 27
        Hint = 'Abandonner la saisie de couverture'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = BCAbandonClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object E_TOTCOUVDEBS: THNumEdit
        Tag = 1
        Left = 120
        Top = 74
        Width = 93
        Height = 21
        Color = clBtnFace
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 3
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = True
        UseRounding = True
        Validate = False
      end
      object E_TOTCOUVCREDS: THNumEdit
        Tag = 1
        Left = 252
        Top = 74
        Width = 93
        Height = 21
        Color = clBtnFace
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 4
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object PFenCouverture: TPanel
        Left = 2
        Top = 2
        Width = 493
        Height = 11
        Align = alTop
        BevelInner = bvLowered
        BevelOuter = bvNone
        Color = clActiveCaption
        Enabled = False
        TabOrder = 5
      end
      object E_TAUXDEV: THNumEdit
        Tag = 1
        Left = 387
        Top = 80
        Width = 21
        Height = 16
        AutoSize = False
        Color = clYellow
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 6
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = True
        UseRounding = True
        Validate = False
      end
      object GCouvDev: TGroupBox
        Left = 2
        Top = 26
        Width = 493
        Height = 40
        Align = alTop
        Color = clBtnFace
        ParentColor = False
        TabOrder = 0
        object H_MontantD: TLabel
          Left = 4
          Top = 15
          Width = 39
          Height = 13
          Caption = 'Montant'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object H_COUVERTURED: TLabel
          Left = 176
          Top = 15
          Width = 52
          Height = 13
          Caption = 'Couverture'
          FocusControl = E_COUVERTURES
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object H_RESTED: TLabel
          Left = 344
          Top = 15
          Width = 28
          Height = 13
          Caption = 'Reste'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object E_MONTANTS: THNumEdit
          Tag = 1
          Left = 48
          Top = 12
          Width = 121
          Height = 21
          Color = clBtnFace
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 0
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object E_COUVERTURES: THNumEdit
          Tag = 1
          Left = 232
          Top = 12
          Width = 109
          Height = 21
          Color = clWhite
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 1
          OnChange = E_COUVERTURESChange
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object E_RESTES: THNumEdit
          Tag = 1
          Left = 380
          Top = 12
          Width = 101
          Height = 21
          Color = clBtnFace
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 2
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0'
          Debit = False
          UseRounding = True
          Validate = False
        end
      end
    end
    object Pages: TPageControl
      Left = 0
      Top = 0
      Width = 613
      Height = 93
      ActivePage = TComptes
      Align = alTop
      TabOrder = 11
      object TComptes: TTabSheet
        Caption = 'Comptes'
        object H_CODEL: TLabel
          Left = 473
          Top = 10
          Width = 39
          Height = 13
          Alignment = taCenter
          Caption = 'Lettrage'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object ISigneEuro: TImage
          Left = 550
          Top = 8
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
        object TG_GENERAL: THLabel
          Left = 9
          Top = 10
          Width = 37
          Height = 13
          Caption = 'G'#233'n'#233'ral'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TT_AUXILIAIRE: THLabel
          Left = 8
          Top = 37
          Width = 41
          Height = 13
          Caption = 'Auxiliaire'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HG_GENERAL: THLabel
          Left = 88
          Top = 10
          Width = 239
          Height = 13
          AutoSize = False
          Caption = 'Compte g'#233'n'#233'ral'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HT_AUXILIAIRE: THLabel
          Left = 88
          Top = 37
          Width = 239
          Height = 13
          AutoSize = False
          Caption = 'Compte auxiliaire'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TSoldes: TLabel
          Left = 354
          Top = 37
          Width = 104
          Height = 13
          Caption = 'Solde s'#233'lection (pivot)'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BvSolde: TBevel
          Left = 482
          Top = 32
          Width = 26
          Height = 23
        end
        object HTitreL: TLabel
          Left = 354
          Top = 10
          Width = 51
          Height = 13
          Caption = 'Code lettre'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BDown: TToolbarButton97
          Left = 56
          Top = 32
          Width = 23
          Height = 22
          Opaque = False
          OnClick = BDownClick
          GlobalIndexImage = 'Z2234_S16G1'
        end
        object BUP: TToolbarButton97
          Left = 56
          Top = 8
          Width = 23
          Height = 22
          Opaque = False
          OnClick = BUPClick
          GlobalIndexImage = 'Z1512_S16G1'
        end
        object LESOLDE: THNumEdit
          Tag = 2
          Left = 467
          Top = 35
          Width = 99
          Height = 17
          BorderStyle = bsNone
          Color = clBtnFace
          Ctl3D = True
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0'
          Debit = False
          UseRounding = True
          Validate = False
        end
      end
      object TParams: TTabSheet
        Caption = 'Param'#232'tres'
        ImageIndex = 1
        object H_REGUL: THLabel
          Left = 2
          Top = 36
          Width = 45
          Height = 13
          Caption = 'Jal. &r'#233'gul.'
          FocusControl = E_REGUL
          OnDblClick = H_REGULDblClick
        end
        object H_ECART: THLabel
          Left = 186
          Top = 36
          Width = 82
          Height = 13
          Caption = 'Jal. '#233'cart &change'
          FocusControl = E_ECART
          OnDblClick = H_ECARTDblClick
        end
        object TE_REFINTERNE: THLabel
          Left = 329
          Top = 6
          Width = 58
          Height = 13
          Caption = 'R'#233'f. &lettrage'
          FocusControl = E_REFLETTRAGE
        end
        object TDATEGENERATION: THLabel
          Left = 408
          Top = 36
          Width = 76
          Height = 13
          Caption = 'Date g'#233'n'#233'ration'
        end
        object CParEnsemble: TCheckBox
          Left = 4
          Top = 5
          Width = 131
          Height = 17
          Alignment = taLeftJustify
          Caption = 'D'#233'lettrage &global'
          Checked = True
          State = cbChecked
          TabOrder = 1
          OnClick = CParEnsembleClick
        end
        object CParCouverture: TCheckBox
          Left = 8
          Top = 4
          Width = 131
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Lettrage par &couverture'
          TabOrder = 0
          OnClick = CParCouvertureClick
        end
        object E_REGUL: THValComboBox
          Left = 54
          Top = 32
          Width = 125
          Height = 21
          Style = csDropDownList
          ItemHeight = 0
          TabOrder = 4
          OnChange = E_REGULChange
          TagDispatch = 0
          Vide = True
          VideString = '<< Aucun >>'
          DataType = 'TTJALREGUL'
        end
        object E_ECART: THValComboBox
          Left = 274
          Top = 32
          Width = 125
          Height = 21
          Style = csDropDownList
          ItemHeight = 0
          TabOrder = 5
          OnChange = E_ECARTChange
          TagDispatch = 0
          Vide = True
          VideString = '<< Aucun >>'
          DataType = 'TTJALECART'
        end
        object CParExcept: TCheckBox
          Left = 160
          Top = 4
          Width = 131
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Lettrage par e&xception'
          TabOrder = 2
          OnClick = CParCouvertureClick
        end
        object E_REFLETTRAGE: TEdit
          Left = 401
          Top = 2
          Width = 170
          Height = 21
          CharCase = ecUpperCase
          Ctl3D = True
          MaxLength = 35
          ParentCtl3D = False
          TabOrder = 3
        end
        object DATEGENERATION: THCritMaskEdit
          Left = 490
          Top = 31
          Width = 80
          Height = 21
          EditMask = '!99/99/0000;1;_'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 10
          ParentFont = False
          TabOrder = 6
          Text = '  /  /    '
          OnExit = DATEGENERATIONExit
          TagDispatch = 0
          Operateur = Egal
          OpeType = otDate
          ElipsisButton = True
          ElipsisAutoHide = True
          ControlerDate = True
        end
      end
    end
    object PCDetail: TPageControl
      Left = 0
      Top = 335
      Width = 613
      Height = 101
      ActivePage = TSDetail
      Align = alBottom
      TabOrder = 12
      object TSDetail: TTabSheet
        Caption = 'D'#233'tail'
        object PDetail: TPanel
          Left = 0
          Top = 0
          Width = 605
          Height = 73
          Align = alClient
          BevelOuter = bvNone
          Enabled = False
          TabOrder = 0
          object BSepare: TBevel
            Left = 300
            Top = 1
            Width = 2
            Height = 89
            Shape = bsLeftLine
            Visible = False
          end
          object H_JOURNAL: TLabel
            Left = 4
            Top = 4
            Width = 34
            Height = 13
            AutoSize = False
            Caption = 'Journal'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clNavy
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object E_JOURNAL: TLabel
            Tag = 1
            Left = 56
            Top = 4
            Width = 96
            Height = 13
            AutoSize = False
            Caption = 'E_JOURNAL'
          end
          object H_NATUREPIECE: TLabel
            Left = 156
            Top = 4
            Width = 32
            Height = 13
            AutoSize = False
            Caption = 'Nature'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clNavy
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object E_NATUREPIECE: TLabel
            Tag = 1
            Left = 200
            Top = 4
            Width = 96
            Height = 13
            AutoSize = False
            Caption = 'E_NATUREPIECE'
          end
          object H_DATEECHEANCE: TLabel
            Left = 4
            Top = 20
            Width = 49
            Height = 13
            AutoSize = False
            Caption = 'Ech'#233'ance'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clNavy
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object E_DATEECHEANCE: TLabel
            Tag = 1
            Left = 56
            Top = 20
            Width = 96
            Height = 13
            AutoSize = False
            Caption = 'E_DATEECHEANCE'
          end
          object H_MODEPAIE: TLabel
            Left = 156
            Top = 36
            Width = 24
            Height = 13
            AutoSize = False
            Caption = 'Paie.'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clNavy
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object E_MODEPAIE: TLabel
            Tag = 1
            Left = 200
            Top = 36
            Width = 96
            Height = 13
            Alignment = taCenter
            AutoSize = False
            Caption = 'E_MODEPAIE'
          end
          object H_LIBELLE: TLabel
            Left = 156
            Top = 20
            Width = 30
            Height = 13
            AutoSize = False
            Caption = 'Libell'#233
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clNavy
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object E_LIBELLE: TLabel
            Tag = 1
            Left = 200
            Top = 20
            Width = 96
            Height = 13
            AutoSize = False
            Caption = 'E_LIBELLE'
          end
          object H_REFLIBRE: TLabel
            Left = 4
            Top = 36
            Width = 42
            Height = 13
            AutoSize = False
            Caption = 'Relev'#233
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clNavy
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object E_REFRELEVE: TLabel
            Tag = 1
            Left = 56
            Top = 36
            Width = 96
            Height = 13
            AutoSize = False
            Caption = 'E_REFRELEVE'
          end
          object H_TOTAL: TLabel
            Left = 1
            Top = 52
            Width = 43
            Height = 13
            AutoSize = False
            Caption = 'Euros'
            FocusControl = E_TOTAL
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clNavy
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object H_JOURNAL_: TLabel
            Left = 304
            Top = 4
            Width = 34
            Height = 13
            AutoSize = False
            Caption = 'Journal'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clMaroon
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object E_JOURNAL_: TLabel
            Tag = 1
            Left = 356
            Top = 4
            Width = 96
            Height = 13
            AutoSize = False
            Caption = 'E_JOURNAL'
          end
          object H_NATUREPIECE_: TLabel
            Left = 456
            Top = 4
            Width = 32
            Height = 13
            AutoSize = False
            Caption = 'Nature'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clMaroon
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object E_NATUREPIECE_: TLabel
            Tag = 1
            Left = 500
            Top = 4
            Width = 95
            Height = 13
            AutoSize = False
            Caption = 'E_NATUREPIECE'
          end
          object H_LIBELLE_: TLabel
            Left = 456
            Top = 20
            Width = 30
            Height = 13
            AutoSize = False
            Caption = 'Libell'#233
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clMaroon
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object E_LIBELLE_: TLabel
            Tag = 1
            Left = 500
            Top = 20
            Width = 96
            Height = 13
            AutoSize = False
            Caption = 'E_LIBELLE'
          end
          object H_DATEECHEANCE_: TLabel
            Left = 304
            Top = 20
            Width = 49
            Height = 13
            AutoSize = False
            Caption = 'Ech'#233'ance'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clMaroon
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object E_DATEECHEANCE_: TLabel
            Tag = 1
            Left = 356
            Top = 20
            Width = 96
            Height = 13
            AutoSize = False
            Caption = 'E_DATEECHEANCE'
          end
          object H_MODEPAIE_: TLabel
            Left = 456
            Top = 36
            Width = 24
            Height = 13
            AutoSize = False
            Caption = 'Paie.'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clMaroon
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object E_MODEPAIE_: TLabel
            Tag = 1
            Left = 500
            Top = 36
            Width = 96
            Height = 13
            AutoSize = False
            Caption = 'E_MODEPAIE'
          end
          object H_REFLIBRE_: TLabel
            Left = 304
            Top = 36
            Width = 42
            Height = 13
            AutoSize = False
            Caption = 'Relev'#233
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clMaroon
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object E_REFRELEVE_: TLabel
            Tag = 1
            Left = 356
            Top = 36
            Width = 96
            Height = 13
            AutoSize = False
            Caption = 'E_REFRELEVE'
          end
          object H_TOTAL_: TLabel
            Left = 306
            Top = 53
            Width = 43
            Height = 13
            AutoSize = False
            Caption = 'Euros'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clMaroon
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object E_TOTAL: THNumEdit
            Tag = 1
            Left = 55
            Top = 52
            Width = 97
            Height = 13
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
            TabOrder = 0
            Decimals = 2
            Digits = 12
            Masks.PositiveMask = '#,##0'
            Debit = False
            UseRounding = True
            Validate = False
          end
          object E_TOTAL_: THNumEdit
            Tag = 1
            Left = 355
            Top = 52
            Width = 97
            Height = 15
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
            TabOrder = 1
            Decimals = 2
            Digits = 12
            Masks.PositiveMask = '#,##0'
            Debit = False
            UseRounding = True
            Validate = False
          end
        end
      end
      object TSDetail1: TTabSheet
        Caption = 'D'#233'tail'
        ImageIndex = 1
        object Label1: TLabel
          Left = 3
          Top = 4
          Width = 34
          Height = 13
          AutoSize = False
          Caption = 'Nature'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object e_refrelevepcl: TLabel
          Tag = 1
          Left = 48
          Top = 19
          Width = 123
          Height = 13
          AutoSize = False
          Caption = 'E_JOURNAL'
        end
        object e_naturepcl: TLabel
          Tag = 1
          Left = 48
          Top = 4
          Width = 128
          Height = 13
          AutoSize = False
          Caption = 'E_JOURNAL'
        end
        object Label3: TLabel
          Left = 3
          Top = 19
          Width = 42
          Height = 13
          AutoSize = False
          Caption = 'Relev'#233
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label7: TLabel
          Left = 3
          Top = 35
          Width = 73
          Height = 14
          AutoSize = False
          Caption = 'Affaire'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object e_affairepcl: TLabel
          Tag = 1
          Left = 48
          Top = 35
          Width = 128
          Height = 13
          AutoSize = False
          Caption = 'E_JOURNAL'
        end
        object Label8: TLabel
          Left = 184
          Top = 35
          Width = 97
          Height = 14
          AutoSize = False
          Caption = 'R'#233'f'#233'rence externe'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label4: TLabel
          Left = 184
          Top = 19
          Width = 89
          Height = 13
          AutoSize = False
          Caption = 'R'#233'f'#233'rence libre'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label2: TLabel
          Left = 376
          Top = 19
          Width = 49
          Height = 13
          AutoSize = False
          Caption = 'Ech'#233'ance'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object e_dateecheancepcl: TLabel
          Tag = 1
          Left = 496
          Top = 19
          Width = 105
          Height = 13
          AutoSize = False
          Caption = 'E_JOURNAL'
        end
        object e_reflibrepcl: TLabel
          Tag = 1
          Left = 280
          Top = 19
          Width = 81
          Height = 13
          AutoSize = False
          Caption = 'E_JOURNAL'
        end
        object e_refexternepcl: TLabel
          Tag = 1
          Left = 280
          Top = 35
          Width = 89
          Height = 13
          AutoSize = False
          Caption = 'E_JOURNAL'
        end
        object Label6: TLabel
          Left = 184
          Top = 4
          Width = 121
          Height = 13
          AutoSize = False
          Caption = 'Etablissement'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label5: TLabel
          Left = 376
          Top = 4
          Width = 121
          Height = 13
          AutoSize = False
          Caption = 'Compte de contrepartie'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object E_ETABLISSEMENTPCL: TLabel
          Tag = 1
          Left = 280
          Top = 4
          Width = 75
          Height = 13
          AutoSize = False
          Caption = 'E_JOURNAL'
        end
        object E_CONTREPARTIEGENPCL: TLabel
          Tag = 1
          Left = 496
          Top = 4
          Width = 113
          Height = 13
          AutoSize = False
          Caption = 'E_JOURNAL'
        end
        object Label11: TLabel
          Left = 376
          Top = 36
          Width = 24
          Height = 13
          AutoSize = False
          Caption = 'Paie.'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object e_modepaiepcl: TLabel
          Tag = 1
          Left = 496
          Top = 36
          Width = 96
          Height = 13
          AutoSize = False
          Caption = 'E_MODEPAIE'
        end
      end
    end
    object PCSoldes: TPageControl
      Left = 0
      Top = 247
      Width = 613
      Height = 88
      ActivePage = TSSoldePCL
      Align = alBottom
      TabOrder = 13
      object TSSoldePME: TTabSheet
        Caption = 'D'#233'tail'
        object H_INFODEBIT: THLabel
          Left = 115
          Top = 8
          Width = 69
          Height = 13
          Caption = 'Solde      Total'
        end
        object H_INFOCREDIT: THLabel
          Left = 416
          Top = 8
          Width = 69
          Height = 13
          Caption = 'Total      Solde'
        end
        object SL_SOLDEDEBIT: THNumEdit
          Tag = 1
          Left = 3
          Top = 4
          Width = 109
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
          UseRounding = True
          Validate = False
        end
        object SL_TOTALDEBIT: THNumEdit
          Tag = 1
          Left = 188
          Top = 4
          Width = 109
          Height = 21
          Color = clBtnFace
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 1
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object SL_TOTALCREDIT: THNumEdit
          Tag = 1
          Left = 304
          Top = 4
          Width = 109
          Height = 21
          Color = clBtnFace
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 2
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0'
          Debit = False
          UseRounding = False
          Validate = False
        end
        object SL_SOLDECREDIT: THNumEdit
          Tag = 1
          Left = 492
          Top = 4
          Width = 109
          Height = 21
          Color = clBtnFace
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 3
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0'
          Debit = False
          UseRounding = True
          Validate = False
        end
      end
      object TSSoldePCL: TTabSheet
        Caption = 'D'#233'tail'
        ImageIndex = 1
        object h1: THLabel
          Left = 96
          Top = 8
          Width = 74
          Height = 13
          Alignment = taRightJustify
          Caption = 'Cumul s'#233'lection'
        end
        object h2: THLabel
          Left = 88
          Top = 32
          Width = 96
          Height = 13
          Alignment = taRightJustify
          Caption = 'Solde cumul lettrage'
        end
        object SL_SOLDEDEBITPCL: THNumEdit
          Left = 187
          Top = 28
          Width = 109
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
          NumericType = ntDecimal
          UseRounding = True
          Validate = False
        end
        object SL_CUMULDEBIT: THNumEdit
          Tag = 1
          Left = 188
          Top = 4
          Width = 109
          Height = 21
          Color = clBtnFace
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 1
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0'
          Debit = False
          NumericType = ntDecimal
          UseRounding = True
          Validate = False
        end
        object SL_CUMULCREDIT: THNumEdit
          Tag = 1
          Left = 304
          Top = 4
          Width = 109
          Height = 21
          Color = clBtnFace
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 2
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0'
          Debit = False
          NumericType = ntDecimal
          UseRounding = False
          Validate = False
        end
        object SL_SOLDECREDITPCL: THNumEdit
          Left = 308
          Top = 28
          Width = 109
          Height = 21
          Color = clBtnFace
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 3
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0'
          Debit = False
          NumericType = ntDecimal
          UseRounding = True
          Validate = False
        end
        object SL_SOLDE: THNumEdit
          Left = 420
          Top = 28
          Width = 109
          Height = 21
          Color = clBtnFace
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 4
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0'
          Debit = False
          NumericType = ntDecimal
          UseRounding = True
          Validate = False
        end
        object SL_SOLDEPCL: THNumEdit
          Tag = 1
          Left = 419
          Top = 4
          Width = 109
          Height = 21
          Color = clBtnFace
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 5
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0'
          Debit = False
          NumericType = ntDecimal
          UseRounding = False
          Validate = False
        end
      end
    end
    object PMessLettAuto: TPanel
      Left = 133
      Top = 80
      Width = 389
      Height = 142
      BevelInner = bvRaised
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 14
      Visible = False
      object Label9: TLabel
        Left = 4
        Top = 28
        Width = 375
        Height = 18
        Alignment = taCenter
        AutoSize = False
        Caption = 'Lettrage automatique en cours. Veuillez patienter...'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
      end
      object EBStop: TToolbarButton97
        Left = 154
        Top = 104
        Width = 93
        Height = 27
        Hint = 'Lancer l'#39'importation'
        Caption = '  A&rr'#234'ter'
        Opaque = False
        ParentShowHint = False
        ShowHint = False
        OnClick = EBStopClick
        GlobalIndexImage = 'Z0107_S16G1'
      end
      object Panel2: TPanel
        Left = 2
        Top = 2
        Width = 385
        Height = 11
        Align = alTop
        BevelInner = bvLowered
        BevelOuter = bvNone
        Color = clActiveCaption
        Enabled = False
        TabOrder = 1
      end
      object GroupBox1: TGroupBox
        Left = 4
        Top = 53
        Width = 379
        Height = 40
        Color = clBtnFace
        ParentColor = False
        TabOrder = 0
        object Label12: TLabel
          Left = 4
          Top = 15
          Width = 87
          Height = 13
          Caption = 'Solutions trouv'#233'es'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label13: TLabel
          Left = 171
          Top = 15
          Width = 34
          Height = 13
          Caption = 'Niveau'
          FocusControl = HNBN
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object THNBP: TLabel
          Left = 275
          Top = 15
          Width = 52
          Height = 13
          Caption = 'Profondeur'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HNBS: THNumEdit
          Tag = 1
          Left = 98
          Top = 12
          Width = 29
          Height = 21
          Color = clBtnFace
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 0
          Decimals = 0
          Digits = 12
          Masks.PositiveMask = '#,##0'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object HNBN: THNumEdit
          Tag = 1
          Left = 213
          Top = 12
          Width = 29
          Height = 21
          Color = clBtnFace
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 1
          Decimals = 0
          Digits = 12
          Masks.PositiveMask = '#,##0'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object HNBP: THNumEdit
          Tag = 1
          Left = 344
          Top = 12
          Width = 29
          Height = 21
          Color = clBtnFace
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 2
          Decimals = 0
          Digits = 12
          Masks.PositiveMask = '#,##0'
          Debit = False
          UseRounding = True
          Validate = False
        end
      end
    end
  end
  object HLettre: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Lettrage manuel;Le param'#233'trage des listes est incorrect;W;O;O;' +
        'O;'
      
        '1;Lettrage manuel;Lettrage '#233'quilibr'#233' : voulez-vous pr'#233'-lettrer l' +
        'a s'#233'lection ?;Q;YN;Y;Y;'
      
        '2;Lettrage manuel;Lettrage d'#233's'#233'quilibr'#233' en devise pivot : Voulez' +
        '-vous g'#233'n'#233'rer l'#39#233'criture d'#39#233'cart de change ?;Q;YN;Y;Y;'
      
        '3;Lettrage manuel;Vous abandonnez un traitement de lettrage en c' +
        'ours. Voulez-vous valider ce lettrage ?;Q;YNC;Y;Y;'
      
        '4;Lettrage manuel;Traitement de pr'#233'-lettrage en cours. Confirmez' +
        '-vous l'#39'abandon de la fonction ?;Q;YNC;O;O;'
      
        '5;Lettrage manuel;Le mouvement est en devise. Voulez-vous lettre' +
        'r en '#39';Q;YN;O;O;'
      
        '6;Lettrage manuel;La r'#233'gularisation que vous avez demand'#233'e est i' +
        'mpossible : le param'#233'trage des comptes de r'#233'gularisation est inc' +
        'omplet;E;O;O;O;'
      
        '7;Lettrage manuel;La r'#233'gularisation que vous avez demand'#233'e est i' +
        'mpossible : les valeurs de seuil pour les d'#233'bits et cr'#233'dits sont' +
        ' nulles;E;O;O;O;'
      
        '8;Lettrage manuel;S'#233'lection impossible : il existe des pi'#232'ces de' +
        ' r'#233'gularisation g'#233'n'#233'r'#233'es pour cette '#233'criture;E;O;O;O;'
      
        '9;Lettrage manuel;La couverture et le montant d'#39'origine doivent ' +
        #234'tre de m'#234'me sens;E;O;O;O;'
      
        '10;Lettrage manuel;La couverture ne peut pas d'#233'passer le montant' +
        ' d'#39'origine;E;O;O;O;'
      
        '11;Lettrage manuel;La somme des couvertures d'#233'bit doit '#233'galer ce' +
        'lle des couvertures cr'#233'dit;E;O;O;O;'
      
        '12;Lettrage manuel;Confirmez-vous l'#39#233'cart entre la couverture pi' +
        'vot saisie et celle calcul'#233'e?;Q;YN;O;O;'
      
        '13;Lettrage manuel;Il n'#39'existe aucun '#233'l'#233'ment '#224' lettrer pour ces ' +
        'crit'#232'res;E;O;O;O;'
      
        '14;Lettrage manuel;Le paquet est d'#233'ja totalement pr'#233'-lettr'#233';E;O;' +
        'O;O;'
      
        '15;Lettrage manuel;Vous devez '#234'tre sur un '#233'l'#233'ment '#224' rapprocher;E' +
        ';O;O;O;'
      
        '16;Lettrage manuel;L'#39#233'l'#233'ment s'#233'lectionn'#233' est d'#233'j'#224' totalement let' +
        'tr'#233';E;O;O;O;'
      
        '17;Lettrage manuel;Rapprochement non effectu'#233' : aucune solution ' +
        'n'#39'a '#233't'#233' trouv'#233'e;E;O;O;O;'
      
        '18;Lettrage manuel;ATTENTION! Vous n'#39'avez pas renseign'#233' le taux ' +
        'de change de la devise '#224' la date du jour;E;O;O;O;'
      
        '19;Lettrage manuel;ATTENTION! Lettrage partiel d'#39'un relev'#233'. Conf' +
        'irmez-vous le lettrage partiel de votre '#233'ch'#233'ance de relev'#233' ?;Q;Y' +
        'NC;Y;Y;'
      
        '20;Lettrage manuel;Visualisation impossible : le mouvement n'#39'est' +
        ' pas lettr'#233';E;O;O;O;'
      
        '21;Lettrage manuel;Lettrage impossible : le compte est interdit;' +
        'E;O;O;O;'
      
        '22;D'#233'lettrage;Vous abandonnez un traitement de d'#233'lettrage en cou' +
        'rs. Voulez-vous valider ce d'#233'lettrage ?;Q;YNC;Y;Y;'
      
        '23;D'#233'lettrage;Traitement de pr'#233'-d'#233'lettrage en cours. Confirmez-v' +
        'ous l'#39'abandon de la fonction ?;Q;YNC;O;O;'
      
        '24;Lettrage manuel;Vous devez s'#233'lectionner au moins un '#233'l'#233'ment '#224 +
        ' rapprocher;E;O;O;O;'
      
        '25;Lettrage manuel;La r'#233'gularisation que vous avez demand'#233' est i' +
        'mpossible : les montants pivot sont nuls;E;O;O;O;'
      '26;?caption?;ATTENTION : Le tiers est ferm'#233';E;O;O;O;'
      
        '27;D'#233'lettrage;Il n'#39'existe aucun '#233'l'#233'ment '#224' d'#233'lettrer pour ces cri' +
        't'#232'res;E;O;O;O;'
      
        '28;?caption?;La r'#233'gularisation est impossible : le journal ne po' +
        'ss'#232'de pas de facturier;E;O;O;O;'
      
        '29;?caption?;Voulez-vous g'#233'n'#233'rer un '#233'cart de conversion ?;Q;YN;O' +
        ';O;'
      
        '30;D'#233'lettrage;ATTENTION! Certaines '#233'ch'#233'ances ont fait l'#39'objet d'#39 +
        'une '#233'dition de Tva. Confirmez-vous le d'#233'lettrage ? ;Q;YNC;Y;Y;'
      
        '31;Lettrage manuel;Confirmez-vous le lettrage partiel de vos '#233'ch' +
        #233'ances ?;Q;YN;Y;Y;'
      
        '32;Lettrage manuel;Le lettrage va g'#233'n'#233'rer des '#233'critures de r'#233'gul' +
        'arisation et/ou d'#39#233'cart. Confirmez-vous l'#39'op'#233'ration ?;Q;YN;Y;Y;'
      
        '33;Lettrage manuel;Les '#233'l'#233'ments retenus par exception ne constit' +
        'uent pas un lettrage '#233'quilibr'#233';E;O;O;O;'
      
        '34;Lettrage manuel;Voulez-vous voir les '#233'critures g'#233'n'#233'r'#233'es ?;Q;Y' +
        'N;Y;Y;'
      
        '35;Lettrage manuel;Confirmez-vous le d'#233'lettrage de tous les '#233'l'#233'm' +
        'ents visibles et charg'#233's du compte ?;Q;YN;Y;Y;'
      
        '36;Lettrage manuel;Confirmez-vous le d'#233'lettrage du compte pour l' +
        'e code ;Q;YN;Y;Y;'
      '37;?caption?;La date de r'#233'gularisation est incorrecte;E;O;O;O;'
      
        '38;?caption?;ATTENTION : Certaines '#233'ch'#233'ances ont une date sup'#233'ri' +
        'eure au 31/12/2001. Confirmez-vous ??;Q;YN;O;O;'
      ''
      ' '
      ' '
      ' ')
    Left = 56
    Top = 195
  end
  object POPS: TPopupMenu
    OnPopup = POPSPopup
    Left = 100
    Top = 147
  end
  object FindLettre: TFindDialog
    OnFind = FindLettreFind
    Left = 56
    Top = 139
  end
  object HDivL: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Lettrage manuel'
      'Lettrage en saisie'
      'D'#233'lettrage'
      'R'#233'gularisation'
      'R'#233'gularisation de lettrage'
      'Ecart de change'
      'Ecart de change'
      'Lettrage/D'#233'lettrage non valid'#233
      
        'Validation impossible! Certaines '#233'ch'#233'ances ont '#233't'#233' modifi'#233'es par' +
        ' un autre utilisateur'
      'Consultation du lettrage'
      'Valider le d'#233'lettrage'
      'Pr'#233'parer le d'#233'lettrage'
      
        'ATTENTION ! Ce tiers est d'#233'j'#224' en cours de lettrage ! Ressortez d' +
        'e la fonction.'
      'ATTENTION : Mauvaise attribution du code lettre !'
      '14;'
      'Ecart de conversion'
      'Ecart de conversion'
      'Paie.'
      'Fact.'
      ' '
      ' ')
    Left = 12
    Top = 195
  end
  object POPZ: TPopupMenu
    Left = 8
    Top = 140
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 96
    Top = 196
  end
  object PopCalcul: TPopupMenu
    Left = 209
    Top = 225
    object Lettragesurpassagedesolde1: TMenuItem
      Caption = 'Lettrage sur passage de solde'
      ShortCut = 32817
      OnClick = Lettragesurpassagedesolde1Click
    end
    object Rapprochementcombinatoire1: TMenuItem
      Caption = 'Rapprochement combinatoire'
      ShortCut = 32818
      OnClick = Rapprochementcombinatoire1Click
    end
    object Rapprochementsimple1: TMenuItem
      Caption = 'Rapprochement sur les montants'
      ShortCut = 32819
      OnClick = Rapprochementsimple1Click
    end
    object Rapprochementsurrfrence1: TMenuItem
      Caption = 'Rapprochement sur r'#233'f'#233'rence'
      ShortCut = 32820
      OnClick = Rapprochementsurrfrence1Click
    end
    object Lettragesurpassagesoldenul1: TMenuItem
      Caption = 'Lettrage sur passage '#224' solde nul'
      ShortCut = 32821
      OnClick = Lettragesurpassagesoldenul1Click
    end
  end
  object POPDeLett: TPopupMenu
    Left = 169
    Top = 225
    object DeLettPartiel: TMenuItem
      Caption = 'Annulation du lettrage selectionn'#233
      OnClick = DeLettPartielClick
    end
    object DELettTotal: TMenuItem
      Caption = 'Annulation totale du lettrage de ce compte'
      OnClick = DELettTotalClick
    end
  end
  object PopForceValide: TPopupMenu
    Left = 249
    Top = 225
    object Lettragepartiel1: TMenuItem
      Caption = 'Lettrage partiel'
      OnClick = Lettragepartiel1Click
    end
    object Passagedunecrituresimplifi1: TMenuItem
      Caption = 'Passage d'#39'une '#233'criture simplifi'#233'e'
      OnClick = Passagedunecrituresimplifi1Click
    end
    object Soldeautomatique1: TMenuItem
      Caption = 'Solde automatique'
      OnClick = Soldeautomatique1Click
    end
    object Lettrageparexception1: TMenuItem
      Caption = 'Lettrage par exception'
      OnClick = Lettrageparexception1Click
    end
  end
  object POPPCL: TPopupMenu
    OnPopup = POPPCLPopup
    Left = 309
    Top = 225
    object Validerlelettrage1: TMenuItem
      Caption = 'Valider le lettrage'
      Enabled = False
      OnClick = Validerlelettrage1Click
    end
    object Prlettrerlaslection1: TMenuItem
      Caption = 'Pr'#233'-lettrer la s'#233'lection'
      OnClick = Validerlelettrage1Click
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object Lettragepartiel2: TMenuItem
      Caption = 'Lettrage partiel'
      OnClick = Lettragepartiel1Click
    end
    object Passagedunecrituresimplifie1: TMenuItem
      Caption = 'Passage d'#39'une '#233'criture simplifi'#233'e'
      OnClick = Passagedunecrituresimplifi1Click
    end
    object Soldeautomatique2: TMenuItem
      Caption = 'Solde automatique'
      OnClick = Soldeautomatique1Click
    end
    object LEX: TMenuItem
      Caption = 'Lettrage par exception'
      OnClick = Lettrageparexception1Click
    end
    object N2: TMenuItem
      Caption = '-'
    end
    object Dfairelaselectionencours1: TMenuItem
      Caption = 'D'#233'faire la selection en cours'
      ShortCut = 16474
      OnClick = BAnnulerSelClick
    end
    object Rechercherdanslaliste1: TMenuItem
      Caption = 'Rechercher dans la liste'
      ShortCut = 16454
      OnClick = BChercherClick
    end
    object Recherchedanslacomptabilit1: TMenuItem
      Caption = 'Recherche de mouvements comptables'
      OnClick = Recherchedanslacomptabilit1Click
    end
    object Enleverlestotalementslettrs1: TMenuItem
      Caption = 'Enlever les totalements lettr'#233'es'
      OnClick = BEnleveClick
    end
    object Remettrelestotalementslettres1: TMenuItem
      Caption = 'Remettre les totalements lettr'#233'es'
      Enabled = False
      OnClick = BRajouteClick
    end
    object Agrandirlaliste1: TMenuItem
      Caption = 'Agrandir la liste'
      OnClick = BAgrandirClick
    end
    object Rduirelaliste1: TMenuItem
      Caption = 'R'#233'duire la liste'
      Enabled = False
      OnClick = BReduireClick
    end
    object POPAnnulSelect: TMenuItem
      Caption = 'Annulation du lettrage s'#233'lectionn'#233
      OnClick = DeLettPartielClick
    end
    object POPAnnulTotal: TMenuItem
      Caption = 'Annulation totale du lettrage du compte'
      Enabled = False
      OnClick = DELettTotalClick
    end
    object Zoom1: TMenuItem
      Caption = 'M'#233'thodes de lettrage'
      object Lettragesurpassagedesolde2: TMenuItem
        Caption = 'Lettrage sur passage de solde'
        ShortCut = 32817
        OnClick = Lettragesurpassagedesolde1Click
      end
      object Rapprochementcombinatoire2: TMenuItem
        Caption = 'Rapprochement combinatoire'
        ShortCut = 32818
        OnClick = Rapprochementcombinatoire1Click
      end
      object Rapprochementsurlesmontants1: TMenuItem
        Caption = 'Rapprochement sur les montants'
        ShortCut = 32819
        OnClick = Rapprochementsurlesmontants1Click
      end
      object Rapprochementssurlesrfrences1: TMenuItem
        Caption = 'Rapprochements sur les r'#233'f'#233'rences'
        ShortCut = 32820
        OnClick = Rapprochementsurrfrence1Click
      end
      object Lettragesurpassagedesoldenul1: TMenuItem
        Caption = 'Lettrage sur passage de solde '#224' nul'
        ShortCut = 32821
        OnClick = Lettragesurpassagesoldenul1Click
      end
    end
    object Zoom2: TMenuItem
      Caption = 'Zoom'
      object Voirlcriture1: TMenuItem
        Caption = 'Voir l'#39#233'criture'
        OnClick = BZoomEcritureClick
      end
      object Voirlecompte1: TMenuItem
        Caption = 'Voir le compte'
        OnClick = BZoomCpteClick
      end
      object Tabledechancellerie1: TMenuItem
        Caption = 'Table de chancellerie'
        OnClick = BChancelClick
      end
      object Dtailsdeschances1: TMenuItem
        Caption = 'D'#233'tails des '#233'ch'#233'ances'
        OnClick = BJustifClick
      end
    end
    object POPAUX: TMenuItem
      Caption = 'Passer au compte ....'
      object POPAuxSuiv: TMenuItem
        Caption = 'Suivant'
        OnClick = POPAUXSuivClick
      end
      object POPAUXPREC: TMenuItem
        Caption = 'Pr'#233'c'#233'dent'
        OnClick = POPAUXPRECClick
      end
    end
  end
end

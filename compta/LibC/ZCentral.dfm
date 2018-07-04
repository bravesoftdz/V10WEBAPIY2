object FCentral: TFCentral
  Left = 305
  Top = 194
  Width = 702
  Height = 345
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Journal Centralisateur'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clBlack
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  ShowHint = True
  OnClose = FormClose
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object DockTop: TDock97
    Left = 0
    Top = 0
    Width = 694
    Height = 9
  end
  object DockRight: TDock97
    Left = 685
    Top = 9
    Width = 9
    Height = 277
    Position = dpRight
  end
  object DockLeft: TDock97
    Left = 0
    Top = 9
    Width = 9
    Height = 277
    Position = dpLeft
  end
  object DockBottom: TDock97
    Left = 0
    Top = 286
    Width = 694
    Height = 32
    Position = dpBottom
    object Valide97: TToolbar97
      Left = 522
      Top = 0
      Caption = 'Validation'
      CloseButton = False
      DockPos = 522
      TabOrder = 0
      object BValide: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 27
        Height = 27
        Hint = 'Valider la saisie'
        DisplayMode = dmGlyphOnly
        Caption = 'Valider'
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
        GlobalIndexImage = 'Z0184_S16G1'
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 27
        Top = 0
        Width = 27
        Height = 27
        Hint = 'Fermer'
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
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
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
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
      Left = 48
      Top = 0
      Caption = 'Outils'
      DockPos = 48
      TabOrder = 1
      object BCherche: TToolbarButton97
        Left = 0
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Appliquer crit'#232'res'
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BChercheClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object BInit: TToolbarButton97
        Left = 28
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Niveau sup'#233'rieur'
        NumGlyphs = 2
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BInitClick
        GlobalIndexImage = 'Z2300_S16G2'
      end
      object BNivPrec: TToolbarButton97
        Left = 56
        Top = 0
        Width = 28
        Height = 28
        Hint = 'Niveau pr'#233'c'#233'dent'
        NumGlyphs = 2
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BNivPrecClick
        GlobalIndexImage = 'Z0077_S16G2'
      end
      object BStop: TToolbarButton97
        Left = 84
        Top = 0
        Width = 28
        Height = 28
        Hint = 'Arr'#234'ter la s'#233'lection'
        DisplayMode = dmGlyphOnly
        Caption = 'Arr'#234'ter l'#39#233'dition'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BStopClick
        GlobalIndexImage = 'Z0107_S16G1'
        IsControl = True
      end
      object BTree: TToolbarButton97
        Left = 112
        Top = 0
        Width = 32
        Height = 28
        Hint = 'Modifier l'#39'affichage'
        DropdownArrow = True
        DropdownMenu = POPB
        Opaque = False
        GlobalIndexImage = 'M0004_S16G1'
      end
    end
  end
  object PFEN: TPanel
    Left = 9
    Top = 9
    Width = 676
    Height = 277
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object Splitter1: TSplitter
      Left = 0
      Top = 29
      Height = 248
    end
    object Splitter2: TSplitter
      Left = 206
      Top = 29
      Height = 248
      OnMoved = Splitter2Moved
    end
    object PEntete: THPanel
      Left = 0
      Top = 0
      Width = 676
      Height = 29
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
      object HLabel3: THLabel
        Left = 9
        Top = 7
        Width = 41
        Height = 13
        Caption = 'Exercice'
        FocusControl = EXO
        Transparent = True
      end
      object HLabel2: THLabel
        Left = 326
        Top = 8
        Width = 36
        Height = 13
        Caption = 'P'#233'riode'
        FocusControl = PERIODE
        Transparent = True
        Visible = False
      end
      object HLabel1: THLabel
        Left = 434
        Top = 8
        Width = 43
        Height = 13
        Caption = 'Journaux'
        FocusControl = JOURNAUX
        Transparent = True
        Visible = False
      end
      object BThread: TToolbarButton97
        Left = 286
        Top = 2
        Width = 27
        Height = 22
        Visible = False
        GlobalIndexImage = 'Z0050_S16G1'
      end
      object EXO: THValComboBox
        Left = 57
        Top = 3
        Width = 226
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = EXOChange
        TagDispatch = 0
        DataType = 'TTEXERCICE'
      end
      object PERIODE: THValComboBox
        Left = 370
        Top = 4
        Width = 56
        Height = 21
        Style = csDropDownList
        Color = clYellow
        ItemHeight = 13
        TabOrder = 1
        Visible = False
        TagDispatch = 0
        Vide = True
      end
      object JOURNAUX: THValComboBox
        Left = 484
        Top = 4
        Width = 96
        Height = 21
        Style = csDropDownList
        Color = clYellow
        ItemHeight = 13
        TabOrder = 2
        Visible = False
        TagDispatch = 0
        Vide = True
        DataType = 'TTJOURNAUX'
      end
    end
    object HPanel1: THPanel
      Left = 3
      Top = 29
      Width = 203
      Height = 248
      Align = alLeft
      Caption = 'HPanel1'
      FullRepaint = False
      TabOrder = 1
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
      object TreeView: TTreeView
        Left = 1
        Top = 1
        Width = 201
        Height = 246
        Align = alClient
        HideSelection = False
        Images = ImageList
        Indent = 21
        ReadOnly = True
        RowSelect = True
        StateImages = ImageList
        TabOrder = 0
        OnChange = TreeViewChange
        OnExpanding = TreeViewExpanding
      end
    end
    object HPanel3: THPanel
      Left = 209
      Top = 29
      Width = 467
      Height = 248
      Align = alClient
      Caption = 'HPanel3'
      FullRepaint = False
      TabOrder = 2
      BackGroundEffect = bdFlat
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      TextEffect = tenone
      object GS: THGrid
        Left = 1
        Top = 1
        Width = 465
        Height = 246
        Align = alClient
        BorderStyle = bsNone
        ColCount = 6
        Ctl3D = True
        DefaultRowHeight = 18
        FixedCols = 2
        RowCount = 100
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goTabs, goRowSelect]
        ParentCtl3D = False
        TabOrder = 0
        OnDblClick = GSDblClick
        SortedCol = -1
        Titres.Strings = (
          'CodeJal'
          'TypeJal'
          'Intitul'#233
          'D'#233'bit'
          'Cr'#233'dit'
          'Mouvements'
          '')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = 13224395
        ColWidths = (
          2
          2
          169
          80
          78
          18)
        RowHeights = (
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18
          18)
      end
    end
  end
  object HCentral: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '<<Tous>>'
      'Bordereau'
      'Pi'#232'ce'
      'Aucun enregistrement ne correspond aux crit'#232'res'
      'Total'
      'Pas de mouvement'
      'Visualisation des '#233'critures'
      'Modification des '#233'critures'
      'courantes'
      'de simulation'
      'de pr'#233'vision'
      'de situation'
      'de r'#233'vision'
      'd'#39#224'-nouveaux'
      
        '14;?caption?;Attention ce traitement peut '#234'tre long#10#13       ' +
        '   Voulez-vous continuer ?;Q;YN;Y;Y;'
      'Total g'#233'n'#233'ral')
    Left = 24
    Top = 229
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 68
    Top = 230
  end
  object ImageList: THImageList
    GlobalIndexImages.Strings = (
      'Z0499_S16G1'
      'Z0127_S16G1'
      'Z1576_S16G1'
      'Z0652_S16G1'
      'Z0206_S16G1'
      'Z1959_S16G1')
    Height = 18
    Width = 18
    Left = 113
    Top = 229
  end
  object POPB: TPopupMenu
    Left = 22
    Top = 44
    object MenuParPeriode: TMenuItem
      Caption = 'Par p'#233'riode'
      Checked = True
      OnClick = MenuParPeriodeClick
    end
    object MenuParJournal: TMenuItem
      Caption = 'Par journal'
      OnClick = MenuParJournalClick
    end
  end
end

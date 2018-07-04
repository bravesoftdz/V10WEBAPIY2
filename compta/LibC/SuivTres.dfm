object FSuiTre: TFSuiTre
  Left = 74
  Top = 160
  Width = 630
  Height = 380
  HelpContext = 7625000
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Analyse de tr'#233'sorerie'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object BGL: TToolbarButton97
    Tag = 100
    Left = 169
    Top = 219
    Width = 25
    Height = 22
    Hint = 'Grand-livre simple'
    Caption = 'GL'
    Enabled = False
    Visible = False
    OnClick = FListeDblClick
  end
  object BMP: TToolbarButton97
    Tag = 100
    Left = 136
    Top = 219
    Width = 25
    Height = 22
    Hint = 'Mode de paiement'
    Caption = 'MP'
    Visible = False
    OnClick = BMPClick
  end
  object BGene: TToolbarButton97
    Tag = 100
    Left = 201
    Top = 219
    Width = 25
    Height = 22
    Hint = 'Compte g'#233'n'#233'ral'
    Caption = 'Gene'
    Visible = False
    OnClick = BGeneClick
  end
  object BFont: TToolbarButton97
    Tag = 100
    Left = 234
    Top = 217
    Width = 33
    Height = 27
    Hint = 'Param'#233'trage de la police'
    Caption = 'Font'
    Visible = False
    OnClick = BFontClick
  end
  object FListe: THGrid
    Left = 0
    Top = 160
    Width = 622
    Height = 157
    Align = alClient
    DefaultRowHeight = 18
    DefaultDrawing = False
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goRowSelect]
    TabOrder = 1
    OnDblClick = FListeDblClick
    OnDrawCell = FListeDrawCell
    SortedCol = -1
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
      125
      120
      120
      120
      120)
  end
  object Dock: TDock97
    Left = 0
    Top = 317
    Width = 622
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 622
      Caption = 'Barre d'#39'outils'
      ClientAreaHeight = 32
      ClientAreaWidth = 622
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BReduire: TToolbarButton97
        Left = 4
        Top = 3
        Width = 28
        Height = 27
        Hint = 'R'#233'duire la liste'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BReduireClick
        GlobalIndexImage = 'Z0910_S16G1'
      end
      object BAgrandir: TToolbarButton97
        Left = 4
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Agrandir la liste'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAgrandirClick
        GlobalIndexImage = 'Z0270_S16G1'
      end
      object BRechercher: TToolbarButton97
        Left = 36
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Rechercher dans la liste'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BRechercherClick
        GlobalIndexImage = 'Z0077_S16G1'
      end
      object BMenuZoom: TToolbarButton97
        Tag = -100
        Left = 67
        Top = 3
        Width = 37
        Height = 27
        Hint = 'Menu zoom'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = PopZ
        Flat = False
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
        OnMouseEnter = BMenuZoomMouseEnter
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
      object BImprimer: TToolbarButton97
        Left = 523
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BAnnuler: TToolbarButton97
        Left = 555
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Flat = False
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 587
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
    end
  end
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 622
    Height = 121
    ActivePage = PCritere
    Align = alTop
    TabOrder = 0
    TabWidth = 85
    object PCritere: TTabSheet
      Caption = 'Standards'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 614
        Height = 93
        Align = alClient
      end
      object TCpte: THLabel
        Left = 312
        Top = 15
        Width = 36
        Height = 13
        Caption = '&Compte'
        FocusControl = Cpte
      end
      object TQualPie: THLabel
        Left = 8
        Top = 43
        Width = 60
        Height = 13
        Caption = '&Type de Mvt'
        FocusControl = QualPie
      end
      object TDev: THLabel
        Left = 312
        Top = 71
        Width = 33
        Height = 13
        Caption = '&Devise'
        FocusControl = Dev
      end
      object TEtab: THLabel
        Left = 8
        Top = 71
        Width = 65
        Height = 13
        Caption = '&Etablissement'
        FocusControl = Etab
      end
      object TFPaie: THLabel
        Left = 312
        Top = 43
        Width = 88
        Height = 13
        Caption = '&Mode de paiement'
        FocusControl = FPaie
      end
      object Cpte: TEdit
        Left = 413
        Top = 11
        Width = 186
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 2
        OnKeyPress = CpteKeyPress
      end
      object Cpte1: THCpteEdit
        Left = 413
        Top = 11
        Width = 186
        Height = 21
        TabOrder = 1
        ZoomTable = tzGbanque
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object QualPie: THValComboBox
        Left = 86
        Top = 39
        Width = 186
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        TagDispatch = 0
        Vide = True
        DataType = 'TTQUALPIECECRIT'
      end
      object Dev: THValComboBox
        Left = 413
        Top = 67
        Width = 186
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 6
        OnChange = DevChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTDEVISE'
      end
      object Etab: THValComboBox
        Left = 86
        Top = 67
        Width = 186
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 5
        TagDispatch = 0
        Vide = True
        DataType = 'TTETABLISSEMENT'
      end
      object FPaie: THMultiValComboBox
        Left = 413
        Top = 39
        Width = 186
        Height = 21
        TabOrder = 4
        Style = csDialog
        Abrege = False
        DataType = 'TTMODEPAIE'
        Complete = False
        OuInclusif = False
      end
      object CbChoix: TRadioGroup
        Left = 8
        Top = 1
        Width = 263
        Height = 34
        Caption = ' Choix sur compte '
        Columns = 2
        ItemIndex = 0
        Items.Strings = (
          '&Mono Compte'
          'Multi &Compte')
        TabOrder = 0
        OnClick = CbChoixClick
      end
    end
    object PPeriode: TTabSheet
      Caption = 'P'#233'riodes'
      object Bevel2: TBevel
        Left = 0
        Top = 0
        Width = 614
        Height = 93
        Align = alClient
      end
      object Tcol: TLabel
        Left = 264
        Top = 25
        Width = 95
        Height = 13
        Caption = '&Nombre de p'#233'riodes'
      end
      object TPerio: TLabel
        Left = 264
        Top = 60
        Width = 49
        Height = 13
        Caption = 'P'#233'rio&dicit'#233
        FocusControl = Perio
      end
      object LDat: TLabel
        Left = 428
        Top = 25
        Width = 71
        Height = 13
        Caption = '&Jusqu'#39#224' la date'
      end
      object BPerio: TToolbarButton97
        Left = 573
        Top = 55
        Width = 22
        Height = 23
        Hint = 'Param'#233'trage des dates par p'#233'riode'
        ParentShowHint = False
        ShowHint = True
        OnClick = BPerioClick
        GlobalIndexImage = 'Z0584_S16G1'
      end
      object NbCol: TSpinEdit
        Left = 365
        Top = 21
        Width = 42
        Height = 22
        EditorEnabled = False
        MaxLength = 2
        MaxValue = 12
        MinValue = 4
        TabOrder = 1
        Value = 4
        OnChange = NbColChange
      end
      object Perio: THValComboBox
        Left = 365
        Top = 56
        Width = 204
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 4
        OnChange = PerioChange
        TagDispatch = 0
        DataType = 'TTSEPAREECHE'
      end
      object RgAnalyse: TRadioGroup
        Left = 12
        Top = 7
        Width = 220
        Height = 38
        Caption = ' Type d'#39'analyse '
        Columns = 2
        ItemIndex = 0
        Items.Strings = (
          '&P'#233'riode'
          '&Jour')
        TabOrder = 0
        OnClick = RgAnalyseClick
      end
      object RgDat: TRadioGroup
        Left = 12
        Top = 47
        Width = 220
        Height = 38
        Caption = ' Analyser par '
        Columns = 2
        ItemIndex = 0
        Items.Strings = (
          'Date &comptable'
          'Date &valeur')
        TabOrder = 3
        OnClick = RgDatClick
      end
      object MEDat: TMaskEdit
        Left = 508
        Top = 21
        Width = 78
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 2
        Text = '31/12/2999'
      end
    end
    object PAffichage: TTabSheet
      Caption = 'Affichage'
      object Bevel4: TBevel
        Left = 0
        Top = 0
        Width = 614
        Height = 93
        Align = alClient
      end
      object RgMp: TRadioGroup
        Left = 4
        Top = 11
        Width = 486
        Height = 51
        Caption = ' Mode de paiement '
        Columns = 3
        ItemIndex = 0
        Items.Strings = (
          '&Sans les modes de paiement'
          '&Avec les modes de paiement'
          '&Que les modes de paiement')
        TabOrder = 0
      end
      object CAfdev: TRadioGroup
        Left = 494
        Top = 11
        Width = 116
        Height = 51
        Caption = ' Affichage en '
        Columns = 2
        ItemIndex = 0
        Items.Strings = (
          '&Pivot'
          '&Devise')
        TabOrder = 1
      end
      object CbJouvre: TCheckBox
        Left = 4
        Top = 72
        Width = 117
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Voir jour non ouvr'#233
        Checked = True
        State = cbChecked
        TabOrder = 2
      end
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 121
    Width = 622
    Height = 39
    AllowDrag = False
    object PFiltres: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 35
      ClientWidth = 622
      ClientAreaHeight = 35
      ClientAreaWidth = 622
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BFiltre: TToolbarButton97
        Left = 5
        Top = 5
        Width = 56
        Height = 21
        Hint = 'Menu filtre'
        DropdownArrow = True
        DropdownMenu = POPF
        Caption = '&Filtres'
        Layout = blGlyphRight
        ParentShowHint = False
        ShowHint = True
      end
      object BCherche: TToolbarButton97
        Left = 586
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Appliquer crit'#232'res'
        ParentShowHint = False
        ShowHint = True
        OnClick = BChercheClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object FFiltres: THValComboBox
        Left = 67
        Top = 5
        Width = 513
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = FFiltresChange
        TagDispatch = 0
      end
    end
  end
  object CbMp: THValComboBox
    Left = 40
    Top = 220
    Width = 81
    Height = 21
    Style = csDropDownList
    Color = clYellow
    ItemHeight = 13
    TabOrder = 5
    Visible = False
    TagDispatch = 0
    DataType = 'TTMODEPAIE'
  end
  object Pdate: TPanel
    Left = 31
    Top = 36
    Width = 533
    Height = 248
    BorderStyle = bsSingle
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
    object TFP1: TLabel
      Left = 6
      Top = 31
      Width = 19
      Height = 13
      AutoSize = False
      Caption = 'Du'
      FocusControl = FP1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFP2: TLabel
      Left = 134
      Top = 31
      Width = 15
      Height = 13
      AutoSize = False
      Caption = 'au'
      FocusControl = FP2
    end
    object TFP3: TLabel
      Left = 6
      Top = 57
      Width = 19
      Height = 14
      AutoSize = False
      Caption = 'Du'
      FocusControl = FP1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFP4: TLabel
      Left = 134
      Top = 57
      Width = 15
      Height = 14
      AutoSize = False
      Caption = 'au'
      FocusControl = FP2
    end
    object TFP5: TLabel
      Left = 6
      Top = 83
      Width = 19
      Height = 13
      AutoSize = False
      Caption = 'Du'
      FocusControl = FP1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFP6: TLabel
      Left = 134
      Top = 83
      Width = 15
      Height = 13
      AutoSize = False
      Caption = 'au'
      FocusControl = FP2
    end
    object TFP7: TLabel
      Left = 6
      Top = 109
      Width = 19
      Height = 13
      AutoSize = False
      Caption = 'Du'
      FocusControl = FP1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFP8: TLabel
      Left = 134
      Top = 109
      Width = 15
      Height = 13
      AutoSize = False
      Caption = 'au'
      FocusControl = FP2
    end
    object TFP9: TLabel
      Left = 6
      Top = 135
      Width = 19
      Height = 13
      AutoSize = False
      Caption = 'Du'
      FocusControl = FP1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFP10: TLabel
      Left = 134
      Top = 135
      Width = 15
      Height = 13
      AutoSize = False
      Caption = 'au'
      FocusControl = FP2
    end
    object TFP11: TLabel
      Left = 6
      Top = 161
      Width = 19
      Height = 13
      AutoSize = False
      Caption = 'Du'
      FocusControl = FP1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFP12: TLabel
      Left = 134
      Top = 161
      Width = 15
      Height = 13
      AutoSize = False
      Caption = 'au'
      FocusControl = FP2
    end
    object TFP13: TLabel
      Left = 274
      Top = 30
      Width = 19
      Height = 13
      AutoSize = False
      Caption = 'Du'
      FocusControl = FP1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFP14: TLabel
      Left = 402
      Top = 30
      Width = 15
      Height = 13
      AutoSize = False
      Caption = 'au'
      FocusControl = FP2
    end
    object TFP15: TLabel
      Left = 274
      Top = 56
      Width = 19
      Height = 13
      AutoSize = False
      Caption = 'Du'
      FocusControl = FP1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFP16: TLabel
      Left = 402
      Top = 56
      Width = 15
      Height = 13
      AutoSize = False
      Caption = 'au'
      FocusControl = FP2
    end
    object TFP17: TLabel
      Left = 274
      Top = 82
      Width = 19
      Height = 13
      AutoSize = False
      Caption = 'Du'
      FocusControl = FP1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFP18: TLabel
      Left = 402
      Top = 82
      Width = 15
      Height = 13
      AutoSize = False
      Caption = 'au'
      FocusControl = FP2
    end
    object TFP19: TLabel
      Left = 274
      Top = 108
      Width = 19
      Height = 13
      AutoSize = False
      Caption = 'Du'
      FocusControl = FP1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFP20: TLabel
      Left = 402
      Top = 108
      Width = 15
      Height = 13
      AutoSize = False
      Caption = 'au'
      FocusControl = FP2
    end
    object TFP21: TLabel
      Left = 274
      Top = 134
      Width = 19
      Height = 13
      AutoSize = False
      Caption = 'Du'
      FocusControl = FP1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFP22: TLabel
      Left = 402
      Top = 134
      Width = 15
      Height = 13
      AutoSize = False
      Caption = 'au'
      FocusControl = FP2
    end
    object TFP23: TLabel
      Left = 274
      Top = 160
      Width = 19
      Height = 13
      AutoSize = False
      Caption = 'Du'
      FocusControl = FP1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFP24: TLabel
      Left = 402
      Top = 160
      Width = 15
      Height = 13
      AutoSize = False
      Caption = 'au'
      FocusControl = FP2
    end
    object Bevel3: TBevel
      Left = 265
      Top = 24
      Width = 2
      Height = 158
    end
    object TPerio1: TLabel
      Left = 149
      Top = 187
      Width = 49
      Height = 13
      Caption = '&P'#233'riodicit'#233
      FocusControl = Perio1
    end
    object Ptop: TPanel
      Left = 1
      Top = 1
      Width = 527
      Height = 22
      Align = alTop
      Alignment = taLeftJustify
      Caption = ' Param'#233'trage des dates'
      Color = clActiveCaption
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clCaptionText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
    end
    object Pbouton1: TPanel
      Left = 1
      Top = 208
      Width = 527
      Height = 35
      Align = alBottom
      BevelInner = bvLowered
      TabOrder = 26
      object BVal1: TToolbarButton97
        Left = 432
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Valider'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BVal1Click
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object BFer1: TToolbarButton97
        Left = 464
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BFer1Click
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object Baide1: TToolbarButton97
        Left = 496
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Aide'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = Baide1Click
        GlobalIndexImage = 'Z1117_S16G1'
      end
    end
    object FP1: TMaskEdit
      Left = 29
      Top = 27
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 1
      Text = '01/01/1900'
    end
    object FP2: TMaskEdit
      Left = 153
      Top = 27
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 2
      Text = '31/12/2999'
    end
    object FP3: TMaskEdit
      Left = 29
      Top = 53
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 3
      Text = '01/01/1900'
    end
    object FP4: TMaskEdit
      Left = 153
      Top = 53
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 4
      Text = '31/12/2999'
    end
    object FP5: TMaskEdit
      Left = 29
      Top = 79
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 5
      Text = '01/01/1900'
    end
    object FP6: TMaskEdit
      Left = 153
      Top = 79
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 6
      Text = '31/12/2999'
    end
    object FP7: TMaskEdit
      Left = 29
      Top = 105
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 7
      Text = '01/01/1900'
    end
    object FP8: TMaskEdit
      Left = 153
      Top = 105
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 8
      Text = '31/12/2999'
    end
    object FP9: TMaskEdit
      Left = 29
      Top = 131
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 9
      Text = '01/01/1900'
    end
    object FP10: TMaskEdit
      Left = 153
      Top = 131
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 10
      Text = '31/12/2999'
    end
    object FP11: TMaskEdit
      Left = 29
      Top = 157
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 11
      Text = '01/01/1900'
    end
    object FP12: TMaskEdit
      Left = 153
      Top = 157
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 12
      Text = '31/12/2999'
    end
    object FP13: TMaskEdit
      Left = 297
      Top = 26
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 13
      Text = '01/01/1900'
    end
    object FP14: TMaskEdit
      Left = 421
      Top = 26
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 14
      Text = '31/12/2999'
    end
    object FP15: TMaskEdit
      Left = 297
      Top = 52
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 15
      Text = '01/01/1900'
    end
    object FP16: TMaskEdit
      Left = 421
      Top = 52
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 16
      Text = '31/12/2999'
    end
    object FP17: TMaskEdit
      Left = 297
      Top = 78
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 17
      Text = '01/01/1900'
    end
    object FP18: TMaskEdit
      Left = 421
      Top = 78
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 18
      Text = '31/12/2999'
    end
    object FP19: TMaskEdit
      Left = 297
      Top = 104
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 19
      Text = '01/01/1900'
    end
    object FP20: TMaskEdit
      Left = 421
      Top = 104
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 20
      Text = '31/12/2999'
    end
    object FP21: TMaskEdit
      Left = 297
      Top = 130
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 21
      Text = '01/01/1900'
    end
    object FP22: TMaskEdit
      Left = 421
      Top = 130
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 22
      Text = '31/12/2999'
    end
    object FP23: TMaskEdit
      Left = 297
      Top = 156
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 23
      Text = '01/01/1900'
    end
    object FP24: TMaskEdit
      Left = 421
      Top = 156
      Width = 100
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 24
      Text = '31/12/2999'
    end
    object Perio1: THValComboBox
      Left = 211
      Top = 183
      Width = 173
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 25
      OnChange = PerioChange
      TagDispatch = 0
      DataType = 'TTSEPAREECHE'
    end
  end
  object FindDialog: TFindDialog
    Left = 251
    Top = 256
  end
  object QCpte: THQuery
    AutoCalcFields = False
    LockType = ltReadOnly
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    Parameters = <>
    SQL.Strings = (
      'Select G_GENERAL From GENERAUX Where'
      '(G_NATUREGENE="BQE" OR G_NATUREGENE="CAI") ')
    dataBaseName = 'SOC'
    MAJAuto = False
    Distinct = False
    Left = 32
    Top = 256
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Suivi de tr'#233'sorerie;Vous devez renseigner le crit'#232're sur les c' +
        'omptes.;W;O;O;O;'
      '&Compte'
      '&Comptes'
      'Compte tr'#233'so :'
      'D'#233'tail compte :'
      'Compte tr'#233'so :'
      '6;'
      'au'
      'Total des recettes'
      'Total des d'#233'penses'
      'Aucun'
      'Total compte :'
      '&Nombre de p'#233'riodes'
      '&Nombre de jours'
      'du'
      '15;Suivi de tr'#233'sorerie;Vous devez choisir une devise.;W;O;O;O;'
      '16;')
    Left = 361
    Top = 256
  end
  object QSom: THQuery
    AutoCalcFields = False
    LockType = ltReadOnly
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    Parameters = <>
    dataBaseName = 'SOC'
    MAJAuto = False
    Distinct = False
    Left = 87
    Top = 256
  end
  object QEcr: THQuery
    AutoCalcFields = False
    LockType = ltReadOnly
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    Parameters = <>
    dataBaseName = 'SOC'
    MAJAuto = False
    Distinct = False
    Left = 142
    Top = 256
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 306
    Top = 256
  end
  object PopZ: TPopupMenu
    Left = 578
    Top = 256
  end
  object ParamFont: TFontDialog
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Device = fdBoth
    Options = [fdEffects, fdForceFontExist]
    Left = 196
    Top = 256
  end
  object POPF: TPopupMenu
    OnPopup = POPFPopup
    Left = 416
    Top = 258
    object BCreerFiltre: TMenuItem
      Caption = '&Cr'#233'er un filtre'
      OnClick = BCreerFiltreClick
    end
    object BSaveFiltre: TMenuItem
      Caption = '&Enregistrer le filtre'
      OnClick = BSaveFiltreClick
    end
    object BDelFiltre: TMenuItem
      Caption = '&Supprimer le filtre'
      OnClick = BDelFiltreClick
    end
    object BRenFiltre: TMenuItem
      Caption = '&Renommer le filtre'
      OnClick = BRenFiltreClick
    end
    object BNouvRech: TMenuItem
      Caption = '&Nouvelle recherche'
      OnClick = BNouvRechClick
    end
  end
end

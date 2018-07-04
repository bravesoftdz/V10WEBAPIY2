object FModTlEcr: TFModTlEcr
  Left = 307
  Top = 172
  Width = 569
  Height = 384
  Caption = 'Modification des tables libres des '#233'critures comptables'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 561
    Height = 111
    ActivePage = PCritere
    Align = alTop
    TabOrder = 0
    TabWidth = 85
    object PCritere: TTabSheet
      Caption = 'Standards'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 553
        Height = 83
        Align = alClient
      end
      object TE_JOURNAL: THLabel
        Left = 295
        Top = 10
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = JOURNAL
      end
      object TE_EXERCICE: THLabel
        Left = 13
        Top = 35
        Width = 41
        Height = 13
        Caption = '&Exercice'
        FocusControl = EXERCICE
      end
      object TE_DATECOMPTABLE: THLabel
        Left = 13
        Top = 60
        Width = 43
        Height = 13
        Caption = '&Dates du'
        FocusControl = DATECOMPTABLE
      end
      object TE_DATECOMPTABLE2: THLabel
        Left = 171
        Top = 60
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = DATECOMPTABLE_
      end
      object TTL: TLabel
        Left = 13
        Top = 10
        Width = 49
        Height = 13
        Caption = '&Table libre'
        FocusControl = TABLE
      end
      object TAncVal: TLabel
        Left = 295
        Top = 35
        Width = 77
        Height = 13
        Caption = '&Ancienne valeur'
        FocusControl = AncVal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TNewVal: TLabel
        Left = 295
        Top = 60
        Width = 74
        Height = 13
        Caption = '&Nouvelle valeur'
        FocusControl = NewVal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object DATECOMPTABLE_: THCritMaskEdit
        Left = 196
        Top = 56
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 3
        Text = '  /  /    '
        OnKeyPress = DATECOMPTABLEKeyPress
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object DATECOMPTABLE: THCritMaskEdit
        Left = 75
        Top = 56
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 2
        Text = '  /  /    '
        OnKeyPress = DATECOMPTABLEKeyPress
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object JOURNAL: THValComboBox
        Left = 377
        Top = 6
        Width = 153
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTJOURNAL'
      end
      object EXERCICE: THValComboBox
        Left = 75
        Top = 31
        Width = 198
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 1
        OnChange = EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object TABLE: THValComboBox
        Left = 75
        Top = 6
        Width = 198
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 4
        OnChange = TABLEChange
        TagDispatch = 0
      end
      object AncVal: THCpteEdit
        Left = 377
        Top = 31
        Width = 153
        Height = 21
        CharCase = ecUpperCase
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 5
        ZoomTable = tzBudgen
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object NewVal: THCpteEdit
        Left = 377
        Top = 56
        Width = 153
        Height = 21
        CharCase = ecUpperCase
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 6
        ZoomTable = tzGeneral
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
    end
    object PComplement: TTabSheet
      Caption = 'Compl'#233'ments'
      object Bevel2: TBevel
        Left = 0
        Top = 0
        Width = 553
        Height = 83
        Align = alClient
      end
      object TE_GENERAL: THLabel
        Left = 241
        Top = 14
        Width = 37
        Height = 13
        Caption = '&G'#233'n'#233'ral'
        FocusControl = GENERAL
      end
      object TAUXILIAIRE: THLabel
        Left = 393
        Top = 14
        Width = 41
        Height = 13
        Caption = '&Auxiliaire'
        FocusControl = AUXILIAIRE
      end
      object TNATUREPIECE: THLabel
        Left = 12
        Top = 46
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = NATUREPIECE
      end
      object TE_NUMEROPIECE: THLabel
        Left = 241
        Top = 46
        Width = 56
        Height = 13
        Caption = 'N'#176' de &pi'#232'ce'
        FocusControl = NUMEROPIECE
      end
      object HLabel1: THLabel
        Left = 410
        Top = 46
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = DATECOMPTABLE_
      end
      object TE_QUALIFPIECE: THLabel
        Left = 12
        Top = 14
        Width = 24
        Height = 13
        Caption = '&Type'
        FocusControl = QUALIFPIECE
      end
      object Axe: THValComboBox
        Left = 56
        Top = 42
        Width = 154
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 2
        Visible = False
        OnChange = AxeChange
        TagDispatch = 0
        DataType = 'TTAXE'
      end
      object NATUREPIECE: THValComboBox
        Left = 56
        Top = 42
        Width = 165
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATUREPIECE'
      end
      object AUXILIAIRE: THCpteEdit
        Left = 447
        Top = 10
        Width = 69
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 1
        ZoomTable = tzTiers
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object GENERAL: THCpteEdit
        Left = 305
        Top = 10
        Width = 68
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 0
        ZoomTable = tzGeneral
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object NUMEROPIECE: THCritMaskEdit
        Left = 305
        Top = 42
        Width = 68
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object NUMEROPIECE_: THCritMaskEdit
        Left = 447
        Top = 42
        Width = 69
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 5
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object QUALIFPIECE: THValComboBox
        Left = 56
        Top = 10
        Width = 165
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 6
        TagDispatch = 0
        Vide = True
        DataType = 'TTQUALPIECE'
      end
    end
  end
  object PanelTxt: TPanel
    Left = 0
    Top = 282
    Width = 561
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 1
    object Nb1: TLabel
      Left = 221
      Top = 11
      Width = 8
      Height = 13
      Caption = '0'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Tex1: TLabel
      Left = 250
      Top = 11
      Width = 104
      Height = 13
      Caption = 'ligne s'#233'lectionn'#233'e'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Panel1: TPanel
      Left = 408
      Top = 2
      Width = 151
      Height = 31
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 0
    end
  end
  object FListe: THGrid
    Left = 0
    Top = 156
    Width = 561
    Height = 126
    Align = alClient
    ColCount = 8
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
    TabOrder = 2
    OnDblClick = FListeDblClick
    OnKeyDown = FListeKeyDown
    OnMouseDown = FListeMouseDown
    SortedCol = -1
    Couleur = True
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = 13224395
    ColWidths = (
      40
      75
      50
      92
      92
      92
      92
      0)
  end
  object Dock: TDock97
    Left = 0
    Top = 317
    Width = 561
    Height = 40
    AllowDrag = False
    Position = dpBottom
    object PanelBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 36
      ClientWidth = 561
      Caption = 'Barre outils multicrit'#232're'
      ClientAreaHeight = 36
      ClientAreaWidth = 561
      DockPos = 0
      FullSize = True
      TabOrder = 0
      DesignSize = (
        561
        36)
      object BTag: TToolbarButton97
        Left = 419
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Tout s'#233'lectionner'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Tout s'#233'lectionner '
        Flat = False
        Glyph.Data = {
          76010000424D7601000000000000760000002800000020000000100000000100
          04000000000000010000130B0000130B00001000000000000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
          333333333333333333FF33333333333330003333333333333888333333333333
          300033FFFFFF3333388839999993333333333888888F3333333F399999933333
          3300388888833333338833333333333333003333333333333388333333333333
          3333333333333333333F3333333E3333330033333F33333333883333333EC333
          330033338F3333333388EEEEEEEECCE3333333F88FFFFFFF3FF3ECCCCCCCCCCC
          399338888888888F88F33CCCCCCCCCCC399338888888888388333333333ECC33
          333333388F33333333FF33333333C33330003333833333333888333333333333
          3000333333333333388833333333333333333333333333333333}
        GlyphMask.Data = {00000000}
        Layout = blGlyphTop
        ModalResult = 2
        NumGlyphs = 2
        Visible = False
        OnClick = BTagClick
      end
      object BValider: TToolbarButton97
        Left = 464
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Valider'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Valider'
        Flat = False
        Layout = blGlyphTop
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object BFerme: TToolbarButton97
        Left = 494
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Anchors = [akTop, akRight]
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
        Layout = blGlyphTop
        ModalResult = 2
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 523
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Flat = False
        Layout = blGlyphTop
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
      object BDetag: TToolbarButton97
        Left = 419
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Tout d'#233'selectionner'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Tout d'#233'selectionner'
        Flat = False
        Layout = blGlyphTop
        NumGlyphs = 2
        OnClick = BdetagClick
        GlobalIndexImage = 'Z0078_S16G2'
      end
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 111
    Width = 561
    Height = 45
    AllowDrag = False
    object PanelFiltres: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 41
      ClientWidth = 561
      Caption = 'Filtres'
      ClientAreaHeight = 41
      ClientAreaWidth = 561
      DockPos = 0
      FullSize = True
      TabOrder = 0
      DesignSize = (
        561
        41)
      object BFiltre: TToolbarButton97
        Left = 6
        Top = 7
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
      object BCherche_: TToolbarButton97
        Left = 525
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Appliquer crit'#232'res'
        Anchors = [akTop, akRight]
        Flat = False
        NumGlyphs = 2
        Visible = False
        OnClick = BChercheClick
        GlobalIndexImage = 'Z0478_S16G2'
      end
      object BCherche: TToolbarButton97
        Left = 525
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Appliquer crit'#232'res'
        Anchors = [akTop, akRight]
        Flat = False
        OnClick = BChercheClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object FFiltres: THValComboBox
        Left = 68
        Top = 7
        Width = 445
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
      end
    end
  end
  object TITREE: THCombobox
    Left = 360
    Top = 192
    Width = 145
    Height = 21
    Color = clYellow
    ItemHeight = 13
    TabOrder = 5
    Text = 'TITREE'
    Visible = False
    Items.Strings = (
      'Jal;C;'
      'Date;C;'
      'Num;C;'
      'G'#233'n'#233'ral;G;'
      'Auxiliaire;G;'
      'D'#233'bit;G;'
      'Cr'#233'dit;G;')
  end
  object TITREA: THCombobox
    Left = 360
    Top = 216
    Width = 145
    Height = 21
    Color = clYellow
    ItemHeight = 13
    TabOrder = 6
    Text = 'TITREA'
    Visible = False
    Items.Strings = (
      'Jal;C;'
      'Date;C;'
      'Num;C;'
      'G'#233'n'#233'ral;G;'
      'Section;G;'
      'D'#233'bit;G;'
      'Cr'#233'dit;G;')
  end
  object TITREU: THCombobox
    Left = 360
    Top = 240
    Width = 145
    Height = 21
    Color = clYellow
    ItemHeight = 13
    TabOrder = 7
    Text = 'TITREU'
    Visible = False
    Items.Strings = (
      'Jal;C;'
      'Date;C;'
      'Num;C;'
      'Nature;C;'
      'G'#233'n'#233'ral;G;'
      'Section;G;'
      'Etablissement;G;')
  end
  object Q: TQuery
    DatabaseName = 'SOC'
    RequestLive = True
    Left = 121
    Top = 212
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 176
    Top = 212
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '&Auxiliaire'
      '&Axe'
      '&Section'
      'ligne s'#233'lectionn'#233'e'
      'lignes s'#233'lectionn'#233'es'
      
        '5;Modification en s'#233'rie des tables libres;Le code renseign'#233' n'#39'ex' +
        'iste pas;W;O;O;O;'
      
        '6;Modification en s'#233'rie des tables libres;Aucun compte ne corres' +
        'pond '#224' votre crit'#232're, aucun traitement '#224' effectuer;W;O;O;O;'
      
        '7;Modification en s'#233'rie des tables libres;Aucun compte n'#39'est s'#233'l' +
        'ectionn'#233', aucun traitement '#224' effectuer;W;O;O;O;'
      
        '8;Modification en s'#233'rie des tables libres;Confirmez-vous le trai' +
        'tement de mise '#224' jour des tables libres s'#233'lectionn'#233'es?;Q;YN;N;N;'
      
        '9;Modification en s'#233'rie des tables libres;La nouvelle valeur est' +
        ' '#233'gale '#224' l'#39'ancienne valeur. Aucun traitement '#224' effectuer !;W;O;O' +
        ';O;'
      'Modification des tables libres des '#233'critures comptables'
      'Modification des tables libres des '#233'critures analytiques'
      'Modification des tables libres des '#233'critures budg'#233'taires')
    Left = 228
    Top = 212
  end
  object SaveDialog: TSaveDialog
    DefaultExt = 'txt'
    Filter = 'Fichiers textes|*.txt'
    Options = [ofOverwritePrompt, ofHideReadOnly]
    Left = 67
    Top = 212
  end
  object POPF: TPopupMenu
    Left = 12
    Top = 212
    object BCreerFiltre: TMenuItem
      Caption = '&Cr'#233'er un filtre'
    end
    object BSaveFiltre: TMenuItem
      Caption = '&Enregistrer le filtre'
    end
    object BDelFiltre: TMenuItem
      Caption = '&Supprimer le filtre'
    end
    object BRenFiltre: TMenuItem
      Caption = '&Renommer le filtre'
    end
    object BNouvRech: TMenuItem
      Caption = '&Nouvelle recherche'
    end
  end
  object FSearchTimer: TTimer
    Enabled = False
    Left = 272
    Top = 188
  end
end

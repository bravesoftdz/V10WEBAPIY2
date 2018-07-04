object FLetRegul: TFLetRegul
  Left = 294
  Top = 185
  Width = 608
  Height = 390
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'R'#233'gularisation de lettrage'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 592
    Height = 114
    ActivePage = Princ
    Align = alTop
    TabOrder = 0
    object Princ: TTabSheet
      Caption = 'Crit'#232'res de s'#233'lection'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 584
        Height = 86
        Align = alClient
      end
      object HLabel4: THLabel
        Left = 9
        Top = 10
        Width = 61
        Height = 13
        Caption = '&G'#233'n'#233'raux de'
        FocusControl = E_GENERAL
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel1: THLabel
        Left = 9
        Top = 34
        Width = 61
        Height = 13
        Caption = '&Auxiliaires de'
        FocusControl = E_AUXILIAIRE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel3: THLabel
        Left = 9
        Top = 58
        Width = 100
        Height = 13
        Caption = '&Dates comptables du'
        FocusControl = E_DATECOMPTABLE
      end
      object HLabel5: THLabel
        Left = 215
        Top = 10
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_GENERAL_
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel2: THLabel
        Left = 215
        Top = 34
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_AUXILIAIRE_
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel6: THLabel
        Left = 213
        Top = 58
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATECOMPTABLE_
      end
      object HLabel8: THLabel
        Left = 345
        Top = 14
        Width = 33
        Height = 13
        Caption = 'D&evise'
        FocusControl = E_DEVISE
      end
      object Label14: THLabel
        Left = 344
        Top = 38
        Width = 58
        Height = 13
        Caption = '&Nature Tiers'
        FocusControl = T_NATUREAUXI
      end
      object ExpliEuro: TLabel
        Left = 440
        Top = 60
        Width = 86
        Height = 13
        Caption = 'Devise du lettrage'
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 115
        Top = 56
        Width = 93
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 4
        Text = '  /  /    '
        OnKeyPress = E_DATECOMPTABLEKeyPress
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_AUXILIAIRE: THCpteEdit
        Tag = 1
        Left = 115
        Top = 32
        Width = 92
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 2
        OnDblClick = E_AUXILIAIREDblClick
        ZoomTable = tzTlettrable
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object E_GENERAL: THCpteEdit
        Tag = 1
        Left = 115
        Top = 8
        Width = 92
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 0
        ZoomTable = tzGLettColl
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object E_DATECOMPTABLE_: THCritMaskEdit
        Left = 231
        Top = 56
        Width = 93
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 5
        Text = '  /  /    '
        OnKeyPress = E_DATECOMPTABLEKeyPress
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_AUXILIAIRE_: THCpteEdit
        Tag = 1
        Left = 231
        Top = 32
        Width = 92
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 3
        OnDblClick = E_AUXILIAIREDblClick
        ZoomTable = tzTlettrable
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object E_GENERAL_: THCpteEdit
        Tag = 1
        Left = 231
        Top = 8
        Width = 92
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 1
        ZoomTable = tzGLettColl
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_NATUREAUXI: THValComboBox
        Left = 409
        Top = 34
        Width = 169
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 7
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATTIERSCPTA'
      end
      object E_DEVISE: THValComboBox
        Tag = 1
        Left = 409
        Top = 8
        Width = 169
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 6
        OnChange = E_DEVISEChange
        TagDispatch = 0
        DataType = 'TTDEVISE'
      end
      object XX_WHERESEL: TEdit
        Left = 101
        Top = 1
        Width = 13
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 8
        Visible = False
      end
    end
    object PLibres: TTabSheet
      Caption = 'Tables libres'
      object Bevel3: TBevel
        Left = 0
        Top = 0
        Width = 592
        Height = 86
        Align = alClient
      end
      object TT_TABLE0: THLabel
        Left = 5
        Top = 5
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE0'
      end
      object TT_TABLE1: THLabel
        Left = 120
        Top = 5
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE1'
      end
      object TT_TABLE2: THLabel
        Left = 235
        Top = 5
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE2'
      end
      object TT_TABLE3: THLabel
        Left = 349
        Top = 5
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE3'
      end
      object TT_TABLE4: THLabel
        Left = 461
        Top = 5
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE4'
      end
      object TT_TABLE5: THLabel
        Left = 5
        Top = 45
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE5'
      end
      object TT_TABLE6: THLabel
        Left = 120
        Top = 45
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE6'
      end
      object TT_TABLE7: THLabel
        Left = 235
        Top = 45
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE7'
      end
      object TT_TABLE8: THLabel
        Left = 349
        Top = 45
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE8'
      end
      object TT_TABLE9: THLabel
        Left = 461
        Top = 44
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE9'
      end
      object T_TABLE4: THCpteEdit
        Left = 461
        Top = 20
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 0
        ZoomTable = tzNatTiers4
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE3: THCpteEdit
        Left = 349
        Top = 20
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 1
        ZoomTable = tzNatTiers3
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE2: THCpteEdit
        Left = 235
        Top = 20
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 2
        ZoomTable = tzNatTiers2
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE1: THCpteEdit
        Left = 120
        Top = 20
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 3
        ZoomTable = tzNatTiers1
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE0: THCpteEdit
        Left = 5
        Top = 20
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 4
        ZoomTable = tzNatTiers0
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE5: THCpteEdit
        Left = 5
        Top = 58
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 5
        ZoomTable = tzNatTiers5
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE6: THCpteEdit
        Left = 120
        Top = 58
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 6
        ZoomTable = tzNatTiers6
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE7: THCpteEdit
        Left = 235
        Top = 58
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 7
        ZoomTable = tzNatTiers7
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE8: THCpteEdit
        Left = 349
        Top = 58
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 8
        ZoomTable = tzNatTiers8
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE9: THCpteEdit
        Left = 461
        Top = 58
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 9
        ZoomTable = tzNatTiers9
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
    end
    object TabSheet1: TTabSheet
      Caption = 'Param'#232'tres de r'#233'gularisation'
      object H_JOURNALE: THLabel
        Left = 9
        Top = 9
        Width = 123
        Height = 13
        Caption = '&Journal d'#39#233'cart de change'
        FocusControl = JOURNAL
        Visible = False
        OnDblClick = H_JOURNALRDblClick
      end
      object Bevel2: TBevel
        Left = 0
        Top = 0
        Width = 592
        Height = 86
        Align = alClient
      end
      object H_JOURNALR: THLabel
        Left = 8
        Top = 9
        Width = 114
        Height = 13
        Caption = '&Journal de r'#233'gularisation'
        Visible = False
        OnDblClick = H_JOURNALRDblClick
      end
      object HLabel9: THLabel
        Left = 9
        Top = 34
        Width = 116
        Height = 13
        Caption = '&Date de comptabilisation'
        FocusControl = DATEGENERATION
      end
      object HLabel7: THLabel
        Left = 9
        Top = 61
        Width = 115
        Height = 13
        Caption = '&R'#233'f'#233'rence r'#233'gularisation'
        FocusControl = REFREGUL
      end
      object HLabel12: THLabel
        Left = 285
        Top = 9
        Width = 62
        Height = 13
        Caption = '&Compte d'#233'bit'
        FocusControl = CPTDEBIT
      end
      object HLabel13: THLabel
        Left = 285
        Top = 34
        Width = 65
        Height = 13
        Caption = 'C&ompte cr'#233'dit'
        FocusControl = CPTCREDIT
      end
      object HLabel10: THLabel
        Left = 285
        Top = 63
        Width = 56
        Height = 13
        Caption = '&Libell'#233' r'#233'gul'
        FocusControl = LIBREGUL
      end
      object H_SD: THLabel
        Left = 465
        Top = 9
        Width = 23
        Height = 13
        Caption = '&Seuil'
        FocusControl = MAXDEBIT
      end
      object H_SC: THLabel
        Left = 465
        Top = 36
        Width = 23
        Height = 13
        Caption = 'S&euil'
        FocusControl = MAXCREDIT
      end
      object H_JOURNALC: THLabel
        Left = 8
        Top = 9
        Width = 116
        Height = 13
        Caption = '&Journal '#233'cart conversion'
        Visible = False
        OnDblClick = H_JOURNALRDblClick
      end
      object REFREGUL: TEdit
        Left = 137
        Top = 58
        Width = 121
        Height = 21
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 6
      end
      object DATEGENERATION: THCritMaskEdit
        Left = 137
        Top = 32
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 1
        Text = '  /  /    '
        OnExit = DATEGENERATIONExit
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object JOURNAL: THValComboBox
        Tag = 1
        Left = 137
        Top = 5
        Width = 145
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 0
        OnChange = JOURNALChange
        TagDispatch = 0
        DataType = 'TTJALREGUL'
      end
      object LIBREGUL: TEdit
        Left = 353
        Top = 59
        Width = 229
        Height = 21
        MaxLength = 35
        TabOrder = 7
      end
      object CPTCREDIT: TComboBox
        Left = 353
        Top = 32
        Width = 105
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 4
      end
      object CPTDEBIT: TComboBox
        Left = 353
        Top = 5
        Width = 105
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 2
      end
      object MAXCREDIT: THNumEdit
        Left = 497
        Top = 32
        Width = 85
        Height = 21
        TabOrder = 5
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object MAXDEBIT: THNumEdit
        Left = 497
        Top = 5
        Width = 85
        Height = 21
        TabOrder = 3
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
    end
  end
  object Dock: TDock97
    Left = 0
    Top = 316
    Width = 592
    Height = 35
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 592
      Caption = 'Actions'
      ClientAreaHeight = 31
      ClientAreaWidth = 592
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BReduire: TToolbarButton97
        Left = 3
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Afficher les crit'#232'res'
        Caption = 'R'#233'duire'
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
        Left = 3
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Agrandir la liste'
        Caption = 'Agrandir'
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
      object BImprimer: TToolbarButton97
        Left = 471
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Caption = 'Imprimer'
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
      object BValider: TToolbarButton97
        Left = 503
        Top = 3
        Width = 28
        Height = 27
        Hint = 'G'#233'n'#233'rer les pi'#232'ces de r'#233'gularisation'
        Caption = 'G'#233'n'#233'rer'
        Default = True
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 3
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0184_S16G1'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 535
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Caption = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 567
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
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BRecherche: TToolbarButton97
        Left = 35
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Rechercher dans la liste'
        Caption = 'Chercher'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BRechercheClick
        GlobalIndexImage = 'Z0077_S16G1'
      end
      object BGenere: TToolbarButton97
        Tag = 1
        Left = 67
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Liste des pi'#232'ces g'#233'n'#233'r'#233'es'
        Caption = 'Liste'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 0
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BGenereClick
        GlobalIndexImage = 'Z1826_S16G1'
        IsControl = True
      end
      object bSelectAll: TToolbarButton97
        Left = 101
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Tout s'#233'lectionner'
        Caption = 'S'#233'lectionner'
        AllowAllUp = True
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        Visible = False
        OnClick = bSelectAllClick
        GlobalIndexImage = 'Z0205_S16G1'
      end
    end
  end
  object Cache: THCpteEdit
    Left = 8
    Top = 158
    Width = 17
    Height = 21
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
    Text = '!!!'
    Visible = False
    ZoomTable = tzGLettColl
    Vide = False
    Bourre = True
    okLocate = False
    SynJoker = False
  end
  object Dock971: TDock97
    Left = 0
    Top = 114
    Width = 592
    Height = 35
    AllowDrag = False
    object PFiltres: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 592
      ClientAreaHeight = 31
      ClientAreaWidth = 592
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BFiltre: TToolbarButton97
        Left = 2
        Top = 5
        Width = 56
        Height = 21
        Hint = 'Menu filtre'
        Caption = '&Filtres'
        DropdownArrow = True
        DropdownMenu = POPF
        Layout = blGlyphRight
        ParentShowHint = False
        ShowHint = True
      end
      object BChercher: TToolbarButton97
        Left = 564
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Appliquer crit'#232'res'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BChercherClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object FFiltres: THValComboBox
        Left = 64
        Top = 5
        Width = 493
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
      end
    end
  end
  object GR: THGrid
    Left = 0
    Top = 149
    Width = 592
    Height = 167
    Align = alClient
    ColCount = 2
    DefaultRowHeight = 18
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
    TabOrder = 4
    SortedCol = -1
    Couleur = False
    MultiSelect = True
    TitleBold = True
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = clSilver
    DBIndicator = True
    ColWidths = (
      10
      64)
  end
  object HMRegul: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'R'#233'gularisations de lettrage'
      'Diff'#233'rences de change'
      '2;?caption?;Vous n'#39'avez rien s'#233'lectionn'#233';E;O;O;O;'
      
        '3;?caption?;Vous devez choisir un journal de r'#233'gularisation;W;O;' +
        'O;O;'
      
        '4;?caption?;Vous devez renseigner les comptes de r'#233'gularisation;' +
        'W;O;O;O;'
      
        '5;?caption?;La date de g'#233'n'#233'ration que vous avez saisie est incor' +
        'recte;W;O;O;O;'
      
        '6;?caption?;La date de g'#233'n'#233'ration est sur un exercice non ouvert' +
        ';W;O;O;O;'
      
        '7;?caption?;La date de g'#233'n'#233'ration est sur un exercice non ouvert' +
        ';W;O;O;O;'
      
        '8;?caption?;La date de g'#233'n'#233'ration est ant'#233'rieure '#224' la cl'#244'ture pr' +
        'ovisoire;W;O;O;O;'
      
        '9;?caption?;La date de g'#233'n'#233'ration est ant'#233'rieure '#224' la cl'#244'ture d'#233 +
        'finitive;W;O;O;O;'
      
        '10;?caption?;Le compte de r'#233'gularisation d'#233'bit n'#39'existe pas;W;O;' +
        'O;O;'
      
        '11;?caption?;Le compte de r'#233'gularisation cr'#233'dit n'#39'existe pas;W;O' +
        ';O;O;'
      
        '12;?caption?;Confirmez-vous la g'#233'n'#233'ration des pi'#232'ces de r'#233'gulari' +
        'sation ?;Q;YN;Y;Y;'
      
        '13;?caption?;Voulez-vous voir la liste des pi'#232'ces g'#233'n'#233'r'#233'es ?;Q;Y' +
        'N;N;N;'
      
        '14;?caption?;Confirmez-vous la g'#233'n'#233'ration des pi'#232'ces d'#39#233'cart de ' +
        'change ?;Q;YN;Y;Y;'
      
        'R'#233'gularisation non effectu'#233'e! Certaines '#233'critures sont en traite' +
        'ment par un autre utilisateur'
      
        '16;?caption?;La r'#233'gularisation est impossible, le journal ne pos' +
        's'#232'de pas de facturier;W;O;O;O;'
      '17;'
      'Ecarts de conversion'
      
        '19;?caption?;Confirmez-vous la g'#233'n'#233'ration des pi'#232'ces d'#39#233'cart de ' +
        'conversion ?;Q;YN;Y;Y;'
      'Devise de l'#39#233'cart'
      'G'#233'n'#233'rer les pi'#232'ces d'#39#233'cart de conversion'
      'D'#233'bit Euro'
      'Cr'#233'dit Euro'
      'Solde Euro'
      'D'#233'bit'
      'Cr'#233'dit'
      'Solde'
      'R'#233'gularisation'
      'Diff'#233'rence de change'
      'Ecart de conversion'
      '31;')
    Left = 107
    Top = 276
  end
  object FindRegul: TFindDialog
    OnFind = FindRegulFind
    Left = 157
    Top = 276
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 256
    Top = 276
  end
  object POPF: TPopupMenu
    Left = 207
    Top = 276
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
end

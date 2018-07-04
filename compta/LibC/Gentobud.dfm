object FGenToBud: TFGenToBud
  Left = 218
  Top = 154
  Width = 554
  Height = 406
  HelpContext = 7577300
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'G'#233'n'#233'ration des comptes budg'#233'taires'
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
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 546
    Height = 114
    ActivePage = Pparam
    Align = alTop
    TabOrder = 0
    object Pparam: TTabSheet
      Caption = 'Param'#233'trage'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 538
        Height = 86
        Align = alClient
      end
      object TG_NATUREGENE: THLabel
        Left = 22
        Top = 12
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = G_NATUREGENE
      end
      object TG_GENERAL: TLabel
        Left = 22
        Top = 40
        Width = 51
        Height = 13
        Caption = '&Compte de'
        FocusControl = G_GENERAL
      end
      object TG_GENERAL_: TLabel
        Left = 284
        Top = 40
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = G_GENERAL_
      end
      object TG_VENTILABLE: THLabel
        Left = 255
        Top = 12
        Width = 87
        Height = 13
        Caption = '&Ventilable sur l'#39'axe'
        FocusControl = G_VENTILABLE
      end
      object Nb1: TLabel
        Left = 93
        Top = 64
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
      end
      object Tex1: TLabel
        Left = 118
        Top = 64
        Width = 132
        Height = 13
        Caption = 'ligne(s) s'#233'lectionn'#233'e(s)'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object G_NATUREGENE: THValComboBox
        Left = 83
        Top = 8
        Width = 155
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = G_NATUREGENEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATGENE'
      end
      object G_GENERAL: THCpteEdit
        Left = 82
        Top = 36
        Width = 155
        Height = 21
        TabOrder = 1
        ZoomTable = tzGeneral
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_GENERAL_: THCpteEdit
        Left = 349
        Top = 36
        Width = 155
        Height = 21
        TabOrder = 2
        ZoomTable = tzGeneral
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_VENTILABLE: THValComboBox
        Left = 349
        Top = 8
        Width = 155
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        TagDispatch = 0
        Vide = True
        DataType = 'TTAXE'
      end
    end
    object PzLibre: TTabSheet
      Caption = 'Tables libres'
      object Bevel5: TBevel
        Left = 0
        Top = 0
        Width = 538
        Height = 86
        Align = alClient
      end
      object TG_TABLE0: TLabel
        Left = 7
        Top = 2
        Width = 93
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE0'
      end
      object TG_TABLE1: TLabel
        Left = 114
        Top = 2
        Width = 93
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE1'
      end
      object TG_TABLE2: TLabel
        Left = 221
        Top = 2
        Width = 93
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE2'
      end
      object TG_TABLE3: TLabel
        Left = 328
        Top = 2
        Width = 93
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE3'
      end
      object TG_TABLE4: TLabel
        Left = 435
        Top = 2
        Width = 93
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE4'
      end
      object TG_TABLE5: TLabel
        Left = 7
        Top = 44
        Width = 93
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE5'
      end
      object TG_TABLE6: TLabel
        Left = 114
        Top = 44
        Width = 93
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE6'
      end
      object TG_TABLE7: TLabel
        Left = 221
        Top = 44
        Width = 93
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE7'
      end
      object TG_TABLE8: TLabel
        Left = 328
        Top = 44
        Width = 93
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE8'
      end
      object TG_TABLE9: TLabel
        Left = 435
        Top = 44
        Width = 93
        Height = 13
        AutoSize = False
        Caption = 'TG_TABLE9'
      end
      object G_TABLE0: THCpteEdit
        Left = 7
        Top = 17
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 0
        ZoomTable = tzNatGene0
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_TABLE1: THCpteEdit
        Left = 114
        Top = 17
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 1
        ZoomTable = tzNatGene1
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_TABLE2: THCpteEdit
        Left = 221
        Top = 17
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 2
        ZoomTable = tzNatGene2
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_TABLE3: THCpteEdit
        Left = 328
        Top = 17
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 3
        ZoomTable = tzNatGene3
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_TABLE4: THCpteEdit
        Left = 435
        Top = 17
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 4
        ZoomTable = tzNatGene4
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_TABLE5: THCpteEdit
        Left = 7
        Top = 59
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 5
        ZoomTable = tzNatGene5
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_TABLE6: THCpteEdit
        Left = 114
        Top = 59
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 6
        ZoomTable = tzNatGene6
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_TABLE7: THCpteEdit
        Left = 221
        Top = 59
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 7
        ZoomTable = tzNatGene7
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_TABLE8: THCpteEdit
        Left = 328
        Top = 59
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 8
        ZoomTable = tzNatGene8
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object G_TABLE9: THCpteEdit
        Left = 435
        Top = 59
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 9
        ZoomTable = tzNatGene9
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
    end
  end
  object FListe: THGrid
    Left = 0
    Top = 153
    Width = 546
    Height = 187
    Align = alClient
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
    TabOrder = 1
    OnKeyDown = FListeKeyDown
    OnMouseDown = FListeMouseDown
    SortedCol = -1
    Titres.Strings = (
      'Code'
      'Libell'#233
      'Sens;C;'
      'Utilisateur;C;')
    Couleur = True
    MultiSelect = True
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = 13224395
    ColWidths = (
      135
      247
      56
      83
      0)
  end
  object Dock: TDock97
    Left = 0
    Top = 340
    Width = 546
    Height = 39
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 35
      ClientWidth = 546
      Caption = 'Actions'
      ClientAreaHeight = 35
      ClientAreaWidth = 546
      DockPos = 0
      FullSize = True
      ParentShowHint = False
      ShowHint = False
      TabOrder = 0
      object Bdetag: TToolbarButton97
        Left = 3
        Top = 2
        Width = 28
        Height = 27
        Hint = 'D'#233'selectionne tout'
        DisplayMode = dmGlyphOnly
        Caption = 'D'#233's'#233'lectionne'
        Flat = False
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BdetagClick
        GlobalIndexImage = 'Z0078_S16G2'
      end
      object BTag: TToolbarButton97
        Left = 3
        Top = 2
        Width = 28
        Height = 27
        Hint = 'S'#233'lectionne tout'
        DisplayMode = dmGlyphOnly
        Caption = 'S'#233'lectionne'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BTagClick
        GlobalIndexImage = 'Z0205_S16G1'
      end
      object BValider: TToolbarButton97
        Left = 447
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Lancer la cr'#233'ation'
        DisplayMode = dmGlyphOnly
        Caption = 'Cr'#233'er'
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
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 479
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 511
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
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
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 114
    Width = 546
    Height = 39
    AllowDrag = False
    object PFiltres: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 35
      ClientWidth = 546
      Caption = ' '
      ClientAreaHeight = 35
      ClientAreaWidth = 546
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BFiltre: TToolbarButton97
        Left = 6
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
        Left = 509
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
        Left = 68
        Top = 5
        Width = 433
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
      end
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    OnBeforeTraduc = HMTradBeforeTraduc
    Traduction = True
    Left = 16
    Top = 244
  end
  object POPF: TPopupMenu
    Left = 16
    Top = 292
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
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Ligne s'#233'lectionn'#233'e'
      'Lignes s'#233'lectionn'#233'es'
      
        '2;G'#233'n'#233'ration des comptes budg'#233'taires;La fourchette de compte n'#39'e' +
        'st pas valide. Aucun traitement a effectu'#233';W;O;O;O;'
      
        '3;G'#233'n'#233'ration des comptes budg'#233'taires;Aucun traitement '#224' effectue' +
        'r;W;O;O;O;'
      
        '4;G'#233'n'#233'ration des comptes budg'#233'taires;D'#233'sirez cr'#233'er ou modifier l' +
        'es comptes budg'#233'taires '#224' partir de votre s'#233'lection;Q;YNC;Y;C;')
    Left = 16
    Top = 196
  end
end

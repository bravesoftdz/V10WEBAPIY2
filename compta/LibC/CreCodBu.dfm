object FCreCodBu: TFCreCodBu
  Left = 227
  Top = 119
  Width = 620
  Height = 384
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 
    'G'#233'n'#233'ration des comptes budg'#233'taires '#224' partir des valeurs des tabl' +
    'es libres'
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
    Width = 612
    Height = 114
    ActivePage = Pparam
    Align = alTop
    TabOrder = 0
    object Pparam: TTabSheet
      Caption = 'Param'#233'trage'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 604
        Height = 86
        Align = alClient
      end
      object TSens: THLabel
        Left = 14
        Top = 20
        Width = 24
        Height = 13
        Caption = '&Sens'
        FocusControl = Sens
      end
      object TSigne: THLabel
        Left = 538
        Top = 9
        Width = 27
        Height = 13
        Caption = 'S&igne'
        Color = clYellow
        FocusControl = Signe
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        Visible = False
      end
      object LTri: TLabel
        Left = 14
        Top = 57
        Width = 59
        Height = 13
        Caption = 'Tables libres'
      end
      object TAxe: THLabel
        Left = 257
        Top = 21
        Width = 18
        Height = 13
        Caption = '&Axe'
        FocusControl = Axe
      end
      object Buse: TToolbarButton97
        Left = 408
        Top = 53
        Width = 25
        Height = 21
        Hint = 'Choisir les tables '#224' utiliser'
        ParentShowHint = False
        ShowHint = True
        OnClick = BuseClick
        GlobalIndexImage = 'Z0008_S16G1'
      end
      object Sens: THValComboBox
        Left = 77
        Top = 16
        Width = 143
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
        DataType = 'TTSENS'
      end
      object Signe: THValComboBox
        Left = 571
        Top = 5
        Width = 32
        Height = 21
        Style = csDropDownList
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
        DataType = 'TTRUBSIGNE'
      end
      object EChoix: TEdit
        Left = 77
        Top = 53
        Width = 329
        Height = 21
        CharCase = ecUpperCase
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = False
        TabOrder = 3
      end
      object BCherche: THBitBtn
        Left = 567
        Top = 46
        Width = 28
        Height = 27
        Hint = 'Rechercher les combinaisons'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 4
        OnClick = BChercheClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object Axe: THValComboBox
        Left = 290
        Top = 17
        Width = 143
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 1
        OnChange = AxeChange
        TagDispatch = 0
        DataType = 'TTAXE'
      end
      object CbRac: TRadioGroup
        Left = 440
        Top = 13
        Width = 112
        Height = 60
        Caption = ' Rubriques '
        ItemIndex = 1
        Items.Strings = (
          '&Racines'
          '&Comptes entiers')
        TabOrder = 5
      end
    end
  end
  object FListe: THGrid
    Left = 0
    Top = 114
    Width = 612
    Height = 208
    Align = alClient
    ColCount = 4
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
      'Rubrique')
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
      120
      195
      275
      0)
  end
  object HPB: TPanel
    Left = 0
    Top = 322
    Width = 612
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    Ctl3D = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentCtl3D = False
    ParentFont = False
    ParentShowHint = False
    ShowHint = False
    TabOrder = 2
    object Tex1: TLabel
      Left = 207
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
    object Nb1: TLabel
      Left = 182
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
    object Panel1: TPanel
      Left = 480
      Top = 2
      Width = 130
      Height = 31
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 0
      DesignSize = (
        130
        31)
      object Bdetag: THBitBtn
        Left = 3
        Top = 2
        Width = 28
        Height = 27
        Hint = 'D'#233'selectionne tout'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = BdetagClick
        NumGlyphs = 2
        GlobalIndexImage = 'Z0078_S16G2'
      end
      object BTag: THBitBtn
        Left = 3
        Top = 2
        Width = 28
        Height = 27
        Hint = 'S'#233'lectionne tout'
        Anchors = [akTop, akRight]
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        Visible = False
        OnClick = BTagClick
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
        NumGlyphs = 2
      end
      object BAide: THBitBtn
        Left = 99
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        HelpContext = 1395200
        Anchors = [akTop, akRight]
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 4
        OnClick = BAideClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BFerme: THBitBtn
        Left = 67
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Anchors = [akTop, akRight]
        Cancel = True
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BValider: THBitBtn
        Left = 35
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Lancer la cr'#233'ation'
        Anchors = [akTop, akRight]
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = BValiderClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z0127_S16G1'
        IsControl = True
      end
    end
  end
  object CbTabli: THValComboBox
    Left = 280
    Top = 152
    Width = 37
    Height = 21
    Style = csDropDownList
    Color = clYellow
    ItemHeight = 13
    TabOrder = 3
    Visible = False
    TagDispatch = 0
  end
  object ListCpte: THValComboBox
    Left = 240
    Top = 200
    Width = 25
    Height = 21
    Style = csDropDownList
    Color = clYellow
    ItemHeight = 13
    TabOrder = 4
    Visible = False
    TagDispatch = 0
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'G'#233'n'#233'ration des comptes budg'#233'taires '#224' partir des tables libres'
      'G'#233'n'#233'ration des sections budg'#233'taires '#224' partir des tables libres'
      
        '2;?Caption?;Un des codes saisis n'#39'existe pas ou est dupliqu'#233'.;W;' +
        'O;O;O;'
      
        '3;?Caption?;Codes '#224' utiliser non renseign'#233'. Saisie obligatoire.;' +
        'W;O;O;O;'
      'Ligne s'#233'lectionn'#233'e'
      'Lignes s'#233'lectionn'#233'es'
      
        '6;?Caption?;La rubrique est trop longue. Elle va '#234'tre tronqu'#233'e.;' +
        'I;O;O;O;'
      
        '7;?Caption?;Des codes budg'#233'taires n'#39'ont pas '#233't'#233' cr'#233#233'. Le code bu' +
        'dget ou le code rubrique existait d'#233'j'#224'. D'#233'sirez-vous les modifie' +
        'r pour les cr'#233#233'r ?;Q;YNC;Y;N;')
    Left = 408
    Top = 200
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 48
    Top = 180
  end
end

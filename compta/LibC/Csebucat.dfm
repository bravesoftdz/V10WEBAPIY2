object FCsebucat: TFCsebucat
  Left = 272
  Top = 151
  Width = 587
  Height = 362
  HelpContext = 7577450
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 
    'G'#233'n'#233'ration des sections budg'#233'taires '#224' partir des cat'#233'gories de b' +
    'udget'
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
    Width = 579
    Height = 105
    ActivePage = Param
    Align = alTop
    TabOrder = 0
    object Param: TTabSheet
      Caption = 'Param'#233'trage'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 571
        Height = 77
        Align = alClient
      end
      object THLabel
        Left = 31
        Top = 22
        Width = 3
        Height = 13
      end
      object TSigne: THLabel
        Left = 314
        Top = 52
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
      object TSens: THLabel
        Left = 16
        Top = 50
        Width = 24
        Height = 13
        Caption = '&Sens'
        FocusControl = Sens
      end
      object TAxe: THLabel
        Left = 16
        Top = 18
        Width = 18
        Height = 13
        Caption = '&Axe'
        FocusControl = Axe
      end
      object TcbCatjal: THLabel
        Left = 263
        Top = 10
        Width = 68
        Height = 29
        AutoSize = False
        Caption = '&Cat'#233'gories de budgets'
        FocusControl = cbCatjal
        WordWrap = True
      end
      object Sens: THValComboBox
        Left = 52
        Top = 46
        Width = 186
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
        DataType = 'TTSENS'
      end
      object Signe: THValComboBox
        Left = 355
        Top = 48
        Width = 27
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
        TabOrder = 1
        Visible = False
        TagDispatch = 0
        DataType = 'TTRUBSIGNE'
      end
      object Axe: THValComboBox
        Left = 52
        Top = 14
        Width = 186
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
        OnChange = AxeChange
        TagDispatch = 0
        DataType = 'TTAXE'
      end
      object BCherche: THBitBtn
        Left = 531
        Top = 43
        Width = 28
        Height = 27
        Hint = 'Appliquer crit'#232'res'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnClick = BChercheClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object cbCatjal: THValComboBox
        Left = 332
        Top = 14
        Width = 186
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 4
        TagDispatch = 0
      end
    end
  end
  object HPB: TPanel
    Left = 0
    Top = 300
    Width = 579
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
    TabOrder = 1
    object Tex1: TLabel
      Left = 210
      Top = 11
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
    object Nb1: TLabel
      Left = 185
      Top = 11
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
    object Panel1: TPanel
      Left = 414
      Top = 2
      Width = 163
      Height = 31
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 0
      DesignSize = (
        163
        31)
      object BTag: THBitBtn
        Left = 3
        Top = 2
        Width = 28
        Height = 27
        Hint = 'S'#233'lectionne tout'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        Visible = False
        OnClick = BTagClick
        GlobalIndexImage = 'Z0205_S16G1'
      end
      object Bdetag: THBitBtn
        Left = 3
        Top = 2
        Width = 28
        Height = 27
        Hint = 'D'#233'selectionne tout'
        Anchors = [akTop, akRight]
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = BdetagClick
        NumGlyphs = 2
        GlobalIndexImage = 'Z0078_S16G2'
      end
      object BAide: THBitBtn
        Left = 132
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
        Left = 100
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
        Left = 68
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Valider'
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
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
      object BCodeAbr: THBitBtn
        Left = 36
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Codes abr'#233'g'#233's des plans de sous sections'
        Anchors = [akTop, akRight]
        ParentShowHint = False
        ShowHint = True
        TabOrder = 5
        OnClick = BCodeAbrClick
        GlobalIndexImage = 'Z0210_S16G1'
      end
    end
  end
  object FListe: THGrid
    Left = 0
    Top = 105
    Width = 579
    Height = 195
    Align = alClient
    ColCount = 4
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
    TabOrder = 2
    OnKeyDown = FListeKeyDown
    OnMouseDown = FListeMouseDown
    SortedCol = -1
    Titres.Strings = (
      'Code'
      'Libell'#233
      'Code rubrique')
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
      175
      260
      120
      0)
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 48
    Top = 180
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
      
        '2;G'#233'n'#233'ration des sections budg'#233'taires '#224' partir des cat'#233'gories de' +
        ' budget;Les sections budg'#233'taires non mouvement'#233'es vont '#234'tre d'#233'tr' +
        'uites et rempla'#231#233'es par celle s'#233'lectionn'#233'es. Confirmez-vous le t' +
        'raitement ?;Q;YN;N;N; '
      'Destruction'
      'Cr'#233'ation'
      
        '5;G'#233'n'#233'ration des sections budg'#233'taires;Vous n'#39'avez pas param'#233'tr'#233' ' +
        'la copie des sections analytiques vers les sections budg'#233'taires ' +
        'dans la fiche soci'#233't'#233'.Vous ne pouvez pas effectuer cette op'#233'rati' +
        'on.;W;O;O;O; '
      
        '6;G'#233'n'#233'ration des sections budg'#233'taires;Des sections budg'#233'taires n' +
        #39'ont pas '#233't'#233' cr'#233#233'. Le code section ou le code rubrique existait ' +
        'd'#233'j'#224' ou n'#39'est pas renseign'#233'. D'#233'sirez-vous les modifier pour les ' +
        'cr'#233#233'r ?;Q;YNC;Y;N;')
    Left = 412
    Top = 164
  end
end

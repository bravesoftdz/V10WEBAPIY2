object FCopiTabL: TFCopiTabL
  Left = 152
  Top = 179
  HelpContext = 1205300
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Recopie de tables libres'
  ClientHeight = 177
  ClientWidth = 544
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object HPB: TPanel
    Left = 0
    Top = 142
    Width = 544
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
    TabOrder = 0
    object Panel1: TPanel
      Left = 439
      Top = 2
      Width = 103
      Height = 31
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 0
      object BAide: THBitBtn
        Left = 71
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        HelpContext = 1395200
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = BAideClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BFerme: THBitBtn
        Left = 39
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = BFermeClick
        Glyph.Data = {
          CE070000424DCE07000000000000360000002800000024000000120000000100
          1800000000009807000000000000000000000000000000000000008284008284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284008284008284008284008284008284008284008284
          0082840082840082840082840082840082848482848482840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284FFFFFF008284008284008284008284008284008284
          0082840082840082840082840082840082840082840082840082840000FF0000
          840000848482840082840082840082840082840082840000FF84828400828400
          8284008284008284008284008284008284008284848284848284FFFFFF008284
          008284008284008284008284008284FFFFFF0082840082840082840082840082
          840082840082840000FF00008400008400008484828400828400828400828400
          00FF000084000084848284008284008284008284008284008284008284848284
          FFFFFF008284848284FFFFFF008284008284008284FFFFFF848284848284FFFF
          FF0082840082840082840082840082840082840000FF00008400008400008400
          00848482840082840000FF000084000084000084000084848284008284008284
          008284008284008284848284FFFFFF008284008284848284FFFFFF008284FFFF
          FF848284008284008284848284FFFFFF00828400828400828400828400828400
          82840000FF000084000084000084000084848284000084000084000084000084
          000084848284008284008284008284008284008284848284FFFFFF0082840082
          84008284848284FFFFFF848284008284008284008284008284848284FFFFFF00
          82840082840082840082840082840082840000FF000084000084000084000084
          0000840000840000840000848482840082840082840082840082840082840082
          84008284848284FFFFFF00828400828400828484828400828400828400828400
          8284FFFFFF848284008284008284008284008284008284008284008284008284
          0000FF0000840000840000840000840000840000848482840082840082840082
          84008284008284008284008284008284008284848284FFFFFF00828400828400
          8284008284008284008284FFFFFF848284008284008284008284008284008284
          0082840082840082840082840082840000840000840000840000840000848482
          8400828400828400828400828400828400828400828400828400828400828400
          8284848284FFFFFF008284008284008284008284008284848284008284008284
          0082840082840082840082840082840082840082840082840082840000FF0000
          8400008400008400008484828400828400828400828400828400828400828400
          8284008284008284008284008284008284848284FFFFFF008284008284008284
          8482840082840082840082840082840082840082840082840082840082840082
          840082840000FF00008400008400008400008400008484828400828400828400
          8284008284008284008284008284008284008284008284008284008284848284
          008284008284008284008284848284FFFFFF0082840082840082840082840082
          840082840082840082840082840000FF00008400008400008484828400008400
          0084000084848284008284008284008284008284008284008284008284008284
          008284008284848284008284008284008284008284008284848284FFFFFF0082
          840082840082840082840082840082840082840082840000FF00008400008400
          00848482840082840000FF000084000084000084848284008284008284008284
          008284008284008284008284008284848284008284008284008284848284FFFF
          FF008284008284848284FFFFFF00828400828400828400828400828400828400
          82840000FF0000840000848482840082840082840082840000FF000084000084
          000084848284008284008284008284008284008284008284848284FFFFFF0082
          84008284848284008284848284FFFFFF008284008284848284FFFFFF00828400
          82840082840082840082840082840082840000FF000084008284008284008284
          0082840082840000FF0000840000840000840082840082840082840082840082
          84008284848284FFFFFFFFFFFF848284008284008284008284848284FFFFFF00
          8284008284848284FFFFFF008284008284008284008284008284008284008284
          0082840082840082840082840082840082840082840000FF0000840000FF0082
          8400828400828400828400828400828400828484828484828400828400828400
          8284008284008284848284FFFFFFFFFFFFFFFFFF848284008284008284008284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284008284008284008284008284848284848284848284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284008284008284008284008284008284008284008284
          008284008284008284008284008284008284}
        NumGlyphs = 2
      end
      object BValider: THBitBtn
        Left = 7
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Lancer la recopie'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnClick = BValiderClick
        Margin = 2
        NumGlyphs = 2
        Spacing = -1
        GlobalIndexImage = 'Z0127_S16G1'
        IsControl = True
      end
    end
  end
  object GBTabli: TGroupBox
    Left = 5
    Top = 4
    Width = 264
    Height = 133
    Caption = ' Source '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clActiveCaption
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    object TCbTlS: TLabel
      Left = 14
      Top = 24
      Width = 27
      Height = 13
      Caption = '&Entit'#233
      Color = clBtnFace
      FocusControl = CbEnS
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object TCbTlD: TLabel
      Left = 14
      Top = 60
      Width = 49
      Height = 13
      Caption = '&Table libre'
      Color = clBtnFace
      FocusControl = CbTlS
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object CbEnS: THValComboBox
      Left = 70
      Top = 20
      Width = 190
      Height = 21
      Style = csDropDownList
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 0
      OnChange = CbEnSChange
      Items.Strings = (
        'Table libre des '#233'critures analytiques'
        'Table libre des comptes budg'#233'taires'
        'Table libre des sections budg'#233'taires'
        'Table libre des '#233'critures comptables'
        'Table libre des comptes g'#233'n'#233'raux'
        'Table libre des sections analytiques'
        'Table libre des comptes auxiliaires'
        'Table libre des '#233'critures budg'#233'taires')
      TagDispatch = 0
      Values.Strings = (
        'A'
        'B'
        'D'
        'E'
        'G'
        'S'
        'T'
        'U')
    end
    object CbTlS: THValComboBox
      Left = 70
      Top = 56
      Width = 190
      Height = 21
      Style = csDropDownList
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 1
      TagDispatch = 0
    end
  end
  object GroupBox1: TGroupBox
    Left = 273
    Top = 4
    Width = 264
    Height = 133
    Caption = ' Destination '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clActiveCaption
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    object Label1: TLabel
      Left = 14
      Top = 24
      Width = 27
      Height = 13
      Caption = 'E&ntit'#233
      Color = clBtnFace
      FocusControl = CbEnS
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object Label2: TLabel
      Left = 14
      Top = 60
      Width = 49
      Height = 13
      Caption = 'T&able libre'
      Color = clBtnFace
      FocusControl = CbTlS
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object TCbAxeD: TLabel
      Left = 14
      Top = 99
      Width = 18
      Height = 13
      Caption = '&Axe'
      FocusControl = CbAxeD
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object CbEnD: THValComboBox
      Left = 68
      Top = 20
      Width = 190
      Height = 21
      Style = csDropDownList
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 0
      OnChange = CbEnSChange
      Items.Strings = (
        'Table libre des '#233'critures analytiques'
        'Table libre des comptes budg'#233'taires'
        'Table libre des sections budg'#233'taires'
        'Table libre des '#233'critures comptables'
        'Table libre des comptes g'#233'n'#233'raux'
        'Table libre des sections analytiques'
        'Table libre des comptes auxiliaires'
        'Table libre des '#233'critures budg'#233'taires')
      TagDispatch = 0
      Values.Strings = (
        'A'
        'B'
        'D'
        'E'
        'G'
        'S'
        'T'
        'U')
    end
    object CbTlD: THValComboBox
      Left = 68
      Top = 56
      Width = 190
      Height = 21
      Style = csDropDownList
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 1
      OnChange = CbTlSChange
      TagDispatch = 0
    end
    object CbAxeD: THValComboBox
      Left = 67
      Top = 95
      Width = 190
      Height = 21
      Style = csDropDownList
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 2
      TagDispatch = 0
      DataType = 'TTAXE'
    end
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Recopie de tables libres;D'#233'sirez vous copier la table source v' +
        'ers la table destination?;Q;YN;Y;N;'
      
        '1;Recopie de tables libres;La table destination est la m'#234'me que ' +
        'la table source!;W;O;O;O;'
      
        '2;Recopie de tables libres;Recopie effectu'#233'e avec succ'#232's.;I;O;O;' +
        'O;'
      
        '3;Recopie de tables libres;Attention : Si la liste des donn'#233'es d' +
        'e la table libre d'#39'origine ne contient pas toutes les valeurs de' +
        ' la table libre de destination et si ces valeurs avaient '#233't'#233' uti' +
        'lis'#233'es dans la table ;E;O;O;O;'
      ''
      ' ')
    Left = 180
    Top = 104
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 32
    Top = 104
  end
end

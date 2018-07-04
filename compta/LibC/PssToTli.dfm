object FPssToTli: TFPssToTli
  Left = 260
  Top = 191
  HelpContext = 1205400
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'G'#233'n'#233'ration tables libres depuis sous-sections'
  ClientHeight = 246
  ClientWidth = 350
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object GBTabli: TGroupBox
    Left = 5
    Top = 6
    Width = 339
    Height = 88
    Caption = ' Source '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clActiveCaption
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    object TCbTlS: TLabel
      Left = 11
      Top = 21
      Width = 98
      Height = 13
      Caption = '&Plan de sous section'
      Color = clBtnFace
      FocusControl = CbPss
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object TCbAxeD: TLabel
      Left = 11
      Top = 54
      Width = 18
      Height = 13
      Caption = '&Axe'
      FocusControl = CbAxeS
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object CbPss: THValComboBox
      Left = 118
      Top = 18
      Width = 210
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
      TagDispatch = 0
    end
    object CbAxeS: THValComboBox
      Left = 118
      Top = 50
      Width = 210
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
      OnChange = CbAxeSChange
      TagDispatch = 0
      DataType = 'TTAXE'
    end
  end
  object GBAxe: TGroupBox
    Left = 5
    Top = 150
    Width = 339
    Height = 54
    Caption = ' Destination '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clActiveCaption
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    object TCbTlD: TLabel
      Left = 15
      Top = 25
      Width = 49
      Height = 13
      Caption = '&Table libre'
      Color = clBtnFace
      FocusControl = CbTlD
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object CbTlD: THValComboBox
      Left = 122
      Top = 21
      Width = 210
      Height = 21
      Style = csDropDownList
      DropDownCount = 10
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 0
      TagDispatch = 0
    end
  end
  object HPB: TPanel
    Left = 0
    Top = 211
    Width = 350
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
    object Panel1: TPanel
      Left = 245
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
        Hint = 'Valider'
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
  object RgEntite: TRadioGroup
    Left = 5
    Top = 98
    Width = 339
    Height = 49
    Caption = ' Entit'#233' de destination '
    Columns = 2
    ItemIndex = 0
    Items.Strings = (
      'Sections a&nalytiques'
      'Sections &budg'#233'taires')
    TabOrder = 3
    OnClick = RgEntiteClick
  end
  object Q1: TQuery
    DatabaseName = 'SOC'
    SQL.Strings = (
      'Select * From NATCPTE')
    Left = 17
    Top = 195
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 68
    Top = 192
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Copie de table libre;D'#233'sirez-vous copier le plan de sous-secti' +
        'ons vers la table libre destination?;Q;YNC;Y;N;')
    Left = 116
    Top = 196
  end
end

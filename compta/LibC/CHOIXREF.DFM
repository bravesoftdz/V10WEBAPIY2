object FChoixRef: TFChoixRef
  Left = 383
  Top = 191
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Choix des r'#233'f'#233'rences de pointage'
  ClientHeight = 185
  ClientWidth = 282
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
  object Liste: THGrid
    Left = 0
    Top = 0
    Width = 282
    Height = 150
    Align = alClient
    ColCount = 2
    DefaultColWidth = 128
    DefaultRowHeight = 17
    RowCount = 8
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goTabs]
    ScrollBars = ssVertical
    TabOrder = 0
    OnDblClick = ListeDblClick
    OnKeyPress = ListeKeyPress
    SortedCol = -1
    Titres.Strings = (
      'Compte g'#233'n'#233'ral'
      'R'#233'f'#233'rence')
    Couleur = False
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = 13224395
    ColWidths = (
      101
      155)
    RowHeights = (
      17
      17
      17
      17
      17
      17
      17
      17)
  end
  object Panel5: TPanel
    Left = 0
    Top = 150
    Width = 282
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 1
    object BValider: THBitBtn
      Tag = 1
      Left = 184
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Mise '#224' jour des r'#233'f'#233'rences'
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
    object BHelp: THBitBtn
      Tag = 1
      Left = 248
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Aide'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z1117_S16G1'
      IsControl = True
    end
    object BAbandon: THBitBtn
      Tag = 1
      Left = 216
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ModalResult = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0021_S16G1'
      IsControl = True
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 120
    Top = 60
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Choix d'#39'une r'#233'f'#233'rence de pointage'
      
        '1;?Caption?;Vous devez indiquer une r'#233'f'#233'rence de pointage pour c' +
        'haque compte !;W;O;O;O;0')
    Left = 156
    Top = 60
  end
end

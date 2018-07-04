object FCumMens: TFCumMens
  Left = 279
  Top = 169
  BorderIcons = [biSystemMenu, biMaximize]
  BorderStyle = bsSingle
  Caption = 'Cumuls mensuels du compte '
  ClientHeight = 377
  ClientWidth = 577
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clBlack
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Fliste: THGrid
    Left = 0
    Top = 60
    Width = 577
    Height = 282
    TabStop = False
    Align = alClient
    ColCount = 4
    Ctl3D = True
    DefaultColWidth = 100
    DefaultRowHeight = 18
    RowCount = 2
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
    ParentCtl3D = False
    ParentFont = False
    ParentShowHint = False
    ScrollBars = ssVertical
    ShowHint = False
    TabOrder = 0
    OnDblClick = FlisteDblClick
    SortedCol = -1
    Titres.Strings = (
      'P'#233'riode;G'
      'D'#233'bit;D'
      'Cr'#233'dit;D'
      'Solde;D')
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = clSilver
    ColWidths = (
      105
      150
      150
      150)
    RowHeights = (
      18
      16)
  end
  object Pbouton: TPanel
    Left = 0
    Top = 342
    Width = 577
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 1
    object LAnouv: TLabel
      Left = 252
      Top = 12
      Width = 54
      Height = 13
      Caption = 'A Nouveau'
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      Visible = False
    end
    object LTotaux: TLabel
      Left = 184
      Top = 12
      Width = 33
      Height = 13
      Caption = 'Totaux'
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      Visible = False
    end
    object NumEditSold: THNumEdit
      Left = 320
      Top = 8
      Width = 45
      Height = 21
      Color = clYellow
      Decimals = 2
      Digits = 15
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      NumericType = ntDC
      ParentFont = False
      TabOrder = 0
      UseRounding = True
      Validate = False
    end
    object TPanel
      Left = 446
      Top = 2
      Width = 129
      Height = 31
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 1
      object BAide: THBitBtn
        Left = 100
        Top = 2
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
        TabOrder = 2
        OnClick = BAideClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BFerme: THBitBtn
        Left = 68
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BImprimer: THBitBtn
        Left = 4
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnClick = BImprimerClick
        Margin = 2
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BValider: THBitBtn
        Left = 36
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Grand livre'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnClick = FlisteDblClick
        Margin = 2
        NumGlyphs = 2
        Spacing = -1
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
    end
  end
  object CBDebDat: TComboBox
    Left = 14
    Top = 126
    Width = 48
    Height = 21
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
  end
  object CBFinDate: TComboBox
    Left = 66
    Top = 126
    Width = 48
    Height = 21
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ItemHeight = 13
    ParentFont = False
    TabOrder = 3
    Visible = False
  end
  object Pcriteres: TPanel
    Left = 0
    Top = 0
    Width = 577
    Height = 60
    Align = alTop
    BevelInner = bvLowered
    TabOrder = 4
    object LExo: TLabel
      Left = 4
      Top = 11
      Width = 41
      Height = 13
      Caption = '&Exercice'
      FocusControl = CbExo
    end
    object LDevise: TLabel
      Left = 249
      Top = 11
      Width = 33
      Height = 13
      Caption = '&Devise'
      FocusControl = CbDevise
    end
    object LType: TLabel
      Left = 4
      Top = 37
      Width = 45
      Height = 13
      Caption = '&Type Mvt'
      FocusControl = CbTypMvt
    end
    object TEtab: THLabel
      Left = 249
      Top = 37
      Width = 34
      Height = 13
      Caption = 'Eta&blis.'
      FocusControl = CbEtab
    end
    object CbExo: THValComboBox
      Left = 51
      Top = 7
      Width = 194
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = CbExoChange
      TagDispatch = 0
      DataType = 'TTEXERCICE'
    end
    object CbDevise: THValComboBox
      Left = 291
      Top = 7
      Width = 150
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 1
      OnChange = CbDeviseChange
      TagDispatch = 0
      Vide = True
      DataType = 'TTDEVISE'
    end
    object CbSoldCum: TCheckBox
      Left = 444
      Top = 37
      Width = 92
      Height = 17
      Alignment = taLeftJustify
      Caption = '&Solde Cumulatif'
      TabOrder = 5
      OnClick = CbSoldCumClick
    end
    object CbTypMvt: THValComboBox
      Left = 51
      Top = 33
      Width = 194
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 3
      OnChange = CbTypMvtChange
      TagDispatch = 0
      DataType = 'TTQUALPIECECRIT'
    end
    object CbEtab: THValComboBox
      Left = 292
      Top = 33
      Width = 150
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 4
      OnChange = CbEtabChange
      TagDispatch = 0
      Vide = True
      DataType = 'TTETABLISSEMENT'
    end
    object Rev: TCheckBox
      Left = 444
      Top = 11
      Width = 92
      Height = 17
      Alignment = taLeftJustify
      AllowGrayed = True
      Caption = 'Ecr de &R'#233'vision'
      State = cbGrayed
      TabOrder = 2
    end
    object BChercher: THBitBtn
      Left = 541
      Top = 27
      Width = 28
      Height = 27
      Hint = 'Appliquer crit'#232'res'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 6
      OnClick = BChercherClick
      GlobalIndexImage = 'Z0217_S16G1'
    end
  end
  object Dexo2: TEdit
    Left = 164
    Top = 142
    Width = 33
    Height = 21
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 5
    Visible = False
  end
  object Dexo1: TEdit
    Left = 200
    Top = 142
    Width = 33
    Height = 21
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 6
    Visible = False
  end
  object BGraph: THBitBtn
    Left = 5
    Top = 348
    Width = 28
    Height = 27
    Hint = 'Graphique'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 7
    OnClick = BGraphClick
    Margin = 2
    GlobalIndexImage = 'Z2172_S16G1'
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Cumuls mensuels du compte g'#233'n'#233'ral: '
      'Cumuls mensuels du compte auxiliaire: '
      'Cumuls mensuels du journal: '
      'Cumuls mensuels de la section: ')
    Left = 512
    Top = 142
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 509
    Top = 110
  end
end

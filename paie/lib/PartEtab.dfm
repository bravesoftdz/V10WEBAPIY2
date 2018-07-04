object FCasPartEtab: TFCasPartEtab
  Left = 368
  Top = 148
  Width = 563
  Height = 341
  Caption = 'Cas Particuliers'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCloseQuery = FormCloseQuery
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object HLabel1: THLabel
    Left = 12
    Top = 20
    Width = 113
    Height = 13
    Caption = 'Profil ou Mod'#232'le Bulletin'
  end
  object HLabel2: THLabel
    Left = 296
    Top = 20
    Width = 89
    Height = 13
    Caption = 'P'#233'riodicit'#233'  Bulletin'
  end
  object HLabel3: THLabel
    Left = 28
    Top = 60
    Width = 71
    Height = 13
    Caption = 'Cas particuliers'
  end
  object Dock971: TDock97
    Left = 0
    Top = 271
    Width = 555
    Height = 36
    Position = dpBottom
    object ToolWindow971: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 544
      Caption = 'ToolWindow971'
      ClientAreaHeight = 32
      ClientAreaWidth = 544
      DockPos = 0
      TabOrder = 0
      object BVALIDER: TToolbarButton97
        Left = 449
        Top = 2
        Width = 27
        Height = 27
        Flat = False
        NumGlyphs = 2
        OnClick = BVALIDERClick
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object BFERMER: TToolbarButton97
        Left = 481
        Top = 2
        Width = 27
        Height = 27
        Flat = False
        ModalResult = 2
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BHELP: TToolbarButton97
        Left = 513
        Top = 2
        Width = 27
        Height = 27
        Flat = False
        GlobalIndexImage = 'Z1117_S16G1'
      end
    end
  end
  object PCas: THGrid
    Left = 12
    Top = 92
    Width = 529
    Height = 149
    ColCount = 2
    DefaultColWidth = 250
    DefaultRowHeight = 18
    TabOrder = 1
    SortedCol = -1
    Titres.Strings = (
      'Th'#232'mes'
      'Profil cas particulier')
    Couleur = False
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    OnRowEnter = PCasRowEnter
    ColCombo = 1
    ValCombo = PPART
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = 13224395
    RowHeights = (
      18
      18
      18
      18
      18)
  end
  object HDBValComboBox1: THDBValComboBox
    Left = 136
    Top = 16
    Width = 145
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 2
    TagDispatch = 0
  end
  object HDBValComboBox2: THDBValComboBox
    Left = 396
    Top = 16
    Width = 145
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 3
    TagDispatch = 0
  end
  object THEME: THValComboBox
    Left = 152
    Top = 56
    Width = 121
    Height = 21
    ItemHeight = 13
    TabOrder = 4
    Text = 'THEME'
    TagDispatch = 0
    DataType = 'PGTHEMEPROFIL'
  end
  object PPART: THValComboBox
    Left = 316
    Top = 52
    Width = 121
    Height = 21
    ItemHeight = 13
    TabOrder = 5
    Text = 'PARTICULIER'
    Visible = False
    TagDispatch = 0
    DataType = 'PGPROFILPARTICULIER'
    DisableTab = True
  end
  object DataSource1: TDataSource
    Left = 135
    Top = 254
  end
  object Table1: TTable
    DatabaseName = 'SOC'
    TableName = 'ETABLISS'
    Left = 207
    Top = 254
  end
  object Msg: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Voulez-vous sauvegarder les modifications ?;Q;YNC;Y;' +
        'C')
    Left = 272
    Top = 252
  end
end

object FCumuls: TFCumuls
  Left = 302
  Top = 208
  Width = 346
  Height = 378
  HelpContext = 7811100
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Param'#233'trage des constantes : '
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pappli: TPanel
    Left = 0
    Top = 66
    Width = 338
    Height = 250
    Align = alClient
    TabOrder = 1
    object DBnav: TDBNavigator
      Left = 92
      Top = 200
      Width = 80
      Height = 18
      VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast]
      ConfirmDelete = False
      TabOrder = 0
      Visible = False
    end
  end
  object Fliste: THGrid
    Left = 0
    Top = 66
    Width = 338
    Height = 250
    Align = alClient
    ColCount = 2
    DefaultColWidth = 200
    DefaultRowHeight = 18
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goTabs]
    TabOrder = 4
    OnKeyDown = FlisteKeyDown
    OnKeyPress = FlisteKeyPress
    SortedCol = -1
    Titres.Strings = (
      'P'#233'riode'
      'Valeur')
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = 13224395
    ColWidths = (
      158
      158)
  end
  object HPB: TPanel
    Left = 0
    Top = 316
    Width = 338
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
    object FAutoSave: TCheckBox
      Left = 240
      Top = 4
      Width = 29
      Height = 17
      Alignment = taLeftJustify
      Caption = 'Auto'
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      TabOrder = 0
      Visible = False
    end
    object Panel1: TPanel
      Left = 168
      Top = 2
      Width = 168
      Height = 31
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 1
      object BAnnuler: THBitBtn
        Left = 3
        Top = 2
        Width = 29
        Height = 27
        Hint = 'Annuler la derni'#232're action'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnClick = BAnnulerClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z0075_S16G1'
        IsControl = True
      end
      object BImprimer: THBitBtn
        Left = 36
        Top = 2
        Width = 29
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
        TabOrder = 1
        OnClick = BImprimerClick
        Margin = 2
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BValider: THBitBtn
        Left = 69
        Top = 2
        Width = 29
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
        TabOrder = 2
        OnClick = BValiderClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
      object BFerme: THBitBtn
        Left = 102
        Top = 2
        Width = 29
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: THBitBtn
        Left = 135
        Top = 2
        Width = 29
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
        TabOrder = 4
        OnClick = BAideClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
    end
  end
  object CU_SUITE: TDBCheckBox
    Left = 304
    Top = 44
    Width = 27
    Height = 17
    Alignment = taLeftJustify
    Caption = 'Suite'
    Color = clYellow
    DataField = 'CU_SUITE'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentColor = False
    ParentFont = False
    TabOrder = 2
    ValueChecked = 'X'
    ValueUnchecked = '-'
    Visible = False
  end
  object XX_WHERE: TPanel
    Left = 117
    Top = 251
    Width = 81
    Height = 33
    Caption = 'WHERE !!!'
    Color = clYellow
    TabOrder = 3
    Visible = False
  end
  object GbCar: TGroupBox
    Left = 0
    Top = 0
    Width = 338
    Height = 66
    Align = alTop
    Caption = ' Caract'#233'ristiques g'#233'n'#233'rales '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clActiveCaption
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 5
    object TCU_ETABLISSEMENT: THLabel
      Left = 36
      Top = 18
      Width = 65
      Height = 13
      Caption = 'E&tablissement'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TCU_EXERCICE: THLabel
      Left = 36
      Top = 44
      Width = 41
      Height = 13
      Caption = '&Exercice'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object CU_EXERCICE: THValComboBox
      Left = 106
      Top = 40
      Width = 195
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
      OnChange = CU_EXERCICEChange
      TagDispatch = 0
      DataType = 'TTEXERCICE'
    end
    object CU_ETABLISSEMENT: THValComboBox
      Left = 106
      Top = 14
      Width = 195
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
      OnChange = CU_ETABLISSEMENTChange
      TagDispatch = 0
      DataType = 'TTETABLISSEMENT'
    end
    object CU_COMPTE1: TEdit
      Left = 5
      Top = 12
      Width = 28
      Height = 21
      Color = clYellow
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      Text = 'CU_COMPTE1'
      Visible = False
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Param'#233'trage des constantes;Voulez-vous enregistrer les modific' +
        'ations?;Q;YNC;Y;C;'
      '1;'
      '2;'
      '3;'
      '4;'
      '5;'
      '6;'
      'L'#39'enregistrement est inaccessible'
      '8;'
      'A Nouveaux'
      'Valeur'
      
        '11;Param'#233'trage des constantes;Cet enregistrement n'#39'existe pas. D' +
        #233'sirez-vous le cr'#233'er?;Q;YNC;Y;C;'
      
        '12;Param'#233'trage des constantes;Vous devez renseigner l'#39#233'tablissem' +
        'ent.;W;O;O;O;'
      'Modification d'#39'enregistrement'
      'Cr'#233'ation d'#39'enregistrement'
      '15;')
    Left = 48
    Top = 164
  end
  object QCumuls: THQuery
    AutoCalcFields = False
    MarshalOptions = moMarshalModifiedOnly
    BeforePost = QCumulsBeforePost
    EnableBCD = False
    Parameters = <>
    SQL.Strings = (
      'Select * from CUMULS Where CU_TYPE="CON" order by CU_COMPTE1')
    dataBaseName = 'SOC'
    RequestLive = True
    MAJAuto = False
    Distinct = False
    Left = 292
    Top = 180
    object QCumulsCU_TYPE: TStringField
      FieldName = 'CU_TYPE'
      Size = 3
    end
    object QCumulsCU_COMPTE1: TStringField
      FieldName = 'CU_COMPTE1'
    end
    object QCumulsCU_COMPTE2: TStringField
      FieldName = 'CU_COMPTE2'
    end
    object QCumulsCU_EXERCICE: TStringField
      FieldName = 'CU_EXERCICE'
      Size = 3
    end
    object QCumulsCU_SUITE: TStringField
      FieldName = 'CU_SUITE'
      Size = 1
    end
    object QCumulsCU_ETABLISSEMENT: TStringField
      FieldName = 'CU_ETABLISSEMENT'
      Size = 3
    end
    object QCumulsCU_DEVQTE: TStringField
      FieldName = 'CU_DEVQTE'
      Size = 3
    end
    object QCumulsCU_QUALIFPIECE: TStringField
      FieldName = 'CU_QUALIFPIECE'
      Size = 3
    end
    object QCumulsCU_DEBIT1: TFloatField
      FieldName = 'CU_DEBIT1'
    end
    object QCumulsCU_DEBIT2: TFloatField
      FieldName = 'CU_DEBIT2'
    end
    object QCumulsCU_DEBIT3: TFloatField
      FieldName = 'CU_DEBIT3'
    end
    object QCumulsCU_DEBIT4: TFloatField
      FieldName = 'CU_DEBIT4'
    end
    object QCumulsCU_DEBIT5: TFloatField
      FieldName = 'CU_DEBIT5'
    end
    object QCumulsCU_DEBIT6: TFloatField
      FieldName = 'CU_DEBIT6'
    end
    object QCumulsCU_DEBIT7: TFloatField
      FieldName = 'CU_DEBIT7'
    end
    object QCumulsCU_DEBIT8: TFloatField
      FieldName = 'CU_DEBIT8'
    end
    object QCumulsCU_DEBIT9: TFloatField
      FieldName = 'CU_DEBIT9'
    end
    object QCumulsCU_DEBIT10: TFloatField
      FieldName = 'CU_DEBIT10'
    end
    object QCumulsCU_DEBIT11: TFloatField
      FieldName = 'CU_DEBIT11'
    end
    object QCumulsCU_DEBIT12: TFloatField
      FieldName = 'CU_DEBIT12'
    end
    object QCumulsCU_CREDIT1: TFloatField
      FieldName = 'CU_CREDIT1'
    end
    object QCumulsCU_CREDIT2: TFloatField
      FieldName = 'CU_CREDIT2'
    end
    object QCumulsCU_CREDIT3: TFloatField
      FieldName = 'CU_CREDIT3'
    end
    object QCumulsCU_CREDIT4: TFloatField
      FieldName = 'CU_CREDIT4'
    end
    object QCumulsCU_CREDIT5: TFloatField
      FieldName = 'CU_CREDIT5'
    end
    object QCumulsCU_CREDIT6: TFloatField
      FieldName = 'CU_CREDIT6'
    end
    object QCumulsCU_CREDIT7: TFloatField
      FieldName = 'CU_CREDIT7'
    end
    object QCumulsCU_CREDIT8: TFloatField
      FieldName = 'CU_CREDIT8'
    end
    object QCumulsCU_CREDIT9: TFloatField
      FieldName = 'CU_CREDIT9'
    end
    object QCumulsCU_CREDIT10: TFloatField
      FieldName = 'CU_CREDIT10'
    end
    object QCumulsCU_CREDIT11: TFloatField
      FieldName = 'CU_CREDIT11'
    end
    object QCumulsCU_CREDIT12: TFloatField
      FieldName = 'CU_CREDIT12'
    end
    object QCumulsCU_CREDITAN: TFloatField
      FieldName = 'CU_CREDITAN'
    end
    object QCumulsCU_DEBITAN: TFloatField
      FieldName = 'CU_DEBITAN'
    end
    object QCumulsCU_QUALIFQTEDIFF: TStringField
      FieldName = 'CU_QUALIFQTEDIFF'
      Size = 1
    end
  end
  object QSuite: THQuery
    AutoCalcFields = False
    LockType = ltReadOnly
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    Parameters = <>
    dataBaseName = 'SOC'
    MAJAuto = False
    Distinct = False
    Left = 236
    Top = 208
  end
  object HMTrad: THSystemMenu
    MaxInsideSize = False
    ActiveResize = False
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 48
    Top = 264
  end
end

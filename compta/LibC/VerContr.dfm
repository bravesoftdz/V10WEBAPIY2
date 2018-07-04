object FVerContr: TFVerContr
  Left = 368
  Top = 180
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Mise '#224' jour des contreparties'
  ClientHeight = 166
  ClientWidth = 391
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
  object TE_EXERCICE: THLabel
    Left = 8
    Top = 16
    Width = 41
    Height = 13
    Caption = '&Exercice'
    FocusControl = E_EXERCICE
  end
  object TE_DATECOMPTABLE: THLabel
    Left = 8
    Top = 44
    Width = 100
    Height = 13
    Caption = '&Dates comptables du'
    FocusControl = E_DATECOMPTABLE
  end
  object TE_DATECOMPTABLE2: THLabel
    Left = 248
    Top = 44
    Width = 12
    Height = 13
    Caption = 'au'
    FocusControl = E_DATECOMPTABLE_
  end
  object TE_NATUREPIECE: THLabel
    Left = 8
    Top = 100
    Width = 32
    Height = 13
    Caption = '&Nature'
    FocusControl = E_NATUREPIECE
  end
  object Label1: THLabel
    Tag = 10
    Left = 248
    Top = 72
    Width = 12
    Height = 13
    Caption = 'au'
    FocusControl = E_JOURNAL_
  end
  object ENATUREGENE: THLabel
    Tag = 10
    Left = 8
    Top = 72
    Width = 49
    Height = 13
    Caption = '&Journal de'
    FocusControl = E_JOURNAL
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object HPB: TPanel
    Left = 0
    Top = 131
    Width = 391
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    Caption = ' '
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
    object PBoutons: TPanel
      Left = 288
      Top = 2
      Width = 101
      Height = 31
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 0
      object BValider: THBitBtn
        Left = 38
        Top = 2
        Width = 29
        Height = 27
        Hint = 'Lancer la v'#233'rification'
        Default = True
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
        Spacing = -1
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
      object BFerme: THBitBtn
        Left = 70
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
    end
  end
  object E_EXERCICE: THValComboBox
    Left = 124
    Top = 12
    Width = 257
    Height = 21
    Style = csDropDownList
    Ctl3D = True
    ItemHeight = 13
    ParentCtl3D = False
    TabOrder = 1
    OnChange = E_EXERCICEChange
    TagDispatch = 0
    Vide = True
    DataType = 'TTEXERCICE'
  end
  object E_DATECOMPTABLE: THCritMaskEdit
    Left = 124
    Top = 40
    Width = 105
    Height = 21
    Ctl3D = True
    EditMask = '!99/99/0000;1;_'
    MaxLength = 10
    ParentCtl3D = False
    TabOrder = 2
    Text = '  /  /    '
    TagDispatch = 0
    Operateur = Superieur
    OpeType = otDate
    ElipsisButton = True
    ControlerDate = True
  end
  object E_DATECOMPTABLE_: THCritMaskEdit
    Left = 276
    Top = 40
    Width = 105
    Height = 21
    Ctl3D = True
    EditMask = '!99/99/0000;1;_'
    MaxLength = 10
    ParentCtl3D = False
    TabOrder = 3
    Text = '  /  /    '
    TagDispatch = 0
    Operateur = Inferieur
    OpeType = otDate
    ElipsisButton = True
    ControlerDate = True
  end
  object E_NATUREPIECE: THValComboBox
    Left = 124
    Top = 96
    Width = 257
    Height = 21
    Style = csDropDownList
    Ctl3D = True
    ItemHeight = 13
    ParentCtl3D = False
    TabOrder = 4
    TagDispatch = 0
    Vide = True
    DataType = 'TTNATUREPIECE'
  end
  object E_JOURNAL: THValComboBox
    Tag = 10
    Left = 124
    Top = 68
    Width = 106
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 5
    TagDispatch = 0
    DataType = 'CPJOURNAL'
  end
  object E_JOURNAL_: THValComboBox
    Tag = 10
    Left = 276
    Top = 68
    Width = 106
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 6
    TagDispatch = 0
    DataType = 'CPJOURNAL'
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 80
    Top = 112
  end
  object Q: TQuery
    AutoCalcFields = False
    DatabaseName = 'SOC'
    Left = 72
    Top = 12
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?Caption?;Confirmez-vous la mise '#224' jour des contreparties ?;Q;' +
        'YN;N;N;'
      '1;?Caption?;Le traitement a '#233't'#233' effectu'#233' !;A;O;O;O;'
      
        '2;?Caption?;Le traitement de certaines pi'#232'ces n'#39'a pas pu '#234'tre ef' +
        'fectu'#233' ! D'#233'sirez-vous les traiter manuellement ?;Q;YN;N;N;'
      '')
    Left = 60
    Top = 72
  end
end

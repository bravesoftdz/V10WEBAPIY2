object FTVAEdite: TFTVAEdite
  Left = 234
  Top = 206
  Width = 360
  Height = 196
  BorderIcons = [biSystemMenu]
  Caption = 'Edition des '#233'tats de TVA sur encaissements'
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
  object Label7: TLabel
    Left = 224
    Top = 52
    Width = 12
    Height = 13
    Caption = 'au'
    FocusControl = FDateCompta2
  end
  object HLabel4: THLabel
    Left = 8
    Top = 24
    Width = 113
    Height = 13
    AutoSize = False
    Caption = '&Exercice'
    FocusControl = FExercice
  end
  object HLabel6: THLabel
    Left = 8
    Top = 52
    Width = 109
    Height = 13
    AutoSize = False
    Caption = '&Dates comptables du'
    FocusControl = FDateCompta1
  end
  object HPB: TPanel
    Left = 0
    Top = 134
    Width = 352
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
      Left = 251
      Top = 2
      Width = 99
      Height = 31
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 0
      object BAide: THBitBtn
        Left = 67
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
        TabOrder = 0
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BFerme: THBitBtn
        Left = 35
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BValider: THBitBtn
        Left = 4
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Valider'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ModalResult = 1
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
  object FDateCompta2: TMaskEdit
    Left = 260
    Top = 48
    Width = 81
    Height = 21
    Ctl3D = True
    EditMask = '!00/00/0000;1;_'
    MaxLength = 10
    ParentCtl3D = False
    TabOrder = 1
    Text = '  /  /    '
    OnKeyPress = FDateCompta2KeyPress
  end
  object FDateCompta1: TMaskEdit
    Left = 124
    Top = 48
    Width = 81
    Height = 21
    Ctl3D = True
    EditMask = '!00/00/0000;1;_'
    MaxLength = 10
    ParentCtl3D = False
    TabOrder = 2
    Text = '  /  /    '
    OnKeyPress = FDateCompta1KeyPress
  end
  object FExercice: THValComboBox
    Left = 124
    Top = 20
    Width = 218
    Height = 21
    Style = csDropDownList
    Ctl3D = True
    ItemHeight = 13
    ParentCtl3D = False
    TabOrder = 3
    OnChange = FExerciceChange
    TagDispatch = 0
    DataType = 'TTEXERCICE'
  end
  object TypeTVA: TRadioGroup
    Left = 9
    Top = 80
    Width = 336
    Height = 41
    Caption = ' TVA '
    Columns = 2
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ItemIndex = 0
    Items.Strings = (
      '&Collect'#233'e'
      '&D'#233'ductible')
    ParentFont = False
    TabOrder = 4
    OnClick = TypeTVAClick
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 104
    Top = 106
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Edition des '#233'tats de TVA sur encaissements'
      'Edition des '#233'tats de TVA sur d'#233'caissements'
      
        '2;?Caption?;Confirmez-vous le traitement de mise '#224' jour ?;Q;YNC;' +
        'Y;Y;'
      '3;?Caption?;Le traitement a '#233't'#233' effectu'#233' avec succ'#232's.;A;O;O;O;')
    Left = 288
    Top = 86
  end
end

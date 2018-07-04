object FVerCai: TFVerCai
  Left = 253
  Top = 212
  Width = 265
  Height = 154
  HelpContext = 7745200
  BorderIcons = [biSystemMenu]
  Caption = 'Contr'#244'le de caisse'
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
  object HPB: TPanel
    Left = 0
    Top = 93
    Width = 257
    Height = 34
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
    object BValider: THBitBtn
      Left = 162
      Top = 4
      Width = 28
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
      Left = 193
      Top = 4
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
    object BStop: THBitBtn
      Left = 131
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Arr'#234'te la v'#233'rification'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = BStopClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0107_S16G1'
      IsControl = True
    end
    object BAide: THBitBtn
      Left = 224
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Aide'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      OnClick = BAideClick
      GlobalIndexImage = 'Z1117_S16G1'
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 257
    Height = 93
    Align = alClient
    TabOrder = 1
    object TFGene: TLabel
      Left = 6
      Top = 12
      Width = 74
      Height = 13
      Caption = 'Compte g'#233'n'#233'ral'
    end
    object TFPiece: TLabel
      Left = 6
      Top = 62
      Width = 40
      Height = 13
      Caption = 'Pi'#232'ce n'#176
    end
    object TFLigne: TLabel
      Left = 164
      Top = 62
      Width = 39
      Height = 13
      Caption = 'Ligne n'#176
    end
    object Label1: TLabel
      Left = 6
      Top = 37
      Width = 90
      Height = 13
      Caption = 'Date comptable au'
    end
    object Bevel1: TBevel
      Left = 100
      Top = 8
      Width = 152
      Height = 22
    end
    object Bevel2: TBevel
      Left = 100
      Top = 34
      Width = 152
      Height = 22
    end
    object Bevel3: TBevel
      Left = 100
      Top = 59
      Width = 57
      Height = 22
    end
    object Bevel4: TBevel
      Left = 206
      Top = 59
      Width = 46
      Height = 22
    end
    object FGene: TEdit
      Left = 104
      Top = 12
      Width = 144
      Height = 14
      BorderStyle = bsNone
      Color = clBtnFace
      Ctl3D = False
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 0
      Text = '123456789'
    end
    object FPiece: TEdit
      Left = 106
      Top = 62
      Width = 41
      Height = 17
      BorderStyle = bsNone
      Color = clBtnFace
      Ctl3D = False
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 1
      Text = 'gfcjfjf'
    end
    object FLigne: TEdit
      Left = 209
      Top = 62
      Width = 41
      Height = 17
      BorderStyle = bsNone
      Color = clBtnFace
      Ctl3D = False
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 2
      Text = 'jfhjjgfcj'
    end
    object FDateCpta: TEdit
      Left = 105
      Top = 38
      Width = 143
      Height = 14
      BorderStyle = bsNone
      Color = clBtnFace
      Ctl3D = False
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 3
      Text = 'jnfjfj'
    end
    object PourClotureOk: TCheckBox
      Left = 128
      Top = 84
      Width = 97
      Height = 17
      AllowGrayed = True
      Caption = 'PourClotureOk'
      Color = clYellow
      ParentColor = False
      TabOrder = 4
      Visible = False
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 44
    Top = 88
  end
  object MsgRien: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Contr'#244'le de caisse;Confirmez-vous l'#39'arr'#234't de la v'#233'rification e' +
        'n cours ?;Q;YN;N;N;'
      '1;Contr'#244'le de caisse;Nb de compte de caisse cr'#233'diteur;W;O;O;O;'
      '2;?caption?;Il n'#39'y a pas de compte de caisse cr'#233'diteur;E;O;O;O;')
    Left = 16
    Top = 87
  end
  object MError: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Contr'#244'le de caisse;Confirmez-vous l'#39'arr'#234't de la v'#233'rification e' +
        'n cours ?;Q;YN;N;N;')
    Left = 80
    Top = 87
  end
end

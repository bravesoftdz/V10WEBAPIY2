object FSaisEnc: TFSaisEnc
  Left = 325
  Top = 268
  HelpContext = 7646200
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'El'#233'ments de TVA'
  ClientHeight = 256
  ClientWidth = 412
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
  object POutils: TPanel
    Left = 0
    Top = 223
    Width = 412
    Height = 33
    Align = alBottom
    BevelInner = bvLowered
    BevelOuter = bvNone
    TabOrder = 0
    object BValide: THBitBtn
      Left = 314
      Top = 3
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
      OnClick = BValideClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0127_S16G1'
      IsControl = True
    end
    object BAbandon: THBitBtn
      Left = 346
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BAbandonClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0021_S16G1'
      IsControl = True
    end
    object BAide: THBitBtn
      Left = 378
      Top = 3
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
  end
  object PEntete: TPanel
    Left = 0
    Top = 0
    Width = 412
    Height = 85
    Align = alTop
    Enabled = False
    TabOrder = 1
    object H_REGIMETVA: THLabel
      Left = 12
      Top = 32
      Width = 63
      Height = 13
      Caption = '&R'#233'gime fiscal'
      FocusControl = E_REGIMETVA
    end
    object H_GENERAL: THLabel
      Left = 12
      Top = 60
      Width = 63
      Height = 13
      Caption = '&Compte  H.T.'
      FocusControl = E_GENERAL
    end
    object G_LIBELLE: THLabel
      Left = 224
      Top = 60
      Width = 181
      Height = 13
      AutoSize = False
      Caption = 'G_LIBELLE'
    end
    object Label1: TLabel
      Left = 12
      Top = 6
      Width = 385
      Height = 18
      Alignment = taCenter
      AutoSize = False
      Caption = 'Informations ligne H.T.'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object E_REGIMETVA: THValComboBox
      Left = 88
      Top = 28
      Width = 129
      Height = 22
      Style = csSimple
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ItemHeight = 13
      ParentFont = False
      TabOrder = 0
      TagDispatch = 0
      DataType = 'TTREGIMETVA'
    end
    object E_GENERAL: THCpteEdit
      Left = 88
      Top = 56
      Width = 129
      Height = 21
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 1
      ZoomTable = tzGeneral
      Vide = False
      Bourre = False
      Libelle = G_LIBELLE
      okLocate = False
      SynJoker = False
    end
  end
  object PTva: TPanel
    Left = 0
    Top = 85
    Width = 412
    Height = 76
    Align = alTop
    TabOrder = 2
    object H_TVA: THLabel
      Left = 224
      Top = 42
      Width = 49
      Height = 13
      Caption = 'Code &TVA'
      FocusControl = E_TVA
    end
    object Label2: TLabel
      Left = 12
      Top = 6
      Width = 385
      Height = 16
      Alignment = taCenter
      AutoSize = False
      Caption = 'Caract'#233'ristiques TVA modifiables'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object RModeTVA: TRadioGroup
      Left = 4
      Top = 24
      Width = 213
      Height = 39
      Caption = 'Exigibilit'#233' de la TVA'
      Color = clBtnFace
      Columns = 2
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Items.Strings = (
        '&D'#233'bits'
        '&Encaissements')
      ParentColor = False
      ParentFont = False
      TabOrder = 0
      OnClick = DeduitTva
    end
    object E_TVA: THValComboBox
      Left = 280
      Top = 38
      Width = 121
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 1
      OnChange = DeduitTva
      TagDispatch = 0
      DataType = 'TTTVA'
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 161
    Width = 412
    Height = 62
    Align = alClient
    Enabled = False
    TabOrder = 3
    object Label3: TLabel
      Left = 12
      Top = 6
      Width = 385
      Height = 16
      Alignment = taCenter
      AutoSize = False
      Caption = 'Informations th'#233'oriques d'#233'duites du param'#233'trage'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object H_CPTTVA: THLabel
      Left = 12
      Top = 32
      Width = 60
      Height = 13
      Caption = 'Compte TVA'
      FocusControl = E_GENERAL
    end
    object H_TXTVA: THLabel
      Left = 224
      Top = 32
      Width = 48
      Height = 13
      Caption = 'Taux TVA'
      FocusControl = E_GENERAL
    end
    object CPTTVA: THCpteEdit
      Left = 88
      Top = 28
      Width = 129
      Height = 21
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
      ZoomTable = tzGeneral
      Vide = False
      Bourre = False
      okLocate = False
      SynJoker = False
    end
    object TXTVA: THNumEdit
      Left = 280
      Top = 28
      Width = 106
      Height = 21
      Ctl3D = True
      Decimals = 2
      Digits = 12
      Masks.PositiveMask = '#,##0'
      Debit = False
      NumericType = ntPercentage
      ParentCtl3D = False
      TabOrder = 1
      UseRounding = True
      Validate = False
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 368
    Top = 12
  end
end

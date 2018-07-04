object FLettreRegu: TFLettreRegu
  Left = 273
  Top = 172
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = 'R'#233'gularisation de lettrage'
  ClientHeight = 145
  ClientWidth = 371
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
  object Label1: TLabel
    Left = 4
    Top = 4
    Width = 361
    Height = 16
    Alignment = taCenter
    AutoSize = False
    Caption = 'Choix d'#39'un compte de r'#233'gularisation'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object BValide: THBitBtn
    Tag = 1
    Left = 307
    Top = 116
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
    TabOrder = 3
    OnClick = BValideClick
    Margin = 2
    Spacing = -1
    GlobalIndexImage = 'Z0127_S16G1'
    IsControl = True
  end
  object BAbandon: THBitBtn
    Tag = 1
    Left = 339
    Top = 116
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
    TabOrder = 4
    Margin = 2
    Spacing = -1
    GlobalIndexImage = 'Z0021_S16G1'
    IsControl = True
  end
  object Panel2: TPanel
    Left = 4
    Top = 114
    Width = 257
    Height = 27
    Enabled = False
    TabOrder = 2
    object H_Solde: TLabel
      Left = 28
      Top = 6
      Width = 113
      Height = 13
      Caption = 'Diff'#233'rence de r'#232'glement'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object E_SOLDE: THNumEdit
      Left = 156
      Top = 3
      Width = 89
      Height = 21
      Color = clBtnFace
      Decimals = 2
      Digits = 12
      Enabled = False
      Masks.PositiveMask = '#,##0'
      Debit = False
      TabOrder = 0
      UseRounding = True
      Validate = False
    end
  end
  object GroupBox1: TGroupBox
    Left = 36
    Top = 20
    Width = 329
    Height = 89
    Enabled = False
    TabOrder = 1
    object H3: THLabel
      Left = 128
      Top = 65
      Width = 197
      Height = 13
      AutoSize = False
      Caption = 'H3'
    end
    object H2: THLabel
      Left = 128
      Top = 41
      Width = 197
      Height = 13
      AutoSize = False
      Caption = 'H2'
    end
    object H1: THLabel
      Left = 128
      Top = 17
      Width = 197
      Height = 13
      AutoSize = False
      Caption = 'H1'
    end
    object C1: THCpteEdit
      Left = 4
      Top = 13
      Width = 121
      Height = 21
      TabOrder = 0
      Text = 'C1'
      ZoomTable = tzGeneral
      Vide = False
      Bourre = False
      Libelle = H1
      okLocate = False
      SynJoker = False
    end
    object C2: THCpteEdit
      Left = 4
      Top = 37
      Width = 121
      Height = 21
      TabOrder = 1
      Text = 'C2'
      ZoomTable = tzGeneral
      Vide = False
      Bourre = False
      Libelle = H2
      okLocate = False
      SynJoker = False
    end
    object C3: THCpteEdit
      Left = 4
      Top = 61
      Width = 121
      Height = 21
      TabOrder = 2
      Text = 'C3'
      ZoomTable = tzGeneral
      Vide = False
      Bourre = False
      Libelle = H3
      okLocate = False
      SynJoker = False
    end
  end
  object R: TRadioGroup
    Left = 4
    Top = 20
    Width = 29
    Height = 87
    ItemIndex = 0
    Items.Strings = (
      ''
      ''
      '')
    TabOrder = 0
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 272
    Top = 48
  end
end

object FCleRepBu: TFCleRepBu
  Left = 205
  Top = 169
  Width = 331
  Height = 148
  HelpContext = 15217200
  BorderIcons = [biSystemMenu]
  Caption = 'Choix de la cl'#233' de r'#233'partition budg'#233'taire'
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
  object HLabel1: THLabel
    Left = 12
    Top = 12
    Width = 79
    Height = 13
    Caption = 'Cl'#233' de r'#233'partition'
    FocusControl = FCleRep
  end
  object HLabel2: THLabel
    Left = 12
    Top = 40
    Width = 102
    Height = 13
    Caption = 'Colonne de r'#233'f'#233'rence'
    FocusControl = FColRef
  end
  object Pied: TPanel
    Left = 0
    Top = 88
    Width = 323
    Height = 33
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 0
    object BValide: THBitBtn
      Tag = 1
      Left = 226
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Valider'
      Default = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ModalResult = 1
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BValideClick
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      GlobalIndexImage = 'Z0003_S16G2'
      IsControl = True
    end
    object BAbandon: THBitBtn
      Tag = 1
      Left = 258
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
    object BAide: THBitBtn
      Tag = 1
      Left = 290
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
      TabOrder = 3
      OnClick = BAideClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z1117_S16G1'
      IsControl = True
    end
    object BZoom: THBitBtn
      Tag = 1
      Left = 193
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Voir la cl'#233' de r'#233'partition'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BZoomClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0016_S16G1'
      IsControl = True
    end
  end
  object FCleRep: THValComboBox
    Left = 124
    Top = 8
    Width = 197
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 1
    TagDispatch = 0
  end
  object FColRef: THValComboBox
    Left = 124
    Top = 36
    Width = 197
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 2
    TagDispatch = 0
  end
  object CCourant: TCheckBox
    Left = 10
    Top = 64
    Width = 309
    Height = 17
    Alignment = taLeftJustify
    Caption = 'R'#233'partition par p'#233'riode sur la ligne en cours seulement'
    Checked = True
    State = cbChecked
    TabOrder = 3
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 180
    Top = 18
  end
end

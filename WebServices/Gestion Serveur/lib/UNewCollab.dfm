object fNewID: TfNewID
  Left = 393
  Top = 316
  Width = 505
  Height = 247
  Caption = 'Affectation des appareils'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object LID: THLabel
    Left = 8
    Top = 18
    Width = 120
    Height = 13
    Caption = 'Identificateur de l'#39'appareil'
  end
  object LLIBELLE: THLabel
    Left = 8
    Top = 55
    Width = 56
    Height = 13
    Caption = 'D'#233'signation'
  end
  object LSERVER: THLabel
    Left = 8
    Top = 78
    Width = 116
    Height = 13
    Caption = 'Nom du serveur de BDD'
  end
  object HLabel1: THLabel
    Left = 8
    Top = 102
    Width = 78
    Height = 13
    Caption = 'Base de donn'#233'e'
  end
  object LRESSOURCE: THLabel
    Left = 8
    Top = 126
    Width = 96
    Height = 13
    Caption = 'Ressource associ'#233'e'
  end
  object LPASSWORD: THLabel
    Left = 8
    Top = 150
    Width = 64
    Height = 13
    Caption = 'Mot de passe'
  end
  object PBAS: TPanel
    Left = 0
    Top = 173
    Width = 489
    Height = 36
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      489
      36)
    object BFERME: TToolbarButton97
      Left = 456
      Top = 4
      Width = 27
      Height = 27
      Anchors = [akRight, akBottom]
      Opaque = False
      OnClick = BFERMEClick
      GlobalIndexImage = 'Z0021_S24G1'
    end
    object BVALIDE: TToolbarButton97
      Left = 427
      Top = 4
      Width = 27
      Height = 27
      Anchors = [akRight, akBottom]
      Opaque = False
      OnClick = BVALIDEClick
      GlobalIndexImage = 'Z0127_S24G1'
    end
    object BDELETE: TToolbarButton97
      Left = 24
      Top = 3
      Width = 27
      Height = 27
      Anchors = [akLeft, akBottom]
      Opaque = False
      Visible = False
      OnClick = BDELETEClick
      GlobalIndexImage = 'Z0005_S24G1'
    end
  end
  object ID: TEdit97
    Left = 151
    Top = 13
    Width = 177
    Height = 22
    CharCase = ecUpperCase
    MaxLength = 50
    TabOrder = 1
    Text = 'ID'
    OnExit = IDExit
  end
  object LIBELLE: TEdit97
    Left = 151
    Top = 50
    Width = 321
    Height = 22
    MaxLength = 120
    TabOrder = 2
  end
  object MRESSOURCE: THValComboBox
    Left = 153
    Top = 122
    Width = 213
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 3
    TagDispatch = 0
    Vide = True
    VideString = '<<Aucun>>'
  end
  object USER: TEdit97
    Left = 364
    Top = 121
    Width = 93
    Height = 22
    MaxLength = 35
    TabOrder = 5
    Visible = False
  end
  object SERVER: THValComboBox
    Left = 153
    Top = 74
    Width = 213
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 6
    TagDispatch = 0
    Vide = True
    VideString = '<<Aucun>>'
  end
  object DATABASE: THValComboBox
    Left = 153
    Top = 98
    Width = 213
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 7
    TagDispatch = 0
    Vide = True
    VideString = '<<Aucun>>'
  end
  object RESSOURCE: TEdit97
    Left = 364
    Top = 97
    Width = 93
    Height = 22
    MaxLength = 35
    TabOrder = 8
    Visible = False
  end
  object PASSWORD: TEdit97
    Left = 151
    Top = 145
    Width = 213
    Height = 21
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Arial Rounded MT Bold'
    Font.Style = []
    MaxLength = 35
    ParentFont = False
    PasswordChar = '*'
    TabOrder = 4
  end
  object HmTrad: THSystemMenu
    Separator = True
    Traduction = False
    Left = 424
    Top = 8
  end
end

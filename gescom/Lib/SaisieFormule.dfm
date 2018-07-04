object FSaisieFormule: TFSaisieFormule
  Left = 270
  Top = 171
  Width = 483
  Height = 341
  Caption = 'Saisie d'#39'une formule de calcul'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  ShowHint = True
  OnClose = FormClose
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 12
    Top = 3
    Width = 37
    Height = 13
    Caption = 'Formule'
  end
  object Label2: TLabel
    Left = 12
    Top = 49
    Width = 96
    Height = 13
    Caption = 'Table de destination'
  end
  object Label3: TLabel
    Left = 160
    Top = 49
    Width = 90
    Height = 13
    Caption = 'Fichier de donn'#233'es'
  end
  object HLabel1: THLabel
    Left = 308
    Top = 184
    Width = 101
    Height = 13
    Caption = 'Fonctions disponibles'
  end
  object ListeTable: TListBox
    Left = 12
    Top = 69
    Width = 140
    Height = 231
    ItemHeight = 13
    TabOrder = 0
    OnDblClick = ListeTableDblClick
  end
  object ListeFichier: TListBox
    Left = 160
    Top = 69
    Width = 140
    Height = 231
    ItemHeight = 13
    TabOrder = 1
    OnDblClick = ListeFichierDblClick
  end
  object Edit1: TEdit
    Left = 12
    Top = 23
    Width = 446
    Height = 21
    ParentShowHint = False
    ShowHint = True
    TabOrder = 2
    Text = 'Edit1'
  end
  object BitBtn1: THBitBtn
    Left = 308
    Top = 69
    Width = 23
    Height = 23
    Caption = 'BitBtn1'
    TabOrder = 3
    OnClick = BitBtn
    Margin = 0
    GlobalIndexImage = 'Z0248_S16G1'
  end
  object BitBtn2: THBitBtn
    Left = 336
    Top = 69
    Width = 23
    Height = 23
    Caption = 'BitBtn1'
    TabOrder = 4
    OnClick = BitBtn
    Margin = 0
    GlobalIndexImage = 'Z0300_S16G1'
  end
  object BitBtn3: THBitBtn
    Left = 364
    Top = 69
    Width = 23
    Height = 23
    Caption = 'BitBtn1'
    TabOrder = 5
    OnClick = BitBtn
    Margin = 0
    GlobalIndexImage = 'Z0269_S16G1'
  end
  object BitBtn4: THBitBtn
    Left = 308
    Top = 97
    Width = 23
    Height = 23
    Caption = 'BitBtn1'
    TabOrder = 6
    OnClick = BitBtn
    Margin = 0
    GlobalIndexImage = 'Z0315_S16G1'
  end
  object BitBtn5: THBitBtn
    Left = 336
    Top = 97
    Width = 23
    Height = 23
    Caption = 'BitBtn1'
    TabOrder = 7
    OnClick = BitBtn
    Margin = 0
    GlobalIndexImage = 'Z0292_S16G1'
  end
  object BitBtn6: THBitBtn
    Left = 364
    Top = 97
    Width = 23
    Height = 23
    Caption = 'BitBtn1'
    TabOrder = 8
    OnClick = BitBtn
    Margin = 0
    GlobalIndexImage = 'Z0265_S16G1'
  end
  object BitBtn7: THBitBtn
    Left = 308
    Top = 125
    Width = 23
    Height = 23
    Caption = 'BitBtn1'
    TabOrder = 9
    OnClick = BitBtn
    Margin = 0
    GlobalIndexImage = 'Z0247_S16G1'
  end
  object BitBtn8: THBitBtn
    Left = 336
    Top = 125
    Width = 23
    Height = 23
    Caption = 'BitBtn1'
    TabOrder = 10
    OnClick = BitBtn
    Margin = 0
    GlobalIndexImage = 'Z0254_S16G1'
  end
  object BitBtn9: THBitBtn
    Left = 364
    Top = 125
    Width = 23
    Height = 23
    Caption = 'BitBtn1'
    TabOrder = 11
    OnClick = BitBtn
    Margin = 0
    GlobalIndexImage = 'Z0277_S16G1'
  end
  object BitBtn10: THBitBtn
    Left = 336
    Top = 153
    Width = 23
    Height = 23
    Caption = 'BitBtn1'
    TabOrder = 12
    OnClick = BitBtn
    Margin = 0
    GlobalIndexImage = 'Z0314_S16G1'
  end
  object BitBtn11: THBitBtn
    Left = 400
    Top = 69
    Width = 23
    Height = 23
    Caption = 'BitBtn1'
    TabOrder = 13
    OnClick = BitBtn
    Margin = 0
    GlobalIndexImage = 'Z0246_S16G1'
  end
  object BitBtn12: THBitBtn
    Left = 400
    Top = 97
    Width = 23
    Height = 23
    Caption = 'BitBtn1'
    TabOrder = 14
    OnClick = BitBtn
    Margin = 0
    GlobalIndexImage = 'Z0261_S16G1'
  end
  object BitBtn13: THBitBtn
    Left = 400
    Top = 125
    Width = 23
    Height = 23
    Caption = 'BitBtn1'
    TabOrder = 15
    OnClick = BitBtn
    Margin = 0
    GlobalIndexImage = 'Z0290_S16G1'
  end
  object BitBtn14: THBitBtn
    Left = 400
    Top = 153
    Width = 23
    Height = 23
    Caption = 'BitBtn1'
    TabOrder = 16
    OnClick = BitBtn
    Margin = 0
    GlobalIndexImage = 'Z0251_S16G1'
  end
  object BitBtn15: THBitBtn
    Left = 436
    Top = 69
    Width = 23
    Height = 23
    Caption = 'BitBtn1'
    TabOrder = 17
    OnClick = BitBtn
    Margin = 0
    GlobalIndexImage = 'Z0267_S16G1'
  end
  object FctsDisp: THValComboBox
    Left = 308
    Top = 202
    Width = 145
    Height = 21
    ItemHeight = 13
    ParentShowHint = False
    ShowHint = False
    TabOrder = 18
    OnChange = FctsDispChange
    Items.Strings = (
      'Compteur'
      'Remplace caract'#232'res'
      'Correspondance'
      'Majuscule'
      'Table'
      'Sous Chaine'
      'Espaces'
      'Cha'#238'ne -> Entier'
      'Entier -> Cha'#238'ne'
      'Date')
    TagDispatch = 0
    Values.Strings = (
      'COMPTEUR'
      'REPLACECAR'
      'CORRESPOND'
      'MAJUSCULE'
      'TABLE'
      'SOUSCHAINE'
      'SPACE'
      'STRINT'
      'INTSTR'
      'DATE')
  end
  object ListeTableCode: TListBox
    Left = 16
    Top = 73
    Width = 140
    Height = 231
    ItemHeight = 13
    TabOrder = 19
    Visible = False
    OnDblClick = ListeTableDblClick
  end
  object ListeFichierCode: TListBox
    Left = 164
    Top = 73
    Width = 140
    Height = 231
    ItemHeight = 13
    TabOrder = 20
    Visible = False
    OnDblClick = ListeFichierDblClick
  end
end

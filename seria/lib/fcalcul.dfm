object calcul: Tcalcul
  Left = 667
  Top = 246
  Width = 270
  Height = 335
  Caption = 'Calcul de la clef d'#39'activation'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 62
    Width = 60
    Height = 13
    Caption = 'Code produit'
  end
  object Label2: TLabel
    Left = 8
    Top = 86
    Width = 34
    Height = 13
    Caption = 'version'
  end
  object Label3: TLabel
    Left = 8
    Top = 110
    Width = 44
    Height = 13
    Caption = 'Nb Poste'
  end
  object Label4: TLabel
    Left = 8
    Top = 214
    Width = 75
    Height = 13
    Caption = 'Clef d'#39'activation'
  end
  object LIDsERv: TLabel
    Left = 10
    Top = 14
    Width = 62
    Height = 13
    Caption = 'Id du serveur'
  end
  object resultat: TEdit
    Left = 89
    Top = 210
    Width = 89
    Height = 21
    ReadOnly = True
    TabOrder = 0
  end
  object Calcul: TButton
    Left = 165
    Top = 56
    Width = 75
    Height = 25
    Caption = 'Calcul'
    TabOrder = 1
    OnClick = CalculClick
  end
  object IdServeur: TMaskEdit
    Left = 88
    Top = 8
    Width = 41
    Height = 21
    EditMask = '00000;1;_'
    MaxLength = 5
    TabOrder = 2
    Text = '     '
  end
  object Cproduit: TMaskEdit
    Left = 88
    Top = 58
    Width = 49
    Height = 21
    EditMask = '00000;1;_'
    MaxLength = 5
    TabOrder = 3
    Text = '     '
  end
  object Cversion: TMaskEdit
    Left = 88
    Top = 82
    Width = 49
    Height = 21
    EditMask = '000;1;_'
    MaxLength = 3
    TabOrder = 4
    Text = '   '
  end
  object CNbr: TEdit
    Left = 88
    Top = 106
    Width = 49
    Height = 21
    MaxLength = 3
    TabOrder = 5
  end
  object GTempo: TGroupBox
    Left = 8
    Top = 137
    Width = 234
    Height = 64
    TabOrder = 6
    object ldtFIn: TLabel
      Left = 8
      Top = 26
      Width = 74
      Height = 13
      Caption = 'Valide Jusqu'#39'au'
      Visible = False
    end
    object CBTEMPO: TCheckBox
      Left = 10
      Top = -2
      Width = 97
      Height = 17
      Caption = 'Temporaire'
      TabOrder = 0
      OnClick = CBTEMPOClick
    end
    object DTFin: TDateTimePicker
      Left = 107
      Top = 22
      Width = 97
      Height = 21
      Date = 40169.401508275460000000
      Time = 40169.401508275460000000
      TabOrder = 1
      Visible = False
    end
  end
end

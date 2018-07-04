object FRecupTempoParam: TFRecupTempoParam
  Left = 271
  Top = 125
  Width = 474
  Height = 407
  Caption = 'Param'#233'trage de la r'#233'cup'#233'ration des fiches de bases'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label4: TLabel
    Left = 12
    Top = 4
    Width = 425
    Height = 13
    Caption = 
      'Par d'#233'faut, l'#39'ensemble de  vos codes statistiques de Tempo seron' +
      't repris. D'#233'cocher  ceux '
  end
  object Label5: TLabel
    Left = 10
    Top = 19
    Width = 168
    Height = 13
    Caption = 'que vous ne souhaitez pas int'#233'grer.'
  end
  object GroupBoxArticle: TGroupBox
    Left = 4
    Top = 272
    Width = 445
    Height = 77
    Caption = 'Prestations / Frais / Fournitures'
    TabOrder = 0
    object Label1: TLabel
      Left = 8
      Top = 14
      Width = 116
      Height = 13
      Caption = 'Nombre de statistiques  :'
    end
    object NbStatPrest: TLabel
      Left = 128
      Top = 14
      Width = 6
      Height = 13
      Caption = '0'
    end
    object PreCompt: TLabel
      Left = 12
      Top = 56
      Width = 77
      Height = 13
      Caption = 'Zone comptable'
    end
    object CheckStat1Prest: TCheckBox
      Left = 156
      Top = 14
      Width = 225
      Height = 17
      Caption = 'Stat 1 Prestations'
      Checked = True
      State = cbChecked
      TabOrder = 0
    end
    object CheckStat2Prest: TCheckBox
      Left = 156
      Top = 31
      Width = 224
      Height = 17
      Caption = 'Stat 2 Prestations'
      Checked = True
      State = cbChecked
      TabOrder = 1
    end
    object CheckCodePres: TCheckBox
      Left = 304
      Top = 52
      Width = 137
      Height = 17
      Caption = 'Ajout lettre devant code '
      TabOrder = 2
    end
    object ComboPreCompt: TComboBox
      Left = 156
      Top = 52
      Width = 145
      Height = 21
      ItemHeight = 13
      TabOrder = 3
      Text = 'Sans Affectation'
      Items.Strings = (
        'Sans affectation'
        'Famille prestation'
        'Nature comptable prestation')
    end
  end
  object GroupBoxEmploye: TGroupBox
    Left = 4
    Top = 163
    Width = 445
    Height = 102
    Caption = 'Employ'#233's'
    TabOrder = 1
    object Label2: TLabel
      Left = 8
      Top = 13
      Width = 116
      Height = 13
      Caption = 'Nombre de statistiques  :'
    end
    object NbStatEmp: TLabel
      Left = 128
      Top = 13
      Width = 6
      Height = 13
      Caption = '0'
    end
    object LabelLiaisonIsis: TLabel
      Left = 8
      Top = 85
      Width = 66
      Height = 13
      Caption = 'Liaison Isis II :'
    end
    object LabelIsis: TLabel
      Left = 81
      Top = 85
      Width = 20
      Height = 13
      Caption = 'Non'
    end
    object Shape1: TShape
      Left = 1
      Top = 84
      Width = 444
      Height = 1
      Pen.Color = clWhite
    end
    object CheckStat1Emp: TCheckBox
      Left = 4
      Top = 29
      Width = 97
      Height = 17
      Caption = 'Stat 1 Employ'#233
      Checked = True
      State = cbChecked
      TabOrder = 0
    end
    object CheckStat2Emp: TCheckBox
      Left = 4
      Top = 53
      Width = 93
      Height = 17
      Caption = 'Stat 2 Employ'#233
      Checked = True
      State = cbChecked
      TabOrder = 1
    end
    object CheckChangementIsis: TCheckBox
      Left = 144
      Top = 85
      Width = 287
      Height = 17
      Caption = 'Renum'#233'rotation des codes salari'#233's entre Isis II et PGI'
      TabOrder = 2
    end
    object ComboStat1Emp: TComboBox
      Left = 112
      Top = 29
      Width = 100
      Height = 21
      ItemHeight = 13
      TabOrder = 3
      Items.Strings = (
        'Table libre'
        'Etablissement'
        'D'#233'partement'
        'Type ressource')
    end
    object ComboStat2Emp: TComboBox
      Left = 112
      Top = 53
      Width = 100
      Height = 21
      ItemHeight = 13
      TabOrder = 4
      Items.Strings = (
        'Table libre'
        'Etablissement'
        'D'#233'partement'
        'Type ressource')
    end
    object CheckStat3Emp: TCheckBox
      Left = 228
      Top = 29
      Width = 97
      Height = 17
      Caption = 'Stat 3 Employ'#233
      Checked = True
      State = cbChecked
      TabOrder = 5
    end
    object ComboStat3Emp: TComboBox
      Left = 336
      Top = 29
      Width = 100
      Height = 21
      ItemHeight = 13
      TabOrder = 6
      Items.Strings = (
        'Table libre'
        'Etablissement'
        'D'#233'partement'
        'Type ressource')
    end
    object CheckStat4Emp: TCheckBox
      Left = 228
      Top = 53
      Width = 93
      Height = 17
      Caption = 'Stat 4 Employ'#233
      Checked = True
      State = cbChecked
      TabOrder = 7
    end
    object ComboStat4Emp: TComboBox
      Left = 336
      Top = 53
      Width = 100
      Height = 21
      ItemHeight = 13
      TabOrder = 8
      Items.Strings = (
        'Table libre'
        'Etablissement'
        'D'#233'partement'
        'Type ressource')
    end
  end
  object GroupBoxAffaire: TGroupBox
    Left = 4
    Top = 32
    Width = 445
    Height = 125
    Caption = 'Affaires'
    TabOrder = 2
    object Label6: TLabel
      Left = 8
      Top = 16
      Width = 116
      Height = 13
      Caption = 'Nombre de statistiques  :'
    end
    object NbStatAff: TLabel
      Left = 130
      Top = 16
      Width = 6
      Height = 13
      Caption = '0'
    end
    object LabelFamilleAff: TLabel
      Left = 8
      Top = 79
      Width = 72
      Height = 13
      Caption = 'Famille d'#39'affaire'
    end
    object AffCompt: THLabel
      Left = 8
      Top = 102
      Width = 77
      Height = 13
      Caption = 'Zone comptable'
    end
    object CheckStat1Aff: TCheckBox
      Left = 8
      Top = 32
      Width = 85
      Height = 17
      Caption = 'Stat 1 Affaire'
      Checked = True
      State = cbChecked
      TabOrder = 0
    end
    object CheckStat2Aff: TCheckBox
      Left = 8
      Top = 56
      Width = 93
      Height = 17
      Caption = 'Stat 2 Affaire'
      Checked = True
      State = cbChecked
      TabOrder = 1
    end
    object RadioFamAffStat: TRadioButton
      Left = 112
      Top = 79
      Width = 77
      Height = 17
      Caption = 'Statistique'
      TabOrder = 2
    end
    object RadioFamAffNon: TRadioButton
      Left = 191
      Top = 79
      Width = 78
      Height = 17
      Caption = 'non utilis'#233
      Checked = True
      TabOrder = 3
      TabStop = True
    end
    object ComboAffCompt: TComboBox
      Left = 112
      Top = 98
      Width = 145
      Height = 21
      ItemHeight = 13
      TabOrder = 4
      Text = 'Sans Affectation'
      Items.Strings = (
        'Sans Affectation'
        'Code famille'
        '')
    end
    object ComboStat1Aff: TComboBox
      Left = 112
      Top = 32
      Width = 100
      Height = 21
      ItemHeight = 13
      TabOrder = 5
      Items.Strings = (
        'Table libre'
        'Etablissement'
        'D'#233'partement'
        'Type reconduction')
    end
    object ComboStat2Aff: TComboBox
      Left = 112
      Top = 56
      Width = 100
      Height = 21
      ItemHeight = 13
      TabOrder = 6
      Items.Strings = (
        'Table libre'
        'Etablissement'
        'D'#233'partement'
        'Type reconduction')
    end
    object CHeckStat3Aff: TCheckBox
      Left = 234
      Top = 32
      Width = 85
      Height = 17
      Caption = 'Stat 3 Affaire'
      Checked = True
      State = cbChecked
      TabOrder = 7
    end
    object ComboStat3aff: TComboBox
      Left = 338
      Top = 32
      Width = 100
      Height = 21
      ItemHeight = 13
      TabOrder = 8
      Items.Strings = (
        'Table libre'
        'Etablissement'
        'D'#233'partement'
        'Type reconduction')
    end
    object CheckStat4Aff: TCheckBox
      Left = 234
      Top = 56
      Width = 93
      Height = 17
      Caption = 'Stat 4 Affaire'
      Checked = True
      State = cbChecked
      TabOrder = 9
    end
    object ComboStat4Aff: TComboBox
      Left = 338
      Top = 56
      Width = 100
      Height = 21
      ItemHeight = 13
      TabOrder = 10
      Items.Strings = (
        'Table libre'
        'Etablissement'
        'D'#233'partement'
        'Type reconduction')
    end
    object CheckApport: TCheckBox
      Left = 260
      Top = 100
      Width = 181
      Height = 17
      Caption = 'Apporteur en ressource libre 1'
      Checked = True
      State = cbChecked
      TabOrder = 11
    end
  end
  object FermerParamFicBase: TButton
    Left = 372
    Top = 354
    Width = 75
    Height = 25
    Caption = '&Fermer'
    TabOrder = 3
    OnClick = FermerParamFicBaseClick
  end
  object HMTrad: THSystemMenu
    MaxInsideSize = False
    ActiveResize = False
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 388
    Top = 11
  end
end

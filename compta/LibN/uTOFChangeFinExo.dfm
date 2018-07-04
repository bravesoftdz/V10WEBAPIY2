inherited FChangeFinExo: TFChangeFinExo
  Left = 375
  Top = 213
  Width = 394
  Height = 196
  Caption = 'Changement de date de fin d'#39'exercice'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  object TLDATEFIN: THLabel [0]
    Left = 22
    Top = 57
    Width = 52
    Height = 13
    Caption = 'Date de fin'
  end
  object TLLIBELLEEXO: THLabel [1]
    Left = 22
    Top = 93
    Width = 92
    Height = 13
    Caption = 'Libellé de l'#39'exercice'
  end
  object HLabel1: THLabel [2]
    Left = 22
    Top = 20
    Width = 41
    Height = 13
    Caption = 'Exercice'
  end
  inherited Dock971: TDock97
    Top = 134
    Width = 386
    inherited PBouton: TToolWindow97
      ClientWidth = 382
      ClientAreaWidth = 382
      inherited BValider: TToolbarButton97
        Left = 286
      end
      inherited BFerme: TToolbarButton97
        Left = 318
      end
      inherited HelpBtn: TToolbarButton97
        Left = 350
      end
      inherited BImprimer: TToolbarButton97
        Left = 254
      end
    end
  end
  object E_DATECOMPTABLE_: THCritMaskEdit [4]
    Left = 124
    Top = 53
    Width = 81
    Height = 21
    EditMask = '!99/99/0000;1;_'
    MaxLength = 10
    TabOrder = 1
    Text = '  /  /    '
    TagDispatch = 0
    Operateur = Egal
    ElipsisButton = True
    DisplayFormat = 'DD/MM/YYYY'
  end
  object E_LIBELLE: THCritMaskEdit [5]
    Left = 124
    Top = 89
    Width = 237
    Height = 21
    MaxLength = 35
    TabOrder = 2
    TagDispatch = 0
  end
  object E_Exercice: THValComboBox [6]
    Left = 124
    Top = 16
    Width = 237
    Height = 21
    ItemHeight = 13
    TabOrder = 3
    OnChange = E_ExerciceChange
    TagDispatch = 0
    DataType = 'TTEXERCICE'
  end
  inherited HMTrad: THSystemMenu
    Left = 228
    Top = 44
  end
end

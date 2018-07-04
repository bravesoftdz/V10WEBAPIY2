inherited FRecherche: TFRecherche
  Left = 294
  Top = 215
  Width = 330
  Height = 157
  Caption = 'Rechercher'
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel [0]
    Left = 16
    Top = 16
    Width = 62
    Height = 13
    Caption = 'Rechercher :'
  end
  inherited Dock971: TDock97
    Top = 88
    Width = 322
    inherited PBouton: TToolWindow97
      ClientWidth = 318
      ClientAreaWidth = 318
      inherited BValider: TToolbarButton97
        Left = 222
        Hint = 'Suivant'
      end
      inherited BFerme: TToolbarButton97
        Left = 254
      end
      inherited HelpBtn: TToolbarButton97
        Left = 286
        Visible = False
      end
      inherited BImprimer: TToolbarButton97
        Left = 190
      end
    end
  end
  object RechText: TEdit [2]
    Left = 80
    Top = 12
    Width = 225
    Height = 21
    TabOrder = 1
  end
  object RespecterMaj: TCheckBox [3]
    Left = 16
    Top = 56
    Width = 193
    Height = 17
    Alignment = taLeftJustify
    Caption = 'Respecter Majuscule/Minuscule'
    TabOrder = 2
  end
end

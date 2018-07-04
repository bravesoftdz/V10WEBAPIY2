inherited fRapport: TfRapport
  Left = 314
  Top = 189
  Width = 348
  Height = 259
  Caption = 'Rapport'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  object HLabel1: THLabel [0]
    Left = 28
    Top = 72
    Width = 84
    Height = 13
    Caption = 'Tps lecture fichier'
  end
  object HLabel2: THLabel [1]
    Left = 28
    Top = 115
    Width = 122
    Height = 13
    Caption = 'Tps traitement des enregs'
  end
  object HLabel3: THLabel [2]
    Left = 28
    Top = 157
    Width = 97
    Height = 13
    Caption = 'Tps ecriture en table'
  end
  object HLabel7: THLabel [3]
    Left = 28
    Top = 35
    Width = 41
    Height = 13
    Caption = 'Tps total'
  end
  inherited Dock971: TDock97
    Top = 190
    Width = 340
    inherited PBouton: TToolWindow97
      ClientWidth = 340
      ClientAreaWidth = 340
      inherited BValider: TToolbarButton97
        Left = 244
      end
      inherited BFerme: TToolbarButton97
        Left = 276
      end
      inherited HelpBtn: TToolbarButton97
        Left = 308
      end
      inherited BImprimer: TToolbarButton97
        Left = 212
      end
    end
  end
  object Tpsfic: THCritMaskEdit [5]
    Left = 177
    Top = 68
    Width = 116
    Height = 21
    TabOrder = 1
    TagDispatch = 0
  end
  object TpsTrait: THCritMaskEdit [6]
    Left = 177
    Top = 111
    Width = 116
    Height = 21
    TabOrder = 2
    TagDispatch = 0
  end
  object TpsEcr: THCritMaskEdit [7]
    Left = 177
    Top = 153
    Width = 116
    Height = 21
    TabOrder = 3
    TagDispatch = 0
  end
  object TpsTotal: THCritMaskEdit [8]
    Left = 177
    Top = 31
    Width = 116
    Height = 21
    TabOrder = 4
    TagDispatch = 0
  end
end

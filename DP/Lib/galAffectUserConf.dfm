inherited FAffectUserConf: TFAffectUserConf
  Left = 370
  Top = 172
  VertScrollBar.Range = 0
  BorderStyle = bsSingle
  Caption = 'Affectation des utilisateurs sélectionnés'
  ClientHeight = 193
  ClientWidth = 382
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  object lblNewGroupe: THLabel [0]
    Left = 40
    Top = 16
    Width = 234
    Height = 26
    Caption = 
      'Affecter les utilisateurs sélectionnés aux groupes de travail su' +
      'ivants :'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    WordWrap = True
  end
  inherited Dock971: TDock97
    Top = 158
    Width = 382
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
        Visible = False
      end
      inherited BImprimer: TToolbarButton97
        Left = 254
      end
    end
  end
  object mvNewGroupe: THMultiValComboBox [2]
    Left = 40
    Top = 68
    Width = 217
    Height = 21
    TabOrder = 1
    Abrege = False
    DataType = 'GROUPECONF'
    Complete = False
    OuInclusif = False
  end
end

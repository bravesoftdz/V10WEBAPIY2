inherited OptionsDlg: TOptionsDlg
  Left = 241
  Top = 279
  Width = 415
  Height = 259
  Caption = 'Options'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 190
    Width = 407
    inherited PBouton: TToolWindow97
      ClientWidth = 407
      ClientAreaWidth = 407
      inherited BValider: TToolbarButton97
        Left = 311
      end
      inherited BFerme: TToolbarButton97
        Left = 343
      end
      inherited HelpBtn: TToolbarButton97
        Left = 375
      end
      inherited BImprimer: TToolbarButton97
        Left = 279
      end
    end
  end
  object PageControl1: TPageControl [1]
    Left = 0
    Top = 0
    Width = 407
    Height = 190
    ActivePage = TabSheet1
    Align = alClient
    TabOrder = 1
    object TabSheet1: TTabSheet
      Caption = 'Paramètres'
      object LTable: TLabel
        Left = 6
        Top = 90
        Width = 84
        Height = 13
        Caption = 'Requête suivante'
      end
      object GroupBox2: TGroupBox
        Left = 4
        Top = 4
        Width = 345
        Height = 73
        Caption = ' &Format numérique '
        TabOrder = 0
        object Label4: TLabel
          Left = 8
          Top = 20
          Width = 100
          Height = 13
          Caption = 'Séparateur de milliers'
          FocusControl = cbSepMil
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label5: TLabel
          Left = 8
          Top = 44
          Width = 91
          Height = 13
          Caption = 'Séparateur décimal'
          FocusControl = cbSepDec
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object cbSepMil: TComboBox
          Left = 144
          Top = 16
          Width = 145
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 0
          OnChange = cbSepMilChange
          Items.Strings = (
            '. - point'
            ', - virgule'
            '  - espace'
            '  - (aucun)')
        end
        object cbSepDec: TComboBox
          Left = 144
          Top = 40
          Width = 145
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 1
          OnChange = cbSepDecChange
          Items.Strings = (
            '. - point'
            ', - virgule'
            '  - (aucun)')
        end
      end
      object Table: THCritMaskEdit
        Left = 102
        Top = 86
        Width = 115
        Height = 21
        MaxLength = 12
        TabOrder = 1
        TagDispatch = 0
        ElipsisButton = True
        OnElipsisClick = TableElipsisClick
      end
    end
  end
end

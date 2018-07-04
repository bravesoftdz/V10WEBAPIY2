inherited FDefvariable: TFDefvariable
  Left = 345
  Top = 192
  Width = 352
  Height = 276
  HelpContext = 1050
  Caption = 'Liste des variables'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  object Bevel1: TBevel [0]
    Left = 7
    Top = 12
    Width = 337
    Height = 161
    Shape = bsFrame
  end
  inherited Dock971: TDock97
    Top = 207
    Width = 344
    inherited PBouton: TToolWindow97
      ClientWidth = 344
      ClientAreaWidth = 344
      inherited BValider: TToolbarButton97
        Left = 215
      end
      inherited BFerme: TToolbarButton97
        Left = 247
      end
      inherited HelpBtn: TToolbarButton97
        Tag = 1050
        Left = 279
      end
      inherited Binsert: TToolbarButton97
        Visible = True
      end
      inherited BDelete: TToolbarButton97
        Visible = True
      end
      inherited BImprimer: TToolbarButton97
        Left = 183
      end
    end
  end
  object ListBox1: TListBox [2]
    Left = 0
    Top = 0
    Width = 344
    Height = 207
    Align = alClient
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -12
    Font.Name = 'BorlandTE'
    Font.Style = []
    ItemHeight = 15
    ParentFont = False
    Sorted = True
    TabOrder = 1
    OnDblClick = ListBox1DblClick
  end
  inherited HMTrad: THSystemMenu
    Left = 168
    Top = 68
  end
end

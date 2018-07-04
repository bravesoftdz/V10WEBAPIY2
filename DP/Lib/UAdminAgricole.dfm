inherited FAdminAgricole: TFAdminAgricole
  Left = 480
  Width = 544
  Height = 385
  Caption = 'Administration informations agricoles'
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 312
    Width = 536
    Height = 39
    inherited PBouton: TToolWindow97
      ClientHeight = 35
      ClientWidth = 536
      ClientAreaHeight = 35
      ClientAreaWidth = 536
      inherited BValider: TToolbarButton97
        Left = 440
      end
      inherited BFerme: TToolbarButton97
        Left = 472
      end
      inherited HelpBtn: TToolbarButton97
        Left = 504
      end
      inherited Binsert: TToolbarButton97
        Visible = True
      end
      inherited BDelete: TToolbarButton97
        Visible = True
      end
      inherited BImprimer: TToolbarButton97
        Left = 408
      end
    end
  end
  object TV: TTreeView [1]
    Left = 0
    Top = 0
    Width = 536
    Height = 312
    Align = alClient
    Images = ImageList
    Indent = 19
    StateImages = ImageList
    TabOrder = 1
  end
  object ImageList: THImageList
    GlobalIndexImages.Strings = (
      'Z0930_S16G1'
      'Z0174_S16G1')
    Left = 48
    Top = 56
  end
end

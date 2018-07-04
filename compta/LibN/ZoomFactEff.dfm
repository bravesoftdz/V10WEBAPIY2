inherited FZoomfactEff: TFZoomfactEff
  Left = 215
  Top = 193
  Caption = 'Liste des mouvements s'#233'lectionn'#233's'
  PixelsPerInch = 96
  TextHeight = 13
  inherited Pages: THPageControl2
    ActivePage = PCritere
    inherited PCritere: THTabSheet
      inherited Bevel1: TBevel
        Width = 592
      end
      object XX_WHERE: TEdit
        Left = 69
        Top = 9
        Width = 13
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
        Visible = False
      end
    end
  end
  inherited Dock971: TDock97
    inherited PFiltres: TToolWindow97
      Visible = False
    end
  end
  inherited Dock: TDock97
    inherited PanelBouton: TToolWindow97
      inherited BAgrandir: TToolbarButton97 [0]
        Visible = False
      end
      inherited BReduire: TToolbarButton97 [1]
      end
      inherited BOuvrir: TToolbarButton97
        Visible = False
      end
    end
  end
end

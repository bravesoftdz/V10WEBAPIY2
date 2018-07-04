inherited FDpTableauBord: TFDpTableauBord
  Left = 264
  Top = 166
  Width = 640
  Height = 480
  Caption = 'FDpTableauBord'
  OldCreateOrder = True
  OnCanResize = FormCanResize
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 407
    Width = 632
    Height = 39
    inherited PBouton: TToolWindow97
      ClientHeight = 35
      ClientWidth = 632
      ClientAreaHeight = 35
      ClientAreaWidth = 632
      inherited BValider: TToolbarButton97
        Left = 215
      end
      inherited BFerme: TToolbarButton97
        Left = 247
      end
      inherited HelpBtn: TToolbarButton97
        Left = 279
      end
      inherited BImprimer: TToolbarButton97
        Left = 183
        Visible = True
        OnClick = BImprimerClick
      end
    end
  end
  object Grille: THGrid [1]
    Left = 0
    Top = 0
    Width = 632
    Height = 407
    Align = alClient
    DefaultRowHeight = 18
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSelect]
    ScrollBars = ssNone
    TabOrder = 1
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = 13224395
  end
end

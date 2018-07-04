inherited FListeDistrib: TFListeDistrib
  Left = 271
  Top = 278
  Width = 552
  Height = 326
  Caption = 'Liste de distribution'
  OldCreateOrder = True
  ShowHint = True
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 260
    Width = 544
    Height = 39
    inherited PBouton: TToolWindow97
      ClientHeight = 35
      ClientWidth = 544
      ClientAreaHeight = 35
      ClientAreaWidth = 544
      ParentShowHint = False
      ShowHint = True
      inherited BValider: TToolbarButton97
        Left = 448
      end
      inherited BFerme: TToolbarButton97
        Left = 480
      end
      inherited HelpBtn: TToolbarButton97
        Left = 512
      end
      inherited Binsert: TToolbarButton97
        Hint = 'Nouvelle liste de distribution'
        ParentShowHint = False
        ShowHint = True
        Visible = True
      end
      inherited BImprimer: TToolbarButton97
        Left = 416
      end
    end
  end
  object Panel1: TPanel [1]
    Left = 0
    Top = 0
    Width = 544
    Height = 260
    Align = alClient
    TabOrder = 1
    object GrilleListeDistrib: THGrid
      Left = 1
      Top = 1
      Width = 542
      Height = 258
      Align = alClient
      ColCount = 3
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSelect]
      ScrollBars = ssVertical
      TabOrder = 0
      OnDblClick = ModifierListeDistrib
      SortedCol = -1
      Titres.Strings = (
        'Nom de la liste de distribution'
        'Description'
        'Priv'#233)
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
end

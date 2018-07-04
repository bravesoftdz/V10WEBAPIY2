inherited FDPTableauBordDetailSolde: TFDPTableauBordDetailSolde
  Left = 331
  Top = 201
  Width = 572
  Height = 402
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Solde Client'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 325
    Width = 564
    Height = 43
    inherited PBouton: TToolWindow97
      ClientHeight = 39
      ClientWidth = 564
      ClientAreaHeight = 39
      ClientAreaWidth = 564
      inherited BValider: TToolbarButton97
        Left = 532
      end
      inherited BFerme: TToolbarButton97
        Left = 500
        Visible = False
      end
      inherited HelpBtn: TToolbarButton97
        Left = 468
        Visible = False
      end
      inherited BImprimer: TToolbarButton97
        Left = 436
      end
    end
  end
  object Panel1: TPanel [1]
    Left = 0
    Top = 0
    Width = 564
    Height = 325
    Align = alClient
    Caption = 'Panel1'
    TabOrder = 1
    object GrilleDetailSolde: THGrid
      Left = 1
      Top = 1
      Width = 562
      Height = 323
      Align = alClient
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSelect]
      TabOrder = 0
      OnDblClick = GrilleDetailSoldeDblClick
      OnKeyDown = GrilleDetailSoldeKeyDown
      SortedCol = -1
      Titres.Strings = (
        'Dossier'
        'Libelle'
        'D'#233'bit'
        'Cr'#233'dit'
        'Solde')
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
  inherited HMTrad: THSystemMenu
    Left = 524
    Top = 8
  end
end

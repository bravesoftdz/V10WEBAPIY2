inherited FRepertoireTel: TFRepertoireTel
  Left = 335
  Top = 168
  Width = 533
  Height = 400
  Caption = 'R'#233'pertoire t'#233'l'#233'phonique'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 327
    Width = 525
    Height = 39
    inherited PBouton: TToolWindow97
      ClientHeight = 35
      ClientWidth = 525
      ClientAreaHeight = 35
      ClientAreaWidth = 525
      inherited BValider: TToolbarButton97
        Left = 429
      end
      inherited BFerme: TToolbarButton97
        Left = 461
      end
      inherited HelpBtn: TToolbarButton97
        Left = 493
      end
      inherited BImprimer: TToolbarButton97
        Left = 397
      end
    end
  end
  object GdRep: THGrid [1]
    Left = 0
    Top = 0
    Width = 525
    Height = 327
    Align = alClient
    ColCount = 8
    DefaultRowHeight = 17
    FixedCols = 0
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
    ParentShowHint = False
    PopupMenu = PMRepert
    ShowHint = True
    TabOrder = 1
    OnDblClick = GdRepDblClick
    SortedCol = -1
    Titres.Strings = (
      ''
      'Civ.'
      'Nom / D'#233'nomination'
      'Pr'#233'nom'
      'T'#233'l'#233'phone'
      'Autre t'#233'l.'
      'T'#233'l'#233'copie'
      'Email')
    Couleur = False
    MultiSelect = False
    TitleBold = True
    TitleCenter = True
    ColCombo = 0
    SortEnabled = True
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = 13224395
  end
  object PMRepert: TPopupMenu
    OnPopup = PMRepertPopup
    Left = 336
    Top = 168
  end
end

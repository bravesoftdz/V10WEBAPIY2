inherited FNewForme: TFNewForme
  Left = 223
  Top = 138
  Width = 707
  Height = 508
  Caption = 'Fonctions associ'#233'es '#224' la forme juridique'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 431
    Width = 699
    Height = 43
    inherited PBouton: TToolWindow97
      ClientHeight = 39
      ClientWidth = 699
      ClientAreaHeight = 39
      ClientAreaWidth = 699
      inherited BValider: TToolbarButton97
        Left = 603
      end
      inherited BFerme: TToolbarButton97
        Left = 635
      end
      inherited HelpBtn: TToolbarButton97
        Left = 667
      end
      inherited BImprimer: TToolbarButton97
        Left = 571
      end
    end
  end
  object PForme_f: TPanel [1]
    Left = 0
    Top = 0
    Width = 699
    Height = 48
    Align = alTop
    Locked = True
    TabOrder = 1
    object LNewForme_f: TLabel
      Left = 12
      Top = 12
      Width = 113
      Height = 13
      Caption = 'Nouvelle forme juridique'
    end
    object eForme_f: TEdit
      Left = 466
      Top = 8
      Width = 84
      Height = 21
      Enabled = False
      TabOrder = 0
    end
    object eFormeLib_f: TEdit
      Left = 152
      Top = 8
      Width = 303
      Height = 21
      Enabled = False
      TabOrder = 1
    end
  end
  object Pfonction_f: TPanel [2]
    Left = 0
    Top = 48
    Width = 699
    Height = 383
    Align = alClient
    TabOrder = 0
    object GFonction_f: THGrid
      Left = 1
      Top = 1
      Width = 316
      Height = 381
      Align = alLeft
      ColCount = 1
      DefaultColWidth = 12
      DefaultRowHeight = 16
      FixedCols = 0
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
      TabOrder = 0
      SortedCol = -1
      Couleur = False
      MultiSelect = True
      TitleBold = True
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = 13224395
    end
    object PBT_f: TPanel
      Left = 317
      Top = 1
      Width = 76
      Height = 381
      Align = alLeft
      TabOrder = 1
      DesignSize = (
        76
        381)
      object BFoncToType_f: TToolbarButton97
        Left = 24
        Top = 177
        Width = 28
        Height = 27
        Hint = 'Ajoute'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Default = True
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        NumGlyphs = 2
        ParentFont = False
        Spacing = -1
        OnClick = BFoncToType_fClick
        GlobalIndexImage = 'Z0056_S16G2'
        IsControl = True
      end
      object BTypeToFonc_f: TToolbarButton97
        Left = 24
        Top = 147
        Width = 29
        Height = 27
        Hint = 'Enl'#232've'
        AllowAllUp = True
        Anchors = [akTop, akRight]
        Default = True
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        NumGlyphs = 2
        ParentFont = False
        Spacing = -1
        OnClick = BTypeToFonc_fClick
        GlobalIndexImage = 'Z0077_S16G2'
        IsControl = True
      end
    end
    object GTypeFonct_f: THGrid
      Left = 393
      Top = 1
      Width = 305
      Height = 381
      Align = alClient
      ColCount = 1
      DefaultColWidth = 12
      DefaultRowHeight = 16
      FixedCols = 0
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
      TabOrder = 2
      SortedCol = -1
      Couleur = False
      MultiSelect = True
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

inherited FRapportVerifAmort: TFRapportVerifAmort
  Left = 342
  Top = 138
  Width = 619
  Height = 321
  Caption = 'Rapport de vérification des immobilisations'
  OldCreateOrder = True
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 248
    Width = 611
    Height = 39
    inherited PBouton: TToolWindow97
      ClientHeight = 35
      ClientWidth = 611
      ClientAreaHeight = 35
      ClientAreaWidth = 611
      inherited BValider: TToolbarButton97
        Left = 515
        Visible = False
      end
      inherited BFerme: TToolbarButton97
        Left = 547
      end
      inherited HelpBtn: TToolbarButton97
        Left = 579
      end
      inherited BImprimer: TToolbarButton97
        Left = 483
        Visible = True
        OnClick = BImprimerClick
      end
      object FBCorrection: TToolbarButton97
        Left = 4
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Correction des immobilisations sélectionnées'
        AllowAllUp = True
        Flat = False
        Glyph.Data = {
          66010000424D6601000000000000760000002800000014000000140000000100
          040000000000F000000000000000000000001000000010000000000000000000
          8000008000000080800080000000800080008080000080808000C0C0C0000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00DDDDDDDDDDDD
          DDDDDDDD0000DDDDDDDDDDDD7070DDDD0000DDDDDDDDDDD0E8E87DDD0000DDDD
          DDDDDD7E8E8E80DD0000DDDDDDDDDD0FE008E70D0000DDDDDD000D7EF00E800D
          0000DDDD008EF007EFE8E77D0000DDD0E8E8EF607EFE080D0000DD0E8E8E8EF6
          070080DD0000DD08E8E8E8EF00770DDD0000D0FE8E707E8EF07DDDDD0000D0EF
          E80008E8E00DDDDD0000D08EFE707E8E800DDDDD0000DD08EFE8E8E8080DDDDD
          0000DD0E8EFE8E8E00DDDDDD0000DDD0E8EFE8E080DDDDDD0000DDDD008EF008
          0DDDDDDD0000DDDDD7000800DDDDDDDD0000DDDDDDD7000DDDDDDDDD0000DDDD
          DDDDDDDDDDDDDDDD0000}
        ParentShowHint = False
        ShowHint = True
        OnClick = FBCorrectionClick
      end
      object CNombre: THLabel
        Left = 208
        Top = 8
        Width = 52
        Height = 13
        Caption = 'CNombre'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object BToutSel: TToolbarButton97
        Left = 36
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Tout sélectionner'
        AllowAllUp = True
        Flat = False
        Glyph.Data = {
          76010000424D7601000000000000760000002800000020000000100000000100
          04000000000000010000130B0000130B00001000000000000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
          333333333333333333FF33333333333330003333333333333888333333333333
          300033FFFFFF3333388839999993333333333888888F3333333F399999933333
          3300388888833333338833333333333333003333333333333388333333333333
          3333333333333333333F3333333E3333330033333F33333333883333333EC333
          330033338F3333333388EEEEEEEECCE3333333F88FFFFFFF3FF3ECCCCCCCCCCC
          399338888888888F88F33CCCCCCCCCCC399338888888888388333333333ECC33
          333333388F33333333FF33333333C33330003333833333333888333333333333
          3000333333333333388833333333333333333333333333333333}
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BToutSelClick
      end
      object BRefresh: TToolbarButton97
        Left = 68
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Relancer la vérification'
        AllowAllUp = True
        Flat = False
        Glyph.Data = {
          E6000000424DE60000000000000076000000280000000F0000000E0000000100
          0400000000007000000000000000000000001000000000000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          777077777777777777707700877777777770770B3087777777707770BB008777
          777077770BB300877770777770BBB00877707777700BBB007770777700BBB007
          77707777700BBB007770777777700BB0077077777777700B0070777777777770
          00707777777777777770}
        ParentShowHint = False
        ShowHint = True
        OnClick = BRefreshClick
      end
      object BReinit: TToolbarButton97
        Left = 100
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Réinitialisation'
        AllowAllUp = True
        Flat = False
        Glyph.Data = {
          F6000000424DF600000000000000760000002800000010000000100000000100
          0400000000008000000000000000000000001000000000000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00FFFFFFFFFFFF
          FFFFFFFFCCFF00FFFFFFFFFFCCF0780FFFFFFFFFFFF7B780FFFFFCCFCC7B7B78
          0FFFFCCFC4BCC7B780FFFFFF7B74CB7B780FFFFCC7B7B7B7BF0FF87C4BCC7B7B
          F8FFF8F7B74CB7BF8FFFFF8F7B7B7BF80FFFFFF8F7B7BF8B0FFFFFFF8F7BF8BB
          8FFFFFFFF8FF8FF8FFFFFFFFFF88888FFFFFFFFFFFFFFFFFFFFF}
        ParentShowHint = False
        ShowHint = True
        OnClick = BReinitClick
      end
    end
  end
  object HPanel1: THPanel [1]
    Left = 0
    Top = 0
    Width = 611
    Height = 248
    Align = alClient
    FullRepaint = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object FListe: THGrid
      Left = 1
      Top = 1
      Width = 609
      Height = 246
      Align = alClient
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
      TabOrder = 0
      OnFlipSelection = FListeFlipSelection
      SortedCol = -1
      Titres.Strings = (
        ''
        'Code'
        'Libellé'
        'Date d'#39'achat'
        'Corrigé')
      Couleur = False
      MultiSelect = True
      TitleBold = True
      TitleCenter = True
      OnRowEnter = FListeRowEnter
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = 13224395
      ColWidths = (
        64
        103
        260
        79
        64)
      RowHeights = (
        18
        18)
    end
  end
  inherited HMTrad: THSystemMenu
    Left = 144
    Top = 236
  end
end

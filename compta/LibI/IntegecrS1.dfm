object Integration: TIntegration
  Left = 228
  Top = 194
  Width = 614
  Height = 302
  BorderIcons = [biSystemMenu]
  Caption = 'Intégration des écritures'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Dock971: TDock97
    Left = 0
    Top = 240
    Width = 606
    Height = 35
    AllowDrag = False
    Position = dpBottom
    object ToolWindow971: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 602
      Caption = 'Barre outils intégration'
      ClientAreaHeight = 31
      ClientAreaWidth = 602
      DockPos = 0
      FullSize = True
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnResize = ToolWindow971Resize
      object BValider: TToolbarButton97
        Left = 508
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Valider'
        Anchors = [akTop, akRight]
        Default = True
        DisplayMode = dmGlyphOnly
        Caption = 'Valider'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Glyph.Data = {
          BE060000424DBE06000000000000360400002800000024000000120000000100
          0800000000008802000000000000000000000001000000010000000000000000
          80000080000000808000800000008000800080800000C0C0C000C0DCC000F0C8
          A400000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          000000000000000000000000000000000000F0FBFF00A4A0A000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00030303030303
          0303030303030303030303030303030303030303030303030303030303030303
          03030303030303030303030303030303030303030303FF030303030303030303
          03030303030303040403030303030303030303030303030303F8F8FF03030303
          03030303030303030303040202040303030303030303030303030303F80303F8
          FF030303030303030303030303040202020204030303030303030303030303F8
          03030303F8FF0303030303030303030304020202020202040303030303030303
          0303F8030303030303F8FF030303030303030304020202FA0202020204030303
          0303030303F8FF0303F8FF030303F8FF03030303030303020202FA03FA020202
          040303030303030303F8FF03F803F8FF0303F8FF03030303030303FA02FA0303
          03FA0202020403030303030303F8FFF8030303F8FF0303F8FF03030303030303
          FA0303030303FA0202020403030303030303F80303030303F8FF0303F8FF0303
          0303030303030303030303FA0202020403030303030303030303030303F8FF03
          03F8FF03030303030303030303030303FA020202040303030303030303030303
          0303F8FF0303F8FF03030303030303030303030303FA02020204030303030303
          03030303030303F8FF0303F8FF03030303030303030303030303FA0202020403
          030303030303030303030303F8FF0303F8FF03030303030303030303030303FA
          0202040303030303030303030303030303F8FF03F8FF03030303030303030303
          03030303FA0202030303030303030303030303030303F8FFF803030303030303
          030303030303030303FA0303030303030303030303030303030303F803030303
          0303030303030303030303030303030303030303030303030303030303030303
          0303}
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 540
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Anchors = [akTop, akRight]
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
        Glyph.Data = {
          CE070000424DCE07000000000000360000002800000024000000120000000100
          1800000000009807000000000000000000000000000000000000008284008284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284008284008284008284008284008284008284008284
          0082840082840082840082840082840082848482848482840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284FFFFFF008284008284008284008284008284008284
          0082840082840082840082840082840082840082840082840082840000FF0000
          840000848482840082840082840082840082840082840000FF84828400828400
          8284008284008284008284008284008284008284848284848284FFFFFF008284
          008284008284008284008284008284FFFFFF0082840082840082840082840082
          840082840082840000FF00008400008400008484828400828400828400828400
          00FF000084000084848284008284008284008284008284008284008284848284
          FFFFFF008284848284FFFFFF008284008284008284FFFFFF848284848284FFFF
          FF0082840082840082840082840082840082840000FF00008400008400008400
          00848482840082840000FF000084000084000084000084848284008284008284
          008284008284008284848284FFFFFF008284008284848284FFFFFF008284FFFF
          FF848284008284008284848284FFFFFF00828400828400828400828400828400
          82840000FF000084000084000084000084848284000084000084000084000084
          000084848284008284008284008284008284008284848284FFFFFF0082840082
          84008284848284FFFFFF848284008284008284008284008284848284FFFFFF00
          82840082840082840082840082840082840000FF000084000084000084000084
          0000840000840000840000848482840082840082840082840082840082840082
          84008284848284FFFFFF00828400828400828484828400828400828400828400
          8284FFFFFF848284008284008284008284008284008284008284008284008284
          0000FF0000840000840000840000840000840000848482840082840082840082
          84008284008284008284008284008284008284848284FFFFFF00828400828400
          8284008284008284008284FFFFFF848284008284008284008284008284008284
          0082840082840082840082840082840000840000840000840000840000848482
          8400828400828400828400828400828400828400828400828400828400828400
          8284848284FFFFFF008284008284008284008284008284848284008284008284
          0082840082840082840082840082840082840082840082840082840000FF0000
          8400008400008400008484828400828400828400828400828400828400828400
          8284008284008284008284008284008284848284FFFFFF008284008284008284
          8482840082840082840082840082840082840082840082840082840082840082
          840082840000FF00008400008400008400008400008484828400828400828400
          8284008284008284008284008284008284008284008284008284008284848284
          008284008284008284008284848284FFFFFF0082840082840082840082840082
          840082840082840082840082840000FF00008400008400008484828400008400
          0084000084848284008284008284008284008284008284008284008284008284
          008284008284848284008284008284008284008284008284848284FFFFFF0082
          840082840082840082840082840082840082840082840000FF00008400008400
          00848482840082840000FF000084000084000084848284008284008284008284
          008284008284008284008284008284848284008284008284008284848284FFFF
          FF008284008284848284FFFFFF00828400828400828400828400828400828400
          82840000FF0000840000848482840082840082840082840000FF000084000084
          000084848284008284008284008284008284008284008284848284FFFFFF0082
          84008284848284008284848284FFFFFF008284008284848284FFFFFF00828400
          82840082840082840082840082840082840000FF000084008284008284008284
          0082840082840000FF0000840000840000840082840082840082840082840082
          84008284848284FFFFFFFFFFFF848284008284008284008284848284FFFFFF00
          8284008284848284FFFFFF008284008284008284008284008284008284008284
          0082840082840082840082840082840082840082840000FF0000840000FF0082
          8400828400828400828400828400828400828484828484828400828400828400
          8284008284008284848284FFFFFFFFFFFFFFFFFF848284008284008284008284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284008284008284008284008284848284848284848284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828400828400828400828400
          8284008284008284008284008284008284008284008284008284008284008284
          008284008284008284008284008284008284}
        ModalResult = 2
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
      end
      object HelpBtn: TToolbarButton97
        Left = 572
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        Anchors = [akTop, akRight]
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Glyph.Data = {
          BE060000424DBE06000000000000360400002800000024000000120000000100
          0800000000008802000000000000000000000001000000010000000000000000
          80000080000000808000800000008000800080800000C0C0C000C0DCC000F0CA
          A600000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          000000000000000000000000000000000000F0FBFF00A4A0A000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00030303030303
          03030606030303030303030303030303030303FFFF0303030303030303030303
          0303030303060404060303030303030303030303030303F8F8FF030303030303
          030303030303030303FE06060403030303030303030303030303F8FF03F8FF03
          0303030303030303030303030303FE060603030303030303030303030303F8FF
          FFF8FF0303030303030303030303030303030303030303030303030303030303
          030303F8F8030303030303030303030303030303030304040603030303030303
          0303030303030303FFFF03030303030303030303030303030306060604030303
          0303030303030303030303F8F8F8FF0303030303030303030303030303FE0606
          0403030303030303030303030303F8FF03F8FF03030303030303030303030303
          03FE06060604030303030303030303030303F8FF03F8FF030303030303030303
          030303030303FE060606040303030303030303030303F8FF0303F8FF03030303
          0303030303030303030303FE060606040303030303030303030303F8FF0303F8
          FF030303030303030303030404030303FE060606040303030303030303FF0303
          F8FF0303F8FF030303030303030306060604030303FE06060403030303030303
          F8F8FF0303F8FF0303F8FF03030303030303FE06060604040406060604030303
          030303F8FF03F8FFFFFFF80303F8FF0303030303030303FE0606060606060606
          06030303030303F8FF0303F8F8F8030303F8FF030303030303030303FEFE0606
          060606060303030303030303F8FFFF030303030303F803030303030303030303
          0303FEFEFEFEFE03030303030303030303F8F8FFFFFFFFFFF803030303030303
          0303030303030303030303030303030303030303030303F8F8F8F8F803030303
          0303}
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        IsControl = True
      end
      object BImprimer: TToolbarButton97
        Left = 476
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Caption = 'Imprimer'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Glyph.Data = {
          4E010000424D4E01000000000000760000002800000012000000120000000100
          040000000000D800000000000000000000001000000010000000000000000000
          BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00DDDDDDDDDDDD
          DDDDDD000000DDD00000000000DDDD000000DD0777777777070DDD000000D000
          000000000070DD000000D0777777FFF77000DD000000D077777799977070DD00
          0000D0000000000000770D000000D0777777777707070D000000DD0000000000
          70700D000000DDD0FFFFFFFF07070D000000DDDD0FCCCCCF0000DD000000DDDD
          0FFFFFFFF0DDDD000000DDDDD0FCCCCCF0DDDD000000DDDDD0FFFFFFFF0DDD00
          0000DDDDDD000000000DDD000000DDDDDDDDDDDDDDDDDD000000DDDDDDDDDDDD
          DDDDDD000000DDDDDDDDDDDDDDDDDD000000}
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
      end
      object bZoom: TToolbarButton97
        Left = 2
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Visualiser les écritures'
        Anchors = [akTop, akRight]
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Glyph.Data = {
          AA040000424DAA04000000000000360000002800000013000000130000000100
          1800000000007404000000000000000000000000000000000000FF00FFFF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF
          00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
          FFFFFFFFFFFFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFFFFFFBF0000BF
          0000FFFFFFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFFFFFFBF000000FFFFFFFF00BF00
          00FFFFFFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
          00FFFF00FFFF00FFFF00FFFFFFFFBF000000FFFFFFFF80FF8000BF0000FF00FF
          FF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFBF000000FFFFFFFF00BF0000BF0000FF00FFFF00FFFF00FF00
          0000FF00FFFF00FFFF00FFFF00FFFFFFFFFFFFFF000000000000000000000000
          BF000000FFFFFFFF00FF8000BF0000FF00FFFF00FFFF00FFFF00FF000000FF00
          FFFF00FFFFFFFFFFFFFF000000000000FF00FFFF00FFFF00FFFF00FFFF00FFBF
          0000BF0000FF0040FF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FF
          FFFFFF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF0000
          00FF0040FF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFFFFFFFFFFFF00
          0000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFFFFFF000000FF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF
          00FFFF00FFFF00FFFF00FF000000FF00FFFFFFFF000000FF00FFFF00FF00FFFF
          FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00
          FFFF00FFFF00FF000000FF00FFFFFFFF000000FF00FFFF00FF00FFFFFF00FFFF
          00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FF
          FF00FF000000FF00FFFF00FF000000FF00FFFF00FF00FFFF00FFFFFF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FF00
          0000FF00FFFF00FFFF00FF000000FF00FFFF00FF00FFFF00FFFF00FFFFFF00FF
          FF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00
          FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
          00FF000000FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FF
          FF00FFFF00FF000000000000FF00FFFF00FFFF00FFFF00FF000000000000FF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF
          00FFFF00FFFF00FF000000000000000000000000FF00FFFF00FFFF00FFFF00FF
          FF00FFFF00FFFF00FFFF00FFFF00FF000000FF00FFFF00FFFF00FFFF00FFFF00
          FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
          00FFFF00FFFF00FFFF00FF000000}
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = bZoomClick
      end
    end
  end
  object PanelEch: TPanel
    Left = 261
    Top = 0
    Width = 345
    Height = 240
    Align = alLeft
    TabOrder = 1
    object HLabel6: THLabel
      Left = 9
      Top = 19
      Width = 133
      Height = 13
      Caption = 'Génération des écritures du '
      Contour = False
      Ombre = False
      OmbreDecalX = 0
      OmbreDecalY = 0
      OmbreColor = clBlack
      Tag2 = 0
      Tag3 = 0
      TagExport = 0
    end
    object HLabel7: THLabel
      Left = 232
      Top = 19
      Width = 12
      Height = 13
      Caption = 'au'
      Contour = False
      Ombre = False
      OmbreDecalX = 0
      OmbreDecalY = 0
      OmbreColor = clBlack
      Tag2 = 0
      Tag3 = 0
      TagExport = 0
    end
    object LibelleCBanque: THLabel
      Left = 496
      Top = 8
      Width = 74
      Height = 13
      Caption = 'LibelleCBanque'
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      Visible = False
      Contour = False
      Ombre = False
      OmbreDecalX = 0
      OmbreDecalY = 0
      OmbreColor = clBlack
      Tag2 = 0
      Tag3 = 0
      TagExport = 0
    end
    object HLabel1: THLabel
      Left = 9
      Top = 44
      Width = 90
      Height = 13
      Caption = 'Génération de type'
      Contour = False
      Ombre = False
      OmbreDecalX = 0
      OmbreDecalY = 0
      OmbreColor = clBlack
      Tag2 = 0
      Tag3 = 0
      TagExport = 0
    end
    object DateDebutEch: THCritMaskEdit
      Left = 145
      Top = 15
      Width = 81
      Height = 21
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      TabOrder = 0
      Text = '  /  /    '
      OnKeyPress = DateKeyPress
      Operateur = Superieur
      OpeType = otDate
      ElipsisButton = True
      ControlerDate = False
    end
    object DateFinEch: THCritMaskEdit
      Left = 252
      Top = 15
      Width = 81
      Height = 21
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      TabOrder = 1
      Text = '  /  /    '
      OnKeyPress = DateKeyPress
      Operateur = Inferieur
      OpeType = otDate
      ElipsisButton = True
      ControlerDate = False
    end
    object GEcheance: TGroupBox
      Left = 69
      Top = 65
      Width = 224
      Height = 72
      Caption = 'GEcheance'
      TabOrder = 2
      object tJEcheance: THLabel
        Left = 8
        Top = 23
        Width = 93
        Height = 13
        Caption = 'Journal d'#39'échéance'
        Enabled = False
        FocusControl = JEcheance
        Contour = False
        Ombre = False
        OmbreDecalX = 0
        OmbreDecalY = 0
        OmbreColor = clBlack
        Tag2 = 0
        Tag3 = 0
        TagExport = 0
      end
      object tCTva: THLabel
        Left = 8
        Top = 48
        Width = 75
        Height = 13
        Caption = 'Compte de &TVA'
        Enabled = False
        FocusControl = CTva
        Contour = False
        Ombre = False
        OmbreDecalX = 0
        OmbreDecalY = 0
        OmbreColor = clBlack
        Tag2 = 0
        Tag3 = 0
        TagExport = 0
      end
      object cbEcrEcheance: TCheckBox
        Left = 6
        Top = 2
        Width = 129
        Height = 13
        Caption = 'Ecritures d'#39'éc&héances'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        OnClick = cbEcrEcheanceClick
      end
      object JEcheance: THValComboBox
        Left = 105
        Top = 19
        Width = 114
        Height = 21
        Style = csDropDownList
        Enabled = False
        ItemHeight = 13
        Sorted = False
        TabOrder = 1
        Exhaustif = exOui
        Plus = 'AND J_NATUREJAL="ACH"'
        Abrege = False
        Vide = False
        HideInGrid = False
        DataType = 'TTJOURNAL'
        DataTypeParametrable = False
        ComboWidth = 0
        DisableTab = False
      end
      object CTva: THCritMaskEdit
        Left = 105
        Top = 44
        Width = 114
        Height = 21
        Enabled = False
        TabOrder = 2
        Plus = 'AND G_NATUREGENE = "TDE"'
        DataType = 'TTGENERAUX'
        OpeType = otString
        ElipsisButton = True
        ControlerDate = False
      end
    end
    object GBanque: TGroupBox
      Left = 69
      Top = 145
      Width = 224
      Height = 76
      Caption = 'GroupBox4'
      TabOrder = 3
      object tCBanque: THLabel
        Left = 16
        Top = 49
        Width = 90
        Height = 13
        Caption = 'Compte de &banque'
        Enabled = False
        FocusControl = CBanque
        Contour = False
        Ombre = False
        OmbreDecalX = 0
        OmbreDecalY = 0
        OmbreColor = clBlack
        Tag2 = 0
        Tag3 = 0
        TagExport = 0
      end
      object tJBanque: THLabel
        Left = 16
        Top = 23
        Width = 88
        Height = 13
        Caption = 'Journal de banque'
        Enabled = False
        FocusControl = JBanque
        Contour = False
        Ombre = False
        OmbreDecalX = 0
        OmbreDecalY = 0
        OmbreColor = clBlack
        Tag2 = 0
        Tag3 = 0
        TagExport = 0
      end
      object cbEcrPaiement: TCheckBox
        Left = 8
        Top = 0
        Width = 121
        Height = 17
        Caption = 'Ecritures de &paiement'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        OnClick = cbEcrPaiementClick
      end
      object JBanque: THValComboBox
        Left = 112
        Top = 19
        Width = 107
        Height = 21
        Style = csDropDownList
        Enabled = False
        ItemHeight = 13
        Sorted = False
        TabOrder = 1
        OnChange = JBanqueChange
        Exhaustif = exOui
        Plus = 'AND J_NATUREJAL="TRE"'
        Abrege = False
        Vide = False
        HideInGrid = False
        DataType = 'TTJOURNAL'
        DataTypeParametrable = False
        ComboWidth = 0
        DisableTab = False
      end
      object CBanque: THCritMaskEdit
        Left = 112
        Top = 45
        Width = 107
        Height = 21
        Enabled = False
        TabOrder = 2
        Plus = 'and G_NATUREGENE="BQE"'
        DataType = 'TTGENERAUX'
        OpeType = otString
        ControlerDate = False
      end
    end
    object cTypEcrEch: THValComboBox
      Left = 145
      Top = 41
      Width = 145
      Height = 21
      ItemHeight = 13
      Sorted = False
      TabOrder = 4
      Items.Strings = (
        'Synthétique'
        'Détaillé')
      Exhaustif = exOui
      Abrege = False
      Vide = False
      HideInGrid = False
      DataTypeParametrable = False
      Values.Strings = (
        '0'
        '2')
      ComboWidth = 0
      DisableTab = False
    end
  end
  object PanelDot: TPanel
    Left = 0
    Top = 0
    Width = 261
    Height = 240
    Align = alLeft
    TabOrder = 2
    object HLabel3: THLabel
      Left = 9
      Top = 12
      Width = 90
      Height = 13
      Caption = 'Génération de type'
      Contour = False
      Ombre = False
      OmbreDecalX = 0
      OmbreDecalY = 0
      OmbreColor = clBlack
      Tag2 = 0
      Tag3 = 0
      TagExport = 0
    end
    object GroupBox2: TGroupBox
      Left = 21
      Top = 41
      Width = 220
      Height = 80
      Caption = 'Dates'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      object HLabelDateCalcul: THLabel
        Left = 6
        Top = 22
        Width = 110
        Height = 13
        Caption = 'Calcul des &dotations au'
        FocusControl = DateCalcul
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Contour = False
        Ombre = False
        OmbreDecalX = 0
        OmbreDecalY = 0
        OmbreColor = clBlack
        Tag2 = 0
        Tag3 = 0
        TagExport = 0
      end
      object HLabel2: THLabel
        Left = 6
        Top = 55
        Width = 103
        Height = 13
        Caption = 'Ecritures &générées au'
        FocusControl = DateGeneration
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Contour = False
        Ombre = False
        OmbreDecalX = 0
        OmbreDecalY = 0
        OmbreColor = clBlack
        Tag2 = 0
        Tag3 = 0
        TagExport = 0
      end
      object DateCalcul: THCritMaskEdit
        Left = 134
        Top = 18
        Width = 79
        Height = 21
        EditMask = '!99/99/0000;1;_'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 10
        ParentFont = False
        TabOrder = 0
        Text = '  /  /    '
        OnKeyPress = DateKeyPress
        OpeType = otDate
        ElipsisButton = True
        ControlerDate = False
      end
      object DateGeneration: THCritMaskEdit
        Left = 134
        Top = 51
        Width = 79
        Height = 21
        EditMask = '!99/99/0000;1;_'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 10
        ParentFont = False
        TabOrder = 1
        Text = '  /  /    '
        OnKeyPress = DateKeyPress
        OpeType = otDate
        ElipsisButton = True
        ControlerDate = False
      end
    end
    object GroupBox1: TGroupBox
      Left = 21
      Top = 133
      Width = 220
      Height = 80
      Caption = 'Journal/Libellé'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      object HLabel4: THLabel
        Left = 10
        Top = 22
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = CodeJournal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Contour = False
        Ombre = False
        OmbreDecalX = 0
        OmbreDecalY = 0
        OmbreColor = clBlack
        Tag2 = 0
        Tag3 = 0
        TagExport = 0
      end
      object HLabel5: THLabel
        Left = 10
        Top = 55
        Width = 65
        Height = 13
        Caption = '&Libellé unique'
        FocusControl = LibelleUnique
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Contour = False
        Ombre = False
        OmbreDecalX = 0
        OmbreDecalY = 0
        OmbreColor = clBlack
        Tag2 = 0
        Tag3 = 0
        TagExport = 0
      end
      object LibelleUnique: TEdit
        Left = 78
        Top = 51
        Width = 127
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
      end
      object CodeJournal: THValComboBox
        Left = 78
        Top = 18
        Width = 127
        Height = 21
        Style = csDropDownList
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        Sorted = False
        TabOrder = 0
        Exhaustif = exOui
        Plus = 'AND J_NATUREJAL="OD"'
        Abrege = False
        Vide = False
        HideInGrid = False
        DataType = 'TTJOURNAL'
        DataTypeParametrable = False
        ComboWidth = 0
        DisableTab = False
      end
    end
    object cTypEcrDot: THValComboBox
      Left = 106
      Top = 9
      Width = 145
      Height = 21
      ItemHeight = 13
      Sorted = False
      TabOrder = 2
      Items.Strings = (
        'Synthétique'
        'Groupé par compte'
        'Détaillé')
      Exhaustif = exOui
      Abrege = False
      Vide = False
      HideInGrid = False
      DataTypeParametrable = False
      Values.Strings = (
        '0'
        '1'
        '2')
      ComboWidth = 0
      DisableTab = False
    end
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Total progressif'
      'Total général'
      'Calcul des échéances au'
      'Intégration des écritures de dotation'
      'Intégration des écritures d'#39'échéances'
      
        '5;?caption?;Cette date n'#39'appartient pas à l'#39'exercice en cours;E;' +
        'O;O;O;'
      '6;?caption?;Problème lors de la mise à jour de la base;E;O;O;O;'
      
        '7;?caption?;Aucune écriture de dotation pour l'#39'exercice en cours' +
        ';E;O;O;O;'
      
        '8;?Caption?;Êtes-vous sûr de vouloir intégrer les écritures dans' +
        ' la comptabilité ?;Q;YN;Y;N;'
      '9;?caption?;L'#39'intégration des écritures a échoué;E;O;O;O;'
      
        '10;?caption?;Le journal n'#39'est pas renseigné correctement.;E;O;O;' +
        'O;'
      '11;?caption?;Aucune écriture n'#39'a été générée.;A;O;O;O;'
      
        '12;?caption?;Les dates doivent être des dates de l'#39'exercice en c' +
        'ours.;A;O;O;O;'
      
        '13;?caption?;La liste d'#39'écritures affichée ne correspond pas aux' +
        ' critères renseignés.;E;O;O;O;'
      '14;?caption?;Date incorrecte.;E;O;O;O;'
      '15;?caption?;Journal de banque incorrect.;E;O;O;O;'
      '16;?caption?;Journal d'#39'échéance incorrect.;E;O;O;O;'
      '17;?caption?;Compte de TVA incorrect.;E;O;O;O;'
      '18;?caption?;Compte de banque incorrect.;E;O;O;O;'
      '19;?caption?;Date de calcul incorrecte.;E;O;O;O;'
      '20;?caption?;Date de génération incorrecte.;E;O;O;O;'
      '21;?caption?;Date de début d'#39'échéance incorrecte.;E;O;O;O;'
      '22;?caption?;Date de fin d'#39'échéance incorrecte.;E;O;O;O;'
      '23;?caption?;L'#39'intégration términé correctement;A;O;O;O;'
      '24;?caption?;Paramètres incorrects;E;O;O;O;'
      '')
    Left = 306
    Top = 220
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 269
    Top = 172
  end
  object POPF: TPopupMenu
    Left = 303
    Top = 172
    object BCreerFiltre: TMenuItem
      Caption = '&Créer un filtre'
    end
    object BSaveFiltre: TMenuItem
      Caption = '&Enregistrer le filtre'
    end
    object BDelFiltre: TMenuItem
      Caption = '&Supprimer le filtre'
    end
    object BRenFiltre: TMenuItem
      Caption = '&Renommer le filtre'
    end
    object BNouvRech: TMenuItem
      Caption = '&Nouvelle recherche'
    end
  end
end

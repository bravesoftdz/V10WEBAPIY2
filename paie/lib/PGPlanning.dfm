object Planning: TPlanning
  Left = 240
  Top = 210
  Width = 604
  Height = 449
  HelpContext = 42255005
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Planning des absences'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PPanel: TPanel
    Left = 0
    Top = 0
    Width = 596
    Height = 415
    HelpContext = 42255005
    Align = alClient
    Caption = 'PPanel'
    TabOrder = 0
    object Dock971: TDock97
      Left = 1
      Top = 363
      Width = 594
      Height = 51
      AllowDrag = False
      Position = dpBottom
      object PBouton: TToolWindow97
        Left = 0
        Top = 0
        ClientHeight = 47
        ClientWidth = 594
        Caption = 'Barre outils fiche'
        ClientAreaHeight = 47
        ClientAreaWidth = 594
        DockPos = 0
        FullSize = True
        TabOrder = 0
        DesignSize = (
          594
          47)
        object BImprimer: TToolbarButton97
          Left = 494
          Top = 2
          Width = 28
          Height = 27
          Hint = 'Imprimer'
          Anchors = [akTop, akRight]
          Flat = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          ParentShowHint = False
          ShowHint = True
          OnClick = BImprimerClick
          GlobalIndexImage = 'Z0369_S16G1'
        end
        object BFerme: TToolbarButton97
          Left = 528
          Top = 2
          Width = 28
          Height = 27
          Hint = 'Fermer'
          Anchors = [akTop, akRight]
          Cancel = True
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
          GlyphMask.Data = {00000000}
          NumGlyphs = 2
          ParentShowHint = False
          ShowHint = True
          OnClick = BFermeClick
        end
        object HelpBtn: TToolbarButton97
          Left = 561
          Top = 2
          Width = 28
          Height = 27
          Hint = 'Aide'
          HelpContext = 42255005
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
          GlyphMask.Data = {00000000}
          NumGlyphs = 2
          ParentFont = False
          ParentShowHint = False
          ShowHint = True
          Spacing = -1
          OnClick = HelpBtnClick
          IsControl = True
        end
        object TRECAP: TToolbarButton97
          Left = 426
          Top = 2
          Width = 28
          Height = 27
          Hint = 'R'#233'capitulatif'
          Anchors = [akTop, akRight]
          Flat = False
          ParentShowHint = False
          ShowHint = True
          OnClick = TRECAPClick
          GlobalIndexImage = 'Z0285_S16G1'
        end
        object Bperso: TToolbarButton97
          Left = 460
          Top = 2
          Width = 28
          Height = 27
          Hint = 'Personnaliser'
          Anchors = [akTop, akRight]
          Flat = False
          ParentShowHint = False
          ShowHint = True
          OnClick = BpersoClick
          GlobalIndexImage = 'Z0025_S16G1'
        end
        object TExcel: TToolbarButton97
          Left = 392
          Top = 2
          Width = 28
          Height = 27
          Hint = 'Excel'
          Anchors = [akTop, akRight]
          Flat = False
          ParentShowHint = False
          ShowHint = True
          OnClick = TExcelClick
          GlobalIndexImage = 'Z2263_S16G1'
        end
        object BExporter: TToolbarButton97
          Left = 358
          Top = 2
          Width = 28
          Height = 27
          Hint = 'Exporter'
          Anchors = [akTop, akRight]
          Flat = False
          ParentShowHint = False
          ShowHint = True
          OnClick = BExporterClick
          GlobalIndexImage = 'Z0724_S16G1'
        end
        object BFirst: TToolbarButton97
          Left = 206
          Top = 2
          Width = 28
          Height = 27
          Hint = 'D'#233'but d'#39'exercice'
          Anchors = [akTop, akRight]
          Flat = False
          ParentShowHint = False
          ShowHint = True
          OnClick = BFirstClick
          GlobalIndexImage = 'Z0301_S16G1'
        end
        object BPrev: TToolbarButton97
          Left = 234
          Top = 2
          Width = 28
          Height = 27
          Hint = 'Mois pr'#233'c'#233'dent'
          Anchors = [akTop, akRight]
          Flat = False
          ParentShowHint = False
          ShowHint = True
          OnClick = BPrevClick
          GlobalIndexImage = 'Z0017_S16G1'
        end
        object BNext: TToolbarButton97
          Left = 262
          Top = 2
          Width = 28
          Height = 27
          Hint = 'Mois suivant'
          Anchors = [akTop, akRight]
          Flat = False
          ParentShowHint = False
          ShowHint = True
          OnClick = BNextClick
          GlobalIndexImage = 'Z0031_S16G1'
        end
        object BLast: TToolbarButton97
          Left = 290
          Top = 2
          Width = 28
          Height = 27
          Hint = 'Fin d'#39'exercice'
          Anchors = [akTop, akRight]
          Flat = False
          ParentShowHint = False
          ShowHint = True
          OnClick = BLastClick
          GlobalIndexImage = 'Z0264_S16G1'
        end
        object BtnDetail: TToolbarButton97
          Left = 324
          Top = 2
          Width = 28
          Height = 27
          Hint = 'Voir r'#233'capitulatif'
          AllowAllUp = True
          Anchors = [akTop, akRight]
          GroupIndex = 1
          Flat = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Layout = blGlyphTop
          ParentFont = False
          ParentShowHint = False
          ShowHint = True
          OnClick = BtnDetailClick
          GlobalIndexImage = 'Z0270_S16G1'
        end
      end
    end
    object PRECAP: TPanel
      Left = 1
      Top = 284
      Width = 594
      Height = 79
      Align = alBottom
      TabOrder = 1
      Visible = False
      DesignSize = (
        594
        79)
      object GBAbsence: TGroupBox
        Left = 2
        Top = 4
        Width = 212
        Height = 73
        Anchors = [akLeft, akTop, akBottom]
        Caption = 'Absence en cours'
        TabOrder = 0
        object Valide: THLabel
          Left = 9
          Top = 31
          Width = 29
          Height = 13
          Caption = 'Valid'#233
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Attente: THLabel
          Left = 9
          Top = 44
          Width = 49
          Height = 13
          Caption = 'En attente'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object EnCoursCP: THLabel
          Left = 69
          Top = 16
          Width = 63
          Height = 13
          Caption = 'En  Cours CP'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object NbValidePri: THLabel
          Left = 97
          Top = 31
          Width = 6
          Height = 13
          Alignment = taRightJustify
          Caption = '0'
        end
        object NbAttentePri: THLabel
          Left = 97
          Top = 44
          Width = 6
          Height = 13
          Alignment = taRightJustify
          Caption = '0'
        end
        object NbValideRTT: THLabel
          Left = 165
          Top = 31
          Width = 6
          Height = 13
          Alignment = taRightJustify
          Caption = '0'
        end
        object NbAttenteRTT: THLabel
          Left = 165
          Top = 44
          Width = 6
          Height = 13
          Alignment = taRightJustify
          Caption = '0'
        end
        object EnCoursRtt: THLabel
          Left = 137
          Top = 16
          Width = 71
          Height = 13
          Caption = 'En  Cours RTT'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
      end
      object GBRecap: TGroupBox
        Left = 213
        Top = 4
        Width = 381
        Height = 73
        Anchors = [akLeft, akTop, akRight, akBottom]
        Caption = 'R'#233'capitulatif absence'
        TabOrder = 1
        object LBN1: THLabel
          Left = 85
          Top = 16
          Width = 89
          Height = 13
          Caption = 'CP Ann'#233'e '#233'coul'#233'e'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object AcquisN1: THLabel
          Left = 113
          Top = 29
          Width = 46
          Height = 13
          Alignment = taRightJustify
          Caption = 'AcquisN1'
        end
        object PrisN1: THLabel
          Left = 128
          Top = 43
          Width = 31
          Height = 13
          Alignment = taRightJustify
          Caption = 'PrisN1'
        end
        object RestantN1: THLabel
          Left = 108
          Top = 56
          Width = 51
          Height = 13
          Alignment = taRightJustify
          Caption = 'RestantN1'
        end
        object LBN: THLabel
          Left = 195
          Top = 16
          Width = 92
          Height = 13
          Caption = 'CP Ann'#233'e en cours'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object AcquisN: THLabel
          Left = 231
          Top = 29
          Width = 40
          Height = 13
          Alignment = taRightJustify
          Caption = 'AcquisN'
        end
        object PrisN: THLabel
          Left = 246
          Top = 43
          Width = 25
          Height = 13
          Alignment = taRightJustify
          Caption = 'PrisN'
        end
        object RestantN: THLabel
          Left = 226
          Top = 56
          Width = 45
          Height = 13
          Alignment = taRightJustify
          Caption = 'RestantN'
        end
        object LBRTT: THLabel
          Left = 332
          Top = 16
          Width = 22
          Height = 13
          Alignment = taCenter
          Caption = 'RTT'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object AcquisRtt: THLabel
          Left = 317
          Top = 29
          Width = 54
          Height = 13
          Alignment = taRightJustify
          Caption = 'AcquisRTT'
        end
        object PrisRTT: THLabel
          Left = 332
          Top = 43
          Width = 39
          Height = 13
          Alignment = taRightJustify
          Caption = 'PrisRTT'
        end
        object RestantRTT: THLabel
          Left = 312
          Top = 56
          Width = 59
          Height = 13
          Alignment = taRightJustify
          Caption = 'RestantRTT'
        end
        object LBAcquis: THLabel
          Left = 4
          Top = 29
          Width = 32
          Height = 13
          Caption = 'Acquis'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LBPris: THLabel
          Left = 4
          Top = 43
          Width = 17
          Height = 13
          Caption = 'Pris'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LBRestant: THLabel
          Left = 4
          Top = 56
          Width = 37
          Height = 13
          Caption = 'Restant'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clTeal
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
      end
    end
  end
end

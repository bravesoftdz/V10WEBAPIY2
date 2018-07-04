object fPlanSbv: TfPlanSbv
  Left = 282
  Top = 123
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Plan d'#39#233'chelonnement de la subvention'
  ClientHeight = 476
  ClientWidth = 395
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object HPanel3: THPanel
    Left = 0
    Top = 249
    Width = 395
    Height = 91
    Align = alClient
    BevelInner = bvRaised
    BevelOuter = bvNone
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object GrilleAmort: THGrid
      Left = 1
      Top = 1
      Width = 393
      Height = 89
      Hint = 'Amortissement de la subvention d'#39#233'quipement'
      Align = alClient
      ColCount = 2
      DefaultColWidth = 110
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
      ParentShowHint = False
      ScrollBars = ssVertical
      ShowHint = True
      TabOrder = 0
      OnClick = GrilleAmortClick
      OnMouseUp = GrilleAmortMouseUp
      SortedCol = -1
      Titres.Strings = (
        'Exercices'
        'Amortissement ')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      OnRowEnter = GrilleAmortRowEnter
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = 13224395
      ColWidths = (
        109
        110)
      RowHeights = (
        18
        18)
    end
  end
  object HPanel2: THPanel
    Left = 0
    Top = 0
    Width = 395
    Height = 249
    Align = alTop
    BevelOuter = bvNone
    FullRepaint = False
    TabOrder = 1
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object Label1: TLabel
      Left = 6
      Top = 11
      Width = 93
      Height = 13
      Caption = 'Code Immobilisation'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object LibelleImmo: THLabel
      Left = 72
      Top = 33
      Width = 257
      Height = 13
      AutoSize = False
      Caption = 'LibelleImmo'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label6: TLabel
      Left = 6
      Top = 55
      Width = 61
      Height = 13
      Caption = 'Date d'#39'achat'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object CodeImmo: THLabel
      Left = 111
      Top = 11
      Width = 130
      Height = 13
      AutoSize = False
      Caption = 'CodeImmo'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object DateAchat: THLabel
      Left = 111
      Top = 55
      Width = 81
      Height = 13
      AutoSize = False
      Caption = 'DateAchat'
    end
    object Label8: TLabel
      Left = 223
      Top = 55
      Width = 48
      Height = 13
      Caption = 'Valeur HT'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object ValeurHT: THLabel
      Left = 288
      Top = 55
      Width = 81
      Height = 13
      AutoSize = False
      Caption = 'ValeurHT'
    end
    object HLabel11: THLabel
      Left = 7
      Top = 33
      Width = 56
      Height = 13
      Caption = 'D'#233'signation'
    end
    object gAmortEco: TGroupBox
      Left = 6
      Top = 68
      Width = 387
      Height = 145
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      object LIB_DUREE: THLabel
        Left = 16
        Top = 100
        Width = 97
        Height = 13
        Caption = 'Dur'#233'e d'#39'inali'#233'nabilit'#233
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object lDuree: TLabel
        Left = 128
        Top = 100
        Width = 29
        Height = 13
        Alignment = taRightJustify
        Caption = 'Dur'#233'e'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object lMethode: TLabel
        Left = 168
        Top = 32
        Width = 66
        Height = 13
        Caption = 'M'#233'thode SBV'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object Label3: TLabel
        Left = 16
        Top = 32
        Width = 120
        Height = 13
        Caption = 'M'#233'thode d'#39'amortissement'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object LabelTauxEco: THLabel
        Left = 16
        Top = 122
        Width = 27
        Height = 13
        Caption = 'Taux '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object lTaux: TLabel
        Left = 80
        Top = 122
        Width = 48
        Height = 13
        Caption = 'Taux SBV'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel5: THLabel
        Left = 16
        Top = 55
        Width = 134
        Height = 13
        Caption = 'Date d'#233'but d'#39'amortissement '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object lDateDeb: THLabel
        Left = 168
        Top = 55
        Width = 71
        Height = 13
        Caption = 'deb amort SBV'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object Hmois: THLabel
        Left = 165
        Top = 100
        Width = 21
        Height = 13
        Caption = 'mois'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object MNTSBV: THLabel
        Left = 168
        Top = 10
        Width = 58
        Height = 13
        Caption = 'montant sbv'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object li_montant: THLabel
        Left = 16
        Top = 10
        Width = 120
        Height = 13
        Caption = 'Montant de la subvention'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object lib_octroi: THLabel
        Left = 16
        Top = 78
        Width = 141
        Height = 13
        Caption = 'Date d'#39'octroi de la subvention'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object lDateOctroi: THLabel
        Left = 168
        Top = 78
        Width = 71
        Height = 13
        Caption = 'dateoctroi SBV'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
    end
    object gb_reprise: THGroupBox
      Left = 6
      Top = 213
      Width = 385
      Height = 33
      TabOrder = 1
      object HLabel8: THLabel
        Left = 16
        Top = 13
        Width = 86
        Height = 13
        Caption = 'Cumuls ant'#233'rieurs '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clGreen
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object reprise: THLabel
        Left = 168
        Top = 13
        Width = 72
        Height = 13
        Caption = 'montant reprise'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clGreen
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
    end
  end
  object HPanel1: THPanel
    Left = 0
    Top = 439
    Width = 395
    Height = 37
    Align = alBottom
    FullRepaint = False
    TabOrder = 2
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    DesignSize = (
      395
      37)
    object ToolbarButton973: TToolbarButton97
      Left = 363
      Top = 6
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
      GlyphMask.Data = {00000000}
      Margin = 2
      NumGlyphs = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      OnClick = ToolbarButton973Click
      IsControl = True
    end
    object BImprimer: TToolbarButton97
      Left = 299
      Top = 6
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
      GlyphMask.Data = {00000000}
      Margin = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      OnClick = BImprimerClick
    end
    object bFerme: TToolbarButton97
      Left = 331
      Top = 6
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Anchors = [akTop, akRight]
      Cancel = True
      Default = True
      Flat = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Glyph.Data = {
        4E010000424D4E01000000000000760000002800000012000000120000000100
        040000000000D800000000000000000000001000000000000000000000000000
        80000080000000808000800000008000800080800000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        3333330000003333883333333333330000003339118333339833330000003339
        1118333911833300000033391111839111183300000033339111181111183300
        0000333339111111118333000000333333911111183333000000333333311111
        8333330000003333333911118333330000003333339111118333330000003333
        3911181118333300000033339111839111833300000033339118333911183300
        0000333339133333911133000000333333333333391933000000333333333333
        333333000000333333333333333333000000}
      GlyphMask.Data = {00000000}
      Margin = 2
      ModalResult = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      IsControl = True
    end
    object bIntermediaire: TToolbarButton97
      Left = 5
      Top = 5
      Width = 28
      Height = 27
      Hint = 'Dotations interm'#233'diaires'
      Flat = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Glyph.Data = {
        42010000424D4201000000000000760000002800000011000000110000000100
        040000000000CC00000000000000000000001000000010000000000000000000
        80000080000000808000800000008000800080800000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
        77777000000077777777777777777000000070000000000007777000000070FF
        FFFFFFFF07777000000070FCCFCCCCCF07777000000070FFFFFFFFFF07777000
        000070FCCFCCCCCF07777000000070FFFFFFFFFF07777000000070FFFFFFF0FF
        07777000000070F00FFF0B0F07770000000070F0F0F0B0F000700000000070FF
        0B0B0F0FBF0000000000700000F0F0FBFBF0000000007777770B0FBFBFB00000
        000077777770FBFBFB0000000000777777770000007000000000777777777777
        777770000000}
      GlyphMask.Data = {00000000}
      Margin = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      OnClick = bIntermediaireClick
    end
    object bAnteriorite: TCheckBox
      Left = 41
      Top = 11
      Width = 177
      Height = 17
      Hint = 'Affichage du plan depuis le premier exercice d'#39'amortissement'
      Caption = '&Affichage de l'#39'ant'#233'riorit'#233' du plan'
      Checked = True
      State = cbChecked
      TabOrder = 0
      OnClick = bAnterioriteClick
    end
  end
  object HPanel4: THPanel
    Left = 0
    Top = 340
    Width = 395
    Height = 99
    Align = alBottom
    FullRepaint = False
    TabOrder = 3
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object Label7: TLabel
      Left = 9
      Top = 36
      Width = 118
      Height = 13
      Caption = 'Cumul de l'#39'amortissement'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label9: TLabel
      Left = 9
      Top = 65
      Width = 103
      Height = 13
      Caption = 'Amortissement restant'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clGreen
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label10: TLabel
      Left = 9
      Top = 10
      Width = 51
      Height = 13
      Caption = 'Dont sortie'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object DontCession: TLabel
      Left = 196
      Top = 10
      Width = 21
      Height = 13
      Alignment = taRightJustify
      Caption = '0.00'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object VNC: TLabel
      Left = 196
      Top = 65
      Width = 21
      Height = 13
      Alignment = taRightJustify
      Caption = '0.00'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clGreen
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object Cumul: TLabel
      Left = 196
      Top = 36
      Width = 21
      Height = 13
      Alignment = taRightJustify
      Caption = '0.00'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Cumul reprise'
      'Cumul dotations'
      'Valeur nette'
      'au %s'
      'Dont cessions'
      
        '5;?Caption?;Pas de m'#233'thodes d'#39'amortissement renseign'#233'es.;W;O;O;O' +
        ';'
      ' Mois  '
      ' %'
      'R'#233'int'#233'gration fiscale'
      'Quote part personnelle'
      'Non applicable'
      'M'#233'thode')
    Left = 250
    Top = 491
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 216
    Top = 493
  end
end

object PlanAmortissement: TPlanAmortissement
  Left = 284
  Top = 66
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Plan d'#39'amortissement'
  ClientHeight = 595
  ClientWidth = 597
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
    Top = 281
    Width = 597
    Height = 151
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
      Width = 595
      Height = 149
      Hint = 'Tableau des dotations'
      Align = alClient
      ColCount = 6
      DefaultColWidth = 110
      DefaultRowHeight = 18
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
      ParentShowHint = False
      ScrollBars = ssVertical
      ShowHint = True
      TabOrder = 0
      OnClick = GrilleAmortClick
      OnDblClick = GrilleAmortDblClick
      OnMouseUp = GrilleAmortMouseUp
      SortedCol = -1
      Titres.Strings = (
        'Exercices'
        'Dotations Eco.'
        'Dotations Fiscales'
        'D'#233'rog./Reprise'
        'R'#233'int'#233'gr./D'#233'duct.')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      OnRowEnter = GrilleAmortRowEnter
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
      ColWidths = (
        109
        110
        111
        110
        110
        110)
      RowHeights = (
        18
        18)
    end
  end
  object HPanel2: THPanel
    Left = 0
    Top = 0
    Width = 597
    Height = 281
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
      Left = 312
      Top = 11
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
      Top = 31
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
      Top = 31
      Width = 81
      Height = 13
      AutoSize = False
      Caption = 'DateAchat'
    end
    object Label8: TLabel
      Left = 247
      Top = 31
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
      Left = 312
      Top = 31
      Width = 81
      Height = 13
      AutoSize = False
      Caption = 'ValeurHT'
    end
    object HLabel11: THLabel
      Left = 247
      Top = 11
      Width = 56
      Height = 13
      Caption = 'D'#233'signation'
    end
    object gAmortFiscal: TGroupBox
      Left = 304
      Top = 49
      Width = 271
      Height = 118
      Caption = 'Amortissement fiscal'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      object HLabel2: THLabel
        Left = 22
        Top = 56
        Width = 29
        Height = 13
        Caption = 'Dur'#233'e'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object lDureeFisc: TLabel
        Left = 107
        Top = 56
        Width = 51
        Height = 13
        AutoSize = False
        Caption = 'Dur'#233'e fisc.'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object lMethodeFisc: TLabel
        Left = 107
        Top = 17
        Width = 64
        Height = 13
        Caption = 'M'#233'thode fisc.'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object lMethodeFiscale: TLabel
        Left = 22
        Top = 17
        Width = 45
        Height = 13
        Caption = 'M'#233'thode '
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object HLabel3: THLabel
        Left = 22
        Top = 76
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
      object lTauxFisc: TLabel
        Left = 107
        Top = 76
        Width = 51
        Height = 13
        AutoSize = False
        Caption = 'Taux fisc.'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object lBaseFisc: TLabel
        Left = 107
        Top = 96
        Width = 46
        Height = 13
        Caption = 'Base fisc.'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel6: THLabel
        Left = 22
        Top = 95
        Width = 24
        Height = 13
        Caption = 'Base'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel12: THLabel
        Left = 22
        Top = 37
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
      object lDateDebFis: THLabel
        Left = 176
        Top = 37
        Width = 60
        Height = 13
        Caption = 'deb amort fis'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
    end
    object gReintegrationFiscale: TGroupBox
      Left = 304
      Top = 172
      Width = 271
      Height = 101
      Caption = 'Donn'#233'es fiscales'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      object Label4: TLabel
        Left = 22
        Top = 19
        Width = 50
        Height = 13
        Caption = 'Quote part'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object Label5: TLabel
        Left = 22
        Top = 39
        Width = 128
        Height = 13
        Caption = 'Plafond d'#233'ductibilit'#233' fiscale'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object MontantReintegration: TLabel
        Left = 159
        Top = 39
        Width = 102
        Height = 13
        Caption = 'MontantReintegration'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object MontantQuotePart: TLabel
        Left = 159
        Top = 19
        Width = 90
        Height = 13
        AutoSize = False
        Caption = 'Quote part'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object DPIaffecte: TLabel
        Left = 22
        Top = 59
        Width = 60
        Height = 13
        Caption = 'DPI affect'#233'e'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object MontantDPI: TLabel
        Left = 159
        Top = 59
        Width = 60
        Height = 13
        Caption = 'Montant DPI'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
    end
    object GroupBox1: TGroupBox
      Left = 8
      Top = 172
      Width = 281
      Height = 101
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 3
      object RepriseEco: TLabel
        Left = 247
        Top = 20
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
      object RepriseFiscal: TLabel
        Left = 247
        Top = 39
        Width = 21
        Height = 13
        Alignment = taRightJustify
        Caption = '0.00'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object RepriseDerog: TLabel
        Left = 247
        Top = 58
        Width = 21
        Height = 13
        Alignment = taRightJustify
        Caption = '0.00'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clOlive
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object HLabel1: THLabel
        Left = 7
        Top = 0
        Width = 83
        Height = 13
        Caption = 'Cumuls ant'#233'rieurs'
      end
      object HLabel8: THLabel
        Left = 21
        Top = 20
        Width = 63
        Height = 13
        Caption = #233'conomiques'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clGreen
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel9: THLabel
        Left = 21
        Top = 39
        Width = 33
        Height = 13
        Caption = 'fiscaux'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel10: THLabel
        Left = 21
        Top = 58
        Width = 58
        Height = 13
        Caption = 'd'#233'rogatoires'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clOlive
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object Lib_Reint: THLabel
        Left = 21
        Top = 77
        Width = 110
        Height = 13
        Caption = 'r'#233'int'#233'gration/d'#233'duction'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object RepriseReint: TLabel
        Left = 247
        Top = 77
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
    end
    object gAmortEco: TGroupBox
      Left = 6
      Top = 49
      Width = 283
      Height = 118
      Caption = 'Amortissement '#233'conomique'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      object HLabel4: THLabel
        Left = 18
        Top = 56
        Width = 29
        Height = 13
        Caption = 'Dur'#233'e'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object lDureeEco: TLabel
        Left = 80
        Top = 56
        Width = 53
        Height = 13
        Caption = 'Dur'#233'e '#233'co.'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object lMethodeEco: TLabel
        Left = 80
        Top = 17
        Width = 66
        Height = 13
        Caption = 'M'#233'thode '#233'co.'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object Label3: TLabel
        Left = 18
        Top = 17
        Width = 45
        Height = 13
        Caption = 'M'#233'thode '
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
        Left = 18
        Top = 76
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
      object lTauxEco: TLabel
        Left = 80
        Top = 76
        Width = 48
        Height = 13
        Caption = 'Taux '#233'co.'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object lBaseEco: TLabel
        Left = 80
        Top = 96
        Width = 45
        Height = 13
        Caption = 'Base '#233'co'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel7: THLabel
        Left = 18
        Top = 96
        Width = 24
        Height = 13
        Caption = 'Base'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel5: THLabel
        Left = 18
        Top = 37
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
      object lDateDebEco: THLabel
        Left = 168
        Top = 37
        Width = 68
        Height = 13
        Caption = 'deb amort eco'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
    end
  end
  object HPanel1: THPanel
    Left = 0
    Top = 563
    Width = 597
    Height = 32
    Align = alBottom
    FullRepaint = False
    TabOrder = 2
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    DesignSize = (
      597
      32)
    object ToolbarButton973: TToolbarButton97
      Left = 565
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
      Left = 501
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
      Margin = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      OnClick = BImprimerClick
      GlobalIndexImage = 'Z0369_S16G1'
    end
    object bFerme: TToolbarButton97
      Left = 533
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
      Margin = 2
      ModalResult = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = -1
      GlobalIndexImage = 'Z0021_S16G1'
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
      Margin = 2
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      OnClick = bIntermediaireClick
      GlobalIndexImage = 'Z0003_S16G1'
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
    Top = 432
    Width = 597
    Height = 131
    Align = alBottom
    FullRepaint = False
    TabOrder = 3
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object Label7: TLabel
      Left = 9
      Top = 52
      Width = 29
      Height = 13
      Caption = 'Cumul'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label9: TLabel
      Left = 9
      Top = 73
      Width = 77
      Height = 13
      Caption = 'Valeur r'#233'siduelle'
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
    object Label11: TLabel
      Left = 9
      Top = 31
      Width = 102
      Height = 13
      Caption = 'Dont except./d'#233'pr'#233'c.'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object DontCessionEco: TLabel
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
    object DontCessionFiscal: TLabel
      Left = 309
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
    object DontExcepEco: TLabel
      Left = 196
      Top = 31
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
    object DontExcepFiscal: TLabel
      Left = 309
      Top = 31
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
    object VNCEco: TLabel
      Left = 196
      Top = 73
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
    object VNCFiscal: TLabel
      Left = 309
      Top = 73
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
    object CumulEco: TLabel
      Left = 196
      Top = 52
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
    object CumulFiscal: TLabel
      Left = 309
      Top = 52
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
    object CumulDerog: TLabel
      Left = 420
      Top = 52
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
    object CumulReint: TLabel
      Left = 530
      Top = 52
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
    object DontCessionDerog: TLabel
      Left = 420
      Top = 12
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
    object DontCessionReint: TLabel
      Left = 530
      Top = 12
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
    object DontVR: TLabel
      Left = 9
      Top = 111
      Width = 230
      Height = 13
      Caption = '*** Dont r'#233'int'#233'gration valeur r'#233'siduelle comptable'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object VR: TLabel
      Left = 530
      Top = 111
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
    object DontDeduc: TLabel
      Left = 9
      Top = 93
      Width = 148
      Height = 13
      Caption = '*** Dont d'#233'duction Montant HT'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object Deduc: TLabel
      Left = 530
      Top = 93
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
    Left = 394
    Top = 491
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 336
    Top = 493
  end
end

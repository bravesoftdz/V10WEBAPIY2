object FResulexo: TFResulexo
  Left = 297
  Top = 124
  Width = 568
  Height = 389
  HelpContext = 7683000
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'R'#233'sultat d'#39'exercice'
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
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object FListe1: THGrid
    Left = 0
    Top = 58
    Width = 560
    Height = 257
    TabStop = False
    Align = alClient
    ColCount = 4
    Ctl3D = True
    DefaultColWidth = 100
    DefaultRowHeight = 18
    RowCount = 2
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
    ParentCtl3D = False
    ParentFont = False
    ParentShowHint = False
    ShowHint = False
    TabOrder = 5
    Visible = False
    SortedCol = -1
    Titres.Strings = (
      'P'#233'riode;G'
      'Charges;D'
      'Produits;D'
      'R'#233'sultats;D')
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = 13224395
    ColWidths = (
      105
      140
      140
      140)
    RowHeights = (
      18
      16)
  end
  object Fliste: THGrid
    Left = 0
    Top = 58
    Width = 560
    Height = 257
    TabStop = False
    Align = alClient
    ColCount = 4
    Ctl3D = True
    DefaultColWidth = 100
    DefaultRowHeight = 18
    RowCount = 2
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
    ParentCtl3D = False
    ParentFont = False
    ParentShowHint = False
    ShowHint = False
    TabOrder = 0
    SortedCol = -1
    Titres.Strings = (
      'P'#233'riode;G'
      'Charges;D'
      'Produits;D'
      'R'#233'sultats;D')
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = 13224395
    ColWidths = (
      105
      140
      140
      140)
    RowHeights = (
      18
      16)
  end
  object Dock: TDock97
    Left = 0
    Top = 315
    Width = 560
    Height = 47
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 43
      ClientWidth = 560
      Caption = 'Actions'
      ClientAreaHeight = 43
      ClientAreaWidth = 560
      DockPos = 0
      DragHandleStyle = dhNone
      FullSize = True
      TabOrder = 0
      object BGraph: TToolbarButton97
        Left = 5
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Graphique'
        DisplayMode = dmGlyphOnly
        Caption = 'Graphique'
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
        OnClick = BGraphClick
        GlobalIndexImage = 'Z2172_S16G1'
      end
      object BStop: TToolbarButton97
        Left = 36
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Stopper le traitement'
        DisplayMode = dmGlyphOnly
        Caption = 'Stopper'
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
        Spacing = -1
        OnClick = BStopClick
        GlobalIndexImage = 'Z0107_S16G1'
        IsControl = True
      end
      object BImprimer: TToolbarButton97
        Left = 423
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        DisplayMode = dmGlyphOnly
        Caption = 'Imprimer'
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
      object BChercher: TToolbarButton97
        Left = 454
        Top = 3
        Width = 27
        Height = 27
        Hint = 'Lancer le traitement'
        DisplayMode = dmGlyphOnly
        Caption = 'Lancer'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BChercherClick
        GlobalIndexImage = 'Z0127_S16G1'
      end
      object BFerme: TToolbarButton97
        Left = 483
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
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
        GlyphMask.Data = {00000000}
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
      end
      object BAide: TToolbarButton97
        Left = 515
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
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
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
    end
  end
  object Pcrit: TPanel
    Left = 0
    Top = 0
    Width = 560
    Height = 58
    Align = alTop
    BevelInner = bvLowered
    TabOrder = 1
    object LExo: TLabel
      Left = 6
      Top = 9
      Width = 41
      Height = 13
      Caption = '&Exercice'
      FocusControl = CbExo
    end
    object LDevise: TLabel
      Left = 270
      Top = 9
      Width = 33
      Height = 13
      Caption = '&Devise'
      FocusControl = CbDevise
    end
    object LType: TLabel
      Left = 5
      Top = 36
      Width = 47
      Height = 13
      Caption = '&Type mvt.'
      FocusControl = CbTypMvt
    end
    object TEtab: THLabel
      Left = 270
      Top = 36
      Width = 34
      Height = 13
      Caption = 'E&tablis.'
      FocusControl = CbEtab
    end
    object CbExo: THValComboBox
      Left = 54
      Top = 5
      Width = 202
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = CbExoChange
      TagDispatch = 0
      DataType = 'TTEXERCICE'
    end
    object CbDevise: THValComboBox
      Left = 307
      Top = 5
      Width = 146
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 1
      OnChange = CbDeviseChange
      TagDispatch = 0
      Vide = True
      DataType = 'TTDEVISE'
    end
    object CbTypMvt: THValComboBox
      Left = 54
      Top = 32
      Width = 202
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 2
      TagDispatch = 0
      DataType = 'TTQUALPIECECRIT'
    end
    object CbEtab: THValComboBox
      Left = 307
      Top = 32
      Width = 146
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 3
      OnChange = CbEtabChange
      TagDispatch = 0
      Vide = True
      DataType = 'TTETABLISSEMENT'
    end
    object Rev: TCheckBox
      Left = 464
      Top = 7
      Width = 65
      Height = 17
      Alignment = taLeftJustify
      AllowGrayed = True
      Caption = '&R'#233'vision'
      State = cbGrayed
      TabOrder = 4
    end
  end
  object Calculsolde: THNumEdit
    Left = 176
    Top = 205
    Width = 81
    Height = 21
    Color = clYellow
    Decimals = 2
    Digits = 15
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Masks.PositiveMask = '#,##0.00'
    Debit = False
    NumericType = ntDC
    ParentFont = False
    TabOrder = 2
    UseRounding = True
    Validate = False
  end
  object CbMess: TComboBox
    Left = 332
    Top = 224
    Width = 73
    Height = 21
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ItemHeight = 13
    ParentFont = False
    TabOrder = 3
    Visible = False
    Items.Strings = (
      'Calcul en cours'
      'Total'
      'R'#233'sultat de l'#39'exercice')
  end
  object CBFinDate: TComboBox
    Left = 110
    Top = 106
    Width = 131
    Height = 21
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ItemHeight = 13
    ParentFont = False
    TabOrder = 6
    Visible = False
  end
  object CBDebDat: TComboBox
    Left = 110
    Top = 130
    Width = 131
    Height = 21
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ItemHeight = 13
    ParentFont = False
    TabOrder = 7
    Visible = False
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 256
    Top = 244
  end
  object Timer1: TTimer
    Interval = 500
    OnTimer = Timer1Timer
    Left = 33
    Top = 174
  end
end

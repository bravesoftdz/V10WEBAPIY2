object FVentil: TFVentil
  Left = 316
  Top = 193
  HelpContext = 1460100
  BorderIcons = [biSystemMenu, biMaximize]
  BorderStyle = bsSingle
  Caption = 'Ventilation par d'#233'faut de : '
  ClientHeight = 284
  ClientWidth = 545
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clBlack
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object HPanel1: TPanel
    Left = 0
    Top = 249
    Width = 545
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    Caption = ' '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    object TCache: TLabel
      Left = 236
      Top = 4
      Width = 38
      Height = 13
      Caption = 'TCache'
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      Visible = False
    end
    object TFType: THLabel
      Left = 5
      Top = 11
      Width = 72
      Height = 13
      Caption = 'Ventilation type'
      FocusControl = FType
    end
    object FType: THValComboBox
      Left = 85
      Top = 7
      Width = 151
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnClick = FTypeClick
      TagDispatch = 0
      DataType = 'TTVENTILTYPE'
    end
    object Panel6: TPanel
      Left = 316
      Top = 2
      Width = 227
      Height = 31
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 1
      object BDelLigne: THBitBtn
        Left = 68
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Supprimer une ligne'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnClick = BDelLigneClick
        GlobalIndexImage = 'Z0072_S16G1'
      end
      object BInsLigne: THBitBtn
        Left = 36
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Ins'#233'rer une ligne'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = BInsLigneClick
        GlobalIndexImage = 'Z0074_S16G1'
      end
      object BImprimer: THBitBtn
        Left = 100
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = BImprimerClick
        Margin = 2
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object OKBtn: THBitBtn
        Left = 131
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Valider'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnClick = OKBtnClick
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
        Spacing = -1
        IsControl = True
      end
      object BFermer: THBitBtn
        Left = 163
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        ParentShowHint = False
        ShowHint = True
        TabOrder = 4
        OnClick = BFermerClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object HelpBtn: THBitBtn
        Left = 195
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 5
        OnClick = HelpBtnClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BCalculUOE: THBitBtn
        Left = 4
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Calculer les pourcentages'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 6
        OnClick = BCalculUOEClick
        GlobalIndexImage = 'Z0123_S16G1'
      end
    end
  end
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 545
    Height = 249
    ActivePage = PAxe1
    Align = alClient
    TabOrder = 1
    object PAxe1: TTabSheet
      Caption = 'Axe n'#176'1'
      object FListe1: THGrid
        Left = 0
        Top = 0
        Width = 537
        Height = 194
        Align = alClient
        ColCount = 7
        DefaultColWidth = 20
        DefaultRowHeight = 18
        RowCount = 10
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goEditing, goTabs, goAlwaysShowEditor]
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 0
        OnDblClick = FListe1DblClick
        OnKeyDown = FListe1KeyDown
        SortedCol = -1
        Titres.Strings = (
          'N'#176';C'
          'Section;G'
          'Libell'#233';G'
          'UO;D'
          '% Valeur;D'
          '% Qt'#233' 1;D'
          '% Qt'#233' 2;D')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = FListe1RowEnter
        OnCellEnter = FListe1CellEnter
        OnCellExit = FListe1CellExit
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        ColWidths = (
          20
          70
          168
          63
          60
          60
          56)
      end
      object Panel1: TPanel
        Left = 0
        Top = 194
        Width = 537
        Height = 27
        Align = alBottom
        BevelOuter = bvNone
        Caption = ' '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object HLabel1: THLabel
          Left = 216
          Top = 7
          Width = 40
          Height = 13
          Caption = 'Totaux'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object FTotQte21: THNumEdit
          Left = 451
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Ctl3D = True
          Enabled = False
          ParentCtl3D = False
          TabOrder = 0
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object FTotQte11: THNumEdit
          Left = 389
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Ctl3D = True
          Enabled = False
          ParentCtl3D = False
          TabOrder = 1
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object FTotVal1: THNumEdit
          Left = 327
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Ctl3D = True
          Enabled = False
          ParentCtl3D = False
          TabOrder = 2
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object FTotalUOe1: THNumEdit
          Left = 265
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Ctl3D = True
          Enabled = False
          ParentCtl3D = False
          TabOrder = 3
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
      end
    end
    object PAxe2: TTabSheet
      Caption = 'Axe n'#176'2'
      object Panel2: TPanel
        Left = 0
        Top = 194
        Width = 537
        Height = 27
        Align = alBottom
        BevelOuter = bvNone
        Caption = ' '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object HLabel2: THLabel
          Left = 216
          Top = 7
          Width = 40
          Height = 13
          Caption = 'Totaux'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object FTotQte22: THNumEdit
          Left = 451
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Enabled = False
          TabOrder = 0
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object FTotQte12: THNumEdit
          Left = 389
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Enabled = False
          TabOrder = 1
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object FTotVal2: THNumEdit
          Left = 327
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Enabled = False
          TabOrder = 2
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object FTotalUOe2: THNumEdit
          Left = 264
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Enabled = False
          TabOrder = 3
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
      end
      object FListe2: THGrid
        Left = 0
        Top = 0
        Width = 537
        Height = 194
        Align = alClient
        ColCount = 7
        DefaultColWidth = 20
        DefaultRowHeight = 18
        RowCount = 10
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goEditing, goTabs, goAlwaysShowEditor]
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 1
        OnDblClick = FListe1DblClick
        OnKeyDown = FListe1KeyDown
        SortedCol = -1
        Titres.Strings = (
          'N'#176';C'
          'Section;G'
          'Libell'#233';G'
          'UO;D'
          '% Valeur;D'
          '% Qt'#233' 1;D'
          '% Qt'#233' 2;D')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = FListe1RowEnter
        OnCellEnter = FListe1CellEnter
        OnCellExit = FListe1CellExit
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        ColWidths = (
          20
          70
          168
          62
          60
          60
          59)
      end
    end
    object PAxe3: TTabSheet
      Caption = 'Axe n'#176'3'
      object Panel3: TPanel
        Left = 0
        Top = 194
        Width = 537
        Height = 27
        Align = alBottom
        BevelOuter = bvNone
        Caption = ' '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object HLabel3: THLabel
          Left = 216
          Top = 7
          Width = 40
          Height = 13
          Caption = 'Totaux'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object FTotQte23: THNumEdit
          Left = 451
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Enabled = False
          TabOrder = 0
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object FTotQte13: THNumEdit
          Left = 389
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Enabled = False
          TabOrder = 1
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object FTotVal3: THNumEdit
          Left = 327
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Enabled = False
          TabOrder = 2
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object FTotalUOe3: THNumEdit
          Left = 264
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Enabled = False
          TabOrder = 3
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
      end
      object FListe3: THGrid
        Left = 0
        Top = 0
        Width = 537
        Height = 194
        Align = alClient
        ColCount = 7
        DefaultColWidth = 20
        DefaultRowHeight = 18
        RowCount = 10
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goEditing, goTabs, goAlwaysShowEditor]
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 1
        OnDblClick = FListe1DblClick
        OnKeyDown = FListe1KeyDown
        SortedCol = -1
        Titres.Strings = (
          'N'#176';C'
          'Section;G'
          'Libell'#233';G'
          'UO;D'
          '% Valeur;D'
          '% Qt'#233' 1;D'
          '% Qt'#233' 2;D')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = FListe1RowEnter
        OnCellEnter = FListe1CellEnter
        OnCellExit = FListe1CellExit
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        ColWidths = (
          20
          70
          168
          60
          60
          60
          53)
      end
    end
    object PAxe4: TTabSheet
      Caption = 'Axe n'#176'4'
      object Panel4: TPanel
        Left = 0
        Top = 194
        Width = 537
        Height = 27
        Align = alBottom
        BevelOuter = bvNone
        Caption = ' '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object HLabel4: THLabel
          Left = 216
          Top = 7
          Width = 40
          Height = 13
          Caption = 'Totaux'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object FTotQte24: THNumEdit
          Left = 451
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Enabled = False
          TabOrder = 0
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object FTotQte14: THNumEdit
          Left = 389
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Enabled = False
          TabOrder = 1
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object FTotVal4: THNumEdit
          Left = 327
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Enabled = False
          TabOrder = 2
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object FTotalUOe4: THNumEdit
          Left = 264
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Enabled = False
          TabOrder = 3
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
      end
      object FListe4: THGrid
        Left = 0
        Top = 0
        Width = 537
        Height = 194
        Align = alClient
        ColCount = 7
        DefaultColWidth = 20
        DefaultRowHeight = 18
        RowCount = 10
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goEditing, goTabs, goAlwaysShowEditor]
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 1
        OnDblClick = FListe1DblClick
        OnKeyDown = FListe1KeyDown
        SortedCol = -1
        Titres.Strings = (
          'N'#176';C'
          'Section;G'
          'Libell'#233';G'
          'UO;D'
          '% Valeur;D'
          '% Qt'#233' 1;D'
          '% Qt'#233' 2;D')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = FListe1RowEnter
        OnCellEnter = FListe1CellEnter
        OnCellExit = FListe1CellExit
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        ColWidths = (
          20
          70
          168
          60
          60
          60
          58)
      end
    end
    object PAxe5: TTabSheet
      Caption = 'Axe n'#176'5'
      object Panel5: TPanel
        Left = 0
        Top = 194
        Width = 537
        Height = 27
        Align = alBottom
        BevelOuter = bvNone
        Caption = ' '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object HLabel5: THLabel
          Left = 216
          Top = 7
          Width = 40
          Height = 13
          Caption = 'Totaux'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object FTotQte25: THNumEdit
          Left = 451
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Enabled = False
          TabOrder = 0
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object FTotQte15: THNumEdit
          Left = 389
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Enabled = False
          TabOrder = 1
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object FTotVal5: THNumEdit
          Left = 327
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Enabled = False
          TabOrder = 2
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object FTotalUOe5: THNumEdit
          Left = 265
          Top = 3
          Width = 60
          Height = 21
          Color = clBtnFace
          Enabled = False
          TabOrder = 3
          Decimals = 3
          Digits = 12
          Masks.PositiveMask = '# ##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
      end
      object FListe5: THGrid
        Left = 0
        Top = 0
        Width = 537
        Height = 194
        Align = alClient
        ColCount = 7
        DefaultColWidth = 20
        DefaultRowHeight = 18
        RowCount = 10
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goEditing, goTabs, goAlwaysShowEditor]
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 1
        OnDblClick = FListe1DblClick
        OnKeyDown = FListe1KeyDown
        SortedCol = -1
        Titres.Strings = (
          'N'#176';C'
          'Section;G'
          'Libell'#233';G'
          'UO;D'
          '% Valeur;D'
          '% Qt'#233' 1;D'
          '% Qt'#233' 2;D')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnRowEnter = FListe1RowEnter
        OnCellEnter = FListe1CellEnter
        OnCellExit = FListe1CellExit
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = True
        AlternateColor = clSilver
        ColWidths = (
          20
          70
          168
          60
          60
          60
          59)
      end
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Ventilations analytiques;Voulez-vous enregistrer les modificat' +
        'ions?;Q;YNC;Y;C;'
      
        '1;Ventilations analytiques;ATTENTION : La section est ferm'#233'e.;E;' +
        'O;O;O;'
      'Sections r'#233'ceptrices de :'
      'Ventilation par d'#233'faut de :'
      'Ventilation type :'
      'Ventilation du compte de charge'
      'Ventilation du compte de produit'
      'Choix d'#39'une section analytique'
      '8;')
    Left = 212
    Top = 116
  end
  object HMTrad: THSystemMenu
    LockedCtrls.Strings = (
      'HPanel1'
      'BDelLigne'
      'BInsLigne'
      'BImprimer'
      'OkBtn'
      'BFermer'
      'HelpBtn')
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 92
    Top = 112
  end
end

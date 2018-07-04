object FDotInt: TFDotInt
  Left = 370
  Top = 213
  HelpContext = 999999961
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Dotations interm'#233'diaires'
  ClientHeight = 245
  ClientWidth = 516
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  OnClose = FormClose
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object HPanel2: THPanel
    Left = 0
    Top = 0
    Width = 516
    Height = 112
    Align = alClient
    Caption = 'HPanel2'
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object HLabel1: THLabel
      Left = 148
      Top = 14
      Width = 114
      Height = 13
      Caption = 'Calcul de la dotation au '
    end
    object fDateCalcul: THCritMaskEdit
      Left = 264
      Top = 10
      Width = 66
      Height = 21
      Hint = 'Date de l'#39'op'#233'ration'
      EditMask = '!99/99/0000;1;_'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 10
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      Text = '  /  /    '
      OnKeyDown = FormKeyDown
      OnKeyPress = fDateCalculKeyPress
      TagDispatch = 0
      OpeType = otDate
      ControlerDate = True
    end
    object fGrid: THGrid
      Left = 12
      Top = 40
      Width = 487
      Height = 72
      DefaultRowHeight = 16
      RowCount = 3
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSelect]
      TabOrder = 1
      OnKeyDown = FormKeyDown
      SortedCol = -1
      Titres.Strings = (
        'Amortissement'
        'Cumul ant'#233'rieur'
        'Dotation calc.'
        'Cumul amort.'
        'V.N.C.'
        ''
        '')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = clSilver
      ColWidths = (
        89
        95
        97
        98
        99)
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 206
    Width = 516
    Height = 39
    AllowDrag = False
    Position = dpBottom
    object PBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 35
      ClientWidth = 516
      Caption = 'Barre outils fiche'
      ClientAreaHeight = 35
      ClientAreaWidth = 516
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object ToolbarButton971: TToolbarButton97
        Left = 386
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Imprimer'
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
      object BAPPLIQUER: TToolbarButton97
        Left = 418
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Valider'
        Default = True
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0127_S16G1'
        IsControl = True
      end
      object ToolbarButton974: TToolbarButton97
        Left = 450
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = bFermeClick
        GlobalIndexImage = 'Z1770_S16G1'
      end
      object BoutonAide: TToolbarButton97
        Left = 482
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
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
        OnClick = BoutonAideClick
        IsControl = True
      end
    end
  end
  object pAmortDerog: TPanel
    Left = 0
    Top = 112
    Width = 516
    Height = 94
    Align = alBottom
    TabOrder = 2
    Visible = False
    object tAmortDerog: THLabel
      Left = 8
      Top = 6
      Width = 168
      Height = 13
      Caption = 'Amortissements d'#233'rogatoire  (145) : '
    end
    object AmortDerog: THLabel
      Left = 193
      Top = 6
      Width = 56
      Height = 13
      Alignment = taRightJustify
      Caption = 'AmortDerog'
    end
    object tReintDeduc: THLabel
      Left = 8
      Top = 28
      Width = 133
      Height = 13
      Caption = 'R'#233'int'#233'gration/d'#233'duction :  : '
    end
    object ReintDeduc: THLabel
      Left = 192
      Top = 28
      Width = 57
      Height = 13
      Alignment = taRightJustify
      Caption = 'ReintDeduc'
    end
    object typederog: THLabel
      Left = 256
      Top = 6
      Width = 47
      Height = 13
      Caption = 'typederog'
    end
    object typereint: THLabel
      Left = 256
      Top = 28
      Width = 40
      Height = 13
      Caption = 'typereint'
    end
    object InfoReint: THLabel
      Left = 264
      Top = 60
      Width = 17
      Height = 13
      Caption = 'info'
    end
    object InfoDeduc: THLabel
      Left = 264
      Top = 44
      Width = 17
      Height = 13
      Caption = 'info'
    end
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Economique'
      'Fiscal'
      
        '2;?Caption?;La date de calcul doit '#234'tre une date d'#39'un exercice o' +
        'uvert.;W;O;O;O;'
      '3;?Caption?;Date incorrecte.;W;O;O;O;'
      '(Dotation)'
      '(Reprise)'
      '(R'#233'int'#233'gration)'
      '(D'#233'duction)')
    Left = 376
    Top = 121
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 328
    Top = 121
  end
end

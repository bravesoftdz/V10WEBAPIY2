object FPlanEche: TFPlanEche
  Left = 325
  Top = 203
  HelpContext = 2101400
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Saisie group'#233'e de l'#39#233'ch'#233'ancier'
  ClientHeight = 264
  ClientWidth = 455
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object HPanel3: THPanel
    Left = 0
    Top = 192
    Width = 455
    Height = 37
    Align = alBottom
    BevelOuter = bvNone
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object HLabel3: THLabel
      Left = 131
      Top = 12
      Width = 76
      Height = 13
      Caption = 'Total g'#233'n'#233'ral'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object TotalValeur: THNumEdit
      Left = 270
      Top = 8
      Width = 95
      Height = 21
      Color = clInactiveBorder
      Ctl3D = True
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Masks.PositiveMask = '#,##0.00'
      Masks.ZeroMask = '0.00'
      Debit = False
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 0
      UseRounding = True
      Validate = False
    end
    object TotalFrais: THNumEdit
      Left = 370
      Top = 8
      Width = 70
      Height = 21
      Color = clInactiveBorder
      Ctl3D = True
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Masks.PositiveMask = '#,##0.00'
      Masks.ZeroMask = '0.00'
      Debit = False
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 1
      UseRounding = True
      Validate = False
    end
    object TotalNombre: THNumEdit
      Left = 216
      Top = 8
      Width = 49
      Height = 21
      Color = clInactiveBorder
      Ctl3D = True
      Decimals = 0
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Masks.PositiveMask = '#'
      Masks.ZeroMask = '0'
      Debit = False
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 2
      UseRounding = True
      Validate = False
    end
  end
  object GridSaisie: THGrid
    Left = 0
    Top = 45
    Width = 455
    Height = 147
    Align = alClient
    ColCount = 6
    DefaultColWidth = 90
    DefaultRowHeight = 16
    RowCount = 3
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goEditing, goTabs]
    PopupMenu = PopupMenu
    TabOrder = 1
    OnGetEditMask = GridSaisieGetEditMask
    SortedCol = -1
    Titres.Strings = (
      ''
      'P'#233'riode du'
      'au'
      'Nombre '
      'Valeur '
      'Frais')
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    OnRowEnter = GridSaisieRowEnter
    OnRowExit = GridSaisieRowExit
    OnColExit = GridSaisieColExit
    OnCellEnter = GridSaisieCellEnter
    OnCellExit = GridSaisieCellExit
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = clSilver
    ColWidths = (
      19
      92
      90
      50
      99
      75)
  end
  object Dock971: TDock97
    Left = 0
    Top = 229
    Width = 455
    Height = 35
    AllowDrag = False
    Position = dpBottom
    object PBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 455
      Caption = 'Barre outils fiche'
      ClientAreaHeight = 31
      ClientAreaWidth = 455
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BAbandon: TToolbarButton97
        Left = 389
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        Flat = False
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BAbandonClick
        GlobalIndexImage = 'Z1770_S16G1'
      end
      object HelpBtn: TToolbarButton97
        Left = 421
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
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = HelpBtnClick
        IsControl = True
      end
      object BImprimer: TToolbarButton97
        Left = 326
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
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BValide: TToolbarButton97
        Left = 358
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
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValideClick
        GlobalIndexImage = 'Z0127_S16G1'
        IsControl = True
      end
      object BInserer: TToolbarButton97
        Left = 2
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Ins'#233'rer'
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
        OnClick = InsererClick
        GlobalIndexImage = 'Z0074_S16G1'
      end
      object BSupprimer: TToolbarButton97
        Left = 34
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Supprimer'
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
        OnClick = SupprimerClick
        GlobalIndexImage = 'Z0072_S16G1'
      end
    end
  end
  object HPanel1: THPanel
    Left = 0
    Top = 0
    Width = 455
    Height = 45
    Align = alTop
    BevelOuter = bvNone
    FullRepaint = False
    TabOrder = 3
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
    object HLabel1: THLabel
      Left = 24
      Top = 16
      Width = 49
      Height = 13
      Caption = 'P'#233'riodicit'#233
    end
    object HLabel2: THLabel
      Left = 181
      Top = 16
      Width = 91
      Height = 13
      Caption = 'Type de versement'
    end
    object PERIODICITE: TEdit
      Left = 80
      Top = 12
      Width = 73
      Height = 21
      Enabled = False
      TabOrder = 0
      Text = 'PERIODICITE'
    end
    object VERSEMENT: TEdit
      Left = 280
      Top = 12
      Width = 65
      Height = 21
      Enabled = False
      TabOrder = 1
      Text = 'VERSEMENT'
    end
  end
  object HM1: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?Caption?;Date d'#39#233'ch'#233'ance incorrecte.;E;O;O;O;'
      '1;?Caption?;Tableau d'#39#233'ch'#233'ances incorrect;E;O;O;O;'
      
        '2;?Caption?;Incoh'#233'rence dans le plan d'#39#233'ch'#233'ances. Veuillez v'#233'rif' +
        'ier les dates.;E;O;O;O;'
      '3;?Caption?;Ligne incompl'#232'te;E;O;O;O;'
      
        '4;?Caption?;Une ligne vide a '#233't'#233' d'#233'tect'#233'e. Veuillez la supprimer' +
        '.;E;O;O;O;'
      
        '5;?caption?;Etes-vous s'#251'r de sortir sans sauvegarder les modific' +
        'ations?;Q;YN;N;N;')
    Left = 82
    Top = 214
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 40
    Top = 104
  end
  object PopupMenu: TPopupMenu
    Left = 176
    Top = 136
    object Inserer: TMenuItem
      Caption = '&Ins'#233'rer ligne'
      OnClick = InsererClick
    end
    object Supprimer: TMenuItem
      Caption = '&Supprimer ligne'
      OnClick = SupprimerClick
    end
  end
end

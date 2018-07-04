object FCasEmplois: TFCasEmplois
  Left = 326
  Top = 181
  Width = 640
  Height = 480
  HelpContext = 11000249
  Caption = 'Cas d'#39'emploi'
  Color = clBtnFace
  Constraints.MinHeight = 440
  Constraints.MinWidth = 620
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PANChoix: TPanel
    Left = 0
    Top = 0
    Width = 624
    Height = 65
    Align = alTop
    TabOrder = 0
    object Label1: TLabel
      Left = 32
      Top = 12
      Width = 57
      Height = 13
      Caption = 'Code Article'
      FocusControl = GA_CODEARTICLE
    end
    object Label2: TLabel
      Left = 300
      Top = 12
      Width = 91
      Height = 13
      Caption = 'Code article unique'
    end
    object Label3: TLabel
      Left = 300
      Top = 39
      Width = 92
      Height = 13
      Caption = 'Code nomenclature'
    end
    object GA_CODEARTICLE: THCritMaskEdit
      Left = 116
      Top = 8
      Width = 121
      Height = 21
      TabOrder = 0
      OnExit = GA_CODEARTICLEExit
      TagDispatch = 0
      ElipsisButton = True
      OnElipsisClick = GA_CODEARTICLEElipsisClick
    end
    object GA_ARTICLE: TEdit
      Left = 428
      Top = 8
      Width = 152
      Height = 21
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 1
    end
    object GNE_NOMENCLATURE: TEdit
      Left = 428
      Top = 35
      Width = 152
      Height = 21
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 2
    end
  end
  object TV_NLIG: TTreeView
    Left = 393
    Top = 65
    Width = 231
    Height = 340
    Align = alClient
    HideSelection = False
    Indent = 19
    PopupMenu = PopupTV_NLIG
    TabOrder = 2
    OnMouseUp = TV_NLIGMouseUp
  end
  object G_NLIG: THGrid
    Left = 0
    Top = 65
    Width = 393
    Height = 340
    Align = alLeft
    ColCount = 3
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 50
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goRowSelect]
    ParentShowHint = False
    PopupMenu = PopupG_NLIG
    ShowHint = True
    TabOrder = 1
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    OnRowEnter = G_NLIGRowEnter
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = clSilver
    ElipsisButton = True
    ColWidths = (
      120
      187
      64)
  end
  object Dock971: TDock97
    Left = 0
    Top = 405
    Width = 624
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object ToolWindow971: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 624
      Caption = 'ToolWindow971'
      ClientAreaHeight = 32
      ClientAreaWidth = 624
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BAide: TToolbarButton97
        Tag = 1
        Left = 593
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        Caption = 'Aide'
        DisplayMode = dmGlyphOnly
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
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 563
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Caption = 'Annuler'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ModalResult = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BValider: TToolbarButton97
        Left = 533
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Enregistrer la nomenclature'
        Caption = 'Enregistrer'
        Default = True
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
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
        GlyphMask.Data = {00000000}
        Layout = blGlyphTop
        ModalResult = 1
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        IsControl = True
      end
      object bPrint: TToolbarButton97
        Left = 503
        Top = 3
        Width = 28
        Height = 27
        Flat = False
        OnClick = bPrintClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BInfos: TToolbarButton97
        Tag = 1
        Left = 10
        Top = 3
        Width = 40
        Height = 27
        Hint = 'Compl'#233'ments d'#39'informations'
        Caption = 'Compl'#233'ments'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = PopupG_NLIG
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        ParentFont = False
        Spacing = -1
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 532
    Top = 76
  end
  object PopupG_NLIG: TPopupMenu
    Left = 504
    Top = 76
    object PopG_NLIG_A: TMenuItem
      Caption = 'Fiche &Article'
      ShortCut = 16449
      OnClick = PopG_NLIG_AClick
    end
    object PopG_NLIG_N: TMenuItem
      Caption = '&Nomenclature'
      ShortCut = 16462
      OnClick = PopG_NLIG_NClick
    end
  end
  object PopupTV_NLIG: TPopupMenu
    Left = 476
    Top = 76
    object PopTV_NLIG_O: TMenuItem
      Caption = 'Tout &Ouvrir'
      ShortCut = 16468
      OnClick = PopTV_NLIG_OClick
    end
    object PopTV_NLIG_F: TMenuItem
      Caption = '&Fermer sous niveaux'
      ShortCut = 16454
      OnClick = PopTV_NLIG_FClick
    end
  end
end

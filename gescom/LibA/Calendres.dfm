object FCalendrier: TFCalendrier
  Left = 56
  Top = 21
  HelpContext = 1700
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Calendrier'
  ClientHeight = 499
  ClientWidth = 683
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnDestroy = FormDestroy
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object LabelMois: TLabel
    Left = 177
    Top = 29
    Width = 300
    Height = 24
    Alignment = taCenter
    AutoSize = False
    Caption = 'Mois en cours'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -19
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object bNext: TToolbarButton97
    Left = 479
    Top = 29
    Width = 23
    Height = 22
    Glyph.Data = {
      36030000424D3603000000000000360000002800000010000000100000000100
      1800000000000003000000000000000000000000000000000000C6C3C6C6C3C6
      C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3
      C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6
      C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6
      C6C3C6C6C3C6C6C3C6848284848284C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3
      C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6FFFFFF00000000000084
      8284C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6
      C6C3C6C6C3C6FFFFFF000000000000000000848284C6C3C6C6C3C6C6C3C6C6C3
      C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6FFFFFF00000000000000
      0000000000848284C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6
      C6C3C6C6C3C6FFFFFF000000000000000000000000000000848284C6C3C6C6C3
      C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6FFFFFF00000000000000
      0000000000000000000000848284C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6
      C6C3C6C6C3C6FFFFFF000000000000000000000000000000000000C6C3C6C6C3
      C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6FFFFFF00000000000000
      0000000000000000C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6
      C6C3C6C6C3C6FFFFFF000000000000000000000000C6C3C6C6C3C6C6C3C6C6C3
      C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6FFFFFF00000000000000
      0000C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6
      C6C3C6C6C3C6FFFFFF000000000000C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3
      C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6FFFFFFFFFFFFC6C3C6C6
      C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6
      C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3
      C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6
      C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6C6C3C6}
    GlyphMask.Data = {00000000}
    OnClick = bNextClick
  end
  object bLast: TToolbarButton97
    Left = 501
    Top = 29
    Width = 23
    Height = 22
    OnClick = bLastClick
    GlobalIndexImage = 'Z0092_S16G1'
  end
  object bPrevious: TToolbarButton97
    Left = 151
    Top = 29
    Width = 23
    Height = 22
    OnClick = bPreviousClick
    GlobalIndexImage = 'Z0017_S16G1'
  end
  object bFirst: TToolbarButton97
    Left = 129
    Top = 29
    Width = 23
    Height = 22
    OnClick = bFirstClick
    GlobalIndexImage = 'Z0095_S16G1'
  end
  object LabelStandard: THLabel
    Left = 4
    Top = 4
    Width = 122
    Height = 16
    Caption = 'Type de Calendrier :'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object NomStandard: THLabel
    Left = 132
    Top = 4
    Width = 84
    Height = 16
    Caption = 'NomStandard'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Cal: THGrid
    Left = 3
    Top = 52
    Width = 466
    Height = 317
    Color = clWhite
    ColCount = 7
    Ctl3D = True
    DefaultColWidth = 90
    DefaultRowHeight = 60
    DefaultDrawing = False
    FixedCols = 0
    RowCount = 7
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected]
    ParentCtl3D = False
    ScrollBars = ssVertical
    TabOrder = 0
    SortedCol = -1
    Couleur = False
    MultiSelect = True
    TitleBold = True
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = clSilver
  end
  object FenHoraire: TToolWindow97
    Left = 475
    Top = 63
    ClientHeight = 135
    ClientWidth = 169
    Caption = 'Horaires'
    ClientAreaHeight = 135
    ClientAreaWidth = 169
    TabOrder = 1
    object bValWindows: TToolbarButton97
      Left = 115
      Top = 112
      Width = 23
      Height = 22
      Flat = False
      OnClick = bValWindowsClick
      GlobalIndexImage = 'Z0127_S16G1'
    end
    object bAnnulWindows: TToolbarButton97
      Left = 141
      Top = 112
      Width = 23
      Height = 22
      Flat = False
      OnClick = bAnnulWindowsClick
      GlobalIndexImage = 'Z0021_S16G1'
    end
    object LDEB: THLabel
      Left = 3
      Top = 12
      Width = 14
      Height = 13
      Caption = '&De'
    end
    object LDEB2: THLabel
      Left = 3
      Top = 38
      Width = 14
      Height = 13
      Caption = 'D&e'
    end
    object LFIN1: THLabel
      Left = 83
      Top = 12
      Width = 6
      Height = 13
      Caption = #224
    end
    object LFIN2: THLabel
      Left = 83
      Top = 38
      Width = 6
      Height = 13
      Caption = #224
    end
    object LDUREE: THLabel
      Left = 7
      Top = 65
      Width = 29
      Height = 13
      Caption = 'D&ur'#233'e'
      FocusControl = HDuree
    end
    object Hdeb1: THCritMaskEdit
      Left = 28
      Top = 8
      Width = 45
      Height = 21
      TabOrder = 0
      OnEnter = Hdeb1Enter
      OnExit = Hdeb1Exit
      TagDispatch = 0
      OpeType = otReel
    end
    object Hfin1: THCritMaskEdit
      Left = 98
      Top = 8
      Width = 45
      Height = 21
      TabOrder = 1
      OnEnter = Hdeb1Enter
      OnExit = Hdeb1Exit
      TagDispatch = 0
      OpeType = otReel
    end
    object HDeb2: THCritMaskEdit
      Left = 28
      Top = 34
      Width = 45
      Height = 21
      TabOrder = 2
      OnEnter = Hdeb1Enter
      OnExit = Hdeb1Exit
      TagDispatch = 0
      OpeType = otReel
    end
    object HFin2: THCritMaskEdit
      Left = 98
      Top = 34
      Width = 45
      Height = 21
      TabOrder = 3
      OnEnter = Hdeb1Enter
      OnExit = Hdeb1Exit
      TagDispatch = 0
      OpeType = otReel
    end
    object HDuree: THCritMaskEdit
      Left = 53
      Top = 61
      Width = 45
      Height = 21
      TabOrder = 4
      TagDispatch = 0
      OpeType = otReel
    end
    object CheckTravailFerie: TCheckBox
      Left = 8
      Top = 91
      Width = 161
      Height = 17
      Caption = '&Travail autoris'#233' ce jour f'#233'ri'#233
      TabOrder = 5
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 463
    Width = 683
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object ToolCalendrier: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 683
      Caption = 'ToolCalendrier'
      ClientAreaHeight = 32
      ClientAreaWidth = 683
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object bValider: TToolbarButton97
        Left = 586
        Top = 2
        Width = 28
        Height = 27
        Flat = False
        OnClick = bValiderClick
        GlobalIndexImage = 'Z0127_S16G1'
      end
      object bAnnuler: TToolbarButton97
        Left = 618
        Top = 2
        Width = 28
        Height = 27
        Flat = False
        OnClick = bAnnulerClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object HelpBtn: TToolbarButton97
        Left = 650
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        HelpContext = 120000502
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
        Spacing = -1
        OnClick = HelpBtnClick
        IsControl = True
      end
      object BImprimer: TToolbarButton97
        Left = 523
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
        Visible = False
        OnClick = BImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object bDetail: TToolbarButton97
        Left = 555
        Top = 2
        Width = 28
        Height = 27
        Hint = 'D'#233'tail journalier'
        DisplayMode = dmGlyphOnly
        Caption = 'D'#233'tail'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        OnClick = bDetailClick
        GlobalIndexImage = 'Z0665_S16G1'
      end
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 628
    Top = 376
  end
end

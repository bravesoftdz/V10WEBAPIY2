object FNomenAPLat: TFNomenAPLat
  Left = 310
  Top = 101
  Width = 620
  Height = 440
  HelpContext = 110000240
  Caption = 'Nomenclature Mise '#224' plat'
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
  object PANNOMENENT: TPanel
    Left = 0
    Top = 0
    Width = 604
    Height = 105
    Align = alTop
    TabOrder = 0
    object TGNE_ARTICLE: TLabel
      Left = 312
      Top = 12
      Width = 103
      Height = 13
      Caption = 'Code Article compos'#233
      FocusControl = GNE_ARTICLE
    end
    object TGNE_NOMENCLATURE: TLabel
      Left = 13
      Top = 12
      Width = 94
      Height = 13
      Caption = 'Code Nomenclature'
      FocusControl = GNE_NOMENCLATURE
    end
    object TGNE_LIBELLE: TLabel
      Left = 13
      Top = 36
      Width = 123
      Height = 13
      Caption = 'Libell'#233' de la nomenclature'
      FocusControl = GNE_LIBELLE
    end
    object Label1: TLabel
      Left = 24
      Top = 60
      Width = 22
      Height = 13
      Caption = 'DPA'
    end
    object Label2: TLabel
      Left = 24
      Top = 83
      Width = 23
      Height = 13
      Caption = 'DPR'
    end
    object Label3: TLabel
      Left = 176
      Top = 62
      Width = 29
      Height = 13
      Caption = 'PVHT'
    end
    object Label4: TLabel
      Left = 176
      Top = 83
      Width = 35
      Height = 13
      Caption = 'PVTTC'
    end
    object Label5: TLabel
      Left = 320
      Top = 60
      Width = 29
      Height = 13
      Caption = 'PAHT'
    end
    object Label6: TLabel
      Left = 320
      Top = 83
      Width = 30
      Height = 13
      Caption = 'PRHT'
    end
    object Label7: TLabel
      Left = 464
      Top = 60
      Width = 30
      Height = 13
      Caption = 'PMAP'
    end
    object Label8: TLabel
      Left = 464
      Top = 83
      Width = 31
      Height = 13
      Caption = 'PMRP'
    end
    object GNE_ARTICLE: TEdit
      Left = 441
      Top = 8
      Width = 121
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 0
    end
    object GNE_NOMENCLATURE: TEdit
      Left = 141
      Top = 8
      Width = 121
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 1
    end
    object GNE_LIBELLE: TEdit
      Left = 141
      Top = 32
      Width = 421
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 2
    end
    object GNE_DPA: THCritMaskEdit
      Left = 76
      Top = 56
      Width = 85
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 3
      TagDispatch = 0
      OpeType = otReel
    end
    object GNE_DPR: THCritMaskEdit
      Left = 76
      Top = 80
      Width = 85
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 4
      TagDispatch = 0
      OpeType = otReel
    end
    object GNE_PVHT: THCritMaskEdit
      Left = 216
      Top = 56
      Width = 85
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 5
      TagDispatch = 0
      OpeType = otReel
    end
    object GNE_PVTTC: THCritMaskEdit
      Left = 216
      Top = 80
      Width = 85
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 6
      TagDispatch = 0
      OpeType = otReel
    end
    object GNE_PRHT: THCritMaskEdit
      Left = 360
      Top = 80
      Width = 85
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 7
      TagDispatch = 0
      OpeType = otReel
    end
    object GNE_PMAP: THCritMaskEdit
      Left = 504
      Top = 56
      Width = 85
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 8
      TagDispatch = 0
      OpeType = otReel
    end
    object GNE_PMRP: THCritMaskEdit
      Left = 504
      Top = 80
      Width = 85
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 9
      TagDispatch = 0
      OpeType = otReel
    end
    object GNE_PAHT: THCritMaskEdit
      Left = 360
      Top = 56
      Width = 85
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 10
      TagDispatch = 0
      OpeType = otReel
    end
  end
  object PANPIED: TPanel
    Left = 0
    Top = 315
    Width = 604
    Height = 50
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 1
    object Label9: TLabel
      Left = 24
      Top = 5
      Width = 22
      Height = 13
      Caption = 'DPA'
    end
    object Label10: TLabel
      Left = 24
      Top = 29
      Width = 23
      Height = 13
      Caption = 'DPR'
    end
    object Label11: TLabel
      Left = 176
      Top = 3
      Width = 29
      Height = 13
      Caption = 'PVHT'
    end
    object Label12: TLabel
      Left = 176
      Top = 29
      Width = 35
      Height = 13
      Caption = 'PVTTC'
    end
    object Label13: TLabel
      Left = 320
      Top = 5
      Width = 29
      Height = 13
      Caption = 'PAHT'
    end
    object Label14: TLabel
      Left = 320
      Top = 29
      Width = 30
      Height = 13
      Caption = 'PRHT'
    end
    object Label15: TLabel
      Left = 464
      Top = 5
      Width = 30
      Height = 13
      Caption = 'PMAP'
    end
    object Label16: TLabel
      Left = 464
      Top = 29
      Width = 31
      Height = 13
      Caption = 'PMRP'
    end
    object GNL_DPA: THCritMaskEdit
      Left = 76
      Top = 1
      Width = 85
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 0
      TagDispatch = 0
      OpeType = otReel
    end
    object GNL_DPR: THCritMaskEdit
      Left = 76
      Top = 25
      Width = 85
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 1
      TagDispatch = 0
      OpeType = otReel
    end
    object GNL_PVHT: THCritMaskEdit
      Left = 216
      Top = 1
      Width = 85
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 2
      TagDispatch = 0
      OpeType = otReel
    end
    object GNL_PVTTC: THCritMaskEdit
      Left = 216
      Top = 25
      Width = 85
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 3
      TagDispatch = 0
      OpeType = otReel
    end
    object GNL_PAHT: THCritMaskEdit
      Left = 360
      Top = 1
      Width = 85
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 4
      TagDispatch = 0
      OpeType = otReel
    end
    object GNL_PRHT: THCritMaskEdit
      Left = 360
      Top = 25
      Width = 85
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 5
      TagDispatch = 0
      OpeType = otReel
    end
    object GNL_PMAP: THCritMaskEdit
      Left = 504
      Top = 1
      Width = 85
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 6
      TagDispatch = 0
      OpeType = otReel
    end
    object GNL_PMRP: THCritMaskEdit
      Left = 504
      Top = 25
      Width = 85
      Height = 21
      TabStop = False
      Color = clBtnFace
      ReadOnly = True
      TabOrder = 7
      TagDispatch = 0
      OpeType = otReel
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 365
    Width = 604
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object ToolWindow971: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 604
      Caption = 'ToolWindow971'
      ClientAreaHeight = 32
      ClientAreaWidth = 604
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BAide: TToolbarButton97
        Tag = 1
        Left = 569
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
        Left = 539
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
        Left = 509
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
  object PANDIM: TPanel
    Left = 0
    Top = 289
    Width = 604
    Height = 26
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 3
    object TGA_CODEDIM1: TLabel
      Left = 24
      Top = 7
      Width = 38
      Height = 13
      Caption = 'Label17'
    end
    object TGA_CODEDIM2: TLabel
      Left = 112
      Top = 7
      Width = 38
      Height = 13
      Caption = 'Label17'
    end
    object TGA_CODEDIM3: TLabel
      Left = 184
      Top = 7
      Width = 38
      Height = 13
      Caption = 'Label17'
    end
    object TGA_CODEDIM4: TLabel
      Left = 252
      Top = 7
      Width = 38
      Height = 13
      Caption = 'Label17'
    end
    object TGA_CODEDIM5: TLabel
      Left = 324
      Top = 7
      Width = 38
      Height = 13
      Caption = 'Label17'
    end
  end
  object G_NLIG: THGrid
    Tag = 1
    Left = 0
    Top = 105
    Width = 389
    Height = 184
    TabStop = False
    Align = alClient
    ColCount = 4
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 50
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goRowSelect]
    PopupMenu = PopupG_NLIG
    TabOrder = 4
    OnDblClick = G_NLIGDblClick
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
      64
      64)
  end
  object TV_NLIG: TTreeView
    Left = 389
    Top = 105
    Width = 215
    Height = 184
    Align = alRight
    HideSelection = False
    Indent = 19
    TabOrder = 5
    TabStop = False
    OnMouseUp = TV_NLIGMouseUp
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 532
    Top = 124
  end
  object PopupG_NLIG: TPopupMenu
    Left = 504
    Top = 124
    object PopG_NLIG_A: TMenuItem
      Caption = '&Fiche Article'
      ShortCut = 16449
      OnClick = PopG_NLIG_AClick
    end
    object PopG_NLIG_C: TMenuItem
      Caption = '&Cas d'#39'emploi'
      ShortCut = 16451
      OnClick = PopG_NLIG_CClick
    end
    object N2: TMenuItem
      Caption = '-'
    end
    object PopG_NLIG_D: TMenuItem
      Caption = 'Voir &Dimensions'
      Checked = True
      ShortCut = 16470
      OnClick = PopG_NLIG_DClick
    end
    object PopG_NLIG_P: TMenuItem
      Caption = 'Voir &Prix'
      Checked = True
      ShortCut = 16464
      OnClick = PopG_NLIG_PClick
    end
  end
end

object FMulLettre: TFMulLettre
  Left = 278
  Top = 247
  Width = 620
  Height = 390
  HelpContext = 7508000
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Crit'#232'res du lettrage'
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
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object iCritGlyphModified: TImage
    Left = 128
    Top = 192
    Width = 32
    Height = 16
    AutoSize = True
    Picture.Data = {
      07544269746D617076010000424D760100000000000076000000280000002000
      000010000000010004000000000000010000120B0000120B0000100000000000
      0000000000000000800000800000008080008000000080008000808000008080
      8000C0C0C0000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
      FF00555555555555555555555555555555555555555555555555555555FF5555
      555555555C05555555555555577FF55555555555CCC05555555555557777F555
      55555555CCC05555555555557777FF555555555CCCCC05555555555777777F55
      555555CCCCCC05555555557777777FF5555554CC05CCC05555555777757777F5
      55554C05555CC05555557775555777FF55555555555CCC05555555555557777F
      555555500055CC05555555555555777FF555555060555CC05555555555555777
      FF5555505055554C05555555555555777FF5550556055554C055555555555557
      77FF506F556055555CC055555555555557770000000005555555555555555555
      5555}
    Visible = False
  end
  object iCritGlyph: TImage
    Left = 96
    Top = 208
    Width = 32
    Height = 16
    AutoSize = True
    Picture.Data = {
      07544269746D617076010000424D760100000000000076000000280000002000
      000010000000010004000000000000010000120B0000120B0000100000000000
      000000000000000080000080000000808000800000008000800080800000C0C0
      C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
      FF00555555555555555555555555555555555555555555555555555555FF5555
      555555555C05555555555555588FF55555555555CCC05555555555558888F555
      55555555CCC05555555555558888FF555555555CCCCC05555555555888888F55
      555555CCCCCC05555555558888888FF5555554CC05CCC05555555888858888F5
      55554C05555CC05555558885555888FF55555555555CCC05555555555558888F
      555555555555CC05555555555555888FF555555555555CC05555555555555888
      FF5555555555554C05555555555555888FF5555555555554C055555555555558
      88FF5555555555555CC055555555555558885555555555555555555555555555
      5555}
    Visible = False
  end
  object GL: THGrid
    Left = 0
    Top = 147
    Width = 612
    Height = 180
    Align = alClient
    ColCount = 2
    DefaultRowHeight = 18
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goRowSelect]
    TabOrder = 2
    OnDblClick = GLDblClick
    OnKeyDown = FormKeyDown
    SortedCol = -1
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = 13224395
    DBIndicator = True
    ColWidths = (
      10
      64)
  end
  object Dock: TDock97
    Left = 0
    Top = 327
    Width = 612
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 612
      Caption = 'Actions'
      ClientAreaHeight = 32
      ClientAreaWidth = 612
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BReduire: TToolbarButton97
        Left = 4
        Top = 3
        Width = 28
        Height = 27
        Hint = 'R'#233'duire la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'R'#233'duire'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BReduireClick
        GlobalIndexImage = 'Z0910_S16G1'
      end
      object BAgrandir: TToolbarButton97
        Left = 4
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Agrandir la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'Agrandir'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAgrandirClick
        GlobalIndexImage = 'Z0270_S16G1'
      end
      object BRecherche: TToolbarButton97
        Left = 36
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Rechercher dans la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'Chercher'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BRechercheClick
        GlobalIndexImage = 'Z0077_S16G1'
      end
      object BImprimer: TToolbarButton97
        Left = 482
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
        Visible = False
        OnClick = BImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BValider: TToolbarButton97
        Left = 514
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Lancer lettrage'
        Default = True
        DisplayMode = dmGlyphOnly
        Caption = 'Valider'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
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
        Margin = 3
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 546
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
        Left = 578
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
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 612
    Height = 112
    ActivePage = Princ
    Align = alTop
    TabOrder = 0
    object Princ: TTabSheet
      Caption = 'Standards'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 604
        Height = 84
        Align = alClient
      end
      object HLabel4: THLabel
        Left = 9
        Top = 11
        Width = 61
        Height = 13
        Caption = '&G'#233'n'#233'raux de'
        FocusControl = E_GENERAL
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel5: THLabel
        Left = 179
        Top = 11
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_GENERAL_
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel2: THLabel
        Left = 179
        Top = 34
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_AUXILIAIRE_
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel1: THLabel
        Left = 9
        Top = 36
        Width = 61
        Height = 13
        Caption = '&Auxiliaires de'
        FocusControl = E_AUXILIAIRE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object Label5: THLabel
        Left = 341
        Top = 36
        Width = 57
        Height = 13
        Caption = '&Intitul'#233' Tiers'
        FocusControl = T_LIBELLE
      end
      object HLabel8: THLabel
        Left = 9
        Top = 61
        Width = 33
        Height = 13
        Caption = '&Devise'
        FocusControl = E_DEVISE
      end
      object TE_ETABLISSEMENT: THLabel
        Left = 341
        Top = 11
        Width = 65
        Height = 13
        Caption = 'Etabli&ssement'
        FocusControl = E_ETABLISSEMENT
      end
      object BFeuVert: THBitBtn
        Left = 283
        Top = 7
        Width = 27
        Height = 46
        Hint = 'Lancer lettrage'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 8
        Visible = False
        OnClick = BFeuVertClick
        Layout = blGlyphRight
        Spacing = -1
        GlobalIndexImage = 'Z0101_S24G1'
      end
      object E_GENERAL: THCritMaskEdit
        Tag = 1
        Left = 84
        Top = 7
        Width = 89
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 0
        OnExit = E_GENERALExit
        TagDispatch = 0
        DataType = 'TZGENERAL'
        Operateur = Superieur
        ElipsisButton = True
        OnElipsisClick = E_GENERALElipsisClick
      end
      object E_GENERAL_: THCritMaskEdit
        Tag = 1
        Left = 192
        Top = 7
        Width = 89
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 1
        OnExit = E_GENERALExit
        TagDispatch = 0
        DataType = 'TZGENERAL'
        Operateur = Inferieur
        ElipsisButton = True
        OnElipsisClick = E_GENERALElipsisClick
      end
      object E_AUXILIAIRE_: THCritMaskEdit
        Tag = 1
        Left = 192
        Top = 32
        Width = 89
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 3
        OnExit = E_AUXILIAIREExit
        TagDispatch = 0
        DataType = 'TZTTOUS'
        Operateur = Inferieur
        ElipsisButton = True
        OnElipsisClick = E_AUXILIAIRE2ElipsisClick
      end
      object T_LIBELLE: TEdit
        Left = 418
        Top = 32
        Width = 165
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 6
      end
      object E_DEVISE: THValComboBox
        Tag = 1
        Left = 84
        Top = 57
        Width = 198
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        Vide = True
        DataType = 'TTDEVISE'
      end
      object Distinguer: TCheckBox
        Left = 341
        Top = 59
        Width = 244
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Distinguer les mouvements'
        TabOrder = 7
      end
      object E_ETABLISSEMENT: THValComboBox
        Left = 418
        Top = 7
        Width = 165
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 5
        OnChange = E_EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTETABLISSEMENT'
      end
      object E_AUXILIAIRE: THCritMaskEdit
        Tag = 1
        Left = 84
        Top = 32
        Width = 89
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 2
        OnExit = E_AUXILIAIREExit
        TagDispatch = 0
        DataType = 'TZTTOUS'
        Operateur = Superieur
        ElipsisButton = True
        OnElipsisClick = E_AUXILIAIRE2ElipsisClick
      end
    end
    object Complements: TTabSheet
      Caption = 'Compl'#233'ments'
      object Bevel2: TBevel
        Left = 0
        Top = 0
        Width = 604
        Height = 84
        Align = alClient
      end
      object Label12: THLabel
        Left = 8
        Top = 12
        Width = 55
        Height = 13
        Caption = 'Ref. &interne'
        FocusControl = E_REFINTERNE
      end
      object HLabel10: THLabel
        Left = 291
        Top = 12
        Width = 69
        Height = 13
        Caption = '&Ech'#233'ances du'
        FocusControl = E_DATEECHEANCE
      end
      object HLabel7: THLabel
        Left = 489
        Top = 12
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATEECHEANCE_
      end
      object TE_EXERCICE: THLabel
        Left = 9
        Top = 36
        Width = 41
        Height = 13
        Caption = 'E&xercice'
        FocusControl = E_EXERCICE
      end
      object HLabel3: THLabel
        Left = 291
        Top = 36
        Width = 100
        Height = 13
        Caption = '&Dates comptables du'
        FocusControl = E_DATECOMPTABLE
      end
      object HLabel6: THLabel
        Left = 489
        Top = 36
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATECOMPTABLE_
      end
      object Label14: THLabel
        Left = 9
        Top = 60
        Width = 63
        Height = 13
        Caption = '&Natures Tiers'
        FocusControl = T_NATUREAUXI
      end
      object E_REFINTERNE: THCritMaskEdit
        Left = 84
        Top = 8
        Width = 177
        Height = 21
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Operateur = Superieur
      end
      object LETTREPARTIEL: TCheckBox
        Left = 290
        Top = 59
        Width = 130
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Ajouter &lettrage partiel'
        Checked = True
        State = cbChecked
        TabOrder = 6
      end
      object E_DATEECHEANCE: THCritMaskEdit
        Left = 406
        Top = 8
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!00/00/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 1
        Text = '  /  /    '
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATEECHEANCE_: THCritMaskEdit
        Left = 516
        Top = 8
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!00/00/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 2
        Text = '  /  /    '
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_EXERCICE: THValComboBox
        Left = 84
        Top = 32
        Width = 177
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 3
        OnChange = E_EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 406
        Top = 32
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!00/00/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 4
        Text = '  /  /    '
        OnKeyPress = E_DATECOMPTABLEKeyPress
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATECOMPTABLE_: THCritMaskEdit
        Left = 516
        Top = 32
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!00/00/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 5
        Text = '  /  /    '
        OnKeyPress = E_DATECOMPTABLEKeyPress
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object CFacNotPay: TCheckBox
        Left = 441
        Top = 59
        Width = 144
        Height = 17
        Alignment = taLeftJustify
        AllowGrayed = True
        Caption = 'Tiers sans payeur associ'#233
        Checked = True
        State = cbChecked
        TabOrder = 7
        OnClick = CFacNotPayClick
      end
      object XX_WHERETP: TEdit
        Left = 265
        Top = 8
        Width = 15
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 8
        Text = 'E_TIERSPAYEUR=""'
        Visible = False
      end
      object T_NATUREAUXI: THValComboBox
        Left = 84
        Top = 57
        Width = 177
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 9
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATTIERSCPTA'
      end
    end
    object Avances: TTabSheet
      Caption = 'Avanc'#233's'
      object Bevel3: TBevel
        Left = 0
        Top = 0
        Width = 604
        Height = 84
        Align = alClient
      end
      object Label7: THLabel
        Left = 12
        Top = 12
        Width = 58
        Height = 13
        Caption = '&Ref. lettrage'
        FocusControl = E_REFLETTRAGE
      end
      object Label3: THLabel
        Left = 352
        Top = 36
        Width = 70
        Height = 13
        Caption = '&Montants pivot'
        FocusControl = TOTAL
      end
      object Label4: THLabel
        Left = 508
        Top = 36
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = TOTAL_
      end
      object HLabel11: THLabel
        Left = 508
        Top = 12
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = LETTRAGE_
      end
      object HLabel9: THLabel
        Left = 352
        Top = 12
        Width = 70
        Height = 13
        Caption = '&Code Lettre de'
        FocusControl = LETTRAGE
      end
      object Label1: THLabel
        Left = 352
        Top = 60
        Width = 57
        Height = 13
        Caption = '&Num'#233'ros de'
        FocusControl = E_NUMEROPIECE
      end
      object Label2: THLabel
        Left = 508
        Top = 60
        Width = 6
        Height = 13
        Caption = #224
      end
      object Label6: THLabel
        Left = 12
        Top = 36
        Width = 47
        Height = 13
        Caption = '&Code tiers'
        FocusControl = T_TIERS
      end
      object E_REFLETTRAGE: THCritMaskEdit
        Left = 83
        Top = 8
        Width = 152
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
      end
      object TOTAL: THCritMaskEdit
        Tag = 3
        Left = 436
        Top = 32
        Width = 65
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 4
        Text = '-9999999999'
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object TOTAL_: THCritMaskEdit
        Tag = 3
        Left = 524
        Top = 32
        Width = 65
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 5
        Text = '9999999999'
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object LETTRAGE_: THCritMaskEdit
        Left = 524
        Top = 8
        Width = 65
        Height = 21
        Ctl3D = True
        MaxLength = 4
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        Operateur = Inferieur
      end
      object LETTRAGE: THCritMaskEdit
        Left = 436
        Top = 8
        Width = 65
        Height = 21
        Ctl3D = True
        MaxLength = 4
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        Operateur = Superieur
      end
      object E_NUMEROPIECE: THCritMaskEdit
        Tag = 2
        Left = 436
        Top = 56
        Width = 65
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 6
        Text = '0'
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_NUMEROPIECE_: THCritMaskEdit
        Tag = 2
        Left = 524
        Top = 56
        Width = 65
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 7
        Text = '99999999'
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object T_TIERS: TEdit
        Left = 83
        Top = 32
        Width = 152
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 1
      end
    end
    object PRefs: TTabSheet
      Caption = 'R'#233'f'#233'rences'
      ImageIndex = 4
      object Bevel5: TBevel
        Left = 0
        Top = 0
        Width = 604
        Height = 84
        Align = alClient
      end
      object TE_REFRELEVE: THLabel
        Left = 20
        Top = 10
        Width = 115
        Height = 13
        Caption = 'Ref'#233'rence de  rele&v'#233' de'
        FocusControl = E_REFRELEVE
      end
      object TE_REFRELEVE_: THLabel
        Left = 332
        Top = 10
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_REFRELEVE_
      end
      object TE_REFLIBRE: THLabel
        Left = 20
        Top = 36
        Width = 87
        Height = 13
        Caption = 'Ref'#233'rence &libre de'
        FocusControl = E_REFLIBRE
      end
      object TE_REFLIBRE_: THLabel
        Left = 332
        Top = 36
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_REFLIBRE_
      end
      object TE_REFEXTERNE: THLabel
        Left = 20
        Top = 62
        Width = 103
        Height = 13
        Caption = 'Ref'#233'rence e&xterne de'
        FocusControl = E_REFEXTERNE
      end
      object TE_REFEXTERNE_: THLabel
        Left = 332
        Top = 62
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_REFEXTERNE_
      end
      object E_REFRELEVE: THCritMaskEdit
        Left = 156
        Top = 6
        Width = 152
        Height = 21
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Operateur = Superieur
      end
      object E_REFRELEVE_: THCritMaskEdit
        Left = 356
        Top = 6
        Width = 152
        Height = 21
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 1
        TagDispatch = 0
        Operateur = Inferieur
      end
      object E_REFLIBRE: THCritMaskEdit
        Left = 156
        Top = 32
        Width = 152
        Height = 21
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        Operateur = Superieur
      end
      object E_REFLIBRE_: THCritMaskEdit
        Left = 356
        Top = 32
        Width = 152
        Height = 21
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        Operateur = Inferieur
      end
      object E_REFEXTERNE: THCritMaskEdit
        Left = 156
        Top = 58
        Width = 152
        Height = 21
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        Operateur = Superieur
      end
      object E_REFEXTERNE_: THCritMaskEdit
        Left = 356
        Top = 58
        Width = 152
        Height = 21
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 5
        TagDispatch = 0
        Operateur = Inferieur
      end
    end
    object Plibres: TTabSheet
      Caption = 'Tables libres'
      object Bevel4: TBevel
        Left = 0
        Top = 0
        Width = 604
        Height = 84
        Align = alClient
      end
      object TT_TABLE0: THLabel
        Left = 11
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE0'
      end
      object TT_TABLE1: THLabel
        Left = 130
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE1'
      end
      object TT_TABLE2: THLabel
        Left = 249
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE2'
      end
      object TT_TABLE3: THLabel
        Left = 367
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE3'
      end
      object TT_TABLE4: THLabel
        Left = 483
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE4'
      end
      object TT_TABLE5: THLabel
        Left = 11
        Top = 42
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE5'
      end
      object TT_TABLE6: THLabel
        Left = 130
        Top = 42
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE6'
      end
      object TT_TABLE7: THLabel
        Left = 249
        Top = 42
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE7'
      end
      object TT_TABLE8: THLabel
        Left = 367
        Top = 42
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE8'
      end
      object TT_TABLE9: THLabel
        Left = 483
        Top = 42
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE9'
      end
      object T_TABLE4: THCritMaskEdit
        Left = 483
        Top = 17
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 4
        TagDispatch = 0
        DataType = 'TZNATTIERS4'
        ElipsisButton = True
      end
      object T_TABLE3: THCritMaskEdit
        Left = 367
        Top = 17
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 3
        TagDispatch = 0
        DataType = 'TZNATTIERS3'
        ElipsisButton = True
      end
      object T_TABLE2: THCritMaskEdit
        Left = 249
        Top = 17
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 2
        TagDispatch = 0
        DataType = 'TZNATTIERS2'
        ElipsisButton = True
      end
      object T_TABLE1: THCritMaskEdit
        Left = 130
        Top = 17
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 1
        TagDispatch = 0
        DataType = 'TZNATTIERS1'
        ElipsisButton = True
      end
      object T_TABLE0: THCritMaskEdit
        Left = 11
        Top = 17
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 0
        TagDispatch = 0
        DataType = 'TZNATTIERS0'
        ElipsisButton = True
      end
      object T_TABLE5: THCritMaskEdit
        Left = 11
        Top = 57
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 5
        TagDispatch = 0
        DataType = 'TZNATTIERS5'
        ElipsisButton = True
      end
      object T_TABLE6: THCritMaskEdit
        Left = 130
        Top = 57
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 6
        TagDispatch = 0
        DataType = 'TZNATTIERS6'
        ElipsisButton = True
      end
      object T_TABLE7: THCritMaskEdit
        Left = 249
        Top = 57
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 7
        TagDispatch = 0
        DataType = 'TZNATTIERS7'
        ElipsisButton = True
      end
      object T_TABLE8: THCritMaskEdit
        Left = 367
        Top = 57
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 8
        TagDispatch = 0
        DataType = 'TZNATTIERS8'
        ElipsisButton = True
      end
      object T_TABLE9: THCritMaskEdit
        Left = 483
        Top = 57
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 9
        TagDispatch = 0
        DataType = 'TZNATTIERS9'
        ElipsisButton = True
      end
    end
  end
  object Cache: THCpteEdit
    Left = 16
    Top = 178
    Width = 17
    Height = 21
    TabStop = False
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
    Text = '!!!'
    Visible = False
    ZoomTable = tzGLettColl
    Vide = False
    Bourre = True
    okLocate = False
    SynJoker = False
  end
  object Dock971: TDock97
    Left = 0
    Top = 112
    Width = 612
    Height = 35
    AllowDrag = False
    object PFiltres: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 612
      Caption = ' '
      ClientAreaHeight = 31
      ClientAreaWidth = 612
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BFiltre: TToolbarButton97
        Left = 5
        Top = 5
        Width = 56
        Height = 21
        Hint = 'Menu filtre'
        DropdownArrow = True
        DropdownMenu = POPF
        Caption = '&Filtres'
        Layout = blGlyphRight
        ParentShowHint = False
        ShowHint = True
      end
      object BChercher: TToolbarButton97
        Left = 577
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Appliquer crit'#232'res'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BChercherClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object FFiltres: THValComboBox
        Left = 67
        Top = 5
        Width = 503
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = FFiltresChange
        TagDispatch = 0
      end
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 226
    Top = 276
  end
  object FindLettre: TFindDialog
    OnFind = FindLettreFind
    Left = 117
    Top = 276
  end
  object HMessMulL: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Lettrage manuel;Les mouvements sont en devise. Voulez-vous let' +
        'trer en ;Q;YNC;Y;Y;'
      
        '1;Lettrage manuel;Le compte g'#233'n'#233'ral que vous avez renseign'#233' est ' +
        'incorrect ou interdit;W;O;O;O;'
      
        '2;Lettrage manuel;Le compte auxiliaire que vous avez renseign'#233' e' +
        'st incorrect ou interdit;W;O;O;O;'
      
        '3;Lettrage manuel;Le compte g'#233'n'#233'ral que vous avez renseign'#233' n'#39'es' +
        't pas collectif;W;O;O;O;'
      
        '4;Lettrage manuel;La s'#233'lection des dates est incoh'#233'rente.;W;O;O;' +
        'O;'
      
        '5;Lettrage manuel;Voulez-vous lettrer vos mouvements en ;Q;YNC;Y' +
        ';Y;'
      'Distinguer les mouvements '
      '7;')
    Left = 171
    Top = 276
  end
  object POPF: TPopupMenu
    OnPopup = POPFPopup
    Left = 272
    Top = 276
    object BCREER: TMenuItem
      Caption = '&Cr'#233'er un filtre'
      OnClick = BCREERClick
    end
    object BENREGISTRE: TMenuItem
      Caption = '&Enregistrer le filtre'
      OnClick = BENREGISTREClick
    end
    object BDUPLIQUER: TMenuItem
      Caption = '&Duppliqsuer le filtre'
      OnClick = BDUPLIQUERClick
    end
    object BSUPPRIMER: TMenuItem
      Caption = '&Supprimer le filtre'
      OnClick = BSUPPRIMERClick
    end
    object BRENOMMER: TMenuItem
      Caption = '&Renommer le filtre'
      OnClick = BRENOMMERClick
    end
    object BRECHERCHER: TMenuItem
      Caption = '&Nouvelle recherche'
      OnClick = BRECHERCHERClick
    end
  end
end

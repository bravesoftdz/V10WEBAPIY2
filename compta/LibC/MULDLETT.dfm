object FMulDLettre: TFMulDLettre
  Left = 275
  Top = 153
  Width = 622
  Height = 420
  HelpContext = 7511000
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Crit'#232'res du d'#233'lettrage'
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
  object Dock: TDock97
    Left = 0
    Top = 357
    Width = 614
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 614
      Caption = 'Actions'
      ClientAreaHeight = 32
      ClientAreaWidth = 614
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BReduire: TToolbarButton97
        Left = 2
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
        Left = 2
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
        Left = 34
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
        Left = 452
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
        Left = 516
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Lancer le d'#233'lettrage'
        Default = True
        DisplayMode = dmGlyphOnly
        Caption = 'D'#233'lettrer'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 3
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0184_S16G1'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 548
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
        Left = 580
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
      object BExec: TToolbarButton97
        Left = 484
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Tout s'#233'lectionner et supprimer le lettrage'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BExecClick
        GlobalIndexImage = 'Z0050_S16G1'
      end
    end
  end
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 614
    Height = 113
    ActivePage = Princ
    Align = alTop
    TabOrder = 0
    object Princ: TTabSheet
      Caption = 'Crit'#232'res standards'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 606
        Height = 85
        Align = alClient
      end
      object HLabel4: THLabel
        Left = 12
        Top = 9
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
        Left = 188
        Top = 9
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
        Left = 188
        Top = 35
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
        Left = 12
        Top = 35
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
      object HLabel9: THLabel
        Left = 360
        Top = 9
        Width = 70
        Height = 13
        Caption = 'Code &Lettre de'
        FocusControl = LETTRAGE
      end
      object HLabel8: THLabel
        Left = 360
        Top = 35
        Width = 33
        Height = 13
        Caption = '&Devise'
        FocusControl = E_DEVISE
      end
      object HLabel11: THLabel
        Left = 510
        Top = 9
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = LETTRAGE_
      end
      object H_REFLETTRAGE: THLabel
        Left = 360
        Top = 61
        Width = 58
        Height = 13
        Caption = '&Ref. lettrage'
        Color = clYellow
        FocusControl = E_REFLETTRAGE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        Visible = False
      end
      object BFeuVert: THBitBtn
        Left = 300
        Top = 5
        Width = 27
        Height = 48
        TabOrder = 4
        Visible = False
        OnClick = BFeuVertClick
        Layout = blGlyphRight
        Spacing = -1
        GlobalIndexImage = 'Z0101_S24G1'
      end
      object E_GENERAL: THCritMaskEdit
        Tag = 1
        Left = 88
        Top = 5
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 0
        OnExit = E_GENERALExit
        TagDispatch = 0
        Operateur = Superieur
        ElipsisButton = True
        OnElipsisClick = E_GENERALElipsisClick
      end
      object E_GENERAL_: THCritMaskEdit
        Tag = 1
        Left = 204
        Top = 5
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 1
        OnExit = E_GENERALExit
        TagDispatch = 0
        Operateur = Inferieur
        ElipsisButton = True
        OnElipsisClick = E_GENERALElipsisClick
      end
      object E_AUXILIAIRE_: THCritMaskEdit
        Tag = 1
        Left = 204
        Top = 31
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 3
        OnExit = E_AUXILIAIREExit
        TagDispatch = 0
        Operateur = Inferieur
        ElipsisButton = True
        OnElipsisClick = E_AUXILIAIREElipsisClick
      end
      object E_AUXILIAIRE: THCritMaskEdit
        Tag = 1
        Left = 88
        Top = 31
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 2
        OnExit = E_AUXILIAIREExit
        TagDispatch = 0
        Operateur = Superieur
        ElipsisButton = True
        OnElipsisClick = E_AUXILIAIREElipsisClick
      end
      object E_DEVISE: THValComboBox
        Tag = 1
        Left = 436
        Top = 31
        Width = 159
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 5
        TagDispatch = 0
        Vide = True
        DataType = 'TTDEVISE'
      end
      object LETTRAGE: THCritMaskEdit
        Left = 436
        Top = 5
        Width = 69
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 4
        ParentCtl3D = False
        TabOrder = 6
        TagDispatch = 0
        Operateur = Superieur
      end
      object LETTRAGE_: THCritMaskEdit
        Left = 524
        Top = 5
        Width = 69
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 4
        ParentCtl3D = False
        TabOrder = 7
        TagDispatch = 0
        Operateur = Inferieur
      end
      object CFacNotPay: TCheckBox
        Left = 9
        Top = 59
        Width = 287
        Height = 17
        Alignment = taLeftJustify
        AllowGrayed = True
        Caption = 'Voir seulement les mouvements sans tiers pa&yeur'
        Checked = True
        State = cbChecked
        TabOrder = 8
      end
      object E_REFLETTRAGE: THCritMaskEdit
        Tag = 1
        Left = 436
        Top = 57
        Width = 159
        Height = 21
        CharCase = ecUpperCase
        Color = clYellow
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 9
        Visible = False
        TagDispatch = 0
      end
    end
    object PLibres: TTabSheet
      Caption = 'Tables libres'
      object Bevel2: TBevel
        Left = 0
        Top = 0
        Width = 606
        Height = 85
        Align = alClient
      end
      object TT_TABLE0: THLabel
        Left = 13
        Top = 1
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE0'
      end
      object TT_TABLE1: THLabel
        Left = 130
        Top = 1
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE1'
      end
      object TT_TABLE2: THLabel
        Left = 247
        Top = 1
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE2'
      end
      object TT_TABLE3: THLabel
        Left = 363
        Top = 1
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE3'
      end
      object TT_TABLE4: THLabel
        Left = 477
        Top = 1
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE4'
      end
      object TT_TABLE5: THLabel
        Left = 13
        Top = 41
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE5'
      end
      object TT_TABLE6: THLabel
        Left = 130
        Top = 41
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE6'
      end
      object TT_TABLE7: THLabel
        Left = 247
        Top = 41
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE7'
      end
      object TT_TABLE8: THLabel
        Left = 363
        Top = 41
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE8'
      end
      object TT_TABLE9: THLabel
        Left = 477
        Top = 41
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE9'
      end
      object T_TABLE4: THCritMaskEdit
        Left = 477
        Top = 16
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 0
        TagDispatch = 0
        DataType = 'TZNATTIERS4'
        ElipsisButton = True
      end
      object T_TABLE3: THCritMaskEdit
        Left = 363
        Top = 16
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 1
        TagDispatch = 0
        DataType = 'TZNATTIERS3'
        ElipsisButton = True
      end
      object T_TABLE2: THCritMaskEdit
        Left = 247
        Top = 16
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
        Top = 16
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 3
        TagDispatch = 0
        DataType = 'TZNATTIERS1'
        ElipsisButton = True
      end
      object T_TABLE0: THCritMaskEdit
        Left = 13
        Top = 16
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 4
        TagDispatch = 0
        DataType = 'TZNATTIERS0'
        ElipsisButton = True
      end
      object T_TABLE5: THCritMaskEdit
        Left = 13
        Top = 56
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
        Top = 56
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 6
        TagDispatch = 0
        DataType = 'TZNATTIERS6'
        ElipsisButton = True
      end
      object T_TABLE7: THCritMaskEdit
        Left = 247
        Top = 56
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 7
        TagDispatch = 0
        DataType = 'TZNATTIERS7'
        ElipsisButton = True
      end
      object T_TABLE8: THCritMaskEdit
        Left = 363
        Top = 56
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 8
        TagDispatch = 0
        DataType = 'TZNATTIERS8'
        ElipsisButton = True
      end
      object T_TABLE9: THCritMaskEdit
        Left = 477
        Top = 56
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
    Left = 8
    Top = 138
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
    TabOrder = 2
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
    Top = 113
    Width = 614
    Height = 39
    AllowDrag = False
    object PFiltres: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 35
      ClientWidth = 614
      ClientAreaHeight = 35
      ClientAreaWidth = 614
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BChercher: TToolbarButton97
        Left = 579
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
      object BFiltre: TToolbarButton97
        Left = 5
        Top = 5
        Width = 56
        Height = 21
        Hint = 'Menu filtre'
        DropdownArrow = True
        Caption = '&Filtres'
        Layout = blGlyphRight
        ParentShowHint = False
        ShowHint = True
      end
      object FFiltres: THValComboBox
        Left = 67
        Top = 5
        Width = 505
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
      end
    end
  end
  object GL: THGrid
    Left = 0
    Top = 152
    Width = 614
    Height = 205
    Align = alClient
    ColCount = 2
    DefaultRowHeight = 18
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSizing, goRowSelect]
    TabOrder = 4
    OnDblClick = GLDblClick
    SortedCol = -1
    Couleur = False
    MultiSelect = True
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
  object FindLettre: TFindDialog
    OnFind = FindLettreFind
    Left = 136
    Top = 276
  end
  object HMessMulDL: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;LIBRE LIBRE LIBRE'
      
        '1;D'#233'lettrage;Le compte g'#233'n'#233'ral que vous avez renseign'#233' est incor' +
        'rect ou interdit;E;O;O;O;'
      
        '2;D'#233'lettrage;Le compte auxiliaire que vous avez renseign'#233' est in' +
        'correct ou interdit;E;O;O;O;'
      
        '3;D'#233'lettrage;Le compte g'#233'n'#233'ral que vous avez renseign'#233' n'#39'est pas' +
        ' collectif;E;O;O;O;'
      '4;D'#233'lettrage;Aucun lettrage n'#39'est s'#233'lectionn'#233';E;O;O;O;')
    Left = 196
    Top = 276
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 316
    Top = 276
  end
  object POPF: TPopupMenu
    Left = 264
    Top = 276
    object BCreerFiltre: TMenuItem
      Caption = '&Cr'#233'er un filtre'
    end
    object BSaveFiltre: TMenuItem
      Caption = '&Enregistrer le filtre'
    end
    object BDelFiltre: TMenuItem
      Caption = '&Supprimer le filtre'
    end
    object BRenFiltre: TMenuItem
      Caption = '&Renommer le filtre'
    end
    object BNouvRech: TMenuItem
      Caption = '&Nouvelle recherche'
    end
  end
end

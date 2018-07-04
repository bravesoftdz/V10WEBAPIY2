object FSaisieFolio: TFSaisieFolio
  Left = 233
  Top = 203
  Width = 740
  Height = 372
  HelpContext = 7242000
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Saisie des '#233'critures'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clBlack
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  PopupMenu = POPS
  Position = poScreenCenter
  ShowHint = True
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object DockTop: TDock97
    Left = 0
    Top = 0
    Width = 724
    Height = 9
  end
  object DockRight: TDock97
    Left = 715
    Top = 9
    Width = 9
    Height = 293
    Position = dpRight
  end
  object DockLeft: TDock97
    Left = 0
    Top = 9
    Width = 9
    Height = 293
    Position = dpLeft
  end
  object Dock: TDock97
    Left = 0
    Top = 302
    Width = 724
    Height = 31
    AllowDrag = False
    Position = dpBottom
    object Valide97: TToolbar97
      Left = 516
      Top = 0
      Caption = 'Validation'
      CloseButton = False
      DockPos = 516
      TabOrder = 0
      object BValide: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Valider la saisie'
        Caption = '7'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValideClick
        GlobalIndexImage = 'Z0184_S16G1'
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 28
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Caption = 'Fermer'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAbandonClick
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 56
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Aide'
        Caption = 'Aide'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
    end
    object Outils: TToolbar97
      Left = 0
      Top = 0
      Caption = 'Outils'
      CloseButton = False
      DockPos = 0
      TabOrder = 1
      object BSolde: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Calcul du solde'
        Caption = 'Solde'
        DisplayMode = dmGlyphOnly
        Layout = blGlyphTop
        NumGlyphs = 2
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BSoldeClick
        GlobalIndexImage = 'Z0051_S16G2'
      end
      object BComplement: TToolbarButton97
        Tag = 1
        Left = 146
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Informations compl'#233'mentaires'
        Caption = 'Compl'#233'ments'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 0
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BComplementClick
        GlobalIndexImage = 'Z0105_S16G2'
        IsControl = True
      end
      object BChercher: TToolbarButton97
        Tag = 1
        Left = 202
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Rechercher dans la pi'#232'ce'
        Caption = 'Chercher'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BChercherClick
        GlobalIndexImage = 'Z0077_S16G1'
        IsControl = True
      end
      object BMenuZoom: TToolbarButton97
        Tag = -100
        Left = 236
        Top = 0
        Width = 37
        Height = 27
        Hint = 'Menu zoom'
        Caption = 'Zooms'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = POPZ
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnMouseEnter = BMenuZoomMouseEnter
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
      object ToolbarSep971: TToolbarSep97
        Left = 56
        Top = 0
      end
      object ToolbarSep972: TToolbarSep97
        Left = 230
        Top = 0
      end
      object Sep97: TToolbarSep97
        Left = 273
        Top = 0
      end
      object BVentil: TToolbarButton97
        Tag = 1
        Left = 62
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Ventilation analytique'
        Caption = 'Analytiques'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BVentilClick
        GlobalIndexImage = 'Z0133_S16G1'
        IsControl = True
      end
      object BEche: TToolbarButton97
        Tag = 1
        Left = 28
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Ech'#233'ances'
        Caption = 'Ech'#233'ance'
        DisplayMode = dmGlyphOnly
        Layout = blGlyphTop
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BEchClick
        GlobalIndexImage = 'Z0041_S16G2'
      end
      object BLibAuto: TToolbarButton97
        Tag = 1
        Left = 174
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Libell'#233's automatiques'
        Caption = 'Libell'#233's'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 0
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BLibAutoClick
        GlobalIndexImage = 'Z0597_S16G2'
        IsControl = True
      end
      object BInsert: TToolbarButton97
        Tag = 1
        Left = 279
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Ins'#233'rer ligne'
        Caption = 'Ins'#233'rer'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BInsertClick
        GlobalIndexImage = 'Z0074_S16G1'
      end
      object BSDel: TToolbarButton97
        Tag = 1
        Left = 307
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Supprimer ligne'
        Caption = 'Effacer'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Layout = blGlyphTop
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BSDelClick
        GlobalIndexImage = 'Z0072_S16G1'
      end
      object BTools: TToolbarButton97
        Left = 419
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Outils PFU'
        Caption = 'Guide'
        DisplayMode = dmGlyphOnly
        Margin = 3
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        Spacing = 0
        Visible = False
        GlobalIndexImage = 'Z0611_S16G1'
      end
      object BTGuide: TToolbarButton97
        Left = 335
        Top = 0
        Width = 28
        Height = 27
        Hint = 'S'#233'lectionner un guide'
        Caption = 'Guide'
        DisplayMode = dmGlyphOnly
        Margin = 3
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        Spacing = 0
        OnClick = BTGuideClick
        GlobalIndexImage = 'Z0908_S16G1'
      end
      object BModifTva: TToolbarButton97
        Tag = 1
        Left = 118
        Top = 0
        Width = 28
        Height = 27
        Hint = 'El'#233'ments de TVA / Bases Hors-taxes'
        Caption = 'TVA'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        NumGlyphs = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        OnClick = BModifTvaClick
        GlobalIndexImage = 'Z0543_S16G2'
        IsControl = True
      end
      object BSelectFolio: TToolbarButton97
        Left = 363
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Journal Centralisateur'
        NumGlyphs = 2
        Opaque = False
        OnClick = BSelectFolioClick
        GlobalIndexImage = 'Z0808_S16G2'
      end
      object BDevise: TToolbarButton97
        Left = 90
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Saisie autres monnaies'
        Opaque = False
        OnClick = BDeviseClick
        GlobalIndexImage = 'Z0653_S16G1'
      end
      object BLettrage: TToolbarButton97
        Tag = 1
        Left = 391
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Lettrage en saisie'
        Caption = 'Ctrl L'
        DisplayMode = dmGlyphOnly
        Margin = 3
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        Spacing = 0
        OnClick = BLettrageClick
        GlobalIndexImage = 'Z0050_S16G1'
      end
      object BScan: TToolbarButton97
        Tag = 1
        Left = 475
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Visualiser le document associ'#233
        Enabled = False
        Glyph.Data = {
          4E010000424D4E01000000000000760000002800000014000000120000000100
          040000000000D800000000000000000000001000000000000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          7777777700007777777777777777777700007777770077777777777700007777
          7077007777777777000077777077770077777777000077770770077700777777
          00007777070FF0077700777700007770770FFFF0077700770000777070F000FF
          F00770770000770770F8DD0FFF0770770000770770FF88F88F07077700007700
          000000F8F077077700007770FFFFF00FF070777700007770F0EEF07007707777
          00007770F000F0777707777700007770FFFFF007770777770000777000000070
          007777770000777777777777777777770000}
        GlyphMask.Data = {00000000}
        Opaque = False
        Visible = False
        OnClick = BScanClick
      end
      object BAcc: TToolbarButton97
        Tag = 1
        Left = 447
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Acc'#233'l'#233'rateur de saisie'
        Enabled = False
        Opaque = False
        OnClick = BAccClick
        GlobalIndexImage = 'Z0123_S16G1'
      end
    end
  end
  object PFEN: TPanel
    Left = 9
    Top = 9
    Width = 706
    Height = 293
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    OnMouseDown = PFENMouseDown
    object GS: THGrid
      Left = 0
      Top = 29
      Width = 706
      Height = 114
      Align = alClient
      BorderStyle = bsNone
      ColCount = 16
      Ctl3D = True
      DefaultRowHeight = 18
      Enabled = False
      FixedCols = 7
      RowCount = 100
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goEditing, goTabs, goAlwaysShowEditor]
      ParentCtl3D = False
      TabOrder = 2
      OnEnter = GSEnter
      OnExit = GSExit
      OnKeyDown = GSKeyDown
      OnKeyPress = GSKeyPress
      OnMouseDown = GSMouseDown
      OnMouseUp = GSMouseUp
      SortedCol = -1
      Titres.Strings = (
        'Etabl.'
        'Qualif'
        'Cr'#233#233'e'
        'Euro'
        'Pi'#232'ce'
        'Ordre'
        'N'#176
        'Nature'
        'Jr'
        'G'#233'n'#233'ral'
        'Auxiliaire'
        'R'#233'f'#233'rence'
        'Libell'#233
        'D'
        'D'#233'bit'
        'Cr'#233'dit')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      OnRowEnter = GSRowEnter
      OnRowExit = GSRowExit
      OnCellEnter = GSCellEnter
      OnCellExit = GSCellExit
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
      OnElipsisClick = GSElipsisClick
      ColWidths = (
        2
        2
        2
        2
        2
        2
        21
        71
        17
        54
        59
        63
        108
        27
        64
        64)
      RowHeights = (
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18
        18)
    end
    object PEntete: THPanel
      Left = 0
      Top = 0
      Width = 706
      Height = 29
      Align = alTop
      FullRepaint = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      OnEnter = PEnteteEnter
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
      object H_NUMEROPIECE: THLabel
        Left = 372
        Top = 8
        Width = 22
        Height = 13
        Caption = 'Folio'
        FocusControl = E_NUMEROPIECE
        Transparent = True
      end
      object H_JOURNAL: THLabel
        Left = 172
        Top = 8
        Width = 34
        Height = 13
        Caption = 'Journal'
        FocusControl = E_JOURNAL
        Transparent = True
      end
      object H_DATECOMPTABLE: THLabel
        Left = 4
        Top = 8
        Width = 36
        Height = 13
        Caption = 'P'#233'riode'
        FocusControl = E_DATECOMPTABLE
        Transparent = True
      end
      object H_DEVISE: THLabel
        Left = 505
        Top = 8
        Width = 33
        Height = 13
        Caption = 'Devise'
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        Transparent = True
        Visible = False
      end
      object H_EXERCICE: THLabel
        Left = 139
        Top = 8
        Width = 14
        Height = 13
        Caption = '(N)'
        FocusControl = E_DATECOMPTABLE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Transparent = True
      end
      object H_ETABL: THLabel
        Left = 508
        Top = 8
        Width = 27
        Height = 13
        Caption = 'Etabl.'
        FocusControl = E_ETABLISSEMENT
        Transparent = True
      end
      object E_JOURNAL: THValComboBox
        Left = 213
        Top = 4
        Width = 137
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 1
        OnChange = E_JOURNALChange
        OnKeyUp = E_JOURNALKeyUp
        TagDispatch = 0
        DataType = 'TTJALFOLIO'
      end
      object E_DATECOMPTABLE: THValComboBox
        Left = 43
        Top = 4
        Width = 96
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 0
        OnChange = E_DATECOMPTABLEChange
        TagDispatch = 0
        Vide = True
      end
      object E_NUMEROPIECE: THValComboBox
        Left = 417
        Top = 4
        Width = 49
        Height = 21
        ItemHeight = 13
        Sorted = True
        TabOrder = 2
        OnExit = E_NUMEROPIECEExit
        OnKeyPress = E_NUMEROPIECEKeyPress
        TagDispatch = 0
      end
      object BHideFocus: TButton
        Left = 672
        Top = 2
        Width = 10
        Height = 25
        Caption = 'BHideFocus'
        TabOrder = 4
        OnEnter = BHideFocusEnter
      end
      object E_NUMEROPIECEC: THValComboBox
        Left = 432
        Top = 4
        Width = 53
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        Sorted = True
        TabOrder = 3
        Visible = False
        OnExit = E_NUMEROPIECEExit
        OnKeyPress = E_NUMEROPIECEKeyPress
        TagDispatch = 0
      end
      object E_ETABLISSEMENT: THValComboBox
        Left = 569
        Top = 4
        Width = 92
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 5
        OnChange = E_ETABLISSEMENTChange
        TagDispatch = 0
        DataType = 'TTETABLISSEMENT'
      end
    end
    object PPied: THPanel
      Left = 0
      Top = 193
      Width = 706
      Height = 100
      Align = alBottom
      BevelOuter = bvNone
      BorderStyle = bsSingle
      Enabled = False
      FullRepaint = False
      TabOrder = 3
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
      object H_SOLDE: THLabel
        Left = 428
        Top = 63
        Width = 78
        Height = 13
        Caption = 'Solde Journal'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object Bevel2: TBevel
        Left = 472
        Top = 37
        Width = 116
        Height = 17
      end
      object Bevel3: TBevel
        Left = 594
        Top = 37
        Width = 116
        Height = 17
      end
      object Bevel4: TBevel
        Left = 594
        Top = 60
        Width = 116
        Height = 17
      end
      object HConf: TToolbarButton97
        Left = 428
        Top = 57
        Width = 25
        Height = 24
        Hint = 'Ecriture confidentielle'
        Enabled = False
        NoBorder = True
        ParentShowHint = False
        ShowHint = True
        Visible = False
        GlobalIndexImage = 'Z0141_S16G2'
      end
      object ISigneEuro: TImage
        Left = 570
        Top = 61
        Width = 16
        Height = 16
        AutoSize = True
        Picture.Data = {
          07544269746D6170F6000000424DF60000000000000076000000280000001000
          0000100000000100040000000000800000000000000000000000100000001000
          0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
          C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
          FF00FFFFFFFFFFFFFFFFFFFFFFF44444444FFFFFF44444444444FFFF4444FFFF
          FFF4FFFF444FFFFFFFFFFFFF44FFFFFFFFFFF44444444444FFFFFF4444444444
          4FFFFFF44FFFFFFFFFFFF444444444444FFFFF444444444444FFFFFF44FFFFFF
          FFF4FFFF444FFFFFFF44FFFFF444FFFFF444FFFFFF4444444444FFFFFFF44444
          4FF4}
        Stretch = True
        Transparent = True
        Visible = False
      end
      object G_LIBELLE: THLabel
        Left = 4
        Top = 37
        Width = 56
        Height = 13
        Caption = 'G_LIBELLE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Transparent = True
      end
      object T_LIBELLE: THLabel
        Left = 4
        Top = 52
        Width = 55
        Height = 13
        Caption = 'T_LIBELLE'
        Transparent = True
      end
      object LSA_SOLDEG: THLabel
        Left = 81
        Top = 37
        Width = 70
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_SOLDEG'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Transparent = True
      end
      object LSA_SOLDET: THLabel
        Left = 80
        Top = 52
        Width = 69
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_SOLDET'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Transparent = True
      end
      object LSA_SOLDE: THLabel
        Left = 433
        Top = 80
        Width = 8
        Height = 13
        Alignment = taRightJustify
        Caption = '0'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object LSA_TOTALCREDIT: THLabel
        Left = 355
        Top = 79
        Width = 8
        Height = 13
        Alignment = taRightJustify
        Caption = '0'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object LSA_TOTALDEBIT: THLabel
        Left = 222
        Top = 78
        Width = 8
        Height = 13
        Alignment = taRightJustify
        Caption = '0'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object Bevel1: TBevel
        Left = 472
        Top = 9
        Width = 116
        Height = 17
      end
      object Bevel5: TBevel
        Left = 594
        Top = 9
        Width = 116
        Height = 17
      end
      object LSA_FOLIODEBIT: THLabel
        Left = 182
        Top = 2
        Width = 8
        Height = 13
        Hint = 'Total Debit du Folio'
        Alignment = taRightJustify
        Caption = '0'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object LSA_FOLIOCREDIT: THLabel
        Left = 304
        Top = 3
        Width = 8
        Height = 13
        Hint = 'Total Cr'#233'dit du Folio'
        Alignment = taRightJustify
        Caption = '0'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object HLabel1: THLabel
        Left = 428
        Top = 11
        Width = 28
        Height = 13
        Caption = 'Folio'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object HLabel2: THLabel
        Left = 428
        Top = 39
        Width = 42
        Height = 13
        Caption = 'Journal'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object BEVELSOLDEPROG: TBevel
        Left = 158
        Top = 12
        Width = 116
        Height = 17
        Visible = False
      end
      object HLSOLDEPROG: THLabel
        Left = 5
        Top = 5
        Width = 90
        Height = 13
        Caption = 'Solde th'#233'orique'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        Visible = False
      end
      object LSA_SOLDEPROG: THLabel
        Left = 545
        Top = 79
        Width = 8
        Height = 13
        Alignment = taRightJustify
        Caption = '0'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        Visible = False
      end
      object FlashGuide: TFlashingLabel
        Left = 7
        Top = 80
        Width = 40
        Height = 13
        Caption = 'GUIDE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        Visible = False
      end
      object FLASHDEVISE: TFlashingLabel
        Left = 63
        Top = 80
        Width = 46
        Height = 13
        Caption = 'DEVISE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        Visible = False
      end
      object LabelInfos: THLabel
        Left = 7
        Top = 11
        Width = 38
        Height = 13
        Caption = 'INFOS'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        Visible = False
      end
      object LabelInfos2: THLabel
        Left = 312
        Top = 80
        Width = 38
        Height = 13
        Caption = 'INFOS'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        Visible = False
      end
      object BAutoSave: TToolbarButton97
        Left = 404
        Top = 58
        Width = 23
        Height = 22
        Visible = False
        GlobalIndexImage = 'Z0050_S16G1'
      end
      object HLSOLDEPROGCB: THLabel
        Left = 5
        Top = 16
        Width = 90
        Height = 13
        Caption = 'Solde th'#233'orique'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        Visible = False
      end
      object YTC_SCHEMAGEN: THLabel
        Left = 4
        Top = 67
        Width = 95
        Height = 13
        Caption = 'YTC_SCHEMAGEN'
        Transparent = True
      end
      object SoldeD: TLabel
        Left = 223
        Top = 37
        Width = 44
        Height = 13
        Hint = 'Solde de l'#39'encours'
        Alignment = taRightJustify
        Caption = 'a enlever'
        Transparent = True
      end
      object SoldeAuxper: TLabel
        Left = 223
        Top = 52
        Width = 44
        Height = 13
        Hint = 'Solde de l'#39'encours'
        Alignment = taRightJustify
        Caption = 'a enlever'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        Transparent = True
      end
      object SA_FOLIODEBIT: THNumEdit
        Tag = 1
        Left = 473
        Top = 12
        Width = 114
        Height = 15
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 4
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object SA_SOLDEG: THNumEdit
        Tag = 2
        Left = 293
        Top = 37
        Width = 101
        Height = 14
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        NumericType = ntDC
        UseRounding = True
        Validate = False
      end
      object SA_SOLDET: THNumEdit
        Tag = 2
        Left = 293
        Top = 52
        Width = 101
        Height = 15
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        NumericType = ntDC
        UseRounding = True
        Validate = False
      end
      object SA_TOTALCREDIT: THNumEdit
        Tag = 1
        Left = 595
        Top = 40
        Width = 114
        Height = 15
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object SA_SOLDE: THNumEdit
        Tag = 1
        Left = 595
        Top = 63
        Width = 114
        Height = 15
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 3
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        NumericType = ntDC
        UseRounding = True
        Validate = False
      end
      object SA_TOTALDEBIT: THNumEdit
        Tag = 1
        Left = 473
        Top = 37
        Width = 114
        Height = 15
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 5
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object SA_FOLIOCREDIT: THNumEdit
        Tag = 1
        Left = 594
        Top = 12
        Width = 114
        Height = 15
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 6
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object SA_SOLDEPROG: THNumEdit
        Tag = 1
        Left = 160
        Top = 11
        Width = 114
        Height = 15
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 7
        Visible = False
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object SA_SoldePerGen: THNumEdit
        Tag = 2
        Left = 317
        Top = 5
        Width = 101
        Height = 14
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 8
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        NumericType = ntDC
        UseRounding = True
        Validate = False
      end
      object SA_SOLDEAUX: THNumEdit
        Tag = 2
        Left = 317
        Top = 20
        Width = 101
        Height = 15
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        ParentCtl3D = False
        TabOrder = 9
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        NumericType = ntDC
        UseRounding = True
        Validate = False
      end
    end
    object E_NATUREPIECE: THValComboBox
      Left = 164
      Top = 138
      Width = 65
      Height = 21
      Style = csDropDownList
      Color = clYellow
      Ctl3D = False
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 4
      TabStop = False
      Visible = False
      TagDispatch = 0
      HideInGrid = True
      DataType = 'TTNATUREPIECE'
      DisableTab = True
    end
    object SELGUIDE: TEdit
      Left = 256
      Top = 160
      Width = 53
      Height = 21
      Color = clYellow
      TabOrder = 5
      Visible = False
    end
    object BZoomJournal: THBitBtn
      Tag = 100
      Left = 225
      Top = 59
      Width = 28
      Height = 27
      Hint = 'Voir journal'
      Caption = 'Jal'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 7
      Visible = False
      OnClick = BZoomJournalClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BZoom_FB19024: THBitBtn
      Tag = 100
      Left = 193
      Top = 59
      Width = 28
      Height = 27
      Hint = 'Voir compte'
      Caption = 'Cpte'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 6
      Visible = False
      OnClick = BZoom_FB19024Click
      Margin = 2
      Spacing = -1
      IsControl = True
    end
    object BZoomTiers: THBitBtn
      Tag = 100
      Left = 260
      Top = 59
      Width = 28
      Height = 27
      Hint = 'Voir tiers'
      Caption = 'Tiers'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 8
      Visible = False
      OnClick = BZoomTiersClick
      Layout = blGlyphTop
      Spacing = -1
      IsControl = True
    end
    object BZoomDevise: THBitBtn
      Tag = 100
      Left = 295
      Top = 59
      Width = 28
      Height = 27
      Hint = 'Voir devise'
      Caption = 'Dev'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 9
      Visible = False
      OnClick = BZoomDeviseClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BZoomEtabl: THBitBtn
      Tag = 100
      Left = 328
      Top = 59
      Width = 28
      Height = 27
      Hint = 'Voir '#233'tablissement'
      Caption = 'Etab'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 10
      Visible = False
      OnClick = BZoomEtablClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object RichEdit: THRichEditOLE
      Left = 340
      Top = 160
      Width = 205
      Height = 21
      Color = clYellow
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      TabOrder = 11
      Visible = False
      Margins.Top = 0
      Margins.Bottom = 0
      Margins.Left = 0
      Margins.Right = 0
      ContainerName = 'Document'
      ObjectMenuPrefix = '&Object'
      LinesRTF.Strings = (
        
          '{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1036{\fonttbl{\f' +
          '0\fswiss Arial;}}'
        '{\*\generator Riched20 10.0.14393}\viewkind4\uc1 '
        '\pard\f0\fs16 RichEdit'
        '\par '
        '\par '
        '\par }')
    end
    object BZoomCpte: THBitBtn
      Tag = 1000
      Left = 337
      Top = 103
      Width = 28
      Height = 27
      Hint = 'Consultation'
      Caption = 'BZoomCpte'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 12
      Visible = False
      OnClick = BZoomCpteClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BMonnaie: THBitBtn
      Tag = 1000
      Left = 193
      Top = 101
      Width = 28
      Height = 27
      Hint = 'Saisie autres monnaies'
      Caption = 'BMonnaie'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 13
      Visible = False
      OnClick = BDeviseClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BZoomEcrs: THBitBtn
      Tag = 100
      Left = 433
      Top = 59
      Width = 28
      Height = 27
      Hint = 'D'#233'tail des mouvements'
      Caption = 'Ecrs'
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 14
      Visible = False
      OnClick = BZoomEcrsClick
      Margin = 2
      Spacing = -1
      IsControl = True
    end
    object BZoomJalCentral: THBitBtn
      Tag = 1000
      Left = 409
      Top = 103
      Width = 28
      Height = 27
      Hint = 'Journal Centralisateur'
      Caption = 'BZoomJalCentral'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 15
      Visible = False
      OnClick = BSelectFolioClick
      Layout = blGlyphTop
      Spacing = -1
    end
    object BSSCANGED: THBitBtn
      Tag = 1
      Left = 496
      Top = 104
      Width = 75
      Height = 25
      Hint = 'Scanner un documents'
      Caption = 'BSSCANGED'
      TabOrder = 16
      Visible = False
    end
    object BAccactif: THBitBtn
      Tag = 1
      Left = 636
      Top = 48
      Width = 35
      Height = 33
      Hint = 'Arr'#234't de l'#39'acc'#233'l'#233'rateur de saisie'
      Caption = 'Acc'#233'l'#233'rateur'
      TabOrder = 0
      Visible = False
      OnClick = BAccactifClick
    end
    object BAccTiers: THBitBtn
      Tag = 1
      Left = 632
      Top = 88
      Width = 35
      Height = 33
      Hint = 'Activer l'#39'acc'#233'l'#233'rateur de saisie'
      Caption = 'Acc'#233'l'#233'rateur'
      TabOrder = 17
      Visible = False
      OnClick = BAccTiersClick
    end
    object PQuantite: TPanel
      Left = 0
      Top = 143
      Width = 706
      Height = 50
      Align = alBottom
      BevelInner = bvLowered
      BevelOuter = bvNone
      TabOrder = 18
      OnExit = PQuantiteExit
      object GQte: TGroupBox
        Left = 1
        Top = 1
        Width = 704
        Height = 48
        Align = alClient
        Caption = 'Saisie des quantit'#233's'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object H_QTE1: THLabel
          Left = 4
          Top = 18
          Width = 49
          Height = 13
          Caption = '&Quantit'#233' 1'
          Color = clBtnFace
          FocusControl = _QTE1
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object H_QTE2: THLabel
          Left = 356
          Top = 18
          Width = 49
          Height = 13
          Caption = 'Quantit'#233' 2'
          Color = clBtnFace
          FocusControl = _QTE2
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object H_QUALIFQTE1: THLabel
          Left = 152
          Top = 18
          Width = 60
          Height = 13
          Caption = 'Qual&if.  qt'#233' 1'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object H_QUALIFQTE2: THLabel
          Left = 504
          Top = 18
          Width = 60
          Height = 13
          Caption = 'Quali&f.  qt'#233' 2'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object _QTE1: THNumEdit
          Tag = 7
          Left = 56
          Top = 14
          Width = 93
          Height = 23
          AutoSize = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnExit = _QTE1Exit
          OnKeyDown = _QTE1KeyDown
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0.000'
          Masks.NegativeMask = '-#,##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object _QTE2: THNumEdit
          Tag = 7
          Left = 408
          Top = 14
          Width = 93
          Height = 20
          AutoSize = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnExit = _QTE1Exit
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0.000'
          Masks.NegativeMask = '-#,##0.000'
          Debit = False
          UseRounding = True
          Validate = False
        end
      end
    end
    object BPDF: THBitBtn
      Tag = 1000
      Left = 496
      Top = 64
      Width = 75
      Height = 25
      Hint = 'Associer un fichier Pdf'
      Caption = 'BPDF'
      TabOrder = 19
      Visible = False
      OnClick = BPDFClick
    end
  end
  object BModifRIB: THBitBtn
    Tag = 1000
    Left = 239
    Top = 110
    Width = 28
    Height = 27
    Hint = 'Affectation du RIB'
    Caption = 'RIB'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 5
    Visible = False
    OnClick = BModifRIBClick
    Margin = 2
    NumGlyphs = 2
    Spacing = -1
    IsControl = True
  end
  object BZoomImmo: THBitBtn
    Tag = 100
    Left = 405
    Top = 68
    Width = 28
    Height = 27
    Hint = 'Voir immobilisation'
    Caption = 'Immo'
    Enabled = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 6
    Visible = False
    OnClick = BZoomImmoClick
    Layout = blGlyphTop
    Spacing = -1
  end
  object BGuide: THBitBtn
    Tag = 1000
    Left = 276
    Top = 110
    Width = 28
    Height = 27
    Hint = 'Appeler un guide'
    Caption = 'AGuide'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 7
    Visible = False
    OnClick = BTGuideClick
    Margin = 0
    Spacing = -1
    IsControl = True
  end
  object POPS: TPopupMenu
    OnPopup = POPSPopup
    Left = 147
    Top = 112
  end
  object FindSais: TFindDialog
    OnFind = FindSaisFind
    Left = 124
    Top = 89
  end
  object POPZ: TPopupMenu
    AutoPopup = False
    OwnerDraw = True
    Left = 148
    Top = 64
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 88
    Top = 104
  end
  object SD: TSaveDialog
    Left = 643
    Top = 136
  end
end

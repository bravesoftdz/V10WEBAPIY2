object FSaisODA: TFSaisODA
  Left = 231
  Top = 109
  Width = 606
  Height = 395
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Ecriture analytique'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  PopupMenu = POPS
  Position = poScreenCenter
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnKeyPress = FormKeyPress
  OnMouseDown = FormMouseDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object DockTop: TDock97
    Left = 0
    Top = 0
    Width = 590
    Height = 9
  end
  object DockRight: TDock97
    Left = 581
    Top = 9
    Width = 9
    Height = 319
    Position = dpRight
  end
  object DockLeft: TDock97
    Left = 0
    Top = 9
    Width = 9
    Height = 319
    Position = dpLeft
  end
  object DockBottom: TDock97
    Left = 0
    Top = 328
    Width = 590
    Height = 28
    Position = dpBottom
    object Outils97: TToolbar97
      Left = 0
      Top = 0
      Caption = 'Outils'
      CloseButton = False
      DockPos = 0
      TabOrder = 0
      object BSolde: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Solde'
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
      object BVentilType: TToolbarButton97
        Tag = 1
        Left = 27
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Ventilations types'
        Caption = 'Ventilations'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 0
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BVentilTypeClick
        GlobalIndexImage = 'Z1883_S16G1'
        IsControl = True
      end
      object BComplement: TToolbarButton97
        Tag = 1
        Left = 54
        Top = 0
        Width = 27
        Height = 24
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
      object BSoldeGS: TToolbarButton97
        Tag = 1
        Left = 81
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Calcul du solde G'#233'n'#233'ral/Section'
        Caption = 'Calcul solde'
        DisplayMode = dmGlyphOnly
        Layout = blGlyphTop
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BSoldeGSClick
        GlobalIndexImage = 'Z0123_S16G1'
      end
      object BChercher: TToolbarButton97
        Tag = 1
        Left = 108
        Top = 0
        Width = 27
        Height = 24
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
        Left = 141
        Top = 0
        Width = 37
        Height = 24
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
        Left = 135
        Top = 0
      end
    end
    object Valide97: TToolbar97
      Left = 495
      Top = 0
      Caption = 'Validation'
      CloseButton = False
      DockPos = 495
      TabOrder = 1
      object BValide: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 27
        Height = 24
        Hint = 'Valider la saisie'
        Caption = 'Valider'
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
        Left = 27
        Top = 0
        Width = 27
        Height = 24
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
        Left = 54
        Top = 0
        Width = 27
        Height = 24
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
  end
  object PFEN: TPanel
    Left = 9
    Top = 9
    Width = 572
    Height = 319
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object PEntete: THPanel
      Left = 0
      Top = 0
      Width = 572
      Height = 53
      Align = alTop
      FullRepaint = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
      object Bevel1: TBevel
        Left = 485
        Top = 28
        Width = 104
        Height = 21
      end
      object H_JOURNAL: THLabel
        Left = 8
        Top = 8
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = Y_JOURNAL
        Transparent = True
        OnDblClick = H_JOURNALDblClick
      end
      object H_DATECOMPTABLE: THLabel
        Left = 300
        Top = 8
        Width = 23
        Height = 13
        Caption = '&Date'
        FocusControl = Y_DATECOMPTABLE
        Transparent = True
      end
      object H_NUMEROPIECE: THLabel
        Left = 412
        Top = 32
        Width = 37
        Height = 13
        Caption = 'Num'#233'ro'
        Transparent = True
      end
      object H_GENERAL: THLabel
        Left = 8
        Top = 32
        Width = 37
        Height = 13
        Caption = '&G'#233'n'#233'ral'
        FocusControl = Y_GENERAL
        Transparent = True
        OnDblClick = H_GENERALDblClick
      end
      object Y_NUMEROPIECE: THLabel
        Left = 487
        Top = 30
        Width = 100
        Height = 16
        Alignment = taCenter
        AutoSize = False
        Caption = 'Num pi'#232'ce'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object H_ETABLISSEMENT: THLabel
        Left = 412
        Top = 8
        Width = 65
        Height = 13
        Caption = '&Etablissement'
        FocusControl = Y_ETABLISSEMENT
        Transparent = True
      end
      object H_AXE: THLabel
        Left = 200
        Top = 8
        Width = 89
        Height = 13
        AutoSize = False
        Transparent = True
      end
      object H_LIBELLE: THLabel
        Left = 192
        Top = 32
        Width = 204
        Height = 13
        AutoSize = False
        Transparent = True
      end
      object Y_JOURNAL: THValComboBox
        Left = 52
        Top = 4
        Width = 137
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = Y_JOURNALChange
        TagDispatch = 0
        DataType = 'TTJALANALYTIQUE'
      end
      object Y_DATECOMPTABLE: TMaskEdit
        Left = 328
        Top = 4
        Width = 69
        Height = 21
        EditMask = '!99/99/0000;1;_'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 10
        ParentFont = False
        TabOrder = 1
        Text = '  /  /    '
        OnChange = Y_DATECOMPTABLEChange
        OnExit = Y_DATECOMPTABLEExit
      end
      object Cache: THCpteEdit
        Left = 276
        Top = 6
        Width = 17
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        Text = '!!!'
        Visible = False
        ZoomTable = tzSection
        Vide = False
        Bourre = True
        okLocate = False
        SynJoker = False
      end
      object Y_ETABLISSEMENT: THValComboBox
        Left = 484
        Top = 4
        Width = 105
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
        TagDispatch = 0
        DataType = 'TTETABLISSEMENT'
      end
      object Y_GENERAL: THCpteEdit
        Left = 52
        Top = 28
        Width = 137
        Height = 21
        MaxLength = 17
        TabOrder = 3
        OnChange = Y_GENERALChange
        OnExit = Y_GENERALExit
        ZoomTable = tzGventilable
        Vide = False
        Bourre = True
        Libelle = H_LIBELLE
        okLocate = False
        SynJoker = False
      end
    end
    object GSO: THGrid
      Left = 0
      Top = 53
      Width = 572
      Height = 212
      Align = alClient
      BorderStyle = bsNone
      ColCount = 10
      DefaultRowHeight = 18
      Enabled = False
      FixedCols = 4
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing]
      TabOrder = 1
      OnDblClick = GSODblClick
      OnEnter = GSOEnter
      OnExit = GSOExit
      OnKeyPress = GSOKeyPress
      OnSetEditText = GSOSetEditText
      SortedCol = -1
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      OnRowEnter = GSORowEnter
      OnRowExit = GSORowExit
      OnCellEnter = GSOCellEnter
      OnCellExit = GSOCellExit
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
      ColWidths = (
        0
        0
        0
        22
        66
        77
        99
        154
        77
        77)
    end
    object BInsert: THBitBtn
      Tag = 1000
      Left = 262
      Top = 175
      Width = 28
      Height = 27
      Hint = 'Ins'#233'rer ligne'
      Caption = 'Ins.'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      Visible = False
      OnClick = BInsertClick
      Margin = 2
    end
    object BSDel: THBitBtn
      Tag = 1000
      Left = 294
      Top = 175
      Width = 28
      Height = 27
      Hint = 'Supprimer ligne'
      Caption = 'Del.'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 4
      Visible = False
      OnClick = BSDelClick
      Margin = 2
    end
    object PPied: THPanel
      Left = 0
      Top = 265
      Width = 572
      Height = 54
      Align = alBottom
      BevelOuter = bvNone
      BorderStyle = bsSingle
      Enabled = False
      FullRepaint = False
      TabOrder = 2
      BackGroundEffect = bdFond
      ColorShadow = clWindowText
      ColorStart = clBtnFace
      ColorNb = 2
      TextEffect = tenone
      object BevCache: TBevel
        Left = 354
        Top = 4
        Width = 118
        Height = 19
      end
      object Bevel3: TBevel
        Left = 476
        Top = 4
        Width = 117
        Height = 19
      end
      object Bevel4: TBevel
        Left = 476
        Top = 28
        Width = 117
        Height = 19
      end
      object H_MONTANTECR: THLabel
        Left = 378
        Top = 7
        Width = 87
        Height = 13
        Caption = 'Montant Ecr. G'#233'n.'
        Transparent = True
        Visible = False
      end
      object H_SOLDE: THLabel
        Left = 412
        Top = 32
        Width = 56
        Height = 13
        Caption = 'Solde pi'#232'ce'
        Transparent = True
      end
      object HConf: TToolbarButton97
        Left = 354
        Top = 24
        Width = 25
        Height = 24
        Hint = 'Ecriture confidentielle'
        Enabled = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        GlobalIndexImage = 'Z0141_S16G2'
      end
      object S_LIBELLE: THLabel
        Left = 4
        Top = 8
        Width = 3
        Height = 13
        Caption = ' '
        Transparent = True
      end
      object GS_LIBELLE: THLabel
        Left = 4
        Top = 30
        Width = 3
        Height = 13
        Caption = ' '
        Transparent = True
      end
      object LOA_SoldeS: THLabel
        Left = 52
        Top = 4
        Width = 61
        Height = 13
        Alignment = taRightJustify
        Caption = 'LOA_SoldeS'
        Transparent = True
        OnDblClick = H_JOURNALDblClick
      end
      object LOA_SoldeGS: THLabel
        Left = 52
        Top = 28
        Width = 69
        Height = 13
        Alignment = taRightJustify
        Caption = 'LOA_SoldeGS'
        Transparent = True
        OnDblClick = H_JOURNALDblClick
      end
      object LOA_TotalDebit: THLabel
        Left = 93
        Top = 4
        Width = 91
        Height = 13
        Alignment = taRightJustify
        Caption = 'LOA_TotalDebit'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        OnDblClick = H_JOURNALDblClick
      end
      object LOA_TotalCredit: THLabel
        Left = 88
        Top = 28
        Width = 94
        Height = 13
        Alignment = taRightJustify
        Caption = 'LOA_TotalCredit'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        OnDblClick = H_JOURNALDblClick
      end
      object LOA_MontantEcr: THLabel
        Left = 185
        Top = 4
        Width = 97
        Height = 13
        Alignment = taRightJustify
        Caption = 'LOA_MontantEcr'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        OnDblClick = H_JOURNALDblClick
      end
      object LOA_Solde: THLabel
        Left = 190
        Top = 28
        Width = 64
        Height = 13
        Alignment = taRightJustify
        Caption = 'LOA_Solde'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        OnDblClick = H_JOURNALDblClick
      end
      object NONCALC: THLabel
        Left = 248
        Top = 30
        Width = 88
        Height = 13
        Alignment = taRightJustify
        Caption = 'Solde Non Calcul'#233
        Transparent = True
        Visible = False
      end
      object OA_TOTALCREDIT: THNumEdit
        Tag = 1
        Left = 477
        Top = 7
        Width = 115
        Height = 15
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        Enabled = False
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
      object OA_MONTANTECR: THNumEdit
        Tag = 1
        Left = 477
        Top = 7
        Width = 115
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
      object OA_TOTALDEBIT: THNumEdit
        Tag = 1
        Left = 356
        Top = 7
        Width = 115
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
        UseRounding = True
        Validate = False
      end
      object OA_SOLDEGS: THNumEdit
        Tag = 2
        Left = 248
        Top = 30
        Width = 100
        Height = 16
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        ParentCtl3D = False
        TabOrder = 1
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        NumericType = ntDC
        UseRounding = True
        Validate = False
      end
      object OA_SOLDES: THNumEdit
        Tag = 2
        Left = 248
        Top = 7
        Width = 100
        Height = 19
        BorderStyle = bsNone
        Color = clBtnFace
        Ctl3D = False
        ParentCtl3D = False
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
      object OA_SOLDE: THNumEdit
        Tag = 1
        Left = 477
        Top = 31
        Width = 115
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
        NumericType = ntDC
        UseRounding = True
        Validate = False
      end
    end
    object BZoomPiece: THBitBtn
      Tag = 100
      Left = 327
      Top = 175
      Width = 28
      Height = 27
      Hint = 'Voir '#233'criture'
      Caption = 'Piec'
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
      OnClick = BZoomPieceClick
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object BZoom: THBitBtn
      Tag = 100
      Left = 359
      Top = 175
      Width = 28
      Height = 27
      Hint = 'Voir Section'
      Caption = 'Sect'
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
      OnClick = BZoomClick
      Margin = 2
      Spacing = -1
      IsControl = True
    end
    object BZoomEtabl: THBitBtn
      Tag = 100
      Left = 391
      Top = 175
      Width = 28
      Height = 27
      Hint = 'Voir Etablissement'
      Caption = 'Etab'
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
      OnClick = BZoomEtablClick
      Margin = 2
      Spacing = -1
      IsControl = True
    end
  end
  object POPS: TPopupMenu
    OnPopup = POPSPopup
    Left = 112
    Top = 182
  end
  object HMessLigne: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Vous devez renseigner une section analytique existan' +
        'te;W;O;O;O;'
      '1;?caption?;Vous devez saisir un montant;W;O;O;O;'
      '2;?caption?;Attention : cette section est ferm'#233'e;E;O;O;O;'
      
        '3;?caption?;Vous ne pouvez pas appeler une section confidentiell' +
        'e en modification des analytiques;W;O;O;O;'
      
        '4;?caption?;Attention : cette section est ferm'#233'e et sold'#233'e. Vous' +
        ' ne pouvez plus l'#39'utiliser en saisie;E;O;O;O;')
    Left = 44
    Top = 232
  end
  object HMessPiece: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Votre pi'#232'ce est incorrecte : vous devez saisir au mo' +
        'ins une ligne;W;O;O;O;'
      '1;?caption?;Confirmez-vous l'#39'abandon de la saisie ?;Q;YN;N;N;'
      
        '2;?caption?;Votre modification est incorrecte : ce compte n'#39'est ' +
        'pas ventilable sur l'#39'axe ;W;O;O;O;'
      
        '3;?caption?;Votre saisie est incorrecte : ce compte g'#233'n'#233'ral n'#39'ex' +
        'iste pas;W;O;N;N;'
      
        '4;?caption?;La pi'#232'ce est valid'#233'e. Seule la consultation est auto' +
        'ris'#233'e;E;O;O;O;'
      
        '5;?caption?;Le solde de la pi'#232'ce diff'#232're du montant de l'#39#233'critur' +
        'e g'#233'n'#233'rale. Vous devez le corriger.;W;O;O;O;'
      
        '6;?caption?;La date que vous avez renseign'#233'e n'#39'est pas valide;W;' +
        'O;O;O;'
      
        '7;?caption?;La date que vous avez renseign'#233'e n'#39'est pas dans un e' +
        'xercice ouvert;W;O;O;O;'
      
        '8;?caption?;La date que vous avez renseign'#233'e est ant'#233'rieure '#224' un' +
        'e cl'#244'ture;W;O;O;O;'
      
        '9;?caption?;La date que vous avez renseign'#233'e est ant'#233'rieure '#224' un' +
        'e cl'#244'ture;W;O;O;O;'
      
        '10;?caption?;La date que vous avez renseign'#233'e est en dehors des ' +
        'limites autoris'#233'es;W;O;O;O;'
      
        '11;?caption?;Ventilation des quantit'#233's impossible : les qualifia' +
        'nts des quantit'#233's lignes sont diff'#233'rents;W;O;O;O;'
      
        '12;?caption?;Ce compte g'#233'n'#233'ral est interdit. Vous devez le modif' +
        'ier;W;O;O;O;'
      
        '13;?caption?;Ce journal est incorrect : il ne poss'#232'de pas de sou' +
        'che de num'#233'rotation;W;O;O;O;'
      
        '14;?caption?;La pi'#232'ce est en devise. Seule la consultation est a' +
        'utoris'#233'e (utilisez la modification des '#233'critures g'#233'n'#233'rales);E;O;' +
        'O;O;'
      
        '15;?caption?;Vous devez renseigner un compte g'#233'n'#233'ral valide;W;O;' +
        'O;O;'
      
        '16;?caption?;Certaines sections sur le compte g'#233'n'#233'ral ne sont pa' +
        's compatibles avec la d'#233'finition du budget. Confirmez-vous la va' +
        'lidation ?;Q;YN;N;N;'
      
        '17;?caption?;La pi'#232'ce a '#233't'#233' saisie en Euro. Seule la consultatio' +
        'n est autoris'#233'e (utilisez la modification des '#233'critures g'#233'n'#233'rale' +
        's);E;O;O;O;'
      
        '18;?caption?;Seule la consultation est autoris'#233'e. La pi'#232'ce a '#233't'#233 +
        ' saisie en;E;O;O;O;'
      
        '19;?caption?;Vous n'#39'avez pas le droit de saisir sur ce journal;W' +
        ';O;O;O;'
      
        '20;?caption?;Vous n'#39'avez pas le droit de saisir sur ce journal. ' +
        'Seule la consultation est autoris'#233'e;W;O;O;O;'
      '21;?caption?;Vous devez '#233'quilibrer l'#39#233'criture;W;O;O;O;'
      ' ')
    Left = 108
    Top = 232
  end
  object FindSais: TFindDialog
    OnFind = FindSaisFind
    Left = 152
    Top = 182
  end
  object HTitres: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Visualisation des '#233'critures analytiques'
      'Modification des '#233'critures analytiques'
      'Saisie des '#233'critures analytiques'
      'Visualisation des '#233'critures analytiques de simulation'
      'Modification des '#233'critures analytiques de simulation'
      'Saisie des '#233'critures analytiques de simulation'
      'Choix d'#39'une ventilation type'
      'Visualisation des A-Nouveaux analytiques'
      'Modification des A-Nouveaux analytiques'
      'Saisie des A-Nouveaux analytiques'
      '10;')
    Left = 160
    Top = 232
  end
  object HDiv: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        'Validation impossible! Certaines '#233'critures ont '#233't'#233' modifi'#233'es par' +
        ' un autre utilisateur'
      'ATTENTION : Pi'#232'ce non valid'#233'e'
      
        'ATTENTION! Le num'#233'ro de pi'#232'ce n'#39'est pas d'#233'finitivement attribu'#233'.' +
        ' Abandonnez la saisie et r'#233'-essayez de nouveau'
      '3;')
    Left = 200
    Top = 232
  end
  object POPZ: TPopupMenu
    Left = 48
    Top = 182
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 240
    Top = 232
  end
end

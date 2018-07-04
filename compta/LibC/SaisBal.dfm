object FSaisieBalance: TFSaisieBalance
  Left = 287
  Top = 152
  Width = 608
  Height = 372
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Saisie balance'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clBlack
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
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
    Width = 592
    Height = 9
  end
  object DockRight: TDock97
    Left = 583
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
  object DockBottom: TDock97
    Left = 0
    Top = 302
    Width = 592
    Height = 31
    Position = dpBottom
    object Valide97: TToolbar97
      Left = 502
      Top = 0
      Caption = 'Validation'
      CloseButton = False
      DockPos = 502
      TabOrder = 0
      object BValide: TToolbarButton97
        Tag = 1
        Left = 0
        Top = 0
        Width = 27
        Height = 27
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
        Left = 54
        Top = 0
        Width = 27
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
        Width = 27
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
      object BChercher: TToolbarButton97
        Tag = 1
        Left = 33
        Top = 0
        Width = 27
        Height = 27
        Hint = 'Rechercher dans la balance'
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
      object ToolbarSep971: TToolbarSep97
        Left = 27
        Top = 0
      end
      object Sep97: TToolbarSep97
        Left = 143
        Top = 0
      end
      object BNew: TToolbarButton97
        Left = 88
        Top = 0
        Width = 27
        Height = 27
        Hint = 'Nouveau Compte'
        Caption = 'Nouveau'
        DisplayMode = dmGlyphOnly
        Layout = blGlyphTop
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BNewClick
        GlobalIndexImage = 'Z0053_S16G1'
      end
      object BImprimer: TToolbarButton97
        Left = 115
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Caption = 'Imprimer'
        DisplayMode = dmGlyphOnly
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object bExport: TToolbarButton97
        Left = 60
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Exporter la liste'
        Caption = 'Exporter'
        DisplayMode = dmGlyphOnly
        Layout = blGlyphTop
        Opaque = False
        OnClick = bExportClick
        GlobalIndexImage = 'Z0724_S16G1'
      end
    end
  end
  object PFEN: TPanel
    Left = 9
    Top = 9
    Width = 574
    Height = 293
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    OnMouseDown = PFENMouseDown
    object GS: THGrid
      Left = 0
      Top = 29
      Width = 574
      Height = 164
      Align = alClient
      BorderStyle = bsNone
      ColCount = 11
      Ctl3D = True
      DefaultRowHeight = 18
      Enabled = False
      FixedCols = 6
      RowCount = 500
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goEditing, goTabs, goAlwaysShowEditor]
      ParentCtl3D = False
      TabOrder = 1
      OnDblClick = GSDblClick
      OnEnter = GSEnter
      OnExit = GSExit
      OnKeyPress = GSKeyPress
      OnMouseDown = GSMouseDown
      SortedCol = -1
      Titres.Strings = (
        'Etabl.'
        'Cr'#233#233'e'
        'Collectif'
        'TypeGene'
        'Lettrable'
        'ModeRegl'
        'G'#233'n'#233'ral'
        'Auxiliaire'
        'Libell'#233
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
      ColWidths = (
        2
        2
        2
        2
        2
        2
        57
        53
        189
        129
        117)
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
        18
        18
        18
        18
        18)
    end
    object PEntete: THPanel
      Left = 0
      Top = 0
      Width = 574
      Height = 29
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
      object H_DEVISE: THLabel
        Left = 441
        Top = 8
        Width = 33
        Height = 13
        Caption = 'Devise'
        FocusControl = HB_DEVISE
        Transparent = True
      end
      object HLBAL: THLabel
        Left = 191
        Top = 8
        Width = 55
        Height = 13
        Caption = 'Balance Au'
        FocusControl = DATE1
        Transparent = True
      end
      object BValEntete: TToolbarButton97
        Left = 160
        Top = 2
        Width = 25
        Height = 25
        Hint = 'Appliquer crit'#232'res'
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BValEnteteClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object HB_DEVISE: THValComboBox
        Left = 483
        Top = 4
        Width = 94
        Height = 21
        Style = csDropDownList
        Enabled = False
        ItemHeight = 13
        TabOrder = 1
        TabStop = False
        OnChange = HB_DEVISEChange
        TagDispatch = 0
        DataType = 'TTDEVISE'
      end
      object DATE1: TMaskEdit
        Left = 280
        Top = 4
        Width = 92
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 0
        Text = '  /  /    '
        OnKeyPress = DATE1KeyPress
      end
      object BHide: TButton
        Left = 406
        Top = 1
        Width = 10
        Height = 25
        Caption = 'BHideFocus'
        TabOrder = 2
        Visible = False
        OnEnter = BHideFocusEnter
      end
      object BHideFocus: TButton
        Left = 422
        Top = 1
        Width = 10
        Height = 25
        Caption = 'BHideFocus'
        TabOrder = 3
        Visible = False
        OnEnter = BHideFocusEnter
      end
    end
    object PPied: THPanel
      Left = 0
      Top = 193
      Width = 574
      Height = 100
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
      object H_SOLDE: THLabel
        Left = 260
        Top = 67
        Width = 93
        Height = 13
        Caption = 'Solde Progressif'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object Bevel2: TBevel
        Left = 356
        Top = 37
        Width = 116
        Height = 17
      end
      object Bevel3: TBevel
        Left = 478
        Top = 37
        Width = 116
        Height = 17
      end
      object Bevel4: TBevel
        Left = 478
        Top = 66
        Width = 116
        Height = 17
      end
      object HConf: TToolbarButton97
        Left = 8
        Top = 73
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
      object LSA_SOLDE: THLabel
        Left = 381
        Top = 66
        Width = 72
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_SOLDE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object LSA_TOTALCREDIT: THLabel
        Left = 358
        Top = 83
        Width = 117
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_TOTALCREDIT'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object LSA_TOTALDEBIT: THLabel
        Left = 254
        Top = 82
        Width = 108
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_TOTALDEBIT'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object Bevel1: TBevel
        Left = 356
        Top = 9
        Width = 116
        Height = 17
      end
      object Bevel5: TBevel
        Left = 478
        Top = 9
        Width = 116
        Height = 17
      end
      object LSA_CUMULDEBIT: THLabel
        Left = 7
        Top = 6
        Width = 111
        Height = 13
        Hint = 'Total Debit du Folio'
        Alignment = taRightJustify
        Caption = 'LSA_CUMULDEBIT'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object LSA_CUMULCREDIT: THLabel
        Left = 131
        Top = 7
        Width = 120
        Height = 13
        Hint = 'Total Cr'#233'dit du Folio'
        Alignment = taRightJustify
        Caption = 'LSA_CUMULCREDIT'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object HLCUMUL: THLabel
        Left = 260
        Top = 11
        Width = 75
        Height = 13
        Caption = 'Cumul classe'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object HLabel2: THLabel
        Left = 260
        Top = 39
        Width = 92
        Height = 13
        Caption = 'Totaux Comptes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object FLASHDEVISE: TFlashingLabel
        Left = 63
        Top = 80
        Width = 46
        Height = 13
        Caption = 'DEVISE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlue
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
        Visible = False
      end
      object Bevel6: TBevel
        Left = 124
        Top = 37
        Width = 116
        Height = 17
      end
      object HLRESULTAT: THLabel
        Left = 32
        Top = 39
        Width = 48
        Height = 13
        Caption = 'R'#233'sultat'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object LSA_RESULTAT: THLabel
        Left = 9
        Top = 18
        Width = 96
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_RESULTAT'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object Bevel7: TBevel
        Left = 125
        Top = 66
        Width = 116
        Height = 17
      end
      object HLabel1: THLabel
        Left = 32
        Top = 67
        Width = 83
        Height = 13
        Caption = 'Solde Balance'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object LSA_BALANCE: THLabel
        Left = 2
        Top = 54
        Width = 87
        Height = 13
        Alignment = taRightJustify
        Caption = 'LSA_BALANCE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object SA_CUMULDEBIT: THNumEdit
        Tag = 1
        Left = 357
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
        TabOrder = 2
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object SA_TOTALCREDIT: THNumEdit
        Tag = 1
        Left = 479
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
        TabOrder = 0
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
        Left = 479
        Top = 69
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
        NumericType = ntDC
        UseRounding = True
        Validate = False
      end
      object SA_TOTALDEBIT: THNumEdit
        Tag = 1
        Left = 357
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
        TabOrder = 3
        OnChange = GereAffSolde
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object SA_CUMULCREDIT: THNumEdit
        Tag = 1
        Left = 478
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
      object SA_RESULTAT: THNumEdit
        Tag = 1
        Left = 126
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
      object SA_BALANCE: THNumEdit
        Tag = 1
        Left = 126
        Top = 67
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
    end
    object GSPrint: THGrid
      Left = 412
      Top = 76
      Width = 156
      Height = 120
      FixedCols = 0
      TabOrder = 3
      Visible = False
      SortedCol = -1
      Titres.Strings = (
        'G'#233'n'#233'ral'
        'Auxiliaire'
        'Libell'#233
        'D'#233'bit'
        'Cr'#233'dit')
      Couleur = False
      MultiSelect = False
      TitleBold = True
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = False
      AlternateColor = clSilver
    end
  end
  object HBalance: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?caption?;Balance non sold'#233'e;E;O;O;O;'
      '1;?caption?;Impossible de sauvegarder la balance;W;O;O;O;'
      
        '2;?caption?;Une saisie balance est en cours sur un autre poste;E' +
        ';O;O;O;'
      'Cumul Classe'
      'Chargement du plan comptable en cours'
      'Balance Au'
      'Perte'
      'B'#233'n'#233'fice'
      'Cr'#233'ation de compte interdit'
      'Cr'#233'ation de compte tiers interdit'
      'Exercices'
      
        '11;?caption?;Vous ne pouvez pas saisir de balance N-1 car il exi' +
        'ste une balance N-1 pour comparatif;E;O;O;O;'
      
        '12;?caption?;Aucun exercice de r'#233'f'#233'rence pour saisir la balance;' +
        'E;O;O;O;'
      
        '13;?caption?;Les A Nouveaux ont '#233't'#233' modifi'#233's (Lettrage et/ou Poi' +
        'ntage);E;O;O;O;'
      
        '14;?caption?;Vous ne pouvez pas saisir de balance N-1 car il exi' +
        'ste une balance N;E;O;O;O;'
      
        '15;?caption?;Vous ne pouvez pas saisir de balance pour comparati' +
        'f car il existe une balance N-1;E;O;O;O;'
      
        '16;?caption?;Des '#233'critures ont d'#233'j'#224' '#233't'#233' saisies sur l'#39'exercice;E' +
        ';O;O;O;'
      
        '17;?caption?;La date est en dehors de l'#39'exercice courant;E;O;O;O' +
        ';'
      
        '18;?caption?;Voulez-vous saisir le d'#233'tail des comptes auxiliaire' +
        's ?;Q;YN;Y;C;'
      'Choisir une balance de situation'
      'Aucune balance de situation'
      
        '21;?caption?;Confirmez-vous la supression de la balance ?;Q;YN;Y' +
        ';C;'
      
        '22;?caption?;Aucune '#233'criture pour la p'#233'riode s'#233'lectionn'#233'e;E;O;O;' +
        'O;'
      '23;?caption?;Veuillez param'#233'trer les comptes d'#39'attente;E;O;O;O;'
      
        '24;?caption?;Le journal de reprise a '#233't'#233' modifi'#233'. Confirmez la m' +
        'odification de balance ?;Q;YN;Y;C;'
      
        '25;?caption?;Cette balance existe d'#233'j'#224'. Voulez-vous la modifier ' +
        '?;Q;YN;Y;C;'
      'Balance N-1 avec A-Nouveaux'
      'Balance N-1 pour comparatif'
      'Balance N'
      'Balance de situation'
      
        '30;?caption?;Vous ne pouvez pas saisir des montants n'#233'gatifs;W;O' +
        ';O;O;'
      
        '31;?caption?;Le montant que vous avez saisi est en dehors de la ' +
        'fourchette autoris'#233'e;W;O;O;O;'
      
        '32;Balance N-1 avec A Nouveaux;Veuillez param'#233'trer le journal d'#39 +
        #39'ouverture ?;E;O;O;O;'
      
        '33;Balance N;Veuillez param'#233'trer le journal de reprise de balanc' +
        'e;E;O;O;O;'
      'Cr'#233'ation'
      'Modification'
      'G'#233'n'#233'ration'
      '37;?caption?;Voulez-vous abandonner la saisie ?;Q;YN;N;N;'
      
        '38;?caption?;Souhaitez-vous pouvoir lettrer et pointer les '#224'-nou' +
        'veaux g'#233'n'#233'r'#233's ?;Q;YN;N;N;'
      
        '39;?caption?;Souhaitez-vous un calcul automatique du r'#233'sultat ?;' +
        'Q;YN;Y;C;')
    Left = 102
    Top = 159
  end
  object FindSais: TFindDialog
    OnFind = FindSaisFind
    Left = 200
    Top = 161
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 152
    Top = 161
  end
  object SD: TSaveDialog
    DefaultExt = 'XLS'
    Filter = 
      'Fichier Texte (*.txt)|*.txt|Fichier Excel (*.xls)|*.xls|Fichier ' +
      'Ascii (*.asc)|*.asc|Fichier Lotus (*.wks)|*.wks'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofPathMustExist, ofNoReadOnlyReturn, ofNoLongNames]
    Title = 'Export'
    Left = 240
    Top = 160
  end
end

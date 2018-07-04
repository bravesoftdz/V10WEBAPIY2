object FEncaDeca: TFEncaDeca
  Left = 243
  Top = 157
  Width = 639
  Height = 408
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Encaissements / D'#233'caissements'
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
  object DockBottom: TDock97
    Left = 0
    Top = 337
    Width = 631
    Height = 44
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 40
      ClientWidth = 631
      Caption = 'Actions'
      ClientAreaHeight = 40
      ClientAreaWidth = 631
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BImprimer: TToolbarButton97
        Tag = 1
        Left = 497
        Top = 7
        Width = 27
        Height = 27
        Hint = 'Imprimer'
        DisplayMode = dmGlyphOnly
        Caption = 'Imprimer'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BValide: TToolbarButton97
        Tag = 1
        Left = 524
        Top = 7
        Width = 27
        Height = 27
        Hint = 'G'#233'n'#233'rer le traitement'
        DisplayMode = dmGlyphOnly
        Caption = 'Valider'
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = BValideClick
        GlobalIndexImage = 'Z0184_S16G1'
      end
      object BAnnuler: TToolbarButton97
        Tag = 1
        Left = 551
        Top = 7
        Width = 27
        Height = 27
        Hint = 'Fermer'
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = BAnnulerClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 578
        Top = 7
        Width = 27
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
      object bSelectAll: TToolbarButton97
        Tag = 1
        Left = 284
        Top = 7
        Width = 27
        Height = 27
        Hint = 'Tout s'#233'lectioner'
        DisplayMode = dmGlyphOnly
        Caption = 'Tout s'#233'lectionner'
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = bSelectAllClick
        GlobalIndexImage = 'Z0205_S16G1'
      end
      object BLanceSaisieBor: TToolbarButton97
        Tag = 1
        Left = 257
        Top = 7
        Width = 27
        Height = 27
        Hint = 'Saisir un bordereau'
        DisplayMode = dmGlyphOnly
        Caption = 'Chercher'
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = BLanceSaisieBorClick
        GlobalIndexImage = 'Z1775_S16G1'
      end
      object BLanceSaisie: TToolbarButton97
        Tag = 1
        Left = 230
        Top = 7
        Width = 27
        Height = 27
        Hint = 'Saisir une pi'#232'ce'
        DisplayMode = dmGlyphOnly
        Caption = 'Agrandir'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BLanceSaisieClick
        GlobalIndexImage = 'Z1995_S16G1'
      end
      object BAgrandir: TToolbarButton97
        Left = 14
        Top = 7
        Width = 27
        Height = 27
        Hint = 'Agrandir la liste'
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Caption = 'Agrandir'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAgrandirClick
        GlobalIndexImage = 'Z0270_S16G1'
      end
      object BRecherche: TToolbarButton97
        Tag = 1
        Left = 41
        Top = 7
        Width = 27
        Height = 27
        Hint = 'Rechercher dans la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'Chercher'
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = BRechercheClick
        GlobalIndexImage = 'Z0077_S16G1'
      end
      object BZoomPiece: TToolbarButton97
        Tag = 1
        Left = 68
        Top = 7
        Width = 27
        Height = 27
        Hint = 'Voir '#233'criture'
        DisplayMode = dmGlyphOnly
        Caption = 'Ecriture'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BZoomPieceClick
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
      object BModifEche: TToolbarButton97
        Tag = 1
        Left = 95
        Top = 7
        Width = 27
        Height = 27
        Hint = 'Modification de l'#39#233'ch'#233'ance'
        DisplayMode = dmGlyphOnly
        Caption = 'Ech'#233'ance'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BModifEcheClick
        GlobalIndexImage = 'Z0357_S16G2'
        IsControl = True
      end
      object bAffectBqe: TToolbarButton97
        Tag = 1
        Left = 122
        Top = 7
        Width = 27
        Height = 27
        Hint = 'Affectation pr'#233'visionnelle des banques'
        DisplayMode = dmGlyphOnly
        Caption = 'Banques'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = bAffectBqeClick
        GlobalIndexImage = 'Z0159_S16G2'
        IsControl = True
      end
      object BCtrlRib: TToolbarButton97
        Tag = 1
        Left = 149
        Top = 7
        Width = 27
        Height = 27
        Hint = 'Contr'#244'ler les RIB de la s'#233'lection'
        DisplayMode = dmGlyphOnly
        Caption = 'RIB'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        OnClick = BCtrlRibClick
        GlobalIndexImage = 'Z0584_S16G1'
        IsControl = True
      end
      object BModifRIB: TToolbarButton97
        Tag = 1
        Left = 176
        Top = 7
        Width = 27
        Height = 27
        Hint = 'Modifier le RIB sur le mouvement'
        DisplayMode = dmGlyphOnly
        Caption = 'RIB'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BModifRIBClick
        GlobalIndexImage = 'Z0003_S16G1'
        IsControl = True
      end
      object BZoomModele: TToolbarButton97
        Tag = 1
        Left = 203
        Top = 7
        Width = 27
        Height = 27
        Hint = 'Mod'#232'le de document'
        DisplayMode = dmGlyphOnly
        Caption = 'RIB'
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BZoomModeleClick
        GlobalIndexImage = 'Z1109_S16G1'
        IsControl = True
      end
      object BReduire: TToolbarButton97
        Left = 13
        Top = 7
        Width = 28
        Height = 27
        Hint = 'R'#233'duire la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'R'#233'duire'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BAgrandirClick
        GlobalIndexImage = 'Z0910_S16G1'
      end
    end
  end
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 631
    Height = 117
    ActivePage = PGenere
    Align = alTop
    TabOrder = 1
    object PComptes: TTabSheet
      Caption = 'Etapes'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 623
        Height = 89
        Align = alClient
      end
      object HLabel4: THLabel
        Left = 4
        Top = 38
        Width = 81
        Height = 13
        Caption = 'Compte &s'#233'lection'
        FocusControl = E_GENERAL
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HEtape: THLabel
        Left = 4
        Top = 12
        Width = 77
        Height = 13
        Caption = '&Etape r'#232'glement'
        FocusControl = ETAPE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HCpteGen: THLabel
        Left = 4
        Top = 64
        Width = 89
        Height = 13
        Caption = 'Compte &g'#233'n'#233'ration'
        FocusControl = GENEGENERATION
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object H_GENERAL: TLabel
        Left = 206
        Top = 37
        Width = 134
        Height = 13
        AutoSize = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object H_GENEGENERATION: TLabel
        Left = 206
        Top = 64
        Width = 134
        Height = 13
        AutoSize = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object BZOOMETAPE: TToolbarButton97
        Left = 311
        Top = 7
        Width = 23
        Height = 23
        Hint = 'Etape de r'#232'glement'
        Layout = blGlyphRight
        ParentShowHint = False
        ShowHint = True
        OnClick = BZOOMETAPEClick
        GlobalIndexImage = 'Z0008_S16G1'
      end
      object H_TOTDEVISE: TLabel
        Left = 344
        Top = 64
        Width = 80
        Height = 13
        Caption = 'Total '#233'ch'#233'ances'
      end
      object E_GENERAL: THCpteEdit
        Left = 101
        Top = 34
        Width = 101
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 17
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
        OnChange = E_GENERALChange
        OnExit = E_GENERALExit
        ZoomTable = tzGEncais
        Vide = True
        Bourre = False
        Libelle = H_GENERAL
        okLocate = False
        SynJoker = False
      end
      object ETAPE: THValComboBox
        Tag = 1
        Left = 101
        Top = 8
        Width = 207
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ItemHeight = 0
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
        OnChange = ETAPEChange
        TagDispatch = 0
      end
      object GENEGENERATION: THCpteEdit
        Left = 101
        Top = 60
        Width = 101
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 17
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
        OnChange = GENEGENERATIONChange
        OnExit = GENEGENERATIONExit
        ZoomTable = tzGeneral
        Vide = True
        Bourre = False
        Libelle = H_GENEGENERATION
        okLocate = False
        SynJoker = False
      end
      object FSelectAll: TCheckBox
        Left = 476
        Top = 12
        Width = 81
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Pr'#233'-s'#233'lection'
        TabOrder = 5
      end
      object rDetail: TRadioButton
        Left = 342
        Top = 13
        Width = 57
        Height = 14
        Caption = 'D'#233'tail'
        Checked = True
        TabOrder = 3
        TabStop = True
        OnClick = cTypeVisuChange
      end
      object rMasse: TRadioButton
        Left = 402
        Top = 13
        Width = 57
        Height = 14
        Caption = 'Masse'
        TabOrder = 4
        OnClick = cTypeVisuChange
      end
      object cBAP: TCheckBox
        Left = 344
        Top = 37
        Width = 79
        Height = 14
        Caption = '&Bon '#224' payer'
        TabOrder = 6
        OnClick = cBAPClick
      end
      object cModeleBAP: THValComboBox
        Left = 428
        Top = 34
        Width = 133
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 7
        TagDispatch = 0
        DataType = 'TTMODELEBAP'
      end
      object ERSOLDED: THNumEdit
        Tag = 2
        Left = 428
        Top = 60
        Width = 133
        Height = 21
        Color = clBtnFace
        Ctl3D = True
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Masks.PositiveMask = '#,##0.00'
        Debit = False
        NumericType = ntDC
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 8
        UseRounding = True
        Validate = False
      end
    end
    object PParam: TTabSheet
      Caption = 'Param'#232'tres'
      object Bevel2: TBevel
        Left = 0
        Top = 0
        Width = 623
        Height = 89
        Align = alClient
      end
      object HLabel7: THLabel
        Left = 296
        Top = 21
        Width = 91
        Height = 13
        Caption = '&Cat'#233'gorie paiement'
        FocusControl = MP_CATEGORIE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object Label14: TLabel
        Left = 296
        Top = 43
        Width = 88
        Height = 13
        Caption = '&Mode de paiement'
        FocusControl = E_MODEPAIE
      end
      object HLabel1: THLabel
        Left = 4
        Top = 43
        Width = 121
        Height = 13
        Caption = 'Guide pi'#232'ce de r'#232'glement'
        FocusControl = GUIDE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel17: THLabel
        Left = 4
        Top = 67
        Width = 33
        Height = 13
        Caption = 'Devise'
        FocusControl = E_DEVISE
      end
      object HLabel9: THLabel
        Left = 4
        Top = 21
        Width = 62
        Height = 13
        Caption = 'Type d'#39#233'tape'
        FocusControl = TYPEETAPE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object Label1: TLabel
        Left = 24
        Top = 4
        Width = 202
        Height = 13
        Caption = 'Param'#232'tres induits par l'#39#233'tape de r'#232'glement'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object Label2: TLabel
        Left = 364
        Top = 4
        Width = 80
        Height = 13
        Caption = 'Param'#232'tres libres'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object MP_CATEGORIE: THValComboBox
        Left = 396
        Top = 17
        Width = 177
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 3
        OnChange = MP_CATEGORIEChange
        TagDispatch = 0
        DataType = 'TTMODEPAIECAT'
      end
      object E_MODEPAIE: THValComboBox
        Left = 396
        Top = 40
        Width = 177
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        Vide = True
        DataType = 'TTMODEPAIE'
      end
      object GUIDE: THValComboBox
        Left = 136
        Top = 40
        Width = 145
        Height = 19
        Style = csSimple
        Ctl3D = False
        Enabled = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 1
        TagDispatch = 0
        DataType = 'TTGUIDEENCDEC'
      end
      object E_DEVISE: THValComboBox
        Tag = 1
        Left = 136
        Top = 63
        Width = 145
        Height = 21
        Style = csSimple
        Ctl3D = True
        Enabled = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 2
        OnChange = E_DEVISEChange
        TagDispatch = 0
        DataType = 'TTDEVISE'
      end
      object TYPEETAPE: THValComboBox
        Left = 136
        Top = 18
        Width = 145
        Height = 19
        Style = csSimple
        Ctl3D = False
        Enabled = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        DataType = 'TTTYPEETAPEREGLE'
      end
      object CodeAFB: THMultiValComboBox
        Left = 216
        Top = 51
        Width = 24
        Height = 21
        Color = clYellow
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 5
        Visible = False
        Abrege = False
        DataType = 'TTAFB'
        Complete = False
        OuInclusif = False
      end
    end
    object PGenere: TTabSheet
      Caption = 'G'#233'n'#233'ration'
      object Bevel5: TBevel
        Left = 0
        Top = 0
        Width = 623
        Height = 89
        Align = alClient
      end
      object H_MODEGENERE: TLabel
        Left = 12
        Top = 36
        Width = 95
        Height = 13
        Caption = 'Mode de &g'#233'n'#233'ration'
        FocusControl = MODEGENERE
      end
      object HDateReg: THLabel
        Left = 295
        Top = 11
        Width = 116
        Height = 13
        Caption = '&Date de comptabilisation'
        FocusControl = DATEGENERATION
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
      end
      object MODEGENERE: THValComboBox
        Left = 116
        Top = 32
        Width = 129
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 1
        TagDispatch = 0
        Vide = True
        DataType = 'TTMODEPAIE'
      end
      object DATEGENERATION: THCritMaskEdit
        Left = 425
        Top = 7
        Width = 105
        Height = 21
        Ctl3D = True
        Enabled = False
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        Text = '  /  /    '
        OnExit = DATEGENERATIONExit
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object CAutoEnreg: TCheckBox
        Left = 10
        Top = 62
        Width = 236
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Enregistrement &automatique des '#233'critures'
        TabOrder = 2
        Visible = False
      end
      object GLOBAL: TCheckBox
        Left = 10
        Top = 9
        Width = 235
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Compte g'#233'n'#233'ral de contrepartie globalis'#233
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
      end
      object DATEECHEANCE: THCritMaskEdit
        Left = 445
        Top = 32
        Width = 85
        Height = 21
        Ctl3D = True
        Enabled = False
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 5
        Text = '  /  /    '
        OnExit = DATEGENERATIONExit
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object CChoixDate: TCheckBox
        Left = 293
        Top = 34
        Width = 145
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Forcer date d'#39#233'ch'#233'ance'
        TabOrder = 4
        OnClick = CChoixDateClick
      end
      object GereAccept: TCheckBox
        Left = 293
        Top = 62
        Width = 237
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Ne prendre que les &effets accept'#233's'
        TabOrder = 6
        Visible = False
      end
    end
    object PGestionPDF: TTabSheet
      Caption = 'Spooler'
      ImageIndex = 7
      object Bevel7: TBevel
        Left = 0
        Top = 0
        Width = 623
        Height = 89
        Align = alClient
      end
      object LNBLIGNES: THLabel
        Left = 312
        Top = 15
        Width = 203
        Height = 13
        Caption = '&Nbre de lignes maximum par pi'#232'ce g'#233'n'#233'r'#233'e'
        Enabled = False
        FocusControl = NBLIGNES
      end
      object TREPSPOOLER: THLabel
        Left = 312
        Top = 41
        Width = 88
        Height = 13
        Caption = '&R'#233'pertoire Spooler'
        Enabled = False
        FocusControl = RepSpooler
      end
      object CAutoEnreg1: TCheckBox
        Left = 10
        Top = 13
        Width = 282
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Traitement gros volumes (plus de100 documents '#233'dit'#233's)'
        TabOrder = 0
        Visible = False
        OnClick = CAutoEnreg1Click
      end
      object NBLIGNES: THCritMaskEdit
        Left = 522
        Top = 11
        Width = 32
        Height = 21
        Ctl3D = True
        Enabled = False
        ParentCtl3D = False
        TabOrder = 1
        Text = '100'
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object Spooler: TCheckBox
        Left = 10
        Top = 39
        Width = 282
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Gestion du spooler d'#39#233'dition (format PDF)'
        TabOrder = 2
        OnClick = SpoolerClick
      end
      object RepSpooler: THCritMaskEdit
        Left = 409
        Top = 37
        Width = 144
        Height = 21
        CharCase = ecUpperCase
        Enabled = False
        MaxLength = 17
        TabOrder = 3
        TagDispatch = 0
        DataType = 'DIRECTORY'
        Operateur = Egal
        ElipsisButton = True
      end
      object XFichierSpooler: TCheckBox
        Left = 310
        Top = 63
        Width = 241
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Un fichier par document'
        Checked = True
        Enabled = False
        State = cbChecked
        TabOrder = 4
        OnClick = SpoolerClick
      end
    end
    object PCircuit: TTabSheet
      Caption = 'Circuit de validation'
      object Label3: TLabel
        Left = 277
        Top = 39
        Width = 17
        Height = 17
        Caption = #232
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -16
        Font.Name = 'Wingdings'
        Font.Style = []
        ParentFont = False
      end
      object Panel1: TPanel
        Left = 41
        Top = 24
        Width = 200
        Height = 41
        BevelOuter = bvLowered
        Caption = 'D'#233'caissement initiaux (aucune '#233'tape)'
        TabOrder = 3
      end
      object g1: TGroupBox
        Left = 20
        Top = 10
        Width = 243
        Height = 67
        Caption = 'Etape d'#39'&origine'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 0
        object HLabel6: THLabel
          Left = 11
          Top = 18
          Width = 74
          Height = 13
          Caption = '&Etape du circuit'
          FocusControl = E_SUIVDEC
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel12: THLabel
          Left = 11
          Top = 42
          Width = 51
          Height = 13
          Caption = '&Nom du lot'
          FocusControl = E_NOMLOT
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object bListeLots: TToolbarButton97
          Left = 212
          Top = 39
          Width = 16
          Height = 20
          Caption = '...'
          OnClick = bListeLotsClick
        end
        object E_SUIVDEC: THValComboBox
          Left = 91
          Top = 14
          Width = 137
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          TagDispatch = 0
          Vide = True
          DataType = 'CPCIRCUITDECAIS'
        end
        object E_NOMLOT: THCritMaskEdit
          Left = 91
          Top = 38
          Width = 122
          Height = 21
          CharCase = ecUpperCase
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 13
          ParentFont = False
          TabOrder = 1
          TagDispatch = 0
          Operateur = Egal
        end
      end
      object g2: TGroupBox
        Left = 307
        Top = 10
        Width = 243
        Height = 67
        Caption = 'Etape de &destination'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clTeal
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object HLabel13: THLabel
          Left = 11
          Top = 17
          Width = 74
          Height = 13
          Caption = '&Etape du circuit'
          FocusControl = CircuitDest
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLotDest: THLabel
          Left = 11
          Top = 42
          Width = 51
          Height = 13
          Caption = '&Nom du lot'
          FocusControl = LotDest
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object CircuitDest: THValComboBox
          Left = 91
          Top = 14
          Width = 137
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          TagDispatch = 0
          Vide = True
          DataType = 'CPCIRCUITDECAIS'
        end
        object LotDest: TEdit
          Left = 91
          Top = 38
          Width = 137
          Height = 21
          CharCase = ecUpperCase
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 13
          ParentFont = False
          TabOrder = 1
        end
      end
      object XX_WHERE1: TEdit
        Left = 273
        Top = 62
        Width = 23
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
        Visible = False
      end
    end
    object PEcritures: TTabSheet
      Caption = 'Ecritures'
      object Bevel3: TBevel
        Left = 0
        Top = 0
        Width = 623
        Height = 89
        Align = alClient
      end
      object TE_JOURNAL: THLabel
        Left = 10
        Top = 38
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = E_JOURNAL
      end
      object TE_NATUREPIECE: THLabel
        Left = 10
        Top = 66
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = E_NATUREPIECE
      end
      object TE_DATECOMPTABLE: THLabel
        Left = 255
        Top = 10
        Width = 103
        Height = 13
        Caption = '&Dates comptables  du'
        FocusControl = E_DATECOMPTABLE
      end
      object TE_DATECOMPTABLE2: THLabel
        Left = 468
        Top = 10
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATECOMPTABLE_
      end
      object TE_EXERCICE: THLabel
        Left = 10
        Top = 10
        Width = 41
        Height = 13
        Caption = 'E&xercice'
        FocusControl = E_EXERCICE
      end
      object TE_DATEECHEANCE: THLabel
        Left = 255
        Top = 38
        Width = 72
        Height = 13
        Caption = '&Ech'#233'ances du '
        FocusControl = E_DATEECHEANCE
      end
      object TE_DATEECHEANCE2: THLabel
        Left = 468
        Top = 38
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATEECHEANCE_
      end
      object E_EXERCICE: THValComboBox
        Left = 58
        Top = 6
        Width = 181
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 0
        OnChange = E_EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object E_JOURNAL: THValComboBox
        Left = 58
        Top = 34
        Width = 180
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 1
        TagDispatch = 0
        Vide = True
        DataType = 'TTJOURNAUX'
      end
      object E_NATUREPIECE: THValComboBox
        Left = 58
        Top = 62
        Width = 180
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATUREPIECE'
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 385
        Top = 6
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 3
        Text = '  /  /    '
        OnKeyPress = E_DATECOMPTABLEKeyPress
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATECOMPTABLE_: THCritMaskEdit
        Left = 489
        Top = 6
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 4
        Text = '  /  /    '
        OnKeyPress = E_DATECOMPTABLEKeyPress
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_NUMECHE: THCritMaskEdit
        Left = 120
        Top = 9
        Width = 12
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 9
        Text = '1'
        Visible = False
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object XX_WHERE: TEdit
        Left = 101
        Top = 9
        Width = 13
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 10
        Text = 'E_ECRANOUVEAU="N" or E_ECRANOUVEAU="H"'
        Visible = False
      end
      object MP_ENCAISSEMENT: THCritMaskEdit
        Left = 136
        Top = 9
        Width = 12
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 11
        Text = 'DEC'
        Visible = False
        TagDispatch = 0
        Operateur = Different
      end
      object E_QUALIFPIECE: THCritMaskEdit
        Left = 152
        Top = 9
        Width = 12
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 12
        Text = 'N'
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object E_ECHE: THCritMaskEdit
        Left = 170
        Top = 9
        Width = 12
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 13
        Text = 'X'
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object E_ETATLETTRAGE: THCritMaskEdit
        Left = 188
        Top = 9
        Width = 12
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 14
        Text = 'PL'
        Visible = False
        TagDispatch = 0
        Operateur = Inferieur
      end
      object E_TRESOLETTRE: THCritMaskEdit
        Left = 204
        Top = 9
        Width = 12
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 15
        Text = 'X'
        Visible = False
        TagDispatch = 0
        Operateur = Different
      end
      object XX_WHEREENC: TEdit
        Left = 84
        Top = 9
        Width = 13
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 16
        Visible = False
      end
      object cVoirCircuit: TCheckBox
        Left = 253
        Top = 65
        Width = 145
        Height = 14
        Alignment = taLeftJustify
        Caption = 'Avec circuit de &validation'
        TabOrder = 7
        OnClick = cVoirCircuitClick
      end
      object XX_WHERE3: TEdit
        Left = 102
        Top = 36
        Width = 23
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 17
        Visible = False
      end
      object E_DATEECHEANCE: THCritMaskEdit
        Left = 385
        Top = 34
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 5
        Text = '  /  /    '
        OnKeyPress = E_DATECOMPTABLEKeyPress
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATEECHEANCE_: THCritMaskEdit
        Left = 489
        Top = 34
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 6
        Text = '  /  /    '
        OnKeyPress = E_DATECOMPTABLEKeyPress
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object cFactCredit: TCheckBox
        Left = 439
        Top = 65
        Width = 127
        Height = 14
        Alignment = taLeftJustify
        Caption = '&Factures au cr'#233'dit'
        TabOrder = 8
      end
    end
    object PComplements: TTabSheet
      Caption = 'Compl'#233'ments'
      object Bevel4: TBevel
        Left = 0
        Top = 0
        Width = 623
        Height = 89
        Align = alClient
      end
      object TE_NUMEROPIECE: THLabel
        Left = 8
        Top = 39
        Width = 86
        Height = 13
        Caption = '&Num'#233'ros de pi'#232'ce'
        FocusControl = E_NUMEROPIECE
      end
      object HLabel2: THLabel
        Left = 184
        Top = 37
        Width = 6
        Height = 13
        Caption = #224
      end
      object HLabel3: THLabel
        Left = 6
        Top = 12
        Width = 79
        Height = 13
        Caption = 'Compte &auxiliaire'
        FocusControl = E_AUXILIAIRE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TE_DEBIT: TLabel
        Left = 284
        Top = 12
        Width = 80
        Height = 13
        Caption = 'Montant &d'#233'bit de'
        FocusControl = E_DEBIT
      end
      object TE_DEBIT_: TLabel
        Left = 471
        Top = 12
        Width = 6
        Height = 13
        Caption = #224
      end
      object TE_CREDIT: TLabel
        Left = 284
        Top = 39
        Width = 83
        Height = 13
        Caption = 'Montant &cr'#233'dit de'
        FocusControl = E_CREDIT
      end
      object TE_CREDIT_: TLabel
        Left = 471
        Top = 39
        Width = 6
        Height = 13
        Caption = #224
      end
      object iTiers: TToolbarButton97
        Tag = 1
        Left = 116
        Top = 45
        Width = 27
        Height = 27
        Hint = 'Ech'#233'ances du tiers'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        OnClick = BModifEcheClick
        GlobalIndexImage = 'Z0041_S16G2'
        IsControl = True
      end
      object iPiece: TToolbarButton97
        Tag = 1
        Left = 150
        Top = 45
        Width = 27
        Height = 27
        Hint = 'Modification de l'#39#233'ch'#233'ance'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        OnClick = BModifEcheClick
        GlobalIndexImage = 'Z0357_S16G2'
        IsControl = True
      end
      object H_OrdreTri: THLabel
        Left = 8
        Top = 66
        Width = 52
        Height = 13
        Caption = 'Ordre de &tri'
        FocusControl = OrdreTri
      end
      object TE_ETABLISSEMENT: THLabel
        Left = 284
        Top = 66
        Width = 65
        Height = 13
        Caption = '&Etablissement'
        FocusControl = E_ETABLISSEMENT
      end
      object E_NUMEROPIECE: THCritMaskEdit
        Left = 104
        Top = 35
        Width = 77
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 1
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_NUMEROPIECE_: THCritMaskEdit
        Left = 196
        Top = 35
        Width = 77
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_AUXILIAIRE: THCpteEdit
        Left = 104
        Top = 8
        Width = 169
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 0
        ZoomTable = tzTlettrable
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object E_DEBIT: THCritMaskEdit
        Left = 375
        Top = 8
        Width = 90
        Height = 21
        TabOrder = 4
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_DEBIT_: THCritMaskEdit
        Left = 485
        Top = 8
        Width = 90
        Height = 21
        TabOrder = 5
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_CREDIT: THCritMaskEdit
        Left = 375
        Top = 35
        Width = 90
        Height = 21
        TabOrder = 6
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_CREDIT_: THCritMaskEdit
        Left = 485
        Top = 35
        Width = 90
        Height = 21
        TabOrder = 7
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object OrdreTri: THValComboBox
        Left = 104
        Top = 62
        Width = 170
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 3
        Items.Strings = (
          'G'#233'n'#233'ral / Date / Auxiliaire'
          'Auxiliaire / G'#233'n'#233'ral / Date'
          'Date / Auxiliaire / G'#233'n'#233'ral'
          'Ech'#233'ance / Auxiliaire / G'#233'n'#233'ral')
        TagDispatch = 0
        Vide = True
        Values.Strings = (
          'GDA'
          'AGD'
          'DAG'
          'EAG')
      end
      object E_ETABLISSEMENT: THValComboBox
        Left = 375
        Top = 62
        Width = 200
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 8
        TagDispatch = 0
        Vide = True
        DataType = 'TTETABLISSEMENT'
      end
    end
    object PLibres: TTabSheet
      Caption = 'Tables libres'
      object Bevel6: TBevel
        Left = 0
        Top = 0
        Width = 623
        Height = 89
        Align = alClient
      end
      object TT_TABLE0: THLabel
        Left = 7
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE0'
      end
      object TT_TABLE1: THLabel
        Left = 122
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE1'
      end
      object TT_TABLE2: THLabel
        Left = 237
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE2'
      end
      object TT_TABLE3: THLabel
        Left = 351
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE3'
      end
      object TT_TABLE4: THLabel
        Left = 463
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE4'
      end
      object TT_TABLE5: THLabel
        Left = 7
        Top = 44
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE5'
      end
      object TT_TABLE6: THLabel
        Left = 122
        Top = 44
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE6'
      end
      object TT_TABLE7: THLabel
        Left = 237
        Top = 44
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE7'
      end
      object TT_TABLE8: THLabel
        Left = 351
        Top = 44
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE8'
      end
      object TT_TABLE9: THLabel
        Left = 463
        Top = 44
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE9'
      end
      object T_TABLE4: THCpteEdit
        Left = 463
        Top = 17
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 0
        ZoomTable = tzNatTiers4
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE3: THCpteEdit
        Left = 351
        Top = 17
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 1
        ZoomTable = tzNatTiers3
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE2: THCpteEdit
        Left = 237
        Top = 17
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 2
        ZoomTable = tzNatTiers2
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE1: THCpteEdit
        Left = 122
        Top = 17
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 3
        ZoomTable = tzNatTiers1
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE0: THCpteEdit
        Left = 7
        Top = 17
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 4
        ZoomTable = tzNatTiers0
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE5: THCpteEdit
        Left = 7
        Top = 59
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 5
        ZoomTable = tzNatTiers5
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE6: THCpteEdit
        Left = 122
        Top = 59
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 6
        ZoomTable = tzNatTiers6
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE7: THCpteEdit
        Left = 237
        Top = 59
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 7
        ZoomTable = tzNatTiers7
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE8: THCpteEdit
        Left = 351
        Top = 59
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 8
        ZoomTable = tzNatTiers8
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE9: THCpteEdit
        Left = 463
        Top = 59
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 9
        ZoomTable = tzNatTiers9
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 117
    Width = 631
    Height = 35
    AllowDrag = False
    object PFiltres: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 631
      ClientAreaHeight = 31
      ClientAreaWidth = 631
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BFiltre: TToolbarButton97
        Left = 5
        Top = 5
        Width = 54
        Height = 21
        Hint = 'Menu filtre'
        DropdownArrow = True
        Caption = '&Filtres'
        Layout = blGlyphRight
        ParentShowHint = False
        ShowHint = True
      end
      object BChercher: TToolbarButton97
        Left = 554
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
        Left = 63
        Top = 5
        Width = 486
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
      end
    end
  end
  object G: THGrid
    Tag = 1
    Left = 0
    Top = 152
    Width = 631
    Height = 185
    Align = alClient
    Ctl3D = True
    DefaultColWidth = 50
    DefaultRowHeight = 18
    FixedCols = 0
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
    ParentCtl3D = False
    TabOrder = 3
    OnDblClick = GDblClick
    OnMouseUp = GMouseUp
    SortedCol = -1
    Couleur = True
    MultiSelect = True
    TitleBold = False
    TitleCenter = True
    OnRowEnter = GRowEnter
    ColCombo = 0
    SortEnabled = True
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = 13224395
    ColWidths = (
      50
      50
      50
      50
      50)
  end
  object QEnc: THQuery
    AutoCalcFields = False
    LockType = ltReadOnly
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    Parameters = <>
    dataBaseName = 'SOC'
    MAJAuto = False
    Distinct = False
    Left = 54
    Top = 285
  end
  object FindMvt: TFindDialog
    OnFind = FindMvtFind
    Left = 97
    Top = 285
  end
  object POPS: TPopupMenu
    OnPopup = POPSPopup
    Left = 141
    Top = 285
  end
  object POPF: TPopupMenu
    Left = 184
    Top = 288
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
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 228
    Top = 285
  end
  object HDiv: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Compte issu du mode de paiement'
      'G'#233'n'#233'ration des encaissements non effectu'#233'e !'
      
        'ATTENTION. Certains mouvements en cours de traitement par un aut' +
        're utilisateur n'#39'ont pas g'#233'n'#233'r'#233' l'#39'op'#233'ration'
      'Date r'#232'glement'
      'Date virement'
      'Date effective de virement'
      'Date pr'#233'l'#232'v.'
      'Date effective de pr'#233'l'#232'vement'
      'Euro'
      'Factures au d'#233'bit'
      '10;')
    Left = 271
    Top = 285
  end
  object HMEnc: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?caption?;Vous devez choisir un compte g'#233'n'#233'ral.;W;O;O;O;'
      'Encaissements'
      'D'#233'caissements'
      '3;'
      
        '4;?caption?;Confirmez-vous l'#39'encaissement des '#233'critures s'#233'lectio' +
        'nn'#233'es ?;Q;YNC;N;N;'
      
        '5;?caption?;Confirmez-vous le d'#233'caissement des '#233'critures s'#233'lecti' +
        'onn'#233'es ?;Q;YNC;N;N;'
      '6;?caption?;Vous n'#39'avez s'#233'lectionn'#233' aucune '#233'criture.;E;O;O;O;'
      
        '7;?caption?;Vous devez renseigner ou param'#233'trer un guide de r'#232'gl' +
        'ement.;W;O;O;O;'
      
        '8;?caption?;Vous devez renseigner un compte g'#233'n'#233'ral de s'#233'lection' +
        ' valide.;W;O;O;O;'
      '9;?caption?;Vous devez renseigner une date valide.;W;O;O;O;'
      
        '10;?caption?;La date de r'#232'glement que vous avez renseign'#233'e est s' +
        'ur un exercice non ouvert.;W;O;O;O;'
      
        '11;?caption?;La date de r'#232'glement que vous avez renseign'#233'e est s' +
        'ur un exercice non ouvert.;W;O;O;O;'
      
        '12;?caption?;La date de r'#232'glement que vous avez renseign'#233'e est a' +
        'nt'#233'rieure '#224' la cl'#244'ture provisoire.;W;O;O;O;'
      
        '13;?caption?;La date de r'#232'glement que vous avez renseign'#233'e est a' +
        'nt'#233'rieure '#224' la cl'#244'ture d'#233'finitive.;W;O;O;O;'
      
        '14;?caption?;Vous devez lancer une recherche d'#39#233'critures au pr'#233'a' +
        'lable.;W;O;O;O;'
      
        '15;?caption?;Votre s'#233'lection contient des mouvements sur des dev' +
        'ises diff'#233'rentes.;W;O;O;O;'
      
        '16;?caption?;Vous devez renseigner une cat'#233'gorie de paiement.;W;' +
        'O;O;O;'
      
        '17;?caption?;Votre guide de r'#232'glement doit contenir deux lignes.' +
        ';W;O;O;O;'
      '18;LIBRE LIBRE'
      
        '19;?caption?;Le guide de r'#232'glement ne correspond '#224' aucune '#233'tape ' +
        'coh'#233'rente.;W;O;O;O;'
      
        '20;?caption?;La devise du guide diff'#232're de celle des mouvements.' +
        ';W;O;O;O;'
      
        '21;?caption?;Le compte de banque ne doit pas '#234'tre collectif.;W;O' +
        ';O;O;'
      
        '22;?caption?;Les modes de paiement induisent des globalisations ' +
        'diff'#233'rentes.;W;O;O;O;'
      
        '23;?caption?;Des modes de paiement diff'#233'rents ne peuvent '#234'tre gl' +
        'obalis'#233's pour un m'#234'me tiers.;W;O;O;O;'
      
        '24;?caption?;Les modes de paiement ou les '#233'ch'#233'ances diff'#232'rent. V' +
        'oulez-vous choisir et rendre identiques ces informations  ?;Q;YN' +
        ';Y;Y;'
      
        '25;?caption?;Vous devez renseigner un compte g'#233'n'#233'ral de g'#233'n'#233'rati' +
        'on;W;O;O;O;'
      '26;?caption?;Vous devez choisir une '#233'tape d'#39'origine;W;O;O;O;'
      '27;?caption?;Vous devez choisir une '#233'tape de r'#232'glement;W;O;O;O;'
      
        '28;?caption?;Les comptes de s'#233'lection et de g'#233'n'#233'ration doivent '#234 +
        'tre diff'#233'rents.;W;O;O;O;'
      
        '29;?caption?;Votre choix est impossible : le compte de s'#233'lection' +
        ' ou de g'#233'n'#233'ration est interdit.;W;O;O;O;'
      
        '30;?caption?;Les modes de paiement ou les '#233'ch'#233'ances diff'#232'rent. C' +
        'es informations seront perdues en d'#233'tail. Confirmez-vous la g'#233'n'#233 +
        'ration ?;Q;YN;Y;Y;'
      
        '31;?caption?;Certaines dates d'#39#233'ch'#233'ance sont en dehors de la pla' +
        'ge de saisie autoris'#233'e. Vous devez modifier votre date de g'#233'n'#233'ra' +
        'tion de r'#232'glement ou la nouvelle date d'#39#233'ch'#233'ance;W;O;O;O;'
      
        '32;?caption?;Vos '#233'ch'#233'ances en devise portent des taux diff'#233'rents' +
        '. La pi'#232'ce ne peut pas '#234'tre g'#233'n'#233'r'#233'e.;W;O;O;O;'
      'Zoom tiers'
      'Zoom pi'#232'ce'
      
        '35;?caption?;Vous devez renseigner un nom de lot destination.;W;' +
        'O;O;O;'
      'Attention : validation non effectu'#233'e !'
      
        '37;?caption?;Vous devez renseigner un nom de lot d'#39'origine.;W;O;' +
        'O;O;'
      
        '38;?caption?;Les '#233'critures s'#233'lectionn'#233'es ont '#233't'#233' trait'#233'es correc' +
        'tement.;I;O;O;O;'
      'Affectation pr'#233'visionnelle des banques'
      'Voir les affectations des banques'
      'Choix d'#39'un nom de lot'
      '&Etape de r'#232'glement'
      '&Etape de s'#233'lection'
      'G'#233'n'#233'ration de r'#232'glements - Lettres-ch'#232'que'
      'Lettres-ch'#232'que'
      
        '46;?caption?;Confirmez-vous la g'#233'n'#233'ration et l'#39#233'mission des ch'#232'q' +
        'ues ?;Q;YNC;N;N;'
      'Lettres-BOR'
      'G'#233'n'#233'ration de r'#232'glements - Lettres-BOR'
      
        '49;?caption?;Confirmez-vous la g'#233'n'#233'ration et l'#39#233'mission des BOR ' +
        '?;Q;YNC;N;N;'
      
        '50;?caption?;Confirmez-vous la g'#233'n'#233'ration et l'#39#233'mission des trai' +
        'tes ?;Q;YNC;N;N;'
      'G'#233'n'#233'ration de r'#232'glements - Lettres-traite'
      
        '52;?caption?;Certains RIB ne sont pas renseign'#233's. La pi'#232'ce ne pe' +
        'ut pas '#234'tre g'#233'n'#233'r'#233'e.;W;O;O;O;'
      
        '53;?caption?;Vous avez affect'#233' pr'#233'visionnellement des banques di' +
        'ff'#233'rentes. Confirmez-vous la g'#233'n'#233'ration ?;Q;YN;Y;Y;'
      '54;?caption?;Confirmez-vous la d'#233'finition du lot;Q;YN;Y;Y;'
      
        '55;?caption?;Le journal de g'#233'n'#233'ration est ferm'#233'. La pi'#232'ce ne peu' +
        't pas '#234'tre g'#233'n'#233'r'#233'e.;W;O;O;O;')
    Left = 313
    Top = 285
  end
end

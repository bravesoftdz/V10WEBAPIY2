object FAffBQEED: TFAffBQEED
  Left = 164
  Top = 130
  Width = 630
  Height = 410
  HelpContext = 7805000
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Affectation des banques'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Dock: TDock97
    Left = 0
    Top = 336
    Width = 614
    Height = 35
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 618
      Caption = 'Actions'
      ClientAreaHeight = 31
      ClientAreaWidth = 618
      DockPos = 0
      TabOrder = 0
      object BCalcul: TToolbarButton97
        Tag = 1
        Left = 489
        Top = 2
        Width = 27
        Height = 27
        Hint = 'Calcul des r'#233'partitions'
        Caption = 'Calculer'
        DisplayMode = dmGlyphOnly
        Flat = False
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
        OnClick = BCalculClick
        GlobalIndexImage = 'Z0184_S16G1'
        IsControl = True
      end
      object BReduire: TToolbarButton97
        Left = 4
        Top = 3
        Width = 28
        Height = 27
        Hint = 'R'#233'duire la liste'
        Caption = 'R'#233'duire'
        DisplayMode = dmGlyphOnly
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
        Caption = 'Agrandir'
        DisplayMode = dmGlyphOnly
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
      object BRechercher: TToolbarButton97
        Left = 36
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Rechercher dans la liste'
        Caption = 'Chercher'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BRechercherClick
        GlobalIndexImage = 'Z0077_S16G1'
      end
      object BImprimer: TToolbarButton97
        Left = 456
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Caption = 'Imprimer'
        DisplayMode = dmGlyphOnly
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
      object BAnnuler: TToolbarButton97
        Left = 556
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Caption = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAnnulerClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 588
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        Caption = 'Aide'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
      object bExport: TToolbarButton97
        Left = 424
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Exporter la liste'
        Caption = 'Exporter'
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        OnClick = bExportClick
        GlobalIndexImage = 'Z0724_S16G1'
      end
      object BSelectMvt: TToolbarButton97
        Tag = 1
        Left = 68
        Top = 3
        Width = 28
        Height = 27
        Hint = 'S'#233'lectionner les mouvements'
        Caption = 'S'#233'lection'
        DisplayMode = dmGlyphOnly
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
        OnClick = BSelectMvtClick
        GlobalIndexImage = 'Z0047_S16G1'
        IsControl = True
      end
      object BOuvrir: TToolbarButton97
        Left = 522
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Valider le lot'
        Caption = 'Valider'
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        NumGlyphs = 2
        OnClick = BOuvrirClick
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object TFType: THLabel
        Left = 223
        Top = 10
        Width = 26
        Height = 13
        Caption = ' Grille'
        FocusControl = CVType
      end
      object BSauveVentil: TToolbarButton97
        Left = 375
        Top = 6
        Width = 21
        Height = 21
        Hint = 'Enregistrer comme ventilation type sur l'#39'axe en cours'
        ParentShowHint = False
        ShowHint = True
        OnClick = BSauveVentilClick
        GlobalIndexImage = 'Z0393_S16G1'
      end
      object CVType: THValComboBox
        Left = 254
        Top = 6
        Width = 121
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = CVTypeChange
        TagDispatch = 0
        DataType = 'CPVENTILAFFBQE'
      end
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 111
    Width = 614
    Height = 35
    AllowDrag = False
    object PFiltres: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 618
      ClientAreaHeight = 31
      ClientAreaWidth = 618
      DockPos = 0
      TabOrder = 0
      object BFiltre: TToolbarButton97
        Left = 5
        Top = 5
        Width = 56
        Height = 21
        Hint = 'Menu filtre'
        Caption = '&Filtres'
        DropdownArrow = True
        DropdownMenu = POPF
        Layout = blGlyphRight
        ParentShowHint = False
        ShowHint = True
      end
      object BCherche: TToolbarButton97
        Left = 584
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Appliquer crit'#232'res'
        DisplayMode = dmGlyphOnly
        ParentShowHint = False
        ShowHint = True
        OnClick = BChercheClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object FFiltres: THValComboBox
        Left = 67
        Top = 5
        Width = 510
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = FFiltresChange
        TagDispatch = 0
      end
    end
  end
  object FListe: THGrid
    Left = 0
    Top = 146
    Width = 614
    Height = 173
    Align = alClient
    ColCount = 8
    DefaultColWidth = 500
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goEditing, goTabs, goAlwaysShowEditor]
    TabOrder = 2
    OnDblClick = FListeDblClick
    OnDragDrop = FListeDragDrop
    OnDragOver = FListeDragOver
    OnEnter = FListeEnter
    OnKeyDown = FListeKeyDown
    OnKeyPress = FListeKeyPress
    OnTopLeftChanged = FListeTopLeftChanged
    SortedCol = -1
    Titres.Strings = (
      'Compte'
      'Banque'
      'Solde '
      'Solde calcul'#233
      '% saisi'
      'Montant saisi'
      '% calcul'#233
      'Montant calcul'#233)
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    OnRowEnter = FListeRowEnter
    OnCellExit = FListeCellExit
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = clSilver
    ColWidths = (
      63
      123
      78
      69
      57
      73
      45
      86)
  end
  object SM: THNumEdit
    Left = 292
    Top = 172
    Width = 50
    Height = 21
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
    Decimals = 2
    Digits = 12
    Masks.PositiveMask = '#,##0.00'
    Debit = False
    NumericType = ntDC
    UseRounding = True
    Validate = False
  end
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 614
    Height = 111
    ActivePage = PCritere
    Align = alTop
    TabOrder = 4
    TabWidth = 85
    object PCritere: TTabSheet
      Caption = 'Standards'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 606
        Height = 83
        Align = alClient
      end
      object HLabel3: THLabel
        Left = 98
        Top = 17
        Width = 51
        Height = 13
        Caption = '&Nom lot de'
        FocusControl = E_NOMLOT
      end
      object HLabel5: THLabel
        Left = 338
        Top = 17
        Width = 6
        Height = 13
        Caption = #224
      end
      object HLabel6: THLabel
        Left = 98
        Top = 57
        Width = 57
        Height = 13
        Caption = '&Banques de'
        FocusControl = E_BANQUEPREVI
      end
      object HLabel8: THLabel
        Left = 338
        Top = 57
        Width = 6
        Height = 13
        Caption = #224
      end
      object E_NOMLOT: THCritMaskEdit
        Left = 182
        Top = 13
        Width = 131
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        DataType = 'CPNOMLOT'
        Operateur = Superieur
        ElipsisButton = True
      end
      object E_NOMLOT_: THCritMaskEdit
        Left = 376
        Top = 13
        Width = 131
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 1
        TagDispatch = 0
        DataType = 'CPNOMLOT'
        Operateur = Inferieur
        ElipsisButton = True
      end
      object E_BANQUEPREVI: THCritMaskEdit
        Left = 182
        Top = 53
        Width = 131
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        DataType = 'TZGBANQUE'
        Operateur = Superieur
        ElipsisButton = True
      end
      object E_BANQUEPREVI_: THCritMaskEdit
        Left = 376
        Top = 53
        Width = 131
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        DataType = 'TZGBANQUE'
        Operateur = Inferieur
        ElipsisButton = True
      end
    end
    object PComplement: TTabSheet
      Caption = 'Compl'#233'ments'
      object Bevel2: TBevel
        Left = 0
        Top = 0
        Width = 614
        Height = 83
        Align = alClient
      end
    end
  end
  object PVentil: TPanel
    Left = 140
    Top = 167
    Width = 321
    Height = 65
    TabOrder = 5
    Visible = False
    object H_CODEVENTIL: THLabel
      Left = 8
      Top = 40
      Width = 25
      Height = 13
      Caption = 'Code'
    end
    object HLabel2: THLabel
      Left = 92
      Top = 40
      Width = 30
      Height = 13
      Caption = 'Libell'#233
    end
    object H_TITREVENTIl: TLabel
      Left = 11
      Top = 12
      Width = 299
      Height = 16
      Alignment = taCenter
      Caption = 'Cr'#233'ation d'#39'une nouvelle grille de r'#233'partition'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Y_CODEVENTIL: TEdit
      Left = 40
      Top = 36
      Width = 37
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 3
      TabOrder = 0
    end
    object Y_LIBELLEVENTIL: TEdit
      Left = 128
      Top = 36
      Width = 149
      Height = 21
      MaxLength = 35
      TabOrder = 1
    end
    object BNewVentil: THBitBtn
      Tag = 1
      Left = 285
      Top = 33
      Width = 28
      Height = 27
      Hint = 'Valider la cr'#233'ation'
      TabOrder = 2
      OnClick = BNewVentilClick
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      GlobalIndexImage = 'Z0003_S16G2'
      IsControl = True
    end
    object PFenCouverture: TPanel
      Left = 1
      Top = 1
      Width = 319
      Height = 9
      Align = alTop
      BevelInner = bvLowered
      BevelOuter = bvNone
      Color = clTeal
      Enabled = False
      TabOrder = 3
    end
  end
  object PCumul: TPanel
    Left = 0
    Top = 319
    Width = 614
    Height = 17
    Align = alBottom
    Alignment = taLeftJustify
    BevelOuter = bvNone
    BorderStyle = bsSingle
    Caption = '  Totaux'
    Color = clWindow
    Ctl3D = False
    Enabled = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentCtl3D = False
    ParentFont = False
    TabOrder = 6
    object FTotPSaisi: THNumEdit
      Left = 149
      Top = -1
      Width = 121
      Height = 19
      Enabled = False
      TabOrder = 0
      Decimals = 2
      Digits = 12
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      NumericType = ntPercentage
      UseRounding = True
      Validate = False
    end
    object FTotMSaisi: THNumEdit
      Left = 273
      Top = -1
      Width = 121
      Height = 19
      Enabled = False
      TabOrder = 1
      Decimals = 2
      Digits = 12
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      UseRounding = True
      Validate = False
    end
    object FTotPCalcul: THNumEdit
      Left = 395
      Top = -1
      Width = 121
      Height = 19
      Enabled = False
      TabOrder = 2
      Decimals = 2
      Digits = 12
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      NumericType = ntPercentage
      UseRounding = True
      Validate = False
    end
    object FTotMCalcul: THNumEdit
      Left = 503
      Top = -1
      Width = 121
      Height = 19
      Enabled = False
      TabOrder = 3
      Decimals = 2
      Digits = 12
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      UseRounding = True
      Validate = False
    end
  end
  object TotRepart: THNumEdit
    Left = 472
    Top = 0
    Width = 121
    Height = 21
    Enabled = False
    TabOrder = 7
    Visible = False
    Decimals = 2
    Digits = 12
    Masks.PositiveMask = '#,##0.00'
    Debit = False
    NumericType = ntDC
    UseRounding = True
    Validate = False
  end
  object WFListe2: TToolWindow97
    Left = 112
    Top = 35
    ClientHeight = 244
    ClientWidth = 563
    Caption = 'Liste des '#233'critures'
    ClientAreaHeight = 244
    ClientAreaWidth = 563
    DockableTo = []
    DockPos = 8
    HideWhenInactive = False
    Resizable = False
    TabOrder = 8
    Visible = False
    object PFListe2: TPanel
      Left = 0
      Top = 210
      Width = 563
      Height = 34
      Align = alBottom
      TabOrder = 0
      object ToolbarButton972: TToolbarButton97
        Left = 499
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Imprimer'
        Caption = 'Imprimer'
        DisplayMode = dmGlyphOnly
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
      object ToolbarButton973: TToolbarButton97
        Left = 532
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Caption = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAnnulerClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object bSelectAll: TToolbarButton97
        Left = 5
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Tout s'#233'lectionner'
        Caption = 'S'#233'lectionner'
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = bSelectAllClick
        GlobalIndexImage = 'Z0205_S16G1'
      end
      object BSelectBqe: TToolbarButton97
        Left = 70
        Top = 4
        Width = 28
        Height = 27
        Hint = 'S'#233'lectionner les '#233'ch'#233'ances de la banque courante'
        Caption = 'S'#233'lectionner'
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = BSelectBqeClick
        GlobalIndexImage = 'Z0210_S16G1'
      end
      object bPoubelle: TToolbarButton97
        Left = 133
        Top = 4
        Width = 28
        Height = 27
        Hint = 'RAZ banque pr'#233'visionnelle'
        Caption = 'S'#233'lectionner'
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = bPoubelleClick
        GlobalIndexImage = 'Z0204_S16G1'
      end
      object BDeselectAll: TToolbarButton97
        Left = 38
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Tout d'#233'-s'#233'lectionner'
        Caption = 'S'#233'lectionner'
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BDeselectAllClick
        GlobalIndexImage = 'Z0078_S16G2'
      end
      object BDeselectBqe: TToolbarButton97
        Left = 101
        Top = 4
        Width = 28
        Height = 27
        Hint = 'D'#233'-selectionner les '#233'ch'#233'ances de la banque courante'
        Caption = 'S'#233'lectionner'
        GroupIndex = 1
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        OnClick = BDeselectBqeClick
        GlobalIndexImage = 'Z0192_S16G1'
      end
    end
    object FListe2: THGrid
      Left = 0
      Top = 0
      Width = 563
      Height = 210
      Align = alClient
      ColCount = 10
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 10
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
      TabOrder = 1
      OnDblClick = FListe2DblClick
      OnDragOver = FListe2DragOver
      OnEndDrag = FListe2EndDrag
      OnMouseMove = FListe2MouseMove
      SortedCol = -1
      Titres.Strings = (
        'E_BANQUEPREVI'
        'E_JOURNAL'
        'E_GENERAL'
        'E_AUXILIAIRE'
        'E_DATECOMPTABLE'
        'E_DATEECHEANCE'
        'E_NUMEROPIECE'
        'E_DEBIT'
        'E_CREDIT'
        'IND')
      Couleur = False
      MultiSelect = True
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = clSilver
      ColWidths = (
        73
        59
        71
        69
        64
        64
        64
        71
        72
        17)
    end
  end
  object FindDialog: TFindDialog
    OnFind = FindDialogFind
    Left = 411
    Top = 288
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?Caption?;Vous devez saisir un code et un libell'#233';W;O;O;O;'
      '1;?Caption?;Le code ou le libell'#233' existe d'#233'j'#224'.;W;O;O;O;'
      
        '2;?Caption?;Confirmez-vous l'#39'arr'#234't du traitement en cours ?;Q;YN' +
        ';N;N;'
      
        '3;?Caption?;Confirmez-vous le traitement d'#39'affection des banques' +
        ' ?;Q;YN;Y;N;'
      
        '4;?Caption?;Voulez-vous annuler les affectations d'#233'j'#224' effectu'#233'es' +
        ' ?;Q;YN;Y;N;'
      '5;?Caption?;Confirmez-vous la validation du lot ?;Q;YN;Y;N;'
      '6;?Caption?;Vous n'#39'avez pas renseign'#233' de code lot.;W;O;O;O;'
      'Attention : Validation non effectu'#233'e.'
      
        '8;?Caption?;Confirmez-vous l'#39'affectation des '#233'ch'#233'ances s'#233'lection' +
        'n'#233'es?;Q;YN;Y;N;'
      
        '9;?Caption?;Confirmez-vous l'#39'annulation de l'#39'affectation des '#233'ch' +
        #233'ances s'#233'lectionn'#233'es?;Q;YN;Y;N;'
      
        '10;?Caption?;Attention ! Certaines '#233'ch'#233'ances n'#39'ont pas '#233't'#233' prise' +
        's en comptes (Total < 100%);W;O;O;O;'
      'Cr'#233'ation d'#39'un lot'
      'Modification d'#39'un lot'
      '13;')
    Left = 588
    Top = 212
  end
  object HMTrad: THSystemMenu
    LockedCtrls.Strings = (
      'WFListe2'
      'FListe2'
      'PFListe2')
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 375
    Top = 268
  end
  object QBud: TQuery
    DatabaseName = 'SOC'
    SQL.Strings = (
      'Select Sum(BE_DEBIT) as Deb, Sum(BE_CREDIT) as Cred From BUDECR'
      'Where BE_BUDJAL=:Jal And BE_EXERCICE=:Exo And '
      'BE_DATECOMPTABLE>=:DDeb And BE_DATECOMPTABLE<=:DFin'
      'And BE_BUDGENE=:Cpte And BE_ETABLISSEMENT=:Etab'
      'And BE_QUALIFPIECE=:Qual')
    Left = 308
    Top = 304
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Jal'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Exo'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'DDeb'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'DFin'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Cpte'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Etab'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'Qual'
        ParamType = ptUnknown
      end>
  end
  object POPF: TPopupMenu
    OnPopup = POPFPopup
    Left = 532
    Top = 264
    object BCreerFiltre: TMenuItem
      Caption = '&Cr'#233'er un filtre'
      OnClick = BCreerFiltreClick
    end
    object BSaveFiltre: TMenuItem
      Caption = '&Enregistrer le filtre'
      OnClick = BSaveFiltreClick
    end
    object BDelFiltre: TMenuItem
      Caption = '&Supprimer le filtre'
      OnClick = BDelFiltreClick
    end
    object BRenFiltre: TMenuItem
      Caption = '&Renommer le filtre'
      OnClick = BRenFiltreClick
    end
    object BNouvRech: TMenuItem
      Caption = '&Nouvelle recherche'
      OnClick = BNouvRechClick
    end
  end
  object SD: TSaveDialog
    DefaultExt = 'XLS'
    Filter = 
      'Fichier Texte (*.txt)|*.txt|Fichier Excel (*.xls)|*.xls|Fichier ' +
      'Ascii (*.asc)|*.asc|Fichier Lotus (*.wks)|*.wks'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofPathMustExist, ofNoReadOnlyReturn, ofNoLongNames]
    Title = 'Export'
    Left = 436
    Top = 269
  end
end

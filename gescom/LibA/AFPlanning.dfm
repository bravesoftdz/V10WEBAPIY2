object AF_Planning: TAF_Planning
  Left = 150
  Top = 241
  Width = 991
  Height = 478
  Caption = 'Planning'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  WindowState = wsMaximized
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PPlanningJour: THPanel
    Left = 0
    Top = 0
    Width = 975
    Height = 312
    Align = alClient
    FullRepaint = False
    TabOrder = 0
    BackGroundEffect = bdFlat
    ColorShadow = clWindowText
    ColorStart = clBtnFace
    TextEffect = tenone
  end
  object PA_SELECTION: TPanel
    Left = 0
    Top = 346
    Width = 975
    Height = 23
    Align = alBottom
    Caption = ' '
    TabOrder = 1
    object LA_SELECTION: TLabel
      Left = 8
      Top = 4
      Width = 76
      Height = 13
      Caption = 'LA_Selection'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object LA_Quantite: TLabel
      Left = 929
      Top = 4
      Width = 71
      Height = 13
      Alignment = taRightJustify
      Caption = 'LA_Quantite'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object PA_COMMENTAIRE: TPanel
    Left = 0
    Top = 369
    Width = 975
    Height = 36
    Align = alBottom
    Caption = ' '
    TabOrder = 2
    object LA_Commentaire: TLabel
      Left = 7
      Top = 4
      Width = 95
      Height = 13
      Caption = 'LA_Commentaire'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object LA_BASE: TLabel
      Left = 947
      Top = 4
      Width = 55
      Height = 13
      Alignment = taRightJustify
      Caption = 'LA_BASE'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object PA_Quick: TPanel
    Left = 0
    Top = 312
    Width = 975
    Height = 34
    Align = alBottom
    Caption = ' '
    TabOrder = 3
    object CB_RAP: THCheckbox
      Left = 285
      Top = 8
      Width = 52
      Height = 17
      Caption = 'RAP'
      Checked = True
      State = cbChecked
      TabOrder = 5
      OnClick = CB_RAPClick
    end
    object ED_TACHE: THCritMaskEdit
      Left = 372
      Top = 6
      Width = 163
      Height = 21
      Enabled = False
      TabOrder = 7
      TagDispatch = 0
    end
    object ED_AFF1: THCritMaskEdit
      Left = 41
      Top = 6
      Width = 55
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 1
      OnChange = ED_AFF1Change
      TagDispatch = 0
    end
    object ED_AFF2: THCritMaskEdit
      Left = 97
      Top = 6
      Width = 55
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 2
      OnChange = ED_AFF2Change
      TagDispatch = 0
    end
    object ED_AFF3: THCritMaskEdit
      Left = 153
      Top = 6
      Width = 35
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 3
      OnChange = ED_AFF3Change
      TagDispatch = 0
    end
    object ED_AVE: THCritMaskEdit
      Left = 189
      Top = 6
      Width = 25
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 4
      OnExit = ED_AVEExit
      TagDispatch = 0
    end
    object BCHERCHEAFF: THBitBtn
      Left = 8
      Top = 4
      Width = 27
      Height = 27
      TabOrder = 0
      OnClick = BCHERCHEAFFClick
      GlobalIndexImage = 'Z1413_S16G1'
    end
    object ED_AFF: THCritMaskEdit
      Left = 542
      Top = 5
      Width = 25
      Height = 21
      Color = clYellow
      TabOrder = 8
      Visible = False
      OnChange = ED_AFF1Change
      TagDispatch = 0
    end
    object ED_TIERS: THCritMaskEdit
      Left = 574
      Top = 5
      Width = 25
      Height = 21
      Color = clYellow
      TabOrder = 9
      Visible = False
      OnChange = ED_AFF1Change
      TagDispatch = 0
    end
    object ED_AFF0: THCritMaskEdit
      Left = 606
      Top = 5
      Width = 25
      Height = 21
      Color = clYellow
      TabOrder = 10
      Visible = False
      OnChange = ED_AFF1Change
      TagDispatch = 0
    end
    object BCHERCHEART: THBitBtn
      Left = 341
      Top = 3
      Width = 27
      Height = 27
      TabOrder = 6
      OnClick = BCHERCHEARTClick
      GlobalIndexImage = 'Z1264_S16G1'
    end
  end
  object Dock972: TDock97
    Left = 0
    Top = 405
    Width = 975
    Height = 34
    BoundLines = [blTop, blBottom, blLeft, blRight]
    Position = dpBottom
    object TB_NAVIGATION: TToolbar97
      Left = 0
      Top = 0
      Caption = 'Navigation'
      DockPos = 0
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      object BFirst: TToolbarButton97
        Left = 0
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Premier'
        Caption = 'Premier'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BFirstClick
        GlobalIndexImage = 'Z0301_S16G1'
      end
      object BPrev: TToolbarButton97
        Left = 28
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Pr'#233'c'#233'dent'
        Caption = 'Pr'#233'c'#233'dent'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BPrevClick
        GlobalIndexImage = 'Z0017_S16G1'
      end
      object BNext: TToolbarButton97
        Left = 56
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Suivant'
        Caption = 'Suivant'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BNextClick
        GlobalIndexImage = 'Z0031_S16G1'
      end
      object BLast: TToolbarButton97
        Left = 84
        Top = 0
        Width = 28
        Height = 27
        Hint = 'Dernier'
        Caption = 'Dernier'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BLastClick
        GlobalIndexImage = 'Z0264_S16G1'
      end
      object BCalendrier: TToolbarButton97
        Left = 112
        Top = 0
        Width = 27
        Height = 28
        Hint = 'Calendrier d'#233'but'
        Flat = False
        Opaque = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BCalendrierClick
        GlobalIndexImage = 'Z0665_S16G1'
      end
      object DateEdit: THCritMaskEdit
        Left = 139
        Top = 3
        Width = 69
        Height = 21
        Color = clWhite
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        TagDispatch = 0
      end
    end
    object TB_SELECTION: TToolbar97
      Left = 309
      Top = 0
      Caption = 'S'#233'lection'
      DockPos = 309
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      DesignSize = (
        81
        28)
      object BTNonProductive: TToolbarButton97
        Left = 0
        Top = 0
        Width = 27
        Height = 28
        Hint = 'Journ'#233'es non-facturables'
        AllowAllUp = True
        Anchors = [akRight, akBottom]
        GroupIndex = 1
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BTNonProductiveClick
        GlobalIndexImage = 'Z2000_S16G1'
      end
      object BT_LIGNES: TToolbarButton97
        Left = 27
        Top = 0
        Width = 27
        Height = 28
        Hint = 'Lignes '#224' planifier'
        AllowAllUp = True
        Anchors = [akRight, akBottom]
        GroupIndex = 1
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BLignesClick
        GlobalIndexImage = 'Z0202_S16G1'
      end
      object BT_EFFACELIGNES: TToolbarButton97
        Left = 54
        Top = 0
        Width = 27
        Height = 28
        Hint = 'Annuler la s'#233'lection'
        Anchors = [akRight, akBottom]
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BT_EFFACELIGNESClick
        GlobalIndexImage = 'M0080_S16G1'
      end
    end
    object TB_ZOOM: TToolbar97
      Left = 531
      Top = 0
      Caption = 'Zoom'
      DockPos = 531
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      DesignSize = (
        96
        28)
      object BEXCEL: TToolbarButton97
        Left = 0
        Top = 0
        Width = 27
        Height = 28
        Hint = 'Excel'
        Anchors = [akRight, akBottom]
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BExcelClick
        GlobalIndexImage = 'Z2263_S16G1'
      end
      object BINTERVENANT: TToolbarButton97
        Left = 27
        Top = 0
        Width = 27
        Height = 28
        Hint = 'Liste des intervenants'
        Anchors = [akRight, akBottom]
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BINTERVENANTClick
        GlobalIndexImage = 'Z1987_S16G1'
      end
      object BAFFAIRE: TToolbarButton97
        Left = 54
        Top = 0
        Width = 42
        Height = 28
        Hint = 'Zoom'
        Alignment = taRightJustify
        Anchors = [akRight, akBottom]
        DropdownArrow = True
        DropdownMenu = POPUPGENERAL
        Flat = False
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z0016_S16G1'
      end
    end
    object TB_ACTIONS: TToolbar97
      Left = 828
      Top = 0
      Caption = 'Actions'
      DockPos = 828
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      DesignSize = (
        136
        28)
      object BT_REFRESH: TToolbarButton97
        Left = 0
        Top = 0
        Width = 27
        Height = 28
        Hint = 'Rafraichir le planning'
        AllowAllUp = True
        Anchors = [akRight, akBottom]
        GroupIndex = 1
        ParentShowHint = False
        ShowHint = True
        OnClick = BT_REFRESHClick
        GlobalIndexImage = 'Z0338_S16G1'
      end
      object BPARAM: TToolbarButton97
        Left = 27
        Top = 0
        Width = 27
        Height = 28
        Hint = 'Param'#232'tres'
        Anchors = [akRight, akBottom]
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BParamClick
        GlobalIndexImage = 'Z0046_S16G1'
      end
      object Bprint: TToolbarButton97
        Left = 54
        Top = 0
        Width = 27
        Height = 28
        Hint = 'Imprimer'
        Anchors = [akRight, akBottom]
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BprintClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object Bannuler: TToolbarButton97
        Left = 81
        Top = 0
        Width = 27
        Height = 28
        Hint = 'Quitter'
        Anchors = [akRight, akBottom]
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BannulerClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 108
        Top = 0
        Width = 28
        Height = 28
        Hint = 'Aide'
        Caption = 'Aide'
        Anchors = [akRight, akBottom]
        DisplayMode = dmGlyphOnly
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z1117_S16G1'
      end
    end
  end
  object POPUPGENERAL: TPopupMenu
    Left = 656
    Top = 552
    object MnAffaire: TMenuItem
      Caption = 'Voir affaire'
      OnClick = MnAffaireClick
    end
    object MnClient: TMenuItem
      Caption = 'Voir client'
      OnClick = MnClientClick
    end
    object MnTache: TMenuItem
      Caption = 'T'#226'che'
      OnClick = MnTacheClick
    end
    object MnPlanning: TMenuItem
      Caption = 'Planning '
      OnClick = MnPlanningClick
    end
    object MnInterventions: TMenuItem
      Caption = 'Liste des interventions'
      OnClick = MnInterventionsClick
    end
    object MnListeInterv: TMenuItem
      Caption = 'Edition de la liste des interventions'
      OnClick = MnListeIntervClick
    end
    object MnEcole: TMenuItem
      Caption = 'Liste  des affaires '#233'coles'
      OnClick = MnEcoleClick
    end
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 840
    Top = 360
  end
end

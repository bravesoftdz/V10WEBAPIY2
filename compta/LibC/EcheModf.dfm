object FEcheModf: TFEcheModf
  Left = 307
  Top = 137
  Width = 608
  Height = 390
  HelpContext = 7529000
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Modification d'#39#233'ch'#233'ances'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 600
    Height = 106
    ActivePage = Princ
    Align = alTop
    TabOrder = 0
    object Princ: TTabSheet
      Caption = 'Standards'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 592
        Height = 78
        Align = alClient
      end
      object HLabel4: THLabel
        Left = 8
        Top = 8
        Width = 74
        Height = 13
        Caption = 'Compte &g'#233'n'#233'ral'
        FocusControl = E_GENERAL
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel1: THLabel
        Left = 8
        Top = 32
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
      object TE_EXERCICE: THLabel
        Left = 281
        Top = 8
        Width = 41
        Height = 13
        Caption = '&Exercice'
        FocusControl = E_EXERCICE
      end
      object HLabel3: THLabel
        Left = 281
        Top = 56
        Width = 100
        Height = 13
        Caption = '&Dates comptables du'
        FocusControl = E_DATECOMPTABLE
      end
      object HLabel6: THLabel
        Left = 474
        Top = 56
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATECOMPTABLE_
      end
      object HLabel10: THLabel
        Left = 281
        Top = 32
        Width = 69
        Height = 13
        Caption = '&Ech'#233'ances du'
        FocusControl = E_DATEECHEANCE
      end
      object HLabel7: THLabel
        Left = 474
        Top = 32
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATEECHEANCE_
      end
      object TE_JOURNAL: THLabel
        Left = 8
        Top = 56
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = E_JOURNAL
      end
      object E_GENERAL: THCpteEdit
        Tag = 1
        Left = 104
        Top = 4
        Width = 145
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 0
        ZoomTable = tzGLettColl
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object E_AUXILIAIRE: THCpteEdit
        Tag = 1
        Left = 104
        Top = 28
        Width = 145
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 1
        ZoomTable = tzTlettrable
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object E_EXERCICE: THValComboBox
        Left = 388
        Top = 4
        Width = 185
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 3
        OnChange = E_EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 388
        Top = 52
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 6
        Text = '01/01/1990'
        OnKeyPress = E_DATEECHEANCEKeyPress
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATECOMPTABLE_: THCritMaskEdit
        Left = 496
        Top = 52
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 7
        Text = '31/12/2099'
        OnKeyPress = E_DATEECHEANCEKeyPress
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATEECHEANCE: THCritMaskEdit
        Left = 388
        Top = 28
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 4
        Text = '01/01/1990'
        OnKeyPress = E_DATEECHEANCEKeyPress
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATEECHEANCE_: THCritMaskEdit
        Left = 496
        Top = 28
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 5
        Text = '31/12/1999'
        OnKeyPress = E_DATEECHEANCEKeyPress
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_ETATLETTRAGE: TEdit
        Left = 164
        Top = 17
        Width = 20
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 8
        Text = 'AL'
        Visible = False
      end
      object E_ECHE: TEdit
        Left = 188
        Top = 17
        Width = 20
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 9
        Text = 'X'
        Visible = False
      end
      object E_NUMECHE: THCritMaskEdit
        Left = 116
        Top = 17
        Width = 21
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 10
        Text = '1'
        Visible = False
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object XX_WHERE: TEdit
        Left = 140
        Top = 17
        Width = 20
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 11
        Text = 'E_ECRANOUVEAU="N" or E_ECRANOUVEAU="H"'
        Visible = False
      end
      object E_TRESOLETTRE: THCritMaskEdit
        Left = 212
        Top = 17
        Width = 21
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 12
        Text = 'X'
        Visible = False
        TagDispatch = 0
        Operateur = Different
      end
      object E_JOURNAL: THValComboBox
        Left = 104
        Top = 52
        Width = 145
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        Vide = True
        DataType = 'TTJOURNAUX'
      end
      object E_ANA: THCritMaskEdit
        Left = 236
        Top = 17
        Width = 21
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 13
        Text = 'X'
        Visible = False
        TagDispatch = 0
        Operateur = Different
      end
      object XX_WHERENATTIERS: TEdit
        Left = 116
        Top = 41
        Width = 20
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 14
        Visible = False
      end
      object XX_WHEREVIDE: TEdit
        Left = 144
        Top = 41
        Width = 20
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 15
        Text = 'E_JOURNAL="aaa"'
        Visible = False
      end
    end
    object Complements: TTabSheet
      Caption = 'Compl'#233'ments'
      object Bevel2: TBevel
        Left = 0
        Top = 0
        Width = 592
        Height = 78
        Align = alClient
      end
      object Label12: THLabel
        Left = 304
        Top = 32
        Width = 55
        Height = 13
        Caption = 'Ref. &interne'
        FocusControl = E_REFINTERNE
      end
      object HLabel8: THLabel
        Left = 304
        Top = 8
        Width = 33
        Height = 13
        Caption = '&Devise'
        FocusControl = E_DEVISE
      end
      object TE_NATUREPIECE: THLabel
        Left = 8
        Top = 32
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = E_NATUREPIECE
      end
      object TE_QUALIFPIECE: THLabel
        Left = 8
        Top = 57
        Width = 24
        Height = 13
        Caption = '&Type'
        FocusControl = E_QUALIFPIECE
      end
      object Label1: THLabel
        Left = 8
        Top = 8
        Width = 57
        Height = 13
        Caption = '&Num'#233'ros de'
        FocusControl = E_NUMEROPIECE
      end
      object Label2: THLabel
        Left = 154
        Top = 8
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_NUMEROPIECE_
      end
      object TE_ETABLISSEMENT: THLabel
        Left = 304
        Top = 57
        Width = 65
        Height = 13
        Caption = '&Etablissement'
        FocusControl = E_ETABLISSEMENT
      end
      object E_REFINTERNE: THCritMaskEdit
        Left = 388
        Top = 28
        Width = 149
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
      end
      object E_DEVISE: THValComboBox
        Tag = 1
        Left = 388
        Top = 4
        Width = 149
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        Vide = True
        DataType = 'TTDEVISE'
      end
      object E_NATUREPIECE: THValComboBox
        Left = 76
        Top = 28
        Width = 163
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
      object E_QUALIFPIECE: THValComboBox
        Left = 76
        Top = 53
        Width = 163
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 5
        TagDispatch = 0
        DataType = 'TTQUALPIECE'
      end
      object E_NUMEROPIECE: THCritMaskEdit
        Tag = 2
        Left = 76
        Top = 4
        Width = 65
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 0
        Text = '0'
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_NUMEROPIECE_: THCritMaskEdit
        Tag = 2
        Left = 172
        Top = 4
        Width = 65
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 1
        Text = '99999999'
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_ETABLISSEMENT: THValComboBox
        Left = 389
        Top = 53
        Width = 149
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 6
        TagDispatch = 0
        Vide = True
        DataType = 'TTETABLISSEMENT'
      end
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 106
    Width = 600
    Height = 39
    AllowDrag = False
    object PFiltres: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 35
      ClientWidth = 600
      Caption = ' '
      ClientAreaHeight = 35
      ClientAreaWidth = 600
      DockableTo = [dpTop, dpLeft]
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
        Left = 564
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Appliquer crit'#232'res'
        ParentShowHint = False
        ShowHint = True
        OnClick = BChercherClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object FFiltres: THValComboBox
        Left = 67
        Top = 5
        Width = 493
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
      end
    end
  end
  object Dock: TDock97
    Left = 0
    Top = 324
    Width = 600
    Height = 39
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 35
      ClientWidth = 600
      Caption = 'Actions'
      ClientAreaHeight = 35
      ClientAreaWidth = 600
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BReduire: TToolbarButton97
        Left = 3
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Afficher les crit'#232'res'
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
        Left = 3
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
      object BImprimer: TToolbarButton97
        Left = 471
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
        OnClick = BImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BValider: TToolbarButton97
        Left = 503
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Modifier l'#39#233'ch'#233'ance'
        Default = True
        DisplayMode = dmGlyphOnly
        Caption = 'Modifier'
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
        OnClick = GLDblClick
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 535
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
        Left = 567
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
      object BRecherche: TToolbarButton97
        Left = 35
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
      object BParamListe: TToolbarButton97
        Left = 67
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Param'#233'trer la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'Param'#232'tres'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BParamListeClick
        GlobalIndexImage = 'Z0025_S16G1'
      end
      object BZoomPiece: TToolbarButton97
        Tag = 1
        Left = 99
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Voir '#233'criture'
        DisplayMode = dmGlyphOnly
        Caption = 'Zoom'
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
        OnClick = BZoomPieceClick
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
    end
  end
  object GL: THDBGrid
    Left = 0
    Top = 145
    Width = 600
    Height = 179
    Align = alClient
    DataSource = SEche
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 3
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    Row = 1
    MultiSelection = False
    SortEnabled = False
    MyDefaultRowHeight = 0
  end
  object FindMvt: THFindDialog
    OnFind = FindMvtFind
    Left = 16
    Top = 272
  end
  object QEche: THQuery
    AutoCalcFields = False
    LockType = ltReadOnly
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    Parameters = <>
    dataBaseName = 'SOC'
    PageCriteres = Pages
    MAJAuto = False
    Distinct = False
    Left = 158
    Top = 272
  end
  object SEche: TDataSource
    DataSet = QEche
    Left = 205
    Top = 272
  end
  object HME: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Modification des '#233'ch'#233'ances;Cette '#233'ch'#233'ance d'#39'A-Nouveau n'#39'est pa' +
        's modifiable;W;O;O;O;'
      
        '1;Modification des '#233'ch'#233'ances;Les crit'#232'res de compte ne sont pas ' +
        'remplis. Confirmez-vous la recherche sur tous les comptes ?;Q;YN' +
        'C;N;N;')
    Left = 63
    Top = 272
  end
  object HMTrad: THSystemMenu
    ResizeDBGrid = True
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 252
    Top = 272
  end
  object POPF: TPopupMenu
    Left = 112
    Top = 272
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

inherited FConsEcr: TFConsEcr
  Left = 291
  Top = 203
  Width = 610
  Height = 386
  HelpContext = 7602000
  Caption = 'Consultation des comptes :'
  OnResize = FormResize
  PixelsPerInch = 96
  TextHeight = 13
  inherited Pages: THPageControl2
    Width = 602
    ActivePage = PCritere
    inherited PCritere: THTabSheet
      Caption = 'Comptes'
      inherited Bevel1: TBevel
        Width = 594
      end
      object TE_GENERAL: THLabel
        Left = 10
        Top = 32
        Width = 37
        Height = 13
        Caption = '&G'#233'n'#233'ral'
        FocusControl = E_GENERAL
      end
      object TE_AUXILIAIRE: THLabel
        Left = 10
        Top = 61
        Width = 41
        Height = 13
        Caption = '&Auxiliaire'
        FocusControl = E_AUXILIAIRE
      end
      object TGLIBELLE: THLabel
        Left = 221
        Top = 32
        Width = 239
        Height = 13
        AutoSize = False
        Caption = 'TGLIBELLE'
        ShowAccelChar = False
      end
      object TTLIBELLE: THLabel
        Left = 222
        Top = 60
        Width = 238
        Height = 13
        AutoSize = False
        Caption = 'TTLIBELLE'
        ShowAccelChar = False
      end
      object TCompte: TLabel
        Left = 77
        Top = 8
        Width = 41
        Height = 13
        Caption = 'Comptes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TLibelle: TLabel
        Left = 229
        Top = 8
        Width = 36
        Height = 13
        Caption = 'Intitul'#233's'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TSoldes: TLabel
        Left = 514
        Top = 8
        Width = 27
        Height = 13
        Caption = 'Solde'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TLeSolde: THLabel
        Left = 509
        Top = 32
        Width = 37
        Height = 13
        Caption = '&G'#233'n'#233'ral'
        FocusControl = E_GENERAL
      end
      object GeneUp: TToolbarButton97
        Left = 171
        Top = 29
        Width = 20
        Height = 19
        Hint = 'G'#233'n'#233'ral pr'#233'c'#233'dent'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = GeneUpClick
        GlobalIndexImage = 'Z0171_S16G1'
      end
      object GeneDown: TToolbarButton97
        Left = 191
        Top = 29
        Width = 20
        Height = 19
        Hint = 'G'#233'n'#233'ral suivant'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = GeneDownClick
        GlobalIndexImage = 'Z0319_S16G1'
      end
      object AuxiUp: TToolbarButton97
        Left = 171
        Top = 57
        Width = 20
        Height = 19
        Hint = 'Auxiliaire pr'#233'c'#233'dent'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = AuxiUpClick
        GlobalIndexImage = 'Z0171_S16G1'
      end
      object AuxiDown: TToolbarButton97
        Left = 191
        Top = 57
        Width = 20
        Height = 19
        Hint = 'Auxiliaire suivant'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = AuxiDownClick
        GlobalIndexImage = 'Z0319_S16G1'
      end
      object LESOLDE: THNumEdit
        Tag = 1
        Left = 470
        Top = 56
        Width = 114
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
        Masks.PositiveMask = '#,##0'
        Debit = False
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
        UseRounding = True
        Validate = False
      end
      object TPASFAIT: THEdit
        Tag = 1
        Left = 470
        Top = 56
        Width = 114
        Height = 21
        Color = clBtnFace
        Ctl3D = True
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 3
        Text = 'Non calcul'#233
        Visible = False
        TagDispatch = 0
      end
      object E_GENERAL: THCpteEdit
        Left = 61
        Top = 28
        Width = 108
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 0
        OnEnter = E_GENERALEnter
        OnExit = E_GENERALExit
        OnKeyDown = E_GENERALKeyDown
        ZoomTable = tzGeneral
        Vide = False
        Bourre = False
        Libelle = TGLIBELLE
        okLocate = True
        SynJoker = False
      end
      object E_AUXILIAIRE: THCpteEdit
        Left = 61
        Top = 56
        Width = 108
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 1
        OnEnter = E_AUXILIAIREEnter
        OnExit = E_AUXILIAIREExit
        ZoomTable = tzTiers
        Vide = False
        Bourre = False
        Libelle = TTLIBELLE
        okLocate = True
        SynJoker = False
      end
    end
    object PEcritures: TTabSheet [1]
      Caption = 'Ecritures'
      ImageIndex = 5
      object Bevel6: TBevel
        Left = 0
        Top = 0
        Width = 594
        Height = 83
        Align = alClient
      end
      object TE_JOURNAL: THLabel
        Left = 13
        Top = 10
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = E_JOURNAL
      end
      object TE_EXERCICE: THLabel
        Left = 13
        Top = 34
        Width = 41
        Height = 13
        Caption = 'E&xercice'
        FocusControl = E_EXERCICE
      end
      object TE_NATUREPIECE: THLabel
        Left = 14
        Top = 60
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = E_NATUREPIECE
      end
      object TE_NUMEROPIECE: THLabel
        Left = 328
        Top = 10
        Width = 56
        Height = 13
        Caption = 'N'#176' de &pi'#232'ce'
        FocusControl = E_NUMEROPIECE
      end
      object HLabel1: THLabel
        Left = 495
        Top = 9
        Width = 6
        Height = 13
        Caption = #224
      end
      object TE_DATECOMPTABLE: THLabel
        Left = 328
        Top = 34
        Width = 85
        Height = 13
        Caption = '&Dates comptables'
        FocusControl = E_DATECOMPTABLE
      end
      object TE_DATECOMPTABLE2: THLabel
        Left = 492
        Top = 32
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATECOMPTABLE_
      end
      object TE_DEVISE: THLabel
        Left = 328
        Top = 60
        Width = 33
        Height = 13
        Caption = 'De&vise'
        FocusControl = E_DEVISE
      end
      object E_NUMEROPIECE: THCritMaskEdit
        Left = 424
        Top = 6
        Width = 65
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_NUMEROPIECE_: THCritMaskEdit
        Left = 512
        Top = 6
        Width = 65
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_JOURNAL: THValComboBox
        Left = 91
        Top = 6
        Width = 183
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTJOURNAUX'
      end
      object E_EXERCICE: THValComboBox
        Left = 91
        Top = 30
        Width = 183
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 1
        OnChange = E_EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object E_NATUREPIECE: THValComboBox
        Left = 91
        Top = 56
        Width = 183
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
        Left = 424
        Top = 31
        Width = 65
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 5
        Text = '  /  /    '
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATECOMPTABLE_: THCritMaskEdit
        Left = 512
        Top = 31
        Width = 65
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 6
        Text = '  /  /    '
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object XX_WHERE1: TEdit
        Left = 45
        Top = 13
        Width = 17
        Height = 19
        TabStop = False
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 8
        Text = 'E_JOURNAL="zzz" AND E_NUMEROPIECE=-1'
        Visible = False
      end
      object XX_WHEREAN: TEdit
        Left = 70
        Top = 56
        Width = 17
        Height = 19
        TabStop = False
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
        Visible = False
      end
      object XX_WHERELET: TEdit
        Left = 70
        Top = 30
        Width = 17
        Height = 19
        TabStop = False
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
        Visible = False
      end
      object E_DEVISE: THValComboBox
        Left = 424
        Top = 56
        Width = 155
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 7
        TagDispatch = 0
        Vide = True
        DataType = 'TTDEVISE'
      end
      object E_CREERPAR: THCritMaskEdit
        Left = 288
        Top = 44
        Width = 17
        Height = 19
        TabStop = False
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
        Text = 'DET'
        Visible = False
        TagDispatch = 0
        Operateur = Different
      end
      object XX_WHEREQUALIF: TEdit
        Left = 286
        Top = 14
        Width = 17
        Height = 19
        TabStop = False
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
        Text = 'E_QUALIFPIECE="N"'
        Visible = False
      end
      object XX_WHEREDET: TEdit
        Left = 35
        Top = 47
        Width = 17
        Height = 19
        TabStop = False
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
        Text = 'E_CREERPAR<>"DET"'
        Visible = False
      end
      object XX_WHEREDC: TEdit
        Left = 305
        Top = 14
        Width = 17
        Height = 19
        TabStop = False
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
        Visible = False
      end
    end
    inherited PComplement: THTabSheet
      inherited Bevel2: TBevel
        Width = 594
      end
      object TE_DATEECHEANCE: THLabel
        Left = 16
        Top = 34
        Width = 54
        Height = 13
        Caption = '&Ech'#233'ances'
        FocusControl = E_DATEECHEANCE
      end
      object TE_DATEECHEANCE2: THLabel
        Left = 160
        Top = 34
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_DATEECHEANCE_
      end
      object TE_REFINTERNE: THLabel
        Left = 16
        Top = 8
        Width = 50
        Height = 13
        Caption = '&R'#233'f'#233'rence'
        FocusControl = E_REFINTERNE
      end
      object HLabel9: THLabel
        Left = 16
        Top = 61
        Width = 55
        Height = 13
        Caption = '&Code Lettre'
        FocusControl = E_LETTRAGE
      end
      object HLabel2: THLabel
        Left = 337
        Top = 34
        Width = 39
        Height = 13
        Caption = 'Lettra&ge'
        FocusControl = EtatLettrage
      end
      object TE_ETABLISSEMENT: THLabel
        Left = 337
        Top = 8
        Width = 65
        Height = 13
        Caption = 'E&tablissement'
        FocusControl = E_ETABLISSEMENT
      end
      object E_DATEECHEANCE: THCritMaskEdit
        Left = 82
        Top = 30
        Width = 71
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 1
        Text = '  /  /    '
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATEECHEANCE_: THCritMaskEdit
        Left = 176
        Top = 30
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 2
        Text = '  /  /    '
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_VALIDE: TCheckBox
        Left = 180
        Top = 59
        Width = 65
        Height = 17
        Alignment = taLeftJustify
        AllowGrayed = True
        Caption = '&Valid'#233'e'
        State = cbGrayed
        TabOrder = 4
      end
      object E_REFINTERNE: TEdit
        Left = 83
        Top = 4
        Width = 163
        Height = 21
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 0
      end
      object E_LETTRAGE: THCritMaskEdit
        Left = 82
        Top = 57
        Width = 69
        Height = 21
        Ctl3D = True
        MaxLength = 4
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
      end
      object EtatLettrage: TComboBox
        Left = 413
        Top = 30
        Width = 157
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 6
        OnChange = EtatLettrageChange
        Items.Strings = (
          'Toutes'
          'Rapproch'#233'es'
          'Non rapproch'#233'es'
          'Partiellement lettr'#233'es '
          'Totalement lettr'#233'es')
      end
      object E_ETABLISSEMENT: THValComboBox
        Left = 413
        Top = 4
        Width = 157
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 5
        TagDispatch = 0
        Vide = True
        DataType = 'TTETABLISSEMENT'
      end
      object CAvecSimu: TCheckBox
        Left = 335
        Top = 59
        Width = 110
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Ajouter &simulation'
        TabOrder = 7
        OnClick = CAvecSimuClick
      end
    end
    object PParams: TTabSheet [3]
      Caption = 'Param'#232'tres'
      ImageIndex = 6
      object Bevel7: TBevel
        Left = 0
        Top = 0
        Width = 594
        Height = 83
        Align = alClient
      end
      object RDefil: TRadioGroup
        Left = 179
        Top = 5
        Width = 159
        Height = 73
        Caption = 'D'#233'filement des comptes'
        ItemIndex = 0
        Items.Strings = (
          '&Tous'
          '&Non sold'#233's'
          '&Mouvement'#233's')
        TabOrder = 0
      end
      object RAction: TRadioGroup
        Left = 12
        Top = 5
        Width = 159
        Height = 73
        Caption = 'Mode d'#39'acc'#232's aux fonctions'
        ItemIndex = 1
        Items.Strings = (
          '&Consultation'
          'Cr'#233'ation / &Modification')
        TabOrder = 1
        OnClick = RActionClick
      end
    end
    object Pzlibre: TTabSheet [4]
      Caption = 'Tables libres'
      object Bevel5: TBevel
        Left = 0
        Top = 0
        Width = 594
        Height = 83
        Align = alClient
      end
      object TE_TABLE0: TLabel
        Left = 26
        Top = 19
        Width = 100
        Height = 13
        AutoSize = False
        Caption = '&Table libre n'#176'1'
        FocusControl = E_TABLE0
      end
      object TE_TABLE2: TLabel
        Left = 26
        Top = 51
        Width = 100
        Height = 13
        AutoSize = False
        Caption = 'Ta&ble libre n'#176'3'
        FocusControl = E_TABLE2
      end
      object TE_TABLE1: TLabel
        Left = 305
        Top = 19
        Width = 100
        Height = 13
        AutoSize = False
        Caption = 'T&able libre n'#176'2'
        FocusControl = E_TABLE1
      end
      object TE_TABLE3: TLabel
        Left = 305
        Top = 51
        Width = 100
        Height = 13
        AutoSize = False
        Caption = 'Tab&le libre n'#176'4'
        FocusControl = E_TABLE3
      end
      object E_TABLE2: THCpteEdit
        Left = 131
        Top = 47
        Width = 142
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 0
        ZoomTable = tzNatEcrE2
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object E_TABLE0: THCpteEdit
        Left = 131
        Top = 15
        Width = 142
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 1
        ZoomTable = tzNatEcrE0
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object E_TABLE3: THCpteEdit
        Left = 408
        Top = 47
        Width = 142
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 2
        ZoomTable = tzNatEcrE3
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object E_TABLE1: THCpteEdit
        Left = 408
        Top = 15
        Width = 142
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 3
        ZoomTable = tzNatEcrE1
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
    end
    inherited PAvance: THTabSheet
      inherited Bevel4: TBevel
        Width = 594
      end
      inherited bEffaceAvance: TToolbarButton97
        Left = 554
      end
      inherited Z_C1: THValComboBox
        Left = 6
        Width = 151
      end
      inherited Z_C2: THValComboBox
        Left = 6
        Width = 151
      end
      inherited Z_C3: THValComboBox
        Left = 6
        Width = 151
      end
      inherited ZO3: THValComboBox
        Left = 163
        Width = 151
      end
      inherited ZO2: THValComboBox
        Left = 163
        Width = 151
      end
      inherited ZO1: THValComboBox
        Left = 163
        Width = 151
      end
      inherited ZV1: THEdit
        Left = 319
      end
      inherited ZV2: THEdit
        Left = 319
      end
      inherited ZV3: THEdit
        Left = 319
      end
      inherited ZG2: THCombobox
        Left = 510
      end
      inherited ZG1: THCombobox
        Left = 510
      end
    end
    inherited PSQL: THTabSheet
      TabVisible = False
      inherited Bevel3: TBevel
        Width = 594
      end
      inherited Z_SQL: THSQLMemo
        Width = 594
      end
    end
  end
  inherited Dock971: TDock97
    Width = 602
    inherited PFiltres: TToolWindow97
      ClientWidth = 602
      ClientAreaWidth = 602
      inherited BCherche: TToolbarButton97
        Left = 388
      end
      inherited lpresentation: THLabel
        Left = 425
      end
      inherited FFiltres: THValComboBox
        Width = 313
      end
      inherited cbPresentations: THValComboBox
        Left = 496
      end
    end
  end
  inherited FListe: THDBGrid
    Width = 573
    Height = 149
    Options = [dgTitles, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit]
    SortEnabled = True
  end
  inherited Panel2: THPanel
    Left = 425
    inherited PListe: THPanel
      Left = 418
    end
  end
  inherited Dock: TDock97
    Top = 323
    Width = 602
    inherited PanelBouton: TToolWindow97
      ClientWidth = 602
      ClientAreaWidth = 602
      inherited bSelectAll: TToolbarButton97
        Left = 131
      end
      inherited BImprimer: TToolbarButton97
        Left = 475
      end
      inherited BOuvrir: TToolbarButton97
        Left = 507
        OnClick = FListeDblClick
      end
      inherited BAnnuler: TToolbarButton97
        Left = 539
      end
      inherited BAide: TToolbarButton97
        Left = 571
      end
      inherited Binsert: TToolbarButton97
        Left = 163
      end
      inherited BBlocNote: TToolbarButton97
        Left = 444
        Visible = True
      end
      object BActionCPT: TToolbarButton97
        Left = 226
        Top = 2
        Width = 39
        Height = 27
        Hint = 'Actions sur le compte'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = PopACPT
        Caption = 'Actions compte'
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z0611_S16G1'
      end
      object BActionEcr: TToolbarButton97
        Left = 270
        Top = 2
        Width = 39
        Height = 27
        Hint = 'Listes et '#233'ditions'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = PopAEdt
        Caption = 'Editions'
        Flat = False
        Layout = blGlyphTop
        ParentShowHint = False
        ShowHint = True
        GlobalIndexImage = 'Z0003_S16G1'
      end
    end
  end
  inherited PCumul: THPanel
    Top = 301
    Width = 602
  end
  inherited PanVBar: THPanel
    Left = 573
    Height = 149
  end
  inherited SQ: TDataSource
    Left = 104
  end
  inherited Q: THQuery
    Left = 15
  end
  inherited FindDialog: THFindDialog
    Left = 160
    Top = 256
  end
  inherited POPF: THPopupMenu
    Left = 236
  end
  inherited HMTrad: THSystemMenu
    Left = 192
  end
  inherited SD: THSaveDialog
    Left = 280
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Visualisation des '#233'critures'
      'Modification des '#233'critures'
      'courantes'
      'de simulation'
      'de pr'#233'vision'
      'de situation'
      'de r'#233'vision'
      'd'#39#224'-nouveaux'
      'Origine'
      '9;?caption?;Vous devez renseigner un compte;E;O;O;O;'
      
        '10;?caption?;Confirmez-vous le d'#233'lettrage total du compte ?;Q;YN' +
        ';Y;Y;'
      
        '11;Lettrage manuel;Les mouvements sont en devise. Voulez-vous le' +
        'ttrer en ;Q;YN;Y;Y;'
      'G'#233'n'#233'ral/Auxiliaire')
    Left = 60
    Top = 244
  end
  object PopACPT: TPopupMenu
    OnPopup = PopACPTPopup
    Left = 325
    Top = 248
    object LETTRAGEM: TMenuItem
      Caption = '&Lettrage manuel'
      OnClick = LETTRAGEMClick
    end
    object LettrageA: TMenuItem
      Caption = 'Lettrage &automatique'
      OnClick = LettrageAClick
    end
    object Delettre: TMenuItem
      Caption = '&D'#233'lettrage partiel du compte'
      OnClick = DelettreClick
    end
    object DelettreAuto: TMenuItem
      Caption = 'D'#233'lettrage total d&u compte'
      OnClick = DelettreAutoClick
    end
    object N7: TMenuItem
      Caption = '-'
    end
    object PointageG: TMenuItem
      Caption = '&Pointage'
      OnClick = PointageGClick
    end
    object DepointageG: TMenuItem
      Caption = 'D'#233'pointa&ge'
      OnClick = DepointageGClick
    end
    object N2: TMenuItem
      Caption = '-'
    end
    object ModifEcr: TMenuItem
      Caption = '&Modifier l'#39#233'criture'
      OnClick = ModifEcrClick
    end
    object InfosComp: TMenuItem
      Caption = 'Informations &compl'#233'mentaires'
      OnClick = InfosCompClick
    end
    object N6: TMenuItem
      Caption = '-'
    end
    object SaisieEcr: TMenuItem
      Caption = '&Saisir une pi'#232'ce'
      OnClick = SaisieEcrClick
    end
    object SaisieBor: TMenuItem
      Caption = 'Saisir un &bordereau'
      OnClick = SaisieBorClick
    end
    object N3: TMenuItem
      Caption = '-'
    end
    object ModifGENE: TMenuItem
      Caption = 'Modifier le compte g'#233'&n'#233'ral'
      OnClick = ModifGENEClick
    end
    object ModifAux: TMenuItem
      Caption = 'Modifier le compte au&xiliaire'
      OnClick = ModifAuxClick
    end
    object N8: TMenuItem
      Caption = '-'
    end
    object ANALREVIS: TMenuItem
      Caption = 'Anal&yse compl'#233'mentaire'
      OnClick = ANALREVISClick
    end
    object N10: TMenuItem
      Caption = '-'
    end
    object ICCDETAIL: TMenuItem
      Caption = 'Acc'#233'der au d'#233'tail du compte ICC'
      OnClick = ICCDETAILClick
    end
    object ICCCALCUL: TMenuItem
      Caption = 'Calcul des int'#233'r'#234'ts du compte courant'
      OnClick = ICCCALCULClick
    end
    object N9: TMenuItem
      Caption = '-'
    end
    object EmpruntCRE: TMenuItem
      Caption = 'Acc'#233'der au tableau d'#39'emprunt CRE'
      OnClick = EmpruntCREClick
    end
  end
  object PopAEdt: TPopupMenu
    OnPopup = PopAEdtPopup
    Left = 374
    Top = 244
    object Analytiques: TMenuItem
      Caption = 'Mouvements &analytiques'
      OnClick = AnalytiquesClick
    end
    object Immos: TMenuItem
      Caption = 'Liste des &immobilisations'
      OnClick = ImmosClick
    end
    object FicheImmo: TMenuItem
      Caption = 'Fiche i&mmobilisation'
      OnClick = FicheImmoClick
    end
    object N4: TMenuItem
      Caption = '-'
    end
    object CumulsGENE: TMenuItem
      Caption = '&Cumuls du compte'
      OnClick = CumulsGENEClick
    end
    object GLGENE: TMenuItem
      Caption = 'Grand-&livre simple'
      OnClick = GLGENEClick
    end
    object JustifGENE: TMenuItem
      Caption = '&Justificatif de solde'
      OnClick = JustifGENEClick
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object CumulsAUX: TMenuItem
      Caption = 'Cumuls au&xiliaire'
      OnClick = CumulsAUXClick
    end
    object GLAUX: TMenuItem
      Caption = 'Grand-livre auxiliaire &simple'
      OnClick = GLAUXClick
    end
    object GLAUXSITU: TMenuItem
      Caption = 'Grand-livre auxiliaire en sit&uation'
      OnClick = GLAUXSITUClick
    end
    object JUSTIFAUX: TMenuItem
      Caption = 'Justificatif de solde du &tiers'
      OnClick = JUSTIFAUXClick
    end
    object JUSTIFAUXSITU: TMenuItem
      Caption = 'Justificatif de solde en situation'
      OnClick = JUSTIFAUXSITUClick
    end
  end
  object PopZ: TPopupMenu
    OnPopup = PopZPopup
    Left = 323
    Top = 180
    object N5: TMenuItem
      Caption = '-'
    end
  end
end

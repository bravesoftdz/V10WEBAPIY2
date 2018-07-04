inherited FCFONBMP: TFCFONBMP
  Left = 162
  Top = 215
  Width = 605
  Height = 400
  Caption = 'Export CFONB'
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock: TDock97 [2]
    Top = 337
    Width = 597
    inherited PanelBouton: TToolWindow97
      ClientWidth = 597
      ClientAreaWidth = 597
      inherited bExport: TToolbarButton97 [0]
        Left = 130
        Visible = False
      end
      inherited BReduire: TToolbarButton97 [1]
      end
      inherited bSelectAll: TToolbarButton97
        Left = 99
        Visible = True
      end
      inherited BAgrandir: TToolbarButton97 [3]
      end
      inherited BRechercher: TToolbarButton97 [4]
      end
      inherited BParamListe: TToolbarButton97 [5]
      end
      inherited BImprimer: TToolbarButton97
        Left = 476
      end
      inherited BOuvrir: TToolbarButton97
        Left = 508
        Hint = 'Valider'
      end
      inherited BAnnuler: TToolbarButton97
        Left = 540
      end
      inherited BAide: TToolbarButton97
        Left = 572
      end
      inherited BBlocNote: TToolbarButton97
        Left = 444
      end
      object bRib: TToolbarButton97
        Left = 131
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Modifier le RIB'
        Enabled = False
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
        OnClick = bRibClick
        GlobalIndexImage = 'Z2098_S16G1'
        IsControl = True
      end
      object BCtrlRib: TToolbarButton97
        Tag = 1
        Left = 197
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Contr'#244'ler les RIB de la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'G'#233'n'#233'ration'
        Flat = False
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
        OnClick = BCtrlRibClick
        GlobalIndexImage = 'Z0040_S16G1'
        IsControl = True
      end
      object FExport: TCheckBox
        Left = 284
        Top = 8
        Width = 104
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Liste d'#39'exportation'
        TabOrder = 0
        Visible = False
      end
    end
  end
  inherited Panel2: THPanel [3]
    Left = 463
    inherited PListe: THPanel
      Left = 413
    end
  end
  inherited Dock971: TDock97 [4]
    Top = 120
    Width = 597
    inherited PFiltres: TToolWindow97
      ClientWidth = 597
      ClientAreaWidth = 597
      inherited BCherche: TToolbarButton97
        Left = 384
      end
      inherited lpresentation: THLabel
        Left = 425
      end
      inherited FFiltres: THValComboBox
        Width = 309
      end
      inherited cbPresentations: THValComboBox
        Left = 496
      end
    end
  end
  inherited Pages: THPageControl2 [5]
    Width = 597
    Height = 120
    ActivePage = PComplement
    inherited PCritere: THTabSheet
      inherited Bevel1: TBevel
        Width = 589
        Height = 92
      end
      object TE_EXERCICE: THLabel
        Left = 293
        Top = 14
        Width = 41
        Height = 13
        Caption = '&Exercice'
        FocusControl = E_EXERCICE
      end
      object HLabel6: THLabel
        Left = 482
        Top = 40
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATECOMPTABLE_
      end
      object HLabel3: THLabel
        Left = 293
        Top = 40
        Width = 100
        Height = 13
        Caption = '&Dates comptables du'
        FocusControl = E_DATECOMPTABLE
      end
      object TE_JOURNAL: THLabel
        Left = 9
        Top = 14
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = E_JOURNAL
      end
      object TE_NATUREPIECE: THLabel
        Left = 9
        Top = 40
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = E_NATUREPIECE
      end
      object HLabel2: THLabel
        Left = 293
        Top = 67
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
      object E_EXERCICE: THValComboBox
        Left = 400
        Top = 10
        Width = 179
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 2
        OnChange = E_EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object E_DATECOMPTABLE_: THCritMaskEdit
        Left = 501
        Top = 36
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 4
        Text = '31/12/2099'
        OnKeyPress = E_DATECOMPTABLEKeyPress
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 400
        Top = 36
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 3
        Text = '01/01/1990'
        OnKeyPress = E_DATECOMPTABLEKeyPress
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_JOURNAL: THValComboBox
        Left = 82
        Top = 10
        Width = 181
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTJALBANQUEOD'
      end
      object E_NATUREPIECE: THValComboBox
        Left = 82
        Top = 36
        Width = 181
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 1
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATPIECEBANQUE'
      end
      object E_ECHE: TEdit
        Left = 121
        Top = 21
        Width = 16
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 7
        Text = 'X'
        Visible = False
      end
      object XX_WHERE1: TEdit
        Left = 140
        Top = 21
        Width = 16
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 10
        Text = 'E_JOURNAL="ZZZ" and E_EXERCICE="ZZZ" AND E_NUMEROPIECE=-1'
        Visible = False
      end
      object E_AUXILIAIRE: THCritMaskEdit
        Left = 160
        Top = 21
        Width = 16
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 11
        Visible = False
        TagDispatch = 0
        Operateur = Superieur
      end
      object E_AUXILIAIRE_: THCritMaskEdit
        Left = 179
        Top = 21
        Width = 16
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 12
        Visible = False
        TagDispatch = 0
        Operateur = Inferieur
      end
      object E_QUALIFPIECE: TEdit
        Left = 199
        Top = 21
        Width = 16
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 8
        Text = 'N'
        Visible = False
      end
      object E_NUMECHE: THCritMaskEdit
        Left = 218
        Top = 21
        Width = 16
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 9
        Text = '0'
        Visible = False
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_GENERAL: THCpteEdit
        Tag = 1
        Left = 400
        Top = 63
        Width = 76
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 5
        OnChange = Auxiliaire1Change
        ZoomTable = tzGCollectif
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object CTIDTIC: TCheckBox
        Left = 9
        Top = 65
        Width = 86
        Height = 17
        Alignment = taLeftJustify
        Caption = '&TIC/TID'
        TabOrder = 6
        OnClick = CTIDTICClick
      end
      object XX_WHERENATTIERS: TEdit
        Left = 96
        Top = 21
        Width = 20
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 13
        Visible = False
      end
    end
    inherited PComplement: THTabSheet
      inherited Bevel2: TBevel
        Width = 589
        Height = 92
      end
      object c: THLabel
        Left = 9
        Top = 40
        Width = 78
        Height = 13
        Caption = '&Modes paiement'
        FocusControl = E_MODEPAIE
      end
      object HAuxiliaire1: THLabel
        Left = 308
        Top = 40
        Width = 61
        Height = 13
        Caption = '&Auxiliaires de'
        FocusControl = Auxiliaire1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HAuxiliaire2: THLabel
        Left = 476
        Top = 40
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = Auxiliaire2
      end
      object HLabel10: THLabel
        Left = 9
        Top = 66
        Width = 69
        Height = 13
        Caption = 'E&ch'#233'ances du'
        FocusControl = E_DATEECHEANCE
      end
      object HLabel7: THLabel
        Left = 174
        Top = 66
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATEECHEANCE_
      end
      object TE_NUMEROPIECE: THLabel
        Left = 308
        Top = 14
        Width = 56
        Height = 13
        Caption = 'N'#176' de &pi'#232'ce'
        FocusControl = E_NUMEROPIECE
      end
      object HLabel4: THLabel
        Left = 476
        Top = 14
        Width = 6
        Height = 13
        Caption = #224
      end
      object TE_DEVISE: THLabel
        Left = 9
        Top = 14
        Width = 33
        Height = 13
        Caption = '&Devise'
        FocusControl = E_DEVISE
      end
      object cExport: TCheckBox
        Left = 488
        Top = 64
        Width = 90
        Height = 17
        Alignment = taLeftJustify
        AllowGrayed = True
        Caption = 'D'#233'j'#224' e&xport'#233's'
        TabOrder = 9
        OnClick = cExportClick
      end
      object cRIB: TCheckBox
        Left = 306
        Top = 64
        Width = 92
        Height = 17
        Alignment = taLeftJustify
        AllowGrayed = True
        Caption = '&RIB renseign'#233
        State = cbGrayed
        TabOrder = 8
        OnClick = cRIBClick
      end
      object XX_WHERERIB: TEdit
        Left = 119
        Top = 21
        Width = 16
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 10
        Visible = False
      end
      object XX_WHEREEXPORT: TEdit
        Left = 142
        Top = 21
        Width = 16
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 11
        Text = 'E_CFONBOK<>"X"'
        Visible = False
      end
      object XX_WHEREAUX: TEdit
        Left = 209
        Top = 21
        Width = 16
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 12
        Text = 'E_AUXILIAIRE<>""'
        Visible = False
      end
      object E_ECRANOUVEAU: TEdit
        Left = 97
        Top = 21
        Width = 16
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 13
        Text = 'N'
        Visible = False
      end
      object XX_WHERENAT: TEdit
        Left = 231
        Top = 21
        Width = 16
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 14
        Text = 
          'E_NATUREPIECE="RC" OR E_NATUREPIECE="RF" OR E_NATUREPIECE="OD" O' +
          'R E_NATUREPIECE="OC" OR E_NATUREPIECE="OF"'
        Visible = False
      end
      object XX_WHEREDC: TEdit
        Left = 251
        Top = 21
        Width = 16
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 15
        Visible = False
      end
      object E_MODEPAIE: THMultiValComboBox
        Left = 91
        Top = 36
        Width = 181
        Height = 21
        TabOrder = 1
        Style = csDialog
        Abrege = False
        DataType = 'TTMODEPAIE'
        Complete = False
        OuInclusif = False
      end
      object Auxiliaire1: THCpteEdit
        Tag = 1
        Left = 376
        Top = 36
        Width = 82
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 6
        OnChange = Auxiliaire1Change
        ZoomTable = tzTlettrable
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object Auxiliaire2: THCpteEdit
        Tag = 1
        Left = 496
        Top = 36
        Width = 82
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 7
        OnChange = Auxiliaire2Change
        ZoomTable = tzTlettrable
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object E_DATEECHEANCE: THCritMaskEdit
        Left = 91
        Top = 62
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 2
        Text = '01/01/1990'
        OnKeyPress = E_DATECOMPTABLEKeyPress
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATEECHEANCE_: THCritMaskEdit
        Left = 195
        Top = 62
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 3
        Text = '31/12/2099'
        OnKeyPress = E_DATECOMPTABLEKeyPress
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_NUMEROPIECE: THCritMaskEdit
        Left = 376
        Top = 10
        Width = 82
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_NUMEROPIECE_: THCritMaskEdit
        Left = 496
        Top = 10
        Width = 82
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 5
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_DEVISE: THValComboBox
        Tag = 1
        Left = 91
        Top = 10
        Width = 181
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTDEVISE'
      end
    end
    object PLibres: TTabSheet [2]
      Caption = 'Tables libres'
      object Bevel5: TBevel
        Left = 0
        Top = 0
        Width = 589
        Height = 92
        Align = alClient
      end
      object TT_TABLE0: THLabel
        Left = 13
        Top = 5
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE0'
      end
      object TT_TABLE1: THLabel
        Left = 128
        Top = 5
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE1'
      end
      object TT_TABLE2: THLabel
        Left = 243
        Top = 5
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE2'
      end
      object TT_TABLE3: THLabel
        Left = 357
        Top = 5
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE3'
      end
      object TT_TABLE4: THLabel
        Left = 469
        Top = 5
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE4'
      end
      object TT_TABLE5: THLabel
        Left = 13
        Top = 47
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE5'
      end
      object TT_TABLE6: THLabel
        Left = 128
        Top = 47
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE6'
      end
      object TT_TABLE7: THLabel
        Left = 243
        Top = 47
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE7'
      end
      object TT_TABLE8: THLabel
        Left = 357
        Top = 47
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE8'
      end
      object TT_TABLE9: THLabel
        Left = 469
        Top = 46
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE9'
      end
      object T_TABLE4: THCpteEdit
        Left = 469
        Top = 20
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
        Left = 357
        Top = 20
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
        Left = 243
        Top = 20
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
        Left = 128
        Top = 20
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
        Left = 13
        Top = 20
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
        Left = 13
        Top = 60
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
        Left = 128
        Top = 60
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
        Left = 243
        Top = 60
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
        Left = 357
        Top = 60
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
        Left = 469
        Top = 60
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
    inherited PAvance: THTabSheet
      inherited Bevel4: TBevel
        Width = 589
        Height = 92
      end
      inherited Z_C1: THValComboBox
        Width = 143
      end
      inherited Z_C2: THValComboBox
        Width = 143
      end
      inherited Z_C3: THValComboBox
        Width = 143
      end
      inherited ZO3: THValComboBox
        Left = 156
        Width = 143
      end
      inherited ZO2: THValComboBox
        Left = 156
        Width = 143
      end
      inherited ZO1: THValComboBox
        Left = 156
        Width = 143
      end
      inherited ZV1: THEdit
        Left = 305
        Width = 197
      end
      inherited ZV2: THEdit
        Left = 305
        Width = 197
      end
      inherited ZV3: THEdit
        Left = 305
        Width = 197
      end
      inherited ZG2: THCombobox
        Left = 513
      end
      inherited ZG1: THCombobox
        Left = 513
      end
    end
    inherited PSQL: THTabSheet
      inherited Bevel3: TBevel
        Width = 589
        Height = 92
      end
      inherited Z_SQL: THSQLMemo
        Width = 589
        Height = 92
      end
    end
  end
  inherited FListe: THDBGrid [6]
    Top = 161
    Width = 568
    Height = 154
    MultiSelection = True
    MultiFieds = 'E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_NUMEROPIECE;E_NUMLIGNE;'
  end
  inherited PCumul: THPanel
    Top = 315
    Width = 597
  end
  inherited PanVBar: THPanel
    Left = 568
    Top = 161
    Height = 154
  end
  inherited SQ: TDataSource
    OnDataChange = SQDataChange
    Left = 45
    Top = 237
  end
  inherited Q: THQuery
    Top = 237
  end
  inherited FindDialog: THFindDialog
    Left = 166
    Top = 237
  end
  inherited POPF: THPopupMenu
    Left = 117
    Top = 237
  end
  inherited HMTrad: THSystemMenu
    Left = 215
    Top = 237
  end
  inherited SD: THSaveDialog
    Left = 312
    Top = 237
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Export ou '#233'mission impossible. Vous n'#39'avez rien s'#233'le' +
        'ctionn'#233';E;O;O;O;'
      
        '1;?caption?;Confirmez-vous l'#39'export ou l'#39#233'mission des '#233'ch'#233'ances ' +
        's'#233'lectionn'#233'es ?;Q;YN;Y;Y;'
      
        '2;?caption?;ATTENTION ! Vous avez tout s'#233'lectionn'#233'. Confirmez-vo' +
        'us l'#39'export ou l'#39#233'mission ?;W;YN;Y;Y;'
      
        '3;?caption?;Vous devez selectionner un mod'#232'le de document pour l' +
        #39#233'mission du bordereau;W;O;O;O;'
      'Emission de bordereaux'
      'Export CFONB des encaissements'
      'Export CFONB des d'#233'caissements'
      'Emission de lettres-BOR'
      
        '8;?caption?;Vous devez selectionner un mod'#232'le de document;W;O;O;' +
        'O;'
      'Mod'#232'le'
      '10;?caption?;Vous devez renseigner un compte g'#233'n'#233'ral;W;O;O;O;'
      'Emission de lettres-traite'
      
        '12;?caption?;ATTENTION ! Vous avez des banques diff'#233'rentes. L'#39'ex' +
        'port sera incoh'#233'rent. Confirmez-vous l'#39'export ou l'#39#233'mission ?;W;' +
        'YN;Y;Y;'
      'Export CFONB des traites-BOR '#224' l'#39'escompte'
      'Export CFONB des traites-BOR '#224' l'#39'encaissement'
      'Export CFONB des pr'#233'l'#232'vements'
      'Export CFONB des virements'
      'Export CFONB des BOR '
      'Export CFONB des virements internationaux'
      '19'
      ' ')
    Left = 263
    Top = 237
  end
end

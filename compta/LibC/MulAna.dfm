inherited FMulAna: TFMulAna
  Left = 407
  Top = 260
  Width = 602
  Caption = 'Ecritures analytiques'
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock: TDock97 [2]
    Width = 594
    inherited PanelBouton: TToolWindow97
      ClientWidth = 594
      ClientAreaWidth = 594
      inherited BImprimer: TToolbarButton97
        Left = 464
      end
      inherited BOuvrir: TToolbarButton97
        Left = 496
        OnClick = FListeDblClick
      end
      inherited BAnnuler: TToolbarButton97
        Left = 528
      end
      inherited BAide: TToolbarButton97
        Left = 560
      end
      inherited BBlocNote: TToolbarButton97
        Left = 432
        Visible = True
      end
      object BTobV: TToolbarButton97
        Left = 348
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Analyse compl'#233'mentaire'
        DisplayMode = dmGlyphOnly
        Caption = 'Analyse'
        Flat = False
        NumGlyphs = 2
        OnClick = BTobVClick
        GlobalIndexImage = 'Z0105_S16G2'
      end
    end
  end
  inherited FListe: THDBGrid [3]
    Width = 565
  end
  inherited Dock971: TDock97 [4]
    Width = 594
    inherited PFiltres: TToolWindow97
      ClientWidth = 594
      ClientAreaWidth = 594
      inherited BCherche: TToolbarButton97
        Left = 382
      end
      inherited FFiltres: THValComboBox
        Width = 309
      end
    end
  end
  inherited Panel2: THPanel
    Left = 308
    Top = 248
    inherited PListe: THPanel
      Left = 410
    end
  end
  inherited Pages: THPageControl2 [6]
    Width = 594
    ActivePage = PCritere
    inherited PCritere: THTabSheet
      inherited Bevel1: TBevel
        Width = 586
      end
      object TY_JOURNAL: TLabel
        Left = 4
        Top = 9
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = Y_JOURNAL
      end
      object TY_EXERCICE: THLabel
        Left = 3
        Top = 34
        Width = 41
        Height = 13
        Caption = '&Exercice'
        FocusControl = Y_EXERCICE
      end
      object TY_AXE: TLabel
        Left = 274
        Top = 60
        Width = 18
        Height = 13
        Caption = '&Axe'
        FocusControl = Y_AXE
      end
      object TY_NUMEROPIECE: TLabel
        Left = 274
        Top = 9
        Width = 67
        Height = 13
        Caption = 'Num'#233'ro &Pi'#232'ce'
        FocusControl = Y_NUMEROPIECE
      end
      object HLabel1: THLabel
        Left = 452
        Top = 9
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = Y_NUMEROPIECE_
      end
      object TY_DATECOMPTABLE_: THLabel
        Left = 451
        Top = 34
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = Y_DATECOMPTABLE_
      end
      object TY_DATECOMPTABLE: THLabel
        Left = 274
        Top = 34
        Width = 85
        Height = 13
        Caption = '&Dates comptables'
        FocusControl = Y_DATECOMPTABLE
      end
      object Y_JOURNAL: THValComboBox
        Left = 54
        Top = 5
        Width = 203
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        MaxLength = 3
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTJOURNAUX'
      end
      object Y_EXERCICE: THValComboBox
        Left = 54
        Top = 30
        Width = 203
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 3
        OnChange = Y_EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object Y_AXE: THValComboBox
        Left = 364
        Top = 56
        Width = 188
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 8
        OnChange = Y_AXEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTAXE'
      end
      object Y_NUMVENTIL: THCritMaskEdit
        Left = 88
        Top = 20
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
        Text = '1'
        Visible = False
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object Y_NUMVENTIL_: THCritMaskEdit
        Left = 112
        Top = 20
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
        Text = '9999'
        Visible = False
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object Y_NUMEROPIECE: THCritMaskEdit
        Left = 364
        Top = 5
        Width = 80
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 1
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object Y_NUMEROPIECE_: THCritMaskEdit
        Left = 471
        Top = 5
        Width = 80
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object Y_DATECOMPTABLE_: THCritMaskEdit
        Left = 470
        Top = 30
        Width = 81
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 5
        Text = '31/12/2099'
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object Y_DATECOMPTABLE: THCritMaskEdit
        Left = 364
        Top = 30
        Width = 81
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 4
        Text = '01/01/1900'
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object Y_TYPEANALYTIQUE: TCheckBox
        Left = 138
        Top = 58
        Width = 119
        Height = 17
        Alignment = taLeftJustify
        AllowGrayed = True
        Caption = '&OD analytiques'
        State = cbGrayed
        TabOrder = 7
      end
      object Y_VALIDE: TCheckBox
        Left = 2
        Top = 58
        Width = 65
        Height = 17
        Alignment = taLeftJustify
        AllowGrayed = True
        Caption = '&Valid'#233'e'
        State = cbGrayed
        TabOrder = 6
      end
      object XX_WHERE: TEdit
        Left = 138
        Top = 20
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
        Text = 'XX_WHERE'
        Visible = False
      end
    end
    inherited PComplement: THTabSheet
      inherited Bevel2: TBevel
        Width = 586
      end
      object TY_QUALIFPIECE: THLabel
        Left = 11
        Top = 9
        Width = 24
        Height = 13
        Caption = '&Type'
        FocusControl = Y_QUALIFPIECE
      end
      object TY_DATECREATION: THLabel
        Left = 11
        Top = 34
        Width = 48
        Height = 13
        Caption = '&Cr'#233#233'es de'
        FocusControl = Y_DATECREATION
      end
      object TY_DATECREATION_: THLabel
        Left = 160
        Top = 34
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = Y_DATECREATION_
      end
      object TY_DEVISE: THLabel
        Left = 13
        Top = 59
        Width = 33
        Height = 13
        Caption = '&Devise'
        FocusControl = Y_DEVISE
      end
      object TY_GENERAL: TLabel
        Left = 268
        Top = 9
        Width = 37
        Height = 13
        Caption = '&G'#233'n'#233'ral'
        FocusControl = Y_GENERAL
      end
      object TY_SECTION: TLabel
        Left = 405
        Top = 9
        Width = 36
        Height = 13
        Caption = '&Section'
        FocusControl = Y_SECTION
      end
      object TY_ETABLISSSEMENT: THLabel
        Left = 268
        Top = 34
        Width = 34
        Height = 13
        Caption = '&Etablis.'
        FocusControl = Y_ETABLISSEMENT
      end
      object Y_QUALIFPIECE: THValComboBox
        Left = 72
        Top = 5
        Width = 181
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 0
        OnChange = Y_QUALIFPIECEChange
        TagDispatch = 0
        DataType = 'TTQUALPIECE'
      end
      object Y_DATECREATION: THCritMaskEdit
        Left = 72
        Top = 30
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 3
        Text = '01/01/1900'
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object Y_DATECREATION_: THCritMaskEdit
        Left = 184
        Top = 30
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 4
        Text = '31/12/2099'
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object Y_DEVISE: THValComboBox
        Left = 72
        Top = 56
        Width = 181
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 6
        TagDispatch = 0
        Vide = True
        DataType = 'TTDEVISE'
      end
      object RUne: TCheckBox
        Left = 266
        Top = 59
        Width = 89
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Que la 1'#176' &ligne'
        TabOrder = 7
        OnClick = RUneClick
      end
      object Y_GENERAL: THCpteEdit
        Left = 311
        Top = 5
        Width = 89
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 1
        ZoomTable = tzGventilable
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object Y_SECTION: THCpteEdit
        Left = 453
        Top = 5
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 2
        ZoomTable = tzSection
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object Y_ETABLISSEMENT: THValComboBox
        Left = 312
        Top = 30
        Width = 233
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 5
        OnChange = Y_QUALIFPIECEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTETABLISSEMENT'
      end
    end
    object Pzlibre: TTabSheet [2]
      Caption = 'Tables libres'
      object Bevel5: TBevel
        Left = 0
        Top = 0
        Width = 586
        Height = 83
        Align = alClient
      end
      object TY_TABLE0: TLabel
        Left = 34
        Top = 18
        Width = 100
        Height = 13
        AutoSize = False
        Caption = '&Table libre n'#176'1'
        FocusControl = Y_TABLE0
      end
      object TY_TABLE2: TLabel
        Left = 34
        Top = 50
        Width = 100
        Height = 13
        AutoSize = False
        Caption = 'Ta&ble libre n'#176'3'
        FocusControl = Y_TABLE2
      end
      object TY_TABLE3: TLabel
        Left = 303
        Top = 50
        Width = 100
        Height = 13
        AutoSize = False
        Caption = 'Tab&le libre n'#176'4'
        FocusControl = Y_TABLE3
      end
      object TY_TABLE1: TLabel
        Left = 303
        Top = 18
        Width = 100
        Height = 13
        AutoSize = False
        Caption = 'T&able libre n'#176'2'
        FocusControl = Y_TABLE1
      end
      object Y_TABLE0: THCpteEdit
        Left = 136
        Top = 14
        Width = 142
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 0
        ZoomTable = tzNatEcrA0
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object Y_TABLE2: THCpteEdit
        Left = 136
        Top = 46
        Width = 142
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 1
        ZoomTable = tzNatEcrA2
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object Y_TABLE3: THCpteEdit
        Left = 404
        Top = 46
        Width = 142
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 2
        ZoomTable = tzNatEcrA3
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object Y_TABLE1: THCpteEdit
        Left = 404
        Top = 14
        Width = 142
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 3
        ZoomTable = tzNatEcrA1
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
    end
    inherited PAvance: THTabSheet
      inherited Bevel4: TBevel
        Width = 586
      end
      inherited bEffaceAvance: TToolbarButton97
        Left = 557
      end
      inherited Z_C3: THValComboBox [2]
        Width = 155
      end
      inherited ZO3: THValComboBox [3]
        Left = 164
        Width = 157
      end
      inherited ZO2: THValComboBox [4]
        Left = 164
        Width = 157
      end
      inherited ZV2: THEdit [5]
        Left = 325
      end
      inherited ZV3: THEdit [6]
        Left = 325
      end
      inherited ZG2: THCombobox [7]
        Left = 514
      end
      inherited ZG1: THCombobox [8]
        Left = 514
      end
      inherited ZV1: THEdit [9]
        Left = 325
      end
      inherited ZO1: THValComboBox [10]
        Left = 164
        Width = 157
      end
      inherited Z_C1: THValComboBox [11]
        Width = 155
      end
      inherited Z_C2: THValComboBox [12]
        Width = 155
      end
    end
    inherited PSQL: THTabSheet
      inherited Bevel3: TBevel
        Width = 586
      end
      inherited Z_SQL: THSQLMemo
        Width = 586
      end
    end
  end
  inherited PCumul: THPanel
    Width = 594
  end
  inherited PanVBar: THPanel
    Left = 565
  end
  inherited SQ: TDataSource
    Left = 50
    Top = 256
  end
  inherited Q: THQuery
    Left = 8
    Top = 256
  end
  inherited FindDialog: THFindDialog
    Left = 216
    Top = 256
  end
  inherited POPF: THPopupMenu
    Left = 174
    Top = 256
  end
  inherited HMTrad: THSystemMenu
    OnBeforeTraduc = HMTradBeforeTraduc
    Left = 131
    Top = 256
  end
  inherited SD: THSaveDialog
    Left = 264
    Top = 256
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Visualisation des '#233'critures analytiques'
      'Modification des '#233'critures analytiques'
      'Origine'
      'Visualisation des '#233'critures d'#39'A-Nouveaux analytiques '
      'Modification des '#233'critures d'#39'A-Nouveaux analytiques '
      
        '5;?caption?;Vous n'#39'avez pas renseign'#233' de journal ni de section. ' +
        'La s'#233'lection peut '#234'tre importante. Confirmez-vous l'#39'analyse ?;Q;' +
        'YN;Y;Y;'
      '6;')
    Left = 89
    Top = 256
  end
end

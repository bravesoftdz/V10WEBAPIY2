inherited FExtourne: TFExtourne
  Left = 379
  Top = 169
  Width = 600
  HelpContext = 7677000
  Caption = 'Extourne d'#39#233'critures'
  PixelsPerInch = 96
  TextHeight = 13
  inherited Pages: THPageControl2
    Width = 592
    ActivePage = PCritere
    inherited PCritere: THTabSheet
      object Bevel5: TBevel
        Left = 0
        Top = 0
        Width = 584
        Height = 83
        Align = alClient
      end
      object TE_EXERCICE: THLabel
        Left = 8
        Top = 12
        Width = 41
        Height = 13
        Caption = '&Exercice'
        FocusControl = E_EXERCICE
      end
      object TE_JOURNAL: THLabel
        Left = 8
        Top = 36
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = E_JOURNAL
      end
      object TE_NUMEROPIECE: THLabel
        Left = 288
        Top = 56
        Width = 37
        Height = 13
        Caption = 'N&um'#233'ro'
        FocusControl = E_NUMEROPIECE
      end
      object HLabel1: THLabel
        Left = 479
        Top = 60
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_NUMEROPIECE_
      end
      object TE_DATECOMPTABLE: THLabel
        Left = 288
        Top = 12
        Width = 100
        Height = 13
        Caption = '&Dates comptables du'
        FocusControl = E_DATECOMPTABLE
      end
      object TE_DATECOMPTABLE2: THLabel
        Left = 476
        Top = 12
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATECOMPTABLE_
      end
      object TE_DATEECHEANCE: THLabel
        Left = 288
        Top = 36
        Width = 72
        Height = 13
        Caption = 'Ec&h'#233'ances de '
        FocusControl = E_DATEECHEANCE
      end
      object TE_DATEECHEANCE2: THLabel
        Left = 479
        Top = 36
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_DATEECHEANCE_
      end
      object TE_NATUREPIECE: THLabel
        Left = 8
        Top = 60
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = E_NATUREPIECE
      end
      object E_EXERCICE: THValComboBox
        Left = 64
        Top = 8
        Width = 195
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 0
        OnChange = E_EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object E_JOURNAL: THValComboBox
        Left = 64
        Top = 32
        Width = 195
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        Vide = True
        DataType = 'TTJALSAISIE'
      end
      object E_NUMEROPIECE: THCritMaskEdit
        Left = 396
        Top = 56
        Width = 69
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 7
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_NUMEROPIECE_: THCritMaskEdit
        Left = 504
        Top = 56
        Width = 69
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 8
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 396
        Top = 8
        Width = 69
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
      object E_DATECOMPTABLE_: THCritMaskEdit
        Left = 504
        Top = 8
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
      object E_NUMECHE: THCritMaskEdit
        Left = 84
        Top = 16
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
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_ECRANOUVEAU: THCritMaskEdit
        Left = 100
        Top = 16
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
        Text = 'N'
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object E_TRESOLETTRE: THCritMaskEdit
        Left = 100
        Top = 32
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
        Text = 'X'
        Visible = False
        TagDispatch = 0
        Operateur = Different
      end
      object XX_WHERENAT: TEdit
        Left = 84
        Top = 32
        Width = 17
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
        Text = 'E_NATUREPIECE<>"ECC"'
        Visible = False
      end
      object E_DATEECHEANCE: THCritMaskEdit
        Left = 396
        Top = 32
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 4
        Text = '  /  /    '
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATEECHEANCE_: THCritMaskEdit
        Left = 504
        Top = 32
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 5
        Text = '  /  /    '
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_NATUREPIECE: THValComboBox
        Left = 64
        Top = 56
        Width = 195
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 6
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATUREPIECE'
      end
      object E_NUMLIGNE: THCritMaskEdit
        Left = 84
        Top = 48
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
        Text = '1'
        Visible = False
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_CREERPAR: THCritMaskEdit
        Left = 100
        Top = 48
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
        Text = 'DET'
        Visible = False
        TagDispatch = 0
        Operateur = Different
      end
      object XX_WHERESAIS: TEdit
        Left = 116
        Top = 32
        Width = 17
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
        Text = 'E_MODESAISIE="-" OR E_MODESAISIE=""'
        Visible = False
      end
    end
    object PComplements: TTabSheet [1]
      Caption = 'Compl'#233'ments'
      ImageIndex = 4
      object TE_REFINTERNE: THLabel
        Left = 13
        Top = 8
        Width = 85
        Height = 13
        Caption = '&R'#233'f'#233'rence interne'
        FocusControl = E_REFINTERNE
      end
      object TE_GENERAL: THLabel
        Left = 13
        Top = 34
        Width = 37
        Height = 13
        Caption = '&G'#233'n'#233'ral'
        FocusControl = E_GENERAL
      end
      object TE_AUXILIAIRE: THLabel
        Left = 13
        Top = 60
        Width = 41
        Height = 13
        Caption = '&Auxiliaire'
        FocusControl = E_AUXILIAIRE
      end
      object E_REFINTERNE: TEdit
        Left = 114
        Top = 4
        Width = 155
        Height = 21
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 0
      end
      object E_GENERAL: THCpteEdit
        Left = 114
        Top = 30
        Width = 155
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 1
        ZoomTable = tzGeneral
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object E_AUXILIAIRE: THCpteEdit
        Left = 114
        Top = 57
        Width = 155
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 2
        ZoomTable = tzTiers
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object E_VALIDE: TCheckBox
        Left = 375
        Top = 6
        Width = 148
        Height = 17
        Alignment = taLeftJustify
        AllowGrayed = True
        Caption = '&Valid'#233'e'
        State = cbGrayed
        TabOrder = 3
      end
      object CDejaExt: TCheckBox
        Left = 375
        Top = 32
        Width = 148
        Height = 17
        Alignment = taLeftJustify
        AllowGrayed = True
        Caption = 'Ecriture d'#233'&ja extourn'#233'e'
        State = cbGrayed
        TabOrder = 4
        OnClick = CDejaExtClick
      end
      object XX_WHEREEXT: TEdit
        Left = 300
        Top = 12
        Width = 17
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
        TabOrder = 5
        Visible = False
      end
    end
    inherited PComplement: THTabSheet
      Caption = 'Param'#232'tres'
      object HLabel2: THLabel
        Left = 8
        Top = 12
        Width = 65
        Height = 13
        Caption = '&Etablissement'
        FocusControl = E_ETABLISSEMENT
      end
      object TE_DEVISE: THLabel
        Left = 8
        Top = 60
        Width = 33
        Height = 13
        Caption = '&Devise'
        FocusControl = E_DEVISE
      end
      object TE_QUALIFPIECE: THLabel
        Left = 8
        Top = 36
        Width = 24
        Height = 13
        Caption = '&Type'
        FocusControl = E_QUALIFPIECE
      end
      object Label2: TLabel
        Left = 364
        Top = 12
        Width = 198
        Height = 13
        Alignment = taCenter
        Caption = 'Param'#232'tres de l'#39#233'criture d'#39'extourne'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object HLabel6: THLabel
        Left = 332
        Top = 60
        Width = 23
        Height = 13
        Caption = 'Date'
        FocusControl = CONTRE_DATE
      end
      object HLabel3: THLabel
        Left = 332
        Top = 36
        Width = 24
        Height = 13
        Caption = '&Type'
        FocusControl = E_QUALIFPIECE
      end
      object E_ETABLISSEMENT: THValComboBox
        Left = 80
        Top = 9
        Width = 185
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTETABLISSEMENT'
      end
      object E_DEVISE: THValComboBox
        Left = 80
        Top = 56
        Width = 185
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        Vide = True
        DataType = 'TTDEVISE'
      end
      object E_QUALIFPIECE: THValComboBox
        Left = 80
        Top = 32
        Width = 185
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 1
        TagDispatch = 0
        DataType = 'TTQUALPIECE'
      end
      object CONTRE_DATE: THCritMaskEdit
        Left = 364
        Top = 56
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 4
        Text = '  /  /    '
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object CONTRE_NEGATIF: TCheckBox
        Left = 438
        Top = 58
        Width = 132
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Contrepasser en &n'#233'gatif'
        TabOrder = 5
      end
      object CONTRE_TYPE: THValComboBox
        Left = 364
        Top = 32
        Width = 207
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        DataType = 'TTQUALPIECE'
      end
    end
    inherited PAvance: THTabSheet
      inherited Bevel4: TBevel
        Width = 584
      end
      inherited bEffaceAvance: TToolbarButton97
        Left = 541
      end
      inherited ZO3: THValComboBox
        Left = 152
      end
      inherited ZO2: THValComboBox
        Left = 152
      end
      inherited ZO1: THValComboBox
        Left = 152
      end
      inherited ZV1: THEdit
        Left = 297
      end
      inherited ZV2: THEdit
        Left = 297
      end
      inherited ZV3: THEdit
        Left = 297
      end
      inherited ZG2: THCombobox
        Left = 497
      end
      inherited ZG1: THCombobox
        Left = 497
      end
    end
    inherited PSQL: THTabSheet
      Caption = '???SQL'
      TabVisible = False
      inherited Bevel3: TBevel
        Width = 584
      end
      inherited Z_SQL: THSQLMemo
        Width = 584
      end
    end
  end
  inherited Dock971: TDock97
    Width = 592
    inherited PFiltres: TToolWindow97
      ClientWidth = 592
      ClientAreaWidth = 592
      inherited BCherche: TToolbarButton97
        Left = 382
      end
      inherited lpresentation: THLabel
        Left = 413
      end
      inherited FFiltres: THValComboBox
        Width = 309
      end
      inherited cbPresentations: THValComboBox
        Left = 484
      end
    end
  end
  inherited FListe: THDBGrid
    Width = 563
    MultiSelection = True
    MultiFieds = 'E_JOURNAL;E_DATECOMPTABLE;E_NUMEROPIECE;'
  end
  inherited Panel2: THPanel
    Left = 458
    inherited PListe: THPanel
      Left = 408
    end
  end
  inherited Dock: TDock97
    Width = 592
    inherited PanelBouton: TToolWindow97
      ClientWidth = 592
      ClientAreaWidth = 592
      inherited bSelectAll: TToolbarButton97
        Left = 164
        Visible = True
      end
      inherited BImprimer: TToolbarButton97
        Left = 464
      end
      inherited BOuvrir: TToolbarButton97
        Left = 496
      end
      inherited BAnnuler: TToolbarButton97
        Left = 528
      end
      inherited BAide: TToolbarButton97
        Left = 560
      end
      inherited Binsert: TToolbarButton97
        Left = 196
      end
      inherited BBlocNote: TToolbarButton97
        Left = 432
      end
      object BListePIECES: TToolbarButton97
        Tag = 1
        Left = 100
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Ecritures d'#39'extourne g'#233'n'#233'r'#233'es'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 0
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BListePIECESClick
        GlobalIndexImage = 'Z1826_S16G1'
        IsControl = True
      end
      object BZoom: TToolbarButton97
        Tag = 1
        Left = 132
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Voir '#233'criture'
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
        OnClick = BZoomClick
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
    end
  end
  inherited PCumul: THPanel
    Width = 592
  end
  inherited TBlocNote: TToolWindow97
    Top = 149
  end
  inherited PanVBar: THPanel
    Left = 563
  end
  inherited SQ: TDataSource
    Left = 64
    Top = 240
  end
  inherited Q: THQuery
    Top = 240
  end
  inherited FindDialog: THFindDialog
    Left = 56
    Top = 189
  end
  inherited POPF: THPopupMenu
    Left = 156
    Top = 240
  end
  inherited HMTrad: THSystemMenu
    Left = 112
    Top = 240
  end
  inherited SD: THSaveDialog
    Left = 16
    Top = 189
  end
  inherited SearchTimer: TTimer
    Top = 189
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Extourne;Confirmez-vous la g'#233'n'#233'ration de l'#39#233'criture d'#39'extourne' +
        ' ?;Q;YNC;Y;Y;'
      
        '1;Extourne;Vous devez s'#233'lectionner au moins une '#233'criture.;E;O;O;' +
        'O;'
      
        '2;Extourne;La date que vous avez renseign'#233'e n'#39'est pas valide;W;O' +
        ';O;O;'
      
        '3;Extourne;La date que vous avez renseign'#233'e n'#39'est pas dans un ex' +
        'ercice ouvert;W;O;O;O;'
      
        '4;Extourne;La date que vous avez renseign'#233'e est ant'#233'rieure '#224' une' +
        ' cl'#244'ture;W;O;O;O;'
      
        '5;Extourne;La date que vous avez renseign'#233'e est ant'#233'rieure '#224' une' +
        ' cl'#244'ture;W;O;O;O;'
      
        '6;Extourne;La date que vous avez renseign'#233'e est en dehors des li' +
        'mites autoris'#233'es;W;O;O;O;'
      'ATTENTION ! G'#233'n'#233'ration non effectu'#233'e ...'
      
        '8;Extourne;Vous ne pouvez pas extourner en '#233'criture courante des' +
        ' '#233'critures de type ;W;O;O;O;'
      'Extourne'
      
        '10;Extourne;Voulez-vous visualiser les '#233'critures g'#233'n'#233'r'#233'es ?;Q;YN' +
        'C;Y;Y;')
    Left = 104
    Top = 189
  end
  object TEcrGen: TQuery
    DatabaseName = 'SOC'
    RequestLive = True
    SQL.Strings = (
      'SELECT * FROM ECRITURE WHERE E_AUXILIAIRE="ZZ#EAe"')
    Left = 224
    Top = 240
  end
  object TEcrAna: TQuery
    DatabaseName = 'SOC'
    RequestLive = True
    SQL.Strings = (
      'SELECT * FROM ANALYTIQ WHERE Y_SECTION="EA#Ewdz"')
    Left = 228
    Top = 189
  end
end

inherited FAssistCom: TFAssistCom
  Left = 190
  Top = 190
  HelpContext = 10002
  Caption = 'Assistant d'#39'envoi ComSx '
  ClientHeight = 401
  ClientWidth = 760
  OnKeyDown = FormKeyDown
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Top = 378
    Visible = False
  end
  inherited lAide: THLabel
    Top = 348
    Width = 572
  end
  inherited bPrecedent: TToolbarButton97
    Left = 515
    Top = 375
  end
  inherited bSuivant: TToolbarButton97
    Left = 598
    Top = 375
  end
  inherited bFin: TToolbarButton97
    Left = 682
    Top = 375
  end
  inherited bAnnuler: TToolbarButton97
    Left = 432
    Top = 375
  end
  inherited bAide: TToolbarButton97
    Left = 349
    Top = 375
    HelpContext = 10002
  end
  object lEtap: THLabel [7]
    Left = 7
    Top = 378
    Width = 68
    Height = 13
    Anchors = [akLeft, akBottom]
    AutoSize = False
    Caption = 'Etape 1/4'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object bClickVisu: TToolbarButton97 [8]
    Left = 318
    Top = 376
    Width = 28
    Height = 23
    Hint = 'Voir le rapport'
    Alignment = taLeftJustify
    Visible = False
    OnClick = bClickVisuClick
    GlobalIndexImage = 'Z0099_S16G1'
  end
  object BFiltre: TToolbarButton97 [9]
    Left = 72
    Top = 375
    Width = 56
    Height = 24
    Hint = 'Menu filtre'
    DropdownArrow = True
    DropdownMenu = POPF
    Caption = 'Filtre'
    Flat = False
    Layout = blGlyphRight
  end
  object BIni: TToolbarButton97 [10]
    Left = 285
    Top = 376
    Width = 28
    Height = 23
    Hint = 'G'#233'n'#233'ration du fichier de commande'
    ParentShowHint = False
    ShowHint = True
    OnClick = BIniClick
    GlobalIndexImage = 'Z0276_S16G1'
  end
  inherited Plan: THPanel
    Top = 32
    Width = 509
    Height = 257
  end
  inherited GroupBox1: THGroupBox
    Top = 365
    Width = 763
  end
  inherited P: THPageControl2
    Left = 175
    Top = 10
    Width = 586
    Height = 327
    ActivePage = TypeEcr
    object Choix: TTabSheet
      Caption = 'Choix'
      object Label7: TLabel
        Left = 159
        Top = 9
        Width = 216
        Height = 13
        Caption = 'Bienvenue dans l'#39'assistant de ComSx '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Bevel2: TBevel
        Left = 159
        Top = 27
        Width = 214
        Height = 8
        Shape = bsTopLine
      end
      object HLabel2: THLabel
        Left = 105
        Top = 98
        Width = 88
        Height = 13
        Caption = 'Nature du transfert'
      end
      object TTYPEFORMAT: THLabel
        Left = 105
        Top = 126
        Width = 32
        Height = 13
        Caption = 'Format'
      end
      object TDATEECR: THLabel
        Left = 105
        Top = 182
        Width = 113
        Height = 13
        Caption = 'Date d'#39'arr'#234't'#233' comptable'
      end
      object THISTOSYNCHRO: THLabel
        Left = 105
        Top = 153
        Width = 115
        Height = 13
        Caption = 'Envoyer synchronisation'
      end
      object Label15: TLabel
        Left = 105
        Top = 206
        Width = 287
        Height = 13
        Caption = 'Aucune '#233'criture ant'#233'rieure '#224' cette p'#233'riode ne sera modifiable'
        FocusControl = FEMail
        Transparent = True
        Visible = False
      end
      object NATURETRANSFERT: THValComboBox
        Left = 236
        Top = 94
        Width = 184
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = NATURETRANSFERTChange
        TagDispatch = 0
        ComboWidth = 300
      end
      object HISTOSYNCHRO: THValComboBox
        Left = 236
        Top = 149
        Width = 183
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
        TagDispatch = 0
      end
      object TYPEFORMAT: THValComboBox
        Left = 236
        Top = 122
        Width = 184
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 1
        Items.Strings = (
          'STANDARD'
          'ETENDU')
        TagDispatch = 0
        Values.Strings = (
          'STD'
          'ETE')
      end
      object DATEECR: THCritMaskEdit
        Left = 236
        Top = 178
        Width = 96
        Height = 21
        EditMask = '!00/00/0000;1;_'
        MaxLength = 10
        TabOrder = 3
        Text = '  /  /    '
        OnChange = DATEECRChange
        OnEnter = DATEECREnter
        OnExit = DATEECRExit
        TagDispatch = 0
        Operateur = Egal
        OpeType = otDate
        ElipsisButton = True
        ElipsisAutoHide = True
        ControlerDate = True
      end
      object BNetExpert: TCheckBox
        Left = 431
        Top = 124
        Width = 55
        Height = 17
        Alignment = taLeftJustify
        Caption = 'ASP'
        TabOrder = 4
        OnClick = BNetExpertClick
      end
      object TRANSFERTVERS: THRadioGroup
        Left = 104
        Top = 38
        Width = 313
        Height = 46
        Caption = 'Transfert vers Cegid'
        Columns = 3
        Items.Strings = (
          'Business Line'
          'Business/Expert'
          'SISCO')
        TabOrder = 5
        OnClick = TRANSFERTVERSClick
        Abrege = False
        Vide = False
        Values.Strings = (
          'S1'
          'S5'
          'SISCO')
      end
      object BGed: TCheckBox
        Left = 347
        Top = 180
        Width = 125
        Height = 17
        Caption = 'Documents associ'#233's'
        TabOrder = 6
      end
    end
    object TypeEcr: TTabSheet
      Caption = 'Type d'#39#233'critures'
      ImageIndex = 1
      object Label9: TLabel
        Left = 159
        Top = 5
        Width = 117
        Height = 13
        Caption = 'Crit'#232'res de s'#233'lection'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Bevel1: TBevel
        Left = 159
        Top = 23
        Width = 120
        Height = 8
        Shape = bsTopLine
      end
      object GRPBAL: TGroupBox
        Left = 256
        Top = 29
        Width = 145
        Height = 63
        Caption = 'Balance'
        TabOrder = 0
        object AUXILIARE: TCheckBox
          Left = 11
          Top = 19
          Width = 97
          Height = 17
          Caption = 'Auxiliaire'
          TabOrder = 0
        end
        object BANA: TCheckBox
          Left = 11
          Top = 40
          Width = 97
          Height = 17
          Caption = 'Analytique'
          TabOrder = 1
        end
      end
      object GRTYPECR: TGroupBox
        Left = 16
        Top = 37
        Width = 228
        Height = 223
        Caption = 'Type de param'#233'trage'
        TabOrder = 1
        object TFTypeEcr: THLabel
          Left = 12
          Top = 76
          Width = 81
          Height = 13
          Caption = 'Type d'#39#233'critures :'
          FocusControl = FTypeEcr
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TExclure: THLabel
          Left = 12
          Top = 199
          Width = 35
          Height = 13
          Caption = 'Exclure'
          FocusControl = Exclure
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object AvecLettrage: TCheckBox
          Left = 12
          Top = 120
          Width = 97
          Height = 17
          Caption = 'Avec le lettrage'
          TabOrder = 1
        end
        object BEXPORT: TCheckBox
          Left = 12
          Top = 138
          Width = 165
          Height = 17
          AllowGrayed = True
          Caption = 'D'#233'j'#224' export'#233's'
          State = cbGrayed
          TabOrder = 2
          OnClick = PARAMDOSClick
        end
        object ANNULVALIDATION: TCheckBox
          Left = 12
          Top = 157
          Width = 172
          Height = 17
          Hint = 
            'Toutes les '#233'critures export'#233'es seront consid'#233'r'#233'es comme non vali' +
            'd'#233'es.'
          Caption = 'D'#233'validation des '#233'critures'
          TabOrder = 3
        end
        object FTypeEcr: THMultiValComboBox
          Left = 12
          Top = 96
          Width = 166
          Height = 21
          TabOrder = 0
          Text = 'N;'
          Abrege = False
          DataType = 'TTQUALPIECE'
          Complete = False
          OuInclusif = False
        end
        object PARAMDOS: TCheckBox
          Left = 12
          Top = 20
          Width = 165
          Height = 17
          Caption = 'Param'#232'tres dossier uniquement'
          TabOrder = 4
          OnClick = PARAMDOSClick
        end
        object PPARAMSOC: TCheckBox
          Left = 12
          Top = 39
          Width = 138
          Height = 17
          Caption = 'Param'#232'tres soci'#233't'#233's'
          Checked = True
          State = cbChecked
          TabOrder = 5
          OnClick = PPARAMSOCClick
        end
        object PEXERCICE: TCheckBox
          Left = 12
          Top = 58
          Width = 138
          Height = 17
          Caption = 'Exercices'
          Checked = True
          State = cbChecked
          TabOrder = 6
          OnClick = PEXERCICEClick
        end
        object PECRBUD: TCheckBox
          Left = 12
          Top = 177
          Width = 157
          Height = 17
          Caption = 'Budget'
          TabOrder = 7
          OnClick = PECRBUDClick
        end
        object Exclure: THMultiValComboBox
          Left = 56
          Top = 195
          Width = 164
          Height = 21
          TabOrder = 8
          Abrege = False
          Values.Strings = (
            'BQC'
            'BQE'
            'RIB'
            'SAT'
            'SSA'
            'ETB'
            'MDP'
            'MDR'
            'DEV'
            'REG'
            'SOU'
            'CORR'
            'TL'
            'REL'
            'CGN'
            'CAE'
            'JAL')
          Items.Strings = (
            'Comptes Bancaires'
            'Banques'
            'RIB'
            'Sections analytiques'
            'Sous-sections analytiques'
            'Etablissements'
            'Mode de paiement'
            'R'#233'glements'
            'Devise'
            'R'#233'gime TVA'
            'Souches'
            'Correspondants'
            'Tables Libres'
            'Relances'
            'G'#233'n'#233'raux'
            'Tiers'
            'Journal')
          Complete = False
          OuInclusif = False
        end
      end
      object PPARAM: TGroupBox
        Left = 256
        Top = 97
        Width = 146
        Height = 152
        Caption = 'Export param'#232'tre'
        TabOrder = 2
        object PGEN: TCheckBox
          Left = 15
          Top = 19
          Width = 97
          Height = 17
          Caption = 'G'#233'n'#233'raux'
          Checked = True
          State = cbChecked
          TabOrder = 0
          OnClick = PGENClick
        end
        object PTIERS: TCheckBox
          Left = 15
          Top = 38
          Width = 106
          Height = 17
          Caption = 'Tiers Comptables'
          Checked = True
          State = cbChecked
          TabOrder = 1
          OnClick = PGENClick
        end
        object PSectionana: TCheckBox
          Left = 15
          Top = 76
          Width = 97
          Height = 17
          Caption = 'Sections'
          Checked = True
          State = cbChecked
          TabOrder = 3
          OnClick = PGENClick
        end
        object PJRL: TCheckBox
          Left = 15
          Top = 94
          Width = 97
          Height = 17
          Caption = 'Journaux'
          Checked = True
          State = cbChecked
          TabOrder = 4
          OnClick = PGENClick
        end
        object PTiersautre: TCheckBox
          Left = 15
          Top = 57
          Width = 106
          Height = 17
          Caption = 'Autres Tiers'
          TabOrder = 2
          OnClick = PTiersautreClick
        end
        object TLibre: TCheckBox
          Left = 15
          Top = 111
          Width = 97
          Height = 17
          Caption = 'Tables libres'
          TabOrder = 5
          OnClick = PGENClick
        end
        object PVentil: TCheckBox
          Left = 15
          Top = 130
          Width = 106
          Height = 17
          Caption = 'Ventilations types'
          TabOrder = 6
          OnClick = PVentilClick
        end
      end
      object PANA: TCheckBox
        Left = 416
        Top = 56
        Width = 97
        Height = 17
        Caption = 'Analytiques'
        Checked = True
        State = cbChecked
        TabOrder = 3
        Visible = False
      end
      object GroupBox3: TGroupBox
        Left = 405
        Top = 148
        Width = 169
        Height = 100
        Caption = 'Modifi'#233's'
        Enabled = False
        TabOrder = 4
        object TDATEMODIF1: TLabel
          Left = 12
          Top = 34
          Width = 12
          Height = 13
          Caption = 'du'
        end
        object TDATEMODIF2: TLabel
          Left = 12
          Top = 71
          Width = 12
          Height = 13
          Caption = 'au'
        end
        object DATEMODIF1: THCritMaskEdit
          Left = 57
          Top = 30
          Width = 105
          Height = 21
          EditMask = '!99/99/0000;1;_'
          MaxLength = 10
          TabOrder = 0
          Text = '01/01/1900'
          TagDispatch = 0
          OpeType = otDate
          ElipsisButton = True
          ElipsisAutoHide = True
          ControlerDate = True
        end
        object DATEMODIF2: THCritMaskEdit
          Left = 57
          Top = 67
          Width = 105
          Height = 21
          EditMask = '!99/99/0000;1;_'
          MaxLength = 10
          TabOrder = 1
          Text = '01/01/2099'
          TagDispatch = 0
          OpeType = otDate
          ElipsisButton = True
          ElipsisAutoHide = True
          ControlerDate = True
        end
      end
      object GComptes: TGroupBox
        Left = 405
        Top = 29
        Width = 168
        Height = 114
        Caption = ' Comptes '
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 5
        object TCPTEDEBUT: THLabel
          Left = 11
          Top = 61
          Width = 12
          Height = 13
          Caption = 'du'
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TCPTEFIN: THLabel
          Left = 11
          Top = 86
          Width = 12
          Height = 13
          Caption = 'au'
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HNature: THLabel
          Left = 11
          Top = 29
          Width = 32
          Height = 13
          Caption = 'Nature'
          FocusControl = Exclure
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object CPTEDEBUT: THCritMaskEdit
          Left = 54
          Top = 57
          Width = 105
          Height = 21
          Enabled = False
          TabOrder = 0
          TagDispatch = 0
          DataType = 'TZGENERAL'
          ElipsisButton = True
          OnElipsisClick = CPTEDEBUTElipsisClick
        end
        object CPTEFIN: THCritMaskEdit
          Left = 54
          Top = 82
          Width = 105
          Height = 21
          Enabled = False
          TabOrder = 1
          TagDispatch = 0
          DataType = 'TZGENERAL'
          ElipsisButton = True
          OnElipsisClick = CPTEDEBUTElipsisClick
        end
        object TNature: THMultiValComboBox
          Left = 54
          Top = 27
          Width = 103
          Height = 21
          TabOrder = 2
          Text = '<<Tous>>'
          Abrege = False
          DataType = 'TTNATTIERS'
          Complete = False
          OuInclusif = False
        end
      end
      object GroupBox4: TGroupBox
        Left = 16
        Top = 259
        Width = 553
        Height = 41
        Caption = 'Options avanc'#233'es'
        TabOrder = 6
        object Z_C1: THValComboBox
          Left = 160
          Top = 14
          Width = 148
          Height = 21
          Style = csDropDownList
          Ctl3D = False
          ItemHeight = 13
          ParentCtl3D = False
          TabOrder = 0
          TagDispatch = 0
        end
        object ZO1: THValComboBox
          Left = 312
          Top = 14
          Width = 78
          Height = 21
          Style = csDropDownList
          Ctl3D = False
          ItemHeight = 13
          ParentCtl3D = False
          TabOrder = 1
          TagDispatch = 0
          DataType = 'TTCOMPARE'
        end
        object ZV1: TEdit
          Left = 392
          Top = 14
          Width = 153
          Height = 21
          Ctl3D = True
          ParentCtl3D = False
          TabOrder = 2
        end
        object CBLib: TCheckBox
          Left = 2
          Top = 16
          Width = 150
          Height = 17
          Caption = 'Voir les libell'#233's des champs'
          TabOrder = 3
          OnClick = CBLibClick
        end
      end
    end
    object PPARAMGENE: TTabSheet
      Caption = 'Crit'#232'res de s'#233'lection'
      ImageIndex = 4
      object TEXERCICE: THLabel
        Left = 6
        Top = 42
        Width = 41
        Height = 13
        Caption = '&Exercice'
      end
      object TNATUREJRL: THLabel
        Left = 6
        Top = 70
        Width = 81
        Height = 13
        Caption = '&Nature du journal'
      end
      object TJOURNAUX: THLabel
        Left = 6
        Top = 99
        Width = 43
        Height = 13
        Caption = '&Journaux'
      end
      object Label1: TLabel
        Left = 159
        Top = 161
        Width = 110
        Height = 13
        Caption = 'Choix de la p'#233'riode'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object TDATEECR1: THLabel
        Left = 95
        Top = 195
        Width = 14
        Height = 13
        Caption = 'Du'
        FocusControl = DATEECR1
      end
      object TDATEECR2: THLabel
        Left = 225
        Top = 195
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = DATEECR2
      end
      object Bevel8: TBevel
        Left = 159
        Top = 178
        Width = 110
        Height = 8
        Shape = bsTopLine
      end
      object TFETAB: THLabel
        Left = 6
        Top = 129
        Width = 65
        Height = 13
        Caption = 'E&tablissement'
      end
      object Label8: TLabel
        Left = 159
        Top = 5
        Width = 117
        Height = 13
        Caption = 'Crit'#232'res de s'#233'lection'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Bevel4: TBevel
        Left = 159
        Top = 23
        Width = 121
        Height = 8
        Shape = bsTopLine
      end
      object EXERCICE: THValComboBox
        Left = 120
        Top = 38
        Width = 257
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 0
        OnChange = EXERCICEChange
        TagDispatch = 0
      end
      object NATUREJRL: THValComboBox
        Left = 120
        Top = 66
        Width = 257
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 1
        OnClick = NATUREJRLClick
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATJAL'
      end
      object JOURNAUX: THMultiValComboBox
        Left = 120
        Top = 94
        Width = 257
        Height = 21
        TabOrder = 2
        Text = '<<Tous>>'
        Abrege = False
        Values.Strings = (
          'xx')
        Items.Strings = (
          'xx')
        Complete = False
        OuInclusif = False
      end
      object DATEECR1: THCritMaskEdit
        Left = 120
        Top = 191
        Width = 96
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 3
        Text = '01/01/1900'
        TagDispatch = 0
        OpeType = otDate
        ElipsisButton = True
        ElipsisAutoHide = True
        ControlerDate = True
      end
      object DATEECR2: THCritMaskEdit
        Left = 243
        Top = 191
        Width = 96
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 4
        Text = '01/01/2099'
        OnExit = DATEECR2Exit
        TagDispatch = 0
        OpeType = otDate
        ElipsisButton = True
        ElipsisAutoHide = True
        ControlerDate = True
      end
      object FETAB: THMultiValComboBox
        Left = 120
        Top = 125
        Width = 257
        Height = 21
        TabOrder = 5
        Text = '<<Tous>>'
        Abrege = False
        DataType = 'TTETABLISSEMENT'
        Complete = False
        OuInclusif = False
      end
      object GESTIONETAB: TCheckBox
        Left = 380
        Top = 127
        Width = 199
        Height = 17
        Caption = 'Etablissement des '#233'critures'
        Checked = True
        State = cbChecked
        TabOrder = 6
        OnClick = GESTIONETABClick
      end
    end
    object PARAMSISCOII: TTabSheet
      Caption = 'Param'#232'trage SISCOII'
      ImageIndex = 5
      object HLabel1: THLabel
        Left = 6
        Top = 16
        Width = 115
        Height = 13
        Caption = 'Traitement du code stat.'
      end
      object TDATEBAL: THLabel
        Left = 14
        Top = 223
        Width = 128
        Height = 13
        Caption = 'Date d'#39'arr'#234't'#233' de la balance'
      end
      object CP_STAT: THValComboBox
        Left = 137
        Top = 12
        Width = 184
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        Items.Strings = (
          'Pas d'#39'export'
          'Export table libre mouvement'
          'Export analytique')
        TagDispatch = 0
        Values.Strings = (
          'RIE'
          'TL'
          'ANA')
      end
      object GroupBox2: TGroupBox
        Left = 1
        Top = 40
        Width = 494
        Height = 49
        TabOrder = 1
        object Label4: TLabel
          Left = 237
          Top = 24
          Width = 12
          Height = 13
          Caption = 'de'
        end
        object Label5: TLabel
          Left = 342
          Top = 24
          Width = 6
          Height = 13
          Caption = #224
        end
        object AJOUTAUX: TToolbarButton97
          Left = 440
          Top = 19
          Width = 23
          Height = 22
          OnClick = AJOUTAUXClick
          GlobalIndexImage = 'Z0519_S16G1'
        end
        object BDelete1: TToolbarButton97
          Left = 466
          Top = 19
          Width = 23
          Height = 22
          OnClick = BDelete1Click
          GlobalIndexImage = 'Z0005_S16G1'
        end
        object COLLECTIFS: THValComboBox
          Left = 9
          Top = 20
          Width = 95
          Height = 21
          ItemHeight = 13
          TabOrder = 0
          OnChange = COLLECTIFSChange
          Items.Strings = (
            'Fournisseur'
            'Client')
          TagDispatch = 0
          Values.Strings = (
            'F'
            'C')
        end
        object COLLECT: THCritMaskEdit
          Left = 110
          Top = 20
          Width = 81
          Height = 21
          EditMask = '40aaaaaaaa'
          MaxLength = 10
          TabOrder = 1
          Text = '40        '
          OnExit = COLLECTExit
          TagDispatch = 0
          ElipsisButton = True
          ElipsisAutoHide = True
          OnElipsisClick = COLLECTElipsisClick
          ControlerDate = True
        end
        object CARAUX: THValComboBox
          Left = 195
          Top = 20
          Width = 40
          Height = 21
          ItemHeight = 13
          TabOrder = 2
          OnChange = CARAUXChange
          Items.Strings = (
            '0'
            'F'
            '9'
            'C')
          TagDispatch = 0
          Values.Strings = (
            '0'
            'F'
            '9'
            'C')
        end
        object AUX1: THCritMaskEdit
          Left = 254
          Top = 20
          Width = 81
          Height = 21
          EditMask = '0aaaaaaaaa'
          MaxLength = 10
          TabOrder = 3
          Text = '0         '
          OnClick = AUX1Click
          OnExit = AUX1Exit
          TagDispatch = 0
          ControlerDate = True
        end
        object AUX2: THCritMaskEdit
          Left = 355
          Top = 20
          Width = 81
          Height = 21
          EditMask = '0aaaaaaaaa'
          MaxLength = 10
          TabOrder = 4
          Text = '0         '
          OnClick = AUX2Click
          OnExit = AUX2Exit
          TagDispatch = 0
          ControlerDate = True
        end
        object Quadra: TCheckBox
          Left = 160
          Top = 0
          Width = 145
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Transfert vers Quadra'
          TabOrder = 5
          OnClick = SANCOLLClick
        end
      end
      object SANCOLL: TCheckBox
        Left = 8
        Top = 40
        Width = 145
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Comptabilit'#233' sans auxiliaire'
        TabOrder = 2
        OnClick = SANCOLLClick
      end
      object THGRID: THGrid
        Left = 14
        Top = 96
        Width = 465
        Height = 121
        ColCount = 3
        FixedCols = 0
        RowCount = 2
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
        TabOrder = 3
        SortedCol = -1
        Titres.Strings = (
          'Collectifs'
          'De'
          'A')
        Couleur = False
        MultiSelect = False
        TitleBold = False
        TitleCenter = True
        OnColEnter = THGRIDColEnter
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = clSilver
        ColWidths = (
          146
          132
          179)
      end
      object SuppCarAuxi: TCheckBox
        Left = 324
        Top = 14
        Width = 237
        Height = 17
        Caption = 'Suppression caract'#232're de bourrage auxiliaires'
        TabOrder = 4
        OnClick = SuppCarAuxiClick
      end
      object DATARRETEBAL: THCritMaskEdit
        Left = 152
        Top = 219
        Width = 96
        Height = 21
        EditMask = '!00/00/0000;1;_'
        MaxLength = 10
        TabOrder = 5
        Text = '  /  /    '
        OnChange = DATARRETEBALChange
        OnEnter = DATARRETEBALEnter
        OnExit = DATARRETEBALExit
        TagDispatch = 0
        Operateur = Egal
        OpeType = otDate
        ElipsisButton = True
        ElipsisAutoHide = True
        ControlerDate = True
      end
    end
    object Mail: TTabSheet
      Caption = 'Mail'
      ImageIndex = 2
      object Label2: TLabel
        Left = 37
        Top = 86
        Width = 3
        Height = 13
        FocusControl = FEMail
        Transparent = True
      end
      object Label10: TLabel
        Left = 159
        Top = 10
        Width = 122
        Height = 13
        Caption = 'Envoi par messagerie'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Bevel5: TBevel
        Left = 159
        Top = 27
        Width = 122
        Height = 8
        Shape = bsTopLine
      end
      object Label11: TLabel
        Left = 36
        Top = 222
        Width = 3
        Height = 13
        FocusControl = FEMail
        Transparent = True
      end
      object Label12: TLabel
        Left = 37
        Top = 112
        Width = 3
        Height = 13
        FocusControl = FEMail
        Transparent = True
      end
      object Label3: TLabel
        Left = 37
        Top = 58
        Width = 3
        Height = 13
        Transparent = True
      end
      object HLabel3: THLabel
        Left = 40
        Top = 32
        Width = 117
        Height = 13
        Caption = 'Envoyer les donn'#233'es par'
      end
      object HLabel4: THLabel
        Left = 40
        Top = 56
        Width = 127
        Height = 13
        Caption = 'Nom du fichier d'#39#233'change :'
      end
      object HLabel5: THLabel
        Left = 40
        Top = 88
        Width = 109
        Height = 13
        Caption = 'Adresse de messagerie'
      end
      object HLabel6: THLabel
        Left = 40
        Top = 112
        Width = 93
        Height = 13
        Caption = 'Corps du message :'
      end
      object HLabel7: THLabel
        Left = 39
        Top = 224
        Width = 413
        Height = 13
        Caption = 
          'ATTENTION : le fichier d'#39#233'change sera envoy'#233' sous forme de pi'#232'ce' +
          ' jointe au message.'
      end
      object FEMail: THCritMaskEdit
        Left = 169
        Top = 82
        Width = 359
        Height = 21
        TabOrder = 0
        Text = ' '
        TagDispatch = 0
        ControlerDate = True
      end
      object FHigh: TCheckBox
        Left = 169
        Top = 112
        Width = 182
        Height = 12
        Caption = 'Message avec importance haute'
        TabOrder = 1
      end
      object FFile: TRadioButton
        Left = 167
        Top = 33
        Width = 55
        Height = 17
        Caption = 'Fichier'
        Checked = True
        TabOrder = 2
        TabStop = True
        OnClick = FFileClick
      end
      object FMail: TRadioButton
        Left = 226
        Top = 33
        Width = 76
        Height = 17
        Caption = 'Messagerie'
        TabOrder = 3
        OnClick = FFileClick
      end
      object Fmonexpert: TRadioButton
        Left = 301
        Top = 33
        Width = 148
        Height = 17
        Caption = 'monexpertcomptable.com'
        TabOrder = 4
        Visible = False
      end
      object FICHENAME: THCritMaskEdit
        Left = 169
        Top = 55
        Width = 361
        Height = 21
        TabOrder = 5
        OnExit = FICHENAMEExit
        TagDispatch = 0
        DataType = 'SAVEFILE(*.TRA;SI*.TRT)'
        ElipsisButton = True
      end
      object ZIPPE: TCheckBox
        Left = 310
        Top = 33
        Width = 145
        Height = 17
        Caption = 'Donn'#233'es compress'#233'es'
        TabOrder = 6
      end
      object FCorpsMail: THMemo
        Left = 40
        Top = 136
        Width = 473
        Height = 85
        Lines.Strings = (
          'FCorpsMail')
        TabOrder = 7
      end
    end
    object Resume: TTabSheet
      Caption = 'Resume'
      ImageIndex = 6
      object Label13: TLabel
        Left = 193
        Top = 12
        Width = 46
        Height = 13
        Caption = 'R'#233'sum'#233
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Bevel6: TBevel
        Left = 195
        Top = 29
        Width = 44
        Height = 8
        Shape = bsTopLine
      end
      object Bevel3: TBevel
        Left = 29
        Top = 56
        Width = 502
        Height = 135
      end
      object FLib1: TLabel
        Left = 32
        Top = 66
        Width = 65
        Height = 13
        AutoSize = False
        Caption = 'Traitement'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object FVal1: TLabel
        Left = 100
        Top = 66
        Width = 342
        Height = 26
        AutoSize = False
        Caption = 'Envoi des '#233'critures '#224' l'#39'expert comptable'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        WordWrap = True
      end
      object FLib2: TLabel
        Left = 32
        Top = 90
        Width = 65
        Height = 13
        AutoSize = False
        Caption = 'P'#233'riode'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object FVal2: TLabel
        Left = 100
        Top = 90
        Width = 342
        Height = 13
        AutoSize = False
        Caption = 'jusqu'#39'au 01/01/99 (avec arr'#234't'#233' p'#233'riodique)'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object FLib3: TLabel
        Left = 32
        Top = 114
        Width = 65
        Height = 13
        AutoSize = False
        Caption = 'Fichier'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object FVal3: TLabel
        Left = 100
        Top = 114
        Width = 342
        Height = 13
        AutoSize = False
        Caption = 'C:\sfdoedf\dfijhdif.pgi'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object FLib4: TLabel
        Left = 32
        Top = 138
        Width = 65
        Height = 13
        AutoSize = False
        Caption = 'Messagerie'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object FVal4: TLabel
        Left = 100
        Top = 138
        Width = 342
        Height = 13
        AutoSize = False
        Caption = 'envoi '#224' l'#39'adresse : jdupont@expert-club.fr'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
    end
    object PEXPORT: TTabSheet
      Caption = 'Compte rendu'
      ImageIndex = 7
      object Label14: TLabel
        Left = 215
        Top = 0
        Width = 79
        Height = 13
        Caption = 'Compte rendu'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Bevel7: TBevel
        Left = 215
        Top = 17
        Width = 82
        Height = 8
        Shape = bsTopLine
      end
      object LISTEEXPORT: TListBox
        Left = 0
        Top = 24
        Width = 572
        Height = 217
        ItemHeight = 13
        TabOrder = 0
      end
    end
  end
  object Dock971: TDock97 [16]
    Left = 0
    Top = 0
    Width = 760
    Height = 9
    AllowDrag = False
  end
  object FFiltres: THValComboBox [17]
    Left = 130
    Top = 377
    Width = 150
    Height = 21
    Hint = 'Rappel des Sc'#233'narios'
    Style = csDropDownList
    Anchors = [akLeft, akTop, akRight]
    ItemHeight = 13
    TabOrder = 6
    TagDispatch = 0
  end
  object FTimer: TTimer
    Enabled = False
    OnTimer = FTimerTimer
    Left = 40
    Top = 248
  end
  object PrinterSetupDialog1: TPrinterSetupDialog
    Left = 121
    Top = 122
  end
  object POPF: TPopupMenu
    Left = 60
    Top = 236
    object BCreerFiltre: TMenuItem
      Caption = '&Cr'#233'er un filtre'
      Enabled = False
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

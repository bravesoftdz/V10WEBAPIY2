inherited FAssistExcelOle: TFAssistExcelOle
  Left = 441
  Top = 117
  Caption = 'Assistant OLE'
  PixelsPerInch = 96
  TextHeight = 13
  inherited bPrecedent: TToolbarButton97
    Enabled = False
  end
  inherited bSuivant: TToolbarButton97
    Enabled = False
  end
  inherited bFin: TToolbarButton97
    ModalResult = 1
  end
  inherited P: THPageControl2
    Top = 5
    ActivePage = TBCaract
    object TBCaract: TTabSheet
      Caption = 'Caract'#233'ristiques'
      object TMULTIDOSSIER: TLabel
        Left = 7
        Top = 193
        Width = 70
        Height = 13
        Caption = 'Regroupement'
        Visible = False
      end
      object Bevel1: TBevel
        Left = 8
        Top = 154
        Width = 321
        Height = 1
      end
      object TSelectionSociete: TLabel
        Left = 7
        Top = 217
        Width = 41
        Height = 13
        Caption = 'Soci'#233't'#233's'
        Visible = False
      end
      object RBBilansocial: TRadioButton
        Left = 44
        Top = 17
        Width = 261
        Height = 17
        Caption = 'Bilan social et tableaux de bord (PaieBSIndOLE)'
        TabOrder = 1
        OnClick = RBBilansocialClick
      end
      object RBBilanSocialSimple: TRadioButton
        Left = 44
        Top = 44
        Width = 261
        Height = 17
        Caption = 'Bilan social simplifi'#233' (PaieBSCalInfo)'
        TabOrder = 2
        OnClick = RBBilanSocialSimpleClick
      end
      object RBCumulPaie: TRadioButton
        Left = 44
        Top = 71
        Width = 261
        Height = 17
        Caption = 'Donn'#233'es historique de paie (PaieCumul)'
        TabOrder = 3
        OnClick = RBCumulPaieClick
      end
      object RBEffectif: TRadioButton
        Left = 44
        Top = 98
        Width = 261
        Height = 17
        Caption = 'Effectif (PaieEffectif)'
        TabOrder = 4
        OnClick = RBEffectifClick
      end
      object RBPaieInfo: TRadioButton
        Left = 44
        Top = 125
        Width = 261
        Height = 17
        Caption = 'Information paie (PaieRendInfo)'
        TabOrder = 5
        OnClick = RBPaieInfoClick
      end
      object MULTIDOSSIER: THValComboBox
        Left = 117
        Top = 190
        Width = 217
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        Visible = False
        OnSelect = MULTIDOSSIERSelect
        TagDispatch = 0
        Vide = True
        VideString = '<Aucun>'
        DataType = 'YYMULTIDOSSIER'
      end
      object RBChoixMultiSOc: TRadioButton
        Left = 44
        Top = 164
        Width = 261
        Height = 17
        Caption = 'Choisir un regroupement mutli-soci'#233't'#233
        TabOrder = 6
        OnClick = RBChoixMultiSOcClick
      end
      object SELECTIONSOCIETE: THMultiValComboBox
        Left = 117
        Top = 214
        Width = 217
        Height = 21
        TabOrder = 7
        Visible = False
        Abrege = False
        Complete = False
        OuInclusif = False
      end
    end
    object TBBilanSocial: TTabSheet
      Caption = 'TBBilanSocial'
      ImageIndex = 1
      object TPBC_INDICATEURBS: THLabel
        Left = 4
        Top = 84
        Width = 47
        Height = 13
        Caption = 'Indicateur'
        FocusControl = PBC_INDICATEURBS
      end
      object TPBC_BSPRESENTATION: THLabel
        Left = 4
        Top = 60
        Width = 59
        Height = 13
        Caption = 'Pr'#233'sentation'
        FocusControl = PBC_BSPRESENTATION
      end
      object TPBC_DATEDEBUT: THLabel
        Left = 4
        Top = 12
        Width = 68
        Height = 13
        Caption = 'Date de d'#233'but'
        FocusControl = PBC_DATEDEBUT
      end
      object TPBC_DATEFIN: THLabel
        Left = 4
        Top = 36
        Width = 52
        Height = 13
        Caption = 'Date de fin'
        FocusControl = PBC_DATEFIN
      end
      object TPBC_ETABLISSEMENT: THLabel
        Left = 4
        Top = 114
        Width = 65
        Height = 13
        Caption = 'Etablissement'
        FocusControl = PBC_ETABLISSEMENT
      end
      object TPBC_CATBILAN: THLabel
        Left = 4
        Top = 138
        Width = 45
        Height = 13
        Caption = 'Cat'#233'gorie'
        FocusControl = PBC_CATBILAN
      end
      object TPBC_LIBELLEEMPLOI: THLabel
        Left = 4
        Top = 168
        Width = 31
        Height = 13
        Caption = 'Emploi'
        FocusControl = PBC_LIBELLEEMPLOI
        Visible = False
      end
      object TPBC_QUALIFICATION: THLabel
        Left = 4
        Top = 216
        Width = 58
        Height = 13
        Caption = 'Qualification'
        FocusControl = PBC_QUALIFICATION
        Visible = False
      end
      object TPBC_COEFFICIENT: THLabel
        Left = 4
        Top = 192
        Width = 50
        Height = 13
        Caption = 'Coefficient'
        FocusControl = PBC_COEFFICIENT
        Visible = False
      end
      object PBC_ETABLISSEMENT: THMultiValComboBox
        Left = 99
        Top = 110
        Width = 238
        Height = 21
        TabOrder = 4
        Abrege = False
        DataType = 'TTETABLISSEMENT'
        Complete = False
        OuInclusif = False
      end
      object PBC_INDICATEURBS: THValComboBox
        Left = 99
        Top = 80
        Width = 238
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 3
        TagDispatch = 0
        DataType = 'PGBSINDIC'
        ComboWidth = 350
      end
      object PBC_BSPRESENTATION: THValComboBox
        Left = 99
        Top = 56
        Width = 238
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 2
        OnChange = PBC_BSPRESENTATIONChange
        TagDispatch = 0
        DataType = 'PGBSINDSELECT'
      end
      object PBC_DATEDEBUT: THCritMaskEdit
        Left = 99
        Top = 8
        Width = 238
        Height = 21
        TabOrder = 0
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        DefaultDate = odDebAnnee
        OnElipsisClick = PgParamDate
        ControlerDate = True
        DisplayFormat = 'dd mmm yyyy'
      end
      object PBC_DATEFIN: THCritMaskEdit
        Left = 99
        Top = 32
        Width = 238
        Height = 21
        TabOrder = 1
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        DefaultDate = odFinAnnee
        OnElipsisClick = PgParamDate
        ControlerDate = True
        DisplayFormat = 'dd mmm yyyy'
      end
      object PBC_CATBILAN: THMultiValComboBox
        Left = 99
        Top = 134
        Width = 238
        Height = 21
        TabOrder = 5
        Abrege = False
        DataType = 'PGCATBILAN'
        Complete = False
        OuInclusif = False
      end
      object PBC_LIBELLEEMPLOI: THMultiValComboBox
        Left = 99
        Top = 164
        Width = 238
        Height = 21
        TabOrder = 6
        Visible = False
        Abrege = False
        DataType = 'PGLIBEMPLOI'
        Complete = False
        OuInclusif = False
      end
      object PBC_COEFFICIENT: THMultiValComboBox
        Left = 99
        Top = 188
        Width = 238
        Height = 21
        TabOrder = 7
        Visible = False
        Abrege = False
        DataType = 'PGLIBCOEFFICIENT'
        Complete = False
        OuInclusif = False
      end
      object PBC_QUALIFICATION: THMultiValComboBox
        Left = 99
        Top = 212
        Width = 238
        Height = 21
        TabOrder = 8
        Visible = False
        Abrege = False
        DataType = 'PGLIBQUALIFICATION'
        Complete = False
        OuInclusif = False
      end
    end
    object TBPAIECUMUL: TTabSheet
      Caption = 'TBPAIECUMUL'
      ImageIndex = 3
      object HDATEDEB: THLabel
        Left = 4
        Top = 12
        Width = 68
        Height = 13
        Caption = 'Date de d'#233'but'
        FocusControl = DATEDEB
      end
      object HDATEFIN: THLabel
        Left = 4
        Top = 36
        Width = 52
        Height = 13
        Caption = 'Date de fin'
        FocusControl = DATEFIN
      end
      object HEtab: TLabel
        Left = 4
        Top = 60
        Width = 65
        Height = 13
        Caption = 'Etablissement'
      end
      object HSAL: TLabel
        Left = 4
        Top = 84
        Width = 32
        Height = 13
        Caption = 'Salari'#233
        FocusControl = SALARIE
      end
      object HTYPEDONNEE: TLabel
        Left = 4
        Top = 108
        Width = 83
        Height = 13
        Caption = 'Type de donn'#233'es'
        FocusControl = TYPEDONNEE
      end
      object HTYPEVALEUR: TLabel
        Left = 4
        Top = 132
        Width = 71
        Height = 13
        Caption = 'Type de valeur'
        FocusControl = TYPEVALEUR
      end
      object HRUBRIQUE: TLabel
        Left = 4
        Top = 156
        Width = 76
        Height = 13
        Caption = 'Rubrique/cumul'
      end
      object HCATDADS: TLabel
        Left = 4
        Top = 180
        Width = 78
        Height = 13
        Caption = 'Cat'#233'gorie DUCS'
      end
      object HDADSPROF: TLabel
        Left = 4
        Top = 204
        Width = 93
        Height = 13
        Caption = 'Statut professionnel'
      end
      object DATEDEB: THCritMaskEdit
        Left = 99
        Top = 8
        Width = 238
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 0
        Text = '  /  /    '
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        DefaultDate = odDebAnnee
        OnElipsisClick = PgParamDate
        ControlerDate = True
        DisplayFormat = 'dd mmm yyyy'
      end
      object DATEFIN: THCritMaskEdit
        Left = 99
        Top = 32
        Width = 237
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 1
        Text = '  /  /    '
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        DefaultDate = odFinAnnee
        OnElipsisClick = PgParamDate
        ControlerDate = True
        DisplayFormat = 'dd mmm yyyy'
      end
      object SALARIE: THCritMaskEdit
        Left = 99
        Top = 80
        Width = 237
        Height = 21
        MaxLength = 10
        TabOrder = 3
        OnExit = SALARIEExit
        TagDispatch = 0
        DataType = 'PGSALARIE'
        ElipsisButton = True
      end
      object TYPEDONNEE: THValComboBox
        Left = 99
        Top = 104
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 4
        Items.Strings = (
          'Valeur d'#39'un cumul'
          'Valeur d'#39'une cotisation'
          'Valeur d'#39'une r'#233'mun'#233'ration'
          'Nombre de salari'#233's')
        TagDispatch = 0
        Values.Strings = (
          'CUM'
          'COT'
          'REM'
          'PAI')
      end
      object TYPEVALEUR: THValComboBox
        Left = 99
        Top = 128
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 5
        OnChange = TYPEVALEURChange
        Items.Strings = (
          'Montant d'#39'un cumul'
          'Base d'#39'une r'#233'mun'#233'ration'
          'Montant d'#39'une r'#233'mun'#233'ration'
          'Base d'#39'une cotisation'
          'Montant salarial d'#39'une cotisation'
          'Montant patronal d'#39'une cotisation'
          'Effectif r'#233'el'
          'Effectif moyen')
        TagDispatch = 0
        Values.Strings = (
          'MC'
          'BR'
          'MR'
          'BC'
          'MS'
          'MP'
          'EFF'
          'EFM')
      end
      object Present: TCheckBox
        Left = 99
        Top = 224
        Width = 234
        Height = 17
        Caption = 'exclure les salari'#233's sortis'
        TabOrder = 8
      end
      object ETABLISSEMENT: THMultiValComboBox
        Left = 99
        Top = 56
        Width = 237
        Height = 21
        TabOrder = 2
        Abrege = False
        DataType = 'TTETABLISSEMENT'
        Complete = False
        OuInclusif = False
      end
      object RUBRIQUE: THMultiValComboBox
        Left = 99
        Top = 152
        Width = 237
        Height = 21
        TabOrder = 9
        Abrege = False
        Complete = False
        OuInclusif = False
      end
      object CATDADS: THMultiValComboBox
        Left = 99
        Top = 176
        Width = 237
        Height = 21
        TabOrder = 6
        Abrege = False
        DataType = 'PGCODECATEGORIE'
        Complete = False
        OuInclusif = False
      end
      object DADSPROF: THMultiValComboBox
        Left = 99
        Top = 200
        Width = 237
        Height = 21
        TabOrder = 7
        Abrege = False
        DataType = 'PGSPROFESSIONNEL'
        Complete = False
        OuInclusif = False
      end
    end
    object TBPaieBSCalInfo: TTabSheet
      Caption = 'TBPaieBSCalInfo'
      ImageIndex = 5
      object HTYPERENS: TLabel
        Left = 4
        Top = 12
        Width = 83
        Height = 13
        Caption = 'Type de donn'#233'es'
        FocusControl = TYPERENS
      end
      object HTYPEVAL1: TLabel
        Left = 4
        Top = 36
        Width = 80
        Height = 13
        Caption = 'Type de valeur 1'
        FocusControl = TYPEVAL1
      end
      object HDATED: THLabel
        Left = 5
        Top = 114
        Width = 68
        Height = 13
        Caption = 'Date de d'#233'but'
        FocusControl = DATED
      end
      object HDATEF: THLabel
        Left = 5
        Top = 138
        Width = 52
        Height = 13
        Caption = 'Date de fin'
        FocusControl = DATEF
      end
      object HCATDUCS: TLabel
        Left = 5
        Top = 168
        Width = 73
        Height = 13
        Caption = 'Cat'#233'gorie Ducs'
        FocusControl = CAT
      end
      object HVAL1: TLabel
        Left = 5
        Top = 60
        Width = 39
        Height = 13
        Caption = 'Valeur 1'
        FocusControl = VAL1
      end
      object HVAL2: TLabel
        Left = 5
        Top = 84
        Width = 39
        Height = 13
        Caption = 'Valeur 2'
        FocusControl = VAL2
      end
      object TYPERENS: THValComboBox
        Left = 99
        Top = 8
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = TYPERENSChange
        Items.Strings = (
          'Condition d'#39'emploi'
          'Sexe du salari'#233
          'Nationalit'#233' du salari'#233
          'Age du salari'#233
          'Anciennet'#233' du salari'#233
          'Motifs d'#39'absence du salari'#233
          'Num'#233'ro du cumul')
        TagDispatch = 0
        Values.Strings = (
          'EMP'
          'SEX'
          'NAT'
          'AGE'
          'ANC'
          'ABS'
          'CUM')
      end
      object TYPEVAL1: THValComboBox
        Left = 99
        Top = 32
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 1
        OnChange = TYPEVALEURChange
        TagDispatch = 0
        DataType = 'PGCONDEMPLOI'
      end
      object DATED: THCritMaskEdit
        Left = 99
        Top = 110
        Width = 237
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 2
        Text = '  /  /    '
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        DefaultDate = odDebAnnee
        OnElipsisClick = PgParamDate
        ControlerDate = True
        DisplayFormat = 'dd mmm yyyy'
      end
      object DATEF: THCritMaskEdit
        Left = 99
        Top = 134
        Width = 237
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 3
        Text = '  /  /    '
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        DefaultDate = odFinAnnee
        OnElipsisClick = PgParamDate
        ControlerDate = True
        DisplayFormat = 'dd mmm yyyy'
      end
      object CAT: THCritMaskEdit
        Left = 99
        Top = 164
        Width = 237
        Height = 21
        MaxLength = 4
        TabOrder = 4
        TagDispatch = 0
        DataType = 'PGCODECATEGORIE2'
        ElipsisButton = True
      end
      object VAL1: THCritMaskEdit
        Left = 99
        Top = 56
        Width = 237
        Height = 21
        Enabled = False
        MaxLength = 4
        TabOrder = 5
        TagDispatch = 0
      end
      object VAL2: THCritMaskEdit
        Left = 99
        Top = 80
        Width = 237
        Height = 21
        Enabled = False
        MaxLength = 4
        TabOrder = 6
        TagDispatch = 0
      end
    end
    object TBPaieRendInfo: TTabSheet
      Caption = 'TBPaieRendInfo'
      ImageIndex = 6
      object HTYPETABLE: TLabel
        Left = 4
        Top = 12
        Width = 27
        Height = 13
        Caption = 'Table'
        FocusControl = TYPETABLE
      end
      object HSALARIE2: TLabel
        Left = 4
        Top = 36
        Width = 32
        Height = 13
        Caption = 'Salari'#233
        FocusControl = SALARIE2
      end
      object Label1: TLabel
        Left = 4
        Top = 60
        Width = 33
        Height = 13
        Caption = 'Champ'
        FocusControl = CHAMP
      end
      object HQUOI: TLabel
        Left = 4
        Top = 84
        Width = 22
        Height = 13
        Caption = 'Quoi'
        FocusControl = QUOI
      end
      object TYPETABLE: THValComboBox
        Left = 99
        Top = 8
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = TYPERENSChange
        Items.Strings = (
          'Salari'#233's')
        TagDispatch = 0
        Values.Strings = (
          'SALARIES')
      end
      object SALARIE2: THCritMaskEdit
        Left = 99
        Top = 32
        Width = 237
        Height = 21
        MaxLength = 10
        TabOrder = 1
        OnExit = SALARIEExit
        TagDispatch = 0
        DataType = 'PGSALARIE'
        ElipsisButton = True
      end
      object CHAMP: THValComboBox
        Left = 99
        Top = 56
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 2
        OnChange = TYPEVALEURChange
        TagDispatch = 0
        DataType = 'PGCHAMPSAL'
      end
      object QUOI: THValComboBox
        Left = 99
        Top = 80
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        OnChange = TYPEVALEURChange
        Items.Strings = (
          'Libell'#233
          'Code + Libell'#233)
        TagDispatch = 0
        Values.Strings = (
          '$'
          '$$')
      end
    end
    object TBPaieEffectif: TTabSheet
      Caption = 'TBPaieEffectif'
      ImageIndex = 7
      object HLabel2: THLabel
        Left = 4
        Top = 12
        Width = 51
        Height = 13
        Caption = 'Mensualit'#233
        FocusControl = MENSUALITE
      end
      object Label2: TLabel
        Left = 4
        Top = 36
        Width = 65
        Height = 13
        Caption = 'Etablissement'
      end
      object HPSEXE: THLabel
        Left = 4
        Top = 60
        Width = 24
        Height = 13
        Caption = 'Sexe'
      end
      object HPTYPECONTRAT: THLabel
        Left = 4
        Top = 90
        Width = 34
        Height = 13
        Caption = 'Contrat'
      end
      object HPENTSORT: THLabel
        Left = 4
        Top = 114
        Width = 63
        Height = 13
        Caption = 'Entr'#233'e/Sortie'
        FocusControl = PENTSORT
      end
      object HPMOTIFCONT: THLabel
        Left = 4
        Top = 138
        Width = 89
        Height = 13
        Caption = 'Motif Entr'#233'e/Sortie'
        FocusControl = PMOTIFCONT
      end
      object HPSOMEFF: THLabel
        Left = 4
        Top = 162
        Width = 35
        Height = 13
        Caption = 'Somme'
        FocusControl = PSOMEFF
      end
      object MENSUALITE: THCritMaskEdit
        Left = 99
        Top = 8
        Width = 237
        Height = 21
        EditMask = '!99/0000;1;_'
        MaxLength = 7
        TabOrder = 0
        Text = '  /    '
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        DefaultDate = odFinAnnee
        OnElipsisClick = PgParamDate
        ControlerDate = True
        DisplayFormat = 'mmm yyyy'
      end
      object PETABLISSEMENT: THMultiValComboBox
        Left = 99
        Top = 32
        Width = 237
        Height = 21
        TabOrder = 1
        Abrege = False
        DataType = 'TTETABLISSEMENT'
        Complete = False
        OuInclusif = False
      end
      object PSEXE: THValComboBox
        Left = 99
        Top = 56
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 2
        TagDispatch = 0
        Vide = True
        VideString = '<<Tous>>'
        DataType = 'PGSEXE'
      end
      object PTYPECONTRAT: THValComboBox
        Left = 99
        Top = 86
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 3
        OnChange = PTYPECONTRATChange
        TagDispatch = 0
        Vide = True
        VideString = '<<Tous>>'
        DataType = 'PGTYPECONTRAT'
      end
      object PENTSORT: THValComboBox
        Left = 99
        Top = 110
        Width = 237
        Height = 21
        Style = csDropDownList
        Enabled = False
        ItemHeight = 13
        TabOrder = 4
        OnChange = PENTSORTChange
        Items.Strings = (
          '<<Aucun>>'
          'Motif d'#39'entr'#233'e'
          'Motif de sortie')
        TagDispatch = 0
        Values.Strings = (
          ''
          'E'
          'S')
      end
      object PMOTIFCONT: THValComboBox
        Left = 99
        Top = 134
        Width = 237
        Height = 21
        Style = csDropDownList
        Enabled = False
        ItemHeight = 13
        TabOrder = 5
        Items.Strings = (
          'Motif d'#39'entr'#233'e'
          'Motif de sortie')
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'PGMOTIFENTREELIGHT'
        Values.Strings = (
          'E'
          'S')
      end
      object PSOMEFF: THValComboBox
        Left = 99
        Top = 158
        Width = 237
        Height = 21
        Style = csDropDownList
        Enabled = False
        ItemHeight = 13
        TabOrder = 6
        Items.Strings = (
          'Motif d'#39'entr'#233'e'
          'Motif de sortie')
        TagDispatch = 0
        DataType = 'PGOUINON'
        Values.Strings = (
          'E'
          'S')
      end
    end
    object TBPaieEffectifCompl: TTabSheet
      Caption = 'TBPaieEffectifCompl'
      ImageIndex = 8
      object HLIBELLEEMPLOI: THLabel
        Left = 4
        Top = 90
        Width = 31
        Height = 13
        Caption = 'Emploi'
        FocusControl = PLIBELLEEMPLOI
      end
      object HPCOEFFICIENT: THLabel
        Left = 4
        Top = 114
        Width = 50
        Height = 13
        Caption = 'Coefficient'
        FocusControl = PCOEFFICIENT
      end
      object HPQUALIFICATION: THLabel
        Left = 4
        Top = 138
        Width = 58
        Height = 13
        Caption = 'Qualification'
        FocusControl = PQUALIFICATION
      end
      object HPDADSCAT: TLabel
        Left = 4
        Top = 60
        Width = 77
        Height = 13
        Caption = 'Statut cat'#233'goriel'
        FocusControl = PDADSCAT
      end
      object HPDADSPROF: TLabel
        Left = 4
        Top = 36
        Width = 93
        Height = 13
        Caption = 'Statut professionnel'
        FocusControl = PDADSPROF
      end
      object HPCATDUCS: TLabel
        Left = 4
        Top = 12
        Width = 78
        Height = 13
        Caption = 'Cat'#233'gorie DUCS'
        FocusControl = PCATDUCS
      end
      object HPNIVEAU: THLabel
        Left = 4
        Top = 162
        Width = 34
        Height = 13
        Caption = 'Niveau'
        FocusControl = PNIVEAU
      end
      object HPINDICE: THLabel
        Left = 4
        Top = 186
        Width = 29
        Height = 13
        Caption = 'Indice'
        FocusControl = PINDICE
      end
      object HPCODEEMPLOI: THLabel
        Left = 4
        Top = 210
        Width = 90
        Height = 13
        Caption = 'Nomenclature PCS'
        FocusControl = PCODEEMPLOI
      end
      object PLIBELLEEMPLOI: THMultiValComboBox
        Left = 99
        Top = 86
        Width = 237
        Height = 21
        TabOrder = 3
        Abrege = False
        DataType = 'PGLIBEMPLOI'
        Complete = False
        OuInclusif = False
      end
      object PCOEFFICIENT: THMultiValComboBox
        Left = 99
        Top = 110
        Width = 237
        Height = 21
        TabOrder = 4
        Abrege = False
        DataType = 'PGLIBCOEFFICIENT'
        Complete = False
        OuInclusif = False
      end
      object PQUALIFICATION: THMultiValComboBox
        Left = 99
        Top = 134
        Width = 237
        Height = 21
        TabOrder = 5
        Abrege = False
        DataType = 'PGLIBQUALIFICATION'
        Complete = False
        OuInclusif = False
      end
      object PDADSCAT: THMultiValComboBox
        Left = 99
        Top = 56
        Width = 237
        Height = 21
        TabOrder = 2
        Abrege = False
        DataType = 'PGSCATEGORIEL'
        Complete = False
        OuInclusif = False
      end
      object PDADSPROF: THMultiValComboBox
        Left = 99
        Top = 32
        Width = 237
        Height = 21
        TabOrder = 1
        Abrege = False
        DataType = 'PGSPROFESSIONNEL'
        Complete = False
        OuInclusif = False
      end
      object PCATDUCS: THMultiValComboBox
        Left = 99
        Top = 8
        Width = 237
        Height = 21
        TabOrder = 0
        Abrege = False
        DataType = 'PGCODECATEGORIE'
        Complete = False
        OuInclusif = False
      end
      object PNIVEAU: THMultiValComboBox
        Left = 99
        Top = 158
        Width = 237
        Height = 21
        TabOrder = 6
        Abrege = False
        DataType = 'PGLIBNIVEAU'
        Complete = False
        OuInclusif = False
      end
      object PINDICE: THMultiValComboBox
        Left = 99
        Top = 182
        Width = 237
        Height = 21
        TabOrder = 7
        Abrege = False
        DataType = 'PGLIBINDICE'
        Complete = False
        OuInclusif = False
      end
      object PCODEEMPLOI: THMultiValComboBox
        Left = 99
        Top = 206
        Width = 237
        Height = 21
        TabOrder = 8
        Abrege = False
        DataType = 'PGCODEEMPLOI'
        Complete = False
        OuInclusif = False
      end
    end
    object TBComplCombo: TTabSheet
      Caption = 'TBComplCombo'
      ImageIndex = 4
      object TPSA_CODESTAT: THLabel
        Left = 12
        Top = 12
        Width = 75
        Height = 13
        Caption = 'Code statistique'
        FocusControl = PSA_CODESTAT
      end
      object TPSA_TRAVAILN1: THLabel
        Left = 12
        Top = 40
        Width = 46
        Height = 13
        Caption = 'TravailN1'
        FocusControl = PSA_TRAVAILN1
      end
      object TPSA_TRAVAILN2: THLabel
        Left = 12
        Top = 64
        Width = 46
        Height = 13
        Caption = 'TravailN2'
        FocusControl = PSA_TRAVAILN2
      end
      object TPSA_TRAVAILN3: THLabel
        Left = 12
        Top = 88
        Width = 46
        Height = 13
        Caption = 'TravailN3'
        FocusControl = PSA_TRAVAILN3
      end
      object TPSA_TRAVAILN4: THLabel
        Left = 12
        Top = 112
        Width = 46
        Height = 13
        Caption = 'TravailN4'
        FocusControl = PSA_TRAVAILN4
      end
      object TPSA_LIBREPCMB1: THLabel
        Left = 12
        Top = 142
        Width = 64
        Height = 13
        Caption = 'Tablettelibre1'
        FocusControl = PSA_LIBREPCMB1
      end
      object TPSA_LIBREPCMB2: THLabel
        Left = 12
        Top = 166
        Width = 64
        Height = 13
        Caption = 'Tablettelibre2'
        FocusControl = PSA_LIBREPCMB2
      end
      object TPSA_LIBREPCMB3: THLabel
        Left = 12
        Top = 190
        Width = 64
        Height = 13
        Caption = 'Tablettelibre3'
        FocusControl = PSA_LIBREPCMB3
      end
      object TPSA_LIBREPCMB4: THLabel
        Left = 12
        Top = 214
        Width = 64
        Height = 13
        Caption = 'Tablettelibre4'
        FocusControl = PSA_LIBREPCMB4
      end
      object PSA_CODESTAT: THValComboBox
        Left = 95
        Top = 8
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'PGCODESTAT'
      end
      object PSA_TRAVAILN1: THValComboBox
        Left = 95
        Top = 36
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 1
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'PGTRAVAILN1'
      end
      object PSA_TRAVAILN2: THValComboBox
        Left = 95
        Top = 60
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 2
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'PGTRAVAILN2'
      end
      object PSA_TRAVAILN3: THValComboBox
        Left = 95
        Top = 84
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 3
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'PGTRAVAILN3'
      end
      object PSA_TRAVAILN4: THValComboBox
        Left = 95
        Top = 108
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 4
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'PGTRAVAILN4'
      end
      object PSA_LIBREPCMB1: THValComboBox
        Left = 95
        Top = 138
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 5
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'PGLIBREPCMB1'
      end
      object PSA_LIBREPCMB2: THValComboBox
        Left = 95
        Top = 162
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 6
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'PGLIBREPCMB2'
      end
      object PSA_LIBREPCMB3: THValComboBox
        Left = 95
        Top = 186
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 7
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'PGLIBREPCMB3'
      end
      object PSA_LIBREPCMB4: THValComboBox
        Left = 95
        Top = 210
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 8
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'PGLIBREPCMB4'
      end
    end
    object TBComplMulti: TTabSheet
      Caption = 'TBComplMulti'
      ImageIndex = 2
      object TMCODESTAT: THLabel
        Left = 12
        Top = 12
        Width = 75
        Height = 13
        Caption = 'Code statistique'
        FocusControl = MCODESTAT
      end
      object TMTRAVAILN1: THLabel
        Left = 12
        Top = 42
        Width = 46
        Height = 13
        Caption = 'TravailN1'
        FocusControl = MTRAVAILN1
      end
      object TMTRAVAILN2: THLabel
        Left = 12
        Top = 66
        Width = 46
        Height = 13
        Caption = 'TravailN2'
        FocusControl = MTRAVAILN2
      end
      object TMTRAVAILN3: THLabel
        Left = 12
        Top = 90
        Width = 46
        Height = 13
        Caption = 'TravailN3'
        FocusControl = MTRAVAILN3
      end
      object TMTRAVAILN4: THLabel
        Left = 12
        Top = 114
        Width = 46
        Height = 13
        Caption = 'TravailN4'
        FocusControl = MTRAVAILN4
      end
      object TMLIBREPCMB1: THLabel
        Left = 12
        Top = 144
        Width = 64
        Height = 13
        Caption = 'Tablettelibre1'
        FocusControl = MLIBREPCMB1
      end
      object TMLIBREPCMB2: THLabel
        Left = 12
        Top = 168
        Width = 64
        Height = 13
        Caption = 'Tablettelibre2'
        FocusControl = MLIBREPCMB2
      end
      object TMLIBREPCMB3: THLabel
        Left = 12
        Top = 192
        Width = 64
        Height = 13
        Caption = 'Tablettelibre3'
        FocusControl = MLIBREPCMB3
      end
      object TMLIBREPCMB4: THLabel
        Left = 12
        Top = 216
        Width = 64
        Height = 13
        Caption = 'Tablettelibre4'
        FocusControl = MLIBREPCMB4
      end
      object MCODESTAT: THMultiValComboBox
        Left = 88
        Top = 8
        Width = 237
        Height = 21
        TabOrder = 0
        Abrege = False
        DataType = 'PGCODESTAT'
        Complete = False
        OuInclusif = False
      end
      object MTRAVAILN1: THMultiValComboBox
        Left = 88
        Top = 38
        Width = 237
        Height = 21
        TabOrder = 1
        Abrege = False
        DataType = 'PGTRAVAILN1'
        Complete = False
        OuInclusif = False
      end
      object MTRAVAILN2: THMultiValComboBox
        Left = 88
        Top = 62
        Width = 237
        Height = 21
        TabOrder = 2
        Abrege = False
        DataType = 'PGTRAVAILN2'
        Complete = False
        OuInclusif = False
      end
      object MTRAVAILN3: THMultiValComboBox
        Left = 88
        Top = 86
        Width = 237
        Height = 21
        TabOrder = 3
        Abrege = False
        DataType = 'PGTRAVAILN3'
        Complete = False
        OuInclusif = False
      end
      object MTRAVAILN4: THMultiValComboBox
        Left = 88
        Top = 110
        Width = 237
        Height = 21
        TabOrder = 4
        Abrege = False
        DataType = 'PGTRAVAILN4'
        Complete = False
        OuInclusif = False
      end
      object MLIBREPCMB1: THMultiValComboBox
        Left = 88
        Top = 140
        Width = 237
        Height = 21
        TabOrder = 5
        Abrege = False
        DataType = 'PGLIBREPCMB1'
        Complete = False
        OuInclusif = False
      end
      object MLIBREPCMB2: THMultiValComboBox
        Left = 88
        Top = 164
        Width = 237
        Height = 21
        TabOrder = 6
        Abrege = False
        DataType = 'PGLIBREPCMB2'
        Complete = False
        OuInclusif = False
      end
      object MLIBREPCMB3: THMultiValComboBox
        Left = 88
        Top = 188
        Width = 237
        Height = 21
        TabOrder = 7
        Abrege = False
        DataType = 'PGLIBREPCMB3'
        Complete = False
        OuInclusif = False
      end
      object MLIBREPCMB4: THMultiValComboBox
        Left = 88
        Top = 212
        Width = 237
        Height = 21
        TabOrder = 8
        Abrege = False
        DataType = 'PGLIBREPCMB4'
        Complete = False
        OuInclusif = False
      end
    end
  end
end

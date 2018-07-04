inherited FAssistEF: TFAssistEF
  Left = 306
  Top = 185
  Caption = 'FAssistEF'
  ClientHeight = 408
  ClientWidth = 546
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Top = 383
  end
  inherited lAide: THLabel
    Left = 172
    Top = 353
    Width = 374
  end
  object lRub: THLabel [2]
    Left = 225
    Top = 5
    Width = 294
    Height = 15
    AutoSize = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object lSource: TLabel [3]
    Left = 174
    Top = 5
    Width = 49
    Height = 15
    AutoSize = False
  end
  inherited bPrecedent: TToolbarButton97
    Left = 295
    Top = 378
  end
  inherited bSuivant: TToolbarButton97
    Left = 378
    Top = 378
  end
  inherited bFin: TToolbarButton97
    Left = 462
    Top = 378
    Enabled = False
  end
  inherited bAnnuler: TToolbarButton97
    Left = 212
    Top = 378
  end
  inherited bAide: TToolbarButton97
    Left = 129
    Top = 376
  end
  inherited Plan: THPanel
    Left = 174
    Top = 30
  end
  inherited GroupBox1: TGroupBox
    Top = 365
    Width = 565
  end
  inherited P: TPageControl
    Left = 170
    Top = 27
    Height = 318
    ActivePage = Criteres
    OnEnter = FlagExtEnter
    object ChoixFormule: TTabSheet
      Caption = 'ChoixFormule'
      object Label2: TLabel
        Left = 16
        Top = 2
        Width = 309
        Height = 13
        AutoSize = False
        Caption = 'Bienvenue dans l'#39'assistant'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label3: TLabel
        Left = 43
        Top = 20
        Width = 250
        Height = 13
        AutoSize = False
        Caption = 'Quel type de donn'#233'es souha'#238'tez-vous extraire ?'
      end
      object HLabel2: THLabel
        Left = 67
        Top = 121
        Width = 262
        Height = 39
        AutoSize = False
        Caption = 
          'Dans ce cas, vous pourrez extraire une information issue d'#39'une f' +
          'iche. Par exemple, le libell'#233' d'#39'un compte ou l'#39'adresse d'#39'un tier' +
          's.'
        WordWrap = True
      end
      object HLabel4: THLabel
        Left = 67
        Top = 57
        Width = 262
        Height = 43
        AutoSize = False
        Caption = 
          'Dans ce cas, vous pouvez extraire une information calcul'#233'e et d'#233 +
          'finir des crit'#232'res personnalis'#233's. Par exemple, vous obtiendrez l' +
          'e cumul d'#39'une rubrique.'
        WordWrap = True
      end
      object HLabel10: THLabel
        Left = 67
        Top = 189
        Width = 262
        Height = 27
        AutoSize = False
        Caption = 
          'Dans ce cas, vous pourrez extraire une constante comme par exemp' +
          'le le nombre de salari'#233's.'
        WordWrap = True
      end
      object bVal: THSpeedButton
        Left = 173
        Top = 166
        Width = 42
        Height = 20
        Hint = 'Extrait la valeur d'#39'une constante'
        GroupIndex = 1
        Caption = 'Valeur'
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = bValClick
      end
      object bLib: THSpeedButton
        Left = 215
        Top = 166
        Width = 42
        Height = 20
        Hint = 'Extrait le libell'#233' d'#39'une constante'
        GroupIndex = 1
        Caption = 'Libell'#233
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = bLibClick
      end
      object Regroupement: TLabel
        Left = 43
        Top = 244
        Width = 134
        Height = 13
        Caption = 'Regroupement multi soci'#233't'#233' '
      end
      object rCumul: TRadioButton
        Left = 66
        Top = 38
        Width = 100
        Height = 14
        Caption = 'un &cumul'
        Checked = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
        TabStop = True
        OnClick = rCumulClick
      end
      object rChamp: TRadioButton
        Left = 65
        Top = 100
        Width = 100
        Height = 17
        Caption = 'un c&hamp'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 1
        OnClick = rChampClick
      end
      object rConstante: TRadioButton
        Left = 65
        Top = 168
        Width = 106
        Height = 17
        Caption = 'une cons&tante'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 2
        OnClick = rConstanteClick
      end
      object RSQL: TRadioButton
        Left = 65
        Top = 224
        Width = 128
        Height = 17
        Caption = 'Une requ'#232'te SQL'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 3
        OnClick = RSQLClick
      end
      object MULTIDOSSIER: THValComboBox
        Left = 65
        Top = 262
        Width = 184
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 4
        TagDispatch = 0
        Vide = True
        DataType = 'YYMULTIDOSSIER'
      end
    end
    object ChoixSource: TTabSheet
      Caption = 'ChoixSource'
      object Label5: TLabel
        Left = 18
        Top = -2
        Width = 165
        Height = 13
        Caption = 'Choisissez la source des donn'#233'es :'
      end
      object lBudget: TLabel
        Left = 36
        Top = 126
        Width = 205
        Height = 13
        Caption = 'Calculer les informations budg'#233'taires sur le :'
      end
      object rSource: TRadioGroup
        Left = 17
        Top = 9
        Width = 297
        Height = 114
        Columns = 2
        ItemIndex = 0
        Items.Strings = (
          'Rubriques'
          'Rubriques budg'#233'taires'
          'Comptes g'#233'n'#233'raux'
          'Tiers'
          'Sections analytiques'
          'Comptes budg'#233'taires'
          'Sections budg'#233'taires'
          'Journaux'
          'Budgets'
          'Soci'#233't'#233
          'G'#233'n'#233'raux/Tiers'
          'G'#233'n'#233'raux/Sections'
          'Comptes/Sections Budget')
        TabOrder = 1
        OnClick = rSourceClick
        OnEnter = rSourceEnter
      end
      object rBudget: TRadioGroup
        Left = 17
        Top = 144
        Width = 297
        Height = 39
        Columns = 2
        ItemIndex = 0
        Items.Strings = (
          'Budget'#233
          'R'#233'alis'#233)
        TabOrder = 2
        OnClick = rBudgetClick
      end
      object gAxe: TGroupBox
        Left = 17
        Top = 216
        Width = 297
        Height = 37
        TabOrder = 5
        object lAxe: THLabel
          Left = 54
          Top = 15
          Width = 73
          Height = 13
          Caption = 'Choisissez l'#39'axe'
        end
        object cAxe: THValComboBox
          Left = 134
          Top = 11
          Width = 76
          Height = 21
          Style = csDropDownList
          ItemHeight = 0
          TabOrder = 0
          OnChange = cAxeChange
          OnEnter = cAxeEnter
          TagDispatch = 0
          DataType = 'TTAXE'
        end
      end
      object GSens: TGroupBox
        Left = 17
        Top = 137
        Width = 297
        Height = 42
        TabOrder = 3
        Visible = False
        object LSens: THLabel
          Left = 34
          Top = 18
          Width = 85
          Height = 13
          Caption = 'Choisissez le sens'
        end
        object CSens: THValComboBox
          Left = 126
          Top = 14
          Width = 113
          Height = 21
          Style = csDropDownList
          ItemHeight = 0
          TabOrder = 0
          TagDispatch = 0
          DataType = 'TTRUBCALCUL'
        end
      end
      object CMonnaie: TRadioGroup
        Left = 17
        Top = 183
        Width = 297
        Height = 35
        Caption = 'Restitution en ... '
        Columns = 2
        ItemIndex = 0
        Items.Strings = (
          'Monnaie de tenue'
          'Contrevaleur')
        TabOrder = 4
      end
      object GC2: TCheckBox
        Left = 190
        Top = -2
        Width = 97
        Height = 17
        Caption = 'Get_Cumul2'
        TabOrder = 0
      end
      object GBalSit: TGroupBox
        Left = 17
        Top = 253
        Width = 297
        Height = 37
        TabOrder = 6
        object CCBalSit: TCheckBox
          Left = 96
          Top = 12
          Width = 97
          Height = 17
          Caption = 'En situation'
          TabOrder = 0
          OnClick = CCBalSitClick
        end
      end
    end
    object PSQL: TTabSheet
      Caption = 'SQL'
      ImageIndex = 12
      object GroupBox2: TGroupBox
        Left = 10
        Top = 16
        Width = 309
        Height = 175
        Caption = 'SQL '
        TabOrder = 0
        object Z_SQL: THSQLMemo
          Left = 2
          Top = 15
          Width = 305
          Height = 158
          Align = alClient
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'Courier New'
          Font.Style = []
          ParentFont = False
          ScrollBars = ssVertical
          TabOrder = 0
          ColorRules = <>
        end
      end
      object FlagExtSQL: TCheckBox
        Left = 9
        Top = 222
        Width = 120
        Height = 17
        Caption = '&R'#233'f'#233'rence de cellule'
        TabOrder = 1
        OnClick = FlagExtClick
        OnEnter = FlagExtEnter
      end
      object RefCelSQL: TEdit
        Left = 133
        Top = 219
        Width = 128
        Height = 21
        Enabled = False
        TabOrder = 2
        OnChange = RefCelRecChange
        OnEnter = RefCelEnter
      end
    end
    object Budget: TTabSheet
      Caption = 'Budget'
      object Label7: TLabel
        Left = 44
        Top = 60
        Width = 34
        Height = 13
        Caption = '&Budget'
      end
      object cBudget: THValComboBox
        Left = 44
        Top = 78
        Width = 253
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 0
        OnEnter = cDeviseEnter
        TagDispatch = 0
        DataType = 'TTBUDJAL'
      end
      object FlagExtBud: TCheckBox
        Left = 44
        Top = 213
        Width = 120
        Height = 17
        Caption = '&R'#233'f'#233'rence de cellule'
        TabOrder = 1
        OnClick = FlagExtClick
        OnEnter = FlagExtEnter
      end
      object RefCelBud: TEdit
        Left = 168
        Top = 210
        Width = 126
        Height = 21
        Enabled = False
        TabOrder = 2
        OnEnter = RefCelEnter
      end
    end
    object CptVariant: TTabSheet
      Caption = 'Compte(s) "variant(s)"'
      ImageIndex = 10
      object Label8: TLabel
        Left = 44
        Top = 42
        Width = 203
        Height = 13
        Caption = 'Param'#232'tre variant de rubriques composites '
      end
      object Label9: TLabel
        Left = 53
        Top = 63
        Width = 158
        Height = 13
        Caption = '(Comptes, racines ou fourchettes)'
      end
      object FlagExtVariant: TCheckBox
        Left = 44
        Top = 213
        Width = 120
        Height = 17
        Caption = '&R'#233'f'#233'rence de cellule'
        TabOrder = 0
        OnClick = FlagExtClick
        OnEnter = FlagExtEnter
      end
      object RefCellVariant: TEdit
        Left = 168
        Top = 210
        Width = 126
        Height = 21
        Enabled = False
        TabOrder = 1
        OnEnter = RefCelEnter
      end
      object cCptVariant: TEdit
        Left = 44
        Top = 92
        Width = 229
        Height = 21
        TabOrder = 2
        OnEnter = cCptVariantEnter
      end
    end
    object ChoixRec: TTabSheet
      Caption = 'ChoixRec'
      object GridRub: THDBGrid
        Left = 9
        Top = 5
        Width = 320
        Height = 183
        DataSource = DS
        Options = [dgTitles, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgAlwaysShowSelection]
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'MS Sans Serif'
        TitleFont.Style = []
        OnDblClick = GridRubDblClick
        OnEnter = GridRubEnter
        Row = 1
        MultiSelection = False
        SortEnabled = False
        MyDefaultRowHeight = 0
        Columns = <
          item
            Expanded = False
            FieldName = 'CODE'
            Title.Caption = 'Code'
            Width = 70
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'LIBELLE'
            Title.Caption = 'Libell'#233
            Width = 230
            Visible = True
          end>
      end
      object FindCode: TEdit
        Left = 9
        Top = 191
        Width = 72
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 1
        OnChange = FindCodeChange
        OnEnter = FindCodeEnter
      end
      object FindLib: TEdit
        Left = 84
        Top = 191
        Width = 227
        Height = 21
        TabOrder = 2
        OnChange = FindLibChange
        OnEnter = FindLibEnter
      end
      object FlagExtRec: TCheckBox
        Left = 9
        Top = 222
        Width = 120
        Height = 17
        Caption = '&R'#233'f'#233'rence de cellule'
        TabOrder = 3
        OnClick = FlagExtClick
        OnEnter = FlagExtEnter
      end
      object RefCelRec: TEdit
        Left = 133
        Top = 219
        Width = 128
        Height = 21
        Enabled = False
        TabOrder = 4
        OnChange = RefCelRecChange
        OnEnter = RefCelEnter
      end
    end
    object ChoixBalSit: TTabSheet
      Caption = 'ChoixBalSit'
      ImageIndex = 14
      object Label10: TLabel
        Left = 44
        Top = 60
        Width = 39
        Height = 13
        Caption = '&Balance'
      end
      object CBalSit: THValComboBox
        Left = 44
        Top = 78
        Width = 253
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 0
        OnEnter = cEtablissementEnter
        TagDispatch = 0
        Vide = True
        DataType = 'CPBALSIT'
      end
      object FlagExtBalSit: TCheckBox
        Left = 44
        Top = 213
        Width = 120
        Height = 17
        Caption = '&R'#233'f'#233'rence de cellule'
        TabOrder = 1
        OnClick = FlagExtClick
      end
      object RefCelBalSit: TEdit
        Left = 168
        Top = 210
        Width = 126
        Height = 21
        Enabled = False
        TabOrder = 2
        OnEnter = RefCelEnter
      end
    end
    object ChoixRec2: TTabSheet
      Caption = 'ChoixRec2'
      ImageIndex = 11
      object FindCode2: TEdit
        Tag = 1
        Left = 9
        Top = 191
        Width = 72
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 0
        OnChange = FindCode2Change
        OnEnter = FindCodeEnter
      end
      object GridRub2: THDBGrid
        Left = 9
        Top = 5
        Width = 320
        Height = 183
        DataSource = DS2
        Options = [dgTitles, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgAlwaysShowSelection]
        TabOrder = 1
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'MS Sans Serif'
        TitleFont.Style = []
        OnDblClick = GridRubDblClick
        OnEnter = GridRubEnter
        Row = 1
        MultiSelection = False
        SortEnabled = False
        MyDefaultRowHeight = 0
        Columns = <
          item
            Expanded = False
            FieldName = 'CODE'
            Title.Caption = 'Code'
            Width = 70
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'LIBELLE'
            Title.Caption = 'Libell'#233
            Width = 230
            Visible = True
          end>
      end
      object FindLib2: TEdit
        Tag = 1
        Left = 84
        Top = 191
        Width = 227
        Height = 21
        TabOrder = 2
        OnChange = FindLib2Change
        OnEnter = FindLibEnter
      end
      object FlagExtRec2: TCheckBox
        Tag = 1
        Left = 9
        Top = 222
        Width = 120
        Height = 17
        Caption = '&R'#233'f'#233'rence de cellule'
        TabOrder = 3
        OnClick = FlagExtClick
        OnEnter = FlagExtEnter
      end
      object RefCelRec2: TEdit
        Tag = 1
        Left = 133
        Top = 219
        Width = 128
        Height = 21
        Enabled = False
        TabOrder = 4
        OnChange = RefCelRecChange
        OnEnter = RefCelEnter
      end
    end
    object Periode: TTabSheet
      Caption = 'Periode'
      object LExercice: THLabel
        Left = 35
        Top = 45
        Width = 41
        Height = 13
        Caption = '&Exercice'
      end
      object HLabel7: THLabel
        Left = 35
        Top = 90
        Width = 36
        Height = 13
        Caption = '&P'#233'riode'
      end
      object lDetailPeriode: THLabel
        Left = 35
        Top = 134
        Width = 63
        Height = 13
        Caption = '&Compl'#233'ments'
        Visible = False
      end
      object LNumperiode1: THLabel
        Left = 195
        Top = 110
        Width = 12
        Height = 13
        AutoSize = False
        Caption = '&'#224
        FocusControl = NumPeriode1
      end
      object cNumPeriode: THValComboBox
        Left = 163
        Top = 106
        Width = 142
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 4
        Visible = False
        OnClick = cNumPeriodeClick
        TagDispatch = 0
        Values.Strings = (
          'N'
          'M'
          'B'
          'T'
          'Q'
          'S')
      end
      object cExercice: THValComboBox
        Left = 35
        Top = 62
        Width = 270
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 1
        OnEnter = cExerciceEnter
        Items.Strings = (
          'Exercice courant'
          'Exercice N-1'
          'Exercice N-2'
          'Exercice N-3'
          'Exercice N+1')
        TagDispatch = 0
        Values.Strings = (
          ''
          '-'
          '--'
          '---'
          '+')
      end
      object cPeriode: THValComboBox
        Left = 35
        Top = 106
        Width = 124
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
        OnChange = cPeriodeChange
        OnEnter = cPeriodeEnter
        Items.Strings = (
          'Exercice entier'
          'Mois'
          'Bimestre'
          'Trimestre'
          'Quadrimestre'
          'Semestre'
          'Semaine')
        TagDispatch = 0
        Values.Strings = (
          'N'
          'M'
          'B'
          'T'
          'Q'
          'S'
          'W')
      end
      object cDetailPeriode: THValComboBox
        Left = 35
        Top = 151
        Width = 270
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 7
        Visible = False
        OnEnter = cDetailPeriodeEnter
        Items.Strings = (
          'Uniquement la p'#233'riode s'#233'lectionn'#233'e'
          'Du d'#233'but de l'#39'exercice jusqu'#39#224' la p'#233'riode (incluse)'
          'Du d'#233'but de la p'#233'riode jusqu'#39#224' la fin de l'#39'exercice')
        TagDispatch = 0
        Values.Strings = (
          ''
          '<'
          '>')
      end
      object NumPeriode: THNumEdit
        Left = 163
        Top = 106
        Width = 26
        Height = 21
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Min = 1.000000000000000000
        Debit = False
        TabOrder = 3
        UseRounding = True
        Value = 1.000000000000000000
        Validate = False
        Visible = False
        OnChange = NumPeriodeChange
        OnEnter = NumPeriodeEnter
        OnExit = NumPeriodeExit
      end
      object CheckExoOK: TCheckBox
        Left = 35
        Top = 14
        Width = 259
        Height = 17
        Caption = 'Les exercices sont cal'#233's sur l'#39'ann'#233'e civile'
        Checked = True
        State = cbChecked
        TabOrder = 0
        OnClick = CheckExoOKClick
      end
      object FlagExtPer: TCheckBox
        Left = 35
        Top = 213
        Width = 120
        Height = 17
        Caption = '&R'#233'f'#233'rence de cellule'
        TabOrder = 8
        OnClick = FlagExtClick
        OnEnter = FlagExtEnter
      end
      object RefCelPer: TEdit
        Left = 159
        Top = 210
        Width = 128
        Height = 21
        Enabled = False
        TabOrder = 9
        OnEnter = RefCelEnter
      end
      object NumPeriode1: THNumEdit
        Left = 209
        Top = 106
        Width = 26
        Height = 21
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Min = 1.000000000000000000
        Debit = False
        TabOrder = 5
        UseRounding = True
        Value = 1.000000000000000000
        Validate = False
        Visible = False
        OnChange = NumPeriodeChange
        OnEnter = NumPeriodeEnter
        OnExit = NumPeriode1Exit
      end
      object cAnnee: THValComboBox
        Left = 245
        Top = 106
        Width = 58
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 6
        Items.Strings = (
          '2010'
          '2009'
          '2008'
          '2007'
          '2006'
          '2005'
          '2004'
          '2003'
          '2002'
          '2001'
          '2000'
          '1999'
          '1998'
          '1997'
          '1996'
          '1995')
        TagDispatch = 0
        Values.Strings = (
          '2010'
          '2009'
          '2008'
          '2007'
          '2006'
          '2005'
          '2004'
          '2003'
          '2002'
          '2001'
          '2000'
          '1999'
          '1998'
          '1997'
          '1996'
          '1995')
      end
    end
    object Criteres: TTabSheet
      Caption = 'Criteres'
      object HLabel5: THLabel
        Left = 43
        Top = 36
        Width = 113
        Height = 13
        AutoSize = False
        Caption = '&Type d'#39#233'critures'
        FocusControl = cNatureEcr
      end
      object HLabel3: THLabel
        Left = 43
        Top = 82
        Width = 169
        Height = 13
        AutoSize = False
        Caption = '&Int'#233'gration des A-nouveaux'
        FocusControl = cNatureEcr
      end
      object cNatureEcr: THValComboBox
        Left = 43
        Top = 53
        Width = 253
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 0
        OnEnter = cNatureEcrEnter
        TagDispatch = 0
        DataType = 'TTQUALPIECECRIT'
      end
      object cIntegAN: THValComboBox
        Left = 43
        Top = 100
        Width = 253
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 1
        OnEnter = cIntegANEnter
        Items.Strings = (
          'Avec les A-nouveaux'
          'Sans les A-nouveaux'
          'Que les A-nouveaux')
        TagDispatch = 0
        Values.Strings = (
          ''
          '-'
          '#')
      end
      object FlagExtEcr: TCheckBox
        Left = 42
        Top = 213
        Width = 120
        Height = 17
        Caption = '&R'#233'f'#233'rence de cellule'
        TabOrder = 4
        OnClick = FlagExtClick
        OnEnter = FlagExtEnter
      end
      object RefCelEcr: TEdit
        Left = 166
        Top = 210
        Width = 128
        Height = 21
        Enabled = False
        TabOrder = 5
        OnEnter = RefCelEnter
      end
      object Revision: TRadioGroup
        Left = 43
        Top = 127
        Width = 253
        Height = 50
        ItemIndex = 1
        Items.Strings = (
          'Inclure les '#233'critures de r'#233'vision'
          'Exclure les '#233'critures de r'#233'vision')
        TabOrder = 2
      end
      object AvecIFRS: TCheckBox
        Left = 42
        Top = 188
        Width = 255
        Height = 17
        Caption = 'Inclure les '#233'critures d'#39'IFRS'
        TabOrder = 3
      end
    end
    object Etablissement: TTabSheet
      Caption = 'Etablissement'
      object Label1: TLabel
        Left = 44
        Top = 60
        Width = 65
        Height = 13
        Caption = '&Etablissement'
      end
      object cEtablissement: THValComboBox
        Left = 44
        Top = 78
        Width = 253
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 0
        OnEnter = cEtablissementEnter
        TagDispatch = 0
        Vide = True
        DataType = 'TTETABLISSEMENT'
      end
      object FlagExtEtab: TCheckBox
        Left = 44
        Top = 213
        Width = 120
        Height = 17
        Caption = '&R'#233'f'#233'rence de cellule'
        TabOrder = 1
        OnClick = FlagExtClick
      end
      object RefCelEtab: TEdit
        Left = 168
        Top = 210
        Width = 126
        Height = 21
        Enabled = False
        TabOrder = 2
        OnEnter = RefCelEnter
      end
    end
    object Devise: TTabSheet
      Caption = 'Devise'
      object Label4: TLabel
        Left = 44
        Top = 60
        Width = 33
        Height = 13
        Caption = '&Devise'
      end
      object cDevise: THValComboBox
        Left = 44
        Top = 78
        Width = 253
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 0
        OnEnter = cDeviseEnter
        TagDispatch = 0
        Vide = True
        DataType = 'TTDEVISE'
      end
      object FlagExtDev: TCheckBox
        Left = 44
        Top = 213
        Width = 120
        Height = 17
        Caption = '&R'#233'f'#233'rence de cellule'
        TabOrder = 1
        OnClick = FlagExtClick
        OnEnter = FlagExtEnter
      end
      object RefCelDev: TEdit
        Left = 168
        Top = 210
        Width = 126
        Height = 21
        Enabled = False
        TabOrder = 2
        OnEnter = RefCelEnter
      end
    end
    object ChoixChamp: TTabSheet
      Caption = 'ChoixChamp'
      object HLabel1: THLabel
        Left = 4
        Top = 8
        Width = 179
        Height = 13
        Alignment = taCenter
        Caption = 'Quel champ souha'#238'tez-vous extraire ?'
      end
      object GridFields: THDBGrid
        Left = 3
        Top = 28
        Width = 332
        Height = 165
        Ctl3D = True
        DataSource = DSFields
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Options = [dgTitles, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgAlwaysShowSelection]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'MS Sans Serif'
        TitleFont.Style = []
        OnDblClick = GridFieldsDblClick
        Row = 1
        MultiSelection = False
        SortEnabled = False
        MyDefaultRowHeight = 0
        Columns = <
          item
            Expanded = False
            FieldName = 'DH_NOMCHAMP'
            Title.Caption = 'Champ'
            Width = 145
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DH_LIBELLE'
            Title.Caption = 'Libell'#233
            Width = 167
            Visible = True
          end>
      end
      object FlagExtField: TCheckBox
        Left = 4
        Top = 222
        Width = 120
        Height = 17
        Caption = '&R'#233'f'#233'rence de cellule'
        TabOrder = 1
        OnClick = FlagExtClick
        OnEnter = FlagExtEnter
      end
      object RefCelField: TEdit
        Left = 128
        Top = 219
        Width = 128
        Height = 21
        Enabled = False
        TabOrder = 2
        OnEnter = RefCelEnter
      end
      object cLibelle: TCheckBox
        Left = 4
        Top = 198
        Width = 215
        Height = 17
        Caption = '&Afficher le libell'#233' complet au lieu du code'
        TabOrder = 3
        OnEnter = FlagExtEnter
      end
    end
    object PAvances: TTabSheet
      Caption = 'PAvances'
      ImageIndex = 13
      object bEffaceAvance: TToolbarButton97
        Left = 300
        Top = 204
        Width = 23
        Height = 22
        DisplayMode = dmGlyphOnly
        OnClick = bEffaceAvanceClick
        GlobalIndexImage = 'Z0080_S16G1'
      end
      object Z_C3: THValComboBox
        Left = 2
        Top = 109
        Width = 102
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
      end
      object Z_C2: THValComboBox
        Left = 2
        Top = 77
        Width = 102
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 1
        TagDispatch = 0
      end
      object Z_C1: THValComboBox
        Left = 2
        Top = 45
        Width = 102
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
      end
      object ZO1: THValComboBox
        Left = 106
        Top = 45
        Width = 78
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        DataType = 'TTCOMPARE'
      end
      object ZO2: THValComboBox
        Left = 106
        Top = 77
        Width = 78
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        DataType = 'TTCOMPARE'
      end
      object ZO3: THValComboBox
        Left = 106
        Top = 109
        Width = 78
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 5
        TagDispatch = 0
        DataType = 'TTCOMPARE'
      end
      object ZV3: TEdit
        Left = 187
        Top = 109
        Width = 109
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 6
      end
      object ZV2: TEdit
        Left = 187
        Top = 77
        Width = 109
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 7
      end
      object ZV1: TEdit
        Left = 187
        Top = 45
        Width = 109
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 8
      end
      object ZG1: TComboBox
        Left = 300
        Top = 45
        Width = 38
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 9
        Items.Strings = (
          'Et'
          'Ou')
      end
      object ZG2: TComboBox
        Left = 300
        Top = 77
        Width = 38
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 10
        Items.Strings = (
          'Et'
          'Ou')
      end
      object Z_C4: THValComboBox
        Left = 2
        Top = 141
        Width = 102
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 11
        TagDispatch = 0
      end
      object Z_C5: THValComboBox
        Left = 2
        Top = 173
        Width = 102
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 12
        TagDispatch = 0
      end
      object Z_C6: THValComboBox
        Left = 2
        Top = 205
        Width = 102
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 13
        TagDispatch = 0
      end
      object ZO6: THValComboBox
        Left = 106
        Top = 205
        Width = 78
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 14
        TagDispatch = 0
        DataType = 'TTCOMPARE'
      end
      object ZO5: THValComboBox
        Left = 106
        Top = 173
        Width = 78
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 15
        TagDispatch = 0
        DataType = 'TTCOMPARE'
      end
      object ZO4: THValComboBox
        Left = 106
        Top = 141
        Width = 78
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 16
        TagDispatch = 0
        DataType = 'TTCOMPARE'
      end
      object ZV4: TEdit
        Left = 187
        Top = 141
        Width = 109
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 17
      end
      object ZV5: TEdit
        Left = 187
        Top = 173
        Width = 109
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 18
      end
      object ZV6: TEdit
        Left = 187
        Top = 205
        Width = 109
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 19
      end
      object ZG4: TComboBox
        Left = 300
        Top = 141
        Width = 38
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 20
        Items.Strings = (
          'Et'
          'Ou')
      end
      object ZG5: TComboBox
        Left = 300
        Top = 173
        Width = 38
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 21
        Items.Strings = (
          'Et'
          'Ou')
      end
      object ZG3: TComboBox
        Left = 300
        Top = 109
        Width = 38
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 22
        Items.Strings = (
          'Et'
          'Ou')
      end
      object CBLib: TCheckBox
        Left = 3
        Top = 10
        Width = 161
        Height = 17
        Caption = 'Voir les libell'#233's des champs'
        TabOrder = 23
        OnClick = CBLibClick
      end
      object rEcr: TRadioGroup
        Left = 169
        Top = 2
        Width = 165
        Height = 33
        Caption = 'Ecritures...'
        Columns = 2
        Enabled = False
        Items.Strings = (
          'G'#233'n'#233'rales'
          'Analytiques')
        TabOrder = 24
        OnClick = rEcrClick
      end
    end
    object Resultat: TTabSheet
      Caption = 'Resultat'
      object Panel1: TPanel
        Left = 0
        Top = 2
        Width = 337
        Height = 287
        BevelOuter = bvLowered
        Caption = ' '
        TabOrder = 0
        object HLabel8: THLabel
          Left = 109
          Top = 8
          Width = 123
          Height = 13
          Alignment = taCenter
          Caption = 'R'#233'sum'#233' de la formule'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object l1: THLabel
          Left = 9
          Top = 36
          Width = 76
          Height = 13
          Caption = 'Type de formule'
        end
        object l2: THLabel
          Left = 9
          Top = 54
          Width = 78
          Height = 13
          Caption = 'Source donn'#233'es'
        end
        object l3: THLabel
          Left = 9
          Top = 73
          Width = 75
          Height = 13
          Caption = 'Type d'#39#233'critures'
        end
        object l4: THLabel
          Left = 9
          Top = 117
          Width = 65
          Height = 13
          Caption = 'Etablissement'
        end
        object l5: THLabel
          Left = 9
          Top = 135
          Width = 33
          Height = 13
          Caption = 'Devise'
        end
        object l6: THLabel
          Left = 9
          Top = 157
          Width = 36
          Height = 13
          Caption = 'P'#233'riode'
        end
        object pFormule: THLabel
          Left = 12
          Top = 239
          Width = 317
          Height = 45
          AutoSize = False
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clMaroon
          Font.Height = -11
          Font.Name = 'Courier New'
          Font.Style = [fsBold]
          ParentColor = False
          ParentFont = False
          WordWrap = True
        end
        object r1: THLabel
          Left = 94
          Top = 37
          Width = 236
          Height = 13
          AutoSize = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object r2: THLabel
          Left = 94
          Top = 55
          Width = 236
          Height = 13
          AutoSize = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object r3: THLabel
          Left = 94
          Top = 73
          Width = 236
          Height = 39
          AutoSize = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object r4: THLabel
          Left = 94
          Top = 118
          Width = 236
          Height = 13
          AutoSize = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object r5: THLabel
          Left = 94
          Top = 136
          Width = 236
          Height = 13
          AutoSize = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object r6: THLabel
          Left = 94
          Top = 152
          Width = 236
          Height = 39
          AutoSize = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object Bevel3: TBevel
          Left = 103
          Top = 3
          Width = 134
          Height = 23
          Shape = bsFrame
          Style = bsRaised
        end
        object Bevel4: TBevel
          Left = 8
          Top = 45
          Width = 322
          Height = 8
          Shape = bsBottomLine
          Style = bsRaised
        end
        object Bevel5: TBevel
          Left = 9
          Top = 63
          Width = 322
          Height = 8
          Shape = bsBottomLine
          Style = bsRaised
        end
        object Bevel6: TBevel
          Left = 8
          Top = 108
          Width = 322
          Height = 8
          Shape = bsBottomLine
          Style = bsRaised
        end
        object Bevel7: TBevel
          Left = 9
          Top = 126
          Width = 322
          Height = 8
          Shape = bsBottomLine
          Style = bsRaised
        end
        object Bevel8: TBevel
          Left = 9
          Top = 143
          Width = 322
          Height = 8
          Shape = bsBottomLine
          Style = bsRaised
        end
        object Bevel9: TBevel
          Left = 9
          Top = 187
          Width = 322
          Height = 8
          Shape = bsBottomLine
          Style = bsRaised
        end
        object Bevel1: TBevel
          Left = 90
          Top = 41
          Width = 5
          Height = 193
          Shape = bsLeftLine
          Style = bsRaised
        end
        object HLabel6: THLabel
          Left = 9
          Top = 197
          Width = 22
          Height = 13
          Caption = 'Filtre'
        end
        object Bevel2: TBevel
          Left = 9
          Top = 227
          Width = 322
          Height = 8
          Shape = bsBottomLine
          Style = bsRaised
        end
        object r7: THLabel
          Left = 94
          Top = 197
          Width = 236
          Height = 34
          AutoSize = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
      end
    end
  end
  object TitleLine: TGroupBox [12]
    Left = 172
    Top = 19
    Width = 350
    Height = 7
    TabOrder = 7
  end
  object cSources: THValComboBox [14]
    Left = 8
    Top = 217
    Width = 62
    Height = 21
    Style = csDropDownList
    Color = clYellow
    ItemHeight = 13
    TabOrder = 5
    Visible = False
    Items.Strings = (
      'Rubrique comptable'
      'Compte g'#233'n'#233'ral'
      'Tiers'
      'Section analytique'
      'Journal'
      'Rubrique budgetaire - budget'#233
      'Rubrique budgetaire - r'#233'alis'#233
      'Compte budg'#233'taire - budget'#233
      'Compte budg'#233'taire - r'#233'alis'#233
      'Section budg'#233'taire - budget'#233
      'Section budg'#233'taire - r'#233'alis'#233
      'Budget'
      'Soci'#233't'#233)
    TagDispatch = 0
  end
  object TabSheet1: TTabSheet [15]
    Caption = 'Devise'
    object Label6: TLabel
      Left = 44
      Top = 60
      Width = 33
      Height = 13
      Caption = '&Devise'
    end
    object HValComboBox1: THValComboBox
      Left = 44
      Top = 78
      Width = 253
      Height = 21
      Style = csDropDownList
      ItemHeight = 0
      TabOrder = 0
      OnEnter = cDeviseEnter
      TagDispatch = 0
      Vide = True
      DataType = 'TTDEVISEETAT'
    end
    object CheckBox1: TCheckBox
      Left = 44
      Top = 213
      Width = 120
      Height = 17
      Caption = '&R'#233'f'#233'rence de cellule'
      TabOrder = 1
      OnClick = FlagExtClick
      OnEnter = FlagExtEnter
    end
    object Edit1: TEdit
      Left = 168
      Top = 210
      Width = 126
      Height = 21
      Enabled = False
      TabOrder = 2
      Text = 'Edit1'
      OnEnter = RefCelEnter
    end
  end
  inherited cControls: TListBox
    Left = 103
    Top = 144
  end
  inherited Msg: THMsgBox
    Mess.Strings = (
      'Etape'
      '1;Assistant;Voulez-vous quitter l'#39'assistant ?;Q;YN;Y;C;'
      
        '2;Moniteur SQL;Vous ne pouvez pas ex'#233'cuter ce type de commande.;' +
        'W;O;O;O;')
    Left = 69
    Top = 6
  end
  object MsgEF: THMsgBox [18]
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Assistant '#233'tats financiers'
      'S'#233'lectionnez le type d'#39#233'critures '#224' prendre en compte.'
      
        'S'#233'lectionnez la mani'#232're d'#39'int'#233'grer les A-nouveaux dans le calcul' +
        '.'
      'S'#233'lectionnez l'#39'exercice concern'#233'.'
      'S'#233'lectionnez le type de p'#233'riode.'
      'S'#233'lectionnez l'#39'indice de la p'#233'riode dans l'#39'exercice.'
      
        'S'#233'lectionnez si la p'#233'riode s'#39#233'tend du d'#233'but ou jusqu'#39#224' la fin de' +
        ' l'#39'exercice.'
      'S'#233'lectionnez l'#39#233'tablissement.'
      'S'#233'lectionnez la devise.'
      'S'#233'lectionnez un '#233'l'#233'ment dans la liste ci-dessus.'
      'Filtrez la liste par code.'
      'Filtrez la liste par libell'#233'.'
      'Aucun '#233'l'#233'ment s'#233'lectionn'#233
      'Janvier'
      'F'#233'vrier'
      'Mars'
      'Avril'
      'Mai'
      'Juin'
      'Juillet'
      'Ao'#251't'
      'Septembre'
      'Octobre'
      'Novembre'
      'D'#233'cembre'
      'de'
      #224
      
        'Renseignez la r'#233'f'#233'rence de la cellule contenant la valeur approp' +
        'ri'#233'e.'
      'Cochez cette case si la valeur est situ'#233'e dans une cellule.'
      'R'#233'f'#233'rence de cellule'
      '(soci'#233't'#233' en cours d'#39'utilisation)'
      'S'#233'lectionnez la source des donn'#233'es.'
      'S'#233'lectionnez l'#39'axe analytique.'
      'Rubriques'
      'Rubriques budg'#233'taires'
      'Comptes g'#233'n'#233'raux'
      'Tiers'
      'Sections analytiques'
      'Comptes budg'#233'taires'
      'Sections budg'#233'taires'
      'Journaux'
      'Budgets'
      'Soci'#233't'#233
      'Rubrique'
      'Rub. bud.'
      'Compte'
      'Tiers'
      'Section'
      'Cpte bud..'
      'Sect. bud.'
      'Journal'
      'Budget'
      'Soci'#233't'#233
      'Constante'
      
        'Pour recopier la formule dans la cellule courante, cliquez sur F' +
        'in.'
      '&Devise'
      '&Budget'
      'S'#233'lectionnez le budget.'
      'S'#233'lectionnez le(s) compte(s) variant(s)'
      '.'
      'Rubriques'
      'Rubriques budg'#233'taires'
      'Comptes g'#233'n'#233'raux'
      'Tiers'
      'Sections analytiques'
      'Comptes budg'#233'taires'
      'Sections budg'#233'taires'
      'Journaux'
      'Budgets'
      'Soci'#233't'#233
      'G'#233'n'#233'raux/Tiers'
      'G'#233'n'#233'raux/Sections'
      'Comptes/Sections Budget'
      ''
      ' ')
    Left = 37
    Top = 7
  end
  object Q: TQuery [19]
    DatabaseName = 'SOC'
    SQL.Strings = (
      
        'SELECT G_GENERAL AS CODE, G_LIBELLE AS LIBELLE FROM GENERAUX ORD' +
        'ER BY G_GENERAL')
    Left = 6
    Top = 252
  end
  object DS: TDataSource [20]
    AutoEdit = False
    DataSet = Q
    OnDataChange = DSDataChange
    Left = 37
    Top = 252
  end
  object QFields: TQuery [21]
    DatabaseName = 'SOC'
    SQL.Strings = (
      'SELECT DH_LIBELLE FROM DECHAMPS WHERE DH_PREFIXE="T"')
    Left = 68
    Top = 253
  end
  object DSFields: TDataSource [22]
    AutoEdit = False
    DataSet = QFields
    OnDataChange = DSFieldsDataChange
    Left = 99
    Top = 253
  end
  object MsgF: THMsgBox [23]
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Cumul'
      'Champ'
      'Constante (valeur)'
      'Constante (libell'#233')'
      'Rubrique :'
      'Rubrique budg'#233'taire :'
      'Compte :'
      'Tiers :'
      'Section :'
      'Compte budg'#233'taire :'
      'Section budg'#233'taire :'
      'Journal :'
      'Budget :'
      'Soci'#233't'#233
      'Axe'
      'R'#233'f'#233'rence de cellule :'
      '(+ r'#233'vision)'
      'Constante :'
      'Champ'
      'Type d'#39#233'critures'
      'Cell.')
    Left = 102
    Top = 6
  end
  inherited HMTrad: THSystemMenu
    Left = 133
    Top = 6
  end
  object TRub: THTable
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    IndexName = 'RB_CLE1'
    TableName = 'RUBRIQUE'
    dataBaseName = 'SOC'
    RequestLive = True
    Left = 5
    Top = 6
  end
  object Q2: TQuery
    DatabaseName = 'SOC'
    SQL.Strings = (
      
        'SELECT G_GENERAL AS CODE, G_LIBELLE AS LIBELLE FROM GENERAUX ORD' +
        'ER BY G_GENERAL')
    Left = 6
    Top = 287
  end
  object DS2: TDataSource
    AutoEdit = False
    DataSet = Q2
    OnDataChange = DS2DataChange
    Left = 37
    Top = 287
  end
end

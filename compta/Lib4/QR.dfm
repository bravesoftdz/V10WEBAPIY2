object FQR: TFQR
  Left = 46
  Top = 0
  Width = 694
  Height = 360
  HorzScrollBar.Range = 1060
  VertScrollBar.Range = 1000
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = '0'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  Scaled = False
  OnActivate = FormActivate
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 409
    Height = 207
    ActivePage = Avances
    ParentShowHint = False
    ShowHint = False
    TabOrder = 0
    object Standards: TTabSheet
      Caption = 'Standards'
      object TFGenJoker: THLabel
        Left = 10
        Top = 26
        Width = 123
        Height = 13
        AutoSize = False
        Caption = 'Comptes &g'#233'n'#233'raux    '
        FocusControl = FCpte1
        Visible = False
      end
      object TFGen: THLabel
        Left = 10
        Top = 26
        Width = 113
        Height = 13
        AutoSize = False
        Caption = 'Comptes &g'#233'n'#233'raux de'
        FocusControl = FCpte1
      end
      object TFaG: TLabel
        Left = 264
        Top = 26
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = FCpte2
      end
      object Label7: TLabel
        Left = 260
        Top = 133
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = FDateCompta2
      end
      object Hlabel2: THLabel
        Left = 9
        Top = 54
        Width = 113
        Height = 13
        AutoSize = False
        Caption = '&Nature de comptes'
        FocusControl = FNatureCpt
      end
      object HLabel4: THLabel
        Left = 9
        Top = 107
        Width = 113
        Height = 13
        AutoSize = False
        Caption = '&Exercice'
        FocusControl = FExercice
      end
      object HLabel5: THLabel
        Left = 9
        Top = 159
        Width = 113
        Height = 13
        AutoSize = False
        Caption = '&Type d'#39#233'critures'
        FocusControl = FNatureEcr
      end
      object HLabel6: THLabel
        Left = 9
        Top = 133
        Width = 113
        Height = 13
        AutoSize = False
        Caption = '&Dates comptables du'
        FocusControl = FDateCompta1
      end
      object TSelectCpte: THLabel
        Left = 10
        Top = 80
        Width = 106
        Height = 13
        AutoSize = False
        Caption = '&Mode de S'#233'lection'
        Color = clBtnFace
        FocusControl = FSelectCpte
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object FJoker: TEdit
        Left = 144
        Top = 22
        Width = 110
        Height = 21
        AutoSize = False
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 7
        Visible = False
        OnChange = FCpte1Change
        OnKeyDown = FJokerKeyDown
      end
      object FCpte1: THCpteEdit
        Left = 144
        Top = 22
        Width = 110
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 0
        OnChange = FCpte1Change
        OnKeyDown = FCpte1KeyDown
        OnKeyPress = FCpte1KeyPress
        ZoomTable = tzGeneral
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object FNatureEcr: THValComboBox
        Left = 144
        Top = 155
        Width = 247
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 6
        TagDispatch = 0
        DataType = 'TTQUALPIECECRIT'
      end
      object FDateCompta2: TMaskEdit
        Left = 281
        Top = 129
        Width = 110
        Height = 21
        Ctl3D = True
        EditMask = '!00/00/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 5
        Text = '  /  /    '
        OnKeyPress = ParamLaDate
      end
      object FDateCompta1: TMaskEdit
        Left = 144
        Top = 129
        Width = 110
        Height = 21
        Ctl3D = True
        EditMask = '!00/00/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 4
        Text = '  /  /    '
        OnKeyPress = ParamLaDate
      end
      object FExercice: THValComboBox
        Left = 144
        Top = 103
        Width = 247
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 3
        OnChange = FExerciceChange
        TagDispatch = 0
        DataType = 'TTEXERCICE'
      end
      object FNatureCpt: THValComboBox
        Left = 144
        Top = 50
        Width = 247
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 2
        OnChange = FNatureCptChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATGENE'
      end
      object FCpte2: THCpteEdit
        Left = 281
        Top = 22
        Width = 110
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 1
        OnChange = FCpte1Change
        ZoomTable = tzGeneral
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object FSelectCpte: THValComboBox
        Left = 144
        Top = 76
        Width = 247
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 8
        TagDispatch = 0
        DataType = 'TTSELECTIONCOMPTE'
      end
    end
    object Avances: TTabSheet
      Caption = 'Avanc'#233's'
      object HLabel7: THLabel
        Left = 10
        Top = 14
        Width = 111
        Height = 13
        AutoSize = False
        Caption = '&Etablissements'
        FocusControl = FEtab
      end
      object HLabel8: THLabel
        Left = 10
        Top = 51
        Width = 111
        Height = 13
        AutoSize = False
        Caption = '&Devises'
        FocusControl = FDevises
      end
      object HLabel1: THLabel
        Left = 10
        Top = 88
        Width = 99
        Height = 13
        AutoSize = False
        Caption = '&Comptes d'#39'exception'
        FocusControl = FExcep
      end
      object FDevises: THValComboBox
        Left = 147
        Top = 47
        Width = 247
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 1
        OnChange = FDevisesChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTDEVISE'
      end
      object FEtab: THValComboBox
        Left = 147
        Top = 10
        Width = 247
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTETABLISSEMENT'
      end
      object FExcep: TEdit
        Left = 147
        Top = 84
        Width = 247
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 2
      end
      object AvecRevision: TCheckBox
        Left = 8
        Top = 160
        Width = 129
        Height = 17
        AllowGrayed = True
        Caption = 'Ecritures de r&'#233'vision'
        TabOrder = 3
        Visible = False
      end
      object OnCum: TCheckBox
        Left = 148
        Top = 160
        Width = 129
        Height = 17
        Caption = 'Sur cumuls'
        TabOrder = 4
        Visible = False
        OnClick = OnCumClick
      end
    end
    object Mise: TTabSheet
      Caption = 'Options d'#39#233'dition'
      object GroupBox1: TGroupBox
        Left = 10
        Top = 54
        Width = 381
        Height = 115
        Caption = ' Afficher aussi... '
        TabOrder = 0
        object Bevel2: TBevel
          Left = 190
          Top = 7
          Width = 1
          Height = 110
        end
        object FMonetaire: TCheckBox
          Left = 197
          Top = 11
          Width = 172
          Height = 17
          Alignment = taLeftJustify
          Caption = '&Les symboles mon'#233'taires'
          TabOrder = 0
        end
        object FReport: TCheckBox
          Left = 197
          Top = 28
          Width = 172
          Height = 17
          Alignment = taLeftJustify
          Caption = '&Reports'
          Checked = True
          State = cbChecked
          TabOrder = 1
        end
        object Avance: TCheckBox
          Left = 197
          Top = 45
          Width = 172
          Height = 17
          Alignment = taLeftJustify
          Caption = '&Param'#233'trage avanc'#233
          TabOrder = 2
          Visible = False
          OnClick = AvanceClick
        end
      end
      object FMontant: TRadioGroup
        Left = 10
        Top = 10
        Width = 381
        Height = 37
        Caption = ' Affichage en... '
        Columns = 3
        ItemIndex = 0
        Items.Strings = (
          '&Euro'
          '&Devise')
        TabOrder = 1
        OnClick = FMontantClick
      end
    end
    object Option: TTabSheet
      Caption = 'Mise en page'
      object Apercu: TCheckBox
        Left = 10
        Top = 162
        Width = 146
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Aper'#231'u avant impression'
        Checked = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        State = cbChecked
        TabOrder = 1
        OnClick = ApercuClick
      end
      object FListe: TCheckBox
        Left = 185
        Top = 162
        Width = 107
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Liste d'#39'exportation'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
      end
      object GroupBox3: TGroupBox
        Left = 10
        Top = 10
        Width = 381
        Height = 105
        Caption = ' Pr'#233'senter avec... '
        TabOrder = 0
        object Bevel1: TBevel
          Left = 190
          Top = 7
          Width = 1
          Height = 100
        end
        object FLigneCpt: TCheckBox
          Left = 9
          Top = 29
          Width = 172
          Height = 17
          Alignment = taLeftJustify
          Caption = 'S'#233'parateurs de c&omptes'
          Checked = True
          State = cbChecked
          TabOrder = 0
        end
        object Reduire: TCheckBox
          Left = 197
          Top = 12
          Width = 172
          Height = 17
          Alignment = taLeftJustify
          Caption = '&2 pages sur 1'
          TabOrder = 1
        end
        object FTrait: TCheckBox
          Left = 9
          Top = 12
          Width = 172
          Height = 17
          Alignment = taLeftJustify
          Caption = 'S'#233'parateurs de &montants'
          Checked = True
          State = cbChecked
          TabOrder = 2
        end
        object FRappelCrit: TCheckBox
          Left = 197
          Top = 28
          Width = 172
          Height = 17
          Alignment = taLeftJustify
          AllowGrayed = True
          Caption = 'Rappel des crit&'#232'res de s'#233'lection'
          TabOrder = 3
        end
      end
      object FCouleur: TCheckBox
        Left = 334
        Top = 162
        Width = 57
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Couleur'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
      end
      object FGroupChoixRupt: TGroupBox
        Left = 10
        Top = 119
        Width = 381
        Height = 39
        Caption = ' Ruptures... '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        object FSansRupt: TRadioButton
          Left = 7
          Top = 15
          Width = 68
          Height = 17
          Caption = 'Sa&ns'
          Checked = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          TabStop = True
          OnClick = FSansRuptClick
        end
        object FAvecRupt: TRadioButton
          Left = 150
          Top = 15
          Width = 72
          Height = 17
          Caption = 'A&vec'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnClick = FSansRuptClick
        end
        object FSurRupt: TRadioButton
          Left = 292
          Top = 15
          Width = 62
          Height = 17
          Caption = 'S&ur'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          OnClick = FSansRuptClick
        end
      end
    end
    object tabSup: TTabSheet
      Caption = 'tabSup'
    end
    object TabRuptures: TTabSheet
      Caption = 'Ruptures'
      object FGroupRuptures: TGroupBox
        Left = 4
        Top = 80
        Width = 392
        Height = 96
        Caption = ' Ruptures... '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        object TFPlanRuptures: TLabel
          Left = 8
          Top = 25
          Width = 24
          Height = 13
          Caption = '&Plan '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TFCodeRupt1: TLabel
          Left = 8
          Top = 64
          Width = 19
          Height = 13
          AutoSize = False
          Caption = '&De'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TFCodeRupt2: TLabel
          Left = 208
          Top = 64
          Width = 19
          Height = 13
          Alignment = taCenter
          AutoSize = False
          Caption = #224
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object FCodeRupt1: TComboBox
          Left = 50
          Top = 60
          Width = 154
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 1
        end
        object FCodeRupt2: TComboBox
          Left = 230
          Top = 60
          Width = 154
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 2
        end
        object FPlanRuptures: THValComboBox
          Left = 50
          Top = 21
          Width = 333
          Height = 21
          Style = csDropDownList
          Ctl3D = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          OnChange = FPlanRupturesChange
          TagDispatch = 0
          DataType = 'TTRUPTSECT1'
        end
      end
      object FGroupLibres: TGroupBox
        Left = 4
        Top = 80
        Width = 392
        Height = 96
        TabOrder = 3
        object TFLibre1: THLabel
          Left = 8
          Top = 25
          Width = 31
          Height = 13
          AutoSize = False
          Caption = '&De'
          FocusControl = FLibre1
        end
        object TFLibre2: THLabel
          Left = 214
          Top = 25
          Width = 6
          Height = 13
          Caption = #224
          FocusControl = FLibre2
        end
        object TFLibTriPar: THLabel
          Left = 8
          Top = 64
          Width = 36
          Height = 13
          Caption = '&Tri'#233' par'
          FocusControl = FLibTriPar
        end
        object FLibre1: TEdit
          Left = 50
          Top = 21
          Width = 154
          Height = 21
          AutoSize = False
          CharCase = ecUpperCase
          MaxLength = 17
          ReadOnly = True
          TabOrder = 0
          OnDblClick = FLibre1DblClick
        end
        object FLibre2: TEdit
          Left = 230
          Top = 21
          Width = 154
          Height = 21
          AutoSize = False
          CharCase = ecUpperCase
          MaxLength = 17
          ReadOnly = True
          TabOrder = 1
          OnDblClick = FLibre1DblClick
        end
        object FLibTriPar: TEdit
          Left = 50
          Top = 60
          Width = 333
          Height = 21
          AutoSize = False
          CharCase = ecUpperCase
          MaxLength = 17
          ReadOnly = True
          TabOrder = 2
          OnDblClick = FLibre1DblClick
        end
      end
      object FGroupQuelleRupture: TGroupBox
        Left = 4
        Top = 6
        Width = 392
        Height = 45
        Caption = ' G'#233'rer les ruptures '#224' partir des ... '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object FRuptures: TRadioButton
          Left = 9
          Top = 19
          Width = 120
          Height = 17
          Caption = 'Plans de &ruptures'
          Checked = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          TabStop = True
          OnClick = FRupturesClick
        end
        object FTablesLibres: TRadioButton
          Left = 132
          Top = 19
          Width = 106
          Height = 17
          Caption = '&Tables libres'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnClick = FRupturesClick
        end
        object FPlansCo: TRadioButton
          Left = 241
          Top = 19
          Width = 147
          Height = 17
          Caption = 'Plans de &correspondances'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          OnClick = FRupturesClick
        end
      end
      object FOnlyCptAssocie: TCheckBox
        Left = 14
        Top = 59
        Width = 172
        Height = 17
        Caption = '&Uniquement Cpts. associ'#233's'
        TabOrder = 1
      end
      object FSautPageRupt: TCheckBox
        Left = 287
        Top = 59
        Width = 96
        Height = 17
        Caption = '&Saut de page'
        TabOrder = 2
        Visible = False
      end
    end
    object TabComplement: TTabSheet
      Caption = 'Compl'#233'ments'
      object TFNat0: THLabel
        Left = 10
        Top = 30
        Width = 113
        Height = 13
        AutoSize = False
        Caption = 'Comptes 1'
        FocusControl = FNat01
      end
      object TFNat1: THLabel
        Left = 10
        Top = 63
        Width = 113
        Height = 13
        AutoSize = False
        Caption = 'Comptes 2'
        FocusControl = FNat11
      end
      object TFNat2: THLabel
        Left = 10
        Top = 95
        Width = 113
        Height = 13
        AutoSize = False
        Caption = 'Comptes 3'
        FocusControl = FNat21
      end
      object TFNat3: THLabel
        Left = 10
        Top = 128
        Width = 113
        Height = 13
        AutoSize = False
        Caption = 'Comptes 4'
        FocusControl = FNat31
      end
      object FNat01: THCpteEdit
        Left = 144
        Top = 26
        Width = 110
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 0
        OnChange = FCpte1Change
        ZoomTable = tzNatEcrE0
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object FNat02: THCpteEdit
        Left = 281
        Top = 26
        Width = 110
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 1
        OnChange = FCpte1Change
        ZoomTable = tzNatEcrE0
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object FNat11: THCpteEdit
        Left = 144
        Top = 59
        Width = 110
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 2
        OnChange = FCpte1Change
        ZoomTable = tzNatEcrE1
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object FNat12: THCpteEdit
        Left = 281
        Top = 59
        Width = 110
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 3
        OnChange = FCpte1Change
        ZoomTable = tzNatEcrE1
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object FNat21: THCpteEdit
        Left = 144
        Top = 91
        Width = 110
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 4
        OnChange = FCpte1Change
        ZoomTable = tzNatEcrE2
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object FNat22: THCpteEdit
        Left = 281
        Top = 91
        Width = 110
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 5
        OnChange = FCpte1Change
        ZoomTable = tzNatEcrE2
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object FNat31: THCpteEdit
        Left = 144
        Top = 124
        Width = 110
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 6
        OnChange = FCpte1Change
        ZoomTable = tzNatEcrE3
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object FNat32: THCpteEdit
        Left = 281
        Top = 124
        Width = 110
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 7
        OnChange = FCpte1Change
        ZoomTable = tzNatEcrE3
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
    end
    object PAvances: TTabSheet
      Caption = 'Avanc'#233's'
      ImageIndex = 7
      object bEffaceAvance: TToolbarButton97
        Left = 344
        Top = 141
        Width = 23
        Height = 22
        DisplayMode = dmGlyphOnly
        OnClick = bEffaceAvanceClick
        GlobalIndexImage = 'Z0080_S16G1'
      end
      object HTri: TLabel
        Left = 379
        Top = 4
        Width = 22
        Height = 13
        Caption = 'TRI'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Z_C3: THValComboBox
        Left = 2
        Top = 70
        Width = 102
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 11
        OnChange = Z_C1Change
        TagDispatch = 0
      end
      object Z_C2: THValComboBox
        Left = 2
        Top = 46
        Width = 102
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 6
        OnChange = Z_C1Change
        TagDispatch = 0
      end
      object Z_C1: THValComboBox
        Left = 2
        Top = 22
        Width = 102
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 1
        OnChange = Z_C1Change
        TagDispatch = 0
      end
      object ZO1: THValComboBox
        Left = 106
        Top = 22
        Width = 103
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        DataType = 'TTCOMPARE'
      end
      object ZO2: THValComboBox
        Left = 106
        Top = 46
        Width = 103
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 7
        TagDispatch = 0
        DataType = 'TTCOMPARE'
      end
      object ZO3: THValComboBox
        Left = 106
        Top = 70
        Width = 103
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 12
        TagDispatch = 0
        DataType = 'TTCOMPARE'
      end
      object ZV3: TEdit
        Left = 211
        Top = 70
        Width = 131
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 13
      end
      object ZV2: TEdit
        Left = 211
        Top = 46
        Width = 131
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 8
      end
      object ZV1: TEdit
        Left = 211
        Top = 22
        Width = 131
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 3
      end
      object ZG1: TComboBox
        Left = 345
        Top = 22
        Width = 38
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 4
        Items.Strings = (
          'Et'
          'Ou')
      end
      object ZG2: TComboBox
        Left = 345
        Top = 46
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
      object Z_C4: THValComboBox
        Left = 2
        Top = 94
        Width = 102
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 16
        OnChange = Z_C1Change
        TagDispatch = 0
      end
      object Z_C5: THValComboBox
        Left = 2
        Top = 118
        Width = 102
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 21
        OnChange = Z_C1Change
        TagDispatch = 0
      end
      object Z_C6: THValComboBox
        Left = 2
        Top = 142
        Width = 102
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 26
        OnChange = Z_C1Change
        TagDispatch = 0
      end
      object ZO6: THValComboBox
        Left = 106
        Top = 142
        Width = 103
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 27
        TagDispatch = 0
        DataType = 'TTCOMPARE'
      end
      object ZO5: THValComboBox
        Left = 106
        Top = 118
        Width = 103
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 22
        TagDispatch = 0
        DataType = 'TTCOMPARE'
      end
      object ZO4: THValComboBox
        Left = 106
        Top = 94
        Width = 103
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 17
        TagDispatch = 0
        DataType = 'TTCOMPARE'
      end
      object ZV4: TEdit
        Left = 211
        Top = 94
        Width = 131
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 18
      end
      object ZV5: TEdit
        Left = 211
        Top = 118
        Width = 131
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 23
      end
      object ZV6: TEdit
        Left = 211
        Top = 142
        Width = 131
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 28
      end
      object ZG4: TComboBox
        Left = 345
        Top = 94
        Width = 38
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 19
        Items.Strings = (
          'Et'
          'Ou')
      end
      object ZG5: TComboBox
        Left = 345
        Top = 118
        Width = 38
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 24
        Items.Strings = (
          'Et'
          'Ou')
      end
      object ZG3: TComboBox
        Left = 345
        Top = 70
        Width = 38
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 14
        Items.Strings = (
          'Et'
          'Ou')
      end
      object CBLib: TCheckBox
        Left = 3
        Top = 2
        Width = 161
        Height = 17
        Caption = 'Voir les libell'#233's des champs'
        TabOrder = 0
        OnClick = CBLibClick
      end
      object CTri1: TCheckBox
        Left = 385
        Top = 22
        Width = 22
        Height = 17
        TabOrder = 5
        OnClick = CTri1Click
      end
      object CTri2: TCheckBox
        Left = 385
        Top = 46
        Width = 22
        Height = 17
        TabOrder = 10
        OnClick = CTri1Click
      end
      object CTri3: TCheckBox
        Left = 385
        Top = 70
        Width = 22
        Height = 17
        TabOrder = 15
        OnClick = CTri1Click
      end
      object CTri4: TCheckBox
        Left = 385
        Top = 94
        Width = 22
        Height = 17
        TabOrder = 20
        OnClick = CTri1Click
      end
      object CTri5: TCheckBox
        Left = 385
        Top = 118
        Width = 22
        Height = 17
        TabOrder = 25
        OnClick = CTri1Click
      end
      object CTri6: TCheckBox
        Left = 385
        Top = 142
        Width = 22
        Height = 17
        TabOrder = 29
        OnClick = CTri1Click
      end
      object FSautPageTRI: TCheckBox
        Left = 302
        Top = 164
        Width = 96
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Saut de page'
        Enabled = False
        TabOrder = 30
      end
    end
  end
  object Panel1: TPanel
    Left = -12
    Top = 356
    Width = 1060
    Height = 645
    Caption = 'Panel1'
    TabOrder = 1
    Visible = False
    object BtitreCrit: TQRBand
      Left = 1
      Top = 1
      Width = 1058
      Height = 86
      Align = alTop
      Color = clWhite
      Enabled = False
      AfterPrint = BtitreCritAfterPrint
      BandType = rbTitle
      BeforePrint = BtitreCritBeforePrint
      ForceNewPage = False
      RAZPage = False
      Frame.Color = clMaroon
      Frame.DrawTop = True
      Frame.DrawBottom = True
      Frame.DrawLeft = True
      Frame.DrawRight = True
      Ruler = qrrCmV
      PrintOnBottom = False
      object QRSysData4: TQRSysData
        Left = 1
        Top = 1
        width = 1057
        height = 37
        Alignment = taCenter
        Caption = '(Titre)'
        Color = clMaroon
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -27
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = True
        Data = qrsReportTitle
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RappelCrit: TQRLabel
        Left = 25
        Top = 46
        width = 179
        height = 14
        Caption = 'Rappel des crit'#232'res de s'#233'lection'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object CadreCrit: TQRShape
        Left = 0
        Top = 65
        width = 720
        Height = 20
        AutoSize = False
        SynLigne = 0
        NoRepeatValue = False
        Ombre = False
        OmbreDecalX = 0
        OmbreDecalY = 0
        OmbreColor = clBlack
        Shape = qrsRectangle
      end
    end
    object EnteteAutrePage: TQRBand
      Left = 1
      Top = 256
      Width = 1058
      Height = 36
      Align = alTop
      Color = clMaroon
      Enabled = False
      AfterPrint = EnteteAutrePageAfterPrint
      BandType = rbPageHeader
      BeforePrint = EnteteAutrePageBeforePrint
      ForceNewPage = False
      RAZPage = False
      Frame.DrawTop = False
      Frame.DrawBottom = True
      Frame.DrawLeft = False
      Frame.DrawRight = False
      Ruler = qrrCmV
      PrintOnBottom = False
      object TitreEntete: TQRSysData
        Left = 484
        Top = 3
        width = 91
        height = 17
        Alignment = taCenter
        Caption = '(Titre)'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        Visible = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = True
        Data = qrsReportTitle
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TitreBarre: TQRShape
        Left = 0
        Top = 21
        width = 1058
        Height = 4
        AutoSize = False
        Visible = False
        SynLigne = 0
        NoRepeatValue = False
        Ombre = False
        OmbreDecalX = 0
        OmbreDecalY = 0
        OmbreColor = clBlack
        Shape = qrsHorLine
      end
    end
    object BDetail: TQRBand
      Left = 1
      Top = 348
      Width = 1058
      Height = 18
      Align = alTop
      Color = clPurple
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -9
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      BandType = rbDetail
      OnCheckForSpace = BDetailCheckForSpace
      ForceNewPage = False
      RAZPage = False
      Frame.DrawTop = False
      Frame.DrawBottom = False
      Frame.DrawLeft = False
      Frame.DrawRight = False
      Ruler = qrrNone
      PrintOnBottom = False
    end
    object BFinEtat: TQRBand
      Left = 1
      Top = 366
      Width = 1058
      Height = 30
      Align = alTop
      Color = clOlive
      BandType = rbSummary
      BeforePrint = BFinEtatBeforePrint
      ForceNewPage = False
      RAZPage = False
      Frame.DrawTop = True
      Frame.DrawBottom = True
      Frame.DrawLeft = False
      Frame.DrawRight = False
      Ruler = qrrNone
      PrintOnBottom = False
    end
    object EntetePage: TQRBand
      Left = 1
      Top = 292
      Width = 1058
      Height = 34
      Align = alTop
      Color = clAqua
      BandType = rbPageHeader
      ForceNewPage = False
      RAZPage = False
      Frame.DrawTop = True
      Frame.DrawBottom = True
      Frame.DrawLeft = False
      Frame.DrawRight = False
      Ruler = qrrCmV
      PrintOnBottom = False
    end
    object TOPREPORT: TQRBand
      Left = 1
      Top = 326
      Width = 1058
      Height = 22
      Align = alTop
      Color = clSilver
      Enabled = False
      AfterPrint = TOPREPORTAfterPrint
      BandType = rbPageHeader
      ForceNewPage = False
      RAZPage = False
      Frame.DrawTop = False
      Frame.DrawBottom = True
      Frame.DrawLeft = False
      Frame.DrawRight = False
      Ruler = qrrNone
      PrintOnBottom = False
      object TITREREPORTH: TQRLabel
        Left = 195
        Top = 0
        width = 219
        height = 14
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'REPORT     '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
    end
    object BOTTOMREPORT: TQRBand
      Tag = 1
      Left = 1
      Top = 589
      Width = 1058
      Height = 22
      Align = alBottom
      Color = clSilver
      AfterPrint = BOTTOMREPORTAfterPrint
      BandType = rbPageFooter
      ForceNewPage = False
      RAZPage = False
      Frame.DrawTop = False
      Frame.DrawBottom = False
      Frame.DrawLeft = False
      Frame.DrawRight = False
      Ruler = qrrNone
      PrintOnBottom = False
      object TITREREPORTB: TQRLabel
        Left = 195
        Top = 0
        width = 219
        height = 14
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'REPORT     '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
    end
    object piedpage: TQRBand
      Left = 1
      Top = 611
      Width = 1058
      Height = 33
      Align = alBottom
      Color = clWhite
      AfterPrint = piedpageAfterPrint
      BandType = rbPageFooter
      ForceNewPage = False
      RAZPage = False
      Frame.DrawTop = True
      Frame.DrawBottom = False
      Frame.DrawLeft = False
      Frame.DrawRight = False
      Ruler = qrrNone
      PrintOnBottom = True
      object RNumversion: TQRLabel
        Left = 0
        Top = 18
        width = 56
        height = 14
        Alignment = taCenter
        Caption = 'N'#176' version :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RCopyright: TQRLabel
        Left = 0
        Top = 4
        width = 1056
        height = 14
        Alignment = taCenter
        AutoSize = False
        Caption = #169' Copyright APALATYS S.A. - Euro-Comptabilit'#233' HALLEY'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = True
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRSysData1: TQRSysData
        Left = 1
        Top = 2
        width = 115
        height = 14
        AutoSize = True
        Caption = 'Imprim'#233' le (Date, Heure)'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        Data = qrsDateTime
        Text = 'Imprim'#233' le '
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRSysData2: TQRSysData
        Left = 826
        Top = 2
        width = 231
        height = 17
        Alignment = taRightJustify
        Caption = 'Page : (Page N'#176')'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        Data = qrsPageNumber
        Text = 'Page : '
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RSociete: TQRLabel
        Left = 1
        Top = 18
        width = 288
        height = 14
        AutoSize = False
        Caption = 'Soci'#233't'#233' :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RUtilisateur: TQRLabel
        Left = 826
        Top = 18
        width = 231
        height = 14
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'Utilisateur :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
    end
    object BtitreZoom: TQRBand
      Left = 1
      Top = 87
      Width = 1058
      Height = 63
      Align = alTop
      Color = clWhite
      Enabled = False
      AfterPrint = BTitreAfterPrint
      BandType = rbTitle
      ForceNewPage = False
      RAZPage = False
      Frame.Color = clMaroon
      Frame.DrawTop = True
      Frame.DrawBottom = True
      Frame.DrawLeft = True
      Frame.DrawRight = True
      Ruler = qrrCmV
      PrintOnBottom = False
      object QRSysData3: TQRSysData
        Left = 1
        Top = 1
        width = 1057
        height = 37
        Alignment = taCenter
        Caption = '(Titre)'
        Color = clMaroon
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -27
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = True
        Data = qrsReportTitle
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRLabel8Zoom: TQRLabel
        Left = 10
        Top = 41
        width = 57
        height = 14
        Caption = 'Comptes :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RCptZoom: TQRLabel
        Left = 70
        Top = 41
        width = 50
        height = 14
        Caption = 'RCptZoom'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRLabel16Zoom: TQRLabel
        Left = 131
        Top = 41
        width = 119
        height = 14
        Caption = 'Dates comptables du '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRLabel21Zoom: TQRLabel
        Left = 459
        Top = 41
        width = 93
        height = 14
        Caption = 'Type d'#39#233'critures '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRLabel25Zoom: TQRLabel
        Left = 336
        Top = 41
        width = 13
        height = 14
        Caption = 'au'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRLabel26Zoom: TQRLabel
        Left = 859
        Top = 41
        width = 86
        height = 14
        Caption = 'Etablissements'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RDateCompta1Zoom: TQRLabel
        Left = 252
        Top = 41
        width = 71
        height = 14
        Caption = 'RDateCompta1'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RDateCompta2Zoom: TQRLabel
        Left = 358
        Top = 41
        width = 71
        height = 14
        Caption = 'RDateCompta2'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RNatureEcrZoom: TQRLabel
        Left = 556
        Top = 41
        width = 55
        height = 14
        Caption = 'RNatureEcr'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object REtabZoom: TQRLabel
        Left = 952
        Top = 41
        width = 28
        height = 14
        Caption = 'REtab'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRLabel38Zoom: TQRLabel
        Left = 670
        Top = 41
        width = 44
        height = 14
        Caption = 'Devises'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RDevisesZoom: TQRLabel
        Left = 722
        Top = 41
        width = 46
        height = 14
        Caption = 'RDevises'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
    end
    object BTitre: TQRBand
      Tag = 100
      Left = 1
      Top = 150
      Width = 1058
      Height = 106
      Align = alTop
      Color = clWhite
      AfterPrint = BTitreAfterPrint
      BandType = rbTitle
      ForceNewPage = False
      RAZPage = False
      Frame.Color = clMaroon
      Frame.DrawTop = True
      Frame.DrawBottom = True
      Frame.DrawLeft = True
      Frame.DrawRight = True
      Ruler = qrrCmV
      PrintOnBottom = False
      object RJoker: TQRLabel
        Left = 10
        Top = 41
        width = 106
        height = 14
        Caption = 'Comptes g'#233'n'#233'raux'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        Visible = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 101
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TTitre: TQRSysData
        Left = 1
        Top = 1
        width = 1057
        height = 37
        Alignment = taCenter
        Caption = '(Titre)'
        Color = clMaroon
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -27
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = True
        Data = qrsReportTitle
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RCpte: TQRLabel
        Left = 10
        Top = 41
        width = 123
        height = 14
        Caption = 'Comptes g'#233'n'#233'raux de'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 101
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TRa: TQRLabel
        Left = 226
        Top = 41
        width = 6
        height = 14
        Caption = #224
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 101
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RCpte1: TQRLabel
        Left = 144
        Top = 41
        width = 35
        height = 14
        Caption = 'RCpte1'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 101
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RCpte2: TQRLabel
        Left = 244
        Top = 41
        width = 35
        height = 14
        Caption = 'RCpte2'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 101
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RExcepGen: TQRLabel
        Left = 144
        Top = 89
        width = 57
        height = 14
        Caption = 'RExcepGen'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 101
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRLabel4: TQRLabel
        Left = 358
        Top = 56
        width = 119
        height = 14
        Caption = 'Dates comptables du '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 201
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRLabel5: TQRLabel
        Left = 358
        Top = 41
        width = 46
        height = 14
        Caption = 'Exercice'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 101
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRLabel7: TQRLabel
        Left = 358
        Top = 72
        width = 93
        height = 14
        Caption = 'Type d'#39#233'critures '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 201
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRLabel11: TQRLabel
        Left = 564
        Top = 56
        width = 13
        height = 14
        Caption = 'au'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 201
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRLabel15: TQRLabel
        Left = 732
        Top = 41
        width = 86
        height = 14
        Caption = 'Etablissements'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 201
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RDateCompta1: TQRLabel
        Left = 478
        Top = 56
        width = 71
        height = 14
        Caption = 'RDateCompta1'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 201
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RDateCompta2: TQRLabel
        Left = 586
        Top = 56
        width = 71
        height = 14
        Caption = 'RDateCompta2'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 201
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RExercice: TQRLabel
        Left = 478
        Top = 41
        width = 49
        height = 14
        Caption = 'RExercice'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 101
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RNatureEcr: TQRLabel
        Left = 478
        Top = 72
        width = 55
        height = 14
        Caption = 'RNatureEcr'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 201
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object REtab: TQRLabel
        Left = 849
        Top = 41
        width = 28
        height = 14
        Caption = 'REtab'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 201
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRLabel18: TQRLabel
        Left = 732
        Top = 56
        width = 44
        height = 14
        Caption = 'Devises'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 101
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RDevises: TQRLabel
        Left = 849
        Top = 56
        width = 46
        height = 14
        Caption = 'RDevises'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 101
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRLabel12: TQRLabel
        Left = 10
        Top = 56
        width = 105
        height = 14
        Caption = 'Nature de comptes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 101
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RNatureCpt: TQRLabel
        Left = 144
        Top = 56
        width = 55
        height = 14
        Caption = 'RNatureCpt'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 101
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRLabel24: TQRLabel
        Left = 10
        Top = 88
        width = 107
        height = 15
        AutoSize = False
        Caption = 'Sauf '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 101
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object RSelectCpte: TQRLabel
        Left = 144
        Top = 72
        width = 59
        height = 14
        Caption = 'RSelectCpte'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 101
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TRSelectCpte: TQRLabel
        Left = 10
        Top = 72
        width = 102
        height = 14
        Caption = 'Mode de S'#233'lection'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynLigne = 0
        NoRepeatValue = False
        AlignToBand = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 101
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
    end
  end
  object Formateur: THNumEdit
    Left = 505
    Top = 82
    Width = 85
    Height = 21
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    Visible = False
    Decimals = 2
    Digits = 12
    Masks.PositiveMask = '#,##0'
    Debit = False
    UseRounding = True
    Validate = False
  end
  object Dock971: TDock97
    Left = 0
    Top = 928
    Width = 1060
    Height = 72
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 36
      ClientHeight = 32
      ClientWidth = 1060
      Caption = 'Barre d'#39'outils'
      ClientAreaHeight = 32
      ClientAreaWidth = 1060
      DockPos = 0
      DockRow = 1
      FullSize = True
      TabOrder = 0
      object BParamListe: TToolbarButton97
        Left = 6
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Param'#233'trer liste'
        Caption = 'Param'#233'trer liste'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BParamListeClick
        GlobalIndexImage = 'Z0714_S16G1'
      end
      object BValider: TToolbarButton97
        Left = 312
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Lancer l'#39#233'tat'
        Caption = 'OK'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Glyph.Data = {
          BE060000424DBE06000000000000360400002800000024000000120000000100
          0800000000008802000000000000000000000001000000010000000000000000
          80000080000000808000800000008000800080800000C0C0C000C0DCC000F0C8
          A400000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          000000000000000000000000000000000000F0FBFF00A4A0A000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00030303030303
          0303030303030303030303030303030303030303030303030303030303030303
          03030303030303030303030303030303030303030303FF030303030303030303
          03030303030303040403030303030303030303030303030303F8F8FF03030303
          03030303030303030303040202040303030303030303030303030303F80303F8
          FF030303030303030303030303040202020204030303030303030303030303F8
          03030303F8FF0303030303030303030304020202020202040303030303030303
          0303F8030303030303F8FF030303030303030304020202FA0202020204030303
          0303030303F8FF0303F8FF030303F8FF03030303030303020202FA03FA020202
          040303030303030303F8FF03F803F8FF0303F8FF03030303030303FA02FA0303
          03FA0202020403030303030303F8FFF8030303F8FF0303F8FF03030303030303
          FA0303030303FA0202020403030303030303F80303030303F8FF0303F8FF0303
          0303030303030303030303FA0202020403030303030303030303030303F8FF03
          03F8FF03030303030303030303030303FA020202040303030303030303030303
          0303F8FF0303F8FF03030303030303030303030303FA02020204030303030303
          03030303030303F8FF0303F8FF03030303030303030303030303FA0202020403
          030303030303030303030303F8FF0303F8FF03030303030303030303030303FA
          0202040303030303030303030303030303F8FF03F8FF03030303030303030303
          03030303FA0202030303030303030303030303030303F8FFF803030303030303
          030303030303030303FA0303030303030303030303030303030303F803030303
          0303030303030303030303030303030303030303030303030303030303030303
          0303}
        GlyphMask.Data = {00000000}
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 344
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Caption = 'Annuler'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Flat = False
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
        GlobalIndexImage = 'Z1770_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 376
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Aide'
        Caption = 'Aide'
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
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
    end
    object PanelFiltre: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 1060
      Caption = 'Filtres'
      ClientAreaHeight = 32
      ClientAreaWidth = 1060
      DockPos = 0
      FullSize = True
      TabOrder = 1
      object BFiltre: TToolbarButton97
        Left = 6
        Top = 6
        Width = 61
        Height = 21
        Hint = 'Menu filtre'
        Caption = 'F&iltres'
        DropdownArrow = True
        DropdownMenu = POPF
        Layout = blGlyphRight
        ParentShowHint = False
        ShowHint = True
      end
      object FFiltres: THValComboBox
        Left = 72
        Top = 6
        Width = 323
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = FFiltresChange
        TagDispatch = 0
      end
    end
  end
  object Q: TQuery
    DatabaseName = 'SOC'
    SQL.Strings = (
      'Select * From GENERAUX Order By G_GENERAL')
    Left = 444
    Top = 108
  end
  object S: TDataSource
    DataSet = Q
    Left = 486
    Top = 154
  end
  object QRP: TQuickReport
    DataSource = S
    Columns = 1
    ColThenRow = False
    DisplayPrintDialog = True
    Orientation = poLandscape
    PageFrame.DrawTop = False
    PageFrame.DrawBottom = False
    PageFrame.DrawLeft = False
    PageFrame.DrawRight = False
    PaperLength = 0
    PaperSize = qrpA4
    PaperWidth = 0
    ReprintGroupHeaders = False
    RestartData = True
    SQLCompatible = False
    TitleBeforeHeader = True
    Matricielle = False
    Left = 445
    Top = 38
  end
  object MsgRien: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?caption?;Aucun enregistrement avec ces crit'#232'res.;E;O;O;O;'
      
        '1;Erreur;Dates de d'#233'but et de fin d'#39'exercice incoh'#233'rentes.;E;O;O' +
        ';O;'
      
        '2;Erreur;Dates de d'#233'but incoh'#233'rentes avec l'#39'exercice d'#233'sir'#233'.;E;O' +
        ';O;O;'
      '3;Erreur de saisie; : Caract'#232're non admis !;E;O;O;O;'
      'Soci'#233't'#233' :'
      'Utilisateur :'
      
        '6;Erreur;Le choix "Compte non sold'#233's" n'#39'est pas autoris'#233'.;E;O;O;' +
        'O;')
    Left = 567
    Top = 94
  end
  object HMCrit: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'G'#233'n'#233'raux :'
      'Tiers :'
      'Sections :'
      'Journaux :'
      'Dates :'
      'Devise :'
      'Etablissement :'
      'Nature : '
      'Type d'#39#233'criture :'
      'Exercice :'
      '10'
      '11'
      '12'
      '13'
      '14'
      '15'
      '16'
      '17'
      '18'
      '19'
      'de'
      #224
      'Tous'
      'Toutes')
    Left = 556
    Top = 48
  end
  object MsgQR: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Vous devez d'#39'abord s'#233'lectionner une devise particuli' +
        #232're.;E;O;O;O;'
      
        '1;?caption?;La fourchette de dates n'#39'appartient pas '#224' un exercic' +
        'e.;E;O;O;O;'
      
        '2;?caption?;La fourchette de p'#233'riodes n'#39'appartient pas '#224' un exer' +
        'cice.;E;O;O;O;'
      'Version'
      'du'
      
        '5;?caption?;Vous devez sp'#233'cifier un tri pour les natures libres.' +
        ';E;O;O;O;')
    Left = 561
    Top = 141
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 490
    Top = 105
  end
  object POPF: TPopupMenu
    OnPopup = POPFPopup
    Left = 484
    Top = 216
    object BCreerFiltre: TMenuItem
      Caption = '&Cr'#233'er un filtre'
      OnClick = BCreerFiltreClick
    end
    object BSaveFiltre: TMenuItem
      Caption = '&Enregistrer le filtre'
      OnClick = BSaveFiltreClick
    end
    object BDupFiltre: TMenuItem
      Caption = '&Dupliquer le filtre'
      OnClick = BDupFiltreClick
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
  object Timer1: TTimer
    Enabled = False
    Interval = 10
    OnTimer = Timer1Timer
    Left = 484
    Top = 38
  end
end

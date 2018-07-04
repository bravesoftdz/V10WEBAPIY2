object EtatPCL: TEtatPCL
  Left = 404
  Top = 178
  Width = 596
  Height = 396
  Caption = 'Etat de synth'#232'se'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 580
    Height = 281
    ActivePage = Complements
    Align = alClient
    ParentShowHint = False
    ShowHint = False
    TabOrder = 0
    object Standards: TTabSheet
      Caption = 'Standards'
      object GB1: TGroupBox
        Tag = 1
        Left = 54
        Top = 39
        Width = 129
        Height = 232
        TabOrder = 0
        object Formule1: TEdit
          Left = 5
          Top = 131
          Width = 119
          Height = 21
          Enabled = False
          TabOrder = 5
          OnEnter = Formule1Enter
          OnExit = Formule1Exit
        end
        object Exo1: THValComboBox
          Left = 5
          Top = 51
          Width = 119
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 0
          ParentCtl3D = False
          TabOrder = 2
          OnChange = Exo1Change
          TagDispatch = 0
          DataType = 'TTEXERCICE'
        end
        object FD11: TMaskEdit
          Left = 5
          Top = 78
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          MaxLength = 10
          ParentCtl3D = False
          TabOrder = 3
          Text = '  /  /    '
          OnExit = FD11Exit
          OnKeyPress = FD11KeyPress
        end
        object FD12: TMaskEdit
          Left = 58
          Top = 105
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          MaxLength = 10
          ParentCtl3D = False
          TabOrder = 4
          Text = '  /  /    '
          OnExit = FD11Exit
          OnKeyPress = FD11KeyPress
        end
        object OkCol1: TCheckBox
          Tag = 1
          Left = 7
          Top = 0
          Width = 69
          Height = 17
          Caption = 'Colonne 1'
          Checked = True
          State = cbChecked
          TabOrder = 0
          OnClick = OkCol1Click
        end
        object Type1: THValComboBox
          Left = 5
          Top = 24
          Width = 119
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 13
          ParentCtl3D = False
          TabOrder = 1
          OnChange = Type1Change
          Items.Strings = (
            'Rubrique'
            'Formule')
          TagDispatch = 0
          Values.Strings = (
            'RUB'
            'FOR')
        end
        object FValVar1: TSpinEdit
          Tag = 1
          Left = 76
          Top = 164
          Width = 41
          Height = 22
          Color = clYellow
          MaxValue = 0
          MinValue = 0
          TabOrder = 8
          Value = 0
          Visible = False
        end
        object Histobal1: THCritMaskEdit
          Left = 5
          Top = 185
          Width = 119
          Height = 21
          Enabled = False
          ParentShowHint = False
          ReadOnly = True
          ShowHint = True
          TabOrder = 7
          TagDispatch = 0
          DataType = 'CPBALSIT'
          ElipsisButton = True
          OnElipsisClick = OnBalSitClick
        end
        object AvecSitu1: TCheckBox
          Tag = 1
          Left = 6
          Top = 158
          Width = 97
          Height = 17
          Caption = 'sur situation'
          Enabled = False
          TabOrder = 6
          OnClick = OnClickAvecSituation
        end
      end
      object GB2: TGroupBox
        Tag = 2
        Left = 186
        Top = 39
        Width = 129
        Height = 232
        TabOrder = 1
        object HLabel2: THLabel
          Left = 231
          Top = 26
          Width = 43
          Height = 13
          Caption = '&Dates du'
          FocusControl = FD21
        end
        object Label1: TLabel
          Left = 350
          Top = 26
          Width = 12
          Height = 13
          Caption = 'au'
          FocusControl = FD22
        end
        object Formule2: TEdit
          Left = 4
          Top = 131
          Width = 119
          Height = 21
          Enabled = False
          TabOrder = 5
          OnEnter = Formule1Enter
          OnExit = Formule1Exit
        end
        object Exo2: THValComboBox
          Left = 6
          Top = 51
          Width = 119
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 0
          ParentCtl3D = False
          TabOrder = 2
          OnChange = Exo1Change
          TagDispatch = 0
          DataType = 'TTEXERCICE'
        end
        object FD21: TMaskEdit
          Left = 6
          Top = 78
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          MaxLength = 10
          ParentCtl3D = False
          TabOrder = 3
          Text = '  /  /    '
          OnExit = FD11Exit
          OnKeyPress = FD11KeyPress
        end
        object FD22: TMaskEdit
          Left = 57
          Top = 105
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          MaxLength = 10
          ParentCtl3D = False
          TabOrder = 4
          Text = '  /  /    '
          OnExit = FD11Exit
          OnKeyPress = FD11KeyPress
        end
        object OkCol2: TCheckBox
          Tag = 2
          Left = 7
          Top = 0
          Width = 69
          Height = 17
          Caption = 'Colonne 2'
          TabOrder = 0
          OnClick = OkCol1Click
        end
        object Type2: THValComboBox
          Left = 6
          Top = 24
          Width = 119
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 13
          ParentCtl3D = False
          TabOrder = 1
          OnChange = Type1Change
          Items.Strings = (
            'Rubrique'
            'Formule')
          TagDispatch = 0
          Values.Strings = (
            'RUB'
            'FOR')
        end
        object FValVar2: TSpinEdit
          Tag = 2
          Left = 75
          Top = 168
          Width = 41
          Height = 22
          Color = clYellow
          MaxValue = 0
          MinValue = 0
          TabOrder = 8
          Value = 0
          Visible = False
        end
        object Histobal2: THCritMaskEdit
          Left = 6
          Top = 185
          Width = 119
          Height = 21
          Enabled = False
          ParentShowHint = False
          ReadOnly = True
          ShowHint = True
          TabOrder = 7
          TagDispatch = 0
          DataType = 'CPBALSIT'
          ElipsisButton = True
          OnElipsisClick = OnBalSitClick
        end
        object AvecSitu2: TCheckBox
          Tag = 2
          Left = 6
          Top = 158
          Width = 97
          Height = 17
          Caption = 'sur situation'
          Enabled = False
          TabOrder = 6
          OnClick = OnClickAvecSituation
        end
      end
      object GB3: TGroupBox
        Tag = 3
        Left = 318
        Top = 39
        Width = 129
        Height = 232
        TabOrder = 2
        object HLabel9: THLabel
          Left = 231
          Top = 26
          Width = 43
          Height = 13
          Caption = '&Dates du'
          FocusControl = FD31
        end
        object Label2: TLabel
          Left = 350
          Top = 26
          Width = 12
          Height = 13
          Caption = 'au'
          FocusControl = FD32
        end
        object HLabel1: THLabel
          Left = 20
          Top = 200
          Width = 3
          Height = 13
          FocusControl = Histobal1
        end
        object Formule3: TEdit
          Left = 6
          Top = 131
          Width = 119
          Height = 21
          Enabled = False
          TabOrder = 5
          OnEnter = Formule1Enter
          OnExit = Formule1Exit
        end
        object Exo3: THValComboBox
          Left = 6
          Top = 51
          Width = 119
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 0
          ParentCtl3D = False
          TabOrder = 2
          OnChange = Exo1Change
          TagDispatch = 0
          DataType = 'TTEXERCICE'
        end
        object FD31: TMaskEdit
          Left = 6
          Top = 78
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          MaxLength = 10
          ParentCtl3D = False
          TabOrder = 3
          Text = '  /  /    '
          OnExit = FD11Exit
          OnKeyPress = FD11KeyPress
        end
        object FD32: TMaskEdit
          Left = 59
          Top = 105
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          MaxLength = 10
          ParentCtl3D = False
          TabOrder = 4
          Text = '  /  /    '
          OnExit = FD11Exit
          OnKeyPress = FD11KeyPress
        end
        object OkCol3: TCheckBox
          Tag = 3
          Left = 7
          Top = 0
          Width = 69
          Height = 17
          Caption = 'Colonne 3'
          TabOrder = 0
          OnClick = OkCol1Click
        end
        object Type3: THValComboBox
          Left = 6
          Top = 24
          Width = 119
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 13
          ParentCtl3D = False
          TabOrder = 1
          OnChange = Type1Change
          Items.Strings = (
            'Rubrique'
            'Formule')
          TagDispatch = 0
          Values.Strings = (
            'RUB'
            'FOR')
        end
        object FValVar3: TSpinEdit
          Tag = 3
          Left = 82
          Top = 172
          Width = 41
          Height = 22
          Color = clYellow
          MaxValue = 0
          MinValue = 0
          TabOrder = 8
          Value = 0
          Visible = False
        end
        object Histobal3: THCritMaskEdit
          Left = 6
          Top = 185
          Width = 119
          Height = 21
          Enabled = False
          ParentShowHint = False
          ReadOnly = True
          ShowHint = True
          TabOrder = 7
          TagDispatch = 0
          DataType = 'CPBALSIT'
          ElipsisButton = True
          OnElipsisClick = OnBalSitClick
        end
        object AvecSitu3: TCheckBox
          Tag = 3
          Left = 6
          Top = 158
          Width = 97
          Height = 17
          Caption = 'sur situation'
          Enabled = False
          TabOrder = 6
          OnClick = OnClickAvecSituation
        end
      end
      object GB4: TGroupBox
        Tag = 4
        Left = 450
        Top = 39
        Width = 129
        Height = 232
        TabOrder = 3
        object HLabel11: THLabel
          Left = 231
          Top = 26
          Width = 43
          Height = 13
          Caption = '&Dates du'
          FocusControl = FD41
        end
        object Label3: TLabel
          Left = 350
          Top = 26
          Width = 12
          Height = 13
          Caption = 'au'
          FocusControl = FD42
        end
        object Formule4: TEdit
          Left = 6
          Top = 131
          Width = 119
          Height = 21
          Enabled = False
          TabOrder = 5
          OnEnter = Formule1Enter
          OnExit = Formule1Exit
        end
        object Exo4: THValComboBox
          Left = 6
          Top = 51
          Width = 119
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 0
          ParentCtl3D = False
          TabOrder = 2
          OnChange = Exo1Change
          TagDispatch = 0
          DataType = 'TTEXERCICE'
        end
        object FD41: TMaskEdit
          Left = 6
          Top = 78
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          MaxLength = 10
          ParentCtl3D = False
          TabOrder = 3
          Text = '  /  /    '
          OnExit = FD11Exit
          OnKeyPress = FD11KeyPress
        end
        object FD42: TMaskEdit
          Left = 59
          Top = 105
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          MaxLength = 10
          ParentCtl3D = False
          TabOrder = 4
          Text = '  /  /    '
          OnExit = FD11Exit
          OnKeyPress = FD11KeyPress
        end
        object OkCol4: TCheckBox
          Tag = 4
          Left = 7
          Top = 0
          Width = 69
          Height = 17
          Caption = 'Colonne 4'
          TabOrder = 0
          OnClick = OkCol1Click
        end
        object Type4: THValComboBox
          Left = 6
          Top = 24
          Width = 119
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 13
          ParentCtl3D = False
          TabOrder = 1
          OnChange = Type1Change
          Items.Strings = (
            'Rubrique'
            'Formule')
          TagDispatch = 0
          Values.Strings = (
            'RUB'
            'FOR')
        end
        object FValVar4: TSpinEdit
          Tag = 4
          Left = 78
          Top = 168
          Width = 41
          Height = 22
          Color = clYellow
          MaxValue = 0
          MinValue = 0
          TabOrder = 8
          Value = 0
          Visible = False
        end
        object Histobal4: THCritMaskEdit
          Left = 6
          Top = 185
          Width = 119
          Height = 21
          Enabled = False
          ParentShowHint = False
          ReadOnly = True
          ShowHint = True
          TabOrder = 7
          TagDispatch = 0
          DataType = 'CPBALSIT'
          ElipsisButton = True
          OnElipsisClick = OnBalSitClick
        end
        object AvecSitu4: TCheckBox
          Tag = 4
          Left = 6
          Top = 158
          Width = 97
          Height = 17
          Caption = 'sur situation'
          Enabled = False
          TabOrder = 6
          OnClick = OnClickAvecSituation
        end
      end
      object GroupBox1: TGroupBox
        Left = 0
        Top = 39
        Width = 51
        Height = 232
        TabOrder = 4
        object HLabel12: THLabel
          Left = 4
          Top = 29
          Width = 24
          Height = 13
          Caption = 'Type'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel14: THLabel
          Left = 4
          Top = 56
          Width = 41
          Height = 13
          Caption = 'Exercice'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel15: THLabel
          Left = 4
          Top = 82
          Width = 38
          Height = 13
          Caption = 'Date du'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel16: THLabel
          Left = 4
          Top = 110
          Width = 12
          Height = 13
          Caption = 'au'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel17: THLabel
          Left = 4
          Top = 134
          Width = 37
          Height = 13
          Caption = 'Formule'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel18: THLabel
          Left = 4
          Top = 160
          Width = 41
          Height = 13
          Caption = 'Situation'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel4: THLabel
          Left = 4
          Top = 189
          Width = 39
          Height = 13
          Caption = 'Balance'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
      end
      object Panel1: TPanel
        Left = 0
        Top = 0
        Width = 572
        Height = 34
        Align = alTop
        BevelOuter = bvLowered
        TabOrder = 5
        object LFile: THLabel
          Left = 24
          Top = 11
          Width = 45
          Height = 13
          Caption = 'Maquette'
        end
        object ZMAQUETTENAME: THLabel
          Left = 140
          Top = 11
          Width = 98
          Height = 13
          Caption = 'ZMAQUETTENAME'
        end
        object ZMAQUETTE: THCritMaskEdit
          Left = 84
          Top = 7
          Width = 51
          Height = 21
          TabOrder = 0
          OnChange = ZMAQUETTEChange
          TagDispatch = 0
          DataType = 'CPSTDMAQ'
          ElipsisButton = True
          Libelle = ZMAQUETTENAME
        end
        object FAvecPourcent: TCheckBox
          Left = 328
          Top = 8
          Width = 110
          Height = 17
          Caption = 'Avec pourcentage'
          TabOrder = 1
        end
        object FAvecDetail: THValComboBox
          Left = 460
          Top = 6
          Width = 105
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 13
          ParentCtl3D = False
          TabOrder = 2
          Items.Strings = (
            'Etat seul'
            'D'#233'tail seulement'
            'Etat plus d'#233'tail')
          TagDispatch = 0
          Values.Strings = (
            'NON'
            'SEUL'
            'PLUS')
        end
      end
    end
    object PBilan: TTabSheet
      Caption = 'Standards'
      ImageIndex = 3
      object GBBILAN: TGroupBox
        Left = 12
        Top = 39
        Width = 377
        Height = 113
        Caption = 'Bilan pour'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object HLabel21: THLabel
          Left = 36
          Top = 24
          Width = 41
          Height = 13
          Caption = 'Exercice'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel22: THLabel
          Left = 36
          Top = 53
          Width = 28
          Height = 13
          Caption = 'Dates'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel23: THLabel
          Left = 240
          Top = 53
          Width = 12
          Height = 13
          Caption = 'au'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BILEXO: THValComboBox
          Left = 145
          Top = 19
          Width = 204
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          OnChange = BILEXOChange
          TagDispatch = 0
          DataType = 'TTEXERCICE'
        end
        object BILDATE: TMaskEdit
          Left = 145
          Top = 49
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 10
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 1
          Text = '  /  /    '
          OnKeyPress = FD11KeyPress
        end
        object BILDATE_: TMaskEdit
          Left = 283
          Top = 49
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 10
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 2
          Text = '  /  /    '
          OnKeyPress = FD11KeyPress
        end
        object BILBALSIT: THCritMaskEdit
          Left = 145
          Top = 79
          Width = 204
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          ParentShowHint = False
          ShowHint = True
          TabOrder = 3
          TagDispatch = 0
          DataType = 'CPBALSIT'
          ElipsisButton = True
          OnElipsisClick = OnBalSitClick
        end
        object BILAVECSIT: TCheckBox
          Left = 36
          Top = 81
          Width = 97
          Height = 17
          Caption = 'sur situation'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 4
          OnClick = BILAVECSITClick
        end
      end
      object GBBILANCOMP: TGroupBox
        Left = 12
        Top = 159
        Width = 377
        Height = 113
        TabOrder = 1
        object HLabel25: THLabel
          Left = 36
          Top = 53
          Width = 28
          Height = 13
          Caption = 'Dates'
        end
        object HLabel26: THLabel
          Left = 240
          Top = 53
          Width = 12
          Height = 13
          Caption = 'au'
        end
        object HLabel27: THLabel
          Left = 36
          Top = 23
          Width = 41
          Height = 13
          Caption = 'Exercice'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BILEXOCOMP: THValComboBox
          Left = 145
          Top = 19
          Width = 204
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 0
          ParentCtl3D = False
          TabOrder = 0
          OnChange = BILEXOChange
          TagDispatch = 0
          DataType = 'TTEXERCICE'
        end
        object BILDATECOMP: TMaskEdit
          Left = 145
          Top = 49
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          MaxLength = 10
          ParentCtl3D = False
          TabOrder = 1
          Text = '  /  /    '
          OnKeyPress = FD11KeyPress
        end
        object BILDATECOMP_: TMaskEdit
          Left = 284
          Top = 49
          Width = 65
          Height = 21
          Ctl3D = True
          EditMask = '!00/00/0000;1;_'
          MaxLength = 10
          ParentCtl3D = False
          TabOrder = 2
          Text = '  /  /    '
          OnKeyPress = FD11KeyPress
        end
        object BILBALSITCOMP: THCritMaskEdit
          Left = 145
          Top = 79
          Width = 204
          Height = 21
          Enabled = False
          ParentShowHint = False
          ShowHint = True
          TabOrder = 3
          TagDispatch = 0
          DataType = 'CPBALSIT'
          ElipsisButton = True
          OnElipsisClick = OnBalSitClick
        end
        object BILAVECSITCOMP: TCheckBox
          Left = 36
          Top = 81
          Width = 97
          Height = 17
          Caption = 'sur situation'
          TabOrder = 4
          OnClick = BILAVECSITClick
        end
      end
      object BILCOMP: TCheckBox
        Left = 21
        Top = 158
        Width = 97
        Height = 17
        Caption = 'Avec comparatif'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 2
        OnClick = BILCOMPClick
      end
      object Panel2: TPanel
        Left = 0
        Top = 0
        Width = 572
        Height = 34
        Align = alTop
        BevelOuter = bvLowered
        TabOrder = 3
        object HLabel24: THLabel
          Left = 24
          Top = 11
          Width = 45
          Height = 13
          Caption = 'Maquette'
        end
        object ZMAQUETTENAMEBIL: THLabel
          Left = 140
          Top = 11
          Width = 114
          Height = 13
          Caption = 'ZMAQUETTENAMEBIL'
        end
        object ZMAQUETTEBIL: THCritMaskEdit
          Left = 92
          Top = 7
          Width = 35
          Height = 21
          TabOrder = 0
          OnChange = ZMAQUETTEChange
          TagDispatch = 0
          DataType = 'CPSTDMAQ'
          ElipsisButton = True
          Libelle = ZMAQUETTENAMEBIL
        end
        object FAvecDetailBil: THValComboBox
          Left = 460
          Top = 6
          Width = 105
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          ItemHeight = 13
          ParentCtl3D = False
          TabOrder = 1
          Items.Strings = (
            'Etat seul'
            'D'#233'tail seulement'
            'Etat plus d'#233'tail')
          TagDispatch = 0
          Values.Strings = (
            'NON'
            'SEUL'
            'PLUS')
        end
      end
    end
    object Tab_Analytique: TTabSheet
      Caption = 'Analytique'
      ImageIndex = 2
      object HLabel3: THLabel
        Left = 65
        Top = 68
        Width = 156
        Height = 13
        AutoSize = False
        Caption = '&Axe'
        FocusControl = FNatureCpt
      end
      object pnl_Comptes: TPanel
        Left = 22
        Top = 98
        Width = 549
        Height = 173
        BevelOuter = bvNone
        TabOrder = 2
        Visible = False
        object HLabel6: THLabel
          Left = 40
          Top = 20
          Width = 151
          Height = 13
          AutoSize = False
          Caption = 'Tables'
        end
        object HLabel10: THLabel
          Left = 40
          Top = 74
          Width = 149
          Height = 13
          AutoSize = False
          Caption = 'Tables except'#233'es'
        end
        object edt_Section: TEdit
          Left = 180
          Top = 16
          Width = 255
          Height = 21
          ReadOnly = True
          TabOrder = 0
          OnDblClick = edt_SectionDblClick
          OnKeyDown = edt_SectionKeyDown
        end
        object edt_SectionEx: TEdit
          Left = 180
          Top = 70
          Width = 255
          Height = 21
          ReadOnly = True
          TabOrder = 1
          OnDblClick = edt_SectionDblClick
          OnKeyDown = edt_SectionKeyDown
        end
      end
      object FNatureCpt: THValComboBox
        Left = 202
        Top = 64
        Width = 256
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 1
        OnChange = FNatureCptChange
        TagDispatch = 0
        DataType = 'AFAXE'
        DisableTab = True
      end
      object pnl_TLibres: TPanel
        Left = 21
        Top = 92
        Width = 549
        Height = 72
        BevelOuter = bvNone
        TabOrder = 3
        object TFaG: TLabel
          Left = 304
          Top = 14
          Width = 13
          Height = 17
          AutoSize = False
          Caption = #224
        end
        object TFGenJoker: THLabel
          Left = 40
          Top = 12
          Width = 164
          Height = 13
          AutoSize = False
          Caption = '&Sections'
          Visible = False
        end
        object TFGen: THLabel
          Left = 40
          Top = 12
          Width = 105
          Height = 21
          AutoSize = False
          Caption = '&Sections de'
        end
        object FUnEtatParSection: TCheckBox
          Left = 40
          Top = 32
          Width = 154
          Height = 17
          Caption = 'Un '#233'tat par section'
          TabOrder = 2
        end
        object FLibelleDansLesTitres: TCheckBox
          Left = 40
          Top = 51
          Width = 154
          Height = 17
          Caption = 'Libell'#233' dans les titres'
          TabOrder = 3
        end
        object FSECTION: THCritMaskEdit
          Left = 181
          Top = 12
          Width = 119
          Height = 21
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
          OnChange = FSECTIONChange
          OnKeyDown = FSECTIONKeyDown
          OnKeyPress = FSECTIONKeyPress
          TagDispatch = 0
          DataType = 'TZSECTION'
          Operateur = Superieur
          ElipsisButton = True
        end
        object FSECTION_: THCritMaskEdit
          Left = 320
          Top = 12
          Width = 119
          Height = 21
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
          OnKeyDown = FSECTIONKeyDown
          OnKeyPress = FSECTIONKeyPress
          TagDispatch = 0
          DataType = 'TZSECTION'
          Operateur = Inferieur
          ElipsisButton = True
        end
      end
      object rdg_Analytique: THRadioGroup
        Left = 22
        Top = 8
        Width = 549
        Height = 47
        Caption = 'Analytique'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        OnClick = rdg_AnalytiqueClick
        Columns = 2
        ItemIndex = 0
        Items.Strings = (
          'Sections'
          'Tables libres')
        Abrege = False
        Vide = False
        Values.Strings = (
          'CPT'
          'TLI')
      end
      object FSansANO: TCheckBox
        Left = 62
        Top = 208
        Width = 203
        Height = 17
        Caption = 'Exclure les A-nouveaux analytiques'
        TabOrder = 4
      end
    end
    object Complements: TTabSheet
      Caption = 'Compl'#233'ments'
      ImageIndex = 1
      object HLabel7: THLabel
        Left = 9
        Top = 86
        Width = 111
        Height = 13
        AutoSize = False
        Caption = '&Etablissements'
        FocusControl = FEtab
      end
      object HLabel8: THLabel
        Left = 9
        Top = 121
        Width = 111
        Height = 13
        AutoSize = False
        Caption = '&Devises'
        FocusControl = FDevises
        Visible = False
      end
      object TFRESOL: THLabel
        Left = 9
        Top = 15
        Width = 59
        Height = 13
        Caption = 'R'#233's&olution...'
        FocusControl = FRESOL
      end
      object HLabel5: THLabel
        Left = 326
        Top = 15
        Width = 73
        Height = 13
        Caption = 'Format &montant'
        FocusControl = FFormat
      end
      object tFInfoLibre: THLabel
        Left = 319
        Top = 199
        Width = 50
        Height = 13
        Caption = 'Infos libres'
        FocusControl = FInfoLibre
        Visible = False
      end
      object HLabel13: THLabel
        Left = 344
        Top = 32
        Width = 49
        Height = 13
        Caption = ' (F5 zoom)'
      end
      object HLabel19: THLabel
        Left = 9
        Top = 51
        Width = 93
        Height = 13
        Caption = 'R'#233's&olution... (d'#233'tail)'
        FocusControl = FRESOLDETAIL
      end
      object HLabel20: THLabel
        Left = 326
        Top = 51
        Width = 66
        Height = 13
        Caption = 'Format (d'#233'tail)'
        FocusControl = FFormatDetail
      end
      object HLabel28: THLabel
        Left = 319
        Top = 245
        Width = 92
        Height = 13
        Caption = 'Mention en filigrane'
        FocusControl = FMentionFiligrane
      end
      object lCORRESPONDANCE: THLabel
        Left = 8
        Top = 156
        Width = 116
        Height = 13
        Caption = 'Plan de correspondance'
      end
      object FDevises: THValComboBox
        Left = 149
        Top = 117
        Width = 164
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 5
        Visible = False
        TagDispatch = 0
        Vide = True
        DataType = 'TTDEVISE'
      end
      object FEtab: THValComboBox
        Left = 149
        Top = 82
        Width = 163
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        Vide = True
        DataType = 'TTETABLISSEMENT'
      end
      object FAffiche: TRadioGroup
        Left = 367
        Top = 148
        Width = 211
        Height = 34
        Caption = ' Affichage en... '
        Columns = 3
        ItemIndex = 0
        Items.Strings = (
          '&Euro'
          '&Devise')
        TabOrder = 7
        OnClick = FAfficheClick
      end
      object FAvec: TGroupBox
        Left = 9
        Top = 177
        Width = 304
        Height = 56
        Caption = 'Int'#233'grer les '#233'critures de ... '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 8
        object FSimu: TCheckBox
          Left = 12
          Top = 24
          Width = 69
          Height = 17
          Caption = 'Simulation'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object FSitu: TCheckBox
          Left = 88
          Top = 24
          Width = 69
          Height = 17
          Caption = 'Situation'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
        object FRevi: TCheckBox
          Left = 156
          Top = 24
          Width = 69
          Height = 17
          Caption = 'R'#233'vision'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
        end
        object FPrevi: TCheckBox
          Left = 226
          Top = 24
          Width = 69
          Height = 17
          Caption = 'Pr'#233'vision'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
        end
        object FIFRS: TCheckBox
          Left = 12
          Top = 46
          Width = 69
          Height = 17
          Caption = 'IAS / IFRS'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 4
        end
      end
      object FVoirNom: TCheckBox
        Left = 11
        Top = 243
        Width = 134
        Height = 17
        Caption = 'Voir le nom des cellules'
        TabOrder = 9
      end
      object FRESOL: THValComboBox
        Left = 149
        Top = 11
        Width = 163
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = FRESOLChange
        Items.Strings = (
          'Avec d'#233'cimales'
          'Sans d'#233'cimales'
          'en K(ilo)'
          'en M('#233'ga)')
        TagDispatch = 0
        Values.Strings = (
          'C'
          'F'
          'K'
          'M')
      end
      object FFormat: TEdit
        Left = 404
        Top = 11
        Width = 174
        Height = 21
        TabOrder = 2
        OnKeyDown = FFormatKeyDown
      end
      object FMontant0: TCheckBox
        Left = 367
        Top = 70
        Width = 113
        Height = 17
        Caption = 'Avec montants '#224' 0'
        TabOrder = 10
      end
      object FInfoLibre: TEdit
        Left = 413
        Top = 195
        Width = 165
        Height = 21
        TabOrder = 11
        Visible = False
      end
      object FRESOLDETAIL: THValComboBox
        Left = 149
        Top = 47
        Width = 163
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 1
        OnChange = FRESOLDetailChange
        Items.Strings = (
          'Avec d'#233'cimales'
          'Sans d'#233'cimales'
          'en K(ilo)'
          'en M('#233'ga)')
        TagDispatch = 0
        Values.Strings = (
          'C'
          'F'
          'K'
          'M')
      end
      object FFormatDetail: TEdit
        Left = 404
        Top = 47
        Width = 174
        Height = 21
        TabOrder = 3
        OnKeyDown = FFormatDetailKeyDown
      end
      object FSansCompte: TCheckBox
        Left = 367
        Top = 90
        Width = 178
        Height = 17
        Caption = 'Affichage du libell'#233' des comptes'
        TabOrder = 12
      end
      object FMentionFiligrane: TEdit
        Left = 413
        Top = 241
        Width = 165
        Height = 21
        TabOrder = 13
      end
      object FLibelleSection: TCheckBox
        Left = 367
        Top = 110
        Width = 178
        Height = 17
        Caption = 'Affichage du libell'#233' des sections'
        TabOrder = 14
      end
      object FImpressionDate: TCheckBox
        Left = 367
        Top = 131
        Width = 186
        Height = 17
        Caption = 'Impression de la date'
        TabOrder = 15
      end
      object FCORRESPONDANCE: THValComboBox
        Left = 149
        Top = 152
        Width = 164
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 6
        TagDispatch = 0
        Vide = True
        VideString = '<< Aucun >>'
        DataType = 'CPPLANCORRESPONDANCE'
      end
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 281
    Width = 580
    Height = 76
    AllowDrag = False
    Position = dpBottom
    object PanelFiltre: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 580
      Caption = 'Filtres'
      ClientAreaHeight = 32
      ClientAreaWidth = 580
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BFiltre: TToolbarButton97
        Left = 6
        Top = 6
        Width = 56
        Height = 21
        Hint = 'Menu filtre'
        Caption = '&Filtres'
        DropdownArrow = True
        DropdownMenu = POPF
        Layout = blGlyphRight
        ParentShowHint = False
        ShowHint = True
      end
      object FFiltres: THValComboBox
        Left = 72
        Top = 6
        Width = 465
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = FFiltresChange
        TagDispatch = 0
      end
    end
    object HPB: TToolWindow97
      Left = 0
      Top = 36
      ClientHeight = 36
      ClientWidth = 580
      Caption = 'Barre d'#39'outils'
      ClientAreaHeight = 36
      ClientAreaWidth = 580
      DockPos = 0
      DockRow = 1
      FullSize = True
      TabOrder = 1
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
        OnClick = BParamListeClick
        GlobalIndexImage = 'Z0714_S16G1'
      end
      object BValider: TToolbarButton97
        Left = 494
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
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 526
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
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 558
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
      object BExport: TToolbarButton97
        Left = 37
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Exporter au format Excel'
        Caption = 'Imprimer'
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
        OnClick = BExportClick
        GlobalIndexImage = 'Z0724_S16G1'
      end
    end
  end
  object FVariation: TRadioGroup
    Left = 444
    Top = 0
    Width = 141
    Height = 105
    Caption = 'Variation'
    Items.Strings = (
      'Simple'
      'Sur 12 mois'
      'Ramen'#233' au plus petit'
      'Ramen'#233' au plus grand')
    TabOrder = 2
    Visible = False
    OnClick = FVariationClick
  end
  object PFormat: TPanel
    Left = 180
    Top = 47
    Width = 321
    Height = 141
    TabOrder = 3
    Visible = False
    object H_TitreGuide: TLabel
      Left = -8
      Top = 9
      Width = 313
      Height = 18
      Alignment = taCenter
      AutoSize = False
      Caption = 'Format des montants dans l'#39#233'tat'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = False
      ParentFont = False
    end
    object EMP: TLabel
      Tag = 1
      Left = 232
      Top = 35
      Width = 73
      Height = 13
      Alignment = taRightJustify
      AutoSize = False
      Caption = '1000.00'
    end
    object EMN: TLabel
      Tag = 2
      Left = 232
      Top = 63
      Width = 73
      Height = 13
      Alignment = taRightJustify
      AutoSize = False
      Caption = '1000.00'
    end
    object EMU: TLabel
      Tag = 3
      Left = 232
      Top = 91
      Width = 73
      Height = 13
      Alignment = taRightJustify
      AutoSize = False
      Caption = '1000.00'
    end
    object PFenGuide: TPanel
      Left = 1
      Top = 1
      Width = 319
      Height = 9
      Align = alTop
      BevelInner = bvLowered
      BevelOuter = bvNone
      Color = clActiveCaption
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
    end
    object BCValide: THBitBtn
      Tag = 1
      Left = 253
      Top = 109
      Width = 28
      Height = 27
      Hint = 'Enregistrer le guide'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BCValideClick
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      GlobalIndexImage = 'Z0003_S16G2'
      IsControl = True
    end
    object BCAbandon: THBitBtn
      Tag = 1
      Left = 284
      Top = 109
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BCAbandonClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0021_S16G1'
      IsControl = True
    end
    object CBMP: TCheckBox
      Tag = 1
      Left = 24
      Top = 33
      Width = 97
      Height = 17
      Caption = 'Montant positif'
      TabOrder = 3
      OnClick = CBMPClick
    end
    object FMP: TEdit
      Tag = 1
      Left = 124
      Top = 31
      Width = 96
      Height = 21
      TabOrder = 4
      OnChange = FMPChange
    end
    object CBMN: TCheckBox
      Tag = 2
      Left = 24
      Top = 61
      Width = 97
      Height = 17
      Caption = 'Montant n'#233'gatif'
      TabOrder = 5
      OnClick = CBMPClick
    end
    object FMN: TEdit
      Tag = 2
      Left = 124
      Top = 59
      Width = 96
      Height = 21
      TabOrder = 6
      OnChange = FMPChange
    end
    object CBMU: TCheckBox
      Tag = 3
      Left = 24
      Top = 89
      Width = 97
      Height = 17
      Caption = 'Montant nul'
      TabOrder = 7
      OnClick = CBMPClick
    end
    object FMU: TEdit
      Tag = 3
      Left = 124
      Top = 87
      Width = 96
      Height = 21
      TabOrder = 8
      OnChange = FMPChange
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Etats de synth'#232'se'
      
        '1;?Caption?;La maquette d'#39#233'tat choisie n'#39'est pas correcte.;W;O;O' +
        ';O;'
      
        '2;R'#233'partition des '#233'ch'#233'ances;Vous ne pouvez pas renseigner des mo' +
        'ntants n'#233'gatifs.;W;O;O;O;'
      
        '3;R'#233'partition des '#233'ch'#233'ances;Les dates d'#39#233'ch'#233'ance doivent respect' +
        'er la plage de saisie autoris'#233'e;W;O;O;O;'
      'Bilan'
      'Compte de r'#233'sultat'
      'S.I.G.'
      'Compte de r'#233'sultat analytique'
      'S.I.G. analytique'
      'Bilan analytique'
      '10;?Caption?;La maquette d'#39#233'tat choisie n'#39'existe pas.;W;O;O;O;'
      '11;?Caption?;Vous devez renseigner un axe analytique.;W;O;O;O;'
      
        '12,?Caption?,Les dates d'#39'exercices ne sont pas correctes.;E;O;O;' +
        'O;'
      ''
      '')
    Left = 488
    Top = 292
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 420
    Top = 300
  end
  object POPF: TPopupMenu
    Left = 452
    Top = 296
    object BCreerFiltre: TMenuItem
      Caption = '&Cr'#233'er un filtre'
      OnClick = BCreerFiltreClick
    end
    object BSaveFiltre: TMenuItem
      Caption = '&Enregistrer le filtre'
      OnClick = BSaveFiltreClick
    end
    object Dupliquerlefiltre1: TMenuItem
      Caption = '&Dupliquer le filtre'
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
  object Sauve: TSaveDialog
    Filter = 'Fichiers textes (*.txt).|*.txt|Tous fichiers (*.*)|*.*'
    Options = [ofHideReadOnly, ofShowHelp, ofExtensionDifferent, ofPathMustExist, ofNoLongNames]
    Title = 'Exportation des mouvements'
    Left = 384
    Top = 300
  end
  object Timer1: TTimer
    Enabled = False
    Interval = 10
    OnTimer = Timer1Timer
    Left = 280
    Top = 26
  end
end

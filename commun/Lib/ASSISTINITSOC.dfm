inherited FAssistInitSoc: TFAssistInitSoc
  Left = 349
  Top = 158
  Caption = 'Assistant d'#39'initialisation de soci'#233't'#233
  ClientHeight = 390
  ClientWidth = 591
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Left = 4
    Top = 353
    Visible = False
  end
  inherited lAide: THLabel
    Left = 167
    Top = 264
  end
  inherited bPrecedent: TToolbarButton97
    Left = 337
    Top = 344
  end
  inherited bSuivant: TToolbarButton97
    Left = 412
    Top = 344
  end
  inherited bFin: TToolbarButton97
    Left = 497
    Top = 344
  end
  inherited bAnnuler: TToolbarButton97
    Left = 253
    Top = 344
  end
  inherited bAide: TToolbarButton97
    Left = 173
    Top = 344
  end
  inherited Plan: THPanel
    Left = 168
    Top = 1
  end
  inherited GroupBox1: THGroupBox
    Left = -6
    Top = 333
    Width = 583
  end
  inherited P: THPageControl2
    Left = 157
    Top = 4
    Width = 400
    Height = 329
    ActivePage = TabSheet0
    object TABGEN: TTabSheet
      Caption = 'M'#233'tier'
      Constraints.MaxHeight = 301
      Constraints.MaxWidth = 393
      object BMetier: TBevel
        Left = 0
        Top = 41
        Width = 392
        Height = 260
        Align = alClient
      end
      object TINTRO: THLabel
        Left = 12
        Top = 69
        Width = 365
        Height = 25
        AutoSize = False
        Caption = 
          'Cet assistant vous guide afin de proc'#233'der '#224' une initialisation d' +
          'e soci'#233't'#233
        WordWrap = True
      end
      object PTITRE: THPanel
        Left = 0
        Top = 0
        Width = 392
        Height = 41
        Align = alTop
        Caption = 'Initialisation de soci'#233't'#233
        FullRepaint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -16
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
      end
      object GB_DOM: TGroupBox
        Left = 10
        Top = 117
        Width = 373
        Height = 117
        Caption = 'Domaines'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object CB_DOM1: TCheckBox
          Left = 19
          Top = 23
          Width = 145
          Height = 17
          Caption = 'Gestion commerciale'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 0
        end
        object CB_DOM2: TCheckBox
          Left = 225
          Top = 23
          Width = 89
          Height = 17
          Caption = 'Comptabilit'#233
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 1
        end
        object CB_DOM3: TCheckBox
          Left = 19
          Top = 57
          Width = 89
          Height = 17
          Caption = 'Affaires'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 2
        end
        object CB_DOM4: TCheckBox
          Left = 225
          Top = 57
          Width = 89
          Height = 17
          Caption = 'Mode'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 3
        end
        object CB_DOM5: TCheckBox
          Left = 19
          Top = 89
          Width = 103
          Height = 17
          Caption = 'Relation clients'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 4
        end
      end
    end
    object TABINIT: TTabSheet
      Caption = 'Utilisat'
      ImageIndex = 1
      object BInfos: TBevel
        Left = 0
        Top = 0
        Width = 392
        Height = 301
        Align = alClient
      end
      object TFRef: THLabel
        Left = 67
        Top = 13
        Width = 118
        Height = 13
        Caption = '&Base de donn'#233'es source'
        FocusControl = FRef
      end
      object FRef: THValComboBox
        Left = 207
        Top = 9
        Width = 175
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        Items.Strings = (
          '')
        TagDispatch = 0
      end
      object GSOCIETE: TGroupBox
        Left = 4
        Top = 35
        Width = 385
        Height = 53
        Hint = 'SO_CLE1'
        Caption = ' Soci'#233't'#233' '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object TSO_SOCIETE: THLabel
          Left = 12
          Top = 23
          Width = 25
          Height = 13
          Caption = '&Code'
          FocusControl = SO_SOCIETE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TSO_LIBELLE: THLabel
          Left = 111
          Top = 23
          Width = 69
          Height = 13
          Caption = '&Raison sociale'
          FocusControl = SO_LIBELLE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object SO_SOCIETE: TEdit
          Left = 53
          Top = 19
          Width = 49
          Height = 21
          CharCase = ecUpperCase
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 3
          ParentFont = False
          TabOrder = 0
        end
        object SO_LIBELLE: TEdit
          Left = 203
          Top = 19
          Width = 175
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
      end
      object GUSERGRP: TGroupBox
        Left = 4
        Top = 91
        Width = 384
        Height = 53
        Hint = 'UG_CLE1'
        Caption = ' Groupe d'#39'utillisateurs '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        object TUG_GROUPE: THLabel
          Left = 12
          Top = 23
          Width = 25
          Height = 13
          Caption = 'C&ode'
          FocusControl = UG_GROUPE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TUG_LIBELLE: THLabel
          Left = 113
          Top = 23
          Width = 73
          Height = 13
          Caption = '&Nom du groupe'
          FocusControl = UG_LIBELLE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object UG_GROUPE: TEdit
          Left = 53
          Top = 19
          Width = 49
          Height = 21
          CharCase = ecUpperCase
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 3
          ParentFont = False
          TabOrder = 0
          OnChange = UG_GROUPEChange
        end
        object UG_NIVEAUACCES: THNumEdit
          Left = 147
          Top = 8
          Width = 13
          Height = 21
          TabStop = False
          Color = clYellow
          Decimals = 0
          Digits = 12
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Masks.PositiveMask = '#,##0'
          Debit = False
          ParentFont = False
          TabOrder = 1
          UseRounding = True
          Value = 9.000000000000000000
          Validate = False
          Visible = False
        end
        object UG_NUMERO: THNumEdit
          Left = 159
          Top = 8
          Width = 13
          Height = 21
          TabStop = False
          Color = clYellow
          Decimals = 0
          Digits = 12
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Masks.PositiveMask = '#,##0'
          Debit = False
          ParentFont = False
          TabOrder = 2
          UseRounding = True
          Value = 1.000000000000000000
          Validate = False
          Visible = False
        end
        object UG_LIBELLE: TEdit
          Left = 202
          Top = 19
          Width = 175
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
        end
      end
      object GUTILISAT: TGroupBox
        Left = 4
        Top = 148
        Width = 383
        Height = 89
        Hint = 'US_CLE1'
        Caption = ' Utilisateur '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        object TUS_UTILISATEUR: THLabel
          Left = 12
          Top = 20
          Width = 25
          Height = 13
          Caption = 'Co&de'
          FocusControl = US_UTILISATEUR
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TUS_PASSWORD: THLabel
          Left = 222
          Top = 43
          Width = 64
          Height = 13
          Caption = '&Mot de passe'
          FocusControl = US_PASSWORD
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TUS_ABREGE: THLabel
          Left = 12
          Top = 43
          Width = 26
          Height = 13
          Caption = '&Login'
          FocusControl = US_ABREGE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TUS_LIBELLE: THLabel
          Left = 111
          Top = 20
          Width = 88
          Height = 13
          Caption = 'Nom de l'#39'&utilisateur'
          FocusControl = US_LIBELLE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TUS_PASSWORD2: THLabel
          Left = 222
          Top = 67
          Width = 58
          Height = 13
          Caption = 'Con&firmation'
          FocusControl = FPassWord2
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object US_UTILISATEUR: TEdit
          Left = 52
          Top = 15
          Width = 49
          Height = 21
          CharCase = ecUpperCase
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 3
          ParentFont = False
          TabOrder = 0
        end
        object US_PASSWORD: TEdit
          Left = 300
          Top = 39
          Width = 77
          Height = 21
          CharCase = ecUpperCase
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          PasswordChar = '*'
          TabOrder = 3
        end
        object US_ABREGE: TEdit
          Left = 52
          Top = 39
          Width = 151
          Height = 21
          CharCase = ecUpperCase
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 17
          ParentFont = False
          TabOrder = 2
        end
        object US_GROUPE: TEdit
          Left = 134
          Top = 63
          Width = 19
          Height = 21
          TabStop = False
          CharCase = ecUpperCase
          Color = clYellow
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 3
          ParentFont = False
          TabOrder = 4
          Visible = False
        end
        object US_SUPERVISEUR: TEdit
          Left = 154
          Top = 63
          Width = 17
          Height = 21
          TabStop = False
          CharCase = ecUpperCase
          Color = clYellow
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 3
          ParentFont = False
          TabOrder = 5
          Text = 'X'
          Visible = False
        end
        object US_LIBELLE: TEdit
          Left = 202
          Top = 15
          Width = 175
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnChange = US_LIBELLEChange
        end
        object FPassWord2: TEdit
          Left = 300
          Top = 63
          Width = 77
          Height = 21
          CharCase = ecUpperCase
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 8
          ParentFont = False
          PasswordChar = '*'
          TabOrder = 6
        end
      end
    end
    object TabSheet0: TTabSheet
      Caption = 'Param g'#233'n'#233'raux'
      ImageIndex = 4
      object BGen: TBevel
        Left = 0
        Top = 0
        Width = 392
        Height = 301
        Align = alClient
      end
      object GBAdmin: TGroupBox
        Left = 6
        Top = 179
        Width = 378
        Height = 52
        Caption = ' Administrateurs '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 10
        object USERGRP: TCheckBox
          Left = 18
          Top = 23
          Width = 144
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Groupes d'#39'utilisateurs'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnClick = USERGRPClick
        end
        object UTILISAT: TCheckBox
          Left = 221
          Top = 23
          Width = 143
          Height = 18
          Alignment = taLeftJustify
          Caption = 'Utilisateurs'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = UTILISATClick
        end
      end
      object HPanel1: THPanel
        Left = 22
        Top = 5
        Width = 350
        Height = 24
        Caption = 'Param'#232'tres g'#233'n'#233'raux'
        FullRepaint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -15
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 11
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
      end
      object PAYS: TCheckBox
        Left = 227
        Top = 54
        Width = 145
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Pa&ys'
        Checked = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        State = cbChecked
        TabOrder = 5
      end
      object REGION: TCheckBox
        Left = 227
        Top = 76
        Width = 145
        Height = 17
        Alignment = taLeftJustify
        Caption = '&R'#233'gions'
        Checked = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        State = cbChecked
        TabOrder = 6
      end
      object CODEPOST: TCheckBox
        Left = 227
        Top = 98
        Width = 145
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Codes postaux'
        Checked = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        State = cbChecked
        TabOrder = 7
      end
      object CODEAFB: TCheckBox
        Left = 22
        Top = 98
        Width = 145
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Codes AFB'
        Checked = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        State = cbChecked
        TabOrder = 2
      end
      object DEVISE: TCheckBox
        Left = 22
        Top = 76
        Width = 145
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Devises'
        Checked = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        State = cbChecked
        TabOrder = 1
      end
      object MODEPAIE: TCheckBox
        Left = 22
        Top = 142
        Width = 145
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Modes de paiement'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        OnClick = MODEPAIEClick
      end
      object MODEREGL: TCheckBox
        Left = 22
        Top = 120
        Width = 145
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Modes de r'#232'glements'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        OnClick = MODEREGLClick
      end
      object MODELES: TCheckBox
        Left = 227
        Top = 120
        Width = 145
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Mod'#232'les de documents'
        Checked = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        State = cbChecked
        TabOrder = 8
        OnClick = MODELESClick
      end
      object ETABLISS: TCheckBox
        Left = 22
        Top = 54
        Width = 145
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Etablissements'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
      end
      object MODEDATA: TCheckBox
        Left = 227
        Top = 142
        Width = 145
        Height = 17
        Alignment = taLeftJustify
        Caption = 'D'#233'tail des mod'#232'les'
        Checked = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        State = cbChecked
        TabOrder = 9
        Visible = False
      end
    end
    object TabSheet1: TTabSheet
      Caption = 'Gescom'
      ImageIndex = 2
      object BGes: TBevel
        Left = 0
        Top = 0
        Width = 392
        Height = 301
        Align = alClient
      end
      object GBTab1: TGroupBox
        Left = 6
        Top = 36
        Width = 186
        Height = 203
        Caption = 'Tables'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object PARPIECE: TCheckBox
          Left = 5
          Top = 24
          Width = 174
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Param'#232'trage des types de pi'#232'ce'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object SOUCHE: TCheckBox
          Left = 5
          Top = 51
          Width = 174
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Souches'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
        object DIMMASQUE: TCheckBox
          Left = 10
          Top = 179
          Width = 19
          Height = 17
          Alignment = taLeftJustify
          Caption = 'DimMasque'
          Color = clYellow
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 2
          Visible = False
        end
        object DIMENSION: TCheckBox
          Left = 5
          Top = 77
          Width = 174
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Dimensions'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          OnClick = DIMENSIONClick
        end
      end
      object GBTT1: TGroupBox
        Left = 199
        Top = 36
        Width = 186
        Height = 203
        Caption = 'Tablettes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object TT_GCCATEGORIETAXE: TCheckBox
          Left = 5
          Top = 24
          Width = 174
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Cat'#233'gories taxes'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object TT_DIMENSIONS: TCheckBox
          Left = 5
          Top = 50
          Width = 174
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Dimensions'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnClick = TT_DIMENSIONSClick
        end
        object TT_GCLIBFAMILLE: TCheckBox
          Left = 5
          Top = 75
          Width = 174
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Libell'#233's familles'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
        end
        object TT_GCZONELIBRE: TCheckBox
          Left = 5
          Top = 101
          Width = 174
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Param'#232'trage des zones libres'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
        end
        object TT_GCCATEGORIEDIM: TCheckBox
          Left = 10
          Top = 179
          Width = 19
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Cat'#233'gories des dimensions'
          Color = clYellow
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 4
          Visible = False
        end
      end
      object PTitre1: THPanel
        Left = 22
        Top = 5
        Width = 350
        Height = 24
        Caption = 'Gestion commerciale'
        FullRepaint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -15
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'Compta'
      ImageIndex = 3
      object BCpta: TBevel
        Left = 0
        Top = 0
        Width = 392
        Height = 301
        Align = alClient
      end
      object GBTT2: TGroupBox
        Left = 199
        Top = 36
        Width = 186
        Height = 203
        Caption = 'Tablettes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
      end
      object GBTab2: TGroupBox
        Left = 6
        Top = 36
        Width = 186
        Height = 203
        Caption = 'Tables'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object EXERCICE: TCheckBox
          Left = 9
          Top = 27
          Width = 154
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Exercices'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object SUIVCPTA: TCheckBox
          Left = 9
          Top = 52
          Width = 154
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Sc'#233'narios de saisie'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
        object CORRESP: TCheckBox
          Left = 9
          Top = 77
          Width = 154
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Plan des correspondances'
          Checked = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          State = cbChecked
          TabOrder = 2
        end
        object REFAUTO: TCheckBox
          Left = 9
          Top = 102
          Width = 154
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Libell'#233's automatiques'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
        end
        object GUIDES: TCheckBox
          Left = 9
          Top = 126
          Width = 154
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Guides'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 4
        end
        object RUBRIQUE: TCheckBox
          Left = 9
          Top = 176
          Width = 154
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Rubriques'
          Checked = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          State = cbChecked
          TabOrder = 6
        end
        object RUPTURE: TCheckBox
          Left = 9
          Top = 151
          Width = 154
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Plan des Ruptures'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 5
        end
      end
      object PTitre2: THPanel
        Left = 22
        Top = 5
        Width = 350
        Height = 24
        Caption = 'Comptabilit'#233
        FullRepaint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -15
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
      end
    end
    object TabSheet3: TTabSheet
      Caption = 'Affaires'
      ImageIndex = 5
      object BAff: TBevel
        Left = 0
        Top = 0
        Width = 392
        Height = 301
        Align = alClient
      end
      object GBTab3: TGroupBox
        Left = 6
        Top = 36
        Width = 186
        Height = 203
        Caption = 'Tables'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
      end
      object GBTT3: TGroupBox
        Left = 199
        Top = 36
        Width = 186
        Height = 203
        Caption = 'Tablettes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object TT_AFTRADUCAFFAIRE: TCheckBox
          Left = 5
          Top = 24
          Width = 174
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Traduction affaires'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object TT_AFTTYPEHEURE: TCheckBox
          Left = 5
          Top = 48
          Width = 174
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Type d'#39'heures'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
        object TT_AFTTYPERESSOURCE: TCheckBox
          Left = 5
          Top = 71
          Width = 174
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Type de ressource '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
        end
      end
      object PTitre3: THPanel
        Left = 22
        Top = 5
        Width = 350
        Height = 24
        Caption = 'Affaires'
        FullRepaint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -15
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
      end
    end
    object TabSheet4: TTabSheet
      Caption = 'Mode'
      ImageIndex = 8
      object BMode: TBevel
        Left = 0
        Top = 0
        Width = 392
        Height = 301
        Align = alClient
      end
      object GBTab4: TGroupBox
        Left = 6
        Top = 36
        Width = 186
        Height = 203
        Caption = 'Tables'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
      end
      object GBTT4: TGroupBox
        Left = 199
        Top = 36
        Width = 186
        Height = 203
        Caption = 'Tablettes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
      end
      object PTitre4: THPanel
        Left = 22
        Top = 5
        Width = 350
        Height = 24
        Caption = 'Mode'
        FullRepaint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -15
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
      end
    end
    object TabSheet5: TTabSheet
      Caption = 'RCli'
      ImageIndex = 6
      object BRCli: TBevel
        Left = 0
        Top = 0
        Width = 392
        Height = 301
        Align = alClient
        Visible = False
      end
      object GBTab5: TGroupBox
        Left = 6
        Top = 36
        Width = 186
        Height = 261
        Caption = 'Tables'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object CHAMPSPRO: TCheckBox
          Left = 4
          Top = 20
          Width = 177
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Param. Champs libres  Prospects'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = CHAMPSPROClick
        end
        object PARACTIONS: TCheckBox
          Left = 4
          Top = 37
          Width = 177
          Height = 23
          Alignment = taLeftJustify
          Caption = 'Param'#233'trage des types d'#39'actions'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnClick = PARACTIONSClick
        end
        object PARSUSPECT: TCheckBox
          Left = 4
          Top = 60
          Width = 177
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Param'#233'trage des Suspects'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          OnClick = PARSUSPECTClick
        end
        object PARSUSPECTLIG: TCheckBox
          Left = 5
          Top = 81
          Width = 177
          Height = 17
          Alignment = taLeftJustify
          Caption = 'D'#233'tail Param'#233'trage des Suspects'
          Color = clYellow
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 3
          Visible = False
        end
      end
      object GBTT5: TGroupBox
        Left = 199
        Top = 36
        Width = 186
        Height = 38
        Caption = ' Tablettes Table Compl'#233'mentaire '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object TT_RTLIBTABLECOMPL: TCheckBox
          Left = 4
          Top = 14
          Width = 177
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Param Col. Table Compl'#233'mentaire'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = TT_RTLIBTABLECOMPLClick
        end
      end
      object PTitre5: THPanel
        Left = 22
        Top = 5
        Width = 350
        Height = 24
        Caption = 'Relation clients'
        FullRepaint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -15
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
      end
      object GroupBox2: TGroupBox
        Left = 199
        Top = 75
        Width = 186
        Height = 70
        Caption = ' Tablettes Propositions '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        object TT_RTTYPEPERSPECTIVE: TCheckBox
          Left = 4
          Top = 31
          Width = 177
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Types de propositions'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object RTRPRLIBPERSPECTIVE: TCheckBox
          Left = 4
          Top = 48
          Width = 177
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Tables libres (3)'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnClick = RTRPRLIBPERSPECTIVEClick
        end
        object TT_RTMOTIFS: TCheckBox
          Left = 4
          Top = 14
          Width = 177
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Motifs de Signature/Perte'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
        end
      end
      object TT_RTRPCLIBTABLE0: TCheckBox
        Left = 12
        Top = 159
        Width = 21
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 4
        Visible = False
      end
      object TT_RTRPCLIBTABLE1: TCheckBox
        Left = 28
        Top = 159
        Width = 21
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 5
        Visible = False
      end
      object TT_RTRPCLIBTABLE2: TCheckBox
        Left = 41
        Top = 159
        Width = 21
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 6
        Visible = False
      end
      object TT_RTRPCLIBTABLE3: TCheckBox
        Left = 61
        Top = 159
        Width = 21
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 7
        Visible = False
      end
      object TT_RTRPCLIBTABLE4: TCheckBox
        Left = 77
        Top = 159
        Width = 21
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 8
        Visible = False
      end
      object TT_RTRPCLIBTABLE5: TCheckBox
        Left = 54
        Top = 179
        Width = 21
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 9
        Visible = False
      end
      object TT_RTRPCLIBTABLE6: TCheckBox
        Left = 97
        Top = 159
        Width = 21
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 10
        Visible = False
      end
      object TT_RTRPCLIBTABLE7: TCheckBox
        Left = 115
        Top = 159
        Width = 21
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 11
        Visible = False
      end
      object TT_RTRPCLIBTABLE8: TCheckBox
        Left = 127
        Top = 159
        Width = 21
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 12
        Visible = False
      end
      object TT_RTRPCLIBTABLE9: TCheckBox
        Left = 143
        Top = 159
        Width = 21
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 13
        Visible = False
      end
      object TT_RTRPCLIBTABLEA: TCheckBox
        Left = 11
        Top = 179
        Width = 21
        Height = 18
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 14
        Visible = False
      end
      object TT_RTRPCLIBTABLEB: TCheckBox
        Left = 27
        Top = 179
        Width = 21
        Height = 18
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 15
        Visible = False
      end
      object TT_RTRPRLIBPERSPECTIVE1: TCheckBox
        Left = 76
        Top = 178
        Width = 25
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 16
        Visible = False
      end
      object TT_RTRPRLIBPERSPECTIVE2: TCheckBox
        Left = 100
        Top = 178
        Width = 25
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 17
        Visible = False
      end
      object TT_RTRPRLIBPERSPECTIVE3: TCheckBox
        Left = 128
        Top = 178
        Width = 25
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 18
        Visible = False
      end
      object GroupBox3: TGroupBox
        Left = 199
        Top = 147
        Width = 186
        Height = 54
        Caption = ' Tablettes Actions '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 19
        object TT_RTIMPORTANCEACT: TCheckBox
          Left = 4
          Top = 14
          Width = 177
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Niveaux d'#39'importance'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object RTRPRLIBACTION: TCheckBox
          Left = 4
          Top = 31
          Width = 177
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Tables libres (3)'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnClick = RTRPRLIBACTIONClick
        end
      end
      object TT_RTRPRLIBACTION1: TCheckBox
        Left = 12
        Top = 198
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 20
        Visible = False
      end
      object TT_RTRPRLIBACTION2: TCheckBox
        Left = 31
        Top = 198
        Width = 19
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 21
        Visible = False
      end
      object TT_RTRPRLIBACTION3: TCheckBox
        Left = 49
        Top = 198
        Width = 19
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 22
        Visible = False
      end
      object TT_RTLIBCHAMPSLIBRES: TCheckBox
        Left = 69
        Top = 198
        Width = 16
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 23
        Visible = False
      end
      object TT_RTRPRLIBTABLE0: TCheckBox
        Left = 12
        Top = 214
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 24
        Visible = False
      end
      object TT_RTRPRLIBTABLE1: TCheckBox
        Left = 22
        Top = 218
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 25
        Visible = False
      end
      object TT_RTRPRLIBTABLE2: TCheckBox
        Left = 34
        Top = 218
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 26
        Visible = False
      end
      object TT_RTRPRLIBTABLE3: TCheckBox
        Left = 45
        Top = 218
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 27
        Visible = False
      end
      object TT_RTRPRLIBTABLE4: TCheckBox
        Left = 57
        Top = 218
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 28
        Visible = False
      end
      object TT_RTRPRLIBTABLE5: TCheckBox
        Left = 71
        Top = 218
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 29
        Visible = False
      end
      object TT_RTRPRLIBTABLE6: TCheckBox
        Left = 85
        Top = 219
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 30
        Visible = False
      end
      object TT_RTRPRLIBTABLE7: TCheckBox
        Left = 97
        Top = 218
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 31
        Visible = False
      end
      object TT_RTRPRLIBTABLE8: TCheckBox
        Left = 110
        Top = 219
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 32
        Visible = False
      end
      object TT_RTRPRLIBTABLE9: TCheckBox
        Left = 122
        Top = 220
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 33
        Visible = False
      end
      object TT_RTRPRLIBTABLE10: TCheckBox
        Left = 83
        Top = 199
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 34
        Visible = False
      end
      object TT_RTRPRLIBTABLE11: TCheckBox
        Left = 99
        Top = 199
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 35
        Visible = False
      end
      object TT_RTRPRLIBTABLE12: TCheckBox
        Left = 113
        Top = 198
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 36
        Visible = False
      end
      object TT_RTRPRLIBTABLE13: TCheckBox
        Left = 126
        Top = 197
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 37
        Visible = False
      end
      object TT_RTRPRLIBTABLE14: TCheckBox
        Left = 138
        Top = 196
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 38
        Visible = False
      end
      object TT_RTRPRLIBTABLE15: TCheckBox
        Left = 150
        Top = 197
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 39
        Visible = False
      end
      object TT_RTRPRLIBTABLE16: TCheckBox
        Left = 162
        Top = 198
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 40
        Visible = False
      end
      object TT_RTRPRLIBTABLE17: TCheckBox
        Left = 175
        Top = 197
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 41
        Visible = False
      end
      object TT_RTRPRLIBTABLE18: TCheckBox
        Left = 143
        Top = 217
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 42
        Visible = False
      end
      object TT_RTRPRLIBTABLE20: TCheckBox
        Left = 171
        Top = 221
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 43
        Visible = False
      end
      object TT_RTRPRLIBTABLE19: TCheckBox
        Left = 155
        Top = 217
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 44
        Visible = False
      end
      object TT_RTRPRLIBTABLE21: TCheckBox
        Left = 11
        Top = 237
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 45
        Visible = False
      end
      object TT_RTRPRLIBTABLE22: TCheckBox
        Left = 23
        Top = 237
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 46
        Visible = False
      end
      object TT_RTRPRLIBTABLE23: TCheckBox
        Left = 35
        Top = 237
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 47
        Visible = False
      end
      object TT_RTRPRLIBTABLE24: TCheckBox
        Left = 47
        Top = 237
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 48
        Visible = False
      end
      object TT_RTRPRLIBTABLE25: TCheckBox
        Left = 63
        Top = 237
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 49
        Visible = False
      end
      object TT_RTLIBCLACTPERSP: TCheckBox
        Left = 176
        Top = 173
        Width = 21
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 50
        Visible = False
      end
      object TT_RTRPRLIBTABMUL1: TCheckBox
        Left = 5
        Top = 252
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 51
        Visible = False
      end
      object TT_RTRPRLIBTABMUL2: TCheckBox
        Left = 17
        Top = 252
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 52
        Visible = False
      end
      object TT_RTRPRLIBTABMUL3: TCheckBox
        Left = 33
        Top = 252
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 53
        Visible = False
      end
      object TT_RTRPRLIBTABMUL4: TCheckBox
        Left = 49
        Top = 252
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 54
        Visible = False
      end
      object TT_RTRPRLIBTABMUL5: TCheckBox
        Left = 65
        Top = 252
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 55
        Visible = False
      end
      object TT_RTRPRLIBTABMUL6: TCheckBox
        Left = 81
        Top = 252
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 56
        Visible = False
      end
      object TT_RTRPRLIBTABMUL7: TCheckBox
        Left = 97
        Top = 252
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 57
        Visible = False
      end
      object TT_RTRPRLIBTABMUL8: TCheckBox
        Left = 113
        Top = 252
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 58
        Visible = False
      end
      object TT_RTRPRLIBTABMUL9: TCheckBox
        Left = 129
        Top = 252
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 59
        Visible = False
      end
      object TT_RTRPRLIBTABMUL0: TCheckBox
        Left = 145
        Top = 252
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 60
        Visible = False
      end
      object GroupBox4: TGroupBox
        Left = 199
        Top = 203
        Width = 186
        Height = 54
        Caption = ' Tablettes Suspects '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 61
        object TT_RTMOTIFFERMETURE: TCheckBox
          Left = 4
          Top = 14
          Width = 177
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Motifs de Fermeture'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object TT_RTLIBCHPLIBSUSPECTS: TCheckBox
          Left = 4
          Top = 31
          Width = 177
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Champs libres'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnClick = TT_RTLIBCHPLIBSUSPECTSClick
        end
      end
      object TT_RTRSCLIBTABLE0: TCheckBox
        Left = 5
        Top = 268
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 62
        Visible = False
      end
      object TT_RTRSCLIBTABLE1: TCheckBox
        Left = 21
        Top = 268
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 63
        Visible = False
      end
      object TT_RTRSCLIBTABLE2: TCheckBox
        Left = 37
        Top = 268
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 64
        Visible = False
      end
      object TT_RTRSCLIBTABLE3: TCheckBox
        Left = 53
        Top = 268
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 65
        Visible = False
      end
      object TT_RTRSCLIBTABLE4: TCheckBox
        Left = 69
        Top = 268
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 66
        Visible = False
      end
      object TT_RTRSCLIBTABLE5: TCheckBox
        Left = 81
        Top = 264
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 67
        Visible = False
      end
      object TT_RTRSCLIBTABLE6: TCheckBox
        Left = 101
        Top = 268
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 68
        Visible = False
      end
      object TT_RTRSCLIBTABLE7: TCheckBox
        Left = 117
        Top = 268
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 69
        Visible = False
      end
      object TT_RTRSCLIBTABLE8: TCheckBox
        Left = 133
        Top = 268
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 70
        Visible = False
      end
      object TT_RTRSCLIBTABLE9: TCheckBox
        Left = 149
        Top = 268
        Width = 17
        Height = 17
        Alignment = taLeftJustify
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 71
        Visible = False
      end
      object GroupBox5: TGroupBox
        Left = 199
        Top = 259
        Width = 186
        Height = 38
        Caption = ' Tablettes Op'#233'rations '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 72
        object TT_RTOBJETOPE: TCheckBox
          Left = 4
          Top = 14
          Width = 177
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Objet de l'#39'op'#233'ration'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = TT_RTLIBTABLECOMPLClick
        end
      end
    end
    object TABRECAP: TTabSheet
      Caption = 'Recap'
      ImageIndex = 5
      object BRecap: TBevel
        Left = 0
        Top = 0
        Width = 392
        Height = 301
        Align = alClient
      end
      object TEtat: THLabel
        Left = 3
        Top = 197
        Width = 386
        Height = 20
        AutoSize = False
        Caption = 'Connexion...'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        Visible = False
        WordWrap = True
      end
      object TProgress: THLabel
        Left = 3
        Top = 217
        Width = 59
        Height = 13
        Caption = 'Connexion...'
        Visible = False
        WordWrap = True
      end
      object EG_Progress: TEnhancedGauge
        Left = 3
        Top = 217
        Width = 385
        Height = 21
        TabOrder = 0
        Visible = False
        ForeColor = clNavy
        BackColor = clBtnFace
        Progress = 0
      end
      object PanelFin: TPanel
        Left = 3
        Top = 3
        Width = 387
        Height = 97
        TabOrder = 1
        object TTextFin1: THLabel
          Left = 15
          Top = 8
          Width = 339
          Height = 26
          Caption = 
            'Le param'#232'trage est correctement renseign'#233' pour permettre l'#39'initi' +
            'alisation de soci'#233't'#233'.'
          WordWrap = True
        end
        object TTextFin2: THLabel
          Left = 15
          Top = 49
          Width = 324
          Height = 26
          Caption = 
            'Si vous d'#233'sirez revoir le param'#233'trage, il suffit de cliquer sur ' +
            'le bouton Pr'#233'c'#233'dent sinon, le bouton Fin, permet de d'#233'buter le t' +
            'raitement.'
          WordWrap = True
        end
      end
    end
  end
  inherited PanelImage: THPanel
    Left = 1
    inherited Image: TToolbarButton97
      Left = 8
      Top = 2
    end
  end
  inherited cControls: THListBox
    Left = 93
    Top = 252
  end
  object TableDefaut: THValComboBox [12]
    Left = 69
    Top = 8
    Width = 59
    Height = 21
    Style = csDropDownList
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ItemHeight = 13
    ParentFont = False
    TabOrder = 5
    Visible = False
    Items.Strings = (
      'AXE'
      'COMMUN'
      'DECHAMPS'
      'DECOMBOS'
      'DELIENS'
      'DETABLES'
      'DEVUES'
      'FMTIMPDE'
      'FMTIMPOR'
      'FORMES'
      'GRAPHS'
      'LISTE'
      'MENU'
      'PARAMLIB'
      'PARAMSOC'
      'PARART'
      'PARPIECE'
      'PLANREF'
      'PROCEDUR'
      'SCRIPTS'
      'TRADUC'
      '')
    TagDispatch = 0
  end
  object TTDefaut: THValComboBox [13]
    Left = 5
    Top = 8
    Width = 59
    Height = 21
    Style = csDropDownList
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ItemHeight = 13
    ParentFont = False
    TabOrder = 6
    Visible = False
    Items.Strings = (
      'AFTRADUCAFFAIRE'
      'AFTTYPEHEURE'
      'AFTTYPERESSOURCE'
      'GCCATEGORIETAXE'
      'GCLIBFAMILLE'
      'GCZONELIBRE'
      'RTLIBTABLECOMPL'
      'DIMENSIONS'
      'GCCATEGORIEDIM')
    TagDispatch = 0
  end
  inherited Msg: THMsgBox
    Mess.Strings = (
      'Etape'
      '1;Assistant;Voulez-vous quitter l'#39'assistant ?;Q;YN;Y;C;'
      
        '2;?Caption?;Confirmez-vous l'#39'initialisation de la soci'#233't'#233' ?;Q;YC' +
        ';C;'
      '3;?Caption?;Vous devez renseigner toutes les zones;E;O;O;O;;'
      
        '4;?Caption?;L'#39'initialisation de la soci'#233't'#233' a '#233't'#233' effectu'#233'e ! ;A;' +
        'O;O;O;;'
      'Cr'#233'ation de la table :'
      'R'#233'cup'#233'ration des donn'#233'es dans la table :'
      'Mise '#224' jour de la table :'
      
        '8;?Caption?;La base de donn'#233'es source est la m'#234'me que la soci'#233't'#233 +
        ' '#224' initialiser !;E;O;O;O;;'
      'Vous pouvez '#224' pr'#233'sent terminer le param'#233'trage de la soci'#233't'#233'.'
      
        '10;?Caption?;La confirmation du mot de passe est diff'#233'rente du m' +
        'ot de passe !;E;O;O;O;0;'
      'Cr'#233'ation de la vue : '
      'Cr'#233'ation de la tablette : ')
    Left = 41
    Top = 256
  end
  inherited HMTrad: THSystemMenu
    Left = 10
    Top = 255
  end
  object Entete: THTable
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    IndexName = 'DT_CLE1'
    TableName = 'DETABLES'
    dataBaseName = 'DBREF'
    RequestLive = True
    Left = 14
    Top = 160
  end
  object LLigne: THTable
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    IndexName = 'DH_CLE1'
    TableName = 'DECHAMPS'
    dataBaseName = 'DBREF'
    RequestLive = True
    Left = 54
    Top = 160
  end
  object DBREF: TDatabase
    SessionName = 'Default'
    Left = 96
    Top = 160
  end
  object DBSOURCE: TDatabase
    SessionName = 'Default'
    Left = 14
    Top = 208
  end
  object DBDest: TDatabase
    SessionName = 'Default'
    Left = 54
    Top = 208
  end
end

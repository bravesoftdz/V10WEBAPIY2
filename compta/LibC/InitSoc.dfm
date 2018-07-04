object FInitSoc: TFInitSoc
  Left = 324
  Top = 278
  Hint = '*object PBouton: TPanel'
  HelpContext = 3130010
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Initialisation de soci'#233't'#233
  ClientHeight = 323
  ClientWidth = 499
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 499
    Height = 288
    ActivePage = PInfos
    Align = alClient
    TabOrder = 1
    object PInfos: TTabSheet
      HelpContext = 3130020
      Caption = 'D'#233'clarations'
      object Bevel2: TBevel
        Left = 0
        Top = 0
        Width = 491
        Height = 260
        Align = alClient
      end
      object HLabel1: THLabel
        Left = 18
        Top = 12
        Width = 118
        Height = 13
        Caption = '&Base de donn'#233'es source'
        FocusControl = FRef
      end
      object FRef: THValComboBox
        Left = 148
        Top = 8
        Width = 314
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = FRefChange
        TagDispatch = 0
      end
      object GSOCIETE: TGroupBox
        Left = 6
        Top = 36
        Width = 468
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
        object HLabel2: THLabel
          Left = 12
          Top = 24
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
        object HLabel3: THLabel
          Left = 116
          Top = 24
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
          Left = 51
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
          Left = 212
          Top = 19
          Width = 245
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
        Left = 5
        Top = 96
        Width = 468
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
        object HLabel7: THLabel
          Left = 12
          Top = 24
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
        object HLabel8: THLabel
          Left = 116
          Top = 24
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
          Left = 52
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
          Left = 144
          Top = 16
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
          Left = 164
          Top = 12
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
          Left = 212
          Top = 19
          Width = 245
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
        Left = 5
        Top = 156
        Width = 468
        Height = 97
        Hint = 'US_CLE1'
        Caption = ' Utilisateur '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        object HLabel4: THLabel
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
        object HLabel6: THLabel
          Left = 308
          Top = 48
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
          Top = 48
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
        object HLabel5: THLabel
          Left = 116
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
          Left = 308
          Top = 76
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
          Left = 380
          Top = 43
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
          Top = 43
          Width = 161
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
          Left = 236
          Top = 47
          Width = 25
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
          Left = 268
          Top = 47
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
          Left = 212
          Top = 15
          Width = 245
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
          Left = 380
          Top = 72
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
    object Precup: TTabSheet
      HelpContext = 3130030
      Caption = 'Recopie des tables'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 491
        Height = 260
        Align = alClient
      end
      object GroupBox1: TGroupBox
        Left = 315
        Top = 4
        Width = 168
        Height = 194
        Caption = ' Comptabilit'#233' '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        object EXERCICE: TCheckBox
          Left = 12
          Top = 23
          Width = 148
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
          Left = 12
          Top = 47
          Width = 148
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
          Left = 12
          Top = 71
          Width = 148
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
          Left = 12
          Top = 95
          Width = 148
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
          Left = 12
          Top = 119
          Width = 148
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
          Left = 12
          Top = 172
          Width = 148
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
          Left = 12
          Top = 145
          Width = 148
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
      object GroupBox3: TGroupBox
        Left = 8
        Top = 4
        Width = 297
        Height = 137
        Caption = ' Boite '#224' outils '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object PAYS: TCheckBox
          Left = 145
          Top = 16
          Width = 144
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
          TabOrder = 1
        end
        object REGION: TCheckBox
          Left = 145
          Top = 40
          Width = 144
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
          TabOrder = 3
        end
        object CODEPOST: TCheckBox
          Left = 145
          Top = 64
          Width = 144
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
          TabOrder = 5
        end
        object CODEAFB: TCheckBox
          Left = 9
          Top = 64
          Width = 129
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
          TabOrder = 4
        end
        object DEVISE: TCheckBox
          Left = 9
          Top = 40
          Width = 129
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
          TabOrder = 2
        end
        object MODEPAIE: TCheckBox
          Left = 145
          Top = 112
          Width = 144
          Height = 17
          Alignment = taLeftJustify
          Caption = '&Modes de paiement'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 8
          OnClick = MODEPAIEClick
        end
        object MODEREGL: TCheckBox
          Left = 145
          Top = 88
          Width = 144
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Conditions de r'#232'glements'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 7
          OnClick = MODEREGLClick
        end
        object MODELES: TCheckBox
          Left = 9
          Top = 88
          Width = 129
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
          TabOrder = 6
          OnClick = MODELESClick
        end
        object ETABLISS: TCheckBox
          Left = 9
          Top = 16
          Width = 129
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
          Left = 9
          Top = 112
          Width = 129
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
      object GroupBox4: TGroupBox
        Left = 8
        Top = 149
        Width = 297
        Height = 49
        Caption = ' Administrateurs '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object USERGRP: TCheckBox
          Left = 9
          Top = 19
          Width = 131
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Groupes d'#39'utilisateurs'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = USERGRPClick
        end
        object UTILISAT: TCheckBox
          Left = 145
          Top = 19
          Width = 144
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Utilisateurs'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnClick = UTILISATClick
        end
      end
    end
  end
  object PBouton: TPanel
    Left = 0
    Top = 288
    Width = 499
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    Caption = ' '
    TabOrder = 0
    object TT: TLabel
      Left = 10
      Top = 10
      Width = 339
      Height = 15
      AutoSize = False
    end
    object Bdetag: THBitBtn
      Left = 360
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Tout d'#233's'#233'lectionner'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      Visible = False
      OnClick = BdetagClick
      NumGlyphs = 2
      GlobalIndexImage = 'Z0078_S16G2'
    end
    object HelpBtn: THBitBtn
      Left = 456
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Aide'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 4
      OnClick = HelpBtnClick
      Glyph.Data = {
        BE060000424DBE06000000000000360400002800000024000000120000000100
        0800000000008802000000000000000000000001000000010000000000000000
        80000080000000808000800000008000800080800000C0C0C000C0DCC000F0CA
        A600000000000000000000000000000000000000000000000000000000000000
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
        03030606030303030303030303030303030303FFFF0303030303030303030303
        0303030303060404060303030303030303030303030303F8F8FF030303030303
        030303030303030303FE06060403030303030303030303030303F8FF03F8FF03
        0303030303030303030303030303FE060603030303030303030303030303F8FF
        FFF8FF0303030303030303030303030303030303030303030303030303030303
        030303F8F8030303030303030303030303030303030304040603030303030303
        0303030303030303FFFF03030303030303030303030303030306060604030303
        0303030303030303030303F8F8F8FF0303030303030303030303030303FE0606
        0403030303030303030303030303F8FF03F8FF03030303030303030303030303
        03FE06060604030303030303030303030303F8FF03F8FF030303030303030303
        030303030303FE060606040303030303030303030303F8FF0303F8FF03030303
        0303030303030303030303FE060606040303030303030303030303F8FF0303F8
        FF030303030303030303030404030303FE060606040303030303030303FF0303
        F8FF0303F8FF030303030303030306060604030303FE06060403030303030303
        F8F8FF0303F8FF0303F8FF03030303030303FE06060604040406060604030303
        030303F8FF03F8FFFFFFF80303F8FF0303030303030303FE0606060606060606
        06030303030303F8FF0303F8F8F8030303F8FF030303030303030303FEFE0606
        060606060303030303030303F8FFFF030303030303F803030303030303030303
        0303FEFEFEFEFE03030303030303030303F8F8FFFFFFFFFFF803030303030303
        0303030303030303030303030303030303030303030303F8F8F8F8F803030303
        0303}
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object BFerme: THBitBtn
      Left = 424
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      ModalResult = 2
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      NumGlyphs = 2
      GlobalIndexImage = 'Z1770_S16G1'
    end
    object BValider: THBitBtn
      Left = 392
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Valider'
      Default = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ModalResult = 1
      ParentFont = False
      TabOrder = 2
      OnClick = BValiderClick
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
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object BTag: THBitBtn
      Left = 360
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Tout s'#233'lectionner'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BTagClick
      GlobalIndexImage = 'Z0205_S16G1'
    end
  end
  object RempliDefaut: THValComboBox
    Left = 176
    Top = 4
    Width = 69
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
    TabOrder = 2
    Visible = False
    Items.Strings = (
      'AXE'
      'COMMUN'
      'DECHAMPS'
      'DECOMBOS'
      'DELIENS'
      'DETABLES'
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
      'TRADUC')
    TagDispatch = 0
  end
  object DBSOURCE: TDatabase
    SessionName = 'Default'
    Left = 252
    Top = 332
  end
  object DBDest: TDatabase
    SessionName = 'Default'
    Left = 308
    Top = 332
  end
  object Msg: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?Caption?;Confirmez-vous l'#39'initialisation de la soci'#233't'#233' ?;Q;YN' +
        'C;N;C;'
      
        '1;?Caption?;Vous devez renseigner la base de donn'#233'es source !;E;' +
        'O;O;O;;'
      '2;?Caption?;Vous devez renseigner toutes les zones !;E;O;O;O;;'
      
        '3;?Caption?;L'#39'initialisation de la soci'#233't'#233' a '#233't'#233' effectu'#233'e ! ;A;' +
        'O;O;O;;'
      'Cr'#233'ation de la table :'
      'R'#233'cup'#233'ration des donn'#233'es dans la table :'
      'Mise '#224' jour de la table :'
      
        '7;?Caption?;La base de donn'#233'es source est la m'#234'me que la soci'#233't'#233 +
        ' '#224' initialiser !;E;O;O;O;;'
      
        'Vous pouvez '#224' pr'#233'sent utiliser l'#39'assistant de param'#233'trage de soc' +
        'i'#233't'#233'.'
      
        '9;?Caption?;La confirmation du mot de passe est diff'#233'rente du mo' +
        't de passe !;E;O;O;O;0;')
    Left = 440
    Top = 332
  end
  object Entete: THTable
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    IndexName = 'DT_CLE1'
    TableName = 'DETABLES'
    dataBaseName = 'DBREF'
    RequestLive = True
    Left = 16
    Top = 332
  end
  object LLigne: THTable
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    IndexName = 'DH_CLE1'
    TableName = 'DECHAMPS'
    dataBaseName = 'DBREF'
    RequestLive = True
    Left = 56
    Top = 332
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 400
    Top = 332
  end
  object DBREF: TDatabase
    SessionName = 'Default'
    Left = 98
    Top = 332
  end
end

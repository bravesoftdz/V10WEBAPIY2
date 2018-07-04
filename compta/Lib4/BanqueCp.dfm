object FbanqueCP: TFbanqueCP
  Left = 273
  Top = 131
  HelpContext = 7109100
  HorzScrollBar.Visible = False
  VertScrollBar.Visible = False
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Compte bancaire : '
  ClientHeight = 397
  ClientWidth = 431
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  Scaled = False
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object HPB: TPanel
    Left = 0
    Top = 362
    Width = 431
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    Ctl3D = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentCtl3D = False
    ParentFont = False
    ParentShowHint = False
    ShowHint = False
    TabOrder = 1
    object FAutoSave: TCheckBox
      Left = 124
      Top = -20
      Width = 16
      Height = 17
      Caption = 'Auto'
      TabOrder = 11
      Visible = False
    end
    object BAide: THBitBtn
      Left = 396
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
      TabOrder = 10
      OnClick = BAideClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z1117_S16G1'
      IsControl = True
    end
    object BAnnuler: THBitBtn
      Left = 120
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Annuler la derni'#232're action'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 4
      OnClick = BAnnulerClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0075_S16G1'
      IsControl = True
    end
    object BValider: THBitBtn
      Left = 332
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
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 8
      OnClick = BValiderClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0127_S16G1'
      IsControl = True
    end
    object BImprimer: THBitBtn
      Left = 300
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Imprimer'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 7
      OnClick = BImprimerClick
      Margin = 2
      GlobalIndexImage = 'Z0369_S16G1'
    end
    object BFirst: THBitBtn
      Left = 4
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Premier'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BFirstClick
      GlobalIndexImage = 'Z0301_S16G1'
    end
    object BPrev: THBitBtn
      Left = 32
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Pr'#233'c'#233'dent'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BPrevClick
      GlobalIndexImage = 'Z0017_S16G1'
    end
    object BNext: THBitBtn
      Left = 60
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Suivant'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = BNextClick
      GlobalIndexImage = 'Z0031_S16G1'
    end
    object BLast: THBitBtn
      Left = 88
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Dernier'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      OnClick = BLastClick
      GlobalIndexImage = 'Z0264_S16G1'
    end
    object BInsert: THBitBtn
      Left = 184
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Nouveau'
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 6
      Visible = False
      OnClick = BInsertClick
      GlobalIndexImage = 'Z0053_S16G1'
    end
    object BDelete: THBitBtn
      Left = 152
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Supprimer'
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 5
      Visible = False
      OnClick = BDeleteClick
      GlobalIndexImage = 'Z0005_S16G1'
    end
    object BFerme: THBitBtn
      Left = 364
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      ParentShowHint = False
      ShowHint = True
      TabOrder = 9
      OnClick = BFermeClick
      GlobalIndexImage = 'Z1770_S16G1'
    end
  end
  object BPages: TPageControl
    Left = 0
    Top = 0
    Width = 431
    Height = 362
    ActivePage = PIdent
    Align = alClient
    TabOrder = 0
    object PIdent: TTabSheet
      HelpContext = 7109100
      Caption = 'Identification'
      object TBQ_LIBELLE: THLabel
        Left = 168
        Top = 51
        Width = 30
        Height = 13
        Caption = '&Libell'#233
        FocusControl = BQ_LIBELLE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBQ_CODEBIC: THLabel
        Left = 8
        Top = 214
        Width = 45
        Height = 13
        Caption = 'Code B&IC'
        FocusControl = BQ_CODEBIC
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBQ_COMMENTAIRE: THLabel
        Left = 8
        Top = 240
        Width = 61
        Height = 13
        Caption = 'Co&mmentaire'
        FocusControl = BQ_COMMENTAIRE
      end
      object TBQ_CODE: THLabel
        Left = 8
        Top = 51
        Width = 25
        Height = 13
        Caption = '&Code'
        FocusControl = BQ_CODE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBQ_GENERAL: THLabel
        Left = 8
        Top = 75
        Width = 37
        Height = 13
        Caption = '&G'#233'n'#233'ral'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBQ_DEVISE: THLabel
        Left = 168
        Top = 75
        Width = 33
        Height = 13
        Caption = '&Devise'
        FocusControl = BQ_DEVISE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel1: THLabel
        Left = 8
        Top = 188
        Width = 53
        Height = 13
        Caption = 'Code I&BAN'
        FocusControl = BQ_CODEBIC
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HGBRIB: TGroupBox
        Left = 2
        Top = 100
        Width = 420
        Height = 77
        Caption = ' Information RIB  '
        Color = clBtnFace
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 5
        object TBQ_ETABBQ: THLabel
          Left = 10
          Top = 36
          Width = 37
          Height = 13
          Caption = '&Banque'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object TBQ_GUICHET: THLabel
          Left = 85
          Top = 36
          Width = 37
          Height = 13
          Caption = '&Guichet'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object TBQ_NUMEROCOMPTE: THLabel
          Left = 160
          Top = 36
          Width = 37
          Height = 13
          Caption = '&Num'#233'ro'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_CLERIB: THLabel
          Left = 353
          Top = 36
          Width = 15
          Height = 13
          Caption = 'C&l'#233
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_DOMICILIATION: THLabel
          Left = 10
          Top = 13
          Width = 59
          Height = 13
          Caption = '&Domiciliation'
          FocusControl = BQ_DOMICILIATION
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BQ_DOMICILIATION: TDBEdit
          Left = 85
          Top = 12
          Width = 265
          Height = 21
          Ctl3D = True
          DataField = 'BQ_DOMICILIATION'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
        end
        object BQ_ETABBQ: THDBEdit
          Left = 10
          Top = 51
          Width = 72
          Height = 21
          DataField = 'BQ_ETABBQ'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 1
          DisplayFormat = '99999'
          EditMask = '99999'
          Obligatory = False
        end
        object BQ_GUICHET: THDBEdit
          Left = 85
          Top = 51
          Width = 72
          Height = 21
          DataField = 'BQ_GUICHET'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 2
          DisplayFormat = '99999'
          EditMask = '99999'
          Obligatory = False
        end
        object BQ_NUMEROCOMPTE: THDBEdit
          Left = 160
          Top = 51
          Width = 190
          Height = 21
          DataField = 'BQ_NUMEROCOMPTE'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 3
          OnKeyPress = BQ_NUMEROCOMPTEKeyPress
          DisplayFormat = 'aaaaaaaaaaa'
          EditMask = 'aaaaaaaaaaa'
          Obligatory = False
        end
        object BQ_CLERIB: THDBEdit
          Left = 353
          Top = 51
          Width = 45
          Height = 21
          DataField = 'BQ_CLERIB'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 4
          OnExit = BQ_CODEIBANEnter
          DisplayFormat = '99'
          EditMask = '99'
          Obligatory = False
        end
      end
      object BQ_LIBELLE: TDBEdit
        Left = 208
        Top = 46
        Width = 211
        Height = 21
        Ctl3D = True
        DataField = 'BQ_LIBELLE'
        DataSource = SBanqueCP
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
      end
      object BQ_CODEBIC: TDBEdit
        Left = 87
        Top = 210
        Width = 331
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        DataField = 'BQ_CODEBIC'
        DataSource = SBanqueCP
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 7
      end
      object BQ_COMMENTAIRE: TDBEdit
        Left = 87
        Top = 236
        Width = 331
        Height = 21
        DataField = 'BQ_COMMENTAIRE'
        DataSource = SBanqueCP
        TabOrder = 8
      end
      object Pbanque: TPanel
        Left = 0
        Top = 0
        Width = 423
        Height = 41
        Align = alTop
        BevelInner = bvLowered
        Caption = 'Pbanque'
        TabOrder = 0
        object TBQ_BANQUE: THLabel
          Left = 10
          Top = 14
          Width = 109
          Height = 13
          Caption = '&Etablissement bancaire'
          FocusControl = BQ_BANQUE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BMCbanq: TToolbarButton97
          Left = 363
          Top = 10
          Width = 22
          Height = 22
          Hint = 'Param'#232'tres de l'#39#233'tablissement bancaire'
          Layout = blGlyphRight
          ParentShowHint = False
          ShowHint = True
          OnClick = BMCbanqClick
          GlobalIndexImage = 'Z0008_S16G1'
        end
        object BQ_BANQUE: THDBValComboBox
          Left = 132
          Top = 10
          Width = 225
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 0
          TagDispatch = 0
          DataType = 'TTBANQUE'
          DataField = 'BQ_BANQUE'
          DataSource = SBanqueCP
        end
      end
      object GbJfer: TGroupBox
        Left = 2
        Top = 267
        Width = 420
        Height = 65
        Caption = ' Jours de fermeture '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 9
        object Cbj2: TCheckBox
          Left = 117
          Top = 16
          Width = 70
          Height = 17
          Alignment = taLeftJustify
          Caption = 'M&ardi'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnMouseDown = Cbj1MouseDown
        end
        object Cbj1: TCheckBox
          Left = 16
          Top = 16
          Width = 68
          Height = 17
          Alignment = taLeftJustify
          Caption = 'L&undi'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnMouseDown = Cbj1MouseDown
        end
        object Cbj3: TCheckBox
          Left = 217
          Top = 16
          Width = 70
          Height = 17
          Alignment = taLeftJustify
          Caption = 'M&ercredi'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          OnMouseDown = Cbj1MouseDown
        end
        object Cbj4: TCheckBox
          Left = 318
          Top = 16
          Width = 70
          Height = 17
          Alignment = taLeftJustify
          Caption = '&Jeudi'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          OnMouseDown = Cbj1MouseDown
        end
        object Cbj5: TCheckBox
          Left = 16
          Top = 40
          Width = 68
          Height = 17
          Alignment = taLeftJustify
          Caption = '&Vendredi'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 4
          OnMouseDown = Cbj1MouseDown
        end
        object Cbj6: TCheckBox
          Left = 117
          Top = 40
          Width = 70
          Height = 17
          Alignment = taLeftJustify
          Caption = '&Samedi'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 5
          OnMouseDown = Cbj1MouseDown
        end
        object Cbj7: TCheckBox
          Left = 217
          Top = 40
          Width = 70
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Dimanc&he'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 6
          OnMouseDown = Cbj1MouseDown
        end
      end
      object BQ_CODE: TDBEdit
        Left = 48
        Top = 46
        Width = 109
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        DataField = 'BQ_CODE'
        DataSource = SBanqueCP
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
      end
      object BQ_DEVISE: THDBValComboBox
        Left = 208
        Top = 72
        Width = 211
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 4
        OnChange = BQ_DEVISEChange
        TagDispatch = 0
        DataType = 'TTDEVISE'
        DataField = 'BQ_DEVISE'
        DataSource = SBanqueCP
      end
      object BQ_GENERAL: THDBCpteEdit
        Left = 48
        Top = 72
        Width = 109
        Height = 21
        TabOrder = 3
        ZoomTable = tzGbanque
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
        DataField = 'BQ_GENERAL'
        DataSource = SBanqueCP
      end
      object BQ_CODEIBAN: THDBEdit
        Left = 87
        Top = 184
        Width = 331
        Height = 21
        CharCase = ecUpperCase
        DataField = 'BQ_CODEIBAN'
        DataSource = SBanqueCP
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 6
        OnEnter = BQ_CODEIBANEnter
        OnKeyPress = BQ_NUMEROCOMPTEKeyPress
        DisplayFormat = 'aaaa aaaa aaaa aaaa aaaa aaaa aaaa aaaa aaaa'
        EditMask = 'aaaa aaaa aaaa aaaa aaaa aaaa aaaa aaaa aaaa;0'
        Obligatory = False
      end
    end
    object PCoordo: TTabSheet
      HelpContext = 7109120
      Caption = 'Coordonn'#233'es'
      object TBQ_CONTACT: THLabel
        Left = 8
        Top = 41
        Width = 37
        Height = 13
        Caption = '&Contact'
        FocusControl = BQ_CONTACT
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBQ_ADRESSE: THLabel
        Left = 8
        Top = 76
        Width = 47
        Height = 13
        Caption = '&Adresse 1'
        FocusControl = BQ_ADRESSE1
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBQ_CODEPOSTAL: THLabel
        Left = 8
        Top = 158
        Width = 56
        Height = 13
        Caption = 'Code &postal'
        FocusControl = BQ_CODEPOSTAL
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBQ_DIVTERRIT: THLabel
        Left = 8
        Top = 183
        Width = 45
        Height = 13
        Caption = 'Div. &territ.'
        FocusControl = BQ_DIVTERRIT
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBQ_TELEPHONE: THLabel
        Left = 242
        Top = 41
        Width = 18
        Height = 13
        Caption = '&T'#233'l.'
        FocusControl = BQ_TELEPHONE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBQ_FAX: THLabel
        Left = 8
        Top = 239
        Width = 17
        Height = 13
        Caption = '&Fax'
        FocusControl = BQ_FAX
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBQ_VILLE: THLabel
        Left = 168
        Top = 158
        Width = 19
        Height = 13
        Caption = '&Ville'
        FocusControl = BQ_VILLE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBQ_TELEX: THLabel
        Left = 229
        Top = 239
        Width = 29
        Height = 13
        Caption = '&E-Mail'
        FocusControl = BQ_TELEX
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TBQ_ADRESSE2: THLabel
        Left = 8
        Top = 101
        Width = 47
        Height = 13
        Caption = 'A&dresse 2'
        FocusControl = BQ_ADRESSE2
      end
      object TBQ_ADRESSE3: THLabel
        Left = 8
        Top = 126
        Width = 47
        Height = 13
        Caption = 'Ad&resse 3'
        FocusControl = BQ_ADRESSE3
      end
      object TBQ_LANGUE: THLabel
        Left = 168
        Top = 183
        Width = 36
        Height = 13
        Caption = '&Langue'
        FocusControl = BQ_LANGUE
      end
      object TBQ_PAYS: THLabel
        Left = 8
        Top = 211
        Width = 23
        Height = 13
        Caption = '&Pays'
        FocusControl = BQ_PAYS
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object BQ_CONTACT: TDBEdit
        Left = 68
        Top = 37
        Width = 140
        Height = 21
        Ctl3D = True
        DataField = 'BQ_CONTACT'
        DataSource = SBanqueCP
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
      end
      object BQ_ADRESSE1: TDBEdit
        Left = 68
        Top = 72
        Width = 342
        Height = 21
        Ctl3D = True
        DataField = 'BQ_ADRESSE1'
        DataSource = SBanqueCP
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
      end
      object BQ_ADRESSE2: TDBEdit
        Left = 68
        Top = 97
        Width = 342
        Height = 21
        Ctl3D = True
        DataField = 'BQ_ADRESSE2'
        DataSource = SBanqueCP
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 3
      end
      object BQ_ADRESSE3: TDBEdit
        Left = 68
        Top = 122
        Width = 342
        Height = 21
        Ctl3D = True
        DataField = 'BQ_ADRESSE3'
        DataSource = SBanqueCP
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 4
      end
      object BQ_CODEPOSTAL: TDBEdit
        Left = 68
        Top = 154
        Width = 81
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        DataField = 'BQ_CODEPOSTAL'
        DataSource = SBanqueCP
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 5
        OnDblClick = BQ_CODEPOSTALDblClick
      end
      object BQ_DIVTERRIT: TDBEdit
        Left = 68
        Top = 179
        Width = 81
        Height = 21
        Ctl3D = True
        DataField = 'BQ_DIVTERRIT'
        DataSource = SBanqueCP
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 7
        OnDblClick = BQ_DIVTERRITDblClick
      end
      object BQ_TELEPHONE: TDBEdit
        Left = 270
        Top = 37
        Width = 140
        Height = 21
        Ctl3D = True
        DataField = 'BQ_TELEPHONE'
        DataSource = SBanqueCP
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
      end
      object BQ_VILLE: TDBEdit
        Left = 213
        Top = 154
        Width = 197
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        DataField = 'BQ_VILLE'
        DataSource = SBanqueCP
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 6
        OnDblClick = BQ_VILLEDblClick
      end
      object BQ_FAX: TDBEdit
        Left = 68
        Top = 234
        Width = 140
        Height = 21
        Ctl3D = True
        DataField = 'BQ_FAX'
        DataSource = SBanqueCP
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 10
      end
      object BQ_TELEX: TDBEdit
        Left = 270
        Top = 234
        Width = 140
        Height = 21
        Ctl3D = True
        DataField = 'BQ_TELEX'
        DataSource = SBanqueCP
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 11
      end
      object BQ_LANGUE: THDBValComboBox
        Left = 213
        Top = 179
        Width = 197
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 8
        TagDispatch = 0
        DataType = 'TTLANGUE'
        DataField = 'BQ_LANGUE'
        DataSource = SBanqueCP
      end
      object BQ_PAYS: THDBValComboBox
        Left = 68
        Top = 206
        Width = 342
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
        TabOrder = 9
        OnChange = BQ_PAYSChange
        TagDispatch = 0
        DataType = 'TTPAYS'
        DataField = 'BQ_PAYS'
        DataSource = SBanqueCP
      end
    end
    object PCommunic: TTabSheet
      HelpContext = 7109130
      Caption = 'Communications'
      object TBQ_DESTINATAIRE: THLabel
        Left = 6
        Top = 255
        Width = 136
        Height = 13
        Caption = 'Code &destinataire ETEBAC 3'
        FocusControl = BQ_DESTINATAIRE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HGroupBox2: TGroupBox
        Left = 2
        Top = 40
        Width = 418
        Height = 177
        Caption = ' Fichiers '#224' transmettre '
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
        object TBQ_REPVIR: THLabel
          Left = 8
          Top = 25
          Width = 41
          Height = 13
          Caption = '&Virement'
          FocusControl = BQ_REPVIR
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_REPPRELEV: THLabel
          Left = 8
          Top = 61
          Width = 59
          Height = 13
          Caption = '&Pr'#233'l'#232'vement'
          FocusControl = BQ_REPPRELEV
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_REPLCR: THLabel
          Left = 8
          Top = 101
          Width = 55
          Height = 13
          Caption = 'L&CR / BOR'
          FocusControl = BQ_REPLCR
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_REPBONAPAYER: THLabel
          Left = 8
          Top = 142
          Width = 57
          Height = 13
          Caption = 'Bon '#224' &payer'
          FocusControl = BQ_REPBONAPAYER
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BQ_REPVIR: TDBEdit
          Left = 115
          Top = 21
          Width = 286
          Height = 21
          Ctl3D = True
          DataField = 'BQ_REPVIR'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
        end
        object BQ_REPPRELEV: TDBEdit
          Left = 115
          Top = 58
          Width = 286
          Height = 21
          Ctl3D = True
          DataField = 'BQ_REPPRELEV'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 1
        end
        object BQ_REPLCR: TDBEdit
          Left = 115
          Top = 98
          Width = 286
          Height = 21
          Ctl3D = True
          DataField = 'BQ_REPLCR'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 2
        end
        object BQ_REPBONAPAYER: TDBEdit
          Left = 115
          Top = 140
          Width = 286
          Height = 21
          Ctl3D = True
          DataField = 'BQ_REPBONAPAYER'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 3
        end
      end
      object BQ_DESTINATAIRE: TDBEdit
        Left = 237
        Top = 251
        Width = 166
        Height = 21
        Ctl3D = True
        DataField = 'BQ_DESTINATAIRE'
        DataSource = SBanqueCP
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
      end
    end
    object PDelai: TTabSheet
      HelpContext = 7109140
      Caption = 'D'#233'lais'
      object HGroupBox4: TGroupBox
        Left = 0
        Top = 0
        Width = 423
        Height = 46
        Align = alTop
        Caption = ' Virements '
        Color = clBtnFace
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
        object TBQ_DELAIVIRORD: THLabel
          Left = 30
          Top = 19
          Width = 42
          Height = 13
          Caption = '&Ordinaire'
          FocusControl = BQ_DELAIVIRORD
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_DELAIVIRCHAUD: THLabel
          Left = 140
          Top = 19
          Width = 31
          Height = 13
          Caption = '&Chaud'
          FocusControl = BQ_DELAIVIRCHAUD
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_DELAIVIRBRULANT: THLabel
          Left = 265
          Top = 19
          Width = 33
          Height = 13
          Caption = '&Brulant'
          FocusControl = BQ_DELAIVIRBRULANT
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BQ_DELAIVIRORD: THDBSpinEdit
          Left = 85
          Top = 15
          Width = 40
          Height = 22
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxValue = 0
          MinValue = 0
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          Value = 0
          DataField = 'BQ_DELAIVIRORD'
          DataSource = SBanqueCP
        end
        object BQ_DELAIVIRCHAUD: THDBSpinEdit
          Left = 195
          Top = 15
          Width = 40
          Height = 22
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxValue = 0
          MinValue = 0
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 1
          Value = 0
          DataField = 'BQ_DELAIVIRCHAUD'
          DataSource = SBanqueCP
        end
        object BQ_DELAIVIRBRULANT: THDBSpinEdit
          Left = 307
          Top = 15
          Width = 40
          Height = 22
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxValue = 0
          MinValue = 0
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 2
          Value = 0
          DataField = 'BQ_DELAIVIRBRULANT'
          DataSource = SBanqueCP
        end
      end
      object HGroupBox5: TGroupBox
        Left = 0
        Top = 46
        Width = 423
        Height = 46
        Caption = ' Pr'#233'l'#232'vements '
        Color = clBtnFace
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
        object TBQ_DELAIPRELVORD: THLabel
          Left = 30
          Top = 18
          Width = 42
          Height = 13
          Caption = 'O&rdinaire'
          FocusControl = BQ_DELAIPRELVORD
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_DELAIPRELVACC: THLabel
          Left = 140
          Top = 18
          Width = 42
          Height = 13
          Caption = '&Acc'#233'l'#233'r'#233
          FocusControl = BQ_DELAIPRELVACC
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BQ_ECHEREPPRELEV: TDBCheckBox
          Left = 250
          Top = 17
          Width = 159
          Height = 17
          Alignment = taLeftJustify
          Caption = '&Report jour ouvrable suivant'
          Ctl3D = True
          DataField = 'BQ_ECHEREPPRELEV'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 2
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object BQ_DELAIPRELVORD: THDBSpinEdit
          Left = 85
          Top = 14
          Width = 40
          Height = 22
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxValue = 0
          MinValue = 0
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          Value = 0
          DataField = 'BQ_DELAIPRELVORD'
          DataSource = SBanqueCP
        end
        object BQ_DELAIPRELVACC: THDBSpinEdit
          Left = 195
          Top = 14
          Width = 40
          Height = 22
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxValue = 0
          MinValue = 0
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 1
          Value = 0
          DataField = 'BQ_DELAIPRELVACC'
          DataSource = SBanqueCP
        end
      end
      object HGroupBox6: TGroupBox
        Left = 0
        Top = 92
        Width = 423
        Height = 89
        Caption = ' LCR / BOR '
        Color = clBtnFace
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
        object TBQ_DELAILCR: THLabel
          Left = 11
          Top = 16
          Width = 62
          Height = 13
          Caption = 'D'#233'lai remises'
          FocusControl = BQ_DELAILCR
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_NUMEMETLCR: THLabel
          Left = 176
          Top = 16
          Width = 80
          Height = 13
          Caption = 'N'#176' e&metteur LCR'
          FocusControl = BQ_NUMEMETLCR
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_ENCOURSLCR: THLabel
          Left = 176
          Top = 40
          Width = 101
          Height = 13
          Caption = 'E&ncours '#224' l'#39'escompte'
          FocusControl = BQ_ENCOURSLCR
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_PLAFONDLCR: THLabel
          Left = 176
          Top = 64
          Width = 98
          Height = 13
          Caption = 'Plafond '#224' l'#39'escompte'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BQ_ECHEREPLCR: TDBCheckBox
          Left = 11
          Top = 40
          Width = 123
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Ech'#233'ance &report'#233'e'
          Ctl3D = True
          DataField = 'BQ_ECHEREPLCR'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 2
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object BQ_DELAILCR: THDBSpinEdit
          Left = 85
          Top = 12
          Width = 49
          Height = 22
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxValue = 0
          MinValue = 0
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          Value = 0
          DataField = 'BQ_DELAILCR'
          DataSource = SBanqueCP
        end
        object BQ_NUMEMETLCR: TDBEdit
          Left = 279
          Top = 12
          Width = 121
          Height = 21
          Ctl3D = True
          DataField = 'BQ_NUMEMETLCR'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 1
        end
        object BQ_ENCOURSLCR: TDBEdit
          Left = 279
          Top = 36
          Width = 121
          Height = 21
          Ctl3D = True
          DataField = 'BQ_ENCOURSLCR'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 3
        end
        object BQ_PLAFONDLCR: TDBEdit
          Left = 279
          Top = 60
          Width = 121
          Height = 21
          DataField = 'BQ_PLAFONDLCR'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 4
        end
      end
      object GbBap: TGroupBox
        Left = 0
        Top = 181
        Width = 423
        Height = 61
        Caption = ' Bon '#224' payer '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        object TBQ_DELAIBAPLCR: THLabel
          Left = 16
          Top = 19
          Width = 29
          Height = 13
          Caption = 'D'#233'lais'
          FocusControl = BQ_DELAIBAPLCR
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_CONVENTIONLCR: THLabel
          Left = 159
          Top = 19
          Width = 119
          Height = 13
          Caption = 'Type de convention LCR'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BQ_RAPPROAUTOLCR: TDBCheckBox
          Left = 14
          Top = 39
          Width = 125
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Rapprochement au&to'
          DataField = 'BQ_RAPPROAUTOLCR'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object BQ_DELAIBAPLCR: THDBSpinEdit
          Left = 60
          Top = 14
          Width = 79
          Height = 22
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxValue = 0
          MinValue = 0
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          Value = 0
          DataField = 'BQ_DELAIBAPLCR'
          DataSource = SBanqueCP
        end
        object BQ_CONVENTIONLCR: TDBEdit
          Left = 283
          Top = 14
          Width = 118
          Height = 21
          Ctl3D = True
          DataField = 'BQ_CONVENTIONLCR'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 1
        end
      end
      object GbTi: TGroupBox
        Left = 0
        Top = 242
        Width = 423
        Height = 66
        Caption = ' Transferts internationaux '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        object TBQ_DELAITRANSINT: THLabel
          Left = 13
          Top = 20
          Width = 29
          Height = 13
          Caption = 'D'#233'lais'
          FocusControl = BQ_DELAITRANSINT
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_COMPTEFRAIS: THLabel
          Left = 195
          Top = 45
          Width = 85
          Height = 13
          Caption = 'Compte  pour frais'
          FocusControl = BQ_COMPTEFRAIS
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_TYPEREMTRANS: THLabel
          Left = 13
          Top = 45
          Width = 35
          Height = 13
          Caption = 'Remise'
          FocusControl = BQ_TYPEREMTRANS
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_INDREMTRANS: THLabel
          Left = 195
          Top = 20
          Width = 62
          Height = 13
          Caption = 'Indice remise'
          FocusControl = BQ_INDREMTRANS
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BQ_DELAITRANSINT: THDBSpinEdit
          Left = 59
          Top = 16
          Width = 121
          Height = 22
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxValue = 0
          MinValue = 0
          ParentFont = False
          TabOrder = 0
          Value = 0
          DataField = 'BQ_DELAITRANSINT'
          DataSource = SBanqueCP
        end
        object BQ_COMPTEFRAIS: THDBCpteEdit
          Left = 288
          Top = 41
          Width = 121
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          ZoomTable = tzGcharge
          Vide = False
          Bourre = False
          okLocate = False
          SynJoker = False
          DataField = 'BQ_COMPTEFRAIS'
          DataSource = SBanqueCP
        end
        object BQ_TYPEREMTRANS: THDBValComboBox
          Left = 59
          Top = 41
          Width = 121
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 2
          TagDispatch = 0
          DataField = 'BQ_TYPEREMTRANS'
          DataSource = SBanqueCP
        end
        object BQ_INDREMTRANS: THDBValComboBox
          Left = 288
          Top = 16
          Width = 121
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 1
          TagDispatch = 0
          DataField = 'BQ_INDREMTRANS'
          DataSource = SBanqueCP
        end
      end
    end
    object PInfo: TTabSheet
      HelpContext = 7109150
      Caption = 'Informations'
      object GbSold: TGroupBox
        Left = 0
        Top = 0
        Width = 423
        Height = 62
        Align = alTop
        Caption = ' Solde '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object TBQ_DATEDERNSOLDE: THLabel
          Left = 14
          Top = 16
          Width = 42
          Height = 13
          Caption = 'Solde au'
          FocusControl = BQ_DATEDERNSOLDE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_DERNSOLDEFRS: THLabel
          Left = 194
          Top = 16
          Width = 62
          Height = 13
          Caption = 'Dernier solde'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_DERNSOLDEDEV: THLabel
          Left = 194
          Top = 39
          Width = 62
          Height = 13
          Caption = 'Dernier solde'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TEdev: TLabel
          Left = 14
          Top = 39
          Width = 33
          Height = 13
          Caption = 'Devise'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BQ_DATEDERNSOLDE: TDBEdit
          Left = 68
          Top = 12
          Width = 117
          Height = 21
          Color = clBtnFace
          Ctl3D = True
          DataField = 'BQ_DATEDERNSOLDE'
          DataSource = SBanqueCP
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
        end
        object BQ_DERNSOLDEFRS: THNumEdit
          Left = 260
          Top = 12
          Width = 145
          Height = 21
          Color = clBtnFace
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object BQ_DERNSOLDEDEV: THNumEdit
          Left = 260
          Top = 36
          Width = 145
          Height = 21
          Color = clBtnFace
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0'
          Debit = False
          UseRounding = True
          Validate = False
        end
        object EDev: TEdit
          Left = 68
          Top = 36
          Width = 117
          Height = 21
          Color = clBtnFace
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
        end
      end
      object GbModele: TGroupBox
        Left = 0
        Top = 66
        Width = 423
        Height = 63
        Caption = ' Mod'#232'le de lettre ou de document'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        object TBQ_LETTREVIR: THLabel
          Left = 227
          Top = 17
          Width = 41
          Height = 13
          Caption = '&Virement'
          FocusControl = BQ_LETTREVIR
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_LETTREPRELV: THLabel
          Left = 9
          Top = 17
          Width = 59
          Height = 13
          Caption = '&Pr'#233'levement'
          FocusControl = BQ_LETTREPRELV
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_LETTRELCR: THLabel
          Left = 9
          Top = 41
          Width = 21
          Height = 13
          Caption = '&LCR'
          FocusControl = BQ_LETTRELCR
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_LETTRECHQ: THLabel
          Left = 227
          Top = 41
          Width = 37
          Height = 13
          Caption = '&Ch'#232'que'
          FocusControl = BQ_LETTRECHQ
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BQ_LETTREVIR: THDBValComboBox
          Left = 272
          Top = 13
          Width = 145
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 1
          TagDispatch = 0
          DataType = 'TTMODELEBOR'
          DataField = 'BQ_LETTREVIR'
          DataSource = SBanqueCP
        end
        object BQ_LETTREPRELV: THDBValComboBox
          Left = 74
          Top = 13
          Width = 145
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 0
          TagDispatch = 0
          DataType = 'TTMODELEBOR'
          DataField = 'BQ_LETTREPRELV'
          DataSource = SBanqueCP
        end
        object BQ_LETTRELCR: THDBValComboBox
          Left = 74
          Top = 37
          Width = 145
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 2
          TagDispatch = 0
          DataType = 'TTMODELEBOR'
          DataField = 'BQ_LETTRELCR'
          DataSource = SBanqueCP
        end
        object BQ_LETTRECHQ: THDBValComboBox
          Left = 272
          Top = 37
          Width = 145
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 3
          TagDispatch = 0
          DataType = 'TTMODELELETTRECHQ'
          DataField = 'BQ_LETTRECHQ'
          DataSource = SBanqueCP
        end
      end
      object GbDiv: TGroupBox
        Left = 0
        Top = 136
        Width = 423
        Height = 61
        Caption = ' Divers '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        object TBQ_NUMEMETVIR: THLabel
          Left = 9
          Top = 15
          Width = 124
          Height = 13
          Caption = 'N'#176' '#233'metteur des virements'
          FocusControl = BQ_NUMEMETVIR
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TBQ_GUIDECOMPATBLE: THLabel
          Left = 229
          Top = 15
          Width = 66
          Height = 13
          Caption = '&Guide compta'
          FocusControl = BQ_GUIDECOMPATBLE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BQ_RELEVEETRANGER: TDBCheckBox
          Left = 9
          Top = 34
          Width = 209
          Height = 17
          Alignment = taLeftJustify
          Caption = '&Relev'#233' '#233'tranger'
          Ctl3D = True
          DataField = 'BQ_RELEVEETRANGER'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 2
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object BQ_GUIDECOMPATBLE: TDBEdit
          Left = 300
          Top = 11
          Width = 93
          Height = 21
          DataField = 'BQ_GUIDECOMPATBLE'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
        object BQ_RAPPAUTOREL: TDBCheckBox
          Left = 229
          Top = 34
          Width = 164
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Rappro&chement auto relev'#233's'
          DataField = 'BQ_RAPPAUTOREL'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object BQ_NUMEMETVIR: TDBEdit
          Left = 154
          Top = 11
          Width = 64
          Height = 21
          Ctl3D = True
          DataField = 'BQ_NUMEMETVIR'
          DataSource = SBanqueCP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
        end
      end
      object TBQ_BLOCNOTE: TGroupBox
        Left = 0
        Top = 204
        Width = 423
        Height = 101
        Caption = ' BlocNote '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object BQ_BLOCNOTE: THDBRichEditOLE
          Left = 9
          Top = 16
          Width = 408
          Height = 77
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          MaxLength = 262140
          ParentCtl3D = False
          ParentFont = False
          ScrollBars = ssVertical
          TabOrder = 0
          Margins.Top = 0
          Margins.Bottom = 0
          Margins.Left = 0
          Margins.Right = 0
          ContainerName = 'Document'
          ObjectMenuPrefix = '&Object'
          LinesRTF.Strings = (
            
              '{\rtf1\ansi\ansicpg1252\deff0\deflang1036{\fonttbl{\f0\fswiss Ar' +
              'ial;}}'
            '{\colortbl ;\red0\green0\blue0;}'
            
              '{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\pard\cf1\f0\' +
              'fs16 BQ_BLOCNOTE'
            '\par }')
          DataField = 'BQ_BLOCNOTE'
          DataSource = SBanqueCP
        end
      end
    end
  end
  object DBNav: TDBNavigator
    Left = 178
    Top = 55
    Width = 80
    Height = 18
    DataSource = SBanqueCP
    VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast]
    ConfirmDelete = False
    TabOrder = 2
    Visible = False
  end
  object SBanqueCP: TDataSource
    DataSet = TBanqueCP
    Left = 366
    Top = 321
  end
  object TBanqueCP: THTable
    MarshalOptions = moMarshalModifiedOnly
    AfterPost = TBanqueCPAfterPost
    AfterDelete = TBanqueCPAfterDelete
    EnableBCD = False
    IndexName = 'BQ_CLE2'
    TableName = 'BANQUECP'
    dataBaseName = 'SOC'
    UpdateMode = upWhereChanged
    RequestLive = True
    Left = 232
    Top = 343
    object TBanqueCPBQ_CODE: TStringField
      FieldName = 'BQ_CODE'
      Size = 17
    end
    object TBanqueCPBQ_GENERAL: TStringField
      FieldName = 'BQ_GENERAL'
      Size = 17
    end
    object TBanqueCPBQ_LIBELLE: TStringField
      FieldName = 'BQ_LIBELLE'
      Size = 35
    end
    object TBanqueCPBQ_DOMICILIATION: TStringField
      FieldName = 'BQ_DOMICILIATION'
      Size = 24
    end
    object TBanqueCPBQ_ADRESSE1: TStringField
      FieldName = 'BQ_ADRESSE1'
      Size = 35
    end
    object TBanqueCPBQ_ADRESSE2: TStringField
      FieldName = 'BQ_ADRESSE2'
      Size = 35
    end
    object TBanqueCPBQ_ADRESSE3: TStringField
      FieldName = 'BQ_ADRESSE3'
      Size = 35
    end
    object TBanqueCPBQ_CODEPOSTAL: TStringField
      FieldName = 'BQ_CODEPOSTAL'
      Size = 9
    end
    object TBanqueCPBQ_VILLE: TStringField
      FieldName = 'BQ_VILLE'
      Size = 35
    end
    object TBanqueCPBQ_DIVTERRIT: TStringField
      FieldName = 'BQ_DIVTERRIT'
      Size = 9
    end
    object TBanqueCPBQ_PAYS: TStringField
      FieldName = 'BQ_PAYS'
      FixedChar = True
      Size = 3
    end
    object TBanqueCPBQ_LANGUE: TStringField
      FieldName = 'BQ_LANGUE'
      FixedChar = True
      Size = 3
    end
    object TBanqueCPBQ_DEVISE: TStringField
      FieldName = 'BQ_DEVISE'
      FixedChar = True
      Size = 3
    end
    object TBanqueCPBQ_TELEPHONE: TStringField
      FieldName = 'BQ_TELEPHONE'
      Size = 25
    end
    object TBanqueCPBQ_FAX: TStringField
      FieldName = 'BQ_FAX'
      Size = 25
    end
    object TBanqueCPBQ_TELEX: TStringField
      FieldName = 'BQ_TELEX'
      Size = 25
    end
    object TBanqueCPBQ_CONTACT: TStringField
      FieldName = 'BQ_CONTACT'
      Size = 35
    end
    object TBanqueCPBQ_ETABBQ: TStringField
      FieldName = 'BQ_ETABBQ'
      Size = 5
    end
    object TBanqueCPBQ_NUMEROCOMPTE: TStringField
      FieldName = 'BQ_NUMEROCOMPTE'
      Size = 11
    end
    object TBanqueCPBQ_CLERIB: TStringField
      FieldName = 'BQ_CLERIB'
      Size = 2
    end
    object TBanqueCPBQ_GUICHET: TStringField
      FieldName = 'BQ_GUICHET'
      Size = 5
    end
    object TBanqueCPBQ_CODEBIC: TStringField
      FieldName = 'BQ_CODEBIC'
      Size = 35
    end
    object TBanqueCPBQ_NUMEMETLCR: TStringField
      FieldName = 'BQ_NUMEMETLCR'
      Size = 6
    end
    object TBanqueCPBQ_CONVENTIONLCR: TStringField
      FieldName = 'BQ_CONVENTIONLCR'
      Size = 6
    end
    object TBanqueCPBQ_NUMEMETVIR: TStringField
      FieldName = 'BQ_NUMEMETVIR'
      Size = 6
    end
    object TBanqueCPBQ_JOURFERMETUE: TStringField
      FieldName = 'BQ_JOURFERMETUE'
      Size = 7
    end
    object TBanqueCPBQ_REPRELEVE: TStringField
      FieldName = 'BQ_REPRELEVE'
      Size = 40
    end
    object TBanqueCPBQ_REPLCR: TStringField
      FieldName = 'BQ_REPLCR'
      Size = 40
    end
    object TBanqueCPBQ_REPLCRFOURN: TStringField
      FieldName = 'BQ_REPLCRFOURN'
      Size = 40
    end
    object TBanqueCPBQ_REPVIR: TStringField
      FieldName = 'BQ_REPVIR'
      Size = 40
    end
    object TBanqueCPBQ_REPPRELEV: TStringField
      FieldName = 'BQ_REPPRELEV'
      Size = 40
    end
    object TBanqueCPBQ_REPBONAPAYER: TStringField
      FieldName = 'BQ_REPBONAPAYER'
      Size = 40
    end
    object TBanqueCPBQ_REPIMPAYELCR: TStringField
      FieldName = 'BQ_REPIMPAYELCR'
      Size = 40
    end
    object TBanqueCPBQ_DELAIVIRORD: TIntegerField
      FieldName = 'BQ_DELAIVIRORD'
    end
    object TBanqueCPBQ_DELAIVIRCHAUD: TIntegerField
      FieldName = 'BQ_DELAIVIRCHAUD'
    end
    object TBanqueCPBQ_DELAIVIRBRULANT: TIntegerField
      FieldName = 'BQ_DELAIVIRBRULANT'
    end
    object TBanqueCPBQ_DELAIPRELVORD: TIntegerField
      FieldName = 'BQ_DELAIPRELVORD'
    end
    object TBanqueCPBQ_DELAIPRELVACC: TIntegerField
      FieldName = 'BQ_DELAIPRELVACC'
    end
    object TBanqueCPBQ_DELAILCR: TIntegerField
      FieldName = 'BQ_DELAILCR'
    end
    object TBanqueCPBQ_GUIDECOMPATBLE: TStringField
      FieldName = 'BQ_GUIDECOMPATBLE'
      Size = 17
    end
    object TBanqueCPBQ_DERNSOLDEFRS: TFloatField
      FieldName = 'BQ_DERNSOLDEFRS'
    end
    object TBanqueCPBQ_DERNSOLDEDEV: TFloatField
      FieldName = 'BQ_DERNSOLDEDEV'
    end
    object TBanqueCPBQ_DATEDERNSOLDE: TDateTimeField
      FieldName = 'BQ_DATEDERNSOLDE'
    end
    object TBanqueCPBQ_RELEVEETRANGER: TStringField
      FieldName = 'BQ_RELEVEETRANGER'
      FixedChar = True
      Size = 1
    end
    object TBanqueCPBQ_CALENDRIER: TStringField
      FieldName = 'BQ_CALENDRIER'
      FixedChar = True
      Size = 3
    end
    object TBanqueCPBQ_RAPPAUTOREL: TStringField
      FieldName = 'BQ_RAPPAUTOREL'
      FixedChar = True
      Size = 1
    end
    object TBanqueCPBQ_RAPPROAUTOLCR: TStringField
      FieldName = 'BQ_RAPPROAUTOLCR'
      FixedChar = True
      Size = 1
    end
    object TBanqueCPBQ_LETTREVIR: TStringField
      FieldName = 'BQ_LETTREVIR'
      FixedChar = True
      Size = 3
    end
    object TBanqueCPBQ_LETTREPRELV: TStringField
      FieldName = 'BQ_LETTREPRELV'
      FixedChar = True
      Size = 3
    end
    object TBanqueCPBQ_LETTRELCR: TStringField
      FieldName = 'BQ_LETTRELCR'
      FixedChar = True
      Size = 3
    end
    object TBanqueCPBQ_ENCOURSLCR: TFloatField
      FieldName = 'BQ_ENCOURSLCR'
    end
    object TBanqueCPBQ_PLAFONDLCR: TFloatField
      FieldName = 'BQ_PLAFONDLCR'
    end
    object TBanqueCPBQ_REPIMPAYEPRELV: TStringField
      FieldName = 'BQ_REPIMPAYEPRELV'
      Size = 40
    end
    object TBanqueCPBQ_ECHEREPPRELEV: TStringField
      FieldName = 'BQ_ECHEREPPRELEV'
      FixedChar = True
      Size = 1
    end
    object TBanqueCPBQ_ECHEREPLCR: TStringField
      FieldName = 'BQ_ECHEREPLCR'
      FixedChar = True
      Size = 1
    end
    object TBanqueCPBQ_SOCIETE: TStringField
      FieldName = 'BQ_SOCIETE'
      FixedChar = True
      Size = 3
    end
    object TBanqueCPBQ_BANQUE: TStringField
      FieldName = 'BQ_BANQUE'
      FixedChar = True
      Size = 3
    end
    object TBanqueCPBQ_BLOCNOTE: TMemoField
      FieldName = 'BQ_BLOCNOTE'
      BlobType = ftMemo
      Size = 1
    end
    object TBanqueCPBQ_DELAIBAPLCR: TIntegerField
      FieldName = 'BQ_DELAIBAPLCR'
    end
    object TBanqueCPBQ_DELAITRANSINT: TIntegerField
      FieldName = 'BQ_DELAITRANSINT'
    end
    object TBanqueCPBQ_COMPTEFRAIS: TStringField
      FieldName = 'BQ_COMPTEFRAIS'
      Size = 17
    end
    object TBanqueCPBQ_TYPEREMTRANS: TStringField
      FieldName = 'BQ_TYPEREMTRANS'
      FixedChar = True
      Size = 3
    end
    object TBanqueCPBQ_INDREMTRANS: TStringField
      FieldName = 'BQ_INDREMTRANS'
      FixedChar = True
      Size = 3
    end
    object TBanqueCPBQ_COMMENTAIRE: TStringField
      FieldName = 'BQ_COMMENTAIRE'
      Size = 35
    end
    object TBanqueCPBQ_LETTRECHQ: TStringField
      FieldName = 'BQ_LETTRECHQ'
      FixedChar = True
      Size = 3
    end
    object TBanqueCPBQ_NUMEMETPRE: TStringField
      FieldName = 'BQ_NUMEMETPRE'
      Size = 6
    end
    object TBanqueCPBQ_DESTINATAIRE: TStringField
      FieldName = 'BQ_DESTINATAIRE'
      Size = 5
    end
    object TBanqueCPBQ_MULTIDEVISE: TStringField
      FieldName = 'BQ_MULTIDEVISE'
      FixedChar = True
      Size = 1
    end
    object TBanqueCPBQ_CODEIBAN: TStringField
      FieldName = 'BQ_CODEIBAN'
      Size = 70
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Comptes bancaires;Voulez-vous enregistrer les modifications?;Q' +
        ';YNC;Y;C;'
      
        '1;Comptes bancaires;Confirmez-vous la suppression de l'#39'enregistr' +
        'ement?;Q;YNC;N;C;'
      '2;Comptes bancaires;Vous devez renseigner un code !;E;O;O;O;'
      '3;Comptes bancaires;Vous devez renseigner un libell'#233' !;E;O;O;O;'
      
        '4;Comptes bancaires;Vous devez renseigner le code banque du RIB;' +
        'E;O;O;O;'
      
        '5;Comptes bancaires;Vous devez renseigner le code guichet du RIB' +
        ';E;O;O;O;'
      
        '6;Comptes bancaires;Vous devez renseigner le num'#233'ro de compte du' +
        ' RIB;E;O;O;O;'
      '7;Comptes bancaires;Vous devez renseigner la cl'#233' du RIB;E;O;O;O;'
      
        '8;Comptes bancaires;Vous devez renseigner le RIB ou le code BIC ' +
        'du compte;E;O;O;O;'
      'L'#39'enregistrement est inaccessible'
      
        '10;Comptes bancaires;Le code que vous avez saisi existe d'#233'ja. Vo' +
        'us devez le modifier;E;O;O;O;'
      
        '11;Comptes bancaires;La cl'#233' RIB est erronn'#233'e. Souhaitez-vous la ' +
        'recalculer ?;Q;YN;Y;C;'
      
        '12;Comptes bancaires;Vous ne pouvez pas affecter cette devise su' +
        'r le compte;E;O;O;O;'
      
        '13;Comptes bancaires;La domiciliation comprte des caract'#232'res int' +
        'erdits !;E;O;O;O;')
    Left = 155
    Top = 345
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 302
    Top = 320
  end
end

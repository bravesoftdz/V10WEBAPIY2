inherited FassistParamSoc: TFassistParamSoc
  Left = 354
  Top = 260
  AlphaBlend = True
  BorderStyle = bsDialog
  Caption = 'Assistant de param'#233'trage soci'#233't'#233
  ClientHeight = 595
  ClientWidth = 791
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Left = 10
    Top = 572
    Width = 90
  end
  inherited lAide: THLabel
    Left = 250
    Top = 542
    Width = 414
  end
  inherited bPrecedent: TToolbarButton97
    Left = 549
    Top = 567
  end
  inherited bSuivant: TToolbarButton97
    Left = 629
    Top = 567
  end
  inherited bFin: TToolbarButton97
    Left = 709
    Top = 567
  end
  inherited bAnnuler: TToolbarButton97
    Left = 461
    Top = 567
  end
  inherited bAide: TToolbarButton97
    Left = 293
    Top = 567
  end
  object BRetMenu: TToolbarButton97 [7]
    Left = 205
    Top = 567
    Width = 75
    Height = 23
    Anchors = [akRight, akBottom]
    Caption = 'Retour Menu'
    Enabled = False
    Flat = False
    OnClick = BRetMenuClick
  end
  object Benregistre: TToolbarButton97 [8]
    Left = 709
    Top = 567
    Width = 75
    Height = 23
    Anchors = [akRight, akBottom]
    Caption = '&Enregistrer'
    Flat = False
    OnClick = bFinClick
  end
  inherited Plan: THPanel
    Left = 258
    Top = 41
    Width = 519
    Height = 416
    Anchors = [akTop, akRight]
  end
  inherited GroupBox1: THGroupBox
    Left = 0
    Top = 554
    Width = 794
  end
  inherited P: THPageControl2
    Left = 254
    Top = 41
    Width = 528
    Height = 496
    ActivePage = ParamSoc2
    Anchors = [akTop, akRight]
    OwnerDraw = True
    object MenuOption: TTabSheet
      Caption = 'Menu des Options'
      ImageIndex = 2
      object PREFERBTP: TToolbarButton97
        Left = 186
        Top = 234
        Width = 150
        Height = 25
        DropdownAlways = True
        DropdownArrow = True
        DropdownMenu = POPBTP
        Caption = 'Pr'#233'f'#233'rences de gestion'
        ImageIndex = 0
        Opaque = False
      end
      object ParamEchangeCpta: TToolbarButton97
        Left = 186
        Top = 204
        Width = 150
        Height = 25
        Caption = 'Echanges comptables'
        ImageIndex = 0
        Opaque = False
        OnClick = ParamEchangeCptaClick
      end
      object BPrefcompta: TToolbarButton97
        Left = 186
        Top = 172
        Width = 150
        Height = 25
        Caption = 'Pr'#233'f'#233'rences Comptables'
        ImageIndex = 0
        Opaque = False
        OnClick = BPrefcomptaClick
      end
      object Label4: TLabel
        Left = 16
        Top = 1
        Width = 449
        Height = 27
        Caption = 'Bienvenue dans l'#39'Assistant de Parametrage'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -23
        Font.Name = 'Cambria'
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = True
      end
      object Label5: TLabel
        Left = 7
        Top = 49
        Width = 467
        Height = 40
        Alignment = taCenter
        Caption = 
          'Cet assistant va vous permettre de param'#233'trer tout votre dossier' +
          ' en toute simplicit'#233' en suivant chacune des '#233'tapes...'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -17
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        Transparent = True
        WordWrap = True
      end
      object Label6: TLabel
        Left = 11
        Top = 311
        Width = 453
        Height = 60
        Alignment = taCenter
        Caption = 
          'Le bouton "Suivant" et "Pr'#233'c'#233'dent" permettent de passer d'#39'une pa' +
          'ge '#224' l'#39'autre de l'#39'assistant. Le bouton "Enregistrer" permet de v' +
          'alider le param'#233'trage effectu'#233
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -17
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        Transparent = True
        WordWrap = True
      end
      object BCoordonnees: TToolbarButton97
        Left = 186
        Top = 110
        Width = 150
        Height = 25
        Caption = 'Pr'#233'f'#233'rences Entreprise'
        ImageIndex = 0
        Opaque = False
        OnClick = bCoordonneesClick
      end
      object BoptionEtats: TToolbarButton97
        Left = 186
        Top = 263
        Width = 150
        Height = 25
        Caption = 'Option Pr'#233'sentation Etats'
        ImageIndex = 0
        Opaque = False
        OnClick = BoptionEtatsClick
      end
      object BExercices: TToolbarButton97
        Left = 186
        Top = 140
        Width = 150
        Height = 25
        Caption = 'Exercices Comptables'
        ImageIndex = 0
        Opaque = False
        OnClick = BExercicesClick
      end
      object GroupBox4: TGroupBox
        Left = 4
        Top = 32
        Width = 486
        Height = 7
        TabOrder = 0
      end
    end
    object ParamSoc1: TTabSheet
      Caption = 'Coordonn'#233'es Entreprises'
      ImageIndex = 3
      object Label1: TLabel
        Left = 6
        Top = 51
        Width = 73
        Height = 13
        Caption = 'Code Soci'#233't'#233' : '
      end
      object LTitre1: THLabel
        Left = 3
        Top = 5
        Width = 556
        Height = 23
        Alignment = taCenter
        AutoSize = False
        Caption = 'Coordonn'#233'es Entreprises'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = 20
        Font.Name = 'Cambria'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label3: TLabel
        Left = 242
        Top = 51
        Width = 83
        Height = 13
        Caption = 'Forme Juridique : '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object Line1: TGroupBox
        Left = 3
        Top = 29
        Width = 498
        Height = 7
        TabOrder = 1
      end
      object SO_SOCIETE: TEdit
        Left = 80
        Top = 47
        Width = 49
        Height = 21
        Enabled = False
        TabOrder = 0
      end
      object GrpAdresse: TGroupBox
        Left = 0
        Top = 70
        Width = 502
        Height = 189
        Caption = 'Adresse'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        object Label2: TLabel
          Left = 338
          Top = 164
          Width = 52
          Height = 13
          Caption = 'Div. Territ :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label7: TLabel
          Left = 6
          Top = 164
          Width = 29
          Height = 13
          Caption = 'Pays :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object SO_ADRESSE1: TEdit
          Left = 6
          Top = 51
          Width = 475
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
        object SO_ADRESSE2: TEdit
          Left = 6
          Top = 78
          Width = 475
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
        end
        object SO_ADRESSE3: TEdit
          Left = 6
          Top = 105
          Width = 475
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
        end
        object SO_CODEPOSTAL: TEdit
          Left = 6
          Top = 132
          Width = 65
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 4
        end
        object SO_VILLE: TEdit
          Left = 90
          Top = 131
          Width = 391
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 5
        end
        object SO_DIVTERRIT: TEdit
          Left = 394
          Top = 160
          Width = 87
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 6
        end
        object SO_LIBELLE: TEdit
          Left = 6
          Top = 25
          Width = 475
          Height = 21
          Hint = 'Raison Sociale'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object SO_PAYS: THValComboBox
          Left = 40
          Top = 158
          Width = 233
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 7
          TagDispatch = 0
          DataType = 'TTPAYS'
          DataTypeParametrable = True
        end
      end
      object SO_NATUREJURIDIQUE: THValComboBox
        Left = 328
        Top = 49
        Width = 139
        Height = 21
        ItemHeight = 13
        TabOrder = 3
        TagDispatch = 0
        DataType = 'TTFORMEJURIDIQUE'
        DataTypeParametrable = True
      end
      object GrpCoordonnees: TGroupBox
        Left = 0
        Top = 259
        Width = 502
        Height = 113
        Caption = 'Coordonn'#233'es'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        object Label10: TLabel
          Left = 10
          Top = 29
          Width = 21
          Height = 13
          Caption = 'T'#233'l :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label11: TLabel
          Left = 282
          Top = 29
          Width = 23
          Height = 13
          Caption = 'Fax :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label12: TLabel
          Left = 10
          Top = 57
          Width = 26
          Height = 13
          Caption = 'Telex'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label13: TLabel
          Left = 10
          Top = 86
          Width = 50
          Height = 13
          Caption = 'Adr. Mail : '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LSO_CONTACT: TLabel
          Left = 218
          Top = 57
          Width = 46
          Height = 13
          Caption = 'Contact : '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object SO_TELEPHONE: TEdit
          Left = 39
          Top = 25
          Width = 171
          Height = 21
          Hint = 'Raison Sociale'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object SO_FAX: TEdit
          Left = 312
          Top = 25
          Width = 169
          Height = 21
          Hint = 'Raison Sociale'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
        object SO_TELEX: TEdit
          Left = 39
          Top = 53
          Width = 171
          Height = 21
          Hint = 'Raison Sociale'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
        end
        object SO_MAIL: TEdit
          Left = 63
          Top = 82
          Width = 418
          Height = 21
          Hint = 'Raison Sociale'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
        end
        object SO_CONTACT: TEdit
          Left = 264
          Top = 53
          Width = 217
          Height = 21
          Hint = 'Raison Sociale'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 4
        end
      end
    end
    object ParamSoc2: TTabSheet
      Caption = 'Renseignements Administratifs'
      ImageIndex = 4
      object HLabel5: THLabel
        Left = 3
        Top = 5
        Width = 486
        Height = 23
        Alignment = taCenter
        AutoSize = False
        Caption = 'Renseignements Administratifs'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = 20
        Font.Name = 'Cambria'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object LTXTJURIDIQUE: TLabel
        Left = 6
        Top = 51
        Width = 81
        Height = 13
        Caption = 'Texte Juridique : '
      end
      object Label8: TLabel
        Left = 8
        Top = 83
        Width = 54
        Height = 13
        Caption = 'Code NIF : '
      end
      object LRVA: TLabel
        Left = 302
        Top = 83
        Width = 56
        Height = 13
        Caption = 'Code RVA :'
      end
      object LAPE: TLabel
        Left = 302
        Top = 106
        Width = 58
        Height = 13
        Caption = 'Code APE : '
      end
      object LSIRET: TLabel
        Left = 8
        Top = 106
        Width = 59
        Height = 13
        Caption = 'N'#176' SIRET :  '
      end
      object LRCS: TLabel
        Left = 304
        Top = 130
        Width = 59
        Height = 13
        Caption = 'Code RCS : '
      end
      object LCAPITAL: TLabel
        Left = 8
        Top = 130
        Width = 38
        Height = 13
        Caption = 'Capital :'
      end
      object LLOGO: TLabel
        Left = 8
        Top = 162
        Width = 30
        Height = 13
        Caption = 'Logo :'
      end
      object TLabel
        Left = 8
        Top = 181
        Width = 362
        Height = 15
        BiDiMode = bdLeftToRight
        Caption = 
          'Ce logo sera imprim'#233' en ent'#234'te des documents (Devis, Facture, ..' +
          '.)'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -13
        Font.Name = 'Cambria'
        Font.Style = []
        ParentBiDiMode = False
        ParentFont = False
      end
      object GroupBox7: TGroupBox
        Left = 3
        Top = 29
        Width = 486
        Height = 7
        TabOrder = 0
      end
      object SO_TXTJURIDIQUE: TEdit
        Left = 88
        Top = 49
        Width = 401
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
      end
      object SO_NIF: TEdit
        Left = 64
        Top = 79
        Width = 129
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
      end
      object SO_RVA: TEdit
        Left = 360
        Top = 79
        Width = 129
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
      end
      object SO_APE: TEdit
        Left = 360
        Top = 102
        Width = 129
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
      end
      object SO_SIRET: TEdit
        Left = 64
        Top = 102
        Width = 129
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 5
      end
      object SO_RC: TEdit
        Left = 360
        Top = 126
        Width = 129
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 6
      end
      object SO_CAPITAL: TEdit
        Left = 64
        Top = 126
        Width = 145
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 7
        OnExit = SO_CAPITALExit
      end
      object SO_GCLOGO_ETATS: THCritMaskEdit
        Left = 64
        Top = 158
        Width = 425
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 8
        TagDispatch = 0
        DataType = 'OPENBMPFILE(*.JPG)'
        ElipsisButton = True
      end
      object GroupBox9: TGroupBox
        Left = 4
        Top = 200
        Width = 487
        Height = 81
        Caption = 'Gestion des M'#233'tr'#233's'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        TabOrder = 9
        object Label9: TLabel
          Left = 10
          Top = 50
          Width = 29
          Height = 13
          Caption = 'R'#233'p : '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object SO_BTMETREDOC: TCheckBox
          Left = 8
          Top = 22
          Width = 113
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Gestion des M'#233'tr'#233's'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object SO_BTREPMETR: THCritMaskEdit
          Left = 46
          Top = 47
          Width = 321
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          TagDispatch = 0
          DataType = 'DIRECTORY'
          ElipsisButton = True
        end
        object SO_METRESEXCEL: TCheckBox
          Left = 160
          Top = 22
          Width = 105
          Height = 17
          Alignment = taLeftJustify
          Caption = 'M'#233'tr'#233's via Excel'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          OnClick = SO_METRESEXCELClick
        end
      end
      object GroupBox16: TGroupBox
        Left = 4
        Top = 288
        Width = 487
        Height = 65
        Caption = 'Compteur de Pi'#232'ces'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 10
        object Label37: TLabel
          Left = 8
          Top = 31
          Width = 36
          Height = 13
          Caption = 'Devis : '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label39: TLabel
          Left = 166
          Top = 31
          Width = 42
          Height = 13
          Caption = 'Facture :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label40: TLabel
          Left = 329
          Top = 31
          Width = 30
          Height = 13
          Caption = 'Avoir :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object CPTDEVIS: THSpinEdit
          Left = 44
          Top = 28
          Width = 113
          Height = 22
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxValue = 0
          MinValue = 0
          ParentFont = False
          TabOrder = 0
          Value = 0
        end
        object CPTFACTURE: THSpinEdit
          Left = 213
          Top = 28
          Width = 109
          Height = 22
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxValue = 0
          MinValue = 0
          ParentFont = False
          TabOrder = 1
          Value = 0
        end
        object CPTAVOIR: THSpinEdit
          Left = 362
          Top = 28
          Width = 121
          Height = 22
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxValue = 0
          MinValue = 0
          ParentFont = False
          TabOrder = 2
          Value = 0
        end
      end
    end
    object CODEARTICLE: TTabSheet
      Caption = 'Codification Article'
      ImageIndex = 7
      object HLabel7: THLabel
        Left = 3
        Top = 5
        Width = 486
        Height = 23
        Alignment = taCenter
        AutoSize = False
        Caption = 'Codification Article'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = 20
        Font.Name = 'Cambria'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object GroupBox10: TGroupBox
        Left = 3
        Top = 29
        Width = 486
        Height = 7
        TabOrder = 0
      end
      object GBSUFFIXE: TGroupBox
        Left = 4
        Top = 215
        Width = 518
        Height = 49
        Caption = 'Gestion Suffixes Article'
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        object SO_GCTYPSUFART: TCheckBox
          Left = 8
          Top = 22
          Width = 193
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Utilisation du Libell'#233' comme suffixe :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
      end
      object GBPREFIXE: TGroupBox
        Left = 4
        Top = 84
        Width = 517
        Height = 129
        Caption = 'Gestion Pr'#233'fixes Article'
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        object LGCPREFIXEART: TLabel
          Left = 10
          Top = 26
          Width = 138
          Height = 13
          Caption = 'Pr'#233'fixe Marchandise (Alpha) :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LGCPREFIXENOM: TLabel
          Left = 10
          Top = 50
          Width = 118
          Height = 13
          Caption = 'Pr'#233'fixe Ouvrage (Alpha) :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LGCPREFIXEPRE: TLabel
          Left = 10
          Top = 74
          Width = 124
          Height = 13
          Caption = 'Pr'#233'fixe Prestation (Alpha) :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LGCPREFIXEPOS: TLabel
          Left = 10
          Top = 101
          Width = 121
          Height = 13
          Caption = 'Pr'#233'fixe Prix Pos'#233' (Alpha) :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object SO_GCPREFIXEART: TEdit
          Left = 158
          Top = 25
          Width = 131
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 9
          ParentFont = False
          TabOrder = 0
        end
        object SO_GCPREFIXENOM: TEdit
          Left = 158
          Top = 49
          Width = 131
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 9
          ParentFont = False
          TabOrder = 1
        end
        object SO_GCPREFIXEPRE: TEdit
          Left = 158
          Top = 73
          Width = 131
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 9
          ParentFont = False
          TabOrder = 2
        end
        object SO_BTPREFIXEPOS: TEdit
          Left = 158
          Top = 97
          Width = 131
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 9
          ParentFont = False
          TabOrder = 3
        end
      end
      object GBPARAMDOC: TGroupBox
        Left = 4
        Top = 266
        Width = 518
        Height = 191
        Caption = 'Param'#232'tres Documents'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        object Label52: TLabel
          Left = 9
          Top = 27
          Width = 159
          Height = 13
          Caption = 'Affectation '#233'cart sur prix March'#233' :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label53: TLabel
          Left = 9
          Top = 58
          Width = 93
          Height = 13
          Caption = 'D'#233'cimale Quantit'#233' :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label54: TLabel
          Left = 174
          Top = 58
          Width = 70
          Height = 13
          Caption = 'Code Devise : '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label55: TLabel
          Left = 318
          Top = 59
          Width = 112
          Height = 13
          Caption = 'D'#233'cimale Prix Unitaire : '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label56: TLabel
          Left = 9
          Top = 87
          Width = 267
          Height = 13
          Caption = 'Type de ressource associ'#233'e '#224' la main d'#39'oeuvre Interne : '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label57: TLabel
          Left = 9
          Top = 115
          Width = 238
          Height = 13
          Caption = 'Article de remplacement pour pr'#233'vision de chantier'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TSO_ECARTFACTURATION: TLabel
          Left = 9
          Top = 139
          Width = 169
          Height = 13
          Caption = 'Ecart constat'#233' sur derni'#232're situation'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TSO_PRESTASSOC: TLabel
          Left = 9
          Top = 163
          Width = 215
          Height = 13
          Caption = 'Prestation associ'#233'e / article prix pos'#233' (d'#233'faut)'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object SO_BTECARTPMA: THCritMaskEdit
          Left = 278
          Top = 23
          Width = 187
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          TagDispatch = 0
          DataType = 'BTARTICLES'
          ElipsisButton = True
        end
        object SO_CPCODEEUROS3: THCritMaskEdit
          Left = 246
          Top = 54
          Width = 67
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          TagDispatch = 0
          DataType = 'TTDEVISE'
          ElipsisButton = True
        end
        object SO_BTMOINTERNE: THMultiValComboBox
          Left = 278
          Top = 84
          Width = 201
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          Abrege = False
          DataType = 'AFTTYPERESSOURCE'
          Complete = False
          OuInclusif = False
        end
        object SO_BTREPLPOURCENT: THCritMaskEdit
          Left = 278
          Top = 111
          Width = 225
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          TagDispatch = 0
          Plus = 'GA_TYPEARTICLE = "MAR"'
          DataType = 'BTARTICLES'
          ElipsisButton = True
        end
        object SO_DECQTE: THSpinEdit
          Left = 112
          Top = 53
          Width = 52
          Height = 22
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxValue = 0
          MinValue = 0
          ParentFont = False
          TabOrder = 4
          Value = 0
        end
        object SO_DECPRIX: THSpinEdit
          Left = 431
          Top = 54
          Width = 44
          Height = 22
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxValue = 0
          MinValue = 0
          ParentFont = False
          TabOrder = 5
          Value = 0
        end
        object SO_ECARTFACTURATION: THCritMaskEdit
          Left = 278
          Top = 135
          Width = 225
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 6
          TagDispatch = 0
          Plus = 'GA_TYPEARTICLE = "MAR"'
          DataType = 'BTARTICLES'
          ElipsisButton = True
        end
        object SO_BTPRESTPRIXPOS: THCritMaskEdit
          Left = 278
          Top = 160
          Width = 153
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 7
          TagDispatch = 0
          Plus = 'GA_TYPEARTICLE = "PRE"'
          DataType = 'BTPRESTATION'
          ElipsisButton = True
        end
      end
      object GBCODEART00: TGroupBox
        Left = 3
        Top = 45
        Width = 517
        Height = 38
        TabOrder = 1
        object LGCLGNUMART: TLabel
          Left = 178
          Top = 13
          Width = 84
          Height = 13
          Caption = 'Longeur du code '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LGCCOMPTEURART: TLabel
          Left = 335
          Top = 13
          Width = 105
          Height = 13
          Caption = 'Dernier Chrono Utilis'#233' '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object SO_GCNUMARTAUTO: TCheckBox
          Left = 3
          Top = 11
          Width = 142
          Height = 17
          Alignment = taLeftJustify
          Caption = 'G'#233'n'#233'ration Automatique '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = SO_GCNUMARTAUTOClick
        end
        object SO_GCLGNUMART: THSpinEdit
          Left = 272
          Top = 8
          Width = 58
          Height = 22
          MaxValue = 0
          MinValue = 0
          TabOrder = 1
          Value = 0
        end
        object SO_GCCOMPTEURART: TEdit
          Left = 442
          Top = 9
          Width = 44
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 18
          ParentFont = False
          TabOrder = 2
          Text = '0'
        end
      end
    end
    object ParamRessources: TTabSheet
      Caption = 'ParamRessources'
      ImageIndex = 11
      object HLabel11: THLabel
        Left = 3
        Top = 5
        Width = 486
        Height = 23
        Alignment = taCenter
        AutoSize = False
        Caption = 'Param'#233'trages Ressources'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = 20
        Font.Name = 'Cambria'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label59: TLabel
        Left = 1
        Top = 47
        Width = 104
        Height = 13
        Caption = 'Prestation par d'#233'faut :'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object LblPrestationRes: TLabel
        Left = 232
        Top = 48
        Width = 9
        Height = 13
        Caption = '...'
      end
      object GroupBox3: TGroupBox
        Left = 3
        Top = 29
        Width = 486
        Height = 7
        TabOrder = 0
      end
      object SO_AFPRESTATIONRES: THCritMaskEdit
        Left = 112
        Top = 43
        Width = 113
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnExit = SO_AFPRESTATIONRESExit
        TagDispatch = 0
        Plus = 'GA_TYPEARTICLE = "PRE"'
        DataType = 'BTPRESTATION'
        ElipsisButton = True
      end
      object GroupBox5: TGroupBox
        Left = 3
        Top = 72
        Width = 493
        Height = 97
        Caption = 'Valorisation des Ressources'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        object Label58: TLabel
          Left = 10
          Top = 60
          Width = 165
          Height = 13
          Caption = 'Taux de frais g'#233'n'#233'raux par d'#233'faut :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object SO_AFRESCALCULPR: TCheckBox
          Left = 11
          Top = 26
          Width = 206
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Calcul Automatique du prix de Revient :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = SO_GCNUMARTAUTOClick
        end
        object SO_AFFRAISGEN1: THCritMaskEdit
          Left = 184
          Top = 56
          Width = 113
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          Text = '0.00000'
          TagDispatch = 0
          OpeType = otReel
        end
      end
    end
    object SauveRestaure: TTabSheet
      Caption = 'Sauvegarde/Restauration'
      ImageIndex = 6
      object HLabel2: THLabel
        Left = 3
        Top = 5
        Width = 486
        Height = 23
        Alignment = taCenter
        AutoSize = False
        Caption = 'Param'#232'tres Sauvegarde/Restauration'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = 20
        Font.Name = 'Cambria'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object LREPSAUVE: TLabel
        Left = 3
        Top = 83
        Width = 120
        Height = 13
        Caption = 'Chemin de Sauvegarde : '
      end
      object LHOWLASTSAVE: TLabel
        Left = 275
        Top = 139
        Width = 51
        Height = 13
        Caption = 'Garder les '
      end
      object LLASTSAVE: TLabel
        Left = 371
        Top = 139
        Width = 109
        Height = 13
        Caption = 'Derni'#232'res sauvegardes'
      end
      object GroupBox6: TGroupBox
        Left = 3
        Top = 29
        Width = 486
        Height = 7
        TabOrder = 0
      end
      object SO_FICHIERSAUVE: THCritMaskEdit
        Left = 128
        Top = 81
        Width = 361
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        TagDispatch = 0
        DataType = 'DIRECTORY'
        ElipsisButton = True
      end
      object SO_SAUVEAUTO: TCheckBox
        Left = 3
        Top = 51
        Width = 286
        Height = 17
        Caption = 'Sauvegarde Automatique '#224' la fermeture de l'#39'application'
        TabOrder = 2
      end
      object SO_ASKBEFORESAVE: TCheckBox
        Left = 3
        Top = 115
        Width = 174
        Height = 17
        Caption = 'Confirmation avant Sauvegarde'
        TabOrder = 3
      end
      object SO_ZIPFILE: TCheckBox
        Left = 3
        Top = 139
        Width = 198
        Height = 17
        Caption = 'Compression du fichier Sauvegard'#233
        TabOrder = 4
      end
      object SO_HOWLASTSAVE: THSpinEdit
        Left = 328
        Top = 134
        Width = 40
        Height = 22
        MaxValue = 0
        MinValue = 0
        TabOrder = 5
        Value = 0
      end
    end
    object ExoComptable: TTabSheet
      Caption = 'Exercice comptable'
      ImageIndex = 9
      object HLabel4: THLabel
        Left = 3
        Top = 5
        Width = 486
        Height = 23
        Alignment = taCenter
        AutoSize = False
        Caption = 'Gestion des Exercices Comptables'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = 20
        Font.Name = 'Cambria'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object BCreatNewExo: TToolbarButton97
        Left = 144
        Top = 263
        Width = 206
        Height = 34
        Alignment = taLeftJustify
        Caption = 'Cr'#233'ation Nouvel Exercice'
        ImageIndex = 0
        Opaque = False
        Visible = False
        OnClick = BCreatNewExoClick
      end
      object GroupBox17: TGroupBox
        Left = 3
        Top = 29
        Width = 486
        Height = 7
        TabOrder = 0
      end
      object GRPEXOENCOURS: TGroupBox
        Left = 0
        Top = 46
        Width = 489
        Height = 115
        Caption = 'Exercice En Cours'
        Font.Charset = ANSI_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 1
        object HLabel6: THLabel
          Left = 268
          Top = 82
          Width = 14
          Height = 13
          Caption = 'Fin'
          FocusControl = EX_DATEFIN
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel10: THLabel
          Left = 10
          Top = 82
          Width = 29
          Height = 13
          Caption = 'D'#233'but'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label41: TLabel
          Left = 10
          Top = 56
          Width = 36
          Height = 13
          Caption = 'Libelle :'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label42: TLabel
          Left = 10
          Top = 28
          Width = 31
          Height = 13
          Caption = 'Code :'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label43: TLabel
          Left = 154
          Top = 28
          Width = 37
          Height = 13
          Caption = 'Abreg'#233' '
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object EX_DATEFIN: TMaskEdit
          Left = 288
          Top = 78
          Width = 76
          Height = 21
          Ctl3D = True
          EditMask = '!99/99/0000;1;_'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          MaxLength = 10
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 4
          Text = '01/01/1900'
        end
        object EX_DATEDEBUT: TMaskEdit
          Left = 53
          Top = 78
          Width = 78
          Height = 21
          Ctl3D = True
          EditMask = '!99/99/0000;1;_'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          MaxLength = 10
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 3
          Text = '01/01/1900'
        end
        object EX_LIBELLE: TEdit
          Left = 53
          Top = 51
          Width = 310
          Height = 21
          Font.Charset = ANSI_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 2
        end
        object EX_EXERCICE: TEdit
          Left = 53
          Top = 24
          Width = 73
          Height = 21
          Enabled = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 0
        end
        object EX_ABREGE: TEdit
          Left = 199
          Top = 24
          Width = 165
          Height = 21
          Font.Charset = ANSI_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 1
        end
      end
      object GroupBox18: TGroupBox
        Left = 0
        Top = 166
        Width = 489
        Height = 79
        Caption = 'Etat '
        Font.Charset = ANSI_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 2
        object Label45: TLabel
          Left = 10
          Top = 28
          Width = 66
          Height = 13
          Caption = 'Comptabilit'#233'  :'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label46: TLabel
          Left = 10
          Top = 52
          Width = 40
          Height = 13
          Caption = 'Budget :'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object EX_ETATCPTA: TEdit
          Left = 82
          Top = 24
          Width = 165
          Height = 21
          Enabled = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 0
        end
        object EX_ETATBUDGET: TEdit
          Left = 82
          Top = 48
          Width = 165
          Height = 21
          Enabled = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 1
        end
      end
    end
    object InfoG: TTabSheet
      Caption = 'Liaison Comptable'
      object LTitre6: THLabel
        Left = 3
        Top = 5
        Width = 486
        Height = 23
        Alignment = taCenter
        AutoSize = False
        Caption = 'Option de Liaison Comptable'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = 20
        Font.Name = 'Cambria'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object MemoLog: TRichEdit
        Left = 8
        Top = 80
        Width = 481
        Height = 281
        BorderStyle = bsNone
        Color = clInactiveCaptionText
        HideScrollBars = False
        Lines.Strings = (
          '')
        PlainText = True
        TabOrder = 3
        Visible = False
      end
      object PTransfert: TPanel
        Left = 64
        Top = 88
        Width = 353
        Height = 225
        BevelOuter = bvNone
        TabOrder = 2
        Visible = False
        object HLabel3: THLabel
          Left = 8
          Top = 11
          Width = 87
          Height = 13
          Caption = 'Fichier de transfert'
        end
        object BLanceRecupParam: TToolbarButton97
          Left = 73
          Top = 78
          Width = 206
          Height = 35
          Alignment = taLeftJustify
          Caption = '   Lancer la r'#233'cup'#233'ration'
          ImageIndex = 0
          Opaque = False
          OnClick = BLanceRecupParamClick
        end
        object FileTrF: THCritMaskEdit
          Left = 6
          Top = 31
          Width = 340
          Height = 21
          TabOrder = 0
          TagDispatch = 0
          ElipsisButton = True
          OnElipsisClick = FileTrFElipsisClick
        end
        object AnimGetParam: TAnimate
          Left = 40
          Top = 144
          Width = 272
          Height = 60
          CommonAVI = aviCopyFiles
          StopFrame = 1
        end
      end
      object GroupBox2: TGroupBox
        Left = 3
        Top = 29
        Width = 486
        Height = 7
        TabOrder = 0
      end
      object PChoixLiaison: TPanel
        Left = 64
        Top = 80
        Width = 353
        Height = 281
        BevelOuter = bvNone
        TabOrder = 1
        object CPTAEX: TRadioButton
          Left = 53
          Top = 21
          Width = 242
          Height = 14
          Caption = 'Comptabilit'#233' tenue par votre expert comptable'
          TabOrder = 0
          OnClick = CPTAEXClick
        end
        object CPTALINE: TRadioButton
          Left = 53
          Top = 42
          Width = 182
          Height = 14
          Caption = 'Comptabilit'#233' Line tenue par vous'
          TabOrder = 1
          OnClick = CPTALINEClick
        end
        object CPTAWINNER: TRadioButton
          Left = 53
          Top = 64
          Width = 262
          Height = 14
          Caption = 'Comptabilit'#233' Win Winner tenue par vous'
          TabOrder = 2
          OnClick = CPTAWINNERClick
        end
        object CPTAQUADRA: TRadioButton
          Left = 53
          Top = 86
          Width = 262
          Height = 14
          Caption = 'Comptabilit'#233' Quadratus tenue par vous'
          TabOrder = 3
          OnClick = CPTAQUADRAClick
        end
        object CPTASUITE: TRadioButton
          Left = 53
          Top = 106
          Width = 262
          Height = 17
          Caption = 'Comptabilit'#233' Suite tenue par vous'
          TabOrder = 4
          OnClick = CPTASUITEClick
        end
        object NOCPTA: TRadioButton
          Left = 53
          Top = 127
          Width = 262
          Height = 17
          Caption = 'Aucune ou autres comptabilit'#233's'
          Checked = True
          TabOrder = 5
          TabStop = True
          OnClick = NOCPTAClick
        end
        object IMPORTPARAM: TCheckBox
          Left = 47
          Top = 188
          Width = 275
          Height = 17
          Hint = 'Pour int'#233'grer vos param'#233'trages comptables depuis un fichier .TRA'
          Alignment = taLeftJustify
          Caption = 'R'#233'cup'#233'ration des param'#232'tres comptables'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 6
        end
      end
    end
    object PREFCOMPTA1: TTabSheet
      Caption = 'Pr'#233'f'#233'rences Comptables'
      ImageIndex = 8
      object HLabel8: THLabel
        Left = 3
        Top = 5
        Width = 486
        Height = 23
        Alignment = taCenter
        AutoSize = False
        Caption = 'Pr'#233'f'#233'rences Comptables (1)'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = 20
        Font.Name = 'Cambria'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Line11: TGroupBox
        Left = 3
        Top = 29
        Width = 486
        Height = 7
        TabOrder = 0
      end
      object GBVENTILE: TGroupBox
        Left = 4
        Top = 51
        Width = 487
        Height = 54
        Caption = 'Ventilation Comptable'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object SO_GCVENTCPTAART: TCheckBox
          Left = 8
          Top = 22
          Width = 113
          Height = 17
          Caption = 'Ventilation Article'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object SO_GCVENTCPTAAFF: TCheckBox
          Left = 360
          Top = 22
          Width = 113
          Height = 17
          Caption = 'Ventilation Chantier'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
        object SO_GCVENTCPTATIERS: TCheckBox
          Left = 182
          Top = 22
          Width = 97
          Height = 17
          Caption = 'Ventilation Client'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
        end
      end
      object GBJOURNAUX: TGroupBox
        Left = 4
        Top = 104
        Width = 487
        Height = 107
        Caption = 'Journaux Comptable'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 2
        object Label17: TLabel
          Left = 10
          Top = 26
          Width = 85
          Height = 13
          Caption = 'Journaux Export : '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label18: TLabel
          Left = 10
          Top = 53
          Width = 143
          Height = 13
          Caption = 'Exclure les Transfert de type...'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label19: TLabel
          Left = 10
          Top = 80
          Width = 68
          Height = 13
          Caption = '... et de type : '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object SO_EXPJRNX: THMultiValComboBox
          Left = 96
          Top = 24
          Width = 385
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          Abrege = False
          DataType = 'CPJOURNAL'
          Complete = False
          OuInclusif = False
        end
        object SO_MBOCPTAEXCLPA1: THMultiValComboBox
          Left = 160
          Top = 48
          Width = 321
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          Abrege = False
          DataType = 'MBOCPTAEXCLPA1'
          Complete = False
          OuInclusif = False
        end
        object SO_MBOCPTAEXCLPA2: THMultiValComboBox
          Left = 81
          Top = 77
          Width = 400
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          Abrege = False
          DataType = 'MBOCPTAEXCLPA2'
          Complete = False
          OuInclusif = False
        end
      end
      object GbPREFAUX: TGroupBox
        Left = 4
        Top = 290
        Width = 487
        Height = 79
        Caption = 'Gestion R'#233'glements/Acomptes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        object Label60: TLabel
          Left = 211
          Top = 24
          Width = 81
          Height = 13
          Caption = 'Lien comptable : '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label61: TLabel
          Left = 211
          Top = 53
          Width = 120
          Height = 13
          Caption = 'Mode Paiement Associ'#233' :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object SO_BTCOMPTAREGL: TCheckBox
          Left = 8
          Top = 22
          Width = 185
          Height = 17
          Caption = 'Comptabilisation des R'#233'glements '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object SO_BTOPTRGCOLLTIE: TCheckBox
          Left = 8
          Top = 51
          Width = 177
          Height = 17
          Caption = 'Ecritures R.G. sur Collectif tiers'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
        object SO_BTLIENCPTAS1: THValComboBox
          Left = 336
          Top = 20
          Width = 146
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 2
          TagDispatch = 0
          DataType = 'CPLIENCOMPTABILITE'
        end
        object SO_BTMODEPAIEASS: THValComboBox
          Left = 336
          Top = 49
          Width = 146
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 3
          TagDispatch = 0
          DataType = 'TTMODEPAIE'
        end
      end
      object GBLGCPT: TGroupBox
        Left = 4
        Top = 210
        Width = 487
        Height = 80
        Caption = 'Longueur des Comptes...'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        object Label20: TLabel
          Left = 10
          Top = 55
          Width = 69
          Height = 13
          Caption = '... Auxilliaires : '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label21: TLabel
          Left = 10
          Top = 29
          Width = 67
          Height = 13
          Caption = '... G'#233'n'#233'raux : '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label22: TLabel
          Left = 151
          Top = 29
          Width = 118
          Height = 13
          Caption = 'Caract'#232'res de Bourrage :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label23: TLabel
          Left = 151
          Top = 55
          Width = 118
          Height = 13
          Caption = 'Caract'#232'res de Bourrage :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label24: TLabel
          Left = 316
          Top = 29
          Width = 113
          Height = 13
          Caption = 'Pr'#233'fixe Auxilliaire Client :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label25: TLabel
          Left = 316
          Top = 55
          Width = 114
          Height = 13
          Caption = 'Pr'#233'fixe Auxilliaire fourn. :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object SO_BOURREGEN: TEdit
          Left = 277
          Top = 25
          Width = 33
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 1
          ParentFont = False
          TabOrder = 0
        end
        object SO_BOURREAUX: TEdit
          Left = 277
          Top = 51
          Width = 33
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 1
          ParentFont = False
          TabOrder = 1
        end
        object SO_GCPREFIXEAUXI: TEdit
          Left = 440
          Top = 25
          Width = 41
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 9
          ParentFont = False
          TabOrder = 2
        end
        object SO_GCPREFIXEAUXIFOU: TEdit
          Left = 440
          Top = 51
          Width = 41
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 9
          ParentFont = False
          TabOrder = 3
        end
        object SO_LGCPTEGEN: THSpinEdit
          Left = 96
          Top = 24
          Width = 45
          Height = 22
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxValue = 0
          MinValue = 0
          ParentFont = False
          TabOrder = 4
          Value = 0
        end
        object SO_LGCPTEAUX: THSpinEdit
          Left = 96
          Top = 48
          Width = 45
          Height = 22
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxValue = 0
          MinValue = 0
          ParentFont = False
          TabOrder = 5
          Value = 0
        end
      end
    end
    object PREFCOMPTA2: TTabSheet
      Caption = 'Pr'#233'f'#233'rences Comptables (Suite)'
      ImageIndex = 9
      object HLabel9: THLabel
        Left = 3
        Top = 5
        Width = 486
        Height = 23
        Alignment = taCenter
        AutoSize = False
        Caption = 'Pr'#233'f'#233'rences Comptables (2)'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = 20
        Font.Name = 'Cambria'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object GroupBox13: TGroupBox
        Left = 3
        Top = 29
        Width = 486
        Height = 7
        TabOrder = 0
      end
      object GRPCOLLECTIF: TGroupBox
        Left = 4
        Top = 51
        Width = 487
        Height = 77
        Caption = 'Comptes Collectifs Par D'#233'faut'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object Label26: TLabel
          Left = 10
          Top = 27
          Width = 35
          Height = 13
          Caption = 'Client : '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label27: TLabel
          Left = 10
          Top = 53
          Width = 57
          Height = 13
          Caption = 'Fournisseur '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label28: TLabel
          Left = 318
          Top = 27
          Width = 46
          Height = 13
          Caption = 'D'#233'biteur :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label29: TLabel
          Left = 318
          Top = 53
          Width = 48
          Height = 13
          Caption = 'Cr'#233'diteur :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object SO_DEFCOLCLI: THCritMaskEdit
          Left = 102
          Top = 23
          Width = 99
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          TagDispatch = 0
          DataType = 'TZGCOLLCLIENT'
          ElipsisButton = True
        end
        object SO_DEFCOLFOU: THCritMaskEdit
          Left = 102
          Top = 49
          Width = 99
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          TagDispatch = 0
          DataType = 'TZGCOLLFOURN'
          ElipsisButton = True
        end
        object SO_DEFCOLDDIV: THCritMaskEdit
          Left = 378
          Top = 23
          Width = 99
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          TagDispatch = 0
          DataType = 'TZGCOLLTOUTDEBIT'
          ElipsisButton = True
        end
        object SO_DEFCOLCDIV: THCritMaskEdit
          Left = 378
          Top = 49
          Width = 99
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          TagDispatch = 0
          DataType = 'TZGCOLLTOUTCREDIT'
          ElipsisButton = True
        end
      end
      object GRPCPTACHATVENTE: TGroupBox
        Left = 4
        Top = 129
        Width = 487
        Height = 182
        Caption = 'Comptes Achats/Ventes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        object Label30: TLabel
          Left = 10
          Top = 31
          Width = 96
          Height = 13
          Caption = 'Compte d'#39'escompte '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label31: TLabel
          Left = 10
          Top = 57
          Width = 118
          Height = 13
          Caption = 'Compte Remise sut Pied '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label32: TLabel
          Left = 10
          Top = 80
          Width = 105
          Height = 13
          Caption = 'Compte HT par d'#233'faut'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label33: TLabel
          Left = 10
          Top = 105
          Width = 86
          Height = 13
          Caption = 'Compte Port & Frais'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label34: TLabel
          Left = 10
          Top = 129
          Width = 141
          Height = 13
          Caption = 'Compte Retenue de Garantie '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label35: TLabel
          Left = 10
          Top = 156
          Width = 104
          Height = 13
          Caption = 'Journal de Ventilation '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object RotateLabel1: TRotateLabel
          Left = 159
          Top = 56
          Width = 23
          Height = 65
          Escapement = 90
          TextStyle = tsRecessed
          AutoSize = False
          Caption = 'ACHATS'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -16
          Font.Name = 'Cambria'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
        end
        object RotateLabel2: TRotateLabel
          Left = 354
          Top = 56
          Width = 23
          Height = 65
          Escapement = 90
          TextStyle = tsRecessed
          AutoSize = False
          Caption = 'VENTES'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -16
          Font.Name = 'Cambria'
          Font.Style = [fsBold]
          ParentFont = False
          Transparent = True
        end
        object SO_GCCPTEESCACH: THCritMaskEdit
          Left = 181
          Top = 27
          Width = 99
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          TagDispatch = 0
          DataType = 'TZGPRODUIT'
          ElipsisButton = True
        end
        object SO_GCCPTEREMACH: THCritMaskEdit
          Left = 181
          Top = 53
          Width = 99
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          TagDispatch = 0
          DataType = 'TZGCHARGE'
          ElipsisButton = True
        end
        object SO_GCCPTEESCVTE: THCritMaskEdit
          Left = 378
          Top = 27
          Width = 99
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          TagDispatch = 0
          DataType = 'TZGCHARGE'
          ElipsisButton = True
        end
        object SO_GCCPTEREMVTE: THCritMaskEdit
          Left = 378
          Top = 53
          Width = 99
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          TagDispatch = 0
          DataType = 'TZGPRODUIT'
          ElipsisButton = True
        end
        object SO_GCCPTEHTACH: THCritMaskEdit
          Left = 181
          Top = 76
          Width = 99
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 4
          TagDispatch = 0
          DataType = 'TZGCHARGE'
          ElipsisButton = True
        end
        object SO_GCCPTEPORTACH: THCritMaskEdit
          Left = 181
          Top = 101
          Width = 99
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 5
          TagDispatch = 0
          DataType = 'TZGCHARGE'
          ElipsisButton = True
        end
        object SO_GCCPTEHTVTE: THCritMaskEdit
          Left = 378
          Top = 76
          Width = 99
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 6
          TagDispatch = 0
          DataType = 'TZGPRODUIT'
          ElipsisButton = True
        end
        object SO_GCCPTEPORTVTE: THCritMaskEdit
          Left = 378
          Top = 101
          Width = 99
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 7
          TagDispatch = 0
          DataType = 'TZGPRODUIT'
          ElipsisButton = True
        end
        object SO_GCCPTERGVTE: THCritMaskEdit
          Left = 378
          Top = 125
          Width = 99
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 8
          TagDispatch = 0
          DataType = 'TZGENERAL'
          ElipsisButton = True
        end
        object GPP_JOURNALCPTAA: THValComboBox
          Left = 181
          Top = 152
          Width = 99
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 9
          TagDispatch = 0
          DataType = 'GCJOURNALCPTA'
        end
        object GPP_JOURNALCPTAV: THValComboBox
          Left = 378
          Top = 152
          Width = 99
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 10
          TagDispatch = 0
          DataType = 'GCJOURNALCPTA'
        end
      end
      object GRPCPTECART: TGroupBox
        Left = 4
        Top = 313
        Width = 487
        Height = 60
        Caption = 'Comptes d'#39#233'cart'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        object Label36: TLabel
          Left = 410
          Top = 12
          Width = 27
          Height = 13
          Caption = 'Cr'#233'dit'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label38: TLabel
          Left = 213
          Top = 12
          Width = 25
          Height = 13
          Caption = 'D'#233'bit'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object SO_GCECARTCREDIT: THCritMaskEdit
          Left = 181
          Top = 25
          Width = 99
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          TagDispatch = 0
          DataType = 'TZGENERAL'
          ElipsisButton = True
        end
        object SO_GCECARTDEBIT: THCritMaskEdit
          Left = 378
          Top = 25
          Width = 99
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          TagDispatch = 0
          DataType = 'TZGENERAL'
          ElipsisButton = True
        end
      end
    end
    object EchangeComptable: TTabSheet
      Caption = 'Echange Comptable'
      ImageIndex = 5
      object HLabel1: THLabel
        Left = 3
        Top = 5
        Width = 486
        Height = 23
        Alignment = taCenter
        AutoSize = False
        Caption = 'Echanges Comptable'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = 20
        Font.Name = 'Cambria'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object LCPRDREPERTOIRE: TLabel
        Left = 2
        Top = 51
        Width = 168
        Height = 13
        Caption = 'Emplacement fichier Export (.TRA) :'
      end
      object Label14: TLabel
        Left = 2
        Top = 115
        Width = 98
        Height = 13
        Caption = 'Fr'#233'quence d'#39'Export :'
      end
      object LCPPOINTAGESX: TLabel
        Left = 258
        Top = 166
        Width = 87
        Height = 13
        Caption = 'Gestion Pointage :'
        Visible = False
      end
      object LSO_BTCHEMINCOMSX: TLabel
        Left = 2
        Top = 75
        Width = 179
        Height = 13
        Caption = 'Emplacement de l'#39'executable COMSX'
      end
      object GroupBox8: TGroupBox
        Left = 3
        Top = 29
        Width = 486
        Height = 7
        TabOrder = 0
      end
      object SO_CPRDREPERTOIRE: THCritMaskEdit
        Left = 184
        Top = 49
        Width = 313
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        TagDispatch = 0
        DataType = 'DIRECTORY'
        ElipsisButton = True
      end
      object SO_FREQUENCESX: THValComboBox
        Left = 104
        Top = 110
        Width = 121
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 0
        ParentFont = False
        TabOrder = 2
        TagDispatch = 0
        DataType = 'DPPERIODICITE'
      end
      object GrpResCabinet: TGroupBox
        Left = 4
        Top = 189
        Width = 487
        Height = 116
        Caption = 'Renseignement Cabinet Comptable'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        object Label15: TLabel
          Left = 10
          Top = 30
          Width = 34
          Height = 13
          Caption = 'Nom  : '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object Label16: TLabel
          Left = 10
          Top = 59
          Width = 69
          Height = 13
          Caption = 'Adresse Mail : '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BEMAIL: TToolbarButton97
          Left = 440
          Top = 49
          Width = 39
          Height = 26
          OnClick = BEMAILClick
        end
        object LCPRDSIRET: TLabel
          Left = 10
          Top = 86
          Width = 56
          Height = 13
          Caption = 'N'#176' SIRET : '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LCPRDDATERECEPTION: TLabel
          Left = 218
          Top = 86
          Width = 128
          Height = 13
          Caption = 'Date Derni'#232're Int'#233'gration : '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object SO_CPRDNOMCLIENT: TEdit
          Left = 48
          Top = 25
          Width = 433
          Height = 21
          Hint = 'Raison Sociale Cabinet comptable'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object SO_CPRDEMAILCLIENT: TEdit
          Left = 80
          Top = 54
          Width = 361
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
        object SO_CPRDSIRET: TEdit
          Left = 72
          Top = 81
          Width = 129
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
        end
        object SO_CPRDDATERECEPTION: TEdit
          Left = 352
          Top = 81
          Width = 129
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
        end
      end
      object SO_CPMODESYNCHRO: TCheckBox
        Left = 8
        Top = 168
        Width = 105
        Height = 17
        Caption = 'Synchronisation '
        TabOrder = 4
        Visible = False
      end
      object SO_CPPOINTAGESX: THValComboBox
        Left = 352
        Top = 161
        Width = 137
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 0
        ParentFont = False
        TabOrder = 5
        Visible = False
        TagDispatch = 0
      end
      object SO_BTCHEMINCOMSX: THCritMaskEdit
        Left = 184
        Top = 73
        Width = 313
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 6
        TagDispatch = 0
        DataType = 'DIRECTORY'
        ElipsisButton = True
      end
    end
    object PRESENTDOC1: TTabSheet
      Caption = 'Pr'#233'sentation Doc (1)'
      ImageIndex = 10
      object LTitre10: THLabel
        Left = 3
        Top = 5
        Width = 486
        Height = 23
        Alignment = taCenter
        AutoSize = False
        Caption = 'Pr'#233'f'#233'rences Document Type Devis'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = 20
        Font.Name = 'Cambria'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object HautDePage: TToolbarButton97
        Left = 273
        Top = 374
        Width = 105
        Height = 25
        Caption = 'Haut de Page'
        Flat = False
        OnClick = HautDePageClick
      end
      object BasDePage: TToolbarButton97
        Left = 385
        Top = 374
        Width = 105
        Height = 25
        Caption = 'Bas de Page'
        Flat = False
        OnClick = BasDePageClick
      end
      object BTYPEDOC: TToolbarButton97
        Left = 161
        Top = 374
        Width = 105
        Height = 25
        Caption = 'Options Edition'
        Flat = False
        OnClick = BTYPEDOCClick
      end
      object C_IMPBASPAGE: TCheckBox
        Left = 248
        Top = 328
        Width = 241
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Impression txt bas de page sur toutes les pages '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 34
        OnClick = C_IMPBASPAGEClick
      end
      object Ligne10: TGroupBox
        Left = 3
        Top = 29
        Width = 486
        Height = 7
        TabOrder = 0
      end
      object GB_Lignes: TGroupBox
        Left = 4
        Top = 50
        Width = 487
        Height = 48
        Caption = 'Pr'#233'sentation Des Lignes '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object CB_CODE: TCheckBox
          Left = 10
          Top = 23
          Width = 49
          Height = 17
          Caption = 'Code'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = CB_CODEClick
        end
        object CB_LIBELLE: TCheckBox
          Left = 81
          Top = 23
          Width = 49
          Height = 17
          Caption = 'Libell'#233
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnClick = CB_LIBELLEClick
        end
        object CB_QTE: TCheckBox
          Left = 161
          Top = 23
          Width = 64
          Height = 17
          Caption = 'Quantit'#233
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          OnClick = CB_QTEClick
        end
        object CB_UNITE: TCheckBox
          Left = 249
          Top = 23
          Width = 49
          Height = 17
          Caption = 'Unit'#233
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          OnClick = CB_UNITEClick
        end
        object CB_PU: TCheckBox
          Left = 312
          Top = 23
          Width = 81
          Height = 17
          Caption = 'Prix Unitaire'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 4
          OnClick = CB_PUClick
        end
        object CB_MONTANT: TCheckBox
          Left = 417
          Top = 23
          Width = 64
          Height = 17
          Caption = 'Montant'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 5
          OnClick = CB_MONTANTClick
        end
      end
      object GB_DESCRIPTIF: TGroupBox
        Left = 4
        Top = 99
        Width = 487
        Height = 66
        Caption = 'Pr'#233'sentation Des Descriptifs'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        object RB_SDESC: TRadioButton
          Left = 10
          Top = 24
          Width = 105
          Height = 17
          Caption = 'Sans Descriptif'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = RB_SDESCClick
        end
        object RB_IDESC: TRadioButton
          Left = 178
          Top = 24
          Width = 105
          Height = 17
          Caption = 'Idem '#224' la Saisie'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnClick = RB_IDESCClick
        end
        object RB_TDESC: TRadioButton
          Left = 360
          Top = 24
          Width = 115
          Height = 17
          Caption = 'Tous les Descriptifs'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          OnClick = RB_TDESCClick
        end
        object C_DESCREMPLACE: TCheckBox
          Left = 144
          Top = 43
          Width = 161
          Height = 17
          Caption = 'En Remplacement du Libell'#233
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          OnClick = C_DESCREMPLACEClick
        end
      end
      object GB_SOUSDETAIL: TGroupBox
        Left = 4
        Top = 165
        Width = 486
        Height = 46
        Caption = 'Pr'#233'sentation Des Sous-D'#233'tails d'#39'Ouvrages'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        object RB_SSD1: TRadioButton
          Left = 10
          Top = 22
          Width = 105
          Height = 17
          Caption = 'Sans Sous-D'#233'tails'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = RB_SSD1Click
        end
        object RB_SSD2: TRadioButton
          Left = 178
          Top = 22
          Width = 105
          Height = 17
          Caption = 'Idem '#224' la Saisie'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnClick = RB_SSD2Click
        end
        object RB_SSD3: TRadioButton
          Left = 360
          Top = 22
          Width = 115
          Height = 17
          Caption = 'Personnalis'#233
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          OnClick = RB_SSD3Click
        end
      end
      object GB_PERSO: TGroupBox
        Left = 4
        Top = 213
        Width = 487
        Height = 48
        Caption = 'Personnalisation des lignes Sous-D'#233'tails d'#39'Ouvrages'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBackground
        Font.Height = -16
        Font.Name = 'Cambria'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        Visible = False
        object CB_CODE1: TCheckBox
          Left = 10
          Top = 23
          Width = 49
          Height = 17
          Caption = 'Code'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = CB_CODE1Click
        end
        object CB_LIBELLE1: TCheckBox
          Left = 81
          Top = 23
          Width = 49
          Height = 17
          Caption = 'Libell'#233
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnClick = CB_LIBELLE1Click
        end
        object CB_QTE1: TCheckBox
          Left = 161
          Top = 23
          Width = 64
          Height = 17
          Caption = 'Quantit'#233
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          OnClick = CB_QTE1Click
        end
        object CB_UNITE1: TCheckBox
          Left = 249
          Top = 23
          Width = 49
          Height = 17
          Caption = 'Unit'#233
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          OnClick = CB_UNITE1Click
        end
        object CB_PU1: TCheckBox
          Left = 312
          Top = 23
          Width = 81
          Height = 17
          Caption = 'Prix Unitaire'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 4
          OnClick = CB_PU1Click
        end
        object CB_MONTANT1: TCheckBox
          Left = 417
          Top = 23
          Width = 64
          Height = 17
          Caption = 'Montant'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 5
          OnClick = CB_MONTANT1Click
        end
      end
      object C_SAUTTXTDEB: TCheckBox
        Left = 10
        Top = 264
        Width = 207
        Height = 17
        Caption = 'Saut de page avant le Texte de d'#233'but'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 5
        OnClick = C_SAUTTXTDEBClick
      end
      object C_SAUTTXTFIN: TCheckBox
        Left = 296
        Top = 264
        Width = 193
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Saut de page avant le Texte de fin'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 6
        OnClick = C_SAUTTXTFINClick
      end
      object C_IMPTOTPAR: TCheckBox
        Left = 10
        Top = 280
        Width = 207
        Height = 17
        Caption = 'Impression des Totaux de Paragraphes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 7
        OnClick = C_IMPTOTPARClick
      end
      object C_IMPTOTSSP: TCheckBox
        Left = 256
        Top = 280
        Width = 233
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Impression des totaux de sous-paragraphes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 8
        OnClick = C_IMPTOTSSPClick
      end
      object C_IMPRECPAR: TCheckBox
        Left = 10
        Top = 296
        Width = 231
        Height = 17
        Caption = 'Impression du r'#233'capitulatif des paragraphes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 9
        OnClick = C_IMPRECPARClick
      end
      object C_IMPMETRE: TCheckBox
        Left = 376
        Top = 296
        Width = 113
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Impression M'#233'tr'#233's'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 10
        OnClick = C_IMPMETREClick
      end
      object C_IMPRECSIT: TCheckBox
        Left = 10
        Top = 312
        Width = 231
        Height = 17
        Caption = 'Impression d'#233'tail des situations pr'#233'c'#233'dentes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 11
        OnClick = C_IMPRECSITClick
      end
      object C_TYPBLOCNOTE: TCheckBox
        Left = 304
        Top = 312
        Width = 185
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Texte du corps sur toute la largeur'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 12
        OnClick = C_TYPBLOCNOTEClick
      end
      object C_IMPTABTOTSIT: TCheckBox
        Left = 10
        Top = 328
        Width = 223
        Height = 17
        Caption = 'Impression tableau totalisations situation'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 13
        OnClick = C_IMPTABTOTSITClick
      end
      object C_IMPRVIATOB: TCheckBox
        Left = 400
        Top = 344
        Width = 89
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Edition directe'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 14
        OnClick = C_IMPRVIATOBClick
      end
      object C_IMPCOLONNES: TCheckBox
        Left = 10
        Top = 344
        Width = 143
        Height = 17
        Caption = 'Impression des colonnes'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 15
        OnClick = C_IMPCOLONNESClick
      end
      object BPD_IMPCOLONNES: TEdit
        Left = 16
        Top = 441
        Width = 25
        Height = 21
        BorderStyle = bsNone
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 16
        Visible = False
      end
      object BPD_TYPEPRES: TSpinEdit
        Left = 18
        Top = 416
        Width = 49
        Height = 22
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxValue = 0
        MinValue = 0
        ParentFont = False
        TabOrder = 17
        Value = 0
        Visible = False
      end
      object BPD_TYPESSD: TSpinEdit
        Left = 144
        Top = 416
        Width = 57
        Height = 22
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxValue = 0
        MinValue = 0
        ParentFont = False
        TabOrder = 18
        Value = 0
        Visible = False
      end
      object BPD_NUMPIECE: TSpinEdit
        Left = 80
        Top = 416
        Width = 57
        Height = 22
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxValue = 0
        MinValue = 0
        ParentFont = False
        TabOrder = 19
        Value = 0
        Visible = False
      end
      object BPD_IMPDESCRIPTIF: TEdit
        Left = 48
        Top = 441
        Width = 25
        Height = 21
        Hint = 'Raison Sociale Cabinet comptable'
        BorderStyle = bsNone
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 20
        Visible = False
      end
      object BPD_IMPRECSIT: TEdit
        Left = 80
        Top = 441
        Width = 25
        Height = 21
        Hint = 'Raison Sociale Cabinet comptable'
        BorderStyle = bsNone
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 21
        Visible = False
      end
      object BPD_IMPRECPAR: TEdit
        Left = 112
        Top = 441
        Width = 25
        Height = 21
        Hint = 'Raison Sociale Cabinet comptable'
        BorderStyle = bsNone
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 22
        Visible = False
      end
      object BPD_SAUTAPRTXTDEB: TEdit
        Left = 144
        Top = 441
        Width = 25
        Height = 21
        Hint = 'Raison Sociale Cabinet comptable'
        BorderStyle = bsNone
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 23
        Visible = False
      end
      object BPD_SAUTAVTTXTFIN: TEdit
        Left = 176
        Top = 441
        Width = 25
        Height = 21
        Hint = 'Raison Sociale Cabinet comptable'
        BorderStyle = bsNone
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 24
        Visible = False
      end
      object BPD_IMPTOTPAR: TEdit
        Left = 208
        Top = 441
        Width = 25
        Height = 21
        Hint = 'Raison Sociale Cabinet comptable'
        BorderStyle = bsNone
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 25
        Visible = False
      end
      object BPD_IMPTOTSSP: TEdit
        Left = 240
        Top = 441
        Width = 25
        Height = 21
        Hint = 'Raison Sociale Cabinet comptable'
        BorderStyle = bsNone
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 26
        Visible = False
      end
      object BPD_NATUREPIECE: TEdit
        Left = 270
        Top = 440
        Width = 25
        Height = 21
        Hint = 'Raison Sociale Cabinet comptable'
        BorderStyle = bsNone
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 27
        Visible = False
      end
      object BPD_SOUCHE: TEdit
        Left = 462
        Top = 440
        Width = 25
        Height = 21
        Hint = 'Raison Sociale Cabinet comptable'
        BorderStyle = bsNone
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 28
        Visible = False
      end
      object BPD_DESCREMPLACE: TEdit
        Left = 430
        Top = 440
        Width = 25
        Height = 21
        Hint = 'Raison Sociale Cabinet comptable'
        BorderStyle = bsNone
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 29
        Visible = False
      end
      object BPD_IMPRVIATOB: TEdit
        Left = 398
        Top = 440
        Width = 25
        Height = 21
        Hint = 'Raison Sociale Cabinet comptable'
        BorderStyle = bsNone
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 30
        Visible = False
      end
      object BPD_TYPBLOCNOTE: TEdit
        Left = 366
        Top = 440
        Width = 25
        Height = 21
        Hint = 'Raison Sociale Cabinet comptable'
        BorderStyle = bsNone
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 31
        Visible = False
      end
      object BPD_IMPTABTOTSIT: TEdit
        Left = 334
        Top = 440
        Width = 25
        Height = 21
        Hint = 'Raison Sociale Cabinet comptable'
        BorderStyle = bsNone
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 32
        Visible = False
      end
      object BPD_IMPMETRE: TEdit
        Left = 302
        Top = 440
        Width = 25
        Height = 21
        Hint = 'Raison Sociale Cabinet comptable'
        BorderStyle = bsNone
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 33
        Visible = False
      end
      object HautBas: THRichEditOLE
        Left = 4
        Top = 48
        Width = 488
        Height = 313
        Color = clInactiveCaptionText
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        TabOrder = 35
        Visible = False
        Margins.Top = 0
        Margins.Bottom = 0
        Margins.Left = 0
        Margins.Right = 0
        ContainerName = 'Document'
        ObjectMenuPrefix = '&Object'
        LinesRTF.Strings = (
          '{\rtf1\ansi\deff0\nouicompat{\fonttbl{\f0\fnil Arial;}}'
          '{\*\generator Riched20 10.0.14393}\viewkind4\uc1 '
          '\pard\f0\fs18\lang1036 '
          '\par '
          '\par '
          '\par '
          '\par }')
      end
      object POPTEDIT: TPanel
        Left = 4
        Top = 47
        Width = 488
        Height = 315
        BevelOuter = bvNone
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 36
        Visible = False
        object GB_OPTEDIT: TGroupBox
          Left = 4
          Top = -3
          Width = 481
          Height = 124
          Caption = 'Edition des ...'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -16
          Font.Name = 'Cambria'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          object Label44: TLabel
            Left = 26
            Top = 30
            Width = 107
            Height = 13
            Caption = 'Nombre d'#39'exemplaire : '
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object Label47: TLabel
            Left = 10
            Top = 64
            Width = 124
            Height = 19
            Caption = 'Edition Document'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBackground
            Font.Height = -16
            Font.Name = 'Cambria'
            Font.Style = []
            ParentFont = False
          end
          object Label48: TLabel
            Left = 10
            Top = 125
            Width = 161
            Height = 19
            Caption = 'Edition au format Etat...'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBackground
            Font.Height = -16
            Font.Name = 'Cambria'
            Font.Style = []
            ParentFont = False
            Visible = False
          end
          object Label49: TLabel
            Left = 26
            Top = 92
            Width = 102
            Height = 13
            Caption = 'Mod'#232'le d'#39'Impression :'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object Label50: TLabel
            Left = 26
            Top = 156
            Width = 96
            Height = 13
            Caption = 'Mod'#232'le Disponible : '
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
            Visible = False
          end
          object Label51: TLabel
            Left = 26
            Top = 180
            Width = 102
            Height = 13
            Caption = 'Mod'#232'le d'#39'Impression :'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
            Visible = False
          end
          object GPP_NBEXEMPLAIRE: TSpinEdit
            Left = 137
            Top = 26
            Width = 49
            Height = 22
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBackground
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            MaxValue = 0
            MinValue = 0
            ParentFont = False
            TabOrder = 0
            Value = 0
          end
          object GPP_IMPIMMEDIATE: TCheckBox
            Left = 312
            Top = 24
            Width = 153
            Height = 17
            Alignment = taLeftJustify
            Caption = 'Impression Imm'#233'diate'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
            TabOrder = 1
          end
          object GPP_APERCUAVIMP: TCheckBox
            Left = 312
            Top = 40
            Width = 153
            Height = 17
            Alignment = taLeftJustify
            Caption = 'Aper'#231'u avant Impression'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
            TabOrder = 2
          end
          object GPP_IMPMODELE: THValComboBox
            Left = 165
            Top = 89
            Width = 249
            Height = 21
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBackground
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ItemHeight = 0
            ParentFont = False
            TabOrder = 3
            TagDispatch = 0
            DataType = 'GCIMPETAT'
          end
          object GPP_IMPETAT: THValComboBox
            Left = 141
            Top = 177
            Width = 249
            Height = 21
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBackground
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ItemHeight = 0
            ParentFont = False
            TabOrder = 4
            Visible = False
            TagDispatch = 0
            DataType = 'GCIMPETAT'
          end
          object _ETATSDISPO: THMultiValComboBox
            Left = 141
            Top = 152
            Width = 325
            Height = 21
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBackground
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
            TabOrder = 5
            Abrege = False
            DataType = 'GCIMPETAT'
            Complete = False
            OuInclusif = False
          end
        end
        object GBFACTUREAVANC: TGroupBox
          Left = 4
          Top = 125
          Width = 481
          Height = 172
          Caption = 'Edition des situations'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBackground
          Font.Height = -16
          Font.Name = 'Cambria'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          Visible = False
          object Label63: TLabel
            Left = 10
            Top = 20
            Width = 93
            Height = 19
            Caption = 'Mode normal'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBackground
            Font.Height = -16
            Font.Name = 'Cambria'
            Font.Style = []
            ParentFont = False
          end
          object Label65: TLabel
            Left = 26
            Top = 44
            Width = 92
            Height = 13
            Caption = 'Mod'#232'le de situation'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object Label62: TLabel
            Left = 26
            Top = 116
            Width = 92
            Height = 13
            Caption = 'Mod'#232'le de situation'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object Label64: TLabel
            Left = 26
            Top = 68
            Width = 79
            Height = 13
            Caption = 'Etat r'#233'captitulatif'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object Label66: TLabel
            Left = 10
            Top = 92
            Width = 91
            Height = 19
            Caption = 'Mode directe'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBackground
            Font.Height = -16
            Font.Name = 'Cambria'
            Font.Style = []
            ParentFont = False
          end
          object Label67: TLabel
            Left = 26
            Top = 139
            Width = 79
            Height = 13
            Caption = 'Etat r'#233'captitulatif'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object SO_BTETATSIT: THValComboBox
            Left = 165
            Top = 41
            Width = 249
            Height = 21
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBackground
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ItemHeight = 0
            ParentFont = False
            TabOrder = 0
            TagDispatch = 0
            DataType = 'GCIMPETAT'
          end
          object SO_BTETATSITDIR: THValComboBox
            Left = 165
            Top = 113
            Width = 249
            Height = 21
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBackground
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ItemHeight = 0
            ParentFont = False
            TabOrder = 1
            TagDispatch = 0
            DataType = 'GCIMPETAT'
          end
          object SO_BTETATSIR: THValComboBox
            Left = 165
            Top = 65
            Width = 249
            Height = 21
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBackground
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ItemHeight = 0
            ParentFont = False
            TabOrder = 2
            TagDispatch = 0
            DataType = 'GCIMPETAT'
          end
          object SO_BTETATSIRDIR: THValComboBox
            Left = 165
            Top = 136
            Width = 249
            Height = 21
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBackground
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ItemHeight = 0
            ParentFont = False
            TabOrder = 3
            TagDispatch = 0
            DataType = 'GCIMPETAT'
          end
        end
      end
      object BPD_IMPBASPAGE: TEdit
        Left = 208
        Top = 415
        Width = 25
        Height = 21
        BorderStyle = bsNone
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 37
        Visible = False
      end
    end
  end
  inherited PanelImage: THPanel
    Left = 7
    Top = 114
    Width = 224
    Height = 271
    inherited Image: TToolbarButton97
      Left = 8
      Top = 10
      Width = 209
      Height = 255
    end
  end
  inherited cControls: THListBox
    Left = 29
    Top = 81
    Width = 65
    Height = 24
    BorderStyle = bsNone
    Font.Color = clMaroon
    ParentFont = False
  end
  object TFINISASSISTANT: TCheckBox [14]
    Left = 674
    Top = 542
    Width = 109
    Height = 17
    Alignment = taLeftJustify
    Anchors = [akRight, akBottom]
    Caption = 'Ne plus afficher'
    TabOrder = 5
    OnClick = TFINISASSISTANTClick
  end
  object PTitre: TPanel [15]
    Left = 0
    Top = 0
    Width = 791
    Height = 41
    Align = alTop
    BevelInner = bvLowered
    BevelOuter = bvLowered
    Caption = 'Assistant de Param'#233'trage Soci'#233't'#233
    Color = clHighlight
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clHighlightText
    Font.Height = -30
    Font.Name = 'Cambria'
    Font.Style = []
    ParentFont = False
    TabOrder = 6
  end
  object SO_BTGAMME: TEdit [16]
    Left = 107
    Top = 81
    Width = 17
    Height = 24
    BorderStyle = bsNone
    Color = clYellow
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clMaroon
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    MaxLength = 9
    ParentFont = False
    TabOrder = 7
    Visible = False
  end
  inherited Msg: THMsgBox
    Left = 191
    Top = 41
  end
  inherited HMTrad: THSystemMenu
    Left = 156
    Top = 41
  end
  object POPCPTA: TPopupMenu
    Left = 24
    Top = 40
    object InfoGen: TMenuItem
      Caption = 'Infos g'#233'n'#233'rale'
      OnClick = InfoGenClick
    end
    object Fourchettesdecomptes1: TMenuItem
      Caption = 'Fourchettes de comptes'
    end
    object Comptespciaux1: TMenuItem
      Caption = 'Dates'
    end
    object GestionTVA1: TMenuItem
      Caption = 'Gestion TVA'
    end
    object Monnaiedetenue1: TMenuItem
      Caption = 'Monnaie de tenue'
    end
    object Divers1: TMenuItem
      Caption = 'Divers'
    end
    object Paramtragedelettrage1: TMenuItem
      Caption = 'Param'#233'trage de lettrage'
    end
    object Modedactionsurlettrage1: TMenuItem
      Caption = 'Mode d'#39'action sur lettrage'
    end
  end
  object ODGetInfosTRA: TOpenDialog
    Filter = 'Fichier transfert PGI|*.TRA'
    Options = [ofHideReadOnly, ofPathMustExist, ofFileMustExist, ofEnableSizing]
    Left = 123
    Top = 40
  end
  object POPBTP: TPopupMenu
    Left = 64
    Top = 40
    object InfoClient: TMenuItem
      Caption = 'Familles Comptables Clients'
      OnClick = InfoclientClick
    end
    object InfoAffaires: TMenuItem
      Caption = 'Familles Comptables Chantiers'
      OnClick = InfoAffairesClick
    end
    object FamCptaArt: TMenuItem
      Caption = 'Familles Comptables Articles'
      OnClick = FamCptaArtClick
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object VentilCpta1: TMenuItem
      Caption = 'Ventilation Comptable'
      OnClick = VentilCpta1Click
    end
    object VentilTVA: TMenuItem
      Caption = 'Taux et Ventilation TVA'
      OnClick = VentilTVAClick
    end
    object N2: TMenuItem
      Caption = '-'
    end
    object LibFamArt: TMenuItem
      Caption = 'Libell'#233's des Familles Articles'
      OnClick = LibFamArtClick
    end
    object MESURE: TMenuItem
      Caption = 'Unit'#233's de Mesure'
      OnClick = MESUREClick
    end
    object ModedePaiement1: TMenuItem
      Caption = 'Mode de Paiement'
      OnClick = ModePaiementClick
    end
    object ConditionRegl: TMenuItem
      Caption = 'Conditions de R'#233'glement'
      OnClick = ConditionReglClick
    end
    object FormesJuridiques: TMenuItem
      Caption = 'Formes Juridiques'
      OnClick = FormesJuridiquesClick
    end
  end
end

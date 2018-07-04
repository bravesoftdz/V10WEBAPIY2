inherited FExportCais: TFExportCais
  Left = 212
  Top = 161
  Caption = 'Constitution du fichier de démarrage d'#39'une caisse'
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Left = 3
    Top = 312
    Visible = False
  end
  object lEtapeManu: THLabel [1]
    Left = 3
    Top = 312
    Width = 65
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
  object bImprimer: TToolbarButton97 [2]
    Left = 199
    Top = 308
    Width = 65
    Height = 23
    Anchors = [akRight, akBottom]
    Caption = 'Imprimer'
    Flat = False
    Visible = False
    OnClick = bImprimerClick
  end
  object bParamPce: TToolbarButton97 [3]
    Left = 133
    Top = 308
    Width = 65
    Height = 23
    Anchors = [akRight, akBottom]
    Caption = 'Param Pce'
    Flat = False
    Visible = False
    OnClick = bParamPceClick
  end
  inherited bPrecedent: TToolbarButton97
    Left = 332
    Width = 65
  end
  inherited bSuivant: TToolbarButton97
    Left = 399
    Width = 65
  end
  inherited bFin: TToolbarButton97
    Left = 466
    Width = 65
  end
  inherited bAnnuler: TToolbarButton97
    Left = 265
    Width = 65
  end
  inherited bAide: TToolbarButton97
    Left = 66
    Width = 65
  end
  inherited P: TPageControl
    Left = 143
    Top = 4
    Width = 383
    Height = 281
    ActivePage = PAGE22
    TabOrder = 4
    object PAGE1: TTabSheet
      Caption = 'PAGE1'
      object LGPK_CAISSE: THLabel
        Left = 9
        Top = 120
        Width = 85
        Height = 13
        Caption = 'Choix de la caisse'
        FocusControl = GPK_CAISSE
      end
      object LSO_LIBELLE: THLabel
        Left = 9
        Top = 152
        Width = 103
        Height = 13
        Caption = 'Enseigne de la caisse'
        FocusControl = SO_LIBELLE
      end
      object LIB_CAISSE: THLabel
        Left = 172
        Top = 120
        Width = 60
        Height = 13
        Caption = 'LIB_CAISSE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object PTITRE1: THPanel
        Left = 16
        Top = 8
        Width = 313
        Height = 89
        FullRepaint = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        object TITRE1: THLabel
          Left = 16
          Top = 12
          Width = 273
          Height = 61
          Alignment = taCenter
          AutoSize = False
          Caption = 
            'Cet assistant va constituer le fichier de démarrage d'#39'une caisse' +
            '.'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
      end
      object GPK_CAISSE: THCritMaskEdit
        Left = 116
        Top = 116
        Width = 45
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 1
        TagDispatch = 0
        DataType = 'GCPCAISSE'
        ElipsisButton = True
        Libelle = LIB_CAISSE
      end
      object SO_LIBELLE: TEdit
        Left = 116
        Top = 148
        Width = 218
        Height = 21
        MaxLength = 35
        TabOrder = 2
        OnEnter = SO_LIBELLEEnter
      end
    end
    object PAGE21: TTabSheet
      Caption = 'PAGE21'
      ImageIndex = 8
      object LREPRESENTANT: THLabel
        Left = 9
        Top = 167
        Width = 92
        Height = 13
        Caption = 'Choix de l'#39'utilisateur'
        FocusControl = US_UTILISATEUR_
      end
      object LNOMREP: THLabel
        Left = 9
        Top = 192
        Width = 88
        Height = 13
        Caption = 'Nom de l'#39'utilisateur'
        FocusControl = US_LIBELLE
      end
      object TDBCENTRAL: TLabel
        Left = 9
        Top = 120
        Width = 65
        Height = 13
        Caption = 'Base centrale'
      end
      object LCONFIGPCP: THLabel
        Left = 9
        Top = 218
        Width = 62
        Height = 13
        Caption = 'Configuration'
        FocusControl = CONFIGPCP
      end
      object HPanel1: THPanel
        Left = 16
        Top = 8
        Width = 313
        Height = 89
        FullRepaint = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        object TITREPOP: THLabel
          Left = 26
          Top = 24
          Width = 257
          Height = 49
          Alignment = taCenter
          AutoSize = False
          Caption = 'Sélection de la base du site central et de l'#39'utilisateur'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
      end
      object US_UTILISATEUR_: THCritMaskEdit
        Left = 119
        Top = 163
        Width = 215
        Height = 21
        CharCase = ecUpperCase
        Enabled = False
        TabOrder = 2
        OnChange = US_UTILISATEUR_Change
        TagDispatch = 0
        DataType = 'YYUTILISATMASTER'
        ElipsisButton = True
      end
      object US_LIBELLE: TEdit
        Left = 119
        Top = 188
        Width = 215
        Height = 21
        TabStop = False
        Enabled = False
        MaxLength = 35
        TabOrder = 3
        OnEnter = SO_LIBELLEEnter
      end
      object DBCENTRAL: TComboBox
        Left = 119
        Top = 116
        Width = 215
        Height = 21
        ItemHeight = 0
        TabOrder = 1
        OnChange = DBCENTRAL_OnChange
      end
      object CONFIGPCP: TEdit
        Left = 119
        Top = 214
        Width = 215
        Height = 21
        TabStop = False
        Enabled = False
        MaxLength = 35
        TabOrder = 4
        OnEnter = SO_LIBELLEEnter
      end
    end
    object PAGE211: TTabSheet
      Caption = 'PAGE211'
      ImageIndex = 10
      object LSO_MAIL: THLabel
        Left = 9
        Top = 226
        Width = 28
        Height = 13
        BiDiMode = bdLeftToRight
        Caption = 'E-mail'
        Color = clBtnFace
        FocusControl = SO_MAIL_
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentBiDiMode = False
        ParentColor = False
        ParentFont = False
      end
      object LSO_ADRESSE: THLabel
        Left = 9
        Top = 119
        Width = 19
        Height = 13
        BiDiMode = bdLeftToRight
        Caption = 'Adr.'
        Color = clBtnFace
        FocusControl = SO_ADRESSE1_
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentBiDiMode = False
        ParentColor = False
        ParentFont = False
      end
      object LSO_CODEPOSTAL: THLabel
        Left = 166
        Top = 119
        Width = 20
        Height = 13
        BiDiMode = bdLeftToRight
        Caption = 'C.P.'
        Color = clBtnFace
        FocusControl = SO_CODEPOSTAL_
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentBiDiMode = False
        ParentColor = False
        ParentFont = False
      end
      object LSO_DIVTERRIT: THLabel
        Left = 244
        Top = 119
        Width = 41
        Height = 13
        BiDiMode = bdLeftToRight
        Caption = 'Div.Terr.'
        Color = clBtnFace
        FocusControl = SO_DIVTERRIT_
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentBiDiMode = False
        ParentColor = False
        ParentFont = False
      end
      object LSO_VILLE: THLabel
        Left = 166
        Top = 144
        Width = 19
        Height = 13
        BiDiMode = bdLeftToRight
        Caption = 'Ville'
        Color = clBtnFace
        FocusControl = SO_VILLE_
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentBiDiMode = False
        ParentColor = False
        ParentFont = False
      end
      object LSO_PAYS: THLabel
        Left = 166
        Top = 172
        Width = 23
        Height = 13
        BiDiMode = bdLeftToRight
        Caption = 'Pays'
        Color = clBtnFace
        FocusControl = SO_PAYS_
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentBiDiMode = False
        ParentColor = False
        ParentFont = False
      end
      object LSO_FAX: THLabel
        Left = 166
        Top = 200
        Width = 17
        Height = 13
        BiDiMode = bdLeftToRight
        Caption = 'Fax'
        Color = clBtnFace
        FocusControl = SO_FAX_
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentBiDiMode = False
        ParentColor = False
        ParentFont = False
      end
      object LSO_TELEPHONE: THLabel
        Left = 9
        Top = 199
        Width = 18
        Height = 13
        BiDiMode = bdLeftToRight
        Caption = 'Tél.'
        Color = clBtnFace
        FocusControl = SO_TELEPHONE_
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentBiDiMode = False
        ParentColor = False
        ParentFont = False
      end
      object HLabel2: THLabel
        Left = 9
        Top = 93
        Width = 22
        Height = 13
        BiDiMode = bdLeftToRight
        Caption = 'Nom'
        Color = clBtnFace
        FocusControl = SO_MAIL_
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentBiDiMode = False
        ParentColor = False
        ParentFont = False
      end
      object HPanel2: THPanel
        Left = 16
        Top = 8
        Width = 313
        Height = 66
        FullRepaint = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        object HLabel1: THLabel
          Left = 14
          Top = 24
          Width = 273
          Height = 20
          Alignment = taCenter
          AutoSize = False
          Caption = 'Coordonnées'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
      end
      object SO_LIBELLE_: TEdit
        Left = 40
        Top = 89
        Width = 124
        Height = 21
        BiDiMode = bdLeftToRight
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentBiDiMode = False
        ParentFont = False
        TabOrder = 1
        OnEnter = SO_LIBELLEEnter
      end
      object SO_MAIL_: TEdit
        Left = 40
        Top = 222
        Width = 291
        Height = 21
        BiDiMode = bdLeftToRight
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentBiDiMode = False
        ParentFont = False
        TabOrder = 2
        OnEnter = SO_LIBELLEEnter
      end
      object SO_ADRESSE1_: TEdit
        Left = 40
        Top = 114
        Width = 124
        Height = 21
        BiDiMode = bdLeftToRight
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentBiDiMode = False
        ParentFont = False
        TabOrder = 3
        OnEnter = SO_LIBELLEEnter
      end
      object SO_CODEPOSTAL_: TEdit
        Left = 189
        Top = 114
        Width = 55
        Height = 21
        BiDiMode = bdLeftToRight
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentBiDiMode = False
        ParentFont = False
        TabOrder = 4
        OnEnter = SO_LIBELLEEnter
      end
      object SO_DIVTERRIT_: TEdit
        Left = 287
        Top = 114
        Width = 43
        Height = 21
        BiDiMode = bdLeftToRight
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentBiDiMode = False
        ParentFont = False
        TabOrder = 5
        OnEnter = SO_LIBELLEEnter
      end
      object SO_VILLE_: TEdit
        Left = 189
        Top = 141
        Width = 144
        Height = 21
        BiDiMode = bdLeftToRight
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentBiDiMode = False
        ParentFont = False
        TabOrder = 6
        OnEnter = SO_LIBELLEEnter
      end
      object SO_ADRESSE2_: TEdit
        Left = 40
        Top = 141
        Width = 124
        Height = 21
        BiDiMode = bdLeftToRight
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentBiDiMode = False
        ParentFont = False
        TabOrder = 7
        OnEnter = SO_LIBELLEEnter
      end
      object SO_ADRESSE3_: TEdit
        Left = 40
        Top = 168
        Width = 124
        Height = 21
        BiDiMode = bdLeftToRight
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentBiDiMode = False
        ParentFont = False
        TabOrder = 8
        OnEnter = SO_LIBELLEEnter
      end
      object SO_PAYS_: THValComboBox
        Left = 189
        Top = 169
        Width = 144
        Height = 21
        BiDiMode = bdLeftToRight
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 0
        ParentBiDiMode = False
        ParentFont = False
        TabOrder = 9
        OnChange = SO_DEVISEPRINCChange
        TagDispatch = 0
        Vide = True
        VideString = '<<Aucun>>'
        DataType = 'TTPAYS'
      end
      object SO_FAX_: TEdit
        Left = 189
        Top = 194
        Width = 141
        Height = 21
        BiDiMode = bdLeftToRight
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentBiDiMode = False
        ParentFont = False
        TabOrder = 10
        OnEnter = SO_LIBELLEEnter
      end
      object SO_TELEPHONE_: TEdit
        Left = 40
        Top = 193
        Width = 124
        Height = 21
        BiDiMode = bdLeftToRight
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentBiDiMode = False
        ParentFont = False
        TabOrder = 11
        OnEnter = SO_LIBELLEEnter
      end
    end
    object PAGE212: TTabSheet
      Caption = 'PAGE212'
      ImageIndex = 11
      object GP_CODEIFART: TGroupBox
        Left = 7
        Top = 133
        Width = 326
        Height = 46
        Caption = 'Codification Article'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 0
        object LSO_GCLGNUMART_: THLabel
          Left = 191
          Top = 21
          Width = 74
          Height = 13
          BiDiMode = bdLeftToRight
          Caption = 'Longueur totale'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentBiDiMode = False
          ParentColor = False
          ParentFont = False
        end
        object LSO_GCPREFIXEART_: THLabel
          Left = 7
          Top = 21
          Width = 32
          Height = 13
          BiDiMode = bdLeftToRight
          Caption = 'Préfixe'
          Color = clBtnFace
          FocusControl = SO_GCPREFIXEART_
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentBiDiMode = False
          ParentColor = False
          ParentFont = False
        end
        object LSO_GCCOMPTEURART_: THLabel
          Left = 101
          Top = 21
          Width = 34
          Height = 13
          BiDiMode = bdLeftToRight
          Caption = 'Chrono'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentBiDiMode = False
          ParentColor = False
          ParentFont = False
        end
        object SO_GCLGNUMART_: THDBSpinEdit
          Left = 267
          Top = 18
          Width = 49
          Height = 22
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxValue = 999999999
          MinValue = 0
          ParentFont = False
          TabOrder = 2
          Value = 0
        end
        object SO_GCPREFIXEART_: TEdit
          Left = 43
          Top = 18
          Width = 46
          Height = 21
          BiDiMode = bdLeftToRight
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 35
          ParentBiDiMode = False
          ParentFont = False
          TabOrder = 0
          OnEnter = SO_LIBELLEEnter
        end
        object SO_GCCOMPTEURART_: THNumEdit
          Left = 138
          Top = 18
          Width = 49
          Height = 21
          Decimals = 2
          Digits = 12
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clGray
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Masks.PositiveMask = '0'
          Debit = False
          ParentFont = False
          TabOrder = 1
          UseRounding = True
          Validate = False
        end
      end
      object HPanel3: THPanel
        Left = 16
        Top = 6
        Width = 313
        Height = 31
        FullRepaint = False
        TabOrder = 1
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        object HLabel3: THLabel
          Left = 14
          Top = 5
          Width = 273
          Height = 20
          Alignment = taCenter
          AutoSize = False
          Caption = 'Paramètres et Codification'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
      end
      object GP_CODIFTIERS: TGroupBox
        Left = 7
        Top = 180
        Width = 326
        Height = 67
        Caption = 'Codification Tiers'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 2
        object LSO_GCLGNUMTIERS_: THLabel
          Left = 191
          Top = 19
          Width = 74
          Height = 13
          BiDiMode = bdLeftToRight
          Caption = 'Longueur totale'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentBiDiMode = False
          ParentColor = False
          ParentFont = False
          Visible = False
        end
        object LSO_GCPREFIXETIERS_: THLabel
          Left = 7
          Top = 19
          Width = 32
          Height = 13
          BiDiMode = bdLeftToRight
          Caption = 'Préfixe'
          Color = clBtnFace
          FocusControl = SO_GCPREFIXETIERS_
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentBiDiMode = False
          ParentColor = False
          ParentFont = False
        end
        object LSO_GCCOMPTEURTIERS_: THLabel
          Left = 101
          Top = 19
          Width = 34
          Height = 13
          BiDiMode = bdLeftToRight
          Caption = 'Chrono'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentBiDiMode = False
          ParentColor = False
          ParentFont = False
        end
        object LSO_GCPREFIXEAUXI_: THLabel
          Left = 7
          Top = 45
          Width = 103
          Height = 13
          BiDiMode = bdLeftToRight
          Caption = 'Préfixe auxiliaire client'
          Color = clBtnFace
          FocusControl = SO_GCPREFIXETIERS_
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentBiDiMode = False
          ParentColor = False
          ParentFont = False
        end
        object LSO_GCPREFIXEAUXIFOU_: THLabel
          Left = 191
          Top = 43
          Width = 51
          Height = 13
          BiDiMode = bdLeftToRight
          Caption = 'fournisseur'
          Color = clBtnFace
          FocusControl = SO_GCPREFIXETIERS_
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentBiDiMode = False
          ParentColor = False
          ParentFont = False
        end
        object SO_GCLGNUMTIERS_: THDBSpinEdit
          Left = 267
          Top = 16
          Width = 49
          Height = 22
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxValue = 999999999
          MinValue = 0
          ParentFont = False
          TabOrder = 2
          Value = 0
          Visible = False
        end
        object SO_GCPREFIXETIERS_: TEdit
          Left = 43
          Top = 16
          Width = 46
          Height = 21
          BiDiMode = bdLeftToRight
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 35
          ParentBiDiMode = False
          ParentFont = False
          TabOrder = 0
          OnEnter = SO_LIBELLEEnter
        end
        object SO_GCCOMPTEURTIERS_: THNumEdit
          Left = 138
          Top = 16
          Width = 49
          Height = 21
          Decimals = 2
          Digits = 12
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clGray
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Masks.PositiveMask = '0'
          Debit = False
          ParentFont = False
          TabOrder = 1
          UseRounding = True
          Validate = False
        end
        object SO_GCPREFIXEAUXI_: TEdit
          Left = 138
          Top = 40
          Width = 50
          Height = 21
          BiDiMode = bdLeftToRight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 35
          ParentBiDiMode = False
          ParentFont = False
          TabOrder = 3
          OnEnter = SO_LIBELLEEnter
        end
        object SO_GCPREFIXEAUXIFOU_: TEdit
          Left = 267
          Top = 40
          Width = 50
          Height = 21
          BiDiMode = bdLeftToRight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 35
          ParentBiDiMode = False
          ParentFont = False
          TabOrder = 4
          OnEnter = SO_LIBELLEEnter
        end
      end
      object GP_PARAMPCES: TGroupBox
        Left = 7
        Top = 40
        Width = 326
        Height = 90
        Caption = 'Paramètre Pièces'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 3
        object GP_VERROUILLEPCES_: TCheckBox
          Left = 7
          Top = 15
          Width = 266
          Height = 17
          Caption = 'Verrouiller les pièces envoyées'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 0
        end
        object GP_IMPIMMEDIATE_: TCheckBox
          Left = 7
          Top = 33
          Width = 130
          Height = 17
          Caption = 'Imprimer à la validation'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 1
          OnClick = GP_IMPIMMEDIATE_Click
        end
        object GP_APERCUAVIMP_: TCheckBox
          Left = 7
          Top = 49
          Width = 146
          Height = 17
          Caption = 'Aperçu avant impression'
          Color = clBtnFace
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 2
        end
        object GP_PARAMPCES_: TCheckBox
          Left = 7
          Top = 69
          Width = 306
          Height = 17
          Caption = 'Détailler le paramétrage des pièces en fin de traitement'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 3
        end
      end
    end
    object PAGE9: TTabSheet
      Caption = 'PAGE9'
      ImageIndex = 8
      object LSO_DEVISEPRINC: THLabel
        Left = 18
        Top = 130
        Width = 78
        Height = 13
        Caption = 'Devise de tenue'
        FocusControl = SO_DEVISEPRINC
      end
      object PTITRE9: THPanel
        Left = 16
        Top = 8
        Width = 313
        Height = 105
        FullRepaint = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        object TITRE9: THLabel
          Left = 16
          Top = 12
          Width = 273
          Height = 37
          Alignment = taCenter
          AutoSize = False
          Caption = 'Indiquez la devise de tenue de la caisse'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
        object TITRE91: THLabel
          Left = 9
          Top = 56
          Width = 296
          Height = 42
          Alignment = taCenter
          AutoSize = False
          Caption = 
            'Si la caisse est gérée en Euro, mentionnez la devise de contre-v' +
            'aleur et activez l'#39'indicateur de tenue en Euro.'
          WordWrap = True
        end
      end
      object SO_DEVISEPRINC: THValComboBox
        Left = 168
        Top = 126
        Width = 145
        Height = 21
        ItemHeight = 0
        TabOrder = 1
        OnChange = SO_DEVISEPRINCChange
        TagDispatch = 0
        DataType = 'TTDEVISE'
      end
      object SO_TENUEEURO: TCheckBox
        Left = 16
        Top = 160
        Width = 165
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Caisse tenue en Euro'
        TabOrder = 2
      end
    end
    object PAGE2: TTabSheet
      Caption = 'PAGE2'
      ImageIndex = 3
      object LSO_REGIMEDEFAUT: THLabel
        Left = 18
        Top = 92
        Width = 104
        Height = 13
        Caption = 'Régime de facturation'
        FocusControl = SO_REGIMEDEFAUT
      end
      object LSO_GCMODEREGLEDEFAUT: THLabel
        Left = 18
        Top = 116
        Width = 113
        Height = 13
        Caption = 'Mode de règlement tiers'
        FocusControl = SO_GCMODEREGLEDEFAUT
      end
      object LSO_GCTIERSDEFAUT: THLabel
        Left = 18
        Top = 139
        Width = 89
        Height = 13
        Caption = 'Tiers des transferts'
        FocusControl = SO_GCTIERSDEFAUT
      end
      object LSO_GCPERBASETARIF: THLabel
        Left = 18
        Top = 162
        Width = 142
        Height = 13
        Caption = 'Période de base pour les tarifs'
        FocusControl = SO_GCPERBASETARIF
      end
      object LSO_GCPHOTOFICHE: THLabel
        Left = 18
        Top = 209
        Width = 99
        Height = 13
        Caption = 'Photo pour les fiches'
        FocusControl = SO_GCPHOTOFICHE
      end
      object LUS_UTILISATEUR: THLabel
        Left = 18
        Top = 186
        Width = 99
        Height = 13
        Caption = 'Utilisateurs autorisés '
        FocusControl = US_UTILISATEUR
      end
      object PTITRE2: THPanel
        Left = 16
        Top = 8
        Width = 313
        Height = 66
        FullRepaint = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        object TITRE2: THLabel
          Left = 16
          Top = 12
          Width = 273
          Height = 38
          Alignment = taCenter
          AutoSize = False
          Caption = 'Indiquez les paramètres de la caisse'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
      end
      object SO_REGIMEDEFAUT: THValComboBox
        Left = 168
        Top = 88
        Width = 145
        Height = 21
        ItemHeight = 0
        TabOrder = 1
        TagDispatch = 0
        DataType = 'TTREGIMETVA'
      end
      object SO_GCMODEREGLEDEFAUT: THValComboBox
        Left = 168
        Top = 112
        Width = 145
        Height = 21
        ItemHeight = 0
        TabOrder = 2
        TagDispatch = 0
        DataType = 'TTMODEREGLE'
      end
      object SO_GCTIERSDEFAUT: THCritMaskEdit
        Left = 168
        Top = 135
        Width = 145
        Height = 21
        TabOrder = 3
        TagDispatch = 0
        DataType = 'GCTIERS'
        ElipsisButton = True
      end
      object SO_GCPERBASETARIF: THValComboBox
        Left = 168
        Top = 158
        Width = 145
        Height = 21
        ItemHeight = 0
        TabOrder = 4
        TagDispatch = 0
        DataType = 'GCTARIFPERBASE'
      end
      object SO_GCPHOTOFICHE: THValComboBox
        Left = 168
        Top = 205
        Width = 145
        Height = 21
        ItemHeight = 0
        TabOrder = 6
        TagDispatch = 0
        DataType = 'GCEMPLOIBLOB'
      end
      object US_UTILISATEUR: THMultiValComboBox
        Left = 168
        Top = 182
        Width = 145
        Height = 21
        TabOrder = 5
        Text = '<<Tous>>'
        Abrege = False
        DataType = 'TTUTILISATEUR'
        Complete = False
        OuInclusif = False
      end
    end
    object PAGE3: TTabSheet
      Caption = 'PAGE3'
      ImageIndex = 1
      object LSITECAISSE: THLabel
        Left = 22
        Top = 90
        Width = 119
        Height = 13
        Caption = 'Choix du site de la caisse'
        FocusControl = SITECAISSE
      end
      object LSITESIEGE: THLabel
        Left = 22
        Top = 212
        Width = 95
        Height = 13
        Caption = 'Choix du site central'
        FocusControl = SITESIEGE
      end
      object BSITECAISSE: TToolbarButton97
        Left = 276
        Top = 86
        Width = 23
        Height = 21
        Hint = 'Création du site de la caisse'
        Flat = False
        Glyph.Data = {
          42010000424D4201000000000000760000002800000011000000110000000100
          040000000000CC000000CE0E0000D80E00001000000000000000000000000000
          BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          77777000000070000000000000077000000070FFF0FFFFFF0F077000000070F0
          F0F0FF0F0F077000000070000000000000077000000077709999999077777000
          0000777090000907777770000000777090709007777770000000777090099700
          77777000000077709099070007777000000077709990770BB077700000007770
          9907770BB07770000000777090777770BB0770000000777007777770B0077000
          00007770777777770BB070000000777777777777000770000000777777777777
          777770000000}
        ParentShowHint = False
        ShowHint = True
        OnClick = BSITECAISSEClick
      end
      object BSITESIEGE: TToolbarButton97
        Left = 276
        Top = 208
        Width = 23
        Height = 21
        Hint = 'Création du site central'
        Flat = False
        Glyph.Data = {
          42010000424D4201000000000000760000002800000011000000110000000100
          040000000000CC000000CE0E0000D80E00001000000000000000000000000000
          BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          77777000000070000000000000077000000070FFF0FFFFFF0F077000000070F0
          F0F0FF0F0F077000000070000000000000077000000077709999999077777000
          0000777090000907777770000000777090709007777770000000777090099700
          77777000000077709099070007777000000077709990770BB077700000007770
          9907770BB07770000000777090777770BB0770000000777007777770B0077000
          00007770777777770BB070000000777777777777000770000000777777777777
          777770000000}
        ParentShowHint = False
        ShowHint = True
        OnClick = BSITESIEGEClick
      end
      object SITECAISSE: THCritMaskEdit
        Left = 151
        Top = 86
        Width = 121
        Height = 21
        TabOrder = 1
        OnDblClick = SITECAISSEDblClick
        TagDispatch = 0
        DataType = 'STOXSITES'
        ElipsisButton = True
      end
      object PTITRE3: THPanel
        Left = 16
        Top = 8
        Width = 313
        Height = 72
        FullRepaint = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        object TITRE3: THLabel
          Left = 16
          Top = 12
          Width = 273
          Height = 45
          Alignment = taCenter
          AutoSize = False
          Caption = 'Indiquez le site qui correspond à la caisse.'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
      end
      object PTITRE31: THPanel
        Left = 16
        Top = 127
        Width = 313
        Height = 72
        FullRepaint = False
        TabOrder = 2
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        object TITRE31: THLabel
          Left = 16
          Top = 12
          Width = 273
          Height = 45
          Alignment = taCenter
          AutoSize = False
          Caption = 'Indiquez le site qui correspond au site central.'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
      end
      object SITESIEGE: THCritMaskEdit
        Left = 151
        Top = 208
        Width = 121
        Height = 21
        TabOrder = 3
        OnDblClick = SITESIEGEDblClick
        TagDispatch = 0
        DataType = 'STOXSITES'
        ElipsisButton = True
      end
    end
    object PAGE22: TTabSheet
      Caption = 'PAGE22'
      ImageIndex = 9
      object PTITRE22: THPanel
        Left = 16
        Top = 8
        Width = 313
        Height = 66
        FullRepaint = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        object TITR22: THLabel
          Left = 16
          Top = 12
          Width = 273
          Height = 38
          Alignment = taCenter
          AutoSize = False
          Caption = 'Indiquez l'#39'événement d'#39'initialisation du poste utilisateur'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
      end
      object GEVTINIT: THGrid
        Left = 16
        Top = 84
        Width = 313
        Height = 157
        ColCount = 2
        FixedCols = 0
        RowCount = 6
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
        TabOrder = 1
        SortedCol = -1
        Titres.Strings = (
          'Code'
          'Description')
        Couleur = False
        MultiSelect = True
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = 13224395
        ColWidths = (
          64
          226)
      end
    end
    object PAGE8: TTabSheet
      Caption = 'PAGE8'
      ImageIndex = 7
      object PTITRE8: THPanel
        Left = 16
        Top = 8
        Width = 313
        Height = 66
        FullRepaint = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        object TITRE8: THLabel
          Left = 16
          Top = 12
          Width = 273
          Height = 38
          Alignment = taCenter
          AutoSize = False
          Caption = 'Indiquez les événements qui seront exécutés par la caisse'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
      end
      object GEVTS: THGrid
        Left = 16
        Top = 84
        Width = 313
        Height = 157
        ColCount = 2
        FixedCols = 0
        RowCount = 6
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
        TabOrder = 1
        SortedCol = -1
        Titres.Strings = (
          'Code'
          'Description')
        Couleur = False
        MultiSelect = True
        TitleBold = True
        TitleCenter = True
        ColCombo = 0
        SortEnabled = False
        SortRowExclude = 0
        TwoColors = False
        AlternateColor = 13224395
        ColWidths = (
          64
          226)
      end
    end
    object PAGE4: TTabSheet
      Caption = 'PAGE4'
      ImageIndex = 2
      object LNOMDOSSIER: THLabel
        Left = 24
        Top = 127
        Width = 73
        Height = 13
        Caption = 'Nom du dossier'
      end
      object PTITRE4: THPanel
        Left = 16
        Top = 8
        Width = 313
        Height = 106
        FullRepaint = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        object TITRE4: THLabel
          Left = 16
          Top = 12
          Width = 273
          Height = 41
          Alignment = taCenter
          AutoSize = False
          Caption = 
            'Indiquez le dossier de stockage du fichier de démarrage de la ca' +
            'isse'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
        object TITRE41: THLabel
          Left = 9
          Top = 55
          Width = 296
          Height = 42
          AutoSize = False
          Caption = 
            'Si le fichier doit être récupéré sur la caisse par télétransmiss' +
            'ion, vous devez indiquer le dossier par défaut du service FTP.'
          WordWrap = True
        end
      end
      object NOMDOSSIER: THCritMaskEdit
        Left = 104
        Top = 123
        Width = 217
        Height = 21
        TabOrder = 1
        TagDispatch = 0
        DataType = 'DIRECTORY'
        ElipsisButton = True
      end
    end
    object PAGE7: TTabSheet
      Caption = 'PAGE7'
      ImageIndex = 6
      object LSEV_CODEEVENT: THLabel
        Left = 16
        Top = 136
        Width = 168
        Height = 13
        Caption = 'Code de l'#39'événement de démarrage'
        FocusControl = SEV_CODEEVENT
      end
      object BCODEEVENT: TToolbarButton97
        Left = 312
        Top = 132
        Width = 23
        Height = 21
        Hint = 'Création de l'#39'événement de démarrage'
        Enabled = False
        Flat = False
        Glyph.Data = {
          42010000424D4201000000000000760000002800000011000000110000000100
          040000000000CC000000CE0E0000D80E00001000000000000000000000000000
          BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          77777000000070000000000000077000000070FFF0FFFFFF0F077000000070F0
          F0F0FF0F0F077000000070000000000000077000000077709999999077777000
          0000777090000907777770000000777090709007777770000000777090099700
          77777000000077709099070007777000000077709990770BB077700000007770
          9907770BB07770000000777090777770BB0770000000777007777770B0077000
          00007770777777770BB070000000777777777777000770000000777777777777
          777770000000}
        ParentShowHint = False
        ShowHint = True
        OnClick = BCODEEVENTClick
      end
      object PTITRE7: THPanel
        Left = 16
        Top = 8
        Width = 313
        Height = 89
        FullRepaint = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        object TITRE7: THLabel
          Left = 16
          Top = 12
          Width = 273
          Height = 61
          Alignment = taCenter
          AutoSize = False
          Caption = 
            'Vous pouvez préparer le fichier contenant toutes les données néc' +
            'essaires à la caisse'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
      end
      object LANCETOX: TCheckBox
        Left = 16
        Top = 108
        Width = 210
        Height = 17
        Caption = 'Démarrage des échanges FO - BO'
        TabOrder = 1
        OnClick = LANCETOXClick
      end
      object SEV_CODEEVENT: THCritMaskEdit
        Left = 191
        Top = 132
        Width = 119
        Height = 21
        Enabled = False
        MaxLength = 17
        TabOrder = 2
        OnDblClick = SEV_CODEEVENTDblClick
        TagDispatch = 0
        DataType = 'STOXEVENTS'
        ElipsisButton = True
      end
    end
    object PAGE5: TTabSheet
      Caption = 'PAGE5'
      ImageIndex = 4
      object PTITRE5: THPanel
        Left = 16
        Top = 8
        Width = 313
        Height = 89
        FullRepaint = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        object TITRE5: THLabel
          Left = 16
          Top = 12
          Width = 293
          Height = 41
          Alignment = taCenter
          AutoSize = False
          Caption = 
            'Le bouton '#39'Fin'#39' lance la construction du fichier de démarrage de' +
            ' la caisse :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
        object TITRE51: THLabel
          Left = 12
          Top = 64
          Width = 293
          Height = 13
          Alignment = taCenter
          AutoSize = False
          Caption = 'TITRE51'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold, fsItalic]
          ParentFont = False
        end
      end
      object CPTRENDU: THRichEditOLE
        Left = 16
        Top = 99
        Width = 313
        Height = 128
        TabStop = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        ScrollBars = ssBoth
        TabOrder = 1
        Margins.Top = 0
        Margins.Bottom = 0
        Margins.Left = 0
        Margins.Right = 0
        ContainerName = 'Document'
        ObjectMenuPrefix = '&Object'
        LinesRTF.Strings = (
          '{\rtf1\ansi\deff0{\fonttbl{\f0\fnil Arial;}}'
          '\viewkind4\uc1\pard\lang1036\f0\fs16 CPTRENDU'
          '\par }')
      end
      object PB: TProgressBar
        Left = 16
        Top = 229
        Width = 313
        Height = 16
        Min = 0
        Max = 100
        TabOrder = 2
      end
      object LstRecap: TListBox
        Left = 16
        Top = 99
        Width = 314
        Height = 129
        BiDiMode = bdLeftToRight
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Courier New'
        Font.Style = []
        ItemHeight = 14
        ParentBiDiMode = False
        ParentFont = False
        TabOrder = 3
      end
    end
  end
  inherited PanelImage: THPanel
    Width = 128
    TabOrder = 2
    inherited Image: TToolbarButton97
      Left = 8
      Width = 115
    end
  end
  inherited cControls: TListBox
    Left = 32
    TabOrder = 3
  end
  inherited Msg: THMsgBox
    Mess.Strings = (
      'Etape'
      '1;?caption?;Voulez-vous quitter l'#39'assistant ?;Q;YN;Y;C;'
      '2;?caption?;Le numéro de caisse est incorrect.;E;O;O;O;'
      '3;?caption?;Le site est incorrect.;E;O;O;O;'
      
        '4;?caption?;Le dossier de stockage du fichier de démarrage est i' +
        'ncorrect.;E;O;O;O;'
      '5;?caption?;L'#39'enseigne de la caisse est obligatoire.;E;O;O;O;'
      '6;?caption?;Le nom de la cobeille est incorrect.;E;O;O;O;'
      
        '7;?caption?;Le code événement de démarrage est incorrect.;E;O;O;' +
        'O;'
      '8;?caption?;Le code utilisateur est incorrect.;E;O;O;O;'
      '9;?caption?;L'#39'enseigne représentant est obligatoire.;E;O;O;O;'
      
        '10;?caption?;Cet utilisateur ne dispose pas de site associé.;E;O' +
        ';O;O;'
      
        '11;?caption?;Aucune pièces n'#39'est paramétrée pour la vente ou l'#39'a' +
        'chat;E;O;O;O;'
      
        '12;?caption?;Aucune pièces n'#39'est paramétrée pour la vente;E;O;O;' +
        'O;'
      
        '13;?caption?;Aucune pièces n'#39'est paramétrée pour l'#39'achat;E;O;O;O' +
        ';'
      '14;?caption?;;E;O;O;O;'
      
        '15;?caption?;Aucun représentant n'#39'est rattaché à cet utilisateur' +
        '.;E;O;O;O;'
      
        '16;?caption?;Aucun préfixe n'#39'est défini pour cet utilisateur.;E;' +
        'O;O;O;'
      '17;?caption?;;E;O;O;O;'
      '18;?caption?;;E;O;O;O;'
      
        '19;?caption?;Cet utilisateur n'#39'est pas paramétré pour la vente o' +
        'u pour l'#39'achat.;E;O;O;O;'
      
        '20;?caption?;Il faut sélectionner au moins un événement d'#39'initia' +
        'lisation.;E;O;O;O;'
      
        '21;?caption?;Il faut sélectionner au moins un événement de remon' +
        'tée.;E;O;O;O;'
      
        '22;?caption?;Le préfixe de l'#39'auxiliaire client est obligatoire.;' +
        'E;O;O;O;'
      
        '23;?caption?;Le préfixe de l'#39'auxiliaire fournisseur est obligato' +
        'ire.;E;O;O;O;'
      
        '24;?caption?;Les préfixes des auxiliaires client et fournisseur ' +
        'doivent être différents.;E;O;O;O;'
      '25;?caption?;La longueur du code tiers est obligatoire.;E;O;O;O;'
      
        '26;?caption?;Le "Contexte Métier" n'#39'est pas renseigné sur la bas' +
        'e centrale..;E;O;O;O;'
      ''
      '')
    Left = 63
    Top = 4
  end
  inherited HMTrad: THSystemMenu
    Top = 4
  end
end

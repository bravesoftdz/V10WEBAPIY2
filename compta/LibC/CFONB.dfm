object FEXPCFONB: TFEXPCFONB
  Left = 256
  Top = 193
  HelpContext = 999999823
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Export CFONB'
  ClientHeight = 428
  ClientWidth = 370
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object H_GENERAL: THLabel
    Left = 15
    Top = 149
    Width = 37
    Height = 13
    Caption = '&Banque'
    FocusControl = BQ_GENERAL
  end
  object H_NOMFICHIER: THLabel
    Left = 15
    Top = 177
    Width = 31
    Height = 13
    Caption = '&Fichier'
    FocusControl = NomFichier
  end
  object H_Document: THLabel
    Left = 15
    Top = 201
    Width = 49
    Height = 13
    Caption = '&Document'
    FocusControl = Document
  end
  object RechFile: TToolbarButton97
    Left = 344
    Top = 172
    Width = 16
    Height = 22
    Caption = '...'
    OnClick = RechFileClick
  end
  object TypeExport: TRadioGroup
    Left = 13
    Top = 6
    Width = 348
    Height = 131
    Caption = 'Format d'#39'export'
    Columns = 2
    Items.Strings = (
      'LCR / BOR encaissement'
      'LCR escompte standard'
      'LCR escompte en valeur'
      'LCR enc. apr'#232's expiration'
      'LCR escompte loi dailly'
      'Virements'
      'Virement '#224' '#233'ch'#233'ance E-2'
      'Virement '#224' '#233'ch'#233'ance E-3'
      'Pr'#233'l'#232'vements'
      'Transferts internationaux')
    TabOrder = 0
    OnClick = TypeExportClick
  end
  object Outils: TPanel
    Left = 0
    Top = 393
    Width = 370
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
    object Panel1: TPanel
      Left = 206
      Top = 2
      Width = 162
      Height = 31
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 3
      object BAide: THBitBtn
        Left = 132
        Top = 2
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
        TabOrder = 0
        OnClick = BAideClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BFerme: THBitBtn
        Left = 100
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BValider: THBitBtn
        Left = 70
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Exporter'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ModalResult = 1
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = BValiderClick
        Margin = 2
        Spacing = -1
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
    end
    object BRIB: THBitBtn
      Left = 4
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Modifier le RIB'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnClick = BRIBClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z2098_S16G1'
      IsControl = True
    end
    object BCompteRendu: THBitBtn
      Left = 69
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Voir le fichier g'#233'n'#233'r'#233
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = BCompteRenduClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z1826_S16G1'
      IsControl = True
    end
    object BParam: THBitBtn
      Tag = 1
      Left = 36
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Zoom param'#233'trage fiche soci'#233't'#233
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BParamClick
      Margin = 0
      Spacing = -1
      GlobalIndexImage = 'Z0008_S16G1'
      IsControl = True
    end
  end
  object BQ_GENERAL: THValComboBox
    Left = 68
    Top = 145
    Width = 293
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 2
    OnChange = BQ_GENERALChange
    TagDispatch = 0
    DataType = 'TTBANQUECP'
  end
  object NomFichier: TEdit
    Left = 68
    Top = 172
    Width = 273
    Height = 21
    TabOrder = 3
  end
  object Document: THValComboBox
    Left = 68
    Top = 197
    Width = 293
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 4
    OnChange = DocumentChange
    TagDispatch = 0
    DataType = 'TTMODELEBOR'
  end
  object Pages: TPageControl
    Left = 0
    Top = 222
    Width = 370
    Height = 171
    ActivePage = TS1
    Align = alBottom
    TabOrder = 5
    object TS1: TTabSheet
      Caption = 'Param'#232'tres'
      object H_DATEREMISE: THLabel
        Left = 10
        Top = 61
        Width = 125
        Height = 13
        Caption = 'Date de remise en ban&que'
        FocusControl = DATEREMISE
      end
      object lbl_NatEco: THLabel
        Left = 10
        Top = 88
        Width = 93
        Height = 13
        Caption = 'Nature '#233'conomique'
      end
      object lbl_Remise: THLabel
        Left = 8
        Top = 115
        Width = 109
        Height = 13
        Caption = 'R'#233'f'#233'rence de la remise'
        FocusControl = DATEREMISE
        Visible = False
      end
      object CRibTiers: TCheckBox
        Left = 8
        Top = 8
        Width = 345
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Affecter le &RIB principal du tiers sur les mouvements sans RIB'
        TabOrder = 0
      end
      object CCumulTiers: TCheckBox
        Left = 8
        Top = 32
        Width = 345
        Height = 17
        Alignment = taLeftJustify
        Caption = '&G'#233'n'#233'rer une ligne cumul'#233'e par tiers dans le fichier d'#39'export'
        TabOrder = 1
      end
      object DATEREMISE: THCritMaskEdit
        Left = 276
        Top = 57
        Width = 77
        Height = 21
        Ctl3D = True
        Enabled = False
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        Text = '  /  /    '
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object RefTireLib: TCheckBox
        Left = 8
        Top = 88
        Width = 345
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Utiliser le libell'#233' comme r'#233'f'#233'rence &tir'#233
        TabOrder = 3
      end
      object cbo_NatEco: THValComboBox
        Left = 108
        Top = 84
        Width = 246
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 4
        TagDispatch = 0
        DataType = 'CPNATECO'
        ComboWidth = 390
      end
      object txt_RefRemise: TEdit
        Left = 204
        Top = 111
        Width = 149
        Height = 21
        MaxLength = 16
        TabOrder = 5
        Visible = False
      end
    end
    object TS2: TTabSheet
      Caption = 'Emission'
      ImageIndex = 1
      object TCDestinataire: THLabel
        Left = 12
        Top = 66
        Width = 56
        Height = 13
        Caption = '&Destinataire'
        FocusControl = CDestinataire
        Visible = False
      end
      object FApercu: TCheckBox
        Left = 9
        Top = 18
        Width = 344
        Height = 13
        Alignment = taLeftJustify
        Caption = '&Aper'#231'u avant impression'
        TabOrder = 0
      end
      object REnvoi: TRadioGroup
        Left = 8
        Top = 91
        Width = 345
        Height = 34
        Caption = 'Monnaie d'#39'envoi'
        Columns = 2
        ItemIndex = 0
        Items.Strings = (
          'Euro'
          'Origine')
        TabOrder = 1
        Visible = False
      end
      object cTeleTrans: TCheckBox
        Left = 9
        Top = 43
        Width = 344
        Height = 13
        Alignment = taLeftJustify
        Caption = '&T'#233'l'#233'transmission des fichiers vers la banque'
        TabOrder = 2
        Visible = False
      end
      object CDestinataire: THValComboBox
        Left = 89
        Top = 62
        Width = 264
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 3
        Visible = False
        TagDispatch = 0
      end
    end
    object TS3: TTabSheet
      Caption = 'Options'
      ImageIndex = 2
      object TypeDebit: TRadioGroup
        Left = 159
        Top = 4
        Width = 197
        Height = 113
        Caption = 'Type de d'#233'bit de la remise'
        ItemIndex = 1
        Items.Strings = (
          'D'#233'bit global'
          'D'#233'bit unitaire par op'#233'ration'
          'D'#233'bit global par devise de transfert')
        TabOrder = 0
      end
      object ImputFrais: TRadioGroup
        Left = 4
        Top = 4
        Width = 149
        Height = 113
        Caption = 'Imputation des frais'
        ItemIndex = 2
        Items.Strings = (
          'B'#233'n'#233'ficiaire'
          'Emetteur et b'#233'n'#233'ficiaire'
          'Emetteur')
        TabOrder = 1
      end
    end
  end
  object Sauve: TSaveDialog
    Filter = 
      'Fichiers textes (*.txt).|*.txt|Fichiers ASCII (*.asc).|*.asc|Fic' +
      'hiers SAARI (*.pn*).|*.pn*|Fichiers EDIFICAS (*.edf).|*.edf|Tous' +
      ' fichiers (*.*).|*.*'
    Options = [ofHideReadOnly, ofShowHelp, ofExtensionDifferent, ofPathMustExist, ofNoReadOnlyReturn, ofNoLongNames]
    Title = 'Exportation des mouvements'
    Left = 176
    Top = 141
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Export CFONB;Export impossible ! Vous devez choisir une banque' +
        ';W;O;O;O;'
      
        '1;Export CFONB;Export impossible ! Vous devez choisir un fichier' +
        ';W;O;O;O;'
      '2;Export CFONB;Confirmez-vous l'#39'export CFONB ?;Q;YN;Y;Y;'
      
        '3;Export CFONB;Vous n'#39'avez pas choisi de mod'#232'le de document. L'#39'e' +
        'xport se fera sans '#233'mission d'#39'un document associ'#233'. Confirmez-vou' +
        's l'#39'op'#233'ration ?;Q;YN;Y;Y;'
      
        'ATTENTION. Certaines '#233'ch'#233'ances sont en cours de traitement par u' +
        'n autre utilisateur. Les '#233'critures s'#233'lectionn'#233'es ne seront pas c' +
        'onsid'#233'r'#233'es comme export'#233'es'
      
        '5;Export CFONB;Certaines '#233'ch'#233'ances ont d'#233'j'#224' '#233't'#233' export'#233'es. Confi' +
        'rmez-vous le traitement ?;Q;YN;Y;Y;'
      
        '6;Export CFONB;Vous venez d'#233'j'#224' d'#39'effectuer un export CFONB. Conf' +
        'irmez-vous le traitement ?;Q;YN;Y;Y;'
      
        '7;Export CFONB;Export impossible ! Vous devez renseigner un num'#233 +
        'ro d'#39#233'metteur;W;O;O;O;'
      
        '8;Export CFONB;Export impossible ! Le RIB de votre banque est in' +
        'complet.;W;O;O;O;'
      
        '9;Export CFONB;Vous avez des RIB incomplets et/ou des montants d' +
        'e mauvais sens. Ces '#233'l'#233'ments ont '#233't'#233' rejet'#233's. Confirmez-vous la ' +
        'suite du traitement ? ;Q;YN;N;N;'
      
        '10;Export CFONB;Export impossible ! Vous devez choisir un format' +
        '.;W;O;O;O;'
      'ATTENTION. Export non valid'#233' !'
      
        '12;Export CFONB;Les '#233'ch'#233'ances s'#233'lectionn'#233'es ont des modes de pai' +
        'ement diff'#233'rents. Confirmez-vous le traitement ?;Q;YN;Y;Y;'
      
        '13;Export CFONB;Les '#233'ch'#233'ances s'#233'lectionn'#233'es comportent des mouve' +
        'ments d'#233'biteurs et cr'#233'diteurs. Confirmez-vous le traitement ?;Q;' +
        'YN;Y;Y;'
      '14;Export CFONB;L'#39'export s'#39'est correctement termin'#233'.;E;O;O;O;'
      'LIBRE==LIBRE==LIBRE'
      
        '16;Export CFONB;Export non effectu'#233' ! Aucune ligne de mouvement ' +
        'n'#39'a pu '#234'tre export'#233'e.;W;O;O;O;'
      'Date de remise en banque'
      'Date de pr'#233'l'#232'vement effectif'
      'Date de virement effectif'
      
        '20;Export CFONB;Les '#233'ch'#233'ances s'#233'lectionn'#233'es ont des banques diff' +
        #233'rentes. Confirmez-vous le traitement ?;Q;YN;Y;Y;'
      
        '21;Export CFONB;La t'#233'l'#233'transmission ETEBAC a '#233'chou'#233'. Voulez-vous' +
        ' continuer le traitement ?;Q;YN;Y;Y;'
      
        '22;Export CFONB;Export impossible ! Les lignes s'#233'lectionn'#233'es ont' +
        ' une date d'#39#233'ch'#233'ance diff'#233'rente.;W;O;O;O;'
      
        '23;Export CFONB;Export impossible ! Le code BIC de la banque s'#233'l' +
        'ectionn'#233'e n'#39'est pas renseign'#233'.;W;O;O;O;'
      
        '24;Export CFONB;Export impossible ! Le code IBAN de la banque s'#233 +
        'lectionn'#233'e n'#39'est pas renseign'#233'.;W;O;O;O;'
      
        '25;Export CFONB;Export impossible ! Le code devise de la banque ' +
        's'#233'lectionn'#233'e n'#39'est pas renseign'#233'.;W;O;O;O;'
      '26;Export CFONB;Le code iban pour l'#39#233'criture n'#176';Q;YN;Y;Y;'
      '27;Export CFONB;Le compte n'#39'a pas de RIB principal.;W;O;O;O;'
      
        '28;Export CFONB;Le code iban du compte n'#39'est pas renseign'#233'.;W;O;' +
        'O;O;'
      '29;Export CFONB;Le code iso du pays;W;O;O;O;'
      '30;Export CFONB;Export impossible ! Le RIB du compte ;W;O;O;O;'
      '31;Export CFONB;La fiche du compte ;W;O;O;O;'
      
        '32;Export CFONB;Export impossible ! Vous devez indiquer une r'#233'f'#233 +
        'rence remise.;W;O;O;O;'
      '33;Export CFONB;Le code BIC du compte ;W;O;O;O;'
      '34;Export CFONB;Le code pays du compte ;W;O;O;O;'
      
        '35;Export CFONB;Export impossible ! Le RIB principal du compte ;' +
        'W;O;O;O;'
      ''
      '')
    Left = 120
    Top = 141
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 222
    Top = 141
  end
end

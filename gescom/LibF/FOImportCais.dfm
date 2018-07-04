inherited FImportCais: TFImportCais
  Left = 325
  Top = 258
  Caption = 'Installation d'#39'une caisse'
  PixelsPerInch = 96
  TextHeight = 13
  object lEtapeManu: THLabel [1]
    Left = 7
    Top = 314
    Width = 68
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
  inherited P: TPageControl
    Top = 4
    Height = 285
    ActivePage = PAGE3
    object PAGE1: TTabSheet
      Caption = 'PAGE1'
      ImageIndex = 4
      object LGPK_CAISSE: TLabel
        Left = 35
        Top = 118
        Width = 138
        Height = 13
        Caption = 'Entrez le numéro de la caisse'
        FocusControl = GPK_CAISSE
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
            'Cet assistant va vous aider à installer une nouvelle caisse sur ' +
            'un poste autonome.'
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
        Left = 186
        Top = 114
        Width = 119
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 1
        TagDispatch = 0
      end
      object GACTIONTABLE: TRadioGroup
        Left = 16
        Top = 172
        Width = 313
        Height = 57
        Caption = 'Action sur les tables existantes'
        Columns = 2
        ItemIndex = 0
        Items.Strings = (
          'Remplacer les tables'
          'Complétez les tables')
        TabOrder = 2
      end
    end
    object PAGE2: TTabSheet
      Caption = 'PAGE2'
      ImageIndex = 1
      object GORIGINE: TRadioGroup
        Left = 16
        Top = 90
        Width = 313
        Height = 50
        ItemIndex = 0
        Items.Strings = (
          'Sur le disque ou une disquette'
          'Sur le serveur du siège (par télétransmission).')
        TabOrder = 1
        OnClick = GORIGINEClick
      end
      object PFICHIER: TPanel
        Left = 16
        Top = 143
        Width = 313
        Height = 31
        TabOrder = 2
        object LNOMDOSSIER: THLabel
          Left = 7
          Top = 8
          Width = 73
          Height = 13
          Caption = 'Nom du dossier'
        end
        object NOMDOSSIER: THCritMaskEdit
          Left = 84
          Top = 4
          Width = 217
          Height = 21
          TabOrder = 0
          Text = 'C:\PGI00\Std'
          OnChange = ORIGINEChange
          TagDispatch = 0
          DataType = 'DIRECTORY'
          ElipsisButton = True
        end
      end
      object PTELECOM: TPanel
        Left = 16
        Top = 178
        Width = 313
        Height = 76
        TabOrder = 3
        Visible = False
        object LADRESSEIP1: THLabel
          Left = 8
          Top = 9
          Width = 51
          Height = 13
          Caption = 'Adresse IP'
          FocusControl = ADRESSEIP1
        end
        object LADRESSEIP2: THLabel
          Left = 133
          Top = 9
          Width = 3
          Height = 13
          Caption = ':'
          FocusControl = ADRESSEIP2
        end
        object LADRESSEIP3: THLabel
          Left = 191
          Top = 9
          Width = 3
          Height = 13
          Caption = ':'
          FocusControl = ADRESSEIP3
        end
        object LUSER: THLabel
          Left = 8
          Top = 32
          Width = 69
          Height = 13
          Caption = 'Utilisateur FTP'
        end
        object LMOTPASSE: THLabel
          Left = 8
          Top = 54
          Width = 64
          Height = 13
          Caption = 'Mot de passe'
          FocusControl = MOTPASSE
        end
        object LADRESSEIP4: THLabel
          Left = 249
          Top = 9
          Width = 3
          Height = 13
          Caption = ':'
          FocusControl = ADRESSEIP4
        end
        object ADRESSEIP1: TSpinEdit
          Left = 84
          Top = 4
          Width = 45
          Height = 22
          MaxValue = 255
          MinValue = 0
          TabOrder = 0
          Value = 10
          OnChange = ORIGINEChange
        end
        object ADRESSEIP2: TSpinEdit
          Left = 146
          Top = 4
          Width = 45
          Height = 22
          MaxValue = 255
          MinValue = 0
          TabOrder = 1
          Value = 1
          OnChange = ORIGINEChange
        end
        object ADRESSEIP3: TSpinEdit
          Left = 200
          Top = 4
          Width = 45
          Height = 22
          MaxValue = 255
          MinValue = 0
          TabOrder = 2
          Value = 1
          OnChange = ORIGINEChange
        end
        object USER: THCritMaskEdit
          Left = 84
          Top = 28
          Width = 218
          Height = 21
          TabOrder = 4
          OnChange = ORIGINEChange
          TagDispatch = 0
        end
        object MOTPASSE: THCritMaskEdit
          Left = 84
          Top = 50
          Width = 218
          Height = 21
          TabOrder = 5
          OnChange = ORIGINEChange
          TagDispatch = 0
        end
        object ADRESSEIP4: TSpinEdit
          Left = 258
          Top = 4
          Width = 45
          Height = 22
          MaxValue = 255
          MinValue = 0
          TabOrder = 3
          Value = 200
          OnChange = ORIGINEChange
        end
      end
      object PTITRE2: THPanel
        Left = 16
        Top = 8
        Width = 313
        Height = 78
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
          Height = 61
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
      end
    end
    object PAGE3: TTabSheet
      Caption = 'PAGE3'
      ImageIndex = 4
      object bCharge: TSpeedButton
        Left = 16
        Top = 92
        Width = 313
        Height = 22
        Caption = 'Chargement des données'
        Glyph.Data = {
          E6000000424DE60000000000000076000000280000000F0000000E0000000100
          0400000000007000000000000000000000001000000000000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          777077777777777777707700877777777770770B3087777777707770BB008777
          777077770BB300877770777770BBB00877707777700BBB007770777700BBB007
          77707777700BBB007770777777700BB0077077777777700B0070777777777770
          00707777777777777770}
        OnClick = bChargeClick
      end
      object CPTRDURECUP: THRichEditOLE
        Left = 16
        Top = 128
        Width = 313
        Height = 113
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
          
            '{\rtf1\ansi\ansicpg1252\deff0\deflang1036{\fonttbl{\f0\fnil Aria' +
            'l;}}'
          '\viewkind4\uc1\pard\f0\fs16 CPTRDURECUP'
          '\par }')
      end
      object PTITRE3: THPanel
        Left = 16
        Top = 8
        Width = 313
        Height = 78
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
          Height = 61
          Alignment = taCenter
          AutoSize = False
          Caption = 
            'Lancez le chargement des données à partir du fichier de démarrag' +
            'e de la caisse.'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
      end
    end
    object PAGE4: TTabSheet
      Caption = 'PAGE4'
      ImageIndex = 3
      object GDEVISE: TGroupBox
        Left = 16
        Top = 135
        Width = 313
        Height = 113
        Caption = 'Devise'
        Enabled = False
        TabOrder = 2
        object LSO_DEVISEPRINC: THLabel
          Left = 6
          Top = 14
          Width = 84
          Height = 13
          Caption = 'Devise du dossier'
          FocusControl = SO_DEVISEPRINC
        end
        object LSO_DECVALEUR: THLabel
          Left = 6
          Top = 66
          Width = 102
          Height = 13
          Caption = 'Nombre de décimales'
          FocusControl = SO_DEVISEPRINC
        end
        object LSO_TAUXEURO: THLabel
          Left = 6
          Top = 93
          Width = 119
          Height = 13
          Caption = 'Parité par rapport à l'#39'Euro'
          FocusControl = SO_TAUXEURO
        end
        object HLabel1: THLabel
          Left = 6
          Top = 42
          Width = 110
          Height = 13
          Caption = 'Devise de contrevaleur'
          FocusControl = SO_DEVISEPRINC
        end
        object SO_DEVISEPRINC: THValComboBox
          Left = 142
          Top = 10
          Width = 162
          Height = 21
          ItemHeight = 0
          TabOrder = 0
          Text = 'SO_DEVISEPRINC'
          TagDispatch = 0
          DataType = 'TTDEVISE'
        end
        object SO_DECVALEUR: TSpinEdit
          Left = 142
          Top = 64
          Width = 45
          Height = 22
          MaxValue = 2
          MinValue = 0
          TabOrder = 2
          Value = 2
        end
        object SO_TAUXEURO: THNumEdit
          Left = 141
          Top = 87
          Width = 88
          Height = 21
          Decimals = 5
          Digits = 12
          Masks.PositiveMask = '#,##0.00000'
          Debit = False
          TabOrder = 4
          UseRounding = True
          Validate = False
        end
        object SO_FONGIBLE: THValComboBox
          Left = 142
          Top = 38
          Width = 162
          Height = 21
          ItemHeight = 0
          TabOrder = 1
          Text = 'SO_FONGIBLE'
          TagDispatch = 0
          DataType = 'TTDEVISETOUTES'
        end
        object SO_TENUEEURO: TCheckBox
          Left = 208
          Top = 66
          Width = 97
          Height = 17
          Caption = 'Tenue en Euro'
          TabOrder = 3
          Visible = False
        end
      end
      object GETAB: TGroupBox
        Left = 16
        Top = 81
        Width = 313
        Height = 49
        Caption = 'Rattachement à'
        Enabled = False
        TabOrder = 1
        object LSO_ETABLISDEFAUT: THLabel
          Left = 6
          Top = 21
          Width = 65
          Height = 13
          Caption = 'Etablissement'
          FocusControl = SO_ETABLISDEFAUT
        end
        object SO_ETABLISDEFAUT: THValComboBox
          Left = 89
          Top = 17
          Width = 217
          Height = 21
          ItemHeight = 0
          TabOrder = 0
          Text = 'SO_ETABLISDEFAUT'
          TagDispatch = 0
          DataType = 'TTETABLISSEMENT'
        end
      end
      object PTITRE4: THPanel
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
        object TITRE4: THLabel
          Left = 16
          Top = 12
          Width = 273
          Height = 38
          Alignment = taCenter
          AutoSize = False
          Caption = 'Vérifiez les paramètres de la caisse'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
      end
    end
    object PAGE5: TTabSheet
      Caption = 'PAGE5'
      ImageIndex = 5
      object GDIVERS: TGroupBox
        Left = 16
        Top = 86
        Width = 313
        Height = 156
        Caption = 'Divers'
        Enabled = False
        TabOrder = 1
        object LSO_REGIMEDEFAUT: THLabel
          Left = 6
          Top = 24
          Width = 104
          Height = 13
          Caption = 'Régime de facturation'
          FocusControl = SO_REGIMEDEFAUT
        end
        object LSO_GCPERBASETARIF: THLabel
          Left = 6
          Top = 104
          Width = 142
          Height = 13
          Caption = 'Période de base pour les tarifs'
          FocusControl = SO_GCPERBASETARIF
        end
        object LSO_GCMODEREGLEDEFAUT: THLabel
          Left = 6
          Top = 51
          Width = 113
          Height = 13
          Caption = 'Mode de règlement tiers'
          FocusControl = SO_GCMODEREGLEDEFAUT
        end
        object LSO_GCTIERSDEFAUT: THLabel
          Left = 6
          Top = 78
          Width = 89
          Height = 13
          Caption = 'Tiers des transferts'
        end
        object LSO_GCPHOTOFICHE: THLabel
          Left = 6
          Top = 131
          Width = 99
          Height = 13
          Caption = 'Photo pour les fiches'
          FocusControl = SO_GCPHOTOFICHE
        end
        object SO_REGIMEDEFAUT: THValComboBox
          Left = 161
          Top = 20
          Width = 145
          Height = 21
          ItemHeight = 0
          TabOrder = 0
          TagDispatch = 0
          DataType = 'TTREGIMETVA'
        end
        object SO_GCPERBASETARIF: THValComboBox
          Left = 161
          Top = 100
          Width = 145
          Height = 21
          ItemHeight = 0
          TabOrder = 3
          TagDispatch = 0
          DataType = 'GCTARIFPERBASE'
        end
        object SO_GCMODEREGLEDEFAUT: THValComboBox
          Left = 161
          Top = 47
          Width = 145
          Height = 21
          ItemHeight = 0
          TabOrder = 1
          TagDispatch = 0
          DataType = 'TTMODEREGLE'
        end
        object SO_GCPHOTOFICHE: THValComboBox
          Left = 161
          Top = 127
          Width = 145
          Height = 21
          ItemHeight = 0
          TabOrder = 4
          TagDispatch = 0
          DataType = 'GCEMPLOIBLOB'
        end
        object SO_GCTIERSDEFAUT: THValComboBox
          Left = 161
          Top = 74
          Width = 145
          Height = 21
          ItemHeight = 0
          TabOrder = 2
          TagDispatch = 0
          DataType = 'GCTIERS'
        end
      end
      object PTITRE5: THPanel
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
        object TITRE5: THLabel
          Left = 16
          Top = 12
          Width = 273
          Height = 38
          Alignment = taCenter
          AutoSize = False
          Caption = 'Vérifiez les paramètres de la caisse'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
      end
    end
    object PAGE6: TTabSheet
      Caption = 'PAGE6'
      ImageIndex = 2
      object GCOORD: TGroupBox
        Left = 12
        Top = 15
        Width = 315
        Height = 225
        Caption = 'Coordonnées'
        TabOrder = 0
        object LSO_ASRESSE1: TLabel
          Left = 16
          Top = 42
          Width = 38
          Height = 13
          Caption = 'Adresse'
          FocusControl = SO_ADRESSE1
        end
        object LSO_MAIL: TLabel
          Left = 16
          Top = 198
          Width = 28
          Height = 13
          Caption = 'E-mail'
          FocusControl = SO_MAIL
        end
        object LSO_FAX: TLabel
          Left = 16
          Top = 176
          Width = 17
          Height = 13
          Caption = 'Fax'
          FocusControl = SO_FAX
        end
        object LSO_TELEPHONE: TLabel
          Left = 16
          Top = 154
          Width = 51
          Height = 13
          Caption = 'Téléphone'
          FocusControl = SO_TELEPHONE
        end
        object SO_SOCIETE: TEdit
          Left = 16
          Top = 15
          Width = 45
          Height = 21
          TabOrder = 0
          Text = 'SO_SOCIETE'
        end
        object SO_LIBELLE: TEdit
          Left = 76
          Top = 15
          Width = 213
          Height = 21
          TabOrder = 1
          Text = 'SO_LIBELLE'
        end
        object SO_ADRESSE1: TEdit
          Left = 76
          Top = 38
          Width = 213
          Height = 21
          TabOrder = 2
          Text = 'SO_ADRESSE1'
        end
        object SO_ADRESSE2: TEdit
          Left = 76
          Top = 60
          Width = 213
          Height = 21
          TabOrder = 3
          Text = 'SO_ADRESSE2'
        end
        object SO_ADRESSE3: TEdit
          Left = 76
          Top = 82
          Width = 213
          Height = 21
          TabOrder = 4
          Text = 'SO_ADRESSE3'
        end
        object SO_CODEPOSTAL: TEdit
          Left = 76
          Top = 104
          Width = 49
          Height = 21
          TabOrder = 5
          Text = 'SO_CODEPOSTAL'
        end
        object SO_VILLE: TEdit
          Left = 133
          Top = 104
          Width = 156
          Height = 21
          TabOrder = 6
          Text = 'SO_VILLE'
        end
        object SO_TELEPHONE: TEdit
          Left = 76
          Top = 150
          Width = 213
          Height = 21
          TabOrder = 8
          Text = 'SO_TELEPHONE'
        end
        object SO_FAX: TEdit
          Left = 76
          Top = 172
          Width = 213
          Height = 21
          TabOrder = 9
          Text = 'SO_FAX'
        end
        object SO_MAIL: TEdit
          Left = 76
          Top = 194
          Width = 213
          Height = 21
          TabOrder = 10
          Text = 'SO_MAIL'
        end
        object SO_PAYS: THValComboBox
          Left = 76
          Top = 126
          Width = 213
          Height = 21
          ItemHeight = 0
          TabOrder = 7
          Text = 'SO_PAYS'
          TagDispatch = 0
          DataType = 'TTPAYS'
        end
      end
    end
    object PAGE8: TTabSheet
      Caption = 'PAGE8'
      ImageIndex = 7
      object PTITRE8: THPanel
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
        object TITRE8: THLabel
          Left = 16
          Top = 12
          Width = 273
          Height = 61
          Alignment = taCenter
          AutoSize = False
          Caption = 
            'Vous pouvez enchaîner sur la récupération des données de la cais' +
            'se disponibles sur le site central.'
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
        Width = 277
        Height = 17
        Caption = 'Démarrage des échanges FO - BO'
        TabOrder = 1
      end
    end
    object PAGE7: TTabSheet
      Caption = 'PAGE7'
      ImageIndex = 6
      object PTITRE7: THPanel
        Left = 16
        Top = 8
        Width = 313
        Height = 78
        FullRepaint = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
        object TITRE7: THLabel
          Left = 12
          Top = 12
          Width = 293
          Height = 41
          Alignment = taCenter
          AutoSize = False
          Caption = 
            'Le bouton '#39'Fin'#39' lance l'#39'intégration du fichier de démarrage de l' +
            'a caisse :'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          WordWrap = True
        end
        object TITRE71: THLabel
          Left = 12
          Top = 57
          Width = 293
          Height = 13
          Alignment = taCenter
          AutoSize = False
          Caption = 'TITRE71'
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
        Top = 104
        Width = 313
        Height = 137
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
    end
  end
  inherited Msg: THMsgBox
    Mess.Strings = (
      'Etape'
      '1;?caption?;Voulez-vous quitter l'#39'assistant ?;Q;YN;Y;C;'
      '2;?caption?;Le compte FTP est obligatoire.;E;O;O;O;'
      '3;?caption?;Le mot de passe FTP est obligatoire.;E;O;O;O;'
      
        '4;?caption?;Le dossier de stockage du fichier de configuration e' +
        'st incorrect.;E;O;O;O;'
      
        '5;?caption?;L'#39'adresse IP du seveur du siège est obligatoire.;E;O' +
        ';O;O;'
      '6;?caption?;Le numéro de caisse est obligatoire.;E;O;O;O;'
      
        '7;?caption?;Vous devez lancer le chargement des données.;E;O;O;O' +
        ';'
      
        '8;?caption?;Le fichier de configuration de la caisse n'#39'existe pa' +
        's dans le dossier indiqué.;E;O;O;O;'
      
        '9;?caption?;Attention : vous devez lancer la bascule EURO du dos' +
        'sier.;W;O;O;O;'
      '10;?caption?;Le code représentant est obligatoire.;E;O;O;O;'
      
        '11;?caption?;Le fichier de configuration du poste n'#39'existe pas d' +
        'ans le dossier indiqué.;E;O;O;O;')
    Left = 59
    Top = 4
  end
  inherited HMTrad: THSystemMenu
    Top = 65535
  end
end

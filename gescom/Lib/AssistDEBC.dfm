inherited FAssistDEBC: TFAssistDEBC
  Left = 308
  Top = 486
  HelpContext = 110000370
  Caption = 'Assistant exportations D.E.B.'
  ClientWidth = 534
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Caption = 'Etape 1/1'
  end
  object PBevel2: TBevel [2]
    Left = 0
    Top = 0
    Width = 534
    Height = 337
    Align = alClient
  end
  object Bevel3: TBevel [3]
    Left = 0
    Top = 0
    Width = 534
    Height = 337
    Align = alClient
  end
  inherited bAide: TToolbarButton97
    Visible = True
  end
  object Bfermer: TButton [9]
    Left = 433
    Top = 308
    Width = 75
    Height = 23
    Caption = '&Fermer'
    TabOrder = 5
    Visible = False
    OnClick = BfermerClick
  end
  inherited Plan: TPanel
    Left = 176
  end
  inherited GroupBox1: TGroupBox
    Enabled = False
  end
  inherited P: TPageControl
    Left = 176
    Top = 0
    Width = 354
    Height = 277
    ActivePage = TabSheet1
    object TabSheet1: TTabSheet
      Caption = 'Général'
      object PBevel1: TBevel
        Left = 0
        Top = 41
        Width = 346
        Height = 208
        Align = alClient
      end
      object TINTRO: THLabel
        Left = 19
        Top = 45
        Width = 312
        Height = 45
        AutoSize = False
        Caption = 
          'Cet assistant vous guide pour réaliser les exportations vers le ' +
          'logiciel IDEP en vue de réaliser vos déclarations d'#39'échange de b' +
          'iens.'
        WordWrap = True
      end
      object PTITRE: THPanel
        Left = 0
        Top = 0
        Width = 346
        Height = 41
        Align = alTop
        Caption = 'Exportations pour la D.E.B.'
        FullRepaint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -19
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
      end
      object GBOPTIONS: TGroupBox
        Left = 8
        Top = 88
        Width = 330
        Height = 89
        Caption = ' Type de fichiers à générer '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object TEXPLIQ: THLabel
          Left = 192
          Top = 20
          Width = 121
          Height = 52
          AutoSize = False
          Caption = 
            'Choisissez ici les fichiers que vous désirez générer pour le log' +
            'iciel IDEP'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object CB_CLIENTS: TCheckBox
          Left = 13
          Top = 15
          Width = 148
          Height = 17
          Caption = 'Fichier &Clients'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 0
          OnClick = CB_CLIENTSOnClick
        end
        object CB_ARTICLES: TCheckBox
          Left = 13
          Top = 33
          Width = 148
          Height = 17
          Caption = 'Fichier &Articles'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 1
          OnClick = CB_ARTICLESOnClick
        end
        object CB_EXPORT: TCheckBox
          Left = 13
          Top = 50
          Width = 148
          Height = 17
          Caption = 'Déclaration à l'#39'&Exportation'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 2
          OnClick = CB_EXPORTOnClick
        end
        object CB_INTRO: TCheckBox
          Left = 13
          Top = 68
          Width = 148
          Height = 17
          Caption = 'Déclaration à l'#39'&Introduction'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 3
          OnClick = CB_INTROOnClick
        end
      end
      object GroupBox2: TGroupBox
        Left = 8
        Top = 184
        Width = 330
        Height = 58
        Caption = 'Format de fichiers à générer '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        object HLabel1: THLabel
          Left = 192
          Top = 9
          Width = 121
          Height = 45
          AutoSize = False
          Caption = 
            'Choisissez ici le format que vous désirez générer pour le logici' +
            'el IDEP'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object FRMT_EXC: TRadioButton
          Left = 13
          Top = 18
          Width = 173
          Height = 17
          Caption = 'Fichier .XLS (version Windows)'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = FRMT_EXCClick
        end
        object FRMT_SDF: TRadioButton
          Left = 13
          Top = 34
          Width = 173
          Height = 17
          Caption = 'Fichier .SDF (version DOS)'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnClick = FRMT_SDFClick
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'Clients'
      ImageIndex = 1
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 346
        Height = 249
        Align = alClient
      end
      object Infoscli: THLabel
        Left = 64
        Top = 20
        Width = 221
        Height = 37
        AutoSize = False
        Caption = 
          'Sélectionnez ici les clients que vous désirez transmettre au log' +
          'iciel IDEP'
        WordWrap = True
      end
      object GBCLIENTS: TGroupBox
        Left = 28
        Top = 61
        Width = 305
        Height = 136
        BiDiMode = bdLeftToRight
        Caption = ' Sélection des clients '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentBiDiMode = False
        ParentFont = False
        TabOrder = 0
        object TDATECLI: THLabel
          Left = 21
          Top = 39
          Width = 117
          Height = 13
          Caption = '&Clients modifiés depuis le'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TFICCLI: THLabel
          Left = 21
          Top = 88
          Width = 68
          Height = 13
          Caption = '&Nom du fichier'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object RechFicCli: TToolbarButton97
          Left = 268
          Top = 83
          Width = 17
          Height = 23
          Hint = 'Parcourir'
          Caption = '...'
          Flat = False
          ParentShowHint = False
          ShowHint = True
          OnClick = RechFicCliOnClick
        end
        object CDATECLI: THCritMaskEdit
          Left = 150
          Top = 35
          Width = 134
          Height = 21
          EditMask = '00/00/0000;1;_'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 10
          ParentFont = False
          TabOrder = 0
          Text = '  /  /    '
          OnExit = CDATECLIOnExit
          TagDispatch = 0
          Operateur = Superieur
          OpeType = otDate
          DefaultDate = odDebMois
          ControlerDate = True
        end
        object NomFicCli: TEdit
          Left = 150
          Top = 84
          Width = 119
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
    end
    object TabSheet3: TTabSheet
      Caption = 'Articles'
      ImageIndex = 2
      object Bevel2: TBevel
        Left = 0
        Top = 0
        Width = 346
        Height = 249
        Align = alClient
      end
      object Infosarticles: THLabel
        Left = 64
        Top = 20
        Width = 221
        Height = 37
        AutoSize = False
        Caption = 
          'Sélectionnez ici les articles que vous désirez transmettre au lo' +
          'giciel IDEP'
        WordWrap = True
      end
      object GBARTICLES: TGroupBox
        Left = 28
        Top = 61
        Width = 305
        Height = 136
        BiDiMode = bdLeftToRight
        Caption = ' Sélection des articles '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentBiDiMode = False
        ParentFont = False
        TabOrder = 0
        object TDATEART: THLabel
          Left = 21
          Top = 39
          Width = 108
          Height = 13
          Caption = '&Articles créés depuis le'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TFICART: THLabel
          Left = 21
          Top = 88
          Width = 68
          Height = 13
          Caption = '&Nom du fichier'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object RechFicArt: TToolbarButton97
          Left = 268
          Top = 83
          Width = 17
          Height = 23
          Hint = 'Parcourir'
          Caption = '...'
          Flat = False
          ParentShowHint = False
          ShowHint = True
          OnClick = RechFicArtOnClick
        end
        object CDATEART: THCritMaskEdit
          Left = 150
          Top = 35
          Width = 134
          Height = 21
          EditMask = '00/00/0000;1;_'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 10
          ParentFont = False
          TabOrder = 0
          Text = '  /  /    '
          OnExit = CDATEARTOnExit
          TagDispatch = 0
          Operateur = Superieur
          OpeType = otDate
          DefaultDate = odDebMois
          ControlerDate = True
        end
        object NomFicArt: TEdit
          Left = 150
          Top = 84
          Width = 119
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
    end
    object TabSheet4: TTabSheet
      Caption = 'DEB export'
      ImageIndex = 3
      object Bevel4: TBevel
        Left = 0
        Top = 0
        Width = 346
        Height = 249
        Align = alClient
      end
      object InfosExport: THLabel
        Left = 20
        Top = 16
        Width = 313
        Height = 29
        AutoSize = False
        Caption = 
          'Paramétrez ici l'#39'extraction de vos ventes en direction de la CEE' +
          ' '
        WordWrap = True
      end
      object GBEXPORT: TGroupBox
        Left = 20
        Top = 44
        Width = 313
        Height = 189
        BiDiMode = bdLeftToRight
        Caption = ' Déclaration à l'#39'export '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentBiDiMode = False
        ParentFont = False
        TabOrder = 0
        object TDATEDE: THLabel
          Left = 12
          Top = 27
          Width = 56
          Height = 13
          Caption = '&Factures du'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TDATEFE: THLabel
          Left = 180
          Top = 27
          Width = 12
          Height = 13
          Caption = '&au'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TKILOE: THLabel
          Left = 12
          Top = 67
          Width = 85
          Height = 13
          Caption = 'Code unité du &kilo'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TLIBREE: THLabel
          Left = 12
          Top = 111
          Width = 151
          Height = 13
          Caption = 'Montant Libre pour &unités suppl.'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TFICEXP: THLabel
          Left = 12
          Top = 153
          Width = 68
          Height = 13
          Caption = '&Nom du fichier'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object RechFicExport: TToolbarButton97
          Left = 280
          Top = 149
          Width = 17
          Height = 21
          Hint = 'Parcourir'
          Caption = '...'
          Flat = False
          ParentShowHint = False
          ShowHint = True
          OnClick = RechFicExportOnClick
        end
        object DateDebExp: THCritMaskEdit
          Left = 76
          Top = 23
          Width = 93
          Height = 21
          EditMask = '00/00/0000;1;_'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 10
          ParentFont = False
          TabOrder = 0
          Text = '  /  /    '
          OnExit = DateDebExpOnExit
          TagDispatch = 0
          Operateur = Superieur
          OpeType = otDate
          DefaultDate = odDebMois
          ControlerDate = True
        end
        object DateFinExp: THCritMaskEdit
          Left = 204
          Top = 23
          Width = 93
          Height = 21
          EditMask = '00/00/0000;1;_'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 10
          ParentFont = False
          TabOrder = 1
          Text = '  /  /    '
          OnExit = DateFinExpOnExit
          TagDispatch = 0
          Operateur = Superieur
          OpeType = otDate
          DefaultDate = odFinMois
          ControlerDate = True
        end
        object KiloE: THValComboBox
          Left = 172
          Top = 63
          Width = 125
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 2
          OnClick = KiloEOnClick
          TagDispatch = 0
          Exhaustif = exNon
          VideString = '<<Aucun>>'
          DataType = 'GCQUALUNITPOIDS'
          ComboWidth = 125
        end
        object MontantLibreE: THValComboBox
          Left = 172
          Top = 107
          Width = 125
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 3
          OnClick = MontantLibreEOnClick
          Items.Strings = (
            '<<Aucun>>'
            'Montant libre 1'
            'Montant libre 2'
            'Montant libre 3 ')
          TagDispatch = 0
          Vide = True
          VideString = '<<Aucun>>'
          Values.Strings = (
            '<<Aucun>>'
            'Montant libre 1'
            'Montant libre 2'
            'Montant libre 3 '
            '')
          ComboWidth = 125
        end
        object NomFicExport: TEdit
          Left = 172
          Top = 149
          Width = 109
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 4
        end
      end
    end
    object TabSheet5: TTabSheet
      Caption = 'DEB intro'
      ImageIndex = 4
      object Bevel5: TBevel
        Left = 0
        Top = 0
        Width = 346
        Height = 249
        Align = alClient
      end
      object InfosIntro: THLabel
        Left = 16
        Top = 16
        Width = 321
        Height = 25
        AutoSize = False
        Caption = 
          'Paramétrez ici l'#39'extraction de vos achats en provenance de la CE' +
          'E '
        WordWrap = True
      end
      object GBINTRO: TGroupBox
        Left = 20
        Top = 44
        Width = 313
        Height = 185
        BiDiMode = bdLeftToRight
        Caption = ' Déclaration à l'#39'Introduction '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentBiDiMode = False
        ParentFont = False
        TabOrder = 0
        object TDATEDI: THLabel
          Left = 12
          Top = 27
          Width = 56
          Height = 13
          Caption = '&Factures du'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TDATEFI: THLabel
          Left = 180
          Top = 27
          Width = 12
          Height = 13
          Caption = '&au'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TKILOI: THLabel
          Left = 12
          Top = 67
          Width = 85
          Height = 13
          Caption = 'Code unité du &kilo'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TLIBREI: THLabel
          Left = 12
          Top = 111
          Width = 151
          Height = 13
          Caption = 'Montant Libre pour &unités suppl.'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TFICINT: THLabel
          Left = 12
          Top = 153
          Width = 68
          Height = 13
          Caption = '&Nom du fichier'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object RechFicIntro: TToolbarButton97
          Left = 280
          Top = 149
          Width = 17
          Height = 21
          Hint = 'Parcourir'
          Caption = '...'
          Flat = False
          ParentShowHint = False
          ShowHint = True
          OnClick = RechFicIntroOnClick
        end
        object DateDebInt: THCritMaskEdit
          Left = 76
          Top = 23
          Width = 93
          Height = 21
          EditMask = '00/00/0000;1;_'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 10
          ParentFont = False
          TabOrder = 0
          Text = '  /  /    '
          OnExit = DateDebIntOnExit
          TagDispatch = 0
          Operateur = Superieur
          OpeType = otDate
          DefaultDate = odDebMois
          ControlerDate = True
        end
        object DateFinInt: THCritMaskEdit
          Left = 204
          Top = 23
          Width = 93
          Height = 21
          EditMask = '00/00/0000;1;_'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 10
          ParentFont = False
          TabOrder = 1
          Text = '  /  /    '
          OnExit = DateFinIntOnExit
          TagDispatch = 0
          Operateur = Superieur
          OpeType = otDate
          DefaultDate = odFinMois
          ControlerDate = True
        end
        object KiloI: THValComboBox
          Left = 172
          Top = 63
          Width = 125
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
          OnClick = KiloIOnClick
          TagDispatch = 0
          VideString = '<<Aucun>>'
          DataType = 'GCQUALUNITPOIDS'
          ComboWidth = 125
        end
        object MontantLibreI: THValComboBox
          Left = 172
          Top = 107
          Width = 125
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 3
          OnClick = MontantLibreIOnClick
          Items.Strings = (
            '<<Aucun>>'
            'Montant libre 1'
            'Montant libre 2'
            'Montant libre 3 ')
          TagDispatch = 0
          Vide = True
          VideString = '<<Aucun>>'
          Values.Strings = (
            '<<Aucun>>'
            'Montant libre 1'
            'Montant libre 2'
            'Montant libre 3 '
            '')
          ComboWidth = 125
        end
        object NomFicIntro: TEdit
          Left = 172
          Top = 149
          Width = 109
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 4
        end
      end
    end
    object TabSheet6: TTabSheet
      Caption = 'Récapitulatif'
      ImageIndex = 5
      object PBevel4: TBevel
        Left = 0
        Top = 0
        Width = 346
        Height = 249
        Align = alClient
      end
      object TRecap: THLabel
        Left = 18
        Top = 108
        Width = 59
        Height = 13
        Caption = 'Récapitulatif'
      end
      object PanelFin: TPanel
        Left = 13
        Top = 4
        Width = 329
        Height = 97
        TabOrder = 0
        object TTextFin1: THLabel
          Left = 23
          Top = 6
          Width = 287
          Height = 39
          Caption = 
            'Le paramètrage est maintenant correctement renseigné pour permet' +
            'tre le lancement de l'#39'extraction des données pour le logiciel ID' +
            'EP de déclaration d'#39'échange de biens.'
          WordWrap = True
        end
        object TTextFin2: THLabel
          Left = 23
          Top = 53
          Width = 285
          Height = 39
          Caption = 
            'Si vous désirez revoir le paramétrage, il suffit de cliquer sur ' +
            'le bouton Précédent sinon, le bouton Fin, permet de débuter le t' +
            'raitement.'
          WordWrap = True
        end
      end
      object ListRecap: TListBox
        Left = 13
        Top = 126
        Width = 329
        Height = 116
        Color = clBtnFace
        ItemHeight = 13
        TabOrder = 1
      end
    end
  end
  inherited Msg: THMsgBox
    Top = 8
  end
  inherited HMTrad: THSystemMenu
    Top = 7
  end
  object LocaliseArt: TSaveDialog
    Filter = 'Fichiers SDF (*.sdf).|*.sdf|Tous fichiers (*.*).|*.*'
    Options = [ofReadOnly, ofHideReadOnly, ofShowHelp, ofExtensionDifferent, ofPathMustExist, ofFileMustExist, ofNoLongNames]
    Title = 'Fichier des articles'
    Left = 79
    Top = 252
  end
  object LocaliseCli: TSaveDialog
    Filter = 'Fichiers SDF (*.sdf).|*.sdf|Tous fichiers (*.*).|*.*'
    Options = [ofReadOnly, ofHideReadOnly, ofShowHelp, ofExtensionDifferent, ofPathMustExist, ofFileMustExist, ofNoLongNames]
    Title = 'Fichier des clients'
    Left = 3
    Top = 252
  end
  object LocaliseExport: TSaveDialog
    Filter = 'Fichiers SDF (*.sdf).|*.sdf|Tous fichiers (*.*).|*.*'
    Options = [ofReadOnly, ofHideReadOnly, ofShowHelp, ofExtensionDifferent, ofPathMustExist, ofFileMustExist, ofNoLongNames]
    Title = 'Déclaration à l'#39'export'
    Left = 27
    Top = 252
  end
  object LocaliseIntro: TSaveDialog
    Filter = 'Fichiers SDF (*.sdf).|*.sdf|Tous fichiers (*.*).|*.*'
    Options = [ofReadOnly, ofHideReadOnly, ofShowHelp, ofExtensionDifferent, ofPathMustExist, ofFileMustExist, ofNoLongNames]
    Title = 'Déclarataion à l'#39'introduction'
    Left = 51
    Top = 252
  end
  object HMsgErr: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'n'#39'est pas une date valide.')
    Left = 60
    Top = 8
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Les dates de l'#39'extraction pour l'#39'exportation sont in' +
        'cohérentes;W;O;O;O;'
      
        '1;?caption?;Les dates de l'#39'extraction pour l'#39'exportation ne conc' +
        'ernent pas le même mois;W;O;O;O;'
      '2;?caption?;Vous devez indiquer le code unité du Kilo;W;O;O;O;'
      
        '3;?caption?;Les dates de l'#39'extraction pour l'#39'introduction sont i' +
        'ncohérentes.;W;O;O;O;'
      
        '4;?caption?;Les dates de l'#39'extraction pour l'#39'introduction ne con' +
        'cernent pas le même mois.;W;O;O;O;'
      
        '5;?caption?;Le fichier Clients existe déjà. Faut-il le supprimer' +
        ' ?;Q;YNC;N;C;'
      '6;?caption?;Impossible d'#39'ouvrir le fichier clients.;W;O;O;O;'
      '7;?caption?;Il n'#39'y a aucun client dans votre sélection.;W;O;O;O;'
      
        '8;?caption?;Il n'#39'y a aucun article dans votre sélection.;W;O;O;O' +
        ';'
      
        '9;?caption?;Le fichier Articles existe déjà. Faut-il le supprime' +
        'r ?;Q;YNC;N;C;'
      '10;?caption?;Impossible d'#39'ouvrir le fichier Articles.;W;O;O;O;'
      
        '11;?caption?;Aucune pièce sélectionnée pour l'#39'exportation.;W;O;O' +
        ';O;'
      
        '12;?caption?;Aucune pièce sélectionnée pour l'#39'introduction.;W;O;' +
        'O;O;'
      
        '13;?caption?;Traitement terminé.... Consultez le compte rendu.;W' +
        ';O;O;O;'
      
        '14;?caption?;Le fichier DEB à l'#39'exportation existe déjà. Faut-il' +
        ' le supprimer ?;Q;YNC;N;C;'
      
        '15;?caption?;Le fichier DEB à l'#39'introduction existe déjà. Faut-i' +
        'l le supprimer ?;Q;YNC;N;C;'
      
        '16;?caption?;Impossible d'#39'ouvrir le fichier DEB à l'#39'exportation.' +
        ';W;O;O;O;'
      
        '17;?caption?;Impossible d'#39'ouvrir le fichier DEB à l'#39'introduction' +
        '.;W;O;O;O;')
    Left = 132
    Top = 8
  end
end

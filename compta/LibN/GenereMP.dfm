object FGenereMP: TFGenereMP
  Left = 285
  Top = 195
  Width = 497
  Height = 324
  HelpContext = 999999212
  BorderIcons = [biSystemMenu]
  Caption = 'Param'#232'tres de g'#233'n'#233'ration'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
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
    Width = 489
    Height = 254
    ActivePage = PGenere
    Align = alClient
    TabOrder = 0
    object PGenere: TTabSheet
      Caption = 'Informations'
      object Bevel7: TBevel
        Left = 0
        Top = 0
        Width = 481
        Height = 226
        Align = alClient
      end
      object H_MODEGENERE: TLabel
        Left = 30
        Top = 79
        Width = 88
        Height = 13
        Caption = '&Mode de paiement'
        FocusControl = MODEGENERE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Visible = False
      end
      object HDateReg: THLabel
        Left = 6
        Top = -4
        Width = 116
        Height = 13
        Caption = '&Date de comptabilisation'
        FocusControl = DATEGENERATION
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
      end
      object HCpteGen: THLabel
        Left = 9
        Top = 39
        Width = 89
        Height = 13
        Caption = 'Compte &g'#233'n'#233'ration'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel4: THLabel
        Left = 230
        Top = 77
        Width = 78
        Height = 13
        Caption = 'N'#176' de &bordereau'
        FocusControl = NUMENCADECA
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Visible = False
      end
      object Label1: TLabel
        Left = 244
        Top = 12
        Width = 52
        Height = 13
        Caption = '&G'#233'n'#233'ration'
        FocusControl = GROUPEENCADECA
      end
      object HJournalGenere: TLabel
        Left = 9
        Top = 12
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = JournalGenere
      end
      object LAnal: TLabel
        Left = 244
        Top = 39
        Width = 49
        Height = 13
        Caption = '&Analytique'
        FocusControl = ANAL
      end
      object TTauxDev: TLabel
        Left = 245
        Top = 215
        Width = 142
        Height = 13
        Caption = 'Taux devise (F5 pour modifier)'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HDevise: TLabel
        Left = 12
        Top = 214
        Width = 33
        Height = 13
        Caption = 'Devise'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TDevise: TLabel
        Left = 64
        Top = 214
        Width = 53
        Height = 13
        Caption = 'D_DEVISE'
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object MODEGENERE: THValComboBox
        Left = 142
        Top = 75
        Width = 107
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 7
        Visible = False
        OnChange = MODEGENEREChange
        TagDispatch = 0
        Vide = True
        VideString = 'Inchang'#233
        DataType = 'TTMODEPAIE'
      end
      object DATEGENERATION: THCritMaskEdit
        Left = 156
        Top = -8
        Width = 85
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 10
        ParentCtl3D = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 8
        Text = '  /  /    '
        Visible = False
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ElipsisButton = True
        ControlerDate = True
      end
      object DATEECHEANCE: THCritMaskEdit
        Left = 332
        Top = -3
        Width = 85
        Height = 21
        Ctl3D = True
        Enabled = False
        EditMask = '!99/99/0000;1;_'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 10
        ParentCtl3D = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 10
        Text = '  /  /    '
        Visible = False
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ElipsisButton = True
        ControlerDate = True
      end
      object CChoixDate: TCheckBox
        Left = 180
        Top = -1
        Width = 145
        Height = 17
        Caption = '&Forcer date d'#39#233'ch'#233'ance'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 9
        Visible = False
        OnClick = CChoixDateClick
      end
      object CompteGeneration: THCritMaskEdit
        Left = 124
        Top = 35
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 17
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
        OnExit = CompteGenerationExit
        TagDispatch = 0
        DataType = 'TZGENERAL'
        ElipsisButton = True
      end
      object GroupBox1: TGroupBox
        Left = 244
        Top = 89
        Width = 233
        Height = 121
        Caption = 'Mouvements sur compte de g'#233'n'#233'ration'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 6
        object TREFECR2: THLabel
          Left = 6
          Top = 24
          Width = 50
          Height = 13
          Caption = 'R'#233'&f'#233'rence'
          FocusControl = REFECR2
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel3: THLabel
          Left = 6
          Top = 48
          Width = 30
          Height = 13
          Caption = 'L&ibell'#233
          FocusControl = LIBECR2
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel7: THLabel
          Left = 7
          Top = 72
          Width = 58
          Height = 13
          Caption = '&Ref. externe'
          FocusControl = REFEXT2
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel8: THLabel
          Left = 7
          Top = 96
          Width = 42
          Height = 13
          Caption = '&Ref. libre'
          FocusControl = REFLIB2
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object REFECR2: THCritMaskEdit
          Tag = 5
          Left = 68
          Top = 20
          Width = 154
          Height = 21
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 35
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          TagDispatch = 0
          ElipsisButton = True
          OnElipsisClick = BAssistClick
        end
        object LIBECR2: THCritMaskEdit
          Tag = 6
          Left = 68
          Top = 44
          Width = 154
          Height = 21
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 35
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 1
          TagDispatch = 0
          ElipsisButton = True
          OnElipsisClick = BAssistClick
        end
        object REFLIB2: THCritMaskEdit
          Tag = 8
          Left = 68
          Top = 92
          Width = 154
          Height = 21
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 35
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 3
          TagDispatch = 0
          ElipsisButton = True
          OnElipsisClick = BAssistClick
        end
        object REFEXT2: THCritMaskEdit
          Tag = 7
          Left = 68
          Top = 68
          Width = 154
          Height = 21
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 35
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 2
          TagDispatch = 0
          ElipsisButton = True
          OnElipsisClick = BAssistClick
        end
      end
      object GroupBox2: TGroupBox
        Left = 6
        Top = 89
        Width = 233
        Height = 121
        Caption = 'Mouvements sur comptes s'#233'lectionn'#233's '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 5
        object TREFECR1: THLabel
          Left = 7
          Top = 24
          Width = 50
          Height = 13
          Caption = '&R'#233'f'#233'rence'
          FocusControl = REFECR1
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel1: THLabel
          Left = 7
          Top = 48
          Width = 30
          Height = 13
          Caption = '&Libell'#233
          FocusControl = LIBECR1
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel5: THLabel
          Left = 7
          Top = 72
          Width = 58
          Height = 13
          Caption = '&Ref. externe'
          FocusControl = REFEXT1
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel6: THLabel
          Left = 7
          Top = 96
          Width = 42
          Height = 13
          Caption = '&Ref. libre'
          FocusControl = REFLIB1
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LIBECR1: THCritMaskEdit
          Tag = 2
          Left = 68
          Top = 44
          Width = 154
          Height = 21
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 35
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 1
          OnChange = REFECR1Change
          TagDispatch = 0
          ElipsisButton = True
          OnElipsisClick = BAssistClick
        end
        object REFECR1: THCritMaskEdit
          Tag = 1
          Left = 68
          Top = 20
          Width = 154
          Height = 21
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 35
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          OnChange = REFECR1Change
          TagDispatch = 0
          ElipsisButton = True
          OnElipsisClick = BAssistClick
        end
        object REFEXT1: THCritMaskEdit
          Tag = 3
          Left = 68
          Top = 68
          Width = 154
          Height = 21
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 35
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 2
          OnChange = REFECR1Change
          TagDispatch = 0
          ElipsisButton = True
          OnElipsisClick = BAssistClick
        end
        object REFLIB1: THCritMaskEdit
          Tag = 4
          Left = 68
          Top = 92
          Width = 154
          Height = 21
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 35
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 3
          OnChange = REFECR1Change
          TagDispatch = 0
          ElipsisButton = True
          OnElipsisClick = BAssistClick
        end
      end
      object NUMENCADECA: THCritMaskEdit
        Left = 380
        Top = 73
        Width = 85
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 17
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 11
        Visible = False
        TagDispatch = 0
        DataType = 'CPNUMENCADECA'
        ElipsisButton = True
      end
      object GROUPEENCADECA: THValComboBox
        Left = 312
        Top = 8
        Width = 154
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        DataType = 'CPGROUPEENCADECA'
      end
      object JournalGenere: THValComboBox
        Left = 124
        Top = 8
        Width = 107
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 0
        OnExit = JournalGenereExit
        TagDispatch = 0
        Plus = 'AND J_MODESAISIE="-"'
        VideString = 'A saisir'
        DataType = 'TTJALBANQUEOD'
      end
      object ANAL: THValComboBox
        Left = 312
        Top = 35
        Width = 154
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 3
        Items.Strings = (
          'Report en d'#233'tail'
          'Pr'#233'-ventilation'
          'Section d'#39'attente')
        TagDispatch = 0
        Values.Strings = (
          'DET'
          'VEN'
          'ATT')
      end
      object TAUXDEV: THNumEdit
        Left = 397
        Top = 211
        Width = 79
        Height = 21
        Decimals = 6
        Digits = 12
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Masks.PositiveMask = '#,##0.000000000'
        Debit = False
        ParentFont = False
        TabOrder = 12
        UseRounding = True
        Validate = False
        OnKeyPress = TAUXDEVKeyPress
      end
      object AlerteEcheMP: TCheckBox
        Left = 8
        Top = 64
        Width = 226
        Height = 17
        Caption = '&Alerte si incoh'#233'rence '#233'ch'#233'ance/paiement'
        TabOrder = 4
        OnClick = CChoixDateClick
      end
      object LettrageAuto: TCheckBox
        Left = 240
        Top = 64
        Width = 169
        Height = 17
        Caption = 'Lettrage des pi'#232'ces en devise'
        Checked = True
        State = cbChecked
        TabOrder = 13
        Visible = False
        OnClick = CChoixDateClick
      end
    end
    object PEmission: TTabSheet
      Caption = 'Emissions'
      ImageIndex = 1
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 481
        Height = 226
        Align = alClient
      end
      object HLabel10: THLabel
        Left = 214
        Top = 41
        Width = 32
        Height = 13
        Caption = '&Format'
        FocusControl = FORMATCFONB
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TER_DOCUMENT: THLabel
        Left = 214
        Top = 97
        Width = 35
        Height = 13
        Caption = '&Mod'#232'le'
        FocusControl = ModeleBordereau
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TER_ENVOITRANS: THLabel
        Left = 33
        Top = 150
        Width = 151
        Height = 13
        Caption = '&T'#233'l'#233'transmission vers la banque'
        FocusControl = ModeTeletrans
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object EXPORTCFONB: TCheckBox
        Left = 33
        Top = 39
        Width = 126
        Height = 17
        Alignment = taLeftJustify
        Caption = 'E&xport CFONB'
        TabOrder = 0
      end
      object FORMATCFONB: THValComboBox
        Left = 264
        Top = 37
        Width = 176
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 1
        TagDispatch = 0
        DataType = 'TTTYPEEXPORTCFONB'
      end
      object EnvoiBORDEREAU: TCheckBox
        Left = 33
        Top = 95
        Width = 126
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Emission &bordereau'
        TabOrder = 2
      end
      object ModeleBordereau: THValComboBox
        Left = 264
        Top = 93
        Width = 176
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 3
        TagDispatch = 0
        DataType = 'TTMODELEBOR'
      end
      object ModeTeletrans: THValComboBox
        Left = 264
        Top = 146
        Width = 176
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 4
        TagDispatch = 0
        DataType = 'TTENVOITRANS'
      end
    end
  end
  object Dock: TDock97
    Left = 0
    Top = 254
    Width = 489
    Height = 36
    AllowDrag = False
    Position = dpBottom
    object PanelBouton: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 32
      ClientWidth = 489
      Caption = 'Barre outils multicrit'#232're'
      ClientAreaHeight = 32
      ClientAreaWidth = 489
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BOuvrir: TToolbarButton97
        Left = 394
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Ouvrir'
        DisplayMode = dmGlyphOnly
        Caption = 'Ouvrir'
        Flat = False
        Layout = blGlyphTop
        NumGlyphs = 2
        OnClick = BOuvrirClick
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object BAnnuler: TToolbarButton97
        Left = 426
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
        Layout = blGlyphTop
        ModalResult = 2
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Left = 458
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Flat = False
        Layout = blGlyphTop
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
      object BFiltre: TToolbarButton97
        Left = 6
        Top = 7
        Width = 61
        Height = 21
        Hint = 'Menu filtre'
        DropdownArrow = True
        DropdownMenu = POPF
        Caption = '&Filtres'
        Layout = blGlyphRight
        ParentShowHint = False
        ShowHint = True
      end
      object BAssist: TToolbarButton97
        Tag = 1
        Left = 377
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Assistant libell'#233' automatique'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BAssistClick
        GlobalIndexImage = 'Z0232_S16G1'
      end
      object FFiltres: THValComboBox
        Left = 73
        Top = 5
        Width = 313
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
      end
    end
  end
  object HMTrad: THSystemMenu
    ResizeDBGrid = True
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 244
    Top = 283
  end
  object POPF: TPopupMenu
    Left = 280
    Top = 283
    object BCreerFiltre: TMenuItem
      Caption = '&Cr'#233'er un filtre'
    end
    object BSaveFiltre: TMenuItem
      Caption = '&Enregistrer le filtre'
    end
    object Dupliquer1: TMenuItem
      Caption = '&Dupliquer le filtre'
    end
    object BDelFiltre: TMenuItem
      Caption = '&Supprimer le filtre'
    end
    object BRenFiltre: TMenuItem
      Caption = '&Renommer le filtre'
    end
    object BNouvRech: TMenuItem
      Caption = '&Nouvelle recherche'
    end
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?caption?;Vous devez renseigner une date valide !;W;O;O;O;'
      
        '1;?caption?;La date d'#39#233'ch'#233'ance doit respecter la plage de saisie' +
        ' autoris'#233'e;W;O;O;O;'
      
        '2;?caption?;Confirmez-vous la g'#233'n'#233'ration des mouvements ?;Q;YN;N' +
        ';N;'
      '3;?caption?;Le compte de g'#233'n'#233'ration n'#39'existe pas;W;O;O;O;'
      '4;?caption?;Le journal de g'#233'n'#233'ration n'#39'existe pas;W;O;O;O;'
      
        '5;?caption?;Ce mode de g'#233'n'#233'ration n'#39'est pas compatible avec un c' +
        'ompte de g'#233'n'#233'ration collectif;W;O;O;O;'
      
        '6;?caption?;Le compte g'#233'n'#233'ral n'#39'est pas la contrepartie du journ' +
        'al de banque;W;O;O;O;'
      
        '7;?caption?;Le journal de g'#233'n'#233'ration doit '#234'tre un journal de ban' +
        'que;W;O;O;O;'
      
        '8?caption?;Le compte de g'#233'n'#233'ration doit '#234'tre un compte de banque' +
        ';W;O;O;O;'
      '9;?caption?;Vous devez renseigner une date valide.;W;O;O;O;'
      
        '10;?caption?;La date de r'#232'glement que vous avez renseign'#233'e est s' +
        'ur un exercice non ouvert.;W;O;O;O;'
      
        '11;?caption?;La date de r'#232'glement que vous avez renseign'#233'e est s' +
        'ur un exercice non ouvert.;W;O;O;O;'
      
        '12;?caption?;La date de r'#232'glement que vous avez renseign'#233'e est a' +
        'nt'#233'rieure '#224' la cl'#244'ture provisoire.;W;O;O;O;'
      
        '13;?caption?;La date de r'#232'glement que vous avez renseign'#233'e est a' +
        'nt'#233'rieure '#224' la cl'#244'ture d'#233'finitive.;W;O;O;O;'
      '14;?caption?;La nature du compte g'#233'n'#233'ral est incorrecte;W;O;O;O;'
      
        '15;?caption?;Le compte de g'#233'n'#233'ration ne doit pas '#234'tre collectif;' +
        'W;O;O;O;'
      
        '16;?caption?;La formule saisie est trop longue. Tout ne sera pas' +
        ' retenu,vous devez la recomposer.;W;O;O;O;'
      
        '17;?caption?;Vous allez g'#233'n'#233'rer des '#233'critures de r'#232'glement en de' +
        'vise sans lettrage automatique. Voulez-vous continuer ?;Q;YN;N;N' +
        ';'
      '18;')
    Left = 315
    Top = 283
  end
end

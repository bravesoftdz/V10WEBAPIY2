object FGenerAbo: TFGenerAbo
  Left = 370
  Top = 181
  Width = 606
  Height = 390
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'G'#233'n'#233'ration des abonnements'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  PopupMenu = POPS
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 598
    Height = 117
    ActivePage = PCriteres
    Align = alTop
    TabOrder = 0
    object PCriteres: TTabSheet
      Caption = 'Crit'#232'res'
      object TG_CONTRAT: THLabel
        Left = 8
        Top = 10
        Width = 75
        Height = 13
        Caption = '&Abonnement de'
      end
      object TCB_RECONDUCTION: TLabel
        Left = 9
        Top = 39
        Width = 67
        Height = 13
        Caption = 'Rec&onduction'
        FocusControl = CB_RECONDUCTION
      end
      object HLabel1: THLabel
        Left = 204
        Top = 10
        Width = 6
        Height = 13
        Caption = #224
      end
      object TCB_QUALIFPIECE: THLabel
        Left = 8
        Top = 69
        Width = 70
        Height = 13
        Caption = '&Type d'#39#233'criture'
        Color = clBtnFace
        FocusControl = CB_QUALIFPIECE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object CB_CONTRAT: THCritMaskEdit
        Left = 88
        Top = 6
        Width = 109
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 0
        OnChange = CB_CONTRATChange
        TagDispatch = 0
        DataType = 'CPLISTEABO'
        Operateur = Superieur
      end
      object CB_CONTRAT_: THCritMaskEdit
        Left = 220
        Top = 6
        Width = 105
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 1
        OnChange = CB_CONTRATChange
        TagDispatch = 0
        DataType = 'CPLISTEABO'
        Operateur = Inferieur
      end
      object CB_RECONDUCTION: THValComboBox
        Left = 88
        Top = 35
        Width = 237
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
        OnChange = CB_CONTRATChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTRECONDUCABO'
      end
      object GroupBox1: TGroupBox
        Left = 332
        Top = 1
        Width = 249
        Height = 57
        Caption = 'P'#233'riode de g'#233'n'#233'ration des abonnements'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        object TE_DATECOMPTABLE: THLabel
          Left = 14
          Top = 21
          Width = 38
          Height = 13
          Caption = '&Date de'
          FocusControl = DATEDEBUT
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TE_DATECOMPTABLE2: THLabel
          Left = 139
          Top = 21
          Width = 6
          Height = 13
          Caption = #224
          FocusControl = DATEFIN
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object DATEDEBUT: THCritMaskEdit
          Left = 60
          Top = 17
          Width = 69
          Height = 21
          Ctl3D = True
          Enabled = False
          EditMask = '!99/99/0000;1;_'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 10
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          Text = '01/01/1900'
          OnChange = CB_CONTRATChange
          TagDispatch = 0
          Operateur = Superieur
          OpeType = otDate
          ControlerDate = True
        end
        object DATEFIN: THCritMaskEdit
          Left = 160
          Top = 17
          Width = 69
          Height = 21
          Ctl3D = True
          EditMask = '!99/99/0000;1;_'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 10
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 1
          Text = '31/12/2099'
          OnChange = CB_CONTRATChange
          OnExit = DATEFINExit
          TagDispatch = 0
          Operateur = Inferieur
          OpeType = otDate
          ControlerDate = True
        end
      end
      object CB_COMPTABLE: TCheckBox
        Left = 120
        Top = 15
        Width = 71
        Height = 12
        Caption = 'Comptable'
        Checked = True
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        State = cbChecked
        TabOrder = 4
        Visible = False
      end
      object CB_QUALIFPIECE: THValComboBox
        Left = 89
        Top = 65
        Width = 237
        Height = 21
        Style = csDropDownList
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 5
        OnChange = CB_QUALIFPIECEChange
        TagDispatch = 0
        DataType = 'TTQUALPIECE'
      end
    end
  end
  object GA: THGrid
    Tag = 1
    Left = 0
    Top = 155
    Width = 598
    Height = 168
    Align = alClient
    Ctl3D = True
    DefaultColWidth = 50
    DefaultRowHeight = 18
    FixedCols = 0
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
    ParentCtl3D = False
    TabOrder = 1
    OnDblClick = GADblClick
    OnMouseUp = GAMouseUp
    SortedCol = -1
    Couleur = False
    MultiSelect = True
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = False
    AlternateColor = 13224395
    ColWidths = (
      50
      50
      50
      50
      51)
  end
  object PModif: TPanel
    Left = 272
    Top = 200
    Width = 309
    Height = 108
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 2
    Visible = False
    object H_TitreCouv: TLabel
      Left = 4
      Top = 18
      Width = 301
      Height = 18
      Alignment = taCenter
      AutoSize = False
      Caption = 'Choix d'#39'une date maximum de g'#233'n'#233'ration'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = False
      ParentFont = False
    end
    object Label1: TLabel
      Left = 8
      Top = 67
      Width = 42
      Height = 13
      Caption = 'Contrat'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object H_CONTRAT: TLabel
      Left = 56
      Top = 67
      Width = 76
      Height = 13
      Caption = 'H_CONTRAT'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label2: TLabel
      Left = 24
      Top = 43
      Width = 91
      Height = 13
      Caption = 'Date limite retenue '
    end
    object H_LIBELLE: TLabel
      Left = 8
      Top = 87
      Width = 197
      Height = 13
      AutoSize = False
      Caption = 'H_LIBELLE'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object BCValide: TToolbarButton97
      Tag = 1
      Left = 244
      Top = 76
      Width = 28
      Height = 27
      Hint = 'Valider la date'
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
      OnClick = BCValideClick
      GlobalIndexImage = 'Z0003_S16G2'
      IsControl = True
    end
    object BCAbandon: TToolbarButton97
      Tag = 1
      Left = 276
      Top = 76
      Width = 28
      Height = 27
      Hint = 'Fermer'
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
      OnClick = BCAbandonClick
      GlobalIndexImage = 'Z0021_S16G1'
      IsControl = True
    end
    object PFenAbo: TPanel
      Left = 2
      Top = 2
      Width = 305
      Height = 13
      Align = alTop
      BevelInner = bvLowered
      BevelOuter = bvNone
      Color = clActiveCaption
      Enabled = False
      TabOrder = 0
    end
    object FComboDate: TComboBox
      Left = 128
      Top = 39
      Width = 117
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 1
    end
  end
  object Dock971: TDock97
    Left = 0
    Top = 117
    Width = 598
    Height = 38
    AllowDrag = False
    object PFiltres: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 34
      ClientWidth = 598
      Caption = 'PFiltres'
      ClientAreaHeight = 34
      ClientAreaWidth = 598
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BFiltre: TToolbarButton97
        Left = 6
        Top = 7
        Width = 56
        Height = 21
        Hint = 'Menu filtre'
        DropdownArrow = True
        DropdownMenu = POPF
        Caption = '&Filtres'
        Layout = blGlyphRight
        ParentShowHint = False
        ShowHint = True
      end
      object BCherche: TToolbarButton97
        Tag = 1
        Left = 560
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Appliquer crit'#232'res'
        ParentShowHint = False
        ShowHint = True
        OnClick = BChercheClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object FFiltres: THValComboBox
        Left = 68
        Top = 7
        Width = 481
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
      end
    end
  end
  object Dock972: TDock97
    Left = 0
    Top = 323
    Width = 598
    Height = 40
    AllowDrag = False
    Position = dpBottom
    object PPied: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 36
      ClientWidth = 598
      Caption = 'Actions'
      ClientAreaHeight = 36
      ClientAreaWidth = 598
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BReduire: TToolbarButton97
        Left = 4
        Top = 4
        Width = 28
        Height = 27
        Hint = 'R'#233'duire la liste'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BReduireClick
        GlobalIndexImage = 'Z0910_S16G1'
      end
      object BAgrandir: TToolbarButton97
        Left = 4
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Agrandir la liste'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAgrandirClick
        GlobalIndexImage = 'Z0270_S16G1'
      end
      object BValide: TToolbarButton97
        Tag = 1
        Left = 496
        Top = 4
        Width = 28
        Height = 27
        Hint = 'G'#233'n'#233'ration des abonnements'
        Enabled = False
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BValideClick
        GlobalIndexImage = 'Z0184_S16G1'
      end
      object BAnnuler: TToolbarButton97
        Tag = 1
        Left = 528
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAnnulerClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 560
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Aide'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
      object BImprimer: TToolbarButton97
        Tag = 1
        Left = 464
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Imprimer'
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
        OnClick = BImprimerClick
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object BRecherche: TToolbarButton97
        Left = 36
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Rechercher dans la liste'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BRechercheClick
        GlobalIndexImage = 'Z0077_S16G1'
      end
      object BZoomAbo: TToolbarButton97
        Tag = 1
        Left = 100
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Zoom abonnements'
        Caption = 'BZoomAbo'
        Flat = False
        Margin = 1
        ParentShowHint = False
        ShowHint = True
        OnClick = BZoomAboClick
        GlobalIndexImage = 'Z0061_S16G1'
      end
      object BListePIECES: TToolbarButton97
        Tag = 1
        Left = 132
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Liste des pi'#232'ces d'#39'abonnement g'#233'n'#233'r'#233'es'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 0
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BListePIECESClick
        GlobalIndexImage = 'Z1826_S16G1'
        IsControl = True
      end
      object BModifGenere: TToolbarButton97
        Tag = 1
        Left = 68
        Top = 4
        Width = 28
        Height = 27
        Hint = 'Modification du nombre de g'#233'n'#233'rations'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 0
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BModifGenereClick
        GlobalIndexImage = 'Z0041_S16G2'
        IsControl = True
      end
    end
  end
  object HGA: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'G'#233'n'#233'ration automatique des abonnements'
      'G'#233'n'#233'ration semi-automatique des abonnements'
      
        '2;?caption?;Vous devez renseigner une date de g'#233'n'#233'ration valide;' +
        'W;O;O;O;'
      
        '3;?caption?;La date de g'#233'n'#233'ration est sur un exercice non ouvert' +
        ';W;O;O;O;'
      
        '4;?caption?;La date de g'#233'n'#233'ration est sur un exercice non ouvert' +
        ';W;O;O;O;'
      
        '5;?caption?;La date de g'#233'n'#233'ration est ant'#233'rieure '#224' la cl'#244'ture pr' +
        'ovisoire;W;O;O;O;'
      
        '6;?caption?;La date de g'#233'n'#233'ration est ant'#233'rieure '#224' la cl'#244'ture d'#233 +
        'finitive;W;O;O;O;'
      '7;?caption?;Vous n'#39'avez s'#233'lectionn'#233' aucun abonnement;E;O;O;O;'
      
        '8;?caption?;La prochaine date de g'#233'n'#233'ration ne correspond pas '#224' ' +
        'la p'#233'riode. Voulez-vous faire une modification automatique de ce' +
        't abonnement ?;Q;YNC;Y;Y;'
      
        '9;?caption?;Confirmez-vous la g'#233'n'#233'ration des abonnements?;Q;YNC;' +
        'Y;Y;'
      
        '10;?caption?;Attention : cet abonnement ne sera pas g'#233'n'#233'r'#233';E;O;O' +
        ';O;'
      
        '11;?caption?;Voulez-vous voir la liste des pi'#232'ces g'#233'n'#233'r'#233'es ?;Q;Y' +
        'N;N;N;'
      
        '12;?caption?;Vous devez vous positionner sur un abonnement s'#233'lec' +
        'tionn'#233';W;O;O;O;'
      
        'ATTENTION! Certains abonnements en cours d'#39'utilisation n'#39'ont pas' +
        ' '#233't'#233' g'#233'n'#233'r'#233's'
      'G'#233'n'#233'ration manuelle des abonnements'
      'EURO'
      
        '16;?caption?;Vous n'#39'avez pas le droit de g'#233'n'#233'rer des '#233'critures d' +
        'e r'#233'vision;W;O;O;O;'
      
        'ATTENTION! Erreur lors de la mise '#224' jour d'#39'un compte g'#233'n'#233'ral. Ce' +
        'rtains abonnements en cours d'#39'utilisation n'#39'ont pas '#233't'#233' g'#233'n'#233'r'#233's'
      
        'ATTENTION! Erreur lors de la mise '#224' jour d'#39'un compte auxiliaire.' +
        ' Certains abonnements en cours d'#39'utilisation n'#39'ont pas '#233't'#233' g'#233'n'#233'r' +
        #233's'
      
        'ATTENTION! Erreur lors de la mise '#224' jour d'#39'un journal. Certains ' +
        'abonnements en cours d'#39'utilisation n'#39'ont pas '#233't'#233' g'#233'n'#233'r'#233's'
      '')
    Left = 140
    Top = 268
  end
  object POPS: TPopupMenu
    OnPopup = POPSPopup
    Left = 96
    Top = 268
  end
  object HMTrad: THSystemMenu
    LockedCtrls.Strings = (
      'PMODIF')
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 177
    Top = 268
  end
  object POPF: TPopupMenu
    Left = 220
    Top = 268
    object BCreerFiltre: TMenuItem
      Caption = '&Cr'#233'er un filtre'
    end
    object BSaveFiltre: TMenuItem
      Caption = '&Enregistrer le filtre'
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
  object FindMvt: THFindDialog
    OnFind = FindMvtFind
    Left = 48
    Top = 264
  end
end

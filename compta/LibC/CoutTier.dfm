object FCoutTiers: TFCoutTiers
  Left = 304
  Top = 118
  Width = 630
  Height = 399
  HelpContext = 7598000
  BorderIcons = [biSystemMenu, biMaximize]
  Caption = 'Estimation des co'#251'ts financiers'
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
  OnDestroy = FormDestroy
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 622
    Height = 109
    ActivePage = Princ
    Align = alTop
    TabOrder = 0
    object Princ: TTabSheet
      Caption = 'Comptes'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 614
        Height = 81
        Align = alClient
      end
      object HLabel4: THLabel
        Left = 13
        Top = 9
        Width = 61
        Height = 13
        Caption = '&G'#233'n'#233'raux de'
        FocusControl = E_GENERAL_SUP
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel5: THLabel
        Left = 189
        Top = 9
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_GENERAL_INF
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel2: THLabel
        Left = 189
        Top = 35
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_AUXILIAIRE_INF
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel1: THLabel
        Left = 13
        Top = 35
        Width = 61
        Height = 13
        Caption = '&Auxiliaires de'
        FocusControl = E_AUXILIAIRE_SUP
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel14: THLabel
        Left = 13
        Top = 61
        Width = 53
        Height = 13
        Caption = 'Total co'#251'ts'
      end
      object E_AUXILIAIRE_SUP: THCpteEdit
        Tag = 1
        Left = 89
        Top = 31
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 2
        ZoomTable = tzTToutDebit
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object rSelection: TRadioGroup
        Left = 317
        Top = 1
        Width = 284
        Height = 73
        Color = clBtnFace
        ItemIndex = 0
        Items.Strings = (
          '&Toutes les pi'#232'ces de la s'#233'lection'
          'A&u moins une '#233'ch'#233'ance non lettr'#233'e'
          'Au &moins une '#233'ch'#233'ance '#233'chue et non lettr'#233'e')
        ParentColor = False
        TabOrder = 4
      end
      object E_GENERAL_SUP: THCpteEdit
        Tag = 1
        Left = 89
        Top = 5
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 0
        ZoomTable = tzGCollectif
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object E_GENERAL_INF: THCpteEdit
        Tag = 1
        Left = 205
        Top = 5
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 1
        ZoomTable = tzGCollectif
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object E_AUXILIAIRE_INF: THCpteEdit
        Tag = 1
        Left = 205
        Top = 31
        Width = 93
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 3
        ZoomTable = tzTToutDebit
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object TotalGeneral: TPanel
        Left = 89
        Top = 58
        Width = 93
        Height = 20
        Alignment = taRightJustify
        BevelOuter = bvLowered
        TabOrder = 5
      end
    end
    object Complements: TTabSheet
      Caption = 'Ecritures'
      object Bevel2: TBevel
        Left = 0
        Top = 0
        Width = 614
        Height = 81
        Align = alClient
      end
      object TE_JOURNAL: THLabel
        Left = 16
        Top = 21
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = E_JOURNAL
      end
      object HLabel11: THLabel
        Left = 16
        Top = 50
        Width = 57
        Height = 13
        Caption = 'Nu&m'#233'ros de'
        FocusControl = E_NUMEROPIECE
      end
      object Label2: THLabel
        Left = 155
        Top = 49
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_NUMEROPIECE_
      end
      object HLabel8: THLabel
        Left = 495
        Top = 20
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATEECHEANCE_
      end
      object HLabel10: THLabel
        Left = 315
        Top = 20
        Width = 69
        Height = 13
        Caption = '&Ech'#233'ances du'
        FocusControl = E_DATEECHEANCE
      end
      object TE_DATECOMPTABLE: THLabel
        Left = 315
        Top = 49
        Width = 95
        Height = 13
        Caption = '&Dates comptable du'
        FocusControl = E_DATECOMPTABLE
      end
      object TE_DATECOMPTABLE2: THLabel
        Left = 495
        Top = 50
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATECOMPTABLE_
      end
      object E_QUALIFPIECE: TEdit
        Left = 280
        Top = 51
        Width = 20
        Height = 21
        TabStop = False
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 9
        Text = 'N'
        Visible = False
      end
      object E_JOURNAL: THValComboBox
        Left = 81
        Top = 17
        Width = 158
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTJALVENOD'
      end
      object E_NUMEROPIECE: THCritMaskEdit
        Tag = 2
        Left = 81
        Top = 47
        Width = 69
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 1
        Text = '0'
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_NUMEROPIECE_: THCritMaskEdit
        Tag = 2
        Left = 168
        Top = 47
        Width = 69
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 2
        Text = '99999999'
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_DATEECHEANCE_: THCritMaskEdit
        Left = 519
        Top = 16
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1; '
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 4
        Text = '31/12/2099'
        OnKeyPress = E_DATEECHEANCEKeyPress
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATEECHEANCE: THCritMaskEdit
        Left = 419
        Top = 16
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1; '
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 3
        Text = '01/01/1900'
        OnKeyPress = E_DATEECHEANCEKeyPress
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 419
        Top = 46
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1; '
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 5
        Text = '01/01/1900'
        OnKeyPress = E_DATEECHEANCEKeyPress
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATECOMPTABLE_: THCritMaskEdit
        Left = 519
        Top = 46
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1; '
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 6
        Text = '31/12/2099'
        OnKeyPress = E_DATEECHEANCEKeyPress
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_ECHE: TEdit
        Left = 278
        Top = 7
        Width = 20
        Height = 21
        TabStop = False
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 8
        Text = 'X'
        Visible = False
      end
      object E_NUMECHE: THCritMaskEdit
        Left = 279
        Top = 30
        Width = 21
        Height = 21
        TabStop = False
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 7
        Text = '1'
        Visible = False
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
    end
    object TabSheet1: TTabSheet
      Caption = 'Tiers'
      object Bevel4: TBevel
        Left = 0
        Top = 0
        Width = 614
        Height = 81
        Align = alClient
      end
      object TT_LIBELLE: THLabel
        Left = 13
        Top = 10
        Width = 71
        Height = 13
        Caption = '&Raison Sociale'
        FocusControl = T_LIBELLE
      end
      object TT_EAN: THLabel
        Left = 13
        Top = 34
        Width = 54
        Height = 13
        Caption = '&Code Client'
        FocusControl = T_EAN
      end
      object TT_ADRESSE1: THLabel
        Left = 13
        Top = 58
        Width = 38
        Height = 13
        Caption = 'A&dresse'
        FocusControl = T_ADRESSE1
      end
      object TT_CODEPOSTAL: THLabel
        Left = 321
        Top = 34
        Width = 57
        Height = 13
        Caption = 'Code &Postal'
        FocusControl = T_CODEPOSTAL
      end
      object TT_VILLE: THLabel
        Left = 457
        Top = 34
        Width = 19
        Height = 13
        Caption = '&Ville'
        FocusControl = T_VILLE
      end
      object TT_Secteur: THLabel
        Left = 323
        Top = 10
        Width = 37
        Height = 13
        Caption = '&Secteur'
        FocusControl = T_SECTEUR
      end
      object HLabel13: THLabel
        Left = 321
        Top = 58
        Width = 51
        Height = 13
        Caption = '&T'#233'l'#233'phone'
        FocusControl = T_TELEPHONE
      end
      object T_LIBELLE: TEdit
        Left = 99
        Top = 6
        Width = 173
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 0
      end
      object T_EAN: TEdit
        Left = 99
        Top = 30
        Width = 173
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 1
      end
      object T_ADRESSE1: TEdit
        Left = 99
        Top = 54
        Width = 173
        Height = 21
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 2
      end
      object T_CODEPOSTAL: TEdit
        Left = 395
        Top = 30
        Width = 50
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 4
      end
      object T_VILLE: TEdit
        Left = 481
        Top = 30
        Width = 121
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 5
      end
      object T_SECTEUR: THValComboBox
        Left = 395
        Top = 6
        Width = 207
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        Vide = True
        DataType = 'TTSECTEUR'
      end
      object T_TELEPHONE: TEdit
        Left = 395
        Top = 54
        Width = 207
        Height = 21
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 6
      end
    end
    object PLibres: TTabSheet
      Caption = 'Tables libres'
      object Bevel5: TBevel
        Left = 0
        Top = 0
        Width = 614
        Height = 81
        Align = alClient
      end
      object TT_TABLE0: THLabel
        Left = 25
        Top = 1
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE0'
      end
      object TT_TABLE1: THLabel
        Left = 140
        Top = 1
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE1'
      end
      object TT_TABLE2: THLabel
        Left = 255
        Top = 1
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE2'
      end
      object TT_TABLE3: THLabel
        Left = 369
        Top = 1
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE3'
      end
      object TT_TABLE4: THLabel
        Left = 481
        Top = 1
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE4'
      end
      object TT_TABLE5: THLabel
        Left = 25
        Top = 41
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE5'
      end
      object TT_TABLE6: THLabel
        Left = 140
        Top = 41
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE6'
      end
      object TT_TABLE7: THLabel
        Left = 255
        Top = 41
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE7'
      end
      object TT_TABLE8: THLabel
        Left = 369
        Top = 41
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE8'
      end
      object TT_TABLE9: THLabel
        Left = 481
        Top = 40
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE9'
      end
      object T_TABLE4: THCpteEdit
        Left = 481
        Top = 16
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 0
        ZoomTable = tzNatTiers4
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE3: THCpteEdit
        Left = 369
        Top = 16
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 1
        ZoomTable = tzNatTiers3
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE2: THCpteEdit
        Left = 255
        Top = 16
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 2
        ZoomTable = tzNatTiers2
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE1: THCpteEdit
        Left = 140
        Top = 16
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 3
        ZoomTable = tzNatTiers1
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE0: THCpteEdit
        Left = 25
        Top = 16
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 4
        ZoomTable = tzNatTiers0
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE5: THCpteEdit
        Left = 25
        Top = 54
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 5
        ZoomTable = tzNatTiers5
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE6: THCpteEdit
        Left = 140
        Top = 54
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 6
        ZoomTable = tzNatTiers6
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE7: THCpteEdit
        Left = 255
        Top = 54
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 7
        ZoomTable = tzNatTiers7
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE8: THCpteEdit
        Left = 369
        Top = 54
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 8
        ZoomTable = tzNatTiers8
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE9: THCpteEdit
        Left = 481
        Top = 54
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 9
        ZoomTable = tzNatTiers9
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'Param'#232'tres'
      object Bevel3: TBevel
        Left = 0
        Top = 0
        Width = 614
        Height = 81
        Align = alClient
      end
      object HLabel6: THLabel
        Left = 333
        Top = 34
        Width = 186
        Height = 13
        Caption = '&Date de r'#233'f'#233'rence pour calcul du retard'
        FocusControl = DateReference
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel3: THLabel
        Left = 8
        Top = 11
        Width = 241
        Height = 13
        Caption = 'Taux '#224' appliquer pour le d'#233'lai de paiement &n'#233'goci'#233
        FocusControl = TauxNormal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel7: THLabel
        Left = 8
        Top = 34
        Width = 215
        Height = 13
        Caption = 'Taux '#224' appliquer pour les &retards de paiement'
        FocusControl = TauxRetard
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel9: THLabel
        Left = 304
        Top = 12
        Width = 8
        Height = 13
        Alignment = taCenter
        Caption = '%'
      end
      object HLabel12: THLabel
        Left = 304
        Top = 34
        Width = 8
        Height = 13
        Alignment = taCenter
        Caption = '%'
      end
      object TRacine: THLabel
        Left = 333
        Top = 11
        Width = 149
        Height = 13
        Caption = 'Racine de comptes d'#39'&escompte'
        FocusControl = RacineEscompte
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TSeuil1: THLabel
        Left = 8
        Top = 58
        Width = 68
        Height = 13
        Caption = '&Co'#251't tiers >= '#224
        FocusControl = MontantSeuil
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TSeuil2: THLabel
        Left = 191
        Top = 58
        Width = 56
        Height = 13
        Caption = '% co'#251't >= '#224
        FocusControl = RatioSeuil
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel20: THLabel
        Left = 304
        Top = 58
        Width = 8
        Height = 13
        Alignment = taCenter
        Caption = '%'
      end
      object cAvoirs: TCheckBox
        Left = 340
        Top = 58
        Width = 56
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Soustraire les &avoirs dans le calcul des co'#251'ts'
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 6
        Visible = False
      end
      object DateReference: THCritMaskEdit
        Left = 526
        Top = 31
        Width = 74
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1; '
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 5
        Text = '31/12/2099'
        OnExit = DateReferenceExit
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object TauxNormal: THNumEdit
        Left = 255
        Top = 7
        Width = 44
        Height = 21
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '##0.#0'
        Debit = False
        TabOrder = 0
        UseRounding = True
        Validate = False
        OnEnter = TauxNormalEnter
      end
      object TauxRetard: THNumEdit
        Left = 255
        Top = 31
        Width = 44
        Height = 21
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '##0.#0'
        Debit = False
        TabOrder = 1
        UseRounding = True
        Validate = False
        OnEnter = TauxRetardEnter
      end
      object RacineEscompte: THCritMaskEdit
        Tag = 1
        Left = 526
        Top = 7
        Width = 74
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 4
        Text = '665'
        TagDispatch = 0
        Operateur = Superieur
      end
      object MontantSeuil: THNumEdit
        Left = 88
        Top = 55
        Width = 91
        Height = 21
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0.#0'
        Debit = False
        TabOrder = 2
        UseRounding = True
        Validate = False
        OnEnter = MontantSeuilEnter
      end
      object RatioSeuil: THNumEdit
        Left = 255
        Top = 55
        Width = 44
        Height = 21
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '##0.#0'
        Debit = False
        TabOrder = 3
        UseRounding = True
        Validate = False
        OnEnter = RatioSeuilEnter
      end
    end
  end
  object GT: THGrid
    Left = 0
    Top = 148
    Width = 622
    Height = 185
    Align = alClient
    ColCount = 12
    DefaultRowHeight = 18
    FixedCols = 0
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSizing, goColSizing, goRowSelect]
    TabOrder = 2
    OnDblClick = GTDblClick
    SortedCol = -1
    Titres.Strings = (
      ';;S'
      ';;S'
      ';;R'
      ';;R'
      ';;R'
      ';;R'
      ';;R'
      ';;R')
    Couleur = False
    MultiSelect = False
    TitleBold = False
    TitleCenter = True
    ColCombo = 0
    SortEnabled = False
    SortRowExclude = 0
    TwoColors = True
    AlternateColor = 13224395
  end
  object Dock971: TDock97
    Left = 0
    Top = 109
    Width = 622
    Height = 39
    AllowDrag = False
    object PFiltres: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 35
      ClientWidth = 622
      Caption = 'PFiltres'
      ClientAreaHeight = 35
      ClientAreaWidth = 622
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object TT_FILTRESCRITERES: TLabel
        Left = 72
        Top = 8
        Width = 27
        Height = 13
        Caption = '&Filtres'
        FocusControl = FFiltres
        Visible = False
      end
      object BFiltre: TToolbarButton97
        Left = 5
        Top = 6
        Width = 59
        Height = 21
        Hint = 'Menu filtre'
        DropdownArrow = True
        DropdownMenu = POPF
        Caption = '&Filtres'
        Layout = blGlyphRight
        ParentShowHint = False
        ShowHint = True
      end
      object BChercher: TToolbarButton97
        Left = 588
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Appliquer crit'#232'res'
        ParentShowHint = False
        ShowHint = True
        OnClick = BChercherClick
        GlobalIndexImage = 'Z0217_S16G1'
      end
      object FFiltres: THValComboBox
        Left = 68
        Top = 5
        Width = 513
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
      end
    end
  end
  object BZoomTiers: THBitBtn
    Tag = 100
    Left = 424
    Top = 280
    Width = 28
    Height = 27
    Hint = 'Voir le tiers'
    Caption = 'Tiers'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 3
    Visible = False
    OnClick = BZoomTiersClick
    Margin = 2
    NumGlyphs = 2
    Spacing = -1
    IsControl = True
  end
  object bZoomMvts: THBitBtn
    Tag = 100
    Left = 429
    Top = 248
    Width = 28
    Height = 27
    Hint = 'D'#233'tail des mouvements'
    Caption = 'Tiers'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 5
    Visible = False
    OnClick = bZoomMvtsClick
    Margin = 2
    NumGlyphs = 2
    Spacing = -1
    IsControl = True
  end
  object bGraph1: THBitBtn
    Tag = 99
    Left = 435
    Top = 0
    Width = 27
    Height = 27
    Hint = 'R'#233'partition des co'#251'ts en pourcentage'
    Caption = 'G1'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 6
    Visible = False
    OnClick = bGraph1Click
  end
  object bGraph2: THBitBtn
    Tag = 99
    Left = 464
    Top = 0
    Width = 27
    Height = 27
    Hint = 'R'#233'partition des co'#251'ts en montant'
    Caption = 'G2'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 7
    Visible = False
    OnClick = bGraph2Click
  end
  object PanelPieces: TPanel
    Left = 22
    Top = 131
    Width = 516
    Height = 285
    TabOrder = 4
    Visible = False
    object TitrePanelPieces: TPanel
      Left = 1
      Top = 1
      Width = 514
      Height = 18
      Align = alTop
      BevelInner = bvLowered
      BevelOuter = bvNone
      Caption = 'D'#233'tail des co'#251'ts par pi'#232'ce'
      Color = clActiveCaption
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWhite
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
    end
    object Panel1: TPanel
      Left = 7
      Top = 114
      Width = 320
      Height = 31
      BevelOuter = bvNone
      TabOrder = 2
    end
    object Panel2: TPanel
      Left = 1
      Top = 250
      Width = 514
      Height = 34
      Align = alBottom
      BevelOuter = bvNone
      TabOrder = 3
      object RatioMoyenP: THLabel
        Left = 167
        Top = 11
        Width = 47
        Height = 13
        AutoSize = False
      end
      object HLabel18: THLabel
        Left = 6
        Top = 11
        Width = 53
        Height = 13
        Caption = 'Total co'#251'ts'
      end
      object bImprimerP: THBitBtn
        Left = 416
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Imprimer le co'#251't financier des tiers'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnClick = bImprimerPClick
        Margin = 2
        GlobalIndexImage = 'Z0369_S16G1'
      end
      object bValideP: THBitBtn
        Left = 448
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Voir la pi'#232'ce'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = bValidePClick
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object bFermerPieces: THBitBtn
        Left = 480
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = bFermerPClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object TotalGeneralP: TPanel
        Left = 64
        Top = 8
        Width = 100
        Height = 20
        Alignment = taRightJustify
        BevelOuter = bvLowered
        TabOrder = 3
      end
    end
    object GP: THGrid
      Left = 1
      Top = 19
      Width = 514
      Height = 231
      Align = alClient
      DefaultRowHeight = 18
      FixedCols = 0
      RowCount = 15
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
      TabOrder = 0
      OnDblClick = GPDblClick
      SortedCol = -1
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = 13224395
    end
  end
  object PanelStats: TPanel
    Left = 96
    Top = 139
    Width = 444
    Height = 301
    TabOrder = 9
    Visible = False
    object HLabel15: THLabel
      Left = 8
      Top = 276
      Width = 74
      Height = 13
      Caption = 'Nombre de tiers'
    end
    object Panel4: TPanel
      Left = 1
      Top = 1
      Width = 442
      Height = 18
      Align = alTop
      BevelInner = bvLowered
      BevelOuter = bvNone
      Caption = 'Statistiques'
      Color = clActiveCaption
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWhite
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
    end
    object GS: THGrid
      Left = 6
      Top = 24
      Width = 431
      Height = 118
      DefaultColWidth = 90
      DefaultRowHeight = 18
      DefaultDrawing = False
      Enabled = False
      RowCount = 6
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine]
      ScrollBars = ssNone
      TabOrder = 1
      OnDrawCell = GSDrawCell
      SortedCol = -1
      Titres.Strings = (
        ';G;'
        ';D;'
        ';D;'
        ';D;'
        ';D;'
        ';D;'
        ';D;'
        '')
      Couleur = False
      MultiSelect = False
      TitleBold = False
      TitleCenter = True
      ColCombo = 0
      SortEnabled = False
      SortRowExclude = 0
      TwoColors = True
      AlternateColor = 13224395
      ColWidths = (
        90
        99
        54
        90
        90)
      RowHeights = (
        18
        18
        18
        18
        18
        18)
    end
    object bFermerS: THBitBtn
      Left = 410
      Top = 269
      Width = 28
      Height = 27
      Hint = 'Fermer'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 6
      OnClick = bFermerSClick
      GlobalIndexImage = 'Z0021_S16G1'
    end
    object bImprimerS: THBitBtn
      Left = 378
      Top = 269
      Width = 28
      Height = 27
      Hint = 'Imprimer les statistiques'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 5
      OnClick = bImprimerSClick
      Margin = 2
      GlobalIndexImage = 'Z0369_S16G1'
    end
    object Chart1: TChart
      Left = 7
      Top = 145
      Width = 214
      Height = 120
      AllowPanning = pmNone
      AllowZoom = False
      BackWall.Brush.Color = clWhite
      BackWall.Brush.Style = bsClear
      BackWall.Pen.Visible = False
      Foot.Font.Charset = DEFAULT_CHARSET
      Foot.Font.Color = clRed
      Foot.Font.Height = -12
      Foot.Font.Name = 'Arial'
      Foot.Font.Style = [fsItalic]
      Foot.Visible = False
      Title.Font.Charset = DEFAULT_CHARSET
      Title.Font.Color = clBlack
      Title.Font.Height = -11
      Title.Font.Name = 'MS Sans Serif'
      Title.Font.Style = []
      Title.Text.Strings = (
        'R'#233'partition des co'#251'ts'
        'en montant')
      AxisVisible = False
      BottomAxis.Visible = False
      Chart3DPercent = 10
      ClipPoints = False
      Frame.Visible = False
      LeftAxis.Visible = False
      Legend.Font.Charset = DEFAULT_CHARSET
      Legend.Font.Color = clBlack
      Legend.Font.Height = -11
      Legend.Font.Name = 'MS Sans Serif'
      Legend.Font.Style = []
      Legend.LegendStyle = lsValues
      Legend.ShadowColor = clGray
      Legend.TextStyle = ltsRightValue
      RightAxis.Visible = False
      TopAxis.Visible = False
      View3DOptions.Elevation = 315
      View3DOptions.Orthogonal = False
      View3DOptions.Perspective = 0
      View3DOptions.Rotation = 360
      View3DWalls = False
      BevelOuter = bvLowered
      ParentShowHint = False
      ShowHint = False
      TabOrder = 2
      object Pie: TPieSeries
        Marks.ArrowLength = 8
        Marks.Style = smsLabelPercent
        Marks.Visible = False
        SeriesColor = clRed
        OtherSlice.Text = 'Other'
        PieValues.DateTime = False
        PieValues.Name = 'Y'
        PieValues.Multiplier = 1.000000000000000000
        PieValues.Order = loNone
        Left = 257
        Top = 277
      end
    end
    object Chart2: TChart
      Left = 225
      Top = 145
      Width = 213
      Height = 120
      AllowPanning = pmNone
      AllowZoom = False
      BackWall.Brush.Color = clWhite
      BackWall.Brush.Style = bsClear
      BackWall.Pen.Visible = False
      Foot.Font.Charset = DEFAULT_CHARSET
      Foot.Font.Color = clRed
      Foot.Font.Height = -12
      Foot.Font.Name = 'Arial'
      Foot.Font.Style = [fsItalic]
      Foot.Visible = False
      Title.Font.Charset = DEFAULT_CHARSET
      Title.Font.Color = clBlack
      Title.Font.Height = -11
      Title.Font.Name = 'MS Sans Serif'
      Title.Font.Style = []
      Title.Text.Strings = (
        'R'#233'partition des co'#251'ts'
        'en % du total')
      AxisVisible = False
      BottomAxis.Visible = False
      Chart3DPercent = 10
      ClipPoints = False
      Frame.Visible = False
      LeftAxis.Visible = False
      Legend.Font.Charset = DEFAULT_CHARSET
      Legend.Font.Color = clBlack
      Legend.Font.Height = -11
      Legend.Font.Name = 'MS Sans Serif'
      Legend.Font.Style = []
      Legend.LegendStyle = lsValues
      Legend.ShadowColor = clGray
      Legend.TextStyle = ltsRightValue
      RightAxis.Visible = False
      TopAxis.Visible = False
      View3DOptions.Elevation = 315
      View3DOptions.Orthogonal = False
      View3DOptions.Perspective = 0
      View3DOptions.Rotation = 360
      View3DWalls = False
      BevelOuter = bvLowered
      ParentShowHint = False
      ShowHint = False
      TabOrder = 3
      object Pie2: TPieSeries
        Marks.ArrowLength = 8
        Marks.Style = smsLabelPercent
        Marks.Visible = False
        SeriesColor = clRed
        OtherSlice.Text = 'Other'
        PieValues.DateTime = False
        PieValues.Name = 'Y'
        PieValues.Multiplier = 1.000000000000000000
        PieValues.Order = loNone
        Left = 289
        Top = 277
      end
    end
    object NbTiers: TPanel
      Left = 89
      Top = 272
      Width = 53
      Height = 22
      BevelOuter = bvLowered
      Caption = ' '
      TabOrder = 4
    end
  end
  object Dock: TDock97
    Left = 0
    Top = 333
    Width = 622
    Height = 39
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 35
      ClientWidth = 622
      Caption = 'Actions'
      ClientAreaHeight = 35
      ClientAreaWidth = 622
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BReduire: TToolbarButton97
        Left = 5
        Top = 2
        Width = 28
        Height = 27
        Hint = 'R'#233'duire la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'R'#233'duire'
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
        Left = 5
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Agrandir la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'Agrandir'
        Enabled = False
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
      object BRecherche: TToolbarButton97
        Tag = 1
        Left = 36
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Rechercher dans la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'Chercher'
        Enabled = False
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BRechercheClick
        GlobalIndexImage = 'Z0077_S16G1'
      end
      object BMenuZoom: TToolbarButton97
        Tag = -100
        Left = 67
        Top = 2
        Width = 37
        Height = 27
        Hint = 'Menu zoom'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = POPZ
        Caption = 'Zooms'
        Enabled = False
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
        OnMouseEnter = BMenuZoomMouseEnter
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
      object BStop: TToolbarButton97
        Left = 460
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Stopper le traitement'
        DisplayMode = dmGlyphOnly
        Caption = 'Stop'
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
        Visible = False
        OnClick = BStopClick
        GlobalIndexImage = 'Z0107_S16G1'
        IsControl = True
      end
      object BGraph: TToolbarButton97
        Tag = -99
        Left = 107
        Top = 2
        Width = 37
        Height = 27
        Hint = 'Graphiques'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = POPZ
        Caption = 'Graphiques'
        Enabled = False
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
        OnMouseEnter = BGraphMouseEnter
        GlobalIndexImage = 'Z2172_S16G1'
      end
      object bStats: TToolbarButton97
        Tag = 1
        Left = 147
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Statistiques'
        DisplayMode = dmGlyphOnly
        Caption = 'Statistiques'
        Enabled = False
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = bStatsClick
        GlobalIndexImage = 'Z0797_S16G1'
      end
      object BImprimer: TToolbarButton97
        Tag = 1
        Left = 492
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Imprimer le co'#251't financier des tiers'
        DisplayMode = dmGlyphOnly
        Caption = 'Imprimer'
        Enabled = False
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
      object BDetail: TToolbarButton97
        Tag = 1
        Left = 524
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Afficher le d'#233'tail des co'#251'ts'
        DisplayMode = dmGlyphOnly
        Caption = 'D'#233'tail'
        Enabled = False
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BDetailClick
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object BAnnuler: TToolbarButton97
        Tag = 1
        Left = 556
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAnnulerClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 588
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
    end
  end
  object FindMvt: THFindDialog
    OnFind = FindMvtFind
    Left = 578
    Top = 179
  end
  object POPS: TPopupMenu
    AutoPopup = False
    OnPopup = POPSPopup
    Left = 592
    Top = 242
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Estimation des co'#251'ts financiers;Les param'#232'tres sont incomplets' +
        '.;W;O;O;O;'
      
        '1;Estimation des co'#251'ts financiers;Aucune information ne correspo' +
        'nd '#224' la s'#233'lection.;E;O;O;O;'
      'Total'
      '%'
      'Moyenne'
      'Ecart-type'
      'Factur'#233
      'Co'#251't'
      '  dont escompte'
      '  dont d'#233'lai normal'
      '  dont retard'
      'D'#233'tail des co'#251'ts de '
      'Autres tiers'
      'Co'#251't < '
      ' et < ')
    Left = 539
    Top = 202
  end
  object POPZ: TPopupMenu
    Left = 543
    Top = 248
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 8
    Top = 212
  end
  object POPF: TPopupMenu
    Left = 572
    Top = 292
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
end

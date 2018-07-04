object FRappAuto: TFRappAuto
  Left = 259
  Top = 210
  HelpContext = 7514000
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Lettrage automatique'
  ClientHeight = 327
  ClientWidth = 590
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Dock: TDock97
    Left = 0
    Top = 290
    Width = 590
    Height = 37
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 33
      ClientWidth = 590
      Caption = 'Actions'
      ClientAreaHeight = 33
      ClientAreaWidth = 590
      DockPos = 0
      FullSize = True
      TabOrder = 0
      object BValide: TToolbarButton97
        Tag = 1
        Left = 490
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Lancer le rapprochement automatique'
        DisplayMode = dmGlyphOnly
        Caption = 'Rapprocher'
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
        OnClick = BValideClick
        GlobalIndexImage = 'Z0184_S16G1'
        IsControl = True
      end
      object BAbandon: TToolbarButton97
        Tag = 1
        Left = 522
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Caption = 'Fermer'
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
        OnClick = BAbandonClick
        GlobalIndexImage = 'Z0021_S16G1'
        IsControl = True
      end
      object BAide: TToolbarButton97
        Tag = 1
        Left = 554
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Aide'
        DisplayMode = dmGlyphOnly
        Caption = 'Aide'
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
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BStop: TToolbarButton97
        Tag = 1
        Left = 66
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Arr'#234'ter le rapprochement automatique'
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
        OnClick = BStopClick
        GlobalIndexImage = 'Z0107_S16G1'
        IsControl = True
      end
      object BGenere: TToolbarButton97
        Tag = 1
        Left = 34
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Liste des rapprochements g'#233'n'#233'r'#233's'
        DisplayMode = dmGlyphOnly
        Caption = 'Liste'
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
        OnClick = BGenereClick
        GlobalIndexImage = 'Z1826_S16G1'
        IsControl = True
      end
      object BParam: TToolbarButton97
        Tag = 1
        Left = 2
        Top = 3
        Width = 28
        Height = 27
        Hint = 'Zoom param'#233'trage fiche soci'#233't'#233
        DisplayMode = dmGlyphOnly
        Caption = 'Param'#232'tres'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BParamClick
        GlobalIndexImage = 'Z0008_S16G1'
        IsControl = True
      end
    end
  end
  object PParam: TPanel
    Left = 0
    Top = 136
    Width = 590
    Height = 154
    Align = alClient
    TabOrder = 1
    object GParam: TGroupBox
      Left = 4
      Top = 2
      Width = 581
      Height = 130
      Caption = 'Param'#232'tres de rapprochement'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      TabOrder = 0
      object H_MaxProf: THLabel
        Left = 310
        Top = 51
        Width = 202
        Height = 13
        Caption = '&Niveau maximum de profondeur du lettrage'
        Enabled = False
        FocusControl = MAXPROF
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object H_MaxDuree: THLabel
        Left = 310
        Top = 78
        Width = 214
        Height = 13
        Caption = 'Secondes maximum de &traitement par compte'
        Enabled = False
        FocusControl = MAXDUREE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel3: THLabel
        Left = 8
        Top = 104
        Width = 58
        Height = 13
        Caption = 'R'#233'f. lettra&ge'
        FocusControl = E_REFLETTRAGE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object CMontant: TCheckBox
        Left = 6
        Top = 49
        Width = 274
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Rapprochement combinatoire sur les &montants'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnClick = CMontantClick
      end
      object CSolde: TCheckBox
        Left = 6
        Top = 76
        Width = 274
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Lettrage sur passage '#224' &solde nul'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
      end
      object CRef: TCheckBox
        Left = 6
        Top = 22
        Width = 274
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Rapprochement sur les &r'#233'f'#233'rences'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        OnClick = CRefClick
      end
      object MAXPROF: TSpinEdit
        Left = 530
        Top = 47
        Width = 41
        Height = 22
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxValue = 10
        MinValue = 1
        ParentFont = False
        TabOrder = 5
        Value = 1
      end
      object MAXDUREE: TSpinEdit
        Left = 530
        Top = 75
        Width = 41
        Height = 22
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxValue = 360
        MinValue = 5
        ParentFont = False
        TabOrder = 4
        Value = 5
      end
      object CODSALFOU: TCheckBox
        Left = 310
        Top = 103
        Width = 260
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Utiliser le param'#233'trage fourn. en  OD salaire'
        Checked = True
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        State = cbChecked
        TabOrder = 6
      end
      object E_REFLETTRAGE: TEdit
        Left = 88
        Top = 101
        Width = 193
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 35
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 3
      end
      object CAnc: TCheckBox
        Left = 310
        Top = 22
        Width = 256
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Rapprochement par anciennet'#233
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 7
        OnClick = CRefClick
      end
    end
  end
  object Pages: TPageControl
    Left = 0
    Top = 0
    Width = 590
    Height = 136
    ActivePage = PCrit
    Align = alTop
    TabOrder = 2
    object PCrit: TTabSheet
      Caption = 'Crit'#232'res'
      object Bevel1: TBevel
        Left = 0
        Top = 0
        Width = 582
        Height = 108
        Align = alClient
      end
      object HLabel4: THLabel
        Left = 8
        Top = 14
        Width = 61
        Height = 13
        Caption = '&G'#233'n'#233'raux de'
        FocusControl = E_GENERAL
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel5: THLabel
        Left = 176
        Top = 14
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_GENERAL_
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object Label14: THLabel
        Left = 292
        Top = 16
        Width = 63
        Height = 13
        Caption = 'Natures &Tiers'
        FocusControl = T_NATUREAUXI
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel1: THLabel
        Left = 8
        Top = 46
        Width = 61
        Height = 13
        Caption = '&Auxiliaires de'
        FocusControl = E_AUXILIAIRE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel2: THLabel
        Left = 176
        Top = 46
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_AUXILIAIRE_
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object E_GENERAL: THCpteEdit
        Left = 80
        Top = 12
        Width = 88
        Height = 21
        CharCase = ecUpperCase
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        ZoomTable = tzGLettColl
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object E_GENERAL_: THCpteEdit
        Left = 192
        Top = 12
        Width = 88
        Height = 21
        CharCase = ecUpperCase
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        ZoomTable = tzGLettColl
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object T_NATUREAUXI: THValComboBox
        Left = 364
        Top = 12
        Width = 205
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATTIERSCPTA'
      end
      object E_AUXILIAIRE: THCpteEdit
        Left = 80
        Top = 44
        Width = 88
        Height = 21
        CharCase = ecUpperCase
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        ZoomTable = tzTlettrable
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object E_AUXILIAIRE_: THCpteEdit
        Left = 192
        Top = 44
        Width = 88
        Height = 21
        CharCase = ecUpperCase
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        ZoomTable = tzTlettrable
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object LETTREPARTIEL: TCheckBox
        Left = 290
        Top = 46
        Width = 277
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Prendre en compte les '#233'l'#233'ments &partiellement lettr'#233's'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 5
      end
      object CDEV: THValComboBox
        Left = 338
        Top = 68
        Width = 37
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 6
        Text = 'CDEV'
        Visible = False
        TagDispatch = 0
      end
      object CAvecOD: TCheckBox
        Left = 6
        Top = 70
        Width = 273
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Inclure aussi les mouvements de nature OD'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 7
        OnClick = CAvecODClick
      end
    end
    object TSEcritures: TTabSheet
      Caption = 'Ecritures'
      ImageIndex = 2
      object TE_JOURNAL: THLabel
        Left = 5
        Top = 14
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = E_JOURNAL
      end
      object TE_ETABLISSEMENT: THLabel
        Left = 268
        Top = 14
        Width = 65
        Height = 13
        Caption = '&Etablissement'
        FocusControl = E_ETABLISSEMENT
      end
      object TE_DEVISE: THLabel
        Left = 5
        Top = 47
        Width = 33
        Height = 13
        Caption = '&Devise'
        FocusControl = E_DEVISE
      end
      object TE_REFINTERNE: THLabel
        Left = 268
        Top = 47
        Width = 50
        Height = 13
        Caption = '&R'#233'f'#233'rence'
        FocusControl = E_REFINTERNE
      end
      object TE_EXERCICE: THLabel
        Left = 5
        Top = 82
        Width = 41
        Height = 13
        Caption = 'E&xercice'
        FocusControl = E_EXERCICE
      end
      object TE_DATECOMPTABLE: THLabel
        Left = 268
        Top = 82
        Width = 100
        Height = 13
        Caption = '&Dates comptables du'
        FocusControl = E_DATECOMPTABLE
      end
      object TE_DATECOMPTABLE2: THLabel
        Left = 466
        Top = 82
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATECOMPTABLE_
      end
      object E_JOURNAL: THValComboBox
        Left = 53
        Top = 10
        Width = 183
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTJOURNAL'
      end
      object E_ETABLISSEMENT: THValComboBox
        Left = 379
        Top = 10
        Width = 193
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        Vide = True
        DataType = 'TTETABLISSEMENT'
      end
      object E_DEVISE: THValComboBox
        Left = 53
        Top = 43
        Width = 183
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 1
        TagDispatch = 0
        Vide = True
        DataType = 'TTDEVISE'
      end
      object E_REFINTERNE: TEdit
        Left = 379
        Top = 43
        Width = 193
        Height = 21
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 3
      end
      object E_EXERCICE: THValComboBox
        Left = 53
        Top = 78
        Width = 183
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 4
        OnChange = E_EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 379
        Top = 78
        Width = 65
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 5
        Text = '  /  /    '
        OnKeyPress = E_DATECOMPTABLEKeyPress
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATECOMPTABLE_: THCritMaskEdit
        Left = 507
        Top = 78
        Width = 65
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 6
        Text = '  /  /    '
        OnKeyPress = E_DATECOMPTABLEKeyPress
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
    end
    object PLibres: TTabSheet
      Caption = 'Tables libres'
      object Bevel2: TBevel
        Left = 0
        Top = 0
        Width = 583
        Height = 108
        Align = alClient
      end
      object TT_TABLE0: THLabel
        Left = 3
        Top = 11
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE0'
      end
      object TT_TABLE1: THLabel
        Left = 115
        Top = 11
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE1'
      end
      object TT_TABLE2: THLabel
        Left = 227
        Top = 11
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE2'
      end
      object TT_TABLE3: THLabel
        Left = 338
        Top = 11
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE3'
      end
      object TT_TABLE4: THLabel
        Left = 450
        Top = 11
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE4'
      end
      object TT_TABLE5: THLabel
        Left = 3
        Top = 61
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE5'
      end
      object TT_TABLE6: THLabel
        Left = 115
        Top = 61
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE6'
      end
      object TT_TABLE7: THLabel
        Left = 227
        Top = 61
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE7'
      end
      object TT_TABLE8: THLabel
        Left = 338
        Top = 61
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE8'
      end
      object TT_TABLE9: THLabel
        Left = 450
        Top = 60
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE9'
      end
      object T_TABLE4: THCpteEdit
        Left = 450
        Top = 26
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
        Left = 338
        Top = 26
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
        Left = 227
        Top = 26
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
        Left = 115
        Top = 26
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
        Left = 3
        Top = 26
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
        Left = 3
        Top = 72
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
        Left = 115
        Top = 72
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
        Left = 227
        Top = 72
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
        Left = 338
        Top = 72
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
        Left = 450
        Top = 72
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
  end
  object HRappro: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Lettrage automatique;Confirmez-vous le traitement automatique ' +
        'de lettrage ?;Q;YN;Y;Y;'
      
        '1;Lettrage automatique;Il n'#39'existe aucun '#233'l'#233'ment '#224' lettrer pour ' +
        'ces crit'#232'res;E;O;O;O;'
      
        '2;Lettrage automatique;Confirmez-vous l'#39'arr'#234't du traitement auto' +
        'matique?;Q;YN;Y;Y;'
      
        '3;Lettrage automatique;Vous devez choisir au moins une m'#233'thode d' +
        'e rapprochement;W;O;O;O;'
      
        '4;Lettrage automatique ;Voulez-vous voir le compte-rendu du trai' +
        'tement ?;Q;YN;Y;Y;'
      
        'ATTENTION! Certaines '#233'critures en cours de traitement n'#39'ont pas ' +
        #233't'#233' modifi'#233'es'
      
        '6;Lettrage automatique ;Traitement impossible. Vous devez rensei' +
        'gner le param'#233'trage dans la fiche soci'#233't'#233';W;O;O;O;'
      'EURO'
      'Ecritures en devise ou d'#39'origine'
      'Ecritures d'#39'origine'
      'Ecritures de toutes origines ou en devise'
      '11;')
    Left = 440
    Top = 73
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 65496
    Top = 144
  end
end

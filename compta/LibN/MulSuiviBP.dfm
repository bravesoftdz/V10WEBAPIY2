inherited FMulSuiviBP: TFMulSuiviBP
  Left = 196
  Top = 193
  Width = 638
  Height = 386
  Caption = 'Ecriture Comptable'
  PixelsPerInch = 96
  TextHeight = 13
  inherited Pages: THPageControl2
    Width = 630
    Height = 117
    ActivePage = PCritere
    inherited PCritere: THTabSheet
      inherited Bevel1: TBevel
        Width = 622
        Height = 89
      end
      object HLabel4: THLabel
        Left = 255
        Top = 66
        Width = 37
        Height = 13
        Caption = '&G'#233'n'#233'ral'
        FocusControl = E_GENERAL
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TE_AUXILIAIRE: THLabel
        Left = 8
        Top = 64
        Width = 41
        Height = 13
        Caption = '&Auxiliaire'
        FocusControl = E_AUXILIAIRE
      end
      object TE_EXERCICE: THLabel
        Left = 10
        Top = 10
        Width = 41
        Height = 13
        Caption = 'E&xercice'
        FocusControl = E_EXERCICE
      end
      object TE_DATECOMPTABLE: THLabel
        Left = 255
        Top = 10
        Width = 103
        Height = 13
        Caption = '&Dates comptables  du'
        FocusControl = E_DATECOMPTABLE
      end
      object TE_DATECOMPTABLE2: THLabel
        Left = 468
        Top = 10
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATECOMPTABLE_
      end
      object TE_DATEECHEANCE2: THLabel
        Left = 468
        Top = 38
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATEECHEANCE_
      end
      object TE_DATEECHEANCE: THLabel
        Left = 255
        Top = 38
        Width = 72
        Height = 13
        Caption = '&Ech'#233'ances du '
        FocusControl = E_DATEECHEANCE
      end
      object TE_JOURNAL: THLabel
        Left = 10
        Top = 38
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = E_JOURNAL
      end
      object E_GENERAL: THCritMaskEdit
        Left = 385
        Top = 62
        Width = 181
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
        TabOrder = 7
        TagDispatch = 0
        ElipsisButton = True
      end
      object E_AUXILIAIRE: THCritMaskEdit
        Left = 58
        Top = 62
        Width = 180
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 6
        TagDispatch = 0
        ElipsisButton = True
      end
      object E_EXERCICE: THValComboBox
        Left = 58
        Top = 6
        Width = 181
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = E_EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 385
        Top = 6
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 1
        Text = '01/01/1900'
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        DefaultDate = od1900
        ElipsisButton = True
        ControlerDate = True
      end
      object E_DATECOMPTABLE_: THCritMaskEdit
        Left = 489
        Top = 6
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 2
        Text = '31/12/2099'
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        DefaultDate = od2099
        ElipsisButton = True
        ControlerDate = True
      end
      object E_DATEECHEANCE_: THCritMaskEdit
        Left = 489
        Top = 34
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 5
        Text = '31/12/2099'
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        DefaultDate = od2099
        ElipsisButton = True
        ControlerDate = True
      end
      object E_DATEECHEANCE: THCritMaskEdit
        Left = 385
        Top = 34
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 4
        Text = '01/01/1900'
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        DefaultDate = od1900
        ElipsisButton = True
        ControlerDate = True
      end
      object E_JOURNAL: THValComboBox
        Left = 58
        Top = 34
        Width = 180
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        TagDispatch = 0
        Vide = True
        DataType = 'TTJOURNAUX'
      end
      object XX_WHEREAN: TEdit
        Left = 101
        Top = 9
        Width = 13
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 8
        Text = 'E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H"'
        Visible = False
      end
      object E_NUMECHE: THCritMaskEdit
        Left = 120
        Top = 9
        Width = 12
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 9
        Text = '1'
        Visible = False
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_QUALIFPIECE: THCritMaskEdit
        Left = 152
        Top = 9
        Width = 12
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 10
        Text = 'N'
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object E_ECHE: THCritMaskEdit
        Left = 170
        Top = 9
        Width = 12
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 11
        Text = 'X'
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object E_ETATLETTRAGE: THCritMaskEdit
        Left = 188
        Top = 9
        Width = 12
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 12
        Text = 'PL'
        Visible = False
        TagDispatch = 0
        Operateur = Inferieur
      end
      object E_TRESOLETTRE: THCritMaskEdit
        Left = 204
        Top = 9
        Width = 12
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 13
        Text = 'X'
        Visible = False
        TagDispatch = 0
        Operateur = Different
      end
      object XX_WHERENTP: TEdit
        Left = 85
        Top = 9
        Width = 13
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 14
        Text = 
          'E_NATUREPIECE="FC" OR E_NATUREPIECE="AC" OR E_NATUREPIECE="OD"  ' +
          'OR E_NATUREPIECE="OC"'
        Visible = False
      end
      object XX_WHEREMONTANT: TEdit
        Left = 69
        Top = 9
        Width = 13
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 15
        Text = 'E_DEBIT+E_CREDIT-E_COUVERTURE>0'
        Visible = False
      end
      object XX_WHEREENC: TEdit
        Left = 220
        Top = 9
        Width = 13
        Height = 19
        Color = clYellow
        Ctl3D = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 16
        Text = 'XX_WHEREENC'
        Visible = False
      end
    end
    inherited PComplement: THTabSheet
      inherited Bevel2: TBevel
        Width = 622
        Height = 89
      end
      object TE_NUMEROPIECE: THLabel
        Left = 8
        Top = 37
        Width = 42
        Height = 13
        Caption = '&Pi'#232'ce de'
        FocusControl = E_NUMEROPIECE
      end
      object HLabel2: THLabel
        Left = 170
        Top = 35
        Width = 6
        Height = 13
        Caption = #224
      end
      object TE_ETABLISSEMENT: THLabel
        Left = 8
        Top = 10
        Width = 65
        Height = 13
        Caption = '&Etablissement'
        FocusControl = E_ETABLISSEMENT
      end
      object Label14: TLabel
        Left = 296
        Top = 37
        Width = 88
        Height = 13
        Caption = '&Mode de paiement'
        FocusControl = E_MODEPAIE
      end
      object HLabel17: THLabel
        Left = 8
        Top = 66
        Width = 33
        Height = 13
        Caption = '&Devise'
        Enabled = False
        FocusControl = E_DEVISE
      end
      object TE_NATUREPIECE: THLabel
        Left = 296
        Top = 10
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = E_NATUREPIECE
      end
      object E_NUMEROPIECE: THCritMaskEdit
        Left = 82
        Top = 33
        Width = 77
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_NUMEROPIECE_: THCritMaskEdit
        Left = 184
        Top = 33
        Width = 77
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_ETABLISSEMENT: THValComboBox
        Left = 82
        Top = 6
        Width = 181
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTETABLISSEMENT'
      end
      object E_MODEPAIE: THValComboBox
        Tag = -9980
        Left = 396
        Top = 33
        Width = 171
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        Vide = True
        DataType = 'TTMODEPAIE'
      end
      object E_DEVISE: THValComboBox
        Tag = 1
        Left = 82
        Top = 62
        Width = 181
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        Enabled = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 5
        TagDispatch = 0
        DataType = 'TTDEVISE'
      end
      object E_NATUREPIECE: THValComboBox
        Left = 396
        Top = 6
        Width = 171
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 1
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATPIECEVENTE'
      end
    end
    object PzLibre: TTabSheet [2]
      Caption = 'Tables libres'
      ImageIndex = 4
      object Bevel5: TBevel
        Left = 0
        Top = 0
        Width = 622
        Height = 89
        Align = alClient
      end
      object TT_TABLE0: THLabel
        Left = 7
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE0'
      end
      object TT_TABLE5: THLabel
        Left = 7
        Top = 44
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE5'
      end
      object TT_TABLE6: THLabel
        Left = 122
        Top = 44
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE6'
      end
      object TT_TABLE1: THLabel
        Left = 122
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE1'
      end
      object TT_TABLE2: THLabel
        Left = 237
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE2'
      end
      object TT_TABLE7: THLabel
        Left = 237
        Top = 44
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE7'
      end
      object TT_TABLE8: THLabel
        Left = 351
        Top = 44
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE8'
      end
      object TT_TABLE3: THLabel
        Left = 351
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE3'
      end
      object TT_TABLE4: THLabel
        Left = 468
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE4'
      end
      object TT_TABLE9: THLabel
        Left = 468
        Top = 44
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE9'
      end
      object T_TABLE0: THCpteEdit
        Left = 7
        Top = 17
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 0
        ZoomTable = tzNatTiers0
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE5: THCpteEdit
        Left = 7
        Top = 59
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 1
        ZoomTable = tzNatTiers5
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE6: THCpteEdit
        Left = 122
        Top = 59
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 2
        ZoomTable = tzNatTiers6
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE1: THCpteEdit
        Left = 122
        Top = 17
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
      object T_TABLE2: THCpteEdit
        Left = 237
        Top = 17
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 4
        ZoomTable = tzNatTiers2
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE7: THCpteEdit
        Left = 237
        Top = 59
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 5
        ZoomTable = tzNatTiers7
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE8: THCpteEdit
        Left = 351
        Top = 59
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 6
        ZoomTable = tzNatTiers8
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE3: THCpteEdit
        Left = 351
        Top = 17
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 7
        ZoomTable = tzNatTiers3
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE4: THCpteEdit
        Left = 468
        Top = 17
        Width = 107
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 8
        ZoomTable = tzNatTiers4
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object T_TABLE9: THCpteEdit
        Left = 468
        Top = 59
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
    inherited PAvance: THTabSheet
      inherited Bevel4: TBevel
        Width = 622
        Height = 89
      end
      inherited bEffaceAvance: TToolbarButton97
        Left = 577
      end
      inherited Z_C1: THValComboBox
        Left = 6
        Width = 151
      end
      inherited Z_C2: THValComboBox
        Left = 6
        Width = 151
      end
      inherited Z_C3: THValComboBox
        Left = 6
        Width = 151
      end
      inherited ZO3: THValComboBox
        Left = 167
        Width = 151
      end
      inherited ZO2: THValComboBox
        Left = 167
        Width = 151
      end
      inherited ZO1: THValComboBox
        Left = 167
        Width = 151
      end
      inherited ZV1: THEdit
        Left = 331
      end
      inherited ZV2: THEdit
        Left = 331
      end
      inherited ZV3: THEdit
        Left = 331
      end
      inherited ZG2: THCombobox
        Left = 531
      end
      inherited ZG1: THCombobox
        Left = 531
      end
    end
    inherited PSQL: THTabSheet
      inherited Bevel3: TBevel
        Width = 622
        Height = 89
      end
      inherited Z_SQL: THSQLMemo
        Width = 622
        Height = 89
      end
    end
  end
  inherited Dock971: TDock97
    Top = 117
    Width = 630
    inherited PFiltres: TToolWindow97
      ClientWidth = 630
      ClientAreaWidth = 630
      inherited BCherche: TToolbarButton97
        Left = 418
      end
      inherited lpresentation: THLabel
        Left = 451
      end
      inherited FFiltres: THValComboBox
        Width = 343
      end
      inherited cbPresentations: THValComboBox
        Left = 522
      end
    end
  end
  inherited FListe: THDBGrid
    Top = 158
    Width = 601
    Height = 107
    OnFlipSelection = FListeFlipSelection
    MultiSelection = True
    MultiFieds = 
      'E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_NUMEROPIECE;E_QUALIFPIECE' +
      ';E_NUMLIGNE;E_NUMECHE;'
    SortEnabled = True
  end
  inherited Panel2: THPanel
    Left = 425
    inherited PListe: THPanel
      Left = 444
    end
  end
  inherited Dock: TDock97
    Top = 323
    Width = 630
    inherited PanelBouton: TToolWindow97
      ClientWidth = 630
      ClientAreaWidth = 630
      inherited Binsert: TToolbarButton97 [0]
        Left = 456
      end
      inherited BReduire: TToolbarButton97 [1]
      end
      inherited BAgrandir: TToolbarButton97 [2]
      end
      inherited bSelectAll: TToolbarButton97 [3]
        Visible = True
      end
      inherited BRechercher: TToolbarButton97 [4]
      end
      inherited BParamListe: TToolbarButton97 [5]
      end
      inherited bExport: TToolbarButton97 [6]
      end
      inherited BImprimer: TToolbarButton97 [7]
        Left = 503
      end
      inherited BOuvrir: TToolbarButton97 [8]
        Left = 535
        Hint = 'Valider la s'#233'lection'
        Caption = 'Valider'
      end
      inherited BAnnuler: TToolbarButton97 [9]
        Left = 567
      end
      inherited BAide: TToolbarButton97 [10]
        Left = 599
      end
      inherited BBlocNote: TToolbarButton97
        Left = 211
        Visible = True
      end
    end
  end
  object FPied: TPanel [7]
    Left = 0
    Top = 287
    Width = 630
    Height = 36
    Align = alBottom
    BevelInner = bvRaised
    BevelOuter = bvNone
    TabOrder = 8
    object FSession: TLabel
      Left = 3
      Top = 1
      Width = 129
      Height = 13
      Alignment = taCenter
      Caption = 'Mouvements s'#233'lectionn'#233's :'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object Label9: TLabel
      Left = 199
      Top = 1
      Width = 14
      Height = 13
      Alignment = taCenter
      Caption = 'Nb'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object Label10: TLabel
      Left = 303
      Top = 1
      Width = 25
      Height = 13
      Caption = 'D'#233'bit'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object Label11: TLabel
      Left = 418
      Top = 1
      Width = 27
      Height = 13
      Caption = 'Cr'#233'dit'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object Label12: TLabel
      Left = 556
      Top = 1
      Width = 27
      Height = 13
      Caption = 'Solde'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object Label13: TLabel
      Left = 20
      Top = 19
      Width = 125
      Height = 15
      AutoSize = False
      Caption = 'Totaux'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clGreen
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      WordWrap = True
    end
    object ISigneEuro1: TImage
      Left = 500
      Top = 3
      Width = 16
      Height = 16
      AutoSize = True
      Picture.Data = {
        07544269746D6170F6000000424DF60000000000000076000000280000001000
        0000100000000100040000000000800000000000000000000000100000001000
        0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
        C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF00FFFFFFFFFFFFFFFFFFFFFFF44444444FFFFFF44444444444FFFF4444FFFF
        FFF4FFFF444FFFFFFFFFFFFF44FFFFFFFFFFF44444444444FFFFFF4444444444
        4FFFFFF44FFFFFFFFFFFF444444444444FFFFF444444444444FFFFFF44FFFFFF
        FFF4FFFF444FFFFFFF44FFFFF444FFFFF444FFFFFF4444444444FFFFFFF44444
        4FF4}
      Stretch = True
      Transparent = True
      Visible = False
    end
    object FNbP: THNumEdit
      Tag = -1
      Left = 153
      Top = 19
      Width = 60
      Height = 15
      TabStop = False
      AutoSize = False
      BorderStyle = bsNone
      Color = clBtnFace
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clGreen
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0'
      Debit = False
      ParentCtl3D = False
      ParentFont = False
      ReadOnly = True
      TabOrder = 0
      UseRounding = True
      Validate = False
    end
    object FDebitPA: THNumEdit
      Left = 216
      Top = 19
      Width = 114
      Height = 15
      TabStop = False
      AutoSize = False
      BorderStyle = bsNone
      Color = clBtnFace
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clGreen
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      ParentCtl3D = False
      ParentFont = False
      ReadOnly = True
      TabOrder = 1
      UseRounding = True
      Validate = False
    end
    object FCreditPA: THNumEdit
      Left = 333
      Top = 19
      Width = 114
      Height = 15
      TabStop = False
      AutoSize = False
      BorderStyle = bsNone
      Color = clBtnFace
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clGreen
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      ParentCtl3D = False
      ParentFont = False
      ReadOnly = True
      TabOrder = 2
      UseRounding = True
      Validate = False
    end
    object FSoldeA: THNumEdit
      Left = 450
      Top = 19
      Width = 135
      Height = 15
      TabStop = False
      AutoSize = False
      BorderStyle = bsNone
      Color = clBtnFace
      Ctl3D = False
      Decimals = 2
      Digits = 12
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clGreen
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Masks.PositiveMask = '#,##0.00'
      Masks.ZeroMask = '0.00'
      Debit = False
      ParentCtl3D = False
      ParentFont = False
      ReadOnly = True
      TabOrder = 3
      UseRounding = True
      Validate = False
    end
  end
  inherited PCumul: THPanel
    Top = 265
    Width = 630
  end
  inherited TBlocNote: TToolWindow97
    Left = 193
    Top = 169
  end
  inherited PanVBar: THPanel
    Left = 601
    Top = 158
    Height = 107
  end
  inherited SQ: TDataSource
    Left = 104
  end
  inherited FindDialog: THFindDialog
    Left = 148
  end
  inherited POPF: THPopupMenu
    Left = 236
  end
  inherited HMTrad: THSystemMenu
    Left = 192
  end
  inherited SD: THSaveDialog
    Left = 280
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?caption?;Validation impossible : Vous n'#39'avez pas s'#233'lectionnn'#233 +
        '.les comptes de l'#39#233'criture en cours de saisie;W;O;O;O;'
      'Pr'#233'paration des encaissements : affectation des banques'
      'Pr'#233'paration des d'#233'caissements : affectation des banques')
    Left = 79
    Top = 197
  end
end

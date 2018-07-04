inherited FImputAna: TFImputAna
  Left = 294
  Top = 205
  Width = 600
  HelpContext = 999999984
  Caption = 'R'#233'-imputations analytiques'
  PixelsPerInch = 96
  TextHeight = 13
  inherited Pages: THPageControl2
    Width = 592
    ActivePage = PCritere
    inherited PCritere: THTabSheet
      object TY_AXE: TLabel
        Left = 8
        Top = 8
        Width = 18
        Height = 13
        Caption = 'A&xe'
        FocusControl = Y_AXE
      end
      object HLabel4: THLabel
        Left = 8
        Top = 34
        Width = 46
        Height = 13
        Caption = '&G'#233'n'#233'raux'
        FocusControl = Y_GENERAL_SUP
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel5: THLabel
        Left = 168
        Top = 34
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = Y_GENERAL_INF
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TE_EXERCICE: THLabel
        Left = 8
        Top = 60
        Width = 41
        Height = 13
        Caption = 'E&xercice'
        FocusControl = Y_EXERCICE
      end
      object TY_SECTION: TLabel
        Left = 316
        Top = 8
        Width = 36
        Height = 13
        Caption = '&Section'
        FocusControl = Y_SECTION
      end
      object TE_ETABLISSEMENT: THLabel
        Left = 316
        Top = 34
        Width = 65
        Height = 13
        Caption = '&Etablissement'
        FocusControl = Y_ETABLISSEMENT
      end
      object TE_DATECOMPTABLE: THLabel
        Left = 316
        Top = 60
        Width = 43
        Height = 13
        Caption = '&Dates du'
        FocusControl = Y_DATECOMPTABLE
      end
      object TE_DATECOMPTABLE2: THLabel
        Left = 476
        Top = 60
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = Y_DATECOMPTABLE_
      end
      object Y_AXE: THValComboBox
        Left = 68
        Top = 4
        Width = 203
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 0
        OnChange = Y_AXEChange
        TagDispatch = 0
        DataType = 'TTAXE'
      end
      object Y_GENERAL_SUP: THCpteEdit
        Tag = 1
        Left = 68
        Top = 30
        Width = 89
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 1
        ZoomTable = tzGventilable
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object Y_GENERAL_INF: THCpteEdit
        Tag = 1
        Left = 182
        Top = 30
        Width = 89
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        MaxLength = 17
        ParentCtl3D = False
        TabOrder = 2
        ZoomTable = tzGventilable
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object Y_EXERCICE: THValComboBox
        Left = 68
        Top = 56
        Width = 203
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 3
        OnChange = Y_EXERCICEChange
        TagDispatch = 0
        DataType = 'TTEXOSAUFPRECEDENT'
      end
      object Y_SECTION: THCpteEdit
        Left = 392
        Top = 4
        Width = 181
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 4
        OnChange = Y_SECTIONChange
        ZoomTable = tzSection
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object Y_ETABLISSEMENT: THValComboBox
        Left = 392
        Top = 30
        Width = 181
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 5
        TagDispatch = 0
        Vide = True
        DataType = 'TTETABLISSEMENT'
      end
      object Y_DATECOMPTABLE: THCritMaskEdit
        Left = 392
        Top = 56
        Width = 73
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 6
        Text = '  /  /    '
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object Y_DATECOMPTABLE_: THCritMaskEdit
        Left = 500
        Top = 56
        Width = 73
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 7
        Text = '  /  /    '
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object Y_ECRANOUVEAU: THCritMaskEdit
        Left = 276
        Top = 10
        Width = 16
        Height = 19
        TabStop = False
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
        Text = 'N'
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object Y_QUALIFPIECE: THCritMaskEdit
        Left = 291
        Top = 10
        Width = 16
        Height = 19
        TabStop = False
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
        Text = 'N'
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object XX_WHERE: TEdit
        Left = 276
        Top = 28
        Width = 16
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
        Text = 'Y_NATUREPIECE<>"ECC"'
        Visible = False
      end
      object Y_TYPEANALYTIQUE: THCritMaskEdit
        Left = 276
        Top = 46
        Width = 16
        Height = 19
        TabStop = False
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
        Text = '-'
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object Y_NUMLIGNE: THCritMaskEdit
        Left = 291
        Top = 46
        Width = 16
        Height = 19
        TabStop = False
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
        Text = '1'
        Visible = False
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
    end
    inherited PComplement: THTabSheet
      object TE_JOURNAL: THLabel
        Left = 12
        Top = 16
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = Y_JOURNAL
      end
      object TE_DEVISE: THLabel
        Left = 12
        Top = 49
        Width = 33
        Height = 13
        Caption = '&Devise'
        FocusControl = Y_DEVISE
      end
      object TE_NATUREPIECE: THLabel
        Left = 308
        Top = 16
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = Y_NATUREPIECE
      end
      object TE_NUMEROPIECE: THLabel
        Left = 308
        Top = 49
        Width = 56
        Height = 13
        Caption = 'N'#176' de &pi'#232'ce'
        FocusControl = Y_NUMEROPIECE
      end
      object HLabel1: THLabel
        Left = 468
        Top = 47
        Width = 6
        Height = 13
        Caption = #224
      end
      object Y_JOURNAL: THValComboBox
        Left = 64
        Top = 12
        Width = 203
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTJALSAISIE'
      end
      object Y_DEVISE: THValComboBox
        Left = 64
        Top = 45
        Width = 203
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
      object Y_NATUREPIECE: THValComboBox
        Left = 384
        Top = 12
        Width = 181
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATUREPIECE'
      end
      object Y_NUMEROPIECE: THCritMaskEdit
        Left = 384
        Top = 45
        Width = 73
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object Y_NUMEROPIECE_: THCritMaskEdit
        Left = 492
        Top = 45
        Width = 73
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
    end
    inherited PAvance: THTabSheet
      Caption = '???Avanc'#233's'
      inherited Bevel4: TBevel
        Width = 584
      end
    end
    inherited PSQL: THTabSheet
      Caption = '???SQL'
      inherited Bevel3: TBevel
        Width = 584
      end
      inherited Z_SQL: THSQLMemo
        Width = 584
      end
    end
    object TabSheet1: TTabSheet
      Caption = 'Param'#232'tres'
      object Bevel5: TBevel
        Left = 0
        Top = 0
        Width = 584
        Height = 83
        Align = alClient
      end
      object HLabel2: THLabel
        Left = 307
        Top = 8
        Width = 85
        Height = 13
        Caption = '&Journal analytique'
        FocusControl = JALGENERE
      end
      object HLabel3: THLabel
        Left = 139
        Top = 8
        Width = 76
        Height = 13
        Caption = '&Date g'#233'n'#233'ration'
        FocusControl = DATEGENERE
      end
      object HLabel6: THLabel
        Left = 139
        Top = 34
        Width = 66
        Height = 13
        Caption = '% '#224' r'#233'-&imputer'
        FocusControl = Ratio
      end
      object HLabel7: THLabel
        Left = 307
        Top = 60
        Width = 87
        Height = 13
        Caption = '&Grille de r'#233'partition'
        FocusControl = VENTTYPE
      end
      object HLabel8: THLabel
        Left = 307
        Top = 34
        Width = 50
        Height = 13
        Caption = '&R'#233'f'#233'rence'
        FocusControl = REFGENERE
      end
      object BZoomVentil: TToolbarButton97
        Left = 555
        Top = 55
        Width = 23
        Height = 23
        Hint = 'Voir la ventilation type'
        Layout = blGlyphRight
        ParentShowHint = False
        ShowHint = True
        OnClick = BZoomVentilClick
        GlobalIndexImage = 'Z0061_S16G1'
      end
      object RImput: TRadioGroup
        Left = 4
        Top = 3
        Width = 125
        Height = 72
        ItemIndex = 1
        Items.Strings = (
          '&Modifier l'#39'imputation'
          'G'#233'n'#233'rer une &OD')
        TabOrder = 0
        OnClick = RImputClick
      end
      object JALGENERE: THValComboBox
        Left = 400
        Top = 4
        Width = 178
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        DataType = 'TTJALANALYTIQUE'
      end
      object DATEGENERE: THCritMaskEdit
        Left = 224
        Top = 4
        Width = 65
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 1
        Text = '  /  /    '
        OnExit = DATEGENEREExit
        TagDispatch = 0
        Operateur = Egal
        OpeType = otDate
        ControlerDate = True
      end
      object ContreP: TCheckBox
        Left = 137
        Top = 58
        Width = 152
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Contrepasser en n'#233'gatif'
        Checked = True
        State = cbChecked
        TabOrder = 3
      end
      object Ratio: THNumEdit
        Left = 225
        Top = 30
        Width = 65
        Height = 21
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Max = 1.000000000000000000
        Min = 0.010000000000000000
        Debit = False
        NumericType = ntPercentage
        TabOrder = 2
        UseRounding = True
        Value = 1.000000000000000000
        Validate = False
        OnExit = RatioExit
      end
      object VENTTYPE: THValComboBox
        Left = 400
        Top = 56
        Width = 150
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 6
        TagDispatch = 0
        DataType = 'TTVENTILTYPE'
      end
      object REFGENERE: TEdit
        Left = 400
        Top = 30
        Width = 178
        Height = 21
        MaxLength = 35
        TabOrder = 5
      end
    end
  end
  inherited Dock971: TDock97
    Width = 592
    inherited PFiltres: TToolWindow97
      ClientWidth = 592
      ClientAreaWidth = 592
      inherited BCherche: TToolbarButton97
        Left = 380
      end
      inherited lpresentation: THLabel
        Left = 413
      end
      inherited FFiltres: THValComboBox
        Width = 305
      end
      inherited cbPresentations: THValComboBox
        Left = 484
      end
    end
  end
  inherited FListe: THDBGrid
    Width = 563
    ReadOnly = True
    MultiSelection = True
    MultiFieds = 
      'Y_JOURNAL;Y_DATECOMPTABLE;Y_NUMEROPIECE;Y_NUMLIGNE;Y_NUMVENTIL;Y' +
      '_GENERAL;'
  end
  inherited Panel2: THPanel
    Left = 458
    inherited PListe: THPanel
      Left = 408
    end
  end
  inherited Dock: TDock97
    Width = 592
    inherited PanelBouton: TToolWindow97
      ClientWidth = 592
      ClientAreaWidth = 592
      inherited bSelectAll: TToolbarButton97
        Visible = True
      end
      inherited BImprimer: TToolbarButton97
        Left = 464
      end
      inherited BOuvrir: TToolbarButton97
        Left = 496
        Hint = 'Lancer la r'#233'-imputation'
      end
      inherited BAnnuler: TToolbarButton97
        Left = 528
      end
      inherited BAide: TToolbarButton97
        Left = 560
      end
      inherited BBlocNote: TToolbarButton97
        Left = 432
      end
      object BListePIECES: TToolbarButton97
        Tag = 1
        Left = 164
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Ecritures de r'#233'-imputation g'#233'n'#233'r'#233'es'
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
        Visible = False
        OnClick = BListePIECESClick
        GlobalIndexImage = 'Z1826_S16G1'
        IsControl = True
      end
    end
  end
  inherited PCumul: THPanel
    Width = 592
  end
  inherited TBlocNote: TToolWindow97
    Left = 282
    Top = 114
  end
  inherited PanVBar: THPanel
    Left = 563
  end
  inherited SQ: TDataSource
    Left = 77
    Top = 248
  end
  inherited Q: THQuery
    Left = 24
    Top = 248
  end
  inherited FindDialog: THFindDialog
    Left = 182
    Top = 248
  end
  inherited POPF: THPopupMenu
    Left = 130
    Top = 248
  end
  inherited HMTrad: THSystemMenu
    Left = 235
    Top = 248
  end
  inherited SD: THSaveDialog
    Left = 288
    Top = 248
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;?caption?;Vous devez renseigner un axe analytique;W;O;O;O;'
      
        '1;?caption?;Vous devez renseigner une section analytique;W;O;O;O' +
        ';'
      
        '2;?caption?;Vous devez renseigner une section analytique valide;' +
        'W;O;O;O;'
      
        '3;?caption?;Vous devez renseigner une grille de r'#233'partition;W;O;' +
        'O;O;'
      '4;?caption?;Vous devez renseigner une date valide;W;O;O;O;'
      
        '5;?caption?;La date que vous avez renseign'#233'e est sur un exercice' +
        ' non ouvert;W;O;O;O;'
      
        '6;?caption?;La date que vous avez renseign'#233'e est sur un exercice' +
        ' non ouvert;W;O;O;O;'
      
        '7;?caption?;La date que vous avez renseign'#233'e est ant'#233'rieure '#224' la' +
        ' cl'#244'ture provisoire;W;O;O;O;'
      
        '8;?caption?;La date que vous avez renseign'#233'e est ant'#233'rieure '#224' la' +
        ' cl'#244'ture d'#233'finitive;W;O;O;O;'
      
        '9;?caption?;Vous devez renseigner un journal d'#39'OD analytique;W;O' +
        ';O;O;'
      
        '10;?caption?;La grille de r'#233'partition doit ventiler '#224' 100%;W;O;O' +
        ';O;'
      
        '11;?caption?;Il n'#39'existe aucune '#233'criture analytique qui correspo' +
        'nd aux crit'#232'res;W;O;O;O;'
      '12;?caption?;Le journal d'#39'OD analytique est ferm'#233';W;O;O;O;'
      
        '13;?caption?;Le journal d'#39'OD analytique ne poss'#232'de pas de factur' +
        'ier;W;O;O;O;'
      
        '14;?caption?;Le journal d'#39'OD analytique n'#39'est pas sur le m'#234'me ax' +
        'e;W;O;O;O;'
      'ATTENTION ! Aucune r'#233'-imputation n'#39'a '#233't'#233' g'#233'n'#233'r'#233'e.'
      
        '16;?caption?;Confirmez-vous le traitement de r'#233'-imputation analy' +
        'tique ?;Q;YNC;Y;Y;'
      '17;?caption?;Le traitement s'#39'est correctement effectu'#233';I;O;O;O;'
      
        '18;?caption?;Vous n'#39'avez s'#233'lectionn'#233' aucune '#233'criture analytique;' +
        'W;O;O;O;'
      
        '19;?caption?;Vous devez lancer une recherche au pr'#233'alable;W;O;O;' +
        'O;'
      
        '20;?caption?;Le traitement s'#39'est correctement effectu'#233'. Voulez-v' +
        'ous visualiser les '#233'critures g'#233'n'#233'r'#233'es ?;Q;YN;Y;Y;'
      
        '21;?caption?;Attention. Seules les ventilations '#224' 100% sur la se' +
        'ction '#224' r'#233'-imputer seront reventil'#233'es sur la grille ;E;O;O;O;'
      
        '22;?caption?;Il existe des incompatibilit'#233's entre le mod'#232'le de v' +
        'entilations '#224' appliquer et les mod'#232'les de restrictions analytiqu' +
        'es attach'#233's au comptes g'#233'n'#233'raux s'#233'lectionn'#233's.;W;O;O;O;')
    Left = 332
    Top = 248
  end
  object QEcrAna: TQuery
    AutoCalcFields = False
    MarshalOptions = moMarshalModifiedOnly
    EnableBCD = False
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM ANALYTIQ WHERE Y_SECTION="EA#Ewdz"')
    dataBaseName = 'SOC'
    RequestLive = True
    Left = 372
    Top = 249
  end
end

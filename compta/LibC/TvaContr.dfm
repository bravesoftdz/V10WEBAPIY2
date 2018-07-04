inherited FTvaContr: TFTvaContr
  Left = 240
  Top = 153
  Width = 613
  HelpContext = 7640000
  Caption = 'Contr'#244'le de TVA des factures'
  PixelsPerInch = 96
  TextHeight = 13
  object BParamTPF: TToolbarButton97 [0]
    Tag = 100
    Left = 220
    Top = 144
    Width = 28
    Height = 27
    Hint = 'Param'#233'trage TPF'
    Caption = 'TPF'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Margin = 2
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    Spacing = -1
    Visible = False
    OnClick = BParamTPFClick
    IsControl = True
  end
  object BParamTVA: TToolbarButton97 [1]
    Tag = 100
    Left = 188
    Top = 144
    Width = 28
    Height = 27
    Hint = 'Param'#233'trage TVA'
    Caption = 'TVA'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Margin = 2
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    Spacing = -1
    Visible = False
    OnClick = BParamTVAClick
    IsControl = True
  end
  object BZoomPiece: TToolbarButton97 [2]
    Tag = 100
    Left = 156
    Top = 144
    Width = 28
    Height = 27
    Hint = 'Voir pi'#232'ce'
    Caption = 'Piec'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Margin = 2
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    Spacing = -1
    Visible = False
    OnClick = BZoomPieceClick
    IsControl = True
  end
  inherited Pages: THPageControl2
    Width = 605
    ActivePage = PCritere
    inherited PCritere: THTabSheet
      Caption = 'Crit'#232'res'
      inherited Bevel1: TBevel
        Width = 597
      end
      object TE_NUMEROPIECE: THLabel
        Left = 296
        Top = 8
        Width = 56
        Height = 13
        Caption = '&N'#176' de pi'#232'ce'
        FocusControl = E_NUMEROPIECE
      end
      object HLabel1: THLabel
        Left = 482
        Top = 8
        Width = 6
        Height = 13
        Caption = #224
      end
      object TE_EXERCICE: THLabel
        Left = 9
        Top = 34
        Width = 41
        Height = 13
        Caption = '&Exercice'
        FocusControl = E_EXERCICE
      end
      object TE_DATECOMPTABLE: THLabel
        Left = 296
        Top = 34
        Width = 100
        Height = 13
        Caption = '&Dates comptables du'
        FocusControl = E_DATECOMPTABLE
      end
      object TE_DATECOMPTABLE2: THLabel
        Left = 482
        Top = 34
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATECOMPTABLE_
      end
      object TE_QUALIFPIECE: THLabel
        Left = 9
        Top = 8
        Width = 24
        Height = 13
        Caption = '&Type'
        FocusControl = E_QUALIFPIECE
      end
      object HLabel2: THLabel
        Left = 8
        Top = 60
        Width = 70
        Height = 13
        Caption = 'Et&ablissements'
        FocusControl = E_ETABLISSEMENT
      end
      object HLabel3: THLabel
        Left = 296
        Top = 60
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = E_JOURNAL
      end
      object E_NUMEROPIECE: THCritMaskEdit
        Left = 402
        Top = 4
        Width = 69
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_NUMEROPIECE_: THCritMaskEdit
        Left = 506
        Top = 4
        Width = 69
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_EXERCICE: THValComboBox
        Left = 82
        Top = 30
        Width = 195
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 1
        OnChange = E_EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 402
        Top = 30
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 5
        Text = '  /  /    '
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATECOMPTABLE_: THCritMaskEdit
        Left = 506
        Top = 30
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 6
        Text = '  /  /    '
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_QUALIFPIECE: THValComboBox
        Left = 82
        Top = 4
        Width = 195
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        DataType = 'TTQUALPIECE'
      end
      object E_JOURNAL: THValComboBox
        Left = 402
        Top = 57
        Width = 175
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 7
        TagDispatch = 0
        Vide = True
        DataType = 'TTJALTVA'
      end
      object E_NUMLIGNE: THCritMaskEdit
        Left = 275
        Top = 24
        Width = 16
        Height = 20
        TabStop = False
        AutoSize = False
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
        Text = '1'
        Visible = False
        TagDispatch = 0
        Operateur = Egal
        OpeType = otReel
      end
      object E_NUMECHE: THCritMaskEdit
        Left = 259
        Top = 4
        Width = 16
        Height = 20
        TabStop = False
        AutoSize = False
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
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_CREERPAR: THCritMaskEdit
        Left = 259
        Top = 24
        Width = 16
        Height = 20
        TabStop = False
        AutoSize = False
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
        Text = 'DET'
        Visible = False
        TagDispatch = 0
        Operateur = Different
      end
      object XX_WHEREMODE: TEdit
        Left = 262
        Top = 44
        Width = 16
        Height = 20
        AutoSize = False
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
        Text = 'E_MODESAISIE="-" OR E_MODESAISIE=""'
        Visible = False
      end
      object E_ETABLISSEMENT: THMultiValComboBox
        Left = 82
        Top = 57
        Width = 195
        Height = 21
        TabOrder = 2
        Style = csDialog
        Abrege = False
        DataType = 'TTETABLISSEMENT'
        Complete = False
        OuInclusif = False
      end
      object XX_WHERE: TEdit
        Left = 278
        Top = 43
        Width = 16
        Height = 20
        AutoSize = False
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
        Text = 'J_NATUREJAL="VTE" or J_NATUREJAL="ACH"'
        Visible = False
      end
    end
    inherited PComplement: THTabSheet
      inherited Bevel2: TBevel
        Width = 597
      end
      object TE_REFINTERNE: THLabel
        Left = 14
        Top = 34
        Width = 50
        Height = 13
        Caption = '&R'#233'f'#233'rence'
        FocusControl = E_REFINTERNE
      end
      object TE_NATUREPIECE: THLabel
        Left = 14
        Top = 9
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = E_NATUREPIECE
      end
      object TE_GENERAL: THLabel
        Left = 323
        Top = 9
        Width = 74
        Height = 13
        Caption = 'Compte &g'#233'n'#233'ral'
        FocusControl = E_GENERAL
      end
      object TE_AUXILIAIRE: THLabel
        Left = 323
        Top = 34
        Width = 79
        Height = 13
        Caption = 'Compte &auxiliaire'
        FocusControl = E_AUXILIAIRE
      end
      object HTOLERECALC: THLabel
        Left = 323
        Top = 60
        Width = 107
        Height = 13
        Caption = '&Tol'#233'rance sur le calcul'
        FocusControl = TOLERECALC
      end
      object E_REFINTERNE: TEdit
        Left = 77
        Top = 30
        Width = 166
        Height = 21
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 1
      end
      object E_SAISIEEURO: TCheckBox
        Left = 12
        Top = 58
        Width = 148
        Height = 17
        Alignment = taLeftJustify
        AllowGrayed = True
        Caption = '&Origine Euro'
        State = cbGrayed
        TabOrder = 2
      end
      object E_NATUREPIECE: THValComboBox
        Left = 77
        Top = 5
        Width = 166
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATUREPIECE'
      end
      object E_GENERAL: THCpteEdit
        Left = 442
        Top = 5
        Width = 137
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 3
        ZoomTable = tzGeneral
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object E_AUXILIAIRE: THCpteEdit
        Left = 442
        Top = 30
        Width = 137
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 4
        ZoomTable = tzTiers
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object TOLERECALC: THCritMaskEdit
        Left = 442
        Top = 56
        Width = 137
        Height = 21
        TabOrder = 5
        Text = '0'
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
    end
    inherited PAvance: THTabSheet
      Caption = '???Avanc'#233's'
      inherited Bevel4: TBevel
        Width = 597
      end
    end
    inherited PSQL: THTabSheet
      Caption = '???SQL'
      inherited Bevel3: TBevel
        Width = 597
      end
      inherited Z_SQL: THSQLMemo
        Width = 597
      end
    end
  end
  inherited Dock971: TDock97
    Width = 605
    inherited PFiltres: TToolWindow97
      ClientWidth = 605
      ClientAreaWidth = 605
      inherited BCherche: TToolbarButton97
        Left = 392
        Flat = True
      end
      inherited lpresentation: THLabel
        Left = 425
      end
      inherited FFiltres: THValComboBox
        Width = 313
      end
      inherited cbPresentations: THValComboBox
        Left = 496
      end
    end
  end
  inherited FListe: THDBGrid
    Width = 576
    ReadOnly = True
    MultiFieds = 'E_JOURNAL;E_DATECOMPTABLE;E_NUMEROPIECE;'
    SortEnabled = True
    SortingColumns = 
      'C0;E_JOURNAL;$E_NATUREPIECE;E_DATECOMPTABLE;E_NUMEROPIECE;E_REFI' +
      'NTERNE;'
  end
  inherited Panel2: THPanel
    Left = 457
    inherited PListe: THPanel
      Left = 421
    end
  end
  inherited Dock: TDock97
    Width = 605
    inherited PanelBouton: TToolWindow97
      ClientWidth = 605
      ClientAreaWidth = 605
      inherited bSelectAll: TToolbarButton97
        Left = 140
        Visible = True
      end
      inherited BParamListe: TToolbarButton97
        Left = 108
        Visible = False
      end
      inherited bExport: TToolbarButton97
        Left = 108
      end
      inherited BImprimer: TToolbarButton97
        Left = 478
      end
      inherited BOuvrir: TToolbarButton97
        Left = 510
        Hint = 'Lancer le contr'#244'le de TVA'
      end
      inherited BAnnuler: TToolbarButton97
        Left = 542
      end
      inherited BAide: TToolbarButton97
        Left = 574
      end
      inherited Binsert: TToolbarButton97
        Left = 172
      end
      inherited BBlocNote: TToolbarButton97
        Left = 388
      end
      object BMenuZoom: TToolbarButton97
        Tag = -100
        Left = 68
        Top = 2
        Width = 37
        Height = 27
        Hint = 'Menu zoom'
        DropdownArrow = True
        DropdownMenu = POPZ
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
    end
  end
  inherited PCumul: THPanel
    Width = 605
  end
  inherited PanVBar: THPanel
    Left = 576
  end
  inherited SQ: TDataSource
    Left = 160
    Top = 264
  end
  inherited Q: THQuery
    Left = 113
    Top = 264
  end
  inherited FindDialog: THFindDialog
    Left = 207
    Top = 264
  end
  inherited POPF: THPopupMenu
    Left = 253
    Top = 264
  end
  inherited HMTrad: THSystemMenu
    Left = 67
    Top = 265
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Contr'#244'le des factures;Confirmez-vous le traitement des pi'#232'ces ' +
        's'#233'lectionn'#233'es ?;Q;YNC;Y;Y;'
      
        '1;Contr'#244'le des factures;Vous n'#39'avez s'#233'lectionn'#233' aucune pi'#232'ce.;E;' +
        'O;O;O;'
      
        'ATTENTION! Certaines pi'#232'ces en cours de traitement ont stopp'#233' la' +
        ' validation ...')
    Left = 20
    Top = 264
  end
  object POPZ: TPopupMenu
    Left = 300
    Top = 264
  end
end

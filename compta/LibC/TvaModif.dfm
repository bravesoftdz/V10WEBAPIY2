inherited FTvaModif: TFTvaModif
  Left = 129
  Top = 133
  Width = 622
  Height = 396
  HelpContext = 7646000
  Caption = 'Modification des '#233'l'#233'ments de TVA sur encaissements'
  PixelsPerInch = 96
  TextHeight = 13
  inherited Pages: THPageControl2
    Width = 614
    ActivePage = PCritere
    inherited PCritere: THTabSheet
      inherited Bevel1: TBevel
        Width = 606
      end
      object TE_JOURNAL: THLabel
        Left = 13
        Top = 8
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = E_JOURNAL
      end
      object TE_EXERCICE: THLabel
        Left = 13
        Top = 34
        Width = 41
        Height = 13
        Caption = 'E&xercice'
        FocusControl = E_EXERCICE
      end
      object TE_NATUREPIECE: THLabel
        Left = 14
        Top = 60
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = E_NATUREPIECE
      end
      object TE_NUMEROPIECE: THLabel
        Left = 313
        Top = 8
        Width = 56
        Height = 13
        Caption = 'N'#176' de &pi'#232'ce'
        FocusControl = E_NUMEROPIECE
      end
      object HLabel1: THLabel
        Left = 487
        Top = 7
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_DATECOMPTABLE_
      end
      object TE_DATECOMPTABLE: THLabel
        Left = 313
        Top = 34
        Width = 100
        Height = 13
        Caption = '&Dates comptables du'
        FocusControl = E_DATECOMPTABLE
      end
      object TE_DATECOMPTABLE2: THLabel
        Left = 487
        Top = 32
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATECOMPTABLE_
      end
      object TE_ETABLISSEMENT: THLabel
        Left = 313
        Top = 60
        Width = 65
        Height = 13
        Caption = '&Etablissement'
        FocusControl = E_ETABLISSEMENT
      end
      object E_JOURNAL: THValComboBox
        Left = 75
        Top = 4
        Width = 191
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTJOURNAL'
      end
      object E_EXERCICE: THValComboBox
        Left = 75
        Top = 30
        Width = 191
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 1
        OnChange = E_EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXOSAUFPRECEDENT'
      end
      object E_NATUREPIECE: THValComboBox
        Left = 75
        Top = 56
        Width = 191
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATUREPIECE'
      end
      object E_NUMEROPIECE: THCritMaskEdit
        Left = 417
        Top = 4
        Width = 65
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_NUMEROPIECE_: THCritMaskEdit
        Left = 505
        Top = 4
        Width = 65
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 417
        Top = 30
        Width = 65
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
        Left = 505
        Top = 30
        Width = 65
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
      object E_ETABLISSEMENT: THValComboBox
        Left = 417
        Top = 56
        Width = 154
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 7
        TagDispatch = 0
        Vide = True
        DataType = 'TTETABLISSEMENT'
      end
      object E_NUMECHE: THCritMaskEdit
        Left = 270
        Top = 6
        Width = 17
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
        Text = '1'
        Visible = False
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_QUALIFPIECE: THCritMaskEdit
        Left = 270
        Top = 24
        Width = 17
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
      object E_TRESOLETTRE: THCritMaskEdit
        Left = 286
        Top = 6
        Width = 17
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
        TabOrder = 10
        Text = 'X'
        Visible = False
        TagDispatch = 0
        Operateur = Different
      end
      object E_ECRANOUVEAU: THCritMaskEdit
        Left = 286
        Top = 24
        Width = 17
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
        Text = 'N'
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object XX_WHERE3: TEdit
        Left = 286
        Top = 42
        Width = 17
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
        Text = 'E_NATUREPIECE<>"ECC"'
        Visible = False
      end
      object E_ECHE: THCritMaskEdit
        Left = 270
        Top = 42
        Width = 17
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
        TabOrder = 13
        Text = 'X'
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object XX_WHEREMODE: TEdit
        Left = 290
        Top = 60
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
        TabOrder = 14
        Text = 'E_MODESAISIE="-" OR E_MODESAISIE=""'
        Visible = False
      end
    end
    inherited PComplement: THTabSheet
      inherited Bevel2: TBevel
        Width = 606
      end
      object TE_GENERAL: THLabel
        Left = 4
        Top = 8
        Width = 37
        Height = 13
        Caption = '&G'#233'n'#233'ral'
        FocusControl = E_GENERAL
      end
      object TE_AUXILIAIRE: THLabel
        Left = 4
        Top = 32
        Width = 41
        Height = 13
        Caption = '&Auxiliaire'
        FocusControl = E_AUXILIAIRE
      end
      object TE_DEVISE: THLabel
        Left = 4
        Top = 60
        Width = 33
        Height = 13
        Caption = '&Devise'
        FocusControl = E_DEVISE
      end
      object TE_DATECREATION: THLabel
        Left = 360
        Top = 8
        Width = 48
        Height = 13
        Caption = '&Cr'#233#233'es de'
        FocusControl = E_DATECREATION
      end
      object TE_DATECREATION_: THLabel
        Left = 492
        Top = 8
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_DATECREATION_
      end
      object TE_DATEECHEANCE: THLabel
        Left = 360
        Top = 34
        Width = 54
        Height = 13
        Caption = '&Ech'#233'ances'
        FocusControl = E_DATEECHEANCE
      end
      object TE_DATEECHEANCE2: THLabel
        Left = 492
        Top = 34
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_DATEECHEANCE_
      end
      object TE_REFINTERNE: THLabel
        Left = 360
        Top = 60
        Width = 50
        Height = 13
        Caption = '&R'#233'f'#233'rence'
        FocusControl = E_REFINTERNE
      end
      object TE_MODEPAIE: THLabel
        Left = 188
        Top = 8
        Width = 24
        Height = 13
        Caption = '&Paie.'
        FocusControl = E_MODEPAIE
      end
      object TE_REGIMETVA: THLabel
        Left = 188
        Top = 32
        Width = 36
        Height = 13
        Caption = 'R'#233'g&ime'
        FocusControl = E_REGIMETVA
      end
      object TE_LIBELLE: THLabel
        Left = 188
        Top = 60
        Width = 30
        Height = 13
        Caption = '&Libell'#233
        FocusControl = E_LIBELLE
      end
      object E_GENERAL: THCpteEdit
        Left = 48
        Top = 4
        Width = 133
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 0
        ZoomTable = tzGLettColl
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object E_AUXILIAIRE: THCpteEdit
        Left = 48
        Top = 30
        Width = 133
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 1
        ZoomTable = tzTiers
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object E_DEVISE: THValComboBox
        Left = 48
        Top = 56
        Width = 133
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        Vide = True
        DataType = 'TTDEVISE'
      end
      object E_DATECREATION: THCritMaskEdit
        Left = 416
        Top = 4
        Width = 71
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
      object E_DATECREATION_: THCritMaskEdit
        Left = 508
        Top = 4
        Width = 69
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
      object E_DATEECHEANCE: THCritMaskEdit
        Left = 416
        Top = 30
        Width = 71
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 8
        Text = '  /  /    '
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATEECHEANCE_: THCritMaskEdit
        Left = 508
        Top = 30
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 9
        Text = '  /  /    '
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_REFINTERNE: TEdit
        Left = 416
        Top = 56
        Width = 161
        Height = 21
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 10
      end
      object E_MODEPAIE: THValComboBox
        Left = 228
        Top = 4
        Width = 125
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        Vide = True
        DataType = 'TTMODEPAIE'
      end
      object E_REGIMETVA: THValComboBox
        Left = 228
        Top = 30
        Width = 125
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        Vide = True
        DataType = 'TTREGIMETVA'
      end
      object E_LIBELLE: TEdit
        Left = 228
        Top = 56
        Width = 125
        Height = 21
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 5
      end
    end
    inherited PAvance: THTabSheet
      inherited Bevel4: TBevel
        Width = 606
      end
      inherited bEffaceAvance: TToolbarButton97
        Left = 577
      end
      inherited Z_C1: THValComboBox
        Width = 143
      end
      inherited Z_C2: THValComboBox
        Width = 143
      end
      inherited Z_C3: THValComboBox
        Width = 143
      end
      inherited ZO3: THValComboBox
        Left = 162
        Width = 147
      end
      inherited ZO2: THValComboBox
        Left = 162
        Width = 147
      end
      inherited ZO1: THValComboBox
        Left = 162
        Width = 147
      end
      inherited ZV1: THEdit
        Left = 321
        Width = 201
      end
      inherited ZV2: THEdit
        Left = 321
        Width = 201
      end
      inherited ZV3: THEdit
        Left = 321
        Width = 201
      end
      inherited ZG2: THCombobox
        Left = 533
      end
      inherited ZG1: THCombobox
        Left = 533
      end
    end
    inherited PSQL: THTabSheet
      inherited Bevel3: TBevel
        Width = 606
      end
      inherited Z_SQL: THSQLMemo
        Width = 606
      end
    end
  end
  inherited Dock971: TDock97
    Width = 614
    inherited PFiltres: TToolWindow97
      ClientWidth = 614
      ClientAreaWidth = 614
      inherited BCherche: TToolbarButton97
        Left = 400
      end
      inherited lpresentation: THLabel
        Left = 433
      end
      inherited FFiltres: THValComboBox
        Left = 72
        Width = 321
      end
      inherited cbPresentations: THValComboBox
        Left = 504
      end
    end
  end
  inherited FListe: THDBGrid
    Width = 585
    Height = 159
  end
  inherited Panel2: THPanel
    Left = 458
    inherited PListe: THPanel
      Left = 430
    end
  end
  inherited Dock: TDock97
    Top = 333
    Width = 614
    inherited PanelBouton: TToolWindow97
      ClientWidth = 614
      ClientAreaWidth = 614
      inherited bSelectAll: TToolbarButton97
        Left = 164
      end
      inherited BImprimer: TToolbarButton97
        Left = 486
      end
      inherited BOuvrir: TToolbarButton97
        Left = 518
        Hint = 'Ouvrir la modification'
        OnClick = FListeDblClick
      end
      inherited BAnnuler: TToolbarButton97
        Left = 550
      end
      inherited BAide: TToolbarButton97
        Left = 582
      end
      inherited BBlocNote: TToolbarButton97
        Left = 454
      end
      object BZoomPiece: TToolbarButton97
        Tag = 1
        Left = 132
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Zoom '#233'criture'
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
        OnClick = BZoomPieceClick
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
    end
  end
  inherited PCumul: THPanel
    Top = 311
    Width = 614
  end
  inherited PanVBar: THPanel
    Left = 585
    Height = 159
  end
  inherited SQ: TDataSource
    Left = 71
    Top = 280
  end
  inherited Q: THQuery
    Left = 12
    Top = 280
  end
  inherited FindDialog: THFindDialog
    Left = 190
    Top = 280
  end
  inherited POPF: THPopupMenu
    Left = 130
    Top = 280
  end
  inherited HMTrad: THSystemMenu
    Left = 249
    Top = 280
  end
  inherited SD: THSaveDialog
    Left = 308
    Top = 280
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'ATTENTION. Modification non enregistr'#233'e !')
    Left = 352
    Top = 280
  end
end

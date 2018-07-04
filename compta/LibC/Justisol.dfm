inherited FJustisol: TFJustisol
  Left = 259
  Top = 187
  Width = 579
  Height = 400
  HelpContext = 7509000
  Caption = 'D'#233'tail des '#233'ch'#233'ances du compte:'
  PixelsPerInch = 96
  TextHeight = 13
  inherited Pages: THPageControl2
    Width = 571
    Height = 109
    ActivePage = PCritere
    inherited PCritere: THTabSheet
      inherited Bevel1: TBevel
        Width = 563
        Height = 81
      end
      object TE_EXERCICE: THLabel
        Left = 8
        Top = 17
        Width = 41
        Height = 13
        Caption = '&Exercice'
        FocusControl = E_EXERCICE
      end
      object TE_ETATLETTRAGE: THLabel
        Left = 8
        Top = 52
        Width = 39
        Height = 13
        Caption = '&Lettrage'
        FocusControl = E_ETATLETTRAGE
      end
      object TE_DATECOMPTABLE: THLabel
        Left = 289
        Top = 17
        Width = 100
        Height = 13
        Caption = '&Dates comptables du'
        FocusControl = E_DATECOMPTABLE
      end
      object TE_DATECOMPTABLE2: THLabel
        Left = 466
        Top = 17
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATECOMPTABLE_
      end
      object TE_DATEECHEANCE_: THLabel
        Left = 466
        Top = 52
        Width = 12
        Height = 13
        Caption = 'au'
      end
      object TE_DATEECHEANCE: THLabel
        Left = 289
        Top = 52
        Width = 99
        Height = 13
        Caption = 'Dates '#233'&ch'#233'ances du'
        FocusControl = E_DATEECHEANCE
      end
      object E_EXERCICE: THValComboBox
        Left = 58
        Top = 13
        Width = 198
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 0
        OnChange = E_EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object E_ETATLETTRAGE: THValComboBox
        Left = 58
        Top = 48
        Width = 198
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 1
        TagDispatch = 0
        Vide = True
        DataType = 'TTETATLETTRE'
      end
      object E_GENERAL: THCpteEdit
        Left = 260
        Top = 47
        Width = 26
        Height = 21
        CharCase = ecUpperCase
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        Text = 'E_GENERAL'
        Visible = False
        ZoomTable = tzGCollectif
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object E_AUXILIAIRE: THCpteEdit
        Left = 260
        Top = 13
        Width = 26
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        Text = 'E_AUXILIAIRE'
        Visible = False
        ZoomTable = tzGeneral
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 394
        Top = 13
        Width = 66
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 4
        Text = '  /  /    '
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATECOMPTABLE_: THCritMaskEdit
        Left = 483
        Top = 13
        Width = 66
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 5
        Text = '  /  /    '
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATEECHEANCE_: THCritMaskEdit
        Left = 483
        Top = 48
        Width = 66
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 6
        Text = '  /  /    '
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_DATEECHEANCE: THCritMaskEdit
        Left = 394
        Top = 48
        Width = 66
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 7
        Text = '  /  /    '
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
    end
    inherited PComplement: THTabSheet
      inherited Bevel2: TBevel
        Width = 563
        Height = 81
      end
      object TE_DEVISE: THLabel
        Left = 26
        Top = 19
        Width = 33
        Height = 13
        Caption = '&Devise'
        FocusControl = E_DEVISE
      end
      object TE_ETABLISSEMENT: THLabel
        Left = 294
        Top = 19
        Width = 65
        Height = 13
        Caption = '&Etablissement'
      end
      object TE_QUALIFPIECE: THLabel
        Left = 26
        Top = 49
        Width = 24
        Height = 13
        Caption = '&Type'
        FocusControl = E_QUALIFPIECE
      end
      object TE_REFINTERNE: THLabel
        Left = 294
        Top = 49
        Width = 55
        Height = 13
        Caption = '&R'#233'f. interne'
        FocusControl = E_REFINTERNE
      end
      object E_DEVISE: THValComboBox
        Left = 93
        Top = 15
        Width = 157
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTDEVISE'
      end
      object E_ETABLISSEMENT: THValComboBox
        Left = 370
        Top = 15
        Width = 157
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 1
        TagDispatch = 0
        Vide = True
        DataType = 'TTETABLISSEMENT'
      end
      object E_QUALIFPIECE: THValComboBox
        Left = 93
        Top = 45
        Width = 157
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 2
        TagDispatch = 0
        Vide = True
        DataType = 'TTQUALPIECE'
      end
      object E_REFINTERNE: TEdit
        Left = 370
        Top = 45
        Width = 157
        Height = 21
        TabOrder = 3
      end
    end
    inherited PAvance: THTabSheet
      inherited Bevel4: TBevel
        Width = 563
        Height = 81
      end
      inherited bEffaceAvance: TToolbarButton97
        Left = 533
      end
      inherited Z_C1: THValComboBox
        Left = 25
        Top = 5
      end
      inherited Z_C2: THValComboBox
        Left = 25
        Top = 29
      end
      inherited Z_C3: THValComboBox
        Left = 25
      end
      inherited ZO3: THValComboBox
        Left = 162
      end
      inherited ZO2: THValComboBox
        Left = 162
        Top = 29
      end
      inherited ZO1: THValComboBox
        Left = 162
        Top = 5
      end
      inherited ZV1: THEdit
        Left = 299
        Top = 5
      end
      inherited ZV2: THEdit
        Left = 299
        Top = 29
      end
      inherited ZV3: THEdit
        Left = 299
      end
      inherited ZG2: THCombobox
        Left = 491
        Top = 29
      end
      inherited ZG1: THCombobox
        Left = 491
        Top = 5
      end
    end
    inherited PSQL: THTabSheet
      inherited Bevel3: TBevel
        Width = 563
        Height = 81
      end
      inherited Z_SQL: THSQLMemo
        Width = 563
        Height = 81
      end
    end
  end
  inherited Dock971: TDock97
    Top = 109
    Width = 571
    inherited PFiltres: TToolWindow97
      ClientWidth = 571
      ClientAreaWidth = 571
      inherited BCherche: TToolbarButton97
        Left = 360
      end
      inherited lpresentation: THLabel
        Left = 393
      end
      inherited FFiltres: THValComboBox
        Width = 285
      end
      inherited cbPresentations: THValComboBox
        Left = 464
      end
    end
  end
  object PPied: TPanel [4]
    Left = 0
    Top = 283
    Width = 571
    Height = 32
    Align = alBottom
    Caption = 'PPied'
    Enabled = False
    TabOrder = 8
    object TT_LIBELLE: TLabel
      Left = 8
      Top = 7
      Width = 230
      Height = 16
      AutoSize = False
      Caption = ' '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object JU_SOLDEG: THNumEdit
      Tag = 2
      Left = 248
      Top = 6
      Width = 93
      Height = 21
      Color = clBtnFace
      Decimals = 2
      Digits = 12
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Masks.PositiveMask = '#,##0.00'
      Debit = False
      NumericType = ntDC
      ParentFont = False
      TabOrder = 0
      UseRounding = True
      Validate = False
    end
    object JU_TOTALDEBIT: THNumEdit
      Tag = 1
      Left = 350
      Top = 6
      Width = 101
      Height = 21
      Color = clBtnFace
      Decimals = 2
      Digits = 12
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Masks.PositiveMask = '#,##0'
      Debit = False
      ParentFont = False
      TabOrder = 1
      UseRounding = True
      Validate = False
    end
    object JU_TOTALCREDIT: THNumEdit
      Tag = 1
      Left = 452
      Top = 6
      Width = 101
      Height = 21
      Color = clBtnFace
      Decimals = 2
      Digits = 12
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Masks.PositiveMask = '#,##0'
      Debit = False
      ParentFont = False
      TabOrder = 2
      UseRounding = True
      Validate = False
    end
  end
  inherited FListe: THDBGrid
    Top = 150
    Width = 542
    Height = 133
  end
  inherited Panel2: THPanel
    Left = 428
    Caption = ''
    inherited PListe: THPanel
      Left = 387
    end
  end
  inherited Dock: TDock97
    Top = 337
    Width = 571
    inherited PanelBouton: TToolWindow97
      ClientWidth = 571
      ClientAreaWidth = 571
      inherited BImprimer: TToolbarButton97
        Left = 473
      end
      inherited BOuvrir: TToolbarButton97
        Left = 505
        OnClick = FListeDblClick
      end
      inherited BAnnuler: TToolbarButton97
        Left = 537
      end
      inherited BAide: TToolbarButton97
        Left = 577
      end
      inherited BBlocNote: TToolbarButton97
        Left = 441
      end
    end
  end
  inherited PCumul: THPanel
    Top = 315
    Width = 571
  end
  inherited PanVBar: THPanel
    Left = 542
    Top = 150
    Height = 133
  end
  inherited SQ: TDataSource
    Left = 77
    Top = 252
  end
  inherited Q: THQuery
    Top = 252
  end
  inherited FindDialog: THFindDialog
    Left = 259
    Top = 252
  end
  inherited POPF: THPopupMenu
    Left = 198
    Top = 252
  end
  inherited HMTrad: THSystemMenu
    Left = 138
    Top = 253
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'D'#233'tail des '#233'ch'#233'ances du compte G'#233'n'#233'ral:'
      'D'#233'tail des '#233'ch'#233'ances du compte de Tiers:')
    Left = 320
    Top = 252
  end
end

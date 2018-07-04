inherited FZoomAna: TFZoomAna
  Left = 157
  Top = 138
  Width = 627
  HelpContext = 7175050
  Caption = 'D'#233'tail des mouvements de la section : '
  PixelsPerInch = 96
  TextHeight = 13
  object BGL: TToolbarButton97 [0]
    Tag = 100
    Left = 344
    Top = 152
    Width = 33
    Height = 27
    Hint = 'Grand-livre simple'
    Caption = 'GL'
    Visible = False
    OnClick = BGLClick
  end
  object BVisuP: TToolbarButton97 [1]
    Tag = 100
    Left = 382
    Top = 152
    Width = 33
    Height = 27
    Hint = 'Voir pi'#232'ce'
    Caption = 'VisuP'
    Visible = False
    OnClick = BVisuPClick
  end
  inherited Pages: THPageControl2
    Width = 619
    Height = 109
    ActivePage = PCritere
    inherited PCritere: THTabSheet
      inherited Bevel1: TBevel
        Width = 611
        Height = 81
      end
      object TY_EXERCICE: TLabel
        Left = 6
        Top = 34
        Width = 41
        Height = 13
        Caption = '&Exercice'
        FocusControl = Y_EXERCICE
      end
      object TY_JOURNAL: TLabel
        Left = 6
        Top = 10
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = Y_JOURNAL
      end
      object TY_NUMEROPIECE: THLabel
        Left = 281
        Top = 34
        Width = 71
        Height = 13
        Caption = 'N'#176' de &pi'#232'ce de'
        FocusControl = Y_NUMEROPIECE
      end
      object TY_DATECOMPTABLE: TLabel
        Left = 281
        Top = 10
        Width = 100
        Height = 13
        Caption = '&Dates comptables du'
        FocusControl = Y_DATECOMPTABLE
      end
      object TY_NUMEROPIECE_: THLabel
        Left = 476
        Top = 34
        Width = 6
        Height = 13
        Caption = #224
      end
      object TY_DATECOMPTABLE_: TLabel
        Left = 474
        Top = 10
        Width = 12
        Height = 13
        Caption = 'au'
      end
      object TY_AXE: TLabel
        Left = 281
        Top = 54
        Width = 18
        Height = 13
        Caption = '&Axe'
        FocusControl = Y_AXE
      end
      object TY_QUALIFPIECE: THLabel
        Left = 6
        Top = 58
        Width = 24
        Height = 13
        Caption = '&Type'
        FocusControl = Y_QUALIFPIECE
      end
      object Y_JOURNAL: THValComboBox
        Left = 55
        Top = 6
        Width = 208
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTJOURNAUX'
      end
      object Y_EXERCICE: THValComboBox
        Left = 55
        Top = 30
        Width = 208
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 1
        OnChange = Y_EXERCICEChange
        TagDispatch = 0
        DataType = 'TTEXERCICE'
      end
      object Y_DATECOMPTABLE: THCritMaskEdit
        Left = 386
        Top = 6
        Width = 77
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 3
        Text = '01/01/1900'
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object Y_NUMEROPIECE: THCritMaskEdit
        Left = 386
        Top = 30
        Width = 77
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 5
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object Y_NUMEROPIECE_: THCritMaskEdit
        Left = 498
        Top = 30
        Width = 77
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 6
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object Y_DATECOMPTABLE_: THCritMaskEdit
        Left = 498
        Top = 6
        Width = 77
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 4
        Text = '31/12/2099'
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object Y_AXE: THValComboBox
        Left = 386
        Top = 54
        Width = 189
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 7
        OnChange = Y_AXEChange
        TagDispatch = 0
        DataType = 'TTAXE'
      end
      object Y_QUALIFPIECE: THValComboBox
        Left = 55
        Top = 54
        Width = 208
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        Vide = True
        DataType = 'TTQUALPIECE'
      end
    end
    inherited PComplement: THTabSheet
      inherited Bevel2: TBevel
        Width = 611
        Height = 81
      end
      object TY_SECTION: TLabel
        Left = 348
        Top = 34
        Width = 36
        Height = 13
        Caption = '&Section'
        FocusControl = Y_SECTION
      end
      object TY_GENERAL: TLabel
        Left = 348
        Top = 9
        Width = 37
        Height = 13
        Caption = '&G'#233'n'#233'ral'
        FocusControl = Y_GENERAL
      end
      object TY_NUMEROLIGNE: THLabel
        Left = 8
        Top = 59
        Width = 71
        Height = 13
        Caption = 'N'#176' de L&igne de'
        FocusControl = Y_NUMLIGNE
      end
      object TY_NUMEROLIGNE_: THLabel
        Left = 192
        Top = 59
        Width = 6
        Height = 13
        Caption = #224
      end
      object TY_REFINTERNE: THLabel
        Left = 8
        Top = 34
        Width = 55
        Height = 13
        Caption = '&R'#233'f. interne'
      end
      object TY_NATUREPIECE: THLabel
        Left = 348
        Top = 59
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = Y_NATUREPIECE
      end
      object TY_ETABLISSSEMENT: THLabel
        Left = 8
        Top = 9
        Width = 65
        Height = 13
        Caption = '&Etablissement'
        FocusControl = Y_ETABLISSEMENT
      end
      object Y_GENERAL: THCpteEdit
        Left = 408
        Top = 5
        Width = 164
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 1
        ZoomTable = tzGeneral
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object Y_SECTION: THCpteEdit
        Left = 408
        Top = 30
        Width = 164
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 2
        ZoomTable = tzTiers
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object Y_NUMLIGNE: THCritMaskEdit
        Left = 84
        Top = 55
        Width = 100
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object Y_NUMLIGNE_: THCritMaskEdit
        Left = 206
        Top = 55
        Width = 100
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object Y_NATUREPIECE: THValComboBox
        Left = 408
        Top = 55
        Width = 164
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 5
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATUREPIECE'
      end
      object Y_REFINTERNE: TEdit
        Left = 84
        Top = 30
        Width = 222
        Height = 21
        MaxLength = 35
        TabOrder = 6
      end
      object Y_ETABLISSEMENT: THValComboBox
        Left = 84
        Top = 5
        Width = 222
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTETABLISSEMENT'
      end
    end
    object Pzlibre: TTabSheet [2]
      Caption = 'Zones libres'
      object Bevel5: TBevel
        Left = 0
        Top = 0
        Width = 611
        Height = 81
        Align = alClient
      end
      object TY_TABLE0: TLabel
        Left = 34
        Top = 18
        Width = 100
        Height = 13
        AutoSize = False
        Caption = '&Table libre n'#176'1'
        FocusControl = Y_TABLE0
      end
      object TY_TABLE2: TLabel
        Left = 34
        Top = 50
        Width = 100
        Height = 13
        AutoSize = False
        Caption = 'Ta&ble libre n'#176'3'
        FocusControl = Y_TABLE2
      end
      object TY_TABLE3: TLabel
        Left = 303
        Top = 50
        Width = 100
        Height = 13
        AutoSize = False
        Caption = 'Tab&le libre n'#176'4'
        FocusControl = Y_TABLE3
      end
      object TY_TABLE1: TLabel
        Left = 303
        Top = 18
        Width = 100
        Height = 13
        AutoSize = False
        Caption = 'T&able libre n'#176'2'
        FocusControl = Y_TABLE1
      end
      object Y_TABLE0: THCpteEdit
        Left = 136
        Top = 14
        Width = 142
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 0
        ZoomTable = tzNatEcrA0
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object Y_TABLE2: THCpteEdit
        Left = 136
        Top = 46
        Width = 142
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 1
        ZoomTable = tzNatEcrA2
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object Y_TABLE3: THCpteEdit
        Left = 404
        Top = 46
        Width = 142
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 2
        ZoomTable = tzNatEcrA3
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
      object Y_TABLE1: THCpteEdit
        Left = 404
        Top = 14
        Width = 142
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 3
        ZoomTable = tzNatEcrA1
        Vide = False
        Bourre = False
        okLocate = True
        SynJoker = False
      end
    end
    inherited PAvance: THTabSheet
      inherited Bevel4: TBevel
        Width = 611
        Height = 81
      end
      inherited bEffaceAvance: TToolbarButton97
        Left = 581
      end
      inherited ZV1: THEdit
        Width = 252
      end
      inherited ZV2: THEdit
        Width = 252
      end
      inherited ZV3: THEdit
        Width = 252
      end
      inherited ZG2: THCombobox
        Left = 537
      end
      inherited ZG1: THCombobox
        Left = 537
      end
    end
    inherited PSQL: THTabSheet
      inherited Bevel3: TBevel
        Width = 611
        Height = 81
      end
      inherited Z_SQL: THSQLMemo
        Width = 611
        Height = 81
      end
    end
  end
  inherited Dock971: TDock97
    Top = 109
    Width = 619
    inherited PFiltres: TToolWindow97
      ClientWidth = 619
      ClientAreaWidth = 619
      inherited BCherche: TToolbarButton97
        Left = 404
      end
      inherited lpresentation: THLabel
        Left = 437
      end
      inherited FFiltres: THValComboBox
        Width = 329
      end
      inherited cbPresentations: THValComboBox
        Left = 508
      end
    end
  end
  inherited FListe: THDBGrid
    Top = 150
    Width = 590
    Height = 168
  end
  inherited Panel2: THPanel
    Left = 455
    inherited PListe: THPanel
      Left = 435
    end
  end
  inherited Dock: TDock97
    Width = 619
    inherited PanelBouton: TToolWindow97
      ClientWidth = 619
      ClientAreaWidth = 619
      inherited bSelectAll: TToolbarButton97
        Left = 174
        Caption = 'S'#233'lection'
      end
      inherited BParamListe: TToolbarButton97
        Caption = 'Param'#232'tres'
      end
      inherited BImprimer: TToolbarButton97
        Left = 490
      end
      inherited BOuvrir: TToolbarButton97
        Left = 522
      end
      inherited BAnnuler: TToolbarButton97
        Left = 554
      end
      inherited BAide: TToolbarButton97
        Left = 586
      end
      inherited Binsert: TToolbarButton97
        Left = 206
      end
      inherited BBlocNote: TToolbarButton97
        Left = 458
        Visible = True
      end
      object BMenuZoom: TToolbarButton97
        Tag = -100
        Left = 134
        Top = 2
        Width = 37
        Height = 27
        Hint = 'Menu zoom'
        DisplayMode = dmGlyphOnly
        DropdownArrow = True
        DropdownMenu = PopZ
        Caption = 'Zooms'
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
    Width = 619
  end
  inherited PanVBar: THPanel
    Left = 590
    Top = 150
    Height = 168
  end
  inherited SQ: TDataSource
    Top = 232
  end
  inherited Q: THQuery
    Top = 232
  end
  inherited FindDialog: THFindDialog
    Left = 177
    Top = 232
  end
  inherited POPF: THPopupMenu
    Left = 161
    Top = 200
  end
  inherited HMTrad: THSystemMenu
    OnBeforeTraduc = HMTradBeforeTraduc
    Left = 97
    Top = 232
  end
  inherited SD: THSaveDialog
    Left = 218
    Top = 232
  end
  object PopZ: TPopupMenu
    Left = 258
    Top = 232
  end
end

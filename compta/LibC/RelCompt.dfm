inherited FRelCompt: TFRelCompt
  Left = 76
  Top = 176
  Width = 593
  Height = 467
  HelpContext = 7589000
  Caption = 'Relev'#233's de comptes'
  PixelsPerInch = 96
  TextHeight = 13
  object HLabel4: THLabel [2]
    Left = 10
    Top = 61
    Width = 56
    Height = 13
    Caption = '&Auxiliaire de'
  end
  object HLabel5: THLabel [3]
    Left = 178
    Top = 61
    Width = 6
    Height = 13
    Caption = #224
  end
  inherited Pages: THPageControl2
    Width = 585
    Height = 121
    ActivePage = PCritere
    TabWidth = 0
    inherited PCritere: THTabSheet
      Caption = 'Standard'
      inherited Bevel1: TBevel
        Width = 577
        Height = 93
      end
      object H_AUXILIAIRE: THLabel
        Left = 6
        Top = 12
        Width = 56
        Height = 13
        Caption = '&Auxiliaire de'
      end
      object H_AUXILIAIRE_: THLabel
        Left = 174
        Top = 12
        Width = 6
        Height = 13
        Caption = #224
      end
      object HLabel1: THLabel
        Left = 6
        Top = 67
        Width = 31
        Height = 13
        Caption = '&Intitul'#233
        FocusControl = T_LIBELLE
      end
      object HLabel3: THLabel
        Left = 324
        Top = 67
        Width = 54
        Height = 13
        Caption = '&Nature tiers'
        FocusControl = ChoixNat
      end
      object HLabel9: THLabel
        Left = 324
        Top = 40
        Width = 33
        Height = 13
        Caption = '&Devise'
        FocusControl = _DEVISE
      end
      object HLabel2: THLabel
        Left = 324
        Top = 12
        Width = 70
        Height = 13
        Caption = '&Etablissements'
        FocusControl = _ETABLISSEMENT
      end
      object H_COLLECTIF: THLabel
        Left = 6
        Top = 40
        Width = 52
        Height = 13
        Caption = '&Collectif de'
      end
      object H_COLLECTIF_: THLabel
        Left = 175
        Top = 41
        Width = 6
        Height = 13
        Caption = #224
      end
      object ChoixNat: TRadioGroup
        Left = 404
        Top = 57
        Width = 167
        Height = 32
        Columns = 2
        ItemIndex = 0
        Items.Strings = (
          '&D'#233'biteurs'
          '&Cr'#233'diteurs')
        TabOrder = 3
        OnClick = ChoixNatClick
      end
      object T_AUXILIAIRE_SUP: THCpteEdit
        Left = 66
        Top = 8
        Width = 101
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 0
        ZoomTable = tzTiers
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object T_AUXILIAIRE_INF: THCpteEdit
        Left = 190
        Top = 8
        Width = 101
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 1
        ZoomTable = tzTiers
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object T_LIBELLE: TEdit
        Left = 66
        Top = 63
        Width = 225
        Height = 21
        MaxLength = 35
        TabOrder = 2
      end
      object XX_WHERE: TEdit
        Left = 92
        Top = 6
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
        TabOrder = 6
        Text = 'T_NATUREAUXI="CLI" or T_NATUREAUXI="AUD"'
        Visible = False
      end
      object _DEVISE: THValComboBox
        Tag = 1
        Left = 404
        Top = 36
        Width = 167
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 5
        TagDispatch = 0
        DataType = 'TTDEVISE'
      end
      object _ETABLISSEMENT: THMultiValComboBox
        Left = 404
        Top = 8
        Width = 167
        Height = 21
        TabOrder = 4
        Style = csDialog
        Abrege = False
        DataType = 'TTETABLISSEMENT'
        Complete = False
        OuInclusif = False
      end
      object T_COLLECTIF_SUP: THCpteEdit
        Left = 66
        Top = 36
        Width = 101
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 7
        ZoomTable = tzGCollectif
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object T_COLLECTIF_INF: THCpteEdit
        Left = 190
        Top = 36
        Width = 101
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 8
        ZoomTable = tzGCollectif
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
    end
    inherited PComplement: THTabSheet
      Caption = 'Ecritures et '#233'dition'
      inherited Bevel2: TBevel
        Width = 577
        Height = 93
      end
      object TE_EXERCICE: THLabel
        Left = 9
        Top = 12
        Width = 41
        Height = 13
        Caption = 'E&xercice'
        FocusControl = _EXERCICE
      end
      object TE_DATECOMPTABLE: THLabel
        Left = 9
        Top = 40
        Width = 80
        Height = 13
        Caption = 'Dates &comptable'
      end
      object TE_DATECOMPTABLE2: THLabel
        Left = 180
        Top = 40
        Width = 12
        Height = 13
        Caption = 'au'
      end
      object H_OrdreEdit: THLabel
        Left = 294
        Top = 40
        Width = 60
        Height = 13
        Caption = '&Ordre '#233'dition'
        FocusControl = ORDREEDIT
      end
      object H_SECTEUR: THLabel
        Left = 294
        Top = 12
        Width = 69
        Height = 13
        Caption = '&Mod'#232'le '#233'dition'
        FocusControl = MODELE
      end
      object BParamModele: TToolbarButton97
        Left = 558
        Top = 7
        Width = 22
        Height = 23
        Hint = 'Param'#233'trer mod'#232'le de relev'#233
        Glyph.Data = {
          42010000424D4201000000000000760000002800000011000000110000000100
          040000000000CC00000000000000000000001000000010000000000000000000
          BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          77777000000070000000000000077000000070FFF0FFFFFF0F077000000070F0
          F0F0FF0F0F077000000070000000000000077000000077709999999077777000
          0000777090000907777770000000777090709007777770000000777090099700
          77777000000077709099070007777000000077709990770BB077700000007770
          9907770BB07770000000777090777770BB0770000000777007777770B0077000
          00007770777777770BB070000000777777777777000770000000777777777777
          777770000000}
        GlyphMask.Data = {00000000}
        ParentShowHint = False
        ShowHint = True
        OnClick = BParamModeleClick
      end
      object APP: TCheckBox
        Left = 292
        Top = 68
        Width = 135
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Aper'#231'u avant impression'
        Checked = True
        State = cbChecked
        TabOrder = 6
      end
      object ORDREEDIT: THValComboBox
        Left = 373
        Top = 36
        Width = 208
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 5
        TagDispatch = 0
        DataType = 'TTORDRERLC'
      end
      object MODELE: THValComboBox
        Left = 373
        Top = 8
        Width = 181
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        DataType = 'TTMODELERLC'
      end
      object _EXERCICE: THValComboBox
        Left = 96
        Top = 8
        Width = 183
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 0
        OnChange = _EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object AvecLettrage: TCheckBox
        Left = 8
        Top = 68
        Width = 160
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Avec les '#233'critures lettr'#233'es'
        Checked = True
        State = cbChecked
        TabOrder = 3
      end
      object _DATECOMPTABLE: THCritMaskEdit
        Left = 96
        Top = 36
        Width = 73
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 1
        Text = '  /  /    '
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object _DATECOMPTABLE_: THCritMaskEdit
        Left = 204
        Top = 36
        Width = 73
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 2
        Text = '  /  /    '
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
    end
    object PLibres: TTabSheet [2]
      Caption = 'Tables libres'
      object Bevel5: TBevel
        Left = 0
        Top = 0
        Width = 577
        Height = 93
        Align = alClient
      end
      object TT_TABLE0: THLabel
        Left = 11
        Top = 5
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE0'
      end
      object TT_TABLE1: THLabel
        Left = 126
        Top = 5
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE1'
      end
      object TT_TABLE2: THLabel
        Left = 241
        Top = 5
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE2'
      end
      object TT_TABLE3: THLabel
        Left = 355
        Top = 5
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE3'
      end
      object TT_TABLE4: THLabel
        Left = 467
        Top = 5
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE4'
      end
      object TT_TABLE5: THLabel
        Left = 11
        Top = 47
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE5'
      end
      object TT_TABLE6: THLabel
        Left = 126
        Top = 47
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE6'
      end
      object TT_TABLE7: THLabel
        Left = 241
        Top = 47
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE7'
      end
      object TT_TABLE8: THLabel
        Left = 355
        Top = 47
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE8'
      end
      object TT_TABLE9: THLabel
        Left = 467
        Top = 46
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE9'
      end
      object T_TABLE4: THCpteEdit
        Left = 467
        Top = 20
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
        Left = 355
        Top = 20
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
        Left = 241
        Top = 20
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
        Left = 126
        Top = 20
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
        Left = 11
        Top = 20
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
        Left = 11
        Top = 61
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
        Left = 126
        Top = 61
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
        Left = 241
        Top = 61
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
        Left = 355
        Top = 61
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
        Left = 467
        Top = 61
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
        Width = 577
        Height = 93
      end
      inherited bEffaceAvance: TToolbarButton97
        Left = 513
      end
    end
    inherited PSQL: THTabSheet
      inherited Bevel3: TBevel
        Width = 577
        Height = 93
      end
      inherited Z_SQL: THSQLMemo
        Width = 577
        Height = 93
      end
    end
  end
  inherited Dock971: TDock97
    Top = 121
    Width = 585
    inherited PFiltres: TToolWindow97
      ClientWidth = 585
      ClientAreaWidth = 585
      inherited BCherche: TToolbarButton97
        Left = 372
      end
      inherited lpresentation: THLabel
        Left = 405
      end
      inherited FFiltres: THValComboBox
        Width = 297
      end
      inherited cbPresentations: THValComboBox
        Left = 476
      end
    end
  end
  inherited FListe: THDBGrid
    Top = 162
    Width = 556
    Height = 208
    MultiSelection = True
    MultiFieds = 'T_AUXILIAIRE;'
  end
  object BZoomPiece: THBitBtn [7]
    Tag = 100
    Left = 132
    Top = 172
    Width = 28
    Height = 27
    Hint = 'Zoom sur les mouvements'
    Caption = 'Piec'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 9
    Visible = False
    OnClick = BZoomPieceClick
    Margin = 2
    NumGlyphs = 2
    Spacing = -1
    IsControl = True
  end
  object BZoom: THBitBtn [8]
    Tag = 100
    Left = 167
    Top = 171
    Width = 28
    Height = 28
    Hint = 'Zoom Tiers'
    Caption = 'Cpte'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    TabOrder = 8
    Visible = False
    OnClick = BZoomClick
    Margin = 2
    Spacing = -1
    IsControl = True
  end
  inherited Panel2: THPanel
    Left = 236
    Top = 176
    inherited PListe: THPanel
      Left = 505
    end
  end
  inherited Dock: TDock97
    Top = 392
    Width = 585
    Height = 48
    inherited PanelBouton: TToolWindow97
      ClientHeight = 44
      ClientWidth = 585
      ClientAreaHeight = 44
      ClientAreaWidth = 585
      inherited bSelectAll: TToolbarButton97
        Left = 201
      end
      inherited bExport: TToolbarButton97
        Left = 99
        Enabled = True
      end
      inherited BImprimer: TToolbarButton97
        Left = 459
      end
      inherited BOuvrir: TToolbarButton97
        Left = 492
        Hint = 'Edition des relev'#233's de compte'
      end
      inherited BAnnuler: TToolbarButton97
        Left = 523
      end
      inherited BAide: TToolbarButton97
        Left = 555
      end
      inherited Binsert: TToolbarButton97
        Left = 169
        Enabled = False
      end
      inherited BBlocNote: TToolbarButton97
        Left = 426
      end
      object BMenuZoom: TToolbarButton97
        Tag = -100
        Left = 129
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
    Top = 370
    Width = 585
  end
  inherited PanVBar: THPanel
    Left = 556
    Top = 162
    Height = 208
    inherited BLineDown: TToolbarButton97
      Top = 207
    end
    inherited BPageDown: TToolbarButton97
      Top = 164
    end
  end
  inherited SQ: TDataSource
    Left = 352
    Top = 236
  end
  inherited Q: THQuery
    Left = 8
    Top = 236
  end
  inherited FindDialog: THFindDialog
    Left = 300
    Top = 236
  end
  inherited POPF: THPopupMenu
    Left = 248
    Top = 236
    inherited BNouvRech: THMenuItem
      Left = 332
    end
  end
  inherited HMTrad: THSystemMenu
    Left = 196
    Top = 237
  end
  inherited SD: THSaveDialog
    Left = 52
    Top = 236
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Relev'#233's de compte;Vous n'#39'avez s'#233'lectionn'#233' aucun tiers.;E;O;O;O' +
        ';'
      
        '1;Relev'#233's de compte;Vous devez choisir un mod'#232'le d'#39#233'dition des r' +
        'elev'#233's;W;O;O;O;'
      
        '2;Relev'#233's de compte;Confirmez-vous l'#39#233'dition des relev'#233's pour le' +
        's tiers s'#233'lectionn'#233's ?;Q;YNC;Y;Y;'
      'EURO'
      '4;')
    Left = 92
    Top = 236
  end
  object POPZ: TPopupMenu
    Left = 144
    Top = 236
  end
end

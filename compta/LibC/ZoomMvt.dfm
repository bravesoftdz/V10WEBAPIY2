inherited FZoomMvt: TFZoomMvt
  Left = 107
  Top = 174
  Width = 601
  Height = 378
  Caption = 'Zoom sur les Mouvements'
  PixelsPerInch = 96
  TextHeight = 13
  inherited Pages: THPageControl2
    Width = 593
    ActivePage = PCritere
    inherited PCritere: THTabSheet
      inherited Bevel1: TBevel
        Width = 585
      end
      object TE_EXERCICE: TLabel
        Left = 5
        Top = 11
        Width = 41
        Height = 13
        Caption = '&Exercice'
        FocusControl = E_EXERCICE
      end
      object TE_JOURNAL: TLabel
        Left = 5
        Top = 35
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = E_JOURNAL
      end
      object TE_NATUREPIECE: THLabel
        Left = 5
        Top = 60
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = E_NATUREPIECE
      end
      object TE_NUMEROPIECE: THLabel
        Left = 278
        Top = 11
        Width = 56
        Height = 13
        Caption = 'N'#176' de &pi'#232'ce'
        FocusControl = E_NUMEROPIECE
      end
      object TE_REFINTERNE: THLabel
        Left = 278
        Top = 35
        Width = 55
        Height = 13
        Caption = '&R'#233'f. interne'
        FocusControl = E_REFINTERNE
      end
      object TE_DATECOMPTABLE: TLabel
        Left = 278
        Top = 60
        Width = 100
        Height = 13
        Caption = 'Dates &comptables du'
        FocusControl = E_DATECOMPTABLE
      end
      object HLabel2: THLabel
        Left = 459
        Top = 11
        Width = 6
        Height = 13
        Caption = #224
      end
      object TE_DATECOMPTABLE_: TLabel
        Left = 458
        Top = 60
        Width = 12
        Height = 13
        Caption = 'au'
      end
      object E_EXERCICE: THValComboBox
        Left = 49
        Top = 6
        Width = 207
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
        DataType = 'TTEXERCICE'
      end
      object E_JOURNAL: THValComboBox
        Left = 49
        Top = 31
        Width = 207
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        TagDispatch = 0
        DataType = 'TTJOURNAUX'
      end
      object E_NATUREPIECE: THValComboBox
        Left = 49
        Top = 56
        Width = 207
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 5
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATUREPIECE'
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 383
        Top = 56
        Width = 73
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 6
        Text = '01/01/1900'
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object E_REFINTERNE: THCritMaskEdit
        Left = 383
        Top = 31
        Width = 164
        Height = 21
        Ctl3D = True
        MaxLength = 35
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        OpeType = otDate
        ControlerDate = True
      end
      object E_NUMEROPIECE: THCritMaskEdit
        Left = 383
        Top = 5
        Width = 73
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 1
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_NUMEROPIECE_: THCritMaskEdit
        Left = 475
        Top = 5
        Width = 73
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_DATECOMPTABLE_: THCritMaskEdit
        Left = 475
        Top = 56
        Width = 73
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 7
        Text = '31/12/2099'
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
    end
    inherited PComplement: THTabSheet
      inherited Bevel2: TBevel
        Width = 585
      end
      object TE_DEBIT: TLabel
        Left = 4
        Top = 35
        Width = 80
        Height = 13
        Caption = 'Montant &d'#233'bit de'
        FocusControl = E_DEBIT
      end
      object TE_DEBIT_: TLabel
        Left = 211
        Top = 35
        Width = 6
        Height = 13
        Caption = #224
      end
      object TE_CREDIT_: TLabel
        Left = 211
        Top = 60
        Width = 6
        Height = 13
        Caption = #224
      end
      object TE_CREDIT: TLabel
        Left = 4
        Top = 60
        Width = 83
        Height = 13
        Caption = 'Montant &cr'#233'dit de'
        FocusControl = E_CREDIT
      end
      object TE_GENERAL: TLabel
        Left = 367
        Top = 35
        Width = 37
        Height = 13
        Caption = '&G'#233'n'#233'ral'
        FocusControl = E_GENERAL
      end
      object TE_AUXILIAIRE: TLabel
        Left = 367
        Top = 60
        Width = 41
        Height = 13
        Caption = '&Auxiliaire'
        FocusControl = E_AUXILIAIRE
      end
      object TE_ETABLISSEMENT: THLabel
        Left = 4
        Top = 8
        Width = 65
        Height = 13
        Caption = '&Etablissement'
        FocusControl = E_ETABLISSEMENT
      end
      object E_DEBIT: THCritMaskEdit
        Left = 91
        Top = 31
        Width = 114
        Height = 21
        TabOrder = 1
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_DEBIT_: THCritMaskEdit
        Left = 225
        Top = 31
        Width = 114
        Height = 21
        TabOrder = 2
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_CREDIT_: THCritMaskEdit
        Left = 225
        Top = 56
        Width = 114
        Height = 21
        TabOrder = 3
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_CREDIT: THCritMaskEdit
        Left = 91
        Top = 56
        Width = 114
        Height = 21
        TabOrder = 4
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_GENERAL: THCpteEdit
        Left = 412
        Top = 31
        Width = 121
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 5
        ZoomTable = tzGeneral
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object E_AUXILIAIRE: THCpteEdit
        Left = 412
        Top = 56
        Width = 121
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 6
        ZoomTable = tzTiers
        Vide = False
        Bourre = False
        okLocate = False
        SynJoker = False
      end
      object E_ETABLISSEMENT: THValComboBox
        Left = 91
        Top = 4
        Width = 249
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
    end
    inherited PAvance: THTabSheet
      inherited Bevel4: TBevel
        Width = 585
      end
      inherited bEffaceAvance: TToolbarButton97
        Left = 557
      end
      inherited ZV1: THEdit
        Width = 228
      end
      inherited ZV2: THEdit
        Width = 228
      end
      inherited ZV3: THEdit
        Width = 228
      end
      inherited ZG2: THCombobox
        Left = 515
      end
      inherited ZG1: THCombobox
        Left = 515
      end
    end
    inherited PSQL: THTabSheet
      inherited Bevel3: TBevel
        Width = 585
      end
      inherited Z_SQL: THSQLMemo
        Width = 585
      end
    end
  end
  inherited Dock971: TDock97
    Width = 593
    inherited PFiltres: TToolWindow97
      ClientWidth = 593
      ClientAreaWidth = 593
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
    Width = 564
    Height = 141
  end
  inherited Panel2: THPanel
    Left = 432
    inherited PListe: THPanel
      Left = 409
    end
  end
  inherited Dock: TDock97
    Top = 315
    Width = 593
    inherited PanelBouton: TToolWindow97
      ClientWidth = 593
      ClientAreaWidth = 593
      inherited BImprimer: TToolbarButton97
        Left = 463
      end
      inherited BOuvrir: TToolbarButton97
        Left = 495
      end
      inherited BAnnuler: TToolbarButton97
        Left = 527
      end
      inherited BAide: TToolbarButton97
        Left = 559
      end
      inherited BBlocNote: TToolbarButton97
        Left = 431
        Visible = True
      end
    end
  end
  inherited PCumul: THPanel
    Top = 293
    Width = 593
  end
  inherited PanVBar: THPanel
    Left = 564
    Height = 141
  end
  inherited SQ: TDataSource
    Left = 88
    Top = 236
  end
  inherited FindDialog: THFindDialog
    Top = 236
  end
  inherited POPF: THPopupMenu
    Left = 232
    Top = 236
  end
  inherited HMTrad: THSystemMenu
    Top = 236
  end
end

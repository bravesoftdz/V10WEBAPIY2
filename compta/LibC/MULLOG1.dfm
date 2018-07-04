inherited FMulLog: TFMulLog
  Left = 322
  Top = 132
  Width = 577
  Height = 400
  HelpContext = 3185000
  Caption = 'Suivi d'#39'activit'#233' des utilisateurs'
  PixelsPerInch = 96
  TextHeight = 13
  inherited Pages: THPageControl2
    Width = 569
    Height = 109
    ActivePage = PCritere
    inherited PCritere: THTabSheet
      inherited Bevel1: TBevel
        Width = 561
        Height = 81
      end
      object TRB_FAMILLERUB: THLabel
        Left = 14
        Top = 18
        Width = 46
        Height = 13
        Caption = '&Utilisateur'
        FocusControl = LG_UTILISATEUR
      end
      object TE_DATECREATION: THLabel
        Left = 14
        Top = 49
        Width = 38
        Height = 13
        Caption = '&Date de'
        FocusControl = LG_DATE
      end
      object TE_DATECREATION_: THLabel
        Left = 160
        Top = 49
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = LG_DATE_
      end
      object HLabel1: THLabel
        Left = 296
        Top = 49
        Width = 44
        Height = 13
        Caption = 'D&ur'#233'e de'
        FocusControl = LG_TEMPS
      end
      object HLabel2: THLabel
        Left = 427
        Top = 49
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = LG_TEMPS
      end
      object LG_UTILISATEUR: THValComboBox
        Left = 66
        Top = 14
        Width = 193
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'TTUTILISATEUR'
      end
      object LG_DATE: THCritMaskEdit
        Left = 67
        Top = 46
        Width = 71
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 1
        Text = '01/01/1900'
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        ControlerDate = True
      end
      object LG_DATE_: THCritMaskEdit
        Left = 190
        Top = 46
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 2
        Text = '31/12/2099'
        OnKeyPress = ParamLaDate
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        ControlerDate = True
      end
      object LG_TEMPS: THCritMaskEdit
        Left = 348
        Top = 46
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!00:00:00;1;_'
        MaxLength = 8
        ParentCtl3D = False
        TabOrder = 3
        Text = '00:00:00'
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otHeure
      end
      object LG_TEMPS_: THCritMaskEdit
        Left = 446
        Top = 46
        Width = 69
        Height = 21
        Ctl3D = True
        EditMask = '!00:00:00;1;_'
        MaxLength = 8
        ParentCtl3D = False
        TabOrder = 4
        Text = '99:99:99'
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otHeure
      end
    end
    inherited PComplement: THTabSheet
      Caption = '??? Compl'#233'ments'
      inherited Bevel2: TBevel
        Width = 561
        Height = 81
      end
    end
    inherited PAvance: THTabSheet
      inherited Bevel4: TBevel
        Width = 561
        Height = 81
      end
      inherited bEffaceAvance: TToolbarButton97
        Left = 533
      end
      inherited Z_C1: THValComboBox
        Top = 6
      end
      inherited Z_C2: THValComboBox
        Top = 30
      end
      inherited Z_C3: THValComboBox
        Top = 54
      end
      inherited ZO3: THValComboBox
        Top = 54
      end
      inherited ZO2: THValComboBox
        Top = 30
      end
      inherited ZO1: THValComboBox
        Top = 6
      end
      inherited ZV1: THEdit
        Top = 6
        Width = 201
      end
      inherited ZV2: THEdit
        Top = 30
        Width = 201
      end
      inherited ZV3: THEdit
        Top = 54
        Width = 201
      end
      inherited ZG2: THCombobox
        Left = 489
        Top = 30
      end
      inherited ZG1: THCombobox
        Left = 489
        Top = 6
      end
    end
    inherited PSQL: THTabSheet
      inherited Bevel3: TBevel
        Width = 561
        Height = 81
      end
      inherited Z_SQL: THSQLMemo
        Width = 561
        Height = 81
      end
    end
  end
  inherited Dock971: TDock97
    Top = 109
    Width = 569
    inherited PFiltres: TToolWindow97
      ClientWidth = 569
      ClientAreaWidth = 569
      inherited BCherche: TToolbarButton97
        Left = 354
      end
      inherited lpresentation: THLabel
        Left = 389
      end
      inherited FFiltres: THValComboBox
        Width = 279
      end
      inherited cbPresentations: THValComboBox
        Left = 460
      end
    end
  end
  inherited FListe: THDBGrid
    Top = 150
    Width = 540
    Height = 165
  end
  object CapPurge: TEdit [5]
    Left = 296
    Top = 216
    Width = 121
    Height = 21
    Color = clYellow
    TabOrder = 8
    Text = 'Purger le fichier "Suivi des utilisateurs"'
    Visible = False
  end
  inherited Panel2: THPanel
    inherited PListe: THPanel
      Left = 385
    end
  end
  inherited Dock: TDock97
    Top = 337
    Width = 569
    inherited PanelBouton: TToolWindow97
      ClientWidth = 569
      ClientAreaWidth = 569
      inherited BParamListe: TToolbarButton97
        Left = 100
        Enabled = False
        Visible = False
      end
      inherited BImprimer: TToolbarButton97
        Left = 442
      end
      inherited BOuvrir: TToolbarButton97
        Left = 474
      end
      inherited BAnnuler: TToolbarButton97
        Left = 506
      end
      inherited BAide: TToolbarButton97
        Left = 538
      end
      inherited BBlocNote: TToolbarButton97
        Left = 410
      end
      object BPurge: TToolbarButton97
        Left = 68
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Purger'
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
        OnClick = BPurgeClick
        GlobalIndexImage = 'Z0204_S16G1'
      end
    end
  end
  inherited PCumul: THPanel
    Top = 315
    Width = 569
  end
  inherited PanVBar: THPanel
    Left = 540
    Top = 150
    Height = 165
  end
  inherited SQ: TDataSource
    Left = 62
    Top = 272
  end
  inherited Q: THQuery
    AfterOpen = QAfterOpen
    Left = 12
    Top = 272
  end
  inherited FindDialog: THFindDialog
    Left = 162
    Top = 272
  end
  inherited POPF: THPopupMenu
    Left = 212
    Top = 272
  end
  inherited HMTrad: THSystemMenu
    Left = 112
    Top = 272
  end
end

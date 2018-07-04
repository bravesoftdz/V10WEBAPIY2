inherited FControlFicU: TFControlFicU
  Left = 328
  Top = 153
  Width = 696
  Height = 581
  Caption = 'Controle de Fichier TRA '
  PixelsPerInch = 96
  TextHeight = 13
  inherited Pages: THPageControl2
    Width = 688
    inherited PCritere: THTabSheet
      inherited Bevel1: TBevel
        Width = 680
      end
      inherited iCritGlyphModified: THImage
        Left = 576
        Top = 35
      end
      inherited iCritGlyph: THImage
        Left = 576
        Top = 51
      end
      object LabelTypeAff: THLabel
        Left = 14
        Top = 32
        Width = 71
        Height = 13
        Caption = 'Type de Ligne:'
      end
      object SelectAffLabel: THLabel
        Left = 16
        Top = 8
        Width = 75
        Height = 13
        Caption = 'Type Affichage:'
        Visible = False
      end
      object LabelNomFichier: THLabel
        Left = 584
        Top = 64
        Width = 79
        Height = 13
        Caption = 'LabelNomFichier'
        Visible = False
      end
      object LabelLigne: THLabel
        Left = 512
        Top = 56
        Width = 3
        Height = 13
      end
      object CmbTypeLigne: THValComboBox
        Left = 104
        Top = 32
        Width = 145
        Height = 21
        ItemHeight = 13
        TabOrder = 0
        OnChange = CmbTypeLigneChange
        TagDispatch = 0
      end
      object RGrpTypeAffich: TRadioGroup
        Left = 296
        Top = 0
        Width = 153
        Height = 73
        Caption = 'Affichage'
        ItemIndex = 0
        Items.Strings = (
          'Toutes les lignes.'
          'Lignes avec erreurs.'
          'Lignes sans erreurs.')
        TabOrder = 1
        OnClick = RGrpTypeAffichClick
      end
      object ComboTypeAffichage: THValComboBox
        Left = 104
        Top = 8
        Width = 145
        Height = 21
        ItemHeight = 13
        TabOrder = 2
        Visible = False
        OnChange = ComboTypeAffichageChange
        Items.Strings = (
          'Par type de lignes'
          'Par type de lettrage'
          'Par num'#233'ro de compte')
        TagDispatch = 0
      end
      object ComboChoixLettrage: THValComboBox
        Left = 104
        Top = 28
        Width = 145
        Height = 21
        ItemHeight = 13
        TabOrder = 3
        Visible = False
        OnChange = ComboChoixLettrageChange
        Items.Strings = (
          'Totalement lettr'#233's'
          'Partielement lettr'#233's'
          'Non Lettrable'
          'A lettrer')
        TagDispatch = 0
      end
      object ComboNumCompt: THValComboBox
        Left = 104
        Top = 28
        Width = 145
        Height = 21
        ItemHeight = 13
        TabOrder = 4
        Visible = False
        OnChange = ComboNumComptChange
        TagDispatch = 0
      end
      object ComboBloc: THValComboBox
        Left = 464
        Top = 8
        Width = 145
        Height = 21
        ItemHeight = 13
        TabOrder = 5
        Visible = False
        OnChange = ComboBlocChange
        TagDispatch = 0
      end
      object ProgressTraitement: TEnhancedGauge
        Left = 0
        Top = 64
        Width = 289
        Height = 17
        Caption = 'Donn'#233'es trait'#233's'
        TabOrder = 6
        ForeColor = clNavy
        BackColor = clBtnFace
        Progress = 0
      end
    end
    inherited PComplement: THTabSheet
      inherited Bevel2: TBevel
        Width = 680
      end
      object HLabel2: THLabel
        Left = 8
        Top = 32
        Width = 102
        Height = 13
        Caption = 'Mode de d'#233'coupage:'
      end
      object HLabel3: THLabel
        Left = 8
        Top = 8
        Width = 91
        Height = 13
        Caption = 'Fichier s'#233'lectionn'#233':'
      end
      object BValidDecoup: TToolbarButton97
        Left = 278
        Top = 50
        Width = 28
        Height = 27
        Enabled = False
        Flat = False
        Glyph.Data = {
          36030000424D3603000000000000360000002800000010000000100000000100
          1800000000000003000000000000000000000000000000000000008284008284
          0082840082840082840082840082840082840082840082840082840082840082
          8400828400828400828400828400828400828400828484000084000000828400
          8284008284008284008284008284008284008284008284008284008284008284
          0082848400000082000082008400000082840082840082840082840082840082
          8400828400828400828400828400828484000000820000820000820000820084
          0000008284008284008284008284008284008284008284008284008284840000
          0082000082000082000082000082000082008400000082840082840082840082
          8400828400828400828484000000820000820000820000FF0000820000820000
          8200008200840000008284008284008284008284008284008284008200008200
          00820000FF0000828400FF000082000082000082008400000082840082840082
          8400828400828400828400FF0000820000FF0000828400828400828400FF0000
          820000820000820084000000828400828400828400828400828400828400FF00
          00828400828400828400828400828400FF000082000082000082008400000082
          8400828400828400828400828400828400828400828400828400828400828400
          828400FF00008200008200008200840000008284008284008284008284008284
          00828400828400828400828400828400828400828400FF000082000082000082
          0084000000828400828400828400828400828400828400828400828400828400
          828400828400828400FF00008200008200008200840000008284008284008284
          00828400828400828400828400828400828400828400828400828400FF000082
          0000820000820084000000828400828400828400828400828400828400828400
          828400828400828400828400828400FF00008200008200840000008284008284
          0082840082840082840082840082840082840082840082840082840082840082
          8400FF0000820000820000828400828400828400828400828400828400828400
          828400828400828400828400828400828400828400FF00008284}
        GlyphMask.Data = {00000000}
        OnClick = BValidDecoupClick
      end
      object LabelPeriode: THLabel
        Left = 8
        Top = 56
        Width = 39
        Height = 13
        Caption = 'Periode:'
        Visible = False
      end
      object ComboModDecoup: THValComboBox
        Left = 120
        Top = 32
        Width = 145
        Height = 21
        ItemHeight = 13
        TabOrder = 0
        OnChange = ComboModDecoupChange
        Items.Strings = (
          'Par journal.'
          'Par p'#232'riode.')
        TagDispatch = 0
      end
      object Edit_NomFichier: TEdit97
        Left = 120
        Top = 8
        Width = 145
        Height = 19
        Enabled = False
        TabOrder = 1
      end
      object ComboPeriode: THValComboBox
        Left = 120
        Top = 56
        Width = 145
        Height = 21
        ItemHeight = 0
        TabOrder = 2
        Visible = False
        OnChange = ComboPeriodeChange
        TagDispatch = 0
      end
    end
    inherited PAvance: THTabSheet
      inherited Bevel4: TBevel
        Width = 680
      end
    end
    inherited PSQL: THTabSheet
      inherited Bevel3: TBevel
        Width = 680
      end
      inherited FSQL: THMemo
        Width = 680
      end
    end
  end
  inherited Dock971: TDock97
    Width = 688
    inherited PFiltres: TToolWindow97
      ClientWidth = 688
      ClientAreaWidth = 688
      inherited BCherche: TToolbarButton97
        Left = 661
      end
      inherited FFiltres: THValComboBox
        Left = 60
        Width = 587
      end
    end
  end
  inherited Dock: TDock97
    Top = 507
    Width = 688
    inherited PanelBouton: TToolWindow97
      ClientWidth = 688
      ClientAreaWidth = 688
      inherited BReduire: TToolbarButton97
        Left = 474
      end
      inherited BAgrandir: TToolbarButton97
        Left = 474
      end
      inherited BRechercher: TToolbarButton97
        Left = 503
      end
      inherited BParamListe: TToolbarButton97
        Left = 533
      end
      inherited bExport: TToolbarButton97
        Left = 564
      end
      inherited BImprimer: TToolbarButton97
        Left = 594
      end
      inherited BAnnuler: TToolbarButton97
        Left = 625
      end
      inherited BAide: TToolbarButton97
        Left = 656
      end
      inherited BPresentation: TToolbarButton97
        DropdownMenu = POPControl
      end
      object BOuvrirTRA: TToolbarButton97 [9]
        Left = 442
        Top = 2
        Width = 28
        Height = 27
        Anchors = [akTop, akRight]
        Flat = False
        Glyph.Data = {
          4E010000424D4E01000000000000760000002800000012000000120000000100
          040000000000D800000000000000000000001000000000000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777
          7777770000007777777777777777770000007777777777777777770000007700
          000000000007770000007008B8B8B8B8B8B07700000070F08B8B8B8B8B807700
          000070B0B8B8B8B8B8B80700000070FB0B8B8B8B8B8B0700000070BF08B8B8B8
          B8B8B000000070FBF000008B8B8B8000000070BFBFBFBF0000000700000070FB
          FBFBFBFBFB077700000070BFBFBFBFBFBF077700000070FBFBF0000000777700
          0000770000077777777777000000777777777777777777000000777777777777
          777777000000777777777777777777000000}
        GlyphMask.Data = {00000000}
        OnClick = BOuvrirTRAClick
      end
      inherited BStop: TToolbarButton97
        Left = 380
      end
    end
  end
  inherited TV: TTobViewer
    Width = 688
    Height = 359
  end
  inherited SortGroups: TToolWindow97
    Left = 326
    Top = 175
  end
  inherited POPD: THPopupMenu
    Left = 612
    Top = 24
  end
  inherited POPF: THPopupMenu
    Left = 640
    Top = 28
  end
  inherited POPC: THPopupMenu
    Left = 644
    Top = 0
  end
  object OpenDialog1: TOpenDialog
    Left = 372
    Top = 416
  end
  object POPControl: TPopupMenu
    AutoPopup = False
    OwnerDraw = True
    Left = 300
    Top = 416
    object TotalDbitCrditparJournal1: TMenuItem
      Caption = 'Total D'#233'bit/Cr'#233'dit par Journal'
    end
    object Equilibrerladate1: TMenuItem
      Caption = 'Calcul du solde progressif'
    end
    object Equilibrerladate2: TMenuItem
      Caption = 'Equilibrer '#224' la date'
    end
  end
end

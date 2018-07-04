inherited FGRJ_H: TFGRJ_H
  Left = 131
  Top = 125
  Width = 640
  Height = 480
  HelpContext = 301100300
  Caption = 'Palmarès Meilleure journée / Meilleure Heure '
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
    Top = 419
    Width = 632
    inherited HPB: TToolWindow97
      ClientWidth = 628
      ClientAreaWidth = 628
      inherited BValider: TToolbarButton97
        Left = 318
      end
      inherited BFerme: TToolbarButton97
        Left = 621
      end
      inherited BAide: TToolbarButton97
        Left = 650
      end
      inherited BDelete: TToolbarButton97
        Left = 593
      end
      inherited BSauve: TToolbarButton97
        Left = 564
      end
      inherited BUndo: TToolbarButton97
        Left = 535
      end
      inherited BGraph: TToolbarButton97
        Left = 352
      end
      inherited BImprimer: TToolbarButton97
        Left = 505
      end
      inherited bAffGraph: TToolbarButton97
        Left = 400
        Down = True
      end
      inherited bListe: TToolbarButton97
        Left = 428
        Down = False
      end
      inherited bExport: TToolbarButton97
        Left = 475
      end
      inherited FFiltres: TComboBox
        Width = 245
      end
    end
  end
  inherited P: TPanel
    Top = 97
    Width = 632
    Height = 322
    inherited Splitter1: TSplitter
      Left = 499
      Height = 223
      Visible = False
    end
    inherited FListe: THGrid
      Width = 630
      ColCount = 7
      FixedCols = 0
      OnDblClick = FListeDblClick
      SortedCol = 2
      Titres.Strings = (
        'Concept'
        'N.Pièces'
        'C.A.'
        'Marge'
        'Quantité')
      ColWidths = (
        9
        94
        82
        68
        67
        68
        71)
    end
    inherited FChart1: TChart
      Width = 498
      Height = 223
      OnClickSeries = FChart1ClickSeries
    end
    inherited FChart2: TChart
      Left = 502
      Width = 129
      Height = 223
      Visible = False
    end
  end
  inherited Pages: TPageControl
    Width = 632
    Height = 97
    TabHeight = 1
    TabWidth = 1
    inherited TabSheet1: TTabSheet
      TabVisible = True
      object TGP_DATEPIECE_: THLabel
        Left = 193
        Top = 12
        Width = 12
        Height = 13
        Caption = '&au'
        FocusControl = GP_DATEPIECE_
      end
      object TGP_DATEPIECE: THLabel
        Left = 2
        Top = 12
        Width = 38
        Height = 13
        Caption = '&Date du'
        FocusControl = GP_DATEPIECE
      end
      object TGP_CAISSE: THLabel
        Left = 193
        Top = 36
        Width = 31
        Height = 13
        Caption = '&Caisse'
        FocusControl = GP_CAISSE
      end
      object TGP_REPRESENTANT: THLabel
        Left = 2
        Top = 36
        Width = 40
        Height = 13
        Caption = '&Vendeur'
        FocusControl = GP_REPRESENTANT
      end
      object TFTypePalmares: THLabel
        Left = 2
        Top = 60
        Width = 69
        Height = 13
        Caption = '&Type palmarès'
        FocusControl = FTypePalmares
      end
      object GP_DATEPIECE_: THCritMaskEdit
        Left = 228
        Top = 8
        Width = 110
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 1
        Text = '  /  /    '
        OnChange = GP_DATEPIECEChange
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        DefaultDate = odFinMois
        ElipsisButton = True
        ControlerDate = True
      end
      object GP_DATEPIECE: THCritMaskEdit
        Left = 74
        Top = 8
        Width = 110
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 0
        Text = '  /  /    '
        OnChange = GP_DATEPIECEChange
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        DefaultDate = odDebMois
        ElipsisButton = True
        ControlerDate = True
      end
      object GRJOUR: TRadioGroup
        Left = 341
        Top = 1
        Width = 120
        Height = 79
        Caption = 'Journée / Heure'
        ItemIndex = 0
        Items.Strings = (
          'Meilleure &Journée'
          'Meilleure &Heure')
        TabOrder = 5
        OnClick = GP_DATEPIECEChange
      end
      object XX_WHERE: TEdit
        Left = 292
        Top = 56
        Width = 46
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 7
        Visible = False
      end
      object GP_REPRESENTANT: THValComboBox
        Left = 74
        Top = 32
        Width = 110
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
        OnChange = GP_DATEPIECEChange
        TagDispatch = 0
        Plus = 'GCL_TYPECOMMERCIAL="VEN"'
        Vide = True
        VideString = '<<Tous>>'
        DataType = 'GCCOMMERCIAL'
      end
      object GP_CAISSE: THValComboBox
        Left = 228
        Top = 32
        Width = 110
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        OnChange = GP_DATEPIECEChange
        TagDispatch = 0
        Vide = True
        VideString = '<<Tous>>'
        DataType = 'GCPCAISSE'
      end
      object FTypePalmares: THValComboBox
        Left = 74
        Top = 56
        Width = 110
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 4
        TagDispatch = 0
        Plus = 'J'
        DataType = 'GCSYNTHESEFO'
      end
      object GRCHOIXJOUR: TGroupBox
        Left = 469
        Top = 1
        Width = 152
        Height = 79
        Caption = 'Choix du jour'
        TabOrder = 6
        Visible = False
        object CBLUNDI: TCheckBox
          Left = 4
          Top = 15
          Width = 65
          Height = 17
          Caption = 'Lundi'
          Checked = True
          State = cbChecked
          TabOrder = 0
          OnClick = CHOIXJOURClick
        end
        object CBMARDI: TCheckBox
          Left = 4
          Top = 30
          Width = 65
          Height = 17
          Caption = 'Mardi'
          Checked = True
          State = cbChecked
          TabOrder = 1
          OnClick = CHOIXJOURClick
        end
        object CBMERCREDI: TCheckBox
          Left = 4
          Top = 45
          Width = 65
          Height = 17
          Caption = 'Mercredi'
          Checked = True
          State = cbChecked
          TabOrder = 2
          OnClick = CHOIXJOURClick
        end
        object CBVENDREDI: TCheckBox
          Left = 80
          Top = 15
          Width = 65
          Height = 21
          Caption = 'Vendredi'
          Checked = True
          State = cbChecked
          TabOrder = 4
          OnClick = CHOIXJOURClick
        end
        object CBDIMANCHE: TCheckBox
          Left = 80
          Top = 45
          Width = 69
          Height = 17
          Caption = 'Dimanche'
          Checked = True
          State = cbChecked
          TabOrder = 6
          OnClick = CHOIXJOURClick
        end
        object CBJEUDI: TCheckBox
          Left = 4
          Top = 60
          Width = 65
          Height = 17
          Caption = 'Jeudi'
          Checked = True
          State = cbChecked
          TabOrder = 3
          OnClick = CHOIXJOURClick
        end
        object CBSAMEDI: TCheckBox
          Left = 80
          Top = 30
          Width = 65
          Height = 17
          Caption = 'Samedi'
          Checked = True
          State = cbChecked
          TabOrder = 5
          OnClick = CHOIXJOURClick
        end
      end
      object GP_ETABLISSEMENT: THValComboBox
        Left = 228
        Top = 56
        Width = 49
        Height = 21
        Style = csDropDownList
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentFont = False
        TabOrder = 8
        Visible = False
        OnChange = GP_DATEPIECEChange
        TagDispatch = 0
        DataType = 'TTETABLISSEMENT'
      end
    end
  end
  inherited POPF: TPopupMenu
    Left = 280
    Top = 170
  end
  inherited Msg: THMsgBox
    Left = 334
    Top = 170
  end
  inherited HMTrad: THSystemMenu
    ResizeDBGrid = True
    Left = 396
    Top = 170
  end
  inherited FindDialog: TFindDialog
    Left = 447
    Top = 170
  end
  inherited PopAff: TPopupMenu
    Left = 222
    Top = 170
  end
  inherited PopZ: TPopupMenu
    Left = 164
    Top = 170
  end
  inherited SD: TSaveDialog
    Left = 100
    Top = 172
  end
  object Q1: THQuery
    PageCriteres = Pages
    MAJAuto = False
    Distinct = False
    Left = 56
    Top = 170
  end
end

inherited FGRStatMR: TFGRStatMR
  Left = 49
  Top = 153
  HelpContext = 301100280
  Caption = 'Répartition par mode de paiement'
  PixelsPerInch = 96
  TextHeight = 13
  inherited P: TPanel
    Top = 97
    Height = 214
    inherited Splitter1: TSplitter
      Height = 115
    end
    inherited FListe: THGrid
      ColCount = 4
      FixedCols = 0
      ScrollBars = ssBoth
      OnDblClick = FListeDblClick
      SortedCol = 2
      Titres.Strings = (
        'Mode Règlement'
        'Nombre règlements'
        'Montant')
      ColWidths = (
        130
        113
        92
        97)
    end
    inherited FChart1: TChart
      Height = 115
      OnClickSeries = FChartClickSeries
      OnGetLegendText = FChart1GetLegendText
    end
    inherited FChart2: TChart
      Height = 115
    end
  end
  inherited Pages: TPageControl
    Height = 97
    TabHeight = 1
    TabWidth = 1
    inherited TabSheet1: TTabSheet
      TabVisible = True
      object TGP_DATEPIECE: THLabel
        Left = 7
        Top = 5
        Width = 52
        Height = 13
        Caption = '&Date pièce'
        FocusControl = GP_DATEPIECE
      end
      object TGP_DATEPIECE_: THLabel
        Left = 196
        Top = 5
        Width = 15
        Height = 13
        Caption = '&au '
        FocusControl = GP_DATEPIECE_
      end
      object TGP_CAISSE: THLabel
        Left = 7
        Top = 51
        Width = 31
        Height = 13
        Caption = '&Caisse'
        FocusControl = GP_CAISSE
      end
      object TGP_NUMZCAISSE: THLabel
        Left = 7
        Top = 28
        Width = 73
        Height = 13
        Caption = '&Nº de la clôture'
        FocusControl = GP_NUMZCAISSE
      end
      object TGP_NUMZCAISSE_: THLabel
        Left = 196
        Top = 28
        Width = 12
        Height = 13
        Caption = 'a&u'
        FocusControl = GP_NUMZCAISSE_
      end
      object GP_DATEPIECE: THCritMaskEdit
        Left = 93
        Top = 1
        Width = 93
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
      object GP_DATEPIECE_: THCritMaskEdit
        Left = 216
        Top = 1
        Width = 93
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
      object GP_CAISSE: THValComboBox
        Left = 93
        Top = 47
        Width = 93
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 4
        OnChange = GP_DATEPIECEChange
        TagDispatch = 0
        Vide = True
        VideString = '<<Tous>>'
        DataType = 'GCPCAISSE'
      end
      object GP_NUMZCAISSE: THCritMaskEdit
        Left = 93
        Top = 24
        Width = 93
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 2
        OnChange = GP_DATEPIECEChange
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object GP_NUMZCAISSE_: THCritMaskEdit
        Left = 216
        Top = 24
        Width = 93
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 3
        OnChange = GP_DATEPIECEChange
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object XX_WHERE: TEdit
        Left = 272
        Top = 47
        Width = 40
        Height = 21
        Color = clYellow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 5
        Text = 'GP_NATUREPIECEG="FFO" and GP_NUMZCAISSE>0'
        Visible = False
      end
      object GVOIR: TGroupBox
        Left = 316
        Top = 1
        Width = 121
        Height = 66
        Caption = 'Afficher en'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 6
        object BMONTANT: TRadioButton
          Left = 8
          Top = 20
          Width = 105
          Height = 17
          Caption = 'Montant'
          Checked = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          TabStop = True
        end
        object BNOMBRE: TRadioButton
          Left = 8
          Top = 40
          Width = 105
          Height = 17
          Caption = 'Nombre'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
      end
      object GETIQ: TGroupBox
        Left = 448
        Top = 1
        Width = 121
        Height = 66
        Caption = 'Etiquettes en'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 7
        object BPOURCENT: TRadioButton
          Left = 8
          Top = 20
          Width = 105
          Height = 17
          Caption = 'Pourcentage'
          Checked = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          TabStop = True
        end
        object BVALEUR: TRadioButton
          Left = 8
          Top = 40
          Width = 105
          Height = 17
          Caption = 'Valeur'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
      end
      object GP_ETABLISSEMENT: THValComboBox
        Left = 216
        Top = 47
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
    Top = 126
  end
  inherited Msg: THMsgBox
    Top = 126
  end
  inherited HMTrad: THSystemMenu
    ResizeDBGrid = True
    Top = 126
  end
  inherited FindDialog: TFindDialog
    Top = 126
  end
  inherited PopAff: TPopupMenu
    Top = 126
  end
  inherited PopZ: TPopupMenu
    Top = 126
  end
  inherited SD: TSaveDialog
    Left = 476
    Top = 128
  end
  object Q1: THQuery
    PageCriteres = Pages
    MAJAuto = False
    Distinct = False
    Left = 96
    Top = 126
  end
end

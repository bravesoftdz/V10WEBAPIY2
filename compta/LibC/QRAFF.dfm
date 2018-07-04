inherited FQRAFF: TFQRAFF
  Left = 57
  Top = 172
  Width = 944
  Height = 664
  Cursor = crHourGlass
  VertScrollBar.Position = 387
  Caption = 'Etat des marges'
  PixelsPerInch = 96
  TextHeight = 13
  inherited Pages: TPageControl
    Top = -387
    Height = 213
    ActivePage = TabRuptures
    inherited Standards: TTabSheet
      inherited TFGenJoker: THLabel
        Left = 7
        Top = 44
        Caption = '&Sections'
      end
      inherited TFGen: THLabel
        Left = 7
        Top = 44
        Caption = '&Sections de'
      end
      inherited TFaG: TLabel
        Left = 267
        Top = 44
      end
      inherited Label7: TLabel
        Left = 264
        Top = 118
      end
      inherited Hlabel2: THLabel
        Left = 7
        Top = 18
        Caption = '&Axe'
      end
      inherited HLabel4: THLabel
        Left = 7
        Top = 93
      end
      inherited HLabel5: THLabel
        Left = 7
        Top = 143
      end
      inherited HLabel6: THLabel
        Left = 7
        Top = 118
      end
      inherited TSelectCpte: THLabel
        Left = 7
        Top = 68
        Width = 113
      end
      inherited FJoker: TEdit
        Left = 147
        Top = 40
        TabOrder = 2
      end
      inherited FCpte1: THCpteEdit
        Left = 147
        Top = 40
        TabOrder = 1
      end
      inherited FNatureEcr: THValComboBox
        Left = 147
        Top = 139
        TabOrder = 8
      end
      inherited FDateCompta2: TMaskEdit
        Left = 284
        Top = 114
        TabOrder = 7
      end
      inherited FDateCompta1: TMaskEdit
        Left = 147
        Top = 114
        TabOrder = 6
        OnChange = FDateCompta1Change
      end
      inherited FExercice: THValComboBox
        Left = 147
        Top = 89
        TabOrder = 5
      end
      inherited FNatureCpt: THValComboBox
        Left = 147
        Top = 14
        Width = 110
        TabOrder = 0
        Vide = False
      end
      inherited FCpte2: THCpteEdit
        Left = 284
        Top = 40
        TabOrder = 3
      end
      inherited FSelectCpte: THValComboBox
        Left = 147
        Top = 64
        TabOrder = 4
        OnChange = FSelectCpteChange
      end
      object FAvecAN: TCheckBox
        Left = 7
        Top = 163
        Width = 152
        Height = 17
        Alignment = taLeftJustify
        Caption = '&Inclure le "cumul au"'
        Checked = True
        State = cbChecked
        TabOrder = 9
        Visible = False
        OnClick = FAvecANClick
      end
    end
    inherited Avances: TTabSheet
      inherited HLabel7: THLabel
        Left = 7
        Top = 27
      end
      inherited HLabel8: THLabel
        Left = 7
        Top = 73
      end
      inherited HLabel1: THLabel
        Left = 7
        Top = 122
        Width = 111
        Caption = '&Sections d'#39'exception'
      end
      inherited FDevises: THValComboBox
        Left = 150
        Top = 69
      end
      inherited FEtab: THValComboBox
        Left = 150
        Top = 23
      end
      inherited FExcep: TEdit
        Left = 150
        Top = 118
      end
      inherited AvecRevision: TCheckBox
        Left = 5
        Width = 158
        Alignment = taLeftJustify
        Caption = 'Ecritures de &révisions'
      end
      inherited OnCum: TCheckBox
        Left = 296
        Width = 99
        Alignment = taLeftJustify
        Caption = 'Sur &cumuls'
      end
    end
    inherited tabSup: TTabSheet [2]
      Caption = 'Colonnes'
      object TNatLib: TLabel
        Left = 27
        Top = 15
        Width = 163
        Height = 13
        Caption = '&Table libre utilisée sur les colonnes'
        FocusControl = NatLib
      end
      object NatLib: THValComboBox
        Left = 202
        Top = 11
        Width = 183
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 0
        TagDispatch = 0
      end
      object OKC1: TCheckBox
        Left = 27
        Top = 46
        Width = 163
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Utilisation 1° colonne'
        TabOrder = 1
        OnClick = OKC1Click
      end
      object OkC2: TCheckBox
        Left = 27
        Top = 78
        Width = 163
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Utilisation 2° colonne'
        TabOrder = 3
        OnClick = OkC2Click
      end
      object OkC3: TCheckBox
        Left = 27
        Top = 109
        Width = 163
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Utilisation 3° colonne'
        TabOrder = 5
        OnClick = OkC3Click
      end
      object OkC4: TCheckBox
        Left = 27
        Top = 140
        Width = 163
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Utilisation 4° colonne'
        TabOrder = 7
        OnClick = OkC4Click
      end
      object SPC1: TEdit
        Left = 202
        Top = 44
        Width = 183
        Height = 21
        ReadOnly = True
        TabOrder = 2
        Visible = False
        OnChange = SPC1Change
        OnDblClick = SPC1DblClick
      end
      object SPC2: TEdit
        Left = 202
        Top = 76
        Width = 183
        Height = 21
        ReadOnly = True
        TabOrder = 4
        Visible = False
        OnChange = SPC1Change
        OnDblClick = SPC1DblClick
      end
      object SPC3: TEdit
        Left = 202
        Top = 107
        Width = 183
        Height = 21
        ReadOnly = True
        TabOrder = 6
        Visible = False
        OnChange = SPC1Change
        OnDblClick = SPC1DblClick
      end
      object SPC4: TEdit
        Left = 202
        Top = 138
        Width = 183
        Height = 21
        ReadOnly = True
        TabOrder = 8
        Visible = False
        OnChange = SPC1Change
        OnDblClick = SPC1DblClick
      end
    end
    inherited TabRuptures: TTabSheet [3]
      object TNbCharRupt: TLabel [0]
        Left = 208
        Top = 60
        Width = 149
        Height = 13
        Caption = 'Nombre de chiffres pour rupture'
        FocusControl = NbCharRupt
      end
      inherited FSautPageRupt: TCheckBox
        TabOrder = 5
      end
      object NbCharRupt: TSpinEdit
        Left = 363
        Top = 56
        Width = 32
        Height = 22
        MaxLength = 17
        MaxValue = 17
        MinValue = 0
        TabOrder = 2
        Value = 6
      end
    end
    inherited Option: TTabSheet [4]
      inherited Apercu: TCheckBox
        Top = 153
      end
      inherited FListe: TCheckBox
        Top = 153
      end
      inherited GroupBox3: TGroupBox
        Top = 3
        Height = 89
        inherited Bevel1: TBevel
          Height = 89
        end
        inherited FLigneCpt: TCheckBox
          Left = 8
          Top = 43
          TabOrder = 2
        end
        inherited Reduire: TCheckBox
          Top = 43
          Caption = '2 pag&es sur 1'
          TabOrder = 3
        end
        inherited FTrait: TCheckBox
          Left = 8
          Top = 19
          TabOrder = 0
        end
        inherited FRappelCrit: TCheckBox
          Left = 8
          Top = 66
          TabOrder = 4
        end
        object FLigneRupt: TCheckBox
          Left = 197
          Top = 19
          Width = 172
          Height = 17
          Alignment = taLeftJustify
          Caption = '&Séparateurs de &ruptures'
          Checked = True
          State = cbChecked
          TabOrder = 1
        end
        object SautDePage: TCheckBox
          Left = 197
          Top = 66
          Width = 172
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Saut de page table libre'
          Checked = True
          State = cbChecked
          TabOrder = 5
        end
      end
      inherited FCouleur: TCheckBox
        Top = 153
        TabOrder = 4
      end
      inherited FGroupChoixRupt: TGroupBox
        Top = 102
        TabOrder = 3
      end
    end
    inherited Mise: TTabSheet [5]
      inherited GroupBox1: TGroupBox
        Top = 67
        Height = 101
        TabOrder = 1
        inherited Bevel2: TBevel
          Height = 101
        end
        inherited FMonetaire: TCheckBox
          Top = 41
          TabOrder = 3
        end
        inherited FReport: TCheckBox
          Top = 60
          Checked = False
          Enabled = False
          State = cbUnchecked
          TabOrder = 5
        end
        inherited Avance: TCheckBox
          Top = 101
          TabOrder = 6
        end
        object FCol1: TCheckBox
          Left = 8
          Top = 21
          Width = 172
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Colonne 1'
          Enabled = False
          TabOrder = 0
        end
        object FCol2: TCheckBox
          Left = 8
          Top = 41
          Width = 172
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Colonne 2'
          Enabled = False
          TabOrder = 2
        end
        object FCol3: TCheckBox
          Left = 8
          Top = 60
          Width = 172
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Colonne 3'
          Enabled = False
          TabOrder = 4
        end
        object FCol4: TCheckBox
          Left = 197
          Top = 21
          Width = 172
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Colonne 4'
          Enabled = False
          TabOrder = 1
        end
      end
      inherited FMontant: TRadioGroup
        Top = 15
        Height = 46
        TabOrder = 0
      end
    end
  end
  inherited Panel1: TPanel
    Top = -31
    inherited EnteteAutrePage: TQRBand
      Top = 258
    end
    inherited BDetail: TQRBand
      Top = 345
      AfterPrint = BDetailAfterPrint
      BeforePrint = BDetailBeforePrint
      LinkBand = BDSMulti
      object S_SECTION: TQRDBText
        Tag = 1
        Left = 0
        Top = 2
        width = 100
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Caption = 'S_SECTION'
        DataSource = S
        DataField = 'S_SECTION'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object S_LIBELLE: TQRDBText
        Tag = 2
        Left = 101
        Top = 2
        width = 155
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Caption = 'S_LIBELLE'
        DataSource = S
        DataField = 'S_LIBELLE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object ANvDEBIT: TQRLabel
        Tag = 3
        Left = 257
        Top = 2
        width = 99
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'ANvDEBIT'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object ANvCREDIT: TQRLabel
        Tag = 4
        Left = 357
        Top = 2
        width = 99
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'ANvCREDIT'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object DEBIT: TQRLabel
        Tag = 5
        Left = 457
        Top = 2
        width = 99
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'DEBIT'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object CREDIT: TQRLabel
        Tag = 6
        Left = 557
        Top = 2
        width = 99
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'CREDIT'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object DEBITsum: TQRLabel
        Tag = 7
        Left = 657
        Top = 2
        width = 99
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'DEBITsum'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object CREDITsum: TQRLabel
        Tag = 8
        Left = 757
        Top = 2
        width = 99
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'CREDITsum'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object SOLDEdeb: TQRLabel
        Tag = 9
        Left = 857
        Top = 2
        width = 99
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'SOLDEdeb'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object SOLDEcre: TQRLabel
        Tag = 10
        Left = 957
        Top = 2
        width = 99
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'SOLDEcre'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
    end
    inherited BFinEtat: TQRBand
      Top = 467
      Height = 26
      object QRLabel33: TQRLabel
        Tag = 2
        Left = 101
        Top = 6
        width = 155
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        AutoSize = False
        Caption = 'Total général'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOT2DEBITanv: TQRLabel
        Tag = 3
        Left = 257
        Top = 6
        width = 99
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOT2DEBITanv'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOT2CREDITanv: TQRLabel
        Tag = 4
        Left = 357
        Top = 6
        width = 99
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOT2CREDITanv'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOT2DEBIT: TQRLabel
        Tag = 5
        Left = 457
        Top = 6
        width = 99
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOT2DEBIT'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOT2CREDIT: TQRLabel
        Tag = 6
        Left = 557
        Top = 6
        width = 99
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOT2CREDIT'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOT2DEBITsum: TQRLabel
        Tag = 7
        Left = 657
        Top = 6
        width = 99
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOT2DEBITsum'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOT2CREDITsum: TQRLabel
        Tag = 8
        Left = 757
        Top = 6
        width = 99
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOT2CREDITsum'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOT2SOLdeb: TQRLabel
        Tag = 9
        Left = 857
        Top = 6
        width = 99
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOT2SOLdeb'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOT2SOLcre: TQRLabel
        Tag = 10
        Left = 957
        Top = 6
        width = 99
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOT2SOLcre'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
    end
    inherited EntetePage: TQRBand
      Top = 316
      Height = 29
      object TitreColCpt: TQRLabel
        Tag = -1
        Left = 0
        Top = 0
        width = 100
        height = 29
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taCenter
        AutoSize = False
        Caption = 'Section'
        AlignToBand = False
        Color = clTeal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = True
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TitreColLibelle: TQRLabel
        Tag = -2
        Left = 101
        Top = 0
        width = 155
        height = 29
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taCenter
        AutoSize = False
        Caption = 'Libellé'
        AlignToBand = False
        Color = clTeal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = True
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TTitreColAvant: TQRLabel
        Tag = 3
        Left = 257
        Top = 0
        width = 199
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taCenter
        AutoSize = False
        Caption = 'Total'
        AlignToBand = False
        Color = clTeal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 100
        SynTitreCol = True
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QrC11: TQRLabel
        Tag = -3
        Left = 257
        Top = 15
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taCenter
        AutoSize = False
        Caption = 'Produit'
        AlignToBand = False
        Color = clTeal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 3
        SynTitreCol = True
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRC12: TQRLabel
        Tag = -4
        Left = 357
        Top = 15
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taCenter
        AutoSize = False
        Caption = 'Charge'
        AlignToBand = False
        Color = clTeal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 3
        SynTitreCol = True
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TTitreColSelection: TQRLabel
        Tag = 5
        Left = 457
        Top = 0
        width = 199
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taCenter
        AutoSize = False
        Caption = 'Marge'
        AlignToBand = False
        Color = clTeal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 100
        SynTitreCol = True
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRC21: TQRLabel
        Tag = -5
        Left = 457
        Top = 15
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taCenter
        AutoSize = False
        Caption = 'En valeur'
        AlignToBand = False
        Color = clTeal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 4
        SynTitreCol = True
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRC22: TQRLabel
        Tag = -6
        Left = 557
        Top = 15
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taCenter
        AutoSize = False
        Caption = 'En %'
        AlignToBand = False
        Color = clTeal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 4
        SynTitreCol = True
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TTitreColApres: TQRLabel
        Tag = 7
        Left = 657
        Top = 0
        width = 199
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taCenter
        AutoSize = False
        Caption = 'Mouvements au :'
        AlignToBand = False
        Color = clTeal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 100
        SynTitreCol = True
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRC31: TQRLabel
        Tag = -7
        Left = 657
        Top = 15
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taCenter
        AutoSize = False
        Caption = 'Débit'
        AlignToBand = False
        Color = clTeal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 5
        SynTitreCol = True
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRC32: TQRLabel
        Tag = -8
        Left = 757
        Top = 15
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taCenter
        AutoSize = False
        Caption = 'Crédit'
        AlignToBand = False
        Color = clTeal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 5
        SynTitreCol = True
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRC41: TQRLabel
        Tag = -9
        Left = 857
        Top = 15
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taCenter
        AutoSize = False
        Caption = 'Débit'
        AlignToBand = False
        Color = clTeal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 6
        SynTitreCol = True
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object QRC42: TQRLabel
        Tag = -10
        Left = 957
        Top = 15
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taCenter
        AutoSize = False
        Caption = 'Crédit'
        AlignToBand = False
        Color = clTeal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 6
        SynTitreCol = True
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TTitreColSolde: TQRLabel
        Tag = 9
        Left = 857
        Top = 0
        width = 199
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taCenter
        AutoSize = False
        Caption = 'Solde au :'
        AlignToBand = False
        Color = clTeal
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWhite
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 100
        SynTitreCol = True
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
    end
    inherited TOPREPORT: TQRBand
      Top = 294
      BeforePrint = TOPREPORTBeforePrint
      inherited TITREREPORTH: TQRLabel
        Tag = 2
        Left = 101
        width = 155
      end
      object REPORT1DEBITanv: TQRLabel
        Tag = 3
        Left = 257
        Top = 0
        width = 99
        height = 17
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'REPORT1DEBITanv'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object REPORT1CREDITanv: TQRLabel
        Tag = 4
        Left = 357
        Top = 0
        width = 99
        height = 17
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'REPORT1CREDITanv'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object REPORT1DEBIT: TQRLabel
        Tag = 5
        Left = 457
        Top = 0
        width = 99
        height = 17
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'REPORT1DEBIT'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object REPORT1CREDIT: TQRLabel
        Tag = 6
        Left = 557
        Top = 0
        width = 99
        height = 17
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'REPORT1CREDIT'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object REPORT1DEBITsum: TQRLabel
        Tag = 7
        Left = 657
        Top = 0
        width = 99
        height = 17
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'REPORT1DEBITsum'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object REPORT1CREDITsum: TQRLabel
        Tag = 8
        Left = 757
        Top = 0
        width = 99
        height = 17
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'REPORT1CREDITsum'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object REPORT1DEBITsol: TQRLabel
        Tag = 9
        Left = 857
        Top = 0
        width = 99
        height = 17
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'REPORT1DEBITsol'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object REPORT1CREDITsol: TQRLabel
        Tag = 10
        Left = 957
        Top = 0
        width = 99
        height = 17
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'REPORT1CREDITsol'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
    end
    inherited BOTTOMREPORT: TQRBand
      Top = 593
      Height = 18
      BeforePrint = BOTTOMREPORTBeforePrint
      inherited TITREREPORTB: TQRLabel
        Tag = 2
        Left = 101
        width = 155
      end
      object REPORT2DEBITanv: TQRLabel
        Tag = 3
        Left = 257
        Top = 0
        width = 99
        height = 17
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'REPORT2DEBITanv'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object REPORT2CREDITanv: TQRLabel
        Tag = 4
        Left = 357
        Top = 0
        width = 99
        height = 17
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'REPORT2CREDITanv'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object REPORT2DEBIT: TQRLabel
        Tag = 5
        Left = 457
        Top = 0
        width = 99
        height = 17
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'REPORT2DEBIT'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object REPORT2CREDIT: TQRLabel
        Tag = 6
        Left = 557
        Top = 0
        width = 99
        height = 17
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'REPORT2CREDIT'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object REPORT2CREDITsum: TQRLabel
        Tag = 8
        Left = 757
        Top = 0
        width = 99
        height = 17
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'REPORT2CREDITsum'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object REPORT2DEBITsum: TQRLabel
        Tag = 7
        Left = 657
        Top = 0
        width = 99
        height = 17
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'REPORT2DEBIT'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object REPORT2DEBITsol: TQRLabel
        Tag = 9
        Left = 857
        Top = 0
        width = 99
        height = 17
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'REPORT2DEBITsol'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object REPORT2CREDITsol: TQRLabel
        Tag = 10
        Left = 957
        Top = 0
        width = 99
        height = 17
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'REPORT2CREDITsol'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
    end
    inherited BTitre: TQRBand
      Height = 108
      inherited RCpte: TQRLabel [0]
        Top = 57
        width = 65
        Caption = 'Sections de'
      end
      inherited RJoker: TQRLabel [1]
        Top = 57
        width = 48
        Caption = 'Sections'
      end
      inherited TTitre: TQRSysData [2]
      end
      inherited TRa: TQRLabel
        Top = 57
      end
      inherited RCpte1: TQRLabel
        Top = 57
      end
      inherited RCpte2: TQRLabel
        Top = 57
      end
      inherited RExcepGen: TQRLabel
        Top = 90
      end
      inherited QRLabel4: TQRLabel
        Top = 57
      end
      inherited QRLabel5: TQRLabel
        SynCritere = 201
      end
      inherited QRLabel7: TQRLabel
        Top = 73
      end
      inherited QRLabel11: TQRLabel
        Top = 57
      end
      inherited QRLabel15: TQRLabel
        Left = 719
        SynCritere = 301
      end
      inherited RDateCompta1: TQRLabel
        Top = 57
      end
      inherited RDateCompta2: TQRLabel
        Top = 57
      end
      inherited RExercice: TQRLabel
        Left = 480
        SynCritere = 201
      end
      inherited RNatureEcr: TQRLabel
        Top = 73
      end
      inherited REtab: TQRLabel
        Left = 840
        SynCritere = 301
      end
      inherited QRLabel18: TQRLabel
        Left = 719
        Top = 57
        SynCritere = 301
      end
      inherited RDevises: TQRLabel
        Left = 840
        Top = 57
        SynCritere = 301
      end
      inherited QRLabel12: TQRLabel
        Top = 41
        width = 21
        Caption = 'Axe'
      end
      inherited RNatureCpt: TQRLabel
        Top = 41
      end
      inherited QRLabel24: TQRLabel
        Top = 89
      end
      inherited RSelectCpte: TQRLabel
        Top = 73
      end
      inherited TRSelectCpte: TQRLabel
        Top = 73
      end
    end
    object BDSMulti: TQRBand
      Left = 1
      Top = 363
      Width = 1058
      Height = 36
      Align = alTop
      BandType = rbSubDetail
      BeforePrint = BDSMultiBeforePrint
      Color = clWhite
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clOlive
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = [fsBold]
      ForceNewPage = False
      RAZPage = False
      Frame.DrawTop = False
      Frame.DrawBottom = False
      Frame.DrawLeft = False
      Frame.DrawRight = False
      ParentFont = False
      Ruler = qrrNone
      PrintOnBottom = False
      object TOTST3: TQRLabel
        Tag = 5
        Left = 457
        Top = 3
        width = 99
        height = 13
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTST3'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTST4: TQRLabel
        Tag = 6
        Left = 557
        Top = 3
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTST4'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TCodeST: TQRLabel
        Tag = 1
        Left = 0
        Top = 3
        width = 253
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        AutoSize = False
        Caption = 'TCodeST'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTST7: TQRLabel
        Tag = 9
        Left = 857
        Top = 3
        width = 99
        height = 13
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTST7'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTST8: TQRLabel
        Tag = 10
        Left = 957
        Top = 3
        width = 99
        height = 13
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTST8'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTST5: TQRLabel
        Tag = 7
        Left = 657
        Top = 3
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTST5'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTST6: TQRLabel
        Tag = 8
        Left = 757
        Top = 3
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTST6'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTST1: TQRLabel
        Tag = 3
        Left = 257
        Top = 3
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTST1'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTST2: TQRLabel
        Tag = 4
        Left = 357
        Top = 3
        width = 99
        height = 13
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTST2'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
    end
    object BRUPT: TQRBand
      Left = 1
      Top = 399
      Width = 1058
      Height = 34
      Align = alTop
      BandType = rbSubDetail
      BeforePrint = BRUPTBeforePrint
      Color = clWhite
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clOlive
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = [fsBold]
      ForceNewPage = False
      RAZPage = False
      Frame.DrawTop = False
      Frame.DrawBottom = False
      Frame.DrawLeft = False
      Frame.DrawRight = False
      ParentFont = False
      Ruler = qrrNone
      PrintOnBottom = False
      object TOTT3: TQRLabel
        Tag = 5
        Left = 457
        Top = 3
        width = 99
        height = 13
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTT3'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clPurple
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold, fsItalic]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTT4: TQRLabel
        Tag = 6
        Left = 557
        Top = 3
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTT4'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clPurple
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold, fsItalic]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TCodRupt: TQRLabel
        Tag = 1
        Left = 0
        Top = 3
        width = 253
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        AutoSize = False
        Caption = 'TCodRupt'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clPurple
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold, fsItalic]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTT7: TQRLabel
        Tag = 9
        Left = 857
        Top = 3
        width = 99
        height = 13
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTT7'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clPurple
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold, fsItalic]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTT8: TQRLabel
        Tag = 10
        Left = 957
        Top = 3
        width = 99
        height = 13
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTT8'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clPurple
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold, fsItalic]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTT5: TQRLabel
        Tag = 7
        Left = 657
        Top = 3
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTT5'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clPurple
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold, fsItalic]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTT6: TQRLabel
        Tag = 8
        Left = 757
        Top = 3
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTT6'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clPurple
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold, fsItalic]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTT1: TQRLabel
        Tag = 3
        Left = 257
        Top = 3
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTT1'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clPurple
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold, fsItalic]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTT2: TQRLabel
        Tag = 4
        Left = 357
        Top = 3
        width = 99
        height = 13
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTT2'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clPurple
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold, fsItalic]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
    end
    object QRBand1: TQRBand
      Left = 1
      Top = 493
      Width = 1058
      Height = 17
      Align = alTop
      BandType = rbOverlay
      Color = clWhite
      ForceNewPage = False
      RAZPage = False
      Frame.DrawTop = False
      Frame.DrawBottom = False
      Frame.DrawLeft = False
      Frame.DrawRight = False
      Ruler = qrrNone
      PrintOnBottom = False
      object Trait2: TQRLigne
        Tag = 2
        Left = 256
        Top = 0
        width = 5
        Height = 17
        SynLigne = 0
        NoRepeatValue = False
        TopBand = EntetePage
        BottomBand = BDetail
      end
      object Trait3: TQRLigne
        Tag = 3
        Left = 456
        Top = 0
        width = 5
        Height = 17
        SynLigne = 0
        NoRepeatValue = False
        TopBand = EntetePage
        BottomBand = BDetail
      end
      object Trait4: TQRLigne
        Tag = 4
        Left = 656
        Top = 0
        width = 5
        Height = 17
        SynLigne = 0
        NoRepeatValue = False
        TopBand = EntetePage
        BottomBand = BDetail
      end
      object Trait5: TQRLigne
        Tag = 5
        Left = 856
        Top = 0
        width = 5
        Height = 17
        SynLigne = 0
        NoRepeatValue = False
        TopBand = EntetePage
        BottomBand = BDetail
      end
      object Trait0: TQRLigne
        Tag = -1
        Left = 0
        Top = 0
        width = 5
        Height = 17
        SynLigne = 0
        NoRepeatValue = False
        TopBand = EntetePage
        BottomBand = BDetail
      end
      object Trait6: TQRLigne
        Tag = 6
        Left = 1056
        Top = 0
        width = 5
        Height = 17
        SynLigne = 0
        NoRepeatValue = False
        TopBand = EntetePage
        BottomBand = BDetail
      end
      object Ligne1: TQRLigne
        Tag = -10
        Left = 0
        Top = 1
        width = 1058
        Height = 5
        SynLigne = 0
        NoRepeatValue = False
        TopBand = BDetail
        BottomBand = BDetail
        HorzBand = BDetail
      end
    end
    object BRecapG: TQRBand
      Left = 1
      Top = 433
      Width = 1058
      Height = 34
      AfterPrint = BRecapGAfterPrint
      Align = alTop
      BandType = rbSubDetail
      BeforePrint = BRecapGBeforePrint
      Color = clWhite
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clOlive
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = [fsBold]
      ForceNewPage = True
      RAZPage = False
      Frame.DrawTop = False
      Frame.DrawBottom = False
      Frame.DrawLeft = False
      Frame.DrawRight = False
      ParentFont = False
      Ruler = qrrNone
      PrintOnBottom = False
      object TOTR3: TQRLabel
        Tag = 5
        Left = 457
        Top = 3
        width = 99
        height = 13
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTR3'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTR4: TQRLabel
        Tag = 6
        Left = 557
        Top = 3
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTR4'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TCodeRecap: TQRLabel
        Tag = 2
        Left = 0
        Top = 3
        width = 66
        height = 15
        SynLigne = 0
        NoRepeatValue = False
        Caption = 'TCodeRecap'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTR7: TQRLabel
        Tag = 9
        Left = 857
        Top = 3
        width = 99
        height = 13
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTR7'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTR8: TQRLabel
        Tag = 10
        Left = 957
        Top = 3
        width = 99
        height = 13
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTR8'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTR5: TQRLabel
        Tag = 7
        Left = 657
        Top = 3
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTR5'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTR6: TQRLabel
        Tag = 8
        Left = 757
        Top = 3
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTR6'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTR1: TQRLabel
        Tag = 3
        Left = 257
        Top = 3
        width = 99
        height = 14
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTR1'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
      object TOTR2: TQRLabel
        Tag = 4
        Left = 357
        Top = 3
        width = 99
        height = 13
        SynLigne = 0
        NoRepeatValue = False
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TOTR2'
        AlignToBand = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -12
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        SynColGroup = 0
        SynTitreCol = False
        SynCritere = 0
        Frame.DrawTop = False
        Frame.DrawBottom = False
        Frame.DrawLeft = False
        Frame.DrawRight = False
      end
    end
  end
  inherited Formateur: THNumEdit
    Top = -305
  end
  inherited Dock971: TDock97
    Top = 541
    inherited PanelFiltre: TToolWindow97
      inherited FFiltres: THValComboBox
        Width = 326
      end
    end
  end
  inherited Q: TQuery
    AfterOpen = QAfterOpen
    SQL.Strings = (
      'Select '
      'S_AXE, S_SECTION, S_LIBELLE, '
      'S_TOTALDEBIT, S_TOTALCREDIT, S_SECTIONTRIE'
      'From SECTION'
      'Order by S_AXE, S_SECTION')
    Left = 551
    Top = 148
  end
  inherited S: TDataSource
    Left = 502
    Top = 148
  end
  object QRDLMulti: TQRDetailLink [6]
    OnNeedData = QRDLMultiNeedData
    DetailBand = BDSMulti
    Master = QRP
    PrintBefore = False
    IsRupture = False
    Resume = False
    Left = 747
    Top = 148
  end
  object DLrupt: TQRDetailLink [7]
    OnNeedData = DLruptNeedData
    DetailBand = BRUPT
    Master = QRP
    PrintBefore = True
    IsRupture = True
    Resume = False
    Left = 649
    Top = 148
  end
  inherited QRP: TQuickReport
    Left = 698
    Top = 148
  end
  inherited MsgRien: THMsgBox
    Mess.Strings = (
      
        '0;?caption?;Les critères que vous avez sélectionnés ne renvoient' +
        ' aucun enregistrement.;W;O;O;O;'
      
        '1;?caption?;Les dates de début et de fin d'#39'exercice sont incohér' +
        'entes;W;O;O;O;'
      
        '2;?caption?;Les dates de début sont incohérentes avec l'#39'exercice' +
        ' que vous avez sélectionné;W;O;O;O;'
      '3;?caption?; : ce caractère n'#39'est pas admis.;W;O;O;O;'
      'Société :'
      'Utilisateur :'
      
        '6;?caption?;Vous ne pouvez pas choisir les "Comptes non soldés";' +
        'W;O;O;O;'
      
        '7;?caption?;L'#39' axe que vous avez choisi n'#39'est pas structuré pour' +
        ' les ruptures;W;O;O;O;'
      
        '8;?caption?;Il n'#39'existe a pas d'#39'enchaînement pour ce plan de rup' +
        'tures;W;O;O;O;'
      
        '9;?caption?;Vous devez sélectionner une colonne de montants;W;O;' +
        'O;O;'
      ''
      '')
    Left = 524
    Top = 28
  end
  inherited HMCrit: THMsgBox
    Left = 478
    Top = 28
  end
  inherited MsgQR: THMsgBox
    Left = 569
    Top = 28
  end
  inherited HMTrad: THSystemMenu
    Left = 600
    Top = 148
  end
  inherited POPF: TPopupMenu
    Left = 453
    Top = 148
  end
  object MsgLibel: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'au'
      'du'
      'Solde au'
      'Société :'
      'Utilisateur :'
      'Solde débiteur  :'
      'Solde créditeur :'
      'REPORT'
      'Société'
      'Utilisateur'
      'Rupture... '
      'Rupture par compte de correspondance... '
      '12 ;'
      'Total'
      'sur la table libre n°'
      'Colonne'
      'Total')
    Left = 433
    Top = 28
  end
  object QSup: TQuery
    DatabaseName = 'SOC'
    Left = 532
    Top = 208
  end
  object QRDLRecapG: TQRDetailLink
    OnNeedData = QRDLRecapGNeedData
    DetailBand = BRecapG
    Master = QRP
    PrintBefore = False
    IsRupture = False
    Resume = True
    Left = 796
    Top = 148
  end
end

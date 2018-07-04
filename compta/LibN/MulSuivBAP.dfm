inherited FMulSuivBAP: TFMulSuivBAP
  Left = 196
  Top = 193
  Width = 658
  Height = 386
  Caption = '85'
  PixelsPerInch = 96
  TextHeight = 13
  inherited Pages: THPageControl2
    Width = 650
    Height = 117
    ActivePage = PCritere
    inherited PCritere: THTabSheet
      inherited Bevel1: TBevel
        Width = 642
        Height = 89
      end
      object HLabel3: THLabel
        Left = 310
        Top = 66
        Width = 82
        Height = 13
        Caption = 'Etat de validation'
        FocusControl = E_TYPEANOUVEAU
      end
      object HLabel1: THLabel
        Left = 9
        Top = 11
        Width = 51
        Height = 13
        Caption = '&Nom lot de'
        FocusControl = E_NOMLOT
      end
      object HLabel5: THLabel
        Left = 181
        Top = 11
        Width = 6
        Height = 13
        Caption = #224
      end
      object HLabel8: THLabel
        Left = 181
        Top = 39
        Width = 6
        Height = 13
        Caption = #224
      end
      object HLabel6: THLabel
        Left = 9
        Top = 39
        Width = 57
        Height = 13
        Caption = '&Banques de'
        FocusControl = E_BANQUEPREVI
      end
      object TE_AUXILIAIRE: THLabel
        Left = 9
        Top = 66
        Width = 56
        Height = 13
        Caption = '&Auxiliaire de'
        FocusControl = E_AUXILIAIRE
      end
      object HLabel4: THLabel
        Left = 181
        Top = 66
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_AUXILIAIRE_
      end
      object HLabel7: THLabel
        Left = 310
        Top = 11
        Width = 37
        Height = 13
        Caption = '&G'#233'n'#233'ral'
        FocusControl = E_GENERAL
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel9: THLabel
        Left = 458
        Top = 11
        Width = 6
        Height = 13
        Caption = #224
        FocusControl = E_GENERAL_
      end
      object HLabel10: THLabel
        Left = 310
        Top = 39
        Width = 74
        Height = 13
        Caption = '&Etape du circuit'
        FocusControl = E_SUIVDEC
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object E_TYPEANOUVEAU: THValComboBox
        Tag = 1
        Left = 402
        Top = 62
        Width = 169
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 13
        ParentCtl3D = False
        TabOrder = 0
        TagDispatch = 0
        Vide = True
        DataType = 'CPVALIDEBAP'
      end
      object E_NOMLOT: THCritMaskEdit
        Left = 72
        Top = 7
        Width = 98
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 1
        OnChange = E_NOMLOTChange
        TagDispatch = 0
        DataType = 'CPNOMLOT'
        Operateur = Superieur
        ElipsisButton = True
      end
      object E_NOMLOT_: THCritMaskEdit
        Left = 200
        Top = 7
        Width = 98
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 2
        OnChange = E_NOMLOTChange
        TagDispatch = 0
        DataType = 'CPNOMLOT'
        Operateur = Inferieur
        ElipsisButton = True
      end
      object E_BANQUEPREVI_: THCritMaskEdit
        Left = 200
        Top = 35
        Width = 98
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 3
        TagDispatch = 0
        DataType = 'TZGBANQUE'
        Operateur = Inferieur
        ElipsisButton = True
      end
      object E_BANQUEPREVI: THCritMaskEdit
        Left = 72
        Top = 35
        Width = 98
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 4
        TagDispatch = 0
        DataType = 'TZGBANQUE'
        Operateur = Superieur
        ElipsisButton = True
      end
      object E_AUXILIAIRE: THCritMaskEdit
        Left = 72
        Top = 62
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 5
        TagDispatch = 0
        DataType = 'TZTTOUTDEBIT'
        Operateur = Superieur
        ElipsisButton = True
      end
      object E_AUXILIAIRE_: THCritMaskEdit
        Left = 200
        Top = 62
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 6
        TagDispatch = 0
        DataType = 'TZTTOUTDEBIT'
        Operateur = Inferieur
        ElipsisButton = True
      end
      object E_GENERAL: THCritMaskEdit
        Left = 354
        Top = 7
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 17
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 7
        TagDispatch = 0
        DataType = 'TZGENCAIS'
        ElipsisButton = True
      end
      object E_GENERAL_: THCritMaskEdit
        Left = 473
        Top = 7
        Width = 98
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 17
        TabOrder = 8
        TagDispatch = 0
        DataType = 'TZTTOUTDEBIT'
        Operateur = Inferieur
        ElipsisButton = True
      end
      object E_SUIVDEC: THValComboBox
        Left = 402
        Top = 35
        Width = 169
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemHeight = 13
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 9
        TagDispatch = 0
        Plus = 'AND CO_LIBELLE<>""'
        Vide = True
        VideString = '<<Tous>>'
        DataType = 'CPCIRCUITDECAIS'
      end
    end
    object PCircuit: TTabSheet [1]
      Caption = 'Circuit de validation'
      ImageIndex = 6
      object Bevel7: TBevel
        Left = 0
        Top = 0
        Width = 642
        Height = 89
        Align = alClient
      end
      object Label3: TLabel
        Left = 277
        Top = 39
        Width = 17
        Height = 17
        Caption = #232
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -16
        Font.Name = 'Wingdings'
        Font.Style = []
        ParentFont = False
      end
      object g1: TGroupBox
        Left = 20
        Top = 10
        Width = 243
        Height = 67
        Caption = 'Etape d'#39'&origine'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clMaroon
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 0
        object HLabel11: THLabel
          Left = 11
          Top = 18
          Width = 74
          Height = 13
          Caption = '&Etape du circuit'
          FocusControl = CircuitSource
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel12: THLabel
          Left = 11
          Top = 42
          Width = 51
          Height = 13
          Caption = '&Nom du lot'
          FocusControl = LotSource
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object CircuitSource: THValComboBox
          Left = 91
          Top = 14
          Width = 137
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          TagDispatch = 0
          Plus = 'AND CO_LIBELLE<>""'
          Vide = True
          DataType = 'CPCIRCUITDECAIS'
        end
        object LotSource: THCritMaskEdit
          Left = 91
          Top = 38
          Width = 133
          Height = 21
          CharCase = ecUpperCase
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 13
          ParentFont = False
          TabOrder = 1
          TagDispatch = 0
          Operateur = Egal
        end
      end
      object g2: TGroupBox
        Left = 307
        Top = 0
        Width = 243
        Height = 85
        Caption = 'Etape de &destination'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clTeal
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object HLabel13: THLabel
          Left = 11
          Top = 15
          Width = 74
          Height = 13
          Caption = '&Etape du circuit'
          FocusControl = CircuitDest
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLotDest: THLabel
          Left = 11
          Top = 40
          Width = 51
          Height = 13
          Caption = '&Nom du lot'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object CircuitDest: THValComboBox
          Left = 91
          Top = 12
          Width = 137
          Height = 21
          Style = csDropDownList
          Ctl3D = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 0
          TagDispatch = 0
          Plus = 'AND CO_LIBELLE<>""'
          Vide = True
          DataType = 'CPCIRCUITDECAIS'
        end
        object cBAP: TCheckBox
          Left = 11
          Top = 63
          Width = 79
          Height = 14
          Caption = '&Bon '#224' payer'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
        end
        object cModeleBAP: THValComboBox
          Left = 91
          Top = 60
          Width = 137
          Height = 21
          Style = csDropDownList
          ItemHeight = 0
          TabOrder = 2
          TagDispatch = 0
          DataType = 'TTMODELEBAP'
        end
        object LotDest: THCritMaskEdit
          Left = 91
          Top = 36
          Width = 137
          Height = 21
          Ctl3D = True
          ParentCtl3D = False
          TabOrder = 3
          TagDispatch = 0
          DataType = 'CPNOMLOT'
          Operateur = Superieur
          ElipsisButton = True
        end
      end
    end
    object PEcritures: TTabSheet [2]
      Caption = 'Ecritures'
      ImageIndex = 5
      object Bevel6: TBevel
        Left = 0
        Top = 0
        Width = 642
        Height = 89
        Align = alClient
      end
      object TE_JOURNAL: THLabel
        Left = 10
        Top = 38
        Width = 34
        Height = 13
        Caption = '&Journal'
        FocusControl = E_JOURNAL
      end
      object TE_NATUREPIECE: THLabel
        Left = 10
        Top = 66
        Width = 32
        Height = 13
        Caption = '&Nature'
        FocusControl = E_NATUREPIECE
      end
      object TE_DATECOMPTABLE: THLabel
        Left = 255
        Top = 10
        Width = 103
        Height = 13
        Caption = '&Dates comptables  du'
        FocusControl = E_DATECOMPTABLE
      end
      object TE_DATECOMPTABLE2: THLabel
        Left = 468
        Top = 10
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATECOMPTABLE_
      end
      object TE_EXERCICE: THLabel
        Left = 10
        Top = 10
        Width = 41
        Height = 13
        Caption = 'E&xercice'
        FocusControl = E_EXERCICE
      end
      object TE_DATEECHEANCE: THLabel
        Left = 255
        Top = 38
        Width = 72
        Height = 13
        Caption = '&Ech'#233'ances du '
        FocusControl = E_DATEECHEANCE
      end
      object TE_DATEECHEANCE2: THLabel
        Left = 468
        Top = 38
        Width = 12
        Height = 13
        Caption = 'au'
        FocusControl = E_DATEECHEANCE_
      end
      object Label14: TLabel
        Left = 255
        Top = 66
        Width = 73
        Height = 13
        Caption = '&Mode paiement'
        FocusControl = E_MODEPAIE
      end
      object E_EXERCICE: THValComboBox
        Left = 58
        Top = 6
        Width = 181
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 0
        OnChange = E_EXERCICEChange
        TagDispatch = 0
        Vide = True
        DataType = 'TTEXERCICE'
      end
      object E_JOURNAL: THValComboBox
        Left = 58
        Top = 34
        Width = 180
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 1
        TagDispatch = 0
        Vide = True
        DataType = 'TTJOURNAUX'
      end
      object E_NATUREPIECE: THValComboBox
        Left = 58
        Top = 62
        Width = 180
        Height = 21
        Style = csDropDownList
        Ctl3D = True
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        Vide = True
        DataType = 'TTNATUREPIECE'
      end
      object E_DATECOMPTABLE: THCritMaskEdit
        Left = 385
        Top = 6
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 3
        Text = '  /  /    '
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        DefaultDate = od1900
        ElipsisButton = True
        ControlerDate = True
      end
      object E_DATECOMPTABLE_: THCritMaskEdit
        Left = 489
        Top = 6
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 4
        Text = '  /  /    '
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        DefaultDate = od2099
        ElipsisButton = True
        ControlerDate = True
      end
      object E_NUMECHE: THCritMaskEdit
        Left = 120
        Top = 9
        Width = 12
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
        Text = '1'
        Visible = False
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object XX_WHEREAN: TEdit
        Left = 101
        Top = 9
        Width = 13
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
        TabOrder = 11
        Text = 'E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H"'
        Visible = False
      end
      object E_QUALIFPIECE: THCritMaskEdit
        Left = 152
        Top = 9
        Width = 12
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
        TabOrder = 7
        Text = 'N'
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object E_ECHE: THCritMaskEdit
        Left = 170
        Top = 9
        Width = 12
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
        TabOrder = 8
        Text = 'X'
        Visible = False
        TagDispatch = 0
        Operateur = Egal
      end
      object E_TRESOLETTRE: THCritMaskEdit
        Left = 204
        Top = 9
        Width = 12
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
        TabOrder = 9
        Text = 'X'
        Visible = False
        TagDispatch = 0
        Operateur = Different
      end
      object XX_WHEREVIDE: TEdit
        Left = 102
        Top = 36
        Width = 23
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
        TabOrder = 10
        Text = 'E_JOURNAL="aaa"'
        Visible = False
      end
      object E_DATEECHEANCE: THCritMaskEdit
        Left = 385
        Top = 34
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 5
        Text = '  /  /    '
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otDate
        DefaultDate = od1900
        ElipsisButton = True
        ControlerDate = True
      end
      object E_DATEECHEANCE_: THCritMaskEdit
        Left = 489
        Top = 34
        Width = 77
        Height = 21
        Ctl3D = True
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        ParentCtl3D = False
        TabOrder = 6
        Text = '  /  /    '
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otDate
        DefaultDate = od2099
        ElipsisButton = True
        ControlerDate = True
      end
      object XX_WHEREMP: TEdit
        Left = 137
        Top = 9
        Width = 13
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
        TabOrder = 13
        Visible = False
      end
      object E_MODEPAIE: THValComboBox
        Tag = -9980
        Left = 385
        Top = 62
        Width = 182
        Height = 21
        Style = csDropDownList
        Ctl3D = False
        ItemHeight = 0
        ParentCtl3D = False
        TabOrder = 14
        TagDispatch = 0
        Vide = True
        DataType = 'TTMODEPAIE'
      end
      object XX_WHERELETTRAGE: TEdit
        Left = 134
        Top = 36
        Width = 23
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
        TabOrder = 15
        Visible = False
      end
      object XX_WHERENATCLI: TEdit
        Left = 162
        Top = 36
        Width = 23
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
        TabOrder = 16
        Text = 'T_NATUREAUXI="CLI" OR T_NATUREAUXI="AUD"'
        Visible = False
      end
    end
    inherited PComplement: THTabSheet
      inherited Bevel2: TBevel
        Width = 642
        Height = 89
      end
      object TE_NUMEROPIECE: THLabel
        Left = 6
        Top = 38
        Width = 52
        Height = 13
        Caption = '&Num'#233'ro de'
        FocusControl = E_NUMEROPIECE
      end
      object HLabel2: THLabel
        Left = 179
        Top = 36
        Width = 6
        Height = 13
        Caption = #224
      end
      object TE_DEBIT: TLabel
        Left = 285
        Top = 12
        Width = 80
        Height = 13
        Caption = 'Montant &d'#233'bit de'
        FocusControl = E_DEBIT
      end
      object TE_DEBIT_: TLabel
        Left = 475
        Top = 12
        Width = 6
        Height = 13
        Caption = #224
      end
      object TE_CREDIT: TLabel
        Left = 285
        Top = 38
        Width = 83
        Height = 13
        Caption = 'Montant &cr'#233'dit de'
        FocusControl = E_CREDIT
      end
      object TE_CREDIT_: TLabel
        Left = 475
        Top = 38
        Width = 6
        Height = 13
        Caption = #224
      end
      object TE_ETABLISSEMENT: THLabel
        Left = 6
        Top = 12
        Width = 65
        Height = 13
        Caption = '&Etablissement'
        FocusControl = E_ETABLISSEMENT
      end
      object E_NUMEROPIECE: THCritMaskEdit
        Left = 91
        Top = 34
        Width = 77
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 1
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_NUMEROPIECE_: THCritMaskEdit
        Left = 195
        Top = 34
        Width = 77
        Height = 21
        Ctl3D = True
        ParentCtl3D = False
        TabOrder = 2
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_DEBIT: THCritMaskEdit
        Left = 377
        Top = 8
        Width = 90
        Height = 21
        TabOrder = 3
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_DEBIT_: THCritMaskEdit
        Left = 485
        Top = 34
        Width = 90
        Height = 21
        TabOrder = 4
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_CREDIT: THCritMaskEdit
        Left = 377
        Top = 34
        Width = 90
        Height = 21
        TabOrder = 5
        TagDispatch = 0
        Operateur = Superieur
        OpeType = otReel
      end
      object E_CREDIT_: THCritMaskEdit
        Left = 485
        Top = 8
        Width = 90
        Height = 21
        TabOrder = 6
        TagDispatch = 0
        Operateur = Inferieur
        OpeType = otReel
      end
      object E_ETABLISSEMENT: THValComboBox
        Left = 91
        Top = 8
        Width = 181
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
    object Pzlibre: TTabSheet [4]
      Caption = 'Tables libres'
      object Bevel5: TBevel
        Left = 0
        Top = 0
        Width = 642
        Height = 89
        Align = alClient
      end
      object TT_TABLE0: THLabel
        Left = 7
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE0'
      end
      object TT_TABLE1: THLabel
        Left = 122
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE1'
      end
      object TT_TABLE2: THLabel
        Left = 237
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE2'
      end
      object TT_TABLE3: THLabel
        Left = 351
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE3'
      end
      object TT_TABLE4: THLabel
        Left = 468
        Top = 2
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE4'
      end
      object TT_TABLE5: THLabel
        Left = 7
        Top = 44
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE5'
      end
      object TT_TABLE6: THLabel
        Left = 122
        Top = 44
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE6'
      end
      object TT_TABLE7: THLabel
        Left = 237
        Top = 44
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE7'
      end
      object TT_TABLE8: THLabel
        Left = 351
        Top = 44
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE8'
      end
      object TT_TABLE9: THLabel
        Left = 468
        Top = 44
        Width = 107
        Height = 13
        AutoSize = False
        Caption = 'TT_TABLE9'
      end
      object T_TABLE4: THCpteEdit
        Left = 468
        Top = 17
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
        Left = 351
        Top = 17
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
        Left = 237
        Top = 17
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
        Left = 122
        Top = 17
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
        Left = 7
        Top = 17
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
        Left = 7
        Top = 59
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
        Left = 122
        Top = 59
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
        Left = 237
        Top = 59
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
        Left = 351
        Top = 59
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
        Left = 468
        Top = 59
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
        Width = 642
        Height = 89
      end
      inherited bEffaceAvance: TToolbarButton97
        Left = 573
      end
      inherited Z_C1: THValComboBox
        Left = 6
        Width = 151
      end
      inherited Z_C2: THValComboBox
        Left = 6
        Width = 151
      end
      inherited Z_C3: THValComboBox
        Left = 6
        Width = 151
      end
      inherited ZO3: THValComboBox
        Left = 167
        Width = 151
      end
      inherited ZO2: THValComboBox
        Left = 167
        Width = 151
      end
      inherited ZO1: THValComboBox
        Left = 167
        Width = 151
      end
      inherited ZV1: THEdit
        Left = 331
      end
      inherited ZV2: THEdit
        Left = 331
      end
      inherited ZV3: THEdit
        Left = 331
      end
      inherited ZG2: THCombobox
        Left = 531
      end
      inherited ZG1: THCombobox
        Left = 531
      end
    end
    inherited PSQL: THTabSheet
      inherited Bevel3: TBevel
        Width = 580
        Height = 89
      end
      inherited Z_SQL: THSQLMemo
        Width = 580
        Height = 89
      end
    end
  end
  inherited Dock971: TDock97
    Top = 117
    Width = 650
    inherited PFiltres: TToolWindow97
      ClientWidth = 650
      ClientAreaWidth = 650
      inherited BCherche: TToolbarButton97
        Left = 440
      end
      inherited lpresentation: THLabel
        Left = 471
      end
      inherited FFiltres: THValComboBox
        Width = 365
      end
      inherited cbPresentations: THValComboBox
        Left = 542
      end
    end
  end
  inherited FListe: THDBGrid
    Top = 158
    Width = 621
    Height = 143
    MultiSelection = True
    MultiFieds = 
      'E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_NUMEROPIECE;E_QUALIFPIECE' +
      ';E_NUMLIGNE;E_NUMECHE;'
    SortEnabled = True
  end
  inherited Panel2: THPanel
    Left = 425
    inherited PListe: THPanel
      Left = 404
    end
  end
  inherited Dock: TDock97
    Top = 323
    Width = 650
    inherited PanelBouton: TToolWindow97
      ClientWidth = 650
      ClientAreaWidth = 650
      inherited bSelectAll: TToolbarButton97
        Visible = True
      end
      inherited BImprimer: TToolbarButton97
        Left = 523
      end
      inherited BOuvrir: TToolbarButton97
        Left = 555
        Hint = 'Lancer le traitement'
        Caption = 'Traitement'
      end
      inherited BAnnuler: TToolbarButton97
        Left = 587
      end
      inherited BAide: TToolbarButton97
        Left = 619
      end
      inherited Binsert: TToolbarButton97
        Left = 312
      end
      inherited BBlocNote: TToolbarButton97
        Left = 231
        Visible = True
      end
      object BCtrlRib: TToolbarButton97
        Tag = 1
        Left = 198
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Contr'#244'ler les RIB de la liste'
        DisplayMode = dmGlyphOnly
        Caption = 'G'#233'n'#233'ration'
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Layout = blGlyphTop
        Margin = 0
        Opaque = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BCtrlRibClick
        GlobalIndexImage = 'Z0040_S16G1'
        IsControl = True
      end
    end
  end
  inherited PCumul: THPanel
    Top = 301
    Width = 650
  end
  inherited PanVBar: THPanel
    Left = 621
    Top = 158
    Height = 143
  end
  inherited SQ: TDataSource
    Left = 104
  end
  inherited FindDialog: THFindDialog
    Left = 148
  end
  inherited POPF: THPopupMenu
    Left = 236
  end
  inherited HMTrad: THSystemMenu
    Left = 192
  end
  inherited SD: THSaveDialog
    Left = 280
  end
  object HM: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Op'#233'ration '#224' effectuer'
      
        'ATTENTION : Certaines '#233'ch'#233'ances en cours de traitement n'#39'ont pas' +
        ' '#233't'#233' mises '#224' jour !'
      '2;?caption?;Confirmez-vous le traitement ?;Q;YN;Y;Y;')
    Left = 60
    Top = 244
  end
end

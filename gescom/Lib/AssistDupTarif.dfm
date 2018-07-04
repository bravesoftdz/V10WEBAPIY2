inherited FAssistDupTarif: TFAssistDupTarif
  Left = 286
  Top = 207
  Caption = 'Assistant de duplication des tarifs'
  ClientHeight = 392
  ClientWidth = 532
  KeyPreview = True
  OnKeyDown = FormKeyDown
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Top = 369
  end
  inherited lAide: THLabel
    Top = 339
    Width = 344
  end
  inherited bPrecedent: TToolbarButton97
    Left = 265
    Top = 363
  end
  inherited bSuivant: TToolbarButton97
    Left = 348
    Top = 363
  end
  inherited bFin: TToolbarButton97
    Left = 432
    Top = 363
  end
  inherited bAnnuler: TToolbarButton97
    Left = 182
    Top = 363
  end
  inherited bAide: TToolbarButton97
    Left = 99
    Top = 363
  end
  inherited GroupBox1: THGroupBox
    Top = 351
    Width = 535
  end
  inherited P: THPageControl2
    Left = 193
    Top = 12
    Width = 334
    Height = 305
    ActivePage = TabSheet6
    Enabled = False
    object TabSheet1: TTabSheet
      Caption = 'TabSheet1'
      object TINTRO: THLabel
        Left = 11
        Top = 52
        Width = 312
        Height = 29
        AutoSize = False
        Caption = 
          'Cet assistant vous guide afin de param'#232'trer la duplication de vo' +
          's tarifs pr'#233'alablement s'#233'lectionn'#233's.'
        WordWrap = True
      end
      object PTITRE: THPanel
        Left = 0
        Top = 0
        Width = 326
        Height = 41
        Align = alTop
        Caption = 'Duplication des tarifs'
        FullRepaint = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -19
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
        BackGroundEffect = bdFlat
        ColorShadow = clWindowText
        ColorStart = clBtnFace
        TextEffect = tenone
      end
      object GroupBox2: TGroupBox
        Left = 4
        Top = 91
        Width = 321
        Height = 148
        Caption = ' Nouvelles valeurs '#224' appliquer'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object TGF_ARTICLE: THLabel
          Left = 34
          Top = 25
          Width = 56
          Height = 13
          Caption = 'Code article'
          Color = clBtnFace
          FocusControl = GF_ARTICLE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object TGF_TIERS: THLabel
          Left = 34
          Top = 49
          Width = 47
          Height = 13
          Caption = 'Code tiers'
          Color = clBtnFace
          FocusControl = GF_TIERS
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object TGF_DEPOT: THLabel
          Left = 34
          Top = 73
          Width = 55
          Height = 13
          Caption = 'Code d'#233'p'#244't'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object TGF_TARIFARTICLE: THLabel
          Left = 34
          Top = 97
          Width = 76
          Height = 13
          Caption = 'Code tarif article'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object TGF_TARIFTIERS: THLabel
          Left = 34
          Top = 121
          Width = 67
          Height = 13
          Caption = 'Code tarif tiers'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
        end
        object GF_ARTICLE: THCritMaskEdit
          Left = 158
          Top = 21
          Width = 121
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          ReadOnly = True
          TabOrder = 0
          TagDispatch = 0
          ElipsisButton = True
          OnElipsisClick = GF_ARTICLEElipsisClick
        end
        object GF_TIERS: THCritMaskEdit
          Left = 158
          Top = 45
          Width = 121
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          ReadOnly = True
          TabOrder = 1
          TagDispatch = 0
          ElipsisButton = True
          OnElipsisClick = GF_TIERSElipsisClick
        end
        object GF_DEPOT: THValComboBox
          Left = 158
          Top = 69
          Width = 122
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 2
          TagDispatch = 0
          DataType = 'GCDEPOT'
        end
        object GF_TARIFARTICLE: THValComboBox
          Left = 158
          Top = 93
          Width = 122
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 3
          TagDispatch = 0
          DataType = 'GCTARIFARTICLE'
        end
        object GF_TARIFTIERS: THValComboBox
          Left = 158
          Top = 118
          Width = 122
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 13
          ParentFont = False
          TabOrder = 4
          TagDispatch = 0
          DataType = 'TTTARIFCLIENT'
        end
      end
      object CBBasculeTTC: TCheckBox
        Left = 36
        Top = 244
        Width = 105
        Height = 17
        Caption = 'Basculer en TTC'
        TabOrder = 2
      end
      object CBBasculeHT: TCheckBox
        Left = 160
        Top = 244
        Width = 101
        Height = 17
        Caption = 'Basculer en HT'
        TabOrder = 3
      end
    end
    object TabSheet3: TTabSheet
      Caption = 'TabSheet3'
      ImageIndex = 3
      object GroupBox3: TGroupBox
        Left = 4
        Top = 102
        Width = 321
        Height = 122
        Caption = 'Type de mise '#224' jour'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object TCGPRIX: THLabel
          Left = 171
          Top = 15
          Width = 145
          Height = 38
          AutoSize = False
          Caption = 
            'Permet de modifier le prix net ou la remise en pr'#233'cisant une dat' +
            'e d'#39'effet.'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object TCBDATE: THLabel
          Left = 171
          Top = 75
          Width = 145
          Height = 39
          AutoSize = False
          Caption = 'Permet de modifier la p'#233'riode de validit'#233' des tarifs.'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object CB_PRIX: TCheckBox
          Left = 12
          Top = 24
          Width = 145
          Height = 17
          Caption = '&Modification du prix'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 0
        end
        object CB_DATE: TCheckBox
          Left = 12
          Top = 80
          Width = 145
          Height = 17
          Caption = 'Modification de la &p'#233'riode'
          Color = clBtnFace
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentColor = False
          ParentFont = False
          TabOrder = 1
        end
      end
    end
    object TabSheet4: TTabSheet
      Caption = 'TabSheet4'
      ImageIndex = 4
      object PBevel2: TBevel
        Left = 0
        Top = 0
        Width = 326
        Height = 277
        Align = alClient
      end
      object TDATEEFFET: THLabel
        Left = 37
        Top = 224
        Width = 55
        Height = 13
        Caption = '&Date d'#39'effet'
        Color = clYellow
        FocusControl = DATEEFFET
        ParentColor = False
        Visible = False
      end
      object GBPRIX: TGroupBox
        Left = 16
        Top = 5
        Width = 305
        Height = 110
        Caption = 'Modification des prix'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object TTYPEMAJPRIX: THLabel
          Left = 20
          Top = 22
          Width = 92
          Height = 13
          Caption = '&Type de mise '#224' jour'
          FocusControl = TYPEMAJPRIX
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TVALEURPRIX: THLabel
          Left = 20
          Top = 50
          Width = 30
          Height = 13
          Caption = '&Valeur'
          FocusControl = VALEURPRIX
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TARRONDIPRIX: THLabel
          Left = 20
          Top = 84
          Width = 88
          Height = 13
          Caption = '&M'#233'thode d"arrondi'
          FocusControl = ARRONDIPRIX
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TYPEMAJPRIX: THValComboBox
          Left = 124
          Top = 18
          Width = 145
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 0
          OnChange = TYPEMAJPRIXChange
          TagDispatch = 0
          Plus = 'AND CO_CODE Like "P%"'
          Vide = True
          VideString = '<<Aucun>>'
          DataType = 'GCMAJTARIF'
        end
        object VALEURPRIX: THNumEdit
          Left = 124
          Top = 50
          Width = 145
          Height = 21
          TabStop = False
          Color = clBtnFace
          Decimals = 3
          Digits = 12
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Masks.PositiveMask = '#,##0.000'
          Debit = False
          ParentFont = False
          TabOrder = 1
          UseRounding = True
          Validate = False
        end
        object ARRONDIPRIX: THValComboBox
          Left = 124
          Top = 80
          Width = 145
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 2
          TagDispatch = 0
          DataType = 'GCCODEARRONDI'
        end
      end
      object GBREMISE: TGroupBox
        Left = 16
        Top = 121
        Width = 305
        Height = 89
        Caption = 'Modification des remises'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentColor = False
        ParentFont = False
        TabOrder = 1
        object TTYPEMAJREMISE: THLabel
          Left = 20
          Top = 25
          Width = 92
          Height = 13
          Caption = 'T&ype de mise '#224' jour'
          FocusControl = TYPEMAJREMISE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TVALEURREMISE: THLabel
          Left = 20
          Top = 59
          Width = 30
          Height = 13
          Caption = 'V&aleur'
          FocusControl = VALEURREMISE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          OnClick = TYPEMAJREMISEChange
        end
        object TYPEMAJREMISE: THValComboBox
          Left = 124
          Top = 22
          Width = 145
          Height = 21
          Style = csDropDownList
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 0
          OnChange = TYPEMAJREMISEChange
          TagDispatch = 0
          Plus = 'AND CO_CODE Like "R%"'
          Vide = True
          VideString = '<<Aucun>>'
          DataType = 'GCMAJTARIF'
        end
        object VALEURREMISE: THNumEdit
          Left = 124
          Top = 55
          Width = 145
          Height = 21
          TabStop = False
          Color = clBtnFace
          Decimals = 2
          Digits = 12
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Masks.PositiveMask = '#,##0.00'
          Debit = False
          ParentFont = False
          TabOrder = 1
          UseRounding = True
          Validate = False
        end
      end
      object DATEEFFET: THCritMaskEdit
        Left = 142
        Top = 220
        Width = 142
        Height = 21
        Color = clYellow
        EditMask = '!99 >L<LL 0000;1;_'
        MaxLength = 11
        TabOrder = 2
        Text = '           '
        Visible = False
        OnExit = DATEEFFETExit
        TagDispatch = 0
        OpeType = otDate
        DefaultDate = odDate
        ControlerDate = True
      end
    end
    object TabSheet5: TTabSheet
      Caption = 'TabSheet5'
      ImageIndex = 5
      object GBPERIODE: TGroupBox
        Left = 8
        Top = 8
        Width = 317
        Height = 185
        Caption = 'Modification de la p'#233'riode'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clActiveCaption
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object TTYPEMAJPERIODE: THLabel
          Left = 20
          Top = 24
          Width = 92
          Height = 13
          Caption = '&Type de mise '#224' jour'
          FocusControl = TYPEMAJPERIODE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TNBREJOUR: THLabel
          Left = 20
          Top = 52
          Width = 72
          Height = 13
          Caption = '&Nombre de jour'
          FocusControl = NBREJOUR
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TDATEDEBUT: THLabel
          Left = 20
          Top = 116
          Width = 67
          Height = 13
          Caption = 'D'#233'&but p'#233'riode'
          FocusControl = DATEDEBUT
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TDATEFIN: THLabel
          Left = 156
          Top = 116
          Width = 52
          Height = 13
          Caption = 'Fi&n p'#233'riode'
          FocusControl = DATEFIN
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TYPEMAJPERIODE: THValComboBox
          Left = 155
          Top = 20
          Width = 145
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ItemHeight = 0
          ParentFont = False
          TabOrder = 0
          OnChange = TYPEMAJPERIODEChange
          TagDispatch = 0
          Plus = 'AND CO_CODE Like "D%"'
          Vide = True
          VideString = '<<Aucun>>'
          DataType = 'GCMAJTARIF'
        end
        object NBREJOUR: THNumEdit
          Left = 155
          Top = 48
          Width = 127
          Height = 21
          TabStop = False
          Color = clBtnFace
          Decimals = 0
          Digits = 4
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          Masks.PositiveMask = '#,##0'
          Max = 365.000000000000000000
          Min = 1.000000000000000000
          Debit = False
          ParentFont = False
          TabOrder = 1
          UseRounding = True
          Value = 1.000000000000000000
          Validate = True
        end
        object UpDownJour: TUpDown
          Left = 282
          Top = 48
          Width = 15
          Height = 21
          Associate = NBREJOUR
          Enabled = False
          Min = 1
          Max = 365
          Position = 1
          TabOrder = 2
        end
        object CBDATEDEBUT: TCheckBox
          Left = 20
          Top = 88
          Width = 121
          Height = 17
          TabStop = False
          Caption = 'Changer date &d'#233'but'
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          OnExit = CBDATEDEBUTClick
        end
        object CBDATEFIN: TCheckBox
          Left = 155
          Top = 88
          Width = 137
          Height = 17
          TabStop = False
          Caption = 'Changer date &fin'
          Enabled = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 4
          OnExit = CBDATEFINClick
        end
        object DATEFIN: THCritMaskEdit
          Left = 155
          Top = 134
          Width = 121
          Height = 21
          TabStop = False
          Color = clBtnFace
          Enabled = False
          EditMask = '!99 >L<LL 0000;1;_'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 11
          ParentFont = False
          TabOrder = 6
          Text = '           '
          OnExit = DATEFINExit
          TagDispatch = 0
          OpeType = otDate
          DefaultDate = od2099
          ControlerDate = True
        end
        object DATEDEBUT: THCritMaskEdit
          Left = 20
          Top = 134
          Width = 121
          Height = 21
          TabStop = False
          Color = clBtnFace
          Enabled = False
          EditMask = '!99 >L<LL 0000;1;_'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 11
          ParentFont = False
          TabOrder = 5
          Text = '           '
          OnExit = DATEDEBUTExit
          TagDispatch = 0
          OpeType = otDate
          DefaultDate = odDate
          ControlerDate = True
        end
      end
    end
    object TabSheet6: TTabSheet
      Caption = 'TabSheet6'
      ImageIndex = 2
      object TRecap: THLabel
        Left = 18
        Top = 108
        Width = 59
        Height = 13
        Caption = 'R'#233'capitulatif'
      end
      object PanelFin: TPanel
        Left = 5
        Top = 4
        Width = 317
        Height = 97
        TabOrder = 0
        object TTextFin1: THLabel
          Left = 23
          Top = 8
          Width = 287
          Height = 26
          Caption = 
            'Le param'#232'trage est maintenant correctement renseign'#233' pour permet' +
            'tre le lancement de la  mise '#224' jour des tarifs.'
          WordWrap = True
        end
        object TTextFin2: THLabel
          Left = 23
          Top = 49
          Width = 285
          Height = 39
          Caption = 
            'Si vous d'#233'sirez revoir le param'#233'trage, il suffit de cliquer sur ' +
            'le bouton Pr'#233'c'#233'dent sinon, le bouton Fin, permet de d'#233'buter le t' +
            'raitement.'
          WordWrap = True
        end
      end
      object ListRecap: TListBox
        Left = 5
        Top = 126
        Width = 317
        Height = 116
        Color = clBtnFace
        ItemHeight = 13
        TabOrder = 1
      end
    end
  end
  inherited PanelImage: THPanel
    Height = 197
  end
  object ART: THCritMaskEdit [12]
    Left = 24
    Top = 268
    Width = 121
    Height = 21
    Color = clYellow
    TabOrder = 5
    Visible = False
    TagDispatch = 0
  end
  object TIERS: THCritMaskEdit [13]
    Left = 24
    Top = 296
    Width = 121
    Height = 21
    Color = clYellow
    TabOrder = 6
    Visible = False
    TagDispatch = 0
  end
  object HMsgErr: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'n'#39'est pas une date valide.')
    Left = 60
    Top = 8
  end
  object HRecap: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Oui'
      'Non'
      'ATTENTION : Tarif non enregistr'#233' !'
      
        'ATTENTION : Ce tarif, en cours de traitement par un autre utilis' +
        'ateur, n'#39'a pas '#233't'#233' enregistr'#233'e !'
      'enregistrement(s) modifi'#233#39's)'
      'Basculement de HT '#224' TTC'
      'Basculement de TTC '#224' HT')
    Left = 132
    Top = 12
  end
end

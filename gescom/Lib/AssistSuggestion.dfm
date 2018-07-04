inherited FAssistReappro: TFAssistReappro
  Left = 847
  Top = 240
  HelpContext = 110000305
  Caption = 'Assistant de suggestion de r'#233'approvisionnement'
  ClientHeight = 340
  ClientWidth = 574
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Top = 317
  end
  inherited lAide: THLabel
    Top = 287
    Width = 386
  end
  inherited bPrecedent: TToolbarButton97
    Left = 307
    Top = 311
  end
  inherited bSuivant: TToolbarButton97
    Left = 390
    Top = 311
  end
  inherited bFin: TToolbarButton97
    Left = 474
    Top = 311
  end
  inherited bAnnuler: TToolbarButton97
    Left = 224
    Top = 311
  end
  inherited bAide: TToolbarButton97
    Left = 141
    Top = 311
    HelpContext = 11000305
  end
  inherited GroupBox1: THGroupBox
    Top = 299
    Width = 577
  end
  inherited P: THPageControl2
    Left = 174
    Width = 391
    Height = 268
    ActivePage = TInitiale
    object TInitiale: TTabSheet
      Caption = 'TInitiale'
      object PENTETE: TPanel
        Left = 6
        Top = 8
        Width = 373
        Height = 33
        Caption = 'Suggestion d'#39'approvisionnement'
        Font.Charset = ANSI_CHARSET
        Font.Color = clNavy
        Font.Height = -16
        Font.Name = 'Arial Black'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
      end
      object RadioGroup1: TRadioGroup
        Left = 6
        Top = 47
        Width = 373
        Height = 186
        Caption = 'Politique d'#39'approvisionnement'
        TabOrder = 1
      end
      object TcouvSimple: TRadioButton
        Left = 12
        Top = 75
        Width = 193
        Height = 17
        Caption = 'Couverture des besoins de livraison'
        Checked = True
        TabOrder = 2
        TabStop = True
        OnClick = TcouvSimpleClick
      end
      object CSelDocs: TCheckBox
        Left = 87
        Top = 99
        Width = 169
        Height = 17
        Caption = 'S'#233'lection des documents'
        TabOrder = 3
      end
      object GestHisto: TCheckBox
        Left = 87
        Top = 180
        Width = 274
        Height = 21
        Caption = 'Prise en compte de l'#39'historique de consommation'
        TabOrder = 4
        OnClick = GestHistoClick
      end
      object TBesoinStock: TRadioButton
        Left = 13
        Top = 160
        Width = 358
        Height = 17
        Caption = 'Calcul des besoins de stocks'
        TabOrder = 5
        OnClick = TBesoinStockClick
      end
      object Panel1: TPanel
        Left = 7
        Top = 139
        Width = 371
        Height = 2
        BevelOuter = bvLowered
        TabOrder = 6
      end
      object CWithStocks: TCheckBox
        Left = 87
        Top = 118
        Width = 178
        Height = 17
        Caption = 'Avec prise en compte des stocks'
        TabOrder = 7
      end
    end
    object TbesoinP0: TTabSheet
      Caption = 'TbesoinP0'
      ImageIndex = 7
      object HLabel8: THLabel
        Left = 43
        Top = 19
        Width = 288
        Height = 26
        Caption = 
          'Veuillez indiquer les '#233'l'#233'ments qui doivent '#234'tre pris en compte p' +
          'our le calcul de la quantit'#233' '#224' r'#233'approvisionner'
        WordWrap = True
      end
      object RGBase: TRadioGroup
        Left = 43
        Top = 152
        Width = 289
        Height = 62
        Caption = '  Calcul bas'#233' sur  '
        ItemIndex = 0
        Items.Strings = (
          'le stock physique'
          'le stock net')
        TabOrder = 0
      end
      object RGValeur: TRadioGroup
        Left = 43
        Top = 72
        Width = 289
        Height = 62
        Caption = '  Quantit'#233' r'#233'approvisionn'#233'e  '
        ItemIndex = 0
        Items.Strings = (
          'au besoin r'#233'el calcul'#233
          'au stock maximum')
        TabOrder = 1
        OnClick = RGValeurClick
      end
    end
    object TbesoinP1: TTabSheet
      Caption = 'TBesoinP1'
      ImageIndex = 2
      object TText4: THLabel
        Left = 36
        Top = 19
        Width = 293
        Height = 26
        Caption = 
          'Veuillez indiquer la p'#233'riode de vente pour laquelle vous voulez ' +
          's'#233'lectionner les articles '#224' traiter'
        WordWrap = True
      end
      object TDATEDEB: THLabel
        Left = 36
        Top = 106
        Width = 82
        Height = 13
        Caption = 'D'#233'but de p'#233'riode'
      end
      object TDATEFIN: THLabel
        Left = 36
        Top = 134
        Width = 67
        Height = 13
        Caption = 'Fin de p'#233'riode'
      end
      object DATEDEB: THCritMaskEdit
        Left = 180
        Top = 104
        Width = 69
        Height = 21
        EditMask = '!99 >L<LL 0000;1;_'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 11
        ParentFont = False
        TabOrder = 0
        Text = '           '
        TagDispatch = 0
        OpeType = otDate
        DefaultDate = odDate
        ControlerDate = True
      end
      object DATEDEB_: THCritMaskEdit
        Left = 180
        Top = 132
        Width = 69
        Height = 21
        EditMask = '!99 >L<LL 0000;1;_'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        MaxLength = 11
        ParentFont = False
        TabOrder = 1
        Text = '           '
        TagDispatch = 0
        OpeType = otDate
        DefaultDate = odDate
        ControlerDate = True
      end
    end
    object TbesoinP2: TTabSheet
      Caption = 'TbesoinP2'
      ImageIndex = 3
      object GB_Histo: TGroupBox
        Left = 7
        Top = 37
        Width = 366
        Height = 152
        Caption = '  Prise en compte de l'#39'historique  '
        TabOrder = 0
        object Label2: TLabel
          Left = 123
          Top = 106
          Width = 64
          Height = 13
          Caption = 'derniers mois '
        end
        object RB_Long: TRadioButton
          Left = 16
          Top = 46
          Width = 273
          Height = 17
          Caption = 'Sur les 2 derni'#232'res ann'#233'es'
          Checked = True
          TabOrder = 0
          TabStop = True
          OnClick = RB_LongClick
        end
        object RB_Court: TRadioButton
          Left = 16
          Top = 104
          Width = 53
          Height = 17
          Caption = 'Sur les'
          TabOrder = 1
          OnClick = RB_CourtClick
        end
        object SE_Histo: TSpinEdit
          Left = 72
          Top = 101
          Width = 45
          Height = 22
          MaxValue = 12
          MinValue = 1
          TabOrder = 2
          Value = 1
        end
      end
    end
    object TbesoinP3: TTabSheet
      Caption = 'TbesoinP3'
      ImageIndex = 3
      object TText3: THLabel
        Left = 21
        Top = 19
        Width = 311
        Height = 26
        Caption = 
          'Vous avez la possibilit'#233' de limiter la s'#233'lection du choix des ar' +
          'ticles en fonction des informations ci dessous. '
        WordWrap = True
      end
      object TGA_FAMILLENIV1: THLabel
        Left = 21
        Top = 69
        Width = 76
        Height = 13
        Caption = 'Famille niveau 1'
      end
      object HLabel1: THLabel
        Left = 169
        Top = 69
        Width = 12
        Height = 13
        Caption = 'de'
      end
      object HLabel4: THLabel
        Left = 169
        Top = 93
        Width = 6
        Height = 13
        Caption = #224
      end
      object TGA_FAMILLENIV2: THLabel
        Left = 21
        Top = 124
        Width = 76
        Height = 13
        Caption = 'Famille niveau 2'
      end
      object HLabel2: THLabel
        Left = 169
        Top = 125
        Width = 12
        Height = 13
        Caption = 'de'
      end
      object HLabel5: THLabel
        Left = 169
        Top = 149
        Width = 6
        Height = 13
        Caption = #224
      end
      object TGA_FAMILLENIV3: THLabel
        Left = 21
        Top = 180
        Width = 76
        Height = 13
        Caption = 'Famille niveau 3'
      end
      object HLabel3: THLabel
        Left = 169
        Top = 181
        Width = 12
        Height = 13
        Caption = 'de'
      end
      object HLabel6: THLabel
        Left = 169
        Top = 205
        Width = 6
        Height = 13
        Caption = #224
      end
      object GA_FAMILLENIV1: THValComboBox
        Left = 193
        Top = 65
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 0
        OnChange = GA_FAMILLENIV1Change
        TagDispatch = 0
        Vide = True
        DataType = 'GCFAMILLENIV1'
      end
      object GA_FAMILLENIV1_: THValComboBox
        Left = 193
        Top = 89
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 1
        TagDispatch = 0
        Vide = True
        DataType = 'GCFAMILLENIV1'
      end
      object GA_FAMILLENIV2: THValComboBox
        Left = 193
        Top = 120
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 2
        OnChange = GA_FAMILLENIV2Change
        TagDispatch = 0
        Vide = True
        DataType = 'GCFAMILLENIV2'
      end
      object GA_FAMILLENIV2_: THValComboBox
        Left = 193
        Top = 144
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 3
        TagDispatch = 0
        Vide = True
        DataType = 'GCFAMILLENIV2'
      end
      object GA_FAMILLENIV3: THValComboBox
        Left = 193
        Top = 176
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 4
        OnChange = GA_FAMILLENIV3Change
        TagDispatch = 0
        Vide = True
        DataType = 'GCFAMILLENIV3'
      end
      object GA_FAMILLENIV3_: THValComboBox
        Left = 193
        Top = 200
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 5
        TagDispatch = 0
        Vide = True
        DataType = 'GCFAMILLENIV3'
      end
    end
    object TbesoinP4: TTabSheet
      Caption = 'TbesoinP4'
      ImageIndex = 4
      object TText2: THLabel
        Left = 35
        Top = 23
        Width = 278
        Height = 39
        Caption = 
          'Vous avez la possibilit'#233' de limiter le calcul en fonction des in' +
          'formations ci dessous. Une zone laiss'#233'e vide signifie que tous l' +
          'es '#233'l'#233'ments correspondants seront trait'#233's.'
        WordWrap = True
      end
      object TSR_ARTICLE: THLabel
        Left = 35
        Top = 95
        Width = 72
        Height = 13
        Caption = 'Du code article'
        FocusControl = GZZ_ARTICLE
      end
      object TSR_ARTICLE_: THLabel
        Left = 35
        Top = 119
        Width = 71
        Height = 13
        Caption = 'Au code article'
        FocusControl = GZZ_ARTICLE_
      end
      object TSR_DEPOT: THLabel
        Left = 35
        Top = 172
        Width = 44
        Height = 13
        Caption = 'Du d'#233'p'#244't'
      end
      object TSR_DEPOT_: THLabel
        Left = 35
        Top = 196
        Width = 43
        Height = 13
        Caption = 'Au d'#233'p'#244't'
      end
      object GZZ_ARTICLE: THCritMaskEdit
        Left = 156
        Top = 91
        Width = 145
        Height = 21
        TabOrder = 0
        TagDispatch = 0
        ElipsisButton = True
        OnElipsisClick = GZZ_ARTICLEElipsisClick
      end
      object GZZ_ARTICLE_: THCritMaskEdit
        Left = 156
        Top = 115
        Width = 145
        Height = 21
        TabOrder = 1
        TagDispatch = 0
        ElipsisButton = True
        OnElipsisClick = GZZ_ARTICLE_ElipsisClick
      end
      object GZZ_DEPOT: THValComboBox
        Left = 156
        Top = 168
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 2
        OnChange = GZZ_DEPOTChange
        TagDispatch = 0
        Vide = True
        VideString = '<<Tous>>'
        DataType = 'GCDEPOT'
      end
      object GZZ_DEPOT_: THValComboBox
        Left = 156
        Top = 192
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 3
        TagDispatch = 0
        Vide = True
        VideString = '<<Tous>>'
        DataType = 'GCDEPOT'
      end
      object CBFermes: TCheckBox
        Left = 64
        Top = 144
        Width = 105
        Height = 17
        Alignment = taLeftJustify
        AllowGrayed = True
        Caption = 'Articles ferm'#233's'
        State = cbGrayed
        TabOrder = 4
      end
    end
    object TbesoinP5: TTabSheet
      Caption = 'TbesoinP5'
      ImageIndex = 5
      object TText6: THLabel
        Left = 28
        Top = 20
        Width = 292
        Height = 39
        Caption = 
          'Vous devez pr'#233'ciser quel est l'#39'ordre de priorit'#233' dans le choix d' +
          'es fournisseurs. Pr'#233'cisez de pr'#233'f'#233'rence 2 niveaux de priorit'#233' af' +
          'in de couvrir un choix plus vaste.'
        WordWrap = True
      end
      object TPRIORITE1: THLabel
        Left = 28
        Top = 116
        Width = 92
        Height = 13
        Caption = 'Niveau de priorit'#233' 1'
      end
      object TPRIORITE2: THLabel
        Left = 28
        Top = 144
        Width = 92
        Height = 13
        Caption = 'Niveau de priorit'#233' 2'
      end
      object TPRIORITE3: THLabel
        Left = 28
        Top = 172
        Width = 92
        Height = 13
        Caption = 'Niveau de priorit'#233' 3'
      end
      object TPRIORITE4: THLabel
        Left = 28
        Top = 201
        Width = 92
        Height = 13
        Caption = 'Niveau de priorit'#233' 4'
      end
      object PRIORITE: THValComboBox
        Left = 180
        Top = 88
        Width = 145
        Height = 21
        Style = csDropDownList
        Color = clYellow
        ItemHeight = 0
        TabOrder = 0
        Visible = False
        TagDispatch = 0
      end
      object PRIORITE1: THValComboBox
        Left = 172
        Top = 112
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 1
        OnChange = PRIORITE1Change
        TagDispatch = 0
      end
      object PRIORITE2: THValComboBox
        Left = 172
        Top = 140
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 2
        OnChange = PRIORITE2Change
        TagDispatch = 0
      end
      object PRIORITE3: THValComboBox
        Left = 172
        Top = 168
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 3
        OnChange = PRIORITE3Change
        TagDispatch = 0
      end
      object PRIORITE4: THValComboBox
        Left = 172
        Top = 197
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 4
        OnChange = PRIORITE4Change
        TagDispatch = 0
      end
    end
    object TRecap: TTabSheet
      Caption = 'TRecap'
      ImageIndex = 6
      object HLabel7: THLabel
        Left = 14
        Top = 89
        Width = 59
        Height = 13
        Caption = 'R'#233'capitulatif'
      end
      object PanelFin: TPanel
        Left = 5
        Top = 4
        Width = 376
        Height = 84
        TabOrder = 0
        object TTextFin1: THLabel
          Left = 17
          Top = 10
          Width = 342
          Height = 26
          Caption = 
            'Le param'#232'trage est maintenant correctement renseign'#233' pour permet' +
            'tre la g'#233'n'#233'ration de la suggestion d'#39'approvisionnement.'
          WordWrap = True
        end
        object TTextFin2: THLabel
          Left = 17
          Top = 44
          Width = 324
          Height = 26
          Caption = 
            'Si vous d'#233'sirez revoir le param'#233'trage, il suffit de cliquer sur ' +
            'le bouton Pr'#233'c'#233'dent sinon, le bouton Fin, permet de d'#233'buter le t' +
            'raitement.'
          WordWrap = True
        end
      end
      object Recap: THRichEditOLE
        Left = 4
        Top = 103
        Width = 377
        Height = 133
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 1
        Margins.Top = 0
        Margins.Bottom = 0
        Margins.Left = 0
        Margins.Right = 0
        ContainerName = 'Document'
        ObjectMenuPrefix = '&Object'
        LinesRTF.Strings = (
          '{\rtf1\ansi\deff0\nouicompat{\fonttbl{\f0\fnil Arial;}}'
          '{\*\generator Riched20 10.0.15063}\viewkind4\uc1 '
          '\pard\f0\fs16\lang1036 '
          '\par '
          '\par '
          '\par '
          '\par '
          '\par '
          '\par }')
      end
    end
  end
  inherited Msg: THMsgBox
    Mess.Strings = (
      'Etape'
      '1;Assistant;Voulez-vous quitter l'#39'assistant ?;Q;YN;Y;C;'
      
        '2;?caption?;Il existe d'#233'j'#224' une suggestion avec cette r'#233'f'#233'rence;W' +
        ';O;O;O;')
  end
end

inherited FAssistCreationAffaire: TFAssistCreationAffaire
  Left = 122
  Top = 162
  HelpContext = 120000503
  Caption = 'Assistant de cr'#233'ation d'#39'affaire'
  ClientWidth = 564
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Caption = 'Etape 1/3'
  end
  object HLabel1: THLabel [2]
    Left = 8
    Top = 52
    Width = 238
    Height = 13
    Caption = 'S'#233'lection de l'#39'affaire mod'#232'le de r'#233'f'#233'rence'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object HLabel2: THLabel [3]
    Left = 10
    Top = 76
    Width = 70
    Height = 13
    Caption = 'Affaire mod'#232'le '
  end
  object HLabel4: THLabel [4]
    Left = 16
    Top = 158
    Width = 168
    Height = 13
    Caption = 'Informations compl'#233'mentaires'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object HLabel5: THLabel [5]
    Left = 29
    Top = 184
    Width = 104
    Height = 13
    Caption = '&Responsable principal'
  end
  inherited bFin: TToolbarButton97
    Caption = '&Cr'#233'ation'
  end
  inherited bAide: TToolbarButton97
    HelpContext = 120000503
  end
  inherited P: THPageControl2
    Left = 174
    Width = 387
    Height = 276
    ActivePage = PAvancee
    object TSCODIFICATION: TTabSheet
      Caption = 'S'#233'lection de l'#39'affaire mod'#232'le '
      object Bevel3: TBevel
        Left = 1
        Top = 32
        Width = 378
        Height = 208
      end
      object BevelDecoupe: TBevel
        Left = 1
        Top = 0
        Width = 378
        Height = 31
      end
      object LibSelectModele: THLabel
        Left = 20
        Top = 40
        Width = 238
        Height = 13
        Caption = 'S'#233'lection de l'#39'affaire mod'#232'le de r'#233'f'#233'rence'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object libselectModele2: THLabel
        Left = 20
        Top = 81
        Width = 93
        Height = 13
        Caption = '&Affaire de r'#233'f'#233'rence'
        FocusControl = AFF_AFFAIRE1
      end
      object LibWelcome1: THLabel
        Left = 12
        Top = 7
        Width = 337
        Height = 16
        Caption = 'Bienvenue dans l'#39'assistant de cr'#233'ation d'#39'affaires'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object LibSelectClient: THLabel
        Left = 20
        Top = 138
        Width = 247
        Height = 13
        Caption = 'S'#233'lection du ou des client(s) de destination'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object LibSelectClient2: THLabel
        Left = 20
        Top = 166
        Width = 29
        Height = 13
        Caption = '&Client '
      end
      object NBMULTICLIENT: THLabel
        Left = 20
        Top = 219
        Width = 116
        Height = 13
        Caption = 'Aucun client s'#233'lectionn'#233
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object BRechAffaire: TToolbarButton97
        Left = 124
        Top = 75
        Width = 23
        Height = 22
        Hint = 'Recherche Affaire'
        Color = clNone
        Opaque = False
        ShowBorderWhenInactive = True
        OnClick = BRechAffaireClick
        GlobalIndexImage = 'Z0002_S16G1'
      end
      object LibModeleTiers: THLabel
        Left = 20
        Top = 105
        Width = 89
        Height = 13
        Caption = 'C&lient de r'#233'f'#233'rence'
        FocusControl = AFF_AFFAIRE1
      end
      object AFFAIREMODELE: THCritMaskEdit
        Left = 300
        Top = 74
        Width = 49
        Height = 21
        Color = clYellow
        TabOrder = 5
        Visible = False
        TagDispatch = 0
        DataType = 'AFLAFFAIREMODELE'
      end
      object AFF_TIERS: THCritMaskEdit
        Left = 152
        Top = 162
        Width = 143
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 6
        TagDispatch = 0
        DataType = 'GCTIERSCLI'
        ElipsisButton = True
        OnElipsisClick = AFF_TIERSElipsisClick
      end
      object CBMultiClient: TCheckBox
        Left = 20
        Top = 196
        Width = 121
        Height = 17
        Caption = 'S'#233'lection multi-clients'
        TabOrder = 7
        OnClick = CBMultiClientClick
      end
      object BtSelectMultiClient: THBitBtn
        Left = 154
        Top = 192
        Width = 139
        Height = 25
        Caption = 'S'#233'lection multi-clients'
        TabOrder = 8
        OnClick = BtSelectMultiClientClick
      end
      object AFF_AFFAIRE0: THCritMaskEdit
        Left = 292
        Top = 75
        Width = 31
        Height = 21
        AutoSize = False
        CharCase = ecUpperCase
        TabOrder = 0
        Visible = False
        TagDispatch = 0
      end
      object AFF_AFFAIRE1: THCritMaskEdit
        Left = 151
        Top = 75
        Width = 42
        Height = 21
        AutoSize = False
        CharCase = ecUpperCase
        TabOrder = 1
        OnExit = AFF_AFFAIRE1Exit
        TagDispatch = 0
      end
      object AFF_AFFAIRE2: THCritMaskEdit
        Left = 193
        Top = 75
        Width = 54
        Height = 21
        AutoSize = False
        CharCase = ecUpperCase
        TabOrder = 2
        OnExit = AFF_AFFAIRE1Exit
        TagDispatch = 0
      end
      object AFF_AFFAIRE3: THCritMaskEdit
        Left = 247
        Top = 75
        Width = 20
        Height = 21
        AutoSize = False
        CharCase = ecUpperCase
        TabOrder = 3
        OnExit = AFF_AFFAIRE1Exit
        TagDispatch = 0
      end
      object AFF_AVENANT: THCritMaskEdit
        Left = 267
        Top = 75
        Width = 22
        Height = 21
        AutoSize = False
        CharCase = ecUpperCase
        TabOrder = 4
        Visible = False
        TagDispatch = 0
      end
      object AFF_TIERSMODELE: THCritMaskEdit
        Left = 152
        Top = 101
        Width = 143
        Height = 21
        TabOrder = 9
        TagDispatch = 0
        DataType = 'GCTIERSCLI'
        ElipsisButton = True
        OnElipsisClick = AFF_TIERSMODELEElipsisClick
      end
      object SurAffModele: TCheckBox
        Left = 20
        Top = 57
        Width = 261
        Height = 17
        Caption = 'Cr'#233'ation '#224' partir d'#39'une affaire mod'#232'le'
        Checked = True
        State = cbChecked
        TabOrder = 10
      end
    end
    object TSCaracteristique: TTabSheet
      Caption = 'Caract'#233'ristiques'
      ImageIndex = 1
      object LibSelectDates: THLabel
        Left = 28
        Top = 34
        Width = 177
        Height = 13
        Caption = 'S'#233'lection des dates de l'#39'affaire'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object TAFF_DATEDEBUT: THLabel
        Left = 41
        Top = 56
        Width = 14
        Height = 13
        Caption = 'D&u'
        FocusControl = AFF_DATEDEBUT
      end
      object TAFF_DATEFIN: THLabel
        Left = 161
        Top = 56
        Width = 15
        Height = 13
        Caption = '&au '
        FocusControl = AFF_DATEFIN
      end
      object HLabel3: THLabel
        Left = 54
        Top = 7
        Width = 264
        Height = 16
        Caption = 'Nouvelles caract'#233'ristiques de l'#39'affaire'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Bevel1: TBevel
        Left = 1
        Top = 0
        Width = 378
        Height = 31
      end
      object Bevel2: TBevel
        Left = 1
        Top = 31
        Width = 378
        Height = 217
      end
      object LibModeFact: THLabel
        Left = 28
        Top = 142
        Width = 119
        Height = 13
        Caption = 'Mode de facturation '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object LibLModeFact2: THLabel
        Left = 41
        Top = 163
        Width = 95
        Height = 13
        Caption = 'Type de g'#233'n'#233'ration '
      end
      object TAFF_MONTANTECHE: THLabel
        Left = 41
        Top = 200
        Width = 88
        Height = 13
        Caption = '&Montant forfaitaire '
        FocusControl = AFF_MONTANTECHE
      end
      object LibLModeFact3: THLabel
        Left = 41
        Top = 182
        Width = 52
        Height = 13
        Caption = 'P'#233'riodicit'#233' '
      end
      object AFF_GENERAUTO: TLabel
        Left = 172
        Top = 164
        Width = 43
        Height = 13
        Caption = 'Manuelle'
      end
      object AFF_PERIODICITE: TLabel
        Left = 172
        Top = 180
        Width = 48
        Height = 13
        Caption = 'Mensuelle'
      end
      object TAFF_DATEDEBGENER: THLabel
        Left = 41
        Top = 225
        Width = 68
        Height = 13
        Caption = 'Facturation du'
        FocusControl = AFF_DATEDEBGENER
      end
      object TAFF_DATEFINGENER: THLabel
        Left = 253
        Top = 225
        Width = 15
        Height = 13
        Caption = '&au '
        FocusControl = AFF_DATEFINGENER
      end
      object AFF_DATEDEBUT: THCritMaskEdit
        Left = 67
        Top = 52
        Width = 70
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 0
        Text = '  /  /    '
        TagDispatch = 0
        OpeType = otDate
        DefaultDate = od1900
        ControlerDate = True
      end
      object AFF_DATEFIN: THCritMaskEdit
        Left = 195
        Top = 52
        Width = 70
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 1
        Text = '  /  /    '
        TagDispatch = 0
        OpeType = otDate
        DefaultDate = od2099
        ControlerDate = True
      end
      object AFF_MONTANTECHE: THCritMaskEdit
        Left = 172
        Top = 196
        Width = 73
        Height = 21
        TabOrder = 2
        TagDispatch = 0
        OpeType = otReel
      end
      object CalculMoisCloture: TCheckBox
        Left = 29
        Top = 78
        Width = 341
        Height = 14
        Caption = 
          '&Calcul automatique des dates de la mission sur le mois Cl'#244'ture ' +
          'client'
        TabOrder = 3
        OnClick = CalculMoisClotureClick
      end
      object CalculExercice: TCheckBox
        Left = 29
        Top = 97
        Width = 341
        Height = 17
        Caption = '&Recalcul de l'#39'exercice sur la base du mois de cl'#244'ture'
        TabOrder = 4
        OnClick = CalculMoisClotureClick
      end
      object AFF_DATEDEBGENER: THCritMaskEdit
        Left = 172
        Top = 221
        Width = 73
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 5
        Text = '  /  /    '
        TagDispatch = 0
        OpeType = otDate
        DefaultDate = od1900
        ControlerDate = True
      end
      object AFF_DATEFINGENER: THCritMaskEdit
        Left = 277
        Top = 221
        Width = 70
        Height = 21
        EditMask = '!99/99/0000;1;_'
        MaxLength = 10
        TabOrder = 6
        Text = '  /  /    '
        TagDispatch = 0
        OpeType = otDate
        DefaultDate = od2099
        ControlerDate = True
      end
      object bDateGener: TCheckBox
        Left = 29
        Top = 78
        Width = 270
        Height = 17
        Caption = '&Reprise des dates de facturation de l'#39'affaire mod'#232'le'
        Checked = True
        State = cbChecked
        TabOrder = 7
        OnClick = bDateGenerClick
      end
      object RepFactAff: TCheckBox
        Left = 29
        Top = 118
        Width = 341
        Height = 17
        Caption = 'Reprise int'#233'grale des '#233'ch'#233'ances de facturation (sinon recalcul)'
        TabOrder = 8
        OnClick = RepFactAffClick
      end
    end
    object PAvancee: TTabSheet
      Caption = 'Param'#232'tres avanc'#233's'
      ImageIndex = 2
      object Bevel4: TBevel
        Left = 1
        Top = 0
        Width = 378
        Height = 31
      end
      object Bevel5: TBevel
        Left = 1
        Top = 31
        Width = 378
        Height = 217
        Hint = 'object HLabel6: THLabelobject HLabel6: THLabel'
      end
      object LibRepriseclientdestination: THLabel
        Left = 24
        Top = 60
        Width = 263
        Height = 13
        Caption = 'Reprises des donn'#233'es du client de destination'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object LibParamavance: THLabel
        Left = 54
        Top = 7
        Width = 144
        Height = 16
        Caption = 'Param'#233'tres avanc'#233's'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object LibInfoComplement: THLabel
        Left = 24
        Top = 125
        Width = 232
        Height = 13
        Caption = 'Informations compl'#233'mentaires de l'#39'affaire'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object TAFF_RESPONSABLE: THLabel
        Left = 18
        Top = 149
        Width = 104
        Height = 13
        Caption = '&Responsable principal'
        FocusControl = AFF_RESPONSABLE
      end
      object TAFF_DEPARTEMENT: THLabel
        Left = 18
        Top = 172
        Width = 61
        Height = 13
        Caption = '&D'#233'partement'
        FocusControl = AFF_DEPARTEMENT
      end
      object LA_Taches: THLabel
        Left = 24
        Top = 201
        Width = 113
        Height = 13
        Caption = 'T'#226'ches du planning'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Escompteclient: TCheckBox
        Left = 18
        Top = 77
        Width = 341
        Height = 17
        Caption = 'Reprise de l'#39'escompte du client de destination'
        TabOrder = 0
        OnClick = CalculMoisClotureClick
      end
      object ModeRegleClient: TCheckBox
        Left = 18
        Top = 95
        Width = 341
        Height = 17
        Caption = 'Reprise du mode de r'#232'glement du client de destination'
        Checked = True
        State = cbChecked
        TabOrder = 1
        OnClick = CalculMoisClotureClick
      end
      object AFF_RESPONSABLE: THCritMaskEdit
        Left = 144
        Top = 145
        Width = 145
        Height = 21
        CharCase = ecUpperCase
        TabOrder = 2
        TagDispatch = 0
        DataType = 'AFLRESSOURCE'
        ElipsisButton = True
        OnElipsisClick = AFF_RESPONSABLEElipsisClick
      end
      object AFF_DEPARTEMENT: THValComboBox
        Left = 144
        Top = 168
        Width = 145
        Height = 21
        ItemHeight = 13
        TabOrder = 3
        TagDispatch = 0
        DataType = 'AFDEPARTEMENT'
      end
      object CB_DupliquerTache: TCheckBox
        Left = 18
        Top = 221
        Width = 341
        Height = 17
        Caption = 'Dupliquer les t'#226'ches de l'#39'affaire mod'#232'le'
        TabOrder = 4
        OnClick = CalculMoisClotureClick
      end
      object RepEtatAff: TCheckBox
        Left = 18
        Top = 37
        Width = 341
        Height = 17
        Caption = 'Reprise zone "'#233'tat affaire" de l'#39'affaire d'#39'origine (sinon ENC)'
        TabOrder = 5
        OnClick = CalculMoisClotureClick
      end
    end
  end
end

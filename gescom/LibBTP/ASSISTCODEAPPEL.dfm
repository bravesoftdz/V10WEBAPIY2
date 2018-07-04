inherited FAssistCodeAppel: TFAssistCodeAppel
  Left = 274
  Top = 205
  Caption = ''
  ClientHeight = 346
  ClientWidth = 584
  PixelsPerInch = 96
  TextHeight = 13
  inherited lEtape: THLabel
    Top = 316
  end
  inherited lAide: THLabel
    Left = 242
    Top = 236
  end
  object LIBLNG: THLabel [2]
    Left = 175
    Top = 283
    Width = 322
    Height = 13
    Caption = 'Longueur totale  : 14 caract'#232'res - Longueur disponible : '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object LNGDISPO: THLabel [3]
    Left = 490
    Top = 283
    Width = 15
    Height = 13
    Caption = '14'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object LIBLNG2: THLabel [4]
    Left = 508
    Top = 283
    Width = 61
    Height = 13
    Caption = 'caract'#232'res'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  inherited bPrecedent: TToolbarButton97
    Top = 310
  end
  inherited bSuivant: TToolbarButton97
    Left = 352
    Top = 310
  end
  inherited bFin: TToolbarButton97
    Top = 310
    Caption = 'V&alidation'
  end
  inherited bAnnuler: TToolbarButton97
    Top = 310
  end
  inherited bAide: TToolbarButton97
    Top = 310
    HelpContext = 120000504
  end
  inherited GroupBox1: THGroupBox
    Top = 298
  end
  inherited P: THPageControl2
    Left = 175
    Top = 3
    Width = 406
    Height = 278
    ActivePage = PARTIE2
    Enabled = False
    object TabSheet1: TTabSheet
      Caption = 'D'#233'coupage du code Appel'
      object LIB1: THLabel
        Left = 18
        Top = 4
        Width = 362
        Height = 16
        Caption = 'Bienvenue dans l'#39'assistant de param'#233'trage de votre'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object HLabel2: THLabel
        Left = 18
        Top = 24
        Width = 151
        Height = 16
        Caption = 'codification d'#39'Appels '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object libgene1: THLabel
        Left = 8
        Top = 70
        Width = 325
        Height = 13
        Caption = 
          'Vous pouvez g'#233'rer de 1 '#224' 3 parties distinctes dans votre codific' +
          'ation.'
      end
      object BevelDecoupe: TBevel
        Left = 5
        Top = 0
        Width = 379
        Height = 46
      end
      object LIBNBPART: THLabel
        Left = 8
        Top = 52
        Width = 146
        Height = 13
        Caption = 'Nombre de parties g'#233'r'#233'es'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object LibGene2: THLabel
        Left = 8
        Top = 87
        Width = 350
        Height = 13
        Caption = 
          'Chaque partie est une source d'#39'analyse distincte. Les 4 types de' +
          ' donn'#233'es '
      end
      object LibGene3: THLabel
        Left = 8
        Top = 103
        Width = 323
        Height = 13
        Caption = 
          'possibles pour chaque partie sont : un compteur num'#233'rique, une l' +
          'iste'
      end
      object LibGene4: THLabel
        Left = 8
        Top = 120
        Width = 236
        Height = 13
        Caption = 'param'#233'trable, une zone de saisie , une valeur fixe.'
      end
      object TSO_APPCODENBPARTIE: THLabel
        Left = 8
        Top = 142
        Width = 140
        Height = 13
        Caption = '&Nombre de parties souhait'#233'es'
        FocusControl = SO_APPCODENBPARTIE
      end
      object SO_APPCODENBPARTIE: TSpinEdit
        Left = 160
        Top = 137
        Width = 49
        Height = 22
        MaxValue = 3
        MinValue = 1
        TabOrder = 0
        Value = 1
        OnChange = SO_APPCODENBPARTIEChange
      end
    end
    object TSCompl: TTabSheet
      Caption = 'Compl'#233'ment'
      ImageIndex = 4
      object HLabel3: THLabel
        Left = 73
        Top = 11
        Width = 246
        Height = 16
        Caption = 'Acceptation et duplication d'#39'Appels'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Bevel1: TBevel
        Left = 5
        Top = 4
        Width = 379
        Height = 31
      end
      object LibGene5: THLabel
        Left = 8
        Top = 45
        Width = 274
        Height = 13
        Caption = 'Compteur distinct propositions d'#39'Appels / appels'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object LibGene6: THLabel
        Left = 8
        Top = 63
        Width = 365
        Height = 13
        Caption = 
          'Si vous utiliser un compteur automatique cette option vous perme' +
          't de codifier '
      end
      object HLabel1: THLabel
        Left = 8
        Top = 78
        Width = 370
        Height = 13
        Caption = 
          'sur la m'#234'me base ou sur deux bases diff'#233'rentes  les Appels et le' +
          's  propositions.'
      end
      object HLabel4: THLabel
        Left = 8
        Top = 133
        Width = 343
        Height = 13
        Caption = 'Recodification du code Appel en acceptation et duplication '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object HLabel5: THLabel
        Left = 8
        Top = 151
        Width = 347
        Height = 13
        Caption = 
          'Lors de l'#39'acceptation d'#39'une proposition ou lors d'#39'une duplicatio' +
          'n , pr'#233'ciser '
      end
      object HLabel6: THLabel
        Left = 8
        Top = 167
        Width = 242
        Height = 13
        Caption = 'si vous souhaitez modifier les parties du code Appel'
      end
      object SO_APPPRODIFFERENT: TCheckBox
        Left = 50
        Top = 98
        Width = 260
        Height = 17
        Caption = 'Compteur distinct propositions / Appels'
        TabOrder = 0
        OnClick = SO_APPPRODIFFERENTClick
      end
      object SO_APPRECODIFICATION: TCheckBox
        Left = 50
        Top = 190
        Width = 252
        Height = 17
        Caption = 'Recodification possible du code Appel'
        TabOrder = 1
        OnClick = SO_APPPRODIFFERENTClick
      end
    end
    object PARTIE1: TTabSheet
      Tag = 1
      Caption = 'Partie 1'
      ImageIndex = 1
      object TSO_APPCO1LIB: THLabel
        Left = 16
        Top = 48
        Width = 69
        Height = 13
        Caption = '&Libell'#233' associ'#233
        FocusControl = SO_APPCO1LIB
      end
      object TSO_APPCO1TYPE: THLabel
        Left = 16
        Top = 81
        Width = 83
        Height = 13
        Caption = '&Type de donn'#233'es'
        FocusControl = SO_APPCO1TYPE
      end
      object TSO_APPCO1LNG: THLabel
        Left = 16
        Top = 114
        Width = 90
        Height = 13
        Caption = 'L&ongueur du code '
        FocusControl = SO_APPCO1LNG
      end
      object Entetepartie1: THLabel
        Left = 109
        Top = 7
        Width = 159
        Height = 16
        Caption = 'Partie 1 du code Appel'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object BevelPartie1: TBevel
        Left = 5
        Top = 2
        Width = 379
        Height = 27
      end
      object TSO_APPCO1VALEUR: THLabel
        Left = 16
        Top = 147
        Width = 30
        Height = 13
        Caption = '&Valeur'
        FocusControl = SO_APPCO1VALEUR
      end
      object TSO_APPCO1VALEURPRO: THLabel
        Left = 16
        Top = 170
        Width = 124
        Height = 13
        Caption = '&Compteur  sur propositions'
        FocusControl = SO_APPCO1VALEURPRO
      end
      object TSO_BTNATAFFAIRE: THLabel
        Left = 2
        Top = 234
        Width = 70
        Height = 13
        Caption = 'Pi'#232'ces Appels '
        Visible = False
      end
      object TSO_BTNATPROPOSITION: THLabel
        Left = 192
        Top = 234
        Width = 94
        Height = 13
        Caption = 'Pi'#232'ces propositions '
        Visible = False
      end
      object SO_APPCO1LIB: THCritMaskEdit
        Left = 147
        Top = 44
        Width = 145
        Height = 21
        TabOrder = 0
        TagDispatch = 0
      end
      object SO_APPCO1TYPE: THValComboBox
        Tag = 1
        Left = 147
        Top = 77
        Width = 145
        Height = 21
        ItemHeight = 13
        TabOrder = 1
        OnClick = SO_APPCO1TYPEClick
        TagDispatch = 0
        DataType = 'BTTYPECODEAPPEL'
      end
      object SO_APPCO1LNG: TSpinEdit
        Tag = 1
        Left = 147
        Top = 105
        Width = 49
        Height = 22
        Enabled = False
        MaxValue = 14
        MinValue = 1
        TabOrder = 2
        Value = 1
        OnChange = SO_APPCO1LNGChange
      end
      object SO_APPCO1VALEUR: THCritMaskEdit
        Left = 147
        Top = 141
        Width = 145
        Height = 21
        MaxLength = 14
        TabOrder = 3
        OnExit = SO_APPCOVALEURExit
        TagDispatch = 0
      end
      object PARAMAFFAIREPART1: TBitBtn
        Tag = 1
        Left = 16
        Top = 197
        Width = 156
        Height = 25
        Caption = 'Param'#233'trage liste des codes '
        TabOrder = 4
        OnClick = PARAMAFFAIREPART1Click
      end
      object SO_APPCO1VALEURPRO: THCritMaskEdit
        Left = 147
        Top = 166
        Width = 145
        Height = 21
        MaxLength = 14
        TabOrder = 5
        OnExit = SO_APPCOVALEURExit
        TagDispatch = 0
      end
      object SO_BTNATAFFAIRE: THValComboBox
        Tag = 1
        Left = 76
        Top = 229
        Width = 111
        Height = 21
        ItemHeight = 13
        TabOrder = 6
        Visible = False
        OnClick = SO_APPCO1TYPEClick
        TagDispatch = 0
        Plus = ' OR GPP_VENTEACHAT="AUT"'
        DataType = 'GCNATUREPIECEGVEN'
      end
      object SO_BTNATPROPOSITION: THValComboBox
        Tag = 1
        Left = 287
        Top = 229
        Width = 111
        Height = 21
        ItemHeight = 13
        TabOrder = 7
        Visible = False
        OnClick = SO_APPCO1TYPEClick
        TagDispatch = 0
        Plus = ' OR GPP_VENTEACHAT="AUT"'
        DataType = 'GCNATUREPIECEGVEN'
      end
      object SO_AAAAMM1: TRadioButton
        Left = 213
        Top = 108
        Width = 77
        Height = 17
        Caption = 'AAAA/MM'
        TabOrder = 8
        Visible = False
        OnClick = SO_AAAAMM1Click
      end
      object SO_AAMM1: TRadioButton
        Left = 147
        Top = 108
        Width = 65
        Height = 17
        Caption = 'AA/MM'
        TabOrder = 9
        Visible = False
        OnClick = SO_AAMM1Click
      end
    end
    object PARTIE2: TTabSheet
      Tag = 2
      Caption = 'Partie 2'
      ImageIndex = 2
      object TSO_APPCO2LIB: THLabel
        Left = 16
        Top = 48
        Width = 69
        Height = 13
        Caption = '&Libell'#233' associ'#233
        FocusControl = SO_APPCO2LIB
      end
      object TSO_APPCO2TYPE: THLabel
        Left = 16
        Top = 81
        Width = 83
        Height = 13
        Caption = '&Type de donn'#233'es'
        FocusControl = SO_APPCO2TYPE
      end
      object TSO_APPCO2LNG: THLabel
        Left = 16
        Top = 114
        Width = 90
        Height = 13
        Caption = 'L&ongueur du code '
        FocusControl = SO_APPCO2LNG
      end
      object LIBPARTIE2: THLabel
        Left = 109
        Top = 7
        Width = 159
        Height = 16
        Caption = 'Partie 2 du code Appel'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object BevelPartie2: TBevel
        Left = 5
        Top = 2
        Width = 379
        Height = 27
      end
      object TSO_APPCO2VALEUR: THLabel
        Left = 16
        Top = 147
        Width = 30
        Height = 13
        Caption = '&Valeur'
        FocusControl = SO_APPCO2VALEUR
      end
      object TSO_APPCO2VALEURPRO: THLabel
        Left = 16
        Top = 170
        Width = 124
        Height = 13
        Caption = '&Compteur  sur propositions'
        FocusControl = SO_APPCO2VALEURPRO
      end
      object SO_APPCO2LIB: THCritMaskEdit
        Left = 147
        Top = 44
        Width = 145
        Height = 21
        TabOrder = 0
        TagDispatch = 0
      end
      object SO_APPCO2TYPE: THValComboBox
        Tag = 2
        Left = 147
        Top = 77
        Width = 145
        Height = 21
        ItemHeight = 13
        TabOrder = 1
        OnClick = SO_APPCO1TYPEClick
        TagDispatch = 0
        DataType = 'BTTYPECODEAPPEL'
      end
      object SO_APPCO2LNG: TSpinEdit
        Left = 147
        Top = 105
        Width = 49
        Height = 22
        MaxValue = 14
        MinValue = 1
        TabOrder = 2
        Value = 1
      end
      object SO_APPCO2VALEUR: THCritMaskEdit
        Left = 147
        Top = 141
        Width = 145
        Height = 21
        MaxLength = 14
        TabOrder = 3
        OnExit = SO_APPCOVALEURExit
        TagDispatch = 0
      end
      object SO_APPCO2VISIBLE: TCheckBox
        Left = 217
        Top = 208
        Width = 168
        Height = 17
        Caption = 'Partie 2 visible pour l'#39'utilisateur'
        TabOrder = 4
        Visible = False
      end
      object PARAMAFFAIREPART2: TBitBtn
        Tag = 2
        Left = 16
        Top = 197
        Width = 156
        Height = 25
        Caption = 'Param'#233'trage liste des codes '
        TabOrder = 5
        OnClick = PARAMAFFAIREPART1Click
      end
      object SO_APPCO2VALEURPRO: THCritMaskEdit
        Left = 147
        Top = 166
        Width = 145
        Height = 21
        MaxLength = 14
        TabOrder = 6
        OnExit = SO_APPCOVALEURExit
        TagDispatch = 0
      end
      object SO_AAAAMM2: TRadioButton
        Left = 213
        Top = 108
        Width = 77
        Height = 17
        Caption = 'AAAA/MM'
        TabOrder = 7
        Visible = False
        OnClick = SO_AAAAMM1Click
      end
      object SO_AAMM2: TRadioButton
        Left = 147
        Top = 108
        Width = 65
        Height = 17
        Caption = 'AA/MM'
        TabOrder = 8
        Visible = False
        OnClick = SO_AAMM1Click
      end
    end
    object PARTIE3: TTabSheet
      Tag = 3
      Caption = 'Partie 3'
      ImageIndex = 3
      object TSO_APPCO3LIB: THLabel
        Left = 16
        Top = 48
        Width = 69
        Height = 13
        Caption = '&Libell'#233' associ'#233
        FocusControl = SO_APPCO3LIB
      end
      object TSO_APPCO3TYPE: THLabel
        Left = 16
        Top = 81
        Width = 83
        Height = 13
        Caption = '&Type de donn'#233'es'
        FocusControl = SO_APPCO3TYPE
      end
      object TSO_APPCO3LNG: THLabel
        Left = 16
        Top = 114
        Width = 90
        Height = 13
        Caption = '&Longueur du code '
        FocusControl = SO_APPCO3LNG
      end
      object LibPartie3: THLabel
        Left = 109
        Top = 7
        Width = 159
        Height = 16
        Caption = 'Partie 3 du code Appel'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object BevelPartie3: TBevel
        Left = 5
        Top = 3
        Width = 379
        Height = 27
      end
      object TSO_APPCO3VALEUR: THLabel
        Left = 16
        Top = 147
        Width = 30
        Height = 13
        Caption = '&Valeur'
        FocusControl = SO_APPCO3VALEUR
      end
      object TSO_APPCO3VALEURPRO: THLabel
        Left = 16
        Top = 170
        Width = 124
        Height = 13
        Caption = '&Compteur  sur propositions'
        FocusControl = SO_APPCO3VALEURPRO
      end
      object SO_APPCO3LIB: THCritMaskEdit
        Left = 147
        Top = 44
        Width = 145
        Height = 21
        TabOrder = 0
        TagDispatch = 0
      end
      object SO_APPCO3TYPE: THValComboBox
        Tag = 3
        Left = 147
        Top = 77
        Width = 145
        Height = 21
        ItemHeight = 13
        TabOrder = 1
        OnClick = SO_APPCO1TYPEClick
        TagDispatch = 0
        DataType = 'BTTYPECODEAPPEL'
      end
      object SO_APPCO3LNG: TSpinEdit
        Tag = 3
        Left = 147
        Top = 105
        Width = 49
        Height = 22
        Enabled = False
        MaxValue = 14
        MinValue = 1
        TabOrder = 2
        Value = 1
      end
      object SO_APPCO3VALEUR: THCritMaskEdit
        Left = 147
        Top = 141
        Width = 145
        Height = 21
        MaxLength = 14
        TabOrder = 3
        OnExit = SO_APPCOVALEURExit
        TagDispatch = 0
      end
      object SO_APPCO3VISIBLE: TCheckBox
        Left = 217
        Top = 208
        Width = 168
        Height = 17
        Caption = 'Partie 3 visible pour l'#39'utilisateur'
        TabOrder = 4
        Visible = False
      end
      object PARAMAFFAIREPART3: TBitBtn
        Tag = 3
        Left = 16
        Top = 197
        Width = 156
        Height = 25
        Caption = 'Param'#233'trage liste des codes '
        TabOrder = 5
        OnClick = PARAMAFFAIREPART1Click
      end
      object SO_APPCO3VALEURPRO: THCritMaskEdit
        Left = 147
        Top = 166
        Width = 145
        Height = 21
        MaxLength = 14
        TabOrder = 6
        OnExit = SO_APPCOVALEURExit
        TagDispatch = 0
      end
      object SO_AAMM3: TRadioButton
        Left = 147
        Top = 108
        Width = 65
        Height = 17
        Caption = 'AA/MM'
        TabOrder = 7
        Visible = False
        OnClick = SO_AAMM1Click
      end
      object SO_AAAAMM3: TRadioButton
        Left = 213
        Top = 108
        Width = 77
        Height = 17
        Caption = 'AAAA/MM'
        TabOrder = 8
        Visible = False
        OnClick = SO_AAAAMM1Click
      end
    end
  end
  inherited PanelImage: THPanel
    inherited Image: TToolbarButton97
      Left = -4
      Top = 38
    end
  end
  inherited cControls: THListBox
    Top = 269
  end
end

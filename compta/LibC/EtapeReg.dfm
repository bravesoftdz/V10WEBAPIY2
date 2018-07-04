inherited FEtapeReg: TFEtapeReg
  Left = 410
  Top = 168
  Caption = 'Etapes de R'#232'glement :'
  ClientHeight = 370
  ClientWidth = 549
  PixelsPerInch = 96
  TextHeight = 13
  inherited FListe: THDBGrid
    Left = 364
    Width = 185
    Height = 328
    Columns = <
      item
        Expanded = False
        Title.Alignment = taCenter
        Title.Caption = 'Code'
        Width = 43
        Visible = True
      end
      item
        Expanded = False
        Title.Alignment = taCenter
        Title.Caption = 'Libell'#233
        Width = 107
        Visible = True
      end>
  end
  inherited Pappli: THPanel
    Width = 364
    Height = 328
    object HLabel3: THLabel
      Left = 8
      Top = 18
      Width = 25
      Height = 13
      Caption = '&Code'
      FocusControl = ER_ETAPE
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HLabel5: THLabel
      Left = 8
      Top = 44
      Width = 30
      Height = 13
      Caption = '&Libell'#233
      FocusControl = ER_LIBELLE
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HLabel8: THLabel
      Left = 168
      Top = 17
      Width = 28
      Height = 13
      Caption = '&Etape'
      FocusControl = ER_TYPEETAPE
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object ER_ETAPE: TDBEdit
      Left = 89
      Top = 14
      Width = 57
      Height = 21
      CharCase = ecUpperCase
      DataField = 'ER_ETAPE'
      DataSource = STa
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
    end
    object ENCDEC: TDBCheckBox
      Left = 48
      Top = 17
      Width = 29
      Height = 17
      TabStop = False
      Alignment = taLeftJustify
      Caption = 'ENCDEC'
      Color = clYellow
      DataField = 'ER_ENCAISSEMENT'
      DataSource = STa
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      TabOrder = 5
      ValueChecked = 'X'
      ValueUnchecked = '-'
      Visible = False
    end
    object ER_LIBELLE: TDBEdit
      Left = 89
      Top = 40
      Width = 265
      Height = 21
      DataField = 'ER_LIBELLE'
      DataSource = STa
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      OnExit = ER_LIBELLEExit
    end
    object ER_TYPEETAPE: THDBValComboBox
      Left = 208
      Top = 14
      Width = 146
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 1
      OnExit = ER_TYPEETAPEExit
      TagDispatch = 0
      DataType = 'TTTYPEETAPEREGLE'
      DataField = 'ER_TYPEETAPE'
      DataSource = STa
    end
    object ER_GLOBALISE: TDBCheckBox
      Left = 5
      Top = 68
      Width = 126
      Height = 17
      Alignment = taLeftJustify
      Caption = 'Contre&partie globalis'#233'e'
      DataField = 'ER_GLOBALISE'
      DataSource = STa
      TabOrder = 3
      ValueChecked = 'X'
      ValueUnchecked = '-'
    end
    object GComptes: TGroupBox
      Left = 2
      Top = 88
      Width = 358
      Height = 71
      Caption = ' Comptes '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 4
      object HLabel4: THLabel
        Left = 8
        Top = 17
        Width = 69
        Height = 13
        Caption = 'Suivi t&r'#233'sorerie'
        FocusControl = ER_CPTEDEPART
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel9: THLabel
        Left = 8
        Top = 44
        Width = 52
        Height = 13
        Caption = '&G'#233'n'#233'ration'
        FocusControl = ER_CPTEARRIVEE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object EDEPART: THLabel
        Left = 188
        Top = 17
        Width = 164
        Height = 13
        AutoSize = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object EARRIVEE: THLabel
        Left = 188
        Top = 44
        Width = 164
        Height = 13
        AutoSize = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object ER_CPTEARRIVEE: THDBCpteEdit
        Left = 88
        Top = 41
        Width = 93
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnExit = ER_CPTEARRIVEEExit
        ZoomTable = tzGEncais
        Vide = False
        Bourre = True
        Libelle = EARRIVEE
        okLocate = False
        SynJoker = False
        DataField = 'ER_CPTEARRIVEE'
        DataSource = STa
      end
      object ER_CPTEDEPART: THDBCpteEdit
        Left = 88
        Top = 14
        Width = 93
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        OnExit = ER_CPTEDEPARTExit
        ZoomTable = tzGEncais
        Vide = False
        Bourre = True
        Libelle = EDEPART
        okLocate = False
        SynJoker = False
        DataField = 'ER_CPTEDEPART'
        DataSource = STa
      end
    end
    object Pages: TPageControl
      Left = 0
      Top = 165
      Width = 364
      Height = 163
      ActivePage = TS2
      Align = alBottom
      TabOrder = 6
      object TS1: TTabSheet
        Caption = 'Param'#232'tres'
        object HLabel1: THLabel
          Left = 8
          Top = 115
          Width = 28
          Height = 13
          Caption = '&Guide'
          FocusControl = LIBGUIDE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BZOOMETAPE: TToolbarButton97
          Left = 325
          Top = 111
          Width = 22
          Height = 21
          Hint = 'Guide de g'#233'n'#233'ration'
          Layout = blGlyphRight
          ParentShowHint = False
          ShowHint = True
          OnClick = BZOOMETAPEClick
          GlobalIndexImage = 'Z0008_S16G1'
        end
        object HLabel6: THLabel
          Left = 8
          Top = 89
          Width = 33
          Height = 13
          Caption = '&Devise'
          FocusControl = ER_DEVISE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TER_CODEAFB: THLabel
          Left = 8
          Top = 62
          Width = 68
          Height = 13
          Caption = 'ou Codes &AFB'
          FocusControl = ER_CODEAFB
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel2: THLabel
          Left = 9
          Top = 35
          Width = 73
          Height = 13
          Caption = '&Mode paiement'
          FocusControl = ER_MODEPAIE
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object HLabel7: THLabel
          Left = 8
          Top = 9
          Width = 77
          Height = 13
          Caption = 'Cat'#233'g. pa&iement'
          FocusControl = ER_CATEGORIEMP
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object LIBGUIDE: TEdit
          Left = 104
          Top = 111
          Width = 217
          Height = 21
          TabOrder = 5
          OnExit = LIBGUIDEExit
          OnKeyDown = LIBGUIDEKeyDown
        end
        object ER_DEVISE: THDBValComboBox
          Left = 104
          Top = 85
          Width = 217
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 3
          TagDispatch = 0
          DataType = 'TTDEVISE'
          DataField = 'ER_DEVISE'
          DataSource = STa
        end
        object ER_GUIDE: TDBEdit
          Left = 316
          Top = 85
          Width = 37
          Height = 21
          Color = clYellow
          Ctl3D = True
          DataField = 'ER_GUIDE'
          DataSource = STa
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentCtl3D = False
          ParentFont = False
          TabOrder = 4
          Visible = False
        end
        object ER_CODEAFB: THDBMultiValComboBox
          Left = 104
          Top = 58
          Width = 217
          Height = 21
          TabOrder = 2
          Text = 'ER_CODEAFB'
          Abrege = False
          DataType = 'TTAFB'
          Complete = False
          OuInclusif = False
          DataField = 'ER_CODEAFB'
          DataSource = STa
        end
        object ER_MODEPAIE: THDBValComboBox
          Left = 104
          Top = 31
          Width = 217
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 1
          OnChange = ER_MODEPAIEChange
          TagDispatch = 0
          DataField = 'ER_MODEPAIE'
          DataSource = STa
        end
        object ER_CATEGORIEMP: THDBValComboBox
          Left = 104
          Top = 5
          Width = 217
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 0
          OnExit = ER_CATEGORIEMPExit
          TagDispatch = 0
          DataType = 'TTMODEPAIECAT'
          DataField = 'ER_CATEGORIEMP'
          DataSource = STa
        end
      end
      object TS2: TTabSheet
        Caption = 'Emissions'
        ImageIndex = 1
        object HLabel10: THLabel
          Left = 158
          Top = 8
          Width = 32
          Height = 13
          Caption = '&Format'
          FocusControl = ER_FORMATCFONB
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TER_DOCUMENT: THLabel
          Left = 158
          Top = 38
          Width = 35
          Height = 13
          Caption = '&Mod'#232'le'
          FocusControl = ER_DOCUMENT
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object TER_ENVOITRANS: THLabel
          Left = 5
          Top = 70
          Width = 151
          Height = 13
          Caption = '&T'#233'l'#233'transmission vers la banque'
          FocusControl = ER_ENVOITRANS
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object ER_EXPORTCFONB: TDBCheckBox
          Left = 5
          Top = 7
          Width = 126
          Height = 17
          Alignment = taLeftJustify
          Caption = 'E&xport CFONB'
          DataField = 'ER_EXPORTCFONB'
          DataSource = STa
          TabOrder = 0
          ValueChecked = 'X'
          ValueUnchecked = '-'
          OnClick = ER_EXPORTCFONBClick
        end
        object ER_FORMATCFONB: THDBValComboBox
          Left = 208
          Top = 5
          Width = 146
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 1
          OnExit = ER_TYPEETAPEExit
          TagDispatch = 0
          DataType = 'TTTYPEEXPORTCFONB'
          DataField = 'ER_FORMATCFONB'
          DataSource = STa
        end
        object ER_BORDEREAU: TDBCheckBox
          Left = 5
          Top = 37
          Width = 126
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Emission &bordereau'
          DataField = 'ER_BORDEREAU'
          DataSource = STa
          TabOrder = 2
          ValueChecked = 'X'
          ValueUnchecked = '-'
          OnClick = ER_BORDEREAUClick
        end
        object ER_DOCUMENT: THDBValComboBox
          Left = 208
          Top = 35
          Width = 146
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 3
          OnExit = ER_TYPEETAPEExit
          TagDispatch = 0
          DataType = 'TTMODELEBOR'
          DataField = 'ER_DOCUMENT'
          DataSource = STa
        end
        object ER_ENVOITRANS: THDBValComboBox
          Left = 208
          Top = 66
          Width = 146
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 4
          OnExit = ER_TYPEETAPEExit
          TagDispatch = 0
          DataType = 'TTENVOITRANS'
          DataField = 'ER_ENVOITRANS'
          DataSource = STa
        end
      end
    end
  end
  inherited DBNav: TDBNavigator
    Left = 211
    Top = 67
    Hints.Strings = ()
  end
  inherited Dock971: TDock97
    Top = 328
    Width = 549
    inherited PBouton: TToolWindow97
      ClientWidth = 549
      ClientAreaWidth = 549
      inherited BImprimer: TToolbarButton97
        Left = 232
      end
      inherited BValider: TToolbarButton97
        Left = 264
      end
      inherited HelpBtn: TToolbarButton97
        Left = 328
      end
      inherited BFerme: TToolbarButton97
        Left = 296
      end
      inherited BFirst: TToolbarButton97
        Left = 364
      end
      inherited BPrev: TToolbarButton97
        Left = 408
      end
      inherited BNext: TToolbarButton97
        Left = 452
      end
      inherited BLast: TToolbarButton97
        Left = 496
      end
    end
  end
  inherited Ta: THTable
    TableName = 'ETAPEREG'
    Left = 392
    object TaER_ENCAISSEMENT: TStringField
      FieldName = 'ER_ENCAISSEMENT'
      Size = 1
    end
    object TaER_ETAPE: TStringField
      FieldName = 'ER_ETAPE'
      Size = 3
    end
    object TaER_LIBELLE: TStringField
      FieldName = 'ER_LIBELLE'
      Size = 70
    end
    object TaER_TYPEETAPE: TStringField
      FieldName = 'ER_TYPEETAPE'
      Size = 3
    end
    object TaER_CPTEDEPART: TStringField
      FieldName = 'ER_CPTEDEPART'
      Size = 17
    end
    object TaER_CPTEARRIVEE: TStringField
      FieldName = 'ER_CPTEARRIVEE'
      Size = 17
    end
    object TaER_CATEGORIEMP: TStringField
      FieldName = 'ER_CATEGORIEMP'
      Size = 3
    end
    object TaER_MODEPAIE: TStringField
      FieldName = 'ER_MODEPAIE'
      Size = 3
    end
    object TaER_DEVISE: TStringField
      FieldName = 'ER_DEVISE'
      Size = 3
    end
    object TaER_GUIDE: TStringField
      FieldName = 'ER_GUIDE'
      Size = 3
    end
    object TaER_GLOBALISE: TStringField
      FieldName = 'ER_GLOBALISE'
      Size = 1
    end
    object TaER_EXPORTCFONB: TStringField
      FieldName = 'ER_EXPORTCFONB'
      Size = 1
    end
    object TaER_FORMATCFONB: TStringField
      FieldName = 'ER_FORMATCFONB'
      Size = 3
    end
    object TaER_BORDEREAU: TStringField
      FieldName = 'ER_BORDEREAU'
      Size = 1
    end
    object TaER_DOCUMENT: TStringField
      FieldName = 'ER_DOCUMENT'
      Size = 3
    end
    object TaER_CODEAFB_SQL: TMemoField
      FieldName = 'ER_CODEAFB'
      BlobType = ftMemo
      Size = 2000
    end
    object TaER_ENVOITRANS: TStringField
      FieldName = 'ER_ENVOITRANS'
      Size = 3
    end
    object TaER_CODEAFB_ORA: TStringField
      FieldName = 'ER_CODEAFB2'
      Size = 2000
    end
  end
  inherited STa: TDataSource
    Left = 444
    Top = 176
  end
  inherited HM: THMsgBox
    Mess.Strings = (
      
        '0;Etapes de r'#232'glement;Voulez-vous enregistrer les modifications ' +
        '?;Q;YNC;Y;C;'
      
        '1;Etapes de r'#232'glement;Confirmez-vous la suppression de l'#39'enregis' +
        'trement ?;Q;YNC;N;C;'
      '2;Etapes de r'#232'glement;Vous devez renseigner un code.;W;O;O;O;'
      '3;Etapes de r'#232'glement;Vous devez renseigner un libell'#233'.;W;O;O;O;'
      
        '4;Etapes de r'#232'glement;Le code que vous avez saisi existe d'#233'j'#224'. V' +
        'ous devez le modifier.;W;O;O;O;'
      'L'#39'enregistrement est inaccessible')
    Left = 390
    Top = 100
  end
  inherited HM2: THMsgBox
    Mess.Strings = (
      'Etape d'#39'encaissement :'
      'Etape de d'#233'caissement :'
      
        '2;Etapes de r'#232'glement;La devise du guide est diff'#233'rente de celle' +
        ' renseign'#233'e;W;O;O;O;'
      
        '3;Etapes de r'#232'glement;Le compte de tr'#233'sorerie du guide est diff'#233 +
        'rent de celui renseign'#233';W;O;O;O;'
      
        '4;Etapes de r'#232'glement;Le compte de g'#233'n'#233'ration du guide est diff'#233 +
        'rent de celui renseign'#233';W;O;O;O;'
      
        '5;Etapes de r'#232'glement;Un guide concernant les comptes et la devi' +
        'se renseign'#233's existe d'#232'j'#224'. Vous devez l'#39'utiliser.;W;O;O;O;'
      
        '6;Etapes de r'#232'glement;Il n'#39'existe pas de guide correspondant '#224' v' +
        'otre saisie, d'#233'sirez-vous en cr'#233'er un nouveau ?;Q;YNC;O;O;'
      'Compte issu du mode de paiement'
      
        '8;Etapes de r'#232'glement;D'#233'sirez-vous visualiser le guide cr'#233#233' ?;Q;' +
        'YNC;O;O;'
      '9'
      '10'
      '11'
      '12'
      '13'
      '14'
      
        '15;Etapes de r'#232'glement;Vous devez renseigner le type d'#39#233'tape.;W;' +
        'O;O;O;'
      
        '16;Etapes de r'#232'glement;Vous devez renseigner le compte de suivi ' +
        'de tr'#233'sorerie.;W;O;O;O;'
      
        '17;Etapes de r'#232'glement;Vous devez renseigner le compte de g'#233'n'#233'ra' +
        'tion.;W;O;O;O;'
      
        '18;Etapes de r'#232'glement;Vous devez renseigner la cat'#233'gorie de pai' +
        'ement.;W;O;O;O;'
      '19;Etapes de r'#232'glement;Vous devez renseigner la devise.;W;O;O;O;'
      
        '20;Etapes de r'#232'glement;Le comptes de suivi de tr'#233'sorerie et de g' +
        #233'n'#233'ration ne peuvent pas '#234'tre identiques;W;O;O;O;'
      
        '21;Etapes de r'#232'glement;La nature du compte de suivi de tr'#233'soreri' +
        'e n'#39'est pas en corr'#233'lation avec le type d'#39#233'tape !;W;O;O;O;'
      
        '22;Etapes de r'#232'glement;La nature ou la caract'#233'ristique de suivi ' +
        'de tr'#233'sorerie du compte de g'#233'n'#233'ration n'#39'est pas en corr'#233'lation a' +
        'vec le type d'#39#233'tape !;W;O;O;O;'
      
        '23;Etapes de r'#232'glement;La devise n'#39'est pas celle du compte banca' +
        'ire associ'#233'.;W;O;O;O;'
      
        '24;Etapes de r'#232'glement;L'#39'export CFONB n'#39'est possible qu'#39'avec la ' +
        'devise pivot !;W;O;O;O;'
      
        '25;Etapes de r'#232'glement;Vous n'#39'avez pas d'#233'fini de journal de Banq' +
        'ue ou d'#39'OD. La cr'#233'ation du guide associ'#233' est impossible;W;O;O;O;'
      ' ')
    Left = 444
    Top = 100
  end
  inherited HMTrad: THSystemMenu
    Left = 396
    Top = 172
  end
  object QMP: TQuery
    DatabaseName = 'SOC'
    UpdateMode = upWhereChanged
    Left = 444
    Top = 132
  end
end

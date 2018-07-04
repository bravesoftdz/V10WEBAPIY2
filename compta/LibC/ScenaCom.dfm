inherited FParamImpSup: TFParamImpSup
  Left = 126
  HelpContext = 1196100
  Caption = 'Sc'#233'nario d'#39'importation de fichier'
  ClientHeight = 380
  ClientWidth = 618
  PixelsPerInch = 96
  TextHeight = 13
  inherited FListe: THDBGrid
    Left = 392
    Width = 226
    Height = 338
    Columns = <
      item
        Expanded = False
        Title.Alignment = taCenter
        Title.Caption = 'Code'
        Width = 66
        Visible = True
      end
      item
        Expanded = False
        Title.Alignment = taCenter
        Title.Caption = 'Libell'#233
        Width = 119
        Visible = True
      end>
  end
  inherited Pappli: THPanel
    Width = 392
    Height = 338
    object TNT_NATURE: THLabel
      Left = 12
      Top = 7
      Width = 25
      Height = 13
      Caption = '&Code'
      FocusControl = FS_CODE
    end
    object TNT_LIBELLE: THLabel
      Left = 12
      Top = 32
      Width = 30
      Height = 13
      Caption = '&Libell'#233
      FocusControl = FS_LIBELLE
    end
    object HLabel1: THLabel
      Left = 12
      Top = 57
      Width = 32
      Height = 13
      Caption = '&Format'
      FocusControl = FS_FORMAT
    end
    object FS_LIBELLE: TDBEdit
      Left = 76
      Top = 28
      Width = 253
      Height = 21
      Ctl3D = True
      DataField = 'FS_LIBELLE'
      DataSource = STa
      ParentCtl3D = False
      TabOrder = 1
    end
    object FS_CODE: TDBEdit
      Left = 76
      Top = 3
      Width = 141
      Height = 21
      CharCase = ecUpperCase
      Ctl3D = True
      DataField = 'FS_CODE'
      DataSource = STa
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 0
    end
    object Pages: TPageControl
      Left = 0
      Top = 73
      Width = 392
      Height = 265
      ActivePage = Tinfos1
      Align = alBottom
      TabOrder = 3
      object Tinfos1: TTabSheet
        Caption = 'G'#233'n'#233'ral'
        object Label11: THLabel
          Left = 42
          Top = 217
          Width = 172
          Height = 13
          Caption = '&Mode de paiement de remplacement'
          FocusControl = FS_MPDEFAUT
        end
        object FS_TRESO: TDBCheckBox
          Left = 4
          Top = 3
          Width = 254
          Height = 17
          Alignment = taLeftJustify
          Caption = '&Fichier pour mouvement de tr'#233'sorerie'
          Ctl3D = True
          DataField = 'FS_TRESO'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 0
          ValueChecked = 'X'
          ValueUnchecked = '-'
          Visible = False
        end
        object FS_MPDEFAUT: THDBValComboBox
          Left = 220
          Top = 213
          Width = 145
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 10
          OnKeyDown = FS_MPDEFAUTKeyDown
          TagDispatch = 0
          DataType = 'TTMODEPAIE'
          DataField = 'FS_MPDEFAUT'
          DataSource = STa
        end
        object FS_TRAITETVA: TDBCheckBox
          Left = 40
          Top = 3
          Width = 255
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Traitement de la TVA '
          Ctl3D = True
          DataField = 'FS_TRAITETVA'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 1
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_TRAITECTR: TDBCheckBox
          Left = 40
          Top = 27
          Width = 255
          Height = 20
          Alignment = taLeftJustify
          Caption = '&Recalcul des contreparties'
          Ctl3D = True
          DataField = 'FS_TRAITECTR'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 2
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_RIBCLIENT: TDBCheckBox
          Left = 40
          Top = 51
          Width = 255
          Height = 20
          Alignment = taLeftJustify
          Caption = '&Affecter le rib principal des tiers clients'
          Ctl3D = True
          DataField = 'FS_RIBCLIENT'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 3
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_ANOUVEAUD: TDBCheckBox
          Left = 40
          Top = 98
          Width = 255
          Height = 20
          Alignment = taLeftJustify
          Caption = 'R'#233'cup'#233'rer les A-&Nouveaux en d'#233'tail'
          Ctl3D = True
          DataField = 'FS_ANOUVEAUD'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 5
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_RIBFOUR: TDBCheckBox
          Left = 40
          Top = 74
          Width = 255
          Height = 20
          Alignment = taLeftJustify
          Caption = '&Affecter le rib principal des tiers fournisseurs'
          Ctl3D = True
          DataField = 'FS_RIBFOUR'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 4
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_VALIDEECR: TDBCheckBox
          Left = 40
          Top = 122
          Width = 255
          Height = 20
          Alignment = taLeftJustify
          Caption = 'Valider les '#233'critures &import'#233'es'
          Ctl3D = True
          DataField = 'FS_VALIDEECR'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 6
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_TAUXDEVOUT: TDBCheckBox
          Left = 40
          Top = 146
          Width = 255
          Height = 20
          Alignment = taLeftJustify
          Caption = 'Recalcul du taux devise &out'
          Ctl3D = True
          DataField = 'FS_TAUXDEVOUT'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 7
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_FILTREAUX: TDBCheckBox
          Left = 40
          Top = 169
          Width = 255
          Height = 20
          Alignment = taLeftJustify
          Caption = 'R'#233#233'quilibrage anal&ytique automatique'
          Ctl3D = True
          DataField = 'FS_FILTREAUX'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 8
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_ECCDIV: TDBCheckBox
          Left = 40
          Top = 193
          Width = 255
          Height = 20
          Alignment = taLeftJustify
          Caption = 'Conversion en majuscule des comptes'
          Ctl3D = True
          DataField = 'FS_ECCDIV'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 9
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
      end
      object TabSheet3: TTabSheet
        Caption = 'Substitution'
        ImageIndex = 4
        object Label12: THLabel
          Left = 75
          Top = 35
          Width = 92
          Height = 13
          Caption = '--> Pour le compte :'
          FocusControl = FS_MPDEFAUT
        end
        object Label13: THLabel
          Left = 75
          Top = 99
          Width = 92
          Height = 13
          Caption = '--> Pour le compte :'
          FocusControl = FS_MPDEFAUT
        end
        object FS_COLLCLI: TDBCheckBox
          Left = 27
          Top = 11
          Width = 256
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Substitution par le collectif de la fiche &client '
          Ctl3D = True
          DataField = 'FS_COLLCLI'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 0
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_CPTCOLLCLI: THDBCpteEdit
          Left = 207
          Top = 32
          Width = 121
          Height = 21
          TabOrder = 1
          ZoomTable = TzGCollClient
          Vide = False
          Bourre = False
          okLocate = False
          SynJoker = False
          DataField = 'FS_CPTCOLLCLI'
          DataSource = STa
        end
        object FS_COLFOU: TDBCheckBox
          Left = 27
          Top = 75
          Width = 256
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Substitution par le collectif de la fiche &fournisseur '
          Ctl3D = True
          DataField = 'FS_COLFOU'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 2
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_CPTCOLLFOU: THDBCpteEdit
          Left = 207
          Top = 96
          Width = 121
          Height = 21
          TabOrder = 3
          ZoomTable = tzGCollFourn
          Vide = False
          Bourre = False
          okLocate = False
          SynJoker = False
          DataField = 'FS_CPTCOLLFOU'
          DataSource = STa
        end
        object GroupBox1: TGroupBox
          Left = 21
          Top = 132
          Width = 344
          Height = 93
          Caption = 'Cr'#233'ation de comptes de tiers... '
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 4
          object Label8: THLabel
            Left = 9
            Top = 28
            Width = 175
            Height = 13
            Caption = '&Mode de r'#232'glement de remplacement'
            FocusControl = FS_MRDEFAUT
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object Label9: THLabel
            Left = 9
            Top = 60
            Width = 159
            Height = 13
            Caption = '&R'#233'gime de TVA de remplacement'
            FocusControl = FS_TREGDEFAUT
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
          end
          object FS_MRDEFAUT: THDBValComboBox
            Left = 189
            Top = 24
            Width = 145
            Height = 21
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ItemHeight = 0
            ParentFont = False
            TabOrder = 0
            OnKeyDown = FS_MRDEFAUTKeyDown
            TagDispatch = 0
            DataType = 'TTMODEREGLE'
            DataField = 'FS_MRDEFAUT'
            DataSource = STa
          end
          object FS_TREGDEFAUT: THDBValComboBox
            Left = 189
            Top = 56
            Width = 145
            Height = 21
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -11
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ItemHeight = 0
            ParentFont = False
            TabOrder = 1
            OnKeyDown = FS_TREGDEFAUTKeyDown
            TagDispatch = 0
            DataType = 'TTREGIMETVA'
            DataField = 'FS_TREGDEFAUT'
            DataSource = STa
          end
        end
      end
      object TInfos2: TTabSheet
        Caption = 'Comptes inexistants'
        object Label1: THLabel
          Left = 40
          Top = 4
          Width = 158
          Height = 13
          Caption = 'Compte &g'#233'n'#233'ral de remplacement'
          FocusControl = FS_GENATTEND
        end
        object Label2: THLabel
          Left = 40
          Top = 28
          Width = 148
          Height = 13
          Caption = 'Compte &client de remplacement'
          FocusControl = FS_CLIATTEND
        end
        object Label3: THLabel
          Left = 40
          Top = 52
          Width = 174
          Height = 13
          Caption = 'Compte &fournisseur de remplacement'
          FocusControl = FS_FOUATTEND
        end
        object Label4: THLabel
          Left = 40
          Top = 76
          Width = 153
          Height = 13
          Caption = 'Compte &salari'#233' de remplacement'
          FocusControl = FS_SALATTEND
        end
        object Label5: THLabel
          Left = 40
          Top = 100
          Width = 173
          Height = 13
          Caption = 'Compte tiers &divers de remplacement'
          FocusControl = FS_DIVATTEND
        end
        object Label6: THLabel
          Left = 40
          Top = 124
          Width = 161
          Height = 13
          Caption = 'Section axe N'#176'&1 de remplacement'
          FocusControl = FS_SECTION1
        end
        object TFS_SECTION2: THLabel
          Left = 40
          Top = 148
          Width = 161
          Height = 13
          Caption = 'Section axe N'#176'&2 de remplacement'
          FocusControl = FS_SECTION2
        end
        object TFS_SECTION3: THLabel
          Left = 40
          Top = 172
          Width = 161
          Height = 13
          Caption = 'Section axe N'#176'&3 de remplacement'
          FocusControl = FS_SECTION3
        end
        object TFS_SECTION4: THLabel
          Left = 40
          Top = 196
          Width = 161
          Height = 13
          Caption = 'Section axe N'#176'&4 de remplacement'
          FocusControl = FS_SECTION4
        end
        object TFS_SECTION5: THLabel
          Left = 40
          Top = 220
          Width = 161
          Height = 13
          Caption = 'Section axe N'#176'&5 de remplacement'
          FocusControl = FS_SECTION5
        end
        object FS_GENATTEND: THDBCpteEdit
          Left = 233
          Top = 0
          Width = 121
          Height = 21
          TabOrder = 0
          ZoomTable = tzGeneral
          Vide = False
          Bourre = False
          okLocate = False
          SynJoker = False
          DataField = 'FS_GENATTEND'
          DataSource = STa
        end
        object FS_CLIATTEND: THDBCpteEdit
          Left = 233
          Top = 24
          Width = 121
          Height = 21
          TabOrder = 1
          ZoomTable = tzTclient
          Vide = False
          Bourre = False
          okLocate = False
          SynJoker = False
          DataField = 'FS_CLIATTEND'
          DataSource = STa
        end
        object FS_FOUATTEND: THDBCpteEdit
          Left = 233
          Top = 48
          Width = 121
          Height = 21
          TabOrder = 2
          ZoomTable = tzTfourn
          Vide = False
          Bourre = False
          okLocate = False
          SynJoker = False
          DataField = 'FS_FOUATTEND'
          DataSource = STa
        end
        object FS_SALATTEND: THDBCpteEdit
          Left = 233
          Top = 72
          Width = 121
          Height = 21
          TabOrder = 3
          ZoomTable = tzTsalarie
          Vide = False
          Bourre = False
          okLocate = False
          SynJoker = False
          DataField = 'FS_SALATTEND'
          DataSource = STa
        end
        object FS_DIVATTEND: THDBCpteEdit
          Left = 233
          Top = 96
          Width = 121
          Height = 21
          TabOrder = 4
          ZoomTable = tzTdivers
          Vide = False
          Bourre = False
          okLocate = False
          SynJoker = False
          DataField = 'FS_DIVATTEND'
          DataSource = STa
        end
        object FS_SECTION1: THDBCpteEdit
          Left = 233
          Top = 120
          Width = 121
          Height = 21
          TabOrder = 5
          ZoomTable = tzSection
          Vide = False
          Bourre = False
          okLocate = False
          SynJoker = False
          DataField = 'FS_SECTION1'
          DataSource = STa
        end
        object FS_SECTION2: THDBCpteEdit
          Left = 233
          Top = 144
          Width = 121
          Height = 21
          TabOrder = 6
          ZoomTable = tzSection2
          Vide = False
          Bourre = False
          okLocate = False
          SynJoker = False
          DataField = 'FS_SECTION2'
          DataSource = STa
        end
        object FS_SECTION3: THDBCpteEdit
          Left = 233
          Top = 168
          Width = 121
          Height = 21
          TabOrder = 7
          ZoomTable = tzSection3
          Vide = False
          Bourre = False
          okLocate = False
          SynJoker = False
          DataField = 'FS_SECTION3'
          DataSource = STa
        end
        object FS_SECTION4: THDBCpteEdit
          Left = 233
          Top = 192
          Width = 121
          Height = 21
          TabOrder = 8
          ZoomTable = tzSection4
          Vide = False
          Bourre = False
          okLocate = False
          SynJoker = False
          DataField = 'FS_SECTION4'
          DataSource = STa
        end
        object FS_SECTION5: THDBCpteEdit
          Left = 233
          Top = 216
          Width = 121
          Height = 21
          TabOrder = 9
          ZoomTable = tzSection5
          Vide = False
          Bourre = False
          okLocate = False
          SynJoker = False
          DataField = 'FS_SECTION5'
          DataSource = STa
        end
      end
      object TabSheet1: TTabSheet
        Caption = 'Correspondance'
        object FS_CORRIMPG: TDBCheckBox
          Left = 35
          Top = 5
          Width = 322
          Height = 17
          Alignment = taLeftJustify
          Caption = '&Utilisation du plan de correspondance des comptes g'#233'n'#233'raux '
          Ctl3D = True
          DataField = 'FS_CORRIMPG'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 0
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_CORRIMPT: TDBCheckBox
          Left = 35
          Top = 32
          Width = 322
          Height = 17
          Alignment = taLeftJustify
          Caption = '&Utilisation du plan de correspondance des &tiers'
          Ctl3D = True
          DataField = 'FS_CORRIMPT'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 1
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_CORRIMP1: TDBCheckBox
          Left = 35
          Top = 58
          Width = 322
          Height = 17
          Alignment = taLeftJustify
          Caption = 
            '&Utilisation du plan de correspondance des sections de l'#39'axe N'#176'&' +
            '1'
          Ctl3D = True
          DataField = 'FS_CORRIMP1'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 2
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_CORRIMP2: TDBCheckBox
          Left = 35
          Top = 85
          Width = 322
          Height = 17
          Alignment = taLeftJustify
          Caption = 
            '&Utilisation du plan de correspondance des sections de l'#39'axe N'#176'&' +
            '2'
          Ctl3D = True
          DataField = 'FS_CORRIMP2'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 3
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_CORRIMP3: TDBCheckBox
          Left = 35
          Top = 111
          Width = 322
          Height = 17
          Alignment = taLeftJustify
          Caption = 
            '&Utilisation du plan de correspondance des sections de l'#39'axe N'#176'&' +
            '3'
          Ctl3D = True
          DataField = 'FS_CORRIMP3'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 4
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_CORRIMP4: TDBCheckBox
          Left = 35
          Top = 138
          Width = 322
          Height = 17
          Alignment = taLeftJustify
          Caption = 
            '&Utilisation du plan de correspondance des sections de l'#39'axe N'#176'&' +
            '4'
          Ctl3D = True
          DataField = 'FS_CORRIMP4'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 5
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_CORRIMP5: TDBCheckBox
          Left = 35
          Top = 164
          Width = 322
          Height = 17
          Alignment = taLeftJustify
          Caption = 
            '&Utilisation du plan de correspondance des sections de l'#39'axe N'#176'&' +
            '5'
          Ctl3D = True
          DataField = 'FS_CORRIMP5'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 6
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_CORRIMPMP: TDBCheckBox
          Left = 35
          Top = 191
          Width = 322
          Height = 17
          Alignment = taLeftJustify
          Caption = '&Utilisation du plan de correspondance des modes de &paiement'
          Ctl3D = True
          DataField = 'FS_CORRIMPMP'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 7
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_CORRIMPJAL: TDBCheckBox
          Left = 35
          Top = 217
          Width = 322
          Height = 17
          Alignment = taLeftJustify
          Caption = '&Utilisation du plan de correspondance des &journaux'
          Ctl3D = True
          DataField = 'FS_CORRIMPJAL'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 8
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
      end
      object TabSheet2: TTabSheet
        Caption = 'Import. auto.'
        object Label14: THLabel
          Left = 60
          Top = 22
          Width = 92
          Height = 13
          Caption = '&Chemin par d'#233'faut :'
          FocusControl = FS_CHEMIN
        end
        object Label15: THLabel
          Left = 60
          Top = 62
          Width = 133
          Height = 13
          Caption = '&Pr'#233'fixe du fichier '#224' importer :'
          FocusControl = FS_PREFIXE
        end
        object Label16: THLabel
          Left = 60
          Top = 102
          Width = 133
          Height = 13
          Caption = '&Suffixe du fichier '#224' importer :'
          FocusControl = FS_SUFFIXE
        end
        object Label17: THLabel
          Left = 27
          Top = 202
          Width = 136
          Height = 13
          Caption = 'Nom du fichier &interm'#233'diaire :'
          FocusControl = FS_NOMINTER
          Visible = False
        end
        object Label18: THLabel
          Left = 27
          Top = 192
          Width = 108
          Height = 13
          Caption = 'Nim du fichier de &rejet :'
          FocusControl = FS_NOMREJET
          Visible = False
        end
        object RechFile: TToolbarButton97
          Left = 328
          Top = 18
          Width = 16
          Height = 21
          Hint = 'Parcourir'
          Caption = '...'
          Flat = False
          ParentShowHint = False
          ShowHint = True
          OnClick = RechFileClick
        end
        object FS_CHEMIN: TDBEdit
          Left = 207
          Top = 18
          Width = 121
          Height = 21
          DataField = 'FS_CHEMIN'
          DataSource = STa
          TabOrder = 0
        end
        object FS_PREFIXE: TDBEdit
          Left = 207
          Top = 58
          Width = 121
          Height = 21
          DataField = 'FS_PREFIXE'
          DataSource = STa
          TabOrder = 1
        end
        object FS_SUFFIXE: TDBEdit
          Left = 207
          Top = 98
          Width = 121
          Height = 21
          DataField = 'FS_SUFFIXE'
          DataSource = STa
          TabOrder = 2
        end
        object FS_NOMINTER: TDBEdit
          Left = 174
          Top = 198
          Width = 121
          Height = 21
          DataField = 'FS_NOMINTER'
          DataSource = STa
          TabOrder = 3
          Visible = False
        end
        object FS_NOMREJET: TDBEdit
          Left = 174
          Top = 188
          Width = 121
          Height = 21
          DataField = 'FS_NOMREJET'
          DataSource = STa
          TabOrder = 4
          Visible = False
        end
        object FS_DOUBLON: TDBCheckBox
          Left = 60
          Top = 134
          Width = 268
          Height = 20
          Alignment = taLeftJustify
          Caption = '&Contr'#244'le de doublons'
          Ctl3D = True
          DataField = 'FS_DOUBLON'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 5
          ValueChecked = 'X'
          ValueUnchecked = '-'
          OnClick = FS_DOUBLONClick
        end
        object FS_FORCEPIECE: TDBCheckBox
          Left = 60
          Top = 167
          Width = 268
          Height = 21
          Alignment = taLeftJustify
          Caption = '&Forcer les num'#233'ros de pi'#232'ces'
          Ctl3D = True
          DataField = 'FS_FORCEPIECE'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 6
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_FILTREGENE: TDBCheckBox
          Left = 60
          Top = 199
          Width = 268
          Height = 22
          Alignment = taLeftJustify
          Caption = '&D'#233'truire le fichier origine'
          Ctl3D = True
          DataField = 'FS_FILTREGENE'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 7
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
      end
    end
    object FS_FORMAT: THDBValComboBox
      Left = 76
      Top = 53
      Width = 145
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 2
      TagDispatch = 0
      DataType = 'TTFORMATECRITURE'
      DataField = 'FS_FORMAT'
      DataSource = STa
    end
  end
  inherited DBNav: TDBNavigator
    Left = 384
    Hints.Strings = ()
  end
  inherited Dock971: TDock97
    Top = 338
    Width = 618
    inherited PBouton: TToolWindow97
      ClientWidth = 618
      ClientAreaWidth = 618
      inherited BImprimer: TToolbarButton97
        Left = 232
      end
      inherited BValider: TToolbarButton97
        Left = 266
      end
      inherited HelpBtn: TToolbarButton97
        Left = 330
      end
      inherited BFerme: TToolbarButton97
        Left = 298
      end
      inherited BFirst: TToolbarButton97
        Left = 364
        Width = 62
      end
      inherited BPrev: TToolbarButton97
        Left = 426
        Width = 62
      end
      inherited BNext: TToolbarButton97
        Left = 488
        Width = 62
      end
      inherited BLast: TToolbarButton97
        Left = 550
        Width = 62
      end
      inherited FAutoSave: THCheckbox
        Top = 4
      end
    end
  end
  inherited Ta: THTable
    IndexName = 'FS_CLE1'
    TableName = 'FMTSUP'
    Left = 440
    Top = 52
    object TaFS_IMPORT: TStringField
      FieldName = 'FS_IMPORT'
      Size = 1
    end
    object TaFS_NATURE: TStringField
      FieldName = 'FS_NATURE'
      Size = 3
    end
    object TaFS_FORMAT: TStringField
      FieldName = 'FS_FORMAT'
      Size = 3
    end
    object TaFS_CODE: TStringField
      FieldName = 'FS_CODE'
      Size = 17
    end
    object TaFS_TRESO: TStringField
      FieldName = 'FS_TRESO'
      Size = 1
    end
    object TaFS_GENATTEND: TStringField
      FieldName = 'FS_GENATTEND'
      Size = 17
    end
    object TaFS_CLIATTEND: TStringField
      FieldName = 'FS_CLIATTEND'
      Size = 17
    end
    object TaFS_FOUATTEND: TStringField
      FieldName = 'FS_FOUATTEND'
      Size = 17
    end
    object TaFS_SALATTEND: TStringField
      FieldName = 'FS_SALATTEND'
      Size = 17
    end
    object TaFS_DIVATTEND: TStringField
      FieldName = 'FS_DIVATTEND'
      Size = 17
    end
    object TaFS_SECTION1: TStringField
      FieldName = 'FS_SECTION1'
      Size = 17
    end
    object TaFS_SECTION2: TStringField
      FieldName = 'FS_SECTION2'
      Size = 17
    end
    object TaFS_SECTION3: TStringField
      FieldName = 'FS_SECTION3'
      Size = 17
    end
    object TaFS_SECTION4: TStringField
      FieldName = 'FS_SECTION4'
      Size = 17
    end
    object TaFS_SECTION5: TStringField
      FieldName = 'FS_SECTION5'
      Size = 17
    end
    object TaFS_COLLCLI: TStringField
      FieldName = 'FS_COLLCLI'
      Size = 1
    end
    object TaFS_CPTCOLLCLI: TStringField
      FieldName = 'FS_CPTCOLLCLI'
      Size = 17
    end
    object TaFS_COLFOU: TStringField
      FieldName = 'FS_COLFOU'
      Size = 1
    end
    object TaFS_CPTCOLLFOU: TStringField
      FieldName = 'FS_CPTCOLLFOU'
      Size = 17
    end
    object TaFS_CORRIMPG: TStringField
      FieldName = 'FS_CORRIMPG'
      Size = 1
    end
    object TaFS_CORRIMPT: TStringField
      FieldName = 'FS_CORRIMPT'
      Size = 1
    end
    object TaFS_CORRIMP1: TStringField
      FieldName = 'FS_CORRIMP1'
      Size = 1
    end
    object TaFS_CORRIMP2: TStringField
      FieldName = 'FS_CORRIMP2'
      Size = 1
    end
    object TaFS_CORRIMP3: TStringField
      FieldName = 'FS_CORRIMP3'
      Size = 1
    end
    object TaFS_CORRIMP4: TStringField
      FieldName = 'FS_CORRIMP4'
      Size = 1
    end
    object TaFS_CORRIMP5: TStringField
      FieldName = 'FS_CORRIMP5'
      Size = 1
    end
    object TaFS_MPDEFAUT: TStringField
      FieldName = 'FS_MPDEFAUT'
      Size = 3
    end
    object TaFS_MRDEFAUT: TStringField
      FieldName = 'FS_MRDEFAUT'
      Size = 3
    end
    object TaFS_CHEMIN: TStringField
      FieldName = 'FS_CHEMIN'
      Size = 200
    end
    object TaFS_PREFIXE: TStringField
      FieldName = 'FS_PREFIXE'
      Size = 35
    end
    object TaFS_SUFFIXE: TStringField
      FieldName = 'FS_SUFFIXE'
      Size = 3
    end
    object TaFS_NOMINTER: TStringField
      FieldName = 'FS_NOMINTER'
      Size = 70
    end
    object TaFS_NOMREJET: TStringField
      FieldName = 'FS_NOMREJET'
      Size = 70
    end
    object TaFS_TREGDEFAUT: TStringField
      FieldName = 'FS_TREGDEFAUT'
      Size = 3
    end
    object TaFS_TMRDEFAUT: TStringField
      FieldName = 'FS_TMRDEFAUT'
      Size = 3
    end
    object TaFS_FILTREGENE: TStringField
      FieldName = 'FS_FILTREGENE'
      Size = 1
    end
    object TaFS_NATUREGENE: TStringField
      FieldName = 'FS_NATUREGENE'
      Size = 3
    end
    object TaFS_FILTREAUX: TStringField
      FieldName = 'FS_FILTREAUX'
      Size = 1
    end
    object TaFS_NATUREAUX: TStringField
      FieldName = 'FS_NATUREAUX'
      Size = 3
    end
    object TaFS_NATMVT: TStringField
      FieldName = 'FS_NATMVT'
      Size = 3
    end
    object TaFS_TRAITETVA: TStringField
      FieldName = 'FS_TRAITETVA'
      Size = 1
    end
    object TaFS_TRAITECTR: TStringField
      FieldName = 'FS_TRAITECTR'
      Size = 1
    end
    object TaFS_LIBMP: TStringField
      FieldName = 'FS_LIBMP'
      Size = 1
    end
    object TaFS_LIBELLE: TStringField
      FieldName = 'FS_LIBELLE'
      Size = 35
    end
    object TaFS_RIBCLIENT: TStringField
      FieldName = 'FS_RIBCLIENT'
      Size = 1
    end
    object TaFS_DOUBLON: TStringField
      FieldName = 'FS_DOUBLON'
      Size = 1
    end
    object TaFS_FORCEPIECE: TStringField
      FieldName = 'FS_FORCEPIECE'
      Size = 1
    end
    object TaFS_ANOUVEAUD: TStringField
      FieldName = 'FS_ANOUVEAUD'
      Size = 1
    end
    object TaFS_RIBFOUR: TStringField
      FieldName = 'FS_RIBFOUR'
      Size = 1
    end
    object TaFS_VALIDEECR: TStringField
      FieldName = 'FS_VALIDEECR'
      Size = 1
    end
    object TaFS_DOUBLONM: TStringField
      FieldName = 'FS_DOUBLONM'
      Size = 1
    end
    object TaFS_FORCEPIECEM: TStringField
      FieldName = 'FS_FORCEPIECEM'
      Size = 1
    end
    object TaFS_CORRIMPMP: TStringField
      FieldName = 'FS_CORRIMPMP'
      Size = 1
    end
    object TaFS_CORRIMPJAL: TStringField
      FieldName = 'FS_CORRIMPJAL'
      Size = 1
    end
    object TaFS_TAUXDEVOUT: TStringField
      FieldName = 'FS_TAUXDEVOUT'
      Size = 1
    end
    object TaFS_ECCDIV: TStringField
      FieldName = 'FS_ECCDIV'
      FixedChar = True
      Size = 1
    end
  end
  inherited HM2: THMsgBox
    Mess.Strings = (
      'Sc'#233'nario d'#39'importation de mouvements : '
      'Sc'#233'nario d'#39'exportation de mouvements : '
      'Format C'#233'gid '#233'tendu'
      'Format C'#233'gid simplifi'#233
      'Format Halley '#233'tendu'
      'Format Halley simplifi'#233
      'Format Saari coeur de gamme'
      'Format Saari N'#233'goce ')
    Left = 556
    Top = 48
  end
  inherited HMTrad: THSystemMenu
    Left = 492
    Top = 52
  end
end

inherited FParamExpSup: TFParamExpSup
  Left = 265
  Top = 191
  HelpContext = 1196100
  Caption = 'Sc'#233'nario d'#39'importation de fichier'
  ClientHeight = 380
  ClientWidth = 618
  PixelsPerInch = 96
  TextHeight = 13
  inherited FListe: THDBGrid
    Left = 366
    Width = 252
    Height = 338
    Columns = <
      item
        Expanded = False
        Title.Alignment = taCenter
        Title.Caption = 'Code'
        Width = 115
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
    Width = 366
    Height = 338
    object TNT_NATURE: THLabel
      Left = 12
      Top = 15
      Width = 25
      Height = 13
      Caption = '&Code'
      FocusControl = FS_CODE
    end
    object TNT_LIBELLE: THLabel
      Left = 12
      Top = 43
      Width = 30
      Height = 13
      Caption = '&Libell'#233
      FocusControl = FS_LIBELLE
    end
    object HLabel1: THLabel
      Left = 12
      Top = 72
      Width = 32
      Height = 13
      Caption = '&Format'
      FocusControl = FS_FORMAT
    end
    object FS_LIBELLE: TDBEdit
      Left = 76
      Top = 39
      Width = 278
      Height = 21
      Ctl3D = True
      DataField = 'FS_LIBELLE'
      DataSource = STa
      ParentCtl3D = False
      TabOrder = 1
    end
    object FS_CODE: TDBEdit
      Left = 76
      Top = 11
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
      Top = 92
      Width = 366
      Height = 246
      ActivePage = Tinfos1
      Align = alBottom
      TabOrder = 3
      object Tinfos1: TTabSheet
        Caption = 'G'#233'n'#233'ral'
        object Label12: TLabel
          Left = 22
          Top = 104
          Width = 154
          Height = 13
          Caption = '--> Nature de compte par d'#233'faut:'
          FocusControl = FS_NATUREGENE
        end
        object Label13: TLabel
          Left = 22
          Top = 178
          Width = 138
          Height = 13
          Caption = '--> Nature de tiers par d'#233'faut:'
          FocusControl = FS_NATUREAUX
        end
        object Label1: TLabel
          Left = 6
          Top = 32
          Width = 106
          Height = 13
          Caption = '&Nature de l'#39'exportation'
          FocusControl = FS_NATMVT
        end
        object FS_FILTREGENE: TDBCheckBox
          Left = 4
          Top = 75
          Width = 255
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Filtre sur les comptes &g'#233'n'#233'raux '
          Ctl3D = True
          DataField = 'FS_FILTREGENE'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 1
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_FILTREAUX: TDBCheckBox
          Left = 4
          Top = 149
          Width = 255
          Height = 17
          Alignment = taLeftJustify
          Caption = 'Filtre sur les comptes &auxiliaires'
          Ctl3D = True
          DataField = 'FS_FILTREAUX'
          DataSource = STa
          ParentCtl3D = False
          TabOrder = 3
          ValueChecked = 'X'
          ValueUnchecked = '-'
        end
        object FS_NATUREGENE: THDBValComboBox
          Left = 196
          Top = 100
          Width = 145
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 2
          TagDispatch = 0
          Vide = True
          DataType = 'TTNATGENE'
          DataField = 'FS_NATUREGENE'
          DataSource = STa
        end
        object FS_NATUREAUX: THDBValComboBox
          Left = 196
          Top = 174
          Width = 145
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 4
          TagDispatch = 0
          Vide = True
          DataType = 'TTNATTIERSCPTA'
          DataField = 'FS_NATUREAUX'
          DataSource = STa
        end
        object FS_NATMVT: THDBValComboBox
          Left = 196
          Top = 28
          Width = 145
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 0
          OnChange = FS_NATMVTChange
          TagDispatch = 0
          DataType = 'TTNATMVTIE'
          DataField = 'FS_NATMVT'
          DataSource = STa
        end
      end
    end
    object FS_FORMAT: THDBValComboBox
      Left = 76
      Top = 68
      Width = 145
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 2
      TagDispatch = 0
      DataField = 'FS_FORMAT'
      DataSource = STa
    end
    object FS_FORMAT2: THValComboBox
      Left = 340
      Top = 68
      Width = 17
      Height = 21
      ItemHeight = 13
      TabOrder = 4
      Visible = False
      TagDispatch = 0
      DataType = 'TTFORMATECRITUREEXP'
    end
  end
  inherited DBNav: TDBNavigator
    Left = 404
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
    end
  end
  inherited Ta: THTable
    IndexName = 'FS_CLE1'
    TableName = 'FMTSUP'
    Left = 472
    Top = 144
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
    object TaFS_LIBELLE: TStringField
      FieldName = 'FS_LIBELLE'
      Size = 35
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
  end
  inherited HM2: THMsgBox
    Mess.Strings = (
      'Sc'#233'nario d'#39'importation de mouvements : '
      'Sc'#233'nario d'#39'exportation de mouvements : ')
    Left = 556
    Top = 48
  end
  inherited HMTrad: THSystemMenu
    Left = 492
    Top = 52
  end
end

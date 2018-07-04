inherited FSouche: TFSouche
  Left = 294
  Top = 175
  HelpContext = 1300000
  Caption = 'Compteur :'
  ClientHeight = 244
  ClientWidth = 570
  PixelsPerInch = 96
  TextHeight = 13
  inherited FListe: THDBGrid
    Left = 384
    Width = 186
    Height = 202
    Columns = <
      item
        Expanded = False
        Title.Alignment = taCenter
        Title.Caption = 'Code'
        Width = 40
        Visible = True
      end
      item
        Expanded = False
        Title.Alignment = taCenter
        Title.Caption = 'Libell'#233
        Width = 111
        Visible = True
      end>
  end
  inherited Pappli: THPanel
    Width = 384
    Height = 202
    object TSH_SOUCHE: TLabel
      Left = 5
      Top = 12
      Width = 25
      Height = 13
      Caption = '&Code'
      FocusControl = SH_SOUCHE
    end
    object TSH_LIBELLE: TLabel
      Left = 115
      Top = 13
      Width = 30
      Height = 13
      Caption = '&Libell'#233
      FocusControl = SH_LIBELLE
    end
    object TSH_ABREGE: TLabel
      Left = 5
      Top = 41
      Width = 34
      Height = 13
      Caption = '&Abr'#233'g'#233
      FocusControl = SH_ABREGE
    end
    object TSH_NUMDEPART: TLabel
      Left = 179
      Top = 41
      Width = 32
      Height = 13
      Caption = '&D'#233'part'
      FocusControl = SH_NUMDEPART
    end
    object TSH_JOURNAL: TLabel
      Left = 179
      Top = 68
      Width = 34
      Height = 13
      Caption = '&Journal'
      FocusControl = SH_JOURNAL
    end
    object TSH_NATUREPIECE: TLabel
      Left = 5
      Top = 68
      Width = 47
      Height = 13
      Caption = '&Nat.Pi'#232'ce'
      Enabled = False
      FocusControl = SH_NATUREPIECE
    end
    object TSH_DATEDEBUT: TLabel
      Left = 83
      Top = 147
      Width = 53
      Height = 13
      Caption = 'Date d'#233'b&ut'
      FocusControl = SH_DATEDEBUT
      Visible = False
    end
    object TSH_DATEFIN: TLabel
      Left = 199
      Top = 146
      Width = 37
      Height = 13
      Caption = 'Da&te fin'
      FocusControl = SH_DATEFIN
      Visible = False
    end
    object Label2: TLabel
      Left = 5
      Top = 122
      Width = 50
      Height = 13
      Caption = '&R'#233'f'#233'rence'
      FocusControl = SH_MASQUENUM
    end
    object TSH_NUMDEPARTS: TLabel
      Left = 101
      Top = 177
      Width = 61
      Height = 13
      Caption = 'D'#233'part (&N+1)'
      Color = clYellow
      FocusControl = SH_NUMDEPARTS
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      Visible = False
    end
    object SH_SOUCHE: TDBEdit
      Left = 62
      Top = 9
      Width = 47
      Height = 21
      CharCase = ecUpperCase
      Ctl3D = True
      DataField = 'SH_SOUCHE'
      DataSource = STa
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 0
    end
    object SH_LIBELLE: TDBEdit
      Left = 151
      Top = 9
      Width = 194
      Height = 21
      Ctl3D = True
      DataField = 'SH_LIBELLE'
      DataSource = STa
      ParentCtl3D = False
      TabOrder = 1
    end
    object SH_ABREGE: TDBEdit
      Left = 62
      Top = 38
      Width = 110
      Height = 21
      Ctl3D = True
      DataField = 'SH_ABREGE'
      DataSource = STa
      ParentCtl3D = False
      TabOrder = 2
    end
    object SH_NUMDEPART: THDBSpinEdit
      Left = 235
      Top = 38
      Width = 110
      Height = 22
      MaxLength = 17
      MaxValue = 2000000000
      MinValue = 1
      TabOrder = 3
      Value = 1
      DataField = 'SH_NUMDEPART'
      DataSource = STa
    end
    object SH_JOURNAL: THDBValComboBox
      Left = 235
      Top = 65
      Width = 110
      Height = 21
      Style = csDropDownList
      Ctl3D = True
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 7
      TagDispatch = 0
      DataType = 'TTJALOD'
      DataField = 'SH_JOURNAL'
      DataSource = STa
    end
    object SH_NATUREPIECE: THDBValComboBox
      Left = 62
      Top = 65
      Width = 110
      Height = 21
      Style = csDropDownList
      Ctl3D = True
      Enabled = False
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 6
      TagDispatch = 0
      DataType = 'TTNATUREPIECE'
      DataField = 'SH_NATUREPIECE'
      DataSource = STa
    end
    object SH_DATEDEBUT: TDBEdit
      Left = 145
      Top = 144
      Width = 45
      Height = 21
      Color = clYellow
      Ctl3D = True
      DataField = 'SH_DATEDEBUT'
      DataSource = STa
      ParentCtl3D = False
      TabOrder = 8
      Visible = False
      OnExit = SH_DATEDEBUTExit
    end
    object SH_DATEFIN: TDBEdit
      Left = 238
      Top = 147
      Width = 45
      Height = 21
      Color = clYellow
      Ctl3D = True
      DataField = 'SH_DATEFIN'
      DataSource = STa
      ParentCtl3D = False
      TabOrder = 9
      Visible = False
      OnExit = SH_DATEFINExit
    end
    object SH_MASQUENUM: TDBEdit
      Left = 62
      Top = 119
      Width = 283
      Height = 21
      Ctl3D = True
      DataField = 'SH_MASQUENUM'
      DataSource = STa
      ParentCtl3D = False
      TabOrder = 14
    end
    object SH_TYPE: TDBCheckBox
      Left = 3
      Top = 93
      Width = 72
      Height = 17
      Alignment = taLeftJustify
      Caption = '&Relev'#233
      DataField = 'SH_TYPE'
      DataSource = STa
      TabOrder = 10
      ValueChecked = 'REL'
      ValueUnchecked = 'CPT'
      OnMouseUp = SH_TYPEMouseUp
    end
    object SH_FERME: TDBCheckBox
      Left = 12
      Top = 151
      Width = 75
      Height = 17
      Alignment = taLeftJustify
      Caption = '&Ferm'#233
      Color = clYellow
      DataField = 'SH_FERME'
      DataSource = STa
      ParentColor = False
      TabOrder = 11
      ValueChecked = 'X'
      ValueUnchecked = '-'
      Visible = False
    end
    object SH_ANALYTIQUE: TDBCheckBox
      Left = 270
      Top = 93
      Width = 75
      Height = 17
      Alignment = taLeftJustify
      Caption = 'Anal&ytique'
      DataField = 'SH_ANALYTIQUE'
      DataSource = STa
      TabOrder = 13
      ValueChecked = 'X'
      ValueUnchecked = '-'
      OnMouseUp = SH_ANALYTIQUEMouseUp
    end
    object SH_SIMULATION: TDBCheckBox
      Left = 139
      Top = 93
      Width = 75
      Height = 17
      Alignment = taLeftJustify
      Caption = 'Si&mulation'
      DataField = 'SH_SIMULATION'
      DataSource = STa
      TabOrder = 12
      ValueChecked = 'X'
      ValueUnchecked = '-'
      OnMouseUp = SH_SIMULATIONMouseUp
    end
    object SH_SOUCHEEXO: TDBCheckBox
      Left = 59
      Top = 176
      Width = 35
      Height = 17
      Alignment = taLeftJustify
      Caption = 'Souche &multi-exercice'
      Color = clYellow
      DataField = 'SH_SOUCHEEXO'
      DataSource = STa
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      TabOrder = 4
      ValueChecked = 'X'
      ValueUnchecked = '-'
      Visible = False
      OnClick = SH_SOUCHEEXOClick
    end
    object SH_NUMDEPARTS: THDBSpinEdit
      Left = 159
      Top = 167
      Width = 43
      Height = 22
      Color = clYellow
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 17
      MaxValue = 2000000000
      MinValue = 1
      ParentFont = False
      TabOrder = 5
      Value = 1
      Visible = False
      DataField = 'SH_NUMDEPARTS'
      DataSource = STa
    end
  end
  inherited DBNav: TDBNavigator
    Left = 376
    Top = 43
    Hints.Strings = ()
  end
  inherited Dock971: TDock97
    Top = 202
    Width = 570
    inherited PBouton: TToolWindow97
      ClientWidth = 570
      ClientAreaWidth = 570
      inherited BImprimer: TToolbarButton97
        Left = 228
      end
      inherited BValider: TToolbarButton97
        Left = 260
      end
      inherited HelpBtn: TToolbarButton97
        Left = 324
      end
      inherited BFerme: TToolbarButton97
        Left = 292
      end
      inherited BFirst: TToolbarButton97
        Left = 353
      end
      inherited BPrev: TToolbarButton97
        Left = 398
      end
      inherited BNext: TToolbarButton97
        Left = 443
      end
      inherited BLast: TToolbarButton97
        Left = 488
      end
      inherited FAutoSave: THCheckbox
        Width = 24
      end
    end
  end
  inherited Ta: THTable
    Filter = 'SH_TYPE='#39'CPT'#39' OR SH_TYPE='#39'REL'#39
    TableName = 'SOUCHE'
    dataBaseName = 'soc'
    Left = 376
    Top = 64
    object TaSH_TYPE: TStringField
      FieldName = 'SH_TYPE'
      Size = 3
    end
    object TaSH_SOUCHE: TStringField
      FieldName = 'SH_SOUCHE'
      Size = 3
    end
    object TaSH_LIBELLE: TStringField
      FieldName = 'SH_LIBELLE'
      Size = 35
    end
    object TaSH_ABREGE: TStringField
      FieldName = 'SH_ABREGE'
      Size = 17
    end
    object TaSH_NATUREPIECE: TStringField
      FieldName = 'SH_NATUREPIECE'
      Size = 3
    end
    object TaSH_NUMDEPART: TIntegerField
      FieldName = 'SH_NUMDEPART'
    end
    object TaSH_SIMULATION: TStringField
      FieldName = 'SH_SIMULATION'
      Size = 1
    end
    object TaSH_JOURNAL: TStringField
      FieldName = 'SH_JOURNAL'
      Size = 3
    end
    object TaSH_MASQUENUM: TStringField
      FieldName = 'SH_MASQUENUM'
      Size = 17
    end
    object TaSH_SOCIETE: TStringField
      FieldName = 'SH_SOCIETE'
      Size = 3
    end
    object TaSH_DATEDEBUT: TDateTimeField
      FieldName = 'SH_DATEDEBUT'
      DisplayFormat = 'dd mmm yyyy'
      EditMask = '!99/99/0000;1;_'
    end
    object TaSH_DATEFIN: TDateTimeField
      FieldName = 'SH_DATEFIN'
      DisplayFormat = 'dd mmm yyyy'
      EditMask = '!99/99/0000;1;_'
    end
    object TaSH_FERME: TStringField
      FieldName = 'SH_FERME'
      Size = 1
    end
    object TaSH_ANALYTIQUE: TStringField
      FieldName = 'SH_ANALYTIQUE'
      Size = 1
    end
    object TaSH_NATUREPIECEG: TStringField
      FieldName = 'SH_NATUREPIECEG'
      Size = 3
    end
    object TaSH_NUMDEPARTS: TIntegerField
      FieldName = 'SH_NUMDEPARTS'
    end
    object TaSH_NUMDEPARTP: TIntegerField
      FieldName = 'SH_NUMDEPARTP'
    end
    object TaSH_SOUCHEEXO: TStringField
      FieldName = 'SH_SOUCHEEXO'
      Size = 1
    end
    object TaSH_ENTITY: TIntegerField
      FieldName = 'SH_ENTITY'
    end
  end
  inherited STa: TDataSource
    Left = 428
    Top = 64
  end
  inherited HM: THMsgBox
    Mess.Strings = (
      
        '0;Compteurs;Voulez-vous enregistrer les modifications ?;Q;YNC;Y;' +
        'C;'
      
        '1;Compteurs;Confirmez-vous la suppression de l'#39'enregistrement ?;' +
        'Q;YNC;N;C;'
      '2;Compteurs;Vous devez renseigner un code.;W;O;O;O;'
      '3;Compteurs;Vous devez renseigner un libell'#233'.;W;O;O;O;'
      
        '4;Compteurs;Le code que vous avez saisi existe d'#233'j'#224'. Vous devez ' +
        'le modifier.;W;O;O;O;'
      'L'#39'enregistrement est inaccessible')
    Left = 428
    Top = 112
  end
  inherited HM2: THMsgBox
    Mess.Strings = (
      
        '0;Compteurs;La date que vous avez renseign'#233'e n'#39'est pas valide.;W' +
        ';O;O;O;'
      
        '1;Compteurs;Vous devez renseigner une dateb de d'#233'but et une date' +
        ' de fin.;W;O;O;O;'
      
        '2;Compteurs;La date de fin doit '#234'tre sup'#233'rieure ou '#233'gale '#224' la da' +
        'te de d'#233'but;W;O;O;O;'
      
        '3;Compteurs;Il ne peut y avoir qu'#39'un seul compteur pour les rele' +
        'v'#233's de factures.;W;O;O;O;'
      
        '4;Compteurs;Ce compteur est utilis'#233' par un journal qui est mouve' +
        'ment'#233'. Vous ne pouvez pas d'#233'cr'#233'menter le num'#233'ro de d'#233'part.;W;O;O' +
        ';O;'
      
        '5;Compteurs;Ce compteur est utilis'#233' par un budget qui est mouvem' +
        'ent'#233'. Vous ne pouvez pas d'#233'cr'#233'menter le num'#233'ro de d'#233'part.;W;O;O;' +
        'O;'
      
        '6;Compteurs;Ce compteur est utilis'#233' par des pi'#232'ces. Vous ne pouv' +
        'ez pas d'#233'cr'#233'menter le num'#233'ro de d'#233'part.;W;O;O;O;'
      '')
    Left = 376
    Top = 112
  end
  inherited HMTrad: THSystemMenu
    Left = 468
    Top = 112
  end
end

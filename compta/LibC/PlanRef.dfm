inherited FPlanRef: TFPlanRef
  Left = 182
  Top = 202
  HelpContext = 1315000
  Caption = 'Plan de r'#233'f'#233'rence : '
  ClientHeight = 364
  ClientWidth = 612
  PixelsPerInch = 96
  TextHeight = 13
  inherited FListe: THDBGrid
    Left = 318
    Width = 294
    Height = 322
    Columns = <
      item
        Expanded = False
        FieldName = 'PR_COMPTE'
        Title.Alignment = taCenter
        Title.Caption = 'Code'
        Width = 57
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'PR_LIBELLE'
        Title.Alignment = taCenter
        Title.Caption = 'Libell'#233
        Width = 204
        Visible = True
      end>
  end
  inherited Pappli: THPanel
    Width = 318
    Height = 322
    object TPR_COMPTE: THLabel
      Left = 5
      Top = 50
      Width = 36
      Height = 13
      Caption = '&Compte'
      FocusControl = PR_COMPTE
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object tPR_ABREGE: THLabel
      Left = 164
      Top = 50
      Width = 34
      Height = 13
      Caption = '&Abr'#233'g'#233
      FocusControl = PR_ABREGE
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TPR_LIBELLE: THLabel
      Left = 5
      Top = 74
      Width = 30
      Height = 13
      Caption = '&Libell'#233
      FocusControl = PR_LIBELLE
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TPR_SENS: THLabel
      Left = 5
      Top = 98
      Width = 24
      Height = 13
      Caption = '&Sens'
      FocusControl = PR_SENS
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TPR_NATUREGENE: THLabel
      Left = 164
      Top = 98
      Width = 32
      Height = 13
      Caption = '&Nature'
      FocusControl = PR_NATUREGENE
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object PR_COMPTE: TDBEdit
      Left = 44
      Top = 46
      Width = 112
      Height = 21
      CharCase = ecUpperCase
      Ctl3D = True
      DataField = 'PR_COMPTE'
      DataSource = STa
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 2
      OnExit = PR_COMPTEExit
    end
    object PR_ABREGE: TDBEdit
      Left = 200
      Top = 46
      Width = 112
      Height = 21
      Ctl3D = True
      DataField = 'PR_ABREGE'
      DataSource = STa
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 3
    end
    object PR_LIBELLE: TDBEdit
      Left = 44
      Top = 70
      Width = 269
      Height = 21
      Ctl3D = True
      DataField = 'PR_LIBELLE'
      DataSource = STa
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 4
    end
    object PR_SENS: THDBValComboBox
      Left = 44
      Top = 94
      Width = 112
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
      TabOrder = 5
      TagDispatch = 0
      DataType = 'TTSENS'
      DataField = 'PR_SENS'
      DataSource = STa
    end
    object PR_NATUREGENE: THDBValComboBox
      Left = 200
      Top = 94
      Width = 112
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
      TabOrder = 6
      OnChange = PR_NATUREGENEChange
      TagDispatch = 0
      DataType = 'TTNATGENE'
      DataField = 'PR_NATUREGENE'
      DataSource = STa
    end
    object PR_COLLECTIF: TDBCheckBox
      Left = 3
      Top = 126
      Width = 65
      Height = 13
      Alignment = taLeftJustify
      Caption = 'C&ollectif'
      Ctl3D = True
      DataField = 'PR_COLLECTIF'
      DataSource = STa
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 7
      ValueChecked = 'X'
      ValueUnchecked = '-'
    end
    object PR_POINTABLE: TDBCheckBox
      Left = 123
      Top = 126
      Width = 65
      Height = 13
      Alignment = taLeftJustify
      Caption = '&Pointable'
      Ctl3D = True
      DataField = 'PR_POINTABLE'
      DataSource = STa
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 8
      ValueChecked = 'X'
      ValueUnchecked = '-'
    end
    object PR_LETTRABLE: TDBCheckBox
      Left = 250
      Top = 126
      Width = 60
      Height = 13
      Alignment = taLeftJustify
      Caption = '&Lettrable'
      Ctl3D = True
      DataField = 'PR_LETTRABLE'
      DataSource = STa
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 9
      ValueChecked = 'X'
      ValueUnchecked = '-'
    end
    object HGBOptionaxes: TGroupBox
      Left = 4
      Top = 147
      Width = 309
      Height = 38
      Caption = ' Ventilation '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 10
      object PR_VENTILABLE1: TDBCheckBox
        Left = 8
        Top = 14
        Width = 48
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Axe 1'
        Ctl3D = True
        DataField = 'PR_VENTILABLE1'
        DataSource = STa
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
        ValueChecked = 'X'
        ValueUnchecked = '-'
      end
      object PR_VENTILABLE2: TDBCheckBox
        Left = 68
        Top = 14
        Width = 48
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Axe 2'
        Ctl3D = True
        DataField = 'PR_VENTILABLE2'
        DataSource = STa
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
        ValueChecked = 'X'
        ValueUnchecked = '-'
      end
      object PR_VENTILABLE3: TDBCheckBox
        Left = 129
        Top = 14
        Width = 48
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Axe 3'
        Ctl3D = True
        DataField = 'PR_VENTILABLE3'
        DataSource = STa
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
        ValueChecked = 'X'
        ValueUnchecked = '-'
      end
      object PR_VENTILABLE4: TDBCheckBox
        Left = 189
        Top = 14
        Width = 48
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Axe 4'
        Ctl3D = True
        DataField = 'PR_VENTILABLE4'
        DataSource = STa
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 3
        ValueChecked = 'X'
        ValueUnchecked = '-'
      end
      object PR_VENTILABLE5: TDBCheckBox
        Left = 249
        Top = 14
        Width = 48
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Axe 5'
        Ctl3D = True
        DataField = 'PR_VENTILABLE5'
        DataSource = STa
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 4
        ValueChecked = 'X'
        ValueUnchecked = '-'
      end
    end
    object HGBOptionImpression: TGroupBox
      Left = 4
      Top = 191
      Width = 309
      Height = 58
      Caption = ' Options d'#39'impression du grand livre '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 11
      object PR_SOLDEPROGRESSIF: TDBCheckBox
        Left = 6
        Top = 17
        Width = 109
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Solde &progressif'
        Ctl3D = True
        DataField = 'PR_SOLDEPROGRESSIF'
        DataSource = STa
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 0
        ValueChecked = 'X'
        ValueUnchecked = '-'
      end
      object PR_TOTAUXMENSUELS: TDBCheckBox
        Left = 184
        Top = 17
        Width = 109
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Totaux &mensuels'
        Ctl3D = True
        DataField = 'PR_TOTAUXMENSUELS'
        DataSource = STa
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 1
        ValueChecked = 'X'
        ValueUnchecked = '-'
      end
      object PR_SAUTPAGE: TDBCheckBox
        Left = 6
        Top = 39
        Width = 109
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Sa&ut de page'
        Ctl3D = True
        DataField = 'PR_SAUTPAGE'
        DataSource = STa
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 2
        ValueChecked = 'X'
        ValueUnchecked = '-'
      end
      object PR_CENTRALISABLE: TDBCheckBox
        Left = 184
        Top = 39
        Width = 109
        Height = 13
        Alignment = taLeftJustify
        Caption = 'Centralisa&ble'
        Ctl3D = True
        DataField = 'PR_CENTRALISABLE'
        DataSource = STa
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        TabOrder = 3
        ValueChecked = 'X'
        ValueUnchecked = '-'
      end
    end
    object TG_BLOCNOTE: TGroupBox
      Left = 4
      Top = 253
      Width = 309
      Height = 68
      Caption = ' BlocNote '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 12
      object PR_BLOCNOTE: THDBRichEditOLE
        Left = 8
        Top = 16
        Width = 293
        Height = 43
        Ctl3D = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = []
        ParentCtl3D = False
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 0
        Margins.Top = 0
        Margins.Bottom = 0
        Margins.Left = 0
        Margins.Right = 0
        ContainerName = 'Document'
        ObjectMenuPrefix = '&Object'
        LinesRTF.Strings = (
          
            '{\rtf1\ansi\ansicpg1252\deff0\deflang1036{\fonttbl{\f0\fswiss Ar' +
            'ial;}}'
          '{\colortbl ;\red0\green0\blue0;}'
          
            '{\*\generator Msftedit 5.41.15.1507;}\viewkind4\uc1\pard\cf1\f0\' +
            'fs16 PR_BLOCNOTE'
          '\par }')
        DataField = 'PR_BLOCNOTE'
        DataSource = STa
      end
    end
    object Pnum: TPanel
      Left = 0
      Top = 0
      Width = 318
      Height = 35
      Align = alTop
      BevelInner = bvLowered
      TabOrder = 0
      object HLabel1: THLabel
        Left = 5
        Top = 10
        Width = 117
        Height = 13
        Caption = 'Choi&x du num'#233'ro de plan'
        FocusControl = FNumPlan
      end
      object FNumPlan: TSpinEdit
        Left = 127
        Top = 6
        Width = 43
        Height = 22
        AutoSize = False
        Ctl3D = True
        EditorEnabled = False
        MaxValue = 10
        MinValue = 1
        ParentCtl3D = False
        TabOrder = 0
        Value = 1
        OnChange = FNumPlanChange
      end
    end
    object PR_PREDEFINI: THDBValComboBox
      Left = 198
      Top = 6
      Width = 112
      Height = 21
      Style = csDropDownList
      Ctl3D = True
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 1
      OnChange = PR_NATUREGENEChange
      TagDispatch = 0
      DataType = 'YYPREDCEGSTD'
      DataField = 'PR_PREDEFINI'
      DataSource = STa
    end
  end
  inherited DBNav: TDBNavigator
    Left = 468
    Top = 64
    Hints.Strings = ()
  end
  inherited Dock971: TDock97
    Top = 322
    Width = 612
    inherited PBouton: TToolWindow97
      ClientWidth = 612
      ClientAreaWidth = 612
      inherited BImprimer: TToolbarButton97
        Left = 188
      end
      inherited BValider: TToolbarButton97
        Left = 220
      end
      inherited HelpBtn: TToolbarButton97
        Left = 284
      end
      inherited BFerme: TToolbarButton97
        Left = 252
      end
      inherited BFirst: TToolbarButton97
        Left = 324
        Width = 70
      end
      inherited BPrev: TToolbarButton97
        Left = 394
        Width = 70
      end
      inherited BNext: TToolbarButton97
        Left = 464
        Width = 70
      end
      inherited BLast: TToolbarButton97
        Left = 534
        Width = 70
      end
    end
  end
  object XX_WHERE: TPanel [4]
    Left = 412
    Top = 188
    Width = 81
    Height = 33
    Hint = 'Where PR_NUMPLAN=NumPlan'
    Caption = 'WHERE !!!'
    Color = clYellow
    TabOrder = 4
    Visible = False
  end
  inherited Ta: THTable
    IndexName = 'PR_CLE1'
    TableName = 'PLANREF'
    object TaPR_NUMPLAN: TIntegerField
      FieldName = 'PR_NUMPLAN'
    end
    object TaPR_COMPTE: TStringField
      FieldName = 'PR_COMPTE'
      Size = 17
    end
    object TaPR_LIBELLE: TStringField
      FieldName = 'PR_LIBELLE'
      Size = 35
    end
    object TaPR_ABREGE: TStringField
      FieldName = 'PR_ABREGE'
      Size = 17
    end
    object TaPR_NATUREGENE: TStringField
      FieldName = 'PR_NATUREGENE'
      Size = 3
    end
    object TaPR_CENTRALISABLE: TStringField
      FieldName = 'PR_CENTRALISABLE'
      Size = 1
    end
    object TaPR_SOLDEPROGRESSIF: TStringField
      FieldName = 'PR_SOLDEPROGRESSIF'
      Size = 1
    end
    object TaPR_SAUTPAGE: TStringField
      FieldName = 'PR_SAUTPAGE'
      Size = 1
    end
    object TaPR_TOTAUXMENSUELS: TStringField
      FieldName = 'PR_TOTAUXMENSUELS'
      Size = 1
    end
    object TaPR_COLLECTIF: TStringField
      FieldName = 'PR_COLLECTIF'
      Size = 1
    end
    object TaPR_BLOCNOTE: TMemoField
      FieldName = 'PR_BLOCNOTE'
      BlobType = ftMemo
      Size = 1
    end
    object TaPR_SENS: TStringField
      DisplayWidth = 1
      FieldName = 'PR_SENS'
      Size = 1
    end
    object TaPR_LETTRABLE: TStringField
      FieldName = 'PR_LETTRABLE'
      Size = 1
    end
    object TaPR_POINTABLE: TStringField
      FieldName = 'PR_POINTABLE'
      Size = 1
    end
    object TaPR_VENTILABLE1: TStringField
      FieldName = 'PR_VENTILABLE1'
      Size = 1
    end
    object TaPR_VENTILABLE2: TStringField
      FieldName = 'PR_VENTILABLE2'
      Size = 1
    end
    object TaPR_VENTILABLE3: TStringField
      FieldName = 'PR_VENTILABLE3'
      Size = 1
    end
    object TaPR_VENTILABLE4: TStringField
      FieldName = 'PR_VENTILABLE4'
      Size = 1
    end
    object TaPR_VENTILABLE5: TStringField
      FieldName = 'PR_VENTILABLE5'
      Size = 1
    end
    object TaPR_REPORTDETAIL: TStringField
      FieldName = 'PR_REPORTDETAIL'
      Size = 1
    end
    object TaPR_PREDEFINI: TStringField
      FieldName = 'PR_PREDEFINI'
      Size = 3
    end
  end
  inherited HM: THMsgBox
    Mess.Strings = (
      
        '0;Plans de r'#233'f'#233'rence;Voulez-vous enregistrer les modifications ?' +
        ';Q;YNC;Y;C;'
      
        '1;Plans de r'#233'f'#233'rence;Confirmez-vous la suppression de l'#39'enregist' +
        'rement ?;Q;YNC;N;C;'
      '2;Plans de r'#233'f'#233'rence;Vous devez renseigner un code.;W;O;O;O;'
      '3;Plans de r'#233'f'#233'rence;Vous devez renseigner un libell'#233'.;W;O;O;O;'
      
        '4;Plans de r'#233'f'#233'rence;Le code que vous avez saisi existe d'#233'j'#224'. Vo' +
        'us devez le modifier.;W;O;O;O;'
      'L'#39'enregistrement est inaccessible'
      'Plan de r'#233'f'#233'rence num'#233'ro ')
    Top = 100
  end
  inherited HM2: THMsgBox
    Top = 100
  end
  inherited HMTrad: THSystemMenu
    Left = 448
    Top = 104
  end
end

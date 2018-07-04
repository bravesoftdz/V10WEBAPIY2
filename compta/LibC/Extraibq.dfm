inherited FExtraibq: TFExtraibq
  Left = 327
  Top = 231
  HelpContext = 7604100
  Caption = 'R'#233'f'#233'rences de pointage :'
  ClientHeight = 312
  ClientWidth = 586
  PixelsPerInch = 96
  TextHeight = 13
  inherited FListe: THDBGrid
    Left = 322
    Width = 264
    Height = 270
    Columns = <
      item
        Expanded = False
        Title.Alignment = taCenter
        Title.Caption = 'Compte'
        Width = 80
        Visible = True
      end
      item
        Expanded = False
        Title.Alignment = taCenter
        Title.Caption = 'R'#233'f.pointage'
        Width = 80
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'EE_DATEPOINTAGE'
        Title.Alignment = taCenter
        Title.Caption = 'Date'
        Width = 69
        Visible = True
      end>
  end
  inherited Pappli: THPanel
    Width = 322
    Height = 270
    object TEE_GENERAL: THLabel
      Left = 8
      Top = 15
      Width = 60
      Height = 13
      Caption = '&Cpte g'#233'n'#233'ral'
    end
    object TEE_REFPOINTAGE: THLabel
      Left = 8
      Top = 43
      Width = 64
      Height = 13
      Caption = 'R'#233'&f. pointage'
      FocusControl = EE_REFPOINTAGE
    end
    object TEE_DATEPOINTAGE: THLabel
      Left = 8
      Top = 71
      Width = 67
      Height = 13
      Caption = '&Date pointage'
      FocusControl = EE_DATEPOINTAGE
    end
    object EE_REFPOINTAGE: TDBEdit
      Left = 77
      Top = 39
      Width = 205
      Height = 21
      CharCase = ecUpperCase
      DataField = 'EE_REFPOINTAGE'
      DataSource = STa
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 1
    end
    object EE_DATEPOINTAGE: TDBEdit
      Left = 77
      Top = 67
      Width = 80
      Height = 21
      DataField = 'EE_DATEPOINTAGE'
      DataSource = STa
      TabOrder = 2
    end
    object GbAncSol: TGroupBox
      Left = 8
      Top = 170
      Width = 285
      Height = 100
      Caption = ' Ancien solde bancaire  '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 4
      Visible = False
      object TEE_DATEOLDSOLDE: THLabel
        Left = 11
        Top = 21
        Width = 23
        Height = 13
        Caption = 'D&ate'
        FocusControl = EE_DATEOLDSOLDE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TEE_OLDSOLDECRE: THLabel
        Left = 11
        Top = 46
        Width = 27
        Height = 13
        Caption = 'Cr'#233'di&t'
        FocusControl = EE_OLDSOLDECRE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TEE_OLDSOLDEDEB: THLabel
        Left = 11
        Top = 71
        Width = 25
        Height = 13
        Caption = 'D'#233'bi&t'
        FocusControl = EE_OLDSOLDEDEB
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object EE_DATEOLDSOLDE: TDBEdit
        Left = 55
        Top = 17
        Width = 80
        Height = 21
        DataField = 'EE_DATEOLDSOLDE'
        DataSource = STa
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
      end
      object EE_OLDSOLDECRE: TDBEdit
        Left = 55
        Top = 42
        Width = 220
        Height = 21
        DataField = 'EE_OLDSOLDECRE'
        DataSource = STa
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
      end
      object EE_OLDSOLDEDEB: TDBEdit
        Left = 55
        Top = 67
        Width = 220
        Height = 21
        DataField = 'EE_OLDSOLDEDEB'
        DataSource = STa
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
      end
    end
    object GbNouSol: TGroupBox
      Left = 8
      Top = 96
      Width = 303
      Height = 69
      Caption = ' Nouveau solde bancaire '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 3
      object TEE_NEWSOLDECRE: THLabel
        Left = 11
        Top = 21
        Width = 27
        Height = 13
        Caption = 'C&r'#233'dit'
        FocusControl = EE_NEWSOLDECRE
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TEE_NEWSOLDEDEB: THLabel
        Left = 11
        Top = 46
        Width = 25
        Height = 13
        Caption = 'D'#233'&bit'
        FocusControl = EE_NEWSOLDEDEB
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object TLibDevise: TLabel
        Left = 192
        Top = 32
        Width = 64
        Height = 13
        Caption = 'Libell'#233' devise'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object EE_NEWSOLDECRE: TDBEdit
        Left = 55
        Top = 17
        Width = 130
        Height = 21
        DataField = 'EE_NEWSOLDECRE'
        DataSource = STa
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
      end
      object EE_NEWSOLDEDEB: TDBEdit
        Left = 55
        Top = 42
        Width = 130
        Height = 21
        DataField = 'EE_NEWSOLDEDEB'
        DataSource = STa
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
      end
    end
    object EE_GENERAL: THDBCpteEdit
      Left = 77
      Top = 11
      Width = 129
      Height = 21
      TabOrder = 0
      OnExit = EE_GENERALExit
      ZoomTable = tzGpointable
      Vide = False
      Bourre = True
      okLocate = False
      SynJoker = False
      DataField = 'EE_GENERAL'
      DataSource = STa
    end
    object GroupBox1: TGroupBox
      Left = 8
      Top = 184
      Width = 303
      Height = 69
      Caption = ' Nouveau solde bancaire '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clActiveCaption
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 5
      object HLabel1: THLabel
        Left = 11
        Top = 21
        Width = 27
        Height = 13
        Caption = 'C&r'#233'dit'
        FocusControl = EE_NEWSOLDECREEURO
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object HLabel2: THLabel
        Left = 11
        Top = 46
        Width = 25
        Height = 13
        Caption = 'D'#233'&bit'
        FocusControl = EE_NEWSOLDEDEBEURO
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object iEuroEuro: TImage
        Left = 229
        Top = 30
        Width = 16
        Height = 16
        AutoSize = True
        Picture.Data = {
          07544269746D6170F6000000424DF60000000000000076000000280000001000
          0000100000000100040000000000800000000000000000000000100000001000
          0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
          C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
          FF00FFFFFFFFFFFFFFFFFFFFFFF44444444FFFFFF44444444444FFFF4444FFFF
          FFF4FFFF444FFFFFFFFFFFFF44FFFFFFFFFFF44444444444FFFFFF4444444444
          4FFFFFF44FFFFFFFFFFFF444444444444FFFFF444444444444FFFFFF44FFFFFF
          FFF4FFFF444FFFFFFF44FFFFF444FFFFF444FFFFFF4444444444FFFFFFF44444
          4FF4}
        Stretch = True
        Transparent = True
      end
      object EE_NEWSOLDECREEURO: TDBEdit
        Left = 55
        Top = 17
        Width = 130
        Height = 21
        DataField = 'EE_NEWSOLDECREEURO'
        DataSource = STa
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
      end
      object EE_NEWSOLDEDEBEURO: TDBEdit
        Left = 55
        Top = 42
        Width = 130
        Height = 21
        DataField = 'EE_NEWSOLDEDEBEURO'
        DataSource = STa
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
      end
    end
  end
  inherited DBNav: TDBNavigator
    Left = 436
    Top = 104
    Hints.Strings = ()
  end
  inherited Dock971: TDock97
    Top = 270
    Width = 586
    inherited PBouton: TToolWindow97
      ClientWidth = 586
      ClientAreaWidth = 586
      inherited BImprimer: TToolbarButton97
        Left = 169
      end
      inherited BValider: TToolbarButton97
        Left = 201
      end
      inherited HelpBtn: TToolbarButton97
        Left = 265
      end
      inherited BFerme: TToolbarButton97
        Left = 233
      end
      inherited BFirst: TToolbarButton97
        Left = 304
        Width = 61
      end
      inherited BPrev: TToolbarButton97
        Left = 365
        Width = 61
      end
      inherited BNext: TToolbarButton97
        Left = 426
        Width = 61
      end
      inherited BLast: TToolbarButton97
        Left = 487
        Width = 61
      end
      inherited FAutoSave: THCheckbox
        Left = 101
        Width = 26
      end
    end
  end
  inherited Ta: THTable
    TableName = 'EEXBQ'
    object TaEE_GENERAL: TStringField
      FieldName = 'EE_GENERAL'
      Size = 17
    end
    object TaEE_REFPOINTAGE: TStringField
      FieldName = 'EE_REFPOINTAGE'
      Size = 17
    end
    object TaEE_DATEOLDSOLDE: TDateTimeField
      FieldName = 'EE_DATEOLDSOLDE'
      DisplayFormat = 'dd mmm yyyy'
      EditMask = '!99/99/0000;1;_'
    end
    object TaEE_OLDSOLDECRE: TFloatField
      FieldName = 'EE_OLDSOLDECRE'
    end
    object TaEE_DATEPOINTAGE: TDateTimeField
      FieldName = 'EE_DATEPOINTAGE'
      DisplayFormat = 'dd mmm yyyy'
      EditMask = '!99/99/0000;1;_'
    end
    object TaEE_NEWSOLDECRE: TFloatField
      FieldName = 'EE_NEWSOLDECRE'
    end
    object TaEE_SOCIETE: TStringField
      FieldName = 'EE_SOCIETE'
      Size = 3
    end
    object TaEE_OLDSOLDEDEB: TFloatField
      FieldName = 'EE_OLDSOLDEDEB'
    end
    object TaEE_NEWSOLDEDEB: TFloatField
      FieldName = 'EE_NEWSOLDEDEB'
    end
    object TaEE_OLDSOLDEDEBEURO: TFloatField
      FieldName = 'EE_OLDSOLDEDEBEURO'
    end
    object TaEE_OLDSOLDECREEURO: TFloatField
      FieldName = 'EE_OLDSOLDECREEURO'
    end
    object TaEE_NEWSOLDEDEBEURO: TFloatField
      FieldName = 'EE_NEWSOLDEDEBEURO'
    end
    object TaEE_NEWSOLDECREEURO: TFloatField
      FieldName = 'EE_NEWSOLDECREEURO'
    end
    object TaEE_NUMERO: TIntegerField
      FieldName = 'EE_NUMERO'
    end
  end
  inherited HM: THMsgBox
    Mess.Strings = (
      
        '0;R'#233'f'#233'rences de pointage ;Voulez-vous enregistrer les modificati' +
        'ons?;Q;YNC;Y;C;'
      
        '1;R'#233'f'#233'rences de pointage ;Confirmez-vous la suppression de l'#39'enr' +
        'egistrement?;Q;YNC;N;C;'
      '2;R'#233'f'#233'rences de pointage ;Vous devez renseigner un code.W;O;O;O;'
      
        '3;R'#233'f'#233'rences de pointage ;Vous devez renseigner un libell'#233'.;W;O;' +
        'O;O;'
      
        '4;R'#233'f'#233'rences de pointage ;Le code que vous avez saisi existe d'#233'j' +
        #224'. Vous devez le modifier.;W;O;O;O;'
      'L'#39'enregistrement est inaccessible'
      
        '6;R'#233'f'#233'rences de pointage ;Le compte que vous avez saisi n'#39'est pa' +
        's un compte pointable.;W;O;O;O;'
      
        '7;R'#233'f'#233'rences de pointage ;Suppression impossible : Des '#233'critures' +
        ' sont point'#233'es avec cette r'#233'f'#233'rence.;W;O;O;O;'
      
        '8;R'#233'f'#233'rences de pointage ;Attention : Ce compte n'#39'a pas de devis' +
        'e associ'#233'e;W;O;O;O;'
      
        '9;R'#233'f'#233'rences de pointage ;Certaines lignes du relev'#233' ne sont pas' +
        ' point'#233'es. Confirmez-vous la suppression de l'#39'enregistrement?;Q;' +
        'YNC;N;C;')
    Top = 96
  end
  inherited HMTrad: THSystemMenu
    Left = 432
    Top = 48
  end
end

object FChancel: TFChancel
  Left = 444
  Top = 245
  HelpContext = 1150100
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Chancellerie : '
  ClientHeight = 360
  ClientWidth = 507
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clBlack
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PDevise: TPanel
    Left = 0
    Top = 0
    Width = 507
    Height = 111
    Align = alTop
    TabOrder = 0
    object TH_DATEDU: THLabel
      Left = 7
      Top = 42
      Width = 48
      Height = 13
      Caption = '&A partir du'
      FocusControl = H_DATEDU
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TH_DATEAU: THLabel
      Left = 158
      Top = 42
      Width = 39
      Height = 13
      Caption = '&jusqu'#39'au'
      FocusControl = H_DATEAU
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TQuotite: THLabel
      Left = 291
      Top = 84
      Width = 82
      Height = 13
      Caption = 'La quotit'#233' est de '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TValeurQ: THLabel
      Left = 440
      Top = 80
      Width = 40
      Height = 13
      AutoSize = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TH_DEVISE: THLabel
      Left = 7
      Top = 13
      Width = 33
      Height = 13
      Caption = '&Devise'
      FocusControl = CodeDevise
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TValable: THLabel
      Left = 291
      Top = 42
      Width = 205
      Height = 35
      AutoSize = False
      Caption = 'Valable jusqu'#39'au'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      WordWrap = True
    end
    object H_DATEDU: TMaskEdit
      Left = 62
      Top = 38
      Width = 69
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 10
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 2
      Text = '01/01/1900'
      OnExit = H_DATEDUExit
    end
    object H_DATEAU: TMaskEdit
      Left = 206
      Top = 38
      Width = 69
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 10
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 3
      Text = '31/12/2099'
      OnExit = H_DATEAUExit
    end
    object OkDate: TCheckBox
      Left = 291
      Top = 12
      Width = 73
      Height = 17
      Alignment = taLeftJustify
      Caption = '&Filtre date'
      Checked = True
      Ctl3D = True
      ParentCtl3D = False
      State = cbChecked
      TabOrder = 1
      OnClick = OkDateClick
    end
    object CodeDevise: THValComboBox
      Left = 62
      Top = 9
      Width = 212
      Height = 21
      Style = csDropDownList
      Ctl3D = True
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 0
      OnChange = CodeDeviseChange
      TagDispatch = 0
    end
    object RGSens: TRadioGroup
      Left = 7
      Top = 64
      Width = 262
      Height = 36
      Caption = 'Sens de saisie'
      Columns = 2
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemIndex = 0
      Items.Strings = (
        '1 <FFF> = xx,xxx $'
        '1 $ = xx,xxx <FFF>')
      ParentFont = False
      TabOrder = 4
      OnClick = RGSensClick
    end
  end
  object panel: TPanel
    Left = 0
    Top = 325
    Width = 507
    Height = 35
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 2
    object FAutoSave: TCheckBox
      Left = 15
      Top = 8
      Width = 61
      Height = 17
      Caption = 'Auto'
      TabOrder = 0
      Visible = False
    end
    object BInsert: THBitBtn
      Left = 152
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Nouveau'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 6
      OnClick = BInsertClick
      Margin = 2
      GlobalIndexImage = 'Z0053_S16G1'
    end
    object HelpBtn: THBitBtn
      Left = 473
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Aide'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 12
      OnClick = HelpBtnClick
      Glyph.Data = {
        BE060000424DBE06000000000000360400002800000024000000120000000100
        0800000000008802000000000000000000000001000000010000000000000000
        80000080000000808000800000008000800080800000C0C0C000C0DCC000F0CA
        A600000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        000000000000000000000000000000000000F0FBFF00A4A0A000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00030303030303
        03030606030303030303030303030303030303FFFF0303030303030303030303
        0303030303060404060303030303030303030303030303F8F8FF030303030303
        030303030303030303FE06060403030303030303030303030303F8FF03F8FF03
        0303030303030303030303030303FE060603030303030303030303030303F8FF
        FFF8FF0303030303030303030303030303030303030303030303030303030303
        030303F8F8030303030303030303030303030303030304040603030303030303
        0303030303030303FFFF03030303030303030303030303030306060604030303
        0303030303030303030303F8F8F8FF0303030303030303030303030303FE0606
        0403030303030303030303030303F8FF03F8FF03030303030303030303030303
        03FE06060604030303030303030303030303F8FF03F8FF030303030303030303
        030303030303FE060606040303030303030303030303F8FF0303F8FF03030303
        0303030303030303030303FE060606040303030303030303030303F8FF0303F8
        FF030303030303030303030404030303FE060606040303030303030303FF0303
        F8FF0303F8FF030303030303030306060604030303FE06060403030303030303
        F8F8FF0303F8FF0303F8FF03030303030303FE06060604040406060604030303
        030303F8FF03F8FFFFFFF80303F8FF0303030303030303FE0606060606060606
        06030303030303F8FF0303F8F8F8030303F8FF030303030303030303FEFE0606
        060606060303030303030303F8FFFF030303030303F803030303030303030303
        0303FEFEFEFEFE03030303030303030303F8F8FFFFFFFFFFF803030303030303
        0303030303030303030303030303030303030303030303F8F8F8F8F803030303
        0303}
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object BPurge: THBitBtn
      Left = 216
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Purger'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 8
      OnClick = BPurgeClick
      Margin = 2
      GlobalIndexImage = 'Z0204_S16G1'
    end
    object BFirst: THBitBtn
      Left = 4
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Premier'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnClick = BFirstClick
      GlobalIndexImage = 'Z0301_S16G1'
    end
    object BPrev: THBitBtn
      Left = 32
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Pr'#233'c'#233'dent'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = BPrevClick
      GlobalIndexImage = 'Z0017_S16G1'
    end
    object BNext: THBitBtn
      Left = 60
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Suivant'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      OnClick = BNextClick
      GlobalIndexImage = 'Z0031_S16G1'
    end
    object BLast: THBitBtn
      Left = 88
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Dernier'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 4
      OnClick = BLastClick
      GlobalIndexImage = 'Z0264_S16G1'
    end
    object BFerme: THBitBtn
      Left = 441
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      ParentShowHint = False
      ShowHint = True
      TabOrder = 11
      OnClick = BFermeClick
      GlobalIndexImage = 'Z1770_S16G1'
    end
    object BDelete: THBitBtn
      Left = 184
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Supprimer'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 7
      OnClick = BDeleteClick
      Margin = 2
      GlobalIndexImage = 'Z0005_S16G1'
    end
    object BImprimer: THBitBtn
      Left = 379
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Imprimer'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 9
      OnClick = BImprimerClick
      Margin = 2
      GlobalIndexImage = 'Z0369_S16G1'
    end
    object BValider: THBitBtn
      Left = 410
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Valider'
      Default = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 10
      OnClick = BValiderClick
      Glyph.Data = {
        BE060000424DBE06000000000000360400002800000024000000120000000100
        0800000000008802000000000000000000000001000000010000000000000000
        80000080000000808000800000008000800080800000C0C0C000C0DCC000F0C8
        A400000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        000000000000000000000000000000000000F0FBFF00A4A0A000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00030303030303
        0303030303030303030303030303030303030303030303030303030303030303
        03030303030303030303030303030303030303030303FF030303030303030303
        03030303030303040403030303030303030303030303030303F8F8FF03030303
        03030303030303030303040202040303030303030303030303030303F80303F8
        FF030303030303030303030303040202020204030303030303030303030303F8
        03030303F8FF0303030303030303030304020202020202040303030303030303
        0303F8030303030303F8FF030303030303030304020202FA0202020204030303
        0303030303F8FF0303F8FF030303F8FF03030303030303020202FA03FA020202
        040303030303030303F8FF03F803F8FF0303F8FF03030303030303FA02FA0303
        03FA0202020403030303030303F8FFF8030303F8FF0303F8FF03030303030303
        FA0303030303FA0202020403030303030303F80303030303F8FF0303F8FF0303
        0303030303030303030303FA0202020403030303030303030303030303F8FF03
        03F8FF03030303030303030303030303FA020202040303030303030303030303
        0303F8FF0303F8FF03030303030303030303030303FA02020204030303030303
        03030303030303F8FF0303F8FF03030303030303030303030303FA0202020403
        030303030303030303030303F8FF0303F8FF03030303030303030303030303FA
        0202040303030303030303030303030303F8FF03F8FF03030303030303030303
        03030303FA0202030303030303030303030303030303F8FFF803030303030303
        030303030303030303FA0303030303030303030303030303030303F803030303
        0303030303030303030303030303030303030303030303030303030303030303
        0303}
      Margin = 2
      NumGlyphs = 2
      Spacing = -1
      IsControl = True
    end
    object BAnnuler: THBitBtn
      Left = 120
      Top = 4
      Width = 28
      Height = 27
      Hint = 'D'#233'faire'
      Cancel = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 5
      OnClick = BAnnulerClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0075_S16G1'
      IsControl = True
    end
  end
  object FListe: THDBGrid
    Left = 0
    Top = 111
    Width = 507
    Height = 214
    Align = alClient
    Ctl3D = False
    DataSource = SChancell
    Options = [dgEditing, dgAlwaysShowEditor, dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs, dgAlwaysShowSelection]
    ParentCtl3D = False
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clBlack
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnKeyDown = FListeKeyDown
    OnRowEnter = FListeRowEnter
    Row = 1
    MultiSelection = False
    SortEnabled = False
    MyDefaultRowHeight = 0
    Columns = <
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'H_DATECOURS'
        Title.Alignment = taCenter
        Title.Caption = 'Date'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'H_COTATION'
        Title.Alignment = taCenter
        Title.Caption = 'Cotation'
        Width = 100
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'H_TAUXREEL'
        Title.Alignment = taCenter
        Title.Caption = 'Taux r'#233'el'
        Width = 100
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'H_TAUXCLOTURE'
        Title.Alignment = taRightJustify
        Title.Caption = 'Taux de cloture'
        Width = 79
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'H_COMMENTAIRE'
        Title.Alignment = taCenter
        Title.Caption = 'Commentaire'
        Width = 166
        Visible = True
      end>
  end
  object DBNav: TDBNavigator
    Left = 88
    Top = 258
    Width = 88
    Height = 18
    DataSource = SChancell
    VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast]
    Ctl3D = True
    ParentCtl3D = False
    ConfirmDelete = False
    TabOrder = 3
    Visible = False
  end
  object SChancell: TDataSource
    DataSet = TChancell
    OnDataChange = SChancellDataChange
    OnUpdateData = SChancellUpdateData
    Left = 96
    Top = 204
  end
  object TChancell: THTable
    MarshalOptions = moMarshalModifiedOnly
    BeforePost = TChancellBeforePost
    OnNewRecord = TChancellNewRecord
    OnPostError = TChancellPostError
    EnableBCD = False
    IndexName = 'H_CLE1'
    TableName = 'CHANCELL'
    dataBaseName = 'SOC'
    UpdateMode = upWhereChanged
    RequestLive = True
    Left = 38
    Top = 204
    object TChancellH_DEVISE: TStringField
      FieldName = 'H_DEVISE'
      Size = 3
    end
    object TChancellH_DATECOURS: TDateTimeField
      FieldName = 'H_DATECOURS'
      EditMask = '!99/99/0000;1;_'
    end
    object TChancellH_TAUXREEL: TFloatField
      FieldName = 'H_TAUXREEL'
    end
    object TChancellH_TAUXCLOTURE: TFloatField
      FieldName = 'H_TAUXCLOTURE'
    end
    object TChancellH_COMMENTAIRE: TStringField
      FieldName = 'H_COMMENTAIRE'
      Size = 17
    end
    object TChancellH_COTATION: TFloatField
      FieldName = 'H_COTATION'
    end
  end
  object MsgBox: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Chancellerie;Voulez-vous enregistrer les modifications?;Q;YNC;' +
        'Y;C;'
      
        '1;Chancellerie;Confirmez-vous la suppression de l'#39'enregistrement' +
        '?;Q;YNC;N;C;'
      '2;Chancellerie;Vous devez renseigner un code.;W;O;O;O;'
      '3;Chancellerie;Vous devez renseigner un libell'#233'.;W;O;O;O;'
      
        '4;Chancellerie;La date de valeur que vous avez saisie est d'#233'j'#224' r' +
        'enseign'#233'e.;W;O;O;O;'
      
        '5;Chancellerie;Vous devez renseigner une date de valeur.;W;O;O;O' +
        ';'
      
        '6;Chancellerie;Vous devez renseigner un taux de change pour chaq' +
        'ue date de valeur.;W;O;O;O;'
      'L'#39'enregistrement est inaccessible'
      
        '8;Chancellerie;L'#39'enregistrement n'#39'a pas pu '#234'tre sauvegard'#233'. La d' +
        'ate est vide ou elle existe d'#233'j'#224'.;W;O;O;O;'
      
        '9;Chancellerie;Vous devez renseigner un taux de change sup'#233'rieur' +
        ' '#224' z'#233'ro.;W;O;O;O;'
      'Valable jusqu'#39'au'
      'Valable '#224' partir du'
      'par rapport '#224' la devise pivot'
      'par rapport '#224' l '#39' Euro'
      
        '14;Chancellerie;La date saisie est sup'#233'rieure '#224' la date d'#39'entr'#233'e' +
        ' en vigueur de l'#39'Euro.;W;O;O;O;'
      
        '15;Chancellerie;La date saisie est inf'#233'rieure '#224' la date d'#39'entr'#233'e' +
        ' en vigueur de l'#39'Euro.;W;O;O;O;'
      '1 <FFF> = xx,xxx <$$$>'
      '1 <$$$> = xx,xxx <FFF>')
    Left = 148
    Top = 204
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 200
    Top = 204
  end
end

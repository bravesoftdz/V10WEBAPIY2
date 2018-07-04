object FMce: TFMce
  Left = 353
  Top = 207
  Width = 454
  Height = 415
  HelpContext = 1732000
  BorderIcons = [biSystemMenu]
  Caption = 'Communication Etebac'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  ShowHint = True
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Dock: TDock97
    Left = 0
    Top = 341
    Width = 438
    Height = 35
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      ClientHeight = 31
      ClientWidth = 442
      ClientAreaHeight = 31
      ClientAreaWidth = 442
      DockPos = 0
      TabOrder = 0
      object BAide: TToolbarButton97
        Left = 412
        Top = 2
        Width = 27
        Height = 28
        Hint = 'Aide'
        Caption = 'Aide'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
      end
      object BAnnuler: TToolbarButton97
        Left = 382
        Top = 2
        Width = 27
        Height = 28
        Hint = 'Annuler'
        Caption = 'Annuler'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BAnnulerClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BValider: TToolbarButton97
        Left = 350
        Top = 2
        Width = 28
        Height = 28
        Hint = 'Valider'
        Caption = 'Valider'
        Default = True
        DisplayMode = dmGlyphOnly
        Flat = False
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0003_S16G2'
      end
      object BConnect: TToolbarButton97
        Left = 96
        Top = 2
        Width = 27
        Height = 28
        Hint = 'Connexions'
        Caption = 'Connexions'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        Visible = False
        OnClick = BConnectClick
        GlobalIndexImage = 'Z0050_S16G1'
      end
      object BNouveau: TToolbarButton97
        Left = 5
        Top = 2
        Width = 27
        Height = 28
        Hint = 'Nouveau'
        Caption = 'Nouveau'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BNouveauClick
        GlobalIndexImage = 'Z0053_S16G1'
      end
      object BSuppr: TToolbarButton97
        Left = 34
        Top = 2
        Width = 27
        Height = 28
        Hint = 'Supprimer'
        Caption = 'Supprimer'
        DisplayMode = dmGlyphOnly
        Flat = False
        ParentShowHint = False
        ShowHint = True
        OnClick = BSupprClick
        GlobalIndexImage = 'Z0005_S16G1'
      end
    end
  end
  object PageControl: TPageControl
    Left = 0
    Top = 0
    Width = 438
    Height = 341
    ActivePage = TabSheet1
    Align = alClient
    TabOrder = 1
    OnChange = PageControlChange
    OnChanging = PageControlChanging
    object TabSheet1: TTabSheet
      Caption = 'Destinataires'
      object Bevel1: TBevel
        Left = 0
        Top = 1
        Width = 440
        Height = 326
      end
      object LBDest: TListBox
        Left = 328
        Top = 9
        Width = 99
        Height = 303
        ItemHeight = 13
        Sorted = True
        TabOrder = 0
        OnClick = LBDestClick
      end
      object GroupBox1: TGroupBox
        Left = 7
        Top = 7
        Width = 306
        Height = 72
        Caption = 'Destinataire'
        TabOrder = 1
        object HLabel1: THLabel
          Left = 11
          Top = 23
          Width = 25
          Height = 13
          Caption = '&Code'
          FocusControl = ECode
        end
        object HLabel2: THLabel
          Left = 163
          Top = 23
          Width = 37
          Height = 13
          Caption = '&Guichet'
          FocusControl = EGuichet
        end
        object HLabel3: THLabel
          Left = 11
          Top = 49
          Width = 22
          Height = 13
          Caption = '&Nom'
          FocusControl = ENomDest
        end
        object ECode: TEdit
          Left = 68
          Top = 20
          Width = 66
          Height = 21
          CharCase = ecUpperCase
          MaxLength = 5
          TabOrder = 0
          OnChange = OnChangeCtl
        end
        object EGuichet: TEdit
          Left = 221
          Top = 20
          Width = 76
          Height = 21
          MaxLength = 5
          TabOrder = 1
          OnChange = OnChangeCtl
        end
        object ENomDest: TEdit
          Left = 68
          Top = 46
          Width = 229
          Height = 21
          MaxLength = 70
          TabOrder = 2
          OnChange = OnChangeCtl
        end
      end
      object GroupBox2: TGroupBox
        Left = 7
        Top = 83
        Width = 306
        Height = 92
        Caption = 'Transpac'
        TabOrder = 2
        object HLabel4: THLabel
          Left = 11
          Top = 20
          Width = 38
          Height = 13
          Caption = 'A&dresse'
          FocusControl = EAdr
        end
        object HLabel5: THLabel
          Left = 11
          Top = 49
          Width = 46
          Height = 13
          Caption = '&Extension'
          FocusControl = EExt
        end
        object EAdr: TEdit
          Left = 68
          Top = 16
          Width = 229
          Height = 21
          MaxLength = 40
          TabOrder = 0
          OnChange = OnChangeCtl
        end
        object EExt: TEdit
          Left = 68
          Top = 46
          Width = 229
          Height = 21
          MaxLength = 40
          TabOrder = 1
          OnChange = OnChangeCtl
        end
        object CPcv: TCheckBox
          Left = 68
          Top = 72
          Width = 129
          Height = 13
          Alignment = taLeftJustify
          Caption = 'N'#39'accepte pas le &PCV'
          TabOrder = 2
          OnClick = OnChangeCtl
        end
      end
      object GroupBox3: TGroupBox
        Left = 7
        Top = 176
        Width = 306
        Height = 73
        Caption = 'Contact'
        TabOrder = 3
        object HLabel6: THLabel
          Left = 11
          Top = 19
          Width = 22
          Height = 13
          Caption = 'N&om'
          FocusControl = ENomContact
        end
        object HLabel7: THLabel
          Left = 11
          Top = 46
          Width = 51
          Height = 13
          Caption = '&T'#233'l'#233'phone'
          FocusControl = ETelContact
        end
        object HLabel8: THLabel
          Left = 173
          Top = 46
          Width = 17
          Height = 13
          Caption = '&Fax'
          FocusControl = EFaxContact
        end
        object ENomContact: TEdit
          Left = 68
          Top = 15
          Width = 229
          Height = 21
          MaxLength = 40
          TabOrder = 0
          OnChange = OnChangeCtl
        end
        object ETelContact: TEdit
          Left = 68
          Top = 43
          Width = 99
          Height = 21
          MaxLength = 20
          TabOrder = 1
          OnChange = OnChangeCtl
        end
        object EFaxContact: TEdit
          Left = 198
          Top = 43
          Width = 99
          Height = 21
          MaxLength = 20
          TabOrder = 2
          OnChange = OnChangeCtl
        end
      end
      object GroupBox4: TGroupBox
        Left = 7
        Top = 250
        Width = 306
        Height = 73
        Caption = 'Contact technique'
        TabOrder = 4
        object HLabel9: THLabel
          Left = 11
          Top = 19
          Width = 22
          Height = 13
          Caption = 'No&m'
          FocusControl = ENomTech
        end
        object HLabel10: THLabel
          Left = 11
          Top = 46
          Width = 51
          Height = 13
          Caption = 'T'#233'l'#233'p&hone'
          FocusControl = ETelTech
        end
        object HLabel11: THLabel
          Left = 173
          Top = 46
          Width = 17
          Height = 13
          Caption = 'F&ax'
          FocusControl = EFaxTech
        end
        object ENomTech: TEdit
          Left = 68
          Top = 15
          Width = 229
          Height = 21
          MaxLength = 40
          TabOrder = 0
          OnChange = OnChangeCtl
        end
        object ETelTech: TEdit
          Left = 68
          Top = 43
          Width = 99
          Height = 21
          MaxLength = 20
          TabOrder = 1
          OnChange = OnChangeCtl
        end
        object EFaxTech: TEdit
          Left = 198
          Top = 43
          Width = 99
          Height = 21
          MaxLength = 20
          TabOrder = 2
          OnChange = OnChangeCtl
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'Cartes d'#39'appel'
      object Bevel2: TBevel
        Left = 0
        Top = 2
        Width = 440
        Height = 325
      end
      object GroupBox5: TGroupBox
        Left = 10
        Top = 13
        Width = 150
        Height = 170
        Caption = 'Destinataires'
        TabOrder = 0
        object LBDest2: TListBox
          Left = 10
          Top = 20
          Width = 131
          Height = 140
          ItemHeight = 13
          Sorted = True
          TabOrder = 0
          OnClick = LBDest2Click
        end
      end
      object GroupBox6: TGroupBox
        Left = 176
        Top = 13
        Width = 254
        Height = 170
        Caption = 'Mod'#232'les'
        TabOrder = 1
        object LBModele: TListBox
          Left = 10
          Top = 20
          Width = 235
          Height = 140
          ItemHeight = 13
          TabOrder = 0
          OnClick = LBModeleClick
          OnDblClick = LBModeleDblClick
        end
      end
      object GroupBox7: TGroupBox
        Left = 10
        Top = 195
        Width = 421
        Height = 108
        Caption = 'Carte d'#39'appel'
        TabOrder = 2
        object Label1: TLabel
          Left = 135
          Top = 20
          Width = 96
          Height = 13
          Caption = 'Position du curseur :'
        end
        object LPos: TLabel
          Left = 245
          Top = 20
          Width = 3
          Height = 13
        end
        object HLabel12: THLabel
          Left = 55
          Top = 78
          Width = 299
          Height = 13
          Caption = 'Pour modifier la carte d'#39'appel, double-cliquez le nom du mod'#232'le.'
        end
        object ECarte: TEdit
          Left = 13
          Top = 46
          Width = 401
          Height = 21
          MaxLength = 80
          TabOrder = 0
        end
      end
    end
    object TabSheet3: TTabSheet
      Caption = 'Port S'#233'rie'
      object Bevel3: TBevel
        Left = 0
        Top = 2
        Width = 440
        Height = 325
      end
      object GroupBox8: TGroupBox
        Left = 10
        Top = 10
        Width = 416
        Height = 69
        Caption = 'Port de communication'
        TabOrder = 0
        object RBCom1: TRadioButton
          Left = 16
          Top = 20
          Width = 76
          Height = 13
          Caption = 'COM1'
          TabOrder = 0
        end
        object RBCom2: TRadioButton
          Left = 120
          Top = 20
          Width = 73
          Height = 13
          Caption = 'COM2'
          TabOrder = 1
        end
        object RBCom3: TRadioButton
          Left = 225
          Top = 20
          Width = 72
          Height = 13
          Caption = 'COM3'
          TabOrder = 2
        end
        object RBCom4: TRadioButton
          Left = 329
          Top = 20
          Width = 72
          Height = 13
          Caption = 'COM4'
          TabOrder = 3
        end
        object RBCom5: TRadioButton
          Left = 16
          Top = 46
          Width = 73
          Height = 13
          Caption = 'COM5'
          TabOrder = 4
        end
        object RBCom6: TRadioButton
          Left = 120
          Top = 46
          Width = 76
          Height = 13
          Caption = 'COM6'
          TabOrder = 5
        end
        object RBCom7: TRadioButton
          Left = 225
          Top = 46
          Width = 75
          Height = 13
          Caption = 'COM7'
          TabOrder = 6
        end
        object RBCom8: TRadioButton
          Left = 329
          Top = 46
          Width = 65
          Height = 13
          Caption = 'COM8'
          TabOrder = 7
        end
      end
      object GroupBox9: TGroupBox
        Left = 10
        Top = 87
        Width = 416
        Height = 69
        Caption = 'Vitesse de communication'
        TabOrder = 1
        object RB300: TRadioButton
          Left = 62
          Top = 20
          Width = 59
          Height = 13
          Caption = '300'
          TabOrder = 0
        end
        object RB9600: TRadioButton
          Left = 185
          Top = 20
          Width = 69
          Height = 13
          Caption = '9600'
          TabOrder = 1
        end
        object RB19200: TRadioButton
          Left = 312
          Top = 20
          Width = 63
          Height = 13
          Caption = '19200'
          TabOrder = 2
        end
        object RB38400: TRadioButton
          Left = 62
          Top = 46
          Width = 69
          Height = 14
          Caption = '38400'
          TabOrder = 3
        end
        object RB57600: TRadioButton
          Left = 189
          Top = 46
          Width = 75
          Height = 14
          Caption = '57600'
          TabOrder = 4
        end
        object RB115200: TRadioButton
          Left = 312
          Top = 46
          Width = 69
          Height = 13
          Caption = '115200'
          TabOrder = 5
        end
      end
      object GroupBox10: TGroupBox
        Left = 10
        Top = 165
        Width = 416
        Height = 49
        Caption = 'Nombre de bits'
        TabOrder = 2
        object RB5b: TRadioButton
          Left = 16
          Top = 23
          Width = 66
          Height = 14
          Caption = '5 bits'
          TabOrder = 0
        end
        object RB6b: TRadioButton
          Left = 120
          Top = 23
          Width = 60
          Height = 14
          Caption = '6 bits'
          TabOrder = 1
        end
        object RB7b: TRadioButton
          Left = 225
          Top = 23
          Width = 92
          Height = 14
          Caption = '7 bits'
          TabOrder = 2
        end
        object RB8b: TRadioButton
          Left = 329
          Top = 23
          Width = 62
          Height = 14
          Caption = '8 bits'
          TabOrder = 3
        end
      end
      object GroupBox11: TGroupBox
        Left = 10
        Top = 219
        Width = 416
        Height = 48
        Caption = 'Parit'#233
        TabOrder = 3
        object RBSans: TRadioButton
          Left = 62
          Top = 21
          Width = 69
          Height = 14
          Caption = 'Sans'
          TabOrder = 0
        end
        object RBPaire: TRadioButton
          Left = 189
          Top = 21
          Width = 69
          Height = 14
          Caption = 'Paire'
          TabOrder = 1
        end
        object RBImpaire: TRadioButton
          Left = 312
          Top = 21
          Width = 79
          Height = 14
          Caption = 'Impaire'
          TabOrder = 2
        end
      end
      object GroupBox12: TGroupBox
        Left = 10
        Top = 271
        Width = 416
        Height = 48
        Caption = 'Nombre de bit(s) stop'
        TabOrder = 4
        object RB1Stop: TRadioButton
          Left = 104
          Top = 23
          Width = 82
          Height = 14
          Caption = '1 bit stop'
          TabOrder = 0
        end
        object RB2Stop: TRadioButton
          Left = 270
          Top = 23
          Width = 92
          Height = 14
          Caption = '2 bits stop'
          TabOrder = 1
        end
      end
    end
    object TabSheet4: TTabSheet
      Caption = 'Modem'
      object Bevel4: TBevel
        Left = 0
        Top = 2
        Width = 440
        Height = 325
      end
      object HLabel13: THLabel
        Left = 16
        Top = 59
        Width = 80
        Height = 13
        Caption = 'Reconnaissance'
      end
      object HLabel14: THLabel
        Left = 120
        Top = 23
        Width = 53
        Height = 13
        Caption = 'Commande'
      end
      object HLabel15: THLabel
        Left = 16
        Top = 101
        Width = 63
        Height = 13
        Caption = 'Num'#233'rotation'
      end
      object HLabel16: THLabel
        Left = 16
        Top = 140
        Width = 67
        Height = 13
        Caption = 'Remise '#224' z'#233'ro'
      end
      object HLabel17: THLabel
        Left = 16
        Top = 178
        Width = 55
        Height = 13
        Caption = 'Autorisation'
      end
      object HLabel18: THLabel
        Left = 16
        Top = 196
        Width = 62
        Height = 13
        Caption = 'appel entrant'
      end
      object HLabel19: THLabel
        Left = 16
        Top = 236
        Width = 54
        Height = 13
        Caption = 'Initialisation'
      end
      object HLabel20: THLabel
        Left = 16
        Top = 273
        Width = 46
        Height = 13
        Caption = 'D'#233'tection'
      end
      object HLabel21: THLabel
        Left = 16
        Top = 288
        Width = 62
        Height = 13
        Caption = 'appel entrant'
      end
      object HLabel22: THLabel
        Left = 206
        Top = 23
        Width = 48
        Height = 13
        Caption = 'Valeur OK'
      end
      object HLabel23: THLabel
        Left = 284
        Top = 23
        Width = 63
        Height = 13
        Caption = 'Valeur '#233'chec'
      end
      object HLabel24: THLabel
        Left = 365
        Top = 23
        Width = 60
        Height = 13
        Caption = 'D'#233'lai attente'
      end
      object ECRec: TEdit
        Left = 114
        Top = 55
        Width = 66
        Height = 21
        TabOrder = 0
      end
      object ECNum: TEdit
        Left = 114
        Top = 98
        Width = 66
        Height = 21
        TabOrder = 1
      end
      object ECRaz: TEdit
        Left = 114
        Top = 140
        Width = 66
        Height = 21
        TabOrder = 2
      end
      object ECAut: TEdit
        Left = 114
        Top = 183
        Width = 66
        Height = 21
        TabOrder = 3
      end
      object ECIni: TEdit
        Left = 114
        Top = 232
        Width = 66
        Height = 21
        TabOrder = 4
      end
      object EORec: TEdit
        Left = 199
        Top = 55
        Width = 66
        Height = 21
        TabOrder = 5
      end
      object EONum: TEdit
        Left = 199
        Top = 98
        Width = 66
        Height = 21
        TabOrder = 6
      end
      object EORaz: TEdit
        Left = 199
        Top = 140
        Width = 66
        Height = 21
        TabOrder = 7
      end
      object EOAut: TEdit
        Left = 199
        Top = 183
        Width = 66
        Height = 21
        TabOrder = 8
      end
      object EOIni: TEdit
        Left = 200
        Top = 232
        Width = 66
        Height = 21
        TabOrder = 9
      end
      object EODet: TEdit
        Left = 199
        Top = 276
        Width = 66
        Height = 21
        TabOrder = 10
      end
      object EERec: TEdit
        Left = 284
        Top = 55
        Width = 66
        Height = 21
        TabOrder = 11
      end
      object EENum: TEdit
        Left = 284
        Top = 98
        Width = 66
        Height = 21
        TabOrder = 12
      end
      object EERaz: TEdit
        Left = 284
        Top = 140
        Width = 66
        Height = 21
        TabOrder = 13
      end
      object EEAut: TEdit
        Left = 284
        Top = 183
        Width = 66
        Height = 21
        TabOrder = 14
      end
      object EEIni: TEdit
        Left = 284
        Top = 232
        Width = 66
        Height = 21
        TabOrder = 15
      end
      object EEDet: TEdit
        Left = 283
        Top = 276
        Width = 66
        Height = 21
        TabOrder = 16
      end
      object EDRec: THNumEdit
        Left = 364
        Top = 55
        Width = 66
        Height = 21
        TabOrder = 17
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object EDNum: THNumEdit
        Left = 364
        Top = 98
        Width = 66
        Height = 21
        TabOrder = 18
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object EDRaz: THNumEdit
        Left = 364
        Top = 140
        Width = 66
        Height = 21
        TabOrder = 19
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object EDAut: THNumEdit
        Left = 364
        Top = 183
        Width = 66
        Height = 21
        TabOrder = 20
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object EDIni: THNumEdit
        Left = 364
        Top = 232
        Width = 66
        Height = 21
        TabOrder = 21
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
      object EDDet: THNumEdit
        Left = 364
        Top = 276
        Width = 66
        Height = 21
        TabOrder = 22
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        UseRounding = True
        Validate = False
      end
    end
    object TabSheet5: TTabSheet
      Caption = 'Communication'
      object Bevel5: TBevel
        Left = 0
        Top = 2
        Width = 440
        Height = 325
      end
      object GroupBox13: TGroupBox
        Left = 10
        Top = 15
        Width = 416
        Height = 64
        Caption = 'Pour acc'#233'der '#224' une ligne ext'#233'rieure, composer le'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object HLabel25: THLabel
          Left = 39
          Top = 29
          Width = 135
          Height = 13
          Caption = '&Num'#233'ro pour ligne ext'#233'rieure'
          FocusControl = EExterieure
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object EExterieure: TEdit
          Left = 202
          Top = 28
          Width = 72
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 10
          ParentFont = False
          TabOrder = 0
        end
      end
      object GroupBox14: TGroupBox
        Left = 10
        Top = 91
        Width = 416
        Height = 64
        Caption = 'Asynchrone'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object HLabel26: THLabel
          Left = 39
          Top = 28
          Width = 102
          Height = 13
          Caption = 'Num'#233'ro de &t'#233'l'#233'phone'
          FocusControl = ETelAsync
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object ETelAsync: TEdit
          Left = 202
          Top = 24
          Width = 98
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 20
          ParentFont = False
          TabOrder = 0
        end
      end
      object GroupBox15: TGroupBox
        Left = 10
        Top = 169
        Width = 416
        Height = 76
        Caption = 'Synchrone'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        object HLabel27: THLabel
          Left = 39
          Top = 47
          Width = 102
          Height = 13
          Caption = 'Num'#233'ro de t'#233'l'#233'&phone'
          FocusControl = ETelSync
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object CBSync: TCheckBox
          Left = 39
          Top = 23
          Width = 176
          Height = 14
          Alignment = taLeftJustify
          Caption = 'Acc'#232's synchrone'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = CBSyncClick
        end
        object ETelSync: TEdit
          Left = 202
          Top = 44
          Width = 98
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 20
          ParentFont = False
          TabOrder = 1
        end
      end
      object GroupBox16: TGroupBox
        Left = 10
        Top = 254
        Width = 416
        Height = 65
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        object HLabel28: THLabel
          Left = 39
          Top = 30
          Width = 126
          Height = 13
          Caption = 'N.U.I. (identifiant transpa&c)'
          FocusControl = EIdent
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object EIdent: TEdit
          Left = 198
          Top = 26
          Width = 203
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          MaxLength = 40
          ParentFont = False
          TabOrder = 0
        end
      end
    end
    object TabSheet6: TTabSheet
      Caption = 'Avanc'#233's'
      object Bevel6: TBevel
        Left = 0
        Top = 2
        Width = 440
        Height = 325
      end
      object GroupBox17: TGroupBox
        Left = 10
        Top = 10
        Width = 416
        Height = 95
        Caption = 'Mode connexion'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        object HLabel29: THLabel
          Left = 364
          Top = 72
          Width = 33
          Height = 13
          Caption = 'jours(s)'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
        end
        object BVoirJal: TToolbarButton97
          Left = 309
          Top = 18
          Width = 92
          Height = 28
          Caption = '&Voir le journal'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          OnClick = BVoirJalClick
          GlobalIndexImage = 'Z0298_S16G1'
        end
        object CBJalConnect: TCheckBox
          Left = 23
          Top = 23
          Width = 202
          Height = 14
          Alignment = taLeftJustify
          Caption = 'Obtenir un &fichier-journal de connexion'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          OnClick = CBJalConnectClick
        end
        object RBEvSuppr: TRadioButton
          Left = 68
          Top = 47
          Width = 251
          Height = 14
          Caption = '&Ecraser les '#233'v'#232'nements '#224' chaque connexion'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          OnClick = RBEvSupprClick
        end
        object RBEvAncien: TRadioButton
          Left = 68
          Top = 72
          Width = 229
          Height = 14
          Caption = 'E&craser les '#233'v'#232'nements plus anciens de'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          OnClick = RBEvAncienClick
        end
        object ENbJours: THNumEdit
          Left = 309
          Top = 68
          Width = 49
          Height = 21
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          Decimals = 2
          Digits = 12
          Masks.PositiveMask = '#,##0'
          Debit = False
          UseRounding = True
          Validate = False
        end
      end
      object GroupBox18: TGroupBox
        Left = 10
        Top = 117
        Width = 416
        Height = 202
        Caption = 'Mode trace'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        object BActualiser: TToolbarButton97
          Left = 309
          Top = 24
          Width = 92
          Height = 27
          Caption = '&Actualiser'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          OnClick = BActualiserClick
          GlobalIndexImage = 'Z1100_S16G1'
        end
        object BCherche: TToolbarButton97
          Left = 234
          Top = 24
          Width = 28
          Height = 27
          Hint = 'Rechercher'
          Visible = False
          OnClick = BChercheClick
          GlobalIndexImage = 'Z0077_S16G1'
        end
        object CBTrace: TCheckBox
          Left = 14
          Top = 31
          Width = 128
          Height = 14
          Alignment = taLeftJustify
          Caption = 'Obtenir un fichier &trace'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          ParentFont = False
          TabOrder = 0
        end
        object LBTrace: TListBox
          Left = 10
          Top = 63
          Width = 397
          Height = 121
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          IntegralHeight = True
          ItemHeight = 13
          ParentFont = False
          TabOrder = 1
        end
      end
    end
  end
  object Msgs: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      '0;Attention;Le code doit '#234'tre renseign'#233';W;O;O;O;'
      '1;Attention;Ce destinataire existe d'#233'j'#224';W;O;O;O;'
      
        '2;Confirmation;Confirmez-vous la suppression du destinataire ?;Q' +
        ';YN;Y;N;'
      '3;Avertissement;Le fichier de trace n'#39'existe pas;W;O;O;O;'
      '4;Avertissement;La recherche a '#233'chou'#233';W;O;O;O;'
      '5;Avertissement;Impossible de lancer NOTEPAD.EXE;W;O;O;O;'
      '6;Avertissement;Veuillez saisir un destinataire;W;O;O;O;')
    Left = 382
    Top = 169
  end
  object Timer1: TTimer
    Interval = 100
    OnTimer = Timer1Timer
    Left = 385
    Top = 216
  end
  object FindDialog: TFindDialog
    Options = [frDown, frHideMatchCase, frHideWholeWord]
    OnFind = FindDialogFind
    Left = 385
    Top = 265
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 381
    Top = 114
  end
end

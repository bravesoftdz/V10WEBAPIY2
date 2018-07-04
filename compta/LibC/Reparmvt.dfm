object FReparMvt: TFReparMvt
  Left = 459
  Top = 119
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'R'#233'paration comptable'
  ClientHeight = 364
  ClientWidth = 388
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Shape1: TShape
    Left = 252
    Top = 104
    Width = 65
    Height = 65
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 388
    Height = 330
    Align = alClient
    TabOrder = 1
    object TFJal: THLabel
      Left = 9
      Top = 34
      Width = 88
      Height = 13
      Caption = 'Codes j&ournaux de'
      FocusControl = FJal1
    end
    object TFaJ: TLabel
      Left = 251
      Top = 32
      Width = 6
      Height = 13
      Caption = #224
      FocusControl = FJal2
    end
    object TFExercice: THLabel
      Left = 9
      Top = 56
      Width = 88
      Height = 13
      AutoSize = False
      Caption = 'E&xercice'
      FocusControl = FExercice
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFDateCpta1: THLabel
      Left = 9
      Top = 81
      Width = 124
      Height = 13
      AutoSize = False
      Caption = '&Date comptable du'
      FocusControl = FDateCpta1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFDateCpta2: TLabel
      Left = 251
      Top = 56
      Width = 15
      Height = 13
      AutoSize = False
      Caption = 'au'
      FocusControl = FDateCpta2
    end
    object TFEtab: THLabel
      Left = 9
      Top = 104
      Width = 88
      Height = 13
      AutoSize = False
      Caption = '&Etablissement'
      FocusControl = FEtab
    end
    object TFTypeEcriture: THLabel
      Left = 9
      Top = 127
      Width = 88
      Height = 13
      AutoSize = False
      Caption = '&Type d'#39#233'critures'
      FocusControl = FTypeEcriture
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFNumPiece1: THLabel
      Left = 9
      Top = 152
      Width = 92
      Height = 13
      AutoSize = False
      Caption = '&N'#176' de pi'#232'ces de'
      FocusControl = FNumPiece1
    end
    object TFNumPiece2: TLabel
      Left = 255
      Top = 152
      Width = 15
      Height = 13
      AutoSize = False
      Caption = #224
      FocusControl = FNumPiece2
    end
    object HLabel1: THLabel
      Left = 9
      Top = 180
      Width = 114
      Height = 13
      Caption = '&Compte pour '#233'quilibrage'
      FocusControl = GeneAttend
    end
    object H_MODEPAIE: THLabel
      Left = 227
      Top = 180
      Width = 44
      Height = 13
      Caption = 'Paiement'
      FocusControl = MPAttend
    end
    object HLabel2: THLabel
      Left = 10
      Top = 9
      Width = 82
      Height = 13
      Caption = 'T&ype de journaux'
      FocusControl = FJal1
    end
    object FExercice: THValComboBox
      Left = 136
      Top = 52
      Width = 245
      Height = 21
      Style = csDropDownList
      Ctl3D = True
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 3
      OnChange = FExerciceChange
      TagDispatch = 0
      DataType = 'TTEXOSAUFPRECEDENT'
    end
    object FDateCpta1: TMaskEdit
      Left = 136
      Top = 76
      Width = 81
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 4
      Text = '01/01/1900'
      OnKeyPress = FDateCpta1KeyPress
    end
    object FDateCpta2: TMaskEdit
      Left = 294
      Top = 76
      Width = 86
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 5
      Text = '31/12/2099'
      OnKeyPress = FDateCpta1KeyPress
    end
    object FEtab: THValComboBox
      Left = 136
      Top = 100
      Width = 245
      Height = 21
      Style = csDropDownList
      Ctl3D = True
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 6
      TagDispatch = 0
      Vide = True
      DataType = 'TTETABLISSEMENT'
    end
    object FTypeEcriture: THValComboBox
      Left = 136
      Top = 124
      Width = 245
      Height = 21
      Style = csDropDownList
      Ctl3D = True
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 7
      TagDispatch = 0
      Vide = True
      DataType = 'TTQUALPIECE'
    end
    object Panel2: TPanel
      Left = 4
      Top = 284
      Width = 373
      Height = 41
      BevelOuter = bvLowered
      TabOrder = 20
      object TErrD: TLabel
        Left = 24
        Top = 13
        Width = 95
        Height = 13
        Caption = 'Erreur(s) d'#233'tect'#233'e(s)'
      end
      object TErrC: TLabel
        Left = 200
        Top = 13
        Width = 91
        Height = 13
        Caption = 'Erreur(s) corrig'#233'e(s)'
      end
      object NbErrD: THNumEdit
        Left = 128
        Top = 9
        Width = 41
        Height = 21
        Color = clBtnFace
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Masks.PositiveMask = '#,##0'
        Debit = False
        ParentFont = False
        TabOrder = 0
        UseRounding = True
        Validate = False
      end
      object NbErrC: THNumEdit
        Left = 304
        Top = 9
        Width = 41
        Height = 21
        Color = clBtnFace
        Decimals = 2
        Digits = 12
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -12
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Masks.PositiveMask = '#,##0'
        Debit = False
        ParentFont = False
        TabOrder = 1
        UseRounding = True
        Validate = False
      end
    end
    object TCEtab: THValComboBox
      Left = 240
      Top = 260
      Width = 38
      Height = 21
      Color = clYellow
      ItemHeight = 13
      TabOrder = 17
      Visible = False
      TagDispatch = 0
      DataType = 'TTETABLISSEMENT'
    end
    object TCNatPiece: TComboBox
      Left = 336
      Top = 260
      Width = 38
      Height = 21
      Color = clYellow
      ItemHeight = 13
      TabOrder = 18
      Text = 'TCNatPiece'
      Visible = False
    end
    object FNumPiece1: TMaskEdit
      Left = 136
      Top = 148
      Width = 81
      Height = 21
      Ctl3D = True
      MaxLength = 9
      ParentCtl3D = False
      TabOrder = 8
    end
    object FNumPiece2: TMaskEdit
      Left = 294
      Top = 148
      Width = 86
      Height = 21
      Ctl3D = True
      MaxLength = 9
      ParentCtl3D = False
      TabOrder = 9
    end
    object TCModP: THValComboBox
      Left = 292
      Top = 260
      Width = 41
      Height = 21
      Color = clYellow
      ItemHeight = 13
      TabOrder = 19
      Visible = False
      TagDispatch = 0
      DataType = 'TTMODEPAIE'
    end
    object GeneAttend: THDBCpteEdit
      Left = 136
      Top = 176
      Width = 81
      Height = 21
      TabOrder = 10
      ZoomTable = tzGNonCollectif
      Vide = False
      Bourre = False
      okLocate = True
      SynJoker = False
    end
    object CtrlEcrGen: TCheckBox
      Left = 8
      Top = 242
      Width = 185
      Height = 17
      Caption = 'Equilibre en comptabilit'#233' g'#233'n'#233'rale'
      TabOrder = 15
    end
    object CtrlEcrAna: TCheckBox
      Left = 194
      Top = 242
      Width = 183
      Height = 17
      Caption = 'Equilibre en comptabilit'#233' analytique'
      TabOrder = 16
    end
    object CtrlDetGen: TCheckBox
      Left = 8
      Top = 222
      Width = 181
      Height = 17
      Caption = 'Lignes de comptabilit'#233' g'#233'n'#233'rale'
      TabOrder = 13
    end
    object CtrlDetAna: TCheckBox
      Left = 194
      Top = 222
      Width = 181
      Height = 17
      Caption = 'Lignes de comptabilit'#233' analytique'
      TabOrder = 14
    end
    object CtrlTreso: TCheckBox
      Left = 8
      Top = 203
      Width = 181
      Height = 17
      Caption = 'Pi'#232'ces de r'#232'glement'
      TabOrder = 11
    end
    object CtrlAnaOff: TCheckBox
      Left = 194
      Top = 203
      Width = 191
      Height = 17
      Caption = 'Ligne analytique sans ligne g'#233'n'#233'rale'
      TabOrder = 12
    end
    object CtrlSouche: TCheckBox
      Left = 8
      Top = 262
      Width = 185
      Height = 17
      Caption = 'Recalcul des num'#233'ros de compteurs'
      Enabled = False
      TabOrder = 21
    end
    object MPAttend: THValComboBox
      Left = 278
      Top = 176
      Width = 101
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 22
      TagDispatch = 0
      DataType = 'TTMODEPAIE'
    end
    object FJal1: THCritMaskEdit
      Left = 136
      Top = 28
      Width = 101
      Height = 21
      TabOrder = 1
      TagDispatch = 0
      DataType = 'TTJALCRIT'
    end
    object FJal2: THCritMaskEdit
      Left = 274
      Top = 28
      Width = 106
      Height = 21
      TabOrder = 2
      TagDispatch = 0
      Plus = 'OR J_MODESAISIE="BOR" OR J_MODESAISIE="LIB"'
      DataType = 'TTJALCRIT'
    end
    object cbTypeJal: THValComboBox
      Left = 136
      Top = 4
      Width = 101
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = cbTypeJalChange
      TagDispatch = 0
      DataType = 'CPMODESAISIEJAL'
    end
  end
  object HPB: TPanel
    Left = 0
    Top = 330
    Width = 388
    Height = 34
    Align = alBottom
    BevelInner = bvLowered
    Caption = ' '
    Ctl3D = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -12
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentCtl3D = False
    ParentFont = False
    ParentShowHint = False
    ShowHint = False
    TabOrder = 0
    object TTravail: TLabel
      Left = 5
      Top = 11
      Width = 241
      Height = 13
      AutoSize = False
      Caption = 'TTravail'
      FocusControl = BStop
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      WordWrap = True
    end
    object BValider: THBitBtn
      Left = 323
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Lancer la v'#233'rification'
      Default = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
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
    object BFerme: THBitBtn
      Left = 354
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Fermer'
      Cancel = True
      ModalResult = 2
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      Glyph.Data = {
        CE070000424DCE07000000000000360000002800000024000000120000000100
        1800000000009807000000000000000000000000000000000000008284008284
        0082840082840082840082840082840082840082840082840082840082840082
        8400828400828400828400828400828400828400828400828400828400828400
        8284008284008284008284008284008284008284008284008284008284008284
        0082840082840082840082840082840082848482848482840082840082840082
        8400828400828400828400828400828400828400828400828400828400828400
        8284008284008284008284FFFFFF008284008284008284008284008284008284
        0082840082840082840082840082840082840082840082840082840000FF0000
        840000848482840082840082840082840082840082840000FF84828400828400
        8284008284008284008284008284008284008284848284848284FFFFFF008284
        008284008284008284008284008284FFFFFF0082840082840082840082840082
        840082840082840000FF00008400008400008484828400828400828400828400
        00FF000084000084848284008284008284008284008284008284008284848284
        FFFFFF008284848284FFFFFF008284008284008284FFFFFF848284848284FFFF
        FF0082840082840082840082840082840082840000FF00008400008400008400
        00848482840082840000FF000084000084000084000084848284008284008284
        008284008284008284848284FFFFFF008284008284848284FFFFFF008284FFFF
        FF848284008284008284848284FFFFFF00828400828400828400828400828400
        82840000FF000084000084000084000084848284000084000084000084000084
        000084848284008284008284008284008284008284848284FFFFFF0082840082
        84008284848284FFFFFF848284008284008284008284008284848284FFFFFF00
        82840082840082840082840082840082840000FF000084000084000084000084
        0000840000840000840000848482840082840082840082840082840082840082
        84008284848284FFFFFF00828400828400828484828400828400828400828400
        8284FFFFFF848284008284008284008284008284008284008284008284008284
        0000FF0000840000840000840000840000840000848482840082840082840082
        84008284008284008284008284008284008284848284FFFFFF00828400828400
        8284008284008284008284FFFFFF848284008284008284008284008284008284
        0082840082840082840082840082840000840000840000840000840000848482
        8400828400828400828400828400828400828400828400828400828400828400
        8284848284FFFFFF008284008284008284008284008284848284008284008284
        0082840082840082840082840082840082840082840082840082840000FF0000
        8400008400008400008484828400828400828400828400828400828400828400
        8284008284008284008284008284008284848284FFFFFF008284008284008284
        8482840082840082840082840082840082840082840082840082840082840082
        840082840000FF00008400008400008400008400008484828400828400828400
        8284008284008284008284008284008284008284008284008284008284848284
        008284008284008284008284848284FFFFFF0082840082840082840082840082
        840082840082840082840082840000FF00008400008400008484828400008400
        0084000084848284008284008284008284008284008284008284008284008284
        008284008284848284008284008284008284008284008284848284FFFFFF0082
        840082840082840082840082840082840082840082840000FF00008400008400
        00848482840082840000FF000084000084000084848284008284008284008284
        008284008284008284008284008284848284008284008284008284848284FFFF
        FF008284008284848284FFFFFF00828400828400828400828400828400828400
        82840000FF0000840000848482840082840082840082840000FF000084000084
        000084848284008284008284008284008284008284008284848284FFFFFF0082
        84008284848284008284848284FFFFFF008284008284848284FFFFFF00828400
        82840082840082840082840082840082840000FF000084008284008284008284
        0082840082840000FF0000840000840000840082840082840082840082840082
        84008284848284FFFFFFFFFFFF848284008284008284008284848284FFFFFF00
        8284008284848284FFFFFF008284008284008284008284008284008284008284
        0082840082840082840082840082840082840082840000FF0000840000FF0082
        8400828400828400828400828400828400828484828484828400828400828400
        8284008284008284848284FFFFFFFFFFFFFFFFFF848284008284008284008284
        0082840082840082840082840082840082840082840082840082840082840082
        8400828400828400828400828400828400828400828400828400828400828400
        8284008284008284008284008284008284008284008284848284848284848284
        0082840082840082840082840082840082840082840082840082840082840082
        8400828400828400828400828400828400828400828400828400828400828400
        8284008284008284008284008284008284008284008284008284008284008284
        008284008284008284008284008284008284}
      NumGlyphs = 2
    end
    object BStop: THBitBtn
      Left = 257
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Arr'#234'te la v'#233'rification'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      Visible = False
      OnClick = BStopClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0107_S16G1'
      IsControl = True
    end
    object BVentil: THBitBtn
      Tag = 1
      Left = 288
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Param'#233'trages analytiques'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      OnClick = BVentilClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0133_S16G1'
      IsControl = True
    end
    object btnError: THBitBtn
      Left = 225
      Top = 3
      Width = 28
      Height = 27
      Hint = 'Voir la rapport d'#39'erreurs'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 4
      OnClick = btnErrorClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0061_S16G1'
      IsControl = True
    end
  end
  object QSum: TQuery
    DatabaseName = 'SOC'
    Left = 114
    Top = 275
  end
  object QEcr: TQuery
    DatabaseName = 'SOC'
    RequestLive = True
    Left = 114
    Top = 303
  end
  object QAnal: TQuery
    DatabaseName = 'SOC'
    Left = 146
    Top = 303
  end
  object MsgRien: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;?Caption?;Aucun enregistrement ne correspond aux crit'#232'res que ' +
        'vous avez s'#233'lectionn'#233's.;W;O;O;O;'
      
        '1;?Caption?;Date que vous avez renseign'#233'e est en dehors des p'#233'ri' +
        'odes de l'#39'exercice;W;O;O;O;'
      '2;?Caption?;Votre s'#233'lection de dates est incoh'#233'rente;W;O;O;O;'
      '3;?Caption?;Les mouvements sont corrects.;A;O;O;O;'
      
        '4;?Caption?;Confirmez-vous l'#39'arr'#234't de la v'#233'rification en cours ?' +
        ';Q;YN;N;N;'
      
        '5;?Caption?;Confirmez-vous la v'#233'rification des fichiers ?;Q;YN;N' +
        ';N;'
      
        '6;?Caption?;Une ou plusieurs '#233'critures ne sont pas '#233'quilibr'#233'es, ' +
        'd'#233'sirez-vous  continuer ?;Q;YN;N;N;'
      
        '7;?Caption?;Le nombre d'#39'erreur est important, les prochaines err' +
        'eurs ne seront pas consultables !;I;O;O;O;'
      
        '8;?Caption?;Une erreur est survenue pendant le traitement. La r'#233 +
        'paration a '#233't'#233' interrompue;W;O;O;O;'
      '9;?Caption?;Le compte g'#233'n'#233'ral d'#39'attente n'#39'existe pas;W;O;O;O;'
      
        '10;?Caption?;Le compte g'#233'n'#233'ral d'#39'attente ne doit pas '#234'tre ventil' +
        'able,lettrable ou collectif;W;O;O;O;'
      
        '11;?Caption?;La section d'#39'attente n'#39'existe pas sur l'#39'axe 1;W;O;O' +
        ';O;'
      
        '12;?Caption?;La section d'#39'attente n'#39'existe pas sur l'#39'axe 2;W;O;O' +
        ';O;'
      
        '13;?Caption?;La section d'#39'attente n'#39'existe pas sur l'#39'axe 3;W;O;O' +
        ';O;'
      
        '14;?Caption?;La section d'#39'attente n'#39'existe pas sur l'#39'axe 4;W;O;O' +
        ';O;'
      
        '15;?Caption?;La section d'#39'attente n'#39'existe pas sur l'#39'axe 5;W;O;O' +
        ';O;'
      '16;?Caption?;Veuillez s'#233'lectionner un type de journal.;W;O;O;O;'
      
        '17;?Caption?;La fourchette des dates doit aller d'#39'un d'#233'but de mo' +
        'is '#224' une fin de mois.;W;O;O;O;')
    Left = 184
    Top = 255
  end
  object MsgLibel: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Pi'#232'ce non '#233'quilibr'#233'e en devise pivot : '
      'Pi'#232'ce non '#233'quilibr'#233'e en devise : '
      'Pi'#232'ce non '#233'quilibr'#233'e en devise et en devise pivot: '
      'Total analytique diff'#233'rent du total de l'#39#233'criture (Pivot)'
      'Total analytique diff'#233'rent du total de l'#39#233'criture (Devise)'
      
        'Total analytique diff'#233'rent du total de l'#39#233'criture (Pivot et Devi' +
        'se)'
      'L'#39#233'criture n'#39'a pas d'#39'analytique(s) rattach'#233'e(s)'
      'Compte de tiers de l'#39#233'criture non renseign'#233
      'Compte de tiers devrait '#234'tre vide'
      'Le compte g'#233'n'#233'ral n'#39'existe pas'
      'Le compte de tiers n'#39'existe pas'
      'la section analytique n'#39'existe pas sur cette axe'
      'le journal n'#39'existe pas'
      
        'Le compte g'#233'n'#233'ral de l'#39#233'criture analytique est diff'#233'rent de la l' +
        'igne g'#233'n'#233'rale"'
      'Le num'#233'ro de pi'#232'ce est incorrect'
      'Total DEBIT '
      'Total CREDIT '
      
        'Les analytiques rattach'#233's '#224' une ligne d'#39' '#233'criture doit '#234'tre d'#39'an' +
        'alytique comptable '
      'Le compte de g'#233'n'#233'ral de type Tiers n'#39'est pas lettrable'
      'Les montants en devise n'#39'existent pas'
      'Fichier Compte non accessible'
      'Fichier Tiers non accessible'
      'Mouvement non-lettr'#233' incorrect '
      'Mouvement lettr'#233' incorrect '
      'Mouvement lettr'#233' non '#233'quilibr'#233' en devise pivot'
      'Mouvement lettr'#233' non '#233'quilibr'#233' en devise'
      'Le journal n'#39' est pas de nature ODA ou ANA'
      'Le journal n'#39' existe pas sur cette axe analytique'
      'Le compte g'#233'n'#233'ral n'#39' est pas ventilable sur cet axe analytique'
      
        'Ecriture : Un compte de charge ou de produit ne devrait pas '#234'tre' +
        ' renseign'#233' '
      
        'Analytique : Un compte de charge ou de produit ne devrait pas '#234't' +
        're renseign'#233' '
      
        'Ecriture : La nature de la pi'#232'ce ne correspond pas '#224' celle de du' +
        ' compte g'#233'n'#233'ral'
      
        'Analytique : La nature de la pi'#232'ce ne correspond pas '#224' celle de ' +
        'du compte g'#233'n'#233'ral'
      
        'Ecriture : Le compte de banque ne doit pas '#234'tre mouvement'#233' avec ' +
        'cette devise'
      'Ecriture : Le compte g'#233'n'#233'ral est interdit sur ce journal'
      'Ecriture : Le mode de paiement n'#39'est pas renseign'#233
      'Analytique : Le journal n'#39'existe pas'
      'DEBIT '
      'CREDIT '
      'Ecriture : Le mode de paiement n'#39'existe pas '
      'Analytique : Le journal ne doit pas '#234'tre de nature ODA ou ANA'
      
        'Ecriture : La nature de pi'#232'ce de type '#233'cart de change n'#39'a pas de' +
        ' montant en Pivot'
      
        'Ecriture : La nature de pi'#232'ce de type '#233'cart de change ne doit pa' +
        's avoir de montant en Devise'
      
        'Analytique : La nature de pi'#232'ce de type '#233'cart de change n'#39'a pas ' +
        'de montant en Pivot'
      
        'Analytique : La nature de pi'#232'ce de type '#233'cart de change ne doit ' +
        'pas avoir de montant en Devise'
      '45 ;'
      'Ecriture : Le journal n'#39'est pas renseign'#233
      'Analytique : Le journal n'#39'est pas renseign'#233
      'Ecriture : La nature de la pi'#232'ce n'#39'est pas renseign'#233'e'
      'Ecriture : La nature de la pi'#232'ce n'#39'existe pas'
      'Analytique : La nature de la pi'#232'ce n'#39'est pas renseign'#233'e'
      'Analytique : La nature de la pi'#232'ce n'#39'existe pas'
      'Ecriture : La devise n'#39'est pas renseign'#233'e'
      'Ecriture : La devise n'#39'existe pas'
      'Analytique : La devise n'#39'est pas renseign'#233'e'
      'Analytique : La devise n'#39'existe pas'
      '56 ;'
      'Ecriture : Le compte g'#233'n'#233'ral n'#39'est pas renseign'#233
      'Analytique : Le compte g'#233'n'#233'ral n'#39'est pas renseign'#233
      'Analytique : Le code section n'#39'est pas renseign'#233
      'Analytique : L'#39' axe n'#39'est pas renseign'#233
      'Analytique : L'#39' axe n'#39'est pas valide'
      
        'Ecriture : La nature de pi'#232'ce de type '#233'cart de change ne doit pa' +
        's '#234'tre mouvement'#233' en Pivot'
      
        'Analytique : La nature de pi'#232'ce de type '#233'cart de change ne doit ' +
        'pas '#234'tre mouvement'#233' en Pivot'
      
        'Ecriture : Un compte g'#233'n'#233'ral extra-comptable ne devrait pas '#234'tre' +
        ' renseign'#233
      
        'Analytique : Un compte g'#233'n'#233'ral extra-comptable ne devrait pas '#234't' +
        're renseign'#233
      '66 ;'
      
        'Ecriture : La date comptable sur une pi'#232'ce d'#39'anouveau ne corresp' +
        'ond pas '#224' la date du d'#233'but de l'#39'exercice'
      
        'Analytique : La date comptable sur une pi'#232'ce d'#39'anouveau ne corre' +
        'spond pas '#224' la date du d'#233'but de l'#39'exercice'
      'Ecriture : L'#39' '#233'tablissement n'#39'est pas renseign'#233
      'Ecriture : L'#39' '#233'tablissement n'#39'existe pas'
      'Analytique : L'#39' '#233'tablissement n'#39'est pas renseign'#233
      'Analytique : L'#39' '#233'tablissement n'#39'existe pas'
      'Date paquet max ou min incorrect'
      'DEBIT DEV'
      'CREDIT DEV'
      '76 ;'
      'Ecriture : Le taux de la devise pivot est incorrect'
      'Ecriture : Le taux de la devise est incorrect'
      'Analytique : Le taux de la devise pivot est incorrect'
      'Analytique : Le taux de la devise est incorrect'
      'Pi'#232'ce non '#233'quilibr'#233'e en Euro :'
      'Total analytique diff'#233'rent du total de l'#39#233'criture (Euro)'
      '83;'
      '84;'
      '85;'
      '86;'
      '87;'
      '88;'
      '89;'
      '90;'
      
        'Ecriture : La ligne avec un compte lettrable a un '#233'tat de lettra' +
        'ge incoh'#233'rent'
      
        'Ecriture : La ligne avec un compte pointable de banque a un '#233'tat' +
        ' de lettrage incoh'#233'rent'
      'Pi'#232'ce dont tous les mouvements sont '#224' z'#233'ro.')
    Left = 10
    Top = 303
  end
  object MsgLibel2: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Champ'
      'incorrect'
      'R'#233'gularisation de l'#39#233'criture')
    Left = 215
    Top = 255
  end
  object MsgBar: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Analyse de l'#39#233'criture en cours'
      'Journal'
      #224' la date'
      'pi'#232'ce n'#176
      'erreurs'
      'erreur'
      'Aucune erreur'
      'Equilibre '#233'criture'
      'Equilibre analytique'
      'Mouvements'
      'Lettrage'
      'Code'
      'tiers'
      'g'#233'n'#233'ral'
      'Pr'#233'paration des fichiers...'
      'Mouvements analytiques'
      'R'#233'paration des fichiers en cours...'
      '17 ;'
      'V'#233'rification en cours...'
      '<<Tous>>'
      'R'#233'paration en cours ...'
      'de temp restant...')
    Left = 55
    Top = 279
  end
  object QInfoAnaG: TQuery
    DatabaseName = 'SOC'
    Left = 78
    Top = 303
  end
  object QAnaGene: TQuery
    DatabaseName = 'SOC'
    Left = 46
    Top = 303
  end
  object QSumGeAn: TQuery
    DatabaseName = 'SOC'
    Left = 210
    Top = 307
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 176
    Top = 304
  end
end

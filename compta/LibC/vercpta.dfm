object FVerCpta: TFVerCpta
  Left = 543
  Top = 213
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'Contr'#244'le des mouvements'
  ClientHeight = 372
  ClientWidth = 382
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
    Width = 382
    Height = 338
    Align = alClient
    TabOrder = 1
    object TFJalJoker: THLabel
      Left = 5
      Top = 37
      Width = 73
      Height = 13
      Caption = 'Codes j&ournaux'
      FocusControl = FJalJoker
      Visible = False
    end
    object TFJal: THLabel
      Left = 5
      Top = 37
      Width = 88
      Height = 13
      Caption = 'Codes j&ournaux de'
      FocusControl = FJal1
    end
    object TFaJ: TLabel
      Left = 251
      Top = 37
      Width = 6
      Height = 13
      Caption = #224
      FocusControl = FJal2
    end
    object TFExercice: THLabel
      Left = 5
      Top = 61
      Width = 88
      Height = 13
      AutoSize = False
      Caption = 'E&xercice'
      FocusControl = FExercice
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFDateCpta1: THLabel
      Left = 5
      Top = 86
      Width = 124
      Height = 13
      AutoSize = False
      Caption = '&Date comptable du'
      FocusControl = FDateCpta1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFDateCpta2: TLabel
      Left = 247
      Top = 83
      Width = 15
      Height = 13
      AutoSize = False
      Caption = 'au'
      FocusControl = FDateCpta2
    end
    object TFEtab: THLabel
      Left = 5
      Top = 109
      Width = 88
      Height = 13
      AutoSize = False
      Caption = '&Etablissement'
      FocusControl = FEtab
    end
    object TFTypeEcriture: THLabel
      Left = 5
      Top = 132
      Width = 88
      Height = 13
      AutoSize = False
      Caption = '&Type d'#39#233'critures'
      FocusControl = FTypeEcriture
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFVerification: THLabel
      Left = 5
      Top = 11
      Width = 104
      Height = 13
      AutoSize = False
      Caption = '&V'#233'rification'
      FocusControl = FVerification
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TFNumPiece1: THLabel
      Left = 5
      Top = 157
      Width = 92
      Height = 13
      AutoSize = False
      Caption = '&N'#176' de pi'#232'ces de'
      FocusControl = FNumPiece1
    end
    object TFNumPiece2: TLabel
      Left = 251
      Top = 157
      Width = 15
      Height = 13
      AutoSize = False
      Caption = #224
      FocusControl = FNumPiece2
    end
    object FJalJoker: TEdit
      Left = 132
      Top = 33
      Width = 81
      Height = 21
      CharCase = ecUpperCase
      Ctl3D = True
      MaxLength = 3
      ParentCtl3D = False
      TabOrder = 3
      Visible = False
      OnChange = FJal1Change
    end
    object FJal1: THCpteEdit
      Left = 132
      Top = 33
      Width = 81
      Height = 21
      CharCase = ecUpperCase
      TabOrder = 1
      OnChange = FJal1Change
      ZoomTable = tzJournal
      Vide = False
      Bourre = False
      okLocate = False
      SynJoker = False
    end
    object FJal2: THCpteEdit
      Left = 290
      Top = 33
      Width = 86
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 3
      TabOrder = 2
      OnChange = FJal1Change
      ZoomTable = tzJournal
      Vide = False
      Bourre = False
      okLocate = False
      SynJoker = False
    end
    object FExercice: THValComboBox
      Left = 132
      Top = 57
      Width = 245
      Height = 21
      Style = csDropDownList
      Ctl3D = True
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 4
      OnChange = FExerciceChange
      TagDispatch = 0
      DataType = 'TTEXERCICE'
    end
    object FDateCpta1: TMaskEdit
      Left = 132
      Top = 81
      Width = 81
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 5
      Text = '01/01/1900'
      OnKeyPress = FDateCpta1KeyPress
    end
    object FDateCpta2: TMaskEdit
      Left = 290
      Top = 81
      Width = 86
      Height = 21
      Ctl3D = True
      EditMask = '!99/99/0000;1;_'
      MaxLength = 10
      ParentCtl3D = False
      TabOrder = 6
      Text = '31/12/2099'
      OnKeyPress = FDateCpta1KeyPress
    end
    object FEtab: THValComboBox
      Left = 132
      Top = 105
      Width = 245
      Height = 21
      Style = csDropDownList
      Ctl3D = True
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 7
      TagDispatch = 0
      Vide = True
      DataType = 'TTETABLISSEMENT'
    end
    object FTypeEcriture: THValComboBox
      Left = 132
      Top = 129
      Width = 245
      Height = 21
      Style = csDropDownList
      Ctl3D = True
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 8
      TagDispatch = 0
      Vide = True
      DataType = 'TTQUALPIECE'
    end
    object Panel2: TPanel
      Left = 4
      Top = 240
      Width = 373
      Height = 91
      BevelOuter = bvLowered
      TabOrder = 11
      object TNBError1: TLabel
        Left = 5
        Top = 2
        Width = 350
        Height = 13
        AutoSize = False
        Caption = 'TNBError1'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        WordWrap = True
      end
      object TNBError2: TLabel
        Left = 5
        Top = 16
        Width = 350
        Height = 13
        AutoSize = False
        Caption = 'TNBError'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        WordWrap = True
      end
      object TNBError3: TLabel
        Left = 5
        Top = 30
        Width = 350
        Height = 13
        AutoSize = False
        Caption = 'TNBError'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        WordWrap = True
      end
      object TNBError5: TLabel
        Left = 5
        Top = 58
        Width = 350
        Height = 13
        AutoSize = False
        Caption = 'TNBError'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        WordWrap = True
      end
      object TNBError4: TLabel
        Left = 5
        Top = 44
        Width = 350
        Height = 13
        AutoSize = False
        Caption = 'TNBError'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        WordWrap = True
      end
      object TTemp: TLabel
        Left = 225
        Top = 59
        Width = 145
        Height = 13
        Alignment = taRightJustify
        AutoSize = False
        Caption = 'TTemp'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clPurple
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        Visible = False
        WordWrap = True
      end
      object TNbError6: TLabel
        Left = 5
        Top = 72
        Width = 350
        Height = 13
        AutoSize = False
        Caption = 'TNBError'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        WordWrap = True
      end
      object TCModP: THValComboBox
        Left = 324
        Top = 0
        Width = 41
        Height = 21
        Color = clYellow
        ItemHeight = 13
        TabOrder = 0
        Visible = False
        TagDispatch = 0
        DataType = 'TTMODEPAIE'
      end
      object FSolde: THNumEdit
        Left = 200
        Top = 12
        Width = 85
        Height = 21
        Color = clYellow
        Decimals = 2
        Digits = 12
        Masks.PositiveMask = '#,##0'
        Debit = False
        TabOrder = 1
        UseRounding = True
        Validate = False
      end
    end
    object FVerification: TComboBox
      Left = 132
      Top = 8
      Width = 245
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = FVerificationChange
      Items.Strings = (
        'Toutes'
        'Comptable'
        'Lettrage'
        'Budg'#233'taire')
    end
    object TCEtab: THValComboBox
      Left = 332
      Top = 176
      Width = 38
      Height = 21
      Color = clYellow
      ItemHeight = 13
      TabOrder = 12
      Visible = False
      TagDispatch = 0
      DataType = 'TTETABLISSEMENT'
    end
    object TCNatPiece: TComboBox
      Left = 328
      Top = 200
      Width = 38
      Height = 21
      Color = clYellow
      ItemHeight = 13
      TabOrder = 13
      Text = 'TCNatPiece'
      Visible = False
    end
    object FNumPiece1: TMaskEdit
      Left = 132
      Top = 153
      Width = 81
      Height = 21
      Ctl3D = True
      MaxLength = 9
      ParentCtl3D = False
      TabOrder = 9
    end
    object FNumPiece2: TMaskEdit
      Left = 290
      Top = 153
      Width = 86
      Height = 21
      Ctl3D = True
      MaxLength = 9
      ParentCtl3D = False
      TabOrder = 10
    end
    object CtrlDetGen: TCheckBox
      Left = 4
      Top = 176
      Width = 181
      Height = 17
      Caption = 'Lignes de comptabilit'#233' g'#233'n'#233'rale'
      Checked = True
      Enabled = False
      State = cbChecked
      TabOrder = 14
    end
    object CtrlEcrGen: TCheckBox
      Left = 4
      Top = 196
      Width = 185
      Height = 17
      Caption = 'Equilibre en comptabilit'#233' g'#233'n'#233'rale'
      Checked = True
      Enabled = False
      State = cbChecked
      TabOrder = 15
    end
    object CtrlEcrAna: TCheckBox
      Left = 190
      Top = 196
      Width = 183
      Height = 17
      Caption = 'Equilibre en comptabilit'#233' analytique'
      Checked = True
      Enabled = False
      State = cbChecked
      TabOrder = 16
    end
    object CtrlDetAna: TCheckBox
      Left = 190
      Top = 176
      Width = 181
      Height = 17
      Caption = 'Lignes de comptabilit'#233' analytique'
      Checked = True
      Enabled = False
      State = cbChecked
      TabOrder = 17
    end
    object CtrlAnaOff: TCheckBox
      Left = 4
      Top = 217
      Width = 191
      Height = 17
      Caption = 'Ligne analytique sans ligne g'#233'n'#233'rale'
      TabOrder = 18
    end
  end
  object HPB: TPanel
    Left = 0
    Top = 338
    Width = 382
    Height = 34
    Align = alBottom
    BevelInner = bvLowered
    Caption = ' '
    Ctl3D = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
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
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      WordWrap = True
    end
    object BValider: THBitBtn
      Left = 319
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Lancer la v'#233'rification'
      Default = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
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
      Left = 350
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
      Left = 288
      Top = 4
      Width = 28
      Height = 27
      Hint = 'Arr'#234'te la v'#233'rification'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = BStopClick
      Margin = 2
      Spacing = -1
      GlobalIndexImage = 'Z0107_S16G1'
      IsControl = True
    end
  end
  object MsgRien: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Contr'#244'le des mouvements;Aucun enregistrement ne correspond aux' +
        ' crit'#232'res que vous avez s'#233'lectionn'#233's.;W;O;O;O;'
      
        '1;Contr'#244'le des mouvements;La date que vous avez renseign'#233'e est e' +
        'n dehors des p'#233'riodes de l'#39'exercice;W;O;O;O;'
      
        '2;Contr'#244'le des mouvements;Votre s'#233'lection de dates est incoh'#233'ren' +
        'te;W;O;O;O;'
      '3;Contr'#244'le des mouvements;Les mouvements sont corrects.;A;O;O;O;'
      
        '4;Contr'#244'le des mouvements;Confirmez-vous l'#39'arr'#234't de la v'#233'rificat' +
        'ion en cours ?;Q;YN;N;N;'
      
        '5;Contr'#244'le des mouvements;Confirmez-vous la v'#233'rification des fic' +
        'hiers ?;Q;YN;N;N;'
      
        '6;Contr'#244'le des mouvements;Une ou plusieurs '#233'critures ne sont pas' +
        ' '#233'quilibr'#233'es, d'#233'sirez-vous  continuer ?;Q;YN;N;N;'
      
        '7;?Caption?;Le nombre d'#39'erreur est important, les prochaines err' +
        'eurs ne seront pas consultables !;I;OC;C;C;')
    Left = 144
    Top = 65535
  end
  object MsgLibel: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clWindowText
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      'Pi'#232'ce non '#233'quilibr'#233'e en devise pivot :'
      'Pi'#232'ce non '#233'quilibr'#233'e en devise :'
      'Pi'#232'ce non '#233'quilibr'#233'e en devise et en devise pivot:'
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
      'Total DEBIT'
      'Total CREDIT'
      
        'Les analytiques rattach'#233's '#224' une ligne d'#39#233'criture doivent '#234'tre d'#39 +
        'analytique comptable'
      'Le compte de g'#233'n'#233'ral de type Tiers n'#39'est pas lettrable'
      'Les montants en devise n'#39'existent pas'
      'Fichier Compte non accessible'
      'Fichier Tiers non accessible'
      'Mouvement non-lettr'#233' incorrect'
      'Mouvement lettr'#233' incorrect'
      'Mouvement lettr'#233' non '#233'quilibr'#233' en devise pivot'
      'Mouvement lettr'#233' non '#233'quilibr'#233' en devise'
      'Le journal n'#39' est pas de nature ODA ou ANA'
      'Le journal n'#39' existe pas sur cette axe analytique'
      'Le compte g'#233'n'#233'ral n'#39'est pas ventilable sur cet axe analytique'
      
        'Ecriture : Un compte de charge ou de produit ne devrait pas '#234'tre' +
        ' renseign'#233
      
        'Analytique : Un compte de charge ou de produit ne devrait pas '#234't' +
        're renseign'#233
      
        'Ecriture : La nature de la pi'#232'ce ne correspond pas '#224' celle de du' +
        ' compte g'#233'n'#233'ral'
      
        'Analytique : La nature de la pi'#232'ce ne correspond pas '#224' celle de ' +
        'du compte g'#233'n'#233'ral'
      
        'Ecriture : Le compte de banque ne doit pas '#234'tre mouvement'#233' avec ' +
        'cette devise'
      'Ecriture : Le compte g'#233'n'#233'ral est interdit sur ce journal'
      'Ecriture : Le mode de paiement n'#39'est pas renseign'#233
      'Analytique : Le journal n'#39'existe pas'
      'DEBIT'
      'CREDIT'
      'Ecriture : Le mode de paiement n'#39'existe pas'
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
      'Analytique : L'#39'axe n'#39'est pas renseign'#233
      'Analytique : L'#39'axe n'#39'est pas valide'
      
        'Ecriture : La nature de pi'#232'ce de type '#233'cart de change ne doit pa' +
        's '#234'tre mouvement'#233'e en Pivot'
      
        'Analytique : La nature de pi'#232'ce de type '#233'cart de change ne doit ' +
        'pas '#234'tre mouvement'#233'e en Pivot'
      
        'Ecriture : Un compte g'#233'n'#233'ral extra-comptable ne devrait pas '#234'tre' +
        ' renseign'#233
      
        'Analytique : Un compte g'#233'n'#233'ral extra-comptable ne devrait pas '#234't' +
        're renseign'#233
      '66 ;'
      
        'Ecriture : La date comptable sur une pi'#232'ce d'#39'a-nouveau ne corres' +
        'pond pas '#224' la date du d'#233'but de l'#39'exercice'
      
        'Analytique : La date comptable sur une pi'#232'ce d'#39'a-nouveau ne corr' +
        'espond pas '#224' la date du d'#233'but de l'#39'exercice'
      'Ecriture : L'#39#233'tablissement n'#39'est pas renseign'#233
      'Ecriture : L'#39#233'tablissement n'#39'existe pas'
      'Analytique : L'#39#233'tablissement n'#39'est pas renseign'#233
      'Analytique : L'#39#233'tablissement n'#39'existe pas'
      'Date paquet max ou min incorrect'
      'DEBIT DEV'
      'CREDIT DEV'
      '76 ;'
      'Ecriture : Le taux de la devise pivot est incorrect'
      'Ecriture : Le taux de la devise est incorrect'
      'Analytique : Le taux de la devise pivot est incorrect'
      'Analytique : Le taux de la devise est incorrect'
      'Pi'#232'ce non '#233'quilibr'#233'e en monnaie de contrevaleur :'
      
        'Total analytique diff'#233'rent du total de l'#39#233'criture (monnaie de co' +
        'ntrevaleur)'
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
      'Pi'#232'ce dont tous les mouvements sont '#224' z'#233'ro.'
      'Ecriture : Le compte de banque n'#39'a pas de devise associ'#233'e (RIB)'
      
        'L'#39'exercice budg'#233'taire n'#39'est pas en corr'#233'lation avec les dates bu' +
        'dg'#233'taires'
      'la section budg'#233'taire n'#39'existe pas sur cet axe'
      'la section budg'#233'taire n'#39'est pas renseign'#233'e'
      'L'#39'axe n'#39'est pas renseign'#233
      'L'#39'axe n'#39'est pas valide'
      'Le budget n'#39'est pas renseign'#233
      'le budget n'#39'existe pas'
      'Le compte budg'#233'taire n'#39'est pas renseign'#233
      'le compte budg'#233'taire n'#39'existe pas'
      
        'Ecriture : La date comptable n'#39'appartient pas '#224' un exercice ouve' +
        'rt'
      'Ecriture : La nature de pi'#232'ce est incoh'#233'rente pour ce journal'
      'Ecriture : Les lignes de la pi'#232'ce'
      'ne sont pas cons'#233'cutives'
      'Analytique : Ce mouvement n'#39'a pas d'#39#233'criture comptable rattach'#233'e'
      'Pi'#232'ce non '#233'quilibr'#233'e en monnaie de contrevaleur :'
      
        'Total analytique diff'#233'rent du total de l'#39#233'criture (monnaie de co' +
        'ntrevaleur)'
      'Ecriture : Le mode de saisie ne correspond '#224' celui du journal'
      'Ecriture : Les dates de la pi'#232'ce '
      'sont diff'#233'rentes '
      'La p'#233'riode est cl'#244'tur'#233'e.'
      'L'#39#233'criture a un type de pi'#232'ce diff'#233'rent de la pi'#232'ce'
      '116'
      'Ecriture : Le journal n'#39'est pas multidevise'
      'Analytique : Le journal n'#39'est pas multidevise'
      ' ')
    Left = 243
    Top = 65535
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
    Left = 295
    Top = 65535
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
      'de temps restant...'
      'Mouvements budg'#233'taires'
      'Contr'#244'le des analytiques sans '#233'criture rattach'#233'e...'
      '24')
    Left = 196
    Top = 65535
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 224
    Top = 120
  end
end

object FCloS: TFCloS
  Left = 320
  Top = 162
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = 'Cl'#244'ture d'#39'exercice provisoire'
  ClientHeight = 328
  ClientWidth = 582
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object GBencours: TGroupBox
    Left = 0
    Top = 0
    Width = 582
    Height = 153
    Align = alTop
    Caption = 'En cours...'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clActiveCaption
    Font.Height = 11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
    object EnCours: TLabel
      Left = 42
      Top = 52
      Width = 513
      Height = 20
      Alignment = taCenter
      AutoSize = False
      Caption = 'EnCours'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object GroupBox2: TGroupBox
      Left = 327
      Top = 76
      Width = 245
      Height = 69
      Caption = 'Comptes trait'#233's'
      Color = clBtnFace
      Ctl3D = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = False
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 1
      object Cpt1: TLabel
        Left = 16
        Top = 16
        Width = 177
        Height = 21
        AutoSize = False
        Caption = 'Cpt1'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Cpt2: TLabel
        Left = 16
        Top = 40
        Width = 193
        Height = 21
        AutoSize = False
        Caption = 'Cpt2'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
    end
    object GroupBox1: TGroupBox
      Left = 12
      Top = 76
      Width = 273
      Height = 69
      Caption = 'Etablissement/Devise'
      Color = clBtnFace
      Ctl3D = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = False
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 0
      object HCpt1: TLabel
        Left = 12
        Top = 16
        Width = 201
        Height = 21
        Alignment = taCenter
        AutoSize = False
        Caption = 'HCpt1'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object HCpt2: TLabel
        Left = 12
        Top = 40
        Width = 201
        Height = 21
        Alignment = taCenter
        AutoSize = False
        Caption = 'HCpt2'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
    end
    object Panel1: TPanel
      Left = 14
      Top = 20
      Width = 557
      Height = 29
      TabOrder = 2
      object Label6: TLabel
        Left = 156
        Top = 5
        Width = 113
        Height = 20
        Alignment = taCenter
        AutoSize = False
        Caption = 'Passage N'#176
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object XX: TLabel
        Left = 260
        Top = 5
        Width = 33
        Height = 20
        Alignment = taCenter
        AutoSize = False
        Caption = 'XX'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label8: TLabel
        Left = 292
        Top = 5
        Width = 25
        Height = 20
        Alignment = taCenter
        AutoSize = False
        Caption = '/'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object YY: TLabel
        Left = 316
        Top = 5
        Width = 33
        Height = 20
        Alignment = taCenter
        AutoSize = False
        Caption = 'YY'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
    end
  end
  object GBOuvre: TGroupBox
    Left = 0
    Top = 0
    Width = 581
    Height = 153
    Caption = 'Param'#232'tres d'#39'ouverture... '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clActiveCaption
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    object HLabel3: THLabel
      Left = 19
      Top = 77
      Width = 23
      Height = 13
      Caption = 'Bilan'
      FocusControl = BilC
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HLabel4: THLabel
      Left = 263
      Top = 104
      Width = 83
      Height = 13
      Caption = 'R'#233'sultat b'#233'n'#233'fice'
      FocusControl = BenC
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HLabel5: THLabel
      Left = 19
      Top = 104
      Width = 66
      Height = 13
      Caption = 'R'#233'sultat perte'
      FocusControl = PerC
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HLabel6: THLabel
      Left = 19
      Top = 49
      Width = 90
      Height = 13
      Caption = 'Journal d'#39'ouverture'
      FocusControl = JalC
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HLabel8: THLabel
      Left = 19
      Top = 22
      Width = 50
      Height = 13
      Caption = 'R'#233'f'#233'rence'
      FocusControl = JalC
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HLabel9: THLabel
      Left = 263
      Top = 22
      Width = 30
      Height = 13
      Caption = 'Libell'#233
      FocusControl = JalC
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HLabel10: THLabel
      Left = 263
      Top = 49
      Width = 104
      Height = 13
      AutoSize = False
      Caption = 'Nouvel exercice de'
      Enabled = False
      FocusControl = DateDebN1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label7: TLabel
      Left = 451
      Top = 49
      Width = 6
      Height = 13
      Caption = #224
      Enabled = False
      FocusControl = DateFinN1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object LPatience: TLabel
      Left = 130
      Top = 127
      Width = 347
      Height = 16
      Alignment = taCenter
      Caption = 'Cl'#244'ture en cours de traitement. Veuillez patienter...'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = False
      ParentFont = False
      Visible = False
    end
    object DateDebN1: TMaskEdit
      Left = 367
      Top = 45
      Width = 75
      Height = 21
      Ctl3D = True
      Enabled = False
      EditMask = '!00/00/0000;1;_'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 10
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 6
      Text = '  /  /    '
    end
    object BilO: THCpteEdit
      Left = 118
      Top = 73
      Width = 121
      Height = 21
      Enabled = False
      TabOrder = 3
      ZoomTable = tzGeneral
      Vide = False
      Bourre = False
      okLocate = False
      SynJoker = False
    end
    object PerO: THCpteEdit
      Left = 118
      Top = 100
      Width = 121
      Height = 21
      Enabled = False
      TabOrder = 4
      ZoomTable = tzGeneral
      Vide = False
      Bourre = False
      okLocate = False
      SynJoker = False
    end
    object BenO: THCpteEdit
      Left = 367
      Top = 100
      Width = 133
      Height = 21
      Enabled = False
      TabOrder = 5
      ZoomTable = tzGeneral
      Vide = False
      Bourre = False
      okLocate = False
      SynJoker = False
    end
    object JalO: THCpteEdit
      Left = 118
      Top = 45
      Width = 121
      Height = 21
      TabOrder = 2
      ZoomTable = tzJAN
      Vide = False
      Bourre = False
      okLocate = True
      SynJoker = False
    end
    object RefO: TEdit
      Left = 118
      Top = 18
      Width = 121
      Height = 21
      TabOrder = 0
    end
    object LibO: TEdit
      Left = 367
      Top = 18
      Width = 173
      Height = 21
      TabOrder = 1
    end
    object DateFinN1: TMaskEdit
      Left = 465
      Top = 45
      Width = 75
      Height = 21
      Ctl3D = True
      Enabled = False
      EditMask = '!00/00/0000;1;_'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 10
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 7
      Text = '  /  /    '
    end
  end
  object GBFerme: TGroupBox
    Left = 0
    Top = 153
    Width = 582
    Height = 138
    Align = alTop
    Caption = 'Param'#232'tres de fermeture... '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clActiveCaption
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    object TSO_FERMEBIL: THLabel
      Left = 19
      Top = 82
      Width = 23
      Height = 13
      Caption = 'Bilan'
      FocusControl = BilC
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TSO_FERMEBEN: THLabel
      Left = 263
      Top = 110
      Width = 83
      Height = 13
      Caption = 'R'#233'sultat b'#233'n'#233'fice'
      FocusControl = BenC
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TSO_FERMEPERTE: THLabel
      Left = 19
      Top = 112
      Width = 66
      Height = 13
      Caption = 'R'#233'sultat perte'
      FocusControl = PerC
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TSO_JALFERME: THLabel
      Left = 19
      Top = 53
      Width = 96
      Height = 13
      Caption = 'Journal de fermeture'
      FocusControl = JalC
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object TSO_RESULTAT: THLabel
      Left = 263
      Top = 82
      Width = 101
      Height = 13
      Caption = 'R'#233'sultat interm'#233'diaire'
      FocusControl = ResC
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HLabel1: THLabel
      Left = 19
      Top = 22
      Width = 50
      Height = 13
      Caption = 'R'#233'f'#233'rence'
      FocusControl = RefC
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HLabel2: THLabel
      Left = 263
      Top = 22
      Width = 30
      Height = 13
      Caption = 'Libell'#233
      FocusControl = LibC
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object HLabel12: THLabel
      Left = 263
      Top = 51
      Width = 104
      Height = 13
      AutoSize = False
      Caption = 'Exercice '#224' cl'#244'turer de'
      Enabled = False
      FocusControl = DateDebN1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label9: TLabel
      Left = 451
      Top = 51
      Width = 6
      Height = 13
      Caption = #224
      Enabled = False
      FocusControl = DateFinN1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object BilC: THCpteEdit
      Left = 122
      Top = 78
      Width = 121
      Height = 21
      Enabled = False
      TabOrder = 3
      ZoomTable = tzGeneral
      Vide = False
      Bourre = False
      okLocate = False
      SynJoker = False
    end
    object ResC: THCpteEdit
      Left = 367
      Top = 78
      Width = 133
      Height = 21
      Enabled = False
      TabOrder = 4
      ZoomTable = tzGeneral
      Vide = False
      Bourre = False
      okLocate = False
      SynJoker = False
    end
    object PerC: THCpteEdit
      Left = 122
      Top = 108
      Width = 121
      Height = 21
      Enabled = False
      TabOrder = 5
      ZoomTable = tzGeneral
      Vide = False
      Bourre = False
      okLocate = False
      SynJoker = False
    end
    object BenC: THCpteEdit
      Left = 367
      Top = 108
      Width = 133
      Height = 21
      Enabled = False
      TabOrder = 6
      ZoomTable = tzGeneral
      Vide = False
      Bourre = False
      okLocate = False
      SynJoker = False
    end
    object JalC: THCpteEdit
      Left = 122
      Top = 47
      Width = 121
      Height = 21
      TabOrder = 2
      ZoomTable = tzJCloture
      Vide = False
      Bourre = False
      okLocate = True
      SynJoker = False
    end
    object RefC: TEdit
      Left = 122
      Top = 18
      Width = 121
      Height = 21
      TabOrder = 0
    end
    object LibC: TEdit
      Left = 367
      Top = 18
      Width = 173
      Height = 21
      TabOrder = 1
    end
    object DateDebN: TMaskEdit
      Left = 367
      Top = 47
      Width = 75
      Height = 21
      Ctl3D = True
      Enabled = False
      EditMask = '!00/00/0000;1;_'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 10
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 7
      Text = '  /  /    '
    end
    object DateFinN: TMaskEdit
      Left = 465
      Top = 47
      Width = 75
      Height = 21
      Ctl3D = True
      Enabled = False
      EditMask = '!00/00/0000;1;_'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MaxLength = 10
      ParentCtl3D = False
      ParentFont = False
      TabOrder = 8
      Text = '  /  /    '
    end
  end
  object GBTraitement: TGroupBox
    Left = 0
    Top = 291
    Width = 582
    Height = 2
    Align = alClient
    Caption = 'Traitements...'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clActiveCaption
    Font.Height = 11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    object Label1: TLabel
      Left = 52
      Top = 36
      Width = 157
      Height = 18
      AutoSize = False
      Caption = 'Calcul des Charges'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = 11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label2: TLabel
      Left = 52
      Top = 60
      Width = 157
      Height = 18
      AutoSize = False
      Caption = 'Calcul des Produits'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = 11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label3: TLabel
      Left = 52
      Top = 85
      Width = 157
      Height = 18
      AutoSize = False
      Caption = 'Calcul du R'#233'sultat'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = 11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label4: TLabel
      Left = 280
      Top = 44
      Width = 253
      Height = 21
      AutoSize = False
      Caption = 'G'#233'n'#233'ration des '#233'critures de Cl'#244'ture'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = 11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label5: TLabel
      Left = 280
      Top = 68
      Width = 273
      Height = 21
      AutoSize = False
      Caption = 'G'#233'n'#233'ration des '#233'critures d'#39'A-Nouveaux'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = 11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object FAF1: TImage
      Left = 32
      Top = 34
      Width = 17
      Height = 19
      Picture.Data = {
        07544269746D61705A010000424D5A0100000000000076000000280000001300
        0000130000000100040000000000E40000000000000000000000100000001000
        0000000000000000800000800000008080008000000080008000808000008080
        8000C0C0C0000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF00888888888888888888800000888888888888888888800000880000000000
        0000888000008800BFBFBBFBFBF08880000088070BFBFFBFBFBF08800000880B
        0FBFBBFBFBFB0880000088070BFBFFBFBFBF08800000880B70BFBBFBFBFBF080
        00008807B0FBFFBFBFBFB0800000880B700000000000008000008807B7B7BAEA
        0B080880000088000B7B80AEA0800880000088888000880AEA0A088000008888
        88888880AEAE088000008888888888880AEA08800000888888888880AEAE0880
        0000888888888800000008800000888888888888888888800000888888888888
        888888800000}
      Stretch = True
    end
    object FOK1: TImage
      Left = 24
      Top = 34
      Width = 21
      Height = 19
      Picture.Data = {
        07544269746D617042010000424D420100000000000076000000280000001100
        0000110000000100040000000000CC0000000000000000000000100000001000
        0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
        C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF007777777777777777700000007777777777777777700000007777777774F7
        77777000000077777777444F77777000000077777774444F7777700000007000
        00444F44F7777000000070FFF444F0744F777000000070F8884FF0774F777000
        000070FFFFFFF07774F77000000070F88888F077774F7000000070FFFFFFF077
        7774F000000070F88777F07777774000000070FFFF00007777777000000070F8
        8707077777777000000070FFFF00777777777000000070000007777777777000
        0000777777777777777770000000}
      Stretch = True
      Visible = False
    end
    object FAF2: TImage
      Left = 32
      Top = 58
      Width = 17
      Height = 19
      Picture.Data = {
        07544269746D61705A010000424D5A0100000000000076000000280000001300
        0000130000000100040000000000E40000000000000000000000100000001000
        0000000000000000800000800000008080008000000080008000808000008080
        8000C0C0C0000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF00888888888888888888800000888888888888888888800000880000000000
        0000888000008800BFBFBBFBFBF08880000088070BFBFFBFBFBF08800000880B
        0FBFBBFBFBFB0880000088070BFBFFBFBFBF08800000880B70BFBBFBFBFBF080
        00008807B0FBFFBFBFBFB0800000880B700000000000008000008807B7B7BAEA
        0B080880000088000B7B80AEA0800880000088888000880AEA0A088000008888
        88888880AEAE088000008888888888880AEA08800000888888888880AEAE0880
        0000888888888800000008800000888888888888888888800000888888888888
        888888800000}
      Stretch = True
    end
    object FOK2: TImage
      Left = 24
      Top = 58
      Width = 21
      Height = 19
      Picture.Data = {
        07544269746D617042010000424D420100000000000076000000280000001100
        0000110000000100040000000000CC0000000000000000000000100000001000
        0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
        C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF007777777777777777700000007777777777777777700000007777777774F7
        77777000000077777777444F77777000000077777774444F7777700000007000
        00444F44F7777000000070FFF444F0744F777000000070F8884FF0774F777000
        000070FFFFFFF07774F77000000070F88888F077774F7000000070FFFFFFF077
        7774F000000070F88777F07777774000000070FFFF00007777777000000070F8
        8707077777777000000070FFFF00777777777000000070000007777777777000
        0000777777777777777770000000}
      Stretch = True
      Visible = False
    end
    object FAF3: TImage
      Left = 32
      Top = 82
      Width = 17
      Height = 19
      Picture.Data = {
        07544269746D61705A010000424D5A0100000000000076000000280000001300
        0000130000000100040000000000E40000000000000000000000100000001000
        0000000000000000800000800000008080008000000080008000808000008080
        8000C0C0C0000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF00888888888888888888800000888888888888888888800000880000000000
        0000888000008800BFBFBBFBFBF08880000088070BFBFFBFBFBF08800000880B
        0FBFBBFBFBFB0880000088070BFBFFBFBFBF08800000880B70BFBBFBFBFBF080
        00008807B0FBFFBFBFBFB0800000880B700000000000008000008807B7B7BAEA
        0B080880000088000B7B80AEA0800880000088888000880AEA0A088000008888
        88888880AEAE088000008888888888880AEA08800000888888888880AEAE0880
        0000888888888800000008800000888888888888888888800000888888888888
        888888800000}
      Stretch = True
    end
    object FOK3: TImage
      Left = 24
      Top = 82
      Width = 21
      Height = 19
      Picture.Data = {
        07544269746D617042010000424D420100000000000076000000280000001100
        0000110000000100040000000000CC0000000000000000000000100000001000
        0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
        C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF007777777777777777700000007777777777777777700000007777777774F7
        77777000000077777777444F77777000000077777774444F7777700000007000
        00444F44F7777000000070FFF444F0744F777000000070F8884FF0774F777000
        000070FFFFFFF07774F77000000070F88888F077774F7000000070FFFFFFF077
        7774F000000070F88777F07777774000000070FFFF00007777777000000070F8
        8707077777777000000070FFFF00777777777000000070000007777777777000
        0000777777777777777770000000}
      Stretch = True
      Visible = False
    end
    object FAF4: TImage
      Left = 256
      Top = 42
      Width = 17
      Height = 19
      Picture.Data = {
        07544269746D61705A010000424D5A0100000000000076000000280000001300
        0000130000000100040000000000E40000000000000000000000100000001000
        0000000000000000800000800000008080008000000080008000808000008080
        8000C0C0C0000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF00888888888888888888800000888888888888888888800000880000000000
        0000888000008800BFBFBBFBFBF08880000088070BFBFFBFBFBF08800000880B
        0FBFBBFBFBFB0880000088070BFBFFBFBFBF08800000880B70BFBBFBFBFBF080
        00008807B0FBFFBFBFBFB0800000880B700000000000008000008807B7B7BAEA
        0B080880000088000B7B80AEA0800880000088888000880AEA0A088000008888
        88888880AEAE088000008888888888880AEA08800000888888888880AEAE0880
        0000888888888800000008800000888888888888888888800000888888888888
        888888800000}
      Stretch = True
    end
    object FOK4: TImage
      Left = 248
      Top = 42
      Width = 21
      Height = 19
      Picture.Data = {
        07544269746D617042010000424D420100000000000076000000280000001100
        0000110000000100040000000000CC0000000000000000000000100000001000
        0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
        C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF007777777777777777700000007777777777777777700000007777777774F7
        77777000000077777777444F77777000000077777774444F7777700000007000
        00444F44F7777000000070FFF444F0744F777000000070F8884FF0774F777000
        000070FFFFFFF07774F77000000070F88888F077774F7000000070FFFFFFF077
        7774F000000070F88777F07777774000000070FFFF00007777777000000070F8
        8707077777777000000070FFFF00777777777000000070000007777777777000
        0000777777777777777770000000}
      Stretch = True
      Visible = False
    end
    object FAF5: TImage
      Left = 256
      Top = 66
      Width = 17
      Height = 19
      Picture.Data = {
        07544269746D61705A010000424D5A0100000000000076000000280000001300
        0000130000000100040000000000E40000000000000000000000100000001000
        0000000000000000800000800000008080008000000080008000808000008080
        8000C0C0C0000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF00888888888888888888800000888888888888888888800000880000000000
        0000888000008800BFBFBBFBFBF08880000088070BFBFFBFBFBF08800000880B
        0FBFBBFBFBFB0880000088070BFBFFBFBFBF08800000880B70BFBBFBFBFBF080
        00008807B0FBFFBFBFBFB0800000880B700000000000008000008807B7B7BAEA
        0B080880000088000B7B80AEA0800880000088888000880AEA0A088000008888
        88888880AEAE088000008888888888880AEA08800000888888888880AEAE0880
        0000888888888800000008800000888888888888888888800000888888888888
        888888800000}
      Stretch = True
    end
    object FOK5: TImage
      Left = 248
      Top = 66
      Width = 21
      Height = 19
      Picture.Data = {
        07544269746D617042010000424D420100000000000076000000280000001100
        0000110000000100040000000000CC0000000000000000000000100000001000
        0000000000000000BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0
        C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFF
        FF007777777777777777700000007777777777777777700000007777777774F7
        77777000000077777777444F77777000000077777774444F7777700000007000
        00444F44F7777000000070FFF444F0744F777000000070F8884FF0774F777000
        000070FFFFFFF07774F77000000070F88888F077774F7000000070FFFFFFF077
        7774F000000070F88777F07777774000000070FFFF00007777777000000070F8
        8707077777777000000070FFFF00777777777000000070000007777777777000
        0000777777777777777770000000}
      Stretch = True
      Visible = False
    end
  end
  object Dock: TDock97
    Left = 0
    Top = 285
    Width = 582
    Height = 43
    AllowDrag = False
    Position = dpBottom
    object HPB: TToolWindow97
      Left = 0
      Top = 0
      Hint = 'Param'#233'trage cl'#244'ture'
      ClientHeight = 39
      ClientWidth = 582
      ClientAreaHeight = 39
      ClientAreaWidth = 582
      DockPos = 0
      FullSize = True
      ParentShowHint = False
      ShowHint = False
      TabOrder = 0
      object BAide: TToolbarButton97
        Left = 513
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Aide'
        Caption = 'Aide'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BAideClick
        GlobalIndexImage = 'Z1117_S16G1'
        IsControl = True
      end
      object BValider: TToolbarButton97
        Left = 449
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Valider'
        Caption = 'Valider'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        NumGlyphs = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BValiderClick
        GlobalIndexImage = 'Z0003_S16G2'
        IsControl = True
      end
      object BFerme: TToolbarButton97
        Left = 481
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Fermer'
        Caption = 'Fermer'
        Cancel = True
        DisplayMode = dmGlyphOnly
        Flat = False
        ModalResult = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = BFermeClick
        GlobalIndexImage = 'Z0021_S16G1'
      end
      object BParam: TToolbarButton97
        Left = 38
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Param'#232'tres soci'#233't'#233
        Caption = 'Param'#232'tres'
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        Visible = False
        OnClick = BParamClick
        GlobalIndexImage = 'Z0008_S16G1'
        IsControl = True
      end
      object BSoc: TToolbarButton97
        Left = 6
        Top = 2
        Width = 28
        Height = 27
        Hint = 'Soci'#233't'#233
        Caption = 'Soci'#233't'#233
        DisplayMode = dmGlyphOnly
        Flat = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Margin = 2
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        Spacing = -1
        OnClick = BSocClick
        GlobalIndexImage = 'Z0061_S16G1'
        IsControl = True
      end
    end
  end
  object PPatience: TPanel
    Left = 89
    Top = 161
    Width = 321
    Height = 77
    TabOrder = 5
    Visible = False
    object H_TitreGuide: TLabel
      Left = 4
      Top = 18
      Width = 313
      Height = 18
      Alignment = taCenter
      AutoSize = False
      Caption = 'Contr'#244'le de la balance en cours.'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = False
      ParentFont = False
    end
    object Label10: TLabel
      Left = 4
      Top = 38
      Width = 313
      Height = 18
      Alignment = taCenter
      AutoSize = False
      Caption = 'Veuillez patienter quelques instants...'
      Color = clBtnFace
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clMaroon
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = False
      ParentFont = False
    end
    object PFenGuide: TPanel
      Left = 1
      Top = 1
      Width = 319
      Height = 15
      Align = alTop
      BevelInner = bvLowered
      BevelOuter = bvNone
      Color = clActiveCaption
      Enabled = False
      TabOrder = 0
    end
  end
  object HMess: THMsgBox
    Police.Charset = DEFAULT_CHARSET
    Police.Color = clBlack
    Police.Height = -11
    Police.Name = 'MS Sans Serif'
    Police.Style = []
    Mess.Strings = (
      
        '0;Cl'#244'ture comptable;Vous devez renseigner tous les comptes;Q;YNC' +
        ';Y;C;'
      
        '1;Cl'#244'ture comptable;Vous devez renseigner tous les comptes;W;O;O' +
        ';O;'
      'Compte g'#233'n'#233'ral'
      'Tiers'
      'Section'
      '5'
      'PREPARATION CLOTURE'
      'CALCULS DES CHARGES'
      'CALCULS DES PRODUITS'
      'CALCUL DU RESULTAT'
      'CREATION DES ECRITURES DE CLOTURE'
      'CREATION DES ECRITURES D'#39'A-NOUVEAUX'
      '12'
      '13'
      '14'
      'VERIFICATION GENERALE'
      '16;Cl'#244'ture comptable;Compte d'#39'ouverture de bilan;E;O;O;O;'
      '17;Cl'#244'ture comptable;Compte d'#39'ouverture de b'#233'n'#233'fice ;E;O;O;O;'
      '18;Cl'#244'ture comptable;Compte d'#39'ouverture de perte ;E;O;O;O;'
      '19;Cl'#244'ture comptable;Journal d'#39'ouverture ;E;O;O;O;'
      '20;Cl'#244'ture comptable;Compte de fermeture de  bilan ;E;O;O;O;'
      '21;Cl'#244'ture comptable;Compte de fermeture de b'#233'n'#233'fice ;E;O;O;O;'
      '22;Cl'#244'ture comptable;Compte de fermeturede perte ;E;O;O;O;'
      '23;Cl'#244'ture comptable;Journal de fermeture ;E;O;O;O;'
      '24;Cl'#244'ture comptable;Compte de r'#233'sultat ;E;O;O;O;'
      'inexistant'
      'collectif'
      'lettrable'
      'pointable'
      'ventilable'
      'de nature incorrecte'
      ': pas de facturier associ'#233
      
        '32;Cl'#244'ture comptable;Attention au compteur du journal de cl'#244'ture' +
        ' : des journaux autres que celui de cl'#244'ture ont le m'#234'me compteur' +
        '. Confirmez-vous ?;Q;YNC;Y;C;'
      
        '33;Cl'#244'ture comptable;Attention au compteur du journal d'#39'A-Nouvea' +
        'u : des journaux autres que celui d'#39'A-Nouveau ont le m'#234'me compte' +
        'ur. Confirmez-vous ?;Q;YNC;Y;C;'
      
        '34;Cl'#244'ture comptable;1'#176' confirmation : confirmez-vous le traitem' +
        'ent de cl'#244'ture?;Q;YN;Y;N'
      
        '35;Cl'#244'ture comptable;2'#176' confirmation : confirmez-vous le traitem' +
        'ent de cl'#244'ture?;Q;YN;Y;N'
      
        '36;Cl'#244'ture comptable;Probl'#232'me : Le traitement de cl'#244'ture a '#233't'#233' i' +
        'nterrompu ;E;O;O;O;'
      
        '37;Cl'#244'ture comptable;Probl'#232'me : Le traitement de cl'#244'ture a '#233't'#233' i' +
        'nterrompu ;E;O;O;O;'
      
        '38;Cl'#244'ture comptable;Le traitement s'#39'est correctement termin'#233'. A' +
        'ucune anomalie n'#39'a '#233't'#233' d'#233'tect'#233'e;E;O;O;O;'
      '39;Cl'#244'ture comptable'
      'ATTENTION : Un incident s'#39'est produit pendant le traitement.'
      'La cl'#244'ture va '#234'tre annul'#233'e.'
      
        'Les informations comptables vont '#234'tre restitu'#233'es telles qu'#39'elles' +
        ' '#233'taient avant cl'#244'ture.'
      ';E;O;O;O;'
      
        '44;Cl'#244'ture comptable;ATTENTION : Un incident s'#39'est produit penda' +
        'nt le traitement.#13#13La cl'#244'ture va '#234'tre annul'#233'e.#13#13Les info' +
        'rmations comptables vont '#234'tre restitu'#233'es telles qu'#39'elles '#233'taient' +
        ' avant cl'#244'ture.;E;O;O;O;'
      
        '45;Cl'#244'ture comptable;La cl'#244'ture ne peut '#234'tre faite car l'#39'exercic' +
        'e suivant n'#39'est pas ouvert.;E;O;O;O;'
      
        '46;Cl'#244'ture comptable;La balance de l'#39'exercice '#224' cl'#244'turer n'#39'est p' +
        'as '#233'quilibr'#233'e.;E;O;O;O;'
      '47'
      'Cl'#244'ture d'#39'exercice d'#233'finitive'
      'Total des '#233'critures point'#233'es'
      'Total des '#233'critures non point'#233'es'
      '51'
      
        '52;Cl'#244'ture comptable;La cl'#244'ture des immobilisations a-t-elle '#233't'#233 +
        ' faite?;Q;YN;Y;N;'
      
        '53;Cl'#244'ture comptable;La balance IFRS de l'#39'exercice '#224' cl'#244'turer n'#39 +
        'est pas '#233'quilibr'#233'e.;E;O;O;O;'
      'PREPARATION TRAITEMENT'
      'CALCULS DES SOLDES EURO DES CHARGES'
      'CALCULS DES SOLDES EURO DES PRODUITS'
      'CALCUL DU SOLDES EURO DU RESULTAT'
      'MISE A JOUR DES DES ECRITURES DE CLOTURE'
      'MISE A JOUR DES ECRITURES D'#39'A-NOUVEAUX'
      
        '60;Cl'#244'ture comptable;Attention : Certains comptes 6 ne sont pas ' +
        'de nature charge.;W;O;O;O;'
      
        '61;Cl'#244'ture comptable;Attention : Certains comptes 7 ne sont pas ' +
        'de nature produit.;W;O;O;O;'
      
        '62;Cl'#244'ture comptable;Attention : Vous devez modifier la nature d' +
        'e ces comptes.;W;O;O;O;'
      
        '63;Cl'#244'ture comptable;Attention : L'#39'appel au processus serveur de' +
        ' cl'#244'ture n'#39'a pu aboutir. Veuillez v'#233'rifier votre installation.;W' +
        ';O;O;O;'
      
        '64;Cl'#244'ture comptable;Attention : Certains comptes 6 ne sont pas ' +
        'de nature charge.#10#13Vous devriez modifier la nature de ces co' +
        'mptes.#10#13Voulez vous continuer ?;Q;YN;Y;N;'
      
        '65;Cl'#244'ture comptable;Attention : Certains comptes 7 ne sont pas ' +
        'de nature produit.#10#13Vous devriez modifier la nature de ces c' +
        'omptes.#10#13Voulez vous continuer ?;Q;YN;Y;N;'
      
        '66;Cl'#244'ture comptable;ATTENTION : Un incident s'#39'est produit penda' +
        'nt le traitement des IFRS.#13#13La cl'#244'ture va '#234'tre annul'#233'e.#13#1' +
        '3Les informations comptables vont '#234'tre restitu'#233'es telles qu'#39'elle' +
        's '#233'taient avant cl'#244'ture.;E;O;O;O;'
      
        '67;Cl'#244'ture comptable;ATTENTION : Pr'#233'sence d'#39#233'critures non valid'#233 +
        'es. Voulez-vous continuer ?;Q;YN;N;N;')
    Left = 520
    Top = 140
  end
  object HMTrad: THSystemMenu
    Caption = '&Personnalisation'
    Separator = True
    Traduction = True
    Left = 492
    Top = 140
  end
  object Timer1: TTimer
    Enabled = False
    OnTimer = Timer1Timer
    Left = 509
    Top = 108
  end
end
